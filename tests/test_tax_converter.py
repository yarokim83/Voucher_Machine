import base64
import logging
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import tax_converter


class TaxConversionTests(unittest.TestCase):
    def test_startup_failure_has_stage_without_sensitive_message(self):
        logger = MagicMock()
        with patch('tax_converter._conversion_logger', return_value=logger), \
             patch('selenium.webdriver.Edge', side_effect=RuntimeError('secret-invoice-data')) as edge:
            with self.assertRaisesRegex(RuntimeError, 'Edge 실행 단계 실패') as error:
                tax_converter.convert_tax_html('invoice.html')
        self.assertNotIn('secret-invoice-data', str(error.exception))
        self.assertNotIn('secret-invoice-data', str(logger.mock_calls))
        self.assertIn('--edge-skip-compat-layer-relaunch', edge.call_args.kwargs['options'].arguments)

    def test_invalid_pdf_removed_and_browser_closed(self):
        driver = MagicMock()
        driver.execute_cdp_cmd.return_value = {'data': base64.b64encode(b'invalid pdf').decode()}
        button = MagicMock()
        button.text = '확인'
        driver.find_elements.return_value = [button]
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'invoice.pdf'
            with patch('tax_converter._conversion_logger'), \
                 patch('selenium.webdriver.Edge', return_value=driver), \
                 patch('selenium.webdriver.support.ui.WebDriverWait'), \
                 patch('tax_converter.temp_output', return_value=str(output)), \
                 patch('tax_converter.pdf_parser.extract_pdf_text_safely', return_value=''):
                with self.assertRaisesRegex(RuntimeError, 'PDF 본문 검증 단계 실패'):
                    tax_converter.convert_tax_html('invoice.html')
            self.assertFalse(output.exists())
        driver.quit.assert_called_once()

    def test_logger_works_without_debug_environment(self):
        logger = logging.getLogger('voucherpass.tax_conversion')
        previous = logger.handlers[:]
        logger.handlers = []
        try:
            with tempfile.TemporaryDirectory() as folder:
                with patch.dict('os.environ', {'APPDATA': folder, 'VOUCHERPASS_DEBUG': '0'}):
                    tax_converter._conversion_logger().info('diagnostic-check')
                    for handler in logger.handlers:
                        handler.close()
                    text = (Path(folder) / 'VoucherPass/tax_conversion.log').read_text(encoding='utf-8')
                    self.assertIn('diagnostic-check', text)
        finally:
            logger.handlers = previous
