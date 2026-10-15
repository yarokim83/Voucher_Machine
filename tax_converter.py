"""Convert an unlocked invoice HTML to a verified PDF without accessing Tk."""
import base64
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import traceback
import uuid
from file_utils import temp_output
import pdf_parser


def _conversion_logger():
    logger = logging.getLogger('voucherpass.tax_conversion')
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        logger.propagate = False
        try:
            folder = Path(os.getenv('APPDATA') or Path.home()) / 'VoucherPass'
            folder.mkdir(parents=True, exist_ok=True)
            handler = RotatingFileHandler(folder / 'tax_conversion.log',
                                         maxBytes=1_000_000, backupCount=3, encoding='utf-8')
            handler.setFormatter(logging.Formatter('%(asctime)s %(message)s'))
            logger.addHandler(handler)
        except OSError:
            logger.addHandler(logging.NullHandler())
    return logger


def convert_tax_html(html_path):
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1200,1600')
    # Edge's compatibility-layer relaunch exits the process watched by WebDriver.
    options.add_argument('--edge-skip-compat-layer-relaunch')
    logger = _conversion_logger()
    run_id = uuid.uuid4().hex[:12]
    stage = 'Edge 실행'

    def enter_stage(value):
        nonlocal stage
        stage = value
        logger.info('run=%s stage=%s', run_id, stage)

    driver = None
    output = None
    try:
        enter_stage('Edge 실행')
        driver = webdriver.Edge(options=options)
        logger.info('run=%s browser_version=%s', run_id, driver.capabilities.get('browserVersion', 'unknown'))
        driver.set_page_load_timeout(30)
        enter_stage('HTML 열기')
        driver.get(Path(html_path).resolve().as_uri())

        def password_field(browser):
            return next((item for item in browser.find_elements(By.CSS_SELECTOR, 'input[type=password], input[type=text]') if item.is_displayed()), False)

        enter_stage('비밀번호 입력')
        field = WebDriverWait(driver, 10, poll_frequency=0.2).until(password_field)
        field.clear()
        field.send_keys(pdf_parser.DEFAULT_PASSWORDS[0])
        buttons = driver.find_elements(By.CSS_SELECTOR, 'button, input[type=button], input[type=submit], a[onclick]')
        confirm = next((button for button in buttons if button.is_displayed() and
                        ('확인' in (button.text or button.get_attribute('value') or '') or
                         any(word in (button.get_attribute('onclick') or '') for word in ('Cri', 'InputPwd')))), None)
        if confirm is None:
            raise ValueError('세금계산서 비밀번호 확인 버튼을 찾을 수 없습니다.')
        confirm.click()
        enter_stage('세금계산서 본문 확인')
        WebDriverWait(driver, 15, poll_frequency=0.2).until(lambda browser: browser.execute_script("""
            var f = document.getElementById('CriMsgPosition');
            if (!f) return false;
            var d = f.contentDocument || f.contentWindow.document;
            var t = d && d.body ? d.body.innerText : '';
            return /세금\\s*계산서/.test(t) && /공\\s*급/.test(t);
        """))
        driver.execute_script("""
            ['CriBtnPosition', 'idPcPwdDlg', 'idMobilePwdDlg'].forEach(function(id) {
                var e = document.getElementById(id); if (e) e.style.display = 'none';
            });
            var f = document.getElementById('CriMsgPosition');
            var d = f.contentDocument || f.contentWindow.document;
            f.style.height = (Math.max(d.body.scrollHeight, d.documentElement.scrollHeight) + 30) + 'px';
            f.style.overflow = 'hidden'; d.body.style.overflow = 'hidden';
        """)
        enter_stage('PDF 생성')
        driver.execute_cdp_cmd('Emulation.setEmulatedMedia', {'media': 'print'})
        result = driver.execute_cdp_cmd('Page.printToPDF', {
            'landscape': False, 'displayHeaderFooter': False, 'printBackground': True,
            'preferCSSPageSize': False, 'paperWidth': 8.27, 'paperHeight': 11.69,
            'marginTop': 0.15, 'marginBottom': 0.15, 'marginLeft': 0.15,
            'marginRight': 0.15, 'scale': 0.95})
        output = Path(temp_output('NTS_eTaxInvoice.pdf'))
        output.write_bytes(base64.b64decode(result['data'], validate=True))
        enter_stage('PDF 본문 검증')
        text = pdf_parser.extract_pdf_text_safely(str(output))
        compact = ''.join(text.split())
        if '세금계산서' not in compact or '공급' not in compact:
            raise ValueError('생성된 PDF에서 세금계산서 본문을 확인하지 못했습니다.')
        enter_stage('날짜·금액 추출')
        data = pdf_parser.parse_tax_invoice_data(str(output))
        if not data.get('date'):
            raise ValueError('생성된 PDF의 작성일자를 확인하지 못했습니다. 원문을 확인해 주세요.')
        enter_stage('완료')
        return str(output), data
    except Exception as exc:
        # Do not persist invoice contents, passwords, URLs or exception messages.
        frames = ' > '.join(f'{Path(f.filename).name}:{f.lineno}:{f.name}'
                            for f in traceback.extract_tb(exc.__traceback__))
        logger.error('run=%s stage=%s error=%s frames=%s',
                     run_id, stage, type(exc).__name__, frames)
        if output:
            output.unlink(missing_ok=True)
        raise RuntimeError(f'{stage} 단계 실패 ({type(exc).__name__}). 진단 ID: {run_id}') from exc
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
