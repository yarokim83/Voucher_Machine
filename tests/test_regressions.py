import ast
import os
from pathlib import Path
import tempfile
import threading
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
import openpyxl
import pdf_parser
import printer_handler
import excel_handler
import outlook_handler
from file_utils import temp_output, safe_name
from voucher_core import DocumentState, extract_amounts, parse_pages, parse_money
from task_runner import TaskRunner


class AmountTests(unittest.TestCase):
    def test_supply_not_total(self):
        data = extract_amounts('공급가액 1,000,000\n부가세 100,000\n합계 1,100,000')
        self.assertEqual((data['amount'], data['vat'], data['total_amount']), (1000000, 100000, 1100000))
        self.assertFalse(data['warnings'])

    def test_total_only_not_guessed(self):
        data = extract_amounts('합계 1,100,000')
        self.assertIsNone(data['amount'])
        self.assertIsNone(data['vat'])

    def test_zero_vat_table(self):
        data = extract_amounts('공급가액 세액 합계금액\n1,000,000 0 1,000,000')
        self.assertEqual(data['vat'], 0)
        self.assertFalse(data['warnings'])

    def test_nts_date_amount_table(self):
        data = extract_amounts('작성일자 공급가액 세액 수정사유\n2026/08/11 2,100,000 210,000')
        self.assertEqual(data['amount'], 2100000)
        self.assertEqual(data['total_amount'], 2310000)

    def test_negative_and_unformatted(self):
        data = extract_amounts('공급가액 -1000\n세액 -100')
        self.assertEqual(data['total_amount'], -1100)

    def test_discrepancy_not_overwritten(self):
        data = extract_amounts('공급가액 1000\n세액 100\n합계 1500')
        self.assertEqual(data['total_amount'], 1500)
        self.assertTrue(data['warnings'])

    def test_conflicting_values_need_review(self):
        data = extract_amounts('공급가액 1000\n공급가액 2000\n세액 100\n합계 1100')
        self.assertIsNone(data['amount'])

    def test_date_and_phone_not_money(self):
        data = extract_amounts('작성일자 2026-09-08\n연락처 010-1234-5678\nKRW 8,000,000')
        self.assertIsNone(data['amount'])

    def test_money_input(self):
        self.assertEqual(parse_money('-1,000'), -1000)
        self.assertIsNone(parse_money(''))
        with self.assertRaises(ValueError):
            parse_money('abc')


class DocumentTests(unittest.TestCase):
    def test_order_independent_pr_po(self):
        pr = {'pr_no': 'PR1', 'pr_title': 'PR title', 'amount': 1000, 'supplier': 'PR vendor'}
        po = {'pr_no': 'PO2', 'pr_title': 'PO title', 'amount': 2000, 'supplier': 'PO vendor'}
        a, b = DocumentState(), DocumentState()
        for kind, data in [('pr', pr), ('po', po)]:
            a.replace(kind, data)
        for kind, data in [('po', po), ('pr', pr)]:
            b.replace(kind, data)
        self.assertEqual(a.form(), b.form())
        self.assertEqual(a.form()['pr_no'], 'PR1')
        self.assertEqual(a.form()['amount'], 1000)
        self.assertEqual(a.supply('pr'), 1000)
        self.assertEqual(a.form()['supplier'], 'PO vendor')

    def test_tax_replacement_clears_values(self):
        state = DocumentState()
        state.replace('tax', {'date': '2026-01-01', 'supply_amount': 1000})
        state.replace('tax', {})
        self.assertEqual(state.form()['date'], '')
        self.assertIsNone(state.supply('tax'))

    def test_failed_pr_does_not_fall_back_to_po(self):
        state = DocumentState()
        state.replace('po', {'amount': 2000})
        state.replace('pr', {})
        self.assertIsNone(state.form()['amount'])


class FilesTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def pdf(self, name, pages=2):
        path = self.root / name
        writer = PdfWriter()
        for _ in range(pages):
            writer.add_blank_page(width=100, height=100)
        with path.open('wb') as stream:
            writer.write(stream)
        return str(path)

    def test_bad_pages_never_print(self):
        pdf = self.pdf('contract.pdf')
        with patch('printer_handler.subprocess.run') as run:
            for pages in ([98], [], [0, 98]):
                with self.subTest(pages=pages), self.assertRaises(ValueError):
                    printer_handler.print_pdf_file(pdf, 'MockPrinter', pages)
            run.assert_not_called()

    def test_print_timeout_does_not_fall_back(self):
        import subprocess
        pdf = self.pdf('invoice.pdf')
        with patch('printer_handler.subprocess.run', side_effect=subprocess.TimeoutExpired('mock', 15)), \
             patch('os.startfile') as start:
            with self.assertRaisesRegex(RuntimeError, '중복'):
                printer_handler.print_pdf_file(pdf, 'MockPrinter')
            start.assert_not_called()

    def test_same_named_outlook_attachments_have_distinct_contents(self):
        def attachment(content):
            return SimpleNamespace(FileName='same.pdf', SaveAsFile=lambda path: Path(path).write_bytes(content))
        items = [attachment(b'first'), attachment(b'second')]
        mail = SimpleNamespace(Attachments=SimpleNamespace(Count=2, Item=lambda n: items[n - 1]))
        outlook = SimpleNamespace(ActiveInspector=lambda: SimpleNamespace(CurrentItem=mail))
        com = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
        client = SimpleNamespace(client=SimpleNamespace(Dispatch=lambda _: outlook))
        with patch.object(outlook_handler, 'HAS_WIN32', True), \
             patch.object(outlook_handler, 'pythoncom', com, create=True), \
             patch.object(outlook_handler, 'win32com', client, create=True), \
             patch('file_utils.tempfile.gettempdir', return_value=self.tmp.name):
            files = outlook_handler.extract_from_active_com_mail()
        self.assertEqual(len(set(files)), 2)
        self.assertEqual([Path(p).read_bytes() for p in files], [b'first', b'second'])

    def test_real_pdf_text_and_encryption(self):
        from pypdf.generic import DictionaryObject, NameObject, DecodedStreamObject
        writer = PdfWriter()
        page = writer.add_blank_page(width=600, height=800)
        font = DictionaryObject({NameObject('/Type'): NameObject('/Font'), NameObject('/Subtype'): NameObject('/Type1'), NameObject('/BaseFont'): NameObject('/Helvetica')})
        page[NameObject('/Resources')] = DictionaryObject({NameObject('/Font'): DictionaryObject({NameObject('/F1'): writer._add_object(font)})})
        stream = DecodedStreamObject()
        lines = ['PR No. PR123', 'SUBJECT: Sample', 'Supply Amount 1,000,000', 'VAT 100,000', 'Grand Total 1,100,000']
        stream.set_data(('BT /F1 12 Tf 20 750 Td ' + ' '.join(f'({line}) Tj 0 -20 Td' for line in lines) + ' ET').encode('ascii'))
        page[NameObject('/Contents')] = writer._add_object(stream)
        writer.encrypt(pdf_parser.DEFAULT_PASSWORDS[0])
        path = self.root / 'encrypted.pdf'
        with path.open('wb') as output:
            writer.write(output)
        data = pdf_parser.parse_pr_pdf(str(path))
        self.assertEqual(data['pr_no'], 'PR123')
        self.assertEqual(data['amount'], 1000000)
        self.assertEqual(data['total_amount'], 1100000)

    def test_invalid_batch_preflight_prints_nothing(self):
        pdf = self.pdf('contract.pdf')
        with patch('printer_handler.print_pdf_file') as send:
            with self.assertRaises(ValueError):
                printer_handler.print_documents({'tax': pdf, 'contract': pdf}, 'Mock', [98])
            send.assert_not_called()

    def test_valid_batch_slices_actual_pdf(self):
        pdf = self.pdf('contract.pdf', 3)
        sizes = []
        with patch('printer_handler.print_pdf_file', side_effect=lambda p, _: sizes.append(len(PdfReader(p).pages)) or True):
            submitted, failed = printer_handler.print_documents({'contract': pdf}, 'Mock', [1, 2])
        self.assertEqual(sizes, [2])
        self.assertEqual(submitted, ['contract'])
        self.assertFalse(failed)

    def test_invalid_page_strings(self):
        for value in ('abc', '0', '1,', '3-1', '-1', '1-99999999'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_pages(value)
        self.assertEqual(parse_pages('1,3-4')[0], [0, 2, 3])

    def test_archive_same_names_and_repeated_calls(self):
        for folder, content in [('a', b'one'), ('b', b'two')]:
            (self.root / folder).mkdir()
            (self.root / folder / 'same.pdf').write_bytes(content)
        paths = {'pr': str(self.root / 'a/same.pdf'), 'po': str(self.root / 'b/same.pdf')}
        data = {'supplier': '../A:B', 'date': '2026/09/08'}
        first, files = excel_handler.archive_voucher_package(data, paths, base_dir=self.root / 'vault')
        second, _ = excel_handler.archive_voucher_package(data, paths, base_dir=self.root / 'vault')
        self.assertNotEqual(first, second)
        self.assertEqual([Path(p).read_bytes() for p in files], [b'one', b'two'])
        self.assertEqual(Path(first).parent, self.root / 'vault')

    def test_output_paths_unique_and_safe(self):
        with patch('file_utils.tempfile.gettempdir', return_value=self.tmp.name):
            a, b = temp_output('../same.pdf'), temp_output('../same.pdf')
        self.assertNotEqual(a, b)
        self.assertEqual(Path(a).parent.parent, self.root / 'VoucherPass')
        self.assertEqual(safe_name('CON'), 'document')

    def test_cp949_invoice(self):
        p = self.root / 'invoice.html'
        p.write_bytes('<div>작성일자 2026-08-24</div><div>공급가액 1,000,000</div><div>세액 0</div>'.encode('cp949'))
        with patch('pdf_parser._log_debug'):
            data = pdf_parser.parse_tax_invoice_data(str(p))
        self.assertEqual(data['date'], '2026-08-24')
        self.assertEqual(data['supply_amount'], 1000000)
        self.assertEqual(data['vat'], 0)

    def test_invalid_calendar_date(self):
        p = self.root / 'invoice.html'
        p.write_text('<div>작성일자 2026-02-31</div>', encoding='utf-8')
        with patch('pdf_parser._log_debug'):
            self.assertEqual(pdf_parser.parse_tax_invoice_data(str(p))['date'], '')

    def test_report_not_po(self):
        p = self.root / 'report.pdf'
        p.write_bytes(b'')
        with patch('pdf_parser.extract_pdf_text_safely', return_value='거래명세서'):
            self.assertEqual(pdf_parser.classify_pdf_type(str(p)), 'spec')

    def test_po_supplier_spacing(self):
        p = self.pdf('po.pdf')
        text = '2 . 공급처 ( SUPPLIER )\nCOMPANY NAME\n고려제강\n3. 납품'
        with patch('pdf_parser.extract_pdf_text_safely', return_value=text):
            self.assertEqual(pdf_parser.parse_pr_pdf(p)['supplier'], '고려제강')

    def test_excel_preserves_explicit_amounts(self):
        p = self.root / 'template.xlsx'
        wb = openpyxl.Workbook()
        wb.save(p)
        wb.close()
        data = dict(pr_no='PR1', pr_title='=1+1', amount=1000, vat=0, total_amount=1000)
        with patch('excel_handler.desktop_dir', return_value=self.root):
            output = excel_handler.generate_voucher_excel(data, str(p))
        wb = openpyxl.load_workbook(output)
        self.addCleanup(wb.close)
        self.assertEqual(wb.active['W9'].value, 1000)
        self.assertEqual(wb.active['W16'].value, 0)
        self.assertEqual(wb.active['D9'].data_type, 's')


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.root = SimpleNamespace(after=lambda *args: 'timer', after_cancel=lambda *args: None,
                                    report_callback_exception=lambda *args: self.fail(str(args)))
        self.runner = TaskRunner(self.root)
        self.addCleanup(self.runner.close)

    def test_old_results_ignored_after_reset(self):
        results = []
        old = object()
        self.runner.tokens['tax'] = old
        self.runner.events.put(('tax', old, results.append, 'old invoice'))
        self.runner.invalidate()
        self.runner._drain()
        self.assertEqual(results, [])

    def test_callback_runs_on_drain_thread(self):
        calls = []
        signal = threading.Event()
        self.runner.submit('pr', lambda: 42, lambda n: calls.append((n, threading.get_ident())), self.fail)
        # A queue get waits for actual work completion without timing assumptions.
        event = self.runner.events.get(timeout=5)
        self.runner.events.put(event)
        self.assertEqual(calls, [])
        self.runner._drain()
        self.assertEqual(calls, [(42, threading.get_ident())])

    def test_duplicate_work_rejected(self):
        gate = threading.Event()
        self.addCleanup(gate.set)
        self.assertTrue(self.runner.submit('outlook', lambda: gate.wait(5), lambda _: None, self.fail))
        self.assertFalse(self.runner.submit('outlook', lambda: None, lambda _: None, self.fail))
        gate.set()


if __name__ == '__main__':
    unittest.main()
