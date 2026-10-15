"""Synthetic purchase-form regressions; no business documents in the repository."""
import unittest
from unittest.mock import patch
from voucher_core import extract_purchase_amounts
import pdf_parser


def pr_table(total='Total KRW 500,000 KRW 1,500,000', basis='Excluding VAT'):
    return f'PURCHASE REQUISITION\nCONTRACT LIST (Unit:KRW / {basis})\n단가/월간금액 총약정금액\n{total}\n⊙ BUDGET\nTotal KRW 99,000,000'


class PurchaseAmountTests(unittest.TestCase):
    def test_amount_wraps_after_unit_price(self):
        data = extract_purchase_amounts(pr_table('Total KRW 500,000\n1,500,000'))
        self.assertEqual(data['amount'], 1500000)

    def test_wrapped_currency_and_smaller_amount(self):
        data = extract_purchase_amounts(pr_table('Total KRW 9,000,000\nKRW\n1,500,000'))
        self.assertEqual(data['amount'], 1500000)

    def test_amount_column_overrides_text_order(self):
        tables = [[['No', 'Amount', 'Unit Price'], ['Total', 'KRW\n1,500,000', 'KRW 500,000']]]
        totals = pdf_parser._purchase_totals_from_tables(tables)
        data = extract_purchase_amounts(pr_table('Total KRW 1,500,000 KRW 500,000'), totals)
        self.assertEqual(data['amount'], 1500000)

    def test_korean_amount_column(self):
        tables = [[['단가/월간금액', '총약정금액', ''], ['500,000', 'KRW\n1,500,000', 'Total']]]
        self.assertEqual(pdf_parser._purchase_totals_from_tables(tables), [1500000])

    def test_unreadable_amount_cell_does_not_use_unit_price(self):
        tables = [[['Unit Price', 'Amount', ''], ['500,000', '', 'Total']]]
        data = extract_purchase_amounts(pr_table('Total KRW 500,000'), pdf_parser._purchase_totals_from_tables(tables))
        self.assertIsNone(data['amount'])
        self.assertTrue(data['warnings'])

    def test_conflicting_table_totals_need_review(self):
        data = extract_purchase_amounts(pr_table(), [1500000, 2000000])
        self.assertIsNone(data['amount'])

    def test_last_total_column_is_net_not_largest_number(self):
        data = extract_purchase_amounts(pr_table('Total KRW 9,000,000 KRW 1,500,000'))
        self.assertEqual((data['amount'], data['vat'], data['total_amount']), (1500000, 150000, 1650000))
        self.assertEqual(data['vat_source'], 'calculated')
        self.assertEqual(data['vat_rate_percent'], 10)

    def test_wrapped_po_total(self):
        text = 'PURCHASING ORDER\n3. 구매 리스트\n(Unit:KRW / Excluding VAT)\n(ORDER LIST)\nKRW KRW\nTotal\n1,500,000 1,500,000\n4. 계약 조건(TERMS AND CONDITIONS)'
        self.assertEqual(extract_purchase_amounts(text)['amount'], 1500000)

    def test_currency_on_separate_line(self):
        data = extract_purchase_amounts(pr_table('Total\nKRW KRW\n500,000 1,500,000'))
        self.assertEqual(data['amount'], 1500000)

    def test_single_column_exclusive_total(self):
        data = extract_purchase_amounts(pr_table('Total KRW 1,500,000'))
        self.assertEqual(data['total_amount'], 1650000)

    def test_inclusive_table_does_not_apply_exclusive_rule(self):
        data = extract_purchase_amounts(pr_table(basis='Including VAT'))
        self.assertIsNone(data['amount'])
        self.assertNotIn('vat_source', data)

    def test_missing_total_does_not_scan_budget(self):
        self.assertIsNone(extract_purchase_amounts(pr_table('1 UT KRW 500,000 KRW 1,500,000'))['amount'])

    def test_conflicting_totals_not_guessed(self):
        data = extract_purchase_amounts(pr_table('Total KRW 1000\nTotal KRW 2000'))
        self.assertIsNone(data['amount'])
        self.assertTrue(data['warnings'])

    def test_explicit_vat_is_preserved(self):
        data = extract_purchase_amounts(pr_table('Total KRW 1,500,000\nVAT 0'))
        self.assertEqual(data['vat'], 0)
        self.assertEqual(data['vat_source'], 'extracted')
        self.assertEqual(data['total_amount'], 1500000)

    def test_rate_is_not_parsed_as_amount(self):
        data = extract_purchase_amounts(pr_table() + '\nVAT rate: 5%')
        self.assertEqual(data['vat'], 75000)
        self.assertEqual(data['vat_rate_percent'], 5)

    def test_conflicting_explicit_vat_not_replaced_with_ten_percent(self):
        data = extract_purchase_amounts(pr_table() + '\nVAT 100\nVAT 200')
        self.assertIsNone(data['vat'])
        self.assertTrue(data['warnings'])

    def test_exempt_needs_confirmation(self):
        data = extract_purchase_amounts(pr_table() + '\n면세 대상')
        self.assertEqual(data['amount'], 1500000)
        self.assertIsNone(data['vat'])

    def test_non_purchase_document_not_subject_to_default_rate(self):
        data = extract_purchase_amounts(pr_table().replace('PURCHASE REQUISITION', 'INVOICE'))
        self.assertIsNone(data['vat'])


class PurchaseSupplierTests(unittest.TestCase):
    def test_pr_company_column_excludes_person_and_contacts(self):
        tables = [[['No', 'Company', 'Name', 'Title', 'TEL.', 'FAX.'],
                   ['1', '테스트기술', '홍길동', '이사', '051 123 4567', '051 123 4568']]]
        self.assertEqual(pdf_parser._supplier_from_tables(tables), '테스트기술')

    def test_po_combined_header_cell(self):
        tables = [[['COMPANY NAME\n테스트기술', 'NAME', 'TEL.\n051 123 4567']]]
        self.assertEqual(pdf_parser._supplier_from_tables(tables), '테스트기술')

    def test_fallback_removes_spaced_phone_without_removing_brand_digits(self):
        text = '2. 공급처(SUPPLIER)\nCOMPANY NAME\n3M코리아 051 123 4567 051 123 4568'
        with patch('pdf_parser.extract_pdf_text_safely', return_value=text), patch('pdf_parser._read_purchase_supplier', return_value=''):
            self.assertEqual(pdf_parser.parse_pr_pdf(__file__)['supplier'], '3M코리아')

    def test_fallback_removes_hyphenated_contacts(self):
        text = '2. 공급처(SUPPLIER)\nCOMPANY NAME\n제일2공업 02-1234-5678'
        with patch('pdf_parser.extract_pdf_text_safely', return_value=text), patch('pdf_parser._read_purchase_supplier', return_value=''):
            self.assertEqual(pdf_parser.parse_pr_pdf(__file__)['supplier'], '제일2공업')
