"""Hidden Tk smoke tests. No Outlook, printer, tray or clipboard interaction."""
import tkinter as tk
import unittest
from unittest.mock import patch
from pathlib import Path
import tempfile
import app_gui


class GuiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = app_gui.TkinterDnD.Tk() if app_gui.HAS_DND else tk.Tk()
        except tk.TclError as error:
            self.skipTest(f'Tk runtime unavailable: {error}')
        self.root.withdraw()
        with patch.object(app_gui.VoucherPassApp, '_setup_system_tray'), \
             patch.object(app_gui.VoucherPassApp, '_setup_global_hotkey'), \
             patch.object(app_gui.VoucherPassApp, '_load_printers'):
            self.app = app_gui.VoucherPassApp(self.root)
        self.addCleanup(self.app.close_app)

    def test_reset_clears_all(self):
        self.app.documents.replace('pr', dict(pr_no='PR1', pr_title='title', amount=1000, vat=100, total_amount=1100))
        self.app._refresh_form()
        self.assertEqual(self.app.amount_var.get(), '1,000')
        self.app.reset_all_data()
        self.assertEqual(self.app.amount_var.get(), '')
        self.assertEqual(self.app.txt_pr_title.get('1.0', 'end-1c'), '')
        self.assertEqual(self.app.documents.documents, {})

    def test_manual_amount_does_not_change_original(self):
        self.app.documents.replace('pr', dict(amount=1000, vat=100, total_amount=1100))
        self.app._refresh_form()
        self.app.amount_var.set('2000')
        self.app._recalc_amounts()
        self.assertEqual(self.app.pr_supply_amount, 1000)
        self.assertEqual(self.app.total_amount_var.get(), '2,100')

    def test_auto_registration_calls_parser_once(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'tax.pdf'
            path.write_bytes(b'')
            with patch.object(self.app.tasks, 'submit') as submit:
                self.assertTrue(self.app.handle_auto_detected_pdf(str(path), 'tax'))
                self.assertEqual(submit.call_count, 1)

    def test_stale_completion_after_reset(self):
        token = object()
        self.app.tasks.tokens['tax'] = token
        self.app.tasks.events.put(('tax', token, self.app._apply_tax_data,
                                   {'date': '2026-09-08', 'supply_amount': 1000}))
        self.app.reset_all_data()
        self.app.tasks._drain()
        self.assertEqual(self.app.date_var.get(), '')

    def test_pr_po_callbacks_have_distinct_owners(self):
        with patch.object(self.app, 'parse_uploaded_pdf') as parse:
            self.app.drop_pr.on_file_selected('a.pdf')
            self.app.drop_po.on_file_selected('b.pdf')
        self.assertEqual([call.args for call in parse.call_args_list], [('a.pdf', 'pr'), ('b.pdf', 'po')])

    def test_calculated_vat_label_and_reset(self):
        self.app.documents.replace('pr', dict(amount=1500000, vat=150000, total_amount=1650000,
                                              vat_source='calculated', vat_rate_percent=10))
        self.app._refresh_form()
        self.assertEqual(self.app.vat_var.get(), '150,000')
        self.assertIn('10% 계산', self.app.lbl_vat.cget('text'))
        self.app.reset_all_data()
        self.assertEqual(self.app.lbl_vat.cget('text'), '• 부가세')


if __name__ == '__main__':
    unittest.main()
