import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pdf_parser
import excel_handler
import printer_handler

class VoucherMachineApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Voucher Machine - 구매 서류 처리 & Voucher 자동 생성/인쇄")
        self.root.geometry("960x780")
        self.root.minsize(900, 700)

        # 변수 정의
        self.pr_pdf_path = tk.StringVar()
        self.spec_pdf_path = tk.StringVar()   # 거래명세서
        self.tax_pdf_path = tk.StringVar()    # 세금계산서
        self.contract_pdf_path = tk.StringVar() # 계약서
        self.contract_page = tk.StringVar(value="1") # 계약서 출력 페이지

        self.template_path = tk.StringVar(value=r"C:\Users\baewoong.kim\Desktop\고려제강(2025).xlsx")
        self.selected_printer = tk.StringVar()

        # 파싱 데이터 변수
        self.pr_no_var = tk.StringVar()
        self.pr_title_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.vat_var = tk.StringVar()
        self.total_amount_var = tk.StringVar()
        self.date_var = tk.StringVar()
        self.supplier_var = tk.StringVar()

        self._setup_style()
        self._create_widgets()
        self._load_printers()

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use('clam')

        PRIMARY = "#1E3A8A"
        SECONDARY = "#2563EB"
        BG_COLOR = "#F8FAFC"
        CARD_BG = "#FFFFFF"

        self.root.configure(bg=BG_COLOR)

        style.configure("Header.TLabel", font=("Malgun Gothic", 16, "bold"), background=BG_COLOR, foreground=PRIMARY)
        style.configure("SubHeader.TLabel", font=("Malgun Gothic", 11, "bold"), background=CARD_BG, foreground="#334155")
        style.configure("Card.TLabelframe", background=CARD_BG, relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", font=("Malgun Gothic", 11, "bold"), foreground=PRIMARY)
        style.configure("Action.TButton", font=("Malgun Gothic", 10, "bold"), background=SECONDARY, foreground="white")
        style.map("Action.TButton", background=[("active", PRIMARY)])

    def _create_widgets(self):
        header_frame = ttk.Frame(self.root, padding=15)
        header_frame.pack(fill="x")
        ttk.Label(header_frame, text="📄 Voucher Machine - 구매 서류 일괄 자동화", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header_frame, text="PR/거래명세서/세금계산서 PDF 업로드 ➔ 데이터 추출 ➔ Voucher 엑셀 작성 ➔ 일괄 인쇄", font=("Malgun Gothic", 9), foreground="#64748B").pack(anchor="w", pady=(2, 0))

        main_canvas = tk.Canvas(self.root, bg="#F8FAFC", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas, padding=15)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 1. PDF 파일 업로드 섹션
        file_frame = ttk.LabelFrame(scrollable_frame, text=" 1. 제출 서류 PDF 업로드 ", style="Card.TLabelframe", padding=15)
        file_frame.pack(fill="x", pady=(0, 15))

        self._create_file_row(file_frame, "① PR Print (구매요청서):", self.pr_pdf_path, self.browse_pr_pdf)
        self._create_file_row(file_frame, "② 거래명세서 PDF:", self.spec_pdf_path, lambda: self.browse_file(self.spec_pdf_path))
        self._create_file_row(file_frame, "③ 전자 세금계산서 PDF:", self.tax_pdf_path, lambda: self.browse_file(self.tax_pdf_path))

        contract_sub_frame = ttk.Frame(file_frame)
        contract_sub_frame.pack(fill="x", pady=4)
        ttk.Label(contract_sub_frame, text="④ 업체 계약서 PDF:", font=("Malgun Gothic", 9, "bold"), width=22).pack(side="left")
        ttk.Entry(contract_sub_frame, textvariable=self.contract_pdf_path, font=("Malgun Gothic", 9)).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(contract_sub_frame, text="찾아보기", command=lambda: self.browse_file(self.contract_pdf_path)).pack(side="left", padx=2)
        
        ttk.Label(contract_sub_frame, text="출력 페이지:", font=("Malgun Gothic", 9)).pack(side="left", padx=(10, 2))
        ttk.Entry(contract_sub_frame, textvariable=self.contract_page, width=6, font=("Malgun Gothic", 9)).pack(side="left")

        btn_parse = tk.Button(file_frame, text="🔍 PDF 데이터 자동 분석", font=("Malgun Gothic", 10, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", command=self.parse_uploaded_pdf, relief="flat", pady=6)
        btn_parse.pack(fill="x", pady=(10, 0))

        # 2. 추출된 데이터 확인 & 수정 섹션
        data_frame = ttk.LabelFrame(scrollable_frame, text=" 2. 추출 데이터 확인 및 수정 (Voucher 자동 입력 항목) ", style="Card.TLabelframe", padding=15)
        data_frame.pack(fill="x", pady=(0, 15))

        grid_opts = {'padx': 8, 'pady': 6, 'sticky': 'w'}

        ttk.Label(data_frame, text="P/R No.:", font=("Malgun Gothic", 9, "bold")).grid(row=0, column=0, **grid_opts)
        ttk.Entry(data_frame, textvariable=self.pr_no_var, font=("Malgun Gothic", 10), width=25).grid(row=0, column=1, **grid_opts)

        ttk.Label(data_frame, text="Date (작성일):", font=("Malgun Gothic", 9, "bold")).grid(row=0, column=2, **grid_opts)
        ttk.Entry(data_frame, textvariable=self.date_var, font=("Malgun Gothic", 10), width=25).grid(row=0, column=3, **grid_opts)

        ttk.Label(data_frame, text="PR Title (품목/내용):", font=("Malgun Gothic", 9, "bold")).grid(row=1, column=0, **grid_opts)
        ttk.Entry(data_frame, textvariable=self.pr_title_var, font=("Malgun Gothic", 10), width=65).grid(row=1, column=1, columnspan=3, **grid_opts)

        ttk.Label(data_frame, text="Payee (거래처명):", font=("Malgun Gothic", 9, "bold")).grid(row=2, column=0, **grid_opts)
        ttk.Entry(data_frame, textvariable=self.supplier_var, font=("Malgun Gothic", 10), width=25).grid(row=2, column=1, **grid_opts)

        ttk.Label(data_frame, text="공급가액 (Amount):", font=("Malgun Gothic", 9, "bold")).grid(row=3, column=0, **grid_opts)
        amt_entry = ttk.Entry(data_frame, textvariable=self.amount_var, font=("Malgun Gothic", 10), width=25)
        amt_entry.grid(row=3, column=1, **grid_opts)
        amt_entry.bind("<KeyRelease>", self._recalc_amounts)

        ttk.Label(data_frame, text="V.A.T. (부가세 10%):", font=("Malgun Gothic", 9, "bold")).grid(row=3, column=2, **grid_opts)
        ttk.Entry(data_frame, textvariable=self.vat_var, font=("Malgun Gothic", 10), width=25).grid(row=3, column=3, **grid_opts)

        ttk.Label(data_frame, text="합계금액 (Total):", font=("Malgun Gothic", 9, "bold")).grid(row=4, column=0, **grid_opts)
        ttk.Entry(data_frame, textvariable=self.total_amount_var, font=("Malgun Gothic", 10, "bold"), width=25, foreground="#1E3A8A").grid(row=4, column=1, **grid_opts)

        # 3. 템플릿 설정 및 프린터 선택
        setting_frame = ttk.LabelFrame(scrollable_frame, text=" 3. Voucher 템플릿 & 프린터 설정 ", style="Card.TLabelframe", padding=15)
        setting_frame.pack(fill="x", pady=(0, 15))

        self._create_file_row(setting_frame, "Voucher 엑셀 템플릿:", self.template_path, lambda: self.browse_file(self.template_path, [("Excel Files", "*.xlsx")]))

        printer_sub = ttk.Frame(setting_frame)
        printer_sub.pack(fill="x", pady=4)
        ttk.Label(printer_sub, text="출력 프린터 선택:", font=("Malgun Gothic", 9, "bold"), width=22).pack(side="left")
        self.printer_combo = ttk.Combobox(printer_sub, textvariable=self.selected_printer, font=("Malgun Gothic", 9), state="readonly")
        self.printer_combo.pack(side="left", fill="x", expand=True, padx=5)

        # 4. 실행 버튼 그룹
        action_frame = ttk.Frame(scrollable_frame)
        action_frame.pack(fill="x", pady=10)

        btn_excel = tk.Button(action_frame, text="📊 Voucher 엑셀 자동 생성", font=("Malgun Gothic", 11, "bold"), bg="#059669", fg="white", activebackground="#047857", activeforeground="white", command=self.create_voucher_excel, relief="flat", pady=10, width=28)
        btn_excel.pack(side="left", padx=(0, 10), expand=True, fill="x")

        btn_print = tk.Button(action_frame, text="🖨️ Voucher & PDF 3종 일괄 인쇄", font=("Malgun Gothic", 11, "bold"), bg="#DC2626", fg="white", activebackground="#B91C1C", activeforeground="white", command=self.print_all_documents, relief="flat", pady=10, width=32)
        btn_print.pack(side="right", expand=True, fill="x")

    def _create_file_row(self, parent, label_text, var, command):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=4)
        ttk.Label(frame, text=label_text, font=("Malgun Gothic", 9, "bold"), width=22).pack(side="left")
        ttk.Entry(frame, textvariable=var, font=("Malgun Gothic", 9)).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(frame, text="찾아보기", command=command).pack(side="left")

    def _load_printers(self):
        printers = printer_handler.get_installed_printers()
        self.printer_combo['values'] = printers
        if printers:
            self.selected_printer.set(printers[0])

    def browse_file(self, var, filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    def browse_pr_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if path:
            self.pr_pdf_path.set(path)
            self.parse_uploaded_pdf()

    def parse_uploaded_pdf(self):
        pr_path = self.pr_pdf_path.get()
        if not pr_path or not os.path.exists(pr_path):
            messagebox.showwarning("경고", "분석할 PR Print (구매요청서) PDF 파일을 먼저 선택하세요.")
            return

        try:
            data = pdf_parser.parse_pr_pdf(pr_path)
            self.pr_no_var.set(data.get('pr_no', ''))
            self.pr_title_var.set(data.get('pr_title', ''))
            
            amt = data.get('amount', 0)
            self.amount_var.set(f"{amt:,}" if amt else "0")
            
            vat = data.get('vat', 0)
            self.vat_var.set(f"{vat:,}" if vat else "0")
            
            tot = data.get('total_amount', 0)
            self.total_amount_var.set(f"{tot:,}" if tot else "0")
            
            self.date_var.set(data.get('date', ''))
            self.supplier_var.set(data.get('supplier', ''))

            messagebox.showinfo("성공", "PR PDF 데이터 분석이 완료되었습니다!")
        except Exception as e:
            messagebox.showerror("오류", f"PDF 데이터 파싱 실패:\n{e}")

    def _recalc_amounts(self, event=None):
        raw = self.amount_var.get().replace(',', '').strip()
        if raw.isdigit():
            amt = int(raw)
            vat = int(amt * 0.1)
            tot = amt + vat
            self.vat_var.set(f"{vat:,}")
            self.total_amount_var.set(f"{tot:,}")

    def _get_form_data(self):
        def _to_int(val_str):
            cleaned = val_str.replace(',', '').strip()
            return int(cleaned) if cleaned.isdigit() else 0

        return {
            'pr_no': self.pr_no_var.get().strip(),
            'pr_title': self.pr_title_var.get().strip(),
            'amount': _to_int(self.amount_var.get()),
            'vat': _to_int(self.vat_var.get()),
            'total_amount': _to_int(self.total_amount_var.get()),
            'date': self.date_var.get().strip(),
            'supplier': self.supplier_var.get().strip(),
        }

    def create_voucher_excel(self):
        data = self._get_form_data()
        tmpl = self.template_path.get()

        if not tmpl or not os.path.exists(tmpl):
            messagebox.showerror("오류", f"Voucher 엑셀 템플릿 파일을 찾을 수 없습니다:\n{tmpl}")
            return None

        try:
            out_path = excel_handler.generate_voucher_excel(data, template_path=tmpl)
            messagebox.showinfo("엑셀 생성 완료", f"Voucher 엑셀 파일이 정상적으로 저장되었습니다:\n\n{out_path}")
            return out_path
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 생성 실패:\n{e}")
            return None

    def print_all_documents(self):
        excel_path = self.create_voucher_excel()
        if not excel_path:
            return

        printer = self.selected_printer.get()
        printed_list = []

        try:
            printer_handler.print_excel_file(excel_path, printer_name=printer)
            printed_list.append("Voucher 엑셀 서식")

            if self.pr_pdf_path.get() and os.path.exists(self.pr_pdf_path.get()):
                printer_handler.print_pdf_file(self.pr_pdf_path.get(), printer_name=printer)
                printed_list.append("PR Print PDF")

            if self.spec_pdf_path.get() and os.path.exists(self.spec_pdf_path.get()):
                printer_handler.print_pdf_file(self.spec_pdf_path.get(), printer_name=printer)
                printed_list.append("거래명세서 PDF")

            if self.tax_pdf_path.get() and os.path.exists(self.tax_pdf_path.get()):
                printer_handler.print_pdf_file(self.tax_pdf_path.get(), printer_name=printer)
                printed_list.append("전자 세금계산서 PDF")

            if self.contract_pdf_path.get() and os.path.exists(self.contract_pdf_path.get()):
                page_str = self.contract_page.get().strip()
                pages = []
                for p in page_str.split(','):
                    if p.strip().isdigit():
                        pages.append(int(p.strip()) - 1)
                if not pages:
                    pages = [0]
                
                printer_handler.print_pdf_file(self.contract_pdf_path.get(), printer_name=printer, page_range=pages)
                printed_list.append(f"업체 계약서 PDF ({page_str} 페이지)")

            summary = "\n- ".join(printed_list)
            messagebox.showinfo("일괄 인쇄 완료", f"다음 문서들이 프린터[{printer}]로 인쇄 요청되었습니다:\n\n- {summary}")

        except Exception as e:
            messagebox.showerror("인쇄 오류", f"문서 일괄 인쇄 중 오류 발생:\n{e}")

def main():
    root = tk.Tk()
    app = VoucherMachineApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
