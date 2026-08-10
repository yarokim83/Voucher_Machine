import os
import sys
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
    TkinterDnD = tk

import pdf_parser
import excel_handler
import printer_handler

class GlassDropZone(tk.Frame):
    """
    애플 글래스모피즘 스타일 드래그 앤 드롭 박스 컴포넌트
    """
    def __init__(self, parent, title, icon, file_var, on_file_selected=None, **kwargs):
        super().__init__(parent, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, bd=0, **kwargs)
        self.file_var = file_var
        self.on_file_selected = on_file_selected

        self.inner_frame = tk.Frame(self, bg="#FFFFFF", padx=12, pady=10)
        self.inner_frame.pack(fill="both", expand=True)

        self.lbl_icon = tk.Label(self.inner_frame, text=icon, font=("Segoe UI Emoji", 18), bg="#FFFFFF", fg="#3B82F6")
        self.lbl_icon.pack(side="left", padx=(0, 8))

        info_box = tk.Frame(self.inner_frame, bg="#FFFFFF")
        info_box.pack(side="left", fill="both", expand=True)

        self.lbl_title = tk.Label(info_box, text=title, font=("Malgun Gothic", 10, "bold"), bg="#FFFFFF", fg="#1E293B", anchor="w")
        self.lbl_title.pack(fill="x")

        self.lbl_status = tk.Label(info_box, text="PDF 파일을 이곳에 드래그 앤 드롭 하거나 클릭하세요", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#94A3B8", anchor="w")
        self.lbl_status.pack(fill="x")

        for widget in (self, self.inner_frame, self.lbl_icon, info_box, self.lbl_title, self.lbl_status):
            widget.config(cursor="hand2")
            widget.bind("<Button-1>", self._browse_file)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._handle_drop)
            self.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            self.dnd_bind('<<DragLeave>>', self._on_drag_leave)

        self.file_var.trace_add("write", self._update_ui_state)

    def _browse_file(self, event=None):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if path:
            self.set_file(path)

    def _handle_drop(self, event):
        self._on_drag_leave()
        raw_path = event.data.strip()
        if raw_path.startswith('{') and raw_path.endswith('}'):
            raw_path = raw_path[1:-1]
        
        paths = re.findall(r'\{[^}]+\}|[^\s]+', raw_path)
        first_path = paths[0].strip('{}') if paths else raw_path

        if first_path.lower().endswith('.pdf') and os.path.exists(first_path):
            self.set_file(first_path)
        else:
            messagebox.showwarning("파일 형식 오류", "PDF 파일만 업로드할 수 있습니다.")

    def _on_drag_enter(self, event=None):
        self.config(bg="#EFF6FF", highlightbackground="#3B82F6", highlightthickness=2)
        self.inner_frame.config(bg="#EFF6FF")
        self.lbl_icon.config(bg="#EFF6FF")
        self.lbl_title.config(bg="#EFF6FF")
        self.lbl_status.config(bg="#EFF6FF", fg="#2563EB")

    def _on_drag_leave(self, event=None):
        bg_col = "#F0FDF4" if self.file_var.get() else "#FFFFFF"
        border_col = "#22C55E" if self.file_var.get() else "#CBD5E1"
        self.config(bg=bg_col, highlightbackground=border_col, highlightthickness=1)
        self.inner_frame.config(bg=bg_col)
        self.lbl_icon.config(bg=bg_col)
        self.lbl_title.config(bg=bg_col)
        self.lbl_status.config(bg=bg_col, fg="#64748B" if self.file_var.get() else "#94A3B8")

    def set_file(self, path):
        self.file_var.set(path)
        if self.on_file_selected:
            self.on_file_selected(path)

    def _update_ui_state(self, *args):
        path = self.file_var.get()
        if path and os.path.exists(path):
            fname = os.path.basename(path)
            self.lbl_status.config(text=f"✓ {fname}", fg="#15803D", font=("Malgun Gothic", 9, "bold"))
            self.config(bg="#F0FDF4", highlightbackground="#22C55E")
            self.inner_frame.config(bg="#F0FDF4")
            self.lbl_icon.config(bg="#F0FDF4", fg="#16A34A")
            self.lbl_title.config(bg="#F0FDF4")
        else:
            self.lbl_status.config(text="PDF 파일을 이곳에 드래그 앤 드롭 하거나 클릭하세요", fg="#94A3B8", font=("Malgun Gothic", 8))
            self.config(bg="#FFFFFF", highlightbackground="#CBD5E1")
            self.inner_frame.config(bg="#FFFFFF")
            self.lbl_icon.config(bg="#FFFFFF", fg="#3B82F6")
            self.lbl_title.config(bg="#FFFFFF")


class VoucherMachineApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Voucher Machine - Apple Glassmorphic Edition (v1.2.0)")
        self.root.geometry("1000x860")
        self.root.minsize(940, 780)

        self.pr_pdf_path = tk.StringVar()
        self.spec_pdf_path = tk.StringVar()   # 거래명세서
        self.tax_pdf_path = tk.StringVar()    # 세금계산서
        self.contract_pdf_path = tk.StringVar() # 계약서
        self.contract_page = tk.StringVar(value="1")

        self.template_path = tk.StringVar(value=r"C:\Users\baewoong.kim\Desktop\고려제강(2025).xlsx")
        self.selected_printer = tk.StringVar()

        self.pr_no_var = tk.StringVar()
        self.pr_title_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.vat_var = tk.StringVar()
        self.total_amount_var = tk.StringVar()
        self.date_var = tk.StringVar()
        self.supplier_var = tk.StringVar()

        self._setup_glass_theme()
        self._build_ui()
        self._load_printers()

    def _setup_glass_theme(self):
        self.bg_main = "#EEF2F6"
        self.glass_bg = "#FFFFFF"
        self.primary_blue = "#2563EB"
        self.root.configure(bg=self.bg_main)

    def _build_ui(self):
        header = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=25, pady=16)
        header.pack(fill="x")

        title_lbl = tk.Label(header, text="✨ Voucher Machine", font=("Malgun Gothic", 18, "bold"), bg="#FFFFFF", fg="#0F172A")
        title_lbl.pack(anchor="w")
        sub_lbl = tk.Label(header, text="Apple Glassmorphism UI  |  구매 서류 드래그 앤 드롭 ➔ 자동 파싱 & 엑셀 인쇄", font=("Malgun Gothic", 9.5), bg="#FFFFFF", fg="#64748B")
        sub_lbl.pack(anchor="w", pady=(2, 0))

        canvas = tk.Canvas(self.root, bg=self.bg_main, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        main_box = tk.Frame(canvas, bg=self.bg_main, padx=20, pady=16)

        main_box.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=main_box, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        sec1_card = tk.LabelFrame(main_box, text=" 📂 1. 제출 서류 Drag & Drop 업로드 ", font=("Malgun Gothic", 11, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=16, pady=14)
        sec1_card.pack(fill="x", pady=(0, 16))

        grid_frame = tk.Frame(sec1_card, bg="#FFFFFF")
        grid_frame.pack(fill="x")

        self.drop_pr = GlassDropZone(grid_frame, "① PR Print (구매요청서)", "📋", self.pr_pdf_path, on_file_selected=self.parse_uploaded_pdf)
        self.drop_pr.grid(row=0, column=0, padx=8, pady=8, sticky="nsew")

        self.drop_spec = GlassDropZone(grid_frame, "② 거래명세서", "📑", self.spec_pdf_path)
        self.drop_spec.grid(row=0, column=1, padx=8, pady=8, sticky="nsew")

        self.drop_tax = GlassDropZone(grid_frame, "③ 전자 세금계산서", "🧾", self.tax_pdf_path)
        self.drop_tax.grid(row=1, column=0, padx=8, pady=8, sticky="nsew")

        contract_container = tk.Frame(grid_frame, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1)
        contract_container.grid(row=1, column=1, padx=8, pady=8, sticky="nsew")

        self.drop_contract = GlassDropZone(contract_container, "④ 업체 계약서", "📜", self.contract_pdf_path)
        self.drop_contract.pack(fill="both", expand=True)

        pg_sub = tk.Frame(contract_container, bg="#F8FAFC", padx=10, pady=4)
        pg_sub.pack(fill="x")
        tk.Label(pg_sub, text="출력 페이지:", font=("Malgun Gothic", 8.5, "bold"), bg="#F8FAFC", fg="#475569").pack(side="left")
        tk.Entry(pg_sub, textvariable=self.contract_page, font=("Malgun Gothic", 9), width=5, relief="solid", bd=1).pack(side="left", padx=5)
        tk.Label(pg_sub, text="(예: 1 또는 1,2)", font=("Malgun Gothic", 8), bg="#F8FAFC", fg="#94A3B8").pack(side="left")

        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)

        sec2_card = tk.LabelFrame(main_box, text=" 📝 2. 파싱 데이터 확인 및 수정 (Voucher 연동) ", font=("Malgun Gothic", 11, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=18, pady=16)
        sec2_card.pack(fill="x", pady=(0, 16))

        lbl_style = {"font": ("Malgun Gothic", 9.5, "bold"), "bg": "#FFFFFF", "fg": "#334155", "anchor": "e"}
        ent_style = {"font": ("Malgun Gothic", 10.5), "bg": "#F8FAFC", "relief": "solid", "bd": 1}

        tk.Label(sec2_card, text="P/R No.:", **lbl_style).grid(row=0, column=0, padx=8, pady=8, sticky="e")
        tk.Entry(sec2_card, textvariable=self.pr_no_var, width=24, **ent_style).grid(row=0, column=1, padx=8, pady=8, sticky="w")

        tk.Label(sec2_card, text="작성일자 (Date):", **lbl_style).grid(row=0, column=2, padx=8, pady=8, sticky="e")
        tk.Entry(sec2_card, textvariable=self.date_var, width=24, **ent_style).grid(row=0, column=3, padx=8, pady=8, sticky="w")

        tk.Label(sec2_card, text="PR Title (품목명):", **lbl_style).grid(row=1, column=0, padx=8, pady=8, sticky="e")
        tk.Entry(sec2_card, textvariable=self.pr_title_var, width=68, **ent_style).grid(row=1, column=1, columnspan=3, padx=8, pady=8, sticky="w")

        tk.Label(sec2_card, text="Payee (거래처명):", **lbl_style).grid(row=2, column=0, padx=8, pady=8, sticky="e")
        tk.Entry(sec2_card, textvariable=self.supplier_var, width=24, **ent_style).grid(row=2, column=1, padx=8, pady=8, sticky="w")

        tk.Label(sec2_card, text="공급가액 (Amount):", **lbl_style).grid(row=3, column=0, padx=8, pady=8, sticky="e")
        amt_ent = tk.Entry(sec2_card, textvariable=self.amount_var, width=24, **ent_style)
        amt_ent.grid(row=3, column=1, padx=8, pady=8, sticky="w")
        amt_ent.bind("<KeyRelease>", self._recalc_amounts)

        tk.Label(sec2_card, text="V.A.T. (부가세 10%):", **lbl_style).grid(row=3, column=2, padx=8, pady=8, sticky="e")
        tk.Entry(sec2_card, textvariable=self.vat_var, width=24, **ent_style).grid(row=3, column=3, padx=8, pady=8, sticky="w")

        tk.Label(sec2_card, text="합계금액 (Total):", **lbl_style).grid(row=4, column=0, padx=8, pady=8, sticky="e")
        tot_ent = tk.Entry(sec2_card, textvariable=self.total_amount_var, width=24, font=("Malgun Gothic", 11, "bold"), bg="#EFF6FF", fg="#1D4ED8", relief="solid", bd=1)
        tot_ent.grid(row=4, column=1, padx=8, pady=8, sticky="w")

        sec3_card = tk.LabelFrame(main_box, text=" ⚙️ 3. Voucher 템플릿 & 인쇄 설정 ", font=("Malgun Gothic", 11, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=18, pady=14)
        sec3_card.pack(fill="x", pady=(0, 16))

        tmpl_row = tk.Frame(sec3_card, bg="#FFFFFF")
        tmpl_row.pack(fill="x", pady=4)
        tk.Label(tmpl_row, text="Voucher 엑셀 템플릿:", font=("Malgun Gothic", 9.5, "bold"), bg="#FFFFFF", fg="#334155", width=18, anchor="e").pack(side="left")
        tk.Entry(tmpl_row, textvariable=self.template_path, font=("Malgun Gothic", 9), bg="#F8FAFC", relief="solid", bd=1).pack(side="left", fill="x", expand=True, padx=8)
        tk.Button(tmpl_row, text="변경", font=("Malgun Gothic", 8.5), bg="#E2E8F0", fg="#334155", relief="flat", command=self.browse_template).pack(side="left")

        prt_row = tk.Frame(sec3_card, bg="#FFFFFF")
        prt_row.pack(fill="x", pady=4)
        tk.Label(prt_row, text="출력 프린터 선택:", font=("Malgun Gothic", 9.5, "bold"), bg="#FFFFFF", fg="#334155", width=18, anchor="e").pack(side="left")
        self.printer_combo = ttk.Combobox(prt_row, textvariable=self.selected_printer, font=("Malgun Gothic", 9.5), state="readonly")
        self.printer_combo.pack(side="left", fill="x", expand=True, padx=8)

        btn_box = tk.Frame(main_box, bg=self.bg_main)
        btn_box.pack(fill="x", pady=10)

        btn_excel = tk.Button(btn_box, text="📊 Voucher 엑셀 자동 작성", font=("Malgun Gothic", 11.5, "bold"), bg="#10B981", fg="white", activebackground="#059669", activeforeground="white", relief="flat", pady=12, cursor="hand2", command=self.create_voucher_excel)
        btn_excel.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_print = tk.Button(btn_box, text="🖨️ Voucher & PDF 서류 일괄 인쇄", font=("Malgun Gothic", 11.5, "bold"), bg="#4F46E5", fg="white", activebackground="#4338CA", activeforeground="white", relief="flat", pady=12, cursor="hand2", command=self.print_all_documents)
        btn_print.pack(side="right", fill="x", expand=True, padx=(10, 0))

    def _load_printers(self):
        printers = printer_handler.get_installed_printers()
        self.printer_combo['values'] = printers
        if printers:
            self.selected_printer.set(printers[0])

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.template_path.set(path)

    def parse_uploaded_pdf(self, pr_path=None):
        if not pr_path:
            pr_path = self.pr_pdf_path.get()

        if not pr_path or not os.path.exists(pr_path):
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

            messagebox.showinfo("파싱 완료", f"PR PDF 데이터 분석 성공!\n- PR No: {data.get('pr_no')}\n- 공급가액: {amt:,} 원")
        except Exception as e:
            messagebox.showerror("파싱 오류", f"PDF 데이터 자동 추출 실패:\n{e}")

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
            messagebox.showinfo("엑셀 생성 성공", f"Voucher 엑셀 서식이 정상 저장되었습니다!\n\n경로: {out_path}")
            return out_path
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 작성 실패:\n{e}")
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
            messagebox.showinfo("일괄 인쇄 요청 완료", f"다음 서류들이 프린터 [{printer}] 로 출력 요청되었습니다:\n\n- {summary}")

        except Exception as e:
            messagebox.showerror("인쇄 오류", f"일괄 인쇄 중 오류 발생:\n{e}")

def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = VoucherMachineApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
