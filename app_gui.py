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
    대형 애플 글래스모피즘 드래그 앤 드롭 박스 컴포넌트
    """
    def __init__(self, parent, title, icon, file_var, on_file_selected=None, **kwargs):
        super().__init__(parent, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1.5, bd=0, **kwargs)
        self.file_var = file_var
        self.on_file_selected = on_file_selected

        self.inner_frame = tk.Frame(self, bg="#FFFFFF", padx=16, pady=22)
        self.inner_frame.pack(fill="both", expand=True)

        self.lbl_icon = tk.Label(self.inner_frame, text=icon, font=("Segoe UI Emoji", 26), bg="#FFFFFF", fg="#2563EB")
        self.lbl_icon.pack(side="top", pady=(4, 6))

        self.lbl_title = tk.Label(self.inner_frame, text=title, font=("Malgun Gothic", 12, "bold"), bg="#FFFFFF", fg="#0F172A")
        self.lbl_title.pack(side="top")

        self.lbl_status = tk.Label(self.inner_frame, text="PDF 파일을 이곳으로 드래그 앤 드롭 하거나 클릭하세요", font=("Malgun Gothic", 9), bg="#FFFFFF", fg="#64748B")
        self.lbl_status.pack(side="top", pady=(4, 4))

        for widget in (self, self.inner_frame, self.lbl_icon, self.lbl_title, self.lbl_status):
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
        self.config(bg="#EFF6FF", highlightbackground="#2563EB", highlightthickness=2.5)
        self.inner_frame.config(bg="#EFF6FF")
        self.lbl_icon.config(bg="#EFF6FF", fg="#1D4ED8")
        self.lbl_title.config(bg="#EFF6FF")
        self.lbl_status.config(bg="#EFF6FF", fg="#1D4ED8")

    def _on_drag_leave(self, event=None):
        bg_col = "#F0FDF4" if self.file_var.get() else "#FFFFFF"
        border_col = "#16A34A" if self.file_var.get() else "#CBD5E1"
        self.config(bg=bg_col, highlightbackground=border_col, highlightthickness=1.5)
        self.inner_frame.config(bg=bg_col)
        self.lbl_icon.config(bg=bg_col)
        self.lbl_title.config(bg=bg_col)
        self.lbl_status.config(bg=bg_col, fg="#15803D" if self.file_var.get() else "#64748B")

    def set_file(self, path):
        self.file_var.set(path)
        if self.on_file_selected:
            self.on_file_selected(path)

    def _update_ui_state(self, *args):
        path = self.file_var.get()
        if path and os.path.exists(path):
            fname = os.path.basename(path)
            self.lbl_status.config(text=f"✓ {fname}", fg="#15803D", font=("Malgun Gothic", 10, "bold"))
            self.config(bg="#F0FDF4", highlightbackground="#16A34A")
            self.inner_frame.config(bg="#F0FDF4")
            self.lbl_icon.config(bg="#F0FDF4", fg="#16A34A")
            self.lbl_title.config(bg="#F0FDF4")
        else:
            self.lbl_status.config(text="PDF 파일을 이곳으로 드래그 앤 드롭 하거나 클릭하세요", fg="#64748B", font=("Malgun Gothic", 9))
            self.config(bg="#FFFFFF", highlightbackground="#CBD5E1")
            self.inner_frame.config(bg="#FFFFFF")
            self.lbl_icon.config(bg="#FFFFFF", fg="#2563EB")
            self.lbl_title.config(bg="#FFFFFF")


class VoucherMachineApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Voucher Machine - Compact Layout Edition (v1.3.1)")
        self.root.geometry("1020x860")
        self.root.minsize(980, 800)

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
        self.bg_main = "#F1F5F9"
        self.root.configure(bg=self.bg_main)

    def _build_ui(self):
        header = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=20, pady=10)
        header.pack(fill="x")

        title_lbl = tk.Label(header, text="✨ Voucher Machine", font=("Malgun Gothic", 16, "bold"), bg="#FFFFFF", fg="#0F172A")
        title_lbl.pack(side="left")
        sub_lbl = tk.Label(header, text="  |  제출 서류 PDF Drag & Drop 업로드 ➔ 자동 파싱 & Voucher 작성/인쇄", font=("Malgun Gothic", 9), bg="#FFFFFF", fg="#64748B")
        sub_lbl.pack(side="left", pady=(3, 0))

        main_box = tk.Frame(self.root, bg=self.bg_main, padx=16, pady=12)
        main_box.pack(fill="both", expand=True)

        # -----------------------------------------------------------
        # Section 1: LARGE Drag & Drop Drop Zone (최대 크기 할당)
        # -----------------------------------------------------------
        sec1_card = tk.LabelFrame(main_box, text=" 📂 1. 제출 서류 PDF 업로드 (Drag & Drop Zone) ", font=("Malgun Gothic", 11, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=14, pady=12)
        sec1_card.pack(fill="both", expand=True, pady=(0, 10))

        grid_frame = tk.Frame(sec1_card, bg="#FFFFFF")
        grid_frame.pack(fill="both", expand=True)

        self.drop_pr = GlassDropZone(grid_frame, "① PR Print (구매요청서)", "📋", self.pr_pdf_path, on_file_selected=self.parse_uploaded_pdf)
        self.drop_pr.grid(row=0, column=0, padx=8, pady=6, sticky="nsew")

        self.drop_spec = GlassDropZone(grid_frame, "② 거래명세서 PDF", "📑", self.spec_pdf_path)
        self.drop_spec.grid(row=0, column=1, padx=8, pady=6, sticky="nsew")

        self.drop_tax = GlassDropZone(grid_frame, "③ 전자 세금계산서 PDF", "🧾", self.tax_pdf_path)
        self.drop_tax.grid(row=1, column=0, padx=8, pady=6, sticky="nsew")

        contract_container = tk.Frame(grid_frame, bg="#FFFFFF")
        contract_container.grid(row=1, column=1, padx=8, pady=6, sticky="nsew")

        self.drop_contract = GlassDropZone(contract_container, "④ 업체 계약서 PDF", "📜", self.contract_pdf_path)
        self.drop_contract.pack(fill="both", expand=True)

        pg_sub = tk.Frame(contract_container, bg="#F8FAFC", padx=8, pady=3, highlightbackground="#CBD5E1", highlightthickness=1)
        pg_sub.pack(fill="x", pady=(4, 0))
        tk.Label(pg_sub, text="📄 인쇄 대상 페이지:", font=("Malgun Gothic", 9, "bold"), bg="#F8FAFC", fg="#475569").pack(side="left")
        tk.Entry(pg_sub, textvariable=self.contract_page, font=("Malgun Gothic", 9), width=4, relief="solid", bd=1).pack(side="left", padx=4)
        tk.Label(pg_sub, text="(예: 1 또는 1,2)", font=("Malgun Gothic", 8), bg="#F8FAFC", fg="#94A3B8").pack(side="left")

        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)
        grid_frame.rowconfigure(0, weight=1)
        grid_frame.rowconfigure(1, weight=1)

        # -----------------------------------------------------------
        # Section 2: COMPACT Extracted Data Form (슬림 최소화)
        # -----------------------------------------------------------
        sec2_card = tk.LabelFrame(main_box, text=" 📝 2. 추출 데이터 (Voucher 자동 연동) ", font=("Malgun Gothic", 10, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=12, pady=6)
        sec2_card.pack(fill="x", pady=(0, 10))

        lbl_style = {"font": ("Malgun Gothic", 9, "bold"), "bg": "#FFFFFF", "fg": "#334155", "anchor": "e"}
        ent_style = {"font": ("Malgun Gothic", 9), "bg": "#F8FAFC", "relief": "solid", "bd": 1}

        r0 = tk.Frame(sec2_card, bg="#FFFFFF")
        r0.pack(fill="x", pady=2)
        
        tk.Label(r0, text="P/R No:", **lbl_style).pack(side="left")
        tk.Entry(r0, textvariable=self.pr_no_var, width=18, **ent_style).pack(side="left", padx=(4, 16))

        tk.Label(r0, text="작성일자:", **lbl_style).pack(side="left")
        tk.Entry(r0, textvariable=self.date_var, width=14, **ent_style).pack(side="left", padx=(4, 16))

        tk.Label(r0, text="거래처명 (Payee):", **lbl_style).pack(side="left")
        tk.Entry(r0, textvariable=self.supplier_var, width=22, **ent_style).pack(side="left", padx=(4, 0))

        r1 = tk.Frame(sec2_card, bg="#FFFFFF")
        r1.pack(fill="x", pady=2)
        tk.Label(r1, text="PR Title:", **lbl_style).pack(side="left")
        tk.Entry(r1, textvariable=self.pr_title_var, **ent_style).pack(side="left", fill="x", expand=True, padx=(4, 0))

        r2 = tk.Frame(sec2_card, bg="#FFFFFF")
        r2.pack(fill="x", pady=2)

        tk.Label(r2, text="공급가액:", **lbl_style).pack(side="left")
        amt_ent = tk.Entry(r2, textvariable=self.amount_var, width=16, **ent_style)
        amt_ent.pack(side="left", padx=(4, 16))
        amt_ent.bind("<KeyRelease>", self._recalc_amounts)

        tk.Label(r2, text="부가세(10%):", **lbl_style).pack(side="left")
        tk.Entry(r2, textvariable=self.vat_var, width=16, **ent_style).pack(side="left", padx=(4, 16))

        tk.Label(r2, text="합계금액:", **lbl_style).pack(side="left")
        tot_ent = tk.Entry(r2, textvariable=self.total_amount_var, width=18, font=("Malgun Gothic", 10, "bold"), bg="#EFF6FF", fg="#1D4ED8", relief="solid", bd=1)
        tot_ent.pack(side="left", padx=(4, 0))

        # -----------------------------------------------------------
        # Section 3 & Actions: ULTRA-SLIM BOTTOM BAR (최소화)
        # -----------------------------------------------------------
        bottom_bar = tk.Frame(main_box, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, padx=12, pady=8)
        bottom_bar.pack(fill="x")

        sett_box = tk.Frame(bottom_bar, bg="#FFFFFF")
        sett_box.pack(side="left", fill="x", expand=True, padx=(0, 10))

        s_r1 = tk.Frame(sett_box, bg="#FFFFFF")
        s_r1.pack(fill="x", pady=1)
        tk.Label(s_r1, text="템플릿:", font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#475569").pack(side="left")
        tk.Entry(s_r1, textvariable=self.template_path, font=("Malgun Gothic", 9), bg="#F8FAFC", relief="solid", bd=1).pack(side="left", fill="x", expand=True, padx=4)
        tk.Button(s_r1, text="변경", font=("Malgun Gothic", 8), bg="#E2E8F0", relief="flat", command=self.browse_template).pack(side="left")

        s_r2 = tk.Frame(sett_box, bg="#FFFFFF")
        s_r2.pack(fill="x", pady=1)
        tk.Label(s_r2, text="프린터:", font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#475569").pack(side="left")
        self.printer_combo = ttk.Combobox(s_r2, textvariable=self.selected_printer, font=("Malgun Gothic", 9), state="readonly")
        self.printer_combo.pack(side="left", fill="x", expand=True, padx=4)

        act_box = tk.Frame(bottom_bar, bg="#FFFFFF")
        act_box.pack(side="right")

        btn_excel = tk.Button(act_box, text="📊 Voucher 엑셀 작성", font=("Malgun Gothic", 10, "bold"), bg="#10B981", fg="white", activebackground="#059669", activeforeground="white", relief="flat", padx=14, pady=8, cursor="hand2", command=self.create_voucher_excel)
        btn_excel.pack(side="left", padx=(0, 6))

        btn_print = tk.Button(act_box, text="🖨️ Voucher & PDF 일괄 인쇄", font=("Malgun Gothic", 10, "bold"), bg="#4F46E5", fg="white", activebackground="#4338CA", activeforeground="white", relief="flat", padx=16, pady=8, cursor="hand2", command=self.print_all_documents)
        btn_print.pack(side="right")

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
