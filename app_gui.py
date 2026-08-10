import os
import sys
import re
import urllib.parse
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

class PastelGlassDropZone(tk.Frame):
    """
    VoucherPass 파스텔 원형 아이콘 글래스모피즘 Drop Zone
    """
    def __init__(self, parent, title, icon, bg_circle_color, icon_color, file_var, on_file_selected=None, **kwargs):
        super().__init__(parent, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, bd=0, **kwargs)
        self.file_var = file_var
        self.on_file_selected = on_file_selected
        self.bg_circle_color = bg_circle_color
        self.icon_color = icon_color

        self.inner = tk.Frame(self, bg="#FFFFFF", padx=16, pady=20)
        self.inner.pack(fill="both", expand=True)

        self.icon_bg = tk.Frame(self.inner, bg=bg_circle_color, padx=12, pady=10)
        self.icon_bg.pack(side="left", padx=(0, 14))

        self.lbl_icon = tk.Label(self.icon_bg, text=icon, font=("Segoe UI Emoji", 20), bg=bg_circle_color, fg=icon_color)
        self.lbl_icon.pack()

        txt_box = tk.Frame(self.inner, bg="#FFFFFF")
        txt_box.pack(side="left", fill="both", expand=True)

        self.lbl_title = tk.Label(txt_box, text=title, font=("Malgun Gothic", 11, "bold"), bg="#FFFFFF", fg="#0F172A", anchor="w")
        self.lbl_title.pack(fill="x")

        self.lbl_status = tk.Label(txt_box, text="PDF 파일을 이곳으로 드래그 앤 드롭\n하거나 클릭하세요", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#94A3B8", anchor="w", justify="left")
        self.lbl_status.pack(fill="x", pady=(2, 0))

        for w in (self, self.inner, self.icon_bg, self.lbl_icon, txt_box, self.lbl_title, self.lbl_status):
            w.config(cursor="hand2")
            w.bind("<Button-1>", self._browse_file)

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
        raw_data = event.data if event and hasattr(event, 'data') else ''
        
        cleaned_data = urllib.parse.unquote(str(raw_data)).strip()
        if cleaned_data.startswith('file:///'):
            cleaned_data = cleaned_data[8:]

        matches = re.findall(r'\{([^}]+)\}|(\S+)', cleaned_data)
        valid_pdf = None

        for m in matches:
            p = m[0] if m[0] else m[1]
            p = p.strip('\"\'{} \t\r\n')
            if p:
                norm_p = os.path.abspath(os.path.normpath(p))
                if norm_p.lower().endswith('.pdf') and os.path.exists(norm_p):
                    valid_pdf = norm_p
                    break

        if valid_pdf:
            self.set_file(valid_pdf)
        else:
            messagebox.showwarning("파일 형식 오류", "올바른 PDF 파일을 업로드해 주세요.")

    def _on_drag_enter(self, event=None):
        self.config(bg="#F0F9FF", highlightbackground="#0284C7", highlightthickness=2)
        self.inner.config(bg="#F0F9FF")

    def _on_drag_leave(self, event=None):
        is_set = bool(self.file_var.get())
        bg_col = "#F0FDF4" if is_set else "#FFFFFF"
        border_col = "#16A34A" if is_set else "#E2E8F0"
        self.config(bg=bg_col, highlightbackground=border_col, highlightthickness=1)
        self.inner.config(bg=bg_col)

    def set_file(self, path):
        self.file_var.set(path)
        if self.on_file_selected:
            self.on_file_selected(path)

    def _update_ui_state(self, *args):
        path = self.file_var.get()
        if path and os.path.exists(path):
            fname = os.path.basename(path)
            self.lbl_status.config(text=f"✓ {fname}", fg="#15803D", font=("Malgun Gothic", 9, "bold"))
            self.config(bg="#F0FDF4", highlightbackground="#16A34A")
            self.inner.config(bg="#F0FDF4")
        else:
            self.lbl_status.config(text="PDF 파일을 이곳으로 드래그 앤 드롭\n하거나 클릭하세요", fg="#94A3B8", font=("Malgun Gothic", 8))
            self.config(bg="#FFFFFF", highlightbackground="#E2E8F0")
            self.inner.config(bg="#FFFFFF")


class VoucherPassApp:
    def __init__(self, root):
        self.root = root
        self.root.title("VoucherPass v1.5.0")
        self.root.geometry("1180x800")
        self.root.minsize(1120, 740)

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

        self._setup_style()
        self._build_layout()
        self._load_printers()

    def _setup_style(self):
        self.bg_app = "#F8FAFC"
        self.sidebar_bg = "#F1F5F9"
        self.root.configure(bg=self.bg_app)

    def _build_layout(self):
        # Header
        header = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=16, pady=10)
        header.pack(fill="x")

        traffic_frame = tk.Frame(header, bg="#FFFFFF")
        traffic_frame.pack(side="left", padx=(0, 12))

        for color in ("#FF5F56", "#FFBD2E", "#27C93F"):
            dot = tk.Label(traffic_frame, text="●", font=("Arial", 10), bg="#FFFFFF", fg=color)
            dot.pack(side="left", padx=2)

        logo_icon = tk.Label(header, text="☑", font=("Segoe UI Emoji", 15, "bold"), bg="#FFFFFF", fg="#2563EB")
        logo_icon.pack(side="left", padx=(6, 4))

        logo_txt = tk.Label(header, text="VoucherPass", font=("Malgun Gothic", 15, "bold"), bg="#FFFFFF", fg="#0F172A")
        logo_txt.pack(side="left")

        ver_badge = tk.Label(header, text="v1.5.0", font=("Malgun Gothic", 8, "bold"), bg="#EFF6FF", fg="#2563EB", padx=6, pady=2)
        ver_badge.pack(side="left", padx=(8, 16))

        sub_desc = tk.Label(header, text="⚙️ 제출 서류 PDF Drag & Drop 업로드 ➔ 자동 파싱 & 원클릭 클립보드 복사", font=("Malgun Gothic", 9), bg="#FFFFFF", fg="#64748B")
        sub_desc.pack(side="left")

        btn_help = tk.Button(header, text="❓ 도움말", font=("Malgun Gothic", 8), bg="#F1F5F9", fg="#475569", relief="flat", padx=10, pady=3, command=lambda: messagebox.showinfo("도움말", "작성일자, PR Title, 공급가액 옆의 [📋 복사] 버튼을 누르면 클립보드에 복사되어 엑셀 등에 바로 붙여넣을 수 있습니다."))
        btn_help.pack(side="right", padx=2)

        btn_sett = tk.Button(header, text="⚙️ 설정", font=("Malgun Gothic", 8), bg="#F1F5F9", fg="#475569", relief="flat", padx=10, pady=3, command=self.browse_template)
        btn_sett.pack(side="right", padx=2)

        body_container = tk.Frame(self.root, bg=self.bg_app)
        body_container.pack(fill="both", expand=True)

        # Left Sidebar
        sidebar = tk.Frame(body_container, bg=self.sidebar_bg, width=210, padx=12, pady=16)
        sidebar.pack(side="left", fill="y")

        nav_items = [
            ("☁️ PDF 업로드", True),
            ("📋 데이터 복사", True),
            ("🖨️ 작성/인쇄", False),
            ("🕒 히스토리", False),
            ("⚙️ 설정", False),
        ]

        for text, active in nav_items:
            bg_c = "#EFF6FF" if active else self.sidebar_bg
            fg_c = "#2563EB" if active else "#475569"
            font_w = "bold" if active else "normal"
            
            btn = tk.Button(sidebar, text=text, font=("Malgun Gothic", 9, font_w), bg=bg_c, fg=fg_c, activebackground="#DBEAFE", activeforeground="#1E40AF", relief="flat", anchor="w", padx=14, pady=9, cursor="hand2")
            btn.pack(fill="x", pady=2)

        safe_card = tk.Frame(sidebar, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=12, pady=12)
        safe_card.pack(side="bottom", fill="x", pady=(0, 10))

        tk.Label(safe_card, text="🛡️ 안전한 데이터 처리", font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#0F172A", anchor="w").pack(fill="x")
        tk.Label(safe_card, text="모든 파일은 안전하게\n처리되며 저장되지 않습니다.", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#94A3B8", justify="left", anchor="w").pack(fill="x", pady=(4, 0))

        # Right Main Panel
        main_panel = tk.Frame(body_container, bg=self.bg_app, padx=18, pady=14)
        main_panel.pack(side="right", fill="both", expand=True)

        # --- Section 1: Drag & Drop Drop Zone ---
        sec1_header = tk.Frame(main_panel, bg=self.bg_app)
        sec1_header.pack(fill="x", pady=(0, 8))

        tk.Label(sec1_header, text="1. 제출 서류 PDF 업로드 (Drag & Drop Zone)", font=("Malgun Gothic", 11, "bold"), bg=self.bg_app, fg="#0F172A").pack(side="left")

        btn_select_file = tk.Button(sec1_header, text="📁 파일 선택", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#334155", relief="flat", highlightbackground="#E2E8F0", highlightthickness=1, padx=10, pady=3, command=lambda: self.drop_pr._browse_file())
        btn_select_file.pack(side="right")

        grid_drop = tk.Frame(main_panel, bg=self.bg_app)
        grid_drop.pack(fill="both", expand=True, pady=(0, 10))

        self.drop_pr = PastelGlassDropZone(grid_drop, "① PR Print (구매요청서)", "🛒", "#EFF6FF", "#2563EB", self.pr_pdf_path, on_file_selected=self.parse_uploaded_pdf)
        self.drop_pr.grid(row=0, column=0, padx=6, pady=6, sticky="nsew")

        self.drop_spec = PastelGlassDropZone(grid_drop, "② 거래명세서 PDF", "📄", "#ECFDF5", "#059669", self.spec_pdf_path)
        self.drop_spec.grid(row=0, column=1, padx=6, pady=6, sticky="nsew")

        self.drop_tax = PastelGlassDropZone(grid_drop, "③ 전자 세금계산서 PDF", "🧾", "#F5F3FF", "#7C3AED", self.tax_pdf_path)
        self.drop_tax.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")

        contract_container = tk.Frame(grid_drop, bg=self.bg_app)
        contract_container.grid(row=1, column=1, padx=6, pady=6, sticky="nsew")

        self.drop_contract = PastelGlassDropZone(contract_container, "④ 업체 계약서 PDF", "📝", "#FFFBEB", "#D97706", self.contract_pdf_path)
        self.drop_contract.pack(fill="both", expand=True)

        pg_sub = tk.Frame(contract_container, bg="#FFFFFF", padx=8, pady=3, highlightbackground="#E2E8F0", highlightthickness=1)
        pg_sub.pack(fill="x", pady=(3, 0))
        tk.Label(pg_sub, text="📄 인쇄 대상 페이지:", font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#475569").pack(side="left")
        tk.Entry(pg_sub, textvariable=self.contract_page, font=("Malgun Gothic", 8), width=4, relief="solid", bd=1).pack(side="left", padx=4)
        tk.Label(pg_sub, text="(예: 1 또는 1,2)", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#94A3B8").pack(side="left")

        grid_drop.columnconfigure(0, weight=1)
        grid_drop.columnconfigure(1, weight=1)
        grid_drop.rowconfigure(0, weight=1)
        grid_drop.rowconfigure(1, weight=1)

        # -----------------------------------------------------------
        # Section 2: Extracted Data & COPY BUTTONS (복사 기능 강조)
        # -----------------------------------------------------------
        sec2_card = tk.LabelFrame(main_panel, text=" 📝 2. 추출 데이터 & 원클릭 복사 (필요 항목만 복사하여 붙여넣기) ", font=("Malgun Gothic", 10, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=12, pady=10)
        sec2_card.pack(fill="x", pady=(0, 10))

        lbl_s = {"font": ("Malgun Gothic", 8.5, "bold"), "bg": "#FFFFFF", "fg": "#334155", "anchor": "e"}
        ent_s = {"font": ("Malgun Gothic", 9), "bg": "#F8FAFC", "relief": "solid", "bd": 1}
        btn_copy_s = {"font": ("Malgun Gothic", 8, "bold"), "bg": "#EFF6FF", "fg": "#2563EB", "activebackground": "#DBEAFE", "activeforeground": "#1E40AF", "relief": "solid", "bd": 1, "padx": 6, "pady": 1, "cursor": "hand2"}

        # Row 0: P/R No & 작성일자 (복사 버튼 1) & 거래처명
        r0 = tk.Frame(sec2_card, bg="#FFFFFF")
        r0.pack(fill="x", pady=3)

        tk.Label(r0, text="P/R No:", **lbl_s).pack(side="left")
        tk.Entry(r0, textvariable=self.pr_no_var, width=16, **ent_s).pack(side="left", padx=(4, 12))

        # 1) 작성일자 + 복사 버튼
        tk.Label(r0, text="📅 작성일자:", **lbl_s).pack(side="left")
        tk.Entry(r0, textvariable=self.date_var, width=13, **ent_s).pack(side="left", padx=(4, 2))
        tk.Button(r0, text="📋 복사", command=lambda: self.copy_to_clipboard(self.date_var.get(), "작성일자"), **btn_copy_s).pack(side="left", padx=(0, 14))

        tk.Label(r0, text="거래처명:", **lbl_s).pack(side="left")
        tk.Entry(r0, textvariable=self.supplier_var, width=18, **ent_s).pack(side="left", padx=(4, 0))

        # Row 1: PR Title + 복사 버튼 2
        r1 = tk.Frame(sec2_card, bg="#FFFFFF")
        r1.pack(fill="x", pady=3)
        tk.Label(r1, text="📌 PR Title:", **lbl_s).pack(side="left")
        tk.Entry(r1, textvariable=self.pr_title_var, **ent_s).pack(side="left", fill="x", expand=True, padx=(4, 4))
        tk.Button(r1, text="📋 복사", command=lambda: self.copy_to_clipboard(self.pr_title_var.get(), "PR Title"), **btn_copy_s).pack(side="right")

        # Row 2: 공급가액 (복사 버튼 3) & 부가세 & 합계금액 + 3종 일괄 복사 버튼
        r2 = tk.Frame(sec2_card, bg="#FFFFFF")
        r2.pack(fill="x", pady=3)

        # 3) 공급가액 + 복사 버튼
        tk.Label(r2, text="💰 공급가액:", **lbl_s).pack(side="left")
        amt_e = tk.Entry(r2, textvariable=self.amount_var, width=15, **ent_s)
        amt_e.pack(side="left", padx=(4, 2))
        amt_e.bind("<KeyRelease>", self._recalc_amounts)
        tk.Button(r2, text="📋 복사", command=lambda: self.copy_to_clipboard(self.amount_var.get(), "공급가액"), **btn_copy_s).pack(side="left", padx=(0, 12))

        tk.Label(r2, text="부가세:", **lbl_s).pack(side="left")
        tk.Entry(r2, textvariable=self.vat_var, width=13, **ent_s).pack(side="left", padx=(4, 12))

        tk.Label(r2, text="합계금액:", **lbl_s).pack(side="left")
        tot_e = tk.Entry(r2, textvariable=self.total_amount_var, width=16, font=("Malgun Gothic", 9, "bold"), bg="#EFF6FF", fg="#1D4ED8", relief="solid", bd=1)
        tot_e.pack(side="left", padx=(4, 12))

        # 3종 전체 복사 버튼 (작성일자 / PR Title / 공급가액 탭구분 일괄 복사)
        btn_copy_all = tk.Button(r2, text="✨ 3종 항목 한꺼번에 복사 (Tab 구분)", font=("Malgun Gothic", 8, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", relief="flat", padx=8, pady=2, cursor="hand2", command=self.copy_all_3items)
        btn_copy_all.pack(side="right")

        # --- Section 3 & Bottom Action Bar ---
        bottom_bar = tk.Frame(main_panel, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=12, pady=8)
        bottom_bar.pack(fill="x")

        sett_b = tk.Frame(bottom_bar, bg="#FFFFFF")
        sett_b.pack(side="left", fill="x", expand=True, padx=(0, 10))

        s1 = tk.Frame(sett_b, bg="#FFFFFF")
        s1.pack(fill="x", pady=1)
        tk.Label(s1, text="템플릿:", font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#475569").pack(side="left")
        tk.Entry(s1, textvariable=self.template_path, font=("Malgun Gothic", 8), bg="#F8FAFC", relief="solid", bd=1).pack(side="left", fill="x", expand=True, padx=4)
        tk.Button(s1, text="📁", font=("Segoe UI Emoji", 8), bg="#F1F5F9", relief="flat", command=self.browse_template).pack(side="left")

        s2 = tk.Frame(sett_b, bg="#FFFFFF")
        s2.pack(fill="x", pady=1)
        tk.Label(s2, text="프린터:", font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#475569").pack(side="left")
        self.printer_combo = ttk.Combobox(s2, textvariable=self.selected_printer, font=("Malgun Gothic", 8), state="readonly")
        self.printer_combo.pack(side="left", fill="x", expand=True, padx=4)

        act_b = tk.Frame(bottom_bar, bg="#FFFFFF")
        act_b.pack(side="right")

        btn_excel = tk.Button(act_b, text="📊 Voucher 엑셀 작성", font=("Malgun Gothic", 9, "bold"), bg="#10B981", fg="white", activebackground="#059669", activeforeground="white", relief="flat", padx=16, pady=8, cursor="hand2", command=self.create_voucher_excel)
        btn_excel.pack(side="left", padx=(0, 6))

        btn_print = tk.Button(act_b, text="🖨️ Voucher & PDF 일괄 인쇄", font=("Malgun Gothic", 9, "bold"), bg="#3B82F6", fg="white", activebackground="#2563EB", activeforeground="white", relief="flat", padx=18, pady=8, cursor="hand2", command=self.print_all_documents)
        btn_print.pack(side="right")

    def copy_to_clipboard(self, text, label_name):
        """
        단일 항목 클립보드 복사
        """
        if not text:
            messagebox.showwarning("복사 실패", f"{label_name} 값이 비어 있습니다.")
            return
        
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        messagebox.showinfo("클립보드 복사 완료", f"[{label_name}] 텍스트가 복사되었습니다!\n\n📋 복사된 내용: {text}")

    def copy_all_3items(self):
        """
        작성일자, PR Title, 공급가액 3종 항목 탭(\t) 구분 일괄 복사 (엑셀 행에 바로 붙여넣기 가능)
        """
        d = self.date_var.get().strip()
        t = self.pr_title_var.get().strip()
        a = self.amount_var.get().strip()

        if not d and not t and not a:
            messagebox.showwarning("복사 실패", "복사할 데이터 항목이 없습니다. PDF 파싱을 먼저 진행하세요.")
            return

        combined_text = f"{d}\t{t}\t{a}"
        self.root.clipboard_clear()
        self.root.clipboard_append(combined_text)
        self.root.update()

        messagebox.showinfo(
            "3종 전체 복사 완료",
            f"다음 3가지 항목이 탭(Tab)으로 구분되어 클립보드에 복사되었습니다.\n엑셀 셀에 바로 Ctrl+V 로 붙여넣으세요!\n\n"
            f"1) 작성일자: {d}\n"
            f"2) PR Title: {t}\n"
            f"3) 공급가액: {a}"
        )

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

            messagebox.showinfo("파싱 완료", f"PR PDF 데이터 분석 성공!\n\n- 작성일자: {data.get('date')}\n- PR Title: {data.get('pr_title')}\n- 공급가액: {amt:,} 원\n\n[📋 복사] 버튼을 눌러 원하시는 곳에 붙여넣으세요.")
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
    app = VoucherPassApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
