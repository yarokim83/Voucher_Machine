import os
import sys
import re
import urllib.parse
import subprocess
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
    TkinterDnD = tk

import pdf_parser
import excel_handler
import printer_handler
import pdf_watcher

class CleanMinimalDropZone(tk.Frame):
    """
    VoucherPass v5.2 Large Font & Zero Right-Margin DropZone
    """
    def __init__(self, parent, title, icon, accent_color, file_var, on_file_selected=None, **kwargs):
        super().__init__(parent, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, bd=0, **kwargs)
        self.file_var = file_var
        self.on_file_selected = on_file_selected
        self.accent_color = accent_color

        self.inner = tk.Frame(self, bg="#FFFFFF", padx=5, pady=3)
        self.inner.pack(fill="both", expand=True)

        # 아이콘 타겟 드롭존 박스
        self.icon_bg = tk.Frame(self.inner, bg="#F8FAFC", padx=6, pady=2, highlightbackground=accent_color, highlightthickness=1)
        self.icon_bg.pack(side="left", padx=(0, 6))

        self.lbl_icon = tk.Label(self.icon_bg, text=icon, font=("Segoe UI Emoji", 16), bg="#F8FAFC", fg=accent_color)
        self.lbl_icon.pack()

        txt_box = tk.Frame(self.inner, bg="#FFFFFF")
        txt_box.pack(side="left", fill="both", expand=True)

        self.lbl_title = tk.Label(txt_box, text=title, font=("Malgun Gothic", 10, "bold"), bg="#FFFFFF", fg="#0F172A", anchor="w")
        self.lbl_title.pack(fill="x")

        self.lbl_status = tk.Label(txt_box, text="파일 드래그 앤 드롭 또는 클릭", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#475569", anchor="w", justify="left")
        self.lbl_status.pack(fill="x", pady=(0, 0))

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
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf"), ("HTML Files", "*.html;*.htm"), ("All Files", "*.*")])
        if path:
            self.set_file(path)

    def _handle_drop(self, event):
        self._on_drag_leave()
        raw_data = event.data if event and hasattr(event, 'data') else ''
        cleaned_data = urllib.parse.unquote(str(raw_data)).strip()
        if cleaned_data.startswith('file:///'):
            cleaned_data = cleaned_data[8:]

        matches = re.findall(r'\{([^}]+)\}|(\S+)', cleaned_data)
        valid_file = None

        for m in matches:
            p = m[0] if m[0] else m[1]
            p = p.strip('\"\'{} \t\r\n')
            if p:
                norm_p = os.path.abspath(os.path.normpath(p))
                ext = norm_p.lower()
                if (ext.endswith('.pdf') or ext.endswith('.html') or ext.endswith('.htm')) and os.path.exists(norm_p):
                    valid_file = norm_p
                    break

        if valid_file:
            self.set_file(valid_file)
        else:
            messagebox.showwarning("파일 형식 오류", "PDF 또는 HTML 파일을 업로드해 주세요.")

    def _on_drag_enter(self, event=None):
        self.config(bg="#EFF6FF", highlightbackground="#2563EB", highlightthickness=2)
        self.inner.config(bg="#EFF6FF")

    def _on_drag_leave(self, event=None):
        is_set = bool(self.file_var.get())
        bg_col = "#F0FDF4" if is_set else "#FFFFFF"
        border_col = "#16A34A" if is_set else "#CBD5E1"
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
            self.lbl_status.config(text="파일 드래그 앤 드롭 또는 클릭", fg="#475569", font=("Malgun Gothic", 8))
            self.config(bg="#FFFFFF", highlightbackground="#CBD5E1")
            self.inner.config(bg="#FFFFFF")


class VoucherPassApp:
    """
    VoucherPass v5.2 Large Font & Zero Margin Engine
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Voucher Pass")
        self.root.geometry("390x560+1100+100")
        self.root.minsize(390, 560)
        self.root.maxsize(390, 560)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#F8FAFC", highlightbackground="#2563EB", highlightthickness=2)

        self._offsetx = 0
        self._offsety = 0

        self.pr_pdf_path = tk.StringVar()
        self.po_pdf_path = tk.StringVar()
        self.spec_pdf_path = tk.StringVar()
        self.tax_pdf_path = tk.StringVar()
        self.contract_pdf_path = tk.StringVar()
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

        self.auto_watch_enabled = tk.BooleanVar(value=True)
        self.folder_watcher = pdf_watcher.DownloadFolderWatcher(self.handle_auto_detected_pdf)

        self._build_widget_layout()
        self._load_printers()
        self._start_folder_watch_timer()

        # PR Title Text - StringVar 양방향 동기화
        self.pr_title_var.trace_add("write", self._on_pr_title_var_changed)

    def _on_pr_title_var_changed(self, *args):
        val = self.pr_title_var.get()
        if hasattr(self, 'txt_pr_title') and self.txt_pr_title.get("1.0", "end-1c") != val:
            self.txt_pr_title.delete("1.0", tk.END)
            self.txt_pr_title.insert("1.0", val)

    def _on_pr_title_txt_changed(self, event=None):
        if hasattr(self, 'txt_pr_title'):
            val = self.txt_pr_title.get("1.0", "end-1c").strip()
            if self.pr_title_var.get() != val:
                self.pr_title_var.set(val)

    def toggle_visibility(self, event=None):
        if self.root.state() == "normal":
            self.root.withdraw()
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

    def _click_title(self, event):
        self._offsetx = event.x
        self._offsety = event.y

    def _drag_title(self, event):
        x = self.root.winfo_pointerx() - self._offsetx
        y = self.root.winfo_pointery() - self._offsety
        self.root.geometry(f"+{x}+{y}")

    def _build_widget_layout(self):
        # 1. Header Bar (HPNT Style Vibrant Blue Header)
        hdr = tk.Frame(self.root, bg="#2563EB", padx=8, pady=5)
        hdr.pack(fill="x")
        hdr.bind("<Button-1>", self._click_title)
        hdr.bind("<B1-Motion>", self._drag_title)

        lbl_logo = tk.Label(hdr, text="⚡ Voucher Pass", font=("Malgun Gothic", 10, "bold"), bg="#2563EB", fg="#FFFFFF")
        lbl_logo.pack(side="left")
        lbl_logo.bind("<Button-1>", self._click_title)
        lbl_logo.bind("<B1-Motion>", self._drag_title)

        ver_b = tk.Label(hdr, text="v5.2", font=("Malgun Gothic", 8, "bold"), bg="#1D4ED8", fg="white", padx=5, pady=1)
        ver_b.pack(side="left", padx=(6, 0))

        btn_min = tk.Label(hdr, text=" ─ ", font=("Arial", 10, "bold"), bg="#2563EB", fg="#DBEAFE", cursor="hand2")
        btn_min.pack(side="right", padx=(3, 0))
        btn_min.bind("<Button-1>", lambda e: self.root.withdraw())

        btn_close = tk.Label(hdr, text=" ✕ ", font=("Arial", 10, "bold"), bg="#2563EB", fg="#FEE2E2", cursor="hand2")
        btn_close.pack(side="right")
        btn_close.bind("<Button-1>", lambda e: self.root.destroy())

        # 2. Main Content Canvas (Soft Ice White Canvas - 좌우 여백 타이트 3px)
        main_box = tk.Frame(self.root, bg="#F8FAFC", padx=3, pady=2)
        main_box.pack(fill="both", expand=True)

        # 3. 5가지 서류 Clean DropZone
        self.drop_tax = CleanMinimalDropZone(main_box, "① 전자 세금계산서 (6068625399)", "🧾", "#7C3AED", self.tax_pdf_path, on_file_selected=self.parse_tax_invoice_uploaded)
        self.drop_tax.pack(fill="x", pady=1)

        self.drop_spec = CleanMinimalDropZone(main_box, "② 거래명세서 PDF", "📄", "#059669", self.spec_pdf_path)
        self.drop_spec.pack(fill="x", pady=1)

        self.drop_pr = CleanMinimalDropZone(main_box, "③ PR Print (구매요청서) PDF", "🛒", "#2563EB", self.pr_pdf_path, on_file_selected=self.parse_uploaded_pdf)
        self.drop_pr.pack(fill="x", pady=1)

        self.drop_po = CleanMinimalDropZone(main_box, "④ 발주서 (PO) PDF", "📦", "#0284C7", self.po_pdf_path, on_file_selected=self.parse_uploaded_pdf)
        self.drop_po.pack(fill="x", pady=1)

        self.drop_contract = CleanMinimalDropZone(main_box, "⑤ 업체 계약서 PDF", "📝", "#D97706", self.contract_pdf_path)
        self.drop_contract.pack(fill="x", pady=1)

        # 4. Clean Minimal Spotlight HUD Data Board (대형 시원한 폰트 보드)
        hud = tk.LabelFrame(main_box, text=" 📝 추출 데이터 7종 ", font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=4, pady=2)
        hud.pack(fill="x", pady=(2, 1))

        lbl_s = {"font": ("Malgun Gothic", 9, "bold"), "bg": "#FFFFFF", "fg": "#0F172A"}
        ent_s = {"font": ("Malgun Gothic", 9, "bold"), "bg": "#F8FAFC", "fg": "#0F172A", "insertbackground": "#0F172A", "relief": "solid", "bd": 1}
        btn_cp = {"font": ("Malgun Gothic", 8, "bold"), "bg": "#EFF6FF", "fg": "#2563EB", "relief": "flat", "padx": 4, "cursor": "hand2"}

        # Row 1: P/R No & Date
        r1 = tk.Frame(hud, bg="#FFFFFF")
        r1.pack(fill="x", pady=1)
        tk.Label(r1, text="P/R No:", **lbl_s).pack(side="left")
        e_prno = tk.Entry(r1, textvariable=self.pr_no_var, width=11, **ent_s)
        e_prno.pack(side="left", padx=(2, 4))

        tk.Label(r1, text="📅 작성일자:", **lbl_s).pack(side="left")
        e_date = tk.Entry(r1, textvariable=self.date_var, font=("Malgun Gothic", 9, "bold"), bg="#F0FDF4", fg="#15803D", width=10, relief="solid", bd=1)
        e_date.pack(side="left", padx=(2, 2))
        tk.Button(r1, text="📋", command=lambda: self.copy_to_clipboard(self.date_var.get(), "작성일자"), **btn_cp).pack(side="left")

        # Row 2: PR Title (가로폭 꽉 차게 width=22 및 우측 복사 버튼 고정 밀착)
        r2 = tk.Frame(hud, bg="#FFFFFF")
        r2.pack(fill="x", pady=1)
        tk.Label(r2, text="📌 PR Title:", **lbl_s).pack(side="left", anchor="n", pady=2)
        
        self.txt_pr_title = tk.Text(r2, width=22, height=2, wrap="word", font=("Malgun Gothic", 9, "bold"), bg="#F8FAFC", fg="#0F172A", insertbackground="#0F172A", relief="solid", bd=1)
        self.txt_pr_title.pack(side="left", padx=(2, 4))
        self.txt_pr_title.bind("<KeyRelease>", self._on_pr_title_txt_changed)

        btn_copy_title = tk.Button(r2, text="📋 복사", command=lambda: self.copy_to_clipboard(self.pr_title_var.get(), "PR Title"), **btn_cp)
        btn_copy_title.pack(side="left", anchor="n", pady=2)

        # Row 3: Supplier & Amount & Copy
        r3 = tk.Frame(hud, bg="#FFFFFF")
        r3.pack(fill="x", pady=1)
        tk.Label(r3, text="🏢 거래처명:", **lbl_s).pack(side="left")
        e_sup = tk.Entry(r3, textvariable=self.supplier_var, width=11, **ent_s)
        e_sup.pack(side="left", padx=(2, 4))

        tk.Label(r3, text="💰 공급가액:", **lbl_s).pack(side="left")
        e_amt = tk.Entry(r3, textvariable=self.amount_var, font=("Malgun Gothic", 9, "bold"), bg="#EFF6FF", fg="#1D4ED8", width=10, relief="solid", bd=1)
        e_amt.pack(side="left", padx=(2, 2))
        e_amt.bind("<KeyRelease>", self._recalc_amounts)
        tk.Button(r3, text="📋", command=lambda: self.copy_to_clipboard(self.amount_var.get(), "공급가액"), **btn_cp).pack(side="left")

        # Row 4: VAT & Total Amount
        r4 = tk.Frame(hud, bg="#FFFFFF")
        r4.pack(fill="x", pady=1)
        tk.Label(r4, text="💵 부가세:", **lbl_s).pack(side="left")
        e_vat = tk.Entry(r4, textvariable=self.vat_var, width=10, **ent_s)
        e_vat.pack(side="left", padx=(2, 4))

        tk.Label(r4, text="💳 합계금액:", **lbl_s).pack(side="left")
        e_tot = tk.Entry(r4, textvariable=self.total_amount_var, font=("Malgun Gothic", 9, "bold"), bg="#EFF6FF", fg="#1D4ED8", width=11, relief="solid", bd=1)
        e_tot.pack(side="left", padx=(2, 0))

        # 5. Clean Minimal Action Buttons (대형 10pt 폰트)
        act_panel = tk.Frame(main_box, bg="#F8FAFC")
        act_panel.pack(fill="x", pady=(1, 0))

        btn_copy_all = tk.Button(act_panel, text="📋 Voucher 엑셀 양식 붙여넣기 (클립보드 복사)", font=("Malgun Gothic", 9, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", relief="flat", padx=6, pady=5, cursor="hand2", command=self.copy_all_3items)
        btn_copy_all.pack(side="top", fill="x", pady=1)

        bot_btn_f = tk.Frame(act_panel, bg="#F8FAFC")
        bot_btn_f.pack(fill="x", pady=0)

        btn_arch = tk.Button(bot_btn_f, text="📂 건별 자동 보관", font=("Malgun Gothic", 9, "bold"), bg="#059669", fg="white", activebackground="#047857", activeforeground="white", relief="flat", padx=4, pady=4, cursor="hand2", command=self.archive_voucher_files)
        btn_arch.pack(side="left", fill="x", expand=True, padx=(0, 1))

        btn_print = tk.Button(bot_btn_f, text="🖨️ 서류 5종 일괄 인쇄 🖨️", font=("Malgun Gothic", 9, "bold"), bg="#7C3AED", fg="white", activebackground="#6D28D9", activeforeground="white", relief="flat", padx=4, pady=4, cursor="hand2", command=self.print_pdf_documents_only)
        btn_print.pack(side="right", fill="x", expand=True, padx=(1, 0))

        # 6. Printer Selector Bar
        prt_bar = tk.Frame(main_box, bg="#F8FAFC")
        prt_bar.pack(fill="x", pady=(1, 0))
        tk.Label(prt_bar, text="🖨️ 프린터:", font=("Malgun Gothic", 9, "bold"), bg="#F8FAFC", fg="#475569").pack(side="left")
        self.printer_combo = ttk.Combobox(prt_bar, textvariable=self.selected_printer, font=("Malgun Gothic", 9), state="readonly")
        self.printer_combo.pack(side="left", fill="x", expand=True, padx=2)

    def _start_folder_watch_timer(self):
        if self.auto_watch_enabled.get():
            try:
                self.folder_watcher.check_new_files()
            except Exception:
                pass
        self.root.after(2000, self._start_folder_watch_timer)

    def handle_auto_detected_pdf(self, pdf_path):
        pdf_type = pdf_parser.classify_pdf_type(pdf_path)
        if pdf_type == 'pr':
            self.drop_pr.set_file(pdf_path)
        elif pdf_type == 'po':
            self.drop_po.set_file(pdf_path)
        elif pdf_type == 'spec':
            self.drop_spec.set_file(pdf_path)
        elif pdf_type == 'tax':
            self.drop_tax.set_file(pdf_path)
        elif pdf_type == 'contract':
            self.drop_contract.set_file(pdf_path)

    def generate_excel_action(self):
        data = self._get_form_data()
        if not data.get('pr_no') and not data.get('pr_title') and not data.get('amount'):
            messagebox.showwarning("엑셀 생성 실패", "기록할 데이터가 없습니다. PDF 서류를 먼저 드롭해 주세요.")
            return

        try:
            tmpl_path = self.template_path.get()
            out_excel = excel_handler.generate_voucher_excel(data, template_path=tmpl_path)
            res = messagebox.askyesno("📊 엑셀 기록 저장 완료", f"고려제강 엑셀 템플릿에 데이터가 정상 기록되었습니다!\n\n📄 엑셀 파일: {out_excel}\n\n지금 해당 엑셀 파일을 열어보시겠습니까?")
            if res:
                os.startfile(out_excel)
        except Exception as e:
            messagebox.showerror("엑셀 저장 오류", f"엑셀 생성 중 오류가 발생했습니다:\n{e}")

    def archive_voucher_files(self):
        data = self._get_form_data()
        pdf_paths = [
            self.pr_pdf_path.get(),
            self.po_pdf_path.get(),
            self.spec_pdf_path.get(),
            self.tax_pdf_path.get(),
            self.contract_pdf_path.get()
        ]

        if not any(pdf_paths):
            messagebox.showwarning("보관 실패", "정리할 PDF 서류가 없습니다.")
            return

        try:
            target_dir, files = excel_handler.archive_voucher_package(data, pdf_paths)
            res = messagebox.askyesno("자동 정리 보관 완료", f"다음 보관소 폴더가 자동 생성되고 서류가 정돈되었습니다!\n\n📂 위치: {target_dir}\n\n지금 해당 폴더를 탐색기로 열어보시겠습니까?")
            if res:
                os.startfile(target_dir)

        except Exception as e:
            messagebox.showerror("보관 오류", f"건별 폴더 정리 중 오류 발생:\n{e}")

    def copy_to_clipboard(self, text, label_name):
        if not text:
            messagebox.showwarning("복사 실패", f"{label_name} 값이 비어 있습니다.")
            return
        
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        messagebox.showinfo("클립보드 복사 완료", f"[{label_name}] 텍스트가 복사되었습니다!\n\n📋 복사된 내용: {text}")

    def copy_all_3items(self):
        d = self.date_var.get().strip()
        t = self.pr_title_var.get().strip()
        a = self.amount_var.get().strip()

        if not d and not t and not a:
            messagebox.showwarning("복사 실패", "복사할 데이터 항목이 없습니다. PDF 파싱을 먼저 진행하세요.")
            return

        combined_text = f"{d}\n{t}\n{a}"
        self.root.clipboard_clear()
        self.root.clipboard_append(combined_text)
        self.root.update()

        messagebox.showinfo(
            "Voucher 엑셀 양식 복사 완료",
            f"다음 3가지 항목이 세로 줄바꿈(\\n)으로 복사되었습니다.\n엑셀 셀 선택 후 Ctrl+V 하시면 세로 3개 셀에 순서대로 들어갑니다!\n\n"
            f"1) 작성일자: {d}\n"
            f"2) PR Title: {t}\n"
            f"3) 공급가액: {a}"
        )

    def _load_printers(self):
        printers = printer_handler.get_installed_printers()
        self.printer_combo['values'] = printers
        if printers:
            self.selected_printer.set(printers[0])

    def parse_uploaded_pdf(self, pr_path=None):
        if not pr_path:
            pr_path = self.pr_pdf_path.get() or self.po_pdf_path.get()

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
            
            if not self.date_var.get() and data.get('date'):
                self.date_var.set(data.get('date', ''))

            self.supplier_var.set(data.get('supplier', ''))

            messagebox.showinfo("파싱 완료", f"PR/발주서 PDF 데이터 분석 성공!\n\n- Title: {data.get('pr_title')}\n- 공급가액: {amt:,} 원")
        except Exception as e:
            messagebox.showerror("파싱 오류", f"PDF 데이터 자동 추출 실패:\n{e}")

    def parse_tax_invoice_uploaded(self, tax_path=None):
        if not tax_path:
            tax_path = self.tax_pdf_path.get()

        if not tax_path or not os.path.exists(tax_path):
            return

        try:
            # HTML 파일일 경우 즉시 무음 자동 해제 및 PDF 생성 파이프라인 작동
            if pdf_parser._is_html_file(tax_path):
                self.auto_unlock_and_print_html_tax_invoice(tax_path)
                return

            # PDF 파일 파싱
            parse_target = tax_path
            try:
                dec_path = pdf_parser.decrypt_pdf_to_temp(tax_path)
                if dec_path and os.path.exists(dec_path):
                    parse_target = dec_path
            except Exception:
                pass

            tax_date = pdf_parser.parse_tax_invoice_date(parse_target)
            if tax_date:
                self.date_var.set(tax_date)
                messagebox.showinfo("세금계산서 파싱 완료", f"전자 세금계산서 작성일자가 성공적으로 추출되었습니다!\n\n📅 작성일자: {tax_date}")
            else:
                messagebox.showwarning("작성일자 미추출", "세금계산서에서 작성일자를 자동으로 찾지 못했습니다. 수동으로 입력해 주세요.")
        except Exception as e:
            messagebox.showerror("세금계산서 파싱 오류", f"세금계산서 작성일자 파싱 중 오류 발생:\n{e}")

    def auto_unlock_and_print_html_tax_invoice(self, html_path):
        import threading, time
        def _worker():
            try:
                import pyautogui, pyperclip
                pyautogui.FAILSAFE = False
                pyautogui.PAUSE = 0.1

                desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
                downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
                
                def _get_pdfs():
                    files = set()
                    for folder in [desktop, downloads]:
                        if os.path.exists(folder):
                            for f in os.listdir(folder):
                                if f.lower().endswith('.pdf'):
                                    files.add(os.path.join(folder, f))
                    return files

                before_pdfs = _get_pdfs()
                pyperclip.copy("6068625399")

                os.startfile(html_path)
                time.sleep(1.2)

                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.3)
                pyautogui.press('enter')

                time.sleep(1.8)
                pyautogui.hotkey('ctrl', 'p')
                time.sleep(1.0)
                pyautogui.press('enter')

                for _ in range(30):
                    time.sleep(1.0)
                    after_pdfs = _get_pdfs()
                    diff = list(after_pdfs - before_pdfs)
                    if diff:
                        newest_pdf = max(diff, key=lambda x: os.path.getmtime(x))
                        self.root.after(0, lambda p=newest_pdf: self._on_new_tax_pdf_saved(p))
                        break

            except Exception as e:
                print(f"Auto unlock & print error: {e}")

        threading.Thread(target=_worker, daemon=True).start()

    def _on_new_tax_pdf_saved(self, pdf_path):
        if not pdf_path or not os.path.exists(pdf_path):
            return

        def _parse_job():
            import time
            # 1. 새로 생성된 PDF 파일 쓰기가 완료될 때까지 최대 3초 대기
            for _ in range(10):
                try:
                    if os.path.getsize(pdf_path) > 500:
                        break
                except Exception:
                    pass
                time.sleep(0.3)

            time.sleep(0.5)
            self.drop_tax.set_file(pdf_path)

            # 2. 복호화 및 날짜 파싱 재시도 3회
            parse_target = pdf_path
            try:
                dec_p = pdf_parser.decrypt_pdf_to_temp(pdf_path)
                if dec_p and os.path.exists(dec_p):
                    parse_target = dec_p
            except Exception:
                pass

            parsed_date = ''
            for _ in range(3):
                parsed_date = pdf_parser.parse_tax_invoice_date(parse_target)
                if parsed_date:
                    break
                time.sleep(0.3)

            if parsed_date:
                self.root.after(0, lambda d=parsed_date: self._apply_tax_date(d))
            else:
                self.root.after(0, lambda: messagebox.showwarning("작성일자 미추출", f"PDF 세금계산서 [{os.path.basename(pdf_path)}] 에서 작성일자를 자동으로 찾지 못했습니다.\n수동으로 작성일자를 입력해 주세요."))

        threading.Thread(target=_parse_job, daemon=True).start()

    def _apply_tax_date(self, tax_date):
        self.date_var.set(tax_date)
        messagebox.showinfo("세금계산서 날짜 추출 성공", f"🎉 새로 저장된 세금계산서 PDF에서 작성일자가 자동으로 추출되었습니다!\n\n📅 작성일자: {tax_date}")

    def _recalc_amounts(self, event=None):
        raw = self.amount_var.get().replace(',', '').strip()
        if raw.isdigit():
            amt = int(raw)
            vat = int(amt * 0.1)
            tot = amt + vat
            self.vat_var.set(f"{vat:,}")
            self.total_amount_var.set(f"{tot:,}")

    def _parse_contract_pages(self, page_str):
        pages = set()
        cleaned = page_str.strip()
        if not cleaned:
            return [0], "1"
        
        parts = cleaned.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                sub = part.split('-')
                if len(sub) == 2 and sub[0].strip().isdigit() and sub[1].strip().isdigit():
                    start = int(sub[0].strip())
                    end = int(sub[1].strip())
                    for p in range(min(start, end), max(start, end) + 1):
                        if p >= 1:
                            pages.add(p - 1)
            elif part.isdigit():
                p = int(part)
                if p >= 1:
                    pages.add(p - 1)

        sorted_pages = sorted(list(pages))
        if not sorted_pages:
            sorted_pages = [0]

        return sorted_pages, cleaned

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

    def print_pdf_documents_only(self):
        printer = self.selected_printer.get()
        printed_list = []

        try:
            if self.tax_pdf_path.get() and os.path.exists(self.tax_pdf_path.get()):
                dec_pdf = pdf_parser.decrypt_pdf_to_temp(self.tax_pdf_path.get())
                printer_handler.print_pdf_file(dec_pdf, printer_name=printer)
                printed_list.append("① 전자 세금계산서 PDF (암호해제 인쇄)")

            if self.spec_pdf_path.get() and os.path.exists(self.spec_pdf_path.get()):
                printer_handler.print_pdf_file(self.spec_pdf_path.get(), printer_name=printer)
                printed_list.append("② 거래명세서 PDF")

            if self.pr_pdf_path.get() and os.path.exists(self.pr_pdf_path.get()):
                printer_handler.print_pdf_file(self.pr_pdf_path.get(), printer_name=printer)
                printed_list.append("③ PR Print PDF (구매요청서)")

            if self.po_pdf_path.get() and os.path.exists(self.po_pdf_path.get()):
                printer_handler.print_pdf_file(self.po_pdf_path.get(), printer_name=printer)
                printed_list.append("④ 발주서 (PO) PDF")

            if self.contract_pdf_path.get() and os.path.exists(self.contract_pdf_path.get()):
                page_str = self.contract_page.get().strip()
                page_indices, label_str = self._parse_contract_pages(page_str)
                printer_handler.print_pdf_file(self.contract_pdf_path.get(), printer_name=printer, page_range=page_indices)
                printed_list.append(f"⑤ 업체 계약서 PDF ({label_str} 페이지)")

            if not printed_list:
                messagebox.showwarning("인쇄할 PDF 없음", "인쇄할 PDF 서류가 하나도 업로드되지 않았습니다.\nPDF 파일을 드래그 앤 드롭 업로드해 주세요.")
                return

            summary = "\n- ".join(printed_list)
            messagebox.showinfo("제출 서류 5종 일괄 인쇄 완료", f"다음 업로드 PDF 서류들이 프린터 [{printer}] 로 출력 요청되었습니다:\n\n- {summary}")

        except Exception as e:
            messagebox.showerror("인쇄 오류", f"PDF 서류 일괄 인쇄 중 오류 발생:\n{e}")

def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = VoucherPassApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
