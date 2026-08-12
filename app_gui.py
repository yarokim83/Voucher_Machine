import os
import sys
import re
import urllib.parse
import subprocess
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
import outlook_handler
import pdf_watcher

class PastelGlassDropZone(tk.Frame):
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
        self.root.title("VoucherPass v2.0.0 (Full Automation)")
        self.root.geometry("1200x820")
        self.root.minsize(1140, 760)

        self.pr_pdf_path = tk.StringVar()
        self.spec_pdf_path = tk.StringVar()   # 거래명세서
        self.tax_pdf_path = tk.StringVar()    # 세금계산서 (암호 6068625399 자동대입)
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

        self.auto_watch_enabled = tk.BooleanVar(value=True)
        self.folder_watcher = pdf_watcher.DownloadFolderWatcher(self.handle_auto_detected_pdf)

        self._setup_app_icon()
        self._setup_style()
        self._build_layout()
        self._load_printers()
        self._start_folder_watch_timer()

    def _setup_app_icon(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        png_icon = os.path.join(base_dir, "assets", "app_icon.png")
        ico_icon = os.path.join(base_dir, "assets", "app_icon.ico")

        if os.path.exists(png_icon):
            try:
                self.app_icon_img = ImageTk.PhotoImage(Image.open(png_icon).resize((32, 32), Image.Resampling.LANCZOS))
                self.root.iconphoto(True, self.app_icon_img)
            except Exception as e:
                print(f"Icon photo error: {e}")
        elif os.path.exists(ico_icon):
            try:
                self.root.iconbitmap(ico_icon)
            except Exception:
                pass

    def _setup_style(self):
        self.bg_app = "#F8FAFC"
        self.sidebar_bg = "#F1F5F9"
        self.root.configure(bg=self.bg_app)

    def _build_layout(self):
        header = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=16, pady=10)
        header.pack(fill="x")

        traffic_frame = tk.Frame(header, bg="#FFFFFF")
        traffic_frame.pack(side="left", padx=(0, 12))

        for color in ("#FF5F56", "#FFBD2E", "#27C93F"):
            dot = tk.Label(traffic_frame, text="●", font=("Arial", 10), bg="#FFFFFF", fg=color)
            dot.pack(side="left", padx=2)

        base_dir = os.path.dirname(os.path.abspath(__file__))
        png_icon = os.path.join(base_dir, "assets", "app_icon.png")
        if os.path.exists(png_icon):
            try:
                self.header_logo_img = ImageTk.PhotoImage(Image.open(png_icon).resize((28, 28), Image.Resampling.LANCZOS))
                logo_lbl = tk.Label(header, image=self.header_logo_img, bg="#FFFFFF")
                logo_lbl.pack(side="left", padx=(6, 4))
            except Exception:
                tk.Label(header, text="☑", font=("Segoe UI Emoji", 15, "bold"), bg="#FFFFFF", fg="#2563EB").pack(side="left", padx=(6, 4))
        else:
            tk.Label(header, text="☑", font=("Segoe UI Emoji", 15, "bold"), bg="#FFFFFF", fg="#2563EB").pack(side="left", padx=(6, 4))

        logo_txt = tk.Label(header, text="VoucherPass", font=("Malgun Gothic", 15, "bold"), bg="#FFFFFF", fg="#0F172A")
        logo_txt.pack(side="left")

        ver_badge = tk.Label(header, text="v2.0 Full Auto", font=("Malgun Gothic", 8, "bold"), bg="#EFF6FF", fg="#2563EB", padx=6, pady=2)
        ver_badge.pack(side="left", padx=(8, 16))

        # 획기적 1초 아웃룩 추출 버튼
        btn_outlook = tk.Button(header, text="📧 아웃룩 첨부파일 1초 가져오기", font=("Malgun Gothic", 9, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", relief="flat", padx=12, pady=4, cursor="hand2", command=self.fetch_from_outlook)
        btn_outlook.pack(side="left", padx=(0, 10))

        # 다운로드 감시 토글
        chk_watch = tk.Checkbutton(header, text="📥 다운로드 실시간 자동 분류", variable=self.auto_watch_enabled, font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#059669", activebackground="#FFFFFF", cursor="hand2")
        chk_watch.pack(side="left")

        btn_help = tk.Button(header, text="❓ 도움말", font=("Malgun Gothic", 8), bg="#F1F5F9", fg="#475569", relief="flat", padx=10, pady=3, command=lambda: messagebox.showinfo("도움말", "1) 세금계산서 암호(6068625399)는 자동 해제 파싱됩니다.\n2) 아웃룩 메일 선택 후 [📧 가져오기] 버튼을 누르면 서류가 1초만에 들어옵니다.\n3) 다운로드 폴더 자동 감시가 켜져 있어 PDF 내려받기 시 카드에 자동 배치됩니다."))
        btn_help.pack(side="right", padx=2)

        body_container = tk.Frame(self.root, bg=self.bg_app)
        body_container.pack(fill="both", expand=True)

        # Sidebar
        sidebar = tk.Frame(body_container, bg=self.sidebar_bg, width=210, padx=12, pady=16)
        sidebar.pack(side="left", fill="y")

        nav_items = [
            ("☁️ 자동 PDF 수집", True),
            ("📋 데이터 세로 복사", True),
            ("📂 건별 자동 정리", True),
            ("🖨️ PDF 일괄 인쇄", True),
        ]

        for text, active in nav_items:
            btn = tk.Button(sidebar, text=text, font=("Malgun Gothic", 9, "bold"), bg="#EFF6FF", fg="#2563EB", activebackground="#DBEAFE", activeforeground="#1E40AF", relief="flat", anchor="w", padx=14, pady=10, cursor="hand2")
            btn.pack(fill="x", pady=4)

        safe_card = tk.Frame(sidebar, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=12, pady=12)
        safe_card.pack(side="bottom", fill="x", pady=(0, 10))

        tk.Label(safe_card, text="🛡️ 암호 자동 대입 해제", font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#0F172A", anchor="w").pack(fill="x")
        tk.Label(safe_card, text="암호화된 세금계산서(6068625399)\n자동 해제 파싱 완료", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#94A3B8", justify="left", anchor="w").pack(fill="x", pady=(4, 0))

        # Main Panel
        main_panel = tk.Frame(body_container, bg=self.bg_app, padx=18, pady=14)
        main_panel.pack(side="right", fill="both", expand=True)

        sec1_header = tk.Frame(main_panel, bg=self.bg_app)
        sec1_header.pack(fill="x", pady=(0, 8))

        tk.Label(sec1_header, text="1. 스마트 제출 서류 PDF 4종 (아웃룩/다운로드 자동 수집)", font=("Malgun Gothic", 11, "bold"), bg=self.bg_app, fg="#0F172A").pack(side="left")

        btn_select_file = tk.Button(sec1_header, text="📁 수동 파일 선택", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#334155", relief="flat", highlightbackground="#E2E8F0", highlightthickness=1, padx=10, pady=3, command=lambda: self.drop_pr._browse_file())
        btn_select_file.pack(side="right")

        grid_drop = tk.Frame(main_panel, bg=self.bg_app)
        grid_drop.pack(fill="both", expand=True, pady=(0, 10))

        self.drop_pr = PastelGlassDropZone(grid_drop, "① PR Print (구매요청서)", "🛒", "#EFF6FF", "#2563EB", self.pr_pdf_path, on_file_selected=self.parse_uploaded_pdf)
        self.drop_pr.grid(row=0, column=0, padx=6, pady=6, sticky="nsew")

        self.drop_spec = PastelGlassDropZone(grid_drop, "② 거래명세서 PDF", "📄", "#ECFDF5", "#059669", self.spec_pdf_path)
        self.drop_spec.grid(row=0, column=1, padx=6, pady=6, sticky="nsew")

        # ③ 전자 세금계산서 PDF (암호 6068625399 자동 대입 작성일자 우선 연동)
        self.drop_tax = PastelGlassDropZone(grid_drop, "③ 전자 세금계산서 (암호 6068625399)", "🧾", "#F5F3FF", "#7C3AED", self.tax_pdf_path, on_file_selected=self.parse_tax_invoice_uploaded)
        self.drop_tax.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")

        contract_container = tk.Frame(grid_drop, bg=self.bg_app)
        contract_container.grid(row=1, column=1, padx=6, pady=6, sticky="nsew")

        self.drop_contract = PastelGlassDropZone(contract_container, "④ 업체 계약서 PDF", "📝", "#FFFBEB", "#D97706", self.contract_pdf_path)
        self.drop_contract.pack(fill="both", expand=True)

        pg_sub = tk.Frame(contract_container, bg="#FFFFFF", padx=8, pady=3, highlightbackground="#E2E8F0", highlightthickness=1)
        pg_sub.pack(fill="x", pady=(3, 0))
        tk.Label(pg_sub, text="📄 인쇄 대상 페이지:", font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#475569").pack(side="left")
        tk.Entry(pg_sub, textvariable=self.contract_page, font=("Malgun Gothic", 8), width=8, relief="solid", bd=1).pack(side="left", padx=4)
        tk.Label(pg_sub, text="(예: 12-13 또는 1,2)", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#94A3B8").pack(side="left")

        grid_drop.columnconfigure(0, weight=1)
        grid_drop.columnconfigure(1, weight=1)
        grid_drop.rowconfigure(0, weight=1)
        grid_drop.rowconfigure(1, weight=1)

        # -----------------------------------------------------------
        # Section 2: Extracted Data & COPY BUTTONS & ARCHIVER
        # -----------------------------------------------------------
        sec2_card = tk.LabelFrame(main_panel, text=" 📝 2. 추출 데이터 & 원클릭 세로 복사 ", font=("Malgun Gothic", 10, "bold"), bg="#FFFFFF", fg="#1E3A8A", bd=1, relief="solid", padx=12, pady=10)
        sec2_card.pack(fill="x", pady=(0, 10))

        lbl_s = {"font": ("Malgun Gothic", 8, "bold"), "bg": "#FFFFFF", "fg": "#334155", "anchor": "e"}
        ent_s = {"font": ("Malgun Gothic", 9), "bg": "#F8FAFC", "relief": "solid", "bd": 1}
        btn_copy_s = {"font": ("Malgun Gothic", 8, "bold"), "bg": "#EFF6FF", "fg": "#2563EB", "activebackground": "#DBEAFE", "activeforeground": "#1E40AF", "relief": "solid", "bd": 1, "padx": 6, "pady": 1, "cursor": "hand2"}

        r0 = tk.Frame(sec2_card, bg="#FFFFFF")
        r0.pack(fill="x", pady=3)

        tk.Label(r0, text="P/R No:", **lbl_s).pack(side="left")
        tk.Entry(r0, textvariable=self.pr_no_var, width=16, **ent_s).pack(side="left", padx=(4, 12))

        tk.Label(r0, text="📅 작성일자(세금계산서):", **lbl_s).pack(side="left")
        tk.Entry(r0, textvariable=self.date_var, width=13, **ent_s).pack(side="left", padx=(4, 2))
        tk.Button(r0, text="📋 복사", command=lambda: self.copy_to_clipboard(self.date_var.get(), "작성일자"), **btn_copy_s).pack(side="left", padx=(0, 14))

        tk.Label(r0, text="거래처명:", **lbl_s).pack(side="left")
        tk.Entry(r0, textvariable=self.supplier_var, width=18, **ent_s).pack(side="left", padx=(4, 0))

        r1 = tk.Frame(sec2_card, bg="#FFFFFF")
        r1.pack(fill="x", pady=3)
        tk.Label(r1, text="📌 PR Title:", **lbl_s).pack(side="left")
        tk.Entry(r1, textvariable=self.pr_title_var, **ent_s).pack(side="left", fill="x", expand=True, padx=(4, 4))
        tk.Button(r1, text="📋 복사", command=lambda: self.copy_to_clipboard(self.pr_title_var.get(), "PR Title"), **btn_copy_s).pack(side="right")

        r2 = tk.Frame(sec2_card, bg="#FFFFFF")
        r2.pack(fill="x", pady=3)

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

        btn_copy_all = tk.Button(r2, text="✨ 3종 항목 세로 복사 (엑셀 붙여넣기용)", font=("Malgun Gothic", 8, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", relief="flat", padx=8, pady=2, cursor="hand2", command=self.copy_all_3items)
        btn_copy_all.pack(side="right")

        # Bottom Action Bar
        bottom_bar = tk.Frame(main_panel, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=12, pady=8)
        bottom_bar.pack(fill="x")

        sett_b = tk.Frame(bottom_bar, bg="#FFFFFF")
        sett_b.pack(side="left", fill="x", expand=True, padx=(0, 10))

        s2 = tk.Frame(sett_b, bg="#FFFFFF")
        s2.pack(fill="x", pady=1)
        tk.Label(s2, text="출력 프린터 선택:", font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#475569").pack(side="left")
        self.printer_combo = ttk.Combobox(s2, textvariable=self.selected_printer, font=("Malgun Gothic", 9), state="readonly")
        self.printer_combo.pack(side="left", fill="x", expand=True, padx=4)

        act_b = tk.Frame(bottom_bar, bg="#FFFFFF")
        act_b.pack(side="right")

        # 자동 보관 버튼
        btn_archive = tk.Button(act_b, text="📂 건별 자동 폴더 생성 & 보관", font=("Malgun Gothic", 9, "bold"), bg="#059669", fg="white", activebackground="#047857", activeforeground="white", relief="flat", padx=14, pady=9, cursor="hand2", command=self.archive_voucher_files)
        btn_archive.pack(side="left", padx=(0, 6))

        btn_print = tk.Button(act_b, text="🖨️ 업로드 PDF 서류 일괄 인쇄 (4종)", font=("Malgun Gothic", 10, "bold"), bg="#3B82F6", fg="white", activebackground="#2563EB", activeforeground="white", relief="flat", padx=18, pady=9, cursor="hand2", command=self.print_pdf_documents_only)
        btn_print.pack(side="right")

    def _start_folder_watch_timer(self):
        """
        다운로드 폴더 실시간 자동 감시 스케줄러 (2초마다 체크)
        """
        if self.auto_watch_enabled.get():
            try:
                self.folder_watcher.check_new_files()
            except Exception:
                pass
        self.root.after(2000, self._start_folder_watch_timer)

    def handle_auto_detected_pdf(self, pdf_path):
        """
        다운로드 폴더에서 새로 감지된 PDF 파일 스마트 자동 분류 및 드롭 연동
        """
        pdf_type = pdf_parser.classify_pdf_type(pdf_path)
        if pdf_type == 'pr':
            self.drop_pr.set_file(pdf_path)
        elif pdf_type == 'spec':
            self.drop_spec.set_file(pdf_path)
        elif pdf_type == 'tax':
            self.drop_tax.set_file(pdf_path)
        elif pdf_type == 'contract':
            self.drop_contract.set_file(pdf_path)

    def fetch_from_outlook(self):
        """
        아웃룩(Outlook) 선택/최신 이메일에서 PDF 4종 1초 자동 수집
        """
        files = outlook_handler.fetch_outlook_attachments()
        if not files:
            messagebox.showwarning("아웃룩 추출", "아웃룩이 켜져 있지 않거나 선택된 메일에 PDF 첨부파일이 없습니다.")
            return

        count = 0
        for f in files:
            self.handle_auto_detected_pdf(f)
            count += 1

        messagebox.showinfo("아웃룩 추출 성공", f"아웃룩 메일에서 {count}개의 PDF 서류가 추출되어 자동 분류 배치되었습니다!")

    def archive_voucher_files(self):
        """
        건별/업체별 (날짜_업체명_PR번호) 자동 폴더 생성 및 4종 PDF/엑셀 일괄 정리 보관
        """
        data = self._get_form_data()
        pdf_paths = [
            self.pr_pdf_path.get(),
            self.spec_pdf_path.get(),
            self.tax_pdf_path.get(),
            self.contract_pdf_path.get()
        ]

        if not any(pdf_paths):
            messagebox.showwarning("보관 실패", "정리할 PDF 서류가 없습니다.")
            return

        try:
            target_dir, files = excel_handler.archive_voucher_package(data, pdf_paths)
            
            # 폴더 열기 동시 지원
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
            "3종 세로 복사 완료",
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
            
            if not self.date_var.get() and data.get('date'):
                self.date_var.set(data.get('date', ''))

            self.supplier_var.set(data.get('supplier', ''))

            messagebox.showinfo("파싱 완료", f"PR PDF 데이터 분석 성공!\n\n- PR Title: {data.get('pr_title')}\n- 공급가액: {amt:,} 원\n\n전자 세금계산서 PDF를 올리시면 작성일자가 세금계산서 날짜로 자동 업데이트됩니다.")
        except Exception as e:
            messagebox.showerror("파싱 오류", f"PDF 데이터 자동 추출 실패:\n{e}")

    def parse_tax_invoice_uploaded(self, tax_path=None):
        if not tax_path:
            tax_path = self.tax_pdf_path.get()

        if not tax_path or not os.path.exists(tax_path):
            return

        tax_date = pdf_parser.parse_tax_invoice_date(tax_path)
        if tax_date:
            self.date_var.set(tax_date)
            messagebox.showinfo("세금계산서 연동 완료", f"전자 세금계산서 서류에서 [작성일자] 데이터가 성공적으로 연동되었습니다!\n\n- 연동된 작성일자: {tax_date}")
        else:
            fname = os.path.basename(tax_path)
            
            # HTML 파일 드롭 시 자동으로 브라우저 실행 & 비밀번호(6068625399) 클립보드 자동 대입 복사!
            if pdf_parser._is_html_file(tax_path):
                try:
                    # 1. 비밀번호 클립보드 자동 대입
                    self.root.clipboard_clear()
                    self.root.clipboard_append("6068625399")
                    self.root.update()
                    
                    # 2. 브라우저로 HTML 자동 실행
                    os.startfile(tax_path)

                    messagebox.showinfo(
                        "🚀 국세청 세금계산서 HTML 자동 실행",
                        f"국세청 보안 세금계산서 파일({fname})이 웹 브라우저로 자동 실행되었습니다!\n\n"
                        f"🔑 비밀번호 [ 6068625399 ] 가 클립보드에 자동 복사되었습니다.\n\n"
                        f"열린 브라우저 암호 칸에 [ Ctrl + V ] 를 누르시면 1초 만에 세금계산서 화면이 바로 해제됩니다!"
                    )
                except Exception as e:
                    messagebox.showwarning("실행 오류", f"HTML 파일 자동 실행 중 오류 발생: {e}")
            else:
                messagebox.showinfo("세금계산서 연결 완료", f"전자 세금계산서 서류({fname})가 업로드되었습니다!\n\n[작성일자] 필드에 2026-08-11 형태로 직접 입력/수정해 주시면 됩니다.")

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
            if self.pr_pdf_path.get() and os.path.exists(self.pr_pdf_path.get()):
                printer_handler.print_pdf_file(self.pr_pdf_path.get(), printer_name=printer)
                printed_list.append("① PR Print PDF (구매요청서)")

            if self.spec_pdf_path.get() and os.path.exists(self.spec_pdf_path.get()):
                printer_handler.print_pdf_file(self.spec_pdf_path.get(), printer_name=printer)
                printed_list.append("② 거래명세서 PDF")

            # 암호화된 세금계산서 PDF 자동 해제 후 인쇄
            if self.tax_pdf_path.get() and os.path.exists(self.tax_pdf_path.get()):
                dec_pdf = pdf_parser.decrypt_pdf_to_temp(self.tax_pdf_path.get())
                printer_handler.print_pdf_file(dec_pdf, printer_name=printer)
                printed_list.append("③ 전자 세금계산서 PDF (암호해제 인쇄)")

            if self.contract_pdf_path.get() and os.path.exists(self.contract_pdf_path.get()):
                page_str = self.contract_page.get().strip()
                page_indices, label_str = self._parse_contract_pages(page_str)
                
                printer_handler.print_pdf_file(self.contract_pdf_path.get(), printer_name=printer, page_range=page_indices)
                printed_list.append(f"④ 업체 계약서 PDF ({label_str} 페이지)")

            if not printed_list:
                messagebox.showwarning("인쇄할 PDF 없음", "인쇄할 PDF 서류가 하나도 업로드되지 않았습니다.\nPDF 파일을 드래그 앤 드롭 업로드해 주세요.")
                return

            summary = "\n- ".join(printed_list)
            messagebox.showinfo("PDF 서류 일괄 인쇄 완료", f"다음 업로드 PDF 서류들이 프린터 [{printer}] 로 출력 요청되었습니다:\n\n- {summary}")

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
