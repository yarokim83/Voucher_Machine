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
    VoucherPass v7.0 Clean UX DropZone (완료: 초록 실선 / 미완료: 점선-연한 테두리)
    """
    def __init__(self, parent, title, icon, accent_color, file_var, on_file_selected=None, on_state_changed=None, **kwargs):
        super().__init__(parent, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, bd=0, **kwargs)
        self.file_var = file_var
        self.on_file_selected = on_file_selected
        self.on_state_changed = on_state_changed
        self.accent_color = accent_color

        self.inner = tk.Frame(self, bg="#FFFFFF", padx=6, pady=3)
        self.inner.pack(fill="both", expand=True)

        self.lbl_icon = tk.Label(self.inner, text=icon, font=("Segoe UI Emoji", 11), bg="#FFFFFF", fg="#475569")
        self.lbl_icon.pack(side="left", padx=(0, 6))

        txt_box = tk.Frame(self.inner, bg="#FFFFFF")
        txt_box.pack(side="left", fill="both", expand=True)

        self.lbl_title = tk.Label(txt_box, text=title, font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#1E293B", anchor="w")
        self.lbl_title.pack(fill="x")

        self.lbl_status = tk.Label(txt_box, text="미완료 (드래그 앤 드롭)", font=("Malgun Gothic", 8), bg="#FFFFFF", fg="#94A3B8", anchor="w")
        self.lbl_status.pack(fill="x")

        for w in (self, self.inner, self.lbl_icon, txt_box, self.lbl_title, self.lbl_status):
            w.config(cursor="hand2")
            w.bind("<Button-1>", self._browse_file)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._handle_drop)
            self.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            self.dnd_bind('<<DragLeave>>', self._on_drag_leave)

        self.file_var.trace_add("write", self._update_ui_state)

    def _browse_file(self, event=None):
        path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
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
                if os.path.exists(norm_p):
                    valid_file = norm_p
                    break

        if not valid_file and cleaned_data:
            possible_p = cleaned_data.strip('\"\'{} \t\r\n')
            if os.path.exists(possible_p):
                valid_file = os.path.abspath(os.path.normpath(possible_p))

        if valid_file:
            self.set_file(valid_file)

    def _on_drag_enter(self, event=None):
        self.config(bg="#EFF6FF", highlightbackground="#2563EB", highlightthickness=2)
        self.inner.config(bg="#EFF6FF")

    def _on_drag_leave(self, event=None):
        self._update_ui_state()

    def set_file(self, path):
        self.file_var.set(path)
        if self.on_file_selected:
            self.on_file_selected(path)

    def _update_ui_state(self, *args):
        path = self.file_var.get()
        if path and os.path.exists(path):
            fname = os.path.basename(path)
            # 완료 상태: 초록 실선 테두리 + 초록 배경 + 초록 체크 표시
            self.config(bg="#F0FDF4", highlightbackground="#16A34A", highlightthickness=2)
            self.inner.config(bg="#F0FDF4")
            self.lbl_icon.config(bg="#F0FDF4", fg="#15803D")
            self.lbl_title.master.config(bg="#F0FDF4")
            self.lbl_title.config(bg="#F0FDF4", fg="#14532D")
            self.lbl_status.config(text=f"✓ 완료: {fname}", fg="#15803D", font=("Malgun Gothic", 8, "bold"), bg="#F0FDF4")
        else:
            # 미완료 상태: 연한 회색 배경 + 연한 테두리
            self.config(bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1)
            self.inner.config(bg="#FFFFFF")
            self.lbl_icon.config(bg="#FFFFFF", fg="#475569")
            self.lbl_title.master.config(bg="#FFFFFF")
            self.lbl_title.config(bg="#FFFFFF", fg="#1E293B")
            self.lbl_status.config(text="미완료 (드래그 앤 드롭)", fg="#94A3B8", font=("Malgun Gothic", 8), bg="#FFFFFF")

        if self.on_state_changed:
            self.on_state_changed()


class VoucherPassApp:
    """
    VoucherPass v6.0 Folder Watcher & Auto PDF Date Extractor Engine
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Voucher Pass")
        self.root.geometry("390x645+1100+50")
        self.root.minsize(390, 645)
        self.root.maxsize(390, 645)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#F8FAFC", highlightbackground="#2563EB", highlightthickness=2)

        icon_path = self._get_icon_file('VoucherPass.ico')
        if icon_path and os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

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

        # 시작프로그램 자동 등록, 바로가기 생성 및 시스템 트레이/핫키 초기화
        self.register_startup_registry()
        self.create_desktop_shortcut()
        self._setup_system_tray()
        self._setup_global_hotkey()

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
            self.popup_at_cursor()

    def popup_at_cursor(self):
        """
        현재 마우스 커서 위치로 위젯을 뿅! 팝업
        """
        try:
            mx = self.root.winfo_pointerx()
            my = self.root.winfo_pointery()

            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()

            w, h = 390, 645
            x = max(10, min(mx - 20, sw - w - 10))
            y = max(10, min(my - 20, sh - h - 40))

            self.root.geometry(f"{w}x{h}+{x}+{y}")
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            self.root.deiconify()

    def is_startup_registered(self):
        """
        시작프로그램 등록 여부 확인
        """
        try:
            import winreg
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
                winreg.QueryValueEx(key, "VoucherPass")
                return True
        except Exception:
            return False

    def register_startup_registry(self):
        """
        윈도우 시작 프로그램 레지스트리 자동 등록
        """
        try:
            import winreg
            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dist', 'VoucherPass.exe')
            if not os.path.exists(exe_path):
                exe_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "VoucherPass", 0, winreg.REG_SZ, f'"{exe_path}"')
            self.set_live_status("🚀 윈도우 시작프로그램에 VoucherPass가 등록되었습니다!", type="success")
        except Exception as e:
            print(f"Startup registry error: {e}")

    def unregister_startup_registry(self):
        """
        윈도우 시작 프로그램 레지스트리 해제
        """
        try:
            import winreg
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, "VoucherPass")
            self.set_live_status("🛑 윈도우 시작프로그램 등록이 해제되었습니다.", type="info")
        except Exception as e:
            print(f"Unregister startup error: {e}")

    def _get_icon_file(self, filename):
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        candidates = [
            os.path.join(base_dir, 'dist', filename),
            os.path.join(base_dir, filename),
            os.path.join(os.path.dirname(sys.executable), 'dist', filename),
            os.path.join(os.path.dirname(sys.executable), filename),
            os.path.join(os.getcwd(), 'dist', filename),
            os.path.join(os.getcwd(), filename),
        ]
        for c in candidates:
            if c and os.path.exists(c):
                return c
        return None

    def _get_config_path(self):
        appdata = os.getenv('APPDATA')
        if appdata:
            config_dir = os.path.join(appdata, 'VoucherPass')
            os.makedirs(config_dir, exist_ok=True)
            return os.path.join(config_dir, 'printer_config.json')
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_dir, 'printer_config.json')

    def _setup_system_tray(self):
        """
        윈도우 우측 하단 시스템 트레이 아이콘 등록 (동적 토글 지원)
        """
        def _tray_worker():
            try:
                import pystray
                from PIL import Image

                icon_path = self._get_icon_file('VoucherPass.ico') or self._get_icon_file('app_icon.png')
                if icon_path and os.path.exists(icon_path):
                    image = Image.open(icon_path)
                else:
                    image = Image.new('RGB', (64, 64), color=(37, 99, 235))

                def _on_toggle(icon, item):
                    self.root.after(0, self.popup_at_cursor)

                def _toggle_startup(icon, item):
                    if self.is_startup_registered():
                        self.unregister_startup_registry()
                    else:
                        self.register_startup_registry()

                def _on_exit(icon, item):
                    icon.stop()
                    self.root.after(0, self.root.destroy)

                menu = pystray.Menu(
                    pystray.MenuItem("⚡ 위젯 열기/숨기기 (Ctrl+Shift+V / 마우스 위치)", _on_toggle, default=True),
                    pystray.MenuItem("🚀 시작프로그램 자동등록", _toggle_startup, checked=lambda item: self.is_startup_registered()),
                    pystray.MenuItem("🖥️ 바탕화면 바로가기 아이콘 생성", lambda icon, item: self.create_desktop_shortcut()),
                    pystray.MenuItem("✕ 종료", _on_exit)
                )

                self.tray_icon = pystray.Icon("VoucherPass", image, "Voucher Pass", menu)
                self.tray_icon.run()
            except Exception as e:
                print(f"System tray error: {e}")

        threading.Thread(target=_tray_worker, daemon=True).start()

    def _setup_global_hotkey(self):
        """
        전역 단축키 (Ctrl+Shift+V) 감지 훅
        """
        def _hotkey_worker():
            try:
                from pynput import keyboard

                def _on_trigger():
                    self.root.after(0, self.popup_at_cursor)

                with keyboard.GlobalHotKeys({
                    '<ctrl>+<shift>+v': _on_trigger,
                    '<ctrl>+<alt>+v': _on_trigger
                }) as h:
                    h.join()
            except Exception as e:
                print(f"Global hotkey error: {e}")

        threading.Thread(target=_hotkey_worker, daemon=True).start()

    def save_sliced_contract_pdf(self):
        """
        업체 계약서 지정 페이지(예: 12-13 -> 2페이지)만 별도 PDF 추출 저장
        """
        contract_pdf = self.contract_pdf_path.get()
        if not contract_pdf or not os.path.exists(contract_pdf):
            self.set_live_status("⚠️ 추출할 업체 계약서 PDF 파일이 올려지지 않았습니다.", type="error")
            return

        page_str = self.contract_page.get().strip()
        page_indices, label_str = self._parse_contract_pages(page_str)

        try:
            from pypdf import PdfReader, PdfWriter
            reader = PdfReader(contract_pdf)
            writer = PdfWriter()
            total_pages = len(reader.pages)

            valid_pages = [p for p in page_indices if 0 <= p < total_pages]
            if not valid_pages:
                self.set_live_status(f"⚠️ 유효하지 않은 페이지 범위입니다. (총 {total_pages}페이지)", type="error")
                return

            for p in valid_pages:
                writer.add_page(reader.pages[p])

            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            out_name = f"계약서_{label_str}p.pdf"
            out_path = os.path.join(desktop, out_name)

            with open(out_path, 'wb') as f_out:
                writer.write(f_out)

            self.set_live_status(f"🎉 업체 계약서 [{label_str}p] ({len(valid_pages)}페이지) 추출 저장 완료! 📂 {out_name}", type="success")
        except Exception as e:
            self.set_live_status(f"⚠️ 계약서 페이지 추출 오류: {e}", type="error")

    def create_desktop_shortcut(self):
        """
        Windows 정식 바로가기 (.lnk) 생성
        """
        try:
            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dist', 'VoucherPass.exe')
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            lnk_path = os.path.join(desktop, 'VoucherPass.lnk')

            vbs_script = f'''
Set WshShell = WScript.CreateObject("WScript.Shell")
Set shortcut = WshShell.CreateShortcut("{lnk_path}")
shortcut.TargetPath = "{exe_path}"
shortcut.WorkingDirectory = "{os.path.dirname(exe_path)}"
shortcut.Description = "VoucherPass Tax & PR Extractor"
shortcut.WindowStyle = 1
shortcut.Save
'''
            vbs_file = os.path.join(os.environ.get('TEMP', '.'), 'create_lnk.vbs')
            with open(vbs_file, 'w', encoding='utf-8') as f:
                f.write(vbs_script)

            subprocess.run(['cscript', '//Nologo', vbs_file], check=True)
            self.set_live_status("🖥️ 바탕화면에 VoucherPass.lnk 바로가기가 생성되었습니다!", type="success")
        except Exception as e:
            print(f"Shortcut creation error: {e}")

    def _click_title(self, event):
        self._offsetx = event.x
        self._offsety = event.y

    def _drag_title(self, event):
        x = self.root.winfo_pointerx() - self._offsetx
        y = self.root.winfo_pointery() - self._offsety
        self.root.geometry(f"+{x}+{y}")

    def _update_progress_summary(self):
        """
        업로드 진행 상태 실시간 집계 & 막대바(Progress Bar) 채우기
        """
        count = sum(1 for var in [self.tax_pdf_path, self.spec_pdf_path, self.pr_pdf_path, self.po_pdf_path, self.contract_pdf_path] if var.get() and os.path.exists(var.get()))
        pct = count * 20
        if hasattr(self, 'progress_bar'):
            self.progress_bar['value'] = pct

        if hasattr(self, 'lbl_progress'):
            if count == 5:
                self.lbl_progress.config(text="진행 상태: 5/5 완료 🎉", fg="#15803D", bg="#DCFCE7")
            else:
                self.lbl_progress.config(text=f"진행 상태: {count}/5 완료", fg="#1D4ED8", bg="#DBEAFE")

    def _build_widget_layout(self):
        # 1. Header Bar (HPNT Style Blue Header & Progress Badge & Progressbar)
        hdr = tk.Frame(self.root, bg="#2563EB", padx=8, pady=5)
        hdr.pack(fill="x")
        hdr.bind("<Button-1>", self._click_title)
        hdr.bind("<B1-Motion>", self._drag_title)

        lbl_logo = tk.Label(hdr, text="⚡ Voucher Pass", font=("Malgun Gothic", 10, "bold"), bg="#2563EB", fg="#FFFFFF")
        lbl_logo.pack(side="left")
        lbl_logo.bind("<Button-1>", self._click_title)
        lbl_logo.bind("<B1-Motion>", self._drag_title)

        ver_b = tk.Label(hdr, text="v8.4.2", font=("Malgun Gothic", 8, "bold"), bg="#1D4ED8", fg="white", padx=4, pady=1)
        ver_b.pack(side="left", padx=(4, 0))

        # 업로드 진행 상태 뱃지 ("1/5 완료") 및 초록색 프로그레스 막대바
        self.lbl_progress = tk.Label(hdr, text="진행 상태: 0/5 완료", font=("Malgun Gothic", 8, "bold"), bg="#DBEAFE", fg="#1D4ED8", padx=4, pady=1)
        self.lbl_progress.pack(side="left", padx=(4, 2))

        # 프로그레스 막대바 채우기 요소
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Green.Horizontal.TProgressbar", foreground='#16A34A', background='#22C55E', thickness=10)
        self.progress_bar = ttk.Progressbar(hdr, style="Green.Horizontal.TProgressbar", orient="horizontal", length=65, mode="determinate")
        self.progress_bar.pack(side="left", padx=(0, 2))

        btn_min = tk.Label(hdr, text=" ─ ", font=("Arial", 10, "bold"), bg="#2563EB", fg="#DBEAFE", cursor="hand2")
        btn_min.pack(side="right", padx=(3, 0))
        btn_min.bind("<Button-1>", lambda e: self.root.withdraw())

        btn_close = tk.Label(hdr, text=" ✕ ", font=("Arial", 10, "bold"), bg="#2563EB", fg="#FEE2E2", cursor="hand2")
        btn_close.pack(side="right")
        btn_close.bind("<Button-1>", lambda e: self.root.destroy())

        # 2. Main Content Canvas
        main_box = tk.Frame(self.root, bg="#F8FAFC", padx=3, pady=2)
        main_box.pack(fill="both", expand=True)

        # 3. 5가지 서류 Clean DropZone
        self.drop_tax = CleanMinimalDropZone(main_box, "① 전자 세금계산서 (6068625399)", "🧾", "#7C3AED", self.tax_pdf_path, on_file_selected=self.parse_tax_invoice_uploaded, on_state_changed=self._update_progress_summary)
        self.drop_tax.pack(fill="x", pady=1)

        self.drop_spec = CleanMinimalDropZone(main_box, "② 거래명세서 PDF", "📄", "#059669", self.spec_pdf_path, on_state_changed=self._update_progress_summary)
        self.drop_spec.pack(fill="x", pady=1)

        self.drop_pr = CleanMinimalDropZone(main_box, "③ PR Print (구매요청서) PDF", "🛒", "#2563EB", self.pr_pdf_path, on_file_selected=self.parse_uploaded_pdf, on_state_changed=self._update_progress_summary)
        self.drop_pr.pack(fill="x", pady=1)

        self.drop_po = CleanMinimalDropZone(main_box, "④ 발주서 (PO) PDF", "📦", "#0284C7", self.po_pdf_path, on_file_selected=self.parse_uploaded_pdf, on_state_changed=self._update_progress_summary)
        self.drop_po.pack(fill="x", pady=1)

        f_contract_f = tk.Frame(main_box, bg="#F8FAFC")
        f_contract_f.pack(fill="x", pady=1)

        self.drop_contract = CleanMinimalDropZone(f_contract_f, "⑤ 업체 계약서 PDF", "📝", "#D97706", self.contract_pdf_path, on_state_changed=self._update_progress_summary)
        self.drop_contract.pack(side="left", fill="x", expand=True)

        # 계약서 전용 페이지 지정 (예: 12-13)
        pg_sub = tk.Frame(f_contract_f, bg="#FFFBEB", highlightbackground="#FCD34D", highlightthickness=1, padx=3, pady=2)
        pg_sub.pack(side="right", padx=(2, 0))
        tk.Label(pg_sub, text="📄 페이지:", font=("Malgun Gothic", 7, "bold"), bg="#FFFBEB", fg="#92400E").pack(side="top")
        self.e_contract_page = tk.Entry(pg_sub, textvariable=self.contract_page, font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#B45309", width=6, justify="center", relief="solid", bd=1)
        self.e_contract_page.pack(side="top", pady=1)
        tk.Button(pg_sub, text="💾 추출저장", font=("Malgun Gothic", 7, "bold"), bg="#F59E0B", fg="white", relief="flat", padx=2, pady=1, cursor="hand2", command=self.save_sliced_contract_pdf).pack(side="top")

        # 4. 데이터 입력 영역 (2열 4행 그리드 + 상단 라벨 카드 컨테이너)
        hud = tk.LabelFrame(main_box, text=" 📝 추출 데이터 7종 ", font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#0F172A", bd=1, relief="solid", padx=4, pady=2)
        hud.pack(fill="x", pady=(2, 1))

        lbl_s = {"font": ("Malgun Gothic", 8, "bold"), "bg": "#FFFFFF", "fg": "#475569"}
        ent_s = {"font": ("Malgun Gothic", 9, "bold"), "bg": "#F8FAFC", "fg": "#0F172A", "insertbackground": "#0F172A", "relief": "solid", "bd": 1}
        btn_cp = {"font": ("Malgun Gothic", 7, "bold"), "bg": "#EFF6FF", "fg": "#2563EB", "relief": "flat", "padx": 2, "cursor": "hand2"}

        grid_f = tk.Frame(hud, bg="#FFFFFF")
        grid_f.pack(fill="x", pady=1)

        # Row 0: Labels Top (P/R No | 📅 작성일자)
        f_c1_r1 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c1_r1.grid(row=0, column=0, sticky="ew", padx=2, pady=1)
        tk.Label(f_c1_r1, text="P/R No:", **lbl_s).pack(anchor="w")
        self.e_prno = tk.Entry(f_c1_r1, textvariable=self.pr_no_var, width=14, **ent_s)
        self.e_prno.pack(fill="x")

        f_c2_r1 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c2_r1.grid(row=0, column=1, sticky="ew", padx=2, pady=1)
        tk.Label(f_c2_r1, text="📅 작성일자:", **lbl_s).pack(anchor="w")
        f_date_box = tk.Frame(f_c2_r1, bg="#FFFFFF")
        f_date_box.pack(fill="x")
        self.e_date = tk.Entry(f_date_box, textvariable=self.date_var, font=("Malgun Gothic", 9, "bold"), bg="#F0FDF4", fg="#15803D", width=11, relief="solid", bd=1)
        self.e_date.pack(side="left", fill="x", expand=True)
        tk.Button(f_date_box, text="📋", command=lambda: self.copy_to_clipboard(self.date_var.get(), "작성일자"), **btn_cp).pack(side="right", padx=(2, 0))

        # Row 1: Labels Top (📌 PR Title | 💰 공급가액)
        f_c1_r2 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c1_r2.grid(row=1, column=0, sticky="ew", padx=2, pady=1)
        f_t_box = tk.Frame(f_c1_r2, bg="#FFFFFF")
        f_t_box.pack(fill="x")
        tk.Label(f_t_box, text="📌 PR Title:", **lbl_s).pack(side="left")
        tk.Button(f_t_box, text="📋", command=lambda: self.copy_to_clipboard(self.pr_title_var.get(), "PR Title"), **btn_cp).pack(side="right")
        self.txt_pr_title = tk.Text(f_c1_r2, width=15, height=2, wrap="word", font=("Malgun Gothic", 9, "bold"), bg="#F8FAFC", fg="#0F172A", insertbackground="#0F172A", relief="solid", bd=1)
        self.txt_pr_title.pack(fill="x")
        self.txt_pr_title.bind("<KeyRelease>", self._on_pr_title_txt_changed)

        f_c2_r2 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c2_r2.grid(row=1, column=1, sticky="ew", padx=2, pady=1)
        tk.Label(f_c2_r2, text="💰 공급가액:", **lbl_s).pack(anchor="w")
        f_amt_box = tk.Frame(f_c2_r2, bg="#FFFFFF")
        f_amt_box.pack(fill="x")
        self.e_amt = tk.Entry(f_amt_box, textvariable=self.amount_var, font=("Malgun Gothic", 9, "bold"), bg="#EFF6FF", fg="#1D4ED8", width=11, relief="solid", bd=1)
        self.e_amt.pack(side="left", fill="x", expand=True)
        self.e_amt.bind("<KeyRelease>", self._recalc_amounts)
        tk.Button(f_amt_box, text="📋", command=lambda: self.copy_to_clipboard(self.amount_var.get(), "공급가액"), **btn_cp).pack(side="right", padx=(2, 0))

        # Row 2: Labels Top (🏢 거래처명 | 💵 부가세)
        f_c1_r3 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c1_r3.grid(row=2, column=0, sticky="ew", padx=2, pady=1)
        tk.Label(f_c1_r3, text="🏢 거래처명:", **lbl_s).pack(anchor="w")
        self.e_sup = tk.Entry(f_c1_r3, textvariable=self.supplier_var, width=14, **ent_s)
        self.e_sup.pack(fill="x")

        f_c2_r3 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c2_r3.grid(row=2, column=1, sticky="ew", padx=2, pady=1)
        tk.Label(f_c2_r3, text="💵 부가세:", **lbl_s).pack(anchor="w")
        self.e_vat = tk.Entry(f_c2_r3, textvariable=self.vat_var, width=14, **ent_s)
        self.e_vat.pack(fill="x")

        # Row 3: Labels Top (💳 합계금액)
        f_c2_r4 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c2_r4.grid(row=3, column=1, sticky="ew", padx=2, pady=1)
        tk.Label(f_c2_r4, text="💳 합계금액:", **lbl_s).pack(anchor="w")
        self.e_tot = tk.Entry(f_c2_r4, textvariable=self.total_amount_var, font=("Malgun Gothic", 9, "bold"), bg="#EFF6FF", fg="#1D4ED8", width=14, relief="solid", bd=1)
        self.e_tot.pack(fill="x")

        grid_f.columnconfigure(0, weight=1)
        grid_f.columnconfigure(1, weight=1)

        # Live Interactive Status Banner
        r5 = tk.Frame(hud, bg="#F0FDF4", highlightbackground="#86EFAC", highlightthickness=1, padx=4, pady=2)
        r5.pack(fill="x", pady=(3, 1))

        self.lbl_live_status = tk.Label(r5, text="✨ 서류를 올려주시면 데이터가 0.1초 만에 쏙! 추출됩니다.", font=("Malgun Gothic", 8, "bold"), bg="#F0FDF4", fg="#15803D", anchor="w")
        self.lbl_live_status.pack(fill="x")

        # 5. 버튼 계층 정리 (메인 강조 버튼 1종 + 보조 2종 나란히)
        act_panel = tk.Frame(main_box, bg="#F8FAFC")
        act_panel.pack(fill="x", pady=(1, 0))

        # 메인 주 동작 버튼 (엑셀 붙여넣기 - 프리미엄 로열 블루 강조)
        btn_copy_all = tk.Button(act_panel, text="📋 Voucher 엑셀 양식 붙여넣기 (클립보드 복사)", font=("Malgun Gothic", 9, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", relief="flat", padx=6, pady=6, cursor="hand2", command=self.copy_all_3items)
        btn_copy_all.pack(side="top", fill="x", pady=1)

        # 보조 동작 버튼 2종 (보관 & 인쇄 1:1 동등 비중 나란히)
        bot_btn_f = tk.Frame(act_panel, bg="#F8FAFC")
        bot_btn_f.pack(fill="x", pady=1)

        btn_arch = tk.Button(bot_btn_f, text="📂 건별 자동 보관", font=("Malgun Gothic", 9, "bold"), bg="#475569", fg="white", activebackground="#334155", activeforeground="white", relief="flat", padx=4, pady=4, cursor="hand2", command=self.archive_voucher_files)
        btn_arch.pack(side="left", fill="x", expand=True, padx=(0, 1))

        btn_print = tk.Button(bot_btn_f, text="🖨️ 서류 5종 일괄 인쇄", font=("Malgun Gothic", 9, "bold"), bg="#059669", fg="white", activebackground="#047857", activeforeground="white", relief="flat", padx=4, pady=4, cursor="hand2", command=self.print_pdf_documents_only)
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
        """
        지정 폴더(다운로드/바탕화면)에 새로운 PDF가 감지되면 자동으로 파싱 및 분류
        """
        if not pdf_path or not os.path.exists(pdf_path):
            return

        pdf_type = pdf_parser.classify_pdf_type(pdf_path)
        if pdf_type == 'tax' or '세금' in os.path.basename(pdf_path) or '조회' in os.path.basename(pdf_path):
            self.drop_tax.set_file(pdf_path)
            self.parse_tax_invoice_uploaded(pdf_path)
        elif pdf_type == 'pr':
            self.drop_pr.set_file(pdf_path)
            self.parse_uploaded_pdf(pdf_path)
        elif pdf_type == 'po':
            self.drop_po.set_file(pdf_path)
            self.parse_uploaded_pdf(pdf_path)
        elif pdf_type == 'spec':
            self.drop_spec.set_file(pdf_path)
        elif pdf_type == 'contract':
            self.drop_contract.set_file(pdf_path)

    def generate_excel_action(self):
        data = self._get_form_data()
        if not data.get('pr_no') and not data.get('pr_title') and not data.get('amount'):
            messagebox.showwarning("엑셀 생성 실패", "기록할 데이터가 없습니다. PDF 서류를 먼저 업로드해 주세요.")
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
            self.set_live_status(f"⚠️ {label_name} 값이 비어 있습니다.", type="error")
            return
        
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self.set_live_status(f"📋 [{label_name}] 텍스트 복사 완료: {text}", type="success")

    def copy_all_3items(self):
        d = self.date_var.get().strip()
        t = self.pr_title_var.get().strip()
        a = self.amount_var.get().strip()

        if not d and not t and not a:
            self.set_live_status("⚠️ 복사할 데이터 항목이 없습니다. 서류를 먼저 올려주세요.", type="error")
            return

        combined_text = f"{d}\n{t}\n{a}"
        self.root.clipboard_clear()
        self.root.clipboard_append(combined_text)
        self.root.update()

        self.set_live_status(f"🎉 Voucher 엑셀 양식 3종 항목 복사 완료! (엑셀 셀 선택 후 Ctrl+V)", type="success")

    def _load_printers(self):
        printers = printer_handler.get_installed_printers()
        self.printer_combo['values'] = printers

        # 저장된 마지막 프린터 설정 로드
        cfg_file = self._get_config_path()
        last_printer = ''
        if os.path.exists(cfg_file):
            try:
                import json
                with open(cfg_file, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                    last_printer = cfg.get('last_printer', '')
            except Exception:
                pass

        if last_printer and last_printer in printers:
            self.selected_printer.set(last_printer)
        elif printers:
            self.selected_printer.set(printers[0])

        self.selected_printer.trace_add("write", self._on_printer_selected)

    def _on_printer_selected(self, *args):
        p = self.selected_printer.get()
        if p:
            cfg_file = self._get_config_path()
            try:
                import json
                with open(cfg_file, 'w', encoding='utf-8') as f:
                    json.dump({'last_printer': p}, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

    def set_live_status(self, msg, type="info"):
        if not hasattr(self, 'lbl_live_status'):
            return
        
        bg_col = "#F0FDF4" if type == "success" else "#EFF6FF" if type == "info" else "#FEF2F2"
        fg_col = "#15803D" if type == "success" else "#1D4ED8" if type == "info" else "#B91C1C"
        
        self.lbl_live_status.master.config(bg=bg_col)
        self.lbl_live_status.config(text=msg, bg=bg_col, fg=fg_col)

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

            if hasattr(self, 'txt_pr_title'):
                self._shake_widget(self.txt_pr_title, highlight_bg="#BAE6FD")
            if hasattr(self, 'e_amt'):
                self._shake_widget(self.e_amt, highlight_bg="#93C5FD")

            self.set_live_status(f"🚀 PR 서류 Title & 공급가액 {amt:,}원 추출 획득 완료! 💎", type="success")
        except Exception as e:
            self.set_live_status(f"⚠️ PR 서류 파싱 실패: {e}", type="error")

    def parse_tax_invoice_uploaded(self, tax_path=None):
        """
        1) HTML/PDF 파일 업로드 감지
        2) HTML일 경우 자동 비밀번호 입력 및 지정 폴더 PDF 저장
        3) 저장된 PDF 파일에서 작성일자(2026/08/11) 자동 파싱
        """
        if not tax_path:
            tax_path = self.tax_pdf_path.get()

        pdf_parser._log_debug(f"parse_tax_invoice_uploaded called with: {tax_path}")

        if not tax_path or not os.path.exists(tax_path):
            pdf_parser._log_debug(f"parse_tax_invoice_uploaded: File missing or invalid: {tax_path}")
            return

        # Step 1 & 2: HTML 세금계산서일 경우 HTML 텍스트에서 즉시 작성일자 1차 파싱 + 자동 해제 & PDF 저장 파이프라인 가동
        if pdf_parser._is_html_file(tax_path):
            pdf_parser._log_debug("Detected HTML file. Attempting immediate HTML text date parsing...")
            try:
                html_date = pdf_parser.parse_tax_invoice_date(tax_path)
                pdf_parser._log_debug(f"Immediate HTML date parsing result: {html_date}")
                if html_date:
                    self.date_var.set(html_date)
                    if hasattr(self, 'e_date'):
                        self._shake_widget(self.e_date, highlight_bg="#FEF08A")
                    self.set_live_status(f"🎉 HTML 작성일자 [{html_date}] 즉시 획득 성공! ✨", type="success")
            except Exception as e_html:
                pdf_parser._log_debug(f"Immediate HTML date parsing exception: {e_html}")

            pdf_parser._log_debug("Triggering auto_unlock_and_save_pdf for PDF conversion...")
            self.auto_unlock_and_save_pdf(tax_path)
            return

        # Step 3: PDF 파일에서 작성일자 파싱
        try:
            tax_date = pdf_parser.parse_tax_invoice_date(tax_path)
            pdf_parser._log_debug(f"parse_tax_invoice_date result for PDF: {tax_date}")
            if tax_date:
                self.date_var.set(tax_date)
                if hasattr(self, 'e_date'):
                    self._shake_widget(self.e_date, highlight_bg="#FEF08A")
                self.root.update_idletasks()
                self.root.update()
                self.set_live_status(f"🎉 작성일자 [{tax_date}] 자동 획득 성공! ✨", type="success")
            else:
                self.set_live_status("⚠️ 세금계산서 작성일자를 찾지 못했습니다. 수동 입력해 주세요.", type="error")
        except Exception as e:
            pdf_parser._log_debug(f"parse_tax_invoice_uploaded exception: {e}")
            self.set_live_status(f"⚠️ 파싱 오류: {e}", type="error")

    def auto_unlock_and_save_pdf(self, html_path):
        """Selenium headless Edge + CDP Page.printToPDF 방식으로
        NTS 세금계산서 HTML을 비밀번호 복호화 후 PDF로 변환.
        브라우저 종류/해상도/프린터 설정에 완전 독립적."""
        import threading, time, tempfile

        def _worker():
            driver = None
            try:
                from selenium import webdriver
                from selenium.webdriver.edge.options import Options
                from selenium.webdriver.common.by import By
                import base64

                pdf_parser._log_debug(f"[auto_unlock] Selenium headless 시작: {html_path}")
                self.root.after(0, lambda: self.set_live_status("🔐 세금계산서 복호화 중...", type="info"))

                # Step 1: headless Edge 브라우저 실행
                options = Options()
                options.add_argument('--headless=new')
                options.add_argument('--disable-gpu')
                options.add_argument('--no-sandbox')
                driver = webdriver.Edge(options=options)

                # Step 2: HTML 파일 열기
                file_uri = 'file:///' + html_path.replace('\\', '/').replace(' ', '%20')
                pdf_parser._log_debug(f"[auto_unlock 1단계] HTML 열기: {file_uri}")
                driver.get(file_uri)
                time.sleep(2)

                # Step 3: 비밀번호 입력
                pdf_parser._log_debug("[auto_unlock 2단계] 비밀번호(6068625399) 입력...")
                pwd_input = None
                inputs = driver.find_elements(By.CSS_SELECTOR, 'input[type=password], input[type=text]')
                for inp in inputs:
                    if inp.is_displayed() and inp.get_attribute('type') in ('password', 'text'):
                        pwd_input = inp
                        break

                if not pwd_input:
                    pdf_parser._log_debug("[auto_unlock ERROR] 비밀번호 입력란을 찾지 못했습니다.")
                    self.root.after(0, lambda: self.set_live_status("⚠️ 비밀번호 입력란 미발견", type="error"))
                    return

                pwd_input.clear()
                pwd_input.send_keys('6068625399')

                # Step 4: 확인 버튼 클릭
                pdf_parser._log_debug("[auto_unlock 3단계] 확인 버튼 클릭...")
                buttons = driver.find_elements(By.CSS_SELECTOR, 'button, input[type=button], input[type=submit], a[onclick]')
                clicked = False
                for btn in buttons:
                    if btn.is_displayed():
                        onclick = btn.get_attribute('onclick') or ''
                        text = btn.text or ''
                        if 'Cri' in onclick or 'InputPwd' in onclick or '확인' in text:
                            btn.click()
                            clicked = True
                            pdf_parser._log_debug(f"[auto_unlock 3단계] 클릭됨: text='{text}', onclick='{onclick[:50]}'")
                            break
                if not clicked:
                    for btn in buttons:
                        if btn.is_displayed():
                            btn.click()
                            clicked = True
                            break

                # Step 5: 세금계산서 렌더링 대기
                pdf_parser._log_debug("[auto_unlock 4단계] 세금계산서 본문 렌더링 대기...")
                self.root.after(0, lambda: self.set_live_status("📄 세금계산서 렌더링 중...", type="info"))
                time.sleep(3)

                # Step 6: 인쇄 UI 정리 (상단 '인쇄/첨부보기' 툴바 및 비밀번호 대화상자 숨김)
                try:
                    driver.execute_script("""
                        var btnBar = document.getElementById('CriBtnPosition');
                        if (btnBar) btnBar.style.display = 'none';
                        var pwdDlg = document.getElementById('idPcPwdDlg');
                        if (pwdDlg) pwdDlg.style.display = 'none';
                        var mobDlg = document.getElementById('idMobilePwdDlg');
                        if (mobDlg) mobDlg.style.display = 'none';
                    """)
                    driver.execute_cdp_cmd('Emulation.setEmulatedMedia', {'media': 'print'})
                except Exception as e_em:
                    pdf_parser._log_debug(f"[auto_unlock 4단계] UI 정리/미디어 설정 보조 예외: {e_em}")

                # Step 7: CDP Page.printToPDF로 온전한 A4 세금계산서 PDF 생성 (하단 잘림 방지 scale: 0.95)
                pdf_parser._log_debug("[auto_unlock 5단계] CDP Page.printToPDF 호출 (scale=0.95, A4 완벽 비율)...")
                self.root.after(0, lambda: self.set_live_status("🖨️ PDF 생성 중...", type="info"))

                params = {
                    'landscape': False,
                    'displayHeaderFooter': False,
                    'printBackground': True,
                    'preferCSSPageSize': False,
                    'paperWidth': 8.27,
                    'paperHeight': 11.69,
                    'marginTop': 0.15,
                    'marginBottom': 0.15,
                    'marginLeft': 0.15,
                    'marginRight': 0.15,
                    'scale': 0.95
                }
                result = driver.execute_cdp_cmd('Page.printToPDF', params)

                # Step 8: PDF 파일 저장
                tempdir = tempfile.gettempdir()
                expected_pdf_path = os.path.join(tempdir, f"NTS_eTaxInvoice_{int(time.time())}.pdf")
                with open(expected_pdf_path, 'wb') as f:
                    f.write(base64.b64decode(result['data']))

                pdf_size = os.path.getsize(expected_pdf_path)
                pdf_parser._log_debug(f"[auto_unlock SUCCESS] PDF 저장 완료: {expected_pdf_path} ({pdf_size} bytes)")

                if pdf_size > 0:
                    self.root.after(0, lambda: self.root.lift())
                    self.root.after(0, lambda p=expected_pdf_path: self._on_new_tax_pdf_saved(p))
                else:
                    pdf_parser._log_debug("[auto_unlock ERROR] PDF 파일 크기가 0입니다.")
                    self.root.after(0, lambda: self.set_live_status("⚠️ PDF 생성 실패 (0 bytes)", type="error"))

            except Exception as e:
                pdf_parser._log_debug(f"[auto_unlock ERROR] Exception: {e}")
                print(f"Auto unlock & save PDF error: {e}")
                self.root.after(0, lambda: self.root.lift())
                self.root.after(0, lambda: self.set_live_status(f"⚠️ PDF 변환 오류: {e}", type="error"))

            finally:
                if driver:
                    try:
                        driver.quit()
                    except Exception:
                        pass

        threading.Thread(target=_worker, daemon=True).start()

    def _on_new_tax_pdf_saved(self, pdf_path):
        if not pdf_path or not os.path.exists(pdf_path):
            pdf_parser._log_debug(f"_on_new_tax_pdf_saved: Invalid path: {pdf_path}")
            return

        def _parse_job():
            import time
            pdf_parser._log_debug(f"_on_new_tax_pdf_saved parse job started for: {pdf_path}")
            # 0.05초 단위 초고속 파일 안정화 대기
            for _ in range(10):
                try:
                    if os.path.getsize(pdf_path) > 100:
                        break
                except Exception:
                    pass
                time.sleep(0.05)

            self.drop_tax.set_file(pdf_path)

            tax_date = pdf_parser.parse_tax_invoice_date(pdf_path)
            if not tax_date:
                time.sleep(0.1)
                pdf_parser._log_debug("Retrying parse_tax_invoice_date second attempt...")
                tax_date = pdf_parser.parse_tax_invoice_date(pdf_path)

            pdf_parser._log_debug(f"_on_new_tax_pdf_saved final tax_date: {tax_date}")
            if tax_date:
                self.root.after(0, lambda d=tax_date: self._apply_tax_date(d))
            else:
                self.root.after(0, lambda: messagebox.showwarning("작성일자 미추출", f"저장된 PDF 세금계산서 [{os.path.basename(pdf_path)}] 에서 작성일자를 발견하지 못했습니다.\n\n(자세한 원인은 voucher_pass_debug.log 파일을 참고하세요)"))

        threading.Thread(target=_parse_job, daemon=True).start()

    def _apply_tax_date(self, tax_date):
        self.date_var.set(tax_date)
        if hasattr(self, 'e_date'):
            self._shake_widget(self.e_date, highlight_bg="#FEF08A")
        self.set_live_status(f"🎉 작성일자 [{tax_date}] 자동 획득 성공! ✨", type="success")

    def _shake_widget(self, widget, highlight_bg="#FEF08A"):
        """
        추출된 데이터 셀창이 좌-우-좌-우 통통통~ 재미있게 흔들리는 쉐이크 애니메이션
        """
        if not widget or not os.path.exists if not hasattr(widget, 'config') else False:
            pass

        try:
            orig_bg = widget.cget("bg")
            widget.config(bg=highlight_bg)

            # 좌우 오프셋 패턴 (Shake Sequence)
            offsets = [5, -5, 4, -4, 2, -2, 0]

            def _animate(idx):
                if idx < len(offsets):
                    dx = offsets[idx]
                    try:
                        widget.pack_configure(padx=(2 + dx, 2 - dx))
                    except Exception:
                        pass
                    self.root.after(25, lambda: _animate(idx + 1))
                else:
                    try:
                        widget.pack_configure(padx=(2, 2))
                        widget.config(bg=orig_bg)
                    except Exception:
                        pass

            _animate(0)
        except Exception:
            pass

    def _apply_tax_date(self, tax_date):
        self.date_var.set(tax_date)
        self.root.update_idletasks()
        self.root.update()
        if hasattr(self, 'e_date'):
            self._shake_widget(self.e_date, highlight_bg="#86EFAC")
        self.set_live_status(f"🎉 작성일자 [{tax_date}] 추출 완료! (셀이 통통~ 튀어올랐어요)", type="success")

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
        failed_list = []

        try:
            if self.tax_pdf_path.get() and os.path.exists(self.tax_pdf_path.get()):
                try:
                    dec_pdf = pdf_parser.decrypt_pdf_to_temp(self.tax_pdf_path.get())
                    ok = printer_handler.print_pdf_file(dec_pdf, printer_name=printer)
                    if ok:
                        printed_list.append("① 전자 세금계산서 PDF")
                    else:
                        failed_list.append("① 전자 세금계산서 PDF")
                except Exception as e1:
                    print(f"Tax print error: {e1}")
                    failed_list.append("① 전자 세금계산서 PDF")

            if self.spec_pdf_path.get() and os.path.exists(self.spec_pdf_path.get()):
                try:
                    ok = printer_handler.print_pdf_file(self.spec_pdf_path.get(), printer_name=printer)
                    if ok:
                        printed_list.append("② 거래명세서 PDF")
                    else:
                        failed_list.append("② 거래명세서 PDF")
                except Exception as e2:
                    print(f"Spec print error: {e2}")
                    failed_list.append("② 거래명세서 PDF")

            if self.pr_pdf_path.get() and os.path.exists(self.pr_pdf_path.get()):
                try:
                    ok = printer_handler.print_pdf_file(self.pr_pdf_path.get(), printer_name=printer)
                    if ok:
                        printed_list.append("③ PR Print PDF (구매요청서)")
                    else:
                        failed_list.append("③ PR Print PDF (구매요청서)")
                except Exception as e3:
                    print(f"PR print error: {e3}")
                    failed_list.append("③ PR Print PDF (구매요청서)")

            if self.po_pdf_path.get() and os.path.exists(self.po_pdf_path.get()):
                try:
                    ok = printer_handler.print_pdf_file(self.po_pdf_path.get(), printer_name=printer)
                    if ok:
                        printed_list.append("④ 발주서 (PO) PDF")
                    else:
                        failed_list.append("④ 발주서 (PO) PDF")
                except Exception as e4:
                    print(f"PO print error: {e4}")
                    failed_list.append("④ 발주서 (PO) PDF")

            if self.contract_pdf_path.get() and os.path.exists(self.contract_pdf_path.get()):
                try:
                    page_str = self.contract_page.get().strip()
                    page_indices, label_str = self._parse_contract_pages(page_str)
                    ok = printer_handler.print_pdf_file(self.contract_pdf_path.get(), printer_name=printer, page_range=page_indices)
                    if ok:
                        printed_list.append(f"⑤ 업체 계약서 PDF ({label_str} 페이지)")
                    else:
                        failed_list.append(f"⑤ 업체 계약서 PDF ({label_str} 페이지)")
                except Exception as e5:
                    print(f"Contract print error: {e5}")
                    failed_list.append(f"⑤ 업체 계약서 PDF ({label_str} 페이지)")

            if not printed_list and not failed_list:
                self.set_live_status("⚠️ 인쇄할 PDF 서류가 업로드되지 않았습니다.", type="error")
                return

            if printed_list:
                summary = ", ".join(printed_list)
                if failed_list:
                    fail_summary = ", ".join(failed_list)
                    self.set_live_status(f"⚠️ 일부 인쇄 성공 [{summary}], 실패 [{fail_summary}] (프린터: {printer})", type="error")
                else:
                    self.set_live_status(f"🖨️ 서류 일괄 인쇄 요청 완료: [{summary}] (프린터: {printer})", type="success")
            else:
                fail_summary = ", ".join(failed_list)
                self.set_live_status(f"⚠️ 인쇄 실패: [{fail_summary}] (프린터: {printer})", type="error")

        except Exception as e:
            self.set_live_status(f"⚠️ 인쇄 오류: {e}", type="error")

def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = VoucherPassApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
