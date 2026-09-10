import os
import sys
import re
import urllib.parse
import subprocess
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw

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
import outlook_handler
from voucher_core import DocumentState, parse_money, parse_pages, validate_pages
from file_utils import desktop_dir, temp_output, safe_name
from task_runner import TaskRunner
from tax_converter import convert_tax_html


class RoundedFrame(tk.Frame):
    """PIL 안티앨리어싱 둥근 모서리 배경을 가진 Tk Frame. 자식 위젯 크기에 따라 자동 높이 조절."""
    def __init__(self, parent, radius=8, bg_color="#FFFFFF", border_color="#E2E8F0", border_width=1, outer_bg="#EEF1F6", **kw):
        super().__init__(parent, bg=outer_bg, bd=0, highlightthickness=0, **kw)
        self.radius = radius
        self._rf_bg = bg_color
        self._rf_border = border_color
        self._rf_bw = border_width
        self._outer_bg = outer_bg
        self._bg_img = None
        self._bg_lbl = tk.Label(self, bg=outer_bg, bd=0, highlightthickness=0)
        self._bg_lbl.place(x=0, y=0, relwidth=1, relheight=1)
        self._bg_lbl.lower()
        self.bind("<Configure>", self._rf_redraw)

    def _rf_redraw(self, event=None):
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        scale = 3
        W, H, R = w * scale, h * scale, self.radius * scale
        img = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle(
            [0, 0, W - 1, H - 1],
            radius=R,
            fill=self._rf_bg,
            outline=self._rf_border,
            width=max(1, int(self._rf_bw * scale)) if self._rf_border else 0
        )
        self._bg_img = ImageTk.PhotoImage(img.resize((w, h), Image.Resampling.LANCZOS))
        self._bg_lbl.configure(image=self._bg_img)
        self._bg_lbl.lower()

    def set_colors(self, bg_color=None, border_color=None, border_width=None):
        if bg_color is not None:
            self._rf_bg = bg_color
        if border_color is not None:
            self._rf_border = border_color
        if border_width is not None:
            self._rf_bw = border_width
        self._rf_redraw()


class RoundedBadge(tk.Canvas):
    """안티앨리어싱 둥근 칩/배지"""
    def __init__(self, parent, text, bg_color="#C94A3C", fg_color="#FFFFFF", font=("Malgun Gothic", 8, "bold"), radius=6, width=28, height=22, border_color=None, outer_bg="#FFFFFF", **kw):
        super().__init__(parent, width=width, height=height, bg=outer_bg, highlightthickness=0, bd=0, **kw)
        self.text = text
        self.bg_color = bg_color
        self.fg_color = fg_color
        self.font = font
        self.radius = radius
        self.w = width
        self.h = height
        self.border_color = border_color
        self.outer_bg = outer_bg
        self._img = None
        self.redraw()

    def set_content(self, text=None, bg_color=None, fg_color=None, border_color=None):
        if text is not None:
            self.text = text
        if bg_color is not None:
            self.bg_color = bg_color
        if fg_color is not None:
            self.fg_color = fg_color
        if border_color is not None:
            self.border_color = border_color
        self.redraw()

    def redraw(self):
        self.config(bg=self.outer_bg)
        scale = 3
        W, H, R = self.w * scale, self.h * scale, self.radius * scale
        img = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle(
            [0, 0, W - 1, H - 1],
            radius=R,
            fill=self.bg_color,
            outline=self.border_color,
            width=1 * scale if self.border_color else 0
        )
        self._img = ImageTk.PhotoImage(img.resize((self.w, self.h), Image.Resampling.LANCZOS))
        self.delete("all")
        self.create_image(0, 0, image=self._img, anchor="nw")
        self.create_text(self.w / 2, self.h / 2, text=self.text, fill=self.fg_color, font=self.font)


class CircularProgressRing(tk.Canvas):
    """원형 도넛 프로그레스 링 (0/5)"""
    def __init__(self, parent, size=28, bg="#3457A8", track_color="#4B6CB7", prog_color="#22C55E", **kwargs):
        super().__init__(parent, width=size, height=size, bg=bg, highlightthickness=0, bd=0, **kwargs)
        self.size = size
        self.track_color = track_color
        self.prog_color = prog_color
        self.count = 0
        self.total = 5
        self.draw_ring()

    def set_progress(self, count, total=5):
        self.count = count
        self.total = total
        self.draw_ring()

    def draw_ring(self):
        self.delete("all")
        pad = 2.5
        w = 3.0
        self.create_oval(pad, pad, self.size - pad, self.size - pad, outline=self.track_color, width=w)
        if self.count > 0:
            extent = -(self.count / max(1, self.total)) * 359.9
            self.create_arc(pad, pad, self.size - pad, self.size - pad, start=90, extent=extent, outline=self.prog_color, width=w, style="arc")
        self.create_text(self.size / 2, self.size / 2, text=f"{self.count}/{self.total}", fill="#FFFFFF", font=("Malgun Gothic", 7, "bold"))


class ModernRoundedButton(tk.Canvas):
    """안티앨리어싱 둥근 모서리 버튼 (고정 높이 Canvas 기반)"""
    def __init__(self, parent, text, command=None, height=36, radius=8, bg_color="#1F9D63", hover_bg="#178350", fg_color="#FFFFFF", border_color=None, font=("Malgun Gothic", 10, "bold"), outer_bg="#EEF1F6", **kwargs):
        super().__init__(parent, height=height, bg=outer_bg, highlightthickness=0, bd=0, cursor="hand2", **kwargs)
        self.text = text
        self.command = command
        self.radius = radius
        self.normal_bg = bg_color
        self.hover_bg = hover_bg
        self.fg_color = fg_color
        self.border_color = border_color
        self.btn_font = font
        self.outer_bg = outer_bg
        self.current_bg = bg_color
        self._last_w = 0
        self._last_h = height
        self._bg_img = None
        self.bind("<Configure>", self._on_resize)
        self.bind("<Enter>", lambda e: self._set_bg(self.hover_bg))
        self.bind("<Leave>", lambda e: self._set_bg(self.normal_bg))
        self.bind("<Button-1>", lambda e: self.command() if self.command else None)

    def _set_bg(self, color):
        self.current_bg = color
        self._redraw()

    def _on_resize(self, event):
        if event.width > 10 and event.height > 10 and (event.width != self._last_w or event.height != self._last_h):
            self._last_w = event.width
            self._last_h = event.height
            self._redraw()

    def _redraw(self):
        w = max(10, self._last_w or self.winfo_width())
        h = max(10, self._last_h or self.winfo_height())
        scale = 3
        W, H, R = int(w * scale), int(h * scale), int(self.radius * scale)
        img = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([0, 0, W - 1, H - 1], radius=R, fill=self.current_bg, outline=self.border_color, width=1 * scale if self.border_color else 0)
        self._bg_img = ImageTk.PhotoImage(img.resize((w, h), Image.Resampling.LANCZOS))
        self.delete("all")
        self.create_image(0, 0, image=self._bg_img, anchor="nw")
        self.create_text(w / 2, h / 2, text=self.text, fill=self.fg_color, font=self.btn_font)


class CleanMinimalDropZone(RoundedFrame):
    """
    VoucherPass Modern Rounded DropZone
    - 안티앨리어싱 둥근 테두리 카드 (RoundedFrame 상속)
    - 안티앨리어싱 둥근 문서 칩 (세, 명, PR, PO, 계)
    - 안티앨리어싱 둥근 대기/완료 배지
    """
    def __init__(self, parent, num_str, title, tag_str, chip_text, chip_bg, file_var, on_file_selected=None, on_state_changed=None, outer_bg="#EEF1F6", **kwargs):
        super().__init__(parent, radius=8, bg_color="#FFFFFF", border_color="#E1E5EC", border_width=1, outer_bg=outer_bg, padx=6, pady=4, **kwargs)
        self.file_var = file_var
        self.on_file_selected = on_file_selected
        self.on_state_changed = on_state_changed
        self.chip_bg = chip_bg
        self.num_str = num_str

        # 1. 문서 종류별 색상 아이콘 칩 (세, 명, PR, PO, 계)
        self.lbl_chip = RoundedBadge(
            self, text=chip_text, bg_color=chip_bg, fg_color="#FFFFFF",
            font=("Malgun Gothic", 9, "bold"), radius=6, width=28, height=22, outer_bg="#FFFFFF"
        )
        self.lbl_chip.pack(side="left", padx=(2, 6))

        # 2. 중앙 텍스트 컨테이너 (제목 + 태그 & 파일 힌트)
        self.txt_box = tk.Frame(self, bg="#FFFFFF")
        self.txt_box.pack(side="left", fill="both", expand=True)

        # 상단 타이틀 행 (순번 원 + 문서명 + 파일형식 태그)
        self.title_row = tk.Frame(self.txt_box, bg="#FFFFFF")
        self.title_row.pack(fill="x", anchor="w")

        self.lbl_num = RoundedBadge(
            self.title_row, text=num_str, bg_color="#E2E8F0", fg_color="#64748B",
            font=("Malgun Gothic", 7, "bold"), radius=8, width=16, height=16, outer_bg="#FFFFFF"
        )
        self.lbl_num.pack(side="left", padx=(0, 4))

        self.lbl_title = tk.Label(
            self.title_row, text=title, font=("Malgun Gothic", 9, "bold"),
            bg="#FFFFFF", fg="#1C2536", anchor="w"
        )
        self.lbl_title.pack(side="left")

        self.lbl_tag = tk.Label(
            self.title_row, text=f" {tag_str}", font=("Malgun Gothic", 8, "bold"),
            bg="#FFFFFF", fg="#64748B", anchor="w"
        )
        self.lbl_tag.pack(side="left", padx=(2, 0))

        # 하단 힌트/완료 안내
        self.lbl_status = tk.Label(
            self.txt_box, text="여기로 파일을 드래그하거나 클릭하여 선택",
            font=("Malgun Gothic", 7), bg="#FFFFFF", fg="#94A3B8", anchor="w"
        )
        self.lbl_status.pack(fill="x", pady=(1, 0))

        # 3. 우측 상태 배지 ([대기] / [완료])
        self.lbl_badge = RoundedBadge(
            self, text="대기", bg_color="#ECEFF4", fg_color="#6B7686",
            font=("Malgun Gothic", 7, "bold"), radius=5, width=36, height=18, outer_bg="#FFFFFF"
        )
        self.lbl_badge.pack(side="right", padx=(2, 2))

        for w in (self, self.lbl_chip, self.txt_box, self.title_row, self.lbl_num, self.lbl_title, self.lbl_tag, self.lbl_status, self.lbl_badge):
            w.config(cursor="hand2")
            w.bind("<Button-1>", self._browse_file)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._handle_drop)
            self.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            self.dnd_bind('<<DragLeave>>', self._on_drag_leave)

        self.file_var.trace_add("write", self._update_ui_state)

    def _browse_file(self, event=None):
        path = filedialog.askopenfilename(filetypes=[("문서", "*.pdf *.html *.htm")])
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
        self.set_colors(bg_color="#EAF0FD", border_color="#3457A8", border_width=2)
        self.txt_box.config(bg="#EAF0FD")
        self.title_row.config(bg="#EAF0FD")
        self.lbl_title.config(bg="#EAF0FD")
        self.lbl_tag.config(bg="#EAF0FD")
        self.lbl_status.config(bg="#EAF0FD")
        self.lbl_chip.outer_bg = "#EAF0FD"
        self.lbl_chip.redraw()
        self.lbl_num.outer_bg = "#EAF0FD"
        self.lbl_num.redraw()
        self.lbl_badge.outer_bg = "#EAF0FD"
        self.lbl_badge.redraw()

    def _on_drag_leave(self, event=None):
        self._update_ui_state()

    def set_file(self, path):
        allowed = ('.pdf', '.html', '.htm') if self.num_str == '1' else ('.pdf',)
        if not os.path.isfile(path) or not path.lower().endswith(allowed):
            messagebox.showwarning('파일 형식 확인', '이 영역에서 지원하는 문서 파일을 선택해 주세요.')
            return False
        self.file_var.set(path)
        if self.on_file_selected:
            self.on_file_selected(path)
        return True

    def _update_ui_state(self, *args):
        path = self.file_var.get()
        if path and os.path.exists(path):
            fname = os.path.basename(path)
            bg_color = "#E6F7EE"
            self.set_colors(bg_color=bg_color, border_color="#1F9D63", border_width=1)
            self.txt_box.config(bg=bg_color)
            self.title_row.config(bg=bg_color)
            self.lbl_title.config(bg=bg_color, fg="#1C2536")
            self.lbl_tag.config(bg=bg_color, fg="#1F9D63")
            self.lbl_status.config(text=f"{fname} 등록됨", fg="#1F9D63", font=("Malgun Gothic", 7, "bold"), bg=bg_color)
            self.lbl_chip.outer_bg = bg_color
            self.lbl_chip.redraw()
            self.lbl_num.set_content(bg_color="#1F9D63", fg_color="#FFFFFF")
            self.lbl_num.outer_bg = bg_color
            self.lbl_num.redraw()
            self.lbl_badge.set_content(text="등록", bg_color="#1F9D63", fg_color="#FFFFFF")
            self.lbl_badge.outer_bg = bg_color
            self.lbl_badge.redraw()
        else:
            bg_color = "#FFFFFF"
            self.set_colors(bg_color=bg_color, border_color="#E1E5EC", border_width=1)
            self.txt_box.config(bg=bg_color)
            self.title_row.config(bg=bg_color)
            self.lbl_title.config(bg=bg_color, fg="#1C2536")
            self.lbl_tag.config(bg=bg_color, fg="#64748B")
            self.lbl_status.config(text="여기로 파일을 드래그하거나 클릭하여 선택", fg="#94A3B8", font=("Malgun Gothic", 7), bg=bg_color)
            self.lbl_chip.outer_bg = bg_color
            self.lbl_chip.redraw()
            self.lbl_num.set_content(bg_color="#E2E8F0", fg_color="#64748B")
            self.lbl_num.outer_bg = bg_color
            self.lbl_num.redraw()
            self.lbl_badge.set_content(text="대기", bg_color="#ECEFF4", fg_color="#6B7686")
            self.lbl_badge.outer_bg = bg_color
            self.lbl_badge.redraw()

        if self.on_state_changed:
            self.on_state_changed()


class VoucherPassApp:
    """
    VoucherPass v8.8.0 document collection and validation
    """
    def __init__(self, root):
        self.root = root
        self.tasks = TaskRunner(root)
        self.documents = DocumentState()
        self._last_discrepancy = None
        self._closed = False
        self.root.title("Voucher Pass")
        self.root.geometry("440x720+1060+30")
        self.root.minsize(440, 720)
        self.root.maxsize(440, 720)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#EEF1F6", highlightbackground="#3457A8", highlightthickness=2)

        icon_path = self._get_icon_file('VoucherPass.ico')
        icon_png = self._get_icon_file('voucherpass_icon.png') or self._get_icon_file('app_icon.png')
        if icon_png and os.path.exists(icon_png):
            try:
                self._app_icon_img = ImageTk.PhotoImage(Image.open(icon_png))
                self.root.iconphoto(True, self._app_icon_img)
            except Exception:
                pass
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

        self.template_path = tk.StringVar()
        self.selected_printer = tk.StringVar()

        self.pr_no_var = tk.StringVar()
        self.pr_title_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.vat_var = tk.StringVar()
        self.total_amount_var = tk.StringVar()
        self.date_var = tk.StringVar()
        self.supplier_var = tk.StringVar()

        # PR 및 세금계산서 공급가액 불일치 검증용 상태 변수
        self.pr_supply_amount = None
        self.tax_supply_amount = None

        self.auto_watch_enabled = tk.BooleanVar(value=False)
        self.folder_watcher = pdf_watcher.DownloadFolderWatcher(self.handle_auto_detected_pdf)

        self._build_widget_layout()
        self._load_printers()
        self._start_folder_watch_timer()

        # 시스템 트레이 및 글로벌 단축키 초기화 (시작프로그램/바로가기는 트레이 메뉴 토글에 따름)
        self._setup_system_tray()
        self._setup_global_hotkey()

        # F4: 아웃룩 첨부파일 1클릭 가져오기, Ctrl+V: 클립보드 첨부파일 붙여넣기
        self.root.bind("<F6>", lambda e: self.generate_excel_action())
        self.root.bind("<F7>", lambda e: self.save_sliced_contract_pdf())
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)
        self.root.bind("<F4>", lambda e: self.fetch_outlook_action())
        self.root.bind("<Control-v>", lambda e: self.paste_from_clipboard_action())

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

            w, h = 440, 720
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
            command = f'"{exe_path}"' if os.path.exists(exe_path) else f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "VoucherPass", 0, winreg.REG_SZ, command)
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
            os.path.join(base_dir, 'assets', filename),
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
                    self.tasks.call_soon(self.popup_at_cursor)

                def _toggle_startup(icon, item):
                    if self.is_startup_registered():
                        self.tasks.call_soon(self.unregister_startup_registry)
                    else:
                        self.tasks.call_soon(self.register_startup_registry)

                def _on_exit(icon, item):
                    icon.stop()
                    self.tasks.call_soon(self.close_app)

                menu = pystray.Menu(
                    pystray.MenuItem("⚡ 위젯 열기/숨기기 (Ctrl+Shift+V / 마우스 위치)", _on_toggle, default=True),
                    pystray.MenuItem("🚀 시작프로그램 자동등록", _toggle_startup, checked=lambda item: self.is_startup_registered()),
                    pystray.MenuItem("🖥️ 바탕화면 바로가기 생성", lambda icon, item: self.tasks.call_soon(self.toggle_desktop_shortcut), checked=lambda item: self.is_desktop_shortcut_existing()),
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
                    self.tasks.call_soon(self.popup_at_cursor)

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
        try:
            page_indices, label_str = self._parse_contract_pages(page_str)
            from pypdf import PdfReader, PdfWriter
            reader = PdfReader(contract_pdf)
            writer = PdfWriter()
            total_pages = len(reader.pages)

            valid_pages = validate_pages(page_indices, total_pages)
            if not valid_pages:
                self.set_live_status(f"⚠️ 유효하지 않은 페이지 범위입니다. (총 {total_pages}페이지)", type="error")
                return

            for p in valid_pages:
                writer.add_page(reader.pages[p])

            desktop = str(desktop_dir())
            out_name = f"계약서_{label_str}p.pdf"
            import tempfile
            fd, out_path = tempfile.mkstemp(prefix=safe_name(f'계약서_{label_str}p')[:60] + '_', suffix='.pdf', dir=desktop)
            os.close(fd)

            with open(out_path, 'wb') as f_out:
                writer.write(f_out)

            self.set_live_status(f"🎉 업체 계약서 [{label_str}p] ({len(valid_pages)}페이지) 추출 저장 완료! 📂 {os.path.basename(out_path)}", type="success")
        except Exception as e:
            self.set_live_status(f"⚠️ 계약서 페이지 추출 오류: {e}", type="error")

    def is_desktop_shortcut_existing(self):
        """바탕화면에 VoucherPass.lnk 바로가기 존재 여부 확인"""
        try:
            desktop = str(desktop_dir())
            lnk_path = os.path.join(desktop, 'VoucherPass.lnk')
            return os.path.exists(lnk_path)
        except Exception:
            return False

    def delete_desktop_shortcut(self):
        """바탕화면 바로가기 삭제"""
        try:
            desktop = str(desktop_dir())
            lnk_path = os.path.join(desktop, 'VoucherPass.lnk')
            if os.path.exists(lnk_path):
                os.remove(lnk_path)
                self.set_live_status("🗑️ 바탕화면 바로가기가 삭제되었습니다.", type="info")
        except Exception as e:
            print(f"Delete shortcut error: {e}")

    def toggle_desktop_shortcut(self):
        """바탕화면 바로가기 생성/삭제 토글"""
        if self.is_desktop_shortcut_existing():
            self.delete_desktop_shortcut()
        else:
            self.create_desktop_shortcut()

    def create_desktop_shortcut(self):
        """
        Windows 정식 바로가기 (.lnk) 생성
        """
        try:
            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dist', 'VoucherPass.exe')
            desktop = str(desktop_dir())
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
            vbs_file = temp_output('create_lnk.vbs')
            with open(vbs_file, 'w', encoding='utf-16') as f:
                f.write(vbs_script)

            subprocess.run(['cscript', '//Nologo', vbs_file], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
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
        업로드 진행 상태 실시간 집계 & 뱃지(0/5) 업데이트
        """
        count = sum(1 for var in [self.tax_pdf_path, self.spec_pdf_path, self.pr_pdf_path, self.po_pdf_path, self.contract_pdf_path] if var.get() and os.path.exists(var.get()))
        if hasattr(self, 'circ_progress'):
            self.circ_progress.set_progress(count, 5)

    def _build_widget_layout(self):
        # Windows 11 DWM 둥근 모서리 (Apple-like Rounded Window Corners)
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            if not hwnd:
                hwnd = self.root.winfo_id()
            DWMWA_WINDOW_CORNER_PREFERENCE = 33
            DWMWCP_ROUND = 2
            val = ctypes.c_int(DWMWCP_ROUND)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ctypes.byref(val), ctypes.sizeof(val))
        except Exception:
            pass

        # 1. Header Bar (Navy/Blue Header & Progress Badge)
        hdr = tk.Frame(self.root, bg="#3457A8", padx=12, pady=8)
        hdr.pack(fill="x")
        hdr.bind("<Button-1>", self._click_title)
        hdr.bind("<B1-Motion>", self._drag_title)

        lbl_logo = tk.Label(hdr, text="⚡ Voucher Pass", font=("Malgun Gothic", 11, "bold"), bg="#3457A8", fg="#FFFFFF")
        lbl_logo.pack(side="left")
        lbl_logo.bind("<Button-1>", self._click_title)
        lbl_logo.bind("<B1-Motion>", self._drag_title)

        ver_b = RoundedBadge(hdr, text="v8.8.0", bg_color="#26407F", fg_color="#FFFFFF", font=("Malgun Gothic", 8, "bold"), radius=6, width=46, height=20, outer_bg="#3457A8")
        ver_b.pack(side="left", padx=(6, 0))

        btn_outlook = RoundedBadge(hdr, text="📬 아웃룩", bg_color="#4F46E5", fg_color="#FFFFFF", font=("Malgun Gothic", 8, "bold"), radius=6, width=62, height=20, outer_bg="#3457A8", cursor="hand2")
        btn_outlook.pack(side="left", padx=(6, 0))
        btn_outlook.bind("<Button-1>", lambda e: self.fetch_outlook_action())
        
        btn_reset = RoundedBadge(hdr, text="🔄 리셋", bg_color="#E02424", fg_color="#FFFFFF", font=("Malgun Gothic", 8, "bold"), radius=6, width=54, height=20, outer_bg="#3457A8", cursor="hand2")
        btn_reset.pack(side="left", padx=(6, 0))
        btn_reset.bind("<Button-1>", lambda e: self.reset_all_data())

        # 우측 창 제어 미니 버튼 및 도넛 프로그레스 링
        btn_close = tk.Label(hdr, text="●", font=("Arial", 10), bg="#3457A8", fg="#FCA5A5", cursor="hand2")
        btn_close.pack(side="right", padx=(4, 0))
        btn_close.bind("<Button-1>", lambda e: self.close_app())

        btn_min = tk.Label(hdr, text="●", font=("Arial", 10), bg="#3457A8", fg="#FDE047", cursor="hand2")
        btn_min.pack(side="right", padx=(4, 0))
        btn_min.bind("<Button-1>", lambda e: self.root.withdraw())

        self.circ_progress = CircularProgressRing(hdr, size=28, bg="#3457A8", track_color="#456BBF", prog_color="#22C55E")
        self.circ_progress.pack(side="right", padx=(0, 6))

        # 2. Main Content Canvas
        main_box = tk.Frame(self.root, bg="#EEF1F6", padx=10, pady=8)
        main_box.pack(fill="both", expand=True)

        # 3. 5가지 서류 Clean DropZone
        drop_container = tk.Frame(main_box, bg="#EEF1F6")
        drop_container.pack(fill="x", pady=(0, 2))

        self.drop_tax = CleanMinimalDropZone(
            drop_container, num_str="1", title="전자 세금계산서", tag_str="· 6068625399",
            chip_text="세", chip_bg="#C94A3C", file_var=self.tax_pdf_path,
            on_file_selected=self.parse_tax_invoice_uploaded, on_state_changed=self._update_progress_summary
        )
        self.drop_tax.pack(fill="x", pady=1.5)

        self.drop_spec = CleanMinimalDropZone(
            drop_container, num_str="2", title="거래명세서", tag_str="· PDF",
            chip_text="명", chip_bg="#3A6FD8", file_var=self.spec_pdf_path,
            on_state_changed=self._update_progress_summary
        )
        self.drop_spec.pack(fill="x", pady=1.5)

        self.drop_pr = CleanMinimalDropZone(
            drop_container, num_str="3", title="PR Print (구매요청서)", tag_str="· PDF",
            chip_text="PR", chip_bg="#8B5FD8", file_var=self.pr_pdf_path,
            on_file_selected=lambda path: self.parse_uploaded_pdf(path, 'pr'), on_state_changed=self._update_progress_summary
        )
        self.drop_pr.pack(fill="x", pady=1.5)

        self.drop_po = CleanMinimalDropZone(
            drop_container, num_str="4", title="발주서 (PO)", tag_str="· PDF",
            chip_text="PO", chip_bg="#E08C2B", file_var=self.po_pdf_path,
            on_file_selected=lambda path: self.parse_uploaded_pdf(path, 'po'), on_state_changed=self._update_progress_summary
        )
        self.drop_po.pack(fill="x", pady=1.5)

        # 5번 업체 계약서 (오른쪽에 페이지 입력란 일체형 통합)
        f_contract_row = tk.Frame(drop_container, bg="#EEF1F6")
        f_contract_row.pack(fill="x", pady=1.5)

        self.drop_contract = CleanMinimalDropZone(
            f_contract_row, num_str="5", title="업체 계약서", tag_str="· PDF",
            chip_text="계", chip_bg="#5A6472", file_var=self.contract_pdf_path,
            on_state_changed=self._update_progress_summary
        )
        self.drop_contract.pack(side="left", fill="x", expand=True)

        pg_box = RoundedFrame(f_contract_row, radius=6, bg_color="#FFFFFF", border_color="#E1E5EC", outer_bg="#EEF1F6", padx=5, pady=3)
        pg_box.pack(side="right", fill="y", padx=(4, 0))
        tk.Label(pg_box, text="페이지", font=("Malgun Gothic", 8, "bold"), bg="#FFFFFF", fg="#6B7686").pack(side="left", padx=(0, 3))
        self.e_contract_page = tk.Entry(pg_box, textvariable=self.contract_page, font=("Malgun Gothic", 9, "bold"), bg="#F8FAFC", fg="#1C2536", width=3, justify="center", relief="flat", bd=0)
        self.e_contract_page.pack(side="left")

        # 4. 추출 데이터 7종 섹션 (라벨-값 분리 & 금액 우측정렬 & 부가세 옆에 합계금액 배치로 세로 압축)
        data_card = RoundedFrame(main_box, radius=10, bg_color="#FFFFFF", border_color="#E1E5EC", outer_bg="#EEF1F6", padx=8, pady=5)
        data_card.pack(fill="x", pady=(2, 3))

        # 데이터 카드 헤더
        data_hdr = tk.Frame(data_card, bg="#FFFFFF")
        data_hdr.pack(fill="x", pady=(0, 2))
        tk.Label(data_hdr, text="📋 추출 데이터 7종", font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#1C2536").pack(side="left")

        # 그리드 컨테이너
        grid_f = tk.Frame(data_card, bg="#FFFFFF")
        grid_f.pack(fill="x")

        lbl_s = {"font": ("Malgun Gothic", 8, "bold"), "bg": "#FFFFFF", "fg": "#6B7686"}

        # Row 0: P/R No. (좌) | 작성일자 (우)
        f_c1_r0 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c1_r0.grid(row=0, column=0, sticky="ew", padx=3, pady=1.5)
        tk.Label(f_c1_r0, text="• P/R No.", **lbl_s).pack(anchor="w")
        box_prno = RoundedFrame(f_c1_r0, radius=5, bg_color="#FFFFFF", border_color="#CBD5E1", outer_bg="#FFFFFF", padx=5, pady=2)
        box_prno.pack(fill="x", pady=(1, 0))
        self.e_prno = tk.Entry(box_prno, textvariable=self.pr_no_var, font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#1C2536", relief="flat", bd=0)
        self.e_prno.pack(fill="x")

        f_c2_r0 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c2_r0.grid(row=0, column=1, sticky="ew", padx=3, pady=1.5)
        f_date_lbl_box = tk.Frame(f_c2_r0, bg="#FFFFFF")
        f_date_lbl_box.pack(fill="x")
        tk.Label(f_date_lbl_box, text="• 작성일자", **lbl_s).pack(side="left")
        btn_cp_date = RoundedBadge(f_date_lbl_box, text="📋", bg_color="#EAF0FD", fg_color="#3457A8", font=("Malgun Gothic", 7, "bold"), radius=4, width=20, height=16, outer_bg="#FFFFFF", cursor="hand2")
        btn_cp_date.pack(side="right")
        btn_cp_date.bind("<Button-1>", lambda e: self.copy_to_clipboard(self.date_var.get(), "작성일자"))
        box_date = RoundedFrame(f_c2_r0, radius=5, bg_color="#F5F8FF", border_color="#CBD5E1", outer_bg="#FFFFFF", padx=5, pady=2)
        box_date.pack(fill="x", pady=(1, 0))
        self.e_date = tk.Entry(box_date, textvariable=self.date_var, font=("Malgun Gothic", 9, "bold"), bg="#F5F8FF", fg="#26407F", relief="flat", bd=0)
        self.e_date.pack(fill="x")

        # Row 1: PR Title (전체 2열 너비 차지)
        f_r1 = tk.Frame(grid_f, bg="#FFFFFF")
        f_r1.grid(row=1, column=0, columnspan=2, sticky="ew", padx=3, pady=1.5)
        f_t_lbl_box = tk.Frame(f_r1, bg="#FFFFFF")
        f_t_lbl_box.pack(fill="x")
        tk.Label(f_t_lbl_box, text="• PR Title", **lbl_s).pack(side="left")
        btn_cp_t = RoundedBadge(f_t_lbl_box, text="📋", bg_color="#EAF0FD", fg_color="#3457A8", font=("Malgun Gothic", 7, "bold"), radius=4, width=20, height=16, outer_bg="#FFFFFF", cursor="hand2")
        btn_cp_t.pack(side="right")
        btn_cp_t.bind("<Button-1>", lambda e: self.copy_to_clipboard(self.pr_title_var.get(), "PR Title"))
        box_t = RoundedFrame(f_r1, radius=5, bg_color="#F5F8FF", border_color="#CBD5E1", outer_bg="#FFFFFF", padx=5, pady=2)
        box_t.pack(fill="x", pady=(1, 0))
        self.txt_pr_title = tk.Text(box_t, width=20, height=2, wrap="word", font=("Malgun Gothic", 9, "bold"), bg="#F5F8FF", fg="#26407F", insertbackground="#26407F", relief="flat", bd=0)
        self.txt_pr_title.pack(fill="x")
        self.txt_pr_title.bind("<KeyRelease>", self._on_pr_title_txt_changed)

        # Row 2: 공급가액 (좌) | 거래처명 (우)
        f_c1_r2 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c1_r2.grid(row=2, column=0, sticky="ew", padx=3, pady=1.5)
        f_amt_lbl_box = tk.Frame(f_c1_r2, bg="#FFFFFF")
        f_amt_lbl_box.pack(fill="x")
        tk.Label(f_amt_lbl_box, text="• 공급가액", **lbl_s).pack(side="left")
        btn_cp_amt = RoundedBadge(f_amt_lbl_box, text="📋", bg_color="#EAF0FD", fg_color="#3457A8", font=("Malgun Gothic", 7, "bold"), radius=4, width=20, height=16, outer_bg="#FFFFFF", cursor="hand2")
        btn_cp_amt.pack(side="right")
        btn_cp_amt.bind("<Button-1>", lambda e: self.copy_to_clipboard(self.amount_var.get(), "공급가액"))
        box_amt = RoundedFrame(f_c1_r2, radius=5, bg_color="#F5F8FF", border_color="#CBD5E1", outer_bg="#FFFFFF", padx=5, pady=2)
        box_amt.pack(fill="x", pady=(1, 0))
        self.e_amt = tk.Entry(box_amt, textvariable=self.amount_var, font=("Consolas", 10, "bold"), bg="#F5F8FF", fg="#26407F", justify="right", relief="flat", bd=0)
        self.e_amt.pack(fill="x")
        self.e_amt.bind("<KeyRelease>", self._recalc_amounts)

        f_c2_r2 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c2_r2.grid(row=2, column=1, sticky="ew", padx=3, pady=1.5)
        tk.Label(f_c2_r2, text="• 거래처명", **lbl_s).pack(anchor="w")
        box_sup = RoundedFrame(f_c2_r2, radius=5, bg_color="#FFFFFF", border_color="#CBD5E1", outer_bg="#FFFFFF", padx=5, pady=2)
        box_sup.pack(fill="x", pady=(1, 0))
        self.e_sup = tk.Entry(box_sup, textvariable=self.supplier_var, font=("Malgun Gothic", 9, "bold"), bg="#FFFFFF", fg="#1C2536", relief="flat", bd=0)
        self.e_sup.pack(fill="x")

        # Row 3: 부가세 (좌) | 합계금액 (우측 골드 카드 - 세로 공간 최적화)
        f_c1_r3 = tk.Frame(grid_f, bg="#FFFFFF")
        f_c1_r3.grid(row=3, column=0, sticky="ew", padx=3, pady=1.5)
        tk.Label(f_c1_r3, text="• 부가세", **lbl_s).pack(anchor="w")
        box_vat = RoundedFrame(f_c1_r3, radius=5, bg_color="#F5F8FF", border_color="#CBD5E1", outer_bg="#FFFFFF", padx=5, pady=2)
        box_vat.pack(fill="x", pady=(1, 0))
        self.e_vat = tk.Entry(box_vat, textvariable=self.vat_var, font=("Consolas", 10, "bold"), bg="#F5F8FF", fg="#26407F", justify="right", relief="flat", bd=0)
        self.e_vat.pack(fill="x")
        self.e_vat.bind("<KeyRelease>", self._recalc_amounts)

        # 합계금액 (우측 1열 차지: 골드 라운드 카드)
        f_c2_r3 = RoundedFrame(grid_f, radius=8, bg_color="#FFF9EC", border_color="#F0A531", outer_bg="#FFFFFF", padx=6, pady=2)
        f_c2_r3.grid(row=3, column=1, sticky="ew", padx=3, pady=1.5)

        f_tot_lbl_box = tk.Frame(f_c2_r3, bg="#FFF9EC")
        f_tot_lbl_box.pack(fill="x")
        tk.Label(f_tot_lbl_box, text="• 합계금액", font=("Malgun Gothic", 8, "bold"), bg="#FFF9EC", fg="#8A5C00").pack(side="left")
        
        f_tot_val_box = tk.Frame(f_c2_r3, bg="#FFF9EC")
        f_tot_val_box.pack(fill="x", pady=(1, 0))
        tk.Label(f_tot_val_box, text="₩ ", font=("Malgun Gothic", 9, "bold"), bg="#FFF9EC", fg="#8A5C00").pack(side="left")
        self.e_tot = tk.Entry(f_tot_val_box, textvariable=self.total_amount_var, font=("Consolas", 10, "bold"), bg="#FFF9EC", fg="#8A5C00", justify="right", relief="flat", bd=0)
        self.e_tot.pack(side="right", fill="x", expand=True)

        grid_f.columnconfigure(0, weight=1)
        grid_f.columnconfigure(1, weight=1)

        # 실시간 상태 바 (Toast)
        toast_box = RoundedFrame(main_box, radius=8, bg_color="#E6F7EE", border_color="#86EFAC", outer_bg="#EEF1F6", padx=8, pady=3)
        toast_box.pack(fill="x", pady=(2, 3))

        self.lbl_live_status = tk.Label(
            toast_box, text="서류를 등록해 주세요. F4: 아웃룩 가져오기",
            font=("Malgun Gothic", 8, "bold"), bg="#E6F7EE", fg="#1F9D63", anchor="w"
        )
        self.lbl_live_status.pack(fill="x")

        # 5. 버튼 계층화 (Actions: '서류 5종 일괄 인쇄'를 상단 풀사이즈 주 버튼으로, 나머지 2개는 보조 버튼으로 하단 배치)
        act_panel = tk.Frame(main_box, bg="#EEF1F6")
        act_panel.pack(fill="x", pady=(0, 2))

        # 주 버튼 (Primary: 상단 풀사이즈 에메랄드 그린 ModernRoundedButton)
        btn_print = ModernRoundedButton(
            act_panel, text="🖨️ 서류 5종 일괄 인쇄", command=self.print_pdf_documents_only,
            height=38, radius=8, bg_color="#1F9D63", hover_bg="#178350", fg_color="#FFFFFF",
            font=("Malgun Gothic", 10, "bold")
        )
        btn_print.pack(side="top", fill="x", pady=(0, 3))

        # 보조 버튼 2종 (Secondary: 하단 2분할 가로 배치 ModernRoundedButton)
        btn_row = tk.Frame(act_panel, bg="#EEF1F6")
        btn_row.pack(fill="x")

        btn_copy_all = ModernRoundedButton(
            btn_row, text="📊 엑셀 양식 붙여넣기", command=self.copy_all_3items,
            height=32, radius=7, bg_color="#FFFFFF", hover_bg="#EAF0FD", fg_color="#3457A8",
            border_color="#3457A8", font=("Malgun Gothic", 8, "bold")
        )
        btn_copy_all.pack(side="left", fill="x", expand=True, padx=(0, 2))

        btn_arch = ModernRoundedButton(
            btn_row, text="📁 건별 자동 보관", command=self.archive_voucher_files,
            height=32, radius=7, bg_color="#FFFFFF", hover_bg="#EAF0FD", fg_color="#3457A8",
            border_color="#3457A8", font=("Malgun Gothic", 8, "bold")
        )
        btn_arch.pack(side="right", fill="x", expand=True, padx=(2, 0))

        # 6. Printer Selector Bar
        prt_bar = tk.Frame(main_box, bg="#EEF1F6")
        prt_bar.pack(fill="x", pady=(2, 0))
        tk.Label(prt_bar, text="🖨️ 프린터:", font=("Malgun Gothic", 8, "bold"), bg="#EEF1F6", fg="#6B7686").pack(side="left")
        self.printer_combo = ttk.Combobox(prt_bar, textvariable=self.selected_printer, font=("Malgun Gothic", 8), state="readonly")
        self.printer_combo.pack(side="left", fill="x", expand=True, padx=(4, 0))

    def _start_folder_watch_timer(self):
        if self.auto_watch_enabled.get():
            try:
                self.folder_watcher.check_new_files()
            except Exception:
                pass
        self.root.after(300, self._start_folder_watch_timer)

    def handle_auto_detected_pdf(self, pdf_path, pdf_type=None):
        if not pdf_path or not os.path.isfile(pdf_path):
            return False
        if pdf_type is None:
            self.tasks.submit('classify:' + pdf_path,
                lambda: pdf_parser.classify_pdf_type(pdf_path),
                lambda kind: self.handle_auto_detected_pdf(pdf_path, kind),
                lambda error: self.set_live_status(error, type='error'))
            return False
        if pdf_type not in ('tax', 'pr', 'po', 'spec', 'contract'):
            self.set_live_status(f'⚠️ 미분류: {os.path.basename(pdf_path)} — 문서 칸에 직접 등록해 주세요.', type='error')
            return False
        var = getattr(self, pdf_type + '_pdf_path')
        if var.get():
            self.set_live_status(f'⚠️ {pdf_type.upper()} 칸이 사용 중입니다. 새 파일은 직접 선택해 주세요.', type='error')
            return False
        return getattr(self, 'drop_' + pdf_type).set_file(pdf_path)

    def reset_all_data(self):
        """업로드된 모든 자료와 입력된 데이터(텍스트박스)를 초기화"""
        if self.tasks.busy & {'print', 'archive', 'excel'}:
            self.set_live_status('인쇄·저장 처리 중입니다. 완료 후 초기화해 주세요.', type='info')
            return
        self.tasks.invalidate()
        self.documents.clear()
        self._last_discrepancy = None
        self.pr_pdf_path.set("")
        self.po_pdf_path.set("")
        self.spec_pdf_path.set("")
        self.tax_pdf_path.set("")
        self.contract_pdf_path.set("")
        self.contract_page.set("1")

        self.pr_no_var.set("")
        self.pr_title_var.set("")
        self.amount_var.set("")
        self.vat_var.set("")
        self.total_amount_var.set("")
        self.date_var.set("")
        self.supplier_var.set("")
        
        self.pr_supply_amount = None
        self.tax_supply_amount = None

        
        if hasattr(self, 'txt_pr_title'):
            self.txt_pr_title.delete("1.0", "end")
            
        self._update_progress_summary()
        self.set_live_status("🔄 모든 데이터가 초기화되었습니다.", type="info")

    def fetch_outlook_action(self):
        def work():
            files, message = outlook_handler.fetch_outlook_attachments()
            return [(path, pdf_parser.classify_pdf_type(path)) for path in files], message
        if self.tasks.submit('outlook', work,
                lambda result: self._on_fetch_outlook_done(*result),
                lambda error: self.set_live_status(f'아웃룩 오류: {error}', type='error')):
            self.set_live_status('⏳ 아웃룩 첨부파일 수집·분류 중...', type='info')

    def _on_fetch_outlook_done(self, files, msg):
        if not files:
            self.set_live_status('⚠️ 아웃룩 연동 실패', type='error')
            messagebox.showinfo('아웃룩 첨부파일 가져오기', msg)
            return
        rejected = []
        for path, kind in files:
            if not self.handle_auto_detected_pdf(path, kind):
                rejected.append(os.path.basename(path))
        if rejected:
            messagebox.showwarning('등록 확인 필요', '미분류 또는 이미 사용 중인 문서 칸: ' + ', '.join(rejected))

    def paste_from_clipboard_action(self):
        # Ordinary text paste must not start COM work or overwrite registered files.
        for path in outlook_handler.extract_files_from_clipboard():
            self.handle_auto_detected_pdf(path)

    def generate_excel_action(self):
        if self.tasks.busy & {'pr', 'po', 'tax', 'outlook'}:
            self.set_live_status('문서 분석 완료 후 저장해 주세요.', type='info')
            return
        try:
            data = self._get_form_data()
            amounts = [data[key] for key in ('amount', 'vat', 'total_amount')]
            if any(value is None for value in amounts) or amounts[0] + amounts[1] != amounts[2]:
                raise ValueError('공급가액·부가세·합계금액을 확인해 주세요.')
            if self.pr_supply_amount is not None and self.tax_supply_amount is not None and self.pr_supply_amount != self.tax_supply_amount:
                if not messagebox.askyesno('원본 금액 불일치', '원본 PR과 세금계산서의 금액이 다릅니다. 현재 입력값으로 저장하시겠습니까?'):
                    return
            template = self.template_path.get()
            if not template or not os.path.isfile(template):
                template = filedialog.askopenfilename(title='전표 Excel 템플릿 선택', filetypes=[('Excel', '*.xlsx')])
                if not template:
                    return
                self.template_path.set(template)
            self.tasks.submit('excel', lambda: excel_handler.generate_voucher_excel(data, template),
                lambda path: self.set_live_status(f'엑셀 저장 완료: {os.path.basename(path)}', type='success'),
                lambda error: self.set_live_status(f'엑셀 저장 실패: {error}', type='error'))
        except ValueError as error:
            self.set_live_status(str(error), type='error')

    def archive_voucher_files(self):
        if self.tasks.busy & {'pr', 'po', 'tax', 'outlook'}:
            self.set_live_status('문서 분석·변환 완료 후 보관해 주세요.', type='info')
            return
        try:
            data = self._get_form_data()
            paths = {kind: getattr(self, kind + '_pdf_path').get() for kind in ('tax', 'spec', 'pr', 'po', 'contract')}
            def done(result):
                target, files = result
                self.set_live_status(f'서류 {len(files)}개 보관 완료', type='success')
                if messagebox.askyesno('자동 보관 완료', f'{target}\n\n보관 폴더를 여시겠습니까?'):
                    os.startfile(target)
            self.tasks.submit('archive', lambda: excel_handler.archive_voucher_package(data, paths), done,
                lambda error: self.set_live_status(f'보관 실패: {error}', type='error'))
        except ValueError as error:
            self.set_live_status(str(error), type='error')

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

        if not all((d, t, a)):
            self.set_live_status('작성일자·제목·공급가액을 모두 확인해 주세요.', type='error')
            return
        try:
            parse_money(a)
        except ValueError as error:
            self.set_live_status(str(error), type='error')
            return
        t = ' '.join(t.splitlines())
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

    def parse_uploaded_pdf(self, pr_path=None, kind='pr'):
        path = pr_path or getattr(self, kind + '_pdf_path').get()
        if not path or not os.path.isfile(path):
            return
        self.documents.replace(kind, {})
        self._refresh_form()
        self.set_live_status(f'⏳ {kind.upper()} 분석 중...', type='info')
        self.tasks.submit(kind, lambda: pdf_parser.parse_pr_pdf(path),
            lambda data: self._apply_document(kind, data),
            lambda error: self.set_live_status(f'{kind.upper()} 분석 실패: {error}', type='error'), replace=True)

    def _refresh_form(self):
        data = self.documents.form()
        for key, value in data.items():
            var = getattr(self, key + '_var')
            var.set('' if value is None else f'{value:,}' if isinstance(value, int) else value)
        self.pr_supply_amount = self.documents.supply('pr')
        self.tax_supply_amount = self.documents.supply('tax')

    def _apply_document(self, kind, data):
        self.documents.replace(kind, data)
        self._refresh_form()
        missing = not data.get('date') if kind == 'tax' else not data.get('pr_no') or not data.get('pr_title')
        warnings = data.get('warnings', [])
        self.set_live_status(f'{kind.upper()} 분석: ' + ('확인 필요 — ' + '; '.join(warnings or ['일부 항목 미추출']) if warnings or missing else '추출 완료'),
                             type='error' if warnings or missing else 'success')
        self._check_amount_discrepancy()

    def parse_tax_invoice_uploaded(self, tax_path=None):
        path = tax_path or self.tax_pdf_path.get()
        if not path or not os.path.isfile(path):
            return
        self.documents.replace('tax', {})
        self._refresh_form()
        if pdf_parser._is_html_file(path):
            self.auto_unlock_and_save_pdf(path)
            return
        self.set_live_status('⏳ 세금계산서 분석 중...', type='info')
        self.tasks.submit('tax', lambda: pdf_parser.parse_tax_invoice_data(path), self._apply_tax_data,
            lambda error: self.set_live_status(f'세금계산서 분석 실패: {error}', type='error'), replace=True)

    def auto_unlock_and_save_pdf(self, html_path):
        self.set_live_status('⏳ 세금계산서 복호화·PDF 변환 중...', type='info')
        self.tasks.submit('tax', lambda: convert_tax_html(html_path),
            lambda result: self._on_new_tax_pdf_saved(*result),
            lambda error: self.set_live_status(f'PDF 변환 실패: {error}', type='error'), replace=True)

    def _on_new_tax_pdf_saved(self, pdf_path, tax_data):
        # Called by TaskRunner on the Tk main thread. Do not parse again.
        self.tax_pdf_path.set(pdf_path)
        self._apply_tax_data(tax_data)

    def _apply_tax_data(self, tax_data):
        self._apply_document('tax', tax_data)

    def _check_amount_discrepancy(self):
        pr_amt, tax_amt = self.pr_supply_amount, self.tax_supply_amount
        pair = (pr_amt, tax_amt)
        if pair == self._last_discrepancy:
            return
        self._last_discrepancy = pair
        if pr_amt is None or tax_amt is None:
            return
        if pr_amt != tax_amt:
            self.set_live_status(f'⚠️ 원본 공급가액 불일치: PR {pr_amt:,} / 세금 {tax_amt:,}', type='error')
            messagebox.showwarning('공급가액 불일치', f'원본 PR: {pr_amt:,}원 / 원본 세금계산서: {tax_amt:,}원\n입력값을 수정해도 원본 비교값은 유지됩니다.')

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

    def _recalc_amounts(self, event=None):
        try:
            amount = parse_money(self.amount_var.get())
            vat = parse_money(self.vat_var.get())
            self.total_amount_var.set(f'{amount + vat:,}' if amount is not None and vat is not None else '')
        except ValueError:
            self.total_amount_var.set('')

    def _parse_contract_pages(self, page_str):
        return parse_pages(page_str)

    def _get_form_data(self):
        data = {key: getattr(self, key + '_var').get().strip() for key in
                ('pr_no', 'pr_title', 'amount', 'vat', 'total_amount', 'date', 'supplier')}
        for key in ('amount', 'vat', 'total_amount'):
            data[key] = parse_money(data[key])
        return data

    def close_app(self):
        if self._closed:
            return
        self._closed = True
        self.tasks.close()
        if hasattr(self, 'tray_icon'):
            self.tray_icon.stop()
        self.root.destroy()

    def print_pdf_documents_only(self):
        if self.tasks.busy & {'pr', 'po', 'tax', 'outlook'}:
            self.set_live_status('문서 분석·변환이 끝난 후 인쇄해 주세요.', type='info')
            return
        try:
            paths = {kind: getattr(self, kind + '_pdf_path').get() for kind in ('tax', 'spec', 'pr', 'po', 'contract')}
            pages, label = self._parse_contract_pages(self.contract_page.get()) if paths['contract'] else (None, '')
            printer = self.selected_printer.get()
            def done(result):
                submitted, failed = result
                self.set_live_status('인쇄 요청: ' + ', '.join(submitted) + (' / 실패: ' + '; '.join(failed) if failed else ''),
                                     type='error' if failed else 'success')
                if failed:
                    messagebox.showwarning('인쇄 요청 확인', '\n'.join(failed))
            if self.tasks.submit('print', lambda: printer_handler.print_documents(paths, printer, pages), done,
                                 lambda error: self.set_live_status(f'인쇄 중단: {error}', type='error')):
                self.set_live_status('⏳ 서류 검증·인쇄 요청 중...', type='info')
        except ValueError as error:
            self.set_live_status(str(error), type='error')

def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = VoucherPassApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
