import os
import sys
import shutil
import winreg
import subprocess
import tkinter as tk
from tkinter import messagebox
from file_utils import desktop_dir, temp_output

def install():
    try:
        target_dir = os.path.join(os.environ.get('PUBLIC', r'C:\Users\Public'), 'VoucherPass')
        os.makedirs(target_dir, exist_ok=True)

        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        source_exe = os.path.join(base_dir, 'VoucherPass.exe')
        target_exe = os.path.join(target_dir, 'VoucherPass.exe')

        if not os.path.isfile(source_exe):
            raise FileNotFoundError('설치 패키지에 VoucherPass.exe가 없습니다.')
        # Stage first, then atomically replace. A running/locked EXE fails clearly.
        staged = os.path.join(target_dir, 'VoucherPass.new.exe')
        shutil.copy2(source_exe, staged)
        os.replace(staged, target_exe)

        # Windows 정식 .lnk 바로가기 아이콘 생성
        desktop = str(desktop_dir())
        lnk_path = os.path.join(desktop, 'VoucherPass.lnk')

        vbs_script = 'Set WshShell = WScript.CreateObject("WScript.Shell")\n' \
                     f'Set shortcut = WshShell.CreateShortcut("{lnk_path}")\n' \
                     f'shortcut.TargetPath = "{target_exe}"\n' \
                     f'shortcut.WorkingDirectory = "{target_dir}"\n' \
                     'shortcut.Description = "VoucherPass Tax & PR Extractor"\n' \
                     'shortcut.WindowStyle = 1\n' \
                     'shortcut.Save\n'

        vbs_file = temp_output('create_lnk_setup.vbs')
        with open(vbs_file, 'w', encoding='utf-16') as f:
            f.write(vbs_script)

        subprocess.run(['cscript', '//Nologo', vbs_file], check=True, creationflags=subprocess.CREATE_NO_WINDOW)

        warnings = []
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "VoucherPass", 0, winreg.REG_SZ, f'"{target_exe}"')
        except Exception as e:
            warnings.append(f"시작프로그램 등록 실패: {e}")

        if os.path.exists(target_exe):
            os.startfile(target_exe)

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("VoucherPass 설치 완료", f"🎉 VoucherPass 프로그램이 성공적으로 설치되었습니다!\n\n📂 설치 위치: {target_dir}\n🖥️ 바탕화면에 정식 바로가기(VoucherPass.lnk)가 생성되었습니다.\n⌨️ 단축키 [Ctrl+Shift+V] 로 언제든 마우스 커서 위치에서 위젯을 호출할 수 있습니다.")
        if warnings:
            messagebox.showwarning("설치 후 확인", "\n".join(warnings))
        root.destroy()

    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("설치 오류", f"프로그램 설치 중 오류가 발생했습니다:\n{e}")
        root.destroy()

if __name__ == '__main__':
    install()
