"""Collision-free output locations; source filenames never determine directories."""
import os
import re
import tempfile
from pathlib import Path


def safe_name(value, fallback='document'):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(value)).strip(' .')[:100]
    if not value or value.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1, 10)), *(f'LPT{i}' for i in range(1, 10))}:
        value = fallback
    return value


def temp_output(filename):
    parent = Path(tempfile.gettempdir()) / 'VoucherPass'
    parent.mkdir(exist_ok=True)
    return str(Path(tempfile.mkdtemp(prefix='job_', dir=parent)) / safe_name(filename))


def desktop_dir():
    if os.name == 'nt':
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders') as key:
                return Path(os.path.expandvars(winreg.QueryValueEx(key, 'Desktop')[0]))
        except OSError:
            pass
    return Path.home() / 'Desktop'
