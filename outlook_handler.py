import os
import glob
import time
import tempfile
import winreg
from file_utils import temp_output

try:
    import win32clipboard
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


def get_outlook_temp_dirs():
    dirs = []
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Office\16.0\Outlook\Security')
        val, _ = winreg.QueryValueEx(key, 'OutlookSecureTempFolder')
        if os.path.exists(val):
            dirs.append(val)
    except Exception:
        pass

    inet_outlook = os.path.join(os.getenv('LOCALAPPDATA', ''), r'Microsoft\Windows\INetCache\Content.Outlook')
    if os.path.exists(inet_outlook):
        for sub in os.listdir(inet_outlook):
            sp = os.path.join(inet_outlook, sub)
            if os.path.isdir(sp) and sp not in dirs:
                dirs.append(sp)
    return dirs


def extract_files_from_clipboard():
    if not HAS_WIN32:
        return []
    saved_files = []

    try:
        win32clipboard.OpenClipboard()
    except Exception:
        return []

    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP):
            file_paths = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
            for p in file_paths:
                if os.path.exists(p) and p.lower().endswith(('.html', '.htm', '.pdf')):
                    saved_files.append(p)
            if saved_files:
                return saved_files

        # Outlook 가상 첨부파일은 F4 COM 가져오기를 사용합니다.
    except Exception:
        pass
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass

    return saved_files


def extract_from_active_com_mail():
    if not HAS_WIN32:
        return []
    try:
        pythoncom.CoInitialize()
        ol = win32com.client.Dispatch('Outlook.Application')
        target_item = None
        insp = ol.ActiveInspector()
        if insp:
            target_item = insp.CurrentItem
        else:
            exp = ol.ActiveExplorer()
            if exp and exp.Selection and exp.Selection.Count > 0:
                target_item = exp.Selection.Item(1)

        if not target_item:
            return []

        atts = getattr(target_item, 'Attachments', None)
        if not atts or atts.Count == 0:
            return []

        results = []
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fn = att.FileName
            if fn.lower().endswith(('.html', '.htm', '.pdf')):
                fp = temp_output(fn)
                att.SaveAsFile(fp)
                results.append(fp)
        return results
    except Exception:
        return []
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def fetch_latest_outlook_file(max_age_seconds=1200):
    dirs = get_outlook_temp_dirs()
    found = []
    now = time.time()
    for d in dirs:
        for ext in ('*.html', '*.htm', '*.pdf'):
            for fp in glob.glob(os.path.join(d, ext)):
                mtime = os.path.getmtime(fp)
                found.append((mtime, fp))
    found.sort(key=lambda x: x[0], reverse=True)
    if found:
        newest_mtime, newest_fp = found[0]
        if now - newest_mtime <= max_age_seconds:
            return newest_fp
    return None


def fetch_outlook_attachments():
    """
    아웃룩에서 현재 선택/열람 중인 메일의 첨부파일을 COM API로 직접 추출.
    COM 실패 시 즉시 실패 메시지 반환 (임시 폴더/클립보드 폴백 없음).
    Returns: (list of file_paths, msg)
    """
    com_files = extract_from_active_com_mail()
    if com_files:
        return com_files, '아웃룩에서 선택한 메일의 첨부파일을 가져왔습니다.'

    return [], '아웃룩 연동 실패: 아웃룩에서 세금계산서 메일을 선택한 상태인지 확인해 주세요.\n\n• 아웃룩이 실행 중이어야 합니다\n• 메일 목록에서 해당 메일을 클릭/선택해 주세요\n• 첨부파일(NTS_eTaxInvoice.html)이 포함된 메일이어야 합니다'


def fetch_outlook_attachment():
    files, msg = fetch_outlook_attachments()
    if files:
        return files[0], msg
    return None, msg

