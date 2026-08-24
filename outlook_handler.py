import os
import glob
import time
import struct
import tempfile
import winreg

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
    temp_dir = os.path.join(tempfile.gettempdir(), 'VoucherPass_Clipboard')
    os.makedirs(temp_dir, exist_ok=True)
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

        fgd_format = win32clipboard.RegisterClipboardFormat('FileGroupDescriptorW')
        fc_format = win32clipboard.RegisterClipboardFormat('FileContents')

        if win32clipboard.IsClipboardFormatAvailable(fgd_format) and win32clipboard.IsClipboardFormatAvailable(fc_format):
            fgd_data = win32clipboard.GetClipboardData(fgd_format)
            num_files = struct.unpack('I', fgd_data[0:4])[0]
            for i in range(num_files):
                offset = 4 + i * 592
                fn_bytes = fgd_data[offset + 76 : offset + 76 + 520]
                fn = fn_bytes.decode('utf-16le', errors='ignore').split(chr(0))[0]
                if not fn:
                    fn = f'outlook_attachment_{i+1}.html'
                fc_data = win32clipboard.GetClipboardData(fc_format)
                fp = os.path.join(temp_dir, fn)
                with open(fp, 'wb') as f:
                    f.write(fc_data if isinstance(fc_data, bytes) else fc_data.encode('utf-8'))
                saved_files.append(fp)
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

        temp_dir = os.path.join(tempfile.gettempdir(), 'VoucherPass_Outlook')
        os.makedirs(temp_dir, exist_ok=True)
        results = []
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fn = att.FileName
            if fn.lower().endswith(('.html', '.htm', '.pdf')):
                fp = os.path.join(temp_dir, fn)
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
        return newest_fp
    return None


def fetch_outlook_attachments(max_age_seconds=600):
    """
    아웃룩에서 세금계산서(HTML/PDF) 및 거래명세서(PDF) 첨부파일들을 스마트 분류하여 일괄 반환.
    max_age_seconds: 임시 폴더 스캔 시 이 시간(초) 이내에 수정된 파일만 인식 (기본 10분).
    Returns: list of file_paths, msg
    """
    # 1. COM 활성 메일 확인 (메일이 열려 있거나 선택된 경우 모든 첨부파일 추출)
    com_files = extract_from_active_com_mail()
    if com_files:
        return com_files, '열려 있는 아웃룩 메일에서 첨부파일을 가져왔습니다.'

    # 2. 클립보드 확인 (Ctrl+C 복사된 첨부파일들)
    clip_files = extract_files_from_clipboard()
    if clip_files:
        return clip_files, '클립보드에서 첨부파일을 가져왔습니다.'

    # 3. 아웃룩 임시 폴더에서 최근 열람된 첨부파일만 스마트 추출
    dirs = get_outlook_temp_dirs()
    now = time.time()
    recent_htmls = []
    recent_pdfs = []

    for d in dirs:
        for fp in glob.glob(os.path.join(d, '*.html')) + glob.glob(os.path.join(d, '*.htm')):
            mtime = os.path.getmtime(fp)
            if now - mtime <= max_age_seconds:
                recent_htmls.append((mtime, fp))
        for fp in glob.glob(os.path.join(d, '*.pdf')):
            mtime = os.path.getmtime(fp)
            if now - mtime <= max_age_seconds:
                recent_pdfs.append((mtime, fp))

    recent_htmls.sort(key=lambda x: x[0], reverse=True)
    recent_pdfs.sort(key=lambda x: x[0], reverse=True)

    results = []

    # 세금계산서 HTML 우선 (NTS_eTaxInvoice 최우선)
    if recent_htmls:
        nts = [fp for _, fp in recent_htmls if any(k in os.path.basename(fp).lower() for k in ['nts', 'tax', 'invoice', '세금'])]
        sel_html = nts[0] if nts else recent_htmls[0][1]
        results.append(sel_html)

    # 최근 열람 PDF 추가
    if recent_pdfs:
        results.append(recent_pdfs[0][1])

    if results:
        return results, '아웃룩에서 최근 열람한 첨부파일을 가져왔습니다.'

    return [], '아웃룩 첨부파일을 찾을 수 없습니다. 아웃룩에서 메일을 열어두거나 첨부파일을 클릭/열람한 후 다시 시도해 주세요.'


def fetch_outlook_attachment():
    files, msg = fetch_outlook_attachments()
    if files:
        return files[0], msg
    return None, msg

