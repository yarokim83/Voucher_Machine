import os
import sys
import tempfile
import subprocess

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import win32api
    import win32print
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

def find_edge_path():
    """
    Windows Edge 설치 경로 탐색
    """
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None

def print_pdf_file(pdf_path, printer_name=None, page_range=None):
    """
    ShellExecute Error 31 방지 다중 안전 인쇄 엔진
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"인쇄할 PDF 파일을 찾을 수 없습니다: {pdf_path}")

    target_pdf = os.path.abspath(pdf_path)

    # 특정 페이지 범위 지정 시 임시 PDF 추출 생성
    if page_range and isinstance(page_range, (list, tuple)):
        try:
            import fitz
            doc = fitz.open(target_pdf)
            new_doc = fitz.open()
            for pno in page_range:
                if 0 <= pno < len(doc):
                    new_doc.insert_pdf(doc, from_page=pno, to_page=pno)
            
            temp_dir = tempfile.gettempdir()
            temp_pdf = os.path.join(temp_dir, f"temp_print_{os.path.basename(target_pdf)}")
            new_doc.save(temp_pdf)
            new_doc.close()
            doc.close()
            target_pdf = temp_pdf
        except Exception as e:
            print(f"Page range extraction error: {e}")

    if sys.platform != 'win32':
        print(f"[Simulation Mode] Printing PDF: {target_pdf}")
        return True

    # -------------------------------------------------------------
    # Method 1: Edge 무음 Headless 프린트 엔진 (ShellExecute Error 31 회피)
    # -------------------------------------------------------------
    edge_exe = find_edge_path()
    if edge_exe:
        try:
            cmd = [edge_exe, "--headless", "--print-to-printer"]
            if printer_name and printer_name != "기본 프린터 (Default Printer)":
                cmd.append(f"--printer-name={printer_name}")
            cmd.append(target_pdf)

            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=15)
            print(f"Edge Direct Print Success: {target_pdf}")
            return True
        except Exception as e1:
            print(f"Method 1 (Edge) Print warning: {e1}")

    # -------------------------------------------------------------
    # Method 2: PowerShell Start-Process Print 엔진
    # -------------------------------------------------------------
    try:
        ps_script = f'Start-Process -FilePath "{target_pdf}" -Verb Print -WindowStyle Hidden'
        subprocess.run(["powershell", "-Command", ps_script], check=True, timeout=15)
        print(f"PowerShell Print Success: {target_pdf}")
        return True
    except Exception as e2:
        print(f"Method 2 (PowerShell) Print warning: {e2}")

    # -------------------------------------------------------------
    # Method 3: win32api ShellExecute 인쇄 엔진 (최후 백업)
    # -------------------------------------------------------------
    if HAS_WIN32:
        try:
            if printer_name and printer_name != "기본 프린터 (Default Printer)":
                try:
                    win32print.SetDefaultPrinter(printer_name)
                except Exception:
                    pass
            win32api.ShellExecute(0, "print", target_pdf, None, ".", 0)
            return True
        except Exception as e3:
            raise RuntimeError(f"모든 인쇄 엔진 시도 실패. 윈도우 프린터 연결을 확인하세요:\n{e3}")

    return False

def print_excel_file(excel_path, printer_name=None):
    """
    엑셀 Voucher 양식 인쇄
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"인쇄할 엑셀 파일을 찾을 수 없습니다: {excel_path}")

    if HAS_WIN32 and sys.platform == 'win32':
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            ws = wb.ActiveSheet
            
            if printer_name and printer_name != "기본 프린터 (Default Printer)":
                try:
                    excel.ActivePrinter = printer_name
                except Exception:
                    pass
                
            ws.PrintOut()
            wb.Close(False)
            excel.Quit()
            return True
        except Exception as e:
            print(f"Excel Print error: {e}")
            win32api.ShellExecute(0, "print", excel_path, None, ".", 0)
            return True
    else:
        print(f"[Simulation Mode] Printing Excel: {excel_path}")
        return False

def get_installed_printers():
    """
    현재 시스템에 설치된 프린터 목록 반환
    """
    printers = []
    if HAS_WIN32 and sys.platform == 'win32':
        try:
            default_p = win32print.GetDefaultPrinter()
            for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
                printers.append(p[2])
            if default_p not in printers:
                printers.insert(0, default_p)
        except Exception as e:
            print(f"Printer list error: {e}")
    if not printers:
        printers = ["기본 프린터 (Default Printer)"]
    return printers

if __name__ == '__main__':
    print("printer_handler loaded")
    print("Printers:", get_installed_printers())
