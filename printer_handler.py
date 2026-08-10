import os
import sys
import tempfile
import subprocess
import pdfplumber
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageWin

try:
    import win32api
    import win32print
    import win32ui
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

def find_sumatra_path():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    bin_sumatra = os.path.join(base_dir, "bin", "SumatraPDF.exe")
    if os.path.exists(bin_sumatra):
        return bin_sumatra
    
    candidates = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe"
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None

def find_edge_path():
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
    [지정 페이지 정밀 인쇄 + SumatraPDF Vector Native Engine]
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"인쇄할 PDF 파일을 찾을 수 없습니다: {pdf_path}")

    target_pdf = os.path.abspath(pdf_path)

    # 특정 페이지 범위 지정 시 pypdf로 해당 페이지만 정확히 분리 추출한 임시 PDF 생성 (예: 12-13)
    if page_range and isinstance(page_range, (list, tuple)):
        try:
            reader = PdfReader(target_pdf)
            total_pdf_pages = len(reader.pages)
            valid_pages = [p for p in page_range if 0 <= p < total_pdf_pages]
            
            if valid_pages:
                writer = PdfWriter()
                for pno in valid_pages:
                    writer.add_page(reader.pages[pno])
                
                temp_dir = tempfile.gettempdir()
                temp_pdf = os.path.join(temp_dir, f"temp_contract_page_{os.path.basename(target_pdf)}")
                with open(temp_pdf, "wb") as f_out:
                    writer.write(f_out)
                target_pdf = temp_pdf
        except Exception as e:
            print(f"Page range extraction error: {e}")

    if sys.platform != 'win32':
        print(f"[Simulation Mode] Printing PDF: {target_pdf}")
        return True

    if not printer_name or printer_name == "기본 프린터 (Default Printer)":
        printer_name = win32print.GetDefaultPrinter()

    # -------------------------------------------------------------
    # Engine 1: SumatraPDF Portable Vector Native Engine (최우선 1순위)
    # -------------------------------------------------------------
    sumatra_exe = find_sumatra_path()
    if sumatra_exe:
        try:
            cmd = [sumatra_exe, "-print-to", printer_name, "-print-settings", "fit", target_pdf]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=20)
            print(f"[Engine 1: Sumatra Vector Native Print] Success to [{printer_name}]: {target_pdf}")
            return True
        except Exception as e1:
            print(f"Engine 1 (Sumatra Vector) warning: {e1}")

    # -------------------------------------------------------------
    # Engine 2: Microsoft Edge Native Vector Spooler
    # -------------------------------------------------------------
    edge_exe = find_edge_path()
    if edge_exe:
        try:
            cmd = [edge_exe, "--no-pdf-header-footer", "--print-to-printer"]
            if printer_name and printer_name != "기본 프린터 (Default Printer)":
                cmd.append(f"--printer-name={printer_name}")
            cmd.append(target_pdf)

            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=20)
            print(f"[Engine 2: Edge Native Vector Print] Success to [{printer_name}]: {target_pdf}")
            return True
        except Exception as e2:
            print(f"Engine 2 (Edge Vector) warning: {e2}")

    # -------------------------------------------------------------
    # Engine 3: PowerShell Native Print Verb
    # -------------------------------------------------------------
    try:
        ps_script = f'Start-Process -FilePath "{target_pdf}" -Verb Print -WindowStyle Hidden'
        subprocess.run(["powershell", "-Command", ps_script], check=True, timeout=20)
        print(f"[Engine 3: PowerShell Vector Print] Success to [{printer_name}]: {target_pdf}")
        return True
    except Exception as e3:
        print(f"Engine 3 (PowerShell) warning: {e3}")

    raise RuntimeError(f"프린터 [{printer_name}] 출력 실패. 프린터 연결 및 용지 상태를 확인하세요.")

def print_excel_file(excel_path, printer_name=None):
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
    print("printer_handler loaded with pypdf range extraction!")
