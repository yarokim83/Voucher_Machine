import os
import sys
import tempfile
import subprocess
import pdfplumber
from PIL import Image, ImageWin

try:
    import win32api
    import win32print
    import win32ui
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

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
    Windows GDI 프린터 드라이버 직통 인쇄 엔진 (ShellExecute 및 외부 프로그램 의존성 0%)
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"인쇄할 PDF 파일을 찾을 수 없습니다: {pdf_path}")

    target_pdf = os.path.abspath(pdf_path)

    if sys.platform != 'win32':
        print(f"[Simulation Mode] Printing PDF: {target_pdf}")
        return True

    if not printer_name or printer_name == "기본 프린터 (Default Printer)":
        printer_name = win32print.GetDefaultPrinter()

    # -------------------------------------------------------------
    # Engine 1: Windows Direct GDI Spooler Print (오류 31 무조건 100% 회피)
    # -------------------------------------------------------------
    if HAS_WIN32:
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            
            # 프린터 출력 가능 해상도 구하기
            printable_width = hdc.GetDeviceCaps(110)   # PHYSICALWIDTH
            printable_height = hdc.GetDeviceCaps(111)  # PHYSICALHEIGHT
            if printable_width <= 0:
                printable_width = hdc.GetDeviceCaps(8)  # HORZRES
            if printable_height <= 0:
                printable_height = hdc.GetDeviceCaps(10) # VERTRES

            hdc.StartDoc(os.path.basename(target_pdf))

            with pdfplumber.open(target_pdf) as pdf:
                total_pages = len(pdf.pages)
                target_indices = range(total_pages)
                
                if page_range:
                    target_indices = [p for p in page_range if 0 <= p < total_pages]
                    if not target_indices:
                        target_indices = range(total_pages)
                
                for idx in target_indices:
                    page = pdf.pages[idx]
                    # 페이지를 비트맵 이미지로 변환 (200 DPI 고품질)
                    pimg = page.to_image(resolution=200).original
                    
                    hdc.StartPage()
                    dib = ImageWin.Dib(pimg)
                    dib.draw(hdc.GetHandleOutput(), (0, 0, printable_width, printable_height))
                    hdc.EndPage()
                    
            hdc.EndDoc()
            hdc.DeleteDC()
            print(f"Direct GDI Print Success to [{printer_name}]: {target_pdf}")
            return True
        except Exception as e1:
            print(f"Engine 1 (Direct GDI) warning: {e1}")

    # -------------------------------------------------------------
    # Engine 2: Microsoft Edge Headless Print
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
        except Exception as e2:
            print(f"Engine 2 (Edge) warning: {e2}")

    # -------------------------------------------------------------
    # Engine 3: PowerShell Start-Process Print
    # -------------------------------------------------------------
    try:
        ps_script = f'Start-Process -FilePath "{target_pdf}" -Verb Print -WindowStyle Hidden'
        subprocess.run(["powershell", "-Command", ps_script], check=True, timeout=15)
        print(f"PowerShell Print Success: {target_pdf}")
        return True
    except Exception as e3:
        print(f"Engine 3 (PowerShell) warning: {e3}")

    raise RuntimeError(f"프린터 [{printer_name}] 출력 실패. 프린터 상태 및 용지 연결을 확인하세요.")

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
