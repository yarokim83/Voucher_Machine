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
    고해상도(300 DPI) + 비율 유지(Aspect Ratio Fit) + Safe Margin 잘림 방지 100% 인쇄 엔진
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
    # Engine 1: 300 DPI High-Res Direct GDI Print with Aspect Ratio & Safe Margin
    # -------------------------------------------------------------
    if HAS_WIN32:
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            
            # 물리적 프린터 해상도 및 여백 정보 구하기
            pw = hdc.GetDeviceCaps(110)   # PHYSICALWIDTH
            ph = hdc.GetDeviceCaps(111)  # PHYSICALHEIGHT
            off_x = hdc.GetDeviceCaps(112) # PHYSICALOFFSETX
            off_y = hdc.GetDeviceCaps(113) # PHYSICALOFFSETY
            res_x = hdc.GetDeviceCaps(88)  # LOGPIXELSX
            res_y = hdc.GetDeviceCaps(90)  # LOGPIXELSY

            if pw <= 0:
                pw = hdc.GetDeviceCaps(8)  # HORZRES
            if ph <= 0:
                ph = hdc.GetDeviceCaps(10) # VERTRES

            if res_x <= 0: res_x = 300
            if res_y <= 0: res_y = 300

            # Safe Margin (약 0.15인치 / 4mm 안전 여백)
            margin_x = int(res_x * 0.15)
            margin_y = int(res_y * 0.15)

            target_w = max(100, pw - (margin_x * 2) - (off_x * 2))
            target_h = max(100, ph - (margin_y * 2) - (off_y * 2))

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
                    # 300 DPI 초고해상도 렌더링 (텍스트 선명도 극대화)
                    pimg = page.to_image(resolution=300).original
                    img_w, img_h = pimg.size

                    # Aspect Ratio (비율 유지 스케일링)
                    scale = min(target_w / img_w, target_h / img_h)
                    final_w = int(img_w * scale)
                    final_h = int(img_h * scale)

                    # 중앙 정렬 위치 좌표
                    pos_x = margin_x + int((target_w - final_w) / 2)
                    pos_y = margin_y + int((target_h - final_h) / 2)

                    hdc.StartPage()
                    dib = ImageWin.Dib(pimg)
                    dib.draw(hdc.GetHandleOutput(), (pos_x, pos_y, pos_x + final_w, pos_y + final_h))
                    hdc.EndPage()
                    
            hdc.EndDoc()
            hdc.DeleteDC()
            print(f"High-Res 300DPI Fit Print Success [{printer_name}]: {target_pdf}")
            return True
        except Exception as e1:
            print(f"Engine 1 (High-Res GDI) warning: {e1}")

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

    raise RuntimeError(f"프린터 [{printer_name}] 출력 실패. 프린터 상태를 확인하세요.")

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
