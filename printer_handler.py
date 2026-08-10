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

def find_sumatra_path():
    """
    프로젝트 bin 디렉토리 또는 시스템의 SumatraPDF.exe 탐색
    """
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
    [벡터 Vector Direct Spooling] 폰트 화질 100% 보존 & 종이 핏팅 잘림 방지 전용 인쇄 엔진
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

    if not printer_name or printer_name == "기본 프린터 (Default Printer)":
        printer_name = win32print.GetDefaultPrinter()

    # -------------------------------------------------------------
    # Engine 1: SumatraPDF Portable Vector Native Engine (최우선 1순위)
    # 텍스트 폰트 100% 칼같은 선명도 유지 + 'fit' 옵션으로 0.1mm도 안 잘림!
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

    # -------------------------------------------------------------
    # Engine 4: 300 DPI High-Res Direct GDI Print (최후 백업)
    # -------------------------------------------------------------
    if HAS_WIN32:
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            
            pw = max(100, hdc.GetDeviceCaps(110))   # PHYSICALWIDTH
            ph = max(100, hdc.GetDeviceCaps(111))  # PHYSICALHEIGHT
            off_x = hdc.GetDeviceCaps(112) # PHYSICALOFFSETX
            off_y = hdc.GetDeviceCaps(113) # PHYSICALOFFSETY
            res_x = hdc.GetDeviceCaps(88)  # LOGPIXELSX
            res_y = hdc.GetDeviceCaps(90)  # LOGPIXELSY

            margin_x = int(res_x * 0.15) if res_x > 0 else 50
            margin_y = int(res_y * 0.15) if res_y > 0 else 50

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
                    pimg = page.to_image(resolution=300).original
                    img_w, img_h = pimg.size

                    scale = min(target_w / img_w, target_h / img_h)
                    final_w = int(img_w * scale)
                    final_h = int(img_h * scale)

                    pos_x = margin_x + int((target_w - final_w) / 2)
                    pos_y = margin_y + int((target_h - final_h) / 2)

                    hdc.StartPage()
                    dib = ImageWin.Dib(pimg)
                    dib.draw(hdc.GetHandleOutput(), (pos_x, pos_y, pos_x + final_w, pos_y + final_h))
                    hdc.EndPage()
                    
            hdc.EndDoc()
            hdc.DeleteDC()
            print(f"[Engine 4: High-Res GDI Fit Print] Success to [{printer_name}]: {target_pdf}")
            return True
        except Exception as e4:
            print(f"Engine 4 (GDI) warning: {e4}")

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
    print("printer_handler Vector Native Architecture Ready!")
    print("Printers:", get_installed_printers())
