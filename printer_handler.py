import os
import sys
import subprocess
import tempfile
from pypdf import PdfReader, PdfWriter

def get_installed_printers():
    printers = []
    if os.name == 'nt':
        try:
            import win32print
            enum_flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printer_objs = win32print.EnumPrinters(enum_flags)
            for p in printer_objs:
                printers.append(p[2])
        except Exception:
            pass
    if not printers:
        printers = ["기본 프린터"]
    return printers

def print_pdf_file(pdf_path, printer_name=None, page_range=None):
    """
    SumatraPDF & Windows ShellExecute Native Print Engine (100% 안전 출력)
    page_range: list of int (0-indexed page indices, e.g. [11, 12] for pages 12-13)
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"인쇄할 PDF 파일을 찾을 수 없습니다: {pdf_path}")

    target_pdf = pdf_path

    # 특정 페이지 슬라이싱 요구 시 pypdf 사용
    if page_range is not None and len(page_range) > 0:
        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            
            valid_pages = [p for p in page_range if 0 <= p < total_pages]
            if valid_pages:
                for page_idx in valid_pages:
                    writer.add_page(reader.pages[page_idx])
                
                temp_dir = tempfile.gettempdir()
                sliced_pdf_path = os.path.join(temp_dir, f"sliced_{os.path.basename(pdf_path)}")
                with open(sliced_pdf_path, "wb") as f_out:
                    writer.write(f_out)
                target_pdf = sliced_pdf_path
        except Exception as e:
            print(f"Page slicing error: {e}, falling back to full print")

    # SumatraPDF.exe 다각 탐색
    candidate_paths = [
        os.path.join(getattr(sys, '_MEIPASS', ''), "bin", "SumatraPDF.exe"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin", "SumatraPDF.exe"),
        os.path.join(os.path.dirname(sys.executable), "bin", "SumatraPDF.exe"),
        os.path.join(os.path.dirname(sys.executable), "SumatraPDF.exe"),
        os.path.join(os.getcwd(), "bin", "SumatraPDF.exe"),
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    
    sumatra_exe = None
    for cp in candidate_paths:
        if cp and os.path.exists(cp):
            sumatra_exe = cp
            break

    # 기본 프린터 정제
    if not printer_name or printer_name == "기본 프린터":
        try:
            import win32print
            printer_name = win32print.GetDefaultPrinter()
        except Exception:
            printer_name = None

    # 1. SumatraPDF Engine
    if sumatra_exe and os.path.exists(sumatra_exe):
        cmd = [sumatra_exe]
        if printer_name:
            cmd.extend(["-print-to", printer_name])
        else:
            cmd.append("-print-default")
        cmd.extend(["-print-settings", "fit", target_pdf])
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
            if res.returncode == 0:
                return True
            else:
                print(f"SumatraPDF exited with code {res.returncode}: {res.stderr}")
        except Exception as e:
            print(f"SumatraPDF print failed: {e}")

    # 2. Windows ShellExecute / os.startfile Fallback
    if os.name == 'nt':
        try:
            import win32api
            import win32print
            if not printer_name:
                printer_name = win32print.GetDefaultPrinter()

            win32api.ShellExecute(0, "printto", target_pdf, f'"{printer_name}"', ".", 0)
            return True
        except Exception as e:
            print(f"ShellExecute print failed: {e}")
            try:
                os.startfile(target_pdf, "print")
                return True
            except Exception as e2:
                print(f"os.startfile print failed: {e2}")

    return False

if __name__ == '__main__':
    print("printer_handler ready without pdfplumber dependency!")
