import os
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
    SumatraPDF Vector Native Print Engine (-print-settings "fit") 사용 100% 선명 벡터 출력
    page_range: list of int (0-indexed page indices to print, e.g. [11, 12] for pages 12-13)
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

    base_dir = os.path.dirname(os.path.abspath(__file__))
    sumatra_exe = os.path.join(base_dir, "bin", "SumatraPDF.exe")

    # 1. SumatraPDF Engine
    if os.path.exists(sumatra_exe):
        cmd = [sumatra_exe, "-print-to", printer_name if printer_name else "default", "-print-settings", "fit", target_pdf]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
            if res.returncode == 0:
                return True
        except Exception as e:
            print(f"SumatraPDF print failed: {e}")

    # 2. Windows ShellExecute Fallback
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

    return False

if __name__ == '__main__':
    print("printer_handler ready without pdfplumber dependency!")
