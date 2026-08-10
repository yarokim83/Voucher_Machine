import os
import sys
import tempfile

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

def print_pdf_file(pdf_path, printer_name=None, page_range=None):
    """
    PDF 파일 또는 특정 페이지 범위 인쇄
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"인쇄할 PDF 파일을 찾을 수 없습니다: {pdf_path}")

    target_pdf = pdf_path

    if HAS_WIN32 and sys.platform == 'win32':
        if printer_name:
            try:
                win32print.SetDefaultPrinter(printer_name)
            except Exception:
                pass
        
        win32api.ShellExecute(0, "print", target_pdf, None, ".", 0)
        return True
    else:
        print(f"[Simulation Mode] Printing PDF: {target_pdf}")
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
            
            if printer_name:
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
