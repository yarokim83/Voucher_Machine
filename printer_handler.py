import os
import sys
import subprocess
import tempfile
from pypdf import PdfReader, PdfWriter
from voucher_core import validate_pages
from file_utils import temp_output


def print_documents(paths, printer_name, contract_pages):
    """Prepare the complete batch before submitting any irreversible print request."""
    import shutil
    import pdf_parser
    prepared = []
    for kind, path in paths.items():
        if not path:
            continue
        if not os.path.isfile(path) or not path.lower().endswith('.pdf'):
            raise ValueError(f'{kind}: PDF 파일 또는 변환 완료 여부를 확인해 주세요.')
        snapshot = temp_output(os.path.basename(path))
        shutil.copy2(path, snapshot)
        target = pdf_parser.decrypt_pdf_to_temp(snapshot)
        reader = PdfReader(target)
        if not reader.pages:
            raise ValueError(f'{kind}: PDF에 페이지가 없습니다.')
        if kind == 'contract':
            pages = validate_pages(contract_pages, len(reader.pages))
            writer = PdfWriter()
            for index in pages:
                writer.add_page(reader.pages[index])
            target = temp_output('contract_pages.pdf')
            with open(target, 'wb') as stream:
                writer.write(stream)
        prepared.append((kind, target))
    if not prepared:
        raise ValueError('인쇄할 서류가 없습니다.')
    submitted, failed = [], []
    for kind, target in prepared:
        try:
            if print_pdf_file(target, printer_name):
                submitted.append(kind)
            else:
                failed.append(f'{kind}: 인쇄 엔진 없음')
        except Exception as exc:
            failed.append(f'{kind}: {exc}')
    return submitted, failed

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

    if not pdf_path.lower().endswith('.pdf'):
        raise ValueError('PDF 변환이 완료된 문서만 인쇄할 수 있습니다.')
    # Validate the entire requested range; never substitute the original document.
    if page_range is not None:
        reader = PdfReader(pdf_path)
        valid_pages = validate_pages(page_range, len(reader.pages))
        writer = PdfWriter()
        for page_idx in valid_pages:
            writer.add_page(reader.pages[page_idx])
        target_pdf = temp_output(f'sliced_{os.path.basename(pdf_path)}')
        with open(target_pdf, 'wb') as stream:
            writer.write(stream)

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
                raise RuntimeError(f"SumatraPDF 인쇄 요청 실패: {res.returncode}")
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("인쇄 요청 시간이 초과됐습니다. 중복 인쇄 방지를 위해 대체 출력을 중단했습니다. 프린터 대기열을 확인해 주세요.") from exc
        except Exception as exc:
            raise RuntimeError(f"인쇄 요청 실패: {exc}") from exc

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
            raise RuntimeError('선택한 프린터에 인쇄 요청을 전달하지 못했습니다.') from e

    return False

if __name__ == '__main__':
    print("printer_handler ready without pdfplumber dependency!")
