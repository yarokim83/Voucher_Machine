import os
import shutil
import tempfile
from pathlib import Path
import openpyxl
from file_utils import desktop_dir, safe_name


def generate_voucher_excel(data, template_path=None):
    if not template_path or not os.path.isfile(template_path):
        raise FileNotFoundError('전표 Excel 템플릿을 선택해 주세요.')
    amounts = [data.get(key) for key in ('amount', 'vat', 'total_amount')]
    if any(value is None for value in amounts) or amounts[0] + amounts[1] != amounts[2]:
        raise ValueError('공급가액·부가세·합계금액을 확인해 주세요.')
    wb = openpyxl.load_workbook(template_path)
    try:
        ws = wb.active
        mapping = {'W6': 'pr_no', 'D9': 'pr_title', 'W9': 'amount', 'W16': 'vat',
                   'W17': 'total_amount', 'A6': 'date', 'J6': 'supplier',
                   'P32': 'total_amount', 'P33': 'amount', 'V34': 'vat'}
        for cell, key in mapping.items():
            value = data.get(key, '')
            ws[cell] = value
            if isinstance(value, str):
                ws[cell].data_type = 's'
        desktop = desktop_dir()
        desktop.mkdir(parents=True, exist_ok=True)
        fd, name = tempfile.mkstemp(prefix=f"Voucher_{safe_name(data.get('pr_no') or 'NO_PR')}_", suffix='.xlsx', dir=desktop)
        os.close(fd)
        try:
            wb.save(name)
        except Exception:
            Path(name).unlink(missing_ok=True)
            raise
        return name
    finally:
        wb.close()


def archive_voucher_package(data, pdf_paths, excel_path=None, base_dir=None):
    items = list(pdf_paths.items()) if isinstance(pdf_paths, dict) else [(str(i), p) for i, p in enumerate(pdf_paths, 1)]
    if excel_path:
        items.append(('voucher', excel_path))
    items = [(kind, Path(path)) for kind, path in items if path]
    if not items:
        raise ValueError('보관할 서류가 없습니다.')
    for _, path in items:
        if not path.is_file():
            raise FileNotFoundError(f'보관할 파일을 찾을 수 없습니다: {path.name}')
    vault = Path(base_dir) if base_dir else desktop_dir() / 'Voucher_보관소'
    vault.mkdir(parents=True, exist_ok=True)
    label = '_'.join(safe_name(data.get(key) or fallback, fallback) for key, fallback in [('date', '날짜미확인'), ('supplier', '거래처미확인'), ('pr_no', 'PR미확인')])
    target = Path(tempfile.mkdtemp(prefix=label[:120] + '_', dir=vault))
    archived = []
    try:
        for index, (kind, path) in enumerate(items, 1):
            dest = target / f'{index:02d}_{safe_name(kind)}_{safe_name(path.name)}'
            shutil.copy2(path, dest)
            archived.append(str(dest))
    except Exception:
        # target is the unique directory created above, never a caller-supplied path.
        shutil.rmtree(target)
        raise
    return str(target), archived
