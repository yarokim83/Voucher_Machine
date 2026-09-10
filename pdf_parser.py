import re
import os
import tempfile
import base64
import time
import html
from datetime import date
from voucher_core import extract_amounts
from file_utils import temp_output

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    from pypdf import PdfReader, PdfWriter
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

DEFAULT_PASSWORDS = [os.getenv('VOUCHERPASS_PDF_PASSWORD', '6068625399'), '']


def _parse_tounicode_cmap(data):
    cmap = {}
    if not data:
        return cmap
    # 1. parse beginbfchar
    for block in re.findall(r'beginbfchar(.*?)endbfchar', data, re.DOTALL):
        for src_hex, dst_hex in re.findall(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', block):
            try:
                src_cid = int(src_hex, 16)
                dst_char = bytes.fromhex(dst_hex).decode('utf-16be', errors='ignore')
                cmap[src_cid] = dst_char
            except Exception:
                pass

    # 2. parse beginbfrange
    for block in re.findall(r'beginbfrange(.*?)endbfrange', data, re.DOTALL):
        for start_hex, end_hex, dst_start_hex in re.findall(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', block):
            try:
                s_cid = int(start_hex, 16)
                e_cid = int(end_hex, 16)
                d_cid = int(dst_start_hex, 16)
                for cid in range(s_cid, e_cid + 1):
                    cmap[cid] = chr(d_cid + (cid - s_cid))
            except Exception:
                pass
    return cmap


def extract_pdf_text_with_cmaps(pdf_path, passwords=DEFAULT_PASSWORDS):
    """
    ToUnicode CMap을 직접 파싱하여 CIDFont로 임베딩된 한글 폰트를 완벽하게 추출.
    """
    if not HAS_PYPDF:
        return ""
    try:
        reader = PdfReader(pdf_path)
    except Exception:
        return ""

    if reader.is_encrypted:
        for pwd in passwords:
            try:
                if reader.decrypt(pwd):
                    break
            except Exception:
                pass

    full_pages = []
    try:
        for page in reader.pages:
            fonts = page.get('/Resources', {}).get('/Font', {})
            cmaps = {}
            for fname, fref in fonts.items():
                try:
                    fobj = fref.get_object()
                    to_u = fobj.get('/ToUnicode')
                    if to_u:
                        data = to_u.get_object().get_data().decode('latin1', errors='ignore')
                        cmaps[fname] = _parse_tounicode_cmap(data)
                    else:
                        cmaps[fname] = {}
                except Exception:
                    cmaps[fname] = {}

            contents = page.get_contents()
            if not contents:
                continue
            raw = contents.get_data() if hasattr(contents, 'get_data') else b''
            current_font = None
            chunks = []

            for b in raw.split(b'BT'):
                if b'ET' not in b:
                    continue
                bt_body = b.split(b'ET')[0]
                font_m = re.findall(rb'(/F\d+)\s+[\d.]+\s+Tf', bt_body)
                if font_m:
                    current_font = font_m[-1].decode('latin1')
                tjs = re.findall(rb'<([0-9A-Fa-f]+)>\s*Tj', bt_body)
                for tj in tjs:
                    cmap = cmaps.get(current_font, {})
                    hex_str = tj.decode('latin1')
                    chars = []
                    for i in range(0, len(hex_str), 4):
                        cid = int(hex_str[i:i+4], 16)
                        if cid in cmap:
                            chars.append(cmap[cid])
                        elif cid == 3:
                            chars.append(' ')
                        elif 32 <= cid <= 126:
                            chars.append(chr(cid))
                    if chars:
                        chunks.append(''.join(chars))
            if chunks:
                full_pages.append(' '.join(chunks))
    except Exception:
        pass

    return '\n'.join(full_pages)


def extract_pdf_text_safely(pdf_path, passwords=DEFAULT_PASSWORDS):
    # Prefer layout-preserving extraction. CMap is a fallback for unreadable text.
    for backend in ('plumber', 'pypdf'):
        for pwd in passwords:
            try:
                if backend == 'plumber' and HAS_PDFPLUMBER:
                    with pdfplumber.open(pdf_path, password=pwd) as pdf:
                        text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
                elif backend == 'pypdf' and HAS_PYPDF:
                    reader = PdfReader(pdf_path)
                    if reader.is_encrypted and not reader.decrypt(pwd):
                        continue
                    text = '\n'.join(page.extract_text() or '' for page in reader.pages)
                else:
                    continue
                if text.strip() and '(cid:' not in text and '\ufffd' not in text:
                    return text
            except Exception:
                continue
    return extract_pdf_text_with_cmaps(pdf_path, passwords)


def _read_html_text(html_path):
    with open(html_path, 'rb') as stream:
        raw = stream.read()
    encodings = ['utf-16'] if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else ['utf-8-sig', 'cp949', 'euc-kr']
    raw_content = None
    for encoding in encodings:
        try:
            raw_content = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if raw_content is None:
        raise ValueError('HTML 문자 인코딩을 확인할 수 없습니다.')
    text = re.sub(r'<(script|style)\b[^>]*>.*?</\1>', '', raw_content, flags=re.I | re.S)
    text = re.sub(r'</(?:tr|p|div)>|<br\s*/?>', '\n', text, flags=re.I)
    text = html.unescape(re.sub(r'<[^>]+>', ' ', text))
    return '\n'.join(re.sub(r'[^\S\n]+', ' ', line).strip() for line in text.splitlines())


def _is_html_file(file_path):
    return file_path.lower().endswith(('.html', '.htm'))

def _log_debug(msg):
    if os.getenv('VOUCHERPASS_DEBUG') != '1':
        return
    log_locations = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'voucher_pass_debug.log'),
        os.path.join(os.getenv('APPDATA', '.'), 'VoucherPass', 'voucher_pass_debug.log'),
    ]
    formatted = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n"
    print(f"[DEBUG] {msg}")
    for log_path in log_locations:
        try:
            os.makedirs(os.path.dirname(log_path), exist_ok=True)
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(formatted)
        except Exception:
            pass

def parse_tax_invoice_data(file_path):
    """
    세금계산서 (HTML 또는 PDF)에서 작성일자 및 공급가액/세액/합계금액을 추출.
    Returns: dict(date=str, supply_amount=int, vat=int, total_amount=int)
    """
    _log_debug(f"=== [parse_tax_invoice_data START] File: {file_path} ===")
    result = {
        'date': '',
        'supply_amount': None,
        'vat': None,
        'total_amount': None
    }
    if not os.path.exists(file_path):
        _log_debug(f"ERROR: File does not exist: {file_path}")
        return result

    _log_debug(f"File size: {os.path.getsize(file_path)} bytes")

    if _is_html_file(file_path):
        _log_debug("File type: HTML")
        full_text = _read_html_text(file_path)
    else:
        _log_debug("File type: PDF")
        try:
            dec_pdf = decrypt_pdf_to_temp(file_path)
            target_p = dec_pdf if (dec_pdf and os.path.exists(dec_pdf)) else file_path
            _log_debug(f"Decrypted PDF path: {target_p}")
        except Exception as e:
            _log_debug(f"Decrypt exception: {e}")
            target_p = file_path

        full_text = extract_pdf_text_safely(target_p)
        _log_debug(f"Safely extracted PDF text length: {len(full_text)}")

        # 표(Table) 셀 텍스트도 추가 평탄화 수집
        if HAS_PDFPLUMBER:
            try:
                for pwd in DEFAULT_PASSWORDS:
                    try:
                        with pdfplumber.open(target_p, password=pwd) as pdf:
                            _log_debug(f"pdfplumber opened PDF. Pages count: {len(pdf.pages)}")
                            for idx, page in enumerate(pdf.pages):
                                tables = page.extract_tables()
                                _log_debug(f"Page {idx+1} tables found: {len(tables)}")
                                for tbl in tables:
                                    for row in tbl:
                                        if row:
                                            row_str = " ".join([str(cell) for cell in row if cell is not None])
                                            full_text += "\n" + row_str
                        break
                    except Exception as pe:
                        _log_debug(f"pdfplumber page table extract exception with pwd '{pwd}': {pe}")
                        continue
            except Exception as pe_outer:
                _log_debug(f"pdfplumber outer exception: {pe_outer}")

    if not full_text or not full_text.strip():
        _log_debug("ERROR: full_text is empty after extraction!")
        return result

    _log_debug(f"Total full_text length: {len(full_text)}")

    # 1. 작성일자 추출
    kw_pos = full_text.find("작성일자")
    if kw_pos == -1:
        m_kw = re.search(r'작\s*성\s*일\s*자', full_text)
        if m_kw:
            kw_pos = m_kw.start()

    if kw_pos != -1:
        sub_text = full_text[kw_pos:kw_pos + 150]
        _log_debug(f"Found '작성일자' at pos {kw_pos}. Subtext: {repr(sub_text)}")

        m = re.search(r'(20[1-3][0-9])[\s/.\-년]+\s*(\d{1,2})[\s/.\-월]+\s*(\d{1,2})', sub_text)
        if m:
            y, month, d = m.group(1), int(m.group(2)), int(m.group(3))
            if 1 <= month <= 12 and 1 <= d <= 31:
                result['date'] = f"{y}-{month:02d}-{d:02d}"
                _log_debug(f"SUCCESS [Step 1-A Keyword '작성일자']: {result['date']}")

        if not result['date']:
            m_dig = re.search(r'(20[1-3][0-9])(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])', sub_text)
            if m_dig:
                y, month, d = m_dig.group(1), int(m_dig.group(2)), int(m_dig.group(3))
                result['date'] = f"{y}-{month:02d}-{d:02d}"
                _log_debug(f"SUCCESS [Step 1-B Keyword '작성일자' digits]: {result['date']}")

    if not result['date']:
        kw_pos_bal = full_text.find("발행일자")
        if kw_pos_bal == -1:
            kw_pos_bal = full_text.find("공급일자")
        if kw_pos_bal == -1:
            m_kw_bal = re.search(r'(?:발\s*행\s*일\s*자|공\s*급\s*일\s*자)', full_text)
            if m_kw_bal:
                kw_pos_bal = m_kw_bal.start()

        if kw_pos_bal != -1:
            sub_text = full_text[kw_pos_bal:kw_pos_bal + 150]
            _log_debug(f"Found '발행일자/공급일자' at pos {kw_pos_bal}. Subtext: {repr(sub_text)}")
            m = re.search(r'(20[1-3][0-9])[\s/.\-년]+\s*(\d{1,2})[\s/.\-월]+\s*(\d{1,2})', sub_text)
            if m:
                y, month, d = m.group(1), int(m.group(2)), int(m.group(3))
                if 1 <= month <= 12 and 1 <= d <= 31:
                    result['date'] = f"{y}-{month:02d}-{d:02d}"
                    _log_debug(f"SUCCESS [Step 2 Keyword '발행일자/공급일자']: {result['date']}")

    if not result['date'] and not _is_html_file(file_path):
        matches = re.findall(r'(20[1-3][0-9])[\s/.\-년]+\s*(\d{1,2})[\s/.\-월]+\s*(\d{1,2})[일\s]?', full_text)
        if matches:
            for mat in matches:
                y, month, d = mat[0], int(mat[1]), int(mat[2])
                if 1 <= month <= 12 and 1 <= d <= 31:
                    result['date'] = f"{y}-{month:02d}-{d:02d}"
                    _log_debug(f"SUCCESS [Step 3-A Fulltext 년월일]: {result['date']}")
                    break

        if not result['date']:
            matches_fmt = re.findall(r'(20[1-3][0-9])[-.\s/](\d{1,2})[-.\s/](\d{1,2})', full_text)
            if matches_fmt:
                for mat in matches_fmt:
                    y, month, d = mat[0], int(mat[1]), int(mat[2])
                    if 1 <= month <= 12 and 1 <= d <= 31:
                        result['date'] = f"{y}-{month:02d}-{d:02d}"
                        _log_debug(f"SUCCESS [Step 3-B Fulltext YYYY-MM-DD]: {result['date']}")
                        break

    amounts = extract_amounts(full_text)
    result.update(supply_amount=amounts['amount'], vat=amounts['vat'],
                  total_amount=amounts['total_amount'], warnings=amounts['warnings'])
    if result['date']:
        try:
            date.fromisoformat(result['date'])
        except ValueError:
            result['date'] = ''
            result['warnings'].append('유효하지 않은 작성일자')

    _log_debug(f"=== [parse_tax_invoice_data END] Result: {result} ===")
    return result

def parse_tax_invoice_date(file_path):
    data = parse_tax_invoice_data(file_path)
    return data.get('date', '')

def parse_pr_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")

    extracted_data = {
        'pr_no': '',
        'pr_title': '',
        'amount': None,
        'vat': None,
        'total_amount': None,
        'date': '',
        'supplier': ''
    }

    full_text = extract_pdf_text_safely(pdf_path)

    pr_match = re.search(r'(?:P/R|PR|PO|발주)\s*No\.?\s*[:\s]*([A-Z0-9]+)', full_text, re.IGNORECASE)
    if not pr_match:
        pr_match = re.search(r'발주번호\s*[:\s]*([A-Z0-9]+)', full_text)
    if not pr_match:
        pr_match = re.search(r'S202[0-9]{8}', full_text)
    
    if pr_match:
        extracted_data['pr_no'] = pr_match.group(1) if len(pr_match.groups()) > 0 else pr_match.group(0)

    date_match = re.search(r'(?:DATE|작성일자|발주일자|발행일자)\s*[:\s]*(\d{4}[-.\s]\d{2}[-.\s]\d{2})', full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', full_text)
    
    if date_match:
        extracted_data['date'] = date_match.group(1).replace('.', '-').strip()

    subj_match = re.search(r'(?:SUBJECT|품명|공사명|건명|발주명)\s*[:\s]*(.+)', full_text, re.IGNORECASE)
    if subj_match:
        extracted_data['pr_title'] = subj_match.group(1).strip()

    extracted_data.update(extract_amounts(full_text))
    if not full_text.strip():
        extracted_data['warnings'].append('텍스트를 읽을 수 없습니다. 스캔/암호화 문서를 확인해 주세요.')

    # 거래처명 (Supplier) 추출
    def _clean_sup(s):
        if not s:
            return ''
        s = re.sub(r'[\r\n\t]+', ' ', s)
        s = re.sub(r'\s*\(\s*', '(', s)
        s = re.sub(r'\s*\)\s*', ')', s)
        s = re.sub(r'^\d+[\s.]*', '', s)
        s = re.sub(r'\s+\d{9,12}.*$', '', s)
        s = re.sub(r'\s+010.*$', '', s)
        s = re.sub(r'\s+[a-zA-Z0-9_.+-]+@.*$', '', s)
        # 한국어 글자 사이 공백 정규화 (예: '코 텍' -> '코텍')
        if re.match(r'^[가-힣]\s+[가-힣](\s+[가-힣])?', s):
            s = re.sub(r'\s+', '', s)
        return s.strip()

    sup_match = None
    
    # 0. 발주서 (PO) - 2. 공급처(SUPPLIER) 섹션 추출 (우선순위 1)
    section = re.search(r'2\s*\.\s*공\s*급\s*처\s*\(\s*SUPPLIER\s*\)', full_text, re.I)
    idx_sup_sec = section.start() if section else -1
    if idx_sup_sec != -1:
        sub_sup = full_text[section.end():section.end() + 400]
        sub_sup = re.split(r'\n\s*3\s*\.', sub_sup, maxsplit=1)[0]
        sub_sup = re.sub(r'(?i)2\.\s*공급처\(SUPPLIER\)', '', sub_sup)
        sub_sup = re.sub(r'(?i)COMPANY\s*NAME', '', sub_sup)
        sub_sup = re.sub(r'(?i)DEPARTMENT', '', sub_sup)
        sub_sup = re.sub(r'(?i)NAME', '', sub_sup)
        sub_sup = re.sub(r'(?i)TITLE', '', sub_sup)
        sub_sup = re.sub(r'(?i)TEL\.?', '', sub_sup)
        sub_sup = re.sub(r'(?i)FAX\.?', '', sub_sup)
        sub_sup = re.sub(r'(?i)E-mail', '', sub_sup)
        
        lines_left = [ln.strip() for ln in sub_sup.split('\n') if ln.strip()]
        if lines_left:
            first_line = lines_left[0]
            parts = re.split(r'\s{2,}', first_line)
            cand = _clean_sup(parts[0])
            if cand and cand not in ('Company', 'Name', 'No', 'Title'):
                extracted_data['supplier'] = cand

    if not extracted_data['supplier']:
        # 1. SUPPLIERS RECOMMENDED 테이블 패턴
        idx_rec = full_text.find('SUPPLIERS')
        if idx_rec != -1:
            sub_rec = full_text[idx_rec:idx_rec + 400]
            sup_match = re.search(r'1\s+([가-힣A-Za-z0-9㈜(주)\s\(\)._-]{2,30}?)(?=\s+[가-힣]{2,4}\s+\d|\s+010|\s+\d{9,12}|\s+[a-zA-Z0-9_.+-]+@|\s+PREPARED|\n|$)', sub_rec)
            if sup_match:
                cand = _clean_sup(sup_match.group(1))
                if cand and cand not in ('Company', 'Name', 'No', 'Title'):
                    extracted_data['supplier'] = cand

    if not extracted_data['supplier']:
        # 2. Company / Vendor 명시적 필드 패턴
        sup_match = re.search(r'Company\s+(?:Name\s+)?(?:Title\s+)?(?:1\s+)?([가-힣A-Za-z0-9㈜(주)\s\(\)._-]{2,30}?)(?=\s+[가-힣]{2,4}\s+\d|\s+010|\s+\d{9,12}|\s+[a-zA-Z0-9_.+-]+@|\s+PREPARED|\n|$)', full_text)
        if sup_match:
            cand = _clean_sup(sup_match.group(1))
            if cand and cand not in ('Company', 'Name', 'No', 'Title'):
                extracted_data['supplier'] = cand

    if not extracted_data['supplier']:
        # 3. 공급자 / 거래처 / 공급업체 라벨 패턴
        sup_match = re.search(r'(?:공급자|공급업체|거래처|업체명|상호|협력업체|추천업체|Vendor|Supplier)\s*[:\s]*([가-힣A-Za-z0-9㈜(주)\s\(\)._-]{2,30})', full_text, re.IGNORECASE)
        if sup_match:
            extracted_data['supplier'] = _clean_sup(sup_match.group(1))

    return extracted_data

def classify_pdf_type(file_path):
    if not os.path.exists(file_path):
        return 'unknown'

    fname = os.path.basename(file_path).lower()
    if 'nts_etaxinvoice' in fname or 'etaxinvoice' in fname:
        return 'tax'

    if '발주' in fname or re.search(r'(?<![a-z])po(?=\d|[^a-z]|$)', fname) or 'order' in fname:
        return 'po'
    if '구매요청' in fname or re.search(r'(?<![a-z])pr(?=\d|[^a-z]|$)', fname) or 'requisition' in fname:
        return 'pr'
    if '명세서' in fname or '거래' in fname or 'spec' in fname:
        return 'spec'
    if '세금' in fname or '계산서' in fname or 'tax' in fname:
        return 'tax'
    if '계약' in fname or 'contract' in fname or '협약' in fname:
        return 'contract'

    if _is_html_file(file_path):
        txt = _read_html_text(file_path)
    else:
        txt = extract_pdf_text_safely(file_path)

    txt_upper = txt.upper()
    if 'PURCHASE ORDER' in txt_upper or '발주서' in txt or 'P.O' in txt_upper:
        return 'po'
    if 'PURCHASE REQUISITION' in txt_upper or '구매요청서' in txt or 'P/R NO' in txt_upper:
        return 'pr'
    if '전자세금계산서' in txt or '세금계산서' in txt or 'TAX INVOICE' in txt_upper:
        return 'tax'
    if '거래명세서' in txt or '명세서' in txt:
        return 'spec'
    if '계약서' in txt or '협약서' in txt or 'CONTRACT' in txt_upper:
        return 'contract'

    return 'unknown'

def decrypt_pdf_to_temp(pdf_path, passwords=DEFAULT_PASSWORDS):
    if not HAS_PYPDF:
        return pdf_path

    try:
        reader = PdfReader(pdf_path)
        if not reader.is_encrypted:
            return pdf_path

        decrypted = False
        for pwd in passwords:
            try:
                if reader.decrypt(pwd):
                    decrypted = True
                    break
            except Exception:
                continue

        if not decrypted:
            return pdf_path

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        temp_pdf = temp_output(f'decrypted_{os.path.basename(pdf_path)}')
        with open(temp_pdf, 'wb') as f:
            writer.write(f)

        return temp_pdf
    except Exception:
        return pdf_path
