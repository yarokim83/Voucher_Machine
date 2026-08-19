import re
import os
import tempfile
import base64
import time

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

DEFAULT_PASSWORDS = ['6068625399', '']

def extract_pdf_text_safely(pdf_path, passwords=DEFAULT_PASSWORDS):
    full_text = ""
    if HAS_PDFPLUMBER:
        for pwd in passwords:
            try:
                with pdfplumber.open(pdf_path, password=pwd) as pdf:
                    for page in pdf.pages:
                        txt = page.extract_text()
                        if txt:
                            full_text += txt + "\n"
                if full_text.strip():
                    return full_text
            except Exception:
                continue

    if HAS_PYPDF:
        for pwd in passwords:
            try:
                reader = PdfReader(pdf_path)
                if reader.is_encrypted:
                    try:
                        reader.decrypt(pwd)
                    except Exception:
                        continue
                for page in reader.pages:
                    txt = page.extract_text()
                    if txt:
                        full_text += txt + "\n"
                if full_text.strip():
                    return full_text
            except Exception:
                continue

    return full_text

def _read_html_text(html_path):
    raw_content = ""
    for enc in ['utf-8', 'euc-kr', 'cp949', 'utf-16']:
        try:
            with open(html_path, 'r', encoding=enc, errors='ignore') as f:
                raw_content = f.read()
            if raw_content:
                break
        except Exception:
            continue

    if not raw_content:
        return ''

    text = re.sub(r'<[^>]+>', ' ', raw_content)
    text = re.sub(r'&[a-zA-Z]+;', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text

def _is_html_file(file_path):
    return file_path.lower().endswith(('.html', '.htm'))

def _log_debug(msg):
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
        'supply_amount': 0,
        'vat': 0,
        'total_amount': 0
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

    # 2. 공급가액(supply_amount) / 세액(vat) / 합계금액(total_amount) 추출
    # 패턴 1: 작성일자 바로 뒤에 나오는 금액들 (예: 2026/08/11 2,100,000 210,000)
    m_amt_seq = re.search(r'20[1-3][0-9][/\-.\s년]+\d{1,2}[/\-.\s월]+\d{1,2}[일\s]*\s+([1-9]\d{0,2}(?:,\d{3})+)\s+([0-9]\d{0,2}(?:,\d{3})*)', full_text)
    if m_amt_seq:
        s_amt = int(m_amt_seq.group(1).replace(',', ''))
        v_amt = int(m_amt_seq.group(2).replace(',', ''))
        if s_amt > 0:
            result['supply_amount'] = s_amt
            result['vat'] = v_amt
            result['total_amount'] = s_amt + v_amt
            _log_debug(f"SUCCESS [Amount Step 1 Date-Sequence]: supply={s_amt:,}, vat={v_amt:,}, total={s_amt+v_amt:,}")

    # 패턴 2: 키워드 '공급가액' 뒤의 금액 추출
    if result['supply_amount'] == 0:
        kw_amt = re.search(r'(?:공\s*급\s*가\s*액)\s*[:\s]*([1-9]\d{0,2}(?:,\d{3})+)', full_text)
        if kw_amt:
            s_amt = int(kw_amt.group(1).replace(',', ''))
            result['supply_amount'] = s_amt
            result['vat'] = int(s_amt * 0.1)
            result['total_amount'] = s_amt + result['vat']
            _log_debug(f"SUCCESS [Amount Step 2 Keyword '공급가액']: supply={s_amt:,}")

    # 패턴 3: 전체 금액 리스트에서 supply + vat = total 수학적 검증
    if result['supply_amount'] == 0:
        all_amounts = re.findall(r'([1-9]\d{0,2}(?:,\d{3})+)', full_text)
        num_set = set([int(a.replace(',', '')) for a in all_amounts if 1000 <= int(a.replace(',', '')) <= 10000000000])
        for a in sorted(num_set, reverse=True):
            for v in num_set:
                if (a + v) in num_set and (v == int(a * 0.1) or v == 0 or v == round(a * 0.1)):
                    result['supply_amount'] = a
                    result['vat'] = v
                    result['total_amount'] = a + v
                    _log_debug(f"SUCCESS [Amount Step 3 Math Verification]: supply={a:,}, vat={v:,}, total={a+v:,}")
                    break
            if result['supply_amount'] > 0:
                break

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
        'amount': 0,
        'vat': 0,
        'total_amount': 0,
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

    pr_no_str = extracted_data['pr_no']
    candidates = []
    
    lines = full_text.split('\n')
    for line in lines:
        if pr_no_str and pr_no_str in line:
            continue
        
        amounts = re.findall(r'(?:KRW|₩|총금액|공급가액|총약정금액|발주금액|Total)?\s*([1-9]\d{0,2}(?:,\d{3})+)', line)
        for amt_str in amounts:
            clean_num = int(amt_str.replace(',', ''))
            if pr_no_str and str(clean_num) in pr_no_str:
                continue
            if len(str(clean_num)) >= 10:
                continue
            if 1000 <= clean_num <= 1000000000:
                is_priority = any(k in line for k in ['KRW', '총약정금액', '발주금액', 'Total', '공급가액', '합계'])
                candidates.append((clean_num, is_priority))

    priority_candidates = [c[0] for c in candidates if c[1]]
    if priority_candidates:
        extracted_data['amount'] = max(priority_candidates)
    elif candidates:
        extracted_data['amount'] = max([c[0] for c in candidates])

    if extracted_data['amount'] > 0:
        extracted_data['vat'] = int(extracted_data['amount'] * 0.1)
        extracted_data['total_amount'] = extracted_data['amount'] + extracted_data['vat']

    sup_match = re.search(r'Company\s+([가-힣A-Za-z0-9㈜(주)]+)', full_text)
    if not sup_match:
        sup_match = re.search(r'SUPPLIERS?\s*RECOMMENDED[\s\S]*?1\s+([가-힣A-Za-z0-9㈜(주)]+)', full_text)
    if not sup_match:
        sup_match = re.search(r'공급자\s*[:\s]*([가-힣A-Za-z0-9㈜(주)]+)', full_text)
    if not sup_match:
        sup_match = re.search(r'금강엔지니어링', full_text)
    
    if sup_match:
        extracted_data['supplier'] = sup_match.group(1) if len(sup_match.groups()) > 0 else sup_match.group(0)

    return extracted_data

def classify_pdf_type(file_path):
    if not os.path.exists(file_path):
        return 'unknown'

    fname = os.path.basename(file_path).lower()
    if 'nts_etaxinvoice' in fname or 'etaxinvoice' in fname:
        return 'tax'

    if '발주' in fname or 'po' in fname or 'order' in fname:
        return 'po'
    if '구매요청' in fname or 'pr' in fname or 'requisition' in fname:
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

    return 'pr'

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

        temp_pdf = os.path.join(tempfile.gettempdir(), f"decrypted_{os.path.basename(pdf_path)}")
        with open(temp_pdf, 'wb') as f:
            writer.write(f)

        return temp_pdf
    except Exception:
        return pdf_path
