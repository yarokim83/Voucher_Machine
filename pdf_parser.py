import re
import os
import tempfile
import base64

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
    try:
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'voucher_pass_debug.log')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass

def parse_tax_invoice_date(file_path):
    import time
    _log_debug(f"parse_tax_invoice_date started for: {file_path}")
    if not os.path.exists(file_path):
        _log_debug("File does not exist!")
        return ''

    if _is_html_file(file_path):
        full_text = _read_html_text(file_path)
    else:
        # 암호 해제 시도 후 PDF 텍스트 추출
        try:
            dec_pdf = decrypt_pdf_to_temp(file_path)
            if dec_pdf and os.path.exists(dec_pdf):
                full_text = extract_pdf_text_safely(dec_pdf)
            else:
                full_text = extract_pdf_text_safely(file_path)
        except Exception:
            full_text = extract_pdf_text_safely(file_path)

    _log_debug(f"Extracted full_text length: {len(full_text)}")
    if not full_text or not full_text.strip():
        _log_debug("full_text is empty!")
        return ''

    _log_debug(f"Full Text Preview: {full_text[:400]}")

    # 1. 작성/발행/공급 키워드 주변 50자 이내 202X 날짜 최우선 탐색
    kw_matches = re.findall(r'(?:작\s*성\s*일\s*자?|발\s*행\s*일\s*자?|공\s*급\s*일\s*자?|작\s*성\s*년\s*월\s*일|작\s*성|일\s*자)\s*[:\s\n]*([^\n]{1,50})', full_text)
    for sub in kw_matches:
        m = re.search(r'(202[0-9]|203[0-9])[\s년\-./\\]+(\d{1,2})[\s월\-./\\]+(\d{1,2})', sub)
        if m:
            y, month, d = m.group(1), int(m.group(2)), int(m.group(3))
            if 1 <= month <= 12 and 1 <= d <= 31:
                res = f"{y}-{month:02d}-{d:02d}"
                _log_debug(f"Match Step 1 Keyword SUCCESS: {res}")
                return res
        
        m_dig = re.search(r'(202[0-9]|203[0-9])(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])', sub)
        if m_dig:
            y, month, d = m_dig.group(1), int(m_dig.group(2)), int(m_dig.group(3))
            res = f"{y}-{month:02d}-{d:02d}"
            _log_debug(f"Match Step 1 Digits SUCCESS: {res}")
            return res

    # 2. 전천후 202X년/월/일 파싱 (다양한 구분자)
    patterns = [
        r'(202[0-9]|203[0-9])[\s년\-./\\]+(\d{1,2})[\s월\-./\\]+(\d{1,2})',
        r'(202[0-9]|203[0-9])[-.\s/](\d{1,2})[-.\s/](\d{1,2})',
        r'(202[0-9]|203[0-9])(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])'
    ]

    for pat in patterns:
        matches = re.findall(pat, full_text)
        if matches:
            for mat in matches:
                y, month, d = mat[0], int(mat[1]), int(mat[2])
                if 1 <= month <= 12 and 1 <= d <= 31:
                    res = f"{y}-{month:02d}-{d:02d}"
                    _log_debug(f"Match Step 2 Pattern SUCCESS: {res}")
                    return res

    # 3. Smart Fallback Engine: 문서 전체에서 2020~2030 날짜 추출
    fallback_matches = re.findall(r'(202[0-9]|203[0-9])[\s년\-./\\]*(\d{1,2})[\s월\-./\\]*(\d{1,2})', full_text)
    if fallback_matches:
        for mat in fallback_matches:
            y, month, d = mat[0], int(mat[1]), int(mat[2])
            if 1 <= month <= 12 and 1 <= d <= 31:
                res = f"{y}-{month:02d}-{d:02d}"
                _log_debug(f"Smart Fallback SUCCESS: {res}")
                return res

    _log_debug("No date match found in full_text!")
    return ''

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
