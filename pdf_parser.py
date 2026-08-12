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

def parse_tax_invoice_date(file_path):
    if not os.path.exists(file_path):
        return ''

    if _is_html_file(file_path):
        full_text = _read_html_text(file_path)
    else:
        full_text = extract_pdf_text_safely(file_path)

    m = re.search(r'작성일자?\s*[:\s]*(\d{4})[년\-.\s/]\s*(\d{1,2})[월\-.\s/]\s*(\d{1,2})[일\s]?', full_text)
    if m and int(m.group(1)) >= 2020:
        y, month, d = m.group(1), int(m.group(2)), int(m.group(3))
        return f"{y}-{month:02d}-{d:02d}"

    m = re.search(r'발[급행]일자?\s*[:\s]*(\d{4})[년\-.\s/]\s*(\d{1,2})[월\-.\s/]\s*(\d{1,2})[일\s]?', full_text)
    if m and int(m.group(1)) >= 2020:
        y, month, d = m.group(1), int(m.group(2)), int(m.group(3))
        return f"{y}-{month:02d}-{d:02d}"

    matches = re.findall(r'(202[0-9])[년\-.\s/]\s*(\d{1,2})[월\-.\s/]\s*(\d{1,2})[일\s]?', full_text)
    if matches:
        y, month, d = matches[0][0], int(matches[0][1]), int(matches[0][2])
        return f"{y}-{month:02d}-{d:02d}"

    matches_fmt = re.findall(r'(202[0-9])[-.\s/](\d{1,2})[-.\s/](\d{1,2})', full_text)
    if matches_fmt:
        y, month, d = matches_fmt[0][0], int(matches_fmt[0][1]), int(matches_fmt[0][2])
        return f"{y}-{month:02d}-{d:02d}"

    matches_digits = re.findall(r'(202[0-9])(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])', full_text)
    if matches_digits:
        y, month, d = matches_digits[0][0], int(matches_digits[0][1]), int(matches_digits[0][2])
        return f"{y}-{month:02d}-{d:02d}"

    return ''

def parse_pr_pdf(pdf_path):
    """
    PR Print (구매요청서) 및 발주서(PO/Purchase Order) PDF 데이터 파싱
    """
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

    # 1. P/R No. 또는 발주번호 (PO No.) 추출
    pr_match = re.search(r'(?:P/R|PR|PO|발주)\s*No\.?\s*[:\s]*([A-Z0-9]+)', full_text, re.IGNORECASE)
    if not pr_match:
        pr_match = re.search(r'발주번호\s*[:\s]*([A-Z0-9]+)', full_text)
    if not pr_match:
        pr_match = re.search(r'S202[0-9]{8}', full_text)
    
    if pr_match:
        extracted_data['pr_no'] = pr_match.group(1) if len(pr_match.groups()) > 0 else pr_match.group(0)

    # 2. 날짜 추출 (DATE / 작성일자 / 발주일자)
    date_match = re.search(r'(?:DATE|작성일자|발주일자|발행일자)\s*[:\s]*(\d{4}[-.\s]\d{2}[-.\s]\d{2})', full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', full_text)
    
    if date_match:
        extracted_data['date'] = date_match.group(1).replace('.', '-').strip()

    # 3. Subject / PR Title / 발명/공사명 추출
    subj_match = re.search(r'(?:SUBJECT|품명|공사명|건명|발주명)\s*[:\s]*(.+)', full_text, re.IGNORECASE)
    if subj_match:
        extracted_data['pr_title'] = subj_match.group(1).strip()

    # 4. 금액 (공급가액 / Total Amount / 발주금액) 파싱
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

    # 5. 거래처명 (Supplier / Payee / 수신/공급자)
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
    """PDF 또는 HTML 파일 내용/파일명으로 4개 카드 타입 분류 (PR / 발주서 통합)"""
    if not os.path.exists(file_path):
        return 'unknown'

    fname = os.path.basename(file_path).lower()
    if 'nts_etaxinvoice' in fname or 'etaxinvoice' in fname:
        return 'tax'

    if '구매요청' in fname or 'pr' in fname or 'requisition' in fname or '발주' in fname or 'po' in fname or 'order' in fname:
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
    if 'PURCHASE REQUISITION' in txt_upper or 'PURCHASE ORDER' in txt_upper or '구매요청서' in txt or '발주서' in txt or 'P/R NO' in txt_upper or 'P.O' in txt_upper:
        return 'pr'
    if '전자세금계산서' in txt or '세금계산서' in txt or 'TAX INVOICE' in txt_upper:
        return 'tax'
    if '거래명세서' in txt or '명세서' in txt:
        return 'spec'
    if '계약서' in txt or '협약서' in txt or 'CONTRACT' in txt_upper:
        return 'contract'

    return 'pr'
