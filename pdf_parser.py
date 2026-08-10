import pdfplumber
import re
import os

def parse_pr_pdf(pdf_path):
    """
    PR Print (구매요청서) PDF 데이터 파싱
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

    full_text = ""
    tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                full_text += txt + "\n"
            tbl = page.extract_tables()
            if tbl:
                tables.extend(tbl)

    # 1. P/R No. 추출
    pr_match = re.search(r'P/R\s*No\.?\s*[:\s]*([A-Z0-9]+)', full_text, re.IGNORECASE)
    if not pr_match:
        pr_match = re.search(r'PR\s*No\.?\s*[:\s]*([A-Z0-9]+)', full_text, re.IGNORECASE)
    if not pr_match:
        pr_match = re.search(r'S202[0-9]{8}', full_text)
    
    if pr_match:
        extracted_data['pr_no'] = pr_match.group(1) if len(pr_match.groups()) > 0 else pr_match.group(0)

    # 2. 날짜 추출 (YYYY-MM-DD 또는 YYYY.MM.DD)
    date_match = re.search(r'DATE\s*[:\s]*(\d{4}[-.\s]\d{2}[-.\s]\d{2})', full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r'작성일자?\s*[:\s]*(\d{4}[-.\s]\d{2}[-.\s]\d{2})', full_text)
    if not date_match:
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', full_text)
    
    if date_match:
        extracted_data['date'] = date_match.group(1).replace('.', '-').strip()

    # 3. Subject / PR Title 추출
    subj_match = re.search(r'SUBJECT\s*[:\s]*(.+)', full_text, re.IGNORECASE)
    if subj_match:
        extracted_data['pr_title'] = subj_match.group(1).strip()

    # 4. 금액 (공급가액 / Total Amount) 파싱
    pr_no_str = extracted_data['pr_no']
    candidates = []
    
    lines = full_text.split('\n')
    for line in lines:
        if pr_no_str and pr_no_str in line:
            continue
        
        amounts = re.findall(r'(?:KRW|₩|총금액|공급가액|총약정금액|Total)?\s*([1-9]\d{0,2}(?:,\d{3})+)', line)
        for amt_str in amounts:
            clean_num = int(amt_str.replace(',', ''))
            if pr_no_str and str(clean_num) in pr_no_str:
                continue
            if len(str(clean_num)) >= 10:
                continue
            if 1000 <= clean_num <= 1000000000:
                is_priority = any(k in line for k in ['KRW', '총약정금액', 'Total', '공급가액', '합계'])
                candidates.append((clean_num, is_priority))

    priority_candidates = [c[0] for c in candidates if c[1]]
    if priority_candidates:
        extracted_data['amount'] = max(priority_candidates)
    elif candidates:
        extracted_data['amount'] = max([c[0] for c in candidates])

    if extracted_data['amount'] > 0:
        extracted_data['vat'] = int(extracted_data['amount'] * 0.1)
        extracted_data['total_amount'] = extracted_data['amount'] + extracted_data['vat']

    # 5. 거래처명 (Supplier / Payee)
    sup_match = re.search(r'Company\s+([가-힣A-Za-z0-9㈜(주)]+)', full_text)
    if not sup_match:
        sup_match = re.search(r'SUPPLIERS?\s*RECOMMENDED[\s\S]*?1\s+([가-힣A-Za-z0-9㈜(주)]+)', full_text)
    if not sup_match:
        sup_match = re.search(r'금강엔지니어링', full_text)
    
    if sup_match:
        extracted_data['supplier'] = sup_match.group(1) if len(sup_match.groups()) > 0 else sup_match.group(0)

    return extracted_data

def parse_tax_invoice_date(pdf_path):
    """
    전자 세금계산서 PDF에서 작성일자/발행일자 정밀 추출
    """
    if not os.path.exists(pdf_path):
        return ""

    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                full_text += txt + "\n"

    # 패턴 1: 작성일자 YYYY년 MM월 DD일 또는 YYYY-MM-DD
    m = re.search(r'작성일자?\s*[:\s]*(\d{4})[년\-.\s/]\s*(\d{1,2})[월\-.\s/]\s*(\d{1,2})[일\s]?', full_text)
    if m:
        y, month, d = m.group(1), int(m.group(2)), int(m.group(3))
        return f"{y}-{month:02d}-{d:02d}"

    # 패턴 2: 발행일자 YYYY-MM-DD
    m = re.search(r'발행일자?\s*[:\s]*(\d{4})[년\-.\s/]\s*(\d{1,2})[월\-.\s/]\s*(\d{1,2})[일\s]?', full_text)
    if m:
        y, month, d = m.group(1), int(m.group(2)), int(m.group(3))
        return f"{y}-{month:02d}-{d:02d}"

    # 패턴 3: 작성일자 8자리 숫자 (20260727)
    m = re.search(r'작성일자?\s*[:\s]*(\d{4})(\d{2})(\d{2})', full_text)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # 패턴 4: 일반 YYYY-MM-DD 날짜 추출
    m = re.search(r'(\d{4}[-.\s/]\d{2}[-.\s/]\d{2})', full_text)
    if m:
        cleaned = m.group(1).replace('.', '-').replace('/', '-').strip()
        parts = cleaned.split('-')
        if len(parts) == 3:
            return f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"

    return ""

if __name__ == '__main__':
    print("pdf_parser loaded with tax invoice date extractor")
