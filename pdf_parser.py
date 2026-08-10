import re
import pdfplumber

def parse_pr_pdf(pdf_path):
    """
    HPNT PURCHASE REQUISITION, 세금계산서, 거래명세서 등 PDF 파싱
    """
    parsed_data = {
        'pr_no': '',
        'pr_title': '',
        'amount': 0,
        'vat': 0,
        'total_amount': 0,
        'date': '',
        'supplier': '',
        'prepared_by': '',
        'doc_type': 'PR'
    }
    
    full_text = ''
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                full_text += (page.extract_text() or '') + '\n'
    except Exception as e:
        print(f"pdfplumber read error: {e}")

    # 1. PR No. (예: S202607270003)
    pr_match = re.search(r'PR\s*No\.?\s*([A-Z0-9]+)', full_text, re.IGNORECASE)
    if pr_match:
        parsed_data['pr_no'] = pr_match.group(1).strip()

    # 2. Date (예: 2026-07-27 또는 2026/07/30)
    date_match = re.search(r'DATE\s*(\d{4}[-/\.]\d{2}[-/\.]\d{2})', full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r'(\d{4}[-/\.]\d{2}[-/\.]\d{2})', full_text)
    if date_match:
        parsed_data['date'] = date_match.group(1).replace('/', '-').replace('.', '-').strip()

    # 3. Subject / PR Title
    subj_match = re.search(r'(?:SUBJECT|⊙\s*SUBJECT)\s*\n?([^\n]+)', full_text)
    if subj_match:
        title = subj_match.group(1).strip()
        title = re.sub(r'^(⊙\s*TYPE OF PURCHASE|TYPE OF PURCHASE).*', '', title).strip()
        parsed_data['pr_title'] = title
    else:
        item_match = re.search(r'([A-Za-z0-9\-]+호\s+[^\n]+교환작업|[A-Za-z0-9\-]+호\s+Hoist wire 교체|EXCHANGE WIRE ROPE[^\n]*)', full_text)
        if item_match:
            parsed_data['pr_title'] = item_match.group(1).strip()

    # 4. Total Amount & VAT 정밀 추출 (오류 방지)
    # PR No(예: 202607270003), 승인번호, 사업자등록번호(541-86-01824) 등 10자리 이상의 일련번호 제거
    clean_text_for_amt = full_text
    if parsed_data['pr_no']:
        clean_text_for_amt = clean_text_for_amt.replace(parsed_data['pr_no'], '')
    
    # 10자리 이상 연속 숫자는 PR No나 승인번호 가능성이 크므로 마스킹
    clean_text_for_amt = re.sub(r'\b\d{10,}\b', '', clean_text_for_amt)

    target_amount = 0

    # 키워드 기반 금액 추출 (우선순위 1)
    # 예: Total KRW 2,100,000 / Total 2,100,000 / 총약정금액 KRW 2,100,000 / 공급가액 2,100,000
    kw_match = re.search(r'(?:Total|총약정금액|공급가액|총액)\s*(?:KRW|₩)?\s*([\d,]{4,})', clean_text_for_amt, re.IGNORECASE)
    if kw_match:
        raw_val = kw_match.group(1).replace(',', '')
        if raw_val.isdigit():
            target_amount = int(raw_val)

    # 우선순위 2: "KRW 2,100,000" 형태 직접 매칭
    if not target_amount:
        krw_match = re.search(r'KRW\s*([\d,]{4,})', clean_text_for_amt, re.IGNORECASE)
        if krw_match:
            raw_val = krw_match.group(1).replace(',', '')
            if raw_val.isdigit():
                target_amount = int(raw_val)

    # 우선순위 3: 테이블 등에서 숫자의 나열 중 합리적인 금액 (1,000원 ~ 10,000,000,000원 범위)
    if not target_amount:
        amt_candidates = re.findall(r'\b([\d,]{4,})\b', clean_text_for_amt)
        valid_amounts = []
        for cand in amt_candidates:
            val_str = cand.replace(',', '')
            if val_str.isdigit():
                val = int(val_str)
                # 1,000원 이상 10억 미만의 현실적 금액만 채택 (PR No 등이 실수로 포함되는 것 방지)
                if 1000 <= val < 1000000000:
                    valid_amounts.append(val)
        if valid_amounts:
            target_amount = max(valid_amounts)

    if target_amount:
        parsed_data['amount'] = target_amount
        parsed_data['vat'] = int(target_amount * 0.1)
        parsed_data['total_amount'] = target_amount + parsed_data['vat']

    # 5. Company / Supplier
    supp_match = re.search(r'SUPPLIERS RECOMMENDED[\s\S]*?\d+\s+([가-힣A-Za-z0-9\((\)]+)', full_text)
    if not supp_match:
        supp_match = re.search(r'공급자[\s\S]*?상호\s*\(법인명\)\s*([가-힣A-Za-z0-9\(\)]+)', full_text)
    if supp_match:
        parsed_data['supplier'] = supp_match.group(1).strip()

    # 6. Prepared By
    prep_match = re.search(r'PREPARED BY[\s\S]*?NAME\s+([가-힣]{2,4})', full_text)
    if prep_match:
        parsed_data['prepared_by'] = prep_match.group(1).strip()

    return parsed_data

def parse_pdf_generic(pdf_path):
    return parse_pr_pdf(pdf_path)
