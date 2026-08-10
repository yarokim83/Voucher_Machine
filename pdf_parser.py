import re
import pdfplumber

def parse_pr_pdf(pdf_path):
    """
    HPNT PURCHASE REQUISITION 및 일반 PR PDF 파싱
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

    # 2. Date (예: 2026-07-27)
    date_match = re.search(r'DATE\s*(\d{4}-\d{2}-\d{2})', full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', full_text)
    if date_match:
        parsed_data['date'] = date_match.group(1).strip()

    # 3. Subject / PR Title (예: ARMGC-251호 Hoist Wire Rope 교환작업)
    subj_match = re.search(r'(?:SUBJECT|⊙\s*SUBJECT)\s*\n?([^\n]+)', full_text)
    if subj_match:
        title = subj_match.group(1).strip()
        title = re.sub(r'^(⊙\s*TYPE OF PURCHASE|TYPE OF PURCHASE).*', '', title).strip()
        parsed_data['pr_title'] = title
    else:
        item_match = re.search(r'([A-Za-z0-9\-]+호\s+[^\n]+교환작업|EXCHANGE WIRE ROPE[^\n]*)', full_text)
        if item_match:
            parsed_data['pr_title'] = item_match.group(1).strip()

    # 4. Total Amount & VAT
    amt_matches = re.findall(r'(?:KRW|총약정금액|Total)?\s*([\d,]{4,})', full_text)
    valid_amounts = []
    for m in amt_matches:
        num_str = m.replace(',', '')
        if num_str.isdigit():
            val = int(num_str)
            if val >= 1000:
                valid_amounts.append(val)

    if valid_amounts:
        base_amt = max(valid_amounts)
        parsed_data['amount'] = base_amt
        parsed_data['vat'] = int(base_amt * 0.1)
        parsed_data['total_amount'] = base_amt + parsed_data['vat']

    # 5. Company / Supplier
    supp_match = re.search(r'SUPPLIERS RECOMMENDED[\s\S]*?\d+\s+([가-힣A-Za-z0-9\((\)]+)', full_text)
    if supp_match:
        parsed_data['supplier'] = supp_match.group(1).strip()

    # 6. Prepared By
    prep_match = re.search(r'PREPARED BY[\s\S]*?NAME\s+([가-힣]{2,4})', full_text)
    if prep_match:
        parsed_data['prepared_by'] = prep_match.group(1).strip()

    return parsed_data

def parse_pdf_generic(pdf_path):
    return parse_pr_pdf(pdf_path)
