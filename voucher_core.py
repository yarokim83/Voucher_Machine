"""UI-independent document state, amount parsing, and page validation."""
import re
from decimal import Decimal
from dataclasses import dataclass, field


MONEY = r"(?<![\w.,])[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?![\w.,])"
LABELS = {
    'amount': r'공\s*급\s*가\s*액|공\s*급\s*금\s*액|Supply\s+Amount|Net\s+Amount|Subtotal',
    'vat': r'부\s*가\s*세|세\s*액|V\.?A\.?T\.?|Tax\s+Amount',
    'total_amount': r'합\s*계\s*금\s*액|총\s*금\s*액|총\s*약\s*정\s*금\s*액|발\s*주\s*금\s*액|Grand\s+Total|Total|합\s*계',
}
LABEL_PATTERN = re.compile('|'.join(f'(?P<{key}>{value})' for key, value in LABELS.items()), re.I)


def extract_amounts(text):
    """Read labeled values; never guess supply from the largest number or tax rate."""
    candidates = {key: set() for key in LABELS}
    lines = text.splitlines()
    for index, line in enumerate(lines):
        labels = list(LABEL_PATTERN.finditer(line))
        for pos, label in enumerate(labels):
            end = labels[pos + 1].start() if pos + 1 < len(labels) else len(line)
            value = re.fullmatch(r'\s*[:：=|]?\s*(?:KRW|₩)?\s*(' + MONEY + r')\s*(?:원|KRW)?\s*', line[label.end():end], re.I)
            if value:
                candidates[label.lastgroup].add(int(value.group(1).replace(',', '')))
        # A header-only row may be followed by a row containing exactly its values.
        if labels and not re.search(r'\d', line) and index + 1 < len(lines):
            row = lines[index + 1].strip()
            # NTS tables put the writing date before supply/tax in the value row.
            if '작성일자' in re.sub(r'\s+', '', line):
                row = re.sub(r'^20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}\s+', '', row)
            numbers = re.findall(MONEY, row)
            residue = re.sub(MONEY, '', row)
            if len(numbers) == len(labels) and not residue.strip(' \t|₩원'):
                for label, value in zip(labels, numbers):
                    candidates[label.lastgroup].add(int(value.replace(',', '')))
    result = {key: next(iter(values)) if len(values) == 1 else None for key, values in candidates.items()}
    warnings = [f'{key}: 서로 다른 금액 후보' for key, values in candidates.items() if len(values) > 1]
    # Derive only an absent value, and only from two unambiguous explicit values.
    if not warnings and sum(v is not None for v in result.values()) == 2:
        a, v, t = (result[k] for k in LABELS)
        if a is None:
            result['amount'] = t - v
        elif v is None:
            result['vat'] = t - a
        else:
            result['total_amount'] = a + v
    a, v, t = (result[k] for k in LABELS)
    if all(value is not None for value in (a, v, t)) and a + v != t:
        warnings.append('공급가액 + 세액 ≠ 합계금액')
    if any(value is None for value in (a, v, t)):
        warnings.append('금액 미추출: 원문 확인 필요')
    result['warnings'] = warnings
    result['explicit_fields'] = tuple(key for key, values in candidates.items() if len(values) == 1)
    return result


def extract_purchase_amounts(text, table_totals=None):
    """Read the last Total column of a KRW, VAT-exclusive purchase table.

    This rule is deliberately separate from invoice parsing: an exclusive Total
    is a net amount, and two monetary columns are unit-price and extended totals.
    VAT not printed in the document is a labeled calculation, never an extraction.
    """
    result = extract_amounts(text)
    if not re.search(r'PURCHASE\s+REQUISITION|PURCHAS(?:E|ING)\s+ORDER', text, re.I):
        return result
    section = re.search(r'CONTRACT\s+LIST|ORDER\s+LIST', text, re.I)
    if not section:
        return result
    # A PO may print the currency/tax basis immediately above ORDER LIST.
    line_start = text.rfind('\n', 0, section.start())
    start = text.rfind('\n', 0, max(0, line_start))
    table = text[max(0, start):]
    end = re.search(r'\n\s*(?:⊙|4\s*\.\s*계약\s*조건|TERMS\s+AND\s+CONDITIONS)', text[section.end():], re.I)
    if end:
        table = text[max(0, start):section.end() + end.start()]
    if not re.search(r'Unit\s*:\s*KRW\s*/\s*Excluding\s+VAT', table, re.I):
        return result
    lines = table.splitlines()
    totals = set()
    for index, line in enumerate(lines):
        total = re.match(r'^\s*Total\b(.*)$', line, re.I)
        if not total:
            continue
        # Currency and values can wrap onto separate lines in browser PDFs.
        values = []
        for row in [total[1]] + lines[index + 1:index + 5]:
            cleaned = re.sub(r'\bKRW\b|₩|원|\|', ' ', row, flags=re.I).strip()
            if not cleaned:
                continue
            tokens = cleaned.split()
            if len(tokens) in (1, 2) and all(re.fullmatch(r'[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.00)?', token) for token in tokens):
                values.extend(int(Decimal(token.replace(',', ''))) for token in tokens)
                if len(values) < 2:
                    continue
            # Never scan past a non-currency row into item/budget numbers.
            break
        if len(values) in (1, 2):
            totals.add(values[-1])
    if table_totals is not None:
        if not table_totals or any(value is None for value in table_totals):
            return dict(amount=None, vat=None, total_amount=None,
                        warnings=['구매 목록 Amount 합계 셀을 읽을 수 없습니다. 원문 확인 필요'])
        totals = set(table_totals)
    if not totals:
        return result
    if len(totals) != 1:
        return dict(amount=None, vat=None, total_amount=None,
                    warnings=['구매 목록에 서로 다른 Total 금액이 있습니다. 원문 확인 필요'])
    amount = totals.pop()
    if 'amount' in result['explicit_fields'] and result['amount'] != amount:
        return dict(amount=None, vat=None, total_amount=None,
                    warnings=['명시된 공급가액과 구매 목록 Total이 다릅니다. 원문 확인 필요'])
    # The template's Total excludes tax. Discard the generic gross-total label.
    vat = result['vat'] if 'vat' in result['explicit_fields'] else None
    data = dict(amount=amount, vat=vat, total_amount=None,
                amount_source='purchase_table_total_excluding_vat', warnings=[])
    if any(warning.startswith(('amount:', 'vat:')) for warning in result['warnings']):
        data['warnings'].append('명시된 금액에 서로 다른 후보가 있습니다. 원문 확인 필요')
        return data
    if vat is None:
        rates = {Decimal(rate) for rate in re.findall(r'\bVAT\s*(?:rate\s*)?[:=]?\s*(\d+(?:\.\d+)?)\s*%', text, re.I)}
        if len(rates) > 1 or any(rate < 0 or rate > 100 for rate in rates) or re.search(r'면세|영세율|VAT\s*exempt', text, re.I):
            data['warnings'].append('부가세율 확인 필요: 부가세를 직접 입력해 주세요.')
            return data
        rate = next(iter(rates)) if rates else Decimal('10')
        vat = int(Decimal(amount) * rate / 100)
        data.update(vat=vat, vat_source='calculated', vat_rate_percent=float(rate))
    else:
        data['vat_source'] = 'extracted'
    data['total_amount'] = amount + vat
    return data


def parse_money(value):
    value = value.replace(',', '').strip()
    if not value:
        return None
    if not re.fullmatch(r'[+-]?\d+', value):
        raise ValueError('금액은 정수로 입력해 주세요.')
    return int(value)


def parse_pages(value):
    value = value.strip() or '1'
    pages = set()
    for part in value.split(','):
        match = re.fullmatch(r'\s*([1-9]\d*)(?:\s*-\s*([1-9]\d*))?\s*', part)
        if not match:
            raise ValueError('페이지는 1, 3-5 형식의 양수로 입력해 주세요.')
        start, end = int(match[1]), int(match[2] or match[1])
        if end < start or end > 10000:
            raise ValueError('페이지 범위는 오름차순, 10000 이하로 입력해 주세요.')
        pages.update(range(start - 1, end))
    return sorted(pages), value


def validate_pages(pages, total):
    if not pages or any(type(p) is not int or p < 0 or p >= total for p in pages):
        raise ValueError(f'페이지 범위를 확인해 주세요. 문서는 총 {total}페이지입니다.')
    return list(pages)


@dataclass
class DocumentState:
    documents: dict = field(default_factory=dict)

    def replace(self, kind, data):
        self.documents[kind] = dict(data)

    def clear(self):
        self.documents.clear()

    def form(self):
        pr = self.documents.get('pr', {})
        po = self.documents.get('po', {})
        tax = self.documents.get('tax', {})
        # An attached PR owns its fields, even if parsing failed. Do not silently
        # substitute PO numbers/amounts into PR fields.
        result = {key: pr.get(key, '') for key in ('pr_no', 'pr_title')}
        owner = pr if 'pr' in self.documents else tax if 'tax' in self.documents else po
        for key in ('amount', 'vat', 'total_amount'):
            result[key] = owner.get('supply_amount' if owner is tax and key == 'amount' else key)
        result['supplier'] = po.get('supplier', '') or pr.get('supplier', '')
        result['date'] = tax.get('date', '') if 'tax' in self.documents else pr.get('date', '') or po.get('date', '')
        return result

    def supply(self, kind):
        return self.documents.get(kind, {}).get('supply_amount' if kind == 'tax' else 'amount')
