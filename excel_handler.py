import os
import datetime
import openpyxl

DEFAULT_TEMPLATE_PATH = r'C:\Users\baewoong.kim\Desktop\고려제강(2025).xlsx'

def generate_voucher_excel(data, template_path=DEFAULT_TEMPLATE_PATH, output_dir=None):
    """
    Voucher 엑셀 템플릿에 파싱된 데이터를 기입하고 저장
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"엑셀 템플릿 파일을 찾을 수 없습니다: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 데이터 기입
    # 1. Date (A6)
    date_val = data.get('date', '')
    if date_val:
        try:
            dt = datetime.datetime.strptime(date_val, '%Y-%m-%d')
            ws['A6'] = dt
            ws['D22'] = dt
        except ValueError:
            ws['A6'] = date_val

    # 2. Payee (J6)
    if data.get('supplier'):
        ws['J6'] = data.get('supplier')

    # 3. P/R No. (W6)
    if data.get('pr_no'):
        ws['W6'] = data.get('pr_no')

    # 4. Description (D9)
    if data.get('pr_title'):
        ws['D9'] = data.get('pr_title')

    # 5. Amount (W9) & VAT (W16) & Total (W17)
    amt = data.get('amount', 0)
    vat = data.get('vat', 0)
    total = data.get('total_amount', 0)

    if amt:
        ws['W9'] = amt
        ws['P32'] = amt
    if vat:
        ws['W16'] = vat
        ws['P33'] = vat
    if total:
        ws['W17'] = total
        ws['V34'] = total

    # 파일 저장
    if not output_dir:
        output_dir = os.path.expanduser('~/Desktop')

    pr_no_str = data.get('pr_no', 'OUTPUT')
    filename = f"Voucher_{pr_no_str}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = os.path.join(output_dir, filename)

    wb.save(output_path)
    wb.close()
    return output_path

if __name__ == '__main__':
    print("excel_handler loaded")
