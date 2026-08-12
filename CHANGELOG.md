# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v3.1.0] - 2026-08-12
### Major Features & Cleanup
- **PR Print(구매요청서) & 발주서(PO) 드롭존 독립 분리**:
  1번 슬롯 ① PR Print (구매요청서) PDF 및 2번 슬롯 ② 발주서 (PO) PDF 로 100% 개별 수집 및 파싱 지원.
- **제출 서류 5종 일괄 인쇄 파이프라인 탑재**:
  업로드된 서류 5가지(① PR, ② 발주서, ③ 거래명세서, ④ 암호해제 세금계산서, ⑤ 지정페이지 계약서)를 원클릭 무음 일괄 인쇄.
- **레가시 아웃룩 코드 100% 제거**:
  사용하지 않는 [📧 아웃룩 1초] 버튼 및 outlook_handler.py 모듈 완전 삭제로 어플리케이션 경량화.
