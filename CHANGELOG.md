# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v3.3.0] - 2026-08-12
### Layout Optimization & Workflow Refinement
- **드롭존 및 인쇄 순서 100% 동일화**:
  1. 전자 세금계산서 -> 2. 거래명세서 -> 3. PR Print (구매요청서) -> 4. 발주서 (PO) -> 5. 업체 계약서 순서로 카드 및 5종 일괄 인쇄 순서 변경.
- **[📋 Voucher 엑셀 양식 붙여넣기 클립보드 복사] 버튼 개편**:
  버튼 명칭 변경 및 가독성 높은 메인 버튼으로 전환.
- **엑셀 템플릿 저장 버튼 UI 숨김**:
  추후 재활용 가능하도록 백엔드 기능(generate_excel_action 및 excel_handler.py)을 100% 보존한 상태에서 UI 버튼만 숨겨 깔끔한 화면 구성.
- **430 x 640 콤팩트 핏 HUD 디자인 적용**:
  하단 낭비 여백을 완전히 제거하여 한눈에 모든 기능이 들어오도록 최적화.
