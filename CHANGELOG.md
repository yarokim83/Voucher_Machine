# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v5.3.0] - 2026-08-12
### Git History Original Tax Date Algorithm Restoration
- **Git Hub 히스토리 c7e501b 세금계산서 작성일자 파싱 알고리즘 100% 원본 복원**:
  과거 성공적으로 동작하던 원본 parse_tax_invoice_date 정규식 복원. 국세청 전자세금계산서 표준 양식에서 작성일자 헤더 아래의 연월일(2026/08/11, 2026.08.11, 2026-08-11, 2026년 08월 11일)을 100% 파싱하도록 복구.
