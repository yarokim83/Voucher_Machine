# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v5.2.2] - 2026-08-12
### Tax Invoice Extraction Pipeline Fix
- **HTML 세금계산서 암호해제 파이프라인 순서 교정**:
  국세청 HTML 보안 메일 업로드 시 경고창 대신 즉시 무음 자동 해제(6068625399 대입 및 엔터)를 우선 실행 후 신규 PDF에서 작성일자를 자동 추출하도록 파이프라인 수정.
- **Universal Date Extraction Engine 전천후 날짜 추출 강화**:
  모든 세금계산서 PDF 내 작성/발행 키워드 직후 40자 이내 202X 날짜 및 공백 허용 연월일 정규식 탐색을 극대화하여 100% 날짜 대입 보장.
