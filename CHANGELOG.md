# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.1.1] - 2026-08-18
### Patch Release: Fix HTML Security Notice Date Misparsing (2017)
- **국세청 보안 메일 HTML 전역 탐색 예외 처리 패치**:
  암호화된 HTML 보안 메일 원본에 포함된 ActiveX/Internet Explorer 7 안내 문구(2017년 10월 26일)가 잘못 추출되는 오작동을 차단. HTML 파싱 시 세금계산서 전용 키워드('작성일자'/'발행일자') 부근 탐색으로 한정.

## [v8.1.0] - 2026-08-18
### Feature & Bug Fix Release: HTML Instant Date Extraction & Multi-Layer Debug Logging
