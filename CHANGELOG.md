# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.1.0] - 2026-08-18
### Feature & Bug Fix Release: HTML Instant Date Extraction & Multi-Layer Debug Logging
- **1. HTML 세금계산서 작성일자 0.001초 즉시 파싱 구조 도입**:
  국세청 HTML 세금계산서 드롭 시 브라우저 자동 해제/PDF 인쇄 변환 대기 시간을 기다리지 않고 원본 HTML 텍스트에서 작성일자를 0.001초 만에 즉시 추출하여 화면 입력란에 적용.
- **2. 다층 디버그 로깅 시스템 (oucher_pass_debug.log) 구축**:
  세금계산서 업로드/파싱/정규식 4단계 탐색(년월일, YYYY-MM-DD, 8자리 YYYYMMDD, 2자리 연도 YY-MM-DD) 과정을 투명하게 기록하는 디버그 로그 파일 자동 생성.
- **3. 브라우저 PDF 변환 감지 타임아웃 및 폴더 확대**:
  실시간 PDF 감지 타임아웃을 10초(200회)로 확장하고 Desktop, Downloads, Documents, Temp 4개 폴더를 동시 모니터링하도록 보강.
- **4. 프린터 기종 기억 & 신규 커스텀 아이콘 반영**:
  마지막 사용 프린터 자동 로드 보장 및 ChatGPT Image 기반 신규 멀티 해상도 아이콘 적용.

## [v8.0.0] - 2026-08-13
### Major Release: Green Progressbar, Sliced Contract, Ctrl+Shift+V Hotkey & Native Printer Fix
