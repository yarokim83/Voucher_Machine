# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.1.2] - 2026-08-18
### Patch Release: Fix HTML Password Focus Timing & Add Step-by-Step Auto Unlock Logging
- **웹브라우저 포커스 대기 시간 1.2초 확대 패치**:
  HTML 세금계산서 열기 후 암호 입력 박스로의 포커스 이동 딜레이(1.2초)를 보장하여 비밀번호(6068625399)가 100% 붙여넣어지도록 안정화.
- **인쇄 및 로딩 딜레이 타이밍 보강**:
  비밀번호 입력 후 본문 로딩(1.0초) 및 PDF 인쇄(0.8초) 딜레이 조율.
- **자동 해제 단계별 상세 디버그 로깅 수집**:
  Step 1 (HTML열기) -> Step 2 (비밀번호 입력) -> Step 3 (인쇄 명령) -> Step 4 (PDF 저장 감지) 전체 4단계를 oucher_pass_debug.log에 실시간 기록.

## [v8.1.1] - 2026-08-18
### Patch Release: Fix HTML Security Notice Date Misparsing (2017)
