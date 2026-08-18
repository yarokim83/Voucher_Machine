# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.3.1] - 2026-08-18
### Feature Release: Exact 4-Step Unlock Sequence (Password -> Print Preview -> PDF Switch -> Save)
- **사용자 요청 4단계 파이프라인 정밀 구현**:
  1. 비밀번호 입력 및 제출 (Ctrl+V -> Enter)
  2. 계산서 인쇄 미리보기 화면 호출 (Ctrl+P)
  3. 미리보기 옵션에서 대상을 'PDF로 저장'으로 변경 (Tab -> Enter -> Down -> Enter)
  4. 인쇄/저장 버튼 클릭 및 파일 저장 확정 (Enter -> Enter)

## [v8.3.0] - 2026-08-18
### Major Reliability Release: Seamless System-Level PDF Printer Switching & Auto Revert
