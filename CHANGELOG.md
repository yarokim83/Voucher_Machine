# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.3.0] - 2026-08-18
### Major Reliability Release: Seamless System-Level PDF Printer Switching & Auto Revert
- **시스템 기본 프린터 일시 전환(Microsoft Print to PDF) 및 자동 복원 아키텍처 도입**:
  키보드 탭/단축키 조작으로 인한 브라우저 검색창 튐(오작동)을 원천 차단하고, 윈도우 레벨에서 기본 프린터를 'Microsoft Print to PDF'로 일시 설정한 후 Ctrl+P ➔ Enter(저장) ➔ Enter(확정) 2단계로 100% 무결점 PDF 저장을 수행한 뒤 원래 프린터로 즉시 자동 복원.

## [v8.2.3] - 2026-08-18
### Patch Release: Bypass Ctrl+P Shortcut Lock via Alt+F -> P Browser Main Menu Command
