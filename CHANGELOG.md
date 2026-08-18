# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.3.6] - 2026-08-18
### Patch Release: Windows 시스템 인쇄 대화상자로 PDF 프린터 자동 선택
- **Edge 인쇄 미리보기 대신 Windows 시스템 인쇄 대화상자 사용 (Ctrl+Shift+P)**:
  Edge 자체 인쇄 미리보기에서 Tab 키로 대상 드롭다운을 탐색하는 방식이 불안정하여, Windows 시스템 인쇄 대화상자를 직접 호출하여 사전에 설정한 기본 프린터 'Microsoft Print to PDF'가 자동 선택되도록 변경.

## [v8.3.5] - 2026-08-18
### Patch Release: Clean Single Paste for Password 6068625399 & Clipboard Isolation
- **비밀번호 단일 클립보드 붙여넣기 및 중복 입력 방지**:
  키 타이핑과 붙여넣기가 중복 실행되거나 클립보드 오염으로 인해 다른 텍스트가 섞여 들어가는 현상을 원천 방지하여, 오직 6068625399 단일 붙여넣기만 깔끔하게 실행되도록 격리.

## [v8.3.4] - 2026-08-18
### Patch Release: Explicit PDF File Path Input into Windows Save As Dialog
