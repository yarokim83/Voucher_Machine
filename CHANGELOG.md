# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.4.1] - 2026-08-18
### Patch Release: PyInstaller Selenium 전체 모듈 번들 포함
- **PyInstaller 빌드 시 `--collect-all selenium` 적용**:
  `No module named 'selenium.webdriver.edge.webdriver'` 번들 누락 오류를 완벽히 해결하여 headless PDF 변환이 정상 실행되도록 수정.

## [v8.4.0] - 2026-08-18
### Feature Release: Selenium Headless PDF 변환 (브라우저/해상도 독립)
- **pyautogui 키보드 조작 완전 제거, Selenium headless Edge + CDP `Page.printToPDF` 로 전환**:
  브라우저 종류(Edge/Chrome/Whale), 화면 해상도, 프린터 설정에 완전히 독립적으로 동작.
  화면에 브라우저 창이 열리지 않으며, 백그라운드에서 비밀번호 입력 → JS 복호화 → PDF 직접 생성.


### Patch Release: Windows 시스템 인쇄 대화상자로 PDF 프린터 자동 선택
- **Edge 인쇄 미리보기 대신 Windows 시스템 인쇄 대화상자 사용 (Ctrl+Shift+P)**:
  Edge 자체 인쇄 미리보기에서 Tab 키로 대상 드롭다운을 탐색하는 방식이 불안정하여, Windows 시스템 인쇄 대화상자를 직접 호출하여 사전에 설정한 기본 프린터 'Microsoft Print to PDF'가 자동 선택되도록 변경.

## [v8.3.5] - 2026-08-18
### Patch Release: Clean Single Paste for Password 6068625399 & Clipboard Isolation
- **비밀번호 단일 클립보드 붙여넣기 및 중복 입력 방지**:
  키 타이핑과 붙여넣기가 중복 실행되거나 클립보드 오염으로 인해 다른 텍스트가 섞여 들어가는 현상을 원천 방지하여, 오직 6068625399 단일 붙여넣기만 깔끔하게 실행되도록 격리.

## [v8.3.4] - 2026-08-18
### Patch Release: Explicit PDF File Path Input into Windows Save As Dialog
