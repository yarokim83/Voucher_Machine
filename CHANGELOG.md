# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.4.3] - 2026-08-18
### Patch Release: 국세청 세금계산서 iframe 스크롤바 원천 제거 및 높이 자동 확장
- **우측 스크롤바 및 하단부 잘림 현상 완벽 해결**:
  국세청 HTML 내부의 세금계산서 본문 `iframe(CriMsgPosition)`의 고정 높이 제약을 해제하고 내부 전체 문서 높이에 맞춰 자동 확장(`overflow: hidden`) 처리하여, 우측 스크롤바가 찍히거나 하단 합계금액이 잘리는 현상을 100% 원천 해결.

## [v8.4.2] - 2026-08-18
### Patch Release: 세금계산서 PDF 하단 잘림 방지 및 A4 최적화
- **PDF 변환 시 세금계산서 전체 레이아웃 100% 온전하게 보존**:
  상단 '인쇄/첨부보기' 툴바 및 비밀번호 창 숨김 처리, `@media print` 에뮬레이션, `scale: 0.95` 적용으로 하단 합계금액 표 및 국세청 안내 문구까지 A4 1페이지에 완벽하게 렌더링되도록 수정.

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
