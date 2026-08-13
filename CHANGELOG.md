# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.0.0] - 2026-08-13
### Major Release: Green Progressbar, Sliced Contract, Ctrl+Shift+V Hotkey & Native Printer Fix
- **1. 글로벌 전역 단축키 Ctrl+Shift+V 변경**:
  어느 화면에서든 Ctrl+Shift+V (또는 Ctrl+Alt+V)를 누르면 마우스 포인터 커서 바로 옆 위치로 위젯 뿅! 팝업.
- **2. 상단 프로그레스 막대바(Progress Bar) 지능적 채우기**:
  서류가 1개씩 업로드될 때마다 초록색 프로그레스 막대바가 0% -> 20% -> 40% -> 60% -> 80% -> 100% 로 부드럽게 채워짐.
- **3. 업체 계약서 특정 범위 페이지 지정 추출 & 저장/인쇄 지원**:
  ⑤ 업체 계약서 카드에 📄 페이지: 입력칸(예: 12-13) 탑재 및 💾 추출저장 버튼 추가! 12-13 입력 시 2페이지 단독 PDF 추출 저장 및 해당 페이지 범위만 콕 찍어 스마트 인쇄.
- **4. 서류 5종 일괄 인쇄 100% 안전 무중단 파이프라인 보강**:
  SumatraPDF 및 win32api/os.startfile 인쇄 엔진을 안전하게 개별 래핑하여 PDF 서류들이 프린터로 100% 연속 출력되도록 수정.
- **버전 동기화 커스텀 룰(Version Sync Rule) 적용**:
  ersion.txt, CHANGELOG.md, pp_gui.py 뱃지를 8.0.0 로 100% 완벽 동기화.
