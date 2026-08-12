# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v6.0.0] - 2026-08-12
### Major Release: Clean PDF Folder Watcher Architecture & Code Refactoring
- **불필요 웹 자동화 모듈 100% 완전 삭제 & 프로젝트 코드 대대적 경량화**:
  속도가 느리고 불안정했던 웹브라우저/PyAutoGUI 자동화 코드를 100% 제거하고, 슬림하고 빠른 클린 파이프라인으로 전환.
- **지정 폴더(다운로드/바탕화면) PDF 자동 감지 & 작성일자 자동 파싱 엔진**:
  사용자가 세금계산서를 지정 폴더에 PDF 로 저장하는 즉시 폴더 감시기(FolderWatcher)가 세금계산서 PDF임을 자동 분류하여 1번 슬롯에 장착하고, 작성일자(2026/08/11)를 100% 자동으로 추출하여 📅 작성일자 칸 및 7종 HUD에 대입.
- **버전 동기화 커스텀 룰(Version Sync Rule) 적용**:
  ersion.txt, CHANGELOG.md, pp_gui.py 뱃지를 6.0.0 로 100% 완벽 동기화.
