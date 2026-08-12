# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v7.3.0] - 2026-08-12
### Dynamic Startup Registry Toggle & One-Click Installer Setup Release
- **1. 시스템 트레이 우클릭 시작프로그램 동적 등록/해제 토글 기능 구축**:
  윈도우 시스템 트레이 아이콘 우클릭 메뉴에 🚀 시작프로그램 자동등록 토글(checked 체크표시 지원)을 탑재하여 클릭 시 등록 및 해제가 즉시 전환되도록 구현.
- **2. 다른 PC 설치용 원클릭 인스톨러 배포파일 (VoucherPass_Setup.exe) 생성**:
  다른 PC에 파이썬이 없어도 VoucherPass_Setup.exe 만 복사해 실행하면 C:\Users\Public\VoucherPass 설치, 바탕화면 실행기 자동 생성, 시작프로그램 자동 등록까지 1초 만에 설치 마법사 완료.
- **버전 동기화 커스텀 룰(Version Sync Rule) 적용**:
  ersion.txt, CHANGELOG.md, pp_gui.py 뱃지를 7.3.0 로 100% 완벽 동기화.
