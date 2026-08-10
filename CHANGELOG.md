# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.9.0] - 2026-08-10
### Added
- 전자 세금계산서 PDF 업로드 시 작성일자(발행일자) 정밀 파싱 및 date_var 자동 연동 기능 구현 (pdf_parser.py, pp_gui.py)
- 사용자 지정 공식 VoucherPass 로고 아이콘(ssets/app_icon.png, ssets/app_icon.ico) 적용 (프로그램 창, 헤더 및 작업표시줄 아이콘)
- PyInstaller 독립 패키징 배포 빌드 완료 (dist/VoucherPass/VoucherPass.exe) 및 바탕화면 실행 아이콘(VoucherPass 실행.bat) 동기화

## [v1.8.0] - 2026-08-10
### Added
- 계약서 페이지 지정 범위 파싱 지원 강화 (12-13)
- 3종 항목 세로 줄바꿈(
) 구분 복사 적용
- 좌측 사이드바 간소화

## [v1.7.0] - 2026-08-10
### Added
- 초경량 무설치 **SumatraPDF Vector Native Print Engine (in/SumatraPDF.exe)** 탑재

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
