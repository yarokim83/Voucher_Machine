# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.9.1] - 2026-08-10
### Fixed
- PyInstaller --hidden-import 옵션에 pdfplumber, pypdf, PIL, openpyxl, 	kinterdnd2 모듈 추가하여 No module named 'pdfplumber' 구동 예외 완전 해결

## [v1.9.0] - 2026-08-10
### Added
- 전자 세금계산서 PDF 업로드 시 작성일자(발행일자) 정밀 파싱 및 date_var 자동 연동 기능 구현
- 사용자 지정 공식 VoucherPass 로고 아이콘 적용
- PyInstaller 독립 패키징 배포 빌드 완료 (dist/VoucherPass/VoucherPass.exe)

## [v1.8.0] - 2026-08-10
### Added
- 계약서 페이지 지정 범위 파싱 지원 강화 (12-13)
- 3종 항목 세로 줄바꿈(\n) 구분 복사 적용
- 좌측 사이드바 간소화

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
