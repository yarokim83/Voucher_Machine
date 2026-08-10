# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.6.2] - 2026-08-10
### Fixed
- Windows GDI Device Context (win32ui/win32print) 직통 비트맵 프린터 드라이버 스풀러 엔진 탑재 (printer_handler.py)로 외부 프로그램 연동 및 ShellExecute 31 (시스템에 부착된 장치가 작동하지 않습니다) 오류 100% 영구 해결

## [v1.6.1] - 2026-08-10
### Fixed
- ShellExecute Error 31 오류 회피용 Edge/PowerShell 무음 인쇄 다중 엔진 구축

## [v1.6.0] - 2026-08-10
### Changed
- 일괄 인쇄 기능 범위 변경: 업로드한 4가지 PDF 서류(PR Print, 거래명세서, 세금계산서, 계약서 지정 페이지)만 순서대로 즉시 일괄 인쇄

## [v1.5.0] - 2026-08-10
### Added
- 2번 추출 데이터 섹션 내 핵심 3개 항목(**작성일자**, **PR Title**, **공급가액**) 원클릭 클립보드 복사 버튼 ([📋 복사]) 구현

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
