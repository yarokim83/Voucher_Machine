# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.7.0] - 2026-08-10
### Added
- 초경량 무설치 **SumatraPDF Vector Native Print Engine (in/SumatraPDF.exe)** 탑재 (printer_handler.py)
- 비트맵 이미지 변환 방식 대신 원본 PDF 벡터(Vector) 폰트 명령어 직접 송출 방식으로 개편하여 **텍스트 뭉개짐/부엿함 없는 100% 칼같은 최고화질 인쇄** 구현
- -print-settings "fit" 옵션 적용으로 프린터 종이 인쇄 영역에 맞춘 100% 비율 자동 핏팅 (테두리 및 텍스트 잘림 현상 0% 해결)

## [v1.6.3] - 2026-08-10
### Fixed
- 300 DPI 초고해상도 비트맵 렌더링 도입

## [v1.6.2] - 2026-08-10
### Fixed
- Windows GDI Device Context 직통 비트맵 프린터 드라이버 스풀러 엔진 탑재

## [v1.6.0] - 2026-08-10
### Changed
- 일괄 인쇄 기능 범위 변경: 업로드한 4가지 PDF 서류만 순서대로 즉시 일괄 인쇄

## [v1.5.0] - 2026-08-10
### Added
- 2번 추출 데이터 섹션 내 핵심 3개 항목 원클릭 클립보드 복사 기능 구현

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
