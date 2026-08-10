# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.6.3] - 2026-08-10
### Fixed
- 300 DPI 초고해상도 비트맵 렌더링 도입으로 텍스트 및 서식 흐릿함/선명도 저하 문제 완벽 해결 (printer_handler.py)
- 종이 가로세로 비율(Aspect Ratio) 100% 유지 + Safe Margin (4mm 안전 여백) 중앙 정렬 핏팅 알고리즘 적용으로 프린터 서식 및 가장자리 텍스트 잘림 현상 0%로 완벽 개편

## [v1.6.2] - 2026-08-10
### Fixed
- Windows GDI Device Context 직통 비트맵 프린터 드라이버 스풀러 엔진 탑재로 ShellExecute 31 오류 영구 해결

## [v1.6.0] - 2026-08-10
### Changed
- 일괄 인쇄 기능 범위 변경: 업로드한 4가지 PDF 서류만 순서대로 즉시 일괄 인쇄

## [v1.5.0] - 2026-08-10
### Added
- 2번 추출 데이터 섹션 내 핵심 3개 항목 원클릭 클립보드 복사 기능 구현

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
