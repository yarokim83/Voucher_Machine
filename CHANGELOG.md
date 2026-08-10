# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.8.0] - 2026-08-10
### Added
- 계약서 페이지 지정 범위 파싱 지원 강화 (12-13, 1-3, 1,2,5 하이픈 구간 및 쉼표 표기 지원)
- pypdf 슬라이싱 엔진 기반 계약서 지정 범위 정밀 PDF 추출 인쇄 구현 (전체 페이지 출력 현상 100% 교정)
- 3종 항목 복사 방식 변경: 줄바꿈(\n) 세로 구분 복사를 적용하여 엑셀 붙여넣기 시 세로 3개 셀에 각각 자동 분리 저장
- 좌측 사이드바 간소화: 사용하지 않는 메뉴 항목을 삭제하고 핵심 3개 메뉴만 깔끔하게 유지

## [v1.7.0] - 2026-08-10
### Added
- 초경량 무설치 **SumatraPDF Vector Native Print Engine (in/SumatraPDF.exe)** 탑재

## [v1.6.3] - 2026-08-10
### Fixed
- 300 DPI 초고해상도 비트맵 렌더링 도입

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
