# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.5.1] - 2026-08-10
### Fixed
- 2번 추출 데이터 라벨 및 복사 버튼 폰트 크기 파라미터 소수점(8.5)을 정수(8/9)로 모두 교정하여 TclError 구동 오류 해결

## [v1.5.0] - 2026-08-10
### Added
- 2번 추출 데이터 섹션 내 핵심 3개 항목(**작성일자**, **PR Title**, **공급가액**) 원클릭 클립보드 복사 버튼 ([📋 복사]) 구현
- 엑셀 열/행에 한 번에 붙여넣을 수 있는 **✨ 3종 항목 한꺼번에 복사 (Tab 구분)** 일괄 복사 버튼 추가

## [v1.4.2] - 2026-08-10
### Fixed
- Windows 탐색기 Drag & Drop 경로 파서 정밀화 (urllib.parse.unquote, os.path.normpath, 중괄호/따옴표 제거)로 .pdf 드롭 실패 오경고 문제 해결

## [v1.4.0] - 2026-08-10
### Added
- 브랜드 이름 **VoucherPass** 및 신규 메인 로고/버전 배지 적용
- macOS 스타일 신호등 제어 아이콘 및 좌측 사이드바 구현
- 4종 파스텔 원형 배경 아이콘 적용 대형 GlassDropZone 카드 디자인 100% 매칭 적용

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
