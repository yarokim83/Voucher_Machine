# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.6.1] - 2026-08-10
### Fixed
- ShellExecute Error 31 (시스템에 부착된 장치가 작동하지 않습니다) 오류 회피를 위해 Microsoft Edge 무음 프린트 엔진 및 PowerShell Start-Process 백업 인쇄 다중 엔진 구축 (printer_handler.py)

## [v1.6.0] - 2026-08-10
### Changed
- 일괄 인쇄 기능 범위 변경: Voucher 엑셀 파일 생성을 거치지 않고 **업로드한 4가지 PDF 서류**(PR Print, 거래명세서, 세금계산서, 계약서 지정 페이지)만 순서대로 즉시 일괄 인쇄하는 print_pdf_documents_only() 로직 적용

## [v1.5.1] - 2026-08-10
### Fixed
- 2번 추출 데이터 라벨 및 복사 버튼 폰트 크기 파라미터 소수점(8.5)을 정수(8/9)로 모두 교정하여 TclError 구동 오류 해결

## [v1.5.0] - 2026-08-10
### Added
- 2번 추출 데이터 섹션 내 핵심 3개 항목(**작성일자**, **PR Title**, **공급가액**) 원클릭 클립보드 복사 버튼 ([📋 복사]) 구현

## [v1.4.0] - 2026-08-10
### Added
- 브랜드 이름 **VoucherPass** 및 신규 메인 로고/버전 배지 적용
- macOS 스타일 신호등 제어 아이콘 및 좌측 사이드바 구현
- 4종 파스텔 원형 배경 아이콘 적용 대형 GlassDropZone 카드 디자인 100% 매칭 적용

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
