# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.4.2] - 2026-08-10
### Fixed
- Windows 탐색기 Drag & Drop 경로 파서 정밀화 (urllib.parse.unquote, os.path.normpath, 중괄호/따옴표 제거)로 .pdf 드롭 실패 및 'PDF 파일만 업로드할 수 있습니다' 오경고 문제 완벽 해결

## [v1.4.1] - 2026-08-10
### Fixed
- 상단 툴바 도움말/설정 버튼 및 사이드바 카드에 남아있던 소수점 폰트 크기(8.5)를 정수(8/9)로 모두 교정하여 TclError 구동 방지

## [v1.4.0] - 2026-08-10
### Added
- 브랜드 이름 **VoucherPass** 및 신규 메인 로고/버전 배지 적용
- macOS 스타일 신호등 제어 아이콘 및 좌측 사이드바 구현
- 4종 파스텔 원형 배경 아이콘 (파랑, 초록, 보라, 오렌지) 적용 대형 GlassDropZone 카드 디자인 100% 매칭 적용

## [v1.3.1] - 2026-08-10
### Fixed
- GlassDropZone 내 소수점 폰트 크기를 모두 정수(9/10/11)로 완벽 교정하여 TclError 구동 방지

## [v1.3.0] - 2026-08-10
### Added
- PDF 제출 서류 Drag & Drop 업로드 영역(1번 섹션) 크기 최대화

## [v1.2.1] - 2026-08-10
### Fixed
- Tkinter Engine Tcl 호환성을 위해 폰트 사이즈 파라미터 소수점 교정

## [v1.2.0] - 2026-08-10
### Added
- Apple 글래스모피즘(Glassmorphism) 컨셉의 모던 디자인 UI 적용 (pp_gui.py)

## [v1.1.1] - 2026-08-10
### Fixed
- PR No. (예: S202607270003) 및 일련번호 숫자에 의한 공급가액 파싱 오류 수정

## [v1.1.0] - 2026-08-10
### Added
- PR Print (구매요청서 PDF), 거래명세서, 전자 세금계산서 PDF 데이터 자동 파싱 기능 (pdf_parser.py)

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
