# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.4.0] - 2026-08-10
### Added
- 브랜드 이름 **VoucherPass** 및 신규 메인 로고/버전 배지 적용
- macOS 스타일 신호등 제어 아이콘 및 좌측 사이드바 (PDF 업로드, 데이터 연동, 작성/인쇄, 히스토리, 설정, 안전 카드) 구현
- 4종 파스텔 원형 배경 아이콘 (파랑, 초록, 보라, 오렌지) 적용 대형 GlassDropZone 카드 디자인 100% 매칭 적용
- 우측 상단 파일 선택, 도움말, 설정 팝업 버튼 및 하단 슬림 바 내 VoucherPass 액션 버튼 디자인 완성

## [v1.3.1] - 2026-08-10
### Fixed
- GlassDropZone 내 ont=("Malgun Gothic", 9.5) 등 남아있던 소수점 폰트 크기를 모두 정수(9/10/11)로 완벽 교정하여 TclError 구동 방지

## [v1.3.0] - 2026-08-10
### Added
- PDF 제출 서류 Drag & Drop 업로드 영역(1번 섹션) 크기 최대화 (시원시원한 대형 Drop Zone 카드 적용)
- 파싱 데이터 확인 및 수정 영역(2번 섹션) 1~2줄 슬림 라인 컴팩트화 (세로 높이 최소화)

## [v1.2.1] - 2026-08-10
### Fixed
- Tkinter Engine Tcl 호환성을 위해 폰트 사이즈 파라미터 소수점(9.5/11.5)을 정수(9/10/11/12)로 교정하여 프로그램 실행 오류 해결

## [v1.2.0] - 2026-08-10
### Added
- Apple 글래스모피즘(Glassmorphism) 컨셉의 모던 디자인 UI 적용 (pp_gui.py)
- 제출 서류 PDF(PR Print, 거래명세서, 세금계산서, 계약서) Drag & Drop (드래그 앤 드롭) 업로드 카드 구현 (GlassDropZone)

## [v1.1.1] - 2026-08-10
### Fixed
- PR No. (예: S202607270003) 및 10자리 이상 승인번호/일련번호 숫자가 공급가액(Amount)으로 잘못 추출되는 파싱 오류 수정

## [v1.1.0] - 2026-08-10
### Added
- PR Print (구매요청서 PDF), 거래명세서, 전자 세금계산서 PDF 데이터 자동 파싱 기능 (pdf_parser.py)
- 바탕화면의 고려제강(2025).xlsx Voucher 템플릿에 PR No., PR Title, 금액, 날짜, 거래처 정보 자동 기입 및 저장 기능 (excel_handler.py)

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
