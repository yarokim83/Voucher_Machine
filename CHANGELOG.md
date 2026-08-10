# Changelog

All notable changes to Voucher_Machine will be documented in this file.

## [v1.2.0] - 2026-08-10
### Added
- Apple 글래스모피즘(Glassmorphism) 컨셉의 모던 디자인 UI 적용 (pp_gui.py)
- 제출 서류 PDF(PR Print, 거래명세서, 세금계산서, 계약서) Drag & Drop (드래그 앤 드롭) 업로드 카드 구현 (GlassDropZone)
- 드래그 앤 드롭 파일 탐지 시 자동 PDF 파싱 연동 및 상태 표시 배지 기능 추가
- 고대비 시인성 개선된 폼 레이아웃 및 애플 아쿠아/에메랄드 인터랙티브 액션 버튼 디자인 적용

## [v1.1.1] - 2026-08-10
### Fixed
- PR No. (예: S202607270003) 및 10자리 이상 승인번호/일련번호 숫자가 공급가액(Amount)으로 잘못 추출되는 파싱 오류 수정
- Total, 총약정금액, 공급가액 키워드 기반 금액 추출 1순위 정밀화

## [v1.1.0] - 2026-08-10
### Added
- PR Print (구매요청서 PDF), 거래명세서, 전자 세금계산서 PDF 데이터 자동 파싱 기능 (pdf_parser.py)
- 바탕화면의 고려제강(2025).xlsx Voucher 템플릿에 PR No., PR Title, 금액, 날짜, 거래처 정보 자동 기입 및 저장 기능 (excel_handler.py)
- Windows 프린터 연동을 통한 Voucher 엑셀 및 PDF 서류 3종 + 계약서 특정 페이지 일괄 인쇄 기능 (printer_handler.py)
- Tkinter 기반 모던 데스크톱 GUI 애플리케이션 화면 제공 (pp_gui.py, main.py)

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
