# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v2.0.0] - 2026-08-12
### Added
- **세금계산서 비밀번호(6068625399) 자동 대입 해제 파싱 및 무음 인쇄 엔진** 구현 (pdf_parser.py, printer_handler.py)
- **Outlook 이메일 첨부파일 1초 자동 추출 버튼 ([📧 아웃룩 첨부파일 가져오기])** 추가 (outlook_handler.py)
- **다운로드 폴더 실시간 자동 감시 & PDF 4종 스마트 자동 수집/분류 기능** 탑재 (pdf_watcher.py)
- **건별/업체별 폴더 자동 생성 & 원클릭 보관소 기능 ([📂 건별 자동 폴더 생성 & 보관])** 구축 (excel_handler.py)
- 2.0.0 Full Automation 단일 포터블 실행 파일 (dist/VoucherPass.exe) 배포 완료

## [v1.9.2] - 2026-08-10
### Fixed
- PyInstaller collect-all 단일 독립 포터블 실행 파일 패치

## [v1.9.0] - 2026-08-10
### Added
- 세금계산서 작성일자 파싱 & 공식 로고 아이콘 적용

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
