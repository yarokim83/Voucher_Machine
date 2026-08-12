# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v2.1.0] - 2026-08-12
### Fixed
- **결정적 근본 원인 해결**: printer_handler.py 5번째 줄의 불필요한 import pdfplumber 의존성 구문을 완전히 제거하여 No module named 'pdfplumber' 예외 100% 영구적으로 소멸
- pdf_parser.py 및 전체 시스템에 pypdf 백업 파서 탑재 완료
- 바탕화면 포터블 독립 실행 파일 VoucherPass.exe 최종 배포

## [v2.0.0] - 2026-08-12
### Added
- 세금계산서 비밀번호(6068625399) 자동 대입 해제 파싱 및 무음 인쇄 엔진 구현
- Outlook 이메일 첨부파일 1초 자동 추출 버튼 구현
- 다운로드 폴더 실시간 자동 감시 & PDF 4종 스마트 자동 분류 기능 구현
- 건별/업체별 폴더 자동 생성 & 원클릭 보관소 기능 구현

## [v1.0.0] - 2026-08-10
### Added
- 초기 프로젝트 생성 및 GitHub 버전 관리 룰 설정
