# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v5.4.0] - 2026-08-12
### National Tax Service Tax Invoice Format Parser & Version Sync Rule
- **국세청 표준 전자세금계산서 작성일자 파서 100% 특화 장착**:
  사용자 스크린샷 표준 양식의 작성일자 헤더 아래 표 셀에 위치한 날짜(2026/08/11, 2026.08.11, 2026-08-11)를 문맥 슬라이싱 정규식으로 100% 명중 추출하여 📅 작성일자 필드에 자동 대입.
- **버전 동기화 커스텀 룰 (Version Sync Rule) 도입**:
  버전 변경 시 ersion.txt, CHANGELOG.md, pp_gui.py 헤더 뱃지 3종이 자동으로 동기화되도록 룰 구축.
