# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v6.2.0] - 2026-08-12
### Sub-Second Ultra Fast PDF Date Extractor Engine
- **작성일자 추출 파이프라인 초고속화 (1초 미만 파싱)**:
  - 브라우저 인쇄 대기 지연 시간을 0.6초/0.15초/0.4초로 최적화.
  - 신규 PDF 저장 감지 폴링 주기를 0.2초(High Frequency Polling)로 단축하여 PDF 가 생성되자마자 0.2초 만에 감지.
  - PDF 텍스트 파싱 대기시간을 0.15초로 단축하여 단 1초 미만에 작성일자(2026/08/11) 자동 대입 완성.
- **버전 동기화 커스텀 룰(Version Sync Rule) 적용**:
  ersion.txt, CHANGELOG.md, pp_gui.py 뱃지를 6.2.0 로 100% 완벽 동기화.
