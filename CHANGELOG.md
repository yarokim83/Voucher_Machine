# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v6.1.0] - 2026-08-12
### 3-Step Pipeline: HTML Drop -> Auto Save PDF -> Parse Date from Saved PDF
- **3단계 연속 파이프라인 탑재**:
  1. HTML 메일 세금계산서 드래그 앤 드롭
  2. 비밀번호(6068625399) 자동 대입 후 지정 폴더에 PDF 세금계산서 자동 저장
  3. 저장된 PDF 세금계산서에서 작성일자(2026/08/11)를 100% 자동 파싱하여 📅 작성일자 칸 및 7종 HUD에 대입.
- **버전 동기화 커스텀 룰(Version Sync Rule) 적용**:
  ersion.txt, CHANGELOG.md, pp_gui.py 뱃지를 6.1.0 로 100% 완벽 동기화.
