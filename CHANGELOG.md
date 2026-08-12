# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v6.3.0] - 2026-08-12
### Instant Intercept Sub-Second Date Extraction Engine
- **PDF 저장이후 1초 미만 (0.05초) 극초고속 추출 파이프라인 탑재**:
  - 기존의 12초 지연을 유발하던 폴더 셋 집합(set diff) 비교 방식을 **0.05초 타임스탬프 실시간 가로채기(Instant Intercept)** 방식으로 전면 교체.
  - PDF 저장이 시작되면 0.05초 단위로 새로 작성된 PDF 세금계산서의 mtime 을 감지하여 저장 완료 즉시(0.1초 이내) 작성일자(2026/08/11)를 추출하여 📅 작성일자 칸 및 7종 HUD 보드에 100% 대입.
- **버전 동기화 커스텀 룰(Version Sync Rule) 적용**:
  ersion.txt, CHANGELOG.md, pp_gui.py 뱃지를 6.3.0 로 100% 완벽 동기화.
