# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v7.0.0] - 2026-08-12
### Major Release: Clean UX & Professional Layout Refactoring
- **1. 업로드 진행 상태 표시 뱃지 탑재**:
  상단 헤더에 진행 상태: 1/5 완료 실시간 뱃지를 배치하고 완료된 항목에 초록 체크(✓ 완료) 표시로 현재 어디까지 진행되었는지 1초 만에 인지.
- **2. 완료 vs 미완료 시각적 대비 극대화**:
  업로드 완료된 서류는 초록 배경(#F0FDF4) + 초록 실선 테두리(#16A34A), 미완료 서류는 연한 배경과 테두리로 명확한 대비 제공.
- **3. 데이터 입력 영역 분리 (2열 4행 카드 그리드 + 상단 라벨)**:
  추출 데이터 7종 섹션을 2열 4행 콤팩트 그리드로 정돈하고 라벨을 상단에 배치하여 시각적 스캔 및 가독성 대폭 향상.
- **4. 버튼 시각적 계층 명확화**:
  메인 버튼(📋 Voucher 엑셀 양식 붙여넣기)을 로열 블루(#2563EB)로 큼직하게 강조하고, 보조 버튼(보관 & 인쇄)은 하단에 1:1 동일 비중으로 나란히 배치.
- **5. 차분하고 전문적인 Clean Soft Slate 톤앤매너 통일**:
  알록달록했던 아이콘 박스를 정돈된 Slate/Cobalt 톤앤매너로 통일하여 전문적인 인상 연출.
- **버전 동기화 커스텀 룰(Version Sync Rule) 적용**:
  ersion.txt, CHANGELOG.md, pp_gui.py 뱃지를 7.0.0 로 100% 완벽 동기화.
