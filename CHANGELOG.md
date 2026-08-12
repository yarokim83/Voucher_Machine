# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v5.2.3] - 2026-08-12
### Tax Invoice PDF Saving & Lock-Safe Extraction Hotfix
- **새로 저장된 세금계산서 PDF 파일 락 해제 대기 파이프라인 탑재**:
  세금계산서 PDF 저장이 실행된 직후, 파일 쓰기 완료(0 bytes -> 500 bytes 이상) 및 OS 파일 락이 풀릴 때까지 최대 3초 대기 후 복호화 템프 생성 및 3회 재시도 파싱 수행.
- **날짜 대입 팝업 100% 보장**:
  추출 완료 시 UI 📅 작성일자 필드에 즉시 입력하고 축하 안내 팝업 표출.
