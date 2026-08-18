# Changelog

All notable changes to VoucherPass will be documented in this file.

## [v8.3.4] - 2026-08-18
### Patch Release: Explicit PDF File Path Input into Windows Save As Dialog
- **Windows 다른 이름으로 저장 대화상자 파일 경로 자동 입력 적용**:
  Save As 대화상자의 파일명이 비어 있어 엔터 키 입력 시 파일 생성이 무시되던 원인을 해결하기 위해, 임시 폴더 내 고유 PDF 전체 경로(NTS_eTaxInvoice_{timestamp}.pdf)를 클립보드로 붙여넣은 뒤 Enter를 실행하여 100% 무조건 PDF 저장을 확정하도록 개선.

## [v8.3.3] - 2026-08-18
### Patch Release: Direct Type-In Password Input & Prevent Text Selection Glitch
