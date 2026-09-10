# v8.8.0 검증 기록

- 검증일: 2026-09-08
- Windows Python 3.12.9에서 `py -3.12 -m unittest discover -s tests -v`: 34개 통과, skip 없음.
- 실제 TkinterDnD 루트의 숨긴 창에서 등록·리셋·이전 비동기 결과 차단 확인.
- 임시 암호화 PDF의 실제 파싱, 계약서 페이지 추출, Excel 셀 저장 확인.
- Outlook COM 및 인쇄 전송은 모의 객체로 검증. 실제 메일·프린터는 조작하지 않음.
- `py -3.12 make_installer.py`: EXE 및 Setup clean 빌드 성공.
- 기존 dist/VoucherPass.exe 접근 오류로 버전별 출력 폴더를 도입하여 기존 파일을 유지.
- 업무 양식별 정확도, 실제 Edge 복호화, 실물 인쇄 및 설치 실행은 추가 현장 확인 대상.

| 파일 | 바이트 | SHA-256 |
|---|---:|---|
| dist/v8.8.0/VoucherPass.exe | 79766308 | 854f6302377e61a70965063fe9911b09d4bfff0df585ecbe8169e4eb1b9a7bd8 |
| dist/v8.8.0/VoucherPass_Setup.exe | 89502771 | 166eea8231dc0810d2ada09b5059f2970c92570fa9950f7645c99db93be1c7fd |
