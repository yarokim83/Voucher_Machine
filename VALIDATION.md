# v8.8.2 검증 기록

- 검증일: 2026-10-15
- Windows Python 3.12 회귀 테스트 54개 통과 (skip 없음).
- 실제 NTS HTML 첨부파일을 수정된 변환 함수로 처리하여 작성일 2026-10-15, 공급가액 2,100,000원, 부가세 210,000원, 합계 2,310,000원 확인.
- Edge 시작 실패 단계 표시, 민감한 예외 메시지 로그 제외, PDF 검증 실패 시 임시 결과 삭제와 브라우저 종료, 디버그 환경변수 없이 진단 로그 기록 검증.
- EXE·Setup clean 빌드 성공. 실행 파일 내 수정 모듈 및 설치 파일 내 EXE 해시 일치 확인.
- 새 EXE 화면에서의 실사용·설치 실행·Outlook·실물 인쇄는 수행하지 않음.
- 실제 업무 첨부파일과 생성 PDF는 Git에 포함하지 않음.

| 파일 | 바이트 | SHA-256 |
|---|---:|---|
| dist/v8.8.2/VoucherPass.exe | 80923215 | 14e3ebb05cd17cd0bc95111bf75aa388ecfa7f96852e616e3b2eb009254a9ec0 |
| dist/v8.8.2/VoucherPass_Setup.exe | 91689951 | adb59d8225d79f4512707dcecd523fc7f71133c17d664a1844010e8fe835353d |

# v8.8.1 검증 기록

- 검증일: 2026-09-14
- Windows Python 3.12.9 회귀 테스트 51개 통과 (skip 없음).
- 실제 구매요청서와 발주서 2개를 읽어 공급가액·계산 부가세·합계·거래처명을 확인.
- 같은 실제 문서를 숨긴 TkinterDnD 화면에 적용하여 값과 부가세 계산 표시 확인.
- PR/PO 표의 마지막 합계 열, 줄바꿈, 부가세 포함/별도, 명시 세액/세율, 충돌 후보, 거래처·담당자·연락처 분리 검증.
- 테스트 저장소에는 합성 예시만 사용. 실제 업무 PDF와 렌더링 이미지는 Git에 포함하지 않음.
- v8.8.1 EXE·Setup clean 빌드 성공. 실제 Outlook·프린터·설치 실행은 수행하지 않음.

| 파일 | 바이트 | SHA-256 |
|---|---:|---|
| dist/v8.8.1/VoucherPass.exe | 80840155 | fa96cb2070a6d041ffe942eef41f1dcc852b7eca76c58b2df8c52be6db59d56e |
| dist/v8.8.1/VoucherPass_Setup.exe | 91606044 | c4894d270b74f0fcead6f263a89912c610076955d5fd8165299199728634405e |

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
