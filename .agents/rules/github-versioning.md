# GitHub 업로드 및 버전 관리 규칙 (Voucher_Machine)

## 대상 저장소
- **GitHub URL**: https://github.com/yarokim83/Voucher_Machine.git
- **로컬 경로**: `c:\Users\baewoong.kim\.gemini\Voucher_Machine`

---

## 🔰 규칙 1: Git 초기화 (최초 1회)

로컬에 `.git` 폴더가 없는 경우, 다음 절차를 **반드시** 먼저 수행한다:

`powershell
git init
git remote add origin https://github.com/yarokim83/Voucher_Machine.git
git branch -M main
`

---

## 📦 규칙 2: 버전 번호 체계 (Semantic Versioning)

버전 번호는 `MAJOR.MINOR.PATCH` 형식을 따른다:

| 구분 | 증가 조건 | 예시 |
|------|-----------|------|
| **MAJOR** | 기존 기능과 호환되지 않는 큰 변경 (UI 전면 개편, DB 구조 변경 등) | `1.0.0 → 2.0.0` |
| **MINOR** | 하위 호환되는 새 기능 추가 | `1.0.0 → 1.1.0` |
| **PATCH** | 버그 수정, 오탈자 수정, 소규모 개선 | `1.0.0 → 1.0.1` |

---

## 📋 규칙 3: 업데이트 전 체크리스트

1. 현재 버전 확인 (CHANGELOG.md 또는 version.txt)
2. 변경 내용 분류 (MAJOR / MINOR / PATCH)
3. version.txt 업데이트
4. CHANGELOG.md 업데이트
5. 핵심 기능 테스트 확인

---

## 🚀 규칙 4: GitHub 업로드 절차

`
git status
git add .
git commit -m "<타입>: <변경 요약> [v<버전번호>]"
git tag -a v<버전번호> -m "Release v<버전번호>: <릴리즈 요약>"
git push origin main
git push origin --tags
`

커밋 타입: feat / fix / refactor / docs / style / chore

---

## 📝 규칙 5: CHANGELOG.md 작성 형식

`
## [v1.2.0] - 2026-08-10
### Added
- 새 기능
### Fixed
- 버그 수정
`

---

## 📁 규칙 6: 필수 관리 파일

- version.txt: 현재 버전 번호
- CHANGELOG.md: 버전별 변경 이력
- README.md: 프로그램 설명
- .gitignore: 추적 제외 목록

---

## 💡 AI 에이전트 행동 지침

GitHub 업로드 요청 시:
1. 버전 자동 제안 (변경 내용 분류 기반)
2. version.txt 및 CHANGELOG.md 먼저 수정
3. 커밋 메시지를 규정 형식으로 작성
4. 태그 포함하여 push
5. 완료 후 버전 및 변경 내용 요약 보고
