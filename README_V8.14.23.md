# V8.14.23 오탐차단 + YML 완성 패키지

## 포함 파일

- `V8.14.23_patch.py`
- `.github/workflows/V8.14.23_오탐차단_모니터링.yml`

## 적용 방법

1. 압축 해제
2. 두 파일을 GitHub 저장소 루트에 업로드
3. `.github/workflows/` 아래의 YML 위치가 유지되도록 업로드
4. GitHub Actions에서 **V8.14.23 오탐차단 모니터링** 실행

## 전제

현재 저장소의 `monitor.py`가 V8.14.22인 것을 기준으로 합니다.

## 주요 변경

- 공모전 개최/모집은 유지
- 공모전 수상작/수상자/후보작/결과/선정/심사 결과 차단
- `/index.do`, `/main.do` 등 홈페이지 차단
- 기관 홈페이지/브랜드 페이지의 BODY 오탐 차단
- 설문조사/국민참여/시민참여 BODY 문맥 강화
- 기존 state, board_cache, Telegram, 중복방지 로직 유지

## 실행 순서

Checkout
→ Python 3.11
→ requests/BeautifulSoup/openpyxl 설치
→ V8.14.23 패치
→ monitor.py 문법 검사
→ 기관 수 검사
→ monitor.py 실행
→ 결과 Artifact 저장
→ 상태 파일과 monitor.py 커밋

## 중요

이 YML은 별도의 게시판 URL 탐색 로직을 추가하지 않습니다.
현재 수동 확인된 게시판 URL 및 기존 모니터링 구조를 그대로 사용합니다.
