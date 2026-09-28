# V8.15 게시판 전용 자동 발견기

## 목적
V8.14.17에서 `HOME_ERROR` / `CANDIDATE_NOT_VERIFIED`로 남는 기관의 실제 공지사항 게시판 URL을 자동으로 찾아 `board_cache.json`에 저장합니다.

### 핵심
- 홈페이지 메뉴 링크 분석
- 하위 알림/공지 링크 1단계 추가 탐색
- 전자정부 표준프레임워크 `selectBoardList.do?bbsId=...` 우선 탐색
- 실제 게시판 페이지 재조회 후 검증
- 검증된 URL은 다음 실행부터 캐시 사용
- 기존 cache 내용 보존

## 우체국물류지원단 사례
실제 게시판은 다음 형태입니다.

https://www.pola.or.kr/web/cop/bbs/selectBoardList.do?bbsId=BBSMSTR_000000000018

이처럼 `selectBoardList.do` + `bbsId=BBSMSTR_...` 형태를 일반화하여 찾습니다.

## 설치
저장소 루트에 아래 파일을 추가합니다.

- `board_discovery_v8_15.py`
- `.github/workflows/v8_15_board_discovery.yml`

필요 패키지:
- requests
- beautifulsoup4

## 실행
GitHub Actions에서
`V8.15 게시판 전용 자동 발견` → `Run workflow`

또는 로컬:
`python board_discovery_v8_15.py`

## 생성 파일
- `board_cache.json`
- `게시판_자동발견_결과.csv`
- `게시판_자동발견_진단.json`

## 주의
이 모듈은 기존 monitor.py를 직접 변경하지 않습니다.
먼저 주 1회 게시판을 발견하고 cache를 축적한 뒤,
V8.15.x에서 monitor.py가 이 cache를 우선 사용하도록 연결하는 것이 안전합니다.
