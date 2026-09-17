# V8.6.4 — 100개 공지사항 오탐 개선

## 목적
- 사전 검증된 100개 공지사항 게시판만 확인
- 검색 키워드는 정확히 4개만 사용:
  - 설문조사
  - 시민참여
  - 국민참여
  - 공모전
- 게시판에서 실제 게시물 상세 링크를 선별한 뒤 제목과 본문을 검색
- 사이트맵/검색/목록/메뉴/페이지네이션 URL 및 숫자·범용 UI 제목을 최대한 제외
- 본문은 신뢰도 높은 article/main/view/content 영역이 확인될 때만 검색하여 공통 메뉴의 키워드 오탐을 줄임
- 동일 실행에서 같은 게시물이 중복 매칭되지 않도록 run-level 중복 제거
- Telegram은 기존 `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID` Secret 사용

## V8.6.3 대비 핵심 변경
1. `/sitemap`, `/search`, `/login`, `/index`, `/main` 등 비게시물 URL 강제 제외
2. `mode=list` 및 페이지네이션 파라미터가 있는 목록 URL 제외
3. 게시물 식별자(`articleNo`, `nttId`, `seq`, `idx`, `wr_id`, `num`, `no` 등) 또는 `view/detail/read/article` 경로가 있는 상세 링크를 우선
4. `기타서비스`, `서브페이지 비쥬얼`, `이전/다음/목록/더보기` 같은 UI 문구 제외 및 제목 정리
5. `content.do?cmsId=...` 같은 일반 콘텐츠 페이지는 명확한 상세 식별자가 없으면 게시물로 취급하지 않음
6. 게시판 자체 URL은 게시물 상세로 처리하지 않음
7. 상세 페이지에서 고신뢰 본문 영역을 찾지 못하면 본문 검색을 하지 않음
8. 동일 게시물은 같은 실행에서 한 번만 처리
9. `monitor_log.json`에 후보 링크 수/상세 링크 수/매칭 상세정보 기록

## 실행
GitHub Actions → `V8.6.4 공지사항 자동검색` → Run workflow

일일 자동 실행은 기존과 같이 UTC 00:10(한국시간 09:10)이다.

## 환경변수
- MAX_CONCURRENCY=12
- TIMEOUT_SECONDS=15
- RECENT_POSTS=15
- MAX_TOTAL_SECONDS=600
- HTTP_RETRIES=2
- TELEGRAM_MAX_SEND=20
