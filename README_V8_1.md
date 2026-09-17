# V8.1 공공기관 공지사항 자동검색

## 핵심 변경
- 홈페이지에서 발견한 링크를 무조건 검색하지 않음
- "공지사항/공지/알림마당/기관소식/새소식/공고" 등의 링크를 발견한 뒤 실제 게시판인지 재검증
- 실제 게시물 3개 이상 + 서로 다른 제목 3개 이상인 경우에만 자동 대상에 추가
- 상세페이지 URL은 자동발견 단계에서 제외
- 기존 monitor_targets.xlsx는 수정하지 않음
- 게시물 제목은 목록 페이지의 제목을 사용하고, 상세페이지는 검증용으로만 사용
- 제목 + 본문에서 키워드 검색
- 첫 실행에서는 기존 매칭을 Telegram으로 보내지 않고 state.json에 등록

## 파일
- monitor.py : 검색 엔진
- monitor_targets.xlsx : 기존 검증 대상
- url_완성.xlsx : 선택사항. 있으면 홈페이지 공지사항 자동발견에 사용
- keywords.txt : 검색 키워드
- state.json : 중복 알림 방지 상태
- monitor_log.json : 실행 결과
- .github/workflows/v8_monitor.yml : 실제 GitHub Actions workflow

## Telegram
GitHub Repository > Settings > Secrets and variables > Actions에 아래 두 개를 등록:
- TELEGRAM_BOT_TOKEN
- TELEGRAM_CHAT_ID

## 주의
루트의 v8_monitor.yml은 사람이 보기 위한 복사본일 뿐이며,
GitHub Actions가 실제 사용하는 파일은 `.github/workflows/v8_monitor.yml`입니다.

## 테스트
Actions > V8.1 공지사항 자동검색 > Run workflow
