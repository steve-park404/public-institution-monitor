# V8.4 공지사항 자동검색

V8.2의 안정적인 일일 모니터링 엔진을 그대로 유지하고,
V8.3 후보검증 결과의 A등급 중 '공지성 게시판'만 보수적으로 통합한 버전입니다.

## 핵심 변경
- 일일 홈페이지 전체 탐색 없음
- 사전 검증된 게시판 URL만 검색
- V8.2 기존 대상 유지
- V8.3 A등급 중 공지/알림/새소식/기관소식 등만 추가
- 기관별 + canonical URL 기준 중복 제거
- jsessionid 등 동적 세션 URL 제외
- 명백한 상세페이지 형태 제외
- 검색 엔진은 V8.2와 동일

## GitHub Actions
실제 workflow는 `.github/workflows/v8_4_monitor.yml`에 있습니다.
편의상 저장소 루트에도 `v8_4_monitor.yml`을 함께 넣었습니다.

## Telegram
GitHub 저장소 Settings → Secrets and variables → Actions에 다음 2개를 설정합니다.
- TELEGRAM_BOT_TOKEN
- TELEGRAM_CHAT_ID

첫 실행은 `SEED_ON_FIRST_RUN=true`이므로 기존 게시물은 알림하지 않고 상태만 저장합니다.
이후 새로 발견된 키워드 일치 게시물부터 Telegram으로 알립니다.
