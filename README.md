# 355개 공공기관 홈페이지 키워드 모니터

구조: 355개 기관 → 게시판 자동 발견(최초 1회) → boards.xlsx → 매일 게시판 검사 → 게시글 날짜 → 제목+본문 → `설문조사/시민참여/국민참여/공모전` → 신규 URL 판별 → Telegram

## 1. GitHub에 올릴 파일
- `url.xlsx` : 사용자의 355개 기관 원본 파일
- `monitor.py`
- `requirements.txt`
- `.github/workflows/monitor.yml`

## 2. Telegram 설정
Telegram에서 BotFather로 봇을 만들고 토큰을 받습니다. 알림을 받을 채팅에서 봇을 시작한 뒤 chat_id를 확인합니다.

GitHub 저장소 → Settings → Secrets and variables → Actions → New repository secret에서:
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

등록.

## 3. 최초 게시판 발견
GitHub Actions의 Actions → `355 Public Institution Monitor` → Run workflow를 한 번 실행하면 `boards.xlsx`가 생성됩니다.

`boards.xlsx`를 내려받아 게시판 URL이 잘못 잡힌 기관만 수정한 뒤 다시 커밋하면 됩니다.

## 4. 매일 자동 실행
Workflow는 UTC 00:00, 한국시간 오전 9시에 실행됩니다. GitHub 서버에서 실행되므로 PC/휴대폰을 켜둘 필요가 없습니다.

## 5. 중복 알림 방지
`seen_posts.json`에 게시글 URL의 SHA-256 식별자를 저장합니다. 이미 알림을 보낸 게시글은 다음 실행에서 다시 알리지 않습니다.

## 6. 수동 실행
GitHub Actions에서 `Run workflow`를 누르면 즉시 실행할 수 있습니다.
