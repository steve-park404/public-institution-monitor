# V8.6 100개 공지사항 자동검색

## 목적
V8.5.1에서 안정적으로 동작한 50개 대상에 검증된 공지사항 게시판 후보 50개를 추가하여 총 100개 기관/게시판을 하루 1회 자동 검색합니다.

## 검색 방식
- 대상: `monitor_targets.xlsx`
- 게시판: 사전 검증된 공지사항 중심
- 최근 게시물: 게시판당 최대 15건
- 검색 범위: 제목 + 필요 시 본문
- 동시 처리: 12개
- 페이지 타임아웃: 15초
- HTTP 재시도: 2회
- 전체 실행 상한: 600초
- 첫 실행: 기존 매칭을 Telegram으로 대량 발송하지 않고 state에 등록
- 이후 실행: 새로 발견된 매칭만 Telegram 발송

## GitHub Actions
`.github/workflows/v8_6_monitor.yml`이 매일 UTC 00:10(한국시간 09:10)에 실행됩니다.
수동 실행도 가능합니다.

## 설치
저장소 루트에 아래 파일/폴더를 그대로 덮어씁니다.

- `monitor.py`
- `monitor_targets.xlsx`
- `keywords.txt`
- `requirements.txt`
- `.github/workflows/v8_6_monitor.yml`

기존 `state.json`, `monitor_log.json`은 유지하세요.

## Telegram Secrets
GitHub Repository → Settings → Secrets and variables → Actions에
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`
를 등록합니다.

## 주의
이번 버전은 100개까지 대상을 확장한 단계입니다. 실행 결과에서 정상 처리율, HTTP 0 오류, 실제 신규 매칭의 품질을 확인한 뒤 150~200개로 확대하는 것을 권장합니다.
