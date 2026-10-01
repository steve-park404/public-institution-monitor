# V8.26 — 723기관 수동 공지게시판 기반 일일 모니터링

## 기준 파일
`monitor_targets_통합.xlsx`

- 대상 기관: 723개
- 공지게시판 URL 입력: 500개
- 나머지는 기존 자동 게시판 탐색 로직으로 보완
- 기존 V8.25의 8개 키워드와 중복방지 로직 유지

## 적용 방법

1. 저장소 루트에 `monitor_targets_통합.xlsx` 업로드
2. `V8.26_patch.py` 업로드
3. `.github/workflows/V8.26_723기관_수동게시판_일일모니터링.yml` 업로드
4. GitHub Actions에서 `V8.26 723기관 수동게시판 일일 모니터링`을 한 번 수동 실행
5. 성공 확인 후 매일 자동 실행

## 자동 실행 시간
GitHub Actions cron은 UTC 기준입니다.

`20 0 * * *` = 한국시간 매일 09:20

## 주의
기존 `V8.25_480기관...` workflow와 `merge_permanent_targets.py`가 동시에 같은 상태 파일을 갱신하면 중복 실행이 생길 수 있으므로,
V8.26 검증 후 기존 V8.25 스케줄 workflow는 Disable 하는 것을 권장합니다.

## Telegram Secrets
기존과 동일하게 다음 Secrets가 필요합니다.

- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`
