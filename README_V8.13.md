# V8.14.1

V8.14.0을 한 번 더 정리한 안정화 버전.

## 핵심

1. V8.13.1의 기존 `seen` URL을 `sent_urls`로 자동 승계.
2. 이미 발송된 URL은 pending에 남아 있어도 재발송하지 않음.
3. pending 자체도 URL 기준으로 정리.
4. 같은 URL의 중복 발견을 canonical URL 기준으로 차단.
5. GitHub Actions 동시 실행 방지.
6. state/pending/diagnostics 충돌 자동 처리.
7. 기존 탐지 조건 유지:
   - 설문조사 / 시민참여 / 국민참여
   - 제목 공모전 제외
   - 최근 30일
   - Telegram 최대 20건

## V8.14.1 첫 실행 확인

diagnostics에서:
- `migrated_from_previous`
- `sent_url_ledger`
- `new_matches`
- `telegram_sent`
- `pending_after`

를 확인한다.

기존 state.json의 `seen`이 정상적으로 남아 있다면 과거에 이미 발송한 URL은
V8.14.1에서 다시 Telegram으로 발송되지 않는다.

주의:
state.json 자체가 과거 실행에서 완전히 유실되어 있다면 그 시점의 실제 발송
이력은 파일만으로 복구할 수 없다. 이 경우에도 V8.14.1은 이후부터
`sent_urls`를 영구 관리하여 같은 문제가 재발하지 않도록 한다.
