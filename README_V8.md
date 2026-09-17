# V8 공공기관 공지사항 자동검색

## 핵심 방식

기존 V7.x처럼 매번 게시판 후보를 대량 발굴하는 방식이 아니라,

1. V7.6 검증 결과를 초기 게시판 목록으로 사용
2. `공지사항` 이름의 게시판을 최우선으로 사용
3. `url_완성.xlsx`가 있으면 각 기관 홈페이지에서도 `공지사항/공지/알림마당` 링크를 자동 발견
4. 게시판에서 최근 게시물 최대 20개 확인
5. 게시물의 제목 + 본문을 키워드 검색
6. 처음 발견한 글은 `state.json`에 기록하여 중복 알림 방지
7. 다음 실행에서 새로 발견된 키워드 게시물만 Telegram으로 전송
8. GitHub Actions가 매일 오전 9:10(KST)에 실행

GitHub Actions의 scheduled workflow는 기본적으로 UTC 기준이며, 현재 GitHub는 IANA timezone 지정도 지원합니다. 이 구성은 호환성을 위해 `00:10 UTC`를 사용해 KST 09:10에 실행되도록 했습니다. 또한 정각 부근에는 스케줄 지연 가능성이 있어 10분을 두었습니다.

## 파일

- `monitor.py` : 실제 검색 엔진
- `monitor_targets.xlsx` : V7.6 기반 초기 모니터링 대상
- `keywords.txt` : 검색 키워드
- `state.json` : 중복 알림 방지용 상태
- `monitor_log.json` : 최근 실행 결과
- `.github/workflows/v8_monitor.yml` : 매일 실행
- `requirements.txt` : Python 패키지

## Telegram 설정

GitHub Repository → Settings → Secrets and variables → Actions → New repository secret

다음 2개를 등록합니다.

- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

Telegram Bot API의 `sendMessage`는 `chat_id`와 `text`를 받아 메시지를 전송합니다.

## 최초 실행

기본값은 `SEED_ON_FIRST_RUN=true`입니다.

따라서 최초 실행에서는 현재 게시물들을 상태에 등록하고 과거 글을 한꺼번에 알리지 않습니다.

두 번째 실행부터 새 게시물의 키워드 매칭을 알립니다.

## 중요한 운영 원칙

`monitor_targets.xlsx`의 V7.6 검증 게시판은 삭제하지 않고 유지합니다.

홈페이지 자동발견 결과는 실행 중 메모리에서만 추가하며, 아직 자동으로 `monitor_targets.xlsx`에 저장하지 않습니다. 즉, 잘못된 자동 발견 URL이 기준 파일을 오염시키지 않습니다.

`공지사항`을 최우선으로 하되, 기존 V7.6에서 확정된 다른 게시판도 보조적으로 검색할 수 있습니다.

## 355개 기관으로 확대

`url_완성.xlsx`를 저장소 루트에 넣으면 홈페이지 자동발견 기능이 활성화됩니다.

권장 운영 방식:

- 1단계: V7.6 검증 대상 + 공지사항 자동발견
- 2단계: 자동발견 결과 중 실제 공지사항으로 확인된 URL을 `monitor_targets.xlsx`에 추가
- 3단계: 기관별 특수 게시판은 별도 추가

이렇게 하면 잘못된 메뉴/기능 URL 때문에 자동검색 품질이 떨어지는 것을 줄일 수 있습니다.
