# V8.15.5 게시판URL 직접검색형

## 목적
V8.14.17의 자동 게시판 발견 기능을 사용하지 않고,
`기관별_상태.csv`에 직접 입력한 `게시판URL`을 기준으로 355개 기관을 검색합니다.

현재 제공된 CSV 기준:
- 전체 기관: 355
- 게시판URL 입력: 348
- 게시판URL 미입력: 7

## 검색 조건
- 최근 30일
- 제목 + 본문
- 키워드: `설문조사`, `시민참여`, `국민참여`
- 제목에 `공모전` 포함 시 제외
- Telegram 최대 20건/회
- URL/fingerprint/게시물 identity 중복 방지

## 설치
이 폴더의 파일을 기존 GitHub 저장소 루트에 복사합니다.

필수:
- `기관별_상태.csv`
- `monitor_v8_15_5.py`
- `requirements.txt`
- `.github/workflows/V8.15.5.yml`

GitHub Secrets:
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

## 주의
이 버전은 `게시판URL`을 직접 지정하는 방식입니다.
따라서 게시판 자동발견으로 `board_cache.json`을 덮어쓰지 않습니다.

`게시판URL`이 비어 있는 7개 기관은 자동 검색 대상에서 제외하고
검증결과에 `URL없음`으로 기록합니다.

## 권장 운영
먼저 GitHub Actions에서 `workflow_dispatch`로 수동 1회 실행하여
`게시판검증결과_V8.15.5.csv`와 `검색결과_V8.15.5.csv`를 확인한 뒤
정상 동작을 확인하면 schedule을 그대로 사용합니다.
