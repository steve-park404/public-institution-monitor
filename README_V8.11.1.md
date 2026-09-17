# V8.11.1 — 355기관 게시판 탐색 진단형

V8.11 #1 결과(`completed=160`, `no_board=195`)를 기준으로 실제 코드 흐름에 연결한 진단 버전입니다.

## 핵심
- 기존 V8.11의 검증된 `discover_board()`를 먼저 실행
- 실패한 기관에만 `discover_board_v8111()` 실행
- 추가 탐색은 홈페이지 → 관련 메뉴 1단계로 제한하여 실행시간 폭증 방지
- 후보 URL 수, 검증 성공 수, 최고 점수 URL을 `board_discovery_log.json`에 기록
- 새로 발견된 게시판은 `state.json`의 `boards`에 저장하여 다음 실행부터 재탐색하지 않음
- Telegram pending queue와 3개 키워드 로직 유지

## GitHub Actions
- `.github/workflows/v8_11_1_355_monitor.yml`

## 확인할 로그
- `SUMMARY`의 `board_discovery_diag`
- `BOARD_DISCOVERY_DIAG`
- 생성되는 `board_discovery_log.json`

예상 진단값:
- `VERIFIED`: 추가 탐색으로 실제 게시판 확인
- `CANDIDATE_NOT_VERIFIED`: 후보는 있었지만 게시판 검증 실패
- `NO_CANDIDATE`: 후보 자체를 찾지 못함
- `HOME_ERROR`: 홈페이지 접근 실패
