V8.14.18 미확인 120개 기관 진단 프로그램

목적
- 기존 355기관 본체 모니터링 코드는 변경하지 않는다.
- 기관별_상태.csv에서 HOME_ERROR와 CANDIDATE_NOT_VERIFIED만 추출한다.
- http/https, www/non-www, 루트 URL 변형을 시험한다.
- 홈페이지 링크와 대표 게시판 경로를 탐색한다.
- 게시판 후보의 구조를 진단한다.

실행
python v8_14_18_unconfirmed_diagnostic.py

생성 파일
- 기관별_미확인120_진단.csv
- 기관별_미확인120_진단.json

판정
- RECOVERED_CANDIDATE: 복구 후보 확인. 다음 본체 버전에 반영 후보
- CANDIDATE_NEEDS_RULE_REVIEW: 후보는 있으나 검증규칙 추가 검토 필요
- HOME_UNRECOVERED: 홈페이지 변형 URL까지 실패
- NO_BOARD_FOUND: 홈페이지는 접근되나 게시판 후보 부족

주의
- Telegram 발송 없음
- state.json / board_cache.json / pending.json 변경 없음
- 본체 모니터링 로직을 직접 변경하지 않음
