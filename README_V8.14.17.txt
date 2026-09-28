V8.14.17

목적
- 355기관 자동 모니터링의 미확인 기관 120개(HOME_ERROR 49 + CANDIDATE_NOT_VERIFIED 71)를 집중 복구.
- 기존 정상 235개 모니터링, Telegram, 중복방지, checked_posts, board_cache, 기관별_상태.csv를 유지.

주요 변경
1. 홈페이지 복구
- 기존 URL 실패 시 HTTPS/HTTP, www/non-www, 명백한 wwww 오타를 자동 보정하여 재접속.
- 복구 성공 URL은 diagnostics.json과 기관별_상태.csv에 기록.

2. 게시판 검증 완화
- 기존 VERIFIED 검증을 우선 유지.
- 엄격 검증 실패 시 상위 후보에 대해 보조 추출(relaxed extraction)을 수행.
- 행에 날짜가 있거나 게시물 식별자(seq/ntt/idx/article/view/read/no/board/bbs)가 있는 실제 게시물 링크를 추출하면 PROBABLE_BOARD로 사용.
- 단순 메뉴/로그인/사이트맵/개인정보/약관 링크는 제외.

3. CSV
- 기관별_상태.csv 355행 강제 검증.
- PROBABLE_BOARD, 홈페이지복구시도/성공, 복구URL, 게시판판정근거, 최고후보URL/점수 기록.

4. 실행시간
- 내부 처리 한도 1200초(20분), GitHub Actions timeout 30분.

설치
- 기존 V8.14.16의 monitor.py와 generate_institution_status_csv.py를 교체.
- .github/workflows/v8_14_16_355_monitor.yml도 함께 교체한다. 파일명은 기존과 같아서 기존 일일 스케줄이 중복 생성되지 않는다.
- workflow permissions는 Settings > Actions > General > Workflow permissions > Read and write permissions가 필요.

주의
- 첫 실행 결과에서 HOME_ERROR/CANDIDATE_NOT_VERIFIED가 감소하는지 확인한다.
- PROBABLE_BOARD가 지나치게 많은 경우 기관별_상태.csv의 게시판URL과 게시물 후보를 기준으로 다음 버전에서 추가 필터링한다.
