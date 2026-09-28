# V8.15.1 게시판 전용 자동 발견기

V8.15 #1의 `targets: 0` 문제를 수정한 버전입니다.

## 핵심 수정
1. 특정 CSV 파일명을 강제하지 않음
2. 저장소 내 CSV를 검사하여 `기관명 + 홈페이지/URL` 컬럼을 자동 탐색
3. 300개 미만이면 즉시 실패하여 `targets: 0` 같은 가짜 성공 방지
4. 355기관을 실제 입력으로 확인
5. eGov `selectBoardList.do?bbsId=...` 우선 탐색
6. 게시판 실제 재접속 검증
7. `board_cache.json`에 검증된 URL 저장

## 실행 후 확인할 값

반드시:
- `targets` ≈ 355
- `found_or_verified` > 0
- `not_found` < targets

가 나와야 합니다.

`targets: 0` 또는 `targets < 300`이면 GitHub Actions가 실패해야 정상입니다.
