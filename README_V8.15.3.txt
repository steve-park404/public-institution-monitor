# V8.15.3 게시판 자동 발견기

V8.15.2의 `Invalid IPv6 URL` 오류를 수정한 버전.

## 수정
- malformed URL을 `normalize_url()`에서 안전하게 폐기
- IPv6/port 파싱 오류가 전체 실행을 중단시키지 않음
- 기관별 게시판 발견 오류도 해당 기관만 `URL_ERROR`/`DISCOVERY_ERROR`로 기록
- 정상 기관은 계속 처리
- 300기관 미만 입력이면 여전히 전체 실행 중단
- eGov selectBoardList.do + bbsId 우선 탐색 유지

## 성공 기준
로그 첫 부분:
[기관목록] source=..., targets=약 355

마지막:
status_counts에 DISCOVERED_EGOV / DISCOVERED_GENERAL / CACHE_VERIFIED 등이 나타나면 정상적으로 탐색이 진행된 것.
