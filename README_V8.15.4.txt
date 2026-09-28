# V8.15.4
V8.14.17의 board_cache.json을 우선 사용합니다.
verified=True이고 board_url/url이 있는 기관은 다시 탐색하지 않습니다.
미확인 기관만 최대 12개 동시 요청으로 탐색합니다.
timeout 8초, 하위 링크 최대 6개.
selectBoardList.do + bbsId를 우선합니다.
최종 기관 데이터가 300개 미만이면 실패합니다.
