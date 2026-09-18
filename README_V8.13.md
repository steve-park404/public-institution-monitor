# V8.13.0

V8.12.9 다음 단계: **실제 게시물 판별 강화 + 오탐 진단 강화**.

핵심 변경:
1. 목록 URL 차단을 게시물 후보 단계에서도 적용.
2. 상세 URL이 실제 목록 페이지인지 DOM 신호로 재검증.
3. 기관 홈페이지/일반 콘텐츠 경로를 상세 게시물에서 차단.
4. 전용 본문 영역 → main/article 순으로 본문을 분리.
5. 사이트명/페이지명만 제목으로 잡힌 경우 차단.
6. 실제 게시물 구조(제목 + 본문 + 작성/등록 메타 또는 충분한 본문)를 요구.
7. `excluded_not_post_structure`, `excluded_no_body`를 diagnostics에 추가.
8. 기존 3개 키워드, 공모전 제목 제외, 최근 30일, pending queue, Telegram 하루 20건 유지.

실행 후 특히 확인할 값:
- `title_matches`
- `body_matches`
- `excluded_list_pages`
- `excluded_generic_pages`
- `excluded_not_post_structure`
- `excluded_no_body`
- `new_matches`

주의:
- V8.13은 탐색 범위를 무리하게 넓히지 않고 **오탐 차단을 우선**한 버전이다.
- 실제 운영 결과에서 `completed/no_board`와 `new_matches`가 지나치게 낮으면 다음 단계에서 게시판 탐색을 별도로 개선한다.
