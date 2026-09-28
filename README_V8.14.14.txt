V8.14.14 - 게시판 검증 개선 + 홈페이지 복구

핵심 변경
1. 게시판 후보를 단순 VERIFIED/실패로 보지 않고 구조적 증거 점수화
2. PROBABLE_BOARD / WEAK_CANDIDATE 상태를 별도로 진단
3. HOME_ERROR 기관은 HTTPS/HTTP 및 www/non-www 변형으로 재접속 시도
4. 기관별_상태.csv에 게시판판정/근거/홈페이지복구 진단 추가
5. 기존 Telegram, 중복방지, checked_posts, daily summary, CSV 커밋 구조 유지

목표
- CANDIDATE_NOT_VERIFIED를 무리하게 모두 통과시키지 않음
- 실제 게시판일 가능성이 높은 후보를 진단 데이터로 확보
- HOME_ERROR 중 URL 변형으로 복구 가능한 기관을 회복
