V8.14.15
목표:
1) CANDIDATE_NOT_VERIFIED 기관의 게시판 후보를 점수화/등급화
2) 홈페이지 URL 복구 후보 생성
3) 기존 Telegram/중복방지/checked_posts/board_cache/CSV/GitHub commit 구조 유지

주의:
- 실제 게시판 선정 로직과의 통합 여부는 실행 결과의 기관별_상태.csv에서 반드시 검증한다.
- PROBABLE_BOARD를 자동 모니터링 대상으로 즉시 편입하지 않고 진단용으로 우선 사용한다.
