# V8.14.6

355기관 티끌 알림봇 실행시간 개선 버전.

## 핵심 변경
- V8.14.5의 board_cache, URL/fingerprint 중복방지, retry_queue 승계
- 최근 게시물 후보 최대 10개로 축소
- `checked_posts.json` 추가: 기관+게시일+게시물 제목 기반 식별자를 기록해 이미 확인한 게시물의 상세페이지 재조회 방지
- URL이 매번 바뀌는 사이트에서도 동일 제목/날짜 게시물의 반복 상세조회 감소
- 실제 신규/미확인 게시물만 상세페이지 조회
- `checked_post_identity_ledger`, `skipped_previously_checked` 진단 추가
- 355기관 전체 처리 검증 및 timeout 검증 유지

## 주의
첫 실행은 기존 checked_posts가 없어 기존과 비슷한 시간이 걸릴 수 있습니다. 이후 일일 실행부터 상세조회가 크게 줄어드는 구조입니다.
