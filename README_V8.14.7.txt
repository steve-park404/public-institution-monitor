# V8.14.7 - 355기관 공지사항 모니터링

## 주요 변경
1. `기관별_상태.csv` 자동 생성
   - 355개 기관 전체를 1행씩 기록
   - 실행시간 초과로 처리되지 않은 기관도 `UNPROCESSED_TIMEOUT`으로 기록
   - 홈페이지/게시판 상태, 게시물 후보, 상세확인, 최근게시물, 매칭, 오류 등을 포함

2. V8.14.6 실행 취소 문제 개선
   - GitHub Actions timeout: 30분
   - monitor 내부 실행 예산: 25분
   - 25분 내 처리 결과를 저장하고 미처리 기관은 재시도 대상으로 기록

3. V8.14.6 집계 오류 방어
   - 숫자형 aggregate에 list/set 등의 자료형이 들어가도 TypeError가 발생하지 않도록 방어

4. 기존 기능 유지
   - 355기관 / 최근 30일
   - 설문조사 / 시민참여 / 국민참여
   - 제목·본문 검색
   - 공모전 제목 제외
   - Telegram 최대 20건
   - URL + fingerprint 중복 방지
   - checked_posts / board_cache / retry_queue
   - 기관유형별 요약
