V8.14.13

목적
- V8.14.12 기관별 상태 CSV 분석 결과를 반영한 게시판 발견률 개선
- 1차 홈페이지 링크 후보와 2차 sitemap/robots/common-path 후보를 분리 진단
- 2차 탐색 후보 수를 6개에서 12개로 확대
- 기관별_상태.csv에 게시판 탐색 과정을 세분화 기록
- GitHub Actions contents: write 명시

기존 기능 유지
- 최근 30일
- 설문조사/시민참여/국민참여
- 제목/본문 검색
- 제목 공모전 제외
- Telegram 일일 20건 제한
- URL + fingerprint 중복 방지
- checked_posts 중복 상세조회 방지
- daily summary 일일 중복 방지
- 기관별_상태.csv 생성 및 커밋

추가 CSV 진단 컬럼
- 1차게시판후보수
- 2차게시판후보수
- 총게시판후보수
- 1차조회수
- 2차조회수
- 1차검증수
- 2차검증수

권장 확인 순서
1. GitHub에 V8.14.13 파일 반영
2. Actions에서 workflow_dispatch 실행
3. 기관별_상태.csv에서 CANDIDATE_NOT_VERIFIED 72개와 HOME_ERROR 47개 확인
4. 2차게시판후보수/2차검증수 분포 확인
5. 검증 성공 기관 수가 증가하는지 비교
