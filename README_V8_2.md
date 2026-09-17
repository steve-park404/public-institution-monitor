# V8.2 공지사항 자동검색
V8.1의 30분 timeout 문제를 해결한 버전입니다.

- 매일 홈페이지 355개를 탐색하지 않음
- `monitor_targets.xlsx`의 검증 게시판만 매일 검색
- 공지사항 게시판을 우선 대상으로 사용
- 게시물 최대 15개 확인
- 제목 + 본문 검색
- 사이트별 15초 timeout, 1회 재시도
- 전체 검색 10분 안전 제한
- Telegram 신규 매칭만 발송
- 첫 실행은 기존 매칭을 발송하지 않고 상태 등록

실제 GitHub Actions 파일은 `.github/workflows/v8_monitor.yml`입니다.
루트 `v8_monitor.yml`은 보기 편한 복사본입니다.
