티끌 알림봇 V8.27.1 — 747기관 설문참여 확장판

중요: GitHub 저장소의 기존 monitor_targets_통합.xlsx를 반드시 이 ZIP의 파일로 교체하세요.
이번 실행 실패 원인은 GitHub에 723기관 엑셀이 남아 있었기 때문입니다.

구성
- monitor_targets_통합.xlsx : 747기관
- V8.27.1_patch.py : 기존 monitor.py에 정밀 필터/8키워드/오늘 요약을 적용
- .github/workflows/V8.27.1_747기관_정밀_일일모니터링.yml : 747기관 검증 후 실행

반영 내용
- 교육지원청 176개는 모니터링 대상에서 제외
- 사용자가 수동 입력한 공지게시판URL 유지
- 8개 키워드 유지
- 공모전 결과발표/수상작 등 오탐 차단 유지
- Telegram 요약은 오늘 모니터링만 표시

GitHub 적용
1. monitor_targets_통합.xlsx를 저장소 루트에 업로드하여 기존 파일을 교체
2. V8.27.1_patch.py를 저장소 루트에 업로드/교체
3. .github/workflows/V8.27.1_747기관_정밀_일일모니터링.yml을 .github/workflows/에 업로드
4. 기존 723기관용 워크플로는 중복 실행 방지를 위해 비활성화/삭제 권장
5. Actions에서 해당 워크플로를 Run workflow

실행 전 검증
워크플로가 다음처럼 표시되어야 합니다.
통합 대상: 747개
수동 입력 공지게시판 URL: 728개 내외
