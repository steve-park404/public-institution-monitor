티끌 알림봇 - 747기관 단일 Workflow 최종 교체용

구성
- monitor_targets.xlsx : 747기관 통합 대상. 수동 입력한 공지게시판URL을 그대로 사용합니다.
- monitor_patch.py : V8.27.1 정밀 필터 + 8개 키워드 + Telegram 오늘 모니터링 요약 패치입니다.
- .github/workflows/tikkle_monitor.yml : 단일 일일 실행 Workflow입니다.

핵심 수정
1. Workflow는 오직 monitor_targets.xlsx만 읽습니다.
2. 723기관 예전 monitor_targets_통합.xlsx를 자동 선택하지 않습니다.
3. 실행 시작 시 기관 수를 정확히 747개인지 검사합니다. 747개가 아니면 즉시 중단합니다.
4. 교육지원청 176개는 실제 모니터링 대상에서 제외되어 있습니다.
5. 기존 수동 입력 공지게시판URL을 유지합니다.
6. Telegram 요약은 '오늘 모니터링'만 표시합니다. 전체기관/공기업·준정부기관/기타공공기관/게시판 미확인/미확인 기관 예시는 표시하지 않습니다.
7. 기존 V8.27.1의 결과발표·수상작·후보작 등 오탐 차단 로직을 유지합니다.

GitHub 적용
- 저장소 루트: monitor_targets.xlsx, monitor_patch.py
- .github/workflows/: tikkle_monitor.yml
- Actions에서 'Tikkle Daily Monitoring - 747 Targets' 실행
- 첫 로그가 반드시 '통합 대상: 747개'여야 합니다.

주의
기존 V8.xx Workflow들은 중복 실행 방지를 위해 비활성화하거나 삭제하고 tikkle_monitor.yml 하나만 운영하는 것을 권장합니다.
