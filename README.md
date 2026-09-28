# 티끌 알림봇 V8.14.21

V8.14.20의 게시판 탐색/캐시/중복방지/Telegram 판정 구조를 유지하고,
**상세페이지 판정 결과를 `진단_후보.csv`로 기록**하는 진단 버전입니다.

## 목적
V8.14.20에서 `new_matches=0`이 나온 상황에서:
- 실제 참여/설문 글을 너무 엄격하게 제외하고 있는지
- 메뉴/푸터/홈페이지 오탐을 제대로 차단하고 있는지
를 다음 실행에서 확인합니다.

## 핵심 변경
- 버전: V8.14.21
- 게시판 탐색 로직: 변경하지 않음
- Telegram 발송 기준: V8.14.20 유지
- 추가: `진단_후보.csv`
- 각 상세페이지의 최종 판정(`NO_KEYWORD`, `GENERIC_PAGE`, `RESULT_ANNOUNCEMENT_TITLE`, `BODY_MATCH` 등)을 기록
- `diagnostics.json`에 `reason_counts` 추가

## 확인 방법
GitHub Actions 실행 후 Artifacts에서 `V8.14.21-results`를 내려받아
`진단_후보.csv`를 확인합니다.

특히 다음 판정을 봅니다.
- `BODY_MATCH`: 실제 알림 후보
- `TITLE_MATCH`: 제목에서 직접 매칭
- `NO_KEYWORD`: 상세페이지까지 확인했지만 키워드 조건 미충족
- `RESULT_ANNOUNCEMENT_TITLE`: 당첨/결과 발표로 제외
- `GENERIC_PAGE`: 홈페이지/공통 페이지로 제외
- `LIST_PAGE`: 게시물 상세페이지가 아닌 목록으로 제외

이번 버전은 **진단용**이므로 실제 Telegram 발송 결과가 0건이어도 오류가 아닙니다.
