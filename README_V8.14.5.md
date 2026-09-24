# 티끌 알림봇 V8.14.5

V8.14.4를 기반으로 실행시간과 전체 처리 신뢰성을 개선한 버전입니다.

## 핵심 변경
- 기존 `board_cache.json`, `state.json`, `pending.json`, fingerprint 이력 승계
- 캐시된 게시판은 재탐색하지 않고 캐시 URL을 우선 사용
- 게시판 후보 탐색량 축소: 1차 후보 최대 12개, 2차 후보 최대 6개
- HTTP timeout 8초 / 재시도 1회
- 동시 처리 25개
- 최근 게시물 후보 최대 15개
- `HOME_ERROR`, `CANDIDATE_NOT_VERIFIED`, `NO_CANDIDATE`를 `retry_queue.json`에 별도 기록
- 모니터링 본 처리와 재시도 이력 분리
- 355개 기관 전체가 처리되지 않으면 `timed_out=true` 또는 `unprocessed_total>0`으로 진단하고 GitHub Actions를 실패 처리
- timeout인데 GitHub Actions가 성공으로 표시되는 문제 방지
- 미완료 실행은 당일 요약을 완료 처리하지 않아 다음 정상 실행에서 다시 요약 발송 가능
- URL + fingerprint 중복방지 유지
- 하루 요약 중복방지 유지

## 목표
- 355기관 전체 처리 여부를 명확하게 확인
- 캐시 기관의 불필요한 게시판 재탐색 감소
- 실행시간 단축
- 미확인 기관은 retry_queue에 누적하여 추적

## 운영
- 워크플로: `.github/workflows/v8_14_5_355_monitor.yml`
- 매일 09:10 KST
- 수동 실행 가능
- 참여정보 최대 20건
