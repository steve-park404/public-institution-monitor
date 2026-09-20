# V8.14.2

V8.14.1 기반 안정화 + 티끌 모니터링 요약.

## 추가 기능
- 기관유형을 `기관유형` 열에서 읽어 공기업/준정부기관/기타공공기관으로 집계
- 공기업 30 + 준정부기관 58 = 88개를 합산한 별도 통계
- 전체/기관유형별 정상 확인, 미확인, 확인율
- 미확인 기관명을 `diagnostics.json`에 저장
- NO_CANDIDATE / HOME_ERROR / CANDIDATE_NOT_VERIFIED 구분
- Telegram으로 매일 모니터링 요약 1건 추가 발송
- 요약 메시지는 참여정보 20건 발송 상한과 별도로 전송
- `post_candidates` 집계를 실제 처리 수와 일치하도록 보강
- 기존 V8.14.1 중복방지/발송이력 승계/충돌 대응 유지

## 중요
기관 목록 Excel에 `기관유형` 열이 있으면 정확한 기관유형별 통계가 생성됩니다.
`diagnostics.json`에서 `institution_type_summary`와 `unconfirmed_institutions`를 확인할 수 있습니다.

Workflow 파일명:
`.github/workflows/v8_14_2_355_monitor.yml`
