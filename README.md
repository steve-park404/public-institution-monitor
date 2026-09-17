# V8.3 결과파일 자동보관 수정본

검증 로직은 변경하지 않았습니다.

V8.3 실행이 끝나면 결과 엑셀을:

1. GitHub Actions → 해당 실행 → Artifacts → `v8_3_validation_result`
2. GitHub 저장소 루트 → `v8_3_notice_validation.xlsx`

두 곳에 보관합니다.

Artifacts는 30일간 보관됩니다.

## 적용
저장소에 기존 V8.3 파일을 유지한 채,
`.github/workflows/v8_3_notice_validation.yml`을 이 파일로 교체하세요.

그리고 Actions → V8.3 공지사항 후보 자동검증 → Run workflow를 실행합니다.
