# V8.23.1 수정 패키지

## 이번 오류
기존 V8.23에서 `read_only=True`로 연 추가 Excel의 `max_row`가 None인 환경에서
`ews.max_row + 1`을 실행하여 TypeError가 발생했습니다.

## 수정
`merge_extra_targets.py`를 `iter_rows()` 기반으로 변경했습니다.

따라서:
- `max_row=None` 오류 방지
- 사용자가 입력한 게시판URL만 편입
- 기존 기관과 중복되면 중복 추가하지 않음
- 기존 monitor.py 검색 로직은 변경하지 않음

## 포함 파일
- merge_extra_targets.py
- 유사기관_추가모니터링_후보목록_URL입력완료 (2).xlsx
- .github/workflows/V8.23.1_유사기관_통합모니터링.yml

기존 저장소에 기존 `monitor.py`는 그대로 둡니다.
