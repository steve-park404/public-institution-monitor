# V8.5.1 오류 수정
V8.5 실행에서 발생한 `TypeError: list indices must be integers or slices, not str`를 수정했습니다.

원인: 이전 state.json의 `seen`이 리스트 형식인데 코드가 딕셔너리로 접근했습니다.
수정: 구버전 list/dict state 자동 변환, 새 state는 dict로 저장, XML 파싱 경고 억제.

**V8.5의 50개 모니터링 대상은 그대로 유지합니다.**
