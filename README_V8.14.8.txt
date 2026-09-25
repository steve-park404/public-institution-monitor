# V8.14.8 - 기관별_상태.csv 생성·커밋 검증 강화

V8.14.7을 그대로 기반으로 하며, 기관별_상태.csv가 GitHub 저장소에 실제로 남도록 생성·검증·커밋 절차를 강화했습니다.

## 주요 변경
1. VERSION = V8.14.8
2. CSV 생성 직후 존재 여부와 파일 크기 검증
3. Actions 로그에 CSV 생성 행 수/파일 크기 출력
4. 별도 Verify 단계에서 CSV 존재 여부, 첫 5줄, 행 수 확인
5. git add -f로 CSV가 .gitignore에 걸려도 강제 추가
6. 커밋 전 git status / staged diff 통계 출력
7. Actions Artifact로 CSV를 14일간 별도 보관
8. 최종 커밋에 기관별_상태.csv가 실제 포함됐는지 확인
9. 기존 V8.14.7 기능 유지
   - 355기관
   - 최근 30일
   - 설문조사 / 시민참여 / 국민참여
   - 제목·본문 검색
   - 공모전 제목 제외
   - Telegram 최대 20건
   - URL + fingerprint 중복 방지
   - checked_posts / board_cache / retry_queue
   - 기관유형별 요약
   - 기관별 상태 CSV 355행
