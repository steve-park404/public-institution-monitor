# 티끌 알림봇 V8.14.4

V8.14.3을 기반으로 URL 변경형 중복과 게시물 제목 추출 문제를 보완한 안정화 버전입니다.

## 핵심 변경
- URL 중복방지 + 내용 기반 fingerprint 중복방지
- 기관명 + 게시일 + 실제 제목 + 본문 기반 SHA-256 fingerprint
- 기존 `sent_urls`, `seen` 상태 승계
- `sent_fingerprints`, `seen_fingerprints` 신규 영구 저장
- 같은 게시물이 실행마다 다른 URL을 받아도 재발송 차단
- 같은 날짜에 workflow를 수동 재실행해도 동일 알림 재발송 차단
- 하루 요약 Telegram은 KST 날짜당 1회만 발송
- 실제 게시물 제목 추출 우선순위 강화(h1/게시물 제목 영역/og:title/제목 테이블)
- 공통 사이트 제목만 있는 경우 제목 오탐 차단
- 게시물 검증을 완화하여 상세 URL + 날짜 + 제목 + 본문이 있으면 통과
- `board_cache.json` 우선 사용 및 기존 캐시 승계
- V8.14.2의 355기관/기관유형/최근30일/3개 키워드/공모전 제외/Telegram/pending 기능 유지

## 운영
- 워크플로: `.github/workflows/v8_14_4_355_monitor.yml`
- 실행: 매일 09:10 KST + 수동 실행 가능
- 요약은 하루 1회
- 참여정보는 신규 fingerprint 기준으로 최대 20건

## 주의
V8.14.4 최초 실행 시 기존 URL 발송 이력은 그대로 승계합니다. 과거에 URL만 저장되어 있고 fingerprint가 없었던 게시물은 실제 페이지 내용 확인 후 fingerprint가 새로 등록됩니다.
