Tikkle V8.27.1 – 747 targets / missed-opportunity precision update

Files
- monitor_targets.xlsx : 747 monitoring targets; manually maintained notice-board URLs preserved.
- monitor_patch.py      : applies V8.27.1 keyword/result filtering, direct title opportunity guard, and compact Telegram summary.
- .github/workflows/tikkle_monitor.yml : single daily workflow.

Important
- The repository must retain the existing monitor.py. This package intentionally does not replace the crawler/state engine.
- The workflow reads ONLY monitor_targets.xlsx.
- Education support offices (176) are not active monitoring targets.

Missed-opportunity fix
- A title such as "[공모] 2026 우리 임산물 숲푸드 콘텐츠 공모전 개최(~10.25)" is accepted directly when it is not a result-stage announcement.
- Result-stage titles such as 수상작/수상 후보/최종 결과 발표 remain excluded.
- A KOFPI preflight check confirms that the official board is reachable and that the known test title is present before the monitor runs.

Telegram summary
Only the following are shown:
- 오늘 모니터링
- 게시물 확인
- 실제 최근 게시물
- 신규 키워드 매칭
- Telegram 참여정보 발송
- 대기
