Tikkle Alert Bot V8.28 - 747 targets / TITLE-ONLY keyword matching

Files in this package:
- .github/workflows/tikkle_monitor.yml
- monitor_targets.xlsx
- monitor_patch.py
- README.txt

Important:
1. Keep the repository's existing monitor.py. This package patches that file; it does not replace it.
2. The active target workbook is monitor_targets.xlsx (747 targets).
3. Education support offices are kept only in the separate candidate sheet and are NOT active monitoring targets.
4. Keyword matching is TITLE ONLY. BODY keyword matching is disabled completely.
5. This prevents false positives such as a page body containing '공모전' or '설문' even when the actual post title is unrelated.
6. Generic page titles such as '이전글', '주요누리집 닫기', '메인', '뉴스/소식 : ... 시민참여 ...' are rejected.
7. Result-stage titles remain excluded: 당첨자, 결과발표, 수상작, 수상 후보, 최종 결과, 선정자 발표, etc.
8. Real opportunity titles such as '[공모] 2026 우리 임산물 숲푸드 콘텐츠 공모전 개최(~10.25)' are accepted by the title regression test.
9. Use this workflow only: .github/workflows/tikkle_monitor.yml
10. Do not run old V8.26/V8.27/V8.28 workflows in parallel.

Telegram summary is kept compact and shows only today's monitoring counts.


[키워드 추가]
제목 기반 키워드에 이벤트, 슬로건, 표어를 추가했습니다. 본문(BODY) 매칭은 계속 사용하지 않습니다.
