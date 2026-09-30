# -*- coding: utf-8 -*-
from pathlib import Path
import re

p = Path("monitor.py")
if not p.exists():
    raise SystemExit("monitor.py가 없습니다.")

s = p.read_text(encoding="utf-8")

s = re.sub(r'(?m)^VERSION\s*=\s*"V[^"]+"',
           'VERSION = "V8.25"', s, count=1)

KEYWORDS = [
    "설문", "의견수렴", "시민참여", "국민참여",
    "공모전", "퀴즈 이벤트", "평가단 모집", "만족도 조사"
]
s = re.sub(r'(?m)^KEYWORDS\s*=\s*\[[^\n]*\]',
           f'KEYWORDS = {KEYWORDS!r}', s, count=1)

s = re.sub(r'(?m)^HTTP_RETRIES\s*=\s*\d+',
           'HTTP_RETRIES = 2', s, count=1)

RESULT_PATTERNS = [
    "당첨자", "결과발표", "결과 발표",
    "조사결과", "조사 결과", "설문결과", "설문 결과",
    "선정결과", "선정 결과", "참여자 발표",
    "수상자", "수상작", "수상자 발표", "수상작 발표",
    "당선자", "당선작", "입상자", "입상작",
    "결과 안내", "결과안내",
    "최종 결과", "최종결과",
    "심사 결과", "심사결과",
    "결과보고", "결과 보고"
]
s = re.sub(r'(?ms)^RESULT_TITLE_PATTERNS\s*=\s*\[.*?\n\]',
           'RESULT_TITLE_PATTERNS = ' + repr(RESULT_PATTERNS),
           s, count=1)

SURVEY_CONTEXT = [
    "설문 참여", "설문에 참여", "설문조사 참여", "설문 조사 참여",
    "설문 응답", "설문에 응답", "설문기간", "설문 기간",
    "설문대상", "설문 대상", "조사 참여", "조사에 참여",
    "조사기간", "조사 기간", "조사대상", "조사 대상",
    "응답기간", "응답 기간", "응답방법", "응답 방법",
    "응답자", "온라인 설문", "온라인 설문조사",
    "설문지", "설문 링크", "설문링크",
    "참여기간", "참여 기간", "참여방법", "참여 방법",
    "참여해 주세요", "참여해주시기 바랍니다",
    "응답해 주세요", "응답해주시기 바랍니다",
    "의견 제출", "의견수렴", "의견조사", "의견 조사",
    "의견을 제출", "의견을 남겨",
    "만족도 조사 참여", "만족도조사 참여",
    "만족도 조사 응답", "만족도조사 응답",
    "만족도 조사기간", "만족도 조사 기간",
    "만족도 조사대상", "만족도 조사 대상",
    "만족도 조사 참여기간", "만족도 조사 참여 기간"
]
s = re.sub(r'(?ms)^SURVEY_ACTION_CONTEXT\s*=\s*\[.*?\n\]',
           'SURVEY_ACTION_CONTEXT = ' + repr(SURVEY_CONTEXT),
           s, count=1)

start = s.find("def meaningful_body_match(")
end = s.find("def is_result_or_announcement_title", start)

if start != -1 and end != -1:
    new_func = """def meaningful_body_match(body, keyword):
    # V8.25: 실제 참여행동 문맥이 있는 경우만 본문 매칭.
    body = norm(body)
    if len(body) < 120:
        return False

    context_map = {
        "공모전": CONTEST_ACTION_CONTEXT,
        "퀴즈 이벤트": QUIZ_EVENT_ACTION_CONTEXT,
        "평가단 모집": EVALUATOR_ACTION_CONTEXT,
    }

    survey_keywords = {"설문", "의견수렴", "만족도 조사"}

    for m in re.finditer(re.escape(keyword), body, re.I):
        idx = m.start()
        ctx = body[max(0, idx - 700):min(len(body), idx + 1000)]

        if keyword in survey_keywords:
            hits = [x for x in SURVEY_ACTION_CONTEXT if x in ctx]
            if not hits:
                continue

            strong = [
                "참여", "응답", "기간", "대상", "방법",
                "설문지", "링크", "의견 제출", "의견수렴",
                "의견조사", "모집", "신청"
            ]
            if not any(x in ctx for x in strong):
                continue

            noise = sum(1 for x in GENERIC_BODY_NOISE if x in ctx)
            if noise >= 4 and len(set(hits)) < 2:
                continue

            return True

        elif keyword in ("국민참여", "시민참여"):
            action_hits = [
                x for x in PARTICIPATION_ACTION_SIGNALS if x in ctx
            ]
            if not action_hits:
                continue

            strong = [x for x in PARTICIPATION_CONTEXT if x in ctx]
            if len(set(strong)) < 2 and not any(
                x in ctx for x in [
                    "설문", "의견", "응답", "모집",
                    "신청", "제안", "조사"
                ]
            ):
                continue

            noise = sum(1 for x in GENERIC_BODY_NOISE if x in ctx)
            if noise >= 3 and len(action_hits) < 2:
                continue

            return True

        else:
            signals = context_map.get(keyword, [])
            hits = [x for x in signals if x in ctx]
            if not hits:
                continue

            if keyword == "공모전" and not any(
                x in ctx for x in [
                    "접수", "응모", "출품", "신청",
                    "모집", "참가", "작품"
                ]
            ):
                continue

            if keyword == "퀴즈 이벤트" and not any(
                x in ctx for x in [
                    "참여", "응모", "정답",
                    "문제", "기간", "방법"
                ]
            ):
                continue

            if keyword == "평가단 모집" and not any(
                x in ctx for x in [
                    "신청", "모집", "지원",
                    "기간", "방법", "선정"
                ]
            ):
                continue

            noise = sum(1 for x in GENERIC_BODY_NOISE if x in ctx)
            if noise >= 3 and len(set(hits)) < 2:
                continue

            return True

    return False

"""
    s = s[:start] + new_func + s[end:]

old = """    body=visible_main_text(soup)
    # 게시물 검증은 기존보다 완화한다. 상세 URL + 날짜 + 제목 + 내용이 있으면 통과시키고,
    # 실제 매칭 여부는 아래 제목/본문 단계에서 판단한다.
    if not body or len(body)<60:
        if not has_post_structure(soup,title): return None,"NOT_POST_STRUCTURE",""
"""
new = """    body=visible_main_text(soup)

    # V8.25: 메인/공통 페이지 문구를 게시물로 오인하지 않도록 차단
    site_like_phrases = [
        "로고 메인페이지로 이동",
        "메인페이지로 이동",
        "홈페이지로 이동",
        "사이트맵",
        "개인정보처리방침",
        "이용약관",
    ]
    if any(x in title for x in site_like_phrases):
        return None,"GENERIC_TITLE",""

    if not body or len(body)<60:
        if not has_post_structure(soup,title):
            return None,"NOT_POST_STRUCTURE",""

    if body:
        noise = sum(
            1 for x in GENERIC_BODY_NOISE
            if x in body[:8000]
        )
        if len(body) < 160 and noise >= 3:
            return None,"GENERIC_PAGE",""
"""
s = s.replace(old, new, 1)

p.write_text(s, encoding="utf-8")

print("V8.25 patch applied")
print("8개 키워드:", ", ".join(KEYWORDS))
print("HTTP_RETRIES=2")
