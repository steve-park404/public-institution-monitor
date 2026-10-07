# -*- coding: utf-8 -*-
"""Tikkle V8.27.1 precision patch + missed-opportunity diagnostics.

This patch is designed to run against the existing monitor.py in the repository.
It keeps the existing crawler/state logic and only tightens the classification and
adds a preflight-friendly set of helpers.
"""
from pathlib import Path
import re

p = Path('monitor.py')
if not p.exists():
    raise SystemExit('monitor.py가 없습니다. 기존 저장소의 monitor.py를 유지한 상태에서 실행하세요.')
s = p.read_text(encoding='utf-8')

# Version / fixed target workbook / keywords
s = re.sub(r'(?m)^VERSION\s*=\s*"V[^"]+"', 'VERSION = "V8.27.1"', s, count=1)
keywords = ["설문","의견수렴","시민참여","국민참여","공모전","퀴즈 이벤트","평가단 모집","만족도 조사"]
s, n = re.subn(r'(?m)^KEYWORDS\s*=\s*\[[^\n]*\]', 'KEYWORDS = ' + repr(keywords), s, count=1)
if n != 1:
    raise SystemExit('KEYWORDS 패치를 찾지 못했습니다.')
s, n = re.subn(r'(?m)^TARGET_FILE_CANDIDATES\s*=\s*\[[^\n]*\]', 'TARGET_FILE_CANDIDATES = ["monitor_targets.xlsx"]', s, count=1)
if n != 1:
    raise SystemExit('TARGET_FILE_CANDIDATES 패치를 찾지 못했습니다.')

# Result-stage titles must remain excluded.
result_patterns = [
    "당첨자","결과발표","결과 발표","조사결과","조사 결과","설문결과","설문 결과",
    "선정결과","선정 결과","참여자 발표","수상자","수상작","수상자 발표","수상작 발표",
    "당선자","당선작","입상자","입상작","결과 안내","결과안내","최종 결과","최종결과",
    "심사 결과","심사결과","결과보고","결과 보고","후보작","수상 후보","공개 검증","검증 결과",
    "선정자 발표","선정자","선정작 발표"
]
s, n = re.subn(r'(?ms)^RESULT_TITLE_PATTERNS\s*=\s*\[.*?\n\]', 'RESULT_TITLE_PATTERNS = ' + repr(result_patterns), s, count=1)
if n != 1:
    raise SystemExit('RESULT_TITLE_PATTERNS 패치를 찾지 못했습니다.')

# Extra generic-shell title patterns.
shell_patterns = ["메인","메인페이지","홈페이지","사이트맵","사이트 맵","대학생활","SEOUL TOURISM ORGANIZATION"]
if 'GENERIC_SHELL_TITLE_PATTERNS' not in s:
    marker = 'GENERIC_URL_HINTS = ['
    if marker in s:
        s = s.replace(marker, 'GENERIC_SHELL_TITLE_PATTERNS = ' + repr(shell_patterns) + '\n' + marker, 1)

# Replace the strict body matcher from V8.27.1.
start = s.find('def meaningful_body_match(body,keyword):')
end = s.find('def is_result_or_announcement_title(title):', start)
if start < 0 or end < 0:
    raise SystemExit('meaningful_body_match 구간을 찾지 못했습니다.')
new_body = '''def meaningful_body_match(body,keyword):
    """V8.27.1: 본문 키워드는 실제 참여행동 문맥이 있을 때만 통과."""
    body=norm(body)
    if len(body)<120: return False
    context_map={
        "공모전": CONTEST_ACTION_CONTEXT,
        "퀴즈 이벤트": QUIZ_EVENT_ACTION_CONTEXT,
        "평가단 모집": EVALUATOR_ACTION_CONTEXT,
    }
    survey_keywords={"설문","설문조사","만족도 조사","의견수렴"}
    for m in re.finditer(re.escape(keyword), body, re.I):
        idx=m.start(); ctx=body[max(0,idx-500):min(len(body),idx+750)]
        if any(x in ctx for x in RESULT_TITLE_PATTERNS): continue
        if keyword in ("국민참여","시민참여"):
            action=[x for x in PARTICIPATION_ACTION_SIGNALS if x in ctx]
            strong=[x for x in PARTICIPATION_CONTEXT if x in ctx]
            if not action: continue
            if len(set(strong))<2 and not any(x in ctx for x in ["설문","의견","응답","모집","신청","제안","조사"]): continue
            if sum(1 for x in GENERIC_BODY_NOISE if x in ctx)>=3 and len(set(action))<2: continue
            return True
        elif keyword in survey_keywords:
            action=[x for x in SURVEY_ACTION_CONTEXT + PARTICIPATION_CONTEXT if x in ctx]
            if not action: continue
            if not any(x in ctx for x in ["참여","응답","기간","대상","방법","설문지","링크","의견"]): continue
            if sum(1 for x in GENERIC_BODY_NOISE if x in ctx)>=3 and len(set(action))<2: continue
            return True
        else:
            signals=context_map.get(keyword,[])
            hits=[x for x in signals if x in ctx]
            if not hits: continue
            if keyword=="공모전":
                strong=["공모전 개최","공모전 참가","공모전 접수","공모전 응모","공모전 신청","공모전 모집","공모전 참여","공모전 출품","아이디어 공모","작품 공모","참가 신청","작품 제출"]
                if any(x in ctx for x in strong): return True
                if len(set(x for x in ["접수","응모","출품","신청","모집","참가","참여","제출","기간","방법"] if x in ctx))<2: continue
            elif keyword=="퀴즈 이벤트":
                if not any(x in ctx for x in ["참여","응모","정답","문제","기간","방법"]): continue
            elif keyword=="평가단 모집":
                if not any(x in ctx for x in ["신청","모집","지원","기간","방법","선정"]): continue
            if sum(1 for x in GENERIC_BODY_NOISE if x in ctx)>=3 and len(set(hits))<2: continue
            return True
    return False

'''
s = s[:start] + new_body + s[end:]

# Site-title false positives: keep result filtering, but don't classify real opportunity titles as shell pages.
old = '''    if len(tc)<=10 and pc and (pc.startswith(tc) or pc.endswith(tc) or tc in pc):
        return True
    return False
'''
new = '''    if len(tc)<=10 and pc and (pc.startswith(tc) or pc.endswith(tc) or tc in pc):
        return True
    if any(x.lower() in t.lower() for x in GENERIC_SHELL_TITLE_PATTERNS):
        return True
    if pc and (tc == pc or tc in pc) and len(tc) <= 100:
        # Opportunity-like titles should not be discarded merely because the page title repeats itself.
        if not any(k in t for k in ["공모전","설문","의견수렴","시민참여","국민참여","퀴즈","평가단","만족도"]):
            return True
    return False
'''
if old in s:
    s = s.replace(old, new, 1)

# More common index URL hints.
s = s.replace('("/index.do","/index.jsp","/main.do","/main.jsp")', '("/index.do","/index.jsp","/index.krc","/main.do","/main.jsp","/main.krc")')

# Add explicit title-level opportunity helper. This is used by the workflow smoke test and
# can also be used by existing matching code when present.
helper = r'''

# --- Tikkle missed-opportunity guard ---
TITLE_DIRECT_OPPORTUNITY_KEYWORDS = ["공모전", "설문", "설문조사", "의견수렴", "시민참여", "국민참여", "퀴즈 이벤트", "평가단 모집", "만족도 조사"]
TITLE_RESULT_EXCLUSIONS = RESULT_TITLE_PATTERNS

def direct_title_opportunity_match(title, keyword=None):
    """Accept a real opportunity title directly, while preserving result-stage exclusions."""
    t = norm(title or "")
    if not t:
        return False
    if any(x in t for x in TITLE_RESULT_EXCLUSIONS):
        return False
    keys = [keyword] if keyword else TITLE_DIRECT_OPPORTUNITY_KEYWORDS
    keys = [k for k in keys if k]
    if not any(k in t for k in keys):
        return False
    # For contest/survey titles, action or recruitment language is preferred.
    if any(k in t for k in ["공모전", "설문", "설문조사", "의견수렴", "시민참여", "국민참여", "만족도 조사"]):
        return any(x in t for x in ["개최","모집","공고","안내","신청","접수","참여","응모","기간","조사","설문","의견","수렴","실시","진행"])
    return any(x in t for x in ["모집","신청","참여","응모","개최","안내","접수","기간"])
'''
if 'def direct_title_opportunity_match(' not in s:
    s += helper

# Telegram summary: TODAY ONLY. Replace if the function exists.
summary_start = s.find('def build_monitoring_summary(')
summary_end = s.find('def telegram_send(item):', summary_start)
if summary_start >= 0 and summary_end >= 0:
    compact_summary = '''def build_monitoring_summary(targets, results, posts_checked, new_matches,
                             sent, pending_after, errors, aggregate,
                             status_counts, sent_url_ledger):
    """Telegram에는 오늘 모니터링 핵심 수치만 표시한다."""
    return "\\n".join([
        "📊 티끌 모니터링",
        "━━━━━━━━━━━━━━",
        f"📅 {datetime.now(KST).strftime('%Y-%m-%d')}",
        "",
        "🔎 오늘 모니터링",
        f"• 게시물 확인 {posts_checked:,}건",
        f"• 실제 최근 게시물 {aggregate.get('recent_posts',0):,}건",
        f"• 신규 키워드 매칭 {new_matches}건",
        f"• Telegram 참여정보 발송 {sent}건",
        f"• 대기 {pending_after}건",
    ])

'''
    s = s[:summary_start] + compact_summary + s[summary_end:]

p.write_text(s, encoding='utf-8')
print('Tikkle precision patch written')
