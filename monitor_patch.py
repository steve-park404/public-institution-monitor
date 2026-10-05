# -*- coding: utf-8 -*-
"""V8.27.1 hotfix for V8.27 keyword/config verification failure."""
from pathlib import Path
import re

p=Path('monitor.py')
if not p.exists(): raise SystemExit('monitor.py가 없습니다.')
s=p.read_text(encoding='utf-8')
s=re.sub(r'(?m)^VERSION\s*=\s*"V[^"]+"','VERSION = "V8.27.1"',s,count=1)

# Keep the user's current 8-keyword policy exactly.
keywords=["설문","의견수렴","시민참여","국민참여","공모전","퀴즈 이벤트","평가단 모집","만족도 조사"]
s,n=re.subn(r'(?m)^KEYWORDS\s*=\s*\[[^\n]*\]', 'KEYWORDS = '+repr(keywords), s, count=1)
if n!=1: raise SystemExit('KEYWORDS 패치를 찾지 못했습니다.')

# Use the manually maintained integrated workbook first.
s,n=re.subn(r'(?m)^TARGET_FILE_CANDIDATES\s*=\s*\[[^\n]*\]', 'TARGET_FILE_CANDIDATES = ["monitor_targets.xlsx", "monitor_targets_통합.xlsx", "monitor_targets_통합_지자체243추가.xlsx", "monitor_targets_통합_지자체243추가 (4).xlsx", "monitor_targets.xlsx", "url.xlsx", "targets.xlsx"]', s, count=1)
if n!=1: raise SystemExit('TARGET_FILE_CANDIDATES 패치를 찾지 못했습니다.')

result_patterns=[
"당첨자","결과발표","결과 발표","조사결과","조사 결과","설문결과","설문 결과",
"선정결과","선정 결과","참여자 발표","수상자","수상작","수상자 발표","수상작 발표",
"당선자","당선작","입상자","입상작","결과 안내","결과안내","최종 결과","최종결과",
"심사 결과","심사결과","결과보고","결과 보고","후보작","수상 후보","공개 검증","검증 결과",
"선정자 발표","선정자","선정작 발표"
]
s,n=re.subn(r'(?ms)^RESULT_TITLE_PATTERNS\s*=\s*\[.*?\n\]', 'RESULT_TITLE_PATTERNS = '+repr(result_patterns), s, count=1)
if n!=1: raise SystemExit('RESULT_TITLE_PATTERNS 패치를 찾지 못했습니다.')

shell_patterns=["메인","메인페이지","홈페이지","사이트맵","사이트 맵","대학생활","SEOUL TOURISM ORGANIZATION"]
if 'GENERIC_SHELL_TITLE_PATTERNS' not in s:
    marker='GENERIC_URL_HINTS = ['
    s=s.replace(marker,'GENERIC_SHELL_TITLE_PATTERNS = '+repr(shell_patterns)+'\n'+marker,1)

# Replace body matcher with strict V8.27 logic, supporting both old/new survey keyword names.
start=s.find('def meaningful_body_match(body,keyword):')
end=s.find('def is_result_or_announcement_title(title):',start)
if start<0 or end<0: raise SystemExit('meaningful_body_match 구간을 찾지 못했습니다.')
new='''def meaningful_body_match(body,keyword):\n    """V8.27.1: 본문 키워드는 실제 참여행동 문맥이 있을 때만 통과."""\n    body=norm(body)\n    if len(body)<120: return False\n    context_map={\n        "공모전": CONTEST_ACTION_CONTEXT,\n        "퀴즈 이벤트": QUIZ_EVENT_ACTION_CONTEXT,\n        "평가단 모집": EVALUATOR_ACTION_CONTEXT,\n    }\n    survey_keywords={"설문","설문조사","만족도 조사","의견수렴"}\n    for m in re.finditer(re.escape(keyword), body, re.I):\n        idx=m.start(); ctx=body[max(0,idx-500):min(len(body),idx+750)]\n        if any(x in ctx for x in RESULT_TITLE_PATTERNS): continue\n        if keyword in ("국민참여","시민참여"):\n            action=[x for x in PARTICIPATION_ACTION_SIGNALS if x in ctx]\n            strong=[x for x in PARTICIPATION_CONTEXT if x in ctx]\n            if not action: continue\n            if len(set(strong))<2 and not any(x in ctx for x in ["설문","의견","응답","모집","신청","제안","조사"]): continue\n            if sum(1 for x in GENERIC_BODY_NOISE if x in ctx)>=3 and len(set(action))<2: continue\n            return True\n        elif keyword in survey_keywords:\n            action=[x for x in SURVEY_ACTION_CONTEXT + PARTICIPATION_CONTEXT if x in ctx]\n            if not action: continue\n            if not any(x in ctx for x in ["참여","응답","기간","대상","방법","설문지","링크","의견"]): continue\n            if sum(1 for x in GENERIC_BODY_NOISE if x in ctx)>=3 and len(set(action))<2: continue\n            return True\n        else:\n            signals=context_map.get(keyword,[])\n            hits=[x for x in signals if x in ctx]\n            if not hits: continue\n            if keyword=="공모전":\n                strong=["공모전 개최","공모전 참가","공모전 접수","공모전 응모","공모전 신청","공모전 모집","공모전 참여","공모전 출품","아이디어 공모","작품 공모","참가 신청","작품 제출"]\n                if any(x in ctx for x in strong): return True\n                if len(set(x for x in ["접수","응모","출품","신청","모집","참가","참여","제출","기간","방법"] if x in ctx))<2: continue\n            elif keyword=="퀴즈 이벤트":\n                if not any(x in ctx for x in ["참여","응모","정답","문제","기간","방법"]): continue\n            elif keyword=="평가단 모집":\n                if not any(x in ctx for x in ["신청","모집","지원","기간","방법","선정"]): continue\n            if sum(1 for x in GENERIC_BODY_NOISE if x in ctx)>=3 and len(set(hits))<2: continue\n            return True\n    return False\n\n'''
s=s[:start]+new+s[end:]

# Add shell-title false-positive checks to existing function.
old='''    if len(tc)<=10 and pc and (pc.startswith(tc) or pc.endswith(tc) or tc in pc):\n        return True\n    return False\n'''
new2='''    if len(tc)<=10 and pc and (pc.startswith(tc) or pc.endswith(tc) or tc in pc):\n        return True\n    if any(x.lower() in t.lower() for x in GENERIC_SHELL_TITLE_PATTERNS):\n        return True\n    if pc and (tc == pc or tc in pc) and len(tc) <= 100:\n        return True\n    return False\n'''
if old not in s: raise SystemExit('is_probably_site_title 본문을 찾지 못했습니다.')
s=s.replace(old,new2,1)

# Generic homepage URL patterns.
s=s.replace('("/index.do","/index.jsp","/main.do","/main.jsp")','("/index.do","/index.jsp","/index.krc","/main.do","/main.jsp","/main.krc")')

p.write_text(s,encoding='utf-8')
print('V8.27.1 patch applied')

# --- V8.27.1 user-requested Telegram summary: TODAY ONLY ---

summary_start = s.find("def build_monitoring_summary(")
summary_end = s.find("def telegram_send(item):", summary_start)
if summary_start < 0 or summary_end < 0:
    raise SystemExit("build_monitoring_summary/telegram_send 구간을 찾지 못했습니다.")
compact_summary = """def build_monitoring_summary(targets, results, posts_checked, new_matches,
                             sent, pending_after, errors, aggregate,
                             status_counts, sent_url_ledger):
    """ + '"""Telegram 요약은 사용자가 요청한 오늘 모니터링만 표시한다."""' + r"""
    return "\n".join([
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
"""
s = s[:summary_start] + compact_summary + s[summary_end:]
