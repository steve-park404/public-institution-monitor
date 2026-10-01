# -*- coding: utf-8 -*-
"""V8.27 precision patch based on real V8.26 Telegram false positives."""
from pathlib import Path
import re

p = Path("monitor.py")
if not p.exists():
    raise SystemExit("monitor.py가 없습니다.")
s = p.read_text(encoding="utf-8")
s = re.sub(r'(?m)^VERSION\s*=\s*"V[^"]+"', 'VERSION = "V8.27"', s, count=1)

result_patterns = [
    "당첨자", "결과발표", "결과 발표", "조사결과", "조사 결과",
    "설문결과", "설문 결과", "선정결과", "선정 결과", "참여자 발표",
    "수상자", "수상작", "수상자 발표", "수상작 발표",
    "당선자", "당선작", "입상자", "입상작", "결과 안내", "결과안내",
    "최종 결과", "최종결과", "심사 결과", "심사결과", "결과보고", "결과 보고",
    "후보작", "수상 후보", "공개 검증", "검증 결과",
    "선정자 발표", "선정자", "선정작 발표",
]
s, n = re.subn(r'(?ms)^RESULT_TITLE_PATTERNS\s*=\s*\[.*?\n\]',
                'RESULT_TITLE_PATTERNS = ' + repr(result_patterns), s, count=1)
if n != 1:
    raise SystemExit("RESULT_TITLE_PATTERNS 패치를 찾지 못했습니다.")

shell_patterns = ["메인", "메인페이지", "홈페이지", "사이트맵", "사이트 맵", "대학생활", "SEOUL TOURISM ORGANIZATION"]
if "GENERIC_SHELL_TITLE_PATTERNS" not in s:
    s = s.replace("GENERIC_URL_HINTS = [", "GENERIC_SHELL_TITLE_PATTERNS = " + repr(shell_patterns) + "\nGENERIC_URL_HINTS = [", 1)

start = s.find('def meaningful_body_match(body,keyword):')
end = s.find('def is_result_or_announcement_title(title):', start)
if start < 0 or end < 0:
    raise SystemExit("meaningful_body_match 구간을 찾지 못했습니다.")
new_func = '''def meaningful_body_match(body,keyword):
    """V8.27: 본문 키워드는 실제 참여행동 문맥이 있을 때만 통과시킨다."""
    body=norm(body)
    if len(body)<120: return False
    context_map={
        "공모전": CONTEST_ACTION_CONTEXT,
        "퀴즈 이벤트": QUIZ_EVENT_ACTION_CONTEXT,
        "평가단 모집": EVALUATOR_ACTION_CONTEXT,
    }
    for m in re.finditer(re.escape(keyword), body, re.I):
        idx=m.start()
        ctx=body[max(0,idx-500):min(len(body),idx+750)]
        if any(x in ctx for x in RESULT_TITLE_PATTERNS):
            continue
        if keyword in ("국민참여","시민참여"):
            action_hits=[x for x in PARTICIPATION_ACTION_SIGNALS if x in ctx]
            strong=[x for x in PARTICIPATION_CONTEXT if x in ctx]
            if not action_hits: continue
            if len(set(strong)) < 2 and not any(x in ctx for x in ["설문","의견","응답","모집","신청","제안","조사"]): continue
            noise_hits=sum(1 for x in GENERIC_BODY_NOISE if x in ctx)
            if noise_hits >= 3 and len(set(action_hits)) < 2: continue
            return True
        elif keyword in ("설문","설문조사","만족도 조사","의견수렴"):
            action_hits=[x for x in SURVEY_ACTION_CONTEXT + PARTICIPATION_CONTEXT if x in ctx]
            if not action_hits: continue
            if not any(x in ctx for x in ["참여","응답","기간","대상","방법","설문지","링크","의견"]): continue
            noise_hits=sum(1 for x in GENERIC_BODY_NOISE if x in ctx)
            if noise_hits >= 3 and len(set(action_hits)) < 2: continue
            return True
        else:
            signals=context_map.get(keyword,[])
            hits=[x for x in signals if x in ctx]
            if not hits: continue
            if keyword=="공모전":
                strong_phrases=[
                    "공모전 개최", "공모전 참가", "공모전 접수", "공모전 응모",
                    "공모전 신청", "공모전 모집", "공모전 참여", "공모전 출품",
                    "아이디어 공모", "작품 공모", "참가 신청", "작품 제출",
                ]
                if any(x in ctx for x in strong_phrases): return True
                action_hits=[x for x in ["접수","응모","출품","신청","모집","참가","참여","제출","기간","방법"] if x in ctx]
                if len(set(action_hits)) < 2: continue
            elif keyword=="퀴즈 이벤트":
                if not any(x in ctx for x in ["참여","응모","정답","문제","기간","방법"]): continue
            elif keyword=="평가단 모집":
                if not any(x in ctx for x in ["신청","모집","지원","기간","방법","선정"]): continue
            noise_hits=sum(1 for x in GENERIC_BODY_NOISE if x in ctx)
            if noise_hits >= 3 and len(set(hits)) < 2: continue
            return True
    return False

'''
s = s[:start] + new_func + s[end:]

old = '''    if not tc: return True
    if len(tc)<=10 and pc and (pc.startswith(tc) or pc.endswith(tc) or tc in pc):
        return True
    return False
'''
new = '''    if not tc: return True
    if len(tc)<=10 and pc and (pc.startswith(tc) or pc.endswith(tc) or tc in pc):
        return True
    if any(x.lower() in t.lower() for x in GENERIC_SHELL_TITLE_PATTERNS):
        return True
    if pc and (tc == pc or tc in pc) and len(tc) <= 100:
        return True
    return False
'''
if old not in s:
    raise SystemExit("is_probably_site_title 본문을 찾지 못했습니다.")
s=s.replace(old,new,1)

old='or path.endswith(("/index.do","/index.jsp","/main.do","/main.jsp")):'
new='or path.endswith(("/index.do","/index.jsp","/index.krc","/main.do","/main.jsp","/main.krc")):'
if old not in s:
    raise SystemExit("generic homepage path 조건을 찾지 못했습니다.")
s=s.replace(old,new,1)

p.write_text(s,encoding="utf-8")
print("V8.27 patch applied")
