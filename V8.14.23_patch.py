# -*- coding: utf-8 -*-
"""V8.14.23 safe patcher for the current monitor.py."""
from pathlib import Path
import re, shutil

FILE = Path("monitor.py")
VERSION_FROM = "V8.14.22"
VERSION_TO = "V8.14.23"

def fail(msg):
    raise SystemExit("PATCH_ABORT: " + msg)

if not FILE.exists():
    fail("monitor.py가 없습니다.")

src = FILE.read_text(encoding="utf-8")

if VERSION_FROM in src:
    src = src.replace(VERSION_FROM, VERSION_TO, 1)

marker = "GENERIC_URL_HINTS = ["
if "CONTEST_RESULT_TITLE_EXCLUDE" not in src:
    insert = r'''
CONTEST_RESULT_TITLE_EXCLUDE = [
    "수상작", "수상자", "수상작 발표", "수상자 발표",
    "당선작", "당선자", "입상작", "입상자",
    "수상 후보", "수상후보", "후보작",
    "공개 검증", "결과 발표", "결과 안내",
    "선정 결과", "최종 결과", "심사 결과",
    "결과보고", "결과 보고"
]

HOMEPAGE_TITLE_HINTS = [
    "메인", "메인페이지", "홈페이지", "home", "homepage",
    "site home", "main", "organization", "tourism organization"
]

'''
    if marker not in src:
        fail("GENERIC_URL_HINTS 위치를 찾지 못했습니다.")
    src = src.replace(marker, insert + marker, 1)

old = '''    if generic_url(r.url) or path in ("","/main","/home","/homepage"):
        return None,"GENERIC_PAGE",""
'''
new = '''    if generic_url(r.url) or path in (
        "", "/main", "/home", "/homepage",
        "/index.do", "/index.jsp", "/main.do", "/main.jsp"
    ) or path.endswith((
        "/index.do", "/index.jsp", "/main.do", "/main.jsp"
    )):
        return None,"GENERIC_PAGE",""
'''
if old in src:
    src = src.replace(old, new, 1)
elif 'path.endswith(("/index.do","/index.jsp","/main.do","/main.jsp"))' not in src:
    fail("홈페이지 URL 판정 블록을 찾지 못했습니다.")

old = '''    if any(x in title for x in EXCLUDE_TITLE): return None,"CONTEST_TITLE",""
'''
new = '''    # 공모전 개최/모집은 유지하고, 결과·수상·후보작 등 종료성 게시물만 제외한다.
    if any(x in title for x in CONTEST_RESULT_TITLE_EXCLUDE):
        return None,"CONTEST_RESULT_TITLE",""
'''
if old in src:
    src = src.replace(old, new, 1)
elif "CONTEST_RESULT_TITLE_EXCLUDE" not in src:
    fail("공모전 제목 판정 블록을 찾지 못했습니다.")

needle = '''    if not body:
        return None,"NO_BODY",""

    for kw in KEYWORDS:
        if meaningful_body_match(body,kw):
'''
replacement = '''    if not body:
        return None,"NO_BODY",""

    # 본문에 키워드가 있어도 기관 홈페이지/브랜드 제목이면 BODY 오탐으로 보지 않는다.
    title_norm = norm(title).lower()
    compact_title = re.sub(r"[^0-9a-z가-힣]+", "", title_norm)
    homepage_title_hit = any(h in title_norm for h in HOMEPAGE_TITLE_HINTS)
    if homepage_title_hit or (
        len(compact_title) <= 8 and
        any(x in title_norm for x in ("메인", "홈", "홈페이지", "home", "main"))
    ):
        return None,"GENERIC_TITLE",""

    for kw in KEYWORDS:
        if meaningful_body_match(body,kw):
'''
if needle in src:
    src = src.replace(needle, replacement, 1)
elif "homepage_title_hit" not in src:
    fail("BODY_MATCH 삽입 위치를 찾지 못했습니다.")

start = src.find("def meaningful_body_match(body,keyword):")
end = src.find("\ndef extract_links(", start)
if start < 0 or end < 0:
    fail("meaningful_body_match 함수를 찾지 못했습니다.")

new_func = '''def meaningful_body_match(body,keyword):
    body=norm(body)
    if len(body)<120: return False

    PARTICIPATION_STRONG_CONTEXT = [
        "설문조사","설문","의견수렴","의견조사","만족도 조사",
        "조사 참여","설문 참여","응답자","응답기간","참여기간",
        "참여방법","응답방법","의견 제출","의견을 제출",
        "의견을 남겨","설문에 참여","조사에 참여","국민 의견",
        "시민 의견","참여자 모집","모집합니다","모집 안내"
    ]
    SURVEY_STRONG_CONTEXT = [
        "설문조사","설문","설문지","설문 참여","설문에 참여",
        "조사 참여","조사기간","응답기간","참여기간","응답방법",
        "참여방법","응답자","조사대상","만족도 조사","의견조사"
    ]
    ACTION_SIGNALS = [
        "참여","신청","응답","제출","모집","조사","설문",
        "기간","대상","방법","링크","온라인"
    ]

    for m in re.finditer(re.escape(keyword),body,re.I):
        idx=m.start()
        ctx=body[max(0,idx-450):min(len(body),idx+650)]

        if keyword in ("국민참여","시민참여"):
            if not any(x in ctx for x in PARTICIPATION_STRONG_CONTEXT):
                continue
            if not any(x in ctx for x in ACTION_SIGNALS):
                continue

        elif keyword=="설문조사":
            if not any(x in ctx for x in SURVEY_STRONG_CONTEXT):
                continue
            if not any(x in ctx for x in ACTION_SIGNALS):
                continue

        if any(x in ctx for x in ACTION_SIGNALS):
            return True

    return False

'''
src = src[:start] + new_func + src[end+1:]

if VERSION_TO not in src:
    fail("V8.14.23 버전 문자열 적용 실패")

backup = FILE.with_suffix(".py.v8.14.23.bak")
shutil.copy2(FILE, backup)
FILE.write_text(src, encoding="utf-8")

print("=== V8.14.23 패치 완료 ===")
print("1) 공모전 결과/수상/후보작/공개검증/선정결과 차단")
print("2) /index.do /main.do 등 홈페이지 URL 차단")
print("3) 홈페이지·기관 브랜드 제목의 BODY 오탐 차단")
print("4) BODY_MATCH 참여/설문 문맥 강화")
print("5) 기존 상태/게시판 URL/중복처리 로직 유지")
