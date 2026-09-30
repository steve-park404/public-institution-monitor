from pathlib import Path
import re

p = Path("monitor.py")
if not p.exists():
    raise SystemExit("monitor.py가 없습니다.")

s = p.read_text(encoding="utf-8")

s = re.sub(r'(?m)^VERSION\s*=\s*"V[^"]+"',
           'VERSION = "V8.24"', s, count=1)

s = re.sub(
    r'(?m)^TARGET_FILE_CANDIDATES\s*=\s*\[[^\n]*\]',
    'TARGET_FILE_CANDIDATES = ["monitor_targets_통합.xlsx", "monitor_targets.xlsx", "url.xlsx", "targets.xlsx"]',
    s, count=1)

KEYWORDS = [
    "설문", "의견수렴", "시민참여", "국민참여",
    "공모전", "퀴즈 이벤트", "평가단 모집", "만족도 조사"
]

s = re.sub(
    r'(?m)^KEYWORDS\s*=\s*\[[^\n]*\]',
    f'KEYWORDS = {KEYWORDS!r}', s, count=1)

if "CONTEST_RESULT_TITLE_EXCLUDE" not in s:
    marker = 'EXCLUDE_TITLE = []'
    replacement = """EXCLUDE_TITLE = []

CONTEST_RESULT_TITLE_EXCLUDE = [
    "수상작", "수상자", "수상작 발표", "수상자 발표",
    "당선작", "당선자", "입상작", "입상자",
    "공개 검증", "결과 발표", "결과 안내", "선정 결과",
    "최종 결과", "심사 결과", "결과보고", "결과 보고"
]"""
    if marker in s:
        s = s.replace(marker, replacement, 1)

old_home = 'if generic_url(r.url) or path in ("","/main","/home","/homepage"):'
new_home = """if generic_url(r.url) or path in ("","/main","/home","/homepage","/index.do","/index.jsp","/main.do","/main.jsp") or path.endswith(("/index.do","/index.jsp","/main.do","/main.jsp")):"""
s = s.replace(old_home, new_home, 1)

old_result = 'if any(x in title for x in EXCLUDE_TITLE): return None,"CONTEST_TITLE",""'
new_result = 'if any(x in title for x in CONTEST_RESULT_TITLE_EXCLUDE): return None,"CONTEST_RESULT_TITLE",""'
s = s.replace(old_result, new_result, 1)

p.write_text(s, encoding="utf-8")

print("V8.24 patch applied")
for i, keyword in enumerate(KEYWORDS, 1):
    print(f"{i}. {keyword}")
