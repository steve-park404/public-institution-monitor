# -*- coding: utf-8 -*-
"""V8.26 patch: use the manually-maintained integrated Excel as the authoritative target list."""
from pathlib import Path
import re

p = Path("monitor.py")
if not p.exists():
    raise SystemExit("monitor.py가 없습니다.")

s = p.read_text(encoding="utf-8")

# Version
s = re.sub(r'(?m)^VERSION\s*=\s*"V[^"]+"',
           'VERSION = "V8.26"', s, count=1)

# Authoritative target file: the manually updated integrated workbook first.
old = r'TARGET_FILE_CANDIDATES\s*=\s*\[[^\n]*\]'
new = 'TARGET_FILE_CANDIDATES = ["monitor_targets_통합.xlsx", "monitor_targets.xlsx", "url.xlsx", "targets.xlsx"]'
s, n = re.subn(r'(?m)^TARGET_FILE_CANDIDATES\s*=\s*\[[^\n]*\]', new, s, count=1)
if n != 1:
    raise SystemExit("TARGET_FILE_CANDIDATES 패치를 찾지 못했습니다.")

# 8 operational keywords
keywords = 'KEYWORDS = ["설문", "의견수렴", "시민참여", "국민참여", "공모전", "퀴즈 이벤트", "평가단 모집", "만족도 조사"]'
s, n = re.subn(r'(?m)^KEYWORDS\s*=\s*\[[^\n]*\]', keywords, s, count=1)
if n != 1:
    raise SystemExit("KEYWORDS 패치를 찾지 못했습니다.")

# More robust retries
s, n = re.subn(r'(?m)^HTTP_RETRIES\s*=\s*\d+', 'HTTP_RETRIES = 2', s, count=1)
if n != 1:
    raise SystemExit("HTTP_RETRIES 패치를 찾지 못했습니다.")

# Broaden result-announcement exclusion.
result_patterns = [
    "당첨자", "결과발표", "결과 발표", "조사결과", "조사 결과",
    "설문결과", "설문 결과", "선정결과", "선정 결과", "참여자 발표",
    "수상자", "수상작", "수상자 발표", "수상작 발표",
    "당선자", "당선작", "입상자", "입상작", "결과 안내", "결과안내",
    "최종 결과", "최종결과", "심사 결과", "심사결과", "결과보고", "결과 보고"
]
s, n = re.subn(r'(?ms)^RESULT_TITLE_PATTERNS\s*=\s*\[.*?\n\]',
               'RESULT_TITLE_PATTERNS = ' + repr(result_patterns), s, count=1)
if n != 1:
    raise SystemExit("RESULT_TITLE_PATTERNS 패치를 찾지 못했습니다.")

p.write_text(s, encoding="utf-8")
print("V8.26 patch applied")
print("authoritative target: monitor_targets_통합.xlsx")
print("8개 키워드:", ", ".join(["설문","의견수렴","시민참여","국민참여","공모전","퀴즈 이벤트","평가단 모집","만족도 조사"]))
