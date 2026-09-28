#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V8.15.1 게시판 전용 자동 발견기
- V8.14.17의 실제 기관 목록을 자동 탐색
- 기관목록 CSV가 없더라도 기관별_상태.csv의 실제 컬럼을 분석
- 기관 수 300 미만이면 안전하게 실패
- eGov selectBoardList.do + bbsId 우선 발견
- 발견 결과를 board_cache.json에 저장
"""

import csv, json, os, re, sys, time
from datetime import datetime, timezone, timedelta
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup

KST = timezone(timedelta(hours=9))
NOW = datetime.now(KST)
UA = "Mozilla/5.0 (compatible; TikkleBoardDiscovery/8.15.1)"

ROOT = os.getcwd()
CACHE_PATH = os.path.join(ROOT, "board_cache.json")
OUT_PATH = os.path.join(ROOT, "게시판_자동발견_결과.csv")
LOG_PATH = os.path.join(ROOT, "게시판_자동발견_진단.json")

MIN_TARGETS = 300
EXPECTED_TARGETS = 355

session = requests.Session()
session.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.5"})

BOARD_WORDS = ["공지사항","공지","알림","알림마당","새소식","기관소식","소식","게시판","국민참여","시민참여","설문","참여"]
EXCLUDE_WORDS = ["입찰","채용","인사","개인정보","사전정보공표","공시","계약","법령","규정","이사회","재무","감사","자료실"]

def norm(s):
    return (s or "").strip()

def normalize_url(url):
    url = norm(url)
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    p = urlparse(url)
    return urlunparse((p.scheme, p.netloc, p.path or "/", p.params, p.query, ""))

def canonical_board_url(url):
    p = urlparse(url)
    q = parse_qs(p.query)
    if "bbsId" in q:
        return urlunparse((p.scheme, p.netloc, p.path, "", urlencode({"bbsId": q["bbsId"][0]}), ""))
    return normalize_url(url)

def load_csv_rows(path):
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        fields = reader.fieldnames or []
        for row in reader:
            rows.append((row, fields))
    return rows

def pick_col(fields, candidates):
    normalized = {re.sub(r"\s+", "", x or ""): x for x in fields}
    for c in candidates:
        key = re.sub(r"\s+", "", c)
        if key in normalized:
            return normalized[key]
    return None

def score_csv(path, rows, fields):
    name_col = pick_col(fields, ["기관명","기관","기관명칭","기관이름"])
    url_col = pick_col(fields, ["URL","url","홈페이지","홈페이지URL","기관홈페이지","사이트","웹사이트"])
    score = 0
    if name_col: score += 50
    if url_col: score += 50
    if len(rows) >= 300: score += 100
    if "기관별_상태" in os.path.basename(path): score += 20
    return score, name_col, url_col

def find_institution_csv():
    candidates = []
    for name in os.listdir(ROOT):
        if not name.lower().endswith(".csv"):
            continue
        path = os.path.join(ROOT, name)
        try:
            rows, fields = load_csv_rows(path)
            sc, nc, uc = score_csv(path, rows, fields)
            if nc and uc:
                candidates.append((sc, path, len(rows), nc, uc))
        except Exception:
            continue

    candidates.sort(reverse=True)
    if not candidates:
        raise RuntimeError("기관명+홈페이지 컬럼을 가진 CSV를 찾지 못했습니다.")

    print("[CSV 후보]")
    for item in candidates[:10]:
        print(" ", item)

    # 가장 많은 기관을 가진 정상 CSV를 우선.
    chosen = max(candidates, key=lambda x: (x[2] >= MIN_TARGETS, x[2], x[0]))
    sc, path, count, nc, uc = chosen

    if count < MIN_TARGETS:
        raise RuntimeError(
            f"기관 목록 검증 실패: {os.path.basename(path)}에서 {count}개만 읽혔습니다. "
            f"최소 {MIN_TARGETS}개가 필요합니다."
        )

    return path, nc, uc, count

def load_targets():
    path, name_col, url_col, count = find_institution_csv()
    rows, fields = load_csv_rows(path)
    targets = []
    seen = set()

    for row in rows:
        name = norm(row.get(name_col))
        url = normalize_url(row.get(url_col))
        if not name or not url or name in seen:
            continue
        seen.add(name)
        targets.append({
            "기관명": name,
            "URL": url,
            "기관유형": norm(row.get("기관유형") or row.get("기관 유형") or "")
        })

    if len(targets) < MIN_TARGETS:
        raise RuntimeError(
            f"실제 유효 기관 수가 {len(targets)}개입니다. "
            f"{MIN_TARGETS}개 미만이면 안전을 위해 실행하지 않습니다."
        )

    return path, targets

def load_cache():
    if not os.path.exists(CACHE_PATH):
        return {}
    try:
        with open(CACHE_PATH, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}

def fetch(url, timeout=10):
    try:
        r = session.get(url, timeout=timeout, allow_redirects=True)
        if r.status_code >= 400:
            return None, r.url, ""
        r.encoding = r.apparent_encoding or r.encoding
        return r, r.url, r.text
    except Exception as e:
        return None, url, ""

def score_candidate(url, anchor, title):
    t = f"{url} {anchor} {title}".lower()
    score = 0
    if "selectboardlist.do" in t: score += 60
    if "bbsid=bbsMSTR_".lower() in t: score += 30
    for w in BOARD_WORDS:
        if w.lower() in anchor.lower(): score += 15
    for w in BOARD_WORDS:
        if w.lower() in title.lower(): score += 5
    for w in EXCLUDE_WORDS:
        if w.lower() in anchor.lower(): score -= 35
    return score

def extract_candidates(base_url, html):
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    out = []

    for a in soup.find_all("a", href=True):
        href = norm(a.get("href"))
        if not href or href.lower().startswith(("javascript:", "#", "mailto:")):
            continue
        u = normalize_url(urljoin(base_url, href))
        txt = a.get_text(" ", strip=True)
        if (
            "selectboardlist.do" in u.lower()
            or "bbsid=" in u.lower()
            or any(w in txt for w in BOARD_WORDS)
        ):
            out.append((score_candidate(u, txt, title), canonical_board_url(u), txt))

    # HTML 내부에 문자열로 존재하는 eGov URL도 탐색
    for m in re.finditer(
        r"""(?:https?:)?//[^"'<>\\s]+/cop/bbs/selectBoardList\.do\?[^"'<>\\s]+""",
        html, re.I
    ):
        u = normalize_url(m.group(0))
        out.append((score_candidate(u, "", title), canonical_board_url(u), ""))

    # 상대경로
    for m in re.finditer(
        r"""[^"'<>\\s]{0,150}/cop/bbs/selectBoardList\.do\?[^"'<>\\s]+""",
        html, re.I
    ):
        u = normalize_url(urljoin(base_url, m.group(0)))
        out.append((score_candidate(u, "", title), canonical_board_url(u), ""))

    best = {}
    for sc, u, txt in out:
        if not u: continue
        if u not in best or sc > best[u][0]:
            best[u] = (sc, u, txt)
    return sorted(best.values(), reverse=True)

def verify_board(url, expected_name):
    r, final_url, html = fetch(url)
    if not r or not html:
        return False, {"reason": "FETCH_FAIL", "url": final_url}

    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    text = soup.get_text(" ", strip=True)

    signals = 0
    if "selectboardlist.do" in final_url.lower(): signals += 2
    if "bbsid=" in final_url.lower(): signals += 2
    if re.search(r"번호\s+제목|제목\s+작성자|작성자\s+날짜|조회", text): signals += 2
    if re.search(r"\d{4}[-./]\d{1,2}[-./]\d{1,2}", text): signals += 1
    if expected_name and expected_name.replace(" ","")[:6] in text.replace(" ",""): signals += 1

    return signals >= 3, {
        "reason": "VERIFIED" if signals >= 3 else "LOW_SIGNAL",
        "url": canonical_board_url(final_url),
        "title": title[:200],
        "signals": signals
    }

def discover(name, homepage, cache):
    old = cache.get(name)
    if isinstance(old, dict):
        old_url = old.get("board_url") or old.get("url")
        if old_url:
            ok, info = verify_board(old_url, name)
            if ok:
                return old_url, "CACHE_VERIFIED", info

    r, final_url, html = fetch(homepage)
    if not r or not html:
        return "", "HOME_ERROR", {"url": final_url}

    candidates = extract_candidates(final_url, html)

    # 공지/알림 관련 하위 링크 최대 10개 추가 탐색
    soup = BeautifulSoup(html, "html.parser")
    children = []
    for a in soup.find_all("a", href=True):
        txt = a.get_text(" ", strip=True)
        if not any(w in txt for w in BOARD_WORDS):
            continue
        u = normalize_url(urljoin(final_url, a["href"]))
        if urlparse(u).netloc == urlparse(final_url).netloc and u not in children:
            children.append(u)

    for child in children[:10]:
        r2, final2, html2 = fetch(child)
        if r2 and html2:
            candidates.extend(extract_candidates(final2, html2))

    best = {}
    for sc, u, txt in candidates:
        if u not in best or sc > best[u][0]:
            best[u] = (sc, u, txt)

    ranked = sorted(best.values(), reverse=True)

    # eGov 우선
    for sc, u, txt in ranked:
        if "selectboardlist.do" in u.lower() and "bbsid=" in u.lower():
            ok, info = verify_board(u, name)
            if ok:
                return u, "DISCOVERED_EGOV", info

    # 일반 게시판
    for sc, u, txt in ranked:
        if sc < 25:
            continue
        ok, info = verify_board(u, name)
        if ok:
            return u, "DISCOVERED_GENERAL", info

    return "", "CANDIDATE_NOT_VERIFIED", {
        "candidate_count": len(ranked),
        "candidates": [{"score":x[0],"url":x[1],"text":x[2][:100]} for x in ranked[:10]]
    }

def save_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def main():
    input_csv, targets = load_targets()
    print(f"[기관목록 검증] {os.path.basename(input_csv)} / {len(targets)}개")

    cache = load_cache()
    results = []
    counts = {}
    found = 0

    for i, row in enumerate(targets, 1):
        name, home = row["기관명"], row["URL"]
        board_url, status, info = discover(name, home, cache)

        if board_url:
            old = cache.get(name, {})
            cache[name] = {
                **(old if isinstance(old, dict) else {}),
                "institution": name,
                "homepage": home,
                "board_url": board_url,
                "url": board_url,
                "verified": True,
                "discovery_status": status,
                "discovered_at": NOW.isoformat(),
                "source": "V8.15.1"
            }
            found += 1
        else:
            if name not in cache:
                cache[name] = {
                    "institution": name,
                    "homepage": home,
                    "verified": False
                }

        counts[status] = counts.get(status, 0) + 1
        results.append({
            "기관명": name,
            "기관유형": row["기관유형"],
            "홈페이지": home,
            "게시판URL": board_url,
            "상태": status,
            "진단": json.dumps(info, ensure_ascii=False)[:1800]
        })

        if i % 25 == 0:
            print(f"[{i}/{len(targets)}] 게시판 확보 {found}")

    save_json(CACHE_PATH, cache)

    with open(OUT_PATH, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["기관명","기관유형","홈페이지","게시판URL","상태","진단"])
        w.writeheader()
        w.writerows(results)

    diag = {
        "version": "V8.15.1",
        "updated_at": NOW.isoformat(),
        "input_csv": input_csv,
        "targets": len(targets),
        "found_or_verified": found,
        "not_found": len(targets)-found,
        "status_counts": counts,
        "safety_min_targets": MIN_TARGETS,
        "expected_targets": EXPECTED_TARGETS
    }
    save_json(LOG_PATH, diag)
    print(json.dumps(diag, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"FATAL: {e}", file=sys.stderr)
        sys.exit(1)
