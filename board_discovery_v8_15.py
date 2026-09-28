#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V8.15 게시판 전용 발견기
- 기존 355기관 monitor.py와 분리 실행 가능
- 기관 홈페이지에서 실제 게시판 URL을 자동 탐색
- 전자정부 표준프레임워크 계열 selectBoardList.do?bbsId=... 우선 처리
- 발견한 게시판을 board_cache.json에 저장
- 기존 캐시는 보존하고 새로 확인된 기관만 갱신
- 우체국물류지원단(POLA) 같은 사례를 일반화
"""

import csv, json, os, re, time
from datetime import datetime, timezone, timedelta
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup

KST = timezone(timedelta(hours=9))
NOW = datetime.now(KST)
UA = "Mozilla/5.0 (compatible; TikkleBoardDiscovery/8.15; +https://github.com/steve-park404/public-institution-monitor)"

ROOT = os.getcwd()
CACHE_PATH = os.path.join(ROOT, "board_cache.json")
OUT_PATH = os.path.join(ROOT, "게시판_자동발견_결과.csv")
LOG_PATH = os.path.join(ROOT, "게시판_자동발견_진단.json")

# 기관목록 CSV 자동 탐색
CSV_CANDIDATES = [
    "기관목록.csv", "기관별_상태.csv", "institutions.csv",
    "targets.csv", "기관목록_355.csv"
]

# 게시판 의미가 강한 링크
BOARD_WORDS = [
    "공지사항", "공지", "알림", "알림마당", "새소식", "기관소식",
    "소식", "게시판", "국민참여", "시민참여", "설문", "참여"
]

# 제외 우선순위가 높은 게시판
EXCLUDE_WORDS = [
    "입찰", "채용", "인사", "개인정보", "사전정보공표", "공시",
    "계약", "법령", "규정", "이사회", "재무", "감사", "자료실"
]

# 전자정부 표준프레임워크 계열
EGOV_RE = re.compile(
    r"(?P<path>/[^\"'<>\\s]*?/cop/bbs/selectBoardList\.do(?:\?[^\"'<>\\s#]*)?)",
    re.I
)
EGOV_BBS_RE = re.compile(
    r"(?:[?&]bbsId=)(BBSMSTR_[A-Za-z0-9_]+)",
    re.I
)

session = requests.Session()
session.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.5"})
session.verify = True

def normalize_url(url):
    if not url:
        return ""
    url = url.strip()
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    p = urlparse(url)
    return urlunparse((p.scheme, p.netloc, p.path or "/", p.params, p.query, ""))

def canonical_board_url(url):
    """게시판 목록 URL의 불필요한 page/search 파라미터를 제거하되 bbsId는 보존."""
    p = urlparse(url)
    q = parse_qs(p.query)
    if "bbsId" not in q:
        return normalize_url(url)
    keep = [("bbsId", q["bbsId"][0])]
    # 검색 페이지에서 발견한 경우에도 1페이지 기본 URL로 저장
    return urlunparse((p.scheme, p.netloc, p.path, "", urlencode(keep), ""))

def fetch(url, timeout=12):
    try:
        r = session.get(url, timeout=timeout, allow_redirects=True)
        if r.status_code >= 400:
            return None, r.url, ""
        r.encoding = r.apparent_encoding or r.encoding
        return r, r.url, r.text
    except Exception as e:
        return None, url, str(e)

def find_targets_csv():
    for name in CSV_CANDIDATES:
        p = os.path.join(ROOT, name)
        if os.path.exists(p):
            return p
    # 기관명/URL/기관유형 헤더를 가진 CSV 자동 검색
    for name in os.listdir(ROOT):
        if not name.lower().endswith(".csv"):
            continue
        p = os.path.join(ROOT, name)
        try:
            with open(p, encoding="utf-8-sig", newline="") as f:
                row = next(csv.reader(f), [])
            if "기관명" in row and "URL" in row:
                return p
        except Exception:
            pass
    return None

def load_targets():
    p = find_targets_csv()
    if not p:
        raise FileNotFoundError("기관명/URL 컬럼을 가진 CSV를 찾지 못했습니다.")
    rows = []
    with open(p, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (row.get("기관명") or row.get("기관명 ") or "").strip()
            url = (row.get("URL") or row.get("url") or "").strip()
            typ = (row.get("기관유형") or "").strip()
            if name and url:
                rows.append({"기관명": name, "URL": url, "기관유형": typ})
    return p, rows

def load_cache():
    if not os.path.exists(CACHE_PATH):
        return {}
    try:
        with open(CACHE_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def cache_get(cache, name):
    return cache.get(name) or cache.get(name.strip())

def looks_like_board(url):
    p = urlparse(url)
    pathq = (p.path + "?" + p.query).lower()
    return (
        "selectboardlist.do" in pathq
        or "board" in p.path.lower()
        or "bbs" in p.path.lower()
        or "notice" in p.path.lower()
        or "news" in p.path.lower()
    )

def score_candidate(url, anchor_text="", page_title="", source="homepage"):
    text = f"{anchor_text} {page_title} {url}".lower()
    score = 0
    if "selectboardlist.do" in text:
        score += 45
    if "bbsid=" in text:
        score += 25
    for w in BOARD_WORDS:
        if w.lower() in anchor_text.lower():
            score += 18
    for w in BOARD_WORDS:
        if w.lower() in page_title.lower():
            score += 10
    for w in EXCLUDE_WORDS:
        if w.lower() in anchor_text.lower():
            score -= 35
    for w in EXCLUDE_WORDS:
        if w.lower() in page_title.lower():
            score -= 20
    if source == "search":
        score -= 5
    return score

def extract_candidates(base_url, html):
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    found = []

    # 1) href 링크
    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        if not href or href.lower().startswith(("javascript:", "#", "mailto:")):
            continue
        u = normalize_url(urljoin(base_url, href))
        if not u.startswith(("http://", "https://")):
            continue
        txt = a.get_text(" ", strip=True)
        if looks_like_board(u) or any(w in txt for w in BOARD_WORDS):
            found.append((score_candidate(u, txt, title), canonical_board_url(u), txt, title))

    # 2) HTML 안에 직접 들어있는 selectBoardList.do 문자열
    for m in EGOV_RE.finditer(html):
        raw = m.group("path")
        u = normalize_url(urljoin(base_url, raw))
        if "bbsId=" in u:
            found.append((score_candidate(u, "", title), canonical_board_url(u), "", title))

    # 3) bbsId만 별도 존재하는 경우
    for m in re.finditer(r"[^\"'<>\\s]{0,200}bbsId=BBSMSTR_[A-Za-z0-9_]+[^\"'<>\\s]{0,200}", html, re.I):
        raw = m.group(0).strip(" '\"")
        if "selectBoardList.do" in raw:
            u = normalize_url(urljoin(base_url, raw))
            found.append((score_candidate(u, "", title), canonical_board_url(u), "", title))

    return found

def verify_board(url, expected_name=""):
    r, final_url, html = fetch(url, timeout=12)
    if not r or not html:
        return False, {"reason": "FETCH_FAIL", "url": final_url}
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    text = soup.get_text(" ", strip=True)

    # 게시판 페이지에서 흔히 나타나는 목록/페이지/날짜 신호
    signals = 0
    if re.search(r"목록|번호|제목|날짜|작성자|조회", text):
        signals += 1
    if re.search(r"\d{4}[-./]\d{1,2}[-./]\d{1,2}", text):
        signals += 1
    if "selectBoardList.do" in final_url.lower():
        signals += 2
    if "bbsId=" in final_url:
        signals += 2
    if expected_name and expected_name.replace(" ", "")[:6] in text.replace(" ", ""):
        signals += 1

    return signals >= 3, {
        "reason": "VERIFIED" if signals >= 3 else "LOW_SIGNAL",
        "url": canonical_board_url(final_url),
        "title": title[:200],
        "signals": signals
    }

def homepage_candidates(name, home_url):
    r, final_url, html = fetch(home_url)
    if not r or not html:
        return [], {"status": "HOME_ERROR", "url": final_url}

    candidates = extract_candidates(final_url, html)

    # 메인페이지에서 바로 안 보이면 흔한 하위 경로 1단계만 추가 탐색
    soup = BeautifulSoup(html, "html.parser")
    child_urls = []
    for a in soup.find_all("a", href=True):
        txt = a.get_text(" ", strip=True)
        if any(w in txt for w in BOARD_WORDS):
            u = normalize_url(urljoin(final_url, a["href"]))
            if u and u not in child_urls and urlparse(u).netloc == urlparse(final_url).netloc:
                child_urls.append(u)

    for child in child_urls[:8]:
        r2, final2, html2 = fetch(child)
        if not r2 or not html2:
            continue
        candidates.extend(extract_candidates(final2, html2))

    # 점수/URL 중복 정리
    best = {}
    for sc, u, txt, title in candidates:
        if not u:
            continue
        if u not in best or sc > best[u][0]:
            best[u] = (sc, u, txt, title)

    ranked = sorted(best.values(), reverse=True)
    return ranked[:15], {"status": "HOME_OK", "url": final_url, "candidate_count": len(ranked)}

def discover(name, home_url, cache):
    # 이미 검증된 캐시가 있으면 유지
    old = cache_get(cache, name)
    if isinstance(old, dict):
        old_url = old.get("board_url") or old.get("url")
        if old_url:
            ok, info = verify_board(old_url, name)
            if ok:
                return old_url, "CACHE_VERIFIED", info

    ranked, diag = homepage_candidates(name, normalize_url(home_url))

    # 우체국물류지원단과 같은 eGov selectBoardList.do 후보를 최우선
    for sc, u, txt, title in ranked:
        if "selectboardlist.do" not in u.lower():
            continue
        ok, info = verify_board(u, name)
        if ok:
            return u, "DISCOVERED_EGOV", info

    # 일반 게시판 후보
    for sc, u, txt, title in ranked:
        if sc < 25:
            continue
        ok, info = verify_board(u, name)
        if ok:
            return u, "DISCOVERED_GENERAL", info

    return "", diag.get("status", "NOT_FOUND"), {
        "candidates": [
            {"score": x[0], "url": x[1], "text": x[2][:100], "title": x[3][:100]}
            for x in ranked[:8]
        ]
    }

def save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def main():
    csv_path, targets = load_targets()
    cache = load_cache()

    results = []
    changed = 0
    stats = {}

    # 기존 캐시 구조가 기관명 -> dict라는 전제를 유지.
    for i, row in enumerate(targets, 1):
        name = row["기관명"]
        home = row["URL"]

        board_url, status, info = discover(name, home, cache)

        if board_url:
            old = cache.get(name, {})
            cache[name] = {
                **(old if isinstance(old, dict) else {}),
                "board_url": board_url,
                "url": board_url,
                "institution": name,
                "homepage": normalize_url(home),
                "discovery_status": status,
                "discovered_at": NOW.isoformat(),
                "verified": True,
                "source": "V8.15_BOARD_DISCOVERY"
            }
            changed += 1
        else:
            cache.setdefault(name, {
                "institution": name,
                "homepage": normalize_url(home),
                "verified": False
            })

        stats[status] = stats.get(status, 0) + 1
        results.append({
            "기관명": name,
            "기관유형": row["기관유형"],
            "홈페이지": normalize_url(home),
            "게시판URL": board_url,
            "상태": status,
            "진단": json.dumps(info, ensure_ascii=False)[:1500]
        })

        if i % 20 == 0:
            print(f"[{i}/{len(targets)}] 발견 {changed}개")

    save_json(CACHE_PATH, cache)

    with open(OUT_PATH, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["기관명","기관유형","홈페이지","게시판URL","상태","진단"])
        w.writeheader()
        w.writerows(results)

    diag = {
        "version": "V8.15",
        "updated_at": NOW.isoformat(),
        "input_csv": csv_path,
        "targets": len(targets),
        "new_or_verified": changed,
        "status_counts": stats,
        "output_csv": OUT_PATH,
        "cache": CACHE_PATH
    }
    save_json(LOG_PATH, diag)

    print(json.dumps(diag, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
