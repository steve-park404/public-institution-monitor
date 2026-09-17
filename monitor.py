#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
V8 public-institution notice monitor
- Prioritizes boards named "공지사항"
- Uses V7.6 validated boards as seed targets
- Optionally discovers "공지사항" links from homepage data in url_완성.xlsx
- Searches both title and body
- Sends only newly detected keyword matches to Telegram
- Saves state.json so the same post is not repeatedly notified
"""

import asyncio
import hashlib
import html
import json
import os
import re
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urljoin, urlparse, urldefrag

import aiohttp
import pandas as pd
from bs4 import BeautifulSoup

BASE = Path(__file__).resolve().parent
TARGET_FILE = BASE / "monitor_targets.xlsx"
HOME_FILE = BASE / "url_완성.xlsx"
KEYWORD_FILE = BASE / "keywords.txt"
STATE_FILE = BASE / "state.json"
LOG_FILE = BASE / "monitor_log.json"

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "20"))
TIMEOUT = int(os.getenv("TIMEOUT_SECONDS", "20"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "20"))
DISCOVER_NOTICE = os.getenv("DISCOVER_NOTICE", "true").lower() == "true"
SEED_ON_FIRST_RUN = os.getenv("SEED_ON_FIRST_RUN", "true").lower() == "true"
MAX_DISCOVERED_PER_HOME = int(os.getenv("MAX_DISCOVERED_PER_HOME", "5"))
MAX_TELEGRAM_MESSAGES = int(os.getenv("MAX_TELEGRAM_MESSAGES", "50"))

KST = timezone(timedelta(hours=9))

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36 "
    "PublicInstitutionNoticeMonitor/8.0"
)

NOTICE_NAMES = {
    "공지사항", "공지", "알림마당", "알림", "소식", "기관소식",
    "새소식", "공고", "공지·공고", "공지/공고", "알림·소식"
}

BAD_LINK_WORDS = {
    "로그인", "회원가입", "사이트맵", "검색", "메뉴", "닫기", "더보기",
    "개인정보", "이용약관", "오시는길", "문의", "예약", "진료", "결제",
    "설문", "캘린더", "행사일정", "교육신청"
}

TITLE_SELECTORS = [
    "h1", "h2", "h3",
    ".subject", ".title", ".tit", ".bbs_title", ".board_title",
    ".view-title", ".article-title", ".notice-title",
    "td.subject", "td.title", "a.subject", "a.title"
]

DATE_PATTERNS = [
    re.compile(r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2})"),
    re.compile(r"(20\d{2}\.\s*\d{1,2}\.\s*\d{1,2})"),
]

def normalize_space(s):
    return re.sub(r"\s+", " ", s or "").strip()

def normalize_url(url):
    if not url:
        return ""
    url, _ = urldefrag(str(url).strip())
    return url.rstrip("/")

def canonical_key(url):
    p = urlparse(url)
    # Keep query strings because many Korean board systems use IDs in query params.
    return f"{p.scheme.lower()}://{p.netloc.lower()}{p.path.rstrip('/')}?{p.query}"

def text_norm(s):
    return normalize_space(s).lower()

def load_keywords():
    if not KEYWORD_FILE.exists():
        return []
    out = []
    for line in KEYWORD_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        out.append(line)
    return out

def load_state():
    if not STATE_FILE.exists():
        return {"posts": {}, "initialized_at": None}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"posts": {}, "initialized_at": None}

def save_state(state):
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def post_id(institution, title, url):
    raw = f"{institution}|{normalize_space(title)}|{normalize_url(url)}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()

def likely_notice_link(text):
    t = normalize_space(text)
    if not t:
        return False
    compact = re.sub(r"[\s·/|_-]+", "", t)
    notice_compact = {re.sub(r"[\s·/|_-]+", "", x) for x in NOTICE_NAMES}
    if compact in notice_compact:
        return True
    return ("공지" in t and len(t) <= 20) or ("알림" in t and len(t) <= 20)

def is_bad_link(text, href):
    t = text_norm(text)
    if any(x in t for x in BAD_LINK_WORDS):
        return True
    if href.startswith("javascript:") or href.startswith("#"):
        return True
    return False

def extract_body_text(soup):
    for tag in soup(["script", "style", "noscript", "svg", "header", "footer", "nav"]):
        tag.decompose()
    return normalize_space(soup.get_text(" ", strip=True))

def extract_title_from_detail(soup):
    # Avoid using <title> or generic h1 alone because Korean sites often
    # use a common page title such as "공지사항 상세" or a site name.
    for sel in TITLE_SELECTORS:
        node = soup.select_one(sel)
        if node:
            txt = normalize_space(node.get_text(" ", strip=True))
            if 3 <= len(txt) <= 300:
                return txt
    return ""

def extract_links_from_board(soup, board_url):
    out = []
    seen = set()
    for a in soup.find_all("a", href=True):
        text = normalize_space(a.get_text(" ", strip=True))
        href = normalize_url(urljoin(board_url, a.get("href")))
        if not text or not href or is_bad_link(text, href):
            continue
        if href == normalize_url(board_url):
            continue
        # Avoid obvious non-detail links.
        if any(x in href.lower() for x in [
            "javascript:", "mailto:", "/login", "/member", "/search"
        ]):
            continue
        key = (text, href)
        if key not in seen:
            seen.add(key)
            out.append((text, href))
    return out

def score_post_candidate(text, href):
    score = 0
    if 4 <= len(text) <= 200:
        score += 2
    if any(x in text for x in ["공지", "안내", "공고", "모집", "채용", "입찰", "교육"]):
        score += 1
    if re.search(r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2})", text):
        score += 2
    if any(x in href.lower() for x in ["view", "detail", "article", "board", "bbs", "seq=", "idx=", "ntt", "no="]):
        score += 2
    return score

def extract_post_candidates(soup, board_url):
    links = extract_links_from_board(soup, board_url)
    scored = [(score_post_candidate(t, h), t, h) for t, h in links]
    scored.sort(key=lambda x: (-x[0], x[1]))
    result = []
    seen_urls = set()
    for score, text, href in scored:
        if score < 2:
            continue
        if href in seen_urls:
            continue
        seen_urls.add(href)
        result.append((text, href))
        if len(result) >= RECENT_POSTS:
            break
    return result

async def fetch(session, url):
    try:
        async with session.get(
            url,
            timeout=aiohttp.ClientTimeout(total=TIMEOUT),
            headers={"User-Agent": USER_AGENT, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7"},
            allow_redirects=True,
        ) as r:
            content_type = r.headers.get("Content-Type", "")
            if r.status >= 400:
                return r.status, "", str(r.url)
            raw = await r.read()
            enc = r.charset or "utf-8"
            try:
                text = raw.decode(enc, errors="ignore")
            except Exception:
                text = raw.decode("utf-8", errors="ignore")
            if "html" not in content_type.lower() and "<html" not in text[:1000].lower():
                return r.status, "", str(r.url)
            return r.status, text, str(r.url)
    except Exception as e:
        return 0, "", str(e)

async def verify_post(session, institution, title_from_list, url):
    status, body, final_url = await fetch(session, url)
    if status != 200 or not body:
        return None
    soup = BeautifulSoup(body, "html.parser")
    text = extract_body_text(soup)
    list_title = normalize_space(title_from_list)
    # The list title should appear in the detail page when possible.
    title = list_title
    if not title:
        title = extract_title_from_detail(soup)
    # Ignore obvious function/detail pages.
    if not title or len(title) < 3:
        return None
    return {
        "title": title[:300],
        "url": normalize_url(final_url if final_url.startswith("http") else url),
        "body": text[:30000],
    }

async def scan_board(session, target):
    institution = target["기관명"]
    board_name = target.get("게시판명", "")
    board_url = normalize_url(target["게시판URL"])

    status, body, final_url = await fetch(session, board_url)
    result = {
        "기관명": institution,
        "게시판명": board_name,
        "게시판URL": board_url,
        "status": status,
        "posts_checked": 0,
        "matches": [],
        "error": "",
    }
    if status != 200 or not body:
        result["error"] = f"HTTP {status}"
        return result

    soup = BeautifulSoup(body, "html.parser")
    candidates = extract_post_candidates(soup, board_url)

    sem = asyncio.Semaphore(MAX_CONCURRENCY)

    async def one(item):
        async with sem:
            return await verify_post(session, institution, item[0], item[1])

    verified = [x for x in await asyncio.gather(*(one(x) for x in candidates)) if x]
    result["posts_checked"] = len(verified)

    keywords = load_keywords()
    for p in verified:
        hay = f'{p["title"]}\n{p["body"]}'.lower()
        hit = [k for k in keywords if k.lower() in hay]
        if hit:
            result["matches"].append({
                "institution": institution,
                "board": board_name or "공지사항",
                "title": p["title"],
                "url": p["url"],
                "keywords": hit,
            })
    return result

async def discover_notice_boards(session):
    if not DISCOVER_NOTICE or not HOME_FILE.exists():
        return []

    try:
        df = pd.read_excel(HOME_FILE)
    except Exception:
        return []

    if not {"기관명", "URL"}.issubset(df.columns):
        return []

    async def discover_one(row):
        institution = str(row["기관명"]).strip()
        home = normalize_url(str(row["URL"]).strip())
        if not home or home.lower() == "nan":
            return []
        status, body, final_url = await fetch(session, home)
        if status != 200 or not body:
            return []
        soup = BeautifulSoup(body, "html.parser")
        found = []
        seen = set()
        for a in soup.find_all("a", href=True):
            text = normalize_space(a.get_text(" ", strip=True))
            href = normalize_url(urljoin(home, a.get("href")))
            if not likely_notice_link(text) or is_bad_link(text, href):
                continue
            if href in seen:
                continue
            seen.add(href)
            found.append({
                "기관명": institution,
                "게시판명": text,
                "게시판URL": href,
                "우선순위": 1,
                "출처": "V8 홈페이지 공지사항 자동발견",
            })
            if len(found) >= MAX_DISCOVERED_PER_HOME:
                break
        return found

    results = await asyncio.gather(*(discover_one(r) for _, r in df.iterrows()))
    out = []
    for x in results:
        out.extend(x)
    return out

async def main():
    keywords = load_keywords()
    if not keywords:
        print("ERROR: keywords.txt is empty.")
        sys.exit(2)

    if not TARGET_FILE.exists():
        print(f"ERROR: {TARGET_FILE.name} not found.")
        sys.exit(2)

    target_df = pd.read_excel(TARGET_FILE)
    required = {"기관명", "게시판URL"}
    if not required.issubset(target_df.columns):
        print(f"ERROR: target file needs columns {required}")
        sys.exit(2)

    targets = []
    for _, r in target_df.iterrows():
        url = normalize_url(str(r["게시판URL"]).strip())
        if not url or url.lower() == "nan":
            continue
        targets.append({
            "기관명": str(r["기관명"]).strip(),
            "게시판명": str(r.get("게시판명", "")).strip(),
            "게시판URL": url,
            "우선순위": int(r.get("우선순위", 9)),
            "출처": str(r.get("출처", "")),
        })

    timeout = aiohttp.ClientTimeout(total=TIMEOUT)
    connector = aiohttp.TCPConnector(limit=MAX_CONCURRENCY, ssl=False)
    async with aiohttp.ClientSession(timeout=timeout, connector=connector) as session:
        discovered = await discover_notice_boards(session)

        # Merge discovered notice boards with validated seeds.
        all_targets = targets + discovered
        dedup = {}
        for t in all_targets:
            key = (t["기관명"], canonical_key(t["게시판URL"]))
            if key not in dedup or t.get("우선순위", 9) < dedup[key].get("우선순위", 9):
                dedup[key] = t
        targets = list(dedup.values())

        print(f"Monitoring {len(targets)} board URLs")

        results = await asyncio.gather(
            *(scan_board(session, t) for t in targets)
        )

    state = load_state()
    posts = state.setdefault("posts", {})
    first_run = not bool(state.get("initialized_at"))

    all_matches = []
    errors = []
    checked = 0

    for res in results:
        checked += res["posts_checked"]
        if res["error"]:
            errors.append({
                "기관명": res["기관명"],
                "게시판": res["게시판명"],
                "URL": res["게시판URL"],
                "error": res["error"],
            })
        for m in res["matches"]:
            pid = post_id(m["institution"], m["title"], m["url"])
            if pid in posts:
                continue
            posts[pid] = {
                "institution": m["institution"],
                "board": m["board"],
                "title": m["title"],
                "url": m["url"],
                "keywords": m["keywords"],
                "first_seen": datetime.now(KST).isoformat(),
            }
            if not (first_run and SEED_ON_FIRST_RUN):
                all_matches.append(m)

    # Prevent state from growing forever.
    if len(posts) > 10000:
        items = sorted(posts.items(), key=lambda kv: kv[1].get("first_seen", ""))
        posts = dict(items[-8000:])
        state["posts"] = posts

    if not state.get("initialized_at"):
        state["initialized_at"] = datetime.now(KST).isoformat()

    save_state(state)

    summary = {
        "run_at": datetime.now(KST).isoformat(),
        "targets": len(targets),
        "posts_checked": checked,
        "new_matches": len(all_matches),
        "errors": len(errors),
        "first_run": first_run,
        "discovered_notice_boards": len(discovered),
    }
    LOG_FILE.write_text(
        json.dumps({"summary": summary, "errors": errors}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    await send_telegram(all_matches, errors, summary)

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if errors:
        print("Errors:")
        for e in errors[:30]:
            print(e)

async def send_telegram(matches, errors, summary):
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        print("Telegram secrets not configured; skipping notification.")
        return

    messages = []
    for m in matches[:MAX_TELEGRAM_MESSAGES]:
        kws = ", ".join(m["keywords"])
        messages.append(
            "🚨 <b>공공기관 공지사항 키워드 감지</b>\n"
            f"<b>기관:</b> {html.escape(m['institution'])}\n"
            f"<b>게시판:</b> {html.escape(m['board'])}\n"
            f"<b>제목:</b> {html.escape(m['title'])}\n"
            f"<b>키워드:</b> {html.escape(kws)}\n"
            f"<a href=\"{html.escape(m['url'], quote=True)}\">게시물 바로가기</a>"
        )

    if errors and not messages:
        messages.append(
            "⚠️ <b>공공기관 모니터링 오류</b>\n"
            f"전체 게시판: {summary['targets']}\n"
            f"접속/처리 오류: {summary['errors']}\n"
            "GitHub Actions 로그를 확인하세요."
        )

    if not messages:
        return

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    timeout = aiohttp.ClientTimeout(total=20)

    async with aiohttp.ClientSession(timeout=timeout) as session:
        for msg in messages:
            try:
                async with session.post(
                    url,
                    json={
                        "chat_id": chat_id,
                        "text": msg[:4096],
                        "parse_mode": "HTML",
                        "disable_web_page_preview": False,
                    },
                ) as r:
                    if r.status >= 400:
                        print("Telegram error:", r.status, await r.text())
            except Exception as e:
                print("Telegram exception:", e)

if __name__ == "__main__":
    asyncio.run(main())
