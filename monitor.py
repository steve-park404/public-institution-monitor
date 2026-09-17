import os
import json
import re
import time
import asyncio
import warnings
from datetime import datetime, timezone
from urllib.parse import urljoin

import aiohttp
import pandas as pd
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

TARGET_FILE = "monitor_targets.xlsx"
KEYWORD_FILE = "keywords.txt"
STATE_FILE = "state.json"
LOG_FILE = "monitor_log.json"

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "12"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "600"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

def load_keywords():
    if not os.path.exists(KEYWORD_FILE):
        return []
    out = []
    with open(KEYWORD_FILE, "r", encoding="utf-8-sig") as f:
        for line in f:
            s = line.strip()
            if s and not s.startswith("#"):
                out.append(s)
    return list(dict.fromkeys(out))

def normalize_state():
    if not os.path.exists(STATE_FILE):
        return {"version": 851, "initialized": False, "seen": {}}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except Exception:
        return {"version": 851, "initialized": False, "seen": {}}

    if isinstance(raw, dict):
        seen = raw.get("seen", {})
        if isinstance(seen, list):
            seen = {str(x): True for x in seen}
        elif not isinstance(seen, dict):
            seen = {}
        return {
            "version": raw.get("version", 851),
            "initialized": bool(raw.get("initialized", False)),
            "seen": seen,
        }
    if isinstance(raw, list):
        return {"version": 851, "initialized": False, "seen": {str(x): True for x in raw}}
    return {"version": 851, "initialized": False, "seen": {}}

def save_state(state):
    # Keep state bounded.
    seen = state.get("seen", {})
    if len(seen) > 10000:
        items = list(seen.items())[-10000:]
        state["seen"] = dict(items)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def clean_text(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def extract_posts(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    posts = []
    seen = set()

    # Prefer links that look like board/list item links.
    for a in soup.find_all("a", href=True):
        title = clean_text(a.get_text(" ", strip=True))
        href = urljoin(base_url, a.get("href"))
        if not title or len(title) < 2:
            continue
        if href in seen:
            continue
        if href.startswith(("javascript:", "mailto:", "#")):
            continue

        # Filter navigation links and common UI labels.
        low_title = title.lower()
        if low_title in {
            "home", "login", "로그인", "회원가입", "이전", "다음", "목록",
            "검색", "확인", "닫기", "more", "more +", "더보기"
        }:
            continue

        # Candidate titles often include notice/article wording.
        if len(title) > 180:
            continue

        posts.append({"title": title, "url": href})
        seen.add(href)

        if len(posts) >= RECENT_POSTS:
            break

    return posts[:RECENT_POSTS]

async def fetch(session, url, sem):
    async with sem:
        last_error = None
        for attempt in range(HTTP_RETRIES + 1):
            try:
                timeout = aiohttp.ClientTimeout(total=TIMEOUT_SECONDS)
                async with session.get(
                    url, headers=HEADERS, timeout=timeout, allow_redirects=True,
                    ssl=False
                ) as resp:
                    if resp.status >= 400:
                        raise RuntimeError(f"HTTP {resp.status}")
                    text = await resp.text(errors="ignore")
                    return text, None
            except Exception as e:
                last_error = str(e)
                if attempt < HTTP_RETRIES:
                    await asyncio.sleep(0.8 * (attempt + 1))
        return None, last_error or "HTTP 0"

async def inspect_target(session, row, keywords, state, sem, first_run):
    inst = clean_text(row["기관명"])
    board = clean_text(row["게시판명"])
    url = clean_text(row["URL"])

    html, err = await fetch(session, url, sem)
    if html is None:
        return {
            "institution": inst, "board": board, "url": url,
            "posts": 0, "matches": [], "error": err or "HTTP 0"
        }

    posts = extract_posts(html, url)
    matches = []

    for p in posts:
        # Title is available immediately; body is fetched only for candidates.
        hay = p["title"]
        hit = [k for k in keywords if k.lower() in hay.lower()]

        if not hit:
            # Fetch detail page to search body for keywords.
            body_html, _ = await fetch(session, p["url"], sem)
            if body_html:
                soup = BeautifulSoup(body_html, "html.parser")
                body = clean_text(soup.get_text(" ", strip=True))
                hit = [k for k in keywords if k.lower() in body.lower()]

        key = f"{inst}|{board}|{p['url']}"
        if hit and key not in state["seen"]:
            matches.append({
                "institution": inst,
                "board": board,
                "title": p["title"],
                "url": p["url"],
                "keywords": hit,
            })

        # First run seeds current matches to avoid alert storm.
        if hit:
            state["seen"][key] = True

    return {
        "institution": inst, "board": board, "url": url,
        "posts": len(posts), "matches": matches, "error": None
    }

async def send_telegram(text):
    if not TOKEN or not CHAT_ID:
        return {"ok": True, "skipped": True, "reason": "telegram secrets not configured"}

    api = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    payload = {"chat_id": CHAT_ID, "text": text, "disable_web_page_preview": True}
    try:
        timeout = aiohttp.ClientTimeout(total=15)
        async with aiohttp.ClientSession(timeout=timeout) as s:
            async with s.post(api, json=payload) as r:
                data = await r.json(content_type=None)
                return data
    except Exception as e:
        return {"ok": False, "error": str(e)}

def format_alert(m):
    return (
        "📢 공공기관 공지사항 키워드 알림\n\n"
        f"기관: {m['institution']}\n"
        f"게시판: {m['board']}\n"
        f"제목: {m['title']}\n"
        f"키워드: {', '.join(m['keywords'])}\n"
        f"링크: {m['url']}"
    )

async def main():
    started = time.time()
    df = pd.read_excel(TARGET_FILE).fillna("")
    keywords = load_keywords()
    state = normalize_state()
    first_run = not bool(state.get("initialized"))

    sem = asyncio.Semaphore(MAX_CONCURRENCY)
    connector = aiohttp.TCPConnector(limit=MAX_CONCURRENCY, ssl=False)

    results = []
    async with aiohttp.ClientSession(connector=connector) as session:
        tasks = [
            inspect_target(session, row, keywords, state, sem, first_run)
            for _, row in df.iterrows()
        ]
        for task in asyncio.as_completed(tasks):
            if time.time() - started > MAX_TOTAL_SECONDS:
                break
            try:
                results.append(await task)
            except Exception as e:
                results.append({
                    "institution": "", "board": "", "url": "",
                    "posts": 0, "matches": [], "error": str(e)
                })

    new_matches = []
    errors = []
    completed = len(results)
    posts_checked = 0

    for r in results:
        posts_checked += int(r.get("posts", 0))
        if r.get("error"):
            errors.append(r)
        new_matches.extend(r.get("matches", []))

    # Mark initialized only after the run finishes.
    state["initialized"] = True
    save_state(state)

    alert_results = []
    for m in new_matches:
        alert_results.append(await send_telegram(format_alert(m)))

    log = {
        "version": "V8.6",
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "targets": len(df),
        "completed": completed,
        "posts_checked": posts_checked,
        "new_matches": len(new_matches),
        "errors": len(errors),
        "timed_out": 1 if (time.time() - started) > MAX_TOTAL_SECONDS else 0,
        "first_run": first_run,
        "discovered_notice_boards": 0,
        "telegram": (
            {"ok": True, "skipped": True}
            if not new_matches
            else {"ok": all(x.get("ok", False) for x in alert_results), "sent": len(alert_results)}
        ),
        "error_details": errors[:20],
    }

    with open(LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

    print(json.dumps(log, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    asyncio.run(main())
