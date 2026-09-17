import asyncio
import json
import os
import re
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

import aiohttp
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import warnings

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

KEYWORDS = ["설문조사", "시민참여", "국민참여", "공모전"]

TARGET_FILE = "monitor_targets.xlsx"
STATE_FILE = "state.json"
LOG_FILE = "monitor_log.json"

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "12"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "600"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))

LIST_PARAM_NAMES = {
    "page", "pageno", "pageNo", "pageIndex", "offset", "article.offset",
    "articleLimit", "limit", "size", "rows"
}

def load_targets():
    import openpyxl
    wb = openpyxl.load_workbook(TARGET_FILE, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(x).strip() if x is not None else "" for x in rows[0]]
    idx = {h: i for i, h in enumerate(header)}
    out = []
    for row in rows[1:]:
        if not row:
            continue
        institution = row[idx.get("기관명", 0)]
        board = row[idx.get("게시판", idx.get("URL", 1))]
        if institution and board:
            out.append({"institution": str(institution).strip(), "board": str(board).strip()})
    return out

def normalize_state(raw):
    if isinstance(raw, dict):
        seen = raw.get("seen", {})
        if isinstance(seen, list):
            seen = {str(x): True for x in seen}
        elif not isinstance(seen, dict):
            seen = {}
        return {"version": 863, "initialized": bool(raw.get("initialized", False)), "seen": seen}
    if isinstance(raw, list):
        return {"version": 863, "initialized": True, "seen": {str(x): True for x in raw}}
    return {"version": 863, "initialized": False, "seen": {}}

def load_state():
    p = Path(STATE_FILE)
    if not p.exists():
        return normalize_state({})
    try:
        return normalize_state(json.loads(p.read_text(encoding="utf-8")))
    except Exception:
        return normalize_state({})

def save_state(state):
    seen = state["seen"]
    if len(seen) > 10000:
        keys = list(seen.keys())[-10000:]
        state["seen"] = {k: True for k in keys}
    Path(STATE_FILE).write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def is_probable_list_url(url):
    q = parse_qs(urlparse(url).query)
    path = urlparse(url).path.lower()
    if any(k.lower() in {x.lower() for x in q.keys()} for k in LIST_PARAM_NAMES):
        # offset/page parameters can also appear on detail pages, so don't reject
        # when a clear detail marker exists.
        if any(marker in q for marker in ["mode", "articleNo", "seq", "nttId", "idx", "no", "num", "view", "wr_id"]):
            return False
        return True
    if re.search(r"/(list|search|index|board/list|notice/list)(/|\.|$)", path):
        return True
    return False

def clean_text(s):
    return re.sub(r"\s+", " ", s or "").strip()

def tag_classes(tag):
    try:
        attrs = getattr(tag, "attrs", None)
        if not isinstance(attrs, dict):
            return ""
        cls = attrs.get("class", [])
        if isinstance(cls, str):
            return cls.lower()
        if isinstance(cls, (list, tuple)):
            return " ".join(str(x) for x in cls).lower()
        return ""
    except Exception:
        return ""

def tag_id(tag):
    try:
        attrs = getattr(tag, "attrs", None)
        if not isinstance(attrs, dict):
            return ""
        return str(attrs.get("id", "")).lower()
    except Exception:
        return ""

EXCLUDE_RE = re.compile(
    r"(header|footer|gnb|lnb|nav|menu|breadcrumb|sitemap|"
    r"related|relation|recommend|popular|search|comment|reply|"
    r"share|sns|banner|quick|skip|pagination|pager|"
    r"prev|next|previous|attach|file|download|login|"
    r"popup|layer|aside|sidebar)",
    re.I
)

CONTENT_HINT_RE = re.compile(
    r"(view|read|content|contents|article|board|bbs|notice|"
    r"detail|detailview|post|body|cont|txt|editor|articleview)",
    re.I
)

def strip_common_noise(soup):
    for tag in soup(["script", "style", "noscript", "template", "svg", "canvas",
                     "form", "iframe", "header", "footer", "nav", "aside"]):
        try:
            tag.decompose()
        except Exception:
            pass

    for tag in list(soup.find_all(True)):
        if tag_classes(tag) and EXCLUDE_RE.search(tag_classes(tag)):
            try:
                tag.decompose()
            except Exception:
                pass
        elif tag_id(tag) and EXCLUDE_RE.search(tag_id(tag)):
            try:
                tag.decompose()
            except Exception:
                pass
    return soup

def extract_title(soup):
    candidates = []
    for selector in [
        "h1", "h2", "h3",
        ".tit", ".title", ".subject", ".bbs-title", ".board-title",
        ".view-title", ".article-title", "[class*='title']",
        "[class*='subject']"
    ]:
        try:
            for x in soup.select(selector):
                t = clean_text(x.get_text(" ", strip=True))
                if 2 <= len(t) <= 300:
                    candidates.append(t)
        except Exception:
            pass
    if candidates:
        # Prefer longer, meaningful candidates; avoid generic page headings.
        candidates.sort(key=lambda x: (len(x) >= 5, len(x)), reverse=True)
        return candidates[0]
    if soup.title:
        return clean_text(soup.title.get_text(" ", strip=True))[:300]
    return ""

def likely_content(soup):
    # Work on a copy so title extraction can still inspect the original cleaned tree.
    soup = strip_common_noise(soup)

    candidates = []
    for tag in soup.find_all(["article", "main", "div", "section", "td"]):
        try:
            attrs = getattr(tag, "attrs", None)
            if not isinstance(attrs, dict):
                attrs = {}
            cls = tag_classes(tag)
            tid = tag_id(tag)
            marker = f"{cls} {tid}"
            if marker and CONTENT_HINT_RE.search(marker):
                txt = clean_text(tag.get_text(" ", strip=True))
                if 80 <= len(txt) <= 200000:
                    score = 0
                    if re.search(r"(content|contents|article|view|detail|body|cont)", marker, re.I):
                        score += 5
                    if tag.name == "article":
                        score += 4
                    if tag.name == "main":
                        score += 3
                    score += min(len(txt) / 50000, 3)
                    candidates.append((score, txt))
        except Exception:
            continue

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        best = candidates[0][1]
    else:
        # Conservative fallback: visible body text after noise removal.
        best = clean_text(soup.get_text(" ", strip=True))

    # Remove repeated UI fragments that often survive DOM filtering.
    lines = [clean_text(x) for x in re.split(r"[\r\n]+", best)]
    lines = [x for x in lines if x]
    return clean_text(" ".join(lines))[:250000]

def extract_links(board_url, soup):
    links = []
    seen = set()
    for a in soup.find_all("a"):
        try:
            attrs = getattr(a, "attrs", None)
            if not isinstance(attrs, dict):
                continue
            href = attrs.get("href")
            if not href or not isinstance(href, str):
                continue
            href = href.strip()
            if href.lower().startswith(("javascript:", "mailto:", "#")):
                continue
            url = urljoin(board_url, href)
            if urlparse(url).scheme not in ("http", "https"):
                continue
            text = clean_text(a.get_text(" ", strip=True))
            if not text:
                continue
            key = url.split("#")[0]
            if key in seen:
                continue
            seen.add(key)
            links.append((text, key))
        except Exception:
            continue
    return links

def detail_score(board_url, url, text):
    score = 0
    q = parse_qs(urlparse(url).query)
    path = urlparse(url).path.lower()
    lowtext = text.lower()

    if is_probable_list_url(url):
        score -= 8
    if any(k.lower() in {x.lower() for x in q.keys()} for k in
           ["articleNo", "nttId", "seq", "idx", "wr_id", "num", "no"]):
        score += 7
    if any(k.lower() in {x.lower() for x in q.keys()} for k in ["mode", "view"]):
        score += 4
    if re.search(r"/view|/detail|/read|/article", path):
        score += 4
    if len(text) >= 4:
        score += 1
    if re.fullmatch(r"\d{1,5}", text):
        score -= 6
    if text in {"더보기", "목록", "이전글", "다음글", "검색"}:
        score -= 8
    if url.rstrip("/") == board_url.rstrip("/"):
        score -= 10
    return score

def find_post_links(board_url, soup):
    raw = extract_links(board_url, soup)
    scored = [(detail_score(board_url, u, t), t, u) for t, u in raw]
    scored = [x for x in scored if x[0] >= 1]
    scored.sort(key=lambda x: x[0], reverse=True)
    # Keep a reasonable number; duplicates are removed by URL.
    out, seen = [], set()
    for score, text, url in scored:
        if url in seen:
            continue
        seen.add(url)
        out.append({"title_hint": text, "url": url, "score": score})
        if len(out) >= RECENT_POSTS * 3:
            break
    return out

async def fetch(session, url):
    last_error = None
    for attempt in range(HTTP_RETRIES + 1):
        try:
            timeout = aiohttp.ClientTimeout(total=TIMEOUT_SECONDS)
            async with session.get(
                url, timeout=timeout, allow_redirects=True,
                headers={"User-Agent": "Mozilla/5.0 public-institution-monitor/8.6.3"}
            ) as r:
                body = await r.text(errors="ignore")
                if r.status >= 400:
                    raise RuntimeError(f"HTTP {r.status}")
                return body, str(r.url)
        except Exception as e:
            last_error = e
            if attempt < HTTP_RETRIES:
                await asyncio.sleep(0.5 * (attempt + 1))
    raise RuntimeError(str(last_error) if last_error else "HTTP 0")

def stable_key(institution, url):
    return f"{institution}|{url.split('#')[0]}"

def match_post(institution, board_url, title, content, url, state):
    title = clean_text(title)
    content = clean_text(content)
    # Reject obviously invalid/list-like pages.
    if len(title) < 2 or re.fullmatch(r"\d{1,5}", title):
        return None
    if is_probable_list_url(url):
        return None

    hits = [kw for kw in KEYWORDS if kw in title or kw in content]
    if not hits:
        return None

    key = stable_key(institution, url)
    if key in state["seen"]:
        return None

    return {
        "institution": institution,
        "board": board_url,
        "keyword": hits,
        "title": title[:300],
        "url": url
    }

async def process_target(target, session, state):
    institution, board_url = target["institution"], target["board"]
    try:
        html, final_board = await fetch(session, board_url)
        soup = BeautifulSoup(html, "html.parser")
        posts = find_post_links(final_board, soup)

        # If the board itself is a detail page, treat it as a single candidate.
        if not posts and not is_probable_list_url(final_board):
            posts = [{"title_hint": "", "url": final_board, "score": 1}]

        posts = posts[:RECENT_POSTS]
        matches = []
        checked = 0

        for p in posts:
            try:
                ph, final_url = await fetch(session, p["url"])
                if is_probable_list_url(final_url):
                    continue
                psoup = BeautifulSoup(ph, "html.parser")
                title = extract_title(psoup) or p["title_hint"]
                content = likely_content(psoup)
                checked += 1
                m = match_post(institution, board_url, title, content, final_url, state)
                if m:
                    matches.append(m)
            except Exception:
                continue

        return {"institution": institution, "board": board_url,
                "checked": checked, "matches": matches, "error": None}
    except Exception as e:
        return {"institution": institution, "board": board_url,
                "checked": 0, "matches": [], "error": str(e)}

async def send_telegram(matches):
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        return {"ok": False, "sent": 0, "attempted": 0, "errors": ["missing secret"]}

    selected = matches[:TELEGRAM_MAX_SEND]
    errors = []
    sent = 0
    async with aiohttp.ClientSession() as s:
        for m in selected:
            text = (
                f"📢 공공기관 참여정보\n"
                f"기관: {m['institution']}\n"
                f"키워드: {', '.join(m['keyword'])}\n"
                f"제목: {m['title']}\n"
                f"링크: {m['url']}"
            )
            try:
                url = f"https://api.telegram.org/bot{token}/sendMessage"
                async with s.post(url, data={"chat_id": chat_id, "text": text},
                                  timeout=aiohttp.ClientTimeout(total=15)) as r:
                    data = await r.json(content_type=None)
                    if data.get("ok"):
                        sent += 1
                    else:
                        errors.append(str(data.get("description", f"HTTP {r.status}")))
            except Exception as e:
                errors.append(str(e))
    return {"ok": sent == len(selected) and not errors,
            "sent": sent, "attempted": len(selected),
            "errors": errors, "truncated": len(matches) > len(selected)}

async def main():
    start = time.monotonic()
    targets = load_targets()
    state = load_state()
    all_matches = []
    results = []
    errors = 0
    completed = 0
    checked = 0

    connector = aiohttp.TCPConnector(limit=MAX_CONCURRENCY, ssl=False)
    async with aiohttp.ClientSession(connector=connector) as session:
        sem = asyncio.Semaphore(MAX_CONCURRENCY)

        async def one(t):
            async with sem:
                return await process_target(t, session, state)

        tasks = [asyncio.create_task(one(t)) for t in targets]
        try:
            for fut in asyncio.as_completed(tasks, timeout=MAX_TOTAL_SECONDS):
                res = await fut
                results.append(res)
                completed += 1
                checked += res["checked"]
                if res["error"]:
                    errors += 1
                    print(f"[ERROR] {res['institution']} | {res['board']} | {res['error']}")
                for m in res["matches"]:
                    all_matches.append(m)
                    print(f"[NEW MATCH {len(all_matches)}]")
                    print(f"기관: {m['institution']}")
                    print(f"키워드: {', '.join(m['keyword'])}")
                    print(f"제목: {m['title']}")
                    print(f"URL: {m['url']}")
        except asyncio.TimeoutError:
            print("[TIMEOUT] MAX_TOTAL_SECONDS reached")

    # Mark only newly detected matches as seen.
    for m in all_matches:
        state["seen"][stable_key(m["institution"], m["url"])] = True
    state["initialized"] = True
    save_state(state)

    telegram = await send_telegram(all_matches)

    log = {
        "version": "8.6.3",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "targets": len(targets),
        "completed": completed,
        "posts_checked": checked,
        "new_matches": len(all_matches),
        "errors": errors,
        "timed_out": completed < len(targets),
        "runtime_seconds": round(time.monotonic() - start, 2),
        "keywords": KEYWORDS,
        "matches": all_matches,
        "telegram": telegram
    }
    Path(LOG_FILE).write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(log, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    asyncio.run(main())
