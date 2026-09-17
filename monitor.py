import os
import re
import json
import time
import warnings
import asyncio
from urllib.parse import urljoin, urlparse, parse_qs, urlunparse

import aiohttp
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import openpyxl

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

KEYWORDS = ["설문조사", "시민참여", "국민참여", "공모전"]

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "12"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "600"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))

STATE_FILE = "state.json"
LOG_FILE = "monitor_log.json"
TARGET_FILE = "monitor_targets.xlsx"

GENERIC_TEXT = {
    "기타서비스", "etc service", "서브페이지 비쥬얼", "이전", "다음", "목록",
    "더보기", "검색", "로그인", "회원가입", "사이트맵", "sitemap", "home",
    "main", "메인", "공지사항", "공지", "게시판", "board"
}

DETAIL_QUERY_KEYS = {
    "articleno", "article_no", "nttid", "ntt_id", "seq", "idx", "wr_id",
    "wrid", "num", "no", "articleid", "article_id", "postid", "post_id",
    "bbsno", "bbs_no", "viewno", "view_no"
}
LIST_QUERY_KEYS = {
    "page", "pageno", "page_no", "pageindex", "page_index", "offset",
    "article.offset", "articlelimit", "article_limit", "limit", "size",
    "rows", "start", "startrow", "start_row"
}
HARD_BAD_PATH = re.compile(
    r"(?:/sitemap(?:/|\.|$)|/search(?:/|\.|$)|/search\.|/login(?:/|\.|$)|"
    r"/member/sitemap|/index(?:/|\.|$)|/main(?:/|\.|$))", re.I
)
UI_EDGE_RE = re.compile(r"^\s*(?:이전|다음|목록|더보기)\s+|\s+(?:이전|다음|목록|더보기)\s*$")

def clean_text(s):
    s = re.sub(r"\s+", " ", s or "").strip()
    s = UI_EDGE_RE.sub("", s).strip()
    s = re.sub(r"^(?:이전|다음|목록|더보기)\s*", "", s).strip()
    s = re.sub(r"\s*(?:이전|다음|목록|더보기)$", "", s).strip()
    return s

def meaningful_title(s):
    s = clean_text(s)
    if len(s) < 3 or len(s) > 300:
        return False
    if re.fullmatch(r"[\d\s\-_./]+", s):
        return False
    low = s.lower()
    if low in GENERIC_TEXT:
        return False
    if any(x in low for x in ("etc service", "서브페이지 비쥬얼", "사이트맵")):
        return False
    return True

def query_dict(url):
    return parse_qs(urlparse(url).query, keep_blank_values=True)

def is_probable_list_url(url):
    p = urlparse(url)
    path = p.path.lower()
    q = query_dict(url)
    if HARD_BAD_PATH.search(path):
        return True
    if q.get("mode", [""])[0].lower() in {"list", "search", "index"}:
        return True
    # Explicit pagination is a list signal unless a strong post identifier exists.
    has_detail = any(k.lower() in DETAIL_QUERY_KEYS for k in q.keys())
    if any(k.lower() in LIST_QUERY_KEYS for k in q.keys()) and not has_detail:
        return True
    if re.search(r"/(?:list|search|index)(?:/|\.|$)", path):
        return True
    return False

def detail_score(url, link_text=""):
    if is_probable_list_url(url):
        return -100
    p = urlparse(url)
    q = query_dict(url)
    path = p.path.lower()
    text = clean_text(link_text)
    score = 0

    if any(k.lower() in DETAIL_QUERY_KEYS for k in q.keys()):
        score += 7
    if q.get("mode", [""])[0].lower() in {"view", "detail", "read"}:
        score += 5
    if re.search(r"/(?:view|detail|read|article)(?:/|\.|$)", path):
        score += 5
    if text and meaningful_title(text):
        score += 3
    if len(text) >= 8:
        score += 1

    # Board/list-like destinations are not acceptable without detail evidence.
    if re.search(r"/board(?:/|\.|$)", path) and not (
        any(k.lower() in DETAIL_QUERY_KEYS for k in q.keys())
        or q.get("mode", [""])[0].lower() in {"view", "detail", "read"}
        or re.search(r"/(?:view|detail|read|article)(?:/|\.|$)", path)
    ):
        score -= 8

    # Generic content.do pages are not post pages unless a clear detail marker exists.
    if path.endswith("/content.do") or path.endswith("content.do"):
        if not (
            any(k.lower() in DETAIL_QUERY_KEYS for k in q.keys())
            or q.get("mode", [""])[0].lower() in {"view", "detail", "read"}
        ):
            score -= 10

    return score

def canonical_url(url):
    p = urlparse(url)
    # Keep query because many Korean board systems encode post IDs there.
    return urlunparse((p.scheme, p.netloc, p.path, p.params, p.query, ""))

def normalize_state(raw):
    if isinstance(raw, dict):
        seen = raw.get("seen", {})
        if isinstance(seen, list):
            seen = {str(x): True for x in seen}
        if not isinstance(seen, dict):
            seen = {}
        return {"version": 864, "initialized": bool(raw.get("initialized", False)), "seen": seen}
    if isinstance(raw, list):
        return {"version": 864, "initialized": False, "seen": {str(x): True for x in raw}}
    return {"version": 864, "initialized": False, "seen": {}}

def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return normalize_state(json.load(f))
    except Exception:
        return normalize_state({})

def save_state(state):
    seen = state.get("seen", {})
    if len(seen) > 10000:
        items = list(seen.items())[-10000:]
        state["seen"] = dict(items)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def load_targets():
    wb = openpyxl.load_workbook(TARGET_FILE, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(x).strip() if x is not None else "" for x in rows[0]]
    idx_name = header.index("기관명") if "기관명" in header else 0
    idx_url = header.index("URL") if "URL" in header else 1
    out = []
    for r in rows[1:]:
        if len(r) <= max(idx_name, idx_url):
            continue
        name, url = r[idx_name], r[idx_url]
        if name and url:
            out.append((str(name).strip(), str(url).strip()))
    return out[:100]

async def fetch(session, url):
    last_err = ""
    for attempt in range(HTTP_RETRIES + 1):
        try:
            async with session.get(url, timeout=TIMEOUT_SECONDS, allow_redirects=True) as resp:
                text = await resp.text(errors="ignore")
                if resp.status >= 400:
                    raise RuntimeError(f"HTTP {resp.status}")
                return text, resp.url.human_repr(), resp.status
        except Exception as e:
            last_err = str(e)
            if attempt < HTTP_RETRIES:
                await asyncio.sleep(0.7 * (attempt + 1))
    raise RuntimeError(last_err or "HTTP 0")

def clean_dom(soup):
    for tag in soup(["script", "style", "noscript", "svg", "header", "footer", "nav",
                     "aside", "form", "iframe", "canvas", "template"]):
        tag.decompose()
    for tag in soup.find_all(True):
        attrs = getattr(tag, "attrs", None) or {}
        ident = " ".join([
            str(attrs.get("id", "")),
            " ".join(map(str, attrs.get("class", []) if isinstance(attrs.get("class", []), list)
                         else [attrs.get("class", "")]))
        ]).lower()
        if any(x in ident for x in (
            "menu", "gnb", "lnb", "breadcrumb", "location", "related", "comment",
            "search", "pagination", "pager", "sns", "share", "banner", "visual",
            "quick", "shortcut", "footer", "header"
        )):
            try:
                tag.decompose()
            except Exception:
                pass

def extract_high_confidence_content(soup):
    clean_dom(soup)
    candidates = []
    selectors = [
        "article", "main", "[role='main']",
        ".article", "#article", ".article_view", "#article_view",
        ".view", "#view", ".view_content", "#view_content",
        ".board_view", "#board_view", ".bbs_view", "#bbs_view",
        ".content_view", "#content_view", ".contents", "#contents",
        ".board-content", ".board_content", ".post-content", ".post_content"
    ]
    for sel in selectors:
        try:
            for node in soup.select(sel):
                txt = clean_text(node.get_text(" ", strip=True))
                if len(txt) >= 80:
                    candidates.append(txt)
        except Exception:
            pass
    if not candidates:
        return ""
    # Prefer the most substantial high-confidence container, but cap to avoid page-wide boilerplate.
    candidates.sort(key=len, reverse=True)
    return candidates[0][:20000]

def extract_title(soup, link_text=""):
    # Link text from the board is often cleaner than page-wide UI titles.
    lt = clean_text(link_text)
    if meaningful_title(lt):
        return lt

    selectors = [
        "h1", "h2", ".subject", "#subject", ".board_subject", "#board_subject",
        ".view_title", "#view_title", ".article_title", "#article_title",
        "meta[property='og:title']", "title"
    ]
    for sel in selectors:
        try:
            node = soup.select_one(sel)
            if not node:
                continue
            if node.name == "meta":
                val = node.get("content", "")
            else:
                val = node.get_text(" ", strip=True)
            val = clean_text(val)
            if meaningful_title(val):
                return val
        except Exception:
            pass
    return ""

def find_post_links(soup, base_url):
    links = []
    seen = set()
    for a in soup.find_all("a"):
        attrs = getattr(a, "attrs", None) or {}
        href = attrs.get("href")
        if not href or not isinstance(href, str):
            continue
        if href.startswith(("#", "javascript:", "mailto:", "tel:")):
            continue
        url = canonical_url(urljoin(base_url, href))
        if url in seen:
            continue
        seen.add(url)
        text = clean_text(a.get_text(" ", strip=True))
        if not meaningful_title(text):
            continue
        sc = detail_score(url, text)
        if sc >= 6:
            links.append((sc, url, text))
    links.sort(key=lambda x: (-x[0], x[1]))
    return links[: max(RECENT_POSTS * 3, RECENT_POSTS)]

def match_post(title, content):
    title = clean_text(title)
    content = content or ""
    if not meaningful_title(title):
        return []
    # Always search the title. Search body only when a high-confidence content block exists.
    hay_title = title.lower()
    hay_body = content.lower()
    return [kw for kw in KEYWORDS if kw.lower() in hay_title or (hay_body and kw.lower() in hay_body)]

async def process_target(session, sem, target, state, seen_this_run):
    name, board_url = target
    result = {
        "institution": name, "board_url": board_url, "posts_checked": 0,
        "new_matches": [], "error": None, "candidate_links": 0, "detail_links": 0
    }
    async with sem:
        try:
            board_html, final_board_url, _ = await fetch(session, board_url)
            board_soup = BeautifulSoup(board_html, "html.parser")
            links = find_post_links(board_soup, final_board_url)
            result["candidate_links"] = len(links)

            # If the board itself somehow looks like a detail page, do not process it.
            links = [(sc, u, t) for sc, u, t in links if not is_probable_list_url(u)]
            result["detail_links"] = len(links)

            for _, post_url, link_text in links[:RECENT_POSTS]:
                if result["posts_checked"] >= RECENT_POSTS:
                    break
                key = f"{name}|{post_url}"
                if key in seen_this_run:
                    continue
                try:
                    post_html, final_post_url, _ = await fetch(session, post_url)
                    psoup = BeautifulSoup(post_html, "html.parser")
                    title = extract_title(psoup, link_text)
                    if not meaningful_title(title):
                        continue
                    # Never accept a fetched page that resolves to a list/sitemap/search page.
                    if is_probable_list_url(final_post_url):
                        continue
                    content = extract_high_confidence_content(psoup)
                    kws = match_post(title, content)
                    result["posts_checked"] += 1
                    if kws:
                        seen_this_run.add(key)
                        if key not in state["seen"]:
                            state["seen"][key] = True
                            result["new_matches"].append({
                                "institution": name,
                                "keyword": kws[0],
                                "keywords": kws,
                                "title": title,
                                "url": final_post_url,
                                "board_url": board_url
                            })
                        else:
                            seen_this_run.add(key)
                except Exception:
                    # A single malformed post must not fail the target.
                    continue
        except Exception as e:
            result["error"] = str(e) or "HTTP 0"
    return result

def telegram_message(item):
    kws = ", ".join(item.get("keywords", []))
    return (
        f"📢 공공기관 참여정보\n"
        f"기관: {item['institution']}\n"
        f"키워드: {kws}\n"
        f"제목: {item['title']}\n"
        f"{item['url']}"
    )

async def send_telegram(items):
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return {"ok": False, "sent": 0, "attempted": 0, "errors": ["Telegram secrets missing"], "truncated": bool(items)}
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    attempted = items[:TELEGRAM_MAX_SEND]
    errors = []
    sent = 0
    timeout = aiohttp.ClientTimeout(total=20)
    async with aiohttp.ClientSession(timeout=timeout) as s:
        for item in attempted:
            try:
                async with s.post(url, data={"chat_id": chat_id, "text": telegram_message(item)}) as r:
                    body = await r.text()
                    if r.status == 200:
                        sent += 1
                    else:
                        errors.append(f"HTTP {r.status}: {body[:200]}")
            except Exception as e:
                errors.append(str(e))
    return {
        "ok": sent == len(attempted) and not errors,
        "sent": sent,
        "attempted": len(attempted),
        "errors": errors,
        "truncated": len(items) > len(attempted)
    }

async def main():
    started = time.time()
    state = load_state()
    targets = load_targets()
    sem = asyncio.Semaphore(MAX_CONCURRENCY)
    seen_this_run = set()
    timeout = aiohttp.ClientTimeout(total=TIMEOUT_SECONDS)
    connector = aiohttp.TCPConnector(limit=MAX_CONCURRENCY, ssl=False)
    headers = {"User-Agent": "Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.6.4)"}

    results = []
    async with aiohttp.ClientSession(timeout=timeout, connector=connector, headers=headers) as session:
        tasks = [process_target(session, sem, t, state, seen_this_run) for t in targets]
        try:
            results = await asyncio.wait_for(asyncio.gather(*tasks), timeout=MAX_TOTAL_SECONDS)
        except asyncio.TimeoutError:
            print(f"[TIMEOUT] MAX_TOTAL_SECONDS={MAX_TOTAL_SECONDS}")
            # Gather completed tasks without blocking indefinitely.
            results = []
            for task in tasks:
                if task.done() and not task.cancelled():
                    try:
                        results.append(task.result())
                    except Exception:
                        pass

    all_matches = []
    errors = []
    posts_checked = 0
    completed = 0
    candidate_links = 0
    detail_links = 0
    for r in results:
        if r:
            completed += 1
            posts_checked += r.get("posts_checked", 0)
            candidate_links += r.get("candidate_links", 0)
            detail_links += r.get("detail_links", 0)
            all_matches.extend(r.get("new_matches", []))
            if r.get("error"):
                errors.append({
                    "institution": r["institution"],
                    "board_url": r["board_url"],
                    "error": r["error"]
                })

    telegram = await send_telegram(all_matches)

    state["initialized"] = True
    state["version"] = 864
    save_state(state)

    log = {
        "version": "8.6.4",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "targets": len(targets),
        "completed": completed,
        "posts_checked": posts_checked,
        "candidate_links": candidate_links,
        "detail_links": detail_links,
        "new_matches": len(all_matches),
        "errors": len(errors),
        "timed_out": completed < len(targets),
        "elapsed_seconds": round(time.time() - started, 1),
        "keywords": KEYWORDS,
        "matches": all_matches,
        "error_details": errors,
        "telegram": telegram
    }
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

    print(f"[V8.6.4] targets={len(targets)} completed={completed} posts_checked={posts_checked} "
          f"candidate_links={candidate_links} detail_links={detail_links} "
          f"new_matches={len(all_matches)} errors={len(errors)} timed_out={completed < len(targets)}")
    print(f"[KEYWORDS] {', '.join(KEYWORDS)}")
    for item in all_matches[:50]:
        print(f"[MATCH] {item['institution']} | {', '.join(item['keywords'])} | {item['title']} | {item['url']}")
    print(f"[TELEGRAM] {telegram}")

if __name__ == "__main__":
    asyncio.run(main())
