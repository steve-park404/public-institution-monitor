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

# 검색 키워드는 반드시 이 4개만 사용
KEYWORDS = ["설문조사", "시민참여", "국민참여"]

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "12"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "600"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))

STATE_FILE = "state.json"
LOG_FILE = "monitor_log.json"
TARGET_FILE = "monitor_targets.xlsx"

GENERIC_TITLES = {
    "기타서비스", "etc service", "서브페이지 비쥬얼", "이전", "다음", "목록",
    "더보기", "검색", "로그인", "회원가입", "사이트맵", "sitemap",
    "home", "main", "메인", "공지사항", "공지", "게시판", "board"
}

DETAIL_QUERY_KEYS = {
    "articleno", "article_no", "nttid", "ntt_id", "seq", "idx", "wr_id",
    "wrid", "num", "no", "articleid", "article_id", "postid", "post_id",
    "bbsno", "bbs_no", "viewno", "view_no", "cntid", "cnt_id"
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

NOISE_ID_CLASS = (
    "menu", "gnb", "lnb", "breadcrumb", "location", "related", "comment",
    "search", "pagination", "pager", "sns", "share", "banner", "visual",
    "quick", "shortcut", "footer", "header", "navigation", "nav"
)

CONTENT_SELECTORS = [
    "article", "main", "[role='main']",
    ".article_view", "#article_view", ".article-view", "#article-view",
    ".article", "#article",
    ".view_content", "#view_content", ".view-content", "#view-content",
    ".board_view", "#board_view", ".board-view", "#board-view",
    ".bbs_view", "#bbs_view", ".bbs-view", "#bbs-view",
    ".content_view", "#content_view", ".content-view", "#content-view",
    ".post_content", "#post_content", ".post-content", "#post-content",
    ".board_content", "#board_content", ".board-content", "#board-content",
    ".contents_view", "#contents_view"
]

UI_WORD_RE = re.compile(
    r"^(?:이전|다음|목록|더보기|검색|공유|인쇄|첨부파일|다운로드)$", re.I
)

def clean_text(s):
    s = re.sub(r"\s+", " ", s or "").strip()
    # 상세 제목에 섞이는 전후 UI 텍스트 제거
    for _ in range(2):
        s = re.sub(r"^\s*(?:이전|다음|목록|더보기)\s+", "", s).strip()
        s = re.sub(r"\s+(?:이전|다음|목록|더보기)\s*$", "", s).strip()
    return s

def meaningful_title(s):
    s = clean_text(s)
    if len(s) < 3 or len(s) > 300:
        return False
    if re.fullmatch(r"[\d\s\-_./]+", s):
        return False
    if UI_WORD_RE.fullmatch(s):
        return False
    low = s.lower()
    if low in GENERIC_TITLES:
        return False
    if any(x in low for x in ("etc service", "서브페이지 비쥬얼", "사이트맵", "sitemap")):
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

    mode = q.get("mode", [""])[0].lower()
    if mode in {"list", "search", "index"}:
        return True

    keys = {k.lower() for k in q.keys()}
    has_detail = bool(keys & DETAIL_QUERY_KEYS)
    if keys & LIST_QUERY_KEYS and not has_detail:
        return True

    if re.search(r"/(?:list|search|index)(?:/|\.|$)", path):
        return True

    return False

def has_strong_detail_evidence(url):
    if is_probable_list_url(url):
        return False
    p = urlparse(url)
    q = query_dict(url)
    keys = {k.lower() for k in q.keys()}
    path = p.path.lower()

    if keys & DETAIL_QUERY_KEYS:
        return True
    if q.get("mode", [""])[0].lower() in {"view", "detail", "read"}:
        return True
    if re.search(r"/(?:view|detail|read|article)(?:/|\.|$)", path):
        return True
    return False

def detail_score(url, link_text=""):
    if not has_strong_detail_evidence(url):
        return -100

    score = 0
    p = urlparse(url)
    q = query_dict(url)
    path = p.path.lower()
    text = clean_text(link_text)

    keys = {k.lower() for k in q.keys()}
    if keys & DETAIL_QUERY_KEYS:
        score += 8
    if q.get("mode", [""])[0].lower() in {"view", "detail", "read"}:
        score += 5
    if re.search(r"/(?:view|detail|read|article)(?:/|\.|$)", path):
        score += 5

    if meaningful_title(text):
        score += 4
    if len(text) >= 8:
        score += 1

    return score

def canonical_url(url):
    p = urlparse(url)
    return urlunparse((p.scheme, p.netloc, p.path, p.params, p.query, ""))

def normalize_state(raw):
    if isinstance(raw, dict):
        seen = raw.get("seen", {})
        if isinstance(seen, list):
            seen = {str(x): True for x in seen}
        if not isinstance(seen, dict):
            seen = {}
        return {
            "version": 866,
            "initialized": bool(raw.get("initialized", False)),
            "seen": seen
        }
    if isinstance(raw, list):
        return {"version": 866, "initialized": False,
                "seen": {str(x): True for x in raw}}
    return {"version": 866, "initialized": False, "seen": {}}

def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return normalize_state(json.load(f))
    except Exception:
        return normalize_state({})

def save_state(state):
    seen = state.get("seen", {})
    if len(seen) > 10000:
        state["seen"] = dict(list(seen.items())[-10000:])
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
        if r[idx_name] and r[idx_url]:
            out.append((str(r[idx_name]).strip(), str(r[idx_url]).strip()))
    return out[:100]

async def fetch(session, url):
    last_err = "HTTP 0"
    for attempt in range(HTTP_RETRIES + 1):
        try:
            async with session.get(url, timeout=TIMEOUT_SECONDS,
                                   allow_redirects=True) as resp:
                text = await resp.text(errors="ignore")
                if resp.status >= 400:
                    raise RuntimeError(f"HTTP {resp.status}")
                return text, resp.url.human_repr(), resp.status
        except Exception as e:
            last_err = str(e) or "HTTP 0"
            if attempt < HTTP_RETRIES:
                await asyncio.sleep(0.7 * (attempt + 1))
    raise RuntimeError(last_err)

def safe_remove(tag):
    try:
        tag.decompose()
    except Exception:
        pass

def tag_identity(tag):
    attrs = getattr(tag, "attrs", None) or {}
    cls = attrs.get("class", [])
    if not isinstance(cls, list):
        cls = [cls]
    return (
        str(attrs.get("id", "")) + " " +
        " ".join(str(x) for x in cls)
    ).lower()

def strip_noise(soup):
    for tag in soup([
        "script", "style", "noscript", "svg", "iframe",
        "canvas", "template", "header", "footer", "nav",
        "aside", "form"
    ]):
        safe_remove(tag)

    # 명확한 공통 UI/메뉴/검색/관련게시물 영역 제거
    for tag in soup.find_all(True):
        ident = tag_identity(tag)
        if any(x in ident for x in NOISE_ID_CLASS):
            safe_remove(tag)

def extract_high_confidence_content(soup):
    strip_noise(soup)
    candidates = []

    for sel in CONTENT_SELECTORS:
        try:
            for node in soup.select(sel):
                txt = clean_text(node.get_text(" ", strip=True))
                if len(txt) < 80:
                    continue

                # 본문 후보에 제목/게시판 공통 영역이 지나치게 많이 포함되면 제외
                low = txt.lower()
                noise_hits = sum(
                    low.count(x) for x in
                    ("이전글", "다음글", "목록", "검색", "관련게시물", "댓글")
                )
                if noise_hits >= 4:
                    continue

                candidates.append(txt)
        except Exception:
            continue

    if not candidates:
        return "", "none"

    # 가장 긴 후보를 사용하되 페이지 전체가 아닌 고신뢰 컨테이너만 허용
    candidates.sort(key=len, reverse=True)
    return candidates[0][:20000], "high"

def extract_title(soup, link_text=""):
    # 게시판 링크의 실제 제목을 최우선으로 사용
    lt = clean_text(link_text)
    if meaningful_title(lt):
        return lt

    selectors = [
        "h1", "h2", "h3",
        ".subject", "#subject", ".board_subject", "#board_subject",
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
            continue
    return ""

def find_post_links(soup, base_url):
    links = []
    seen = set()

    for a in soup.find_all("a"):
        attrs = getattr(a, "attrs", None) or {}
        href = attrs.get("href")
        if not isinstance(href, str) or not href:
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

        score = detail_score(url, text)
        if score >= 6:
            links.append((score, url, text))

    # 같은 상세 URL은 링크가 여러 개여도 한 번만
    links.sort(key=lambda x: (-x[0], x[1]))
    return links[:max(RECENT_POSTS * 3, RECENT_POSTS)]

def match_post(title, content, content_confidence):
    title = clean_text(title)
    if not meaningful_title(title):
        return []

    # 제목 매칭은 가장 신뢰도가 높으므로 그대로 유지
    title_low = title.lower()
    title_hits = [kw for kw in KEYWORDS if kw.lower() in title_low]
    if title_hits:
        return title_hits

    # 본문은 고신뢰 컨테이너가 확인된 경우에만 검색
    if content_confidence != "high" or not content:
        return []

    body_low = content.lower()

    # 본문 전체에 한 번만 등장하는 일반 단어/하단 공통문구의 오탐을 줄임.
    # 실제 게시물 본문에서 발견되더라도, 제목에 없는 경우에는
    # "키워드가 문맥상 핵심인지"를 간단히 확인한다.
    hits = []
    for kw in KEYWORDS:
        positions = [m.start() for m in re.finditer(re.escape(kw.lower()), body_low)]
        if not positions:
            continue

        # 키워드 주변 ±80자의 문맥에서 공모/참여/설문 관련 표현이 있는 경우 우선 인정.
        # 단, 키워드 자체가 이미 충분히 구체적이므로 1회라도 인정한다.
        # 대신 페이지 하단/공통문구 가능성이 높은 매우 짧은 후보는 제외한다.
        if len(body_low) >= 80:
            hits.append(kw)

    return hits

async def process_target(session, sem, target, state, seen_this_run):
    name, board_url = target
    result = {
        "institution": name,
        "board_url": board_url,
        "posts_checked": 0,
        "new_matches": [],
        "error": None,
        "candidate_links": 0,
        "detail_links": 0
    }

    async with sem:
        try:
            board_html, final_board_url, _ = await fetch(session, board_url)
            board_soup = BeautifulSoup(board_html, "html.parser")
            links = find_post_links(board_soup, final_board_url)
            result["candidate_links"] = len(links)

            # 최종적으로도 상세 페이지 증거를 재확인
            links = [
                item for item in links
                if has_strong_detail_evidence(item[1])
            ]
            result["detail_links"] = len(links)

            for _, post_url, link_text in links[:RECENT_POSTS]:
                key = f"{name}|{post_url}"
                if key in seen_this_run:
                    continue
                seen_this_run.add(key)

                try:
                    post_html, final_post_url, _ = await fetch(session, post_url)
                    if not has_strong_detail_evidence(final_post_url):
                        continue

                    psoup = BeautifulSoup(post_html, "html.parser")
                    title = extract_title(psoup, link_text)
                    if not meaningful_title(title):
                        continue

                    content, confidence = extract_high_confidence_content(psoup)
                    kws = match_post(title, content, confidence)
                    result["posts_checked"] += 1

                    if not kws:
                        continue

                    if key not in state["seen"]:
                        state["seen"][key] = True
                        result["new_matches"].append({
                            "institution": name,
                            "keyword": kws[0],
                            "keywords": kws,
                            "match_type": "title" if kws[0].lower() in title.lower() else "body",
                            "title": title,
                            "url": final_post_url,
                            "board_url": board_url
                        })
                except Exception:
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
        return {
            "ok": False, "sent": 0, "attempted": 0,
            "errors": ["Telegram secrets missing"],
            "truncated": bool(items)
        }

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    attempted = items[:TELEGRAM_MAX_SEND]
    errors = []
    sent = 0

    timeout = aiohttp.ClientTimeout(total=20)
    async with aiohttp.ClientSession(timeout=timeout) as s:
        for item in attempted:
            try:
                async with s.post(
                    url,
                    data={"chat_id": chat_id, "text": telegram_message(item)}
                ) as r:
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
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.6.6)"
    }

    results = []
    timed_out = False

    async with aiohttp.ClientSession(
        timeout=timeout, connector=connector, headers=headers
    ) as session:
        tasks = [
            asyncio.create_task(
                process_target(session, sem, target, state, seen_this_run)
            )
            for target in targets
        ]
        try:
            results = await asyncio.wait_for(
                asyncio.gather(*tasks), timeout=MAX_TOTAL_SECONDS
            )
        except asyncio.TimeoutError:
            timed_out = True
            print(f"[TIMEOUT] MAX_TOTAL_SECONDS={MAX_TOTAL_SECONDS}")
            for task in tasks:
                if task.done() and not task.cancelled():
                    try:
                        results.append(task.result())
                    except Exception:
                        pass
            for task in tasks:
                if not task.done():
                    task.cancel()

    all_matches = []
    errors = []
    posts_checked = 0
    completed = 0
    candidate_links = 0
    detail_links = 0

    for r in results:
        if not r:
            continue
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
    state["version"] = 865
    save_state(state)

    log = {
        "version": "8.6.6",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "targets": len(targets),
        "completed": completed,
        "posts_checked": posts_checked,
        "candidate_links": candidate_links,
        "detail_links": detail_links,
        "new_matches": len(all_matches),
        "errors": len(errors),
        "timed_out": timed_out,
        "elapsed_seconds": round(time.time() - started, 1),
        "keywords": KEYWORDS,
        "matches": all_matches,
        "error_details": errors,
        "telegram": telegram
    }

    with open(LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

    print(
        f"[V8.6.5] targets={len(targets)} completed={completed} "
        f"posts_checked={posts_checked} candidate_links={candidate_links} "
        f"detail_links={detail_links} new_matches={len(all_matches)} "
        f"errors={len(errors)} timed_out={timed_out}"
    )
    print(f"[KEYWORDS] {', '.join(KEYWORDS)}")

    for item in all_matches[:50]:
        print(
            f"[MATCH] {item['institution']} | "
            f"{', '.join(item['keywords'])} | "
            f"{item['match_type']} | {item['title']} | {item['url']}"
        )

    print(f"[TELEGRAM] {telegram}")

if __name__ == "__main__":
    asyncio.run(main())
