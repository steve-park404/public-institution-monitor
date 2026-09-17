import asyncio
import json
import os
import re
import time
import warnings
from datetime import datetime, timezone

import aiohttp
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
from openpyxl import load_workbook

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

KEYWORDS = ["설문조사", "시민참여", "국민참여", "공모전"]
STATE_FILE = "state.json"
LOG_FILE = "monitor_log.json"

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "12"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "600"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.6.1)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}

def load_keywords():
    # 검색 키워드는 반드시 이 4개만 사용한다.
    return KEYWORDS.copy()

def normalize_state(raw):
    if isinstance(raw, dict):
        seen = raw.get("seen", {})
        if isinstance(seen, list):
            seen = {str(x): True for x in seen}
        if not isinstance(seen, dict):
            seen = {}
        return {
            "version": 861,
            "initialized": bool(raw.get("initialized", False)),
            "seen": seen,
        }
    if isinstance(raw, list):
        return {"version": 861, "initialized": False,
                "seen": {str(x): True for x in raw}}
    return {"version": 861, "initialized": False, "seen": {}}

def load_state():
    if not os.path.exists(STATE_FILE):
        return normalize_state({})
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return normalize_state(json.load(f))
    except Exception:
        return normalize_state({})

def save_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def clean_text(s):
    return re.sub(r"\s+", " ", s or "").strip()

def extract_posts(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    # 링크를 가진 요소를 후보로 수집. 게시판 구조가 달라도 최대한 보수적으로 잡는다.
    candidates = []
    seen_urls = set()
    for a in soup.find_all("a", href=True):
        title = clean_text(a.get_text(" ", strip=True))
        href = a.get("href", "")
        if not title or len(title) < 2:
            continue
        if title in {"이전", "다음", "목록", "검색", "확인", "닫기", "더보기"}:
            continue
        if href.startswith("javascript:") or href.startswith("#"):
            continue
        from urllib.parse import urljoin
        url = urljoin(base_url, href)
        if url in seen_urls:
            continue
        # 게시물 링크에서 흔히 쓰이는 패턴을 우선하되, 키워드가 포함된 링크는 허용
        low = (href + " " + title).lower()
        looks_post = any(x in low for x in [
            "view", "detail", "board", "bbs", "article", "notice", "seq=", "idx=", "ntt"
        ])
        if looks_post or any(k in title for k in KEYWORDS):
            seen_urls.add(url)
            candidates.append({"title": title, "url": url})
    return candidates[:RECENT_POSTS]

async def fetch(session, url):
    last_error = None
    for attempt in range(HTTP_RETRIES + 1):
        try:
            timeout = aiohttp.ClientTimeout(total=TIMEOUT_SECONDS)
            async with session.get(url, headers=HEADERS, timeout=timeout,
                                   allow_redirects=True, ssl=False) as r:
                if r.status >= 400:
                    return None, f"HTTP {r.status}"
                text = await r.text(errors="ignore")
                return text, None
        except Exception as e:
            last_error = f"{type(e).__name__}: {e}"
            if attempt < HTTP_RETRIES:
                await asyncio.sleep(0.8 * (attempt + 1))
    return None, "HTTP 0"

def make_key(institution, board, url):
    return f"{institution}|{board}|{url}"

async def process_target(session, sem, target):
    institution, board, url = target
    async with sem:
        html, err = await fetch(session, url)
        if err:
            return {
                "institution": institution, "board": board, "url": url,
                "posts": 0, "matches": [], "error": err
            }

        posts = extract_posts(html, url)
        matches = []
        # 목록에서 제목이 키워드에 걸리면 우선 매칭.
        for p in posts:
            title = p["title"]
            title_hits = [k for k in KEYWORDS if k in title]
            body = ""
            # 본문 검색: 개별 상세페이지를 가져오되, 시간 절약을 위해 제목 히트는 바로 기록.
            if title_hits:
                body = title
            else:
                # 상세 페이지는 후보를 제한해서 요청
                detail_html, detail_err = await fetch(session, p["url"])
                if detail_html:
                    dsoup = BeautifulSoup(detail_html, "html.parser")
                    body = clean_text(dsoup.get_text(" ", strip=True))
                else:
                    body = title

            hits = [k for k in KEYWORDS if k in (title + " " + body)]
            if hits:
                matches.append({
                    "institution": institution,
                    "board": board,
                    "keyword": hits,
                    "title": title,
                    "url": p["url"],
                })

        return {
            "institution": institution, "board": board, "url": url,
            "posts": len(posts), "matches": matches
        }

def load_targets():
    wb = load_workbook("monitor_targets.xlsx", read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    headers = [str(x).strip() if x is not None else "" for x in rows[0]]
    idx = {h: i for i, h in enumerate(headers)}
    required = ["기관명", "게시판", "URL"]
    # 이전 파일의 컬럼명 변형도 지원
    board_col = idx.get("게시판", idx.get("board", 1))
    url_col = idx.get("URL", idx.get("url", 2))
    inst_col = idx.get("기관명", idx.get("institution", 0))
    targets = []
    for row in rows[1:]:
        if len(row) <= max(inst_col, board_col, url_col):
            continue
        inst = str(row[inst_col] or "").strip()
        board = str(row[board_col] or "").strip()
        url = str(row[url_col] or "").strip()
        if inst and url:
            targets.append((inst, board, url))
    return targets[:100]

async def telegram_send(session, matches):
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id or not matches:
        return {"ok": False, "sent": 0, "reason": "credentials_or_no_matches"}

    sent = 0
    errors = []
    api = f"https://api.telegram.org/bot{token}/sendMessage"
    for m in matches[:TELEGRAM_MAX_SEND]:
        kws = ", ".join(m["keyword"])
        text = (
            f"🔔 공공기관 알림\n"
            f"기관: {m['institution']}\n"
            f"게시판: {m['board']}\n"
            f"키워드: {kws}\n"
            f"제목: {m['title']}\n"
            f"URL: {m['url']}"
        )
        try:
            async with session.post(api, data={"chat_id": chat_id, "text": text},
                                    timeout=aiohttp.ClientTimeout(total=15)) as r:
                data = await r.json(content_type=None)
                if r.status == 200 and data.get("ok") is True:
                    sent += 1
                else:
                    errors.append({
                        "status": r.status,
                        "description": data.get("description", "unknown")
                    })
        except Exception as e:
            errors.append({"status": 0, "description": f"{type(e).__name__}: {e}"})
        await asyncio.sleep(0.25)

    return {
        "ok": sent == min(len(matches), TELEGRAM_MAX_SEND) and not errors,
        "sent": sent,
        "attempted": min(len(matches), TELEGRAM_MAX_SEND),
        "errors": errors[:5],
        "truncated": len(matches) > TELEGRAM_MAX_SEND,
    }

async def main():
    started = time.monotonic()
    state = load_state()
    first_run = not state["initialized"]
    targets = load_targets()
    keywords = load_keywords()

    sem = asyncio.Semaphore(MAX_CONCURRENCY)
    connector = aiohttp.TCPConnector(limit=MAX_CONCURRENCY, ssl=False)
    async with aiohttp.ClientSession(connector=connector) as session:
        tasks = [asyncio.create_task(process_target(session, sem, t)) for t in targets]
        results = []
        try:
            results = await asyncio.wait_for(
                asyncio.gather(*tasks), timeout=MAX_TOTAL_SECONDS
            )
        except asyncio.TimeoutError:
            for t in tasks:
                if not t.done():
                    t.cancel()
            await asyncio.gather(*tasks, return_exceptions=True)
            results = [r for r in tasks if hasattr(r, "result") and r.done() and not r.cancelled()]

        all_matches = []
        errors = []
        completed = 0
        posts_checked = 0
        for r in results:
            if isinstance(r, dict):
                completed += 1
                posts_checked += r.get("posts", 0)
                if r.get("error"):
                    errors.append(r)
                all_matches.extend(r.get("matches", []))

        # 상태 기반 신규 매칭 판정
        new_matches = []
        for m in all_matches:
            # 게시물 단위로 중복 제거
            key = f"{m['institution']}|{m['board']}|{m['url']}"
            if key not in state["seen"]:
                if first_run:
                    state["seen"][key] = True
                else:
                    new_matches.append(m)
                    state["seen"][key] = True

        state["initialized"] = True
        # seen 상한
        if len(state["seen"]) > 10000:
            keys = list(state["seen"].keys())[-10000:]
            state["seen"] = {k: True for k in keys}
        save_json(STATE_FILE, state)

        # 매칭 상세를 로그에 저장
        log = {
            "version": "V8.6.1",
            "timestamp_utc": datetime.now(timezone.utc).isoformat(),
            "keywords": keywords,
            "targets": len(targets),
            "completed": completed,
            "posts_checked": posts_checked,
            "new_matches": len(new_matches),
            "errors": len(errors),
            "timed_out": completed < len(targets),
            "first_run": first_run,
            "discovered_notice_boards": 0,
            "matches": new_matches,
            "error_details": errors,
        }

        # Actions 콘솔에 상세 출력
        print("\n===== 신규 매칭 상세 =====")
        if new_matches:
            for i, m in enumerate(new_matches, 1):
                print(f"[NEW MATCH {i}]")
                print(f"기관: {m['institution']}")
                print(f"게시판: {m['board']}")
                print(f"키워드: {', '.join(m['keyword'])}")
                print(f"제목: {m['title']}")
                print(f"URL: {m['url']}")
                print("-" * 60)
        else:
            print("신규 매칭 없음")
        print("===== 신규 매칭 상세 끝 =====\n")

        tg = await telegram_send(session, new_matches)
        log["telegram"] = tg
        save_json(LOG_FILE, log)

        print(json.dumps(log, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    asyncio.run(main())
