import os
import re
import json
import hashlib
import asyncio
from datetime import datetime, timezone, timedelta
from urllib.parse import urljoin, urlparse

import aiohttp
from bs4 import BeautifulSoup

BASE = os.path.dirname(os.path.abspath(__file__))
TARGET_FILE = os.path.join(BASE, "monitor_targets.xlsx")
KEYWORD_FILE = os.path.join(BASE, "keywords.txt")
STATE_FILE = os.path.join(BASE, "state.json")
LOG_FILE = os.path.join(BASE, "monitor_log.json")

KST = timezone(timedelta(hours=9))

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "15"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "25"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "20"))
DISCOVER_NOTICE = os.getenv("DISCOVER_NOTICE", "true").lower() == "true"
SEED_ON_FIRST_RUN = os.getenv("SEED_ON_FIRST_RUN", "true").lower() == "true"

NOTICE_LABELS = [
    "공지사항", "공지", "알림마당", "기관소식", "새소식", "공고",
    "알림", "소식", "공지·공고", "공지/공고"
]
BAD_LINK_WORDS = [
    "로그인", "회원가입", "예약", "진료", "검색", "설문", "채용", "입찰",
    "자료실", "뉴스레터", "소식지", "이용안내", "찾아오시는길",
    "개인정보", "약관", "사이트맵", "교육", "신청", "민원", "발급",
    "결제", "일정", "문의", "다운로드"
]
DETAIL_HINTS = [
    "list_no=", "articleNo=", "article_no=", "seq=", "idx=", "bbsId=",
    "view.do", "view.jsp", "view.aspx", "act=view", "mode=view"
]

def now_kst():
    return datetime.now(KST).isoformat()

def load_keywords():
    if not os.path.exists(KEYWORD_FILE):
        return []
    out = []
    for line in open(KEYWORD_FILE, encoding="utf-8"):
        s = line.strip()
        if s and not s.startswith("#"):
            out.append(s)
    return list(dict.fromkeys(out))

def load_state():
    if not os.path.exists(STATE_FILE):
        return {"seen": {}, "initialized": False}
    try:
        return json.load(open(STATE_FILE, encoding="utf-8"))
    except Exception:
        return {"seen": {}, "initialized": False}

def save_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def clean_text(s):
    return re.sub(r"\s+", " ", s or "").strip()

def normalize_url(url):
    if not url:
        return ""
    return url.strip().replace(" ", "%20")

def is_detail_url(url):
    low = url.lower()
    return any(x.lower() in low for x in DETAIL_HINTS)

def looks_like_notice_label(text):
    t = clean_text(text)
    if not t:
        return False
    return any(label in t for label in NOTICE_LABELS)

def is_bad_link_text(text):
    t = clean_text(text)
    return any(x in t for x in BAD_LINK_WORDS)

def title_is_bad(title):
    t = clean_text(title)
    if len(t) < 3 or len(t) > 250:
        return True
    bad_exact = {
        "이전글이 없습니다.", "다음글이 없습니다.", "스킵네비게이션",
        "메뉴", "본문", "홈", "공지사항 상세", "새소식 공지사항 상세"
    }
    if t in bad_exact:
        return True
    if any(x in t.lower() for x in ["skip navigation", "privacy policy"]):
        return True
    return False

def get_title_from_anchor(a):
    for attr in ("title", "aria-label"):
        v = clean_text(a.get(attr))
        if v:
            return v
    return clean_text(a.get_text(" ", strip=True))

def extract_list_posts(soup, base_url):
    candidates = []

    # 우선 table/list의 반복 행에서 추출
    containers = soup.select("table, ul, ol, .board-list, .bbs-list, .list, [class*='board'], [class*='bbs']")
    seen = set()

    for container in containers:
        anchors = container.find_all("a", href=True)
        local = []
        for a in anchors:
            title = get_title_from_anchor(a)
            href = normalize_url(urljoin(base_url, a.get("href")))
            if not title or not href or href in seen:
                continue
            if title_is_bad(title) or is_bad_link_text(title):
                continue
            if href.rstrip("/") == base_url.rstrip("/"):
                continue
            if href.lower().startswith(("javascript:", "mailto:", "tel:")):
                continue
            # 메뉴/기능 링크 제외
            cls = " ".join(a.get("class", [])).lower()
            parent_text = clean_text(a.parent.get_text(" ", strip=True)) if a.parent else ""
            if any(x in (cls + " " + title.lower()) for x in [
                "login", "search", "menu", "gnb", "lnb", "reservation"
            ]):
                continue
            local.append((title, href))
        # 같은 컨테이너에서 여러 게시물 후보가 나오는 경우에만 채택
        if len(local) >= 3:
            for x in local:
                seen.add(x[1])
                candidates.append(x)

    # fallback: 페이지의 앵커 중 상세 URL 패턴이 명확한 것
    if len(candidates) < 3:
        for a in soup.find_all("a", href=True):
            title = get_title_from_anchor(a)
            href = normalize_url(urljoin(base_url, a.get("href")))
            if title_is_bad(title) or is_bad_link_text(title):
                continue
            if href.lower().startswith(("javascript:", "mailto:", "tel:")):
                continue
            if is_detail_url(href) and href not in seen:
                seen.add(href)
                candidates.append((title, href))

    # URL 중복 제거 + 제목 다양성 확보
    out = []
    seen_urls = set()
    for title, href in candidates:
        if href in seen_urls:
            continue
        seen_urls.add(href)
        out.append({"title": title, "url": href})
        if len(out) >= RECENT_POSTS:
            break
    return out

def extract_body_text(soup):
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    return clean_text(soup.get_text(" ", strip=True))

def discover_notice_boards(home_url, soup):
    found = []
    seen = set()
    for a in soup.find_all("a", href=True):
        text = get_title_from_anchor(a)
        href = normalize_url(urljoin(home_url, a.get("href")))
        if not looks_like_notice_label(text):
            continue
        if is_bad_link_text(text) and "공지" not in text:
            continue
        if href.lower().startswith(("javascript:", "mailto:", "tel:")):
            continue
        if is_detail_url(href):
            continue
        # 외부 도메인 제외
        if urlparse(href).netloc and urlparse(href).netloc != urlparse(home_url).netloc:
            continue
        if href not in seen:
            seen.add(href)
            found.append((text, href))
    return found

async def fetch(session, url):
    try:
        async with session.get(
            url,
            timeout=aiohttp.ClientTimeout(total=TIMEOUT_SECONDS),
            allow_redirects=True,
            headers={
                "User-Agent": "Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.1)"
            },
            ssl=False,
        ) as r:
            text = await r.text(errors="ignore")
            return r.status, str(r.url), text
    except Exception as e:
        return 0, url, f"{type(e).__name__}: {e}"

async def verify_notice_board(session, label, url):
    status, final_url, html = await fetch(session, url)
    if status != 200:
        return None, {"label": label, "url": url, "error": f"HTTP {status}"}
    soup = BeautifulSoup(html, "html.parser")
    posts = extract_list_posts(soup, final_url)
    titles = {clean_text(p["title"]) for p in posts if not title_is_bad(p["title"])}
    # 게시물 3개 이상 + 서로 다른 제목 3개 이상이면 실제 목록으로 인정
    if len(posts) >= 3 and len(titles) >= 3:
        return {"label": label, "url": final_url, "posts": posts}, None
    return None, {"label": label, "url": url, "error": f"NOT_BOARD posts={len(posts)} unique_titles={len(titles)}"}

async def verify_post(session, post):
    status, final_url, html = await fetch(session, post["url"])
    if status != 200:
        return False
    soup = BeautifulSoup(html, "html.parser")
    body = extract_body_text(soup)
    title = clean_text(post["title"])
    # 상세 페이지의 제목이 별도로 깨져도 목록 제목을 기준으로 검증
    if len(title) >= 4 and title in body:
        return True
    # 제목의 핵심어 2개 이상이 본문에 존재하면 보조 인정
    words = [w for w in re.split(r"\s+", title) if len(w) >= 2]
    hits = sum(1 for w in words[:8] if w in body)
    return hits >= 2 and len(body) >= 100

async def process_target(session, sem, row, keywords):
    async with sem:
        institution = clean_text(str(row.get("기관명", "")))
        board = clean_text(str(row.get("게시판명", "")))
        url = normalize_url(str(row.get("게시판URL", "")))
        if not url or url.lower() == "nan":
            return [], {"기관명": institution, "게시판": board, "URL": url, "error": "EMPTY_URL"}

        status, final_url, html = await fetch(session, url)
        if status != 200:
            return [], {"기관명": institution, "게시판": board, "URL": url, "error": f"HTTP {status}"}

        soup = BeautifulSoup(html, "html.parser")
        posts = extract_list_posts(soup, final_url)
        results = []

        for p in posts:
            matched = [kw for kw in keywords if kw.casefold() in p["title"].casefold()]
            if not matched:
                ok = await verify_post(session, p)
                if ok:
                    # 본문 확인
                    st2, _, html2 = await fetch(session, p["url"])
                    if st2 == 200:
                        s2 = BeautifulSoup(html2, "html.parser")
                        body = extract_body_text(s2)
                        matched = [kw for kw in keywords if kw.casefold() in body.casefold()]
            if matched:
                key = hashlib.sha256(
                    f"{institution}|{p['title']}|{p['url']}".encode("utf-8")
                ).hexdigest()
                results.append({
                    "key": key, "기관명": institution, "게시판": board,
                    "title": p["title"], "url": p["url"], "keywords": matched
                })
        return results, None

async def send_telegram(matches):
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        return {"ok": False, "error": "TELEGRAM secrets missing"}

    api = f"https://api.telegram.org/bot{token}/sendMessage"
    text = "🚨 공공기관 공지사항 키워드 발견\n\n"
    for m in matches[:20]:
        text += f"🏢 {m['기관명']}\n📌 {m['게시판']}\n📝 {m['title']}\n🔑 {', '.join(m['keywords'])}\n🔗 {m['url']}\n\n"

    async with aiohttp.ClientSession() as session:
        try:
            async with session.post(api, data={"chat_id": chat_id, "text": text}, timeout=20) as r:
                body = await r.text()
                if r.status != 200:
                    return {"ok": False, "error": f"HTTP {r.status}: {body}"}
                return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": f"{type(e).__name__}: {e}"}

async def main():
    import pandas as pd

    keywords = load_keywords()
    state = load_state()

    if not keywords:
        print("No keywords configured.")
        return

    if not os.path.exists(TARGET_FILE):
        print(f"Target file not found: {TARGET_FILE}")
        return

    df = pd.read_excel(TARGET_FILE)
    rows = df.fillna("").to_dict("records")

    timeout = aiohttp.ClientTimeout(total=TIMEOUT_SECONDS)
    connector = aiohttp.TCPConnector(limit=MAX_CONCURRENCY, limit_per_host=4, ssl=False)
    sem = asyncio.Semaphore(MAX_CONCURRENCY)

    discovered = []
    discovery_errors = []

    async with aiohttp.ClientSession(timeout=timeout, connector=connector) as session:
        # 홈페이지에서 공지사항 후보 발견: 별도 검색 대상으로 추가하기 전에 실제 게시판 검증
        if DISCOVER_NOTICE and os.path.exists(os.path.join(BASE, "url_완성.xlsx")):
            hdf = pd.read_excel(os.path.join(BASE, "url_완성.xlsx"))
            for r in hdf.fillna("").to_dict("records"):
                institution = clean_text(str(r.get("기관명", "")))
                home = normalize_url(str(r.get("URL", "")))
                if not home or home.lower() == "nan":
                    continue
                st, final_home, html = await fetch(session, home)
                if st != 200:
                    continue
                soup = BeautifulSoup(html, "html.parser")
                for label, notice_url in discover_notice_boards(final_home, soup):
                    verified, err = await verify_notice_board(session, label, notice_url)
                    if verified:
                        discovered.append({
                            "기관명": institution,
                            "게시판명": label,
                            "게시판URL": verified["url"],
                            "priority": 1,
                            "source": "homepage_discovery"
                        })
                    elif err:
                        discovery_errors.append({"기관명": institution, **err})

        # 기존 대상 + 검증된 자동발견 게시판만 병합
        all_rows = rows + discovered
        unique = {}
        for r in all_rows:
            u = normalize_url(str(r.get("게시판URL", "")))
            if not u or u.lower() == "nan":
                continue
            key = (clean_text(str(r.get("기관명", ""))), u)
            if key not in unique:
                unique[key] = r
            else:
                # 공지사항 명칭을 우선
                old = unique[key]
                if "공지" in clean_text(str(r.get("게시판명", ""))) and "공지" not in clean_text(str(old.get("게시판명", ""))):
                    unique[key] = r

        targets = list(unique.values())
        print(f"Monitoring {len(targets)} verified board URLs")
        print(f"Homepage notice discoveries verified: {len(discovered)}")

        tasks = [process_target(session, sem, r, keywords) for r in targets]
        outputs = await asyncio.gather(*tasks)

    all_matches = []
    errors = []
    posts_checked = 0

    for matches, err in outputs:
        if matches:
            all_matches.extend(matches)
            posts_checked += len(matches)
        if err:
            errors.append(err)

    first_run = not state.get("initialized", False)
    new_matches = []

    for m in all_matches:
        if m["key"] not in state.get("seen", {}):
            state.setdefault("seen", {})[m["key"]] = {
                "seen_at": now_kst(),
                "title": m["title"],
                "url": m["url"]
            }
            if not first_run or not SEED_ON_FIRST_RUN:
                new_matches.append(m)

    state["initialized"] = True
    state["last_run"] = now_kst()
    save_json(STATE_FILE, state)

    telegram_result = {"ok": True, "skipped": True}
    if new_matches:
        telegram_result = await send_telegram(new_matches)

    log = {
        "run_at": now_kst(),
        "targets": len(targets),
        "posts_checked": posts_checked,
        "new_matches": len(new_matches),
        "errors": len(errors),
        "first_run": first_run,
        "discovered_notice_boards": len(discovered),
        "telegram": telegram_result,
        "errors_detail": errors[:100]
    }
    save_json(LOG_FILE, log)
    print(json.dumps(log, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    asyncio.run(main())
