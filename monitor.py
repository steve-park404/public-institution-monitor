VERSION = "V8.12"
import os, re, json, time, html, warnings
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, parse_qs
import requests
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import openpyxl

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# 운영 키워드는 정확히 3개만 사용
KEYWORDS = ["설문조사", "시민참여", "국민참여"]
EXCLUDE_TITLE = ["공모전"]

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "20"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
RECENT_DAYS = int(os.getenv("RECENT_DAYS", "30"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "1200"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))

# 하루 최대 Telegram 발송량. 초과분은 pending.json에 보관하여 다음 실행으로 이월.
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))
MAX_PENDING = int(os.getenv("MAX_PENDING", "10000"))

UA = "Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.11.1)"

NOISE = ["script","style","noscript","svg","header","footer","nav","aside","form","iframe","canvas","template"]
BOARD_WORDS = ["공지사항","공지","알림마당","알림","소식","새소식","기관소식","게시판","뉴스","보도자료","공고"]
URL_HINTS = ["notice","noti","board","bbs","news","announcement","plaza","inform"]
BAD = ["채용","입찰","계약","로그인","회원","사이트맵","개인정보","이용약관"]
PARTICIPATION_BOARD_WORDS = ["국민참여","시민참여","참여마당","고객참여","소통","설문","이벤트","공모전","동반성장","사회공헌"]
NOTICE_STRONG_WORDS = ["공지사항","새소식","알림마당","기관소식","보도자료","공고"]
PARTICIPATION_CONTEXT = ["설문","의견수렴","의견조사","만족도","조사","응답","설문지","참여단","시민의견","국민의견"]
KST = ZoneInfo("Asia/Seoul")

session = requests.Session()
session.headers.update({
    "User-Agent": UA,
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.5"
})


def norm(s):
    return re.sub(r"\s+", " ", html.unescape(str(s or ""))).strip()

def get(url):
    for i in range(HTTP_RETRIES + 1):
        try:
            r = session.get(url, timeout=TIMEOUT_SECONDS, allow_redirects=True)
            r.raise_for_status()
            if not r.encoding or r.encoding.lower() == "iso-8859-1":
                r.encoding = r.apparent_encoding
            return r
        except Exception:
            if i < HTTP_RETRIES:
                time.sleep(.4 * (i + 1))
    return None

def same_domain(a, b):
    try:
        return urlparse(a).netloc.lower().replace("www.", "") == urlparse(b).netloc.lower().replace("www.", "")
    except Exception:
        return False

def title_of(soup):
    return norm(soup.title.get_text(" ", strip=True) if soup.title else "")

def clean(soup):
    for tag in NOISE:
        for x in soup.find_all(tag):
            x.decompose()
    return soup

def extract_links(page_url, soup):
    out, seen = [], set()
    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        if not href or href.startswith(("javascript:", "mailto:", "tel:", "#")):
            continue
        u = urljoin(page_url, href)
        if not u.startswith(("http://", "https://")) or not same_domain(page_url, u):
            continue
        if u in seen:
            continue
        seen.add(u)
        out.append((u, norm(a.get_text(" ", strip=True))))
    return out

def detail_signal(u):
    x = u.lower()
    q = parse_qs(urlparse(x).query)
    if any(k.lower() in {"seq","no","idx","nttno","articleid","article_id","id","view"} for k in q):
        return True
    return any(v in x for v in ["/view", "/detail", "/read", "/article", "/contents/view", "/board/view"])

def detail_url(u):
    x = u.lower()
    if any(v in x for v in ["login", "member", "delete", "write", "modify"]):
        return False
    return detail_signal(u)

def board_score(u, t):
    s = 0
    x, tt = u.lower(), t.lower()
    s += sum(7 for w in BOARD_WORDS if w.lower() in tt)
    s += sum(3 for h in URL_HINTS if h in x)
    s -= sum(5 for b in BAD if b.lower() in tt)
    if detail_signal(u):
        s -= 4
    return s



DATE_PATTERNS = [
    re.compile(r"(?P<y>20\d{2})[-./년]\s*(?P<m>\d{1,2})[-./월]\s*(?P<d>\d{1,2})일?"),
    re.compile(r"(?P<y>\d{2})[-./]\s*(?P<m>\d{1,2})[-./]\s*(?P<d>\d{1,2})"),
]

def parse_date_text(text):
    text = norm(text)
    for pat in DATE_PATTERNS:
        m = pat.search(text)
        if not m:
            continue
        y, mo, d = int(m.group("y")), int(m.group("m")), int(m.group("d"))
        if y < 100:
            y += 2000
        try:
            return datetime(y, mo, d, tzinfo=KST)
        except ValueError:
            pass
    return None

def recent_cutoff():
    return datetime.now(KST) - timedelta(days=RECENT_DAYS)

def is_recent_date(dt):
    return bool(dt and dt >= recent_cutoff())

def row_context(a):
    # Prefer the nearest table/list row; fall back to a small parent container.
    for tag_name in ("tr", "li"):
        x = a.find_parent(tag_name)
        if x:
            return norm(x.get_text(" ", strip=True))
    x = a.parent
    for _ in range(3):
        if x is None:
            break
        txt = norm(x.get_text(" ", strip=True))
        if len(txt) >= 20:
            return txt
        x = x.parent
    return norm(a.get_text(" ", strip=True))

def extract_post_candidates(list_url, soup):
    """Return detail candidates with best-effort list-row date/title evidence."""
    out, seen = [], set()
    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        if not href:
            continue
        u = urljoin(list_url, href)
        if not u.startswith(("http://", "https://")) or not same_domain(list_url, u):
            continue
        if not detail_url(u) or u in seen:
            continue
        text = norm(a.get_text(" ", strip=True)) or norm(a.get("aria-label") or a.get("title"))
        if not text or text.lower() in {"prev","next","이전","다음","목록","검색","확인","닫기"}:
            continue
        ctx = row_context(a)
        dt = parse_date_text(ctx)
        seen.add(u)
        out.append({"url": u, "anchor_title": text, "date": dt.isoformat() if dt else None, "row_text": ctx[:500]})
    return out

def notice_board_score(url, page_text="", page_title=""):
    u = url.lower()
    txt = norm(f"{page_title} {page_text}").lower()
    score = 0
    for w in NOTICE_STRONG_WORDS:
        if w.lower() in txt:
            score += 8
    for h in URL_HINTS:
        if h in u:
            score += 2
    for w in PARTICIPATION_BOARD_WORDS:
        if w.lower() in txt:
            score -= 8
    for b in BAD:
        if b.lower() in txt:
            score -= 5
    return score

def board_identity_text(soup):
    if not soup:
        return ""
    parts = []
    if soup.title:
        parts.append(soup.title.get_text(" ", strip=True))
    for tag in soup.find_all(["h1", "h2", "h3"], limit=12):
        parts.append(tag.get_text(" ", strip=True))
    main = soup.find("main") or soup.find("article")
    if main:
        parts.append(main.get_text(" ", strip=True)[:1800])
    return norm(" ".join(parts))

def looks_like_wrong_board(url, soup):
    identity = board_identity_text(soup)
    score = notice_board_score(url, identity, title_of(soup) if soup else "")
    low = identity.lower()
    explicit_bad = any(w.lower() in low for w in PARTICIPATION_BOARD_WORDS)
    strong_notice = any(w.lower() in low for w in NOTICE_STRONG_WORDS)
    return explicit_bad and not strong_notice and score < 5

def recent_detail_urls(list_url, soup):
    items = extract_post_candidates(list_url, soup)
    dated = [x for x in items if x.get("date") and is_recent_date(datetime.fromisoformat(x["date"]))]
    # If dates are present, strictly enforce the 30-day window.
    if dated:
        dated.sort(key=lambda x: x["date"], reverse=True)
        return dated[:RECENT_POSTS]
    # No extractable dates: keep a very small bounded set. Detail-page date
    # validation in match_post remains mandatory, preventing old alerts.
    return items[:RECENT_POSTS]

def discover_board_v811(home_url, fetch_html_func):
    """
    V8.11: layered board discovery.
    Returns the highest-scoring verified board URL, or None.
    fetch_html_func(url) must return HTML text or None.
    """
    from bs4 import BeautifulSoup

    visited = set()
    queue = [(home_url, 0)]
    candidates = []

    while queue and len(visited) < BOARD_DISCOVERY_MAX_LINKS:
        url, depth = queue.pop(0)
        if not url or url in visited or depth > BOARD_DISCOVERY_DEPTH:
            continue
        visited.add(url)
        try:
            html = fetch_html_func(url)
        except Exception:
            html = None
        if not html:
            continue

        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a.get("href", "").strip()
            text = a.get_text(" ", strip=True)
            try:
                abs_url = urljoin(url, href)
            except Exception:
                continue
            if not abs_url.startswith(("http://", "https://")):
                continue
            if urlparse(abs_url).netloc != urlparse(home_url).netloc:
                continue

            if _v811_is_candidate_link(text, abs_url):
                score = _v811_score_link(text, abs_url)
                candidates.append((score, abs_url, text))
            elif depth < BOARD_DISCOVERY_DEPTH:
                # Follow useful-looking internal navigation pages.
                nt = _v811_norm(text)
                if nt and any(k.lower() in nt for k in ("알림", "소식", "참여", "게시", "고객", "정보")):
                    queue.append((abs_url, depth + 1))

    # De-duplicate and verify candidates by looking for list-like content.
    seen_urls = set()
    verified = []
    for score, url, text in sorted(candidates, key=lambda x: (-x[0], x[1])):
        if url in seen_urls or score < BOARD_DISCOVERY_MIN_SCORE:
            continue
        seen_urls.add(url)
        try:
            html = fetch_html_func(url)
        except Exception:
            html = None
        if not html:
            continue
        soup = BeautifulSoup(html, "html.parser")
        body_text = soup.get_text(" ", strip=True)
        # A board/list page normally has repeated links, pagination, or notice/list terms.
        link_count = len(soup.find_all("a", href=True))
        list_terms = sum(body_text.count(x) for x in ("공지", "번호", "제목", "등록일", "조회", "목록", "페이지"))
        if link_count >= 5 or list_terms >= 2:
            verified.append((score, url, text, link_count, list_terms))
            if len(verified) >= 5:
                break

    if not verified:
        return None

    verified.sort(key=lambda x: (-x[0], -x[4], -x[3], x[1]))
    return verified[0][1]


# =========================
# V8.11.1: targeted board-discovery diagnostics
# =========================
V8111_MAX_NAV_PAGES = int(os.getenv("V8111_MAX_NAV_PAGES", "8"))
V8111_MAX_CANDIDATES = int(os.getenv("V8111_MAX_CANDIDATES", "25"))
V8111_MIN_SCORE = int(os.getenv("V8111_MIN_SCORE", "4"))
V8111_DIAG_FILE = "board_discovery_log.json"

V8111_BOARD_TEXT = [
    "공지사항", "공지", "알림마당", "알림", "소식", "새소식", "기관소식",
    "게시판", "뉴스", "보도자료", "자료실", "고시", "공고", "참여", "소통",
    "국민참여", "시민참여", "고객참여"
]
V8111_BOARD_URL = [
    "notice", "noti", "board", "bbs", "news", "announcement",
    "plaza", "article", "list"
]
V8111_BAD_TEXT = ["채용", "입찰", "계약", "로그인", "회원", "사이트맵", "개인정보", "이용약관"]


def v8111_score(url, text):
    u, t = url.lower(), norm(text).lower()
    score = 0
    for w in V8111_BOARD_TEXT:
        if w.lower() in t:
            score += 7 if w in ("공지사항", "새소식", "알림마당", "게시판") else 3
    for h in V8111_BOARD_URL:
        if h in u:
            score += 2
    for w in PARTICIPATION_BOARD_WORDS:
        if w.lower() in t:
            score -= 7
    for b in V8111_BAD_TEXT:
        if b.lower() in t:
            score -= 5
    if detail_signal(url):
        score -= 4
    return score


def discover_board_v8111(home):
    """Targeted fallback for institutions still marked NO_BOARD.
    Returns (board_url, recent_detail_urls, diagnostic_record).
    It is deliberately bounded so the daily 355-site run does not balloon.
    """
    diag = {
        "home": home, "result": "HOME_ERROR", "nav_pages": 0,
        "candidates": 0, "verified": 0, "best_score": 0,
        "best_url": None, "candidate_samples": []
    }
    r = get(home)
    if not r:
        return None, [], diag

    visited_pages = {r.url}
    queue = []
    candidates = []
    first_links = extract_links(r.url, BeautifulSoup(r.text, "html.parser"))

    def add_candidate(u, tx):
        if not u.startswith(("http://", "https://")) or not same_domain(home, u):
            return
        sc = v8111_score(u, tx)
        if sc >= V8111_MIN_SCORE:
            candidates.append((sc, u, tx))

    # Home page: collect board candidates and a small number of navigation pages.
    for u, tx in first_links:
        add_candidate(u, tx)
        nt = norm(tx).lower()
        if any(k in nt for k in ("알림", "소식", "참여", "소통", "게시", "정보", "고객")):
            queue.append(u)

    # One bounded navigation hop. This catches sites where the board is behind a top menu.
    for nav in list(dict.fromkeys(queue))[:V8111_MAX_NAV_PAGES]:
        if nav in visited_pages:
            continue
        rr = get(nav)
        visited_pages.add(nav)
        diag["nav_pages"] += 1
        if not rr:
            continue
        for u, tx in extract_links(rr.url, BeautifulSoup(rr.text, "html.parser")):
            add_candidate(u, tx)

    # Deduplicate and verify candidate pages as actual list boards.
    unique = {}
    for sc, u, tx in candidates:
        if u not in unique or sc > unique[u][0]:
            unique[u] = (sc, u, tx)
    ranked = sorted(unique.values(), key=lambda x: (-x[0], x[1]))[:V8111_MAX_CANDIDATES]
    diag["candidates"] = len(unique)
    diag["candidate_samples"] = [
        {"score": sc, "url": u, "text": norm(tx)[:80]} for sc, u, tx in ranked[:8]
    ]

    verified = []
    for sc, u, tx in ranked:
        rr = get(u)
        if not rr:
            continue
        ss = BeautifulSoup(rr.text, "html.parser")
        details = recent_detail_urls(rr.url, ss)
        if looks_like_wrong_board(rr.url, ss):
            continue
        if len(details) >= 2:
            verified.append((sc, rr.url, [x["url"] for x in details], tx))

    diag["verified"] = len(verified)
    if verified:
        verified.sort(key=lambda x: (-x[0], -len(x[2]), x[1]))
        sc, board, details, tx = verified[0]
        diag["result"] = "VERIFIED"
        diag["best_score"] = sc
        diag["best_url"] = board
        return board, details, diag

    diag["result"] = "CANDIDATE_NOT_VERIFIED" if unique else "NO_CANDIDATE"
    return None, [], diag


def write_v8111_diag(records):
    try:
        with open(V8111_DIAG_FILE, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def discover_board(home):
    r = get(home)
    if not r:
        return None, [], "HOME_ERROR"
    soup = BeautifulSoup(r.text, "html.parser")
    links = extract_links(r.url, soup)
    scored = sorted([(board_score(u, t), u, t) for u, t in links], reverse=True)
    for score, u, t in scored[:35]:
        if score < 4:
            break
        rr = get(u)
        if not rr:
            continue
        ss = BeautifulSoup(rr.text, "html.parser")
        details = recent_detail_urls(rr.url, ss)
        if looks_like_wrong_board(rr.url, ss):
            continue
        if len(details) >= 2:
            return rr.url, [x["url"] for x in details], "DISCOVERED"
    return None, [], "NO_BOARD"

def visible_main_text(soup):
    soup = clean(soup)
    main = soup.find("main") or soup.find("article")
    if not main:
        candidates = soup.find_all(
            ["div", "section"],
            id=re.compile(r"(content|contents|sub|body|article)", re.I)
        )
        main = max(candidates, key=lambda x: len(x.get_text(" ", strip=True))) if candidates else soup.body
    if not main:
        return ""
    for x in main.find_all(["ul", "ol"]):
        txt = norm(x.get_text(" ", strip=True))
        if len(txt) < 400 and sum(1 for k in BOARD_WORDS if k in txt) >= 1:
            x.decompose()
    return norm(main.get_text(" ", strip=True))

def meaningful_body_match(body, kw):
    for m in re.finditer(re.escape(kw), body):
        a = max(0, m.start() - 160)
        b = min(len(body), m.end() + 220)
        ctx = body[a:b]
        if len(ctx) < 80:
            continue
        if kw in ("국민참여", "시민참여"):
            if not any(x in ctx for x in PARTICIPATION_CONTEXT):
                continue
        if any(p in ctx for p in [".", "다.", "요.", "습니다", "한다", "안내", "실시", "모집", "참여", "응답", "기간"]):
            return True
    return False

def extract_detail_date(soup):
    # Look at common metadata/date containers first, then visible text.
    selectors = [
        "time", ".date", ".regdate", ".reg_date", ".view_date", ".board-date",
        ".bbs-date", ".article-date", ".date-info", ".info", ".view_info"
    ]
    for sel in selectors:
        for x in soup.select(sel):
            dt = parse_date_text(x.get_text(" ", strip=True))
            if dt:
                return dt
    # Search only a bounded head/body prefix to avoid dates in attachments/history.
    txt = norm(soup.get_text(" ", strip=True))
    return parse_date_text(txt[:5000])

def match_post(u):
    r = get(u)
    if not r:
        return None
    soup = BeautifulSoup(r.text, "html.parser")
    title = title_of(soup)
    if not title:
        return None

    # Hard date gate: old posts can never alert.
    post_date = extract_detail_date(soup)
    if not is_recent_date(post_date):
        return None

    if any(x in title for x in EXCLUDE_TITLE):
        return None

    for kw in KEYWORDS:
        if kw in title:
            return {
                "url": r.url, "title": title, "keyword": kw,
                "where": "title", "date": post_date.strftime("%Y-%m-%d")
            }

    body = visible_main_text(soup)
    if len(body) < 120:
        return None
    for kw in KEYWORDS:
        if meaningful_body_match(body, kw):
            return {
                "url": r.url, "title": title, "keyword": kw,
                "where": "body", "date": post_date.strftime("%Y-%m-%d")
            }
    return None

def load_targets():
    wb = openpyxl.load_workbook("monitor_targets.xlsx", read_only=True, data_only=True)
    ws = wb["355기관"]
    hs = [c.value for c in next(ws.iter_rows())]
    ix = {str(v): i for i, v in enumerate(hs)}
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[ix["기관명"]]:
            continue
        out.append({
            "name": str(row[ix["기관명"]]).strip(),
            "home": str(row[ix["홈페이지URL"]] or "").strip(),
            "board": str(row[ix["공지게시판URL"]] or "").strip()
        })
    return out[:355]

def load_state():
    try:
        with open("state.json", "r", encoding="utf-8") as f:
            s = json.load(f)
    except Exception:
        s = {}
    if not isinstance(s, dict):
        s = {}
    if not isinstance(s.get("seen"), dict):
        s["seen"] = {str(x): 1 for x in s.get("seen", [])[-10000:]}
    if not isinstance(s.get("boards"), dict):
        s["boards"] = {}
    if not isinstance(s.get("stats"), dict):
        s["stats"] = {}
    return s

def save_state(s):
    s["seen"] = dict(list(s.get("seen", {}).items())[-10000:])
    with open("state.json", "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)

def load_pending():
    try:
        with open("pending.json", "r", encoding="utf-8") as f:
            p = json.load(f)
    except Exception:
        p = []
    if not isinstance(p, list):
        return []
    # 구버전/중복 데이터 방어
    out, seen = [], set()
    for m in p:
        if not isinstance(m, dict):
            continue
        key = match_key(m)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(m)
    return out[-MAX_PENDING:]

def save_pending(pending):
    with open("pending.json", "w", encoding="utf-8") as f:
        json.dump(pending[-MAX_PENDING:], f, ensure_ascii=False, indent=2)

def match_key(m):
    return f"{m.get('url','')}|{m.get('keyword','')}"

def priority(m):
    # 제목 매칭을 먼저, 키워드는 설문조사 → 시민참여 → 국민참여 순
    where_rank = 0 if m.get("where") == "title" else 1
    kw_rank = {k: i for i, k in enumerate(KEYWORDS)}.get(m.get("keyword"), 99)
    return (where_rank, kw_rank, m.get("기관명", ""), m.get("title", ""), m.get("url", ""))

def process(t, state, deadline):
    name, home, seed = t["name"], t["home"], t["board"]
    if time.time() >= deadline:
        return {"name": name, "status": "DEADLINE", "posts": 0, "matches": [], "diag": None}

    board = seed or state["boards"].get(name, "")
    details = []
    diag = None

    if board:
        rr = get(board)
        if rr:
            ss = BeautifulSoup(rr.text, "html.parser")
            if looks_like_wrong_board(rr.url, ss):
                board = ""
            else:
                details = [x["url"] for x in recent_detail_urls(rr.url, ss)]
        else:
            board = ""

    # Proven discovery first. V8.11.1 fallback is called ONLY when the proven
    # one fails, so normal sites keep the previous runtime characteristics.
    if not board and home:
        board, details, status = discover_board(home)

    if not board and home and time.time() < deadline:
        board, details, diag = discover_board_v8111(home)

    if not board:
        return {"name": name, "status": "NO_BOARD", "posts": 0, "matches": [], "diag": diag}

    matches = []
    for u in details:
        if time.time() >= deadline:
            break
        m = match_post(u)
        if m:
            m["기관명"] = name
            matches.append(m)

    return {
        "name": name,
        "status": "OK",
        "posts": len(details),
        "matches": matches,
        "board": board,
        "diag": diag
    }

def telegram_send(m):
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat:
        return False, "CONFIG_MISSING"

    text = (
        f"📢 공공기관 참여정보 알림\n"
        f"기관: {m['기관명']}\n"
        f"키워드: {m['keyword']}\n"
        f"제목: {m['title']}\n"
        f"검색위치: {m['where']}\n"
        f"{m['url']}"
    )
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat, "text": text},
            timeout=20
        )
        if r.ok and r.json().get("ok"):
            return True, "SENT"
        print("TELEGRAM_ERROR", r.status_code, r.text[:300])
        return False, f"HTTP_{r.status_code}"
    except Exception as e:
        print("TELEGRAM_EXCEPTION", repr(e))
        return False, "EXCEPTION"

def main():
    start = time.time()
    deadline = start + MAX_TOTAL_SECONDS

    targets = load_targets()
    state = load_state()
    pending = load_pending()

    pending_before = len(pending)
    results = []
    discovered = []

    # 병렬 수집 단계에서는 seen/pending을 변경하지 않음.
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as ex:
        fs = [ex.submit(process, t, state, deadline) for t in targets]
        for f in as_completed(fs):
            try:
                r = f.result()
            except Exception as e:
                r = {
                    "name": "?",
                    "status": "ERROR",
                    "posts": 0,
                    "matches": [],
                    "error": repr(e)
                }
            results.append(r)
            discovered.extend(r.get("matches", []))

    # Persist newly discovered boards for future runs.
    for r in results:
        if r.get("status") == "OK" and r.get("board"):
            state["boards"][r["name"]] = r["board"]

    # V8.11.1 diagnostics: only fallback attempts are recorded.
    diag_records = []
    for r in results:
        d = r.get("diag")
        if d:
            d = dict(d)
            d["institution"] = r.get("name")
            diag_records.append(d)
    write_v8111_diag(diag_records)
    diag_counts = {}
    for d in diag_records:
        k = d.get("result", "UNKNOWN")
        diag_counts[k] = diag_counts.get(k, 0) + 1
    print("BOARD_DISCOVERY_DIAG", json.dumps(diag_counts, ensure_ascii=False))

    # V8.12: revalidate legacy pending items so old/false alerts are purged.
    cleaned_pending = []
    pending_removed_invalid = 0
    for old in pending:
        try:
            refreshed = match_post(old.get("url", ""))
        except Exception:
            refreshed = None
        if refreshed:
            refreshed["기관명"] = old.get("기관명", refreshed.get("기관명", ""))
            cleaned_pending.append(refreshed)
        else:
            pending_removed_invalid += 1
    pending = cleaned_pending

    # 신규 매칭을 queue에 넣는다. 이미 seen 또는 pending이면 중복 삽입하지 않는다.
    seen_keys = set(state.get("seen", {}).keys())
    pending_keys = {match_key(m) for m in pending}

    new_matches = []
    for m in discovered:
        k = match_key(m)
        if not k or k in seen_keys or k in pending_keys:
            continue
        pending.append(m)
        pending_keys.add(k)
        new_matches.append(m)

    # 발송 우선순위: 제목 매칭 → 키워드 순 → 기관명/제목/URL
    pending.sort(key=priority)
    pending = pending[-MAX_PENDING:]

    sent_items = []
    send_attempts = min(TELEGRAM_MAX_SEND, len(pending))

    for m in pending[:send_attempts]:
        ok, status = telegram_send(m)
        if ok:
            sent_items.append(m)
            state["seen"][match_key(m)] = int(time.time())
        else:
            # 실패/미설정이면 queue에서 제거하지 않는다.
            if status == "CONFIG_MISSING":
                print("TELEGRAM_CONFIG_MISSING")
            break

    sent_keys = {match_key(m) for m in sent_items}
    if sent_keys:
        pending = [m for m in pending if match_key(m) not in sent_keys]

    save_pending(pending)
    save_state(state)

    summary = {
        "targets": len(targets),
        "completed": sum(r["status"] == "OK" for r in results),
        "no_board": sum(r["status"] == "NO_BOARD" for r in results),
        "posts_checked": sum(r.get("posts", 0) for r in results),
        "new_matches": len(new_matches),
        "pending_before": pending_before,
        "pending_removed_invalid": pending_removed_invalid,
        "recent_days": RECENT_DAYS,
        "telegram_sent": len(sent_items),
        "pending_after": len(pending),
        "errors": sum(r["status"] == "ERROR" for r in results),
        "elapsed_seconds": round(time.time() - start, 1),
        "timed_out": time.time() >= deadline,
        "board_discovery_diag": diag_counts
    }

    with open("monitor_log.json", "w", encoding="utf-8") as f:
        json.dump({
            "summary": summary,
            "sent": sent_items[-TELEGRAM_MAX_SEND:],
            "new_matches": new_matches[-500:],
            "pending_after": pending[-500:],
            "results": results
        }, f, ensure_ascii=False, indent=2)

    print("TELEGRAM_SENT", len(sent_items))
    print("SUMMARY", json.dumps(summary, ensure_ascii=False))

if __name__ == "__main__":
    main()
