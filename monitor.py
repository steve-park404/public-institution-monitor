# -*- coding: utf-8 -*-
"""
Public Institution Monitor V8.12.6

Architecture:
355기관 → 공식 홈페이지 → 공지/새소식/알림/참여 관련 게시판 탐색
→ 최근 30일 게시물 → 상세페이지 → 제목/본문 키워드 매칭 → Telegram

Operational keywords:
- 설문조사
- 시민참여
- 국민참여

Title exclusion:
- 공모전

V8.12.6 changes:
1. 최근 30일 날짜 게이트를 최종 상세페이지에서 강제
2. 1순위 게시판 키워드 12개를 우선 탐색
3. 참여 게시판을 무조건 배제하지 않고 점수로 처리
4. 채용/입찰/계약/자료실/교육/공모전/동반성장/사회공헌/구매 등은 강제 배제
5. 게시판 검증을 "최근 게시물 2개 이상"에서 "최근 게시물 1개 이상"으로 완화
6. 게시물 제목도 게시판 점수에 반영
7. V8.11.1의 누락된 fallback 함수/상수 의존성 제거
8. 진단 로그 강화
9. pending queue 유지 및 하루 Telegram 최대 20건
10. 실제 목록 URL 보수적 차단
11. 기관별 상세 URL 패턴 학습
12. 기관별 상태/오류 유형 집계 및 미처리 기관 수 진단
13. 일반 메뉴명이 제목으로 추출되어도 게시물 구조가 명확하면 본문 검사를 계속
"""

VERSION = "V8.12.7"

import os
import re
import json
import time
import html
import warnings
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, parse_qs

import requests
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import openpyxl

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# -----------------------------
# Config
# -----------------------------
KEYWORDS = ["설문조사", "시민참여", "국민참여"]
EXCLUDE_TITLE = ["공모전"]

MAX_CONCURRENCY = 20
TIMEOUT_SECONDS = 15
HTTP_RETRIES = 2

RECENT_POSTS = 20
RECENT_DAYS = 30

MAX_TOTAL_SECONDS = 1200
TELEGRAM_MAX_SEND = 20
MAX_PENDING = 10000

BOARD_DISCOVERY_MAX_LINKS = 100
BOARD_DISCOVERY_MAX_FETCH = 35
BOARD_DISCOVERY_DEPTH = 1

UA = f"Mozilla/5.0 (compatible; PublicInstitutionMonitor/{VERSION})"

NOISE = [
    "script", "style", "noscript", "svg", "header", "footer", "nav",
    "aside", "form", "iframe", "canvas", "template"
]

FIRST_PRIORITY_BOARD_WORDS = [
    "공지사항", "공지", "새소식", "알림마당", "알림", "이벤트",
    "설문", "설문조사", "국민참여", "시민참여", "참여마당", "소통"
]

BOARD_WORDS = FIRST_PRIORITY_BOARD_WORDS + [
    "소식", "기관소식", "게시판", "뉴스", "보도자료", "공고"
]

URL_HINTS = [
    "notice", "noti", "board", "bbs", "news", "announcement",
    "plaza", "inform", "ntt", "article", "list"
]

HARD_EXCLUDE_BOARD_WORDS = [
    "채용", "입찰", "계약", "자료실", "교육", "공모전",
    "동반성장", "사회공헌", "구매", "구매계약"
]

BAD = HARD_EXCLUDE_BOARD_WORDS + [
    "로그인", "회원", "사이트맵", "개인정보", "이용약관"
]

PARTICIPATION_BOARD_WORDS = [
    "국민참여", "시민참여", "참여마당", "고객참여", "소통",
    "설문", "이벤트"
]

NOTICE_STRONG_WORDS = FIRST_PRIORITY_BOARD_WORDS

PARTICIPATION_CONTEXT = [
    "설문", "의견수렴", "의견조사", "만족도", "조사", "응답",
    "설문지", "참여단", "시민의견", "국민의견", "의견",
    "참여기간", "응답기간", "조사기간", "참여방법", "응답방법",
    "참여해 주세요", "응답해 주세요", "설문에 참여", "조사에 참여",
    "의견을 제출", "의견을 남겨", "설문링크", "조사대상"
]

SURVEY_ACTION_CONTEXT = [
    "설문", "조사", "응답", "만족도", "의견수렴", "의견조사",
    "설문지", "참여기간", "응답기간", "조사기간", "참여방법",
    "응답방법", "참여해 주세요", "응답해 주세요", "설문에 참여",
    "조사에 참여", "의견을 제출", "의견을 남겨", "설문링크",
    "조사대상", "응답자", "참여자"
]

GENERIC_PAGE_TITLES = {
    "사이트맵", "사이트 맵", "알림마당", "공지사항", "공지", "새소식",
    "사업소개", "연구", "국민소통", "시민참여", "국민참여", "참여마당",
    "개인정보처리방침", "개인정보 처리방침", "이용약관", "로그인", "회원가입"
}

GENERIC_URL_HINTS = [
    "sitemap", "privacy", "terms", "login", "contents.do", "programproposal"
]

DATE_SELECTORS = [
    ".date", ".regdate", ".reg-date", ".write-date", ".wdate",
    ".board-date", ".bbs-date", ".ntt-date", ".article-date",
    "[class*='date']", "[class*='Date']",
    "[class*='regist']", "[class*='Regist']",
    "[class*='write']", "[class*='Write']",
    "time", "td"
]

KST = ZoneInfo("Asia/Seoul")

STATE_FILE = "state.json"
PENDING_FILE = "pending.json"
TARGET_FILE_CANDIDATES = [
    "monitor_targets.xlsx",
    "url.xlsx",
    "targets.xlsx",
]

# -----------------------------
# HTTP
# -----------------------------
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": UA,
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7",
})

def get(url):
    if not url:
        return None

    for attempt in range(HTTP_RETRIES + 1):
        try:
            r = SESSION.get(
                url,
                timeout=TIMEOUT_SECONDS,
                allow_redirects=True,
                verify=True,
            )
            if r.status_code >= 200 and r.status_code < 400:
                if not r.encoding or r.encoding.lower() == "iso-8859-1":
                    r.encoding = r.apparent_encoding or "utf-8"
                return r
        except Exception:
            pass

        if attempt < HTTP_RETRIES:
            time.sleep(0.35 * (attempt + 1))

    return None

# -----------------------------
# General helpers
# -----------------------------
def norm(s):
    return re.sub(r"\s+", " ", html.unescape(str(s or ""))).strip()

def lower_url(u):
    return (u or "").lower()

def same_domain(a, b):
    try:
        da = urlparse(a).netloc.lower().split(":")[0]
        db = urlparse(b).netloc.lower().split(":")[0]
        if da.startswith("www."):
            da = da[4:]
        if db.startswith("www."):
            db = db[4:]
        return da == db
    except Exception:
        return False

def absolute(base, href):
    try:
        return urljoin(base, href)
    except Exception:
        return ""

def clean_url(u):
    try:
        p = urlparse(u)
        return p._replace(fragment="").geturl()
    except Exception:
        return u

def text_of(tag):
    if not tag:
        return ""
    return norm(tag.get_text(" ", strip=True))

def looks_like_wrong_board(url, soup):
    text = ""
    title = ""

    if soup:
        title = text_of(soup.title)
        text = norm(soup.get_text(" ", strip=True))[:12000]

    sample = f"{url} {title} {text[:2500]}".lower()

    for bad in HARD_EXCLUDE_BOARD_WORDS:
        if bad.lower() in sample:
            # If the page is clearly a generic notice page but merely
            # mentions "교육" or "구매" in body text, don't reject solely
            # from body text. Reject when it appears in URL/title or
            # several times near navigation/list labels.
            u = lower_url(url)
            if bad.lower() in u or bad.lower() in title.lower():
                return True

    return False

# -----------------------------
# Date parsing
# -----------------------------
DATE_PATTERNS = [
    r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})",
    r"(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일",
    r"(20\d{2})(\d{2})(\d{2})",
]

def parse_date_text(s):
    s = norm(s)
    if not s:
        return None

    # Avoid interpreting ordinary numeric IDs as dates.
    for pat in DATE_PATTERNS:
        m = re.search(pat, s)
        if not m:
            continue
        try:
            y, mo, d = map(int, m.groups())
            if 2000 <= y <= 2099 and 1 <= mo <= 12 and 1 <= d <= 31:
                return datetime(y, mo, d, tzinfo=KST)
        except Exception:
            pass

    return None

def recent_cutoff():
    now = datetime.now(KST)
    return now - timedelta(days=RECENT_DAYS)

def is_recent_date(dt):
    if not dt:
        return False
    now = datetime.now(KST)
    return recent_cutoff() <= dt <= now + timedelta(days=1)

# -----------------------------
# Text extraction
# -----------------------------
def visible_main_text(soup):
    if not soup:
        return ""

    for tag in soup.find_all(NOISE):
        tag.decompose()

    candidates = []

    for selector in ["main", "article"]:
        for tag in soup.select(selector):
            t = text_of(tag)
            if len(t) >= 100:
                candidates.append(t)

    for tag in soup.find_all(["div", "section"]):
        ident = f"{tag.get('id','')} {' '.join(tag.get('class',[]) or [])}".lower()
        if re.search(r"(content|contents|sub|body|article|board|bbs|view|detail)", ident):
            t = text_of(tag)
            if len(t) >= 120:
                candidates.append(t)

    if candidates:
        candidates.sort(key=len, reverse=True)
        return candidates[0]

    body = soup.body or soup
    return text_of(body)

def meaningful_body_match(body, keyword):
    body = norm(body)
    if len(body) < 120:
        return False

    idxs = [m.start() for m in re.finditer(re.escape(keyword), body, re.I)]
    if not idxs:
        return False

    for idx in idxs:
        left = max(0, idx - 350)
        right = min(len(body), idx + 500)
        ctx = body[left:right]

        # 국민참여/시민참여는 일반 메뉴명이나 기관의 상시 참여창구가
        # 본문에 섞이는 경우가 많으므로 실제 조사/응답 행동 문맥을 요구.
        if keyword in ("국민참여", "시민참여"):
            if not any(x in ctx for x in PARTICIPATION_CONTEXT):
                continue

        # 설문조사 역시 단어 하나만 존재하는 정책/개인정보/사이트 안내는 제외.
        # 주변 문맥에서 실제 조사/응답 행위를 나타내는 표현을 요구한다.
        if keyword == "설문조사":
            if not any(x in ctx for x in SURVEY_ACTION_CONTEXT):
                continue

        sentence_markers = [
            "안내", "참여", "신청", "설문", "의견", "조사",
            "응답", "만족도", "기간", "대상", "방법", "모집",
            "제출", "온라인", "링크"
        ]
        if sum(1 for x in sentence_markers if x in ctx) >= 1:
            return True

    return False

# -----------------------------
# Detail date/title
# -----------------------------
def extract_title(soup):
    if not soup:
        return ""

    # 실제 게시물 제목 영역을 우선한다. 메뉴/브레드크럼의 h1을 잘못
    # 가져오는 것을 막기 위해 일반적인 UI 영역은 제외한다.
    selectors = [
        ".view-title", ".board-title", ".article-title", ".bbs-title",
        ".board_view .subject", ".boardView .subject", ".view .subject",
        ".view_subject", ".viewSubject", ".subject",
        "article h1", "article h2", "main h1", "main h2",
        "[class*='view'][class*='title']",
        "[class*='board'][class*='title']",
        "[class*='article'][class*='title']"
    ]

    vals = []
    for sel in selectors:
        try:
            tags = soup.select(sel)[:10]
        except Exception:
            tags = []
        for tag in tags:
            t = text_of(tag)
            if 2 <= len(t) <= 300 and t not in GENERIC_PAGE_TITLES:
                vals.append(t)

    if vals:
        # 게시물 제목은 지나치게 짧은 메뉴명보다 적당한 길이의 후보를 우선.
        vals.sort(key=lambda x: (len(x) < 4, len(x) > 150, len(x)))
        return vals[0]

    # h1/h2 전체 후보에서 generic 제목을 제외
    for tag in soup.find_all(["h1", "h2", "h3"])[:30]:
        t = text_of(tag)
        if 4 <= len(t) <= 300 and t not in GENERIC_PAGE_TITLES:
            return t

    # og:title은 보조 수단.
    og = soup.select_one("meta[property='og:title']")
    if og and og.get("content"):
        t = norm(og.get("content"))
        if t and t not in GENERIC_PAGE_TITLES:
            return t[:300]

    return ""

def is_probable_content_page(url, soup, title=""):
    u = lower_url(url)
    t = norm(title)
    # 명백한 콘텐츠/정책/사이트맵/개인정보 등은 게시물로 취급하지 않는다.
    if any(h in u for h in GENERIC_URL_HINTS):
        return True
    if t in GENERIC_PAGE_TITLES:
        return True
    if soup:
        pt = text_of(soup.title)
        if pt in GENERIC_PAGE_TITLES:
            return True
    return False

def has_post_structure(soup, title=""):
    if not soup:
        return False
    text = visible_main_text(soup)
    if len(text) < 120:
        return False

    signals = 0
    raw = norm(soup.get_text(" ", strip=True))[:15000]
    if title and title not in GENERIC_PAGE_TITLES:
        signals += 1
    if re.search(r"(등록일|작성일|게시일|작성자|조회수|첨부파일|첨부|이전글|다음글)", raw):
        signals += 1
    if soup.select(".file, .attach, [class*='attach'], [class*='file']"):
        signals += 1
    if soup.select("article, [class*='view'], [class*='board-view'], [class*='bbs-view'], [class*='contents']"):
        signals += 1
    return signals >= 2

def extract_detail_date(soup):
    if not soup:
        return None

    # meta/time 우선
    for tag in soup.find_all(["meta", "time"])[:100]:
        attrs = " ".join(str(v) for v in tag.attrs.values())
        content = tag.get("content") or tag.get("datetime") or tag.get_text(" ", strip=True)
        dt = parse_date_text(f"{attrs} {content}")
        if dt:
            return dt

    for sel in DATE_SELECTORS:
        try:
            tags = soup.select(sel)
        except Exception:
            tags = []

        for tag in tags[:30]:
            dt = parse_date_text(text_of(tag))
            if dt:
                return dt

    # 상세 본문 상단에서 날짜 탐색
    txt = visible_main_text(soup)
    return parse_date_text(txt[:5000])

# -----------------------------
# Link extraction / board discovery
# -----------------------------
def link_label(a):
    parts = [
        text_of(a),
        a.get("title", ""),
        a.get("aria-label", ""),
    ]
    return norm(" ".join(parts))

def extract_links(base_url, soup):
    out = []
    seen = set()

    if not soup:
        return out

    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        if not href or href.lower().startswith(("javascript:", "#", "mailto:", "tel:")):
            continue

        u = clean_url(absolute(base_url, href))
        if not u or not same_domain(base_url, u):
            continue

        label = link_label(a)
        key = (u, label)
        if key in seen:
            continue
        seen.add(key)
        out.append((u, label))

        if len(out) >= BOARD_DISCOVERY_MAX_LINKS:
            break

    return out

def hard_exclude_score(u, title):
    s = f"{u} {title}".lower()
    score = 0
    for x in HARD_EXCLUDE_BOARD_WORDS:
        if x.lower() in s:
            score -= 14
    return score

def board_score(u, title, recent_titles=""):
    s = f"{u} {title}".lower()
    score = 0

    for w in FIRST_PRIORITY_BOARD_WORDS:
        if w.lower() in s:
            score += 18

    for w in BOARD_WORDS:
        if w.lower() in s:
            score += 7

    for h in URL_HINTS:
        if h in lower_url(u):
            score += 3

    score += hard_exclude_score(u, title)

    # 참여 게시판은 허용하되 공지형 게시판보다 우선순위를 낮추는 효과
    # 를 주기 위해 별도 가산은 하지 않는다.
    if recent_titles:
        rt = recent_titles.lower()
        for kw in KEYWORDS:
            if kw.lower() in rt:
                score += 3

    return score

def notice_board_score(u, title):
    s = f"{u} {title}".lower()
    score = 0

    for w in FIRST_PRIORITY_BOARD_WORDS:
        if w.lower() in s:
            score += 12

    for w in ["소식", "기관소식", "게시판", "뉴스", "보도자료", "공고"]:
        if w.lower() in s:
            score += 5

    for h in URL_HINTS:
        if h in lower_url(u):
            score += 2

    for x in HARD_EXCLUDE_BOARD_WORDS:
        if x.lower() in s:
            score -= 12

    return score

# -----------------------------
# List-page post candidate extraction
# -----------------------------
def looks_like_detail_link(u, title):
    """게시물 후보 링크인지 넓게 잡는다. 최종 진위 판정은 상세페이지에서 한다."""
    s = f"{u} {title}".lower()

    if any(x in s for x in ["login", "logout", "sitemap", "privacy", "terms"]):
        return False

    # 게시물 링크 자체에 채용/입찰 등이 포함됐다는 이유만으로
    # 공지게시판의 정상 게시물을 버리지 않는다. 게시판 탐색 단계의
    # HARD_EXCLUDE와 게시물 후보 단계는 분리한다.
    if any(x in lower_url(u) for x in [
        "view", "read", "detail", "article", "ntt", "bbs",
        "board", "idx=", "seq=", "no=", "mode=view", "view.do", "read.do"
    ]):
        return True

    return 2 <= len(norm(title)) <= 300

def extract_post_candidates(base_url, soup):
    """
    반환:
      [{"url": ..., "title": ..., "date": ...}, ...]
    """
    out = []
    seen = set()

    if not soup:
        return out

    # 표 형식 게시판
    for row in soup.find_all("tr"):
        links = row.find_all("a", href=True)
        if not links:
            continue

        row_text = text_of(row)
        row_date = parse_date_text(row_text)

        for a in links:
            title = link_label(a)
            u = clean_url(absolute(base_url, a.get("href", "")))

            if not u or not same_domain(base_url, u):
                continue
            if not looks_like_detail_link(u, title):
                continue

            key = u
            if key in seen:
                continue

            seen.add(key)
            out.append({
                "url": u,
                "title": title,
                "date": row_date,
            })

    # 일반 div/ul/li 게시판
    for a in soup.find_all("a", href=True):
        title = link_label(a)
        u = clean_url(absolute(base_url, a.get("href", "")))

        if not u or not same_domain(base_url, u):
            continue
        if not looks_like_detail_link(u, title):
            continue
        if len(title) < 2:
            continue

        parent = a.parent
        parent_text = text_of(parent)
        dt = parse_date_text(parent_text)

        key = u
        if key in seen:
            continue

        seen.add(key)
        out.append({
            "url": u,
            "title": title,
            "date": dt,
        })

        if len(out) >= 80:
            break

    return out

def recent_detail_urls(base_url, soup):
    candidates = extract_post_candidates(base_url, soup)
    if not candidates:
        return []

    dated = [x for x in candidates if x.get("date")]
    if dated:
        dated.sort(key=lambda x: x["date"], reverse=True)
        recent = [x for x in dated if is_recent_date(x["date"])]
        # 목록 날짜가 비정상적으로 파싱된 경우에도 상세페이지에서 재검증할 수 있도록
        # 일부 후보를 보존한다. 최근 날짜 후보가 있으면 그것을 우선한다.
        if recent:
            return recent[:RECENT_POSTS]
        return dated[:RECENT_POSTS]

    return candidates[:RECENT_POSTS]

def inspect_board(board_url):
    r = get(board_url)
    if not r:
        return None, [], {
            "status": "BOARD_FETCH_ERROR",
            "url": board_url,
            "candidate_posts": 0,
            "recent_posts": 0,
        }

    soup = BeautifulSoup(r.text, "html.parser")
    if looks_like_wrong_board(r.url, soup):
        return r.url, [], {
            "status": "HARD_EXCLUDED",
            "url": r.url,
            "candidate_posts": 0,
            "recent_posts": 0,
        }

    candidates = recent_detail_urls(r.url, soup)
    recent = []
    for x in candidates:
        if x.get("date") and is_recent_date(x["date"]):
            recent.append(x)
        elif not x.get("date"):
            recent.append(x)

    # 날짜가 목록에서 오래되게 잡혀도 상세페이지에서 최종 검증한다.
    # 따라서 후보가 있으면 게시판 확보 단계에서는 VERIFIED로 본다.
    recent_titles = " ".join(x.get("title", "") for x in candidates[:10])

    return r.url, candidates[:RECENT_POSTS], {
        "status": "VERIFIED" if candidates else "NO_RECENT_CANDIDATE",
        "url": r.url,
        "candidate_posts": len(candidates),
        "recent_posts": len(recent),
        "recent_titles": recent_titles[:1000],
        "score_hint": notice_board_score(r.url, text_of(soup.title)),
    }

def discover_board(home):
    diag = {
        "home": home, "status": "START", "candidate_count": 0,
        "fetched_count": 0, "verified_count": 0, "selected": "",
        "selected_score": None, "selected_status": "", "candidates": [],
        "candidate_errors": 0,
    }

    r = get(home)
    if not r:
        diag["status"] = "HOME_ERROR"
        return None, [], "HOME_ERROR", diag

    soup = BeautifulSoup(r.text, "html.parser")
    links = extract_links(r.url, soup)
    scored = sorted([(board_score(u, t), u, t) for u, t in links], key=lambda x: x[0], reverse=True)
    diag["candidate_count"] = len(scored)

    if not scored:
        diag["status"] = "NO_CANDIDATE"
        return None, [], "NO_CANDIDATE", diag

    best = None
    for score, u, t in scored[:BOARD_DISCOVERY_MAX_FETCH]:
        if score < -8:
            continue
        if looks_like_wrong_board(u, None):
            continue
        diag["fetched_count"] += 1
        try:
            bu, details, info = inspect_board(u)
        except Exception as e:
            diag["candidate_errors"] += 1
            diag["candidates"].append({"score":score,"url":u,"title":t[:120],"status":"INSPECT_ERROR","error_type":classify_exception(e) if 'classify_exception' in globals() else type(e).__name__})
            continue

        diag["candidates"].append({
            "score": score, "url": u, "title": t[:120],
            "status": info.get("status"),
            "recent_posts": info.get("recent_posts", 0),
            "candidate_posts": info.get("candidate_posts", 0),
        })
        if info.get("status") != "VERIFIED":
            continue

        diag["verified_count"] += 1
        combined = score + info.get("score_hint", 0)
        rt = info.get("recent_titles", "")
        for kw in KEYWORDS:
            if kw in rt:
                combined += 2
        # 실제 후보 수가 있는 게시판을 약간 우선
        combined += min(info.get("candidate_posts", 0), 10)

        if best is None or combined > best["score"]:
            best = {"score":combined,"url":bu,"details":details,"title":t}

    if best:
        diag["status"] = "VERIFIED"
        diag["selected"] = best["url"]
        diag["selected_score"] = best["score"]
        diag["selected_status"] = "VERIFIED"
        return best["url"], best["details"][:RECENT_POSTS], "DISCOVERED", diag

    diag["status"] = "CANDIDATE_NOT_VERIFIED"
    return None, [], "CANDIDATE_NOT_VERIFIED", diag

# -----------------------------
# V8.12.6 detail/list diagnostics
# -----------------------------
LIST_ONLY_PATHS = {
    "/list", "/lists", "/index", "/events", "/event",
    "/notice", "/notices", "/news", "/board", "/bbs"
}

def is_list_only_url(url):
    try:
        p = urlparse(url)
        path = (p.path or "").rstrip("/").lower()
        q = (p.query or "").lower()
        if path in LIST_ONLY_PATHS and not q:
            return True
        if re.search(r"(^|&)(page|pageindex|pageno|page_no|mode=list|act=list)=", q):
            return True
    except Exception:
        pass
    return False

def infer_detail_pattern(board_url, detail_urls):
    if not board_url or not detail_urls:
        return ""
    try:
        bp=urlparse(board_url); b=(bp.path or "").rstrip("/")
        for u in detail_urls:
            dp=urlparse(u); d=(dp.path or "").rstrip("/")
            if b and d.startswith(b+"/"):
                return b+"/"
    except Exception:
        pass
    return ""

def classify_exception(exc):
    if isinstance(exc, requests.exceptions.Timeout):
        return "TIMEOUT"
    if isinstance(exc, requests.exceptions.ConnectionError):
        return "CONNECTION_ERROR"
    if isinstance(exc, (KeyError, IndexError, TypeError, AttributeError)):
        return "CODE_ERROR"
    if isinstance(exc, ValueError):
        return "VALUE_ERROR"
    return "PROCESS_ERROR"

# -----------------------------
# Detail matching
# -----------------------------
def match_post_detailed(u):
    r = get(u)
    if not r:
        return None, "DETAIL_FETCH_ERROR", "FETCH_ERROR"

    soup = BeautifulSoup(r.text, "html.parser")
    if is_list_only_url(r.url):
        return None, "LIST_PAGE", ""

    if is_probable_content_page(r.url, soup, ""):
        return None, "GENERIC_PAGE", ""

    dt = extract_detail_date(soup)
    if not dt or not is_recent_date(dt):
        return None, "OLD_OR_NO_DATE", ""

    title = extract_title(soup)
    if not title:
        return None, "NO_TITLE", ""

    # 명백한 사이트 공통 페이지는 게시물로 보지 않는다.
    if title in GENERIC_PAGE_TITLES:
        return None, "GENERIC_TITLE", ""

    if any(x in title for x in EXCLUDE_TITLE):
        return None, "CONTEST_TITLE", ""

    if not has_post_structure(soup, title):
        return None, "NOT_POST_STRUCTURE", ""

    for kw in KEYWORDS:
        if kw in title:
            return {
                "url": r.url, "title": title[:300],
                "date": dt.strftime("%Y-%m-%d"),
                "keyword": kw, "match_type": "TITLE",
            }, "TITLE_MATCH", kw

    body = visible_main_text(soup)
    if body:
        for kw in KEYWORDS:
            if meaningful_body_match(body, kw):
                return {
                    "url": r.url, "title": title[:300],
                    "date": dt.strftime("%Y-%m-%d"),
                    "keyword": kw, "match_type": "BODY",
                }, "BODY_MATCH", kw

    return None, "NO_KEYWORD", ""

def match_post(u):
    m, _, _ = match_post_detailed(u)
    return m

# -----------------------------
# State / pending
# -----------------------------
def load_json(path, default):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default

def save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def normalize_pending_item(x):
    if not isinstance(x, dict):
        return None
    if not x.get("url"):
        return None

    return {
        "url": x.get("url", ""),
        "title": x.get("title", "")[:300],
        "date": x.get("date", ""),
        "keyword": x.get("keyword", ""),
        "match_type": x.get("match_type", ""),
        "institution": x.get("institution", ""),
        "added_at": x.get("added_at", datetime.now(KST).isoformat()),
    }

def pending_key(x):
    return x.get("url", "") if isinstance(x, dict) else ""

def revalidate_pending(pending):
    """
    오래된 pending은 제거.
    상세페이지를 다시 확인할 수 있으면 날짜/제목/키워드를 갱신.
    """
    valid = []
    removed = 0

    for item in pending[:MAX_PENDING]:
        item = normalize_pending_item(item)
        if not item:
            removed += 1
            continue

        u = item["url"]
        r = get(u)
        if not r:
            # 일시적 장애는 pending 유지
            valid.append(item)
            continue

        soup = BeautifulSoup(r.text, "html.parser")
        dt = extract_detail_date(soup)

        if not dt or not is_recent_date(dt):
            removed += 1
            continue

        title = extract_title(soup)
        if any(x in title for x in EXCLUDE_TITLE):
            removed += 1
            continue

        new_match = match_post(u)
        if new_match:
            new_match["institution"] = item.get("institution", "")
            new_match["added_at"] = item.get("added_at", datetime.now(KST).isoformat())
            valid.append(new_match)
        else:
            removed += 1

    return valid, removed

# -----------------------------
# Targets
# -----------------------------
def find_target_file():
    for f in TARGET_FILE_CANDIDATES:
        if os.path.exists(f):
            return f
    return None

def load_targets():
    path = find_target_file()
    if not path:
        raise FileNotFoundError(
            f"기관 목록 파일을 찾을 수 없습니다. "
            f"다음 중 하나가 필요합니다: {TARGET_FILE_CANDIDATES}"
        )

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    ws = wb["355기관"] if "355기관" in wb.sheetnames else wb[wb.sheetnames[0]]

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [norm(x) for x in rows[0]]

    def col(*names):
        for name in names:
            if name in headers:
                return headers.index(name)
        return None

    idx_name = col("기관명", "기관")
    idx_home = col("홈페이지URL", "URL", "홈페이지")
    idx_board = col("공지게시판URL", "공지사항URL", "게시판URL")

    if idx_name is None or idx_home is None:
        raise ValueError(
            f"기관명/홈페이지URL 열을 찾지 못했습니다. 현재 열: {headers}"
        )

    out = []
    for row in rows[1:]:
        if not row:
            continue

        name = norm(row[idx_name]) if idx_name < len(row) else ""
        home = norm(row[idx_home]) if idx_home < len(row) else ""
        board = norm(row[idx_board]) if idx_board is not None and idx_board < len(row) else ""

        if not name or not home:
            continue

        out.append({
            "기관명": name,
            "홈페이지URL": home,
            "공지게시판URL": board,
        })

    return out

# -----------------------------
# Institution process
# -----------------------------
def process(target, state):
    name = target["기관명"]
    home = target["홈페이지URL"]
    seed = target.get("공지게시판URL", "")
    result = {
        "기관명": name, "home": home, "board": "", "status": "",
        "posts_checked": 0, "matches": [], "error": "",
        "error_type": "", "diag": {},
        "detail_checked": 0, "recent_posts": 0,
        "title_matches": 0, "body_matches": 0,
        "excluded_list_pages": 0, "excluded_generic_pages": 0,
        "excluded_contest_titles": 0, "detail_errors": 0,
    }

    try:
        cached = state.get("boards", {}).get(home, "")
        board = cached or seed
        details = []
        recoverable_warning = ""

        if board:
            try:
                bu, details, info = inspect_board(board)
            except Exception as e:
                # 캐시/시드 검증 실패는 복구 가능한 경고다.
                recoverable_warning = classify_exception(e)
                bu, details, info = None, [], {"status":"BOARD_FETCH_ERROR"}
            if info.get("status") == "VERIFIED":
                board = bu
                result["board"] = bu
                result["status"] = "CACHED_OR_SEED"
            else:
                board = ""
                details = []

        if not board:
            try:
                bu, details, status, diag = discover_board(home)
                result["diag"] = diag or {}
                if bu and details:
                    board = bu
                    result["board"] = bu
                    result["status"] = "DISCOVERED"
                else:
                    result["status"] = status
                    return result
            except Exception as e:
                et = classify_exception(e)
                result.update({"status":"DISCOVERY_ERROR", "error_type":et, "error":str(e)[:500]})
                result["diag"] = {"status":"DISCOVERY_ERROR","error_type":et,"error":str(e)[:500]}
                return result

        unique=[]; seen_urls=set()
        for item in details:
            u = item.get("url") if isinstance(item, dict) else item
            if u and u not in seen_urls:
                seen_urls.add(u); unique.append(item)
        result["posts_checked"] = len(unique)
        result["detail_checked"] = len(unique)

        pattern = infer_detail_pattern(board, [x.get("url") if isinstance(x,dict) else x for x in unique])
        result.setdefault("diag", {})
        if pattern:
            result["detail_pattern"] = pattern
            result["diag"]["detail_pattern"] = pattern

        matches=[]
        for item in unique:
            u = item.get("url") if isinstance(item,dict) else item
            try:
                m, reason, kw = match_post_detailed(u)
                if reason == "LIST_PAGE": result["excluded_list_pages"] += 1
                elif reason in ("GENERIC_PAGE", "GENERIC_TITLE"): result["excluded_generic_pages"] += 1
                elif reason == "CONTEST_TITLE": result["excluded_contest_titles"] += 1
                elif reason == "DETAIL_FETCH_ERROR": result["detail_errors"] += 1
                elif reason == "OLD_OR_NO_DATE": pass
                elif reason == "TITLE_MATCH": result["title_matches"] += 1
                elif reason == "BODY_MATCH": result["body_matches"] += 1
                if m:
                    m["institution"] = name
                    matches.append(m)
            except Exception as e:
                result["detail_errors"] += 1
        result["matches"] = matches
        result["recent_posts"] = max(0, len(unique) - result["excluded_list_pages"] - result["excluded_generic_pages"] - result["excluded_contest_titles"])
        # recoverable warning은 fatal error_type으로 올리지 않는다.
        if recoverable_warning:
            result.setdefault("diag", {})["recoverable_board_warning"] = recoverable_warning
        return result

    except Exception as e:
        et=classify_exception(e)
        result.update({"status":"PROCESS_ERROR","error_type":et,"error":str(e)[:500]})
        result["diag"]={**(result.get("diag") or {}),"status":"PROCESS_ERROR","error_type":et,"error":str(e)[:500]}
        return result

# -----------------------------
# Telegram
# -----------------------------
def telegram_send(item):
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.environ.get("TELEGRAM_CHAT_ID", "").strip()

    if not token or not chat_id:
        return False

    title = item.get("title", "")
    institution = item.get("institution", "")
    keyword = item.get("keyword", "")
    date = item.get("date", "")
    url = item.get("url", "")

    text = (
        f"📢 [{institution}] 공공기관 참여/설문 게시물\n\n"
        f"제목: {title}\n"
        f"일자: {date}\n"
        f"키워드: {keyword}\n"
        f"매칭: {item.get('match_type','')}\n"
        f"링크: {url}"
    )

    api = f"https://api.telegram.org/bot{token}/sendMessage"

    try:
        r = SESSION.post(
            api,
            data={
                "chat_id": chat_id,
                "text": text[:4000],
                "disable_web_page_preview": False,
            },
            timeout=20,
        )
        return r.ok
    except Exception:
        return False

# -----------------------------
# Main
# -----------------------------
def main():
    started = time.time()

    targets = load_targets()
    state = load_json(STATE_FILE, {
        "seen": [],
        "boards": {},
        "updated_at": "",
        "version": VERSION,
    })

    pending = load_json(PENDING_FILE, [])
    if not isinstance(pending, list):
        pending = []

    # 기존 pending 재검증
    pending, pending_removed_invalid = revalidate_pending(pending)

    seen = set(state.get("seen", []))
    results = []

    completed = 0
    no_board = 0
    errors = 0
    posts_checked = 0
    new_matches = 0
    status_counts = {}
    error_type_counts = {}
    aggregate = {
        "board_candidates": 0, "post_candidates": 0, "detail_checked": 0,
        "recent_posts": 0, "title_matches": 0, "body_matches": 0,
        "excluded_list_pages": 0, "excluded_generic_pages": 0,
        "excluded_contest_titles": 0, "detail_errors": 0, "fatal_errors": 0,
    }

    deadline = started + MAX_TOTAL_SECONDS

    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as ex:
        future_map = {
            ex.submit(process, target, state): target
            for target in targets
        }

        for fut in as_completed(future_map):
            if time.time() >= deadline:
                break

            target = future_map[fut]

            try:
                result = fut.result()
                results.append(result)

                status = result.get("status", "") or "UNKNOWN"
                status_counts[status] = status_counts.get(status, 0) + 1
                if result.get("board"):
                    completed += 1
                    state.setdefault("boards", {})[result["home"]] = result["board"]
                    dp=(result.get("diag") or {}).get("detail_pattern", "")
                    if dp:
                        state.setdefault("detail_patterns", {})[result["home"]] = dp
                else:
                    no_board += 1

                posts_checked += result.get("posts_checked", 0)
                aggregate["detail_checked"] += result.get("detail_checked", 0)
                aggregate["recent_posts"] += result.get("recent_posts", 0)
                aggregate["title_matches"] += result.get("title_matches", 0)
                aggregate["body_matches"] += result.get("body_matches", 0)
                aggregate["excluded_list_pages"] += result.get("excluded_list_pages", 0)
                aggregate["excluded_generic_pages"] += result.get("excluded_generic_pages", 0)
                aggregate["excluded_contest_titles"] += result.get("excluded_contest_titles", 0)
                aggregate["detail_errors"] += result.get("detail_errors", 0)
                d = result.get("diag") or {}
                aggregate["board_candidates"] += d.get("candidate_count", 0) or 0
                aggregate["post_candidates"] += result.get("posts_checked", 0)
                et=result.get("error_type", "")
                if et:
                    error_type_counts[et]=error_type_counts.get(et,0)+1
                    errors += 1

                for m in result.get("matches", []):
                    u = m.get("url", "")
                    if not u:
                        continue

                    # 이미 seen 또는 pending이면 중복 방지
                    if u in seen:
                        continue
                    if any(pending_key(x) == u for x in pending):
                        continue

                    m["added_at"] = datetime.now(KST).isoformat()
                    pending.append(m)
                    new_matches += 1

            except Exception as e:
                errors += 1
                et=classify_exception(e)
                error_type_counts[et]=error_type_counts.get(et,0)+1
                status_counts["FUTURE_EXCEPTION"]=status_counts.get("FUTURE_EXCEPTION",0)+1

    # pending 중복 제거
    dedup = {}
    for item in pending:
        k = pending_key(item)
        if k:
            dedup[k] = item
    pending = list(dedup.values())

    # FIFO
    pending.sort(key=lambda x: x.get("added_at", ""))

    # 최대 보관
    if len(pending) > MAX_PENDING:
        pending = pending[-MAX_PENDING:]

    pending_before_send = len(pending)

    # Telegram 하루 최대 20건
    sent = 0
    remaining = []

    for item in pending:
        if sent >= TELEGRAM_MAX_SEND:
            remaining.append(item)
            continue

        if telegram_send(item):
            sent += 1
            seen.add(item.get("url", ""))
        else:
            remaining.append(item)

    pending = remaining

    state["seen"] = list(seen)[-50000:]
    state["updated_at"] = datetime.now(KST).isoformat()
    state["version"] = VERSION

    save_json(STATE_FILE, state)
    save_json(PENDING_FILE, pending)

    elapsed = time.time() - started
    timed_out = elapsed >= MAX_TOTAL_SECONDS

    # 진단 파일
    diag_counts = {}
    for r in results:
        d = r.get("diag") or {}
        st = d.get("status")
        if st:
            diag_counts[st] = diag_counts.get(st, 0) + 1

    diagnostics = {
        "version": VERSION,
        "updated_at": datetime.now(KST).isoformat(),
        "targets": len(targets),
        "completed": completed,
        "no_board": no_board,
        "posts_checked": posts_checked,
        "new_matches": new_matches,
        "pending_before": pending_before_send,
        "pending_removed_invalid": pending_removed_invalid,
        "telegram_sent": sent,
        "pending_after": len(pending),
        "errors": errors,
        "elapsed_seconds": round(elapsed, 1),
        "timed_out": timed_out,
        "recent_days": RECENT_DAYS,
        "telegram_max_send": TELEGRAM_MAX_SEND,
        "board_discovery_status": diag_counts,
        "status_counts": status_counts,
        "error_type_counts": error_type_counts,
        "pipeline_counts": aggregate,
        "processed_total": len(results),
        "unprocessed_total": max(0, len(targets) - len(results)),
        "board_details": [
            {
                "기관명": r.get("기관명"),
                "status": r.get("status"),
                "board": r.get("board"),
                "posts_checked": r.get("posts_checked"),
                "detail_checked": r.get("detail_checked"),
                "recent_posts": r.get("recent_posts"),
                "title_matches": r.get("title_matches"),
                "body_matches": r.get("body_matches"),
                "excluded_list_pages": r.get("excluded_list_pages"),
                "excluded_generic_pages": r.get("excluded_generic_pages"),
                "excluded_contest_titles": r.get("excluded_contest_titles"),
                "detail_errors": r.get("detail_errors"),
                "diag_status": (r.get("diag") or {}).get("status"),
                "candidate_count": (r.get("diag") or {}).get("candidate_count"),
                "fetched_count": (r.get("diag") or {}).get("fetched_count"),
                "verified_count": (r.get("diag") or {}).get("verified_count"),
                "selected_score": (r.get("diag") or {}).get("selected_score"),
                "detail_pattern": (r.get("diag") or {}).get("detail_pattern", ""),
                "error_type": r.get("error_type", ""),
                "error": r.get("error", "")[:500],
            }
            for r in results
        ],
    }

    save_json("diagnostics.json", diagnostics)

    print(json.dumps(diagnostics, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
