import asyncio
import aiohttp
import pandas as pd
import re
import ssl
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse
from collections import deque
from pathlib import Path

# ============================================================
# V6.1 - 미발견 기관 안정형 심층 탐색
# ============================================================

INPUT_FILE = "boards_missing.xlsx"
OUTPUT_FILE = "boards_v6.xlsx"

# 안정성 우선
CONCURRENCY = 6
TIMEOUT = 15
INSTITUTION_TIMEOUT = 90

# 기관당 탐색 페이지 수
MAX_PAGES = 20

# 페이지당 링크 최대
MAX_LINKS_PER_PAGE = 100

# 후보 검증 개수
MAX_BOARD_VERIFY = 5
MAX_POST_VERIFY = 5

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/140.0.0.0 Safari/537.36",

    "Mozilla/5.0 (X11; Linux x86_64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/140.0.0.0 Safari/537.36",

    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) "
    "Version/17.0 Mobile/15E148 Safari/604.1"
]

BOARD_WORDS = [
    "공지사항", "공지", "알림", "소식", "새소식",
    "게시판", "공지·공고", "공지/공고",
    "공고", "채용", "입찰", "보도자료", "자료실",
    "뉴스", "정보마당", "고시", "공시"
]

NOTICE_WORDS = [
    "공지사항", "공지", "알림", "새소식",
    "소식", "공지·공고", "공지/공고"
]

BAD_IFRAME_WORDS = [
    "google.com/maps",
    "youtube.com/embed",
    "googletagmanager.com",
    "google-analytics.com",
    "doubleclick.net",
    "facebook.com/plugins",
    "instagram.com"
]

API_WORDS = [
    "fetch(",
    "axios",
    "$.ajax",
    "$.get",
    "$.post",
    "XMLHttpRequest",
    "ajax/",
    "api/",
    "/api"
]

BOARD_URL_WORDS = [
    "board",
    "bbs",
    "notice",
    "news",
    "announce",
    "bulletin",
    "community",
    "boardList",
    "boardView",
    "board.do",
    "bbs.do",
    "notice.do",
    "list.do",
    "view.do"
]

POST_PARAM_WORDS = [
    "nttId",
    "ntt_id",
    "articleId",
    "article_id",
    "seq",
    "idx",
    "no",
    "bbsSeq",
    "boardSeq"
]

CONTENT_TYPES = [
    "text/html",
    "application/xhtml+xml",
    "application/xml",
    "text/xml"
]

SSL_CONTEXT = ssl.create_default_context()
SSL_CONTEXT.check_hostname = False
SSL_CONTEXT.verify_mode = ssl.CERT_NONE


# ============================================================
# 기본 함수
# ============================================================

def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def same_domain(base, target):
    try:
        b = urlparse(base).netloc.lower().replace("www.", "")
        t = urlparse(target).netloc.lower().replace("www.", "")
        return b == t
    except Exception:
        return False


def clean_url(url):
    try:
        p = urlparse(url)

        query = parse_qs(
            p.query,
            keep_blank_values=True
        )

        remove = {
            "utm_source",
            "utm_medium",
            "utm_campaign",
            "utm_term",
            "utm_content"
        }

        query = {
            k: v
            for k, v in query.items()
            if k not in remove
        }

        new_query = urlencode(query, doseq=True)

        return urlunparse((
            p.scheme,
            p.netloc,
            p.path,
            p.params,
            new_query,
            ""
        ))

    except Exception:
        return url


# ============================================================
# HTTP
# ============================================================

async def fetch(session, url, retries=1):

    url = normalize_url(url)

    last_error = ""

    for attempt in range(retries + 1):

        headers = {
            "User-Agent": USER_AGENTS[
                attempt % len(USER_AGENTS)
            ],
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml,text/xml;q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
            "Connection": "close"
        }

        try:

            timeout = aiohttp.ClientTimeout(
                total=TIMEOUT,
                connect=8,
                sock_read=TIMEOUT
            )

            async with session.get(
                url,
                headers=headers,
                timeout=timeout,
                ssl=SSL_CONTEXT,
                allow_redirects=True,
                max_redirects=5
            ) as r:

                content_type = (
                    r.headers.get(
                        "Content-Type",
                        ""
                    ).lower()
                )

                # 너무 큰 파일 방지
                content_length = r.headers.get(
                    "Content-Length"
                )

                if content_length:

                    try:
                        if int(content_length) > 5_000_000:
                            return {
                                "ok": False,
                                "status": r.status,
                                "url": str(r.url),
                                "content_type": content_type,
                                "text": "",
                                "error": "too_large"
                            }
                    except Exception:
                        pass

                text = await r.text(
                    encoding=None,
                    errors="ignore"
                )

                return {
                    "ok": r.status < 400,
                    "status": r.status,
                    "url": str(r.url),
                    "content_type": content_type,
                    "text": text,
                    "error": ""
                }

        except asyncio.TimeoutError:

            last_error = "Timeout"

        except aiohttp.ClientError as e:

            last_error = type(e).__name__

        except Exception as e:

            last_error = type(e).__name__

        if attempt < retries:
            await asyncio.sleep(0.5)

    return {
        "ok": False,
        "status": 0,
        "url": url,
        "content_type": "",
        "text": "",
        "error": last_error
    }


# ============================================================
# HTML 분석
# ============================================================

def make_soup(html):

    if not html:
        return None

    try:
        # XML이면 XML parser 사용
        stripped = html.lstrip()

        if (
            stripped.startswith("<?xml")
            or "<rss" in stripped[:500].lower()
            or "<feed" in stripped[:500].lower()
        ):
            return BeautifulSoup(
                html,
                "xml"
            )

        return BeautifulSoup(
            html,
            "html.parser"
        )

    except Exception:

        try:
            return BeautifulSoup(
                html,
                "html.parser"
            )
        except Exception:
            return None


def extract_links(html, base_url):

    soup = make_soup(html)

    if soup is None:
        return []

    links = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith((
            "javascript:",
            "mailto:",
            "tel:",
            "#"
        )):
            continue

        try:

            absolute = clean_url(
                urljoin(
                    base_url,
                    href
                )
            )

            if absolute.startswith((
                "http://",
                "https://"
            )):

                links.append({
                    "url": absolute,
                    "text": a.get_text(
                        " ",
                        strip=True
                    )
                })

        except Exception:
            pass

    # 중복 제거
    unique = []
    seen = set()

    for item in links:

        if item["url"] not in seen:

            seen.add(item["url"])
            unique.append(item)

    return unique[:MAX_LINKS_PER_PAGE]


def extract_iframes(html, base_url):

    soup = make_soup(html)

    if soup is None:
        return []

    result = []

    for iframe in soup.find_all("iframe"):

        src = (
            iframe.get("src")
            or iframe.get("data-src")
            or iframe.get("data-url")
            or ""
        ).strip()

        if not src:
            continue

        try:
            result.append(
                urljoin(
                    base_url,
                    src
                )
            )
        except Exception:
            pass

    return result


# ============================================================
# JS / CMS
# ============================================================

def analyze_js(html):

    low = html.lower()

    score = 0
    findings = []
    endpoints = []

    for word in API_WORDS:

        if word.lower() in low:

            score += 1
            findings.append(word)

    patterns = [

        r'["\']([^"\']{0,200}'
        r'(?:api|ajax|board|bbs|notice)'
        r'[^"\']{0,200})["\']',

        r'url\s*:\s*["\']([^"\']+)["\']',

        r'fetch\s*\(\s*["\']([^"\']+)["\']',

        r'axios\.(?:get|post)'
        r'\s*\(\s*["\']([^"\']+)["\']'
    ]

    for pattern in patterns:

        try:

            matches = re.findall(
                pattern,
                html,
                re.I
            )

            for x in matches:

                if isinstance(x, tuple):
                    x = x[0]

                x = str(x).strip()

                if (
                    x
                    and len(x) < 500
                    and x not in endpoints
                ):

                    endpoints.append(x)

        except Exception:
            pass

    return (
        score,
        findings,
        endpoints[:20]
    )


def detect_cms(html):

    low = html.lower()

    cms = []

    if "k2web" in low:
        cms.append("K2Web")

    if "joomla" in low or "/components/com_" in low:
        cms.append("Joomla")

    if "wordpress" in low:
        cms.append("WordPress")

    if "drupal" in low or "drupalsettings" in low:
        cms.append("Drupal")

    if (
        "gnu board" in low
        or "gnuboard" in low
        or "g5_" in low
    ):
        cms.append("그누보드")

    return cms


# ============================================================
# 게시판 판별
# ============================================================

def board_score(url, text, html=""):

    target = (
        str(url)
        + " "
        + str(text)
    )

    score = 0
    reasons = []

    for word in BOARD_URL_WORDS:

        if word.lower() in target.lower():

            score += 2
            reasons.append(
                f"URL:{word}"
            )

    for word in BOARD_WORDS:

        if word.lower() in str(text).lower():

            score += 2
            reasons.append(
                f"TEXT:{word}"
            )

    if html:

        soup = make_soup(html)

        if soup:

            body_text = soup.get_text(
                " ",
                strip=True
            )

            date_count = len(
                re.findall(
                    r"\b20\d{2}[./-]"
                    r"\d{1,2}[./-]"
                    r"\d{1,2}\b",
                    body_text
                )
            )

            if date_count >= 2:

                score += 4
                reasons.append(
                    "날짜복수"
                )

            pagination_words = [
                "다음",
                "이전",
                "1",
                "2",
                "3"
            ]

            page_count = sum(
                1
                for x in pagination_words
                if x in body_text
            )

            if page_count >= 3:

                score += 2
                reasons.append(
                    "페이지네이션"
                )

    return score, reasons


def looks_like_post(url, text=""):

    low = (
        str(url)
        + " "
        + str(text)
    ).lower()

    parsed = urlparse(url)
    query = parsed.query.lower()

    for p in POST_PARAM_WORDS:

        if p.lower() + "=" in query:
            return True

    patterns = [
        r"/view/",
        r"/view\?",
        r"/read/",
        r"/read\?",
        r"/article/",
        r"/article\?",
        r"/detail/",
        r"/detail\?",
        r"/board/view",
        r"/bbs/view",
        r"/notice/view",
        r"/board/read"
    ]

    for pattern in patterns:

        if re.search(
            pattern,
            low
        ):
            return True

    return False


def parent_candidates(url):

    result = []

    try:

        p = urlparse(url)

        path = p.path.rstrip("/")

        parts = path.split("/")

        if len(parts) > 1:

            parent = "/".join(
                parts[:-1]
            )

            result.append(
                urlunparse((
                    p.scheme,
                    p.netloc,
                    parent + "/",
                    "",
                    "",
                    ""
                ))
            )

        if p.query:

            result.append(
                urlunparse((
                    p.scheme,
                    p.netloc,
                    p.path,
                    "",
                    "",
                    ""
                ))
            )

        for token in [
            "/view",
            "/read",
            "/detail",
            "/article"
        ]:

            if token in path.lower():

                candidate = re.sub(
                    token,
                    "/list",
                    path,
                    flags=re.I
                )

                result.append(
                    urlunparse((
                        p.scheme,
                        p.netloc,
                        candidate,
                        "",
                        "",
                        ""
                    ))
                )

    except Exception:
        pass

    return list(
        dict.fromkeys(result)
    )


def valid_iframe(url):

    low = url.lower()

    for bad in BAD_IFRAME_WORDS:

        if bad in low:
            return False

    return True


# ============================================================
# 기관 1개 분석
# ============================================================

async def analyze_institution(
    session,
    semaphore,
    row
):

    name = str(
        row.get(
            "기관명",
            ""
        )
    )

    homepage = normalize_url(
        row.get(
            "URL",
            ""
        )
    )

    result = {

        "기관명": name,
        "홈페이지": homepage,
        "접속상태": "",
        "탐색페이지수": 0,
        "CMS": "",
        "iframe": "",
        "JS": "",
        "후보수": 0,
        "게시물후보수": 0,
        "게시판URL": "",
        "게시물URL": "",
        "iframeURL": "",
        "API후보": "",
        "탐색방법": "",
        "신뢰도": "",
        "실패단계": "",
        "오류": ""
    }

    async with semaphore:

        try:

            # ------------------------------------------------
            # 전체 기관 제한시간
            # ------------------------------------------------

            return await asyncio.wait_for(
                analyze_institution_inner(
                    session,
                    result
                ),
                timeout=INSTITUTION_TIMEOUT
            )

        except asyncio.TimeoutError:

            result["실패단계"] = "기관전체 Timeout"
            result["오류"] = (
                f"{INSTITUTION_TIMEOUT}초 초과"
            )

            return result

        except Exception as e:

            result["실패단계"] = "전체분석"
            result["오류"] = (
                f"{type(e).__name__}: {str(e)[:200]}"
            )

            return result


async def analyze_institution_inner(
    session,
    result
):

    homepage = result["홈페이지"]

    # --------------------------------------------------------
    # 1. 홈페이지 접속
    # --------------------------------------------------------

    first = await fetch(
        session,
        homepage,
        retries=1
    )

    if not first["ok"]:

        result["접속상태"] = (
            f"실패:{first['error'] or first['status']}"
        )

        result["실패단계"] = "홈페이지 접속"

        return result

    result["접속상태"] = (
        f"정상:{first['status']}"
    )

    # 실제 리다이렉트된 주소
    homepage_final = first["url"]

    visited = set()

    queue = deque([
        homepage_final
    ])

    candidate_boards = []
    candidate_posts = []
    candidate_iframes = []
    candidate_api = []

    js_score = 0
    cms_findings = []

    pages = 0

    # --------------------------------------------------------
    # 2. BFS 탐색
    # --------------------------------------------------------

    while queue and pages < MAX_PAGES:

        current = clean_url(
            queue.popleft()
        )

        if current in visited:
            continue

        if not same_domain(
            homepage_final,
            current
        ):
            continue

        visited.add(current)

        data = await fetch(
            session,
            current,
            retries=0
        )

        if not data["ok"]:
            continue

        content_type = data[
            "content_type"
        ]

        if not any(
            x in content_type
            for x in CONTENT_TYPES
        ):
            continue

        html = data["text"]

        if not html:
            continue

        pages += 1

        result["탐색페이지수"] = pages

        low = html.lower()

        # ----------------------------------------------------
        # CMS
        # ----------------------------------------------------

        cms_findings.extend(
            detect_cms(html)
        )

        # ----------------------------------------------------
        # iframe
        # ----------------------------------------------------

        iframe_urls = extract_iframes(
            html,
            current
        )

        for iframe_url in iframe_urls:

            if not valid_iframe(
                iframe_url
            ):
                continue

            if iframe_url not in candidate_iframes:

                candidate_iframes.append(
                    iframe_url
                )

                # iframe은 소수만 실제 탐색
                if len(candidate_iframes) <= 3:

                    if same_domain(
                        homepage_final,
                        iframe_url
                    ):

                        queue.append(
                            iframe_url
                        )

        # ----------------------------------------------------
        # JS / API
        # ----------------------------------------------------

        score, findings, endpoints = analyze_js(
            html
        )

        js_score += score

        for endpoint in endpoints:

            if endpoint not in candidate_api:

                candidate_api.append(
                    endpoint
                )

                if endpoint.startswith("/"):

                    api_url = urljoin(
                        current,
                        endpoint
                    )

                    if same_domain(
                        homepage_final,
                        api_url
                    ):

                        queue.append(
                            api_url
                        )

        # ----------------------------------------------------
        # 링크
        # ----------------------------------------------------

        links = extract_links(
            html,
            current
        )

        for item in links:

            link = item["url"]
            text = item["text"]

            if not same_domain(
                homepage_final,
                link
            ):
                continue

            # ----------------------------------------------
            # 게시판 후보
            # ----------------------------------------------

            score, reasons = board_score(
                link,
                text,
                html
            )

            if score >= 5:

                candidate = {
                    "url": link,
                    "text": text,
                    "score": score,
                    "reasons": ",".join(
                        reasons[:10]
                    )
                }

                if not any(
                    x["url"] == link
                    for x in candidate_boards
                ):

                    candidate_boards.append(
                        candidate
                    )

            # ----------------------------------------------
            # 게시물 후보
            # ----------------------------------------------

            if looks_like_post(
                link,
                text
            ):

                if link not in candidate_posts:

                    candidate_posts.append(
                        link
                    )

                    for parent in parent_candidates(
                        link
                    ):

                        if same_domain(
                            homepage_final,
                            parent
                        ):

                            queue.append(
                                parent
                            )

            # ----------------------------------------------
            # 우선 탐색 링크
            # ----------------------------------------------

            link_text = (
                text
                + " "
                + link
            ).lower()

            priority = 0

            for word in NOTICE_WORDS:

                if word.lower() in link_text:

                    priority += 5

            for word in BOARD_URL_WORDS:

                if word.lower() in link.lower():

                    priority += 2

            if (
                "k2web" in link.lower()
                or "fnctid" in link.lower()
                or "nttid" in link.lower()
            ):

                priority += 5

            if (
                priority >= 3
                and link not in visited
            ):

                queue.append(link)

        # ----------------------------------------------------
        # 초반 메뉴 추가 탐색
        # ----------------------------------------------------

        if pages <= 5:

            for item in links:

                link = item["url"]
                text = item["text"]

                if not same_domain(
                    homepage_final,
                    link
                ):
                    continue

                target = (
                    text
                    + " "
                    + link
                ).lower()

                if any(
                    x.lower() in target
                    for x in BOARD_WORDS
                ):

                    if
