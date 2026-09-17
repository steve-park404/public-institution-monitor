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
# V6.1 안정형
# 미발견 기관만 대상으로 게시판/게시물 구조 탐색
# ============================================================

INPUT_FILE = "boards_missing.xlsx"
OUTPUT_FILE = "boards_v6.xlsx"

# 동시 접속 기관 수
CONCURRENCY = 6

# 개별 HTTP 요청 제한시간
TIMEOUT = 15

# 기관 하나에 허용되는 최대 분석시간
INSTITUTION_TIMEOUT = 90

# 기관당 최대 탐색 페이지
MAX_PAGES = 20

# 페이지당 최대 링크
MAX_LINKS_PER_PAGE = 100

# 게시판 후보 실제 검증 개수
MAX_BOARD_VERIFY = 5

# 게시물 후보 역추적 개수
MAX_POST_VERIFY = 5


# ============================================================
# User-Agent
# ============================================================

USER_AGENTS = [
    (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    (
        "Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    (
        "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
        "AppleWebKit/605.1.15 (KHTML, like Gecko) "
        "Version/17.0 Mobile/15E148 Safari/604.1"
    )
]


# ============================================================
# 게시판 관련 키워드
# ============================================================

BOARD_WORDS = [
    "공지사항",
    "공지",
    "알림",
    "소식",
    "새소식",
    "게시판",
    "공지·공고",
    "공지/공고",
    "공고",
    "채용",
    "입찰",
    "보도자료",
    "자료실",
    "뉴스",
    "정보마당",
    "고시",
    "공시"
]


NOTICE_WORDS = [
    "공지사항",
    "공지",
    "알림",
    "새소식",
    "소식",
    "공지·공고",
    "공지/공고"
]


BOARD_URL_WORDS = [
    "board",
    "bbs",
    "notice",
    "news",
    "announce",
    "bulletin",
    "community",
    "boardlist",
    "boardview",
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


# ============================================================
# iframe / JS
# ============================================================

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


# ============================================================
# 허용 Content-Type
# ============================================================

CONTENT_TYPES = [
    "text/html",
    "application/xhtml+xml",
    "application/xml",
    "text/xml"
]


# ============================================================
# SSL
# ============================================================

SSL_CONTEXT = ssl.create_default_context()
SSL_CONTEXT.check_hostname = False
SSL_CONTEXT.verify_mode = ssl.CERT_NONE


# ============================================================
# URL 함수
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
        base_domain = (
            urlparse(base)
            .netloc
            .lower()
            .replace("www.", "")
        )

        target_domain = (
            urlparse(target)
            .netloc
            .lower()
            .replace("www.", "")
        )

        return base_domain == target_domain

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

        new_query = urlencode(
            query,
            doseq=True
        )

        return urlunparse(
            (
                p.scheme,
                p.netloc,
                p.path,
                p.params,
                new_query,
                ""
            )
        )

    except Exception:
        return url


# ============================================================
# HTTP 요청
# ============================================================

async def fetch(session, url, retries=0):

    url = normalize_url(url)

    last_error = ""

    for attempt in range(retries + 1):

        headers = {
            "User-Agent": USER_AGENTS[
                attempt % len(USER_AGENTS)
            ],
            "Accept": (
                "text/html,"
                "application/xhtml+xml,"
                "application/xml,"
                "text/xml,"
                "*/*;q=0.8"
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
            ) as response:

                content_type = (
                    response.headers
                    .get("Content-Type", "")
                    .lower()
                )

                content_length = (
                    response.headers
                    .get("Content-Length")
                )

                if content_length:

                    try:
                        if int(content_length) > 5000000:
                            return {
                                "ok": False,
                                "status": response.status,
                                "url": str(response.url),
                                "content_type": content_type,
                                "text": "",
                                "error": "too_large"
                            }

                    except Exception:
                        pass

                text = await response.text(
                    encoding=None,
                    errors="ignore"
                )

                return {
                    "ok": response.status < 400,
                    "status": response.status,
                    "url": str(response.url),
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
# BeautifulSoup
# XML이면 XML parser 사용
# ============================================================

def make_soup(html):

    if not html:
        return None

    try:

        start = html.lstrip()[:500].lower()

        if (
            start.startswith("<?xml")
            or "<rss" in start
            or "<feed" in start
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


# ============================================================
# 링크 추출
# ============================================================

def extract_links(html, base_url):

    soup = make_soup(html)

    if soup is None:
        return []

    links = []
    seen = set()

    for a in soup.find_all("a", href=True):

        href = a.get(
            "href",
            ""
        ).strip()

        if not href:
            continue

        if href.startswith(
            (
                "javascript:",
                "mailto:",
                "tel:",
                "#"
            )
        ):
            continue

        try:

            absolute = clean_url(
                urljoin(
                    base_url,
                    href
                )
            )

            if not absolute.startswith(
                (
                    "http://",
                    "https://"
                )
            ):
                continue

            if absolute in seen:
                continue

            seen.add(absolute)

            links.append(
                {
                    "url": absolute,
                    "text": a.get_text(
                        " ",
                        strip=True
                    )
                }
            )

        except Exception:
            continue

    return links[:MAX_LINKS_PER_PAGE]


# ============================================================
# iframe 추출
# ============================================================

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


def valid_iframe(url):

    low = url.lower()

    for word in BAD_IFRAME_WORDS:

        if word in low:
            return False

    return True


# ============================================================
# CMS 확인
# ============================================================

def detect_cms(html):

    low = html.lower()

    result = []

    if "k2web" in low:
        result.append("K2Web")

    if (
        "joomla" in low
        or "/components/com_" in low
    ):
        result.append("Joomla")

    if "wordpress" in low:
        result.append("WordPress")

    if (
        "drupal" in low
        or "drupalsettings" in low
    ):
        result.append("Drupal")

    if (
        "gnu board" in low
        or "gnuboard" in low
        or "g5_" in low
    ):
        result.append("그누보드")

    return result


# ============================================================
# JS / API 분석
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

            for match in matches:

                if isinstance(
                    match,
                    tuple
                ):
                    match = match[0]

                match = str(match).strip()

                if (
                    match
                    and len(match) < 500
                    and match not in endpoints
                ):
                    endpoints.append(match)

        except Exception:
            pass

    return (
        score,
        findings,
        endpoints[:20]
    )


# ============================================================
# 게시판 점수
# ============================================================

def board_score(
    url,
    text,
    html=""
):

    target = (
        str(url)
        + " "
        + str(text)
    )

    score = 0
    reasons = []

    target_low = target.lower()
    text_low = str(text).lower()

    for word in BOARD_URL_WORDS:

        if word.lower() in target_low:

            score += 2

            reasons.append(
                "URL:" + word
            )

    for word in BOARD_WORDS:

        if word.lower() in text_low:

            score += 2

            reasons.append(
                "TEXT:" + word
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

            pagination_count = 0

            for word in [
                "다음",
                "이전",
                "1",
                "2",
                "3"
            ]:

                if word in body_text:
                    pagination_count += 1

            if pagination_count >= 3:

                score += 2
                reasons.append(
                    "페이지네이션"
                )

    return score, reasons


# ============================================================
# 게시물 URL 판별
# ============================================================

def looks_like_post(
    url,
    text=""
):

    low = (
        str(url)
        + " "
        + str(text)
    ).lower()

    parsed = urlparse(url)

    query = parsed.query.lower()

    for param in POST_PARAM_WORDS:

        if param.lower() + "=" in query:
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


# ============================================================
# 게시물 → 상위 게시판 URL 추정
# ============================================================

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
                urlunparse(
                    (
                        p.scheme,
                        p.netloc,
                        parent + "/",
                        "",
                        "",
                        ""
                    )
                )
            )

        if p.query:

            result.append(
                urlunparse(
                    (
                        p.scheme,
                        p.netloc,
                        p.path,
                        "",
                        "",
                        ""
                    )
                )
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
                    urlunparse(
                        (
                            p.scheme,
                            p.netloc,
                            candidate,
                            "",
                            "",
                            ""
                        )
                    )
                )

    except Exception:
        pass

    return list(
        dict.fromkeys(result)
    )


# ============================================================
# 기관 기본 결과
# ============================================================

def create_result(
    name,
    homepage
):

    return {
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


# ============================================================
# 기관 분석 내부
# ============================================================

async def analyze_institution_inner(
    session,
    result
):

    homepage = result["홈페이지"]

    # --------------------------------------------------------
    # 1. 홈페이지
    # --------------------------------------------------------

    first = await fetch(
        session,
        homepage,
        retries=1
    )

    if not first["ok"]:

        result["접속상태"] = (
            "실패:"
            + str(
                first["error"]
                or first["status"]
            )
        )

        result["실패단계"] = (
            "홈페이지 접속"
        )

        return result

    result["접속상태"] = (
        "정상:"
        + str(first["status"])
    )

    homepage_final = first["url"]

    # --------------------------------------------------------
    # BFS
    # --------------------------------------------------------

    visited = set()

    queue = deque()

    queue.append(
        homepage_final
    )

    candidate_boards = []
    candidate_posts = []
    candidate_iframes = []
    candidate_api = []

    cms_findings = []

    js_score = 0

    pages = 0

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
            item in content_type
            for item in CONTENT_TYPES
        ):
            continue

        html = data["text"]

        if not html:
            continue

        pages += 1

        result["탐색페이지수"] = pages

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

        js_result = analyze_js(
            html
        )

        score = js_result[0]
        endpoints = js_result[2]

        js_score += score

        for endpoint in endpoints:

            if endpoint in candidate_api:
                continue

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

                exists = any(
                    x["url"] == link
                    for x in candidate_boards
                )

                if not exists:

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
            # 우선 탐색
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

                queue.append(
                    link
                )

        # ----------------------------------------------------
        # 초기 메뉴 탐색
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
                    word.lower() in target
                    for word in BOARD_WORDS
                ):

                    if link not in visited:

                        queue.append(
                            link
                        )

        # ----------------------------------------------------
        # queue 폭주 방지
        # ----------------------------------------------------

        max_queue = MAX_PAGES * 3

        if len(queue) > max_queue:

            queue = deque(
                list(queue)[:max_queue]
            )

    # --------------------------------------------------------
    # 후보 정리
    # --------------------------------------------------------

    candidate_boards = sorted(
        candidate_boards,
        key=lambda x: x["score"],
        reverse=True
    )[:10]

    candidate_posts = list(
        dict.fromkeys(
            candidate_posts
        )
    )[:10]

    # --------------------------------------------------------
    # 게시판 실제 검증
    # --------------------------------------------------------

    verified_boards = []

    for candidate in candidate_boards[
        :MAX_BOARD_VERIFY
    ]:

        check = await fetch(
            session,
            candidate["url"],
            retries=0
        )

        if not check["ok"]:
            continue

        score, reasons = board_score(
            candidate["url"],
            candidate["text"],
            check["text"]
        )

        if score >= 6:

            verified_boards.append(
                {
                    "url": candidate["url"],
                    "score": score,
                    "reasons": reasons
                }
            )

    # --------------------------------------------------------
    # 게시물에서 게시판 역추적
    # --------------------------------------------------------

    for post in candidate_posts[
        :MAX_POST_VERIFY
    ]:

        parents = parent_candidates(
            post
        )

        for parent in parents[:3]:

            check = await fetch(
                session,
                parent,
                retries=0
            )

            if not check["ok"]:
                continue

            score, reasons = board_score(
                parent,
                "",
                check["text"]
            )

            if score >= 6:

                verified_boards.append(
                    {
                       
