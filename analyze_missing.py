import asyncio
import aiohttp
import pandas as pd
import re
import os
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from collections import Counter
from openpyxl import Workbook


# =========================================================
# 설정
# =========================================================

INPUT_FILE = "boards_missing.xlsx"
OUTPUT_FILE = "missing_structure_analysis.xlsx"

CONCURRENCY = 10
TIMEOUT = 20

MAX_PAGES = 20

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    )
}


# =========================================================
# 키워드
# =========================================================

BOARD_WORDS = [
    "공지사항",
    "공지",
    "알림마당",
    "알림",
    "소식",
    "게시판",
    "자료실",
    "자료",
    "참여",
    "국민참여",
    "시민참여",
    "고객참여",
    "소통",
    "커뮤니티",
    "뉴스",
    "보도자료",
    "공모",
    "공모전",
    "설문",
    "이벤트",
    "새소식",
    "입찰",
    "채용",
    "고시공고",
    "공지·공고",
    "정보마당",
    "경영공시",
    "전자민원",
]

POST_WORDS = [
    "view",
    "detail",
    "read",
    "article",
    "board",
    "bbs",
    "ntt",
    "seq",
    "idx",
    "no=",
    "articleid",
    "boardid",
]

ERROR_WORDS = [
    "페이지를 찾을 수 없습니다",
    "페이지가 없습니다",
    "존재하지 않는 페이지",
    "요청하신 페이지",
    "404",
    "오류가 발생",
    "에러가 발생",
    "error",
    "not found",
    "접근할 수 없습니다",
    "서비스 이용에 불편",
]

CMS_PATTERNS = {
    "K2Web": [
        "k2web",
        "k2webwizard",
        "ntt",
        "siteId",
        "fnctId",
    ],
    "Drupal": [
        "drupal",
        "/node/",
        "views-row",
    ],
    "WordPress": [
        "wp-content",
        "wp-includes",
        "wordpress",
    ],
    "Joomla": [
        "joomla",
        "/component/",
    ],
    "그누보드": [
        "gnuboard",
        "bo_table",
        "wr_id",
    ],
    "아파치/일반": [
        "apache",
    ],
}


# =========================================================
# 유틸
# =========================================================

def clean_text(text):
    if not text:
        return ""

    text = re.sub(r"\s+", " ", text)
    return text.strip()


def same_domain(url1, url2):
    try:
        return urlparse(url1).netloc == urlparse(url2).netloc
    except Exception:
        return False


def is_html_response(content_type):
    if not content_type:
        return True

    content_type = content_type.lower()

    return (
        "text/html" in content_type
        or "application/xhtml" in content_type
    )


def detect_cms(html, url):
    text = (html + " " + url).lower()

    detected = []

    for cms, patterns in CMS_PATTERNS.items():
        for pattern in patterns:
            if pattern.lower() in text:
                detected.append(cms)
                break

    if not detected:
        return "일반 HTML/미확인"

    return ", ".join(dict.fromkeys(detected))


def detect_iframe(soup):
    iframe_count = len(soup.find_all("iframe"))

    if iframe_count > 0:
        return f"있음({iframe_count}개)"

    return "없음"


def detect_js_dependency(soup, html):
    script_count = len(soup.find_all("script"))

    js_patterns = [
        "javascript:",
        "onclick=",
        "fetch(",
        "axios",
        "ajax",
        "$.ajax",
        "$.get",
        "$.post",
        "XMLHttpRequest",
        "addEventListener",
    ]

    found = []

    lower_html = html.lower()

    for pattern in js_patterns:
        if pattern.lower() in lower_html:
            found.append(pattern)

    if script_count >= 20 or len(found) >= 2:
        return "높음"

    if script_count >= 8 or len(found) >= 1:
        return "중간"

    return "낮음"


def detect_error_page(text, status):
    lower = text.lower()

    if status >= 400:
        return True

    for word in ERROR_WORDS:
        if word.lower() in lower:
            return True

    return False


def classify_url(url, text):
    """
    URL이 게시판/게시물 형태인지 대략 분류
    """

    lower_url = url.lower()
    lower_text = text.lower()

    board_score = 0
    post_score = 0

    for word in BOARD_WORDS:
        if word.lower() in lower_text:
            board_score += 1

    for word in POST_WORDS:
        if word.lower() in lower_url:
            post_score += 1

    # URL 패턴
    url_patterns = [
        r"/board",
        r"/bbs",
        r"/notice",
        r"/community",
        r"/news",
        r"/data",
        r"/archive",
        r"/particip",
        r"/event",
        r"/survey",
        r"ntt",
        r"boardid",
        r"bo_table",
    ]

    for pattern in url_patterns:
        if re.search(pattern, lower_url):
            board_score += 2

    if post_score >= 2:
        return "게시물 상세 후보"

    if board_score >= 3:
        return "게시판 후보"

    return "일반 페이지"


def extract_links(base_url, soup):
    results = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith("#"):
            continue

        if href.lower().startswith("javascript:"):
            continue

        if href.lower().startswith("mailto:"):
            continue

        if href.lower().startswith("tel:"):
            continue

        try:
            absolute = urljoin(base_url, href)
        except Exception:
            continue

        if not same_domain(base_url, absolute):
            continue

        text = clean_text(a.get_text(" ", strip=True))

        results.append({
            "url": absolute,
            "text": text,
        })

    return results


# =========================================================
# 페이지 분석
# =========================================================

async def fetch_page(session, url):
    try:

        async with session.get(
            url,
            headers=HEADERS,
            timeout=aiohttp.ClientTimeout(total=TIMEOUT),
            allow_redirects=True,
            ssl=False,
        ) as response:

            status = response.status

            content_type = response.headers.get(
                "Content-Type",
                ""
            )

            final_url = str(response.url)

            if not is_html_response(content_type):
                return {
                    "success": False,
                    "status": status,
                    "url": final_url,
                    "html": "",
                    "error": f"비HTML 응답: {content_type}",
                }

            html = await response.text(
                errors="ignore"
            )

            return {
                "success": True,
                "status": status,
                "url": final_url,
                "html": html,
                "error": "",
            }

    except asyncio.TimeoutError:

        return {
            "success": False,
            "status": 0,
            "url": url,
            "html": "",
            "error": "Timeout",
        }

    except aiohttp.ClientConnectorError:

        return {
            "success": False,
            "status": 0,
            "url": url,
            "html": "",
            "error": "ConnectionError",
        }

    except Exception as e:

        return {
            "success": False,
            "status": 0,
            "url": url,
            "html": "",
            "error": type(e).__name__,
        }


# =========================================================
# 후보 URL 생성
# =========================================================

def generate_candidate_urls(homepage):
    """
    홈페이지가 정상 접속되는 경우
    공공기관에서 자주 사용하는 경로를 추가 탐색
    """

    candidates = []

    base = homepage.rstrip("/") + "/"

    paths = [
        "notice",
        "notices",
        "news",
        "board",
        "bbs",
        "community",
        "customer",
        "customer/notice",
        "contents",
        "content",
        "data",
        "archive",
        "information",
        "info",
        "participation",
        "particip",
        "citizen",
        "event",
        "events",
        "survey",
        "media",
        "press",
        "pds",
        "reference",
        "download",
        "communication",
        "sotong",
        "알림",
        "알림마당",
        "공지사항",
        "소식",
        "자료실",
    ]

    for path in paths:
        candidates.append(urljoin(base, path))

    return candidates


# =========================================================
# 페이지별 분석
# =========================================================

def analyze_page(url, html, status):
    soup = BeautifulSoup(html, "html.parser")

    title = ""

    if soup.title:
        title = clean_text(
            soup.title.get_text(" ", strip=True)
        )

    visible_text = clean_text(
        soup.get_text(" ", strip=True)
    )

    links = extract_links(url, soup)

    board_candidates = []
    post_candidates = []

    for link in links:

        link_url = link["url"]
        link_text = link["text"]

        combined = f"{link_text} {link_url}"

        classification = classify_url(
            link_url,
            combined
        )

        if classification == "게시판 후보":
            board_candidates.append(
                (link_text, link_url)
            )

        elif classification == "게시물 상세 후보":
            post_candidates.append(
                (link_text, link_url)
            )

    # 현재 페이지 자체도 분석
    page_type = classify_url(
        url,
        f"{title} {visible_text[:5000]}"
    )

    return {
        "title": title,
        "text": visible_text,
        "iframe": detect_iframe(soup),
        "js": detect_js_dependency(
            soup,
            html
        ),
        "cms": detect_cms(
            html,
            url
        ),
        "links": links,
        "board_candidates": board_candidates,
        "post_candidates": post_candidates,
        "page_type": page_type,
        "error_page": detect_error_page(
            visible_text,
            status
        ),
    }


# =========================================================
# 기관 하나 분석
# =========================================================

async def analyze_institution(
    session,
    semaphore,
    institution,
    homepage,
    index,
):

    result = {
        "기관명": institution,
        "URL": homepage,
        "접속상태": "",
        "HTTP상태": "",
        "최종URL": "",
        "CMS/프레임워크": "",
        "iframe 사용": "",
        "JS 의존도": "",
        "메뉴 링크 발견": 0,
        "공지사항 후보": 0,
        "게시판 URL 패턴": "",
        "게시물 상세 URL 발견": 0,
        "게시판 유형": "",
        "추정 원인": "",
        "6차 공략법": "",
        "탐색 페이지 수": 0,
        "발견 URL 예시": "",
        "오류": "",
    }

    async with semaphore:

        print(
            f"[{index:03d}] {institution}"
        )

        # -------------------------------------------------
        # 1차: 홈페이지 접속
        # -------------------------------------------------

        first = await fetch_page(
            session,
            homepage
        )

        if not first["success"]:

            result["접속상태"] = "실패"
            result["HTTP상태"] = first["status"]
            result["최종URL"] = first["url"]
            result["추정 원인"] = first["error"]
            result["6차 공략법"] = (
                "접속 문제 해결 후 재탐색"
            )
            result["오류"] = first["error"]

            return result

        result["접속상태"] = "정상"
        result["HTTP상태"] = first["status"]
        result["최종URL"] = first["url"]

        first_analysis = analyze_page(
            first["url"],
            first["html"],
            first["status"]
        )

        result["CMS/프레임워크"] = (
            first_analysis["cms"]
        )

        result["iframe 사용"] = (
            first_analysis["iframe"]
        )

        result["JS 의존도"] = (
            first_analysis["js"]
        )

        # -------------------------------------------------
        # 탐색 큐
        # -------------------------------------------------

        queue = []

        visited = set()

        def add_url(url):
            if not url:
                return

            if url in visited:
                return

            if len(queue) >= MAX_PAGES:
                return

            queue.append(url)

        add_url(first["url"])

        # 홈페이지에서 발견된 링크
        for link in first_analysis["links"]:

            text = link["text"]
            url = link["url"]

            combined = (
                f"{text} {url}"
            ).lower()

            if any(
                word.lower() in combined
                for word in BOARD_WORDS
            ):
                add_url(url)

        # 공통 경로도 추가
        for url in generate_candidate_urls(
            first["url"]
        ):
            add_url(url)

        # -------------------------------------------------
        # BFS 탐색
        # -------------------------------------------------

        pages_analyzed = 0

        all_board_candidates = []
        all_post_candidates = []

        cms_counter = Counter()
        page_type_counter = Counter()

        discovered_urls = []

        while queue and pages_analyzed < MAX_PAGES:

            current_url = queue.pop(0)

            if current_url in visited:
                continue

            visited.add(current_url)

            page = await fetch_page(
                session,
                current_url
            )

            if not page["success"]:
                continue

            pages_analyzed += 1

            analysis = analyze_page(
                page["url"],
                page["html"],
                page["status"]
            )

            discovered_urls.append(
                page["url"]
            )

            cms_counter[
                analysis["cms"]
            ] += 1

            page_type_counter[
                analysis["page_type"]
            ] += 1

            # 게시판 후보
            all_board_candidates.extend(
                analysis["board_candidates"]
            )

            # 게시물 후보
            all_post_candidates.extend(
                analysis["post_candidates"]
            )

            # -------------------------------------------------
           
