# -*- coding: utf-8 -*-

"""
355개 공공기관 홈페이지 게시판 모니터링 시스템
5차 버전

핵심 변경
1. 실제 게시판 구조를 확인하여 오탐 감소
2. 오류 페이지 / 일반 콘텐츠 페이지 제거
3. 실제 게시글 링크가 있는 게시판 우선
4. 기관당 무조건 10개를 채우지 않음
5. 참여·공모·설문 관련 게시판 우선
6. 미발견 기관 재탐색 강화
7. 일일 모니터링은 최근 게시글 중심
8. Telegram 알림
9. seen_posts.json 중복 방지

환경변수
FORCE_DISCOVER_BOARDS=true
    -> 게시판 재탐색

FORCE_DISCOVER_BOARDS=false
    -> boards.xlsx를 이용하여 일일 모니터링
"""

import asyncio
import aiohttp
import pandas as pd
import os
import re
import json
import html

from urllib.parse import (
    urljoin,
    urlparse,
    urlunparse,
    parse_qs,
)

from bs4 import BeautifulSoup
from collections import deque


# =========================================================
# 기본 설정
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

URL_FILE = os.path.join(BASE_DIR, "url_완성.xlsx")
BOARDS_FILE = os.path.join(BASE_DIR, "boards.xlsx")
MISSING_FILE = os.path.join(BASE_DIR, "boards_missing.xlsx")
SEEN_FILE = os.path.join(BASE_DIR, "seen_posts.json")


TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")


FORCE_DISCOVER_BOARDS = (
    os.getenv("FORCE_DISCOVER_BOARDS", "false").lower()
    in ("true", "1", "yes")
)


# =========================================================
# 탐색 성능 설정
# =========================================================

CONCURRENCY = 12

# 기관당 최대 탐색 페이지
MAX_PAGES_PER_ORG = 50

# 홈페이지 → 메뉴 → 게시판 탐색 깊이
MAX_DEPTH = 4

# 기관당 저장할 최대 유효 게시판
# 4차의 10개 강제 구조를 완화
MAX_BOARDS_PER_ORG = 6

# HTTP timeout
REQUEST_TIMEOUT = 15

# 게시판 최소 점수
MIN_BOARD_SCORE = 7

# 게시판에서 확인할 최신 게시글 수
MAX_POSTS_PER_BOARD = 15


# =========================================================
# 검색 키워드
# =========================================================

KEYWORDS = [
    "설문조사",
    "시민참여",
    "국민참여",
    "공모전",
]


# =========================================================
# 게시판 관련 단어
# =========================================================

BOARD_WORDS = [
    "공지사항",
    "공지",
    "알림마당",
    "알림",
    "새소식",
    "소식",
    "게시판",
    "국민참여",
    "시민참여",
    "참여",
    "공모전",
    "공모",
    "설문조사",
    "설문",
    "고시공고",
    "고시·공고",
    "고시",
    "공고",
    "입찰",
    "뉴스",
    "보도자료",
    "자료실",
    "소식지",
    "행사",
    "이벤트",
    "채용",
    "교육",
    "정책자료",
    "정보공개",
    "입법예고",
    "의견수렴",
]


# =========================================================
# 우리가 특히 관심있는 게시판
# =========================================================

HIGH_VALUE_BOARD_WORDS = [
    "국민참여",
    "시민참여",
    "참여",
    "공모전",
    "공모",
    "설문조사",
    "설문",
    "의견수렴",
    "이벤트",
    "행사",
]


# =========================================================
# 사이트 구조 관련 단어
# =========================================================

SITE_MAP_WORDS = [
    "사이트맵",
    "사이트 맵",
    "전체메뉴",
    "전체 메뉴",
    "홈페이지맵",
    "홈페이지 맵",
]


# =========================================================
# URL 패턴
# =========================================================

BOARD_URL_PATTERNS = [
    "board",
    "bbs",
    "notice",
    "news",
    "community",
    "particip",
    "participation",
    "event",
    "survey",
    "list",
    "pds",
    "data",
    "announce",
    "public",
    "customer",
    "communication",
    "opinion",
]


DETAIL_PATTERNS = [
    "view",
    "detail",
    "read",
    "articleview",
    "articleview",
    "nttid",
    "bbsno",
    "boardno",
    "seq",
    "idx",
    "wr_id",
    "postid",
    "newsid",
]


IGNORE_EXTENSIONS = (
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".svg",
    ".webp",
    ".ico",
    ".css",
    ".js",
    ".woff",
    ".woff2",
    ".ttf",
    ".pdf",
    ".zip",
    ".hwp",
    ".hwpx",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
    ".mp4",
    ".avi",
    ".mov",
)


# =========================================================
# 오류 페이지 판별 단어
# =========================================================

ERROR_WORDS = [
    "404",
    "403",
    "500",
    "502",
    "503",
    "오류",
    "에러",
    "ERROR",
    "ERROR PAGE",
    "PAGE NOT FOUND",
    "NOT FOUND",
    "페이지를 찾을 수 없습니다",
    "페이지가 없습니다",
    "존재하지 않는 페이지",
    "잘못된 요청",
    "접근할 수 없습니다",
    "서비스 오류",
    "Basic Sample",
]


# =========================================================
# 일반 콘텐츠 페이지 판별 단어
# =========================================================

CONTENT_ONLY_WORDS = [
    "인사말",
    "기관소개",
    "기관 안내",
    "조직도",
    "조직 및 직원",
    "찾아오시는 길",
    "오시는 길",
    "연혁",
    "비전",
    "미션",
    "설립목적",
    "경영목표",
    "경영공시",
    "직원검색",
    "업무안내",
    "부서안내",
    "전화번호",
    "이용안내",
    "개인정보처리방침",
    "저작권정책",
    "이용약관",
    "사이트맵",
]


# =========================================================
# HTTP Header
# =========================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,*/*;q=0.8"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}


# =========================================================
# 공통 함수
# =========================================================

def clean_text(value):
    if value is None:
        return ""

    value = html.unescape(str(value))
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url.startswith(("http://", "https://")):
        return ""

    parsed = urlparse(url)

    scheme = parsed.scheme.lower()
    netloc = parsed.netloc.lower()

    path = parsed.path or "/"
    path = re.sub(r"/{2,}", "/", path)

    return urlunparse(
        (
            scheme,
            netloc,
            path,
            "",
            parsed.query,
            "",
        )
    )


def same_domain(url1, url2):

    try:

        a = (
            urlparse(url1)
            .netloc
            .lower()
            .replace("www.", "")
        )

        b = (
            urlparse(url2)
            .netloc
            .lower()
            .replace("www.", "")
        )

        return a == b

    except Exception:
        return False


def is_http_url(url):
    return url.startswith(("http://", "https://"))


def is_ignored_file(url):

    path = urlparse(url).path.lower()

    return any(
        path.endswith(ext)
        for ext in IGNORE_EXTENSIONS
    )


# =========================================================
# 상세 페이지 판별
# =========================================================

def is_detail_url(url):

    lower = url.lower()

    query = parse_qs(
        urlparse(url).query
    )

    for key in query.keys():

        key_lower = key.lower()

        if key_lower in [
            "nttid",
            "bbsno",
            "boardno",
            "seq",
            "idx",
            "wr_id",
            "postid",
            "newsid",
        ]:
            return True

    for pattern in DETAIL_PATTERNS:

        if pattern.lower() in lower:
            return True

    return False


# =========================================================
# 오류 페이지 판별
# =========================================================

def is_error_page(soup, url=""):

    title = ""

    title_tag = soup.find("title")

    if title_tag:
        title = clean_text(
            title_tag.get_text()
        )

    body_text = clean_text(
        soup.get_text(" ", strip=True)
    )

    sample = (
        title[:500]
        + " "
        + body_text[:3000]
    ).upper()

    for word in ERROR_WORDS:

        if word.upper() in sample:
            return True

    return False


# =========================================================
# 일반 콘텐츠 페이지 판별
# =========================================================

def is_content_only_page(soup, title=""):

    text = clean_text(
        soup.get_text(" ", strip=True)
    )

    title_lower = title.lower()

    matched = 0

    for word in CONTENT_ONLY_WORDS:

        if word.lower() in title_lower:
            matched += 2

        elif word in text[:1500]:
            matched += 1

    # 일반 콘텐츠 페이지에서 흔히 나타나는 구조
    if matched >= 3:

        # 단, 실제 게시판 구조가 있으면 예외
        board_signal = detect_real_board_structure(
            soup
        )

        if board_signal["is_board"]:
            return False

        return True

    return False


# =========================================================
# 게시판 메뉴명 판별
# =========================================================

def link_text_is_board(text):

    text = clean_text(text)

    if not text:
        return False

    return any(
        word in text
        for word in BOARD_WORDS
    )


# =========================================================
# URL 게시판 가능성
# =========================================================

def url_looks_like_board(url):

    lower = url.lower()

    return any(
        pattern in lower
        for pattern in BOARD_URL_PATTERNS
    )


# =========================================================
# 실제 게시판 구조 탐지
# =========================================================

def detect_real_board_structure(soup):

    score = 0
    reasons = []

    if is_error_page(soup):
        return {
            "is_board": False,
            "score": -20,
            "reasons": "오류페이지",
        }

    text = clean_text(
        soup.get_text(" ", strip=True)
    )

    # -----------------------------------------------------
    # 게시판 컬럼
    # -----------------------------------------------------

    board_columns = [
        "번호",
        "제목",
        "등록일",
        "작성일",
        "조회수",
        "작성자",
        "첨부파일",
    ]

    found_columns = []

    for word in board_columns:

        if word in text:
            found_columns.append(word)

    if len(found_columns) >= 3:

        score += 7

        reasons.append(
            "게시판컬럼:"
            + ",".join(found_columns[:6])
        )

    elif len(found_columns) == 2:

        score += 4

        reasons.append(
            "게시판컬럼:"
            + ",".join(found_columns)
        )

    elif len(found_columns) == 1:

        score += 1

    # -----------------------------------------------------
    # table
    # -----------------------------------------------------

    tables = soup.find_all("table")

    if tables:

        for table in tables:

            rows = table.find_all("tr")

            if len(rows) >= 3:

                score += 4

                reasons.append(
                    f"table행:{len(rows)}"
                )

                break

    # -----------------------------------------------------
    # 게시글 링크 후보
    # -----------------------------------------------------

    post_link_count = 0

    for a in soup.find_all(
        "a",
        href=True,
    ):

        href = a.get(
            "href",
            "",
        ).strip()

        text_value = clean_text(
            a.get_text(
                " ",
                strip=True,
            )
        )

        if not text_value:
            continue

        if len(text_value) < 2:
            continue

        if is_detail_url(
            urljoin("https://dummy.local", href)
        ):
            post_link_count += 1

    if post_link_count >= 5:

        score += 6

        reasons.append(
            f"게시글링크:{post_link_count}"
        )

    elif post_link_count >= 2:

        score += 3

        reasons.append(
            f"게시글링크:{post_link_count}"
        )

    # -----------------------------------------------------
    # 페이지 이동
    # -----------------------------------------------------

    page_words = [
        "다음",
        "이전",
        "페이지",
        "목록",
        "검색",
    ]

    page_found = [
        word
        for word in page_words
        if word in text
    ]

    if len(page_found) >= 3:

        score += 3

        reasons.append(
            "페이지요소:"
            + ",".join(page_found)
        )

    # -----------------------------------------------------
    # 검색 form
    # -----------------------------------------------------

    if soup.find("form"):

        score += 1

        reasons.append(
            "검색form"
        )

    # -----------------------------------------------------
    # 최종 판단
    # -----------------------------------------------------

    is_board = False

    if (
        score >= 8
        and (
            len(found_columns) >= 2
            or post_link_count >= 3
            or len(tables) >= 1
        )
    ):
        is_board = True

    return {
        "is_board": is_board,
        "score": score,
        "reasons": "; ".join(
            dict.fromkeys(reasons)
        ),
    }


# =========================================================
# 후보 점수
# =========================================================

def score_candidate(
    url,
    link_text,
    soup=None,
):

    score = 0
    reasons = []

    text = clean_text(link_text)
    lower_text = text.lower()
    lower_url = url.lower()

    # -----------------------------------------------------
    # 메뉴명
    # -----------------------------------------------------

    for word in BOARD_WORDS:

        if word.lower() in lower_text:

            score += 5

            reasons.append(
                f"메뉴:{word}"
            )

            break

    # -----------------------------------------------------
    # 고가치 게시판
    # -----------------------------------------------------

    for word in HIGH_VALUE_BOARD_WORDS:

        if word.lower() in lower_text:

            score += 5

            reasons.append(
                f"관심메뉴:{word}"
            )

            break

    # -----------------------------------------------------
    # URL
    # -----------------------------------------------------

    for pattern in BOARD_URL_PATTERNS:

        if pattern.lower() in lower_url:

            score += 2

            reasons.append(
                f"URL:{pattern}"
            )

            break

    # -----------------------------------------------------
    # 상세페이지 제외
    # -----------------------------------------------------

    if is_detail_url(url):

        score -= 10

        reasons.append(
            "상세페이지"
        )

    # -----------------------------------------------------
    # 실제 HTML 구조
    # -----------------------------------------------------

    if soup is not None:

        structure = detect_real_board_structure(
            soup
        )

        if structure["is_board"]:

            score += structure["score"]

            reasons.append(
                structure["reasons"]
            )

        else:

            score += max(
                0,
                structure["score"] - 3,
            )

    return {
        "score": int(score),
        "reasons": "; ".join(
            dict.fromkeys(reasons)
        ),
    }


# =========================================================
# HTTP
# =========================================================

async def fetch(
    session,
    url,
):

    try:

        timeout = aiohttp.ClientTimeout(
            total=REQUEST_TIMEOUT
        )

        async with session.get(
            url,
            headers=HEADERS,
            timeout=timeout,
            allow_redirects=True,
            ssl=False,
        ) as response:

            status = response.status

            final_url = str(
                response.url
            )

            if status >= 400:

                return (
                    None,
                    status,
                    final_url,
                )

            content_type = (
                response.headers
                .get(
                    "Content-Type",
                    "",
                )
                .lower()
            )

            if (
                "html" not in content_type
                and "xml" not in content_type
                and content_type
            ):

                return (
                    None,
                    status,
                    final_url,
                )

            raw = await response.read()

            try:

                text = raw.decode(
                    "utf-8"
                )

            except UnicodeDecodeError:

                try:

                    text = raw.decode(
                        "cp949"
                    )

                except UnicodeDecodeError:

                    text = raw.decode(
                        "euc-kr",
                        errors="ignore",
                    )

            return (
                text,
                status,
                final_url,
            )

    except asyncio.TimeoutError:

        return (
            None,
            "timeout",
            url,
        )

    except Exception as e:

        return (
            None,
            type(e).__name__,
            url,
        )


# =========================================================
# 링크 추출
# =========================================================

def extract_links(
    soup,
    base_url,
):

    results = []

    for a in soup.find_all(
        "a",
        href=True,
    ):

        href = a.get(
            "href",
            "",
        ).strip()

        if not href:
            continue

        if href.startswith(
            (
                "javascript:",
                "mailto:",
                "tel:",
                "#",
            )
        ):
            continue

        full_url = normalize_url(
            urljoin(
                base_url,
                href,
            )
        )

        if not is_http_url(
            full_url
        ):
            continue

        if is_ignored_file(
            full_url
        ):
            continue

        text = clean_text(
            a.get_text(
                " ",
                strip=True,
            )
        )

        results.append(
            (
                full_url,
                text,
            )
        )

    # iframe
    for tag in soup.find_all(
        [
            "iframe",
            "frame",
        ]
    ):

        src = tag.get("src")

        if not src:
            continue

        full_url = normalize_url(
            urljoin(
                base_url,
                src,
            )
        )

        if is_http_url(
            full_url
        ):

            results.append(
                (
                    full_url,
                    "iframe",
                )
            )

    unique = []

    seen = set()

    for url, text in results:

        if url in seen:
            continue

        seen.add(url)

        unique.append(
            (
                url,
                text,
            )
        )

    return unique


# =========================================================
# robots sitemap
# =========================================================

async def discover_robots(
    session,
    homepage,
):

    parsed = urlparse(
        homepage
    )

    robots_url = urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            "/robots.txt",
            "",
            "",
            "",
        )
    )

    text, status, _ = await fetch(
        session,
        robots_url,
    )

    if not text:
        return []

    urls = []

    for line in text.splitlines():

        if line.lower().startswith(
            "sitemap:"
        ):

            sitemap = line.split(
                ":",
                1,
            )[1].strip()

            sitemap = normalize_url(
                sitemap
            )

            if sitemap:
                urls.append(
                    sitemap
                )

    return urls


# =========================================================
# sitemap
# =========================================================

async def parse_sitemap(
    session,
    sitemap_url,
):

    text, status, final_url = await fetch(
        session,
        sitemap_url,
    )

    if not text:
        return []

    soup = BeautifulSoup(
        text,
        "xml",
    )

    urls = []

    for loc in soup.find_all(
        "loc"
    ):

        value = clean_text(
            loc.get_text()
        )

        value = normalize_url(
            value
        )

        if value:
            urls.append(
                value
            )

    return urls


# =========================================================
# 공통 URL
# =========================================================

def common_paths(
    homepage
):

    parsed = urlparse(
        homepage
    )

    base = (
        f"{parsed.scheme}://"
        f"{parsed.netloc}"
    )

    paths = [
        "/notice",
        "/board",
        "/bbs",
        "/news",
        "/community",
        "/participation",
        "/particip",
        "/event",
        "/survey",
        "/customer",
        "/contents",
        "/content",
        "/sitemap",
        "/siteMap",
        "/main",
        "/index",
    ]

    return [
        normalize_url(
            urljoin(
                base,
                path,
            )
        )
        for path in paths
    ]


# =========================================================
# 후보 추가
# =========================================================

def add_candidate(
    candidates,
    url,
    name,
    score,
    reasons,
):

    url = normalize_url(url)

    if not url:
        return

    if is_detail_url(url):
        return

    candidate = {
        "url": url,
        "name": (
            clean_text(name)
            or "게시판"
        ),
        "score": int(score),
        "reasons": clean_text(
            reasons
        ),
    }

    existing = candidates.get(
        url
    )

    if existing is None:

        candidates[url] = candidate

        return

    if (
        int(candidate["score"])
        > int(
            existing.get(
                "score",
                -999,
            )
        )
    ):

        candidates[url] = candidate


# =========================================================
# 게시판 실제 검증
# =========================================================

async def validate_board(
    session,
    url,
):

    text, status, final_url = await fetch(
        session,
        url,
    )

    if not text:

        return None

    try:

        soup = BeautifulSoup(
            text,
            "html.parser",
        )

    except Exception:

        return None

    if is_error_page(
        soup,
        url,
    ):

        return None

    title = ""

    title_tag = soup.find(
        "title"
    )

    if title_tag:

        title = clean_text(
            title_tag.get_text()
        )

    if is_content_only_page(
        soup,
        title,
    ):

        return None

    structure = (
        detect_real_board_structure(
            soup
        )
    )

    if not structure[
        "is_board"
    ]:

        return None

    return {
        "title": title,
        "score": structure[
            "score"
        ],
        "reasons": structure[
            "reasons"
        ],
    }


# =========================================================
# 기관 1개 탐색
# =========================================================

async def discover_one(
    session,
    semaphore,
    index,
    total,
    row,
):

    name = clean_text(
        row.get(
            "기관명",
            "",
        )
    )

    institution_type = clean_text(
        row.get(
            "기관유형",
            "",
        )
    )

    homepage = normalize_url(
        row.get(
            "URL",
            "",
        )
    )

    result = {
        "기관명": name,
        "기관유형": institution_type,
        "홈페이지": homepage,
        "boards": [],
        "status": "실패",
        "note": "",
        "pages": 0,
    }

    if not homepage:

        result["note"] = (
            "홈페이지 URL 없음"
        )

        print(
            f"[{index:03d}/{total}] "
            f"{name} → 0개 "
            "(홈페이지 URL 없음)"
        )

        return result

    async with semaphore:

        try:

            # ------------------------------------------------
            # 홈페이지
            # ------------------------------------------------

            text, status, final_url = await fetch(
                session,
                homepage,
            )

            if not text:

                result["note"] = (
                    "홈페이지 접속 시간 초과"
                    if status == "timeout"
                    else
                    f"홈페이지 접속 실패: {status}"
                )

                print(
                    f"[{index:03d}/{total}] "
                    f"{name} → 0개 "
                    f"({result['note']})"
                )

                return result

            homepage = (
                normalize_url(
                    final_url
                )
                or homepage
            )

            queue = deque()

            queue.append(
                (
                    homepage,
                    0,
                    "홈페이지",
                )
            )

            visited = set()

            candidates = {}

            # ------------------------------------------------
            # robots sitemap
            # ------------------------------------------------

            try:

                robots_sitemaps = (
                    await discover_robots(
                        session,
                        homepage,
                    )
                )

                for sitemap in robots_sitemaps:

                    queue.append(
                        (
                            sitemap,
                            1,
                            "robots",
                        )
                    )

            except Exception:
                pass

            # ------------------------------------------------
            # 공통경로
            # ------------------------------------------------

            for common_url in common_paths(
                homepage
            ):

                queue.append(
                    (
                        common_url,
                        1,
                        "공통경로",
                    )
                )

            # ------------------------------------------------
            # 탐색
            # ------------------------------------------------

            while (
                queue
                and len(visited)
                < MAX_PAGES_PER_ORG
            ):

                current_url, depth, source = (
                    queue.popleft()
                )

                current_url = normalize_url(
                    current_url
                )

                if not current_url:
                    continue

                if current_url in visited:
                    continue

                visited.add(
                    current_url
                )

                if not is_http_url(
                    current_url
                ):
                    continue

                if is_ignored_file(
                    current_url
                ):
                    continue

                if not same_domain(
                    current_url,
                    homepage,
                ):
                    continue

                page_text, page_status, final = (
                    await fetch(
                        session,
                        current_url,
                    )
                )

                if not page_text:
                    continue

                result["pages"] += 1

                try:

                    soup = BeautifulSoup(
                        page_text,
                        "html.parser",
                    )

                except Exception:

                    continue

                # 오류 페이지는 탐색하지 않음
                if is_error_page(
                    soup,
                    current_url,
                ):
                    continue

                title = ""

                title_tag = soup.find(
                    "title"
                )

                if title_tag:

                    title = clean_text(
                        title_tag.get_text()
                    )

                # ------------------------------------------------
                # 현재 페이지 자체가 게시판인지
                # ------------------------------------------------

                structure = (
                    detect_real_board_structure(
                        soup
                    )
                )

                if structure[
                    "is_board"
                ]:

                    score_data = score_candidate(
                        current_url,
                        title,
                        soup,
                    )

                    total_score = (
                        score_data["score"]
                        + structure["score"]
                    )

                    if (
                        total_score
                        >= MIN_BOARD_SCORE
                    ):

                        add_candidate(
                            candidates,
                            current_url,
                            title or "게시판",
                            total_score,
                            (
                                score_data["reasons"]
                                + ";"
                                + structure[
                                    "reasons"
                                ]
                            ),
                        )

                # ------------------------------------------------
                # 링크
                # ------------------------------------------------

                links = extract_links(
                    soup,
                    current_url,
                )

                for (
                    link_url,
                    link_text,
                ) in links:

                    if not same_domain(
                        link_url,
                        homepage,
                    ):
                        continue

                    if is_detail_url(
                        link_url
                    ):
                        continue

                    data = score_candidate(
                        link_url,
                        link_text,
                        None,
                    )

                    menu_match = (
                        link_text_is_board(
                            link_text
                        )
                    )

                    url_match = (
                        url_looks_like_board(
                            link_url
                        )
                    )

                    site_map_match = any(
                        word in link_text
                        for word in SITE_MAP_WORDS
                    )

                    # ------------------------------------------------
                    # 후보 등록
                    # ------------------------------------------------

                    if (
                        data["score"]
                        >= MIN_BOARD_SCORE
                    ):

                        add_candidate(
                            candidates,
                            link_url,
                            link_text or "게시판",
                            data["score"],
                            data["reasons"],
                        )

                    # ------------------------------------------------
                    # 다음 페이지 탐색
                    # ------------------------------------------------

                    if (
                        depth
                        >= MAX_DEPTH
                    ):
                        continue

                    if (
                        menu_match
                        or url_match
                        or site_map_match
                        or depth == 0
                    ):

                        if (
                            link_url
                            not in visited
                        ):

                            queue.append(
                                (
                                    link_url,
                                    depth + 1,
                                    "링크",
                                )
                            )

            # =====================================================
            # 후보 최종 검증
            # =====================================================

            preliminary = sorted(
                candidates.values(),
                key=lambda x: int(
                    x.get(
                        "score",
                        0,
                    )
                ),
                reverse=True,
            )

            final_candidates = []

            seen_urls = set()

            # 상위 후보를 실제 HTTP로 재검증
            for candidate in preliminary:

                if len(
                    final_candidates
                ) >= MAX_BOARDS_PER_ORG:
                    break

                board_url = normalize_url(
                    candidate.get(
                        "url",
                        "",
                    )
                )

                if not board_url:
                    continue

                if board_url in seen_urls:
                    continue

                if is_detail_url(
                    board_url
                ):
                    continue

                # 실제 게시판 검증
                validated = await validate_board(
                    session,
                    board_url,
                )

                if not validated:
                    continue

                final_score = (
                    int(
                        candidate.get(
                            "score",
                            0,
                        )
                    )
                    + int(
                        validated.get(
                            "score",
                            0,
                        )
                    )
                )

                # 관심 키워드 게시판 추가 가점
                board_name = clean_text(
                    candidate.get(
                        "name",
                        "",
                    )
                )

                for word in HIGH_VALUE_BOARD_WORDS:

                    if word in board_name:

                        final_score += 5

                        break

                if final_score < (
                    MIN_BOARD_SCORE + 3
                ):
                    continue

                seen_urls.add(
                    board_url
                )

                final_candidates.append(
                    {
                        "url": board_url,
                        "name": (
                            board_name
                            or validated.get(
                                "title",
                                "게시판",
                            )
                        ),
                        "score": final_score,
                        "reasons": (
                            candidate.get(
                                "reasons",
                                "",
                            )
                            + ";검증:"
                            + validated.get(
                                "reasons",
                                "",
                            )
                        ),
                    }
                )

            # ------------------------------------------------
            # 점수순 정렬
            # ------------------------------------------------

            final_candidates = sorted(
                final_candidates,
                key=lambda x: int(
                    x.get(
                        "score",
                        0,
                    )
                ),
                reverse=True,
            )

            result["boards"] = (
                final_candidates[
                    :MAX_BOARDS_PER_ORG
                ]
            )

            # ------------------------------------------------
            # 결과
            # ------------------------------------------------

            if result["boards"]:

                result["status"] = "정상"

                result["note"] = (
                    "유효 게시판 "
                    f"{len(result['boards'])}개"
                )

                print(
                    f"[{index:03d}/{total}] "
                    f"{name} → "
                    f"{len(result['boards'])}개 "
                    f"(탐색 {result['pages']}페이지)"
                )

            else:

                result["status"] = "실패"

                result["note"] = (
                    "유효 게시판 없음"
                )

                print(
                    f"[{index:03d}/{total}] "
                    f"{name} → 0개 "
                    f"(탐색 {result['pages']}페이지)"
                )

            return result

        except Exception as e:

            result["status"] = "실패"

            result["note"] = (
                f"탐색 오류: "
                f"{type(e).__name__}: {e}"
            )

            print(
                f"[{index:03d}/{total}] "
                f"{name} → 0개 "
                f"(탐색 오류: "
                f"{type(e).__name__}: {e})"
            )

            return result


# =========================================================
# 전체 기관 탐색
# =========================================================

async def discover_all():

    print("=" * 70)
    print("5차 게시판 탐색 시작")
    print("=" * 70)

    print(
        f"전체 기관 : "
        f"{len(institutions)}개"
    )

    print(
        f"동시 탐색 : "
        f"{CONCURRENCY}개"
    )

    print(
        f"기관당 최대 페이지 : "
        f"{MAX_PAGES_PER_ORG}"
    )

    print(
        f"최대 탐색 깊이 : "
        f"{MAX_DEPTH}"
    )

    print(
        f"기관당 최대 게시판 : "
        f"{MAX_BOARDS_PER_ORG}개"
    )

    print(
        f"최소 점수 : "
        f"{MIN_BOARD_SCORE}"
    )

    print("=" * 70)

    semaphore = asyncio.Semaphore(
        CONCURRENCY
    )

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY * 2,
        ssl=False,
        ttl_dns_cache=300,
    )

    timeout = aiohttp.ClientTimeout(
        total=REQUEST_TIMEOUT
    )

    async with aiohttp.ClientSession(
        connector=connector,
        timeout=timeout,
        headers=HEADERS,
    ) as session:

        async def worker(
            index,
            row,
        ):

            return await discover_one(
                session,
                semaphore,
                index,
                len(institutions),
                row,
            )

        tasks = [
            worker(
                i + 1,
                row,
            )
            for i, row
            in institutions.iterrows()
        ]

        raw_results = await asyncio.gather(
            *tasks,
            return_exceptions=True,
        )

    results = []

    for i, item in enumerate(
        raw_results,
        start=1,
    ):

        if isinstance(
            item,
            Exception,
        ):

            row = institutions.iloc[
                i - 1
            ]

            results.append(
                {
                    "기관명": clean_text(
                        row.get(
                            "기관명",
                            "",
                        )
                    ),
                    "기관유형": clean_text(
                        row.get(
                            "기관유형",
                            "",
                        )
                    ),
                    "홈페이지": normalize_url(
                        row.get(
                            "URL",
                            "",
                        )
                    ),
                    "boards": [],
                    "status": "실패",
                    "note": (
                        f"worker 오류: "
                        f"{type(item).__name__}: "
                        f"{item}"
                    ),
                    "pages": 0,
                }
            )

        else:

            results.append(item)

    # =====================================================
    # boards.xlsx
    # =====================================================

    board_rows = []

    for result in results:

        for board in result.get(
            "boards",
            [],
        ):

            board_rows.append(
                {
                    "기관명": result.get(
                        "기관명",
                        "",
                    ),
                    "기관유형": result.get(
                        "기관유형",
                        "",
                    ),
                    "홈페이지": result.get(
                        "홈페이지",
                        "",
                    ),
                    "게시판명": board.get(
                        "name",
                        "게시판",
                    ),
                    "게시판URL": board.get(
                        "url",
                        "",
                    ),
                    "점수": int(
                        board.get(
                            "score",
                            0,
                        )
                    ),
                    "판별근거": board.get(
                        "reasons",
                        "",
                    ),
                }
            )

    boards_df = pd.DataFrame(
        board_rows,
        columns=[
            "기관명",
            "기관유형",
            "홈페이지",
            "게시판명",
            "게시판URL",
            "점수",
            "판별근거",
        ],
    )

    if not boards_df.empty:

        boards_df = (
            boards_df
            .drop_duplicates(
                subset=[
                    "기관명",
                    "게시판URL",
                ]
            )
            .sort_values(
                [
                    "기관명",
                    "점수",
                ],
                ascending=[
                    True,
                    False,
                ],
            )
        )

    boards_df.to_excel(
        BOARDS_FILE,
        index=False,
    )

    # =====================================================
    # boards_missing.xlsx
    # =====================================================

    missing_rows = []

    for result in results:

        if not result.get(
            "boards"
        ):

            missing_rows.append(
                {
                    "기관명": result.get(
                        "기관명",
                        "",
                    ),
                    "기관유형": result.get(
                        "기관유형",
                        "",
                    ),
                    "URL": result.get(
                        "홈페이지",
                        "",
                    ),
                    "상태": result.get(
                        "status",
                        "",
                    ),
                    "비고": result.get(
                        "note",
                        "",
                    ),
                    "탐색페이지": result.get(
                        "pages",
                        0,
                    ),
                }
            )

    missing_df = pd.DataFrame(
        missing_rows,
        columns=[
            "기관명",
            "기관유형",
            "URL",
            "상태",
            "비고",
            "탐색페이지",
        ],
    )

    missing_df.to_excel(
        MISSING_FILE,
        index=False,
    )

    # =====================================================
    # 결과
    # =====================================================

    total_orgs = len(
        institutions
    )

    found_orgs = sum(
        1
        for result in results
        if result.get("boards")
    )

    missing_orgs = (
        total_orgs
        - found_orgs
    )

    coverage = (
        found_orgs
        / total_orgs
        * 100
        if total_orgs
        else 0
    )

    print()

    print("=" * 70)
    print("5차 게시판 탐색 완료")
    print("=" * 70)

    print(
        f"전체 기관       : "
        f"{total_orgs}개"
    )

    print(
        f"게시판 발견 기관 : "
        f"{found_orgs}개"
    )

    print(
        f"게시판 미발견    : "
        f"{missing_orgs}개"
    )

    print(
        f"기관 기준 발견률 : "
        f"{coverage:.1f}%"
    )

    print(
        f"게시판 총 개수   : "
        f"{len(boards_df)}개"
    )

    print(
        f"boards.xlsx      : "
        f"{BOARDS_FILE}"
    )

    print(
        f"boards_missing   : "
        f"{MISSING_FILE}"
    )

    print("=" * 70)

    return boards_df


# =========================================================
# seen_posts
# =========================================================

def load_seen():

    if not os.path.exists(
        SEEN_FILE
    ):
        return set()

    try:

        with open(
            SEEN_FILE,
            "r",
            encoding="utf-8",
        ) as f:

            data = json.load(f)

        if isinstance(
            data,
            list,
        ):

            return set(data)

        if isinstance(
            data,
            dict,
        ):

            return set(
                data.keys()
            )

    except Exception as e:

        print(
            "seen_posts.json 읽기 오류:",
            e,
        )

    return set()


def save_seen(seen):

    try:

        with open(
            SEEN_FILE,
            "w",
            encoding="utf-8",
        ) as f:

            json.dump(
                sorted(
                    list(seen)
                ),
                f,
                ensure_ascii=False,
                indent=2,
            )

    except Exception as e:

        print(
            "seen_posts.json 저장 오류:",
            e,
        )


# =========================================================
# 게시글 정보
# =========================================================

def extract_post_info(
    soup,
    url,
):

    title = ""

    for selector in [
        "h1",
        "h2",
        "h3",
        ".title",
        ".subject",
        ".view_title",
        ".board_title",
        ".bbs_title",
    ]:

        tag = soup.select_one(
            selector
        )

        if tag:

            value = clean_text(
                tag.get_text(
                    " ",
                    strip=True,
                )
            )

            if (
                2
                <= len(value)
                <= 300
            ):

                title = value

                break

    if not title:

        tag = soup.find(
            "title"
        )

        if tag:

            title = clean_text(
                tag.get_text()
            )

    content = ""

    for selector in [
        ".view",
        ".view_cont",
        ".view_content",
        ".board_view",
        ".bbs_view",
        ".content",
        ".contents",
        "article",
        "main",
    ]:

        tag = soup.select_one(
            selector
        )

        if tag:

            value = clean_text(
                tag.get_text(
                    " ",
                    strip=True,
                )
            )

            if len(value) > len(
                content
            ):

                content = value

    if not content:

        content = clean_text(
            soup.get_text(
                " ",
                strip=True,
            )
        )

    return {
        "url": url,
        "title": title,
        "content": content,
    }


# =========================================================
# 게시글 링크
# =========================================================

def extract_post_links(
    soup,
    board_url,
):

    links = []

    # -----------------------------------------------------
    # table 우선
    # -----------------------------------------------------

    tables = soup.find_all(
        "table"
    )

    if tables:

        for table in tables:

            for a in table.find_all(
                "a",
                href=True,
            ):

                href = a.get(
                    "href",
                    "",
                ).strip()

                if not href:
                    continue

                if href.startswith(
                    (
                        "javascript:",
                        "mailto:",
                        "#",
                    )
                ):
                    continue

                full_url = normalize_url(
                    urljoin(
                        board_url,
                        href,
                    )
                )

                if not is_http_url(
                    full_url
                ):
                    continue

                if is_ignored_file(
                    full_url
                ):
                    continue

                if not same_domain(
                    full_url,
                    board_url,
                ):
                    continue

                text = clean_text(
                    a.get_text(
                        " ",
                        strip=True,
                    )
                )

                if not text:
                    continue

                links.append(
                    (
                        full_url,
                        text,
                    )
                )

    # -----------------------------------------------------
    # 일반 링크
    # -----------------------------------------------------

    if not links:

        for a in soup.find_all(
            "a",
            href=True,
        ):

            href = a.get(
                "href",
                "",
            ).strip()

            if not href:
                continue

            if href.startswith(
                (
                    "javascript:",
                    "mailto:",
                    "#",
                )
            ):
                continue

            full_url = normalize_url(
                urljoin(
                    board_url,
                    href,
                )
            )

            if not is_http_url(
                full_url
            ):
                continue

            if is_ignored_file(
                full_url
            ):
                continue

            if not same_domain(
                full_url,
                board_url,
            ):
                continue

            text = clean_text(
                a.get_text(
                    " ",
                    strip=True,
                )
            )

            if not text:
                continue

            if is_detail_url(
                full_url
            ):

                links.append(
                    (
                        full_url,
                        text,
                    )
                )

    # -----------------------------------------------------
    # 중복 제거
    # -----------------------------------------------------

    result = []

    seen = set()

    for url, text in links:

        if url in seen:
            continue

        seen.add(url)

        result.append(
            (
                url,
                text,
            )
        )

        if (
            len(result)
            >= MAX_POSTS_PER_BOARD
        ):
            break

    return result


# =========================================================
# 키워드
# =========================================================

def find_keywords(
    title,
    content,
):

    combined = (
        clean_text(title)
        + " "
        + clean_text(content)
    )

    return [
        keyword
        for keyword in KEYWORDS
        if keyword in combined
    ]


# =========================================================
# Telegram
# =========================================================

async def send_telegram(
    session,
    message,
):

    if not TELEGRAM_BOT_TOKEN:

        print(
            "Telegram BOT TOKEN 없음"
        )

        return False

    if not TELEGRAM_CHAT_ID:

        print(
            "Telegram CHAT ID 없음"
        )

        return False

    url = (
        "https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/"
        "sendMessage"
    )

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "disable_web_page_preview": True,
    }

    try:

        async with session.post(
            url,
            data=payload,
            timeout=20,
            ssl=False,
        ) as response:

            if response.status == 200:

                return True

            body = await response.text()

            print(
                "Telegram 오류:",
                response.status,
                body[:500],
            )

            return False

    except Exception as e:

        print(
            "Telegram 전송 오류:",
            e,
        )

        return False


# =========================================================
# 게시판 모니터링
# =========================================================

async def monitor_board(
    session,
    board_row,
    seen,
):

    institution = clean_text(
        board_row.get(
            "기관명",
            "",
        )
    )

    board_name = clean_text(
        board_row.get(
            "게시판명",
            "게시판",
        )
    )

    board_url = normalize_url(
        board_row.get(
            "게시판URL",
            "",
        )
    )

    if not board_url:

        return 0

    text, status, final_url = (
        await fetch(
            session,
            board_url,
        )
    )

    if not text:

        print(
            f"[게시판 접속 실패] "
            f"{institution} / "
            f"{board_name}"
        )

        return 0

    try:

        soup = BeautifulSoup(
            text,
            "html.parser",
        )

    except Exception:

        return 0

    # 오류 페이지
    if is_error_page(
        soup,
        board_url,
    ):

        return 0

    post_links = extract_post_links(
        soup,
        board_url,
    )

    alerts = 0

    for (
        post_url,
        link_text,
    ) in post_links:

        if post_url in seen:

            continue

        post_text, post_status, _ = (
            await fetch(
                session,
                post_url,
            )
        )

        if not post_text:

            continue

        try:

            post_soup = BeautifulSoup(
                post_text,
                "html.parser",
            )

        except Exception:

            continue

        if is_error_page(
            post_soup,
            post_url,
        ):

            continue

        info = extract_post_info(
            post_soup,
            post_url,
        )

        title = (
            info["title"]
            or link_text
        )

        content = info[
            "content"
        ]

        found = find_keywords(
            title,
            content,
        )

        # 게시글은 확인했으므로 seen 처리
        seen.add(
            post_url
        )

        if not found:

            continue

        message = (
            "🔔 공공기관 게시글 알림\n\n"
            f"기관명: {institution}\n"
            f"게시판: {board_name}\n"
            f"제목: {title}\n"
            f"키워드: {', '.join(found)}\n\n"
            f"{post_url}"
        )

        success = await send_telegram(
            session,
            message,
        )

        if success:

            alerts += 1

            print(
                f"[알림] "
                f"{institution} / "
                f"{title} / "
                f"{', '.join(found)}"
            )

    return alerts


# =========================================================
# 전체 모니터링
# =========================================================

async def monitor_all():

    if not os.path.exists(
        BOARDS_FILE
    ):

        print(
            "boards.xlsx가 없습니다."
        )

        print(
            "먼저 게시판 탐색을 실행하세요."
        )

        return

    try:

        boards_df = pd.read_excel(
            BOARDS_FILE
        )

    except Exception as e:

        print(
            f"boards.xlsx 읽기 실패: {e}"
        )

        return

    if boards_df.empty:

        print(
            "boards.xlsx에 게시판이 없습니다."
        )

        return

    seen = load_seen()

    print("=" * 70)
    print(
        "일일 게시판 모니터링 시작"
    )
    print("=" * 70)

    print(
        f"게시판 수 : "
        f"{len(boards_df)}개"
    )

    print(
        f"기존 seen : "
        f"{len(seen)}개"
    )

    print(
        "검색 키워드 : "
        + ", ".join(KEYWORDS)
    )

    print("=" * 70)

    connector = aiohttp.TCPConnector(
        limit=20,
        ssl=False,
        ttl_dns_cache=300,
    )

    timeout = aiohttp.ClientTimeout(
        total=REQUEST_TIMEOUT
    )

    total_alerts = 0

    async with aiohttp.ClientSession(
        connector=connector,
        timeout=timeout,
        headers=HEADERS,
    ) as session:

        for index, row in boards_df.iterrows():

            try:

                alerts = await monitor_board(
                    session,
                    row,
                    seen,
                )

                total_alerts += alerts

                if (
                    (index + 1) % 20
                    == 0
                ):

                    print(
                        f"진행: "
                        f"{index + 1}/"
                        f"{len(boards_df)}"
                    )

            except Exception as e:

                print(
                    f"[모니터링 오류] "
                    f"{row.get('기관명', '')} / "
                    f"{row.get('게시판명', '')} / "
                    f"{type(e).__name__}: {e}"
                )

                continue

    save_seen(
        seen
    )

    print("=" * 70)

    print(
        "일일 모니터링 완료"
    )

    print("=" * 70)

    print(
        f"Telegram 알림 : "
        f"{total_alerts}건"
    )

    print(
        f"seen_posts    : "
        f"{len(seen)}건"
    )

    print("=" * 70)


# =========================================================
# 기관 데이터
# =========================================================

def load_institutions():

    if not os.path.exists(
        URL_FILE
    ):

        raise FileNotFoundError(
            f"{URL_FILE} 파일이 없습니다."
        )

    df = pd.read_excel(
        URL_FILE
    )

    df.columns = [
        clean_text(c)
        for c in df.columns
    ]

    required = [
        "기관명",
        "URL",
    ]

    for col in required:

        if col not in df.columns:

            raise ValueError(
                f"url_완성.xlsx에 "
                f"'{col}' 컬럼이 없습니다."
            )

    if "기관유형" not in df.columns:

        df["기관유형"] = ""

    df["URL"] = df[
        "URL"
    ].apply(
        normalize_url
    )

    df = df[
        df["기관명"]
        .astype(str)
        .str.strip()
        != ""
    ].copy()

    return df.reset_index(
        drop=True
    )


# =========================================================
# Main
# =========================================================

def main():

    global institutions

    institutions = (
        load_institutions()
    )

    print()

    print(
        f"기관 데이터 로드 완료: "
        f"{len(institutions)}개"
    )

    print(
        f"FORCE_DISCOVER_BOARDS = "
        f"{FORCE_DISCOVER_BOARDS}"
    )

    # -----------------------------------------------------
    # 강제 탐색
    # -----------------------------------------------------

    if FORCE_DISCOVER_BOARDS:

        print()

        print(
            "FORCE_DISCOVER_BOARDS=true"
        )

        print(
            "5차 게시판 탐색을 실행합니다."
        )

        asyncio.run(
            discover_all()
        )

        return

    # -----------------------------------------------------
    # boards.xlsx 없으면 최초 탐색
    # -----------------------------------------------------

    if not os.path.exists(
        BOARDS_FILE
    ):

        print()

        print(
            "boards.xlsx가 없으므로 "
            "5차 게시판 탐색을 실행합니다."
        )

        asyncio.run(
            discover_all()
        )

        return

    # -----------------------------------------------------
    # 정상 일일 모니터링
    # -----------------------------------------------------

    print()

    print(
        "기존 boards.xlsx를 사용하여 "
        "일일 모니터링을 실행합니다."
    )

    asyncio.run(
        monitor_all()
    )


if __name__ == "__main__":

    main()
