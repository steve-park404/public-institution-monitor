# -*- coding: utf-8 -*-

"""
355개 공공기관 홈페이지 게시판 모니터링 시스템
4차 버전

주요 기능
1. url_완성.xlsx에서 355개 기관 홈페이지 읽기
2. 게시판 최초 탐색
3. boards.xlsx 생성
4. boards_missing.xlsx 생성
5. 일상 실행 시 boards.xlsx의 게시판만 모니터링
6. 제목 + 본문에서 키워드 검색
7. Telegram 알림
8. seen_posts.json으로 중복 알림 방지

환경변수
FORCE_DISCOVER_BOARDS=true
    -> boards.xlsx가 있어도 게시판을 다시 탐색

FORCE_DISCOVER_BOARDS=false
    -> 기존 boards.xlsx를 이용해 게시글 모니터링
"""

import asyncio
import aiohttp
import pandas as pd
import os
import re
import json
import time
import html
from urllib.parse import (
    urljoin,
    urlparse,
    urlunparse,
    parse_qs,
)
from bs4 import BeautifulSoup
from collections import deque


# ============================================================
# 기본 설정
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

URL_FILE = os.path.join(BASE_DIR, "url_완성.xlsx")
BOARDS_FILE = os.path.join(BASE_DIR, "boards.xlsx")
MISSING_FILE = os.path.join(BASE_DIR, "boards_missing.xlsx")
SEEN_FILE = os.path.join(BASE_DIR, "seen_posts.json")


# ============================================================
# Telegram 설정
# ============================================================

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")


# ============================================================
# 실행 모드
# ============================================================

FORCE_DISCOVER_BOARDS = (
    os.getenv("FORCE_DISCOVER_BOARDS", "false").lower()
    in ("true", "1", "yes")
)


# ============================================================
# 탐색 설정
# ============================================================

CONCURRENCY = 12

# 기관 하나당 최대 탐색 페이지
MAX_PAGES_PER_ORG = 50

# 홈페이지에서 링크를 따라갈 최대 깊이
MAX_DEPTH = 4

# 기관 하나당 저장할 게시판 수
MAX_BOARDS_PER_ORG = 10

# HTTP timeout
REQUEST_TIMEOUT = 15

# 게시판으로 인정하기 위한 최소 점수
MIN_BOARD_SCORE = 5

# 게시글 상세페이지 탐색 최대 수
MAX_POSTS_PER_BOARD = 30


# ============================================================
# 검색 키워드
# ============================================================

KEYWORDS = [
    "설문조사",
    "시민참여",
    "국민참여",
    "공모전",
]


# ============================================================
# 게시판 관련 단어
# ============================================================

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


# ============================================================
# 사이트 구조 관련 단어
# ============================================================

SITE_MAP_WORDS = [
    "사이트맵",
    "사이트 맵",
    "전체메뉴",
    "전체 메뉴",
    "홈페이지맵",
    "홈페이지 맵",
    "메뉴",
]


# ============================================================
# URL에서 게시판 가능성이 높은 패턴
# ============================================================

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
    "contents",
    "content",
    "menu",
    "article",
    "program",
    "announce",
    "public",
    "customer",
    "communication",
    "opinion",
]


# ============================================================
# 상세 게시글 URL 패턴
# ============================================================

DETAIL_PATTERNS = [
    "view",
    "detail",
    "read",
    "articleview",
    "articleView",
    "nttId",
    "nttid",
    "bbsNo",
    "bbsno",
    "boardNo",
    "boardno",
    "seq",
    "idx",
    "wr_id",
    "postid",
    "newsid",
]


# ============================================================
# 무시할 확장자
# ============================================================

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


# ============================================================
# User-Agent
# ============================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/139.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,*/*;q=0.8"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}


# ============================================================
# 공통 함수
# ============================================================

def clean_text(value):
    """HTML/공백 제거"""
    if value is None:
        return ""

    value = html.unescape(str(value))
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_url(url):
    """URL 정규화"""
    if not url:
        return ""

    url = str(url).strip()

    if not url.startswith(("http://", "https://")):
        return ""

    parsed = urlparse(url)

    scheme = parsed.scheme.lower()
    netloc = parsed.netloc.lower()

    # fragment 제거
    path = parsed.path or "/"

    # // 중복 제거
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
    """같은 도메인인지 확인"""
    try:
        a = urlparse(url1).netloc.lower()
        b = urlparse(url2).netloc.lower()

        a = a.replace("www.", "")
        b = b.replace("www.", "")

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


def is_detail_url(url):
    """
    게시판 목록 URL인지 상세 게시글 URL인지 추정
    """

    lower = url.lower()

    # query parameter
    query = parse_qs(urlparse(url).query)

    # key=123 같은 CMS 방식은 목록일 수도 있으므로
    # 무조건 상세페이지로 보지 않는다.
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

    # 일반적인 상세 URL 패턴
    for pattern in DETAIL_PATTERNS:
        if pattern.lower() in lower:
            return True

    return False


def link_text_is_board(text):
    text = clean_text(text)

    if not text:
        return False

    return any(
        word in text
        for word in BOARD_WORDS
    )


def url_looks_like_board(url):
    lower = url.lower()

    return any(
        pattern in lower
        for pattern in BOARD_URL_PATTERNS
    )


# ============================================================
# HTML 구조 분석
# ============================================================

def extract_visible_text(soup):
    try:
        for tag in soup(
            ["script", "style", "noscript", "svg"]
        ):
            tag.decompose()

        return clean_text(
            soup.get_text(" ", strip=True)
        )
    except Exception:
        return ""


def detect_board_structure(soup):
    """
    HTML 자체가 게시판처럼 생겼는지 검사
    """

    score = 0
    reasons = []

    text = extract_visible_text(soup)

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

    if len(found_columns) >= 2:
        score += 4
        reasons.append(
            "게시판 항목: " + ",".join(found_columns[:5])
        )

    elif len(found_columns) == 1:
        score += 2
        reasons.append(
            "게시판 항목: " + found_columns[0]
        )

    # table
    tables = soup.find_all("table")

    if tables:
        score += 2
        reasons.append("table 존재")

    # 목록 형태
    if soup.find_all("ul"):
        score += 1
        reasons.append("ul 목록 존재")

    # paging
    page_words = [
        "다음",
        "이전",
        "페이지",
        "목록",
        "검색",
    ]

    page_found = [
        w for w in page_words
        if w in text
    ]

    if len(page_found) >= 2:
        score += 2
        reasons.append(
            "목록/페이지 요소"
        )

    # 글 수가 많은 링크
    links = soup.find_all("a")

    if len(links) >= 10:
        score += 1
        reasons.append("링크 다수")

    return score, reasons


# ============================================================
# 게시판 후보 점수
# ============================================================

def score_candidate(url, link_text, soup=None):
    """
    게시판 후보 점수 계산

    중요:
    모든 반환값은 반드시 score/reasons를 가진다.
    """

    score = 0
    reasons = []

    text = clean_text(link_text)
    lower_text = text.lower()
    lower_url = url.lower()

    # --------------------------------------------------------
    # 링크 텍스트
    # --------------------------------------------------------

    for word in BOARD_WORDS:
        if word.lower() in lower_text:
            score += 5
            reasons.append(
                f"메뉴명:{word}"
            )
            break

    # --------------------------------------------------------
    # URL
    # --------------------------------------------------------

    for pattern in BOARD_URL_PATTERNS:
        if pattern.lower() in lower_url:
            score += 3
            reasons.append(
                f"URL:{pattern}"
            )
            break

    # --------------------------------------------------------
    # 사이트맵 / 메뉴
    # --------------------------------------------------------

    for word in SITE_MAP_WORDS:
        if word.lower() in lower_text:
            score += 2
            reasons.append(
                f"메뉴구조:{word}"
            )
            break

    # --------------------------------------------------------
    # 상세 URL은 감점
    # --------------------------------------------------------

    if is_detail_url(url):
        score -= 5
        reasons.append("상세페이지 추정")

    # --------------------------------------------------------
    # HTML 구조
    # --------------------------------------------------------

    if soup is not None:
        structure_score, structure_reasons = (
            detect_board_structure(soup)
        )

        score += structure_score
        reasons.extend(structure_reasons)

    # --------------------------------------------------------
    # 최소 점수
    # --------------------------------------------------------

    return {
        "score": int(score),
        "reasons": "; ".join(
            dict.fromkeys(reasons)
        ),
    }


# ============================================================
# HTTP 요청
# ============================================================

async def fetch(session, url):
    """
    HTML 가져오기

    오류가 발생해도 전체 프로그램을 중단하지 않는다.
    """

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

            if status >= 400:
                return None, status, str(response.url)

            content_type = (
                response.headers.get(
                    "Content-Type", ""
                ).lower()
            )

            # HTML/XML이 아니면 무시
            if (
                "html" not in content_type
                and "xml" not in content_type
                and content_type
            ):
                return None, status, str(response.url)

            raw = await response.read()

            # 인코딩 자동 처리
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                try:
                    text = raw.decode("cp949")
                except UnicodeDecodeError:
                    text = raw.decode(
                        "euc-kr",
                        errors="ignore",
                    )

            return text, status, str(response.url)

    except asyncio.TimeoutError:
        return None, "timeout", url

    except Exception as e:
        return None, type(e).__name__, url


# ============================================================
# 링크 추출
# ============================================================

def extract_links(soup, base_url):
    """
    일반 링크 + iframe/frame 링크 추출
    """

    results = []

    # a
    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith(
            ("javascript:", "mailto:", "tel:", "#")
        ):
            continue

        full_url = urljoin(base_url, href)
        full_url = normalize_url(full_url)

        if not is_http_url(full_url):
            continue

        if is_ignored_file(full_url):
            continue

        text = clean_text(
            a.get_text(" ", strip=True)
        )

        results.append(
            (full_url, text)
        )

    # iframe / frame
    for tag in soup.find_all(
        ["iframe", "frame"]
    ):
        src = tag.get("src")

        if not src:
            continue

        full_url = normalize_url(
            urljoin(base_url, src)
        )

        if is_http_url(full_url):
            results.append(
                (full_url, "iframe")
            )

    # 중복 제거
    unique = []
    seen = set()

    for url, text in results:
        if url in seen:
            continue

        seen.add(url)
        unique.append(
            (url, text)
        )

    return unique


# ============================================================
# robots.txt
# ============================================================

async def discover_robots(session, homepage):
    """
    robots.txt에서 Sitemap 주소 탐색
    """

    parsed = urlparse(homepage)

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

        if line.lower().startswith("sitemap:"):
            sitemap = line.split(
                ":", 1
            )[1].strip()

            sitemap = normalize_url(sitemap)

            if sitemap:
                urls.append(sitemap)

    return urls


# ============================================================
# sitemap.xml
# ============================================================

async def parse_sitemap(session, sitemap_url):
    """
    sitemap XML에서 URL 추출
    """

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

    # sitemap index
    for loc in soup.find_all("loc"):

        value = clean_text(
            loc.get_text()
        )

        value = normalize_url(value)

        if value:
            urls.append(value)

    return urls


# ============================================================
# 공통 CMS URL 추정
# ============================================================

def common_paths(homepage):
    """
    홈페이지에서 자주 사용되는 메뉴/게시판 경로
    """

    parsed = urlparse(homepage)

    base = f"{parsed.scheme}://{parsed.netloc}"

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
        "/menu",
        "/sitemap",
        "/siteMap",
        "/main",
        "/index",
    ]

    return [
        normalize_url(
            urljoin(base, path)
        )
        for path in paths
    ]


# ============================================================
# 게시판 후보 저장
# ============================================================

def add_candidate(
    candidates,
    url,
    name,
    score,
    reasons,
):
    """
    후보를 안전하게 저장한다.

    이전 3차 버전의
    KeyError: 'score'
    방지 핵심 함수.
    """

    url = normalize_url(url)

    if not url:
        return

    # 상세 게시글은 원칙적으로 제외
    if is_detail_url(url):
        return

    existing = candidates.get(url)

    candidate = {
        "url": url,
        "name": clean_text(name)
        or "게시판",
        "score": int(score),
        "reasons": clean_text(reasons),
    }

    if existing is None:
        candidates[url] = candidate
        return

    # score가 없더라도 절대 오류가 나지 않도록 처리
    existing_score = int(
        existing.get("score", -999)
    )

    candidate_score = int(
        candidate.get("score", 0)
    )

    if candidate_score > existing_score:
        candidates[url] = candidate


# ============================================================
# 기관 하나 탐색
# ============================================================

async def discover_one(
    session,
    semaphore,
    index,
    total,
    row,
):
    """
    기관 하나의 홈페이지를 탐색한다.

    절대로 예외를 밖으로 던지지 않는다.
    """

    name = clean_text(
        row.get("기관명", "")
    )

    institution_type = clean_text(
        row.get("기관유형", "")
    )

    homepage = normalize_url(
        row.get("URL", "")
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
        result["status"] = "실패"
        result["note"] = "홈페이지 URL 없음"

        print(
            f"[{index:03d}/{total}] "
            f"{name} → 0개 "
            f"(홈페이지 URL 없음)"
        )

        return result

    async with semaphore:

        try:

            # ------------------------------------------------
            # 홈페이지 접속
            # ------------------------------------------------

            text, status, final_url = (
                await fetch(
                    session,
                    homepage,
                )
            )

            if not text:

                result["status"] = "실패"

                if status == "timeout":
                    result["note"] = (
                        "홈페이지 접속 시간 초과"
                    )
                else:
                    result["note"] = (
                        f"홈페이지 접속 실패: {status}"
                    )

                print(
                    f"[{index:03d}/{total}] "
                    f"{name} → 0개 "
                    f"({result['note']})"
                )

                return result

            homepage = (
                normalize_url(final_url)
                or homepage
            )

            # ------------------------------------------------
            # BFS
            # ------------------------------------------------

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

            # robots sitemap
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
                            "robots sitemap",
                        )
                    )

            except Exception:
                pass

            # common paths
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
            # BFS 탐색
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

                visited.add(current_url)

                if not is_http_url(current_url):
                    continue

                if is_ignored_file(current_url):
                    continue

                # 외부 도메인 제외
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

                # ------------------------------------------------
                # 현재 페이지 자체가 게시판인지 평가
                # ------------------------------------------------

                current_title = ""

                title_tag = soup.find("title")

                if title_tag:
                    current_title = clean_text(
                        title_tag.get_text()
                    )

                current_score_data = (
                    score_candidate(
                        current_url,
                        current_title,
                        soup,
                    )
                )

                if (
                    current_score_data["score"]
                    >= MIN_BOARD_SCORE
                ):
                    add_candidate(
                        candidates,
                        current_url,
                        current_title
                        or "게시판",
                        current_score_data[
                            "score"
                        ],
                        current_score_data[
                            "reasons"
                        ],
                    )

                # ------------------------------------------------
                # 링크 분석
                # ------------------------------------------------

                links = extract_links(
                    soup,
                    current_url,
                )

                for link_url, link_text in links:

                    if not same_domain(
                        link_url,
                        homepage,
                    ):
                        continue

                    # 상세 게시글 URL은
                    # 게시판 후보로 저장하지 않는다.
                    detail = is_detail_url(
                        link_url
                    )

                    data = score_candidate(
                        link_url,
                        link_text,
                        None,
                    )

                    # 현재 페이지의 HTML 구조 점수는
                    # 링크 자체에는 적용하지 않는다.

                    if (
                        not detail
                        and data["score"]
                        >= MIN_BOARD_SCORE
                    ):
                        add_candidate(
                            candidates,
                            link_url,
                            link_text
                            or "게시판",
                            data["score"],
                            data["reasons"],
                        )

                    # ------------------------------------------------
                    # 다음 페이지 탐색 여부
                    # ------------------------------------------------

                    if depth >= MAX_DEPTH:
                        continue

                    if detail:
                        continue

                    # 탐색 가치가 높은 링크
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

                    # 1단계는 폭넓게
                    if depth == 0:
                        should_follow = True

                    else:
                        should_follow = (
                            menu_match
                            or url_match
                            or site_map_match
                        )

                    if should_follow:

                        if link_url not in visited:
                            queue.append(
                                (
                                    link_url,
                                    depth + 1,
                                    "링크",
                                )
                            )

                # ------------------------------------------------
                # iframe 내부
                # ------------------------------------------------

                for frame in soup.find_all(
                    ["iframe", "frame"]
                ):

                    src = frame.get("src")

                    if not src:
                        continue

                    frame_url = normalize_url(
                        urljoin(
                            current_url,
                            src,
                        )
                    )

                    if (
                        frame_url
                        and same_domain(
                            frame_url,
                            homepage,
                        )
                        and frame_url
                        not in visited
                    ):

                        queue.append(
                            (
                                frame_url,
                                depth + 1,
                                "iframe",
                            )
                        )

            # ------------------------------------------------
            # 후보를 점수순 정렬
            # ------------------------------------------------

            sorted_candidates = sorted(
                candidates.values(),
                key=lambda x: (
                    int(
                        x.get(
                            "score",
                            0,
                        )
                    ),
                    len(
                        x.get(
                            "reasons",
                            "",
                        )
                    ),
                ),
                reverse=True,
            )

            # ------------------------------------------------
            # 너무 낮은 후보 제거
            # ------------------------------------------------

            final_candidates = []

            seen_board_urls = set()

            for candidate in sorted_candidates:

                board_url = normalize_url(
                    candidate.get("url", "")
                )

                if not board_url:
                    continue

                if board_url in seen_board_urls:
                    continue

                # 상세페이지 방지
                if is_detail_url(
                    board_url
                ):
                    continue

                score = int(
                    candidate.get(
                        "score",
                        0,
                    )
                )

                if score < MIN_BOARD_SCORE:
                    continue

                seen_board_urls.add(
                    board_url
                )

                final_candidates.append(
                    candidate
                )

                if len(final_candidates) >= (
                    MAX_BOARDS_PER_ORG
                ):
                    break

            result["boards"] = (
                final_candidates
            )

            if final_candidates:

                result["status"] = "정상"

                result["note"] = (
                    f"게시판 후보 "
                    f"{len(final_candidates)}개"
                )

                print(
                    f"[{index:03d}/{total}] "
                    f"{name} → "
                    f"{len(final_candidates)}개 "
                    f"(탐색 {result['pages']}페이지)"
                )

            else:

                result["status"] = "실패"

                result["note"] = (
                    "게시판 후보 없음"
                )

                print(
                    f"[{index:03d}/{total}] "
                    f"{name} → 0개 "
                    f"(게시판 후보 없음, "
                    f"탐색 {result['pages']}페이지)"
                )

            return result

        except Exception as e:

            # ★ 가장 중요
            # 기관 하나의 오류가 전체 탐색을
            # 중단시키지 않도록 한다.

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


# ============================================================
# 전체 기관 탐색
# ============================================================

async def discover_all():

    print("=" * 70)
    print("4차 게시판 탐색 시작")
    print("=" * 70)

    print(
        f"전체 기관 : {len(institutions)}개"
    )

    print(
        f"동시 탐색 : {CONCURRENCY}개"
    )

    print(
        f"기관당 최대 페이지 : "
        f"{MAX_PAGES_PER_ORG}"
    )

    print(
        f"최대 탐색 깊이 : {MAX_DEPTH}"
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

        async def worker(index, row):

            return await discover_one(
                session,
                semaphore,
                index,
                len(institutions),
                row,
            )

        tasks = []

        for i, row in institutions.iterrows():

            tasks.append(
                worker(
                    i + 1,
                    row,
                )
            )

        # gather 자체의 예외 방지를 위해
        # return_exceptions=True 사용
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

    # ========================================================
    # boards.xlsx
    # ========================================================

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

    # 중복 제거
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

    # ========================================================
    # boards_missing.xlsx
    # ========================================================

    missing_rows = []

    for result in results:

        if not result.get("boards"):

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

    # ========================================================
    # 결과 통계
    # ========================================================

    total_orgs = len(institutions)

    found_orgs = sum(
        1
        for result in results
        if result.get("boards")
    )

    missing_orgs = (
        total_orgs - found_orgs
    )

    coverage = (
        found_orgs / total_orgs * 100
        if total_orgs
        else 0
    )

    print()
    print("=" * 70)
    print("4차 게시판 탐색 완료")
    print("=" * 70)

    print(
        f"전체 기관       : {total_orgs}개"
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


# ============================================================
# seen_posts.json
# ============================================================

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

        if isinstance(data, list):
            return set(data)

        if isinstance(data, dict):
            return set(data.keys())

    except Exception as e:

        print(
            f"seen_posts.json 읽기 오류: {e}"
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
                sorted(list(seen)),
                f,
                ensure_ascii=False,
                indent=2,
            )

    except Exception as e:

        print(
            f"seen_posts.json 저장 오류: {e}"
        )


# ============================================================
# 게시글 상세 페이지 분석
# ============================================================

def extract_post_info(
    soup,
    url,
):
    """
    상세페이지에서 제목 + 본문 추출
    """

    title = ""

    # title
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

            if 2 <= len(value) <= 300:
                title = value
                break

    if not title:

        tag = soup.find("title")

        if tag:
            title = clean_text(
                tag.get_text()
            )

    # 본문
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

            if len(value) > len(content):
                content = value

    if not content:
        content = extract_visible_text(
            soup
        )

    return {
        "url": url,
        "title": title,
        "content": content,
    }


# ============================================================
# 게시판에서 게시글 링크 추출
# ============================================================

def extract_post_links(
    soup,
    board_url,
):
    """
    게시판 목록에서 상세 게시글 링크 추출
    """

    links = []

    # table 내부 링크 우선
    tables = soup.find_all("table")

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

                # 빈 링크 제외
                if not text:
                    continue

                # 숫자만 있는 링크도 허용
                links.append(
                    (
                        full_url,
                        text,
                    )
                )

    # table이 없으면 일반 링크
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

            # 상세페이지 가능성이 있는 것
            if is_detail_url(
                full_url
            ):
                links.append(
                    (
                        full_url,
                        text,
                    )
                )

    # 중복 제거
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

        if len(result) >= MAX_POSTS_PER_BOARD:
            break

    return result


# ============================================================
# 키워드 검색
# ============================================================

def find_keywords(
    title,
    content,
):
    combined = (
        clean_text(title)
        + " "
        + clean_text(content)
    )

    found = []

    for keyword in KEYWORDS:

        if keyword in combined:

            found.append(
                keyword
            )

    return found


# ============================================================
# Telegram
# ============================================================

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


# ============================================================
# 게시판 모니터링
# ============================================================

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

    post_links = extract_post_links(
        soup,
        board_url,
    )

    alerts = 0

    for post_url, link_text in post_links:

        # 이미 본 게시글은 건너뜀
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

        info = extract_post_info(
            post_soup,
            post_url,
        )

        title = (
            info["title"]
            or link_text
        )

        content = info["content"]

        found = find_keywords(
            title,
            content,
        )

        # 신규 게시글이라는 것은
        # 일단 seen에 기록
        seen.add(post_url)

        if not found:
            continue

        # ----------------------------------------------------
        # Telegram 메시지
        # ----------------------------------------------------

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


# ============================================================
# 전체 모니터링
# ============================================================

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
    print("일일 게시판 모니터링 시작")
    print("=" * 70)

    print(
        f"게시판 수 : {len(boards_df)}개"
    )

    print(
        f"기존 seen : {len(seen)}개"
    )

    print(
        f"검색 키워드 : {', '.join(KEYWORDS)}"
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
                    (index + 1) % 20 == 0
                ):
                    print(
                        f"진행: "
                        f"{index + 1}/"
                        f"{len(boards_df)}"
                    )

            except Exception as e:

                print(
                    f"[모니터링 오류] "
                    f"{row.get('기관명', '')} "
                    f"/ "
                    f"{row.get('게시판명', '')} "
                    f"/ "
                    f"{type(e).__name__}: {e}"
                )

                # 한 게시판 오류로
                # 전체 모니터링이 중단되지 않음
                continue

    save_seen(seen)

    print("=" * 70)
    print("일일 모니터링 완료")
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


# ============================================================
# URL 파일 읽기
# ============================================================

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

    # 컬럼명 공백 제거
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

    # 기관유형 없으면 생성
    if "기관유형" not in df.columns:
        df["기관유형"] = ""

    # URL 정리
    df["URL"] = df["URL"].apply(
        normalize_url
    )

    # 기관명 없는 행 제거
    df = df[
        df["기관명"]
        .astype(str)
        .str.strip()
        != ""
    ].copy()

    return df.reset_index(
        drop=True
    )


# ============================================================
# MAIN
# ============================================================

def main():

    global institutions

    institutions = load_institutions()

    print()
    print(
        f"기관 데이터 로드 완료: "
        f"{len(institutions)}개"
    )

    print(
        f"FORCE_DISCOVER_BOARDS = "
        f"{FORCE_DISCOVER_BOARDS}"
    )

    # --------------------------------------------------------
    # 강제 게시판 탐색
    # --------------------------------------------------------

    if FORCE_DISCOVER_BOARDS:

        print()
        print(
            "FORCE_DISCOVER_BOARDS=true"
        )

        print(
            "boards.xlsx 존재 여부와 관계없이 "
            "게시판을 다시 탐색합니다."
        )

        asyncio.run(
            discover_all()
        )

        return

    # --------------------------------------------------------
    # boards.xlsx가 없으면 최초 탐색
    # --------------------------------------------------------

    if not os.path.exists(
        BOARDS_FILE
    ):

        print()
        print(
            "boards.xlsx가 없으므로 "
            "게시판 최초 탐색을 실행합니다."
        )

        asyncio.run(
            discover_all()
        )

        return

    # --------------------------------------------------------
    # 정상 일일 모니터링
    # --------------------------------------------------------

    print()
    print(
        "기존 boards.xlsx를 사용하여 "
        "일일 모니터링을 실행합니다."
    )

    asyncio.run(
        monitor_all()
    )


# ============================================================
# 실행
# ============================================================

if __name__ == "__main__":
    main()
