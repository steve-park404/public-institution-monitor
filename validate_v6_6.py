import asyncio
import aiohttp
import pandas as pd
import re
import os

from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs


# =========================================================
# V6.6 설정
# =========================================================

INPUT_FILE = "boards_v6_5_validation.xlsx"
OUTPUT_FILE = "boards_v6_6_validation.xlsx"

TIMEOUT = aiohttp.ClientTimeout(total=40)

CONCURRENCY = 5

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    )
}


# =========================================================
# URL 분류 기준
# =========================================================

HOMEPAGE_PATHS = {
    "",
    "/",
    "/main",
    "/main/",
    "/main/index.do",
    "/home",
    "/home/",
    "/home/index.do",
    "/index.do",
}


DETAIL_QUERY_KEYS = {
    "act",
    "articleNo",
    "nttId",
    "list_no",
    "boardId",
    "bbsIdx",
    "idx",
    "seq",
}


DETAIL_PATH_WORDS = [
    "/view",
    "/detail",
    "boardview",
    "selectnewsarticle",
]


LIST_PATTERNS = [
    "/list",
    "list.do",
    "list.jsp",
    "list.asp",
    "/bbs",
    "/notice",
    "/board",
    "boardlist",
    "postlist",
    "article/list",
]


LIST_TEXT = [
    "번호",
    "제목",
    "등록일",
    "작성일",
    "조회",
    "공지사항",
    "채용",
    "자료실",
    "소식",
    "뉴스",
]


LIST_BUTTON_TEXT = [
    "목록",
    "리스트",
    "전체목록",
    "공지사항",
    "공지",
    "게시판",
]


DATE_PATTERNS = [
    r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}",
    r"\d{4}\.\d{1,2}\.\d{1,2}",
    r"\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}",
]


# =========================================================
# 기본 함수
# =========================================================

def normalize_url(url):

    if not url:
        return ""

    url = str(url).strip()

    if not url:
        return ""

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def same_host(a, b):

    try:

        h1 = (
            urlparse(a)
            .netloc
            .lower()
            .replace("www.", "")
        )

        h2 = (
            urlparse(b)
            .netloc
            .lower()
            .replace("www.", "")
        )

        return h1 == h2

    except Exception:

        return False


def normalize_compare_url(url):

    try:

        p = urlparse(url)

        path = p.path.rstrip("/")

        return (
            p.scheme.lower(),
            p.netloc.lower().replace("www.", ""),
            path.lower(),
            p.query.lower(),
        )

    except Exception:

        return ("", "", "", "")


def same_url(a, b):

    return normalize_compare_url(a) == normalize_compare_url(b)


def clean_text(text):

    if not text:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(text)
    ).strip()


# =========================================================
# URL 분류
# =========================================================

def classify_url(url):

    try:

        p = urlparse(url)

        path = p.path.lower()

        query = parse_qs(p.query)

        if path in HOMEPAGE_PATHS:

            return "홈페이지"

        query_keys = {
            x.lower()
            for x in DETAIL_QUERY_KEYS
        }

        for key in query.keys():

            if key.lower() in query_keys:

                return "상세"

        for values in query.values():

            for value in values:

                if str(value).lower() == "view":

                    return "상세"

        for word in DETAIL_PATH_WORDS:

            if word in path:

                return "상세"

        for pattern in LIST_PATTERNS:

            if pattern in path:

                return "목록"

        return "미분류"

    except Exception:

        return "오류"


# =========================================================
# HTTP
# =========================================================

async def fetch(session, url):

    try:

        async with session.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
            allow_redirects=True,
            ssl=False
        ) as response:

            text = await response.text(
                errors="ignore"
            )

            return {
                "status": response.status,
                "url": str(response.url),
                "text": text,
            }

    except Exception as e:

        return {
            "status": 0,
            "url": url,
            "text": "",
            "error": str(e),
        }


# =========================================================
# 제목
# =========================================================

def extract_title(soup):

    selectors = [
        "h1",
        "h2",
        "h3",
        ".subject",
        ".title",
        ".board-title",
        ".view-title",
        ".bbs-title",
        ".tit",
        "title",
    ]

    for selector in selectors:

        try:

            el = soup.select_one(
                selector
            )

            if el:

                text = clean_text(
                    el.get_text(
                        " ",
                        strip=True
                    )
                )

                if 3 <= len(text) <= 300:

                    return text

        except Exception:

            pass

    return ""


# =========================================================
# 날짜
# =========================================================

def extract_date(text):

    for pattern in DATE_PATTERNS:

        match = re.search(
            pattern,
            text
        )

        if match:

            return match.group(0)

    return ""


# =========================================================
# 게시물 제목으로 보이는 링크인지 판단
# =========================================================

def looks_like_post_title(text):

    text = clean_text(text)

    if len(text) < 2:
        return False

    if len(text) > 250:
        return False

    # 메뉴성 텍스트 제거
    menu_words = [
        "로그인",
        "회원가입",
        "사이트맵",
        "홈",
        "HOME",
        "메인",
        "검색",
        "전체메뉴",
        "메뉴",
        "이전",
        "다음",
        "목록",
        "리스트",
        "맨위로",
        "TOP",
        "닫기",
        "확인",
        "취소",
    ]

    for word in menu_words:

        if text.upper() == word.upper():

            return False

    return True


# =========================================================
# 게시물 링크 추출
# =========================================================

def extract_post_links(base_url, soup):

    candidates = []

    # -----------------------------------------------------
    # table / tr 구조 우선
    # -----------------------------------------------------

    rows = soup.find_all("tr")

    for row in rows:

        links = row.find_all(
            "a",
            href=True
        )

        if not links:
            continue

        row_text = clean_text(
            row.get_text(
                " ",
                strip=True
            )
        )

        has_date = bool(
            re.search(
                r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}",
                row_text
            )
        )

        for a in links:

            href = a.get(
                "href",
                ""
            ).strip()

            text = clean_text(
                a.get_text(
                    " ",
                    strip=True
                )
            )

            if not href:
                continue

            if href.startswith(
                (
                    "javascript:",
                    "#",
                    "mailto:",
                )
            ):
                continue

            if not looks_like_post_title(text):
                continue

            absolute = urljoin(
                base_url,
                href
            )

            if not same_host(
                base_url,
                absolute
            ):
                continue

            kind = classify_url(
                absolute
            )

            score = 0

            if kind == "상세":

                score += 6

            elif kind == "목록":

                score += 1

            if has_date:

                score += 2

            lower = absolute.lower()

            if any(
                x in lower
                for x in [
                    "article",
                    "nttid",
                    "articleid",
                    "boardid",
                    "bbsidx",
                    "article_no",
                    "idx=",
                    "seq=",
                    "no=",
                ]
            ):

                score += 3

            if 5 <= len(text) <= 200:

                score += 1

            if score >= 5:

                candidates.append(
                    (
                        absolute,
                        text,
                        score
                    )
                )

    # -----------------------------------------------------
    # table이 아닌 일반 리스트 구조
    # -----------------------------------------------------

    for container in soup.find_all(
        ["li", "div"]
    ):

        links = container.find_all(
            "a",
            href=True
        )

        if not links:
            continue

        container_text = clean_text(
            container.get_text(
                " ",
                strip=True
            )
        )

        has_date = bool(
            re.search(
                r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}",
                container_text
            )
        )

        # 너무 큰 div는 메뉴/전체 페이지일 가능성이 높음
        if len(container_text) > 500:
            continue

        for a in links:

            href = a.get(
                "href",
                ""
            ).strip()

            text = clean_text(
                a.get_text(
                    " ",
                    strip=True
                )
            )

            if not href:
                continue

            if not looks_like_post_title(text):
                continue

            absolute = urljoin(
                base_url,
                href
            )

            if not same_host(
                base_url,
                absolute
            ):
                continue

            kind = classify_url(
                absolute
            )

            score = 0

            if kind == "상세":

                score += 6

            if has_date:

                score += 2

            lower = absolute.lower()

            if any(
                x in lower
                for x in [
                    "article",
                    "nttid",
                    "articleid",
                    "boardid",
                    "bbsidx",
                    "article_no",
                    "idx=",
                    "seq=",
                    "no=",
                ]
            ):

                score += 3

            if 5 <= len(text) <= 200:

                score += 1

            if score >= 5:

                candidates.append(
                    (
                        absolute,
                        text,
                        score
                    )
                )

    # -----------------------------------------------------
    # URL 중복 제거
    # -----------------------------------------------------

    unique = {}

    for url, text, score in candidates:

        key = normalize_compare_url(
            url
        )

        if key not in unique:

            unique[key] = (
                url,
                text,
                score
            )

        else:

            old = unique[key]

            if score > old[2]:

                unique[key] = (
                    url,
                    text,
                    score
                )

    result = list(
        unique.values()
    )

    result.sort(
        key=lambda x: x[2],
        reverse=True
    )

    return result


# =========================================================
# 게시판 구조 분석
# =========================================================

def inspect_board_structure(
    list_url,
    soup
):

    text = clean_text(
        soup.get_text(
            " ",
            strip=True
        )
    )

    lower_text = text.lower()

    score = 0

    evidence = []

    # -----------------------------------------------------
    # 1. 목록 키워드
    # -----------------------------------------------------

    keyword_count = 0

    for keyword in LIST_TEXT:

        if keyword in text:

            keyword_count += 1

    if keyword_count >= 5:

        score += 4
        evidence.append(
            "목록 키워드 5개 이상"
        )

    elif keyword_count >= 3:

        score += 3
        evidence.append(
            "목록 키워드 3개 이상"
        )

    elif keyword_count >= 2:

        score += 1
        evidence.append(
            "목록 키워드 2개 이상"
        )

    # -----------------------------------------------------
    # 2. TABLE 구조
    # -----------------------------------------------------

    tables = soup.find_all("table")

    if tables:

        score += 3

        evidence.append(
            f"TABLE {len(tables)}개"
        )

    # -----------------------------------------------------
    # 3. 목록 태그
    # -----------------------------------------------------

    list_items = soup.find_all("li")

    if len(list_items) >= 5:

        score += 1

        evidence.append(
            f"LI {len(list_items)}개"
        )

    # -----------------------------------------------------
    # 4. HTML 내 게시판 관련 단어
    # -----------------------------------------------------

    html = str(
        soup
    ).lower()

    board_words = [
        "board",
        "bbs",
        "notice",
        "post",
        "article",
    ]

    board_word_count = sum(
        html.count(word)
        for word in board_words
    )

    if board_word_count >= 3:

        score += 2

        evidence.append(
            "게시판 관련 HTML 구조 확인"
        )

    elif board_word_count >= 1:

        score += 1

    # -----------------------------------------------------
    # 5. 페이지네이션
    # -----------------------------------------------------

    pagination = detect_pagination(
        soup
    )

    if pagination:

        score += 2

        evidence.append(
            "페이지네이션 확인"
        )

    # -----------------------------------------------------
    # 6. 게시물 링크
    # -----------------------------------------------------

    post_links = extract_post_links(
        list_url,
        soup
    )

    if len(post_links) >= 10:

        score += 4

        evidence.append(
            f"게시물 링크 {len(post_links)}개"
        )

    elif len(post_links) >= 5:

        score += 3

        evidence.append(
            f"게시물 링크 {len(post_links)}개"
        )

    elif len(post_links) >= 3:

        score += 2

        evidence.append(
            f"게시물 링크 {len(post_links)}개"
        )

    elif len(post_links) >= 1:

        score += 1

    return {
        "score": score,
        "keyword_count": keyword_count,
        "post_links": post_links,
        "pagination": pagination,
        "evidence": evidence,
        "table_count": len(tables),
        "li_count": len(list_items),
    }


# =========================================================
# 페이지네이션
# =========================================================

def detect_pagination(soup):

    number_count = 0

    for a in soup.find_all(
        "a",
        href=True
    ):

        text = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

        if re.fullmatch(
            r"\d{1,3}",
            text
        ):

            number_count += 1

    if number_count >= 2:

        return True

    # 실제 페이지 이동 링크 확인
    for a in soup.find_all(
        "a",
        href=True
    ):

        text = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

        href = a.get(
            "href",
            ""
        ).lower()

        if any(
            word in text
            for word in [
                "다음",
                "이전",
                "첫 페이지",
                "마지막 페이지",
            ]
        ):

            return True

        if any(
            word in href
            for word in [
                "page=",
                "pageNo=",
                "pageno=",
                "currentpage=",
            ]
        ):

            return True

    return False


# =========================================================
# 실제 게시물 검증
# =========================================================

async def verify_posts(
    session,
    post_links
):

    checked = []

    seen_titles = set()

    # 최대 10개만 실제 접근
    for (
        post_url,
        link_text,
        link_score
    ) in post_links[:10]:

        result = await fetch(
            session,
            post_url
        )

        if (
            result["status"] < 200
            or result["status"] >= 400
            or not result["text"]
        ):

            continue

        soup = BeautifulSoup(
            result["text"],
            "html.parser"
        )

        final_url = result["url"]

        final_kind = classify_url(
            final_url
        )

        # 상세페이지가 아닌 홈페이지로 이동하면 제외
        if final_kind == "홈페이지":

            continue

        title = extract_title(
            soup
        )

        body = clean_text(
            soup.get_text(
                " ",
                strip=True
            )
        )

        date = extract_date(
            body
        )

        if not title:

            continue

        if len(body) < 80:

            continue

        title_key = clean_text(
            title
        ).lower()

        if title_key in seen_titles:

            continue

        seen_titles.add(
            title_key
        )

        checked.append(
            {
                "url": final_url,
                "title": title,
                "date": date,
                "body_length": len(body),
            }
        )

    return checked


# =========================================================
# 상세페이지 → 목록 복구
# =========================================================

async def recover_list_from_detail(
    session,
    detail_url
):

    result = await fetch(
        session,
        detail_url
    )

    if (
        result["status"] < 200
        or result["status"] >= 400
        or not result["text"]
    ):

        return []

    soup = BeautifulSoup(
        result["text"],
        "html.parser"
    )

    candidates = []

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = a.get(
            "href",
            ""
        ).strip()

        if not href:

            continue

        if href.startswith(
            (
                "javascript:",
                "#",
            )
        ):

            continue

        absolute = urljoin(
            result["url"],
            href
        )

        if not same_host(
            detail_url,
            absolute
        ):

            continue

        if same_url(
            detail_url,
            absolute
        ):

            continue

        text = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

        kind = classify_url(
            absolute
        )

        score = 0

        # "목록" 링크
        for word in LIST_BUTTON_TEXT:

            if word in text:

                score += 10

        # URL 자체가 목록 형태
        if kind == "목록":

            score += 8

        # 상세 URL이면 강하게 감점
        if kind == "상세":

            score -= 10

        # href에 list 관련 단어
        lower = absolute.lower()

        for word in [
            "list",
            "bbs",
            "board",
            "notice",
            "postlist",
        ]:

            if word in lower:

                score += 2

        if score >= 5:

            candidates.append(
                (
                    score,
                    absolute
                )
            )

    # 중복 제거
    unique = {}

    for score, url in candidates:

        key = normalize_compare_url(
            url
        )

        if key not in unique:

            unique[key] = (
                score,
                url
            )

        elif score > unique[key][0]:

            unique[key] = (
                score,
                url
            )

    candidates = list(
        unique.values()
    )

    candidates.sort(
        reverse=True
    )

    return [
        url
        for score, url
        in candidates[:15]
    ]


# =========================================================
# 홈페이지 → 목록 후보 탐색
# =========================================================

async def discover_from_homepage(
    session,
    homepage
):

    result = await fetch(
        session,
        homepage
    )

    if (
        result["status"] < 200
        or result["status"] >= 400
        or not result["text"]
    ):

        return []

    soup = BeautifulSoup(
        result["text"],
        "html.parser"
    )

    candidates = []

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = a.get(
            "href",
            ""
        ).strip()

        if not href:

            continue

        if href.startswith(
            (
                "javascript:",
                "#",
                "mailto:",
            )
        ):

            continue

        absolute = urljoin(
            result["url"],
            href
        )

        if not same_host(
            homepage,
            absolute
        ):

            continue

        kind = classify_url(
            absolute
        )

        if kind in [
            "홈페이지",
            "상세",
        ]:

            continue

        text = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

        score = 0

        for word in LIST_BUTTON_TEXT:

            if word in text:

                score += 7

        if kind == "목록":

            score += 8

        lower = absolute.lower()

        for pattern in LIST_PATTERNS:

            if pattern in lower:

                score += 3

        if score >= 5:

            candidates.append(
                (
                    score,
                    absolute
                )
            )

    # 중복 제거
    unique = {}

    for score, url in candidates:

        key = normalize_compare_url(
            url
        )

        if key not in unique:

            unique[key] = (
                score,
                url
            )

        elif score > unique[key][0]:

            unique[key] = (
                score,
                url
            )

    candidates = list(
        unique.values()
    )

    candidates.sort(
        reverse=True
    )

    return [
        url
        for score, url
        in candidates[:20]
    ]


# =========================================================
# 핵심 목록 URL 검증
# =========================================================

async def verify_list_url(
    session,
    list_url
):

    result = await fetch(
        session,
        list_url
    )

    if (
        result["status"] < 200
        or result["status"] >= 400
        or not result["text"]
    ):

        return {
            "access": False,
            "auto_confirm": False,
            "reason": "페이지 접속 실패",
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": "",
            "post": None,
            "structure": None,
        }

    final_url = result["url"]

    final_kind = classify_url(
        final_url
    )

    # -----------------------------------------------------
    # 핵심: 홈페이지는 무조건 탈락
    # -----------------------------------------------------

    if final_kind == "홈페이지":

        return {
            "access": True,
            "auto_confirm": False,
            "reason": "홈페이지로 연결됨",
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": final_url,
            "post": None,
            "structure": None,
        }

    # -----------------------------------------------------
    # 상세페이지는 최종 목록으로 인정하지 않음
    # -----------------------------------------------------

    if final_kind == "상세":

        return {
            "access": True,
            "auto_confirm": False,
            "reason": "상세페이지 URL",
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": final_url,
            "post": None,
            "structure": None,
        }

    soup = BeautifulSoup(
        result["text"],
        "html.parser"
    )

    structure = inspect_board_structure(
        final_url,
        soup
    )

    post_links = structure[
        "post_links"
    ]

    verified_posts = await verify_posts(
        session,
        post_links
    )

    unique_titles = {
        clean_text(
            x["title"]
        ).lower()
        for x in verified_posts
    }

    verified_count = len(
        unique_titles
    )

    score = structure[
        "score"
    ]

    # 실제 게시물 검증 점수
    if verified_count >= 5:

        score += 7

    elif verified_count >= 4:

        score += 6

    elif verified_count >= 3:

        score += 5

    elif verified_count >= 2:

        score += 2

    elif verified_count >= 1:

        score += 1

    # URL 자체가 목록 패턴이면 보너스
    if final_kind == "목록":

        score += 2

    # -----------------------------------------------------
    # V6.6 자동확정 기준
    # -----------------------------------------------------
    #
    # 1. 최종 URL이 실제 목록 형태
    # 2. 게시물 링크 최소 3개
    # 3. 실제 서로 다른 게시물 최소 3개
    # 4. 목록 키워드 최소 2개
    # 5. 구조 점수 최소 8
    # 6. 실제 게시물 3개 이상이면 강한 증거
    #
    # -----------------------------------------------------

    auto_confirm = (
        final_kind == "목록"
        and len(post_links) >= 3
        and verified_count >= 3
        and structure["keyword_count"] >= 2
        and structure["score"] >= 8
    )

    reasons = []

    if final_kind != "목록":

        reasons.append(
            "최종 URL이 목록형 URL이 아님"
        )

    if len(post_links) < 3:

        reasons.append(
            f"게시물 링크 부족({len(post_links)}개)"
        )

    if verified_count < 3:

        reasons.append(
            f"실제 게시물 부족({verified_count}개)"
        )

    if structure["keyword_count"] < 2:

        reasons.append(
            f"목록 키워드 부족({structure['keyword_count']}개)"
        )

    if structure["score"] < 8:

        reasons.append(
            f"목록 구조 점수 부족({structure['score']}점)"
        )

    if auto_confirm:

        reason = (
            "실제 게시판 목록 구조 확인 / "
            f"게시물 링크 {len(post_links)}개 / "
            f"서로 다른 실제 게시물 {verified_count}개 / "
            f"구조점수 {structure['score']} / "
            f"총점 {score}"
        )

    else:

        if not reasons:

            reasons.append(
                "자동확정 기준 미충족"
            )

        reason = ", ".join(
            reasons
        )

    return {
        "access": True,
        "auto_confirm": auto_confirm,
        "reason": reason,
        "score": score,
        "post_count": len(post_links),
        "verified_post_count": verified_count,
        "pagination": structure[
            "pagination"
        ],
        "final_url": final_url,
        "post": (
            verified_posts[0]
            if verified_posts
            else None
        ),
        "structure": structure,
    }


# =========================================================
# 기관별 후보 처리
# =========================================================

async def process_candidate(
    session,
    semaphore,
    institution,
    candidate_url
):

    async with semaphore:

        candidate_type = classify_url(
            candidate_url
        )

        final_list_url = ""

        best_result = {
            "access": False,
            "auto_confirm": False,
            "reason": "",
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": "",
            "post": None,
            "structure": None,
        }

        error = ""

        try:

            # =================================================
            # 1. 이미 목록형 URL이면 직접 검증
            # =================================================

            if candidate_type == "목록":

                result = await verify_list_url(
                    session,
                    candidate_url
                )

                best_result = result

                if result["auto_confirm"]:

                    final_list_url = result[
                        "final_url"
                    ]

            # =================================================
            # 2. 상세 URL이면 목록 URL 복구
            # =================================================

            elif candidate_type == "상세":

                list_candidates = (
                    await recover_list_from_detail(
                        session,
                        candidate_url
                    )
                )

                for list_url in list_candidates:

                    result = await verify_list_url(
                        session,
                        list_url
                    )

                    if (
                        result["score"]
                        >
                        best_result["score"]
                    ):

                        best_result = result

                    if result["auto_confirm"]:

                        final_list_url = result[
                            "final_url"
                        ]

                        break

                if not list_candidates:

                    best_result[
                        "reason"
                    ] = (
                        "상세페이지에서 목록 후보를 "
                        "찾지 못함"
                    )

            # =================================================
            # 3. 홈페이지면 게시판 후보 탐색
            # =================================================

            elif candidate_type == "홈페이지":

                list_candidates = (
                    await discover_from_homepage(
                        session,
                        candidate_url
                    )
                )

                for list_url in list_candidates:

                    result = await verify_list_url(
                        session,
                        list_url
                    )

                    if (
                        result["score"]
                        >
                        best_result["score"]
                    ):

                        best_result = result

                    if result["auto_confirm"]:

                        final_list_url = result[
                            "final_url"
                        ]

                        break

                if not list_candidates:

                    best_result[
                        "reason"
                    ] = (
                        "홈페이지에서 게시판 목록 "
                        "후보를 찾지 못함"
                    )

            # =================================================
            # 4. 미분류 URL
            # =================================================

            else:

                result = await verify_list_url(
                    session,
                    candidate_url
                )

                best_result = result

                if result["auto_confirm"]:

                    final_list_url = result[
                        "final_url"
                    ]

        except Exception as e:

            error = str(e)

            best_result = {
                "access": False,
                "auto_confirm": False,
                "reason": "검증 중 오류",
                "score": 0,
                "post_count": 0,
                "verified_post_count": 0,
                "pagination": False,
                "final_url": "",
                "post": None,
                "structure": None,
            }

        # =====================================================
        # 결과 판정
        # =====================================================

        if error:

            status = "오류"

        elif best_result[
            "auto_confirm"
        ]:

            status = "자동확정"

        elif best_result[
            "access"
        ]:

            status = "수동확인"

        else:

            status = "제외"

        post = best_result.get(
            "post"
        )

        structure = best_result.get(
            "structure"
        )

        if structure:

            evidence = " / ".join(
                structure.get(
                    "evidence",
                    []
                )
            )

            table_count = structure.get(
                "table_count",
                0
            )

            li_count = structure.get(
                "li_count",
                0
            )

            keyword_count = structure.get(
                "keyword_count",
                0
            )

        else:

            evidence = ""

            table_count = 0

            li_count = 0

            keyword_count = 0

        return {
            "기관명": institution,

            "V6.6후보URL": candidate_url,

            "V6.6후보유형": candidate_type,

            "V6.6최종목록URL": final_list_url,

            "V6.6최종URL유형": (
                classify_url(
                    final_list_url
                )
                if final_list_url
                else ""
            ),

            "V6.6목록접속": (
                "TRUE"
                if best_result["access"]
                else "FALSE"
            ),

            "V6.6게시물링크수": best_result[
                "post_count"
            ],

            "V6.6실제게시물수": best_result[
                "verified_post_count"
            ],

            "V6.6페이지네이션": (
                "TRUE"
                if best_result["pagination"]
                else "FALSE"
            ),

            "V6.6목록키워드수": keyword_count,

            "V6.6TABLE수": table_count,

            "V6.6LI수": li_count,

            "V6.6구조근거": evidence,

            "V6.6검증게시물URL": (
                post["url"]
                if post
                else ""
            ),

            "V6.6검증게시물제목": (
                post["title"]
                if post
                else ""
            ),

            "V6.6검증게시물날짜": (
                post["date"]
                if post
                else ""
            ),

            "V6.6검증본문길이": (
                post["body_length"]
                if post
                else 0
            ),

            "V6.6점수": best_result[
                "score"
            ],

            "V6.6결과": status,

            "V6.6사유": best_result[
                "reason"
            ],

            "V6.6오류": error,
        }


# =========================================================
# MAIN
# =========================================================

async def main():

    # -----------------------------------------------------
    # 입력파일 확인
    # -----------------------------------------------------

    if not os.path.exists(
        INPUT_FILE
    ):

        raise FileNotFoundError(
            f"입력파일 없음: {INPUT_FILE}"
        )

    df = pd.read_excel(
        INPUT_FILE
    )

    # -----------------------------------------------------
    # URL 컬럼 결정
    # -----------------------------------------------------

    url_col = None

    # V6.5에서 자동확정된 최종목록URL 우선
    if "V6.5최종목록URL" in df.columns:

        url_col = "V6.5최종목록URL"

    elif "V6.4검증_목록URL" in df.columns:

        url_col = "V6.4검증_목록URL"

    elif "V6.3게시판URL" in df.columns:

        url_col = "V6.3게시판URL"

    elif "V6.5후보URL" in df.columns:

        url_col = "V6.5후보URL"

    else:

        raise ValueError(
            "검증 URL 컬럼을 찾을 수 없습니다."
        )

    # -----------------------------------------------------
    # 기관명
    # -----------------------------------------------------

    if "기관명" not in df.columns:

        raise ValueError(
            "기관명 컬럼 없음"
        )

    candidates = []

    for _, row in df.iterrows():

        url = row.get(
            url_col
        )

        # -------------------------------------------------
        # URL fallback
        # -------------------------------------------------

        if (
            pd.isna(url)
            or not str(url).strip()
        ):

            fallback_columns = [
                "V6.5후보URL",
                "V6.4검증_목록URL",
                "V6.3게시판URL",
            ]

            for fallback in fallback_columns:

                if fallback not in df.columns:

                    continue

                value = row.get(
                    fallback
                )

                if (
                    not pd.isna(value)
                    and str(value).strip()
                ):

                    url = value

                    break

        if pd.isna(url):

            continue

        url = normalize_url(
            url
        )

        if not url:

            continue

        institution = str(
            row.get(
                "기관명",
                ""
            )
        ).strip()

        if not institution:

            continue

        candidates.append(
            (
                institution,
                url
            )
        )

    # -----------------------------------------------------
    # 시작
    # -----------------------------------------------------

    print("=" * 75)

    print(
        "V6.6 Structural Board Validation"
    )

    print("=" * 75)

    print(
        f"입력파일 : {INPUT_FILE}"
    )

    print(
        f"URL 컬럼 : {url_col}"
    )

    print(
        f"검증 대상 : {len(candidates)}"
    )

    print("=" * 75)

    # -----------------------------------------------------
    # HTTP
    # -----------------------------------------------------

    semaphore = asyncio.Semaphore(
        CONCURRENCY
    )

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=False
    )

    async with aiohttp.ClientSession(
        connector=connector,
        timeout=TIMEOUT,
        headers=HEADERS
    ) as session:

        tasks = [
            process_candidate(
                session,
                semaphore,
                institution,
                url
            )
            for institution, url
            in candidates
        ]

        results = []

        for number, task in enumerate(
            asyncio.as_completed(tasks),
            start=1
        ):

            try:

                result = await task

            except Exception as e:

                result = {
                    "기관명": "UNKNOWN",
                    "V6.6후보URL": "",
                    "V6.6후보유형": "",
                    "V6.6최종목록URL": "",
                    "V6.6최종URL유형": "",
                    "V6.6목록접속": "FALSE",
                    "V6.6게시물링크수": 0,
                    "V6.6실제게시물수": 0,
                    "V6.6페이지네이션": "FALSE",
                    "V6.6목록키워드수": 0,
                    "V6.6TABLE수": 0,
                    "V6.6LI수": 0,
                    "V6.6구조근거": "",
                    "V6.6검증게시물URL": "",
                    "V6.6검증게시물제목": "",
                    "V6.6검증게시물날짜": "",
                    "V6.6검증본문길이": 0,
                    "V6.6점수": 0,
                    "V6.6결과": "오류",
                    "V6.6사유": "작업 처리 오류",
                    "V6.6오류": str(e),
                }

            results.append(
                result
            )

            print(
                f"[V6.6] "
                f"{number}/{len(candidates)} | "
                f"{result['기관명']} | "
                f"{result['V6.6결과']} | "
                f"점수 {result['V6.6점수']} | "
                f"링크 {result['V6.6게시물링크수']} | "
                f"실제게시물 {result['V6.6실제게시물수']}"
            )

    # -----------------------------------------------------
    # 순서 복원
    # -----------------------------------------------------

    order = {
        institution: i
        for i, (
            institution,
            _
        ) in enumerate(candidates)
    }

    results.sort(
        key=lambda x:
        order.get(
            x["기관명"],
            999999
        )
    )

    # -----------------------------------------------------
    # 기존 데이터 + V6.6 결과
    # -----------------------------------------------------

    output = df.copy()

    result_columns = [
        "V6.6후보URL",
        "V6.6후보유형",
        "V6.6최종목록URL",
        "V6.6최종URL유형",
        "V6.6목록접속",
        "V6.6게시물링크수",
        "V6.6실제게시물수",
        "V6.6페이지네이션",
        "V6.6목록키워드수",
        "V6.6TABLE수",
        "V6.6LI수",
       
