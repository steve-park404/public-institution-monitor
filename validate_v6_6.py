import asyncio
import aiohttp
import pandas as pd
import re
import os

from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs


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
# URL 기준
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
    "articleno",
    "nttid",
    "list_no",
    "boardid",
    "bbsidx",
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
]


# =========================================================
# 공통
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


def clean_text(text):

    if not text:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(text)
    ).strip()


def same_host(a, b):

    try:

        h1 = urlparse(a).netloc.lower().replace("www.", "")
        h2 = urlparse(b).netloc.lower().replace("www.", "")

        return h1 == h2

    except Exception:

        return False


def same_url(a, b):

    try:

        p1 = urlparse(a)
        p2 = urlparse(b)

        return (
            p1.netloc.lower().replace("www.", "")
            == p2.netloc.lower().replace("www.", "")
            and
            p1.path.rstrip("/").lower()
            == p2.path.rstrip("/").lower()
            and
            p1.query.lower()
            == p2.query.lower()
        )

    except Exception:

        return False


# =========================================================
# URL 분류
# =========================================================

def classify_url(url):

    try:

        p = urlparse(url)

        path = p.path.lower()

        if path in HOMEPAGE_PATHS:
            return "홈페이지"

        query = parse_qs(p.query)

        for key in query.keys():

            if key.lower() in DETAIL_QUERY_KEYS:

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
# 제목 / 날짜
# =========================================================

def extract_title(soup):

    selectors = [
        ".subject",
        ".title",
        ".board-title",
        ".view-title",
        ".bbs-title",
        "h1",
        "h2",
        "h3",
        ".tit",
        "title",
    ]

    for selector in selectors:

        try:

            el = soup.select_one(selector)

            if not el:
                continue

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
# 게시물 제목 링크 판별
# =========================================================

def looks_like_post_title(text):

    text = clean_text(text)

    if len(text) < 2:
        return False

    if len(text) > 250:
        return False

    bad_words = [
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

    for word in bad_words:

        if text.upper() == word.upper():

            return False

    return True


# =========================================================
# 게시물 링크 추출
# =========================================================

def extract_post_links(base_url, soup):

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

        text = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

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

        if same_url(
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

        # 게시물 행에 날짜가 있는 경우 가점
        parent = a.parent

        if parent:

            parent_text = clean_text(
                parent.get_text(
                    " ",
                    strip=True
                )
            )

            if re.search(
                r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}",
                parent_text
            ):

                score += 2

        if score >= 5:

            candidates.append(
                (
                    absolute,
                    text,
                    score
                )
            )

    # URL 기준 중복 제거
    unique = {}

    for url, text, score in candidates:

        key = (
            urlparse(url).netloc.lower(),
            urlparse(url).path.lower(),
            urlparse(url).query.lower(),
        )

        if key not in unique:

            unique[key] = (
                url,
                text,
                score
            )

        elif score > unique[key][2]:

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
            x in text
            for x in [
                "다음",
                "이전",
                "첫 페이지",
                "마지막 페이지",
            ]
        ):

            return True

        if any(
            x in href
            for x in [
                "page=",
                "pageno=",
                "pageNo=",
                "currentpage=",
            ]
        ):

            return True

    return False


# =========================================================
# 게시판 구조 검사
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

    score = 0

    evidence = []

    keyword_count = 0

    for keyword in LIST_TEXT:

        if keyword in text:

            keyword_count += 1

    if keyword_count >= 5:

        score += 4

        evidence.append(
            f"목록키워드 {keyword_count}개"
        )

    elif keyword_count >= 3:

        score += 3

        evidence.append(
            f"목록키워드 {keyword_count}개"
        )

    elif keyword_count >= 2:

        score += 1

        evidence.append(
            f"목록키워드 {keyword_count}개"
        )

    tables = soup.find_all("table")

    if tables:

        score += 3

        evidence.append(
            f"TABLE {len(tables)}개"
        )

    list_items = soup.find_all("li")

    if len(list_items) >= 5:

        score += 1

        evidence.append(
            f"LI {len(list_items)}개"
        )

    html = str(
        soup
    ).lower()

    board_word_count = sum(
        html.count(x)
        for x in [
            "board",
            "bbs",
            "notice",
            "post",
            "article",
        ]
    )

    if board_word_count >= 3:

        score += 2

        evidence.append(
            "게시판 HTML 구조"
        )

    elif board_word_count >= 1:

        score += 1

    pagination = detect_pagination(
        soup
    )

    if pagination:

        score += 2

        evidence.append(
            "페이지네이션"
        )

    post_links = extract_post_links(
        list_url,
        soup
    )

    if len(post_links) >= 10:

        score += 4

        evidence.append(
            f"게시물링크 {len(post_links)}개"
        )

    elif len(post_links) >= 5:

        score += 3

        evidence.append(
            f"게시물링크 {len(post_links)}개"
        )

    elif len(post_links) >= 3:

        score += 2

        evidence.append(
            f"게시물링크 {len(post_links)}개"
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
# 실제 게시물 검증
# =========================================================

async def verify_posts(
    session,
    post_links
):

    checked = []

    seen_titles = set()

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

        if classify_url(
            final_url
        ) == "홈페이지":

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

        if not title:
            continue

        if len(body) < 80:
            continue

        title_key = title.lower()

        if title_key in seen_titles:
            continue

        seen_titles.add(
            title_key
        )

        checked.append(
            {
                "url": final_url,
                "title": title,
                "date": extract_date(body),
                "body_length": len(body),
            }
        )

    return checked


# =========================================================
# 목록 URL 검증
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
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": "",
            "post": None,
            "reason": "페이지 접속 실패",
            "structure": None,
        }

    final_url = result["url"]

    final_kind = classify_url(
        final_url
    )

    # 홈페이지로 리다이렉트
    if final_kind == "홈페이지":

        return {
            "access": True,
            "auto_confirm": False,
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": final_url,
            "post": None,
            "reason": "홈페이지로 리다이렉트",
            "structure": None,
        }

    # 상세페이지는 목록으로 인정하지 않음
    if final_kind == "상세":

        return {
            "access": True,
            "auto_confirm": False,
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": final_url,
            "post": None,
            "reason": "상세페이지",
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

    verified_count = len(
        {
            x["title"].lower()
            for x in verified_posts
        }
    )

    score = structure[
        "score"
    ]

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

    if final_kind == "목록":

        score += 2

    # -----------------------------------------------------
    # V6.6 자동확정 기준
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
            "목록형 URL 아님"
        )

    if len(post_links) < 3:

        reasons.append(
            f"게시물링크 부족({len(post_links)}개)"
        )

    if verified_count < 3:

        reasons.append(
            f"실제게시물 부족({verified_count}개)"
        )

    if structure["keyword_count"] < 2:

        reasons.append(
            f"목록키워드 부족({structure['keyword_count']}개)"
        )

    if structure["score"] < 8:

        reasons.append(
            f"구조점수 부족({structure['score']}점)"
        )

    if auto_confirm:

        reason = (
            "실제 게시판 목록 구조 확인 / "
            f"게시물링크 {len(post_links)}개 / "
            f"실제게시물 {verified_count}개 / "
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
        "score": score,
        "post_count": len(post_links),
        "verified_post_count": verified_count,
        "pagination": structure["pagination"],
        "final_url": final_url,
        "post": (
            verified_posts[0]
            if verified_posts
            else None
        ),
        "reason": reason,
        "structure": structure,
    }


# =========================================================
# 상세페이지 → 목록 후보
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

        for word in LIST_BUTTON_TEXT:

            if word in text:

                score += 10

        if kind == "목록":

            score += 8

        if kind == "상세":

            score -= 10

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

    # 중복
    unique = {}

    for score, url in candidates:

        key = (
            urlparse(url).netloc.lower(),
            urlparse(url).path.lower(),
            urlparse(url).query.lower(),
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
# 홈페이지 → 목록 후보
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

    unique = {}

    for score, url in candidates:

        key = (
            urlparse(url).netloc.lower(),
            urlparse(url).path.lower(),
            urlparse(url).query.lower(),
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
# 기관 처리
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

        best_result = {
            "access": False,
            "auto_confirm": False,
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "final_url": "",
            "post": None,
            "reason": "",
            "structure": None,
        }

        final_list_url = ""
        error = ""

        try:

            # -------------------------------------------------
            # 목록
            # -------------------------------------------------

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

            # -------------------------------------------------
            # 상세
            # -------------------------------------------------

            elif candidate_type == "상세":

                candidates = (
                    await recover_list_from_detail(
                        session,
                        candidate_url
                    )
                )

                for list_url in candidates:

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

                if not candidates:

                    best_result[
                        "reason"
                    ] = (
                        "상세페이지에서 "
                        "목록 후보를 찾지 못함"
                    )

            # -------------------------------------------------
            # 홈페이지
            # -------------------------------------------------

            elif candidate_type == "홈페이지":

                candidates = (
                    await discover_from_homepage(
                        session,
                        candidate_url
                    )
                )

                for list_url in candidates:

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

                if not candidates:

                    best_result[
                        "reason"
                    ] = (
                        "홈페이지에서 "
                        "목록 후보를 찾지 못함"
                    )

            # -------------------------------------------------
            # 미분류
            # -------------------------------------------------

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
                "score": 0,
                "post_count": 0,
                "verified_post_count": 0,
                "pagination": False,
                "final_url": "",
                "post": None,
                "reason": "검증 중 오류",
                "structure": None,
            }

        # -----------------------------------------------------
        # 최종 상태
        # -----------------------------------------------------

        if error:

            status = "오류"

        elif best_result["auto_confirm"]:

            status = "자동확정"

        elif best_result["access"]:

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

            keyword_count = structure.get(
                "keyword_count",
                0
            )

            table_count = structure.get(
                "table_count",
                0
            )

            li_count = structure.get(
                "li_count",
                0
            )

        else:

            evidence = ""
            keyword_count = 0
            table_count = 0
            li_count = 0

        return {
            "기관명": institution,
            "V6.6후보URL": candidate_url,
            "V6.6후보유형": candidate_type,
            "V6.6최종목록URL": final_list_url,
            "V6.6최종URL유형": (
                classify_url(final_list_url)
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
    # URL 컬럼
    # -----------------------------------------------------

    url_col = None

    for candidate in [
        "V6.5최종목록URL",
        "V6.5후보URL",
        "V6.4검증_목록URL",
        "V6.3게시판URL",
    ]:

        if candidate in df.columns:

            url_col = candidate

            break

    if not url_col:

        raise ValueError(
            "검증 URL 컬럼을 찾을 수 없습니다."
        )

    if "기관명" not in df.columns:

        raise ValueError(
            "기관명 컬럼 없음"
        )

    candidates = []

    for _, row in df.iterrows():

        url = row.get(
            url_col
        )

        # 빈 경우 fallback
        if (
            pd.isna(url)
            or not str(url).strip()
        ):

            for fallback in [
                "V6.5후보URL",
                "V6.4검증_목록URL",
                "V6.3게시판URL",
            ]:

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
    print("V6.6 Structural Board Validation")
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
            for institution, url in candidates
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
    # 원래 기관 순서
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
    # 결과 컬럼
    # -----------------------------------------------------

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
        "V6.6구조근거",
        "V6.6검증게시물URL",
        "V6.6검증게시물제목",
        "V6.6검증게시물날짜",
        "V6.6검증본문길이",
        "V6.6점수",
        "V6.6결과",
        "V6.6사유",
        "V6.6오류",
    ]

    string_columns = [
        "V6.6후보URL",
        "V6.6후보유형",
        "V6.6최종목록URL",
        "V6.6최종URL유형",
        "V6.6목록접속",
        "V6.6페이지네이션",
        "V6.6구조근거",
        "V6.6검증게시물URL",
        "V6.6검증게시물제목",
        "V6.6검증게시물날짜",
        "V6.6결과",
        "V6.6사유",
        "V6.6오류",
    ]

    numeric_columns = [
        "V6.6게시물링크수",
        "V6.6실제게시물수",
        "V6.6목록키워드수",
        "V6.6TABLE수",
        "V6.6LI수",
        "V6.6검증본문길이",
        "V6.6점수",
    ]

    # -----------------------------------------------------
    # 결과 DataFrame
    # -----------------------------------------------------

    result_df = pd.DataFrame(
        results
    )

    for col in result_columns:

        if col not in result_df.columns:

            result_df[col] = ""

    result_df = result_df[
        ["기관명"] + result_columns
    ]

    # -----------------------------------------------------
    # 원본 + V6.6 결과
    # -----------------------------------------------------

    output = df.copy()

    for col in string_columns:

        output[col] = ""

    for col in numeric_columns:

        output[col] = 0

    result_map = {
        r["기관명"]: r
        for r in results
    }

    for idx in output.index:

        institution = str(
            output.at[
                idx,
                "기관명"
            ]
        ).strip()

        if institution not in result_map:

            continue

        r = result_map[
            institution
        ]

        for col in string_columns:

            value = r.get(
                col,
                ""
            )

            if value is None:

                value = ""

            output.at[
                idx,
                col
            ] = str(value)

        for col in numeric_columns:

            value = r.get(
                col,
                0
            )

            try:

                value = int(
                    float(value)
                )

            except Exception:

                value = 0

            output.at[
                idx,
                col
            ] = value

    # -----------------------------------------------------
    # dtype 정리
    # -----------------------------------------------------

    for col in string_columns:

        output[col] = (
            output[col]
            .fillna("")
            .astype(str)
        )

    for col in numeric_columns:

        output[col] = pd.to_numeric(
            output[col],
            errors="coerce"
        ).fillna(0).astype(int)

    # -----------------------------------------------------
    # 저장
    # -----------------------------------------------------

    output.to_excel(
        OUTPUT_FILE,
        index=False,
        engine="openpyxl"
    )

    # -----------------------------------------------------
    # 통계
    # -----------------------------------------------------

    auto_count = sum(
        x["V6.6결과"] == "자동확정"
        for x in results
    )

    manual_count = sum(
        x["V6.6결과"] == "수동확인"
        for x in results
    )

    exclude_count = sum(
        x["V6.6결과"] == "제외"
        for x in results
    )

    error_count = sum(
        x["V6.6결과"] == "오류"
        for x in results
    )

    print()
    print("=" * 75)
    print("V6.6 구조검증 완료")
    print("=" * 75)

    print(
        f"자동확정 : {auto_count}"
    )

    print(
        f"수동확인 : {manual_count}"
    )

    print(
        f"제외     : {exclude_count}"
    )

    print(
        f"오류     : {error_count}"
    )

    print(
        f"결과파일 : {OUTPUT_FILE}"
    )

    print("=" * 75)


if __name__ == "__main__":

    asyncio.run(
        main()
    )
