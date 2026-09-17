import asyncio
import aiohttp
import pandas as pd
import re
import os
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs

INPUT_FILE = "boards_v6_4_validation.xlsx"
OUTPUT_FILE = "boards_v6_5_validation.xlsx"

TIMEOUT = aiohttp.ClientTimeout(total=40)
CONCURRENCY = 6

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    )
}

# --------------------------------------------------
# URL 분류
# --------------------------------------------------

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
]

DATE_PATTERNS = [
    r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}",
    r"\d{4}\.\d{1,2}\.\d{1,2}",
    r"\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}",
]


def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def same_host(a, b):
    try:
        h1 = urlparse(a).netloc.lower().replace("www.", "")
        h2 = urlparse(b).netloc.lower().replace("www.", "")
        return h1 == h2
    except Exception:
        return False


def classify_url(url):
    try:
        p = urlparse(url)
        path = p.path.lower()
        query = parse_qs(p.query)

        if path in HOMEPAGE_PATHS:
            return "홈페이지"

        for key in query.keys():
            if key.lower() in {x.lower() for x in DETAIL_QUERY_KEYS}:
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


def clean_text(text):
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


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
            el = soup.select_one(selector)

            if el:
                text = clean_text(
                    el.get_text(" ", strip=True)
                )

                if 3 <= len(text) <= 300:
                    return text

        except Exception:
            pass

    return ""


def extract_date(text):
    for pattern in DATE_PATTERNS:
        m = re.search(pattern, text)

        if m:
            return m.group(0)

    return ""


# --------------------------------------------------
# HTTP
# --------------------------------------------------

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


# --------------------------------------------------
# 게시물 링크 추출
# --------------------------------------------------

def extract_post_links(base_url, soup):

    candidates = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith(
            ("javascript:", "#", "mailto:")
        ):
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

        kind = classify_url(absolute)

        if kind == "홈페이지":
            continue

        text = clean_text(
            a.get_text(" ", strip=True)
        )

        if len(text) < 2:
            continue

        score = 0
        lower = absolute.lower()

        # 상세 URL 패턴
        if kind == "상세":
            score += 5

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

        # 제목처럼 보이는 링크
        if 5 <= len(text) <= 200:
            score += 1

        if score >= 4:

            candidates.append(
                (absolute, text, score)
            )

    # URL 중복 제거
    unique = {}

    for url, text, score in candidates:

        if url not in unique:
            unique[url] = (
                text,
                score
            )

    result = []

    for url, value in unique.items():

        result.append(
            (
                url,
                value[0],
                value[1]
            )
        )

    return result


# --------------------------------------------------
# 페이지네이션
# --------------------------------------------------

def detect_pagination(soup):

    number_count = 0

    for a in soup.find_all("a", href=True):

        text = clean_text(
            a.get_text(" ", strip=True)
        )

        if re.fullmatch(
            r"\d{1,3}",
            text
        ):
            number_count += 1

    if number_count >= 2:
        return True

    page_words = [
        "다음",
        "이전",
        "페이지",
        "첫 페이지",
        "마지막 페이지",
        "page",
    ]

    text = clean_text(
        soup.get_text(" ", strip=True)
    ).lower()

    return any(
        word.lower() in text
        for word in page_words
    )


# --------------------------------------------------
# 목록 구조 검사
# --------------------------------------------------

def inspect_list_structure(
    list_url,
    soup
):

    text = clean_text(
        soup.get_text(" ", strip=True)
    )

    score = 0

    keyword_count = 0

    for keyword in LIST_TEXT:

        if keyword in text:
            keyword_count += 1

    if keyword_count >= 4:
        score += 3

    elif keyword_count >= 2:
        score += 1

    if soup.find("table"):
        score += 2

    html = str(soup).lower()

    if any(
        word in html
        for word in [
            "board",
            "bbs",
            "notice",
            "list",
        ]
    ):
        score += 1

    post_links = extract_post_links(
        list_url,
        soup
    )

    if len(post_links) >= 5:
        score += 3

    elif len(post_links) >= 3:
        score += 2

    elif len(post_links) >= 2:
        score += 1

    pagination = detect_pagination(soup)

    if pagination:
        score += 1

    return {
        "score": score,
        "post_links": post_links,
        "pagination": pagination,
        "keyword_count": keyword_count,
    }


# --------------------------------------------------
# 게시물 실제 검증
# --------------------------------------------------

async def verify_posts(
    session,
    list_url,
    post_links
):

    checked = []

    # 최대 8개 실제 게시물 검사
    for post_url, link_text, link_score in post_links[:8]:

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

        # 홈페이지로 튀면 제외
        if classify_url(final_url) == "홈페이지":
            continue

        title = extract_title(soup)

        body = clean_text(
            soup.get_text(
                " ",
                strip=True
            )
        )

        date = extract_date(body)

        if not title:
            continue

        if len(body) < 40:
            continue

        checked.append(
            {
                "url": final_url,
                "title": title,
                "date": date,
                "body_length": len(body),
            }
        )

    return checked


# --------------------------------------------------
# 상세페이지 → 목록 URL
# --------------------------------------------------

async def recover_list_from_detail(
    session,
    detail_url
):

    result = await fetch(
        session,
        detail_url
    )

    if (
        result["status"] >= 400
        or not result["text"]
    ):
        return ""

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

        text = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

        score = 0

        for word in LIST_BUTTON_TEXT:

            if word in text:
                score += 8

        if classify_url(
            absolute
        ) == "목록":
            score += 5

        if classify_url(
            absolute
        ) == "상세":
            score -= 5

        if score > 0:
            candidates.append(
                (score, absolute)
            )

    candidates.sort(
        reverse=True
    )

    # 후보를 실제 목록으로 검증
    for _, candidate in candidates[:10]:

        verified = await verify_list_url(
            session,
            candidate
        )

        if verified["auto_confirm"]:
            return candidate

    return ""


# --------------------------------------------------
# 홈페이지 → 게시판 찾기
# --------------------------------------------------

async def discover_from_homepage(
    session,
    homepage
):

    result = await fetch(
        session,
        homepage
    )

    if (
        result["status"] >= 400
        or not result["text"]
    ):
        return ""

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
                score += 5

        if kind == "목록":
            score += 4

        if any(
            pattern in absolute.lower()
            for pattern in LIST_PATTERNS
        ):
            score += 3

        if score >= 5:
            candidates.append(
                (score, absolute)
            )

    candidates.sort(
        reverse=True
    )

    for _, candidate in candidates[:15]:

        verified = await verify_list_url(
            session,
            candidate
        )

        if verified["auto_confirm"]:
            return candidate

    return ""


# --------------------------------------------------
# 최종 목록 URL 검증
# --------------------------------------------------

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
            "reason": "목록 페이지 접속 실패",
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "post": None,
            "final_url": "",
        }

    final_url = result["url"]

    # 최종적으로 홈페이지가 된 경우
    if classify_url(final_url) == "홈페이지":

        return {
            "access": True,
            "auto_confirm": False,
            "reason": "홈페이지로 리다이렉트",
            "score": 0,
            "post_count": 0,
            "verified_post_count": 0,
            "pagination": False,
            "post": None,
            "final_url": final_url,
        }

    soup = BeautifulSoup(
        result["text"],
        "html.parser"
    )

    structure = inspect_list_structure(
        final_url,
        soup
    )

    post_links = structure[
        "post_links"
    ]

    # 실제 게시물 검증
    verified_posts = await verify_posts(
        session,
        final_url,
        post_links
    )

    # 서로 다른 게시물 확인
    unique_titles = set()

    for post in verified_posts:
        unique_titles.add(
            post["title"]
        )

    verified_count = len(
        unique_titles
    )

    score = structure["score"]

    # 실제 게시물
    if verified_count >= 5:
        score += 5

    elif verified_count >= 3:
        score += 4

    elif verified_count >= 2:
        score += 2

    elif verified_count >= 1:
        score += 1

    # 실제 목록 URL 판정
    list_kind = classify_url(
        final_url
    )

    if list_kind == "목록":
        score += 2

    # 자동확정 조건
    auto_confirm = (
        list_kind == "목록"
        and len(post_links) >= 3
        and verified_count >= 3
        and structure["keyword_count"] >= 2
        and score >= 10
    )

    if auto_confirm:

        reason = (
            f"실제 목록 URL 확인 / "
            f"게시물 링크 {len(post_links)}개 / "
            f"실제 게시물 {verified_count}개 확인 / "
            f"점수 {score}"
        )

    else:

        reasons = []

        if list_kind != "목록":
            reasons.append(
                "목록 URL 패턴 불충분"
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
                "목록 구조 약함"
            )

        if score < 10:
            reasons.append(
                f"점수 부족({score})"
            )

        reason = ", ".join(reasons)

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
        "post": (
            verified_posts[0]
            if verified_posts
            else None
        ),
        "final_url": final_url,
    }


# --------------------------------------------------
# 기관별 처리
# --------------------------------------------------

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

        result = None
        final_list_url = ""
        error = ""

        try:

            # 1. 목록
            if candidate_type == "목록":

                result = await verify_list_url(
                    session,
                    candidate_url
                )

                if result[
                    "auto_confirm"
                ]:
                    final_list_url = result[
                        "final_url"
                    ]

            # 2. 상세
            elif candidate_type == "상세":

                list_url = await recover_list_from_detail(
                    session,
                    candidate_url
                )

                if list_url:

                    result = await verify_list_url(
                        session,
                        list_url
                    )

                    if result[
                        "auto_confirm"
                    ]:
                        final_list_url = result[
                            "final_url"
                        ]

                else:

                    result = {
                        "access": False,
                        "auto_confirm": False,
                        "reason": "상세페이지에서 실제 목록 URL을 찾지 못함",
                        "score": 0,
                        "post_count": 0,
                        "verified_post_count": 0,
                        "pagination": False,
                        "post": None,
                        "final_url": "",
                    }

            # 3. 홈페이지
            elif candidate_type == "홈페이지":

                list_url = await discover_from_homepage(
                    session,
                    candidate_url
                )

                if list_url:

                    result = await verify_list_url(
                        session,
                        list_url
                    )

                    if result[
                        "auto_confirm"
                    ]:
                        final_list_url = result[
                            "final_url"
                        ]

                else:

                    result = {
                        "access": False,
                        "auto_confirm": False,
                        "reason": "홈페이지에서 실제 게시판 목록을 찾지 못함",
                        "score": 0,
                        "post_count": 0,
                        "verified_post_count": 0,
                        "pagination": False,
                        "post": None,
                        "final_url": "",
                    }

            # 4. 미분류
            else:

                result = await verify_list_url(
                    session,
                    candidate_url
                )

                if result[
                    "auto_confirm"
                ]:
                    final_list_url = result[
                        "final_url"
                    ]

        except Exception as e:

            error = str(e)

            result = {
                "access": False,
                "auto_confirm": False,
                "reason": "검증 오류",
                "score": 0,
                "post_count": 0,
                "verified_post_count": 0,
                "pagination": False,
                "post": None,
                "final_url": "",
            }

        if error:
            status = "오류"

        elif result[
            "auto_confirm"
        ]:
            status = "자동확정"

        elif result[
            "access"
        ]:
            status = "수동확인"

        else:
            status = "제외"

        post = result.get(
            "post"
        )

        return {
            "기관명": institution,
            "V6.5후보URL": candidate_url,
            "V6.5후보유형": candidate_type,
            "V6.5최종목록URL": final_list_url,
            "V6.5목록접속": result[
                "access"
            ],
            "V6.5게시물링크수": result[
                "post_count"
            ],
            "V6.5실제게시물수": result[
                "verified_post_count"
            ],
            "V6.5페이지네이션": result[
                "pagination"
            ],
            "V6.5검증게시물URL": (
                post["url"]
                if post
                else ""
            ),
            "V6.5검증게시물제목": (
                post["title"]
                if post
                else ""
            ),
            "V6.5검증게시물날짜": (
                post["date"]
                if post
                else ""
            ),
            "V6.5검증본문길이": (
                post["body_length"]
                if post
                else 0
            ),
            "V6.5점수": result[
                "score"
            ],
            "V6.5결과": status,
            "V6.5사유": result[
                "reason"
            ],
            "V6.5오류": error,
        }


# --------------------------------------------------
# MAIN
# --------------------------------------------------

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

    # V6.4 최종 후보 URL을 우선 사용
    url_col = None

    for col in df.columns:

        if str(col) == "V6.4검증_목록URL":
            url_col = col
            break

    # 없으면 V6.3 후보 사용
    if url_col is None:

        for col in df.columns:

            if "V6.3게시판URL" in str(col):
                url_col = col
                break

    if url_col is None:
        raise ValueError(
            "검증 URL 컬럼을 찾을 수 없습니다."
        )

    name_col = None

    for col in df.columns:

        if str(col).strip() == "기관명":
            name_col = col
            break

    if name_col is None:
        raise ValueError(
            "기관명 컬럼 없음"
        )

    candidates = []

    for _, row in df.iterrows():

        url = row.get(
            url_col
        )

        if pd.isna(url):
            continue

        url = normalize_url(
            url
        )

        if not url:
            continue

        institution = str(
            row.get(
                name_col,
                ""
            )
        ).strip()

        candidates.append(
            (
                institution,
                url
            )
        )

    print("=" * 70)
    print("V6.5 Deep Board Validation")
    print("=" * 70)
    print(f"입력파일 : {INPUT_FILE}")
    print(f"URL 컬럼 : {url_col}")
    print(f"검증 대상 : {len(candidates)}")
    print("=" * 70)

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

            result = await task

            results.append(
                result
            )

            print(
                f"[V6.5] "
                f"{number}/{len(candidates)} | "
                f"{result['기관명']} | "
                f"{result['V6.5결과']} | "
                f"점수 {result['V6.5점수']} | "
                f"링크 {result['V6.5게시물링크수']} | "
                f"실제게시물 {result['V6.5실제게시물수']}"
            )

    # 기관명 기준 원래 순서
    order = {
        institution: i
        for i, (institution, _)
        in enumerate(candidates)
    }

    results.sort(
        key=lambda x:
        order.get(
            x["기관명"],
            999999
        )
    )

    result_map = {
        r["기관명"]: r
        for r in results
    }

    output = df.copy()

    columns = [
        "V6.5후보URL",
        "V6.5후보유형",
        "V6.5최종목록URL",
        "V6.5목록접속",
        "V6.5게시물링크수",
        "V6.5실제게시물수",
        "V6.5페이지네이션",
        "V6.5검증게시물URL",
        "V6.5검증게시물제목",
        "V6.5검증게시물날짜",
        "V6.5검증본문길이",
        "V6.5점수",
        "V6.5결과",
        "V6.5사유",
        "V6.5오류",
    ]

    for col in columns:
        output[col] = ""

    for idx, row in output.iterrows():

        institution = str(
            row[name_col]
        ).strip()

        if institution not in result_map:
            continue

        r = result_map[
            institution
        ]

        for col in columns:

            output.at[
                idx,
                col
            ] = r.get(
                col,
                ""
            )

    output.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print()
    print("=" * 70)
    print("V6.5 검증 완료")
    print("=" * 70)

    print(
        "자동확정 :",
        sum(
            x["V6.5결과"]
            == "자동확정"
            for x in results
        )
    )

    print(
        "수동확인 :",
        sum(
            x["V6.5결과"]
            == "수동확인"
            for x in results
        )
    )

    print(
        "제외     :",
        sum(
            x["V6.5결과"]
            == "제외"
            for x in results
        )
    )

    print(
        "오류     :",
        sum(
            x["V6.5결과"]
            == "오류"
            for x in results
        )
    )

    print(
        "결과파일 :",
        OUTPUT_FILE
    )

    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(main())
