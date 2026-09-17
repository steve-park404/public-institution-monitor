import asyncio
import aiohttp
import pandas as pd
import re
import os
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs, urlunparse

INPUT_FILE = "boards_v6_3_analysis.xlsx"
OUTPUT_FILE = "boards_v6_4_validation.xlsx"

TIMEOUT = aiohttp.ClientTimeout(total=45)
CONCURRENCY = 8

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    )
}

LIST_TEXT_KEYWORDS = [
    "번호",
    "제목",
    "등록일",
    "작성일",
    "조회",
    "공지",
    "채용",
    "소식",
    "뉴스",
    "자료실",
    "알림",
]

POST_TEXT_KEYWORDS = [
    "제목",
    "작성일",
    "등록일",
    "조회",
    "첨부파일",
    "첨부",
]

LIST_WORDS = [
    "목록",
    "리스트",
    "전체목록",
    "공지사항",
    "공지",
    "채용",
    "뉴스",
    "소식",
    "알림",
    "자료실",
]

DETAIL_QUERY_KEYS = [
    "act",
    "articleNo",
    "nttId",
    "list_no",
    "boardId",
    "bbsIdx",
    "idx",
    "seq",
]

DETAIL_QUERY_VALUES = [
    "view",
    "detail",
]

DETAIL_PATH_WORDS = [
    "/view",
    "/detail",
    "boardview",
    "selectarticle",
]

HOMEPAGE_PATHS = [
    "",
    "/",
    "/main",
    "/main/",
    "/main/index.do",
    "/home",
    "/home/",
    "/home/index.do",
    "/index.do",
]


def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def same_host(url1, url2):
    try:
        h1 = urlparse(url1).netloc.lower()
        h2 = urlparse(url2).netloc.lower()

        if h1 == h2:
            return True

        # www 차이 허용
        h1 = h1.replace("www.", "")
        h2 = h2.replace("www.", "")

        return h1 == h2

    except Exception:
        return False


def classify_url(url):
    """
    후보 URL이 홈페이지 / 목록 / 상세 / 외부서비스 중 무엇인지 판별
    """

    try:
        p = urlparse(url)
        path = p.path.lower()
        query = parse_qs(p.query)

        # 홈페이지
        if path in [x.lower() for x in HOMEPAGE_PATHS]:
            return "홈페이지"

        # 명백한 상세 페이지
        for key in DETAIL_QUERY_KEYS:
            if key.lower() in [k.lower() for k in query.keys()]:
                return "상세"

        for value_list in query.values():
            for value in value_list:
                if str(value).lower() in DETAIL_QUERY_VALUES:
                    return "상세"

        for word in DETAIL_PATH_WORDS:
            if word in path:
                return "상세"

        # 목록 URL 패턴
        list_patterns = [
            "/list",
            "list.do",
            "list.jsp",
            "list.asp",
            "list/",
            "/bbs",
            "/notice",
            "/board",
            "boardlist",
            "listboard",
            "postlist",
        ]

        for pattern in list_patterns:
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
    """
    게시물 제목 추출
    """

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
                text = clean_text(el.get_text(" ", strip=True))

                if 3 <= len(text) <= 300:
                    return text
        except Exception:
            pass

    return ""


def extract_post_links(base_url, soup):
    """
    목록 페이지에서 게시물 링크 후보 추출
    """

    results = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith(("javascript:", "#", "mailto:")):
            continue

        absolute = urljoin(base_url, href)

        if not same_host(base_url, absolute):
            continue

        text = clean_text(a.get_text(" ", strip=True))

        if len(text) < 2 or len(text) > 300:
            continue

        lower_href = absolute.lower()

        score = 0

        # 상세페이지에서 흔히 발견되는 패턴
        if any(x in lower_href for x in [
            "article",
            "view",
            "detail",
            "nttid",
            "articleid",
            "boardid",
            "bbsidx",
            "idx=",
            "seq=",
            "no=",
            "list_no",
        ]):
            score += 3

        # 숫자 파라미터
        if re.search(r"(\?|&)(idx|seq|no|id|articleNo|nttId|list_no|bbsIdx|boardId)=", lower_href):
            score += 3

        # 링크 텍스트가 게시물 제목처럼 보이는 경우
        if len(text) >= 5:
            score += 1

        # 상세 URL 특성
        if "/view" in lower_href or "act=view" in lower_href:
            score += 2

        if "/detail" in lower_href:
            score += 2

        if score >= 3:
            results.append((absolute, text, score))

    # URL 기준 중복 제거
    unique = {}

    for url, text, score in results:
        if url not in unique:
            unique[url] = (text, score)

    return [
        (url, data[0], data[1])
        for url, data in unique.items()
    ]


def detect_pagination(soup):
    """
    페이지네이션 존재 여부
    """

    text = clean_text(soup.get_text(" ", strip=True))

    # 숫자 페이지 링크
    number_links = 0

    for a in soup.find_all("a", href=True):
        t = clean_text(a.get_text(" ", strip=True))

        if re.fullmatch(r"\d+", t):
            number_links += 1

    if number_links >= 2:
        return True

    # 페이지 관련 문구
    pagination_words = [
        "다음",
        "이전",
        "페이지",
        "page",
        "첫 페이지",
        "마지막 페이지",
    ]

    lower_text = text.lower()

    for word in pagination_words:
        if word.lower() in lower_text:
            return True

    return False


def list_semantics_score(soup):
    """
    목록 페이지인지 판단
    """

    text = clean_text(soup.get_text(" ", strip=True))

    score = 0

    for keyword in LIST_TEXT_KEYWORDS:
        if keyword in text:
            score += 1

    # table 존재
    if soup.find("table"):
        score += 2

    # 목록/게시판 관련 class
    html = str(soup).lower()

    if any(x in html for x in [
        "board",
        "bbs",
        "notice",
        "list",
        "table",
    ]):
        score += 2

    return score


def post_semantics_score(soup):
    text = clean_text(soup.get_text(" ", strip=True))

    score = 0

    for keyword in POST_TEXT_KEYWORDS:
        if keyword in text:
            score += 1

    return score


async def fetch(session, url):
    try:
        async with session.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
            allow_redirects=True,
            ssl=False,
        ) as response:

            body = await response.text(errors="ignore")

            return {
                "status": response.status,
                "url": str(response.url),
                "text": body,
            }

    except Exception as e:

        return {
            "status": 0,
            "url": url,
            "text": "",
            "error": str(e),
        }


async def find_list_from_detail(session, detail_url):
    """
    상세페이지에서 '목록' 버튼을 찾아 실제 목록 URL을 복원
    """

    result = await fetch(session, detail_url)

    if result["status"] >= 400 or not result["text"]:
        return ""

    soup = BeautifulSoup(result["text"], "html.parser")

    candidates = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        absolute = urljoin(result["url"], href)

        if not same_host(detail_url, absolute):
            continue

        text = clean_text(a.get_text(" ", strip=True))
        lower = absolute.lower()

        score = 0

        # 버튼 문구
        for word in LIST_WORDS:
            if word in text:
                score += 5

        # 목록 URL 패턴
        for pattern in [
            "/list",
            "list.do",
            "list.asp",
            "list.jsp",
            "/bbs",
            "/notice",
            "/board",
            "boardlist",
        ]:
            if pattern in lower:
                score += 4

        # 상세 URL이면 감점
        if classify_url(absolute) == "상세":
            score -= 5

        if classify_url(absolute) == "홈페이지":
            score -= 5

        if score > 0:
            candidates.append((score, absolute))

    candidates.sort(reverse=True)

    if candidates:
        return candidates[0][1]

    return ""


async def discover_list_from_homepage(session, homepage_url):
    """
    홈페이지 후보인 경우 홈페이지 내부에서 게시판 목록을 탐색
    """

    result = await fetch(session, homepage_url)

    if result["status"] >= 400 or not result["text"]:
        return ""

    soup = BeautifulSoup(result["text"], "html.parser")

    candidates = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        absolute = urljoin(result["url"], href)

        if not same_host(homepage_url, absolute):
            continue

        kind = classify_url(absolute)

        if kind == "상세" or kind == "홈페이지":
            continue

        text = clean_text(a.get_text(" ", strip=True))
        lower = absolute.lower()

        score = 0

        # 메뉴 텍스트
        for word in LIST_WORDS:
            if word in text:
                score += 4

        # URL 패턴
        for pattern in [
            "/list",
            "list.do",
            "list.asp",
            "list.jsp",
            "/bbs",
            "/notice",
            "/board",
            "boardlist",
        ]:
            if pattern in lower:
                score += 3

        if kind == "목록":
            score += 3

        if score >= 4:
            candidates.append((score, absolute, text))

    # 높은 점수 우선
    candidates.sort(reverse=True)

    # 상위 후보를 실제 검증
    for _, candidate_url, _ in candidates[:10]:

        verified = await verify_list_page(
            session,
            candidate_url
        )

        if verified["confirmed"]:
            return candidate_url

    return ""


async def verify_list_page(session, list_url):
    """
    실제 반복 모니터링 가능한 게시판 목록인지 검증
    """

    result = await fetch(session, list_url)

    if result["status"] >= 400 or not result["text"]:
        return {
            "confirmed": False,
            "list_access": False,
            "post_links": 0,
            "post_url": "",
            "post_title": "",
            "body_length": 0,
            "pagination": False,
            "score": 0,
            "reason": "목록 페이지 접속 실패",
        }

    soup = BeautifulSoup(result["text"], "html.parser")

    # 홈페이지가 최종적으로 반환된 경우
    if classify_url(str(result["url"])) == "홈페이지":
        return {
            "confirmed": False,
            "list_access": True,
            "post_links": 0,
            "post_url": "",
            "post_title": "",
            "body_length": 0,
            "pagination": False,
            "score": 0,
            "reason": "홈페이지로 리다이렉트됨",
        }

    semantics = list_semantics_score(soup)

    post_links = extract_post_links(result["url"], soup)

    pagination = detect_pagination(soup)

    # 점수
    score = 0

    # 목록 URL 형태
    if classify_url(list_url) == "목록":
        score += 2

    # 목록 페이지 구조
    if semantics >= 4:
        score += 2
    elif semantics >= 2:
        score += 1

    # 게시물 링크
    if len(post_links) >= 5:
        score += 3
    elif len(post_links) >= 3:
        score += 3
    elif len(post_links) >= 2:
        score += 2

    # 페이지네이션
    if pagination:
        score += 1

    post_url = ""
    post_title = ""
    body_length = 0

    # 게시물 실제 접근
    for candidate_url, _, _ in post_links[:5]:

        post_result = await fetch(session, candidate_url)

        if (
            post_result["status"] >= 400
            or not post_result["text"]
        ):
            continue

        post_soup = BeautifulSoup(
            post_result["text"],
            "html.parser"
        )

        title = extract_title(post_soup)

        body_text = clean_text(
            post_soup.get_text(" ", strip=True)
        )

        body_len = len(body_text)

        post_score = post_semantics_score(post_soup)

        if title and body_len >= 40:
            post_url = candidate_url
            post_title = title
            body_length = body_len

            score += 2

            if post_score >= 2:
                score += 1

            break

    confirmed = (
        score >= 10
        and len(post_links) >= 3
        and bool(post_url)
        and body_length >= 40
    )

    if confirmed:
        reason = (
            f"목록 구조 확인, 게시물 {len(post_links)}개 이상, "
            f"실제 게시물 접근 및 본문 확인"
        )
    else:

        reasons = []

        if semantics < 4:
            reasons.append("목록 구조 약함")

        if len(post_links) < 3:
            reasons.append(
                f"게시물 링크 부족({len(post_links)}개)"
            )

        if not post_url:
            reasons.append("실제 게시물 확인 실패")

        if body_length < 40:
            reasons.append("본문 부족")

        if not pagination:
            reasons.append("페이지네이션 미확인")

        reason = ", ".join(reasons)

    return {
        "confirmed": confirmed,
        "list_access": True,
        "post_links": len(post_links),
        "post_url": post_url,
        "post_title": post_title,
        "body_length": body_length,
        "pagination": pagination,
        "score": score,
        "reason": reason,
    }


async def process_candidate(
    session,
    semaphore,
    institution,
    candidate_url,
    candidate_type
):

    async with semaphore:

        final_list_url = ""
        result = None
        error = ""

        try:

            # -----------------------------------------
            # 1. 목록 URL
            # -----------------------------------------
            if candidate_type == "목록":

                result = await verify_list_page(
                    session,
                    candidate_url
                )

                if result["confirmed"]:
                    final_list_url = candidate_url

            # -----------------------------------------
            # 2. 상세 URL
            # -----------------------------------------
            elif candidate_type == "상세":

                list_url = await find_list_from_detail(
                    session,
                    candidate_url
                )

                if list_url:

                    result = await verify_list_page(
                        session,
                        list_url
                    )

                    if result["confirmed"]:
                        final_list_url = list_url

                else:

                    result = {
                        "confirmed": False,
                        "list_access": False,
                        "post_links": 0,
                        "post_url": "",
                        "post_title": "",
                        "body_length": 0,
                        "pagination": False,
                        "score": 0,
                        "reason": "상세페이지에서 목록 URL을 찾지 못함",
                    }

            # -----------------------------------------
            # 3. 홈페이지
            # -----------------------------------------
            elif candidate_type == "홈페이지":

                list_url = await discover_list_from_homepage(
                    session,
                    candidate_url
                )

                if list_url:

                    result = await verify_list_page(
                        session,
                        list_url
                    )

                    if result["confirmed"]:
                        final_list_url = list_url

                else:

                    result = {
                        "confirmed": False,
                        "list_access": False,
                        "post_links": 0,
                        "post_url": "",
                        "post_title": "",
                        "body_length": 0,
                        "pagination": False,
                        "score": 0,
                        "reason": "홈페이지에서 게시판 목록을 찾지 못함",
                    }

            # -----------------------------------------
            # 4. 미분류
            # -----------------------------------------
            else:

                # 먼저 목록으로 간주
                result = await verify_list_page(
                    session,
                    candidate_url
                )

                if result["confirmed"]:
                    final_list_url = candidate_url

                else:

                    # 상세일 가능성 확인
                    list_url = await find_list_from_detail(
                        session,
                        candidate_url
                    )

                    if list_url:

                        result = await verify_list_page(
                            session,
                            list_url
                        )

                        if result["confirmed"]:
                            final_list_url = list_url

        except Exception as e:

            error = str(e)

            result = {
                "confirmed": False,
                "list_access": False,
                "post_links": 0,
                "post_url": "",
                "post_title": "",
                "body_length": 0,
                "pagination": False,
                "score": 0,
                "reason": "검증 중 오류",
            }

        if result is None:

            result = {
                "confirmed": False,
                "list_access": False,
                "post_links": 0,
                "post_url": "",
                "post_title": "",
                "body_length": 0,
                "pagination": False,
                "score": 0,
                "reason": "검증 결과 없음",
            }

        if error:
            status = "오류"

        elif result["confirmed"]:
            status = "확정"

        elif result["list_access"]:
            status = "추가검증"

        else:
            status = "제외"

        return {
            "기관명": institution,
            "후보URL": candidate_url,
            "후보유형": candidate_type,
            "목록URL": final_list_url,
            "목록접속": result["list_access"],
            "게시물링크수": result["post_links"],
            "게시물URL": result["post_url"],
            "게시물접속": bool(result["post_url"]),
            "게시물제목": result["post_title"],
            "본문길이": result["body_length"],
            "페이지네이션": result["pagination"],
            "점수": result["score"],
            "결과": status,
            "사유": result["reason"],
            "오류": error,
        }


async def main():

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(
            f"입력파일이 없습니다: {INPUT_FILE}"
        )

    df = pd.read_excel(INPUT_FILE)

    # V6.3 URL 컬럼 자동 탐색
    url_col = None

    for col in df.columns:

        col_str = str(col).lower()

        if (
            "v6.3" in col_str
            and "url" in col_str
        ):
            url_col = col
            break

    if url_col is None:

        for col in df.columns:

            if "V6.3게시판URL" in str(col):
                url_col = col
                break

    if url_col is None:
        raise ValueError(
            "V6.3 게시판 URL 컬럼을 찾을 수 없습니다."
        )

    # 기관명 컬럼
    name_col = None

    for col in df.columns:

        if str(col).strip() == "기관명":
            name_col = col
            break

    if name_col is None:
        raise ValueError(
            "기관명 컬럼을 찾을 수 없습니다."
        )

    candidates = []

    for _, row in df.iterrows():

        url = row.get(url_col)

        if pd.isna(url):
            continue

        url = normalize_url(url)

        if not url:
            continue

        institution = str(
            row.get(name_col, "")
        ).strip()

        candidate_type = classify_url(url)

        candidates.append(
            (
                institution,
                url,
                candidate_type
            )
        )

    print("=" * 70)
    print("V6.4 Precision Board Validation")
    print("=" * 70)
    print(f"입력파일 : {INPUT_FILE}")
    print(f"URL 컬럼 : {url_col}")
    print(f"검증 대상 : {len(candidates)}")
    print("=" * 70)

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=False
    )

    semaphore = asyncio.Semaphore(CONCURRENCY)

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
                url,
                candidate_type
            )
            for institution, url, candidate_type
            in candidates
        ]

        results = []

        for index, task in enumerate(
            asyncio.as_completed(tasks),
            start=1
        ):

            result = await task

            results.append(result)

            print(
                f"[V6.4 검증] "
                f"{index}/{len(candidates)} | "
                f"{result['기관명']} | "
                f"{result['결과']} | "
                f"점수 {result['점수']} | "
                f"게시물 {result['게시물링크수']}개"
            )

    # 원본 순서 유지
    order = {
        institution: i
        for i, (institution, _, _)
        in enumerate(candidates)
    }

    results.sort(
        key=lambda x: order.get(
            x["기관명"],
            999999
        )
    )

    result_df = pd.DataFrame(results)

    # 원본 데이터와 합치기
    output = df.copy()

    result_map = {
        r["기관명"]: r
        for r in results
    }

    output["V6.4검증_후보URL"] = ""
    output["V6.4검증_후보유형"] = ""
    output["V6.4검증_목록URL"] = ""
    output["V6.4검증_목록접속"] = False
    output["V6.4검증_게시물링크수"] = 0
    output["V6.4검증_게시물URL"] = ""
    output["V6.4검증_게시물접속"] = False
    output["V6.4검증_게시물제목"] = ""
    output["V6.4검증_본문길이"] = 0
    output["V6.4검증_페이지네이션"] = False
    output["V6.4검증_점수"] = 0
    output["V6.4검증_결과"] = ""
    output["V6.4검증_사유"] = ""
    output["V6.4검증_오류"] = ""

    for idx, row in output.iterrows():

        institution = str(
            row[name_col]
        ).strip()

        if institution not in result_map:
            continue

        r = result_map[institution]

        output.at[
            idx,
            "V6.4검증_후보URL"
        ] = r["후보URL"]

        output.at[
            idx,
            "V6.4검증_후보유형"
        ] = r["후보유형"]

        output.at[
            idx,
            "V6.4검증_목록URL"
        ] = r["목록URL"]

        output.at[
            idx,
            "V6.4검증_목록접속"
        ] = r["목록접속"]

        output.at[
            idx,
            "V6.4검증_게시물링크수"
        ] = r["게시물링크수"]

        output.at[
            idx,
            "V6.4검증_게시물URL"
        ] = r["게시물URL"]

        output.at[
            idx,
            "V6.4검증_게시물접속"
        ] = r["게시물접속"]

        output.at[
            idx,
            "V6.4검증_게시물제목"
        ] = r["게시물제목"]

        output.at[
            idx,
            "V6.4검증_본문길이"
        ] = r["본문길이"]

        output.at[
            idx,
            "V6.4검증_페이지네이션"
        ] = r["페이지네이션"]

        output.at[
            idx,
            "V6.4검증_점수"
        ] = r["점수"]

        output.at[
            idx,
            "V6.4검증_결과"
        ] = r["결과"]

        output.at[
            idx,
            "V6.4검증_사유"
        ] = r["사유"]

        output.at[
            idx,
            "V6.4검증_오류"
        ] = r["오류"]

    output.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print()
    print("=" * 70)
    print("V6.4 검증 완료")
    print("=" * 70)

    print(
        f"확정       : "
        f"{sum(x['결과'] == '확정' for x in results)}"
    )

    print(
        f"추가검증   : "
        f"{sum(x['결과'] == '추가검증' for x in results)}"
    )

    print(
        f"제외       : "
        f"{sum(x['결과'] == '제외' for x in results)}"
    )

    print(
        f"오류       : "
        f"{sum(x['결과'] == '오류' for x in results)}"
    )

    print(
        f"결과파일   : {OUTPUT_FILE}"
    )

    print("=" * 70)


if __name__ == "__main__":
    asyncio.run(main())
