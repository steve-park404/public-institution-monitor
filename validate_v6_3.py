import asyncio
import re
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

import aiohttp
import pandas as pd
from bs4 import BeautifulSoup


# ============================================================
# 설정
# ============================================================

INPUT = "boards_v6_3_analysis.xlsx"
OUTPUT = "boards_v6_3_validation.xlsx"

CONCURRENCY = 8
REQUEST_TIMEOUT = 20
INSTITUTION_TIMEOUT = 45

MAX_POST_LINKS = 8
MAX_LIST_LINKS = 10

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
}


# ============================================================
# 컬럼 자동 탐색
# ============================================================

def find_column(df, aliases):
    cols = list(df.columns)

    # 정확 일치
    for alias in aliases:
        for col in cols:
            if str(col).strip() == alias:
                return col

    # 부분 일치
    for alias in aliases:
        for col in cols:
            if alias.lower() in str(col).lower():
                return col

    return None


def detect_columns(df):
    name_col = find_column(
        df,
        ["기관명", "기관", "기관명칭", "기관 이름"]
    )

    candidate_col = find_column(
        df,
        [
            "V6.3게시판URL",
            "게시판URL",
            "V6.3 게시판 URL",
            "V6.3게시판",
        ]
    )

    confidence_col = find_column(
        df,
        [
            "V6.3신뢰도",
            "신뢰도",
        ]
    )

    homepage_col = find_column(
        df,
        [
            "홈페이지",
            "URL",
            "홈페이지URL",
        ]
    )

    print()
    print("=" * 70)
    print("입력 파일 컬럼 확인")
    print("=" * 70)

    for c in df.columns:
        print(" -", c)

    print()
    print("기관명       :", name_col)
    print("게시판 후보  :", candidate_col)
    print("신뢰도       :", confidence_col)
    print("홈페이지     :", homepage_col)
    print("=" * 70)

    if name_col is None:
        raise ValueError("기관명 컬럼을 찾을 수 없습니다.")

    if candidate_col is None:
        raise ValueError("게시판 URL 컬럼을 찾을 수 없습니다.")

    return name_col, candidate_col, confidence_col, homepage_col


# ============================================================
# URL
# ============================================================

def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if url.lower() in ["nan", "none", "null"]:
        return ""

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def clean_url(url):
    try:
        p = urlparse(url)

        # fragment 제거
        return p._replace(fragment="").geturl()

    except Exception:
        return url


def same_host(a, b):
    try:
        ha = urlparse(a).netloc.lower()
        hb = urlparse(b).netloc.lower()

        ha = ha.split(":")[0]
        hb = hb.split(":")[0]

        return ha == hb

    except Exception:
        return False


# ============================================================
# URL 판별
# ============================================================

BOARD_WORDS = [
    "board",
    "bbs",
    "notice",
    "news",
    "community",
    "list",
    "article",
    "press",
    "recruit",
    "notification",
    "plaza",
    "communication",
    "post",
]

POST_WORDS = [
    "view",
    "detail",
    "articleNo",
    "nttId",
    "boardId",
    "list_no",
    "seq",
    "idx",
    "wr_id",
    "article",
]


def looks_like_board(url):
    u = url.lower()

    score = 0

    for word in BOARD_WORDS:
        if word in u:
            score += 1

    if "list" in u:
        score += 2

    if "board" in u or "bbs" in u:
        score += 2

    return score >= 2


def looks_like_post(url):
    u = url.lower()

    score = 0

    for word in POST_WORDS:
        if word.lower() in u:
            score += 1

    query = parse_qs(urlparse(url).query)

    for key in [
        "articleno",
        "nttid",
        "boardid",
        "list_no",
        "idx",
        "seq",
        "wr_id",
    ]:
        if key.lower() in [k.lower() for k in query.keys()]:
            score += 2

    if "mode=view" in u:
        score += 3

    if "/view" in u or "/detail" in u:
        score += 2

    return score >= 2


# ============================================================
# HTML 분석
# ============================================================

def get_text(soup):
    return soup.get_text(" ", strip=True)


def extract_links(base_url, soup):

    result = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith(("#", "javascript:", "mailto:", "tel:")):
            continue

        url = clean_url(urljoin(base_url, href))

        if not url.startswith(("http://", "https://")):
            continue

        text = a.get_text(" ", strip=True)

        result.append(
            {
                "url": url,
                "text": text,
            }
        )

    return result


def find_title(soup):

    # title
    if soup.title:
        title = soup.title.get_text(" ", strip=True)

        if title:
            return title[:300]

    # h1
    for tag in soup.find_all(["h1", "h2"]):

        text = tag.get_text(" ", strip=True)

        if 2 <= len(text) <= 200:
            return text

    return ""


def find_post_content(soup):

    # 흔한 본문 selector
    selectors = [
        ".view_content",
        ".board_view",
        ".board-view",
        ".bbs_view",
        ".bbs-view",
        ".article_view",
        ".article-view",
        ".view_cont",
        ".view-cont",
        ".contents",
        ".content",
        "#content",
        "#contents",
        ".txt",
        ".article",
        "article",
    ]

    candidates = []

    for selector in selectors:

        try:
            for tag in soup.select(selector):

                text = tag.get_text(" ", strip=True)

                if len(text) >= 40:
                    candidates.append(text)

        except Exception:
            pass

    if not candidates:
        return ""

    candidates.sort(key=len, reverse=True)

    return candidates[0][:5000]


def extract_title_from_post(soup):

    # 게시물 제목에 자주 사용되는 selector
    selectors = [
        ".view_title",
        ".view-title",
        ".board_title",
        ".board-title",
        ".bbs_title",
        ".bbs-title",
        ".article_title",
        ".article-title",
        "h1",
        "h2",
        ".subject",
        ".title",
    ]

    candidates = []

    for selector in selectors:

        try:
            for tag in soup.select(selector):

                text = tag.get_text(" ", strip=True)

                if 2 <= len(text) <= 300:
                    candidates.append(text)

        except Exception:
            pass

    if candidates:
        candidates.sort(key=len)

        return candidates[0]

    return find_title(soup)


# ============================================================
# HTTP
# ============================================================

async def fetch(session, url):

    try:

        timeout = aiohttp.ClientTimeout(
            total=REQUEST_TIMEOUT,
            connect=10,
        )

        async with session.get(
            url,
            headers=HEADERS,
            timeout=timeout,
            allow_redirects=True,
            ssl=False,
        ) as response:

            content_type = response.headers.get(
                "Content-Type",
                ""
            )

            text = await response.text(
                errors="ignore"
            )

            return {
                "ok": True,
                "status": response.status,
                "url": str(response.url),
                "text": text,
                "content_type": content_type,
                "error": "",
            }

    except Exception as e:

        return {
            "ok": False,
            "status": 0,
            "url": url,
            "text": "",
            "content_type": "",
            "error": str(e),
        }


# ============================================================
# 게시판 검증
# ============================================================

async def validate_candidate(session, homepage, board_url):

    result = {
        "접속성공": "N",
        "HTTP상태": "",
        "최종URL": "",
        "게시판판정": "미확인",
        "게시물링크수": 0,
        "게시물URL": "",
        "게시물접속": "N",
        "게시물제목": "",
        "게시물본문길이": 0,
        "목록페이지판정": "N",
        "검증점수": 0,
        "검증결과": "제외",
        "검증사유": "",
        "오류": "",
    }

    board_url = normalize_url(board_url)

    if not board_url:
        result["검증사유"] = "게시판 URL 없음"
        return result

    r = await fetch(session, board_url)

    result["HTTP상태"] = r["status"]
    result["최종URL"] = r["url"]

    if not r["ok"]:

        result["오류"] = r["error"]
        result["검증사유"] = "게시판 접속 실패"

        return result

    result["접속성공"] = "Y"

    if r["status"] >= 400:

        result["검증사유"] = (
            f"HTTP 오류 {r['status']}"
        )

        return result

    html = r["text"]

    if len(html) < 100:

        result["검증사유"] = "HTML 내용 부족"

        return result

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    text = get_text(soup)

    score = 0

    # --------------------------------------------------------
    # 게시판 페이지 판정
    # --------------------------------------------------------

    links = extract_links(
        r["url"],
        soup
    )

    post_links = []

    for item in links:

        url = item["url"]
        link_text = item["text"]

        if not same_host(homepage, url):
            # 외부 서브도메인 게시판은 허용
            hp = urlparse(homepage)
            up = urlparse(url)

            if not (
                hp.netloc.endswith(
                    "." + up.netloc
                )
                or up.netloc.endswith(
                    "." + hp.netloc
                )
            ):
                continue

        if looks_like_post(url):

            post_links.append(
                (
                    url,
                    link_text
                )
            )

    # 중복 제거
    seen = set()
    unique_posts = []

    for url, text_ in post_links:

        if url in seen:
            continue

        seen.add(url)

        unique_posts.append(
            (url, text_)
        )

        if len(unique_posts) >= MAX_POST_LINKS:
            break

    result["게시물링크수"] = len(unique_posts)

    if unique_posts:
        score += 3
        result["목록페이지판정"] = "Y"

    # 게시판 관련 텍스트
    board_text_words = [
        "공지",
        "게시판",
        "등록일",
        "조회",
        "작성일",
        "번호",
        "제목",
        "첨부파일",
        "새소식",
        "알림",
        "채용",
    ]

    matched_words = sum(
        1
        for word in board_text_words
        if word in text
    )

    if matched_words >= 2:
        score += 2

    if matched_words >= 4:
        score += 1

    # URL 자체가 게시판 형태
    if looks_like_board(r["url"]):
        score += 1

    # --------------------------------------------------------
    # 실제 게시물 접속
    # --------------------------------------------------------

    post_url = ""

    if unique_posts:

        # 링크 텍스트가 제목처럼 보이는 링크 우선
        for url, link_text in unique_posts:

            if 3 <= len(link_text) <= 150:
                post_url = url
                break

        if not post_url:
            post_url = unique_posts[0][0]

    if post_url:

        result["게시물URL"] = post_url

        pr = await fetch(
            session,
            post_url
        )

        if pr["ok"] and pr["status"] < 400:

            result["게시물접속"] = "Y"

            psoup = BeautifulSoup(
                pr["text"],
                "html.parser"
            )

            title = extract_title_from_post(
                psoup
            )

            content = find_post_content(
                psoup
            )

            result["게시물제목"] = title[:300]
            result["게시물본문길이"] = len(content)

            if title:
                score += 2

            if len(content) >= 40:
                score += 3

            elif len(content) >= 15:
                score += 1

        else:

            if pr["error"]:
                result["오류"] = pr["error"]

    # --------------------------------------------------------
    # 최종 판정
    # --------------------------------------------------------

    result["검증점수"] = score

    if (
        result["접속성공"] == "Y"
        and result["목록페이지판정"] == "Y"
        and result["게시물접속"] == "Y"
        and (
            result["게시물제목"]
            or result["게시물본문길이"] >= 40
        )
        and score >= 7
    ):

        result["검증결과"] = "확정"
        result["게시판판정"] = "실제 게시판"

        result["검증사유"] = (
            "게시판 목록 및 실제 게시물 "
            "접속·제목/본문 확인"
        )

    elif (
        result["접속성공"] == "Y"
        and score >= 4
    ):

        result["검증결과"] = "추가검증"
        result["게시판판정"] = "게시판 후보"

        result["검증사유"] = (
            "게시판 구조는 확인되나 "
            "게시물 또는 본문 검증 부족"
        )

    else:

        result["검증결과"] = "제외"
        result["게시판판정"] = "게시판 아님"

        if not result["검증사유"]:
            result["검증사유"] = (
                "실제 게시판 구조 확인 실패"
            )

    return result


# ============================================================
# 기관 1개
# ============================================================

async def validate_one(
    session,
    row,
    idx,
    total,
    name_col,
    candidate_col,
    homepage_col,
):

    name = str(
        row.get(name_col, "")
    ).strip()

    board_url = str(
        row.get(candidate_col, "")
    ).strip()

    homepage = ""

    if homepage_col:
        homepage = normalize_url(
            row.get(homepage_col, "")
        )

    if not homepage:
        homepage = board_url

    print(
        f"[V6.3 검증] {idx}/{total} {name}"
    )

    try:

        async with asyncio.timeout(
            INSTITUTION_TIMEOUT
        ):

            validation = await validate_candidate(
                session,
                homepage,
                board_url
            )

    except Exception as e:

        validation = {
            "접속성공": "N",
            "HTTP상태": "",
            "최종URL": "",
            "게시판판정": "오류",
            "게시물링크수": 0,
            "게시물URL": "",
            "게시물접속": "N",
            "게시물제목": "",
            "게시물본문길이": 0,
            "목록페이지판정": "N",
            "검증점수": 0,
            "검증결과": "오류",
            "검증사유": "기관 검증 시간초과/오류",
            "오류": str(e),
        }

    return validation


# ============================================================
# 메인
# ============================================================

async def main():

    input_path = Path(INPUT)

    if not input_path.exists():
        raise FileNotFoundError(
            f"입력 파일 없음: {INPUT}"
        )

    print()
    print("=" * 70)
    print("V6.3 게시판 후보 자동 검증")
    print("=" * 70)
    print("입력 :", INPUT)
    print("출력 :", OUTPUT)
    print("=" * 70)

    df = pd.read_excel(
        input_path
    )

    print(
        f"전체 입력 행: {len(df)}"
    )

    (
        name_col,
        candidate_col,
        confidence_col,
        homepage_col,
    ) = detect_columns(df)

    # --------------------------------------------------------
    # 게시판 후보만 추출
    # --------------------------------------------------------

    targets = []

    for idx, row in df.iterrows():

        candidate = row.get(
            candidate_col,
            ""
        )

        candidate = str(candidate).strip()

        if (
            not candidate
            or candidate.lower()
            in ["nan", "none", "null"]
        ):
            continue

        confidence = ""

        if confidence_col:
            confidence = str(
                row.get(confidence_col, "")
            ).strip()

        # V6.3에서 발견된 후보만
        targets.append(
            (
                idx,
                row,
                confidence,
                candidate,
            )
        )

    print()
    print(
        f"검증 대상 후보: {len(targets)}개"
    )

    if not targets:

        print(
            "검증할 게시판 후보가 없습니다."
        )

        # 결과 파일은 만들어 둔다.
        df.to_excel(
            OUTPUT,
            index=False
        )

        return

    # --------------------------------------------------------
    # 세션
    # --------------------------------------------------------

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        limit_per_host=3,
        ssl=False,
    )

    async with aiohttp.ClientSession(
        connector=connector,
        headers=HEADERS,
    ) as session:

        semaphore = asyncio.Semaphore(
            CONCURRENCY
        )

        async def worker(item):

            async with semaphore:

                idx, row, confidence, candidate = item

                result = await validate_one(
                    session,
                    row,
                    idx + 1,
                    len(targets),
                    name_col,
                    candidate_col,
                    homepage_col,
                )

                return (
                    idx,
                    confidence,
                    candidate,
                    result,
                )

        tasks = [
            worker(item)
            for item in targets
        ]

        results = await asyncio.gather(
            *tasks,
            return_exceptions=False
        )

    # --------------------------------------------------------
    # 결과 컬럼
    # --------------------------------------------------------

    result_rows = []

    for (
        idx,
        confidence,
        candidate,
        result,
    ) in results:

        row = df.loc[idx].to_dict()

        row["V6.3검증_기존신뢰도"] = confidence
        row["V6.3검증_후보URL"] = candidate

        for key, value in result.items():

            row[
                "V6.3검증_" + key
            ] = value

        result_rows.append(row)

    result_df = pd.DataFrame(
        result_rows
    )

    # --------------------------------------------------------
    # 결과 정렬
    # --------------------------------------------------------

    order = {
        "확정": 0,
        "추가검증": 1,
        "제외": 2,
        "오류": 3,
    }

    result_df["_sort"] = (
        result_df[
            "V6.3검증_검증결과"
        ]
        .map(order)
        .fillna(9)
    )

    result_df = result_df.sort_values(
        "_sort"
    ).drop(
        columns=["_sort"]
    )

    # --------------------------------------------------------
    # 저장
    # --------------------------------------------------------

    result_df.to_excel(
        OUTPUT,
        index=False
    )

    # --------------------------------------------------------
    # 통계
    # --------------------------------------------------------

    counts = (
        result_df[
            "V6.3검증_검증결과"
        ]
        .value_counts()
    )

    print()
    print("=" * 70)
    print("V6.3 게시판 검증 완료")
    print("=" * 70)

    print(
        "검증 대상 :", len(result_df)
    )

    print(
        "확정       :",
        counts.get("확정", 0)
    )

    print(
        "추가검증   :",
        counts.get("추가검증", 0)
    )

    print(
        "제외       :",
        counts.get("제외", 0)
    )

    print(
        "오류       :",
        counts.get("오류", 0)
    )

    print()
    print("===== 확정 게시판 =====")

    confirmed = result_df[
        result_df[
            "V6.3검증_검증결과"
        ] == "확정"
    ]

    for _, r in confirmed.iterrows():

        print(
            "-",
            r[name_col],
            "|",
            r["V6.3검증_후보URL"],
            "| 점수:",
            r["V6.3검증_검증점수"]
        )

    print()
    print("===== 추가검증 =====")

    extra = result_df[
        result_df[
            "V6.3검증_검증결과"
        ] == "추가검증"
    ]

    for _, r in extra.iterrows():

        print(
            "-",
            r[name_col],
            "|",
            r["V6.3검증_후보URL"],
            "| 점수:",
            r["V6.3검증_검증점수"]
        )

    print()
    print("=" * 70)
    print(
        "결과파일:",
        OUTPUT
    )
    print("=" * 70)


if __name__ == "__main__":

    asyncio.run(main())
