import asyncio
import re
from urllib.parse import urljoin, urlparse

import aiohttp
import pandas as pd
from bs4 import BeautifulSoup


# ============================================================
# 기본 설정
# ============================================================

INPUT = "boards_v6.xlsx"
OUTPUT = "boards_v6_3_analysis.xlsx"

CONCURRENCY = 8
TIMEOUT = 25
MAX_PAGES = 35
MAX_DEPTH = 2
MAX_LINKS_PER_PAGE = 120

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    )
}


# ============================================================
# 게시판 관련 키워드
# ============================================================

BOARD_WORDS = [
    "공지",
    "공지사항",
    "알림",
    "소식",
    "뉴스",
    "자료실",
    "게시판",
    "공고",
    "채용",
    "입찰",
    "보도자료",
    "새소식",
    "고시",
    "공시",
    "정보마당",
    "홍보",
    "전자민원",
    "고객센터",
]


# ============================================================
# JS / API 관련 키워드
# ============================================================

API_WORDS = [
    "api",
    "ajax",
    "json",
    "rest",
    "graphql",
    "axios",
    "fetch(",
    "xmlhttprequest",
    "$.ajax",
    "$.get(",
    "$.post(",
]


# ============================================================
# CMS 탐지 패턴
# ============================================================

CMS_PATTERNS = {
    "K2Web": [
        r"k2web",
        r"webcontents",
    ],
    "Drupal": [
        r"drupal",
    ],
    "WordPress": [
        r"wp-content",
        r"wordpress",
    ],
    "Joomla": [
        r"joomla",
    ],
    "그누보드": [
        r"gnuboard",
        r"bbs/board\.php",
    ],
    "XpressEngine": [
        r"xpressengine",
        r"xe/",
    ],
    "Rhymix": [
        r"rhymix",
    ],
}


# ============================================================
# URL 유틸
# ============================================================

def clean_url(url):
    if not url:
        return ""

    return str(url).strip()


def same_host(a, b):

    try:

        host_a = urlparse(a).netloc.lower().split(":")[0]
        host_b = urlparse(b).netloc.lower().split(":")[0]

        return host_a == host_b

    except:

        return False


# ============================================================
# 게시판 URL 점수
# ============================================================

def score_link(text, href):

    text = str(text or "")
    href = str(href or "")

    value = (text + " " + href).lower()

    score = 0

    for word in BOARD_WORDS:

        if word.lower() in value:

            score += 3


    for word in [
        "list",
        "board",
        "bbs",
        "notice",
        "news",
        "article",
        "view",
        "detail",
        "content",
    ]:

        if word in value:

            score += 2


    for word in [
        "login",
        "privacy",
        "sitemap",
        "search",
        "member",
    ]:

        if word in value:

            score -= 3


    return score


# ============================================================
# 기존 V6.2 결과를 기준으로 1차 공략방법 분류
# ============================================================

def classify_method(row):

    status = str(row.get("접속상태", ""))
    cms = str(row.get("CMS", ""))
    iframe = str(row.get("iframe", ""))
    js = str(row.get("JS", ""))
    fail = str(row.get("실패단계", ""))

    try:
        old_board = int(row.get("게시판후보수", 0) or 0)
    except:
        old_board = 0

    try:
        old_post = int(row.get("게시물후보수", 0) or 0)
    except:
        old_post = 0


    if (
        "실패" in status
        or "Timeout" in status
        or "실패" in fail
    ):

        return "A_접속실패_재시도"


    if old_board > 0 or old_post > 0:

        return "B_기존후보_정밀검증"


    if "k2web" in cms.lower():

        return "C_K2Web_정밀탐색"


    if (
        "있음" in iframe.lower()
        or "true" in iframe.lower()
        or "yes" in iframe.lower()
    ):

        return "D_iframe_실제주소추적"


    if any(
        word in js.lower()
        for word in [
            "높음",
            "high",
            "동적",
            "yes",
        ]
    ):

        return "E_JS_API_추적"


    return "F_메뉴_사이트맵_URL패턴"


# ============================================================
# HTTP 요청
# ============================================================

async def fetch(session, url):

    try:

        async with session.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
            allow_redirects=True,
            ssl=False,
        ) as response:

            html = await response.text(
                errors="ignore"
            )

            return (
                response.status,
                str(response.url),
                html,
                "",
            )

    except Exception as e:

        return (
            None,
            url,
            "",
            f"{type(e).__name__}: {e}",
        )


# ============================================================
# 기관 하나 분석
# ============================================================

async def analyze_one(
    session,
    semaphore,
    record,
):

    async with semaphore:

        homepage = clean_url(
            record.get("홈페이지", "")
        )

        result = dict(record)

        result.update({

            "V6.3상태": "",

            "V6.3탐색페이지": 0,

            "V6.3게시판후보": 0,

            "V6.3게시물후보": 0,

            "V6.3iframe": 0,

            "V6.3API후보": 0,

            "V6.3CMS": "",

            "V6.3게시판URL": "",

            "V6.3게시물URL": "",

            "V6.3iframeURL": "",

            "V6.3API예시": "",

            "V6.3탐색방법": classify_method(record),

            "V6.3신뢰도": "미발견",

            "V6.3오류": "",

        })


        # 홈페이지 없음

        if (
            not homepage
            or homepage.lower() == "nan"
        ):

            result["V6.3상태"] = "홈페이지없음"

            return result


        # 탐색 큐

        queue = [
            (homepage, 0)
        ]

        seen = set()

        board_urls = []

        post_urls = []

        iframe_urls = []

        api_urls = []

        cms_hits = set()


        # ====================================================
        # BFS 탐색
        # ====================================================

        while (
            queue
            and len(seen) < MAX_PAGES
        ):

            url, depth = queue.pop(0)


            if url in seen:

                continue


            seen.add(url)


            status, final_url, html, error = (
                await fetch(
                    session,
                    url,
                )
            )


            # 오류

            if error:

                if not result["V6.3오류"]:

                    result["V6.3오류"] = error

                continue


            # 첫 페이지 상태

            if len(seen) == 1:

                result["V6.3상태"] = (
                    f"정상:{status}"
                )


            if status >= 400:

                continue


            soup = BeautifulSoup(
                html,
                "html.parser",
            )


            html_lower = html.lower()


            # =================================================
            # CMS 탐지
            # =================================================

            for (
                cms_name,
                patterns,
            ) in CMS_PATTERNS.items():

                for pattern in patterns:

                    if re.search(
                        pattern,
                        html_lower,
                    ):

                        cms_hits.add(
                            cms_name
                        )

                        break


            # =================================================
            # iframe / frame
            # =================================================

            for tag in soup.find_all(
                ["iframe", "frame"]
            ):

                src = tag.get("src")

                if not src:

                    continue


                target = urljoin(
                    final_url,
                    src,
                )


                if target not in iframe_urls:

                    iframe_urls.append(
                        target
                    )


                # iframe도 탐색 큐에 추가
                # 단, 외부 사이트도 추적 가능

                if target not in seen:

                    queue.append(
                        (
                            target,
                            depth + 1,
                        )
                    )


            # =================================================
            # JS / API 탐지
            # =================================================

            for script in soup.find_all(
                "script"
            ):

                src = script.get("src")

                inline_code = (
                    script.string
                    or ""
                )


                blob = (
                    str(src or "")
                    + " "
                    + inline_code[:20000]
                ).lower()


                if any(
                    word in blob
                    for word in API_WORDS
                ):

                    # 외부 JS 파일

                    if src:

                        target = urljoin(
                            final_url,
                            src,
                        )

                        if (
                            target
                            not in api_urls
                        ):

                            api_urls.append(
                                target
                            )


                    # 코드 내부 API URL

                    patterns = [

                        r'https?://[^"\']+',

                        r'/[^"\']*(?:api|ajax|json)[^"\']*',

                        r'["\']([^"\']*(?:api|ajax|json)[^"\']*)["\']',

                    ]


                    for pattern in patterns:

                        try:

                            matches = re.findall(
                                pattern,
                                inline_code,
                                re.I,
                            )

                        except:

                            matches = []


                        for match in matches:

                            if isinstance(
                                match,
                                tuple,
                            ):

                                match = match[0]


                            target = urljoin(
                                final_url,
                                str(match),
                            )


                            if (
                                target
                                not in api_urls
                            ):

                                api_urls.append(
                                    target
                                )


            # =================================================
            # 링크 분석
            # =================================================

            links = []


            for a in soup.find_all(
                "a",
                href=True,
            ):

                href = a.get("href")

                if not href:

                    continue


                target = urljoin(
                    final_url,
                    href,
                )


                if not target.startswith(
                    (
                        "http://",
                        "https://",
                    )
                ):

                    continue


                text = a.get_text(
                    " ",
                    strip=True,
                )


                score = score_link(
                    text,
                    target,
                )


                # ---------------------------------------------
                # 게시판 후보
                # ---------------------------------------------

                if score >= 3:

                    board_urls.append(
                        (
                            score,
                            target,
                            text[:100],
                        )
                    )


                # ---------------------------------------------
                # 게시물 후보
                # ---------------------------------------------

                path_query = (
                    urlparse(target).path.lower()
                    + " "
                    + urlparse(target).query.lower()
                )


                post_patterns = [

                    "view",

                    "detail",

                    "article",

                    "read",

                    "contentview",

                    "boardview",

                    "view.do",

                    "read.do",

                    "seq=",

                    "idx=",

                    "ntt",

                    "ntt_id",

                    "bbsno=",

                    "boardno=",

                    "articleid=",

                ]


                if any(
                    pattern
                    in path_query
                    for pattern
                    in post_patterns
                ):

                    post_urls.append(
                        (
                            score + 1,
                            target,
                            text[:100],
                        )
                    )


                # ---------------------------------------------
                # 다음 페이지
                # ---------------------------------------------

                if (
                    same_host(
                        homepage,
                        target,
                    )
                    and depth < MAX_DEPTH
                ):

                    links.append(
                        (
                            score,
                            target,
                        )
                    )


            # 점수 높은 링크 우선

            links.sort(
                reverse=True
            )


            for (
                _,
                target,
            ) in links[
                :MAX_LINKS_PER_PAGE
            ]:

                if target not in seen:

                    queue.append(
                        (
                            target,
                            depth + 1,
                        )
                    )


        # ====================================================
        # 결과 정리
        # ====================================================

        board_urls = sorted(
            set(board_urls),
            reverse=True,
        )


        post_urls = sorted(
            set(post_urls),
            reverse=True,
        )


        result[
            "V6.3탐색페이지"
        ] = len(seen)


        result[
            "V6.3게시판후보"
        ] = len(board_urls)


        result[
            "V6.3게시물후보"
        ] = len(post_urls)


        result[
            "V6.3iframe"
        ] = len(iframe_urls)


        result[
            "V6.3API후보"
        ] = len(api_urls)


        result[
            "V6.3CMS"
        ] = ", ".join(
            sorted(cms_hits)
        )


        if board_urls:

            result[
                "V6.3게시판URL"
            ] = board_urls[0][1]


        if post_urls:

            result[
                "V6.3게시물URL"
            ] = post_urls[0][1]


        if iframe_urls:

            result[
                "V6.3iframeURL"
            ] = iframe_urls[0]


        if api_urls:

            result[
                "V6.3API예시"
            ] = api_urls[0]


        # ====================================================
        # 신뢰도
        # ====================================================

        if (
            board_urls
            and post_urls
        ):

            result[
                "V6.3신뢰도"
            ] = "높음"


        elif board_urls:

            result[
                "V6.3신뢰도"
            ] = "중간"


        elif (
            post_urls
            or iframe_urls
            or api_urls
        ):

            result[
                "V6.3신뢰도"
            ] = "낮음-추가검증"


        elif str(
            result["V6.3상태"]
        ).startswith("정상"):

            result[
                "V6.3신뢰도"
            ] = "미발견"


        return result


# ============================================================
# 메인
# ============================================================

async def main():

    print("=" * 80)

    print(
        "V6.3 정밀 게시판 탐색 시작"
    )

    print("=" * 80)


    # 입력 파일 확인

    if not Path(INPUT).exists():

        raise FileNotFoundError(
            f"{INPUT} 파일이 없습니다."
        )


    df = pd.read_excel(
        INPUT
    ).fillna("")


    print(
        f"분석 기관 : {len(df)}개"
    )


    records = df.to_dict(
        "records"
    )


    semaphore = asyncio.Semaphore(
        CONCURRENCY
    )


    timeout = aiohttp.ClientTimeout(
        total=TIMEOUT
    )


    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=False,
    )


    async with aiohttp.ClientSession(
        timeout=timeout,
        connector=connector,
    ) as session:


        tasks = [

            analyze_one(
                session,
                semaphore,
                record,
            )

            for record
            in records

        ]


        results = []


        completed = 0


        for task in asyncio.as_completed(
            tasks
        ):

            try:

                result = await task

                results.append(
                    result
                )

            except Exception as e:

                print(
                    "[ERROR]",
                    repr(e)
                )


            completed += 1


            if (
                completed % 10 == 0
                or completed == len(tasks)
            ):

                print(
                    f"[V6.3] "
                    f"{completed}/"
                    f"{len(tasks)} 완료"
                )


    # ========================================================
    # 결과 저장
    # ========================================================

    out = pd.DataFrame(
        results
    )


    original_columns = list(
        df.columns
    )


    new_columns = [

        c

        for c in out.columns

        if c not in original_columns

    ]


    out = out[
        original_columns
        + new_columns
    ]


    out.to_excel(
        OUTPUT,
        index=False,
    )


    # ========================================================
    # 요약
    # ========================================================

    print()

    print("=" * 80)

    print(
        "V6.3 정밀탐색 완료"
    )

    print("=" * 80)


    print(
        f"기관 수 : {len(out)}"
    )


    print(
        f"결과 파일 : {OUTPUT}"
    )


    print()

    print(
        "[신뢰도별 결과]"
    )


    print(
        out[
            "V6.3신뢰도"
        ]
        .value_counts(
            dropna=False
        )
        .to_string()
    )


    print()

    print(
        "[탐색방법별 결과]"
    )


    print(
        out[
            "V6.3탐색방법"
        ]
        .value_counts(
            dropna=False
        )
        .to_string()
    )


    print()

    print("=" * 80)


if __name__ == "__main__":

    asyncio.run(
        main()
    )
