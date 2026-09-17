import asyncio
import re
from pathlib import Path
from urllib.parse import urljoin, urlparse

import aiohttp
import pandas as pd
from bs4 import BeautifulSoup


# ============================================================
# V6.3 설정
# ============================================================

INPUT = "boards_v6.xlsx"
OUTPUT = "boards_v6_3_analysis.xlsx"

CONCURRENCY = 8
REQUEST_TIMEOUT = 15

# 기관 1개당 최대 페이지
MAX_PAGES = 25

# 탐색 깊이
MAX_DEPTH = 2

# 한 페이지에서 추가 탐색할 최대 링크
MAX_LINKS_PER_PAGE = 60

# 기관 하나 전체 최대 실행시간
INSTITUTION_TIMEOUT = 60


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,*/*;q=0.8"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.7,en;q=0.5",
}


# ============================================================
# 게시판 키워드
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
    "사업공고",
    "입찰공고",
    "채용공고",
    "언론보도",
    "알림마당",
    "공지/공고",
]


# ============================================================
# 게시판 URL 패턴
# ============================================================

BOARD_URL_WORDS = [
    "board",
    "bbs",
    "notice",
    "news",
    "announce",
    "announcement",
    "data",
    "archive",
    "community",
    "list",
    "article",
    "content",
    "pds",
    "webzine",
    "press",
    "job",
    "recruit",
    "information",
]


# ============================================================
# 게시물 URL 패턴
# ============================================================

POST_URL_WORDS = [
    "view",
    "detail",
    "article",
    "read",
    "contentview",
    "boardview",
    "view.do",
    "read.do",
    "detail.do",
    "seq=",
    "idx=",
    "ntt",
    "ntt_id",
    "bbsno=",
    "boardno=",
    "articleid=",
    "article_no",
    "menu_no",
    "mode=view",
]


# ============================================================
# JS / API 패턴
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
# CMS
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
        r"/xe/",
    ],

    "Rhymix": [
        r"rhymix",
    ],
}


# ============================================================
# 공통 경로
# ============================================================

COMMON_PATHS = [

    "/sitemap.xml",
    "/robots.txt",

    "/notice",
    "/notices",
    "/notice/list",
    "/notice.do",

    "/news",
    "/news/list",
    "/news.do",

    "/board",
    "/board/list",
    "/board/notice",
    "/board/notice/list",

    "/bbs",
    "/bbs/list",
    "/bbs/board",
    "/bbs/board.php",

    "/data",
    "/data/list",

    "/pds",
    "/pds/list",

    "/community",
    "/community/notice",

    "/information",
    "/info",
    "/archive",

    "/contents/notice",
    "/contents/news",

    "/sub/notice",
    "/sub/news",

    "/ko/notice",
    "/kr/notice",

]


# ============================================================
# 숫자 변환
# ============================================================

def to_int(value):

    try:
        return int(float(value))
    except:
        return 0


# ============================================================
# 문자열
# ============================================================

def clean(value):

    if value is None:
        return ""

    if pd.isna(value):
        return ""

    return str(value).strip()


# ============================================================
# 컬럼 자동 탐색
# ============================================================

def find_column(df, candidates):

    # 정확히 일치
    for candidate in candidates:

        if candidate in df.columns:
            return candidate

    # 부분 일치
    for col in df.columns:

        col_str = str(col).strip().lower()

        for candidate in candidates:

            if candidate.lower() in col_str:
                return col

    return None


# ============================================================
# 컬럼 구조 자동 분석
# ============================================================

def detect_columns(df):

    institution_col = find_column(
        df,
        [
            "기관명",
            "기관",
            "기관명칭",
            "기관 이름",
        ],
    )


    homepage_col = find_column(
        df,
        [
            "홈페이지",
            "URL",
            "홈페이지URL",
            "홈페이지 주소",
            "주소",
        ],
    )


    board_count_col = find_column(
        df,
        [
            "게시판후보수",
            "게시판 후보수",
            "게시판후보",
            "게시판 후보",
            "게시판수",
        ],
    )


    post_count_col = find_column(
        df,
        [
            "게시물후보수",
            "게시물 후보수",
            "게시물후보",
            "게시물 후보",
            "게시물수",
        ],
    )


    board_url_col = find_column(
        df,
        [
            "게시판URL",
            "게시판 URL",
            "게시판주소",
        ],
    )


    post_url_col = find_column(
        df,
        [
            "게시물URL",
            "게시물 URL",
            "게시물주소",
        ],
    )


    status_col = find_column(
        df,
        [
            "접속상태",
            "상태",
            "접속 상태",
        ],
    )


    cms_col = find_column(
        df,
        [
            "CMS",
            "CMS/프레임워크",
            "프레임워크",
        ],
    )


    iframe_col = find_column(
        df,
        [
            "iframe",
            "iframe 사용",
        ],
    )


    js_col = find_column(
        df,
        [
            "JS",
            "JS 의존도",
            "자바스크립트",
        ],
    )


    print()
    print("=" * 80)
    print("V6.3 입력파일 컬럼 자동인식")
    print("=" * 80)

    print("기관명      :", institution_col)
    print("홈페이지    :", homepage_col)
    print("게시판후보  :", board_count_col)
    print("게시물후보  :", post_count_col)
    print("게시판URL   :", board_url_col)
    print("게시물URL   :", post_url_col)
    print("접속상태    :", status_col)
    print("CMS         :", cms_col)
    print("iframe      :", iframe_col)
    print("JS          :", js_col)

    print()
    print("전체 컬럼:")

    for col in df.columns:
        print(" -", col)

    print("=" * 80)


    if institution_col is None:

        raise ValueError(
            "기관명 컬럼을 찾을 수 없습니다."
        )


    if homepage_col is None:

        raise ValueError(
            "홈페이지/URL 컬럼을 찾을 수 없습니다."
        )


    return {
        "institution": institution_col,
        "homepage": homepage_col,
        "board_count": board_count_col,
        "post_count": post_count_col,
        "board_url": board_url_col,
        "post_url": post_url_col,
        "status": status_col,
        "cms": cms_col,
        "iframe": iframe_col,
        "js": js_col,
    }


# ============================================================
# URL 정리
# ============================================================

def normalize_url(url):

    url = clean(url)

    if not url:
        return ""

    if not url.startswith(
        (
            "http://",
            "https://",
        )
    ):

        url = "https://" + url

    return url


# ============================================================
# Host 비교
# ============================================================

def same_host(a, b):

    try:

        host_a = (
            urlparse(a)
            .netloc
            .lower()
            .split(":")[0]
        )

        host_b = (
            urlparse(b)
            .netloc
            .lower()
            .split(":")[0]
        )

        return host_a == host_b

    except:

        return False


# ============================================================
# 게시판 링크 점수
# ============================================================

def score_link(text, href):

    value = (
        clean(text)
        + " "
        + clean(href)
    ).lower()


    score = 0


    for word in BOARD_WORDS:

        if word.lower() in value:

            score += 4


    for word in BOARD_URL_WORDS:

        if word in value:

            score += 2


    for word in POST_URL_WORDS:

        if word in value:

            score += 1


    for word in [
        "login",
        "privacy",
        "sitemap",
        "search",
        "member",
        "terms",
        "copyright",
    ]:

        if word in value:

            score -= 3


    return score


# ============================================================
# 게시판 URL 여부
# ============================================================

def is_board_url(url):

    value = clean(url).lower()

    return any(
        word in value
        for word in BOARD_URL_WORDS
    )


# ============================================================
# 게시물 URL 여부
# ============================================================

def is_post_url(url):

    value = clean(url).lower()

    return any(
        word in value
        for word in POST_URL_WORDS
    )


# ============================================================
# HTTP
# ============================================================

async def fetch(
    session,
    url,
):

    try:

        async with session.get(
            url,
            headers=HEADERS,
            timeout=REQUEST_TIMEOUT,
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
# 기관 1개 분석
# ============================================================

async def analyze_one(
    session,
    semaphore,
    record,
    columns,
):

    async with semaphore:

        try:

            return await asyncio.wait_for(

                analyze_one_inner(
                    session,
                    record,
                    columns,
                ),

                timeout=INSTITUTION_TIMEOUT,

            )


        except asyncio.TimeoutError:

            result = dict(record)

            result.update({

                "V6.3상태":
                    "기관 분석시간 초과",

                "V6.3탐색페이지":
                    0,

                "V6.3게시판후보":
                    0,

                "V6.3게시물후보":
                    0,

                "V6.3iframe":
                    0,

                "V6.3API후보":
                    0,

                "V6.3CMS":
                    "",

                "V6.3게시판URL":
                    "",

                "V6.3게시물URL":
                    "",

                "V6.3iframeURL":
                    "",

                "V6.3API예시":
                    "",

                "V6.3탐색방법":
                    "시간초과",

                "V6.3信頼度":
                    "미발견",

                "V6.3신뢰도":
                    "미발견",

                "V6.3오류":
                    "INSTITUTION_TIMEOUT",

            })

            return result


# ============================================================
# 실제 분석
# ============================================================

async def analyze_one_inner(
    session,
    record,
    columns,
):

    homepage = normalize_url(
        record.get(
            columns["homepage"],
            "",
        )
    )


    result = dict(record)


    # ========================================================
    # 결과 필드
    # ========================================================

    result.update({

        "V6.3상태":
            "",

        "V6.3탐색페이지":
            0,

        "V6.3게시판후보":
            0,

        "V6.3게시물후보":
            0,

        "V6.3iframe":
            0,

        "V6.3API후보":
            0,

        "V6.3CMS":
            "",

        "V6.3게시판URL":
            "",

        "V6.3게시물URL":
            "",

        "V6.3iframeURL":
            "",

        "V6.3API예시":
            "",

        "V6.3탐색방법":
            "",

        "V6.3신뢰도":
            "미발견",

        "V6.3오류":
            "",

    })


    # ========================================================
    # 홈페이지 없음
    # ========================================================

    if not homepage:

        result[
            "V6.3상태"
        ] = "홈페이지없음"

        result[
            "V6.3탐색방법"
        ] = "홈페이지없음"

        return result


    # ========================================================
    # 큐
    # ========================================================

    queue = [
        (
            homepage,
            0,
        )
    ]


    seen = set()

    board_candidates = []

    post_candidates = []

    iframe_urls = []

    api_candidates = []

    cms_hits = set()


    # ========================================================
    # 공통 URL
    # ========================================================

    base = homepage.rstrip("/")


    for path in COMMON_PATHS:

        target = urljoin(
            base + "/",
            path.lstrip("/"),
        )

        queue.append(
            (
                target,
                1,
            )
        )


    # ========================================================
    # BFS
    # ========================================================

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


        if error:

            if not result[
                "V6.3오류"
            ]:

                result[
                    "V6.3오류"
                ] = error

            continue


        if status is None:

            continue


        if len(seen) == 1:

            result[
                "V6.3상태"
            ] = f"정상:{status}"


        if status >= 400:

            continue


        # ====================================================
        # HTML
        # ====================================================

        soup = BeautifulSoup(
            html,
            "html.parser",
        )


        html_lower = html.lower()


        # ====================================================
        # CMS
        # ====================================================

        for (
            cms_name,
            patterns,
        ) in CMS_PATTERNS.items():

            for pattern in patterns:

                try:

                    if re.search(
                        pattern,
                        html_lower,
                    ):

                        cms_hits.add(
                            cms_name
                        )

                        break

                except:

                    pass


        # ====================================================
        # iframe
        # ====================================================

        for tag in soup.find_all(
            [
                "iframe",
                "frame",
            ]
        ):

            src = tag.get(
                "src"
            )


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


            if (
                target not in seen
                and len(queue)
                < 20
            ):

                queue.append(
                    (
                        target,
                        depth + 1,
                    )
                )


        # ====================================================
        # JS / API
        # ====================================================

        for script in soup.find_all(
            "script"
        ):

            src = script.get(
                "src"
            )


            inline_code = (
                script.string
                or ""
            )


            blob = (
                clean(src)
                + " "
                + inline_code[:30000]
            ).lower()


            if any(
                word in blob
                for word in API_WORDS
            ):

                if src:

                    target = urljoin(
                        final_url,
                        src,
                    )


                    if (
                        target
                        not in api_candidates
                    ):

                        api_candidates.append(
                            target
                        )


                patterns = [

                    r'https?://[^"\']+',

                    r'["\']([^"\']*(?:/api/|/ajax/|/json/)[^"\']*)["\']',

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


                        if not match:

                            continue


                        target = urljoin(
                            final_url,
                            str(match),
                        )


                        if (
                            target
                            not in api_candidates
                        ):

                            api_candidates.append(
                                target
                            )


        # ====================================================
        # 링크
        # ====================================================

        next_links = []


        for a in soup.find_all(
            "a",
            href=True,
        ):

            href = a.get(
                "href"
            )


            if not href:

                continue


            href_lower = href.lower()


            if href_lower.startswith(
                (
                    "javascript:",
                    "mailto:",
                    "tel:",
                    "#",
                )
            ):

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


            # -----------------------------------------------
            # 게시판
            # -----------------------------------------------

            if (
                score >= 4
                or is_board_url(target)
            ):

                board_candidates.append(
                    (
                        score,
                        target,
                        text[:120],
                    )
                )


            # -----------------------------------------------
            # 게시물
            # -----------------------------------------------

            if is_post_url(
                target
            ):

                post_candidates.append(
                    (
                        score + 1,
                        target,
                        text[:120],
                    )
                )


            # -----------------------------------------------
            # 추가 탐색
            # -----------------------------------------------

            if (
                same_host(
                    homepage,
                    target,
                )
                and depth < MAX_DEPTH
                and target not in seen
            ):

                next_links.append(
                    (
                        score,
                        target,
                    )
                )


        # ====================================================
        # 점수 높은 링크 우선
        # ====================================================

        next_links.sort(
            reverse=True
        )


        for (
            score,
            target,
        ) in next_links[
            :MAX_LINKS_PER_PAGE
        ]:

            if target in seen:

                continue


            if any(
                target == item[0]
                for item in queue
            ):

                continue


            queue.append(
                (
                    target,
                    depth + 1,
                )
            )


    # ========================================================
    # 중복 제거
    # ========================================================

    board_candidates = sorted(
        set(board_candidates),
        reverse=True,
    )


    post_candidates = sorted(
        set(post_candidates),
        reverse=True,
    )


    iframe_urls = list(
        dict.fromkeys(
            iframe_urls
        )
    )


    api_candidates = list(
        dict.fromkeys(
            api_candidates
        )
    )


    # ========================================================
    # 결과
    # ========================================================

    result[
        "V6.3탐색페이지"
    ] = len(seen)


    result[
        "V6.3게시판후보"
    ] = len(
        board_candidates
    )


    result[
        "V6.3게시물후보"
    ] = len(
        post_candidates
    )


    result[
        "V6.3iframe"
    ] = len(
        iframe_urls
    )


    result[
        "V6.3API후보"
    ] = len(
        api_candidates
    )


    result[
        "V6.3CMS"
    ] = ", ".join(
        sorted(cms_hits)
    )


    if board_candidates:

        result[
            "V6.3게시판URL"
        ] = board_candidates[0][1]


    if post_candidates:

        result[
            "V6.3게시물URL"
        ] = post_candidates[0][1]


    if iframe_urls:

        result[
            "V6.3iframeURL"
        ] = iframe_urls[0]


    if api_candidates:

        result[
            "V6.3API예시"
        ] = api_candidates[0]


    # ========================================================
    # 공략방법
    # ========================================================

    old_cms = clean(
        record.get(
            columns["cms"],
            "",
        )
        if columns["cms"]
        else ""
    )


    old_iframe = clean(
        record.get(
            columns["iframe"],
            "",
        )
        if columns["iframe"]
        else ""
    )


    old_js = clean(
        record.get(
            columns["js"],
            "",
        )
        if columns["js"]
        else ""
    )


    if "K2Web" in result[
        "V6.3CMS"
    ] or "k2web" in old_cms.lower():

        method = "B_K2Web"


    elif iframe_urls or (
        "있음" in old_iframe.lower()
        or "true" in old_iframe.lower()
        or "yes" in old_iframe.lower()
    ):

        method = "C_iframe"


    elif api_candidates:

        method = "D_JS_API"


    else:

        method = "E_메뉴_사이트맵_URL"


    result[
        "V6.3탐색방법"
    ] = method


    # ========================================================
    # 신뢰도
    # ========================================================

    if (
        board_candidates
        and post_candidates
    ):

        result[
            "V6.3신뢰도"
        ] = "높음"


    elif board_candidates:

        result[
            "V6.3신뢰도"
        ] = "중간"


    elif (
        post_candidates
        or iframe_urls
        or api_candidates
    ):

        result[
            "V6.3신뢰도"
        ] = "낮음-추가검증"


    else:

        result[
            "V6.3신뢰도"
        ] = "미발견"


    return result


# ============================================================
# 메인
# ============================================================

async def main():

    print("=" * 80)
    print("V6.3 미발견 기관 정밀탐색")
    print("=" * 80)


    # ========================================================
    # 파일
    # ========================================================

    if not Path(INPUT).exists():

        raise FileNotFoundError(
            f"{INPUT} 파일이 없습니다."
        )


    df = pd.read_excel(
        INPUT
    ).fillna("")


    print(
        f"V6.2 전체 기관 : {len(df)}개"
    )


    # ========================================================
    # 컬럼 자동 인식
    # ========================================================

    columns = detect_columns(
        df
    )


    # ========================================================
    # 미발견 기관 선별
    #
    # 컬럼이 있으면 숫자로 판단
    # 컬럼이 없으면 URL/후보 존재 여부로 판단
    # ========================================================

    target_rows = []


    for _, row in df.iterrows():

        # ----------------------------------------------------
        # 게시판 후보 수
        # ----------------------------------------------------

        if columns["board_count"]:

            board_count = to_int(
                row.get(
                    columns["board_count"],
                    0,
                )
            )

        else:

            board_count = 0


        # ----------------------------------------------------
        # 게시물 후보 수
        # ----------------------------------------------------

        if columns["post_count"]:

            post_count = to_int(
                row.get(
                    columns["post_count"],
                    0,
                )
            )

        else:

            post_count = 0


        # ----------------------------------------------------
        # 게시판 URL
        # ----------------------------------------------------

        board_url = ""

        if columns["board_url"]:

            board_url = clean(
                row.get(
                    columns["board_url"],
                    "",
                )
            )


        # ----------------------------------------------------
        # 게시물 URL
        # ----------------------------------------------------

        post_url = ""

        if columns["post_url"]:

            post_url = clean(
                row.get(
                    columns["post_url"],
                    "",
                )
            )


        # ----------------------------------------------------
        # 발견 여부
        # ----------------------------------------------------

        already_found = (

            board_count > 0

            or post_count > 0

            or bool(board_url)

            or bool(post_url)

        )


        if not already_found:

            target_rows.append(
                row.to_dict()
            )


    target_df = pd.DataFrame(
        target_rows
    )


    print()
    print(
        f"V6.3 정밀탐색 대상 : "
        f"{len(target_df)}개"
    )


    print(
        f"기존 후보 발견/제외 : "
        f"{len(df) - len(target_df)}개"
    )


    # ========================================================
    # 대상 없음
    # ========================================================

    if target_df.empty:

        print()
        print(
            "정밀탐색 대상 기관이 없습니다."
        )

        return


    # ========================================================
    # 기관명 목록
    # ========================================================

    print()
    print("=" * 80)
    print("정밀탐색 대상")
    print("=" * 80)


    for i, (_, row) in enumerate(
        target_df.iterrows(),
        1,
    ):

        name = clean(
            row.get(
                columns["institution"],
                "",
            )
        )

        print(
            f"{i:3d}. {name}"
        )


    # ========================================================
    # 비동기 실행
    # ========================================================

    semaphore = asyncio.Semaphore(
        CONCURRENCY
    )


    timeout = aiohttp.ClientTimeout(
        total=REQUEST_TIMEOUT
    )


    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        limit_per_host=2,
        ssl=False,
        ttl_dns_cache=300,
    )


    results = []


    async with aiohttp.ClientSession(
        timeout=timeout,
        connector=connector,
    ) as session:


        tasks = [

            analyze_one(
                session,
                semaphore,
                row,
                columns,
            )

            for row
            in target_df.to_dict(
                "records"
            )

        ]


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
    # 결과
    # ========================================================

    result_df = pd.DataFrame(
        results
    )


    original_columns = list(
        df.columns
    )


    new_columns = [

        c

        for c in result_df.columns

        if c not in original_columns

    ]


    result_df = result_df[
        original_columns
        + new_columns
    ]


    # ========================================================
    # 저장
    # ========================================================

    result_df.to_excel(
        OUTPUT,
        index=False,
    )


    # ========================================================
    # 최종 통계
    # ========================================================

    print()
    print("=" * 80)
    print("V6.3 정밀탐색 완료")
    print("=" * 80)


    print(
        f"V6.2 전체 : {len(df)}개"
    )


    print(
        f"V6.3 탐색 : {len(result_df)}개"
    )


    print(
        f"제외       : "
        f"{len(df) - len(result_df)}개"
    )


    print(
        f"결과파일   : {OUTPUT}"
    )


    print()


    # ========================================================
    # 신뢰도
    # ========================================================

    print(
        "===== 신뢰도 ====="
    )


    if not result_df.empty:

        print(
            result_df[
                "V6.3신뢰도"
            ]
            .value_counts(
                dropna=False
            )
            .to_string()
        )


    print()


    # ========================================================
    # 게시판 후보
    # ========================================================

    found = result_df[
        result_df[
            "V6.3게시판후보"
        ]
        .apply(to_int)
        > 0
    ]


    print(
        "===== V6.3 게시판 후보 발견 ====="
    )


    print(
        f"{len(found)}개 기관"
    )


    for _, row in found.iterrows():

        print(
            f"- "
            f"{clean(row.get(columns['institution'], ''))}"
            f" | "
            f"{clean(row.get('V6.3게시판URL', ''))}"
            f" | "
            f"{clean(row.get('V6.3신뢰도', ''))}"
        )


    print()
    print("=" * 80)
    print("V6.3 종료")
    print("=" * 80)


# ============================================================
# 실행
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
