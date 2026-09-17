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

# 동시 접속 수
CONCURRENCY = 8

# 개별 HTTP 요청 제한
REQUEST_TIMEOUT = 15

# 기관 하나당 최대 탐색 페이지
MAX_PAGES = 25

# 링크 탐색 깊이
MAX_DEPTH = 2

# 한 페이지에서 다음 탐색 대상으로 사용할 최대 링크
MAX_LINKS_PER_PAGE = 60

# 기관 하나의 전체 분석 제한시간
INSTITUTION_TIMEOUT = 70


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,*/*;q=0.8"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.7,en;q=0.5",
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
    "사업공고",
    "입찰공고",
    "채용공고",
    "언론보도",
    "보도",
    "알림마당",
    "공지/공고",
]


# ============================================================
# URL 패턴
# ============================================================

BOARD_URL_WORDS = [
    "board",
    "bbs",
    "notice",
    "news",
    "announce",
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
]


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
# API / JS 패턴
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
# CMS 패턴
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
# 공통 URL
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
# URL 정리
# ============================================================

def normalize_url(url):

    if not url:
        return ""

    url = str(url).strip()

    if not url:
        return ""

    return url


# ============================================================
# Host 비교
# ============================================================

def same_host(a, b):

    try:

        host_a = urlparse(a).netloc.lower().split(":")[0]
        host_b = urlparse(b).netloc.lower().split(":")[0]

        return host_a == host_b

    except:

        return False


# ============================================================
# 링크 점수
# ============================================================

def score_link(text, href):

    text = str(text or "")
    href = str(href or "")

    value = (
        text
        + " "
        + href
    ).lower()

    score = 0


    # 한글 게시판 키워드

    for word in BOARD_WORDS:

        if word.lower() in value:

            score += 4


    # URL 패턴

    for word in BOARD_URL_WORDS:

        if word in value:

            score += 2


    # 게시물 패턴

    for word in POST_URL_WORDS:

        if word in value:

            score += 1


    # 불필요한 페이지 감점

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
# V6.2 결과를 이용한 공략방법
# ============================================================

def classify_method(row):

    status = str(
        row.get("접속상태", "")
    )

    cms = str(
        row.get("CMS", "")
    )

    iframe = str(
        row.get("iframe", "")
    )

    js = str(
        row.get("JS", "")
    )

    fail = str(
        row.get("실패단계", "")
    )


    old_board = to_int(
        row.get("게시판후보수", 0)
    )

    old_post = to_int(
        row.get("게시물후보수", 0)
    )


    # 이미 후보가 있으면 이번 V6.3 대상에서
    # 원칙적으로 제외

    if old_board > 0 or old_post > 0:

        return "제외_기존후보"


    # 접속 실패

    if (
        "실패" in status
        or "Timeout" in status
        or "실패" in fail
    ):

        return "A_접속실패_재시도"


    # K2Web

    if "k2web" in cms.lower():

        return "B_K2Web"


    # iframe

    if any(
        x in iframe.lower()
        for x in [
            "있음",
            "true",
            "yes",
        ]
    ):

        return "C_iframe"


    # JS

    if any(
        x in js.lower()
        for x in [
            "높음",
            "high",
            "동적",
            "yes",
        ]
    ):

        return "D_JS_API"


    return "E_메뉴_사이트맵_URL"


# ============================================================
# HTTP 요청
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

            body = await response.text(
                errors="ignore"
            )

            return (
                response.status,
                str(response.url),
                body,
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
# URL에서 게시판 패턴 검사
# ============================================================

def is_board_url(url):

    value = str(url).lower()

    return any(
        word in value
        for word in BOARD_URL_WORDS
    )


# ============================================================
# URL에서 게시물 패턴 검사
# ============================================================

def is_post_url(url):

    value = str(url).lower()

    return any(
        word in value
        for word in POST_URL_WORDS
    )


# ============================================================
# 기관 분석
# ============================================================

async def analyze_one(
    session,
    semaphore,
    record,
):

    async with semaphore:

        return await asyncio.wait_for(
            analyze_one_inner(
                session,
                record,
            ),
            timeout=INSTITUTION_TIMEOUT,
        )


# ============================================================
# 실제 기관 분석
# ============================================================

async def analyze_one_inner(
    session,
    record,
):

    homepage = normalize_url(
        record.get(
            "홈페이지",
            "",
        )
    )


    result = dict(record)


    # 결과 필드

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

        "V6.3탐색방법": classify_method(
            record
        ),

        "V6.3신뢰도": "미발견",

        "V6.3오류": "",

    })


    # ========================================================
    # 홈페이지 없음
    # ========================================================

    if (
        not homepage
        or homepage.lower() == "nan"
    ):

        result[
            "V6.3상태"
        ] = "홈페이지없음"

        return result


    # ========================================================
    # 탐색 큐
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
    # 홈페이지의 공통 URL도 큐에 추가
    # ========================================================

    homepage_base = homepage.rstrip("/")


    for path in COMMON_PATHS:

        target = urljoin(
            homepage_base + "/",
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


        # 오류

        if error:

            if not result["V6.3오류"]:

                result[
                    "V6.3오류"
                ] = error

            continue


        # 첫 번째 정상 접속

        if (
            len(seen) == 1
            and status is not None
        ):

            result[
                "V6.3상태"
            ] = f"정상:{status}"


        if status is None:

            continue


        if status >= 400:

            continue


        # ====================================================
        # HTML 분석
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


            # iframe은 외부 도메인도 추적

            if (
                target not in seen
                and len(seen)
                + len(queue)
                < MAX_PAGES + 10
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

            src = script.get("src")

            inline_code = (
                script.string
                or ""
            )


            blob = (
                str(src or "")
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


                # API 주소 추출

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


            # javascript 제거

            if href.lower().startswith(
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


            # =================================================
            # 게시판 후보
            # =================================================

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


            # =================================================
            # 게시물 후보
            # =================================================

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


            # =================================================
            # 다음 페이지
            # =================================================

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
        # 점수 높은 링크부터 탐색
        # ====================================================

        next_links.sort(
            reverse=True
        )


        added = 0


        for (
            score,
            target,
        ) in next_links:

            if target in seen:

                continue


            if target in [
                x[0]
                for x in queue
            ]:

                continue


            queue.append(
                (
                    target,
                    depth + 1,
                )
            )


            added += 1


            if added >= MAX_LINKS_PER_PAGE:

                break


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
    # 입력파일 확인
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
    # 미발견 기관만 선택
    #
    # 게시판 후보 = 0
    # AND
    # 게시물 후보 = 0
    # ========================================================

    target_df = df[
        (
            df["게시판후보수"]
            .apply(to_int)
            == 0
        )
        &
        (
            df["게시물후보수"]
            .apply(to_int)
            == 0
        )
    ].copy()


    print(
        f"V6.3 정밀탐색 대상 : "
        f"{len(target_df)}개"
    )


    print(
        f"제외 기관 : "
        f"{len(df) - len(target_df)}개"
    )


    # ========================================================
    # 대상이 없으면 종료
    # ========================================================

    if target_df.empty:

        print(
            "정밀탐색 대상이 없습니다."
        )

        return


    records = target_df.to_dict(
        "records"
    )


    # ========================================================
    # HTTP 세션
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
                record,
            )

            for record
            in records

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


            except asyncio.TimeoutError:

                print(
                    "[TIMEOUT] 기관 분석시간 초과"
                )


            except Exception as e:

                print(
                    "[ERROR]",
                    repr(e)
                )


            completed += 1


            if (
                completed % 10 == 0
                or completed == len(records)
            ):

                print(
                    f"[V6.3] "
                    f"{completed}/"
                    f"{len(records)} 완료"
                )


    # ========================================================
    # 결과 DataFrame
    # ========================================================

    result_df = pd.DataFrame(
        results
    )


    # ========================================================
    # 원래 컬럼 순서 + V6.3 컬럼
    # ========================================================

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
    # 결과 요약
    # ========================================================

    print()

    print("=" * 80)
    print("V6.3 정밀탐색 완료")
    print("=" * 80)


    print(
        f"전체 V6.2 기관 : {len(df)}개"
    )


    print(
        f"V6.3 분석 기관 : "
        f"{len(result_df)}개"
    )


    print(
        f"결과 파일 : {OUTPUT}"
    )


    print()


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


    print(
        "===== 공략방법 ====="
    )


    if not result_df.empty:

        print(
            result_df[
                "V6.3탐색방법"
            ]
            .value_counts(
                dropna=False
            )
            .to_string()
        )


    print()


    print(
        "===== 게시판 후보 발견 기관 ====="
    )


    found = result_df[
        result_df[
            "V6.3게시판후보"
        ]
        .apply(to_int)
        > 0
    ]


    print(
        f"{len(found)}개 기관"
    )


    for _, row in found.iterrows():

        print(
            f"- "
            f"{row.get('기관명', '')} | "
            f"{row.get('V6.3게시판URL', '')} | "
            f"{row.get('V6.3신뢰도', '')}"
        )


    print()

    print("=" * 80)
    print("분석 종료")
    print("=" * 80)


# ============================================================
# 실행
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
