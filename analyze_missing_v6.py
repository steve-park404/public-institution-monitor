import asyncio
import aiohttp
import pandas as pd
import re
import ssl
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse
from collections import deque
from pathlib import Path


# =========================================================
# 기본 설정
# =========================================================

INPUT_FILE = "boards_missing.xlsx"
OUTPUT_FILE = "boards_v6.xlsx"

CONCURRENCY = 8
TIMEOUT = 25
MAX_PAGES = 35
MAX_LINKS_PER_PAGE = 150

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/140.0 Safari/537.36",

    "Mozilla/5.0 (X11; Linux x86_64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/140.0 Safari/537.36",

    "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) "
    "Version/18.0 Mobile/15E148 Safari/604.1",
]


# =========================================================
# 키워드
# =========================================================

BOARD_WORDS = [
    "공지사항",
    "공지",
    "알림",
    "소식",
    "새소식",
    "게시판",
    "공지·공고",
    "공지/공고",
    "공고",
    "채용",
    "입찰",
    "보도자료",
    "자료실",
    "뉴스",
    "정보마당",
    "고시",
    "공시",
]

NOTICE_WORDS = [
    "공지사항",
    "공지",
    "알림",
    "새소식",
    "소식",
    "공지·공고",
    "공지/공고",
]

BAD_IFRAME_WORDS = [
    "google.com/maps",
    "youtube.com/embed",
    "googletagmanager.com",
    "google-analytics.com",
    "doubleclick.net",
    "facebook.com/plugins",
    "instagram.com",
]

API_WORDS = [
    "fetch(",
    "axios",
    "$.ajax",
    "$.get",
    "$.post",
    "XMLHttpRequest",
    "ajax/",
    "api/",
    "/api",
]

BOARD_URL_WORDS = [
    "board",
    "bbs",
    "notice",
    "news",
    "announce",
    "bulletin",
    "community",
    "boardList",
    "boardView",
    "board.do",
    "bbs.do",
    "notice.do",
    "list.do",
    "view.do",
]

K2WEB_WORDS = [
    "k2web",
    "fnctId",
    "nttId",
    "siteId",
    "bbsId",
]

POST_PARAM_WORDS = [
    "nttId",
    "ntt_id",
    "articleId",
    "article_id",
    "seq",
    "idx",
    "no",
    "bbsSeq",
    "boardSeq",
]

CONTENT_TYPES = [
    "text/html",
    "application/xhtml+xml",
]


# =========================================================
# SSL
# =========================================================

SSL_CONTEXT = ssl.create_default_context()
SSL_CONTEXT.check_hostname = False
SSL_CONTEXT.verify_mode = ssl.CERT_NONE


# =========================================================
# URL
# =========================================================

def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def same_domain(base, target):
    try:
        b = urlparse(base).netloc.lower().replace("www.", "")
        t = urlparse(target).netloc.lower().replace("www.", "")
        return b == t
    except:
        return False


def clean_url(url):
    try:
        p = urlparse(url)

        query = parse_qs(
            p.query,
            keep_blank_values=True
        )

        # 추적용 파라미터 제거
        remove = {
            "utm_source",
            "utm_medium",
            "utm_campaign",
            "utm_term",
            "utm_content",
        }

        query = {
            k: v
            for k, v in query.items()
            if k not in remove
        }

        new_query = urlencode(query, doseq=True)

        return urlunparse((
            p.scheme,
            p.netloc,
            p.path,
            p.params,
            new_query,
            ""
        ))

    except:
        return url


# =========================================================
# HTML 가져오기
# =========================================================

async def fetch(session, url, retries=2):

    url = normalize_url(url)

    last_error = ""

    for attempt in range(retries + 1):

        headers = {
            "User-Agent": USER_AGENTS[
                attempt % len(USER_AGENTS)
            ],
            "Accept":
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language":
                "ko-KR,ko;q=0.9,en-US;q=0.7",
            "Connection": "keep-alive",
        }

        try:

            timeout = aiohttp.ClientTimeout(
                total=TIMEOUT + attempt * 10
            )

            async with session.get(
                url,
                headers=headers,
                timeout=timeout,
                ssl=SSL_CONTEXT,
                allow_redirects=True,
            ) as r:

                content_type = (
                    r.headers.get(
                        "Content-Type",
                        ""
                    ).lower()
                )

                text = await r.text(
                    errors="ignore"
                )

                return {
                    "ok": r.status < 400,
                    "status": r.status,
                    "url": str(r.url),
                    "content_type": content_type,
                    "text": text,
                    "error": "",
                }

        except Exception as e:

            last_error = type(e).__name__

            await asyncio.sleep(
                1 + attempt
            )

    return {
        "ok": False,
        "status": 0,
        "url": url,
        "content_type": "",
        "text": "",
        "error": last_error,
    }


# =========================================================
# 링크 추출
# =========================================================

def extract_links(html, base_url):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    links = []

    for a in soup.find_all("a", href=True):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith(
            ("javascript:", "mailto:", "tel:", "#")
        ):
            continue

        try:
            absolute = urljoin(
                base_url,
                href
            )

            absolute = clean_url(
                absolute
            )

            if absolute.startswith(
                ("http://", "https://")
            ):
                links.append({
                    "url": absolute,
                    "text": a.get_text(
                        " ",
                        strip=True
                    )
                })

        except:
            pass

    return links[:MAX_LINKS_PER_PAGE]


# =========================================================
# iframe 추출
# =========================================================

def extract_iframes(html, base_url):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    result = []

    for iframe in soup.find_all(
        "iframe"
    ):

        src = (
            iframe.get("src")
            or iframe.get("data-src")
            or iframe.get("data-url")
            or ""
        ).strip()

        if not src:
            continue

        absolute = urljoin(
            base_url,
            src
        )

        result.append(absolute)

    return result


# =========================================================
# JS/API 분석
# =========================================================

def extract_scripts(html, base_url):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    scripts = []

    for s in soup.find_all("script"):

        src = s.get("src")

        if src:
            scripts.append(
                urljoin(base_url, src)
            )

        else:
            text = s.get_text(
                " ",
                strip=True
            )

            if text:
                scripts.append(
                    text[:20000]
                )

    return scripts


def analyze_js(html):

    low = html.lower()

    score = 0
    findings = []

    for word in API_WORDS:

        if word.lower() in low:

            score += 1
            findings.append(word)

    endpoints = []

    patterns = [
        r'["\']([^"\']{0,200}(?:api|ajax|board|bbs|notice)[^"\']{0,200})["\']',
        r'url\s*:\s*["\']([^"\']+)["\']',
        r'fetch\s*\(\s*["\']([^"\']+)["\']',
        r'axios\.(?:get|post)\s*\(\s*["\']([^"\']+)["\']',
    ]

    for pattern in patterns:

        try:

            matches = re.findall(
                pattern,
                html,
                re.I
            )

            for x in matches:

                if isinstance(x, tuple):
                    x = x[0]

                x = str(x).strip()

                if (
                    x
                    and len(x) < 500
                    and x not in endpoints
                ):
                    endpoints.append(x)

        except:
            pass

    return score, findings, endpoints[:30]


# =========================================================
# 게시판 후보 판정
# =========================================================

def board_score(url, text, html=""):

    target = (
        str(url)
        + " "
        + str(text)
    ).lower()

    score = 0
    reasons = []

    for word in BOARD_URL_WORDS:

        if word.lower() in target:

            score += 2
            reasons.append(
                f"URL:{word}"
            )

    for word in BOARD_WORDS:

        if word.lower() in text.lower():

            score += 2
            reasons.append(
                f"TEXT:{word}"
            )

    # 목록 구조
    if html:

        soup = BeautifulSoup(
            html,
            "html.parser"
        )

        body_text = soup.get_text(
            " ",
            strip=True
        )

        # 게시물 날짜가 여러 개 존재
        date_count = len(
            re.findall(
                r"\b20\d{2}[./-]\d{1,2}[./-]\d{1,2}\b",
                body_text
            )
        )

        if date_count >= 2:

            score += 4
            reasons.append(
                "날짜복수"
            )

        # 페이지네이션
        page_words = [
            "다음",
            "이전",
            "1",
            "2",
            "3",
        ]

        page_count = sum(
            1
            for x in page_words
            if x in body_text
        )

        if page_count >= 3:

            score += 2
            reasons.append(
                "페이지네이션"
            )

    return score, reasons


# =========================================================
# 게시물 상세 URL 판정
# =========================================================

def looks_like_post(url, text=""):

    low = (
        str(url)
        + " "
        + str(text)
    ).lower()

    # 상세 페이지에서 흔히 사용되는 파라미터
    parsed = urlparse(url)

    query = parsed.query.lower()

    for p in POST_PARAM_WORDS:

        if (
            p.lower() + "="
        ) in query:

            return True

    patterns = [
        r"/view/",
        r"/view\?",
        r"/read/",
        r"/read\?",
        r"/article/",
        r"/article\?",
        r"/detail/",
        r"/detail\?",
        r"/board/view",
        r"/bbs/view",
        r"/notice/view",
        r"/board/read",
    ]

    for pattern in patterns:

        if re.search(
            pattern,
            low
        ):
            return True

    return False


# =========================================================
# 부모 게시판 URL 추정
# =========================================================

def parent_candidates(url):

    result = []

    try:

        p = urlparse(url)

        path = p.path.rstrip("/")

        parts = path.split("/")

        # 마지막 path 제거
        if len(parts) > 1:

            parent = "/".join(
                parts[:-1]
            )

            result.append(
                urlunparse((
                    p.scheme,
                    p.netloc,
                    parent + "/",
                    "",
                    "",
                    ""
                ))
            )

        # query parameter 제거
        if p.query:

            result.append(
                urlunparse((
                    p.scheme,
                    p.netloc,
                    p.path,
                    "",
                    "",
                    ""
                ))
            )

        # view -> list
        for token in [
            "/view",
            "/read",
            "/detail",
            "/article",
        ]:

            if token in path.lower():

                candidate = re.sub(
                    token,
                    "/list",
                    path,
                    flags=re.I
                )

                result.append(
                    urlunparse((
                        p.scheme,
                        p.netloc,
                        candidate,
                        "",
                        "",
                        ""
                    ))
                )

    except:
        pass

    return list(
        dict.fromkeys(result)
    )


# =========================================================
# iframe 유효성
# =========================================================

def valid_iframe(url):

    low = url.lower()

    for bad in BAD_IFRAME_WORDS:

        if bad in low:
            return False

    return True


# =========================================================
# 기관 1개 분석
# =========================================================

async def analyze_institution(
    session,
    semaphore,
    row
):

    async with semaphore:

        name = str(
            row.get(
                "기관명",
                ""
            )
        )

        homepage = normalize_url(
            row.get(
                "URL",
                ""
            )
        )

        result = {
            "기관명": name,
            "홈페이지": homepage,
            "접속상태": "",
            "CMS": "",
            "iframe": "",
            "JS": "",
            "후보수": 0,
            "게시물후보수": 0,
            "게시판URL": "",
            "게시물URL": "",
            "iframeURL": "",
            "API후보": "",
            "탐색방법": "",
            "신뢰도": "",
            "오류": "",
        }

        first = await fetch(
            session,
            homepage,
            retries=2
        )

        if not first["ok"]:

            result["접속상태"] = (
                f"실패:{first['error'] or first['status']}"
            )

            return result

        result["접속상태"] = (
            f"정상:{first['status']}"
        )

        visited = set()
        queue = deque()

        queue.append(
            homepage
        )

        candidate_boards = []
        candidate_posts = []
        candidate_iframes = []
        candidate_api = []

        js_score = 0
        cms_findings = []

        pages = 0

        while queue and pages < MAX_PAGES:

            current = queue.popleft()

            current = clean_url(
                current
            )

            if current in visited:
                continue

            if not same_domain(
                homepage,
                current
            ):
                continue

            visited.add(current)

            data = await fetch(
                session,
                current,
                retries=1
            )

            if not data["ok"]:
                continue

            content_type = data[
                "content_type"
            ]

            if not any(
                x in content_type
                for x in CONTENT_TYPES
            ):
                continue

            html = data["text"]

            pages += 1

            # ---------------------------------------------
            # CMS 탐지
            # ---------------------------------------------

            low = html.lower()

            if "k2web" in low:

                cms_findings.append(
                    "K2Web"
                )

            if (
                "joomla" in low
                or "/components/com_" in low
            ):

                cms_findings.append(
                    "Joomla"
                )

            if "wordpress" in low:

                cms_findings.append(
                    "WordPress"
                )

            if (
                "drupal" in low
                or "drupalsettings" in low
            ):

                cms_findings.append(
                    "Drupal"
                )

            if (
                "gnu board" in low
                or "gnuboard" in low
                or "g5_" in low
            ):

                cms_findings.append(
                    "그누보드"
                )

            # ---------------------------------------------
            # iframe
            # ---------------------------------------------

            iframe_urls = extract_iframes(
                html,
                current
            )

            for iframe_url in iframe_urls:

                if valid_iframe(
                    iframe_url
                ):

                    if iframe_url not in candidate_iframes:

                        candidate_iframes.append(
                            iframe_url
                        )

                        if len(candidate_iframes) < 10:

                            queue.append(
                                iframe_url
                            )

            # ---------------------------------------------
            # JS/API
            # ---------------------------------------------

            score, findings, endpoints = (
                analyze_js(html)
            )

            js_score += score

            for endpoint in endpoints:

                if endpoint not in candidate_api:

                    candidate_api.append(
                        endpoint
                    )

                    # 상대경로 API
                    if endpoint.startswith("/"):

                        api_url = urljoin(
                            current,
                            endpoint
                        )

                        if same_domain(
                            homepage,
                            api_url
                        ):
                            queue.append(
                                api_url
                            )

            # ---------------------------------------------
            # 링크
            # ---------------------------------------------

            links = extract_links(
                html,
                current
            )

            for item in links:

                link = item["url"]
                text = item["text"]

                score, reasons = board_score(
                    link,
                    text,
                    html
                )

                if score >= 5:

                    candidate = {
                        "url": link,
                        "text": text,
                        "score": score,
                        "reasons": ",".join(
                            reasons[:10]
                        ),
                    }

                    if not any(
                        x["url"] == link
                        for x in candidate_boards
                    ):

                        candidate_boards.append(
                            candidate
                        )

                if looks_like_post(
                    link,
                    text
                ):

                    if link not in candidate_posts:

                        candidate_posts.append(
                            link
                        )

                        # 부모 게시판 추적
                        for parent in parent_candidates(
                            link
                        ):

                            if same_domain(
                                homepage,
                                parent
                            ):

                                queue.append(
                                    parent
                                )

                # -----------------------------------------
                # BFS
                # -----------------------------------------

                link_text = (
                    text
                    + " "
                    + link
                ).lower()

                priority = 0

                for word in NOTICE_WORDS:

                    if word.lower() in link_text:
                        priority += 5

                for word in BOARD_URL_WORDS:

                    if word.lower() in link.lower():
                        priority += 2

                if (
                    "k2web" in link.lower()
                    or "fnctid" in link.lower()
                    or "nttid" in link.lower()
                ):

                    priority += 5

                if priority >= 3:

                    if (
                        link not in visited
                        and same_domain(
                            homepage,
                            link
                        )
                    ):

                        queue.append(
                            link
                        )

            # 일반 메뉴도 제한적으로 탐색
            if pages <= 8:

                for item in links:

                    link = item["url"]
                    text = item["text"]

                    if not same_domain(
                        homepage,
                        link
                    ):
                        continue

                    if len(
                        queue
                    ) >= MAX_PAGES * 2:
                        break

                    # 메뉴/공지 관련 링크
                    target = (
                        text
                        + " "
                        + link
                    ).lower()

                    if any(
                        x.lower() in target
                        for x in BOARD_WORDS
                    ):

                        queue.append(
                            link
                        )

        # =================================================
        # 후보 정리
        # =================================================

        candidate_boards = sorted(
            candidate_boards,
            key=lambda x: x["score"],
            reverse=True
        )

        candidate_boards = (
            candidate_boards[:10]
        )

        candidate_posts = list(
            dict.fromkeys(
                candidate_posts
            )
        )[:10]

        # =================================================
        # 실제 게시판 검증
        # =================================================

        verified_boards = []

        for candidate in candidate_boards[:8]:

            check = await fetch(
                session,
                candidate["url"],
                retries=1
            )

            if not check["ok"]:
                continue

            score, reasons = board_score(
                candidate["url"],
                candidate["text"],
                check["text"]
            )

            if score >= 6:

                verified_boards.append({
                    "url": candidate["url"],
                    "score": score,
                    "reasons": reasons,
                })

        # =================================================
        # 부모 게시판 검증
        # =================================================

        for post in candidate_posts[:8]:

            for parent in parent_candidates(
                post
            )[:5]:

                check = await fetch(
                    session,
                    parent,
                    retries=1
                )

                if not check["ok"]:
                    continue

                score, reasons = board_score(
                    parent,
                    "",
                    check["text"]
                )

                if score >= 6:

                    verified_boards.append({
                        "url": parent,
                        "score": score,
                        "reasons": [
                            "상세URL역추적"
                        ] + reasons,
                    })

        # =================================================
        # 중복 제거
        # =================================================

        unique = {}

        for item in verified_boards:

            url = clean_url(
                item["url"]
            )

            if url not in unique:

                unique[url] = item

            elif item["score"] > unique[url]["score"]:

                unique[url] = item

        verified_boards = sorted(
            unique.values(),
            key=lambda x: x["score"],
            reverse=True
        )

        # =================================================
        # 최종 결과
        # =================================================

        result["CMS"] = ", ".join(
            sorted(
                set(cms_findings)
            )
        ) or "일반 HTML"

        if candidate_iframes:

            result["iframe"] = (
                f"{len(candidate_iframes)}개"
            )

            result["iframeURL"] = (
                " | ".join(
                    candidate_iframes[:5]
                )
            )

        else:

            result["iframe"] = "없음"

        if js_score >= 5:

            result["JS"] = "높음"

        elif js_score >= 2:

            result["JS"] = "중간"

        else:

            result["JS"] = "낮음"

        result["후보수"] = len(
            candidate_boards
        )

        result["게시물후보수"] = len(
            candidate_posts
        )

        if verified_boards:

            result["게시판URL"] = (
                " | ".join(
                    x["url"]
                    for x in verified_boards[:5]
                )
            )

            result["신뢰도"] = (
                str(
                    verified_boards[0]["score"]
                )
            )

            result["탐색방법"] = (
                "실제페이지검증"
            )

        elif candidate_posts:

            result["게시물URL"] = (
                " | ".join(
                    candidate_posts[:5]
                )
            )

            result["탐색방법"] = (
                "게시물상세후보"
            )

        elif candidate_iframes:

            result["탐색방법"] = (
                "iframe"
            )

        elif candidate_api:

            result["탐색방법"] = (
                "JS/API"
            )

        else:

            result["탐색방법"] = (
                "메뉴/BFS"
            )

        result["API후보"] = (
            " | ".join(
                candidate_api[:10]
            )
        )

        return result


# =========================================================
# 메인
# =========================================================

async def main():

    print("=" * 80)
    print("6차 미발견 기관 게시판 심층 탐색")
    print("=" * 80)

    if not Path(
        INPUT_FILE
    ).exists():

        print(
            f"입력 파일 없음: {INPUT_FILE}"
        )

        return

    df = pd.read_excel(
        INPUT_FILE
    )

    print(
        f"대상 기관: {len(df)}개"
    )

    # 필요한 컬럼 확인
    required = [
        "기관명",
        "URL",
    ]

    for col in required:

        if col not in df.columns:

            print(
                f"필수 컬럼 없음: {col}"
            )

            return

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=SSL_CONTEXT,
        limit_per_host=3
    )

    semaphore = asyncio.Semaphore(
        CONCURRENCY
    )

    async with aiohttp.ClientSession(
        connector=connector
    ) as session:

        tasks = []

        for _, row in df.iterrows():

            tasks.append(
                analyze_institution(
                    session,
                    semaphore,
                    row
                )
            )

        results = []

        for i, task in enumerate(
            asyncio.as_completed(tasks),
            start=1
        ):

            try:

                result = await task

                results.append(
                    result
                )

                print(
                    f"[{i}/{len(tasks)}] "
                    f"{result['기관명']} | "
                    f"{result['접속상태']} | "
                    f"게시판={result['게시판URL'][:80]}"
                )

            except Exception as e:

                print(
                    f"[{i}/{len(tasks)}] 오류: {e}"
                )

    result_df = pd.DataFrame(
        results
    )

    # 기관명 순서 정렬
    result_df = result_df.sort_values(
        "기관명"
    )

    result_df.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print()
    print("=" * 80)
    print("6차 탐색 완료")
    print("=" * 80)

    print(
        f"전체 기관: {len(result_df)}개"
    )

    print(
        "실제 게시판 발견:",
        (
            result_df["게시판URL"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .sum()
        ),
        "개"
    )

    print(
        "게시물 후보 발견:",
        (
            result_df["게시물URL"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .sum()
        ),
        "개"
    )

    print(
        "iframe 발견:",
        (
            result_df["iframe"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .sum()
        ),
        "개"
    )

    print()
    print(
        f"결과 파일: {OUTPUT_FILE}"
    )


if __name__ == "__main__":

    asyncio.run(
        main()
    )
