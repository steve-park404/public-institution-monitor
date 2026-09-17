import asyncio
import aiohttp
import pandas as pd
import re
import ssl

from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse
from collections import deque


# ============================================================
# V6.2 안정형
# 미발견 기관만 대상으로 게시판 구조 탐색
# ============================================================

INPUT_FILE = "boards_missing.xlsx"
OUTPUT_FILE = "boards_v6.xlsx"

# 동시에 분석할 기관 수
CONCURRENCY = 6

# HTTP 1회 요청 제한
REQUEST_TIMEOUT = 15

# 기관 1개 전체 분석 제한
INSTITUTION_TIMEOUT = 90

# 기관당 최대 페이지
MAX_PAGES = 15

# 페이지당 최대 링크
MAX_LINKS = 80

# 실제 게시판 검증
MAX_VERIFY = 5


# ============================================================
# SSL
# ============================================================

SSL_CONTEXT = ssl.create_default_context()
SSL_CONTEXT.check_hostname = False
SSL_CONTEXT.verify_mode = ssl.CERT_NONE


# ============================================================
# User-Agent
# ============================================================

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/140.0.0.0 Safari/537.36"
)


# ============================================================
# 키워드
# ============================================================

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

BOARD_URL_WORDS = [
    "board",
    "bbs",
    "notice",
    "news",
    "announce",
    "bulletin",
    "community",
    "boardlist",
    "boardview",
    "board.do",
    "bbs.do",
    "notice.do",
    "list.do",
    "view.do",
]

POST_PARAMS = [
    "nttid",
    "ntt_id",
    "articleid",
    "article_id",
    "seq",
    "idx",
    "no",
    "bbsseq",
    "boardseq",
]

CMS_WORDS = [
    ("K2Web", "k2web"),
    ("WordPress", "wordpress"),
    ("Drupal", "drupal"),
    ("Joomla", "joomla"),
    ("그누보드", "gnuboard"),
]


# ============================================================
# URL
# ============================================================

def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def clean_url(url):
    try:
        p = urlparse(url)

        query = parse_qs(
            p.query,
            keep_blank_values=True
        )

        remove_keys = {
            "utm_source",
            "utm_medium",
            "utm_campaign",
            "utm_term",
            "utm_content",
        }

        query = {
            k: v
            for k, v in query.items()
            if k.lower() not in remove_keys
        }

        new_query = urlencode(
            query,
            doseq=True
        )

        return urlunparse(
            (
                p.scheme,
                p.netloc,
                p.path,
                p.params,
                new_query,
                "",
            )
        )

    except Exception:
        return url


def same_domain(base, target):
    try:
        a = urlparse(base).netloc.lower()
        b = urlparse(target).netloc.lower()

        a = a.replace("www.", "")
        b = b.replace("www.", "")

        return a == b

    except Exception:
        return False


# ============================================================
# HTTP
# ============================================================

async def fetch(session, url):
    url = normalize_url(url)

    if not url:
        return None

    headers = {
        "User-Agent": USER_AGENT,
        "Accept": (
            "text/html,"
            "application/xhtml+xml,"
            "*/*;q=0.8"
        ),
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Connection": "close",
    }

    timeout = aiohttp.ClientTimeout(
        total=REQUEST_TIMEOUT,
        connect=8,
        sock_read=REQUEST_TIMEOUT,
    )

    try:
        async with session.get(
            url,
            headers=headers,
            timeout=timeout,
            ssl=SSL_CONTEXT,
            allow_redirects=True,
            max_redirects=5,
        ) as response:

            content_type = (
                response.headers
                .get("Content-Type", "")
                .lower()
            )

            if response.status >= 400:
                return {
                    "ok": False,
                    "status": response.status,
                    "url": str(response.url),
                    "content_type": content_type,
                    "text": "",
                    "error": "HTTP " + str(response.status),
                }

            text = await response.text(
                encoding=None,
                errors="ignore",
            )

            return {
                "ok": True,
                "status": response.status,
                "url": str(response.url),
                "content_type": content_type,
                "text": text,
                "error": "",
            }

    except asyncio.TimeoutError:
        return {
            "ok": False,
            "status": 0,
            "url": url,
            "content_type": "",
            "text": "",
            "error": "Timeout",
        }

    except aiohttp.ClientError as e:
        return {
            "ok": False,
            "status": 0,
            "url": url,
            "content_type": "",
            "text": "",
            "error": type(e).__name__,
        }

    except Exception as e:
        return {
            "ok": False,
            "status": 0,
            "url": url,
            "content_type": "",
            "text": "",
            "error": type(e).__name__,
        }


# ============================================================
# HTML
# ============================================================

def soup_from_html(html):
    if not html:
        return None

    try:
        return BeautifulSoup(
            html,
            "html.parser"
        )
    except Exception:
        return None


def extract_links(html, base_url):
    soup = soup_from_html(html)

    if soup is None:
        return []

    result = []
    seen = set()

    for tag in soup.find_all("a", href=True):

        href = tag.get("href", "").strip()

        if not href:
            continue

        if href.startswith(
            (
                "javascript:",
                "mailto:",
                "tel:",
                "#",
            )
        ):
            continue

        try:
            url = clean_url(
                urljoin(
                    base_url,
                    href,
                )
            )
        except Exception:
            continue

        if not url.startswith(
            ("http://", "https://")
        ):
            continue

        if url in seen:
            continue

        seen.add(url)

        text = tag.get_text(
            " ",
            strip=True,
        )

        result.append(
            {
                "url": url,
                "text": text,
            }
        )

        if len(result) >= MAX_LINKS:
            break

    return result


# ============================================================
# CMS
# ============================================================

def detect_cms(html):
    low = html.lower()
    result = []

    for name, word in CMS_WORDS:
        if word in low:
            result.append(name)

    return list(dict.fromkeys(result))


# ============================================================
# iframe
# ============================================================

def detect_iframe(html, base_url):
    soup = soup_from_html(html)

    if soup is None:
        return []

    result = []

    for tag in soup.find_all("iframe"):

        src = (
            tag.get("src")
            or tag.get("data-src")
            or tag.get("data-url")
            or ""
        ).strip()

        if not src:
            continue

        try:
            url = urljoin(
                base_url,
                src,
            )

            low = url.lower()

            if any(
                x in low
                for x in [
                    "youtube.com",
                    "google.com/maps",
                    "googletagmanager",
                    "google-analytics",
                    "facebook.com/plugins",
                ]
            ):
                continue

            result.append(url)

        except Exception:
            continue

    return result


# ============================================================
# JS / API
# ============================================================

def analyze_js(html):
    low = html.lower()

    findings = []

    words = [
        "fetch(",
        "axios",
        "$.ajax",
        "$.get(",
        "$.post(",
        "xmlhttprequest",
        "/api/",
        "api/",
        "ajax/",
    ]

    for word in words:
        if word.lower() in low:
            findings.append(word)

    return list(dict.fromkeys(findings))


# ============================================================
# 게시판 후보 점수
# ============================================================

def board_score(url, text, html=""):
    target = (
        str(url)
        + " "
        + str(text)
    )

    low = target.lower()

    score = 0
    reasons = []

    for word in BOARD_URL_WORDS:
        if word.lower() in low:
            score += 2
            reasons.append(
                "URL:" + word
            )

    text_low = str(text).lower()

    for word in BOARD_WORDS:
        if word.lower() in text_low:
            score += 3
            reasons.append(
                "TEXT:" + word
            )

    if html:

        soup = soup_from_html(html)

        if soup:

            body = soup.get_text(
                " ",
                strip=True,
            )

            dates = re.findall(
                r"20\d{2}[./-]\d{1,2}[./-]\d{1,2}",
                body,
            )

            if len(dates) >= 2:
                score += 4
                reasons.append("날짜복수")

            if any(
                word in body
                for word in [
                    "다음",
                    "이전",
                    "페이지",
                    "검색",
                ]
            ):
                score += 2
                reasons.append("게시판요소")

            if len(soup.find_all("tr")) >= 3:
                score += 2
                reasons.append("목록행")

    return score, list(
        dict.fromkeys(reasons)
    )


# ============================================================
# 게시물 판별
# ============================================================

def looks_like_post(url):
    low = url.lower()

    parsed = urlparse(url)
    query = parsed.query.lower()

    for param in POST_PARAMS:
        if param + "=" in query:
            return True

    patterns = [
        "/view/",
        "/view?",
        "/read/",
        "/read?",
        "/article/",
        "/article?",
        "/detail/",
        "/detail?",
        "/board/view",
        "/bbs/view",
        "/notice/view",
        "/board/read",
    ]

    for pattern in patterns:
        if pattern in low:
            return True

    return False


# ============================================================
# 게시물에서 상위 URL 추정
# ============================================================

def parent_candidates(url):
    result = []

    try:
        p = urlparse(url)

        path = p.path.rstrip("/")
        parts = path.split("/")

        if len(parts) > 1:
            parent_path = "/".join(
                parts[:-1]
            )

            result.append(
                urlunparse(
                    (
                        p.scheme,
                        p.netloc,
                        parent_path + "/",
                        "",
                        "",
                        "",
                    )
                )
            )

        if p.query:
            result.append(
                urlunparse(
                    (
                        p.scheme,
                        p.netloc,
                        p.path,
                        "",
                        "",
                        "",
                    )
                )
            )

        lower_path = path.lower()

        for token in [
            "/view",
            "/read",
            "/detail",
            "/article",
        ]:

            if token in lower_path:

                new_path = re.sub(
                    token,
                    "/list",
                    path,
                    flags=re.I,
                )

                result.append(
                    urlunparse(
                        (
                            p.scheme,
                            p.netloc,
                            new_path,
                            "",
                            "",
                            "",
                        )
                    )
                )

    except Exception:
        pass

    return list(
        dict.fromkeys(result)
    )


# ============================================================
# 결과 기본값
# ============================================================

def make_result(name, homepage):

    return {
        "기관명": name,
        "홈페이지": homepage,
        "접속상태": "",
        "탐색페이지수": 0,
        "CMS": "",
        "iframe": "",
        "JS": "",
        "게시판후보수": 0,
        "게시물후보수": 0,
        "게시판URL": "",
        "게시물URL": "",
        "iframeURL": "",
        "API후보": "",
        "탐색방법": "",
        "신뢰도": "",
        "실패단계": "",
        "오류": "",
    }


# ============================================================
# 기관 1개 분석
# ============================================================

async def analyze_institution(
    session,
    name,
    homepage,
):

    result = make_result(
        name,
        homepage,
    )

    try:

        return await asyncio.wait_for(
            analyze_core(
                session,
                result,
            ),
            timeout=INSTITUTION_TIMEOUT,
        )

    except asyncio.TimeoutError:

        result["실패단계"] = (
            "기관 전체 분석시간 초과"
        )

        result["오류"] = (
            str(INSTITUTION_TIMEOUT)
            + "초 초과"
        )

        result["탐색방법"] = (
            "V6.2 기관시간제한"
        )

        return result

    except Exception as e:

        result["실패단계"] = "분석 오류"
        result["오류"] = type(e).__name__

        return result


# ============================================================
# 실제 분석
# ============================================================

async def analyze_core(
    session,
    result,
):

    homepage = normalize_url(
        result["홈페이지"]
    )

    first = await fetch(
        session,
        homepage,
    )

    if not first or not first["ok"]:

        error = "접속 실패"

        if first:
            error = first["error"]

        result["접속상태"] = "실패"
        result["실패단계"] = "홈페이지"
        result["오류"] = error
        result["탐색방법"] = "V6.2"

        return result

    result["접속상태"] = (
        "정상:"
        + str(first["status"])
    )

    homepage = first["url"]

    queue = deque()
    queue.append(homepage)

    visited = set()

    boards = []
    posts = []
    iframes = []
    apis = []
    cms = []
    js_findings = []

    pages = 0

    while queue and pages < MAX_PAGES:

        current = clean_url(
            queue.popleft()
        )

        if current in visited:
            continue

        if not same_domain(
            homepage,
            current,
        ):
            continue

        visited.add(current)

        data = await fetch(
            session,
            current,
        )

        if not data or not data["ok"]:
            continue

        content_type = data[
            "content_type"
        ]

        if (
            "html" not in content_type
            and "xhtml" not in content_type
            and content_type
        ):
            continue

        html = data["text"]

        if not html:
            continue

        pages += 1
        result["탐색페이지수"] = pages

        # ----------------------------------------------------
        # CMS
        # ----------------------------------------------------

        cms.extend(
            detect_cms(html)
        )

        # ----------------------------------------------------
        # iframe
        # ----------------------------------------------------

        for iframe in detect_iframe(
            html,
            current,
        ):

            if iframe not in iframes:
                iframes.append(iframe)

                if same_domain(
                    homepage,
                    iframe,
                ):
                    if iframe not in visited:
                        queue.append(iframe)

        # ----------------------------------------------------
        # JS
        # ----------------------------------------------------

        js_findings.extend(
            analyze_js(html)
        )

        # ----------------------------------------------------
        # 링크
        # ----------------------------------------------------

        links = extract_links(
            html,
            current,
        )

        for item in links:

            link = item["url"]
            text = item["text"]

            if not same_domain(
                homepage,
                link,
            ):
                continue

            # ----------------------------------------------
            # 게시판 후보
            # ----------------------------------------------

            score, reasons = board_score(
                link,
                text,
            )

            if score >= 5:

                exists = False

                for old in boards:
                    if old["url"] == link:
                        exists = True
                        break

                if not exists:

                    boards.append(
                        {
                            "url": link,
                            "text": text,
                            "score": score,
                            "reasons": ",".join(
                                reasons
                            ),
                        }
                    )

            # ----------------------------------------------
            # 게시물 후보
            # ----------------------------------------------

            if looks_like_post(link):

                if link not in posts:
                    posts.append(link)

            # ----------------------------------------------
            # 우선 탐색
            # ----------------------------------------------

            target = (
                text
                + " "
                + link
            ).lower()

            priority = 0

            for word in NOTICE_WORDS:
                if word.lower() in target:
                    priority += 5

            for word in BOARD_URL_WORDS:
                if word.lower() in link.lower():
                    priority += 2

            if "k2web" in target:
                priority += 5

            if "nttid" in target:
                priority += 5

            if priority >= 3:
                if link not in visited:
                    queue.append(link)

        # ----------------------------------------------------
        # 초기 메뉴 우선 탐색
        # ----------------------------------------------------

        if pages <= 5:

            for item in links:

                link = item["url"]
                text = item["text"]

                target = (
                    text
                    + " "
                    + link
                ).lower()

                if any(
                    word.lower() in target
                    for word in BOARD_WORDS
                ):

                    if link not in visited:
                        queue.append(link)

        # ----------------------------------------------------
        # Queue 폭주 방지
        # ----------------------------------------------------

        limit = MAX_PAGES * 3

        if len(queue) > limit:

            queue = deque(
                list(queue)[:limit]
            )

    # ========================================================
    # 후보 정리
    # ========================================================

    boards.sort(
        key=lambda x: x["score"],
        reverse=True,
    )

    boards = boards[:10]

    posts = list(
        dict.fromkeys(posts)
    )[:10]

    # ========================================================
    # 실제 게시판 검증
    # ========================================================

    verified = []

    for candidate in boards[:MAX_VERIFY]:

        check = await fetch(
            session,
            candidate["url"],
        )

        if not check or not check["ok"]:
            continue

        score, reasons = board_score(
            candidate["url"],
            candidate["text"],
            check["text"],
        )

        if score >= 7:

            verified.append(
                {
                    "url": candidate["url"],
                    "score": score,
                    "reasons": ",".join(
                        reasons
                    ),
                }
            )

    # ========================================================
    # 게시물 역추적
    # ========================================================

    for post in posts[:5]:

        parents = parent_candidates(
            post
        )

        for parent in parents[:3]:

            if not same_domain(
                homepage,
                parent,
            ):
                continue

            check = await fetch(
                session,
                parent,
            )

            if not check or not check["ok"]:
                continue

            score, reasons = board_score(
                parent,
                "",
                check["text"],
            )

            if score >= 7:

                verified.append(
                    {
                        "url": parent,
                        "score": score,
                        "reasons": ",".join(
                            reasons
                        ),
                    }
                )

    # ========================================================
    # 최종 게시판 URL
    # ========================================================

    verified.sort(
        key=lambda x: x["score"],
        reverse=True,
    )

    unique_verified = []
    seen_urls = set()

    for item in verified:

        url = item["url"]

        if url in seen_urls:
            continue

        seen_urls.add(url)
        unique_verified.append(item)

    # ========================================================
    # 결과 입력
    # ========================================================

    result["CMS"] = ",".join(
        list(dict.fromkeys(cms))
    )

    result["iframe"] = str(
        len(iframes)
    )

    result["JS"] = ",".join(
        list(dict.fromkeys(js_findings))
    )

    result["게시판후보수"] = len(boards)
    result["게시물후보수"] = len(posts)

    if unique_verified:

        result["게시판URL"] = (
            unique_verified[0]["url"]
        )

        result["신뢰도"] = (
            "높음"
        )

        result["탐색방법"] = (
            "V6.2 게시판 직접검증"
        )

    elif posts:

        result["게시물URL"] = posts[0]

        result["신뢰도"] = (
            "중간"
        )

        result["탐색방법"] = (
            "V6.2 게시물 역추적"
        )

    elif boards:

        result["게시판URL"] = (
            boards[0]["url"]
        )

        result["신뢰도"] = (
            "후보"
        )

        result["탐색방법"] = (
            "V6.2 게시판 후보"
        )

    elif iframes:

        result["iframeURL"] = (
            iframes[0]
        )

        result["신뢰도"] = (
            "iframe"
        )

        result["탐색방법"] = (
            "V6.2 iframe 탐색"
        )

    elif js_findings:

        result["API후보"] = ",".join(
            list(dict.fromkeys(
                js_findings
            ))
        )

        result["신뢰도"] = (
            "JS"
        )

        result["탐색방법"] = (
            "V6.2 JS 구조분석"
        )

    else:

        result["신뢰도"] = (
            "미발견"
        )

        result["탐색방법"] = (
            "V6.2 전체탐색"
        )

    return result


# ============================================================
# 전체 분석
# ============================================================

async def run():

    print("=" * 70)
    print("V6.2 미발견 기관 게시판 구조 분석 시작")
    print("=" * 70)

    try:
        df = pd.read_excel(
            INPUT_FILE
        )
    except Exception as e:
        print("입력파일 오류:", e)
        return

    required = [
        "기관명",
        "URL",
    ]

    for col in required:

        if col not in df.columns:
            print(
                "필수 컬럼 없음:",
                col,
            )
            return

    total = len(df)

    print("분석 대상 기관:", total)
    print("동시 분석:", CONCURRENCY)
    print("기관별 제한:", INSTITUTION_TIMEOUT, "초")
    print("기관당 최대 페이지:", MAX_PAGES)
    print()

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY * 2,
        limit_per_host=2,
        ssl=SSL_CONTEXT,
    )

    async with aiohttp.ClientSession(
        connector=connector
    ) as session:

        semaphore = asyncio.Semaphore(
            CONCURRENCY
        )

        results = [None] * total

        async def worker(index, row):

            async with semaphore:

                name = str(
                    row["기관명"]
                )

                homepage = str(
                    row["URL"]
                )

                print(
                    f"[{index + 1}/{total}] "
                    f"{name}"
                )

                result = await analyze_institution(
                    session,
                    name,
                    homepage,
                )

                results[index] = result

                status = result[
                    "접속상태"
                ]

                board = result[
                    "게시판URL"
                ]

                print(
                    f"    → {status} | "
                    f"페이지 {result['탐색페이지수']} | "
                    f"후보 {result['게시판후보수']} | "
                    f"게시판 {'발견' if board else '미발견'}"
                )

        tasks = []

        for index, row in df.iterrows():

            tasks.append(
                asyncio.create_task(
                    worker(
                        index,
                        row,
                    )
                )
            )

        await asyncio.gather(
            *tasks
        )

    # ========================================================
    # 결과 저장
    # ========================================================

    final_results = []

    for result in results:

        if result is not None:
            final_results.append(
                result
            )

    output = pd.DataFrame(
        final_results
    )

    output.to_excel(
        OUTPUT_FILE,
        index=False,
    )

    # ========================================================
    # 요약
    # ========================================================

    normal = len(
        output[
            output["접속상태"].astype(str).str.startswith(
                "정상"
            )
        ]
    )

    discovered = len(
        output[
            output["게시판URL"].astype(str).str.len() > 5
        ]
    )

    post_found = len(
        output[
            output["게시물URL"].astype(str).str.len() > 5
        ]
    )

    print()
    print("=" * 70)
    print("V6.2 분석 완료")
    print("=" * 70)
    print("전체:", len(output))
    print("정상 접속:", normal)
    print("게시판 발견:", discovered)
    print("게시물 역추적:", post_found)
    print("결과 파일:", OUTPUT_FILE)
    print("=" * 70)


# ============================================================
# 실행
# ============================================================

if __name__ == "__main__":
    asyncio.run(run())
