import asyncio
import aiohttp
import pandas as pd
import re
import os
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from collections import Counter
from openpyxl import load_workbook


# =========================================================
# 설정
# =========================================================

INPUT_FILE = "boards_missing.xlsx"
OUTPUT_FILE = "missing_structure_analysis.xlsx"

CONCURRENCY = 10
TIMEOUT = 20

# 기관별 최대 탐색 페이지
MAX_PAGES = 20

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0 Safari/537.36"
    )
}


# =========================================================
# 게시판 관련 키워드
# =========================================================

BOARD_WORDS = [
    "공지사항",
    "공지",
    "알림마당",
    "알림",
    "소식",
    "게시판",
    "자료실",
    "자료",
    "참여",
    "국민참여",
    "시민참여",
    "고객참여",
    "소통",
    "커뮤니티",
    "뉴스",
    "보도자료",
    "공모",
    "공모전",
    "설문",
    "이벤트",
    "새소식",
    "입찰",
    "채용",
    "고시공고",
    "공지·공고",
    "정보마당",
    "경영공시",
    "전자민원",
]


POST_WORDS = [
    "view",
    "detail",
    "read",
    "article",
    "board",
    "bbs",
    "ntt",
    "seq",
    "idx",
    "no=",
    "articleid",
    "boardid",
    "wr_id",
]


# =========================================================
# 오류 페이지 키워드
# =========================================================

ERROR_WORDS = [
    "페이지를 찾을 수 없습니다",
    "페이지가 없습니다",
    "존재하지 않는 페이지",
    "요청하신 페이지",
    "404",
    "오류가 발생",
    "에러가 발생",
    "error",
    "not found",
    "접근할 수 없습니다",
    "서비스 이용에 불편",
]


# =========================================================
# CMS 패턴
# =========================================================

CMS_PATTERNS = {

    "K2Web": [
        "k2web",
        "k2webwizard",
        "ntt",
        "siteId",
        "fnctId",
    ],

    "Drupal": [
        "drupal",
        "/node/",
        "views-row",
    ],

    "WordPress": [
        "wp-content",
        "wp-includes",
        "wordpress",
    ],

    "Joomla": [
        "joomla",
        "/component/",
    ],

    "그누보드": [
        "gnuboard",
        "bo_table",
        "wr_id",
    ],

    "Apache/일반": [
        "apache",
    ],
}


# =========================================================
# 유틸
# =========================================================

def clean_text(text):

    if not text:
        return ""

    text = re.sub(r"\s+", " ", str(text))

    return text.strip()


def same_domain(url1, url2):

    try:

        host1 = urlparse(url1).netloc.lower()
        host2 = urlparse(url2).netloc.lower()

        return host1 == host2

    except Exception:

        return False


def is_html_response(content_type):

    if not content_type:
        return True

    content_type = content_type.lower()

    return (
        "text/html" in content_type
        or "application/xhtml" in content_type
    )


# =========================================================
# CMS 탐지
# =========================================================

def detect_cms(html, url):

    text = (
        str(html)
        + " "
        + str(url)
    ).lower()

    detected = []

    for cms, patterns in CMS_PATTERNS.items():

        for pattern in patterns:

            if pattern.lower() in text:

                detected.append(cms)

                break

    if not detected:

        return "일반 HTML/미확인"

    return ", ".join(
        dict.fromkeys(detected)
    )


# =========================================================
# iframe 탐지
# =========================================================

def detect_iframe(soup):

    iframes = soup.find_all("iframe")

    if not iframes:

        return "없음"

    sources = []

    for iframe in iframes:

        src = iframe.get("src", "").strip()

        if src:

            sources.append(src)

    if sources:

        return (
            f"있음({len(iframes)}개): "
            + " | ".join(sources[:3])
        )

    return f"있음({len(iframes)}개)"


# =========================================================
# JS 의존도
# =========================================================

def detect_js_dependency(soup, html):

    script_count = len(
        soup.find_all("script")
    )

    js_patterns = [
        "javascript:",
        "onclick=",
        "fetch(",
        "axios",
        "ajax",
        "$.ajax",
        "$.get",
        "$.post",
        "xmlhttprequest",
        "addeventlistener",
    ]

    found = []

    lower_html = html.lower()

    for pattern in js_patterns:

        if pattern.lower() in lower_html:

            found.append(pattern)

    if script_count >= 20 or len(found) >= 2:

        return "높음"

    if script_count >= 8 or len(found) >= 1:

        return "중간"

    return "낮음"


# =========================================================
# 오류 페이지 탐지
# =========================================================

def detect_error_page(text, status):

    if status >= 400:

        return True

    lower = text.lower()

    for word in ERROR_WORDS:

        if word.lower() in lower:

            return True

    return False


# =========================================================
# URL 분류
# =========================================================

def classify_url(url, text):

    lower_url = url.lower()
    lower_text = text.lower()

    board_score = 0
    post_score = 0

    # 텍스트 키워드
    for word in BOARD_WORDS:

        if word.lower() in lower_text:

            board_score += 1

    # 게시물 URL 키워드
    for word in POST_WORDS:

        if word.lower() in lower_url:

            post_score += 1

    # URL 패턴
    url_patterns = [

        r"/board",
        r"/bbs",
        r"/notice",
        r"/community",
        r"/news",
        r"/data",
        r"/archive",
        r"/particip",
        r"/event",
        r"/survey",
        r"ntt",
        r"boardid",
        r"bo_table",
        r"wr_id",
    ]

    for pattern in url_patterns:

        if re.search(
            pattern,
            lower_url
        ):

            board_score += 2

    if post_score >= 2:

        return "게시물 상세 후보"

    if board_score >= 3:

        return "게시판 후보"

    return "일반 페이지"


# =========================================================
# 링크 추출
# =========================================================

def extract_links(base_url, soup):

    results = []

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

        href_lower = href.lower()

        if href.startswith("#"):

            continue

        if href_lower.startswith(
            "javascript:"
        ):

            continue

        if href_lower.startswith(
            "mailto:"
        ):

            continue

        if href_lower.startswith(
            "tel:"
        ):

            continue

        try:

            absolute = urljoin(
                base_url,
                href
            )

        except Exception:

            continue

        if not same_domain(
            base_url,
            absolute
        ):

            continue

        text = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

        results.append({
            "url": absolute,
            "text": text,
        })

    return results


# =========================================================
# 페이지 접속
# =========================================================

async def fetch_page(
    session,
    url
):

    try:

        async with session.get(

            url,

            headers=HEADERS,

            timeout=aiohttp.ClientTimeout(
                total=TIMEOUT
            ),

            allow_redirects=True,

            ssl=False,

        ) as response:

            status = response.status

            content_type = response.headers.get(
                "Content-Type",
                ""
            )

            final_url = str(
                response.url
            )

            if not is_html_response(
                content_type
            ):

                return {

                    "success": False,

                    "status": status,

                    "url": final_url,

                    "html": "",

                    "error":
                        f"비HTML 응답: {content_type}",

                }

            html = await response.text(
                errors="ignore"
            )

            return {

                "success": True,

                "status": status,

                "url": final_url,

                "html": html,

                "error": "",

            }

    except asyncio.TimeoutError:

        return {

            "success": False,
            "status": 0,
            "url": url,
            "html": "",
            "error": "Timeout",

        }

    except aiohttp.ClientConnectorError:

        return {

            "success": False,
            "status": 0,
            "url": url,
            "html": "",
            "error": "ConnectionError",

        }

    except aiohttp.ClientSSLError:

        return {

            "success": False,
            "status": 0,
            "url": url,
            "html": "",
            "error": "SSLError",

        }

    except Exception as e:

        return {

            "success": False,
            "status": 0,
            "url": url,
            "html": "",
            "error": type(e).__name__,

        }


# =========================================================
# 공통 후보 URL
# =========================================================

def generate_candidate_urls(homepage):

    candidates = []

    base = homepage.rstrip("/") + "/"

    paths = [

        "notice",
        "notices",

        "news",

        "board",
        "bbs",

        "community",

        "customer",
        "customer/notice",

        "contents",
        "content",

        "data",
        "archive",

        "information",
        "info",

        "participation",
        "particip",

        "citizen",

        "event",
        "events",

        "survey",

        "media",
        "press",

        "pds",
        "reference",
        "download",

        "communication",
        "sotong",

        "알림",
        "알림마당",
        "공지사항",
        "소식",
        "자료실",

    ]

    for path in paths:

        candidates.append(
            urljoin(
                base,
                path
            )
        )

    return candidates


# =========================================================
# 페이지 분석
# =========================================================

def analyze_page(
    url,
    html,
    status
):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    title = ""

    if soup.title:

        title = clean_text(
            soup.title.get_text(
                " ",
                strip=True
            )
        )

    visible_text = clean_text(
        soup.get_text(
            " ",
            strip=True
        )
    )

    links = extract_links(
        url,
        soup
    )

    board_candidates = []
    post_candidates = []

    for link in links:

        link_url = link["url"]

        link_text = link["text"]

        combined = (
            f"{link_text} {link_url}"
        )

        classification = classify_url(
            link_url,
            combined
        )

        if classification == "게시판 후보":

            board_candidates.append(
                (
                    link_text,
                    link_url
                )
            )

        elif classification == "게시물 상세 후보":

            post_candidates.append(
                (
                    link_text,
                    link_url
                )
            )

    page_type = classify_url(
        url,
        f"{title} {visible_text[:5000]}"
    )

    return {

        "title": title,

        "text": visible_text,

        "iframe": detect_iframe(
            soup
        ),

        "js": detect_js_dependency(
            soup,
            html
        ),

        "cms": detect_cms(
            html,
            url
        ),

        "links": links,

        "board_candidates":
            board_candidates,

        "post_candidates":
            post_candidates,

        "page_type":
            page_type,

        "error_page":
            detect_error_page(
                visible_text,
                status
            ),

    }


# =========================================================
# 기관 하나 분석
# =========================================================

async def analyze_institution(

    session,
    semaphore,
    institution,
    homepage,
    index

):

    result = {

        "기관명": institution,

        "URL": homepage,

        "접속상태": "",

        "HTTP상태": "",

        "최종URL": "",

        "CMS/프레임워크": "",

        "iframe 사용": "",

        "JS 의존도": "",

        "메뉴 링크 발견": 0,

        "공지사항 후보": 0,

        "게시판 URL 패턴": "",

        "게시물 상세 URL 발견": 0,

        "게시판 유형": "",

        "추정 원인": "",

        "6차 공략법": "",

        "탐색 페이지 수": 0,

        "발견 URL 예시": "",

        "오류": "",

    }

    async with semaphore:

        print(
            f"[{index:03d}] {institution}",
            flush=True
        )

        # =================================================
        # 홈페이지 접속
        # =================================================

        first = await fetch_page(
            session,
            homepage
        )

        if not first["success"]:

            result["접속상태"] = "실패"

            result["HTTP상태"] = (
                first["status"]
            )

            result["최종URL"] = (
                first["url"]
            )

            result["추정 원인"] = (
                first["error"]
            )

            result["6차 공략법"] = (
                "접속 문제 해결 후 재탐색"
            )

            result["오류"] = (
                first["error"]
            )

            return result

        # =================================================
        # 정상 접속
        # =================================================

        result["접속상태"] = "정상"

        result["HTTP상태"] = (
            first["status"]
        )

        result["최종URL"] = (
            first["url"]
        )

        first_analysis = analyze_page(

            first["url"],

            first["html"],

            first["status"]

        )

        result["CMS/프레임워크"] = (
            first_analysis["cms"]
        )

        result["iframe 사용"] = (
            first_analysis["iframe"]
        )

        result["JS 의존도"] = (
            first_analysis["js"]
        )

        # =================================================
        # 탐색 큐
        # =================================================

        queue = []

        visited = set()

        def add_url(url):

            if not url:
                return

            if url in visited:
                return

            if url in queue:
                return

            if len(queue) >= MAX_PAGES:
                return

            queue.append(url)

        # 홈페이지
        add_url(
            first["url"]
        )

        # 홈페이지 메뉴
        for link in first_analysis["links"]:

            text = link["text"]

            url = link["url"]

            combined = (
                f"{text} {url}"
            ).lower()

            if any(

                word.lower()
                in combined

                for word in BOARD_WORDS

            ):

                add_url(url)

        # 공통 URL
        for url in generate_candidate_urls(
            first["url"]
        ):

            add_url(url)

        # =================================================
        # BFS 탐색
        # =================================================

        pages_analyzed = 0

        all_board_candidates = []

        all_post_candidates = []

        all_links = []

        cms_counter = Counter()

        page_type_counter = Counter()

        discovered_urls = []

        board_pattern_counter = Counter()

        # -------------------------------------------------
        # 탐색 시작
        # -------------------------------------------------

        while (

            queue
            and
            pages_analyzed < MAX_PAGES

        ):

            current_url = queue.pop(0)

            if current_url in visited:

                continue

            visited.add(
                current_url
            )

            page = await fetch_page(
                session,
                current_url
            )

            if not page["success"]:

                continue

            pages_analyzed += 1

            analysis = analyze_page(

                page["url"],

                page["html"],

                page["status"]

            )

            discovered_urls.append(
                page["url"]
            )

            # -------------------------------------------------
            # CMS
            # -------------------------------------------------

            cms_counter[
                analysis["cms"]
            ] += 1

            # -------------------------------------------------
            # 페이지 유형
            # -------------------------------------------------

            page_type_counter[
                analysis["page_type"]
            ] += 1

            # -------------------------------------------------
            # 링크
            # -------------------------------------------------

            all_links.extend(
                analysis["links"]
            )

            # -------------------------------------------------
            # 게시판 후보
            # -------------------------------------------------

            for candidate in (
                analysis[
                    "board_candidates"
                ]
            ):

                all_board_candidates.append(
                    candidate
                )

                candidate_url = candidate[1]

                parsed = urlparse(
                    candidate_url
                )

                path = parsed.path.lower()

                if path:

                    board_pattern_counter[
                        path
                    ] += 1

            # -------------------------------------------------
            # 게시물 후보
            # -------------------------------------------------

            all_post_candidates.extend(
                analysis[
                    "post_candidates"
                ]
            )

            # -------------------------------------------------
            # 다음 링크 추가
            # -------------------------------------------------

            for link in analysis["links"]:

                link_url = link["url"]

                link_text = link["text"]

                combined = (
                    f"{link_text} {link_url}"
                ).lower()

                # 게시판 관련 링크만 다음 탐색
                if any(

                    word.lower()
                    in combined

                    for word in BOARD_WORDS

                ):

                    add_url(
                        link_url
                    )

        # =================================================
        # 중복 제거
        # =================================================

        unique_board_candidates = []

        seen_board_urls = set()

        for text, url in all_board_candidates:

            if url in seen_board_urls:

                continue

            seen_board_urls.add(url)

            unique_board_candidates.append(
                (text, url)
            )

        unique_post_candidates = []

        seen_post_urls = set()

        for text, url in all_post_candidates:

            if url in seen_post_urls:

                continue

            seen_post_urls.add(url)

            unique_post_candidates.append(
                (text, url)
            )

        # =================================================
        # 결과 입력
        # =================================================

        result["탐색 페이지 수"] = (
            pages_analyzed
        )

        result["메뉴 링크 발견"] = (
            len(all_links)
        )

        result["공지사항 후보"] = (
            len(unique_board_candidates)
        )

        result["게시물 상세 URL 발견"] = (
            len(unique_post_candidates)
        )

        # -------------------------------------------------
        # CMS 대표값
        # -------------------------------------------------

        if cms_counter:

            result["CMS/프레임워크"] = (
                cms_counter.most_common(1)[0][0]
            )

        # -------------------------------------------------
        # 게시판 URL 패턴
        # -------------------------------------------------

        if board_pattern_counter:

            top_patterns = (
                board_pattern_counter
                .most_common(5)
            )

            result["게시판 URL 패턴"] = (
                " | ".join(
                    p[0]
                    for p in top_patterns
                )
            )

        else:

            result["게시판 URL 패턴"] = (
                "미발견"
            )

        # -------------------------------------------------
        # 게시판 유형
        # -------------------------------------------------

        if unique_board_candidates:

            result["게시판 유형"] = (
                "게시판 후보 발견"
            )

        elif unique_post_candidates:

            result["게시판 유형"] = (
                "게시물 상세 후보만 발견"
            )

        else:

            result["게시판 유형"] = (
                "게시판 미발견"
            )

        # =================================================
        # URL 예시
        # =================================================

        examples = []

        for text, url in unique_board_candidates[:5]:

            if text:

                examples.append(
                    f"{text} → {url}"
                )

            else:

                examples.append(
                    url
                )

        if not examples:

            examples = discovered_urls[:5]

        result["발견 URL 예시"] = (
            " | ".join(examples)
        )

        # =================================================
        # 원인 및 6차 공략법
        # =================================================

        cms = result["CMS/프레임워크"]

        iframe = result["iframe 사용"]

        js = result["JS 의존도"]

        boards = len(
            unique_board_candidates
        )

        posts = len(
            unique_post_candidates
        )

        # -------------------------------------------------
        # 게시판을 찾은 경우
        # -------------------------------------------------

        if boards > 0:

            result["추정 원인"] = (
                "기존 탐색에서 게시판 후보 "
                "판정 기준 미충족 가능성"
            )

            result["6차 공략법"] = (
                "발견 후보를 실제 게시판인지 "
                "검증 후 boards.xlsx에 선택 등록"
            )

        # -------------------------------------------------
        # 게시물만 발견
        # -------------------------------------------------

        elif posts > 0:

            result["추정 원인"] = (
                "게시판 목록보다 게시물 상세 "
                "URL 구조가 먼저 노출되는 사이트"
            )

            result["6차 공략법"] = (
                "게시물 상세 URL에서 "
                "상위 게시판 URL 역추적"
            )

        # -------------------------------------------------
        # iframe
        # -------------------------------------------------

        elif "있음" in iframe:

            result["추정 원인"] = (
                "게시판이 iframe 내부에 "
                "별도 구성되어 있을 가능성"
            )

            result["6차 공략법"] = (
                "iframe src 직접 접속 및 "
                "iframe 내부 게시판 탐색"
            )

        # -------------------------------------------------
        # JS 높음
        # -------------------------------------------------

        elif js == "높음":

            result["추정 원인"] = (
                "JavaScript 기반 동적 메뉴/게시판 "
                "사용 가능성"
            )

            result["6차 공략법"] = (
                "API/AJAX 호출 URL 또는 "
                "렌더링된 페이지 구조 별도 분석"
            )

        # -------------------------------------------------
        # K2Web
        # -------------------------------------------------

        elif "K2Web" in cms:

            result["추정 원인"] = (
                "K2Web 계열 사이트 구조로 "
                "일반 링크 탐색만으로 누락 가능"
            )

            result["6차 공략법"] = (
                "K2Web의 fnctId/nttId/siteId "
                "구조 집중 탐색"
            )

        # -------------------------------------------------
        # 일반 HTML
        # -------------------------------------------------

        elif cms == "일반 HTML/미확인":

            result["추정 원인"] = (
                "게시판 URL이 일반적인 "
                "board/bbs/notice 패턴과 다를 가능성"
            )

            result["6차 공략법"] = (
                "메뉴 구조와 실제 게시물 링크를 "
                "중점적으로 역추적"
            )

        # -------------------------------------------------
        # 기타
        # -------------------------------------------------

        else:

            result["추정 원인"] = (
                "게시판 구조 확인 필요"
            )

            result["6차 공략법"] = (
                "사이트별 URL 구조를 "
                "개별 분석하여 공략"
            )

        return result


# =========================================================
# 메인
# =========================================================

async def main():

    print("=" * 70)

    print(
        "미발견 기관 구조 분석 시작"
    )

    print("=" * 70)

    # =====================================================
    # 입력파일 확인
    # =====================================================

    if not os.path.exists(
        INPUT_FILE
    ):

        raise FileNotFoundError(
            f"입력 파일이 없습니다: {INPUT_FILE}"
        )

    print(
        f"입력파일: {INPUT_FILE}"
    )

    print(
        f"출력파일: {OUTPUT_FILE}"
    )

    # =====================================================
    # Excel 읽기
    # =====================================================

    df = pd.read_excel(
        INPUT_FILE
    )

    print(
        f"입력 행 수: {len(df)}"
    )

    # =====================================================
    # 컬럼 확인
    # =====================================================

    print(
        "입력 컬럼:",
        list(df.columns)
    )

    # 기관명 컬럼
    institution_col = None

    for col in [
        "기관명",
        "기관",
        "기관명칭",
    ]:

        if col in df.columns:

            institution_col = col

            break

    if institution_col is None:

        raise ValueError(
            "기관명 컬럼을 찾을 수 없습니다."
        )

    # URL 컬럼
    url_col = None

    for col in [
        "URL",
        "url",
        "홈페이지",
        "홈페이지 URL",
    ]:

        if col in df.columns:

            url_col = col

            break

    if url_col is None:

        raise ValueError(
            "URL 컬럼을 찾을 수 없습니다."
        )

    # =====================================================
    # 분석 대상 생성
    # =====================================================

    targets = []

    for _, row in df.iterrows():

        institution = str(
            row[institution_col]
        ).strip()

        homepage = str(
            row[url_col]
        ).strip()

        if not institution:

            continue

        if homepage.lower() in [
            "",
            "nan",
            "none",
        ]:

            continue

        targets.append(
            (
                institution,
                homepage
            )
        )

    print(
        f"실제 분석 대상: {len(targets)}개"
    )

    # =====================================================
    # HTTP 세션
    # =====================================================

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        limit_per_host=3,
        ssl=False,
    )

    timeout = aiohttp.ClientTimeout(
        total=TIMEOUT
    )

    semaphore = asyncio.Semaphore(
        CONCURRENCY
    )

    results = []

    async with aiohttp.ClientSession(

        connector=connector,

        timeout=timeout,

        headers=HEADERS,

    ) as session:

        tasks = []

        for index, (
            institution,
            homepage
        ) in enumerate(
            targets,
            start=1
        ):

            tasks.append(
                analyze_institution(

                    session,

                    semaphore,

                    institution,

                    homepage,

                    index,

                )
            )

        # =================================================
        # 병렬 실행
        # =================================================

        raw_results = await asyncio.gather(

            *tasks,

            return_exceptions=True

        )

        # =================================================
        # 결과 처리
        # =================================================

        for index, item in enumerate(
            raw_results,
            start=1
        ):

            if isinstance(
                item,
                Exception
            ):

                institution = (
                    targets[index - 1][0]
                )

                homepage = (
                    targets[index - 1][1]
                )

                print(
                    f"[{index:03d}] 예외 발생: "
                    f"{institution}: {item}"
                )

                results.append({

                    "기관명":
                        institution,

                    "URL":
                        homepage,

                    "접속상태":
                        "분석오류",

                    "HTTP상태":
                        "",

                    "최종URL":
                        "",

                    "CMS/프레임워크":
                        "",

                    "iframe 사용":
                        "",

                    "JS 의존도":
                        "",

                    "메뉴 링크 발견":
                        0,

                    "공지사항 후보":
                        0,

                    "게시판 URL 패턴":
                        "",

                    "게시물 상세 URL 발견":
                        0,

                    "게시판 유형":
                        "분석오류",

                    "추정 원인":
                        type(item).__name__,

                    "6차 공략법":
                        "개별 재분석",

                    "탐색 페이지 수":
                        0,

                    "발견 URL 예시":
                        "",

                    "오류":
                        str(item),

                })

            else:

                results.append(item)

    # =====================================================
    # DataFrame
    # =====================================================

    result_df = pd.DataFrame(
        results
    )

    # =====================================================
    # 결과 정렬
    # =====================================================

    desired_columns = [

        "기관명",
        "URL",
        "접속상태",
        "HTTP상태",
        "최종URL",
        "CMS/프레임워크",
        "iframe 사용",
        "JS 의존도",
        "메뉴 링크 발견",
        "공지사항 후보",
        "게시판 URL 패턴",
        "게시물 상세 URL 발견",
        "게시판 유형",
        "추정 원인",
        "6차 공략법",
        "탐색 페이지 수",
        "발견 URL 예시",
        "오류",

    ]

    result_df = result_df[
        [
            col
            for col in desired_columns
            if col in result_df.columns
        ]
    ]

    # =====================================================
    # Excel 저장
    # =====================================================

    result_df.to_excel(
        OUTPUT_FILE,
        index=False,
        engine="openpyxl"
    )

    # =====================================================
    # Excel 서식
    # =====================================================

    try:

        wb = load_workbook(
            OUTPUT_FILE
        )

        ws = wb.active

        # 첫 행 고정
        ws.freeze_panes = "A2"

        # 자동 필터
        ws.auto_filter.ref = (
            ws.dimensions
        )

        # 열 너비
        widths = {

            "A": 25,
            "B": 45,
            "C": 12,
            "D": 12,
            "E": 45,
            "F": 25,
            "G": 30,
            "H": 12,
            "I": 15,
            "J": 15,
            "K": 45,
            "L": 18,
            "M": 25,
            "N": 45,
            "O": 55,
            "P": 15,
            "Q": 80,
            "R": 30,

        }

        for col, width in widths.items():

            ws.column_dimensions[
                col
            ].width = width

        wb.save(
            OUTPUT_FILE
        )

    except Exception as e:

        print(
            "Excel 서식 적용 중 경고:",
            e
        )

    # =====================================================
    # 최종 검증
    # =====================================================

    if not os.path.exists(
        OUTPUT_FILE
    ):

        raise RuntimeError(
            "결과 Excel 파일 생성에 실패했습니다."
        )

    file_size = os.path.getsize(
        OUTPUT_FILE
    )

    if file_size <= 0:

        raise RuntimeError(
            "결과 Excel 파일 크기가 0입니다."
        )

    # =====================================================
    # 요약
    # =====================================================

    normal_count = len(
        result_df[
            result_df["접속상태"]
            == "정상"
        ]
    )

    fail_count = len(
        result_df[
            result_df["접속상태"]
            == "실패"
        ]
    )

    error_count = len(
        result_df[
            result_df["접속상태"]
            == "분석오류"
        ]
    )

    board_count = len(
        result_df[
            result_df["공지사항 후보"]
            > 0
        ]
    )

    post_count = len(
        result_df[
            result_df[
                "게시물 상세 URL 발견"
            ]
            > 0
        ]
    )

    print()
    print("=" * 70)
    print(
        "미발견 기관 구조 분석 완료"
    )
    print("=" * 70)

    print(
        f"전체 분석 기관 : {len(result_df)}개"
    )

    print(
        f"정상 접속      : {normal_count}개"
    )

    print(
        f"접속 실패      : {fail_count}개"
    )

    print(
        f"분석 오류      : {error_count}개"
    )

    print(
        f"게시판 후보 발견: {board_count}개"
    )

    print(
        f"게시물 URL 발견 : {post_count}개"
    )

    print(
        f"결과 파일       : {OUTPUT_FILE}"
    )

    print(
        f"파일 크기       : {file_size:,} bytes"
    )

    print("=" * 70)


# =========================================================
# 실행
# =========================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
