import os
import re
import json
import asyncio
from urllib.parse import urljoin, urlparse, urldefrag, parse_qs

import aiohttp
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook


# ============================================================
# 기본 설정
# ============================================================

URL_FILE = "url_완성.xlsx"
BOARDS_FILE = "boards.xlsx"
MISSING_FILE = "boards_missing.xlsx"
SEEN_FILE = "seen_posts.json"

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")

# ------------------------------------------------------------
# 게시판 강제 재탐색
#
# GitHub Actions에서
# FORCE_DISCOVER_BOARDS=true
# 로 설정하면 boards.xlsx가 있어도 재탐색한다.
# ------------------------------------------------------------

FORCE_DISCOVER_BOARDS = (
    os.getenv(
        "FORCE_DISCOVER_BOARDS",
        "false"
    ).lower()
    in ("true", "1", "yes")
)

# ------------------------------------------------------------
# 실제 모니터링 키워드
# ------------------------------------------------------------

KEYWORDS = [
    "설문조사",
    "시민참여",
    "국민참여",
    "공모전",
]

# ============================================================
# 게시판 관련 단어
# ============================================================

BOARD_WORDS = [
    "공지사항",
    "공지",
    "알림마당",
    "알림",
    "새소식",
    "소식",
    "게시판",
    "국민참여",
    "시민참여",
    "참여",
    "공모전",
    "공모",
    "설문조사",
    "설문",
    "고시공고",
    "고시·공고",
    "고시",
    "공고",
    "입찰",
    "뉴스",
    "보도자료",
    "자료실",
    "소식지",
    "행사",
    "이벤트",
    "채용",
    "교육",
]

SITE_MAP_WORDS = [
    "사이트맵",
    "사이트 맵",
    "전체메뉴",
    "전체 메뉴",
    "메뉴",
    "홈페이지맵",
    "이용안내",
]

BOARD_URL_PATTERNS = [
    "board",
    "bbs",
    "notice",
    "news",
    "community",
    "particip",
    "participation",
    "event",
    "survey",
    "list",
    "pds",
    "data",
    "contents",
    "content",
    "menu",
    "article",
    "program",
]

DETAIL_PATTERNS = [
    "view",
    "detail",
    "read",
    "articleview",
    "articleView",
    "nttId",
    "nttid",
    "bbsNo",
    "bbsno",
    "boardNo",
    "boardno",
    "seq",
    "idx",
    "wr_id",
    "postid",
    "newsid",
]

# ============================================================
# 탐색 설정
# ============================================================

MAX_DEPTH = 3

# 기관별 최종 게시판 최대 개수
MAX_BOARDS_PER_ORG = 10

# 기관별 최대 페이지 탐색
MAX_PAGES_PER_ORG = 30

# 한 페이지에서 가져올 링크 최대 수
MAX_LINKS_PER_PAGE = 250

# 동시에 탐색할 기관 수
CONCURRENCY = 12

# HTTP timeout
REQUEST_TIMEOUT = 18

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}


# ============================================================
# URL 함수
# ============================================================

def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()

    if not url:
        return ""

    if not re.match(
        r"^https?://",
        url,
        re.I,
    ):
        url = "https://" + url

    url = urldefrag(url)[0]

    parsed = urlparse(url)

    netloc = parsed.netloc.lower()

    if (
        netloc.endswith(":80")
        and parsed.scheme == "http"
    ):
        netloc = netloc[:-3]

    if (
        netloc.endswith(":443")
        and parsed.scheme == "https"
    ):
        netloc = netloc[:-4]

    path = parsed.path or "/"

    path = re.sub(
        r"/+",
        "/",
        path,
    )

    return parsed._replace(
        netloc=netloc,
        path=path,
    ).geturl()


def same_domain(url1, url2):
    try:
        a = urlparse(url1).netloc.lower()
        b = urlparse(url2).netloc.lower()

        a = re.sub(
            r"^www\.",
            "",
            a,
        )

        b = re.sub(
            r"^www\.",
            "",
            b,
        )

        return a == b

    except Exception:
        return False


def is_http(url):
    return bool(
        url
        and url.lower().startswith(
            (
                "http://",
                "https://",
            )
        )
    )


def clean_text(text):
    if not text:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(text),
    ).strip()


# ============================================================
# 상세 페이지 여부
# ============================================================

def is_detail_url(url):
    lower = url.lower()

    parsed = urlparse(url)

    # URL path에 상세 패턴
    for pattern in DETAIL_PATTERNS:

        if pattern.lower() in lower:
            return True

    # query parameter 검사
    try:

        params = parse_qs(
            parsed.query
        )

        detail_keys = [
            "seq",
            "idx",
            "no",
            "nttid",
            "article",
            "articleno",
            "bbsno",
            "boardno",
            "wr_id",
            "uid",
            "postid",
            "newsid",
            "key",
        ]

        for key in params.keys():

            if key.lower() in [
                x.lower()
                for x in detail_keys
            ]:

                # key는 예외적으로
                # contents.do?key=123 형태가
                # 게시판 메뉴일 수도 있으므로
                # URL 전체를 보고 추가 판단
                if (
                    key.lower() == "key"
                    and (
                        "contents.do"
                        in parsed.path.lower()
                        or "menu.do"
                        in parsed.path.lower()
                    )
                ):
                    continue

                return True

    except Exception:
        pass

    return False


# ============================================================
# 게시판 후보 점수
# ============================================================

def score_candidate(
    url,
    text="",
    html=None,
):
    score = 0
    reasons = []

    lower_url = url.lower()
    lower_text = text.lower()

    # --------------------------------------------------------
    # 메뉴명
    # --------------------------------------------------------

    for word in BOARD_WORDS:

        if word.lower() in lower_text:

            score += 5
            reasons.append(
                f"메뉴:{word}"
            )

    # --------------------------------------------------------
    # URL
    # --------------------------------------------------------

    for pattern in BOARD_URL_PATTERNS:

        if pattern.lower() in lower_url:

            score += 2
            reasons.append(
                f"url:{pattern}"
            )

    # --------------------------------------------------------
    # 사이트맵
    # --------------------------------------------------------

    for word in SITE_MAP_WORDS:

        if word.lower() in lower_text:

            score += 2
            reasons.append(
                f"메뉴탐색:{word}"
            )

    # --------------------------------------------------------
    # 상세 URL 감점
    # --------------------------------------------------------

    if is_detail_url(url):

        score -= 10
        reasons.append(
            "상세페이지가능성"
        )

    # --------------------------------------------------------
    # HTML 검사
    # --------------------------------------------------------

    if html:

        soup = BeautifulSoup(
            html,
            "html.parser",
        )

        visible_text = clean_text(
            soup.get_text(
                " ",
                strip=True,
            )
        ).lower()

        # 게시판에서 흔한 단어
        structure_words = [
            "번호",
            "제목",
            "등록일",
            "작성일",
            "조회수",
            "작성자",
            "첨부파일",
            "목록",
        ]

        hits = sum(
            1
            for word in structure_words
            if word in visible_text
        )

        if hits >= 1:

            score += 3
            reasons.append(
                "게시판용어"
            )

        if hits >= 3:

            score += 5
            reasons.append(
                "게시판구조"
            )

        if hits >= 5:

            score += 5
            reasons.append(
                "게시판구조강함"
            )

        # table
        tables = soup.find_all(
            "table"
        )

        if tables:

            score += 2
            reasons.append(
                "table"
            )

        # 목록 형태의 링크가 여러 개 있는지
        links = soup.find_all(
            "a",
            href=True,
        )

        detail_like = 0

        for a in links:

            href = a.get(
                "href",
                "",
            ).lower()

            if any(
                x.lower() in href
                for x in DETAIL_PATTERNS
            ):

                detail_like += 1

        if detail_like >= 2:

            score += 3
            reasons.append(
                "게시글링크"
            )

        if detail_like >= 5:

            score += 4
            reasons.append(
                "게시글링크다수"
            )

    return score, reasons


# ============================================================
# HTML fetch
# ============================================================

async def fetch(
    session,
    url,
):
    try:

        async with session.get(
            url,
            allow_redirects=True,
            ssl=False,
            timeout=aiohttp.ClientTimeout(
                total=REQUEST_TIMEOUT
            ),
        ) as response:

            if response.status >= 400:
                return None, None

            content_type = response.headers.get(
                "Content-Type",
                "",
            ).lower()

            # HTML이 아닌 파일은 제외
            if content_type and not any(
                x in content_type
                for x in [
                    "text/html",
                    "application/xhtml",
                ]
            ):

                return None, None

            html = await response.text(
                errors="ignore"
            )

            return (
                str(response.url),
                html,
            )

    except Exception:
        return None, None


# ============================================================
# 링크 추출
# ============================================================

def extract_links(
    base_url,
    html,
):
    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    results = []
    seen = set()

    # --------------------------------------------------------
    # 일반 <a>
    # --------------------------------------------------------

    for a in soup.find_all(
        "a",
        href=True,
    ):

        href = a.get(
            "href",
            "",
        ).strip()

        if not href:
            continue

        if href.lower().startswith(
            (
                "javascript:",
                "mailto:",
                "tel:",
                "#",
            )
        ):
            continue

        url = normalize_url(
            urljoin(
                base_url,
                href,
            )
        )

        if not is_http(url):
            continue

        if not same_domain(
            base_url,
            url,
        ):
            continue

        lower = url.lower()

        # 파일 제외
        if lower.endswith(
            (
                ".jpg",
                ".jpeg",
                ".png",
                ".gif",
                ".svg",
                ".pdf",
                ".hwp",
                ".hwpx",
                ".xls",
                ".xlsx",
                ".zip",
                ".doc",
                ".docx",
                ".ppt",
                ".pptx",
                ".mp4",
                ".mp3",
            )
        ):
            continue

        if url in seen:
            continue

        seen.add(url)

        text = clean_text(
            a.get_text(
                " ",
                strip=True,
            )
        )

        # 상위 요소 텍스트
        parent_text = ""

        parent = a.parent

        if parent:

            parent_text = clean_text(
                parent.get_text(
                    " ",
                    strip=True,
                )
            )

        combined = clean_text(
            f"{text} {parent_text}"
        )

        results.append(
            {
                "url": url,
                "text": combined[:500],
            }
        )

        if len(results) >= MAX_LINKS_PER_PAGE:
            break

    # --------------------------------------------------------
    # iframe / frame
    # --------------------------------------------------------

    for frame in soup.find_all(
        [
            "iframe",
            "frame",
        ],
        src=True,
    ):

        src = frame.get(
            "src",
            "",
        ).strip()

        if not src:
            continue

        url = normalize_url(
            urljoin(
                base_url,
                src,
            )
        )

        if not is_http(url):
            continue

        if not same_domain(
            base_url,
            url,
        ):
            continue

        if url in seen:
            continue

        seen.add(url)

        results.append(
            {
                "url": url,
                "text": "iframe 게시판",
            }
        )

    return results


# ============================================================
# 기관 1곳 탐색
# ============================================================

async def discover_one(
    session,
    institution,
):
    org_name = institution["기관명"]
    homepage = normalize_url(
        institution["URL"]
    )

    if not homepage:

        return {
            "institution": institution,
            "boards": [],
            "reason": "홈페이지 URL 없음",
        }

    # queue:
    # url, depth, source_text
    queue = [
        (
            homepage,
            0,
            "홈페이지",
        )
    ]

    visited = set()
    candidates = {}

    pages = 0

    homepage_failed = False

    while (
        queue
        and pages < MAX_PAGES_PER_ORG
    ):

        current_url, depth, source_text = (
            queue.pop(0)
        )

        current_url = normalize_url(
            current_url
        )

        if current_url in visited:
            continue

        if depth > MAX_DEPTH:
            continue

        visited.add(
            current_url
        )

        final_url, html = await fetch(
            session,
            current_url,
        )

        if not html:

            if (
                depth == 0
                and current_url == homepage
            ):
                homepage_failed = True

            continue

        pages += 1

        # ----------------------------------------------------
        # 현재 페이지 자체 평가
        # ----------------------------------------------------

        score, reasons = score_candidate(
            final_url,
            source_text,
            html,
        )

        # 홈페이지 자체는 게시판으로 저장하지 않음
        if (
            depth > 0
            and score >= 5
            and not is_detail_url(final_url)
        ):

            key = normalize_url(
                final_url
            )

            existing = candidates.get(
                key
            )

            if (
                existing is None
                or score > existing["score"]
            ):

                candidates[key] = {
                    "기관명": org_name,
                    "게시판명": (
                        source_text
                        or "게시판"
                    ),
                    "게시판URL": final_url,
                    "점수": score,
                    "판별근거": ", ".join(
                        reasons
                    ),
                }

        # ----------------------------------------------------
        # 링크 추출
        # ----------------------------------------------------

        links = extract_links(
            final_url,
            html,
        )

        next_links = []

        for item in links:

            link_url = item["url"]
            link_text = item["text"]

            link_score, link_reasons = (
                score_candidate(
                    link_url,
                    link_text,
                )
            )

            # -----------------------------------------------
            # 후보 저장
            # -----------------------------------------------

            # 게시판 가능성이 조금만 있어도 후보로 저장
            if (
                link_score >= 4
                and not is_detail_url(
                    link_url
                )
            ):

                key = normalize_url(
                    link_url
                )

                existing = candidates.get(
                    key
                )

                candidate = {
                    "기관명": org_name,
                    "게시판명": (
                        link_text
                        or "게시판"
                    ),
                    "게시판URL": link_url,
                    "점수": link_score,
                    "판별근거": ", ".join(
                        link_reasons
                    ),
                }

                if (
                    existing is None
                    or link_score > existing["score"]
                ):

                    candidates[key] = candidate

            # -----------------------------------------------
            # 다음 단계 탐색 후보
            # -----------------------------------------------

            # 게시판 상세페이지는 탐색하지 않음
            if is_detail_url(
                link_url
            ):
                continue

            # 이미 방문
            if link_url in visited:
                continue

            # 메뉴명이 의미있는 경우
            menu_match = any(
                word.lower()
                in link_text.lower()
                for word in (
                    BOARD_WORDS
                    + SITE_MAP_WORDS
                )
            )

            # URL이 의미있는 경우
            url_match = any(
                word.lower()
                in link_url.lower()
                for word in BOARD_URL_PATTERNS
            )

            if (
                menu_match
                or url_match
                or depth < 1
            ):

                next_links.append(
                    (
                        link_score,
                        link_url,
                        link_text,
                    )
                )

        # ----------------------------------------------------
        # 다음 탐색 우선순위
        # ----------------------------------------------------

        next_links.sort(
            key=lambda x: x[0],
            reverse=True,
        )

        added = 0

        for (
            link_score,
            link_url,
            link_text,
        ) in next_links:

            if link_url in visited:
                continue

            queue.append(
                (
                    link_url,
                    depth + 1,
                    link_text,
                )
            )

            added += 1

            # 한 페이지에서 너무 많은 메뉴로
            # 확장되지 않도록 제한
            if added >= 20:
                break

    # --------------------------------------------------------
    # 후보 정리
    # --------------------------------------------------------

    candidate_list = list(
        candidates.values()
    )

    # 점수 높은 순
    candidate_list.sort(
        key=lambda x: x["점수"],
        reverse=True,
    )

    # --------------------------------------------------------
    # URL path 기준 중복 제거
    # --------------------------------------------------------

    final = []

    seen_keys = set()

    for item in candidate_list:

        url = item["게시판URL"]

        parsed = urlparse(url)

        key = (
            parsed.netloc.lower(),
            parsed.path.lower(),
            parsed.query.lower(),
        )

        if key in seen_keys:
            continue

        seen_keys.add(key)

        final.append(
            item
        )

        if (
            len(final)
            >= MAX_BOARDS_PER_ORG
        ):
            break

    # --------------------------------------------------------
    # 결과
    # --------------------------------------------------------

    if final:

        return {
            "institution": institution,
            "boards": final,
            "reason": (
                f"{len(final)}개 후보 발견"
            ),
        }

    if homepage_failed:

        return {
            "institution": institution,
            "boards": [],
            "reason": "홈페이지 접속 실패",
        }

    if pages == 0:

        return {
            "institution": institution,
            "boards": [],
            "reason": "HTML 페이지 확보 실패",
        }

    return {
        "institution": institution,
        "boards": [],
        "reason": (
            f"게시판 후보 없음 "
            f"(탐색 {pages}페이지)"
        ),
    }


# ============================================================
# 전체 기관 병렬 탐색
# ============================================================

async def discover_all():
    institutions = load_institutions()

    total = len(
        institutions
    )

    print()
    print("=" * 75)
    print("3차 게시판 탐색 시작")
    print("=" * 75)
    print(
        f"전체 기관 : {total}개"
    )
    print(
        f"동시 탐색 : {CONCURRENCY}개"
    )
    print(
        f"기관당 최대 페이지 : "
        f"{MAX_PAGES_PER_ORG}"
    )
    print("=" * 75)
    print()

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY * 2,
        ssl=False,
    )

    timeout = aiohttp.ClientTimeout(
        total=REQUEST_TIMEOUT
    )

    semaphore = asyncio.Semaphore(
        CONCURRENCY
    )

    async with aiohttp.ClientSession(
        connector=connector,
        timeout=timeout,
        headers=HEADERS,
    ) as session:

        async def worker(
            index,
            institution,
        ):

            async with semaphore:

                result = await discover_one(
                    session,
                    institution,
                )

                boards = result[
                    "boards"
                ]

                print(
                    f"[{index:03d}/{total}] "
                    f"{institution['기관명']} "
                    f"→ "
                    f"{len(boards)}개 "
                    f"({result['reason']})"
                )

                return result

        tasks = [
            asyncio.create_task(
                worker(
                    index,
                    institution,
                )
            )
            for index, institution
            in enumerate(
                institutions,
                start=1,
            )
        ]

        results = await asyncio.gather(
            *tasks
        )

    # ========================================================
    # 결과 정리
    # ========================================================

    board_rows = []
    missing_rows = []

    found_orgs = set()

    for result in results:

        institution = result[
            "institution"
        ]

        boards = result[
            "boards"
        ]

        if boards:

            found_orgs.add(
                institution["기관명"]
            )

            for board in boards:

                board["기관유형"] = (
                    institution[
                        "기관유형"
                    ]
                )

                board["홈페이지"] = (
                    institution["URL"]
                )

                board_rows.append(
                    board
                )

        else:

            missing_rows.append(
                {
                    **institution,
                    "상태": "미발견",
                    "비고": result[
                        "reason"
                    ],
                }
            )

    # 저장
    save_boards(
        board_rows
    )

    save_missing(
        missing_rows
    )

    # ========================================================
    # 통계
    # ========================================================

    found_count = len(
        found_orgs
    )

    missing_count = (
        total - found_count
    )

    coverage = (
        found_count
        / total
        * 100
        if total
        else 0
    )

    print()
    print("=" * 75)
    print("3차 게시판 탐색 완료")
    print("=" * 75)

    print(
        f"전체 기관          : "
        f"{total}개"
    )

    print(
        f"게시판 발견 기관    : "
        f"{found_count}개"
    )

    print(
        f"게시판 미발견 기관  : "
        f"{missing_count}개"
    )

    print(
        f"기관 발견률         : "
        f"{coverage:.1f}%"
    )

    print(
        f"최종 게시판 수      : "
        f"{len(board_rows)}개"
    )

    print()
    print(
        f"생성 파일 : "
        f"{BOARDS_FILE}"
    )

    print(
        f"생성 파일 : "
        f"{MISSING_FILE}"
    )

    print("=" * 75)

    # ========================================================
    # 미발견 기관 요약
    # ========================================================

    if missing_rows:

        print()
        print("=" * 75)
        print(
            "게시판 미발견 기관 목록"
        )
        print("=" * 75)

        for row in missing_rows:

            print(
                f"- {row['기관명']}"
                f" | {row['비고']}"
            )

        print("=" * 75)

    print()
    print(
        "※ 이번 실행은 게시판 탐색만 수행했습니다."
    )
    print(
        "※ Telegram 알림은 발송하지 않았습니다."
    )


# ============================================================
# Excel
# ============================================================

def load_institutions():

    if not os.path.exists(
        URL_FILE
    ):
        raise FileNotFoundError(
            f"{URL_FILE} 파일이 없습니다."
        )

    wb = load_workbook(
        URL_FILE,
        read_only=True,
        data_only=True,
    )

    ws = wb.active

    rows = list(
        ws.iter_rows(
            values_only=True
        )
    )

    wb.close()

    if not rows:
        return []

    headers = [
        str(x).strip()
        if x is not None
        else ""
        for x in rows[0]
    ]

    header_map = {
        header: index
        for index, header
        in enumerate(headers)
    }

    org_idx = header_map.get(
        "기관명"
    )

    url_idx = header_map.get(
        "URL"
    )

    type_idx = header_map.get(
        "기관유형"
    )

    if org_idx is None:
        raise ValueError(
            "기관명 컬럼이 없습니다."
        )

    if url_idx is None:
        raise ValueError(
            "URL 컬럼이 없습니다."
        )

    institutions = []

    for row in rows[1:]:

        if not row:
            continue

        org_name = (
            str(
                row[org_idx]
            ).strip()
            if (
                org_idx < len(row)
                and row[org_idx]
                is not None
            )
            else ""
        )

        homepage = (
            str(
                row[url_idx]
            ).strip()
            if (
                url_idx < len(row)
                and row[url_idx]
                is not None
            )
            else ""
        )

        org_type = ""

        if (
            type_idx is not None
            and type_idx < len(row)
            and row[type_idx]
            is not None
        ):

            org_type = str(
                row[type_idx]
            ).strip()

        if not org_name:
            continue

        institutions.append(
            {
                "기관명": org_name,
                "URL": homepage,
                "기관유형": org_type,
            }
        )

    return institutions


def save_boards(rows):

    wb = Workbook()

    ws = wb.active

    ws.title = "게시판"

    headers = [
        "기관명",
        "기관유형",
        "홈페이지",
        "게시판명",
        "게시판URL",
        "점수",
        "판별근거",
    ]

    ws.append(
        headers
    )

    for row in rows:

        ws.append(
            [
                row.get(
                    "기관명",
                    "",
                ),
                row.get(
                    "기관유형",
                    "",
                ),
                row.get(
                    "홈페이지",
                    "",
                ),
                row.get(
                    "게시판명",
                    "",
                ),
                row.get(
                    "게시판URL",
                    "",
                ),
                row.get(
                    "점수",
                    "",
                ),
                row.get(
                    "판별근거",
                    "",
                ),
            ]
        )

    widths = {
        "A": 30,
        "B": 20,
        "C": 55,
        "D": 35,
        "E": 90,
        "F": 10,
        "G": 60,
    }

    for col, width in widths.items():

        ws.column_dimensions[
            col
        ].width = width

    ws.freeze_panes = "A2"

    wb.save(
        BOARDS_FILE
    )


def save_missing(rows):

    wb = Workbook()

    ws = wb.active

    ws.title = "미발견"

    headers = [
        "기관명",
        "기관유형",
        "URL",
        "상태",
        "비고",
    ]

    ws.append(
        headers
    )

    for row in rows:

        ws.append(
            [
                row.get(
                    "기관명",
                    "",
                ),
                row.get(
                    "기관유형",
                    "",
                ),
                row.get(
                    "URL",
                    "",
                ),
                row.get(
                    "상태",
                    "",
                ),
                row.get(
                    "비고",
                    "",
                ),
            ]
        )

    widths = {
        "A": 30,
        "B": 20,
        "C": 60,
        "D": 15,
        "E": 60,
    }

    for col, width in widths.items():

        ws.column_dimensions[
            col
        ].width = width

    ws.freeze_panes = "A2"

    wb.save(
        MISSING_FILE
    )


# ============================================================
# 게시판 로드
# ============================================================

def load_boards():

    if not os.path.exists(
        BOARDS_FILE
    ):
        return []

    wb = load_workbook(
        BOARDS_FILE,
        read_only=True,
        data_only=True,
    )

    ws = wb.active

    rows = list(
        ws.iter_rows(
            values_only=True
        )
    )

    wb.close()

    if not rows:
        return []

    headers = [
        str(x).strip()
        if x is not None
        else ""
        for x in rows[0]
    ]

    result = []

    for row in rows[1:]:

        item = {}

        for index, header in enumerate(
            headers
        ):

            if index < len(row):

                item[header] = (
                    str(
                        row[index]
                    ).strip()
                    if row[index]
                    is not None
                    else ""
                )

            else:

                item[header] = ""

        if item.get(
            "게시판URL"
        ):

            result.append(
                item
            )

    return result


# ============================================================
# 게시물 추출
# ============================================================

def extract_posts(
    board_url,
    html,
):
    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    posts = []
    seen = set()

    # table 우선
    tables = soup.find_all(
        "table"
    )

    for table in tables:

        for a in table.find_all(
            "a",
            href=True,
        ):

            title = clean_text(
                a.get_text(
                    " ",
                    strip=True,
                )
            )

            href = a.get(
                "href",
                "",
            ).strip()

            if not title:
                continue

            if href.lower().startswith(
                (
                    "javascript:",
                    "#",
                    "mailto:",
                )
            ):
                continue

            url = normalize_url(
                urljoin(
                    board_url,
                    href,
                )
            )

            if not is_http(url):
                continue

            if not same_domain(
                board_url,
                url,
            ):
                continue

            if (
                url
                == normalize_url(
                    board_url
                )
            ):
                continue

            if url.lower().endswith(
                (
                    ".pdf",
                    ".hwp",
                    ".hwpx",
                    ".xls",
                    ".xlsx",
                    ".zip",
                    ".jpg",
                    ".png",
                )
            ):
                continue

            key = (
                url,
                title,
            )

            if key in seen:
                continue

            seen.add(key)

            posts.append(
                {
                    "title": title,
                    "url": url,
                }
            )

    # 일반 상세 링크
    if len(posts) < 3:

        for a in soup.find_all(
            "a",
            href=True,
        ):

            title = clean_text(
                a.get_text(
                    " ",
                    strip=True,
                )
            )

            href = a.get(
                "href",
                "",
            ).strip()

            if not title:
                continue

            url = normalize_url(
                urljoin(
                    board_url,
                    href,
                )
            )

            if not is_http(url):
                continue

            if not same_domain(
                board_url,
                url,
            ):
                continue

            if not is_detail_url(
                url
            ):
                continue

            key = (
                url,
                title,
            )

            if key in seen:
                continue

            seen.add(key)

            posts.append(
                {
                    "title": title,
                    "url": url,
                }
            )

    return posts[:100]


# ============================================================
# 게시물 상세 내용
# ============================================================

async def fetch_post_content(
    session,
    url,
):
    final_url, html = await fetch(
        session,
        url,
    )

    if not html:
        return ""

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    for tag in soup.find_all(
        [
            "script",
            "style",
            "noscript",
            "header",
            "footer",
            "nav",
        ]
    ):

        tag.decompose()

    return clean_text(
        soup.get_text(
            " ",
            strip=True,
        )
    )


# ============================================================
# Keyword
# ============================================================

def find_keywords(text):

    if not text:
        return []

    lower = text.lower()

    return [
        keyword
        for keyword in KEYWORDS
        if keyword.lower()
        in lower
    ]


# ============================================================
# seen_posts
# ============================================================

def load_seen():

    if not os.path.exists(
        SEEN_FILE
    ):
        return set()

    try:

        with open(
            SEEN_FILE,
            "r",
            encoding="utf-8",
        ) as f:

            data = json.load(
                f
            )

        if isinstance(
            data,
            list,
        ):

            return set(
                data
            )

    except Exception:
        pass

    return set()


def save_seen(seen):

    with open(
        SEEN_FILE,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            sorted(
                seen
            ),
            f,
            ensure_ascii=False,
            indent=2,
        )


# ============================================================
# Telegram
# ============================================================

async def send_telegram(
    session,
    message,
):
    if not TELEGRAM_BOT_TOKEN:
        print(
            "TELEGRAM_BOT_TOKEN 없음"
        )
        return False

    if not TELEGRAM_CHAT_ID:
        print(
            "TELEGRAM_CHAT_ID 없음"
        )
        return False

    url = (
        "https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    )

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "disable_web_page_preview": True,
    }

    try:

        async with session.post(
            url,
            data=payload,
            timeout=20,
        ) as response:

            if response.status == 200:
                return True

            print(
                "Telegram 오류:",
                response.status,
            )

    except Exception as e:

        print(
            "Telegram 오류:",
            e,
        )

    return False


# ============================================================
# 일일 모니터링
# ============================================================

async def monitor():

    boards = load_boards()

    if not boards:

        print(
            "boards.xlsx에 게시판이 없습니다."
        )

        return

    seen = load_seen()

    print()
    print("=" * 75)
    print("일일 게시판 모니터링")
    print("=" * 75)
    print(
        f"게시판 수 : {len(boards)}"
    )
    print(
        f"키워드 : {', '.join(KEYWORDS)}"
    )
    print("=" * 75)

    connector = aiohttp.TCPConnector(
        limit=20,
        ssl=False,
    )

    timeout = aiohttp.ClientTimeout(
        total=REQUEST_TIMEOUT
    )

    new_count = 0

    async with aiohttp.ClientSession(
        connector=connector,
        timeout=timeout,
        headers=HEADERS,
    ) as session:

        for index, board in enumerate(
            boards,
            start=1,
        ):

            org_name = board.get(
                "기관명",
                "",
            )

            board_name = board.get(
                "게시판명",
                "",
            )

            board_url = board.get(
                "게시판URL",
                "",
            )

            print(
                f"[{index}/{len(boards)}] "
                f"{org_name} / "
                f"{board_name}"
            )

            final_url, html = await fetch(
                session,
                board_url,
            )

            if not html:

                print(
                    "   → 접속 실패"
                )

                continue

            posts = extract_posts(
                final_url
                or board_url,
                html,
            )

            for post in posts:

                title = post[
                    "title"
                ]

                post_url = post[
                    "url"
                ]

                title_keywords = (
                    find_keywords(
                        title
                    )
                )

                content = title

                if not title_keywords:

                    content = (
                        await fetch_post_content(
                            session,
                            post_url,
                        )
                    )

                found = find_keywords(
                    content
                )

                if not found:
                    continue

                post_id = post_url

                if post_id in seen:
                    continue

                seen.add(
                    post_id
                )

                new_count += 1

                message = (
                    "🚨 공공기관 "
                    "참여/공모 관련 게시물\n\n"
                    f"기관: {org_name}\n"
                    f"게시판: {board_name}\n"
                    f"제목: {title}\n"
                    f"키워드: "
                    f"{', '.join(found)}\n\n"
                    f"{post_url}"
                )

                print(
                    "   ★ 신규:",
                    title,
                )

                await send_telegram(
                    session,
                    message,
                )

    save_seen(
        seen
    )

    print()
    print("=" * 75)
    print(
        f"모니터링 완료 "
        f"| 신규 {new_count}건"
    )
    print("=" * 75)


# ============================================================
# Main
# ============================================================

async def main():

    # --------------------------------------------------------
    # 강제 재탐색
    # --------------------------------------------------------

    if FORCE_DISCOVER_BOARDS:

        print()
        print(
            "FORCE_DISCOVER_BOARDS=true"
        )

        print(
            "→ boards.xlsx 존재 여부와 관계없이 "
            "게시판을 다시 탐색합니다."
        )

        await discover_all()

        return

    # --------------------------------------------------------
    # boards.xlsx가 없으면 자동 탐색
    # --------------------------------------------------------

    if not os.path.exists(
        BOARDS_FILE
    ):

        print(
            "boards.xlsx가 없습니다."
        )

        print(
            "→ 게시판 탐색을 시작합니다."
        )

        await discover_all()

        return

    # --------------------------------------------------------
    # 정상 운영
    # --------------------------------------------------------

    print(
        "boards.xlsx가 존재합니다."
    )

    print(
        "→ 일일 모니터링을 시작합니다."
    )

    await monitor()


if __name__ == "__main__":
    asyncio.run(
        main()
    )
