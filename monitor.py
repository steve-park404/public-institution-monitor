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

KEYWORDS = [
    "설문조사",
    "시민참여",
    "국민참여",
    "공모전",
]

# 게시판으로 판단할 때 사용하는 단어
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
    "입찰",
    "뉴스",
    "보도자료",
    "자료실",
    "소식지",
    "행사",
    "이벤트",
]

# URL에서 게시판 성격을 판단할 때 사용하는 단어
BOARD_URL_WORDS = [
    "bbs",
    "board",
    "notice",
    "news",
    "list",
    "community",
    "particip",
    "event",
    "survey",
    "data",
    "pds",
    "contents",
    "menu",
]

# 상세 게시물 URL에서 자주 발견되는 패턴
DETAIL_URL_WORDS = [
    "view",
    "detail",
    "read",
    "article",
    "articleNo",
    "nttid",
    "nttId",
    "bbsno",
    "boardno",
    "seq=",
    "idx=",
    "no=",
    "wr_id",
    "uid=",
    "postid",
    "newsid",
]

# 탐색 깊이
MAX_DEPTH = 2

# 기관별 최종 게시판 최대 개수
MAX_BOARDS_PER_ORG = 8

# 한 페이지에서 추출할 링크 최대 개수
MAX_LINKS_PER_PAGE = 150

# HTTP
REQUEST_TIMEOUT = 20

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}


# ============================================================
# 공통 함수
# ============================================================

def normalize_url(url):
    """URL을 비교하기 쉽게 정규화한다."""
    if not url:
        return ""

    url = str(url).strip()

    if not url:
        return ""

    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url

    url = urldefrag(url)[0]

    parsed = urlparse(url)

    # 기본 포트 제거
    netloc = parsed.netloc.lower()

    if netloc.endswith(":80") and parsed.scheme == "http":
        netloc = netloc[:-3]

    if netloc.endswith(":443") and parsed.scheme == "https":
        netloc = netloc[:-4]

    path = parsed.path or "/"

    # 중복 슬래시 제거
    path = re.sub(r"/+", "/", path)

    # query는 대부분 유지
    return parsed._replace(
        netloc=netloc,
        path=path,
    ).geturl()


def same_domain(url1, url2):
    """동일 도메인인지 확인."""
    try:
        a = urlparse(url1).netloc.lower()
        b = urlparse(url2).netloc.lower()

        # www 제거
        a = re.sub(r"^www\.", "", a)
        b = re.sub(r"^www\.", "", b)

        return a == b
    except Exception:
        return False


def is_http_url(url):
    return bool(url and url.lower().startswith(("http://", "https://")))


def clean_text(text):
    if not text:
        return ""

    text = re.sub(r"\s+", " ", text)
    return text.strip()


def is_probably_detail_url(url):
    """게시판이 아니라 개별 게시물일 가능성이 높은 URL인지 판단."""
    lower = url.lower()

    # 명확한 상세 URL 패턴
    for word in DETAIL_URL_WORDS:
        if word.lower() in lower:
            return True

    # query parameter가 지나치게 많은 경우
    try:
        parsed = urlparse(url)
        query = parse_qs(parsed.query)

        detail_keys = [
            "seq",
            "idx",
            "no",
            "nttid",
            "article",
            "articleNo",
            "bbsno",
            "boardno",
            "wr_id",
            "uid",
            "postid",
            "newsid",
        ]

        if any(k.lower() in [x.lower() for x in query.keys()]
               for k in detail_keys):
            return True
    except Exception:
        pass

    return False


def keyword_in_text(text):
    """키워드가 제목/본문에 포함되는지."""
    if not text:
        return []

    found = []

    for keyword in KEYWORDS:
        if keyword.lower() in text.lower():
            found.append(keyword)

    return found


# ============================================================
# HTML 가져오기
# ============================================================

async def fetch(session, url):
    try:
        async with session.get(
            url,
            timeout=aiohttp.ClientTimeout(total=REQUEST_TIMEOUT),
            allow_redirects=True,
            ssl=False,
        ) as response:

            if response.status >= 400:
                return None, None

            content_type = response.headers.get("Content-Type", "")

            if (
                "text/html" not in content_type
                and "application/xhtml" not in content_type
                and content_type
            ):
                return None, None

            text = await response.text(
                errors="ignore"
            )

            final_url = str(response.url)

            return final_url, text

    except Exception:
        return None, None


# ============================================================
# 링크 추출
# ============================================================

def extract_links(base_url, html):
    """HTML에서 내부 링크와 링크 텍스트를 추출."""
    soup = BeautifulSoup(html, "html.parser")

    results = []
    seen = set()

    for tag in soup.find_all("a", href=True):

        href = tag.get("href", "").strip()

        if not href:
            continue

        # javascript / mailto / tel 제외
        if href.lower().startswith(
            ("javascript:", "mailto:", "tel:", "#")
        ):
            continue

        full_url = normalize_url(urljoin(base_url, href))

        if not is_http_url(full_url):
            continue

        if not same_domain(base_url, full_url):
            continue

        # 이미지/파일 등 제외
        lower = full_url.lower()

        excluded_extensions = [
            ".jpg",
            ".jpeg",
            ".png",
            ".gif",
            ".svg",
            ".pdf",
            ".hwp",
            ".hwpx",
            ".xlsx",
            ".xls",
            ".zip",
            ".doc",
            ".docx",
            ".ppt",
            ".pptx",
        ]

        if any(lower.endswith(ext) for ext in excluded_extensions):
            continue

        if full_url in seen:
            continue

        seen.add(full_url)

        text = clean_text(
            tag.get_text(" ", strip=True)
        )

        # 부모 메뉴 텍스트도 일부 활용
        parent_text = ""

        parent = tag.parent

        if parent:
            parent_text = clean_text(
                parent.get_text(" ", strip=True)
            )

        combined_text = clean_text(
            f"{text} {parent_text}"
        )

        results.append(
            {
                "url": full_url,
                "text": combined_text[:300],
            }
        )

        if len(results) >= MAX_LINKS_PER_PAGE:
            break

    return results


# ============================================================
# 게시판 후보 점수 계산
# ============================================================

def score_board_candidate(url, link_text, html=None):
    """
    URL / 메뉴명 / 실제 HTML 구조를 종합해서
    게시판일 가능성을 점수화한다.
    """

    score = 0
    reasons = []

    url_lower = url.lower()
    text_lower = link_text.lower()

    # --------------------------------------------------------
    # 1. 메뉴명
    # --------------------------------------------------------

    for word in BOARD_WORDS:
        if word.lower() in text_lower:
            score += 6
            reasons.append(f"메뉴:{word}")

    # --------------------------------------------------------
    # 2. URL
    # --------------------------------------------------------

    for word in BOARD_URL_WORDS:
        if word.lower() in url_lower:
            score += 3
            reasons.append(f"url:{word}")

    # --------------------------------------------------------
    # 3. 상세 게시물 URL이면 감점
    # --------------------------------------------------------

    if is_probably_detail_url(url):
        score -= 12
        reasons.append("상세URL")

    # --------------------------------------------------------
    # 4. HTML 구조
    # --------------------------------------------------------

    if html:
        soup = BeautifulSoup(html, "html.parser")

        page_text = clean_text(
            soup.get_text(" ", strip=True)
        ).lower()

        # 게시판에서 흔히 나타나는 단어
        board_structure_words = [
            "번호",
            "제목",
            "등록일",
            "조회수",
            "작성일",
            "첨부파일",
            "작성자",
            "목록",
        ]

        structure_hits = 0

        for word in board_structure_words:
            if word.lower() in page_text:
                structure_hits += 1

        if structure_hits >= 2:
            score += 8
            reasons.append("게시판구조")

        if structure_hits >= 4:
            score += 5
            reasons.append("게시판구조강함")

        # table 존재
        tables = soup.find_all("table")

        if tables:
            score += 3
            reasons.append("table")

        # 게시글처럼 보이는 링크 숫자
        detail_links = 0

        for a in soup.find_all("a", href=True):
            href = a.get("href", "").lower()

            if any(
                word.lower() in href
                for word in DETAIL_URL_WORDS
            ):
                detail_links += 1

        if detail_links >= 3:
            score += 5
            reasons.append("게시글링크")

        if detail_links >= 8:
            score += 4
            reasons.append("게시글링크다수")

        # 번호/제목이 실제 표에 같이 있는 경우
        for table in tables:
            table_text = clean_text(
                table.get_text(" ", strip=True)
            ).lower()

            if "번호" in table_text and "제목" in table_text:
                score += 7
                reasons.append("번호+제목표")

                break

    return score, reasons


# ============================================================
# 기관별 게시판 탐색
# ============================================================

async def discover(session, org_name, homepage):
    """
    홈페이지에서 게시판 후보를 탐색한다.

    BFS 방식:
        홈페이지
          ↓
        메뉴/사이트맵
          ↓
        게시판 후보
          ↓
        실제 HTML 검증
    """

    homepage = normalize_url(homepage)

    if not homepage:
        return [], "홈페이지 URL 없음"

    queue = [
        (homepage, 0, "")
    ]

    visited = set()

    candidates = {}

    while queue:

        current_url, depth, source_text = queue.pop(0)

        current_url = normalize_url(current_url)

        if current_url in visited:
            continue

        visited.add(current_url)

        if depth > MAX_DEPTH:
            continue

        final_url, html = await fetch(
            session,
            current_url
        )

        if not html:
            continue

        # ----------------------------------------------------
        # 현재 페이지 자체가 게시판인지 검사
        # ----------------------------------------------------

        score, reasons = score_board_candidate(
            final_url,
            source_text,
            html,
        )

        if score >= 12 and not is_probably_detail_url(final_url):

            key = normalize_url(final_url)

            if key not in candidates or score > candidates[key]["score"]:

                candidates[key] = {
                    "기관명": org_name,
                    "게시판명": source_text or "게시판",
                    "게시판URL": final_url,
                    "점수": score,
                    "판별근거": ", ".join(reasons),
                }

        # ----------------------------------------------------
        # 링크 탐색
        # ----------------------------------------------------

        links = extract_links(
            final_url,
            html,
        )

        # 점수가 높은 링크를 먼저 탐색
        scored_links = []

        for item in links:

            link_url = item["url"]
            link_text = item["text"]

            link_score, link_reasons = score_board_candidate(
                link_url,
                link_text,
            )

            scored_links.append(
                (
                    link_score,
                    link_url,
                    link_text,
                )
            )

            # ------------------------------------------------
            # 게시판 후보 등록
            # ------------------------------------------------

            if (
                link_score >= 8
                and not is_probably_detail_url(link_url)
            ):

                key = normalize_url(link_url)

                if (
                    key not in candidates
                    or link_score > candidates[key]["score"]
                ):

                    candidates[key] = {
                        "기관명": org_name,
                        "게시판명": link_text or "게시판",
                        "게시판URL": link_url,
                        "점수": link_score,
                        "판별근거": ", ".join(link_reasons),
                    }

        # ----------------------------------------------------
        # 다음 탐색 페이지 선정
        # ----------------------------------------------------

        scored_links.sort(
            key=lambda x: x[0],
            reverse=True,
        )

        added = 0

        for link_score, link_url, link_text in scored_links:

            if link_url in visited:
                continue

            # 게시판 후보는 굳이 깊게 들어갈 필요 없음
            if link_score >= 8:
                continue

            # 메뉴/사이트맵 성격의 링크만 추가
            if (
                link_score >= 3
                or any(
                    word.lower() in link_text.lower()
                    for word in BOARD_WORDS
                )
            ):

                queue.append(
                    (
                        link_url,
                        depth + 1,
                        link_text,
                    )
                )

                added += 1

            if added >= 25:
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

    # 같은 기관의 비슷한 URL 제거
    final = []
    seen_paths = set()

    for item in candidate_list:

        url = item["게시판URL"]

        parsed = urlparse(url)

        path_key = (
            parsed.netloc.lower(),
            parsed.path.lower(),
        )

        if path_key in seen_paths:
            continue

        seen_paths.add(path_key)

        final.append(item)

        if len(final) >= MAX_BOARDS_PER_ORG:
            break

    if final:
        return final, "게시판 발견"

    return [], "게시판 후보 미발견"


# ============================================================
# Excel 읽기
# ============================================================

def load_institutions():
    if not os.path.exists(URL_FILE):
        raise FileNotFoundError(
            f"{URL_FILE} 파일을 찾을 수 없습니다."
        )

    wb = load_workbook(
        URL_FILE,
        read_only=True,
        data_only=True,
    )

    ws = wb.active

    rows = list(
        ws.iter_rows(values_only=True)
    )

    if not rows:
        return []

    headers = [
        str(x).strip() if x is not None else ""
        for x in rows[0]
    ]

    header_map = {
        header: idx
        for idx, header in enumerate(headers)
    }

    # 예상 컬럼
    org_idx = header_map.get("기관명")
    url_idx = header_map.get("URL")
    type_idx = header_map.get("기관유형")

    if org_idx is None:
        raise ValueError(
            "url_완성.xlsx에 '기관명' 컬럼이 없습니다."
        )

    if url_idx is None:
        raise ValueError(
            "url_완성.xlsx에 'URL' 컬럼이 없습니다."
        )

    institutions = []

    for row in rows[1:]:

        if not row:
            continue

        org_name = (
            str(row[org_idx]).strip()
            if org_idx < len(row)
            and row[org_idx] is not None
            else ""
        )

        homepage = (
            str(row[url_idx]).strip()
            if url_idx < len(row)
            and row[url_idx] is not None
            else ""
        )

        org_type = ""

        if (
            type_idx is not None
            and type_idx < len(row)
            and row[type_idx] is not None
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

    wb.close()

    return institutions


# ============================================================
# boards.xlsx 저장
# ============================================================

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

    ws.append(headers)

    for row in rows:

        ws.append(
            [
                row.get("기관명", ""),
                row.get("기관유형", ""),
                row.get("홈페이지", ""),
                row.get("게시판명", ""),
                row.get("게시판URL", ""),
                row.get("점수", ""),
                row.get("판별근거", ""),
            ]
        )

    # 열 너비
    widths = {
        "A": 30,
        "B": 20,
        "C": 50,
        "D": 30,
        "E": 80,
        "F": 10,
        "G": 50,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(BOARDS_FILE)


def save_missing(rows):
    wb = Workbook()
    ws = wb.active

    ws.title = "게시판 미발견"

    headers = [
        "기관명",
        "기관유형",
        "홈페이지",
        "상태",
        "비고",
    ]

    ws.append(headers)

    for row in rows:

        ws.append(
            [
                row.get("기관명", ""),
                row.get("기관유형", ""),
                row.get("URL", ""),
                row.get("상태", ""),
                row.get("비고", ""),
            ]
        )

    widths = {
        "A": 30,
        "B": 20,
        "C": 60,
        "D": 20,
        "E": 50,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(MISSING_FILE)


# ============================================================
# 게시판 전체 탐색
# ============================================================

async def discover_all():
    institutions = load_institutions()

    total = len(institutions)

    print("=" * 70)
    print("게시판 전체 탐색 시작")
    print(f"대상 기관 수 : {total}")
    print("=" * 70)

    if total == 0:
        print("처리할 기관이 없습니다.")
        return

    connector = aiohttp.TCPConnector(
        limit=15,
        ssl=False,
    )

    timeout = aiohttp.ClientTimeout(
        total=REQUEST_TIMEOUT
    )

    all_rows = []
    missing_rows = []

    found_orgs = set()

    async with aiohttp.ClientSession(
        connector=connector,
        timeout=timeout,
        headers=HEADERS,
    ) as session:

        for index, institution in enumerate(
            institutions,
            start=1,
        ):

            org_name = institution["기관명"]
            homepage = institution["URL"]
            org_type = institution["기관유형"]

            print(
                f"[{index}/{total}] "
                f"{org_name}"
            )

            if not homepage:
                missing_rows.append(
                    {
                        **institution,
                        "상태": "누락",
                        "비고": "홈페이지 URL 없음",
                    }
                )

                print(
                    "   → 홈페이지 URL 없음"
                )

                continue

            try:

                boards, status = await discover(
                    session,
                    org_name,
                    homepage,
                )

            except Exception as e:

                boards = []

                status = (
                    f"탐색 오류: "
                    f"{type(e).__name__}"
                )

            if boards:

                found_orgs.add(org_name)

                for board in boards:

                    board["기관유형"] = org_type
                    board["홈페이지"] = homepage

                    all_rows.append(board)

                print(
                    f"   → 게시판 {len(boards)}개 발견"
                )

                for board in boards[:5]:
                    print(
                        f"      • "
                        f"{board['게시판명'][:35]} "
                        f""
                        f"[{board['점수']}]"
                    )

            else:

                missing_rows.append(
                    {
                        **institution,
                        "상태": "누락",
                        "비고": status,
                    }
                )

                print(
                    f"   → 게시판 미발견 "
                    f"({status})"
                )

    # --------------------------------------------------------
    # 저장
    # --------------------------------------------------------

    save_boards(all_rows)
    save_missing(missing_rows)

    found_count = len(found_orgs)
    missing_count = total - found_count

    coverage = (
        found_count / total * 100
        if total
        else 0
    )

    print()
    print("=" * 70)
    print("게시판 탐색 완료")
    print("=" * 70)

    print(f"전체 기관       : {total}개")
    print(f"게시판 발견 기관 : {found_count}개")
    print(f"게시판 누락 기관 : {missing_count}개")
    print(f"게시판 발견률    : {coverage:.1f}%")
    print(f"게시판 총 개수   : {len(all_rows)}개")

    print()
    print(f"생성 파일 : {BOARDS_FILE}")
    print(f"누락 파일 : {MISSING_FILE}")

    print()
    print("-" * 70)
    print("게시판 미발견 기관")
    print("-" * 70)

    for row in missing_rows:
        print(
            f"- {row['기관명']} "
            f"| {row['비고']}"
        )

    print("=" * 70)

    # 중요:
    # 게시판 탐색 직후에는 모니터링하지 않는다.
    # 기존 게시물을 새 글로 오인하여 Telegram이 폭주하는 것을 방지.
    print()
    print(
        "게시판 탐색만 완료했습니다."
    )
    print(
        "이번 실행에서는 Telegram 알림을 보내지 않습니다."
    )


# ============================================================
# 게시판 읽기
# ============================================================

def load_boards():
    if not os.path.exists(BOARDS_FILE):
        return []

    wb = load_workbook(
        BOARDS_FILE,
        read_only=True,
        data_only=True,
    )

    ws = wb.active

    rows = list(
        ws.iter_rows(values_only=True)
    )

    wb.close()

    if not rows:
        return []

    headers = [
        str(x).strip() if x is not None else ""
        for x in rows[0]
    ]

    result = []

    for row in rows[1:]:

        item = {}

        for idx, header in enumerate(headers):

            if idx < len(row):
                item[header] = (
                    str(row[idx]).strip()
                    if row[idx] is not None
                    else ""
                )
            else:
                item[header] = ""

        if item.get("게시판URL"):
            result.append(item)

    return result


# ============================================================
# 게시판에서 게시글 추출
# ============================================================

def extract_posts(board_url, html):
    soup = BeautifulSoup(html, "html.parser")

    posts = []

    seen = set()

    # --------------------------------------------------------
    # 방법 1 : table 안의 링크
    # --------------------------------------------------------

    tables = soup.find_all("table")

    for table in tables:

        for a in table.find_all(
            "a",
            href=True,
        ):

            text = clean_text(
                a.get_text(" ", strip=True)
            )

            href = a.get("href", "").strip()

            if not text or len(text) < 2:
                continue

            if href.lower().startswith(
                ("javascript:", "#", "mailto:")
            ):
                continue

            url = normalize_url(
                urljoin(board_url, href)
            )

            if not is_http_url(url):
                continue

            if not same_domain(
                board_url,
                url,
            ):
                continue

            # 목록 자체 링크 제외
            if url == normalize_url(board_url):
                continue

            # 첨부파일 제외
            lower = url.lower()

            if lower.endswith(
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

            key = (url, text)

            if key in seen:
                continue

            seen.add(key)

            posts.append(
                {
                    "title": text,
                    "url": url,
                }
            )

    # --------------------------------------------------------
    # 방법 2 : 일반 링크
    # --------------------------------------------------------

    if len(posts) < 3:

        for a in soup.find_all(
            "a",
            href=True,
        ):

            text = clean_text(
                a.get_text(" ", strip=True)
            )

            href = a.get("href", "").strip()

            if not text or len(text) < 2:
                continue

            url = normalize_url(
                urljoin(board_url, href)
            )

            if not is_http_url(url):
                continue

            if not same_domain(
                board_url,
                url,
            ):
                continue

            if url == normalize_url(board_url):
                continue

            if not is_probably_detail_url(url):
                continue

            key = (url, text)

            if key in seen:
                continue

            seen.add(key)

            posts.append(
                {
                    "title": text,
                    "url": url,
                }
            )

    return posts[:100]


# ============================================================
# 게시글 상세 내용
# ============================================================

async def fetch_post_content(
    session,
    post_url,
):
    final_url, html = await fetch(
        session,
        post_url,
    )

    if not html:
        return ""

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    # 불필요한 영역 제거
    for tag in soup(
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

    text = clean_text(
        soup.get_text(
            " ",
            strip=True,
        )
    )

    return text


# ============================================================
# seen_posts
# ============================================================

def load_seen():
    if not os.path.exists(SEEN_FILE):
        return set()

    try:

        with open(
            SEEN_FILE,
            "r",
            encoding="utf-8",
        ) as f:

            data = json.load(f)

        if isinstance(data, list):
            return set(data)

        return set()

    except Exception:
        return set()


def save_seen(seen):
    with open(
        SEEN_FILE,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            sorted(seen),
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
            "TELEGRAM_BOT_TOKEN이 없습니다."
        )
        return False

    if not TELEGRAM_CHAT_ID:
        print(
            "TELEGRAM_CHAT_ID가 없습니다."
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

            text = await response.text()

            print(
                "Telegram 오류:",
                response.status,
                text[:300],
            )

            return False

    except Exception as e:

        print(
            "Telegram 전송 오류:",
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
            "boards.xlsx에 게시판 정보가 없습니다."
        )
        print(
            "게시판 탐색을 먼저 실행해야 합니다."
        )
        return

    seen = load_seen()

    print("=" * 70)
    print("일일 게시판 모니터링 시작")
    print("=" * 70)
    print(
        f"감시 게시판 수 : {len(boards)}개"
    )
    print(
        f"키워드         : {', '.join(KEYWORDS)}"
    )

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

            if not board_url:
                continue

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
                final_url or board_url,
                html,
            )

            for post in posts:

                title = post["title"]
                post_url = post["url"]

                # ------------------------------------------------
                # 제목에 키워드가 있으면 우선 처리
                # ------------------------------------------------

                title_keywords = keyword_in_text(
                    title
                )

                # ------------------------------------------------
                # 상세 페이지
                # ------------------------------------------------

                content = ""

                if title_keywords:
                    content = title

                else:
                    content = await fetch_post_content(
                        session,
                        post_url,
                    )

                found_keywords = keyword_in_text(
                    content
                )

                if not found_keywords:
                    continue

                # URL을 고유 ID로 사용
                post_id = post_url

                if post_id in seen:
                    continue

                # 새로운 게시물
                seen.add(post_id)

                new_count += 1

                message = (
                    "🚨 공공기관 참여/공모 관련 게시물\n\n"
                    f"기관: {org_name}\n"
                    f"게시판: {board_name}\n"
                    f"제목: {title}\n"
                    f"키워드: {', '.join(found_keywords)}\n\n"
                    f"{post_url}"
                )

                print(
                    "   ★ 신규 발견:",
                    title,
                )

                await send_telegram(
                    session,
                    message,
                )

    save_seen(seen)

    print()
    print("=" * 70)
    print("모니터링 완료")
    print(f"신규 발견 : {new_count}건")
    print(f"누적 확인 : {len(seen)}건")
    print("=" * 70)


# ============================================================
# 최초 실행 / 일일 실행 구분
# ============================================================

async def main():

    # --------------------------------------------------------
    # boards.xlsx가 없으면 게시판 탐색
    # --------------------------------------------------------

    if not os.path.exists(
        BOARDS_FILE
    ):

        print(
            "boards.xlsx가 없습니다."
        )

        print(
            "→ 게시판 전체 탐색을 시작합니다."
        )

        await discover_all()

        return

    # --------------------------------------------------------
    # boards.xlsx가 있으면 일일 모니터링
    # --------------------------------------------------------

    print(
        "boards.xlsx가 존재합니다."
    )

    print(
        "→ 일일 모니터링 모드로 실행합니다."
    )

    await monitor()


if __name__ == "__main__":
    asyncio.run(main())
