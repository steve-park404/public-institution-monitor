#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
V6.8 Deep Board Re-validation
Input : boards_v6_7_validation.xlsx
Output: boards_v6_8_validation.xlsx

목표
- V6.7의 후보를 새로 대량 탐색하지 않고 정밀 재검증
- 후보 URL 자체가 실제 목록이면 우선 유지
- 후보가 상세/홈페이지인 경우에만 목록 URL 복구
- 사진자료실/자료실 등 "게시판 구조는 있으나 모니터링 목적과 맞지 않는 곳"을 별도 판정
- 제목 파싱 실패를 단독 탈락 사유로 사용하지 않음
- 서로 다른 게시물 URL과 상세페이지의 실제 내용으로 게시판 여부 확인
- boards.xlsx는 절대 수정하지 않음

환경변수는 사용하지 않습니다.
"""

import re
import time
import html
import unicodedata
from urllib.parse import urljoin, urlparse, parse_qs, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed


INPUT_FILE = "boards_v6_7_validation.xlsx"
OUTPUT_FILE = "boards_v6_8_validation.xlsx"

TIMEOUT = 35
CONCURRENCY = 6
MAX_POST_CHECK = 6
MAX_DISCOVERY_LINKS = 80

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7",
}

LIST_WORDS = [
    "공지", "알림", "소식", "게시판", "목록", "채용", "입찰", "공고",
    "보도자료", "자료실", "뉴스", "notice", "board", "bbs", "list",
    "recruit", "news", "press", "community", "notification"
]

DETAIL_PATTERNS = [
    r"(^|/)(view|detail|read|article|readView|boardView)(/|$)",
    r"([?&])(idx|id|seq|no|articleNo|bbsNo|nttId|boardId|list_no)=",
    r"([?&])act=(view|read|detail)",
]

LIST_PATTERNS = [
    r"(^|/)(list|boardList|bbsList|list\.do|list\.asp|board\.do|index\.do)(/|$)",
    r"([?&])(page|pageNo|pageIndex|p)=",
]

LOW_PRIORITY_WORDS = [
    "사진", "포토", "photo", "gallery", "갤러리", "영상", "동영상",
    "자료", "archive", "ebook", "간행물", "홍보물", "행사사진"
]


def clean_text(s):
    if s is None:
        return ""
    s = html.unescape(str(s))
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def norm_url(url, base=None):
    if not url:
        return ""

    try:
        if base:
            url = urljoin(base, url)

        p = urlparse(url)

        if p.scheme not in ("http", "https") or not p.netloc:
            return ""

        p = p._replace(fragment="")

        return urlunparse(p)

    except Exception:
        return ""


def same_host(a, b):
    try:
        return urlparse(a).netloc.lower() == urlparse(b).netloc.lower()
    except Exception:
        return False


def fetch(url, session):
    try:
        r = session.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
            allow_redirects=True
        )

        r.encoding = r.apparent_encoding or r.encoding

        return r

    except Exception:
        return None


def soup_from_response(r):
    if r is None:
        return None

    try:
        return BeautifulSoup(r.text, "html.parser")
    except Exception:
        return None


def classify_url(url):
    if not url:
        return "미분류"

    u = url.lower()

    path = urlparse(u).path
    query = urlparse(u).query

    for pat in DETAIL_PATTERNS:
        if re.search(pat, path) or re.search(pat, query):
            return "상세"

    for pat in LIST_PATTERNS:
        if re.search(pat, path) or re.search(pat, query):
            return "목록"

    if any(
        w in path
        for w in [
            "/board",
            "/bbs",
            "/notice",
            "/news",
            "/recruit",
            "/community"
        ]
    ):
        if not any(
            x in path
            for x in [
                "/view",
                "/detail",
                "/read"
            ]
        ):
            return "목록"

    return "미분류"


def is_probable_detail(href, text=""):
    u = href.lower()

    for pat in DETAIL_PATTERNS:
        if re.search(pat, u):
            return True

    q = parse_qs(urlparse(u).query)

    ids = {
        "idx",
        "id",
        "seq",
        "no",
        "articleNo",
        "bbsNo",
        "nttId",
        "boardId",
        "list_no"
    }

    if ids.intersection(q.keys()):
        return True

    return False


def looks_like_post_text(text):
    t = clean_text(text)

    if not t:
        return False

    if len(t) < 2 or len(t) > 300:
        return False

    bad = [
        "로그인",
        "회원가입",
        "검색",
        "메뉴",
        "더보기",
        "이전",
        "다음",
        "페이지"
    ]

    if t in bad:
        return False

    return True


def extract_candidate_links(soup, page_url):
    out = []

    if not soup:
        return out

    for a in soup.find_all("a", href=True):

        href = norm_url(
            a.get("href"),
            page_url
        )

        if not href:
            continue

        if not same_host(href, page_url):
            continue

        if href.lower().startswith(
            ("javascript:", "mailto:", "tel:")
        ):
            continue

        text = clean_text(
            a.get_text(" ", strip=True)
        )

        title = clean_text(
            a.get("title")
        )

        aria = clean_text(
            a.get("aria-label")
        )

        label = " ".join(
            x for x in [text, title, aria]
            if x
        )

        out.append(
            (href, label)
        )

    return out


def extract_post_candidates(soup, page_url):
    candidates = []

    if not soup:
        return candidates

    for a in soup.find_all("a", href=True):

        href = norm_url(
            a.get("href"),
            page_url
        )

        if not href:
            continue

        if not same_host(href, page_url):
            continue

        text = clean_text(
            a.get_text(" ", strip=True)
        )

        title_attr = clean_text(
            a.get("title")
        )

        aria = clean_text(
            a.get("aria-label")
        )

        label = next(
            (
                x
                for x in [
                    text,
                    title_attr,
                    aria
                ]
                if looks_like_post_text(x)
            ),
            ""
        )

        if not is_probable_detail(
            href,
            label
        ):
            continue

        if label in {
            "수정",
            "삭제",
            "답글",
            "댓글",
            "로그인",
            "목록",
            "목록으로"
        }:
            continue

        candidates.append(
            (href, label)
        )

    seen = set()
    result = []

    for href, label in candidates:

        key = href.rstrip("/")

        if key in seen:
            continue

        seen.add(key)

        result.append(
            (href, label)
        )

        if len(result) >= MAX_DISCOVERY_LINKS:
            break

    return result


def extract_title(soup, response_url=""):
    if not soup:
        return ""

    selectors = [
        (
            "meta",
            {"property": "og:title"},
            "content"
        ),
        (
            "meta",
            {"name": "twitter:title"},
            "content"
        ),
        (
            "meta",
            {"name": "title"},
            "content"
        ),
    ]

    for tag, attrs, attr in selectors:

        x = soup.find(
            tag,
            attrs=attrs
        )

        if x and x.get(attr):

            t = clean_text(
                x.get(attr)
            )

            if len(t) >= 2:
                return t

    for sel in [
        "h1",
        "h2",
        "h3",
        ".subject",
        ".title",
        ".board-title",
        ".view-title",
        ".article-title",
        "[class*='subject']",
        "[class*='title']"
    ]:

        try:

            for x in soup.select(sel)[:10]:

                t = clean_text(
                    x.get_text(
                        " ",
                        strip=True
                    )
                )

                if 2 <= len(t) <= 300:
                    return t

        except Exception:
            pass

    title = soup.title

    if title:

        t = clean_text(
            title.get_text(
                " ",
                strip=True
            )
        )

        if t:
            return t

    for x in soup.find_all(
        ["td", "div", "p", "span"]
    ):

        t = clean_text(
            x.get_text(
                " ",
                strip=True
            )
        )

        if 4 <= len(t) <= 200:

            if not any(
                k in t.lower()
                for k in [
                    "copyright",
                    "privacy",
                    "로그인"
                ]
            ):
                return t

    return ""


def extract_date(soup):
    if not soup:
        return ""

    text = clean_text(
        soup.get_text(
            " ",
            strip=True
        )
    )

    patterns = [
        r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2})",
        r"(20\d{2}년\s*\d{1,2}월\s*\d{1,2}일)",
        r"(\d{4}[./-]\d{1,2}[./-]\d{1,2})",
    ]

    for p in patterns:

        m = re.search(
            p,
            text
        )

        if m:
            return m.group(1)

    return ""


def extract_body_length(soup):
    if not soup:
        return 0

    s = BeautifulSoup(
        str(soup),
        "html.parser"
    )

    for tag in s(
        [
            "script",
            "style",
            "noscript",
            "svg"
        ]
    ):
        tag.decompose()

    text = clean_text(
        s.get_text(
            " ",
            strip=True
        )
    )

    return len(text)


def list_structure_score(
    soup,
    page_url
):
    if not soup:

        return {
            "links": [],
            "keywords": 0,
            "rows": 0,
            "pagination": False,
            "score": 0
        }

    links = extract_post_candidates(
        soup,
        page_url
    )

    all_text = clean_text(
        soup.get_text(
            " ",
            strip=True
        )
    ).lower()

    keyword_hits = sum(
        1
        for w in LIST_WORDS
        if w.lower() in all_text
    )

    rows = 0

    for tr in soup.find_all("tr"):

        cells = tr.find_all(
            ["td", "th"]
        )

        if len(cells) >= 2:

            txt = clean_text(
                tr.get_text(
                    " ",
                    strip=True
                )
            )

            if len(txt) >= 4:
                rows += 1

    if rows < 3:

        for li in soup.find_all("li"):

            txt = clean_text(
                li.get_text(
                    " ",
                    strip=True
                )
            )

            if (
                len(txt) >= 8
                and any(
                    c.isdigit()
                    for c in txt
                )
            ):
                rows += 1

    pagination = False

    for a in soup.find_all(
        "a",
        href=True
    ):

        t = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        ).lower()

        if t in {
            "1",
            "2",
            "3",
            "다음",
            "next",
            "마지막",
            "last"
        }:

            pagination = True
            break

        href = a.get(
            "href",
            ""
        ).lower()

        if any(
            k in href
            for k in [
                "page=",
                "pageindex=",
                "pageno=",
                "page_no="
            ]
        ):

            pagination = True
            break

    score = 0

    if len(links) >= 5:
        score += 5

    elif len(links) >= 3:
        score += 3

    elif len(links) >= 1:
        score += 1

    if keyword_hits >= 3:
        score += 4

    elif keyword_hits >= 2:
        score += 3

    elif keyword_hits >= 1:
        score += 1

    if rows >= 5:
        score += 4

    elif rows >= 3:
        score += 2

    if pagination:
        score += 2

    return {
        "links": links,
        "keywords": keyword_hits,
        "rows": rows,
        "pagination": pagination,
        "score": score
    }


def discover_list_from_page(
    soup,
    page_url
):
    candidates = []

    for href, label in extract_candidate_links(
        soup,
        page_url
    ):

        u = href.lower()

        score = 0

        if any(
            k in u
            for k in [
                "/board",
                "/bbs",
                "/notice",
                "/news",
                "/recruit",
                "/community"
            ]
        ):
            score += 5

        if any(
            k in u
            for k in [
                "list",
                "boardlist",
                "bbslist",
                "index.do"
            ]
        ):
            score += 4

        if any(
            k in label.lower()
            for k in [
                "공지",
                "알림",
                "채용",
                "입찰",
                "공고",
                "보도자료",
                "게시판",
                "notice",
                "board",
                "news"
            ]
        ):
            score += 4

        if classify_url(href) == "목록":
            score += 5

        if classify_url(href) == "상세":
            score -= 5

        if score > 0:

            candidates.append(
                (
                    score,
                    href,
                    label
                )
            )

    candidates.sort(
        key=lambda x: (
            -x[0],
            len(x[1])
        )
    )

    return [
        (u, label, score)
        for score, u, label
        in candidates[:30]
    ]


def verify_post(
    session,
    href,
    label
):
    r = fetch(
        href,
        session
    )

    if r is None or r.status_code >= 400:
        return None

    soup = soup_from_response(r)

    if soup is None:
        return None

    title = extract_title(
        soup,
        r.url
    )

    date = extract_date(
        soup
    )

    body_len = extract_body_length(
        soup
    )

    if title:

        title = re.sub(
            r"\s*[\|\-–—]\s*(홈페이지|한국.*|공공기관.*)$",
            "",
            title
        ).strip()

    detail_like = (
        is_probable_detail(
            r.url,
            label
        )
        or classify_url(r.url) == "상세"
    )

    actual = (
        detail_like
        and bool(title)
        and (
            bool(date)
            or body_len >= 100
        )
    )

    return {
        "url": r.url,
        "title": title,
        "date": date,
        "body_len": body_len,
        "actual": actual,
    }


def validate_row(row):

    institution = clean_text(
        row.get(
            "기관명",
            ""
        )
    )

    candidate = clean_text(
        row.get(
            "V6.7후보URL",
            ""
        )
    )

    if not candidate:

        for c in [
            "V6.7최종목록URL",
            "V6.7후보URL",
            "V6.7 URL",
            "URL"
        ]:

            if (
                c in row
                and clean_text(
                    row.get(c)
                )
            ):

                candidate = clean_text(
                    row.get(c)
                )

                break

    result = {

        "V6.8후보URL": candidate,

        "V6.8후보유형":
            classify_url(candidate),

        "V6.8최종목록URL": "",

        "V6.8최종유형": "",

        "V6.8목록접속": "실패",

        "V6.8게시물링크수": 0,

        "V6.8고유게시물링크수": 0,

        "V6.8실제게시물수": 0,

        "V6.8고유게시물제목수": 0,

        "V6.8목록키워드수": 0,

        "V6.8목록행수": 0,

        "V6.8페이지네이션": "없음",

        "V6.8구조점수": 0,

        "V6.8검증게시물URL": "",

        "V6.8검증게시물제목": "",

        "V6.8검증게시물날짜": "",

        "V6.8검증본문길이": "",

        "V6.8결과": "제외",

        "V6.8사유": "",

        "V6.8오류": "",
    }

    if (
        not candidate
        or not candidate.startswith(
            ("http://", "https://")
        )
    ):

        result["V6.8사유"] = (
            "유효한 후보 URL 없음"
        )

        return result

    session = requests.Session()

    try:

        r = fetch(
            candidate,
            session
