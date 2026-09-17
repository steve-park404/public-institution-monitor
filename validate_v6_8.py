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

# 일반적인 상세 URL 패턴
DETAIL_PATTERNS = [
    r"(^|/)(view|detail|read|article|readView|boardView)(/|$)",
    r"([?&])(idx|id|seq|no|articleNo|bbsNo|nttId|boardId|list_no)=",
    r"([?&])act=(view|read|detail)",
]

# 목록 URL에서 흔히 발견되는 구조
LIST_PATTERNS = [
    r"(^|/)(list|boardList|bbsList|list\.do|list\.asp|board\.do|index\.do)(/|$)",
    r"([?&])(page|pageNo|pageIndex|p)=",
]

# 모니터링 목적상 자동확정에서 제외할 가능성이 높은 자료 성격
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
        # fragment 제거
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
        r = session.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
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

    # path/query에 board/bbs/notice/list 등이 명확하면 후보 목록
    if any(w in path for w in ["/board", "/bbs", "/notice", "/news", "/recruit", "/community"]):
        if not any(x in path for x in ["/view", "/detail", "/read"]):
            return "목록"

    return "미분류"


def is_probable_detail(href, text=""):
    u = href.lower()
    for pat in DETAIL_PATTERNS:
        if re.search(pat, u):
            return True

    # 쿼리의 식별자 + 짧은 텍스트는 상세 가능성이 높음
    q = parse_qs(urlparse(u).query)
    ids = {"idx", "id", "seq", "no", "articleNo", "bbsNo", "nttId", "boardId", "list_no"}
    if ids.intersection(q.keys()):
        return True
    return False


def looks_like_post_text(text):
    t = clean_text(text)
    if not t:
        return False
    if len(t) < 2 or len(t) > 300:
        return False
    bad = ["로그인", "회원가입", "검색", "메뉴", "더보기", "이전", "다음", "페이지"]
    if t in bad:
        return False
    return True


def extract_candidate_links(soup, page_url):
    out = []
    if not soup:
        return out

    for a in soup.find_all("a", href=True):
        href = norm_url(a.get("href"), page_url)
        if not href or not same_host(href, page_url):
            continue
        text = clean_text(a.get_text(" ", strip=True))
        title = clean_text(a.get("title"))
        aria = clean_text(a.get("aria-label"))
        label = " ".join(x for x in [text, title, aria] if x)

        # javascript / mailto 등 제거
        if href.lower().startswith(("javascript:", "mailto:", "tel:")):
            continue

        out.append((href, label))
    return out


def extract_post_candidates(soup, page_url):
    """
    게시물 후보 URL을 폭넓게 수집.
    제목 추출은 여기서 확정하지 않는다.
    """
    candidates = []
    if not soup:
        return candidates

    for a in soup.find_all("a", href=True):
        href = norm_url(a.get("href"), page_url)
        if not href or not same_host(href, page_url):
            continue

        text = clean_text(a.get_text(" ", strip=True))
        title_attr = clean_text(a.get("title"))
        aria = clean_text(a.get("aria-label"))
        label = next((x for x in [text, title_attr, aria] if looks_like_post_text(x)), "")

        if not is_probable_detail(href, label):
            continue

        # 너무 명백한 UI 링크 제외
        if label in {"수정", "삭제", "답글", "댓글", "로그인", "목록", "목록으로"}:
            continue

        candidates.append((href, label))

    # URL 자체가 상세 패턴인 링크 중 중복 제거
    seen = set()
    result = []
    for href, label in candidates:
        key = href.rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        result.append((href, label))
        if len(result) >= MAX_DISCOVERY_LINKS:
            break
    return result


def extract_title(soup, response_url=""):
    """
    다중 방식 제목 추출.
    하나가 실패해도 다른 소스를 사용한다.
    """
    if not soup:
        return ""

    selectors = [
        ("meta", {"property": "og:title"}, "content"),
        ("meta", {"name": "twitter:title"}, "content"),
        ("meta", {"name": "title"}, "content"),
    ]
    for tag, attrs, attr in selectors:
        x = soup.find(tag, attrs=attrs)
        if x and x.get(attr):
            t = clean_text(x.get(attr))
            if len(t) >= 2:
                return t

    for sel in ["h1", "h2", "h3", ".subject", ".title", ".board-title",
                ".view-title", ".article-title", "[class*='subject']",
                "[class*='title']"]:
        try:
            for x in soup.select(sel)[:10]:
                t = clean_text(x.get_text(" ", strip=True))
                if 2 <= len(t) <= 300:
                    return t
        except Exception:
            pass

    title = soup.title
    if title:
        t = clean_text(title.get_text(" ", strip=True))
        if t:
            return t

    # body 상단에서 긴 메뉴가 아닌 텍스트 후보
    for x in soup.find_all(["td", "div", "p", "span"]):
        t = clean_text(x.get_text(" ", strip=True))
        if 4 <= len(t) <= 200:
            if not any(k in t.lower() for k in ["copyright", "privacy", "로그인"]):
                return t

    return ""


def extract_date(soup):
    if not soup:
        return ""

    text = clean_text(soup.get_text(" ", strip=True))
    patterns = [
        r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2})",
        r"(20\d{2}년\s*\d{1,2}월\s*\d{1,2}일)",
        r"(\d{4}[./-]\d{1,2}[./-]\d{1,2})",
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(1)
    return ""


def extract_body_length(soup):
    if not soup:
        return 0
    # script/style/nav/footer 제거 후 본문 후보 계산
    s = BeautifulSoup(str(soup), "html.parser")
    for tag in s(["script", "style", "noscript", "svg"]):
        tag.decompose()
    text = clean_text(s.get_text(" ", strip=True))
    return len(text)


def list_structure_score(soup, page_url):
    if not soup:
        return {
            "links": [],
            "keywords": 0,
            "rows": 0,
            "pagination": False,
            "score": 0
        }

    links = extract_post_candidates(soup, page_url)
    all_text = clean_text(soup.get_text(" ", strip=True)).lower()

    keyword_hits = sum(1 for w in LIST_WORDS if w.lower() in all_text)

    # 행 구조 추정
    rows = 0
    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if len(cells) >= 2:
            txt = clean_text(tr.get_text(" ", strip=True))
            if len(txt) >= 4:
                rows += 1

    # li 기반 게시판
    if rows < 3:
        for li in soup.find_all("li"):
            txt = clean_text(li.get_text(" ", strip=True))
            if len(txt) >= 8 and any(c.isdigit() for c in txt):
                rows += 1

    pagination = False
    for a in soup.find_all("a", href=True):
        t = clean_text(a.get_text(" ", strip=True)).lower()
        if t in {"1", "2", "3", "다음", "next", "마지막", "last"}:
            pagination = True
            break
        href = a.get("href", "").lower()
        if any(k in href for k in ["page=", "pageindex=", "pageno=", "page_no="]):
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


def discover_list_from_page(soup, page_url):
    """
    홈페이지/상세 페이지에서 같은 호스트의 목록 후보를 발견.
    우선순위는 명시적인 board/notice/list 경로.
    """
    candidates = []
    for href, label in extract_candidate_links(soup, page_url):
        u = href.lower()
        score = 0

        if any(k in u for k in ["/board", "/bbs", "/notice", "/news", "/recruit", "/community"]):
            score += 5
        if any(k in u for k in ["list", "boardlist", "bbslist", "index.do"]):
            score += 4
        if any(k in label.lower() for k in ["공지", "알림", "채용", "입찰", "공고", "보도자료", "게시판", "notice", "board", "news"]):
            score += 4
        if classify_url(href) == "목록":
            score += 5
        if classify_url(href) == "상세":
            score -= 5

        if score > 0:
            candidates.append((score, href, label))

    candidates.sort(key=lambda x: (-x[0], len(x[1])))
    return [(u, label, score) for score, u, label in candidates[:30]]


def verify_post(session, href, label):
    r = fetch(href, session)
    if r is None or r.status_code >= 400:
        return None

    soup = soup_from_response(r)
    if soup is None:
        return None

    title = extract_title(soup, r.url)
    date = extract_date(soup)
    body_len = extract_body_length(soup)

    # 상세 페이지의 제목은 너무 일반적인 사이트 title보다 본문 제목이 우선되도록 보정
    if title:
        title = re.sub(r"\s*[\|\-–—]\s*(홈페이지|한국.*|공공기관.*)$", "", title).strip()

    # 실제 게시물 판단:
    # URL이 상세형이고 제목 또는 날짜 또는 충분한 본문이 있으면 인정
    detail_like = is_probable_detail(r.url, label) or classify_url(r.url) == "상세"
    actual = detail_like and bool(title) and (bool(date) or body_len >= 100)

    return {
        "url": r.url,
        "title": title,
        "date": date,
        "body_len": body_len,
        "actual": actual,
    }


def validate_row(row):
    institution = clean_text(row.get("기관명", ""))
    candidate = clean_text(row.get("V6.7후보URL", ""))

    # V6.7 컬럼명이 다를 경우 대비
    if not candidate:
        for c in ["V6.7최종목록URL", "V6.7후보URL", "V6.7 URL", "URL"]:
            if c in row and clean_text(row.get(c)):
                candidate = clean_text(row.get(c))
                break

    result = {
        "V6.8후보URL": candidate,
        "V6.8후보유형": classify_url(candidate),
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

    if not candidate or not candidate.startswith(("http://", "https://")):
        result["V6.8사유"] = "유효한 후보 URL 없음"
        return result

    session = requests.Session()

    try:
        r = fetch(candidate, session)
        if r is None or r.status_code >= 400:
            result["V6.8사유"] = "후보 URL 접속 실패"
            return result

        final_url = norm_url(r.url)
        soup = soup_from_response(r)
        if soup is None:
            result["V6.8사유"] = "HTML 파싱 실패"
            return result

        final_type = classify_url(final_url)

        # 1차: 후보 자체가 목록이면 반드시 후보 자체를 우선 검증
        list_url = ""
        list_soup = None

        if final_type == "목록":
            list_url = final_url
            list_soup = soup
        else:
            # 2차: 후보가 상세/홈페이지면 목록 후보 탐색
            discovered = discover_list_from_page(soup, final_url)

            # 후보의 기존 V6.7 목록 URL도 보존하여 검토
            old_list = clean_text(row.get("V6.7최종목록URL", ""))
            if old_list:
                discovered.insert(0, (old_list, "V6.7기존목록URL", 100))

            seen = set()
            for u, label, score in discovered:
                u = norm_url(u)
                if not u or u in seen:
                    continue
                seen.add(u)

                rr = fetch(u, session)
                if rr is None or rr.status_code >= 400:
                    continue
                ss = soup_from_response(rr)
                if ss is None:
                    continue

                typ = classify_url(rr.url)
                st = list_structure_score(ss, rr.url)

                # 목록 후보는 실제 구조가 있는 경우에만 인정
                if typ == "목록" and (
                    len(st["links"]) >= 2 or st["rows"] >= 3 or st["keywords"] >= 2
                ):
                    list_url = norm_url(rr.url)
                    list_soup = ss
                    break

        if not list_url or list_soup is None:
            result["V6.8사유"] = "검증 가능한 목록 URL을 확보하지 못함"
            return result

        st = list_structure_score(list_soup, list_url)
        post_links = st["links"]

        unique_urls = []
        seen = set()
        for href, label in post_links:
            if href.rstrip("/") not in seen:
                seen.add(href.rstrip("/"))
                unique_urls.append((href, label))

        verified = []
        with ThreadPoolExecutor(max_workers=MAX_POST_CHECK) as ex:
            futs = [
                ex.submit(verify_post, session, href, label)
                for href, label in unique_urls[:MAX_POST_CHECK]
            ]
            for f in as_completed(futs):
                try:
                    x = f.result()
                    if x and x["actual"]:
                        verified.append(x)
                except Exception:
                    pass

        # URL 순서 안정화
        verified.sort(key=lambda x: x["url"])

        titles = []
        for x in verified:
            t = clean_text(x["title"])
            if t and t not in titles:
                titles.append(t)

        # 게시판 성격 판정
        page_text = clean_text(list_soup.get_text(" ", strip=True)).lower()
        path_text = (list_url + " " + page_text).lower()

        low_priority_hits = sum(1 for w in LOW_PRIORITY_WORDS if w.lower() in path_text)

        strong_structure = (
            len(unique_urls) >= 5
            and len(verified) >= 3
            and st["keywords"] >= 2
            and st["rows"] >= 3
        )

        very_strong = (
            len(unique_urls) >= 8
            and len(verified) >= 4
            and st["keywords"] >= 3
            and st["rows"] >= 5
        )

        # 제목 파싱 실패는 탈락 사유가 아니라 감점/수동확인 사유
        if very_strong and len(titles) >= 2 and low_priority_hits < 4:
            decision = "자동확정"
            reason = "목록 URL + 다수 게시물 URL + 실제 상세페이지 + 구조 반복 확인"
        elif strong_structure:
            decision = "수동확인"
            if len(titles) < 2:
                reason = "게시판 구조와 실제 게시물은 확인되나 제목 파싱 신뢰도가 낮음"
            elif low_priority_hits >= 4:
                reason = "게시판 구조는 확인되나 자료/사진 등 모니터링 목적 적합성 추가 확인 필요"
            else:
                reason = "게시판 구조와 실제 게시물은 확인되나 자동확정 기준 미충족"
        elif len(unique_urls) >= 3 and len(verified) >= 2:
            decision = "수동확인"
            reason = "실제 게시물은 확인되나 목록 구조 증거가 부족함"
        else:
            decision = "제외"
            reason = "게시판 구조 또는 실제 게시물 증거 부족"

        result.update({
            "V6.8최종목록URL": list_url,
            "V6.8최종유형": classify_url(list_url),
            "V6.8목록접속": "성공",
            "V6.8게시물링크수": len(post_links),
            "V6.8고유게시물링크수": len(unique_urls),
            "V6.8실제게시물수": len(verified),
            "V6.8고유게시물제목수": len(titles),
            "V6.8목록키워드수": st["keywords"],
            "V6.8목록행수": st["rows"],
            "V6.8페이지네이션": "있음" if st["pagination"] else "없음",
            "V6.8구조점수": st["score"],
            "V6.8결과": decision,
            "V6.8사유": reason,
        })

        if verified:
            # 가장 내용이 풍부한 검증 게시물 1건 기록
            best = sorted(verified, key=lambda x: (x["body_len"], len(x["title"])), reverse=True)[0]
            result.update({
                "V6.8검증게시물URL": best["url"],
                "V6.8검증게시물제목": best["title"],
                "V6.8검증게시물날짜": best["date"],
                "V6.8검증본문길이": best["body_len"],
            })

        return result

    except Exception as e:
        result["V6.8결과"] = "오류"
        result["V6.8오류"] = f"{type(e).__name__}: {str(e)[:500]}"
        return result


def main():
    print("==============================================")
    print("V6.8 Deep Board Re-validation")
    print(f"입력파일 : {INPUT_FILE}")
    print(f"출력파일 : {OUTPUT_FILE}")
    print("boards.xlsx는 수정하지 않습니다.")
    print("==============================================")

    df = pd.read_excel(INPUT_FILE, dtype=object).fillna("")

    # 전체 행을 유지하되, 실제 검증 대상은 V6.7에서 후보/수동확인/자동확정으로 남은 행
    target_mask = pd.Series(False, index=df.index)

    if "V6.7결과" in df.columns:
        target_mask = df["V6.7결과"].astype(str).isin(["자동확정", "수동확인"])
    elif "V6.7후보URL" in df.columns:
        target_mask = df["V6.7후보URL"].astype(str).str.startswith(("http://", "https://"))

    # 만약 V6.7 결과 컬럼이 없어도 URL이 있는 행만 검증
    targets = df[target_mask].copy()

    print(f"전체 행       : {len(df)}")
    print(f"정밀 검증 대상 : {len(targets)}")

    if len(targets) == 0:
        print("검증 대상이 없습니다.")
        return

    records = []
    with ThreadPoolExecutor(max_workers=CONCURRENCY) as ex:
        futures = {
            ex.submit(validate_row, row.to_dict()): idx
            for idx, row in targets.iterrows()
        }

        done = 0
        for f in as_completed(futures):
            idx = futures[f]
            done += 1
            try:
                records.append((idx, f.result()))
                rr = f.result()
                print(
                    f"[V6.8] {done}/{len(targets)} | "
                    f"{df.loc[idx, '기관명'] if '기관명' in df.columns else ''} | "
                    f"{rr.get('V6.8결과','')} | "
                    f"점수 {rr.get('V6.8구조점수',0)} | "
                    f"링크 {rr.get('V6.8고유게시물링크수',0)} | "
                    f"실제게시물 {rr.get('V6.8실제게시물수',0)}"
                )
            except Exception as e:
                records.append((idx, {
                    "V6.8결과": "오류",
                    "V6.8오류": f"{type(e).__name__}: {str(e)[:500]}"
                }))

    # V6.8 컬럼을 원본에 추가/갱신
    new_cols = [
        "V6.8후보URL", "V6.8후보유형", "V6.8최종목록URL", "V6.8최종유형",
        "V6.8목록접속", "V6.8게시물링크수", "V6.8고유게시물링크수",
        "V6.8실제게시물수", "V6.8고유게시물제목수", "V6.8목록키워드수",
        "V6.8목록행수", "V6.8페이지네이션", "V6.8구조점수",
        "V6.8검증게시물URL", "V6.8검증게시물제목", "V6.8검증게시물날짜",
        "V6.8검증본문길이", "V6.8결과", "V6.8사유", "V6.8오류"
    ]

    for c in new_cols:
        if c not in df.columns:
            df[c] = ""

    for idx, rr in records:
        for c in new_cols:
            if c in rr:
                df.at[idx, c] = rr[c]

    # 숫자형 명시
    for c in [
        "V6.8게시물링크수", "V6.8고유게시물링크수", "V6.8실제게시물수",
        "V6.8고유게시물제목수", "V6.8목록키워드수", "V6.8목록행수",
        "V6.8구조점수", "V6.8검증본문길이"
    ]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    df.to_excel(OUTPUT_FILE, index=False)

    print("----------------------------------------------")
    print("V6.8 결과")
    print(f"자동확정 : {(df['V6.8결과'] == '자동확정').sum()}")
    print(f"수동확인 : {(df['V6.8결과'] == '수동확인').sum()}")
    print(f"제외     : {(df['V6.8결과'] == '제외').sum()}")
    print(f"오류     : {(df['V6.8결과'] == '오류').sum()}")
    print(f"결과파일 : {OUTPUT_FILE}")
    print("----------------------------------------------")


if __name__ == "__main__":
    main()
