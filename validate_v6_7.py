# -*- coding: utf-8 -*-
"""
V6.7 Final Board URL Discovery & Validation

목적
- boards_v6_6_validation.xlsx의 20개 후보를 대상으로
  실제 모니터링에 사용할 '목록 URL'을 재탐색/검증한다.
- 홈페이지/상세페이지를 최종 목록 URL로 절대 확정하지 않는다.
- 실제 게시물 링크/제목/본문을 확인하고, 목록 구조를 보수적으로 판정한다.
- boards.xlsx는 절대 수정하지 않는다.

출력
- boards_v6_7_validation.xlsx
"""

import re
import time
from urllib.parse import urljoin, urlparse, urlunparse, parse_qsl, urlencode
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from bs4 import BeautifulSoup

INPUT_FILE = "boards_v6_6_validation.xlsx"
OUTPUT_FILE = "boards_v6_7_validation.xlsx"

TIMEOUT = 35
CONCURRENCY = 5
MAX_DISCOVERED_LINKS = 80
MAX_VERIFY_POSTS = 5

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7",
}

LIST_KEYWORDS = [
    "공지", "알림", "소식", "게시판", "목록", "번호", "제목", "등록일",
    "작성일", "조회", "채용", "입찰", "보도자료", "자료실", "공고",
    "notice", "board", "list", "bbs", "news", "recruit", "employment",
]

DETAIL_HINTS = [
    "view", "detail", "read", "article", "select", "contents", "readView",
    "view.do", "detail.do", "articleNo", "nttId", "seq", "idx", "list_no",
    "act=view",
]

HOME_PATHS = {"", "/", "/index", "/index.html", "/index.do", "/main", "/main/"}

BAD_EXT = (
    ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".pdf",
    ".zip", ".hwp", ".hwpx", ".xlsx", ".xls", ".doc", ".docx",
    ".ppt", ".pptx", ".mp4", ".mp3",
)

def clean_url(url):
    if not url:
        return ""
    url = str(url).strip()
    if not url:
        return ""
    p = urlparse(url)
    if not p.scheme:
        return ""
    return urlunparse((p.scheme, p.netloc, p.path or "/", p.params, p.query, ""))

def same_host(a, b):
    try:
        return urlparse(a).netloc.lower() == urlparse(b).netloc.lower()
    except Exception:
        return False

def normalize_text(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def fetch(url, session):
    try:
        r = session.get(url, timeout=TIMEOUT, allow_redirects=True)
        if r.status_code >= 400:
            return None, "", r.url, f"HTTP {r.status_code}"
        r.encoding = r.apparent_encoding or r.encoding
        return r, r.text, r.url, ""
    except Exception as e:
        return None, "", url, f"{type(e).__name__}: {e}"

def classify_url(url):
    if not url:
        return "없음"
    p = urlparse(url)
    path = (p.path or "/").lower()
    q = (p.query or "").lower()

    if path.rstrip("/") in ("", "/index", "/index.html", "/index.do", "/main", "/main/index.do"):
        return "홈페이지"

    detail_score = 0
    for h in DETAIL_HINTS:
        if h in path or h in q:
            detail_score += 1

    list_score = 0
    for k in ["list", "board", "bbs", "notice", "news", "recruit", "edu", "search"]:
        if k in path:
            list_score += 1

    if detail_score >= 2 and list_score == 0:
        return "상세"

    if re.search(r"(list|board|bbs|notice|news|recruit|edu|search)", path, re.I):
        return "목록"

    if re.search(r"(view|detail|read|article|select|contents)", path, re.I):
        return "상세"

    if re.search(r"(act=view|articleNo=|nttId=|list_no=|seq=|idx=)", q, re.I):
        return "상세"

    return "미분류"

def is_probable_post_url(url, base_url):
    if not url or not same_host(url, base_url):
        return False
    low = url.lower()
    if any(low.split("?")[0].endswith(ext) for ext in BAD_EXT):
        return False
    p = urlparse(url)
    pathq = (p.path + "?" + p.query).lower()

    # 명백한 목록/검색/홈페이지는 게시물로 세지 않는다.
    if re.search(r"/(list|index|search|main)(/|\.|$)", p.path.lower()):
        return False
    if re.search(r"(page=|pageindex=|pageNo=|currentPage=|offset=)", p.query.lower()):
        return False

    score = 0
    for h in DETAIL_HINTS:
        if h.lower() in pathq:
            score += 1

    # 숫자형 상세 URL도 후보로 인정
    if re.search(r"(articleNo|nttId|seq|idx|list_no)=\d+", p.query, re.I):
        score += 2
    if re.search(r"/\d{2,}(/|$)", p.path):
        score += 1

    return score >= 1

def extract_title(soup):
    selectors = [
        "meta[property='og:title']",
        "meta[name='twitter:title']",
        "h1", "h2", "h3",
        ".subject", ".title", ".view_title", ".board-title",
        ".bbs-title", ".article-title",
        "title",
    ]
    for sel in selectors:
        el = soup.select_one(sel)
        if not el:
            continue
        text = normalize_text(el.get("content") if el.name == "meta" else el.get_text(" ", strip=True))
        if 3 <= len(text) <= 200:
            return text
    return ""

def extract_body_text(soup):
    candidates = [
        "article", "main", "#content", "#contents", ".content",
        ".contents", ".board_view", ".view", ".article", ".bbs-view",
        ".board-view", ".view-content",
    ]
    best = ""
    for sel in candidates:
        for el in soup.select(sel):
            txt = normalize_text(el.get_text(" ", strip=True))
            if len(txt) > len(best):
                best = txt
    if not best:
        best = normalize_text(soup.get_text(" ", strip=True))
    return best[:20000]

def verify_post(url, base_url, session):
    r, html, final_url, err = fetch(url, session)
    if not r:
        return {
            "url": url, "final_url": final_url, "ok": False,
            "title": "", "date": "", "body_len": 0, "error": err
        }

    soup = BeautifulSoup(html, "html.parser")
    title = extract_title(soup)
    body = extract_body_text(soup)

    date = ""
    date_patterns = [
        r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2})",
        r"(20\d{2}\.\d{1,2}\.\d{1,2})",
    ]
    for pat in date_patterns:
        m = re.search(pat, body)
        if m:
            date = m.group(1)
            break

    # 게시물 페이지는 제목과 일정량의 본문이 있는 경우를 우선 인정
    ok = bool(title) and len(body) >= 80

    return {
        "url": url, "final_url": final_url, "ok": ok,
        "title": title, "date": date, "body_len": len(body),
        "error": "",
    }

def page_structure(url, html):
    soup = BeautifulSoup(html, "html.parser")

    anchors = []
    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        text = normalize_text(a.get_text(" ", strip=True))
        if not href or href.startswith(("javascript:", "mailto:", "tel:", "#")):
            continue
        full = clean_url(urljoin(url, href))
        if not full or not same_host(full, url):
            continue
        if not text:
            continue
        anchors.append((full, text))

    # 중복 제거
    seen = set()
    uniq = []
    for u, t in anchors:
        key = (u, t)
        if key not in seen:
            seen.add(key)
            uniq.append((u, t))

    post_candidates = []
    for u, t in uniq:
        if is_probable_post_url(u, url):
            post_candidates.append((u, t))

    # 표/리스트 반복 구조
    table_rows = 0
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) >= 3:
            table_rows = max(table_rows, len(rows))

    li_rows = 0
    for ul in soup.find_all(["ul", "ol"]):
        lis = ul.find_all("li", recursive=False)
        if len(lis) >= 3:
            li_rows = max(li_rows, len(lis))

    text = normalize_text(soup.get_text(" ", strip=True)).lower()
    keyword_count = sum(1 for k in LIST_KEYWORDS if k.lower() in text)

    pagination = bool(soup.select(
        ".pagination, .paging, .page, .paginate, "
        "[class*='pagination'], [class*='paging'], "
        "a[href*='page='], a[href*='pageNo='], a[href*='pageIndex=']"
    ))

    structural_rows = max(table_rows, li_rows)

    return {
        "anchors": uniq,
        "post_candidates": post_candidates[:MAX_DISCOVERED_LINKS],
        "table_rows": table_rows,
        "li_rows": li_rows,
        "structural_rows": structural_rows,
        "keyword_count": keyword_count,
        "pagination": pagination,
        "anchor_count": len(uniq),
    }

def make_list_variants(url):
    """상세 URL에서 흔한 목록 URL 패턴을 제한적으로 생성한다."""
    out = []
    if not url:
        return out
    p = urlparse(url)
    path = p.path
    q = dict(parse_qsl(p.query, keep_blank_values=True))

    # query에서 view/detail 전용 파라미터 제거
    for k in ["act", "articleNo", "nttId", "seq", "idx", "list_no"]:
        q.pop(k, None)

    if q:
        out.append(urlunparse((p.scheme, p.netloc, path, p.params, urlencode(q), "")))

    # path의 detail/view/read 계열을 list/board로 치환
    replacements = [
        (r"/detail(?:/[^/]*)?/?$", "/list"),
        (r"/view(?:/[^/]*)?/?$", "/list"),
        (r"/read(?:/[^/]*)?/?$", "/list"),
        (r"/article(?:/[^/]*)?/?$", "/list"),
        (r"/boardView\.do$", "/boardList.do"),
        (r"/view\.do$", "/list.do"),
        (r"/detail\.do$", "/list.do"),
    ]
    for pat, repl in replacements:
        np = re.sub(pat, repl, path, flags=re.I)
        if np != path:
            out.append(urlunparse((p.scheme, p.netloc, np, p.params, "", "")))

    # 같은 디렉터리의 list.do
    directory = path.rsplit("/", 1)[0] if "/" in path else ""
    if directory:
        for name in ["list.do", "list", "boardList.do", "board/list", "bbs/list"]:
            out.append(urlunparse((p.scheme, p.netloc, directory + "/" + name, "", "", "")))

    return list(dict.fromkeys(clean_url(x) for x in out if x))

def score_list(url, html, structure):
    typ = classify_url(url)
    score = 0

    if typ == "목록":
        score += 4
    elif typ == "상세":
        score -= 8
    elif typ == "홈페이지":
        score -= 12

    pc = len(structure["post_candidates"])
    score += min(pc, 8)

    if structure["structural_rows"] >= 5:
        score += 4
    elif structure["structural_rows"] >= 3:
        score += 2

    if structure["keyword_count"] >= 5:
        score += 4
    elif structure["keyword_count"] >= 3:
        score += 3
    elif structure["keyword_count"] >= 2:
        score += 2

    if structure["pagination"]:
        score += 2

    return score

def discover_from_page(start_url, session):
    """후보 페이지에서 목록 URL을 탐색한다."""
    candidates = []
    visited = set()

    def add_candidate(url, source, depth):
        url = clean_url(url)
        if not url or url in visited:
            return
        visited.add(url)
        candidates.append((url, source, depth))

    add_candidate(start_url, "원후보", 0)

    # 원페이지의 링크를 먼저 수집
    r, html, final_url, err = fetch(start_url, session)
    if r:
        st = page_structure(final_url, html)
        for u, text in st["anchors"]:
            low = (u + " " + text).lower()
            if any(k.lower() in low for k in [
                "공지", "알림", "게시판", "notice", "board", "bbs",
                "소식", "채용", "공고", "입찰", "보도자료", "자료실",
                "news", "recruit"
            ]):
                add_candidate(u, "메뉴/링크", 1)

        for u in make_list_variants(final_url):
            add_candidate(u, "상세→목록 패턴", 1)

    # 후보 최대 15개만 실제 접속
    tested = []
    for url, source, depth in candidates[:15]:
        rr, hh, fu, ee = fetch(url, session)
        if not rr:
            continue
        st = page_structure(fu, hh)
        sc = score_list(fu, hh, st)
        tested.append({
            "url": fu,
            "source": source,
            "depth": depth,
            "score": sc,
            "type": classify_url(fu),
            "html": hh,
            "structure": st,
        })

    # 목록형 우선, 상세/홈페이지는 최종 후보에서 제외
    valid = [
        x for x in tested
        if x["type"] == "목록"
        and len(x["structure"]["post_candidates"]) >= 2
        and x["structure"]["keyword_count"] >= 2
        and x["structure"]["structural_rows"] >= 3
    ]

    valid.sort(key=lambda x: (
        x["score"],
        len(x["structure"]["post_candidates"]),
        x["structure"]["structural_rows"],
        x["structure"]["keyword_count"],
    ), reverse=True)

    return {
        "tested": tested,
        "best": valid[0] if valid else None,
        "initial_final_url": final_url if r else start_url,
        "initial_error": err,
    }

def validate_one(row):
    institution = str(row.get("기관명", "")).strip()

    # V6.5 최종목록URL → 후보 URL
    start = clean_url(row.get("V6.5최종목록URL", ""))

    # 혹시 V6.5 URL이 비어 있으면 이전 후보 컬럼들을 순서대로 사용
    if not start:
        for col in [
            "V6.5검증_목록URL", "V6.4검증_목록URL",
            "V6.3게시판URL", "후보URL", "URL"
        ]:
            if col in row.index:
                start = clean_url(row.get(col, ""))
                if start:
                    break

    result = {
        "V6.7후보URL": start,
        "V6.7후보유형": classify_url(start),
        "V6.7최종목록URL": "",
        "V6.7최종유형": "",
        "V6.7목록접속": "실패",
        "V6.7게시물링크수": 0,
        "V6.7고유게시물링크수": 0,
        "V6.7실제게시물수": 0,
        "V6.7고유게시물제목수": 0,
        "V6.7목록키워드수": 0,
        "V6.7목록행수": 0,
        "V6.7페이지네이션": "없음",
        "V6.7구조점수": 0,
        "V6.7검증게시물URL": "",
        "V6.7검증게시물제목": "",
        "V6.7검증게시물날짜": "",
        "V6.7검증본문길이": 0,
        "V6.7결과": "제외",
        "V6.7사유": "",
        "V6.7오류": "",
    }

    if not start:
        result["V6.7결과"] = "제외"
        result["V6.7사유"] = "검증할 후보 URL 없음"
        return result

    session = requests.Session()
    session.headers.update(HEADERS)

    try:
        discovery = discover_from_page(start, session)
        best = discovery["best"]

        if not best:
            result["V6.7사유"] = (
                "실제 목록으로 인정할 URL을 찾지 못함 "
                "(목록형 URL + 반복 행 + 게시물 링크 + 목록 의미를 동시에 충족하지 못함)"
            )
            return result

        final_url = clean_url(best["url"])
        st = best["structure"]

        # 안전장치: 홈페이지/상세는 절대 최종 목록 URL로 기록하지 않음
        if classify_url(final_url) != "목록":
            result["V6.7사유"] = "최종 후보가 목록 URL이 아니어서 제외"
            return result

        posts = st["post_candidates"][:MAX_VERIFY_POSTS]

        verified = []
        for post_url, post_text in posts:
            v = verify_post(post_url, final_url, session)
            if v["ok"]:
                verified.append(v)

        titles = []
        unique_verified = []
        seen_titles = set()
        for v in verified:
            title = normalize_text(v["title"])
            if title and title not in seen_titles:
                seen_titles.add(title)
                titles.append(title)
                unique_verified.append(v)

        result.update({
            "V6.7최종목록URL": final_url,
            "V6.7최종유형": "목록",
            "V6.7목록접속": "성공",
            "V6.7게시물링크수": len(st["post_candidates"]),
            "V6.7고유게시물링크수": len(set(u for u, _ in st["post_candidates"])),
            "V6.7실제게시물수": len(verified),
            "V6.7고유게시물제목수": len(titles),
            "V6.7목록키워드수": st["keyword_count"],
            "V6.7목록행수": st["structural_rows"],
            "V6.7페이지네이션": "있음" if st["pagination"] else "없음",
            "V6.7구조점수": best["score"],
        })

        if unique_verified:
            result["V6.7검증게시물URL"] = " | ".join(v["final_url"] for v in unique_verified[:3])
            result["V6.7검증게시물제목"] = " | ".join(v["title"] for v in unique_verified[:3])
            result["V6.7검증게시물날짜"] = " | ".join(v["date"] for v in unique_verified[:3])
            result["V6.7검증본문길이"] = max(v["body_len"] for v in unique_verified)

        # 자동확정은 매우 보수적으로
        # 1) 진짜 목록 URL
        # 2) 게시물 링크 >=5
        # 3) 실제 검증 게시물 >=3
        # 4) 서로 다른 제목 >=3
        # 5) 목록 의미 >=3
        # 6) 반복 행 >=5
        # 7) 홈페이지/상세가 아님
        strong = (
            classify_url(final_url) == "목록"
            and len(set(u for u, _ in st["post_candidates"])) >= 5
            and len(verified) >= 3
            and len(titles) >= 3
            and st["keyword_count"] >= 3
            and st["structural_rows"] >= 5
        )

        if strong:
            result["V6.7결과"] = "자동확정"
            result["V6.7사유"] = (
                "목록 URL + 반복 목록 구조 + 충분한 게시물 링크 + "
                "3개 이상 실제 게시물/고유 제목을 모두 확인"
            )
        else:
            result["V6.7결과"] = "수동확인"
            reasons = []
            if len(st["post_candidates"]) < 5:
                reasons.append("게시물 링크 부족")
            if len(verified) < 3:
                reasons.append("실제 게시물 검증 부족")
            if len(titles) < 3:
                reasons.append("고유 게시물 제목 부족")
            if st["keyword_count"] < 3:
                reasons.append("목록 의미 키워드 부족")
            if st["structural_rows"] < 5:
                reasons.append("반복 행 구조 부족")
            result["V6.7사유"] = ", ".join(reasons)

        return result

    except Exception as e:
        result["V6.7결과"] = "오류"
        result["V6.7오류"] = f"{type(e).__name__}: {e}"
        result["V6.7사유"] = "검증 중 예외 발생"
        return result

def main():
    print("=" * 75)
    print("V6.7 Final Board URL Discovery & Validation")
    print("=" * 75)
    print(f"입력파일 : {INPUT_FILE}")
    print(f"출력파일 : {OUTPUT_FILE}")

    df = pd.read_excel(INPUT_FILE, dtype=object).fillna("")

    print(f"검증 대상 : {len(df)}")
    print("=" * 75)

    results = [None] * len(df)

    with ThreadPoolExecutor(max_workers=CONCURRENCY) as executor:
        futures = {
            executor.submit(validate_one, row): i
            for i, (_, row) in enumerate(df.iterrows())
        }

        done = 0
        for future in as_completed(futures):
            i = futures[future]
            try:
                results[i] = future.result()
            except Exception as e:
                results[i] = {
                    "V6.7결과": "오류",
                    "V6.7오류": f"{type(e).__name__}: {e}",
                    "V6.7사유": "작업 예외",
                }
            done += 1

            inst = str(df.iloc[i].get("기관명", "")).strip()
            rr = results[i]
            print(
                f"[V6.7] {done}/{len(df)} | {inst} | "
                f"{rr.get('V6.7결과','')} | "
                f"점수 {rr.get('V6.7구조점수',0)} | "
                f"링크 {rr.get('V6.7게시물링크수',0)} | "
                f"실제게시물 {rr.get('V6.7실제게시물수',0)}"
            )

    result_df = pd.DataFrame(results).fillna("")

    # 입력 컬럼을 그대로 보존하고 V6.7 컬럼만 추가
    for col in result_df.columns:
        df[col] = result_df[col]

    # URL/텍스트는 문자열로 고정
    for col in result_df.columns:
        df[col] = df[col].astype(str)

    # 숫자형 컬럼은 숫자로 저장
    numeric_cols = [
        "V6.7게시물링크수", "V6.7고유게시물링크수",
        "V6.7실제게시물수", "V6.7고유게시물제목수",
        "V6.7목록키워드수", "V6.7목록행수", "V6.7구조점수",
        "V6.7검증본문길이",
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df.to_excel(OUTPUT_FILE, index=False)

    counts = df["V6.7결과"].value_counts()
    print()
    print("=" * 75)
    print("V6.7 최종 URL 검증 완료")
    print("=" * 75)
    print(f"자동확정 : {int(counts.get('자동확정', 0))}")
    print(f"수동확인 : {int(counts.get('수동확인', 0))}")
    print(f"제외     : {int(counts.get('제외', 0))}")
    print(f"오류     : {int(counts.get('오류', 0))}")
    print(f"결과파일 : {OUTPUT_FILE}")
    print("=" * 75)

if __name__ == "__main__":
    main()
