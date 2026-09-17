# validate_v7_2.py
# V7.2 2단계 실제 게시판 구조 검증
# 입력: boards_cleaned.xlsx
# 출력: boards_v7_2_validation.xlsx
#
# 핵심 변경
# 1) 게시물 링크를 페이지 전체의 <a>에서 무차별 수집하지 않고,
#    table/list/card 등 반복되는 "게시판 영역" 안에서 우선 추출
# 2) 메뉴/푸터/사이트맵/검색/로그인 등의 링크를 강하게 제외
# 3) list/search/homepage URL을 게시물 상세 URL로 오인하지 않도록 차단
# 4) 목록의 anchor 제목과 실제 상세페이지 제목을 비교
# 5) 공통/사이트 제목을 실제 게시물 제목으로 인정하지 않음
# 6) 1단계는 787개 전체의 구조를 판별하고,
#    2단계는 실제 게시판 후보만 개별 게시물까지 정밀 검증
# 7) 기존 boards.xlsx / boards_cleaned.xlsx는 수정하지 않음

import re
import time
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

warnings.filterwarnings("ignore")

INPUT_FILE = "boards_cleaned.xlsx"
OUTPUT_FILE = "boards_v7_2_validation.xlsx"

TIMEOUT = 20
CONCURRENCY = 8
MAX_HTML = 2_000_000
MAX_POST_CHECK = 5

LIST_WORDS = (
    "공지", "공고", "채용", "입찰", "보도자료", "새소식", "소식",
    "알림", "자료실", "고시", "게시판", "notice", "news", "board",
    "bbs", "recruit", "list"
)

COMMON_PAGE_WORDS = {
    "공지사항", "새소식", "알림마당", "게시판", "목록", "더보기",
    "로그인", "회원가입", "사이트맵", "개인정보처리방침", "홈",
    "home", "notice", "news", "board", "k-water", "kwater",
    "한국수자원공사", "국민소통", "검색", "통합검색"
}

FUNCTION_WORDS = {
    "로그인", "회원가입", "로그아웃", "검색", "통합검색", "사이트맵",
    "개인정보처리방침", "이용약관", "오시는길", "찾아오시는길",
    "홈", "home", "메뉴", "닫기", "열기", "더보기", "more",
    "이전", "다음", "처음", "마지막", "prev", "next",
    "sns", "facebook", "youtube", "instagram", "twitter"
}

BAD_PATH_WORDS = (
    "login", "logout", "sitemap", "privacy", "member", "mypage",
    "search", "javascript", "mailto:", "tel:", "rss", "feed"
)

DETAIL_QUERY_KEYS = (
    "mode=view", "act=view", "articleno=", "article_no=", "article_seq=",
    "articleid=", "boardidx=", "bbidx=", "nttid=", "wr_id=",
    "vno=", "linkid=", "idx="
)

LIST_PATH_WORDS = (
    "list", "board", "bbs", "notice", "news", "recruit", "data", "community"
)

DATE_RE = re.compile(r"(20\d{2}[./-]\s*\d{1,2}[./-]\s*\d{1,2})")


def clean_text(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()


def normalize_url(url, base=None):
    if not url:
        return ""
    if base:
        url = urljoin(base, url)
    p = urlparse(url.strip())
    if p.scheme not in ("http", "https") or not p.netloc:
        return ""
    return urlunparse((p.scheme, p.netloc, p.path, p.params, p.query, ""))


def same_domain(a, b):
    try:
        return urlparse(a).netloc.lower().replace("www.", "") == urlparse(b).netloc.lower().replace("www.", "")
    except Exception:
        return False


def build_session():
    s = requests.Session()
    retry = Retry(
        total=1, connect=1, read=1, backoff_factor=0.3,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"]
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=16, pool_maxsize=16)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/131.0 Safari/537.36"
        ),
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7",
    })
    return s


def fetch(session, url):
    try:
        r = session.get(url, timeout=TIMEOUT, allow_redirects=True, stream=True)
        status = r.status_code
        ctype = (r.headers.get("content-type") or "").lower()
        final_url = str(r.url)

        if "text/html" not in ctype and "application/xhtml" not in ctype:
            r.close()
            return None, status, ctype, final_url, "HTML 아님"

        chunks, total = [], 0
        for chunk in r.iter_content(16384):
            if not chunk:
                continue
            chunks.append(chunk)
            total += len(chunk)
            if total >= MAX_HTML:
                break
        raw = b"".join(chunks)[:MAX_HTML]
        enc = r.encoding or "utf-8"
        r.close()

        try:
            html = raw.decode(enc, errors="replace")
        except Exception:
            html = raw.decode("utf-8", errors="replace")

        return html, status, ctype, final_url, ""
    except Exception as e:
        return None, 0, "", url, f"{type(e).__name__}: {str(e)[:180]}"


def title_from_anchor(a):
    for attr in ("data-title", "title", "aria-label"):
        v = clean_text(a.get(attr, ""))
        if 4 <= len(v) <= 300:
            return v
    v = clean_text(a.get_text(" ", strip=True))
    if 4 <= len(v) <= 300:
        return v
    return ""


def is_function_link(title, href):
    t = title.lower().strip()
    h = href.lower()
    if t in {x.lower() for x in FUNCTION_WORDS}:
        return True
    if any(x in h for x in BAD_PATH_WORDS):
        return True
    return False


def is_list_url(url):
    p = urlparse(url.lower())
    q = p.query.lower()
    path = p.path.lower()

    # 상세형 쿼리는 목록 URL보다 우선적으로 상세로 본다.
    if any(k in q for k in DETAIL_QUERY_KEYS):
        return False

    # 명시적 list 경로
    if re.search(r"(?:^|/)(list|list\.do|list\.asp)(?:$|[/?])", path):
        return True

    # board/bbs + list 계열
    if any(w in path for w in LIST_PATH_WORDS) and not re.search(r"/view(?:\.|/)", path):
        return True

    # 페이지 번호/검색 파라미터를 가진 게시판 URL
    if any(k in q for k in ("page=", "pageindex=", "pageno=", "bbsid=", "boardkey=")):
        return True

    return False


def is_detail_url(url):
    p = urlparse(url.lower())
    q = p.query.lower()
    path = p.path.lower()

    if any(x in path for x in ("/noticeview", "/detail/", "/detailview", "/newsview", "/recruitview")):
        return True
    if any(k in q for k in DETAIL_QUERY_KEYS):
        return True
    if re.search(r"/view(?:\.do|\.asp|\.php|/)(?:$|[^/]*)", path):
        return True
    return False


def looks_like_post_title(title):
    t = clean_text(title)
    tl = t.lower()
    if not (5 <= len(t) <= 300):
        return False
    if tl in {x.lower() for x in COMMON_PAGE_WORDS}:
        return False
    if tl in {x.lower() for x in FUNCTION_WORDS}:
        return False
    # 단순 사이트 공통 title
    if tl.startswith("새소식 공지사항 상세") or tl.startswith("공지사항 상세"):
        return False
    return True


def score_post_link(a, href, anchor_title):
    score = 0
    t = anchor_title.lower()
    h = href.lower()

    if is_detail_url(href):
        score += 5
    if looks_like_post_title(anchor_title):
        score += 2
    if any(w in t for w in ("공지", "공고", "채용", "모집", "안내", "보도", "입찰", "소식")):
        score += 1
    if DATE_RE.search(anchor_title):
        score += 1

    # 링크가 table row 안에 있으면 게시판 구조 신호
    tr = a.find_parent("tr")
    if tr:
        score += 2

    # 목록/검색 페이지 자체 링크는 감점
    if is_list_url(href):
        score -= 6

    return score


def find_board_containers(soup):
    containers = []

    # 표 기반 게시판: 행이 여러 개인 table
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) >= 3:
            containers.append(("table", table, len(rows)))

    # ul/ol 기반 게시판: li가 반복되는 영역
    for tag in soup.find_all(["ul", "ol"]):
        lis = tag.find_all("li", recursive=False)
        if len(lis) >= 3:
            containers.append(("list", tag, len(lis)))

    # 게시판에서 흔히 사용하는 class/id
    for tag in soup.find_all(["div", "section", "article"]):
        ident = clean_text(" ".join([
            str(tag.get("id", "")),
            " ".join(tag.get("class", []) if isinstance(tag.get("class"), list) else [str(tag.get("class", ""))])
        ])).lower()
        if any(w in ident for w in ("board", "bbs", "notice", "news", "list", "table", "content")):
            links = tag.find_all("a", href=True)
            if len(links) >= 3:
                containers.append(("block", tag, len(links)))

    # 큰 컨테이너부터가 아니라 반복성이 높은 순으로 정렬
    containers.sort(key=lambda x: x[2], reverse=True)
    return containers[:20]


def extract_candidates_from_container(container, base_url):
    _, node, _ = container
    out, seen = [], set()

    for a in node.find_all("a", href=True):
        href = normalize_url(a.get("href"), base_url)
        if not href or not same_domain(href, base_url):
            continue

        title = title_from_anchor(a)
        if not looks_like_post_title(title):
            continue
        if is_function_link(title, href):
            continue

        score = score_post_link(a, href, title)

        # 상세 URL 또는 목록 내부 반복행의 강한 증거가 필요
        if score < 5:
            continue
        if is_list_url(href):
            continue

        key = href.rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        out.append((href, title, score))

    out.sort(key=lambda x: x[2], reverse=True)
    return out


def extract_post_candidates(soup, base_url):
    # 1차: 게시판 컨테이너
    candidates = []
    for c in find_board_containers(soup):
        candidates.extend(extract_candidates_from_container(c, base_url))

    # 중복 제거
    unique = []
    seen = set()
    for item in sorted(candidates, key=lambda x: x[2], reverse=True):
        if item[0].rstrip("/") in seen:
            continue
        seen.add(item[0].rstrip("/"))
        unique.append(item)

    # 2차 보완: 컨테이너에서 못 찾았지만 상세 URL + 제목이 명확한 경우만 허용
    if len(unique) < 3:
        for a in soup.find_all("a", href=True):
            href = normalize_url(a.get("href"), base_url)
            if not href or not same_domain(href, base_url):
                continue
            title = title_from_anchor(a)
            if not looks_like_post_title(title) or is_function_link(title, href):
                continue
            if not is_detail_url(href) or is_list_url(href):
                continue

            score = score_post_link(a, href, title)
            if score < 7:
                continue
            if href.rstrip("/") in seen:
                continue
            seen.add(href.rstrip("/"))
            unique.append((href, title, score))

    return unique[:30]


def extract_date(text):
    m = DATE_RE.search(text)
    if not m:
        return ""
    return re.sub(r"\s+", "", m.group(1)).replace(".", "-").replace("/", "-")


def normalize_title(t):
    t = clean_text(t)
    t = re.sub(r"\s*[-|:]\s*(K[- ]?water|한국수자원공사|홈페이지).*$", "", t, flags=re.I)
    t = re.sub(r"\s*[-|:]\s*공지사항 상세.*$", "", t, flags=re.I)
    return clean_text(t)


def extract_detail_title(soup, anchor_title=""):
    selectors = [
        ".view-title", ".board-title", ".bbs-title", ".post-title",
        ".subject", ".article-title", ".tit", "article h1", "article h2",
        "main h1", "main h2", "h1", "h2", "h3",
        "meta[property='og:title']"
    ]

    values = []
    for sel in selectors:
        for node in soup.select(sel)[:3]:
            if node.name == "meta":
                v = clean_text(node.get("content", ""))
            else:
                v = clean_text(node.get_text(" ", strip=True))
            if 4 <= len(v) <= 500:
                nv = normalize_title(v)
                if looks_like_post_title(nv):
                    values.append(nv)

    # anchor 제목도 증거로 보존하되, 공통 제목은 제외
    if looks_like_post_title(anchor_title):
        values.append(normalize_title(anchor_title))

    # 가장 짧은 공통 사이트 제목보다 실제 게시물 제목 후보를 우선
    seen = set()
    for v in values:
        if v and v.lower() not in seen:
            seen.add(v.lower())
            return v
    return ""


def verify_post(session, url, anchor_title):
    html, status, ctype, final_url, err = fetch(session, url)
    if not html:
        return {
            "verified": False, "title": "", "anchor_title": anchor_title,
            "date": "", "body_len": 0, "status": status,
            "final_url": final_url, "error": err
        }

    soup = BeautifulSoup(html, "html.parser")
    body = clean_text(soup.get_text(" ", strip=True))
    title = extract_detail_title(soup, anchor_title)
    date = extract_date(body)

    # 상세페이지로 이동했는지 확인
    moved_to_detail = is_detail_url(final_url) or is_detail_url(url)

    # 목록/검색 페이지가 다시 나온 경우 검증 실패
    repeated_links = len(extract_post_candidates(soup, final_url))
    verified = (
        status == 200
        and len(body) >= 300
        and len(title) >= 5
        and moved_to_detail
        and repeated_links < 8
    )

    return {
        "verified": verified, "title": title, "anchor_title": anchor_title,
        "date": date, "body_len": len(body), "status": status,
        "final_url": final_url, "error": err
    }


def phase1(row):
    session = build_session()
    institution = clean_text(row["기관명"])
    board_name = clean_text(row["게시판명"])
    url = clean_text(row["게시판URL"])

    base = {
        "기관명": institution, "게시판명": board_name, "게시판URL": url,
        "V7.2최종URL": "", "V7.2HTTP": 0, "V7.2접속": "실패",
        "V7.2최종유형": "접속실패", "V7.2목록컨테이너": 0,
        "V7.2목록행수": 0, "V7.2게시물후보수": 0,
        "V7.2강한게시물후보수": 0, "V7.2페이지네이션": "없음",
        "V7.2결과": "제외", "V7.2사유": "", "V7.2오류": ""
    }

    if not url.startswith(("http://", "https://")):
        base["V7.2사유"] = "HTTP/HTTPS URL 아님"
        return base

    html, status, ctype, final_url, err = fetch(session, url)
    base["V7.2HTTP"] = status
    base["V7.2최종URL"] = final_url
    base["V7.2접속"] = "성공" if html and status == 200 else "실패"

    if not html:
        base["V7.2사유"] = err or "HTML 응답 없음"
        base["V7.2오류"] = err
        return base

    soup = BeautifulSoup(html, "html.parser")
    containers = find_board_containers(soup)
    candidates = extract_post_candidates(soup, final_url)

    strong = [x for x in candidates if x[2] >= 7]
    rows = len(soup.find_all("tr"))

    text = clean_text(soup.get_text(" ", strip=True)).lower()
    pag = soup.select("a[rel='next'], .pagination a, .paging a, .paginate a")
    has_pagination = bool(pag) or any(x in text for x in ("다음", "next", "페이지"))

    base["V7.2목록컨테이너"] = len(containers)
    base["V7.2목록행수"] = rows
    base["V7.2게시물후보수"] = len(candidates)
    base["V7.2강한게시물후보수"] = len(strong)
    base["V7.2페이지네이션"] = "있음" if has_pagination else "없음"

    if is_detail_url(final_url) and len(candidates) < 3:
        base["V7.2최종유형"] = "상세"
        base["V7.2결과"] = "제외"
        base["V7.2사유"] = "개별 게시물 상세페이지"
        return base

    # 명백한 목록 후보
    if len(strong) >= 5 and (len(containers) >= 1 or rows >= 3):
        base["V7.2최종유형"] = "목록"
        base["V7.2결과"] = "2단계정밀검증"
        base["V7.2사유"] = "반복 게시판 구조 + 강한 게시물 후보 5개 이상"
    elif len(strong) >= 3 and (len(containers) >= 1 or rows >= 3):
        base["V7.2최종유형"] = "목록"
        base["V7.2결과"] = "2단계정밀검증"
        base["V7.2사유"] = "게시판 구조 및 게시물 후보 확인"
    else:
        base["V7.2최종유형"] = "미확정"
        base["V7.2결과"] = "수동확인"
        base["V7.2사유"] = f"게시판 구조 증거 부족: 후보 {len(candidates)}, 강한후보 {len(strong)}"

    return base


def phase2(rec):
    session = build_session()
    url = rec["V7.2최종URL"] or rec["게시판URL"]

    html, status, ctype, final_url, err = fetch(session, url)
    if not html:
        rec["V7.2결과"] = "수동확인"
        rec["V7.2사유"] = "2단계 재접속 실패"
        rec["V7.2오류"] = err
        return rec

    soup = BeautifulSoup(html, "html.parser")
    candidates = extract_post_candidates(soup, final_url)

    checked = []
    for post_url, anchor_title, score in candidates[:MAX_POST_CHECK]:
        info = verify_post(session, post_url, anchor_title)
        if info["verified"]:
            checked.append(info)

    titles = []
    for x in checked:
        t = normalize_title(x["title"])
        if t and t.lower() not in {z.lower() for z in titles}:
            titles.append(t)

    rec["V7.2게시물후보수"] = len(candidates)
    rec["V7.2강한게시물후보수"] = sum(1 for x in candidates if x[2] >= 7)
    rec["V7.2검증게시물수"] = len(checked)
    rec["V7.2고유제목수"] = len(titles)
    rec["V7.2검증제목"] = " | ".join(titles[:5])
    rec["V7.2검증날짜"] = " | ".join(x["date"] for x in checked[:5])
    rec["V7.2검증본문길이"] = " | ".join(str(x["body_len"]) for x in checked[:5])

    # 최종 확정 기준: 실제 게시물 3건 이상 + 서로 다른 제목 3개 이상
    if (
        is_list_url(final_url)
        and len(candidates) >= 5
        and len(checked) >= 3
        and len(titles) >= 3
    ):
        rec["V7.2최종유형"] = "목록"
        rec["V7.2결과"] = "확정"
        rec["V7.2사유"] = "2단계 실제 게시물 3건 이상 및 서로 다른 제목 3개 이상 확인"
    elif (
        len(candidates) >= 3
        and len(checked) >= 1
        and len(titles) >= 1
    ):
        rec["V7.2최종유형"] = "목록"
        rec["V7.2결과"] = "수동확인"
        rec["V7.2사유"] = "게시판 가능성은 있으나 최종 확정 증거 부족"
    elif is_detail_url(final_url):
        rec["V7.2최종유형"] = "상세"
        rec["V7.2결과"] = "제외"
        rec["V7.2사유"] = "상세페이지"
    else:
        rec["V7.2결과"] = "수동확인"
        rec["V7.2사유"] = "2단계에서도 실제 게시판 증거 부족"

    return rec


def main():
    print("=" * 70)
    print("V7.2 2단계 실제 게시판 구조 검증")
    print("=" * 70)

    df = pd.read_excel(INPUT_FILE, dtype=object)
    required = {"기관명", "게시판명", "게시판URL"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"필수 컬럼 없음: {missing}")

    print(f"입력파일 : {INPUT_FILE}")
    print(f"전체대상 : {len(df)}")
    print()

    # -------------------------
    # 1단계: 전체 구조 판별
    # -------------------------
    phase1_records = []
    start = time.time()

    with ThreadPoolExecutor(max_workers=CONCURRENCY) as ex:
        futures = {ex.submit(phase1, row.to_dict()): i for i, (_, row) in enumerate(df.iterrows())}
        done = 0
        for future in as_completed(futures):
            done += 1
            idx = futures[future]
            try:
                rec = future.result()
            except Exception as e:
                rec = {
                    "기관명": clean_text(df.iloc[idx]["기관명"]),
                    "게시판명": clean_text(df.iloc[idx]["게시판명"]),
                    "게시판URL": clean_text(df.iloc[idx]["게시판URL"]),
                    "V7.2최종URL": "", "V7.2HTTP": 0, "V7.2접속": "오류",
                    "V7.2최종유형": "오류", "V7.2목록컨테이너": 0,
                    "V7.2목록행수": 0, "V7.2게시물후보수": 0,
                    "V7.2강한게시물후보수": 0, "V7.2페이지네이션": "",
                    "V7.2결과": "오류", "V7.2사유": "1단계 예외",
                    "V7.2오류": str(e)[:500]
                }
            phase1_records.append(rec)
            if done % 25 == 0 or done == len(df):
                print(f"[1단계] {done}/{len(df)}")

    # 원본 순서 복원
    order = {clean_text(v): i for i, v in enumerate(df["게시판URL"].tolist())}
    phase1_records.sort(key=lambda r: order.get(clean_text(r["게시판URL"]), 999999))

    # -------------------------
    # 2단계: 정밀검증
    # -------------------------
    targets = [
        r for r in phase1_records
        if r["V7.2결과"] == "2단계정밀검증"
    ]
    print(f"1단계 정밀검증 대상: {len(targets)}")

    phase2_map = {}
    with ThreadPoolExecutor(max_workers=CONCURRENCY) as ex:
        futures = {ex.submit(phase2, dict(r)): i for i, r in enumerate(targets)}
        done = 0
        for future in as_completed(futures):
            done += 1
            try:
                rec = future.result()
            except Exception as e:
                rec = dict(targets[futures[future]])
                rec["V7.2결과"] = "수동확인"
                rec["V7.2사유"] = "2단계 예외"
                rec["V7.2오류"] = str(e)[:500]
            phase2_map[clean_text(rec["게시판URL"])] = rec
            if done % 10 == 0 or done == len(targets):
                print(f"[2단계] {done}/{len(targets)}")

    final_records = []
    for r in phase1_records:
        final_records.append(phase2_map.get(clean_text(r["게시판URL"]), r))

    result_df = pd.DataFrame(final_records)
    vcols = [c for c in result_df.columns if c.startswith("V7.2")]
    result_df = result_df[["기관명", "게시판명", "게시판URL"] + vcols]

    # 시트별 분리
    summary = result_df["V7.2결과"].value_counts(dropna=False).rename_axis("결과").reset_index(name="건수")
    confirmed = result_df[result_df["V7.2결과"] == "확정"].copy()
    manual = result_df[result_df["V7.2결과"] == "수동확인"].copy()
    excluded = result_df[result_df["V7.2결과"] == "제외"].copy()
    errors = result_df[result_df["V7.2결과"] == "오류"].copy()
    phase2_sheet = result_df[result_df["V7.2결과"].isin(["확정", "수동확인"])].copy()

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="V7.2검증결과", index=False)
        summary.to_excel(writer, sheet_name="요약", index=False)
        confirmed.to_excel(writer, sheet_name="확정후보", index=False)
        manual.to_excel(writer, sheet_name="수동확인", index=False)
        excluded.to_excel(writer, sheet_name="제외", index=False)
        errors.to_excel(writer, sheet_name="오류", index=False)
        phase2_sheet.to_excel(writer, sheet_name="확정_수동후보", index=False)

    elapsed = time.time() - start
    print()
    print("=" * 70)
    print("V7.2 완료")
    print("=" * 70)
    print(summary.to_string(index=False))
    print(f"소요시간 : {elapsed/60:.1f}분")
    print(f"2단계 대상 : {len(targets)}")
    print(f"결과파일 : {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
