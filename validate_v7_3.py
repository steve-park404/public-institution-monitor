# validate_v7_3.py
# V7.3 - 2등급(소폭 코드 보완 후 자동화 가능) 정밀 자동화 검증
#
# 목적:
# V7.2에서 2등급으로 분류된 URL만 대상으로,
# "목록 -> 실제 게시물 -> 제목/날짜" 추출 성공률을 높이는
# 사이트 공통 보완 로직을 적용한다.
#
# 기존 boards.xlsx / boards_cleaned.xlsx는 수정하지 않는다.
# 입력: boards_v7_3_4grade.xlsx (2등급_보완 시트)
# 출력: boards_v7_3_grade2_validation.xlsx

import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, urlunparse, parse_qsl, urlencode

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

INPUT_FILE = "boards_v7_3_4grade.xlsx"
OUTPUT_FILE = "boards_v7_3_grade2_validation.xlsx"

TIMEOUT = 20
CONCURRENCY = 8
MAX_HTML = 2_000_000
MAX_POSTS = 5

DATE_RE = re.compile(r"(20\d{2}[./-]\s*\d{1,2}[./-]\s*\d{1,2})")
DATE_RE_KR = re.compile(r"(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일")

BAD_TEXT = {
    "로그인", "회원가입", "로그아웃", "검색", "통합검색", "사이트맵",
    "개인정보처리방침", "이용약관", "홈", "home", "더보기", "more",
    "이전", "다음", "처음", "마지막", "목록", "목록보기", "글쓰기",
    "등록", "수정", "삭제", "확인", "취소", "다운로드"
}

BAD_PATH = (
    "login", "logout", "sitemap", "privacy", "member", "mypage",
    "javascript:", "mailto:", "tel:"
)

DETAIL_MARKERS = (
    "mode=view", "act=view", "articleNo=", "article_no=", "article_seq=",
    "articleid=", "articleId=", "boardIdx=", "boardidx=", "bbsId=",
    "bbsid=", "nttId=", "nttid=", "idx=", "linkId=", "linkid=",
    "/detail/", "detail.do", "detailview", "view.do", "view.asp",
    "view.php", "view?","/view/"
)

LIST_MARKERS = (
    "/list", "list.do", "list.asp", "list.php",
    "board", "bbs", "notice", "news", "recruit", "community"
)

COMMON_TITLE_PATTERNS = (
    "공지사항 상세", "새소식 공지사항 상세", "알림마당", "공지사항",
    "새소식", "게시판", "목록", "k-water", "kwater"
)


def text(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def norm_url(url, base=""):
    if not url:
        return ""
    u = urljoin(base, str(url).strip()) if base else str(url).strip()
    p = urlparse(u)
    if p.scheme not in ("http", "https") or not p.netloc:
        return ""
    return urlunparse((p.scheme, p.netloc, p.path, p.params, p.query, ""))


def same_domain(a, b):
    try:
        aa = urlparse(a).netloc.lower().replace("www.", "")
        bb = urlparse(b).netloc.lower().replace("www.", "")
        return aa == bb
    except Exception:
        return False


def session():
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
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 Chrome/131.0 Safari/537.36",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7"
    })
    return s


def fetch(s, url):
    try:
        r = s.get(url, timeout=TIMEOUT, allow_redirects=True, stream=True)
        status = r.status_code
        final = str(r.url)
        ctype = (r.headers.get("content-type") or "").lower()
        if "html" not in ctype and "xhtml" not in ctype:
            r.close()
            return None, status, final, ctype, "HTML 아님"

        chunks, total = [], 0
        for chunk in r.iter_content(16384):
            if chunk:
                chunks.append(chunk)
                total += len(chunk)
                if total >= MAX_HTML:
                    break
        raw = b"".join(chunks)[:MAX_HTML]
        enc = r.encoding or "utf-8"
        r.close()
        html = raw.decode(enc, errors="replace")
        return html, status, final, ctype, ""
    except Exception as e:
        return None, 0, url, "", f"{type(e).__name__}: {str(e)[:180]}"


def is_bad_link(title, href):
    t = text(title).lower()
    h = href.lower()
    if t in {x.lower() for x in BAD_TEXT}:
        return True
    if any(x in h for x in BAD_PATH):
        return True
    if h.startswith(("javascript:", "mailto:", "tel:")):
        return True
    return False


def is_detail(url):
    u = url.lower()
    return any(x.lower() in u for x in DETAIL_MARKERS)


def is_list(url):
    u = url.lower()
    if is_detail(u):
        return False
    return any(x in u for x in LIST_MARKERS)


def anchor_title(a):
    for attr in ("data-title", "title", "aria-label"):
        v = text(a.get(attr, ""))
        if 4 <= len(v) <= 300:
            return v
    return text(a.get_text(" ", strip=True))


def valid_title(t):
    t = text(t)
    if not 5 <= len(t) <= 300:
        return False
    if t.lower() in {x.lower() for x in BAD_TEXT}:
        return False
    if any(t.lower().startswith(x.lower()) for x in COMMON_TITLE_PATTERNS):
        return False
    return True


def normalize_title(t):
    t = text(t)
    t = re.sub(r"\s*[-|:]\s*(K[- ]?water|한국수자원공사|홈페이지).*$", "", t, flags=re.I)
    t = re.sub(r"\s*[-|:]\s*공지사항 상세.*$", "", t, flags=re.I)
    return text(t)


def extract_date(raw):
    raw = text(raw)
    m = DATE_RE.search(raw)
    if m:
        return re.sub(r"\s+", "", m.group(1)).replace(".", "-").replace("/", "-")
    m = DATE_RE_KR.search(raw)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""


def container_candidates(soup, base):
    """반복 구조를 우선하여 게시물 링크 추출"""
    containers = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) >= 3:
            containers.append((3, table, len(rows)))

    for tag in soup.find_all(["ul", "ol"]):
        lis = tag.find_all("li", recursive=False)
        if len(lis) >= 3:
            containers.append((2, tag, len(lis)))

    for tag in soup.find_all(["div", "section", "article"]):
        ident = text(
            " ".join([
                str(tag.get("id", "")),
                " ".join(tag.get("class", []) if isinstance(tag.get("class"), list)
                         else [str(tag.get("class", ""))])
            ])
        ).lower()
        if any(k in ident for k in ("board", "bbs", "notice", "news", "list", "table", "content")):
            links = tag.find_all("a", href=True)
            if len(links) >= 3:
                containers.append((1, tag, len(links)))

    containers.sort(key=lambda x: (x[0], x[2]), reverse=True)

    candidates = []
    seen = set()

    for _, node, _ in containers[:30]:
        for a in node.find_all("a", href=True):
            href = norm_url(a.get("href"), base)
            if not href or not same_domain(href, base):
                continue

            title = anchor_title(a)
            if not valid_title(title):
                continue
            if is_bad_link(title, href):
                continue
            if is_list(href):
                continue

            # 실제 게시물 링크에 흔한 신호
            score = 0
            if is_detail(href):
                score += 5
            if a.find_parent("tr"):
                score += 2
            if len(title) >= 8:
                score += 1

            # href가 단순히 같은 목록 URL이면 제외
            if href.rstrip("/") == base.rstrip("/"):
                continue

            if score < 3:
                continue

            key = href.rstrip("/")
            if key in seen:
                continue
            seen.add(key)
            candidates.append((href, title, score))

    candidates.sort(key=lambda x: x[2], reverse=True)
    return candidates


def fallback_candidates(soup, base):
    """일반 HTML 게시판의 보완 추출"""
    out, seen = [], set()

    for a in soup.find_all("a", href=True):
        href = norm_url(a.get("href"), base)
        if not href or not same_domain(href, base):
            continue
        title = anchor_title(a)
        if not valid_title(title) or is_bad_link(title, href):
            continue
        if not is_detail(href) or is_list(href):
            continue

        # 링크 주변 텍스트에 날짜가 있으면 강한 후보
        parent = a.find_parent(["li", "tr", "div", "article"])
        nearby = text(parent.get_text(" ", strip=True) if parent else "")
        score = 5 + (2 if DATE_RE.search(nearby) or DATE_RE_KR.search(nearby) else 0)
        if len(title) >= 8:
            score += 1

        if href.rstrip("/") not in seen:
            seen.add(href.rstrip("/"))
            out.append((href, title, score))

    out.sort(key=lambda x: x[2], reverse=True)
    return out


def post_candidates(soup, base):
    c = container_candidates(soup, base)
    if len(c) < 5:
        f = fallback_candidates(soup, base)
        seen = {x[0].rstrip("/") for x in c}
        c.extend(x for x in f if x[0].rstrip("/") not in seen)
    return c[:30]


def detail_title(soup, anchor):
    candidates = []

    # 게시물 제목에 우선순위를 주는 선택자
    for sel in [
        ".view-title", ".board-title", ".bbs-title", ".post-title",
        ".article-title", ".subject", ".tit", "article h1", "article h2",
        "main h1", "main h2"
    ]:
        for n in soup.select(sel)[:3]:
            v = normalize_title(n.get_text(" ", strip=True))
            if valid_title(v):
                candidates.append(v)

    # og:title는 사이트 전체 제목일 수 있어 낮은 우선순위
    for n in soup.select("meta[property='og:title']")[:1]:
        v = normalize_title(n.get("content", ""))
        if valid_title(v):
            candidates.append(v)

    # anchor 제목은 마지막 보완
    av = normalize_title(anchor)
    if valid_title(av):
        candidates.append(av)

    # title 태그는 최후 보완
    if soup.title:
        v = normalize_title(soup.title.get_text(" ", strip=True))
        if valid_title(v):
            candidates.append(v)

    seen = set()
    for v in candidates:
        k = v.lower()
        if k not in seen:
            seen.add(k)
            return v
    return ""


def verify(s, url, anchor):
    html, status, final, ctype, err = fetch(s, url)
    if not html:
        return {"ok": False, "title": "", "date": "", "body": 0,
                "final": final, "status": status, "error": err}

    soup = BeautifulSoup(html, "html.parser")
    body = text(soup.get_text(" ", strip=True))
    title = detail_title(soup, anchor)
    date = extract_date(body)

    # 상세페이지 확인: 원래 링크 또는 최종 URL 중 하나라도 상세형
    detail_ok = is_detail(url) or is_detail(final)

    # 실제 상세 페이지에서 게시판 목록이 다시 나타나는 경우는 감점
    repeated = len(post_candidates(soup, final))

    ok = (
        status == 200
        and detail_ok
        and len(body) >= 250
        and valid_title(title)
        and repeated < 10
    )

    return {"ok": ok, "title": title, "date": date, "body": len(body),
            "final": final, "status": status, "error": ""}


def run_one(row):
    s = session()
    inst = text(row["기관명"])
    name = text(row["게시판명"])
    url = text(row["게시판URL"])

    rec = {
        "기관명": inst,
        "게시판명": name,
        "게시판URL": url,
        "V7.3최종URL": "",
        "V7.3접속": "실패",
        "V7.3HTTP": 0,
        "V7.3후보수": 0,
        "V7.3강한후보수": 0,
        "V7.3검증수": 0,
        "V7.3고유제목수": 0,
        "V7.3검증제목": "",
        "V7.3검증날짜": "",
        "V7.3검증본문": "",
        "V7.3판정": "수동확인",
        "V7.3사유": "",
        "V7.3오류": ""
    }

    html, status, final, ctype, err = fetch(s, url)
    rec["V7.3최종URL"] = final
    rec["V7.3HTTP"] = status
    rec["V7.3접속"] = "성공" if html and status == 200 else "실패"

    if not html:
        rec["V7.3판정"] = "수동확인"
        rec["V7.3사유"] = "접속 실패"
        rec["V7.3오류"] = err
        return rec

    soup = BeautifulSoup(html, "html.parser")
    candidates = post_candidates(soup, final)
    strong = [x for x in candidates if x[2] >= 5]

    rec["V7.3후보수"] = len(candidates)
    rec["V7.3강한후보수"] = len(strong)

    checked = []
    for u, t, score in candidates[:MAX_POSTS]:
        info = verify(s, u, t)
        if info["ok"]:
            checked.append(info)

    titles = []
    for x in checked:
        t = normalize_title(x["title"])
        if t and t.lower() not in {z.lower() for z in titles}:
            titles.append(t)

    rec["V7.3검증수"] = len(checked)
    rec["V7.3고유제목수"] = len(titles)
    rec["V7.3검증제목"] = " | ".join(titles[:5])
    rec["V7.3검증날짜"] = " | ".join(x["date"] for x in checked[:5])
    rec["V7.3검증본문"] = " | ".join(str(x["body"]) for x in checked[:5])

    # 최종 확정
    if (
        len(candidates) >= 5
        and len(checked) >= 3
        and len(titles) >= 3
        and all(len(t) >= 5 for t in titles[:3])
    ):
        rec["V7.3판정"] = "확정"
        rec["V7.3사유"] = "실제 게시물 3건 이상 + 서로 다른 제목 3개 이상 확인"
    elif len(candidates) >= 3 and len(checked) >= 1:
        rec["V7.3판정"] = "수동확인"
        rec["V7.3사유"] = "게시판 구조 확인되나 최종 확정 증거 부족"
    else:
        rec["V7.3판정"] = "수동확인"
        rec["V7.3사유"] = "실제 게시물 검증 부족"

    return rec


def main():
    print("=" * 70)
    print("V7.3 - 2등급 153개 정밀 자동화 검증")
    print("=" * 70)

    df = pd.read_excel(INPUT_FILE, sheet_name="2등급_보완", dtype=object)
    print(f"대상: {len(df)}개")
    print(f"기관: {df['기관명'].nunique()}개")

    results = []
    start = time.time()

    with ThreadPoolExecutor(max_workers=CONCURRENCY) as ex:
        futures = {ex.submit(run_one, row.to_dict()): i
                   for i, (_, row) in enumerate(df.iterrows())}
        done = 0
        for f in as_completed(futures):
            done += 1
            try:
                results.append(f.result())
            except Exception as e:
                i = futures[f]
                row = df.iloc[i]
                results.append({
                    "기관명": text(row["기관명"]),
                    "게시판명": text(row["게시판명"]),
                    "게시판URL": text(row["게시판URL"]),
                    "V7.3판정": "오류",
                    "V7.3사유": "실행 예외",
                    "V7.3오류": str(e)[:500]
                })
            if done % 20 == 0 or done == len(df):
                print(f"{done}/{len(df)}")

    # 원래 순서 복원
    order = {text(v): i for i, v in enumerate(df["게시판URL"].tolist())}
    results.sort(key=lambda r: order.get(text(r.get("게시판URL")), 999999))
    result = pd.DataFrame(results)

    summary = (
        result["V7.3판정"].value_counts(dropna=False)
        .rename_axis("판정").reset_index(name="건수")
    )

    confirmed = result[result["V7.3판정"] == "확정"].copy()
    manual = result[result["V7.3판정"] == "수동확인"].copy()
    errors = result[result["V7.3판정"] == "오류"].copy()

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="V7.3결과", index=False)
        summary.to_excel(writer, sheet_name="요약", index=False)
        confirmed.to_excel(writer, sheet_name="확정", index=False)
        manual.to_excel(writer, sheet_name="수동확인", index=False)
        errors.to_excel(writer, sheet_name="오류", index=False)

    elapsed = time.time() - start
    print()
    print("=" * 70)
    print("V7.3 완료")
    print("=" * 70)
    print(summary.to_string(index=False))
    print(f"소요시간: {elapsed/60:.1f}분")
    print(f"결과: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
