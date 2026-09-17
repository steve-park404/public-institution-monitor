import re
import time
from urllib.parse import urljoin, urlparse, parse_qs

import pandas as pd
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed


INPUT_FILE = "boards_v6_8_validation.xlsx"
OUTPUT_FILE = "boards_v6_9_validation.xlsx"

TIMEOUT = 30
CONCURRENCY = 4
MAX_POSTS = 8

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}

TARGET_RESULTS = {"수동확인", "자동확정"}


def clean_text(value):
    if value is None:
        return ""
    value = re.sub(r"\s+", " ", str(value))
    return value.strip()


def norm_url(url):
    if not url:
        return ""
    url = str(url).strip()
    if not url:
        return ""
    return url


def same_host(a, b):
    try:
        return urlparse(a).netloc.lower() == urlparse(b).netloc.lower()
    except Exception:
        return False


def fetch(session, url):
    try:
        r = session.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
            allow_redirects=True,
        )
        if r.status_code >= 400:
            return None
        if not r.encoding:
            r.encoding = r.apparent_encoding
        return r
    except Exception:
        return None


def soup_of(response):
    if response is None:
        return None
    try:
        return BeautifulSoup(response.text, "html.parser")
    except Exception:
        return None


def unique_keep_order(items):
    out = []
    seen = set()
    for x in items:
        x = clean_text(x)
        if not x:
            continue
        key = x.lower()
        if key not in seen:
            seen.add(key)
            out.append(x)
    return out


def is_noise_title(title):
    t = clean_text(title).lower()
    if not t:
        return True

    noise = [
        "k-water 한국수자원공사",
        "한국수자원공사",
        "k-water",
        "한국여성정책연구원 korean women's development institute",
        "korean women's development institute",
        "home",
        "홈",
        "공지사항",
        "알림마당",
        "더보기",
        "목록",
        "이전",
        "다음",
        "검색",
        "로그인",
    ]

    if t in noise:
        return True

    if len(t) < 4:
        return True

    return False


def looks_like_date(text):
    t = clean_text(text)
    patterns = [
        r"\b20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}\b",
        r"\b20\d{2}\.\d{1,2}\.\d{1,2}\b",
        r"\b20\d{2}-\d{1,2}-\d{1,2}\b",
    ]
    return any(re.search(p, t) for p in patterns)


def extract_date(soup):
    if soup is None:
        return ""

    # 1. time tags
    for tag in soup.find_all("time"):
        txt = clean_text(tag.get("datetime") or tag.get_text(" ", strip=True))
        if looks_like_date(txt):
            return txt

    # 2. meta date fields
    for meta in soup.find_all("meta"):
        key = clean_text(
            meta.get("property")
            or meta.get("name")
            or meta.get("itemprop")
        ).lower()
        val = clean_text(meta.get("content"))
        if val and any(x in key for x in ["date", "published", "article:published"]):
            if looks_like_date(val):
                return val

    # 3. text search
    text = soup.get_text(" ", strip=True)
    m = re.search(r"(20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2})", text)
    return m.group(1) if m else ""


def title_from_meta(soup):
    if soup is None:
        return []

    values = []

    for meta in soup.find_all("meta"):
        key = clean_text(
            meta.get("property")
            or meta.get("name")
            or meta.get("itemprop")
        ).lower()
        val = clean_text(meta.get("content"))

        if not val:
            continue

        if key in {
            "og:title",
            "twitter:title",
            "title",
            "headline",
            "dc.title",
            "dcterms.title",
        }:
            values.append(val)

    if soup.title:
        values.append(soup.title.get_text(" ", strip=True))

    return unique_keep_order(values)


def title_from_headings(soup):
    if soup is None:
        return []

    values = []

    # h1/h2 are usually strong candidates, but exclude generic page headings.
    for tag in soup.find_all(["h1", "h2", "h3"]):
        txt = clean_text(tag.get_text(" ", strip=True))
        if txt and not is_noise_title(txt):
            values.append(txt)

    return unique_keep_order(values)


def title_from_structured_data(soup):
    if soup is None:
        return []

    values = []

    for script in soup.find_all("script", type=re.compile("ld\\+json", re.I)):
        raw = script.string or script.get_text(" ", strip=True)
        if not raw:
            continue

        for key in ["headline", "name", "title"]:
            for m in re.finditer(
                rf'"{key}"\s*:\s*"([^"]{{4,300}})"',
                raw,
                flags=re.I,
            ):
                values.append(m.group(1))

    return unique_keep_order(values)


def score_title_candidate(title, context=""):
    t = clean_text(title)
    if is_noise_title(t):
        return -100

    score = 0

    if 6 <= len(t) <= 180:
        score += 3
    elif len(t) > 180:
        score -= 1

    if re.search(r"(공지|공고|채용|입찰|모집|안내|등록|알림|연구|보고|회의|행사|정책|자료)", t):
        score += 2

    if looks_like_date(t):
        score -= 1

    if context:
        c = clean_text(context)
        if t in c:
            score += 1

    # Site-name-only strings should be strongly penalized.
    if len(t.split()) <= 2 and any(
        x in t.lower()
        for x in [
            "k-water",
            "한국수자원공사",
            "한국여성정책연구원",
            "women's development institute",
        ]
    ):
        score -= 8

    return score


def extract_detail_title(soup):
    if soup is None:
        return ""

    candidates = []

    # Highest confidence: common detail-page title containers.
    selectors = [
        ".view_title",
        ".board_view_title",
        ".bbs_view_title",
        ".view-tit",
        ".board-title",
        ".bbs-title",
        ".subject",
        ".viewSubject",
        ".view_tit",
        ".article-title",
        ".article_tit",
        ".tit",
        "h1",
        "h2",
    ]

    for selector in selectors:
        for tag in soup.select(selector):
            txt = clean_text(tag.get_text(" ", strip=True))
            if txt:
                candidates.append((score_title_candidate(txt, "detail"), txt, selector))

    # Metadata
    for txt in title_from_meta(soup):
        candidates.append((score_title_candidate(txt, "meta"), txt, "meta"))

    # JSON-LD
    for txt in title_from_structured_data(soup):
        candidates.append((score_title_candidate(txt, "jsonld"), txt, "jsonld"))

    # Headings
    for txt in title_from_headings(soup):
        candidates.append((score_title_candidate(txt, "heading"), txt, "heading"))

    candidates.sort(key=lambda x: (x[0], len(x[1])), reverse=True)

    for score, txt, _source in candidates:
        if score >= 2 and not is_noise_title(txt):
            return txt

    return candidates[0][1] if candidates else ""


def extract_list_links(list_url, soup):
    if soup is None:
        return []

    links = []

    for a in soup.find_all("a", href=True):
        href = clean_text(a.get("href"))
        if not href:
            continue

        absolute = urljoin(list_url, href)

        if not same_host(list_url, absolute):
            continue

        text = clean_text(a.get_text(" ", strip=True))
        if not text:
            # aria-label/title can contain the post subject.
            text = clean_text(a.get("aria-label") or a.get("title"))

        if not text:
            continue

        if is_noise_title(text):
            continue

        # Avoid obvious navigation.
        low = text.lower()
        if low in {"prev", "next", "이전", "다음", "목록", "검색", "확인", "닫기"}:
            continue

        links.append({
            "url": absolute,
            "anchor_title": text,
        })

    # Deduplicate by URL.
    out = []
    seen = set()

    for item in links:
        key = item["url"]
        if key in seen:
            continue
        seen.add(key)
        out.append(item)

    return out


def candidate_post_links(list_url, soup):
    links = extract_list_links(list_url, soup)

    scored = []

    for item in links:
        url = item["url"]
        text = item["anchor_title"]

        score = 0

        # Detail-page URL patterns.
        path = urlparse(url).path.lower()
        query = urlparse(url).query.lower()

        if any(x in path for x in [
            "view", "detail", "article", "read", "content",
            "noticeview", "boardview", "view.do"
        ]):
            score += 4

        if any(x in query for x in [
            "seq=", "idx=", "article", "ntt", "bbsctt",
            "boardid=", "list_no="
        ]):
            score += 3

        if 5 <= len(text) <= 180:
            score += 2

        if looks_like_date(text):
            score -= 1

        if is_noise_title(text):
            score -= 5

        if score >= 3:
            scored.append((score, item))

    scored.sort(key=lambda x: x[0], reverse=True)

    out = []
    seen = set()

    for _, item in scored:
        if item["url"] in seen:
            continue
        seen.add(item["url"])
        out.append(item)

    return out[:MAX_POSTS]


def verify_detail(session, item):
    r = fetch(session, item["url"])

    if r is None:
        return {
            "url": item["url"],
            "anchor_title": item["anchor_title"],
            "title": "",
            "date": "",
            "body_len": 0,
            "title_sources": "",
            "ok": False,
        }

    soup = soup_of(r)
    if soup is None:
        return {
            "url": r.url,
            "anchor_title": item["anchor_title"],
            "title": "",
            "date": "",
            "body_len": 0,
            "title_sources": "",
            "ok": False,
        }

    title = extract_detail_title(soup)

    # Fallback to list anchor text when detail title extraction fails.
    if not title or is_noise_title(title):
        title = item["anchor_title"]

    body_text = clean_text(soup.get_text(" ", strip=True))

    return {
        "url": r.url,
        "anchor_title": item["anchor_title"],
        "title": title,
        "date": extract_date(soup),
        "body_len": len(body_text),
        "title_sources": ",".join(
            ["detail", "anchor"] if title == item["anchor_title"]
            else ["detail"]
        ),
        "ok": bool(title) and len(body_text) >= 100,
    }


def validate_row(row):
    institution = clean_text(row.get("기관명", ""))

    # Prefer V6.8's verified list URL.
    list_url = clean_text(row.get("V6.8최종목록URL", ""))

    if not list_url:
        list_url = clean_text(row.get("V6.8후보URL", ""))

    if not list_url:
        return {
            "V6.9최종목록URL": "",
            "V6.9목록접속": "실패",
            "V6.9게시물링크수": 0,
            "V6.9검증성공수": 0,
            "V6.9고유게시물제목수": 0,
            "V6.9검증게시물URL": "",
            "V6.9검증게시물제목": "",
            "V6.9검증게시물날짜": "",
            "V6.9검증본문길이": "",
            "V6.9제목추출근거": "",
            "V6.9결과": "제외",
            "V6.9사유": "검증할 목록 URL이 없음",
            "V6.9오류": "",
        }

    with requests.Session() as session:
        session.headers.update(HEADERS)

        r = fetch(session, list_url)

        if r is None:
            return {
                "V6.9최종목록URL": list_url,
                "V6.9목록접속": "실패",
                "V6.9게시물링크수": 0,
                "V6.9검증성공수": 0,
                "V6.9고유게시물제목수": 0,
                "V6.9검증게시물URL": "",
                "V6.9검증게시물제목": "",
                "V6.9검증게시물날짜": "",
                "V6.9검증본문길이": "",
                "V6.9제목추출근거": "",
                "V6.9결과": "제외",
                "V6.9사유": "목록 페이지 접속 실패",
                "V6.9오류": "",
            }

        final_url = r.url
        soup = soup_of(r)

        if soup is None:
            return {
                "V6.9최종목록URL": final_url,
                "V6.9목록접속": "실패",
                "V6.9게시물링크수": 0,
                "V6.9검증성공수": 0,
                "V6.9고유게시물제목수": 0,
                "V6.9검증게시물URL": "",
                "V6.9검증게시물제목": "",
                "V6.9검증게시물날짜": "",
                "V6.9검증본문길이": "",
                "V6.9제목추출근거": "",
                "V6.9결과": "제외",
                "V6.9사유": "목록 HTML 파싱 실패",
                "V6.9오류": "",
            }

        posts = candidate_post_links(final_url, soup)

        verified = []

        for item in posts:
            result = verify_detail(session, item)
            if result["ok"]:
                verified.append(result)

        titles = unique_keep_order([
            x["title"] for x in verified
            if x["title"] and not is_noise_title(x["title"])
        ])

        urls = unique_keep_order([x["url"] for x in verified])

        # A title is considered independently credible when it differs from
        # generic site/page titles and is supported by an actual detail page.
        credible_titles = [
            t for t in titles
            if score_title_candidate(t) >= 2
        ]

        if len(credible_titles) >= 3 and len(verified) >= 3:
            result_status = "자동확정"
            reason = (
                f"실제 게시물 {len(verified)}개와 고유 게시물 제목 "
                f"{len(credible_titles)}개를 독립적으로 확인"
            )
        elif len(verified) >= 3 and len(credible_titles) >= 2:
            result_status = "수동확인"
            reason = (
                f"실제 게시물 {len(verified)}개 확인, 제목 {len(credible_titles)}개 "
                f"추출되었으나 추가 확인 필요"
            )
        elif len(verified) >= 1:
            result_status = "수동확인"
            reason = (
                f"실제 게시물 {len(verified)}개 확인되었으나 "
                "제목 추출 신뢰도가 충분하지 않음"
            )
        else:
            result_status = "제외"
            reason = "실제 게시물 상세페이지를 확인하지 못함"

        return {
            "V6.9최종목록URL": final_url,
            "V6.9목록접속": "성공",
            "V6.9게시물링크수": len(posts),
            "V6.9검증성공수": len(verified),
            "V6.9고유게시물제목수": len(credible_titles),
            "V6.9검증게시물URL": "\n".join(urls[:8]),
            "V6.9검증게시물제목": "\n".join(credible_titles[:8]),
            "V6.9검증게시물날짜": "\n".join([
                x["date"] for x in verified[:8]
            ]),
            "V6.9검증본문길이": "\n".join([
                str(x["body_len"]) for x in verified[:8]
            ]),
            "V6.9제목추출근거": "\n".join([
                f'{x["title"]} [{x["title_sources"]}]'
                for x in verified[:8]
            ]),
            "V6.9결과": result_status,
            "V6.9사유": reason,
            "V6.9오류": "",
        }


def main():
    print("=" * 70)
    print("V6.9 Title Extraction Deep Validation")
    print(f"입력파일 : {INPUT_FILE}")
    print(f"출력파일 : {OUTPUT_FILE}")
    print("boards.xlsx는 수정하지 않습니다.")
    print("=" * 70)

    df = pd.read_excel(INPUT_FILE, dtype=object).fillna("")

    if "V6.8결과" not in df.columns:
        raise ValueError("V6.8결과 컬럼이 없습니다.")

    mask = df["V6.8결과"].astype(str).str.strip().isin(TARGET_RESULTS)
    target_indices = list(df.index[mask])

    print(f"전체 행       : {len(df)}")
    print(f"정밀 검증 대상 : {len(target_indices)}")

    results = {}

    with ThreadPoolExecutor(max_workers=CONCURRENCY) as executor:
        future_map = {
            executor.submit(validate_row, df.loc[idx].to_dict()): idx
            for idx in target_indices
        }

        completed = 0

        for future in as_completed(future_map):
            idx = future_map[future]
            completed += 1

            try:
                rr = future.result()
            except Exception as e:
                rr = {
                    "V6.9최종목록URL": "",
                    "V6.9목록접속": "오류",
                    "V6.9게시물링크수": 0,
                    "V6.9검증성공수": 0,
                    "V6.9고유게시물제목수": 0,
                    "V6.9검증게시물URL": "",
                    "V6.9검증게시물제목": "",
                    "V6.9검증게시물날짜": "",
                    "V6.9검증본문길이": "",
                    "V6.9제목추출근거": "",
                    "V6.9결과": "오류",
                    "V6.9사유": "",
                    "V6.9오류": repr(e),
                }

            results[idx] = rr

            print(
                f"[V6.9] {completed}/{len(target_indices)} | "
                f"{df.at[idx, '기관명']} | "
                f"{rr.get('V6.9결과', '')} | "
                f"링크 {rr.get('V6.9게시물링크수', 0)} | "
                f"실제게시물 {rr.get('V6.9검증성공수', 0)} | "
                f"고유제목 {rr.get('V6.9고유게시물제목수', 0)}"
            )

    new_cols = [
        "V6.9최종목록URL",
        "V6.9목록접속",
        "V6.9게시물링크수",
        "V6.9검증성공수",
        "V6.9고유게시물제목수",
        "V6.9검증게시물URL",
        "V6.9검증게시물제목",
        "V6.9검증게시물날짜",
        "V6.9검증본문길이",
        "V6.9제목추출근거",
        "V6.9결과",
        "V6.9사유",
        "V6.9오류",
    ]

    # Avoid pandas StringDtype assignment problems and fragmentation.
    additions = pd.DataFrame(
        "",
        index=df.index,
        columns=new_cols,
        dtype=object,
    )

    for idx, rr in results.items():
        for c in new_cols:
            value = rr.get(c, "")
            if value is None:
                value = ""
            additions.at[idx, c] = str(value)

    # Replace previous V6.9 columns if the workflow is re-run locally.
    df = df.drop(columns=[c for c in new_cols if c in df.columns], errors="ignore")
    df = pd.concat([df, additions], axis=1)

    # Convert count columns to numeric only at the end.
    for c in [
        "V6.9게시물링크수",
        "V6.9검증성공수",
        "V6.9고유게시물제목수",
    ]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    df.to_excel(OUTPUT_FILE, index=False)

    counts = (
        df.loc[target_indices, "V6.9결과"]
        .astype(str)
        .value_counts()
        .to_dict()
    )

    print("-" * 70)
    print("V6.9 결과")
    print(f"자동확정 : {counts.get('자동확정', 0)}")
    print(f"수동확인 : {counts.get('수동확인', 0)}")
    print(f"제외     : {counts.get('제외', 0)}")
    print(f"오류     : {counts.get('오류', 0)}")
    print(f"결과파일 : {OUTPUT_FILE}")
    print("-" * 70)


if __name__ == "__main__":
    main()
