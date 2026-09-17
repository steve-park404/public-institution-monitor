# validate_v7.py
# V7 실제 웹페이지 구조 검증
# 입력: boards_cleaned.xlsx
# 출력: boards_v7_validation.xlsx
#
# 목적:
# 1. 실제 URL 접속 여부 확인
# 2. 홈페이지/상세/목록 구조 판별
# 3. 목록 페이지에서 실제 게시물 링크 탐지
# 4. 게시물 제목/날짜/본문 일부 확인
# 5. 기존 boards.xlsx 및 boards_cleaned.xlsx는 수정하지 않음

import re
import time
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, parse_qsl, urlencode, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

warnings.filterwarnings("ignore")

INPUT_FILE = "boards_cleaned.xlsx"
OUTPUT_FILE = "boards_v7_validation.xlsx"

TIMEOUT = 25
CONCURRENCY = 8
MAX_POST_CHECK = 5
MAX_HTML = 2_500_000

LIST_KEYWORDS = [
    "공지사항", "알림", "소식", "새소식", "게시판", "목록",
    "채용", "입찰", "보도자료", "자료실", "고시", "공고",
    "notice", "news", "board", "bbs", "list", "recruit"
]

DETAIL_SIGNALS = [
    "mode=view", "act=view", "article", "articleNo", "article_seq",
    "boardIdx", "bbIdx", "nttId", "wr_id", "list_no", "view.do",
    "detailView.do", "noticeView", "/detail/"
]

NON_POST_TEXT = {
    "로그인", "회원가입", "메뉴", "닫기", "검색", "더보기",
    "사이트맵", "이전", "다음", "처음", "마지막", "바로가기",
    "home", "next", "prev", "more", "login"
}

def clean_text(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()

def normalize_url(url, base=None):
    if not url:
        return ""
    url = url.strip()
    if base:
        url = urljoin(base, url)
    p = urlparse(url)
    if p.scheme not in ("http", "https") or not p.netloc:
        return ""
    # fragment 제거
    return urlunparse((p.scheme, p.netloc, p.path, p.params, p.query, ""))

def same_domain(a, b):
    try:
        return urlparse(a).netloc.lower().replace("www.", "") == urlparse(b).netloc.lower().replace("www.", "")
    except Exception:
        return False

def build_session():
    s = requests.Session()
    retry = Retry(
        total=2,
        connect=2,
        read=2,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"]
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/131.0 Safari/537.36"
        ),
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7",
    })
    return s

def fetch(session, url):
    try:
        r = session.get(url, timeout=TIMEOUT, allow_redirects=True, stream=True)
        content_type = (r.headers.get("content-type") or "").lower()

        if "text/html" not in content_type and "application/xhtml" not in content_type:
            r.close()
            return None, r.status_code, content_type, str(r.url), "HTML 아님"

        chunks = []
        total = 0
        for chunk in r.iter_content(16384):
            if not chunk:
                continue
            chunks.append(chunk)
            total += len(chunk)
            if total >= MAX_HTML:
                break
        raw = b"".join(chunks)[:MAX_HTML]
        final_url = str(r.url)
        status = r.status_code
        r.close()

        # 인코딩 방어
        enc = r.encoding or "utf-8"
        try:
            html = raw.decode(enc, errors="replace")
        except Exception:
            html = raw.decode("utf-8", errors="replace")

        return html, status, content_type, final_url, ""
    except Exception as e:
        return None, 0, "", url, f"{type(e).__name__}: {str(e)[:160]}"

def extract_title_from_anchor(a):
    # aria/title/data-title 우선
    for attr in ("title", "aria-label", "data-title"):
        v = clean_text(a.get(attr, ""))
        if len(v) >= 4 and v.lower() not in NON_POST_TEXT:
            return v

    # anchor 자체 텍스트
    v = clean_text(a.get_text(" ", strip=True))
    if 4 <= len(v) <= 300 and v.lower() not in NON_POST_TEXT:
        return v

    # 하위 heading
    for tag in a.find_all(["h1", "h2", "h3", "h4", "strong", "b"]):
        v = clean_text(tag.get_text(" ", strip=True))
        if 4 <= len(v) <= 300:
            return v
    return ""

def looks_like_detail(url):
    s = url.lower()
    p = urlparse(s)
    q = s.split("?", 1)[1] if "?" in s else ""
    if any(x in s for x in ["/noticeview", "/detail/", "/detailview", "/recruitview", "/newsview"]):
        return True
    if any(x in q for x in [
        "mode=view", "act=view", "articleno=", "article_seq=", "articleid=",
        "boardidx=", "bbidx=", "nttid=", "wr_id=", "list_no=", "vno=", "linkid="
    ]):
        return True
    if re.search(r"/view\.(do|asp)(?:\?|$)", p.path):
        return True
    return False

def classify_url(url, final_url, soup, links):
    if looks_like_detail(final_url):
        # 단, 실제 페이지에 다수의 게시물 링크가 있으면 목록으로 재분류
        if len(links) >= 5:
            return "목록"
        return "상세"

    path = urlparse(final_url).path.lower()
    text = clean_text(soup.get_text(" ", strip=True))[:50000].lower()

    list_kw = sum(1 for k in LIST_KEYWORDS if k.lower() in text or k.lower() in path)

    # 반복되는 내부 링크가 있으면 목록 가능성이 높음
    if len(links) >= 5 and list_kw >= 1:
        return "목록"
    if len(links) >= 3 and list_kw >= 2:
        return "목록"
    if list_kw >= 3 and any(x in path for x in ("list", "board", "bbs", "notice", "news")):
        return "목록"

    # 홈페이지
    if path in ("", "/") and len(links) < 5:
        return "홈페이지"

    return "미분류"

def extract_post_links(soup, base_url):
    candidates = []
    seen = set()

    for a in soup.find_all("a", href=True):
        href = normalize_url(a.get("href"), base_url)
        if not href or not same_domain(href, base_url):
            continue

        title = extract_title_from_anchor(a)
        if len(title) < 4:
            continue

        low = href.lower()
        title_low = title.lower()

        # 메뉴/기능 링크 제외
        if title_low in NON_POST_TEXT:
            continue
        if any(x in low for x in (
            "login", "logout", "search", "sitemap", "privacy",
            "member", "mypage", "javascript:"
        )):
            continue

        # 게시물 링크 신호
        score = 0
        if looks_like_detail(href):
            score += 4
        if any(k in title_low for k in ("공지", "공고", "채용", "모집", "안내", "보도", "입찰", "소식")):
            score += 1
        if re.search(r"\d{4}[-./]\d{1,2}[-./]\d{1,2}", title):
            score += 1
        if len(title) >= 8:
            score += 1

        if score < 2:
            continue

        key = href.rstrip("/")
        if key in seen:
            continue
        seen.add(key)

        candidates.append((href, title, score))

    candidates.sort(key=lambda x: x[2], reverse=True)
    return candidates

def extract_date(text):
    patterns = [
        r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2})",
        r"(20\d{2}\.\s*\d{1,2}\.\s*\d{1,2})",
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return re.sub(r"\s+", "", m.group(1)).replace(".", "-").replace("/", "-")
    return ""

def extract_detail_info(session, url, anchor_title=""):
    html, status, ctype, final_url, err = fetch(session, url)
    if not html:
        return {
            "url": url, "title": anchor_title, "date": "", "body_len": 0,
            "status": status, "final_url": final_url, "verified": False,
            "error": err
        }

    soup = BeautifulSoup(html, "lxml")

    title = ""
    # 여러 방법을 병행하여 제목 추출
    selectors = [
        "h1", "h2", "h3",
        ".view-title", ".board-title", ".bbs-title",
        ".subject", ".tit", ".title",
        "meta[property='og:title']",
        "meta[name='twitter:title']",
        "title"
    ]

    for sel in selectors:
        node = soup.select_one(sel)
        if not node:
            continue
        if node.name == "meta":
            v = clean_text(node.get("content", ""))
        else:
            v = clean_text(node.get_text(" ", strip=True))
        if 4 <= len(v) <= 500:
            title = v
            break

    if not title:
        title = anchor_title

    body = clean_text(soup.get_text(" ", strip=True))
    date = extract_date(body)

    # 너무 짧은 에러/리다이렉트 페이지 제외
    verified = status == 200 and len(body) >= 300 and len(title) >= 4

    return {
        "url": url,
        "title": title[:500],
        "date": date,
        "body_len": len(body),
        "status": status,
        "final_url": final_url,
        "verified": verified,
        "error": err
    }

def validate_one(row):
    session = build_session()
    institution = clean_text(row["기관명"])
    board_name = clean_text(row["게시판명"])
    url = clean_text(row["게시판URL"])

    result = {
        "기관명": institution,
        "게시판명": board_name,
        "게시판URL": url,
        "V7최종URL": "",
        "V7HTTP": 0,
        "V7접속": "실패",
        "V7최종유형": "접속실패",
        "V7리다이렉트": "",
        "V7목록키워드수": 0,
        "V7게시물링크수": 0,
        "V7검증게시물수": 0,
        "V7고유제목수": 0,
        "V7목록행수": 0,
        "V7페이지네이션": "없음",
        "V7검증제목": "",
        "V7검증날짜": "",
        "V7검증본문길이": "",
        "V7결과": "제외",
        "V7사유": "",
        "V7오류": ""
    }

    if not url.startswith(("http://", "https://")):
        result["V7사유"] = "HTTP/HTTPS URL 아님"
        return result

    html, status, ctype, final_url, err = fetch(session, url)
    result["V7HTTP"] = status
    result["V7최종URL"] = final_url
    result["V7접속"] = "성공" if html and status == 200 else "실패"
    result["V7리다이렉트"] = "있음" if final_url.rstrip("/") != url.rstrip("/") else "없음"

    if not html:
        result["V7사유"] = err or "HTML 응답 없음"
        result["V7오류"] = err
        return result

    soup = BeautifulSoup(html, "lxml")
    post_links = extract_post_links(soup, final_url)

    # 목록 행 수 추정
    row_count = len(soup.find_all("tr"))
    list_kw = sum(
        1 for k in LIST_KEYWORDS
        if k.lower() in clean_text(soup.get_text(" ", strip=True))[:50000].lower()
    )

    result["V7목록행수"] = row_count
    result["V7목록키워드수"] = list_kw
    result["V7게시물링크수"] = len(post_links)

    # pagination 신호
    pag_text = clean_text(soup.get_text(" ", strip=True)).lower()
    pag_links = soup.select("a[rel='next'], .pagination a, .paging a, .paginate a")
    if pag_links or any(x in pag_text for x in ("다음", "next", "1 2 3", "페이지")):
        result["V7페이지네이션"] = "있음"

    final_type = classify_url(url, final_url, soup, post_links)
    result["V7최종유형"] = final_type

    if final_type == "홈페이지":
        result["V7결과"] = "제외"
        result["V7사유"] = "홈페이지 성격"
        return result

    # 실제 게시물 최대 5개 확인
    checked = []
    for post_url, anchor_title, _score in post_links[:MAX_POST_CHECK]:
        info = extract_detail_info(session, post_url, anchor_title)
        if info["verified"]:
            checked.append(info)

    titles = [x["title"] for x in checked if x["title"]]
    unique_titles = list(dict.fromkeys(titles))

    result["V7검증게시물수"] = len(checked)
    result["V7고유제목수"] = len(unique_titles)
    result["V7검증제목"] = " | ".join(unique_titles[:5])
    result["V7검증날짜"] = " | ".join(x["date"] for x in checked[:5])
    result["V7검증본문길이"] = " | ".join(str(x["body_len"]) for x in checked[:5])

    # 보수적 판정
    if (
        final_type == "목록"
        and len(post_links) >= 5
        and len(checked) >= 3
        and len(unique_titles) >= 3
    ):
        result["V7결과"] = "확정"
        result["V7사유"] = "실제 목록 + 게시물 링크 + 개별 게시물 3건 이상 검증"
    elif (
        final_type == "목록"
        and len(post_links) >= 3
        and len(checked) >= 1
    ):
        result["V7결과"] = "수동확인"
        result["V7사유"] = "목록 구조는 확인되나 게시물 검증 수가 부족"
    elif final_type == "상세":
        result["V7결과"] = "제외"
        result["V7사유"] = "개별 게시물 상세페이지"
    else:
        result["V7결과"] = "수동확인"
        result["V7사유"] = f"페이지 유형={final_type}, 게시물링크={len(post_links)}, 검증={len(checked)}"

    return result

def main():
    print("=" * 70)
    print("V7 실제 웹페이지 구조 검증")
    print("=" * 70)

    df = pd.read_excel(INPUT_FILE, dtype=object)
    required = {"기관명", "게시판명", "게시판URL"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"필수 컬럼 없음: {missing}")

    print(f"입력파일 : {INPUT_FILE}")
    print(f"검증대상 : {len(df)}")
    print(f"동시접속 : {CONCURRENCY}")
    print()

    records = []
    start = time.time()

    with ThreadPoolExecutor(max_workers=CONCURRENCY) as ex:
        futures = {
            ex.submit(validate_one, row.to_dict()): i
            for i, (_, row) in enumerate(df.iterrows(), start=1)
        }

        done = 0
        for future in as_completed(futures):
            done += 1
            idx = futures[future]
            try:
                rec = future.result()
            except Exception as e:
                rec = {
                    "기관명": clean_text(df.iloc[idx-1]["기관명"]),
                    "게시판명": clean_text(df.iloc[idx-1]["게시판명"]),
                    "게시판URL": clean_text(df.iloc[idx-1]["게시판URL"]),
                    "V7최종URL": "",
                    "V7HTTP": 0,
                    "V7접속": "오류",
                    "V7최종유형": "오류",
                    "V7리다이렉트": "",
                    "V7목록키워드수": 0,
                    "V7게시물링크수": 0,
                    "V7검증게시물수": 0,
                    "V7고유제목수": 0,
                    "V7목록행수": 0,
                    "V7페이지네이션": "",
                    "V7검증제목": "",
                    "V7검증날짜": "",
                    "V7검증본문길이": "",
                    "V7결과": "오류",
                    "V7사유": "예외 발생",
                    "V7오류": str(e)[:500],
                }

            records.append(rec)

            if done % 10 == 0 or done == len(df):
                print(
                    f"[V7] {done}/{len(df)} | "
                    f"{rec['기관명']} | {rec['V7결과']} | "
                    f"유형 {rec['V7최종유형']} | "
                    f"링크 {rec['V7게시물링크수']} | "
                    f"검증 {rec['V7검증게시물수']}"
                )

    result_df = pd.DataFrame(records)

    # 원본 컬럼 순서를 먼저 유지하고 V7 컬럼을 뒤에 추가
    v7_cols = [c for c in result_df.columns if c.startswith("V7")]
    result_df = result_df[["기관명", "게시판명", "게시판URL"] + v7_cols]

    # 기관명/URL/게시판명 기준 정렬은 하지 않고 원본 순서 복원
    order = {clean_text(v): i for i, v in enumerate(df["게시판URL"].tolist())}
    result_df["_order"] = result_df["게시판URL"].map(order).fillna(999999)
    result_df = result_df.sort_values("_order").drop(columns=["_order"])

    # Excel 저장
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="V7검증결과", index=False)

        summary = (
            result_df["V7결과"]
            .value_counts(dropna=False)
            .rename_axis("결과")
            .reset_index(name="건수")
        )
        summary.to_excel(writer, sheet_name="요약", index=False)

        confirmed = result_df[result_df["V7결과"] == "확정"].copy()
        confirmed.to_excel(writer, sheet_name="확정후보", index=False)

        manual = result_df[result_df["V7결과"] == "수동확인"].copy()
        manual.to_excel(writer, sheet_name="수동확인", index=False)

        excluded = result_df[result_df["V7결과"] == "제외"].copy()
        excluded.to_excel(writer, sheet_name="제외", index=False)

    elapsed = time.time() - start

    print()
    print("=" * 70)
    print("V7 검증 완료")
    print("=" * 70)
    print(result_df["V7결과"].value_counts().to_string())
    print(f"소요시간 : {elapsed/60:.1f}분")
    print(f"결과파일 : {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
