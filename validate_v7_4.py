import re
import time
from urllib.parse import urljoin, urlparse, parse_qsl, urlencode, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

INPUT = "boards_v7_3_grade2_validation.xlsx"
SHEET = "확정"
OUTPUT = "boards_v7_4_final_validation.xlsx"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
TIMEOUT = 15

# 게시물 제목으로 보기 어려운 문구
BAD_TITLE_PATTERNS = [
    r"^공지사항$", r"^공지$", r"^알림$", r"^새소식$", r"^소식$",
    r"^더보기$", r"^목록$", r"^상세$", r"^게시판$", r"^검색$",
    r"^홈$", r"^로그인$", r"^회원가입$", r"^사이트맵$",
    r"^외래진료일정", r"^진료비", r"^진료예약", r"^예약$",
    r"^이전글(이 없습니다)?$", r"^다음글(이 없습니다)?$",
    r"^한국고전선집$", r"^고전대중화 도서$",
]

FUNCTION_WORDS = [
    "로그인", "회원가입", "검색", "사이트맵", "메뉴", "바로가기",
    "예약", "진료", "결제", "오시는 길", "찾아오시는 길",
    "개인정보처리방침", "이용약관", "저작권", "문의", "뉴스레터",
]

DATE_RE = re.compile(
    r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})"
    r"|20\d{2}\s*[년./-]\s*\d{1,2}\s*[월./-]\s*\d{1,2}"
)

def clean_text(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()

def norm_url(u):
    if not u or not isinstance(u, str):
        return ""
    u = u.strip()
    p = urlparse(u)
    # fragment 제거
    return urlunparse((p.scheme, p.netloc, p.path, p.params, p.query, ""))

def same_host(a, b):
    try:
        return urlparse(a).netloc.lower() == urlparse(b).netloc.lower()
    except Exception:
        return False

def looks_function(text):
    t = clean_text(text)
    if len(t) < 2:
        return True
    return any(w in t for w in FUNCTION_WORDS)

def good_title(t, institution=""):
    t = clean_text(t)
    if len(t) < 5 or len(t) > 180:
        return False
    for p in BAD_TITLE_PATTERNS:
        if re.search(p, t, re.I):
            return False
    if t == clean_text(institution):
        return False
    if t.count("http") >= 1 and len(t) < 80:
        return False
    # 메뉴/기능성 문구만 있는 경우 제외
    if looks_function(t) and not any(k in t for k in ["공고", "모집", "안내", "공지", "채용", "입찰"]):
        return False
    return True

def fetch(session, url):
    try:
        r = session.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        r.encoding = r.apparent_encoding or r.encoding
        return r, BeautifulSoup(r.text, "html.parser")
    except Exception:
        return None, None

def extract_candidates(soup, base_url, institution):
    # 게시물 링크 후보를 찾되 메뉴/푸터/검색 영역을 최대한 배제
    roots = []
    selectors = [
        "table tbody tr", "table tr",
        "ul.board_list li", "ol.board_list li",
        "ul.bbs_list li", "ol.bbs_list li",
        "div.board_list li", "div.bbs_list li",
        "div.notice_list li", "div.list li",
        "article", ".board-list", ".bbs-list", ".notice-list",
    ]
    for sel in selectors:
        try:
            roots.extend(soup.select(sel))
        except Exception:
            pass

    # root 중복 제거
    seen_roots = set()
    unique_roots = []
    for x in roots:
        key = str(x)[:1500]
        if key not in seen_roots:
            seen_roots.add(key)
            unique_roots.append(x)

    out = []
    seen_urls = set()

    def add(a, title, score):
        href = a.get("href")
        if not href:
            return
        u = norm_url(urljoin(base_url, href))
        title = clean_text(title)
        if not u or not same_host(u, base_url):
            return
        if u in seen_urls or not good_title(title, institution):
            return
        # javascript/mailto/tel 등
        if href.lower().startswith(("javascript:", "mailto:", "tel:", "#")):
            return
        seen_urls.add(u)
        out.append((u, title, score))

    for root in unique_roots:
        for a in root.select("a[href]"):
            txt = clean_text(a.get_text(" ", strip=True))
            if not good_title(txt, institution):
                continue
            score = 0
            href = a.get("href", "")
            parent_text = clean_text(root.get_text(" ", strip=True))
            if DATE_RE.search(parent_text):
                score += 2
            if any(k in txt for k in ["공지", "공고", "안내", "모집", "채용", "입찰"]):
                score += 1
            if any(k in href.lower() for k in ["view", "detail", "article", "read", "seq", "idx", "no="]):
                score += 1
            add(a, txt, score)

    # 구조화된 목록을 못 잡은 경우: 본문 영역의 링크를 보조 탐색
    if len(out) < 5:
        for a in soup.select("main a[href], #content a[href], #contents a[href], .content a[href]"):
            txt = clean_text(a.get_text(" ", strip=True))
            if not good_title(txt, institution):
                continue
            add(a, txt, 1)

    # 날짜가 있는 후보 우선
    out.sort(key=lambda x: x[2], reverse=True)
    return out[:20]

def detail_title(soup, institution):
    sels = [
        "h1", "h2", "h3", ".view_title", ".board_view_title",
        ".bbs_view_title", ".subject", ".title", "article h1", "article h2",
        "meta[property='og:title']", "title"
    ]
    candidates = []
    for sel in sels:
        for el in soup.select(sel):
            if el.name == "meta":
                t = clean_text(el.get("content", ""))
            else:
                t = clean_text(el.get_text(" ", strip=True))
            if good_title(t, institution):
                candidates.append(t)
    # 가장 구체적인 제목을 선택
    if candidates:
        # site title이 뒤에 붙은 긴 문자열보다 짧고 의미 있는 제목 선호
        candidates = list(dict.fromkeys(candidates))
        candidates.sort(key=lambda x: (len(x) > 120, len(x)))
        return candidates[0]
    return ""

def detail_date(soup):
    text = clean_text(soup.get_text(" ", strip=True))
    m = DATE_RE.search(text)
    return m.group(0) if m else ""

def detail_body_score(soup, title):
    # 실제 상세 페이지인지 확인하는 약한 증거
    text = clean_text(soup.get_text(" ", strip=True))
    score = 0
    if len(text) >= 500:
        score += 1
    if len(text) >= 1000:
        score += 1
    if DATE_RE.search(text):
        score += 1
    if title and title in text:
        score += 1
    # 첨부파일/본문/목록 버튼 등의 흔적
    if any(k in text for k in ["첨부파일", "첨부", "다운로드", "목록", "본문"]):
        score += 1
    return score

def verify_detail(session, url, candidate_title, institution):
    r, soup = fetch(session, url)
    if soup is None or r is None or r.status_code >= 400:
        return False, "", "", 0
    title = detail_title(soup, institution)
    if not title:
        title = candidate_title if good_title(candidate_title, institution) else ""
    score = detail_body_score(soup, title)
    # 실제 상세로 판단할 최소 조건
    ok = bool(title) and score >= 2
    return ok, title, detail_date(soup), score

def classify(row, session):
    inst = clean_text(row.get("기관명", ""))
    original = norm_url(row.get("게시판URL", ""))
    v73 = norm_url(row.get("V7.3최종URL", "")) or original
    result = dict(row)

    r, soup = fetch(session, v73)
    if soup is None or r is None:
        result.update({
            "V7.4판정": "오류",
            "V7.4사유": "게시판 URL 접속 실패",
            "V7.4후보수": 0, "V7.4검증수": 0, "V7.4고유제목수": 0,
            "V7.4검증제목": "", "V7.4검증날짜": "", "V7.4본문점수": "",
        })
        return result

    final_url = norm_url(r.url)
    cands = extract_candidates(soup, final_url, inst)

    verified = []
    for u, title, sc in cands[:10]:
        ok, dt, date, body_score = verify_detail(session, u, title, inst)
        if ok:
            verified.append((u, dt or title, date, body_score))
        time.sleep(0.15)

    titles = []
    for _, t, _, _ in verified:
        t = clean_text(t)
        if good_title(t, inst) and t not in titles:
            titles.append(t)

    dates = [x[2] for x in verified if x[2]]
    body_scores = [str(x[3]) for x in verified]

    # 중복/품질 이상 탐지
    duplicate_title_ratio = 0
    if verified:
        raw_titles = [clean_text(x[1]) for x in verified]
        duplicate_title_ratio = 1 - (len(set(raw_titles)) / len(raw_titles))

    reasons = []
    # URL이 다른 게시판 URL로 정규화/리디렉션되는 경우
    if original and final_url and urlparse(original).path != urlparse(final_url).path:
        reasons.append("최종 URL 경로 변경")
    if duplicate_title_ratio >= 0.6:
        reasons.append("검증 제목 중복률 높음")
    if len(titles) >= 3:
        reasons.append("서로 다른 실제 제목 확인")
    if any(len(t) > 120 for t in titles):
        reasons.append("장문 제목 포함")
    if any(not DATE_RE.search(t) and len(t) < 8 for t in titles):
        reasons.append("짧은 제목 포함")

    # 최종 판정
    # A: 실제 게시물 5건 이상, 고유 제목 4건 이상, 본문/상세 증거 충분, 중복률 낮음
    if len(verified) >= 5 and len(titles) >= 4 and duplicate_title_ratio < 0.5:
        decision = "A-최종확정"
        reasons.insert(0, "실제 상세 게시물 5건 이상 + 고유 제목 4건 이상")
    # B: 자동화 가능하지만 일부 품질/중복 문제가 있음
    elif len(verified) >= 3 and len(titles) >= 3:
        decision = "B-보완후자동화"
        reasons.insert(0, "실제 게시물은 확인되나 제목/구조 보완 필요")
    elif len(verified) >= 1:
        decision = "C-추가분석"
        reasons.insert(0, "실제 상세 게시물 증거 부족")
    else:
        decision = "D-제외"
        reasons.insert(0, "실제 상세 게시물 확인 실패")

    result.update({
        "V7.4최종URL": final_url,
        "V7.4HTTP": r.status_code,
        "V7.4후보수": len(cands),
        "V7.4검증수": len(verified),
        "V7.4고유제목수": len(titles),
        "V7.4중복제목률": round(duplicate_title_ratio, 2),
        "V7.4검증제목": " | ".join(titles[:8]),
        "V7.4검증날짜": " | ".join(dates[:8]),
        "V7.4본문점수": " | ".join(body_scores[:8]),
        "V7.4판정": decision,
        "V7.4사유": " / ".join(reasons),
        "V7.4오류": "",
    })
    return result

def main():
    if not __import__("os").path.exists(INPUT):
        raise SystemExit(f"ERROR: {INPUT} not found.")

    df = pd.read_excel(INPUT, sheet_name=SHEET)
    # URL 중복은 모두 검사하되, 결과에서 중복 그룹을 표시
    session = requests.Session()

    rows = []
    total = len(df)
    print(f"V7.4 대상: {total}개")
    for i, row in df.iterrows():
        print(f"[{i+1}/{total}] {row.get('기관명','')} | {row.get('게시판명','')}")
        try:
            rows.append(classify(row.to_dict(), session))
        except Exception as e:
            d = dict(row)
            d.update({
                "V7.4판정": "오류", "V7.4사유": "검증 중 예외 발생",
                "V7.4오류": str(e),
            })
            rows.append(d)

    out = pd.DataFrame(rows)

    # 동일한 최종 URL은 중복 그룹 번호 부여
    counts = out["V7.4최종URL"].fillna("").value_counts()
    out["V7.4중복URL"] = out["V7.4최종URL"].map(lambda x: "중복" if x and counts.get(x, 0) > 1 else "")
    out["V7.4중복수"] = out["V7.4최종URL"].map(lambda x: int(counts.get(x, 0)) if x else 0)

    # 기관별 A/B/C/D 요약
    summary = (
        out.groupby("V7.4판정", dropna=False)
        .size().reset_index(name="게시판수")
        .sort_values("게시판수", ascending=False)
    )
    inst_summary = (
        out.groupby(["기관명", "V7.4판정"], dropna=False)
        .size().reset_index(name="게시판수")
        .sort_values(["기관명", "게시판수"], ascending=[True, False])
    )

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="V7.4전체", index=False)
        out[out["V7.4판정"] == "A-최종확정"].to_excel(writer, sheet_name="A_최종확정", index=False)
        out[out["V7.4판정"] == "B-보완후자동화"].to_excel(writer, sheet_name="B_보완", index=False)
        out[out["V7.4판정"] == "C-추가분석"].to_excel(writer, sheet_name="C_추가분석", index=False)
        out[out["V7.4판정"] == "D-제외"].to_excel(writer, sheet_name="D_제외", index=False)
        out[out["V7.4중복URL"] == "중복"].to_excel(writer, sheet_name="중복URL", index=False)
        summary.to_excel(writer, sheet_name="요약", index=False)
        inst_summary.to_excel(writer, sheet_name="기관별요약", index=False)

    print("\n=== V7.4 완료 ===")
    print(summary.to_string(index=False))
    print(f"결과파일: {OUTPUT}")

if __name__ == "__main__":
    main()
