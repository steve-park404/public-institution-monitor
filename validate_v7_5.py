import os
import re
import time
from urllib.parse import urljoin, urlparse, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

INPUT = "boards_v7_4_final_validation.xlsx"
OUTPUT = "boards_v7_5_final_validation.xlsx"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
TIMEOUT = 15

BAD_TEXT = [
    "스킵네비게이션", "본문 바로가기", "주메뉴", "전체메뉴", "로그인",
    "회원가입", "사이트맵", "검색", "더보기", "목록", "이전글",
    "다음글", "이전페이지", "다음페이지", "맨위로", "맨위로가기",
]
BAD_TITLE_RE = [
    r"^공지사항$", r"^공지$", r"^알림$", r"^새소식$", r"^소식$",
    r"^게시판$", r"^목록$", r"^상세$", r"^검색$", r"^홈$",
    r"^스킵네비게이션$", r"^NRC 소개$", r"^한국체육산업개발 홈페이지$",
    r"^이전글(이 없습니다)?$", r"^다음글(이 없습니다)?$",
    r"^외래진료일정", r"^진료비", r"^진료예약", r"^예약$",
]
FUNCTION_WORDS = [
    "로그인", "회원가입", "사이트맵", "검색", "바로가기", "메뉴",
    "예약", "진료", "결제", "오시는 길", "찾아오시는 길",
    "개인정보처리방침", "이용약관", "저작권",
]
DATE_RE = re.compile(
    r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})"
    r"|20\d{2}\s*[년./-]\s*\d{1,2}\s*[월./-]\s*\d{1,2}"
)

def clean(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()

def norm_url(u):
    if not isinstance(u, str) or not u.strip():
        return ""
    p = urlparse(u.strip())
    return urlunparse((p.scheme, p.netloc, p.path, p.params, p.query, ""))

def same_host(a, b):
    try:
        return urlparse(a).netloc.lower() == urlparse(b).netloc.lower()
    except Exception:
        return False

def good_title(t, inst=""):
    t = clean(t)
    if len(t) < 4 or len(t) > 220:
        return False
    if t == clean(inst):
        return False
    if any(re.search(p, t, re.I) for p in BAD_TITLE_RE):
        return False
    if any(x == t for x in BAD_TEXT):
        return False
    if t.lower().startswith(("http://", "https://", "javascript:")):
        return False
    # 순수 기능 문구는 제외하되 '공고/모집/채용/안내' 등이 포함된 실제 제목은 허용
    if any(w in t for w in FUNCTION_WORDS):
        if not any(k in t for k in ["공고", "모집", "채용", "안내", "입찰", "공지"]):
            return False
    return True

def fetch(s, url):
    try:
        r = s.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        r.encoding = r.apparent_encoding or r.encoding
        return r, BeautifulSoup(r.text, "html.parser")
    except Exception:
        return None, None

def find_row_candidates(soup, base_url, inst):
    """
    목록의 '행'을 먼저 찾고, 각 행에서 제목과 링크를 함께 가져온다.
    V7.5는 상세페이지의 <title>/<h1>을 게시물 제목으로 사용하지 않는다.
    """
    containers = []
    for sel in [
        "table tbody tr", "table tr",
        "ul.board_list li", "ol.board_list li",
        "ul.bbs_list li", "ol.bbs_list li",
        "div.board_list li", "div.bbs_list li",
        "div.notice_list li", "div.notice-list li",
        "div.board-list li", "article",
    ]:
        try:
            containers.extend(soup.select(sel))
        except Exception:
            pass

    seen = set()
    rows = []
    for c in containers:
        key = clean(c.get_text(" ", strip=True))
        if not key or key in seen:
            continue
        seen.add(key)
        links = c.select("a[href]")
        if not links:
            continue

        # 행 내부의 링크 중 제목 후보를 선정
        for a in links:
            href = a.get("href", "")
            title = clean(a.get_text(" ", strip=True))
            if not href or href.lower().startswith(("javascript:", "mailto:", "tel:", "#")):
                continue
            u = norm_url(urljoin(base_url, href))
            if not same_host(u, base_url):
                continue
            if not good_title(title, inst):
                continue

            text = clean(c.get_text(" ", strip=True))
            score = 0
            if DATE_RE.search(text):
                score += 3
            if any(k in title for k in ["공지", "공고", "안내", "모집", "채용", "입찰"]):
                score += 1
            if any(k in href.lower() for k in ["view", "detail", "article", "read", "seq", "idx", "no="]):
                score += 1
            # 행 안에 날짜/번호/조회수 등이 있으면 게시물 행일 가능성 상승
            if re.search(r"\b\d{1,6}\b", text):
                score += 1
            rows.append({"url": u, "title": title, "row_text": text, "score": score})

    # fallback: content 영역에서 링크 탐색
    if len(rows) < 5:
        for a in soup.select("main a[href], #content a[href], #contents a[href], .content a[href]"):
            title = clean(a.get_text(" ", strip=True))
            if not good_title(title, inst):
                continue
            u = norm_url(urljoin(base_url, a.get("href", "")))
            if same_host(u, base_url):
                rows.append({"url": u, "title": title, "row_text": title, "score": 0})

    # URL 중복 제거
    out = []
    seen_urls = set()
    for x in sorted(rows, key=lambda z: z["score"], reverse=True):
        if x["url"] not in seen_urls:
            seen_urls.add(x["url"])
            out.append(x)
    return out[:15]

def title_in_detail(soup, list_title, inst):
    """
    상세페이지의 공통 title/h1은 사용하지 않는다.
    목록에서 확보한 제목이 상세 본문에 존재하는지 확인한다.
    """
    lt = clean(list_title)
    if not good_title(lt, inst):
        return False, 0

    text = clean(soup.get_text(" ", strip=True))
    score = 0

    if lt in text:
        score += 3
    # 제목의 핵심어가 본문에 존재하는 경우
    words = [w for w in re.split(r"\s+", lt) if len(w) >= 2]
    if words:
        hit = sum(1 for w in words[:10] if w in text)
        if hit >= max(2, min(4, len(words))):
            score += 1

    if DATE_RE.search(text):
        score += 1
    if len(text) >= 500:
        score += 1
    if any(k in text for k in ["첨부파일", "첨부", "다운로드", "조회수", "작성자", "등록일", "본문"]):
        score += 1

    return score >= 3, score

def verify_post(s, item, inst, board_url):
    r, soup = fetch(s, item["url"])
    if r is None or soup is None or r.status_code >= 400:
        return False, "", "", 0, "접속실패"

    final = norm_url(r.url)
    # 다른 게시판/홈페이지로 이동하면 게시물 검증 실패
    if not same_host(final, board_url):
        return False, "", final, 0, "외부도메인"

    ok, score = title_in_detail(soup, item["title"], inst)
    if not ok:
        return False, "", final, score, "목록제목-본문불일치"

    text = clean(soup.get_text(" ", strip=True))
    date = ""
    m = DATE_RE.search(text)
    if m:
        date = m.group(0)

    return True, item["title"], final, score, "목록제목-본문일치"

def classify(row, s):
    inst = clean(row.get("기관명", ""))
    board = norm_url(row.get("V7.4최종URL", "") or row.get("게시판URL", ""))
    result = dict(row)

    r, soup = fetch(s, board)
    if r is None or soup is None:
        result.update({
            "V7.5판정": "C-추가분석",
            "V7.5사유": "게시판 접속 실패",
            "V7.5후보수": 0, "V7.5검증수": 0, "V7.5고유제목수": 0,
            "V7.5검증제목": "", "V7.5상세검증": "",
        })
        return result

    final_board = norm_url(r.url)
    candidates = find_row_candidates(soup, final_board, inst)

    verified = []
    for item in candidates[:10]:
        ok, title, final, score, reason = verify_post(s, item, inst, final_board)
        if ok:
            verified.append({
                "title": title, "url": item["url"], "detail_url": final,
                "score": score, "reason": reason
            })
        time.sleep(0.12)

    unique_titles = []
    for x in verified:
        if x["title"] not in unique_titles:
            unique_titles.append(x["title"])

    # 목록 제목을 기준으로 판단하므로 동일한 사이트 공통 h1/title 오인 문제를 차단
    duplicate_ratio = 0
    if verified:
        duplicate_ratio = 1 - len(unique_titles) / len(verified)

    if len(verified) >= 5 and len(unique_titles) >= 4 and duplicate_ratio < 0.5:
        decision = "A-최종확정"
        reason = "목록 제목 보존 + 실제 상세 본문 일치 게시물 5건 이상"
    elif len(verified) >= 3 and len(unique_titles) >= 3:
        decision = "B-보완후자동화"
        reason = "실제 게시물 확인. 모니터링용 구조 보완 가능"
    elif len(verified) >= 1:
        decision = "C-추가분석"
        reason = "실제 게시물 증거가 충분하지 않음"
    else:
        decision = "D-제외"
        reason = "목록 링크를 실제 게시물로 검증하지 못함"

    result.update({
        "V7.5최종URL": final_board,
        "V7.5HTTP": r.status_code,
        "V7.5후보수": len(candidates),
        "V7.5검증수": len(verified),
        "V7.5고유제목수": len(unique_titles),
        "V7.5중복제목률": round(duplicate_ratio, 2),
        "V7.5검증제목": " | ".join(unique_titles[:10]),
        "V7.5검증URL": " | ".join(x["detail_url"] for x in verified[:10]),
        "V7.5상세검증": " | ".join(f'{x["score"]}:{x["reason"]}' for x in verified[:10]),
        "V7.5판정": decision,
        "V7.5사유": reason,
    })
    return result

def main():
    if not os.path.exists(INPUT):
        raise SystemExit(f"ERROR: {INPUT} not found.")

    # V7.4의 A만 대상으로 하지 않고, V7.4 결과 전체를 읽되
    # 이번 작업의 핵심은 V7.3에서 '확정'된 52개를 다시 검증하는 것.
    xls = pd.ExcelFile(INPUT)
    if "V7.4전체" in xls.sheet_names:
        df = pd.read_excel(INPUT, sheet_name="V7.4전체")
    else:
        # 파일 구조가 다를 경우 A_최종확정 + B/C/D를 합쳐 중복 제거
        parts = []
        for sh in xls.sheet_names:
            if sh in ["A_최종확정", "B_보완", "C_추가분석", "D_제외"]:
                parts.append(pd.read_excel(INPUT, sheet_name=sh))
        if not parts:
            raise SystemExit("ERROR: V7.4 result sheets not found.")
        df = pd.concat(parts, ignore_index=True)

    # V7.3 확정 52개가 들어 있는 V7.4 전체에서 중복 제거
    if "게시판URL" in df.columns:
        df["_key"] = df["게시판URL"].astype(str).map(norm_url)
        df = df.drop_duplicates("_key").drop(columns=["_key"])

    print(f"V7.5 대상: {len(df)}개")

    s = requests.Session()
    rows = []
    for i, row in df.iterrows():
        print(f"[{i+1}/{len(df)}] {row.get('기관명','')} | {row.get('게시판명','')}")
        try:
            rows.append(classify(row.to_dict(), s))
        except Exception as e:
            d = dict(row)
            d.update({
                "V7.5판정": "C-추가분석",
                "V7.5사유": "검증 예외: " + str(e),
            })
            rows.append(d)

    out = pd.DataFrame(rows)

    # 최종 URL 중복
    vc = out["V7.5최종URL"].fillna("").value_counts()
    out["V7.5중복URL"] = out["V7.5최종URL"].map(lambda x: "중복" if x and vc.get(x, 0) > 1 else "")
    out["V7.5중복수"] = out["V7.5최종URL"].map(lambda x: int(vc.get(x, 0)) if x else 0)

    summary = out.groupby("V7.5판정", dropna=False).size().reset_index(name="게시판수")
    inst = out.groupby(["기관명", "V7.5판정"], dropna=False).size().reset_index(name="게시판수")

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as w:
        out.to_excel(w, sheet_name="V7.5전체", index=False)
        out[out["V7.5판정"]=="A-최종확정"].to_excel(w, sheet_name="A_최종확정", index=False)
        out[out["V7.5판정"]=="B-보완후자동화"].to_excel(w, sheet_name="B_보완", index=False)
        out[out["V7.5판정"]=="C-추가분석"].to_excel(w, sheet_name="C_추가분석", index=False)
        out[out["V7.5판정"]=="D-제외"].to_excel(w, sheet_name="D_제외", index=False)
        out[out["V7.5중복URL"]=="중복"].to_excel(w, sheet_name="중복URL", index=False)
        summary.to_excel(w, sheet_name="요약", index=False)
        inst.to_excel(w, sheet_name="기관별요약", index=False)

    print("\n=== V7.5 완료 ===")
    print(summary.to_string(index=False))
    print(f"결과파일: {OUTPUT}")

if __name__ == "__main__":
    main()
