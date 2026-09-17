import os
import re
import time
from urllib.parse import urljoin, urlparse, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

INPUT = "boards_v7_5_final_validation.xlsx"
OUTPUT = "boards_v7_6_final_validation.xlsx"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
TIMEOUT = 15

DATE_RE = re.compile(
    r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})"
    r"|20\d{2}\s*[년./-]\s*\d{1,2}\s*[월./-]\s*\d{1,2}"
)
BAD = {
    "공지사항", "공지", "알림", "새소식", "소식", "더보기", "목록",
    "상세", "검색", "홈", "로그인", "회원가입", "사이트맵",
    "스킵네비게이션", "본문 바로가기", "이전글", "다음글",
    "이전글이 없습니다.", "다음글이 없습니다.",
}

def clean(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()

def norm_url(u):
    if not isinstance(u, str) or not u.strip():
        return ""
    p = urlparse(u.strip())
    return urlunparse((p.scheme.lower(), p.netloc.lower(), p.path, p.params, p.query, ""))

def same_host(a, b):
    try:
        return urlparse(a).netloc.lower() == urlparse(b).netloc.lower()
    except Exception:
        return False

def fetch(s, url):
    try:
        r = s.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        r.encoding = r.apparent_encoding or r.encoding
        return r, BeautifulSoup(r.text, "html.parser")
    except Exception:
        return None, None

def good_title(t, inst=""):
    t = clean(t)
    if len(t) < 4 or len(t) > 220 or t == clean(inst):
        return False
    if t in BAD:
        return False
    if re.match(r"^(http|javascript|mailto|tel):", t, re.I):
        return False
    if any(x in t for x in ["로그인", "회원가입", "사이트맵", "스킵네비게이션"]):
        return False
    return True

def list_rows(soup, base, inst):
    """
    게시판의 반복되는 행/리스트를 찾는다.
    각 행에서 제목 링크와 날짜/번호 등의 구조적 증거를 같이 수집한다.
    """
    rows = []
    selectors = [
        "table tbody tr", "table tr",
        "ul.board_list li", "ol.board_list li",
        "ul.bbs_list li", "ol.bbs_list li",
        "div.board_list li", "div.bbs_list li",
        "div.notice_list li", "div.notice-list li",
        "div.board-list li", "article",
    ]
    for sel in selectors:
        try:
            rows.extend(soup.select(sel))
        except Exception:
            pass

    out, seen = [], set()
    for row in rows:
        txt = clean(row.get_text(" ", strip=True))
        if not txt or txt in seen:
            continue
        seen.add(txt)
        links = row.select("a[href]")
        for a in links:
            title = clean(a.get_text(" ", strip=True))
            href = a.get("href", "")
            if not href or not good_title(title, inst):
                continue
            if href.lower().startswith(("javascript:", "mailto:", "tel:", "#")):
                continue
            u = norm_url(urljoin(base, href))
            if not same_host(u, base):
                continue

            score = 0
            if DATE_RE.search(txt):
                score += 3
            if re.search(r"\b\d{1,7}\b", txt):
                score += 1
            if any(k in title for k in ["공지", "공고", "안내", "모집", "채용", "입찰"]):
                score += 1
            if any(k in href.lower() for k in ["view", "detail", "article", "read", "seq", "idx", "no="]):
                score += 1

            out.append({
                "title": title,
                "url": u,
                "row_text": txt,
                "score": score
            })

    # fallback
    if len(out) < 5:
        for a in soup.select("main a[href], #content a[href], #contents a[href], .content a[href]"):
            title = clean(a.get_text(" ", strip=True))
            if not good_title(title, inst):
                continue
            u = norm_url(urljoin(base, a.get("href", "")))
            if same_host(u, base):
                out.append({"title": title, "url": u, "row_text": title, "score": 0})

    result, seen_url = [], set()
    for x in sorted(out, key=lambda z: z["score"], reverse=True):
        if x["url"] not in seen_url:
            seen_url.add(x["url"])
            result.append(x)
    return result[:15]

def page_signature(soup, final_url, inst):
    """
    URL 자체가 독립적인 게시판 목록인지 평가한다.
    게시물 링크가 많다는 사실만으로는 통과시키지 않는다.
    """
    text = clean(soup.get_text(" ", strip=True))
    rows = list_rows(soup, final_url, inst)
    titles = []
    for x in rows:
        if x["title"] not in titles:
            titles.append(x["title"])

    signals = {
        "row_5plus": len(rows) >= 5,
        "unique_title_3plus": len(titles) >= 3,
        "date_2plus": len([x for x in rows if DATE_RE.search(x["row_text"])]) >= 2,
        "pagination": bool(soup.select(
            "a[href*='page'], a[href*='Page'], .pagination, .paging, "
            ".paginate, .pager, .board_paging, .bbs_paging"
        )),
        "table": bool(soup.select("table")),
        "list_class": bool(soup.select(
            ".board_list, .bbs_list, .notice_list, .board-list, .bbs-list, "
            ".notice-list"
        )),
        "list_words": sum(
            1 for k in ["공지사항", "게시판", "등록일", "작성일", "조회수", "번호", "제목"]
            if k in text
        ) >= 2,
    }

    # 게시판 고유 구조 점수
    score = 0
    if signals["row_5plus"]: score += 3
    if signals["unique_title_3plus"]: score += 2
    if signals["date_2plus"]: score += 2
    if signals["pagination"]: score += 2
    if signals["table"]: score += 1
    if signals["list_class"]: score += 1
    if signals["list_words"]: score += 1

    # 홈/메뉴 페이지의 위험 신호
    menu_risk = 0
    if len(rows) < 3:
        menu_risk += 2
    if not signals["date_2plus"]:
        menu_risk += 2
    if not signals["pagination"] and not signals["table"] and not signals["list_class"]:
        menu_risk += 2

    return score, menu_risk, rows, titles, signals

def verify_post(s, item, board_url, inst):
    r, soup = fetch(s, item["url"])
    if r is None or soup is None or r.status_code >= 400:
        return False, "접속실패", ""
    final = norm_url(r.url)
    if not same_host(final, board_url):
        return False, "외부도메인", final

    text = clean(soup.get_text(" ", strip=True))
    title = item["title"]
    # V7.5 원칙 유지: 상세 <title>/<h1>은 게시물 제목으로 사용하지 않는다.
    score = 0
    if title in text:
        score += 4
    words = [w for w in re.split(r"\s+", title) if len(w) >= 2]
    if words:
        hit = sum(1 for w in words[:10] if w in text)
        if hit >= max(2, min(4, len(words))):
            score += 1
    if DATE_RE.search(text): score += 1
    if len(text) >= 500: score += 1
    if any(k in text for k in ["첨부파일", "첨부", "다운로드", "작성자", "등록일", "조회수", "본문"]):
        score += 1

    return score >= 4, "목록제목-상세본문 일치" if score >= 4 else "상세본문 증거 부족", final

def content_fingerprint(soup):
    """
    게시판 중복 판단용 간단 fingerprint.
    실제 게시물 제목/링크의 집합을 기준으로 동일 게시판을 탐지한다.
    """
    text = clean(soup.get_text(" ", strip=True))
    return re.sub(r"\d+", "#", text[:5000]).lower()

def classify(rows, s):
    results = []
    for i, row in rows.iterrows():
        inst = clean(row.get("기관명", ""))
        board = norm_url(row.get("V7.5최종URL", "") or row.get("게시판URL", ""))
        r, soup = fetch(s, board)

        d = dict(row)
        if r is None or soup is None or r.status_code >= 400:
            d.update({
                "V7.6판정": "ANALYZE",
                "V7.6사유": "게시판 접속 실패",
                "V7.6게시판점수": 0,
                "V7.6메뉴위험도": 99,
                "V7.6목록후보수": 0,
                "V7.6실제게시물수": 0,
                "V7.6고유제목수": 0,
                "V7.6최종URL": board,
                "V7.6Fingerprint": "",
            })
            results.append(d)
            continue

        final = norm_url(r.url)
        score, risk, candidates, titles, signals = page_signature(soup, final, inst)

        verified = []
        for item in candidates[:10]:
            ok, reason, detail = verify_post(s, item, final, inst)
            if ok:
                verified.append((item["title"], detail))
            time.sleep(0.10)

        unique_verified = list(dict.fromkeys(x[0] for x in verified))

        # 독립 게시판 판정
        independent = (
            score >= 7
            and len(candidates) >= 5
            and len(titles) >= 3
            and (signals["date_2plus"] or signals["pagination"] or signals["table"])
        )

        if independent and len(verified) >= 5 and len(unique_verified) >= 4:
            decision = "FINAL"
            reason = "독립 게시판 구조 + 실제 상세 게시물 5건 이상 검증"
        elif independent and len(verified) >= 3 and len(unique_verified) >= 3:
            decision = "KEEP"
            reason = "독립 게시판 구조 확인, 일부 보완 필요"
        elif len(verified) >= 1:
            decision = "ANALYZE"
            reason = "게시물은 확인되나 독립 게시판 구조 증거 부족"
        else:
            decision = "REMOVE"
            reason = "독립 게시판 또는 실제 게시물 검증 실패"

        d.update({
            "V7.6최종URL": final,
            "V7.6HTTP": r.status_code,
            "V7.6게시판점수": score,
            "V7.6메뉴위험도": risk,
            "V7.6목록후보수": len(candidates),
            "V7.6실제게시물수": len(verified),
            "V7.6고유제목수": len(unique_verified),
            "V7.6검증제목": " | ".join(unique_verified[:10]),
            "V7.6검증URL": " | ".join(x[1] for x in verified[:10]),
            "V7.6구조신호": "; ".join(k for k,v in signals.items() if v),
            "V7.6판정": decision,
            "V7.6사유": reason,
            "V7.6Fingerprint": content_fingerprint(soup),
        })
        results.append(d)

    return pd.DataFrame(results)

def dedupe_final(df):
    """
    동일 최종 URL 또는 사실상 동일한 게시판 fingerprint를 기관별로 제거.
    자동 삭제하지 않고 DUPLICATE로 표시하여 사람이 확인할 수 있게 한다.
    """
    df = df.copy()
    df["V7.6중복URL"] = ""
    df["V7.6중복수"] = 0
    df["V7.6중복Fingerprint"] = ""

    url_counts = df["V7.6최종URL"].fillna("").value_counts()
    for idx in df.index:
        u = df.at[idx, "V7.6최종URL"]
        if u and url_counts.get(u, 0) > 1:
            df.at[idx, "V7.6중복URL"] = "중복"
            df.at[idx, "V7.6중복수"] = int(url_counts[u])

    # 기관별 fingerprint 중복
    if "V7.6Fingerprint" in df.columns:
        for inst, g in df.groupby("기관명"):
            fp_counts = g["V7.6Fingerprint"].fillna("").value_counts()
            for idx in g.index:
                fp = df.at[idx, "V7.6Fingerprint"]
                if fp and fp_counts.get(fp, 0) > 1:
                    df.at[idx, "V7.6중복Fingerprint"] = "유사/중복"

    # FINAL 중복은 자동 확정에서 제외
    dup = (df["V7.6중복URL"] == "중복") | (df["V7.6중복Fingerprint"] == "유사/중복")
    df.loc[(df["V7.6판정"] == "FINAL") & dup, "V7.6판정"] = "KEEP"
    df.loc[(df["V7.6판정"] == "KEEP") & dup, "V7.6사유"] = "중복 게시판 후보: 대표 URL 1개 선정 필요"

    return df

def main():
    if not os.path.exists(INPUT):
        raise SystemExit(f"ERROR: {INPUT} not found.")

    xls = pd.ExcelFile(INPUT)
    if "V7.5전체" in xls.sheet_names:
        df = pd.read_excel(INPUT, sheet_name="V7.5전체")
    elif "A_최종확정" in xls.sheet_names:
        df = pd.read_excel(INPUT, sheet_name="A_최종확정")
    else:
        raise SystemExit("ERROR: V7.5 result sheet not found.")

    # V7.5 52개 대상. 동일 게시판URL은 사전에 1회만 검사.
    if "게시판URL" in df.columns:
        df["_key"] = df["게시판URL"].astype(str).map(norm_url)
        df = df.drop_duplicates("_key").drop(columns=["_key"])

    print(f"V7.6 대상: {len(df)}개")
    s = requests.Session()
    out = classify(df, s)
    out = dedupe_final(out)

    summary = out.groupby("V7.6판정", dropna=False).size().reset_index(name="게시판수")
    inst = out.groupby(["기관명", "V7.6판정"], dropna=False).size().reset_index(name="게시판수")

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as w:
        out.to_excel(w, sheet_name="V7.6전체", index=False)
        out[out["V7.6판정"]=="FINAL"].to_excel(w, sheet_name="FINAL", index=False)
        out[out["V7.6판정"]=="KEEP"].to_excel(w, sheet_name="KEEP", index=False)
        out[out["V7.6판정"]=="ANALYZE"].to_excel(w, sheet_name="ANALYZE", index=False)
        out[out["V7.6판정"]=="REMOVE"].to_excel(w, sheet_name="REMOVE", index=False)
        out[(out["V7.6중복URL"]=="중복") | (out["V7.6중복Fingerprint"]=="유사/중복")].to_excel(
            w, sheet_name="DUPLICATE", index=False
        )
        summary.to_excel(w, sheet_name="요약", index=False)
        inst.to_excel(w, sheet_name="기관별요약", index=False)

    print("\n=== V7.6 완료 ===")
    print(summary.to_string(index=False))
    print(f"결과파일: {OUTPUT}")

if __name__ == "__main__":
    main()
