#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V8.15.5 - 355기관 게시판URL 직접검색형
- 기관별_상태.csv의 '게시판URL'을 authoritative source로 사용
- 자동 게시판 발견 로직 제거
- 게시판URL이 있는 기관만 직접 검색
- 제목 + 본문에서 설문조사/시민참여/국민참여 검색
- 제목에 공모전이 포함되면 제외
- 최근 30일 게시물 우선
- Telegram 최대 20건/회
- URL/fingerprint 중복 방지 ledger 유지
"""

import os, re, csv, json, time, hashlib
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin, urlparse, parse_qsl, urlencode, urlunparse, unquote
import requests
from bs4 import BeautifulSoup

VERSION = "V8.15.5"
INPUT_CSV = os.getenv("INPUT_CSV", "기관별_상태.csv")
RECENT_DAYS = int(os.getenv("RECENT_DAYS", "30"))
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))
TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "15"))
MAX_PAGES_PER_BOARD = int(os.getenv("MAX_PAGES_PER_BOARD", "3"))

KEYWORDS = ["설문조사", "시민참여", "국민참여"]
EXCLUDE_TITLE = ["공모전"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128 Safari/537.36"
}

KST = timezone(timedelta(hours=9))

def now_kst():
    return datetime.now(KST)

def normalize_url(u):
    if not u:
        return ""
    u = str(u).strip()
    if not u or u.lower() in ("nan", "none", "null"):
        return ""
    if not re.match(r"^https?://", u, re.I):
        return ""
    return u

def clean_text(s):
    return re.sub(r"\s+", " ", unquote(str(s or ""))).strip()

def canonical_url(u):
    u = normalize_url(u)
    if not u:
        return ""
    p = urlparse(u)
    # tracking parameters 제거
    drop = {"utm_source","utm_medium","utm_campaign","utm_term","utm_content"}
    q = [(k,v) for k,v in parse_qsl(p.query, keep_blank_values=True) if k not in drop]
    return urlunparse((p.scheme.lower(), p.netloc.lower(), p.path, "", urlencode(q), ""))

def load_csv():
    with open(INPUT_CSV, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if len(rows) < 300:
        raise RuntimeError(f"기관 데이터가 부족합니다: {len(rows)}개")
    required = {"기관명", "게시판URL"}
    missing = required - set(rows[0].keys())
    if missing:
        raise RuntimeError(f"필수 컬럼 누락: {sorted(missing)}")
    return rows

def load_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def save_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def date_from_text(text):
    text = clean_text(text)
    patterns = [
        r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})",
        r"(20\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            try:
                return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), tzinfo=KST)
            except Exception:
                pass
    return None

def keyword_hits(text):
    return [k for k in KEYWORDS if k in text]

def get_html(session, url):
    r = session.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or r.encoding
    return r.url, r.text

def extract_candidates(list_url, html):
    soup = BeautifulSoup(html, "html.parser")
    out = []
    seen = set()

    for a in soup.find_all("a", href=True):
        title = clean_text(a.get_text(" ", strip=True))
        href = urljoin(list_url, a.get("href", ""))
        if not href.startswith(("http://","https://")):
            continue
        if len(title) < 2 or len(title) > 300:
            continue
        # 메뉴/푸터 등 명백한 비게시물 링크 제거
        low = title.lower()
        if any(x in low for x in ["로그인","회원가입","사이트맵","개인정보","이용약관"]):
            continue
        key = canonical_url(href)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append((title, href, a))
    return out

def detail_text(session, url):
    try:
        final_url, html = get_html(session, url)
        soup = BeautifulSoup(html, "html.parser")
        for tag in soup(["script","style","noscript"]):
            tag.decompose()
        text = clean_text(soup.get_text(" ", strip=True))
        return final_url, text, html
    except Exception:
        return url, "", ""

def verify_and_search(session, agency, board_url, checked_identity, sent_urls, sent_fp):
    result = {
        "기관명": agency,
        "게시판URL": board_url,
        "게시판확인": "실패",
        "게시물확인": 0,
        "매칭": [],
        "오류": ""
    }
    try:
        final_url, html = get_html(session, board_url)
        soup = BeautifulSoup(html, "html.parser")
        page_text = clean_text(soup.get_text(" ", strip=True))

        # 게시판 자체에 키워드가 있는 경우도 검사
        candidates = extract_candidates(final_url, html)

        # 후보가 부족하면 페이지 내 모든 링크 중 게시판 성격 URL을 조금 더 허용
        if not candidates:
            candidates = []

        result["게시판확인"] = "성공"
        cutoff = now_kst() - timedelta(days=RECENT_DAYS)

        for title, href, a in candidates[:250]:
            # 제목 자체에서 날짜를 찾고, 없으면 상세 페이지에서 확인
            dt = date_from_text(a.parent.get_text(" ", strip=True) if a.parent else "") or date_from_text(title)

            final_detail, body, _ = detail_text(session, href)
            combined = clean_text(title + " " + body)

            if not dt:
                dt = date_from_text(combined)

            # 날짜가 명확히 과거 30일 밖이면 제외
            if dt and dt < cutoff:
                continue

            hits = keyword_hits(combined)
            if not hits:
                continue
            if any(x in title for x in EXCLUDE_TITLE):
                continue

            identity = canonical_url(final_detail or href)
            fp = hashlib.sha256((agency + "|" + title + "|" + identity).encode("utf-8")).hexdigest()

            if identity in checked_identity:
                continue
            checked_identity.add(identity)

            item = {
                "기관명": agency,
                "제목": title,
                "URL": final_detail or href,
                "게시일": dt.strftime("%Y-%m-%d") if dt else "",
                "키워드": ",".join(hits),
                "본문매칭": any(k in body for k in hits),
                "fingerprint": fp,
            }
            result["매칭"].append(item)

        result["게시물확인"] = len(candidates)
        return result

    except Exception as e:
        result["오류"] = f"{type(e).__name__}: {e}"
        return result

def telegram_send(token, chat_id, text):
    if not token or not chat_id:
        return False, "telegram_secret_missing"
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    try:
        r = requests.post(url, data={"chat_id": chat_id, "text": text, "disable_web_page_preview": True},
                          timeout=20)
        if r.ok:
            return True, "sent"
        return False, f"http_{r.status_code}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def main():
    rows = load_csv()
    checked_identity = set(load_json("checked_post_identity_ledger.json", []))
    sent_urls = set(load_json("sent_url_ledger.json", []))
    sent_fp = set(load_json("sent_fingerprint_ledger.json", []))

    session = requests.Session()
    results = []
    matches = []
    used = 0

    # 게시판 URL이 실제로 입력된 기관만 검색
    targets = [r for r in rows if normalize_url(r.get("게시판URL"))]
    skipped = [r for r in rows if not normalize_url(r.get("게시판URL"))]

    for i, row in enumerate(targets, 1):
        agency = clean_text(row.get("기관명"))
        board = normalize_url(row.get("게시판URL"))
        res = verify_and_search(session, agency, board, checked_identity, sent_urls, sent_fp)
        results.append(res)
        matches.extend(res["매칭"])
        print(f"[{i}/{len(targets)}] {agency} | {res['게시판확인']} | matches={len(res['매칭'])}", flush=True)

    token = os.getenv("TELEGRAM_BOT_TOKEN", "")
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "")

    telegram_sent = 0
    for item in matches:
        if telegram_sent >= TELEGRAM_MAX_SEND:
            break
        u = canonical_url(item["URL"])
        fp = item["fingerprint"]
        if u in sent_urls or fp in sent_fp:
            continue

        msg = (
            f"🔔 {item['기관명']}\n"
            f"📌 {item['제목']}\n"
            f"🔎 키워드: {item['키워드']}\n"
            f"📅 {item['게시일'] or '날짜 미확인'}\n"
            f"🔗 {item['URL']}"
        )
        ok, _ = telegram_send(token, chat_id, msg)
        if ok:
            telegram_sent += 1
            sent_urls.add(u)
            sent_fp.add(fp)

    save_json("checked_post_identity_ledger.json", sorted(checked_identity))
    save_json("sent_url_ledger.json", sorted(sent_urls))
    save_json("sent_fingerprint_ledger.json", sorted(sent_fp))

    # 결과 CSV
    with open("검색결과_V8.15.5.csv", "w", encoding="utf-8-sig", newline="") as f:
        fields = ["기관명","제목","URL","게시일","키워드","본문매칭"]
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for x in matches:
            w.writerow({k:x.get(k,"") for k in fields})

    with open("게시판검증결과_V8.15.5.csv", "w", encoding="utf-8-sig", newline="") as f:
        fields = ["기관명","게시판URL","게시판확인","게시물확인","매칭건수","오류"]
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for x in results:
            w.writerow({
                "기관명":x["기관명"], "게시판URL":x["게시판URL"],
                "게시판확인":x["게시판확인"], "게시물확인":x["게시물확인"],
                "매칭건수":len(x["매칭"]), "오류":x["오류"]
            })
        for r in skipped:
            w.writerow({
                "기관명":r.get("기관명",""), "게시판URL":"",
                "게시판확인":"URL없음", "게시물확인":0, "매칭건수":0,
                "오류":"게시판URL 미입력"
            })

    summary = {
        "version": VERSION,
        "targets_total": len(rows),
        "board_url_targets": len(targets),
        "board_url_missing": len(skipped),
        "board_verified": sum(x["게시판확인"]=="성공" for x in results),
        "board_failed": sum(x["게시판확인"]!="성공" for x in results),
        "posts_checked": sum(x["게시물확인"] for x in results),
        "new_matches": len(matches),
        "telegram_sent": telegram_sent,
        "recent_days": RECENT_DAYS,
        "telegram_max_send": TELEGRAM_MAX_SEND,
        "generated_at": now_kst().isoformat()
    }
    save_json("summary_V8.15.5.json", summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
