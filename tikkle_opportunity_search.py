
# 티끌 알림봇 V2.0 - 기회 중심 웹 검색 엔진
# 목적: 기관 목록을 늘리는 것이 아니라, 실제 참여 가능한 보상형 기회를 직접 찾는다.

import os, re, json, time, hashlib
from datetime import datetime, timedelta, timezone
from urllib.parse import urlparse
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

VERSION = "V2.0"
KST = timezone(timedelta(hours=9))

SEARCH_QUERIES = [
    '"설문조사" "기프티콘" 모집',
    '"설문조사" "상품권" 모집',
    '"설문조사" "커피쿠폰" 참여',
    '"설문조사" "참여혜택"',
    '"설문조사" "참여자 모집"',
    '"온라인 설문" "기프티콘"',
    '"설문" "참여수당"',
    '"패널 모집" "사례비"',
    '"패널 모집" "상품권"',
    '"패널 모집" "기프티콘"',
    '"인터뷰 참여자 모집" "사례비"',
    '"인터뷰 참여자" "상품권"',
    '"FGI 모집" "사례비"',
    '"좌담회 모집" "사례비"',
    '"사용성 테스트" "사례비"',
    '"사용성 테스트" "상품권"',
    '"제품 테스트" "참여자 모집"',
    '"서비스 테스트" "참여자 모집"',
    '"베타테스터" 모집 혜택',
    '"모니터단" 모집 "상품권"',
    '"시민참여" "상품권"',
    '"국민참여" "기프티콘"',
    '"의견수렴" "상품권"',
    '"정책 설문" "기프티콘"',
    '"만족도 조사" "경품"',
    '"소비자 조사" "사례비"',
    '"소비자 패널" 모집',
    '"리서치" "참여자 모집" "사례비"',
    '"설문 이벤트" "경품"'
]

EXCLUDE_TERMS = [
    "공모전", "채용", "입사지원", "구인", "구직", "취업", "장학생",
    "장학금", "보도자료", "정책자료", "입찰", "계약", "제안서",
]

REWARD_PATTERNS = [
    (r'(\d[\d,]*)\s*원', 25),
    (r'(\d[\d,]*)\s*만원', 30),
    (r'상품권', 18),
    (r'기프티콘', 18),
    (r'커피쿠폰|커피\s*쿠폰|스타벅스', 16),
    (r'사례비|참여비|참여수당', 25),
    (r'경품|추첨', 10),
    (r'페이|포인트', 12),
]

OPPORTUNITY_PATTERNS = [
    (r'설문조사|설문', 12),
    (r'패널', 14),
    (r'인터뷰', 18),
    (r'FGI|좌담회', 20),
    (r'사용성\s*테스트|UX\s*테스트', 22),
    (r'제품\s*테스트|서비스\s*테스트', 18),
    (r'베타테스터|베타\s*테스트', 15),
    (r'모니터단|모니터링단', 12),
    (r'소비자\s*조사|리서치', 12),
    (r'의견수렴|시민참여|국민참여', 10),
]

RESTRICTION_PATTERNS = [
    (r'만\s*(?:가능|대상)|(?:대학생|대학원생|교직원|의료인|공무원|청소년|임산부|환자|사업자)', -7),
    (r'특정\s*지역|해당\s*지역', -5),
]

def norm(s):
    return re.sub(r'\s+', ' ', (s or '')).strip()

def canonical_url(url):
    try:
        p = urlparse(url)
        host = p.netloc.lower().replace("www.", "")
        path = re.sub(r'/+$', '', p.path)
        return f"{host}{path}"
    except:
        return url

def sha(s):
    return hashlib.sha256(s.encode("utf-8", "ignore")).hexdigest()[:20]

def fetch_page(url):
    try:
        r = requests.get(
            url,
            timeout=12,
            headers={"User-Agent": "Mozilla/5.0 (compatible; TikkleOpportunityBot/2.0)"}
        )
        if r.status_code >= 400:
            return ""
        r.encoding = r.apparent_encoding or r.encoding
        soup = BeautifulSoup(r.text, "lxml")
        for x in soup(["script","style","noscript","svg"]):
            x.decompose()
        return norm(soup.get_text(" ", strip=True))[:12000]
    except Exception:
        return ""

def extract_reward(text):
    hits = []
    for pat, pts in REWARD_PATTERNS:
        m = re.search(pat, text, re.I)
        if m:
            hits.append(m.group(0))
    return ", ".join(dict.fromkeys(hits))[:200]

def extract_deadline(text):
    patterns = [
        r'(?:마감|접수|신청|모집)[^0-9]{0,20}(20\d{2}[.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2})',
        r'(20\d{2}[.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2})[^가-힣]{0,10}(?:까지|마감)',
        r'(\d{1,2}월\s*\d{1,2}일)[^가-힣]{0,10}(?:까지|마감)',
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(1)
    return ""

def score_item(title, body):
    text = f"{title} {body}"
    score = 0
    reasons = []
    for pat, pts in OPPORTUNITY_PATTERNS:
        if re.search(pat, text, re.I):
            score += pts
            reasons.append(pat.replace("\\s*", " "))
    for pat, pts in REWARD_PATTERNS:
        if re.search(pat, text, re.I):
            score += pts
            reasons.append("보상")
    for pat, pts in RESTRICTION_PATTERNS:
        if re.search(pat, text, re.I):
            score += pts
    if extract_deadline(text):
        score += 5
    if any(x in title for x in EXCLUDE_TERMS):
        score -= 60
    return max(0, min(100, score)), list(dict.fromkeys(reasons))

def search_web():
    results = []
    with DDGS(timeout=15) as ddgs:
        for q in SEARCH_QUERIES:
            try:
                rows = ddgs.text(q, max_results=6, safesearch="moderate")
                for x in rows:
                    title = norm(x.get("title"))
                    url = x.get("href") or x.get("url") or ""
                    snippet = norm(x.get("body") or x.get("snippet"))
                    if not title or not url:
                        continue
                    results.append({
                        "query": q, "title": title, "url": url, "snippet": snippet
                    })
            except Exception:
                continue
    return results

def build_candidates(raw):
    by_url = {}
    for x in raw:
        key = canonical_url(x["url"])
        if key not in by_url:
            by_url[key] = x
        else:
            # 같은 URL이 여러 검색어에 걸리면 검색어를 누적
            by_url[key]["query"] += " | " + x["query"]
    return list(by_url.values())

def enrich(items):
    out = []
    for i, x in enumerate(items):
        body = fetch_page(x["url"])
        combined = norm(x["title"] + " " + x["snippet"] + " " + body)
        score, reasons = score_item(x["title"], combined)
        if score < 25:
            continue
        if any(t in x["title"] for t in EXCLUDE_TERMS):
            continue
        out.append({
            "발견일": datetime.now(KST).strftime("%Y-%m-%d"),
            "점수": score,
            "제목": x["title"],
            "보상정보": extract_reward(combined),
            "마감정보": extract_deadline(combined),
            "검색근거": x["query"][:300],
            "URL": x["url"],
            "도메인": urlparse(x["url"]).netloc,
            "판정근거": ", ".join(reasons[:8]),
            "미리보기": norm(body or x["snippet"])[:500],
        })
    return out

def load_state(path="opportunity_state.json"):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except:
        return {"sent": {}, "seen": {}}

def save_state(state, path="opportunity_state.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def send_telegram(items):
    token = os.getenv("TELEGRAM_BOT_TOKEN", "")
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "")
    if not token or not chat_id:
        return 0
    sent = 0
    for x in sorted(items, key=lambda z: z["점수"], reverse=True)[:20]:
        msg = (
            f"🎯 티끌 기회 발견 [{x['점수']}점]\n\n"
            f"📌 {x['제목']}\n"
            f"🎁 {x['보상정보'] or '보상 확인 필요'}\n"
            f"📅 {x['마감정보'] or '마감일 확인 필요'}\n"
            f"🌐 {x['도메인']}\n\n"
            f"🔗 {x['URL']}"
        )
        try:
            r = requests.post(
                f"https://api.telegram.org/bot{token}/sendMessage",
                json={"chat_id": chat_id, "text": msg},
                timeout=10
            )
            if r.ok:
                sent += 1
        except Exception:
            pass
    return sent

def main():
    print(f"=== 티끌 알림봇 {VERSION} ===")
    print("기관 중심이 아닌 참여기회 중심 웹 검색을 시작합니다.")
    raw = search_web()
    print(f"검색 결과: {len(raw)}")
    candidates = build_candidates(raw)
    print(f"URL 중복 제거 후: {len(candidates)}")
    items = enrich(candidates)
    items.sort(key=lambda x: x["점수"], reverse=True)

    state = load_state()
    new_items = []
    for x in items:
        fp = sha(canonical_url(x["URL"]) + "|" + x["제목"])
        if fp not in state["seen"]:
            state["seen"][fp] = {
                "title": x["제목"], "url": x["URL"],
                "first_seen": x["발견일"], "score": x["점수"]
            }
            new_items.append(x)

    # 상위 후보만 Telegram 전송
    sent = send_telegram(new_items)
    for x in new_items[:20]:
        state["sent"][sha(canonical_url(x["URL"]) + "|" + x["제목"])] = datetime.now(KST).isoformat()

    save_state(state)

    import csv
    fields = ["발견일","점수","제목","보상정보","마감정보","검색근거","URL","도메인","판정근거","미리보기"]
    with open("티끌_기회_검색결과.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(items)

    summary = {
        "version": VERSION,
        "검색어수": len(SEARCH_QUERIES),
        "검색원시결과": len(raw),
        "URL중복제거후": len(candidates),
        "유효후보": len(items),
        "신규후보": len(new_items),
        "텔레그램전송": sent,
        "상위점수": [x["점수"] for x in items[:10]],
        "실행시각": datetime.now(KST).isoformat(),
        "운영원칙": {
            "기관중심아님": True,
            "참여기회중심": True,
            "공모전_채용_입찰_제외": True,
            "중복URL제거": True,
            "보상정보추출": True,
            "마감정보추출": True
        }
    }
    with open("티끌_기회_검색결과_요약.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(json.dumps(summary, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
