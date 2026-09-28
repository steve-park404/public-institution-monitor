# -*- coding: utf-8 -*-
"""
기관 후보 발굴 V1.2

목표
- ALIO Plus에서 '실제 공공기관' 후보만 보수적으로 추출
- 기관 홈페이지/서비스/SNS/포털/통계값/기관유형 문자열 등의 오탐 제거
- 지부/지역본부/사업소/교육원 등 부속조직은 ALIO가 부모기관을 표시할 경우 부모기관으로 통합
- 기존 355기관과 중복 제거
- ALIO Plus 설문정보는 '후보 생성'이 아니라 '참여기회 신호' 점수화에만 사용
- 후보 결과를 GitHub Actions에서 CSV/JSON으로 자동 보존

V1.2 핵심 변경
1. 기관정보 파싱을 단순 '앞 5줄' 방식에서 DOM 컨테이너 + 주소 + 기관유형 교차검증 방식으로 변경
2. 기관유형은 허용된 5개 유형만 인정
3. 서비스/SNS/포털/숫자/유형명 오탐 차단
4. 부모기관 표기가 있는 부속기관은 부모기관으로 canonicalize
5. 기관명 정규화/부분문자열 중복 제거
6. 설문정보는 기존 후보와 매칭되는 경우에만 점수에 반영
7. 검증통계를 별도 JSON에 기록
"""

import argparse, json, os, re, time
from urllib.parse import urljoin
import requests
import pandas as pd
from bs4 import BeautifulSoup, NavigableString

VERSION = "V1.2"
ALIO_ORG = "https://www.alioplus.go.kr/organization/organByTypeList.do"
ALIO_REGION = "https://www.alioplus.go.kr/organization/organByRegionList.do"
ALIO_POLL = "https://www.alioplus.go.kr/nation/pollList.do"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128 Safari/537.36"

ALLOWED_TYPES = {
    "공기업(시장형)", "공기업(준시장형)",
    "준정부기관(기금관리형)", "준정부기관(위탁집행형)",
    "기타공공기관"
}
TYPE_RE = re.compile(r"^(공기업\((?:시장형|준시장형)\)|준정부기관\((?:기금관리형|위탁집행형)\)|기타공공기관)$")
ADDRESS_RE = re.compile(r"(?:특별시|광역시|특별자치시|특별자치도|도)\s+[^\n]{1,80}(?:로|길|대로|번길|읍|면|동|가|리)\s*\d")

# 실제 기관명에 나타날 수 있으나 서비스/사이트명에 가까운 표현은 강하게 배제
SERVICE_PATTERNS = [
    "기관 홈페이지", "인스타그램", "페이스북", "유튜브", "블로그",
    "Q-net", "Q net", "KIPRO", "TAAS", "공공누리",
    "통합접수시스템", "전자신고포털", "정보포털", "정보시스템",
    "통합전산망", "온라인시스템", "전자민원", "서비스", "홈페이지",
    "포털", "시스템", "앱", "어플", "플랫폼", "정보광장",
    "지식정보", "종합정보서비스", "정보서비스", "신청", "피해구제",
    "클린사업", "소재은행", "고용산재보험토탈서비스", "근로복지서비스",
]
GENERIC_NAMES = set(ALLOWED_TYPES) | {"기관정보", "유형별", "지역별", "통합검색", "기관", "본/지점"}

PARTICIPATION_KEYWORDS = [
    "설문", "설문조사", "투표", "퀴즈", "이벤트", "국민참여", "시민참여",
    "의견수렴", "만족도", "조사", "참여자", "모집", "패널", "체험단"
]
REWARD_KEYWORDS = ["경품", "상품권", "상금", "기프티콘", "사은품", "포인트", "보상", "응모"]


def norm(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip()
    s = s.replace("㈜", "(주)").replace("ㆍ", "·")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def keynorm(s):
    return re.sub(r"[^0-9a-z가-힣]", "", norm(s).lower())


def get(session, url, params=None, timeout=20):
    for attempt in range(3):
        try:
            r = session.get(url, params=params, timeout=timeout)
            r.raise_for_status()
            r.encoding = r.apparent_encoding or r.encoding
            return r
        except Exception:
            if attempt == 2:
                return None
            time.sleep(1.0 + attempt * 1.5)
    return None


def clean_candidate_name(name):
    name = norm(name)
    if not name:
        return ""
    # 기관유형/기타 메타가 이름 뒤에 붙는 경우 제거
    for t in sorted(ALLOWED_TYPES, key=len, reverse=True):
        name = re.sub(rf"\s*/\s*{re.escape(t)}$", "", name)
    name = re.sub(r"\s*/\s*(?:공기업|준정부기관|기타공공기관).*$", "", name)
    name = re.sub(r"^\d{1,6}$", "", name).strip()
    return name


def is_bad_name(name):
    n = clean_candidate_name(name)
    if not n or n in GENERIC_NAMES:
        return True
    if len(n) < 2 or len(n) > 80:
        return True
    if re.fullmatch(r"[0-9\s.,:/-]+", n):
        return True
    if any(p.lower() in n.lower() for p in SERVICE_PATTERNS):
        return True
    # 공공기관 유형 자체가 기관명으로 추출된 경우
    if TYPE_RE.fullmatch(n):
        return True
    # 단순 메뉴명/버튼명
    if n in {"검색하기", "엑셀다운로드", "상세보기", "더보기", "로그인", "회원가입"}:
        return True
    return False


def type_info(raw):
    raw = norm(raw)
    if TYPE_RE.fullmatch(raw):
        return raw, ""
    # ALIO Plus가 부모기관 / 기관유형 형태로 표시하는 경우
    m = re.match(r"^(.+?)\s*/\s*(공기업\((?:시장형|준시장형)\)|준정부기관\((?:기금관리형|위탁집행형)\)|기타공공기관)$", raw)
    if m:
        return m.group(2), clean_candidate_name(m.group(1))
    return "", ""


def likely_name_from_container(container, type_text):
    """기관 카드/목록 컨테이너에서 서비스 링크가 아니라 대표 기관명을 우선 추출."""
    # 1) heading/strong/dt 요소 우선
    preferred = []
    for tag in container.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "dt", "strong", "b"], limit=20):
        t = clean_candidate_name(tag.get_text(" ", strip=True))
        if t and not is_bad_name(t):
            preferred.append(t)
    if preferred:
        return preferred[0]

    # 2) 링크 중 기관명다운 텍스트. 단, 하위 서비스 링크는 배제
    for a in container.find_all("a", href=True, limit=30):
        t = clean_candidate_name(a.get_text(" ", strip=True))
        if t and not is_bad_name(t):
            # 너무 긴 문장/URL 링크는 제외
            if len(t) <= 80 and not re.search(r"https?://", t):
                return t

    # 3) direct text nodes를 역순/순서대로 확인
    for node in container.find_all(string=True):
        if not isinstance(node, NavigableString):
            continue
        t = clean_candidate_name(str(node))
        if t and not is_bad_name(t) and not ADDRESS_RE.search(t):
            return t
    return ""


def find_card_ancestor(node):
    """기관유형 텍스트에서 주소까지 포함하는 가장 작은 실질 컨테이너를 찾는다."""
    cur = node.parent
    best = None
    for _ in range(10):
        if cur is None:
            break
        txt = norm(cur.get_text(" ", strip=True))
        if ADDRESS_RE.search(txt) and len(txt) <= 1800:
            best = cur
            # 너무 큰 페이지 전체 컨테이너로 올라가기 전에 중단
            if cur.name in {"li", "article", "dl", "dd", "section"}:
                break
        cur = cur.parent
    return best


def parse_org_page(html, source_url):
    soup = BeautifulSoup(html, "lxml")
    rows = []
    seen = set()

    # 정확한 기관유형 텍스트 노드만 시작점으로 사용
    for node in soup.find_all(string=True):
        raw = norm(str(node))
        typ, parent_org = type_info(raw)
        if not typ:
            continue
        card = find_card_ancestor(node)
        if card is None:
            continue
        card_text = norm(card.get_text(" ", strip=True))
        if not ADDRESS_RE.search(card_text):
            continue

        name = likely_name_from_container(card, typ)
        if not name:
            continue

        # 부모기관 표기가 있으면 본/지점명이 아니라 부모기관을 후보명으로 사용
        if parent_org and not is_bad_name(parent_org):
            canonical = parent_org
        else:
            canonical = name

        canonical = clean_candidate_name(canonical)
        if is_bad_name(canonical):
            continue

        k = keynorm(canonical)
        if not k or k in seen:
            continue
        seen.add(k)
        rows.append({
            "기관명": canonical,
            "기관유형": typ,
            "후보출처": "ALIO_PLUS_기관정보",
            "출처URL": source_url,
            "검증근거": "기관유형+주소+DOM기관카드"
        })
    return rows


def parse_poll_page(html):
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n", strip=True)
    lines = [norm(x) for x in text.splitlines() if norm(x)]
    records = []
    date_re = re.compile(r"^20\d{2}-\d{2}-\d{2}")

    # 날짜가 있는 항목 주변에서 기관명 후보를 수집하되, 후보 생성에는 사용하지 않는다.
    for i, line in enumerate(lines):
        if not date_re.match(line):
            continue
        block = lines[max(0, i-8):min(len(lines), i+4)]
        for x in block:
            if is_bad_name(x) or date_re.match(x):
                continue
            if any(t in x for t in ["공사", "공단", "재단", "진흥원", "연구원", "공공", "원", "센터", "협회", "대학교", "병원"]):
                records.append(x)
    return records


def load_existing(path):
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    if path.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path, encoding="utf-8-sig")
    if "기관명" not in df.columns:
        raise ValueError("입력 파일에 '기관명' 열이 없습니다.")
    return df


def build_existing(df):
    out = {}
    for _, r in df.iterrows():
        n = clean_candidate_name(r.get("기관명", ""))
        if not n:
            continue
        out[keynorm(n)] = n
    return out


def canonicalize_candidates(rows):
    merged = {}
    for x in rows:
        name = clean_candidate_name(x.get("기관명", ""))
        if is_bad_name(name):
            continue
        k = keynorm(name)
        if not k:
            continue
        if k not in merged:
            merged[k] = dict(x)
        else:
            # 더 구체적인 기관유형/근거가 있으면 유지
            if not merged[k].get("기관유형") and x.get("기관유형"):
                merged[k]["기관유형"] = x["기관유형"]
    return list(merged.values())


def match_poll_signals(candidates, poll_names):
    candidate_keys = {keynorm(x["기관명"]): x for x in candidates}
    counts = {k: 0 for k in candidate_keys}
    kw_hits = {k: 0 for k in candidate_keys}
    for raw in poll_names:
        rk = keynorm(raw)
        if not rk:
            continue
        # 정확/포함 매칭. 짧은 기관명은 오탐 방지를 위해 4자 이상만 부분매칭
        matched = []
        for ck in candidate_keys:
            if rk == ck or (len(ck) >= 4 and (ck in rk or rk in ck)):
                matched.append(ck)
        for ck in matched:
            counts[ck] += 1
            kw_hits[ck] += 1
    return counts, kw_hits


def score_candidate(x, poll_count, kw_hits):
    name = x["기관명"]
    score = 40  # 실제 기관 검증 통과 기본점수
    if poll_count:
        score += min(35, poll_count * 8)
    if kw_hits:
        score += min(15, kw_hits * 3)
    if any(t in name for t in ["관광", "문화", "체육", "복지", "환경", "청년", "교육", "연구", "진흥", "콘텐츠", "소비자", "안전", "농업", "수산"]):
        score += 10
    score = min(100, score)
    grade = "A" if score >= 75 else ("B" if score >= 55 else "C")
    return score, grade


def discover(input_path, out_csv, max_org_pages=160, max_poll_pages=30, sleep=0.15):
    base = load_existing(input_path)
    existing = build_existing(base)
    session = requests.Session()
    session.headers.update({"User-Agent": UA})

    raw_rows = []
    page_stats = []

    # 기관정보: V1과 동일한 범위로 충분히 순회하되, 파싱은 V1.2에서 보수적으로 수행
    for p in range(1, max_org_pages + 1):
        r = get(session, ALIO_ORG, params={"pageNo": p})
        if not r:
            page_stats.append({"page": p, "status": "FETCH_ERROR", "rows": 0})
            continue
        rows = parse_org_page(r.text, r.url)
        page_stats.append({"page": p, "status": "OK", "rows": len(rows)})
        raw_rows.extend(rows)
        if not rows and p > 20:
            # 연속 빈 페이지가 끝을 의미하는 경우 조기 종료
            empty_recent = sum(1 for z in page_stats[-5:] if z["rows"] == 0)
            if empty_recent >= 5:
                break
        time.sleep(sleep)

    all_candidates = canonicalize_candidates(raw_rows)
    before_filter = len(all_candidates)
    candidates = [x for x in all_candidates if keynorm(x["기관명"]) not in existing]

    # 설문정보는 후보를 생성하지 않고, 이미 검증된 후보에 대한 보조점수만 제공
    poll_names = []
    poll_pages_ok = 0
    for p in range(1, max_poll_pages + 1):
        r = get(session, ALIO_POLL, params={"pageNo": p})
        if not r:
            continue
        poll_pages_ok += 1
        poll_names.extend(parse_poll_page(r.text))
        time.sleep(sleep)

    poll_counts, kw_hits = match_poll_signals(candidates, poll_names)

    result = []
    for x in candidates:
        k = keynorm(x["기관명"])
        pc = poll_counts.get(k, 0)
        kw = kw_hits.get(k, 0)
        score, grade = score_candidate(x, pc, kw)
        y = dict(x)
        y.update({
            "기존355": "N",
            "설문/참여건수": pc,
            "참여정보키워드": kw,
            "참여정보점수": min(50, pc * 8 + kw * 3),
            "후보점수": score,
            "추천등급": grade,
        })
        result.append(y)

    cols = ["기관명", "기관유형", "후보출처", "출처URL", "검증근거", "기존355",
            "설문/참여건수", "참여정보키워드", "참여정보점수", "후보점수", "추천등급"]
    df = pd.DataFrame(result, columns=cols)
    if not df.empty:
        df = df.sort_values(["후보점수", "설문/참여건수", "기관명"], ascending=[False, False, True])
    df.to_csv(out_csv, index=False, encoding="utf-8-sig")

    summary = {
        "version": VERSION,
        "기준기관": len(base),
        "원시기관카드": len(raw_rows),
        "기관카드정제후": before_filter,
        "기존355제거후신규후보": len(df),
        "신규후보": len(df),
        "A등급": int((df["추천등급"] == "A").sum()) if len(df) else 0,
        "B등급": int((df["추천등급"] == "B").sum()) if len(df) else 0,
        "C등급": int((df["추천등급"] == "C").sum()) if len(df) else 0,
        "설문페이지정상조회": poll_pages_ok,
        "설문기관명신호수": len(poll_names),
        "source": ["ALIO Plus 기관정보", "ALIO Plus 설문정보"],
        "정책": {
            "실제기관검증": True,
            "서비스_SNS_포털_유형문자열제외": True,
            "부속조직부모기관통합": True,
            "설문정보는후보생성에사용하지않음": True
        }
    }
    summary_path = out_csv.replace(".csv", "_요약.json")
    diag_path = out_csv.replace(".csv", "_진단.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    diagnostics = {
        "version": VERSION,
        "page_stats": page_stats,
        "raw_rows": len(raw_rows),
        "clean_rows": before_filter,
        "new_rows": len(df),
        "excluded_existing": len(all_candidates) - len(candidates),
        "poll_pages_ok": poll_pages_ok,
        "poll_name_signals": len(poll_names),
    }
    with open(diag_path, "w", encoding="utf-8") as f:
        json.dump(diagnostics, f, ensure_ascii=False, indent=2)

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"생성: {out_csv}")
    print(f"생성: {summary_path}")
    print(f"생성: {diag_path}")

    # 최소 품질검증: 후보 중 명백한 오탐 패턴이 남아 있으면 실패
    bad_left = []
    for n in df["기관명"].tolist() if len(df) else []:
        if is_bad_name(n):
            bad_left.append(n)
    if bad_left:
        raise SystemExit(f"V1.2 품질검증 실패: 명백한 비기관 후보 {bad_left[:10]}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="기관별_상태.csv")
    ap.add_argument("--output", default="기관후보_발굴결과.csv")
    ap.add_argument("--org-pages", type=int, default=160)
    ap.add_argument("--poll-pages", type=int, default=30)
    ap.add_argument("--sleep", type=float, default=0.15)
    args = ap.parse_args()
    discover(args.input, args.output, args.org_pages, args.poll_pages, args.sleep)
