
# -*- coding: utf-8 -*-
"""
기관 후보 발굴 프로그램 V1
- 기존 355개 기관 CSV를 기준으로 중복 제거
- ALIO Plus 기관정보를 대량 수집
- ALIO Plus 설문정보를 수집해 '참여정보 발생 가능성' 점수화
- 후보 기관의 홈페이지/기관유형/출처를 CSV로 저장
- 기존 모니터링 프로그램과 직접 결합하지 않고 '승인대기' 단계로 분리

Colab:
  pip install -q requests beautifulsoup4 pandas lxml
  python candidate_discovery.py --input 기관별_상태.csv
"""

import argparse, re, time, json, os
from urllib.parse import urljoin
import requests
import pandas as pd
from bs4 import BeautifulSoup

ALIO_ORG = "https://www.alioplus.go.kr/organization/organByTypeList.do"
ALIO_POLL = "https://www.alioplus.go.kr/nation/pollList.do"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128 Safari/537.36"

KEYWORDS_HIGH = ["설문", "투표", "퀴즈", "이벤트", "공모", "국민참여", "시민참여"]
ORG_SUFFIX = re.compile(r"\s*(주식회사|\(주\)|재단법인|사단법인|\(재\)|\(사\)|법인)\s*")

def norm(s):
    if pd.isna(s): return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", "", s)
    s = s.replace("㈜","(주)").replace("ㆍ","·")
    return s

def get(session, url, params=None, timeout=20):
    for attempt in range(3):
        try:
            r = session.get(url, params=params, timeout=timeout)
            r.raise_for_status()
            r.encoding = r.apparent_encoding or r.encoding
            return r
        except Exception:
            if attempt == 2: return None
            time.sleep(1.2 * (attempt + 1))
    return None

def parse_org_page(html, source_url):
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n", strip=True)
    rows = []

    # ALIO Plus 기관정보 화면은 기관명 -> 주소 -> 기관유형 순서가 반복된다.
    # DOM 구조 변경에 대비해 주소/유형을 앵커로 주변 텍스트를 함께 분석한다.
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    type_pat = re.compile(r"(공기업|준정부기관|기타공공기관|출연기관|출자기관)")
    region_pat = re.compile(r"(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)")
    addr_pat = re.compile(r"(특별시|광역시|특별자치시|특별자치도|도)\s")

    for i, line in enumerate(lines):
        if not type_pat.search(line):
            continue
        typ = line
        # 바로 앞 1~5개 라인 중 기관명 후보 선택
        for j in range(max(0, i-5), i):
            cand = lines[j]
            if len(cand) < 2 or len(cand) > 80:
                continue
            if region_pat.match(cand) or addr_pat.search(cand):
                continue
            if cand in ("기관정보", "유형별", "통합검색"):
                continue
            # 주소가 기관명으로 잡히는 것을 제외
            if re.search(r"(로|길|번길|대로|읍|면|동)\s*\d", cand):
                continue
            rows.append({"기관명": cand, "기관유형": typ, "후보출처": "ALIO_PLUS_기관정보", "출처URL": source_url})
            break

    # 중복 제거
    out, seen = [], set()
    for x in rows:
        k = (norm(x["기관명"]), x["기관유형"])
        if k not in seen:
            seen.add(k); out.append(x)
    return out

def parse_poll_page(html, source_url):
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n", strip=True)
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    out = []
    # 날짜/기관명 패턴을 이용한 보수적 추출
    date_pat = re.compile(r"^\d{4}-\d{2}-\d{2}")
    for i, line in enumerate(lines):
        if not date_pat.match(line):
            continue
        # 날짜 이전의 제목/기관명 블록에서 기관명 후보
        block = lines[max(0,i-4):i+1]
        # 기관명은 보통 날짜 바로 앞 또는 제목 뒤에 위치
        for cand in reversed(block[:-1]):
            if len(cand) >= 2 and len(cand) <= 80 and not date_pat.match(cand):
                if any(x in cand for x in ["공사","공단","재단","진흥원","공기업","공공","원","센터","협회","연구원","공사","공단"]):
                    out.append(cand)
                    break
    return out

def discover(input_csv, out_csv, max_org_pages=160, max_poll_pages=30, sleep=0.15):
    base = pd.read_csv(input_csv, encoding="utf-8-sig")
    if "기관명" not in base.columns:
        raise ValueError("입력 CSV에 '기관명' 열이 없습니다.")

    existing = {}
    for _, r in base.iterrows():
        existing[norm(r["기관명"])] = {
            "기관명": str(r["기관명"]),
            "기관유형": str(r.get("기관유형","")),
            "홈페이지URL": str(r.get("홈페이지URL","")),
            "기존355": "Y",
        }

    s = requests.Session()
    s.headers.update({"User-Agent": UA})

    candidates = {}
    # 1) ALIO Plus 기관정보
    for p in range(1, max_org_pages+1):
        url = ALIO_ORG
        r = get(s, url, params={"pageNo": p})
        if not r:
            continue
        rows = parse_org_page(r.text, r.url)
        if not rows:
            # 페이지가 끝났을 가능성
            if p > 10:
                break
        for x in rows:
            k = norm(x["기관명"])
            if not k or k in existing:
                continue
            candidates.setdefault(k, {
                **x, "기존355":"N", "설문/참여건수":0, "참여정보키워드":0,
                "참여정보점수":0, "후보점수":0
            })
        time.sleep(sleep)

    # 2) ALIO Plus 설문정보: 기관명/키워드 발생을 별도 점수로 누적
    poll_counts = {}
    poll_kw = {}
    for p in range(1, max_poll_pages+1):
        r = get(s, ALIO_POLL, params={"pageNo": p})
        if not r: continue
        soup = BeautifulSoup(r.text, "lxml")
        text = soup.get_text(" ", strip=True)
        # 알려진 설문 목록에서 기관명은 날짜 앞/뒤에 노출되므로 후보 이름 토큰을 추정
        for phrase in re.findall(r"([가-힣A-Za-z0-9·()㈜&\-\s]{2,80})\s+(?:\d{4}-\d{2}-\d{2})", text):
            phrase = re.sub(r"\s+"," ",phrase).strip()
            # 가장 오른쪽의 기관명다운 구간을 찾는다
            parts = [x.strip() for x in phrase.split("  ") if x.strip()]
            for part in reversed(parts):
                if any(t in part for t in ["공사","공단","재단","진흥원","원","센터","협회","공공"]):
                    k = norm(part)
                    poll_counts[k] = poll_counts.get(k,0)+1
                    break
        for kw in KEYWORDS_HIGH:
            if kw in text:
                # 페이지 단위 보조 신호
                for k in list(candidates.keys()):
                    if k in norm(text):
                        poll_kw[k] = poll_kw.get(k,0)+1
        time.sleep(sleep)

    # 3) 점수화
    for k, x in candidates.items():
        pc = poll_counts.get(k,0)
        kw = poll_kw.get(k,0)
        x["설문/참여건수"] = pc
        x["참여정보키워드"] = kw
        # 공식 기관정보 + 참여정보 발생 신호를 중심으로 점수
        score = 30
        if pc > 0: score += min(40, pc*10)
        if kw > 0: score += min(15, kw*5)
        name = x["기관명"]
        if any(t in name for t in ["관광","문화","체육","복지","환경","청년","교육","연구","진흥","콘텐츠","소비자","안전","농업","수산"]):
            score += 10
        x["후보점수"] = min(100, score)
        x["추천등급"] = "A" if score >= 70 else ("B" if score >= 50 else "C")

    df = pd.DataFrame(candidates.values())
    if df.empty:
        df = pd.DataFrame(columns=["기관명","기관유형","홈페이지URL","기존355","후보출처","출처URL","설문/참여건수","참여정보키워드","참여정보점수","후보점수","추천등급"])
    else:
        df = df.sort_values(["후보점수","설문/참여건수","기관명"], ascending=[False,False,True])
    df.to_csv(out_csv, index=False, encoding="utf-8-sig")

    summary = {
        "기준기관": len(base),
        "신규후보": len(df),
        "A등급": int((df["추천등급"]=="A").sum()) if len(df) else 0,
        "B등급": int((df["추천등급"]=="B").sum()) if len(df) else 0,
        "C등급": int((df["추천등급"]=="C").sum()) if len(df) else 0,
        "source": ["ALIO Plus 기관정보", "ALIO Plus 설문정보"],
    }
    with open(out_csv.replace(".csv","_요약.json"),"w",encoding="utf-8") as f:
        json.dump(summary,f,ensure_ascii=False,indent=2)
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    print("생성:", out_csv)

if __name__ == "__main__":
    ap=argparse.ArgumentParser()
    ap.add_argument("--input", default="기관별_상태.csv")
    ap.add_argument("--output", default="기관후보_발굴결과.csv")
    ap.add_argument("--org-pages", type=int, default=160)
    ap.add_argument("--poll-pages", type=int, default=30)
    args=ap.parse_args()
    discover(args.input,args.output,args.org_pages,args.poll_pages)


# V1.1 output verification
from pathlib import Path as _Path
_csv_out = _Path("기관후보_발굴결과.csv")
_summary_out = _Path("기관후보_발굴결과_요약.json")
print(f"[V1.1] CSV 확인: {_csv_out} / 존재={_csv_out.exists()} / 크기={_csv_out.stat().st_size if _csv_out.exists() else 0} bytes")
print(f"[V1.1] 요약 JSON 확인: {_summary_out} / 존재={_summary_out.exists()} / 크기={_summary_out.stat().st_size if _summary_out.exists() else 0} bytes")
if not _csv_out.exists():
    raise SystemExit("기관후보_발굴결과.csv가 생성되지 않았습니다.")
