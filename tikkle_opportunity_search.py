
import os, re, json, hashlib
from datetime import datetime, timedelta, timezone, date
from urllib.parse import urlparse
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

VERSION="V2.3"
KST=timezone(timedelta(hours=9))
TODAY=datetime.now(KST).date()

SEARCH_QUERIES=[
    '"설문조사" "기프티콘" "모집"',
    '"설문조사" "상품권" "모집"',
    '"설문조사" "커피쿠폰" "참여자"',
    '"설문조사" "참여자 모집"',
    '"온라인 설문" "기프티콘"',
    '"설문" "참여수당"',
    '"패널 모집" "사례비"',
    '"패널 모집" "상품권"',
    '"인터뷰 참여자 모집" "사례비"',
    '"인터뷰 참여자 모집" "상품권"',
    '"사용성 테스트" "사례비"',
    '"사용성 테스트" "모집"',
    '"제품 테스트" "참여자 모집"',
    '"서비스 테스트" "참여자 모집"',
    '"베타테스터" "모집"',
    '"모니터단" "모집" "상품권"',
    '"시민참여" "기프티콘" "모집"',
    '"국민참여" "기프티콘" "모집"',
    '"의견수렴" "상품권" "참여자"',
    '"정책 설문" "기프티콘"',
    '"만족도 조사" "경품" "참여"',
    '"소비자 조사" "사례비"',
    '"소비자 패널" "모집"',
    '"리서치" "참여자 모집" "사례비"',
]

# 사용자가 참여할 수 없는 유형은 검색/판정 단계 모두 제외
EXCLUDE_TERMS=[
    "공모전","채용","입사지원","구인","구직","취업","장학생","장학금",
    "입찰","계약","제안서","fgI","fgi","좌담회","집단면접","포커스그룹",
    "focus group","집단 인터뷰"
]
INFO_TERMS=[
    "가이드","완벽 가이드","방법","후기","리뷰","정리","소개","개념",
    "팁","노하우","칼럼","대행","에이전시","플랫폼","서비스 소개",
    "추천","사례","사용기"
]
RECRUIT_TERMS=[
    "참여자 모집","참가자 모집","모집공고","모집합니다","모집 안내",
    "설문 참여","설문조사 참여","패널 모집","인터뷰 참여자",
    "사용성 테스트 참가","사용성 테스트 참여","제품 테스트 참여",
    "서비스 테스트 참여","베타테스터 모집","모니터단 모집",
    "리서치 참여자","소비자 조사 참여","현재 모집","상시 모집"
]
REWARD_PATTERNS=[
    (r'(\d[\d,]*)\s*만원',30),
    (r'(\d[\d,]*)\s*원',25),
    (r'상품권',18),(r'기프티콘',18),(r'커피\s*쿠폰|스타벅스',16),
    (r'사례비|참여비|참여수당',25),(r'경품|추첨',10),(r'페이|포인트',12)
]

def norm(s): return re.sub(r'\s+',' ',(s or '')).strip()

def canonical_url(url):
    try:
        p=urlparse(url); return p.netloc.lower().replace("www.","")+re.sub(r'/+$','',p.path)
    except: return url

def fetch_page(url):
    try:
        r=requests.get(url,timeout=12,headers={"User-Agent":"Mozilla/5.0 (compatible; TikkleOpportunityBot/2.3)"})
        if r.status_code>=400: return ""
        r.encoding=r.apparent_encoding or r.encoding
        soup=BeautifulSoup(r.text,"lxml")
        for x in soup(["script","style","noscript","svg"]): x.decompose()
        return norm(soup.get_text(" ",strip=True))[:16000]
    except Exception: return ""

def parse_date(s):
    s=s.replace(" ","")
    m=re.search(r'(20\d{2})[./-](\d{1,2})[./-](\d{1,2})',s)
    if m:
        try: return date(int(m.group(1)),int(m.group(2)),int(m.group(3)))
        except: return None
    m=re.search(r'(20\d{2})년(\d{1,2})월(\d{1,2})일?',s)
    if m:
        try: return date(int(m.group(1)),int(m.group(2)),int(m.group(3)))
        except: return None
    return None

def extract_date_ranges(text):
    # 명시적 모집/접수 기간
    pairs=[]
    pats=[
      r'(20\d{2}[./-]\d{1,2}[./-]\d{1,2})\s*(?:~|[-–—]|부터)\s*(20\d{2}[./-]\d{1,2}[./-]\d{1,2})',
      r'(20\d{2}년\s*\d{1,2}월\s*\d{1,2}일?)\s*(?:~|[-–—]|부터)\s*(20\d{2}년\s*\d{1,2}월\s*\d{1,2}일?)'
    ]
    for p in pats:
        for m in re.finditer(p,text):
            a,b=parse_date(m.group(1)),parse_date(m.group(2))
            if a and b: pairs.append((a,b))
    return pairs

def extract_deadline(text):
    patterns=[
      r'(?:마감|접수마감|신청마감|모집마감|접수|신청|모집)[^0-9]{0,25}(20\d{2}[./-]\s*\d{1,2}[./-]\s*\d{1,2})',
      r'(20\d{2}[./-]\s*\d{1,2}[./-]\s*\d{1,2})[^가-힣]{0,15}(?:까지|마감)',
      r'(20\d{2}년\s*\d{1,2}월\s*\d{1,2}일?)[^가-힣]{0,15}(?:까지|마감)'
    ]
    for p in patterns:
        m=re.search(p,text)
        if m:
            d=parse_date(m.group(1))
            if d: return d,m.group(1)
    return None,""

def current_status(text):
    t=text.lower()
    if any(x in t for x in ["상시 모집","상시모집","현재 모집 중","현재모집중","지금 모집","모집 중","모집중"]):
        return "CURRENT_SIGNAL"
    ranges=extract_date_ranges(text)
    if ranges:
        active=[(a,b) for a,b in ranges if a<=TODAY<=b]
        if active: return "ACTIVE_RANGE"
        future=[(a,b) for a,b in ranges if b>=TODAY]
        if future: return "FUTURE_RANGE"
        return "EXPIRED_RANGE"
    d,raw=extract_deadline(text)
    if d:
        return "ACTIVE_OR_UNKNOWN" if d>=TODAY else "EXPIRED_DEADLINE"
    # 명시 날짜가 전혀 없으면, 모집 표현과 최신성 신호를 함께 본다.
    if any(x.lower() in t for x in RECRUIT_TERMS):
        return "NO_DATE_RECRUIT_SIGNAL"
    return "NO_DATE"

def extract_reward(text):
    hits=[]
    for p,_ in REWARD_PATTERNS:
        m=re.search(p,text,re.I)
        if m: hits.append(m.group(0))
    return ", ".join(dict.fromkeys(hits))[:200]

def score(text):
    s=0
    if any(x.lower() in text.lower() for x in RECRUIT_TERMS): s+=30
    for p,pts in REWARD_PATTERNS:
        if re.search(p,text,re.I): s+=pts
    if re.search(r'온라인|비대면|모바일|웹|링크|폼|구글폼|네이버폼',text,re.I): s+=8
    st=current_status(text)
    if st in ("CURRENT_SIGNAL","ACTIVE_RANGE","ACTIVE_OR_UNKNOWN"): s+=25
    elif st=="FUTURE_RANGE": s+=20
    elif st=="NO_DATE_RECRUIT_SIGNAL": s+=2
    return min(100,s),st

def exclusion_reason(title,text):
    tl=(title+" "+text).lower()
    if any(x.lower() in tl for x in EXCLUDE_TERMS): return "사용자 제외유형(FGI/좌담회/공모전/채용 등)"
    if any(x.lower() in title.lower() for x in INFO_TERMS): return "정보성/후기/가이드 제목"
    st=current_status(text)
    if st in ("EXPIRED_RANGE","EXPIRED_DEADLINE"): return "모집기간/마감일이 과거"
    if st=="NO_DATE" and not any(x.lower() in tl for x in RECRUIT_TERMS): return "모집 근거 부족"
    # 연도만 오래된 경우
    years=[int(y) for y in re.findall(r'20(\d{2})',text)]
    if years and max(years)<TODAY.year-1 and st not in ("CURRENT_SIGNAL","ACTIVE_RANGE"): return "오래된 자료"
    return ""

def search():
    out=[]
    with DDGS(timeout=15) as ddgs:
        for q in SEARCH_QUERIES:
            try:
                for x in ddgs.text(q,max_results=6,safesearch="moderate"):
                    title=norm(x.get("title")); url=x.get("href") or x.get("url") or ""
                    if title and url: out.append({"query":q,"title":title,"url":url,"snippet":norm(x.get("body") or x.get("snippet"))})
            except Exception: pass
    return out

def main():
    print(f"=== 티끌 알림봇 {VERSION} ===")
    raw=search()
    by={}
    for x in raw: by.setdefault(canonical_url(x["url"]),x)
    candidates=list(by.values())
    valid=[]; excluded=[]
    for x in candidates:
        body=fetch_page(x["url"])
        text=norm(x["title"]+" "+x["snippet"]+" "+body)
        reason=exclusion_reason(x["title"],text)
        if reason:
            excluded.append({"제목":x["title"],"URL":x["url"],"제외사유":reason})
            continue
        sc,st=score(text)
        if sc<35:
            excluded.append({"제목":x["title"],"URL":x["url"],"제외사유":"참여/보상 신호 부족"})
            continue
        valid.append({
            "발견일":TODAY.isoformat(),"점수":sc,"유형":detect_type(text),
            "제목":x["title"],"보상정보":extract_reward(text),
            "모집상태":st,"마감정보":extract_deadline(text)[1],
            "URL":x["url"],"도메인":urlparse(x["url"]).netloc,
            "미리보기":norm(body or x["snippet"])[:500]
        })
    # 상태 우선 + 점수
    valid.sort(key=lambda z:(z["모집상태"] not in ("CURRENT_SIGNAL","ACTIVE_RANGE","ACTIVE_OR_UNKNOWN"),-z["점수"]))
    state=load_state()
    new=[]
    for x in valid:
        fp=hashlib.sha256((canonical_url(x["URL"])+"|"+x["제목"]).encode()).hexdigest()[:20]
        if fp not in state["seen"]:
            state["seen"][fp]=TODAY.isoformat(); new.append(x)
    telegram_targets=new[:10]
    sent=send_telegram(telegram_targets)
    for x in telegram_targets:
        state["sent"][hashlib.sha256((canonical_url(x["URL"])+"|"+x["제목"]).encode()).hexdigest()[:20]]=datetime.now(KST).isoformat()
    save_state(state)
    import csv
    fields=["발견일","점수","유형","제목","보상정보","모집상태","마감정보","URL","도메인","미리보기"]
    with open("티끌_기회_검색결과.csv","w",newline="",encoding="utf-8-sig") as f:
        w=csv.DictWriter(f,fieldnames=fields); w.writeheader(); w.writerows(valid)
    with open("티끌_기회_제외결과.csv","w",newline="",encoding="utf-8-sig") as f:
        w=csv.DictWriter(f,fieldnames=["제목","URL","제외사유"]); w.writeheader(); w.writerows(excluded)
    summary={
      "version":VERSION,"기준일":TODAY.isoformat(),"검색어수":len(SEARCH_QUERIES),
      "검색원시결과":len(raw),"URL중복제거후":len(candidates),
      "유효후보":len(valid),"신규후보":len(new),"Telegram대상":len(telegram_targets),
      "텔레그램전송":sent,"제외후보":len(excluded),
      "운영원칙":{"FGI_좌담회_집단면접제외":True,"과거모집공고제외":True,
      "모집기간판정":True,"현재모집우선":True,"정보성글제외":True,
      "공모전채용입찰제외":True,"최대Telegram":10}
    }
    with open("티끌_기회_검색결과_요약.json","w",encoding="utf-8") as f: json.dump(summary,f,ensure_ascii=False,indent=2)
    print(json.dumps(summary,ensure_ascii=False,indent=2))

def detect_type(text):
    if re.search(r'사용성\s*테스트|제품\s*테스트|서비스\s*테스트|베타테스터',text,re.I): return "테스트"
    if re.search(r'인터뷰',text,re.I): return "인터뷰"
    if re.search(r'패널',text,re.I): return "패널"
    if re.search(r'설문|설문조사',text,re.I): return "설문"
    if re.search(r'모니터단',text,re.I): return "모니터단"
    return "기타"

def load_state():
    try:
        with open("opportunity_state.json",encoding="utf-8") as f: return json.load(f)
    except: return {"seen":{},"sent":{}}

def save_state(s):
    with open("opportunity_state.json","w",encoding="utf-8") as f: json.dump(s,f,ensure_ascii=False,indent=2)

def send_telegram(items):
    token=os.getenv("TELEGRAM_BOT_TOKEN",""); chat=os.getenv("TELEGRAM_CHAT_ID","")
    if not token or not chat: return 0
    sent=0
    for x in items:
        msg=(f"🎯 티끌 기회 발견 [{x['점수']}점]\n\n📌 {x['제목']}\n"
             f"🧩 유형: {x['유형']}\n🎁 {x['보상정보'] or '보상 확인 필요'}\n"
             f"📅 {x['마감정보'] or x['모집상태']}\n🌐 {x['도메인']}\n\n🔗 {x['URL']}")
        try:
            r=requests.post(f"https://api.telegram.org/bot{token}/sendMessage",json={"chat_id":chat,"text":msg},timeout=10)
            if r.ok: sent+=1
        except Exception: pass
    return sent

if __name__=="__main__": main()
