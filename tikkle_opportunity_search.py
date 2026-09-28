
# 티끌 알림봇 V2.1 - 실제 모집기회 필터 강화
import os,re,json,hashlib
from datetime import datetime,timedelta,timezone
from urllib.parse import urlparse
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS

VERSION="V2.2"
KST=timezone(timedelta(hours=9))

SEARCH_QUERIES=[
'"설문조사" "참여자 모집" "기프티콘"',
'"설문조사" "참여자 모집" "상품권"',
'"설문조사" "사례비" 모집',
'"설문" "참여자 모집" "사례비"',
'"패널 모집" "사례비"',
'"패널 모집" "기프티콘"',
'"인터뷰 참여자 모집" "사례비"',
'"사용성 테스트" "참가자 모집" "사례비"',
'"UX 테스트" "참가자 모집"',
'"제품 테스트" "참여자 모집"',
'"서비스 테스트" "참여자 모집"',
'"베타테스터" "모집" "사례비"',
'"모니터단 모집" "상품권"',
'"소비자 조사" "참여자 모집"',
'"소비자 패널" "모집" "사례비"',
'"시민참여" "설문" "기프티콘"',
'"국민참여" "설문" "상품권"',
'"정책 설문" "참여자 모집"',
'"만족도 조사" "참여자 모집" "경품"',
'"설문 이벤트" "참여자 모집"',
'"리서치 참여자 모집" "사례비"',
]

# V2.0에서 발견된 오탐 유형
HARD_EXCLUDE_TITLE=[
"FGI","fgi","좌담회","집단면접",
"가이드","완벽 가이드","방법","후기","후기｜","후기 |","리뷰","정리","소개",
"뜻","개념","팁","노하우","대행","에이전시","전문 에이전시","서비스 소개",
"플랫폼","교환","추천","광고","홍보","블로그","칼럼","매뉴얼"
]
HARD_EXCLUDE_DOMAIN=[
"eventhouse.kr","pickply.com","superplanning.co.kr"
]
EXCLUDE_TITLE=[
"공모전","채용","입찰","구인","구직","장학생","장학금","보도자료","대행업체"
]
# 홈/카테고리/플랫폼 자체 페이지는 모집 공고로 취급하지 않음
GENERIC_PATH=["/","/category/","/search","/tag/","/archive","/blog","/service","/gig/"]
RECRUIT_PATTERNS=[
r"참여자\s*모집",r"참가자\s*모집",r"모집\s*(?:공고|안내)",
r"패널\s*모집",r"설문\s*(?:참여|응답)\s*모집",r"인터뷰\s*참여자",
r"사용성\s*테스트.*모집",
r"테스트.*참여자.*모집",r"베타테스터.*모집",r"모니터단.*모집",
r"리서치.*참여자.*모집"
]
REWARD_PATTERNS=[
(r"(\d[\d,]*)\s*만원",35),(r"(\d[\d,]*)\s*원",25),
(r"사례비|참여비|참여수당",28),(r"상품권|기프티콘|커피\s*쿠폰|스타벅스",18),
(r"경품|추첨",8),(r"페이|포인트",10)
]
TYPE_PATTERNS=[
(r"설문조사|설문",10,"설문"),(r"패널",12,"패널"),(r"인터뷰",18,"인터뷰"),
(r"사용성\s*테스트|ux\s*테스트",22,"사용성테스트"),
(r"제품\s*테스트|서비스\s*테스트",18,"제품/서비스테스트"),
(r"베타테스터",15,"베타테스트"),(r"모니터단",12,"모니터단")
]
RESTRICTION=[(r"대학생|대학원생|교직원|의료인|공무원|청소년|임산부|환자", -8),
(r"특정\s*지역|해당\s*지역|서울\s*거주|부산\s*거주", -5)]

def norm(s): return re.sub(r"\s+"," ",s or "").strip()
def canon(url):
    try:
        p=urlparse(url); return p.netloc.lower().replace("www.","")+re.sub(r"/+$","",p.path)
    except:return url
def fetch(url):
    try:
        r=requests.get(url,timeout=10,headers={"User-Agent":"Mozilla/5.0 TikkleBot/2.1"})
        if r.status_code>=400:return ""
        r.encoding=r.apparent_encoding or r.encoding
        s=BeautifulSoup(r.text,"lxml")
        for x in s(["script","style","noscript","svg","nav","footer"]): x.decompose()
        return norm(s.get_text(" ",strip=True))[:15000]
    except:return ""

def parse_date(text):
    pats=[
      r"(20\d{2})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})",
      r"(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일",
    ]
    dates=[]
    for p in pats:
        for m in re.finditer(p,text):
            try: dates.append(datetime(int(m.group(1)),int(m.group(2)),int(m.group(3)),tzinfo=KST))
            except: pass
    return max(dates) if dates else None

def deadline(text):
    patterns=[
      r"(?:마감|접수|신청|모집)[^0-9]{0,25}(20\d{2}[.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2})",
      r"(20\d{2}[.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2})[^가-힣]{0,15}(?:까지|마감)",
      r"(\d{1,2}월\s*\d{1,2}일)[^가-힣]{0,15}(?:까지|마감)"
    ]
    for p in patterns:
        m=re.search(p,text)
        if m:return m.group(1)
    return ""

def reward(text):
    out=[]
    for p,_ in REWARD_PATTERNS:
        m=re.search(p,text,re.I)
        if m: out.append(m.group(0))
    return ", ".join(dict.fromkeys(out))[:200]

def classify(title,text,url):
    t=(title+" "+text).lower()
    if any(x.lower() in title.lower() for x in EXCLUDE_TITLE): return False,"제외: 공모전/채용/입찰 등",""
    if any(x.lower() in title.lower() for x in HARD_EXCLUDE_TITLE): return False,"제외: 정보성/후기/가이드",""
    domain=urlparse(url).netloc.lower().replace("www.","")
    if domain in HARD_EXCLUDE_DOMAIN: return False,"제외: 정보성/업체 페이지", ""
    # 강한 모집 표현이 없고 단순 정보/후기이면 제외
    types=[lab for p,_,lab in TYPE_PATTERNS if re.search(p,t,re.I)]
    recruit=any(re.search(p,t,re.I) for p in RECRUIT_PATTERNS)
    if not recruit:
        # 제목에 모집/참여자/사례비가 없으면 일반 정보로 판단
        if not re.search(r"모집|참여자|참가자|사례비|참여수당",title,re.I):
            return False,"제외: 모집 의도 불명확",""
    # 과거 날짜가 제목/본문에 여러 번 나오고 최근 모집 신호가 없으면 제외
    cutoff=datetime.now(KST)-timedelta(days=180)
    ds=parse_date(text[:12000])
    dl=deadline(text)
    if ds and ds < cutoff and not re.search(r"현재|모집중|모집\s*중|진행중|진행\s*중|2026",t,re.I):
        return False,"제외: 오래된 페이지",""
    score=0
    reasons=[]
    if recruit: score+=35; reasons.append("모집공고")
    for p,pts,lab in TYPE_PATTERNS:
        if re.search(p,t,re.I): score+=pts; reasons.append(lab)
    for p,pts in REWARD_PATTERNS:
        if re.search(p,t,re.I): score+=pts; reasons.append("보상")
    for p,pts in RESTRICTION:
        if re.search(p,t,re.I): score+=pts
    if dl: score+=8; reasons.append("마감정보")
    # 온라인 참여 신호
    if re.search(r"온라인|비대면|모바일|PC|링크로\s*참여",t,re.I):
        score+=5; reasons.append("온라인")
    # 금액이 명시된 실제 사례비는 가점
    if re.search(r"\d[\d,]*\s*만원|\d[\d,]*\s*원",t):
        score+=8; reasons.append("금액명시")
    return score>=40,"유효" if score>=40 else "점수미달",",".join(dict.fromkeys(reasons))

def search():
    out=[]
    with DDGS(timeout=15) as d:
        for q in SEARCH_QUERIES:
            try:
                for x in d.text(q,max_results=8,safesearch="moderate"):
                    title=norm(x.get("title")); url=x.get("href") or x.get("url") or ""; sn=norm(x.get("body") or x.get("snippet"))
                    if title and url: out.append({"query":q,"title":title,"url":url,"snippet":sn})
            except: pass
    return out

def main():
    raw=search()
    unique={}
    for x in raw:
        k=canon(x["url"])
        if k not in unique: unique[k]=x
    candidates=list(unique.values())
    valid=[]; excluded=[]
    for x in candidates:
        body=fetch(x["url"])
        combined=norm(x["title"]+" "+x["snippet"]+" "+body)
        ok,reason,reasons=classify(x["title"],combined,x["url"])
        if not ok:
            excluded.append({"제목":x["title"],"URL":x["url"],"제외사유":reason})
            continue
        # reward must be present for Telegram priority; keep low-score candidates only in CSV
        sc=0
        if re.search(r"모집|참여자|참가자",combined,re.I): sc+=35
        for p,pts,lab in TYPE_PATTERNS:
            if re.search(p,combined,re.I): sc+=pts
        for p,pts in REWARD_PATTERNS:
            if re.search(p,combined,re.I): sc+=pts
        if deadline(combined): sc+=8
        if re.search(r"온라인|비대면|모바일|PC",combined,re.I): sc+=5
        if re.search(r"\d[\d,]*\s*만원|\d[\d,]*\s*원",combined): sc+=8
        for p,pts in RESTRICTION:
            if re.search(p,combined,re.I): sc+=pts
        valid.append({
          "발견일":datetime.now(KST).strftime("%Y-%m-%d"),
          "점수":max(0,min(100,sc)),
          "유형":",".join(dict.fromkeys([lab for p,_,lab in TYPE_PATTERNS if re.search(p,combined,re.I)])),
          "제목":x["title"],"보상정보":reward(combined),"마감정보":deadline(combined),
          "URL":x["url"],"도메인":urlparse(x["url"]).netloc,
          "판정근거":reasons,"미리보기":norm(body or x["snippet"])[:500]
        })
    valid.sort(key=lambda z:z["점수"],reverse=True)

    state={}
    try:
        state=json.load(open("opportunity_state.json",encoding="utf-8"))
    except: state={"seen":{},"sent":{}}
    new=[]
    for x in valid:
        fp=hashlib.sha256((canon(x["URL"])+"|"+x["제목"]).encode()).hexdigest()[:20]
        if fp not in state["seen"]:
            state["seen"][fp]={"title":x["제목"],"url":x["URL"],"first_seen":x["발견일"],"score":x["점수"]}
            new.append(x)

    # Telegram은 '실제 모집 + 보상 명시' 중심으로 최대 10건.
    notify=[x for x in new if x["보상정보"] and x["점수"]>=60][:10]
    token=os.getenv("TELEGRAM_BOT_TOKEN",""); chat=os.getenv("TELEGRAM_CHAT_ID","")
    sent=0
    for x in notify:
        msg=(f"🎯 티끌 기회 [{x['점수']}점]\n\n📌 {x['제목']}\n"
             f"🧩 유형: {x['유형'] or '참여기회'}\n🎁 보상: {x['보상정보']}\n"
             f"📅 마감: {x['마감정보'] or '확인 필요'}\n🌐 {x['도메인']}\n\n🔗 {x['URL']}")
        try:
            r=requests.post(f"https://api.telegram.org/bot{token}/sendMessage",json={"chat_id":chat,"text":msg},timeout=10)
            if r.ok: sent+=1
        except: pass
    for x in notify:
        fp=hashlib.sha256((canon(x["URL"])+"|"+x["제목"]).encode()).hexdigest()[:20]
        state["sent"][fp]=datetime.now(KST).isoformat()
    json.dump(state,open("opportunity_state.json","w",encoding="utf-8"),ensure_ascii=False,indent=2)

    import csv
    fields=["발견일","점수","유형","제목","보상정보","마감정보","URL","도메인","판정근거","미리보기"]
    with open("티끌_기회_검색결과.csv","w",newline="",encoding="utf-8-sig") as f:
        w=csv.DictWriter(f,fieldnames=fields); w.writeheader(); w.writerows(valid)
    with open("티끌_기회_제외결과.csv","w",newline="",encoding="utf-8-sig") as f:
        w=csv.DictWriter(f,fieldnames=["제목","URL","제외사유"]); w.writeheader(); w.writerows(excluded)

    summary={"version":VERSION,"검색어수":len(SEARCH_QUERIES),"검색원시결과":len(raw),
             "URL중복제거후":len(candidates),"유효후보":len(valid),"신규후보":len(new),
             "Telegram대상":len(notify),"텔레그램전송":sent,
             "제외후보":len(excluded),
             "운영원칙":{"실제모집중심":True,"보상명시우선":True,"정보성글제외":True,
                         "오래된글제외":True,"공모전채용입찰제외":True,"최대Telegram":10}}
    json.dump(summary,open("티끌_기회_검색결과_요약.json","w",encoding="utf-8"),ensure_ascii=False,indent=2)
    print(json.dumps(summary,ensure_ascii=False,indent=2))

if __name__=="__main__": main()
