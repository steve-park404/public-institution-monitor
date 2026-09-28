import os,re,csv,json,time,hashlib,warnings
from datetime import datetime,timedelta,timezone
from urllib.parse import urljoin,urlparse,parse_qsl,urlencode,urlunparse
from concurrent.futures import ThreadPoolExecutor,as_completed
import requests
from bs4 import BeautifulSoup,XMLParsedAsHTMLWarning
warnings.filterwarnings("ignore",category=XMLParsedAsHTMLWarning)

VERSION="V8.15.6"; INPUT_CSV=os.getenv("INPUT_CSV","기관별_상태.csv")
RECENT_DAYS=int(os.getenv("RECENT_DAYS","30")); MAX_SEND=int(os.getenv("TELEGRAM_MAX_SEND","20"))
WORKERS=int(os.getenv("WORKERS","12")); TIMEOUT=int(os.getenv("REQUEST_TIMEOUT","10"))
MAX_LINKS=int(os.getenv("MAX_LIST_LINKS","120")); MAX_DETAIL=int(os.getenv("MAX_DETAIL_PER_BOARD","15"))
KEYWORDS=["설문조사","시민참여","국민참여"]; EXCLUDE=["공모전"]; KST=timezone(timedelta(hours=9))
HEAD={"User-Agent":"Mozilla/5.0 Chrome/128 Safari/537.36"}

def clean(x): return re.sub(r"\s+"," ",str(x or "")).strip()
def norm(u):
    u=str(u or "").strip()
    return u if re.match(r"^https?://",u,re.I) else ""
def canon(u):
    u=norm(u)
    if not u:return ""
    p=urlparse(u); drop={"utm_source","utm_medium","utm_campaign","utm_term","utm_content"}
    q=[x for x in parse_qsl(p.query,keep_blank_values=True) if x[0] not in drop]
    return urlunparse((p.scheme.lower(),p.netloc.lower(),p.path,"",urlencode(q),""))
def dtxt(s):
    for p in [r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})",r"(20\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일"]:
        m=re.search(p,clean(s))
        if m:
            try:return datetime(*map(int,m.groups()),tzinfo=KST)
            except:pass
    return None
def hits(s): return [k for k in KEYWORDS if k in s]
def loadj(p,d):
    try:return json.load(open(p,encoding="utf-8"))
    except:return d
def savej(p,d):
    with open(p+".tmp","w",encoding="utf-8") as f:json.dump(d,f,ensure_ascii=False,indent=2)
    os.replace(p+".tmp",p)
def get(s,u):
    r=s.get(u,headers=HEAD,timeout=TIMEOUT,allow_redirects=True);r.raise_for_status()
    r.encoding=r.apparent_encoding or r.encoding;return r.url,r.text

def candidates(url,html):
    soup=BeautifulSoup(html,"html.parser"); out=[];seen=set()
    for a in soup.find_all("a",href=True):
        t=clean(a.get_text(" ",strip=True)); u=urljoin(url,a["href"])
        if not u.startswith(("http://","https://")) or len(t)<2 or len(t)>300:continue
        if t in {"홈","로그인","회원가입","사이트맵","개인정보처리방침"}:continue
        c=canon(u)
        if not c or c in seen:continue
        seen.add(c); pt=clean(a.parent.get_text(" ",strip=True)) if a.parent else t
        out.append((t,u,dtxt(pt) or dtxt(t)))
    return out[:MAX_LINKS]

def one(row):
    agency=clean(row.get("기관명")); board=norm(row.get("게시판URL"))
    out={"기관명":agency,"게시판URL":board,"게시판확인":"실패","게시물확인":0,"매칭":[],"오류":""}
    if not board:out["게시판확인"]="URL없음";return out
    s=requests.Session();cut= datetime.now(KST)-timedelta(days=RECENT_DAYS)
    try:
        fu,html=get(s,board);out["게시판확인"]="성공"
        cs=candidates(fu,html);out["게시물확인"]=len(cs); targets=[]
        for c in cs:
            t,u,d=c
            if any(x in t for x in EXCLUDE) or (d and d<cut):continue
            h=hits(t)
            if h or d: targets.append(c)
            if len(targets)>=MAX_DETAIL:break
        for t,u,d in targets:
            try:
                du,bh=get(s,u); soup=BeautifulSoup(bh,"html.parser")
                for z in soup(["script","style","noscript"]):z.decompose()
                body=clean(soup.get_text(" ",strip=True)); hh=hits(t+" "+body)
                dd=d or dtxt(body)
                if not hh or any(x in t for x in EXCLUDE) or (dd and dd<cut):continue
                ident=canon(du or u); fp=hashlib.sha256((agency+"|"+t+"|"+ident).encode()).hexdigest()
                out["매칭"].append({"기관명":agency,"제목":t,"URL":du or u,"게시일":dd.strftime("%Y-%m-%d") if dd else "","키워드":",".join(hh),"본문매칭":any(k in body for k in hh),"fingerprint":fp})
            except:pass
        return out
    except Exception as e:out["오류"]=f"{type(e).__name__}: {e}";return out

def tg(token,chat,msg):
    try:return requests.post(f"https://api.telegram.org/bot{token}/sendMessage",data={"chat_id":chat,"text":msg,"disable_web_page_preview":True},timeout=15).ok
    except:return False

rows=list(csv.DictReader(open(INPUT_CSV,encoding="utf-8-sig",newline="")))
if len(rows)<300:raise RuntimeError(f"기관 데이터가 부족합니다: {len(rows)}개")
targets=[r for r in rows if norm(r.get("게시판URL"))];missing=[r for r in rows if not norm(r.get("게시판URL"))]
print(f"VERSION={VERSION} targets_total={len(rows)} board_url_targets={len(targets)} board_url_missing={len(missing)} workers={WORKERS}")

results=[];matches=[]
with ThreadPoolExecutor(max_workers=WORKERS) as ex:
    fs={ex.submit(one,r):r.get("기관명","") for r in targets}
    for i,f in enumerate(as_completed(fs),1):
        r=f.result();results.append(r);matches+=r["매칭"]
        print(f"[{i}/{len(targets)}] {r['기관명']} | {r['게시판확인']} | matches={len(r['매칭'])}",flush=True)

sent=set(loadj("sent_url_ledger.json",[])); sentfp=set(loadj("sent_fingerprint_ledger.json",[])); checked=set(loadj("checked_post_identity_ledger.json",[]))
token=os.getenv("TELEGRAM_BOT_TOKEN","");chat=os.getenv("TELEGRAM_CHAT_ID","");sentn=0
for x in matches:
    u=canon(x["URL"])
    if u in sent or x["fingerprint"] in sentfp:continue
    if sentn<MAX_SEND and tg(token,chat,f"🔔 {x['기관명']}\n📌 {x['제목']}\n🔎 {x['키워드']}\n📅 {x['게시일'] or '날짜 미확인'}\n🔗 {x['URL']}"):
        sent.add(u);sentfp.add(x["fingerprint"]);sentn+=1
    checked.add(u)
savej("sent_url_ledger.json",sorted(sent));savej("sent_fingerprint_ledger.json",sorted(sentfp));savej("checked_post_identity_ledger.json",sorted(checked))

with open("검색결과_V8.15.6.csv","w",encoding="utf-8-sig",newline="") as f:
    w=csv.DictWriter(f,fieldnames=["기관명","제목","URL","게시일","키워드","본문매칭"]);w.writeheader()
    for x in matches:w.writerow({k:x.get(k,"") for k in w.fieldnames})
with open("게시판검증결과_V8.15.6.csv","w",encoding="utf-8-sig",newline="") as f:
    fs=["기관명","게시판URL","게시판확인","게시물확인","매칭건수","오류"];w=csv.DictWriter(f,fieldnames=fs);w.writeheader()
    for x in results:w.writerow({"기관명":x["기관명"],"게시판URL":x["게시판URL"],"게시판확인":x["게시판확인"],"게시물확인":x["게시물확인"],"매칭건수":len(x["매칭"]),"오류":x["오류"]})
    for x in missing:w.writerow({"기관명":x.get("기관명",""),"게시판URL":"","게시판확인":"URL없음","게시물확인":0,"매칭건수":0,"오류":"게시판URL 미입력"})
summary={"version":VERSION,"targets_total":len(rows),"board_url_targets":len(targets),"board_url_missing":len(missing),"board_verified":sum(x["게시판확인"]=="성공" for x in results),"board_failed":sum(x["게시판확인"]=="실패" for x in results),"posts_checked":sum(x["게시물확인"] for x in results),"new_matches":len(matches),"telegram_sent":sentn,"elapsed_seconds":None}
savej("summary_V8.15.6.json",summary);print(json.dumps(summary,ensure_ascii=False,indent=2))
