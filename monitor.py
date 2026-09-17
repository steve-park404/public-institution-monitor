import os,re,json,asyncio,hashlib
from datetime import datetime,timezone,timedelta
from urllib.parse import urljoin
import aiohttp
from bs4 import BeautifulSoup
import pandas as pd

BASE=os.path.dirname(os.path.abspath(__file__))
TARGET=os.path.join(BASE,"monitor_targets.xlsx")
KEYWORDS=os.path.join(BASE,"keywords.txt")
STATE=os.path.join(BASE,"state.json")
LOG=os.path.join(BASE,"monitor_log.json")
KST=timezone(timedelta(hours=9))

CONC=int(os.getenv("MAX_CONCURRENCY","12"))
TIMEOUT=int(os.getenv("TIMEOUT_SECONDS","15"))
RECENT=int(os.getenv("RECENT_POSTS","15"))
MAX_SECONDS=int(os.getenv("MAX_TOTAL_SECONDS","600"))
SEED=os.getenv("SEED_ON_FIRST_RUN","true").lower()=="true"

DETAIL=("act=view","mode=view","list_no=","articleNo=","article_no=","seq=","idx=","view.do","view.jsp","view.aspx")
BAD=("로그인","회원가입","예약","진료","검색","설문","채용","입찰","자료실","뉴스레터","소식지","개인정보","약관","사이트맵","교육신청","민원","발급","결제","일정","문의")

def clean(x): return re.sub(r"\s+"," ",str(x or "")).strip()
def now(): return datetime.now(KST).isoformat()
def load_keywords():
    if not os.path.exists(KEYWORDS): return []
    return list(dict.fromkeys(clean(x) for x in open(KEYWORDS,encoding="utf-8") if clean(x) and not clean(x).startswith("#")))
def load_state():
    try: return json.load(open(STATE,encoding="utf-8"))
    except: return {"seen":{},"initialized":False}
def save(path,obj):
    tmp=path+".tmp"
    json.dump(obj,open(tmp,"w",encoding="utf-8"),ensure_ascii=False,indent=2)
    os.replace(tmp,path)
def is_detail(u):
    u=u.lower()
    return any(x.lower() in u for x in DETAIL)
def bad_title(t):
    t=clean(t)
    return len(t)<3 or len(t)>300 or t in {"이전글이 없습니다.","다음글이 없습니다.","스킵네비게이션","본문","메뉴","홈","공지사항 상세"}
def atitle(a):
    return clean(a.get("title") or a.get("aria-label") or a.get_text(" ",strip=True))
def extract_posts(soup,base):
    out=[]; seen=set()
    selectors=["table tbody tr","table tr","ul.board-list li","ul.bbs-list li","ul.notice-list li",
               "div.board-list li","div.bbs-list li","div[class*='board'] li","div[class*='bbs'] li"]
    for sel in selectors:
        for row in soup.select(sel):
            vals=[]
            for a in row.find_all("a",href=True):
                t=atitle(a); u=urljoin(base,a["href"]).strip()
                if bad_title(t) or any(w in t for w in BAD) or not is_detail(u): continue
                vals.append((t,u))
            if vals:
                t,u=max(vals,key=lambda z:len(z[0]))
                if u not in seen: seen.add(u); out.append({"title":t,"url":u})
        if len(out)>=RECENT: break
    if len(out)<3:
        for a in soup.find_all("a",href=True):
            t=atitle(a); u=urljoin(base,a["href"]).strip()
            if bad_title(t) or any(w in t for w in BAD) or not is_detail(u) or u in seen: continue
            seen.add(u); out.append({"title":t,"url":u})
            if len(out)>=RECENT: break
    return out[:RECENT]
def body(soup):
    for x in soup(["script","style","noscript","svg","iframe"]): x.decompose()
    return clean(soup.get_text(" ",strip=True))

async def fetch(s,url,retry=True):
    for n in range(2 if retry else 1):
        try:
            async with s.get(url,allow_redirects=True,ssl=False,
                timeout=aiohttp.ClientTimeout(total=TIMEOUT),
                headers={"User-Agent":"Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.2)"}) as r:
                return r.status,str(r.url),await r.text(errors="ignore")
        except Exception as e:
            err=f"{type(e).__name__}: {e}"
            if n==0 and retry: await asyncio.sleep(1)
    return 0,url,err

async def check(s,sem,row,keywords,deadline):
    async with sem:
        if asyncio.get_running_loop().time()>=deadline: return {"timeout":1,"row":row}
        org=clean(row.get("기관명")); board=clean(row.get("게시판명")); url=clean(row.get("게시판URL"))
        if not url or url.lower()=="nan": return {"error":"EMPTY_URL","row":row}
        st,final,html=await fetch(s,url)
        if st!=200: return {"error":f"HTTP {st}","row":row}
        posts=extract_posts(BeautifulSoup(html,"html.parser"),final)
        matches=[]
        for p in posts:
            if asyncio.get_running_loop().time()>=deadline: break
            hit=[k for k in keywords if k.casefold() in p["title"].casefold()]
            if not hit:
                st2,_,h2=await fetch(s,p["url"],retry=False)
                if st2==200:
                    txt=body(BeautifulSoup(h2,"html.parser"))
                    hit=[k for k in keywords if k.casefold() in txt.casefold()]
            if hit:
                key=hashlib.sha256(f"{org}|{p['title']}|{p['url']}".encode()).hexdigest()
                matches.append({"key":key,"기관명":org,"게시판":board,"title":p["title"],"url":p["url"],"keywords":hit})
        return {"ok":1,"checked":len(posts),"matches":matches,"row":row}

async def telegram(matches):
    token=os.getenv("TELEGRAM_BOT_TOKEN","").strip(); chat=os.getenv("TELEGRAM_CHAT_ID","").strip()
    if not token or not chat: return {"ok":False,"error":"TELEGRAM secrets missing"}
    api=f"https://api.telegram.org/bot{token}/sendMessage"
    chunks=[]; cur="🚨 공공기관 공지사항 키워드 발견\n\n"
    for m in matches:
        b=f"🏢 {m['기관명']}\n📌 {m['게시판']}\n📝 {m['title']}\n🔑 {', '.join(m['keywords'])}\n🔗 {m['url']}\n\n"
        if len(cur)+len(b)>3800: chunks.append(cur); cur="🚨 공공기관 공지사항 키워드 발견\n\n"
        cur+=b
    chunks.append(cur)
    async with aiohttp.ClientSession() as s:
        for text in chunks[:10]:
            try:
                async with s.post(api,data={"chat_id":chat,"text":text},timeout=20) as r:
                    b=await r.text()
                    if r.status!=200: return {"ok":False,"error":f"HTTP {r.status}: {b}"}
            except Exception as e: return {"ok":False,"error":f"{type(e).__name__}: {e}"}
    return {"ok":True,"sent_chunks":len(chunks)}

async def main():
    kw=load_keywords()
    if not kw: raise SystemExit("keywords.txt에 키워드가 없습니다.")
    if not os.path.exists(TARGET): raise SystemExit("monitor_targets.xlsx가 없습니다.")
    df=pd.read_excel(TARGET).fillna("")
    rows=[]; seen_targets=set()
    for r in df.to_dict("records"):
        k=(clean(r.get("기관명")),clean(r.get("게시판URL")))
        if k[1] and k not in seen_targets: seen_targets.add(k); rows.append(r)

    state=load_state(); first=not state.get("initialized",False); state.setdefault("seen",{})
    deadline=asyncio.get_running_loop().time()+MAX_SECONDS
    sem=asyncio.Semaphore(CONC)
    conn=aiohttp.TCPConnector(limit=CONC,limit_per_host=3,ssl=False)
    results=[]; errors=[]; checked=0; completed=0; timed=0

    async with aiohttp.ClientSession(connector=conn) as s:
        tasks=[check(s,sem,r,kw,deadline) for r in rows]
        for f in asyncio.as_completed(tasks):
            if asyncio.get_running_loop().time()>=deadline: timed+=1; break
            x=await f
            if x.get("ok"):
                completed+=1; checked+=x["checked"]; results.extend(x["matches"])
            elif x.get("error"):
                r=x["row"]; errors.append({"기관명":clean(r.get("기관명")),"게시판":clean(r.get("게시판명")),"URL":clean(r.get("게시판URL")),"error":x["error"]})

    new=[]
    for m in results:
        if m["key"] not in state["seen"]:
            state["seen"][m["key"]]={"seen_at":now(),"title":m["title"],"url":m["url"]}
            if not first or not SEED: new.append(m)
    state["initialized"]=True; state["last_run"]=now(); save(STATE,state)

    tg={"ok":True,"skipped":True}
    if new: tg=await telegram(new)
    log={"run_at":now(),"targets":len(rows),"completed":completed,"posts_checked":checked,
         "new_matches":len(new),"errors":len(errors),"timed_out":timed,"first_run":first,
         "discovered_notice_boards":0,"telegram":tg,"errors_detail":errors[:100]}
    save(LOG,log); print(json.dumps(log,ensure_ascii=False,indent=2))

if __name__=="__main__": asyncio.run(main())
