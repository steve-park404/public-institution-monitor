
import os, re, json, time, html, warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, parse_qs
import requests
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import openpyxl

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

KEYWORDS = ["설문조사", "시민참여", "국민참여"]
EXCLUDE_TITLE = ["공모전"]
UA = "Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.9)"

MAX_CONCURRENCY=int(os.getenv("MAX_CONCURRENCY","20"))
TIMEOUT_SECONDS=int(os.getenv("TIMEOUT_SECONDS","15"))
RECENT_POSTS=int(os.getenv("RECENT_POSTS","15"))
MAX_TOTAL_SECONDS=int(os.getenv("MAX_TOTAL_SECONDS","1200"))
HTTP_RETRIES=int(os.getenv("HTTP_RETRIES","2"))
TELEGRAM_MAX_SEND=int(os.getenv("TELEGRAM_MAX_SEND","20"))

NOISE=["script","style","noscript","svg","header","footer","nav","aside","form","iframe","canvas","template"]
BOARD_WORDS=["공지사항","공지","알림마당","알림","소식","새소식","기관소식","게시판","뉴스","보도자료"]
URL_HINTS=["notice","noti","board","bbs","news","announcement","community","plaza","inform","particip"]
BAD=["채용","입찰","계약","로그인","회원","사이트맵","개인정보","이용약관"]

session=requests.Session()
session.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.5"})

def norm(s): return re.sub(r"\s+"," ",html.unescape(str(s or ""))).strip()

def get(url):
    for i in range(HTTP_RETRIES+1):
        try:
            r=session.get(url,timeout=TIMEOUT_SECONDS,allow_redirects=True)
            r.raise_for_status()
            if not r.encoding or r.encoding.lower()=="iso-8859-1":
                r.encoding=r.apparent_encoding
            return r
        except Exception:
            if i<HTTP_RETRIES: time.sleep(.4*(i+1))
    return None

def same_domain(a,b):
    try:
        return urlparse(a).netloc.lower().replace("www.","")==urlparse(b).netloc.lower().replace("www.","")
    except: return False

def title_of(soup):
    return norm(soup.title.get_text(" ",strip=True) if soup.title else "")

def clean(soup):
    for tag in NOISE:
        for x in soup.find_all(tag): x.decompose()
    return soup

def extract_links(page_url,soup):
    out=[]; seen=set()
    for a in soup.find_all("a",href=True):
        href=a.get("href","").strip()
        if not href or href.startswith(("javascript:","mailto:","tel:","#")): continue
        u=urljoin(page_url,href)
        if not u.startswith(("http://","https://")) or not same_domain(page_url,u): continue
        if u in seen: continue
        seen.add(u)
        out.append((u,norm(a.get_text(" ",strip=True))))
    return out

def detail_signal(u):
    x=u.lower(); q=parse_qs(urlparse(x).query)
    if any(k.lower() in {"seq","no","idx","nttno","articleid","article_id","id","view"} for k in q): return True
    return any(v in x for v in ["/view","/detail","/read","/article","/contents/view","/board/view"])

def detail_url(u):
    x=u.lower()
    if any(v in x for v in ["login","member","delete","write","modify"]): return False
    return detail_signal(u)

def board_score(u,t):
    s=0; x=u.lower(); tt=t.lower()
    s += sum(7 for w in BOARD_WORDS if w.lower() in tt)
    s += sum(3 for h in URL_HINTS if h in x)
    s -= sum(5 for b in BAD if b.lower() in tt)
    if detail_signal(u): s-=4
    return s

def discover_board(home):
    r=get(home)
    if not r: return None,[],"HOME_ERROR"
    soup=BeautifulSoup(r.text,"html.parser")
    links=extract_links(r.url,soup)
    scored=sorted([(board_score(u,t),u,t) for u,t in links],reverse=True)
    for score,u,t in scored[:35]:
        if score<4: break
        rr=get(u)
        if not rr: continue
        ss=BeautifulSoup(rr.text,"html.parser")
        details=list(dict.fromkeys([x for x,tx in extract_links(rr.url,ss) if detail_url(x)]))
        if len(details)>=2:
            return rr.url,details[:RECENT_POSTS],"DISCOVERED"
    return None,[],"NO_BOARD"

def visible_main_text(soup):
    soup=clean(soup)
    main=soup.find("main") or soup.find("article")
    if not main:
        candidates=soup.find_all(["div","section"],id=re.compile(r"(content|contents|sub|body|article)",re.I))
        main=max(candidates,key=lambda x:len(x.get_text(" ",strip=True))) if candidates else soup.body
    if not main: return ""
    # Remove common navigation-like nodes nested in content
    for x in main.find_all(["ul","ol"]):
        txt=norm(x.get_text(" ",strip=True))
        if len(txt)<400 and sum(1 for k in BOARD_WORDS if k in txt)>=1:
            x.decompose()
    return norm(main.get_text(" ",strip=True))

def meaningful_body_match(body,kw):
    # Require keyword in a reasonably long content segment and avoid menus/metadata.
    for m in re.finditer(re.escape(kw),body):
        a=max(0,m.start()-140); b=min(len(body),m.end()+180)
        ctx=body[a:b]
        # Context should contain sentence-like content, not just a navigation cluster.
        if len(ctx)>=80 and any(p in ctx for p in [".","다.","요.","습니다","한다","안내","실시","모집","참여","응답","기간"]):
            return True
    return False

def match_post(u):
    r=get(u)
    if not r: return None
    soup=BeautifulSoup(r.text,"html.parser")
    title=title_of(soup)
    if not title: return None
    # Exact operational policy: 공모전 제목은 제외.
    if any(x in title for x in EXCLUDE_TITLE): return None
    for kw in KEYWORDS:
        if kw in title:
            return {"url":r.url,"title":title,"keyword":kw,"where":"title"}
    body=visible_main_text(soup)
    if len(body)<120: return None
    for kw in KEYWORDS:
        if meaningful_body_match(body,kw):
            return {"url":r.url,"title":title,"keyword":kw,"where":"body"}
    return None

def load_targets():
    wb=openpyxl.load_workbook("monitor_targets.xlsx",read_only=True,data_only=True)
    ws=wb["355기관"]; hs=[c.value for c in next(ws.iter_rows())]
    ix={str(v):i for i,v in enumerate(hs)}
    out=[]
    for row in ws.iter_rows(min_row=2,values_only=True):
        if not row[ix["기관명"]]: continue
        out.append({
            "name":str(row[ix["기관명"]]).strip(),
            "home":str(row[ix["홈페이지URL"]] or "").strip(),
            "board":str(row[ix["공지게시판URL"]] or "").strip()
        })
    return out[:355]

def load_state():
    try:
        with open("state.json","r",encoding="utf-8") as f:s=json.load(f)
    except: s={}
    if not isinstance(s,dict): s={}
    if not isinstance(s.get("seen"),dict):
        s["seen"]={str(x):1 for x in s.get("seen",[])[-10000:]}
    if not isinstance(s.get("boards"),dict): s["boards"]={}
    if not isinstance(s.get("stats"),dict): s["stats"]={}
    return s

def save_state(s):
    s["seen"]=dict(list(s.get("seen",{}).items())[-10000:])
    with open("state.json","w",encoding="utf-8") as f: json.dump(s,f,ensure_ascii=False,indent=2)

def process(t,state,deadline):
    name,home,seed=t["name"],t["home"],t["board"]
    if time.time()>=deadline:return {"name":name,"status":"DEADLINE","posts":0,"matches":[]}
    board=seed or state["boards"].get(name,"")
    details=[]
    if board:
        rr=get(board)
        if rr:
            ss=BeautifulSoup(rr.text,"html.parser")
            details=list(dict.fromkeys([u for u,tx in extract_links(rr.url,ss) if detail_url(u)]))[:RECENT_POSTS]
        else: board=""
    if not board and home:
        board,details,status=discover_board(home)
        if board: state["boards"][name]=board
    if not board:return {"name":name,"status":"NO_BOARD","posts":0,"matches":[]}
    matches=[]
    for u in details:
        if time.time()>=deadline: break
        m=match_post(u)
        if not m: continue
        key=m["url"]+"|"+m["keyword"]
        if key not in state["seen"]:
            state["seen"][key]=1; m["기관명"]=name; matches.append(m)
    return {"name":name,"status":"OK","posts":len(details),"matches":matches,"board":board}

def send(items):
    token=os.getenv("TELEGRAM_BOT_TOKEN","").strip(); chat=os.getenv("TELEGRAM_CHAT_ID","").strip()
    if not token or not chat: print("TELEGRAM_CONFIG_MISSING"); return 0
    n=0
    for m in items[:TELEGRAM_MAX_SEND]:
        text=f"📢 공공기관 참여정보 알림\n기관: {m['기관명']}\n키워드: {m['keyword']}\n제목: {m['title']}\n검색위치: {m['where']}\n{m['url']}"
        try:
            r=requests.post(f"https://api.telegram.org/bot{token}/sendMessage",json={"chat_id":chat,"text":text},timeout=20)
            if r.ok and r.json().get("ok"): n+=1
            else: print("TELEGRAM_ERROR",r.status_code,r.text[:300])
        except Exception as e: print("TELEGRAM_EXCEPTION",repr(e))
    print("TELEGRAM_SENT",n); return n

def main():
    start=time.time(); deadline=start+MAX_TOTAL_SECONDS
    targets=load_targets(); state=load_state(); results=[]; matches=[]
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as ex:
        fs=[ex.submit(process,t,state,deadline) for t in targets]
        for f in as_completed(fs):
            try:r=f.result()
            except Exception as e:r={"name":"?","status":"ERROR","posts":0,"matches":[],"error":repr(e)}
            results.append(r); matches.extend(r.get("matches",[]))
    save_state(state); sent=send(matches)
    summary={
        "targets":len(targets),
        "completed":sum(r["status"]=="OK" for r in results),
        "no_board":sum(r["status"]=="NO_BOARD" for r in results),
        "posts_checked":sum(r.get("posts",0) for r in results),
        "new_matches":len(matches),
        "telegram_sent":sent,
        "errors":sum(r["status"]=="ERROR" for r in results),
        "elapsed_seconds":round(time.time()-start,1),
        "timed_out":time.time()>=deadline
    }
    with open("monitor_log.json","w",encoding="utf-8") as f:
        json.dump({"summary":summary,"results":results,"matches":matches[-200:]},f,ensure_ascii=False,indent=2)
    print("SUMMARY",json.dumps(summary,ensure_ascii=False))

if __name__=="__main__": main()
