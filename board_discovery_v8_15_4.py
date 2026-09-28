import concurrent.futures,csv,json,os,re,sys,threading
from datetime import datetime,timezone,timedelta
from urllib.parse import urljoin,urlparse,parse_qs,urlencode,urlunparse
import requests
from bs4 import BeautifulSoup

KST=timezone(timedelta(hours=9)); NOW=datetime.now(KST)
ROOT=os.getcwd(); CACHE_PATH=os.path.join(ROOT,"board_cache.json")
OUT_PATH=os.path.join(ROOT,"게시판_자동발견_V8.15.4_결과.csv")
LOG_PATH=os.path.join(ROOT,"게시판_자동발견_V8.15.4_진단.json")
MAX_WORKERS=12; REQUEST_TIMEOUT=8
BOARD_WORDS=["공지사항","공지","알림","알림마당","새소식","기관소식","소식","게시판","국민참여","시민참여","설문","참여"]
thread_local=threading.local()

def session():
    if not hasattr(thread_local,"s"):
        s=requests.Session(); s.headers.update({"User-Agent":"Mozilla/5.0 (compatible; TikkleBoardDiscovery/8.15.4)","Accept-Language":"ko-KR,ko;q=0.9,en;q=0.5"}); thread_local.s=s
    return thread_local.s
def norm(v): return str(v or "").strip()
def normalize_url(url):
    url=norm(url).strip(" ,;)]}")
    if not url:return ""
    if not re.match(r"^https?://",url,re.I):url="https://"+url
    try:
        p=urlparse(url); _=p.hostname; _=p.port
        return urlunparse((p.scheme,p.netloc,p.path or "/",p.params,p.query,"")) if p.netloc else ""
    except (ValueError,TypeError): return ""
def canonical(url):
    try:
        p=urlparse(url); q=parse_qs(p.query)
        return urlunparse((p.scheme,p.netloc,p.path,"",urlencode({"bbsId":q["bbsId"][0]}),"")) if "bbsId" in q else normalize_url(url)
    except Exception:return ""
def load_cache():
    try:
        with open(CACHE_PATH,encoding="utf-8") as f:d=json.load(f)
        return d if isinstance(d,dict) else {}
    except Exception:return {}
def save_cache(c):
    tmp=CACHE_PATH+".tmp"
    with open(tmp,"w",encoding="utf-8") as f:json.dump(c,f,ensure_ascii=False,indent=2)
    os.replace(tmp,CACHE_PATH)
def institution_like(s):
    s=norm(s)
    return bool(s and len(s)<=100 and re.search(r"[가-힣A-Za-z]",s) and not normalize_url(s))
def load_targets(cache):
    t={}
    for name,item in cache.items():
        if isinstance(item,dict):
            home=normalize_url(item.get("homepage") or item.get("home_url") or item.get("URL"))
            if name and home:t[name]=home
    # 보조: 기존 CSV의 URL 열에서 홈페이지를 복원. 이미 캐시된 기관은 유지.
    for fn in os.listdir(ROOT):
        if not fn.lower().endswith(".csv"):continue
        try:
            with open(os.path.join(ROOT,fn),encoding="utf-8-sig",newline="") as f:rows=list(csv.reader(f))
            for row in rows[1:]:
                urls=[normalize_url(x) for x in row if normalize_url(x)]
                if not urls:continue
                home=urls[0]
                names=[norm(x) for x in row if institution_like(x)]
                if names:
                    name=max(names,key=len); t.setdefault(name,home)
        except Exception:continue
    return t
def fetch(url):
    try:
        r=session().get(url,timeout=REQUEST_TIMEOUT,allow_redirects=True)
        if r.status_code>=400:return None,r.url,""
        r.encoding=r.apparent_encoding or r.encoding
        return r,r.url,r.text
    except Exception:return None,url,""
def score(u,a,t):
    x=f"{u} {a} {t}".lower();s=0
    if "selectboardlist.do" in x:s+=70
    if "bbsid=bbsMSTR_".lower() in x:s+=30
    for w in BOARD_WORDS:
        if w.lower() in a.lower():s+=15
    return s
def candidates(base,html):
    soup=BeautifulSoup(html,"html.parser"); title=soup.title.get_text(" ",strip=True) if soup.title else ""; out=[]
    for a in soup.find_all("a",href=True):
        h=norm(a.get("href"))
        if not h or h.lower().startswith(("javascript:","#","mailto:")):continue
        u=normalize_url(urljoin(base,h)); txt=a.get_text(" ",strip=True)
        if u and ("selectboardlist.do" in u.lower() or "bbsid=" in u.lower() or any(w in txt for w in BOARD_WORDS)):out.append((score(u,txt,title),canonical(u),txt))
    for m in re.finditer(r"""[^"'<>\\s]{0,150}/cop/bbs/selectBoardList\.do\?[^"'<>\\s]+""",html,re.I):
        u=normalize_url(urljoin(base,m.group(0)))
        if u:out.append((score(u,"",title),canonical(u),""))
    d={}
    for x in out:
        if x[1] and (x[1] not in d or x[0]>d[x[1]][0]):d[x[1]]=x
    return sorted(d.values(),reverse=True)
def verify(u,name):
    r,final,html=fetch(u)
    if not r or not html:return False,{"reason":"FETCH_FAIL","url":final}
    soup=BeautifulSoup(html,"html.parser"); title=soup.title.get_text(" ",strip=True) if soup.title else ""; body=soup.get_text(" ",strip=True); sig=0
    if "selectboardlist.do" in final.lower():sig+=2
    if "bbsid=" in final.lower():sig+=2
    if re.search(r"번호\s+제목|제목\s+작성자|작성자\s+날짜|조회",body):sig+=2
    if re.search(r"\d{4}[-./]\d{1,2}[-./]\d{1,2}",body):sig+=1
    if name.replace(" ","")[:6] in body.replace(" ",""):sig+=1
    return sig>=3,{"reason":"VERIFIED" if sig>=3 else "LOW_SIGNAL","url":canonical(final),"title":title[:200],"signals":sig}
def discover(name,home):
    r,final,html=fetch(home)
    if not r or not html:return name,"","HOME_ERROR",{"homepage":home}
    cand=candidates(final,html); soup=BeautifulSoup(html,"html.parser"); children=[]
    for a in soup.find_all("a",href=True):
        txt=a.get_text(" ",strip=True)
        if not any(w in txt for w in BOARD_WORDS):continue
        u=normalize_url(urljoin(final,a["href"]))
        try:same=urlparse(u).netloc==urlparse(final).netloc
        except Exception:same=False
        if u and same and u not in children:children.append(u)
    for child in children[:6]:
        r2,f2,h2=fetch(child)
        if r2 and h2:cand.extend(candidates(f2,h2))
    d={}
    for x in cand:
        if x[1] and (x[1] not in d or x[0]>d[x[1]][0]):d[x[1]]=x
    ranked=sorted(d.values(),reverse=True)
    for sc,u,a in ranked:
        if "selectboardlist.do" in u.lower() and "bbsid=" in u.lower():
            ok,info=verify(u,name)
            if ok:return name,u,"DISCOVERED_EGOV",info
    for sc,u,a in ranked:
        if sc>=25:
            ok,info=verify(u,name)
            if ok:return name,u,"DISCOVERED_GENERAL",info
    return name,"","CANDIDATE_NOT_VERIFIED",{"candidate_count":len(ranked),"top_candidates":[{"score":x[0],"url":x[1],"anchor":x[2][:100]} for x in ranked[:5]]}

def main():
    cache=load_cache(); targets=load_targets(cache)
    if len(targets)<300:raise RuntimeError(f"기관 데이터가 부족합니다: {len(targets)}개")
    unresolved=[]
    for name,home in targets.items():
        item=cache.get(name,{})
        verified=isinstance(item,dict) and item.get("verified") is True and bool(item.get("board_url") or item.get("url"))
        if not verified:unresolved.append((name,home))
    print(f"[전체 기관] {len(targets)}개")
    print(f"[기존 검증 완료] {len(targets)-len(unresolved)}개")
    print(f"[탐색 대상] {len(unresolved)}개")
    if not unresolved:return
    rows=[]; counts={}; found=0
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures={ex.submit(discover,n,h):(n,h) for n,h in unresolved}
        for i,f in enumerate(concurrent.futures.as_completed(futures),1):
            n,h=futures[f]
            try:name,board,status,info=f.result()
            except Exception as e:name,board,status,info=n,"","DISCOVERY_ERROR",{"reason":type(e).__name__,"message":str(e)}
            counts[status]=counts.get(status,0)+1
            if board:
                old=cache.get(name,{})
                cache[name]={**(old if isinstance(old,dict) else {}),"institution":name,"homepage":h,"board_url":board,"url":board,"verified":True,"discovery_status":status,"discovered_at":NOW.isoformat(),"source":"V8.15.4"};found+=1
            rows.append({"기관명":name,"홈페이지":h,"게시판URL":board,"상태":status,"진단":json.dumps(info,ensure_ascii=False)[:1600]})
            print(f"[{i}/{len(unresolved)}] {name} -> {status}")
    save_cache(cache)
    with open(OUT_PATH,"w",encoding="utf-8-sig",newline="") as f:
        w=csv.DictWriter(f,fieldnames=["기관명","홈페이지","게시판URL","상태","진단"]);w.writeheader();w.writerows(rows)
    d={"version":"V8.15.4","updated_at":NOW.isoformat(),"total_institution_data":len(targets),"previously_verified":len(targets)-len(unresolved),"unresolved_targets":len(unresolved),"newly_found":found,"remaining_unresolved":len(unresolved)-found,"max_workers":MAX_WORKERS,"request_timeout":REQUEST_TIMEOUT,"status_counts":counts}
    with open(LOG_PATH,"w",encoding="utf-8") as f:json.dump(d,f,ensure_ascii=False,indent=2)
    print(json.dumps(d,ensure_ascii=False,indent=2))
if __name__=="__main__":
    try:main()
    except Exception as e:print(f"FATAL: {type(e).__name__}: {e}",file=sys.stderr);sys.exit(1)
