import os, re, json, time, asyncio, html
from urllib.parse import urljoin, urlparse
import aiohttp
from bs4 import BeautifulSoup
import openpyxl

KEYWORDS = ["설문조사", "시민참여", "국민참여", "공모전"]
STATE_FILE = "state.json"
LOG_FILE = "monitor_log.json"

TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "12"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "600"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))

def load_targets():
    wb = openpyxl.load_workbook("monitor_targets.xlsx", data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    name_i = headers.index("기관명")
    url_i = headers.index("URL")
    rows=[]
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[url_i]:
            rows.append({"institution": str(r[name_i] or "").strip(),
                         "board": str(r[url_i]).strip()})
    return rows

def normalize_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            s=json.load(f)
    except Exception:
        s={}
    if isinstance(s, list):
        s={"version":862, "initialized": True, "seen": {str(x): True for x in s}}
    if not isinstance(s, dict):
        s={}
    if not isinstance(s.get("seen"), dict):
        s["seen"]={}
    s["version"]=8621
    return s

def clean_text(s):
    s=html.unescape(s or "")
    return re.sub(r"\s+", " ", s).strip()

def is_listing_url(url):
    p=urlparse(url)
    q=p.query.lower()
    path=p.path.lower()
    bad=["list","search","page","offset","boardlist","selectnoticelist"]
    return any(x in q or x in path for x in bad) and "view" not in q and "article" not in q

def extract_candidates(base_url, soup):
    base_host=urlparse(base_url).netloc
    out=[]
    seen=set()
    for a in soup.find_all("a"):
        attrs = getattr(a, "attrs", None) or {}
        href = str(attrs.get("href") or "").strip()
        if not href:
            continue
        text=clean_text(a.get_text(" ", strip=True))
        if not href or href.startswith(("javascript:", "#","mailto:")):
            continue
        u=urljoin(base_url, href)
        pu=urlparse(u)
        if pu.netloc != base_host:
            continue
        if is_listing_url(u):
            continue
        if len(text) < 2 or len(text) > 250:
            continue
        # likely detail links: query has article/id/no/seq/view or href contains view/detail
        marker=(pu.query+" "+pu.path).lower()
        if not re.search(r"(article(no)?|board(no)?|seq|ntt|bbs|idx|no=|view|detail|read|viewpage)", marker):
            continue
        key=(u,text)
        if key not in seen:
            seen.add(key); out.append({"url":u,"title":text})
    return out[:RECENT_POSTS]

def likely_content(soup):
    # Remove site-wide and non-content areas first.
    for tag in soup(["script","style","noscript","template","svg","header","footer","nav","aside","form"]):
        tag.decompose()
    # Remove obvious lists/menus/related/search areas.
    for tag in soup.find_all(["ul","ol"]):
        attrs = getattr(tag, "attrs", None) or {}
        cls_val = attrs.get("class", [])
        if isinstance(cls_val, str):
            cls_val = [cls_val]
        cls=" ".join(str(x) for x in cls_val).lower()
        tid=str(attrs.get("id") or "").lower()
        if any(x in (cls+" "+tid) for x in ["menu","nav","gnb","lnb","related","recommend","search","list"]):
            tag.decompose()
    selectors=[
        "[class*='view']", "[class*='content']", "[class*='detail']",
        "[id*='view']", "[id*='content']", "[id*='detail']",
        "article", "main"
    ]
    candidates=[]
    for sel in selectors:
        for x in soup.select(sel):
            t=clean_text(x.get_text(" ", strip=True))
            if len(t)>=40:
                candidates.append((len(t),t))
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]
    return clean_text(soup.get_text(" ", strip=True))

def keyword_hits(title, body):
    # Title/body are the only searchable content. Exclude page chrome by extraction.
    hits=[]
    for kw in KEYWORDS:
        if kw in title or kw in body:
            hits.append(kw)
    return hits

async def fetch(session, url):
    last=None
    for attempt in range(HTTP_RETRIES+1):
        try:
            timeout=aiohttp.ClientTimeout(total=TIMEOUT_SECONDS)
            async with session.get(url, timeout=timeout, allow_redirects=True,
                                   headers={"User-Agent":"Mozilla/5.0 public-institution-monitor/8.6.2"}) as r:
                if r.status >= 400:
                    last=f"HTTP {r.status}"
                else:
                    return r.status, await r.text(errors="ignore")
        except Exception as e:
            last=f"HTTP 0"
        if attempt < HTTP_RETRIES:
            await asyncio.sleep(0.7*(attempt+1))
    return 0, ""

async def process_target(session, sem, target, state):
    async with sem:
        try:
            return await _process_target_inner(session, target, state)
        except Exception as e:
            return {"institution":target["institution"],"board":target["board"],"url":target["board"],
                    "posts":0,"matches":[],"error":f"{type(e).__name__}: {e}"}

async def _process_target_inner(session, target, state):
        board=target["board"]
        status, raw=await fetch(session, board)
        if not raw:
            return {"institution":target["institution"],"board":board,"url":board,
                    "posts":0,"matches":[],"error":f"HTTP {status}"}
        soup=BeautifulSoup(raw,"html.parser")
        candidates=extract_candidates(board,soup)
        matches=[]
        checked=0
        for c in candidates:
            status2, raw2=await fetch(session,c["url"])
            if not raw2:
                continue
            checked+=1
            ds=BeautifulSoup(raw2,"html.parser")
            title=clean_text(ds.title.get_text(" ",strip=True) if ds.title else c["title"])
            # Prefer link text when page title is generic.
            if c["title"] and len(c["title"]) >= 2 and not re.search(r"(홈|공지사항|목록|사이트)", c["title"]):
                if len(c["title"]) <= 180:
                    title=c["title"]
            body=likely_content(ds)
            hits=keyword_hits(title,body)
            if not hits:
                continue
            key=f"{target['institution']}|{c['url']}"
            if key not in state["seen"]:
                matches.append({"institution":target["institution"],"board":board,
                                "keyword":hits,"title":title,"url":c["url"]})
        return {"institution":target["institution"],"board":board,"url":board,
                "posts":checked,"matches":matches,"error":None}

async def main_async():
    targets=load_targets()
    state=normalize_state()
    initialized=bool(state.get("initialized"))
    state["initialized"]=True
    sem=asyncio.Semaphore(MAX_CONCURRENCY)
    connector=aiohttp.TCPConnector(limit=MAX_CONCURRENCY, ssl=False)
    started=time.time()
    results=[]
    async with aiohttp.ClientSession(connector=connector) as session:
        tasks=[asyncio.create_task(process_target(session,sem,t,state)) for t in targets]
        try:
            results=await asyncio.wait_for(asyncio.gather(*tasks, return_exceptions=False),
                                           timeout=MAX_TOTAL_SECONDS)
        except asyncio.TimeoutError:
            for t in tasks:
                if not t.done(): t.cancel()
            results=[x for x in await asyncio.gather(*tasks, return_exceptions=True) if isinstance(x,dict)]
    new_matches=[]
    errors=[]
    posts=0
    for r in results:
        posts += r.get("posts",0)
        new_matches.extend(r.get("matches",[]))
        if r.get("error"): errors.append(r)
    # On first run, seed matches but don't alert.
    if not initialized:
        for m in new_matches:
            state["seen"][f"{m['institution']}|{m['url']}"]=True
        alert_matches=[]
    else:
        alert_matches=new_matches[:TELEGRAM_MAX_SEND]
        for m in new_matches:
            state["seen"][f"{m['institution']}|{m['url']}"]=True
    state["seen"]=dict(list(state["seen"].items())[-10000:])
    log={"version":"8.6.2","timestamp":time.strftime("%Y-%m-%dT%H:%M:%S"),
         "targets":len(targets),"completed":len(results),"posts_checked":posts,
         "new_matches":len(new_matches),"errors":len(errors),
         "timed_out": len(results)<len(targets),"first_run":not initialized,
         "matches":new_matches,"error_details":errors}
    tg=await send_telegram(alert_matches)
    log["telegram"]=tg
    with open(STATE_FILE,"w",encoding="utf-8") as f: json.dump(state,f,ensure_ascii=False,indent=2)
    with open(LOG_FILE,"w",encoding="utf-8") as f: json.dump(log,f,ensure_ascii=False,indent=2)
    print(json.dumps(log,ensure_ascii=False,indent=2))

async def send_telegram(matches):
    token=os.getenv("TELEGRAM_BOT_TOKEN","").strip()
    chat=os.getenv("TELEGRAM_CHAT_ID","").strip()
    if not matches:
        return {"ok":True,"sent":0,"attempted":0,"errors":[],"truncated":False}
    if not token or not chat:
        return {"ok":False,"sent":0,"attempted":0,"errors":["missing Telegram secret"],"truncated":False}
    url=f"https://api.telegram.org/bot{token}/sendMessage"
    sent=0; errs=[]
    async with aiohttp.ClientSession() as s:
        for m in matches:
            text=f"🔔 {m['institution']}\n키워드: {', '.join(m['keyword'])}\n{m['title']}\n{m['url']}"
            try:
                async with s.post(url,json={"chat_id":chat,"text":text},timeout=15) as r:
                    data=await r.json(content_type=None)
                    if data.get("ok"):
                        sent+=1
                    else:
                        errs.append(str(data.get("description","Telegram API error")))
            except Exception as e:
                errs.append(str(e))
    return {"ok":len(errs)==0,"sent":sent,"attempted":len(matches),
            "errors":errs,"truncated":len(matches)>TELEGRAM_MAX_SEND}

if __name__=="__main__":
    asyncio.run(main_async())
