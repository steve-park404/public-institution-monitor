import os, re, json, time, html
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, parse_qs
import requests
from bs4 import BeautifulSoup

KEYWORDS = ["설문조사", "시민참여", "국민참여"]
EXCLUDE_KEYWORDS = ["공모전"]
UA = "Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.8; +https://github.com/steve-park404/public-institution-monitor)"

MAX_CONCURRENCY = int(os.getenv("MAX_CONCURRENCY", "20"))
TIMEOUT_SECONDS = int(os.getenv("TIMEOUT_SECONDS", "15"))
RECENT_POSTS = int(os.getenv("RECENT_POSTS", "15"))
MAX_TOTAL_SECONDS = int(os.getenv("MAX_TOTAL_SECONDS", "1200"))
HTTP_RETRIES = int(os.getenv("HTTP_RETRIES", "2"))
TELEGRAM_MAX_SEND = int(os.getenv("TELEGRAM_MAX_SEND", "20"))

NOISE_TAGS = ["script","style","noscript","svg","header","footer","nav","aside","form","iframe","canvas","template"]
BOARD_WORDS = ["공지사항","공지","알림마당","알림","소식","새소식","기관소식","게시판","뉴스","보도자료","참여","국민참여","시민참여"]
URL_HINTS = ["notice","noti","board","bbs","news","announcement","community","plaza","inform","particip"]
BAD_WORDS = ["채용","입찰","계약","로그인","회원","사이트맵","개인정보","이용약관","검색"]

session = requests.Session()
session.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.5"})

def norm(s):
    return re.sub(r"\s+", " ", html.unescape(str(s or ""))).strip()

def get(url):
    for i in range(HTTP_RETRIES + 1):
        try:
            r = session.get(url, timeout=TIMEOUT_SECONDS, allow_redirects=True)
            r.raise_for_status()
            if not r.encoding or r.encoding.lower() == "iso-8859-1":
                r.encoding = r.apparent_encoding
            return r
        except Exception:
            if i < HTTP_RETRIES:
                time.sleep(0.4 * (i + 1))
    return None

def same_domain(a, b):
    try:
        return urlparse(a).netloc.lower().replace("www.","") == urlparse(b).netloc.lower().replace("www.","")
    except Exception:
        return False

def clean_soup(soup):
    for tag in soup(NOISE_TAGS):
        tag.decompose()
    return soup

def page_text(soup):
    return norm(clean_soup(soup).get_text(" ", strip=True))

def title_of(soup):
    t = soup.title.get_text(" ", strip=True) if soup.title else ""
    return norm(t)

def is_probable_list_url(url):
    u = url.lower()
    p = urlparse(u)
    path = p.path
    if any(x in u for x in ["sitemap","robots.txt","login","member","search?","/search/"]):
        return False
    q = parse_qs(p.query)
    if any(k.lower() in {"page","pageindex","pageidx","offset","start","rows"} for k in q):
        if not any(k.lower() in {"seq","no","idx","nttno","articleid","article_id","id"} for k in q):
            return True
    return True

def has_detail_signal(url):
    u = url.lower()
    p = urlparse(u)
    q = parse_qs(p.query)
    if any(k.lower() in {"seq","no","idx","nttno","articleid","article_id","id","view"} for k in q):
        return True
    if any(x in u for x in ["/view","/detail","/read","/article","/contents/view","/board/view"]):
        return True
    return False

def is_probable_detail_url(url):
    return has_detail_signal(url) and not any(x in url.lower() for x in ["login","member","delete","write","modify"])

def link_text(a):
    return norm(a.get_text(" ", strip=True))

def extract_links(page_url, soup):
    out = []
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a.get("href","").strip()
        if not href or href.startswith(("javascript:","mailto:","tel:","#")):
            continue
        u = urljoin(page_url, href)
        if not u.startswith(("http://","https://")) or not same_domain(page_url,u):
            continue
        if u in seen:
            continue
        seen.add(u)
        txt = link_text(a)
        out.append((u, txt))
    return out

def score_board_link(base_url, url, txt):
    s = 0
    t = txt.lower()
    u = url.lower()
    for w in BOARD_WORDS:
        if w.lower() in t: s += 7
    for h in URL_HINTS:
        if h in u: s += 3
    if any(b.lower() in t for b in BAD_WORDS): s -= 5
    if has_detail_signal(url): s -= 4
    if len(txt) > 60: s -= 2
    return s

def discover_board(home_url):
    r = get(home_url)
    if not r:
        return None, [], "homepage_error"
    soup = BeautifulSoup(r.text, "html.parser")
    links = extract_links(r.url, soup)
    scored = sorted([(score_board_link(r.url,u,t),u,t) for u,t in links], reverse=True)
    # 상위 후보를 실제로 열어 게시물 링크가 있는지 확인
    checked = 0
    for score,u,t in scored[:30]:
        if score < 4:
            break
        rr = get(u)
        checked += 1
        if not rr:
            continue
        ss = BeautifulSoup(rr.text, "html.parser")
        details = [x for x,tx in extract_links(rr.url, ss) if is_probable_detail_url(x)]
        unique = list(dict.fromkeys(details))
        if len(unique) >= 2:
            return rr.url, unique[:RECENT_POSTS], f"discovered:{score}"
    return None, [], f"no_board:{checked}"

def extract_high_confidence_content(soup):
    for tag in NOISE_TAGS:
        for x in soup.find_all(tag):
            x.decompose()
    main = soup.find("main") or soup.find("article") or soup.find(id=re.compile(r"(content|contents|sub|body)", re.I))
    text = norm(main.get_text(" ", strip=True) if main else soup.get_text(" ", strip=True))
    # 공통 메뉴/헤더성 짧은 반복 텍스트를 지나치게 신뢰하지 않도록 길이 제한
    return text

def match_post(url):
    r = get(url)
    if not r:
        return None
    soup = BeautifulSoup(r.text, "html.parser")
    title = title_of(soup)
    # 제목 우선
    if title and not any(x in title for x in EXCLUDE_KEYWORDS):
        for kw in KEYWORDS:
            if kw in title:
                return {"url": r.url, "title": title, "keyword": kw, "where": "title"}
    body = extract_high_confidence_content(soup)
    if len(body) < 80:
        return None
    if any(x in body for x in EXCLUDE_KEYWORDS):
        # '공모전'이 본문에 단순 언급되는 경우까지 무조건 배제하지 않고,
        # 제목이 공모전이 아니면 키워드 매칭은 허용하되 제목 우선 정책을 유지한다.
        pass
    for kw in KEYWORDS:
        if kw in body:
            return {"url": r.url, "title": title, "keyword": kw, "where": "body"}
    return None

def load_targets():
    import openpyxl
    wb = openpyxl.load_workbook("monitor_targets.xlsx", read_only=True, data_only=True)
    ws = wb["355기관"]
    headers = [c.value for c in next(ws.iter_rows())]
    idx = {str(v):i for i,v in enumerate(headers)}
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["기관명"]]:
            out.append({
                "name": str(row[idx["기관명"]]).strip(),
                "home": str(row[idx["홈페이지URL"]] or "").strip(),
                "board": str(row[idx["공지게시판URL"]] or "").strip(),
            })
    return out[:355]

def load_state():
    if not os.path.exists("state.json"):
        return {"seen": {}, "boards": {}}
    try:
        with open("state.json","r",encoding="utf-8") as f:
            s=json.load(f)
        if isinstance(s, list):
            return {"seen": {str(x):1 for x in s[-10000:]}, "boards": {}}
        if not isinstance(s, dict):
            return {"seen": {}, "boards": {}}
        if not isinstance(s.get("seen"), dict):
            s["seen"] = {str(x):1 for x in s.get("seen", [])[-10000:]}
        if not isinstance(s.get("boards"), dict):
            s["boards"] = {}
        return s
    except Exception:
        return {"seen": {}, "boards": {}}

def save_state(s):
    s["seen"] = dict(list(s.get("seen",{}).items())[-10000:])
    with open("state.json","w",encoding="utf-8") as f:
        json.dump(s,f,ensure_ascii=False,indent=2)

def process_target(t, state, deadline):
    name, home, seed_board = t["name"], t["home"], t["board"]
    if time.time() >= deadline:
        return {"name":name,"status":"deadline","posts":0,"matches":[],"board":""}
    board = seed_board or state.get("boards",{}).get(name,"")
    details = []
    source = "cached"
    if board:
        r = get(board)
        if r:
            soup = BeautifulSoup(r.text, "html.parser")
            details = list(dict.fromkeys([u for u,tx in extract_links(r.url,soup) if is_probable_detail_url(u)]))[:RECENT_POSTS]
        else:
            board = ""
    if not board and home:
        board, details, source = discover_board(home)
        if board:
            state.setdefault("boards",{})[name] = board
    if not board:
        return {"name":name,"status":"no_board","posts":0,"matches":[],"board":""}
    matches=[]
    for u in details[:RECENT_POSTS]:
        if time.time() >= deadline:
            break
        m=match_post(u)
        if not m:
            continue
        key=m["url"]+"|"+m["keyword"]
        if key not in state["seen"]:
            state["seen"][key]=1
            m["기관명"]=name
            matches.append(m)
    return {"name":name,"status":"ok","posts":len(details),"matches":matches,"board":board,"source":source}

def send_telegram(items):
    token=os.getenv("TELEGRAM_BOT_TOKEN","").strip()
    chat=os.getenv("TELEGRAM_CHAT_ID","").strip()
    if not token or not chat:
        print("TELEGRAM_CONFIG_MISSING")
        return 0
    sent=0
    for m in items[:TELEGRAM_MAX_SEND]:
        text=(f"📢 공공기관 참여정보 알림\n"
              f"기관: {m['기관명']}\n"
              f"키워드: {m['keyword']}\n"
              f"제목: {m['title']}\n"
              f"검색위치: {m['where']}\n"
              f"{m['url']}")
        try:
            rr=requests.post(
                f"https://api.telegram.org/bot{token}/sendMessage",
                json={"chat_id":chat,"text":text},
                timeout=20
            )
            if rr.ok and rr.json().get("ok"):
                sent += 1
            else:
                print("TELEGRAM_ERROR", rr.status_code, rr.text[:300])
        except Exception as e:
            print("TELEGRAM_EXCEPTION", repr(e))
    print("TELEGRAM_SENT", sent)
    return sent

def main():
    start=time.time()
    deadline=start+MAX_TOTAL_SECONDS
    targets=load_targets()
    state=load_state()
    results=[]
    matches=[]
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as ex:
        futs=[ex.submit(process_target,t,state,deadline) for t in targets]
        for f in as_completed(futs):
            try:
                r=f.result()
            except Exception as e:
                r={"name":"?","status":"error","posts":0,"matches":[],"error":repr(e)}
            results.append(r)
            matches.extend(r.get("matches",[]))
    save_state(state)
    sent=send_telegram(matches)
    completed=sum(r.get("status")=="ok" for r in results)
    posts=sum(r.get("posts",0) for r in results)
    no_board=sum(r.get("status")=="no_board" for r in results)
    errors=sum(r.get("status")=="error" for r in results)
    summary={
        "targets":len(targets),
        "completed":completed,
        "no_board":no_board,
        "posts_checked":posts,
        "new_matches":len(matches),
        "telegram_sent":sent,
        "errors":errors,
        "elapsed_seconds":round(time.time()-start,1),
        "timed_out":time.time()>=deadline
    }
    with open("monitor_log.json","w",encoding="utf-8") as f:
        json.dump({"summary":summary,"results":results,"matches":matches[-100:]},f,ensure_ascii=False,indent=2)
    print("SUMMARY",json.dumps(summary,ensure_ascii=False))

if __name__=="__main__":
    main()
