# -*- coding: utf-8 -*-
"""
Public Institution Monitor V8.13.0
Precision-first actual-post verification.

Architecture:
355기관 → 공식 홈페이지 → 게시판 → 최근 30일 후보 → 상세페이지
→ 실제 게시물 구조 검증 → 제목/본문 키워드 → Telegram

Operational keywords:
- 설문조사
- 시민참여
- 국민참여

Title exclusion:
- 공모전
"""

VERSION = "V8.13.0"

import os, re, json, time, html, warnings
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import openpyxl

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

KEYWORDS = ["설문조사", "시민참여", "국민참여"]
EXCLUDE_TITLE = ["공모전"]

MAX_CONCURRENCY = 20
TIMEOUT_SECONDS = 15
HTTP_RETRIES = 2
RECENT_POSTS = 20
RECENT_DAYS = 30
MAX_TOTAL_SECONDS = 1200
TELEGRAM_MAX_SEND = 20
MAX_PENDING = 10000

BOARD_DISCOVERY_MAX_LINKS = 100
BOARD_DISCOVERY_MAX_FETCH = 35

UA = f"Mozilla/5.0 (compatible; PublicInstitutionMonitor/{VERSION})"
KST = ZoneInfo("Asia/Seoul")

STATE_FILE = "state.json"
PENDING_FILE = "pending.json"
TARGET_FILE_CANDIDATES = ["monitor_targets.xlsx", "url.xlsx", "targets.xlsx"]

NOISE = ["script","style","noscript","svg","header","footer","nav","aside",
         "form","iframe","canvas","template"]

FIRST_PRIORITY_BOARD_WORDS = [
    "공지사항","공지","새소식","알림마당","알림","이벤트",
    "설문","설문조사","국민참여","시민참여","참여마당","소통"
]
BOARD_WORDS = FIRST_PRIORITY_BOARD_WORDS + [
    "소식","기관소식","게시판","뉴스","보도자료","공고"
]
URL_HINTS = [
    "notice","noti","board","bbs","news","announcement","plaza",
    "inform","ntt","article","list"
]
HARD_EXCLUDE_BOARD_WORDS = [
    "채용","입찰","계약","자료실","교육","공모전","동반성장",
    "사회공헌","구매","구매계약"
]

GENERIC_PAGE_TITLES = {
    "사이트맵","사이트 맵","알림마당","공지사항","공지","새소식",
    "사업소개","주요사업","연구","국민소통","시민참여","국민참여",
    "참여마당","개인정보처리방침","개인정보 처리방침","이용약관",
    "로그인","회원가입","참여소통","기관소개","홈"
}
GENERIC_URL_HINTS = [
    "sitemap","privacy","terms","login","programproposal","consulting",
    "issuedata","homepage"
]

PARTICIPATION_CONTEXT = [
    "설문","의견수렴","의견조사","만족도","조사","응답","설문지",
    "참여단","시민의견","국민의견","의견","참여기간","응답기간",
    "조사기간","참여방법","응답방법","참여해 주세요","응답해 주세요",
    "설문에 참여","조사에 참여","의견을 제출","의견을 남겨",
    "설문링크","조사대상","응답자","참여자"
]
SURVEY_ACTION_CONTEXT = [
    "설문","조사","응답","만족도","의견수렴","의견조사","설문지",
    "참여기간","응답기간","조사기간","참여방법","응답방법",
    "참여해 주세요","응답해 주세요","설문에 참여","조사에 참여",
    "의견을 제출","의견을 남겨","설문링크","조사대상","응답자","참여자"
]
SENTENCE_MARKERS = [
    "안내","참여","신청","설문","의견","조사","응답","만족도",
    "기간","대상","방법","모집","제출","온라인","링크"
]

DATE_PATTERNS = [
    r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})",
    r"(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일",
    r"(20\d{2})(\d{2})(\d{2})"
]
DATE_SELECTORS = [
    ".date",".regdate",".reg-date",".write-date",".wdate",".board-date",
    ".bbs-date",".ntt-date",".article-date","[class*='date']",
    "[class*='Date']","[class*='regist']","[class*='Regist']",
    "[class*='write']","[class*='Write']","time","td"
]

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": UA,
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7",
})

def get(url):
    if not url:
        return None
    for attempt in range(HTTP_RETRIES + 1):
        try:
            r = SESSION.get(url, timeout=TIMEOUT_SECONDS,
                            allow_redirects=True, verify=True)
            if 200 <= r.status_code < 400:
                if not r.encoding or r.encoding.lower() == "iso-8859-1":
                    r.encoding = r.apparent_encoding or "utf-8"
                return r
        except Exception:
            pass
        if attempt < HTTP_RETRIES:
            time.sleep(0.35 * (attempt + 1))
    return None

def norm(s):
    return re.sub(r"\s+", " ", html.unescape(str(s or ""))).strip()

def lower_url(u): return (u or "").lower()

def same_domain(a,b):
    try:
        da=urlparse(a).netloc.lower().split(":")[0]
        db=urlparse(b).netloc.lower().split(":")[0]
        return da.removeprefix("www.") == db.removeprefix("www.")
    except Exception:
        return False

def absolute(base, href):
    try: return urljoin(base, href)
    except Exception: return ""

def clean_url(u):
    try: return urlparse(u)._replace(fragment="").geturl()
    except Exception: return u

def text_of(tag):
    return norm(tag.get_text(" ", strip=True)) if tag else ""

def parse_date_text(s):
    s=norm(s)
    if not s: return None
    for pat in DATE_PATTERNS:
        m=re.search(pat,s)
        if not m: continue
        try:
            y,mo,d=map(int,m.groups())
            if 2000 <= y <= 2099 and 1 <= mo <= 12 and 1 <= d <= 31:
                return datetime(y,mo,d,tzinfo=KST)
        except Exception:
            pass
    return None

def recent_cutoff():
    return datetime.now(KST)-timedelta(days=RECENT_DAYS)

def is_recent_date(dt):
    return bool(dt and recent_cutoff() <= dt <= datetime.now(KST)+timedelta(days=1))

def is_list_only_url(url):
    try:
        p=urlparse(url)
        path=(p.path or "").rstrip("/").lower()
        q=(p.query or "").lower()
        full=path+"?"+q
        if path in {"/list","/lists","/index","/events","/event","/notice",
                    "/notices","/news","/board","/bbs"} and not q:
            return True
        if re.search(r"(^|&)(page|pageindex|pageno|page_no|mode=list|act=list)=[^&]*",q):
            return True
        if re.search(r"/(?:notice|notices|news|board|bbs|list|announcement|announcements)(?:/index)?$",path):
            return True
        if re.search(r"/(?:selectnttlist|selectbbslist|boardlist|selectboardlist|list)\.do$",path):
            return True
        if re.search(r"(?:selectnttlist|selectbbslist|boardlist|selectboardlist|mode=list|act=list)",full):
            return True
    except Exception:
        pass
    return False

def generic_url(url):
    p=lower_url(url)
    return any(x in p for x in GENERIC_URL_HINTS)

def clean_container(tag):
    clone=BeautifulSoup(str(tag),"html.parser")
    for x in clone.find_all(["header","nav","footer","aside","form","script","style","noscript"]):
        x.decompose()
    # Remove navigation-like micro blocks.
    for x in clone.find_all(["div","section","ul","ol","li","p","dl"]):
        tx=norm(x.get_text(" ",strip=True))
        if tx and len(tx)<=500 and re.search(
            r"^(이전글|다음글|목록|공유|인쇄|관련글|관련사이트|첨부파일|만족도|스크랩)$",tx):
            x.decompose()
    return text_of(clone)

def visible_main_text(soup):
    if not soup: return ""
    for tag in soup.find_all(NOISE):
        tag.decompose()

    selectors=[
        ".view-content",".view_cont",".view-cont",".board-content",".board_cont",
        ".board-view-content",".bbs-content",".bbs_cont",".article-content",
        ".article_cont",".contents-view",".content-view",".nttCn",".ntt_cn",
        "[class*='view'][class*='content']","[class*='board'][class*='content']",
        "[class*='bbs'][class*='content']","[class*='article'][class*='content']"
    ]
    candidates=[]
    for sel in selectors:
        try: tags=soup.select(sel)[:10]
        except Exception: tags=[]
        for tag in tags:
            t=clean_container(tag)
            if 80 <= len(t) <= 30000:
                candidates.append(t)
    if candidates:
        # Prefer compact dedicated content over huge page shells.
        candidates.sort(key=lambda x:(len(x)>18000, len(x)>12000, -len(x)))
        return candidates[0]

    for root_sel in ["article","main"]:
        candidates=[]
        for root in soup.select(root_sel):
            t=clean_container(root)
            if 100 <= len(t) <= 30000:
                candidates.append(t)
        if candidates:
            candidates.sort(key=lambda x:(len(x)>18000,-len(x)))
            return candidates[0]
    return ""

def extract_title(soup):
    if not soup: return ""
    def valid(t):
        t=norm(t)
        if not (2 <= len(t) <= 300): return False
        if t in GENERIC_PAGE_TITLES: return False
        if len(re.sub(r"[^0-9A-Za-z가-힣]","",t)) < 3: return False
        return True

    selectors=[
        ".view-title",".board-title",".article-title",".bbs-title",
        ".board_view .subject",".boardView .subject",".view .subject",
        ".view_subject",".viewSubject",".post-title",".post_title",
        ".nttTitle",".ntt-title",".subject",".bbs-subject",
        "[class*='view'][class*='title']","[class*='board'][class*='title']",
        "[class*='article'][class*='title']","[class*='post'][class*='title']"
    ]
    vals=[]
    for sel in selectors:
        try: tags=soup.select(sel)[:10]
        except Exception: tags=[]
        for tag in tags:
            t=text_of(tag)
            if valid(t): vals.append(t)
    if vals:
        # Avoid tiny menu labels; prefer plausible post titles.
        vals.sort(key=lambda x:(len(x)<5, len(x)>180, -len(x)))
        return vals[0]

    for sel in ["meta[property='og:title']","meta[name='twitter:title']"]:
        og=soup.select_one(sel)
        if og and og.get("content"):
            parts=[norm(x) for x in re.split(r"\s*[|｜]\s*",og.get("content")) if norm(x)]
            for p in parts:
                if valid(p): return p[:300]

    if soup.title:
        parts=[norm(x) for x in re.split(r"\s*[|｜]\s*",soup.title.get_text(" ",strip=True)) if norm(x)]
        # Prefer the longest plausible segment, not the institution name.
        vals=[p for p in parts if valid(p)]
        if vals:
            return sorted(vals,key=len,reverse=True)[0][:300]
    return ""

def extract_detail_date(soup):
    if not soup: return None
    for tag in soup.find_all(["meta","time"])[:100]:
        attrs=" ".join(str(v) for v in tag.attrs.values())
        content=tag.get("content") or tag.get("datetime") or tag.get_text(" ",strip=True)
        dt=parse_date_text(f"{attrs} {content}")
        if dt: return dt
    for sel in DATE_SELECTORS:
        try: tags=soup.select(sel)
        except Exception: tags=[]
        for tag in tags[:30]:
            dt=parse_date_text(text_of(tag))
            if dt: return dt
    txt=visible_main_text(soup)
    return parse_date_text(txt[:5000])

def has_post_structure(soup,title):
    if not soup or not title: return False
    txt=norm(soup.get_text(" ",strip=True))[:12000]
    meta=bool(re.search(r"(작성자|작성일|등록일|조회수|등록자|게시일)",txt))
    content=bool(visible_main_text(soup))
    title_len=2 <= len(title) <= 300
    return title_len and content and (meta or len(visible_main_text(soup)) >= 160)

def looks_like_list_dom(soup):
    if not soup: return True
    raw=norm(soup.get_text(" ",strip=True))[:20000]
    signals=0
    if re.search(r"(공지사항 리스트|게시물 리스트|목록|총\s*\d+\s*건)",raw): signals+=1
    if len(soup.select("table tbody tr")) >= 3: signals+=1
    if len(soup.select("a[href*='list_no'], a[href*='article_no'], a[href*='nttId']")) >= 3: signals+=1
    if signals >= 2 and not re.search(r"(작성자|작성일|등록일|조회수)",raw[:10000]):
        return True
    return False

def meaningful_body_match(body,keyword):
    body=norm(body)
    if len(body)<120: return False
    for m in re.finditer(re.escape(keyword),body,re.I):
        idx=m.start()
        ctx=body[max(0,idx-350):min(len(body),idx+500)]
        if keyword in ("국민참여","시민참여"):
            if not any(x in ctx for x in PARTICIPATION_CONTEXT): continue
        if keyword=="설문조사":
            if not any(x in ctx for x in SURVEY_ACTION_CONTEXT): continue
        if sum(x in ctx for x in SENTENCE_MARKERS)>=1:
            return True
    return False

def extract_links(base,soup):
    out=[]; seen=set()
    if not soup: return out
    for a in soup.find_all("a",href=True):
        href=a.get("href","").strip()
        if not href or href.lower().startswith(("javascript:","#","mailto:","tel:")): continue
        u=clean_url(absolute(base,href))
        if not u or not same_domain(base,u): continue
        label=norm(" ".join([text_of(a),a.get("title",""),a.get("aria-label","")]))
        if (u,label) in seen: continue
        seen.add((u,label)); out.append((u,label))
        if len(out)>=BOARD_DISCOVERY_MAX_LINKS: break
    return out

def board_score(u,title):
    s=f"{u} {title}".lower(); score=0
    for w in FIRST_PRIORITY_BOARD_WORDS:
        if w.lower() in s: score+=18
    for w in BOARD_WORDS:
        if w.lower() in s: score+=7
    for h in URL_HINTS:
        if h in lower_url(u): score+=3
    for x in HARD_EXCLUDE_BOARD_WORDS:
        if x.lower() in s: score-=14
    return score

def looks_like_detail_link(u,title):
    s=f"{u} {title}".lower()
    if any(x in s for x in ["login","logout","sitemap","privacy","terms"]): return False
    if is_list_only_url(u): return False
    if any(x in lower_url(u) for x in [
        "view","read","detail","article","ntt","bbs","board",
        "idx=","seq=","no=","mode=view","view.do","read.do"
    ]): return True
    return 2 <= len(norm(title)) <= 300

def extract_post_candidates(base,soup):
    out=[]; seen=set()
    if not soup: return out

    for row in soup.find_all("tr"):
        links=row.find_all("a",href=True)
        if not links: continue
        row_date=parse_date_text(text_of(row))
        for a in links:
            title=norm(" ".join([text_of(a),a.get("title",""),a.get("aria-label","")]))
            u=clean_url(absolute(base,a.get("href","")))
            if not u or not same_domain(base,u) or not looks_like_detail_link(u,title): continue
            if u in seen: continue
            seen.add(u); out.append({"url":u,"title":title,"date":row_date})

    for a in soup.find_all("a",href=True):
        title=norm(" ".join([text_of(a),a.get("title",""),a.get("aria-label","")]))
        u=clean_url(absolute(base,a.get("href","")))
        if not u or not same_domain(base,u) or not looks_like_detail_link(u,title): continue
        if len(title)<2: continue
        if u in seen: continue
        seen.add(u)
        out.append({"url":u,"title":title,"date":parse_date_text(text_of(a.parent))})
        if len(out)>=80: break
    return out

def recent_detail_urls(base,soup):
    c=extract_post_candidates(base,soup)
    if not c: return []
    dated=[x for x in c if x.get("date")]
    if dated:
        dated.sort(key=lambda x:x["date"],reverse=True)
        recent=[x for x in dated if is_recent_date(x["date"])]
        return (recent or dated)[:RECENT_POSTS]
    return c[:RECENT_POSTS]

def inspect_board(board_url):
    r=get(board_url)
    if not r:
        return None,[],{"status":"BOARD_FETCH_ERROR","url":board_url,"candidate_posts":0}
    soup=BeautifulSoup(r.text,"html.parser")
    c=recent_detail_urls(r.url,soup)
    return r.url,c,{"status":"VERIFIED" if c else "NO_RECENT_CANDIDATE",
                   "url":r.url,"candidate_posts":len(c),
                   "recent_titles":" ".join(x.get("title","") for x in c[:10])[:1000]}

def discover_board(home):
    diag={"home":home,"status":"START","candidate_count":0,"fetched_count":0,
          "verified_count":0,"selected":"","selected_score":None,"candidates":[]}
    r=get(home)
    if not r:
        diag["status"]="HOME_ERROR"; return None,[],"HOME_ERROR",diag
    soup=BeautifulSoup(r.text,"html.parser")
    links=extract_links(r.url,soup)
    scored=sorted([(board_score(u,t),u,t) for u,t in links],reverse=True)
    diag["candidate_count"]=len(scored)
    if not scored:
        diag["status"]="NO_CANDIDATE"; return None,[],"NO_CANDIDATE",diag

    best=None
    for score,u,t in scored[:BOARD_DISCOVERY_MAX_FETCH]:
        if score < -8: continue
        diag["fetched_count"]+=1
        try:
            bu,details,info=inspect_board(u)
        except Exception as e:
            diag["candidates"].append({"score":score,"url":u,"title":t[:120],
                                       "status":"INSPECT_ERROR","error_type":type(e).__name__})
            continue
        diag["candidates"].append({"score":score,"url":u,"title":t[:120],
                                   "status":info.get("status"),
                                   "candidate_posts":info.get("candidate_posts",0)})
        if info.get("status")!="VERIFIED": continue
        diag["verified_count"]+=1
        combined=score+min(info.get("candidate_posts",0),10)
        rt=info.get("recent_titles","")
        if any(k in rt for k in KEYWORDS): combined+=2
        if best is None or combined>best["score"]:
            best={"score":combined,"url":bu,"details":details}
    if best:
        diag["status"]="VERIFIED"; diag["selected"]=best["url"]; diag["selected_score"]=best["score"]
        return best["url"],best["details"][:RECENT_POSTS],"DISCOVERED",diag
    diag["status"]="CANDIDATE_NOT_VERIFIED"
    return None,[],"CANDIDATE_NOT_VERIFIED",diag

def classify_exception(exc):
    if isinstance(exc,requests.exceptions.Timeout): return "TIMEOUT"
    if isinstance(exc,requests.exceptions.ConnectionError): return "CONNECTION_ERROR"
    if isinstance(exc,(KeyError,IndexError,TypeError,AttributeError)): return "CODE_ERROR"
    if isinstance(exc,ValueError): return "VALUE_ERROR"
    return "PROCESS_ERROR"

def match_post_detailed(u):
    r=get(u)
    if not r: return None,"DETAIL_FETCH_ERROR","FETCH_ERROR"
    soup=BeautifulSoup(r.text,"html.parser")

    if is_list_only_url(r.url): return None,"LIST_PAGE",""
    if looks_like_list_dom(soup): return None,"LIST_PAGE",""

    path=(urlparse(r.url).path or "").lower().rstrip("/")
    if generic_url(r.url) or path in ("","/main","/home","/homepage"):
        return None,"GENERIC_PAGE",""

    dt=extract_detail_date(soup)
    if not dt or not is_recent_date(dt): return None,"OLD_OR_NO_DATE",""

    title=extract_title(soup)
    if not title: return None,"NO_TITLE",""
    if title in GENERIC_PAGE_TITLES: return None,"GENERIC_TITLE",""

    # Reject site-name-only titles.
    page_title=norm(soup.title.get_text(" ",strip=True)) if soup.title else ""
    tc=re.sub(r"[^0-9A-Za-z가-힣]","",title).lower()
    pc=re.sub(r"[^0-9A-Za-z가-힣]","",page_title).lower()
    if tc and pc and (tc==pc or tc in pc and len(tc)>=4 and len(tc)/max(len(pc),1)>0.8):
        return None,"GENERIC_TITLE",""

    if any(x in title for x in EXCLUDE_TITLE): return None,"CONTEST_TITLE",""
    if not has_post_structure(soup,title): return None,"NOT_POST_STRUCTURE",""

    # TITLE match is strongest and does not require body extraction.
    for kw in KEYWORDS:
        if kw in title:
            return {"url":r.url,"title":title[:300],"date":dt.strftime("%Y-%m-%d"),
                    "keyword":kw,"match_type":"TITLE"},"TITLE_MATCH",kw

    body=visible_main_text(soup)
    if not body:
        return None,"NO_BODY",""

    for kw in KEYWORDS:
        if meaningful_body_match(body,kw):
            return {"url":r.url,"title":title[:300],"date":dt.strftime("%Y-%m-%d"),
                    "keyword":kw,"match_type":"BODY"},"BODY_MATCH",kw
    return None,"NO_KEYWORD",""

def load_json(path,default):
    try:
        if os.path.exists(path):
            with open(path,"r",encoding="utf-8") as f: return json.load(f)
    except Exception: pass
    return default

def save_json(path,data):
    tmp=path+".tmp"
    with open(tmp,"w",encoding="utf-8") as f: json.dump(data,f,ensure_ascii=False,indent=2)
    os.replace(tmp,path)

def normalize_pending_item(x):
    if not isinstance(x,dict) or not x.get("url"): return None
    return {"url":x.get("url",""),"title":x.get("title","")[:300],
            "date":x.get("date",""),"keyword":x.get("keyword",""),
            "match_type":x.get("match_type",""),"institution":x.get("institution",""),
            "added_at":x.get("added_at",datetime.now(KST).isoformat())}

def pending_key(x): return x.get("url","") if isinstance(x,dict) else ""

def revalidate_pending(pending):
    valid=[]; removed=0
    for item in pending[:MAX_PENDING]:
        item=normalize_pending_item(item)
        if not item: removed+=1; continue
        r=get(item["url"])
        if not r: valid.append(item); continue
        m,reason,_=match_post_detailed(item["url"])
        if m:
            m["institution"]=item.get("institution","")
            m["added_at"]=item.get("added_at",datetime.now(KST).isoformat())
            valid.append(m)
        else:
            removed+=1
    return valid,removed

def find_target_file():
    for f in TARGET_FILE_CANDIDATES:
        if os.path.exists(f): return f
    return None

def load_targets():
    path=find_target_file()
    if not path: raise FileNotFoundError(f"기관 목록 파일이 없습니다: {TARGET_FILE_CANDIDATES}")
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    ws=wb["355기관"] if "355기관" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows=list(ws.iter_rows(values_only=True))
    if not rows: return []
    headers=[norm(x) for x in rows[0]]
    def col(*names):
        for name in names:
            if name in headers: return headers.index(name)
        return None
    idx_name=col("기관명","기관")
    idx_home=col("홈페이지URL","URL","홈페이지")
    idx_board=col("공지게시판URL","공지사항URL","게시판URL")
    if idx_name is None or idx_home is None: raise ValueError(f"기관명/홈페이지URL 열 없음: {headers}")
    out=[]
    for row in rows[1:]:
        if not row: continue
        name=norm(row[idx_name]) if idx_name<len(row) else ""
        home=norm(row[idx_home]) if idx_home<len(row) else ""
        board=norm(row[idx_board]) if idx_board is not None and idx_board<len(row) else ""
        if name and home: out.append({"기관명":name,"홈페이지URL":home,"공지게시판URL":board})
    return out

def process(target,state):
    name=target["기관명"]; home=target["홈페이지URL"]; seed=target.get("공지게시판URL","")
    result={"기관명":name,"home":home,"board":"","status":"","posts_checked":0,"matches":[],
            "error":"","error_type":"","diag":{},"detail_checked":0,"recent_posts":0,
            "title_matches":0,"body_matches":0,"excluded_list_pages":0,
            "excluded_generic_pages":0,"excluded_contest_titles":0,"detail_errors":0,
            "excluded_not_post_structure":0,"excluded_no_body":0}
    try:
        board=state.get("boards",{}).get(home,"") or seed
        details=[]
        if board:
            try: bu,details,info=inspect_board(board)
            except Exception: bu,details,info=None,[],{"status":"BOARD_FETCH_ERROR"}
            if info.get("status")=="VERIFIED":
                board=bu; result["board"]=bu; result["status"]="CACHED_OR_SEED"
            else: board=""; details=[]

        if not board:
            bu,details,status,diag=discover_board(home)
            result["diag"]=diag or {}
            if bu and details:
                board=bu; result["board"]=bu; result["status"]="DISCOVERED"
            else:
                result["status"]=status; return result

        unique=[]; seen=set()
        for item in details:
            u=item.get("url") if isinstance(item,dict) else item
            if u and u not in seen: seen.add(u); unique.append(item)
        result["posts_checked"]=len(unique); result["detail_checked"]=len(unique)

        matches=[]
        for item in unique:
            u=item.get("url") if isinstance(item,dict) else item
            try:
                m,reason,kw=match_post_detailed(u)
                if reason=="LIST_PAGE": result["excluded_list_pages"]+=1
                elif reason in ("GENERIC_PAGE","GENERIC_TITLE"): result["excluded_generic_pages"]+=1
                elif reason=="CONTEST_TITLE": result["excluded_contest_titles"]+=1
                elif reason=="NOT_POST_STRUCTURE": result["excluded_not_post_structure"]+=1
                elif reason=="NO_BODY": result["excluded_no_body"]+=1
                elif reason=="DETAIL_FETCH_ERROR": result["detail_errors"]+=1
                elif reason=="TITLE_MATCH": result["title_matches"]+=1
                elif reason=="BODY_MATCH": result["body_matches"]+=1
                if m:
                    m["institution"]=name; matches.append(m)
            except Exception as e:
                result["detail_errors"]+=1
        result["matches"]=matches
        result["recent_posts"]=max(0,len(unique)-result["excluded_list_pages"]-
            result["excluded_generic_pages"]-result["excluded_contest_titles"]-
            result["excluded_not_post_structure"]-result["excluded_no_body"])
        return result
    except Exception as e:
        et=classify_exception(e)
        result.update({"status":"PROCESS_ERROR","error_type":et,"error":str(e)[:500]})
        return result

def telegram_send(item):
    token=os.environ.get("TELEGRAM_BOT_TOKEN","").strip()
    chat_id=os.environ.get("TELEGRAM_CHAT_ID","").strip()
    if not token or not chat_id: return False
    text=(f"📢 [{item.get('institution','')}] 공공기관 참여/설문 게시물\n\n"
          f"제목: {item.get('title','')}\n일자: {item.get('date','')}\n"
          f"키워드: {item.get('keyword','')}\n매칭: {item.get('match_type','')}\n"
          f"링크: {item.get('url','')}")
    try:
        r=SESSION.post(f"https://api.telegram.org/bot{token}/sendMessage",
                       data={"chat_id":chat_id,"text":text[:4000],
                             "disable_web_page_preview":False},timeout=20)
        return r.ok
    except Exception:
        return False

def main():
    started=time.time()
    targets=load_targets()
    state=load_json(STATE_FILE,{"seen":[],"boards":{},"updated_at":"","version":VERSION})
    pending=load_json(PENDING_FILE,[])
    if not isinstance(pending,list): pending=[]

    pending,pending_removed_invalid=revalidate_pending(pending)
    seen=set(state.get("seen",[]))
    results=[]; completed=no_board=errors=posts_checked=new_matches=0
    status_counts={}; error_type_counts={}
    aggregate={"board_candidates":0,"post_candidates":0,"detail_checked":0,
               "recent_posts":0,"title_matches":0,"body_matches":0,
               "excluded_list_pages":0,"excluded_generic_pages":0,
               "excluded_contest_titles":0,"excluded_not_post_structure":0,
               "excluded_no_body":0,"detail_errors":0}

    deadline=started+MAX_TOTAL_SECONDS
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as ex:
        fmap={ex.submit(process,t,state):t for t in targets}
        for fut in as_completed(fmap):
            if time.time()>=deadline: break
            try:
                result=fut.result(); results.append(result)
            except Exception as e:
                errors+=1; et=classify_exception(e)
                error_type_counts[et]=error_type_counts.get(et,0)+1
                status_counts["FUTURE_EXCEPTION"]=status_counts.get("FUTURE_EXCEPTION",0)+1
                continue

            st=result.get("status") or "UNKNOWN"
            status_counts[st]=status_counts.get(st,0)+1
            if result.get("board"):
                completed+=1; state.setdefault("boards",{})[result["home"]]=result["board"]
            else: no_board+=1

            posts_checked+=result.get("posts_checked",0)
            for k in aggregate:
                aggregate[k]+=result.get(k,0)
            d=result.get("diag") or {}
            aggregate["board_candidates"]+=d.get("candidate_count",0) or 0

            et=result.get("error_type","")
            if et:
                errors+=1; error_type_counts[et]=error_type_counts.get(et,0)+1

            for m in result.get("matches",[]):
                u=m.get("url","")
                if not u or u in seen or any(pending_key(x)==u for x in pending): continue
                m["added_at"]=datetime.now(KST).isoformat()
                pending.append(m); new_matches+=1

    dedup={}
    for item in pending:
        if pending_key(item): dedup[pending_key(item)]=item
    pending=list(dedup.values())
    pending.sort(key=lambda x:x.get("added_at",""))
    if len(pending)>MAX_PENDING: pending=pending[-MAX_PENDING:]

    pending_before_send=len(pending); sent=0; remaining=[]
    for item in pending:
        if sent>=TELEGRAM_MAX_SEND:
            remaining.append(item); continue
        if telegram_send(item):
            sent+=1; seen.add(item.get("url",""))
        else: remaining.append(item)
    pending=remaining

    state["seen"]=list(seen)[-50000:]
    state["updated_at"]=datetime.now(KST).isoformat()
    state["version"]=VERSION
    save_json(STATE_FILE,state); save_json(PENDING_FILE,pending)

    elapsed=time.time()-started
    diag_counts={}
    for r in results:
        st=(r.get("diag") or {}).get("status")
        if st: diag_counts[st]=diag_counts.get(st,0)+1

    diagnostics={
        "version":VERSION,"updated_at":datetime.now(KST).isoformat(),
        "targets":len(targets),"completed":completed,"no_board":no_board,
        "posts_checked":posts_checked,"new_matches":new_matches,
        "pending_before":pending_before_send,"pending_removed_invalid":pending_removed_invalid,
        "telegram_sent":sent,"pending_after":len(pending),"errors":errors,
        "elapsed_seconds":round(elapsed,1),"timed_out":elapsed>=MAX_TOTAL_SECONDS,
        "recent_days":RECENT_DAYS,"telegram_max_send":TELEGRAM_MAX_SEND,
        "board_discovery_status":diag_counts,"status_counts":status_counts,
        "error_type_counts":error_type_counts,"pipeline_counts":aggregate,
        "processed_total":len(results),"unprocessed_total":max(0,len(targets)-len(results)),
        "board_details":[{
            "기관명":r.get("기관명"),"status":r.get("status"),"board":r.get("board"),
            "posts_checked":r.get("posts_checked"),"detail_checked":r.get("detail_checked"),
            "recent_posts":r.get("recent_posts"),"title_matches":r.get("title_matches"),
            "body_matches":r.get("body_matches"),
            "excluded_list_pages":r.get("excluded_list_pages"),
            "excluded_generic_pages":r.get("excluded_generic_pages"),
            "excluded_contest_titles":r.get("excluded_contest_titles"),
            "excluded_not_post_structure":r.get("excluded_not_post_structure"),
            "excluded_no_body":r.get("excluded_no_body"),
            "detail_errors":r.get("detail_errors"),
            "diag_status":(r.get("diag") or {}).get("status"),
            "candidate_count":(r.get("diag") or {}).get("candidate_count"),
            "fetched_count":(r.get("diag") or {}).get("fetched_count"),
            "verified_count":(r.get("diag") or {}).get("verified_count"),
            "selected_score":(r.get("diag") or {}).get("selected_score"),
            "error_type":r.get("error_type",""),"error":r.get("error","")[:500]
        } for r in results]
    }
    save_json("diagnostics.json",diagnostics)
    print(json.dumps(diagnostics,ensure_ascii=False,indent=2))

if __name__=="__main__":
    main()
