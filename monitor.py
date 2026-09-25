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

VERSION = "V8.14.6"

import os, re, json, time, html, warnings, hashlib
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

MAX_CONCURRENCY = 25
TIMEOUT_SECONDS = 8
HTTP_RETRIES = 1
RECENT_POSTS = 10
RECENT_DAYS = 30
MAX_TOTAL_SECONDS = 1080
TELEGRAM_MAX_SEND = 20
MAX_PENDING = 10000

BOARD_DISCOVERY_MAX_LINKS = 100
BOARD_DISCOVERY_MAX_FETCH = 12
BOARD_DISCOVERY_MAX_SECONDARY = 6
BOARD_CACHE_FILE = "board_cache.json"
MAX_FINGERPRINTS = 100000
DAILY_SUMMARY_FILE = "daily_summary.json"
RETRY_QUEUE_FILE = "retry_queue.json"
MAX_RETRY_QUEUE = 120
CHECKED_POSTS_FILE = "checked_posts.json"
MAX_CHECKED_POSTS = 50000

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
    headers_variants=[{}, {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128 Safari/537.36"}]
    for attempt in range(HTTP_RETRIES + 1):
        try:
            extra=headers_variants[min(attempt, len(headers_variants)-1)]
            r=SESSION.get(url, timeout=TIMEOUT_SECONDS, allow_redirects=True, verify=True, headers=extra)
            if 200 <= r.status_code < 400:
                if not r.encoding or r.encoding.lower() == "iso-8859-1":
                    r.encoding = r.apparent_encoding or "utf-8"
                return r
        except Exception:
            pass
        if attempt < HTTP_RETRIES:
            time.sleep(0.5 * (attempt + 1))
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
    """실제 게시물 제목을 우선 추출한다. 사이트 공통 title은 최후순위로 사용."""
    if not soup: return ""
    def valid(t):
        t=norm(t)
        if not (2 <= len(t) <= 300): return False
        if t in GENERIC_PAGE_TITLES: return False
        if len(re.sub(r"[^0-9A-Za-z가-힣]","",t)) < 3: return False
        return True

    # 1) 게시물 전용 제목 영역
    selectors=[
        "h1.view-title","h1.board-title","h1.article-title","h1.bbs-title",
        ".view-title h1",".board-title h1",".article-title h1",
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
        vals.sort(key=lambda x:(len(x)<5, len(x)>180, -len(x)))
        return vals[0][:300]

    # 2) OpenGraph / social title
    for sel in ["meta[property='og:title']","meta[name='twitter:title']"]:
        og=soup.select_one(sel)
        if og and og.get("content"):
            parts=[norm(x) for x in re.split(r"\s*[|｜]\s*",og.get("content")) if norm(x)]
            vals=[p for p in parts if valid(p)]
            if vals:
                return sorted(vals,key=len,reverse=True)[0][:300]

    # 3) 흔한 게시물 제목 테이블/정의목록 구조
    for tag in soup.find_all(["th","dt","strong"]):
        label=norm(tag.get_text(" ",strip=True))
        if label in ("제목","게시물 제목","글제목","내용"):
            sib=tag.find_next_sibling()
            if sib:
                t=text_of(sib)
                if valid(t): return t[:300]

    # 4) document title은 공통 사이트 제목을 최대한 배제
    if soup.title:
        parts=[norm(x) for x in re.split(r"\s*[|｜]\s*",soup.title.get_text(" ",strip=True)) if norm(x)]
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

def secondary_discovery(home):
    """1차 링크 탐색 실패 시 sitemap/robots/대표 게시판 경로를 이용한 2차 탐색."""
    candidates=[]; seen_urls=set()
    try:
        base=urlparse(home); root=f"{base.scheme or 'https'}://{base.netloc}"
    except Exception:
        return candidates
    for suffix in ["/sitemap.xml","/robots.txt"]:
        u=root+suffix; r=get(u)
        if not r: continue
        text=r.text or ""
        for x in re.findall(r'https?://[^\s<>"]+',text)[:200]:
            x=clean_url(x)
            if same_domain(root,x) and x not in seen_urls:
                seen_urls.add(x); candidates.append((board_score(x,""),x,"secondary-sitemap"))
        if suffix.endswith("robots.txt"):
            for line in text.splitlines():
                if line.lower().startswith("sitemap:"):
                    su=line.split(":",1)[1].strip(); rr=get(su) if su else None
                    if rr:
                        for x in re.findall(r'https?://[^\s<>"]+',rr.text or "")[:200]:
                            x=clean_url(x)
                            if same_domain(root,x) and x not in seen_urls:
                                seen_urls.add(x); candidates.append((board_score(x,""),x,"secondary-robots-sitemap"))
    paths=["/board","/bbs","/notice","/news","/community","/participation","/board/list","/bbs/list","/notice/list","/boardList.do","/bbsList.do","/board/list.do","/bbs/list.do","/contents/board","/site/board"]
    for path in paths:
        u=root+path
        if u not in seen_urls:
            seen_urls.add(u); candidates.append((board_score(u,path),u,"secondary-common-path"))
    candidates.sort(reverse=True)
    return candidates[:BOARD_DISCOVERY_MAX_SECONDARY]

def discover_board(home):
    diag={"home":home,"status":"START","candidate_count":0,"fetched_count":0,"verified_count":0,"selected":"","selected_score":None,"candidates":[],"secondary_candidate_count":0}
    r=get(home)
    if not r:
        diag["status"]="HOME_ERROR"; return None,[],"HOME_ERROR",diag
    soup=BeautifulSoup(r.text,"html.parser")
    links=extract_links(r.url,soup)
    scored=sorted([(board_score(u,t),u,t) for u,t in links],reverse=True)
    diag["candidate_count"]=len(scored)
    best=None
    for score,u,t in scored[:BOARD_DISCOVERY_MAX_FETCH]:
        if score < -8: continue
        diag["fetched_count"]+=1
        try: bu,details,info=inspect_board(u)
        except Exception as e:
            diag["candidates"].append({"score":score,"url":u,"title":t[:120],"status":"INSPECT_ERROR","error_type":type(e).__name__}); continue
        diag["candidates"].append({"score":score,"url":u,"title":t[:120],"status":info.get("status"),"candidate_posts":info.get("candidate_posts",0),"source":"primary"})
        if info.get("status")!="VERIFIED": continue
        diag["verified_count"]+=1
        combined=score+min(info.get("candidate_posts",0),10)
        if any(k in info.get("recent_titles","") for k in KEYWORDS): combined+=2
        if best is None or combined>best["score"]: best={"score":combined,"url":bu,"details":details}
    if best is None:
        secondary=secondary_discovery(home); diag["secondary_candidate_count"]=len(secondary)
        for score,u,t in secondary:
            diag["fetched_count"]+=1
            try: bu,details,info=inspect_board(u)
            except Exception as e:
                diag["candidates"].append({"score":score,"url":u,"title":t[:120],"status":"INSPECT_ERROR","source":"secondary","error_type":type(e).__name__}); continue
            diag["candidates"].append({"score":score,"url":u,"title":t[:120],"status":info.get("status"),"candidate_posts":info.get("candidate_posts",0),"source":"secondary"})
            if info.get("status")!="VERIFIED": continue
            diag["verified_count"]+=1
            combined=score+min(info.get("candidate_posts",0),10)
            if any(k in info.get("recent_titles","") for k in KEYWORDS): combined+=2
            if best is None or combined>best["score"]: best={"score":combined,"url":bu,"details":details}
    if best:
        diag["status"]="VERIFIED"; diag["selected"]=best["url"]; diag["selected_score"]=best["score"]
        return best["url"],best["details"][:RECENT_POSTS],"DISCOVERED",diag
    if not scored and not diag["secondary_candidate_count"]:
        diag["status"]="NO_CANDIDATE"; return None,[],"NO_CANDIDATE",diag
    diag["status"]="CANDIDATE_NOT_VERIFIED"; return None,[],"CANDIDATE_NOT_VERIFIED",diag


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

    # 페이지 전체 title과 동일할 때만 공통 페이지 제목으로 간주.
    page_title=norm(soup.title.get_text(" ",strip=True)) if soup.title else ""
    tc=re.sub(r"[^0-9A-Za-z가-힣]","",title).lower()
    pc=re.sub(r"[^0-9A-Za-z가-힣]","",page_title).lower()
    if tc and pc and tc==pc and len(tc)>=6:
        return None,"GENERIC_TITLE",""

    if any(x in title for x in EXCLUDE_TITLE): return None,"CONTEST_TITLE",""

    body=visible_main_text(soup)
    # 게시물 검증은 기존보다 완화한다. 상세 URL + 날짜 + 제목 + 내용이 있으면 통과시키고,
    # 실제 매칭 여부는 아래 제목/본문 단계에서 판단한다.
    if not body or len(body)<60:
        if not has_post_structure(soup,title): return None,"NOT_POST_STRUCTURE",""

    for kw in KEYWORDS:
        if kw in title:
            item={"url":r.url,"title":title[:300],"date":dt.strftime("%Y-%m-%d"),
                  "keyword":kw,"match_type":"TITLE","body":body}
            item["fingerprint"]=make_fingerprint(item,body)
            item.pop("body",None)
            return item,"TITLE_MATCH",kw

    if not body:
        return None,"NO_BODY",""

    for kw in KEYWORDS:
        if meaningful_body_match(body,kw):
            item={"url":r.url,"title":title[:300],"date":dt.strftime("%Y-%m-%d"),
                  "keyword":kw,"match_type":"BODY","body":body}
            item["fingerprint"]=make_fingerprint(item,body)
            item.pop("body",None)
            return item,"BODY_MATCH",kw
    return None,"NO_KEYWORD",""

def canonical_url(url):
    """동일 게시물 URL의 표기 차이를 줄인다."""
    try:
        p=urlparse(clean_url(url))
        scheme=(p.scheme or "https").lower()
        host=(p.netloc or "").lower()
        if host.startswith("www."): host=host[4:]
        path=(p.path or "/").rstrip("/") or "/"
        query=p.query or ""
        return f"{scheme}://{host}{path}" + (f"?{query}" if query else "")
    except Exception:
        return clean_url(url)

def normalize_fingerprint_text(text):
    text=norm(text)
    # 날짜/시간, 조회수 등 실행마다 변할 수 있는 숫자성 표시를 약간 완화한다.
    text=re.sub(r"20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}(?:\s+\d{1,2}:\d{2})?", "DATE", text)
    text=re.sub(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", "TIME", text)
    text=re.sub(r"(조회수|조회)\s*[:：]?\s*\d+", r"\1", text)
    return text[:8000]

def make_fingerprint(item, body=""):
    """URL이 바뀌어도 같은 게시물로 판별하기 위한 내용 기반 fingerprint."""
    # 기관명은 제외한다. 기존 V8.14.x 발송 이력에는 기관명이 저장되지 않은 경우가 있어
    # 과거 URL에서 재생성한 fingerprint와 신규 탐지 fingerprint가 동일해야 한다.
    date=norm(item.get("date",""))
    title=normalize_fingerprint_text(item.get("title",""))
    body_norm=normalize_fingerprint_text(body)
    # 본문이 너무 긴 경우 앞/뒤를 함께 사용해 동적 footer 영향을 줄인다.
    if len(body_norm)>5000:
        body_norm=body_norm[:3500]+body_norm[-1200:]
    raw="|".join([date,title,body_norm])
    return "fp:"+hashlib.sha256(raw.encode("utf-8","ignore")).hexdigest()

def alert_key(item):
    """하위호환용 URL 키. 실제 중복판정은 URL + fingerprint를 함께 사용."""
    return canonical_url(item.get("url", ""))

def migrate_sent_ledger(state):
    if not isinstance(state,dict): state={}
    state.setdefault("seen",[])
    old=state.get("sent_urls",{})
    if not isinstance(old,dict): old={}
    migrated=dict(old)
    for u in state.get("seen",[]):
        k=canonical_url(u)
        if k and k not in migrated:
            migrated[k]=state.get("updated_at",datetime.now(KST).isoformat())
    state["sent_urls"]=migrated
    state.setdefault("sent_fingerprints",{})
    if not isinstance(state["sent_fingerprints"],dict): state["sent_fingerprints"]={}
    state.setdefault("seen_fingerprints",[])
    if not isinstance(state["seen_fingerprints"],list): state["seen_fingerprints"]=[]
    return state

def clean_pending_against_ledger(pending, sent_keys, seen, sent_fingerprints, seen_fingerprints):
    out=[]; seen_pending=set(); seen_fp_pending=set()
    for item in pending if isinstance(pending,list) else []:
        if not isinstance(item,dict): continue
        u=canonical_url(item.get("url","")); fp=item.get("fingerprint","")
        if not u: continue
        if u in sent_keys or u in seen or (fp and (fp in sent_fingerprints or fp in seen_fingerprints)): continue
        if u in seen_pending or (fp and fp in seen_fp_pending): continue
        item["url"]=u
        seen_pending.add(u)
        if fp: seen_fp_pending.add(fp)
        out.append(item)
    return out

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
            "fingerprint":x.get("fingerprint",""),
            "added_at":x.get("added_at",datetime.now(KST).isoformat())}

def pending_key(x): return canonical_url(x.get("url","")) if isinstance(x,dict) else ""

def revalidate_pending(pending, seen, sent_keys, seen_fingerprints, sent_fingerprints):
    valid=[]; removed=0
    for item in pending[:MAX_PENDING]:
        item=normalize_pending_item(item)
        if not item:
            removed+=1; continue
        key=canonical_url(item["url"])
        fp=item.get("fingerprint","")
        if key in seen or key in sent_keys or (fp and (fp in seen_fingerprints or fp in sent_fingerprints)):
            removed+=1; continue
        m,reason,_=match_post_detailed(key)
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
    idx_type=col("기관유형","기관 유형","유형","기관구분","구분")
    if idx_name is None or idx_home is None: raise ValueError(f"기관명/홈페이지URL 열 없음: {headers}")
    out=[]
    for row in rows[1:]:
        if not row: continue
        name=norm(row[idx_name]) if idx_name<len(row) else ""
        home=norm(row[idx_home]) if idx_home<len(row) else ""
        board=norm(row[idx_board]) if idx_board is not None and idx_board<len(row) else ""
        org_type=norm(row[idx_type]) if idx_type is not None and idx_type<len(row) else ""
        if name and home:
            out.append({
                "기관명":name,"홈페이지URL":home,"공지게시판URL":board,
                "기관유형":org_type
            })
    return out

def load_board_cache():
    data=load_json(BOARD_CACHE_FILE,{})
    return data if isinstance(data,dict) else {}

def cache_board(board_cache,name,home,board,status="VERIFIED"):
    if board:
        board_cache[home]={"기관명":name,"홈페이지URL":home,"게시판URL":board,"status":status,"updated_at":datetime.now(KST).isoformat(),"version":VERSION}

def board_from_cache(board_cache,home):
    x=board_cache.get(home,{})
    return x.get("게시판URL","") if isinstance(x,dict) else ""

def load_retry_queue():
    data=load_json(RETRY_QUEUE_FILE,{})
    return data if isinstance(data,dict) else {}

def save_retry_queue(data):
    save_json(RETRY_QUEUE_FILE,data)

def update_retry_queue(queue,result):
    name=result.get("기관명","")
    if not name: return
    st=result.get("status","")
    retryable=st in ("HOME_ERROR","CANDIDATE_NOT_VERIFIED","NO_CANDIDATE")
    now=datetime.now(KST).isoformat()
    item=queue.get(name,{}) if isinstance(queue.get(name,{}),dict) else {}
    if retryable:
        item.update({"기관명":name,"홈페이지URL":result.get("home",""),"status":st,
                     "attempts":int(item.get("attempts",0))+1,"last_attempt":now})
        queue[name]=item
    else:
        queue.pop(name,None)

def post_identity(institution, item):
    """URL이 매번 바뀌는 사이트에서도 이미 확인한 동일 게시물을 식별한다."""
    date=norm(str(item.get("date") or ""))
    title=normalize_fingerprint_text(str(item.get("title") or ""))
    if not date and not title:
        return ""
    raw="|".join([norm(institution), date, title])
    return "pid:"+hashlib.sha256(raw.encode("utf-8","ignore")).hexdigest()

def load_checked_posts():
    try:
        with open(CHECKED_POSTS_FILE,"r",encoding="utf-8") as f:
            x=json.load(f)
        return x if isinstance(x,dict) else {}
    except Exception:
        return {}

def save_checked_posts(data):
    try:
        if len(data)>MAX_CHECKED_POSTS:
            data=dict(list(data.items())[-MAX_CHECKED_POSTS:])
        atomic_write_json(CHECKED_POSTS_FILE,data)
    except Exception:
        pass

def process(target,state,board_cache,checked_posts):
    name=target["기관명"]; home=target["홈페이지URL"]; seed=target.get("공지게시판URL","")
    result={"기관명":name,"기관유형":target.get("기관유형",""),"home":home,"board":"","status":"","posts_checked":0,"matches":[],
            "error":"","error_type":"","diag":{},"post_candidates":0,"detail_checked":0,"recent_posts":0,
            "title_matches":0,"body_matches":0,"excluded_list_pages":0,
            "excluded_generic_pages":0,"excluded_contest_titles":0,"detail_errors":0,
            "excluded_not_post_structure":0,"excluded_no_body":0}
    try:
        board=board_from_cache(board_cache,home) or state.get("boards",{}).get(home,"") or seed
        details=[]
        if board:
            try: bu,details,info=inspect_board(board)
            except Exception: bu,details,info=None,[],{"status":"BOARD_FETCH_ERROR"}
            if info.get("status")=="VERIFIED":
                board=bu; result["board"]=bu; result["status"]="CACHED_OR_SEED"; cache_board(board_cache,name,home,bu,"VERIFIED")
            else: board=""; details=[]

        if not board:
            bu,details,status,diag=discover_board(home)
            result["diag"]=diag or {}
            if bu and details:
                board=bu; result["board"]=bu; result["status"]="DISCOVERED"; cache_board(board_cache,name,home,bu,"VERIFIED")
            else:
                result["status"]=status; return result

        unique=[]; seen=set(); skipped_checked=0
        for item in details:
            if not isinstance(item,dict):
                item={"url":item}
            u=item.get("url")
            if not u or u in seen:
                continue
            seen.add(u)
            pid=post_identity(name,item)
            if pid and pid in checked_posts:
                skipped_checked+=1
                continue
            item["post_identity"]=pid
            unique.append(item)
        result["post_candidates"]=len(details)
        result["skipped_previously_checked"]=skipped_checked
        result["posts_checked"]=len(unique); result["detail_checked"]=len(unique)

        matches=[]
        checked_ids=[]
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
            finally:
                pid=item.get("post_identity","")
                if pid: checked_ids.append(pid)
        result["checked_post_identities"]=checked_ids
        result["matches"]=matches
        result["recent_posts"]=max(0,len(unique)-result["excluded_list_pages"]-
            result["excluded_generic_pages"]-result["excluded_contest_titles"]-
            result["excluded_not_post_structure"]-result["excluded_no_body"])
        return result
    except Exception as e:
        et=classify_exception(e)
        result.update({"status":"PROCESS_ERROR","error_type":et,"error":str(e)[:500]})
        return result


def classify_institution_type(value):
    """기관유형 원본 표현을 3개 통계군으로 표준화한다."""
    t=norm(value).replace(" ","")
    if "공기업" in t:
        return "공기업"
    if "준정부기관" in t:
        return "준정부기관"
    if "기타공공기관" in t:
        return "기타공공기관"
    return "기타/미분류"

def summarize_group(items):
    total=len(items)
    ok=sum(1 for x in items if x.get("status") in ("CACHED_OR_SEED","DISCOVERED"))
    no_board=total-ok
    rate=(ok/total*100) if total else 0
    return {"대상":total,"정상확인":ok,"미확인":no_board,"확인율":round(rate,1)}

def telegram_send_text(text):
    token=os.environ.get("TELEGRAM_BOT_TOKEN","").strip()
    chat_id=os.environ.get("TELEGRAM_CHAT_ID","").strip()
    if not token or not chat_id:
        return False
    try:
        r=SESSION.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            data={"chat_id":chat_id,"text":text[:4000],
                  "disable_web_page_preview":True},
            timeout=20
        )
        return r.ok
    except Exception:
        return False

def build_monitoring_summary(targets, results, posts_checked, new_matches,
                             sent, pending_after, errors, aggregate,
                             status_counts, sent_url_ledger):
    groups={
        "공기업": [r for r in results if classify_institution_type(r.get("기관유형",""))=="공기업"],
        "준정부기관": [r for r in results if classify_institution_type(r.get("기관유형",""))=="준정부기관"],
        "공기업·준정부기관": [
            r for r in results
            if classify_institution_type(r.get("기관유형","")) in ("공기업","준정부기관")
        ],
        "기타공공기관": [
            r for r in results if classify_institution_type(r.get("기관유형",""))=="기타공공기관"
        ],
    }

    total=len(targets)
    processed=len(results)
    ok=sum(1 for r in results if r.get("status") in ("CACHED_OR_SEED","DISCOVERED"))
    unconfirmed=[r for r in results if r.get("status") not in ("CACHED_OR_SEED","DISCOVERED")]
    unprocessed=max(0,total-processed)

    def fmt(g):
        x=summarize_group(g)
        return f"• 대상 {x['대상']} / 정상 {x['정상확인']} / 미확인 {x['미확인']} / 확인율 {x['확인율']}%"

    no_candidate=sum(1 for r in results if r.get("status")=="NO_CANDIDATE")
    home_error=sum(1 for r in results if r.get("status")=="HOME_ERROR")
    candidate_not_verified=sum(1 for r in results if r.get("status")=="CANDIDATE_NOT_VERIFIED")

    fail_names=[r.get("기관명","") for r in unconfirmed if r.get("기관명")]
    fail_preview=fail_names[:10]

    lines=[
        "📊 티끌 모니터링 요약",
        "━━━━━━━━━━━━━━",
        f"📅 {datetime.now(KST).strftime('%Y-%m-%d')}",
        "",
        "🏢 전체 기관",
        f"• 대상 {total} / 정상 확인 {ok} / 미확인 {len(unconfirmed)} / 확인율 {(ok/processed*100 if processed else 0):.1f}%",
        f"• 미처리 {unprocessed}",
        f"• 전체 처리 완료 {processed}/{total}" + (" ⚠️" if unprocessed else ""),
        "",
        "🏛 공기업·준정부기관",
        fmt(groups["공기업·준정부기관"]),
        "",
        "🏢 기타공공기관",
        fmt(groups["기타공공기관"]),
        "",
        "🔎 오늘 모니터링",
        f"• 게시물 확인 {posts_checked:,}건",
        f"• 실제 최근 게시물 {aggregate.get('recent_posts',0):,}건",
        f"• 신규 키워드 매칭 {new_matches}건",
        f"• Telegram 참여정보 발송 {sent}건",
        f"• 대기 {pending_after}건",
        "",
        "⚠️ 게시판 미확인",
        f"• 후보 없음(NO_CANDIDATE) {no_candidate}개",
        f"• 홈페이지 오류(HOME_ERROR) {home_error}개",
        f"• 후보 검증 실패 {candidate_not_verified}개",
        f"• 기타 오류/미처리 {max(0,len(unconfirmed)-no_candidate-home_error-candidate_not_verified)}개",
        "",
        f"💾 누적 발송 URL {sent_url_ledger:,}개",
        f"• 실행 오류 {errors}건",
    ]

    if fail_preview:
        lines += ["", "📌 미확인 기관 예시(최대 10개)"]
        lines += [f"• {x}" for x in fail_preview]

    return "\n".join(lines)

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

def migrate_existing_url_fingerprints(sent_urls, sent_fingerprints):
    """V8.14.4 최초 1회: 기존 URL 발송 이력에서 내용 fingerprint를 생성해 중복 재발송을 방지."""
    added=0
    for u in list(sent_urls.keys())[-500:]:
        if len(sent_fingerprints)>=MAX_FINGERPRINTS: break
        try:
            m,reason,_=match_post_detailed(u)
            if m and m.get("fingerprint") and m["fingerprint"] not in sent_fingerprints:
                sent_fingerprints.add(m["fingerprint"]); added+=1
        except Exception:
            continue
    return added

def load_daily_summary_state():
    x=load_json(DAILY_SUMMARY_FILE,{})
    return x if isinstance(x,dict) else {}

def save_daily_summary_state(x):
    save_json(DAILY_SUMMARY_FILE,x)

def main():
    started=time.time()
    targets=load_targets()
    state=load_json(
        STATE_FILE,
        {"seen":[],"sent_urls":{},"boards":{},"updated_at":"","version":VERSION}
    )
    if not isinstance(state,dict):
        state={"seen":[],"sent_urls":{},"boards":{},"updated_at":"","version":VERSION}

    # V8.13.1의 seen 기록을 V8.14.1의 영구 발송 이력으로 승계.
    state=migrate_sent_ledger(state)
    board_cache=load_board_cache()
    checked_posts=load_checked_posts()

    seen=set(canonical_url(x) for x in state.get("seen",[]) if x)
    sent_urls=state.get("sent_urls",{})
    sent_keys=set(canonical_url(x) for x in sent_urls.keys() if x)
    sent_fingerprints=set(state.get("sent_fingerprints",{}).keys())
    seen_fingerprints=set(state.get("seen_fingerprints",[]))
    fingerprint_migrated=bool(state.get("v8144_fingerprint_migration_done"))
    migrated_fingerprint_count=0
    if not fingerprint_migrated:
        migrated_fingerprint_count=migrate_existing_url_fingerprints(sent_urls, sent_fingerprints)
        state["v8144_fingerprint_migration_done"]=True

    pending=load_json(PENDING_FILE,[])
    pending=clean_pending_against_ledger(pending, sent_keys, seen, sent_fingerprints, seen_fingerprints)

    pending,pending_removed_invalid=revalidate_pending(pending, seen, sent_keys, seen_fingerprints, sent_fingerprints)

    # V8.14.1 최초 전환 여부를 기록한다.
    # 기존 state가 존재하면 기존 seen/sent 이력을 그대로 승계하고,
    # 신규 게시물만 기존 로직에 따라 탐지한다.
    migrated_from_previous = bool(state.get("updated_at")) and not state.get("v8141_migration_done")
    state["v8141_migration_done"] = True

    results=[]; completed=no_board=errors=posts_checked=new_matches=0
    status_counts={}; error_type_counts={}
    aggregate={"board_candidates":0,"post_candidates":0,"detail_checked":0,
               "recent_posts":0,"title_matches":0,"body_matches":0,
               "excluded_list_pages":0,"excluded_generic_pages":0,
               "excluded_contest_titles":0,"excluded_not_post_structure":0,
               "excluded_no_body":0,"detail_errors":0,"skipped_previously_checked":0}

    deadline=started+MAX_TOTAL_SECONDS
    executor=ThreadPoolExecutor(max_workers=MAX_CONCURRENCY)
    fmap={executor.submit(process,t,state,board_cache,checked_posts):t for t in targets}
    timed_out=False
    try:
        pending_futures=set(fmap)
        while pending_futures:
            remaining=max(0,deadline-time.time())
            if remaining <= 0:
                timed_out=True
                break
            done_now=set()
            for fut in list(pending_futures):
                if fut.done(): done_now.add(fut)
            if not done_now:
                time.sleep(0.15)
                continue
            for fut in done_now:
                pending_futures.discard(fut)
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
                    completed+=1; state.setdefault("boards",{})[result["home"]]=result["board"]; cache_board(board_cache,result.get("기관명",""),result["home"],result["board"],"VERIFIED")
                else: no_board+=1

                posts_checked+=result.get("posts_checked",0)
                for pid in result.get("checked_post_identities",[]):
                    checked_posts[pid]=datetime.now(KST).isoformat()
                aggregate.setdefault("skipped_previously_checked",0)
                aggregate["skipped_previously_checked"]+=result.get("skipped_previously_checked",0)
                for k in aggregate:
                    if k=="skipped_previously_checked": continue
                    v=result.get(k,0)
                    # aggregate에는 숫자형 지표만 누적한다.
                    # checked_post_identities 같은 리스트형 결과가 섞여도
                    # int + list TypeError가 발생하지 않도록 방어한다.
                    if isinstance(v,(int,float)):
                        aggregate[k]+=v
                d=result.get("diag") or {}
                aggregate["board_candidates"]+=d.get("candidate_count",0) or 0
                et=result.get("error_type","")
                if et:
                    errors+=1; error_type_counts[et]=error_type_counts.get(et,0)+1
                for m in result.get("matches",[]):
                    u=m.get("url",""); key=alert_key(m); fp=m.get("fingerprint","")
                    if (not u or not key or key in seen or key in sent_keys
                        or (fp and (fp in seen_fingerprints or fp in sent_fingerprints))
                        or any(alert_key(x)==key or (fp and x.get("fingerprint")==fp) for x in pending)):
                        continue
                    m["url"]=canonical_url(u); m["added_at"]=datetime.now(KST).isoformat(); pending.append(m); new_matches+=1
    finally:
        if 'pending_futures' in locals() and pending_futures:
            timed_out=True
            for fut in pending_futures: fut.cancel()
        executor.shutdown(wait=False, cancel_futures=True)

    retry_queue=load_retry_queue()
    for r in results:
        update_retry_queue(retry_queue,r)
    retry_queue={k:v for k,v in list(retry_queue.items())[-MAX_RETRY_QUEUE:]}
    save_retry_queue(retry_queue)
    save_checked_posts(checked_posts)

    dedup={}
    for item in pending:
        k=item.get("fingerprint") or alert_key(item)
        if k: dedup[k]=item
    pending=list(dedup.values())
    pending.sort(key=lambda x:x.get("added_at",""))
    if len(pending)>MAX_PENDING: pending=pending[-MAX_PENDING:]

    pending_before_send=len(pending); sent=0; remaining=[]
    for item in pending:
        if sent>=TELEGRAM_MAX_SEND:
            remaining.append(item)
            continue

        key=alert_key(item)
        fp=item.get("fingerprint","")
        if not key:
            continue

        if key in sent_keys or key in seen or (fp and (fp in sent_fingerprints or fp in seen_fingerprints)):
            seen.add(key)
            if fp: seen_fingerprints.add(fp)
            continue

        if telegram_send(item):
            sent+=1
            seen.add(key)
            sent_keys.add(key)
            sent_urls[key]=datetime.now(KST).isoformat()
            if fp:
                sent_fingerprints.add(fp)
        else:
            remaining.append(item)
    pending=remaining

    state["seen"]=list(seen)[-100000:]
    state["seen_fingerprints"]=list(seen_fingerprints)[-MAX_FINGERPRINTS:]
    state["sent_urls"]=dict(list(sent_urls.items())[-100000:])
    state["sent_fingerprints"]={fp:datetime.now(KST).isoformat() for fp in list(sent_fingerprints)[-MAX_FINGERPRINTS:]}
    state["updated_at"]=datetime.now(KST).isoformat()
    state["version"]=VERSION
    save_json(STATE_FILE,state); save_json(PENDING_FILE,pending)

    elapsed=time.time()-started
    diag_counts={}
    for r in results:
        st=(r.get("diag") or {}).get("status")
        if st: diag_counts[st]=diag_counts.get(st,0)+1

    # 기관유형별 진단 통계
    type_summary={}
    for label in ["공기업","준정부기관","공기업·준정부기관","기타공공기관","기타/미분류"]:
        if label=="공기업":
            group=[r for r in results if classify_institution_type(r.get("기관유형",""))=="공기업"]
        elif label=="준정부기관":
            group=[r for r in results if classify_institution_type(r.get("기관유형",""))=="준정부기관"]
        elif label=="공기업·준정부기관":
            group=[r for r in results if classify_institution_type(r.get("기관유형","")) in ("공기업","준정부기관")]
        elif label=="기타공공기관":
            group=[r for r in results if classify_institution_type(r.get("기관유형",""))=="기타공공기관"]
        else:
            group=[r for r in results if classify_institution_type(r.get("기관유형",""))=="기타/미분류"]
        type_summary[label]=summarize_group(group)

    summary_text=build_monitoring_summary(
        targets,results,posts_checked,new_matches,sent,len(pending),errors,
        aggregate,status_counts,len(sent_urls)
    )
    today_key=datetime.now(KST).strftime("%Y-%m-%d")
    daily_summary=load_daily_summary_state()
    complete_run=(len(results)==len(targets) and not timed_out)
    if daily_summary.get("last_sent_date")==today_key and complete_run:
        summary_sent=False
        summary_skipped_duplicate=True
    else:
        summary_sent=telegram_send_text(summary_text)
        summary_skipped_duplicate=not summary_sent
        if summary_sent and complete_run:
            daily_summary={"last_sent_date":today_key,"updated_at":datetime.now(KST).isoformat(),"version":VERSION}
            save_daily_summary_state(daily_summary)

    diagnostics={
        "version":VERSION,"updated_at":datetime.now(KST).isoformat(),
        "targets":len(targets),"completed":completed,"no_board":no_board,
        "posts_checked":posts_checked,"new_matches":new_matches,
        "pending_before":pending_before_send,"pending_removed_invalid":pending_removed_invalid,
        "telegram_sent":sent,"summary_telegram_sent":summary_sent,
        "pending_after":len(pending),"sent_url_ledger":len(sent_urls),
        "sent_fingerprint_ledger":len(sent_fingerprints),
        "checked_post_identity_ledger":len(checked_posts),
        "migrated_fingerprint_count":migrated_fingerprint_count,
        "summary_skipped_duplicate":summary_skipped_duplicate,
        "board_cache_total":len(board_cache),
        "board_cache_verified":sum(1 for x in board_cache.values() if isinstance(x,dict) and x.get("게시판URL")),"migrated_from_previous":migrated_from_previous,"errors":errors,
        "elapsed_seconds":round(elapsed,1),"timed_out":bool(timed_out or len(results)<len(targets)),
        "recent_days":RECENT_DAYS,"telegram_max_send":TELEGRAM_MAX_SEND,
        "board_discovery_status":diag_counts,"status_counts":status_counts,
        "error_type_counts":error_type_counts,"pipeline_counts":aggregate,
        "processed_total":len(results),"unprocessed_total":max(0,len(targets)-len(results)),
        "retry_queue_total":len(retry_queue),
        "institution_type_summary":type_summary,
        "unconfirmed_institutions":[
            {
                "기관명":r.get("기관명"),"기관유형":r.get("기관유형",""),
                "status":r.get("status"),"error_type":r.get("error_type",""),
                "error":r.get("error","")[:500]
            }
            for r in results
            if r.get("status") not in ("CACHED_OR_SEED","DISCOVERED")
        ],
        "board_details":[{
            "기관명":r.get("기관명"),"기관유형":r.get("기관유형",""),
            "status":r.get("status"),"board":r.get("board"),
            "post_candidates":r.get("post_candidates"),"posts_checked":r.get("posts_checked"),"detail_checked":r.get("detail_checked"),
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
    if len(results) < len(targets) or timed_out:
        raise SystemExit(1)

if __name__=="__main__":
    main()
