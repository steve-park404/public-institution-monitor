#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V8.15.3 게시판 자동 발견기
핵심 변경:
- CSV 헤더 이름에 의존하지 않음
- CSV의 실제 값으로 기관명/URL 컬럼 추론
- CSV에서 URL을 못 찾으면 monitor.py 등 Python 소스에서 URL/기관 데이터를 추출
- 최종 기관 수 300 미만이면 실패
- eGov selectBoardList.do + bbsId 우선 탐색
"""

import ast, csv, json, os, re, sys
from datetime import datetime, timezone, timedelta
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup

KST = timezone(timedelta(hours=9))
NOW = datetime.now(KST)
ROOT = os.getcwd()
CACHE_PATH = os.path.join(ROOT, "board_cache.json")
OUT_PATH = os.path.join(ROOT, "게시판_자동발견_결과.csv")
LOG_PATH = os.path.join(ROOT, "게시판_자동발견_진단.json")
MIN_TARGETS = 300
UA = "Mozilla/5.0 (compatible; TikkleBoardDiscovery/8.15.2)"

session = requests.Session()
session.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.5"})

BOARD_WORDS = ["공지사항","공지","알림","알림마당","새소식","기관소식","소식","게시판","국민참여","시민참여","설문","참여"]
EXCLUDE_WORDS = ["입찰","채용","인사","개인정보","사전정보공표","공시","계약","법령","규정","이사회","재무","감사","자료실"]

URL_RE = re.compile(r'https?://[^\s\'"<>\\]+', re.I)

def norm(v): return str(v or "").strip()

def normalize_url(url):
    """
    문자열을 URL로 정규화하되 잘못된 URL은 예외를 발생시키지 않고 ""로 버린다.
    V8.15.3의 Invalid IPv6 URL 오류 방지.
    """
    url = norm(url).strip(" ,;)]}")
    if not url:
        return ""
    # 흔한 Excel/문자열 잔여물 제거
    url = url.replace("\\n", "").replace("\\r", "").strip()
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    try:
        p = urlparse(url)
        if not p.netloc:
            return ""
        # 잘못된 IPv6/포트 등 urlparse가 내부적으로 평가하는 값도 여기서 차단
        _ = p.hostname
        _ = p.port
        return urlunparse((p.scheme, p.netloc, p.path or "/", p.params, p.query, ""))
    except (ValueError, TypeError):
        return ""

def canonical_board_url(url):
    p = urlparse(url); q = parse_qs(p.query)
    if "bbsId" in q:
        return urlunparse((p.scheme,p.netloc,p.path,"",urlencode({"bbsId":q["bbsId"][0]}),""))
    return normalize_url(url)

def load_cache():
    try:
        with open(CACHE_PATH, encoding="utf-8") as f:
            d=json.load(f)
        return d if isinstance(d,dict) else {}
    except Exception:
        return {}

def url_like(v):
    u = normalize_url(v)
    if not u:
        return False
    try:
        p = urlparse(u)
        host = p.hostname or ""
        if not host or "." not in host:
            return False
        if p.scheme not in ("http", "https"):
            return False
        if u.lower().endswith((".png",".jpg",".gif",".css",".js")):
            return False
        return True
    except (ValueError, TypeError):
        return False

def institution_like(v):
    s=norm(v)
    if not s or url_like(s): return False
    if len(s)>120: return False
    # 기관명은 한글/영문/숫자를 포함하고 너무 일반적인 상태값은 제외
    if len(re.findall(r"[가-힣]",s)) < 2 and not re.search(r"[A-Za-z]{3}",s): return False
    bad=["정상","오류","미확인","게시판","상태","검증","없음","true","false"]
    if s.lower() in {x.lower() for x in bad}: return False
    return True

def csv_candidates():
    out=[]
    for fn in os.listdir(ROOT):
        if not fn.lower().endswith(".csv"): continue
        path=os.path.join(ROOT,fn)
        try:
            with open(path,encoding="utf-8-sig",newline="") as f:
                reader=csv.reader(f)
                rows=list(reader)
            if not rows: continue
            header=rows[0]
            data=rows[1:]
            if len(data)<MIN_TARGETS: continue
            cols=list(zip(*data)) if data else []
            for idx,col in enumerate(cols):
                url_count=sum(url_like(x) for x in col)
                name_count=sum(institution_like(x) for x in col)
                if url_count>=MIN_TARGETS:
                    out.append((path,"url",idx,url_count,header[idx] if idx<len(header) else ""))
                if name_count>=MIN_TARGETS:
                    out.append((path,"name",idx,name_count,header[idx] if idx<len(header) else ""))
        except Exception:
            pass
    return out

def load_from_csv():
    candidates=csv_candidates()
    urls=[x for x in candidates if x[1]=="url"]
    names=[x for x in candidates if x[1]=="name"]
    if not urls or not names: return None

    # 같은 파일 우선, 그중 URL/기관명 쌍의 유효행 최대화
    best=None
    for u in urls:
        for n in names:
            if u[0]!=n[0] or u[2]==n[2]: continue
            try:
                with open(u[0],encoding="utf-8-sig",newline="") as f:
                    rows=list(csv.reader(f))
                score=0
                pairs=[]
                for row in rows[1:]:
                    if max(u[2],n[2])>=len(row): continue
                    name=norm(row[n[2]]); url=normalize_url(row[u[2]])
                    if institution_like(name) and url_like(url):
                        pairs.append((name,url))
                score=len(pairs)
                if best is None or score>best[0]:
                    best=(score,u,n,pairs)
            except Exception:
                pass
    if best and best[0]>=MIN_TARGETS:
        return best[3], f"CSV:{os.path.basename(best[1][0])}"
    return None

def flatten_constants(node):
    vals=[]
    if isinstance(node,ast.Constant) and isinstance(node.value,str):
        vals.append(node.value)
    elif isinstance(node,(ast.List,ast.Tuple,ast.Set)):
        for x in node.elts: vals += flatten_constants(x)
    elif isinstance(node,ast.Dict):
        for k,v in zip(node.keys,node.values):
            vals += flatten_constants(k); vals += flatten_constants(v)
    return vals

def source_pairs(path):
    try:
        src=open(path,encoding="utf-8").read()
        tree=ast.parse(src)
    except Exception:
        return []
    strings=[]
    for node in ast.walk(tree):
        if isinstance(node,ast.Constant) and isinstance(node.value,str):
            strings.append(node.value)
    urls=[normalize_url(x) for x in strings if url_like(x)]
    urls=[x for x in urls if x]
    # 각 URL 주변/동일 데이터 구조에서 기관명을 찾기 위한 문자열 후보
    names=[x.strip() for x in strings if institution_like(x)]
    pairs=[]
    # 같은 상수 컬렉션/딕셔너리의 name/url을 우선
    for node in ast.walk(tree):
        if isinstance(node,ast.Dict):
            kv={}
            for k,v in zip(node.keys,node.values):
                if isinstance(k,ast.Constant) and isinstance(k.value,str):
                    vals=flatten_constants(v)
                    kv[k.value]=vals
            urlvals=[]
            namevals=[]
            for k,vals in kv.items():
                kl=k.lower()
                if any(z in kl for z in ["url","homepage","site","website","주소"]):
                    urlvals += [normalize_url(v) for v in vals if url_like(v)]
                if any(z in kl for z in ["기관","name","org","target"]):
                    namevals += [v for v in vals if institution_like(v)]
            if urlvals and namevals:
                for u in urlvals:
                    for n in namevals[:5]:
                        pairs.append((n,u))
    # URL 문자열에 대응하는 가장 가까운 한글 문자열: 같은 소스에 기관명과 URL이 충분한 경우
    if len(urls)>=MIN_TARGETS and len(names)>=MIN_TARGETS:
        # URL이 들어있는 문자열의 도메인과 이름 후보를 순서대로 1:1 매칭
        # 실제 monitor.py가 리스트/튜플이면 대개 동일 순서이므로 이를 보조수단으로 사용
        pairs += list(zip(names[:len(urls)], urls[:len(names)]))
    # 중복 제거
    seen=set(); out=[]
    for n,u in pairs:
        k=(n,u)
        if k not in seen:
            seen.add(k); out.append(k)
    return out

def load_from_python():
    candidates=[]
    for fn in os.listdir(ROOT):
        if not fn.endswith(".py"): continue
        path=os.path.join(ROOT,fn)
        pairs=source_pairs(path)
        if len(pairs)>=MIN_TARGETS:
            candidates.append((len(pairs),path,pairs))
    if not candidates: return None
    candidates.sort(reverse=True)
    return candidates[0][2], f"PYTHON:{os.path.basename(candidates[0][1])}"

def load_targets():
    result=load_from_csv()
    if result: return result
    result=load_from_python()
    if result: return result
    raise RuntimeError(
        "355기관 데이터를 자동 복원하지 못했습니다. "
        "CSV의 헤더명에 의존하지 않고 값 기반/monitor.py 분석까지 수행했지만 "
        f"{MIN_TARGETS}개 이상 기관-URL 쌍을 확보하지 못했습니다."
    )

def fetch(url,timeout=10):
    try:
        r=session.get(url,timeout=timeout,allow_redirects=True)
        if r.status_code>=400: return None,r.url,""
        r.encoding=r.apparent_encoding or r.encoding
        return r,r.url,r.text
    except Exception:
        return None,url,""

def score_candidate(url,anchor,title):
    t=f"{url} {anchor} {title}".lower()
    s=0
    if "selectboardlist.do" in t: s+=60
    if "bbsid=bbsMSTR_".lower() in t: s+=30
    for w in BOARD_WORDS:
        if w.lower() in anchor.lower(): s+=15
    for w in EXCLUDE_WORDS:
        if w.lower() in anchor.lower(): s-=35
    return s

def extract_candidates(base,html):
    soup=BeautifulSoup(html,"html.parser")
    title=soup.title.get_text(" ",strip=True) if soup.title else ""
    out=[]
    for a in soup.find_all("a",href=True):
        href=norm(a.get("href"))
        if not href or href.lower().startswith(("javascript:","#","mailto:")): continue
        u=normalize_url(urljoin(base,href)); txt=a.get_text(" ",strip=True)
        if "selectboardlist.do" in u.lower() or "bbsid=" in u.lower() or any(w in txt for w in BOARD_WORDS):
            out.append((score_candidate(u,txt,title),canonical_board_url(u),txt))
    for m in re.finditer(r"""[^"'<>\\s]{0,150}/cop/bbs/selectBoardList\.do\?[^"'<>\\s]+""",html,re.I):
        u=normalize_url(urljoin(base,m.group(0)))
        out.append((score_candidate(u,"",title),canonical_board_url(u),""))
    best={}
    for x in out:
        if x[1] not in best or x[0]>best[x[1]][0]: best[x[1]]=x
    return sorted(best.values(),reverse=True)

def verify(url,name):
    r,final,html=fetch(url)
    if not r or not html: return False,{"reason":"FETCH_FAIL","url":final}
    soup=BeautifulSoup(html,"html.parser"); title=soup.title.get_text(" ",strip=True) if soup.title else ""
    text=soup.get_text(" ",strip=True)
    sig=0
    if "selectboardlist.do" in final.lower(): sig+=2
    if "bbsid=" in final.lower(): sig+=2
    if re.search(r"번호\s+제목|제목\s+작성자|작성자\s+날짜|조회",text): sig+=2
    if re.search(r"\d{4}[-./]\d{1,2}[-./]\d{1,2}",text): sig+=1
    if name.replace(" ","")[:6] in text.replace(" ",""): sig+=1
    return sig>=3,{"reason":"VERIFIED" if sig>=3 else "LOW_SIGNAL","url":canonical_board_url(final),"title":title[:200],"signals":sig}

def discover(name,home,cache):
    old=cache.get(name)
    if isinstance(old,dict):
        oldurl=old.get("board_url") or old.get("url")
        if oldurl:
            ok,info=verify(oldurl,name)
            if ok: return oldurl,"CACHE_VERIFIED",info
    r,final,html=fetch(home)
    if not r or not html: return "","HOME_ERROR",{"url":final}
    candidates=extract_candidates(final,html)
    soup=BeautifulSoup(html,"html.parser")
    children=[]
    for a in soup.find_all("a",href=True):
        txt=a.get_text(" ",strip=True)
        if not any(w in txt for w in BOARD_WORDS): continue
        u=normalize_url(urljoin(final,a["href"]))
        if urlparse(u).netloc==urlparse(final).netloc and u not in children: children.append(u)
    for child in children[:10]:
        r2,f2,h2=fetch(child)
        if r2 and h2: candidates.extend(extract_candidates(f2,h2))
    best={}
    for x in candidates:
        if x[1] not in best or x[0]>best[x[1]][0]: best[x[1]]=x
    ranked=sorted(best.values(),reverse=True)
    for sc,u,txt in ranked:
        if "selectboardlist.do" in u.lower() and "bbsid=" in u.lower():
            ok,info=verify(u,name)
            if ok: return u,"DISCOVERED_EGOV",info
    for sc,u,txt in ranked:
        if sc<25: continue
        ok,info=verify(u,name)
        if ok: return u,"DISCOVERED_GENERAL",info
    return "","CANDIDATE_NOT_VERIFIED",{"candidate_count":len(ranked),"candidates":[{"score":x[0],"url":x[1],"text":x[2][:100]} for x in ranked[:8]]}

def save_json(path,obj):
    tmp=path+".tmp"
    with open(tmp,"w",encoding="utf-8") as f: json.dump(obj,f,ensure_ascii=False,indent=2)
    os.replace(tmp,path)

def main():
    targets,source=load_targets()
    if len(targets)<MIN_TARGETS: raise RuntimeError(f"안전장치: 기관 수 {len(targets)} < {MIN_TARGETS}")
    # 이름 기준 중복 제거
    uniq={}
    malformed = 0
    for n,u in targets:
        if institution_like(n) and url_like(u):
            uniq[n]=(n,u)
        else:
            malformed += 1
    if malformed:
        print(f"[URL 필터] 잘못된 기관-URL 후보 {malformed}개 제외")
    targets=list(uniq.values())
    if len(targets)<MIN_TARGETS: raise RuntimeError(f"안전장치: 중복 제거 후 기관 수 {len(targets)} < {MIN_TARGETS}")
    print(f"[기관목록] source={source}, targets={len(targets)}")
    cache=load_cache(); rows=[]; counts={}; found=0
    for i,(name,home) in enumerate(targets,1):
        try:
            board,status,info=discover(name,home,cache)
        except (ValueError, TypeError) as e:
            board = ""
            status = "URL_ERROR"
            info = {"reason": type(e).__name__, "message": str(e), "homepage": home}
        except Exception as e:
            board = ""
            status = "DISCOVERY_ERROR"
            info = {"reason": type(e).__name__, "message": str(e), "homepage": home}
        if board:
            old=cache.get(name,{})
            cache[name]={**(old if isinstance(old,dict) else {}),"institution":name,"homepage":home,"board_url":board,"url":board,"verified":True,"discovery_status":status,"discovered_at":NOW.isoformat(),"source":"V8.15.3"}
            found+=1
        counts[status]=counts.get(status,0)+1
        rows.append({"기관명":name,"홈페이지":home,"게시판URL":board,"상태":status,"진단":json.dumps(info,ensure_ascii=False)[:1600]})
        if i%25==0: print(f"[{i}/{len(targets)}] 확보 {found}")
    save_json(CACHE_PATH,cache)
    with open(OUT_PATH,"w",encoding="utf-8-sig",newline="") as f:
        w=csv.DictWriter(f,fieldnames=["기관명","홈페이지","게시판URL","상태","진단"]); w.writeheader(); w.writerows(rows)
    diag={"version":"V8.15.3","updated_at":NOW.isoformat(),"source":source,"targets":len(targets),"found_or_verified":found,"not_found":len(targets)-found,"status_counts":counts}
    save_json(LOG_PATH,diag); print(json.dumps(diag,ensure_ascii=False,indent=2))

if __name__=="__main__":
    try: main()
    except Exception as e:
        print("FATAL:",e,file=sys.stderr); sys.exit(1)
