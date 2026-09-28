# -*- coding: utf-8 -*-
"""Public Institution Monitor V8.14.18
355기관 본체를 건드리지 않고 미확인 기관만 진단/복구 후보를 탐색한다.
"""
import csv, json, re, time
from datetime import datetime
from urllib.parse import urlparse, urlunparse, urljoin
import requests
from bs4 import BeautifulSoup

VERSION='V8.14.18-DIAG'
INPUT='기관별_상태.csv'
OUT_CSV='기관별_미확인120_진단.csv'
OUT_JSON='기관별_미확인120_진단.json'
TIMEOUT=8
RETRIES=1
COMMON_PATHS=['/notice','/board','/bbs','/news','/community','/participation','/notice/list','/board/list','/bbs/list','/boardList.do','/bbsList.do','/board/list.do','/bbs/list.do','/contents/board','/site/board']
KEYWORDS=['공지사항','공지','새소식','알림마당','알림','설문','국민참여','시민참여','참여마당','소통','게시판','뉴스']

S=requests.Session(); S.headers.update({'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128 Safari/537.36','Accept-Language':'ko-KR,ko;q=0.9,en;q=0.7'})

def get(u):
    for i in range(RETRIES+1):
        try:
            r=S.get(u,timeout=TIMEOUT,allow_redirects=True,verify=True)
            if 200<=r.status_code<400:
                if not r.encoding or r.encoding.lower()=='iso-8859-1': r.encoding=r.apparent_encoding or 'utf-8'
                return r,None
            err=f'HTTP_{r.status_code}'
        except requests.exceptions.Timeout: err='TIMEOUT'
        except requests.exceptions.SSLError: err='SSL_ERROR'
        except requests.exceptions.ConnectionError: err='CONNECTION_ERROR'
        except Exception as e: err=type(e).__name__
        if i<RETRIES: time.sleep(.4)
    return None,err

def variants(home):
    try:
        p=urlparse(home); host=p.netloc.split(':')[0]; basepath=p.path or '/'
        hosts=[]
        for h in [host, host[4:] if host.startswith('www.') else 'www.'+host]:
            if h and h not in hosts: hosts.append(h)
        schemes=['https','http']
        out=[]
        for sc in schemes:
            for h in hosts:
                u=urlunparse((sc,h,basepath,'','',''))
                if u not in out: out.append(u)
        rootset=[]
        for sc in schemes:
            for h in hosts:
                root=urlunparse((sc,h,'/','','',''))
                rootset.append(root)
        for root in rootset:
            if root not in out: out.append(root)
        return out[:8]
    except Exception: return [home]

def norm(x): return re.sub(r'\s+',' ',str(x or '')).strip()

def same_domain(a,b):
    try: return urlparse(a).netloc.split(':')[0].removeprefix('www.')==urlparse(b).netloc.split(':')[0].removeprefix('www.')
    except: return False

def link_candidates(base,soup):
    out=[]; seen=set()
    for a in soup.find_all('a',href=True):
        href=a.get('href','').strip()
        if not href or href.lower().startswith(('javascript:','#','mailto:','tel:')): continue
        u=urljoin(base,href).split('#')[0]
        if not same_domain(base,u) or u in seen: continue
        t=norm(' '.join([a.get_text(' ',strip=True),a.get('title',''),a.get('aria-label','')]))
        score=0; s=(u+' '+t).lower()
        score += sum(15 for k in KEYWORDS if k.lower() in s)
        score += sum(3 for k in ['notice','board','bbs','news','list','ntt','article'] if k in u.lower())
        if score>0: out.append((score,u,t[:150])) ; seen.add(u)
    return sorted(out,reverse=True)[:15]

def inspect_candidate(u):
    r,err=get(u)
    if not r: return {'url':u,'status':'FETCH_FAIL','error':err,'posts':0}
    soup=BeautifulSoup(r.text,'html.parser'); text=norm(soup.get_text(' ',strip=True))[:30000]
    title_links=[]
    for a in soup.find_all('a',href=True):
        t=norm(a.get_text(' ',strip=True))
        if 2<=len(t)<=300 and any(x in (a.get('href','').lower()+t.lower()) for x in ['view','read','article','nttid','ntt','seq=','idx=','detail']): title_links.append(t)
    rows=len(soup.select('table tr'))
    date_signal=bool(re.search(r'(20\\d{2}[./-]\\d{1,2}[./-]\\d{1,2}|등록일|작성일|게시일)',text))
    list_signal=rows>=3 or bool(re.search(r'(공지사항|게시물|목록|총\\s*\\d+\\s*건)',text))
    return {'url':r.url,'status':'BOARD_LIKE' if list_signal else 'PAGE','error':'','posts':min(len(title_links),80),'rows':rows,'date_signal':date_signal,'keyword_signal':any(k in text for k in KEYWORDS),'title':norm(soup.title.get_text(' ',strip=True)) if soup.title else ''}

def diagnose(row):
    home=row['홈페이지URL']; out={'기관명':row['기관명'],'기관유형':row.get('기관유형',''),'원래URL':home,'기존상태':row['상태'],'복구URL':'','홈페이지접속':'','홈페이지오류':'','후보수':0,'검증후보':'','판정':'','근거':'','후보목록':''}
    best=None; errs=[]
    for u in variants(home):
        r,e=get(u)
        if r:
            out['홈페이지접속']='성공'; out['복구URL']=r.url if r.url!=home else ''
            soup=BeautifulSoup(r.text,'html.parser')
            cs=link_candidates(r.url,soup)
            out['후보수']=len(cs)
            candidates=cs[:]
            for path in COMMON_PATHS:
                x=urljoin(r.url,path)
                if same_domain(r.url,x) and all(x!=z[1] for z in candidates): candidates.append((1,x,'common-path'))
            checks=[]
            for sc,x,t in candidates[:20]:
                q=inspect_candidate(x); q['score']=sc; q['label']=t; checks.append(q)
            verified=[q for q in checks if q.get('status')=='BOARD_LIKE' and (q.get('posts',0)>0 or q.get('date_signal'))]
            if verified:
                v=max(verified,key=lambda q:(q.get('posts',0),q.get('score',0)))
                best=v; out['검증후보']=v['url']; out['판정']='RECOVERED_CANDIDATE'; out['근거']=f"게시판형 구조/게시물 후보 {v.get('posts',0)}건"
                out['후보목록']=' | '.join(q['url'] for q in checks[:10]); return out
            if checks and best is None:
                best=max(checks,key=lambda q:(q.get('posts',0),q.get('score',0)))
                out['검증후보']=best['url']
                out['후보목록']=' | '.join(q['url'] for q in checks[:10])
        else: errs.append(f'{u}:{e}')
    if out['홈페이지접속']!='성공':
        out['홈페이지접속']='실패'; out['홈페이지오류']='; '.join(errs[:5]); out['판정']='HOME_UNRECOVERED'; out['근거']='홈페이지 변형 URL까지 접속 실패'
    elif best:
        out['판정']='CANDIDATE_NEEDS_RULE_REVIEW'; out['근거']='후보 페이지는 있으나 현재 검증기준으로 게시판 확정 불가'
    else:
        out['판정']='NO_BOARD_FOUND'; out['근거']='홈페이지 접근 성공했으나 게시판 후보 부족'
    return out

def main():
    with open(INPUT,encoding='utf-8-sig',newline='') as f: rows=list(csv.DictReader(f))
    targets=[r for r in rows if r.get('상태') in ('HOME_ERROR','CANDIDATE_NOT_VERIFIED')]
    results=[]
    for i,r in enumerate(targets,1):
        results.append(diagnose(r)); print(f'[{i}/{len(targets)}] {r["기관명"]} -> {results[-1]["판정"]}')
    fields=['기관명','기관유형','원래URL','기존상태','복구URL','홈페이지접속','홈페이지오류','후보수','검증후보','판정','근거','후보목록']
    with open(OUT_CSV,'w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=fields); w.writeheader(); w.writerows(results)
    summary={}
    for x in results: summary[x['판정']]=summary.get(x['판정'],0)+1
    with open(OUT_JSON,'w',encoding='utf-8') as f: json.dump({'version':VERSION,'generated_at':datetime.now().isoformat(),'targets':len(targets),'summary':summary,'results':results},f,ensure_ascii=False,indent=2)
    print(json.dumps(summary,ensure_ascii=False,indent=2))

if __name__=='__main__': main()
