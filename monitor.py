import asyncio, aiohttp, pandas as pd, re, json, os, hashlib, requests
from bs4 import BeautifulSoup
from datetime import datetime
from urllib.parse import urljoin, urlparse, urldefrag

KEYWORDS=['설문조사','시민참여','국민참여','공모전']
BOARD_WORDS=['공지사항','공지','알림마당','알림','새소식','소식','게시판','국민참여','시민참여','참여','공모전','공모','설문조사','설문']
URL_FILE='url_완성.xlsx'; INPUT_FILE='url.xlsx'; BOARDS_FILE='boards.xlsx'; STATE_FILE='seen_posts.json'
CONCURRENCY=20; TIMEOUT=20
UA='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140 Safari/537.36'

def norm(s): return re.sub(r'\s+',' ',str(s or '')).strip()
def nname(s): return re.sub(r'\s+','',norm(s).replace('㈜','(주)'))
def clean_url(x):
    if pd.isna(x): return ''
    x=str(x).strip(); m=re.search(r'\]\((https?://[^)]+)\)',x)
    if m: x=m.group(1)
    x=x.strip('[]() ')
    return x if re.match(r'^https?://',x,re.I) else ('https://'+x if x else '')
def canon(u): return urldefrag(u or '')[0].rstrip('/')
def same_domain(a,b): return urlparse(a).netloc.lower().lstrip('www.')==urlparse(b).netloc.lower().lstrip('www.')
def date_from(s):
    m=re.search(r'(20\d{2})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})',str(s or ''))
    if not m: return ''
    try: return f'{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    except: return ''
def hits(s):
    t=norm(s).lower(); return [k for k in KEYWORDS if k.lower() in t]

async def fetch(session,url):
    try:
        async with session.get(url,timeout=aiohttp.ClientTimeout(total=TIMEOUT),allow_redirects=True,ssl=False) as r:
            if r.status>=400: return None,str(r.url)
            raw=await r.read(); return raw.decode(r.charset or 'utf-8',errors='ignore'),str(r.url)
    except Exception: return None,url

def score_link(text,href):
    c=(norm(text)+' '+href).lower(); score=0
    for w in BOARD_WORDS:
        if w.lower() in c: score+=2
    if re.search(r'(view|detail|read|articleNo|nttId|bbsno|boardno|seq=|idx=)',href,re.I): score-=4
    if re.search(r'(list|board|notice|news)',href,re.I): score+=2
    return score

async def discover(session,name,home):
    html,final=await fetch(session,home)
    if not html: return []
    soup=BeautifulSoup(html,'html.parser'); cand={}
    for a in soup.find_all('a',href=True):
        href=a.get('href','').strip(); text=a.get_text(' ',strip=True)
        if not href or href.lower().startswith(('javascript:','mailto:','#')): continue
        u=canon(urljoin(final,href))
        if not same_domain(u,final): continue
        s=score_link(text,href)
        if s>=2: cand[u]=(s,norm(text)[:100])
    return [{'기관명':name,'게시판명':v[1],'게시판URL':u,'점수':v[0]} for u,v in sorted(cand.items(),key=lambda x:x[1][0],reverse=True)[:5]]

def parse_posts(html,board):
    soup=BeautifulSoup(html,'html.parser'); out={}
    for tr in soup.find_all('tr'):
        row=norm(tr.get_text(' ',strip=True)); d=date_from(row)
        for a in tr.find_all('a',href=True):
            title=norm(a.get_text(' ',strip=True)); h=a['href']
            if len(title)<2 or h.lower().startswith(('javascript:','mailto:','#')): continue
            u=canon(urljoin(board,h))
            if same_domain(u,board): out[u]={'title':title,'url':u,'date':d}
    if len(out)<3:
        for a in soup.find_all('a',href=True):
            title=norm(a.get_text(' ',strip=True)); h=a['href']
            if len(title)<4 or h.lower().startswith(('javascript:','mailto:','#')): continue
            u=canon(urljoin(board,h))
            if same_domain(u,board) and re.search(r'(view|detail|articleNo|nttId|bbsno|seq=|idx=)',u,re.I): out[u]={'title':title,'url':u,'date':''}
    return list(out.values())[:20]

async def inspect(session,p):
    html,final=await fetch(session,p['url'])
    if not html:return None
    soup=BeautifulSoup(html,'html.parser')
    for x in soup(['script','style','noscript']): x.decompose()
    body=norm(soup.get_text(' ',strip=True)); ks=hits(p['title']+' '+body)
    if not ks:return None
    return {'url':canon(final),'title':p['title'],'date':p['date'] or date_from(body[:5000]),'keywords':ks}

def state_load():
    try:
        with open(STATE_FILE,encoding='utf-8') as f:return json.load(f)
    except:return {}
def state_save(s):
    with open(STATE_FILE,'w',encoding='utf-8') as f:json.dump(s,f,ensure_ascii=False,indent=2)

def make_url_file():
    if os.path.exists(URL_FILE): return
    df=pd.read_excel(INPUT_FILE)
    year=datetime.now().year
    u=f'https://alio.go.kr/download/statisticsDown.json?f=%EC%9D%BC%EB%B0%98%ED%98%84%ED%99%A9_{year}.xlsx&s=general_status_{year}.xlsx'
    r=requests.get(u,timeout=30); r.raise_for_status(); open(f'alio_{year}.xlsx','wb').write(r.content)
    raw=pd.read_excel(f'alio_{year}.xlsx',sheet_name='일반현황',header=None)
    hr=next(i for i in range(min(10,len(raw))) if '기관명' in raw.iloc[i].astype(str).tolist())
    alio=pd.read_excel(f'alio_{year}.xlsx',sheet_name='일반현황',header=hr)
    mp={nname(r['기관명']):clean_url(r['홈페이지']) for _,r in alio.iterrows() if clean_url(r.get('홈페이지',''))}
    df['URL']=[mp.get(nname(x),'') for x in df['기관명']]; df.to_excel(URL_FILE,index=False)

async def discover_all():
    make_url_file(); df=pd.read_excel(URL_FILE); rows=[]
    conn=aiohttp.TCPConnector(limit=CONCURRENCY,ssl=False)
    async with aiohttp.ClientSession(connector=conn,headers={'User-Agent':UA}) as s:
        sem=asyncio.Semaphore(CONCURRENCY)
        async def one(r):
            async with sem:return await discover(s,r['기관명'],clean_url(r['URL'])) if clean_url(r['URL']) else []
        for group in await asyncio.gather(*[one(r) for _,r in df.iterrows()]): rows+=group
    out=pd.DataFrame(rows).drop_duplicates(['기관명','게시판URL']) if rows else pd.DataFrame(columns=['기관명','게시판명','게시판URL','점수'])
    out.to_excel(BOARDS_FILE,index=False); print(f'게시판 후보 {len(out)}개 저장: {BOARDS_FILE}')

async def monitor():
    if not os.path.exists(BOARDS_FILE): await discover_all()
    boards=pd.read_excel(BOARDS_FILE).fillna(''); state=state_load(); found=[]
    conn=aiohttp.TCPConnector(limit=CONCURRENCY,ssl=False)
    async with aiohttp.ClientSession(connector=conn,headers={'User-Agent':UA}) as s:
        sem=asyncio.Semaphore(CONCURRENCY)
        async def one(r):
            async with sem:
                h,_=await fetch(s,r['게시판URL'])
                if not h:return []
                res=[]
                for p in parse_posts(h,r['게시판URL']):
                    x=await inspect(s,p)
                    if x:
                        k=hashlib.sha256(x['url'].encode()).hexdigest()
                        if k not in state: x.update({'key':k,'기관명':r['기관명'],'게시판명':r['게시판명']}); res.append(x)
                return res
        for group in await asyncio.gather(*[one(r) for _,r in boards.iterrows()]): found+=group
    token=os.getenv('TELEGRAM_BOT_TOKEN'); chat=os.getenv('TELEGRAM_CHAT_ID')
    for x in found:
        msg=f"🔔 신규 키워드 게시글\n\n기관: {x['기관명']}\n게시판: {x['게시판명']}\n키워드: {', '.join(x['keywords'])}\n제목: {x['title']}\n등록일: {x['date'] or '확인불가'}\n링크: {x['url']}"
        if token and chat:
            try:
                rr=requests.post(f'https://api.telegram.org/bot{token}/sendMessage',data={'chat_id':chat,'text':msg,'disable_web_page_preview':True},timeout=20); rr.raise_for_status()
            except Exception as e: print('Telegram 실패',e); continue
        state[x['key']]={'기관명':x['기관명'],'title':x['title'],'url':x['url'],'date':x['date'],'saved_at':datetime.now().isoformat()}
    state_save(state); print('신규 알림 대상:',len(found))

if __name__=='__main__': asyncio.run(discover_all() if os.getenv('MODE')=='discover' else monitor())
