import asyncio, aiohttp, json, os, re, time
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urldefrag
import pandas as pd

INPUT="v8_3_notice_candidates.xlsx"
OUTPUT="v8_3_notice_validation.xlsx"
TIMEOUT=int(os.getenv("TIMEOUT_SECONDS","15"))
CONCURRENCY=int(os.getenv("MAX_CONCURRENCY","15"))
POST_CHECK=int(os.getenv("POST_CHECK","8"))

DETAIL_PAT=re.compile(r"(?:mode=view|act=view|view\\.do|detail\\.do|detail\\.asp|articleview|boardview|selectpstinfo|/view[/?])",re.I)
FUNCTION_PAT=re.compile(r"(login|search|sitemap|calendar|reservation|reservation|map|location|print|share|download|file|member|mypage|survey|payment)",re.I)
LIST_WORDS=["공지사항","공지","알림","새소식","기관소식","소식","게시판","목록","등록일","작성일","제목","조회"]
DATE_PAT=re.compile(r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2}|\d{4}[./-]\d{1,2}[./-]\d{1,2})")

def clean(s):
    return re.sub(r"\s+"," ",str(s or "")).strip()

def norm_url(base, href):
    if not href: return ""
    href=href.strip()
    if href.startswith(("javascript:","#","mailto:","tel:")): return ""
    return urldefrag(urljoin(base,href))[0]

def is_bad_link(href, text=""):
    h=(href or "").lower()
    t=clean(text).lower()
    if FUNCTION_PAT.search(h) and not DETAIL_PAT.search(h): return True
    if t in ["검색","로그인","회원가입","사이트맵","전체메뉴","인쇄","공유","더보기","다음","이전","목록"]: return True
    return False

def title_from_row(a):
    txt=clean(a.get_text(" ",strip=True))
    for x in ["새글","공지","N","NEW"]:
        txt=re.sub(rf"^{re.escape(x)}\s+","",txt,flags=re.I)
    return txt[:300]

def extract_posts(url, html):
    soup=BeautifulSoup(html,"html.parser")
    for tag in soup(["script","style","noscript","svg"]): tag.decompose()
    posts=[]
    seen=set()

    # 반복 행 우선
    rows=soup.select("table tr, ul li, ol li, .list li, .board_list li, .bbs_list li, .notice_list li")
    for row in rows:
        links=[]
        for a in row.find_all("a",href=True):
            href=norm_url(url,a.get("href"))
            title=title_from_row(a)
            if not href or is_bad_link(a.get("href"),title): continue
            if len(title)<4 or len(title)>300: continue
            if not DETAIL_PAT.search(href): continue
            if href in seen: continue
            seen.add(href)
            txt=clean(row.get_text(" ",strip=True))
            if not (DATE_PAT.search(txt) or any(w in txt for w in LIST_WORDS)):
                continue
            posts.append({"title":title,"url":href})
            if len(posts)>=POST_CHECK: return posts

    # fallback: 상세형 링크 중 제목다운 텍스트
    for a in soup.find_all("a",href=True):
        href=norm_url(url,a.get("href")); title=title_from_row(a)
        if not href or href in seen or not DETAIL_PAT.search(href): continue
        if is_bad_link(a.get("href"),title): continue
        if len(title)<8 or len(title)>250: continue
        seen.add(href); posts.append({"title":title,"url":href})
        if len(posts)>=POST_CHECK: break
    return posts

async def fetch(session,url,sem):
    async with sem:
        try:
            async with session.get(url,timeout=aiohttp.ClientTimeout(total=TIMEOUT),allow_redirects=True,
                                   headers={"User-Agent":"Mozilla/5.0 (compatible; PublicInstitutionMonitor/8.3)"}) as r:
                return r.status, str(r.url), await r.text(errors="ignore")
        except Exception as e:
            return 0,url,""

async def verify_one(session,row,sem):
    institution=str(row["기관명"]); name=str(row["게시판명"]); url=str(row["게시판URL"])
    status,final,html=await fetch(session,url,sem)
    result={
        "기관명":institution,"게시판명":name,"게시판URL":url,"HTTP":status,
        "최종URL":final,"게시물후보":0,"검증게시물":0,"고유제목":0,
        "목록단서":0,"판정":"","사유":"","검증제목":""
    }
    if status!=200 or not html:
        result["판정"]="D-제외"; result["사유"]="접속실패"
        return result

    soup=BeautifulSoup(html,"html.parser")
    visible=clean(soup.get_text(" ",strip=True))
    list_clues=sum(1 for w in LIST_WORDS if w in visible)
    result["목록단서"]=list_clues
    posts=extract_posts(final,html)
    result["게시물후보"]=len(posts)

    verified=[]; titles=[]
    for p in posts[:POST_CHECK]:
        st,fu,body=await fetch(session,p["url"],sem)
        if st!=200 or not body: continue
        b=clean(BeautifulSoup(body,"html.parser").get_text(" ",strip=True))
        t=clean(p["title"])
        # 목록 제목이 상세 본문에 나타나는지 확인
        ok=len(t)>=6 and t[:80].lower() in b.lower()
        if ok:
            verified.append(p); titles.append(t)

    uniq=[]
    for t in titles:
        if t not in uniq: uniq.append(t)
    result["검증게시물"]=len(verified)
    result["고유제목"]=len(uniq)
    result["검증제목"]=" | ".join(uniq[:8])

    if len(verified)>=5 and len(uniq)>=4 and list_clues>=3:
        result["판정"]="A-최종확정"; result["사유"]="실제 게시물 5건 이상 + 서로 다른 제목 4개 이상 + 목록구조 확인"
    elif len(verified)>=3 and len(uniq)>=3 and list_clues>=2:
        result["판정"]="B-자동화가능"; result["사유"]="실제 게시물 3건 이상 + 서로 다른 제목 3개 이상"
    elif len(verified)>=2 or (len(posts)>=3 and list_clues>=2):
        result["판정"]="C-추가분석"; result["사유"]="게시판 후보 증거는 있으나 자동확정 기준 미달"
    else:
        result["판정"]="D-제외"; result["사유"]="독립 게시판/실제 게시물 증거 부족"
    return result

async def main():
    df=pd.read_excel(INPUT,sheet_name="공지후보")
    sem=asyncio.Semaphore(CONCURRENCY)
    connector=aiohttp.TCPConnector(limit=CONCURRENCY,ssl=False)
    async with aiohttp.ClientSession(connector=connector) as session:
        tasks=[verify_one(session,row,sem) for _,row in df.iterrows()]
        results=[]
        for coro in asyncio.as_completed(tasks):
            results.append(await coro)
    out=pd.DataFrame(results)
    with pd.ExcelWriter(OUTPUT,engine="openpyxl") as w:
        out.to_excel(w,index=False,sheet_name="전체검증")
        for s,cond in [
            ("A_최종확정",out["판정"]=="A-최종확정"),
            ("B_자동화가능",out["판정"]=="B-자동화가능"),
            ("C_추가분석",out["판정"]=="C-추가분석"),
            ("D_제외",out["판정"]=="D-제외")
        ]: out[cond].to_excel(w,index=False,sheet_name=s)
        summary=pd.DataFrame({
            "판정":["A-최종확정","B-자동화가능","C-추가분석","D-제외","전체"],
            "건수":[int((out["판정"]=="A-최종확정").sum()),int((out["판정"]=="B-자동화가능").sum()),
                    int((out["판정"]=="C-추가분석").sum()),int((out["판정"]=="D-제외").sum()),len(out)],
        })
        summary.to_excel(w,index=False,sheet_name="요약")
    print(json.dumps({"total":len(out),
                      "A":int((out["판정"]=="A-최종확정").sum()),
                      "B":int((out["판정"]=="B-자동화가능").sum()),
                      "C":int((out["판정"]=="C-추가분석").sum()),
                      "D":int((out["판정"]=="D-제외").sum())},ensure_ascii=False))

if __name__=="__main__":
    asyncio.run(main())
