import asyncio
import aiohttp
import pandas as pd
import re
import json
import os
import hashlib
import requests

from bs4 import BeautifulSoup
from datetime import datetime
from urllib.parse import urljoin, urlparse, urldefrag


# ============================================================
# 기본 설정
# ============================================================

KEYWORDS = ['설문조사', '시민참여', '국민참여', '공모전']

BOARD_WORDS = [
    '공지사항', '공지', '알림마당', '알림', '새소식', '소식',
    '게시판', '국민참여', '시민참여', '참여', '공모전',
    '공모', '설문조사', '설문'
]

# GitHub 저장소에 이미 올려놓은 완성 URL 파일
URL_FILE = 'url_완성.xlsx'

# 게시판 후보 저장 파일
BOARDS_FILE = 'boards.xlsx'

# 중복 알림 방지 파일
STATE_FILE = 'seen_posts.json'

CONCURRENCY = 20
TIMEOUT = 20

UA = (
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
    'AppleWebKit/537.36 (KHTML, like Gecko) '
    'Chrome/140 Safari/537.36'
)


# ============================================================
# 공통 함수
# ============================================================

def norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def nname(s):
    return re.sub(
        r'\s+',
        '',
        norm(s).replace('㈜', '(주)')
    )


def clean_url(x):
    if pd.isna(x):
        return ''

    x = str(x).strip()

    m = re.search(
        r'\]\((https?://[^)]+)\)',
        x
    )

    if m:
        x = m.group(1)

    x = x.strip('[]() ')

    if re.match(r'^https?://', x, re.I):
        return x

    return 'https://' + x if x else ''


def canon(u):
    return urldefrag(u or '')[0].rstrip('/')


def same_domain(a, b):
    return (
        urlparse(a).netloc.lower().lstrip('www.')
        ==
        urlparse(b).netloc.lower().lstrip('www.')
    )


def date_from(s):
    m = re.search(
        r'(20\d{2})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})',
        str(s or '')
    )

    if not m:
        return ''

    try:
        return (
            f'{int(m.group(1)):04d}-'
            f'{int(m.group(2)):02d}-'
            f'{int(m.group(3)):02d}'
        )
    except Exception:
        return ''


def hits(s):
    t = norm(s).lower()

    return [
        k for k in KEYWORDS
        if k.lower() in t
    ]


# ============================================================
# 웹페이지 요청
# ============================================================

async def fetch(session, url):
    try:
        async with session.get(
            url,
            timeout=aiohttp.ClientTimeout(total=TIMEOUT),
            allow_redirects=True,
            ssl=False
        ) as r:

            if r.status >= 400:
                return None, str(r.url)

            raw = await r.read()

            return (
                raw.decode(
                    r.charset or 'utf-8',
                    errors='ignore'
                ),
                str(r.url)
            )

    except Exception:
        return None, url


# ============================================================
# 게시판 후보 점수 계산
# ============================================================

def score_link(text, href):
    c = (
        norm(text) + ' ' + href
    ).lower()

    score = 0

    for w in BOARD_WORDS:
        if w.lower() in c:
            score += 2

    # 게시글 상세 페이지로 보이는 URL은 감점
    if re.search(
        r'(view|detail|read|articleNo|nttId|bbsno|boardno|seq=|idx=)',
        href,
        re.I
    ):
        score -= 4

    # 게시판/목록 페이지로 보이는 URL은 가점
    if re.search(
        r'(list|board|notice|news)',
        href,
        re.I
    ):
        score += 2

    return score


# ============================================================
# 기관 홈페이지에서 게시판 후보 발견
# ============================================================

async def discover(session, name, home):
    html, final = await fetch(session, home)

    if not html:
        return []

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    cand = {}

    for a in soup.find_all('a', href=True):

        href = a.get('href', '').strip()
        text = a.get_text(' ', strip=True)

        if not href:
            continue

        if href.lower().startswith(
            ('javascript:', 'mailto:', '#')
        ):
            continue

        u = canon(
            urljoin(final, href)
        )

        if not same_domain(u, final):
            continue

        s = score_link(text, href)

        if s >= 2:
            cand[u] = (
                s,
                norm(text)[:100]
            )

    return [
        {
            '기관명': name,
            '게시판명': v[1],
            '게시판URL': u,
            '점수': v[0]
        }
        for u, v in sorted(
            cand.items(),
            key=lambda x: x[1][0],
            reverse=True
        )[:5]
    ]


# ============================================================
# 게시판에서 게시글 목록 추출
# ============================================================

def parse_posts(html, board):
    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    out = {}

    # 일반적인 테이블형 게시판
    for tr in soup.find_all('tr'):

        row = norm(
            tr.get_text(' ', strip=True)
        )

        d = date_from(row)

        for a in tr.find_all('a', href=True):

            title = norm(
                a.get_text(' ', strip=True)
            )

            h = a['href']

            if len(title) < 2:
                continue

            if h.lower().startswith(
                ('javascript:', 'mailto:', '#')
            ):
                continue

            u = canon(
                urljoin(board, h)
            )

            if same_domain(u, board):

                out[u] = {
                    'title': title,
                    'url': u,
                    'date': d
                }

    # 테이블 구조가 아닌 게시판에 대한 보조 탐색
    if len(out) < 3:

        for a in soup.find_all('a', href=True):

            title = norm(
                a.get_text(' ', strip=True)
            )

            h = a['href']

            if len(title) < 4:
                continue

            if h.lower().startswith(
                ('javascript:', 'mailto:', '#')
            ):
                continue

            u = canon(
                urljoin(board, h)
            )

            if (
                same_domain(u, board)
                and re.search(
                    r'(view|detail|articleNo|nttId|bbsno|seq=|idx=)',
                    u,
                    re.I
                )
            ):
                out[u] = {
                    'title': title,
                    'url': u,
                    'date': ''
                }

    return list(out.values())[:20]


# ============================================================
# 게시글 본문 검사
# ============================================================

async def inspect(session, p):

    html, final = await fetch(
        session,
        p['url']
    )

    if not html:
        return None

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    for x in soup(
        ['script', 'style', 'noscript']
    ):
        x.decompose()

    body = norm(
        soup.get_text(' ', strip=True)
    )

    ks = hits(
        p['title'] + ' ' + body
    )

    if not ks:
        return None

    return {
        'url': canon(final),
        'title': p['title'],
        'date': (
            p['date']
            or date_from(body[:5000])
        ),
        'keywords': ks
    }


# ============================================================
# 중복 알림 상태 관리
# ============================================================

def state_load():

    try:
        with open(
            STATE_FILE,
            encoding='utf-8'
        ) as f:
            return json.load(f)

    except Exception:
        return {}


def state_save(s):

    with open(
        STATE_FILE,
        'w',
        encoding='utf-8'
    ) as f:

        json.dump(
            s,
            f,
            ensure_ascii=False,
            indent=2
        )


# ============================================================
# 최초 게시판 발견
#
# 중요:
# ALIO 다운로드 기능을 완전히 제거했습니다.
# GitHub 저장소에 이미 있는 url_완성.xlsx를 직접 사용합니다.
# ============================================================

async def discover_all():

    if not os.path.exists(URL_FILE):

        raise FileNotFoundError(
            f'{URL_FILE} 파일이 없습니다. '
            'GitHub 저장소에 url_완성.xlsx를 올려주세요.'
        )

    df = pd.read_excel(
        URL_FILE
    )

    print(
        f'기관 수: {len(df)}'
    )

    rows = []

    conn = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=False
    )

    async with aiohttp.ClientSession(
        connector=conn,
        headers={'User-Agent': UA}
    ) as s:

        sem = asyncio.Semaphore(
            CONCURRENCY
        )

        async def one(r):

            async with sem:

                url = clean_url(
                    r.get('URL', '')
                )

                if not url:
                    return []

                return await discover(
                    s,
                    r['기관명'],
                    url
                )

        results = await asyncio.gather(
            *[
                one(r)
                for _, r in df.iterrows()
            ]
        )

        for i, group in enumerate(
            results,
            start=1
        ):

            rows += group

            if (
                i % 30 == 0
                or i == len(results)
            ):
                print(
                    f'{i}/{len(results)} '
                    '기관 게시판 탐색 완료'
                )

    if rows:

        out = pd.DataFrame(
            rows
        ).drop_duplicates(
            ['기관명', '게시판URL']
        )

    else:

        out = pd.DataFrame(
            columns=[
                '기관명',
                '게시판명',
                '게시판URL',
                '점수'
            ]
        )

    out.to_excel(
        BOARDS_FILE,
        index=False
    )

    print(
        f'게시판 후보 {len(out)}개 저장: '
        f'{BOARDS_FILE}'
    )


# ============================================================
# 매일 게시판 모니터링
# ============================================================

async def monitor():

    if not os.path.exists(
        BOARDS_FILE
    ):

        print(
            'boards.xlsx가 없습니다.'
        )

        print(
            '최초 게시판 발견을 시작합니다.'
        )

        await discover_all()

    boards = pd.read_excel(
        BOARDS_FILE
    ).fillna('')

    state = state_load()

    found = []

    print(
        f'모니터링 게시판 수: {len(boards)}'
    )

    conn = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=False
    )

    async with aiohttp.ClientSession(
        connector=conn,
        headers={'User-Agent': UA}
    ) as s:

        sem = asyncio.Semaphore(
            CONCURRENCY
        )

        async def one(r):

            async with sem:

                h, _ = await fetch(
                    s,
                    r['게시판URL']
                )

                if not h:
                    return []

                res = []

                posts = parse_posts(
                    h,
                    r['게시판URL']
                )

                for p in posts:

                    x = await inspect(
                        s,
                        p
                    )

                    if x:

                        k = hashlib.sha256(
                            x['url'].encode()
                        ).hexdigest()

                        if k not in state:

                            x.update({
                                'key': k,
                                '기관명': r['기관명'],
                                '게시판명': r['게시판명']
                            })

                            res.append(x)

                return res

        results = await asyncio.gather(
            *[
                one(r)
                for _, r in boards.iterrows()
            ]
        )

        for group in results:
            found += group

    # ========================================================
    # Telegram 알림
    # ========================================================

    token = os.getenv(
        'TELEGRAM_BOT_TOKEN'
    )

    chat = os.getenv(
        'TELEGRAM_CHAT_ID'
    )

    for x in found:

        msg = (
            '🔔 신규 키워드 게시글\n\n'
            f"기관: {x['기관명']}\n"
            f"게시판: {x['게시판명']}\n"
            f"키워드: {', '.join(x['keywords'])}\n"
            f"제목: {x['title']}\n"
            f"등록일: {x['date'] or '확인불가'}\n"
            f"링크: {x['url']}"
        )

        if token and chat:

            try:

                rr = requests.post(
                    f'https://api.telegram.org/bot{token}/sendMessage',
                    data={
                        'chat_id': chat,
                        'text': msg,
                        'disable_web_page_preview': True
                    },
                    timeout=20
                )

                rr.raise_for_status()

            except Exception as e:

                print(
                    'Telegram 실패',
                    e
                )

                continue

        state[x['key']] = {
            '기관명': x['기관명'],
            'title': x['title'],
            'url': x['url'],
            'date': x['date'],
            'saved_at': datetime.now().isoformat()
        }

    state_save(state)

    print(
        '신규 알림 대상:',
        len(found)
    )


# ============================================================
# 프로그램 실행
# ============================================================

if __name__ == '__main__':

    asyncio.run(
        monitor()
    )
