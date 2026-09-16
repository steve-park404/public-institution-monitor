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
# 설정
# ============================================================

KEYWORDS = [
    '설문조사',
    '시민참여',
    '국민참여',
    '공모전'
]

BOARD_WORDS = [
    '공지사항',
    '공지',
    '알림마당',
    '알림',
    '새소식',
    '소식',
    '게시판',
    '국민참여',
    '시민참여',
    '참여',
    '공모전',
    '공모',
    '설문조사',
    '설문',
    '고시공고',
    '입찰',
    '뉴스',
    '보도자료',
    '자료실'
]

# 게시판으로 판단하기 좋은 URL 단어
BOARD_URL_WORDS = [
    'bbs',
    'board',
    'notice',
    'news',
    'list',
    'community',
    'particip',
    'event',
    'survey',
    'data',
    'pds'
]

# 게시글 상세페이지로 판단되는 패턴
DETAIL_PATTERNS = [
    r'view',
    r'detail',
    r'read',
    r'article',
    r'articleNo',
    r'nttId',
    r'nttid',
    r'bbsno',
    r'boardno',
    r'seq=',
    r'idx=',
    r'no=',
    r'wr_id',
    r'uid=',
    r'postid',
    r'newsid'
]

URL_FILE = 'url_완성.xlsx'
BOARDS_FILE = 'boards.xlsx'
MISSING_FILE = 'boards_missing.xlsx'
STATE_FILE = 'seen_posts.json'

CONCURRENCY = 20
TIMEOUT = 20

# 홈페이지 → 메뉴 → 게시판까지 탐색
MAX_DEPTH = 2

# 기관당 저장할 게시판 후보 수
MAX_BOARDS_PER_ORG = 8

UA = (
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
    'AppleWebKit/537.36 (KHTML, like Gecko) '
    'Chrome/140 Safari/537.36'
)


# ============================================================
# 기본 함수
# ============================================================

def norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def clean_url(x):
    if pd.isna(x):
        return ''

    x = str(x).strip()

    m = re.search(r'\]\((https?://[^)]+)\)', x)

    if m:
        x = m.group(1)

    x = x.strip('[]() ')

    if not x:
        return ''

    if re.match(r'^https?://', x, re.I):
        return x

    return 'https://' + x


def canon(u):
    return urldefrag(u or '')[0].rstrip('/')


def same_domain(a, b):
    return (
        urlparse(a).netloc.lower().lstrip('www.')
        == urlparse(b).netloc.lower().lstrip('www.')
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
    text = norm(s).lower()

    return [
        k for k in KEYWORDS
        if k.lower() in text
    ]


# ============================================================
# URL이 게시글 상세페이지인지 판단
# ============================================================

def is_detail_url(url):
    low = url.lower()

    for pattern in DETAIL_PATTERNS:
        if re.search(pattern, low, re.I):
            return True

    return False


# ============================================================
# URL이 게시판 후보인지 판단
# ============================================================

def url_board_score(url):
    low = url.lower()

    score = 0

    for word in BOARD_URL_WORDS:
        if word in low:
            score += 2

    if is_detail_url(url):
        score -= 7

    return score


# ============================================================
# 링크 점수
# ============================================================

def score_link(text, href):

    text = norm(text)
    combined = (text + ' ' + href).lower()

    score = 0

    # 메뉴 이름
    for word in BOARD_WORDS:

        if word.lower() in text.lower():
            score += 5

        elif word.lower() in combined:
            score += 2

    # URL 패턴
    score += url_board_score(href)

    # 목록/게시판
    if re.search(
        r'(list|board|bbs|notice|news|community)',
        href,
        re.I
    ):
        score += 3

    # 상세페이지 감점
    if is_detail_url(href):
        score -= 8

    # 너무 긴 링크 문구는 게시글 제목일 가능성이 높음
    if len(text) > 70:
        score -= 5

    # 날짜가 포함된 링크는 게시글 제목일 가능성이 높음
    if date_from(text):
        score -= 4

    # 일반적인 게시글 제목 형태
    if re.search(
        r'(모집|공고|안내|알려드립니다|참여자|접수|신청)',
        text
    ) and len(text) > 25:
        score -= 3

    return score


# ============================================================
# 링크 추출
# ============================================================

def extract_links(html, base_url):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    result = []

    for a in soup.find_all(
        'a',
        href=True
    ):

        href = a.get('href', '').strip()
        text = norm(
            a.get_text(' ', strip=True)
        )

        if not href:
            continue

        if href.lower().startswith(
            ('javascript:', 'mailto:', '#', 'tel:')
        ):
            continue

        url = canon(
            urljoin(base_url, href)
        )

        if not same_domain(
            url,
            base_url
        ):
            continue

        if url == canon(base_url):
            continue

        result.append({
            'url': url,
            'text': text
        })

    return result


# ============================================================
# 게시판 후보 검증
# ============================================================

def validate_board_html(html, url):

    """
    실제 게시판 페이지에 가까운지 HTML 구조를 확인한다.
    """

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    text = norm(
        soup.get_text(
            ' ',
            strip=True
        )
    )

    score = 0

    # 게시판에서 흔히 나타나는 단어
    for word in [
        '번호',
        '제목',
        '등록일',
        '조회수',
        '작성일',
        '공지',
        '목록'
    ]:

        if word in text:
            score += 1

    # 여러 개의 링크가 있는지
    links = soup.find_all(
        'a',
        href=True
    )

    if len(links) >= 10:
        score += 2

    # 표 형태
    if soup.find('table'):
        score += 3

    # 게시글 상세 링크가 여러 개 있는지
    detail_count = 0

    for a in links:

        href = a.get(
            'href',
            ''
        )

        if is_detail_url(href):
            detail_count += 1

    if detail_count >= 3:
        score += 3

    return score


# ============================================================
# 기관별 게시판 탐색
# ============================================================

async def discover(session, name, home):

    first_html, final_home = await fetch(
        session,
        home
    )

    if not first_html:
        return []

    candidates = {}

    # --------------------------------------------------------
    # BFS 방식으로 홈페이지 → 메뉴 → 게시판 탐색
    # --------------------------------------------------------

    queue = [
        (
            canon(final_home),
            first_html,
            0
        )
    ]

    visited = set()

    while queue:

        current_url, html, depth = queue.pop(0)

        current_url = canon(
            current_url
        )

        if current_url in visited:
            continue

        visited.add(
            current_url
        )

        links = extract_links(
            html,
            current_url
        )

        # ----------------------------------------------------
        # 현재 페이지의 링크 분석
        # ----------------------------------------------------

        for item in links:

            url = item['url']
            text = item['text']

            score = score_link(
                text,
                url
            )

            # 후보 점수가 충분히 높을 경우
            if score >= 5:

                candidates[url] = max(
                    candidates.get(
                        url,
                        {
                            'score': -999,
                            'name': ''
                        }
                    )['score'],
                    score
                ), norm(text)[:100]

        # ----------------------------------------------------
        # 2단계 탐색
        # ----------------------------------------------------

        if depth >= MAX_DEPTH:
            continue

        # 메뉴/게시판으로 보이는 링크만 다음 단계 탐색
        next_links = []

        for item in links:

            url = item['url']
            text = item['text']

            score = score_link(
                text,
                url
            )

            if score >= 2 and not is_detail_url(url):

                next_links.append(
                    (score, url)
                )

        # 점수 높은 링크부터 최대 10개 탐색
        next_links.sort(
            reverse=True
        )

        for _, next_url in next_links[:10]:

            if next_url in visited:
                continue

            next_html, next_final = await fetch(
                session,
                next_url
            )

            if not next_html:
                continue

            # 실제 게시판 구조인지 확인
            board_quality = validate_board_html(
                next_html,
                next_final
            )

            # 게시판 구조가 명확하면 점수 추가
            if board_quality >= 4:

                old = candidates.get(
                    canon(next_final),
                    {
                        'score': 0,
                        'name': ''
                    }
                )

                candidates[
                    canon(next_final)
                ] = (
                    old['score'] + board_quality + 3,
                    old['name']
                    or norm(next_url)
                )

            # 다음 단계 탐색
            if depth + 1 < MAX_DEPTH:

                queue.append(
                    (
                        canon(next_final),
                        next_html,
                        depth + 1
                    )
                )

    # --------------------------------------------------------
    # 최종 후보 정리
    # --------------------------------------------------------

    result = []

    for url, value in candidates.items():

        score, board_name = value

        # 너무 낮은 후보 제외
        if score < 5:
            continue

        # 상세페이지는 제외
        if is_detail_url(url):
            continue

        result.append({
            '기관명': name,
            '게시판명': board_name,
            '게시판URL': url,
            '점수': score
        })

    result.sort(
        key=lambda x: x['점수'],
        reverse=True
    )

    return result[
        :MAX_BOARDS_PER_ORG
    ]


# ============================================================
# 웹페이지 요청
# ============================================================

async def fetch(session, url):

    try:

        async with session.get(
            url,
            timeout=aiohttp.ClientTimeout(
                total=TIMEOUT
            ),
            allow_redirects=True,
            ssl=False
        ) as response:

            if response.status >= 400:
                return None, str(
                    response.url
                )

            raw = await response.read()

            return (
                raw.decode(
                    response.charset or 'utf-8',
                    errors='ignore'
                ),
                str(response.url)
            )

    except Exception:

        return None, url


# ============================================================
# 전체 기관 게시판 수집
# ============================================================

async def discover_all():

    if not os.path.exists(
        URL_FILE
    ):

        raise FileNotFoundError(
            f'{URL_FILE} 파일이 없습니다.'
        )

    df = pd.read_excel(
        URL_FILE
    )

    results = []

    missing = []

    conn = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=False
    )

    async with aiohttp.ClientSession(
        connector=conn,
        headers={
            'User-Agent': UA
        }
    ) as session:

        sem = asyncio.Semaphore(
            CONCURRENCY
        )

        async def one(index, row):

            async with sem:

                name = norm(
                    row.get(
                        '기관명',
                        ''
                    )
                )

                url = clean_url(
                    row.get(
                        'URL',
                        ''
                    )
                )

                if not url:

                    return (
                        index,
                        name,
                        [],
                        'URL 없음'
                    )

                try:

                    found = await discover(
                        session,
                        name,
                        url
                    )

                    if not found:

                        return (
                            index,
                            name,
                            [],
                            '게시판 후보 없음'
                        )

                    return (
                        index,
                        name,
                        found,
                        ''
                    )

                except Exception as e:

                    return (
                        index,
                        name,
                        [],
                        str(e)
                    )

        tasks = [
            one(
                i,
                row
            )
            for i, (_, row)
            in enumerate(
                df.iterrows(),
                start=1
            )
        ]

        for i, task in enumerate(
            asyncio.as_completed(tasks),
            start=1
        ):

            try:

                index, name, found, error = (
                    await task
                )

                if found:

                    results.extend(
                        found
                    )

                else:

                    missing.append({
                        '기관명': name,
                        '사유': error
                    })

            except Exception as e:

                print(
                    f'처리 오류: {e}'
                )

            if (
                i % 25 == 0
                or i == len(tasks)
            ):

                print(
                    f'기관 {i}/{len(tasks)} '
                    f'확인 완료 '
                    f'(게시판 후보 {len(results)}개)'
                )

    # --------------------------------------------------------
    # boards.xlsx
    # --------------------------------------------------------

    if results:

        boards = pd.DataFrame(
            results
        )

        boards = boards.drop_duplicates(
            ['기관명', '게시판URL']
        )

        boards = boards.sort_values(
            ['기관명', '점수'],
            ascending=[True, False]
        )

    else:

        boards = pd.DataFrame(
            columns=[
                '기관명',
                '게시판명',
                '게시판URL',
                '점수'
            ]
        )

    boards.to_excel(
        BOARDS_FILE,
        index=False
    )

    # --------------------------------------------------------
    # boards_missing.xlsx
    # --------------------------------------------------------

    missing_df = pd.DataFrame(
        missing
    )

    if not missing_df.empty:

        missing_df = missing_df.drop_duplicates(
            ['기관명']
        )

    missing_df.to_excel(
        MISSING_FILE,
        index=False
    )

    # --------------------------------------------------------
    # 결과 출력
    # --------------------------------------------------------

    total_orgs = len(df)

    found_orgs = (
        boards['기관명'].nunique()
        if not boards.empty
        else 0
    )

    missing_orgs = (
        missing_df['기관명'].nunique()
        if not missing_df.empty
        else 0
    )

    print()
    print('=' * 60)
    print('게시판 탐색 완료')
    print('=' * 60)

    print(
        f'전체 기관: {total_orgs}'
    )

    print(
        f'게시판 발견 기관: {found_orgs}'
    )

    print(
        f'게시판 미발견 기관: {missing_orgs}'
    )

    print(
        f'게시판 후보: {len(boards)}개'
    )

    print(
        f'저장 파일: {BOARDS_FILE}'
    )

    print(
        f'미발견 기관: {MISSING_FILE}'
    )

    print('=' * 60)

    print()
    print(
        '첫 실행이므로 게시글 모니터링은 하지 않습니다.'
    )


# ============================================================
# 게시글 목록 추출
# ============================================================

def parse_posts(html, board):

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    output = {}

    # 일반적인 테이블형 게시판
    for tr in soup.find_all('tr'):

        row_text = norm(
            tr.get_text(
                ' ',
                strip=True
            )
        )

        date = date_from(
            row_text
        )

        for a in tr.find_all(
            'a',
            href=True
        ):

            title = norm(
                a.get_text(
                    ' ',
                    strip=True
                )
            )

            href = a['href']

            if len(title) < 2:
                continue

            if href.lower().startswith(
                (
                    'javascript:',
                    'mailto:',
                    '#'
                )
            ):
                continue

            url = canon(
                urljoin(
                    board,
                    href
                )
            )

            if not same_domain(
                url,
                board
            ):
                continue

            if is_detail_url(url):

                output[url] = {
                    'title': title,
                    'url': url,
                    'date': date
                }

    # 비테이블형 게시판
    if len(output) < 3:

        for a in soup.find_all(
            'a',
            href=True
        ):

            title = norm(
                a.get_text(
                    ' ',
                    strip=True
                )
            )

            href = a['href']

            if len(title) < 4:
                continue

            if href.lower().startswith(
                (
                    'javascript:',
                    'mailto:',
                    '#'
                )
            ):
                continue

            url = canon(
                urljoin(
                    board,
                    href
                )
            )

            if not same_domain(
                url,
                board
            ):
                continue

            if is_detail_url(url):

                output[url] = {
                    'title': title,
                    'url': url,
                    'date': ''
                }

    return list(
        output.values()
    )[:20]


# ============================================================
# 게시글 상세 검사
# ============================================================

async def inspect(session, post):

    html, final = await fetch(
        session,
        post['url']
    )

    if not html:
        return None

    soup = BeautifulSoup(
        html,
        'html.parser'
    )

    for element in soup([
        'script',
        'style',
        'noscript'
    ]):

        element.decompose()

    body = norm(
        soup.get_text(
            ' ',
            strip=True
        )
    )

    keywords = hits(
        post['title'] + ' ' + body
    )

    if not keywords:
        return None

    return {
        'url': canon(final),
        'title': post['title'],
        'date': (
            post['date']
            or date_from(
                body[:5000]
            )
        ),
        'keywords': keywords
    }


# ============================================================
# 상태 파일
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


def state_save(state):

    with open(
        STATE_FILE,
        'w',
        encoding='utf-8'
    ) as f:

        json.dump(
            state,
            f,
            ensure_ascii=False,
            indent=2
        )


# ============================================================
# 모니터링
# ============================================================

async def monitor():

    # --------------------------------------------------------
    # boards.xlsx가 없으면
    # 게시판 수집만 하고 종료
    # --------------------------------------------------------

    if not os.path.exists(
        BOARDS_FILE
    ):

        print(
            'boards.xlsx가 없습니다.'
        )

        print(
            '개선된 게시판 탐색을 시작합니다.'
        )

        await discover_all()

        return

    # --------------------------------------------------------
    # boards.xlsx가 있으면
    # 실제 모니터링
    # --------------------------------------------------------

    print(
        'boards.xlsx 확인됨.'
    )

    boards = pd.read_excel(
        BOARDS_FILE
    ).fillna('')

    state = state_load()

    found = []

    conn = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=False
    )

    async with aiohttp.ClientSession(
        connector=conn,
        headers={
            'User-Agent': UA
        }
    ) as session:

        sem = asyncio.Semaphore(
            CONCURRENCY
        )

        async def one(row):

            async with sem:

                board_url = clean_url(
                    row['게시판URL']
                )

                if not board_url:
                    return []

                html, _ = await fetch(
                    session,
                    board_url
                )

                if not html:
                    return []

                posts = parse_posts(
                    html,
                    board_url
                )

                results = []

                for post in posts:

                    result = await inspect(
                        session,
                        post
                    )

                    if not result:
                        continue

                    key = hashlib.sha256(
                        result['url'].encode()
                    ).hexdigest()

                    if key in state:
                        continue

                    result.update({
                        'key': key,
                        '기관명': row['기관명'],
                        '게시판명': row['게시판명']
                    })

                    results.append(
                        result
                    )

                return results

        tasks = [
            one(row)
            for _, row
            in boards.iterrows()
        ]

        for i, task in enumerate(
            asyncio.as_completed(tasks),
            start=1
        ):

            try:

                result = await task

                found.extend(
                    result
                )

            except Exception as e:

                print(
                    f'게시판 모니터링 오류: {e}'
                )

            if (
                i % 50 == 0
                or i == len(tasks)
            ):

                print(
                    f'게시판 {i}/{len(tasks)} '
                    f'확인 완료'
                )

    # --------------------------------------------------------
    # Telegram
    # --------------------------------------------------------

    token = os.getenv(
        'TELEGRAM_BOT_TOKEN'
    )

    chat = os.getenv(
        'TELEGRAM_CHAT_ID'
    )

    for item in found:

        message = (
            f"🔔 신규 키워드 게시글\n\n"
            f"기관: {item['기관명']}\n"
            f"게시판: {item['게시판명']}\n"
            f"키워드: {', '.join(item['keywords'])}\n"
            f"제목: {item['title']}\n"
            f"등록일: {item['date'] or '확인불가'}\n"
            f"링크: {item['url']}"
        )

        if token and chat:

            try:

                response = requests.post(
                    f'https://api.telegram.org/bot{token}/sendMessage',
                    data={
                        'chat_id': chat,
                        'text': message,
                        'disable_web_page_preview': True
                    },
                    timeout=20
                )

                response.raise_for_status()

                print(
                    f'Telegram 발송 완료: '
                    f'{item["title"]}'
                )

            except Exception as e:

                print(
                    f'Telegram 발송 실패: {e}'
                )

                continue

        state[item['key']] = {
            '기관명': item['기관명'],
            'title': item['title'],
            'url': item['url'],
            'date': item['date'],
            'saved_at': datetime.now().isoformat()
        }

    state_save(state)

    print()
    print(
        f'신규 알림 대상: {len(found)}개'
    )


# ============================================================
# 실행
# ============================================================

if __name__ == '__main__':
    asyncio.run(
        monitor()
    )
