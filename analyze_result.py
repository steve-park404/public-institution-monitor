import pandas as pd
from collections import Counter

# ============================================================
# 설정
# ============================================================

FILE = "missing_structure_analysis.xlsx"

# ============================================================
# 엑셀 읽기
# ============================================================

print("=" * 80)
print("미발견 기관 구조 분석 결과 확인")
print("=" * 80)

try:
    df = pd.read_excel(FILE)
except Exception as e:
    print(f"❌ 엑셀 파일을 읽을 수 없습니다: {e}")
    raise

print(f"\n전체 분석 기관 수 : {len(df)}개")


# ============================================================
# 컬럼 확인
# ============================================================

print("\n" + "=" * 80)
print("1. 엑셀 컬럼 확인")
print("=" * 80)

for i, col in enumerate(df.columns, 1):
    print(f"{i:2d}. {col}")


# ============================================================
# 공통 출력 함수
# ============================================================

def print_distribution(title, column):
    print("\n" + "-" * 80)
    print(title)
    print("-" * 80)

    if column not in df.columns:
        print(f"⚠️ '{column}' 컬럼이 없습니다.")
        return

    values = (
        df[column]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    counts = values.value_counts()

    if len(counts) == 0:
        print("데이터 없음")
        return

    for value, count in counts.items():
        if value == "":
            value = "(빈 값)"
        print(f"{value:<40} : {count:>4}개")


# ============================================================
# 2. 접속 상태
# ============================================================

print_distribution(
    "2. 접속 상태",
    "접속상태"
)


# ============================================================
# 3. CMS / 프레임워크
# ============================================================

print_distribution(
    "3. CMS / 프레임워크",
    "CMS/프레임워크"
)


# ============================================================
# 4. iframe 사용 여부
# ============================================================

print_distribution(
    "4. iframe 사용 여부",
    "iframe 사용"
)


# ============================================================
# 5. JavaScript 의존도
# ============================================================

print_distribution(
    "5. JavaScript 의존도",
    "JS 의존도"
)


# ============================================================
# 6. 공지사항 / 게시판 후보
# ============================================================

print_distribution(
    "6. 게시판 후보 발견 여부",
    "공지사항 후보"
)


# ============================================================
# 7. 게시물 상세 URL
# ============================================================

print_distribution(
    "7. 게시물 상세 URL 발견 여부",
    "게시물 상세 URL 발견"
)


# ============================================================
# 8. 게시판 유형
# ============================================================

print_distribution(
    "8. 게시판 유형",
    "게시판 유형"
)


# ============================================================
# 9. 추정 원인
# ============================================================

print_distribution(
    "9. 게시판 미발견 추정 원인",
    "추정 원인"
)


# ============================================================
# 10. 6차 공략법
# ============================================================

print_distribution(
    "10. 6차 탐색 공략법",
    "6차 공략법"
)


# ============================================================
# 11. 게시판 후보가 발견된 기관 목록
# ============================================================

print("\n" + "=" * 80)
print("11. 게시판 후보가 발견된 기관")
print("=" * 80)

if "공지사항 후보" in df.columns:

    mask = (
        df["공지사항 후보"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .isin([
            "yes",
            "y",
            "true",
            "1",
            "있음",
            "발견",
            "후보 있음"
        ])
    )

    result = df[mask]

    print(f"후보 발견 기관 : {len(result)}개\n")

    for _, row in result.iterrows():

        name = row.get("기관명", "")
        url = row.get("URL", "")
        cms = row.get("CMS/프레임워크", "")
        iframe = row.get("iframe 사용", "")
        js = row.get("JS 의존도", "")
        reason = row.get("추정 원인", "")
        method = row.get("6차 공략법", "")

        print("-" * 80)
        print(f"기관명   : {name}")
        print(f"URL      : {url}")
        print(f"CMS      : {cms}")
        print(f"iframe   : {iframe}")
        print(f"JS       : {js}")
        print(f"추정원인 : {reason}")
        print(f"6차공략  : {method}")


# ============================================================
# 12. 게시물 URL이 발견된 기관 목록
# ============================================================

print("\n" + "=" * 80)
print("12. 게시물 상세 URL이 발견된 기관")
print("=" * 80)

if "게시물 상세 URL 발견" in df.columns:

    values = (
        df["게시물 상세 URL 발견"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    mask = values.isin([
        "yes",
        "y",
        "true",
        "1",
        "있음",
        "발견"
    ])

    result = df[mask]

    print(f"게시물 URL 발견 기관 : {len(result)}개\n")

    for _, row in result.iterrows():

        print("-" * 80)
        print(f"기관명   : {row.get('기관명', '')}")
        print(f"URL      : {row.get('URL', '')}")
        print(f"게시판   : {row.get('공지사항 후보', '')}")
        print(f"게시물   : {row.get('게시물 상세 URL 발견', '')}")
        print(f"원인     : {row.get('추정 원인', '')}")
        print(f"6차공략  : {row.get('6차 공략법', '')}")


# ============================================================
# 13. 탐색 페이지 수 분석
# ============================================================

print("\n" + "=" * 80)
print("13. 탐색 페이지 수")
print("=" * 80)

if "탐색 페이지 수" in df.columns:

    pages = pd.to_numeric(
        df["탐색 페이지 수"],
        errors="coerce"
    ).fillna(0)

    print(f"최소 : {pages.min():.0f}")
    print(f"최대 : {pages.max():.0f}")
    print(f"평균 : {pages.mean():.1f}")


# ============================================================
# 14. 실제 발견 URL 예시
# ============================================================

print("\n" + "=" * 80)
print("14. 발견 URL 예시가 있는 기관")
print("=" * 80)

if "발견 URL 예시" in df.columns:

    values = (
        df["발견 URL 예시"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    result = df[values != ""]

    print(f"발견 URL이 있는 기관 : {len(result)}개\n")

    for _, row in result.iterrows():

        print("-" * 80)
        print(f"기관명 : {row.get('기관명', '')}")
        print(f"URL    : {row.get('URL', '')}")
        print(f"발견URL: {row.get('발견 URL 예시', '')}")


# ============================================================
# 15. 6차 탐색 우선순위 자동 분류
# ============================================================

print("\n" + "=" * 80)
print("15. 6차 탐색 우선순위 분석")
print("=" * 80)

priority = {
    "A. 바로 재검증": [],
    "B. iframe 전용 탐색": [],
    "C. K2Web/CMS 전용 탐색": [],
    "D. JS/API 전용 탐색": [],
    "E. 외부 게시판/별도 도메인 탐색": [],
    "F. 접속 불가": [],
    "G. 추가 분석 필요": []
}


for _, row in df.iterrows():

    name = str(row.get("기관명", ""))
    status = str(row.get("접속상태", "")).lower()
    cms = str(row.get("CMS/프레임워크", "")).lower()
    iframe = str(row.get("iframe 사용", "")).lower()
    js = str(row.get("JS 의존도", "")).lower()
    board = str(row.get("공지사항 후보", "")).lower()
    post = str(row.get("게시물 상세 URL 발견", "")).lower())
    reason = str(row.get("추정 원인", "")).lower()

    # --------------------------------------------------------
    # 접속 실패
    # --------------------------------------------------------

    if (
        "실패" in status
        or "fail" in status
        or "timeout" in status
        or "error" in status
        or "접속불가" in status
    ):
        priority["F. 접속 불가"].append(name)
        continue

    # --------------------------------------------------------
    # iframe
    # --------------------------------------------------------

    if (
        "true" in iframe
        or "yes" in iframe
        or "있음" in iframe
        or "사용" in iframe
    ):
        priority["B. iframe 전용 탐색"].append(name)
        continue

    # --------------------------------------------------------
    # K2Web / CMS
    # --------------------------------------------------------

    if (
        "k2web" in cms
        or "k2web" in reason
        or "cms" in cms
    ):
        priority["C. K2Web/CMS 전용 탐색"].append(name)
        continue

    # --------------------------------------------------------
    # JS
    # --------------------------------------------------------

    if (
        "높음" in js
        or "high" in js
        or "동적" in js
        or "api" in reason
        or "javascript" in reason
    ):
        priority["D. JS/API 전용 탐색"].append(name)
        continue

    # --------------------------------------------------------
    # 게시판/게시물 후보가 이미 있는 경우
    # --------------------------------------------------------

    if (
        "yes" in board
        or "true" in board
        or "있음" in board
        or "발견" in board
        or "yes" in post
        or "true" in post
        or "있음" in post
        or "발견" in post
    ):
        priority["A. 바로 재검증"].append(name)
        continue

    # --------------------------------------------------------
    # 외부 게시판
    # --------------------------------------------------------

    if (
        "외부" in reason
        or "external" in reason
        or "별도" in reason
        or "도메인" in reason
    ):
        priority["E. 외부 게시판/별도 도메인 탐색"].append(name)
        continue

    # --------------------------------------------------------
    # 나머지
    # --------------------------------------------------------

    priority["G. 추가 분석 필요"].append(name)


# ============================================================
# 우선순위 결과 출력
# ============================================================

for category, institutions in priority.items():

    print("\n" + "-" * 80)
    print(f"{category} : {len(institutions)}개")
    print("-" * 80)

    for name in institutions:
        print(f"- {name}")


# ============================================================
# 16. 최종 요약
# ============================================================

print("\n" + "=" * 80)
print("16. 최종 요약")
print("=" * 80)

print(f"전체 미발견 기관 : {len(df)}개")

for category, institutions in priority.items():
    print(f"{category:<35} : {len(institutions):>4}개")

print("\n" + "=" * 80)
print("분석 완료")
print("=" * 80)
