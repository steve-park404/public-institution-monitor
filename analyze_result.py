import pandas as pd

FILE = "missing_structure_analysis.xlsx"

print("=" * 80)
print("미발견 기관 6차 탐색용 상세 분석")
print("=" * 80)

df = pd.read_excel(FILE)

print(f"\n전체 기관 : {len(df)}개")


# ============================================================
# 1. 컬럼 확인
# ============================================================

print("\n" + "=" * 80)
print("1. 컬럼")
print("=" * 80)

for col in df.columns:
    print("-", col)


# ============================================================
# 공통 출력 함수
# ============================================================

def show_count(title, column):
    print("\n" + "-" * 80)
    print(title)
    print("-" * 80)

    if column not in df.columns:
        print("컬럼 없음")
        return

    s = df[column].fillna("").astype(str).str.strip()

    for value, count in s.value_counts().items():
        if value == "":
            value = "(빈 값)"
        print(f"{value} : {count}개")


# ============================================================
# 2~10. 기본 분류
# ============================================================

show_count("2. 접속상태", "접속상태")

show_count("3. CMS/프레임워크", "CMS/프레임워크")

show_count("4. iframe 사용", "iframe 사용")

show_count("5. JS 의존도", "JS 의존도")

show_count("6. 공지사항 후보", "공지사항 후보")

show_count("7. 게시물 상세 URL 발견", "게시물 상세 URL 발견")

show_count("8. 게시판 유형", "게시판 유형")

show_count("9. 추정 원인", "추정 원인")

show_count("10. 6차 공략법", "6차 공략법")


# ============================================================
# 11. 게시판 후보 또는 게시물 URL이 발견된 기관
# ============================================================

print("\n" + "=" * 80)
print("11. 게시판 후보 또는 게시물 URL이 발견된 기관")
print("=" * 80)

for _, r in df.iterrows():

    board = str(r.get("공지사항 후보", "")).lower()
    post = str(r.get("게시물 상세 URL 발견", "")).lower()

    if (
        "yes" in board
        or "true" in board
        or "있음" in board
        or "발견" in board
        or
        "yes" in post
        or "true" in post
        or "있음" in post
        or "발견" in post
    ):

        print("\n----------------------------------------")
        print("기관명 :", r.get("기관명", ""))
        print("URL    :", r.get("URL", ""))
        print("CMS    :", r.get("CMS/프레임워크", ""))
        print("iframe :", r.get("iframe 사용", ""))
        print("JS     :", r.get("JS 의존도", ""))
        print("게시판 :", r.get("공지사항 후보", ""))
        print("게시물 :", r.get("게시물 상세 URL 발견", ""))
        print("유형   :", r.get("게시판 유형", ""))
        print("원인   :", r.get("추정 원인", ""))
        print("공략법 :", r.get("6차 공략법", ""))
        print("URL예시:", r.get("발견 URL 예시", ""))


# ============================================================
# 12. 6차 탐색 우선순위
# ============================================================

print("\n" + "=" * 80)
print("12. 6차 탐색 우선순위별 기관")
print("=" * 80)

groups = {
    "A. 바로 재검증": [],
    "B. iframe 전용": [],
    "C. K2Web/CMS 전용": [],
    "D. JS/API 전용": [],
    "E. 외부 게시판 전용": [],
    "F. 접속 실패": [],
    "G. 기타/추가 분석": []
}


for _, r in df.iterrows():

    name = str(r.get("기관명", ""))
    status = str(r.get("접속상태", "")).lower()
    cms = str(r.get("CMS/프레임워크", "")).lower()
    iframe = str(r.get("iframe 사용", "")).lower()
    js = str(r.get("JS 의존도", "")).lower()
    reason = str(r.get("추정 원인", "")).lower()
    board = str(r.get("공지사항 후보", "")).lower()
    post = str(r.get("게시물 상세 URL 발견", "")).lower()

    # --------------------------------------------------------
    # 접속 실패
    # --------------------------------------------------------

    if (
        "실패" in status
        or "timeout" in status
        or "error" in status
        or "접속불가" in status
    ):
        groups["F. 접속 실패"].append(name)

    # --------------------------------------------------------
    # iframe
    # --------------------------------------------------------

    elif (
        "true" in iframe
        or "yes" in iframe
        or "있음" in iframe
        or "사용" in iframe
    ):
        groups["B. iframe 전용"].append(name)

    # --------------------------------------------------------
    # K2Web / CMS
    # --------------------------------------------------------

    elif (
        "k2web" in cms
        or "k2web" in reason
        or "cms" in cms
    ):
        groups["C. K2Web/CMS 전용"].append(name)

    # --------------------------------------------------------
    # JS / API
    # --------------------------------------------------------

    elif (
        "높음" in js
        or "high" in js
        or "동적" in js
        or "javascript" in reason
        or "api" in reason
    ):
        groups["D. JS/API 전용"].append(name)

    # --------------------------------------------------------
    # 외부 게시판
    # --------------------------------------------------------

    elif (
        "외부" in reason
        or "external" in reason
        or "별도" in reason
    ):
        groups["E. 외부 게시판 전용"].append(name)

    # --------------------------------------------------------
    # 게시판/게시물 후보가 이미 있는 경우
    # --------------------------------------------------------

    elif (
        "yes" in board
        or "true" in board
        or "있음" in board
        or "발견" in board
        or "yes" in post
        or "true" in post
        or "있음" in post
        or "발견" in post
    ):
        groups["A. 바로 재검증"].append(name)

    # --------------------------------------------------------
    # 나머지
    # --------------------------------------------------------

    else:
        groups["G. 기타/추가 분석"].append(name)


# ============================================================
# 13. 우선순위 결과 출력
# ============================================================

for group, names in groups.items():

    print("\n" + "-" * 80)
    print(f"{group} : {len(names)}개")
    print("-" * 80)

    for name in names:
        print(name)


# ============================================================
# 14. 정상 접속했지만 게시판/게시물 URL을 찾지 못한 기관
# ============================================================

print("\n" + "=" * 80)
print("14. 정상 접속했지만 게시판/게시물 URL을 찾지 못한 기관")
print("=" * 80)

count = 0

for _, r in df.iterrows():

    status = str(r.get("접속상태", "")).lower()
    board = str(r.get("공지사항 후보", "")).lower()
    post = str(r.get("게시물 상세 URL 발견", "")).lower()

    failed = (
        "실패" in status
        or "timeout" in status
        or "error" in status
        or "접속불가" in status
    )

    found_board = (
        "yes" in board
        or "true" in board
        or "있음" in board
        or "발견" in board
    )

    found_post = (
        "yes" in post
        or "true" in post
        or "있음" in post
        or "발견" in post
    )

    if not failed and not found_board and not found_post:

        count += 1

        print(
            f"{count:3d}. "
            f"{r.get('기관명', '')} | "
            f"{r.get('URL', '')} | "
            f"CMS={r.get('CMS/프레임워크', '')} | "
            f"iframe={r.get('iframe 사용', '')} | "
            f"JS={r.get('JS 의존도', '')} | "
            f"원인={r.get('추정 원인', '')}"
        )


# ============================================================
# 15. 최종 요약
# ============================================================

print("\n" + "=" * 80)
print("15. 최종 요약")
print("=" * 80)

print(f"전체 미발견 기관 : {len(df)}개")

for group, names in groups.items():
    print(f"{group:<35} : {len(names):>4}개")

print("\n" + "=" * 80)
print("분석 완료")
print("=" * 80)
