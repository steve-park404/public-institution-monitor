# -*- coding: utf-8 -*-
"""
V8.23.1 유사기관 URL 병합
- 기존 monitor.py 검색 로직은 변경하지 않음
- 사용자가 수동 입력한 게시판URL만 추가
- read_only worksheet의 max_row=None 문제를 방지하기 위해 iter_rows 사용
"""

from pathlib import Path
import re
import openpyxl

BASE_CANDIDATES = ["monitor_targets.xlsx", "url.xlsx", "targets.xlsx"]
EXTRA_CANDIDATES = [
    "유사기관_추가모니터링_후보목록_URL입력완료 (2).xlsx",
    "유사기관_추가모니터링_후보목록_URL입력완료.xlsx",
    "유사기관_추가모니터링_후보목록.xlsx",
]

def norm(x):
    return re.sub(r"\s+", "", str(x or "")).strip()

def valid_url(x):
    return bool(re.match(r"^https?://", str(x or "").strip(), re.I))

def find_file(candidates):
    for name in candidates:
        p = Path(name)
        if p.exists():
            return p
    return None

def get_headers(ws):
    return {
        norm(cell.value): idx
        for idx, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1)), start=1)
    }

def main():
    base = find_file(BASE_CANDIDATES)
    extra = find_file(EXTRA_CANDIDATES)

    if not base:
        raise SystemExit(
            "기존 기관목록 Excel을 찾지 못했습니다. "
            f"필요 파일: {BASE_CANDIDATES}"
        )
    if not extra:
        raise SystemExit("유사기관 추가 Excel을 찾지 못했습니다.")

    wb = openpyxl.load_workbook(base)
    ws = wb["355기관"] if "355기관" in wb.sheetnames else wb[wb.sheetnames[0]]
    hm = get_headers(ws)

    if "기관명" not in hm or "홈페이지URL" not in hm:
        raise SystemExit(
            f"기존 Excel의 기관명/홈페이지URL 열을 찾지 못했습니다: {list(hm)}"
        )

    for colname in ["공지게시판URL", "기관유형"]:
        if colname not in hm:
            idx = ws.max_column + 1
            ws.cell(1, idx).value = colname
            hm[colname] = idx

    existing = {}
    for r in range(2, ws.max_row + 1):
        name = norm(ws.cell(r, hm["기관명"]).value)
        if name:
            existing[name] = r

    # 추가 파일은 read_only로 열되 max_row를 사용하지 않고 행 자체를 순회
    ewb = openpyxl.load_workbook(extra, read_only=True, data_only=True)
    ews = ewb["추가후보"] if "추가후보" in ewb.sheetnames else ewb[ewb.sheetnames[0]]
    rows = ews.iter_rows(values_only=True)

    try:
        header_row = next(rows)
    except StopIteration:
        raise SystemExit("추가 Excel의 '추가후보' 시트가 비어 있습니다.")

    eh = {
        norm(value): idx
        for idx, value in enumerate(header_row)
    }

    if "기관명" not in eh or "게시판URL" not in eh:
        raise SystemExit(
            f"추가 Excel의 기관명/게시판URL 열을 찾지 못했습니다: {list(eh)}"
        )

    added = 0
    supplemented = 0
    skipped = 0
    valid_board_urls = 0

    for values in rows:
        def val(col):
            idx = eh.get(col)
            return values[idx] if idx is not None and idx < len(values) else ""

        name = norm(val("기관명"))
        board = str(val("게시판URL") or "").strip()
        home = str(val("홈페이지URL") or "").strip()
        org_type = str(val("기관유형") or "").strip()

        if not name or not valid_url(board):
            continue

        valid_board_urls += 1

        if name in existing:
            rr = existing[name]
            old_board = str(ws.cell(rr, hm["공지게시판URL"]).value or "").strip()
            changed = False

            if not valid_url(old_board):
                ws.cell(rr, hm["공지게시판URL"]).value = board
                changed = True

            if not str(ws.cell(rr, hm["홈페이지URL"]).value or "").strip() and home:
                ws.cell(rr, hm["홈페이지URL"]).value = home
                changed = True

            if changed:
                supplemented += 1
            else:
                skipped += 1
            continue

        rr = ws.max_row + 1
        ws.cell(rr, hm["기관명"]).value = name
        ws.cell(rr, hm["홈페이지URL"]).value = home or board
        ws.cell(rr, hm["공지게시판URL"]).value = board
        ws.cell(rr, hm["기관유형"]).value = org_type

        existing[name] = rr
        added += 1

    ewb.close()
    wb.save(base)

    print("=== V8.23.1 유사기관 URL 병합 완료 ===")
    print(f"기존 대상 파일: {base}")
    print(f"수동 게시판 URL 확인 건수: {valid_board_urls}")
    print(f"신규 기관 추가: {added}")
    print(f"기존 기관 URL 보완: {supplemented}")
    print(f"기존값 유지: {skipped}")
    print(f"최종 기관 수: {ws.max_row - 1}")

if __name__ == "__main__":
    main()
