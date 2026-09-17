import pandas as pd
p = "monitor_targets.xlsx"
d = pd.read_excel(p)
print("전체 대상:", len(d))
print("기관 수:", d["기관명"].nunique())
print("\n우선순위별")
print(d["우선순위"].value_counts().sort_index())
print("\n게시판명에 공지사항 포함:")
print(d[d["게시판명"].fillna("").astype(str).str.contains("공지사항")][["기관명","게시판명","게시판URL"]].to_string(index=False))
