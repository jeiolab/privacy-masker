import app

A = app.App

# _split_headline
cases = [
    "✅ 남은 개인정보 없음  (저장: output\\AI용\\x.md)",
    "⚠ 남은 개인정보 2건 발견 - 결과를 AI에 붙여 넣지 마세요 (저장: output\\보관용\\y.md)",
    "⚠ 스캔된 쪽: 3쪽 - 처리 불가",
]
print("[_split_headline]")
for c in cases:
    print("  ", A._split_headline(c))

# _hide_md_separators
md = "| 이름 | 점수 |\n| --- | --- |\n| 학생003 | 90 |\n|:---:|:--:|\n본문 --- 유지\n| 홍길동 | 80 |"
print("[_hide_md_separators] 결과(미리보기용):")
print(A._hide_md_separators(md))
print("원본 줄수:", len(md.split(chr(10))), "-> 미리보기 줄수:", len(A._hide_md_separators(md).split(chr(10))))
