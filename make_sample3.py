"""samples/가짜_명렬표.xlsx 의 학생 정보로 테스트용 가짜 PDF(samples/가짜_명단.pdf)를 만든다.

- 1쪽: 명단 표 (학번, 성명, 보호자명, 연락처) 15명
- 2쪽: 나머지 15명
- 3쪽: 상담 기록 문단 5개 (성명 전체, 성 뺀 이름+조사, 보호자명, 전화번호가 섞임)
- 문서 정보(제목, 작성자)에 학생 이름 하나와 "홍교사"

한글 글꼴은 PyMuPDF 내장 한국어 글꼴("korea")을 쓴다.
make_sample.py 를 먼저 실행해 명렬표가 있어야 한다.   실행: python3 make_sample3.py
"""

from pathlib import Path

import pymupdf

from make_sample2 import given_name, load_roster, phone_fmt

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "samples" / "가짜_명단.pdf"

FONT = "korea"            # PyMuPDF 내장 한국어 글꼴
PAGE_W, PAGE_H = pymupdf.paper_size("a4")
MARGIN = 50
COLUMNS = [("학번", 90), ("성명", 120), ("보호자명", 120), ("연락처", 165)]
ROW_H = 34
TABLE_TOP = 110
PER_PAGE = 15


def text_box(page, rect, text, size=11, align=0):
    """칸 안에 글을 넣는다. 들어가지 않으면 오류로 알린다."""
    rc = page.insert_textbox(rect, text, fontname=FONT, fontsize=size, align=align)
    if rc < 0:
        raise RuntimeError(f"글이 칸에 들어가지 않습니다: {text[:20]}...")


def add_table_page(doc, students, title, page_no, total):
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    text_box(page, pymupdf.Rect(MARGIN, 50, PAGE_W - MARGIN, 85), title, size=18)

    x_edges = [MARGIN]
    for _, w in COLUMNS:
        x_edges.append(x_edges[-1] + w)
    rows = [[c for c, _ in COLUMNS]] + [
        [s["학번"], s["성명"], s["보호자명"], s["학생연락처"]] for s in students
    ]

    for r, row in enumerate(rows):
        y0 = TABLE_TOP + r * ROW_H
        if r == 0:
            page.draw_rect(pymupdf.Rect(x_edges[0], y0, x_edges[-1], y0 + ROW_H), fill=(0.9, 0.93, 1))
        for c, value in enumerate(row):
            cell = pymupdf.Rect(x_edges[c], y0, x_edges[c + 1], y0 + ROW_H)
            page.draw_rect(cell, width=0.6)
            # 세로 가운데 맞춤을 위해 위쪽 여백을 조금 둔다
            text_box(page, pymupdf.Rect(cell.x0, y0 + 9, cell.x1, y0 + ROW_H), str(value), align=1)

    text_box(page, pymupdf.Rect(MARGIN, PAGE_H - 50, PAGE_W - MARGIN, PAGE_H - 30),
             f"- {page_no} / {total} -", size=9, align=1)


def counsel_sentences(students):
    """상담 기록 5문장. 성명, 성 뺀 이름+조사, 보호자명, 전화번호가 섞인다."""
    a, b, c, d, e = (students[i] for i in (0, 4, 10, 17, 23))
    ga, gb, gc, gd, ge = (given_name(s["성명"]) for s in (a, b, c, d, e))
    return [
        f"{a['성명']} 학생이 수업 시간에 자주 졸아 면담함. {ga}이는 요즘 늦게 잔다고 했고, "
        f"보호자 {a['보호자명']}에게 {phone_fmt(a['보호자연락처'], 'hyphen')} 로 연락하기로 함.",
        f"{b['성명']}은 모둠 활동에서 갈등이 있었음. {gb}이가 먼저 사과했으며 "
        f"{b['보호자명']} 보호자가 {phone_fmt(b['보호자연락처'], 'plain')} 로 전화하여 감사 인사를 전함.",
        f"{c['성명']} 학생 보호자 {c['보호자명']} 님과 면담. {gc}아, 다음 주까지 과제를 내기로 약속함. "
        f"학생 연락처 {phone_fmt(c['학생연락처'], 'dot')} 확인.",
        f"{d['성명']}의 체험학습 신청서가 제출됨. {gd}이를 칭찬해 주기로 하고, "
        f"{d['보호자명']} 보호자 번호 {phone_fmt(d['보호자연락처'], 'space')} 로 결과를 알림.",
        f"{e['성명']} 학생 상담 5회차. {ge}은 방과후 수업을 희망하며, 보호자 {e['보호자명']} 님은 "
        f"{phone_fmt(e['보호자연락처'], 'hyphen')} 로 연락 가능한 시간을 알려 주심.",
    ]


def add_counsel_page(doc, students, page_no, total):
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    text_box(page, pymupdf.Rect(MARGIN, 50, PAGE_W - MARGIN, 85), "상담 기록", size=18)
    y = 110
    for i, text in enumerate(counsel_sentences(students), start=1):
        text_box(page, pymupdf.Rect(MARGIN, y, PAGE_W - MARGIN, y + 90), f"{i}. {text}", size=12)
        y += 100
    text_box(page, pymupdf.Rect(MARGIN, PAGE_H - 50, PAGE_W - MARGIN, PAGE_H - 30),
             f"- {page_no} / {total} -", size=9, align=1)


def main():
    roster = Path(__file__).resolve().parent / "samples" / "가짜_명렬표.xlsx"
    if not roster.exists():
        raise SystemExit("samples/가짜_명렬표.xlsx 가 없습니다. 먼저 python3 make_sample.py 를 실행하세요.")
    students = load_roster()
    if len(students) < 2 * PER_PAGE:
        raise SystemExit(f"학생이 {2 * PER_PAGE}명 이상 필요합니다. (현재 {len(students)}명)")

    doc = pymupdf.open()
    add_table_page(doc, students[:PER_PAGE], "1학년 명단 (1)", 1, 3)
    add_table_page(doc, students[PER_PAGE:2 * PER_PAGE], "1학년 명단 (2)", 2, 3)
    add_counsel_page(doc, students, 3, 3)

    doc.set_metadata({
        "title": f"{students[0]['성명']} 학생 명단 및 상담 기록",   # 학생 이름 하나
        "author": "홍교사",
    })
    doc.save(OUTPUT)
    doc.close()
    print(f"만들었습니다: {OUTPUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
