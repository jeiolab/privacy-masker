"""엑셀 파일을 가명처리하는 어댑터.

글자 치환은 모두 core.masker.mask_text 가 한다. 이 파일은 엑셀의 칸·시트 이름·메모를
돌아다니며 mask_text 를 부르고, 그 결과를 새 파일에 저장하는 일만 한다.
원본 파일은 읽기만 하고 수정하지 않는다.
"""
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from core.masker import mask_text
from core.restorer import unmask_text

_FORMULA_STR = re.compile(r'"((?:[^"]|"")*)"')   # 수식 안 "글자" (안의 "" 는 따옴표 하나)

SHEET_NAME_MAX = 31
_BAD_SHEET_CHARS = re.compile(r"[\\/?*\[\]:]")
_BAD_FILE_CHARS = re.compile(r'[\\/:*?"<>|\[\]]')


@dataclass
class Change:
    sheet: str      # 원래 시트 이름 (파일 이름 변경이면 "(파일 이름)")
    ai_sheet: str   # 가명처리 파일에서의 시트 이름
    cell: str       # 칸 위치 (예: B3, "B3 (메모)", "(시트 이름)")
    original: str
    new: str
    kind: str


def safe_filename(text):
    """파일 이름에 쓸 수 없는 글자를 없앤다 (코드의 [ ] 도 뺀다)."""
    cleaned = _BAD_FILE_CHARS.sub("", text).strip()
    return cleaned or "파일"


def _unique_title(base, used):
    title = _BAD_SHEET_CHARS.sub("", base).strip().strip("'")[:SHEET_NAME_MAX] or "시트"
    n = 1
    candidate = title
    while candidate.lower() in used:
        n += 1
        suffix = f"_{n}"
        candidate = title[: SHEET_NAME_MAX - len(suffix)] + suffix
    return candidate


def _stringify_cell(value):
    """칸 값을 글자로. 서식 있는 글(CellRichText)도 일반 글로 본다."""
    if value is None or isinstance(value, bool):
        return None
    try:
        from openpyxl.cell.rich_text import CellRichText
        if isinstance(value, CellRichText):
            value = str(value)
    except ImportError:
        pass
    return value


def _mask_value(value, dictionary, kinds=None):
    """칸 값 하나를 치환한다. 바뀌면 (새 값, [[원래, 바뀐, 종류], ...]), 아니면 None."""
    value = _stringify_cell(value)
    if isinstance(value, bool) or value is None:
        return None

    if isinstance(value, str):
        if value.startswith("="):          # 수식은 건드리지 않음
            return None
        new, ch = mask_text(value, dictionary, kinds)
        return (new, [[o, n, k] for o, n, k, _ in ch]) if ch else None

    if isinstance(value, (int, float)):
        if isinstance(value, float):
            if not value.is_integer():
                return None
            value = int(value)
        shown = str(value)
        # 엑셀에서 앞 0이 빠진 휴대폰 번호(1012345678)는 0을 붙여 전화번호로 본다
        text = "0" + shown if len(shown) == 10 and shown.startswith("10") else shown
        new, ch = mask_text(text, dictionary, kinds)
        return (new, [[shown, n, k] for _, n, k, _ in ch]) if ch else None

    if isinstance(value, (datetime, date)):  # 생년월일이 날짜 칸으로 들어 있는 경우
        shown = value.strftime("%Y-%m-%d")
        new, ch = mask_text(shown, dictionary, kinds)
        return (new, [[o, n, k] for o, n, k, _ in ch]) if ch else None

    return None


def _walk_workbook(src_path, dst_path, fn):
    """모든 시트 이름, 칸, 메모를 fn 으로 바꿔서 dst_path 에 저장한다 (서식 유지, 수식은 건드리지 않음).

    fn(값) -> None(안 바뀜) 또는 (새 값, [[원래, 바뀐, 종류], ...], [경고, ...])
    (변경 목록, 경고 목록)을 돌려준다.
    """
    wb = load_workbook(src_path)   # 서식 유지를 위해 일반 모드로 연다 (원본은 저장하지 않음)
    changes, warnings = [], []

    def warn_all(sheet, where, ws_warnings):
        warnings.extend(f"{sheet} / {where}: {w}" for w in ws_warnings)

    # 1) 시트 이름: 먼저 새 이름을 정한다 (엑셀이 쓸 수 없는 글자 제거, 31자 제한, 중복 방지)
    sheets = list(wb.worksheets)
    old_titles = [ws.title for ws in sheets]
    used, new_titles, title_results = set(), [], []
    for old in old_titles:
        r = fn(old)
        title_results.append(r)
        final = _unique_title(r[0] if r else old, used)
        used.add(final.lower())
        new_titles.append(final)
    for i, ws in enumerate(sheets):          # 이름이 서로 겹쳐 밀리는 일이 없게 임시 이름을 거친다
        ws.title = f"__tmp{i}__"
    for ws, final in zip(sheets, new_titles):
        ws.title = final
    for old, final, r in zip(old_titles, new_titles, title_results):
        if r:
            changes.extend(Change(old, final, "(시트 이름)", o, n, k) for o, n, k in r[1])
            warn_all(old, "(시트 이름)", r[2])
    renamed = {o: n for o, n in zip(old_titles, new_titles) if o != n}

    # 2) 칸, 메모
    for ws, old in zip(sheets, old_titles):
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                v = cell.value

                if cell.data_type == "f" or (isinstance(v, str) and v.startswith("=")):
                    for o in renamed:   # 바뀐 시트 이름을 가리키는 수식은 깨질 수 있다
                        if isinstance(v, str) and o in v:
                            warnings.append(f"{ws.title}!{cell.coordinate}: 수식이 바뀐 시트 이름 '{o}'을(를) 참조함")
                else:
                    r = fn(v)
                    if r:
                        cell.value = r[0]
                        changes.extend(Change(old, ws.title, cell.coordinate, o, n, k) for o, n, k in r[1])
                        warn_all(old, cell.coordinate, r[2])

                if cell.comment is not None:
                    c = cell.comment
                    text, author = c.text, c.author
                    t = fn(text)
                    a = fn(author) if author else None
                    if t or a:
                        new = Comment(t[0] if t else text, a[0] if a else author)
                        new.width, new.height = c.width, c.height
                        cell.comment = new
                    for res, label in ((t, "메모"), (a, "메모 작성자")):
                        if res:
                            loc = f"{cell.coordinate} ({label})"
                            changes.extend(Change(old, ws.title, loc, o, n, k) for o, n, k in res[1])
                            warn_all(old, loc, res[2])

    # 3) 칸 밖: 수식 안 글자, 머리글·바닥글, 데이터 유효성 목록, 이름 정의, 문서 속성 (2026-10-25 빈틈 시험)
    def put(sheet, title, where, text):
        if not isinstance(text, str) or not text:
            return None
        r = fn(text)
        if not r:
            return None
        changes.extend(Change(sheet, title, where, o, n, k) for o, n, k in r[1])
        warn_all(sheet, where, r[2])
        return r[0]

    def literals(sheet, title, where, formula):   # 수식은 큰따옴표 안 글자만 바꾼다(참조·함수는 그대로)
        if not isinstance(formula, str) or '"' not in formula:
            return formula
        return _FORMULA_STR.sub(lambda m: '"' + (put(sheet, title, where, m.group(1)) or m.group(1)) + '"', formula)

    for ws, old in zip(sheets, old_titles):
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell, MergedCell) and cell.data_type == "f" and isinstance(cell.value, str):
                    cell.value = literals(old, ws.title, f"{cell.coordinate} (수식)", cell.value)
        for hf_name in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"):
            hf = getattr(ws, hf_name, None)
            for part in ("left", "center", "right"):
                p = getattr(hf, part, None) if hf is not None else None
                new = put(old, ws.title, f"({hf_name})", getattr(p, "text", None))
                if new is not None:
                    p.text = new
        for dv in ws.data_validations.dataValidation:
            for attr in ("formula1", "formula2", "prompt", "promptTitle", "error", "errorTitle"):
                v = getattr(dv, attr, None)
                nv = literals(old, ws.title, "(데이터 유효성)", v) if attr.startswith("formula") else put(old, ws.title, "(데이터 유효성)", v)
                if nv is not None and nv != v:
                    setattr(dv, attr, nv)
        for dn in getattr(ws, "defined_names", {}).values():
            dn.attr_text = literals(old, ws.title, f"(이름 정의 {dn.name})", dn.attr_text)
    for dn in wb.defined_names.values():
        dn.attr_text = literals("(통합 문서)", "", f"(이름 정의 {dn.name})", dn.attr_text)
    props = wb.properties
    for attr in ("creator", "lastModifiedBy", "title", "subject", "description", "keywords", "category"):
        new = put("(문서 속성)", "", attr, getattr(props, attr, None))
        if new is not None:
            setattr(props, attr, new)

    wb.save(dst_path)
    return changes, warnings


def _cell_scan_text(value):
    """가릴 항목을 찾을 때 칸 하나를 글자로. 가리기와 같은 모양(앞 0이 빠진 번호 포함)으로 본다."""
    value = _stringify_cell(value)
    if isinstance(value, bool) or value is None:
        return ""
    if isinstance(value, str):
        return "" if value.startswith("=") else value
    if isinstance(value, (int, float)):
        if isinstance(value, float):
            if not value.is_integer():
                return ""
            value = int(value)
        shown = str(value)
        return "0" + shown if len(shown) == 10 and shown.startswith("10") else shown
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return ""


def workbook_plain_text(path):
    """시트 이름, 칸, 메모를 한 글로 잇는다. 수식 칸은 빼 둔다. 원본은 저장하지 않는다."""
    wb = load_workbook(path)
    parts = []
    try:
        for ws in wb.worksheets:
            parts.append(ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue
                    if cell.data_type == "f":
                        text = ""
                    else:
                        text = _cell_scan_text(cell.value)
                    if text:
                        parts.append(text)
                    if cell.comment is not None:
                        if cell.comment.text:
                            parts.append(cell.comment.text)
                        if cell.comment.author:
                            parts.append(cell.comment.author)
    finally:
        wb.close()
    return "\n".join(parts)


def mask_workbook(src_path, dst_path, dictionary, kinds=None):
    """src_path 를 가명처리해서 dst_path 에 저장. (변경 목록, 경고 목록)을 돌려준다."""
    def fn(value):
        r = _mask_value(value, dictionary, kinds)
        return (r[0], r[1], []) if r else None
    return _walk_workbook(src_path, dst_path, fn)


def restore_workbook(src_path, dst_path, dictionary, extra=None):
    """가명처리된 엑셀(AI가 돌려준 파일)을 실명으로 복원해서 저장. (변경 목록, 경고 목록)."""
    extra = extra or {}
    def fn(value):
        if not isinstance(value, str) or value.startswith("="):
            return None
        new, ch, warns = unmask_text(value, dictionary, extra)
        if not ch and not warns:
            return None
        return new, [[o, n, k] for o, n, k, _ in ch], warns
    return _walk_workbook(src_path, dst_path, fn)


def apply_extra_names_workbook(path, words, start=1):
    """가명처리된 엑셀에서 이름 단어를 [이름N] 으로 바꿔 덮어쓴다.

    (Change 목록, {코드: 단어}) 를 돌려준다. 수식 칸은 건드리지 않는다.
    """
    uniq = list(dict.fromkeys(w for w in words if w))
    name_map = {f"[이름{i}]": w for i, w in enumerate(uniq, start=start)}
    if not name_map:
        return [], {}
    word_to_code = {w: c for c, w in name_map.items()}
    # 꼬리말/조사(씨, 에게, 와 ...)가 바로 뒤에 붙어도 이름을 바꾸도록 오른쪽 경계는 두지 않는다
    alt = "|".join(re.escape(w) for w in sorted(word_to_code, key=lambda w: (-len(w), w)))
    pat = re.compile(rf"(?<![가-힣])(?:{alt})")
    changes = []
    wb = load_workbook(path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                v = cell.value
                if not isinstance(v, str) or v.startswith("="):
                    continue
                hits = []

                def repl(m, _hits=hits):
                    _hits.append(m.group(0))
                    return word_to_code[m.group(0)]

                new = pat.sub(repl, v)
                if hits:
                    cell.value = new
                    for w in hits:
                        changes.append(Change(ws.title, ws.title, cell.coordinate, w,
                                              word_to_code[w], "이름의심확정"))
    wb.save(path)
    return changes, name_map


def append_name_map_sheet(wb, name_map):
    """변경기록 엑셀에 [이름N] ↔ 원래 단어 시트를 붙인다. 없으면 만들지 않는다."""
    if not name_map:
        return
    ws = wb.create_sheet("이름대응")
    ws.append(["코드", "원래 단어"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for code, word in name_map.items():
        ws.append([code, word])
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
            cell.number_format = "@"
    for i in range(1, 3):
        letter = get_column_letter(i)
        longest = max((len(str(c.value or "")) for c in ws[letter]), default=8)
        ws.column_dimensions[letter].width = min(max(longest + 3, 10), 40)
    ws.freeze_panes = "A2"


def load_name_map(path):
    """변경기록 엑셀의 이름대응 시트에서 {코드: 원래 단어} 를 읽는다. 없으면 빈 dict."""
    path = Path(path)
    if not path.exists():
        return {}
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if "이름대응" not in wb.sheetnames:
            return {}
        out = {}
        for i, row in enumerate(wb["이름대응"].iter_rows(values_only=True)):
            if i == 0 or not row:
                continue
            code = str(row[0] or "").strip()
            word = str(row[1] or "").strip() if len(row) > 1 else ""
            if code.startswith("[이름") and word:
                out[code] = word
        return out
    finally:
        wb.close()


def write_change_log(changes, path, name_map=None):
    """보관용 변경기록 엑셀: 시트, 칸 위치, 원래 값, 바뀐 값, 종류."""
    wb = Workbook()
    ws = wb.active
    ws.title = "변경기록"
    ws.append(["시트", "칸 위치", "원래 값", "바뀐 값", "종류", "AI용 파일의 시트 이름"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for ch in changes:
        ws.append([ch.sheet, ch.cell, ch.original, ch.new, ch.kind, ch.ai_sheet])
    for row in ws.iter_rows():
        for cell in row:   # 모든 값을 글자로 저장 (앞 0·'=' 로 시작하는 글이 변하지 않게)
            if isinstance(cell.value, str):
                cell.data_type = "s"
            cell.number_format = "@"
    for i in range(1, 7):
        letter = get_column_letter(i)
        longest = max((len(str(c.value or "")) for c in ws[letter]), default=8)
        ws.column_dimensions[letter].width = min(max(longest + 3, 10), 60)
    ws.freeze_panes = "A2"
    append_name_map_sheet(wb, name_map)
    wb.save(path)
