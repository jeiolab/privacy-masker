#!/usr/bin/env python3
"""명렬표 엑셀을 읽어서 '학생 사전' 엑셀을 만든다.

사용법: python3 build_dict.py samples/가짜_명렬표.xlsx
결과  : data/학생사전.xlsx  (원본 파일은 읽기만 하고 수정하지 않는다)
"""
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from core import config

BASE_DIR = config.user_data_dir()   # 문서\개인정보가리개_데이터 (경로는 config 한 곳)


# ---------------------------------------------------------------- 유틸
def norm(text):
    """비교용: 문자열로 바꾸고 모든 공백 제거."""
    return re.sub(r"\s+", "", str(text)) if text is not None else ""


def cell_to_str(value):
    """셀 값을 문자열로 (숫자 .0 제거, 날짜는 YYYY-MM-DD)."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def format_phone(raw):
    """전화번호를 하이픈 형식으로. (바꾸지 못하면 None)"""
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 10 and digits.startswith("10"):  # 엑셀이 앞 0을 지운 경우
        digits = "0" + digits
    if digits.startswith("02"):
        if len(digits) == 9:
            return f"{digits[:2]}-{digits[2:5]}-{digits[5:]}"
        if len(digits) == 10:
            return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"
    elif len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    elif len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return None


def split_given_name(name):
    """성 뺀 이름. 두 글자 성은 config 목록 기준."""
    for sur in config.COMPOUND_SURNAMES:
        if name.startswith(sur) and len(name) > len(sur):
            return name[len(sur):]
    return name[1:] if len(name) > 1 else name


# ---------------------------------------------------------------- 열 찾기
def classify_header(header):
    """머리글 한 칸이 어떤 필드인지 반환 (없으면 None)."""
    h = norm(header)
    if not h:
        return None
    cands = {f: [norm(c) for c in cs] for f, cs in config.COLUMN_CANDIDATES.items()}
    for field, cs in cands.items():
        if h in cs:
            return field
    # 학생 연락처 후보 단어가 들어 있고 보호자/학부모/부모도 들어 있으면 보호자연락처
    has_guardian = any(k in h for k in config.GUARDIAN_KEYWORDS)
    if has_guardian and any(c in h for c in cands["학생연락처"]):
        return "보호자연락처"
    return None


def find_header_row(rows):
    """위에서 N줄 안에서 후보 단어가 2개 이상 든 첫 줄의 인덱스."""
    for i, row in enumerate(rows[: config.HEADER_SCAN_ROWS]):
        hits = sum(1 for v in row if classify_header(v))
        if hits >= config.HEADER_MIN_MATCHES:
            return i
    return None


def map_columns(header_row):
    """필드 -> 열 인덱스. 같은 필드가 여러 열이면 맨 왼쪽만 사용."""
    mapping = {}
    for idx, h in enumerate(header_row):
        field = classify_header(h)
        if field and field not in mapping:
            mapping[field] = idx
    return mapping


def load_rows(path):
    """원본을 읽기 전용으로 열어 (시트명, 행 목록) 반환."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if find_header_row(rows) is not None:
                return ws.title, rows
    finally:
        wb.close()
    return None, None


# ---------------------------------------------------------------- 사전 만들기
def build_students(rows, header_idx, mapping, warnings):
    def get(row, field):
        idx = mapping.get(field)
        if idx is None or idx >= len(row):
            return ""
        return cell_to_str(row[idx])

    students = []
    for order, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        name = norm(get(row, "성명"))
        if not name:
            warnings.append(f"{order}행: 성명이 비어 있어 건너뜀")
            continue
        s = {
            "행": order,
            "학번": get(row, "학번"),
            "성명": name,
            "보호자명": get(row, "보호자명"),
            "학생연락처": get(row, "학생연락처"),
            "보호자연락처": get(row, "보호자연락처"),
            "주소": get(row, "주소"),
            "생년월일": get(row, "생년월일"),
        }
        for f in ("학생연락처", "보호자연락처"):
            if s[f]:
                fixed = format_phone(s[f])
                if fixed:
                    s[f] = fixed
                else:
                    warnings.append(f"{name}: {f} '{s[f]}' 형식을 인식하지 못해 그대로 둠")
        s["성뺀이름"] = split_given_name(name)
        students.append(s)
    return students


def sort_key(s):
    sid = s["학번"]
    return (0, int(sid), s["행"]) if sid.isdigit() else (1, sid, s["행"])


def assign_codes(students, mapping, warnings):
    if "학번" in mapping:
        students.sort(key=sort_key)
        ids = [s["학번"] for s in students if s["학번"]]
        for sid in sorted({i for i in ids if ids.count(i) > 1}):
            warnings.append(f"학번 {sid} 이(가) 중복됨")
    else:
        students.sort(key=lambda s: s["행"])  # 학번이 없으면 파일 순서
        warnings.append("학번 열이 없어 파일에 나온 순서대로 번호를 붙임")

    width = max(3, len(str(len(students))))
    for n, s in enumerate(students, start=1):
        num = str(n).zfill(width)
        s["번호"] = num
        s["학생코드"] = f"학생{num}"
        s["S코드"] = f"S{num}"
        s["보호자코드"] = f"보호자{num}" if s["보호자명"] else ""
        s["학생전화코드"] = f"[학생전화{num}]" if s["학생연락처"] else ""
        s["보호자전화코드"] = f"[보호자전화{num}]" if s["보호자연락처"] else ""
        s["주소코드"] = f"[주소{num}]" if s["주소"] else ""
        s["생년월일코드"] = f"[생년월일{num}]" if s["생년월일"] else ""


def add_cautions(students, warnings):
    by_name = defaultdict(list)
    for s in students:
        by_name[s["성명"]].append(s)

    common = set(config.COMMON_WORD_NAMES)
    for s in students:
        notes = []
        same = by_name[s["성명"]]
        if len(same) > 1:
            notes.append("동명이인")
        if len(s["성뺀이름"]) == 1:
            notes.append("외자 이름")
        if s["성뺀이름"] in common:
            notes.append("일반 단어와 겹침")
        s["주의"] = ", ".join(notes)

    for name, group in by_name.items():
        if len(group) > 1:
            nums = ", ".join(f"{g['번호']}번({g['학번'] or '학번없음'})" for g in group)
            warnings.append(f"동명이인: {name} × {len(group)}명 [{nums}]")
    for s in students:
        if "외자 이름" in s["주의"]:
            warnings.append(f"외자 이름: {s['성명']} → '{s['성뺀이름']}' ({s['번호']}번)")
        if "일반 단어와 겹침" in s["주의"]:
            warnings.append(f"일반 단어와 겹치는 이름: {s['성명']} → '{s['성뺀이름']}' ({s['번호']}번)")


# ---------------------------------------------------------------- 저장
OUTPUT_COLUMNS = [
    "번호", "학번", "성명", "성뺀이름", "보호자명",
    "학생연락처", "보호자연락처", "주소", "생년월일",
    "학생코드", "S코드", "보호자코드",
    "학생전화코드", "보호자전화코드", "주소코드", "생년월일코드", "주의",
]
# 코드 열의 머리글 표시 이름
HEADER_LABELS = {
    "학생코드": "학생코드", "S코드": "S코드", "보호자코드": "보호자코드",
    "학생전화코드": "학생전화코드", "보호자전화코드": "보호자전화코드",
    "주소코드": "주소코드", "생년월일코드": "생년월일코드",
}


STUDENT_SHEET = "학생사전"
STAFF_SHEET = "교직원"
INFO_SHEET = "사전정보"
SHEET_ORDER = [STUDENT_SHEET, STAFF_SHEET, INFO_SHEET]


def _open_or_new_workbook(out_path):
    """사전 파일이 있으면 열어서(다른 시트 보존), 없으면 빈 Workbook 을 만든다."""
    out_path = Path(out_path)
    if out_path.exists():
        return load_workbook(out_path)
    wb = Workbook()
    wb.remove(wb.active)   # 기본 'Sheet' 제거 (시트는 우리가 직접 만든다)
    return wb


def _fresh_sheet(wb, title):
    """같은 이름 시트가 있으면 지우고 새로 만든다 (내용만 갈아끼운다)."""
    if title in wb.sheetnames:
        del wb[title]
    return wb.create_sheet(title)


def _write_format_info(wb):
    """사전형식버전을 '사전정보' 시트에 기록한다. (예전 사전도 읽기만 하면 된다)"""
    ws = _fresh_sheet(wb, INFO_SHEET)
    _fill_sheet(ws, ["항목", "값"], [["사전형식버전", str(config.DICT_FORMAT_VERSION)]])


def read_dict_format_version(path):
    """사전 파일의 사전형식버전. 시트가 없으면 0 (예전 형식)."""
    path = Path(path)
    if not path.exists():
        return 0
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if INFO_SHEET not in wb.sheetnames:
            return 0
        rows = list(wb[INFO_SHEET].iter_rows(values_only=True))
    finally:
        wb.close()
    for row in rows[1:]:
        if row and str(row[0] or "").strip() == "사전형식버전":
            try:
                return int(row[1])
            except (TypeError, ValueError):
                return 0
    return 0


def _order_sheets(wb):
    """학생사전 → 교직원 순서로 시트를 정렬한다 (알려진 시트부터, 나머지는 뒤에)."""
    known = [t for t in SHEET_ORDER if t in wb.sheetnames]
    rest = [t for t in wb.sheetnames if t not in SHEET_ORDER]
    wb._sheets.sort(key=lambda ws: (known + rest).index(ws.title))


def _fill_sheet(ws, header_cells, rows):
    """머리글 + 행을 채우고 모든 값은 텍스트(@) 형식으로, 보기 좋게 열 너비를 맞춘다."""
    ws.append(header_cells)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.number_format = "@"   # 전화번호/번호가 숫자로 바뀌지 않도록
    for i, head in enumerate(header_cells, start=1):
        longest = max([len(str(head))] + [len(str(r[i - 1])) for r in rows]) if rows else len(str(head))
        ws.column_dimensions[get_column_letter(i)].width = min(longest + 4, 50)
    ws.freeze_panes = "A2"


def _write_workbook(students, cols, out_path):
    """학생 목록을 cols 순서대로 '학생사전' 시트로 저장한다 (교직원 시트는 그대로 보존)."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = _open_or_new_workbook(out_path)
    ws = _fresh_sheet(wb, STUDENT_SHEET)
    header = [HEADER_LABELS.get(c, c) for c in cols]
    rows = [[s.get(c, "") for c in cols] for s in students]
    _fill_sheet(ws, header, rows)
    _write_format_info(wb)
    _order_sheets(wb)
    wb.save(out_path)


def write_dict(students, mapping, out_path):
    # 원본에 없는 열(보호자명 등)에서 나온 빈 열은 제외
    skip = set()
    if "보호자명" not in mapping:
        skip |= {"보호자명", "보호자코드"}
    if "학생연락처" not in mapping:
        skip |= {"학생연락처", "학생전화코드"}
    if "보호자연락처" not in mapping:
        skip |= {"보호자연락처", "보호자전화코드"}
    if "주소" not in mapping:
        skip |= {"주소", "주소코드"}
    if "생년월일" not in mapping:
        skip |= {"생년월일", "생년월일코드"}
    if "학번" not in mapping:
        skip.add("학번")
    cols = [c for c in OUTPUT_COLUMNS if c not in skip]
    _write_workbook(students, cols, out_path)


# ---------------------------------------------------------------- 추가 모드(사전 업데이트)
BACKUP_DIR = config.backup_dir()

# 업데이트 때 값만 갱신하는 정보 열(코드 번호는 절대 바꾸지 않는다)
INFO_FIELDS = ["보호자명", "학생연락처", "보호자연락처", "주소", "생년월일"]

# 기존 사전을 다시 student dict 로 복원할 때 쓰는 모든 열
_DICT_READ_COLUMNS = OUTPUT_COLUMNS


def load_existing_dict(path):
    """기존 학생사전 엑셀을 읽어 student dict 목록으로 되살린다. (파일이 없으면 None)"""
    path = Path(path)
    if not path.exists():
        return None
    wb = load_workbook(path, data_only=True)
    try:
        ws = wb["학생사전"] if "학생사전" in wb.sheetnames else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    if not rows:
        return []
    header = [norm(h) for h in rows[0]]
    students = []
    for row in rows[1:]:
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        s = {h: cell_to_str(v) for h, v in zip(header, row) if h}
        if not s.get("성명"):
            continue
        students.append(s)
    return students


def backup_dict(path):
    """사전을 바꾸기 전에 data/백업/학생사전_날짜시간.xlsx 로 복사한다. (백업 경로 반환)"""
    import shutil

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = BACKUP_DIR / f"학생사전_{stamp}.xlsx"
    shutil.copy2(path, dst)
    return dst


def _derive_codes(s):
    """'번호'를 기준으로 코드 열을 다시 계산한다. (번호 자체는 절대 바꾸지 않는다)"""
    num = s["번호"]
    s["학생코드"] = f"학생{num}"
    s["S코드"] = f"S{num}"
    s["보호자코드"] = f"보호자{num}" if s.get("보호자명") else ""
    s["학생전화코드"] = f"[학생전화{num}]" if s.get("학생연락처") else ""
    s["보호자전화코드"] = f"[보호자전화{num}]" if s.get("보호자연락처") else ""
    s["주소코드"] = f"[주소{num}]" if s.get("주소") else ""
    s["생년월일코드"] = f"[생년월일{num}]" if s.get("생년월일") else ""


def merge_students(old_students, new_students, mapping, warnings):
    """기존 사전 + 새 명렬표를 병합한다. 한 번 붙은 코드 번호는 절대 바꾸지 않는다.

    Returns:
        (merged, stats) — stats = {"kept","updated","added","missing"}
    """
    use_id = ("학번" in mapping) and any(s.get("학번") for s in old_students)
    if not use_id:
        warnings.append("학번이 없어 성명으로 비교합니다(동명이인이 있으면 부정확할 수 있음).")

    def key_of(s):
        return s["학번"] if use_id and s.get("학번") else s.get("성명", "")

    old_by_key, new_by_key = {}, {}
    for s in old_students:
        old_by_key.setdefault(key_of(s), s)
    for s in new_students:
        new_by_key.setdefault(key_of(s), s)

    # 새로 붙일 번호: 기존 최대 번호 다음부터. 자릿수는 기존과 동일하게 유지.
    existing_nums = [int(s["번호"]) for s in old_students if str(s.get("번호", "")).isdigit()]
    next_num = (max(existing_nums) + 1) if existing_nums else 1
    width = max([1] + [len(str(s.get("번호", ""))) for s in old_students if s.get("번호")])

    merged = []
    kept = updated = added = missing = 0

    # 1) 기존 학생: 코드 번호 유지. 명렬표에 있으면 정보 갱신, 없으면 보관 표시.
    for old in old_students:
        s = dict(old)
        k = key_of(old)
        if k in new_by_key:
            new = new_by_key[k]
            changed = False
            for f in INFO_FIELDS:
                if (new.get(f) or "") != (s.get(f) or ""):
                    s[f] = new.get(f, "")
                    changed = True
            _derive_codes(s)
            if changed:
                updated += 1
            else:
                kept += 1
        else:
            s["_빠짐"] = True
            missing += 1
        merged.append(s)

    # 2) 새 학생(전학 등): 가장 큰 번호 다음 번호를 붙여 추가 (학번 순으로 정렬해서 번호 부여)
    new_only = [s for s in new_students if key_of(s) not in old_by_key]
    new_only.sort(key=sort_key)
    for new in new_only:
        s = dict(new)
        s["번호"] = str(next_num).zfill(width)
        next_num += 1
        _derive_codes(s)
        merged.append(s)
        added += 1

    # 이름이 바뀌었을 수 있으니 성 뺀 이름은 항상 다시 계산
    for s in merged:
        s["성뺀이름"] = split_given_name(s["성명"])

    return merged, {"kept": kept, "updated": updated, "added": added, "missing": missing}


def apply_missing_notes(merged, warnings):
    """명렬표에서 빠진 학생의 '주의' 열에 '명렬표에 없음'을 표시한다. (삭제하지 않는다)"""
    for s in merged:
        if s.pop("_빠짐", False):
            note = "명렬표에 없음"
            s["주의"] = note if not s.get("주의") else f"{note}, {s['주의']}"
            warnings.append(f"명렬표에 없음: {s['성명']} ({s['번호']}번) - 사전에 그대로 보관함")


def _num_key(s):
    num = s.get("번호", "")
    return (0, int(num)) if str(num).isdigit() else (1, num)


def columns_for(students):
    """실제 값이 있는 열만 고른다(핵심 열은 항상 포함). 업데이트 저장용."""
    always = {"번호", "성명", "성뺀이름", "학생코드", "S코드", "주의"}
    return [c for c in OUTPUT_COLUMNS if c in always or any(s.get(c) for s in students)]


def write_dict_incremental(students, out_path):
    """업데이트된 학생 목록을 저장한다 (값이 있는 열만)."""
    _write_workbook(students, columns_for(students), out_path)


# ---------------------------------------------------------------- 교직원 명단
STAFF_COLUMNS = ["번호", "성명", "성뺀이름", "교사코드", "주의"]
STAFF_CODE_WIDTH = 3   # 교사001, 교사002 ...


def _staff_code(num):
    return f"교사{num}"


def _extract_staff_names(rows):
    """명단 표에서 교직원 '성명'만 뽑는다. (성명 열이 있으면 그 열, 없으면 한 열 목록으로 본다)"""
    if not rows:
        return []
    name_terms = {norm(c) for c in config.COLUMN_CANDIDATES["성명"]}
    skip_terms = name_terms | {"번호", "순번", "연번", "no", "교직원", "직원", "교사", "선생님"}

    # 1) 어느 줄이든 '성명/이름' 머리글 칸이 있으면 그 열을 쓴다
    for i, row in enumerate(rows[: config.HEADER_SCAN_ROWS]):
        for j, cell in enumerate(row):
            if norm(cell) in name_terms:
                return [norm(cell_to_str(r[j])) for r in rows[i + 1:]
                        if j < len(r) and norm(cell_to_str(r[j]))]

    # 2) 머리글이 없으면: 각 줄에서 맨 앞 값이 든 칸을 이름으로 본다 (머리글 글자는 건너뜀)
    names = []
    for row in rows:
        first = next((cell_to_str(c) for c in row if cell_to_str(c).strip()), "")
        n = norm(first)
        if n and n not in skip_terms:
            names.append(n)
    return names


def load_staff_names(path):
    """교직원 명단 엑셀을 읽어 (시트명, 이름 목록)을 돌려준다. (성명 한 열만 있어도 됨)"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            names = _extract_staff_names(rows)
            if names:
                return ws.title, names
    finally:
        wb.close()
    return None, []


def load_existing_staff(path):
    """학생사전 파일의 '교직원' 시트를 읽는다. 파일이 없으면 None, 시트가 없으면 []."""
    path = Path(path)
    if not path.exists():
        return None
    wb = load_workbook(path, data_only=True)
    try:
        if STAFF_SHEET not in wb.sheetnames:
            return []
        rows = [list(r) for r in wb[STAFF_SHEET].iter_rows(values_only=True)]
    finally:
        wb.close()
    if not rows:
        return []
    header = [norm(h) for h in rows[0]]
    staff = []
    for row in rows[1:]:
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        s = {h: cell_to_str(v) for h, v in zip(header, row) if h}
        if s.get("성명"):
            staff.append(s)
    return staff


def merge_staff(old_staff, new_names, warnings):
    """기존 교직원 + 새 명단을 병합한다. 한 번 붙은 교사 코드 번호는 절대 바꾸지 않는다.

    Returns:
        (merged, stats) — stats = {"kept","added","missing"}
    """
    old_staff = old_staff or []
    # 새 명단의 중복 이름은 한 번만 쓴다 (순서 유지)
    uniq_names, seen_new = [], set()
    for name in new_names:
        if name and name not in seen_new:
            seen_new.add(name)
            uniq_names.append(name)
        elif name:
            warnings.append(f"교직원 명단에 '{name}' 이(가) 여러 번 있어 한 번만 넣음")
    new_set = set(uniq_names)

    old_by_name = {}
    for s in old_staff:
        old_by_name.setdefault(s.get("성명", ""), s)

    existing_nums = [int(s["번호"]) for s in old_staff if str(s.get("번호", "")).isdigit()]
    next_num = (max(existing_nums) + 1) if existing_nums else 1
    width = max([STAFF_CODE_WIDTH] + [len(str(s.get("번호", ""))) for s in old_staff if s.get("번호")])

    merged = []
    kept = missing = added = 0

    # 1) 기존 교직원: 코드 번호 유지. 명단에 없으면 보관하되 표시.
    for old in old_staff:
        s = dict(old)
        s.pop("주의", None)   # '명단에 없음' 표시는 매번 다시 계산
        if s.get("성명") in new_set:
            kept += 1
        else:
            s["주의"] = "명단에 없음"
            missing += 1
            warnings.append(f"명단에 없음: {s.get('성명')} (교사{s.get('번호')}번) - 사전에 그대로 보관함")
        merged.append(s)

    # 2) 새 교직원: 가장 큰 번호 다음 번호를 붙여 추가
    for name in uniq_names:
        if name in old_by_name:
            continue
        s = {"번호": str(next_num).zfill(width), "성명": name}
        next_num += 1
        added += 1
        merged.append(s)

    for s in merged:
        s["성뺀이름"] = split_given_name(s["성명"])
        s["교사코드"] = _staff_code(s["번호"])

    merged.sort(key=_num_key)
    return merged, {"kept": kept, "added": added, "missing": missing}


def write_staff_sheet(staff, out_path):
    """교직원 목록을 '교직원' 시트로 저장한다 (학생사전 시트는 그대로 보존).

    학생사전 시트가 아직 없으면 머리글만 있는 빈 학생사전 시트를 함께 만들어 둔다.
    """
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = _open_or_new_workbook(out_path)
    if STUDENT_SHEET not in wb.sheetnames:   # 학생 없이 교직원만 먼저 넣는 경우
        base = wb.create_sheet(STUDENT_SHEET)
        _fill_sheet(base, OUTPUT_COLUMNS, [])
    ws = _fresh_sheet(wb, STAFF_SHEET)
    rows = [[s.get(c, "") for c in STAFF_COLUMNS] for s in staff]
    _fill_sheet(ws, STAFF_COLUMNS, rows)
    _write_format_info(wb)
    _order_sheets(wb)
    wb.save(out_path)


# ---------------------------------------------------------------- 실행
def main(argv):
    if len(argv) != 2:
        print("사용법: python3 build_dict.py <명렬표.xlsx>")
        return 1

    src = Path(argv[1]).expanduser()
    out_path = config.dict_path()
    if not src.is_file():
        print(f"[오류] 파일을 찾을 수 없습니다: {src}")
        return 1
    if src.resolve() == out_path.resolve():
        print("[오류] 입력 파일이 결과 파일과 같습니다.")
        return 1

    try:
        sheet, rows = load_rows(src)
    except Exception as e:  # 손상/비밀번호 파일 등
        print(f"[오류] 엑셀을 읽지 못했습니다: {e}")
        return 1
    if rows is None:
        print(f"[오류] 위에서 {config.HEADER_SCAN_ROWS}줄 안에서 머리글 줄을 찾지 못했습니다.")
        return 1

    header_idx = find_header_row(rows)
    mapping = map_columns(rows[header_idx])
    warnings = []

    found = [f for f in config.COLUMN_CANDIDATES if f in mapping]
    missing = [f for f in config.COLUMN_CANDIDATES if f not in mapping]

    if "성명" not in mapping:
        print("[오류] 성명 열을 찾지 못해 사전을 만들 수 없습니다.")
        print(f"  못 찾은 열: {', '.join(missing)}")
        return 1
    for f in missing:
        warnings.append(f"{f} 열을 찾지 못해 건너뜀")

    new_students = build_students(rows, header_idx, mapping, warnings)

    existing = load_existing_dict(out_path)
    stats = backup = None
    if existing is None:
        # data/학생사전.xlsx 가 없으면: 지금처럼 새로 만든다
        students = new_students
        assign_codes(students, mapping, warnings)
        add_cautions(students, warnings)
        write_dict(students, mapping, out_path)
        mode = "새로 만들기"
    else:
        # 이미 있으면: 바꾸기 전에 백업하고, 기존 코드 번호를 유지한 채 업데이트한다
        backup = backup_dict(out_path)
        students, stats = merge_students(existing, new_students, mapping, warnings)
        add_cautions(students, warnings)
        apply_missing_notes(students, warnings)
        students.sort(key=_num_key)
        write_dict_incremental(students, out_path)
        mode = "사전 업데이트"

    print(f"원본: {src} (시트 '{sheet}', 머리글 {header_idx + 1}행)")
    print(f"모드: {mode}")
    if backup is not None:
        print(f"백업: {backup}")
    print("찾은 열  : " + ", ".join(
        f"{f}('{cell_to_str(rows[header_idx][mapping[f]])}')" for f in found))
    print("못 찾은 열: " + (", ".join(missing) if missing else "없음"))
    print(f"총 학생 수: {len(students)}명")
    if stats is not None:
        print(f"요약: 유지 {stats['kept']}명 / 새로 추가 {stats['added']}명 / "
              f"정보 갱신 {stats['updated']}명 / 명렬표에 없음 {stats['missing']}명")
    print(f"경고 {len(warnings)}건" + (":" if warnings else ""))
    for w in warnings:
        print(f"  - {w}")
    print(f"결과: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
