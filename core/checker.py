"""가명처리 결과 엑셀에 개인정보가 남았는지 검사한다.

결과 파일을 다시 읽어서 (칸, 시트 이름, 메모) 사전의 성명 전체, 보호자명, 학번,
전화번호(숫자만 비교)가 남아 있는지 본다. 수식 칸도 글자로 읽어서 함께 검사한다.
"""
import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from core import config


@dataclass
class Finding:
    sheet: str
    location: str   # 칸 위치 / "(시트 이름)" / "B3 (메모)"
    kind: str       # 성명 / 보호자명 / 학번 / 전화번호
    value: str      # 남아 있는 글자


@dataclass
class Suspect:
    """사전에 없지만 사람 이름처럼 보이는 단어 (자동으로 바꾸지 않고 경고만 한다)."""
    location: str   # "N번째 줄" / 칸 위치
    word: str       # 의심되는 이름 (예: 김민준)
    context: str    # 앞뒤 문장


# 이미 코드로 바뀐 부분(학생003, 교사001, S003, [학생전화003], [이름1] 등)은 의심에서 제외
_CODE_RE = re.compile(
    r"(?:학생|보호자|교사)\s?\d{1,4}(?:또는\d{1,4})*"
    r"|S\s?\d{3,4}(?:또는\d{3,4})*"
    r"|\[(?:학생전화|보호자전화|주소|생년월일|이름)\d{0,4}(?:또는\d+)*\]"
    r"|\[(?:전화번호|주민번호|이메일)\]"
)


def _suspect_re():
    surnames = sorted(config.KOREAN_SURNAMES, key=lambda s: (-len(s), s))
    suffixes = sorted(config.NAME_SUSPECT_SUFFIXES, key=lambda s: (-len(s), s))
    sur = "|".join(re.escape(s) for s in surnames)
    suf = "|".join(re.escape(s) for s in suffixes)
    # 성씨로 시작하는 2~4글자(이름 부분은 되도록 짧게) + (공백?) + 의심 꼬리말.
    # 꼬리말 뒤에는 조사(학생과, 씨에게 ...)가 붙을 수 있으므로 오른쪽 경계는 두지 않는다.
    return re.compile(
        rf"(?<![가-힣])(?P<word>(?:{sur})[가-힣]{{1,3}}?)[ \t]?(?P<suffix>{suf})"
    )


def _default_ignore_path():
    return config.ignore_words_path()


def load_ignored_words(path=None):
    """data/무시단어.txt (또는 path) 에서 '항상 무시' 단어를 읽는다. 파일이 없으면 빈 목록."""
    p = Path(path) if path is not None else _default_ignore_path()
    if not p.exists():
        return []
    words = []
    for line in p.read_text(encoding="utf-8").splitlines():
        w = line.strip()
        if w and not w.startswith("#"):
            words.append(w)
    return words


def _given_name(name):
    """성 뺀 이름. 두 글자 성은 config 목록 기준."""
    name = name or ""
    for sur in config.COMPOUND_SURNAMES:
        if name.startswith(sur) and len(name) > len(sur):
            return name[len(sur):]
    return name[1:] if len(name) > 1 else name


def real_name_words(dictionary):
    """사전에 있는 실명: 성명 전체, 성 뺀 이름, 보호자명, 교직원명."""
    names = set()
    for s in dictionary.students:
        names.add(s.get("성명", ""))
        names.add(s.get("성뺀이름") or _given_name(s.get("성명", "")))
        names.add(s.get("보호자명", ""))
    for t in getattr(dictionary, "staff", []):
        names.add(t.get("성명", ""))
        names.add(t.get("성뺀이름") or _given_name(t.get("성명", "")))
    names.discard("")
    return names


def add_ignored_word(word, dictionary, path=None):
    """단어를 '항상 무시' 목록에 저장한다.
    성공하면 None, 사전 실명이면 '실제 이름은 항상 무시할 수 없어요' 를 돌려준다.
    """
    word = (word or "").strip()
    if not word:
        return None
    if word in real_name_words(dictionary):
        return "명단에 있는 이름은 빼 둘 수 없어요"
    p = Path(path) if path is not None else _default_ignore_path()
    if word in load_ignored_words(p):
        return None
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("a", encoding="utf-8", newline="\n") as f:
        f.write(word + "\n")
    return None


def remove_ignored_words(words, path=None):
    """무시 목록에서 고른 단어를 뺀다. (지운 개수)"""
    p = Path(path) if path is not None else _default_ignore_path()
    drop = {w.strip() for w in words if (w or "").strip()}
    kept = [w for w in load_ignored_words(p) if w not in drop]
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(("\n".join(kept) + "\n") if kept else "", encoding="utf-8")
    return len(drop)


def _is_exception(word, exceptions):
    if word in exceptions:
        return True
    return word.endswith("들") and word[:-1] in exceptions   # 남학생들 → 남학생


def find_name_suspects(text, dictionary, ignore=(), ignore_path=None):
    """글에서 '사전에 없지만 이름처럼 보이는 단어'를 찾는다. (단어별로 한 번씩, 첫 문장 맥락과 함께)

    - 이미 코드로 바뀐 부분과 config 의 예외 단어, 사전에 든 이름은 제외한다.
    - ignore 에 든 단어와 data/무시단어.txt 의 '항상 무시' 단어도 제외한다.
    - 자동 치환은 하지 않는다.
    """
    pat = _suspect_re()
    exceptions = set(config.NAME_SUSPECT_EXCEPTIONS)
    particles = set(config.NAME_SUSPECT_PARTICLES)   # 성씨 뒤가 조사·시간 표현뿐이면 제외
    ignore = set(ignore) | set(load_ignored_words(ignore_path))
    dict_names = set()
    for s in dictionary.students:
        dict_names.add(s.get("성명", ""))
        dict_names.add(s.get("보호자명", ""))
    for t in getattr(dictionary, "staff", []):
        dict_names.add(t.get("성명", ""))
    dict_names.discard("")

    suspects, seen = [], set()
    for n, line in enumerate(text.splitlines(), start=1):
        # 코드로 바뀐 구간은 비워서, 그 안/경계에서 의심 단어를 찾지 않게 한다
        safe = _CODE_RE.sub(lambda m: " " * len(m.group(0)), line)
        for m in pat.finditer(safe):
            word = m.group("word")
            if _is_exception(word, exceptions) or word in dict_names or word in ignore or word in seen:
                continue
            if config.is_field_title(word):
                continue
            if word[1:] in particles:   # 성씨 한 글자 뒤가 조사·시간 표현뿐 → 이름 아님
                continue
            seen.add(word)
            s, e = m.start("word"), m.end("word")
            lo, hi = max(0, s - 15), min(len(line), e + 15)
            context = ("…" if lo > 0 else "") + line[lo:hi].strip() + ("…" if hi < len(line) else "")
            suspects.append(Suspect(f"{n}번째 줄", word, context))
    return suspects


def _alt(keys):
    return "|".join(re.escape(k) for k in sorted(keys, key=lambda k: (-len(k), k)))


def _on(kinds, key):
    return kinds is None or key in kinds


def _build_patterns(dictionary, kinds=None):
    names, guardians, ids, phones = set(), set(), set(), set()
    for s in dictionary.students:
        if _on(kinds, "성명") and len(s.get("성명", "")) >= 2:
            names.add(s["성명"])
        if _on(kinds, "보호자명") and len(s.get("보호자명", "")) >= 2:
            guardians.add(s["보호자명"])
        if _on(kinds, "학번") and s.get("학번"):
            ids.add(s["학번"])
        phone_fields = []
        if _on(kinds, "학생전화"):
            phone_fields.append("학생연락처")
        if _on(kinds, "보호자전화"):
            phone_fields.append("보호자연락처")
        for f in phone_fields:
            d = re.sub(r"\D", "", s.get(f, ""))
            if len(d) >= 9:
                phones.add(d)
                if d.startswith("0"):
                    phones.add(d[1:])   # 엑셀에서 앞 0이 빠진 모양도 잡는다

    # 전화번호: 숫자 사이에 하이픈·점·공백이 있어도 같은 번호로 본다
    phone_alt = "|".join(r"[-. ]?".join(d) for d in sorted(phones, key=lambda d: (-len(d), d)))
    return {
        "names": names,
        "guardians": guardians,
        "ids": ids,
        "phones": phones,
        "name_re": re.compile(_alt(names | guardians)) if names | guardians else None,
        "id_re": re.compile(r"(?<![0-9])(?:" + _alt(ids) + r")(?![0-9])") if ids else None,
        "phone_re": re.compile(r"(?<![0-9])(?:" + phone_alt + r")(?![0-9])") if phones else None,
    }


def _scan(text, p):
    """글 하나에서 남은 개인정보를 (종류, 글자) 목록으로."""
    found = []
    if p["name_re"]:
        for m in p["name_re"].finditer(text):
            w = m.group(0)
            kinds = [k for k, s in (("성명", p["names"]), ("보호자명", p["guardians"])) if w in s]
            found.append(("/".join(kinds), w))
    if p["id_re"]:
        found += [("학번", m.group(0)) for m in p["id_re"].finditer(text)]
    if p["phone_re"]:
        found += [("전화번호", m.group(0)) for m in p["phone_re"].finditer(text)]
    return found


def _as_text(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def check_file(path, dictionary, kinds=None):
    """결과 엑셀을 다시 읽어 남은 개인정보 목록(Finding)을 돌려준다.

    kinds를 주면 일부러 남긴 항목은 남은 개인정보로 세지 않는다.
    """
    p = _build_patterns(dictionary, kinds)
    wb = load_workbook(path)
    findings = []

    def check(sheet, location, text):
        for kind, word in _scan(text, p):
            findings.append(Finding(sheet, location, kind, word))

    for ws in wb.worksheets:
        check(ws.title, "(시트 이름)", ws.title)
        for row in ws.iter_rows():
            for cell in row:
                check(ws.title, cell.coordinate, _as_text(cell.value))
                if cell.comment is not None:
                    check(ws.title, f"{cell.coordinate} (메모)", cell.comment.text or "")
                    check(ws.title, f"{cell.coordinate} (메모 작성자)", cell.comment.author or "")
    # 칸 밖(수식·머리글·유효성·이름 정의·문서 속성·피벗·차트 등)은 파일 속 XML 글자를 통째로 본다 (2026-10-25)
    seen = {(f.kind, f.value) for f in findings}
    for part, text in _xml_texts(path):
        for kind, word in _scan(text, p):
            if (kind, word) not in seen:
                seen.add((kind, word))
                findings.append(Finding("(칸 밖)", part, kind, word))
    return findings


_ATTR = re.compile(r"""([\w:.-]+)\s*=\s*(["'])(.*?)\2""", re.S)   # 여는 따옴표와 같은 따옴표에서만 끝난다
# 칸 주소·행 번호·서식 번호처럼 내용이 아닌 속성 — 「A1001」「r="1001"」을 학번으로 보지 않게
_STRUCT_ATTRS = {
    "r", "ref", "sqref", "spans", "s", "t", "ht", "min", "max", "width", "style", "xfId", "numFmtId",
    "fontId", "fillId", "borderId", "count", "uniqueCount", "id", "r:id", "sheetId", "topLeftCell",
    "activeCell", "activeTab", "tabSelected", "workbookViewId", "pane", "xSplit", "ySplit", "x", "y",
    "outlineLevel", "customHeight", "customWidth", "baseColWidth", "defaultRowHeight", "dyDescent",
    "si", "cm", "vm", "ca", "aca", "dxfId", "priority", "rgb", "theme", "tint", "indexed", "sz", "val",
    "xmlns", "Type", "TargetMode", "mc:Ignorable", "xr:uid", "x14ac:dyDescent", "fld", "item",
}


def _attr_values(tag):
    from urllib.parse import unquote
    # 링크 주소(Target)는 검사하고, %로 감싼 글자도 풀어서 본다 (강희 지적)
    return [unquote(v) if k == "Target" else v for k, _, v in _ATTR.findall(tag) if k not in _STRUCT_ATTRS and not k.startswith("xmlns")]


# 공유문자열 칸 <c t="s"><v>N</v></c> 의 N 은 글이 아니라 sharedStrings 번호 — 학번으로 보지 않게 걷어 낸다 (강희 지적)
# 글자 조각(run)을 감싸는 태그 — 서식만 다른 조각이 한 낱말이면 이어 붙여야 이름이 보인다(「가상」+「인물」, 강희 지적).
#   서식 속성값(lang="ko-KR", sz="1100")도 조각 사이에 끼지 않게 통째로 지운다
_RUN_TAGS = {"t", "r", "rPr", "rFont", "sz", "b", "i", "u", "strike", "color", "family", "scheme", "charset", "outline",
             "shadow", "condense", "extend", "vertAlign", "latin", "ea", "cs", "sym", "solidFill", "srgbClr", "schemeClr"}
_TAG = re.compile(r"<(/?)([\w.-]+:)?([\w.-]+)([^>]*)>")


def _tag_text(m):
    if m.group(3) in _RUN_TAGS: return ""
    return " " + " ".join(_attr_values(m.group(0))) + " "


# 수식·차트 참조식(<f>, <c:f>) 속 칸 주소(Sheet1!$A$1001)는 학번이 아니다 — 큰따옴표 안 글자는 그대로 검사한다 (강희 지적)
#   유효성 목록(<formula1>)·이름 정의(<definedName>)의 범위 참조도 같다 (강희 지적)
_REF_TAGS = {"f", "formula", "formula1", "formula2", "definedName"}
_F_ELEM = re.compile(r"(<(?:[\w.-]+:)?(f|formula|formula1|formula2|definedName)\b[^>]*(?<!/)>)(.*?)(</(?:[\w.-]+:)?\2>)", re.S)
#   행 통째 범위($1001:$1001 — 인쇄 제목 같은 것)도 칸 주소다 (강희 지적)
_CELL_REF = re.compile(r"(?<![\w$.])(?:\$?[A-Za-z]{1,3}\$?\d+(?![\w(])|\$?\d+:\$?\d+(?![\w.]))")


def _drop_refs(formula):
    return "".join(p if p.startswith('"') else _CELL_REF.sub(" ", p) for p in re.split(r'("[^"]*")', formula))


#   빈 칸 <c t="s"/> 은 닫는 태그가 없으니 걸지 않는다 — 걸면 다음 칸의 숫자(진짜 학번)까지 걷힌다
_SST_CELL = re.compile(r"""(<c\b[^>]*\bt\s*=\s*(["'])s\2[^>]*(?<!/)>)(.*?)(</c>)""", re.S)


def _local(name):
    return name.rsplit("}", 1)[-1]


def _parsed_text(data):
    """XML 바이트를 파서로 읽어 검사할 글자로. 못 읽으면 None.
    글자 조각(run)은 이어 붙이고, 그 밖의 요소 사이엔 빈칸. 구조 속성·서식 속성은 빼고, 공유문자열 번호·참조식 칸 주소는 걷는다."""
    import xml.etree.ElementTree as ET
    from urllib.parse import unquote
    try:
        root = ET.fromstring(data)
    except (ET.ParseError, ValueError):
        return None
    out = []

    def walk(el, skip_v):
        tag = _local(el.tag)
        run = tag in _RUN_TAGS
        if not run:
            out.append(" ")
            for k, v in el.attrib.items():
                kl = _local(k)
                if kl in _STRUCT_ATTRS or k in _STRUCT_ATTRS or k.startswith("xmlns"): continue
                out.append(" " + (unquote(v) if kl == "Target" else v) + " ")
        t = el.text or ""
        if tag == "v" and skip_v: t = ""
        elif tag in _REF_TAGS: t = _drop_refs(t)
        out.append(t)
        sst = tag == "c" and el.get("t") == "s"
        for ch in el:
            walk(ch, sst)
            out.append(ch.tail or "")
        if not run: out.append(" ")

    walk(root, False)
    return "".join(out)


def _xml_texts(path):
    """엑셀(zip) 안 XML 마다 (부분 이름, 태그를 걷어 낸 글자)."""
    import html, zipfile
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return []
    out = []
    with z:
        for name in z.namelist():
            if name.endswith((".xml", ".rels", ".vml")):
                data = z.read(name)
                text = _parsed_text(data)   # XML 파서로 — 속성값 안 「>」, UTF-16 파일도 바르게 읽는다 (강희 지적)
                if text is not None:
                    out.append((name, text)); continue
                # 파서가 못 읽는 것(깨진 XML·vml)은 아래 정규식 길로
                raw = data.decode("utf-8", "replace")
                raw = _SST_CELL.sub(lambda m: m.group(1) + re.sub(r"<v>\s*\d+\s*</v>", "", m.group(3)) + m.group(4), raw)
                # 태그는 걷되 속성값은 남긴다 (피벗 캐시 <s v="…"/>, 유효성 prompt="…" 같은 자리)
                raw = _F_ELEM.sub(lambda m: m.group(1) + _drop_refs(m.group(3)) + m.group(4), raw)
                text = re.sub(r"<[^>]*>", " ", _TAG.sub(_tag_text, raw))   # 남은 <?xml …?>·<!-- --> 같은 것은 걷는다
                out.append((name, html.unescape(text)))
    return out


def check_text(text, dictionary, kinds=None):
    """글 전체(여러 줄)에서 남은 개인정보를 찾는다. 위치는 'N번째 줄'."""
    p = _build_patterns(dictionary, kinds)
    findings = []
    for n, line in enumerate(text.splitlines(), start=1):
        for kind, word in _scan(line, p):
            findings.append(Finding("", f"{n}번째 줄", kind, word))
    return findings


def check_md_file(path, dictionary, kinds=None):
    """결과 md(텍스트) 파일을 다시 읽어서 남은 개인정보를 찾는다."""
    with open(path, encoding="utf-8") as f:
        return check_text(f.read(), dictionary, kinds)
