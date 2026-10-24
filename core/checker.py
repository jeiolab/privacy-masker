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


def _build_patterns(dictionary):
    names, guardians, ids, phones = set(), set(), set(), set()
    for s in dictionary.students:
        if len(s.get("성명", "")) >= 2:
            names.add(s["성명"])
        if len(s.get("보호자명", "")) >= 2:
            guardians.add(s["보호자명"])
        if s.get("학번"):
            ids.add(s["학번"])
        for f in ("학생연락처", "보호자연락처"):
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


def check_file(path, dictionary):
    """결과 엑셀을 다시 읽어 남은 개인정보 목록(Finding)을 돌려준다."""
    p = _build_patterns(dictionary)
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
    return findings


def check_text(text, dictionary):
    """글 전체(여러 줄)에서 남은 개인정보를 찾는다. 위치는 'N번째 줄'."""
    p = _build_patterns(dictionary)
    findings = []
    for n, line in enumerate(text.splitlines(), start=1):
        for kind, word in _scan(line, p):
            findings.append(Finding("", f"{n}번째 줄", kind, word))
    return findings


def check_md_file(path, dictionary):
    """결과 md(텍스트) 파일을 다시 읽어서 남은 개인정보를 찾는다."""
    with open(path, encoding="utf-8") as f:
        return check_text(f.read(), dictionary)
