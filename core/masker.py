"""글자 치환 엔진 (문자열만 다룸. 파일/화면 처리는 하지 않는다).

사용 예:
    d = load_dictionary("data/학생사전.xlsx")
    새글, 변경목록 = mask_text("홍길동 학생 010-1234-5678", d)

변경목록의 각 항목: [원래 값, 바뀐 값, 종류, 위치]
  - 위치: 원래 글에서 해당 부분이 시작하는 글자 번호(0부터)
  - 종류: 성명 / 보호자명 / 성뺀이름 / 학번 / 학생전화 / 보호자전화 / 주소 /
          생년월일 / 동명이인확인 / 중복확인 / 사전에없음 / 주민번호 / 이메일

모든 치환은 컴파일된 정규식 한 번(re.sub)으로 처리한다.
- 같은 자리에서는 긴 문자열이 먼저 맞도록 정렬해서 넣는다.
- 한 번 훑으면서 원문만 보고 바꾸므로, 이미 바뀐 코드(학생003 등)는 다시 바뀌지 않는다.
"""
import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from core import config

BASE_DIR = config.user_data_dir()

# 이름 바로 뒤에 올 수 있는 조사 (성 뺀 이름은 이런 조사가 붙을 때만 바꾼다)
_PARTICLES = [
    "에게서", "에게", "한테서", "한테", "께서", "께", "이라고", "이랑", "이라는",
    "이란", "라고", "이나", "으로", "처럼", "같이", "보다", "부터", "까지",
    "조차", "마저", "보고", "이가", "은", "는", "이", "가", "을", "를", "도",
    "의", "와", "과", "랑", "야", "아", "만",
]
_PARTICLE_END = (
    r"(?:" + "|".join(_PARTICLES) + r")(?:은|는|도|만|의|까지|부터)?(?![가-힣])"
)

_EMAIL = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}"
_RRN = r"(?<!\d)\d{6}-\d{7}(?!\d)"
_PHONE = r"(?<!\d)0(?:2|[1-8]\d)[-. ]?\d{3,4}[-. ]?\d{4}(?!\d)"
_BIRTH = r"(?<!\d)(?P<by>\d{4})(?P<bs>[-./]?)(?P<bm>\d{2})(?P=bs)(?P<bd>\d{2})(?!\d)"

_NAME_KINDS = {"성명", "보호자명", "성뺀이름", "교직원"}

# 화면에서 고르는 항목. (엔진 키, 보이는 이름). 성 뺀 이름은 '학생 이름'에 포함된다.
MASK_OPTIONS = (
    ("성명", "학생 이름"),
    ("학번", "학번"),
    ("학생전화", "학생 연락처"),
    ("보호자명", "보호자 이름"),
    ("보호자전화", "보호자 연락처"),
    ("주소", "주소"),
    ("생년월일", "생년월일"),
    ("교직원", "교직원"),
    ("이메일", "이메일"),
    ("주민번호", "주민번호"),
    ("사전에없음", "다른 전화번호"),
)
_KIND_OPTION = {
    "성명": "성명",
    "성뺀이름": "성명",
    "보호자명": "보호자명",
    "학번": "학번",
    "학생전화": "학생전화",
    "보호자전화": "보호자전화",
    "주소": "주소",
    "생년월일": "생년월일",
    "교직원": "교직원",
    "이메일": "이메일",
    "주민번호": "주민번호",
    "사전에없음": "사전에없음",
}


def kind_enabled(kinds, key):
    """kinds가 None이면 모든 항목을 가린다. 아니면 고른 키만."""
    return kinds is None or key in kinds


def _hangul_span(text, start, end):
    """맞춰진 구간을 붙여 쓴 한글 단어까지 넓힌다. '연락처' 안의 짧은 일치도 제목으로 보게."""
    while start > 0 and "가" <= text[start - 1] <= "힣":
        start -= 1
    n = len(text)
    while end < n and "가" <= text[end] <= "힣":
        end += 1
    return text[start:end]


def _is_title_span(text, start, end):
    return config.is_field_title(_hangul_span(text, start, end))

# 교직원 이름 뒤에 붙는 호칭 (이 호칭이 붙으면 '성 뺀 이름'도 바꾼다)
_TEACHER_TITLE = r"(?:선생님|선생|쌤|샘)"


@dataclass(frozen=True)
class _Entry:
    code: str   # 바꿀 코드 (예: 학생003)
    num: str    # 번호 (예: 003)
    sid: str    # 이 항목이 속한 학생의 학번
    kind: str   # 변경 종류


class Dictionary:
    """학생사전. load_dictionary()로 만든다."""

    def __init__(self, students, staff=None):
        self.students = students
        self.staff = staff or []
        self._lit = {}          # 성명·보호자명·주소·교직원 성명 → [_Entry]
        self._given = {}        # 학생 성 뺀 이름 → [_Entry]
        self._staff_given = {}  # 교직원 성 뺀 이름(호칭이 붙을 때만) → [_Entry]
        self._ids = {}          # 학번 → [_Entry]
        self._phones = {}       # 숫자만 남긴 전화번호 → [_Entry]
        self._births = {}       # YYYYMMDD → [_Entry]
        self._build_index()
        self._pattern = self._compile()

    # ------------------------------------------------------------ 색인
    @staticmethod
    def _add(index, key, s, code_field, kind):
        code = s.get(code_field, "")
        if key and code:
            index.setdefault(key, []).append(_Entry(code, s.get("번호", ""), s.get("학번", ""), kind))

    def _build_index(self):
        for s in self.students:
            name = s.get("성명", "")
            if len(name) >= 2:
                self._add(self._lit, name, s, "학생코드", "성명")
            guardian = s.get("보호자명", "")
            if len(guardian) >= 2:
                self._add(self._lit, guardian, s, "보호자코드", "보호자명")
            self._add(self._lit, s.get("주소", ""), s, "주소코드", "주소")

            # 외자 이름·일반 단어 겹침 학생은 성 뺀 이름을 바꾸지 않는다
            caution = s.get("주의", "")
            given = s.get("성뺀이름", "")
            if len(given) >= 2 and "외자" not in caution and "일반 단어" not in caution:
                self._add(self._given, given, s, "학생코드", "성뺀이름")

            self._add(self._ids, s.get("학번", ""), s, "S코드", "학번")
            self._add(self._phones, re.sub(r"\D", "", s.get("학생연락처", "")), s, "학생전화코드", "학생전화")
            self._add(self._phones, re.sub(r"\D", "", s.get("보호자연락처", "")), s, "보호자전화코드", "보호자전화")
            birth = re.sub(r"\D", "", s.get("생년월일", ""))
            if len(birth) == 8:
                self._add(self._births, birth, s, "생년월일코드", "생년월일")

        # 교직원: 성명은 어디서나, 성 뺀 이름은 '선생님/쌤' 같은 호칭이 붙을 때만 바꾼다
        for t in self.staff:
            name = t.get("성명", "")
            if len(name) >= 2:
                self._add(self._lit, name, t, "교사코드", "교직원")
            given = t.get("성뺀이름", "")
            if len(given) >= 1 and "일반 단어" not in t.get("주의", ""):
                self._add(self._staff_given, given, t, "교사코드", "교직원")

    @staticmethod
    def _alt(keys):
        """긴 문자열이 먼저 맞도록 정렬한 정규식 선택지."""
        return "|".join(re.escape(k) for k in sorted(keys, key=lambda k: (-len(k), k)))

    def _compile(self):
        parts = [
            f"(?P<email>{_EMAIL})",
            f"(?P<rrn>{_RRN})",
            f"(?P<phone>{_PHONE})",
            f"(?P<birth>{_BIRTH})",
        ]
        if self._lit:
            parts.append(f"(?P<lit>{self._alt(self._lit)})")
        if self._given:
            parts.append(
                f"(?P<given>(?<![가-힣])(?P<gn>{self._alt(self._given)})"
                f"(?:이?[ \\t]*학생|(?:이(?={_PARTICLE_END}))?(?={_PARTICLE_END})))"
            )
        if self._staff_given:
            parts.append(
                f"(?P<sgiven>(?<![가-힣])(?P<sgn>{self._alt(self._staff_given)})"
                f"(?=[ \\t]*{_TEACHER_TITLE}))"
            )
        if self._ids:
            parts.append(f"(?P<sid>(?<![0-9])(?:{self._alt(self._ids)})(?![0-9]))")
        return re.compile("|".join(parts))

    # ------------------------------------------------------------ 치환 도우미
    @staticmethod
    def _combine(cands):
        """후보가 여럿일 때 '학생003또는017' 꼴로 합친다."""
        cands = sorted(cands, key=lambda e: (e.num, e.code))
        pre, suf = set(), set()
        for e in cands:
            i = e.code.rfind(e.num) if e.num else -1
            if i < 0:
                return "또는".join(e.code for e in cands)
            pre.add(e.code[:i])
            suf.add(e.code[i + len(e.num):])
        if len(pre) == 1 and len(suf) == 1:
            return pre.pop() + "또는".join(e.num for e in cands) + suf.pop()
        return "또는".join(e.code for e in cands)

    def _resolve(self, entries, text, pos):
        """(바뀐 값, 종류). 후보가 여럿이면 같은 줄의 학번으로 좁히고, 못 좁히면 합친다."""
        uniq = {}
        for e in entries:
            uniq.setdefault(e.code, e)
        cands = list(uniq.values())
        if len(cands) > 1:
            start = text.rfind("\n", 0, pos) + 1
            end = text.find("\n", pos)
            line = text[start:] if end < 0 else text[start:end]
            hit = [
                e for e in cands
                if e.sid and re.search(r"(?<![0-9])" + re.escape(e.sid) + r"(?![0-9])", line)
            ]
            if hit:
                cands = hit
        if len(cands) == 1:
            return cands[0].code, cands[0].kind
        kind = "동명이인확인" if all(e.kind in _NAME_KINDS for e in cands) else "중복확인"
        return self._combine(cands), kind

    def _keep(self, entries, kinds):
        """고른 항목에 해당하는 후보만. kinds가 None이면 전부."""
        if not entries:
            return []
        if kinds is None:
            return list(entries)
        return [e for e in entries if kind_enabled(kinds, _KIND_OPTION.get(e.kind, e.kind))]

    # ------------------------------------------------------------ 치환
    def mask(self, text, kinds=None):
        """kinds가 None이면 전부 가린다. 집합이면 그 항목만 가리고 나머지는 원문 그대로 둔다."""
        changes = []

        def repl(m):
            original = m.group(0)
            pos = m.start()
            if _is_title_span(text, pos, pos + len(original)):
                return original   # 학번·이름·연락처 같은 제목은 바꾸지 않는다
            value = kind = None

            if m.group("email") is not None:
                if kind_enabled(kinds, "이메일"):
                    value, kind = "[이메일]", "이메일"
            elif m.group("rrn") is not None:
                if kind_enabled(kinds, "주민번호"):
                    value, kind = "[주민번호]", "주민번호"
            elif m.group("phone") is not None:
                raw = self._phones.get(re.sub(r"\D", "", original))
                entries = self._keep(raw, kinds)
                if entries:
                    value, kind = self._resolve(entries, text, pos)
                elif not raw and kind_enabled(kinds, "사전에없음"):
                    value, kind = "[전화번호]", "사전에없음"
            elif m.group("birth") is not None:
                entries = self._keep(self._births.get(m.group("by") + m.group("bm") + m.group("bd")), kinds)
                if entries:  # 사전에 없는 날짜는 그대로 둔다
                    value, kind = self._resolve(entries, text, pos)
            elif "lit" in m.re.groupindex and m.group("lit") is not None:
                entries = self._keep(self._lit[original], kinds)
                if entries:
                    value, kind = self._resolve(entries, text, pos)
            elif "given" in m.re.groupindex and m.group("given") is not None:
                entries = self._keep(self._given[m.group("gn")], kinds)
                if entries:
                    value, kind = self._resolve(entries, text, pos)
            elif "sgiven" in m.re.groupindex and m.group("sgiven") is not None:
                entries = self._keep(self._staff_given[m.group("sgn")], kinds)
                if entries:
                    value, kind = self._resolve(entries, text, pos)
            elif "sid" in m.re.groupindex and m.group("sid") is not None:
                entries = self._keep(self._ids[original], kinds)
                if entries:
                    value, kind = self._resolve(entries, text, pos)

            if value is None:
                return original
            changes.append([original, value, kind, pos])
            return value

        return self._pattern.sub(repl, text), changes

    def present_kinds(self, text):
        """이 글에서 실제로 가릴 수 있는 항목 키. 없는 항목은 넣지 않는다."""
        found = set()
        if not text:
            return found

        def take(entries):
            for entry in entries or []:
                key = _KIND_OPTION.get(entry.kind)
                if key:
                    found.add(key)

        def repl(m):
            original = m.group(0)
            if _is_title_span(text, m.start(), m.end()):
                return original
            if m.group("email") is not None:
                found.add("이메일")
            elif m.group("rrn") is not None:
                found.add("주민번호")
            elif m.group("phone") is not None:
                raw = self._phones.get(re.sub(r"\D", "", original))
                if raw:
                    take(raw)
                else:
                    found.add("사전에없음")
            elif m.group("birth") is not None:
                take(self._births.get(m.group("by") + m.group("bm") + m.group("bd")))
            elif "lit" in m.re.groupindex and m.group("lit") is not None:
                take(self._lit[original])
            elif "given" in m.re.groupindex and m.group("given") is not None:
                take(self._given[m.group("gn")])
            elif "sgiven" in m.re.groupindex and m.group("sgiven") is not None:
                take(self._staff_given[m.group("sgn")])
            elif "sid" in m.re.groupindex and m.group("sid") is not None:
                take(self._ids[original])
            return original

        self._pattern.sub(repl, text)
        return found


def load_dictionary(path=None):
    """build_dict.py가 만든 학생사전 엑셀을 읽는다. path 를 생략하면 config.dict_path()."""
    p = Path(path).expanduser() if path else config.dict_path()
    if not p.is_absolute() and not p.exists():
        p = BASE_DIR / p
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        if "학생사전" in wb.sheetnames:
            ws = wb["학생사전"]
        else:
            ws = next((s for s in wb.worksheets if s.title != "사전정보"), wb.worksheets[0])
        rows = list(ws.iter_rows(values_only=True))
        staff_rows = list(wb["교직원"].iter_rows(values_only=True)) if "교직원" in wb.sheetnames else None
    finally:
        wb.close()
    if not rows:
        raise ValueError(f"명단이 비어 있어요: {p}")

    students = _rows_to_dicts(rows)
    staff = _rows_to_dicts(staff_rows) if staff_rows else []
    return Dictionary(students, staff)


def _rows_to_dicts(rows):
    """[머리글행, 값행...] → 성명이 있는 행만 dict 목록으로."""
    if not rows:
        return []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    out = []
    for row in rows[1:]:
        s = {h: ("" if v is None else str(v).strip()) for h, v in zip(header, row) if h}
        if s.get("성명"):
            out.append(s)
    return out


def mask_text(text, dictionary, kinds=None):
    """text를 치환해서 (바뀐 글자, 변경 목록)을 돌려준다.

    kinds가 None이면 전부 가린다. 아니면 MASK_OPTIONS의 키 중에서 고른 것만 가린다.
    """
    if not text:
        return text, []
    return dictionary.mask(text, kinds)
