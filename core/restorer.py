"""가명처리된 글(코드)을 실명으로 되돌리는 엔진. 문자열만 다룬다.

    복원글, 변경목록, 경고목록 = unmask_text(text, dictionary)

변경목록의 각 항목: [코드 원문, 복원한 값, 종류, 위치]  (위치: 글자 번호, 0부터)
경고목록: 사람이 읽을 문장 (복원 못 한 것, 확인이 필요한 것)

- 학생003 → 성명, 보호자003 → 보호자명, S003 → 학번
- [학생전화003] [보호자전화003] [주소003] [생년월일003] → 사전의 원래 값
- AI가 바꿔 쓴 모양(학생 003, 학생3, 학생03, 학생003번, S 003, 학생전화003 ...)도 인식
- 이름 바로 뒤 조사는 받침에 맞게 고친다 (학생003가 → 홍길동이)
- 복원할 수 없는 것은 그대로 두고 경고한다
"""
import re

NB = r"(?<![가-힣A-Za-z0-9])"   # 앞에 글자·숫자가 붙어 있으면 코드로 보지 않음 (대학생3 등)
# 조사 뒤는 공백·문장부호·줄끝일 때만 조사로 본다 (학생003가방의 '가'는 조사가 아님)
END = r"(?=[\s.,!?;:()\[\]{}'\"“”‘’「」『』…·~*_/-]|$)"
PARTICLE = r"(?:이라고|라고|이랑|랑|으로|로|은|는|이|가|을|를|과|와|아|야)"
# '학생 3명', '학생 2학년' 같은 일반 문장은 코드로 보지 않는다
UNITS = r"(?:번째|명|개|반|학년|학기|분|가지|곳|등|위|시|일|월|주|차|회|점|장|권|%|교시|시간|살|세|년|원|건|배|편|쪽|층)"

_PATTERN = re.compile("|".join([
    # [학생전화003] [주소 003] [학생전화003또는017]
    r"(?P<bracket>\[\s*(?P<bk>학생전화|보호자전화|주소|생년월일)\s*"
    r"(?P<bn>\d{1,4}(?:\s*또는\s*\d{1,4})*)\s*\])",
    # 원래 값이 사전에 없는 표시
    r"(?P<unknown>\[(?P<un>전화번호|주민번호|이메일)\])",
    # '이 단어도 가리기'로 바꾼 이름 ([이름1], [이름2] ...) - 변경기록으로 복원
    r"(?P<ename>\[이름\d{1,4}\])",
    # 괄호 없는 학생전화003 / 보호자전화003
    rf"(?P<tel>{NB}(?P<pk>학생전화|보호자전화)\s?(?P<pn>\d{{1,4}}(?:또는\d{{1,4}})*)(?!\d))",
    # 학생003 / 학생 3 / 학생03번 / 학생003또는017 + 조사
    rf"(?P<name>{NB}(?P<nk>학생|보호자)\s?"
    rf"(?P<nn>\d{{1,4}}(?:또는(?:(?:학생|보호자)\s?)?\d{{1,4}})*)(?!\d|{UNITS})"
    rf"(?:번(?!째))?(?P<part>{PARTICLE}{END})?)",
    # S003 / S 003 (S3처럼 짧은 것은 다른 말과 헷갈려서 인식하지 않음)
    rf"(?P<sid>{NB}S\s?(?P<sn>\d{{3,4}}(?:또는\d{{3,4}})*)(?!\d))",
    # 교사001 / 교사 001 / 교사001또는002 (+ 조사)
    rf"(?P<staff>{NB}교사\s?(?P<tn>\d{{1,4}}(?:또는\d{{1,4}})*)(?!\d|{UNITS})"
    rf"(?:번(?!째))?(?P<tpart>{PARTICLE}{END})?)",
]))

_VALUE_FIELD = {
    "학생전화": ("학생연락처", "학생전화"),
    "보호자전화": ("보호자연락처", "보호자전화"),
    "주소": ("주소", "주소"),
    "생년월일": ("생년월일", "생년월일"),
}
_PARTICLE_PAIRS = [("이", "가"), ("은", "는"), ("을", "를"), ("과", "와"),
                   ("아", "야"), ("이랑", "랑"), ("이라고", "라고")]
DUP_MARK = "(동명이인)"
_FIELD_SAY = {
    "성명": "이름",
    "보호자명": "보호자 이름",
    "학번": "학번",
    "학생연락처": "학생 연락처",
    "보호자연락처": "보호자 연락처",
    "주소": "주소",
    "생년월일": "생년월일",
}


def _say_field(field):
    return _FIELD_SAY.get(field, field)


def fix_particle(name, particle):
    """name 마지막 글자의 받침에 맞는 조사로 고친다. (한글이 아니면 그대로)"""
    last = name[-1]
    if not "가" <= last <= "힣":
        return particle
    jong = (ord(last) - 0xAC00) % 28
    if particle in ("으로", "로"):
        return "로" if jong in (0, 8) else "으로"   # 8 = ㄹ 받침
    for with_batchim, without in _PARTICLE_PAIRS:
        if particle in (with_batchim, without):
            return with_batchim if jong else without
    return particle


def _index(dictionary):
    """번호(정수) → 학생. 한 번 만들어 사전 객체에 붙여 둔다."""
    idx = getattr(dictionary, "_restore_index", None)
    if idx is None:
        idx = {}
        for s in dictionary.students:
            if s.get("번호", "").isdigit():
                idx.setdefault(int(s["번호"]), s)
        dictionary._restore_index = idx
    return idx


def _staff_index(dictionary):
    """번호(정수) → 교직원. 한 번 만들어 사전 객체에 붙여 둔다."""
    idx = getattr(dictionary, "_restore_staff_index", None)
    if idx is None:
        idx = {}
        for s in getattr(dictionary, "staff", []):
            if str(s.get("번호", "")).isdigit():
                idx.setdefault(int(s["번호"]), s)
        dictionary._restore_staff_index = idx
    return idx


def _unique(seq):
    return list(dict.fromkeys(seq))


def unmask_text(text, dictionary, extra=None):
    """text 안의 코드를 실명·원래 값으로 되돌린다. (복원글, 변경목록, 경고목록)

    extra: '이 단어도 가리기'로 바꾼 {코드([이름1] 등): 원래 단어} 매핑. 변경기록에서 가져온다.
    """
    if not text:
        return text, [], []
    by_num = _index(dictionary)
    by_staff = _staff_index(dictionary)
    extra = extra or {}
    changes, warnings = [], []
    multiline = "\n" in text

    def warn(pos, msg):
        where = f"{text.count(chr(10), 0, pos) + 1}번째 줄, " if multiline else ""
        warnings.append(where + msg)

    def lookup(items, field, pos, original, strict):
        """items=[(접두어, 번호)] → 값 목록. 못 찾으면 None (strict면 경고)."""
        values = []
        for prefix, n in items:
            s = by_num.get(n)
            if s is None:
                if strict:
                    warn(pos, f"{original}: 명단에 없는 번호예요. 그대로 두었어요")
                return None
            f = field(prefix)
            if not s.get(f):
                warn(pos, f"{original}: 명단에 {_say_field(f)} 값이 없어 되돌리지 못하고 그대로 두었어요")
                return None
            values.append(s[f])
        return values

    def finish(original, pos, value, kind):
        changes.append([original, value, kind, pos])
        return value

    def restore_values(original, pos, field_name, kind, numbers, strict=True):
        """이름이 아닌 값(학번, 전화, 주소, 생년월일)."""
        values = lookup([(None, n) for n in numbers], lambda _: field_name, pos, original, strict)
        if values is None:
            return original
        uniq = _unique(values)
        if len(uniq) == 1:
            return finish(original, pos, uniq[0], kind)
        out = " 또는 ".join(uniq) + "(확인해 주세요)"
        warn(pos, f"{original}: 후보가 여러 개라 하나로 정하지 못했어요 → {out}")
        return finish(original, pos, out, "중복확인")

    def repl(m):
        original, pos = m.group(0), m.start()

        if m.group("unknown") is not None:
            warn(pos, f"{original}: 명단에 없는 글이라 되돌리지 못하고 그대로 두었어요")
            return original

        if m.group("ename") is not None:
            word = extra.get(original)
            if word:
                return finish(original, pos, word, "이름의심")
            warn(pos, f"{original}: 이 파일의 기록이 없어 되돌리지 못하고 그대로 두었어요")
            return original

        if m.group("bracket") is not None:
            field_name, kind = _VALUE_FIELD[m.group("bk")]
            nums = [int(x) for x in re.findall(r"\d+", m.group("bn"))]
            return restore_values(original, pos, field_name, kind, nums)

        if m.group("tel") is not None:
            field_name, kind = _VALUE_FIELD[m.group("pk")]
            nums = [int(x) for x in m.group("pn").split("또는")]
            return restore_values(original, pos, field_name, kind, nums)

        if m.group("sid") is not None:
            nums = [int(x) for x in m.group("sn").split("또는")]
            return restore_values(original, pos, "학번", "학번", nums)

        if m.group("staff") is not None:
            nums = [int(x) for x in m.group("tn").split("또는")]
            part = m.group("tpart")
            code = original[:-len(part)] if part else original
            values = []
            for n in nums:
                s = by_staff.get(n)
                if s is None:
                    warn(pos, f"{code}: 명단에 없는 선생님 번호예요. 그대로 두었어요")
                    return original
                if not s.get("성명"):
                    warn(pos, f"{code}: 명단에 이름이 없어 되돌리지 못하고 그대로 두었어요")
                    return original
                values.append(s["성명"])
            uniq = _unique(values)
            if len(uniq) == 1:
                out = uniq[0] + (fix_particle(uniq[0], part) if part else "")
                return finish(original, pos, out, "교직원")
            out = "/".join(uniq) + DUP_MARK + (part or "")
            warn(pos, f"{code}: 같은 이름이 여러 명이라 한 사람으로 정하지 못했어요 → {out}")
            return finish(original, pos, out, "동명이인확인")

        # 학생003 / 보호자003 (이름)
        cur, items, first_digits = m.group("nk"), [], None
        for i, p in enumerate(m.group("nn").split("또는")):
            mm = re.fullmatch(r"(?:(학생|보호자)\s?)?(\d+)", p)
            if mm.group(1):
                cur = mm.group(1)
            items.append((cur, int(mm.group(2))))
            if i == 0:
                first_digits = mm.group(2)
        strict = len(first_digits) >= 3   # 학생 3 / 학생03 처럼 짧은 형태는 사전에 없으면 조용히 둔다
        part = m.group("part")
        code = original[:-len(part)] if part else original   # 경고에는 조사를 뺀 코드만 적는다
        values = lookup(items, lambda pre: "성명" if pre == "학생" else "보호자명", pos, code, strict)
        if values is None:
            return original


        if len(items) == 1:
            kind = "성명" if items[0][0] == "학생" else "보호자명"
            out = values[0]
            if part:
                out += fix_particle(values[0], part)
            elif text[m.end():m.end() + 2] == "학생":      # 학생003학생 → 홍길동 학생
                out += " "
            return finish(original, pos, out, kind)

        uniq = _unique(values)
        if all(p == "보호자" for p, _ in items) and len(uniq) == 1:
            out = uniq[0]                                  # 형제의 같은 보호자: 한 사람
            if part:
                out += fix_particle(out, part)
            return finish(original, pos, out, "보호자명")
        out = "/".join(uniq) + DUP_MARK
        warn(pos, f"{code}: 같은 이름이 여러 명이라 한 사람으로 정하지 못했어요 → {out}")
        if part:
            out += part
        return finish(original, pos, out, "동명이인확인")

    return _PATTERN.sub(repl, text), changes, warnings
