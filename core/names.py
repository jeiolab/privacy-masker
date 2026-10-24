"""사람 이름처럼 보이는 단어를 칸·글에서 찾는다.

사전 열(성명/학번) 밖에 있는 전출·전입·메모 칸도 같은 규칙으로 본다.
학교 이름·머리글·일반 단어는 제외한다.
"""
import re

from core import config

_SCHOOL_TAILS = (
    "중학교", "고등학교", "초등학교", "여자중학교", "여자고등학교",
    "여자중", "여자고", "여중", "여고", "분교장", "분교", "학교",
)
_SKIP_LABELS = {
    "전출", "전입", "자퇴", "편입", "출신중", "학번", "성명", "이름", "학생명",
    "학생이름", "전교생", "총계", "남", "여", "계", "보호자", "보호자명",
}
_SPLIT = re.compile(r"[,，/·;|\n]+")
_TRAIL_ASCII = re.compile(r"[A-Za-z0-9._]+$")
_PARENS = re.compile(r"\([^)]*\)")
_HANGUL_ONLY = re.compile(r"^[가-힣]+$")

_SUR = sorted(list(config.COMPOUND_SURNAMES) + list(config.KOREAN_SURNAMES),
              key=lambda s: (-len(s), s))
_NAME_TOKEN = re.compile(
    r"(?<![가-힣])(?:" + "|".join(re.escape(s) for s in _SUR) + r")[가-힣]{1,3}(?![가-힣])"
)


def normalize_name_token(text):
    """칸 값에서 이름만 남긴다. 예: '박선영(3.1.)' → '박선영', '조은호B' → '조은호'."""
    s = re.sub(r"\s+", "", str(text or ""))
    s = _PARENS.sub("", s)
    s = _TRAIL_ASCII.sub("", s)
    return s


def looks_like_person_name(word):
    """성씨로 시작하는 2~4글자 한글이고, 학교명·예외 단어가 아니면 True."""
    word = normalize_name_token(word)
    if not _HANGUL_ONLY.match(word) or not (2 <= len(word) <= 4):
        return False
    if word in _SKIP_LABELS or word in config.NAME_SUSPECT_EXCEPTIONS:
        return False
    if word in config.COMMON_WORD_NAMES:
        return False
    if any(word.endswith(t) for t in _SCHOOL_TAILS):
        return False
    if word.endswith("중") and len(word) >= 3:
        return False
    sur = next((s for s in _SUR if word.startswith(s) and len(word) > len(s)), None)
    if not sur:
        return False
    given = word[len(sur):]
    if given in config.NAME_SUSPECT_PARTICLES:
        return False
    return True


def iter_person_names(text):
    """글 안에 있는 사람 이름 후보를 나온 순서로 (중복 없이)."""
    if text is None:
        return []
    raw = str(text)
    found, seen = [], set()

    def add(token):
        name = normalize_name_token(token)
        if name and looks_like_person_name(name) and name not in seen:
            seen.add(name)
            found.append(name)

    for part in _SPLIT.split(raw):
        add(part.strip())
        for m in _NAME_TOKEN.finditer(part):
            add(m.group(0))
    return found


def leftover_person_names(text, dictionary, extra_ignore=()):
    """가명처리 뒤에 남은, 사전에 없는 사람 이름 목록."""
    known = set()
    for s in getattr(dictionary, "students", []) or []:
        known.add(s.get("성명", ""))
        known.add(s.get("보호자명", ""))
    for t in getattr(dictionary, "staff", []) or []:
        known.add(t.get("성명", ""))
    known.discard("")
    known.update(extra_ignore)
    out, seen = [], set()
    for name in iter_person_names(text):
        if name in known or name in seen:
            continue
        seen.add(name)
        out.append(name)
    return out
