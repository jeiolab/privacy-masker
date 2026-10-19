#!/usr/bin/env python3
"""글자 치환 엔진 확인용 스크립트.

사전(data/학생사전.xlsx)에서 실제 이름·학번·전화번호를 꺼내 테스트 문장을 만들고,
'바꾸기 전 → 바꾼 후'와 변경 목록을 출력한다.   실행: python3 test_masker.py

build_cases() 는 test_restore.py (왕복 테스트)에서도 같은 문장을 쓰려고 따로 뺐다.
"""
import re
import sys
from collections import Counter

from core.masker import load_dictionary, mask_text

failures = 0


def josa(name, with_batchim, without):
    """이름 마지막 글자에 받침이 있으면 with_batchim, 없으면 without 조사."""
    return with_batchim if (ord(name[-1]) - 0xAC00) % 28 else without


def build_cases(d):
    """테스트 문장 목록. 각 항목은 dict:
    title, text, absent(결과에 없어야 할 글자), present(있어야 할 글자), counts((글자, 개수)),
    roundtrip(True면 mask→unmask 했을 때 원문과 같아야 함, 아니면 제외 사유 글),
    restore(왕복 대신 복원 결과에 대한 기대: {'count': [(글자, 개수)], 'warnings': 개수})
    """
    S = d.students
    name_cnt = Counter(s["성명"] for s in S)
    given_cnt = Counter(s["성뺀이름"] for s in S)
    guardian_cnt = Counter(s["보호자명"] for s in S if s.get("보호자명"))
    phone_cnt = Counter(re.sub(r"\D", "", s.get("학생연락처", "")) for s in S)
    all_digits = {re.sub(r"\D", "", s.get(f, "")) for s in S for f in ("학생연락처", "보호자연락처")}
    plain = [s for s in S if not s.get("주의") and name_cnt[s["성명"]] == 1 and given_cnt[s["성뺀이름"]] == 1]
    cases = []

    def add(title, text, absent=(), present=(), counts=(), roundtrip=True, restore=None):
        cases.append(dict(title=title, text=text, absent=list(absent), present=list(present),
                          counts=list(counts), roundtrip=roundtrip, restore=restore or {}))

    # 1) 성명 전체
    s = plain[0]
    n = s["성명"]
    add("1) 성명 전체", f"{n}{josa(n, '은', '는')} 오늘 학급 회의에서 발표를 맡았다.",
        absent=[n], present=[s["학생코드"]])

    # 2) 성 뺀 이름 + 조사
    s = plain[1]
    g = s["성뺀이름"]
    add("2) 성 뺀 이름 + 조사",
        f"{g}이가 발표를 잘했다. {g}이는 책을 읽었다. {g}아, 이리 오렴. "
        f"{g}은 성실하다. {g}이를 칭찬했다. {g} 학생을 격려했다.",
        absent=[g], counts=[(s["학생코드"], 6)],
        roundtrip="성 뺀 이름은 성명 전체로 복원됨",
        restore={"count": [(s["성명"], 6)]})

    # 3) 일반 단어와 겹치는 이름
    s = next((x for x in S if x["성뺀이름"] == "하늘" and "일반 단어" in x.get("주의", "")), None) \
        or next((x for x in S if "일반 단어" in x.get("주의", "")), None)
    if s:
        g = s["성뺀이름"]
        add("3) 일반 단어와 겹치는 이름", f"{s['성명']} 학생이 발표했다. 오늘 {g}이 맑다.",
            absent=[s["성명"]], present=[s["학생코드"], f"오늘 {g}이 맑다"])

    # 4) 외자 이름
    s = next((x for x in S if "외자" in x.get("주의", "")), None)
    if s:
        g = s["성뺀이름"]
        add("4) 외자 이름", f"{s['성명']} 학생이 {g}비물을 챙겼다.",
            absent=[s["성명"]], present=[s["학생코드"], f"{g}비물을 챙겼다"])

    # 5) 동명이인
    dup = next((nm for nm, c in name_cnt.items() if c > 1), None)
    if dup:
        a, b = sorted((x for x in S if x["성명"] == dup), key=lambda x: x["번호"])[:2]
        add("5-1) 동명이인 + 같은 줄에 학번 있음", f"{dup}({a['학번']}) 학생이 발표했다.",
            absent=[dup, a["학번"]], present=[a["학생코드"], a["S코드"]])
        add("5-2) 동명이인 + 학번 없음", f"{dup}{josa(dup, '이', '가')} 복도에서 뛰었다.",
            absent=[dup], present=[f"{a['학생코드']}또는{b['번호']}"],
            roundtrip="누구인지 정할 수 없음 (동명이인-확인필요)",
            restore={"contains": [f"{dup}(동명이인-확인필요)"], "warnings": 1})

    # 6) 학번이 더 긴 숫자 안에 들어간 경우
    s = plain[2]
    sid = s["학번"]
    add("6) 학번이 더 긴 숫자 안에 있음", f"문서번호 1{sid}2 는 그대로, 학번 {sid} 은 바꿈.",
        present=[f"1{sid}2", s["S코드"]], counts=[(s["S코드"], 1)])

    # 7) 같은 전화번호 4가지 형식
    s = next(x for x in plain if x.get("학생연락처") and phone_cnt[re.sub(r"\D", "", x["학생연락처"])] == 1)
    dg = re.sub(r"\D", "", s["학생연락처"])
    a, b, c = dg[:3], dg[3:7], dg[7:]
    forms = [dg, f"{a}-{b}-{c}", f"{a} {b} {c}", f"{a}.{b}.{c}"]
    add("7) 같은 전화번호 4가지 형식", "연락처: " + " / ".join(forms),
        absent=[b + c, f"{b}-{c}"], counts=[(s["학생전화코드"], 4)],
        roundtrip="전화번호가 모두 하이픈 형식으로 통일됨",
        restore={"count": [(s["학생연락처"], 4)]})

    # 8) 사전에 없는 전화번호
    unk = next(f"010-{k:04d}-{k:04d}" for k in range(1000, 9999) if f"010{k:04d}{k:04d}" not in all_digits)
    add("8) 사전에 없는 전화번호", f"급하면 {unk} 또는 02-123-4567 로 연락 바람.",
        absent=[unk, "123-4567"], counts=[("[전화번호]", 2)],
        roundtrip="원래 값이 사전에 없어 복원 불가",
        restore={"count": [("[전화번호]", 2)], "warnings": 2})

    # 9) 주민번호, 이메일 섞임
    add("9) 주민번호 + 이메일", "주민번호 900101-1234567, 메일 hong.gildong@example.com 으로 보내 주세요.",
        absent=["1234567", "example.com"], present=["[주민번호]", "[이메일]"],
        roundtrip="원래 값이 사전에 없어 복원 불가",
        restore={"count": [("[주민번호]", 1), ("[이메일]", 1)], "warnings": 2})

    # 10) 보호자명이 들어간 상담 기록
    s = next(x for x in S if guardian_cnt[x["보호자명"]] == 1 and x["보호자명"] not in name_cnt
             and name_cnt[x["성명"]] == 1)
    gd = s["보호자명"]
    add("10) 보호자명 상담 기록",
        f"{gd} 어머님과 상담했다. {s['성명']}의 학교생활을 이야기했고, {gd}{josa(gd, '은', '는')} 협조를 약속했다.",
        absent=[gd, s["성명"]], counts=[(s["보호자코드"], 2), (s["학생코드"], 1)])
    return cases


def run(case, d):
    """문장 하나를 치환해서 보여주고, 기대한 결과인지 확인한다."""
    global failures
    text = case["text"]
    out, changes = mask_text(text, d)
    print(f"\n■ {case['title']}")
    print(f"  전: {text}")
    print(f"  후: {out}")
    if changes:
        print("  변경 목록:")
        for orig, new, kind, pos in changes:
            print(f"    - [{orig!r}, {new!r}, {kind!r}, {pos}]")
    else:
        print("  변경 목록: (없음)")

    problems = [f"남아 있으면 안 됨: {a}" for a in case["absent"] if a in out]
    problems += [f"없음: {p}" for p in case["present"] if p not in out]
    problems += [f"'{w}' {out.count(w)}개 (기대 {k}개)" for w, k in case["counts"] if out.count(w) != k]
    if problems:
        failures += 1
        print("  결과: FAIL - " + " / ".join(problems))
    else:
        print("  결과: OK")


def main():
    d = load_dictionary("data/학생사전.xlsx")
    print(f"사전 로드: 학생 {len(d.students)}명")
    for case in build_cases(d):
        run(case, d)
    print(f"\n==== 전체 결과: {'모두 OK' if not failures else f'FAIL {failures}건'} ====")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
