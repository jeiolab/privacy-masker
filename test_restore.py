#!/usr/bin/env python3
"""복원 엔진 확인용 스크립트.   실행: python3 test_restore.py

1) 왕복 테스트   : test_masker.py 의 문장을 mask_text → unmask_text 해서 원문과 같은지 (실제 사전 사용)
2) AI 변형 테스트: AI가 코드를 바꿔 쓴 모양을 알아보는지   (가짜 사전 사용)
3) 조사 보정 테스트: 받침 있는 이름 / 없는 이름 / ㄹ 받침 이름   (가짜 사전 사용)
"""
import sys

from core.masker import Dictionary, load_dictionary, mask_text
from core.restorer import unmask_text
from test_masker import build_cases

passed = failed = 0


def check(title, ok, detail=""):
    global passed, failed
    if ok:
        passed += 1
        print(f"  PASS  {title}")
    else:
        failed += 1
        print(f"  FAIL  {title}")
    if detail and not ok:
        print(f"        {detail}")


# ---------------------------------------------------------------- 가짜 사전
def fake_student(num, name, sid, phone, gname, gphone, addr, birth):
    return {
        "번호": num, "학번": sid, "성명": name, "성뺀이름": name[1:], "보호자명": gname,
        "학생연락처": phone, "보호자연락처": gphone, "주소": addr, "생년월일": birth,
        "학생코드": f"학생{num}", "S코드": f"S{num}", "보호자코드": f"보호자{num}",
        "학생전화코드": f"[학생전화{num}]", "보호자전화코드": f"[보호자전화{num}]",
        "주소코드": f"[주소{num}]", "생년월일코드": f"[생년월일{num}]", "주의": "",
    }


def fake_dictionary():
    return Dictionary([
        fake_student("001", "홍길동", "10101", "010-1111-1111", "홍판서", "010-9999-0000", "서울시 가로 1", "2019-01-01"),
        fake_student("002", "김민지", "10102", "010-2222-2222", "홍판서", "010-9999-0000", "서울시 나로 2", "2019-02-02"),
        fake_student("003", "최영일", "10103", "010-3333-3333", "박말순", "010-8888-0000", "서울시 다로 3", "2019-03-03"),
        fake_student("004", "정하윤", "10104", "010-4444-4444", "정아빠", "010-7777-0001", "서울시 라로 4", "2019-04-04"),
        fake_student("005", "정하윤", "10205", "010-5555-5555", "정엄마", "010-7777-0002", "서울시 마로 5", "2019-05-05"),
    ])


# ---------------------------------------------------------------- 1) 왕복
def test_roundtrip():
    print("\n[1] 왕복 테스트 (mask_text → unmask_text)")
    d = load_dictionary("data/학생사전.xlsx")
    for case in build_cases(d):
        title, text = case["title"], case["text"]
        masked, _ = mask_text(text, d)
        restored, _, warns = unmask_text(masked, d)
        detail = f"원문: {text}\n        복원: {restored}"

        if case["roundtrip"] is True:
            check(f"{title}: 원래 문장과 같음", restored == text, detail)
            continue

        exp = case["restore"]
        if not exp:
            print(f"  제외  {title}: {case['roundtrip']}")
            continue
        ok = True
        for word, n in exp.get("count", []):
            ok &= restored.count(word) == n
        for word in exp.get("contains", []):
            ok &= word in restored
        if "warnings" in exp:
            ok &= len(warns) == exp["warnings"]
        check(f"{title}: ({case['roundtrip']}) 기대한 모양으로 복원/경고", ok, detail + f"\n        경고: {warns}")


# ---------------------------------------------------------------- 2) AI 변형
def test_variants():
    print("\n[2] AI 변형 테스트")
    d = fake_dictionary()
    cases = [
        # (입력, 기대 출력)
        ("학생 2는 발표했다", "김민지는 발표했다"),
        ("학생 3은 성실하다", "최영일은 성실하다"),
        ("학생1을 만났다", "홍길동을 만났다"),
        ("학생01을 만났다", "홍길동을 만났다"),
        ("학생001번의 의견", "홍길동의 의견"),
        ("학생001가 발표했다", "홍길동이 발표했다"),
        ("학생002가 발표했다", "김민지가 발표했다"),
        ("학생 003 이 발표했다", "최영일 이 발표했다"),    # 조사가 떨어져 있으면 조사로 보지 않음
        ("학생001 학생은 성실하다", "홍길동 학생은 성실하다"),
        ("학생001학생은 성실하다", "홍길동 학생은 성실하다"),
        ("**학생001**", "**홍길동**"),
        ("학생001, 학생002", "홍길동, 김민지"),
        ("학생 001과 학생 002는", "홍길동과 김민지는"),
        ("보호자003는 오셨다", "박말순은 오셨다"),
        ("S001 / S 002 / S003", "10101 / 10102 / 10103"),
        ("학생전화001 로 연락", "010-1111-1111 로 연락"),
        ("[학생전화001], [보호자전화002]", "010-1111-1111, 010-9999-0000"),
        ("[주소 003] / [생년월일003]", "서울시 다로 3 / 2019-03-03"),
        ("보호자001또는002는 오셨다", "홍판서는 오셨다"),          # 형제의 같은 보호자 → 한 사람
        ("학생001또는002", "홍길동/김민지(동명이인)"),
        ("학생004또는005이 뛰었다", "정하윤(동명이인)이 뛰었다"),
        # 건드리면 안 되는 것
        ("학생3명이 왔다", "학생3명이 왔다"),
        ("학생 2학년 3반", "학생 2학년 3반"),
        ("학생 1번째 줄", "학생 1번째 줄"),
        ("대학생3 이다", "대학생3 이다"),
        ("Galaxy S3 와 S 7", "Galaxy S3 와 S 7"),
        ("학생 99은 없다", "학생 99은 없다"),
        ("학생001가방을 들었다", "홍길동가방을 들었다"),    # 가방의 '가'는 조사가 아님
    ]
    for text, expected in cases:
        out, _, _ = unmask_text(text, d)
        check(f"{text!r} → {expected!r}", out == expected, f"실제: {out!r}")

    warn_cases = [
        ("학생099 이 발표했다", "학생099 이 발표했다", 1),      # 사전에 없는 코드
        ("학생099은 발표했다", "학생099은 발표했다", 1),
        ("[전화번호] [주민번호] [이메일]", "[전화번호] [주민번호] [이메일]", 3),
        ("학생004또는005", "정하윤(동명이인)", 1),
        ("[학생전화001또는002]", "010-1111-1111 또는 010-2222-2222(확인해 주세요)", 1),
    ]
    for text, expected, n in warn_cases:
        out, _, warns = unmask_text(text, d)
        check(f"{text!r} → {expected!r}, 경고 {n}건", out == expected and len(warns) == n,
              f"실제: {out!r}, 경고: {warns}")


# ---------------------------------------------------------------- 3) 조사 보정
def test_particles():
    print("\n[3] 조사 보정 테스트")
    d = fake_dictionary()
    # 이름: (코드 번호, 받침 있는 조사 / 받침 없는 조사 중 맞는 쪽)
    names = {"홍길동(받침 있음)": ("001", "홍길동", 0),
             "김민지(받침 없음)": ("002", "김민지", 1),
             "최영일(ㄹ 받침)": ("003", "최영일", 0)}
    pairs = [("이", "가"), ("은", "는"), ("을", "를"), ("과", "와"),
             ("아", "야"), ("이랑", "랑"), ("이라고", "라고")]
    for label, (num, name, idx) in names.items():
        for with_b, without in pairs:
            right = (with_b, without)[idx]
            for given in (with_b, without):          # AI가 어느 쪽으로 써도 맞게 고쳐져야 함
                for tail in (" 입니다", ".", ""):
                    out, _, _ = unmask_text(f"학생{num}{given}{tail}", d)
                    check(f"{label}: 학생{num}{given}{tail!r} → {name}{right}{tail!r}",
                          out == f"{name}{right}{tail}", f"실제: {out!r}")
    # 로 / 으로 (ㄹ 받침은 '로')
    for label, (num, name, _) in names.items():
        right = "으로" if name == "홍길동" else "로"
        for given in ("로", "으로"):
            out, _, _ = unmask_text(f"학생{num}{given} 갔다", d)
            check(f"{label}: 학생{num}{given} → {name}{right}", out == f"{name}{right} 갔다", f"실제: {out!r}")
    # 조사 뒤에 글자가 이어지면 조사가 아님
    for text, exp in [("학생001이다", "홍길동이다"), ("학생002가방", "김민지가방"), ("학생001은행", "홍길동은행")]:
        out, _, _ = unmask_text(text, d)
        check(f"조사 아님: {text!r} → {exp!r}", out == exp, f"실제: {out!r}")


def main():
    test_roundtrip()
    test_variants()
    test_particles()
    print(f"\n==== 전체 결과: 통과 {passed}건, 실패 {failed}건 ====")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
