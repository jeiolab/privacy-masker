#!/usr/bin/env python3
"""지금까지 만든 기능을 한 번에 점검하는 통합 테스트.   실행: python run_all_tests.py

반드시 samples/ 의 가짜 데이터와 '임시로 만든 가짜 사전'만 쓴다.
실제 사전(data/학생사전.xlsx)은 절대 읽거나 바꾸지 않는다.
(기존 기능 코드는 수정하지 않고, 그 함수들을 불러서 쓰기만 한다.)

점검 항목
  1. 사전 만들기 + 추가 모드 (한 번 붙은 번호가 바뀌지 않는지)
  2. 치환 엔진        (test_masker 의 문장들)
  3. 복원            (왕복 / 조사 보정 / AI 변형 코드)
  4. 엑셀 / PDF / md 처리 후 남은 개인정보 없음
  5. 한글(.hwp) 처리  (Windows + 한글이 있을 때만, 없으면 건너뜀)
  6. 교직원 치환 · 이름 의심 경고

항목마다 ✅ 통과 / ❌ 실패 / ⏭ 건너뜀 으로 보여 주고,
실패하면 어떤 문장·파일에서 무엇이 틀렸는지 알려 준다.
테스트 중 만든 임시 파일은 끝나면 모두 지운다.
"""
import io
import shutil
import sys
import tempfile
import traceback
from contextlib import redirect_stdout
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
SAMPLES = BASE / "samples"

import build_dict
from adapters.excel_adapter import mask_workbook
from adapters.md_adapter import apply_extra_names, mask_markdown, to_markdown
from core.checker import add_ignored_word, check_file, check_text, find_name_suspects
from core.masker import load_dictionary, mask_text
from core.restorer import unmask_text
from test_masker import build_cases
import test_restore

PASS, FAIL, SKIP = "✅ 통과", "❌ 실패", "⏭ 건너뜀"
ROSTER = SAMPLES / "가짜_명렬표.xlsx"
STAFF = SAMPLES / "가짜_교직원.xlsx"


# ---------------------------------------------------------------- 임시 가짜 사전
def build_temp_student_dict(tmp):
    """samples/가짜_명렬표.xlsx 로 임시 학생사전을 만든다. (실제 사전은 건드리지 않는다)"""
    dict_path = tmp / "임시_학생사전.xlsx"
    _, rows = build_dict.load_rows(ROSTER)
    hi = build_dict.find_header_row(rows)
    mapping = build_dict.map_columns(rows[hi])
    students = build_dict.build_students(rows, hi, mapping, [])
    build_dict.assign_codes(students, mapping, [])
    build_dict.add_cautions(students, [])
    build_dict.write_dict(students, mapping, dict_path)
    return dict_path, rows, hi, mapping, students


# ---------------------------------------------------------------- 1) 사전 + 추가 모드
def item1(dict_path, rows, hi, mapping, students):
    bad = []
    if len(students) != 30:
        bad.append(f"사전 만들기: 학생 수 {len(students)}명 (기대 30명)")
    orig = {s["학번"]: s["번호"] for s in students}
    width = max(len(n) for n in orig.values())

    # 학생 추가 모드: 한 명 빼고(=명렬표에 없음), 한 명 새로 추가
    existing = build_dict.load_existing_dict(dict_path)
    mod = [r[:] for r in rows]
    dropped = mod.pop()                                   # 파일 마지막 학생 제거
    dropped_id = build_dict.cell_to_str(dropped[mapping["학번"]])
    dropped_name = build_dict.norm(dropped[mapping["성명"]])
    new_row = [""] * len(rows[hi])
    new_row[mapping["학번"]] = "19999"
    new_row[mapping["성명"]] = "신규신입"
    new_row[mapping["학생연락처"]] = "010-7777-0007"
    new_row[mapping["보호자명"]] = "신규보호자"
    new_row[mapping["보호자연락처"]] = "010-6666-0006"
    new_row[mapping["주소"]] = "서울 테스트구 테스트로 1"
    new_row[mapping["생년월일"]] = "2015-12-31"
    mod.append(new_row)
    new_students = build_dict.build_students(mod, hi, mapping, [])
    merged, stats = build_dict.merge_students(existing, new_students, mapping, [])
    build_dict.apply_missing_notes(merged, [])
    codes = {s["학번"]: s["번호"] for s in merged}

    for sid, num in orig.items():
        if codes.get(sid) != num:
            bad.append(f"학생 번호가 바뀜: 학번 {sid} {num} → {codes.get(sid)}")
    expect_new = str(max(int(n) for n in orig.values()) + 1).zfill(width)
    if codes.get("19999") != expect_new:
        bad.append(f"새 학생 번호: {codes.get('19999')} (기대 {expect_new})")
    dmerged = next((s for s in merged if s["학번"] == dropped_id), None)
    if not dmerged or "명렬표에 없음" not in dmerged.get("주의", ""):
        bad.append(f"빠진 학생 표시 안 됨: {dropped_name}({dropped_id})")
    if stats.get("added") != 1 or stats.get("missing") != 1:
        bad.append(f"요약 이상: {stats}")

    # 교직원 추가 모드: 번호 유지 + 새 교직원 추가 + 빠진 교직원 표시
    _, names = build_dict.load_staff_names(STAFF)
    m1, _ = build_dict.merge_staff([], names, [])
    c1 = {x["성명"]: x["번호"] for x in m1}
    m2, s2 = build_dict.merge_staff(m1, names[1:] + ["신규교사"], [])
    c2 = {x["성명"]: x["번호"] for x in m2}
    for nm, num in c1.items():
        if c2.get(nm) != num:
            bad.append(f"교직원 번호가 바뀜: {nm} {num} → {c2.get(nm)}")
    if c2.get("신규교사") != str(len(names) + 1).zfill(3):
        bad.append(f"새 교직원 번호: {c2.get('신규교사')} (기대 {len(names) + 1:03d})")
    sdrop = next((x for x in m2 if x["성명"] == names[0]), None)
    if not sdrop or "명단에 없음" not in sdrop.get("주의", ""):
        bad.append(f"빠진 교직원 표시 안 됨: {names[0]}")

    if bad:
        return FAIL, bad
    return PASS, [f"학생 30명 생성 · 추가 모드 요약 {stats} · 교직원 번호도 유지됨"]


# ---------------------------------------------------------------- 2) 치환 엔진
def item2(d):
    cases = build_cases(d)
    bad = []
    for case in cases:
        out, _ = mask_text(case["text"], d)
        probs = [f"남으면 안 되는 '{a}' 가 남음" for a in case["absent"] if a in out]
        probs += [f"'{p}' 가 없음" for p in case["present"] if p not in out]
        probs += [f"'{w}' {out.count(w)}개 (기대 {k}개)" for w, k in case["counts"] if out.count(w) != k]
        if probs:
            bad.append(f"[{case['title']}] " + " / ".join(probs) + f"\n      후: {out}")
    if bad:
        return FAIL, bad
    return PASS, [f"문장 {len(cases)}개 모두 기대대로 치환됨"]


# ---------------------------------------------------------------- 3) 복원
def item3(d):
    bad = []
    # (a) 왕복 테스트 (test_masker 문장 → mask → unmask)
    for case in build_cases(d):
        masked, _ = mask_text(case["text"], d)
        restored, _, warns = unmask_text(masked, d)
        if case["roundtrip"] is True:
            if restored != case["text"]:
                bad.append(f"[왕복: {case['title']}] 원문과 다름\n      원: {case['text']}\n      복: {restored}")
            continue
        exp = case["restore"]
        if not exp:
            continue
        ok = True
        for word, n in exp.get("count", []):
            ok &= restored.count(word) == n
        for word in exp.get("contains", []):
            ok &= word in restored
        if "warnings" in exp:
            ok &= len(warns) == exp["warnings"]
        if not ok:
            bad.append(f"[왕복: {case['title']}] 기대 복원/경고와 다름\n      복: {restored}\n      경고: {warns}")

    # (b) AI 변형 / (c) 조사 보정 : test_restore 의 '가짜 사전' 전용 함수 재사용
    for label, fn in (("AI 변형", test_restore.test_variants), ("조사 보정", test_restore.test_particles)):
        before = test_restore.failed
        buf = io.StringIO()
        with redirect_stdout(buf):
            fn()
        if test_restore.failed > before:
            fail_lines = [ln.strip() for ln in buf.getvalue().splitlines() if "FAIL" in ln]
            bad.append(f"[{label}] {test_restore.failed - before}건 실패:\n      " + "\n      ".join(fail_lines[:12]))

    if bad:
        return FAIL, bad
    return PASS, ["왕복 · AI 변형 · 조사 보정 모두 통과"]


# ---------------------------------------------------------------- 4) 엑셀/PDF/md
def _finding_text(findings):
    return "; ".join(
        f"{getattr(f, 'sheet', '') and f.sheet + '/'}{f.location} {f.kind}:{f.value}" for f in findings[:5]
    )


def item4(d, tmp):
    bad = []
    # 엑셀 2개 (원본 형식 유지 방식)
    for name in ("가짜_상담기록.xlsx", "가짜_성적.xlsx"):
        src = SAMPLES / name
        if not src.exists():
            bad.append(f"[엑셀 {name}] 샘플 파일이 없음")
            continue
        dst = tmp / f"out_{name}"
        mask_workbook(src, dst, d)
        findings = check_file(dst, d)
        if findings:
            bad.append(f"[엑셀 {name}] 남은 개인정보 {len(findings)}건: {_finding_text(findings)}")

    # PDF (md 변환 방식)
    pdf = SAMPLES / "가짜_명단.pdf"
    if not pdf.exists():
        bad.append("[PDF 가짜_명단.pdf] 샘플 파일이 없음")
    else:
        try:
            masked, _ = mask_markdown(to_markdown(pdf), d)
            findings = check_text(masked, d)
            if findings:
                bad.append(f"[PDF 가짜_명단.pdf] 남은 개인정보 {len(findings)}건: {_finding_text(findings)}")
        except Exception as e:
            bad.append(f"[PDF 가짜_명단.pdf] 처리 오류: {type(e).__name__}: {e}")

    # md (임시로 만든 md 파일)
    md_file = tmp / "테스트_상담.md"
    plain = [s for s in d.students if not s.get("주의")][:3]
    lines = [f"{s['성명']} 학생({s['학번']}) 연락처 {s['학생연락처']}, 보호자 {s['보호자명']}님." for s in plain]
    md_file.write_text("# 상담 기록\n\n" + "\n".join(lines) + "\n", encoding="utf-8")
    masked, _ = mask_markdown(to_markdown(md_file), d)
    findings = check_text(masked, d)
    if findings:
        bad.append(f"[md 테스트_상담.md] 남은 개인정보 {len(findings)}건: {_finding_text(findings)}")

    if bad:
        return FAIL, bad
    return PASS, ["엑셀 2개 · PDF · md 모두 남은 개인정보 없음"]


# ---------------------------------------------------------------- 5) 한글(.hwp)
def item5(d, tmp):
    if sys.platform != "win32":
        return SKIP, ["Windows 가 아니라 한글 파일을 처리할 수 없음"]
    try:
        import pyhwpx  # noqa: F401
    except Exception:
        return SKIP, ["pyhwpx 가 설치되어 있지 않음"]
    try:   # 아래아한글이 실제로 실행되는지 가볍게 확인
        from pyhwpx import Hwp
        probe = Hwp(new=True, visible=False)
        try:
            probe.quit()
        except Exception:
            pass
    except Exception as e:
        return SKIP, [f"아래아한글 실행 불가(미설치 등): {type(e).__name__}: {e}"]

    src = SAMPLES / "가짜_상담기록.hwp"
    if not src.exists():
        return SKIP, ["samples/가짜_상담기록.hwp 샘플이 없음"]
    from adapters.hwp_adapter import mask_hwp
    res = mask_hwp(src, tmp / "out_상담기록.hwp", d)
    findings = res.get("findings", [])
    if findings:
        return FAIL, [f"[한글 가짜_상담기록.hwp] 남은 개인정보 {len(findings)}건: {_finding_text(findings)}"]
    return PASS, [f"한글 처리 후 남은 개인정보 없음 (치환 {len(res.get('changes', []))}종)"]


# ---------------------------------------------------------------- 6) 교직원 · 이름 의심
def item6(dict_path, tmp):
    bad = []
    # 임시 사전에 '교직원' 시트를 추가한다 (학생 시트는 그대로)
    _, names = build_dict.load_staff_names(STAFF)
    staff, _ = build_dict.merge_staff([], names, [])
    build_dict.write_staff_sheet(staff, dict_path)
    d2 = load_dictionary(dict_path)
    if len(d2.staff) != len(names):
        bad.append(f"교직원 로드 수 {len(d2.staff)} (기대 {len(names)})")
    if "박영희" not in names:
        bad.append("가짜_교직원.xlsx 에 '박영희' 가 없어 테스트 문장과 맞지 않음")

    for sent in ("박영희 선생님께 상담 의뢰", "영희 쌤이 말씀하심"):
        m, _ = mask_text(sent, d2)
        r, _, _ = unmask_text(m, d2)
        if "박영희" in m or "교사" not in m:
            bad.append(f"[교직원 치환] {sent!r} → {m!r} (교사 코드로 안 바뀜)")
        if "박영희" not in r:
            bad.append(f"[교직원 복원] {m!r} → {r!r} (실명으로 안 돌아옴)")

    # 이름 의심 경고
    text = "상담 결과 김배추 학생과 최수호 선생님이 참석. 홍길순씨에게 전달. 이번 주 정말 고마웠다."
    words = {s.word for s in find_name_suspects(text, d2)}
    missing = {"김배추", "최수호", "홍길순"} - words
    if missing:
        bad.append(f"[이름 의심] 못 찾은 이름: {missing} (찾은 것: {words})")
    wrong = {"이번", "정말"} & words
    if wrong:
        bad.append(f"[이름 의심] 예외 단어를 잘못 잡음: {wrong}")

    # 조사·시간 표현 오탐 제외: "주까지"(주+까지) 등은 이름으로 잡히면 안 된다
    text_time = "다음 주까지 과제를 내기로 함"
    words_time = {s.word for s in find_name_suspects(text_time, d2)}
    if words_time:
        bad.append(f"[이름 의심] 조사·시간 표현을 이름으로 오탐: {words_time} (문장: {text_time!r})")

    # 예외 단어: 이름 자리에 와도 경고하지 않음 / 실제 이름 형태는 여전히 경고
    ignore_file = tmp / "무시단어.txt"   # 실제 data/무시단어.txt 는 건드리지 않는다
    for sent in ("남학생들과 상담", "국어 선생님께 전달", "이번 주까지 제출"):
        got = {s.word for s in find_name_suspects(sent, d2, ignore_path=ignore_file)}
        if got:
            bad.append(f"[이름 의심] 예외 문장에서 경고가 뜸: {sent!r} → {got}")
    got_name = {s.word for s in find_name_suspects("학원 최민준 선생님", d2, ignore_path=ignore_file)}
    if "최민준" not in got_name:
        bad.append(f"[이름 의심] '학원 최민준 선생님' 을 못 찾음 (찾은 것: {got_name})")

    # 사전 실명은 '항상 무시' 목록에 넣을 수 없다
    reject = "실제 이름은 항상 무시할 수 없어요"
    real = next((s.get("성명") for s in d2.students if s.get("성명")), "")
    msg = add_ignored_word(real, d2, path=ignore_file)
    if msg != reject:
        bad.append(f"[항상 무시] 사전 실명 {real!r} 저장이 거부되지 않음 (반환: {msg!r})")
    if ignore_file.exists() and real in ignore_file.read_text(encoding="utf-8").splitlines():
        bad.append(f"[항상 무시] 거부된 실명 {real!r} 이 파일에 저장됨")

    # '이 단어도 가리기' → [이름N] → 복원
    masked2, changes, name_map = apply_extra_names(text, ["김배추"])
    if "[이름1]" not in masked2 or "김배추" in masked2:
        bad.append(f"[이름 의심] '이 단어도 가리기' 실패: {masked2!r}")
    if not any(c[3] == "이름의심확정" for c in changes):
        bad.append("[이름 의심] 변경기록에 '이름의심확정' 이 없음 (복원용 기록 누락)")
    r2, _, _ = unmask_text(masked2, d2, name_map)
    if "김배추" not in r2 or "[이름1]" in r2:
        bad.append(f"[이름 의심] 복원 실패: {r2!r}")

    if bad:
        return FAIL, bad
    return PASS, ["교직원 치환/복원 OK · 이름 의심 경고·가리기·복원 OK"]


# ---------------------------------------------------------------- 7) 데이터 폴더 위치
def item7():
    """쓰는 폴더(data/output/logs)가 프로그램 폴더 밖에 있는지 본다. (실제 파일은 건드리지 않음)"""
    from core import config
    prog = config.program_dir().resolve()
    root = config.user_data_dir().resolve()
    docs = config.documents_dir().resolve()
    bad = []
    if root.parent != docs:
        bad.append(f"데이터 루트가 문서 폴더 아래가 아님: {root} (문서: {docs})")
    if root.name != config.USER_DATA_DIR_NAME:
        bad.append(f"데이터 폴더 이름이 {config.USER_DATA_DIR_NAME!r} 가 아님: {root.name}")
    for label, folder in (("data", config.data_dir()), ("output", config.output_dir()),
                          ("logs", config.logs_dir())):
        p = folder.resolve()
        if p == prog or prog in p.parents:
            bad.append(f"{label} 폴더가 프로그램 폴더 안에 있음: {p}")
    if not bad:
        return PASS, [f"데이터 위치: {root} (프로그램: {prog})"]
    return FAIL, bad


# ---------------------------------------------------------------- 실행
def main():
    print("개인정보 가리개 - 통합 테스트")
    print("(samples/ 가짜 데이터 + 임시 가짜 사전만 사용 · 실제 사전은 건드리지 않음)")
    tmp = Path(tempfile.mkdtemp(prefix="run_all_tests_"))
    n_pass = n_fail = n_skip = 0
    try:
        dict_path, rows, hi, mapping, students = build_temp_student_dict(tmp)
        d = load_dictionary(dict_path)

        plan = [
            ("1. 사전 만들기 + 추가 모드 (번호 유지)", lambda: item1(dict_path, rows, hi, mapping, students)),
            ("2. 치환 엔진 (test_masker 문장)", lambda: item2(d)),
            ("3. 복원 (왕복 · 조사 보정 · AI 변형)", lambda: item3(d)),
            ("4. 엑셀 / PDF / md 처리 후 남은 개인정보 없음", lambda: item4(d, tmp)),
            ("5. 한글(.hwp) 처리 (Windows+한글일 때만)", lambda: item5(d, tmp)),
            ("6. 교직원 치환 · 이름 의심 경고", lambda: item6(dict_path, tmp)),
            ("7. 데이터 폴더가 프로그램 폴더 밖에 있는지", item7),
        ]
        for title, fn in plan:
            try:
                status, details = fn()
            except Exception:
                tail = traceback.format_exc().splitlines()[-4:]
                status, details = FAIL, ["예외 발생:\n      " + "\n      ".join(tail)]
            print(f"\n{status}  {title}")
            for line in details:
                print(f"   - {line}")
            n_pass += status == PASS
            n_fail += status == FAIL
            n_skip += status == SKIP
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    total = n_pass + n_fail
    print("\n" + "=" * 60)
    if n_skip:
        print(f"전체 {total}개 중 {n_pass}개 통과  (⏭ 건너뜀 {n_skip}개)")
    else:
        print(f"전체 {total}개 중 {n_pass}개 통과")
    return 1 if n_fail else 0


if __name__ == "__main__":
    sys.exit(main())
