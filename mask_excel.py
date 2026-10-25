#!/usr/bin/env python3
"""엑셀 파일을 가명처리한다.

사용법: python3 mask_excel.py samples/가짜_성적.xlsx
결과  : output/AI용/원래이름_가명.xlsx        (AI에 올릴 파일)
        output/보관용/원래이름_변경기록.xlsx   (바꾼 내용 기록 - 실제 이름이 들어 있으니 AI에 올리지 말 것)
원본 파일은 읽기만 하고 수정하지 않는다. (먼저 python3 build_dict.py 로 학생사전을 만들어 둘 것)
"""
import sys
from collections import Counter
from pathlib import Path

from adapters.excel_adapter import Change, mask_workbook, safe_filename, write_change_log
from core import config
from core.checker import check_file
from core.masker import load_dictionary, mask_text

BASE_DIR = config.user_data_dir()
DICT_PATH = config.dict_path()
AI_DIR = config.ai_dir()
KEEP_DIR = config.keep_dir()
SPECIAL_KINDS = ("사전에없음", "동명이인확인")


def main(argv):
    if len(argv) != 2:
        print("사용법: python3 mask_excel.py <엑셀파일.xlsx>")
        return 1
    src = Path(argv[1]).expanduser()
    if not src.is_file():
        print(f"[오류] 파일을 찾을 수 없습니다: {src}")
        return 1
    if src.suffix.lower() != ".xlsx":
        print("[오류] .xlsx 파일만 처리할 수 있습니다.")
        return 1
    try:
        dictionary = load_dictionary(DICT_PATH)
    except FileNotFoundError:
        print(f"[오류] {DICT_PATH} 가 없습니다. 먼저 python3 build_dict.py <명렬표.xlsx> 를 실행하세요.")
        return 1

    # 결과 파일 이름 (파일 이름에 학생 정보가 있으면 치환)
    stem, name_changes = mask_text(src.stem, dictionary)
    stem = safe_filename(stem)
    ai_path = AI_DIR / f"{stem}_가명.xlsx"
    keep_path = KEEP_DIR / f"{stem}_변경기록.xlsx"
    AI_DIR.mkdir(parents=True, exist_ok=True)
    KEEP_DIR.mkdir(parents=True, exist_ok=True)
    if src.resolve() in (ai_path.resolve(), keep_path.resolve()):
        print("[오류] 입력 파일이 결과 파일과 같습니다.")
        return 1

    try:
        changes, warnings = mask_workbook(src, ai_path, dictionary)
    except Exception as e:  # 손상/비밀번호 파일 등
        print(f"[오류] 엑셀을 처리하지 못했습니다: {e}")
        return 1
    changes = [Change("(파일 이름)", "", "-", o, n, k) for o, n, k, _ in name_changes] + changes
    write_change_log(changes, keep_path)

    findings = check_file(ai_path, dictionary)

    print(f"원본  : {src}")
    print(f"AI용  : {ai_path.relative_to(BASE_DIR)}")
    print(f"보관용: {keep_path.relative_to(BASE_DIR)}")

    print(f"\n[바뀐 개수] 총 {len(changes)}건")
    for kind, n in sorted(Counter(c.kind for c in changes).items(), key=lambda x: -x[1]):
        print(f"  - {kind}: {n}")

    for kind in SPECIAL_KINDS:
        items = [c for c in changes if c.kind == kind]
        if items:
            print(f"\n[{kind}] {len(items)}건 - 직접 확인하세요")
            for c in items:
                print(f"  - {c.sheet} / {c.cell}: {c.original} → {c.new}")

    if warnings:
        print(f"\n[경고] {len(warnings)}건")
        for w in warnings:
            print(f"  - {w}")

    print("\n[남은 개인정보 검사]")
    if findings:
        print("⚠ 남은 개인정보")
        for f in findings:
            print(f"  - {f.sheet} / {f.location}: [{f.kind}] {f.value}")
        ai_path.unlink(missing_ok=True)   # 남으면 AI용 파일을 남기지 않는다 (2026-10-25)
        print(f"→ AI용 파일을 지웠습니다: {ai_path.name}")
        return 2
    print("✅ 남은 개인정보 없음")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
