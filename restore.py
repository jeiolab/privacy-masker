#!/usr/bin/env python3
"""가명처리된 글을 실명으로 되돌린다.

사용법:
  python3 restore.py --clip          클립보드의 글을 복원해서 다시 클립보드에 넣음
  python3 restore.py 파일.txt        output/복원/파일_복원.txt 로 저장 (.md 도 가능)
  python3 restore.py 파일.xlsx       모든 칸·시트 이름·메모를 복원해서 output/복원/파일_복원.xlsx 로 저장
사전: data/학생사전.xlsx   (원본 파일은 읽기만 하고 수정하지 않는다)
"""
import sys
from collections import Counter
from pathlib import Path

from adapters.excel_adapter import load_name_map, restore_workbook, safe_filename
from core import config
from core.masker import load_dictionary
from core.restorer import unmask_text
from workflows import find_keep_log

BASE_DIR = config.user_data_dir()
DICT_PATH = config.dict_path()
OUT_DIR = config.restore_dir()
USAGE = "사용법: python3 restore.py --clip | 파일.txt | 파일.md | 파일.xlsx"


def read_text(path):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("글자를 읽지 못했어요. UTF-8 또는 한글(CP949) 파일만 돼요.")


def print_summary(changes, warnings):
    kinds = [c.kind if hasattr(c, "kind") else c[2] for c in changes]
    print(f"\n[복원] 총 {len(changes)}건")
    for kind, n in sorted(Counter(kinds).items(), key=lambda x: -x[1]):
        print(f"  - {kind}: {n}")
    if warnings:
        print(f"\n[경고] {len(warnings)}건 - 직접 확인하세요")
        for w in warnings:
            print(f"  - {w}")
    else:
        print("\n[경고] 없음")


def main(argv):
    if len(argv) != 2:
        print(USAGE)
        return 1
    try:
        dictionary = load_dictionary(DICT_PATH)
    except FileNotFoundError:
        print(f"[오류] {DICT_PATH} 가 없습니다. 먼저 python3 build_dict.py <명렬표.xlsx> 를 실행하세요.")
        return 1

    arg = argv[1]

    if arg == "--clip":
        try:
            import pyperclip
        except ImportError:
            print("[오류] pyperclip 이 없습니다. pip install -r requirements.txt 를 실행하세요.")
            return 1
        text = pyperclip.paste()
        if not text.strip():
            print("[오류] 클립보드가 비어 있습니다.")
            return 1
        restored, changes, warnings = unmask_text(text, dictionary)
        pyperclip.copy(restored)
        print("클립보드의 글을 복원해서 다시 클립보드에 넣었습니다.")
        print_summary(changes, warnings)
        return 0

    src = Path(arg).expanduser()
    if not src.is_file():
        print(f"[오류] 파일을 찾을 수 없습니다: {src}")
        return 1
    ext = src.suffix.lower()
    if ext not in (".txt", ".md", ".xlsx"):
        print(f"[오류] .txt, .md, .xlsx 파일만 처리할 수 있습니다. ({USAGE})")
        return 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stem = safe_filename(src.stem)
    extra = {}
    keep = find_keep_log(src)
    if keep:
        extra = load_name_map(keep)
    try:
        if ext == ".xlsx":
            out = OUT_DIR / f"{stem}_복원.xlsx"
            if out.resolve() == src.resolve():
                print("[오류] 입력 파일이 결과 파일과 같습니다.")
                return 1
            changes, warnings = restore_workbook(src, out, dictionary, extra=extra)
        else:
            out = OUT_DIR / f"{stem}_복원.txt"
            restored, changes, warnings = unmask_text(read_text(src), dictionary, extra)
            out.write_text(restored, encoding="utf-8")
    except Exception as e:  # 손상 파일, 인코딩 문제 등
        print(f"[오류] 복원하지 못했습니다: {e}")
        return 1

    print(f"원본: {src}")
    print(f"결과: {out.relative_to(BASE_DIR)}")
    print_summary(changes, warnings)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
