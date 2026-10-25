#!/usr/bin/env python3
"""여러 형식의 파일을 md로 바꾼 뒤 가명처리한다.

사용법: python3 mask_md.py samples/가짜_명단.pdf
지원  : .pdf .docx .xlsx .pptx .csv (markitdown 으로 변환) / .txt .md (그대로 읽음)
결과  : output/AI용/원래이름_가명.md
        output/보관용/원래이름_변경기록.xlsx   (실제 이름이 들어 있으니 AI에 올리지 말 것)

- 변환된 원본 md(실명이 있는 글)는 파일로 저장하지 않고 메모리에서만 처리한다.
- 인터넷 없이 동작한다 (LLM·이미지 설명·외부 서비스 기능을 쓰지 않음).
- 원본 파일은 읽기만 하고 수정하지 않는다. (먼저 python3 build_dict.py 로 학생사전을 만들어 둘 것)
"""
import sys
from collections import Counter
from pathlib import Path

from adapters.excel_adapter import safe_filename
from adapters.md_adapter import (
    HWP_EXTS, HWP_WINDOWS_ONLY_MSG, SUPPORTED_EXTS, find_scanned_pages,
    mask_markdown, prepare_excel_markdown, to_markdown, write_md_change_log,
)
from core import config
from core.checker import check_md_file
from core.masker import load_dictionary, mask_text

BASE_DIR = config.user_data_dir()
DICT_PATH = config.dict_path()
AI_DIR = config.ai_dir()
KEEP_DIR = config.keep_dir()
SPECIAL_KINDS = ("사전에없음", "동명이인확인")


def main(argv):
    if len(argv) != 2:
        print("사용법: python3 mask_md.py <파일>   (지원: " + " ".join(sorted(SUPPORTED_EXTS)) + ")")
        return 1
    src = Path(argv[1]).expanduser()
    if not src.is_file():
        print(f"[오류] 파일을 찾을 수 없습니다: {src}")
        return 1
    ext = src.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        print(f"[오류] 지원하지 않는 형식입니다: {ext}  (지원: " + " ".join(sorted(SUPPORTED_EXTS)) + ")")
        return 1
    # .hwp 만 한글 프로그램이 필요하다. .hwpx 는 ZIP+XML 이라 어디서나 된다.
    if ext == ".hwp" and sys.platform != "win32":
        print(HWP_WINDOWS_ONLY_MSG)
        return 1
    try:
        dictionary = load_dictionary(DICT_PATH)
    except FileNotFoundError:
        print(f"[오류] {DICT_PATH} 가 없습니다. 먼저 python3 build_dict.py <명렬표.xlsx> 를 실행하세요.")
        return 1

    # 스캔된 PDF는 글자를 읽을 수 없으니, 아무것도 만들지 않고 멈춘다
    if ext == ".pdf":
        try:
            scanned = find_scanned_pages(src)
        except Exception as e:
            print(f"[오류] PDF를 열지 못했습니다: {e}")
            return 1
        if scanned:
            print(f"⚠ 스캔된 쪽: {', '.join(map(str, scanned))}쪽 - 처리 불가")
            print("  (글자가 거의 없는 쪽이 있어 결과 파일을 만들지 않고 멈춥니다)")
            return 3

    try:
        text = to_markdown(src)          # 메모리에서만 다룸 (파일로 저장하지 않음)
    except Exception as e:
        print(f"[오류] 변환하지 못했습니다: {e}")
        return 1
    if not text.strip():
        print("[오류] 변환된 글이 없습니다.")
        return 1

    added_zero = frozenset()
    if ext in (".xlsx", ".csv"):   # 엑셀·엑셀에서 내보낸 csv는 앞 0이 빠진 번호가 흔하다
        text, added_zero = prepare_excel_markdown(text)
    masked, changes = mask_markdown(text, dictionary, added_zero)

    # 결과 파일 이름 (파일 이름에 학생 정보가 있으면 치환)
    stem, name_changes = mask_text(src.stem, dictionary)
    stem = safe_filename(stem)
    changes = [["(파일 이름)", o, n, k] for o, n, k, _ in name_changes] + changes
    ai_path = AI_DIR / f"{stem}_가명.md"
    keep_path = KEEP_DIR / f"{stem}_변경기록.xlsx"
    if src.resolve() in (ai_path.resolve(), keep_path.resolve()):
        print("[오류] 입력 파일이 결과 파일과 같습니다.")
        return 1
    overwritten = [p for p in (ai_path, keep_path) if p.exists()]

    AI_DIR.mkdir(parents=True, exist_ok=True)
    KEEP_DIR.mkdir(parents=True, exist_ok=True)
    ai_path.write_text(masked, encoding="utf-8")
    write_md_change_log(changes, keep_path)

    findings = check_md_file(ai_path, dictionary)

    print(f"원본  : {src}")
    print(f"AI용  : {ai_path.relative_to(BASE_DIR)}")
    print(f"보관용: {keep_path.relative_to(BASE_DIR)}")
    for p in overwritten:
        print(f"  (기존 파일을 덮어썼습니다: {p.relative_to(BASE_DIR)})")

    print(f"\n[바뀐 개수] 총 {len(changes)}건")
    for kind, n in sorted(Counter(c[3] for c in changes).items(), key=lambda x: -x[1]):
        print(f"  - {kind}: {n}")

    for kind in SPECIAL_KINDS:
        items = [c for c in changes if c[3] == kind]
        if items:
            print(f"\n[{kind}] {len(items)}건 - 직접 확인하세요")
            for line, original, new, _ in items:
                where = line if isinstance(line, str) else f"{line}번째 줄"
                print(f"  - {where}: {original} → {new}")

    print("\n[남은 개인정보 검사]")
    if findings:
        print("⚠ 남은 개인정보")
        for f in findings:
            print(f"  - {f.location}: [{f.kind}] {f.value}")
        ai_path.unlink(missing_ok=True)   # 남으면 AI용 파일을 남기지 않는다 (2026-10-25)
        print(f"→ AI용 파일을 지웠습니다: {ai_path.name}")
        return 2
    print("✅ 남은 개인정보 없음")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
