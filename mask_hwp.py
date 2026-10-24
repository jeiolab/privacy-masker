#!/usr/bin/env python3
"""한글 파일(.hwp/.hwpx)을 서식 그대로 가명처리한다.

사용법: py mask_hwp.py samples/가짜_상담기록.hwp
결과  : output/AI용/원래이름_가명.hwp     (원본과 같은 형식, 서식 유지)
        output/보관용/원래이름_변경기록.xlsx (실제 이름이 들어 있으니 AI에 올리지 말 것)

방법:
  1. 원본을 output/AI용/원래이름_가명.hwp 로 복사한 뒤 그 복사본만 연다 (원본은 절대 열지 않음).
  2. 문서 전체 글자(표·머리말·꼬리말·메모 포함)를 꺼내 core.masker.mask_text 로 변경 목록을 얻는다.
  3. '원래 값 → 바뀐 값'을 긴 것부터 한글 '모두 바꾸기'로 본문·표·머리말·꼬리말·메모에 적용한다.
  4. 문서 정보(제목/지은이/주제/키워드)는 모두 비운다.
  5. 저장 후 다시 열어서 글자를 꺼내 core.checker 로 남은 개인정보를 검사한다.

- 한글은 백그라운드로 실행하고, 끝나면 반드시 닫는다. (Windows + 한글 설치 환경 전용)
- 원본 파일은 열지도 수정하지도 않는다. (먼저 py build_dict.py 로 학생사전을 만들어 둘 것)
"""
import sys
from collections import Counter
from pathlib import Path

from adapters.excel_adapter import safe_filename
from adapters.hwp_adapter import (
    DUPLICATE_WARNING, HWP_EXTS, HWP_WINDOWS_ONLY_MSG, mask_hwp,
    write_hwp_change_log,
)
from adapters.hwpx_adapter import mask_hwpx
from core import config
from core.masker import load_dictionary, mask_text

BASE_DIR = config.user_data_dir()
DICT_PATH = config.dict_path()
AI_DIR = config.ai_dir()
KEEP_DIR = config.keep_dir()
SPECIAL_KINDS = ("사전에없음", "동명이인확인")


def _ask_yes_no(question):
    """y/n 를 물어본다. y 면 True, 그 외엔 False."""
    try:
        answer = input(f"{question} (y/n) ").strip().lower()
    except EOFError:
        return False
    return answer in ("y", "yes")


def main(argv):
    if len(argv) != 2:
        print("사용법: py mask_hwp.py <한글파일>   (지원: " + " ".join(sorted(HWP_EXTS)) + ")")
        return 1
    src = Path(argv[1]).expanduser()
    if not src.is_file():
        print(f"[오류] 파일을 찾을 수 없습니다: {src}")
        return 1
    ext = src.suffix.lower()
    if ext not in HWP_EXTS:
        print(f"[오류] 한글 파일이 아닙니다: {ext}  (지원: " + " ".join(sorted(HWP_EXTS)) + ")")
        return 1
    if ext == ".hwp" and sys.platform != "win32":
        print(HWP_WINDOWS_ONLY_MSG)
        return 1

    try:
        dictionary = load_dictionary(DICT_PATH)
    except FileNotFoundError:
        print(f"[오류] {DICT_PATH} 가 없습니다. 먼저 py build_dict.py <명렬표.xlsx> 를 실행하세요.")
        return 1

    # 결과 파일 이름 (파일 이름에 학생 정보가 있으면 치환)
    stem, name_changes = mask_text(src.stem, dictionary)
    stem = safe_filename(stem)
    ai_path = AI_DIR / f"{stem}_가명{src.suffix.lower()}"
    keep_path = KEEP_DIR / f"{stem}_변경기록.xlsx"
    if src.resolve() == ai_path.resolve():
        print("[오류] 입력 파일이 결과 파일과 같습니다.")
        return 1
    overwritten = [p for p in (ai_path, keep_path) if p.exists()]

    # 동명이인이 있으면 바꾸기 전에 안내하고 계속할지 물어본다
    def confirm_duplicate():
        print(DUPLICATE_WARNING)
        return _ask_yes_no("계속할까요?")

    try:
        mask_fn = mask_hwpx if ext == ".hwpx" else mask_hwp
        result = mask_fn(src, ai_path, dictionary, confirm_duplicate=confirm_duplicate)
    except Exception as e:
        print(f"[오류] 한글 파일을 처리하지 못했습니다: {e}")
        return 1

    if result["status"] == "aborted":
        print("중단했습니다. (동명이인) md 방식(mask_md.py)을 이용해 주세요.")
        return 3

    # 파일 이름에서 바뀐 것도 변경기록에 함께 남긴다
    changes = [(o, n, k, 1) for o, n, k, _ in name_changes] + list(result["changes"])

    KEEP_DIR.mkdir(parents=True, exist_ok=True)
    write_hwp_change_log(changes, keep_path)

    print(f"원본  : {src}")
    print(f"AI용  : {ai_path.relative_to(BASE_DIR)}")
    print(f"보관용: {keep_path.relative_to(BASE_DIR)}")
    for p in overwritten:
        print(f"  (기존 파일을 덮어썼습니다: {p.relative_to(BASE_DIR)})")

    total = sum(c[3] for c in changes)
    print(f"\n[바뀐 개수] 종류 {len(changes)}가지 / 적용 {total}건")
    for kind, n in sorted(Counter(c[2] for c in changes).items(), key=lambda x: -x[1]):
        print(f"  - {kind}: {n}")

    for kind in SPECIAL_KINDS:
        items = [c for c in changes if c[2] == kind]
        if items:
            print(f"\n[{kind}] {len(items)}건 - 직접 확인하세요")
            for original, new, _, _ in items:
                print(f"  - {original} → {new}")

    print("\n[남은 개인정보 검사]")
    findings = result["findings"]
    if findings:
        print("⚠ 남은 개인정보")
        for f in findings:
            print(f"  - {f.location}: [{f.kind}] {f.value}")
        return 2
    print("✅ 남은 개인정보 없음")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
