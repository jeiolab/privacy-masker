#!/usr/bin/env python3
"""개인정보 가리개 - 배포용 빌드 스크립트.

쓰는 법:  python build_exe.py

하는 일 (순서대로):
 1) PyInstaller 로 폴더형(--onedir) · 콘솔 없음(--noconsole) 으로 빌드
    -> dist/개인정보가리개/개인정보가리개.exe
 2) ui/ 를 결과 폴더(exe 옆)에 함께 넣는다. samples/ 는 테스트용이라 넣지 않는다
 3) 결과 폴더에 data/·output/·logs/ 가 있으면 반드시 뺀다 (실제 학생 사전이 섞이면 안 됨)
 4) dist/개인정보가리개/ 를 dist/개인정보가리개.zip 으로 압축
 5) zip 안에 data 폴더나 학생사전.xlsx 가 없는지 검사해서 결과를 출력
 6) 빌드 중 경고·오류를 정리해서 보여 준다

참고: Python 이 없는 다른 Windows 에서도 이 exe 하나(폴더째)로 실행된다.
      data/·output/·logs/ 는 exe 옆에 자동으로 생긴다 (임시폴더가 아님).
"""
import importlib.util
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

PROJECT = Path(__file__).resolve().parent
APP_NAME = "개인정보가리개"
ENTRY = "launch.py"                     # exe 진입점 (app.py 는 그대로 둠)
DIST = PROJECT / "dist"
BUILD = PROJECT / "build"
APP_DIR = DIST / APP_NAME               # 최종 결과 폴더
EXE = APP_DIR / f"{APP_NAME}.exe"
ZIP_PATH = DIST / f"{APP_NAME}.zip"

INCLUDE_DIRS = ["ui"]                    # exe 옆에 함께 넣을 폴더. samples/ 는 배포하지 않는다
INCLUDE_FILES = ["사용설명서.md", "assets/jeio_logo.png"]       # exe 옆에 함께 넣을 파일
EXCLUDE_DIRS = ["data", "output", "logs"]   # 배포에서 반드시 빼는 폴더

# 빠지면 안 되는 부품 (설치돼 있을 때만 넣는다)
COLLECT_ALL = ["tkinterdnd2", "markitdown", "magika", "pyhwpx", "pywebview", "charset_normalizer"]
COPY_METADATA = ["markitdown"]
HIDDEN_IMPORTS = [
    "fitz", "pymupdf",                                  # pdf
    "win32com", "win32com.client", "pythoncom", "pywintypes", "win32timezone",  # pyhwpx(pywin32)
    "openpyxl", "xlrd", "pyperclip", "pywebview",
    "docx", "pptx",
]


def _has(name):
    """해당 모듈/패키지가 설치돼 있으면 True (실제로 불러오지는 않는다)."""
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError, ModuleNotFoundError):
        return False


def ensure_pyinstaller():
    if _has("PyInstaller"):
        return
    print("PyInstaller 가 없어 설치합니다...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)


def clean():
    for p in (BUILD, APP_DIR):
        if p.exists():
            shutil.rmtree(p, ignore_errors=True)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()


def build():
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onedir", "--noconsole",
        "--name", APP_NAME,
        "--distpath", str(DIST),
        "--workpath", str(BUILD),
        "--specpath", str(BUILD),
        "--paths", str(PROJECT),
    ]
    for pkg in COLLECT_ALL:
        if _has(pkg):
            cmd += ["--collect-all", pkg]
        else:
            print(f"  (건너뜀) --collect-all {pkg}: 설치돼 있지 않음")
    for pkg in COPY_METADATA:
        if _has(pkg):
            cmd += ["--copy-metadata", pkg]
    added = set()
    for mod in HIDDEN_IMPORTS:
        top = mod.split(".")[0]
        if _has(top) and mod not in added:
            cmd += ["--hidden-import", mod]
            added.add(mod)
        elif not _has(top):
            print(f"  (건너뜀) --hidden-import {mod}: 설치돼 있지 않음")
    cmd.append(ENTRY)

    print("\n[빌드] 다음 명령을 실행합니다:\n  " + " ".join(cmd) + "\n")
    proc = subprocess.run(cmd, cwd=str(PROJECT), text=True, encoding="utf-8",
                          errors="replace", capture_output=True)
    sys.stdout.write(proc.stdout or "")
    sys.stderr.write(proc.stderr or "")
    return proc


def add_extra_dirs():
    """ui/ · samples/ 를 결과 폴더(exe 옆)에 복사해 함께 포함한다."""
    for name in INCLUDE_DIRS:
        src = PROJECT / name
        if not src.exists():
            print(f"  (없음) {name}/ 폴더가 프로젝트에 없어 건너뜀")
            continue
        dst = APP_DIR / name
        if dst.exists():
            shutil.rmtree(dst, ignore_errors=True)
        shutil.copytree(src, dst)
        print(f"  포함: {name}/")
    for name in INCLUDE_FILES:
        src = PROJECT / name
        if not src.exists():
            print(f"  (없음) {name} 파일이 프로젝트에 없어 건너뜀")
            continue
        dst = APP_DIR / name
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        print(f"  포함: {name}")


def strip_private_dirs():
    """결과 폴더에서 data/·output/·logs/ 를 지운다 (실제 사전·결과가 섞이지 않도록)."""
    for name in EXCLUDE_DIRS:
        target = APP_DIR / name
        if target.exists():
            shutil.rmtree(target, ignore_errors=True)
            print(f"  제거: {name}/ (배포에서 제외)")


def make_zip():
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    shutil.make_archive(str(DIST / APP_NAME), "zip", root_dir=str(DIST), base_dir=APP_NAME)
    size_mb = ZIP_PATH.stat().st_size / (1024 * 1024)
    print(f"  압축: {ZIP_PATH}  ({size_mb:.1f} MB)")


def verify_zip():
    """zip 안에 (앱 최상위의) data 폴더나 학생사전.xlsx 가 없는지 검사한다.
    주의: _internal 안쪽 패키지들이 가진 'data' 하위 폴더는 정상이므로 문제 삼지 않는다."""
    bad = []
    with zipfile.ZipFile(ZIP_PATH) as zf:
        for member in zf.namelist():
            norm = member.replace("\\", "/")
            parts = norm.split("/")
            # 실제 학생 사전이 들어가는 자리: 개인정보가리개/data/... (앱 폴더 바로 아래)
            if len(parts) >= 2 and parts[0] == APP_NAME and parts[1] in EXCLUDE_DIRS:
                bad.append(member)
            elif Path(norm).name == "학생사전.xlsx":
                bad.append(member)
    if bad:
        print("  [검사 실패] zip 안에 들어가면 안 되는 항목이 있습니다:")
        for b in bad:
            print(f"     - {b}")
        return False
    print("  [검사 통과] zip 안에 data/output/logs 폴더도, 학생사전.xlsx 도 없습니다.")
    return True


def report_warnings(proc):
    """빌드 로그와 PyInstaller 경고 파일에서 중요한 경고·오류만 모아 보여 준다."""
    lines = []
    for text in (proc.stdout or "", proc.stderr or ""):
        for ln in text.splitlines():
            s = ln.strip()
            if s.startswith("ERROR") or ": ERROR:" in s or s.startswith("WARNING") or ": WARNING:" in s:
                lines.append(s)

    missing = []
    warn_file = BUILD / APP_NAME / f"warn-{APP_NAME}.txt"
    if warn_file.exists():
        for ln in warn_file.read_text(encoding="utf-8", errors="replace").splitlines():
            if "missing module named" in ln:
                name = ln.split("missing module named", 1)[1].strip()
                missing.append(name)

    # 흔한(무해한) 표준 라이브러리 미발견은 걸러 낸다
    ignore_prefix = ("_", "java", "org.", "com.sun", "test", "pyimod", "ctypes.macholib",
                     "AppKit", "Foundation", "Carbon", "gi.", "gtk", "readline", "termios",
                     "posix", "resource", "_posixsubprocess", "vms_lib", "multiprocessing.")
    notable = sorted({m for m in missing if not m.startswith(ignore_prefix)})

    print("\n================ 빌드 경고·오류 요약 ================")
    errs = [l for l in lines if "ERROR" in l]
    warns = [l for l in lines if "ERROR" not in l]
    if errs:
        print(f"[오류 {len(errs)}건]")
        for l in errs[:30]:
            print("  " + l)
    else:
        print("[오류] 없음")
    print(f"[경고 {len(warns)}건] (PyInstaller 일반 경고)")
    for l in warns[:15]:
        print("  " + l)
    if len(warns) > 15:
        print(f"  ... 외 {len(warns) - 15}건")
    if notable:
        print(f"[참고] 빌드 중 '찾지 못한 모듈' 후보 {len(notable)}개 "
              f"(대부분 선택적 의존성이라 무시해도 됩니다):")
        for m in notable[:25]:
            print("  - " + m)
        if len(notable) > 25:
            print(f"  ... 외 {len(notable) - 25}개")
    print("====================================================\n")


def main():
    for stream in (sys.stdout, sys.stderr):   # cp949 콘솔에서도 한글·기호 출력이 안 깨지게
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    print("개인정보 가리개 - 배포용 빌드 시작")
    from core.updater import write_version_json
    from core.version import VERSION
    dest, info = write_version_json(PROJECT / "version.json",
                                   notes="폴더 통째로 가리기(90_가림본), 바뀐 것만 다시, 결과 표")
    print(f"  버전 {VERSION} → {dest.name}")
    if info.get("download"):
        print(f"  받는 주소: {info['download']}")
    else:
        print("  (안내) config.DOWNLOAD_URL 을 넣으면 알림창에 받는 링크가 같이 나옵니다.")
        print("  (안내) config.UPDATE_URL 에 version.json 주소를 넣으면 켤 때 새 버전을 확인합니다.")
    ensure_pyinstaller()
    clean()

    proc = build()
    if proc.returncode != 0 or not EXE.exists():
        report_warnings(proc)
        print("[빌드 실패] exe 가 만들어지지 않았습니다. 위 오류를 확인하세요.")
        sys.exit(1)
    print(f"\n[1/5] 빌드 완료: {EXE}")

    print("[2/5] ui/ · 사용설명서 포함")
    add_extra_dirs()

    print("[3/5] data/ · output/ · logs/ 제외")
    strip_private_dirs()

    print("[4/5] zip 압축")
    make_zip()

    print("[5/5] zip 검사")
    ok = verify_zip()

    report_warnings(proc)

    print(f"결과 폴더 : {APP_DIR}")
    print(f"실행 파일 : {EXE}")
    print(f"배포 zip  : {ZIP_PATH}")
    if not ok:
        print("[주의] zip 검사에서 문제가 발견됐습니다. 위 목록을 확인하세요.")
        sys.exit(2)
    print("\n[완료] 모두 끝났습니다. dist/개인정보가리개.zip 를 그대로 나눠 주면 됩니다.")


if __name__ == "__main__":
    main()
