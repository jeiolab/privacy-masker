"""오류 보고용 글. 메시지·경로는 반드시 mask_text 로 가명처리한 뒤에만 복사·저장한다."""
from datetime import datetime
from pathlib import Path
import platform
import sys
import traceback

from core import config
from core.masker import Dictionary, load_dictionary, mask_text
from core.version import VERSION

LOG_NAME = "실행오류.txt"
_dict_cache = None


def _dictionary():
    """사전이 있으면 쓰고, 없거나 읽기 실패하면 빈 사전(전화·주민번호만이라도 가림)."""
    global _dict_cache
    if _dict_cache is not None:
        return _dict_cache
    try:
        p = config.dict_path()
        if p.exists():
            _dict_cache = load_dictionary(p)
            return _dict_cache
    except Exception:
        pass
    _dict_cache = Dictionary([])
    return _dict_cache


def mask_for_report(text):
    """오류 글·경로를 mask_text 로 가명처리한다. 실패해도 원문을 그대로 내보내지 않는다."""
    raw = "" if text is None else str(text)
    try:
        masked, _ = mask_text(raw, _dictionary())
        return masked
    except Exception:
        return "(개인정보를 가리지 못해 원문은 넣지 않았어요)"


def windows_version():
    if sys.platform.startswith("win"):
        rel, ver, *_ = platform.win32_ver()
        return f"Windows {rel} ({ver})".strip()
    return platform.platform()


def _rel_code_path(filename):
    """프로그램 폴더 기준 상대 경로. 밖이면 이름만 (사용자 폴더 경로를 남기지 않음)."""
    try:
        return str(Path(filename).resolve().relative_to(config.program_dir().resolve()))
    except Exception:
        return Path(filename).name


def _code_location(exc):
    frames = traceback.extract_tb(exc.__traceback__) if getattr(exc, "__traceback__", None) else []
    if not frames:
        return "(위치 없음)"
    prog = str(config.program_dir().resolve())
    app = [f for f in frames if str(Path(f.filename).resolve()).startswith(prog)]
    pick = app[-1] if app else frames[-1]
    loc = f"{_rel_code_path(pick.filename)}:{pick.lineno} ({pick.name})"
    return mask_for_report(loc)


def format_error_report(exc):
    """복사·저장용 오류 보고 글 (메시지·경로는 가명처리됨)."""
    kind = type(exc).__name__
    message = mask_for_report(f"{kind}: {exc}")
    frames = traceback.extract_tb(exc.__traceback__) if getattr(exc, "__traceback__", None) else []
    lines = []
    for f in frames:
        lines.append(f"  {_rel_code_path(f.filename)}:{f.lineno} in {f.name}")
    stack = mask_for_report("\n".join(lines)) if lines else "  (없음)"
    return (
        f"프로그램 버전: {VERSION}\n"
        f"Windows 버전: {windows_version()}\n"
        f"오류 종류: {kind}\n"
        f"오류가 난 코드 위치: {_code_location(exc)}\n"
        f"오류 메시지: {message}\n"
        f"호출 위치:\n{stack}\n"
    )


def append_error_log(report):
    """logs/실행오류.txt 에 가명처리된 보고를 덧붙인다."""
    try:
        folder = config.logs_dir()
        folder.mkdir(parents=True, exist_ok=True)
        with open(folder / LOG_NAME, "a", encoding="utf-8") as f:
            f.write(f"\n----- {datetime.now().isoformat(timespec='seconds')} -----\n")
            f.write(report if report.endswith("\n") else report + "\n")
    except OSError:
        pass


class MaskingLogWriter:
    """한 줄씩 mask_text 한 뒤 파일에 쓴다. (exe 의 표준 오류 출력용)"""

    def __init__(self, path):
        self.path = Path(path)
        self._buf = ""

    def write(self, text):
        self._buf += text if isinstance(text, str) else str(text)
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            self._write_line(line + "\n")
        return len(text) if text else 0

    def flush(self):
        if self._buf:
            self._write_line(self._buf)
            self._buf = ""

    def _write_line(self, line):
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(mask_for_report(line) if line.strip() else line)
        except OSError:
            pass
