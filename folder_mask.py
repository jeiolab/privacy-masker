#!/usr/bin/env python3
"""폴더 통째로 가리기 (1.3.0, 2026-10-24 캡틴: 원본은 제자리, 가린 복사본은 90_가림본 에).

NAS 구조
  <뿌리>/20_업무_School/…/상담.hwpx              ← 원본 (그대로 둔다)
  <뿌리>/90_가림본/20_업무_School/…/<가린 이름>_가림.hwpx   ← 가린 복사본만
  <뿌리>/20_업무_School/_가림기록/…/<가린 이름>_변경기록.xlsx ← 보관용 기록(실명 대응표). 원본 쪽 숨김 폴더
  <뿌리>/90_가림본/_가림목록.json                 ← 무엇을 언제 가렸나(경로 해시·파일 해시·날짜만, 실명 없음)
  <뿌리>/20_업무_School/_가림기록/_원본목록.json     ← 원본 경로 ↔ 가림본 경로 (원본 이름이 들어 있어 기록 쪽에 둔다)

- 원본 폴더에는 _가림기록 말고 아무것도 쓰지 않는다. 원본은 읽기만 한다.
- 바뀐 것만 다시: 크기·수정 시각이 같으면 건너뛰고, 다르면 해시를 견줘 내용이 같으면 시각만 고친다.
- 원본이 없어진 가림본은 지우지 않고 결과 표에 「원본 없음」으로만 알린다.
- 결과 표에는 실명을 내지 않는다(파일 이름도 가린 이름으로).

사용법: python3 folder_mask.py <원본 폴더> [--가림본 <가림본 뿌리 폴더>]
"""
import hashlib
import json
import os
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import workflows

가림본이름 = "90_가림본"
기록이름 = "_가림기록"
목록이름 = "_가림목록.json"
원본목록이름 = "_원본목록.json"
번호폴더 = re.compile(r"^[0-6]\d_")          # 00_ ~ 69_ (가림본 90_ 은 아님)
원래형식 = {".xlsx", ".xls", ".hwp", ".hwpx"}   # 엑셀·한글은 원래 형식으로, 나머지는 AI에 올릴 글(md)로


@dataclass
class FolderRow:
    상태: str          # 가림 / 건너뜀 / 개인정보 남음 / 확인 필요 / 못 가림 / 원본 없음
    파일: str          # 가림본 쪽 상대 경로(가린 이름) — 실명 없음
    내용: str = ""
    위험: bool = False  # 이 줄의 가림본을 AI에 올리면 안 되는가(원본 없음이어도 지난번 경고가 있으면 True)


@dataclass
class FolderResult:
    가림본뿌리: Path
    기록폴더: Path
    rows: list = field(default_factory=list)
    대상아님: int = 0

    def count(self, 상태):
        return sum(1 for r in self.rows if r.상태 == 상태)

    @property
    def headline(self):
        n_ok, n_skip = self.count("가림"), self.count("건너뜀")
        n_left, n_fail, n_gone = self.count("개인정보 남음"), self.count("못 가림"), self.count("원본 없음")
        n_chk = self.count("확인 필요")
        parts = [f"가림 {n_ok}", f"건너뜀 {n_skip}"]
        if n_left: parts.append(f"개인정보 남음 {n_left}")
        if n_chk: parts.append(f"확인 필요 {n_chk}")
        if n_fail: parts.append(f"못 가림 {n_fail}")
        if n_gone: parts.append(f"원본 없음 {n_gone}")
        if self.대상아님: parts.append(f"대상 아님 {self.대상아님}")
        mark = "✅" if self.clean else "⚠"
        tail = "  남은 개인정보 없음" if self.clean else "  남은 개인정보·이름 후보·못 가린 파일을 확인하세요"
        return f"{mark} " + " · ".join(parts) + tail

    @property
    def clean(self):
        return not any(r.위험 or r.상태 in ("개인정보 남음", "확인 필요", "못 가림") for r in self.rows)


def find_layout(src, mask_root=None):
    """(원본 기준 폴더, 가림본 뿌리, 기록 폴더, 원본의 기준 상대 경로) 를 정한다.

    원본 경로에서 00_~69_ 번호 폴더를 찾으면 그 바로 위가 NAS 뿌리다.
    못 찾으면 mask_root(사용자가 고른 가림본 뿌리)가 있어야 하고, 그때 기록은 원본 폴더 안 _가림기록/ 에 둔다.
    """
    src = Path(src).resolve()
    # 가장 바깥 번호 폴더를 최상위로 — 「20_업무/01_평가」를 골라도 NAS 뿌리는 같은 곳 (강희 지적)
    for p in reversed([src, *src.parents]):
        if 번호폴더.match(p.name) and p.parent != p:
            nas, top = p.parent, p
            root = Path(mask_root).resolve() if mask_root else nas / 가림본이름
            return nas, root, top / 기록이름, src.relative_to(nas)
    if not mask_root:
        raise ValueError("NAS 번호 폴더(00_~60_)를 찾지 못했어요. 가림본을 둘 폴더를 골라 주세요.")
    return src.parent, Path(mask_root).resolve(), src / 기록이름, Path(src.name)


def _hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _key(rel):
    return hashlib.sha256(str(rel).encode("utf-8")).hexdigest()[:32]   # 원본 경로는 해시로만(실명이 든 이름일 수 있다)


def _load(path):
    try:
        d = json.loads(Path(path).read_text(encoding="utf-8"))
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def _save(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + f".{os.getpid()}.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)


def _walk(src, skip_dirs, 못읽은=None):
    # 읽지 못한 폴더는 조용히 넘기지 않고 모은다 — 결과 표에 「못 가림」으로 (강희 지적)
    for dirpath, dirnames, filenames in os.walk(src, onerror=(lambda e: 못읽은.append(e)) if 못읽은 is not None else None):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith((".", "#", "@")) and d != 기록이름
                             and Path(dirpath, d).resolve() not in skip_dirs)
        for name in sorted(filenames):
            if name.startswith((".", "~$")):
                continue
            yield Path(dirpath) / name


def mask_folder(src, mask_root=None, dictionary=None, progress=None):
    """원본 폴더 전체를 가려 가림본에 쓴다. progress(i, n) 가 있으면 진행을 알린다."""
    src = Path(src).resolve()
    if not src.is_dir():
        raise ValueError("원본 폴더를 찾을 수 없어요.")
    nas, root, keep_root, rel_base = find_layout(src, mask_root)
    if root == src or root in src.parents:
        raise ValueError("가림본 폴더 안을 원본으로 고를 수 없어요.")
    if 기록이름 in root.parts or root == keep_root or root in keep_root.parents or keep_root in root.parents:
        raise ValueError("가림본 폴더와 실명 기록(_가림기록) 폴더가 겹쳐요. 다른 곳을 골라 주세요.")   # 대응표가 가림본에 섞이지 않게 (강희 지적)
    dictionary = dictionary or workflows.get_dictionary()
    사전표 = _dict_mark(dictionary)                   # 사전이 바뀌면(새 이름 추가) 원본이 같아도 다시 가린다 (강희 지적)
    목록 = _load(root / 목록이름)
    원본목록 = _load(keep_root / 원본목록이름)
    result = FolderResult(root, keep_root)
    # 가림본의 최상위 폴더 이름: 가린 이름. 번호 폴더가 아닌 곳(고른 뿌리)은 원본 위치 해시를 붙인다 —
    #   다른 곳의 같은 이름 폴더(A/보고·B/보고)가 한 가림본 뿌리에서 기록 연결을 덮어쓰지 않게 (강희 지적)
    top_out = _masked_part(rel_base.parts[0], dictionary)
    if not 번호폴더.match(rel_base.parts[0]):
        top_out += "_" + hashlib.sha256(str(src).encode("utf-8")).hexdigest()[:6]
    elif root != (nas / 가림본이름).resolve():          # 번호 폴더라도 가림본을 따로 고르면 NAS 위치를 붙인다 — 다른 NAS 의 같은 번호 폴더와 섞이지 않게 (강희 지적)
        top_out += "_" + hashlib.sha256(str(nas).encode("utf-8")).hexdigest()[:6]
    if root != (nas / 가림본이름).resolve():             # 기본 90_가림본 이 아니면 기록도 가림본 뿌리마다 따로 — 두 출력이 대응표를 덮어쓰지 않게 (강희 지적)
        keep_root = keep_root / f"가림본_{hashlib.sha256(str(root).encode('utf-8')).hexdigest()[:6]}"
        result.기록폴더 = keep_root
        원본목록 = _load(keep_root / 원본목록이름)
    top_out = _unique_top(root, top_out, keep_root)     # 다른 최상위 폴더가 가린 뒤 같은 이름(20_[A]·20_A)이면 기록 연결을 덮어쓰지 않게 (강희 지적)
    _remember_keep(root, top_out, keep_root)            # 조회는 가린 이름으로 하니 가린 최상위 이름으로 남긴다 (강희 지적)

    files, 대상아님 = [], 0
    못읽은 = []
    for p in _walk(src, {root}, 못읽은):
        if p.suffix.lower() in workflows.MASK_EXTS:
            files.append(p)
        else:
            대상아님 += 1
    result.대상아님 = 대상아님
    for e in 못읽은:
        try:
            where = _safe_rel(rel_base / Path(e.filename).relative_to(src) / "_", dictionary, top_out)
        except Exception:
            where = "(폴더)"
        result.rows.append(FolderRow("못 가림", where, "폴더를 읽지 못함(권한·연결) — 안의 파일은 확인 안 됨"))
    seen = set()
    # 중간 결과는 이 작업만의 임시 폴더에 — 단일 파일 가리기의 AI용·보관용 결과와 기록을 덮어쓰지 않게 (강희 지적)
    old_dirs = (workflows.AI_DIR, workflows.KEEP_DIR)
    import tempfile
    Path(workflows.BASE_DIR).mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="폴더가리기_", dir=workflows.BASE_DIR) as 임시:   # 데이터 폴더 안 — 결과 안내(저장: …)가 그 기준으로 경로를 적는다
        workflows.AI_DIR, workflows.KEEP_DIR = Path(임시) / "AI용", Path(임시) / "보관용"
        try:
            for i, p in enumerate(files, 1):
                if progress:
                    progress(i, len(files))
                rel = rel_base / p.relative_to(src)
                k = _key(p.resolve())                  # 키는 원본의 실제 위치로 — 다른 부모 아래 같은 상대 경로가 섞이지 않게 (강희 지적)
                seen.add(k)
                try:
                    _one(p, rel, k, root, keep_root, 목록, 원본목록, dictionary, 사전표, result, top_out)
                except Exception as e:                   # 한 파일이 무엇으로 실패해도 나머지와 목록 저장은 이어 간다 (강희 지적)
                    result.rows.append(FolderRow("못 가림", _safe_rel(rel, dictionary, top_out), _reason(e)))
        finally:
            workflows.AI_DIR, workflows.KEEP_DIR = old_dirs
    # 원본이 없어진 가림본: 이번에 고른 폴더 아래였던 것만 본다. 지우지 않는다
    for k, v in sorted(원본목록.items()):
        원 = Path(str(v.get("원본", "")))
        if k in seen or not (원 == rel_base or rel_base in 원.parents):   # 경로 마디로 견준다 — 「평가」를 골랐는데 「평가자료」가 섞이지 않게 (강희 지적)
            continue
        gb = root / v.get("가림본", "")
        if not (nas / 원).exists() and gb.exists():
            기록 = 목록.get(k) or {}
            # 다시 검증할 원본이 없으니, 지난번에 깨끗했고 사전이 같고 가림본이 그대로일 때만 안전으로 본다 — 하나라도 모르면 경고 (강희 지적)
            나쁨 = not (기록.get("깨끗") is True and 기록.get("사전") == 사전표
                       and bool(기록.get("가림해시")) and _hash(gb) == 기록["가림해시"])
            result.rows.append(FolderRow("원본 없음", v.get("가림본", ""), "가림본은 그대로 두었어요" + (" · 지난번 개인정보·이름 후보 경고 그대로 — AI에 올리지 마세요" if 나쁨 else ""), 위험=나쁨))
    _save(root / 목록이름, 목록)
    _save(keep_root / 원본목록이름, 원본목록)
    return result


def _masked_part(part, dictionary):
    return workflows.safe_filename(workflows.mask_text(part, dictionary)[0]) or "_"


def _masked_rel(rel, dictionary, top_out=None):
    """폴더 이름까지 가린 상대 경로(표·가림본·목록에 쓰는 것) — 상위 폴더 이름에 든 실명도 남기지 않는다 (강희 지적)."""
    parts = [_masked_part(x, dictionary) for x in rel.parent.parts]
    if parts and top_out:
        parts[0] = top_out
    return str(Path(*parts, f"{workflows._masked_stem(rel, dictionary)[0]}{rel.suffix}"))


def _path_bad(path_text, dictionary):
    글 = " ".join(Path(path_text).parts).replace("_", " ")
    return bool(workflows.check_text(글, dictionary)) or bool(workflows.find_name_suspects(글, dictionary))


def _safe_rel(rel, dictionary, top_out=None):
    """표에 싣는 경로 — 가린 뒤에도 이름 후보가 남으면 경로를 싣지 않는다(실패 줄도, 강희 지적)."""
    shown = _masked_rel(rel, dictionary, top_out)
    if _path_bad(shown, dictionary):
        return f"(이름 후보가 든 경로의 {rel.suffix.lstrip('.') or '폴더'})"
    return shown[:-2] if shown.endswith("/_") else shown


def _dict_mark(dictionary):
    try:
        무시 = sorted(workflows.load_ignored_words())    # 「이름 아님」 무시 단어도 이름 검사에 쓰인다 — 바뀌면 다시 (강희 지적)
    except Exception:
        무시 = []
    data = {"s": getattr(dictionary, "students", []), "t": getattr(dictionary, "staff", []), "i": 무시}
    return hashlib.sha256(json.dumps(data, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()[:16]


def _one(p, rel, k, root, keep_root, 목록, 원본목록, dictionary, 사전표, result, top_out=None):
    st = p.stat()
    old = 목록.get(k) or {}
    old_out = root / old["가림본"] if old.get("가림본") else None
    같은사전 = old.get("사전") == 사전표
    # 짝인 복원 기록이 없어졌거나, 가림본을 누가 고쳤으면(가림본 해시가 다르면) 다시 만든다 — 지난 판정을 그대로 믿지 않게 (강희 지적)
    if (old and old_out and old_out.exists() and 같은사전 and keep_log_for(old_out)
            and old.get("가림해시") and _hash(old_out) == old["가림해시"]):
        same = old.get("크기") == st.st_size and old.get("수정") == st.st_mtime_ns   # 나노초 — 같은 초에 고친 것도 (강희 지적)
        h = None if same else _hash(p)
        if same or old.get("해시") == h:
            if not same:
                old.update(크기=st.st_size, 수정=st.st_mtime_ns)
            if old.get("깨끗") is True:
                result.rows.append(FolderRow("건너뜀", old["가림본"], "바뀐 것 없음"))
            else:                                       # 지난번에 개인정보가 남았던 파일은 다시 돌려도 그대로 경고한다 (강희 지적)
                result.rows.append(FolderRow("개인정보 남음", old["가림본"], "지난번 개인정보·이름 후보 경고 그대로(원본 안 바뀜) — AI에 올리지 마세요"))
            return
    h = _hash(p)
    shown = _masked_rel(rel, dictionary, top_out)
    masked_name = workflows._masked_stem(p, dictionary)[0]
    mode = "xlsx" if p.suffix.lower() in 원래형식 else "md"
    r = workflows.mask_file(p, mode, dictionary)
    if not getattr(r, "ok", False) or not r.ai_path:   # 스캔 PDF 처럼 결과를 못 만든 경우 (강희 지적)
        result.rows.append(FolderRow("못 가림", _safe_rel(rel, dictionary, top_out), _reason(ValueError(r.headline))))
        return
    ai, keep = Path(r.ai_path), Path(r.keep_path) if r.keep_path else None
    # 원래 형식이 아닌 글(md)로 바뀐 파일은 원래 확장자를 이름에 넣는다 — 「상담.pdf」와 「상담.docx」가 같은 「상담_가림.md」로 겹치지 않게
    tag = "" if ai.suffix.lower() == p.suffix.lower() else f"_{p.suffix.lower().lstrip('.')}"
    out_rel = Path(shown).parent / f"{masked_name}{tag}_가림{ai.suffix}"
    쓰임 = {str(v.get("가림본", "")).casefold(): kk for kk, v in 목록.items() if kk != k}   # 대소문자 무시 — 윈도우·맥 기본 디스크는 「A」와 「a」를 같은 파일로 본다 (강희 지적)
    n = 2
    while str(out_rel).casefold() in 쓰임:            # 다른 원본이 같은 이름을 쓰고 있으면(a.pdf·a_pdf.md) 번호를 붙인다 — 덮어쓰지 않게 (강희 지적)
        out_rel = Path(shown).parent / f"{masked_name}{tag}_가림_{n}{ai.suffix}"; n += 1
    # 경로(폴더·파일 이름)를 쓰기 전에 검사한다 — 사전에 없는 이름이 이름에만 있으면 본문 검사로는 못 잡고,
    #   그대로 쓰면 가림본 폴더 이름·목록에 실명이 남는다. 걸리면 가림본에 쓰지 않는다 (강희 지적)
    if _path_bad(str(out_rel), dictionary):
        for t in (ai, keep):
            if t and t.exists():
                t.unlink()                               # 중간 결과(이 컴퓨터 데이터 폴더)도 지운다
        result.rows.append(FolderRow("확인 필요", f"(이름 후보가 든 경로의 {p.suffix.lstrip('.') or '?'} 파일)",   # 이름은 싣지 않는다 (강희 지적)
                                     "폴더·파일 이름에 이름 후보 — 원본 이름을 바꾸거나 사전에 넣은 뒤 다시. 가림본에 쓰지 않았어요"))
        return
    out = root / out_rel
    # 대응표를 먼저 옮기고 가림본은 그다음 — 대응표 저장이 실패하면 옛 가림본·옛 대응표 짝이 그대로 남아 엇갈리지 않는다 (강희 지적)
    if keep and keep.exists():
        log = keep_root / Path(*out_rel.parent.parts[1:])  # 기록 폴더는 최상위 폴더 안 — 그 아래(가린 이름) 경로를 그대로
        log.mkdir(parents=True, exist_ok=True)
        새기록 = log / f"{out.name}_변경기록.xlsx"            # 확장자까지 넣는다 — a.xlsx·a.hwpx 의 기록이 겹치지 않게 (강희 지적)
        임시기록 = 새기록.with_name(새기록.name + ".새것")
        shutil.move(str(keep), str(임시기록))
    else:
        새기록 = 임시기록 = None
    out.parent.mkdir(parents=True, exist_ok=True)
    임시가림 = out.with_name(out.name + ".새것")
    shutil.move(str(ai), str(임시가림))
    옛기록 = 새기록.with_name(새기록.name + ".옛것") if 새기록 and 새기록.exists() else None
    if 옛기록:
        os.replace(새기록, 옛기록)
    try:
        if 임시기록:
            os.replace(임시기록, 새기록)                    # 둘 다 준비된 뒤에 바꿔 끼운다(같은 폴더 안 이름 바꾸기)
        os.replace(임시가림, out)
    except Exception:
        if 옛기록:                                       # 가림본 교체가 실패하면 옛 대응표를 되돌려 옛 가림본과 짝을 지킨다 (강희 지적)
            os.replace(옛기록, 새기록)
        elif 새기록 and 새기록.exists():
            새기록.unlink()
        raise
    if 옛기록 and 옛기록.exists():
        옛기록.unlink()
    if old_out and old_out.exists() and old_out.resolve() != out.resolve():
        old_out.unlink()                                 # 원본이 바뀌어 가린 이름이 달라졌으면 옛 가림본만 지운다(원본은 그대로)
    목록[k] = {"가림본": str(out_rel), "크기": st.st_size, "수정": st.st_mtime_ns, "해시": h, "사전": 사전표, "가림해시": _hash(out),
              "가린날": datetime.now().strftime("%Y-%m-%d %H:%M"), "깨끗": bool(r.clean) and not r.suspects}   # 이름 후보가 있으면 깨끗하지 않다 (강희 지적)
    원본목록[k] = {"원본": str(rel), "가림본": str(out_rel)}
    # 경로(폴더·파일 이름)도 검사한다 — 사전에 없는 이름이 이름에만 있으면 본문 검사로는 못 잡는다 (강희 지적)
    안내 = sorted({str(x[0]) for x in (r.review or []) if x and x[0]})     # 교직원 사전 없음·수식 참조 같은 안내는 종류만 (내용엔 실명이 섞일 수 있다, 강희 지적)
    덧 = f" · 안내 {len(r.review)}건({', '.join(안내)})" if r.review else ""
    if r.clean and r.suspects:
        result.rows.append(FolderRow("확인 필요", str(out_rel), f"이름 후보 {len(r.suspects)}개 — 프로그램에서 파일을 열어 확인한 뒤 올리세요" + 덧))
    elif r.clean:
        result.rows.append(FolderRow("가림", str(out_rel), f"{r.total_changes}곳 가림" + 덧))
    else:
        result.rows.append(FolderRow("개인정보 남음", str(out_rel), f"{len(r.findings)}건 남음 — AI에 올리지 마세요" + 덧))


def _위치파일():
    return Path(workflows.BASE_DIR) / "가림본위치.json"   # 이 컴퓨터의 데이터 폴더 — 가림본 뿌리 ↔ 기록 폴더(경로만)


def _unique_top(root, top, keep_root):
    """가림본 최상위 이름을 기록 폴더마다 하나로 — 이미 다른 기록 폴더가 쓰는 이름이면 기록 위치 해시를 붙인다."""
    keep = str(Path(keep_root).resolve())
    맵 = {}
    for t, v in (_load(_위치파일()).get(str(Path(root).resolve())) or {}).items():   # 대소문자 무시 — 「20_A」·「20_a」는 한 폴더가 된다 (강희 지적)
        if 맵.get(t.casefold()) in (None, keep):
            맵[t.casefold()] = v
    top_cf = top.casefold()
    if 맵.get(top_cf, keep) == keep:
        return top
    h = hashlib.sha256(keep.encode("utf-8")).hexdigest()
    n = 6
    while True:                                          # 해시 붙인 이름도 남이 쓰고 있으면 더 길게, 끝까지 겹치면 번호 (강희 지적)
        붙임 = f"{top}_{h[:n]}" if n <= len(h) else f"{top}_{h}_{n - len(h)}"
        if 맵.get(붙임.casefold(), keep) == keep:                  # 비었거나 지난번 이 기록 폴더가 쓰던 이름
            return 붙임
        n += 2


def _remember_keep(root, top, keep_root):
    """다른 이름의 가림본 폴더를 골랐어도 되돌리기가 기록을 찾게 위치를 남긴다 (강희 지적)."""
    d = _load(_위치파일())
    d.setdefault(str(Path(root).resolve()), {})[top] = str(Path(keep_root).resolve())
    _save(_위치파일(), d)


def _reason(e):
    msg = str(e)
    if "암호" in msg or "password" in msg.lower() or "encrypted" in msg.lower():
        return "암호가 걸린 파일"
    if "Windows" in msg or "한글이 설치" in msg:
        return "한글(.hwp)은 한글이 있는 Windows 에서만"
    if "스캔" in msg or "그림" in msg:
        return "그림(스캔) PDF — 글자가 없어 가릴 수 없음"
    if "markitdown" in msg or "No module named" in msg:
        return "이 컴퓨터에 변환 도구가 없음(설치 필요)"
    return f"처리 오류({type(e).__name__})"           # 그 밖의 오류 글은 싣지 않는다 — 원본 경로·이름이 섞일 수 있다 (강희 지적)


def is_folder_output(masked_path):
    """폴더 가리기로 만든 가림본인가(90_가림본 아래이거나 기록된 가림본 뿌리 아래) — 기록이 없을 때 다른 작업의 기록을 쓰지 않게."""
    p = Path(masked_path).resolve()
    위치 = _load(_위치파일())
    return any(r.name == 가림본이름 or str(r) in 위치 for r in p.parents)


def keep_log_for(masked_path):
    """가림본 파일의 짝이 되는 보관용 기록(실명 대응표)을 찾는다 — 되돌리기용."""
    p = Path(masked_path).resolve()
    위치 = _load(_위치파일())
    for root in p.parents:
        rel = p.relative_to(root)
        if len(rel.parts) < 2:
            continue
        keep_root = (위치.get(str(root)) or {}).get(rel.parts[0])
        if not keep_root and root.name == 가림본이름:
            keep_root = str(root.parent / rel.parts[0] / 기록이름)
        if keep_root:
            keep = Path(keep_root) / Path(*rel.parts[1:-1]) / f"{p.name}_변경기록.xlsx"
            return keep if keep.exists() else None
    return None


def main(argv):
    args = list(argv[1:])
    mask_root = None
    if "--가림본" in args:
        i = args.index("--가림본")
        if i + 1 >= len(args):
            print("사용법: python3 folder_mask.py <원본 폴더> [--가림본 <가림본 뿌리 폴더>]")
            return 1
        mask_root = args[i + 1]
        del args[i:i + 2]
    if len(args) != 1:
        print("사용법: python3 folder_mask.py <원본 폴더> [--가림본 <가림본 뿌리 폴더>]")
        return 1
    try:
        res = mask_folder(Path(args[0]).expanduser(), mask_root,
                          progress=lambda i, n: print(f"\r{i}/{n}", end="", flush=True))
    except Exception as e:
        print(f"\n[오류] {e}")
        return 1
    print()
    for r in res.rows:
        if r.상태 != "건너뜀":
            print(f"{r.상태:8} {r.파일}  {r.내용}")
    print(res.headline)
    print(f"가림본: {res.가림본뿌리}")
    return 0 if res.clean else 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
