"""한글 .hwpx 를 한컴 없이 가명처리하는 어댑터.

.hwpx 는 ZIP + XML 이다. 원본은 읽기만 하고, 결과는 dst 에 새로 쓴다.
- Contents/section*.xml 의 <hp:t> 글자를 치환한다 (표·머리말·꼬리말 포함).
- Preview/PrvText.txt 도 같이 치환한다 (미리보기에 실명이 남지 않게).
- Contents/content.hpf 의 제목/지은이 등 문서 정보는 비운다.
- 한 칸에 이름이 쪼개져 있어도, 이어 붙여 찾아 바꾼다.
"""
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as _xml_escape
from xml.sax.saxutils import unescape as _xml_unescape

from core.checker import check_text
from core.masker import mask_text

from adapters.hwp_adapter import _dedupe_changes

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
_HP_T = re.compile(r"(<hp:t(?:\s[^>]*)?>)([^<]*)(</hp:t>)")
_TEXT_XML = re.compile(r"(?:^|/)Contents/section\d+\.xml$", re.I)


def _xml_files(zf):
    return [n for n in zf.namelist() if _TEXT_XML.search(n.replace("\\", "/"))]


def _collect_hp_t(xml):
    return [_xml_unescape(m.group(2)) for m in _HP_T.finditer(xml)]


def _replace_across(texts, old, new):
    """이어진 글자 목록에서 old 를 찾아 new 로 바꾼다. (칸이 쪼개져 있어도)"""
    if not old:
        return
    while True:
        concat = "".join(texts)
        pos = concat.find(old)
        if pos < 0:
            return
        acc = 0
        start_i = offset = 0
        for i, t in enumerate(texts):
            if acc + len(t) > pos:
                start_i, offset = i, pos - acc
                break
            acc += len(t)
        remaining = old
        i = start_i
        first = True
        while remaining and i < len(texts):
            t = texts[i]
            start = offset if first else 0
            first = False
            chunk = t[start:]
            take = min(len(chunk), len(remaining))
            if chunk[:take] != remaining[:take]:
                return
            if take == len(remaining):
                texts[i] = t[:start] + new + t[start + take:]
                remaining = ""
            else:
                texts[i] = t[:start]
                remaining = remaining[take:]
                i += 1


def _apply_items_to_texts(texts, items):
    for old, new, _kind, _count in items:
        _replace_across(texts, old, new)


def _apply_items_to_xml(xml, items):
    texts = _collect_hp_t(xml)
    if not texts:
        return xml
    _apply_items_to_texts(texts, items)
    parts, last = [], 0
    for m, new_text in zip(_HP_T.finditer(xml), texts):
        parts.append(xml[last:m.start()])
        parts.append(m.group(1) + _xml_escape(new_text) + m.group(3))
        last = m.end()
    parts.append(xml[last:])
    return "".join(parts)


def _apply_items_to_plain(text, items):
    for old, new, _kind, _count in items:
        if old:
            text = text.replace(old, new)
    return text


def _clear_package_meta(xml):
    xml = re.sub(r"(<opf:title>)(.*?)(</opf:title>)", r"\1\3", xml)
    xml = re.sub(
        r'(<opf:meta\b[^>]*\bname="(?:creator|subject|keyword|description)"[^>]*>)(.*?)(</opf:meta>)',
        r"\1\3",
        xml,
    )
    return xml


def _rewrite_zip(src_path, dst_path, updates):
    src_path, dst_path = Path(src_path), Path(dst_path)
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src_path) as zin, zipfile.ZipFile(dst_path, "w") as zout:
        for info in zin.infolist():
            data = updates.get(info.filename, zin.read(info.filename))
            dest = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
            dest.compress_type = info.compress_type
            dest.external_attr = info.external_attr
            zout.writestr(dest, data)


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _p_text(p):
    return "".join("".join(t.itertext()) for t in p.iter(f"{HP}t"))


def _tbl_md(tbl):
    rows = []
    for tr in tbl.findall(f"{HP}tr"):
        cells = []
        for tc in tr.findall(f"{HP}tc"):
            bits = [t for t in (_p_text(p).strip() for p in tc.iter(f"{HP}p")) if t]
            cells.append("<br>".join(bits))
        if cells:
            rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    norm = [r + [""] * (width - len(r)) for r in rows]
    lines = [
        "| " + " | ".join(norm[0]) + " |",
        "| " + " | ".join("---" for _ in range(width)) + " |",
    ]
    for r in norm[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


def _walk_md(el, out):
    tag = _local(el.tag)
    if tag == "tbl":
        md = _tbl_md(el)
        if md:
            out.append(md)
        return
    if tag == "p":
        if next(el.iter(f"{HP}tbl"), None) is not None:
            for child in el:
                _walk_md(child, out)
            return
        text = _p_text(el).strip()
        if text:
            out.append(text)
        return
    for child in el:
        _walk_md(child, out)


def hwpx_to_markdown(path):
    """.hwpx 본문·표를 md 글로 바꾼다 (한컴 불필요)."""
    parts = []
    with zipfile.ZipFile(path) as zf:
        for name in _xml_files(zf):
            root = ET.fromstring(zf.read(name))
            chunk = []
            _walk_md(root, chunk)
            if chunk:
                parts.extend(chunk)
    return "\n\n".join(parts).strip() + ("\n" if parts else "")


def hwpx_text(path):
    """문서의 모든 <hp:t> 글자와 미리보기 글을 이은 문자열."""
    chunks = []
    with zipfile.ZipFile(path) as zf:
        for name in _xml_files(zf):
            xml = zf.read(name).decode("utf-8")
            texts = [t for t in _collect_hp_t(xml) if t]
            if texts:
                chunks.append("\n".join(texts))
        if "Preview/PrvText.txt" in zf.namelist():
            prv = zf.read("Preview/PrvText.txt").decode("utf-8")
            if prv.strip():
                chunks.append(prv)
    return "\n".join(chunks)


def mask_hwpx(src_path, dst_path, dictionary, confirm_duplicate=None):
    """원본 .hwpx 를 읽어 서식 그대로 가명처리한 뒤 dst 에 저장한다."""
    src_path, dst_path = Path(src_path), Path(dst_path)
    text = hwpx_text(src_path)
    _, raw_changes = mask_text(text, dictionary)
    items = _dedupe_changes(raw_changes)
    duplicates = any(kind == "동명이인확인" for _, _, kind, _ in items)
    if duplicates and confirm_duplicate is not None and not confirm_duplicate():
        return {"status": "aborted", "changes": items, "findings": [], "duplicates": duplicates}

    updates = {}
    with zipfile.ZipFile(src_path) as zf:
        for name in _xml_files(zf):
            xml = zf.read(name).decode("utf-8")
            updates[name] = _apply_items_to_xml(xml, items).encode("utf-8")
        if "Preview/PrvText.txt" in zf.namelist():
            prv = zf.read("Preview/PrvText.txt").decode("utf-8")
            updates["Preview/PrvText.txt"] = _apply_items_to_plain(prv, items).encode("utf-8")
        if "Contents/content.hpf" in zf.namelist():
            meta = zf.read("Contents/content.hpf").decode("utf-8")
            updates["Contents/content.hpf"] = _clear_package_meta(
                _apply_items_to_plain(meta, items)
            ).encode("utf-8")

    _rewrite_zip(src_path, dst_path, updates)
    findings = check_text(hwpx_text(dst_path), dictionary)
    return {"status": "done", "changes": items, "findings": findings, "duplicates": duplicates}


def apply_extra_names_hwpx(path, words):
    """가명처리된 .hwpx 에서 사용자가 고른 단어를 [이름N] 으로 바꾼다.

    (변경 목록 [(원래, 바뀐, 종류, 횟수), ...], {코드: 단어}) 를 돌려준다.
    """
    uniq = list(dict.fromkeys(w for w in words if w))
    name_map = {f"[이름{i}]": w for i, w in enumerate(uniq, start=1)}
    if not name_map:
        return [], {}
    items = [(w, c, "이름의심확정", 1) for c, w in name_map.items()]
    items.sort(key=lambda t: (-len(t[0]), t[0]))
    path = Path(path)
    updates = {}
    with zipfile.ZipFile(path) as zf:
        for name in _xml_files(zf):
            xml = zf.read(name).decode("utf-8")
            updates[name] = _apply_items_to_xml(xml, items).encode("utf-8")
        if "Preview/PrvText.txt" in zf.namelist():
            prv = zf.read("Preview/PrvText.txt").decode("utf-8")
            updates["Preview/PrvText.txt"] = _apply_items_to_plain(prv, items).encode("utf-8")
    tmp = path.with_suffix(path.suffix + ".tmp")
    _rewrite_zip(path, tmp, updates)
    tmp.replace(path)

    text = hwpx_text(path)
    counted = []
    for old, new, kind, _ in items:
        n = text.count(new)
        if n:
            counted.append((old, new, kind, n))
    return counted, name_map
