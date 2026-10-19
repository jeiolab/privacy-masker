#!/usr/bin/env python3
"""개인정보 가리개 - 화면.   실행: python3 app.py

화면은 이 파일, 실제 작업은 workflows.py 가 core/·adapters/ 의 함수를 불러서 한다.
- 가리기: 왼쪽에서 파일·방식을 고르고, 오른쪽에서 결과를 확인한다.
- 되돌리기: 왼쪽에 AI 답변을 붙여 넣고, 오른쪽에서 복원 결과를 본다.
- 어떤 오류가 나도 프로그램이 꺼지지 않고, 알림창으로 오류 내용을 보여 준다.
"""
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog
from tkinter import font as tkfont
from tkinter import messagebox, ttk

import workflows
from core.restorer import unmask_text
from core.updater import fetch_update_info, is_newer, open_download_page, update_how_to
from core.version import VERSION

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:   # tkinterdnd2 가 없으면 끌어다 놓기만 빠지고 나머지는 그대로 동작
    HAS_DND = False

APP_TITLE = "개인정보 가리개"
WINDOW_SIZE = "1080x800"
MIN_SIZE = (920, 700)

# 맥·윈도우·리눅스에서 한글이 잘 보이는 글꼴 후보 (앞에서부터 설치된 것을 쓴다)
FONT_CANDIDATES = [
    "Apple SD Gothic Neo", "AppleGothic",       # 맥
    "맑은 고딕", "Malgun Gothic",                 # 윈도우
    "Noto Sans CJK KR", "NanumGothic",           # 리눅스
]

# ---- 색 (DESIGN.md 의 색 이름을 그대로 썼다) ----
MIST = "#F5F5F7"        # 화면 바탕
WHITE = "#FFFFFF"       # 카드
INK = "#1D1D1F"         # 본문 글자
SLATE = "#6E6E73"       # 보조 글자
STEEL = "#86868B"       # 아이콘·테두리
HAIR = "#D6D6D6"        # 가는 선
CONTROL = "#E6E6E8"     # 세그먼트 바탕·비활성
ACTION = "#0071E3"      # 주 버튼
ACTION_HOVER = "#0077ED"
ACTION_PRESS = "#0062C4"
DROP_BG = "#F5F9FF"
DROP_HOVER = "#E3EFFD"
DROP_BORDER = "#9DBBE4"
SELECT_BG = "#BBD7FB"

# 상태 줄 색: (글자, 바탕)
STATUS_COLORS = {
    "ok": ("#1E7B3A", "#E7F5EB"),
    "warn": ("#B64400", "#FFF2E6"),
    "bad": ("#C62828", "#FDECEE"),
    "neutral": (SLATE, MIST),
}

# 미리보기에서 색을 다르게 보여 줄 코드들 (태그 이름, 정규식, 글자색, 배경색)
CODE_STYLES = [
    ("code_unknown", r"\[(?:전화번호|주민번호|이메일)\]", "#B71C1C", "#FDE7E7"),
    ("code_ename", r"\[이름\d+\]", "#0F766E", "#D7F2EC"),
    ("code_bracket", r"\[(?:학생전화|보호자전화|주소|생년월일)\d{3,}(?:또는\d{3,})*\]", "#6A1B9A", "#F3E5F9"),
    ("code_sid", r"S\d{3,}(?:또는\d{3,})*", "#B45309", "#FEF0DB"),
    ("code_name", r"(?:학생|보호자|교사)\d{3,}(?:또는\d{3,})*", "#1D4ED8", "#E3ECFD"),
]
CODE_REGEX = re.compile("|".join(f"(?P<{tag}>{pat})" for tag, pat, _, _ in CODE_STYLES))

NO_DICT_MESSAGE = "학생사전 없음 - 먼저 사전을 만들어 주세요"
INPUT_HINT = "여기를 눌러 AI 답변을 붙여 넣으세요 (붙여넣기: Cmd+V / Ctrl+V)"
OUTPUT_HINT = "[복원]을 누르면 실명으로 되돌린 글이 여기에 나와요."
EXAMPLE_PREVIEW = "학생003은 S003이며 연락처는 [학생전화003]입니다. 보호자003과 상담함."
MASK_FILETYPES = [("지원하는 파일", "*.pdf *.docx *.xlsx *.xls *.pptx *.csv *.txt *.md *.hwp *.hwpx"),
                  ("모든 파일", "*.*")]
HAND_CURSOR = "pointinghand" if sys.platform == "darwin" else "hand2"


def open_folder(path):
    """폴더를 연다 (맥·윈도우·리눅스)."""
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    if sys.platform.startswith("win"):
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=True)
    else:
        subprocess.run(["xdg-open", str(path)], check=True)


def split_warning(text):
    """'[위치: ]N번째 줄, 코드: 이유' 모양의 경고를 (코드, 이유)로 나눈다."""
    parts = text.split(": ")
    if len(parts) < 2:
        return "", text
    reason, code, where = parts[-1], parts[-2], ": ".join(parts[:-2])
    m = re.match(r"(\d+번째 줄), (.*)", code)
    if m:
        where = f"{where} {m.group(1)}".strip()
        code = m.group(2)
    return code, f"[{where}] {reason}" if where else reason


def pick_font_family(root):
    available = set(tkfont.families(root))
    for name in FONT_CANDIDATES:
        if name in available:
            return name
    return "TkDefaultFont"


# ====================================================================== 화면 부품
def _solid_round(canvas, x1, y1, x2, y2, r, color):
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    d = 2 * r
    kw = dict(fill=color, outline="")
    canvas.create_oval(x1, y1, x1 + d, y1 + d, **kw)
    canvas.create_oval(x2 - d, y1, x2, y1 + d, **kw)
    canvas.create_oval(x1, y2 - d, x1 + d, y2, **kw)
    canvas.create_oval(x2 - d, y2 - d, x2, y2, **kw)
    canvas.create_rectangle(x1 + r, y1, x2 - r, y2, **kw)
    canvas.create_rectangle(x1, y1 + r, x2, y2 - r, **kw)


def round_rect(canvas, x1, y1, x2, y2, r, fill, outline=None):
    """모서리가 둥근 사각형 (r 이 높이의 절반이면 알약 모양). 테두리는 바깥 모양 위에 안쪽 모양을 겹쳐 만든다."""
    h = y2 - y1
    if abs(r - h / 2) < 1 and x2 - x1 > h:   # 알약 모양
        def pill(inset, color):
            canvas.create_line(x1 + h / 2, (y1 + y2) / 2, x2 - h / 2, (y1 + y2) / 2, width=h - 2 * inset,
                               capstyle="round", fill=color)
        if outline and outline != fill:
            pill(0, outline)
            pill(1, fill)
        else:
            pill(0, fill)
    elif outline and outline != fill:
        _solid_round(canvas, x1, y1, x2, y2, r, outline)
        _solid_round(canvas, x1 + 1, y1 + 1, x2 - 1, y2 - 1, r - 1, fill)
    else:
        _solid_round(canvas, x1, y1, x2, y2, r, fill)


class Card(tk.Frame):
    """둥근 모서리 카드. 내용은 card.inner 안에 넣는다.

    뒤에 둥근 배경을 그리는 캔버스를 깔고, 그 위에 안쪽 틀을 여백만큼 띄워 얹는다.
    (안쪽 내용 크기에 따라 카드 크기가 저절로 맞춰지고, 남는 자리는 안쪽 내용이 채운다.)
    """

    def __init__(self, parent, parent_bg=MIST, fill=WHITE, radius=20, pad=22, expand=False, border=None):
        super().__init__(parent, bg=parent_bg, bd=0, highlightthickness=0)
        self.fill, self.radius, self.pad, self.border = fill, radius, pad, border
        self._bg = tk.Canvas(self, bg=parent_bg, highlightthickness=0, bd=0)
        self._bg.place(x=0, y=0, relwidth=1, relheight=1)
        self._bg.bind("<Configure>", self._draw)
        self.inner = tk.Frame(self, bg=fill)
        self.inner.pack(fill="both", expand=True, padx=pad, pady=pad)

    def _draw(self, event):
        self._bg.delete("all")
        round_rect(self._bg, 0, 0, event.width, event.height, self.radius, fill=self.fill, outline=self.border)


class PillButton(tk.Canvas):
    """알약 모양 버튼. kind: primary(파랑, 화면에서 가장 중요한 버튼) / outline(테두리만)."""

    KINDS = {
        "primary": dict(fill=ACTION, hover=ACTION_HOVER, press=ACTION_PRESS, fg=WHITE, outline=None),
        "outline": dict(fill=WHITE, hover=MIST, press=CONTROL, fg=INK, outline=STEEL),
    }

    def __init__(self, parent, text, command, font, kind="outline", height=40, padx=22, parent_bg=WHITE):
        self._font = font
        self._text = text
        self._command = command
        self._kind = kind
        self._disabled = False
        self._hover = False
        self._pressed = False
        width = font.measure(text) + 2 * padx
        super().__init__(parent, width=width, height=height, bg=parent_bg, highlightthickness=0, bd=0)
        try:
            self.configure(cursor=HAND_CURSOR)
        except tk.TclError:
            pass
        self.bind("<Configure>", lambda _e: self._redraw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self._redraw()

    # ttk 버튼처럼 쓸 수 있게 맞춘 부분
    def state(self, flags):
        for f in flags:
            if f == "disabled":
                self._disabled = True
            elif f == "!disabled":
                self._disabled = False
        self._hover = self._pressed = False
        try:
            self.configure(cursor="" if self._disabled else HAND_CURSOR)
        except tk.TclError:
            pass
        self._redraw()

    def instate(self, flags):
        for f in flags:
            if f == "disabled":
                return self._disabled
            if f == "!disabled":
                return not self._disabled
        return False

    def configure(self, cnf=None, **kw):
        if "text" in kw:
            self._text = kw.pop("text")
            self._redraw()
        if cnf or kw:
            return super().configure(cnf, **kw)
        return None

    config = configure

    def cget(self, key):
        return self._text if key == "text" else super().cget(key)

    def set_parent_bg(self, color):
        super().configure(bg=color)

    # 그리기
    def _redraw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1 or h <= 1:
            w, h = int(float(self["width"])), int(float(self["height"]))
        style = self.KINDS[self._kind]
        if self._disabled:
            fill, fg, outline = CONTROL, STEEL, None
        else:
            fill = style["press"] if self._pressed else style["hover"] if self._hover else style["fill"]
            fg, outline = style["fg"], style["outline"]
        round_rect(self, 0, 0, w, h, h / 2, fill=fill, outline=outline)
        self.create_text(w / 2, h / 2, text=self._text, font=self._font, fill=fg)

    def _on_enter(self, _e):
        if not self._disabled:
            self._hover = True
            self._redraw()

    def _on_leave(self, _e):
        self._hover = self._pressed = False
        self._redraw()

    def _on_press(self, _e):
        if not self._disabled:
            self._pressed = True
            self._redraw()

    def _on_release(self, e):
        was_pressed = self._pressed
        self._pressed = False
        self._redraw()
        inside = 0 <= e.x < self.winfo_width() and 0 <= e.y < self.winfo_height()
        if was_pressed and inside and not self._disabled:
            self._command()


class Segmented(tk.Canvas):
    """알약 모양 선택 스위치 (라디오 버튼 대신). options: [(값, 글자, 작은 꼬리표 또는 None)]"""

    def __init__(self, parent, options, variable, font, tag_font, height=40, parent_bg=WHITE,
                 min_seg=0, command=None):
        self.options, self.var, self.command = options, variable, command
        self.font, self.tag_font, self.pad = font, tag_font, 3
        self.disabled = set()
        seg = max([font.measure(label) + (tag_font.measure(tag) + 6 if tag else 0) + 40
                   for _, label, tag in options] + [min_seg])
        self.seg_w = seg
        super().__init__(parent, width=seg * len(options) + 2 * self.pad, height=height, bg=parent_bg,
                         highlightthickness=0, bd=0)
        try:
            self.configure(cursor=HAND_CURSOR)
        except tk.TclError:
            pass
        self.bind("<Configure>", lambda _e: self._redraw())
        self.bind("<ButtonRelease-1>", self._on_click)
        variable.trace_add("write", lambda *_: self._redraw())
        self._redraw()

    def set_enabled(self, value, enabled):
        (self.disabled.discard if enabled else self.disabled.add)(value)
        self._redraw()

    def _redraw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1 or h <= 1:
            w, h = int(float(self["width"])), int(float(self["height"]))
        n, pad = len(self.options), self.pad
        sw = (w - 2 * pad) / n
        round_rect(self, 0, 0, w, h, h / 2, fill=CONTROL)
        for i, (value, label, tag) in enumerate(self.options):
            selected = self.var.get() == value
            off = value in self.disabled
            x0 = pad + i * sw
            if selected:
                round_rect(self, x0, pad, x0 + sw, h - pad, (h - 2 * pad) / 2, fill=WHITE, outline="#D2D2D6")
            fg = STEEL if off else INK
            font = self.font
            lw = font.measure(label)
            tw = self.tag_font.measure(tag) + 6 if tag else 0
            start = x0 + sw / 2 - (lw + tw) / 2
            self.create_text(start, h / 2, text=label, font=font, fill=fg, anchor="w")
            if tag:
                self.create_text(start + lw + 6, h / 2, text=tag, font=self.tag_font,
                                 fill=STEEL if off else "#B64400", anchor="w")

    def _on_click(self, e):
        n = len(self.options)
        sw = (self.winfo_width() - 2 * self.pad) / n
        i = int((e.x - self.pad) // sw)
        if 0 <= i < n and 0 <= e.y <= self.winfo_height():
            value = self.options[i][0]
            if value in self.disabled:
                return
            if self.var.get() != value:
                self.var.set(value)
                if self.command:
                    self.command(value)


# ====================================================================== 앱
class App:
    def __init__(self):
        self.root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*MIN_SIZE)
        self.root.configure(bg=MIST)

        self.family = pick_font_family(self.root)
        self._setup_fonts_and_styles()

        self.file_var = tk.StringVar(value="아직 선택한 파일이 없어요")
        self.mode_var = tk.StringVar(value="md")
        self.tab_var = tk.StringVar(value="mask")
        self.selected_path = None
        self.result = None
        self._confirmed_names = []      # '이 단어도 가리기'로 고른 이름 단어 (순서 유지)
        self._ignored_names = set()     # '무시'한 이름 의심 단어

        self._build_top_bar()
        self._build_tab_switch()
        self.content = tk.Frame(self.root, bg=MIST)
        self.content.pack(fill="both", expand=True, padx=24, pady=(0, 22))
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)
        self.mask_tab = tk.Frame(self.content, bg=MIST)
        self.restore_tab = tk.Frame(self.content, bg=MIST)
        for tab in (self.mask_tab, self.restore_tab):   # 두 화면을 같은 자리에 겹쳐 두고 앞으로만 꺼낸다
            tab.grid(row=0, column=0, sticky="nsew")
        self._build_mask_tab()
        self._build_restore_tab()
        self.show_tab("mask")

        self.root.report_callback_exception = self._on_callback_error   # 어떤 오류도 알림창으로
        self.refresh_dictionary()
        self.root.after(600, self._startup_migrate_and_cleanup)

    # ------------------------------------------------------------ 글꼴·스타일
    def _setup_fonts_and_styles(self):
        fam = self.family
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            tkfont.nametofont(name).configure(family=fam, size=13)
        self.f_body = tkfont.Font(family=fam, size=13)
        self.f_bold = tkfont.Font(family=fam, size=13, weight="bold")
        self.f_small = tkfont.Font(family=fam, size=12)
        self.f_small_bold = tkfont.Font(family=fam, size=12, weight="bold")
        self.f_tag = tkfont.Font(family=fam, size=11, weight="bold")
        self.f_title = tkfont.Font(family=fam, size=16, weight="bold")
        self.f_app = tkfont.Font(family=fam, size=17, weight="bold")
        self.f_tab = tkfont.Font(family=fam, size=14, weight="bold")
        self.f_btn = tkfont.Font(family=fam, size=13, weight="bold")
        self.f_primary = tkfont.Font(family=fam, size=16, weight="bold")
        self.f_text = tkfont.Font(family=fam, size=14)
        self.f_text_bold = tkfont.Font(family=fam, size=14, weight="bold")

        style = ttk.Style(self.root)
        style.theme_use("clam")   # 맥·윈도우에서 같은 모양으로 보이게
        style.configure("Treeview", font=self.f_body, rowheight=32, background=WHITE, fieldbackground=WHITE,
                        foreground=INK, borderwidth=1, relief="flat", bordercolor=HAIR, lightcolor=WHITE,
                        darkcolor=WHITE)
        style.configure("Treeview.Heading", font=self.f_small_bold, background=MIST, foreground=SLATE,
                        relief="flat", borderwidth=0, padding=(10, 7))
        style.map("Treeview", background=[("selected", DROP_HOVER)], foreground=[("selected", INK)])
        style.map("Treeview.Heading", background=[("active", MIST)])
        style.layout("Slim.Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {
            "sticky": "ns", "children": [("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        style.configure("Slim.Vertical.TScrollbar", background=HAIR, troughcolor=MIST, bordercolor=MIST,
                        lightcolor=HAIR, darkcolor=HAIR, relief="flat", width=10)
        style.map("Slim.Vertical.TScrollbar", background=[("active", STEEL)])

    # ------------------------------------------------------------ 공통 조각
    def button(self, parent, text, command, kind="outline", height=40, padx=22, parent_bg=WHITE, font=None):
        return PillButton(parent, text, self.safe(command), font or self.f_btn, kind=kind, height=height,
                          padx=padx, parent_bg=parent_bg)

    def card_title(self, parent, number, text):
        row = tk.Frame(parent, bg=WHITE)
        badge = tk.Canvas(row, width=26, height=26, bg=WHITE, highlightthickness=0, bd=0)
        badge.create_oval(1, 1, 25, 25, fill=INK, outline=INK)
        badge.create_text(13, 13, text=number, fill=WHITE, font=self.f_small_bold)
        badge.pack(side="left")
        tk.Label(row, text=text, font=self.f_title, bg=WHITE, fg=INK).pack(side="left", padx=(10, 0))
        return row

    def make_text(self, parent, readonly=False):
        """연한 회색 둥근 글상자 + 얇은 스크롤바. (box, text) 를 돌려준다."""
        box = Card(parent, parent_bg=WHITE, fill=MIST, radius=16, pad=4, expand=True)
        text = tk.Text(box.inner, wrap="word", font=self.f_text, bg=MIST, fg=INK, relief="flat", bd=0,
                       highlightthickness=0, padx=14, pady=10, spacing1=2, spacing3=2, insertbackground=INK,
                       selectbackground=SELECT_BG, selectforeground=INK, undo=not readonly, width=10, height=4)
        bar = ttk.Scrollbar(box.inner, orient="vertical", command=text.yview, style="Slim.Vertical.TScrollbar")
        text.configure(yscrollcommand=self.autohide(bar, text))
        text.pack(side="left", fill="both", expand=True)
        return box, text

    @staticmethod
    def autohide(bar, widget):
        """스크롤바는 내용이 넘칠 때만 보인다 (평소에는 숨겨서 화면을 깔끔하게)."""
        def update(lo, hi):
            if float(lo) <= 0.0 and float(hi) >= 1.0:
                bar.pack_forget()
            elif not bar.winfo_ismapped():
                bar.pack(side="right", fill="y", before=widget)
            bar.set(lo, hi)
        return update

    def make_table(self, parent, columns, height):
        """(종류, 위치, 내용) 같은 열이 있는 목록. (frame, tree) 를 돌려준다."""
        frame = tk.Frame(parent, bg=WHITE)
        tree = ttk.Treeview(frame, columns=[c for c, _, _ in columns], show="headings", height=height,
                            selectmode="browse")
        for name, width, stretch in columns:
            tree.heading(name, text=name, anchor="w")
            tree.column(name, width=width, minwidth=60, anchor="w", stretch=stretch)
        tree._max_rows = height
        bar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview, style="Slim.Vertical.TScrollbar")
        tree.configure(yscrollcommand=self.autohide(bar, tree))
        tree.pack(side="left", fill="both", expand=True)
        return frame, tree

    def set_text(self, text_widget, content, highlight=False, placeholder=None):
        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")
        text_widget.insert("1.0", content)
        if highlight:
            self.highlight_codes(text_widget)
        if not content and placeholder:
            text_widget.tag_configure("ph", foreground=STEEL)
            text_widget.insert("1.0", placeholder, "ph")
        text_widget.configure(state="disabled")

    @staticmethod
    def attach_placeholder(text, message):
        """비어 있을 때만 보이는 안내 글 (글상자를 누르면 사라진다)."""
        text._ph_on = False
        text.tag_configure("ph", foreground=STEEL)

        def show(_e=None):
            if not text.get("1.0", "end-1c"):
                text.insert("1.0", message, "ph")
                text._ph_on = True

        def hide(_e=None):
            if text._ph_on:
                text.delete("1.0", "end")
                text._ph_on = False
        text.bind("<FocusIn>", hide)
        text.bind("<FocusOut>", show)
        show()

    def highlight_codes(self, text_widget):
        """학생003, S003, [학생전화003] 같은 코드에 색을 입힌다."""
        for tag, _, fg, bg in CODE_STYLES:
            text_widget.tag_configure(tag, foreground=fg, background=bg, font=self.f_text_bold)
            text_widget.tag_remove(tag, "1.0", "end")
        for line_no, line in enumerate(text_widget.get("1.0", "end-1c").split("\n"), start=1):
            for m in CODE_REGEX.finditer(line):
                text_widget.tag_add(m.lastgroup, f"{line_no}.{m.start()}", f"{line_no}.{m.end()}")

    _SEVERITY = {"neutral": 0, "ok": 1, "warn": 2, "bad": 3}   # 빨강(bad) > 주황(warn) > 초록(ok)

    def set_status(self, text, kind="neutral"):
        """결과 상태 한 덩어리(모두 같은 색). 여러 줄이면 전부 같은 kind 색으로 보여 준다."""
        self._render_status([(line, kind) for line in text.split("\n")], kind)

    def _render_status(self, lines, box_kind, path=None):
        """줄마다 (글, kind)로 색을 달리 칠한다. 상자 바탕색은 box_kind(가장 심각한 상태) 기준.
        path 가 있으면 상자 아래에 작은 회색 글씨로 저장 위치를 한 줄 보여 준다."""
        box_bg = STATUS_COLORS[box_kind][1]
        self.status_box.configure(bg=box_bg)
        self.status_inner.configure(bg=box_bg)
        for w in self.status_inner.winfo_children():
            w.destroy()
        self._status_line_labels = []
        for text, kind in lines:
            lbl = tk.Label(self.status_inner, text=text, font=self.f_bold, bg=box_bg,
                           fg=STATUS_COLORS[kind][0], anchor="w", justify="left")
            lbl.pack(fill="x", anchor="w")
            self._status_line_labels.append(lbl)
        if path:
            self.status_path.configure(text=path)
            self.status_path.pack(fill="x", pady=(4, 0))
        else:
            self.status_path.pack_forget()
        self._reflow_status()

    def _reflow_status(self):
        w = self.status_inner.winfo_width()
        if w <= 1:
            return
        for lbl in self._status_line_labels:
            lbl.configure(wraplength=max(60, w - 4))

    @staticmethod
    def _split_headline(headline):
        """'… (저장: 경로)' 에서 (상태글, '저장: 경로') 를 나눈다. 경로가 없으면 (원문, None)."""
        m = re.search(r"\s*\((저장:[^()]*)\)\s*$", headline)
        if m:
            return headline[:m.start()].rstrip(), m.group(1)
        return headline, None

    # ------------------------------------------------------------ 토스트(잠깐 떴다 사라지는 알림)
    def toast(self, message, ms=3200):
        """화면 아래쪽에 둥근 알림을 잠깐 띄웠다가 스스로 사라진다. 실패해도 프로그램에 영향 없음."""
        try:
            tip = tk.Toplevel(self.root)
            tip.overrideredirect(True)
            tip.configure(bg=MIST)
            for attr in (("-topmost", True), ("-transparentcolor", MIST)):
                try:
                    tip.attributes(*attr)      # 투명색은 Windows 전용, 안 되면 그냥 넘어감
                except tk.TclError:
                    pass
            pad_x, pad_y = 22, 14
            w = self.f_bold.measure(message) + 2 * pad_x
            h = self.f_bold.metrics("linespace") + 2 * pad_y
            cv = tk.Canvas(tip, width=w, height=h, bg=MIST, highlightthickness=0, bd=0)
            cv.pack()
            round_rect(cv, 1, 1, w - 1, h - 1, h / 2, fill=INK)
            cv.create_text(w / 2, h / 2, text=message, font=self.f_bold, fill=WHITE)
            self.root.update_idletasks()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            x = rx + (rw - w) // 2
            y = ry + rh - h - 48
            tip.geometry(f"{int(w)}x{int(h)}+{max(0, int(x))}+{max(0, int(y))}")
            tip.after(ms, tip.destroy)
        except tk.TclError:
            pass

    # ------------------------------------------------------------ output 폴더 정리
    def _startup_migrate_and_cleanup(self):
        """처음 켤 때: 예전 프로그램 폴더의 data/output 을 문서 아래로 옮긴 뒤, 오래된 결과를 정리한다."""
        from core import config
        try:
            moved = config.migrate_legacy_data()
        except Exception:
            moved = None
        if moved:
            messagebox.showinfo(APP_TITLE, moved)
            self.refresh_dictionary()
        self._startup_cleanup()
        self.root.after(800, lambda: self.check_update(manual=False))

    def _startup_cleanup(self):
        """프로그램 시작 때: 오래된 결과 파일을 지우고, 지운 게 있으면 토스트로 알린다."""
        try:
            removed = workflows.cleanup_old_outputs()
        except Exception:
            removed = 0                       # 정리 실패해도 프로그램은 정상 실행
        if removed:
            self.toast(f"오래된 결과 파일 {removed}개를 정리했어요")

    # ------------------------------------------------------------ 업데이트 확인
    def check_update(self, manual=False):
        """version.json 만 GET 으로 읽는다. 실패하면 조용히 넘어간다. (오류창 없음)"""
        def work():
            info = fetch_update_info()
            try:
                self.root.after(0, lambda: self._on_update_info(info, manual))
            except tk.TclError:
                pass

        threading.Thread(target=work, daemon=True).start()

    def _on_update_info(self, info, manual):
        if not info or not is_newer(info.get("version", "")):
            return   # 최신이거나 못 읽음 → 오류창 없이 넘어감 (켤 때·클릭 모두)
        self._show_update_dialog(info)

    def _show_update_dialog(self, info):
        """새 버전이 있을 때만: notes 와 [다운로드 페이지 열기] [나중에]."""
        win = tk.Toplevel(self.root)
        win.title(APP_TITLE)
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text=f"새 버전 {info['version']}이 있어요", font=self.f_title,
                 bg=WHITE, fg=INK).pack(anchor="w")
        note = info.get("notes") or ""
        if info.get("date"):
            note = f"{info['date']}\n{note}".strip()
        if note:
            tk.Label(box, text=note, font=self.f_small, bg=WHITE, fg=SLATE,
                     justify="left", wraplength=400, anchor="w").pack(anchor="w", pady=(8, 0))
        tk.Label(box, text=update_how_to(info.get("download")), font=self.f_small, bg=WHITE, fg=INK,
                 justify="left", wraplength=400, anchor="w").pack(anchor="w", pady=(12, 16))

        row = tk.Frame(box, bg=WHITE)
        row.pack(fill="x")

        def open_page():
            open_download_page(info.get("download"))
            win.destroy()

        PillButton(row, "다운로드 페이지 열기", open_page, self.f_btn, kind="primary",
                   height=40, padx=18, parent_bg=WHITE).pack(side="left")
        PillButton(row, "나중에", win.destroy, self.f_btn, kind="outline",
                   height=40, padx=18, parent_bg=WHITE).pack(side="left", padx=(10, 0))

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - w) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - h) // 3)
        win.geometry(f"+{x}+{y}")
        win.grab_set()

    def clear_outputs_now(self):
        """[지금 모두 정리] 버튼: 확인 후 output/ 의 결과 파일을 전부 비운다. (data/ 는 건드리지 않음)"""
        if not messagebox.askyesno(
                APP_TITLE,
                "output 폴더의 결과 파일(AI용·보관용·복원)을 모두 지웁니다.\n"
                "사전과 백업이 든 data 폴더는 건드리지 않아요.\n\n계속할까요?",
                icon="warning"):
            return
        removed = workflows.clear_all_outputs()
        self.toast(f"결과 파일 {removed}개를 모두 정리했어요" if removed else "지울 결과 파일이 없어요")

    # ------------------------------------------------------------ 맨 위 줄
    def _build_top_bar(self):
        bar = tk.Frame(self.root, bg=WHITE)
        bar.pack(fill="x")
        row = tk.Frame(bar, bg=WHITE)
        row.pack(fill="x", padx=24, pady=10)
        tk.Label(row, text=APP_TITLE, font=self.f_app, bg=WHITE, fg=INK).pack(side="left")
        self.version_label = tk.Label(row, text=f"v{VERSION}", font=self.f_small, bg=WHITE, fg=STEEL,
                                      cursor=HAND_CURSOR)
        self.version_label.pack(side="right")
        self.version_label.bind("<Button-1>", lambda _e: self.check_update(manual=True))
        self.button(row, "사전 업데이트", self.build_dictionary, height=36, padx=18).pack(
            side="right", padx=(0, 12))
        self.button(row, "무시 단어", self.show_ignored_words, height=36, padx=14).pack(
            side="right", padx=(0, 8))
        self.button(row, "지금 모두 정리", self.clear_outputs_now, height=36, padx=18).pack(
            side="right", padx=(0, 8))
        self.dict_label = tk.Label(row, text="", font=self.f_small, bg=WHITE, fg=SLATE)
        self.dict_label.pack(side="right", padx=(0, 16))
        self.dict_dot = tk.Label(row, text="●", font=self.f_small, bg=WHITE, fg=STEEL)
        self.dict_dot.pack(side="right", padx=(0, 6))
        tk.Frame(bar, bg=HAIR, height=1).pack(fill="x")

    def _build_tab_switch(self):
        row = tk.Frame(self.root, bg=MIST)
        row.pack(fill="x", pady=(16, 14))
        Segmented(row, [("mask", "가리기", None), ("restore", "되돌리기", None)], self.tab_var, self.f_tab,
                  self.f_tag, height=42, parent_bg=MIST, min_seg=150, command=self.show_tab).pack()

    def show_tab(self, which):
        (self.mask_tab if which == "mask" else self.restore_tab).tkraise()

    # ------------------------------------------------------------ 탭 1: 가리기
    def _build_mask_tab(self):
        tab = self.mask_tab
        tab.grid_columnconfigure(0, weight=2, uniform="mask")
        tab.grid_columnconfigure(1, weight=3, uniform="mask")
        tab.grid_rowconfigure(0, weight=1)
        left = tk.Frame(tab, bg=MIST)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right = tk.Frame(tab, bg=MIST)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        # ---- 왼쪽: 1 파일 → 2 방식 → 실행
        file_card = Card(left)
        file_card.pack(fill="x")
        box = file_card.inner
        self.card_title(box, "1", "파일 선택").pack(anchor="w")

        # 파일 선택 상자: 높이 180 고정, 평소엔 회색 점선 테두리. 끌어다 놓는 중에만 파란색으로.
        self._drop_border_color = STEEL
        self.drop = tk.Canvas(box, height=180, bg=WHITE, highlightthickness=0, bd=0)
        self.drop.pack(fill="x", pady=(14, 12))
        self.drop.bind("<Configure>", lambda _e: self._draw_drop_border())
        inner = tk.Frame(self.drop, bg=WHITE)
        inner.place(relx=0.5, rely=0.5, anchor="center")
        self.drop_icon = tk.Canvas(inner, width=48, height=56, bg=WHITE, highlightthickness=0, bd=0)
        self.drop_icon.create_polygon(8, 4, 30, 4, 42, 16, 42, 52, 8, 52, outline=STEEL, fill=WHITE, width=2)
        self.drop_icon.create_line(30, 4, 30, 16, 42, 16, fill=STEEL, width=2)
        self.drop_icon.create_line(15, 28, 35, 28, 15, 36, 35, 36, 15, 44, 28, 44, fill=STEEL, width=2)
        self.drop_icon.pack()
        hint = "파일을 여기로 끌어다 놓으세요" if HAS_DND else "파일을 여기로 끌어다 놓으세요 (tkinterdnd2 없음)"
        self.drop_label = tk.Label(inner, text=hint, font=self.f_bold, bg=WHITE, fg=INK)
        self.drop_label.pack(pady=(10, 2))
        self.drop_sub = tk.Label(inner, text="pdf · docx · xlsx · pptx · csv · txt · md", font=self.f_small,
                                 bg=WHITE, fg=SLATE)
        self.drop_sub.pack()
        self.choose_button = self.button(inner, "파일 선택", self.choose_mask_file, parent_bg=WHITE)
        self.choose_button.pack(pady=(14, 0))
        if HAS_DND:
            self._enable_drop([self.drop, inner, self.drop_icon, self.drop_label, self.drop_sub])
        self._drop_widgets = [self.drop, inner, self.drop_icon, self.drop_label, self.drop_sub]

        self.file_label = tk.Label(box, textvariable=self.file_var, font=self.f_bold, bg=WHITE, fg=SLATE,
                                   anchor="w", justify="left")
        self.file_label.pack(fill="x")
        self.file_label.bind("<Configure>", lambda e: self.file_label.configure(wraplength=max(60, e.width - 4)))

        mode_card = Card(left)
        mode_card.pack(fill="x", pady=(16, 0))
        box = mode_card.inner
        self.card_title(box, "2", "방식 선택").pack(anchor="w")
        self.mode_seg = Segmented(box, [("md", "AI용 md", "추천"), ("xlsx", "엑셀 원본 형식", None)],
                                  self.mode_var, self.f_btn, self.f_tag, height=40, min_seg=130)
        self.mode_seg.pack(anchor="w", pady=(14, 8))
        tk.Label(box, text="AI에 올릴 때는 md를 추천해요.", font=self.f_small, bg=WHITE, fg=SLATE,
                 anchor="w").pack(fill="x")

        self.run_button = self.button(left, "가명처리 실행", self.run_mask, kind="primary", height=56,
                                      font=self.f_primary, parent_bg=MIST)
        self.run_button.pack(fill="x", pady=(16, 0))

        # ---- 오른쪽: 결과
        result_card = Card(right, expand=True)
        result_card.pack(fill="both", expand=True)
        box = result_card.inner
        box.grid_columnconfigure(0, weight=1)
        box.grid_rowconfigure(5, weight=1)
        self.card_title(box, "3", "결과 확인").grid(row=0, column=0, sticky="w")

        # 결과 상태: 줄마다 색을 달리할 수 있는 상자 + 그 아래 작은 회색 저장경로 한 줄
        self.status_wrap = tk.Frame(box, bg=WHITE)
        self.status_wrap.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        self.status_box = tk.Frame(self.status_wrap, bg=STATUS_COLORS["neutral"][1])
        self.status_box.pack(fill="x")
        self.status_inner = tk.Frame(self.status_box, bg=STATUS_COLORS["neutral"][1])
        self.status_inner.pack(fill="x", padx=16, pady=11)
        self.status_inner.bind("<Configure>", lambda _e: self._reflow_status())
        self._status_line_labels = []
        self.status_path = tk.Label(self.status_wrap, text="", font=self.f_small, bg=WHITE, fg=SLATE,
                                    anchor="w", justify="left")
        self.status_path.bind("<Configure>",
                              lambda e: self.status_path.configure(wraplength=max(60, e.width - 4)))

        # 확인 필요 목록: 내용이 있을 때만 보인다
        self.review_wrap = tk.Frame(box, bg=WHITE)
        self.review_wrap.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        tk.Label(self.review_wrap, text="확인 필요 목록 (사전에없음 / 동명이인확인)", font=self.f_small_bold,
                 bg=WHITE, fg="#B64400", anchor="w").pack(fill="x", pady=(0, 4))
        table, self.review_table = self.make_table(
            self.review_wrap, [("종류", 130, False), ("위치", 110, False), ("내용", 300, True)], 4)
        table.pack(fill="x")
        self.review_wrap.grid_remove()

        # 이름 의심 목록: 사전에 없지만 이름 같은 단어 (내용이 있을 때만 보인다)
        self.suspect_wrap = tk.Frame(box, bg=WHITE)
        self.suspect_wrap.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        self.suspect_count = tk.Label(self.suspect_wrap, text="이름 의심", font=self.f_small_bold,
                                      bg=WHITE, fg="#B64400", anchor="w")
        self.suspect_count.pack(fill="x", pady=(0, 4))
        self.suspect_list = tk.Frame(self.suspect_wrap, bg=WHITE)
        self.suspect_list.pack(fill="x")
        self.suspect_wrap.grid_remove()

        self.preview_title = tk.Label(box, text="미리보기 (예시)", font=self.f_small_bold, bg=WHITE, fg=SLATE,
                                      anchor="w")
        self.preview_title.grid(row=4, column=0, sticky="w", pady=(14, 6))
        preview_box, self.preview = self.make_text(box, readonly=True)
        preview_box.grid(row=5, column=0, sticky="nsew")
        self.set_text(self.preview, EXAMPLE_PREVIEW, highlight=True)

        buttons = tk.Frame(box, bg=WHITE)
        buttons.grid(row=6, column=0, sticky="w", pady=(16, 0))
        self.copy_result_button = self.button(buttons, "결과 내용 복사", self.copy_result)
        self.copy_result_button.pack(side="left")
        self.copy_result_button.state(["disabled"])   # 결과를 만들기 전에는 복사할 것이 없다
        self.button(buttons, "결과 폴더 열기", lambda: open_folder(workflows.AI_DIR)).pack(side="left", padx=10)

        self.set_status("아직 결과가 없어요 - 파일을 고르고 [가명처리 실행]을 눌러 주세요", "neutral")

    def _enable_drop(self, widgets):
        for w in widgets:
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            w.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            w.dnd_bind("<<Drop>>", self.safe(self._on_drop))

    def _draw_drop_border(self):
        """드롭 상자 테두리를 점선으로 그린다 (색은 self._drop_border_color)."""
        c = self.drop
        c.delete("border")
        w, h = c.winfo_width(), c.winfo_height()
        if w <= 1:
            w = int(c["width"]) or 320
        if h <= 1:
            h = int(c["height"]) or 180
        c.create_rectangle(2, 2, w - 2, h - 2, outline=self._drop_border_color, width=2,
                           dash=(5, 4), tags="border")

    def _paint_drop(self, color):
        for w in self._drop_widgets:
            w.configure(bg=color)
        self.choose_button.set_parent_bg(color)
        self._draw_drop_border()

    def _on_drop_enter(self, event):
        self._drop_border_color = ACTION          # 끌어다 놓는 중: 파란 테두리
        self._paint_drop(DROP_HOVER)
        return event.action

    def _on_drop_leave(self, event):
        self._drop_border_color = STEEL           # 평소: 회색 점선
        self._paint_drop(WHITE)
        return event.action

    def _on_drop(self, event):
        """끌어다 놓은 파일을 고른 것으로 처리한다 (처리 실행은 [가명처리 실행]을 눌러야 시작)."""
        self._drop_border_color = STEEL
        self._paint_drop(WHITE)
        files = self.root.tk.splitlist(event.data)
        if files:
            self.select_mask_file(files[0], extra=len(files) - 1)
        return event.action

    # ------------------------------------------------------------ 탭 2: 되돌리기
    def _build_restore_tab(self):
        tab = self.restore_tab
        tab.grid_columnconfigure(0, weight=1, uniform="restore")
        tab.grid_columnconfigure(1, weight=1, uniform="restore")
        tab.grid_rowconfigure(0, weight=1)

        # 왼쪽: AI 답변
        left = Card(tab, expand=True)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        box = left.inner
        box.grid_columnconfigure(0, weight=1)
        box.grid_rowconfigure(2, weight=1)
        self.card_title(box, "1", "AI 답변 붙여 넣기").grid(row=0, column=0, sticky="w")
        tk.Label(box, text="AI가 돌려준 글을 그대로 붙여 넣으세요.", font=self.f_small, bg=WHITE, fg=SLATE,
                 anchor="w").grid(row=1, column=0, sticky="w", pady=(4, 10))
        input_box, self.input_text = self.make_text(box)
        input_box.grid(row=2, column=0, sticky="nsew")
        self.attach_placeholder(self.input_text, INPUT_HINT)
        buttons = tk.Frame(box, bg=WHITE)
        buttons.grid(row=3, column=0, sticky="w", pady=(16, 0))
        self.restore_button = self.button(buttons, "복원", self.restore_from_box, kind="primary", height=44,
                                          padx=30, font=self.f_primary)
        self.restore_button.pack(side="left")
        self.restore_clip_button = self.button(buttons, "클립보드에서 바로 복원", self.restore_from_clipboard,
                                               height=44, padx=18)
        self.restore_clip_button.pack(side="left", padx=(10, 0))
        self.restore_file_button = self.button(buttons, "파일 복원", self.restore_from_file, height=44, padx=18)
        self.restore_file_button.pack(side="left", padx=(10, 0))

        # 오른쪽: 복원 결과
        right = Card(tab, expand=True)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        box = right.inner
        box.grid_columnconfigure(0, weight=1)
        box.grid_rowconfigure(2, weight=1)
        self.card_title(box, "2", "복원 결과").grid(row=0, column=0, sticky="w")
        tk.Label(box, text="실명이 들어 있으니 AI에 다시 올리지 마세요.", font=self.f_small, bg=WHITE, fg=SLATE,
                 anchor="w").grid(row=1, column=0, sticky="w", pady=(4, 10))
        output_box, self.output_text = self.make_text(box, readonly=True)
        output_box.grid(row=2, column=0, sticky="nsew")
        self.restored = ""
        self.set_text(self.output_text, "", placeholder=OUTPUT_HINT)
        self.copy_restored_button = self.button(box, "복사", self.copy_restored, height=44, padx=26)
        self.copy_restored_button.grid(row=3, column=0, sticky="w", pady=(16, 0))

        # 아래: 확인 필요 (복원 못 한 코드가 있을 때만 보인다)
        self.warn_card = Card(tab)
        self.warn_card.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        box = self.warn_card.inner
        tk.Label(box, text="확인 필요 - 복원하지 못한 코드", font=self.f_title, bg=WHITE, fg="#B64400",
                 anchor="w").pack(fill="x", pady=(0, 8))
        table, self.warn_table = self.make_table(box, [("코드", 160, False), ("이유", 600, True)], 5)
        table.pack(fill="x")
        self.warn_card.grid_remove()

    # ------------------------------------------------------------ 오류 처리·공통 동작
    def safe(self, fn):
        """버튼이 눌렸을 때 오류가 나도 프로그램이 꺼지지 않고 알림창으로 보여 준다."""
        def wrapper(*args, **kwargs):
            try:
                return fn(*args, **kwargs)
            except Exception as exc:   # noqa: BLE001 - 무엇이 터져도 알림창으로 보여 주는 것이 목적
                self._show_error(exc)
        return wrapper

    def _on_callback_error(self, exc_type, exc, tb):
        self._busy_end()
        self._show_error(exc)

    def _show_error(self, exc):
        """오류 알림. 복사·로그에는 가명처리한 보고만 넣는다."""
        from core.error_report import append_error_log, format_error_report, mask_for_report
        self._busy_end()
        report = format_error_report(exc)
        append_error_log(report)
        short = mask_for_report(f"{type(exc).__name__}: {exc}")
        self._error_dialog(short, report)

    def _error_dialog(self, short, report):
        win = tk.Toplevel(self.root)
        win.title("오류")
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text="오류", font=self.f_title, bg=WHITE, fg=INK).pack(anchor="w")
        tk.Label(box, text=short, font=self.f_small, bg=WHITE, fg=SLATE,
                 justify="left", wraplength=420, anchor="w").pack(anchor="w", pady=(8, 16))
        row = tk.Frame(box, bg=WHITE)
        row.pack(fill="x")
        PillButton(row, "오류 내용 복사", lambda: self._copy_error_report(report),
                   self.f_btn, kind="primary", height=40, padx=18, parent_bg=WHITE).pack(side="left")
        PillButton(row, "닫기", win.destroy, self.f_btn, kind="outline",
                   height=40, padx=18, parent_bg=WHITE).pack(side="left", padx=(10, 0))

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - w) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - h) // 3)
        win.geometry(f"+{x}+{y}")

    def _copy_error_report(self, report):
        """복사 전에 가명처리된 내용을 보여주고, 확인한 뒤에만 클립보드에 넣는다."""
        win = tk.Toplevel(self.root)
        win.title("오류 내용 복사")
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text="복사할 내용", font=self.f_title, bg=WHITE, fg=INK).pack(anchor="w")
        tk.Label(box, text="학생 정보가 없는지 확인한 뒤 제작자에게 보내 주세요.",
                 font=self.f_small, bg=WHITE, fg="#B64400",
                 wraplength=420, justify="left", anchor="w").pack(anchor="w", pady=(6, 10))
        preview, text = self.make_text(box, readonly=True)
        preview.pack(fill="both", expand=True)
        self.set_text(text, report)
        text.configure(width=56, height=12)

        def do_copy():
            self.root.clipboard_clear()
            self.root.clipboard_append(report)
            self.root.update()
            win.destroy()
            messagebox.showinfo(APP_TITLE, "오류 내용을 복사했어요.")

        row = tk.Frame(box, bg=WHITE)
        row.pack(fill="x", pady=(16, 0))
        PillButton(row, "이 내용 복사", do_copy, self.f_btn, kind="primary",
                   height=40, padx=18, parent_bg=WHITE).pack(side="left")
        PillButton(row, "닫기", win.destroy, self.f_btn, kind="outline",
                   height=40, padx=18, parent_bg=WHITE).pack(side="left", padx=(10, 0))

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - w) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - h) // 3)
        win.geometry(f"+{x}+{y}")
        win.grab_set()

    def _busy_start(self, message):
        self.root.configure(cursor="watch")
        self.set_status(message, "neutral")
        self.root.update()

    def _busy_end(self):
        self.root.configure(cursor="")

    def copy_to_clipboard(self, text, button):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        old = button.cget("text")
        button.configure(text="복사됨 ✓")
        self.root.after(1500, lambda: button.configure(text=old))

    @staticmethod
    def fill_table(tree, rows, wrap):
        """목록을 채운다. 내용이 없으면 목록 자리 자체를 숨겨서 화면을 깔끔하게 둔다."""
        tree.delete(*tree.get_children())
        for row in rows:
            tree.insert("", "end", values=row)
        tree.configure(height=max(1, min(len(rows), tree._max_rows)))   # 항목 수만큼만 보이게
        if rows:
            wrap.grid()
        else:
            wrap.grid_remove()

    # ------------------------------------------------------------ 맨 위 줄: 사전
    def refresh_dictionary(self):
        """사전을 읽어서 학생 수와 수정 날짜를 보여 주고, 없으면 버튼을 잠근다."""
        try:
            info = workflows.dictionary_info()
            problem = NO_DICT_MESSAGE if info is None else None
        except Exception as exc:   # noqa: BLE001 - 사전이 깨져 있어도 프로그램은 떠야 한다
            info, problem = None, f"학생사전을 읽을 수 없음 - {exc}"
        if info is None:
            self.dict_label.configure(text=problem, fg=STATUS_COLORS["bad"][0])
            self.dict_dot.configure(fg=STATUS_COLORS["bad"][0])
        else:
            students, staff, modified = info
            staff_text = "교직원 미등록" if staff == 0 else f"교직원 {staff}명"
            self.dict_label.configure(text=f"학생 {students}명 · {staff_text} · {modified}", fg=SLATE)
            self.dict_dot.configure(fg=STATUS_COLORS["ok"][0])
        for b in (self.run_button, self.restore_button, self.restore_clip_button, self.restore_file_button):
            b.state(["!disabled"] if info else ["disabled"])

    def build_dictionary(self):
        kind = self._ask_dict_kind()   # "학생" / "교직원" / None(취소)
        if not kind:
            return
        noun = "학생" if kind == "학생" else "교직원"
        title = "학생 명렬표 엑셀 선택" if kind == "학생" else "교직원 명단 엑셀 선택"
        path = filedialog.askopenfilename(title=title,
                                          filetypes=[("엑셀 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")])
        if not path:
            return
        self.root.configure(cursor="watch")
        self.root.update()
        try:
            result = workflows.build_dictionary(path, kind=kind)
        finally:
            self._busy_end()
        self.refresh_dictionary()
        self._mark_result_stale(f"{noun} 명단으로 사전을 업데이트했습니다. [가명처리 실행]을 다시 눌러 주세요.")

        lines = [f"{noun} {result.count}명으로 사전을 업데이트했습니다."]
        if kind == "학생":
            lines += [f"찾은 열: {', '.join(result.found)}",
                      f"못 찾은 열: {', '.join(result.missing) if result.missing else '없음'}"]
        if result.warnings:
            lines += ["", f"경고 {len(result.warnings)}건:"]
            lines += [f"- {w}" for w in result.warnings[:25]]
            if len(result.warnings) > 25:
                lines.append(f"... 외 {len(result.warnings) - 25}건")
            messagebox.showwarning(APP_TITLE, "\n".join(lines))
        else:
            messagebox.showinfo(APP_TITLE, "\n".join(lines))

    def _ask_dict_kind(self):
        """[사전 업데이트]: '학생 명렬표 / 교직원 명단' 중 무엇을 넣을지 고른다. (고른 값 또는 None)"""
        win = tk.Toplevel(self.root)
        win.title(APP_TITLE)
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)
        choice = {"value": None}

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text="무엇을 넣을까요?", font=self.f_title, bg=WHITE, fg=INK).pack(anchor="w")
        tk.Label(box, text="학생 명렬표(다른 반 포함)인지, 교직원 명단인지 골라 주세요.",
                 font=self.f_small, bg=WHITE, fg=SLATE).pack(anchor="w", pady=(6, 16))
        row = tk.Frame(box, bg=WHITE)
        row.pack(fill="x")

        def pick(value):
            choice["value"] = value
            win.destroy()

        PillButton(row, "학생 명렬표", lambda: pick("학생"), self.f_btn, kind="primary",
                   height=48, padx=26, parent_bg=WHITE).pack(side="left")
        PillButton(row, "교직원 명단", lambda: pick("교직원"), self.f_btn, kind="outline",
                   height=48, padx=26, parent_bg=WHITE).pack(side="left", padx=(12, 0))

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - w) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - h) // 3)
        win.geometry(f"+{x}+{y}")
        win.grab_set()
        win.wait_window()
        return choice["value"]

    # ------------------------------------------------------------ 탭 1: 가리기
    def choose_mask_file(self):
        path = filedialog.askopenfilename(title="가명처리할 파일 선택", filetypes=MASK_FILETYPES)
        if path:
            self.select_mask_file(path)

    def select_mask_file(self, path, extra=0):
        path = Path(path)
        ext = path.suffix.lower()
        if ext in workflows.HWP_EXTS:
            messagebox.showinfo(APP_TITLE, workflows.HWP_MESSAGE)
            return
        if ext not in workflows.MASK_EXTS:
            messagebox.showwarning(APP_TITLE, f"지원하지 않는 형식입니다: {ext or '(확장자 없음)'}\n"
                                              "지원: pdf, docx, xlsx, xls, pptx, csv, txt, md")
            return
        self.selected_path = path
        self.file_var.set(f"선택된 파일: {path.name}" + (f" 외 {extra}개" if extra > 0 else ""))
        self.file_label.configure(fg=INK)
        self._confirmed_names = []
        self._ignored_names = set()
        # '엑셀 원본 형식 유지'는 .xlsx 일 때만 고를 수 있다
        if ext == ".xlsx":
            self.mode_seg.set_enabled("xlsx", True)
        else:
            self.mode_var.set("md")
            self.mode_seg.set_enabled("xlsx", False)
        self._mark_result_stale("새 파일을 골랐습니다. [가명처리 실행]을 눌러 주세요.")

    def _mark_result_stale(self, message):
        """이전 결과가 새 파일/새 사전과 맞지 않으니, 복사 버튼을 잠그고 안내만 보여 준다."""
        if self.result is None and self.copy_result_button.instate(["disabled"]):
            return   # 아직 결과를 만든 적이 없으면 처음 화면(예시)을 그대로 둔다
        self.result = None
        self.copy_result_button.state(["disabled"])
        self.suspect_wrap.grid_remove()
        self.set_status(message, "neutral")

    def run_mask(self):
        if self.selected_path is None:
            messagebox.showinfo(APP_TITLE, "먼저 가명처리할 파일을 선택하세요.")
            return
        dictionary = workflows.get_dictionary()
        self._confirmed_names = []
        self._ignored_names = set()
        self.copy_result_button.state(["disabled"])
        self._busy_start("처리 중입니다...")
        try:
            result = workflows.mask_file(self.selected_path, self.mode_var.get(), dictionary)
        except Exception:
            self.result = None
            self.set_status("⚠ 처리하지 못했습니다 - 오류 내용은 알림창을 확인하세요", "bad")
            raise
        finally:
            self._busy_end()
        self.show_mask_result(result)

    def show_mask_result(self, result):
        self.result = result
        rows = [r for r in result.review]
        rows += [(f"남은 개인정보({kind})", where, value) for where, kind, value in result.findings]
        self.fill_table(self.review_table, rows, self.review_wrap)
        self.preview_title.configure(text="미리보기")
        # 미리보기에서만 md 표 구분선(| --- | --- |)을 숨긴다. 복사·저장되는 result.text 는 그대로 둔다.
        display = self._hide_md_separators(result.text) if result.ok else ""
        self.set_text(self.preview, display, highlight=True)
        # 남은 개인정보가 있으면 복사를 막는다. 이름 의심은 복사를 막지 않는다(복사 때 안내만).
        self.copy_result_button.state(["!disabled"] if result.can_copy else ["disabled"])
        self._render_suspects(self._pending_suspects())
        self._set_mask_status()

    # ---- 이름 의심 ----
    def _pending_suspects(self):
        """아직 '가리기/무시'로 처리하지 않은 이름 의심 목록."""
        if not self.result:
            return []
        return [s for s in self.result.suspects if s.word not in self._ignored_names]

    def _set_mask_status(self):
        r = self.result
        if r is None:
            return
        head_text, path = self._split_headline(r.headline)   # 저장 경로는 상태 줄에서 뺀다
        n = len(self._pending_suspects())
        if not (r.ok and r.clean):
            lines = [(head_text, "bad")]                      # 위험(빨강) 한 줄
        else:
            lines = [(head_text, "ok")]                       # "남은 개인정보 없음" 초록 줄
            if n:
                extra = f" · 확인 필요 {len(r.review)}건" if r.review else ""
                lines.append((f"⚠ 이름 의심 {n}건 - 아래에서 가릴지 정해 주세요{extra}", "warn"))
            elif r.review:
                lines.append((f"⚠ 확인 필요 {len(r.review)}건 - 아래 목록에서 위치를 확인하세요", "warn"))
        box_kind = max((k for _, k in lines), key=lambda k: self._SEVERITY[k])   # 상자 바탕=가장 심각
        self._render_status(lines, box_kind, path)

    @staticmethod
    def _hide_md_separators(text):
        """미리보기용: md 표의 구분선 줄(| --- | --- |)만 뺀 글을 돌려준다. (원본은 바꾸지 않음)"""
        kept = []
        for line in text.split("\n"):
            s = line.strip()
            if s and "|" in s and set(s) <= set("|-: "):
                continue   # 칸이 대시·콜론·막대뿐인 줄 = 표 구분선 → 숨김
            kept.append(line)
        return "\n".join(kept)

    def _render_suspects(self, suspects):
        for w in self.suspect_list.winfo_children():
            w.destroy()
        if not suspects:
            self.suspect_wrap.grid_remove()
            return
        self.suspect_wrap.grid()
        self.suspect_count.configure(text=f"이름 의심 {len(suspects)}건 (사전에 없지만 이름 같은 단어)")
        cap = 12
        for s in suspects[:cap]:
            row = tk.Frame(self.suspect_list, bg=WHITE)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=s.word, font=self.f_small_bold, bg=WHITE, fg=INK,
                     width=7, anchor="w").pack(side="left")
            self.button(row, "항상 무시", lambda w=s.word: self._always_ignore_suspect(w),
                        height=30, padx=12).pack(side="right")
            self.button(row, "무시", lambda w=s.word: self._ignore_suspect(w),
                        height=30, padx=12).pack(side="right", padx=(0, 6))
            self.button(row, "이 단어도 가리기", lambda w=s.word: self._confirm_suspect(w),
                        kind="primary", height=30, padx=12).pack(side="right", padx=(0, 6))
            tk.Label(row, text=s.context, font=self.f_small, bg=WHITE, fg=SLATE,
                     anchor="w", justify="left").pack(side="left", fill="x", expand=True, padx=(6, 8))
        if len(suspects) > cap:
            tk.Label(self.suspect_list, text=f"… 외 {len(suspects) - cap}건은 미리보기에서 확인하세요",
                     font=self.f_small, bg=WHITE, fg=SLATE, anchor="w").pack(fill="x", pady=(2, 0))

    def _confirm_suspect(self, word):
        """[이 단어도 가리기]: 그 단어를 [이름N] 으로 바꾸고 결과를 다시 만든다."""
        if self.selected_path is None:
            return
        if word not in self._confirmed_names:
            self._confirmed_names.append(word)
        self._ignored_names.discard(word)
        self._busy_start("다시 처리 중입니다...")
        try:
            result = workflows.mask_file(self.selected_path, self.mode_var.get(),
                                         workflows.get_dictionary(), extra_names=self._confirmed_names)
        finally:
            self._busy_end()
        self.show_mask_result(result)

    def _ignore_suspect(self, word):
        """[무시]: 이 단어를 의심 목록에서만 뺀다 (글은 그대로 둔다)."""
        self._ignored_names.add(word)
        self._render_suspects(self._pending_suspects())
        self._set_mask_status()

    def _always_ignore_suspect(self, word):
        """[항상 무시]: 확인 후 data/무시단어.txt 에 저장하고, 이번 목록에서도 뺀다."""
        if not messagebox.askyesno(
                APP_TITLE,
                f"'{word}'를 앞으로 계속 무시할까요? 사람 이름이면 누르지 마세요."):
            return
        refused = workflows.add_ignored_word(word)
        if refused:
            messagebox.showwarning(APP_TITLE, refused)
            return
        self._ignored_names.add(word)
        self._render_suspects(self._pending_suspects())
        self._set_mask_status()

    def show_ignored_words(self):
        """상단 [무시 단어]: 저장된 목록을 보고 골라 지울 수 있는 작은 창."""
        win = tk.Toplevel(self.root)
        win.title("무시 단어")
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text="무시 단어", font=self.f_title, bg=WHITE, fg=INK).pack(anchor="w")
        tk.Label(box, text="이름 의심에서 빼 둔 단어입니다. 고른 뒤 지울 수 있어요.",
                 font=self.f_small, bg=WHITE, fg=SLATE).pack(anchor="w", pady=(6, 12))

        frame = tk.Frame(box, bg=WHITE)
        frame.pack(fill="both", expand=True)
        listbox = tk.Listbox(frame, font=self.f_body, height=10, width=28, selectmode="extended",
                             bg=MIST, fg=INK, relief="flat", highlightthickness=0, bd=0,
                             selectbackground=DROP_HOVER, selectforeground=INK)
        bar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview,
                            style="Slim.Vertical.TScrollbar")
        listbox.configure(yscrollcommand=self.autohide(bar, listbox))
        listbox.pack(side="left", fill="both", expand=True)
        empty = tk.Label(box, text="저장된 무시 단어가 없어요.", font=self.f_small,
                         bg=WHITE, fg=SLATE, anchor="w")

        def refresh():
            words = workflows.load_ignored_words()
            listbox.delete(0, "end")
            for w in words:
                listbox.insert("end", w)
            if words:
                empty.pack_forget()
            else:
                empty.pack(fill="x", pady=(8, 0))

        def delete_selected():
            picked = [listbox.get(i) for i in listbox.curselection()]
            if not picked:
                messagebox.showinfo(APP_TITLE, "지울 단어를 먼저 골라 주세요.")
                return
            if not messagebox.askyesno(APP_TITLE, f"고른 {len(picked)}개를 무시 목록에서 지울까요?"):
                return
            workflows.remove_ignored_words(picked)
            refresh()

        buttons = tk.Frame(box, bg=WHITE)
        buttons.pack(fill="x", pady=(16, 0))
        self.button(buttons, "선택한 단어 지우기", delete_selected, height=36, padx=16).pack(side="left")
        self.button(buttons, "닫기", win.destroy, height=36, padx=16).pack(side="left", padx=(8, 0))
        refresh()

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - w) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - h) // 3)
        win.geometry(f"+{x}+{y}")

    def copy_result(self):
        if self.result is None or not self.result.can_copy:
            messagebox.showwarning(APP_TITLE, "복사할 수 있는 결과가 없습니다.")
            return
        pending = self._pending_suspects()
        if pending:
            messagebox.showwarning(
                APP_TITLE,
                f"확인하지 않은 이름 의심 {len(pending)}건이 있어요.\n"
                "그래도 복사합니다. 필요하면 [이 단어도 가리기]로 가린 뒤 다시 복사하세요.")
        self.copy_to_clipboard(self.result.text, self.copy_result_button)

    # ------------------------------------------------------------ 탭 2: 되돌리기
    def _name_map(self):
        """현재 결과에서 '이 단어도 가리기'로 바꾼 {코드: 단어} (없으면 None). [이름N] 복원에 쓴다."""
        return self.result.name_map if self.result and self.result.name_map else None

    def _show_restored(self, restored, warnings):
        self.restored = restored
        self.set_text(self.output_text, restored, placeholder=OUTPUT_HINT)
        rows = [split_warning(w) for w in warnings]
        self.fill_table(self.warn_table, rows, self.warn_card)

    def restore_from_box(self):
        text = "" if self.input_text._ph_on else self.input_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showinfo(APP_TITLE, "AI 답변을 먼저 붙여 넣으세요.")
            return
        restored, _, warnings = unmask_text(text, workflows.get_dictionary(), self._name_map())
        self._show_restored(restored, warnings)

    def restore_from_clipboard(self):
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            text = ""
        if not text.strip():
            messagebox.showinfo(APP_TITLE, "클립보드에 글이 없습니다.")
            return
        restored, _, warnings = unmask_text(text, workflows.get_dictionary(), self._name_map())
        self._show_restored(restored, warnings)
        self.root.clipboard_clear()
        self.root.clipboard_append(restored)
        self.root.update()

    def restore_from_file(self):
        path = filedialog.askopenfilename(
            title="복원할 파일 선택",
            filetypes=[("텍스트·엑셀", "*.txt *.md *.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self.root.configure(cursor="watch")
        self.root.update()
        try:
            result = workflows.restore_file(path, workflows.get_dictionary(), extra=self._name_map())
        finally:
            self._busy_end()
        saved = result.out_path.relative_to(workflows.BASE_DIR)
        shown = result.text if result.text is not None else f"엑셀 파일을 복원했습니다.\n저장 위치: {saved}"
        self._show_restored(shown, result.warnings)
        messagebox.showinfo(APP_TITLE, f"복원했습니다 ({result.changes}건, 경고 {len(result.warnings)}건)\n"
                                       f"저장 위치: {saved}")

    def copy_restored(self):
        text = self.restored
        if not text.strip():
            messagebox.showinfo(APP_TITLE, "복사할 복원 결과가 없습니다.")
            return
        self.copy_to_clipboard(text, self.copy_restored_button)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
