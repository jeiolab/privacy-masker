#!/usr/bin/env python3
"""개인정보 가리개 - 화면.   실행: python3 app.py

화면은 이 파일, 실제 작업은 workflows.py 가 core/·adapters/ 의 함수를 불러서 한다.
- 가리기: 왼쪽에서 파일·방식을 고르고, 오른쪽에서 결과를 확인한다.
- 되돌리기: 왼쪽에 AI 답변을 붙여 넣고, 오른쪽에서 복원 결과를 본다.
- 어떤 오류가 나도 프로그램이 꺼지지 않고, 알림창으로 오류 내용을 보여 준다.
"""
import base64
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
from core.config import program_dir
from core.masker import MASK_OPTIONS
from core.restorer import unmask_text
from core.updater import fetch_update_info, is_newer, open_download_page, update_how_to
from core.version import VERSION

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:   # tkinterdnd2 가 없으면 끌어다 놓기만 빠지고 나머지는 그대로 동작
    HAS_DND = False

APP_TITLE = "개인정보 가리개"
WINDOW_SIZE = "1180x1140"
MIN_SIZE = (1040, 1120)
COPYRIGHT_NOTICE = "© 2026 능주고등학교 교사 JEIO"
LOGO_BASELINE = 13                       # jeio_logo.png 안에서 E·I·O 글자 밑줄 위치

# 맥·윈도우·리눅스에서 한글이 잘 보이는 글꼴 후보 (앞에서부터 설치된 것을 쓴다)
FONT_CANDIDATES = [
    "Noto Sans KR",                              # 윈도우에 있으면 화면 글꼴로 쓴다
    "Apple SD Gothic Neo", "AppleGothic",        # 맥
    "맑은 고딕", "Malgun Gothic",
    "Noto Sans CJK KR", "NanumGothic",
]

# ---- 색: 흰 카드 + 옅은 회색 바탕, 강조색은 짙은 초록 하나 (지키는 도구라는 느낌, 파랑은 쓰지 않음)
MIST = "#F2F3F5"        # 화면 바탕
WHITE = "#FFFFFF"       # 카드
INK = "#16181D"         # 본문 글자
SLATE = "#5F6672"       # 보조 글자 (예전보다 진하게 — 흐려서 안 읽히던 것)
STEEL = "#A3A9B3"       # 비활성
HAIR = "#E3E5E9"        # 가는 선
CONTROL = "#EEF0F3"     # 세그먼트·보조 버튼 바탕
ACCENT = "#1D5C4B"      # 강조: 주 버튼, 단계 번호, 고른 탭
ACCENT_HOVER = "#174B3D"
ACCENT_PRESS = "#123B30"
ACCENT_SOFT = "#E7F1EE"  # 강조색 옅은 바탕 (끌어다 놓는 중, 단계 완료)
ACTION = ACCENT         # 주 버튼 (예전 이름을 쓰는 곳이 있어 남겨 둠)
ACTION_HOVER = ACCENT_HOVER
ACTION_PRESS = ACCENT_PRESS
DROP_BG = "#F7F8FA"     # 파일 놓는 칸
DROP_HOVER = ACCENT_SOFT
SELECT_BG = "#DCE9E4"
MARK = "#9A5B13"        # 확인 필요 같은 작은 표시 (주황 갈색)

# 상태 줄 색: (글자, 바탕)
STATUS_COLORS = {
    "ok": ("#1F7A4D", WHITE),
    "warn": (MARK, WHITE),
    "bad": ("#B3261E", WHITE),
    "neutral": (SLATE, WHITE),
}

# 미리보기 아래 「색 읽는 법」: (보이는 이름, CODE_STYLES 태그)
LEGEND = [("학생·보호자·교사 이름", "code_name"), ("학번", "code_sid"), ("연락처·주소·생일", "code_bracket"),
          ("명단 밖 이름", "code_ename"), ("가린 번호", "code_unknown")]

# 미리보기에서 색을 다르게 보여 줄 코드들 (태그 이름, 정규식, 글자색, 배경색)
CODE_STYLES = [
    ("code_unknown", r"\[(?:전화번호|주민번호|이메일)\]", "#8E2F2F", "#F8EBEB"),
    ("code_ename", r"\[이름\d+\]", "#1F4B45", "#E5F1EE"),
    ("code_bracket", r"\[(?:학생전화|보호자전화|주소|생년월일)\d{3,}(?:또는\d{3,})*\]", "#5C3D78", "#F3EDF6"),
    ("code_sid", r"S\d{3,}(?:또는\d{3,})*", "#8A5A2B", "#F8F0E4"),
    ("code_name", r"(?:학생|보호자|교사)\d{3,}(?:또는\d{3,})*", "#1E4D8C", "#E8F0F8"),
]
CODE_REGEX = re.compile("|".join(f"(?P<{tag}>{pat})" for tag, pat, _, _ in CODE_STYLES))

KIND_LABEL = {
    "사전에없음": "명단에 없음",
    "동명이인확인": "동명이인",
    "동명이인": "동명이인",
    "중복확인": "여러 후보",
    "오류": "문제",
    "성명": "이름",
    "보호자명": "보호자 이름",
    "학번": "학번",
    "전화번호": "전화번호",
    "이메일": "이메일",
    "주민번호": "주민번호",
}


def say_kind(kind):
    """표에 보이는 종류 이름. 저장용 코드는 그대로 두고 화면만 바꾼다."""
    if not kind:
        return kind
    if kind in KIND_LABEL:
        return KIND_LABEL[kind]
    if "/" in kind:
        return " · ".join(KIND_LABEL.get(part, part) for part in kind.split("/"))
    return kind


NO_DICT_MESSAGE = "명단이 없어요. 위쪽 [명단 넣기]를 눌러 주세요"
INPUT_HINT = "여기를 눌러 AI 답변을 붙여 넣으세요"
OUTPUT_HINT = "[되돌리기]를 누르면 실명으로 돌아온 글이 여기에 나와요."
EXAMPLE_PREVIEW = ("원래 글\n홍길동(10203) 학생은 010-1234-5678로 연락해요. 어머니 김영희 님께 안내했어요.\n\n"
                   "가린 글\n학생003(S003) 학생은 [학생전화003]로 연락해요. 어머니 보호자003 님께 안내했어요.")
MASK_FILETYPES = [("지원하는 파일", "*.pdf *.docx *.xlsx *.xls *.pptx *.csv *.txt *.md *.hwp *.hwpx"),
                  ("모든 파일", "*.*")]
HAND_CURSOR = "pointinghand" if sys.platform == "darwin" else "hand2"


def open_path(path, make_dir=False):
    """파일이나 폴더를 기본 프로그램으로 연다 (맥·윈도우·리눅스)."""
    path = Path(path)
    if make_dir:
        path.mkdir(parents=True, exist_ok=True)
    if sys.platform.startswith("win"):
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=True)
    else:
        subprocess.run(["xdg-open", str(path)], check=True)


def open_folder(path):
    """폴더를 연다 (맥·윈도우·리눅스)."""
    open_path(path, make_dir=True)


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


def pick_title_family(root, body):
    """제목만 한 단계 두꺼운 글꼴이 있으면 그걸 쓴다."""
    if body == "Noto Sans KR" and "Noto Sans KR Medium" in set(tkfont.families(root)):
        return "Noto Sans KR Medium"
    return body


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


class StepBadge(tk.Canvas):
    """카드 제목 앞 동그란 단계 번호. done 이면 체크 표시."""

    def __init__(self, parent, number, font, size=26, parent_bg=WHITE):
        super().__init__(parent, width=size, height=size, bg=parent_bg, highlightthickness=0, bd=0)
        self.number, self.font, self.size, self.done, self.off = number, font, size, False, False
        self._redraw()

    def set(self, done=None, off=None):
        if done is not None:
            self.done = done
        if off is not None:
            self.off = off
        self._redraw()

    def _redraw(self):
        self.delete("all")
        n = self.size
        fill = HAIR if self.off else ACCENT
        self.create_oval(1, 1, n - 1, n - 1, fill=fill, outline="")
        fg = SLATE if self.off else WHITE
        if self.done and not self.off:
            self.create_line(n * .30, n * .52, n * .45, n * .66, n * .71, n * .37, fill=WHITE, width=2.2,
                             capstyle="round", joinstyle="round")
        else:
            self.create_text(n / 2, n / 2, text=str(self.number), font=self.font, fill=fg)


def draw_logo(canvas, x, y, size):
    """로고: 둥근 사각형 안에 가려진 글줄 세 개 (가운데 줄만 짧게 — 「가린 글」)."""
    round_rect(canvas, x, y, x + size, y + size, size * .28, fill=ACCENT)
    pad, h = size * .22, size * .1
    for i, ratio in enumerate((1.0, .55, .8)):
        top = y + size * .27 + i * size * .19
        canvas.create_line(x + pad + h / 2, top + h / 2, x + pad + (size - 2 * pad) * ratio - h / 2, top + h / 2,
                           width=h, capstyle="round", fill=WHITE if i != 1 else "#9CC7B8")


def draw_doc_icon(canvas, cx, cy, color, accent):
    """파일 놓는 칸 가운데의 얇은 문서 그림 + 위쪽 화살표."""
    w, h, fold = 34, 42, 10
    x0, y0 = cx - w / 2, cy - h / 2
    canvas.create_line(x0, y0, x0 + w - fold, y0, x0 + w, y0 + fold, x0 + w, y0 + h, x0, y0 + h, x0, y0,
                       fill=color, width=1.6, joinstyle="round")
    canvas.create_line(x0 + w - fold, y0, x0 + w - fold, y0 + fold, x0 + w, y0 + fold, fill=color, width=1.6)
    canvas.create_line(cx, y0 + h - 10, cx, y0 + 15, fill=accent, width=2, capstyle="round")
    canvas.create_line(cx - 6, y0 + 21, cx, y0 + 15, cx + 6, y0 + 21, fill=accent, width=2, capstyle="round",
                       joinstyle="round")


def dashed_round_rect(canvas, x1, y1, x2, y2, r, color, width=1.5, dash=(5, 4)):
    """점선 둥근 사각형 (변은 선, 모서리는 호)."""
    kw = dict(fill=color, width=width, dash=dash)
    canvas.create_line(x1 + r, y1, x2 - r, y1, **kw)
    canvas.create_line(x1 + r, y2, x2 - r, y2, **kw)
    canvas.create_line(x1, y1 + r, x1, y2 - r, **kw)
    canvas.create_line(x2, y1 + r, x2, y2 - r, **kw)
    akw = dict(style="arc", outline=color, width=width, dash=dash)
    canvas.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, **akw)
    canvas.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, **akw)
    canvas.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, **akw)
    canvas.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, **akw)


class Tick(tk.Canvas):
    """18px 체크. 켜지면 초록 네모와 흰 체크, 꺼지면 흰 네모."""

    def __init__(self, parent, variable, command=None, parent_bg=WHITE):
        super().__init__(parent, width=18, height=18, bg=parent_bg, highlightthickness=0, bd=0)
        self.var = variable
        self.command = command
        try:
            self.configure(cursor=HAND_CURSOR)
        except tk.TclError:
            pass
        self.bind("<Button-1>", self._toggle)
        self._draw()

    def set(self, value):
        self.var.set(bool(value))
        self._draw()
        if self.command:
            self.command()

    def _toggle(self, _event):
        self.set(not self.var.get())

    def _draw(self):
        self.delete("all")
        if self.var.get():
            round_rect(self, 0, 0, 18, 18, 4, fill=ACCENT)
            self.create_line(4.2, 9.2, 7.4, 12.6, 13.6, 5.4, fill=WHITE, width=1.8,
                             capstyle="round", joinstyle="round")
        else:
            round_rect(self, 0.5, 0.5, 17.5, 17.5, 4, fill=WHITE, outline="#CDD1D8")


class Card(tk.Frame):
    """둥근 모서리 카드. 내용은 card.inner 안에 넣는다.

    뒤에 둥근 배경을 그리는 캔버스를 깔고, 그 위에 안쪽 틀을 여백만큼 띄워 얹는다.
    (안쪽 내용 크기에 따라 카드 크기가 저절로 맞춰지고, 남는 자리는 안쪽 내용이 채운다.)
    """

    def __init__(self, parent, parent_bg=MIST, fill=WHITE, radius=22, pad=26, expand=False, border=HAIR):
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
        "outline": dict(fill=WHITE, hover=CONTROL, press=HAIR, fg=INK, outline="#CDD1D8"),
        "quiet": dict(fill=None, hover=CONTROL, press=HAIR, fg=SLATE, outline=None),
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
            if not fill:
                fill = self["bg"]
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
                round_rect(self, x0, pad, x0 + sw, h - pad, (h - 2 * pad) / 2, fill=WHITE, outline=HAIR)
            fg = STEEL if off else INK
            font = self.font
            lw = font.measure(label)
            tw = self.tag_font.measure(tag) + 8 if tag else 0
            start = x0 + sw / 2 - (lw + tw) / 2
            self.create_text(start, h / 2, text=label, font=font, fill=fg, anchor="w")
            if tag:
                self.create_text(start + lw + 8, h / 2, text=tag, font=self.tag_font,
                                 fill=STEEL if off else MARK, anchor="w")

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


class TextTabs(tk.Frame):
    """글자 탭. 고른 항목 아래에만 짧은 선을 둔다."""

    def __init__(self, parent, options, variable, font, command=None):
        super().__init__(parent, bg=WHITE)
        self.var, self.command = variable, command
        self._cells = []
        for value, label in options:
            cell = tk.Frame(self, bg=WHITE, cursor=HAND_CURSOR)
            cell.pack(side="left", padx=(0, 28))
            text = tk.Label(cell, text=label, font=font, bg=WHITE, fg=SLATE, cursor=HAND_CURSOR)
            text.pack(anchor="w")
            line = tk.Frame(cell, bg=WHITE, height=2)
            line.pack(fill="x", pady=(8, 0))
            for w in (cell, text, line):
                w.bind("<Button-1>", lambda _e, v=value: self._pick(v))
            self._cells.append((value, text, line))
        variable.trace_add("write", lambda *_: self._paint())
        self._paint()

    def _pick(self, value):
        if self.var.get() != value:
            self.var.set(value)
            if self.command:
                self.command(value)
        else:
            self._paint()

    def _paint(self):
        current = self.var.get()
        for value, text, line in self._cells:
            on = value == current
            text.configure(fg=INK if on else SLATE)
            line.configure(bg=INK if on else WHITE)


# ====================================================================== 앱
class App:
    def __init__(self):
        self.root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*MIN_SIZE)
        self.root.configure(bg=MIST)

        self.family = pick_font_family(self.root)
        self.title_family = pick_title_family(self.root, self.family)
        self._setup_fonts_and_styles()

        self.file_var = tk.StringVar(value="아직 고른 파일이 없어요")
        self.mode_var = tk.StringVar(value="md")
        self.tab_var = tk.StringVar(value="mask")
        self._dict_ok = False
        self.selected_path = None
        self.selected_paths = []
        self._kind_vars = {}
        self._kind_status = "hidden"
        self.result = None
        self._confirmed_names = []      # '이 단어도 가리기'로 고른 이름 단어 (순서 유지)
        self._ignored_names = set()     # '무시'한 이름 의심 단어

        self._build_top_bar()
        self._build_copyright()
        self.content = tk.Frame(self.root, bg=MIST)
        self.content.pack(fill="both", expand=True, padx=28, pady=(20, 24))
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
        fam, title = self.family, self.title_family
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            tkfont.nametofont(name).configure(family=fam, size=12)
        self.f_body = tkfont.Font(family=fam, size=12)
        self.f_bold = tkfont.Font(family=title, size=12)
        self.f_small = tkfont.Font(family=fam, size=10)
        self.f_small_bold = tkfont.Font(family=title, size=10)
        self.f_tag = tkfont.Font(family=fam, size=10)
        self.f_title = tkfont.Font(family=title, size=16)
        self.f_app = tkfont.Font(family=title, size=20)
        self.f_tab = tkfont.Font(family=title, size=13)
        self.f_btn = tkfont.Font(family=fam, size=12)
        self.f_primary = tkfont.Font(family=title, size=13)
        self.f_text = tkfont.Font(family=fam, size=13)
        self.f_text_bold = tkfont.Font(family=title, size=13)

        style = ttk.Style(self.root)
        style.theme_use("clam")   # 맥·윈도우에서 같은 모양으로 보이게
        style.configure("Treeview", font=self.f_body, rowheight=30, background=WHITE, fieldbackground=WHITE,
                        foreground=INK, borderwidth=0, relief="flat", bordercolor=WHITE, lightcolor=WHITE,
                        darkcolor=WHITE)
        style.configure("Treeview.Heading", font=self.f_small, background=DROP_BG, foreground=SLATE,
                        relief="flat", borderwidth=0, padding=(10, 8))
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

    def card_title(self, parent, text):
        row = tk.Frame(parent, bg=WHITE)
        tk.Label(row, text=text, font=self.f_title, bg=WHITE, fg=INK).pack(side="left")
        return row

    def make_text(self, parent, readonly=False):
        """연한 회색 둥근 글상자 + 얇은 스크롤바. (box, text) 를 돌려준다."""
        box = Card(parent, parent_bg=WHITE, fill=DROP_BG, radius=12, pad=2, expand=True, border=None)
        text = tk.Text(box.inner, wrap="word", font=self.f_text, bg=DROP_BG, fg=INK, relief="flat", bd=0,
                       highlightthickness=0, padx=16, pady=14, spacing1=4, spacing3=8, insertbackground=INK,
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
            text_widget.tag_configure(tag, foreground=fg, font=self.f_text_bold)
            text_widget.tag_remove(tag, "1.0", "end")
        for line_no, line in enumerate(text_widget.get("1.0", "end-1c").split("\n"), start=1):
            for m in CODE_REGEX.finditer(line):
                text_widget.tag_add(m.lastgroup, f"{line_no}.{m.start()}", f"{line_no}.{m.end()}")

    _SEVERITY = {"neutral": 0, "ok": 1, "warn": 2, "bad": 3}   # 빨강(bad) > 주황(warn) > 초록(ok)

    def set_status(self, text, kind="neutral"):
        """결과 상태 한 덩어리(모두 같은 색). 여러 줄이면 전부 같은 kind 색으로 보여 준다."""
        self._render_status([(line, kind) for line in text.split("\n")], kind)

    def _render_status(self, lines, box_kind, path=None):
        """점 하나와 글자로 상태를 보여 준다. path 가 있으면 그 아래에 저장 위치를 둔다."""
        self.status_dot.delete("all")
        self.status_dot.create_oval(1, 5, 9, 13, fill=STATUS_COLORS[box_kind][0], outline="")
        for w in self.status_inner.winfo_children():
            w.destroy()
        self._status_line_labels = []
        for i, (text, kind) in enumerate(lines):
            lbl = tk.Label(self.status_inner, text=text, font=self.f_bold if i == 0 else self.f_small,
                           bg=WHITE, fg=STATUS_COLORS[kind][0], anchor="w", justify="left")
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
            self.toast(f"오래된 결과 파일 {removed}개를 지웠어요")

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
                "가린 결과와 되돌린 파일을 모두 지웁니다.\n"
                "명단은 지우지 않아요.\n\n계속할까요?",
                icon="warning"):
            return
        removed = workflows.clear_all_outputs()
        self.toast(f"결과 파일 {removed}개를 지웠어요" if removed else "지울 결과 파일이 없어요")

    # ------------------------------------------------------------ 맨 위 줄
    def _build_top_bar(self):
        bar = tk.Frame(self.root, bg=WHITE)
        bar.pack(fill="x")
        top = tk.Frame(bar, bg=WHITE)
        top.pack(fill="x", padx=32, pady=(18, 16))
        top.grid_columnconfigure(1, weight=1)

        # 왼쪽: 로고 + 이름 + 한 줄 설명
        logo = tk.Canvas(top, width=44, height=44, bg=WHITE, highlightthickness=0, bd=0)
        draw_logo(logo, 1, 1, 42)
        logo.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 14))
        tk.Label(top, text=APP_TITLE, font=self.f_app, bg=WHITE, fg=INK).grid(row=0, column=1, sticky="sw")
        tk.Label(top, text="학생 이름·학번·연락처를 가린 뒤 AI에 올리고, 답은 다시 실명으로 되돌려요",
                 font=self.f_small, bg=WHITE, fg=SLATE).grid(row=1, column=1, sticky="nw", pady=(2, 0))

        # 오른쪽 위: 명단 상태 알약
        status = tk.Frame(top, bg=WHITE)
        status.grid(row=0, column=2, rowspan=2, sticky="e")
        self.dict_dot = tk.Label(status, text="●", font=self.f_small, bg=WHITE, fg=STEEL)
        self.dict_dot.pack(side="left", padx=(0, 6))
        self.dict_label = tk.Label(status, text="", font=self.f_small, bg=WHITE, fg=SLATE)
        self.dict_label.pack(side="left")
        self.button(top, "명단 넣기", self.build_dictionary, height=34, padx=16).grid(
            row=0, column=3, rowspan=2, sticky="e", padx=(10, 0))

        # 둘째 줄: 단계 탭(가리기 → 되돌리기) + 작은 메뉴
        nav = tk.Frame(bar, bg=WHITE)
        nav.pack(fill="x", padx=32, pady=(0, 14))
        nav.grid_columnconfigure(1, weight=1)
        self.tab_seg = Segmented(nav, [("mask", "1  가리기", None), ("restore", "2  되돌리기", None)],
                                 self.tab_var, self.f_tab, self.f_tag, height=40, min_seg=150,
                                 command=self.show_tab)
        self.tab_seg.grid(row=0, column=0, sticky="w")
        tk.Label(nav, text="파일을 가려 AI에 올리고 → AI 답을 받아 되돌려요", font=self.f_small, bg=WHITE,
                 fg=STEEL).grid(row=0, column=1, sticky="w", padx=(16, 0))
        actions = tk.Frame(nav, bg=WHITE)
        actions.grid(row=0, column=2, sticky="e")
        self.version_label = tk.Label(actions, text=f"v{VERSION}", font=self.f_small, bg=WHITE, fg=STEEL,
                                      cursor=HAND_CURSOR)
        self.version_label.pack(side="right", padx=(10, 0))
        self.version_label.bind("<Button-1>", lambda _e: self.check_update(manual=True))
        self.button(actions, "도움말", self.show_usage, kind="quiet", height=32, padx=10).pack(side="right")
        self.button(actions, "예외 단어", self.show_ignored_words, kind="quiet", height=32, padx=10).pack(
            side="right", padx=(0, 2))
        self.button(actions, "결과 지우기", self.clear_outputs_now, kind="quiet", height=32, padx=10).pack(
            side="right", padx=(0, 2))
        tk.Frame(bar, bg=HAIR, height=1).pack(fill="x")

    def _build_copyright(self):
        """맨 아래: JEIO 로고와 저작권 문구. 로고 글자 밑줄과 문구 밑줄을 맞춘다."""
        foot = tk.Frame(self.root, bg=MIST)
        foot.pack(side="bottom", fill="x", padx=32, pady=(0, 12))
        font = self.f_small
        ascent = font.metrics("ascent")
        descent = font.metrics("descent")
        logo_path = program_dir() / "assets" / "jeio_logo.png"
        logo_w = logo_h = 0
        if logo_path.is_file():
            self._jeio_logo = tk.PhotoImage(data=base64.b64encode(logo_path.read_bytes()))
            logo_w, logo_h = self._jeio_logo.width(), self._jeio_logo.height()
        gap = 8 if logo_w else 0
        baseline = max(ascent, LOGO_BASELINE)
        canvas_h = max(baseline + descent, baseline - LOGO_BASELINE + logo_h, 1)
        canvas_w = logo_w + gap + font.measure(COPYRIGHT_NOTICE) + 6
        mark = tk.Canvas(foot, width=canvas_w, height=canvas_h, bg=MIST, highlightthickness=0, bd=0)
        mark.pack(side="left")
        if logo_w:
            mark.create_image(0, baseline - LOGO_BASELINE, image=self._jeio_logo, anchor="nw")
        mark.create_text(logo_w + gap, baseline + descent, text=COPYRIGHT_NOTICE, font=font,
                         fill=SLATE, anchor="sw")

    def _step_title(self, parent, number, text, sub=None):
        """카드 제목: 단계 번호 동그라미 + 제목 (+ 오른쪽 작은 설명). (줄, 번호) 를 돌려준다."""
        row = tk.Frame(parent, bg=WHITE)
        badge = StepBadge(row, number, self.f_small_bold)
        badge.pack(side="left", padx=(0, 10))
        tk.Label(row, text=text, font=self.f_title, bg=WHITE, fg=INK).pack(side="left")
        if sub:
            tk.Label(row, text=sub, font=self.f_small, bg=WHITE, fg=SLATE).pack(side="left", padx=(10, 0), pady=(3, 0))
        return row, badge

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

        # ---- 왼쪽: 1 파일 → 2 형식 → 3 가리기
        file_card = Card(left)
        file_card.pack(fill="x")
        box = file_card.inner
        row, self.step_file = self._step_title(box, 1, "파일 고르기")
        row.pack(anchor="w")

        # 점선 둥근 칸: 뒤에 캔버스로 테두리를 그리고 그 위에 안내를 얹는다
        self._drop_fill = DROP_BG
        self.drop = tk.Frame(box, bg=WHITE, height=168)
        self.drop.pack(fill="x", pady=(16, 12))
        self.drop.pack_propagate(False)
        self._drop_canvas = tk.Canvas(self.drop, bg=WHITE, highlightthickness=0, bd=0)
        self._drop_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self._drop_canvas.bind("<Configure>", lambda _e: self._draw_drop())
        inner = tk.Frame(self.drop, bg=DROP_BG)
        inner.place(relx=0.5, rely=0.5, anchor="center")
        self._drop_icon = tk.Canvas(inner, width=48, height=52, bg=DROP_BG, highlightthickness=0, bd=0)
        self._drop_icon.pack()
        hint = "여기에 파일을 끌어다 놓으세요" if HAS_DND else "가릴 파일을 고르세요"
        self.drop_label = tk.Label(inner, text=hint, font=self.f_bold, bg=DROP_BG, fg=INK)
        self.drop_label.pack(pady=(8, 0))
        self.drop_sub = tk.Label(inner, text="여러 개도 돼요 · pdf, 워드, 엑셀, 한글, 피피티",
                                 font=self.f_small, bg=DROP_BG, fg=SLATE)
        self.drop_sub.pack(pady=(4, 0))
        self.choose_button = self.button(inner, "파일 고르기", self.choose_mask_file, parent_bg=DROP_BG, height=34,
                                         padx=18)
        self.choose_button.pack(pady=(12, 0))
        self._drop_inner = inner
        if HAS_DND:
            self._enable_drop([self.drop, self._drop_canvas, inner, self._drop_icon, self.drop_label, self.drop_sub])
        self._drop_widgets = [inner, self._drop_icon, self.drop_label, self.drop_sub]

        self.file_label = tk.Label(box, textvariable=self.file_var, font=self.f_body, bg=WHITE, fg=SLATE,
                                   anchor="w", justify="left")
        self.file_label.pack(fill="x")
        self.file_label.bind("<Configure>", lambda e: self.file_label.configure(wraplength=max(60, e.width - 4)))

        mode_card = Card(left)
        mode_card.pack(fill="x", pady=(14, 0))
        box = mode_card.inner
        row, self.step_mode = self._step_title(box, 2, "어떻게 저장할까요")
        row.pack(anchor="w")
        self.mode_seg = Segmented(box, [("md", "AI에 올릴 글", "추천"), ("xlsx", "원래 형식", None)],
                                  self.mode_var, self.f_btn, self.f_tag, height=40, min_seg=150)
        self.mode_seg.pack(anchor="w", pady=(14, 8))
        self.mode_hint = tk.Label(box, text="AI에 보낼 때는 ‘AI에 올릴 글’을 고르세요.",
                                  font=self.f_small, bg=WHITE, fg=SLATE, anchor="w", justify="left")
        self.mode_hint.pack(fill="x")
        self.mode_hint.bind("<Configure>", lambda e: self.mode_hint.configure(wraplength=max(60, e.width - 4)))

        # 가릴 항목은 파일을 고르기 전에는 숨긴다. 고른 뒤에 그 파일에서 찾은 항목만 채운다.
        self.kind_wrap = tk.Frame(box, bg=WHITE)
        kind_head = tk.Frame(self.kind_wrap, bg=WHITE)
        kind_head.pack(fill="x", pady=(0, 4))
        tk.Label(kind_head, text="가릴 항목", font=self.f_small_bold, bg=WHITE, fg=INK).pack(side="left")
        self.kind_toggle = tk.Label(kind_head, text="선택 해제", font=self.f_small, bg=WHITE, fg=SLATE,
                                    cursor=HAND_CURSOR)
        self.kind_toggle.pack(side="right")
        self.kind_toggle.bind("<Button-1>", lambda _e: self._toggle_all_kinds())
        self.kind_note = tk.Label(self.kind_wrap, text="이 파일에서 찾은 항목이에요. 체크를 끄면 글에 그대로 남아요.",
                                  font=self.f_small, bg=WHITE, fg=SLATE, anchor="w", justify="left")
        self.kind_note.pack(fill="x", pady=(0, 8))
        self.kind_note.bind("<Configure>", lambda e: self.kind_note.configure(wraplength=max(60, e.width - 4)))
        self.kind_empty = tk.Label(self.kind_wrap, text="", font=self.f_small, bg=WHITE, fg=SLATE,
                                   anchor="w", justify="left")
        self.kind_empty.bind("<Configure>", lambda e: self.kind_empty.configure(wraplength=max(60, e.width - 4)))
        self.kind_grid = tk.Frame(self.kind_wrap, bg=WHITE)
        self.kind_grid.grid_columnconfigure(0, weight=1)
        self.kind_grid.grid_columnconfigure(1, weight=1)
        self._kind_ticks = []

        run_row = tk.Frame(left, bg=MIST)
        run_row.pack(fill="x", pady=(16, 0))
        self.run_button = self.button(run_row, "3  가리기", self.run_mask, kind="primary", height=50,
                                      padx=26, font=self.f_primary, parent_bg=MIST)
        self.run_button.pack(fill="x")
        self.run_hint = tk.Label(run_row, text="", font=self.f_small, bg=MIST, fg=SLATE, anchor="w",
                                 justify="left")
        self.run_hint.pack(fill="x", pady=(8, 0))
        self.run_hint.bind("<Configure>", lambda e: self.run_hint.configure(wraplength=max(60, e.width - 4)))

        # ---- 오른쪽: 결과
        result_card = Card(right, expand=True)
        result_card.pack(fill="both", expand=True)
        box = result_card.inner
        box.grid_columnconfigure(0, weight=1)
        box.grid_rowconfigure(5, weight=1)
        head = tk.Frame(box, bg=WHITE)
        head.grid(row=0, column=0, sticky="ew")
        tk.Label(head, text="결과", font=self.f_title, bg=WHITE, fg=INK).pack(side="left")

        self.status_wrap = tk.Frame(box, bg=WHITE)
        self.status_wrap.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        status_head = tk.Frame(self.status_wrap, bg=WHITE)
        status_head.pack(fill="x")
        self.status_dot = tk.Canvas(status_head, width=14, height=18, bg=WHITE, highlightthickness=0, bd=0)
        self.status_dot.pack(side="left", anchor="n")
        self.status_inner = tk.Frame(status_head, bg=WHITE)
        self.status_inner.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.status_inner.bind("<Configure>", lambda _e: self._reflow_status())
        self._status_line_labels = []
        self.status_path = tk.Label(self.status_wrap, text="", font=self.f_small, bg=WHITE, fg=SLATE,
                                    anchor="w", justify="left")
        self.status_path.bind("<Configure>",
                              lambda e: self.status_path.configure(wraplength=max(60, e.width - 4)))

        # 확인 필요 목록: 내용이 있을 때만 보인다
        self.review_wrap = tk.Frame(box, bg=WHITE)
        self.review_wrap.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        tk.Label(self.review_wrap, text="확인 필요", font=self.f_small_bold,
                 bg=WHITE, fg=MARK, anchor="w").pack(fill="x", pady=(0, 6))
        table, self.review_table = self.make_table(
            self.review_wrap, [("종류", 130, False), ("위치", 110, False), ("내용", 300, True)], 4)
        table.pack(fill="x")
        self.review_wrap.grid_remove()

        # 이름 의심 목록: 사전에 없지만 이름 같은 단어 (내용이 있을 때만 보인다)
        self.suspect_wrap = tk.Frame(box, bg=WHITE)
        self.suspect_wrap.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        self.suspect_count = tk.Label(self.suspect_wrap, text="확인할 이름", font=self.f_small_bold,
                                      bg=WHITE, fg=MARK, anchor="w")
        self.suspect_count.pack(fill="x", pady=(0, 4))
        self.suspect_list = tk.Frame(self.suspect_wrap, bg=WHITE)
        self.suspect_list.pack(fill="x")
        self.suspect_wrap.grid_remove()

        preview_head = tk.Frame(box, bg=WHITE)
        preview_head.grid(row=4, column=0, sticky="ew", pady=(16, 8))
        self.preview_title = tk.Label(preview_head, text="미리보기", font=self.f_small_bold, bg=WHITE, fg=INK,
                                      anchor="w")
        self.preview_title.pack(side="left")
        self.example_badge = tk.Label(preview_head, text=" 예시 ", font=self.f_tag, bg=CONTROL, fg=SLATE)
        self.example_badge.pack(side="left", padx=(8, 0))
        self.preview_note = tk.Label(preview_head, text="", font=self.f_small, bg=WHITE, fg=SLATE,
                                     anchor="w")
        self.preview_note.pack(side="left", padx=(12, 0))
        preview_box, self.preview = self.make_text(box, readonly=True)
        preview_box.grid(row=5, column=0, sticky="nsew")
        self.set_text(self.preview, EXAMPLE_PREVIEW, highlight=True)

        # 색 읽는 법: 미리보기에서 어떤 색이 무엇을 가린 것인지
        legend = tk.Frame(box, bg=WHITE)
        legend.grid(row=6, column=0, sticky="w", pady=(10, 0))
        styles = {tag: fg for tag, _, fg, _ in CODE_STYLES}
        for label, tag in LEGEND:
            item = tk.Frame(legend, bg=WHITE)
            item.pack(side="left", padx=(0, 14))
            tk.Label(item, text="■", font=self.f_tag, bg=WHITE, fg=styles[tag]).pack(side="left")
            tk.Label(item, text=label, font=self.f_tag, bg=WHITE, fg=SLATE).pack(side="left", padx=(3, 0))

        self.loading = tk.Frame(box, bg=WHITE)
        self.loading.grid(row=1, column=0, rowspan=6, sticky="nsew")
        loading_box = tk.Frame(self.loading, bg=WHITE)
        loading_box.place(relx=0.5, rely=0.46, anchor="center")
        self.loading_title = tk.Label(loading_box, text="가리는 중", font=self.f_title, bg=WHITE, fg=INK)
        self.loading_title.pack()
        self.loading_sub = tk.Label(loading_box, text="", font=self.f_body, bg=WHITE, fg=SLATE,
                                    justify="center")
        self.loading_sub.pack(pady=(12, 0))
        self.loading.grid_remove()

        buttons = tk.Frame(box, bg=WHITE)
        buttons.grid(row=7, column=0, sticky="ew", pady=(16, 0))
        self.copy_result_button = self.button(buttons, "결과 복사", self.copy_result, kind="primary", height=42,
                                              padx=24)
        self.copy_result_button.pack(side="left")
        self.copy_result_button.state(["disabled"])   # 결과를 만들기 전에는 복사할 것이 없다
        self.open_result_button = self.button(buttons, "결과 파일 열기", self.open_result_files, height=42)
        self.open_result_button.pack(side="left", padx=(10, 0))
        self.open_result_button.state(["disabled"])
        self.button(buttons, "결과 폴더", lambda: open_folder(workflows.AI_DIR), kind="quiet", height=42,
                    padx=14).pack(side="right")

        self.set_status("아직 결과가 없어요. 파일을 고르고 [가리기]를 눌러 주세요", "neutral")

    def _draw_drop(self):
        """점선 칸을 지금 색으로 다시 그린다 (끌어다 놓는 중이면 강조색)."""
        c = self._drop_canvas
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w <= 1 or h <= 1:
            return
        hot = self._drop_fill == DROP_HOVER
        round_rect(c, 1, 1, w - 1, h - 1, 16, fill=self._drop_fill)
        dashed_round_rect(c, 1, 1, w - 2, h - 2, 16, ACCENT if hot else "#BFC5CE", width=2 if hot else 1.5)
        self._drop_icon.delete("all")
        draw_doc_icon(self._drop_icon, 24, 26, ACCENT if hot else SLATE, ACCENT)

    def _selected_kinds(self):
        """지금 목록에서 체크된 항목 키. 목록이 없으면 빈 목록."""
        return [key for key, _ in MASK_OPTIONS if self._kind_vars.get(key) and self._kind_vars[key].get()]

    def _kinds_for_mask(self):
        """모두 켜져 있으면 전부 가린다. 일부를 끄면 체크된 것만 가린다."""
        chosen = self._selected_kinds()
        if chosen and len(chosen) == len(self._kind_vars):
            return None
        return frozenset(chosen)

    def _clear_kind_list(self):
        for child in self.kind_grid.winfo_children():
            child.destroy()
        self._kind_vars = {}
        self._kind_ticks = []

    def _fill_kind_grid(self, keys):
        self._clear_kind_list()
        labels = dict(MASK_OPTIONS)
        for i, key in enumerate(keys):
            var = tk.BooleanVar(value=True)
            self._kind_vars[key] = var
            cell = tk.Frame(self.kind_grid, bg=WHITE)
            cell.grid(row=i // 2, column=i % 2, sticky="w", pady=2)
            tick = Tick(cell, var, command=self._on_kind_change, parent_bg=WHITE)
            self._kind_ticks.append(tick)
            tick.pack(side="left")
            name = tk.Label(cell, text=labels[key], font=self.f_small, bg=WHITE, fg=INK, cursor=HAND_CURSOR)
            name.pack(side="left", padx=(6, 0))
            name.bind("<Button-1>", lambda _e, t=tick: t.set(not t.var.get()))
        self.kind_toggle.configure(text="선택 해제")

    def _refresh_found_kinds(self):
        """고른 파일에서 가릴 수 있는 항목만 목록으로 보여 준다. 파일이 없으면 목록을 숨긴다."""
        paths = self._mask_paths()
        self._clear_kind_list()
        if not paths or not self._dict_ok:
            self._kind_status = "hidden"
            self.kind_wrap.pack_forget()
            return
        if not self.kind_wrap.winfo_ismapped():
            self.kind_wrap.pack(fill="x", pady=(16, 0))
        self.kind_note.pack_forget()
        self.kind_grid.pack_forget()
        self.kind_toggle.pack_forget()
        self.kind_empty.configure(text="파일에서 가릴 항목을 찾고 있어요.")
        self.kind_empty.pack(fill="x")
        self.run_hint.configure(text="파일에서 가릴 항목을 찾고 있어요.")
        self.root.update_idletasks()
        try:
            found, errors = workflows.kinds_in_files(paths)
        except Exception as exc:
            found, errors = set(), [str(exc)]
        keys = [key for key, _ in MASK_OPTIONS if key in found]
        if not keys:
            self._kind_status = "error" if errors else "empty"
            message = "파일을 읽지 못해서 가릴 항목을 못 찾았어요." if errors else "이 파일에서 가릴 항목을 찾지 못했어요."
            self.kind_empty.configure(text=message)
            return
        self._kind_status = "ready"
        self.kind_empty.pack_forget()
        self.kind_toggle.pack(side="right")
        self.kind_note.pack(fill="x", pady=(0, 8))
        self.kind_grid.pack(fill="x")
        self._fill_kind_grid(keys)

    def _on_kind_change(self):
        chosen = self._selected_kinds()
        self.kind_toggle.configure(text="선택 해제" if chosen and len(chosen) == len(self._kind_vars) else "모두 선택")
        self._update_run_hint()
        self._mark_result_stale("가릴 항목을 바꿨어요. [가리기]를 다시 눌러 주세요.")

    def _toggle_all_kinds(self):
        if not self._kind_vars:
            return
        turn_on = len(self._selected_kinds()) != len(self._kind_vars)
        for var in self._kind_vars.values():
            var.set(turn_on)
        self._redraw_kind_ticks()
        self._on_kind_change()

    def _redraw_kind_ticks(self):
        for tick in self._kind_ticks:
            tick._draw()

    def _update_run_hint(self):
        """[가리기] 아래 한 줄: 무엇이 빠져서 누를 수 없는지, 또는 지금 고른 것."""
        paths = self._mask_paths()
        kinds = self._selected_kinds()
        off = [label for key, label in MASK_OPTIONS if key in self._kind_vars and not self._kind_vars[key].get()]
        if not self._dict_ok:
            text = "먼저 위쪽 [명단 넣기]로 학생 명렬표를 넣어 주세요."
        elif not paths:
            text = "파일을 고르면 그 안에 있는 가릴 항목이 나와요."
        elif self._kind_status == "error":
            text = "파일을 읽지 못해서 가릴 항목을 못 찾았어요."
        elif not self._kind_vars:
            text = "이 파일에서 가릴 항목을 찾지 못했어요."
        elif not kinds:
            text = "가릴 항목을 하나 이상 체크해 주세요."
        elif off:
            shown = ", ".join(off[:4])
            more = f" 외 {len(off) - 4}개" if len(off) > 4 else ""
            text = f"그대로 둘 항목: {shown}{more}. 원본 파일은 바뀌지 않아요."
        else:
            text = "가린 결과는 오른쪽에 나와요. 원본 파일은 바뀌지 않아요."
        self.run_hint.configure(text=text)
        self.step_file.set(done=bool(paths))
        self.step_mode.set(done=bool(paths) and bool(kinds))

    def _enable_drop(self, widgets):
        for w in widgets:
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            w.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            w.dnd_bind("<<Drop>>", self.safe(self._on_drop))

    def _paint_drop(self, color):
        self._drop_fill = color
        for w in self._drop_widgets:
            w.configure(bg=color)
        self.choose_button.set_parent_bg(color)
        self._draw_drop()

    def _on_drop_enter(self, event):
        self._paint_drop(DROP_HOVER)
        return event.action

    def _on_drop_leave(self, event):
        self._paint_drop(DROP_BG)
        return event.action

    def _on_drop(self, event):
        """끌어다 놓은 파일을 고른 것으로 처리한다 (처리 실행은 [가명처리 실행]을 눌러야 시작)."""
        self._paint_drop(DROP_BG)
        files = self.root.tk.splitlist(event.data)
        if files:
            self.select_mask_files(files)
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
        row, _ = self._step_title(box, 1, "AI 답변 붙여 넣기")
        row.grid(row=0, column=0, sticky="w")
        tk.Label(box, text="AI가 돌려준 글을 그대로 붙여 넣으세요. 학생003 같은 코드가 실명으로 돌아와요.",
                 font=self.f_small, bg=WHITE, fg=SLATE, anchor="w", justify="left").grid(
            row=1, column=0, sticky="ew", pady=(8, 12))
        input_box, self.input_text = self.make_text(box)
        input_box.grid(row=2, column=0, sticky="nsew")
        self.attach_placeholder(self.input_text, INPUT_HINT)
        buttons = tk.Frame(box, bg=WHITE)
        buttons.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        self.restore_button = self.button(buttons, "되돌리기", self.restore_from_box, kind="primary", height=46,
                                          padx=30, font=self.f_primary)
        self.restore_button.pack(side="left")
        self.restore_clip_button = self.button(buttons, "복사한 글 바로 되돌리기", self.restore_from_clipboard,
                                               height=46, padx=16)
        self.restore_clip_button.pack(side="left", padx=(10, 0))
        self.restore_file_button = self.button(buttons, "파일로", self.restore_from_file, kind="quiet",
                                               height=46, padx=12)
        self.restore_file_button.pack(side="right")

        # 오른쪽: 복원 결과
        right = Card(tab, expand=True)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        box = right.inner
        box.grid_columnconfigure(0, weight=1)
        box.grid_rowconfigure(2, weight=1)
        row, _ = self._step_title(box, 2, "되돌린 글")
        row.grid(row=0, column=0, sticky="w")
        note = tk.Frame(box, bg="#FBF3E8")
        note.grid(row=1, column=0, sticky="ew", pady=(10, 12))
        tk.Label(note, text="실명이 들어 있어요. 이 글은 AI에 다시 올리지 마세요.", font=self.f_small,
                 bg="#FBF3E8", fg=MARK, anchor="w").pack(fill="x", padx=12, pady=7)
        output_box, self.output_text = self.make_text(box, readonly=True)
        output_box.grid(row=2, column=0, sticky="nsew")
        self.restored = ""
        self.set_text(self.output_text, "", placeholder=OUTPUT_HINT)
        self.copy_restored_button = self.button(box, "복사", self.copy_restored, height=46, padx=28)
        self.copy_restored_button.grid(row=3, column=0, sticky="w", pady=(16, 0))

        # 아래: 확인 필요 (복원 못 한 코드가 있을 때만 보인다)
        self.warn_card = Card(tab)
        self.warn_card.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        box = self.warn_card.inner
        tk.Label(box, text="확인 필요 · 되돌리지 못한 코드", font=self.f_title, bg=WHITE, fg=MARK,
                 anchor="w").pack(fill="x", pady=(0, 8))
        table, self.warn_table = self.make_table(box, [("가린 글", 160, False), ("이유", 600, True)], 5)
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
        win.title("문제")
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text="문제가 생겼어요", font=self.f_title, bg=WHITE, fg=INK).pack(anchor="w")
        tk.Label(box, text=short, font=self.f_small, bg=WHITE, fg=SLATE,
                 justify="left", wraplength=420, anchor="w").pack(anchor="w", pady=(8, 16))
        row = tk.Frame(box, bg=WHITE)
        row.pack(fill="x")
        PillButton(row, "내용 복사", lambda: self._copy_error_report(report),
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
        win.title("내용 복사")
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text="복사할 내용", font=self.f_title, bg=WHITE, fg=INK).pack(anchor="w")
        tk.Label(box, text="학생 이름이 없는지 본 뒤, 만든 사람에게 보내 주세요.",
                 font=self.f_small, bg=WHITE, fg=MARK,
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
            messagebox.showinfo(APP_TITLE, "내용을 복사했어요.")

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

    def _mask_wait_copy(self):
        """가리는 동안 보여줄 안내. 원래 형식이면 파일이 언제 준비되는지도 적는다."""
        if self.mode_var.get() != "xlsx":
            return "끝나면 여기에 AI에 올릴 글이 나와요."
        exts = {p.suffix.lower() for p in self._mask_paths()}
        if exts and exts <= {".hwp", ".hwpx"}:
            kind = "한글 파일"
        elif exts <= {".xlsx", ".xls"}:
            kind = "엑셀 파일"
        else:
            kind = "엑셀·한글 파일"
        return f"끝나면 {kind}이 준비돼요.\n[결과 파일 열기]로 원래 모양을 보세요."

    def _show_mask_loading(self, again=False):
        self.example_badge.pack_forget()
        self.review_wrap.grid_remove()
        self.suspect_wrap.grid_remove()
        self.preview_note.configure(text="")
        self.set_text(self.preview, "")
        self.loading_title.configure(text="다시 가리는 중" if again else "가리는 중")
        self.loading_sub.configure(text=self._mask_wait_copy())
        self.loading.grid()
        self.loading.tkraise()
        self.run_button.state(["disabled"])
        self.root.configure(cursor="watch")
        self.set_status("가리는 중...", "neutral")
        self.root.update()

    def _hide_mask_loading(self):
        self.loading.grid_remove()
        self.run_button.state(["!disabled"])
        self.root.configure(cursor="")

    def copy_to_clipboard(self, text, button):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        old = button.cget("text")
        button.configure(text="복사했어요")
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
            info, problem = None, f"명단을 읽을 수 없어요 - {exc}"
        if info is None:
            self.dict_label.configure(text=problem, fg=STATUS_COLORS["bad"][0])
            self.dict_dot.configure(fg=STATUS_COLORS["bad"][0])
        else:
            students, staff, modified = info
            staff_text = "교직원 명단 없음" if staff == 0 else f"교직원 {staff}명"
            self.dict_label.configure(text=f"학생 {students}명 · {staff_text} · {modified}", fg=SLATE)
            self.dict_dot.configure(fg=STATUS_COLORS["ok"][0])
        for b in (self.run_button, self.restore_button, self.restore_clip_button, self.restore_file_button):
            b.state(["!disabled"] if info else ["disabled"])
        self._dict_ok = info is not None
        self._update_run_hint()

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
        self._mark_result_stale(f"{noun} 명단을 넣었어요. [가리기]를 다시 눌러 주세요.")

        lines = [f"{noun} {result.count}명을 넣었어요."]
        if kind == "학생":
            lines += [f"찾은 칸: {', '.join(result.found)}",
                      f"없는 칸: {', '.join(result.missing) if result.missing else '없음'}"]
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
        paths = filedialog.askopenfilenames(title="가릴 파일 선택", filetypes=MASK_FILETYPES)
        if paths:
            self.select_mask_files(paths)

    def select_mask_file(self, path, extra=0):
        self.select_mask_files([path])

    def select_mask_files(self, paths):
        ok, bad = [], []
        for raw in paths:
            path = Path(raw)
            if path.suffix.lower() in workflows.MASK_EXTS:
                ok.append(path)
            else:
                bad.append(path)
        if not ok:
            messagebox.showwarning(APP_TITLE, "이 파일은 쓸 수 없어요.\n"
                                              "pdf, 워드, 엑셀, 한글, 피피티, txt, md")
            return
        if bad:
            names = ", ".join(p.name for p in bad[:5])
            more = f" 외 {len(bad) - 5}개" if len(bad) > 5 else ""
            messagebox.showwarning(APP_TITLE, f"쓸 수 없는 파일은 빼 두었어요: {names}{more}")
        self.selected_paths = ok
        self.selected_path = ok[0]
        extra = len(ok) - 1
        self.file_var.set(f"고른 파일: {ok[0].name}" + (f" 외 {extra}개" if extra else ""))
        self.file_label.configure(fg=INK)
        self._confirmed_names = []
        self._ignored_names = set()
        # 엑셀·한글은 같은 형식 파일로 저장한다. 그 외는 글(md)만 된다.
        if all(p.suffix.lower() in {".xlsx", ".xls", ".hwp", ".hwpx"} for p in ok):
            self.mode_seg.set_enabled("xlsx", True)
            self.mode_var.set("xlsx")
            self.mode_hint.configure(text="엑셀·한글은 같은 형식 파일로 저장돼요. AI에 보낼 때만 ‘AI에 올릴 글’을 고르세요.")
        else:
            self.mode_var.set("md")
            self.mode_seg.set_enabled("xlsx", False)
            self.mode_hint.configure(text="이 파일은 글(md)로만 저장돼요. 엑셀·한글만 원래 형식을 유지해요.")
        self._refresh_found_kinds()
        self._update_run_hint()
        self._mark_result_stale("새 파일을 골랐어요. [가리기]를 눌러 주세요.")

    def _mask_paths(self):
        return list(self.selected_paths) or ([self.selected_path] if self.selected_path else [])

    def _mark_result_stale(self, message):
        """이전 결과가 새 파일/새 사전과 맞지 않으니, 복사 버튼을 잠그고 안내만 보여 준다."""
        if self.result is None and self.copy_result_button.instate(["disabled"]):
            return   # 아직 결과를 만든 적이 없으면 처음 화면(예시)을 그대로 둔다
        self.result = None
        self.copy_result_button.state(["disabled"])
        self.open_result_button.state(["disabled"])
        self.suspect_wrap.grid_remove()
        self.set_status(message, "neutral")

    def run_mask(self):
        paths = self._mask_paths()
        if not paths:
            messagebox.showinfo(APP_TITLE, "먼저 가릴 파일을 고르세요.")
            return
        if not self._kind_vars:
            message = ("파일을 읽지 못해서 가릴 항목을 못 찾았어요."
                       if self._kind_status == "error" else "이 파일에서 가릴 항목을 찾지 못했어요.")
            messagebox.showinfo(APP_TITLE, message)
            return
        kinds = self._kinds_for_mask()
        if not kinds and kinds is not None:
            messagebox.showinfo(APP_TITLE, "가릴 항목을 하나 이상 체크해 주세요.")
            return
        dictionary = workflows.get_dictionary()
        self._confirmed_names = []
        self._ignored_names = set()
        self.copy_result_button.state(["disabled"])
        self.open_result_button.state(["disabled"])
        self._show_mask_loading()
        try:
            result = workflows.mask_files(paths, self.mode_var.get(), dictionary, kinds=kinds)
        except Exception:
            self.result = None
            self.open_result_button.state(["disabled"])
            self.preview_title.configure(text="미리보기")
            self.set_status("⚠ 가리지 못했어요. 알림창의 내용을 확인해 주세요", "bad")
            raise
        finally:
            self._hide_mask_loading()
        self.show_mask_result(result)

    def show_mask_result(self, result):
        self.example_badge.pack_forget()   # 여기부터는 진짜 결과
        self.result = result
        rows = [(say_kind(kind), where, text) for kind, where, text in result.review]
        rows += [(f"남은 {say_kind(kind)}", where, value) for where, kind, value in result.findings]
        self.fill_table(self.review_table, rows, self.review_wrap)
        files = self._result_files()
        if self.mode_var.get() == "xlsx" and files:
            exts = {p.suffix.lower() for p in files}
            if exts <= {".hwp", ".hwpx"}:
                kind = "한글 파일"
            elif exts <= {".xlsx", ".xls"}:
                kind = "엑셀 파일"
            else:
                kind = "엑셀·한글 파일"
            self.preview_title.configure(text="글 미리보기")
            self.preview_note.configure(text=f"{kind}은 [결과 파일 열기]로 보세요.")
        else:
            self.preview_title.configure(text="미리보기")
            self.preview_note.configure(text="")
        # 미리보기에서만 md 표 구분선(| --- | --- |)을 숨긴다. 복사·저장되는 result.text 는 그대로 둔다.
        display = self._hide_md_separators(result.text) if result.ok else ""
        self.set_text(self.preview, display, highlight=True)
        # 남은 개인정보가 있으면 복사를 막는다. 이름 의심은 복사를 막지 않는다(복사 때 안내만).
        self.copy_result_button.state(["!disabled"] if result.can_copy else ["disabled"])
        self.open_result_button.state(["!disabled"] if self._result_files() else ["disabled"])
        self._render_suspects(self._pending_suspects())
        self._set_mask_status()

    def _result_files(self):
        """이번에 만든 결과 파일. 없는 경로는 빼 둔다."""
        if not self.result or not self.result.ok:
            return []
        paths = list(self.result.ai_paths or [])
        if not paths and self.result.ai_path:
            paths = [self.result.ai_path]
        return [Path(p) for p in paths if p and Path(p).exists()]

    def open_result_files(self):
        """가린 파일을 엑셀·한글 등 기본 프로그램으로 연다."""
        paths = self._result_files()
        if not paths:
            messagebox.showinfo(APP_TITLE, "열 결과 파일이 없어요.")
            return
        if len(paths) > 5:
            open_folder(paths[0].parent)
            self.toast(f"결과 파일이 {len(paths)}개라 폴더를 열었어요")
            return
        for path in paths:
            open_path(path)

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
            tone = "warn" if "그대로 둔 항목" in head_text else "ok"
            lines = [(head_text, tone)]
            if n:
                extra = f" · 확인 필요 {len(r.review)}건" if r.review else ""
                lines.append((f"⚠ 확인할 이름 {n}건 - 아래에서 가릴지 정해 주세요{extra}", "warn"))
            elif r.review:
                lines.append((f"⚠ 확인 필요 {len(r.review)}건 - 아래 목록을 봐 주세요", "warn"))
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
        self.suspect_count.configure(text=f"확인할 이름 {len(suspects)}건 · 명단에 없는 단어")
        cap = 12
        for s in suspects[:cap]:
            row = tk.Frame(self.suspect_list, bg=WHITE)
            row.pack(fill="x", pady=(0, 12))
            head = tk.Frame(row, bg=WHITE)
            head.pack(fill="x")
            tk.Label(head, text=s.word, font=self.f_bold, bg=WHITE, fg=INK).pack(side="left")
            context = tk.Label(head, text=s.context, font=self.f_small, bg=WHITE, fg=SLATE,
                               anchor="w", justify="left")
            context.pack(side="left", fill="x", expand=True, padx=(12, 0))
            head.bind("<Configure>", lambda e, c=context: c.configure(wraplength=max(80, e.width - 140)))
            buttons = tk.Frame(row, bg=WHITE)
            buttons.pack(anchor="e", pady=(8, 0))
            self.button(buttons, "이 이름도 가리기", lambda w=s.word: self._confirm_suspect(w),
                        kind="primary", height=32, padx=12).pack(side="left")
            self.button(buttons, "그냥 두기", lambda w=s.word: self._ignore_suspect(w),
                        height=32, padx=12).pack(side="left", padx=(8, 0))
            self.button(buttons, "계속 빼 두기", lambda w=s.word: self._always_ignore_suspect(w),
                        height=32, padx=12).pack(side="left", padx=(8, 0))
        if len(suspects) > cap:
            tk.Label(self.suspect_list, text=f"… 외 {len(suspects) - cap}건은 미리보기에서 봐 주세요",
                     font=self.f_small, bg=WHITE, fg=SLATE, anchor="w").pack(fill="x", pady=(2, 0))

    def _confirm_suspect(self, word):
        """[이 단어도 가리기]: 그 단어를 [이름N] 으로 바꾸고 결과를 다시 만든다."""
        paths = self._mask_paths()
        if not paths:
            return
        if word not in self._confirmed_names:
            self._confirmed_names.append(word)
        self._ignored_names.discard(word)
        self._show_mask_loading(again=True)
        try:
            result = workflows.mask_files(paths, self.mode_var.get(),
                                          workflows.get_dictionary(), extra_names=self._confirmed_names,
                                          kinds=self._kinds_for_mask())
        finally:
            self._hide_mask_loading()
        self.show_mask_result(result)

    def show_usage(self):
        """상단 [사용 설명]: 선생님용 짧은 안내를 기본 앱으로 연다."""
        from core import config
        path = config.program_dir() / "사용설명서.md"
        if not path.exists():
            messagebox.showinfo(APP_TITLE, "도움말 파일을 찾지 못했어요.")
            return
        if sys.platform.startswith("win"):
            os.startfile(path)
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])

    def _ignore_suspect(self, word):
        """[무시]: 이 단어를 의심 목록에서만 뺀다 (글은 그대로 둔다)."""
        self._ignored_names.add(word)
        self._render_suspects(self._pending_suspects())
        self._set_mask_status()

    def _always_ignore_suspect(self, word):
        """[항상 무시]: 확인 후 data/무시단어.txt 에 저장하고, 이번 목록에서도 뺀다."""
        if not messagebox.askyesno(
                APP_TITLE,
                f"'{word}'를 앞으로도 이름으로 보지 않을까요? 사람 이름이면 누르지 마세요."):
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
        win.title("예외 단어")
        win.configure(bg=MIST)
        win.transient(self.root)
        win.resizable(False, False)

        card = Card(win, parent_bg=MIST, expand=True)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        box = card.inner
        tk.Label(box, text="예외 단어", font=self.f_title, bg=WHITE, fg=INK).pack(anchor="w")
        tk.Label(box, text="이름으로 보지 않기로 한 단어예요. 고른 뒤 지울 수 있어요.",
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
        empty = tk.Label(box, text="빼 둔 단어가 없어요.", font=self.f_small,
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
            if not messagebox.askyesno(APP_TITLE, f"고른 {len(picked)}개를 예외 단어에서 지울까요?"):
                return
            workflows.remove_ignored_words(picked)
            refresh()

        buttons = tk.Frame(box, bg=WHITE)
        buttons.pack(fill="x", pady=(16, 0))
        self.button(buttons, "고른 단어 지우기", delete_selected, height=36, padx=16).pack(side="left")
        self.button(buttons, "닫기", win.destroy, height=36, padx=16).pack(side="left", padx=(8, 0))
        refresh()

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - w) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - h) // 3)
        win.geometry(f"+{x}+{y}")

    def copy_result(self):
        if self.result is None or not self.result.can_copy:
            messagebox.showwarning(APP_TITLE, "복사할 결과가 없어요.")
            return
        pending = self._pending_suspects()
        if pending:
            messagebox.showwarning(
                APP_TITLE,
                f"아직 정하지 않은 이름 {len(pending)}건이 있어요.\n"
                "그대로 복사합니다. 필요하면 [이 이름도 가리기]로 가린 뒤 다시 복사하세요.")
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
            messagebox.showinfo(APP_TITLE, "복사한 글이 없어요.")
            return
        restored, _, warnings = unmask_text(text, workflows.get_dictionary(), self._name_map())
        self._show_restored(restored, warnings)
        self.root.clipboard_clear()
        self.root.clipboard_append(restored)
        self.root.update()

    def restore_from_file(self):
        path = filedialog.askopenfilename(
            title="되돌릴 파일 선택",
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
        shown = result.text if result.text is not None else f"엑셀 파일을 되돌렸어요.\n저장 위치: {saved}"
        self._show_restored(shown, result.warnings)
        messagebox.showinfo(APP_TITLE, f"되돌렸어요 ({result.changes}건, 경고 {len(result.warnings)}건)\n"
                                       f"저장 위치: {saved}")

    def copy_restored(self):
        text = self.restored
        if not text.strip():
            messagebox.showinfo(APP_TITLE, "되돌린 글이 없어요.")
            return
        self.copy_to_clipboard(text, self.copy_restored_button)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
