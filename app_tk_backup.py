#!/usr/bin/env python3
"""개인정보 가리개 - 화면.   실행: python3 app.py

화면은 이 파일, 실제 작업은 workflows.py 가 core/·adapters/ 의 함수를 불러서 한다.
- 결과를 처음 만들기 전까지 결과 영역에는 모양을 보기 위한 "(예시)" 글자가 보인다.
- 어떤 오류가 나도 프로그램이 꺼지지 않고, 알림창으로 오류 내용을 보여 준다.
"""
import os
import re
import subprocess
import sys
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog
from tkinter import font as tkfont
from tkinter import messagebox, ttk

import workflows
from core.restorer import unmask_text

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:   # tkinterdnd2 가 없으면 끌어다 놓기만 빠지고 나머지는 그대로 동작
    HAS_DND = False

APP_TITLE = "개인정보 가리개"
WINDOW_SIZE = "900x700"

# 맥·윈도우·리눅스에서 한글이 잘 보이는 글꼴 후보 (앞에서부터 설치된 것을 쓴다)
FONT_CANDIDATES = [
    "Apple SD Gothic Neo", "AppleGothic",       # 맥
    "맑은 고딕", "Malgun Gothic",                 # 윈도우
    "Noto Sans CJK KR", "NanumGothic",           # 리눅스
]

COLORS = {
    "ok": "#1B7F3B",
    "bad": "#C62828",
    "muted": "#6B7280",
    "drop_bg": "#F3F7FD",
    "drop_hover": "#DCE9FB",
    "drop_border": "#7FA0D0",
}

# 미리보기에서 색을 다르게 보여 줄 코드들 (태그 이름, 정규식, 글자색, 배경색)
CODE_STYLES = [
    ("code_unknown", r"\[(?:전화번호|주민번호|이메일)\]", "#B71C1C", "#FDE7E7"),
    ("code_bracket", r"\[(?:학생전화|보호자전화|주소|생년월일)\d{3,}(?:또는\d{3,})*\]", "#6A1B9A", "#F3E5F9"),
    ("code_sid", r"S\d{3,}(?:또는\d{3,})*", "#B45309", "#FEF0DB"),
    ("code_name", r"(?:학생|보호자)\d{3,}(?:또는\d{3,})*", "#1D4ED8", "#E3ECFD"),
]
CODE_REGEX = re.compile("|".join(f"(?P<{tag}>{pat})" for tag, pat, _, _ in CODE_STYLES))

NO_DICT_MESSAGE = "학생사전 없음 - 먼저 사전을 만들어 주세요"
MASK_FILETYPES = [("지원하는 파일", "*.pdf *.docx *.xlsx *.xls *.pptx *.csv *.txt *.md *.hwp *.hwpx"),
                  ("모든 파일", "*.*")]


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


class App:
    def __init__(self):
        self.root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(800, 640)

        self.family = pick_font_family(self.root)
        self._setup_fonts_and_styles()

        self.file_var = tk.StringVar(value="선택된 파일: (없음)")
        self.mode_var = tk.StringVar(value="md")

        self._build_top_bar()
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.mask_tab = ttk.Frame(notebook, padding=10)
        self.restore_tab = ttk.Frame(notebook, padding=10)
        notebook.add(self.mask_tab, text="  가리기  ")
        notebook.add(self.restore_tab, text="  되돌리기  ")
        self._build_mask_tab()
        self._build_restore_tab()

        self.selected_path = None
        self.result = None
        self.root.report_callback_exception = self._on_callback_error   # 어떤 오류도 알림창으로
        self.refresh_dictionary()

    # ------------------------------------------------------------ 글꼴·스타일
    def _setup_fonts_and_styles(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            tkfont.nametofont(name).configure(family=self.family, size=12)
        self.font = (self.family, 12)
        self.font_bold = (self.family, 12, "bold")
        self.font_big = (self.family, 16, "bold")
        self.font_drop = (self.family, 15)

        style = ttk.Style(self.root)
        style.configure("TNotebook.Tab", font=(self.family, 13, "bold"), padding=(14, 6))
        style.configure("Big.TButton", font=(self.family, 15, "bold"), padding=(10, 12))
        style.configure("TButton", font=self.font, padding=(8, 5))
        style.configure("Treeview", font=self.font, rowheight=26)
        style.configure("Treeview.Heading", font=self.font_bold)
        style.configure("TLabelframe.Label", font=self.font_bold)
        style.configure("TRadiobutton", font=self.font)

    # ------------------------------------------------------------ 공통 조각
    def button(self, parent, text, command, style="TButton"):
        return ttk.Button(parent, text=text, style=style, command=self.safe(command))

    def make_text(self, parent, height, readonly=False):
        """스크롤바가 붙은 글상자. (frame, text) 를 돌려준다."""
        frame = ttk.Frame(parent)
        text = tk.Text(frame, height=height, wrap="word", font=self.font, relief="solid", borderwidth=1,
                       padx=8, pady=6, undo=not readonly)
        bar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)
        return frame, text

    def make_table(self, parent, columns, height):
        """(종류, 위치, 내용) 같은 열이 있는 목록. (frame, tree) 를 돌려준다."""
        frame = ttk.Frame(parent)
        tree = ttk.Treeview(frame, columns=[c for c, _ in columns], show="headings", height=height)
        for name, width in columns:
            tree.heading(name, text=name)
            tree.column(name, width=width, anchor="w")
        bar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        return frame, tree

    def set_text(self, text_widget, content, highlight=False):
        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")
        text_widget.insert("1.0", content)
        if highlight:
            self.highlight_codes(text_widget)
        text_widget.configure(state="disabled")

    def highlight_codes(self, text_widget):
        """학생003, S003, [학생전화003] 같은 코드에 색을 입힌다."""
        for tag, _, fg, bg in CODE_STYLES:
            text_widget.tag_configure(tag, foreground=fg, background=bg, font=self.font_bold)
            text_widget.tag_remove(tag, "1.0", "end")
        for line_no, line in enumerate(text_widget.get("1.0", "end-1c").split("\n"), start=1):
            for m in CODE_REGEX.finditer(line):
                text_widget.tag_add(m.lastgroup, f"{line_no}.{m.start()}", f"{line_no}.{m.end()}")

    # ------------------------------------------------------------ 맨 위 줄
    def _build_top_bar(self):
        bar = ttk.Frame(self.root, padding=(12, 10))
        bar.pack(fill="x")
        self.dict_label = ttk.Label(bar, text="📘 학생사전: 30명 (2026-10-18)", font=self.font_big)
        self.dict_label.pack(side="left")
        self.button(bar, "사전 새로 만들기", self.build_dictionary).pack(side="right")

    # ------------------------------------------------------------ 탭 1: 가리기
    def _build_mask_tab(self):
        tab = self.mask_tab

        # 1. 끌어다 놓는 영역
        self.drop = tk.Frame(tab, bg=COLORS["drop_bg"], highlightthickness=2,
                             highlightbackground=COLORS["drop_border"], height=130)
        self.drop.pack(fill="x")
        self.drop.pack_propagate(False)
        inner = tk.Frame(self.drop, bg=COLORS["drop_bg"])
        inner.place(relx=0.5, rely=0.5, anchor="center")
        hint = "여기에 파일을 끌어다 놓으세요" if HAS_DND else "여기에 파일을 끌어다 놓으세요 (tkinterdnd2 없음)"
        self.drop_label = tk.Label(inner, text=hint, font=self.font_drop, bg=COLORS["drop_bg"], fg="#334155")
        self.drop_label.pack(pady=(0, 8))
        self.button(inner, "파일 선택", self.choose_mask_file).pack()
        ttk.Label(tab, textvariable=self.file_var, foreground=COLORS["muted"]).pack(anchor="w", pady=(6, 8))
        if HAS_DND:
            self._enable_drop([self.drop, inner, self.drop_label])

        # 2. 방식 선택
        mode = ttk.Frame(tab)
        mode.pack(fill="x")
        ttk.Label(mode, text="방식:", font=self.font_bold).pack(side="left", padx=(0, 10))
        self.radio_md = ttk.Radiobutton(mode, text="AI용 md (추천)", value="md", variable=self.mode_var)
        self.radio_md.pack(side="left", padx=(0, 18))
        self.radio_xlsx = ttk.Radiobutton(mode, text="엑셀 원본 형식 유지", value="xlsx", variable=self.mode_var)
        self.radio_xlsx.pack(side="left")

        # 3. 실행 버튼
        self.run_button = self.button(tab, "가명처리 실행", self.run_mask, style="Big.TButton")
        self.run_button.pack(fill="x", pady=10)

        # 4. 결과 영역
        result = ttk.LabelFrame(tab, text=" 결과 ", padding=10)
        result.pack(fill="both", expand=True)

        self.check_label = tk.Label(result, text="✅ 남은 개인정보 없음  (예시)", font=self.font_bold,
                                    fg=COLORS["ok"], anchor="w")
        self.check_label.pack(fill="x")

        ttk.Label(result, text="확인 필요 목록 (사전에없음 / 동명이인확인)", foreground=COLORS["muted"]).pack(
            anchor="w", pady=(8, 2))
        table_frame, self.review_table = self.make_table(result, [("종류", 120), ("위치", 200), ("내용", 480)], 3)
        table_frame.pack(fill="x")
        self.review_table.insert("", "end", values=("사전에없음", "(예시) 3번째 줄", "010-0000-0000 → [전화번호]"))
        self.review_table.insert("", "end", values=("동명이인확인", "(예시) 7번째 줄", "정하윤 → 학생007또는026"))

        ttk.Label(result, text="미리보기", foreground=COLORS["muted"]).pack(anchor="w", pady=(8, 2))
        preview_frame, self.preview = self.make_text(result, 5, readonly=True)
        preview_frame.pack(fill="both", expand=True)
        self.set_text(self.preview, "학생003은 S003이며 연락처는 [학생전화003]입니다.", highlight=True)

        buttons = ttk.Frame(result)
        buttons.pack(fill="x", pady=(8, 0))
        self.copy_result_button = self.button(buttons, "결과 내용 복사", self.copy_result)
        self.copy_result_button.pack(side="left")
        self.copy_result_button.state(["disabled"])   # 결과를 만들기 전에는 복사할 것이 없다
        self.button(buttons, "결과 폴더 열기", lambda: open_folder(workflows.AI_DIR)).pack(side="left", padx=8)

    def _enable_drop(self, widgets):
        for w in widgets:
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            w.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            w.dnd_bind("<<Drop>>", self.safe(self._on_drop))

    def _paint_drop(self, color):
        for w in (self.drop, self.drop.winfo_children()[0], self.drop_label):
            w.configure(bg=color)

    def _on_drop_enter(self, event):
        self._paint_drop(COLORS["drop_hover"])
        return event.action

    def _on_drop_leave(self, event):
        self._paint_drop(COLORS["drop_bg"])
        return event.action

    def _on_drop(self, event):
        """끌어다 놓은 파일 이름만 보여 준다 (처리는 하지 않음)."""
        self._paint_drop(COLORS["drop_bg"])
        files = self.root.tk.splitlist(event.data)
        if files:
            self.select_mask_file(files[0], extra=len(files) - 1)
        return event.action

    # ------------------------------------------------------------ 탭 2: 되돌리기
    def _build_restore_tab(self):
        tab = self.restore_tab

        ttk.Label(tab, text="AI 답변을 붙여 넣으세요", font=self.font_bold).pack(anchor="w", pady=(0, 4))
        input_frame, self.input_text = self.make_text(tab, 8)
        input_frame.pack(fill="both", expand=True)

        buttons = ttk.Frame(tab)
        buttons.pack(fill="x", pady=8)
        self.restore_button = self.button(buttons, "복원", self.restore_from_box, style="Big.TButton")
        self.restore_button.pack(side="left")
        self.restore_clip_button = self.button(buttons, "클립보드에서 바로 복원", self.restore_from_clipboard)
        self.restore_clip_button.pack(side="left", padx=8)
        self.restore_file_button = self.button(buttons, "파일 복원", self.restore_from_file)
        self.restore_file_button.pack(side="left")

        head = ttk.Frame(tab)
        head.pack(fill="x", pady=(2, 4))
        ttk.Label(head, text="복원 결과", font=self.font_bold).pack(side="left")
        self.copy_restored_button = self.button(head, "복사", self.copy_restored)
        self.copy_restored_button.pack(side="right")
        output_frame, self.output_text = self.make_text(tab, 8, readonly=True)
        output_frame.pack(fill="both", expand=True)
        self.set_text(self.output_text, "")

        ttk.Label(tab, text="경고 목록 (복원 못 한 코드)", font=self.font_bold).pack(anchor="w", pady=(10, 4))
        warn_frame, self.warn_table = self.make_table(tab, [("코드", 160), ("이유", 640)], 4)
        warn_frame.pack(fill="x")
        self.warn_table.insert("", "end", values=("(예시) 학생099", "사전에 없는 번호입니다 (AI가 만든 코드일 수 있음)"))

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
        traceback.print_exception(exc_type, exc, tb)
        self._busy_end()
        messagebox.showerror("오류", f"{exc_type.__name__}: {exc}")

    def _show_error(self, exc):
        traceback.print_exc()
        self._busy_end()
        messagebox.showerror("오류", f"{type(exc).__name__}: {exc}")

    def _busy_start(self, message):
        self.root.configure(cursor="watch")
        self.check_label.configure(text=message, fg=COLORS["muted"])
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
    def fill_table(tree, rows):
        tree.delete(*tree.get_children())
        for row in rows:
            tree.insert("", "end", values=row)

    # ------------------------------------------------------------ 맨 위 줄: 사전
    def refresh_dictionary(self):
        """사전을 읽어서 학생 수와 수정 날짜를 보여 주고, 없으면 버튼을 잠근다."""
        try:
            info = workflows.dictionary_info()
            problem = NO_DICT_MESSAGE if info is None else None
        except Exception as exc:   # noqa: BLE001 - 사전이 깨져 있어도 프로그램은 떠야 한다
            info, problem = None, f"학생사전을 읽을 수 없음 - {exc}"
        if info is None:
            self.dict_label.configure(text=problem, foreground=COLORS["bad"])
        else:
            count, modified = info
            self.dict_label.configure(text=f"📘 학생사전: {count}명 ({modified})", foreground="")
        for b in (self.run_button, self.restore_button, self.restore_clip_button, self.restore_file_button):
            b.state(["!disabled"] if info else ["disabled"])

    def build_dictionary(self):
        path = filedialog.askopenfilename(title="명렬표 엑셀 선택",
                                          filetypes=[("엑셀 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")])
        if not path:
            return
        if workflows.DICT_PATH.exists() and not messagebox.askyesno(
                APP_TITLE, "덮어쓰면 기존 코드 번호가 바뀔 수 있습니다.\n계속할까요?", icon="warning"):
            return
        self.root.configure(cursor="watch")
        self.root.update()
        try:
            result = workflows.build_dictionary(path)
        finally:
            self._busy_end()
        self.refresh_dictionary()
        self._mark_result_stale("사전을 새로 만들었습니다. [가명처리 실행]을 다시 눌러 주세요.")

        lines = [f"학생 {result.count}명으로 사전을 만들었습니다.",
                 f"찾은 열: {', '.join(result.found)}",
                 f"못 찾은 열: {', '.join(result.missing) if result.missing else '없음'}"]
        if result.warnings:
            lines += ["", f"경고 {len(result.warnings)}건:"]
            lines += [f"- {w}" for w in result.warnings[:25]]
            if len(result.warnings) > 25:
                lines.append(f"... 외 {len(result.warnings) - 25}건")
            messagebox.showwarning(APP_TITLE, "\n".join(lines))
        else:
            messagebox.showinfo(APP_TITLE, "\n".join(lines))

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
        # '엑셀 원본 형식 유지'는 .xlsx 일 때만 고를 수 있다
        if ext == ".xlsx":
            self.radio_xlsx.state(["!disabled"])
        else:
            self.mode_var.set("md")
            self.radio_xlsx.state(["disabled"])
        self._mark_result_stale("새 파일을 골랐습니다. [가명처리 실행]을 눌러 주세요.")

    def _mark_result_stale(self, message):
        """이전 결과가 새 파일/새 사전과 맞지 않으니, 복사 버튼을 잠그고 안내만 보여 준다."""
        if self.result is None and self.copy_result_button.instate(["disabled"]):
            return   # 아직 결과를 만든 적이 없으면 처음 화면(예시)을 그대로 둔다
        self.result = None
        self.copy_result_button.state(["disabled"])
        self.check_label.configure(text=message, fg=COLORS["muted"])

    def run_mask(self):
        if self.selected_path is None:
            messagebox.showinfo(APP_TITLE, "먼저 가명처리할 파일을 선택하세요.")
            return
        dictionary = workflows.get_dictionary()
        self.copy_result_button.state(["disabled"])
        self._busy_start("처리 중입니다...")
        try:
            result = workflows.mask_file(self.selected_path, self.mode_var.get(), dictionary)
        except Exception:
            self.result = None
            self.check_label.configure(text="⚠ 처리하지 못했습니다 - 오류 내용은 알림창을 확인하세요", fg=COLORS["bad"])
            raise
        finally:
            self._busy_end()
        self.show_mask_result(result)

    def show_mask_result(self, result):
        self.result = result
        color = COLORS["ok"] if result.clean and result.ok else COLORS["bad"]
        self.check_label.configure(text=result.headline, fg=color)

        rows = [r for r in result.review]
        rows += [(f"남은 개인정보({kind})", where, value) for where, kind, value in result.findings]
        self.fill_table(self.review_table, rows)
        self.set_text(self.preview, result.text if result.ok else "", highlight=True)
        # 남은 개인정보가 있으면 복사를 막아서, 위험한 결과를 AI에 붙여 넣지 못하게 한다
        self.copy_result_button.state(["!disabled"] if result.can_copy else ["disabled"])

    def copy_result(self):
        if self.result is None or not self.result.can_copy:
            messagebox.showwarning(APP_TITLE, "복사할 수 있는 결과가 없습니다.")
            return
        self.copy_to_clipboard(self.result.text, self.copy_result_button)

    # ------------------------------------------------------------ 탭 2: 되돌리기
    def _show_restored(self, restored, warnings):
        self.set_text(self.output_text, restored)
        self.fill_table(self.warn_table, [split_warning(w) for w in warnings])

    def restore_from_box(self):
        text = self.input_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showinfo(APP_TITLE, "AI 답변을 먼저 붙여 넣으세요.")
            return
        restored, _, warnings = unmask_text(text, workflows.get_dictionary())
        self._show_restored(restored, warnings)

    def restore_from_clipboard(self):
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            text = ""
        if not text.strip():
            messagebox.showinfo(APP_TITLE, "클립보드에 글이 없습니다.")
            return
        restored, _, warnings = unmask_text(text, workflows.get_dictionary())
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
            result = workflows.restore_file(path, workflows.get_dictionary())
        finally:
            self._busy_end()
        saved = result.out_path.relative_to(workflows.BASE_DIR)
        shown = result.text if result.text is not None else f"엑셀 파일을 복원했습니다.\n저장 위치: {saved}"
        self._show_restored(shown, result.warnings)
        messagebox.showinfo(APP_TITLE, f"복원했습니다 ({result.changes}건, 경고 {len(result.warnings)}건)\n"
                                       f"저장 위치: {saved}")

    def copy_restored(self):
        text = self.output_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showinfo(APP_TITLE, "복사할 복원 결과가 없습니다.")
            return
        self.copy_to_clipboard(text, self.copy_restored_button)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
