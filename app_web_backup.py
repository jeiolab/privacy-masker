#!/usr/bin/env python3
"""개인정보 가리개 - pywebview 화면 (화면 모양만, 기능은 아직 연결되지 않음).

실행: python3 app.py
- 이 파일은 창을 여는 일만 한다. 화면은 ui/index.html, ui/style.css, ui/app.js.
- 모든 버튼은 누르면 "아직 연결 안 됨" 알림창만 뜬다.
- 이전 tkinter 화면은 app_tk_backup.py 에 그대로 있다.
"""
from pathlib import Path

import webview

UI_DIR = Path(__file__).resolve().parent / "ui"


def main():
    webview.create_window(
        "개인정보 가리개",
        str(UI_DIR / "index.html"),
        width=1180,
        height=860,
        min_size=(420, 700),
        resizable=True,
        text_select=True,   # 미리보기·글상자의 글자를 드래그해서 선택할 수 있게
    )
    webview.start()


if __name__ == "__main__":
    main()
