#!/usr/bin/env python3
"""배포용(exe) 진입점.

- 파이썬으로 쓸 때는 지금처럼 `python app.py` 를 바로 실행하면 된다. (이 파일은 안 써도 됨)
- PyInstaller 로 만든 exe 로 켜지면(검은 콘솔 창이 없음) 시작 도중 생긴 오류가 그냥 사라지므로,
  문서\\개인정보가리개_데이터\\logs\\실행오류.txt 에 오류를 남기고 화면(app.App)을 띄운다.
- data/·output/·logs/ 위치는 core/config.py 한 곳에서 정한다 (파이썬·exe 모두 문서 아래).
"""
import sys


def _setup_logs():
    """exe 로 실행될 때만: 문서 아래 logs/ 를 만들고 오류 출력을 그쪽 파일로 보낸다."""
    if not getattr(sys, "frozen", False):
        return
    try:
        from core import config
        from core.error_report import LOG_NAME, MaskingLogWriter
        logs = config.logs_dir()
        logs.mkdir(parents=True, exist_ok=True)
        stream = MaskingLogWriter(logs / LOG_NAME)
        sys.stderr = stream
        sys.stdout = stream
    except OSError:
        pass   # 로그를 못 남겨도 프로그램은 그대로 뜬다


def main():
    _setup_logs()
    from app import App
    App().run()


if __name__ == "__main__":
    main()
