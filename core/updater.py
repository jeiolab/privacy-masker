"""업데이트 확인. version.json 을 GET 으로만 읽고, 로컬 VERSION 과 비교한다.

보안:
- GET 한 번으로 version.json 만 읽는다. 쿼리·본문·사전·파일 이름·사용 기록·컴퓨터 정보는 보내지 않는다.
- 파일을 내려받거나 실행하지 않는다. 다운로드는 브라우저로 페이지만 연다.
- 주소가 비어 있거나, 인터넷이 없거나, 3초 안에 못 읽으면 None 을 돌려주고 오류를 내지 않는다.
"""
import json
import re
import urllib.request
import webbrowser

from core import config
from core.version import VERSION

def update_how_to(download=""):
    """알림창·받는 페이지에 보여 줄 안내. download 주소가 있으면 1번에 같이 적는다."""
    url = (download or getattr(config, "DOWNLOAD_URL", "") or "").strip()
    step1 = "1. 개인정보가리개.zip 을 받으세요"
    if url:
        step1 += f"\n   {url}"
    return (
        "이렇게 바꾸세요.\n"
        f"{step1}\n"
        "2. 압축을 풀어 지금 쓰는 프로그램 폴더에 덮어쓰세요\n"
        "3. 개인정보가리개.exe 를 다시 실행하세요\n"
        "\n"
        "문서 폴더의 개인정보가리개_데이터 는 지우지 마세요. 명단과 결과가 들어 있어요."
    )


# 예전 이름 유지 (안내 문구만 필요할 때)
UPDATE_HOW_TO = update_how_to()


def write_version_json(path=None, notes=""):
    """core/version.py · DOWNLOAD_URL 을 version.json 에 맞춰 쓴다. (빌드 때 자동)"""
    from datetime import date
    from pathlib import Path
    dest = Path(path) if path else Path(__file__).resolve().parent.parent / "version.json"
    data = {
        "version": VERSION,
        "date": date.today().isoformat(),
        "notes": notes or "새 버전",
        "download": (getattr(config, "DOWNLOAD_URL", "") or "").strip(),
    }
    dest.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return dest, data


def _version_tuple(text):
    nums = [int(x) for x in re.findall(r"\d+", str(text or ""))]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums[:3])


def is_newer(remote, local=VERSION):
    return _version_tuple(remote) > _version_tuple(local)


def fetch_update_info(url=None, timeout=None):
    """원격 version.json 을 읽는다. 실패하면 None. 식별 정보는 요청에 넣지 않는다."""
    url = (url if url is not None else config.UPDATE_URL) or ""
    url = url.strip()
    if not url.startswith(("https://", "http://")):
        return None
    seconds = config.UPDATE_TIMEOUT if timeout is None else timeout
    try:
        req = urllib.request.Request(
            url,
            method="GET",
            headers={
                "Accept": "application/json",
                "User-Agent": "개인정보가리개-update-check",  # 고정 문구. OS·버전·경로 없음
            },
        )
        with urllib.request.urlopen(req, timeout=seconds) as resp:
            raw = resp.read(4096)   # version.json 은 짧다. 큰 파일은 받지 않는다
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict) or not data.get("version"):
            return None
        download = str(data.get("download", "")).strip() or (getattr(config, "DOWNLOAD_URL", "") or "").strip()
        return {
            "version": str(data.get("version", "")).strip(),
            "date": str(data.get("date", "")).strip(),
            "notes": str(data.get("notes", "")).strip(),
            "download": download,
        }
    except Exception:
        return None


def open_download_page(url):
    """브라우저로 페이지만 연다. http(s) 가 아니면 열지 않는다. 파일 실행 없음."""
    url = (url or "").strip()
    if not url.startswith(("https://", "http://")):
        return False
    webbrowser.open(url)
    return True
