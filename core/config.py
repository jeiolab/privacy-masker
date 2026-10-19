"""설정값 모음 (열 이름 후보, 성씨, 일반 단어, 데이터 폴더 경로 등).

중요 규칙 (어떤 업데이트에서도 지키세요):
- 기존 학생 코드 번호(학생001, S001, 보호자001 등)는 절대 바뀌면 안 된다.
- 교직원 코드 번호(교사001 …)도 한 번 붙으면 바꾸지 않는다.
- 사전 구조가 바뀌어도 예전 사전은 읽을 수 있어야 한다 (사전형식버전 참고).
"""
import os
import shutil
import sys
from pathlib import Path

# 사전 파일 안의 형식 번호. 구조가 바뀌면 숫자를 올리고, 예전 번호도 계속 읽는다.
DICT_FORMAT_VERSION = 1

# 업데이트 확인: 이 주소의 version.json 만 GET 으로 읽는다.
# (사전·파일 이름·사용 기록·컴퓨터 정보는 보내지 않는다. 파일 자동 다운로드/실행도 하지 않는다.)
# 나중에 GitHub raw 주소로 바꾸면 된다. 비어 있으면 확인을 건너뛴다.
UPDATE_URL = "https://raw.githubusercontent.com/jeiolab/privacy-masker/main/version.json"
# 선생님이 zip 을 받는 페이지. 알림창 안내와 version.json 의 download 에 그대로 들어간다.
DOWNLOAD_URL = "https://github.com/jeiolab/privacy-masker/releases/latest"
UPDATE_TIMEOUT = 3   # 초. 이 안에 못 읽으면 조용히 포기

# ---------------------------------------------------------------- 쓰는 폴더 (이 파일에서만 정한다)
# 문서\개인정보가리개_데이터\ 아래에 data / output / logs 를 둔다.
# 파이썬으로 실행하든 exe 로 실행하든 같은 위치를 쓴다.
USER_DATA_DIR_NAME = "개인정보가리개_데이터"
DICT_FILENAME = "학생사전.xlsx"
IGNORE_FILENAME = "무시단어.txt"
OUTPUT_PATH = "data/학생사전.xlsx"          # user_data_dir() 기준 상대 경로 (예전 코드 호환)
IGNORE_WORDS_PATH = "data/무시단어.txt"    # user_data_dir() 기준 상대 경로


def program_dir():
    """프로그램이 있는 폴더 (exe 옆 / 프로젝트 폴더). 예전 data·output 을 찾을 때만 쓴다."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def documents_dir():
    """사용자 '문서' 폴더."""
    if sys.platform.startswith("win"):
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(260)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf) == 0 and buf.value:
                p = Path(buf.value)
                if p.is_dir():
                    return p
        except Exception:
            pass
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    for name in ("Documents", "문서"):
        p = home / name
        if p.is_dir():
            return p
    return home / "Documents"


def user_data_dir():
    """문서\\개인정보가리개_데이터"""
    return documents_dir() / USER_DATA_DIR_NAME


def data_dir():
    return user_data_dir() / "data"


def output_dir():
    return user_data_dir() / "output"


def logs_dir():
    return user_data_dir() / "logs"


def dict_path():
    return data_dir() / DICT_FILENAME


def ignore_words_path():
    return data_dir() / IGNORE_FILENAME


def backup_dir():
    return data_dir() / "백업"


def ai_dir():
    return output_dir() / "AI용"


def keep_dir():
    return output_dir() / "보관용"


def restore_dir():
    return output_dir() / "복원"


def app_base_dir():
    """쓰는 폴더의 기준 = 문서\\개인정보가리개_데이터. (파이썬·exe 모두 동일)"""
    return user_data_dir()


def ensure_user_data_dirs():
    """data / output / logs 와 결과 하위 폴더를 만든다."""
    for p in (data_dir(), ai_dir(), keep_dir(), restore_dir(), logs_dir(), backup_dir()):
        p.mkdir(parents=True, exist_ok=True)


def _has_any(folder):
    if not folder.exists() or not folder.is_dir():
        return False
    try:
        next(folder.iterdir())
        return True
    except StopIteration:
        return False


def _relocate(src, dst, skip_names=()):
    """src 내용을 dst 로 옮긴다. dst 에 이미 있는 파일은 덮어쓰지 않는다. (옮긴 게 있으면 True)"""
    if not src.exists() or src.resolve() == dst.resolve():
        return False
    dst.mkdir(parents=True, exist_ok=True)
    moved = False
    for item in list(src.iterdir()):
        if item.name in skip_names:
            continue
        dest = dst / item.name
        if dest.exists():
            if item.is_dir():
                moved = _relocate(item, dest) or moved
            continue
        shutil.move(str(item), str(dest))
        moved = True
    try:
        if src.exists() and not any(src.iterdir()):
            src.rmdir()
    except OSError:
        pass
    return moved


def migrate_legacy_data():
    """예전 위치(프로그램 폴더의 data/output/logs)에 파일이 있으면 문서 아래로 옮긴다.
    새 위치에 이미 사전이 있으면 그 사전은 덮어쓰지 않는다.
    알림 문구를 돌려주고, 옮길 것이 없으면 None.
    """
    old_root = program_dir()
    pairs = (
        (old_root / "data", data_dir()),
        (old_root / "output", output_dir()),
        (old_root / "logs", logs_dir()),
    )
    if not any(_has_any(src) for src, _ in pairs):
        return None

    ensure_user_data_dirs()
    already = dict_path().exists()
    skip = {DICT_FILENAME} if already else set()
    moved = False
    for src, dst in pairs:
        extra = skip if src.name == "data" else set()
        moved = _relocate(src, dst, skip_names=extra) or moved

    if already:
        return "문서\\개인정보가리개_데이터 에 이미 사전이 있어 예전 사전을 덮어쓰지 않았어요"
    if moved:
        return "데이터 폴더를 문서\\개인정보가리개_데이터 로 옮겼어요"
    return None


# 필드별 열 이름 후보 (비교할 때 공백은 무시합니다)
COLUMN_CANDIDATES = {
    "성명": ["성명", "이름", "학생명", "학생이름"],
    "학번": ["학번"],
    "보호자명": ["보호자명", "보호자", "학부모명", "부모명"],
    "학생연락처": ["학생연락처", "연락처", "휴대폰", "전화번호"],
    "보호자연락처": ["보호자연락처", "학부모연락처", "보호자 휴대폰"],
    "주소": ["주소"],
    "생년월일": ["생년월일", "생일"],
}

# 열 이름에 이 단어가 들어간 연락처는 보호자연락처로 구분
GUARDIAN_KEYWORDS = ["보호자", "학부모", "부모"]

# 머리글 줄 찾기: 위에서 N줄 안에서 후보 단어가 M개 이상 든 줄
HEADER_SCAN_ROWS = 10
HEADER_MIN_MATCHES = 2

# 두 글자 성
COMPOUND_SURNAMES = ["남궁", "제갈", "선우", "황보", "독고", "사공", "서문"]

# 일반 단어와 겹치는 이름 (성 뺀 이름 기준)
COMMON_WORD_NAMES = [
    "하늘", "나라", "바다", "한결", "사랑", "우리", "하나", "보람",
    "슬기", "다솜", "소망", "기쁨", "나래", "가람", "별",
    "봄", "여름", "가을", "겨울", "새벽", "노을", "이슬", "구름", "단비", "한별", "하람",
    "아름", "미소", "온유", "다움", "겨레", "누리", "가온", "마루", "소리", "이음", "다온",
    "하랑", "시원", "나리", "보미",
    "지혜", "은혜", "믿음", "평화", "자유", "승리", "정의", "행복", "희망", "선물", "성실",
]

# ---------------------------------------------------------------- output 폴더 자동 정리
# 프로그램을 켤 때, 아래 일수가 지난 결과 파일을 자동으로 지운다. (data/ 사전·백업은 절대 건드리지 않음)
# 폴더 이름(output/ 아래) → 보관 일수
CLEANUP_DAYS = {
    "AI용": 7,     # AI에 올리고 나면 필요 없으니 짧게 둔다
    "보관용": 30,
    "복원": 30,
}

# ---------------------------------------------------------------- 이름 의심 탐지
# 가명처리 뒤에도 남은 '사전에 없지만 사람 이름처럼 보이는 단어'를 찾을 때 쓰는 주요 한국 성씨
KOREAN_SURNAMES = [
    "김", "이", "박", "최", "정", "강", "조", "윤", "장", "임",
    "한", "오", "서", "신", "권", "황", "안", "송", "전", "홍",
    "유", "고", "문", "양", "손", "배", "백", "허", "남", "심",
    "노", "하", "곽", "성", "차", "주", "우", "구", "민", "류",
    "나", "진", "지", "엄", "채", "원", "천", "방", "공", "현",
    "함", "변", "염", "여", "추", "도", "소", "석", "선", "설",
    "마", "길", "연", "위", "표", "명", "기", "반", "라", "왕",
    "금", "옥", "육", "인", "맹", "제", "모", "탁", "국", "어",
]

# 이름(성씨로 시작하는 2~4글자) 뒤에 이런 말이 붙으면 사람으로 의심한다
NAME_SUSPECT_SUFFIXES = [
    "선생님", "어머니", "아버지", "보호자", "학생", "쌤",
    "이가", "이는", "에게", "씨", "군", "양", "와", "과",
]

# 성씨 한 글자 뒤에 '붙은 말'이 아래 조사·시간 표현뿐이면 사람 이름이 아니라고 보고 의심에서 제외한다.
# 예: "다음 주까지"(주+까지), "이번 주에"(주+에), "한 주 동안"(주+동안) → 경고하지 않음
NAME_SUSPECT_PARTICLES = [
    "까지", "부터", "에", "에서", "에는", "에도", "으로", "로",
    "마다", "동안", "간", "말", "초", "중", "내", "쯤", "경", "께",
]

# 성씨로 시작하지만 사람 이름이 아닌 흔한 단어 (의심에서 제외)
NAME_SUSPECT_EXCEPTIONS = [
    "이번", "정말", "조금", "장난", "조차", "조건", "정도", "정보", "정리", "정식",
    "강조", "강의", "강력", "강하", "주의", "주소", "주장", "오늘", "이것", "이거",
    "저것", "그것", "서로", "신고", "신청", "신분", "신기", "양쪽", "양측", "양해",
    "문제", "문의", "유리", "고장", "구조", "구분", "구별", "조절", "조정", "방법",
    "방금", "임시", "전화", "전체", "전부", "최고", "최근", "최대", "최소", "민감",
    "안내", "안전", "남녀", "하나", "소문", "배려", "백지", "고려", "노력", "성적",
    "성실", "차이", "주말", "임무", "한번", "한참", "한편", "현재", "현실", "모두",
    "제목", "제출", "제일", "지금", "지난", "도움", "도착", "방향", "최종", "정확",
    "남학생", "여학생", "전학생", "방과후", "진학", "신입",
    "국어", "한문", "기술", "음악", "진로", "상담", "사서", "강사", "원어민", "교장", "교감",
    "이제", "이후", "이유", "조퇴", "조회", "한국", "유지", "오후", "공부", "방학",
    "지각", "추천", "하교", "함께",
]
