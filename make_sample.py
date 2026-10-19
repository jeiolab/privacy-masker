"""테스트용 가짜 명렬표 엑셀을 만든다.

1학년 1반~3반, 반마다 10명. 이름·전화·주소는 가짜이며,
동명이인, 일반 단어와 겹치는 이름, 외자 이름을 일부러 넣는다.
"""

import calendar
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "samples" / "가짜_명렬표.xlsx"
RANDOM_SEED = 42

COLUMNS = [
    "번호",
    "학번",
    "성명",
    "성별",
    "생년월일",
    "학생연락처",
    "보호자명",
    "관계",
    "보호자연락처",
    "주소",
]

SURNAMES = [
    "김", "이", "박", "최", "정", "강", "조", "윤", "장", "임",
    "한", "오", "서", "신", "권", "황", "안", "송", "전", "홍",
]

GIVEN_NAMES = {
    "남": [
        "민준", "서준", "도윤", "예준", "시우", "주원", "하준", "지호",
        "준서", "건우", "현우", "지훈", "우진", "선우", "연우", "유준",
        "은우", "민재", "승현", "태윤",
    ],
    "여": [
        "서연", "서윤", "지우", "하은", "민서", "지민", "채원", "수아",
        "지아", "다은", "예은", "수빈", "윤서", "소율", "아린", "예린",
        "하린", "지윤", "유진", "채은",
    ],
}

REGIONS = [
    ("서울특별시", ["강남구", "서초구", "마포구", "송파구", "노원구", "은평구"]),
    ("부산광역시", ["해운대구", "수영구", "부산진구", "동래구"]),
    ("대구광역시", ["수성구", "달서구"]),
    ("인천광역시", ["연수구", "남동구", "부평구"]),
    ("경기도", ["성남시 분당구", "고양시 일산서구", "수원시 팔달구", "용인시 기흥구", "화성시"]),
    ("강원특별자치도", ["춘천시", "원주시"]),
    ("충청북도", ["청주시 흥덕구", "충주시"]),
    ("충청남도", ["천안시 동남구", "아산시"]),
    ("전북특별자치도", ["전주시 완산구", "군산시"]),
    ("전라남도", ["여수시", "순천시"]),
    ("경상북도", ["포항시 북구", "경주시"]),
    ("경상남도", ["창원시 성산구", "김해시"]),
    ("제주특별자치도", ["제주시", "서귀포시"]),
]

ROADS = [
    "중앙로", "학교로", "행복길", "벚꽃로", "문화로", "백제로",
    "한강대로", "대학로", "솔밭로", "무지개로", "늘봄로", "푸른길",
]

# (반, 번호)에 고정할 까다로운 사례. 나머지는 시드로 무작위 생성한다.
SPECIAL_CASES = {
    (1, 3): {"성명": "이나라", "성별": "여"},
    (1, 7): {"성명": "정하윤", "성별": "여"},
    (2, 4): {"성명": "김하늘", "성별": "여"},
    (3, 1): {"성명": "박준", "성별": "남"},
    (3, 6): {"성명": "정하윤", "성별": "여"},
}

RESERVED_NAMES = {case["성명"] for case in SPECIAL_CASES.values()}


def student_id(class_no, number):
    return f"1{class_no:02d}{number:02d}"


def make_name(rng, gender, used_names):
    for _ in range(1000):
        name = rng.choice(SURNAMES) + rng.choice(GIVEN_NAMES[gender])
        if name not in used_names and name not in RESERVED_NAMES:
            used_names.add(name)
            return name
    raise RuntimeError("고유한 가짜 이름을 더 만들 수 없습니다.")


def make_phone(rng, used_phones):
    for _ in range(1000):
        phone = f"010-{rng.randint(1000, 9999)}-{rng.randint(1000, 9999)}"
        if phone not in used_phones:
            used_phones.add(phone)
            return phone
    raise RuntimeError("고유한 가짜 전화번호를 더 만들 수 없습니다.")


def make_birth(rng):
    month = rng.randint(1, 12)
    day = rng.randint(1, calendar.monthrange(2019, month)[1])
    return f"2019-{month:02d}-{day:02d}"


def make_address(rng):
    sido, districts = rng.choice(REGIONS)
    district = rng.choice(districts)
    road = rng.choice(ROADS)
    number = rng.randint(1, 180)
    if rng.random() < 0.7:
        return f"{sido} {district} {road} {number} {rng.randint(1, 15)}동 {rng.randint(101, 1504)}호"
    return f"{sido} {district} {road} {number}"


def make_guardian(rng, child_name, used_names, used_phones):
    if rng.random() < 0.08:
        gender = rng.choice(["남", "여"])
        relation = "조부" if gender == "남" else "조모"
    else:
        gender = rng.choices(["남", "여"], weights=[4, 6])[0]
        relation = "부" if gender == "남" else "모"
    name = make_name(rng, gender, used_names)
    while name == child_name:
        used_names.discard(name)
        name = make_name(rng, gender, used_names)
    return name, relation, make_phone(rng, used_phones)


def build_rows(rng):
    used_names = set(RESERVED_NAMES)
    used_phones = set()
    rows = []

    for class_no in range(1, 4):
        for number in range(1, 11):
            special = SPECIAL_CASES.get((class_no, number))
            if special:
                name = special["성명"]
                gender = special["성별"]
            else:
                gender = rng.choice(["남", "여"])
                name = make_name(rng, gender, used_names)

            guardian_name, relation, guardian_phone = make_guardian(
                rng, name, used_names, used_phones
            )
            rows.append(
                {
                    "번호": number,
                    "학번": student_id(class_no, number),
                    "성명": name,
                    "성별": gender,
                    "생년월일": make_birth(rng),
                    "학생연락처": make_phone(rng, used_phones),
                    "보호자명": guardian_name,
                    "관계": relation,
                    "보호자연락처": guardian_phone,
                    "주소": make_address(rng),
                }
            )
    return rows


def validate(rows):
    if len(rows) != 30:
        raise RuntimeError(f"학생 수가 30명이 아닙니다: {len(rows)}")

    ids = [row["학번"] for row in rows]
    expected_ids = [
        student_id(class_no, number)
        for class_no in range(1, 4)
        for number in range(1, 11)
    ]
    if ids != expected_ids:
        raise RuntimeError("학번이 1학년 1반~3반, 번호 01~10 형식과 맞지 않습니다.")

    names = [row["성명"] for row in rows]
    duplicated = sorted({name for name in names if names.count(name) > 1})
    if duplicated != ["정하윤"]:
        raise RuntimeError(f"동명이인은 정하윤 한 쌍만 있어야 합니다: {duplicated}")

    same_name_classes = {
        row["학번"][1:3] for row in rows if row["성명"] == "정하윤"
    }
    if len(same_name_classes) != 2:
        raise RuntimeError("동명이인은 서로 다른 반에 있어야 합니다.")

    by_name = {row["성명"]: row for row in rows}
    for name in ("김하늘", "이나라", "박준"):
        if name not in by_name:
            raise RuntimeError(f"필수 이름이 없습니다: {name}")
    if by_name["박준"]["성명"].removeprefix("박") != "준":
        raise RuntimeError("박준은 외자 이름이어야 합니다.")


def write_xlsx(rows, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "명렬표"
    sheet.append(COLUMNS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append([row[column] for column in COLUMNS])
    for column in sheet.columns:
        letter = column[0].column_letter
        width = max(len(str(cell.value or "")) for cell in column) + 2
        sheet.column_dimensions[letter].width = min(width, 42)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)


def main():
    rng = random.Random(RANDOM_SEED)
    rows = build_rows(rng)
    validate(rows)
    write_xlsx(rows, OUTPUT)
    print(f"저장: {OUTPUT}")
    print(f"학생 수: {len(rows)}")


if __name__ == "__main__":
    main()
