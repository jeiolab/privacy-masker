# -*- coding: utf-8 -*-
"""
pyhwpx(아래아한글 자동화)로 테스트용 '가짜' 상담기록 한글 문서를 생성하는 스크립트.

생성물:
  - samples/가짜_상담기록.hwp
  - samples/가짜_상담기록.hwpx   (동일 내용)

학생 정보 출처:
  - samples/가짜_명렬표.xlsx  (1학년 1반 = 상위 10명 사용)

문서 구성:
  - 머리말: "1학년 1반 ○○○ 담임"(가짜 교사 이름) + 학생 이름 하나
  - 제목: "2학기 학생 상담 기록"
  - 표: 학번 / 성명 / 보호자명 / 보호자연락처 (10명)
  - 표 아래 상담 문단 5개: 성명 전체 / 성 뺀 이름+조사 / 보호자명 / 전화번호 / 학번을 섞어서 서술
  - 문단 하나에 한글 메모(주석 메모)를 달고 그 안에 학생 이름 포함
  - 문서 정보(요약)의 제목·지은이에 학생 이름 포함

주의:
  - 한글을 백그라운드(화면에 안 보이게)로 실행하고, 끝나면 종료한다.
  - 최초 실행 시 뜨는 '접근 허용' 보안창은 보안모듈 자동 등록으로 더 이상 뜨지 않게 처리한다.
  - 윈도우 + 아래아한글 설치 환경 전용.
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import openpyxl

try:
    import pandas as pd
except Exception:
    pd = None

from pyhwpx import Hwp

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE, "samples", "가짜_명렬표.xlsx")
HWP_OUT = os.path.join(BASE, "samples", "가짜_상담기록.hwp")
HWPX_OUT = os.path.join(BASE, "samples", "가짜_상담기록.hwpx")

# 학생 명단에 존재하지 않는 '가짜 교사' 이름 (머리말 담임용)
FAKE_TEACHER = "박지훈"


def load_students(n=10):
    """명렬표 xlsx에서 상위 n명(1학년 1반)의 학번/성명/보호자명/보호자연락처를 읽는다."""
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb.active
    # 헤더: 번호(0) 학번(1) 성명(2) 성별(3) 생년월일(4) 학생연락처(5) 보호자명(6) 관계(7) 보호자연락처(8) 주소(9)
    students = []
    for row in list(ws.iter_rows(min_row=2, values_only=True))[:n]:
        students.append(
            {
                "학번": str(row[1]).strip(),
                "성명": str(row[2]).strip(),
                "보호자명": str(row[6]).strip(),
                "보호자연락처": str(row[8]).strip(),
            }
        )
    return students


def has_final_consonant(ch):
    """한글 음절의 받침 유무."""
    code = ord(ch)
    if 0xAC00 <= code <= 0xD7A3:
        return (code - 0xAC00) % 28 != 0
    return False


def josa(word, with_final, without_final):
    """단어 끝 받침에 따라 조사를 붙인다. 예) josa('준서','은','는') -> '준서는'."""
    if not word:
        return word
    return word + (with_final if has_final_consonant(word[-1]) else without_final)


def given_name(full_name):
    """한국식 1글자 성을 제외한 이름 부분."""
    return full_name[1:] if len(full_name) >= 2 else full_name


def build_counseling_paragraphs(students):
    """상담 문단 5개를 생성한다. 각 문단에 성명 전체/이름+조사/보호자명/전화번호/학번을 섞는다."""
    paras = []
    s = students

    # 1
    gn = given_name(s[0]["성명"])
    paras.append(
        f"학번 {s[0]['학번']}번 {s[0]['성명']} 학생과 2학기 학업 및 교우관계에 대하여 상담하였다. "
        f"{josa(gn, '은', '는')} 진로에 대한 고민을 이야기하였고, "
        f"보호자 {s[0]['보호자명']}님({s[0]['보호자연락처']})께 상담 내용을 안내하였다."
    )
    # 2
    gn = given_name(s[1]["성명"])
    paras.append(
        f"{s[1]['성명']} 학생은 최근 학습 태도가 눈에 띄게 향상되었다. "
        f"{josa(gn, '이', '가')} 수학 성적 향상을 희망하여, "
        f"보호자 {s[1]['보호자명']}님께 {s[1]['보호자연락처']}로 연락하여 가정학습을 당부하였다. (학번 {s[1]['학번']})"
    )
    # 3
    gn = given_name(s[2]["성명"])
    paras.append(
        f"{s[2]['성명']}({s[2]['학번']}) 학생 상담을 진행하였다. "
        f"{josa(gn, '을', '를')} 격려하며 동아리 활동 참여를 권유하였고, "
        f"보호자 {s[2]['보호자명']}님({s[2]['보호자연락처']})과 추후 상담 일정을 협의하였다."
    )
    # 4
    gn = given_name(s[6]["성명"])
    paras.append(
        f"학번 {s[6]['학번']} {s[6]['성명']} 학생은 출결이 안정적이다. "
        f"{josa(gn, '은', '는')} 봉사활동 계획을 세웠으며, "
        f"보호자 {s[6]['보호자명']}님(연락처 {s[6]['보호자연락처']})께 진로 상담 결과를 전달하였다."
    )
    # 5 (이 문단에 메모를 단다)
    gn = given_name(s[8]["성명"])
    paras.append(
        f"{s[8]['성명']} 학생과 교우관계 상담을 하였다. "
        f"{josa(gn, '은', '는')} 학급 활동에 적극적으로 참여하였고, "
        f"보호자 {s[8]['보호자명']}님({s[8]['보호자연락처']}, 학번 {s[8]['학번']})께 상담 내용을 공유하였다."
    )
    return paras


def main():
    students = load_students(10)
    print(f"[1] 학생 {len(students)}명 로드 완료")

    # 백그라운드 실행 + 보안모듈 자동 등록(register_module=True 가 기본, 접근허용 창 방지)
    hwp = Hwp(new=True, visible=False, register_module=True)
    hwp.set_visible(False)
    print("[2] 한글 백그라운드 실행 및 보안모듈 등록")

    try:
        # ---------- 머리말 ----------
        act = hwp.HAction
        pset = hwp.HParameterSet.HHeaderFooter
        act.GetDefault("HeaderFooter", pset.HSet)
        act.Execute("HeaderFooter", pset.HSet)
        header_text = (
            f"1학년 1반 {FAKE_TEACHER} 담임" + " " * 10 + f"{students[0]['성명']}"
        )
        hwp.insert_text(header_text)
        hwp.CloseEx()  # 머리말 편집영역 빠져나오기
        print("[3] 머리말 삽입 완료")

        # ---------- 제목 ----------
        title = "2학기 학생 상담 기록"
        hwp.ParagraphShapeAlignCenter()
        hwp.set_font(Bold=True, Height=18)
        hwp.insert_text(title)
        hwp.BreakPara()
        # 본문용 서식 복원
        hwp.set_font(Bold=False, Height=10)
        hwp.ParagraphShapeAlignLeft()
        hwp.BreakPara()
        print("[4] 제목 삽입 완료")

        # ---------- 표 (학번/성명/보호자명/보호자연락처 x 10명) ----------
        cols = ["학번", "성명", "보호자명", "보호자연락처"]
        hwp.create_table(rows=len(students) + 1, cols=len(cols), treat_as_char=True, header=True)
        # 캐럿은 첫 셀에 위치. 헤더 행 입력.
        for ci, col in enumerate(cols):
            hwp.insert_text(col)
            hwp.TableRightCell()
        # 데이터 행 입력
        for si, stu in enumerate(students):
            for ci, col in enumerate(cols):
                hwp.insert_text(str(stu[col]))
                last_cell = (si == len(students) - 1) and (ci == len(cols) - 1)
                if not last_cell:
                    hwp.TableRightCell()
        # 표 밖(문서 끝)으로 이동
        hwp.MoveDocEnd()
        hwp.BreakPara()
        print("[5] 표 삽입 완료")

        # ---------- 상담 문단 5개 ----------
        paras = build_counseling_paragraphs(students)
        memo_student = students[8]["성명"]
        for idx, para in enumerate(paras):
            hwp.insert_text(para)
            if idx == len(paras) - 1:
                # 마지막 문단에 한글 메모(주석 메모) 달기 — 메모 내용에 학생 이름 포함
                hwp.insert_memo(
                    f"상담 메모: {memo_student} 학생 보호자 재상담 예정. 개인정보 취급 주의.",
                    memo_type="memo",
                )
            hwp.BreakPara()
        print(f"[6] 상담 문단 {len(paras)}개 + 메모 삽입 완료")

        # ---------- 문서 정보(요약): 제목 / 지은이에 학생 이름 ----------
        si = hwp.XHwpDocuments.Item(0).XHwpSummaryInfo
        si.Title = f"{students[0]['성명']} 학생 2학기 상담 기록"
        si.Author = students[6]["성명"]
        print("[7] 문서 요약(제목/지은이) 설정 완료")

        # ---------- 저장 (hwp + hwpx) ----------
        for path in (HWP_OUT, HWPX_OUT):
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass
        hwp.save_as(HWP_OUT, format="HWP")
        hwp.save_as(HWPX_OUT)  # 확장자 .hwpx -> 자동으로 HWPX 포맷 저장
        print(f"[8] 저장 완료:\n    - {HWP_OUT}\n    - {HWPX_OUT}")
    finally:
        # ---------- 한글 종료 ----------
        hwp.quit()
        print("[9] 한글 종료")


if __name__ == "__main__":
    main()
