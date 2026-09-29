# -*- coding: utf-8 -*-
"""4개 파일 교차 점검(앞뒤가 맞는지). 사용: python3 crosscheck.py <key> <xlsx>"""
import sys, re, openpyxl
sys.path.insert(0, ".")
import verify

TEACHERS = ["박윤아", "강희강", "임충만", "김규리"]
OTHER_STU = {"park": ["성O아", "장O준", "중2 여학생", "중1", "중3"], "kang": ["이O오", "장O준", "중2 여학생", "고2", "중3"],
             "lim": ["이O오", "성O아", "중2 여학생", "고2", "중1"], "kim": ["이O오", "성O아", "장O준", "고2", "중1", "중3"]}
NAMES = {"park": "박윤아", "kang": "강희강", "lim": "임충만", "kim": "김규리"}
CODES = {"park": ["○", "△", "이"], "kang": ["○", "요", "T"]}


def all_text(wb):
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    yield ws.title, c.coordinate, c.value


def run(key, path):
    wb = openpyxl.load_workbook(path)
    issues = []
    names = wb.sheetnames
    # 1) 다른 선생님 이름 / 다른 학생 잔존
    for t, co, v in all_text(wb):
        for tn in TEACHERS:
            if tn != NAMES[key] and tn in v:
                issues.append(("다른 선생님 이름", t, co, v[:60]))
        for so in OTHER_STU[key]:
            if so in v and not v.startswith("="):
                issues.append(("다른 학생 표기", t, co, v[:60]))
    # 2) 본문이 가리키는 탭 이름이 실제 존재하는가
    for t, co, v in all_text(wb):
        for m in re.findall(r"[‘'「]?((?:[0-5]\.|양식\d_|인쇄_)[가-힣A-Za-z0-9_·\.\(\) ]*?)[’'」]", v):
            pass
    tab_names = set(names)
    for t, co, v in all_text(wb):
        for m in re.finditer(r"(양식\d_[가-힣A-Za-z0-9_\(\)]+|인쇄_[가-힣A-Za-z0-9_\.]+|5\.학생 점검 카드|[0-4]\.(?:진행 지도|정의|기록지|요약|중재 설계\(초안\)))", v):
            nm = m.group(1)
            if nm not in tab_names and not v.startswith("="):
                issues.append(("존재하지 않는 탭 언급", t, co, f"{nm} ← {v[:50]}"))
    # 3) DV 목록 ↔ 1.정의 코드표
    d = wb["1.정의"]
    codes_in_def = set()
    for row in d.iter_rows(min_col=1, max_col=1):
        for c in row:
            if isinstance(c.value, str):
                codes_in_def.add(c.value.strip())
    rec = wb["2.기록지"]
    for dv in rec.data_validations.dataValidation:
        if dv.type == "list" and dv.formula1:
            for item in dv.formula1.strip('"').split(","):
                ok = item in codes_in_def or any(item in x for x in codes_in_def)
                if not ok and item not in ("기초선", "중재", "유지") and not re.match(r"\d교시", item):
                    issues.append(("드롭다운 값이 1.정의 코드표에 없음", "2.기록지", str(dv.sqref), item))
    # 4) 양식6 목표 == 양식7 LTO 링크
    b15 = wb["양식6_행동중재계획"]["B15"].value
    if b15 != "='양식7_데이터시트'!$E$5":
        issues.append(("양식6 목표가 양식7 LTO와 연결되지 않음", "양식6", "B15", str(b15)[:50]))
    # 5) 색 규칙
    for t, co, why in verify.color_scan(path):
        issues.append(("금지 색", t, co, why))
    # 6) '회기' 사용처(슈퍼비전 회기만 허용) — 목록으로 출력
    hoe = [(t, co, v[:70]) for t, co, v in all_text(wb) if "회기" in v and not v.startswith("=")]
    return issues, hoe


if __name__ == "__main__":
    issues, hoe = run(sys.argv[1], sys.argv[2])
    print(f"[{sys.argv[1]}] 문제 {len(issues)}건")
    for i in issues[:40]:
        print("  ", i)
    print("  '회기' 사용처:", len(hoe))
    for h in hoe:
        print("    ", h)
