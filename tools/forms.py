# -*- coding: utf-8 -*-
"""원본 xlsx에서 표준 양식 탭(양식1·2·3·5·6·7 등)을 가져와 정리하고, 새로 만든 탭과 합쳐 최종 통합본을 만든다."""
import re
from copy import copy
import openpyxl
from openpyxl.styles import Alignment
from lib import *
import sheets, records, summary

KEEP_FORMS = ["양식1_ABC기록지", "양식2_선호자극", "양식3_QABF", "양식5_FBA결과보고서", "양식6_행동중재계획",
              "양식7_데이터시트", "5.학생 점검 카드"]


def _clear(ws, rng):
    for row in ws[rng]:
        for c in row:
            c.value = None


def _set(ws, ref, value, color=None, bg="keep", italic=None, size=None, bold=None):
    c = ws[ref]
    # 병합 셀의 앵커가 아니면 건드리지 않는다
    c.value = value
    f = copy(c.font)
    kw = dict(name=f.name or FONT, size=size or f.size or 10, bold=f.bold if bold is None else bold,
              italic=f.italic if italic is None else italic, color=color if color else f.color)
    c.font = Font(**kw)
    if bg != "keep":
        c.fill = fill(bg) if bg else PatternFill(fill_type=None)


def patch_form1_3(wb, cfg):
    for name in ("양식1_ABC기록지", "양식2_선호자극", "양식3_QABF"):
        ws = wb[name]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    v = c.value
                    v2 = v.replace("5분", "일정 시간") if name.startswith("양식1") and "5분" in v else v
                    if v2 != v:
                        c.value = v2


def patch_form5(wb, cfg):
    ws = wb["양식5_FBA결과보고서"]
    F = cfg["fba"]
    ws["D3"].value = f'{cfg["teacher"]} 선생님'
    rows = [r for r in range(1, ws.max_row + 1) if ws[f"A{r}"].value and str(ws[f"A{r}"].value).strip() == "요약"]
    r = rows[-1]
    for col, key in zip("BDFHJ", ["E", "A", "B", "C", "F"]):
        c = ws[f"{col}{r}"]
        c.value = F[key]
        c.font = Font(name=FONT, size=9, color=EX_TXT, italic=False)
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    ws.row_dimensions[r].height = 118
    ws[f"A{r}"].value = "요약\n(초안)"
    sc = [x for x in range(1, ws.max_row + 1) if ws[f"A{x}"].value and str(ws[f"A{x}"].value).strip() == "점수"][-1]
    for col in "BDFHJ":
        ws[f"{col}{sc}"].value = None
    ws["A1"].value = "기능적행동평가(FBA) 결과 보고서    ※ 회색 글씨 = 슈퍼바이저 초안(예시). 4회기에 모은 자료로 확인·수정하세요."
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[1].height = 34
    return r


def _unify_terms(ws):
    """양식6·7의 '회기'(교수 회차)를 '회차'로 통일 — '회기'는 슈퍼비전 모임에만 쓴다."""
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and not c.value.startswith("=") and "회기" in c.value:
                c.value = c.value.replace("회기", "회차")


def patch_form6(wb, cfg):
    ws = wb["양식6_행동중재계획"]
    _unify_terms(ws)
    B = cfg["bip"]
    ws["A1"].value = "행동중재계획    ※ 회색 글씨 = 슈퍼바이저 초안(예시). 4회기 이후 함께 확인·수정하세요."
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[1].height = 34
    ws["D6"].value = f'{cfg["teacher"]} 선생님'
    ws["B7"].value = "중재1"
    ws["D7"].value = None
    ws["D7"].fill = fill(YELLOW)
    ws["C7"].value = "적용기간(시작일~)"
    ws["C8"].value = "약물복용 여부"
    ws["B9"].value = "='1.정의'!$B$8&\" / 증가: \"&'1.정의'!$B$15&\" — \"&'1.정의'!$B$16"
    ws["C11"].value = B["rank1"]
    ws["C12"].value = B["rank2"]
    ws["B13"].value = B["reinf"]
    ws["B14"].value = B["med"]
    ws["B15"].value = "='양식7_데이터시트'!$E$5"                 # 양식7 LTO와 항상 같은 문장
    ws["B16"].value = B["prevent"]
    ws["B17"].value = B["teach"]
    ws["B18"].value = B["reinforce"]
    for i, (a, b) in enumerate(B["crisis"]):
        ws[f"B{20+i}"].value = a
        ws[f"C{20+i}"].value = b
    for ref in ["C11", "C12", "B13", "B14", "B15", "B16", "B17", "B18"] + [f"{c}{r}" for r in (20, 21, 22) for c in "BC"]:
        c = ws[ref]
        c.font = Font(name=FONT, size=9, color=EX_TXT)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws["B15"].font = Font(name=FONT, size=9, color="000000")
    return ws


def patch_form7(wb, cfg):
    ws = wb["양식7_데이터시트"]
    _unify_terms(ws)
    Z = cfg["form7"]
    ws["E3"].value = Z["grade"]
    ws["E4"].value = f'{cfg["teacher"]} 선생님'
    ws["M4"].value = "중재 시작일 ~ (기초선 종료 후)"
    ws["E5"].value = Z["lto"].replace("회기", "회차")
    ws["E5"].font = Font(name=FONT, size=10, color=NAVY, bold=True)
    ws["E6"].value = "='양식6_행동중재계획'!$B$13"
    if isinstance(ws["E11"].value, str):
        ws["E11"].value = ws["E11"].value.replace("O = 촉구 없이", "촉구 없음")
    for k, v in Z.get("cells", {}).items():
        ws[k].value = v.replace("회기", "회차")
    # 예시 자료 삭제: 18~20행(STO·날짜) + 21~40행(시도)
    for row in ws["B18:P40"]:
        for c in row:
            c.value = None
    ws["A43"].value = ("한 교수 회차 = 시도 10회(21~30행 사용, 31~40행은 시도를 늘릴 때만). 결과는 O(정반응) · □(무반응) · X(오반응)로 고르고, "
                       "18행에 STO 번호, 19~20행에 시작·종료일을 넣으면 그래프가 그려집니다. 정반응률은 시도 수가 달라도 %로 비교됩니다.")
    ws["A44"].value = ("O는 '그 STO의 촉구 수준에서 맞게 함'입니다. 2.기록지의 ○(도움 없이 시작)와는 다른 기록이라 값이 달라도 정상입니다. "
                       "인쇄: 이 탭에서 Ctrl+P → A4 가로 → 배율 '너비 맞춤'.")
    ws["A44"].font = Font(name=FONT, size=10, color=BLUE, bold=True)
    ws["A45"].value = None
    return ws


def strip_predictions(wb):
    """양식1·2·3에 미리 들어 있던 '예측(가상)' 입력값을 지운다. 지운 값은 (탭, 칸, 값)으로 돌려줘 설명용 사본에서 복원할 수 있다."""
    out = []

    def clr(ws, coord):
        c = ws[coord]
        if c.value in (None, "") or (isinstance(c.value, str) and c.value.startswith("=")):
            return
        out.append((ws.title, coord, c.value))
        c.value = None
        f = copy(c.font)
        c.font = Font(name=f.name or FONT, size=f.size or 10, bold=f.bold, italic=False, color="000000")

    ws = wb["양식1_ABC기록지"]
    for r in range(7, 18):
        for col in range(2, 15):
            clr(ws, f"{CL(col)}{r}")
    ws = wb["양식3_QABF"]
    for r in range(8, 33):
        clr(ws, f"A{r}")
    clr(ws, "G3")
    ws = wb["양식2_선호자극"]
    hdr = [r for r in range(1, ws.max_row + 1) if isinstance(ws[f"A{r}"].value, str) and "좋아하는 것 BEST" in ws[f"A{r}"].value][0]
    for r in range(hdr + 1, hdr + 16):
        clr(ws, f"C{r}")
        v = ws[f"F{r}"].value
        if not (isinstance(v, str) and v.startswith("제한 없이")):
            clr(ws, f"F{r}")
    for row in ws.iter_rows():
        for c in row:
            if c.value == "✓":
                clr(ws, c.coordinate)
    return out


def restore_predictions(wb, stripped):
    for tab, coord, val in stripped:
        c = wb[tab][coord]
        c.value = val
        f = copy(c.font)
        c.font = Font(name=f.name or FONT, size=f.size or 10, italic=True, color=EX_TXT)


def clear_hyperlink_formulas(wb):
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=HYPERLINK("):
                    c.value = None


def finalize(wb, cfg, out_path, order):
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [s for s in wb._sheets if s.title not in order]
    wb.active = 0
    for ws in wb:
        ws.sheet_view.tabSelected = (ws.title == order[0])
    wb.calculation.fullCalcOnLoad = True
    tab_colors(wb)
    wb.properties.creator = ""
    wb.properties.lastModifiedBy = ""
    wb.save(out_path)


def tab_colors(wb):
    """탭 색: 남색=길잡이 · 주황=기록하는 곳 · 파랑=자동 결과 · 회색=정의/설계 · 보라=설명 스크립트(내부)"""
    for ws in wb:
        n = ws.title
        if n.startswith("0."):
            ws.sheet_properties.tabColor = "1F4E79"
        elif n in ("1.정의", "4.중재 설계(초안)"):
            ws.sheet_properties.tabColor = "7F7F7F"
        elif n.startswith("2.") or n.startswith("인쇄_") or n.startswith("양식4"):
            ws.sheet_properties.tabColor = "FFC000"
        elif n.startswith("3."):
            ws.sheet_properties.tabColor = "5B9BD5"
        elif n.startswith("설명 스크립트"):
            ws.sheet_properties.tabColor = "7030A0"
