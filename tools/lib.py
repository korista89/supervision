# -*- coding: utf-8 -*-
"""공통 스타일·헬퍼. 규칙: 노란 칸=4회기 전까지 채울(확인할) 칸 / 파란 글씨=꼭 지킬 지침 /
회색 칸=자동 계산(입력 금지) / 회색 기울임 글씨=예시. 빨간 글씨·하늘색 칸은 쓰지 않는다."""
import math
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

FONT = "맑은 고딕"
NAVY, YELLOW, AUTO, LABEL, SECT, BLUE = "1F4E79", "FFFF00", "EDEDED", "F2F2F2", "DDE7F2", "0000FF"
GRAY_TXT, EX_TXT = "555555", "808080"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


def font(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


def put(ws, ref, value=None, *, bold=False, color="000000", bg=None, size=10, italic=False,
        wrap=True, h="left", v="center", border=False, nf=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font(bold, color, size, italic)
    if bg:
        c.fill = fill(bg)
    c.alignment = Alignment(wrap_text=wrap, horizontal=h, vertical=v)
    if border:
        c.border = BORDER
    if nf:
        c.number_format = nf
    return c


def dispw(s):
    """대략적인 표시 폭(한글=2, 그 외=1.05)."""
    w = 0.0
    for ch in s:
        w += 2.0 if ord(ch) > 0x2E80 or ch in "○△×◎★㉮㉯㉰①②③④⑤⑥⑦" else 1.05
    return w


def lines_needed(text, width_units):
    n = 0
    for ln in str(text).split("\n"):
        n += max(1, math.ceil(dispw(ln) * 1.12 / max(width_units - 2.5, 4)))
    return n


def fit_row(ws, row, text, width_units, size=10, minh=18, pad=5):
    lh = 13.8 if size <= 10 else 15.5
    h = max(minh, lines_needed(text, width_units) * lh + pad)
    cur = ws.row_dimensions[row].height or 0
    ws.row_dimensions[row].height = max(cur, h)


def colw(ws, widths):
    for k, v in widths.items():
        ws.column_dimensions[k].width = v


def span_width(ws, c1, c2):
    tot = 0
    for i in range(c1, c2 + 1):
        tot += ws.column_dimensions[CL(i)].width or 8.43
    return tot


def merged(ws, row, c1, c2, value, **kw):
    """row의 c1~c2 열을 병합하고 값을 넣는다(높이 자동 추정)."""
    ref = f"{CL(c1)}{row}"
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    size = kw.get("size", 10)
    put(ws, ref, value, **kw)
    if kw.get("bg"):
        for i in range(c1, c2 + 1):
            ws.cell(row=row, column=i).fill = fill(kw["bg"])
    if isinstance(value, str) and not value.startswith("="):
        fit_row(ws, row, value, span_width(ws, c1, c2), size=size)
    return ref


def title(ws, text, last_col, row=1):
    merged(ws, row, 1, last_col, text, bold=True, color="FFFFFF", bg=NAVY, size=13, v="center")
    ws.row_dimensions[row].height = 28


def dv_list(ws, rng, items, allow_blank=True, prompt=None):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=allow_blank,
                        showErrorMessage=True, errorTitle="목록에서 고르세요",
                        error="정해진 값만 입력할 수 있습니다: " + " / ".join(items))
    if prompt:
        dv.promptTitle, dv.prompt, dv.showInputMessage = "입력 안내", prompt, True
    ws.add_data_validation(dv)
    dv.add(rng)
    return dv


def dv_date(ws, rng):
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                        showErrorMessage=True, errorTitle="날짜", error="날짜 형식으로 입력하세요. 예: 10/6")
    ws.add_data_validation(dv)
    dv.add(rng)


def dv_whole(ws, rng, lo=0, hi=99):
    dv = DataValidation(type="whole", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True,
                        showErrorMessage=True, errorTitle="숫자", error=f"{lo}~{hi} 사이의 정수를 입력하세요.")
    ws.add_data_validation(dv)
    dv.add(rng)


def print_setup(ws, area, landscape=True, fit=True):
    ws.print_area = area
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = 9
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0 if fit else 1
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.print_options.horizontalCentered = True


def legend_row(ws, row, spans):
    """노랑/파랑/회색 사용 규칙을 한 줄로. spans = [(c1,c2),(c1,c2),(c1,c2)]"""
    items = [("노란 칸 = 4회기 전까지 채우거나 확인할 칸", YELLOW, "000000", True),
             ("파란 글씨 = 꼭 지킬 지침", None, BLUE, True),
             ("회색 칸 = 자동 계산, 입력하지 않기", AUTO, GRAY_TXT, True)]
    for (c1, c2), (txt, bg, col, bold) in zip(spans, items):
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        put(ws, f"{CL(c1)}{row}", txt, bold=bold, color=col, bg=bg, h="center", size=9)
        for c in range(c1, c2 + 1):
            ws.cell(row=row, column=c).border = BORDER
            if bg:
                ws.cell(row=row, column=c).fill = fill(bg)
    ws.row_dimensions[row].height = 26


def hide_gridlines(ws):
    ws.sheet_view.showGridLines = False
