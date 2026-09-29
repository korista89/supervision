# -*- coding: utf-8 -*-
"""2.기록지 · 인쇄 양식 · 분포도 생성기 (기록 방식 3종)
  seq    : 박윤아·강희강 — 한 줄=한 회차, 칸 10개에 한 글자씩 (지시/문항 순서)
  days   : 임충만 — 한 줄=하루, 교시별 횟수(벽 분포도를 옮겨 적음)
  counts : 김규리 — 한 줄=한 교시, ○△× 개수(벽 분포도를 옮겨 적음)
"""
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from lib import *

DATA_ROWS = 200
OK_FILL, HALF_FILL, BAD_FILL = "D9EAD3", "FCE5CD", "F4CCCC"     # 결과 칸 색(노랑과 겹치지 않게)


def _steps_block(ws, r0, lines, last_col):
    r = r0
    for i, (txt, blue) in enumerate(lines):
        merged(ws, r, 1, last_col, txt, color=BLUE if blue else "000000", bold=blue, bg="F7F9FC")
        r += 1
    return r


def _cf_outcome(ws, rng, cf):
    """cf = [(값, 채움색), ...] — 결과 칸 색(노랑과 겹치지 않는 색만 사용)"""
    for val, color in cf:
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{val}"'], fill=fill(color)))


# ═════════════ seq : 박윤아 / 강희강 ═════════════
def build_record_seq(wb, cfg):
    R = cfg["rec"]
    ws = wb.create_sheet("2.기록지")
    hide_gridlines(ws)
    n = R["n_items"]                                  # 10
    c_first = 6 if R['has_med'] else 5                # 지시/문항 첫 열
    c_last = c_first + n - 1                           # O
    c_extra = c_last + 1                               # P (턱 때리기 / 메모)
    # 열 정의: 입력열(A날짜 B회차 C단계 D교시 [E아침약]) ...
    ncol = R["ncol"]
    colw(ws, R["widths"])
    title(ws, f'="2.기록지 — {R["title"]}   (한 줄 = 한 회차 = {R["unit_txt"]})"', ncol)
    r = _steps_block(ws, 2, R["how"], ncol)
    # 예시 줄 설명
    hdr_grp = r + 1
    hdr = r + 2
    sub = r + 3           # 항목(지시/문항) 내용 줄
    ex = r + 4
    first = r + 5
    last = first + DATA_ROWS - 1
    R["_rows"] = (first, last)
    put(ws, f"{CL(c_first)}{hdr_grp}", R["seq_label"], bold=True, bg=SECT, h="center", border=True)
    ws.merge_cells(start_row=hdr_grp, start_column=c_first, end_row=hdr_grp, end_column=c_last)
    auto0 = R["auto_start"]
    put(ws, f"{CL(auto0)}{hdr_grp}", "자동(입력 금지)", bold=True, bg=AUTO, color=GRAY_TXT, h="center", border=True, size=9)
    ws.merge_cells(start_row=hdr_grp, start_column=auto0, end_row=hdr_grp, end_column=ncol)
    for i, h in enumerate(R["headers_left"]):
        put(ws, f"{CL(1+i)}{hdr}", h, bold=True, bg=SECT, h="center", border=True, size=9)
    for i in range(n):
        put(ws, f"{CL(c_first+i)}{hdr}", i + 1, bold=True, bg=SECT, h="center", border=True)
        put(ws, f"{CL(c_first+i)}{sub}", f"=IF('1.정의'!$B${cfg['_listrows'][i]}=\"\",\"\",'1.정의'!$B${cfg['_listrows'][i]})", size=8, h="center", border=True,
            color=GRAY_TXT, bg=LABEL)
    put(ws, f"A{sub}", R["sub_label"], size=8, color=GRAY_TXT, h="center")
    ws.merge_cells(start_row=sub, start_column=1, end_row=sub, end_column=5 if R["has_med"] else 4)
    for i, h in enumerate(R["headers_right"]):
        c = c_last + 1 + i
        put(ws, f"{CL(c)}{hdr}", h, bold=True, bg=SECT, h="center", border=True, size=9)
    for i, h in enumerate(R["headers_auto"]):
        put(ws, f"{CL(auto0+i)}{hdr}", h, bold=True, bg=AUTO, color=GRAY_TXT, h="center", border=True, size=9)
    ws.row_dimensions[hdr].height = 30
    ws.row_dimensions[sub].height = 46
    # 예시 줄(계산에 들어가지 않음)
    put(ws, f"A{ex}", "예시 →", italic=True, color=EX_TXT, bold=True, h="center", border=True)
    for i, v in enumerate(R["example"]):
        put(ws, f"{CL(2+i)}{ex}", v, italic=True, color=EX_TXT, h="center", border=True, size=9)
    ws.row_dimensions[ex].height = 22
    # 데이터 줄
    for rr in range(first, last + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=rr, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(c == R["memo_col"]))
            cell.font = font(size=10)
        ws[f"A{rr}"].number_format = "m/d"
        put(ws, f"B{rr}", f'=IF($A{rr}="","",ROW()-{first-1})', bg=AUTO, color=GRAY_TXT, h="center", border=True)
        for i, f in enumerate(R["auto_formulas"]):
            put(ws, f"{CL(auto0+i)}{rr}", f.format(r=rr), bg=AUTO, color=GRAY_TXT, h="center", border=True, size=9,
                nf=R.get("auto_nf", {}).get(i))
        if rr < first + R["yellow_rows"]:                     # 4회기 전까지 최소 이만큼
            for c in list(range(1, 5 if not R["has_med"] else 6)) + list(range(c_first, ncol_input_end(R) + 1)):
                if c == 2:
                    continue
                ws.cell(row=rr, column=c).fill = fill(YELLOW)
    # 드롭다운
    dv_date(ws, f"A{first}:A{last}")
    dv_list(ws, f"C{first}:C{last}", ["기초선", "중재", "유지"])
    dv_list(ws, f"D{first}:D{last}", [f"{i}교시" for i in range(1, 8)])
    if R["has_med"]:
        dv_list(ws, f"E{first}:E{last}", ["복용", "미복용", "모름"])
    dv_list(ws, f"{CL(c_first)}{first}:{CL(c_last)}{last}", R["codes"], prompt=R["code_prompt"])
    if R.get("count_col"):
        dv_whole(ws, f"{R['count_col']}{first}:{R['count_col']}{last}", 0, 99)
    rng = f"{CL(c_first)}{first}:{CL(c_last)}{last}"
    _cf_outcome(ws, rng, R["cf"])
    ws.freeze_panes = ws[f"C{first}"]
    return ws


def ncol_input_end(R):
    """입력 열의 마지막(=자동 열 바로 앞)."""
    return R["auto_start"] - 1


def build_print_seq(wb, cfg):
    """A4 가로 인쇄 양식 — 2.기록지와 열 순서가 같아 1:1로 옮겨 적는다."""
    R = cfg["rec"]
    P = cfg["print"]
    ws = wb.create_sheet("인쇄_기록지")
    hide_gridlines(ws)
    n = R["n_items"]
    c_first = 6 if R['has_med'] else 5
    c_last = c_first + n - 1
    pcols = P["ncol"]
    colw(ws, P["widths"])
    title(ws, P["title"], pcols)
    merged(ws, 2, 1, pcols, P["how"], color="000000", size=9, bg="F7F9FC")
    merged(ws, 3, 1, pcols, P["who"], size=10)
    hdr, sub = 4, 5
    for i, h in enumerate(P["headers"]):
        put(ws, f"{CL(1+i)}{hdr}", h, bold=True, bg=SECT, h="center", border=True, size=9)
    for i in range(n):
        put(ws, f"{CL(c_first+i)}{hdr}", i + 1, bold=True, bg=SECT, h="center", border=True)
        put(ws, f"{CL(c_first+i)}{sub}", f"=IF('1.정의'!$B${cfg['_listrows'][i]}=\"\",\"\",'1.정의'!$B${cfg['_listrows'][i]})", size=7, h="center", border=True,
            color=GRAY_TXT, bg=LABEL)
    put(ws, f"A{sub}", P["sub_label"], size=8, color=GRAY_TXT, h="center")
    ws.merge_cells(start_row=sub, start_column=1, end_row=sub, end_column=5 if R["has_med"] else 4)
    ws.row_dimensions[hdr].height = 28
    ws.row_dimensions[sub].height = 44
    for rr in range(sub + 1, sub + 1 + P["rows"]):
        ws.row_dimensions[rr].height = 27
        for c in range(1, pcols + 1):
            put(ws, f"{CL(c)}{rr}", None, border=True, h="center")
    merged(ws, sub + P["rows"] + 1, 1, pcols, P["foot"], size=9, color=GRAY_TXT)
    print_setup(ws, f"A1:{CL(pcols)}{sub+P['rows']+1}", landscape=True)
    return ws


# ═════════════ days : 임충만 ═════════════
def build_record_days(wb, cfg):
    R = cfg["rec"]
    ws = wb.create_sheet("2.기록지")
    hide_gridlines(ws)
    colw(ws, R["widths"])
    ncol = R["ncol"]
    title(ws, f'="2.기록지 — 감소: "&\'1.정의\'!$B$7&"   (한 줄 = 하루)"', ncol)
    r = _steps_block(ws, 2, R["how"], ncol)
    hdr = r + 1
    ex = r + 2
    first = r + 3
    last = first + DATA_ROWS - 1
    R["_rows"] = (first, last)
    put(ws, f"D{hdr-0}", None)
    for i, h in enumerate(R["headers"]):
        auto = i >= R["auto_start"] - 1
        put(ws, f"{CL(1+i)}{hdr}", h, bold=True, bg=AUTO if auto else SECT, color=GRAY_TXT if auto else "000000",
            h="center", border=True, size=9)
    ws.row_dimensions[hdr].height = 44
    put(ws, f"A{ex}", "예시 →", italic=True, color=EX_TXT, bold=True, h="center", border=True)
    for i, v in enumerate(R["example"]):
        put(ws, f"{CL(2+i)}{ex}", v, italic=True, color=EX_TXT, h="center", border=True, size=9)
    ws.row_dimensions[ex].height = 22
    a0 = R["auto_start"]
    for rr in range(first, last + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=rr, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(c == R["memo_col"]))
            cell.font = font(size=10)
        ws[f"A{rr}"].number_format = "m/d"
        put(ws, f"B{rr}", f'=IF($A{rr}="","",ROW()-{first-1})', bg=AUTO, color=GRAY_TXT, h="center", border=True)
        for i, f in enumerate(R["auto_formulas"]):
            put(ws, f"{CL(a0+i)}{rr}", f.format(r=rr), bg=AUTO, color=GRAY_TXT, h="center", border=True, size=9,
                nf=R.get("auto_nf", {}).get(i))
        if rr < first + R["yellow_rows"]:
            for c in range(1, a0):
                if c != 2:
                    ws.cell(row=rr, column=c).fill = fill(YELLOW)
    dv_date(ws, f"A{first}:A{last}")
    dv_list(ws, f"C{first}:C{last}", ["기초선", "중재", "유지"])
    p1, p2 = R["period_cols"]
    dv_whole(ws, f"{p1}{first}:{p2}{last}", 0, 99)
    dv_list(ws, f"{R['star_col']}{first}:{R['star_col']}{last}", ["★"])
    dv_whole(ws, f"{R['req_col']}{first}:{R['req_col']}{last}", 0, 99)
    ws.conditional_formatting.add(f"{R['star_col']}{first}:{R['star_col']}{last}",
                                  CellIsRule(operator="equal", formula=['"★"'], fill=fill(BAD_FILL)))
    ws.freeze_panes = ws[f"D{first}"]
    return ws


def build_wall_days(wb, cfg):
    """임충만: 교실 벽에 붙이는 분포도(교시 × 요일). 한 칸 = 그 교시 횟수(+★)."""
    W = cfg["wall"]
    ws = wb.create_sheet("양식4_분포도(인쇄용)")
    hide_gridlines(ws)
    colw(ws, {"A": 12, "B": 14, "C": 17, "D": 17, "E": 17, "F": 17, "G": 17})
    title(ws, '="분포도 기록지 — "&\'1.정의\'!$B$7&"   (교실 벽에 붙여 두고 매 교시 기록)"', 7)
    merged(ws, 2, 1, 7, f'="학생: "&\'1.정의\'!$B$2&"        기간:  ____월 ____일(월) ~ ____월 ____일(금)        담임: {cfg["teacher"]} 선생님"',
           size=10)
    merged(ws, 3, 1, 7, W["ask"], color=BLUE, bold=True)
    merged(ws, 4, 1, 7, W["star"], color=BLUE, bold=True)
    for i, h in enumerate(["교시", "시간"] + ["월", "화", "수", "목", "금"]):
        put(ws, f"{CL(1+i)}5", h, bold=True, bg=SECT, h="center", border=True)
    ws.row_dimensions[5].height = 24
    r = 5
    for lab, tm in W["periods"]:
        r += 1
        put(ws, f"A{r}", lab, bold=True, h="center", border=True)
        put(ws, f"B{r}", tm, bg=YELLOW, h="center", border=True, size=9)
        for c in range(3, 8):
            put(ws, f"{CL(c)}{r}", None, border=True)
        ws.row_dimensions[r].height = 38
    for lab in W["bottom"]:
        r += 1
        put(ws, f"A{r}", lab, bold=True, bg=LABEL, border=True, size=9)
        ws.merge_cells(f"A{r}:B{r}")
        ws[f"B{r}"].border = BORDER
        for c in range(3, 8):
            put(ws, f"{CL(c)}{r}", None, border=True)
        ws.row_dimensions[r].height = 30
    r += 1
    merged(ws, r, 1, 7, W["foot"], size=9, color=GRAY_TXT)
    print_setup(ws, f"A1:G{r}", landscape=True)
    return ws


# ═════════════ counts : 김규리 ═════════════
def build_record_counts(wb, cfg):
    R = cfg["rec"]
    ws = wb.create_sheet("2.기록지")
    hide_gridlines(ws)
    colw(ws, R["widths"])
    ncol = R["ncol"]
    title(ws, f'="2.기록지 — "&\'1.정의\'!$B$7&"   (한 줄 = 한 교시)"', ncol)
    r = _steps_block(ws, 2, R["how"], ncol)
    hdr = r + 1
    ex = r + 2
    first = r + 3
    last = first + DATA_ROWS - 1
    R["_rows"] = (first, last)
    for i, h in enumerate(R["headers"]):
        auto = i >= R["auto_start"] - 1
        put(ws, f"{CL(1+i)}{hdr}", h, bold=True, bg=AUTO if auto else SECT, color=GRAY_TXT if auto else "000000",
            h="center", border=True, size=9)
    ws.row_dimensions[hdr].height = 44
    put(ws, f"A{ex}", "예시 →", italic=True, color=EX_TXT, bold=True, h="center", border=True)
    for i, v in enumerate(R["example"]):
        put(ws, f"{CL(2+i)}{ex}", v, italic=True, color=EX_TXT, h="center", border=True, size=9)
    ws.row_dimensions[ex].height = 22
    a0 = R["auto_start"]
    for rr in range(first, last + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=rr, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(c == R["memo_col"]))
            cell.font = font(size=10)
        ws[f"A{rr}"].number_format = "m/d"
        prev = rr - 1
        put(ws, f"B{rr}", f'=IF($A{rr}="","",IF(ISNUMBER($B{prev}),IF($A{rr}=$A{prev},$B{prev},$B{prev}+1),1))',
            bg=AUTO, color=GRAY_TXT, h="center", border=True)
        for i, f in enumerate(R["auto_formulas"]):
            put(ws, f"{CL(a0+i)}{rr}", f.format(r=rr), bg=AUTO, color=GRAY_TXT, h="center", border=True, size=9)
        if rr < first + R["yellow_rows"]:
            for c in range(1, a0):
                if c != 2:
                    ws.cell(row=rr, column=c).fill = fill(YELLOW)
    dv_date(ws, f"A{first}:A{last}")
    dv_list(ws, f"C{first}:C{last}", ["기초선", "중재", "유지"])
    dv_list(ws, f"D{first}:D{last}", [f"{i}교시" for i in range(1, 8)])
    dv_whole(ws, f"{R['count_cols'][0]}{first}:{R['count_cols'][1]}{last}", 0, 30)
    dv_list(ws, f"{R['part_col']}{first}:{R['part_col']}{last}", ["Y", "N"])
    dv_list(ws, f"{R['crisis_col']}{first}:{R['crisis_col']}{last}", ["있음"])
    res = R["result_col"]
    rr_ = f"{res}{first}:{res}{last}"
    _cf_outcome(ws, rr_, [("◎", OK_FILL), ("○", HALF_FILL), ("×", BAD_FILL)])
    ws.freeze_panes = ws[f"E{first}"]
    return ws


def build_wall_counts(wb, cfg):
    """김규리: 통합학급 교과 선생님도 쓰는 분포도. 교시 × 요일, 칸마다 ○ △ × 개수 + 끝까지 참여."""
    W = cfg["wall"]
    ws = wb.create_sheet("양식4_분포도(인쇄용)")
    hide_gridlines(ws)
    widths = {"A": 9, "B": 12}
    for c in range(3, 23):
        widths[CL(c)] = 5.2
    colw(ws, widths)
    title(ws, '="분포도 기록지 — "&\'1.정의\'!$B$7&"   (교실 벽에 붙여 두고 매 교시 기록)"', 22)
    merged(ws, 2, 1, 22, f'="학생: "&\'1.정의\'!$B$2&"        기간:  ____월 ____일(월) ~ ____월 ____일(금)        담임: {cfg["teacher"]} 선생님"',
           size=10)
    merged(ws, 3, 1, 22, W["ask"], color=BLUE, bold=True)
    merged(ws, 4, 1, 22, W["safe"], color=BLUE, bold=True)
    put(ws, "A5", "교시", bold=True, bg=SECT, h="center", border=True)
    put(ws, "B5", "시간·교과", bold=True, bg=SECT, h="center", border=True)
    ws.merge_cells("A5:A6")
    ws.merge_cells("B5:B6")
    days = ["월", "화", "수", "목", "금"]
    for d in range(5):
        c0 = 3 + d * 4
        ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 3)
        put(ws, f"{CL(c0)}5", days[d], bold=True, bg=SECT, h="center", border=True)
        for k in range(4):
            ws.cell(row=5, column=c0 + k).border = BORDER
        for k, h in enumerate(["○", "△", "×", "끝"]):
            put(ws, f"{CL(c0+k)}6", h, bold=True, bg=SECT, h="center", border=True)
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 22
    r = 6
    for lab, tm in W["periods"]:
        r += 1
        put(ws, f"A{r}", lab, bold=True, h="center", border=True)
        put(ws, f"B{r}", tm, bg=YELLOW, h="center", border=True, size=8)
        for c in range(3, 23):
            put(ws, f"{CL(c)}{r}", None, border=True, h="center")
        ws.row_dimensions[r].height = 34
    r += 1
    put(ws, f"A{r}", "담임 확인", bold=True, bg=LABEL, border=True, size=9)
    ws.merge_cells(f"A{r}:B{r}")
    for c in range(3, 23):
        put(ws, f"{CL(c)}{r}", None, border=True)
    ws.row_dimensions[r].height = 26
    r += 1
    merged(ws, r, 1, 22, W["foot"], size=9, color=GRAY_TXT)
    print_setup(ws, f"A1:V{r}", landscape=True)
    return ws
