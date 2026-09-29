# -*- coding: utf-8 -*-
"""3.요약 탭(표 + 기초선 판정 + 그래프) 공통 생성기."""
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.line import LineProperties
from lib import *
from openpyxl.formatting.rule import FormulaRule

ROW0, ROW1 = 6, 45          # 표 시작/끝(회차 1~40)
HELP0 = 24                  # 그래프용 계산 열 시작(X)
C_BASE, C_INT, C_BAR = "1F4E79", "C55A11", "BFBFBF"


def build_summary(wb, cfg):
    S = cfg["summary"]
    ws = wb.create_sheet("3.요약")
    hide_gridlines(ws)
    R0, R1 = S["rec_rows"]
    rec = "'2.기록지'!"
    cols = S["cols"]
    ncol = 3 + len(cols)                                  # A~C + 지표 열
    last_tab_col = ncol
    colw(ws, {"A": 7, "B": 9, "C": 8})
    for i, c in enumerate(cols):
        ws.column_dimensions[CL(4 + i)].width = c.get("w", 11)
    box_c = last_tab_col + 2                              # 판정 상자 시작 열
    for j, w in enumerate([22, 20, 20]):
        ws.column_dimensions[CL(box_c + j)].width = w
    title(ws, f'="3.요약·그래프 — "&\'1.정의\'!$B$7', box_c + 2)
    merged(ws, 2, 1, box_c + 2, S["desc"], color=GRAY_TXT, size=10)
    merged(ws, 3, 1, box_c + 2, S["guide"], color=BLUE, bold=True, size=10)
    # ── 표 머리 ──
    heads = [S["unit"], "날짜", "단계"] + [c["h"] for c in cols]
    for i, h in enumerate(heads):
        put(ws, f"{CL(1+i)}5", h, bold=True, bg=SECT, h="center", border=True, size=9)
    ws.row_dimensions[5].height = 42
    for r in range(ROW0, ROW1 + 1):
        k = r - ROW0 + 1
        put(ws, f"A{r}", k, h="center", border=True, bg=AUTO, color=GRAY_TXT)
        put(ws, f"B{r}", f'=IFERROR(INDEX({rec}$A${R0}:$A${R1},MATCH($A{r},{rec}$B${R0}:$B${R1},0)),"")',
            h="center", border=True, bg=AUTO, nf="m/d", color=GRAY_TXT)
        put(ws, f"C{r}", f'=IFERROR(INDEX({rec}$C${R0}:$C${R1},MATCH($A{r},{rec}$B${R0}:$B${R1},0)),"")',
            h="center", border=True, bg=AUTO, color=GRAY_TXT)
        for i, c in enumerate(cols):
            f = c["f"].format(r=r, k=f"$A{r}", rec=rec, R0=R0, R1=R1)
            put(ws, f"{CL(4+i)}{r}", f, h="center", border=True, bg=AUTO, nf=c.get("nf"), color=GRAY_TXT)
    colkey = {c["key"]: CL(4 + i) for i, c in enumerate(cols)}
    cfg["_sumcol"] = colkey
    # ── 기초선 판정 상자 ──
    b0, b1, b2 = CL(box_c), CL(box_c + 1), CL(box_c + 2)
    put(ws, f"{b0}5", "기초선 간이 판정", bold=True, color="FFFFFF", bg=NAVY, border=True)
    m = S["box"]                                          # [(key, label, nf), (key, label, nf)]
    put(ws, f"{b1}5", m[0][1], bold=True, bg=SECT, h="center", border=True, size=9)
    put(ws, f"{b2}5", m[1][1] if len(m) > 1 else "", bold=True, bg=SECT, h="center", border=True, size=9)
    labels = ["기초선 점 수", "기초선 평균", "기초선 최소~최대", "중재·유지 평균", "판정(기초선)"]
    for i, lab in enumerate(labels):
        put(ws, f"{b0}{6+i}", lab, bold=True, bg=LABEL, border=True)
    for j, (key, lab, nf) in enumerate(m):
        col = colkey[key]
        X = f"${col}${ROW0}:${col}${ROW1}"
        P = f"$C${ROW0}:$C${ROW1}"
        bc = CL(box_c + 1 + j)
        put(ws, f"{bc}6", f'=COUNTIFS({P},"기초선",{X},">=0")', h="center", border=True, bg=AUTO)
        put(ws, f"{bc}7", f'=IF({bc}6=0,"",AVERAGEIFS({X},{P},"기초선"))', h="center", border=True, bg=AUTO, nf=nf)
        put(ws, f"{bc}8", f'=IF({bc}6=0,"",TEXT(_xlfn.MINIFS({X},{P},"기초선"),"{nf}")&" ~ "&TEXT(_xlfn.MAXIFS({X},{P},"기초선"),"{nf}"))',
            h="center", border=True, bg=AUTO)
        put(ws, f"{bc}9", f'=IFERROR(AVERAGEIFS({X},{P},"중재",{X},">=0"),"")' if False else
            f'=IF(COUNTIFS({P},"중재",{X},">=0")+COUNTIFS({P},"유지",{X},">=0")=0,"",'
            f'(SUMIFS({X},{P},"중재")+SUMIFS({X},{P},"유지"))/(COUNTIFS({P},"중재",{X},">=0")+COUNTIFS({P},"유지",{X},">=0")))',
            h="center", border=True, bg=AUTO, nf=nf)
        put(ws, f"{bc}10",
            f'=IF({bc}6=0,"기록 시작",IF({bc}6<3,"자료점 부족: 3점까지 계속",'
            f'IF(_xlfn.MAXIFS({X},{P},"기초선")=0,"모두 0: 0도 자료",'
            f'IF((_xlfn.MAXIFS({X},{P},"기초선")-_xlfn.MINIFS({X},{P},"기초선"))/{bc}7>1,"변동 큼: 2~3점 추가","안정: 중재 시작 협의"))))',
            h="center", border=True, bg=AUTO, size=9)
    cfg["_box"] = {"n1": f"{b1}6", "n2": f"{b2}6"}
    for i, (lab, f) in enumerate(S.get("extra_box", [])):
        f = f.replace("{R0}", str(R0)).replace("{R1}", str(R1))
        r = 11 + i
        put(ws, f"{b0}{r}", lab, bold=True, bg=LABEL, border=True)
        ws.merge_cells(f"{b1}{r}:{b2}{r}")
        put(ws, f"{b1}{r}", f, h="center", border=True, bg=AUTO)
        ws[f"{b2}{r}"].border = BORDER
    nr = 11 + len(S.get("extra_box", []))
    ws.merge_cells(f"{b0}{nr}:{b2}{nr+2}")
    put(ws, f"{b0}{nr}", "판정은 점 수·범위만 보는 간이 기준입니다. 이미 좋아지는 중인지(추세)는 아래 그래프로 함께 봅니다. "
        "안전 문제는 판정과 관계없이 예방·위기 절차를 바로 적용합니다.", color=GRAY_TXT, size=9, v="top")
    # ── 그래프용 계산 열 ──
    hx = CL(HELP0)
    put(ws, f"{hx}4", "그래프용 계산(자동) — 건드리지 마세요", color=EX_TXT, size=8, wrap=False)
    put(ws, f"{hx}5", S["unit"], size=8, color=EX_TXT)
    for r in range(ROW0, ROW1 + 1):
        put(ws, f"{hx}{r}", f'=IF($B{r}="","",$A{r})', size=8, color=EX_TXT)
    chart_cols = {}
    hc = HELP0 + 1
    for ch in S["charts"]:
        col = colkey[ch["key"]]
        mult = "*100" if ch["pct"] else ""
        cb, ci, cbar = CL(hc), CL(hc + 1), CL(hc + 2)
        chart_cols[ch["key"]] = (cb, ci, cbar)
        put(ws, f"{cb}5", "기초선", size=8, color=EX_TXT)
        put(ws, f"{ci}5", "중재·유지", size=8, color=EX_TXT)
        put(ws, f"{cbar}5", "단계 전환", size=8, color=EX_TXT)
        for r in range(ROW0, ROW1 + 1):
            v = f"{col}{r}"
            put(ws, f"{cb}{r}", f'=IF(AND($B{r}<>"",$C{r}="기초선",ISNUMBER({v})),{v}{mult},NA())', size=8, color=EX_TXT)
            put(ws, f"{ci}{r}", f'=IF(AND($B{r}<>"",OR($C{r}="중재",$C{r}="유지"),ISNUMBER({v})),{v}{mult},NA())', size=8, color=EX_TXT)
            top = "100" if ch["pct"] else f"MAX(1,MAX(${col}${ROW0}:${col}${ROW1})*1.2)"
            put(ws, f"{cbar}{r}", f'=IF(OR(AND($C{r}="중재",$C{r-1}="기초선"),AND($C{r}="유지",$C{r-1}="중재")),{top},NA())',
                size=8, color=EX_TXT)
        hc += 3
    # 그래프용 계산 열의 #N/A는 글자를 흰색으로 가려 오류처럼 보이지 않게 한다
    hl = CL(HELP0 + len(S["charts"]) * 3)
    ws.conditional_formatting.add(f"{hx}{ROW0}:{hl}{ROW1}",
                                  FormulaRule(formula=[f"ISNA({hx}{ROW0})"], font=Font(color="FFFFFF")))
    # ── 그래프 ──
    anchor_row = 14
    for ch in S["charts"]:
        cb, ci, cbar = chart_cols[ch["key"]]
        _phase_chart(ws, ch, hx, cb, ci, cbar, f"{b0}{anchor_row}")
        anchor_row += 17
    ws.freeze_panes = "A6"
    return ws


def _col_idx(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def _phase_chart(ws, ch, xcol, cb, ci, cbar, anchor):
    xref = Reference(ws, min_col=_col_idx(xcol), min_row=ROW0, max_row=ROW1)
    bar = BarChart()
    bar.type = "col"
    bar.gapWidth = 20
    s = Reference(ws, min_col=_col_idx(cbar), min_row=5, max_row=ROW1)
    bar.add_data(s, titles_from_data=True)
    bar.set_categories(xref)
    bar.series[0].graphicalProperties = GraphicalProperties(solidFill=C_BAR)
    bar.series[0].graphicalProperties.line.noFill = True
    line = LineChart()
    for colL, color in ((cb, C_BASE), (ci, C_INT)):
        line.add_data(Reference(ws, min_col=_col_idx(colL), min_row=5, max_row=ROW1), titles_from_data=True)
        sr = line.series[-1]
        sr.marker.symbol = "circle"
        sr.marker.size = 8
        sr.marker.graphicalProperties = GraphicalProperties(solidFill=color)
        sr.marker.graphicalProperties.line.solidFill = color
        sr.graphicalProperties = GraphicalProperties()
        sr.graphicalProperties.line = LineProperties(solidFill=color, w=22000)
        sr.smooth = False
    line.set_categories(xref)
    bar.y_axis.scaling.min = 0
    if ch["pct"]:
        bar.y_axis.scaling.max = 100
        bar.y_axis.majorUnit = 20
    bar.y_axis.title = ch.get("ytitle", "%" if ch["pct"] else "횟수")
    bar.x_axis.title = ws.cell(row=5, column=1).value
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.title = ch["title"]
    bar.legend.position = "b"
    bar.height, bar.width = 8.2, 17.5
    bar.dispBlanksAs = "gap"
    bar.visible_cells_only = False
    bar += line
    ws.add_chart(bar, anchor)
