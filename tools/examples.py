# -*- coding: utf-8 -*-
"""설명용(내부) 사본에 넣는 예시(가상) 자료 — 기초선 → 중재 → 유지가 표·그래프로 어떻게 채워지는지 보여 주기 위한 것.
최종 배포본에는 넣지 않는다. 난수는 고정 시드라 실행할 때마다 같은 자료가 나온다."""
import datetime, random, openpyxl

import os
SRC = os.environ.get("SRC_DIR", os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "source"))


def school_days(start, n):
    d, out = start, []
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d += datetime.timedelta(days=1)
    return out


def _first_row(ws):
    return [r for r in range(1, 40) if ws[f"A{r}"].value == "예시 →"][0] + 1


def _setdate(ws, ref, d):
    ws[ref].value = d
    ws[ref].number_format = "m/d"


def form7_from_source(wb, cfg):
    src = openpyxl.load_workbook(f"{SRC}/{cfg['src']}")["양식7_데이터시트"]
    dst = wb["양식7_데이터시트"]
    for row in src["B18:P40"]:
        for c in row:
            v = c.value
            if v is not None and hasattr(v, "year"):
                v = v
            dst[c.coordinate].value = v
    for r in (19, 20):
        for c in "BCDEFGHIJKLMNOP":
            if dst[f"{c}{r}"].value is not None:
                dst[f"{c}{r}"].number_format = "m/d"


def fill_park(wb, cfg):
    s = openpyxl.load_workbook(f"{SRC}/{cfg['src']}")["2.기록지"]
    ws = wb["2.기록지"]
    first = _first_row(ws)
    r_out = first
    for r in range(14, 34):
        if s[f"A{r}"].value is None:
            continue
        _setdate(ws, f"A{r_out}", s[f"A{r}"].value)
        for col in "CDEFGHIJKLMNOPQ":
            v = s[f"{col}{r}"].value
            if col == "Q" and isinstance(v, str):
                v = v.replace("〔예시〕", "").strip() or None
            if col == "P" and v is not None:
                v = int(v)
            ws[f"{col}{r_out}"].value = v
        r_out += 1


def fill_kang(wb, cfg):
    rnd = random.Random(11)
    ws = wb["2.기록지"]
    first = _first_row(ws)
    plan = [("기초선", 0.80, 0.02)] * 5 + [("중재", p, q) for p, q in
                                          [(0.6, 0.2), (0.5, 0.3), (0.45, 0.4), (0.35, 0.5), (0.3, 0.6), (0.25, 0.6), (0.2, 0.7), (0.15, 0.8), (0.1, 0.8), (0.1, 0.8)]] + \
           [("유지", 0.1, 0.8)] * 2
    days = school_days(datetime.datetime(2026, 10, 5), len(plan))
    for i, ((ph, pT, pQ), d) in enumerate(zip(plan, days)):
        r = first + i
        _setdate(ws, f"A{r}", d)
        ws[f"C{r}"].value = ph
        ws[f"D{r}"].value = "3교시"
        for j in range(10):
            new = j >= 5
            if rnd.random() < pT * (0.85 if not new else 1.15):
                v = "T"
            elif new and rnd.random() < pQ:
                v = "요"
            else:
                v = "○"
            ws.cell(row=r, column=5 + j).value = v
        if i == 2:
            ws[f"O{r}"].value = "8번: 선생님을 부른 뒤 답함"


def fill_lim(wb, cfg):
    rnd = random.Random(5)
    ws = wb["2.기록지"]
    first = _first_row(ws)
    plan = [("기초선", 1.1)] * 6 + [("중재", m) for m in (0.9, 0.8, 0.65, 0.55, 0.5, 0.4, 0.35, 0.3, 0.2, 0.2)] + [("유지", 0.1)] * 2
    days = school_days(datetime.datetime(2026, 10, 5), len(plan))
    for i, ((ph, mean), d) in enumerate(zip(plan, days)):
        r = first + i
        _setdate(ws, f"A{r}", d)
        ws[f"C{r}"].value = ph
        star = False
        for j in range(8):
            if i == 4 and j == 7:
                continue                                    # 못 본 교시는 빈칸
            v = max(0, int(round(rnd.gauss(mean, 0.7))))
            if j == 3:                                      # 점심은 낮게
                v = min(v, 1)
            ws.cell(row=r, column=4 + j).value = v
            star = star or (v >= 3)
        if i in (1, 3):
            star = True
        ws[f"L{r}"].value = "★" if star else None
        idx = i - 6
        ws[f"M{r}"].value = 0 if idx < 0 else min(6, 1 + idx // 2 + (1 if idx > 4 else 0))
        if i == 3:
            ws[f"N{r}"].value = "5교시 이마 붉어져 보건실 확인"
        if i == 7:
            ws[f"N{r}"].value = "중재 시작"


def fill_kim(wb, cfg):
    rnd = random.Random(23)
    ws = wb["2.기록지"]
    first = _first_row(ws)
    plan = [("기초선", (1, 1, 3))] * 6 + [("중재", t) for t in
                                          [(2, 1, 2), (2, 1, 2), (3, 1, 1), (3, 1, 1), (3, 0, 1), (4, 1, 1), (4, 0, 0), (4, 0, 1), (5, 0, 0), (5, 0, 0)]] + [("유지", (5, 0, 0))] * 3
    days = school_days(datetime.datetime(2026, 10, 5), len(plan))
    subj = ["국어", "수학", "사회", "과학", "통합체육", "음악"]
    r = first
    for i, ((ph, (a, b, c)), d) in enumerate(zip(plan, days)):
        for k, per in enumerate(("2교시", "5교시")):
            jitter = lambda x: max(0, x + rnd.choice((-1, 0, 0, 1)) if x else x)
            o, t_, x = jitter(a), jitter(b), jitter(c)
            _setdate(ws, f"A{r}", d)
            ws[f"C{r}"].value = ph
            ws[f"D{r}"].value = per
            ws[f"E{r}"].value = subj[(i + k) % len(subj)]
            ws[f"F{r}"].value, ws[f"G{r}"].value, ws[f"H{r}"].value = o, t_, x
            ws[f"I{r}"].value = "Y"
            if ph == "기초선" and i == 3 and k == 1:
                ws[f"I{r}"].value = "N"
                ws[f"J{r}"].value = "있음"
                ws[f"K{r}"].value = "소리 지르며 교실 이탈 — 안전 절차"
            r += 1


FILL = {"park": fill_park, "kang": fill_kang, "lim": fill_lim, "kim": fill_kim}


def fill_all(wb, cfg):
    FILL[cfg["key"]](wb, cfg)
    form7_from_source(wb, cfg)
