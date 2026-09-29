# -*- coding: utf-8 -*-
"""0.진행 지도 · 1.정의 · 4.중재 설계 공통 생성기."""
from lib import *


# ───────────────────────── 0.진행 지도 ─────────────────────────
def build_progress(wb, cfg):
    ws = wb.create_sheet("0.진행 지도")
    hide_gridlines(ws)
    colw(ws, {"A": 6, "B": 36, "C": 24, "D": 16, "E": 22, "F": 11, "G": 30})
    title(ws, f'0.진행 지도 — 7분임 {cfg["teacher"]} 선생님 · {cfg["student"]}', 7)
    legend_row(ws, 2, [(1, 2), (3, 4), (5, 7)])
    merged(ws, 3, 1, 7, cfg["progress_intro"], color=BLUE, bold=True)
    heads = ["단계", "할 일", "탭", "누가", "언제", "상태", "자동 확인"]
    for i, h in enumerate(heads):
        put(ws, f"{CL(1+i)}4", h, bold=True, bg=SECT, h="center", border=True)
    steps = cfg["steps"]
    for i, st in enumerate(steps):
        r = 5 + i
        yel = YELLOW if st.get("y") else None
        put(ws, f"A{r}", i + 1, h="center", border=True, bg=yel)
        put(ws, f"B{r}", st["task"], border=True, bg=yel)
        put(ws, f"C{r}", st["tab"], border=True, bg=yel, size=9)
        put(ws, f"D{r}", st["who"], border=True, bg=yel, size=9, h="center")
        put(ws, f"E{r}", st["when"], border=True, bg=yel, size=9, color=BLUE if st.get("y") else "000000")
        put(ws, f"F{r}", "미시작", border=True, bg=yel, h="center", size=9)
        auto = st.get("auto")
        auto = auto(cfg) if callable(auto) else auto
        put(ws, f"G{r}", auto, border=True, bg=AUTO, size=9, color=GRAY_TXT)
        fit_row(ws, r, st["task"], 36)
        fit_row(ws, r, st["when"], 22, size=9)
    last = 4 + len(steps)
    dv_list(ws, f"F5:F{last}", ["미시작", "진행 중", "완료"])
    # 인쇄 안내
    r = last + 2
    merged(ws, r, 1, 7, "■ 인쇄해서 쓰는 종이 양식 — 인쇄법: 해당 탭을 열고 Ctrl+P → 용지 A4 · 가로 · 배율 '너비 맞춤' → 인쇄",
           bold=True, bg=SECT, color=BLUE)
    r += 1
    for i, h in enumerate(["", "탭", "", "쓰임"]):
        pass
    put(ws, f"B{r}", "탭", bold=True, bg=LABEL, border=True)
    ws.merge_cells(f"C{r}:G{r}")
    put(ws, f"C{r}", "쓰임", bold=True, bg=LABEL, border=True)
    for p in cfg["print_list"]:
        r += 1
        put(ws, f"B{r}", p[0], border=True)
        ws.merge_cells(f"C{r}:G{r}")
        put(ws, f"C{r}", p[1], border=True)
        fit_row(ws, r, p[1], span_width(ws, 3, 7))
    # 두 기록의 차이
    if cfg.get("progress_notes"):
        r += 2
        merged(ws, r, 1, 7, "■ 헷갈리기 쉬운 것 — 미리 정리", bold=True, bg=SECT)
        for t in cfg["progress_notes"]:
            r += 1
            merged(ws, r, 1, 7, t, border=False)
    ws.freeze_panes = "A5"
    return ws


# ───────────────────────── 1.정의 ─────────────────────────
def build_definition(wb, cfg):
    D = cfg["def"]
    ws = wb.create_sheet("1.정의")
    hide_gridlines(ws)
    colw(ws, {"A": 20, "B": 112})
    title(ws, f'표적행동 정의·측정 통합시트 — 7분임 {cfg["teacher"]} 선생님', 2)

    def kv(r, label, text, yellow=False, blue=False, bold=False, size=10):
        put(ws, f"A{r}", label, bold=True, bg=LABEL, border=True, size=size)
        put(ws, f"B{r}", text, bg=YELLOW if yellow else None, color=BLUE if blue else "000000", bold=bold,
            border=True, size=size)
        if isinstance(text, str) and not text.startswith("="):
            fit_row(ws, r, text, 112, size=size)

    def sect(r, text, bg=SECT):
        merged(ws, r, 1, 2, text, bold=True, bg=bg)

    kv(2, "학생", cfg["student"])
    kv(3, "한눈에", D["glance"], blue=True, bold=True)
    kv(4, "사용법", D["howto"])
    sect(6, "■ 감소 표적행동", "FDE9E7")
    kv(7, "행동 이름", D["dec"]["name"], yellow=True)
    kv(8, "조작적 정의", D["dec"]["defn"], yellow=True)
    kv(9, "비예시/제외기준", D["dec"]["excl"], yellow=True)
    for i in range(4):                                         # 10~13
        r = 10 + i
        if i < len(D["dec"]["rules"]):
            lab, txt = D["dec"]["rules"][i]
            kv(r, lab, txt, yellow=True)
    sect(14, "■ 증가 표적행동 (대체행동)", "E8F4EA")
    kv(15, "행동 이름", D["inc"]["name"], yellow=True)
    kv(16, "정의", D["inc"]["defn"], yellow=True)
    kv(17, "참고", D["inc"]["note"])
    r = 19
    if D.get("list"):
        sect(r, D["list"]["title"])
        for i, lab in enumerate(D["list"]["labels"]):
            r += 1
            kv(r, lab, D["list"].get("prefill", [""] * 99)[i], yellow=True)
            cfg.setdefault("_listrows", []).append(r)
        r += 2
    sect(r, "■ 기록 코드 — 2.기록지에서 고르는 값과 똑같습니다")
    r += 1
    put(ws, f"A{r}", "코드", bold=True, bg=LABEL, border=True)
    put(ws, f"B{r}", "뜻", bold=True, bg=LABEL, border=True)
    for grp, items in D["codes"]:
        r += 1
        put(ws, f"A{r}", grp, bold=True, color=NAVY, border=True)
        put(ws, f"B{r}", "", border=True)
        for code, mean in items:
            r += 1
            put(ws, f"A{r}", code, border=True, h="center", bold=True)
            put(ws, f"B{r}", mean, border=True)
            fit_row(ws, r, mean, 112)
    r += 2
    sect(r, "■ 기록할 때 지킬 것 (기초선 포함)")
    for i, item in enumerate(D["keep"]):
        r += 1
        txt, isblue = item if isinstance(item, tuple) else (item, False)
        put(ws, f"A{r}", i + 1, h="center", bg=LABEL, border=True)
        put(ws, f"B{r}", txt, color=BLUE if isblue else "000000", bold=isblue, border=True)
        fit_row(ws, r, txt, 112)
    r += 2
    sect(r, "■ 판정 예시")
    for lab, txt in D["examples"]:
        r += 1
        kv(r, lab, txt)
    if D.get("extra_sect"):
        r += 2
        sect(r, D["extra_sect"]["title"])
        for i, txt in enumerate(D["extra_sect"]["items"]):
            r += 1
            put(ws, f"A{r}", i + 1, h="center", bg=LABEL, border=True)
            put(ws, f"B{r}", txt, border=True)
            fit_row(ws, r, txt, 112)
    r += 2
    sect(r, "■ 왜 이 행동인가", "F7F7F7")
    r += 1
    put(ws, f"A{r}", "선정 이유", bold=True, bg=LABEL, border=True)
    put(ws, f"B{r}", D["why"], border=True)
    fit_row(ws, r, D["why"], 112)
    ws.freeze_panes = "A2"
    return ws


# ───────────────────────── 4.중재 설계(초안) ─────────────────────────
def build_design(wb, cfg):
    G = cfg["design"]
    ws = wb.create_sheet("4.중재 설계(초안)")
    hide_gridlines(ws)
    colw(ws, {"A": 20, "B": 44, "C": 70, "D": 40})
    title(ws, '="4.중재 설계(초안) — "&\'1.정의\'!$B$7', 4)
    merged(ws, 2, 1, 4, "기능과 구체 중재 방법은 기능적행동평가를 위해 모은 자료로 4회기에 확정합니다. 중재는 기초선이 끝난 뒤 시작합니다.",
           color=GRAY_TXT)
    merged(ws, 3, 1, 4, "기초선 원칙 — 중재 요소(새 강화 체계·점검 카드·새 반응 규칙·새 촉구)는 기초선에 넣지 않습니다. "
           "기초선은 '좋게 만드는 기간'이 아니라 '같게 유지하는 기간'이므로, 이미 하던 지원은 거두지 않고 그 조건을 메모합니다.",
           color=BLUE, bold=True)
    r = 4

    def block(head, rows, twocol=True):
        nonlocal r
        r += 1
        merged(ws, r, 1, 4, head, bold=True, bg=SECT)
        for lab, txt in rows:
            r += 1
            put(ws, f"A{r}", lab, bold=True, bg=LABEL, border=True)
            ws.merge_cells(f"B{r}:D{r}")
            put(ws, f"B{r}", txt, border=True)
            for cc in "CD":
                ws[f"{cc}{r}"].border = BORDER
            fit_row(ws, r, txt, 154)

    block("■ 지금 할 수 있는 일 (기초선 동안)", G["now"])
    block("■ 중재의 일반적 틀", G["frame"])
    block("■ 강화와 기록을 겸하는 설계(초안)", G["token"])
    block("■ 기초선에서 모아 중재 계획에 반영할 것", G["collect"])
    block("■ 실행·기록 가능성", G["feasible"])
    r += 2
    merged(ws, r, 1, 4, "■ 4회기 확정 항목 — 오른쪽 노란 칸은 4회기에 함께 채웁니다", bold=True, bg=SECT)
    r += 1
    for i, h in enumerate(["확정 항목", "4회기에 가져올 정보", "초안(4회기 전 슈퍼바이저 작성)", "4회기 합의·수정"]):
        put(ws, f"{CL(1+i)}{r}", h, bold=True, bg=LABEL, border=True, h="center")
    for lab, info, draft in G["confirm"]:
        r += 1
        put(ws, f"A{r}", lab, bold=True, bg=LABEL, border=True)
        put(ws, f"B{r}", info, border=True, size=9)
        put(ws, f"C{r}", draft, border=True, size=9, color=EX_TXT if False else "000000", bg="FFFBE6")
        put(ws, f"D{r}", None, border=True, bg=YELLOW)
        h = max(lines_needed(info, 44), lines_needed(draft, 70)) * 13 + 6
        ws.row_dimensions[r].height = h
    ws.freeze_panes = "A2"
    return ws
