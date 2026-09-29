# -*- coding: utf-8 -*-
"""빌드 산출물 검증: (1) 예시 자료를 넣은 사본을 LibreOffice로 재계산 → 오류 셀·핵심 수치 확인 (2) 옛 용어 잔존 검사."""
import sys, os, subprocess, shutil, re, datetime, openpyxl
import tempfile
WORK = os.path.join(tempfile.gettempdir(), "supervision_work")
os.makedirs(WORK, exist_ok=True)

STALE = ["정정", "칸당", "부딪히기", "5분", "10초 규칙", "IHTX", "T1234", "VMTW", "㉮", "㉯", "㉰", "최고 단계",
         "상해 코드", "목록 탭", "Apps Script", "백업", "예시 자료 위치", "가상", "HYPERLINK", "회색 글씨 = 예시(가상)", "자동 반영"]


def recalc(path):
    out = os.path.join(WORK, "recalc")
    os.makedirs(out, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, path], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    return os.path.join(out, os.path.basename(path))


def scan_errors(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    bad = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and (v.startswith("#") or v.startswith("Err:")) and v != "#N/A":
                    bad.append((ws.title, c.coordinate, v))
    return bad


def stale_scan(path):
    wb = openpyxl.load_workbook(path)
    hits = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    for t in STALE:
                        if t in c.value:
                            hits.append((ws.title, c.coordinate, t))
    return hits


def color_scan(path):
    """빨간 글씨·하늘색(00FFFF) 칸이 남았는지."""
    wb = openpyxl.load_workbook(path)
    out = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                col = c.font.color.rgb if c.font and c.font.color and c.font.color.type == "rgb" else None
                if col and str(col)[-6:].upper() in ("FF0000", "C00000", "CC0000", "EA4335"):
                    out.append((ws.title, c.coordinate, "red"))
                if c.fill and c.fill.fill_type == "solid" and str(c.fill.fgColor.rgb)[-6:].upper() == "00FFFF":
                    out.append((ws.title, c.coordinate, "cyan"))
    return out
