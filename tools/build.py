# -*- coding: utf-8 -*-
"""사용: python3 build.py park|kang|lim|kim [출력경로]"""
import sys, os, importlib, openpyxl
sys.path.insert(0, os.path.dirname(__file__))
import records, sheets, summary, forms
from lib import *

SRC = os.environ.get("SRC_DIR", os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "source"))


def build(key, out, variant='final'):
    cfg = importlib.import_module(f"cfg_{key}").make()
    wb = openpyxl.load_workbook(os.path.join(SRC, cfg["src"]))
    for name in list(wb.sheetnames):
        if name not in forms.KEEP_FORMS:
            del wb[name]
    kind = cfg["rec"]["kind"]
    # 정의를 먼저(문항 목록 행 번호가 기록지 머리 참조에 필요)
    sheets.build_definition(wb, cfg)
    if kind == "seq":
        records.build_record_seq(wb, cfg)
        records.build_print_seq(wb, cfg)
    elif kind == "days":
        records.build_record_days(wb, cfg)
        records.build_wall_days(wb, cfg)
    else:
        records.build_record_counts(wb, cfg)
        records.build_wall_counts(wb, cfg)
    cfg["summary"]["rec_rows"] = cfg["rec"]["_rows"]
    summary.build_summary(wb, cfg)
    sheets.build_design(wb, cfg)
    forms.patch_form1_3(wb, cfg)
    cfg['_stripped'] = forms.strip_predictions(wb)
    cfg["_fba_row"] = forms.patch_form5(wb, cfg)
    forms.patch_form6(wb, cfg)
    forms.patch_form7(wb, cfg)
    forms.clear_hyperlink_formulas(wb)
    sheets.build_progress(wb, cfg)
    forms.finalize(wb, cfg, out, cfg["order"])
    return cfg


def make_clone(key, out_final, out_clone):
    """설명용(내부) 사본 = 최종본 + 예방 예측 자료 복원 + 예시(가상) 자료 + 설명 스크립트 탭."""
    import scripts, examples
    cfg = build(key, out_final)
    wb = openpyxl.load_workbook(out_final)
    forms.restore_predictions(wb, cfg["_stripped"])
    examples.fill_all(wb, cfg)
    scripts.build_script_sheet(wb, cfg)
    wb.active = 0
    for ws in wb:
        ws.sheet_view.tabSelected = (ws.title == wb.sheetnames[0])
    wb.calculation.fullCalcOnLoad = True
    forms.tab_colors(wb)
    wb.properties.creator = ''
    wb.save(out_clone)
    return cfg


if __name__ == "__main__":
    key = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else f"/tmp/{key}.xlsx"
    build(key, out)
    print("saved", out)
