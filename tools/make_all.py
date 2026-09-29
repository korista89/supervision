# -*- coding: utf-8 -*-
"""배포용 최종본 4개 + 설명용(내부) 사본 4개를 한 번에 만든다.  사용: python3 tools/make_all.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NAMES = {"park": "박윤아", "kang": "강희강", "lim": "임충만", "kim": "김규리"}
os.makedirs(os.path.join(ROOT, "final"), exist_ok=True)
os.makedirs(os.path.join(ROOT, "internal"), exist_ok=True)
for key, name in NAMES.items():
    build.make_clone(key,
                     os.path.join(ROOT, "final", f"7분임_{name}T_행동중재전문가_슈퍼비전_양식.xlsx"),
                     os.path.join(ROOT, "internal", f"7분임_{name}T_슈퍼비전_설명용(내부).xlsx"))
    print("완료:", name)
