# -*- coding: utf-8 -*-
"""CI：检查所有 Python 脚本语法。"""
import ast
import glob
import sys

bad = 0
for f in glob.glob("skills/k3-bom-grid-entry/scripts/*.py"):
    try:
        ast.parse(open(f, encoding="utf-8").read())
        print("ok  ", f)
    except SyntaxError as e:
        print("SYNTAX", f, e)
        bad = 1
sys.exit(bad)
