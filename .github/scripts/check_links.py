# -*- coding: utf-8 -*-
"""CI：检查 Markdown 相对链接是否有效。"""
import glob
import os
import re
import sys

bad = 0
for f in glob.glob("**/*.md", recursive=True):
    s = open(f, encoding="utf-8").read()
    for m in re.finditer(r"\]\((?!https?://|mailto:|#)([^)#]+)", s):
        target = m.group(1).strip()
        if not target:
            continue
        p = os.path.normpath(os.path.join(os.path.dirname(f), target))
        if not os.path.exists(p):
            print("BROKEN %s -> %s" % (f, target))
            bad = 1
print("links ok" if not bad else "broken links found")
sys.exit(bad)
