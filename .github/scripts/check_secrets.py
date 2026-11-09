# -*- coding: utf-8 -*-
"""CI: scan for credentials, internal hosts, and local private paths.

Two pattern groups:
  SECRET  -- real credential shapes (assignment of password/token, private key blocks, cloud keys)
  PRIVATE -- local info that must not appear in a public repo (absolute home paths, real emails, UNC shares)

Docs naturally mention words like "password" / "token"; they do not match the assignment
shapes, so they stay clean. To exempt a single line, append the marker ci-allow.
"""
import glob
import re
import sys

SECRET = re.compile(
    r"(passwd|password|api[_-]?key|secret|token|access[_-]?key)\s*[:=]\s*[\"'][^\"']{6,}"
    r"|BEGIN [A-Z ]*PRIVATE KEY"
    r"|gh[pousr]_[A-Za-z0-9]{20,}"
    r"|github_pat_[A-Za-z0-9_]{20,}"
    r"|AKIA[0-9A-Z]{16}",
    re.I,
)

PRIVATE = re.compile(
    r"[A-Za-z]:\\{1,2}(Users|Documents and Settings)\\{1,2}(?!<|xxx|\*|name\b|user\b)[A-Za-z0-9_.\-]{2,}"
    r"|/home/(?!user\b|xxx|\*)[a-z0-9_\-]{3,}/"
    r"|\\\\[A-Za-z0-9_\-]{2,}\\(share|d\$|c\$)"
    r"|[\w.\-]+@(?!example\.(com|org)|users\.noreply\.github\.com|email\.com)[\w\-]+\.[a-z]{2,}",
    re.I,
)

SKIP = (".git/", ".github/scripts/check_secrets.py")
ALLOW_MARK = "ci-allow"

bad = 0
checked = 0
for ext in ("*.py", "*.md", "*.json", "*.csv", "*.yml", "*.yaml", "*.txt", "*.example"):
    for f in glob.glob("**/" + ext, recursive=True):
        f = f.replace("\\", "/")
        if any(s in f for s in SKIP):
            continue
        checked += 1
        try:
            fh = open(f, encoding="utf-8", errors="replace")
        except OSError:
            continue
        for i, line in enumerate(fh, 1):
            if ALLOW_MARK in line:
                continue
            for tag, pat in (("SECRET", SECRET), ("PRIVATE", PRIVATE)):
                if pat.search(line):
                    out = line.strip()[:120].encode("ascii", "backslashreplace").decode()
                    print("%s %s:%d %s" % (tag, f, i, out))
                    bad = 1

print("scanned %d files" % checked)
print("clean" if not bad else "found suspicious lines")
sys.exit(bad)
