#!/usr/bin/env python3
"""data/*.json を1つのバンドルにまとめる。"""
import json, glob, os
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")

qs = []
for path in sorted(glob.glob(os.path.join(DATA, "*.json"))):
    d = json.load(open(path, encoding="utf-8"))
    for q in d["questions"]:
        q["field"] = d["field"]
        qs.append(q)

by_type = {}
for q in qs:
    by_type[q["type"]] = by_type.get(q["type"], 0) + 1

bundle = {"generated": date.today().isoformat(), "exam": "令和9年度 行政書士試験",
          "questions": qs}
with open(os.path.join(ROOT, "questions_all.json"), "w", encoding="utf-8") as f:
    json.dump(bundle, f, ensure_ascii=False, indent=1)

print(f"全{len(qs)}問")
for t, n in sorted(by_type.items(), key=lambda x: -x[1]):
    print(f"  {t:6s} {n:4d}")
