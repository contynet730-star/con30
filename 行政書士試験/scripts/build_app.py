#!/usr/bin/env python3
"""app/template.html に問題データを埋め込み、公開用の1ファイルHTMLを書き出す。"""
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL = os.path.join(ROOT, "app", "template.html")
OUT = os.path.join(ROOT, "app", "毎日一問一答.html")
DATA = os.path.join(ROOT, "questions_all.json")

with open(DATA, encoding="utf-8") as f:
    data = json.load(f)

# ページに埋め込む分は最小限に絞る（daily と questions のみ）
slim = {
    "days": data["days"],
    "questions": data["questions"],
}
payload = json.dumps(slim, ensure_ascii=False, separators=(",", ":"))
# </script> でスクリプトブロックが閉じないようにする
payload = payload.replace("</", "<\\/")

with open(TPL, encoding="utf-8") as f:
    html = f.read()

if "/*__DATA__*/" not in html:
    raise SystemExit("template.html に /*__DATA__*/ が見つかりません")

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html.replace("/*__DATA__*/", payload))

print(f"生成: {OUT}  ({os.path.getsize(OUT):,} bytes、{len(slim['questions'])}問)")
