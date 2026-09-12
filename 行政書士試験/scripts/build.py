#!/usr/bin/env python3
"""行政書士試験（令和9年度）向け 日割り問題集ビルダー

data/*.json の問題データを読み込み、
  - daily/day-XXX.md  … 1日分（5問・10〜20分）
  - daily/INDEX.md    … 学習カレンダー
  - questions_all.json … 全問題を1ファイルに統合（Webアプリ用）
を生成する。
"""
import json
import glob
import os
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
DAILY = os.path.join(ROOT, "daily")

# 学習ロードマップ順（ファイル名の接頭番号順）
DRILL_FILES = [
    "01_基礎法学.json", "02_憲法.json", "03_行政法総論.json", "04_行政手続法.json",
    "05_行政不服審査法.json", "06_行政事件訴訟法.json", "07_国家賠償法.json",
    "08_地方自治法.json", "09_民法総則物権.json", "10_民法債権.json",
    "11_民法親族相続.json", "12_商法会社法.json", "13_行政書士法.json",
    "14_情報個人情報.json", "15_一般知識.json", "16_文章理解.json",
]
DESC_FILE = "17_記述式.json"
MULTI_FILE = "18_多肢選択式.json"

PER_DAY = 5          # 1日の問題数
REVIEW_EVERY = 7     # 7日目ごとに週末チェック


def load(name):
    with open(os.path.join(DATA, name), encoding="utf-8") as f:
        return json.load(f)


def build_pool():
    """分野順に並べた出題プールと、全問題の辞書を返す。"""
    pool, index = [], {}
    for name in DRILL_FILES:
        d = load(name)
        for q in d["questions"]:
            q["field"] = d["field"]
            pool.append(q)
            index[q["id"]] = q
    return pool, index


def build_days():
    pool, index = build_pool()
    desc = load(DESC_FILE)["questions"]
    multi = load(MULTI_FILE)["questions"]
    for q in desc:
        q["field"] = "記述式"
        index[q["id"]] = q
    for q in multi:
        q["field"] = "多肢選択式"
        index[q["id"]] = q

    days, cursor, drill_no = [], 0, 0
    di, mi = 0, 0
    recent = []  # 直近の演習問題ID（週末チェックの復習用）

    while cursor < len(pool):
        day_no = len(days) + 1
        if day_no % REVIEW_EVERY == 0:
            # 週末チェック：記述1・多肢1・直近からの復習3問
            items = []
            if di < len(desc):
                items.append(desc[di]); di += 1
            if mi < len(multi):
                items.append(multi[mi]); mi += 1
            step = max(1, len(recent) // 3) if recent else 1
            items += [index[i] for i in recent[::step][:3]]
            days.append({"no": day_no, "kind": "review", "title": "週末チェック",
                         "items": items})
            recent = []
        else:
            chunk = pool[cursor:cursor + PER_DAY]
            cursor += PER_DAY
            drill_no += 1
            fields = []
            for q in chunk:
                if q["field"] not in fields:
                    fields.append(q["field"])
            days.append({"no": day_no, "kind": "drill",
                         "title": " / ".join(fields), "items": chunk})
            recent += [q["id"] for q in chunk]

    # 余った記述式・多肢選択式は最終の総仕上げ日にまとめる
    leftovers = desc[di:] + multi[mi:]
    while leftovers:
        chunk, leftovers = leftovers[:5], leftovers[5:]
        days.append({"no": len(days) + 1, "kind": "review",
                     "title": "記述式・多肢選択式 総仕上げ", "items": chunk})
    return days, index


def fmt_question(n, q, show_answer=False):
    lines = []
    if q["type"] == "ox":
        lines.append(f"**問{n}**（{q['field']}・{q['topic']}）〇か×か\n")
        lines.append(f"> {q['q']}\n")
    elif q["type"] == "mc":
        lines.append(f"**問{n}**（{q['field']}・{q['topic']}）\n")
        lines.append(f"> {q['q']}\n")
        for i, c in enumerate(q["choices"]):
            lines.append(f"> {i + 1}. {c}")
        lines.append("")
    elif q["type"] == "desc":
        lines.append(f"**問{n}**（記述式・{q['topic']}）40字程度で記述\n")
        lines.append(f"> {q['q']}\n")
    elif q["type"] == "multi":
        lines.append(f"**問{n}**（多肢選択式・{q['topic']}）\n")
        lines.append("> " + q["q"].replace("\n", "\n> ") + "\n")
        opts = "　".join(f"{i + 1}. {c}" for i, c in enumerate(q["choices"]))
        lines.append(f"> 選択肢：{opts}\n")
    return "\n".join(lines)


def fmt_answer(n, q):
    lines = [f"**問{n}　{q['id']}**\n"]
    if q["type"] == "ox":
        lines.append(f"答：**{'〇' if q['a'] else '×'}**\n")
    elif q["type"] == "mc":
        lines.append(f"答：**{q['a'] + 1}**\n")
    elif q["type"] == "desc":
        lines.append(f"解答例：**{q['model']}**（{len(q['model'])}字）\n")
    elif q["type"] == "multi":
        labels = q.get("blanks", ["ア", "イ", "ウ", "エ"])
        ans = "　".join(f"{lb}＝{q['choices'][i]}（{i + 1}）"
                        for lb, i in zip(labels, q["a"]))
        lines.append(f"答：**{ans}**\n")
    lines.append(f"{q['exp']}\n")
    if q.get("ref"):
        lines.append(f"根拠：{q['ref']}\n")
    return "\n".join(lines)


def write_days(days):
    os.makedirs(DAILY, exist_ok=True)
    for old in glob.glob(os.path.join(DAILY, "*.md")):
        os.remove(old)
    for d in days:
        n = d["no"]
        est = "15〜20分" if d["kind"] == "review" else "10〜15分"
        out = [f"# Day {n:03d}　{d['title']}", "",
               f"目安時間：{est}　／　{len(d['items'])}問", "",
               "解答は各問の下の「解答・解説を見る」を開くと確認できます。"
               "先に全問解いてから開くこと。", "", "---", ""]
        for i, q in enumerate(d["items"], 1):
            out.append(fmt_question(i, q))
            out.append("<details><summary>解答・解説を見る</summary>\n")
            out.append(fmt_answer(i, q))
            out.append("</details>\n")
            out.append("---\n")
        prev = f"[← Day {n-1:03d}](day-{n-1:03d}.md)　" if n > 1 else ""
        nxt = f"[Day {n+1:03d} →](day-{n+1:03d}.md)" if n < len(days) else ""
        out.append(f"{prev}[目次](INDEX.md)　{nxt}")
        with open(os.path.join(DAILY, f"day-{n:03d}.md"), "w",
                  encoding="utf-8") as f:
            f.write("\n".join(out) + "\n")


def write_index(days):
    out = ["# 学習カレンダー（1周分）", "",
           f"全{len(days)}日。1日10〜20分。7日目ごとに週末チェック"
           "（記述式・多肢選択式＋直近の復習）が入ります。", "",
           "| Day | 分野 | 問数 | 区分 |", "| --- | --- | --- | --- |"]
    for d in days:
        kind = "週末チェック" if d["kind"] == "review" else "演習"
        out.append(f"| [Day {d['no']:03d}](day-{d['no']:03d}.md) | "
                   f"{d['title']} | {len(d['items'])} | {kind} |")
    with open(os.path.join(DAILY, "INDEX.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")


def write_bundle(index, days):
    bundle = {
        "generated": date.today().isoformat(),
        "exam": "令和9年度 行政書士試験",
        "days": [{"no": d["no"], "kind": d["kind"], "title": d["title"],
                  "ids": [q["id"] for q in d["items"]]} for d in days],
        "questions": list(index.values()),
    }
    with open(os.path.join(ROOT, "questions_all.json"), "w",
              encoding="utf-8") as f:
        json.dump(bundle, f, ensure_ascii=False, indent=1)


def main():
    days, index = build_days()
    write_days(days)
    write_index(days)
    write_bundle(index, days)
    drills = sum(1 for d in days if d["kind"] == "drill")
    print(f"生成: {len(days)}日分（演習{drills}日／週末チェック"
          f"{len(days) - drills}日）、全{len(index)}問")


if __name__ == "__main__":
    main()
