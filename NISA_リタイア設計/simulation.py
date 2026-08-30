# -*- coding: utf-8 -*-
"""
60歳リタイア逆算プラン — 試算スクリプト

前提:
  45歳 / 金融資産 1,500万円 / 毎月の積立 5万円 / 60歳で退職
  生活防衛資金 300万円は現金で保持し、残り 1,200万円を投資に回す。
  金額の単位はすべて「万円」。名目リターン（インフレ未考慮）。
実行: python3 simulation.py
"""

CASH   = 300    # 生活防衛資金（投資しない）
LUMP   = 1200   # 既存資金のうち投資に回す分
YEARS  = 15     # 45 -> 60歳
TAX    = 0.20315  # 課税口座の譲渡益課税
INCOME_TAX_RATE = 0.20  # iDeCo所得控除の想定税率（所得税10%+住民税10%）


def fv_lump(rate, annual=240, n=5, years=YEARS):
    """年 annual 万円を n 年間、年初に投入して years 年運用した将来価値。"""
    bal = 0.0
    for y in range(years):
        bal = (bal + (annual if y < n else 0)) * (1 + rate)
    return bal


def fv_monthly(rate, monthly, months=YEARS * 12):
    """毎月 monthly 万円を積み立てた将来価値（月複利）。"""
    r, acc = rate / 12, 0.0
    for _ in range(months):
        acc = (acc + monthly) * (1 + r)
    return acc


def life(start, spend, pension, ret, pension_age=65, work=0, work_to=65, cap=105):
    """60歳から取り崩しを開始し、資産が尽きる年齢を返す（cap 以上なら枯渇しない）。"""
    bal, age = start, 60
    while age < cap:
        income = (pension if age >= pension_age else 0) + (work if age < work_to else 0)
        bal = (bal - (spend - income)) * (1 + ret)
        if bal <= 0:
            return age
        age += 1
    return cap


def rule(title):
    print("\n" + "=" * 72 + f"\n{title}\n" + "=" * 72)


# ── 1. NISA生涯枠の消化スケジュール ────────────────────────────
rule("1. NISA枠の埋め方（成長枠 240万×5年 + つみたて枠 36万×15年）")
cum = 0
for y in range(YEARS):
    growth = 240 if y < 5 else 0
    cum += growth + 36
    if y < 5 or y == YEARS - 1:
        print(f"  {45+y}歳  成長枠 {growth:3d}万 + つみたて枠 36万 → NISA累計 {cum:5,d}万")
print(f"\n  15年間の投入計 {cum:,}万円 / 生涯枠 1,800万円（残 {1800-cum}万円）")
print(f"  成長投資枠は上限 1,200万円ちょうど。年間投資枠 360万円も超えない。")

# ── 2. 60歳時点の資産 ─────────────────────────────────────
rule("2. 60歳時点の金融資産（iDeCo 2万 + NISAつみたて 3万 に配分）")
totals = {}
for r in (0.03, 0.05, 0.07):
    g, n, i = fv_lump(r), fv_monthly(r, 3), fv_monthly(r, 2)
    totals[r] = g + n + i + CASH
    print(f"  年率{r*100:.0f}%: 成長枠 {g:6,.0f} / つみたて {n:5,.0f} / iDeCo {i:5,.0f}"
          f" / 現金 {CASH} = {totals[r]:6,.0f}万円")
print(f"\n  拠出合計 {CASH + LUMP + 5*12*YEARS:,}万円（うち投資 {LUMP + 5*12*YEARS:,}万円）")
print(f"  iDeCo所得控除による節税累計: 約 {2*12*YEARS*INCOME_TAX_RATE:.0f}万円")

# ── 3. NISA単独 vs iDeCo併用 ──────────────────────────────
rule("3. 月5万円の配分比較（60歳時点の税引後・単位 万円）")
for r in (0.03, 0.05, 0.07):
    growth = fv_lump(r)
    # (a) 全額NISA: 枠は54歳で満杯 → 55〜59歳の月5万は課税口座
    a_nisa = fv_monthly(r, 5, 10 * 12) * (1 + r) ** 5
    a_taxed = fv_monthly(r, 5, 5 * 12)
    a_net = a_taxed - (a_taxed - 300) * TAX
    a = growth + a_nisa + a_net
    # (b) iDeCo 2万 + NISA 3万
    b = growth + fv_monthly(r, 3) + fv_monthly(r, 2)
    saved = 2 * 12 * YEARS * INCOME_TAX_RATE
    print(f"  年率{r*100:.0f}%:  全額NISA {a:6,.0f}  |  併用 {b:6,.0f} + 節税 {saved:.0f}"
          f"  →  併用が {b + saved - a:+5,.0f}万円 有利")

# ── 4. リタイア判定 ────────────────────────────────────────
rule("4. 60歳リタイア判定（公的年金 年180万円 / 65歳〜）")
head = " ".join(f"{r*100:.0f}%成長({totals[r]:,.0f}万)".rjust(18) for r in (0.03, 0.05, 0.07))
print(f"  {'月の生活費':<10}{head}")
for spend in (20, 22, 25, 28, 30):
    cells = []
    for r in (0.03, 0.05, 0.07):
        end = life(totals[r], spend * 12, 180, r)
        cells.append(("100歳超も安泰" if end >= 105 else f"{end}歳で枯渇").rjust(18))
    print(f"  {spend:>6}万円  " + " ".join(cells))

# ── 5. レバー効果 ─────────────────────────────────────────
rule("5. 足りない場合のレバー（3%成長 / 生活費 月25万円 / 年金 年180万円）")
base = totals[0.03]
levers = [
    ("そのまま（基準）",                dict(start=base, spend=300, pension=180, ret=0.03)),
    ("① 積立を月5万→月8万に増額",       dict(start=base + fv_monthly(0.03, 3), spend=300, pension=180, ret=0.03)),
    ("② 60〜65歳だけ月8万円ぶん働く",    dict(start=base, spend=300, pension=180, ret=0.03, work=96)),
    ("③ 生活費を月25万→22万に圧縮",     dict(start=base, spend=264, pension=180, ret=0.03)),
    ("④ 年金を70歳まで繰下げ（+42%）",  dict(start=base, spend=300, pension=180 * 1.42, ret=0.03, pension_age=70)),
    ("②+③ の合わせ技",                dict(start=base, spend=264, pension=180, ret=0.03, work=96)),
]
for label, kw in levers:
    end = life(**kw)
    print(f"  {label:<26} → " + ("100歳超も安泰" if end >= 105 else f"{end}歳で枯渇"))
