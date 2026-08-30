# -*- coding: utf-8 -*-
"""
60歳リタイア逆算プラン — 試算スクリプト

前提:
  夫婦とも45歳 / 年収 各900万円（世帯1,800万円・手取り約1,320万円）
  金融資産 1,500万円 / 毎月の積立 5万円 / 60歳で退職
  生活防衛資金 300万円は現金で保持し、残り1,200万円を夫婦のNISA成長投資枠へ3年で投入。
  公的年金は夫婦合計 457万円/年（65歳〜）。推計の内訳は pension.py を参照。
  金額の単位はすべて「万円」。名目リターン（インフレ未考慮）。

実行:
  python3 simulation.py     全試算
  python3 pension.py        公的年金の推計のみ
"""

CASH, LUMP, YEARS = 300, 1200, 15
PENSION  = 457      # 夫婦合計の公的年金 年額（65歳〜）
TAX_RATE = 0.30     # 年収900万円の限界税率（所得税20%＋住民税10%）
NET_INCOME = 1320   # 世帯手取り 年額
SCHED = [480, 480, 240]   # 夫婦の成長投資枠で1,200万円を3年投入（年720万円まで可）


def fv_lump(rate, sched=SCHED, years=YEARS):
    """年ごとの投入額 sched を years 年運用した将来価値。"""
    bal = 0.0
    for y in range(years):
        bal = (bal + (sched[y] if y < len(sched) else 0)) * (1 + rate)
    return bal


def fv_monthly(rate, monthly, months=YEARS * 12):
    """毎月 monthly 万円を積み立てた将来価値（月複利）。"""
    r, acc = rate / 12, 0.0
    for _ in range(months):
        acc = (acc + monthly) * (1 + r)
    return acc


def life(start, spend, pension, ret, p_age=65, cap=105):
    """60歳から取り崩しを開始し、資産が尽きる年齢を返す（cap 以上なら枯渇しない）。"""
    bal, age = start, 60
    while age < cap:
        bal = (bal - (spend - (pension if age >= p_age else 0))) * (1 + ret)
        if bal <= 0:
            return age
        age += 1
    return cap


def need(spend, pension, ret, p_age=65):
    """100歳まで持たせるために60歳時点で必要な資産（二分探索）。"""
    lo, hi = 0.0, 20000.0
    for _ in range(60):
        mid = (lo + hi) / 2
        if life(mid, spend, pension, ret, p_age) >= 105:
            hi = mid
        else:
            lo = mid
    return hi


def rule(title):
    print("\n" + "=" * 76 + f"\n{title}\n" + "=" * 76)


ASSETS = {m: {r: fv_lump(r) + fv_monthly(r, m) + CASH for r in (0.03, 0.05, 0.07)}
          for m in (5, 10, 15, 25)}

# ── 1. NISA枠 ───────────────────────────────────────────────
rule("1. NISA枠は夫婦で3,600万円。既存1,200万円は3年で投入できる")
print("  年間投資枠は1人360万円（つみたて120＋成長240）→ 夫婦で720万円/年")
cum = 0
for age, g in zip((45, 46, 47), SCHED):
    cum += g
    print(f"    {age}歳: 成長投資枠に {g}万円（夫 {g//2}万＋妻 {g//2}万） → 累計 {cum:,}万円")
used = LUMP + 5 * 12 * YEARS
print(f"  月5万円の積立を15年続けても NISA使用は {used:,}万円。生涯枠3,600万円に対し {3600-used:,}万円 余る。")
print("  → この世帯に「NISA枠が足りない」制約は存在しない。iDeCoを使う理由は節税だけになる。")

# ── 2. 60歳時点の資産 ────────────────────────────────────────
rule("2. 60歳時点の資産（既存1,200万円の投資分＋現金300万円を含む）")
print(f"  {'毎月の積立':<12}{'手取り比':>8}" + "".join(f"{str(int(r*100))+'%成長':>12}" for r in (0.03, 0.05, 0.07)))
for m in (5, 10, 15, 25):
    share = m * 12 / NET_INCOME * 100
    print(f"  月{m:2d}万円{'  ← 現状' if m == 5 else '      '}{share:6.1f}%"
          + "".join(f"{ASSETS[m][r]:11,.0f}万" for r in (0.03, 0.05, 0.07)))

# ── 3. 必要資産の逆算 ────────────────────────────────────────
rule("3. 60歳で必要な資産額（100歳まで持たせる / 年率3%運用）")
print(f"  公的年金 夫婦計 {PENSION}万円/年（月{PENSION/12:.1f}万円）が65歳から入る前提")
have = ASSETS[5][0.03]
print(f"  月5万円の積立を続けた場合の60歳資産: {have:,.0f}万円\n")
for s in (30, 35, 40, 45, 50, 55):
    n = need(s * 12, PENSION, 0.03)
    gap = have - n
    verdict = f"届く（{gap:,.0f}万円の余裕）" if gap >= 0 else f"{-gap:,.0f}万円 不足"
    print(f"  生活費 月{s}万円（年{s*12}万円） → 必要 {n:7,.0f}万円   {verdict}")
print("\n  ※月40万円と月45万円の間で必要額が約1,500万円跳ねる。年金でカバーできる水準を")
print("    超えると不足分が生涯積み上がるため。この段差の手前が実質的な合否ライン。")

# ── 4. 判定マトリクス ────────────────────────────────────────
rule("4. 判定マトリクス（年率3%・年金は65歳から夫婦計457万円）")
print(f"  {'生活費':<10}" + "".join(f"{'月'+str(m)+'万積立':>15}" for m in (5, 10, 15, 25)))
for s in (35, 40, 45, 50, 55):
    cells = []
    for m in (5, 10, 15, 25):
        e = life(ASSETS[m][0.03], s * 12, PENSION, 0.03)
        cells.append("100歳超 ◎" if e >= 105 else (f"{e}歳 △" if e >= 90 else f"{e}歳 ×"))
    print(f"  月{s}万円   " + "".join(f"{c:>15}" for c in cells))

# ── 5. 退職金の感度 ─────────────────────────────────────────
rule("5. 退職金を入れた場合（積立は月5万円のまま・年率3%）")
print(f"  {'夫婦の退職金合計':<16}{'60歳資産':>10}" + "".join(f"{'月'+str(s)+'万':>12}" for s in (45, 50, 55)))
for t in (0, 1000, 2000, 3000, 4000):
    tot = ASSETS[5][0.03] + t
    cells = []
    for s in (45, 50, 55):
        e = life(tot, s * 12, PENSION, 0.03)
        cells.append("100歳超 ◎" if e >= 105 else (f"{e}歳 △" if e >= 90 else f"{e}歳 ×"))
    print(f"  {t:>10,}万円  {tot:8,.0f}万" + "".join(f"{c:>12}" for c in cells))
print("\n  勤続38年の退職所得控除は 800万＋70万×(38−20)＝2,060万円/人。夫婦で4,120万円まで非課税。")
print("  ただしiDeCo一時金を同年に受け取ると枠を食い合う（2026年1月から重複調整が5年→10年に延長）。")

# ── 6. 繰下げ戦略 ───────────────────────────────────────────
rule("6. 繰下げは「資産が足りている」場合だけ有効（生活費 月45万円・3%成長）")
for m in (5, 10, 15):
    print(f"\n  ● 積立 月{m}万円 → 60歳資産 {ASSETS[m][0.03]:,.0f}万円")
    for lbl, p, pa in [("2人とも65歳から受給         ", 457, 65),
                       ("1人だけ70歳まで繰下げ       ", 553, 70),
                       ("2人とも70歳まで繰下げ(+42%) ", 649, 70)]:
        e = life(ASSETS[m][0.03], 45 * 12, p, 0.03, pa)
        print(f"    {lbl}→ " + ("100歳超も安泰 ◎" if e >= 105 else f"{e}歳で枯渇"))
print("\n  ※繰下げは65〜70歳の穴を資産で埋められて初めて機能する。資産不足だと逆効果。")

# ── 7. iDeCoの節税 ──────────────────────────────────────────
rule("7. iDeCoの節税効果（限界税率30%・夫婦それぞれが拠出）")
for lbl, m in [("現行の上限 月2.3万円×2人", 2.3), ("2026年12月〜 月6.2万円×2人", 6.2)]:
    y = m * 12 * 2
    print(f"  {lbl:26s}: 年間拠出 {y:5.1f}万円 → 節税 年 {y*TAX_RATE:5.1f}万円 / 15年で {y*TAX_RATE*15:6.0f}万円")
saved_m = 2.3 * 12 * 2 * TAX_RATE / 12
r = 0.03
print(f"\n  節税分 月{saved_m:.1f}万円も再投資した場合の60歳資産（3%）: "
      f"{fv_lump(r) + fv_monthly(r, 5) + fv_monthly(r, saved_m) + CASH:,.0f}万円"
      f"（+{fv_monthly(r, saved_m):,.0f}万円）")
print("  出口は60〜64歳の年金形式受給が有利。公的年金等控除（65歳未満 年60万円）が丸ごと使え、")
print("  退職金の退職所得控除と競合せず、年金が出ない5年間の穴埋めにもなる。")
