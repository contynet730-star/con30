"""
体育・保健体育の方向性 指導講評 PPTX 生成スクリプト v3
- 第10回WG資料（令和8年6月5日「単元構想に向けた『統合的な理解』『総合的な発揮』の活用イメージ」）を反映
- 体育館対応：本文 22-32pt、見出し 30-34pt
- 体育科特化／文科省調査官の視点
- スライド10枚、各テキストボックス寸法計算済み
"""
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ── カラーパレット ─────────────────────────────────
NAVY      = RGBColor(0x1F, 0x3A, 0x5F)
BLUE      = RGBColor(0x2E, 0x6F, 0xB8)
LIGHTBLUE = RGBColor(0xEC, 0xF3, 0xFA)
ORANGE    = RGBColor(0xE8, 0x7A, 0x1E)
LIGHTORG  = RGBColor(0xFF, 0xF2, 0xE0)
GREEN     = RGBColor(0x2E, 0x8B, 0x57)
LIGHTGRN  = RGBColor(0xE6, 0xF4, 0xEB)
RED       = RGBColor(0xC0, 0x39, 0x2B)
LIGHTRED  = RGBColor(0xFC, 0xEC, 0xEA)
PURPLE    = RGBColor(0x6B, 0x46, 0x9C)
LIGHTPURPL= RGBColor(0xF1, 0xEB, 0xF7)
GRAY_TXT  = RGBColor(0x2B, 0x2B, 0x2B)
GRAY_SUB  = RGBColor(0x55, 0x55, 0x55)
GRAY_LINE = RGBColor(0xBB, 0xBB, 0xBB)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

FONT_JP = "メイリオ"

prs = Presentation()
prs.slide_width  = Cm(33.867)
prs.slide_height = Cm(19.05)
SW = prs.slide_width
SH = prs.slide_height
TOTAL = 10
blank = prs.slide_layouts[6]


def _set_font(run, size, bold, color, name=FONT_JP):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("a:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("a:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set("eastAsia", name)
    rFonts.set("ascii", name)
    rFonts.set("hAnsi", name)


def add_text(slide, left, top, width, height, text, size=24, bold=False,
             color=GRAY_TXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.25):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Cm(0.12)
    tf.margin_right = Cm(0.12)
    tf.margin_top = Cm(0.05)
    tf.margin_bottom = Cm(0.05)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.auto_size = MSO_AUTO_SIZE.NONE
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        _set_font(run, size=size, bold=bold, color=color)
    return tb


def add_rect(slide, left, top, width, height, fill_color,
             line_color=None, line_width=0.75):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(line_width)
    shp.shadow.inherit = False
    return shp


def add_round(slide, left, top, width, height, fill_color,
              line_color=None, line_width=0.75, corner=0.06):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(line_width)
    try:
        shp.adjustments[0] = corner
    except Exception:
        pass
    shp.shadow.inherit = False
    return shp


def add_arrow(slide, left, top, width, height, color, shape=MSO_SHAPE.DOWN_ARROW):
    arr = slide.shapes.add_shape(shape, left, top, width, height)
    arr.fill.solid()
    arr.fill.fore_color.rgb = color
    arr.line.fill.background()
    arr.shadow.inherit = False
    return arr


def add_header(slide, num, title, page):
    add_rect(slide, Cm(0), Cm(0), SW, Cm(2.6), NAVY)
    add_rect(slide, Cm(0), Cm(2.6), SW, Cm(0.2), ORANGE)
    add_round(slide, Cm(0.8), Cm(0.55), Cm(2.2), Cm(1.5),
              ORANGE, corner=0.18)
    add_text(slide, Cm(0.8), Cm(0.55), Cm(2.2), Cm(1.5),
             num, size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Cm(3.4), Cm(0.55), Cm(26.5), Cm(1.5),
             title, size=28, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Cm(30.0), Cm(0.55), Cm(3.0), Cm(1.5),
             f"{page}/{TOTAL}", size=16, color=WHITE,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def add_footer(slide):
    add_rect(slide, Cm(0), SH - Cm(0.7), SW, Cm(0.7), NAVY)
    add_text(slide, Cm(0.8), SH - Cm(0.7), Cm(20), Cm(0.7),
             "練馬区立小学校 体育研究会  指導講評",
             size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Cm(13.0), SH - Cm(0.7), Cm(20), Cm(0.7),
             "出典：中央教育審議会 体育・保健体育、健康、安全WG（第6～10回）",
             size=10, color=WHITE, align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================
# Slide 1 : 表紙
# =====================================================================
s = prs.slides.add_slide(blank)
add_rect(s, Cm(0), Cm(0), SW, SH, WHITE)
# 上下装飾バー
add_rect(s, Cm(0), Cm(0), SW, Cm(2.5), NAVY)
add_rect(s, Cm(0), Cm(2.5), SW, Cm(0.25), ORANGE)
add_rect(s, Cm(0), SH - Cm(1.0), SW, Cm(1.0), NAVY)
add_rect(s, Cm(0), SH - Cm(1.15), SW, Cm(0.15), ORANGE)

# 上部見出し
add_text(s, Cm(0.8), Cm(0.55), Cm(32), Cm(1.5),
         "練馬区立小学校 体育研究会  指導講評",
         size=24, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# 中央パネル
add_round(s, Cm(2.5), Cm(4.0), Cm(28.8), Cm(10.5), LIGHTBLUE,
          line_color=BLUE, line_width=1.0, corner=0.03)

# 副題
add_text(s, Cm(2.5), Cm(4.5), Cm(28.8), Cm(1.2),
         "中教審WG 最新資料（令和8年6月5日 第10回）から",
         size=20, bold=True, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# メインタイトル
add_text(s, Cm(2.5), Cm(6.2), Cm(28.8), Cm(3.0),
         "これからの 体育科 が\n目指す 方向性",
         size=44, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.2)

# キーメッセージ
add_text(s, Cm(2.5), Cm(10.0), Cm(28.8), Cm(3.0),
         "─ 単元構想 を、\n  「高次の資質・能力」から 逆算する ─",
         size=26, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4)

# 下：本日テーマ
add_text(s, Cm(2.5), Cm(15.0), Cm(28.8), Cm(1.4),
         "本日の授業を踏まえて／次期学習指導要領の動向を共有",
         size=14, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Cm(0.8), SH - Cm(1.0), Cm(32), Cm(1.0),
         "練馬区立小学校 体育研究会  指導講評資料",
         size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================
# Slide 2 : 改訂の3つの方向性
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "1", "次期改訂を貫く 3つの方向性", 2)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "論点整理（令和7年9月）で示された 体育科 の進む方向",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

col_w = Cm(10.5)
col_h = Cm(11.6)
col_top = Cm(4.7)
gap = Cm(0.5)
left0 = Cm(0.85)

cols = [
    ("①", "Excellence", "卓越性",  NAVY,   LIGHTBLUE,
     "深い学び",
     "高次の資質・能力を\n授業の軸にする"),
    ("②", "Equity",     "公正性",  GREEN,  LIGHTGRN,
     "多様性の包摂",
     "全ての子どもが\n楽しさを味わう"),
    ("③", "Feasibility","持続性",  ORANGE, LIGHTORG,
     "実現可能性",
     "余白と創意工夫を\n大切にする"),
]

for i, (num, en, jp, color, bg, sub, body) in enumerate(cols):
    left = left0 + (col_w + gap) * i
    add_round(s, left, col_top, col_w, col_h, bg,
              line_color=color, line_width=1.8, corner=0.03)
    add_rect(s, left, col_top, col_w, Cm(2.0), color)
    add_text(s, left, col_top, col_w, Cm(2.0),
             num, size=48, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left, col_top + Cm(2.2), col_w, Cm(1.5),
             en, size=30, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left, col_top + Cm(3.8), col_w, Cm(0.9),
             f"（{jp}）", size=18, color=GRAY_SUB,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left + Cm(1.5), col_top + Cm(5.0),
             col_w - Cm(3.0), Cm(0.08), color)
    add_text(s, left, col_top + Cm(5.3), col_w, Cm(1.3),
             sub, size=26, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Cm(0.4), col_top + Cm(7.2),
             col_w - Cm(0.8), Cm(3.6),
             body, size=22, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 3 : 体育科の中心メッセージ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "2", "体育科の 中心となる 考え方", 3)

# 上：これまで
tb = Cm(3.5)
add_round(s, Cm(2.0), tb, Cm(29.8), Cm(4.2),
          LIGHTRED, line_color=RED, line_width=1.5, corner=0.05)
add_text(s, Cm(2.5), tb + Cm(0.4), Cm(28.8), Cm(1.1),
         "これまでの 課題",
         size=20, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.5), tb + Cm(1.6), Cm(28.8), Cm(2.4),
         "種目を 教える 授業",
         size=44, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_arrow(s, Cm(15.5), tb + Cm(4.6), Cm(3.0), Cm(1.8), ORANGE)

# 下：これから
af = Cm(10.8)
add_round(s, Cm(2.0), af, Cm(29.8), Cm(5.2),
          LIGHTBLUE, line_color=BLUE, line_width=2.0, corner=0.05)
add_text(s, Cm(2.5), af + Cm(0.5), Cm(28.8), Cm(1.1),
         "次期改訂の 方向性",
         size=20, bold=True, color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.5), af + Cm(1.8), Cm(28.8), Cm(3.0),
         "資質・能力を 育てる 授業へ",
         size=48, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Cm(0.8), Cm(16.7), Cm(32.3), Cm(1.3),
         "─ 種目は、資質・能力を 育てる 媒介（手段） ─",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 4 : 高次の資質・能力
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "3", "授業の軸「高次の資質・能力」", 4)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "個別の知・技を 関連付け、深い学びへ つなぐ 2つの姿",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

box_top = Cm(4.7)
box_h = Cm(11.0)
box_w = Cm(15.8)

# 左：統合的な理解
add_round(s, Cm(0.85), box_top, box_w, box_h, LIGHTBLUE,
          line_color=BLUE, line_width=2.0, corner=0.04)
add_rect(s, Cm(0.85), box_top, box_w, Cm(2.2), BLUE)
add_text(s, Cm(0.85), box_top, box_w, Cm(2.2),
         "統合的な 理解",
         size=32, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(0.85), box_top + Cm(2.4), box_w, Cm(0.9),
         "知識 及び 技能",
         size=18, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(1.4), box_top + Cm(3.7), box_w - Cm(1.1), Cm(7.0),
         "知 と 技 が\n関連付けられ\n一般化 された 姿\n\n「だから こう動く」",
         size=26, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4)

# 右：総合的な発揮
add_round(s, Cm(17.2), box_top, box_w, box_h, LIGHTGRN,
          line_color=GREEN, line_width=2.0, corner=0.04)
add_rect(s, Cm(17.2), box_top, box_w, Cm(2.2), GREEN)
add_text(s, Cm(17.2), box_top, box_w, Cm(2.2),
         "総合的な 発揮",
         size=32, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(17.2), box_top + Cm(2.4), box_w, Cm(0.9),
         "思考力・判断力・表現力 等",
         size=18, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(17.75), box_top + Cm(3.7), box_w - Cm(1.1), Cm(7.0),
         "状況に応じて 選び\n仲間と 課題解決する 姿\n\n「だから こう工夫する」",
         size=26, bold=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4)

add_text(s, Cm(0.8), Cm(16.4), Cm(32.3), Cm(1.3),
         "▶ この2つから 単元を 逆算設計する",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 5 : 単元構想の起点 ← 【最新WG資料の中核】
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "4", "単元構想 = 高次の資質・能力から 逆算する", 5)

# サブ
add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "最新WG資料（R8.6.5 第10回）が示した 設計の原理",
         size=18, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 上：到達点（ゴール）
add_round(s, Cm(2.0), Cm(4.6), Cm(29.8), Cm(3.5),
          LIGHTBLUE, line_color=BLUE, line_width=2.0, corner=0.04)
add_text(s, Cm(2.5), Cm(4.7), Cm(28.8), Cm(0.9),
         "①  単元の 到達点 を確認する",
         size=18, bold=True, color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.5), Cm(5.7), Cm(28.8), Cm(2.3),
         "「統合的な理解」 ＋ 「総合的な発揮」",
         size=28, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 矢印１
add_arrow(s, Cm(16.0), Cm(8.2), Cm(1.85), Cm(1.2), GRAY_LINE)

# 中：単元目標
add_round(s, Cm(5.0), Cm(9.5), Cm(23.8), Cm(2.0),
          LIGHTGRN, line_color=GREEN, line_width=1.5, corner=0.05)
add_text(s, Cm(5.0), Cm(9.5), Cm(23.8), Cm(2.0),
         "②  単元目標 ／ 評価規準 を立てる",
         size=22, bold=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 矢印２
add_arrow(s, Cm(16.0), Cm(11.7), Cm(1.85), Cm(1.2), GRAY_LINE)

# 下：各時の学習
add_round(s, Cm(2.0), Cm(13.0), Cm(29.8), Cm(3.7),
          LIGHTORG, line_color=ORANGE, line_width=1.5, corner=0.04)
add_text(s, Cm(2.5), Cm(13.1), Cm(28.8), Cm(0.9),
         "③  各時間の 学習活動 を配置する",
         size=18, bold=True, color=ORANGE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.5), Cm(14.0), Cm(28.8), Cm(2.5),
         "「到達点 に どう つながるか」 を見える化",
         size=24, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 6 : 単元構想図 のイメージ ← 【最新WG資料の中核】
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "5", "単元構想図 の イメージ", 6)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.1),
         "ネット型 ボール運動（例） ─ 各時の学習が 到達点 へ どうつながるか",
         size=18, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 上部：高次の資質・能力（2つを横並び）
goal_top = Cm(4.3)
goal_h = Cm(2.3)
# 統合的
add_round(s, Cm(0.85), goal_top, Cm(15.8), goal_h,
          BLUE, corner=0.05)
add_text(s, Cm(0.85), goal_top, Cm(15.8), Cm(0.9),
         "統合的な理解",
         size=16, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(0.85), goal_top + Cm(0.8), Cm(15.8), Cm(1.5),
         "攻防を展開し、楽しさを味わう\n知・技 を 理解する",
         size=14, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.3)
# 総合的
add_round(s, Cm(17.2), goal_top, Cm(15.8), goal_h,
          GREEN, corner=0.05)
add_text(s, Cm(17.2), goal_top, Cm(15.8), Cm(0.9),
         "総合的な発揮",
         size=16, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(17.2), goal_top + Cm(0.8), Cm(15.8), Cm(1.5),
         "誰もが楽しむために必要なことを考え\n練習方法・作戦を 工夫する",
         size=14, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.3)

# 矢印（下向き、二本）
add_arrow(s, Cm(7.0), Cm(6.8), Cm(1.5), Cm(0.8), GRAY_LINE)
add_arrow(s, Cm(25.0), Cm(6.8), Cm(1.5), Cm(0.8), GRAY_LINE)

# 中央：単元目標
add_round(s, Cm(2.0), Cm(7.8), Cm(29.85), Cm(1.4),
          NAVY, corner=0.2)
add_text(s, Cm(2.0), Cm(7.8), Cm(29.85), Cm(1.4),
         "単元目標 ／ 評価規準",
         size=18, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 矢印
add_arrow(s, Cm(16.4), Cm(9.4), Cm(1.0), Cm(0.7), GRAY_LINE)

# 下：時間ブロック（6時間として簡略化）
t_top = Cm(10.4)
t_h = Cm(5.5)
t_w = Cm(5.0)
t_gap = Cm(0.27)
t_left0 = Cm(0.85)

lessons = [
    ("1時", "オリエン\nテーション", "試しのゲーム",   NAVY),
    ("2時", "基本の\nボール操作", "簡易ゲーム",     BLUE),
    ("3時", "場や\nルールの工夫", "誰もが楽しむ",   GREEN),
    ("4時", "課題の\n発見と練習", "チームで協働",   ORANGE),
    ("5時", "フェアな\nプレイ",   "公正・共生",     PURPLE),
    ("6時", "リーグ戦\nまとめ",   "知・技の発揮",   RED),
]
for i, (t, content, focus, c) in enumerate(lessons):
    left = t_left0 + (t_w + t_gap) * i
    add_round(s, left, t_top, t_w, t_h, WHITE,
              line_color=c, line_width=1.5, corner=0.05)
    # 時数バンド
    add_rect(s, left, t_top, t_w, Cm(1.0), c)
    add_text(s, left, t_top, t_w, Cm(1.0),
             t, size=15, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 学習内容
    add_text(s, left + Cm(0.2), t_top + Cm(1.2),
             t_w - Cm(0.4), Cm(2.0),
             content, size=13, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.3)
    # 仕切り線
    add_rect(s, left + Cm(0.5), t_top + Cm(3.4),
             t_w - Cm(1.0), Cm(0.05), c)
    # 重点
    add_text(s, left + Cm(0.2), t_top + Cm(3.5),
             t_w - Cm(0.4), Cm(2.0),
             focus, size=12, color=c,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 下メッセージ
add_text(s, Cm(0.8), Cm(16.3), Cm(32.3), Cm(1.4),
         "▶ 各時の学習が、何 へ つながるか を 教師が 説明できる単元へ",
         size=18, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 7 : 教師の思考プロセス ← 【最新WG資料】
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "6", "単元を 構想する 教師の 思考プロセス", 7)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "最新WG資料が示す 4つの 問い",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 4つの吹き出し
items = [
    ("Q1", "この単元の\n『高次の資質・能力』 は 何か？",
     "まず 到達点 を 確認する",          BLUE,   LIGHTBLUE),
    ("Q2", "楽しさを 味わうには、\nどんな 知・技 が 必要か？",
     "ゴールから 必要な要素 を 洗い出す", GREEN,  LIGHTGRN),
    ("Q3", "知・技 を 思・判・表 と\nどう 関連付けて 学ぶか？",
     "「考えながら できる」 学習過程",   ORANGE, LIGHTORG),
    ("Q4", "各時の学習が、到達点 に\nどう つながるか？",
     "学習活動 と 評価機会 を配置",      RED,    LIGHTRED),
]

g_top = Cm(4.7)
g_h = Cm(2.85)
g_left = Cm(0.85)
g_w = Cm(32.15)

for i, (q, body, ans, color, bg) in enumerate(items):
    top = g_top + (g_h + Cm(0.18)) * i
    add_round(s, g_left, top, g_w, g_h, bg,
              line_color=color, line_width=1.5, corner=0.04)
    # Qバッジ
    add_round(s, g_left + Cm(0.4), top + Cm(0.4),
              Cm(2.3), Cm(2.1), color, corner=0.15)
    add_text(s, g_left + Cm(0.4), top + Cm(0.4),
             Cm(2.3), Cm(2.1),
             q, size=24, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 問い
    add_text(s, g_left + Cm(3.0), top + Cm(0.3),
             Cm(17.5), Cm(2.3),
             body, size=18, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.35)
    # 矢印
    add_arrow(s, g_left + Cm(20.7), top + Cm(1.1),
              Cm(0.7), Cm(0.7), GRAY_LINE,
              shape=MSO_SHAPE.RIGHT_ARROW)
    # 答え
    add_text(s, g_left + Cm(21.7), top + Cm(0.3),
             Cm(10.2), Cm(2.3),
             ans, size=16, color=GRAY_TXT,
             anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 8 : 「する・みる・支える・知る」
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "7", "全員が 主役になる 体育へ", 8)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "「できる/できない」 を超えて、関わり方を広げる",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

items = [
    ("する",   "やって楽しむ",  NAVY,   LIGHTBLUE),
    ("みる",   "観察して学ぶ",  BLUE,   LIGHTBLUE),
    ("支える", "仲間を助ける",  GREEN,  LIGHTGRN),
    ("知る",   "意味を知る",    ORANGE, LIGHTORG),
]

c_top = Cm(4.9)
c_w = Cm(7.6)
c_h = Cm(9.0)
c_gap = Cm(0.45)
c_left0 = Cm(1.05)

for i, (word, desc, color, bg) in enumerate(items):
    left = c_left0 + (c_w + c_gap) * i
    add_round(s, left, c_top, c_w, c_h, bg,
              line_color=color, line_width=2.0, corner=0.05)
    add_rect(s, left, c_top, c_w, Cm(0.7), color)
    add_text(s, left + Cm(0.2), c_top + Cm(1.4),
             c_w - Cm(0.4), Cm(3.6),
             word, size=60, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left + Cm(1.5), c_top + Cm(5.4),
             c_w - Cm(3.0), Cm(0.08), color)
    add_text(s, left + Cm(0.2), c_top + Cm(5.9),
             c_w - Cm(0.4), Cm(2.5),
             desc, size=22, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_round(s, Cm(0.85), Cm(14.5), Cm(32.15), Cm(3.4),
          LIGHTORG, line_color=ORANGE, line_width=1.5, corner=0.08)
add_text(s, Cm(1.5), Cm(14.5), Cm(31.0), Cm(3.4),
         "用具・ルール・場の 工夫 で、\n運動が苦手な子も 自分の変化を 実感できる",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 9 : 本日の授業を振り返る視点
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "8", "本日の授業を 振り返る 4つの視点", 9)

items = [
    ("①", "単元構想 の 視点",
     "高次の資質・能力 から 逆算して 各時間が 配置されているか",
     BLUE,   LIGHTBLUE),
    ("②", "高次の 資質・能力",
     "どんな動き・楽しみ方が 育まれたか",
     GREEN,  LIGHTGRN),
    ("③", "多様性の 包摂",
     "全員が参加し 自分の変化を 実感できたか",
     ORANGE, LIGHTORG),
    ("④", "余白と 創意工夫",
     "教師の意図 と 子供の試行錯誤 が 両立したか",
     RED,    LIGHTRED),
]

g_top = Cm(3.2)
g_h = Cm(3.4)
g_left = Cm(0.85)
g_w = Cm(32.15)

for i, (num, t, body, color, bg) in enumerate(items):
    top = g_top + (g_h + Cm(0.2)) * i
    add_round(s, g_left, top, g_w, g_h, bg,
              line_color=color, line_width=1.5, corner=0.04)
    add_round(s, g_left + Cm(0.5), top + Cm(0.55),
              Cm(2.2), Cm(2.2), color, corner=0.5)
    add_text(s, g_left + Cm(0.5), top + Cm(0.55),
             Cm(2.2), Cm(2.2),
             num, size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, g_left + Cm(3.2), top + Cm(0.4),
             Cm(28.5), Cm(1.1),
             t, size=22, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, g_left + Cm(3.2), top + Cm(1.7),
             Cm(28.5), Cm(1.4),
             body, size=18, color=GRAY_TXT,
             anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 10 : 練馬の先生方へ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "9", "練馬の 先生方へ", 10)

add_round(s, Cm(1.5), Cm(3.3), Cm(30.85), Cm(11.5),
          NAVY, corner=0.03)
add_rect(s, Cm(9.0), Cm(4.5), Cm(15.85), Cm(0.14), ORANGE)

add_text(s, Cm(2.0), Cm(4.8), Cm(29.85), Cm(9.8),
         "種目を 教える 授業 から、\n資質・能力を 育てる 授業へ。\n\n単元を 高次の資質・能力 から 逆算 し、\n一人一人の 変化と工夫 を 見取る。",
         size=28, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.55)

add_round(s, Cm(1.5), Cm(15.1), Cm(30.85), Cm(2.7),
          WHITE, line_color=NAVY, line_width=1.0, corner=0.04)
add_text(s, Cm(2.0), Cm(15.2), Cm(29.85), Cm(0.8),
         "▼ 主な参考資料",
         size=13, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.0), Cm(16.0), Cm(29.85), Cm(1.7),
         "中央教育審議会 体育・保健体育、健康、安全ワーキンググループ 第6～10回\n特に第10回（R8.6.5）「単元構想に向けた『統合的な理解』『総合的な発揮』の活用イメージ」",
         size=12, color=GRAY_TXT, line_spacing=1.4)

add_footer(s)

out_path = "/home/user/con30/体育・保健体育の方向性_指導講評.pptx"
prs.save(out_path)
print("Saved:", out_path)
