"""
体育・保健体育の方向性 指導講評 PPTX 生成スクリプト
- 体育館対応：本文 24-32pt、見出し 30-36pt、巨大ワード 60-72pt
- 体育科特化
- 文科省調査官の視点
- スライド10枚、テキストボックスはすべて寸法計算済み
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
GRAY_TXT  = RGBColor(0x2B, 0x2B, 0x2B)
GRAY_SUB  = RGBColor(0x55, 0x55, 0x55)
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
             line_spacing=1.25, autosize=True):
    """テキストボックスを追加。autosize=True で枠に合わせて自動縮小"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Cm(0.12)
    tf.margin_right = Cm(0.12)
    tf.margin_top = Cm(0.05)
    tf.margin_bottom = Cm(0.05)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    # 自動縮小（オーバーフロー時に文字サイズを下げる）
    if autosize:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_SHAPE
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


def add_header(slide, num, title, page):
    add_rect(slide, Cm(0), Cm(0), SW, Cm(2.6), NAVY)
    add_rect(slide, Cm(0), Cm(2.6), SW, Cm(0.2), ORANGE)
    add_round(slide, Cm(0.8), Cm(0.55), Cm(2.2), Cm(1.5),
              ORANGE, corner=0.18)
    add_text(slide, Cm(0.8), Cm(0.55), Cm(2.2), Cm(1.5),
             num, size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    add_text(slide, Cm(3.4), Cm(0.55), Cm(26.5), Cm(1.5),
             title, size=28, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, autosize=False)
    add_text(slide, Cm(30.0), Cm(0.55), Cm(3.0), Cm(1.5),
             f"{page}/{TOTAL}", size=16, color=WHITE,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)


def add_footer(slide):
    add_rect(slide, Cm(0), SH - Cm(0.7), SW, Cm(0.7), NAVY)
    add_text(slide, Cm(0.8), SH - Cm(0.7), Cm(20), Cm(0.7),
             "練馬区立小学校 体育研究会  指導講評",
             size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, autosize=False)
    add_text(slide, Cm(13.0), SH - Cm(0.7), Cm(20), Cm(0.7),
             "出典：中央教育審議会 体育・保健体育、健康、安全WG（第6～9回）",
             size=10, color=WHITE, align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.MIDDLE, autosize=False)


# 共通：本文エリアの開始Y
BODY_TOP = Cm(3.1)


# =====================================================================
# Slide 1 : 改訂の3つの方向性
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "1", "次期改訂を貫く 3つの方向性", 1)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "論点整理（令和7年9月）で示された 体育科 の進む方向",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

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
    # 番号ヘッダー
    add_rect(s, left, col_top, col_w, Cm(2.0), color)
    add_text(s, left, col_top, col_w, Cm(2.0),
             num, size=48, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    # 英語
    add_text(s, left, col_top + Cm(2.2), col_w, Cm(1.5),
             en, size=30, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    # 日本語
    add_text(s, left, col_top + Cm(3.8), col_w, Cm(0.9),
             f"（{jp}）", size=18, color=GRAY_SUB,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    # 仕切り線
    add_rect(s, left + Cm(1.5), col_top + Cm(5.0),
             col_w - Cm(3.0), Cm(0.08), color)
    # サブテーマ
    add_text(s, left, col_top + Cm(5.3), col_w, Cm(1.3),
             sub, size=26, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    # 本文（2行 × 24pt × 1.5 = 1.8cm  ＜ 3.4cm 容器）
    add_text(s, left + Cm(0.4), col_top + Cm(7.2),
             col_w - Cm(0.8), Cm(3.6),
             body, size=22, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.5, autosize=False)

add_footer(s)


# =====================================================================
# Slide 2 : 体育科の中心メッセージ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "2", "体育科の 中心となる 考え方", 2)

# 上：これまで
tb = Cm(3.5)
add_round(s, Cm(2.0), tb, Cm(29.8), Cm(4.2),
          LIGHTRED, line_color=RED, line_width=1.5, corner=0.05)
add_text(s, Cm(2.5), tb + Cm(0.4), Cm(28.8), Cm(1.1),
         "これまでの 課題",
         size=20, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
# 1行 × 34pt = 36pt より少し小さく、確実に収める
add_text(s, Cm(2.5), tb + Cm(1.6), Cm(28.8), Cm(2.4),
         "種目を 教える 授業",
         size=44, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 矢印
arrow = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,
                           Cm(15.5), tb + Cm(4.6), Cm(3.0), Cm(1.8))
arrow.fill.solid()
arrow.fill.fore_color.rgb = ORANGE
arrow.line.fill.background()
arrow.shadow.inherit = False

# 下：これから
af = Cm(10.8)
add_round(s, Cm(2.0), af, Cm(29.8), Cm(5.2),
          LIGHTBLUE, line_color=BLUE, line_width=2.0, corner=0.05)
add_text(s, Cm(2.5), af + Cm(0.5), Cm(28.8), Cm(1.1),
         "次期改訂の 方向性",
         size=20, bold=True, color=BLUE, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
# 「資質・能力を育てる授業へ」13文字 × 48pt × 0.035 ≈ 22cm → 28.8cm に収まる
add_text(s, Cm(2.5), af + Cm(1.8), Cm(28.8), Cm(3.0),
         "資質・能力を 育てる 授業へ",
         size=48, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# キャプション
add_text(s, Cm(0.8), Cm(16.7), Cm(32.3), Cm(1.3),
         "─ 種目は、資質・能力を 育てる 媒介（手段） ─",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

add_footer(s)


# =====================================================================
# Slide 3 : 高次の資質・能力（本文を圧縮）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "3", "授業の軸「高次の資質・能力」", 3)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "個別の知・技を 関連付け、深い学びへ つなぐ 2つの姿",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

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
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(0.85), box_top + Cm(2.4), box_w, Cm(0.9),
         "知識 及び 技能",
         size=18, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
# 4行 × 26pt × 1.5 = 5.5cm  容器 7.5cm
add_text(s, Cm(1.4), box_top + Cm(3.7), box_w - Cm(1.1), Cm(7.0),
         "知 と 技 が\n関連付けられ\n一般化 された 姿\n\n「だから こう動く」",
         size=26, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4, autosize=False)

# 右：総合的な発揮
add_round(s, Cm(17.2), box_top, box_w, box_h, LIGHTGRN,
          line_color=GREEN, line_width=2.0, corner=0.04)
add_rect(s, Cm(17.2), box_top, box_w, Cm(2.2), GREEN)
add_text(s, Cm(17.2), box_top, box_w, Cm(2.2),
         "総合的な 発揮",
         size=32, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(17.2), box_top + Cm(2.4), box_w, Cm(0.9),
         "思考力・判断力・表現力 等",
         size=18, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(17.75), box_top + Cm(3.7), box_w - Cm(1.1), Cm(7.0),
         "状況に応じて 選び\n仲間と 課題解決する 姿\n\n「だから こう工夫する」",
         size=26, bold=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4, autosize=False)

# 下段
add_text(s, Cm(0.8), Cm(16.4), Cm(32.3), Cm(1.3),
         "▶  2つを セット で、 単元を通して 育てる",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

add_footer(s)


# =====================================================================
# Slide 4 : 学習指導要領の示し方（本文を圧縮）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "4", "学習指導要領の 示し方が変わる", 4)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "発達段階に応じ、示し方を 工夫する方向 で検討中",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 左：小1〜4
add_round(s, Cm(0.85), Cm(4.7), Cm(15.8), Cm(12.0),
          LIGHTBLUE, line_color=BLUE, line_width=1.8, corner=0.04)
add_rect(s, Cm(0.85), Cm(4.7), Cm(15.8), Cm(2.2), BLUE)
add_text(s, Cm(0.85), Cm(4.7), Cm(15.8), Cm(2.2),
         "小1 〜 4年",
         size=34, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(0.85), Cm(7.0), Cm(15.8), Cm(1.0),
         "運動の基礎を培う時期",
         size=18, bold=True, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
# 4行 × 26pt × 1.5 = 5.5cm  容器 8.0cm
add_text(s, Cm(1.4), Cm(8.5), Cm(14.7), Cm(8.0),
         "どんな 動き を\n意図しているか\nを軸に\n資質・能力を示す",
         size=28, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4, autosize=False)

# 右：小5以降
add_round(s, Cm(17.2), Cm(4.7), Cm(15.8), Cm(12.0),
          LIGHTORG, line_color=ORANGE, line_width=1.8, corner=0.04)
add_rect(s, Cm(17.2), Cm(4.7), Cm(15.8), Cm(2.2), ORANGE)
add_text(s, Cm(17.2), Cm(4.7), Cm(15.8), Cm(2.2),
         "小5 以降",
         size=34, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(17.2), Cm(7.0), Cm(15.8), Cm(1.0),
         "スポーツを 豊かに経験する時期",
         size=18, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(17.75), Cm(8.5), Cm(14.7), Cm(8.0),
         "種目の示し方を\n柔軟にして\n教師の創意工夫を\n最大限引き出す",
         size=28, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4, autosize=False)

add_footer(s)


# =====================================================================
# Slide 5 : 「余白」の創出
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "5", "「余白」を生み出す 体育へ", 5)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "深い学びの余地 と、多様性を 包む余地 を",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 中央メッセージ
add_round(s, Cm(2.0), Cm(4.7), Cm(29.8), Cm(4.7),
          NAVY, corner=0.04)
add_text(s, Cm(2.5), Cm(5.0), Cm(28.8), Cm(1.0),
         "今回の改訂が 目指す",
         size=18, color=RGBColor(0xCC, 0xDD, 0xEE),
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(2.5), Cm(6.2), Cm(28.8), Cm(2.8),
         "教師の余白 ＝ 児童の 深い学びの余地",
         size=32, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 3つの支え
mini_top = Cm(10.0)
mini_w = Cm(10.2)
mini_h = Cm(6.4)
mini_gap = Cm(0.45)
mini_left0 = Cm(1.3)

items = [
    ("①", "内容の精選",       "種目を限定せず\n資質・能力で示す",   BLUE),
    ("②", "指導参考資料",     "活動例は別途\n現場の工夫を活かす",   GREEN),
    ("③", "調整授業時数制度", "1割以上を柔軟に\n上乗せ・新教科等", ORANGE),
]
for i, (n, h, b, c) in enumerate(items):
    l = mini_left0 + (mini_w + mini_gap) * i
    add_round(s, l, mini_top, mini_w, mini_h, WHITE,
              line_color=c, line_width=1.5, corner=0.04)
    # 番号丸
    add_round(s, l + Cm(0.4), mini_top + Cm(0.4),
              Cm(1.5), Cm(1.5), c, corner=0.5)
    add_text(s, l + Cm(0.4), mini_top + Cm(0.4), Cm(1.5), Cm(1.5),
             n, size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    # 見出し
    add_text(s, l + Cm(2.1), mini_top + Cm(0.55),
             mini_w - Cm(2.3), Cm(1.2),
             h, size=20, bold=True, color=c,
             anchor=MSO_ANCHOR.MIDDLE, autosize=False)
    # 本文（2行 × 20pt × 1.5 = 2.1cm  容器 3.6cm）
    add_text(s, l + Cm(0.4), mini_top + Cm(2.6),
             mini_w - Cm(0.8), Cm(3.6),
             b, size=20, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.5, autosize=False)

add_footer(s)


# =====================================================================
# Slide 6 : 「する・みる・支える・知る」
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "6", "全員が 主役になる 体育へ", 6)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "「できる/できない」を超えて、関わり方を広げる",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

items = [
    ("する",    "やって楽しむ",   NAVY,   LIGHTBLUE),
    ("みる",    "観察して学ぶ",   BLUE,   LIGHTBLUE),
    ("支える",  "仲間を助ける",   GREEN,  LIGHTGRN),
    ("知る",    "意味を知る",     ORANGE, LIGHTORG),
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
    # 巨大ワード（3文字「支える」が最長 = 3 × 60pt × 0.035 ≈ 6.3cm < 7.6cm）
    add_text(s, left + Cm(0.2), c_top + Cm(1.4), c_w - Cm(0.4), Cm(3.6),
             word, size=60, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    # 仕切り線
    add_rect(s, left + Cm(1.5), c_top + Cm(5.4),
             c_w - Cm(3.0), Cm(0.08), color)
    # 説明
    add_text(s, left + Cm(0.2), c_top + Cm(5.9), c_w - Cm(0.4), Cm(2.5),
             desc, size=22, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)

# 下メッセージ
add_round(s, Cm(0.85), Cm(14.5), Cm(32.15), Cm(3.4),
          LIGHTORG, line_color=ORANGE, line_width=1.5, corner=0.08)
add_text(s, Cm(1.5), Cm(14.5), Cm(31.0), Cm(3.4),
         "用具・ルール・場の工夫で、\n運動が苦手な子も 自分の変化を 実感できる",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5, autosize=False)

add_footer(s)


# =====================================================================
# Slide 7 : ICT × 体育（本文を圧縮）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "7", "ICT × 体育 ─ 深い学びの道具として", 7)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "「何のために使うか」を 常に 問い直す",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 左：活かす
left_w = Cm(15.8)
add_round(s, Cm(0.85), Cm(4.7), left_w, Cm(12.0),
          LIGHTBLUE, line_color=BLUE, line_width=1.8, corner=0.04)
add_rect(s, Cm(0.85), Cm(4.7), left_w, Cm(2.0), BLUE)
add_text(s, Cm(0.85), Cm(4.7), left_w, Cm(2.0),
         "活かす 使い方",
         size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
# 3項目 × 22pt × 1.35 = 6行で約7.5cm 容器 9.5cm
add_text(s, Cm(1.4), Cm(7.1), left_w - Cm(1.1), Cm(9.4),
         "● 動きの 可視化\n　（動画で 振り返り）\n● データ で 検証\n　（タイム・歩数 等）\n● クラウド で 共有\n　（作戦・気付き）",
         size=22, color=GRAY_TXT, line_spacing=1.65,
         autosize=False)

# 右：避ける
right_l = Cm(17.2)
right_w = Cm(15.8)
add_round(s, right_l, Cm(4.7), right_w, Cm(12.0),
          LIGHTRED, line_color=RED, line_width=1.8, corner=0.04)
add_rect(s, right_l, Cm(4.7), right_w, Cm(2.0), RED)
add_text(s, right_l, Cm(4.7), right_w, Cm(2.0),
         "避けたい こと",
         size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, right_l + Cm(0.55), Cm(7.1), right_w - Cm(1.1), Cm(9.4),
         "● 機器操作の 目的化\n　（触ることが目的に）\n● 身体活動時間の 減少\n　（動く時間を 削らない）\n● 端末頼みの 一斉視聴\n　（主体性を 奪わない）",
         size=22, color=GRAY_TXT, line_spacing=1.65,
         autosize=False)

add_footer(s)


# =====================================================================
# Slide 8 : 発育・発達と安全（巨大文字を圧縮、本文を3行に）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "8", "発育・発達を 踏まえた指導へ", 8)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.3),
         "段階性に応じた指導 が、安全 と 深い学び を支える",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 中央パネル
add_round(s, Cm(2.0), Cm(4.7), Cm(29.8), Cm(11.5),
          LIGHTBLUE, line_color=BLUE, line_width=1.5, corner=0.04)

# 上見出し
add_text(s, Cm(2.5), Cm(5.1), Cm(28.8), Cm(1.0),
         "神経系の発達が完成に近づく時期",
         size=20, bold=True, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 巨大文字（短く・スペース無し）
# 「ゴールデンエイジ」8文字 × 40pt × 0.035 = 11.2cm + 「9-12歳」5文字 ≈ 17cm < 28.8cm
add_text(s, Cm(2.5), Cm(6.5), Cm(28.8), Cm(2.5),
         "9〜12歳  ／  ゴールデンエイジ",
         size=40, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

# 線
add_rect(s, Cm(9.0), Cm(9.4), Cm(15.85), Cm(0.1), ORANGE)

# 本文（3項目 行間で空け、容器 5.8cm に収める）
add_text(s, Cm(3.0), Cm(9.8), Cm(28.0), Cm(6.0),
         "● 新しい動作や技を 短時間で 習得できる時期\n● 多様な動きの経験 が一生の財産になる\n● 過度な同一運動・早期専門化は 障害のリスク",
         size=22, color=GRAY_TXT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=2.2, autosize=False)

# 下メッセージ
add_round(s, Cm(0.85), Cm(16.5), Cm(32.15), Cm(1.4),
          NAVY, corner=0.2)
add_text(s, Cm(0.85), Cm(16.5), Cm(32.15), Cm(1.4),
         "体力テストで計れない 多様な動きを 小学校で経験させる",
         size=20, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)

add_footer(s)


# =====================================================================
# Slide 9 : 本日の授業を振り返る視点
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "9", "本日の授業を 振り返る 4つの視点", 9)

items = [
    ("①", "高次の 資質・能力",
     "どんな動き・楽しみ方が 育まれたか",
     BLUE,   LIGHTBLUE),
    ("②", "多様性の 包摂",
     "全員が参加し 自分の変化を 実感できたか",
     GREEN,  LIGHTGRN),
    ("③", "ICTの 活用",
     "ICTは「何のため」か／身体活動時間 は確保できたか",
     ORANGE, LIGHTORG),
    ("④", "余白と 創意工夫",
     "教師の意図 と 子供の試行錯誤 が両立したか",
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
    # 番号丸
    add_round(s, g_left + Cm(0.5), top + Cm(0.55),
              Cm(2.2), Cm(2.2), color, corner=0.5)
    add_text(s, g_left + Cm(0.5), top + Cm(0.55),
             Cm(2.2), Cm(2.2),
             num, size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             autosize=False)
    # タイトル
    add_text(s, g_left + Cm(3.2), top + Cm(0.4),
             Cm(28.5), Cm(1.1),
             t, size=22, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE, autosize=False)
    # 本文
    add_text(s, g_left + Cm(3.2), top + Cm(1.7),
             Cm(28.5), Cm(1.4),
             body, size=18, color=GRAY_TXT,
             anchor=MSO_ANCHOR.MIDDLE, autosize=False)

add_footer(s)


# =====================================================================
# Slide 10 : 練馬の先生方へ（メッセージを5行に圧縮）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "10", "練馬の 先生方へ", 10)

# 大メッセージ枠
add_round(s, Cm(1.5), Cm(3.3), Cm(30.85), Cm(11.5),
          NAVY, corner=0.03)
# オレンジ線
add_rect(s, Cm(9.0), Cm(4.5), Cm(15.85), Cm(0.14), ORANGE)

# 5行 × 28pt × 1.55 = 6.8cm  容器 10.0cm
add_text(s, Cm(2.0), Cm(4.8), Cm(29.85), Cm(9.8),
         "種目を 教える 授業 から、\n資質・能力を 育てる 授業へ。\n\n一人一人の 変化と工夫 を 見取り、\n練馬の研究の中で 一緒に 描いていきたい。",
         size=28, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.55, autosize=False)

# 参考資料
add_round(s, Cm(1.5), Cm(15.1), Cm(30.85), Cm(2.7),
          WHITE, line_color=NAVY, line_width=1.0, corner=0.04)
add_text(s, Cm(2.0), Cm(15.2), Cm(29.85), Cm(0.8),
         "▼ 主な参考資料",
         size=13, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE,
         autosize=False)
add_text(s, Cm(2.0), Cm(16.0), Cm(29.85), Cm(1.7),
         "中央教育審議会 体育・保健体育、健康、安全ワーキンググループ\n第6回(R8.1.16)／第7回(R8.2.19)／第8回(R8.3.27)／第9回(R8.4.24)",
         size=12, color=GRAY_TXT, line_spacing=1.4,
         autosize=False)

add_footer(s)

out_path = "/home/user/con30/体育・保健体育の方向性_指導講評.pptx"
prs.save(out_path)
print("Saved:", out_path)
