from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
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


def set_run_font(run, size=28, bold=False, color=GRAY_TXT, name=FONT_JP):
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


def add_text(slide, left, top, width, height, text, size=28, bold=False,
             color=GRAY_TXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.3):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Cm(0.15)
    tf.margin_right = Cm(0.15)
    tf.margin_top = Cm(0.05)
    tf.margin_bottom = Cm(0.05)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        set_run_font(run, size=size, bold=bold, color=color)
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
    add_rect(slide, Cm(0), Cm(0), SW, Cm(2.8), NAVY)
    add_rect(slide, Cm(0), Cm(2.8), SW, Cm(0.22), ORANGE)
    add_round(slide, Cm(0.8), Cm(0.55), Cm(2.4), Cm(1.7),
              ORANGE, corner=0.18)
    add_text(slide, Cm(0.8), Cm(0.55), Cm(2.4), Cm(1.7),
             num, size=32, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Cm(3.6), Cm(0.55), Cm(26), Cm(1.7),
             title, size=32, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Cm(30.2), Cm(0.55), Cm(3.1), Cm(1.7),
             f"{page}/{TOTAL}", size=18, color=WHITE,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def add_footer(slide):
    add_rect(slide, Cm(0), SH - Cm(0.75), SW, Cm(0.75), NAVY)
    add_text(slide, Cm(0.8), SH - Cm(0.75), Cm(20), Cm(0.75),
             "練馬区立小学校 体育研究会  指導講評",
             size=12, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Cm(13.0), SH - Cm(0.75), Cm(20), Cm(0.75),
             "出典：中央教育審議会 体育・保健体育、健康、安全WG（第6～9回）",
             size=11, color=WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================
# Slide 1 : 改訂の3つの方向性
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "1", "次 期 改 訂 を 貫 く 3 つ の 方 向 性", 1)

add_text(s, Cm(0.8), Cm(3.5), Cm(32.3), Cm(1.4),
         "論点整理（令和 7 年 9 月）で示された 体育科 の進む方向",
         size=22, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

col_w = Cm(10.5)
col_h = Cm(11.4)
col_top = Cm(5.2)
gap = Cm(0.5)
left0 = Cm(0.85)

cols = [
    ("①", "Excellence", "卓越性",  NAVY,   LIGHTBLUE,
     "深い学び",
     "高次の資質・能力\nを 軸にする"),
    ("②", "Equity",     "公正性",  GREEN,  LIGHTGRN,
     "多様性 の 包摂",
     "全ての子どもが\n楽しさ を 味わう"),
    ("③", "Feasibility","持続性",  ORANGE, LIGHTORG,
     "実現可能性",
     "余白 と 創意工夫\nを 大切にする"),
]

for i, (num, en, jp, color, bg, sub, body) in enumerate(cols):
    left = left0 + (col_w + gap) * i
    add_round(s, left, col_top, col_w, col_h, bg,
              line_color=color, line_width=1.8, corner=0.03)
    add_rect(s, left, col_top, col_w, Cm(2.1), color)
    add_text(s, left, col_top, col_w, Cm(2.1),
             num, size=52, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left, col_top + Cm(2.3), col_w, Cm(1.6),
             en, size=32, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left, col_top + Cm(4.0), col_w, Cm(0.9),
             f"（{jp}）", size=20, color=GRAY_SUB,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left + Cm(1.5), col_top + Cm(5.2),
             col_w - Cm(3.0), Cm(0.08), color)
    add_text(s, left, col_top + Cm(5.5), col_w, Cm(1.5),
             sub, size=28, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Cm(0.4), col_top + Cm(7.4),
             col_w - Cm(0.8), col_h - Cm(7.6),
             body, size=24, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 2 : 体育科の中心メッセージ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "2", "体 育 科 の 中 心 と な る 考 え 方", 2)

# 上：これまで
tb = Cm(3.7)
add_round(s, Cm(2.0), tb, Cm(29.8), Cm(4.0),
          LIGHTRED, line_color=RED, line_width=1.5, corner=0.05)
add_text(s, Cm(2.5), tb + Cm(0.3), Cm(28.8), Cm(1.2),
         "これまでの 課題",
         size=22, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.5), tb + Cm(1.6), Cm(28.8), Cm(2.4),
         "種目 （バレー・跳び箱…） を 教える 授業",
         size=36, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 矢印
arrow = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,
                           Cm(15.5), tb + Cm(4.5), Cm(3.0), Cm(1.8))
arrow.fill.solid()
arrow.fill.fore_color.rgb = ORANGE
arrow.line.fill.background()
arrow.shadow.inherit = False

# 下：これから
af = Cm(11.0)
add_round(s, Cm(2.0), af, Cm(29.8), Cm(5.0),
          LIGHTBLUE, line_color=BLUE, line_width=2.0, corner=0.05)
add_text(s, Cm(2.5), af + Cm(0.4), Cm(28.8), Cm(1.2),
         "次期改訂の 方向性",
         size=22, bold=True, color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.5), af + Cm(1.7), Cm(28.8), Cm(3.0),
         "資質・能力 を 育てる 授業 へ",
         size=48, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# キャプション
add_text(s, Cm(0.8), Cm(16.7), Cm(32.3), Cm(1.3),
         "種目 は、 資質・能力 を 育てるための 媒介 （手段）",
         size=24, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 3 : 高次の資質・能力
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "3", "授 業 づ く り の 軸 「 高 次 の 資 質 ・ 能 力 」", 3)

add_text(s, Cm(0.8), Cm(3.4), Cm(32.3), Cm(1.4),
         "個別の 知 ・ 技 を 関連付け、 深い学び に つなぐ 2 つ の 姿",
         size=22, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

box_top = Cm(5.3)
box_h = Cm(10.5)
box_w = Cm(15.8)

# 左：統合的な理解
add_round(s, Cm(0.85), box_top, box_w, box_h, LIGHTBLUE,
          line_color=BLUE, line_width=2.0, corner=0.04)
add_rect(s, Cm(0.85), box_top, box_w, Cm(2.3), BLUE)
add_text(s, Cm(0.85), box_top, box_w, Cm(2.3),
         "統合的な 理解",
         size=34, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(0.85), box_top + Cm(2.5), box_w, Cm(0.9),
         "知 識 及 び 技 能",
         size=18, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(1.4), box_top + Cm(3.9), box_w - Cm(1.1), Cm(6.0),
         "知 と 技 が\n関連付け られ\n一般化 された 姿\n\n「だから こう動く」",
         size=28, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

# 右：総合的な発揮
add_round(s, Cm(17.2), box_top, box_w, box_h, LIGHTGRN,
          line_color=GREEN, line_width=2.0, corner=0.04)
add_rect(s, Cm(17.2), box_top, box_w, Cm(2.3), GREEN)
add_text(s, Cm(17.2), box_top, box_w, Cm(2.3),
         "総合的な 発揮",
         size=34, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(17.2), box_top + Cm(2.5), box_w, Cm(0.9),
         "思 ・ 判 ・ 表",
         size=18, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(17.75), box_top + Cm(3.9), box_w - Cm(1.1), Cm(6.0),
         "状況に応じ 選び\n仲間と 課題解決 する 姿\n\n「だから こう工夫する」",
         size=28, bold=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

add_text(s, Cm(0.8), Cm(16.4), Cm(32.3), Cm(1.4),
         "▶  2 つ を セット で、 単元 を 通して 育てる",
         size=24, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 4 : 学習指導要領の示し方
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "4", "学 習 指 導 要 領 の 示 し 方 が 変 わ る", 4)

add_text(s, Cm(0.8), Cm(3.4), Cm(32.3), Cm(1.4),
         "発達段階 に 応じ、 示し方 を 工夫する 方向で 検討中",
         size=22, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 左：小1〜4
add_round(s, Cm(0.85), Cm(5.2), Cm(15.8), Cm(11.6),
          LIGHTBLUE, line_color=BLUE, line_width=1.8, corner=0.04)
add_rect(s, Cm(0.85), Cm(5.2), Cm(15.8), Cm(2.2), BLUE)
add_text(s, Cm(0.85), Cm(5.2), Cm(15.8), Cm(2.2),
         "小  1  〜  4  年",
         size=32, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(0.85), Cm(7.5), Cm(15.8), Cm(1.0),
         "運動の 基礎 を 培う 時期",
         size=18, bold=True, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(1.4), Cm(9.0), Cm(14.7), Cm(7.4),
         "どんな 動き を\n意図しているか\n\nを 軸 に\n資質・能力 を\n分かりやすく 示す",
         size=26, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

# 右：小5以降
add_round(s, Cm(17.2), Cm(5.2), Cm(15.8), Cm(11.6),
          LIGHTORG, line_color=ORANGE, line_width=1.8, corner=0.04)
add_rect(s, Cm(17.2), Cm(5.2), Cm(15.8), Cm(2.2), ORANGE)
add_text(s, Cm(17.2), Cm(5.2), Cm(15.8), Cm(2.2),
         "小  5  以  降",
         size=32, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(17.2), Cm(7.5), Cm(15.8), Cm(1.0),
         "スポーツ を 豊かに 経験する 時期",
         size=18, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(17.75), Cm(9.0), Cm(14.7), Cm(7.4),
         "種目 の 示し方 を\n柔軟 に\n\n教師 の 創意工夫\n余白 を\n最大限 引き出す",
         size=26, bold=True, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 5 : 「余白」の創出
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "5", "「 余 白 」 を 生 み 出 す 体 育 へ", 5)

add_text(s, Cm(0.8), Cm(3.4), Cm(32.3), Cm(1.4),
         "深い学び の 余地 と、 多様性 を 包む 余地 を",
         size=22, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 中央メッセージ
add_round(s, Cm(2.0), Cm(5.2), Cm(29.8), Cm(5.2),
          NAVY, corner=0.04)
add_text(s, Cm(2.5), Cm(5.6), Cm(28.8), Cm(1.2),
         "今 回 の 改 訂 が 目 指 す",
         size=20, color=RGBColor(0xCC, 0xDD, 0xEE),
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.5), Cm(7.0), Cm(28.8), Cm(3.0),
         "教師 の 余白  ＝  児童 の 深い学び の 余地",
         size=36, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 3つの支え（調整授業時数等）
mini_top = Cm(11.0)
mini_w = Cm(10.2)
mini_h = Cm(5.5)
mini_gap = Cm(0.45)
mini_left0 = Cm(1.3)

items = [
    ("①", "内容 の 精選",     "種目を限定せず\n資質・能力で示す",  BLUE),
    ("②", "指導参考資料",     "活動例 は 別途\n現場の工夫を阻まない", GREEN),
    ("③", "調整授業時数 制度", "1割以上を 柔軟に\n上乗せ・新教科 等", ORANGE),
]
for i, (n, h, b, c) in enumerate(items):
    l = mini_left0 + (mini_w + mini_gap) * i
    add_round(s, l, mini_top, mini_w, mini_h, WHITE,
              line_color=c, line_width=1.5, corner=0.05)
    # 番号丸
    add_round(s, l + Cm(0.4), mini_top + Cm(0.4),
              Cm(1.5), Cm(1.5), c, corner=0.5)
    add_text(s, l + Cm(0.4), mini_top + Cm(0.4), Cm(1.5), Cm(1.5),
             n, size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, l + Cm(2.1), mini_top + Cm(0.55),
             mini_w - Cm(2.3), Cm(1.2),
             h, size=20, bold=True, color=c,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, l + Cm(0.4), mini_top + Cm(2.4),
             mini_w - Cm(0.8), Cm(2.8),
             b, size=18, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 6 : 「する・みる・支える・知る」
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "6", "全 員 が 主 役 に な る 体 育 へ", 6)

add_text(s, Cm(0.8), Cm(3.4), Cm(32.3), Cm(1.4),
         "「できる ／ できない」 を 超えて、 関わり方 を 広げる",
         size=22, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

items = [
    ("する",    "やって 楽しむ",    NAVY,   LIGHTBLUE),
    ("みる",    "観察して 学ぶ",   BLUE,   LIGHTBLUE),
    ("支える",  "仲間を 助ける",   GREEN,  LIGHTGRN),
    ("知る",    "意味を 知る",     ORANGE, LIGHTORG),
]

c_top = Cm(5.4)
c_w = Cm(7.6)
c_h = Cm(9.2)
c_gap = Cm(0.45)
c_left0 = Cm(1.05)

for i, (word, desc, color, bg) in enumerate(items):
    left = c_left0 + (c_w + c_gap) * i
    add_round(s, left, c_top, c_w, c_h, bg,
              line_color=color, line_width=2.0, corner=0.05)
    add_rect(s, left, c_top, c_w, Cm(0.7), color)
    add_text(s, left, c_top + Cm(1.5), c_w, Cm(3.8),
             word, size=66, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left + Cm(1.5), c_top + Cm(5.6),
             c_w - Cm(3.0), Cm(0.1), color)
    add_text(s, left, c_top + Cm(6.2), c_w, Cm(2.6),
             desc, size=24, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_round(s, Cm(0.85), Cm(15.2), Cm(32.15), Cm(2.7),
          LIGHTORG, line_color=ORANGE, line_width=1.5, corner=0.1)
add_text(s, Cm(0.85), Cm(15.2), Cm(32.15), Cm(2.7),
         "用具 ・ ルール ・ 場 の 工夫 で、\n運動が苦手な子も 自分の 変化 を 実感できる",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.45)

add_footer(s)


# =====================================================================
# Slide 7 : ICT × 体育
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "7", "I C T × 体 育  ─  深 い 学 び の 道 具 と し て", 7)

add_text(s, Cm(0.8), Cm(3.4), Cm(32.3), Cm(1.4),
         "「何のため に 使うか」 を 常に 問い直す",
         size=22, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 左：活かす
left_w = Cm(15.8)
add_round(s, Cm(0.85), Cm(5.2), left_w, Cm(11.7),
          LIGHTBLUE, line_color=BLUE, line_width=1.8, corner=0.04)
add_rect(s, Cm(0.85), Cm(5.2), left_w, Cm(2.0), BLUE)
add_text(s, Cm(0.85), Cm(5.2), left_w, Cm(2.0),
         "活 か す  使 い 方",
         size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(1.4), Cm(7.5), left_w - Cm(1.1), Cm(9.0),
         "● 動きの 可視化\n　 （動画で 自分・仲間を 見る）\n\n● データ で 振り返り\n　 （ラップタイム・心拍 等）\n\n● クラウド で 即時共有\n　 （作戦 ・ 気付き）",
         size=22, color=GRAY_TXT, line_spacing=1.5)

# 右：避ける
right_l = Cm(17.2)
right_w = Cm(15.8)
add_round(s, right_l, Cm(5.2), right_w, Cm(11.7),
          LIGHTRED, line_color=RED, line_width=1.8, corner=0.04)
add_rect(s, right_l, Cm(5.2), right_w, Cm(2.0), RED)
add_text(s, right_l, Cm(5.2), right_w, Cm(2.0),
         "避 け た い こ と",
         size=26, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, right_l + Cm(0.5), Cm(7.5), right_w - Cm(1.0), Cm(9.0),
         "● 機器操作 の 目的化\n　 （触ることが 目的になる）\n\n● 身体活動時間 の 減少\n　 （動く時間を 削らない）\n\n● 端末頼み の 一斉視聴\n　 （主体性 を 奪わない）",
         size=22, color=GRAY_TXT, line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 8 : 発育・発達と安全
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "8", "発 育 ・ 発 達 を 踏 ま え た 指 導 へ", 8)

add_text(s, Cm(0.8), Cm(3.4), Cm(32.3), Cm(1.4),
         "段階性 に 応じた 指導 が、 安全 と 深い学び を 支える",
         size=22, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 中央タイムライン
add_round(s, Cm(2.0), Cm(5.3), Cm(29.8), Cm(10.5),
          LIGHTBLUE, line_color=BLUE, line_width=1.5, corner=0.04)

# 上：神経系
add_text(s, Cm(2.5), Cm(5.6), Cm(28.8), Cm(1.2),
         "神 経 系 の 発 達 が 完 成 に 近 づ く 時 期",
         size=22, bold=True, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 巨大年齢
add_text(s, Cm(2.5), Cm(7.0), Cm(28.8), Cm(3.0),
         "9 〜 12 歳  ／  ゴ ー ル デ ン エ イ ジ",
         size=44, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 線
add_rect(s, Cm(8.0), Cm(10.3), Cm(17.85), Cm(0.12), ORANGE)

# 下：意味
add_text(s, Cm(2.5), Cm(10.7), Cm(28.8), Cm(4.8),
         "● 新しい 動作 や 技 を 短時間で 正確に 習得できる\n\n● この時期の 多様な動きの経験 が 一生の財産 に\n\n● 過度な 同一運動 ・ 早期専門化 は スポーツ障害 のリスク",
         size=24, color=GRAY_TXT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

# 下メッセージ
add_round(s, Cm(0.85), Cm(16.3), Cm(32.15), Cm(1.6),
          NAVY, corner=0.2)
add_text(s, Cm(0.85), Cm(16.3), Cm(32.15), Cm(1.6),
         "「 体力テスト で 計れない 多様な動き 」 を 小学校 で 経験させる",
         size=20, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 9 : 本日の授業を振り返る視点
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "9", "本 日 の 授 業 を 振 り 返 る 4 つ の 視 点", 9)

items = [
    ("①", "高次の 資質・能力",
     "どんな 動き ・ 楽しみ方 が 育まれたか",
     BLUE,   LIGHTBLUE),
    ("②", "多様性 の 包摂",
     "全員が 参加 し 自分の 変化 を 実感できたか",
     GREEN,  LIGHTGRN),
    ("③", "I C T の 活用",
     "I C T は 何のため か ／ 身体活動時間 を 確保できたか",
     ORANGE, LIGHTORG),
    ("④", "余白 と 創意工夫",
     "教師の 意図 と 子供の 試行錯誤 が 両立 したか",
     RED,    LIGHTRED),
]

g_top = Cm(3.5)
g_h = Cm(3.3)
g_left = Cm(0.85)
g_w = Cm(32.15)

for i, (num, t, body, color, bg) in enumerate(items):
    top = g_top + (g_h + Cm(0.18)) * i
    add_round(s, g_left, top, g_w, g_h, bg,
              line_color=color, line_width=1.5, corner=0.04)
    add_round(s, g_left + Cm(0.5), top + Cm(0.55),
              Cm(2.2), Cm(2.2), color, corner=0.5)
    add_text(s, g_left + Cm(0.5), top + Cm(0.55),
             Cm(2.2), Cm(2.2),
             num, size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, g_left + Cm(3.2), top + Cm(0.35),
             Cm(12.5), Cm(1.1),
             t, size=24, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, g_left + Cm(3.2), top + Cm(1.6),
             Cm(28.5), Cm(1.5),
             body, size=20, color=GRAY_TXT,
             anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 10 : 練馬の先生方へ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "10", "練 馬 の 先 生 方 へ", 10)

add_round(s, Cm(1.5), Cm(3.6), Cm(30.85), Cm(11.5),
          NAVY, corner=0.03)
add_rect(s, Cm(8.0), Cm(5.0), Cm(17.85), Cm(0.14), ORANGE)

add_text(s, Cm(2.0), Cm(5.5), Cm(29.85), Cm(9.0),
         "種目 を 教える 授業 から、\n資質・能力 を 育てる 授業 へ。\n\n「できる ／ できない」 を 超え、\n一人一人 の 変化 と 工夫 を 見取る。\n\n練馬 の 体育研究 の 中で、\n一緒に 描いて いきたい。",
         size=32, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.6)

# 参考資料
add_round(s, Cm(1.5), Cm(15.3), Cm(30.85), Cm(2.6),
          WHITE, line_color=NAVY, line_width=1.0, corner=0.04)
add_text(s, Cm(2.0), Cm(15.4), Cm(29.85), Cm(0.9),
         "▼ 主 な 参 考 資 料",
         size=14, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.0), Cm(16.2), Cm(29.85), Cm(1.6),
         "中央教育審議会 体育・保健体育、健康、安全 ワーキンググループ\n第 6 回 (R8.1.16) ／ 第 7 回 (R8.2.19) ／ 第 8 回 (R8.3.27) ／ 第 9 回 (R8.4.24)",
         size=13, color=GRAY_TXT, line_spacing=1.5)

add_footer(s)

out_path = "/home/user/con30/体育・保健体育の方向性_指導講評.pptx"
prs.save(out_path)
print("Saved:", out_path)
