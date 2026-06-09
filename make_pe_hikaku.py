"""
体育・保健体育 取りまとめ骨子（案）比較資料 PPTX 生成
- WG取りまとめ骨子（案）令和8年6月5日 第10回 を反映
- 「現行 ⇔ 今後」の比較に特化
- 小学校体育に特化／文科省調査官の視点
- 体育館対応：本文 22-32pt
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
GRAY_LINE = RGBColor(0xBB, 0xBB, 0xBB)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

FONT_JP = "メイリオ"

prs = Presentation()
prs.slide_width  = Cm(33.867)
prs.slide_height = Cm(19.05)
SW = prs.slide_width
SH = prs.slide_height
TOTAL = 11
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


def add_arrow(slide, left, top, width, height, color, shape=MSO_SHAPE.RIGHT_ARROW):
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
             title, size=27, bold=True, color=WHITE,
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
             "出典：WG取りまとめ骨子（案）令和8年6月5日 第10回",
             size=10, color=WHITE, align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.MIDDLE)


def add_compare_header(slide, top, left_w, right_w, gap, left_label, right_label):
    """現行（赤）・今後（青）の見出しバンドを返す"""
    l_left = Cm(0.85)
    r_left = Cm(0.85) + left_w + gap
    add_rect(slide, l_left, top, left_w, Cm(1.3), RED)
    add_text(slide, l_left, top, left_w, Cm(1.3),
             left_label, size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, r_left, top, right_w, Cm(1.3), BLUE)
    add_text(slide, r_left, top, right_w, Cm(1.3),
             right_label, size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return l_left, r_left


# =====================================================================
# Slide 1 : 表紙
# =====================================================================
s = prs.slides.add_slide(blank)
add_rect(s, Cm(0), Cm(0), SW, SH, WHITE)
add_rect(s, Cm(0), Cm(0), SW, Cm(2.5), NAVY)
add_rect(s, Cm(0), Cm(2.5), SW, Cm(0.25), ORANGE)
add_rect(s, Cm(0), SH - Cm(1.0), SW, Cm(1.0), NAVY)
add_rect(s, Cm(0), SH - Cm(1.15), SW, Cm(0.15), ORANGE)

add_text(s, Cm(0.8), Cm(0.55), Cm(32), Cm(1.5),
         "練馬区立小学校 体育研究会  指導講評",
         size=24, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

add_round(s, Cm(2.5), Cm(3.8), Cm(28.8), Cm(11.0), LIGHTBLUE,
          line_color=BLUE, line_width=1.0, corner=0.03)

add_text(s, Cm(2.5), Cm(4.3), Cm(28.8), Cm(1.2),
         "WG取りまとめ骨子（案） 令和8年6月5日 第10回 から",
         size=19, bold=True, color=BLUE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Cm(2.5), Cm(5.8), Cm(28.8), Cm(3.2),
         "小学校体育は\nこう変わる",
         size=46, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.15)

add_text(s, Cm(2.5), Cm(10.0), Cm(28.8), Cm(3.2),
         "─ 現行 と 今後 の 比較 ─\n種目を 教える から、資質・能力を 育てる へ",
         size=24, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.4)

add_text(s, Cm(2.5), Cm(13.6), Cm(28.8), Cm(1.0),
         "本日の授業を踏まえて／次期学習指導要領の動向",
         size=14, color=GRAY_TXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Cm(0.8), SH - Cm(1.0), Cm(32), Cm(1.0),
         "練馬区立小学校 体育研究会  指導講評資料",
         size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================
# Slide 2 : 全体像（3つの大変化）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "1", "小学校体育 3つの 大変化", 2)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.2),
         "骨子（案）が示す、小学校体育に関わる 特に大きな 3つ",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

cards = [
    ("①", "「運動遊び」\nの拡大", "小1〜4 が\n「運動遊び」を\n通した 学びに", NAVY, LIGHTBLUE),
    ("②", "「態度」評価\nの転換", "公正・協力 等が\n「知・技」中心の\n『関わり方』に", GREEN, LIGHTGRN),
    ("③", "「余白」の\n創出", "種目を 限定せず\n資質・能力 を\n育てる 示し方に", ORANGE, LIGHTORG),
]
col_w = Cm(10.5)
col_h = Cm(11.6)
col_top = Cm(4.7)
gap = Cm(0.5)
left0 = Cm(0.85)

for i, (num, title, body, color, bg) in enumerate(cards):
    left = left0 + (col_w + gap) * i
    add_round(s, left, col_top, col_w, col_h, bg,
              line_color=color, line_width=1.8, corner=0.03)
    add_rect(s, left, col_top, col_w, Cm(2.0), color)
    add_text(s, left, col_top, col_w, Cm(2.0),
             num, size=48, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Cm(0.3), col_top + Cm(2.3), col_w - Cm(0.6), Cm(2.6),
             title, size=26, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.2)
    add_rect(s, left + Cm(1.5), col_top + Cm(5.2),
             col_w - Cm(3.0), Cm(0.08), color)
    add_text(s, left + Cm(0.4), col_top + Cm(5.6), col_w - Cm(0.8), Cm(5.6),
             body, size=23, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.4)

add_footer(s)


# =====================================================================
# Slide 3 : 大変化① 運動遊びの拡大（現行⇔今後）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "2", "大変化① 「運動遊び」が 広がる", 3)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.1),
         "小1〜4 が 「運動遊び」を通した学び に再整理される",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

left_w = Cm(15.5)
right_w = Cm(15.5)
gap = Cm(1.05)
top0 = Cm(4.4)
add_compare_header(s, top0, left_w, right_w, gap, "現　行", "今　後")

l_left = Cm(0.85)
r_left = Cm(0.85) + left_w + gap

# 現行ボックス
add_round(s, l_left, top0 + Cm(1.5), left_w, Cm(10.8), LIGHTRED,
          line_color=RED, line_width=1.2, corner=0.03)
add_text(s, l_left + Cm(0.5), top0 + Cm(1.8), left_w - Cm(1.0), Cm(10.3),
         "低学年 → 運動遊び\n中学年 → 「○○運動」\n\n● 遊び＝低学年 だけ\n● 中学年から「運動の\n　 やり方」を 教える\n● 指定運動の実施 が\n　 目的化 しがち",
         size=22, color=GRAY_TXT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.4)

# 矢印
add_arrow(s, Cm(16.0), top0 + Cm(5.6), Cm(1.85), Cm(1.6), ORANGE)

# 今後ボックス
add_round(s, r_left, top0 + Cm(1.5), right_w, Cm(10.8), LIGHTBLUE,
          line_color=BLUE, line_width=1.2, corner=0.03)
add_text(s, r_left + Cm(0.5), top0 + Cm(1.8), right_w - Cm(1.0), Cm(10.3),
         "小1〜4 →「運動遊び」\n　を通した 学びに\n\n● 中学年まで 遊び の要素\n● 「やってみたい」を\n　 引き出す 場づくり\n● 夢中で動く中で\n　 ねらう 動き が育つ",
         size=22, color=GRAY_TXT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.4)

add_footer(s)


# =====================================================================
# Slide 4 : 大変化①補足 運動遊びの授業イメージ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "3", "「運動遊び」の 授業イメージ", 4)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.1),
         "発達段階に応じて 内発的動機 を引き出す（ボール運動系の例）",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 3段階のカード
stages = [
    ("低学年", "ボールを使っての\n運動遊び", "コーンに\n当てたい！", NAVY, LIGHTBLUE),
    ("中学年", "ボール運動遊び", "たくさん\nゴールしたい！", BLUE, LIGHTBLUE),
    ("高学年", "ボール運動", "仲間と連携して\nゴールしたい！", GREEN, LIGHTGRN),
]
c_top = Cm(4.6)
c_w = Cm(10.3)
c_h = Cm(8.5)
c_gap = Cm(0.5)
c_left0 = Cm(0.95)

for i, (grade, ryouiki, motiv, color, bg) in enumerate(stages):
    left = c_left0 + (c_w + c_gap) * i
    add_round(s, left, c_top, c_w, c_h, bg,
              line_color=color, line_width=1.8, corner=0.04)
    add_rect(s, left, c_top, c_w, Cm(1.6), color)
    add_text(s, left, c_top, c_w, Cm(1.6),
             grade, size=26, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Cm(0.3), c_top + Cm(1.9), c_w - Cm(0.6), Cm(2.4),
             ryouiki, size=21, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.2)
    # 吹き出し風
    add_round(s, left + Cm(0.8), c_top + Cm(4.6), c_w - Cm(1.6), Cm(3.3),
              WHITE, line_color=color, line_width=1.5, corner=0.1)
    add_text(s, left + Cm(0.8), c_top + Cm(4.6), c_w - Cm(1.6), Cm(3.3),
             motiv, size=24, bold=True, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.3)

# 矢印（カード間）
for i in range(2):
    ax = c_left0 + c_w + c_gap * 0.5 + (c_w + c_gap) * i - Cm(0.35)
    add_arrow(s, ax, c_top + Cm(4.0), Cm(0.7), Cm(0.8), ORANGE)

# 下メッセージ
add_round(s, Cm(0.85), Cm(14.0), Cm(32.15), Cm(3.4),
          LIGHTORG, line_color=ORANGE, line_width=1.5, corner=0.08)
add_text(s, Cm(1.5), Cm(14.0), Cm(31.0), Cm(3.4),
         "先生の役割 ＝ 「思わず○○したくなる」 場や用具の工夫\n動きの習得 が軸。でも それ だけ が 目的ではない",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 5 : 大変化② 態度評価の転換（現行⇔今後）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "4", "大変化② 「態度」評価 が 変わる", 5)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.1),
         "公正・協力・責任 等が 「知・技」中心の『運動との関わり方』に",
         size=19, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

left_w = Cm(15.5)
right_w = Cm(15.5)
gap = Cm(1.05)
top0 = Cm(4.4)
add_compare_header(s, top0, left_w, right_w, gap, "現　行", "今　後")

l_left = Cm(0.85)
r_left = Cm(0.85) + left_w + gap

add_round(s, l_left, top0 + Cm(1.5), left_w, Cm(10.8), LIGHTRED,
          line_color=RED, line_width=1.2, corner=0.03)
add_text(s, l_left + Cm(0.5), top0 + Cm(1.8), left_w - Cm(1.0), Cm(10.3),
         "【区分】\n学びに向かう力・人間性等\n\n【見取り】\n「公正に取り組もう\n　としているか」（意思）\n\n【課題】形式的になりがち",
         size=22, color=GRAY_TXT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.4)

add_arrow(s, Cm(16.0), top0 + Cm(5.6), Cm(1.85), Cm(1.6), ORANGE)

add_round(s, r_left, top0 + Cm(1.5), right_w, Cm(10.8), LIGHTBLUE,
          line_color=BLUE, line_width=1.2, corner=0.03)
add_text(s, r_left + Cm(0.5), top0 + Cm(1.8), right_w - Cm(1.0), Cm(10.3),
         "【区分】知識・技能 中心\n（＝運動との関わり方）\n\n【見取り】\n「公正とは 何か・\n　なぜ大切か」（理解）\n\n【効果】客観的に評価でき",
         size=22, color=GRAY_TXT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.4)

add_footer(s)


# =====================================================================
# Slide 6 : 大変化③ 余白の創出（現行⇔今後）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "5", "大変化③ 「余白」を 創り出す", 6)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.1),
         "「○○運動では××をする」 を 改め、教師の創意工夫の 余地 を確保",
         size=19, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

left_w = Cm(15.5)
right_w = Cm(15.5)
gap = Cm(1.05)
top0 = Cm(4.4)
add_compare_header(s, top0, left_w, right_w, gap, "現　行", "今　後")

l_left = Cm(0.85)
r_left = Cm(0.85) + left_w + gap

add_round(s, l_left, top0 + Cm(1.5), left_w, Cm(10.8), LIGHTRED,
          line_color=RED, line_width=1.2, corner=0.03)
add_text(s, l_left + Cm(0.5), top0 + Cm(1.8), left_w - Cm(1.0), Cm(10.3),
         "【書き方】\n「○○運動では\n　××をする」（限定的）\n\n● 指定運動を全部こなす\n　ことが 課題に\n● 教師・子ども 双方の負担\n● 多様性の包摂 が難しい",
         size=22, color=GRAY_TXT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.4)

add_arrow(s, Cm(16.0), top0 + Cm(5.6), Cm(1.85), Cm(1.6), ORANGE)

add_round(s, r_left, top0 + Cm(1.5), right_w, Cm(10.8), LIGHTBLUE,
          line_color=BLUE, line_width=1.2, corner=0.03)
add_text(s, r_left + Cm(0.5), top0 + Cm(1.8), right_w - Cm(1.0), Cm(10.3),
         "【書き方】\n小1〜4「どんな 動きか」\n小5〜「種目を 柔軟に」\n\n● 育てる 資質・能力を\n　分かりやすく\n● 具体例は 指導参考資料で\n● 余白＝深い学び の余地",
         size=22, color=GRAY_TXT,
         anchor=MSO_ANCHOR.TOP, line_spacing=1.4)

add_footer(s)


# =====================================================================
# Slide 7 : 高次の資質・能力（新しい軸）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "6", "新しい軸 「高次の資質・能力」", 7)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.1),
         "単元の 到達点 を示す 羅針盤（小・ボール運動 の例）",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

box_top = Cm(4.5)
box_h = Cm(10.8)
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
add_text(s, Cm(1.4), box_top + Cm(3.6), box_w - Cm(1.1), Cm(7.0),
         "操作 と 連携 で\nゲームを展開し\n楽しさを味わえると\n理解する\n\n「だから こう動く」",
         size=24, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.35)

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
add_text(s, Cm(17.75), box_top + Cm(3.6), box_w - Cm(1.1), Cm(7.0),
         "「誰もが」 楽しむために\n必要なことを考え\nルールや作戦を 選ぶ\n\n「だから こう工夫する」",
         size=24, bold=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.35)

add_text(s, Cm(0.8), Cm(15.5), Cm(32.3), Cm(1.3),
         "▶ キーワードは 「誰もが」 ── 全員が楽しむ工夫 を 考えさせる",
         size=22, bold=True, color=ORANGE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 8 : 学習の系統性（12年間）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "7", "12年間の 学びの 見通し", 8)

add_text(s, Cm(0.8), Cm(3.0), Cm(32.3), Cm(1.1),
         "4年ごとに 区切り、発達段階に応じた 学びへ",
         size=20, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

stages = [
    ("小1〜4", "各種の運動の\n基礎を培う時期",
     "「運動遊び」を通して\n「好き」「楽しい」\n「やってみたい」を育む", NAVY, LIGHTBLUE),
    ("小5〜中2", "多くの領域を経験し\nスポーツの基礎\nに触れる時期",
     "簡易ルールの「運動」で\n全員が安心して\n取り組める機会を保障", GREEN, LIGHTGRN),
    ("中3〜高3", "卒業後も多様に\n関わる時期",
     "選択制で\n生涯スポーツの\n楽しみ方を考える", ORANGE, LIGHTORG),
]
c_top = Cm(4.6)
c_w = Cm(10.3)
c_h = Cm(9.5)
c_gap = Cm(0.5)
c_left0 = Cm(0.95)

for i, (grade, period, body, color, bg) in enumerate(stages):
    left = c_left0 + (c_w + c_gap) * i
    add_round(s, left, c_top, c_w, c_h, bg,
              line_color=color, line_width=1.8, corner=0.04)
    add_rect(s, left, c_top, c_w, Cm(1.5), color)
    add_text(s, left, c_top, c_w, Cm(1.5),
             grade, size=26, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Cm(0.3), c_top + Cm(1.7), c_w - Cm(0.6), Cm(2.8),
             period, size=19, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.2)
    add_rect(s, left + Cm(1.2), c_top + Cm(4.7),
             c_w - Cm(2.4), Cm(0.06), color)
    add_text(s, left + Cm(0.3), c_top + Cm(5.0), c_w - Cm(0.6), Cm(4.2),
             body, size=20, color=GRAY_TXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.4)

# 下帯
add_round(s, Cm(0.85), Cm(15.0), Cm(32.15), Cm(2.9),
          NAVY, corner=0.1)
add_text(s, Cm(1.5), Cm(15.0), Cm(31.0), Cm(2.9),
         "小5〜中2 は 「運動嫌い」 が 発生しやすい段階。\n技能の水準を下げるのではなく、全員参加 を 保障する趣旨。",
         size=20, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.5)

add_footer(s)


# =====================================================================
# Slide 9 : 指導方法の Before→After（一覧）
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "8", "指導方法は こう変わる", 9)

# ヘッダー行
row_top = Cm(3.2)
col1_l = Cm(0.85)
col1_w = Cm(6.5)
col2_l = Cm(7.55)
col2_w = Cm(12.4)
col3_l = Cm(20.15)
col3_w = Cm(13.0)

add_rect(s, col1_l, row_top, col1_w, Cm(1.2), NAVY)
add_text(s, col1_l, row_top, col1_w, Cm(1.2), "場面",
         size=18, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, col2_l, row_top, col2_w, Cm(1.2), RED)
add_text(s, col2_l, row_top, col2_w, Cm(1.2), "現行の指導",
         size=18, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, col3_l, row_top, col3_w, Cm(1.2), BLUE)
add_text(s, col3_l, row_top, col3_w, Cm(1.2), "今後の指導",
         size=18, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

rows = [
    ("授業の\n起点", "種目・技を どう教えるか", "どんな 動き・楽しみ方 を育てるか"),
    ("低中\n学年", "運動の 正しいやり方 を教える", "「やってみたい」を引き出す 場づくり"),
    ("苦手な\n子", "できるように 補助する", "ルール・場の工夫 で 全員参加"),
    ("公正\n協力", "「がんばろう」と 声かけ", "「なぜ大切か」を 理解 させる"),
    ("評価", "態度を 主観 で見取る", "知識・技能 として 客観的 に"),
    ("計画", "毎年 同じ時期に 同じ運動", "2年間 で 計画的 に"),
]
r_h = Cm(1.95)
r_gap = Cm(0.12)
y = row_top + Cm(1.3)
for i, (scene, before, after) in enumerate(rows):
    bg = RGBColor(0xF7, 0xF7, 0xF7) if i % 2 == 0 else WHITE
    add_rect(s, col1_l, y, col1_w, r_h, NAVY)
    add_text(s, col1_l, y, col1_w, r_h, scene,
             size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
    add_rect(s, col2_l, y, col2_w, r_h, LIGHTRED)
    add_text(s, col2_l + Cm(0.3), y, col2_w - Cm(0.6), r_h, before,
             size=17, color=GRAY_TXT,
             anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, col3_l, y, col3_w, r_h, LIGHTBLUE)
    add_text(s, col3_l + Cm(0.3), y, col3_w - Cm(0.6), r_h, after,
             size=17, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    y += r_h + r_gap

add_footer(s)


# =====================================================================
# Slide 10 : 明日から意識したい5つの問い
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "9", "明日から 意識したい 5つの問い", 10)

questions = [
    "この単元、どんな 「動き」「楽しみ方」 を 育てたいか？",
    "全員が 「やってみたい」 と思える 仕掛け はあるか？",
    "「誰もが楽しむ工夫」 を 子どもに 考えさせて いるか？",
    "公正・協力を 「態度」ではなく 「理解」 で見ているか？",
    "ICTは 身体活動時間 を 削っていないか？",
]
colors = [NAVY, GREEN, ORANGE, BLUE, RED]
bgs = [LIGHTBLUE, LIGHTGRN, LIGHTORG, LIGHTBLUE, LIGHTRED]

g_top = Cm(3.3)
g_h = Cm(2.65)
g_left = Cm(0.85)
g_w = Cm(32.15)

for i, q in enumerate(questions):
    top = g_top + (g_h + Cm(0.18)) * i
    add_round(s, g_left, top, g_w, g_h, bgs[i],
              line_color=colors[i], line_width=1.5, corner=0.06)
    add_round(s, g_left + Cm(0.4), top + Cm(0.4),
              Cm(1.85), Cm(1.85), colors[i], corner=0.5)
    add_text(s, g_left + Cm(0.4), top + Cm(0.4),
             Cm(1.85), Cm(1.85),
             f"Q{i+1}", size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, g_left + Cm(2.6), top, g_w - Cm(3.0), g_h,
             q, size=21, bold=True, color=GRAY_TXT,
             anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)


# =====================================================================
# Slide 11 : 練馬の先生方へ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header(s, "10", "練馬の 先生方へ", 11)

add_round(s, Cm(1.5), Cm(3.3), Cm(30.85), Cm(11.5),
          NAVY, corner=0.03)
add_rect(s, Cm(9.0), Cm(4.5), Cm(15.85), Cm(0.14), ORANGE)

add_text(s, Cm(2.0), Cm(4.6), Cm(29.85), Cm(10.0),
         "種目を 教える 授業 から、\n資質・能力を 育てる 授業へ。\n\n子どもの「やってみたい」を引き出し、\n「誰もが楽しむ工夫」を 考えさせる。",
         size=28, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.6)

add_round(s, Cm(1.5), Cm(15.1), Cm(30.85), Cm(2.7),
          WHITE, line_color=NAVY, line_width=1.0, corner=0.04)
add_text(s, Cm(2.0), Cm(15.2), Cm(29.85), Cm(0.8),
         "▼ 主な参考資料",
         size=13, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Cm(2.0), Cm(16.0), Cm(29.85), Cm(1.7),
         "中央教育審議会 体育・保健体育、健康、安全ワーキンググループ\n「取りまとめ骨子（案）」（令和8年6月5日 第10回 資料2-1）",
         size=12, color=GRAY_TXT, line_spacing=1.4)

add_footer(s)

out_path = "/home/user/con30/体育科_現行と今後の比較_指導講評.pptx"
prs.save(out_path)
print("Saved:", out_path)
