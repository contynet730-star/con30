"""
石神井南中学校 6/3 校内研究 講義スライド生成スクリプト
file527.pptx (石村指導主事) のデザインを踏襲した青系統のスライド30枚
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from copy import deepcopy

# file527と同じ「レトロスペクト」テーマカラー
NAVY = RGBColor(0x34, 0x40, 0x68)        # dk2: 濃紺（タイトル文字）
LIGHT_BLUE = RGBColor(0x1C, 0xAD, 0xE4)   # accent1: 明るい青
DEEP_BLUE = RGBColor(0x26, 0x83, 0xC6)    # accent2: 深い青（バー・見出し）
TEAL = RGBColor(0x28, 0xC4, 0xCC)         # accent3: 青緑
GREEN = RGBColor(0x42, 0xBA, 0x97)        # accent4: 緑
GRAY_LT = RGBColor(0xD9, 0xE0, 0xE6)      # lt2: 薄いグレー
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
DARK_GRAY = RGBColor(0x40, 0x40, 0x40)

# プレゼンテーション初期化（16:9 ワイド）
prs = Presentation()
prs.slide_width = Emu(12192000)
prs.slide_height = Emu(6858000)

SW = prs.slide_width
SH = prs.slide_height

def add_blank_slide():
    """白紙スライドを追加"""
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)

def add_bottom_bar(slide, color=DEEP_BLUE):
    """下部に青いバーを追加（file527 風）"""
    # 細い線
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(6334316), SW, Emu(66484))
    line.fill.solid()
    line.fill.fore_color.rgb = LIGHT_BLUE
    line.line.fill.background()
    # メインバー
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(6400800), SW, Emu(457200))
    bar.fill.solid()
    bar.fill.fore_color.rgb = DEEP_BLUE
    bar.line.fill.background()
    return bar

def add_page_number(slide, num):
    """ページ番号"""
    tb = slide.shapes.add_textbox(Emu(400000), Emu(280000), Emu(800000), Emu(400000))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = str(num)
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = DEEP_BLUE
    r.font.name = "メイリオ"

def add_section_marker(slide, section_text):
    """右上のセクションマーカー"""
    tb = slide.shapes.add_textbox(Emu(7000000), Emu(280000), Emu(4900000), Emu(400000))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = section_text
    r.font.size = Pt(16)
    r.font.color.rgb = DEEP_BLUE
    r.font.name = "メイリオ"

def add_title_box(slide, title_text, top=Emu(700000), height=Emu(900000),
                  size=32, color=NAVY, bold=True, align=PP_ALIGN.LEFT,
                  left=Emu(700000), width=None):
    """大見出し（スライドタイトル）"""
    if width is None:
        width = SW - Emu(1400000)
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = title_text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "メイリオ"
    return tb

def add_text_box(slide, left, top, width, height, lines,
                 default_size=20, default_color=BLACK, line_spacing=1.2,
                 align=PP_ALIGN.LEFT):
    """汎用テキストボックス。"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Emu(60000); tf.margin_right = Emu(60000)
    tf.margin_top = Emu(30000); tf.margin_bottom = Emu(30000)
    for i, item in enumerate(lines):
        if isinstance(item, tuple):
            text, opts = item
        else:
            text, opts = item, {}
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = opts.get('align', align)
        p.line_spacing = opts.get('line_spacing', line_spacing)
        if 'space_before' in opts:
            p.space_before = Pt(opts['space_before'])
        r = p.add_run()
        r.text = text
        r.font.size = Pt(opts.get('size', default_size))
        r.font.bold = opts.get('bold', False)
        r.font.color.rgb = opts.get('color', default_color)
        r.font.name = opts.get('font', "メイリオ")
    return tb

def add_filled_box(slide, left, top, width, height, fill_color, text=None,
                   text_color=WHITE, size=22, bold=True, align=PP_ALIGN.CENTER):
    """色付きボックス（見出し用バナーなど）"""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = fill_color
    box.line.fill.background()
    if text:
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.margin_left = Emu(100000); tf.margin_right = Emu(100000)
        tf.margin_top = Emu(30000); tf.margin_bottom = Emu(30000)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = text_color
        r.font.name = "メイリオ"
    return box

def add_outline_box(slide, left, top, width, height, border_color=DEEP_BLUE,
                    fill_color=None, line_width=2.0):
    """枠線ボックス"""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill_color:
        box.fill.solid()
        box.fill.fore_color.rgb = fill_color
    else:
        box.fill.background()
    box.line.color.rgb = border_color
    box.line.width = Pt(line_width)
    return box

# ========================================================================
# スライド1: 表紙
# ========================================================================
s = add_blank_slide()
# 上部装飾バー
bar_top = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Emu(120000))
bar_top.fill.solid(); bar_top.fill.fore_color.rgb = DEEP_BLUE; bar_top.line.fill.background()
bar_top2 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(120000), SW, Emu(40000))
bar_top2.fill.solid(); bar_top2.fill.fore_color.rgb = LIGHT_BLUE; bar_top2.line.fill.background()

# タイトル
add_text_box(s, Emu(800000), Emu(1500000), Emu(10600000), Emu(1200000), [
    ("指導と評価の一体化", {'size': 60, 'bold': True, 'color': NAVY}),
], line_spacing=1.0)
add_text_box(s, Emu(800000), Emu(2700000), Emu(10600000), Emu(800000), [
    ("─ 生徒の姿で語る授業改善 ─", {'size': 36, 'bold': True, 'color': DEEP_BLUE}),
], line_spacing=1.0)
add_text_box(s, Emu(800000), Emu(3600000), Emu(10600000), Emu(600000), [
    ("〜明日、何を変えるか〜", {'size': 28, 'color': DARK_GRAY}),
], line_spacing=1.0)

# 区切り線
sep = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(800000), Emu(4500000), Emu(3000000), Emu(40000))
sep.fill.solid(); sep.fill.fore_color.rgb = LIGHT_BLUE; sep.line.fill.background()

add_text_box(s, Emu(800000), Emu(4650000), Emu(10600000), Emu(1500000), [
    ("令和8年6月3日(水)　石神井南中学校 校内研究", {'size': 22, 'color': NAVY}),
    ("", {'size': 8}),
    ("練馬区教育委員会 指導主事", {'size': 20, 'color': DARK_GRAY}),
    ("紺多　章一郎", {'size': 28, 'bold': True, 'color': NAVY}),
], line_spacing=1.3)

# 下部バー
bar_bot = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(6700000), SW, Emu(158000))
bar_bot.fill.solid(); bar_bot.fill.fore_color.rgb = DEEP_BLUE; bar_bot.line.fill.background()

# ========================================================================
# スライド2: ウォームアップ（生徒の意識調査データ）
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 2)
add_section_marker(s, "ウォームアップ")
add_title_box(s, "生徒は授業をどう見ているか", size=32)

add_filled_box(s, Emu(700000), Emu(1700000), Emu(11000000), Emu(450000), DEEP_BLUE,
               "文部科学省「義務教育に関する意識に係る調査」（令和5年12月）",
               WHITE, 16, True, align=PP_ALIGN.LEFT)
add_text_box(s, Emu(700000), Emu(2200000), Emu(11000000), Emu(400000), [
    ("中学生に「普段の授業について思うこと」を聞いた結果 ──",
     {'size': 18, 'color': DARK_GRAY}),
])

# 肯定的な3項目
y = 2700000
data = [
    ("友達と一緒に学ぶことができて楽しい", "84.1%", GREEN),
    ("授業で学ぶことが、将来役に立つ", "73.8%", DEEP_BLUE),
    ("授業で学ぶ内容は面白い", "67.3%", LIGHT_BLUE),
]
for i, (q, pct, color) in enumerate(data):
    yy = y + i * 700000
    add_outline_box(s, Emu(700000), Emu(yy), Emu(8400000), Emu(600000),
                    border_color=color, line_width=2.0)
    add_text_box(s, Emu(900000), Emu(yy + 80000), Emu(8000000), Emu(450000), [
        (q, {'size': 20, 'color': NAVY}),
    ])
    add_filled_box(s, Emu(9200000), Emu(yy), Emu(2500000), Emu(600000),
                   color, pct, WHITE, 28, True)

add_filled_box(s, Emu(700000), Emu(5100000), Emu(11000000), Emu(400000), NAVY,
               "── では、私たちの評価は、彼らの「学び」を捉えられているか？",
               WHITE, 18, True)
add_text_box(s, Emu(700000), Emu(5550000), Emu(11000000), Emu(700000), [
    ("「これでBにしてよいのか」「あの生徒の頑張りをどう拾うか」",
     {'size': 18, 'color': DEEP_BLUE, 'bold': True, 'align': PP_ALIGN.CENTER}),
    ("そんな迷いに、本日アプローチします",
     {'size': 18, 'color': GREEN, 'bold': True, 'align': PP_ALIGN.CENTER}),
], line_spacing=1.3)

# ========================================================================
# スライド3: 本日のゴール
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 3)
add_section_marker(s, "はじめに")
add_title_box(s, "本日のゴール")

add_text_box(s, Emu(700000), Emu(1900000), Emu(11000000), Emu(500000), [
    ("校長先生から3つの問いをいただきました。", {'size': 22, 'color': DARK_GRAY}),
])

# 3つの問いをボックスで
y = 2600000
for i, (label, txt) in enumerate([
    ("❶", "どんな授業を行うか（指導と評価の一体化）"),
    ("❷", "どんな生徒の姿を見取るか"),
    ("❸", "どう授業改善につなげるか"),
]):
    add_filled_box(s, Emu(900000), Emu(y + i*900000), Emu(700000), Emu(700000),
                   DEEP_BLUE, label, WHITE, 32, True)
    add_text_box(s, Emu(1750000), Emu(y + i*900000 + 100000), Emu(9500000), Emu(600000), [
        (txt, {'size': 26, 'bold': True, 'color': NAVY}),
    ])

add_text_box(s, Emu(700000), Emu(5500000), Emu(11000000), Emu(500000), [
    ("この3つに、30分でお答えします。", {'size': 24, 'bold': True, 'color': GREEN, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド4: Part1 セクション扉
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 4)

# 大きなPart表示
add_filled_box(s, 0, Emu(1800000), SW, Emu(800000), DEEP_BLUE,
               "Part 1", WHITE, 44, True)
add_filled_box(s, 0, Emu(2700000), SW, Emu(1300000), LIGHT_BLUE,
               "なぜ今、「指導と評価の一体化」なのか", WHITE, 40, True)
add_text_box(s, Emu(700000), Emu(4300000), Emu(11000000), Emu(500000), [
    ("（5分）", {'size': 28, 'color': DARK_GRAY, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド5: 学習評価とは何か（学指総則）
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 5)
add_section_marker(s, "Part 1　学習評価の充実")
add_title_box(s, "学習指導要領 総則 ── 評価の本質")

add_filled_box(s, Emu(700000), Emu(1900000), Emu(11000000), Emu(500000), DEEP_BLUE,
               "第3-2 学習評価の充実", WHITE, 22, True, align=PP_ALIGN.LEFT)
add_outline_box(s, Emu(700000), Emu(2400000), Emu(11000000), Emu(2400000),
                border_color=DEEP_BLUE, fill_color=GRAY_LT, line_width=2.0)
add_text_box(s, Emu(900000), Emu(2550000), Emu(10600000), Emu(2200000), [
    ("「単元や題材など内容や時間のまとまりを見通しながら", {'size': 22, 'color': NAVY}),
    ("評価の場面や方法を工夫して、学習の過程や成果を評価し、", {'size': 22, 'color': NAVY}),
    ("指導の改善や学習意欲の向上を図り、", {'size': 22, 'color': NAVY, 'bold': True}),
    ("資質・能力の育成に生かすようにすること」", {'size': 22, 'color': NAVY, 'bold': True}),
    ("", {'size': 10}),
    ("　　　 ── 中学校学習指導要領（平成29年告示）", {'size': 18, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
], line_spacing=1.4)

add_text_box(s, Emu(700000), Emu(5100000), Emu(11000000), Emu(900000), [
    ("ここに評価の本質が詰まっています", {'size': 24, 'bold': True, 'color': GREEN, 'align': PP_ALIGN.CENTER}),
    ("キーワード： 見通し　工夫　指導の改善　意欲の向上", {'size': 20, 'color': DEEP_BLUE, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド6: 評価の目的は2つ（+1）
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 6)
add_section_marker(s, "Part 1　学習評価の目的")
add_title_box(s, "学習評価を行う目的")

y = 1900000
items = [
    ("▶", "生徒の「学習改善」につなげる", "次は○○に取り組むよう生徒に伝える"),
    ("▶", "教師の「指導改善」につなげる", "次の授業で○○を重点的に指導する"),
    ("▶", "慣行として行ってきたことの「見直し」", "今までこうだったから、を勇気をもって見直す"),
]
for i, (mark, main, sub) in enumerate(items):
    bg_color = LIGHT_BLUE if i < 2 else GREEN
    add_filled_box(s, Emu(700000), Emu(y + i*1300000), Emu(900000), Emu(1100000),
                   bg_color, mark, WHITE, 44, True)
    add_outline_box(s, Emu(1700000), Emu(y + i*1300000), Emu(10000000), Emu(1100000),
                    border_color=bg_color, line_width=2.0)
    add_text_box(s, Emu(1850000), Emu(y + i*1300000 + 100000), Emu(9800000), Emu(900000), [
        (main, {'size': 24, 'bold': True, 'color': NAVY}),
        ("例： " + sub, {'size': 18, 'color': DARK_GRAY}),
    ], line_spacing=1.2)

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 中央教育審議会 H31.1 「児童生徒の学習評価の在り方について（報告）」", {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド7: 現状の課題 と 一体化のメリット
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 7)
add_section_marker(s, "Part 1　課題と展望")
add_title_box(s, "課題と「一体化」のメリット", size=32)

# 左：課題
add_filled_box(s, Emu(700000), Emu(1850000), Emu(5400000), Emu(500000),
               RGBColor(0xC0, 0x40, 0x40), "全国の課題", WHITE, 22, True)
add_outline_box(s, Emu(700000), Emu(2350000), Emu(5400000), Emu(2800000),
                border_color=RGBColor(0xC0, 0x40, 0x40), line_width=2.0)
add_text_box(s, Emu(900000), Emu(2450000), Emu(5000000), Emu(2700000), [
    ("○ 学習改善につながっていない", {'size': 17, 'color': NAVY}),
    ("○ 教師により方針が異なる", {'size': 17, 'color': NAVY}),
    ("○ 記録に労力が割かれる", {'size': 17, 'color': NAVY}),
    ("○ 「指導に生かす」と「記録に残す」", {'size': 17, 'color': NAVY}),
    ("　 が区別されていない", {'size': 17, 'color': NAVY}),
    ("○ 観点別評価の誤解", {'size': 17, 'color': NAVY}),
], line_spacing=1.4)

# 右：一体化のメリット
add_filled_box(s, Emu(6300000), Emu(1850000), Emu(5400000), Emu(500000), GREEN,
               "「一体化」で得られるもの", WHITE, 22, True)
add_outline_box(s, Emu(6300000), Emu(2350000), Emu(5400000), Emu(2800000),
                border_color=GREEN, line_width=2.0)
add_text_box(s, Emu(6500000), Emu(2450000), Emu(5000000), Emu(2700000), [
    ("○ 生徒の学習意欲の向上", {'size': 19, 'color': NAVY, 'bold': True}),
    ("", {'size': 4}),
    ("○ 学習効果の向上", {'size': 19, 'color': NAVY, 'bold': True}),
    ("", {'size': 4}),
    ("○ 教師の指導力の向上", {'size': 19, 'color': NAVY, 'bold': True}),
    ("", {'size': 4}),
    ("○ 教育活動全体の質の向上", {'size': 19, 'color': NAVY, 'bold': True}),
], line_spacing=1.3)

add_filled_box(s, Emu(700000), Emu(5400000), Emu(11000000), Emu(700000), DEEP_BLUE,
               "課題を抱えているのは石神井南中だけではありません",
               WHITE, 22, True)

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 中教審 H31.1報告 ／ 文科省 教育課程部会 R8.3資料1-1", {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド8: 最新動向（R8.3 論点整理）
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 8)
add_section_marker(s, "Part 1　最新動向")
add_title_box(s, "評価は今、どこへ向かっているか")

add_filled_box(s, Emu(700000), Emu(1850000), Emu(11000000), Emu(500000), DEEP_BLUE,
               "文部科学省 教育課程部会 総則・評価特別部会（令和8年3月 資料1-1より）",
               WHITE, 18, True, align=PP_ALIGN.LEFT)

y = 2500000
arrows = [
    "形式的かつ過度な「評価材料集め」を抑制する",
    "多様な子供一人一人の「良さや成長」を肯定的に評価する",
    "「学びに向かう力、人間性等」の見取り方を見直す",
]
for i, txt in enumerate(arrows):
    add_filled_box(s, Emu(700000), Emu(y + i*800000), Emu(500000), Emu(600000),
                   LIGHT_BLUE, "▶", WHITE, 28, True)
    add_text_box(s, Emu(1300000), Emu(y + i*800000 + 50000), Emu(10400000), Emu(550000), [
        (txt, {'size': 24, 'color': NAVY}),
    ])

add_filled_box(s, Emu(700000), Emu(5400000), Emu(11000000), Emu(800000), GREEN,
               "「頑張って記録を増やす」時代は終わりつつある", WHITE, 26, True)

# ========================================================================
# スライド9: 「妥当性」と「信頼性」
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 9)
add_section_marker(s, "Part 1　評価の質")
add_title_box(s, "評価が満たすべき2つの条件")

# 左：妥当性
add_filled_box(s, Emu(700000), Emu(1850000), Emu(5300000), Emu(600000), DEEP_BLUE,
               "妥当性", WHITE, 30, True)
add_outline_box(s, Emu(700000), Emu(2450000), Emu(5300000), Emu(2200000),
                border_color=DEEP_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(2600000), Emu(4900000), Emu(2000000), [
    ("評価の対象である", {'size': 20, 'color': NAVY}),
    ("「資質・能力」を", {'size': 20, 'color': NAVY}),
    ("適切に反映していること", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
    ("", {'size': 8}),
    ("→ 個人で確保", {'size': 22, 'bold': True, 'color': GREEN}),
], line_spacing=1.3)

# 右：信頼性
add_filled_box(s, Emu(6400000), Emu(1850000), Emu(5300000), Emu(600000), LIGHT_BLUE,
               "信頼性", WHITE, 30, True)
add_outline_box(s, Emu(6400000), Emu(2450000), Emu(5300000), Emu(2200000),
                border_color=LIGHT_BLUE, line_width=2.0)
add_text_box(s, Emu(6600000), Emu(2600000), Emu(4900000), Emu(2000000), [
    ("教師の主観に流れず、", {'size': 20, 'color': NAVY}),
    ("誰が評価しても", {'size': 20, 'color': NAVY}),
    ("同じ結果になること", {'size': 22, 'bold': True, 'color': LIGHT_BLUE}),
    ("", {'size': 8}),
    ("→ 組織で確保", {'size': 22, 'bold': True, 'color': GREEN}),
], line_spacing=1.3)

add_filled_box(s, Emu(700000), Emu(5000000), Emu(11000000), Emu(800000), GREEN,
               "妥当性は「個人」で、信頼性は「組織」で確保する", WHITE, 26, True)

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 東京都教育委員会「指導と評価の一体化を目指して」（令和2年）", {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド10: Part2 セクション扉
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 10)

add_filled_box(s, 0, Emu(1800000), SW, Emu(800000), DEEP_BLUE,
               "Part 2", WHITE, 44, True)
add_filled_box(s, 0, Emu(2700000), SW, Emu(700000), LIGHT_BLUE,
               "校長の問い❶", WHITE, 32, True)
add_filled_box(s, 0, Emu(3500000), SW, Emu(1000000), TEAL,
               "── どんな授業を行うか ──", WHITE, 36, True)
add_text_box(s, Emu(700000), Emu(4700000), Emu(11000000), Emu(500000), [
    ("（8分）", {'size': 28, 'color': DARK_GRAY, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド11: 単元で評価をデザインする
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 11)
add_section_marker(s, "Part 2　評価のデザイン")
add_title_box(s, "「毎時間 × 全観点」をやめる")

# ×
add_filled_box(s, Emu(700000), Emu(1850000), Emu(900000), Emu(900000),
               RGBColor(0xC0, 0x40, 0x40), "×", WHITE, 48, True)
add_outline_box(s, Emu(1700000), Emu(1850000), Emu(10000000), Emu(900000),
                border_color=RGBColor(0xC0, 0x40, 0x40), line_width=2.0)
add_text_box(s, Emu(1900000), Emu(1900000), Emu(9700000), Emu(800000), [
    ("毎時間、全観点を評価しようとする", {'size': 24, 'bold': True, 'color': NAVY}),
    ("→ 教師は疲弊、生徒は評価されすぎる", {'size': 20, 'color': DARK_GRAY}),
], line_spacing=1.2)

# ○
add_filled_box(s, Emu(700000), Emu(2950000), Emu(900000), Emu(900000),
               GREEN, "○", WHITE, 48, True)
add_outline_box(s, Emu(1700000), Emu(2950000), Emu(10000000), Emu(900000),
                border_color=GREEN, line_width=2.0)
add_text_box(s, Emu(1900000), Emu(3000000), Emu(9700000), Emu(800000), [
    ("単元の中で、いつ・何を見るかを設計する", {'size': 24, 'bold': True, 'color': NAVY}),
    ("→ 重点を決め、記録を絞る", {'size': 20, 'color': DARK_GRAY}),
], line_spacing=1.2)

add_filled_box(s, Emu(700000), Emu(4500000), Emu(11000000), Emu(900000), DEEP_BLUE,
               "これが「指導と評価の一体化」の第一歩", WHITE, 28, True)

add_text_box(s, Emu(700000), Emu(5600000), Emu(11000000), Emu(700000), [
    ("文科省も「形式的・過度な評価材料集めの抑制」を求めています（R8.3論点整理）",
     {'size': 18, 'color': DARK_GRAY, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド12: 評価計画の型（表）
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 12)
add_section_marker(s, "Part 2　評価のデザイン")
add_title_box(s, "単元評価計画の例")

# 表を作成
rows, cols = 8, 5
left = Emu(800000); top = Emu(1900000)
width = Emu(10600000); height = Emu(3400000)
table_shape = s.shapes.add_table(rows, cols, left, top, width, height)
table = table_shape.table

# 列幅
table.columns[0].width = Emu(800000)
table.columns[1].width = Emu(3200000)
table.columns[2].width = Emu(1400000)
table.columns[3].width = Emu(1400000)
table.columns[4].width = Emu(3800000)

# ヘッダー
headers = ["時", "ねらい・学習活動", "重点", "記録", "評価方法"]
for j, h in enumerate(headers):
    cell = table.cell(0, j)
    cell.fill.solid()
    cell.fill.fore_color.rgb = DEEP_BLUE
    tf = cell.text_frame
    tf.margin_left = Emu(40000); tf.margin_right = Emu(40000)
    tf.margin_top = Emu(20000); tf.margin_bottom = Emu(20000)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = ""
    r = p.add_run(); r.text = h
    r.font.size = Pt(18); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = "メイリオ"

# データ
data = [
    ("1", "導入", "知", "", "行動観察"),
    ("2", "展開1", "知 態", "○", "知：小テスト  態：学習シート"),
    ("3", "展開2", "思", "", "行動観察"),
    ("4", "展開3", "思", "○", "ノート"),
    ("5", "展開4", "知", "", "小テスト"),
    ("6", "展開5", "知 態", "○", "知：行動観察  態：学習シート"),
    ("7", "まとめ", "知 思", "○○", "単元テスト"),
]
for i, row in enumerate(data, start=1):
    for j, val in enumerate(row):
        cell = table.cell(i, j)
        if i % 2 == 1:
            cell.fill.solid()
            cell.fill.fore_color.rgb = GRAY_LT
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE
        tf = cell.text_frame
        tf.margin_left = Emu(40000); tf.margin_right = Emu(40000)
        tf.margin_top = Emu(15000); tf.margin_bottom = Emu(15000)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if j != 4 else PP_ALIGN.LEFT
        p.text = ""
        r = p.add_run(); r.text = val
        r.font.size = Pt(16); r.font.color.rgb = NAVY
        r.font.name = "メイリオ"
        if j in (2, 3) and val:
            r.font.bold = True
            r.font.color.rgb = DEEP_BLUE

# 補足
add_text_box(s, Emu(700000), Emu(5450000), Emu(11000000), Emu(820000), [
    ("★ 「重点」 ＝ 指導に生かす評価（毎時間）",
     {'size': 20, 'bold': True, 'color': DEEP_BLUE}),
    ("★ 「記録」 ＝ 総括（評定）に残す評価（節目だけ）",
     {'size': 20, 'bold': True, 'color': GREEN}),
], line_spacing=1.25)

# ========================================================================
# スライド13: 2種類の評価を区別する
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 13)
add_section_marker(s, "Part 2　2種類の評価")
add_title_box(s, "「指導に生かす評価」と「記録に残す評価」")

# 左
add_filled_box(s, Emu(700000), Emu(1850000), Emu(5300000), Emu(700000), LIGHT_BLUE,
               "指導に生かす評価", WHITE, 26, True)
add_filled_box(s, Emu(700000), Emu(2550000), Emu(5300000), Emu(400000), GRAY_LT,
               "（形成的評価）", NAVY, 18, False)
add_outline_box(s, Emu(700000), Emu(2950000), Emu(5300000), Emu(2300000),
                border_color=LIGHT_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(3100000), Emu(4900000), Emu(2100000), [
    ("○ 毎時間、自然な観察で見取る", {'size': 19, 'color': NAVY}),
    ("○ 授業を変えるための材料", {'size': 19, 'color': NAVY}),
    ("○ 記録は必須ではない", {'size': 19, 'bold': True, 'color': DEEP_BLUE}),
    ("", {'size': 6}),
    ("≪具体例≫", {'size': 18, 'color': DARK_GRAY}),
    ("机間指導での観察、振り返りシート、", {'size': 17, 'color': DARK_GRAY}),
    ("ミニテスト、学習シート", {'size': 17, 'color': DARK_GRAY}),
], line_spacing=1.3)

# 右
add_filled_box(s, Emu(6400000), Emu(1850000), Emu(5300000), Emu(700000), DEEP_BLUE,
               "記録に残す評価", WHITE, 26, True)
add_filled_box(s, Emu(6400000), Emu(2550000), Emu(5300000), Emu(400000), GRAY_LT,
               "（総括的評価）", NAVY, 18, False)
add_outline_box(s, Emu(6400000), Emu(2950000), Emu(5300000), Emu(2300000),
                border_color=DEEP_BLUE, line_width=2.0)
add_text_box(s, Emu(6600000), Emu(3100000), Emu(4900000), Emu(2100000), [
    ("○ 単元の節目で計画的に見取る", {'size': 19, 'color': NAVY}),
    ("○ 評定の根拠になる", {'size': 19, 'color': NAVY}),
    ("○ 記録は必須・組織で揃える", {'size': 19, 'bold': True, 'color': DEEP_BLUE}),
    ("", {'size': 6}),
    ("≪具体例≫", {'size': 18, 'color': DARK_GRAY}),
    ("単元テスト、観点別評価、", {'size': 17, 'color': DARK_GRAY}),
    ("成果物（レポート・作品）", {'size': 17, 'color': DARK_GRAY}),
], line_spacing=1.3)

add_text_box(s, Emu(700000), Emu(5500000), Emu(11000000), Emu(560000), [
    ("「2つを区別する」だけで、評価の負担は劇的に減る",
     {'size': 22, 'bold': True, 'color': GREEN, 'align': PP_ALIGN.CENTER}),
])

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 国立教育政策研究所「学習評価の在り方ハンドブック（中学校編）」",
     {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド14: 3種類の評価
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 14)
add_section_marker(s, "Part 2　評価の整理")
add_title_box(s, "学習評価の3種類")

y_base = 1900000
items = [
    ("診断的評価", "学習前", "レディネステスト、コース分けの小テスト", LIGHT_BLUE),
    ("形成的評価", "学習過程", "振り返りシート、教師の観察、ミニテスト", DEEP_BLUE),
    ("総括的評価", "単元末・学期末", "単元テスト、観点別評価、評定", TEAL),
]
for i, (name, when, example, color) in enumerate(items):
    y = y_base + i * 1300000
    add_filled_box(s, Emu(700000), Emu(y), Emu(3200000), Emu(1100000), color,
                   name, WHITE, 32, True)
    add_outline_box(s, Emu(4000000), Emu(y), Emu(7700000), Emu(1100000),
                    border_color=color, line_width=2.0)
    add_text_box(s, Emu(4200000), Emu(y + 100000), Emu(7400000), Emu(900000), [
        ("【タイミング】 " + when, {'size': 18, 'color': NAVY}),
        ("【例】 " + example, {'size': 18, 'color': DARK_GRAY}),
    ], line_spacing=1.3)

add_text_box(s, Emu(700000), Emu(5850000), Emu(11000000), Emu(380000), [
    ("形成的評価で日常的に見取り、総括的評価で記録に残す",
     {'size': 20, 'bold': True, 'color': GREEN, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド15: 単元の授業デザイン
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 15)
add_section_marker(s, "Part 2　単元デザイン")
add_title_box(s, "単元の授業デザイン（10時間モデル）")

# 3つのフェーズ
phases = [
    ("導入", "2時間", "学習の見通しをもつ\n学習計画づくり", "主体的", LIGHT_BLUE),
    ("追究", "6時間", "課題解決に向けた話し合い\n情報精査・考えの形成", "対話的・深い", DEEP_BLUE),
    ("まとめ", "2時間", "学習を評価し修正する振り返り\n新たな問いを見いだす", "主体的・深い", TEAL),
]
x_base = 700000
width = 3650000
for i, (name, t, content, learning, color) in enumerate(phases):
    x = x_base + i * (width + 100000)
    add_filled_box(s, Emu(x), Emu(1900000), Emu(width), Emu(700000), color,
                   name + "　" + t, WHITE, 26, True)
    add_outline_box(s, Emu(x), Emu(2600000), Emu(width), Emu(2200000),
                    border_color=color, line_width=2.0)
    add_text_box(s, Emu(x + 200000), Emu(2750000), Emu(width - 400000), Emu(1500000), [
        (content.split('\n')[0], {'size': 18, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
        (content.split('\n')[1], {'size': 18, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
    ], line_spacing=1.4)
    add_filled_box(s, Emu(x + 300000), Emu(4250000), Emu(width - 600000), Emu(450000),
                   GREEN, "▶ " + learning + "な学び", WHITE, 18, True)

add_filled_box(s, Emu(700000), Emu(5300000), Emu(11000000), Emu(800000), NAVY,
               "教科は違っても、この骨格は共通です", WHITE, 24, True)

# ========================================================================
# スライド16: 問い❶への回答（まとめ）
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 16)
add_section_marker(s, "Part 2　まとめ")
add_title_box(s, "校長の問い❶への回答")

add_filled_box(s, Emu(700000), Emu(1850000), Emu(11000000), Emu(700000), DEEP_BLUE,
               "「どんな授業を行えばよいか」", WHITE, 28, True)

y = 2700000
answers = [
    "単元のスタートで評価計画を立てる",
    "「重点」と「記録」を区別する",
    "形成的評価で授業を変え、総括的評価で記録する",
]
for i, ans in enumerate(answers):
    add_filled_box(s, Emu(900000), Emu(y + i*700000), Emu(600000), Emu(550000),
                   GREEN, "→", WHITE, 28, True)
    add_text_box(s, Emu(1600000), Emu(y + i*700000 + 50000), Emu(10000000), Emu(550000), [
        (ans, {'size': 24, 'color': NAVY}),
    ])

# 明日からできること
add_filled_box(s, Emu(700000), Emu(5000000), Emu(11000000), Emu(450000), TEAL,
               "≪明日からできること≫", WHITE, 20, True, align=PP_ALIGN.LEFT)
add_outline_box(s, Emu(700000), Emu(5470000), Emu(11000000), Emu(550000),
                border_color=TEAL, line_width=2.0)
add_text_box(s, Emu(900000), Emu(5520000), Emu(10600000), Emu(450000), [
    ("次の単元の指導計画に、「記録」を入れる欄を1列足す",
     {'size': 20, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
])

# キーメッセージ
add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("【キーメッセージ①】評価は「ためる」のではなく「使う」もの",
     {'size': 10, 'color': DEEP_BLUE, 'bold': True, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド17: Part3 セクション扉
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 17)

add_filled_box(s, 0, Emu(1500000), SW, Emu(800000), DEEP_BLUE,
               "Part 3", WHITE, 44, True)
add_filled_box(s, 0, Emu(2400000), SW, Emu(700000), LIGHT_BLUE,
               "校長の問い❷", WHITE, 32, True)
add_filled_box(s, 0, Emu(3200000), SW, Emu(1000000), TEAL,
               "── どんな姿を見取るか ──", WHITE, 36, True)
add_text_box(s, Emu(700000), Emu(4400000), Emu(11000000), Emu(500000), [
    ("（9分 ── 本日の核）", {'size': 28, 'color': GREEN, 'bold': True, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド18: 観点別評価の基本構造
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 18)
add_section_marker(s, "Part 3　3観点の整理")
add_title_box(s, "各教科の評価 ── 3つの観点")

points = [
    ("知識・技能", "何を理解しているか／何ができるか", LIGHT_BLUE),
    ("思考・判断・表現", "理解していること・できることをどう使うか", DEEP_BLUE),
    ("主体的に学習に取り組む態度", "どのように学び続けようとしているか", TEAL),
]
y_base = 1900000
for i, (obs, desc, color) in enumerate(points):
    y = y_base + i * 1200000
    add_filled_box(s, Emu(700000), Emu(y), Emu(4500000), Emu(1000000), color,
                   obs, WHITE, 24, True)
    add_outline_box(s, Emu(5300000), Emu(y), Emu(6400000), Emu(1000000),
                    border_color=color, line_width=2.0)
    add_text_box(s, Emu(5500000), Emu(y + 200000), Emu(6000000), Emu(800000), [
        (desc, {'size': 22, 'color': NAVY}),
    ], line_spacing=1.3)

# ========================================================================
# スライド19: 「知識・技能」の見取りどころ
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 19)
add_section_marker(s, "Part 3　知識・技能")
add_title_box(s, "「知識・技能」を見取る")

add_filled_box(s, Emu(700000), Emu(1850000), Emu(11000000), Emu(700000), LIGHT_BLUE,
               "★ 既有の知識と関連付けて使えるか／概念として理解しているか",
               WHITE, 22, True, align=PP_ALIGN.LEFT)

add_filled_box(s, Emu(700000), Emu(2700000), Emu(11000000), Emu(400000), GRAY_LT,
               "【評価の場面】", NAVY, 20, True, align=PP_ALIGN.LEFT)
add_outline_box(s, Emu(700000), Emu(3100000), Emu(11000000), Emu(2400000),
                border_color=LIGHT_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(3250000), Emu(10600000), Emu(2200000), [
    ("○ ペーパーテスト", {'size': 24, 'color': NAVY}),
    ("　 → 事実的知識と概念的理解のバランスを問う", {'size': 18, 'color': DARK_GRAY}),
    ("○ 知識・技能を実際に使う場面の学習活動", {'size': 24, 'color': NAVY}),
    ("　 → 観察・実験・式やグラフでの表現", {'size': 18, 'color': DARK_GRAY}),
    ("○ 説明させる・例を挙げさせる発問", {'size': 24, 'color': NAVY}),
    ("　 → 「概念として理解しているか」を見るのが鍵", {'size': 18, 'color': DARK_GRAY}),
], line_spacing=1.3)

add_filled_box(s, Emu(700000), Emu(5700000), Emu(11000000), Emu(600000), GREEN,
               "テストだけが「知識・技能」の評価ではない", WHITE, 24, True)

# ========================================================================
# スライド20: 「思考・判断・表現」の見取りどころ
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 20)
add_section_marker(s, "Part 3　思考・判断・表現")
add_title_box(s, "「思考・判断・表現」を見取る")

add_filled_box(s, Emu(700000), Emu(1850000), Emu(11000000), Emu(700000), DEEP_BLUE,
               "★ 知識・技能を使って課題を解決できるか／過程を表現できるか",
               WHITE, 22, True, align=PP_ALIGN.LEFT)

add_filled_box(s, Emu(700000), Emu(2700000), Emu(11000000), Emu(400000), GRAY_LT,
               "【評価の場面】", NAVY, 20, True, align=PP_ALIGN.LEFT)
add_outline_box(s, Emu(700000), Emu(3100000), Emu(11000000), Emu(2400000),
                border_color=DEEP_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(3250000), Emu(10600000), Emu(2200000), [
    ("○ 論述・レポート", {'size': 24, 'color': NAVY}),
    ("○ グループでの話合い", {'size': 24, 'color': NAVY}),
    ("○ 作品の制作・表現活動", {'size': 24, 'color': NAVY}),
    ("○ ポートフォリオの活用", {'size': 24, 'color': NAVY}),
], line_spacing=1.4)

add_filled_box(s, Emu(700000), Emu(5700000), Emu(11000000), Emu(600000), GREEN,
               "「過程」を見る ── 結果だけではない", WHITE, 24, True)

# ========================================================================
# スライド21: 「主体的に学習に取り組む態度」の2側面 +キーワード
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 21)
add_section_marker(s, "Part 3　主体的に学習に取り組む態度")
add_title_box(s, "「主体的に学習に取り組む態度」── 2つの側面", size=32)

# 左：粘り強く
add_filled_box(s, Emu(700000), Emu(1900000), Emu(5400000), Emu(700000), LIGHT_BLUE,
               "① 粘り強く取り組む", WHITE, 24, True)
add_outline_box(s, Emu(700000), Emu(2600000), Emu(5400000), Emu(2400000),
                border_color=LIGHT_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(2750000), Emu(5000000), Emu(2200000), [
    ("【姿のキーワード】", {'size': 18, 'color': DARK_GRAY, 'bold': True}),
    ("「積極的に」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
    ("「進んで」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
    ("「最後まで」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
    ("「あきらめずに」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
], line_spacing=1.3)

# 右：自ら調整
add_filled_box(s, Emu(6300000), Emu(1900000), Emu(5400000), Emu(700000), DEEP_BLUE,
               "② 自ら学習を調整する", WHITE, 24, True)
add_outline_box(s, Emu(6300000), Emu(2600000), Emu(5400000), Emu(2400000),
                border_color=DEEP_BLUE, line_width=2.0)
add_text_box(s, Emu(6500000), Emu(2750000), Emu(5000000), Emu(2200000), [
    ("【姿のキーワード】", {'size': 18, 'color': DARK_GRAY, 'bold': True}),
    ("「見通しをもって」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
    ("「学習課題に沿って」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
    ("「振り返って」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
    ("「見直して」", {'size': 22, 'bold': True, 'color': DEEP_BLUE}),
], line_spacing=1.3)

add_filled_box(s, Emu(700000), Emu(5200000), Emu(11000000), Emu(400000), GREEN,
               "①と②は別々ではなく、相互に関わり合う",
               WHITE, 18, True)
add_filled_box(s, Emu(700000), Emu(5650000), Emu(11000000), Emu(400000), NAVY,
               "①A・②C という姿は一般的ではない", WHITE, 18, True)

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 中央教育審議会 H31.1報告 ／ 国立教育政策研究所",
     {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド22: Bと判断する姿（数学の具体例）
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 22)
add_section_marker(s, "Part 3　Bと判断する姿")
add_title_box(s, "「Bと判断する姿」を具体化する", size=32)

# 上：共通の姿
add_filled_box(s, Emu(700000), Emu(1750000), Emu(11000000), Emu(500000), TEAL,
               "教科横断で見える共通の姿", WHITE, 22, True, align=PP_ALIGN.LEFT)
add_outline_box(s, Emu(700000), Emu(2250000), Emu(11000000), Emu(600000),
                border_color=TEAL, line_width=2.0)
add_text_box(s, Emu(900000), Emu(2350000), Emu(10600000), Emu(500000), [
    ("○ 立ち止まる　 ○ 振り返る　 ○ やり直す　 ○ 試行錯誤する",
     {'size': 22, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
])

# 下：数学の具体例
add_filled_box(s, Emu(700000), Emu(3000000), Emu(11000000), Emu(500000), DEEP_BLUE,
               "具体例： 数学「二次方程式」── 学習シートに「気を付けるポイント」を記述",
               WHITE, 18, True, align=PP_ALIGN.LEFT)

# B
add_filled_box(s, Emu(700000), Emu(3550000), Emu(5400000), Emu(500000), LIGHT_BLUE,
               "▼ おおむね満足（B）", WHITE, 20, True)
add_outline_box(s, Emu(700000), Emu(4050000), Emu(5400000), Emu(1700000),
                border_color=LIGHT_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(4150000), Emu(5000000), Emu(1500000), [
    ("ポイントが書かれている", {'size': 18, 'color': NAVY, 'bold': True}),
    ("", {'size': 6}),
    ("例：別の文字に置き換えて", {'size': 17, 'color': DARK_GRAY}),
    ("解く方法を使えるように", {'size': 17, 'color': DARK_GRAY}),
    ("する", {'size': 17, 'color': DARK_GRAY}),
], line_spacing=1.2)

# A
add_filled_box(s, Emu(6300000), Emu(3550000), Emu(5400000), Emu(500000), GREEN,
               "▼ 十分満足（A）", WHITE, 20, True)
add_outline_box(s, Emu(6300000), Emu(4050000), Emu(5400000), Emu(1700000),
                border_color=GREEN, line_width=2.0)
add_text_box(s, Emu(6500000), Emu(4150000), Emu(5000000), Emu(1500000), [
    ("ポイント＋理由が書かれている", {'size': 18, 'color': NAVY, 'bold': True}),
    ("", {'size': 6}),
    ("例：式を見ずに展開すると時間も", {'size': 15, 'color': DARK_GRAY}),
    ("間違いも増えるから、置き換えを", {'size': 15, 'color': DARK_GRAY}),
    ("使えるようにする", {'size': 15, 'color': DARK_GRAY}),
], line_spacing=1.2)

add_filled_box(s, Emu(700000), Emu(5770000), Emu(11000000), Emu(340000), NAVY,
               "同じ気付きでも「理由まで書けるか」でAとBが分かれる",
               WHITE, 18, True)

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 国立教育政策研究所「指導と評価の一体化のための参考資料」（数学）",
     {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド23: 「指導してから評価する」── 本日最重要メッセージ
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 23)
add_section_marker(s, "Part 3　最重要メッセージ")
add_title_box(s, "《本日の最重要メッセージ》", size=28, color=GREEN)

# 大きく中央
add_filled_box(s, Emu(700000), Emu(1900000), Emu(11000000), Emu(900000), NAVY,
               "「主体的に学習に取り組む態度」は", WHITE, 28, True)
add_filled_box(s, Emu(700000), Emu(2800000), Emu(11000000), Emu(900000), GREEN,
               "“指導してから評価する”", WHITE, 36, True)

# ×と○
add_filled_box(s, Emu(700000), Emu(4000000), Emu(900000), Emu(900000),
               RGBColor(0xC0, 0x40, 0x40), "×", WHITE, 44, True)
add_outline_box(s, Emu(1700000), Emu(4000000), Emu(10000000), Emu(900000),
                border_color=RGBColor(0xC0, 0x40, 0x40), line_width=2.0)
add_text_box(s, Emu(1900000), Emu(4080000), Emu(9700000), Emu(800000), [
    ("「個人任せ」で見取ろうとする → 育たない・評価できない",
     {'size': 22, 'color': NAVY}),
], line_spacing=1.2)

add_filled_box(s, Emu(700000), Emu(5100000), Emu(900000), Emu(900000),
               GREEN, "○", WHITE, 44, True)
add_outline_box(s, Emu(1700000), Emu(5100000), Emu(10000000), Emu(900000),
                border_color=GREEN, line_width=2.0)
add_text_box(s, Emu(1900000), Emu(5180000), Emu(9700000), Emu(800000), [
    ("「学び方」を指導してから、その姿を見取る → 育つ・評価できる",
     {'size': 22, 'bold': True, 'color': NAVY}),
], line_spacing=1.2)

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 国立教育政策研究所「指導と評価の一体化のための参考資料」",
     {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド24: 「学び方」を指導する9つの働きかけ
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 24)
add_section_marker(s, "Part 3　教師の働きかけ")
add_title_box(s, "主体的な学びを引き出す ── 9つの働きかけ", size=30)

# 3×3グリッド
items = [
    "①既習事項を振り返らせる",
    "②具体物で引きつける",
    "③「明らかにしたい」課題設定",
    "④解決の方向性に見通しをもたせる",
    "⑤生徒の思考を見守る",
    "⑥生徒の思考に即して展開",
    "⑦生徒の考えを生かしてまとめる",
    "⑧その日の学びを振り返らせる",
    "⑨新たな学びに意識を向けさせる",
]
x_base = 700000
y_base = 1900000
col_w = 3650000
row_h = 1100000
for i, item in enumerate(items):
    col = i % 3
    row = i // 3
    x = x_base + col * (col_w + 100000)
    y = y_base + row * (row_h + 100000)
    color = [LIGHT_BLUE, DEEP_BLUE, TEAL][row]
    add_filled_box(s, Emu(x), Emu(y), Emu(col_w), Emu(row_h), color,
                   item, WHITE, 18, True)

add_text_box(s, Emu(700000), Emu(5500000), Emu(11000000), Emu(400000), [
    ("明日の授業で、9つのうち1つを意識する。それだけで授業が変わる。",
     {'size': 22, 'bold': True, 'color': GREEN, 'align': PP_ALIGN.CENTER}),
])

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 文部科学省 国立教育政策研究所（2020）",
     {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド25: 行動観察 ── 何を見るか
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 25)
add_section_marker(s, "Part 3　行動観察")
add_title_box(s, "行動観察 ── 何を見るか", size=32)

# 粘り強さ
add_filled_box(s, Emu(700000), Emu(1850000), Emu(5400000), Emu(600000), LIGHT_BLUE,
               "【粘り強く取り組む姿】", WHITE, 22, True)
add_outline_box(s, Emu(700000), Emu(2450000), Emu(5400000), Emu(2700000),
                border_color=LIGHT_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(2570000), Emu(5000000), Emu(2500000), [
    ("○ 前時のノートを見返している", {'size': 18, 'color': NAVY}),
    ("", {'size': 4}),
    ("○ 別の方法での解決を試みる", {'size': 18, 'color': NAVY}),
    ("", {'size': 4}),
    ("○ 図や言葉を書き加えて", {'size': 18, 'color': NAVY}),
    ("　 理解しようとしている", {'size': 18, 'color': NAVY}),
    ("", {'size': 4}),
    ("○ 説明を念頭に書き加える", {'size': 18, 'color': NAVY}),
], line_spacing=1.2)

# 自己調整
add_filled_box(s, Emu(6300000), Emu(1850000), Emu(5400000), Emu(600000), DEEP_BLUE,
               "【自己調整する姿】", WHITE, 22, True)
add_outline_box(s, Emu(6300000), Emu(2450000), Emu(5400000), Emu(2700000),
                border_color=DEEP_BLUE, line_width=2.0)
add_text_box(s, Emu(6500000), Emu(2570000), Emu(5000000), Emu(2500000), [
    ("○ 解決方法を振り返り、修正する", {'size': 18, 'color': NAVY}),
    ("", {'size': 4}),
    ("○ よりよい方法を探そうとする", {'size': 18, 'color': NAVY}),
    ("", {'size': 4}),
    ("○ 振り返りシートに気付きを書く", {'size': 18, 'color': NAVY}),
    ("", {'size': 4}),
    ("○ 学習計画を立て直す", {'size': 18, 'color': NAVY}),
], line_spacing=1.2)

add_filled_box(s, Emu(700000), Emu(5350000), Emu(11000000), Emu(800000), GREEN,
               "全員を毎回見なくてよい ── 気になる3〜5人を単元で複数回見る",
               WHITE, 22, True)

# ========================================================================
# スライド26: 見取りの留意点
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 26)
add_section_marker(s, "Part 3　留意点")
add_title_box(s, "見取るときの留意点")

# ×項目
add_filled_box(s, Emu(700000), Emu(1850000), Emu(5400000), Emu(500000),
               RGBColor(0xC0, 0x40, 0x40), "× こんな見取りはNG", WHITE, 20, True)
add_outline_box(s, Emu(700000), Emu(2350000), Emu(5400000), Emu(2400000),
                border_color=RGBColor(0xC0, 0x40, 0x40), line_width=2.0)
add_text_box(s, Emu(900000), Emu(2470000), Emu(5000000), Emu(2200000), [
    ("× 挙手の回数・発言の多寡", {'size': 18, 'color': NAVY}),
    ("× ノート提出の有無だけ", {'size': 18, 'color': NAVY}),
    ("× 性格・態度（外向き/真面目）", {'size': 18, 'color': NAVY}),
    ("× 「個人任せ」で態度を見る", {'size': 18, 'color': NAVY, 'bold': True}),
], line_spacing=1.4)

# ○項目
add_filled_box(s, Emu(6300000), Emu(1850000), Emu(5400000), Emu(500000), GREEN,
               "○ こう見取る", WHITE, 20, True)
add_outline_box(s, Emu(6300000), Emu(2350000), Emu(5400000), Emu(2400000),
                border_color=GREEN, line_width=2.0)
add_text_box(s, Emu(6500000), Emu(2470000), Emu(5000000), Emu(2200000), [
    ("○ 学習目標に向かう「姿」を見る", {'size': 18, 'color': NAVY}),
    ("○ 複数の場面・方法で確認する", {'size': 18, 'color': NAVY}),
    ("○ 「学び方」を指導した上で見る", {'size': 18, 'color': NAVY, 'bold': True}),
    ("○ ばらつき（知CC・態度A）に気づく", {'size': 18, 'color': NAVY}),
], line_spacing=1.4)

add_filled_box(s, Emu(700000), Emu(4950000), Emu(11000000), Emu(700000), DEEP_BLUE,
               "考えにくい評価 ── 知CC・思CC・態度A の組合せに要注意",
               WHITE, 22, True)
add_text_box(s, Emu(700000), Emu(5750000), Emu(11000000), Emu(400000), [
    ("【キーメッセージ②】生徒の姿は「単元の中」で、「学び方を指導した上で」見取る",
     {'size': 11, 'color': DEEP_BLUE, 'bold': True, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド27: Part4 セクション扉
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 27)

add_filled_box(s, 0, Emu(1800000), SW, Emu(800000), DEEP_BLUE,
               "Part 4", WHITE, 44, True)
add_filled_box(s, 0, Emu(2700000), SW, Emu(700000), LIGHT_BLUE,
               "校長の問い❸", WHITE, 32, True)
add_filled_box(s, 0, Emu(3500000), SW, Emu(1000000), TEAL,
               "── どう授業改善につなげるか ──", WHITE, 32, True)
add_text_box(s, Emu(700000), Emu(4700000), Emu(11000000), Emu(500000), [
    ("（4分）", {'size': 28, 'color': DARK_GRAY, 'align': PP_ALIGN.CENTER}),
])

# ========================================================================
# スライド28: 評価から授業改善へのサイクル
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 28)
add_section_marker(s, "Part 4　改善サイクル")
add_title_box(s, "評価 → 授業改善 のサイクル", size=32)

# 6つのステップを縦に
steps = [
    ("①", "単元・本時の目標を明確化", LIGHT_BLUE),
    ("②", "評価規準・場面を計画", LIGHT_BLUE),
    ("③", "授業実施 ＋ 形成的評価", DEEP_BLUE),
    ("④", "「指導に生かす」── 次時の指導を修正", DEEP_BLUE),
    ("⑤", "単元末「記録に残す」── 評定へ", TEAL),
    ("⑥", "次の単元・学年へ引き継ぐ", GREEN),
]
y_base = 1750000
row_h = 650000
for i, (num, txt, color) in enumerate(steps):
    y = y_base + i * row_h
    add_filled_box(s, Emu(700000), Emu(y), Emu(800000), Emu(550000), color,
                   num, WHITE, 28, True)
    add_outline_box(s, Emu(1600000), Emu(y), Emu(10100000), Emu(550000),
                    border_color=color, line_width=1.5)
    add_text_box(s, Emu(1800000), Emu(y + 60000), Emu(9700000), Emu(450000), [
        (txt, {'size': 22, 'color': NAVY, 'bold': True}),
    ])

add_filled_box(s, Emu(700000), Emu(5700000), Emu(11000000), Emu(600000), NAVY,
               "評価は単元の終わりではなく、次の出発点", WHITE, 24, True)

# ========================================================================
# スライド29: 組織で支える信頼性
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 29)
add_section_marker(s, "Part 4　組織で支える")
add_title_box(s, "妥当性は「個人」で、信頼性は「組織」で", size=30)

# 教科会
add_filled_box(s, Emu(700000), Emu(1850000), Emu(5400000), Emu(500000), DEEP_BLUE,
               "【教科会で揃えるべきもの】", WHITE, 20, True)
add_outline_box(s, Emu(700000), Emu(2350000), Emu(5400000), Emu(2500000),
                border_color=DEEP_BLUE, line_width=2.0)
add_text_box(s, Emu(900000), Emu(2470000), Emu(5000000), Emu(2300000), [
    ("○ 評価規準（Bの姿の言葉）", {'size': 18, 'color': NAVY}),
    ("○ 評価場面・評価方法", {'size': 18, 'color': NAVY}),
    ("○ 評定の総括方法", {'size': 18, 'color': NAVY}),
    ("○ 保護者への説明の仕方", {'size': 18, 'color': NAVY}),
], line_spacing=1.4)

# 管理職
add_filled_box(s, Emu(6300000), Emu(1850000), Emu(5400000), Emu(500000), TEAL,
               "【管理職・主任の役割】", WHITE, 20, True)
add_outline_box(s, Emu(6300000), Emu(2350000), Emu(5400000), Emu(2500000),
                border_color=TEAL, line_width=2.0)
add_text_box(s, Emu(6500000), Emu(2470000), Emu(5000000), Emu(2300000), [
    ("○ 教科会の機能化", {'size': 18, 'color': NAVY}),
    ("○ 評定のチェック体制", {'size': 18, 'color': NAVY}),
    ("○ 保護者対応の組織化", {'size': 18, 'color': NAVY}),
    ("○ 「特異なパターン」確認", {'size': 18, 'color': NAVY}),
], line_spacing=1.4)

add_filled_box(s, Emu(700000), Emu(5050000), Emu(11000000), Emu(700000), GREEN,
               "東京都が示す特異パターン例： AAA→「5/4」でない、ABCの偏り 等",
               WHITE, 18, True)
add_text_box(s, Emu(700000), Emu(5780000), Emu(11000000), Emu(300000), [
    ("【キーメッセージ③】評価の信頼性は「組織」で担保する",
     {'size': 11, 'color': DEEP_BLUE, 'bold': True, 'align': PP_ALIGN.CENTER}),
])

add_text_box(s, Emu(700000), Emu(6120000), Emu(11000000), Emu(214000), [
    ("出典： 東京都教育委員会「指導と評価の一体化を目指して」",
     {'size': 10, 'color': DARK_GRAY, 'align': PP_ALIGN.RIGHT}),
])

# ========================================================================
# スライド30: まとめ／協議会への橋渡し
# ========================================================================
s = add_blank_slide()
add_bottom_bar(s)
add_page_number(s, 30)
add_section_marker(s, "本日のまとめ")
add_title_box(s, "本日のまとめ", size=36)

# 3つのキーメッセージ
y = 1700000
messages = [
    ("①", "評価は「ためる」のではなく「使う」もの", DEEP_BLUE),
    ("②", "生徒の姿は「単元の中」で、「学び方を指導した上で」見取る", LIGHT_BLUE),
    ("③", "評価の信頼性は「組織」で担保する", TEAL),
]
for i, (num, msg, color) in enumerate(messages):
    add_filled_box(s, Emu(700000), Emu(y + i*650000), Emu(700000), Emu(550000),
                   color, num, WHITE, 28, True)
    add_outline_box(s, Emu(1500000), Emu(y + i*650000), Emu(10200000), Emu(550000),
                    border_color=color, line_width=2.0)
    add_text_box(s, Emu(1700000), Emu(y + i*650000 + 70000), Emu(9800000), Emu(450000), [
        (msg, {'size': 20, 'bold': True, 'color': NAVY}),
    ])

# 明日からできること
add_filled_box(s, Emu(700000), Emu(3800000), Emu(11000000), Emu(450000), GREEN,
               "≪明日からできること（1つ選んでください）≫", WHITE, 20, True)
add_outline_box(s, Emu(700000), Emu(4250000), Emu(11000000), Emu(1200000),
                border_color=GREEN, line_width=2.0)
add_text_box(s, Emu(900000), Emu(4320000), Emu(10600000), Emu(1100000), [
    ("□ 次の単元の指導計画に「記録」欄を1列足す", {'size': 17, 'color': NAVY}),
    ("□ 教科会で「Bの姿」を1つ言語化して揃える", {'size': 17, 'color': NAVY}),
    ("□ 9つの働きかけから1つ選んで使う", {'size': 17, 'color': NAVY}),
    ("□ 振り返りシートを1回入れてみる", {'size': 17, 'color': NAVY}),
], line_spacing=1.2)

# 協議会へ
add_filled_box(s, Emu(700000), Emu(5500000), Emu(11000000), Emu(380000), NAVY,
               "── このあと協議会で「セルフチェックシート」をお使いください ──",
               WHITE, 14, True)
add_text_box(s, Emu(700000), Emu(5910000), Emu(11000000), Emu(280000), [
    ("ご清聴ありがとうございました",
     {'size': 14, 'color': DEEP_BLUE, 'bold': True, 'align': PP_ALIGN.CENTER}),
])

# 保存
prs.save("/tmp/output/講義スライド_石神井南中_20260603.pptx")
print("OK: PPTX created")
print("Slides:", len(prs.slides))
