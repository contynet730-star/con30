"""
石神井南中学校 6/3 校内研究 講義スライド生成スクリプト
参考: 練馬中PPTX のテイスト
- フォント: メイリオ
- 文字大きく（28-40pt）
- 枠の色: 水色
- 構成: 上部章バー + 全幅サブ見出し + 大きな本文枠
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

# ========== 色定義 ==========
WATER_BLUE = RGBColor(0x5B, 0xC0, 0xDE)    # 水色（枠の主色）
LIGHT_WATER = RGBColor(0xD6, 0xEE, 0xFA)   # 薄い水色（背景）
DEEP_WATER = RGBColor(0x2A, 0x96, 0xC2)    # 濃い水色
NAVY = RGBColor(0x1F, 0x3A, 0x6E)          # 濃紺
ORANGE_KW = RGBColor(0xEB, 0x6C, 0x15)     # オレンジ（強調）
RED = RGBColor(0xC0, 0x40, 0x40)
GREEN = RGBColor(0x07, 0xA9, 0x73)
YELLOW = RGBColor(0xFF, 0xE0, 0x66)        # 薄い黄
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
DARK_GRAY = RGBColor(0x40, 0x40, 0x40)

# ========== フォント ==========
F_MAIN = "メイリオ"

prs = Presentation()
prs.slide_width = Emu(12192000)
prs.slide_height = Emu(6858000)
SW = prs.slide_width
SH = prs.slide_height

# ========== ヘルパー関数 ==========
def add_blank_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])

def add_chapter_bar(slide, chapter_text):
    """上部章バー（水色背景、白文字、全幅、高さ720k）"""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Emu(720000))
    bar.fill.solid(); bar.fill.fore_color.rgb = DEEP_WATER
    bar.line.fill.background()
    tf = bar.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Emu(300000); tf.margin_right = Emu(100000)
    tf.margin_top = Emu(20000); tf.margin_bottom = Emu(20000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = chapter_text
    r.font.size = Pt(32)
    r.font.bold = True
    r.font.color.rgb = WHITE
    r.font.name = F_MAIN

def add_sub_heading(slide, text, width_full=True):
    """全幅サブ見出し（白背景、下部に水色アンダーバー、左寄せ、32pt）"""
    w = SW if width_full else Emu(9000000)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(830000), w, Emu(700000))
    bar.fill.solid(); bar.fill.fore_color.rgb = LIGHT_WATER
    bar.line.fill.background()
    tf = bar.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Emu(300000); tf.margin_right = Emu(200000)
    tf.margin_top = Emu(20000); tf.margin_bottom = Emu(20000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.size = Pt(32)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r.font.name = F_MAIN
    # 水色下線
    ul = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(1530000), w, Emu(40000))
    ul.fill.solid(); ul.fill.fore_color.rgb = WATER_BLUE
    ul.line.fill.background()

def add_page_number(slide, num):
    """右下ページ番号"""
    tb = slide.shapes.add_textbox(Emu(11250000), Emu(6420000), Emu(800000), Emu(360000))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = Emu(40000); tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = str(num)
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = DARK_GRAY
    r.font.name = F_MAIN

def add_source_line(slide, text):
    """下部出典テキスト（小さく、右寄せ）"""
    tb = slide.shapes.add_textbox(Emu(300000), Emu(6430000), Emu(10800000), Emu(360000))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Emu(50000); tf.margin_right = Emu(50000)
    tf.margin_top = Emu(10000); tf.margin_bottom = Emu(10000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "出典：" + text
    r.font.size = Pt(13)
    r.font.bold = False
    r.font.color.rgb = DARK_GRAY
    r.font.name = F_MAIN

def add_body_box(slide, left, top, width, height, lines,
                 default_size=28, default_color=BLACK, line_spacing=1.4,
                 align=PP_ALIGN.LEFT, font=F_MAIN, fill_color=None,
                 border_color=WATER_BLUE, border_width=3.0,
                 vertical_anchor=MSO_ANCHOR.TOP, default_bold=True):
    """本文枠（水色枠）。
       linesの各要素は str / (text, opts) / [(text, opts), ...]"""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill_color:
        box.fill.solid(); box.fill.fore_color.rgb = fill_color
    else:
        box.fill.solid(); box.fill.fore_color.rgb = WHITE
    if border_color:
        box.line.color.rgb = border_color
        box.line.width = Pt(border_width)
    else:
        box.line.fill.background()
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Emu(160000); tf.margin_right = Emu(160000)
    tf.margin_top = Emu(100000); tf.margin_bottom = Emu(100000)
    tf.vertical_anchor = vertical_anchor

    for i, item in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if isinstance(item, list):
            for j, run_item in enumerate(item):
                if isinstance(run_item, tuple):
                    rt, ropts = run_item
                else:
                    rt, ropts = run_item, {}
                if j == 0:
                    p.alignment = ropts.get('align', align)
                    p.line_spacing = ropts.get('line_spacing', line_spacing)
                r = p.add_run()
                r.text = rt
                r.font.size = Pt(ropts.get('size', default_size))
                r.font.bold = ropts.get('bold', default_bold)
                r.font.color.rgb = ropts.get('color', default_color)
                r.font.name = ropts.get('font', font)
        else:
            if isinstance(item, tuple):
                text, opts = item
            else:
                text, opts = item, {}
            p.alignment = opts.get('align', align)
            p.line_spacing = opts.get('line_spacing', line_spacing)
            if 'space_before' in opts:
                p.space_before = Pt(opts['space_before'])
            r = p.add_run()
            r.text = text
            r.font.size = Pt(opts.get('size', default_size))
            r.font.bold = opts.get('bold', default_bold)
            r.font.color.rgb = opts.get('color', default_color)
            r.font.name = opts.get('font', font)

def add_plain_text(slide, left, top, width, height, lines,
                   default_size=24, default_color=NAVY, line_spacing=1.3,
                   align=PP_ALIGN.LEFT, font=F_MAIN,
                   vertical_anchor=MSO_ANCHOR.TOP, default_bold=False):
    """枠なしテキスト"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Emu(50000); tf.margin_right = Emu(50000)
    tf.margin_top = Emu(30000); tf.margin_bottom = Emu(30000)
    tf.vertical_anchor = vertical_anchor
    for i, item in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if isinstance(item, list):
            for j, run_item in enumerate(item):
                if isinstance(run_item, tuple):
                    rt, ropts = run_item
                else:
                    rt, ropts = run_item, {}
                if j == 0:
                    p.alignment = ropts.get('align', align)
                    p.line_spacing = ropts.get('line_spacing', line_spacing)
                r = p.add_run()
                r.text = rt
                r.font.size = Pt(ropts.get('size', default_size))
                r.font.bold = ropts.get('bold', default_bold)
                r.font.color.rgb = ropts.get('color', default_color)
                r.font.name = ropts.get('font', font)
        else:
            if isinstance(item, tuple):
                text, opts = item
            else:
                text, opts = item, {}
            p.alignment = opts.get('align', align)
            p.line_spacing = opts.get('line_spacing', line_spacing)
            r = p.add_run()
            r.text = text
            r.font.size = Pt(opts.get('size', default_size))
            r.font.bold = opts.get('bold', default_bold)
            r.font.color.rgb = opts.get('color', default_color)
            r.font.name = opts.get('font', font)

# ============================================================
# スライド1: 表紙
# ============================================================
s = add_blank_slide()
# 上部水色帯
top_band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Emu(900000))
top_band.fill.solid(); top_band.fill.fore_color.rgb = DEEP_WATER
top_band.line.fill.background()
# 下部水色帯
bot_band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(6300000), SW, Emu(558000))
bot_band.fill.solid(); bot_band.fill.fore_color.rgb = DEEP_WATER
bot_band.line.fill.background()
# 細い水色ライン
ln1 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(910000), SW, Emu(40000))
ln1.fill.solid(); ln1.fill.fore_color.rgb = WATER_BLUE
ln1.line.fill.background()
ln2 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Emu(6250000), SW, Emu(40000))
ln2.fill.solid(); ln2.fill.fore_color.rgb = WATER_BLUE
ln2.line.fill.background()

add_plain_text(s, Emu(700000), Emu(180000), Emu(10800000), Emu(540000), [
    ("令和８年度　石神井南中学校　校内研究", {'size': 24, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.LEFT}),
], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE)

# メインタイトル
add_plain_text(s, Emu(400000), Emu(1700000), Emu(11400000), Emu(1400000), [
    ("指導と評価の一体化", {'size': 60, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=60, vertical_anchor=MSO_ANCHOR.MIDDLE)

# 副題ボックス（水色枠）
add_body_box(s, Emu(1800000), Emu(3250000), Emu(8600000), Emu(900000), [
    ("生徒の姿で語る授業改善", {'size': 36, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
], default_size=36, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=4.0)

add_plain_text(s, Emu(400000), Emu(4350000), Emu(11400000), Emu(600000), [
    ("〜明日、何を変えるか〜", {'size': 24, 'color': DARK_GRAY, 'align': PP_ALIGN.CENTER}),
], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE)

# 日付・講師
add_plain_text(s, Emu(400000), Emu(5050000), Emu(11400000), Emu(440000), [
    ("令和８年６月３日（水）", {'size': 24, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_plain_text(s, Emu(400000), Emu(5520000), Emu(11400000), Emu(380000), [
    ("練馬区教育委員会　指導主事", {'size': 18, 'color': DARK_GRAY, 'align': PP_ALIGN.CENTER}),
], default_size=18, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_plain_text(s, Emu(400000), Emu(5900000), Emu(11400000), Emu(340000), [
    ("紺多　章一郎", {'size': 22, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE)

# ============================================================
# スライド2: ウォームアップ
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "ウォームアップ")
add_sub_heading(s, "この数字は何を表しているでしょう？")
add_page_number(s, 2)

# 3列で大きく数字
add_body_box(s, Emu(400000), Emu(1850000), Emu(3700000), Emu(1500000), [
    ("47.2", {'size': 50, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
    ("30.5　／　22.3", {'size': 28, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=4.0)
add_plain_text(s, Emu(400000), Emu(3400000), Emu(3700000), Emu(400000), [
    ("A中学校", {'size': 22, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_body_box(s, Emu(4250000), Emu(1850000), Emu(3700000), Emu(1500000), [
    ("8.1", {'size': 50, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
    ("65.8　／　26.1", {'size': 28, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=4.0)
add_plain_text(s, Emu(4250000), Emu(3400000), Emu(3700000), Emu(400000), [
    ("B中学校", {'size': 22, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_body_box(s, Emu(8100000), Emu(1850000), Emu(3700000), Emu(1500000), [
    ("12.4", {'size': 50, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
    ("50.3　／　37.3", {'size': 28, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=4.0)
add_plain_text(s, Emu(8100000), Emu(3400000), Emu(3700000), Emu(400000), [
    ("C中学校", {'size': 22, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE)

# 答え
add_body_box(s, Emu(400000), Emu(4150000), Emu(11400000), Emu(900000), [
    ("ヒント：ある教科・ある観点の評定割合（％）",
     {'size': 26, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=3.0)

add_body_box(s, Emu(400000), Emu(5150000), Emu(11400000), Emu(1100000), [
    ("答え：区内3校の同一教科「主体的に学習に取り組む態度」の評定割合",
     {'size': 22, 'bold': True, 'color': ORANGE_KW, 'align': PP_ALIGN.CENTER}),
    ("同じ教科でも、学校で大きく異なる現実",
     {'size': 24, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

# ============================================================
# スライド3: 本日の流れ
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "本日の流れ（30分）")
add_sub_heading(s, "Part1〜Part5の構成")
add_page_number(s, 3)

agenda = [
    ("Part1", "評価をめぐる「いま」", "5分"),
    ("Part2", "石神井南中の現状と課題", "5分"),
    ("Part3", "「指導と評価の一体化」の核", "13分"),
    ("Part4", "明日から動かす", "5分"),
    ("Part5", "協議会へのつなぎ", "2分"),
]
y = 1850000
for i, (part, title, time) in enumerate(agenda):
    yy = y + i * 870000
    # Partラベル
    add_body_box(s, Emu(400000), Emu(yy), Emu(1700000), Emu(750000), [
        (part, {'size': 28, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
    ], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
    # タイトル
    add_body_box(s, Emu(2200000), Emu(yy), Emu(7800000), Emu(750000), [
        (title, {'size': 28, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.LEFT}),
    ], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=2.0)
    # 時間
    add_body_box(s, Emu(10100000), Emu(yy), Emu(1700000), Emu(750000), [
        (time, {'size': 24, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
    ], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=2.0)

# ============================================================
# スライド4: Part1 章扉
# ============================================================
s = add_blank_slide()
add_page_number(s, 4)
# 中央の大ボックス
add_body_box(s, Emu(1000000), Emu(2100000), Emu(10200000), Emu(2700000), [
    ("Part 1", {'size': 36, 'bold': True, 'color': WATER_BLUE, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 14}),
    ("評価をめぐる「いま」", {'size': 52, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 8}),
    ("─ 学習評価とは何か。なぜ今、見直されているのか ─",
     {'size': 20, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=5.0, fill_color=LIGHT_WATER)

# ============================================================
# スライド5: 「学習評価」と「評定」
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 1　評価をめぐる「いま」")
add_sub_heading(s, "「学習評価」と「評定」とは")
add_page_number(s, 5)

add_plain_text(s, Emu(400000), Emu(1800000), Emu(11400000), Emu(440000), [
    ("●「学習評価」とは…", {'size': 30, 'bold': True, 'color': NAVY}),
], default_size=30, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_body_box(s, Emu(400000), Emu(2300000), Emu(11400000), Emu(1500000), [
    [("各教科の目標に照らして、生徒の実現状況を　", {'size': 28, 'color': BLACK}),
     ("観点ごとに評価", {'size': 30, 'color': ORANGE_KW})],
    [("し、学習状況を　", {'size': 28, 'color': BLACK}),
     ("分析的に捉える", {'size': 30, 'color': ORANGE_KW}),
     ("もの", {'size': 28, 'color': BLACK})],
], default_size=28, line_spacing=1.4)

add_plain_text(s, Emu(400000), Emu(3950000), Emu(11400000), Emu(440000), [
    ("●「評定」とは…", {'size': 30, 'bold': True, 'color': NAVY}),
], default_size=30, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_body_box(s, Emu(400000), Emu(4450000), Emu(11400000), Emu(1500000), [
    [("観点別評価を基に、", {'size': 28, 'color': BLACK}),
     ("総括的", {'size': 30, 'color': ORANGE_KW}),
     ("に学習状況を示す", {'size': 28, 'color': BLACK})],
    ("中学校は ５段階 で表記する", {'size': 28, 'color': BLACK}),
], default_size=28, line_spacing=1.4)

add_source_line(s, "中学校学習指導要領　総則（平成29年告示）／文部科学省 総則・評価特別部会資料")

# ============================================================
# スライド6: 学習評価を行う目的
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 1　評価をめぐる「いま」")
add_sub_heading(s, "学習評価を行う目的")
add_page_number(s, 6)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(900000), [
    [("▶  教師の　", {'size': 32, 'color': NAVY}),
     ("指導改善", {'size': 36, 'color': ORANGE_KW}),
     ("　につなげる", {'size': 32, 'color': NAVY})],
], default_size=32, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0, fill_color=LIGHT_WATER)

add_body_box(s, Emu(800000), Emu(2850000), Emu(10800000), Emu(800000), [
    ("→ 次の授業では「○○を重点的に」と判断する", {'size': 26, 'color': BLACK}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE, border_color=None, fill_color=WHITE,
   default_bold=False)

add_body_box(s, Emu(400000), Emu(3850000), Emu(11400000), Emu(900000), [
    [("▶  生徒の　", {'size': 32, 'color': NAVY}),
     ("学習改善", {'size': 36, 'color': ORANGE_KW}),
     ("　につなげる", {'size': 32, 'color': NAVY})],
], default_size=32, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0, fill_color=LIGHT_WATER)

add_body_box(s, Emu(800000), Emu(4850000), Emu(10800000), Emu(1100000), [
    ("→ 生徒に「次は○○に取り組んだ方がよい」と", {'size': 26, 'color': BLACK}),
    ("　 具体的に伝える", {'size': 26, 'color': BLACK}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE, border_color=None, fill_color=WHITE,
   default_bold=False, line_spacing=1.3)

# ============================================================
# スライド7: キーメッセージ①
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 1　評価をめぐる「いま」")
add_sub_heading(s, "【キーメッセージ①】")
add_page_number(s, 7)

add_body_box(s, Emu(800000), Emu(2200000), Emu(10600000), Emu(1500000), [
    ("評価＝「ためる」のでなく", {'size': 38, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
    [("「", {'size': 44, 'color': NAVY}),
     ("使う", {'size': 48, 'color': ORANGE_KW}),
     ("」もの", {'size': 44, 'color': NAVY})],
], default_size=42, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=5.0,
   fill_color=LIGHT_WATER, align=PP_ALIGN.CENTER)

add_body_box(s, Emu(800000), Emu(4100000), Emu(10600000), Emu(1700000), [
    ("評価は、", {'size': 28, 'color': BLACK, 'align': PP_ALIGN.CENTER}),
    [("明日の授業を変えるため", {'size': 32, 'color': ORANGE_KW, 'align': PP_ALIGN.CENTER})],
    ("に行う。", {'size': 28, 'color': BLACK, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

add_source_line(s, "中央教育審議会 H31.1 報告 ／ 文科省 R8.3 資料1-1")

# ============================================================
# スライド8: 全国の課題
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 1　評価をめぐる「いま」")
add_sub_heading(s, "学習評価の現状と課題（全国）")
add_page_number(s, 8)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(4350000), [
    [("○ 評価の結果が　", {'size': 26, 'color': BLACK}),
     ("学習改善につながっていない", {'size': 28, 'color': ORANGE_KW})],
    ("", {'size': 10}),
    [("○ 教師によって　", {'size': 26, 'color': BLACK}),
     ("評価の方針が異なる", {'size': 28, 'color': ORANGE_KW})],
    ("", {'size': 10}),
    [("○ 評価のための　", {'size': 26, 'color': BLACK}),
     ("記録に労力が割かれる", {'size': 28, 'color': ORANGE_KW})],
    ("", {'size': 10}),
    [("○ ", {'size': 26, 'color': BLACK}),
     ("「指導に生かす」と「記録に残す」", {'size': 28, 'color': ORANGE_KW})],
    ("　 が区別されずに混在", {'size': 26, 'color': BLACK}),
], default_size=26, line_spacing=1.3, border_width=4.0, fill_color=LIGHT_WATER)

add_source_line(s, "平成31年１月 中央教育審議会「児童生徒の学習評価の在り方について（報告）」")

# ============================================================
# スライド9: 国の最新動向
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 1　評価をめぐる「いま」")
add_sub_heading(s, "国はどう動いているか（令和8年3月）")
add_page_number(s, 9)

arrows = [
    "形式的・過度な「評価材料集め」を抑制する",
    "一人一人の「良さや成長」を肯定的に評価する",
    "「学びに向かう力・人間性等」の見取り方を見直す",
]
y = 1900000
for i, txt in enumerate(arrows):
    yy = y + i * 1100000
    add_body_box(s, Emu(400000), Emu(yy), Emu(11400000), Emu(950000), [
        [("▶  ", {'size': 32, 'color': WATER_BLUE}),
         (txt, {'size': 28, 'color': NAVY})],
    ], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

add_source_line(s, "令和８年３月 文部科学省 教育課程部会 総則・評価特別部会 資料１-１")

# ============================================================
# スライド10: Part2 章扉
# ============================================================
s = add_blank_slide()
add_page_number(s, 10)
add_body_box(s, Emu(1000000), Emu(2100000), Emu(10200000), Emu(2700000), [
    ("Part 2", {'size': 36, 'bold': True, 'color': WATER_BLUE, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 14}),
    ("石神井南中の現状と課題", {'size': 50, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 8}),
    ("─ データと校長のオーダーから読み解く ─",
     {'size': 20, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=5.0, fill_color=LIGHT_WATER)

# ============================================================
# スライド11: 現状①観点別評定割合
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 2　石神井南中の現状")
add_sub_heading(s, "現状①　観点別評定割合〔R7後期〕")
add_page_number(s, 11)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(900000), [
    ("3観点とも　A : B : C = およそ　1.5 : 7.5 : 1.0",
     {'size': 30, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=30, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=4.0)

add_body_box(s, Emu(400000), Emu(2900000), Emu(11400000), Emu(3300000), [
    ("○ Bが大半 → Bの「幅」が広すぎる可能性", {'size': 26, 'color': BLACK}),
    ("", {'size': 8}),
    [("○ ", {'size': 26, 'color': BLACK}),
     ("「Bと判断する姿」が曖昧", {'size': 28, 'color': ORANGE_KW}),
     ("だと判断が割れる", {'size': 26, 'color': BLACK})],
    ("", {'size': 8}),
    [("○ 観点ごとに ", {'size': 26, 'color': BLACK}),
     ("評価規準を言語化", {'size': 28, 'color': ORANGE_KW}),
     (" する必要", {'size': 26, 'color': BLACK})],
    ("", {'size': 8}),
    ("○ 教師により方針が異なる → 校内・教科会で揃える", {'size': 26, 'color': BLACK}),
], default_size=26, line_spacing=1.3, border_width=3.0)

add_source_line(s, "石神井南中学校提供データ R7後期成績分布より")

# ============================================================
# スライド12: 現状②学校評価アンケート
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 2　石神井南中の現状")
add_sub_heading(s, "現状②　学校評価アンケート〔R7〕")
add_page_number(s, 12)

add_body_box(s, Emu(400000), Emu(1850000), Emu(5600000), Emu(640000), [
    ("保護者から", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

add_body_box(s, Emu(400000), Emu(2490000), Emu(5600000), Emu(2400000), [
    [("○ ", {'size': 22, 'color': BLACK}),
     ("補習・個別支援", {'size': 24, 'color': ORANGE_KW})],
    ("　 が不十分との指摘", {'size': 22, 'color': BLACK}),
    ("", {'size': 8}),
    [("○ ", {'size': 22, 'color': BLACK}),
     ("学校HPで情報", {'size': 24, 'color': ORANGE_KW})],
    ("　 が得にくい", {'size': 22, 'color': BLACK}),
], default_size=22, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(6200000), Emu(1850000), Emu(5600000), Emu(640000), [
    ("生徒から", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=WATER_BLUE, border_color=WATER_BLUE, border_width=2.0)

add_body_box(s, Emu(6200000), Emu(2490000), Emu(5600000), Emu(2400000), [
    [("○ ", {'size': 22, 'color': BLACK}),
     ("先生によって対応が違う", {'size': 24, 'color': ORANGE_KW})],
    ("　 と感じる場面がある", {'size': 22, 'color': BLACK}),
    ("", {'size': 8}),
    [("○ ", {'size': 22, 'color': BLACK}),
     ("HP・通信から情報", {'size': 24, 'color': ORANGE_KW})],
    ("　 を得にくい", {'size': 22, 'color': BLACK}),
], default_size=22, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(400000), Emu(5100000), Emu(11400000), Emu(1100000), [
    ("教師ごとの差異 → 評価の「信頼性」課題に直結",
     {'size': 28, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=4.0)

add_source_line(s, "石神井南中学校 R7学校評価アンケート結果")

# ============================================================
# スライド13: 校長オーダーの3観点
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 2　石神井南中の現状")
add_sub_heading(s, "校長から ─ 教員に意識してほしい3点")
add_page_number(s, 13)

items = [
    ("①", "「指導と評価の一体化」", "を意識した授業展開"),
    ("②", "「妥当性・信頼性」", "ある評価"),
    ("③", "「主体的に学習に取り組む態度」", "の評価精度"),
]
y = 1850000
for i, (num, kw, tail) in enumerate(items):
    yy = y + i * 1370000
    add_body_box(s, Emu(400000), Emu(yy), Emu(1100000), Emu(1100000), [
        (num, {'size': 50, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
    ], default_size=50, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
    add_body_box(s, Emu(1600000), Emu(yy), Emu(10200000), Emu(1100000), [
        [(kw, {'size': 26, 'color': ORANGE_KW}),
         (tail, {'size': 26, 'color': NAVY})],
    ], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

add_source_line(s, "校長依頼内容（5/30）─ 本日の講義の出発点")

# ============================================================
# スライド14: Part3 章扉
# ============================================================
s = add_blank_slide()
add_page_number(s, 14)
add_body_box(s, Emu(1000000), Emu(2100000), Emu(10200000), Emu(2700000), [
    ("Part 3", {'size': 36, 'bold': True, 'color': WATER_BLUE, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 14}),
    ("「指導と評価の一体化」の核", {'size': 46, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 8}),
    ("─ 妥当性 × 信頼性 × Bと判断する姿 ─",
     {'size': 20, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=5.0, fill_color=LIGHT_WATER)

# ============================================================
# スライド15: 評価の3つの種類
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "評価の3つの種類")
add_page_number(s, 15)

kinds = [
    ("診断的評価", "学習前", "例：レディネステスト・コース分け"),
    ("形成的評価", "学習過程", "例：振り返り・単元テスト・観察"),
    ("総括的評価", "学習後", "例：観点別評価・評定"),
]
y = 1850000
for i, (name, when, ex) in enumerate(kinds):
    yy = y + i * 1380000
    add_body_box(s, Emu(400000), Emu(yy), Emu(3700000), Emu(1100000), [
        (name, {'size': 30, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
    ], default_size=30, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
    add_body_box(s, Emu(4200000), Emu(yy), Emu(7600000), Emu(1100000), [
        [(when, {'size': 26, 'color': ORANGE_KW}),
         ("　に行う評価", {'size': 24, 'color': NAVY})],
        (ex, {'size': 22, 'color': DARK_GRAY}),
    ], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0, line_spacing=1.3)

add_source_line(s, "国立教育政策研究所「学習評価の在り方ハンドブック（中学校編）」")

# ============================================================
# スライド16: 妥当性
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "「妥当性」のある評価とは")
add_page_number(s, 16)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(640000), [
    ("◆ 妥当性とは…", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.LEFT}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

add_body_box(s, Emu(400000), Emu(2500000), Emu(11400000), Emu(1100000), [
    [("対象である ", {'size': 28, 'color': BLACK}),
     ("資質・能力を適切に反映", {'size': 30, 'color': ORANGE_KW}),
     (" していること", {'size': 28, 'color': BLACK})],
], default_size=28, line_spacing=1.3, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

add_body_box(s, Emu(400000), Emu(3750000), Emu(11400000), Emu(640000), [
    ("◆ 妥当性を確保するには…", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.LEFT}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

add_body_box(s, Emu(400000), Emu(4400000), Emu(11400000), Emu(1850000), [
    [("○ 指導の ", {'size': 24, 'color': BLACK}),
     ("ねらいを明確", {'size': 26, 'color': ORANGE_KW}),
     (" にし、内容を設定", {'size': 24, 'color': BLACK})],
    ("", {'size': 6}),
    [("○ ねらいに ", {'size': 24, 'color': BLACK}),
     ("ふさわしい評価場面・方法", {'size': 26, 'color': ORANGE_KW}),
     (" を選ぶ", {'size': 24, 'color': BLACK})],
    ("", {'size': 6}),
    [("○ ", {'size': 24, 'color': BLACK}),
     ("「測りたい力」と「評価方法」", {'size': 26, 'color': ORANGE_KW}),
     (" を一致させる", {'size': 24, 'color': BLACK})],
], default_size=24, line_spacing=1.3, border_width=3.0)

add_source_line(s, "令和２年９月 東京都教育委員会「指導と評価の一体化を目指して」")

# ============================================================
# スライド17: 信頼性
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "「信頼性」のある評価とは")
add_page_number(s, 17)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(640000), [
    ("◆ 信頼性とは…", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.LEFT}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

add_body_box(s, Emu(400000), Emu(2500000), Emu(11400000), Emu(1100000), [
    [("教師の主観に流れず ", {'size': 28, 'color': BLACK}),
     ("誰が評価しても同じ結果", {'size': 30, 'color': ORANGE_KW}),
     (" になる", {'size': 28, 'color': BLACK})],
], default_size=28, line_spacing=1.3, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

add_body_box(s, Emu(400000), Emu(3750000), Emu(11400000), Emu(640000), [
    ("◆ 信頼性を確保するには…", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.LEFT}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

add_body_box(s, Emu(400000), Emu(4400000), Emu(11400000), Emu(1850000), [
    [("○ 適切な ", {'size': 24, 'color': BLACK}),
     ("評価規準・評価方法", {'size': 26, 'color': ORANGE_KW}),
     (" を整える", {'size': 24, 'color': BLACK})],
    ("", {'size': 6}),
    [("○ ", {'size': 24, 'color': BLACK}),
     ("学校全体で組織的・計画的", {'size': 26, 'color': ORANGE_KW}),
     (" に行う", {'size': 24, 'color': BLACK})],
    ("", {'size': 6}),
    [("○ 教科会で ", {'size': 24, 'color': BLACK}),
     ("Bの姿を揃える", {'size': 26, 'color': ORANGE_KW}),
     ("　→揺れない評価へ", {'size': 24, 'color': BLACK})],
], default_size=24, line_spacing=1.3, border_width=3.0)

add_source_line(s, "令和２年９月 東京都教育委員会「指導と評価の一体化を目指して」")

# ============================================================
# スライド18: 評価計画は単元でデザイン
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "評価計画は「単元」でデザインする")
add_page_number(s, 18)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(800000), [
    ("毎時間 × 3観点 で評価する必要はない",
     {'size': 30, 'bold': True, 'color': ORANGE_KW, 'align': PP_ALIGN.CENTER}),
], default_size=30, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=4.0)

# 2列
add_body_box(s, Emu(400000), Emu(2800000), Emu(5600000), Emu(640000), [
    ("毎時間の評価", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=WATER_BLUE, border_color=WATER_BLUE, border_width=2.0)
add_body_box(s, Emu(400000), Emu(3440000), Emu(5600000), Emu(2400000), [
    [("→ ", {'size': 24, 'color': NAVY}),
     ("学習改善・指導改善", {'size': 26, 'color': ORANGE_KW})],
    ("　 に生かす評価", {'size': 24, 'color': BLACK}),
    ("", {'size': 8}),
    ("（記録は最小限でOK）", {'size': 22, 'color': DARK_GRAY}),
], default_size=24, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(6200000), Emu(2800000), Emu(5600000), Emu(640000), [
    ("総括用の評価", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
add_body_box(s, Emu(6200000), Emu(3440000), Emu(5600000), Emu(2400000), [
    [("→ ", {'size': 24, 'color': NAVY}),
     ("記録に残す評価", {'size': 26, 'color': ORANGE_KW})],
    ("　 を計画的に", {'size': 24, 'color': BLACK}),
    ("", {'size': 8}),
    ("（単元の節目で1観点ずつ）", {'size': 22, 'color': DARK_GRAY}),
], default_size=24, line_spacing=1.3, border_width=3.0)

add_plain_text(s, Emu(400000), Emu(5950000), Emu(11400000), Emu(360000), [
    ("【キーメッセージ②】単元でデザイン＝全部記録しなくていい",
     {'size': 16, 'color': NAVY, 'bold': True, 'align': PP_ALIGN.CENTER}),
], default_size=16, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_source_line(s, "国立教育政策研究所「学習評価の在り方ハンドブック（中学校編）」")

# ============================================================
# スライド19: 3観点の関係
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "3観点の関係を押さえる")
add_page_number(s, 19)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(900000), [
    ("知識・技能　×　思考・判断・表現　×　主体的に学習に取り組む態度",
     {'size': 24, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

add_body_box(s, Emu(400000), Emu(2900000), Emu(11400000), Emu(1100000), [
    [("「態度」だけが突出 → ", {'size': 26, 'color': BLACK}),
     ("評価のバランスを疑う", {'size': 28, 'color': ORANGE_KW})],
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=3.0)

add_body_box(s, Emu(400000), Emu(4150000), Emu(11400000), Emu(2050000), [
    [("○ ", {'size': 24, 'color': BLACK}),
     ("「主体的」", {'size': 26, 'color': ORANGE_KW}),
     ("は「知・技」「思・判・表」を", {'size': 24, 'color': BLACK})],
    ("　 育てる場面で同時に育てる ─ 別物ではない", {'size': 24, 'color': BLACK}),
    ("", {'size': 6}),
    [("○ バランスがおかしい → ", {'size': 24, 'color': BLACK}),
     ("改善を図る", {'size': 26, 'color': ORANGE_KW})],
], default_size=24, line_spacing=1.3, border_width=3.0)

add_source_line(s, "平成31年１月 中央教育審議会「児童生徒の学習評価の在り方について（報告）」")

# ============================================================
# スライド20: 「主体的に取り組む態度」の2側面
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "「主体的に学習に取り組む態度」の2側面")
add_page_number(s, 20)

add_body_box(s, Emu(400000), Emu(1850000), Emu(5600000), Emu(700000), [
    ("① 粘り強い取組", {'size': 28, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
add_body_box(s, Emu(400000), Emu(2550000), Emu(5600000), Emu(1900000), [
    ("自らの学習を", {'size': 24, 'color': BLACK}),
    [("　", {'size': 24}),
     ("粘り強く調整", {'size': 28, 'color': ORANGE_KW})],
    ("しようとしている姿", {'size': 24, 'color': BLACK}),
], default_size=24, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(6200000), Emu(1850000), Emu(5600000), Emu(700000), [
    ("② 自己調整", {'size': 28, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=WATER_BLUE, border_color=WATER_BLUE, border_width=2.0)
add_body_box(s, Emu(6200000), Emu(2550000), Emu(5600000), Emu(1900000), [
    ("自らの学習状況を", {'size': 24, 'color': BLACK}),
    [("　", {'size': 24}),
     ("把握し、調整", {'size': 28, 'color': ORANGE_KW})],
    ("しようとしている姿", {'size': 24, 'color': BLACK}),
], default_size=24, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(400000), Emu(4700000), Emu(11400000), Emu(1500000), [
    ("「挙手の回数」「ノートの取り方」だけで判断しない",
     {'size': 28, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=4.0)

add_source_line(s, "国立教育政策研究所「指導と評価の一体化のための参考資料」")

# ============================================================
# スライド21: 「Bと判断する姿」共通の姿
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "「Bと判断する姿」を具体化する　①")
add_page_number(s, 21)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(700000), [
    ("教科横断で見える「共通の姿」", {'size': 28, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

labels = ["立ち止まる", "振り返る", "やり直す", "試行錯誤する"]
x_start = 400000
w = 2800000
gap = 80000
for i, lbl in enumerate(labels):
    xx = x_start + i * (w + gap)
    add_body_box(s, Emu(xx), Emu(2800000), Emu(w), Emu(1100000), [
        (lbl, {'size': 26, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
    ], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=LIGHT_WATER, border_width=3.0)

add_body_box(s, Emu(400000), Emu(4150000), Emu(11400000), Emu(2050000), [
    [("○ これらの姿が ", {'size': 24, 'color': BLACK}),
     ("単元の途中で", {'size': 26, 'color': ORANGE_KW}),
     (" 見えるか", {'size': 24, 'color': BLACK})],
    ("", {'size': 6}),
    [("○ ", {'size': 24, 'color': BLACK}),
     ("教科ごとに具体例", {'size': 26, 'color': ORANGE_KW}),
     (" を言語化する", {'size': 24, 'color': BLACK})],
    ("", {'size': 6}),
    ("○ 教科会で1単元1事例ずつ揃えるだけでOK", {'size': 24, 'color': BLACK}),
], default_size=24, line_spacing=1.3, border_width=3.0)

add_source_line(s, "国立教育政策研究所「指導と評価の一体化のための参考資料」")

# ============================================================
# スライド22: 数学の具体例
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "「Bと判断する姿」を具体化する　②")
add_page_number(s, 22)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(640000), [
    ("例：数学「二次方程式」── 学習シートに「気を付けるポイント」を記述",
     {'size': 20, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.LEFT}),
], default_size=20, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

# B
add_body_box(s, Emu(400000), Emu(2550000), Emu(5600000), Emu(640000), [
    ("▼ おおむね満足（B）", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=WATER_BLUE, border_color=WATER_BLUE, border_width=2.0)
add_body_box(s, Emu(400000), Emu(3190000), Emu(5600000), Emu(2200000), [
    [("ポイントが ", {'size': 24, 'color': BLACK}),
     ("書かれている", {'size': 26, 'color': ORANGE_KW})],
    ("", {'size': 8}),
    ("例：別の文字に置き換えて", {'size': 20, 'color': DARK_GRAY}),
    ("　 解く方法を使えるようにする", {'size': 20, 'color': DARK_GRAY}),
], default_size=22, line_spacing=1.3, border_width=3.0)

# A
add_body_box(s, Emu(6200000), Emu(2550000), Emu(5600000), Emu(640000), [
    ("▼ 十分満足（A）", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
add_body_box(s, Emu(6200000), Emu(3190000), Emu(5600000), Emu(2200000), [
    [("ポイント＋", {'size': 24, 'color': BLACK}),
     ("理由", {'size': 26, 'color': ORANGE_KW}),
     ("が書かれている", {'size': 24, 'color': BLACK})],
    ("", {'size': 8}),
    ("例：式を見ずに展開すると時間も", {'size': 20, 'color': DARK_GRAY}),
    ("　 間違いも増えるから、", {'size': 20, 'color': DARK_GRAY}),
    ("　 置き換えを使えるようにする", {'size': 20, 'color': DARK_GRAY}),
], default_size=22, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(400000), Emu(5500000), Emu(11400000), Emu(800000), [
    ("同じ気付きでも「理由まで書けるか」でA/Bが分かれる",
     {'size': 24, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=4.0)

add_source_line(s, "国立教育政策研究所「指導と評価の一体化のための参考資料」（数学）")

# ============================================================
# スライド23: 最重要メッセージ
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "《本日の最重要メッセージ》")
add_page_number(s, 23)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(800000), [
    ("「主体的に学習に取り組む態度」は",
     {'size': 28, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)

add_body_box(s, Emu(400000), Emu(2700000), Emu(11400000), Emu(900000), [
    ("“指導してから評価する”",
     {'size': 38, 'bold': True, 'color': ORANGE_KW, 'align': PP_ALIGN.CENTER}),
], default_size=38, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=5.0)

# ×
add_body_box(s, Emu(400000), Emu(3850000), Emu(900000), Emu(900000), [
    ("×", {'size': 44, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=44, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=RED, border_color=RED, border_width=2.0)
add_body_box(s, Emu(1400000), Emu(3850000), Emu(10400000), Emu(900000), [
    [("「個人任せ」で見取ろうとする → ", {'size': 22, 'color': NAVY}),
     ("育たない・評価できない", {'size': 24, 'color': RED})],
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE, border_color=RED, border_width=3.0)

# ○
add_body_box(s, Emu(400000), Emu(4850000), Emu(900000), Emu(900000), [
    ("○", {'size': 44, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=44, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=GREEN, border_color=GREEN, border_width=2.0)
add_body_box(s, Emu(1400000), Emu(4850000), Emu(10400000), Emu(900000), [
    [("「学び方」を指導してから見取る → ", {'size': 22, 'color': NAVY}),
     ("育つ・評価できる", {'size': 24, 'color': GREEN})],
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE, border_color=GREEN, border_width=3.0)

add_plain_text(s, Emu(400000), Emu(5900000), Emu(11400000), Emu(380000), [
    ("見取る前に「学び方」を授業の中で教える",
     {'size': 18, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=18, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_source_line(s, "国立教育政策研究所「指導と評価の一体化のための参考資料」")

# ============================================================
# スライド24: 「学び方」9つの働きかけ
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 3　一体化の核")
add_sub_heading(s, "「学び方」を指導する9つの働きかけ")
add_page_number(s, 24)

works = [
    "1. 学習目標の共有",
    "2. 既習との関連付け",
    "3. 見通しを持たせる",
    "4. 振り返りの場面設定",
    "5. 自己評価の機会",
    "6. ペア・グループ対話",
    "7. 学習の選択肢提示",
    "8. ICTでの可視化",
    "9. 教師のモデリング",
]
y_start = 1850000
x_start = 400000
w = 3750000
h = 850000
gx = 80000
gy = 130000
for i, w_text in enumerate(works):
    row, col = divmod(i, 3)
    xx = x_start + col * (w + gx)
    yy = y_start + row * (h + gy)
    add_body_box(s, Emu(xx), Emu(yy), Emu(w), Emu(h), [
        (w_text, {'size': 22, 'color': NAVY, 'align': PP_ALIGN.LEFT}),
    ], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=LIGHT_WATER, border_width=3.0)

add_body_box(s, Emu(400000), Emu(5650000), Emu(11400000), Emu(640000), [
    ("1つ選んで、次の単元から実践する",
     {'size': 24, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=YELLOW, border_color=ORANGE_KW, border_width=3.0)

add_source_line(s, "国立教育政策研究所（2020）参考資料より整理")

# ============================================================
# スライド25: Part4 章扉
# ============================================================
s = add_blank_slide()
add_page_number(s, 25)
add_body_box(s, Emu(1000000), Emu(2100000), Emu(10200000), Emu(2700000), [
    ("Part 4", {'size': 36, 'bold': True, 'color': WATER_BLUE, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 14}),
    ("明日から動かす", {'size': 50, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
    ("", {'size': 8}),
    ("─ 単元計画／教科会／組織で支える ─",
     {'size': 20, 'bold': True, 'color': DEEP_WATER, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=5.0, fill_color=LIGHT_WATER)

# ============================================================
# スライド26: 単元計画に「記録」欄
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 4　明日から動かす")
add_sub_heading(s, "単元計画に「記録」欄を1列足す")
add_page_number(s, 26)

add_body_box(s, Emu(400000), Emu(1850000), Emu(11400000), Emu(900000), [
    ("「いつ／何で／どう見取るか」を可視化",
     {'size': 28, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=28, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=4.0)

# 表風
headers = ["時", "学習内容", "観点", "評価方法・場面"]
header_w = [1100000, 4500000, 2000000, 3800000]
x = 400000
y = 3000000
for i, hh in enumerate(headers):
    add_body_box(s, Emu(x), Emu(y), Emu(header_w[i]), Emu(600000), [
        (hh, {'size': 22, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
    ], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=1.5)
    x += header_w[i]

rows = [
    ("3-4", "気付きの記述", "態度", "学習シート"),
    ("7", "ペーパーテスト", "知技", "小テスト"),
]
y += 650000
for r in rows:
    x = 400000
    for i, txt in enumerate(r):
        add_body_box(s, Emu(x), Emu(y), Emu(header_w[i]), Emu(700000), [
            (txt, {'size': 22, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
        ], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=1.5)
        x += header_w[i]
    y += 730000

add_body_box(s, Emu(400000), Emu(5650000), Emu(11400000), Emu(640000), [
    ("全部記録しなくていい ─ 単元の節目で1観点ずつ",
     {'size': 22, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=YELLOW, border_color=ORANGE_KW, border_width=3.0)

add_source_line(s, "国立教育政策研究所「指導と評価の一体化のための参考資料」より作成")

# ============================================================
# スライド27: 教科会で揃える3つ
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 4　明日から動かす")
add_sub_heading(s, "教科会で揃える3つ")
add_page_number(s, 27)

items = [
    ("①", "Bの姿", "観点別評価で「B」と判断する具体的な姿"),
    ("②", "評価場面", "いつ・どの活動で・どんな方法で見取るか"),
    ("③", "総括方法", "観点別評価から評定への総括ルール"),
]
y = 1850000
for i, (num, kw, desc) in enumerate(items):
    yy = y + i * 1380000
    add_body_box(s, Emu(400000), Emu(yy), Emu(1100000), Emu(1100000), [
        (num, {'size': 50, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
    ], default_size=50, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
    add_body_box(s, Emu(1600000), Emu(yy), Emu(10200000), Emu(1100000), [
        (kw, {'size': 28, 'color': ORANGE_KW}),
        (desc, {'size': 22, 'color': NAVY}),
    ], default_size=24, line_spacing=1.2, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

add_source_line(s, "東京都教育委員会「指導と評価の一体化を目指して」")

# ============================================================
# スライド28: 組織で支える信頼性
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 4　明日から動かす")
add_sub_heading(s, "妥当性は「個人」で、信頼性は「組織」で")
add_page_number(s, 28)

add_body_box(s, Emu(400000), Emu(1850000), Emu(5600000), Emu(640000), [
    ("教科会で揃える", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
add_body_box(s, Emu(400000), Emu(2490000), Emu(5600000), Emu(2400000), [
    ("○ 評価規準（Bの姿）", {'size': 22, 'color': NAVY}),
    ("", {'size': 6}),
    ("○ 評価場面・方法", {'size': 22, 'color': NAVY}),
    ("", {'size': 6}),
    ("○ 評定の総括方法", {'size': 22, 'color': NAVY}),
    ("", {'size': 6}),
    ("○ 保護者説明の仕方", {'size': 22, 'color': NAVY}),
], default_size=22, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(6200000), Emu(1850000), Emu(5600000), Emu(640000), [
    ("管理職・主任の役割", {'size': 26, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE,
   fill_color=WATER_BLUE, border_color=WATER_BLUE, border_width=2.0)
add_body_box(s, Emu(6200000), Emu(2490000), Emu(5600000), Emu(2400000), [
    ("○ 教科会の機能化", {'size': 22, 'color': NAVY}),
    ("", {'size': 6}),
    ("○ 評定のチェック体制", {'size': 22, 'color': NAVY}),
    ("", {'size': 6}),
    ("○ 保護者対応の組織化", {'size': 22, 'color': NAVY}),
    ("", {'size': 6}),
    ("○ 「特異なパターン」確認", {'size': 22, 'color': NAVY}),
], default_size=22, line_spacing=1.3, border_width=3.0)

add_body_box(s, Emu(400000), Emu(5100000), Emu(11400000), Emu(700000), [
    ("東京都の特異パターン：AAA→「5/4」でない／ABCの偏り 等",
     {'size': 20, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=20, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=YELLOW, border_color=ORANGE_KW, border_width=2.0)

add_plain_text(s, Emu(400000), Emu(5860000), Emu(11400000), Emu(380000), [
    ("【キーメッセージ③】信頼性は「組織」で担保する",
     {'size': 16, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=16, vertical_anchor=MSO_ANCHOR.MIDDLE)

add_source_line(s, "東京都教育委員会「指導と評価の一体化を目指して」")

# ============================================================
# スライド29: 明日からできること
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 4　明日から動かす")
add_sub_heading(s, "明日からできること ─ 1つ選んでください")
add_page_number(s, 29)

choices = [
    "□ 次の単元の指導計画に「記録」欄を1列足す",
    "□ 教科会で「Bの姿」を1つ言語化して揃える",
    "□ 9つの働きかけから1つ選んで授業に入れる",
    "□ 振り返りシートを1回入れてみる",
]
y = 1850000
for i, txt in enumerate(choices):
    yy = y + i * 970000
    add_body_box(s, Emu(400000), Emu(yy), Emu(11400000), Emu(820000), [
        (txt, {'size': 26, 'color': NAVY}),
    ], default_size=26, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=LIGHT_WATER, border_width=3.0)

add_body_box(s, Emu(400000), Emu(5780000), Emu(11400000), Emu(500000), [
    ("やることを「絞る」から、続く",
     {'size': 22, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.CENTER}),
], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=YELLOW, border_color=ORANGE_KW, border_width=3.0)

# ============================================================
# スライド30: まとめ／協議会へ
# ============================================================
s = add_blank_slide()
add_chapter_bar(s, "Part 5　本日のまとめ")
add_sub_heading(s, "3つのキーメッセージ")
add_page_number(s, 30)

msgs = [
    ("①", "評価は「ためる」より「使う」もの"),
    ("②", "生徒の姿は「単元の中」「学び方を指導した上で」見取る"),
    ("③", "評価の信頼性は「組織」で担保する"),
]
y = 1850000
for i, (num, msg) in enumerate(msgs):
    yy = y + i * 900000
    add_body_box(s, Emu(400000), Emu(yy), Emu(1000000), Emu(770000), [
        (num, {'size': 36, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}),
    ], default_size=36, vertical_anchor=MSO_ANCHOR.MIDDLE,
       fill_color=DEEP_WATER, border_color=DEEP_WATER, border_width=2.0)
    add_body_box(s, Emu(1500000), Emu(yy), Emu(10300000), Emu(770000), [
        (msg, {'size': 22, 'color': NAVY}),
    ], default_size=22, vertical_anchor=MSO_ANCHOR.MIDDLE, border_width=3.0)

add_body_box(s, Emu(400000), Emu(4700000), Emu(11400000), Emu(640000), [
    ("▶  このあとの協議会で…",
     {'size': 24, 'bold': True, 'color': NAVY, 'align': PP_ALIGN.LEFT}),
], default_size=24, vertical_anchor=MSO_ANCHOR.MIDDLE, fill_color=YELLOW, border_color=ORANGE_KW, border_width=3.0)

add_body_box(s, Emu(400000), Emu(5400000), Emu(11400000), Emu(900000), [
    [("教科会単位で ", {'size': 22, 'color': BLACK}),
     ("「Bの姿」を1つ", {'size': 24, 'color': ORANGE_KW}),
     (" 言語化してみてください", {'size': 22, 'color': BLACK})],
    ("　 単元・観点は自由 ／ 30〜45字以内で記述", {'size': 18, 'color': DARK_GRAY}),
], default_size=22, line_spacing=1.3, border_width=3.0)

# ========== 保存 ==========
import os
os.makedirs("/tmp/output", exist_ok=True)
prs.save("/tmp/output/講義スライド_石神井南中_20260603.pptx")
print("OK: PPTX created")
print(f"Slides: {len(prs.slides)}")
