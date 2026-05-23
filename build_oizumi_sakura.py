"""
小中一貫教育（大泉桜学園）プレゼン作成スクリプト
教育委員会職員視点で、新区長に対し新規小中一貫教育校開設の効果がないことを示す資料。
"""

import copy
from pptx import Presentation
from pptx.util import Emu, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ===== Color palette (sample PPTXより抽出) =====
COLOR_HEADER_BLUE = RGBColor(0x0F, 0x9E, 0xD5)  # ヘッダー青
COLOR_ACCENT_YELLOW = RGBColor(0xFF, 0xC0, 0x00)  # 強調黄
COLOR_LIGHT_YELLOW = RGBColor(0xFF, 0xFB, 0xC1)  # 薄黄
COLOR_LIGHT_BLUE = RGBColor(0xE8, 0xF7, 0xF9)  # 薄水色
COLOR_DARK_GRAY = RGBColor(0x40, 0x40, 0x40)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_RED = RGBColor(0xC0, 0x00, 0x00)
COLOR_DARK_BLUE = RGBColor(0x1F, 0x4E, 0x79)
COLOR_GREEN = RGBColor(0x00, 0x6F, 0x3C)
COLOR_LIGHT_PINK = RGBColor(0xFC, 0xE4, 0xE4)
COLOR_LIGHT_GREEN = RGBColor(0xE2, 0xF0, 0xD9)
COLOR_GRAY_BORDER = RGBColor(0x80, 0x80, 0x80)

SLIDE_W = 9144000
SLIDE_H = 5143500

FONT_JA = "Meiryo UI"


def set_font(run, size_pt=None, bold=False, color=None, font_name=FONT_JA):
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = font_name
    if color is not None:
        run.font.color.rgb = color
    # set East Asian font
    rPr = run._r.get_or_add_rPr()
    # remove old ea, latin
    for tag in ["{http://schemas.openxmlformats.org/drawingml/2006/main}ea",
                "{http://schemas.openxmlformats.org/drawingml/2006/main}latin"]:
        for el in rPr.findall(tag):
            rPr.remove(el)
    ea = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
    ea.set("typeface", font_name)
    latin = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
    latin.set("typeface", font_name)


def add_textbox(slide, x, y, w, h, text, size_pt=11, bold=False, color=COLOR_DARK_GRAY,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None, line=None, font_name=FONT_JA):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Emu(36000)
    tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000)
    tf.margin_bottom = Emu(18000)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, text_line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = text_line
        set_font(r, size_pt=size_pt, bold=bold, color=color, font_name=font_name)
    if fill is not None:
        tb.fill.solid()
        tb.fill.fore_color.rgb = fill
    if line is not None:
        tb.line.color.rgb = line
        tb.line.width = Pt(0.75)
    return tb


def add_shape(slide, shape_type, x, y, w, h, text=None, size_pt=11, bold=False,
              text_color=COLOR_DARK_GRAY, fill=None, line=None, align=PP_ALIGN.CENTER,
              anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_JA, line_width=0.75):
    sh = slide.shapes.add_shape(shape_type, x, y, w, h)
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_width)
    sh.shadow.inherit = False
    if text is not None:
        tf = sh.text_frame
        tf.margin_left = Emu(36000)
        tf.margin_right = Emu(36000)
        tf.margin_top = Emu(18000)
        tf.margin_bottom = Emu(18000)
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        lines = text.split("\n")
        for i, sub_line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = sub_line
            set_font(r, size_pt=size_pt, bold=bold, color=text_color, font_name=font_name)
    return sh


def add_header(slide, title_text, page_num):
    """ヘッダー（青いhomePlate型）を追加"""
    bar = slide.shapes.add_shape(MSO_SHAPE.PENTAGON, 0, Emu(42000), SLIDE_W, Emu(461000))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_HEADER_BLUE
    bar.line.fill.background()
    bar.shadow.inherit = False
    tf = bar.text_frame
    tf.margin_left = Emu(180000)
    tf.margin_top = Emu(20000)
    tf.margin_bottom = Emu(20000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title_text
    set_font(r, size_pt=22, bold=True, color=COLOR_WHITE)

    # 部署タグ（右下）
    dept = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Emu(7250000), Emu(4830000),
                                  Emu(1380000), Emu(260000))
    dept.fill.solid()
    dept.fill.fore_color.rgb = COLOR_HEADER_BLUE
    dept.line.fill.background()
    dept.shadow.inherit = False
    tf = dept.text_frame
    tf.margin_left = Emu(36000)
    tf.margin_right = Emu(36000)
    tf.margin_top = Emu(10000)
    tf.margin_bottom = Emu(10000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "教育振興部教育指導課"
    set_font(r, size_pt=10, bold=True, color=COLOR_WHITE)

    # ページ番号
    num = slide.shapes.add_textbox(Emu(8740000), Emu(4830000), Emu(280000), Emu(260000))
    tf = num.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = str(page_num)
    set_font(r, size_pt=12, bold=False, color=COLOR_DARK_GRAY)


def add_section_label(slide, x, y, w, h, text, fill=COLOR_ACCENT_YELLOW, text_color=COLOR_DARK_GRAY, size_pt=12):
    """セクションラベル（小見出し）"""
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.margin_left = Emu(60000)
    tf.margin_right = Emu(60000)
    tf.margin_top = Emu(15000)
    tf.margin_bottom = Emu(15000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    set_font(r, size_pt=size_pt, bold=True, color=text_color)
    return sh


def add_table(slide, x, y, w, h, data, col_widths=None, header_fill=COLOR_HEADER_BLUE,
              header_text_color=COLOR_WHITE, body_size=9, header_size=10, highlight_cells=None,
              highlight_color=COLOR_LIGHT_YELLOW, highlight_text_color=COLOR_RED, body_text_color=COLOR_DARK_GRAY,
              first_col_fill=None, first_col_bold=True):
    """テーブル追加。data はリストのリスト [[c1, c2, ...], ...]。
    highlight_cells: [(row_idx, col_idx), ...] で強調セル指定。"""
    rows = len(data)
    cols = len(data[0])
    tbl_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    tbl = tbl_shape.table
    if col_widths is not None:
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = cw
    if highlight_cells is None:
        highlight_cells = []
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.margin_left = Emu(36000)
            cell.margin_right = Emu(36000)
            cell.margin_top = Emu(18000)
            cell.margin_bottom = Emu(18000)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = str(val)
            is_header = ri == 0
            is_first_col = ci == 0
            is_hl = (ri, ci) in highlight_cells
            if is_hl:
                cell.fill.solid()
                cell.fill.fore_color.rgb = highlight_color
                set_font(r, size_pt=body_size, bold=True, color=highlight_text_color)
            elif is_header:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_fill
                set_font(r, size_pt=header_size, bold=True, color=header_text_color)
            elif is_first_col and first_col_fill is not None:
                cell.fill.solid()
                cell.fill.fore_color.rgb = first_col_fill
                set_font(r, size_pt=body_size, bold=first_col_bold, color=body_text_color)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLOR_WHITE
                set_font(r, size_pt=body_size, bold=False, color=body_text_color)
    return tbl_shape


def add_arrow_down(slide, x, y, w, h, fill=COLOR_ACCENT_YELLOW):
    ar = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, x, y, w, h)
    ar.fill.solid()
    ar.fill.fore_color.rgb = fill
    ar.line.fill.background()
    ar.shadow.inherit = False
    return ar


# ===== Open and modify the presentation =====
src = "/home/user/con30/output/working.pptx"
prs = Presentation(src)

# Delete slides 2-7, keep slide 1
xml_slides = prs.slides._sldIdLst
slides_to_remove = list(xml_slides)[1:]  # all except slide 1
for sld_elem in slides_to_remove:
    rId = sld_elem.get(qn("r:id"))
    prs.part.drop_rel(rId)
    xml_slides.remove(sld_elem)

# Get the blank layout (白紙 = layout index 6)
blank_layout = prs.slide_layouts[6]

# =====================================================================
# SLIDE 2: ①+② 桜学園・小中一貫教育で目指すもの＋成果と課題（声）
# =====================================================================
slide2 = prs.slides.add_slide(blank_layout)
add_header(slide2, "１　桜学園・小中一貫教育で目指すもの／成果と課題（声）", 1)

# 左：目指すもの（基本方針H20.11より5本柱）
add_section_label(slide2, Emu(180000), Emu(580000), Emu(4200000), Emu(320000),
                  "■ 小中一貫教育校設置の効果（基本方針 H20.11）", COLOR_ACCENT_YELLOW, size_pt=11)

aim_x = Emu(180000)
aim_y = Emu(950000)
aim_w = Emu(4200000)
aim_text = (
    "① 一貫した教育課程（9年間を見通した学習・生活指導の充実）\n"
    "② 滑らかな接続（中1ギャップの解消、不登校等の減少）\n"
    "③ 異年齢集団活動（豊かな人間性・社会性の育成）\n"
    "④ 教員の相互協力（学力・体力の向上等の高い教育効果）\n"
    "⑤ 地域社会との連携（学校と地域社会の活性化）"
)
add_textbox(slide2, aim_x, aim_y, aim_w, Emu(1500000), aim_text,
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_BLUE, line=COLOR_HEADER_BLUE)

# 左下：声のテーブル（学校評価R6＆検証報告書H27.10）
add_section_label(slide2, Emu(180000), Emu(2550000), Emu(4200000), Emu(320000),
                  "■ 子供・保護者・教職員の声（肯定的回答率）", COLOR_LIGHT_GREEN,
                  text_color=COLOR_DARK_GRAY, size_pt=11)

voice_data = [
    ["項目", "対象", "肯定率"],
    ["学校に楽しく通っている", "児童生徒", "93%"],
    ["授業改善につながった", "教員", "80%"],
    ["9年間継続した指導・見守り", "保護者", "84%"],
    ["異年齢交流が人間性育成に有効", "保護者", "91%"],
    ["7年進級時の選択肢があるとよい", "保護者", "88%"],
]
add_table(slide2, Emu(180000), Emu(2900000), Emu(4200000), Emu(1700000),
          voice_data, body_size=9, header_size=10,
          col_widths=[Emu(2100000), Emu(1100000), Emu(1000000)])

# 右：成果と課題（2列）
# 上：成果
add_section_label(slide2, Emu(4550000), Emu(580000), Emu(4400000), Emu(320000),
                  "■ 検証で確認された主な成果（H27検証報告書）", COLOR_LIGHT_GREEN,
                  text_color=COLOR_DARK_GRAY, size_pt=11)
success_text = (
    "○ 9年間の系統性を意識した指導の改善\n"
    "○ 全教職員で全児童生徒を見守る体制\n"
    "○ 4-3-2区切りで子供たちが成長（4年生がリーダー経験）\n"
    "○ 一部教科担任制が成長に合致（保護者89% 関係者90%）\n"
    "○ PTA・町会と窓口一本化で地域連携が進展"
)
add_textbox(slide2, Emu(4550000), Emu(950000), Emu(4400000), Emu(1500000), success_text,
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_BLUE, line=COLOR_HEADER_BLUE)

# 下：課題
add_section_label(slide2, Emu(4550000), Emu(2550000), Emu(4400000), Emu(320000),
                  "■ 検証で確認された課題（H27検証報告書）", COLOR_LIGHT_PINK,
                  text_color=COLOR_DARK_GRAY, size_pt=11)
issue_text = (
    "△ 7年生から外部入学する生徒の保護者に不安\n"
    "△ 体育施設・特別教室の小中共用に課題（体格差・教材の違い）\n"
    "△ 5・6年生が部活動に入る活動は限定的（活発化57%）\n"
    "△ 活用力を問う学力では課題（基礎学力は概ね定着）\n"
    "△ 異学年交流の効果は質的には認められるが、定量的な学力・\n"
    "　体力向上効果のエビデンスは限定的"
)
add_textbox(slide2, Emu(4550000), Emu(2900000), Emu(4400000), Emu(1700000), issue_text,
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW)

# 下部メッセージ
add_shape(slide2, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(4680000),
          Emu(8780000), Emu(360000),
          text="検証報告書では「効果」と「課題」が併記されており、定量的な教育効果のエビデンスは限定的。次スライド以降で大泉桜学園の現状データを確認する。",
          size_pt=11, bold=True, text_color=COLOR_WHITE, fill=COLOR_HEADER_BLUE)

# =====================================================================
# SLIDE 3: ③ 児童生徒数の推移（小規模校が解決していないことを示す）
# =====================================================================
slide3 = prs.slides.add_slide(blank_layout)
add_header(slide3, "２　児童生徒数の推移（小規模校化は解決していない）", 2)

# 左：桜中進学率（学園内）
add_section_label(slide3, Emu(180000), Emu(580000), Emu(4400000), Emu(320000),
                  "■ 大泉桜学園 小→中進学状況", COLOR_ACCENT_YELLOW, size_pt=11)
sakura_data = [
    ["", "R3入学\n(現高2)", "R4入学\n(現高1)", "R5入学\n(現中3)", "R6入学\n(現中2)", "R7入学\n(現中1)"],
    ["区立(桜中)", "56", "73", "50", "39", "41"],
    ["区立(桜中以外)", "6", "3", "4", "7", "12"],
    ["区立計", "62", "76", "54", "46", "53"],
    ["区立外", "8", "13", "10", "10", "8"],
    ["合計", "70", "89", "64", "56", "61"],
    ["桜中進学率", "80%", "82%", "78%", "70%", "67%"],
]
add_table(slide3, Emu(180000), Emu(950000), Emu(4400000), Emu(2400000),
          sakura_data, body_size=8, header_size=8,
          col_widths=[Emu(1300000), Emu(620000), Emu(620000), Emu(620000), Emu(620000), Emu(620000)],
          highlight_cells=[(6, 4), (6, 5)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 右：区内全体進学率
add_section_label(slide3, Emu(4750000), Emu(580000), Emu(4200000), Emu(320000),
                  "■ 区内全体（指定校）進学状況", COLOR_LIGHT_BLUE,
                  text_color=COLOR_DARK_GRAY, size_pt=11)
zen_data = [
    ["（参考）区内全体", "R3", "R4", "R5", "R6", "R7"],
    ["区立(指定校)", "3,855", "3,636", "3,723", "3,565", "3,562"],
    ["区立(指定校以外)", "645", "664", "704", "587", "705"],
    ["区立計", "4,500", "4,300", "4,427", "4,152", "4,267"],
    ["区立外", "1,374", "1,395", "1,500", "1,418", "1,455"],
    ["合計", "5,874", "5,695", "5,927", "5,570", "5,722"],
    ["指定校進学率", "66%", "64%", "63%", "64%", "62%"],
]
add_table(slide3, Emu(4750000), Emu(950000), Emu(4200000), Emu(2400000),
          zen_data, body_size=8, header_size=8,
          col_widths=[Emu(1400000), Emu(560000), Emu(560000), Emu(560000), Emu(560000), Emu(560000)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 下部：解説3つ並列
# 左下
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(3500000),
          Emu(2870000), Emu(700000),
          text="① 桜中進学率は下降\n80%(R3)→67%(R7)　▲13ポイント",
          size_pt=11, bold=True, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW, anchor=MSO_ANCHOR.MIDDLE)

# 中下
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(3120000), Emu(3500000),
          Emu(2870000), Emu(700000),
          text="② 区全体の指定校進学率は安定\n（66→62%、2-4ポイント減）",
          size_pt=11, bold=True, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW, anchor=MSO_ANCHOR.MIDDLE)

# 右下
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(6060000), Emu(3500000),
          Emu(2900000), Emu(700000),
          text="③ 人口推計（区将来人口）\n今後20年大きな増減なし",
          size_pt=11, bold=True, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW, anchor=MSO_ANCHOR.MIDDLE)

add_arrow_down(slide3, Emu(4322000), Emu(4280000), Emu(500000), Emu(280000))

# 結論ボックス
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(4620000),
          Emu(8780000), Emu(420000),
          text="小中一貫教育校化しても、桜中進学率はむしろ低下傾向。区全体より進学率の減少幅が大きい。\n小規模校化（中学校生徒数の減少）の課題は、小中一貫教育では解決できていない。",
          size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# SLIDE 4: ④ 桜学園の現状（学力・体力・不登校）
# =====================================================================
slide4 = prs.slides.add_slide(blank_layout)
add_header(slide4, "３　大泉桜学園の現状（学力・体力・不登校）", 3)

# 上左：学力（国語）
add_section_label(slide4, Emu(180000), Emu(580000), Emu(4200000), Emu(280000),
                  "■ 学力調査（国語：平均正答率%）", COLOR_ACCENT_YELLOW, size_pt=10)
kokugo_data = [
    ["国語", "小R4", "小R5", "小R6", "中R4", "中R5", "中R6"],
    ["大泉桜学園", "60", "69", "62", "72", "78", "60"],
    ["練馬区", "69", "70", "71", "72", "74", "62"],
    ["東京都", "69", "69", "70", "70", "72", "61"],
    ["全国", "65.6", "67.2", "67.7", "69", "69.8", "58.1"],
]
add_table(slide4, Emu(180000), Emu(880000), Emu(4200000), Emu(1100000),
          kokugo_data, body_size=8, header_size=8,
          col_widths=[Emu(1080000), Emu(520000), Emu(520000), Emu(520000), Emu(520000), Emu(520000), Emu(520000)],
          highlight_cells=[(1, 1), (1, 3), (1, 6)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 上中：算数・数学
add_section_label(slide4, Emu(4500000), Emu(580000), Emu(4450000), Emu(280000),
                  "■ 学力調査（算数・数学：平均正答率%）", COLOR_ACCENT_YELLOW, size_pt=10)
sansu_data = [
    ["算数/数学", "小R4", "小R5", "小R6", "中R4", "中R5", "中R6"],
    ["大泉桜学園", "62", "69", "56", "58", "62", "55"],
    ["練馬区", "68", "69", "70", "57", "57", "60"],
    ["東京都", "67", "67", "68", "54", "54", "57"],
    ["全国", "63.2", "62.5", "63.4", "51.4", "51", "52.5"],
]
add_table(slide4, Emu(4500000), Emu(880000), Emu(4450000), Emu(1100000),
          sansu_data, body_size=8, header_size=8,
          col_widths=[Emu(1130000), Emu(553000), Emu(553000), Emu(553000), Emu(553000), Emu(553000), Emu(553000)],
          highlight_cells=[(1, 1), (1, 3), (1, 6)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 中段左：区内順位
add_section_label(slide4, Emu(180000), Emu(2050000), Emu(4200000), Emu(280000),
                  "■ 区内順位（学力調査合計点）", COLOR_LIGHT_PINK,
                  text_color=COLOR_DARK_GRAY, size_pt=10)
rank_data = [
    ["", "H29", "H30", "R1", "R3", "R4", "R5", "R6"],
    ["小学校", "59", "56", "62", "53", "62", "31", "63"],
    ["中学校", "22", "31", "33", "7", "5", "3", "25"],
]
add_table(slide4, Emu(180000), Emu(2350000), Emu(4200000), Emu(700000),
          rank_data, body_size=8, header_size=9,
          col_widths=[Emu(720000), Emu(497000), Emu(497000), Emu(497000), Emu(497000), Emu(497000), Emu(497000), Emu(498000)],
          highlight_cells=[(1, 7), (2, 7)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 中段右：体力（R6）
add_section_label(slide4, Emu(4500000), Emu(2050000), Emu(4450000), Emu(280000),
                  "■ 体力調査（R6 参考）", COLOR_LIGHT_PINK,
                  text_color=COLOR_DARK_GRAY, size_pt=10)
tairyoku_text = (
    "小学部：男女ともに区平均と概ね同等または下回る種目が多い\n"
    "中学部：男女とも握力・持久走で都平均を下回る傾向\n"
    "→ 小中一貫教育の「体力向上」効果は明確に確認できず"
)
add_textbox(slide4, Emu(4500000), Emu(2350000), Emu(4450000), Emu(700000),
            tairyoku_text, size_pt=10, color=COLOR_DARK_GRAY,
            fill=COLOR_WHITE, line=COLOR_LIGHT_PINK)

# 下段左：不登校児童生徒数
add_section_label(slide4, Emu(180000), Emu(3150000), Emu(4200000), Emu(280000),
                  "■ 不登校児童生徒数（人）", COLOR_RED,
                  text_color=COLOR_WHITE, size_pt=10)
futoukou_data = [
    ["", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5"],
    ["小学部計", "5", "1", "3", "1", "2", "5", "10", "7"],
    ["中学部計", "8", "12", "14", "8", "13", "21", "23", "21"],
    ["全体計", "13", "13", "17", "9", "15", "26", "33", "28"],
]
add_table(slide4, Emu(180000), Emu(3450000), Emu(4200000), Emu(950000),
          futoukou_data, body_size=8, header_size=8,
          col_widths=[Emu(840000), Emu(420000), Emu(420000), Emu(420000), Emu(420000), Emu(420000), Emu(420000), Emu(420000), Emu(420000)],
          highlight_cells=[(3, 6), (3, 7), (3, 8)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 下段右：不登校出現率
add_section_label(slide4, Emu(4500000), Emu(3150000), Emu(4450000), Emu(280000),
                  "■ 不登校出現率（%）", COLOR_RED,
                  text_color=COLOR_WHITE, size_pt=10)
shutsugen_data = [
    ["", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5"],
    ["桜・小", "1.32", "0.22", "0.68", "0.24", "0.99", "1.28", "2.77", "2.02"],
    ["区(小)", "0.68", "0.61", "0.82", "1.00", "1.12", "1.31", "1.67", "2.14"],
    ["桜・中", "3.38", "5.13", "5.93", "3.52", "5.63", "10.50", "11.11", "10.40"],
    ["区(中)", "3.42", "3.20", "3.26", "4.35", "4.80", "5.23", "6.13", "6.90"],
    ["都(中)", "3.6", "3.78", "4.33", "4.76", "4.93", "5.76", "6.85", "7.80"],
]
add_table(slide4, Emu(4500000), Emu(3450000), Emu(4450000), Emu(1100000),
          shutsugen_data, body_size=7, header_size=8,
          col_widths=[Emu(610000), Emu(480000), Emu(480000), Emu(480000), Emu(480000), Emu(480000), Emu(480000), Emu(480000), Emu(480000)],
          highlight_cells=[(3, 6), (3, 7), (3, 8)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 一番下：要因と結論
add_shape(slide4, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(4620000),
          Emu(8780000), Emu(420000),
          text="主な要因：学力不足、コミュニケーションスキル不足による集団不適応／本人の特性（発達障害・情緒的課題等）の可能性\n"
               "→ 中学部不登校出現率10.40%（R5）は、東京都平均7.80%・練馬区平均6.90%を大きく上回る",
          size_pt=10, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# SLIDE 5: ⑤ 結論 - 小中一貫教育校は、効果がない
# =====================================================================
slide5 = prs.slides.add_slide(blank_layout)
add_header(slide5, "４　結論：新規の小中一貫教育校開設は効果がない", 4)

# 上部：4つの観点で効果なしを示す
add_section_label(slide5, Emu(180000), Emu(580000), Emu(8780000), Emu(320000),
                  "■ 大泉桜学園15年の実績から見る「小中一貫教育校」の効果検証",
                  COLOR_HEADER_BLUE, text_color=COLOR_WHITE, size_pt=13)

# 4象限
box_y = Emu(950000)
box_h = Emu(1500000)

# 左上：学力
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), box_y, Emu(4400000), Emu(380000),
          text="① 学力向上効果", size_pt=13, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(180000), Emu(1330000), Emu(4400000), Emu(1120000),
            "● 学力調査は区・都平均を継続的に下回る項目が多い\n"
            "● R6中学校：国語60(区62)、数学55(区60)\n"
            "● 区内順位：中学校R6=25位（R5=3位から急落）\n"
            "● 小学校R6=63位（直近年で最低水準）\n"
            "→ 9年間一貫指導でも、学力向上の明確な効果なし",
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 右上：不登校
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(4750000), box_y, Emu(4200000), Emu(380000),
          text="② 不登校（中1ギャップ）の解消", size_pt=13, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(4750000), Emu(1330000), Emu(4200000), Emu(1120000),
            "● 中学部不登校率：R5=10.40%（区6.90%・都7.80%）\n"
            "● 全体の不登校児童生徒数はH28=13人→R5=28人と倍増\n"
            "● 小中一貫の最大目的「中1ギャップ解消」は未達成\n"
            "● 第8・9学年で不登校が集中（R5：第9学年11人）\n"
            "→ 「滑らかな接続」の効果は、データ上確認できない",
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 左下：小規模校
box_y2 = Emu(2530000)
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), box_y2, Emu(4400000), Emu(380000),
          text="③ 小規模校化への対応", size_pt=13, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(180000), Emu(2910000), Emu(4400000), Emu(1120000),
            "● 桜中進学率：80%(R3)→67%(R7)と低下\n"
            "● 中学部入学者数：89人(R4)→61人(R7)へ減少\n"
            "● 区全体は安定（指定校進学率62-66%で推移）\n"
            "● 今後20年区人口は大きな増減なし＝外的要因ではない\n"
            "→ 一貫教育校化しても、生徒流出は止まっていない",
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 右下：コスト・運営面
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(4750000), box_y2, Emu(4200000), Emu(380000),
          text="④ 運営面・施設面の課題", size_pt=13, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(4750000), Emu(2910000), Emu(4200000), Emu(1120000),
            "● 体育施設・特別教室の小中共用に課題（体格差等）\n"
            "● 5・6年生の部活動参加活発化は57%にとどまる\n"
            "● 外部入学者の保護者に不安。情報発信の追加負担\n"
            "● 新規開設には、用地確保・改築・教員配置の高コスト\n"
            "→ 既存中学校の連携強化等、代替策の検討余地あり",
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 矢印
add_arrow_down(slide5, Emu(4322000), Emu(4090000), Emu(500000), Emu(220000))

# 最終結論
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(4360000),
          Emu(8780000), Emu(680000),
          text="【結論】 新規の小中一貫教育校開設は、教育的効果の観点から積極的に推進すべき施策ではない。\n"
               "大泉桜学園15年の検証データは、学力・不登校・小規模校対策のいずれにおいても明確な優位性を示していない。\n"
               "既存校の小中連携強化、教員加配、ICT・少人数指導の充実等、より費用対効果の高い施策を優先すべき。",
          size_pt=12, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_DARK_BLUE, anchor=MSO_ANCHOR.MIDDLE)

# Save
out_path = "/home/user/con30/output/小中一貫教育（大泉桜学園）_改良版.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
print(f"Total slides: {len(prs.slides)}")
