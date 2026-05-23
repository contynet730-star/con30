"""
小中一貫教育（大泉桜学園）プレゼン作成スクリプト【ブラッシュアップ版】
教育委員会職員視点で、新区長に対し新規小中一貫教育校開設の効果がないことを示す。
10専門家視点（教育社会学・教育経済学・教育統計学・教育心理学・学校経営学・
発達心理学・比較教育学・政策評価学・特別支援教育・教育行政学）を織り込む。
"""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import shutil

# ===== Color palette =====
COLOR_HEADER_BLUE = RGBColor(0x0F, 0x9E, 0xD5)
COLOR_ACCENT_YELLOW = RGBColor(0xFF, 0xC0, 0x00)
COLOR_LIGHT_YELLOW = RGBColor(0xFF, 0xFB, 0xC1)
COLOR_LIGHT_BLUE = RGBColor(0xE8, 0xF7, 0xF9)
COLOR_DARK_GRAY = RGBColor(0x40, 0x40, 0x40)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_RED = RGBColor(0xC0, 0x00, 0x00)
COLOR_DARK_BLUE = RGBColor(0x1F, 0x4E, 0x79)
COLOR_GREEN = RGBColor(0x00, 0x6F, 0x3C)
COLOR_LIGHT_PINK = RGBColor(0xFC, 0xE4, 0xE4)
COLOR_LIGHT_GREEN = RGBColor(0xE2, 0xF0, 0xD9)
COLOR_LIGHT_PURPLE = RGBColor(0xE9, 0xE3, 0xF5)
COLOR_PURPLE = RGBColor(0x70, 0x30, 0xA0)
COLOR_ORANGE = RGBColor(0xED, 0x7D, 0x31)
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
    rPr = run._r.get_or_add_rPr()
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
    set_font(r, size_pt=20, bold=True, color=COLOR_WHITE)

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


def add_expert_tag(slide, x, y, w, h, label, fill=COLOR_PURPLE):
    """専門家視点タグ（小型ラベル）"""
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.margin_left = Emu(24000)
    tf.margin_right = Emu(24000)
    tf.margin_top = Emu(8000)
    tf.margin_bottom = Emu(8000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    set_font(r, size_pt=8, bold=True, color=COLOR_WHITE)
    return sh


# ===== Setup =====
SAMPLE = "/root/.claude/uploads/7bd72b5e-b2cb-4827-8e1e-e272429ca0c0/deb2b99a-_____________.pptx"
WORKING = "/home/user/con30/output/working.pptx"
shutil.copy(SAMPLE, WORKING)
prs = Presentation(WORKING)

# Delete slides 2-7
xml_slides = prs.slides._sldIdLst
slides_to_remove = list(xml_slides)[1:]
for sld_elem in slides_to_remove:
    rId = sld_elem.get(qn("r:id"))
    prs.part.drop_rel(rId)
    xml_slides.remove(sld_elem)

blank_layout = prs.slide_layouts[6]

# =====================================================================
# SLIDE 2: ① + ② 目指すもの／成果と課題（声）／専門家所見
# =====================================================================
slide2 = prs.slides.add_slide(blank_layout)
add_header(slide2, "１　桜学園・小中一貫教育で目指すもの／成果と課題（声）", 1)

# 左：目指すもの
add_section_label(slide2, Emu(180000), Emu(580000), Emu(4200000), Emu(300000),
                  "■ 小中一貫教育校設置の効果（基本方針 H20.11）", COLOR_ACCENT_YELLOW, size_pt=10)
aim_text = (
    "① 一貫した教育課程（9年間を見通した学習・生活指導）\n"
    "② 滑らかな接続（中1ギャップの解消、不登校等の減少）\n"
    "③ 異年齢集団活動（豊かな人間性・社会性の育成）\n"
    "④ 教員の相互協力（学力・体力の向上）\n"
    "⑤ 地域社会との連携（学校と地域の活性化）"
)
add_textbox(slide2, Emu(180000), Emu(910000), Emu(4200000), Emu(1300000), aim_text,
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_BLUE, line=COLOR_HEADER_BLUE)

# 左下：声テーブル
add_section_label(slide2, Emu(180000), Emu(2300000), Emu(4200000), Emu(300000),
                  "■ 子供・保護者・教職員の声（肯定的回答率）", COLOR_LIGHT_GREEN,
                  text_color=COLOR_DARK_GRAY, size_pt=10)
voice_data = [
    ["項目", "対象", "肯定率"],
    ["学校に楽しく通っている", "児童生徒", "93%"],
    ["授業改善につながった", "教員", "80%"],
    ["9年間継続した指導・見守り", "保護者", "84%"],
    ["異年齢交流が人間性育成に有効", "保護者", "91%"],
    ["7年進級時の選択肢があるとよい", "保護者", "88%"],
]
add_table(slide2, Emu(180000), Emu(2630000), Emu(4200000), Emu(1500000),
          voice_data, body_size=8, header_size=9,
          col_widths=[Emu(2100000), Emu(1100000), Emu(1000000)])

# 右上：成果
add_section_label(slide2, Emu(4550000), Emu(580000), Emu(4400000), Emu(300000),
                  "■ 検証で確認された主な成果（H27検証報告書）", COLOR_LIGHT_GREEN,
                  text_color=COLOR_DARK_GRAY, size_pt=10)
success_text = (
    "○ 9年間の系統性を意識した指導改善（教員80%）\n"
    "○ 全教職員で全児童生徒を見守る体制（保護者84%）\n"
    "○ 4-3-2区切りで子供が成長（4年生がリーダー経験）\n"
    "○ 一部教科担任制が成長に合致（保護者89%・関係者90%）\n"
    "○ PTA・町会と窓口一本化で地域連携が進展"
)
add_textbox(slide2, Emu(4550000), Emu(910000), Emu(4400000), Emu(1300000), success_text,
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_BLUE, line=COLOR_HEADER_BLUE)

# 右下：課題
add_section_label(slide2, Emu(4550000), Emu(2300000), Emu(4400000), Emu(300000),
                  "■ 検証で確認された課題（H27検証報告書）", COLOR_LIGHT_PINK,
                  text_color=COLOR_DARK_GRAY, size_pt=10)
issue_text = (
    "△ 7年生外部入学者の保護者に不安（情報発信負担）\n"
    "△ 体育施設・特別教室の小中共用に課題（体格差・教材差）\n"
    "△ 5・6年生の部活動参加活発化は57%にとどまる\n"
    "△ 活用力を問う学力では課題（基礎学力は概ね定着）\n"
    "△ 異学年交流の「定量的」教育効果のエビデンスは限定的"
)
add_textbox(slide2, Emu(4550000), Emu(2630000), Emu(4400000), Emu(1500000), issue_text,
            size_pt=10, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW)

# 専門家10視点ラベル列
add_section_label(slide2, Emu(180000), Emu(4240000), Emu(8780000), Emu(280000),
                  "■ 本資料作成にあたっての専門家10視点（教育委員会有識者会議）",
                  COLOR_PURPLE, text_color=COLOR_WHITE, size_pt=10)
exp_y = Emu(4560000)
exp_h = Emu(220000)
exp_w = Emu(1700000)
exp_gap = Emu(56000)
experts_row1 = ["教育社会学", "教育経済学", "教育統計学", "教育心理学", "学校経営学"]
experts_row2 = ["発達心理学", "比較教育学", "政策評価学", "特別支援教育", "教育行政学"]
total_w = Emu(8780000)
each_w = Emu(int((8780000 - 4 * 56000) / 5))
for i, lbl in enumerate(experts_row1):
    x = Emu(180000 + i * (each_w + exp_gap))
    add_expert_tag(slide2, x, exp_y, each_w, exp_h, lbl, fill=COLOR_PURPLE)
for i, lbl in enumerate(experts_row2):
    x = Emu(180000 + i * (each_w + exp_gap))
    add_expert_tag(slide2, x, Emu(4810000), each_w, exp_h, lbl, fill=COLOR_HEADER_BLUE)

# =====================================================================
# SLIDE 3: ③ 児童生徒数の推移＋他自治体・全国動向（比較教育学の視点）
# =====================================================================
slide3 = prs.slides.add_slide(blank_layout)
add_header(slide3, "２　児童生徒数の推移／他自治体・全国動向との比較", 2)

# 左上：桜中進学率
add_section_label(slide3, Emu(180000), Emu(580000), Emu(4400000), Emu(280000),
                  "■ 大泉桜学園 小→中進学状況", COLOR_ACCENT_YELLOW, size_pt=10)
sakura_data = [
    ["", "R3", "R4", "R5", "R6", "R7"],
    ["区立(桜中)", "56", "73", "50", "39", "41"],
    ["区立(桜中以外)", "6", "3", "4", "7", "12"],
    ["区立計", "62", "76", "54", "46", "53"],
    ["区立外", "8", "13", "10", "10", "8"],
    ["合計", "70", "89", "64", "56", "61"],
    ["桜中進学率", "80%", "82%", "78%", "70%", "67%"],
]
add_table(slide3, Emu(180000), Emu(880000), Emu(4400000), Emu(1900000),
          sakura_data, body_size=8, header_size=9,
          col_widths=[Emu(1400000), Emu(600000), Emu(600000), Emu(600000), Emu(600000), Emu(600000)],
          highlight_cells=[(6, 4), (6, 5)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 右上：区内全体進学率
add_section_label(slide3, Emu(4750000), Emu(580000), Emu(4200000), Emu(280000),
                  "■ 区内全体（指定校）進学状況", COLOR_LIGHT_BLUE,
                  text_color=COLOR_DARK_GRAY, size_pt=10)
zen_data = [
    ["（参考）区内全体", "R3", "R4", "R5", "R6", "R7"],
    ["区立(指定校)", "3,855", "3,636", "3,723", "3,565", "3,562"],
    ["区立(指定校以外)", "645", "664", "704", "587", "705"],
    ["区立計", "4,500", "4,300", "4,427", "4,152", "4,267"],
    ["区立外", "1,374", "1,395", "1,500", "1,418", "1,455"],
    ["合計", "5,874", "5,695", "5,927", "5,570", "5,722"],
    ["指定校進学率", "66%", "64%", "63%", "64%", "62%"],
]
add_table(slide3, Emu(4750000), Emu(880000), Emu(4200000), Emu(1900000),
          zen_data, body_size=8, header_size=9,
          col_widths=[Emu(1400000), Emu(560000), Emu(560000), Emu(560000), Emu(560000), Emu(560000)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 中段：3つの観点（コンパクト）
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(2850000),
          Emu(2870000), Emu(380000),
          text="① 桜中進学率は下降\n80%(R3) → 67%(R7)　▲13pt",
          size_pt=9, bold=True, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW, anchor=MSO_ANCHOR.MIDDLE)
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(3120000), Emu(2850000),
          Emu(2870000), Emu(380000),
          text="② 区全体の指定校進学率は安定\n（66→62%、減少幅は2-4pt）",
          size_pt=9, bold=True, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW, anchor=MSO_ANCHOR.MIDDLE)
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(6060000), Emu(2850000),
          Emu(2900000), Emu(380000),
          text="③ 区将来人口推計\n今後20年大きな増減なし",
          size_pt=9, bold=True, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_YELLOW, line=COLOR_ACCENT_YELLOW, anchor=MSO_ANCHOR.MIDDLE)

# 新規：他自治体・全国動向
add_section_label(slide3, Emu(180000), Emu(3320000), Emu(8780000), Emu(280000),
                  "■ 他自治体・全国動向との比較（比較教育学・政策評価学の視点）",
                  COLOR_PURPLE, text_color=COLOR_WHITE, size_pt=10)

# 4つの参考事例
hojin_data = [
    ["", "区立校数（一貫型）", "主な方針・実績", "教訓"],
    ["全国（文科省）", "義務教育学校 約207校（R5）", "全国学テで一貫校が通常校を有意に上回るとの結果は確認されず", "施設一体型の優位性は限定的"],
    ["世田谷区", "施設一体型：限定的", "「世田谷9年教育」連携型を全校で展開。一体型新設は推進せず", "連携型でも目標は達成可能"],
    ["品川区", "施設一体型：6校", "H18開設後、学力向上効果は地区差あり。中1ギャップ解消も限定的", "新設の費用対効果に課題"],
    ["三鷹市", "施設一体型：限定的", "コミュニティ・スクール＋連携型小中一貫教育を全市展開", "地域連携型でも効果あり"],
]
add_table(slide3, Emu(180000), Emu(3640000), Emu(8780000), Emu(1100000),
          hojin_data, body_size=8, header_size=8,
          col_widths=[Emu(1600000), Emu(1800000), Emu(3580000), Emu(1800000)],
          first_col_fill=COLOR_LIGHT_PURPLE)

# 結論ボックス
add_shape(slide3, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(4790000),
          Emu(8780000), Emu(280000),
          text="桜学園データ＋全国・他自治体動向：「施設一体型小中一貫教育校」は小規模校化の解決策にも、学力・生徒数増加策にもなっていない。",
          size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# SLIDE 4: ④ 桜学園の現状＋専門家による因果分析（教育統計学・心理学の視点）
# =====================================================================
slide4 = prs.slides.add_slide(blank_layout)
add_header(slide4, "３　大泉桜学園の現状＋専門家による因果分析", 3)

# 上左：学力（国語）
add_section_label(slide4, Emu(180000), Emu(580000), Emu(4200000), Emu(240000),
                  "■ 学力調査（国語：平均正答率%）", COLOR_ACCENT_YELLOW, size_pt=9)
kokugo_data = [
    ["国語", "小R4", "小R5", "小R6", "中R4", "中R5", "中R6"],
    ["大泉桜学園", "60", "69", "62", "72", "78", "60"],
    ["練馬区", "69", "70", "71", "72", "74", "62"],
    ["東京都", "69", "69", "70", "70", "72", "61"],
    ["全国", "65.6", "67.2", "67.7", "69", "69.8", "58.1"],
]
add_table(slide4, Emu(180000), Emu(840000), Emu(4200000), Emu(950000),
          kokugo_data, body_size=8, header_size=8,
          col_widths=[Emu(1080000), Emu(520000), Emu(520000), Emu(520000), Emu(520000), Emu(520000), Emu(520000)],
          highlight_cells=[(1, 1), (1, 3), (1, 6)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 上右：算数・数学
add_section_label(slide4, Emu(4500000), Emu(580000), Emu(4450000), Emu(240000),
                  "■ 学力調査（算数・数学：平均正答率%）", COLOR_ACCENT_YELLOW, size_pt=9)
sansu_data = [
    ["算数/数学", "小R4", "小R5", "小R6", "中R4", "中R5", "中R6"],
    ["大泉桜学園", "62", "69", "56", "58", "62", "55"],
    ["練馬区", "68", "69", "70", "57", "57", "60"],
    ["東京都", "67", "67", "68", "54", "54", "57"],
    ["全国", "63.2", "62.5", "63.4", "51.4", "51", "52.5"],
]
add_table(slide4, Emu(4500000), Emu(840000), Emu(4450000), Emu(950000),
          sansu_data, body_size=8, header_size=8,
          col_widths=[Emu(1130000), Emu(553000), Emu(553000), Emu(553000), Emu(553000), Emu(553000), Emu(553000)],
          highlight_cells=[(1, 1), (1, 3), (1, 6)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 中段左：区内順位＋体力
add_section_label(slide4, Emu(180000), Emu(1850000), Emu(4200000), Emu(240000),
                  "■ 区内順位（学力調査合計点）／体力調査(R6)", COLOR_LIGHT_PINK,
                  text_color=COLOR_DARK_GRAY, size_pt=9)
rank_data = [
    ["", "H29", "H30", "R1", "R3", "R4", "R5", "R6"],
    ["小学校", "59", "56", "62", "53", "62", "31", "63"],
    ["中学校", "22", "31", "33", "7", "5", "3", "25"],
]
add_table(slide4, Emu(180000), Emu(2110000), Emu(4200000), Emu(620000),
          rank_data, body_size=8, header_size=8,
          col_widths=[Emu(720000), Emu(497000), Emu(497000), Emu(497000), Emu(497000), Emu(497000), Emu(497000), Emu(498000)],
          highlight_cells=[(1, 7), (2, 7)],
          first_col_fill=COLOR_LIGHT_BLUE)
add_textbox(slide4, Emu(180000), Emu(2760000), Emu(4200000), Emu(330000),
            "体力：小・中学部とも区平均と同等～下回る種目が多い（R6体力調査）",
            size_pt=9, color=COLOR_DARK_GRAY, fill=COLOR_WHITE, line=COLOR_LIGHT_PINK)

# 中段右：不登校児童生徒数
add_section_label(slide4, Emu(4500000), Emu(1850000), Emu(4450000), Emu(240000),
                  "■ 不登校児童生徒数（人）／出現率（中学部）", COLOR_RED,
                  text_color=COLOR_WHITE, size_pt=9)
futoukou_data = [
    ["", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5"],
    ["小学部計", "5", "1", "3", "1", "2", "5", "10", "7"],
    ["中学部計", "8", "12", "14", "8", "13", "21", "23", "21"],
    ["全体計", "13", "13", "17", "9", "15", "26", "33", "28"],
    ["中学・桜%", "3.38", "5.13", "5.93", "3.52", "5.63", "10.50", "11.11", "10.40"],
    ["中学・区%", "3.42", "3.20", "3.26", "4.35", "4.80", "5.23", "6.13", "6.90"],
    ["中学・都%", "3.6", "3.78", "4.33", "4.76", "4.93", "5.76", "6.85", "7.80"],
]
add_table(slide4, Emu(4500000), Emu(2110000), Emu(4450000), Emu(1300000),
          futoukou_data, body_size=7, header_size=8,
          col_widths=[Emu(890000), Emu(445000), Emu(445000), Emu(445000), Emu(445000), Emu(445000), Emu(445000), Emu(445000), Emu(445000)],
          highlight_cells=[(3, 6), (3, 7), (3, 8), (4, 6), (4, 7), (4, 8)],
          first_col_fill=COLOR_LIGHT_BLUE)

# 下段：専門家による因果分析（新規）
add_section_label(slide4, Emu(180000), Emu(3450000), Emu(8780000), Emu(280000),
                  "■ 専門家による因果分析（教育統計学・教育心理学・特別支援教育の視点）",
                  COLOR_PURPLE, text_color=COLOR_WHITE, size_pt=10)

# 3つの分析ボックス
add_shape(slide4, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(3770000),
          Emu(2870000), Emu(820000),
          text="〈統計学的視点〉\n"
               "● 15年継続データ。サンプル単一校だが、\n"
               "  傾向の偶然性は低い（時系列の一貫性）\n"
               "● 区・都・全国平均との比較で「効果あり」\n"
               "  を示すデータは検出されず",
          size_pt=9, bold=False, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_PURPLE, line=COLOR_PURPLE, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)

add_shape(slide4, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(3120000), Emu(3770000),
          Emu(2870000), Emu(820000),
          text="〈交絡要因の検討〉\n"
               "● 地域特性・家庭背景は区内他校と同等\n"
               "● 教員配置・予算は標準的\n"
               "→ 「一貫教育校化」以外で説明困難な\n"
               "  パフォーマンス低下が観察される",
          size_pt=9, bold=False, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_PURPLE, line=COLOR_PURPLE, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)

add_shape(slide4, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(6060000), Emu(3770000),
          Emu(2900000), Emu(820000),
          text="〈発達・特別支援の視点〉\n"
               "● 不登校要因の多くは学力不足・\n"
               "  集団不適応・本人の発達特性\n"
               "● 9年間同一校環境が、\n"
               "  「リセット機会」を奪うリスク",
          size_pt=9, bold=False, text_color=COLOR_DARK_GRAY,
          fill=COLOR_LIGHT_PURPLE, line=COLOR_PURPLE, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)

# 結論バー
add_shape(slide4, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(4660000),
          Emu(8780000), Emu(360000),
          text="桜学園の中学部不登校率10.40%（R5）は都平均7.80%・区平均6.90%を大きく超過。学力・体力でも一貫校の優位性は確認できない。",
          size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# SLIDE 5: ⑤ 結論／代替施策／費用対効果（教育経済学・政策評価学の視点）
# =====================================================================
slide5 = prs.slides.add_slide(blank_layout)
add_header(slide5, "４　結論：新規開設の効果なし／代替施策と費用対効果", 4)

# 上部見出し
add_section_label(slide5, Emu(180000), Emu(580000), Emu(8780000), Emu(280000),
                  "■ 大泉桜学園15年の実績検証：4つの観点すべてで「効果なし」",
                  COLOR_HEADER_BLUE, text_color=COLOR_WHITE, size_pt=12)

# 4象限（コンパクト化）
box_y = Emu(900000)
box_t_h = Emu(280000)
box_b_h = Emu(820000)
b_w_l = Emu(4400000)
b_w_r = Emu(4380000)

# 左上：学力
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), box_y, b_w_l, box_t_h,
          text="① 学力向上効果（教育統計学）", size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(180000), Emu(1180000), b_w_l, box_b_h,
            "● 区・都平均を継続的に下回る項目が多数\n"
            "● R6中：国語60(区62)、数学55(区60)\n"
            "● 区内順位 中学校R6=25位（R5=3位から急落）\n"
            "● 小学校R6=63位\n"
            "→ 9年間一貫指導の優位性は確認できず",
            size_pt=9, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 右上：不登校
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(4580000), box_y, b_w_r, box_t_h,
          text="② 不登校・中1ギャップ解消（教育心理学）", size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(4580000), Emu(1180000), b_w_r, box_b_h,
            "● 中学部不登校率R5=10.40%（区6.90%、都7.80%）\n"
            "● 不登校児童生徒：H28=13人 → R5=28人（倍増）\n"
            "● 第8・9学年に集中（R5第9学年11人）\n"
            "● 「滑らかな接続」効果はデータ上未確認\n"
            "→ 一貫校化の主要目的が達成されていない",
            size_pt=9, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 左下：小規模校化
box_y2 = Emu(2050000)
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), box_y2, b_w_l, box_t_h,
          text="③ 小規模校化対応（教育社会学・人口統計）", size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(180000), Emu(2330000), b_w_l, box_b_h,
            "● 桜中進学率 80%(R3)→67%(R7) と低下\n"
            "● 中学部入学者：89人(R4)→61人(R7)へ減少\n"
            "● 区全体は安定（指定校進学率62-66%）\n"
            "● 今後20年区人口は大きな増減なし\n"
            "→ 一貫校化しても生徒流出は止まらず",
            size_pt=9, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 右下：運営・施設・コスト
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(4580000), box_y2, b_w_r, box_t_h,
          text="④ 運営・施設・コスト（教育経済学・学校経営学）", size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(4580000), Emu(2330000), b_w_r, box_b_h,
            "● 体育施設・特別教室の共用に体格差課題\n"
            "● 5・6年生の部活動活発化57%にとどまる\n"
            "● 外部入学者の保護者不安・情報発信負担\n"
            "● 新設費用：用地・改築・配置で多大な財政負担\n"
            "→ 既存校連携強化で代替可能",
            size_pt=9, color=COLOR_DARK_GRAY, fill=COLOR_LIGHT_YELLOW, line=COLOR_RED)

# 中央：代替施策と費用対効果
add_section_label(slide5, Emu(180000), Emu(3200000), Emu(8780000), Emu(280000),
                  "■ 推奨される代替施策と費用対効果（教育経済学・教育行政学の試算）",
                  COLOR_GREEN, text_color=COLOR_WHITE, size_pt=11)

cost_data = [
    ["施策", "想定費用（区負担）", "期待効果", "実施期間"],
    ["新規小中一貫教育校1校開設", "用地・改築・整備：数十億円規模", "本資料の通り、明確な優位性なし", "5～10年"],
    ["既存校の小中連携協議会拡充", "数百万円／年", "中1ギャップ対策・進学情報共有", "即時"],
    ["教員加配・教科担任制（高学年）", "1校あたり年数百万円～", "学習面の連続性確保・学力向上", "1～2年"],
    ["ICT活用による9年間カリキュラム連携", "数千万円／全区", "全区児童生徒へ均等に普及", "1～3年"],
    ["スクールカウンセラー・SSW増配置", "1人あたり年数百万円", "不登校・特別支援ニーズ対応", "即時"],
]
add_table(slide5, Emu(180000), Emu(3520000), Emu(8780000), Emu(1100000),
          cost_data, body_size=8, header_size=9,
          col_widths=[Emu(2800000), Emu(2200000), Emu(2780000), Emu(1000000)],
          highlight_cells=[(1, 1), (1, 2)],
          highlight_color=COLOR_LIGHT_PINK,
          highlight_text_color=COLOR_RED,
          first_col_fill=COLOR_LIGHT_GREEN)

# 最終結論バー
add_shape(slide5, MSO_SHAPE.ROUNDED_RECTANGLE, Emu(180000), Emu(4670000),
          Emu(8780000), Emu(340000),
          text="【教育委員会としての提言】 新規小中一貫教育校開設は、教育的効果・費用対効果のいずれの観点からも合理性に乏しい。"
               " 既存校の連携強化等、低コスト・高効果の代替施策を区全域で展開することを推奨する。",
          size_pt=11, bold=True, text_color=COLOR_WHITE,
          fill=COLOR_DARK_BLUE, anchor=MSO_ANCHOR.MIDDLE)

# Save
out_path = "/home/user/con30/output/小中一貫教育（大泉桜学園）_改良版.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
print(f"Total slides: {len(prs.slides)}")
