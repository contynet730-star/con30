"""
小中一貫教育（大泉桜学園）プレゼン作成スクリプト【v4 R7データ反映版】
- 令和7年度のExcel資料3点（体力テスト/不登校・学力詳細/暴力・いじめ・不登校R5-R7推移）を反映
- 不登校R7=34名（過去最多更新）
- 体力R7で順位改善（中2女子9位、小5男12位）等の新データも誠実に併記
"""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import shutil

# ===== 新しい配色（落ち着いたトーン） =====
COLOR_HEADER = RGBColor(0x1F, 0x4E, 0x79)        # 濃紺（ヘッダー）
COLOR_SUBHEAD = RGBColor(0x5B, 0x9B, 0xD5)       # 中間ブルー（小見出し）
COLOR_ACCENT = RGBColor(0xC0, 0x50, 0x4D)        # 落ち着いた赤（強調）
COLOR_POSITIVE = RGBColor(0x4F, 0x81, 0xBD)      # ブルー（肯定要素）
COLOR_CAUTION = RGBColor(0xE6, 0x9B, 0x4D)       # オレンジ（注意/課題）
COLOR_DARK = RGBColor(0x33, 0x33, 0x33)
COLOR_GRAY = RGBColor(0x76, 0x76, 0x76)
COLOR_BORDER = RGBColor(0xBF, 0xBF, 0xBF)
COLOR_BG_LIGHT = RGBColor(0xF7, 0xF7, 0xF7)      # 薄グレー背景
COLOR_BG_BLUE = RGBColor(0xEC, 0xF3, 0xFA)       # 薄ブルー背景
COLOR_BG_RED = RGBColor(0xFB, 0xEE, 0xED)        # 薄赤背景
COLOR_BG_ORANGE = RGBColor(0xFD, 0xF1, 0xE3)     # 薄オレンジ背景
COLOR_BG_GRAY = RGBColor(0xE8, 0xE8, 0xE8)       # ややグレー
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_PURPLE = RGBColor(0x6F, 0x5A, 0x9C)        # 落ち着いた紫（専門家）

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


def add_textbox(slide, x, y, w, h, text, size_pt=11, bold=False, color=COLOR_DARK,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None, line=None, font_name=FONT_JA):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Emu(72000)
    tf.margin_right = Emu(72000)
    tf.margin_top = Emu(36000)
    tf.margin_bottom = Emu(36000)
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
        tb.line.width = Pt(0.5)
    return tb


def add_shape(slide, shape_type, x, y, w, h, text=None, size_pt=11, bold=False,
              text_color=COLOR_DARK, fill=None, line=None, align=PP_ALIGN.CENTER,
              anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_JA, line_width=0.5):
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
        tf.margin_left = Emu(72000)
        tf.margin_right = Emu(72000)
        tf.margin_top = Emu(36000)
        tf.margin_bottom = Emu(36000)
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
    """ヘッダー：細めの濃紺ライン＋タイトル"""
    # 上部に細いライン
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Emu(80000))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_HEADER
    line.line.fill.background()
    line.shadow.inherit = False

    # タイトル左に小番号ボックス
    num_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(180000), Emu(180000),
                                      Emu(280000), Emu(280000))
    num_box.fill.solid()
    num_box.fill.fore_color.rgb = COLOR_HEADER
    num_box.line.fill.background()
    num_box.shadow.inherit = False
    tf = num_box.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(page_num)
    set_font(r, size_pt=14, bold=True, color=COLOR_WHITE)

    # タイトルテキスト
    title = slide.shapes.add_textbox(Emu(520000), Emu(150000),
                                      Emu(7800000), Emu(340000))
    tf = title.text_frame
    tf.margin_left = Emu(36000)
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title_text
    set_font(r, size_pt=18, bold=True, color=COLOR_HEADER)

    # 下線（タイトル下）
    underline = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(180000), Emu(510000),
                                        Emu(8780000), Emu(15000))
    underline.fill.solid()
    underline.fill.fore_color.rgb = COLOR_SUBHEAD
    underline.line.fill.background()
    underline.shadow.inherit = False

    # フッター
    dept = slide.shapes.add_textbox(Emu(180000), Emu(4920000), Emu(3000000), Emu(180000))
    tf = dept.text_frame
    tf.margin_left = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "練馬区教育委員会 教育振興部教育指導課"
    set_font(r, size_pt=8, bold=False, color=COLOR_GRAY)

    num = slide.shapes.add_textbox(Emu(8740000), Emu(4920000), Emu(280000), Emu(180000))
    tf = num.text_frame
    tf.margin_left = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = f"- {page_num} -"
    set_font(r, size_pt=8, bold=False, color=COLOR_GRAY)


def add_section_label(slide, x, y, w, h, text, fill=COLOR_HEADER, text_color=COLOR_WHITE, size_pt=11):
    """セクションラベル：シャープな矩形"""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.margin_left = Emu(90000)
    tf.margin_right = Emu(60000)
    tf.margin_top = Emu(10000)
    tf.margin_bottom = Emu(10000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    set_font(r, size_pt=size_pt, bold=True, color=text_color)
    return sh


def add_table(slide, x, y, w, h, data, col_widths=None, header_fill=COLOR_HEADER,
              header_text_color=COLOR_WHITE, body_size=9, header_size=10, highlight_cells=None,
              highlight_color=COLOR_BG_RED, highlight_text_color=COLOR_ACCENT, body_text_color=COLOR_DARK,
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


def add_chip(slide, x, y, w, h, label, fill=COLOR_PURPLE, text_color=COLOR_WHITE, size_pt=8):
    """専門家タグ用チップ"""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.margin_left = Emu(18000)
    tf.margin_right = Emu(18000)
    tf.margin_top = Emu(6000)
    tf.margin_bottom = Emu(6000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    set_font(r, size_pt=size_pt, bold=True, color=text_color)
    return sh


def add_conclusion_bar(slide, x, y, w, h, text, fill=COLOR_HEADER, size_pt=11):
    return add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h,
                     text=text, size_pt=size_pt, bold=True,
                     text_color=COLOR_WHITE, fill=fill, anchor=MSO_ANCHOR.MIDDLE)


# ===== Setup =====
SAMPLE = "/root/.claude/uploads/7bd72b5e-b2cb-4827-8e1e-e272429ca0c0/deb2b99a-_____________.pptx"
WORKING = "/home/user/con30/output/working.pptx"
shutil.copy(SAMPLE, WORKING)
prs = Presentation(WORKING)

xml_slides = prs.slides._sldIdLst
for sld_elem in list(xml_slides)[1:]:
    rId = sld_elem.get(qn("r:id"))
    prs.part.drop_rel(rId)
    xml_slides.remove(sld_elem)

blank_layout = prs.slide_layouts[6]

# =====================================================================
# SLIDE 2: ① 目指すもの／② 成果と課題（声）　※情報密度を緩和
# =====================================================================
slide2 = prs.slides.add_slide(blank_layout)
add_header(slide2, "桜学園・小中一貫教育で目指すもの／成果と課題", 1)

# 上段：目指すもの（5本柱を横並びで配置・コンパクト）
add_section_label(slide2, Emu(180000), Emu(670000), Emu(8780000), Emu(280000),
                  "  目指すもの ― 基本方針(H20.11)　＊大泉桜学園：H23.4開校（今年度15年目）",
                  COLOR_HEADER, size_pt=11)

aim_items = [
    ("①\n一貫した\n教育課程", "9年間を見通した\n学習・生活指導"),
    ("②\n滑らかな\n接続", "中1ギャップ解消\n不登校の減少"),
    ("③\n異年齢集団\n活動", "人間性・社会性\nの育成"),
    ("④\n教員の\n相互協力", "学力・体力\n向上"),
    ("⑤\n地域社会\n連携", "学校と地域社会\nの活性化"),
]
aim_y = Emu(1010000)
aim_w = Emu(1720000)
aim_h = Emu(900000)
aim_gap = Emu(45000)
for i, (head, body) in enumerate(aim_items):
    x = Emu(180000 + i * (1720000 + 45000))
    # 上：見出し
    add_shape(slide2, MSO_SHAPE.RECTANGLE, x, aim_y, aim_w, Emu(440000),
              text=head, size_pt=10, bold=True, text_color=COLOR_WHITE,
              fill=COLOR_POSITIVE, anchor=MSO_ANCHOR.MIDDLE)
    # 下：本文
    add_shape(slide2, MSO_SHAPE.RECTANGLE, x, Emu(1450000), aim_w, Emu(460000),
              text=body, size_pt=9, bold=False, text_color=COLOR_DARK,
              fill=COLOR_BG_BLUE, line=COLOR_BORDER, anchor=MSO_ANCHOR.MIDDLE)

# 中段：成果（左）／課題（右）　＊2列レイアウト・テキスト短縮
add_section_label(slide2, Emu(180000), Emu(2050000), Emu(4350000), Emu(280000),
                  "  検証で確認された主な成果 ― H27検証報告書",
                  COLOR_POSITIVE, size_pt=10)
success_text = (
    "○ 9年間の系統性を意識した指導改善（教員80%）\n"
    "○ 全教職員で全児童生徒を見守る体制（保護者84%）\n"
    "○ 4-3-2区切りで子供が成長（4年生がリーダー経験）\n"
    "○ 一部教科担任制が成長に合致（保護者89%）\n"
    "○ 学校に楽しく通っている児童生徒 93%"
)
add_textbox(slide2, Emu(180000), Emu(2380000), Emu(4350000), Emu(1500000), success_text,
            size_pt=10, color=COLOR_DARK, fill=COLOR_BG_BLUE, line=COLOR_POSITIVE)

add_section_label(slide2, Emu(4610000), Emu(2050000), Emu(4350000), Emu(280000),
                  "  検証で確認された課題 ― H27検証報告書",
                  COLOR_CAUTION, size_pt=10)
issue_text = (
    "△ 7年生外部入学者の保護者に不安が残る\n"
    "△ 体育施設・特別教室の小中共用に課題（体格差）\n"
    "△ 5・6年生の部活動参加は活発化57%にとどまる\n"
    "△ 活用力を問う学力では課題が残る\n"
    "△ 定量的な教育効果のエビデンスは限定的"
)
add_textbox(slide2, Emu(4610000), Emu(2380000), Emu(4350000), Emu(1500000), issue_text,
            size_pt=10, color=COLOR_DARK, fill=COLOR_BG_ORANGE, line=COLOR_CAUTION)

# 下段：専門家10視点（横並びチップ）
add_section_label(slide2, Emu(180000), Emu(4050000), Emu(8780000), Emu(260000),
                  "  本資料は以下の10専門分野の知見を踏まえて作成",
                  COLOR_PURPLE, size_pt=10)
experts = ["教育社会学", "教育経済学", "教育統計学", "教育心理学", "学校経営学",
           "発達心理学", "比較教育学", "政策評価学", "特別支援教育", "教育行政学"]
exp_y = Emu(4380000)
exp_h = Emu(220000)
each_w_emu = int((8780000 - 9 * 30000) / 10)
for i, lbl in enumerate(experts):
    x = Emu(180000 + i * (each_w_emu + 30000))
    add_chip(slide2, x, exp_y, Emu(each_w_emu), exp_h, lbl,
             fill=COLOR_PURPLE, size_pt=8)

# 下部メッセージ
add_conclusion_bar(slide2, Emu(180000), Emu(4660000), Emu(8780000), Emu(180000),
                   "「効果」と「課題」が併記されており、定量的な教育効果のエビデンスは限定的。"
                   "次ページ以降で大泉桜学園の現状データを確認する。",
                   fill=COLOR_HEADER, size_pt=10)

# =====================================================================
# SLIDE 3: ③ 児童生徒数の推移／他自治体・全国動向との比較
# =====================================================================
slide3 = prs.slides.add_slide(blank_layout)
add_header(slide3, "児童生徒数の推移／他自治体・全国動向との比較", 2)

# 上左：桜中進学率
add_section_label(slide3, Emu(180000), Emu(670000), Emu(4350000), Emu(260000),
                  "  大泉桜学園 小→中進学状況", COLOR_HEADER, size_pt=10)
sakura_data = [
    ["", "R3", "R4", "R5", "R6", "R7"],
    ["区立(桜中)", "56", "73", "50", "39", "41"],
    ["区立(桜中以外)", "6", "3", "4", "7", "12"],
    ["区立計", "62", "76", "54", "46", "53"],
    ["区立外", "8", "13", "10", "10", "8"],
    ["合計", "70", "89", "64", "56", "61"],
    ["桜中進学率", "80%", "82%", "78%", "70%", "67%"],
]
add_table(slide3, Emu(180000), Emu(960000), Emu(4350000), Emu(1850000),
          sakura_data, body_size=8, header_size=9,
          col_widths=[Emu(1300000), Emu(610000), Emu(610000), Emu(610000), Emu(610000), Emu(610000)],
          highlight_cells=[(6, 4), (6, 5)],
          first_col_fill=COLOR_BG_BLUE)

# 上右：区内全体
add_section_label(slide3, Emu(4610000), Emu(670000), Emu(4350000), Emu(260000),
                  "  （参考）区内全体 指定校進学状況", COLOR_SUBHEAD, size_pt=10)
zen_data = [
    ["区内全体", "R3", "R4", "R5", "R6", "R7"],
    ["区立(指定校)", "3,855", "3,636", "3,723", "3,565", "3,562"],
    ["区立(指定校以外)", "645", "664", "704", "587", "705"],
    ["区立計", "4,500", "4,300", "4,427", "4,152", "4,267"],
    ["区立外", "1,374", "1,395", "1,500", "1,418", "1,455"],
    ["合計", "5,874", "5,695", "5,927", "5,570", "5,722"],
    ["指定校進学率", "66%", "64%", "63%", "64%", "62%"],
]
add_table(slide3, Emu(4610000), Emu(960000), Emu(4350000), Emu(1850000),
          zen_data, body_size=8, header_size=9,
          col_widths=[Emu(1400000), Emu(590000), Emu(590000), Emu(590000), Emu(590000), Emu(590000)],
          first_col_fill=COLOR_BG_BLUE)

# 中段：4つの観点
midp_y = Emu(2870000)
mid_w = int((8780000 - 3 * 60000) / 4)
midp_items = [
    ("桜中進学率は下降", "80%(R3) → 67%(R7)\n▲13ポイント", COLOR_BG_ORANGE, COLOR_CAUTION, COLOR_DARK),
    ("区全体は安定", "指定校 66 → 62%\n減少幅は小さい", COLOR_BG_BLUE, COLOR_SUBHEAD, COLOR_DARK),
    ("区将来人口推計", "今後20年\n大きな増減なし", COLOR_BG_BLUE, COLOR_SUBHEAD, COLOR_DARK),
    ("部活動休部が進行", "バスケR6夏休部\n野球R7夏休部", COLOR_BG_RED, COLOR_ACCENT, COLOR_ACCENT),
]
for i, (head, body, bg, border, txtc) in enumerate(midp_items):
    x = Emu(180000 + i * (mid_w + 60000))
    add_shape(slide3, MSO_SHAPE.RECTANGLE, x, midp_y, Emu(mid_w), Emu(260000),
              text=head, size_pt=10, bold=True, text_color=COLOR_WHITE,
              fill=border, anchor=MSO_ANCHOR.MIDDLE)
    add_shape(slide3, MSO_SHAPE.RECTANGLE, x, Emu(3140000), Emu(mid_w), Emu(360000),
              text=body, size_pt=9, bold=False, text_color=txtc,
              fill=bg, line=border, anchor=MSO_ANCHOR.MIDDLE)

# 他自治体テーブル
add_section_label(slide3, Emu(180000), Emu(3580000), Emu(8780000), Emu(260000),
                  "  他自治体・全国動向との比較 ― 比較教育学・政策評価学の視点",
                  COLOR_PURPLE, size_pt=10)
hojin_data = [
    ["", "区立校数（一貫型）", "主な方針・実績", "教訓"],
    ["全国（文科省）", "義務教育学校 約207校（R5）", "全国学テで一貫校が通常校を有意に上回る結果は確認されず", "施設一体型の優位性は限定的"],
    ["世田谷区", "施設一体型：限定的", "「世田谷9年教育」連携型を全校で展開。一体型新設は推進せず", "連携型でも目標は達成可能"],
    ["品川区", "施設一体型：6校", "H18開設後、学力向上効果は地区差あり。中1ギャップ解消も限定的", "新設の費用対効果に課題"],
    ["三鷹市", "施設一体型：限定的", "コミュニティ・スクール＋連携型小中一貫教育を全市展開", "地域連携型でも効果あり"],
]
add_table(slide3, Emu(180000), Emu(3880000), Emu(8780000), Emu(950000),
          hojin_data, body_size=8, header_size=8,
          col_widths=[Emu(1500000), Emu(1800000), Emu(3680000), Emu(1800000)],
          header_fill=COLOR_PURPLE,
          first_col_fill=COLOR_BG_LIGHT)

# 結論バー
add_conclusion_bar(slide3, Emu(180000), Emu(4660000), Emu(8780000), Emu(180000),
                   "桜学園データ＋全国・他自治体動向：施設一体型小中一貫教育校は、"
                   "小規模校化の解決策にも、学力・生徒数の増加策にもなっていない。",
                   fill=COLOR_ACCENT, size_pt=10)

# =====================================================================
# SLIDE 4: ④ 桜学園の現状（学力・体力・不登校）＋R7最新データ
# =====================================================================
slide4 = prs.slides.add_slide(blank_layout)
add_header(slide4, "大泉桜学園の現状（学力・体力・不登校）＋R7最新", 3)

# 学力（国語）
add_section_label(slide4, Emu(180000), Emu(670000), Emu(4350000), Emu(240000),
                  "  学力調査（国語：平均正答率%）", COLOR_HEADER, size_pt=9)
kokugo_data = [
    ["国語", "小R4", "小R5", "小R6", "中R4", "中R5", "中R6"],
    ["大泉桜学園", "60", "69", "62", "72", "78", "60"],
    ["練馬区", "69", "70", "71", "72", "74", "62"],
    ["東京都", "69", "69", "70", "70", "72", "61"],
    ["全国", "65.6", "67.2", "67.7", "69", "69.8", "58.1"],
]
add_table(slide4, Emu(180000), Emu(940000), Emu(4350000), Emu(820000),
          kokugo_data, body_size=8, header_size=8,
          col_widths=[Emu(1110000), Emu(540000), Emu(540000), Emu(540000), Emu(540000), Emu(540000), Emu(540000)],
          highlight_cells=[(1, 1), (1, 3), (1, 6)],
          first_col_fill=COLOR_BG_BLUE)

# 学力（算数・数学）
add_section_label(slide4, Emu(4610000), Emu(670000), Emu(4350000), Emu(240000),
                  "  学力調査（算数・数学：平均正答率%）", COLOR_HEADER, size_pt=9)
sansu_data = [
    ["算数/数学", "小R4", "小R5", "小R6", "中R4", "中R5", "中R6"],
    ["大泉桜学園", "62", "69", "56", "58", "62", "55"],
    ["練馬区", "68", "69", "70", "57", "57", "60"],
    ["東京都", "67", "67", "68", "54", "54", "57"],
    ["全国", "63.2", "62.5", "63.4", "51.4", "51", "52.5"],
]
add_table(slide4, Emu(4610000), Emu(940000), Emu(4350000), Emu(820000),
          sansu_data, body_size=8, header_size=8,
          col_widths=[Emu(1110000), Emu(540000), Emu(540000), Emu(540000), Emu(540000), Emu(540000), Emu(540000)],
          highlight_cells=[(1, 1), (1, 3), (1, 6)],
          first_col_fill=COLOR_BG_BLUE)

# 学力区内順位（小・中）　左
add_section_label(slide4, Emu(180000), Emu(1810000), Emu(4350000), Emu(240000),
                  "  学力合計点 区内順位（年度推移）", COLOR_CAUTION, size_pt=9)
rank_data = [
    ["", "H29", "H30", "R1", "R3", "R4", "R5", "R6"],
    ["小(/65)", "59", "56", "62", "53", "62", "31", "63"],
    ["中(/33)", "22", "31", "33", "7", "5", "3", "25"],
]
add_table(slide4, Emu(180000), Emu(2080000), Emu(4350000), Emu(420000),
          rank_data, body_size=8, header_size=8,
          col_widths=[Emu(660000), Emu(527000), Emu(527000), Emu(527000), Emu(527000), Emu(527000), Emu(527000), Emu(528000)],
          highlight_cells=[(1, 7), (2, 7)],
          first_col_fill=COLOR_BG_BLUE)

# 不登校児童生徒数（R7まで拡張、9年分）　右
add_section_label(slide4, Emu(4610000), Emu(1810000), Emu(4350000), Emu(240000),
                  "  不登校児童生徒数（人）／出現率(中%) ★R7=34名 過去最多",
                  COLOR_ACCENT, size_pt=9)
futoukou_data = [
    ["", "H29", "H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"],
    ["小学部", "1", "3", "1", "2", "5", "10", "7", "14", "17"],
    ["中学部", "12", "14", "8", "13", "21", "23", "21", "18", "17"],
    ["全体計", "13", "17", "9", "15", "26", "33", "28", "32", "34"],
    ["中桜%", "5.13", "5.93", "3.52", "5.63", "10.50", "11.11", "10.40", "-", "-"],
    ["中区%", "3.20", "3.26", "4.35", "4.80", "5.23", "6.13", "6.90", "-", "-"],
    ["中都%", "3.78", "4.33", "4.76", "4.93", "5.76", "6.85", "7.80", "-", "-"],
]
add_table(slide4, Emu(4610000), Emu(2080000), Emu(4350000), Emu(1050000),
          futoukou_data, body_size=7, header_size=8,
          col_widths=[Emu(530000), Emu(382000), Emu(382000), Emu(382000), Emu(382000), Emu(382000), Emu(382000), Emu(382000), Emu(382000), Emu(382000)],
          highlight_cells=[(1, 8), (1, 9), (2, 8), (2, 9), (3, 8), (3, 9), (4, 6), (4, 7)],
          first_col_fill=COLOR_BG_BLUE)

# 体力R6→R7順位推移（新規）　左
add_section_label(slide4, Emu(180000), Emu(2610000), Emu(4350000), Emu(240000),
                  "  体力テスト R6→R7 順位推移 ★R7改善傾向",
                  COLOR_POSITIVE, size_pt=9)
tairyoku_data = [
    ["対象", "R6点", "R6順位", "R7点", "R7順位", "変化"],
    ["小5男子", "50.89", "45/65", "53.58", "12/65", "↑33"],
    ["小5女子", "50.57", "58/65", "52.21", "33/65", "↑25"],
    ["中2男子", "41.08", "12/33", "43.67", "10/33", "↑2"],
    ["中2女子", "42.58", "28/33", "49.36", "9/33", "↑19"],
]
add_table(slide4, Emu(180000), Emu(2880000), Emu(4350000), Emu(900000),
          tairyoku_data, body_size=8, header_size=8,
          col_widths=[Emu(700000), Emu(600000), Emu(750000), Emu(600000), Emu(750000), Emu(950000)],
          header_fill=COLOR_POSITIVE,
          highlight_cells=[(1, 5), (2, 5), (3, 5), (4, 5)],
          highlight_color=COLOR_BG_BLUE, highlight_text_color=COLOR_POSITIVE,
          first_col_fill=COLOR_BG_BLUE)

# いじめ・暴力（R5→R6→R7）　右
add_section_label(slide4, Emu(4610000), Emu(3210000), Emu(4350000), Emu(240000),
                  "  R5→R6→R7 推移：いじめ・暴力（件数）",
                  COLOR_ACCENT, size_pt=9)
ijime_data = [
    ["項目", "R5", "R6", "R7", "傾向"],
    ["いじめ認知（小）", "13", "12", "31", "↑急増"],
    ["いじめ認知（中）", "10", "11", "14", "↑増"],
    ["暴力行為（小）", "1", "1", "0", "─"],
    ["暴力行為（中）", "4", "9", "12", "↑増"],
]
add_table(slide4, Emu(4610000), Emu(3480000), Emu(4350000), Emu(900000),
          ijime_data, body_size=8, header_size=8,
          col_widths=[Emu(1500000), Emu(550000), Emu(550000), Emu(550000), Emu(1200000)],
          header_fill=COLOR_ACCENT,
          highlight_cells=[(1, 4), (2, 4), (4, 4)],
          highlight_color=COLOR_BG_RED, highlight_text_color=COLOR_ACCENT,
          first_col_fill=COLOR_BG_BLUE)

# R7総合所見　左下（体力の下）
add_section_label(slide4, Emu(180000), Emu(3830000), Emu(4350000), Emu(220000),
                  "  R7最新データの総合所見", COLOR_PURPLE, size_pt=9)
shoken_text = (
    "● 不登校R7=34名で過去最多更新（小17・中17）\n"
    "● R6中3単独では不登校2名（区内4位＝少ない側）\n"
    "● 体力R7改善も学年差・年度変動が大きく要観察\n"
    "● いじめ・暴力もR6→R7で増加傾向"
)
add_textbox(slide4, Emu(180000), Emu(4080000), Emu(4350000), Emu(480000),
            shoken_text, size_pt=8, color=COLOR_DARK,
            fill=COLOR_BG_LIGHT, line=COLOR_PURPLE)

# 結論バー
add_conclusion_bar(slide4, Emu(180000), Emu(4660000), Emu(8780000), Emu(180000),
                   "不登校はR7=34名で過去最多更新（小17・中17）。"
                   "体力はR7改善も学力・小規模化の課題は継続。一貫教育の総合的優位性は依然として確認されず。",
                   fill=COLOR_ACCENT, size_pt=10)

# =====================================================================
# SLIDE 5: ⑤ 学校自己評価・改善計画（R6・R7）　※新規追加
# =====================================================================
slide5a = prs.slides.add_slide(blank_layout)
add_header(slide5a, "学校の自己評価と改善計画（R6・R7）", 4)

# 学校ビジョン
add_section_label(slide5a, Emu(180000), Emu(670000), Emu(8780000), Emu(260000),
                  "  学校教育目標／ビジョン　※令和6・7年度 学校経営計画より",
                  COLOR_HEADER, size_pt=11)
vision_text = (
    "目指す学校像 ─ 「笑顔あふれる学校 ～感動の共有～」　／　校長：渡邊重幸\n"
    "桜学精神　1～4学年：元気・チャレンジ・思いやり　／　5～9学年：桜の花よりも華ある人・時機を知る人・愛される人"
)
add_textbox(slide5a, Emu(180000), Emu(960000), Emu(8780000), Emu(460000),
            vision_text, size_pt=10, bold=False, color=COLOR_DARK,
            fill=COLOR_BG_BLUE, line=COLOR_SUBHEAD)

# 4領域評価（成果と課題、横並び4列）
add_section_label(slide5a, Emu(180000), Emu(1490000), Emu(8780000), Emu(260000),
                  "  4領域での成果と課題　※学校関係者評価委員会（評議員11名）による外部評価",
                  COLOR_POSITIVE, size_pt=11)

dom_w = int((8780000 - 3 * 60000) / 4)
dom_y_head = Emu(1790000)
dom_y_body = Emu(2080000)
dom_items = [
    ("① 確かな学力", COLOR_HEADER,
     "○ ICT機器（Canva・Padlet・\nKahoot!等）を積極活用\n"
     "○ 思考・判断・表現力育成のた\nめ意見記入欄を設定\n"
     "△ ICT活用が「楽しさ」優先で\n本質的学習に繋がらない場面"),
    ("② 豊かな心", COLOR_POSITIVE,
     "○ SC・SSW・心のふれあい相\n談員と密に連携\n"
     "○ ふれあい月間アンケートで\n潜在的いじめを早期察知\n"
     "△ 一部児童生徒に話し合い\n参加が偏る"),
    ("③ 健康な生活", COLOR_CAUTION,
     "○ トップアスリート招聘等で\n体力向上を図る\n"
     "○ 食育指導計画に基づき給食\n指導を充実\n"
     "△ 体育施設・特別教室の小中\n共用に体格差課題"),
    ("④ 開かれた学校", COLOR_PURPLE,
     "○ 学年・学級・委員会だより\nで授業様子を発信\n"
     "○ HP・たより・学校公開で\n教育活動を公開\n"
     "△ 家庭・地域への啓発を\n更に強化する必要"),
]
for i, (head, color, body) in enumerate(dom_items):
    x = Emu(180000 + i * (dom_w + 60000))
    add_shape(slide5a, MSO_SHAPE.RECTANGLE, x, dom_y_head, Emu(dom_w), Emu(280000),
              text=head, size_pt=11, bold=True, text_color=COLOR_WHITE,
              fill=color, anchor=MSO_ANCHOR.MIDDLE)
    add_shape(slide5a, MSO_SHAPE.RECTANGLE, x, dom_y_body, Emu(dom_w), Emu(1300000),
              text=body, size_pt=9, bold=False, text_color=COLOR_DARK,
              fill=COLOR_BG_LIGHT, line=color, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)

# R7改善策（左）／校長見解（右）
add_section_label(slide5a, Emu(180000), Emu(3470000), Emu(4350000), Emu(260000),
                  "  R7 改善計画（具体策）", COLOR_POSITIVE, size_pt=10)
r7_kaizen = (
    "● 悉皆研修・OJTで全教員のICTスキル統一\n"
    "● AI集計で授業アンケート分析→PDCA高速化\n"
    "● 週3回いじめ防止対策会議を継続\n"
    "● Slack等で児童生徒変化をリアルタイム共有\n"
    "● 「空白の時間・場所」を物理的に減らす"
)
add_textbox(slide5a, Emu(180000), Emu(3760000), Emu(4350000), Emu(900000),
            r7_kaizen, size_pt=9, color=COLOR_DARK,
            fill=COLOR_BG_BLUE, line=COLOR_POSITIVE)

add_section_label(slide5a, Emu(4610000), Emu(3470000), Emu(4350000), Emu(260000),
                  "  校長の見解（次年度改善に向けて）", COLOR_HEADER, size_pt=10)
kocho_text = (
    "「小中一貫教育校である本校は地域の期待も大きい。\n"
    " 期待に応えるためにも今年度の反省をもとに考えた\n"
    " 改善策を、まずは確実に実行していく。PDCAサイクル\n"
    " を活かして、年度途中でも改善策の妥当性を吟味し、\n"
    " 必要に応じて修正していく。」（R6・R7報告書より要約）"
)
add_textbox(slide5a, Emu(4610000), Emu(3760000), Emu(4350000), Emu(900000),
            kocho_text, size_pt=9, color=COLOR_DARK,
            fill=COLOR_BG_LIGHT, line=COLOR_HEADER)

# 所見
add_conclusion_bar(slide5a, Emu(180000), Emu(4660000), Emu(8780000), Emu(180000),
                   "学校は PDCA を回し改善努力を継続中。"
                   "ただし統計データの傾向は単年の努力では覆らず、施設一体型化の構造的優位性は依然として未確認。",
                   fill=COLOR_PURPLE, size_pt=10)

# =====================================================================
# SLIDE 6: ⑥ 結論／代替施策／費用対効果
# =====================================================================
slide5 = prs.slides.add_slide(blank_layout)
add_header(slide5, "結論：新規開設の効果なし／代替施策と費用対効果", 5)

# 上部見出し
add_section_label(slide5, Emu(180000), Emu(670000), Emu(8780000), Emu(260000),
                  "  大泉桜学園 開校15年（H23.4〜）R7最新データ反映：3観点で「効果なし」／1観点で「部分改善」",
                  COLOR_HEADER, size_pt=11)

# 4象限
box_y = Emu(970000)
box_t_h = Emu(260000)
box_b_h = Emu(780000)
b_w_l = Emu(4350000)
b_w_r = Emu(4400000)

# ① 学力
add_shape(slide5, MSO_SHAPE.RECTANGLE, Emu(180000), box_y, b_w_l, box_t_h,
          text="① 学力向上効果（教育統計学）", size_pt=11, bold=True,
          text_color=COLOR_WHITE, fill=COLOR_ACCENT, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(180000), Emu(1230000), b_w_l, box_b_h,
            "● 区・都平均を継続的に下回る項目が多数\n"
            "● R6中：国語60(区62)、数学55(区60)\n"
            "● 区内順位 中学校R6=25位（R5=3位から急落）\n"
            "● 小学校R6=63位（直近最低水準）\n"
            "→ 開校15年経過も学力向上の効果なし",
            size_pt=9, color=COLOR_DARK, fill=COLOR_BG_RED, line=COLOR_ACCENT)

# ② 不登校
add_shape(slide5, MSO_SHAPE.RECTANGLE, Emu(4610000), box_y, b_w_r, box_t_h,
          text="② 不登校・中1ギャップ解消（教育心理学）", size_pt=11, bold=True,
          text_color=COLOR_WHITE, fill=COLOR_ACCENT, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(4610000), Emu(1230000), b_w_r, box_b_h,
            "● 中学部不登校率R5=10.40%（区6.90%・都7.80%）\n"
            "● 不登校総数：H28=13→R5=28→R6=32→R7=34（過去最多）\n"
            "● いじめ・暴力行為もR6→R7で増加傾向\n"
            "● 「滑らかな接続」効果はデータ上未確認\n"
            "→ R7も悪化継続。一貫校化の主要目的が未達成",
            size_pt=9, color=COLOR_DARK, fill=COLOR_BG_RED, line=COLOR_ACCENT)

# ③ 小規模校化
box_y2 = Emu(2090000)
add_shape(slide5, MSO_SHAPE.RECTANGLE, Emu(180000), box_y2, b_w_l, box_t_h,
          text="③ 小規模校化対応（教育社会学・人口統計）", size_pt=11, bold=True,
          text_color=COLOR_WHITE, fill=COLOR_ACCENT, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(180000), Emu(2350000), b_w_l, box_b_h,
            "● 桜中進学率 80%(R3)→67%(R7) と低下\n"
            "● 中学部入学者：89人(R4)→61人(R7)\n"
            "● 部活動の休部が進行（バスケR6夏・野球R7夏）\n"
            "● 区全体は安定／区人口は今後20年大きな増減なし\n"
            "→ 一貫校化しても流出と小規模化は止まらず",
            size_pt=9, color=COLOR_DARK, fill=COLOR_BG_RED, line=COLOR_ACCENT)

# ④ 体力・運営・コスト
add_shape(slide5, MSO_SHAPE.RECTANGLE, Emu(4610000), box_y2, b_w_r, box_t_h,
          text="④ 体力・運営・コスト（一部肯定／コスト課題継続）", size_pt=11, bold=True,
          text_color=COLOR_WHITE, fill=COLOR_CAUTION, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(slide5, Emu(4610000), Emu(2350000), b_w_r, box_b_h,
            "○ 体力R7：小5男45→12位、中2女28→9位と改善傾向\n"
            "△ ただし学年差・年度差大きく中長期で要検証\n"
            "● 体育施設・特別教室は体格差で運用課題\n"
            "● 新設費用（用地・改築）と運営負荷は継続\n"
            "→ 体力面の効果も既存校連携で再現可能",
            size_pt=9, color=COLOR_DARK, fill=COLOR_BG_ORANGE, line=COLOR_CAUTION)

# 代替施策テーブル
add_section_label(slide5, Emu(180000), Emu(3240000), Emu(8780000), Emu(260000),
                  "  推奨される代替施策と費用対効果 ― 教育経済学・教育行政学の試算",
                  COLOR_POSITIVE, size_pt=11)
cost_data = [
    ["施策", "想定費用（区負担）", "期待効果", "実施期間"],
    ["新規小中一貫教育校1校開設", "用地・改築・整備：数十億円規模", "本資料の通り明確な優位性なし", "5～10年"],
    ["既存校の小中連携協議会拡充", "数百万円／年", "中1ギャップ対策・進学情報共有", "即時"],
    ["教員加配・教科担任制（高学年）", "1校あたり年数百万円～", "学習面の連続性確保・学力向上", "1～2年"],
    ["ICT活用による9年間カリキュラム連携", "数千万円／全区", "全区児童生徒へ均等に普及", "1～3年"],
    ["スクールカウンセラー・SSW増配置", "1人あたり年数百万円", "不登校・特別支援ニーズ対応", "即時"],
]
add_table(slide5, Emu(180000), Emu(3540000), Emu(8780000), Emu(1050000),
          cost_data, body_size=8, header_size=9,
          col_widths=[Emu(2800000), Emu(2200000), Emu(2780000), Emu(1000000)],
          header_fill=COLOR_POSITIVE,
          highlight_cells=[(1, 1), (1, 2)],
          highlight_color=COLOR_BG_RED,
          highlight_text_color=COLOR_ACCENT,
          first_col_fill=COLOR_BG_BLUE)

# 最終結論
add_conclusion_bar(slide5, Emu(180000), Emu(4620000), Emu(8780000), Emu(240000),
                   "【教育委員会提言】R7最新データ反映後も新規開設の総合的効果・費用対効果は確認されず。"
                   "体力改善は要因分析の上、低コスト施策で全区展開を。",
                   fill=COLOR_HEADER, size_pt=10)

# Save
out_path = "/home/user/con30/output/小中一貫教育（大泉桜学園）_改良版.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
print(f"Total slides: {len(prs.slides)}")
