# -*- coding: utf-8 -*-
"""
自殺企図に関する相談対応フロー図（こたエール経由）を PowerPoint で生成するスクリプト。
A4縦1枚・スイムレーン形式（関係機関を横に並べ、時系列を上から下に配置）。
"""
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE as MSO_LINE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

JP_FONT = "Meiryo UI"
TEXT_DARK = RGBColor(0x1F, 0x29, 0x37)
LINE_DARK = RGBColor(0x37, 0x47, 0x5A)
LIFELINE = RGBColor(0xAA, 0xB4, 0xBE)
BADGE_FILL = RGBColor(0x2F, 0x54, 0x96)

# レーン定義: (機関名, 補足, 塗り色, 枠色)
LANES = [
    ("児童生徒\n（自宅）",       "区立小中学校",        RGBColor(0xFD, 0xEB, 0xD0), RGBColor(0xB9, 0x77, 0x0E)),
    ("こたエール",               "東京都の相談窓口",    RGBColor(0xD6, 0xEA, 0xF8), RGBColor(0x21, 0x61, 0x8C)),
    ("警察署",                   "練馬区内管轄",        RGBColor(0xD1, 0xE7, 0xDD), RGBColor(0x1E, 0x66, 0x41)),
    ("当該校",                   "児童生徒の在籍校",    RGBColor(0xFC, 0xF3, 0xCF), RGBColor(0x9A, 0x76, 0x0A)),
    ("教育指導課\n指導主事",     "練馬区教育委員会",    RGBColor(0xE8, 0xDA, 0xEF), RGBColor(0x6C, 0x34, 0x83)),
    ("教育ICT\n環境整備係",      "教育施設課",          RGBColor(0xE5, 0xE8, 0xE8), RGBColor(0x51, 0x5A, 0x5A)),
]

# 手順定義: (発信レーン, 受信レーン, ラベル, フォントpt)
STEPS = [
    (0, 1, "自殺企図に関する相談（メール）", 8),
    (1, 4, "情報提供 ※1", 8),
    (4, 5, "メールアドレスによる\n児童生徒の照会依頼", 8),
    (1, 2, "情報提供", 8),
    (2, 4, "当該児童生徒の照会連絡", 8),
    (4, 2, "児童生徒が判明している場合、\n①学校名 ②学年 ③氏名を伝達\n＋当該校へ事前連絡する旨を伝える", 7.5),
    (2, 4, "当該校へ伝える内容の助言 ※2", 8),
    (4, 3, "事前連絡\n（警察から連絡が入る旨・対応方法）", 8),
    (2, 3, "連絡・在籍確認\n（在籍していれば ①氏名\n②保護者の連絡先 ③住所を確認）", 7.5),
    (2, 0, "自宅を訪問し、安否確認", 8),
    (2, 3, "安否確認の結果を連絡", 8),
    (3, 4, "教育指導課へ報告", 8),
]

FOOTNOTES = [
    "※1　児童生徒が相談に用いたメールアドレスには、こたエール職員から返信することができないため、教育委員会へ情報提供を行う。",
    "※2　助言の例：事前に保護者に連絡しておく／保護者からの虐待が疑われるため、警察が連絡するまでは何もしない　等",
]

SLIDE_W, SLIDE_H = Cm(21.0), Cm(29.7)
MARGIN = 0.7
LANE_W = (21.0 - MARGIN * 2) / len(LANES)

HEADER_TOP, HEADER_H = 1.9, 1.5
LIFE_TOP, LIFE_BOTTOM = HEADER_TOP + HEADER_H, 26.9
ROW_TOP, ROW_BOTTOM = 4.7, 26.5
FOOT_TOP = 27.3


def lane_x(i):
    return MARGIN + LANE_W * (i + 0.5)


def set_text(tf, lines, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=True):
    """lines: [(text, size, bold, color), ...] 1要素=1段落。"""
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(0.05)
    tf.margin_top = tf.margin_bottom = Cm(0.03)
    for i, (text, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = 1.0
        first = True
        for seg in text.split("\n"):
            if not first:
                r = p.add_run()
                r.text = seg
                br = r._r.makeelement(qn('a:br'), {})
                r._r.addprevious(br)
            else:
                r = p.add_run()
                r.text = seg
                first = False
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = JP_FONT
            rPr = r._r.get_or_add_rPr()
            ea = rPr.find(qn('a:ea'))
            if ea is None:
                ea = rPr.makeelement(qn('a:ea'), {})
                rPr.append(ea)
            ea.set('typeface', JP_FONT)


def add_box(slide, shape_type, x, y, w, h, fill, line_color, line_w=1.0):
    sp = slide.shapes.add_shape(shape_type, Cm(x), Cm(y), Cm(w), Cm(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line_color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line_color
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp


def add_arrow(slide, x1, y, x2, color=LINE_DARK, weight=1.5):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y), Cm(x2), Cm(y))
    conn.line.color.rgb = color
    conn.line.width = Pt(weight)
    conn.shadow.inherit = False
    ln = conn.line._get_or_add_ln()
    tail = ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
    ln.append(tail)
    return conn


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # タイトル
    tb = slide.shapes.add_textbox(Cm(MARGIN), Cm(0.35), Cm(21.0 - MARGIN * 2), Cm(1.4))
    set_text(tb.text_frame, [
        ("児童生徒の自殺企図に関する相談への対応フロー", 15, True, TEXT_DARK),
        ("（こたエールにメールで相談が入った場合）", 10, False, TEXT_DARK),
    ])

    # 同時進行の点線枠（手順3〜5をまとめる。矢印より先に描いて背面に置く）
    n_rows = len(STEPS)
    dy = (ROW_BOTTOM - ROW_TOP) / (n_rows - 1)
    y3, y5 = ROW_TOP + dy * 2, ROW_TOP + dy * 4
    grp = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                  MARGIN + 0.1, y3 - 1.35, 21.0 - MARGIN * 2 - 0.2, (y5 - y3) + 1.8,
                  None, RGBColor(0xC0, 0x63, 0x0C), 1.0)
    grp.line.dash_style = MSO_LINE.DASH
    grp.adjustments[0] = 0.05
    tag = slide.shapes.add_textbox(Cm(MARGIN + 0.25), Cm(y3 - 1.32), Cm(3.0), Cm(0.5))
    set_text(tag.text_frame, [("【同時進行】", 8.5, True, RGBColor(0xC0, 0x63, 0x0C))],
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    # レーン見出しとライフライン（縦の点線）
    for i, (name, sub, fill, edge) in enumerate(LANES):
        x = lane_x(i)
        life = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, Cm(x), Cm(LIFE_TOP), Cm(x), Cm(LIFE_BOTTOM))
        life.line.color.rgb = LIFELINE
        life.line.width = Pt(1.0)
        life.line.dash_style = MSO_LINE.DASH
        life.shadow.inherit = False

        hdr = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                      x - (LANE_W - 0.15) / 2, HEADER_TOP, LANE_W - 0.15, HEADER_H,
                      fill, edge, 1.25)
        set_text(hdr.text_frame, [(name, 9.5, True, TEXT_DARK), (f"（{sub}）", 7, False, TEXT_DARK)])

    # 手順（矢印・番号・ラベル）
    for idx, (src, dst, label, fsize) in enumerate(STEPS):
        y = ROW_TOP + dy * idx
        x1, x2 = lane_x(src), lane_x(dst)
        add_arrow(slide, x1, y, x2)

        # ラベル（矢印の直上・下寄せ）
        left, right = min(x1, x2), max(x1, x2)
        width = right - left
        if width < 5.2:
            c = (left + right) / 2
            left, width = c - 2.6, 5.2
        left = max(0.3, min(left, 21.0 - 0.3 - width))
        lab = slide.shapes.add_textbox(Cm(left), Cm(y - 1.48), Cm(width), Cm(1.15))
        set_text(lab.text_frame, [(label, fsize, False, TEXT_DARK)], anchor=MSO_ANCHOR.BOTTOM)

        # 番号バッジ（発信側のライフライン上）
        num = str(idx + 1)
        bw = 0.6 if len(num) == 1 else 0.72
        badge = add_box(slide, MSO_SHAPE.OVAL, x1 - bw / 2, y - 0.3, bw, 0.6, BADGE_FILL, None)
        set_text(badge.text_frame, [(num, 8.5, True, RGBColor(0xFF, 0xFF, 0xFF))], wrap=False)

    # 注記
    foot = slide.shapes.add_textbox(Cm(MARGIN), Cm(FOOT_TOP), Cm(21.0 - MARGIN * 2), Cm(2.2))
    set_text(foot.text_frame, [(t, 7.5, False, TEXT_DARK) for t in FOOTNOTES],
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

    out = "自殺企図相談対応フロー図.pptx"
    prs.save(out)
    print(f"saved: {out}")


if __name__ == "__main__":
    main()
