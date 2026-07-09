# -*- coding: utf-8 -*-
"""
児童生徒の自殺企図に関する相談対応フロー図（こたエール経由）を PowerPoint で生成するスクリプト。
A4縦・縦型フロー（上から下）・3ページ構成。
  1ページ目: フェーズ1 覚知〜児童生徒の特定
  2ページ目: フェーズ2 警察との連携〜安否確認
  3ページ目: フェーズ3 事後の支援 ＋ 根拠・参考資料
"""
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE as MSO_LINE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

JP_FONT = "Meiryo UI"
TEXT_DARK = RGBColor(0x1F, 0x29, 0x37)
LINE_DARK = RGBColor(0x37, 0x47, 0x5A)
BADGE_FILL = RGBColor(0x2F, 0x54, 0x96)
NOTE_EDGE = RGBColor(0x8A, 0x93, 0x9E)
NOTE_FILL = RGBColor(0xF4, 0xF6, 0xF8)
WARN_EDGE = RGBColor(0xB0, 0x21, 0x21)
WARN_FILL = RGBColor(0xFD, 0xEC, 0xEC)
ORANGE = RGBColor(0xC0, 0x63, 0x0C)

# 主体ごとの配色: (表示名, 枠色, 塗り色)
ACTORS = {
    "student": ("児童生徒（区立小中学校）",        RGBColor(0xB9, 0x77, 0x0E), RGBColor(0xFD, 0xEB, 0xD0)),
    "kotaeru": ("こたエール職員",                  RGBColor(0x21, 0x61, 0x8C), RGBColor(0xD6, 0xEA, 0xF8)),
    "police":  ("警察署（練馬区内管轄）",          RGBColor(0x1E, 0x66, 0x41), RGBColor(0xD1, 0xE7, 0xDD)),
    "school":  ("当該校（校長・副校長）",          RGBColor(0x9A, 0x76, 0x0A), RGBColor(0xFC, 0xF3, 0xCF)),
    "shidou":  ("教育指導課 指導主事",             RGBColor(0x6C, 0x34, 0x83), RGBColor(0xE8, 0xDA, 0xEF)),
    "ict":     ("教育ICT環境整備係（教育施設課）", RGBColor(0x51, 0x5A, 0x5A), RGBColor(0xE5, 0xE8, 0xE8)),
    "joint":   ("当該校・教育指導課",              RGBColor(0x6C, 0x34, 0x83), RGBColor(0xE8, 0xDA, 0xEF)),
}

SLIDE_W, SLIDE_H = Cm(21.0), Cm(29.7)
MG = 1.0                 # 左右余白
BOX_X, BOX_W = 1.0, 11.4  # 本流ボックス
NOTE_X, NOTE_W = 12.9, 7.1  # 補足ボックス
CENTER = BOX_X + BOX_W / 2


def set_text(tf, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True,
             space_after=2):
    """lines: [(text, size, bold, color), ...] 1要素=1段落。\n は段落内改行。"""
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(0.15)
    tf.margin_top = tf.margin_bottom = Cm(0.08)
    for i, (text, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = 1.05
        p.space_after = Pt(space_after)
        for j, seg in enumerate(text.split("\n")):
            r = p.add_run()
            r.text = seg
            if j > 0:
                br = r._r.makeelement(qn('a:br'), {})
                r._r.addprevious(br)
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = JP_FONT
            rPr = r._r.get_or_add_rPr()
            ea = rPr.makeelement(qn('a:ea'), {})
            ea.set('typeface', JP_FONT)
            rPr.append(ea)


def add_box(slide, shape_type, x, y, w, h, fill, line_color, line_w=1.25):
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


def add_arrow(slide, x1, y1, x2, y2, color=LINE_DARK, weight=2.0, dash=None):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(weight)
    if dash:
        conn.line.dash_style = dash
    conn.shadow.inherit = False
    ln = conn.line._get_or_add_ln()
    tail = ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
    ln.append(tail)
    return conn


def add_line(slide, x1, y1, x2, y2, color=NOTE_EDGE, weight=1.0, dash=MSO_LINE.DASH):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(weight)
    conn.line.dash_style = dash
    conn.shadow.inherit = False
    return conn


def badge(slide, cx, cy, num):
    d = 0.78
    b = add_box(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, BADGE_FILL, RGBColor(0xFF, 0xFF, 0xFF), 1.5)
    set_text(b.text_frame, [(str(num), 12, True, RGBColor(0xFF, 0xFF, 0xFF))],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)


def flow_box(slide, num, y, h, actor, title, body=None, note=None, note_h=None,
             x=BOX_X, w=BOX_W, title_size=11, body_size=10):
    """本流の手順ボックス＋番号バッジ＋（任意で）右側の補足ボックスを描く。"""
    name, edge, fill = ACTORS[actor]
    box = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, edge, 1.5)
    box.adjustments[0] = 0.08
    lines = [(f"【{name}】", 9.5, True, edge), (title, title_size, True, TEXT_DARK)]
    if body:
        lines.append((body, body_size, False, TEXT_DARK))
    set_text(box.text_frame, lines, anchor=MSO_ANCHOR.MIDDLE)
    if num is not None:
        badge(slide, x - 0.52, y + h / 2, num)
    if note:
        nh = note_h or h
        nb = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, NOTE_X, y, NOTE_W, nh, NOTE_FILL, NOTE_EDGE, 1.0)
        nb.adjustments[0] = 0.06
        set_text(nb.text_frame, note, anchor=MSO_ANCHOR.MIDDLE)
        add_line(slide, x + w, y + min(h, nh) / 2, NOTE_X, y + min(h, nh) / 2)
    return box


def page_header(slide, page, phase, phase_sub, timing):
    tb = slide.shapes.add_textbox(Cm(MG), Cm(0.4), Cm(21.0 - MG * 2), Cm(1.0))
    set_text(tb.text_frame, [
        ("児童生徒の自殺企図に関する相談対応フロー", 16, True, TEXT_DARK),
    ], align=PP_ALIGN.CENTER)
    sub = slide.shapes.add_textbox(Cm(MG), Cm(1.35), Cm(21.0 - MG * 2), Cm(0.6))
    set_text(sub.text_frame, [
        ("〜こたエール（東京都）にメールで相談が入った場合〜", 10, False, TEXT_DARK),
    ], align=PP_ALIGN.CENTER)
    band = add_box(slide, MSO_SHAPE.RECTANGLE, MG, 2.05, 21.0 - MG * 2, 0.95, BADGE_FILL, None)
    set_text(band.text_frame, [(f"{phase}　{phase_sub}", 12.5, True, RGBColor(0xFF, 0xFF, 0xFF))],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    chip = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 21.0 - MG - 3.6, 2.2, 3.4, 0.65,
                   RGBColor(0xFF, 0xFF, 0xFF), BADGE_FILL, 1.0)
    set_text(chip.text_frame, [(timing, 8.5, True, BADGE_FILL)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    foot = slide.shapes.add_textbox(Cm(MG), Cm(28.9), Cm(21.0 - MG * 2), Cm(0.6))
    set_text(foot.text_frame, [(f"練馬区教育委員会 教育指導課　　{page} / 3", 8.5, False, NOTE_EDGE)],
             align=PP_ALIGN.RIGHT)


def down_arrow(slide, y1, y2, x=CENTER, label=None):
    add_arrow(slide, x, y1, x, y2)
    if label:
        tb = slide.shapes.add_textbox(Cm(x + 0.15), Cm((y1 + y2) / 2 - 0.35), Cm(2.5), Cm(0.6))
        set_text(tb.text_frame, [(label, 10, True, TEXT_DARK)], wrap=False)


def chevron(slide, y, text, w=11.0, x=None):
    x = CENTER - w / 2 if x is None else x
    cv = add_box(slide, MSO_SHAPE.PENTAGON, x, y, w, 1.15, LINE_DARK, None)
    set_text(cv.text_frame, [(text, 11.5, True, RGBColor(0xFF, 0xFF, 0xFF))],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return cv


# ---------------------------------------------------------------- ページ1
def slide1(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    page_header(s, 1, "フェーズ1", "覚知 〜 児童生徒の特定", "覚知後 直ちに")

    # 対応の3原則
    pr = add_box(s, MSO_SHAPE.ROUNDED_RECTANGLE, MG, 3.25, 21.0 - MG * 2, 1.05,
                 RGBColor(0xEA, 0xF0, 0xF8), BADGE_FILL, 1.25)
    pr.adjustments[0] = 0.12
    set_text(pr.text_frame, [
        ("対応の3原則　① 最優先で即時対応　② 一人で抱え込まず組織で対応　③ 時系列で記録（日時・相手・内容）",
         10.5, True, BADGE_FILL)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 緊急時の注意
    wn = add_box(s, MSO_SHAPE.ROUNDED_RECTANGLE, MG, 4.55, 21.0 - MG * 2, 1.15, WARN_FILL, WARN_EDGE, 1.5)
    wn.adjustments[0] = 0.1
    set_text(wn.text_frame, [
        ("⚠ 相談内容から切迫した危険（今まさに実行するおそれ・具体的な方法や日時の記載）が読み取れる場合は、\n以下のフローと並行して、直ちに警察（110番）へ通報する。",
         10, True, WARN_EDGE)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    y = 6.1
    flow_box(s, 1, y, 2.2, "student",
             "こたエール（東京都）に自殺企図に関する相談メールを送信",
             "区が貸与している学習用アカウント（メールアドレス）から送信される。",
             note=[("こたエール＝東京都のネット・スマホ", 9, True, TEXT_DARK),
                   ("トラブル相談窓口。都内在住・在学・在勤の子供等が電話・メール・LINEで相談できる。", 9, False, TEXT_DARK)],
             note_h=2.2)
    down_arrow(s, y + 2.2, y + 2.65)

    y = 8.75
    flow_box(s, 2, y, 2.7, "kotaeru",
             "教育指導課の指導主事へ電話で情報提供",
             "※区貸与アカウントは外部からのメール受信が制限されており、こたエールから本人に返信できないため、教育委員会へ連絡が入る。",
             note=[("指導主事が聞き取ること", 9.5, True, BADGE_FILL),
                   ("□ 相談メールのアドレス\n□ 受信日時\n□ 相談内容（方法・時期の具体性＝切迫度）\n□ 氏名・学校名等の手がかり\n□ こたエールの対応状況・折返し連絡先", 9, False, TEXT_DARK)],
             note_h=3.1)
    down_arrow(s, y + 3.1, y + 3.55)

    y = 12.3
    flow_box(s, 3, y, 2.6, "shidou",
             "直ちに課内で共有（課長・統括指導主事へ報告）し、時系列記録を開始",
             "以後の全てのやり取り（日時・相手・内容）を記録する。対応は必ず複数人で行う。",
             note=[("個人情報の共有はためらわない", 9.5, True, BADGE_FILL),
                   ("人の生命・身体の保護のために必要がある場合、本人（保護者）の同意なく個人情報を提供・取得できる（個人情報保護法第27条第1項第2号）。", 9, False, TEXT_DARK)],
             note_h=2.6)
    down_arrow(s, y + 2.6, y + 3.05)

    # 同時進行ブロック
    gy = 15.5
    grp = add_box(s, MSO_SHAPE.ROUNDED_RECTANGLE, MG - 0.15, gy, 21.0 - MG * 2 + 0.3, 4.5, None, ORANGE, 1.25)
    grp.line.dash_style = MSO_LINE.DASH
    grp.adjustments[0] = 0.04
    tag = add_box(s, MSO_SHAPE.ROUNDED_RECTANGLE, MG + 0.3, gy - 0.35, 3.2, 0.7,
                  RGBColor(0xFF, 0xFF, 0xFF), ORANGE, 1.25)
    set_text(tag.text_frame, [("【同時進行】", 10.5, True, ORANGE)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)

    bw = (21.0 - MG * 2 - 1.7) / 2
    x4, x5 = MG + 0.25, MG + 0.25 + bw + 1.2
    flow_box(s, 4, gy + 0.7, 3.4, "shidou",
             "教育ICT環境整備係（教育施設課）へ児童生徒の照会を依頼",
             "こたエールから聞き取ったメールアドレスを伝える。区貸与アカウントのため、在籍校・学年・氏名を特定できる。",
             x=x4, w=bw, title_size=10.5, body_size=9.5)
    flow_box(s, 5, gy + 0.7, 3.4, "police",
             "警察署から指導主事へ、当該児童生徒の照会の連絡が入る",
             "こたエールは警察にも情報提供している。担当係・担当者名・折返し先を必ず控える。",
             x=x5, w=bw, title_size=10.5, body_size=9.5)

    # 判断（ひし形）
    dy = 20.55
    dw, dh = 9.4, 2.7
    dx = CENTER - dw / 2
    add_arrow(s, x4 + bw / 2, gy + 4.5, CENTER - 1.2, dy + 0.35)
    add_arrow(s, x5 + bw / 2, gy + 4.5, CENTER + 1.2, dy + 0.35)
    dm = add_box(s, MSO_SHAPE.DIAMOND, dx, dy, dw, dh, RGBColor(0xFF, 0xFF, 0xFF), LINE_DARK, 1.75)
    set_text(dm.text_frame, [("この時点で児童生徒を\n特定できているか", 11, True, TEXT_DARK)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # いいえ → 待機ボックス
    nx = 15.5
    nb = add_box(s, MSO_SHAPE.ROUNDED_RECTANGLE, nx, dy + 0.15, 21.0 - MG - nx, 2.4, NOTE_FILL, LINE_DARK, 1.25)
    nb.adjustments[0] = 0.08
    set_text(nb.text_frame, [
        ("照会結果を待つ", 10.5, True, TEXT_DARK),
        ("判明し次第、直ちに警察の担当者へ連絡する。", 9.5, False, TEXT_DARK)],
        anchor=MSO_ANCHOR.MIDDLE)
    add_arrow(s, dx + dw, dy + dh / 2, nx, dy + dh / 2)
    lb = s.shapes.add_textbox(Cm(dx + dw + 0.1), Cm(dy + dh / 2 - 0.75), Cm(2.0), Cm(0.55))
    set_text(lb.text_frame, [("いいえ", 10, True, TEXT_DARK)], wrap=False)

    # はい → 次ページへ
    cy = 24.15
    down_arrow(s, dy + dh, cy, label="はい")
    add_arrow(s, nx + (21.0 - MG - nx) / 2, dy + 2.55, CENTER + 4.5, cy - 0.05)
    chevron(s, cy, "特定できたら → フェーズ2（次ページ）へ")

    note = s.shapes.add_textbox(Cm(MG), Cm(25.7), Cm(21.0 - MG * 2), Cm(2.6))
    set_text(note.text_frame, [
        ("※ 児童生徒の特定を待つ間も、警察・こたエールとの連絡体制を維持する。特定に時間を要する場合は、その旨を警察に伝える。", 9.5, False, TEXT_DARK),
        ("※ 電話でのやり取りは、聞き取った内容を復唱して確認し、直後に記録する。", 9.5, False, TEXT_DARK)])


# ---------------------------------------------------------------- ページ2
def slide2(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    page_header(s, 2, "フェーズ2", "警察との連携 〜 安否確認", "特定後 直ちに")

    y = 3.5
    flow_box(s, 6, y, 3.0, "shidou",
             "警察へ ①学校名 ②学年 ③氏名 を伝える",
             "あわせて、警察が当該校へ連絡する前に、指導主事から当該校へ事前連絡することを警察に伝え、当該校に何を伝えるべきか助言をもらう。",
             note=[("警察からの助言の例", 9.5, True, ORANGE),
                   ("・事前に保護者に連絡しておく\n・保護者からの虐待が疑われるため、警察が連絡するまでは何もしない　等\n→ 助言内容をそのまま記録し、当該校に正確に伝える。", 9, False, TEXT_DARK)],
             note_h=3.0)
    down_arrow(s, y + 3.0, y + 3.45)

    y = 6.95
    flow_box(s, 7, y, 2.9, "shidou",
             "当該校の校長（不在時は副校長）へ電話で事前連絡",
             "・警察から在籍確認の連絡が入ること\n・警察の助言に基づく対応方法（保護者への連絡の要否 等）\n・校内で知る教職員は必要最小限にとどめること",
             note=[("虐待が疑われる場合", 9.5, True, WARN_EDGE),
                   ("保護者への連絡は行わない。学校・教育委員会は子供家庭支援センター・児童相談所へ通告する（児童虐待防止法第6条）。", 9, False, TEXT_DARK)],
             note_h=2.9)
    down_arrow(s, y + 2.9, y + 3.35)

    y = 10.3
    flow_box(s, 8, y, 2.3, "police",
             "当該校へ連絡し、在籍確認",
             "在籍していれば ①氏名 ②保護者の連絡先 ③住所 を確認する。",
             note=[("学校側の準備", 9.5, True, BADGE_FILL),
                   ("指名された対応者（管理職）が指導要録・学籍情報をすぐ出せるよう手元に準備しておく。", 9, False, TEXT_DARK)],
             note_h=2.3)
    down_arrow(s, y + 2.3, y + 2.75)

    y = 13.05
    flow_box(s, 9, y, 2.2, "police",
             "当該児童生徒の自宅を訪問し、安否確認",
             "必要に応じて保護・救急要請等の措置がとられる。",
             note=[("根拠", 9.5, True, BADGE_FILL),
                   ("自殺企図のおそれのある者の保護は、警察官職務執行法第3条に基づき警察が行う。学校・教委は警察の活動を妨げない。", 9, False, TEXT_DARK)],
             note_h=2.6)
    down_arrow(s, y + 2.6, y + 3.05)

    y = 16.1
    flow_box(s, 10, y, 1.9, "police",
             "安否確認の結果を当該校へ連絡")
    down_arrow(s, y + 1.9, y + 2.35)

    y = 18.45
    flow_box(s, 11, y, 3.1, "school",
             "警察からの情報を教育指導課へ報告",
             "①安否・現在の状況　②警察がとった措置　③保護者の状況・受け止め\n④学校が把握している本人の様子（欠席・友人関係・いじめの有無 等）",
             note=[("報告を受けた指導主事は", 9.5, True, BADGE_FILL),
                   ("課長へ報告し、記録を整理。学校の翌日以降の対応（フェーズ3）を校長と確認する。", 9, False, TEXT_DARK)],
             note_h=2.6)

    cy = 22.05
    down_arrow(s, 21.55, cy)
    chevron(s, cy, "フェーズ3（次ページ：事後の支援）へ")

    note = s.shapes.add_textbox(Cm(MG), Cm(23.8), Cm(21.0 - MG * 2), Cm(3.5))
    set_text(note.text_frame, [
        ("※ 安否確認が完了するまでは、指導主事・当該校とも電話が受けられる体制を維持する（勤務時間外に及ぶ場合は連絡先を警察・相互に共有）。", 9.5, False, TEXT_DARK),
        ("※ 本人・保護者への学校からの接触は、必ず警察の助言（手順6）に従う。独自の判断で先に連絡しない。", 9.5, False, TEXT_DARK),
        ("※ 校内・課内とも、情報は「知る必要のある者」に限定して共有し、憶測が広がらないよう管理する。", 9.5, False, TEXT_DARK)])


# ---------------------------------------------------------------- ページ3
def slide3(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    page_header(s, 3, "フェーズ3", "事後の支援・継続的な見守り", "翌日以降 継続")

    y = 3.4
    flow_box(s, 12, y, 5.3, "joint",
             "校内支援体制を立ち上げ、継続的に支援する",
             "・校内委員会（管理職・担任・養護教諭・SC・SSW）で支援方針を決定\n・本人との面接（SC等）、保護者との連携、必要に応じ医療機関へ接続\n・背景にいじめ等が疑われる場合は、いじめ防止対策推進法に基づく対応を並行\n・経過を教育指導課へ定期的に報告（当面は毎日→安定後は週次など）",
             note=[("TALKの原則（本人への関わり方）", 9.5, True, BADGE_FILL),
                   ("Tell：心配していることを言葉で伝える\nAsk：死にたい気持ちの有無を率直に尋ねる\nListen：批判せず、訴えを傾聴する\nKeep safe：一人にせず、安全を確保する", 9, False, TEXT_DARK)],
             note_h=3.4)

    # 根拠・参考資料
    ry = 9.1
    rt = add_box(s, MSO_SHAPE.RECTANGLE, MG, ry, 21.0 - MG * 2, 0.8, RGBColor(0xEA, 0xF0, 0xF8), BADGE_FILL, 1.25)
    set_text(rt.text_frame, [("根拠・参考資料", 12, True, BADGE_FILL)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    rb = add_box(s, MSO_SHAPE.RECTANGLE, MG, ry + 0.8, 21.0 - MG * 2, 7.4,
                 RGBColor(0xFF, 0xFF, 0xFF), BADGE_FILL, 1.25)
    set_text(rb.text_frame, [
        ("【国の手引き・計画】", 10, True, TEXT_DARK),
        ("・文部科学省「教師が知っておきたい子どもの自殺予防」（平成21年3月）…TALKの原則、組織的対応、記録の重要性", 9.5, False, TEXT_DARK),
        ("・文部科学省「子供に伝えたい自殺予防（学校における自殺予防教育導入の手引）」（平成26年7月）", 9.5, False, TEXT_DARK),
        ("・自殺総合対策大綱（令和4年10月 閣議決定）", 9.5, False, TEXT_DARK),
        ("・こども家庭庁ほか「こどもの自殺対策緊急強化プラン」（令和5年6月）", 9.5, False, TEXT_DARK),
        ("【法令】", 10, True, TEXT_DARK),
        ("・個人情報の保護に関する法律 第27条第1項第2号 …生命・身体の保護に必要な場合、本人同意なく第三者提供が可能", 9.5, False, TEXT_DARK),
        ("・児童虐待の防止等に関する法律 第6条 …虐待が疑われる場合の通告義務（子供家庭支援センター・児童相談所等へ）", 9.5, False, TEXT_DARK),
        ("・警察官職務執行法 第3条 …自殺企図のおそれのある者の保護", 9.5, False, TEXT_DARK),
        ("・自殺対策基本法（平成18年法律第85号）", 9.5, False, TEXT_DARK),
        ("【相談窓口】", 10, True, TEXT_DARK),
        ("・こたエール（東京こどもネット・ケータイヘルプデスク）…東京都のネット・スマホトラブル相談窓口", 9.5, False, TEXT_DARK),
    ], space_after=3)

    # 連絡先メモ
    my = ry + 8.8
    mt = add_box(s, MSO_SHAPE.RECTANGLE, MG, my, 21.0 - MG * 2, 0.8, RGBColor(0xFD, 0xF3, 0xE7), ORANGE, 1.25)
    set_text(mt.text_frame, [("主な連絡先（年度当初に記入し、課内で共有しておく）", 12, True, ORANGE)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    mb = add_box(s, MSO_SHAPE.RECTANGLE, MG, my + 0.8, 21.0 - MG * 2, 4.6,
                 RGBColor(0xFF, 0xFF, 0xFF), ORANGE, 1.25)
    set_text(mb.text_frame, [
        ("・こたエール　　　　　　　　　　　　　　TEL：　　　　　　　　　　担当：", 10, False, TEXT_DARK),
        ("・管轄警察署（生活安全課等）　　　　　　TEL：　　　　　　　　　　担当：", 10, False, TEXT_DARK),
        ("・教育ICT環境整備係（教育施設課）　　　 内線：　　　　　　　　　 担当：", 10, False, TEXT_DARK),
        ("・子供家庭支援センター　　　　　　　　　TEL：　　　　　　　　　　担当：", 10, False, TEXT_DARK),
        ("・児童相談所　　　　　　　　　　　　　　TEL：　　　　　　　　　　担当：", 10, False, TEXT_DARK),
        ("・課長・統括指導主事（勤務時間外）　　　TEL：", 10, False, TEXT_DARK),
    ], space_after=6)


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slide1(prs)
    slide2(prs)
    slide3(prs)
    out = "自殺企図相談対応フロー図.pptx"
    prs.save(out)
    print(f"saved: {out}")


if __name__ == "__main__":
    main()
