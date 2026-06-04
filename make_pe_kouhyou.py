from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy

# ── カラーパレット ─────────────────────────────────
NAVY      = RGBColor(0x1F, 0x3A, 0x5F)   # メインの濃紺
BLUE      = RGBColor(0x2E, 0x6F, 0xB8)   # 青
LIGHTBLUE = RGBColor(0xE8, 0xF1, 0xFA)   # 背景薄青
ORANGE    = RGBColor(0xE8, 0x7A, 0x1E)   # アクセント橙
GREEN     = RGBColor(0x2E, 0x8B, 0x57)   # アクセント緑
RED       = RGBColor(0xC0, 0x39, 0x2B)   # アクセント赤
GRAY_TXT  = RGBColor(0x33, 0x33, 0x33)
GRAY_BG   = RGBColor(0xF4, 0xF4, 0xF4)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

FONT_JP = "メイリオ"
FONT_EN = "メイリオ"

# 16:9
prs = Presentation()
prs.slide_width  = Cm(33.867)
prs.slide_height = Cm(19.05)
SW = prs.slide_width
SH = prs.slide_height

blank = prs.slide_layouts[6]


def set_run_font(run, size=14, bold=False, color=GRAY_TXT, name=FONT_JP):
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


def add_textbox(slide, left, top, width, height, text, size=14, bold=False,
                color=GRAY_TXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                line_spacing=1.15):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
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


def add_filled_rect(slide, left, top, width, height, fill_color,
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


def add_round_rect(slide, left, top, width, height, fill_color,
                   line_color=None, line_width=0.75, corner=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(line_width)
    # 角丸調整
    try:
        shp.adjustments[0] = corner
    except Exception:
        pass
    shp.shadow.inherit = False
    return shp


def add_header_band(slide, title, subtitle=None, page_no=None, total=5):
    # 上部の濃紺バンド
    add_filled_rect(slide, Cm(0), Cm(0), SW, Cm(1.9), NAVY)
    # 細いオレンジ帯
    add_filled_rect(slide, Cm(0), Cm(1.9), SW, Cm(0.15), ORANGE)
    # タイトル
    add_textbox(slide, Cm(0.8), Cm(0.25), Cm(28), Cm(1.4),
                title, size=22, bold=True, color=WHITE,
                anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_textbox(slide, Cm(0.8), Cm(1.15), Cm(28), Cm(0.7),
                    subtitle, size=11, color=RGBColor(0xCC, 0xDD, 0xEE),
                    anchor=MSO_ANCHOR.MIDDLE)
    if page_no:
        add_textbox(slide, Cm(31.5), Cm(0.25), Cm(2), Cm(1.4),
                    f"{page_no} / {total}", size=11, color=WHITE,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def add_footer(slide, text="練馬区立小学校 体育研究会  指導講評資料"):
    add_filled_rect(slide, Cm(0), SH - Cm(0.6), SW, Cm(0.6), NAVY)
    add_textbox(slide, Cm(0.8), SH - Cm(0.6), Cm(32), Cm(0.6),
                text, size=9, color=WHITE,
                anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================
# Slide 1 : 表紙
# =====================================================================
s = prs.slides.add_slide(blank)
# 背景
add_filled_rect(s, Cm(0), Cm(0), SW, SH, WHITE)
# 上下のアクセントバー
add_filled_rect(s, Cm(0), Cm(0), SW, Cm(2.2), NAVY)
add_filled_rect(s, Cm(0), Cm(2.2), SW, Cm(0.25), ORANGE)
add_filled_rect(s, Cm(0), SH - Cm(0.9), SW, Cm(0.9), NAVY)
add_filled_rect(s, Cm(0), SH - Cm(1.05), SW, Cm(0.15), ORANGE)

# 中央パネル
add_round_rect(s, Cm(2.5), Cm(4.2), Cm(28.8), Cm(10.0), LIGHTBLUE,
               line_color=BLUE, line_width=0.75, corner=0.04)

# サブタイトル(上)
add_textbox(s, Cm(2.5), Cm(4.8), Cm(28.8), Cm(1.0),
            "中央教育審議会 体育・保健体育、健康、安全ワーキンググループ（第6～9回）から",
            size=16, bold=True, color=BLUE, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)

# メインタイトル
add_textbox(s, Cm(2.5), Cm(6.2), Cm(28.8), Cm(3.2),
            "これからの体育・保健体育\nが目指す方向性",
            size=44, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15)

# キーメッセージ
add_textbox(s, Cm(2.5), Cm(10.4), Cm(28.8), Cm(2.8),
            "──「できる／できない」を超えて、「変化と工夫」を支える授業へ──",
            size=18, bold=True, color=ORANGE, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)

# 帯（タイトルバンド内）
add_textbox(s, Cm(0.8), Cm(0.3), Cm(28), Cm(1.5),
            "練馬区立小学校 体育研究会  指導講評",
            size=22, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# 下部
add_textbox(s, Cm(2.5), Cm(15.0), Cm(28.8), Cm(1.4),
            "本日の授業を踏まえて／次期学習指導要領に向けた論点の共有",
            size=14, color=GRAY_TXT, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)

# フッター
add_textbox(s, Cm(0.8), SH - Cm(0.9), Cm(32), Cm(0.9),
            "練馬区立小学校 体育研究会  指導講評資料",
            size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# Slide 2 : 改訂の3つの視点
# =====================================================================
s = prs.slides.add_slide(blank)
add_header_band(s, "改訂を貫く 3 つの視点",
                "中教審 論点整理（令和7年9月）／第9回WG資料より",
                page_no=2)

# リード
add_textbox(s, Cm(0.8), Cm(2.5), Cm(32.3), Cm(1.2),
            "次期学習指導要領の改訂は、次の 3 つの観点を同時に満たすことを目指している。",
            size=15, color=GRAY_TXT, anchor=MSO_ANCHOR.MIDDLE)

# 3カラム
col_w = Cm(9.8)
col_h = Cm(11.5)
col_top = Cm(4.0)
gap = Cm(0.6)
left0 = Cm(1.0)

cols = [
    ("①  深い学びの実装",        "Excellence",  "卓越性",  NAVY,
     "「高次の資質・能力」を軸に、\n断片的な知識・技能の習得を超え、\n単元を貫いて考え・工夫し続ける学び。\n\n● 統合的な理解（知・技）\n● 総合的な発揮（思・判・表）\nの両輪で授業を構想する。"),
    ("②  多様性の包摂",          "Equity",      "公正性",  GREEN,
     "体力差・性別・障害の有無を超え、\n誰もが楽しさ・喜びを味わえる設計へ。\n\n● 「する」だけでない関わり方\n　（みる・支える・知る）\n● 用具・ルール・場の工夫で\n　全員参加の単元計画を。"),
    ("③  実現可能性の確保",      "Feasibility", "持続性",  ORANGE,
     "教師と子供の双方に「余白」を生む。\n内容の精選と、調整授業時数制度等\nを活かした柔軟な単元設計。\n\n● 種目を限定しない示し方\n● デジタル学習基盤の活用\n● 外部人材・地域連携の充実"),
]

for i, (title, en, jp, color, body) in enumerate(cols):
    left = left0 + (col_w + gap) * i
    # カード
    add_round_rect(s, left, col_top, col_w, col_h, WHITE,
                   line_color=color, line_width=1.5, corner=0.04)
    # ヘッダー帯
    add_filled_rect(s, left, col_top, col_w, Cm(1.7), color)
    add_textbox(s, left, col_top, col_w, Cm(1.7),
                title, size=18, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 英語キーワード
    add_textbox(s, left, col_top + Cm(1.9), col_w, Cm(1.2),
                en, size=22, bold=True, color=color,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(s, left, col_top + Cm(3.1), col_w, Cm(0.7),
                f"（{jp}）", size=12, color=GRAY_TXT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 区切り線
    add_filled_rect(s, left + Cm(1.0), col_top + Cm(4.0),
                    col_w - Cm(2.0), Cm(0.05), color)
    # 本文
    add_textbox(s, left + Cm(0.6), col_top + Cm(4.3),
                col_w - Cm(1.2), col_h - Cm(4.5),
                body, size=12, color=GRAY_TXT, anchor=MSO_ANCHOR.TOP,
                line_spacing=1.25)

# 下段メッセージ
add_round_rect(s, Cm(1.0), Cm(16.0), Cm(31.8), Cm(2.0), LIGHTBLUE,
               line_color=BLUE, line_width=0.75, corner=0.1)
add_textbox(s, Cm(1.5), Cm(16.0), Cm(31.0), Cm(2.0),
            "▶ この 3 つは、いずれか一つを優先するのではなく、 同時に追究する ことが求められている。",
            size=14, bold=True, color=NAVY,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)

# =====================================================================
# Slide 3 : 高次の資質・能力 と「余白」
# =====================================================================
s = prs.slides.add_slide(blank)
add_header_band(s, "授業づくりの軸となる「高次の資質・能力」",
                "種目を「教える」から、資質・能力を「育てる」授業へ",
                page_no=3)

# 左側：構造図
left_panel_l = Cm(0.8)
left_panel_w = Cm(15.5)
add_round_rect(s, left_panel_l, Cm(2.7), left_panel_w, Cm(13.0), LIGHTBLUE,
               line_color=BLUE, line_width=1.0, corner=0.03)
add_textbox(s, left_panel_l, Cm(2.9), left_panel_w, Cm(0.9),
            "■  高次の資質・能力 とは",
            size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)

# 二つのボックス
b_top = Cm(4.1)
b_w = Cm(7.0)
b_h = Cm(4.6)
b_gap = Cm(0.3)

# 統合的な理解
bx1_l = left_panel_l + Cm(0.4)
add_round_rect(s, bx1_l, b_top, b_w, b_h, WHITE,
               line_color=NAVY, line_width=1.2, corner=0.05)
add_filled_rect(s, bx1_l, b_top, b_w, Cm(1.0), NAVY)
add_textbox(s, bx1_l, b_top, b_w, Cm(1.0),
            "統合的な理解（知・技）",
            size=14, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(s, bx1_l + Cm(0.3), b_top + Cm(1.2),
            b_w - Cm(0.6), b_h - Cm(1.4),
            "個々の知識・技能が相互に関連付けられ、\n一般化された「分かった」「できた」の姿。\n\n例：ネット型のゲームの特性等に応じて、\nボール操作と仲間との連携で攻防を展開し\n楽しさや喜びを味わえることを理解する。",
            size=11, color=GRAY_TXT, line_spacing=1.25)

# 総合的な発揮
bx2_l = bx1_l + b_w + b_gap
add_round_rect(s, bx2_l, b_top, b_w, b_h, WHITE,
               line_color=GREEN, line_width=1.2, corner=0.05)
add_filled_rect(s, bx2_l, b_top, b_w, Cm(1.0), GREEN)
add_textbox(s, bx2_l, b_top, b_w, Cm(1.0),
            "総合的な発揮（思・判・表）",
            size=14, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(s, bx2_l + Cm(0.3), b_top + Cm(1.2),
            b_w - Cm(0.6), b_h - Cm(1.4),
            "知識・技能を、状況に応じて選択・組合せ、\n仲間と協働して課題解決していく姿。\n\n例：自他が楽しさや喜びを味わうために\n必要なことを考え、練習方法や作戦・攻防\nを工夫する。",
            size=11, color=GRAY_TXT, line_spacing=1.25)

# 下部メッセージ
add_round_rect(s, left_panel_l + Cm(0.4), Cm(9.2),
               left_panel_w - Cm(0.8), Cm(2.4),
               WHITE, line_color=ORANGE, line_width=1.2, corner=0.05)
add_textbox(s, left_panel_l + Cm(0.7), Cm(9.3),
            left_panel_w - Cm(1.4), Cm(2.2),
            "● どんな「動き」「楽しみ方」「関わり方」を\n　経験させたかったのか── を起点に。\n● 種目はあくまで 資質・能力を育てる「媒介」。",
            size=12, bold=True, color=NAVY,
            anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.35)

# 系統性
add_textbox(s, left_panel_l + Cm(0.4), Cm(11.9),
            left_panel_w - Cm(0.8), Cm(0.7),
            "■  小1～4 と 小5以降 で示し方が変わる",
            size=13, bold=True, color=NAVY)
add_textbox(s, left_panel_l + Cm(0.4), Cm(12.7),
            left_panel_w - Cm(0.8), Cm(2.8),
            "小1～4（運動の基礎を培う時期）：\n　 神経系の発達を踏まえ、「どんな動きを意図するか」を軸に。\n小5以降（多くの領域を経験する時期）：\n　 運動・種目の示し方を柔軟にし、教師の創意工夫を一層促す方向で検討。",
            size=11, color=GRAY_TXT, line_spacing=1.3)

# 右側：余白の創出
right_panel_l = Cm(16.8)
right_panel_w = Cm(16.3)
add_round_rect(s, right_panel_l, Cm(2.7), right_panel_w, Cm(13.0),
               RGBColor(0xFF, 0xF6, 0xE8),
               line_color=ORANGE, line_width=1.0, corner=0.03)
add_textbox(s, right_panel_l, Cm(2.9), right_panel_w, Cm(0.9),
            "■  「 余 白 」 の 創 出 と 教 師 の 創 意 工 夫",
            size=15, bold=True, color=ORANGE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 課題→改善
add_round_rect(s, right_panel_l + Cm(0.5), Cm(4.1),
               right_panel_w - Cm(1.0), Cm(2.6), WHITE,
               line_color=RED, line_width=1.0, corner=0.05)
add_textbox(s, right_panel_l + Cm(0.7), Cm(4.2),
            right_panel_w - Cm(1.4), Cm(0.8),
            "【現状の課題】",
            size=12, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE)
add_textbox(s, right_panel_l + Cm(0.7), Cm(4.9),
            right_panel_w - Cm(1.4), Cm(1.8),
            "「○○運動では××をする」と読める記述が、\n運動自体の目的化を招き、児童生徒の内発的動機\nに基づく活動や多様性の包摂を難しくしている面。",
            size=11, color=GRAY_TXT, line_spacing=1.3)

# 矢印
arrow = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,
                           right_panel_l + Cm(7.6), Cm(6.85),
                           Cm(1.0), Cm(0.7))
arrow.fill.solid()
arrow.fill.fore_color.rgb = ORANGE
arrow.line.fill.background()
arrow.shadow.inherit = False

add_round_rect(s, right_panel_l + Cm(0.5), Cm(7.7),
               right_panel_w - Cm(1.0), Cm(7.5), WHITE,
               line_color=BLUE, line_width=1.0, corner=0.04)
add_textbox(s, right_panel_l + Cm(0.7), Cm(7.8),
            right_panel_w - Cm(1.4), Cm(0.8),
            "【これからの方向性】",
            size=12, bold=True, color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
items = [
    "● 学習指導要領は 期待する 資質・能力 を分かりやすく",
    "● 具体的な活動例は 指導参考資料 で示し、創意工夫を妨げない",
    "● 2 年間で計画的に学ぶ内容の示し方を一層分かりやすく",
    "● 体育と保健の関連性が高い部分は 系統性 を踏まえ精選",
    "● 教師の余白＝児童の 深い学び の余地",
]
add_textbox(s, right_panel_l + Cm(0.9), Cm(8.7),
            right_panel_w - Cm(1.8), Cm(6.5),
            "\n".join(items),
            size=12, color=GRAY_TXT, line_spacing=1.55)

add_footer(s)

# =====================================================================
# Slide 4 : 授業改善の具体的な視点
# =====================================================================
s = prs.slides.add_slide(blank)
add_header_band(s, "授業改善 ─ 4 つの具体的な視点",
                "第6～9回 WG委員意見・主資料より",
                page_no=4)

# 2x2 グリッド
items = [
    ("①  デジタル学習基盤の活用",          BLUE,
     "● 動画で動きの可視化／チームの作戦検討\n● ウェアラブル端末等によるデータ活用\n● クラウド共有で相互の考えを瞬時に交流\n● 機器操作の目的化を避け、身体活動時間を確保",
     "宇山委員・中村委員・第9回主資料"),
    ("②  多様性の包摂（インクルーシブ）",  GREEN,
     "● 用具・ルール・場の工夫で全員参加\n● 「する」だけでなく「みる・支える・知る」\n● 柔らかいボール・ネット高調整等の簡易化\n● 「できる/できない」を超え 変化・工夫 を見取る",
     "宇山委員・岩佐委員・第9回主資料"),
    ("③  体育と保健の関連／教科横断",      ORANGE,
     "● 保健「心の健康」×体育「体ほぐしの運動」\n● 安全教育を各教科×安全で実践（SPS）\n● ヘルスリテラシー：身近な情報を批判的に読む\n● 探究の 4 プロセスで「自分ごと」の保健へ",
     "南委員・藤田委員・渡邉委員・柏原(奈)委員"),
    ("④  調整授業時数制度等の活用",        RED,
     "● 標準授業時数の 1 割以上を調整可能\n● 既存教科上乗せ／新教科／裁量的時間\n● 外部専門家連携・地域人材の活用\n● 体育と保健の関連を深める単元設計",
     "第9回主資料・第8回資料2"),
]

g_top = Cm(2.7)
g_w = Cm(15.9)
g_h = Cm(6.5)
g_left = Cm(0.8)
g_gap_x = Cm(0.5)
g_gap_y = Cm(0.4)

for idx, (title, color, body, ref) in enumerate(items):
    row = idx // 2
    col = idx % 2
    left = g_left + (g_w + g_gap_x) * col
    top  = g_top + (g_h + g_gap_y) * row

    add_round_rect(s, left, top, g_w, g_h, WHITE,
                   line_color=color, line_width=1.5, corner=0.03)
    # ヘッダー
    add_filled_rect(s, left, top, g_w, Cm(1.2), color)
    add_textbox(s, left + Cm(0.4), top, g_w - Cm(0.8), Cm(1.2),
                title, size=16, bold=True, color=WHITE,
                anchor=MSO_ANCHOR.MIDDLE)
    # 本文
    add_textbox(s, left + Cm(0.6), top + Cm(1.5),
                g_w - Cm(1.2), g_h - Cm(2.8),
                body, size=13, color=GRAY_TXT, line_spacing=1.45)
    # 出典
    add_textbox(s, left + Cm(0.6), top + g_h - Cm(0.9),
                g_w - Cm(1.2), Cm(0.7),
                f"［ 参照：{ref} ］", size=9.5, color=color,
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# 下段メッセージ
add_round_rect(s, Cm(0.8), Cm(16.4), Cm(32.3), Cm(1.7), NAVY,
               corner=0.15)
add_textbox(s, Cm(1.2), Cm(16.4), Cm(31.5), Cm(1.7),
            "▶ 4 つはバラバラではなく、 一つの単元の中で重ね合わせて 設計するもの。",
            size=14, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s)

# =====================================================================
# Slide 5 : 本日の授業を踏まえて／練馬の先生方へ
# =====================================================================
s = prs.slides.add_slide(blank)
add_header_band(s, "本日の授業から ── これからの体育研究へ",
                "練馬区の先生方へのメッセージ",
                page_no=5)

# 左：本日の授業から見えたもの（記入欄イメージ）
left_l = Cm(0.8)
left_w = Cm(15.5)
add_round_rect(s, left_l, Cm(2.7), left_w, Cm(13.0), LIGHTBLUE,
               line_color=BLUE, line_width=1.0, corner=0.03)
add_textbox(s, left_l, Cm(2.85), left_w, Cm(0.9),
            "■  本 日 の 授 業 か ら 見 え た こ と",
            size=15, bold=True, color=NAVY,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 観点ごと
points = [
    ("● 高次の資質・能力 の視点",
     "どんな『動き』『楽しみ方』『関わり方』が育まれたか／\n単元のどの場面で『統合的な理解』『総合的な発揮』が現れたか。"),
    ("● 多様性の包摂 の視点",
     "運動を苦手とする児童も含め、全員が参加し、自分の『変化』を実感\nできていたか。用具・ルール・場の工夫は機能していたか。"),
    ("● デジタル基盤の活用 の視点",
     "ICT は『何のため』に使われていたか／機器操作の目的化が起きていな\nかったか。身体活動時間は確保されていたか。"),
    ("● 余白と創意工夫 の視点",
     "教師の意図と、子供たちの試行錯誤の余地が両立していたか。\n指導と評価の計画が、見取りたい姿と整合していたか。"),
]
y = 4.0
for title, body in points:
    add_textbox(s, left_l + Cm(0.5), Cm(y), left_w - Cm(1.0), Cm(0.8),
                title, size=12.5, bold=True, color=BLUE)
    add_textbox(s, left_l + Cm(0.9), Cm(y + 0.7),
                left_w - Cm(1.4), Cm(1.9),
                body, size=11, color=GRAY_TXT, line_spacing=1.35)
    y += 2.65

# 右：メッセージ
right_l = Cm(16.8)
right_w = Cm(16.3)
add_round_rect(s, right_l, Cm(2.7), right_w, Cm(8.4), NAVY, corner=0.03)
add_textbox(s, right_l, Cm(3.0), right_w, Cm(1.1),
            "練馬区の先生方へ",
            size=20, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# オレンジ線
add_filled_rect(s, right_l + Cm(5.0), Cm(4.2),
                right_w - Cm(10.0), Cm(0.08), ORANGE)

msg = (
    "種目を 教える 授業 から、\n"
    "資質・能力を 育てる 授業 へ。\n\n"
    "「できる／できない」の二元論を超え、\n"
    "子供一人一人の 変化 と 工夫 を見取り、\n"
    "それを支える指導と評価を、\n"
    "練馬の研究の中で 一緒に 描いていきたい。"
)
add_textbox(s, right_l + Cm(1.0), Cm(4.6),
            right_w - Cm(2.0), Cm(6.3),
            msg, size=14, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.5)

# 出典パネル
add_round_rect(s, right_l, Cm(11.4), right_w, Cm(4.3),
               WHITE, line_color=NAVY, line_width=1.0, corner=0.04)
add_textbox(s, right_l + Cm(0.5), Cm(11.5), right_w - Cm(1.0), Cm(0.8),
            "▼ 参考資料 ／ 主な出典",
            size=12, bold=True, color=NAVY,
            anchor=MSO_ANCHOR.MIDDLE)
ref = (
    "・中央教育審議会 体育・保健体育、健康、安全WG\n"
    "  　第6回（R8.1.16）／第7回（R8.2.19）／第8回（R8.3.27）／第9回（R8.4.24）\n"
    "・教育課程企画特別部会 論点整理（令和7年9月）\n"
    "・第9回WG資料「体育・保健体育の指導と評価の改善・充実について」"
)
add_textbox(s, right_l + Cm(0.7), Cm(12.3),
            right_w - Cm(1.4), Cm(3.3),
            ref, size=10.5, color=GRAY_TXT, line_spacing=1.4)

add_footer(s)

# 保存
out_path = "/home/user/con30/体育・保健体育の方向性_指導講評.pptx"
prs.save(out_path)
print("Saved:", out_path)
