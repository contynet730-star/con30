# -*- coding: utf-8 -*-
"""
今井様あて「学校公開日（土曜授業）の日程について（回答）」改訂版（500字程度）。

上司指示：「土曜日の背景・変えられないこと等に触れないと納得されない」を踏まえ、
  ① 背景（※区の取組のみ。国・都の経緯は割愛）
  ② 実施日を個別に変更できない理由（変えられないこと）
  ③ 令和8年度からの日程柔軟化（前向きな対応）
を、本文おおむね500字程度に凝縮して構成。
出典：添付3資料（PDF＝土曜授業の在り方／PPT＝R8の在り方／当初回答案docx）
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ── 本文テキスト（文字数管理のため変数化）──────────────
FRONT = [
    "　日頃より、練馬区の学校教育にご理解とご協力を賜り、厚く御礼申し上げます。",
    "　このたびお問い合わせのありました件について、下記のとおり回答いたします。",
]
SECTIONS = [
    ("１　土曜授業の意義と実施の背景について",
     "　土曜日は、区立学校の管理運営に関する規則上、本来「休業日」とされておりますが、"
     "本区では、①必要な授業時数を確保すること、②保護者や地域の皆様に教育活動を広く"
     "公開し、学校教育への理解と信頼を高めること、を意義として、平成24年度から、"
     "教育委員会の関与の下、振替休業日を設定しない土曜授業を全小・中学校で実施して"
     "おります。"),
    ("２　「原則第二土曜日」としている経緯と理由について",
     "　実施日は、平成24年度以降、原則として第二土曜日を基本としてまいりました。"
     "これは、実施日をあらかじめ定まった日とすることで、保護者の皆様が参観のご予定を"
     "立てやすく、また地域や区の行事計画の見通しにもなり、より多くの方にご参加いただける"
     "ためです。なお、教員の働き方改革により、実施回数は令和6年度に年8回から年4回程度へ"
     "見直しており、これらの点は変更が難しい状況にございます。"),
    ("３　令和8年度からの見直しについて",
     "　一方、日程の固定によるご負担は、区としても改善すべき課題と受け止めております。"
     "令和8年度からは、第二土曜日以外でも実施できるよう運用を改めましたので、お子様が"
     "通われる学校の予定をご確認のうえ、ご相談くださいますようお願いいたします。"),
]
CLOSING = "　引き続き、本区の教育活動にご理解とご協力を賜りますよう、お願い申し上げます。"

# ── 文字数チェック（本文：前文＋各節＋結語）────────────
body_chars = "".join(FRONT) + "".join(h + b for h, b in SECTIONS) + CLOSING
count = len(body_chars.replace("　", "").replace(" ", ""))  # 字下げ等の空白を除く
count_with_space = len(body_chars)
print(f"本文文字数（空白除く）：{count}字　／　（字下げ含む）：{count_with_space}字")

# ── 文書生成 ─────────────────────────────────────────
doc = Document()
section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width = Cm(21.0)
section.top_margin = Cm(2.2)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

MINCHO = "ＭＳ 明朝"
GOTHIC = "ＭＳ ゴシック"


def set_font(run, size, bold=False, color=None, name=MINCHO):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def para(text="", size=10.5, bold=False, align=None, name=MINCHO,
         before=0, after=4, line=18, first_indent=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = Pt(line)
    if align is not None:
        p.alignment = align
    if first_indent is not None:
        pf.first_line_indent = Cm(first_indent)
    if text:
        set_font(p.add_run(text), size, bold=bold, name=name)
    return p


# 発信日・宛名・発信者
para("令和８年６月　　日", align=WD_ALIGN_PARAGRAPH.RIGHT, after=10)
para("今井　美咲　様", size=11, after=10)
para("練馬区教育委員会", align=WD_ALIGN_PARAGRAPH.RIGHT, after=0, line=15)
para("教育振興部　教育指導課", align=WD_ALIGN_PARAGRAPH.RIGHT, after=14, line=15)

# 件名
para("学校公開日（土曜授業）の日程について（回答）", size=13, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, name=GOTHIC, after=12)

# 前文
for t in FRONT:
    para(t, after=2)

# 記
para("記", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=8)

# 各節
for h, b in SECTIONS:
    para(h, size=11, bold=True, name=GOTHIC, before=6, after=3)
    para(b, after=4, line=19)

# 結語
para(CLOSING, before=4, after=6)
para("以上", align=WD_ALIGN_PARAGRAPH.RIGHT, after=12)

# 担当
para("【担当】", size=10, name=GOTHIC, after=0, line=15)
para("練馬区教育委員会　教育振興部　教育指導課", size=10, after=0, line=15)
para("電話　５９８４－５７５９", size=10, after=0, line=15)

doc.save("/home/user/con30/学校公開日の日程について（回答・改訂版）.docx")
print("Saved: 学校公開日の日程について（回答・改訂版）.docx")
