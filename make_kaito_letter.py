# -*- coding: utf-8 -*-
"""
今井様あて「学校公開日（土曜授業）の日程について（回答）」改訂版を生成する。

上司からの指示：
  「土曜日の背景・変えられないこと等に触れないと（相手は）納得されない」
を踏まえ、当初案（重要な視点として受け止めるが趣旨をご理解ください、のみ）に
  ① 土曜授業を実施している背景（経緯・目的）
  ② 実施日を個別に変更できない理由（変えられないこと）
  ③ 令和8年度からの日程柔軟化（前向きな対応）
を盛り込んだ回答文書。
出典：添付3資料（PDF＝土曜授業の在り方／PPT＝R8の在り方／当初回答案docx）
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()

# ── ページ設定（A4・標準余白）─────────────────────────
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
         before=0, after=4, line=18, first_indent=None, color=None):
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
        r = p.add_run(text)
        set_font(r, size, bold=bold, color=color, name=name)
    return p


# ── 発信日（右寄せ）──────────────────────────────────
para("令和８年６月　　日", size=10.5, align=WD_ALIGN_PARAGRAPH.RIGHT, after=10)

# ── 宛名（左）────────────────────────────────────────
para("今井　美咲　様", size=11, after=10)

# ── 発信者（右寄せ）─────────────────────────────────
para("練馬区教育委員会", size=10.5, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0, line=15)
para("教育振興部　教育指導課", size=10.5, align=WD_ALIGN_PARAGRAPH.RIGHT, before=0, after=14, line=15)

# ── 件名（中央・ゴシック）──────────────────────────
para("学校公開日（土曜授業）の日程について（回答）", size=13, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, name=GOTHIC, after=12)

# ── 前文 ─────────────────────────────────────────────
para("　日頃より、練馬区の学校教育にご理解とご協力を賜り、厚く御礼申し上げます。", after=2)
para("　このたびお問い合わせをいただきました学校公開日（土曜授業）の日程につきまして、"
     "これまでの経緯及び区の考え方を含め、下記のとおり回答いたします。", after=6)

# ── 記 ───────────────────────────────────────────────
para("記", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)


def heading(num, text):
    para(f"{num}　{text}", size=11, bold=True, name=GOTHIC, before=6, after=3, line=18)


def body(text):
    para(text, size=10.5, after=4, line=19, first_indent=0.36)


# 1 背景
heading("１", "土曜授業（学校公開）を実施している背景について")
body("学校は、平成14年度からの完全学校週５日制の下、学校・家庭・地域が連携し、社会全体で"
     "子供を育てることを基本としております。こうした中、国（文部科学省）は平成25年11月に"
     "学校教育法施行規則を改正し、各設置者の判断により土曜日に授業を実施できることを"
     "明確にしました。また東京都教育委員会からも、保護者や地域に開かれた学校づくりを"
     "進める観点から、土曜日の授業実施に係る留意点が示されております。")
body("本区におきましても、これらの趣旨を踏まえ、平成24年度から、振替休業日を設定しない"
     "土曜授業を全小・中学校で実施しております。これは、①児童・生徒に必要な授業時数を"
     "確保すること、②保護者や地域の皆様に教育活動を広く公開し、学校教育への理解と信頼を"
     "高めること、を主なねらいとするものです。")

# 2 変えられないこと
heading("２", "実施日の設定の考え方と、変更が難しい点について")
body("土曜授業は、上記のとおり「より多くの保護者・地域の皆様に学校を開く」ことを目的と"
     "しているため、特定のご家庭ごとに日程を設定するのではなく、区立学校全体で共通の"
     "考え方の下、計画的に実施日を設定しております。実施日は、地域の行事や区の各種事業を"
     "計画する際の指標にもなっており、この点は制度の趣旨や学校運営上、変更が難しいもので"
     "ございます。")
body("また、教員の働き方改革の観点から、児童・生徒及び教職員の負担軽減を図るため、"
     "実施回数は令和６年度に従来の年８回から年４回程度へ見直したところであり、回数を"
     "増やして対応することも困難な状況にございます。これらの事情について、何卒ご理解を"
     "賜りますようお願い申し上げます。")

# 3 前向きな対応（納得の決め手）
heading("３", "実施日の柔軟化に向けた令和８年度からの見直しについて")
body("一方で、今井様からご指摘いただきました「実施日が固定されていることによるご負担」に"
     "つきましては、各学校からも同様の声が寄せられており、区としても改善すべき課題と"
     "受け止めております。")
body("そこで令和８年度からは、これまで「原則第２土曜日」としておりました実施日を見直し、"
     "第２土曜日以外の土曜日にも実施できるよう運用を改めました。各学校が地域や保護者の"
     "実情に応じて公開日を設定できるようになりますので、お子様が通われる学校の年間行事"
     "予定をご確認いただくとともに、ご都合等につきましては学校へご相談くださいますよう"
     "お願いいたします。")
body("なお、土曜授業のほかにも、各学校では平日等に学校公開や授業参観の機会を設けており"
     "ます。お子様の学習の様子をご覧いただける機会として、あわせてご活用いただければ"
     "幸いです。")

# ── 結語 ─────────────────────────────────────────────
para("　引き続き、練馬区立学校の教育活動にご理解とご協力を賜りますよう、お願い申し上げます。",
     before=4, after=6)
para("以上", align=WD_ALIGN_PARAGRAPH.RIGHT, after=12)

# ── 担当 ─────────────────────────────────────────────
para("【担当】", size=10, name=GOTHIC, after=0, line=15)
para("練馬区教育委員会　教育振興部　教育指導課", size=10, after=0, line=15)
para("電話　５９８４－５７５９", size=10, after=0, line=15)

doc.save("/home/user/con30/学校公開日の日程について（回答・改訂版）.docx")
print("Saved: 学校公開日の日程について（回答・改訂版）.docx")
