# -*- coding: utf-8 -*-
"""令和8年度9月合同副校園長会（教務関係）発表原稿（推敲版）をWord形式で出力する。

pptx のノート欄を原本とし、〔省略可〕で始まる段落は時間調整用として灰色で表示する。
"""
import zipfile
from xml.dom import minidom

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

DECK = "令和8年度9月合同副校園長会_教務関係.pptx"
OUT = "令和8年度9月合同副校園長会_教務関係_発表原稿（推敲版）.docx"
OPTIONAL = "〔省略可〕"
CPM = 300  # 読み上げ速度の目安（字／分）

SLIDE_TITLES = {
    1: "表紙　令和8年度9月合同副校園長会　教務関係について",
    2: "1　実施授業時数について（1学期）",
    3: "2　全国学力・学習状況調査の結果および活用について　－教科に関する調査－",
    4: "2　全国学力・学習状況調査の結果および活用について　－質問紙調査①－",
    5: "2　全国学力・学習状況調査の結果および活用について　－質問紙調査②－",
    6: "3　令和8年度授業改善推進プランの作成および授業改善推進プランに基づく教育活動の実施について",
    7: "4　その他　－東京都中学校英語スピーキングテスト（ESAT-J）－",
}
GRAY = (0x7F, 0x7F, 0x7F)
NAVY = (0x1F, 0x3B, 0x63)


def read_notes(deck):
    z = zipfile.ZipFile(deck)
    out = {}
    for i in range(1, 8):
        d = minidom.parseString(z.read(f"ppt/notesSlides/notesSlide{i}.xml"))
        paras = []
        for p in d.getElementsByTagName("a:p"):
            s = "".join(t.firstChild.nodeValue if t.firstChild else "" for t in p.getElementsByTagName("a:t"))
            if s.strip() and s.strip() != str(i):
                paras.append(s)
        out[i] = paras
    return out


def set_font(run, size, bold=False, color=None, name="MS Gothic"):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def add_para(doc, text, size=11, bold=False, color=None, before=0, after=6,
             line=None, indent=None, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    if line:
        pf.line_spacing = Pt(line)
    if indent:
        pf.left_indent = Cm(indent)
    if text:
        set_font(p.add_run(text), size, bold, color)
    return p


def minutes(chars):
    return chars / CPM


notes = read_notes(DECK)
full = {i: sum(len(p) for p in ps) for i, ps in notes.items()}
opt = {i: sum(len(p) for p in ps if p.startswith(OPTIONAL)) for i, ps in notes.items()}
full_all, base_all = sum(full.values()), sum(full[i] - opt[i] for i in full)

doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
sec.top_margin = sec.bottom_margin = Cm(1.8)
sec.left_margin = sec.right_margin = Cm(2.0)

add_para(doc, "令和8年度9月　合同副校園長会", 12, bold=True, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
add_para(doc, "教務関係　発表原稿", 18, bold=True, after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
add_para(doc, "令和8年9月10日（木）　教育指導課　指導主事", 10.5, after=10,
         align=WD_ALIGN_PARAGRAPH.CENTER)

add_para(doc, "■　時間の目安（持ち時間15分）", 10.5, bold=True, color=NAVY, after=4)
rows = [("スライド", "標準", "全文")]
for i in range(1, 8):
    rows.append((f"{i}", f"{minutes(full[i] - opt[i]):.1f}分", f"{minutes(full[i]):.1f}分"))
rows.append(("合計", f"約{minutes(base_all):.0f}分", f"約{minutes(full_all):.0f}分"))

table = doc.add_table(rows=len(rows), cols=3)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.LEFT
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = table.cell(r, c)
        cell.width = Cm(3.0)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
        set_font(p.add_run(val), 9, bold=(r == 0 or r == len(rows) - 1))

add_para(doc, f"「標準」は灰色の{OPTIONAL}段落を読まない場合、「全文」はすべて読む場合。"
              f"読み上げ速度は{CPM}字／分で換算。数表を指し示す間を含めても15分に収まる。",
         9, color=GRAY, before=4, after=14)

for i in range(1, 8):
    add_para(doc, f"【スライド{i}】{SLIDE_TITLES[i]}", 11, bold=True, color=NAVY,
             before=12 if i > 1 else 0, after=6)
    for para in notes[i]:
        is_opt = para.startswith(OPTIONAL)
        add_para(doc, para, 11, color=GRAY if is_opt else None, after=8, line=20, indent=0.4)

doc.save(OUT)
print("wrote", OUT, f"| 標準 約{minutes(base_all):.1f}分 / 全文 約{minutes(full_all):.1f}分")
