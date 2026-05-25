"""豊玉小学校 創立150周年記念誌 教育長前文（Z3・Z4 2案比較版）"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

JP_MINCHO = "MS 明朝"
JP_GOTHIC = "MS ゴシック"


def set_font(run, size=10.5, bold=False, name=JP_MINCHO, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)


def para(text, size=10.5, bold=False, align=None, name=JP_MINCHO,
         line_spacing=22, before=0, after=6, color=None):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = Pt(line_spacing)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold, name=name, color=color)
    return p


def section_label(label, total_chars):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:color"), "1F4E79")
    pBdr.append(bottom)
    pPr.append(pBdr)
    pf = p.paragraph_format
    pf.space_before = Pt(18)
    pf.space_after = Pt(8)
    run = p.add_run(label)
    set_font(run, size=13, bold=True, name=JP_GOTHIC, color=(0x1F, 0x4E, 0x79))
    run2 = p.add_run(f"　［全体 {total_chars} 文字］")
    set_font(run2, size=10, name=JP_GOTHIC, color=(0x60, 0x60, 0x60))


def build_full_letter(history_paragraph):
    # ①祝意
    para(
        "　練馬区立豊玉小学校が、開校百五十周年を迎えられましたことを心より"
        "お祝い申し上げます。",
        size=11
    )
    # ②歴史（差し替え対象）
    para("　" + history_paragraph, size=11)
    # ③実績
    para(
        "　この間、文部科学省、東京都および練馬区の研究指定を度々受けながら"
        "教育目標「考える子　ねばり強い子　心ゆたかな子」の育成に努め、"
        "平成二十九年度には全国学校体育研究優良校として表彰されました。"
        "令和の新時代となってからは小学校教科担任制を推進され、中学校への"
        "円滑な接続の面からも成果を挙げています。四十分午前五時間授業の実施や"
        "探究的な学びにも力を入れ、令和七・八年度には練馬区教育委員会教育課題"
        "研究指定校として、今年度からは文部科学省教育課程サキドリ研究指定校"
        "として実践を深めています。",
        size=11
    )
    # ④感謝
    para(
        "　こうした本校の発展は、歴代校長先生をはじめ教職員の皆様、学校を"
        "支えてくださった地域や保護者の皆様のご尽力の賜と、深く感謝申し上げます。",
        size=11
    )
    # ⑤結語
    para(
        "　豊玉小学校が歴史と伝統を次代に引き継ぐとともに、新たな時代の創造に"
        "向けて益々のご発展を遂げられますことを祈念し、お祝いのことばといたします。",
        size=11
    )


doc = Document()
section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width = Cm(21.0)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)

# 表紙
para("豊玉小学校 創立150周年記念誌 教育長前文",
     size=17, bold=True, name=JP_GOTHIC,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=4)
para("【Z3・Z4 ２案比較】",
     size=13, bold=True, name=JP_GOTHIC,
     align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=10,
     color=(0x1F, 0x4E, 0x79))
para("作成日：2026年5月　／　書き手：練馬区教育委員会 教育長",
     size=10, align=WD_ALIGN_PARAGRAPH.RIGHT, after=6,
     color=(0x60, 0x60, 0x60))
para(
    "第2段落（学校の歴史紹介・創立経緯）について、骨組みから組み替えた"
    "2案を併記しました。どちらも史実（明治九年・南蔵院本堂・教員二名・"
    "児童八十名・校名「豊玉」由来・卒業生一万二千六百六十八名）は完全保持。",
    size=10, color=(0x40, 0x40, 0x40), after=8
)

# ==========================
# 比較表
# ==========================
section_label("◆ 2案の比較", total_chars=0)

table = doc.add_table(rows=4, cols=3)
table.style = "Table Grid"
headers = ["観点", "Z3：校名を入口に", "Z4：連続性を主題に"]
rows = [
    ["1文目で語る主題", "校名「豊玉」の意味", "百五十年歩みを止めなかった理由"],
    ["全体構造", "校名 → 創立 → 継承", "連続性の表明 → 創立＋校名 → 同窓"],
    ["140周年版との差異", "校名を冒頭に立てる時点で構造が異なる", "「なぜ続いてきたか」を主題に据え、史実を説明として位置づけ"],
]
for i, h in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = ""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "D9E2F3")
    tcPr.append(shd)
    run = cell.paragraphs[0].add_run(h)
    set_font(run, size=10.5, bold=True, name=JP_GOTHIC)
for r_idx, row in enumerate(rows, start=1):
    for c_idx, val in enumerate(row):
        cell = table.rows[r_idx].cells[c_idx]
        cell.text = ""
        run = cell.paragraphs[0].add_run(val)
        set_font(run, size=10)
para("", after=2)

# ==========================
# 【Z3】全文
# ==========================
section_label("◆ 【Z3案】校名を入口に　── 全体 579 文字", total_chars=579)

build_full_letter(
    "「豊玉」── この校名は、明治九年、地域の方々が南蔵院本堂の一隅に集った"
    "教員二名と児童八十名に贈られた祈りそのものでございます。"
    "稲穂のように、子どもの心も豊かに実れ。"
    "この祈りは百五十年絶えることなく、一万二千六百六十八名の同窓のもとで"
    "確かな実りを結んでまいりました。"
)

para("令和8年〇月〇日　練馬区教育委員会　教育長　三浦　康彰",
     size=10.5, align=WD_ALIGN_PARAGRAPH.RIGHT, before=4, after=4)

# ==========================
# 【Z4】全文
# ==========================
section_label("◆ 【Z4案】連続性を主題に　── 全体 590 文字", total_chars=590)

build_full_letter(
    "本校が百五十年の永きにわたり歩みを止めなかったのは、明治九年に地域の方々が"
    "学び舎に贈られた「豊玉」── 稲穂のように子どもの心も豊かに実れ ── との祈りが、"
    "南蔵院本堂で迎えられた教員二名・児童八十名から、一万二千六百六十八名の同窓を"
    "介して、今日の教室まで脈々と継がれてきたからでございます。"
)

para("令和8年〇月〇日　練馬区教育委員会　教育長　三浦　康彰",
     size=10.5, align=WD_ALIGN_PARAGRAPH.RIGHT, before=4, after=4)

# ==========================
# 補足
# ==========================
section_label("◆ それぞれの「読み味」", total_chars=0)
para(
    "【Z3】校名を入口に：",
    size=10.5, bold=True, name=JP_GOTHIC, after=2
)
para(
    "　「豊玉」という二文字から入り、その意味を解読してから史実を語る順序。"
    "教育長として「この学校名そのものが、地域の祈りなのです」と冒頭で言い切る"
    "強さがあり、聞き手の心に校名の意味を最初に刻み込む構成。"
    "短文「稲穂のように、子どもの心も豊かに実れ。」が余韻を残す。",
    size=10.5, after=8
)
para(
    "【Z4】連続性を主題に：",
    size=10.5, bold=True, name=JP_GOTHIC, after=2
)
para(
    "　「百五十年歩みを止めなかった」という一点に焦点を絞り、すべての史実を"
    "「なぜ続いてきたか」の答えとして畳みかける一文構成。"
    "教育長の視点で「この学校が続いてきたこと自体の意味」を語る重厚感があり、"
    "150周年という節目の重みと最も響き合う。",
    size=10.5, after=4
)

output_path = "/home/user/con30/R8_豊玉小学校150周年_教育長前文_2案比較.docx"
doc.save(output_path)
print(f"saved: {output_path}")
