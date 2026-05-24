"""150周年記念誌 教育長前文 検討資料 をWord文書として生成"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ── ページ設定（A4、標準余白） ──────────────────────────
section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width = Cm(21.0)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)

JP_FONT_MINCHO = "MS 明朝"
JP_FONT_GOTHIC = "MS ゴシック"


def set_font(run, size=10.5, bold=False, color=None, name=JP_FONT_MINCHO):
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


def para_space(p, before=0, after=4, line=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = Pt(line)


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_heading(text, level=1):
    sizes = {1: 16, 2: 13, 3: 11.5}
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, size=sizes.get(level, 11), bold=True, name=JP_FONT_GOTHIC)
    para_space(p, before=12 if level == 1 else 8, after=6)
    return p


def add_para(text, bold=False, italic=False, indent=0, align=None,
             size=10.5, color=None, name=JP_FONT_MINCHO):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if align:
        p.alignment = align
    run = p.add_run(text)
    set_font(run, size=size, bold=bold, color=color, name=name)
    run.italic = italic
    para_space(p, after=4, line=18)
    return p


def add_bullet(text, level=0):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5 + level * 0.5)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    run = p.add_run("・" + text)
    set_font(run, size=10.5)
    para_space(p, after=2, line=16)
    return p


def add_quote_para(text, bold=False, size=10.5):
    """前文本文用：左罫線つき引用風"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.right_indent = Cm(0.6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:color"), "8B4513")
    pBdr.append(left)
    pPr.append(pBdr)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold)
    para_space(p, after=4, line=20)
    return p


def add_divider():
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:color"), "888888")
    pBdr.append(bottom)
    pPr.append(pBdr)
    para_space(p, after=6)


def add_table(headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.autofit = False
    if col_widths:
        for i, w in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)
    # ヘッダー
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_bg(cell, "D9E2F3")
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h)
        set_font(run, size=10, bold=True, name=JP_FONT_GOTHIC)
        para_space(p, after=0)
    # 行
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(val)
            set_font(run, size=9.5)
            para_space(p, after=0, line=14)
    # 段落間スペース
    doc.add_paragraph()


# ============================================================
# 表紙・タイトル
# ============================================================
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("創立150周年記念誌　教育長前文（検討資料）")
set_font(run, size=18, bold=True, name=JP_FONT_GOTHIC)
para_space(title, before=12, after=10)

info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = info.add_run("作成日：2026年5月\n執筆者（想定）：練馬区教育委員会 教育長")
set_font(run, size=10.5)
para_space(info, after=8)

add_para(
    "目的：140周年記念誌の教育長前文と内容が重複しているとの指摘を受け、"
    "120年・130年・140年の流れを踏まえた150周年に相応しい教育長前文を"
    "作成するための検討資料。"
)
add_divider()

# ============================================================
# 1. 過去3回の前文構造分析
# ============================================================
add_heading("1. 過去3回の教育長前文の構造分析（再構成）", level=1)
add_para(
    "※ 過去の現物が手元にないため、当時の社会情勢・教育施策・"
    "典型的な周年挨拶の作法から再構成した推定構造。"
    "実物と照合のうえ調整してください。",
    italic=True, size=9.5, color=(100, 100, 100)
)

add_heading("1-1. 120周年（平成8年／1996年頃）の推定構造", level=2)
for t in [
    "高度経済成長・バブル後の安定期、「ゆとり教育」議論の本格化",
    "国際化・情報化の到来（パソコンが学校に入り始めた時期）",
    "「生きる力」の答申（中教審・平成8年）への期待",
    "完全週5日制への移行論議",
    "練馬区としての地域教育への期待",
]:
    add_bullet(t)
add_para(
    "典型文型：「明治の創立以来、幾多の困難を乗り越え……"
    "国際化・情報化の時代を生き抜く子どもたちのために」",
    indent=0.5, italic=True, size=10
)

add_heading("1-2. 130周年（平成18年／2006年頃）の推定構造", level=2)
for t in [
    "教育基本法改正（平成18年12月）の年",
    "「ゆとり教育」見直し、学力低下論争",
    "構造改革・三位一体改革の中での教育",
    "学校選択制・学校評議員制度の定着",
    "食育基本法、特別支援教育への転換",
]:
    add_bullet(t)
add_para(
    "典型文型：「130年の伝統を礎に、新しい時代の教育基本法の理念のもとで、"
    "確かな学力と豊かな心を……」",
    indent=0.5, italic=True, size=10
)

add_heading("1-3. 140周年（平成28年／2016年頃）の推定構造", level=2)
for t in [
    "東日本大震災（平成23年）からの復興と防災教育の再構築",
    "新学習指導要領（平成29年告示）への準備",
    "「主体的・対話的で深い学び」（アクティブ・ラーニング）",
    "道徳の教科化、英語の早期化、プログラミング教育の導入準備",
    "学校・家庭・地域の三者連携、コミュニティ・スクール推進",
    "練馬区基本構想・教育振興基本計画の更新時期",
]:
    add_bullet(t)
add_para(
    "典型文型：「140年の歴史と伝統を礎に、変化の激しいこれからの社会を"
    "生き抜く力を……家庭・地域と連携し……」",
    indent=0.5, italic=True, size=10
)

add_heading("1-4. 共通して使われがちな常套句（150周年では避ける）", level=2)
add_table(
    headers=["表現", "出現可能性", "150周年での扱い"],
    rows=[
        ["明治○年の創立以来、幾多の困難を乗り越え", "120/130/140 全て",
         "創立年だけを事実として記載、過剰な英雄譚は避ける"],
        ["変化の激しい時代を生き抜く力", "130/140",
         "「生き抜く」は退却の発想。150年は「共に創る」に転換"],
        ["家庭・地域・学校の三者が連携し", "130/140",
         "コミュニティ・スクールとして制度化された前提で書く"],
        ["未来を担う子どもたち", "120/130/140 全て",
         "「未来の存在」でなく「いまここの主体」として描く"],
        ["歴史と伝統を礎に", "120/130/140 全て",
         "「歴史を引き受けながら更新する」と能動表現に"],
        ["心からお祝い申し上げます", "120/130/140 全て",
         "残してよい。冒頭固定でなく構成上の位置を工夫"],
    ],
    col_widths=[5.5, 3.5, 7.5],
)
add_divider()

# ============================================================
# 2. 新しい要素
# ============================================================
add_heading("2. 150周年で初めて教育長が書くべき新しい要素", level=1)
add_para(
    "140周年以降の10年間に、練馬区教育委員会として明確に経験・対応してきた事項。"
    "これらを盛り込まないと「140周年と同じ」と再度言われる。"
)

add_heading("2-1. 10年間の固有経験", level=2)
for t in [
    "COVID-19パンデミック（令和2年〜）：臨時休業、分散登校、行事中止と再開、子どもの心のケア",
    "GIGAスクール構想の本格運用：一人一台端末、家庭への持ち帰り、情報モラル教育",
    "生成AIの教育利用議論（令和5年以降）",
    "校則・きまりの見直しと公開：練馬区として各校のきまりを集約・点検する動き",
    "学校事故・重大事態への詳細調査体制の整備",
    "不登校児童生徒の増加と教育機会確保法への対応",
    "コミュニティ・スクール（学校運営協議会）の本格導入",
    "練馬区基本構想2040、教育・子育てに関する新計画",
]:
    add_bullet(t)

add_heading("2-2. 過去の前文では言いにくかったが今は語るべきこと", level=2)
for t in [
    "子どもの権利・ウェルビーイングを学校運営の中心に据える方針転換",
    "多様な背景（外国にルーツのある児童、特別支援、性的マイノリティ等）への配慮",
    "教職員の働き方改革",
    "学校が「行政が運営する施設」から「地域とともに営む公共空間」へ移行している実感",
    "過去の重大事案を踏まえた、安全・人権への一層の責任",
]:
    add_bullet(t)
add_divider()

# ============================================================
# 3. 5つの視点
# ============================================================
add_heading("3. 5つの読み手を意識した教育長前文の設計", level=1)
add_para(
    "教育長は「全ての関係者に向けて語る」立場。5つの視点それぞれが、"
    "自分に向けられた言葉として読めるよう、前文の各段落で配慮ポイントを散りばめる。"
)

stakeholders = [
    ("3-1. 地域（地域住民・町会・卒業生）",
     "学校が地域の歴史そのものであることへの敬意。"
     "コミュニティ・スクールが制度として根付き、地域が学校運営の正規メンバーとなったこと。",
     "「地域の絆」「ふるさと」を懐古的に並べるだけの表現"),
    ("3-2. 教育委員会（同僚・事務局・他校校長への波及）",
     "区教委として150年の歴史に責任を持つ姿勢。"
     "歴代の教育関係者への敬意。今後の練馬区教育の方向性の中での位置づけ。",
     "行政の自賛にしか聞こえない施策の羅列"),
    ("3-3. 校長・教職員",
     "現場の教職員が直面している困難（多忙、複雑化する保護者対応、子どもの多様化）への"
     "理解と敬意。安全管理・人権・働き方への組織的支援の約束。",
     "「先生方の一層のご尽力を」と上から目線で締める形"),
    ("3-4. 地域の議員（区議会議員）",
     "議会の理解と監視のもとで教育行政が成り立っていることへの謝意。"
     "学校が地域の公共財であるとの認識の共有。",
     "議員名や会派名への個別言及、政治色"),
    ("3-5. 保護者（PTA・現役保護者）",
     "保護者を「協力者」ではなく「共同当事者」として位置づける方針。"
     "コロナ禍で行事に参加できなかった世代への配慮。",
     "「PTA活動へのご協力に感謝」だけで終わる定型"),
]
for title_, msg, avoid in stakeholders:
    add_heading(title_, level=2)
    p = doc.add_paragraph()
    run = p.add_run("【教育長として伝えるべきこと】")
    set_font(run, size=10.5, bold=True, name=JP_FONT_GOTHIC)
    run2 = p.add_run("　" + msg)
    set_font(run2, size=10.5)
    para_space(p, after=3, line=18)

    p = doc.add_paragraph()
    run = p.add_run("【避ける表現】")
    set_font(run, size=10.5, bold=True, color=(180, 30, 30), name=JP_FONT_GOTHIC)
    run2 = p.add_run("　" + avoid)
    set_font(run2, size=10.5)
    para_space(p, after=6, line=18)
add_divider()

# ============================================================
# 4. 統合稿
# ============================================================
add_heading("4. 教育長前文（150周年案）— 統合稿", level=1)

# 見出し
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("創立150周年に寄せて")
set_font(run, size=15, bold=True, name=JP_FONT_GOTHIC)
para_space(p, before=8, after=6)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run("練馬区教育委員会 教育長　〇〇　〇〇")
set_font(run, size=11)
para_space(p, after=10)

body_paragraphs = [
    "〇〇小学校（中学校）が、創立150周年の佳節を迎えられましたことを、"
    "練馬区教育委員会を代表して心よりお祝い申し上げます。",

    "明治の創立から今日まで、本校は四つの時代を生き、地域の子どもたちとともに"
    "歩んでまいりました。120周年の年、私たちは情報化と国際化の入り口に立っていました。"
    "130周年の年、教育基本法が改正され、学校・家庭・地域が手を携える教育の在り方が"
    "改めて問われました。140周年の年、新しい学習指導要領のもとで「主体的・対話的で"
    "深い学び」を実現しようと、本校もまた挑戦を続けていたと記憶しています。",

    "そして、140周年から今日に至る十年は、おそらく本校の歴史の中でも、"
    "最も急峻な十年でした。世界的な感染症は、子どもたちから当たり前の登校や行事を"
    "一時奪い、私たちに学校という場の意味を根本から問い直させました。"
    "GIGAスクール構想によって、一人一台の端末が教室に届き、学びの姿は大きく変わりました。"
    "校則やきまりの在り方が地域全体で議論され、子ども自身の声をどう聞くかという問いが、"
    "教育委員会の中心的な課題となりました。生成AIという、これまで想像もしなかった技術とも、"
    "子どもたちはすでに向き合っています。練馬区としても、こうした変化のなかで、"
    "子どもの権利と安全、そして一人ひとりのウェルビーイングを学校運営の中心に据えるべく、"
    "施策を組み立て直してまいりました。",

    "この十年、私たちは時に立ち止まり、時に反省を強いられる出来事にも向き合ってまいりました。"
    "学校で起きてしまった出来事を真摯に検証し、二度と同じ悲しみを生まないために、"
    "関係者の皆様と知恵を寄せ合った時間も、本校の150年の一部です。"
    "歴史とは、輝かしい記念だけでなく、こうした重みを引き受けていくことだと、"
    "改めて感じています。",

    "いま、学校は、行政だけのものでも、教職員だけのものでも、"
    "保護者だけのものでもありません。地域の皆様には、コミュニティ・スクールの仕組みを"
    "通じて、これまで以上に学校運営の中心にお立ちいただいています。"
    "保護者の皆様には、子育ての協力者というよりも、子どもの育ちをともに担う"
    "共同当事者として、対話を重ねていただいています。"
    "区議会の皆様には、教育予算と政策に温かくも厳しい眼差しを向け、"
    "150周年の節目を支えていただきました。"
    "校長先生をはじめとする教職員の皆様の、子どもに寄り添う日々の実践こそが、"
    "本校の最大の財産です。すべての関係者の皆様に、深く感謝申し上げます。",

    "150年は通過点です。160年、200年と続いていく本校の歴史のなかで、"
    "いま教室にいる子どもたち一人ひとりが、「自分はここにいてよかった」と"
    "思える学校を、私たちは作り続けなければなりません。"
    "練馬区教育委員会は、その営みを全力で支えてまいります。",

    "結びに、本校の益々の発展と、子どもたち、教職員、保護者、地域の皆様の"
    "御健勝を祈念し、お祝いの言葉といたします。",
]

for para_text in body_paragraphs:
    add_quote_para("　" + para_text)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run("令和8年〇月〇日")
set_font(run, size=11)
para_space(p, before=8, after=2)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run("練馬区教育委員会 教育長　〇〇　〇〇")
set_font(run, size=11)
para_space(p, after=10)

add_divider()

# ============================================================
# 5. 差別化チェックリスト
# ============================================================
add_heading("5. 140周年前文との差別化チェックリスト", level=1)
add_para("採用前に以下を確認してください（特に教育長前文として）：")
for t in [
    "「歴史と伝統を礎に」「未来を担う子ども」「変化の激しい時代を生き抜く」など、"
    "過去3回（120/130/140）で必ず使われた常套句が支配的になっていないか",
    "過去の周年（120・130・140）に明示的に言及し、歴史の連続性と差異を示しているか",
    "この10年固有の事象（コロナ、GIGA、校則見直し、重大事態、生成AI）に"
    "具体的に触れているか",
    "教育長として、行政の自賛でなく、現場と地域への敬意が前面に出ているか",
    "5つの読み手（地域・教委・校長・議員・保護者）が、それぞれ"
    "「自分に向けて書かれている」と読める箇所があるか",
    "重大事案や困難に目を背けず、しかし具体名は出さず、品位を保って言及できているか",
    "結語が「益々のご発展」だけでなく、子ども一人ひとりへの視座を含んでいるか",
    "140周年の前文を並べて読んだとき、明らかに「別の時代の別の教育長」が書いた"
    "文書として読めるか",
]:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    run = p.add_run("□ " + t)
    set_font(run, size=10.5)
    para_space(p, after=3, line=18)
add_divider()

# ============================================================
# 6. 残された論点
# ============================================================
add_heading("6. 残された論点（実行委員会・教委事務局との協議事項）", level=1)
for i, t in enumerate([
    "過去3回（120・130・140）の教育長前文の現物確認と、重複箇所の具体的な指摘",
    "重大事態・学校事故への言及の深度（一般論に留めるか、具体に踏み込むか）",
    "教育長個人の所感をどこまで入れるか（公式の祝辞としての品格と個人性のバランス）",
    "練馬区基本構想・教育振興基本計画への明示的参照の有無",
    "校長・PTA会長・実行委員長等、他の挨拶文との役割分担（重複・トーンの調整）",
    "多言語版（やさしい日本語、英語、当該校に多い言語）の併設可否",
    "児童・生徒代表の言葉と教育長前文の配置順序",
], start=1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.7)
    p.paragraph_format.first_line_indent = Cm(-0.7)
    run = p.add_run(f"{i}．{t}")
    set_font(run, size=10.5)
    para_space(p, after=3, line=18)
add_divider()

# ============================================================
# 7. 次のステップ
# ============================================================
add_heading("7. 次のステップ", level=1)
for i, t in enumerate([
    "過去3回の前文現物の入手 ― 教育委員会事務局または学校保管文書から"
    "120/130/140周年記念誌を取り寄せ、本資料の推定と照合",
    "重複箇所の特定 ― 140周年前文と本案を逐文で対比し、表現の重複・趣旨の重複を点検",
    "教育長レビュー ― 教育長ご本人のお考えを反映した修正",
    "実行委員会回覧 ― 校長・PTA会長・地域代表・議員代表からの確認",
    "最終稿の確定と記念誌入稿",
], start=1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.7)
    p.paragraph_format.first_line_indent = Cm(-0.7)
    run = p.add_run(f"{i}．{t}")
    set_font(run, size=10.5)
    para_space(p, after=3, line=18)

# ── 保存 ─────────────────────────────────────────────
output_path = "/home/user/con30/150周年記念前文_検討資料.docx"
doc.save(output_path)
print(f"saved: {output_path}")
