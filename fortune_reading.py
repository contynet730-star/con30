from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width  = Cm(21.0)
section.top_margin    = Cm(1.5)
section.bottom_margin = Cm(1.5)
section.left_margin   = Cm(2.0)
section.right_margin  = Cm(2.0)

def set_font(run, size, bold=False, color=None, name="MS Gothic"):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)

def para_space(p, before=0, after=0, line=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)
    if line:
        pf.line_spacing = Pt(line)

def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

# ══ タイトル帯 ══════════════════════════════════════════════
tbl = doc.add_table(rows=1, cols=1)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.style = "Table Grid"
cell = tbl.rows[0].cells[0]
set_cell_bg(cell, "1A1A2E")
p = cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
para_space(p, before=4, after=2, line=16)
r = p.add_run("六星占術　運命鑑定書")
set_font(r, 18, bold=True, color=(212, 175, 55), name="MS Mincho")
p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
para_space(p2, before=0, after=4, line=12)
r2 = p2.add_run("細木数子式　令和８年（2026年）７月以降の運勢")
set_font(r2, 10, color=(200, 200, 200))

# ══ 基本情報 ═══════════════════════════════════════════════
p = doc.add_paragraph()
para_space(p, before=6, after=2, line=13)
r = p.add_run("■ 鑑定対象者データ")
set_font(r, 10, bold=True, color=(100, 50, 0))

tbl2 = doc.add_table(rows=4, cols=2)
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl2.style = "Table Grid"
data = [
    ("生年月日", "昭和56年（1981年）7月30日生まれ"),
    ("血液型",   "A型"),
    ("出身地",   "石川県白山市"),
    ("現住所",   "東京都"),
]
label_color = "E8E0F0"
for i, (label, value) in enumerate(data):
    cl = tbl2.rows[i].cells[0]
    cr = tbl2.rows[i].cells[1]
    set_cell_bg(cl, label_color)
    p_l = cl.paragraphs[0]
    para_space(p_l, before=2, after=2, line=12)
    r_l = p_l.add_run(label)
    set_font(r_l, 9, bold=True, color=(60, 0, 100))
    p_r = cr.paragraphs[0]
    para_space(p_r, before=2, after=2, line=12)
    r_r = p_r.add_run(value)
    set_font(r_r, 9)

# ══ 命星の判定 ══════════════════════════════════════════════
p = doc.add_paragraph()
para_space(p, before=8, after=2, line=13)
r = p.add_run("■ あなたの命星（六星占術）")
set_font(r, 10, bold=True, color=(100, 50, 0))

tbl3 = doc.add_table(rows=1, cols=1)
tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl3.style = "Table Grid"
cb = tbl3.rows[0].cells[0]
set_cell_bg(cb, "FDF6E3")
p = cb.paragraphs[0]
para_space(p, before=4, after=1, line=14)
r = p.add_run("★　金星人（マイナス）　★")
set_font(r, 14, bold=True, color=(180, 130, 0), name="MS Mincho")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

p2 = cb.add_paragraph()
para_space(p2, before=2, after=4, line=13)
r2 = p2.add_run(
    "　昭和56年7月30日生まれのあなたは、六星占術において「金星人マイナス」に属します。\n"
    "　金星人は美的感覚と社交性に恵まれ、人を引きつける魅力を持つ星。\n"
    "　マイナスの気質は、慎重さと深い思慮を兼ね備え、物事の本質を見抜く力に優れています。\n"
    "　A型の血液型と相まって、几帳面で責任感が強く、周囲から信頼を集める運命にあります。\n"
    "　また、白山市という霊山・白山を擁する土地の出身は、先祖の加護が強い証拠。\n"
    "　この地脈のエネルギーが、東京という大都市でのあなたを見えないところで支えています。"
)
set_font(r2, 9, name="MS Mincho")

# ══ 2026年後半の年運 ════════════════════════════════════════
p = doc.add_paragraph()
para_space(p, before=8, after=2, line=13)
r = p.add_run("■ 令和８年（2026年）後半　金星人マイナスの年運")
set_font(r, 10, bold=True, color=(100, 50, 0))

tbl4 = doc.add_table(rows=1, cols=1)
tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl4.style = "Table Grid"
cb2 = tbl4.rows[0].cells[0]
set_cell_bg(cb2, "FFF0F0")
p = cb2.paragraphs[0]
para_space(p, before=3, after=1, line=13)
r = p.add_run("【 乱 気 】― 試練と変容の年")
set_font(r, 12, bold=True, color=(160, 0, 0), name="MS Mincho")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

p2 = cb2.add_paragraph()
para_space(p2, before=3, after=4, line=14)
r2 = p2.add_run(
    "　いいですか、あなたは今年「乱気」の年に入っています。\n"
    "　乱気というのはね、嵐の前の空模様みたいなものなの。\n"
    "　一見、何でもうまくいくように思えて、気が大きくなりやすい。\n"
    "　でもそこが落とし穴。この時期に大きな決断や新しいことを始めると、\n"
    "　後になって「あのときやらなければよかった」となることが多いのよ。\n\n"
    "　とくに7月以降は、乱気の影響が色濃く出てきます。\n"
    "　心が揺らぎやすく、判断力が鈍る。感情に流されると痛い目を見る。\n"
    "　これはあなたの責任じゃなくて、星の流れがそうなっているの。\n"
    "　だから余計に、意識して慎重に動くことが必要なの。"
)
set_font(r2, 9.5, name="MS Mincho")

# ══ 7月以降　分野別の注意点 ═══════════════════════════════
p = doc.add_paragraph()
para_space(p, before=8, after=2, line=13)
r = p.add_run("■ ７月以降に気をつけるべきこと（分野別）")
set_font(r, 10, bold=True, color=(100, 50, 0))

sections_data = [
    (
        "【仕事・キャリア】",
        "1B4332",
        "EAFAF1",
        "0,100,0",
        (
            "　仕事面では、今いる場所をしっかり守ることを第一に考えなさい。\n"
            "　転職、独立、新しいプロジェクトへの参加――こういった動きは乱気の年は厳禁。\n"
            "　「チャンスが来た」と思っても、今年の後半は保留にしておきなさい。\n"
            "　来年、再来年に同じチャンスが来たとき、そのときこそ動きなさい。\n"
            "　今は種を蒔く時期。目立たずコツコツと実績を積み上げることが大切です。\n"
            "　A型の几帳面さを活かして、足元の仕事を丁寧にこなしていけば、\n"
            "　必ず報われる時期が来ます。焦りは禁物よ。"
        )
    ),
    (
        "【人間関係・恋愛・結婚】",
        "1A3A5C",
        "EAF2FB",
        "0,60,130",
        (
            "　金星人は本来、人を引きつける魅力があります。でも乱気の年はその魅力が\n"
            "　逆方向に働くことがある。つまり、あなたを利用しようとする人間が近づいてくる。\n"
            "　新しく出会った人、急速に親しくなろうとする人には要注意。\n"
            "　「なんとなく嫌だな」という直感を大切にしなさい。\n"
            "　恋愛に関しては、今ある関係を大切に。新しい出会いを追いかける時期ではない。\n"
            "　結婚を考えているなら、来年以降に先送りすることを強くすすめます。\n"
            "　東京での人間関係は広がりやすいけれど、深さが大事。表面的なつながりより\n"
            "　本当に信頼できる人を一人でも持つことの方がよほど価値があります。"
        )
    ),
    (
        "【健康・体】",
        "5C1A1A",
        "FDF2F2",
        "130,0,0",
        (
            "　乱気の年は、身体に無理が来やすい。特に注意すべきは「気の張りすぎ」です。\n"
            "　A型は本来、ストレスをため込みやすい血液型。それが乱気の年と重なると、\n"
            "　免疫力が落ちたり、慢性的な疲労が出てきたりします。\n"
            "　７月以降は特に、睡眠を削るような生活習慣を改めなさい。\n"
            "　「ちょっとくらい大丈夫」が積み重なって大事になるのが乱気の怖さ。\n"
            "　消化器系、特に胃腸には気をつけること。食事の時間を乱さないように。\n"
            "　白山の水のような、清らかで規則正しいリズムで生活することが大切です。"
        )
    ),
    (
        "【お金・財産】",
        "3D2B1E",
        "FDF5E6",
        "100,60,0",
        (
            "　乱気の年は金銭的な判断が狂いやすい。「今だ！」と思ったときほど、\n"
            "　立ち止まって三日間考えなさい。衝動的な投資、多額の買い物、\n"
            "　人への貸し付け――これらは乱気の年には必ず後悔します。\n"
            "　株や投資に手を出している人は、ポジションを減らすことを考えなさい。\n"
            "　今年後半は守りの財布を意識して。蓄えることが来年以降の飛躍につながります。\n"
            "　人に頼まれてお金を貸すのも今年は断る勇気を持ちなさい。"
        )
    ),
]

for title, header_color, bg_color, text_color_str, content in sections_data:
    tc_rgb = tuple(int(x) for x in text_color_str.split(","))
    tbl_s = doc.add_table(rows=1, cols=1)
    tbl_s.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_s.style = "Table Grid"
    cb_s = tbl_s.rows[0].cells[0]
    p_title = cb_s.paragraphs[0]
    set_cell_bg(cb_s, bg_color)
    para_space(p_title, before=3, after=1, line=13)
    rt = p_title.add_run(title)
    set_font(rt, 10, bold=True, color=tc_rgb, name="MS Mincho")

    p_body = cb_s.add_paragraph()
    para_space(p_body, before=1, after=4, line=14)
    rb = p_body.add_run(content)
    set_font(rb, 9, name="MS Mincho")

    p_sep = doc.add_paragraph()
    para_space(p_sep, before=2, after=0)

# ══ 先祖供養のすすめ ════════════════════════════════════════
p = doc.add_paragraph()
para_space(p, before=4, after=2, line=13)
r = p.add_run("■ 最も大切なこと ―― 先祖供養について")
set_font(r, 10, bold=True, color=(100, 50, 0))

tbl5 = doc.add_table(rows=1, cols=1)
tbl5.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl5.style = "Table Grid"
cb5 = tbl5.rows[0].cells[0]
set_cell_bg(cb5, "F5EEF8")
p = cb5.paragraphs[0]
para_space(p, before=4, after=4, line=15)
r = p.add_run(
    "　あなた、一番大事なことを言いますよ。\n"
    "　白山市というのはね、白山という霊峰を持つ、特別な土地なの。\n"
    "　そこで生まれたあなたには、先祖の霊気が強く宿っています。\n"
    "　でも東京に来てから、お墓参りはちゃんとできていますか？\n"
    "　先祖供養を怠ると、どんなに自分が頑張っても、見えないところで\n"
    "　引っ張られるような感覚が出てくる。運気が上がろうとしても上がれない。\n\n"
    "　乱気の年こそ、先祖に手を合わせることが一番の開運法です。\n"
    "　お盆の時期（7月・8月）に必ず石川に帰って、お墓参りをしなさい。\n"
    "　仏壇があれば毎日手を合わせる。水とお花を絶やさないこと。\n"
    "　これが今のあなたにできる、最大の運気アップです。\n"
    "　星の力を借りるより、先祖の加護の方がよほど強い。覚えておきなさい。"
)
set_font(r, 9.5, name="MS Mincho", color=(60, 0, 100))

# ══ 月別ポイント ════════════════════════════════════════════
p = doc.add_paragraph()
para_space(p, before=8, after=2, line=13)
r = p.add_run("■ 月別　重要ポイント（2026年７月〜12月）")
set_font(r, 10, bold=True, color=(100, 50, 0))

months = [
    ("７月", "FFF9C4", "乱気の影響が本格化。感情的になりやすい。大きな決断は必ず先送り。先祖供養を始める絶好の機会。"),
    ("８月", "FFCCBC", "エネルギーの消耗が激しい月。無理をしない。夏バテに注意。お盆に必ず帰省し墓参りを。"),
    ("９月", "E8F5E9", "少し落ち着きを取り戻せる月。ただし油断禁物。人間関係の整理に良い時期。信頼できる人を選ぶ。"),
    ("10月", "E3F2FD", "金銭面の注意が必要な月。支出を見直す。投資・保証人・大きな買い物は厳禁。"),
    ("11月", "F3E5F5", "仕事で評価される兆しが見える月。ただし出しゃばらず、謙虚さを保つことが吉。"),
    ("12月", "FFF3E0", "年末の疲れが出やすい。体調管理を最優先に。来年への準備を静かに始める時期。"),
]

tbl6 = doc.add_table(rows=len(months), cols=2)
tbl6.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl6.style = "Table Grid"
for i, (month, color, text) in enumerate(months):
    cm = tbl6.rows[i].cells[0]
    ct = tbl6.rows[i].cells[1]
    set_cell_bg(cm, color)
    set_cell_bg(ct, "FAFAFA")
    p_m = cm.paragraphs[0]
    para_space(p_m, before=3, after=3, line=12)
    p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rm = p_m.add_run(month)
    set_font(rm, 11, bold=True, color=(80, 40, 0), name="MS Mincho")
    p_t = ct.paragraphs[0]
    para_space(p_t, before=3, after=3, line=12)
    rt = p_t.add_run(text)
    set_font(rt, 8.5, name="MS Mincho")

# ══ 総括メッセージ ══════════════════════════════════════════
p = doc.add_paragraph()
para_space(p, before=8, after=2, line=13)
r = p.add_run("■ 細木数子からのメッセージ")
set_font(r, 10, bold=True, color=(100, 50, 0))

tbl7 = doc.add_table(rows=1, cols=1)
tbl7.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl7.style = "Table Grid"
cb7 = tbl7.rows[0].cells[0]
set_cell_bg(cb7, "1A1A2E")
p = cb7.paragraphs[0]
para_space(p, before=5, after=5, line=16)
r = p.add_run(
    "　あなたはね、本当にいい星を持っている人なの。\n"
    "　金星人というのはね、磨けば磨くほど輝く星なのよ。\n\n"
    "　でも今年の後半は、磨く前にまず「土台を固める」時期。\n"
    "　輝こうとして動き回るんじゃなくて、じっとして根を張る。\n"
    "　木だって、冬の間に根を張るから春に花が咲くでしょう。\n\n"
    "　東京という土地はエネルギーが強すぎる。\n"
    "　だからこそ、石川の故郷のことを忘れないでいなさい。\n"
    "　白山の霊気は、あなたを守り続けているの。\n\n"
    "　７月以降、嵐が来ても怖くない。\n"
    "　先祖を大切にして、足元を固めれば、\n"
    "　来年からのあなたは本当に花開きます。\n"
    "　信じていいのよ、あなたの運命を。"
)
set_font(r, 10, name="MS Mincho", color=(212, 175, 55))

# ══ フッター ════════════════════════════════════════════════
p = doc.add_paragraph()
para_space(p, before=6, after=0, line=10)
r = p.add_run(
    "※ 本鑑定は六星占術の理論に基づき作成したものです。"
    "　鑑定結果はあくまでも参考としてお取り扱いください。"
    "　　　　令和８年（2026年）６月吉日"
)
set_font(r, 7, color=(120, 120, 120))

output_path = "/home/user/con30/fortune_reading_20260730生まれA型.docx"
doc.save(output_path)
print(f"Done: {output_path}")
