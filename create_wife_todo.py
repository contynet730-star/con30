#!/usr/bin/env python3
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# 日本語フォント設定
style = doc.styles['Normal']
style.font.name = 'MS Gothic'
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'MS Gothic')
style.font.size = Pt(10.5)

# ページ余白
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)


def set_jp_font(run, size=10.5, bold=False, color=None):
    run.font.name = 'MS Gothic'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'MS Gothic')
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def add_heading(text, level=1, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    sizes = {1: 18, 2: 14, 3: 12}
    set_jp_font(run, size=sizes.get(level, 11), bold=True, color=color)
    return p


def add_para(text, bold=False, color=None, size=10.5):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_jp_font(run, size=size, bold=bold, color=color)
    return p


def add_bullet(text, indent=0):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    set_jp_font(run, size=10.5)
    p.paragraph_format.left_indent = Cm(0.5 + indent * 0.5)
    return p


def add_table(headers, rows, col_widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = ''
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(h)
        set_jp_font(run, size=10, bold=True)

    for row_data in rows:
        row_cells = table.add_row().cells
        for i, cell_data in enumerate(row_data):
            row_cells[i].text = ''
            p = row_cells[i].paragraphs[0]
            run = p.add_run(str(cell_data))
            set_jp_font(run, size=10)

    if col_widths:
        for row in table.rows:
            for i, width in enumerate(col_widths):
                row.cells[i].width = Cm(width)
    return table


# ====== タイトル ======
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('妻のやることリスト')
set_jp_font(run, size=22, bold=True, color=RGBColor(0x1F, 0x49, 0x7D))

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('家計改善・節税・障害年金 統合プラン')
set_jp_font(run, size=12, color=RGBColor(0x4F, 0x4F, 0x4F))

info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = info.add_run('作成日：2026年5月31日 / 対象期間：44歳〜60歳定年')
set_jp_font(run, size=9, color=RGBColor(0x80, 0x80, 0x80))

doc.add_paragraph()

# ====== 概要 ======
add_heading('■ 全体サマリ', level=1, color=RGBColor(0x1F, 0x49, 0x7D))

add_table(
    ['項目', '内容'],
    [
        ['対象', '妻（44歳・東京都公立小学校 主任教諭）'],
        ['年収', '額面 約750万円（手取り 約530-540万円）'],
        ['持病', '1型糖尿病（月3万円の医療費自己負担）'],
        ['退職予定', '60歳定年（勤続36年）'],
        ['退職金見込', '約 2,500万円'],
        ['年金見込（65歳〜）', '月 約 19.8万円'],
    ],
    col_widths=[5, 11]
)

doc.add_paragraph()
add_para('このリストは「すぐやる順」に並んでいます。AからHまで順番に着手してください。', bold=True)
doc.add_paragraph()

# ====== A. 障害年金申請 ======
add_heading('A. 障害年金の申請（最優先）', level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_para('【期待効果】月 約6万円（年72万円）の非課税収入。生涯約 1,500万円。', bold=True, color=RGBColor(0xC0, 0x00, 0x00))
doc.add_paragraph()

add_heading('▼ Step 1：主治医に相談（次回通院時）', level=3)
add_bullet('「障害年金の対象になりますか？」と質問する')
add_bullet('Cペプチド値（血清）を確認してもらう（0.3ng/mL未満なら該当の可能性大）')
add_bullet('「障害年金の診断書を書いていただけますか？」と確認')
add_bullet('初診日（糖尿病と診断された日）を確認')

doc.add_paragraph()
add_heading('▼ Step 2：初診医療機関に証明書依頼', level=3)
add_bullet('現在の主治医が初診医なら不要')
add_bullet('別の病院が初診なら「受診状況等証明書」を依頼（費用 約3,000-5,000円）')

doc.add_paragraph()
add_heading('▼ Step 3：公立学校共済組合に書類請求', level=3)
add_table(
    ['項目', '内容'],
    [
        ['提出先', '公立学校共済組合 東京支部'],
        ['住所', '東京都新宿区西新宿2-7-1（小田急第一生命ビル）'],
        ['電話', '03-5320-7421'],
        ['請求書類', '年金請求書（障害厚生年金）一式'],
    ],
    col_widths=[4, 12]
)

doc.add_paragraph()
add_heading('▼ Step 4：必要書類を揃える', level=3)
add_bullet('年金請求書（障害厚生年金）')
add_bullet('診断書（糖尿病用）※主治医作成 / 費用 約5,000-10,000円')
add_bullet('受診状況等証明書（初診医療機関で取得）')
add_bullet('病歴・就労状況等申立書（本人記入）')
add_bullet('戸籍謄本（区役所 450円）')
add_bullet('住民票（区役所 300円）')
add_bullet('年金手帳・基礎年金番号通知書')
add_bullet('預金通帳の写し')
add_bullet('印鑑')

doc.add_paragraph()
add_heading('▼ Step 5：提出 → 審査（3〜4ヶ月）→ 認定', level=3)
add_bullet('提出から受給開始まで 約半年')
add_bullet('認定後、過去5年分まで遡及受給可能（最大300-400万円）')

doc.add_paragraph()
add_heading('▼ 社労士活用の選択肢', level=3)
add_para('自分で申請が不安なら、障害年金専門社労士への依頼も可。')
add_bullet('「障害年金 社労士 東京 糖尿病」で検索')
add_bullet('初回相談 無料が多い')
add_bullet('成功報酬：年金額の1-2ヶ月分（約8-12万円）')
add_bullet('認定率：自分で50-60% → 社労士80-90%')

doc.add_page_break()

# ====== B. 医療費控除 ======
add_heading('B. 医療費控除の申請（過去5年分も含む）', level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_para('【期待効果】年 約7.8万円の節税。過去5年遡及で 約40万円の還付。', bold=True, color=RGBColor(0xC0, 0x00, 0x00))
doc.add_paragraph()

add_heading('▼ Step 1：マイナポータルで医療費通知を取得', level=3)
add_bullet('マイナポータルにログイン → 「医療費通知」')
add_bullet('過去5年分（2021-2025年）をダウンロード')
add_bullet('e-Tax連携で確定申告に自動取込可能')

doc.add_paragraph()
add_heading('▼ Step 2：対象となる費用を確認', level=3)
add_bullet('インスリン製剤・自己注射用品')
add_bullet('血糖測定器・センサー（リブレ等）')
add_bullet('糖尿病内科通院費（診察料・検査料）')
add_bullet('処方薬（家族全員分も合算可）')
add_bullet('通院の交通費（電車・バス）')
add_bullet('歯科治療（家族全員）')
add_bullet('市販薬（家族全員分）')

doc.add_paragraph()
add_heading('▼ Step 3：過去5年分を遡及申告（最重要）', level=3)
add_para('未申告だった場合、過去5年分まで遡って還付請求可能（更正の請求）。')

add_table(
    ['年度', '還付見込額'],
    [
        ['2021年分', '約 7.8万円'],
        ['2022年分', '約 7.8万円'],
        ['2023年分', '約 7.8万円'],
        ['2024年分', '約 7.8万円'],
        ['2025年分', '約 7.8万円'],
        ['合計', '約 39万円'],
    ],
    col_widths=[6, 10]
)

doc.add_paragraph()
add_heading('▼ Step 4：今後は毎年確定申告で申請', level=3)
add_bullet('妻名義で確定申告（家族全員の医療費合算）')
add_bullet('e-Taxで提出（自宅から完結）')
add_bullet('還付は約1ヶ月後に指定口座へ')

doc.add_paragraph()

# ====== C. iDeCo開設 ======
add_heading('C. iDeCoの開設（節税効果大）', level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_para('【期待効果】年 約7.9万円の節税。16年累計 約126万円。', bold=True, color=RGBColor(0xC0, 0x00, 0x00))
doc.add_paragraph()

add_heading('▼ Step 1：楽天証券で申込（10分）', level=3)
add_bullet('楽天証券にログイン → iDeCoタブ')
add_bullet('「個人型確定拠出年金 申込」')
add_bullet('職業：「公務員」を選択')
add_bullet('拠出額：月2万円（上限満額）に設定')
add_bullet('運用商品：楽天・オールカントリー または eMAXIS Slim オルカン 100%')

doc.add_paragraph()
add_heading('▼ Step 2：勤務先証明書の取得', level=3)
add_bullet('「事業所登録申請書 兼 第2号加入者に係る事業主の証明書」が必要')
add_bullet('勤務先（学校事務 or 教育委員会）に依頼')
add_bullet('発行に1-2週間')

doc.add_paragraph()
add_heading('▼ Step 3：書類提出 → 約1-2ヶ月で開始', level=3)
add_bullet('証明書を楽天証券に郵送 or アップロード')
add_bullet('国民年金基金連合会の審査（1-2ヶ月）')
add_bullet('開設後、月2万円自動拠出開始')

doc.add_paragraph()
add_heading('▼ 注意点', level=3)
add_bullet('60歳まで引き出し不可')
add_bullet('年末調整で「小規模企業共済等掛金控除」として申告')
add_bullet('医療費控除と併用可')

doc.add_paragraph()

# ====== D. ふるさと納税 ======
add_heading('D. ふるさと納税のフル活用', level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_para('【期待効果】年 約13万円の上限活用で、返礼品4-5万円相当＋ポイント還元。', bold=True, color=RGBColor(0xC0, 0x00, 0x00))
doc.add_paragraph()

add_heading('▼ 上限額（妻の場合）', level=3)
add_table(
    ['想定', '上限'],
    [
        ['扶養家族なし（高3を夫の扶養）', '約 13万円'],
        ['iDeCo月2万考慮後', '約 12.5万円'],
    ],
    col_widths=[10, 6]
)

doc.add_paragraph()
add_heading('▼ おすすめ配分（年12.5万円）', level=3)
add_table(
    ['区分', '金額', '備考'],
    [
        ['米（4人家族）', '4万円', '常備品'],
        ['肉類（牛・豚・鶏）', '3万円', '冷凍保存可'],
        ['日用品（トイレットペーパー等）', '2万円', '消耗品'],
        ['果物・嗜好品', '2万円', '季節物'],
        ['調味料', '1.5万円', '長期保存可'],
    ],
    col_widths=[6, 3, 7]
)

doc.add_paragraph()
add_heading('▼ おすすめサイト', level=3)
add_bullet('楽天ふるさと納税（楽天ポイント還元あり）')
add_bullet('5と0のつく日に寄附でポイント+2倍')
add_bullet('SPU活用で還元率最大化')

doc.add_paragraph()

# ====== E. NISA積立 ======
add_heading('E. NISA積立の継続（現状維持）', level=1, color=RGBColor(0x1F, 0x49, 0x7D))

add_para('【方針】現状の月3.5万円積立を継続。教育費別途確保と両立。', bold=True)
doc.add_paragraph()

add_heading('▼ 現状', level=3)
add_table(
    ['項目', '内容'],
    [
        ['積立額', '月 3.5万円（年42万円）'],
        ['商品', 'オルカン or S&P500（推奨）'],
        ['年間運用益（5%想定）', '約 2万円'],
        ['16年後の評価額', '約 1,250万円'],
    ],
    col_widths=[7, 9]
)

doc.add_paragraph()
add_heading('▼ いつ増額するか', level=3)
add_bullet('上の子大学卒業後（妻48歳・2030年）→ 教育費浮く')
add_bullet('障害年金認定後 → 月6万円の追加収入があれば積立増額可')
add_bullet('iDeCo節税分（年7.9万）と医療費控除還付（年7.8万）を積立に上乗せも可')

doc.add_paragraph()

# ====== F. 教育費別途確保 ======
add_heading('F. 教育費の別途確保（継続）', level=1, color=RGBColor(0x1F, 0x49, 0x7D))

add_para('【方針】現状の方針を継続。投資には回さず、安全資産で確保。', bold=True)
doc.add_paragraph()

add_heading('▼ 必要額の確認', level=3)
add_table(
    ['対象', '時期', '想定額'],
    [
        ['上の子（高3）クラーク', '2026年度', '年 80-160万'],
        ['上の子 大学進学', '2027-2030年度', '年 130-200万'],
        ['下の子（小5）高校', '2032-2034年度', '年 60-100万'],
        ['下の子 大学進学', '2035-2038年度', '年 130-200万'],
    ],
    col_widths=[6, 5, 5]
)

doc.add_paragraph()
add_heading('▼ 上の子の進路次第で大きく変動', level=3)
add_bullet('芸能事務所所属 → 学費激減')
add_bullet('一般大学進学 → 通常通り')
add_bullet('芸術系大学・海外 → +大幅増')
add_bullet('進路決定後に再度家計見直し')

doc.add_paragraph()

# ====== G. 健康管理 ======
add_heading('G. 1型糖尿病の継続管理', level=1, color=RGBColor(0x1F, 0x49, 0x7D))

add_para('【方針】60歳定年まで安定就労するため、健康管理を最優先。', bold=True)
doc.add_paragraph()

add_heading('▼ 定期チェック', level=3)
add_bullet('糖尿病内科 月1回通院（継続）')
add_bullet('HbA1c値・Cペプチド値の定期測定')
add_bullet('合併症チェック（網膜症・腎症・神経障害）')
add_bullet('歯科健診 年2回')
add_bullet('人間ドック 年1回')

doc.add_paragraph()
add_heading('▼ 治療選択肢の検討', level=3)
add_bullet('インスリンポンプ療法（コントロール改善）')
add_bullet('持続血糖モニター CGM（リブレ等）')
add_bullet('糖尿病療養指導士のサポート活用')

doc.add_paragraph()
add_heading('▼ 高額療養費の限度額適用認定証取得', level=3)
add_bullet('公立学校共済組合に申請')
add_bullet('窓口での自己負担が事前に上限額に抑えられる')

doc.add_paragraph()

# ====== H. 民間保険確認 ======
add_heading('H. 民間医療保険の保障内容確認', level=1, color=RGBColor(0x1F, 0x49, 0x7D))

add_para('【方針】既加入保険は絶対解約しない。保障内容を最大限活用。', bold=True)
doc.add_paragraph()

add_heading('▼ 確認すべき項目', level=3)
add_bullet('医療保険：入院給付金、手術給付金の額')
add_bullet('通院給付金：糖尿病治療継続中の通院は対象か')
add_bullet('特定疾病一時金：糖尿病が対象か')
add_bullet('がん保険：保障内容')
add_bullet('生命保険：保障額・受取人')
add_bullet('教員共済の付加給付（高額療養費後の追加給付）')

doc.add_paragraph()
add_heading('▼ 絶対NG行為', level=3)
add_bullet('新しい医療保険への乗換（既存解約しない）')
add_bullet('保険料節約のための保障減額')

doc.add_paragraph()
add_heading('▼ 推奨アクション', level=3)
add_bullet('既加入保険の証券を全て確認')
add_bullet('保険会社のコールセンターで「請求漏れがないか」確認')
add_bullet('過去の入院・通院で請求できるものがあれば申請')

doc.add_page_break()

# ====== スケジュール ======
add_heading('■ 全体スケジュール', level=1, color=RGBColor(0x1F, 0x49, 0x7D))

add_table(
    ['時期', 'やること', '優先度'],
    [
        ['2026年6月（今月）', '主治医に障害年金相談・Cペプチド値確認', '★★★'],
        ['2026年6月', 'マイナポータルで医療費通知ダウンロード', '★★★'],
        ['2026年7月', '楽天証券でiDeCo申込・勤務先証明書依頼', '★★★'],
        ['2026年7-8月', '障害年金 必要書類取得（診断書・証明書）', '★★★'],
        ['2026年9月', '公立学校共済組合に書類提出', '★★★'],
        ['2026年10-12月', 'ふるさと納税 年内に上限まで', '★★'],
        ['2026年12月', '年末調整でiDeCo控除申告', '★★'],
        ['2027年2-3月', '確定申告（医療費控除・過去5年分も）', '★★★'],
        ['2027年3-4月', '障害年金認定・受給開始', '★★★'],
        ['2027年5月', '過去5年分の還付金受領', '★★★'],
        ['毎月', '糖尿病内科通院・健康管理', '★★★'],
        ['毎年', '医療費控除・ふるさと納税の継続', '★★'],
    ],
    col_widths=[4, 9, 3]
)

doc.add_paragraph()

# ====== 期待効果まとめ ======
add_heading('■ 全部実行した場合の経済効果', level=1, color=RGBColor(0x00, 0x70, 0xC0))

add_table(
    ['項目', '年間効果', '生涯累計（30年）'],
    [
        ['障害年金（3級認定時）', '+72万円', '+1,150万円（16年）'],
        ['障害年金 過去5年遡及', '一括', '+300-360万円'],
        ['医療費控除（毎年）', '+7.8万円', '+230万円'],
        ['医療費控除 過去5年遡及', '一括', '+39万円'],
        ['iDeCo節税', '+7.9万円', '+126万円（16年）'],
        ['ふるさと納税', '+2.8万円相当', '+84万円'],
        ['ふるさと納税 返礼品', '4-5万円相当', '+120-150万円'],
        ['合計（生涯）', '—', '約 2,000-2,200万円'],
    ],
    col_widths=[7, 4, 5]
)

doc.add_paragraph()

# ====== 連絡先一覧 ======
add_heading('■ 連絡先・参考情報', level=1, color=RGBColor(0x1F, 0x49, 0x7D))

add_table(
    ['用途', '連絡先・URL'],
    [
        ['公立学校共済組合 東京支部', '03-5320-7421'],
        ['年金事務所（一般相談）', 'ねんきんダイヤル 0570-05-1165'],
        ['国税庁 e-Tax', 'https://www.e-tax.nta.go.jp/'],
        ['マイナポータル', 'https://myna.go.jp/'],
        ['楽天証券（iDeCo）', 'https://www.rakuten-sec.co.jp/'],
        ['楽天ふるさと納税', 'https://event.rakuten.co.jp/furusato/'],
        ['東京都教育委員会', '03-5320-6720'],
    ],
    col_widths=[6, 10]
)

doc.add_paragraph()

# ====== 最後のメモ ======
add_heading('■ 大切なメッセージ', level=1, color=RGBColor(0xC0, 0x00, 0x00))

p = doc.add_paragraph()
run = p.add_run('障害年金の申請は、「障害者になる」ことではなく、')
set_jp_font(run, size=11)

p = doc.add_paragraph()
run = p.add_run('「1型糖尿病という持病に対する国の制度を使う」だけのことです。')
set_jp_font(run, size=11)

p = doc.add_paragraph()
run = p.add_run('既に治療を続けているあなたが、その努力に対して受け取るべき正当な権利です。')
set_jp_font(run, size=11, bold=True, color=RGBColor(0xC0, 0x00, 0x00))

doc.add_paragraph()

p = doc.add_paragraph()
run = p.add_run('60歳定年まで健康に勤続し、悠々自適のセカンドライフを迎えるために、')
set_jp_font(run, size=11)

p = doc.add_paragraph()
run = p.add_run('AからCまでだけでも、ぜひ今月中に着手してください。')
set_jp_font(run, size=11, bold=True)

doc.add_paragraph()

footer = doc.add_paragraph()
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer.add_run('— 以上 —')
set_jp_font(run, size=10, color=RGBColor(0x80, 0x80, 0x80))

# 保存
output_path = '/home/user/con30/妻のやることリスト.docx'
doc.save(output_path)
print(f'Created: {output_path}')
