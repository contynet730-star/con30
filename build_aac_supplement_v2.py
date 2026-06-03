# -*- coding: utf-8 -*-
"""
補足説明資料 v2（Opus 4.8 再整理版を全面採用）：
「AAC・CCAは本当にあり得ないのか？」
「行動観察・授業中の発言の評価は主観的ではないのか？」
─ 公的資料の徹底調査に基づく、現場の8割が納得できる回答 ─
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

doc = Document()
for sec in doc.sections:
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0); sec.right_margin = Cm(2.0)
style = doc.styles['Normal']
style.font.name = 'メイリオ'; style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'メイリオ')

NAVY=RGBColor(0x1F,0x3A,0x6E); ORANGE=RGBColor(0xEB,0x6C,0x15); GREEN=RGBColor(0x07,0x70,0x40)
GRAY=RGBColor(0x40,0x40,0x40); DEEP_WATER=RGBColor(0x2A,0x96,0xC2); RED=RGBColor(0xC0,0x40,0x40)
PURPLE=RGBColor(0x5B,0x2C,0x6F); LIGHT_YELLOW='FFF7D9'; LIGHT_BLUE='E5F2FA'

def H1(t):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(t); r.font.size=Pt(20); r.font.bold=True; r.font.color.rgb=NAVY
    r.font.name='メイリオ'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')

def H2(t):
    p=doc.add_paragraph(); r=p.add_run(t)
    r.font.size=Pt(15); r.font.bold=True; r.font.color.rgb=NAVY
    r.font.name='メイリオ'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')
    p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(6)

def H3(t, color=ORANGE):
    p=doc.add_paragraph(); r=p.add_run(t)
    r.font.size=Pt(12); r.font.bold=True; r.font.color.rgb=color
    r.font.name='メイリオ'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')
    p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(4)

def P(t, size=10.5, bold=False, color=None, indent=False):
    p=doc.add_paragraph()
    if indent: p.paragraph_format.left_indent=Cm(0.8)
    p.paragraph_format.line_spacing=1.5; p.paragraph_format.space_after=Pt(3)
    r=p.add_run(t); r.font.size=Pt(size); r.font.bold=bold
    if color: r.font.color.rgb=color
    r.font.name='メイリオ'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')

def QUOTE(text, source, color_bg=LIGHT_YELLOW):
    """引用ボックス"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6); p.paragraph_format.right_indent = Cm(0.6)
    p.paragraph_format.line_spacing = 1.4; p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_bg}"/>')
    pPr.append(shd)
    r = p.add_run('▼ ')
    r.font.size=Pt(10); r.font.color.rgb=ORANGE; r.font.bold=True
    r.font.name='メイリオ'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')
    r2 = p.add_run(text)
    r2.font.size=Pt(10.5)
    r2.font.name='メイリオ'; r2._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')
    p2 = doc.add_paragraph()
    p2.paragraph_format.left_indent = Cm(0.6)
    p2.paragraph_format.space_after = Pt(6)
    rs = p2.add_run(f'　─ {source}')
    rs.font.size=Pt(9); rs.font.italic=True; rs.font.color.rgb=GRAY
    rs.font.name='メイリオ'; rs._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')

def HR():
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    r=p.add_run('─'*45); r.font.color.rgb=GRAY; r.font.size=Pt(9)

def CAUTION(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3); p.paragraph_format.right_indent = Cm(0.3)
    p.paragraph_format.line_spacing = 1.4; p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    pPr = p._p.get_or_add_pPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFE5E5"/>')
    pPr.append(shd)
    r = p.add_run('【重要】')
    r.font.size=Pt(11); r.font.color.rgb=RED; r.font.bold=True
    r.font.name='メイリオ'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')
    r2 = p.add_run(text)
    r2.font.size=Pt(10.5); r2.font.bold=True
    r2.font.name='メイリオ'; r2._element.rPr.rFonts.set(qn('w:eastAsia'),'メイリオ')

# ============================================================
# 表紙
# ============================================================
H1('石神井南中学校　校内研修　補足説明資料')
H1('「AAC・CCA問題」と「行動観察・発言の評価」')
P('')
P('─ 公的資料の徹底調査に基づく、現場で納得していただける回答 ─', color=DEEP_WATER, size=14, bold=True)
P('')
P('日時：令和８年６月３日（水）研修当日に出された質問への回答', bold=True)
P('回答者：練馬区教育委員会　指導主事　紺多　章一郎')
HR()
P('【今日いただいた3つの質問】', bold=True, color=NAVY, size=12)
P('質問① 「AAC」「CCA」は「あり得ない」と言われるが、実際の現場では', size=11)
P('　　　 そうした評価になる生徒がいる。なぜ「あり得ない」のか？', size=11)
P('質問② 授業中の発言を毎時間拾うのは現実的でない。また「発言を評価する」', size=11)
P('　　　 と宣言すると、賢い生徒だけが評価目当てで発言しないか？', size=11)
P('質問③ 教師による行動観察は、主観が入るのではないか？', size=11)
HR()
P('【本資料の方針】', bold=True, color=NAVY, size=12)
P('・公的資料の原文を直接引用（中教審報告／国研参考資料／東京都教委）', size=10.5)
P('・「ありえない」の正しい意味（＝禁止ではなく診断アラーム）の提示', size=10.5)
P('・現場で「ああ、なるほど」と納得していただける根拠提示を目指す', size=10.5, color=ORANGE, bold=True)
P('・2025年7月の国の動向（次期指導要領）も援護射撃として明記', size=10.5)

doc.add_page_break()

# ============================================================
# 第0部　まず結論
# ============================================================
H1('第0部　まず結論 ─「ありえない」の正しい意味')

H2('先生方の引っかかりに、まず正面から答えます')

CAUTION('「ありえない」＝「付けてはいけない／だから帳尻を合わせろ」'
        'という意味ではありません。公的資料が言っているのは、'
        '「単元末・学期末の総括でAAC/CCAが出たら、生徒の問題ではなく、'
        '評価方法や指導が機能していない可能性を示す“診断アラーム”として扱い、'
        '原因を検討して改善せよ」ということです。')

H3('国立教育政策研究所（NIER）の核心的な記述')

QUOTE('3つの観点の評価の結果は、ばらつき（「CCA」や「AAC」、「CAA」や「ACA」等）が'
      '生じなくなるよう指導や学習改善を図る必要がある。',
      '国立教育政策研究所「『指導と評価の一体化』のための学習評価に関する参考資料」第1編 総説',
      color_bg=LIGHT_BLUE)

P('ここでの言い方に注目してください ──', bold=True)
P('・「生じてはならない」（禁止）ではない', color=NAVY)
P('・「生じなくなるよう（指導や学習改善を）図る必要がある」（教師側へのアクション要求）', color=NAVY, bold=True)
P('')
P('つまり、先生方は「帳尻合わせ」を求められているのではなく、', bold=True, color=ORANGE)
P('「立ち止まって点検せよ」と言われているのです。', bold=True, color=ORANGE)
P('')
P('ここを最初に共有していただくと、現場の納得感がまるで変わります。', bold=True)

doc.add_page_break()

# ============================================================
# 第Ⅰ部　なぜ理論上AAC・CCAは整合しないのか
# ============================================================
H1('第Ⅰ部　なぜ理論上 AAC・CCA は整合しないのか')

H2('鍵：「主体的に学習に取り組む態度」は性格や努力量を測る独立観点ではない')

P('中教審H31報告が定義した「主体的に学習に取り組む態度」の中身は、次の2側面 ── ', bold=True)

QUOTE('「主体的に学習に取り組む態度」の評価については、'
      '①知識及び技能を獲得したり、思考力、判断力、表現力等を身に付けたりすることに向けた'
      '粘り強い取組を行おうとする側面と、'
      '②①の粘り強い取組を行う中で、自らの学習を調整しようとする側面、'
      'という二つの側面を評価することが求められる。',
      '中央教育審議会報告（平成31年1月21日）')

P('この定義から見えてくることは ──', bold=True)
P('・「主体的態度」は、知識・思考を生み出す“その同じ学習プロセス”の意思的な側面（＝エンジン）', color=NAVY)
P('・出力（知識・思考）とエンジン（態度）が極端に食い違うのは、論理的に不自然', color=NAVY, bold=True)

H2('AAC・CCA それぞれの「最頻原因」')

H3('AAC（知識Ａ・思考Ａ・態度Ｃ）が示すもの')

P('A・Aに到達した以上、相応の「粘り強さ」と「自己調整」があったはず。', bold=True)
P('にもかかわらず態度だけCになるのは ──', bold=True, color=RED)
P('・態度を「挙手・私語・忘れ物・提出の雑さ・態度の悪さ」など、', indent=True)
P('  学習と切り離した形式面・性格面で測っているサイン', indent=True, color=RED)
P('・中教審が「性格や行動面の傾向の評価ではない」と明言した、', indent=True)
P('  まさにその誤りの疑い', indent=True, color=RED)

H3('CCA（知識Ｃ・思考Ｃ・態度Ａ）が示すもの')

P('自己調整（「分かっていない」と気づき方法を変える）が本当に働いていれば、', bold=True)
P('単元を通じて知識か思考のどちらかは動くはず。', bold=True)
P('両方Cのまま態度だけAになるのは ──', bold=True, color=RED)
P('・「がんばっている」「真面目」「前向き」という、', indent=True)
P('  努力姿勢や人柄を「態度A」にしている可能性', indent=True, color=RED)
P('・旧「関心・意欲・態度」の“がんばり賞”を、', indent=True)
P('  新観点が最も排除したかった失敗形', indent=True, color=RED)

H2('中教審・NIERが繰り返し戒めるポイント')

QUOTE('挙手の回数等、その形式的態度を評価することは適当ではなく、'
      '他の観点に関わる児童生徒の学習状況と照らし合わせながら'
      '学習や指導の改善を図ることが重要である。',
      '中央教育審議会報告（平成31年1月21日）')

QUOTE('ノートにおける特定の記述などを取り出して、他の観点から切り離して'
      '「主体的に学習に取り組む態度」として評価することは適切ではない。',
      '奈良県教育委員会「主体的に学習に取り組む態度の評価に関するガイドライン」')

doc.add_page_break()

# ============================================================
# 第Ⅱ部　「でも実際にそういう生徒はいる」── この反論は正しい
# ============================================================
H1('第Ⅱ部　「でも実際にそういう生徒はいる」── この反論は正しい')

H2('現場の観察は間違っていません。NIERも認めています')

QUOTE('単元の導入の段階では観点別の学習状況にばらつきが生じるとしても、'
      '指導と評価の取組を重ねながら授業を展開することにより、'
      '単元末や学期末、学年末の結果として算出される3段階の観点別学習状況の評価については、'
      '観点ごとに大きな差は生じないものと考えられる。',
      '国立教育政策研究所「『指導と評価の一体化』のための学習評価に関する参考資料」第1編 総説',
      color_bg=LIGHT_BLUE)

P('つまりNIERの公式見解は、', bold=True, color=NAVY)
P('「瞬間的・途中的なAAC/CCAは当然ある」', bold=True, color=NAVY)
P('「問題は、それをそのまま“通知表の総括”として固定すること」', bold=True, color=NAVY)

H2('現場で観察される“AAC/CCAっぽい生徒”は、4つに整理できる')

P('先生方には、「その生徒はこの4つのどれですか？」と一緒に切り分けるのが有効です。', bold=True)

H3('1. 態度評価が形式・性格面に汚染されている（最頻原因）')
P('→ 評価方法を直す', indent=True, bold=True, color=ORANGE)
P('  「挙手」「私語」「忘れ物」「提出の有無」など、学習と切り離した形式・性格面を', indent=True)
P('  態度の評価に混入させていないか教科会で点検する', indent=True)

H3('2. 1時間・1方法の点で見ている（過程を見ていない）')
P('→ 単元全体・複数資料に広げる', indent=True, bold=True, color=ORANGE)
P('  ノート記述・振り返り・自己評価・行動観察を組み合わせて、単元の経過で見る', indent=True)

H3('3. 本当に単元途中の一過性の状態')
P('→ 想定内。記録ではなく指導でギャップを閉じる', indent=True, bold=True, color=ORANGE)
P('  単元末まで放置せず、形成的評価で早期に修正するのが本旨', indent=True)

H3('4. 評価規準が観点の趣旨とズレている')
P('→ 規準を作り直す', indent=True, bold=True, color=ORANGE)
P('  「Ｂの姿」を、性格や努力量ではなく、', indent=True)
P('  「学習目標に向けた粘り強さ・自己調整の姿」として教科会で再言語化', indent=True)

H2('本物のエッジケース ── 正直に扱うべき2つ')

CAUTION('以下の2ケースは、現場で本当に悩ましい「本物のエッジケース」です。'
        'ごまかすと現場の信頼を失いますので、正直に扱います。')

H3('エッジケース① AAC型：できてしまう子が退屈して見える')

P('・態度は「努力の見た目の量」ではない', color=NAVY)
P('・効率よく自己調整して楽々ゴールに達するのも、立派に主体的', color=NAVY)
P('・本当に退屈・不関与なら、それは「課題がその子に易しすぎるサイン」', color=NAVY)
P('・→ 教師の指導改善（挑戦的課題・発展学習）で応える', color=ORANGE, bold=True)
P('・Cは生徒の人格評価ではなく、教師の課題設計へのフィードバック', color=ORANGE, bold=True)

H3('エッジケース② CCA型：懸命に取り組み、助けを求め、修正もするが、出力が伸びない子')

P('・いちばん難しいケース', color=NAVY)
P('・現行制度の答え：', color=NAVY, bold=True)
P('　(i) 態度の規準を「努力」ではなく「この単元の目標に向けた自己調整」に結び直す', indent=True, color=ORANGE)
P('　(ii) 両方Cのままなのは、自己調整を支える指導が届いていない教師側のシグナル', indent=True, color=ORANGE)
P('　 → 指導方法を変える（足場かけ・段階的支援・方法の例示）', indent=True, color=ORANGE, bold=True)

doc.add_page_break()

# ============================================================
# 第Ⅲ部　「なりそうなとき」どう回避・指導するか（段階別）
# ============================================================
H1('第Ⅲ部　「なりそうなとき」どう回避・指導するか（段階別）')

H2('現場で一番効くのは「②形成的評価（単元途中）」での早期発見・指導')

P('「単元末に出てから困る」のではなく、「途中で見つけて指導で閉じる」', bold=True, color=ORANGE)
P('これが“指導と評価の一体化”の本旨で、AAC/CCA回避の核心です。', bold=True, color=ORANGE)

H3('① 単元設計（事前）')
P('やること：', bold=True)
P('・3観点の規準を、同じ学習活動から導く', indent=True)
P('・態度の規準に「挙手・提出・忘れ物・私語」など形式要素を入れない', indent=True)
P('・態度は「粘り強さ」「自己調整」の2側面のみ', indent=True)
P('公的根拠：中教審報告（2側面）／NIER参考資料（規準は観点の趣旨から）', color=GRAY, size=10)

H3('② 形成的評価（単元途中）── ★最も効く段階')
P('やること：', bold=True)
P('・中間で「態度C傾向」を早期発見し、原因診断', indent=True)
P('・原因に応じて指導を入れる：', indent=True)
P('　 「課題が難しすぎる」→ 足場かけ', indent=True)
P('　 「方法が分からない」→ 方法の例示', indent=True)
P('　 「見通しがない」→ 振り返りの問い直し', indent=True)
P('・AAC傾向（優秀で不関与）→ 挑戦的課題で応える', indent=True)
P('公的根拠：NIER「指導と評価の取組を重ねて差を縮める」', color=GRAY, size=10)

H3('③ 総括（単元末）')
P('やること：', bold=True)
P('・1時間・1方法で決めない', indent=True)
P('・ノート記述、発言、行動観察、自己評価、相互評価を', indent=True)
P('  「材料の一つ」として、学習過程全体から総合判断', indent=True)
P('・他観点と照らし合わせて評価する', indent=True)
P('公的根拠：東京都教委「指導と評価の一体化を目指して」理論編／中教審「他の観点と照らし合わせ」', color=GRAY, size=10)

H3('④ 組織的担保')
P('やること：', bold=True)
P('・教科会・学年でモデレーション（評価のすり合わせ）を行う', indent=True)
P('・ルーブリックを生徒と事前共有する', indent=True)
P('公的根拠：NIER／各教委ガイドライン', color=GRAY, size=10)

H3('⑤ 結果点検')
P('やること：', bold=True)
P('・それでもAAC/CCAが出たら、生徒の評定を変えるのではなく、', indent=True)
P('  形式要素の混入・規準のズレ・指導の届きを点検する', indent=True)
P('公的根拠：NIER「原因を検討し速やかに学習・指導の改善」／奈良県教委', color=GRAY, size=10)

doc.add_page_break()

# ============================================================
# 第Ⅳ部　最新の国の動き ── 先生方の悩みは制度側も認めた
# ============================================================
H1('第Ⅳ部　最新の国の動き ─ 先生方の悩みは“制度側も認めた”')

H2('2025年7月：国は「態度を評定から外す」方針を示しました')

P('まさにAAC/CCAの構造的な難しさが理由で、国は態度の扱いを変えようとしています。', bold=True)

P('・2025年7月4日、中教審 教育課程企画特別部会で、', bold=True)
P('  「主体的に学習に取り組む態度」を含む「学びに向かう力・人間性等」を、', bold=True)
P('  評定（目標準拠評価＝A/B/C）から外し、', bold=True, color=ORANGE)
P('  教育課程全体を通じた個人内評価とする方針が示された', bold=True, color=ORANGE)
P('  （2030年度以降の次期学習指導要領に向けて検討中）', bold=True)

H3('国が挙げた理由 ── 先生方の悩みと地続き')
P('国自身が挙げている理由は ── ', bold=True)
P('・「提出物等の形式的な勤勉さで測られてしまう」', indent=True)
P('・「教師の期待する振る舞いを子どもが過度に意識する」', indent=True)
P('・「態度の評価が難しく評定が低く／付けられず、', indent=True)
P('  　不登校児童生徒等で不利になる」', indent=True)
P('')
P('これらは ── まさに今日、先生方が直面しているAAC/CCA問題と地続きです。', bold=True, color=ORANGE)

H3('現場へのメッセージ')
CAUTION('先生方の違和感は正しい。国もその難しさを認め、'
        '態度を評定から外す方向で動いている。'
        'ただし現行制度では態度もA/B/Cで付ける必要があるので、'
        'それまでは「診断アラームとして使い、指導で閉じる」運用が正解です。')

P('※ これは検討段階の今後の方向性です。現行の通知表は引き続き', size=10, color=GRAY)
P('　 態度も観点別A/B/C対象ですので、混同しないようご注意ください。', size=10, color=GRAY)

doc.add_page_break()

# ============================================================
# 第Ⅴ部　「行動観察」と「授業中の発言」の評価
# ============================================================
H1('第Ⅴ部　「行動観察」と「授業中の発言」の評価')

H2('公的資料が示す「主体的態度」の評価方法')

QUOTE('「主体的に学習に取り組む態度」については、ノートやレポート等における記述、'
      '授業中の発言、教師による行動観察や、児童生徒による自己評価や相互評価等の状況を、'
      '教師が評価を行う際に考慮する材料の一つとして用いることなどが考えられる。',
      '中央教育審議会報告（平成31年1月21日）')

P('注目すべきは ──', bold=True)
P('・「評価する材料」ではなく「考慮する材料の一つ」と書かれている', color=ORANGE, bold=True)
P('・つまり、行動観察も発言も、単独で評価を決定するものではない', color=ORANGE, bold=True)
P('・複数の材料を「組み合わせて」評価する、というのが公的趣旨', color=ORANGE, bold=True)

H2('質問への回答：「授業中の発言」を毎時間拾うのか？')

H3('回答：毎時間拾う必要はありません')

P('公的指針は、明確にこう述べています ──')
QUOTE('「主体的に学習に取り組む態度」については、行動の様子の観察等を通じ、'
      '長期的な視野で評価するものであり、必ずしも毎時間や単元の冒頭で'
      '評価する必要はない。',
      '中央教育審議会報告（平成31年1月21日）の趣旨')

P('現場での実践 ──', bold=True)
P('① 単元全体（例えば8〜10時間）を通じて、主体的態度を見取る場面を1〜2回設計', indent=True)
P('② 毎時間記録するのは負担も大きく、本来の趣旨でもない', indent=True)
P('③ 単元末の振り返りシート、レポート、自己評価などを中心に見取る', indent=True)
P('④ 授業中の発言は「補助的な材料」として、印象に残った生徒のみ簡単にメモ', indent=True)

H2('質問への回答：「発言を評価する」と宣言すると、賢い子だけが発言しないか？')

H3('回答：ご指摘は正しい。だからこそ「発言量」は評価しないのが原則です')

QUOTE('単に継続的な行動や積極的な発言等を行うなど、'
      '性格や行動面の傾向を評価するということではない。',
      '中央教育審議会報告（平成31年1月21日）')

P('「発言を評価する」と教師が宣言した時点で、公的趣旨と外れます。', bold=True, color=RED)
P('発言の「回数」「積極性」を評価することは、中教審H31報告で明確に否定されています。', bold=True)

H3('では、何を見るのか')
P('・発言の「内容」（学習目標に対する深まり、他者の意見への応答、自分の考えの修正）', indent=True)
P('・発言だけでなく、ノート記述、振り返り、自己評価などを「組み合わせて」見る', indent=True)
P('・「発言が苦手な生徒」が不利にならない設計が必須', indent=True)

H3('現場での実践指針')
P('① 「発言の回数」「挙手」は評価対象としない、と教師自身が明確に意識する', indent=True)
P('② 生徒に「発言を評価する」とは絶対に言わない', indent=True)
P('③ 代わりに「振り返りを丁寧に書こう」「ノートに気付きを書こう」と促す', indent=True)
P('④ 発言が少ない生徒の「ノート・振り返り」を必ず読む', indent=True)
P('⑤ 教科会で「発言量を見るのではない」ことを共通理解する', indent=True)

H2('質問への回答：「行動観察」は教師の主観が入るのではないか？')

H3('回答：ご指摘は正しい。だからこそ「主観を排除する4つの仕組み」が必要です')

QUOTE('評価の妥当性や信頼性を高めるとともに、児童生徒に学習の見通しをもたせるため、'
      '学習評価の方針を事前に児童生徒と共有する場面を必要に応じて設けたり、'
      '評価規準等を組織的に共有することなどが大切である。',
      '中央教育審議会報告（平成31年1月21日）')

H3('主観を排除する4つの仕組み')

P('【仕組み①】事前の評価規準の明確化', bold=True, color=NAVY)
P('・「どんな姿が見えたらA／B／C」を、観察の前に教科会で言語化', indent=True)
P('・「印象」で判定せず、規準と照合する', indent=True)

P('【仕組み②】複数の材料を組み合わせる（公的指針）', bold=True, color=NAVY)
P('・行動観察だけで決めない。ノート・レポート・自己評価と組み合わせる', indent=True)
P('・1つの観察結果が外れていても、他の材料と照合することで主観の影響を相殺', indent=True)

P('【仕組み③】ルーブリックの活用', bold=True, color=NAVY)
P('・観察ポイントを「観察可能な具体的行動」として記述化', indent=True)
P('・例：「前時のノートを見返している」「方法を変えて試している」など', indent=True)
P('・複数の教師で同じ生徒を観察して、判定が一致するかチェック', indent=True)

P('【仕組み④】教科会での事例共有・摺り合わせ（モデレーション）', bold=True, color=NAVY)
P('・「この生徒のこの記述／行動は、A／Bどちらか」を持ち寄って議論', indent=True)
P('・教師間の判定の癖を顕在化させて、組織として揃える', indent=True)

CAUTION('「行動観察に主観が入る」のは事実です。これを否定するのは不誠実です。'
        'しかし、複数の材料を組み合わせ、教科会で規準を揃え、'
        'ルーブリックで観察ポイントを明確化することで、'
        '「主観の影響を最小化する」ことができます。')

doc.add_page_break()

# ============================================================
# 第Ⅵ部　先生方への「返しのセリフ」例
# ============================================================
H1('第Ⅵ部　先生方への「返しのセリフ」例')

H2('質問への返し方の骨子 ── 今日の質疑応答でそのまま使える形')

P('（以下は、質疑応答や教科会で、そのまま声に出せるよう整理した返しの例です）', size=10, color=GRAY)
P('')

CAUTION('国・都が「AAC・CCAはありえない」と言うのは、'
        '「付けるな」という意味ではなく、'
        '「単元末に出たら、評価方法か指導が機能していない赤信号として扱い、'
        '原因を点検して改善せよ」という意味です（国研の参考資料の言葉どおり）。')

P('')
P('理由は、態度が ──', bold=True)
P('「性格や努力量」ではなく ──', bold=True)
P('「知識・思考を生み出す同じ学びの意思的な側面（粘り強さ・自己調整）」', bold=True, color=ORANGE)
P('だからです。', bold=True)
P('')
P('だから知識・思考と極端にズレたら、たいていは態度を', bold=True)
P('挙手や提出など「学習と切り離した形式面」で測ってしまっているサイン、ということになります。', bold=True)
P('')

P('実際にそう見える生徒は確かにいます。', bold=True, color=NAVY)
P('でも公式も「途中のばらつきはある」と認めていて、', bold=True, color=NAVY)
P('大事なのは ── 単元の途中で見つけて、指導で閉じることです。', bold=True, color=ORANGE)
P('')
P('・できる子が退屈してC → 課題が易しすぎるサイン', indent=True)
P('・懸命なのに伸びずCC → 自己調整を支える指導が届いていないサイン', indent=True)
P('どちらも、生徒ではなく、こちらの課題設計・指導へのフィードバックとして読みます。', bold=True)

P('')
P('しかも国は2025年に、まさにこの難しさを理由に、', bold=True)
P('次期指導要領で態度を評定から外し、個人内評価にする方針を出しました。', bold=True, color=ORANGE)
P('皆さんの違和感は、制度的にも正しかった、ということです。', bold=True, color=ORANGE)

doc.add_page_break()

# ============================================================
# 第Ⅶ部　明日からの教科会で取り組む5点
# ============================================================
H1('第Ⅶ部　明日からの教科会で取り組む5点')

P('')
P('① 主体的態度の「Ａ／Ｂ／Ｃの姿」を、教科会で具体的に言語化する', size=12, bold=True, color=NAVY)
P('　 → 「真面目さ・発言量・提出回数」ではなく', size=10.5, indent=True)
P('　 　「学習目標に向けた粘り強さ・自己調整の姿」で記述する', size=10.5, indent=True)
P('')
P('② AAC・CCAが出ている生徒について、教科会で個別ケース検討', size=12, bold=True, color=NAVY)
P('　 → 4つに切り分ける：', size=10.5, indent=True)
P('　 　形式要素の混入か／一時的なものか／単元途中の状態か／規準のズレか', size=10.5, indent=True)
P('')
P('③ 行動観察用のミニルーブリック（観察ポイント表）を教科会で作成', size=12, bold=True, color=NAVY)
P('　 → 例：「前時のノートを見返している」「方法を変えて試している」など', size=10.5, indent=True)
P('　 　「観察可能な具体的行動」で記述する', size=10.5, indent=True)
P('')
P('④ 「発言を評価する」とは生徒に言わない／教科会で共通理解', size=12, bold=True, color=NAVY)
P('　 → 代わりに「振り返りを丁寧に書こう」「ノートに気付きを書こう」と促す', size=10.5, indent=True)
P('')
P('⑤ 単元途中（形成的評価の段階）で態度C傾向の生徒を早期発見し、指導で閉じる', size=12, bold=True, color=NAVY)
P('　 → 単元末まで放置しない', size=10.5, indent=True)
P('　 → 「指導で閉じる」が指導と評価の一体化の本旨', size=10.5, indent=True)

HR()

H2('出典（公的資料）')
P('● 中央教育審議会「児童生徒の学習評価の在り方について（報告）」', size=10)
P('　 平成31年1月21日　文部科学省', size=10, color=GRAY)
P('● 国立教育政策研究所「『指導と評価の一体化』のための学習評価に関する参考資料」', size=10, bold=True)
P('　 第1編 総説（各教科共通）／第2編（各教科別）　令和2年3月', size=10, color=GRAY, bold=True)
P('● 中学校学習指導要領 解説 総則編　文部科学省（平成29年7月）', size=10)
P('　 第3章第3節2「学習評価の充実」', size=10, color=GRAY)
P('● 東京都教育委員会「子供たちに未来の創り手となるために必要な資質・能力を育む', size=10)
P('　 指導と評価の一体化を目指して」理論編・実践編（令和2年9月）', size=10, color=GRAY)
P('● 学習評価の在り方ハンドブック（小・中学校編）', size=10)
P('　 文部科学省 国立教育政策研究所教育課程研究センター（令和元年6月）', size=10, color=GRAY)
P('● 「主体的に学習に取り組む態度」の評価に関するガイドライン', size=10)
P('　 奈良県教育委員会事務局 学ぶ力はぐくみ課', size=10, color=GRAY)
P('● 中央教育審議会 教育課程企画特別部会「論点整理」', size=10)
P('　 令和7年9月25日（次期学習指導要領の方向性）', size=10, color=GRAY)

P('')
P('※ 練馬区独自の文書はオンラインで確認できませんでした。', size=10, color=GRAY)
P('　 練馬区は国・都の枠組みを実施する立場のため、上記が実質的な根拠になります。', size=10, color=GRAY)

HR()
P('以上、本日の研修当日にいただいた3つの質問への、', color=GRAY)
P('公的資料に基づく、誠実な回答です。', color=GRAY)
P('続きはいつでもお気軽にお問い合わせください。 ─ 指導主事 紺多', color=GRAY)

import os
os.makedirs('/tmp/output', exist_ok=True)
out = '/tmp/output/補足説明_AACCCA_行動観察_v2_石神井南中_20260603.docx'
doc.save(out)
print('saved:', out)
print('approx chars:', sum(len(p.text) for p in doc.paragraphs))
