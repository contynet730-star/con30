// 令和８年度 教育指導課訪問（練馬区立みらい青空学園）５校時 指導・講評　発表原稿【プランＢ】
// 生成: node make_genkou_planB.js
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType, PageBreak,
} = require("docx");
const fs = require("fs");

const FONT = "ＭＳ 明朝";
const FONT_H = "ＭＳ ゴシック";
const TEAL = "2F8F79"; // 資料の見出し色（印刷を考慮して濃度を上げたもの）
const GRAY = "595959";

/* ---------- パーツ ---------- */

// 章見出し（スライド番号・経過時間つき）
function head(slideLabel, time, title) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "BFDED5" } },
    children: [
      new TextRun({ text: slideLabel, font: FONT_H, size: 20, bold: true, color: TEAL }),
      new TextRun({ text: "　" + title, font: FONT_H, size: 22, bold: true, color: "1A1A1A" }),
      new TextRun({ text: "　　" + time, font: FONT_H, size: 18, color: GRAY }),
    ],
  });
}

// 読み上げる本文（１文＝１段落。棒読みにならないよう改行位置で息継ぎを示す）
function say(text) {
  return new Paragraph({
    spacing: { after: 60, line: 340 },
    children: [new TextRun({ text, font: FONT, size: 21 })],
  });
}

// ト書き（読み上げない指示）
function cue(text) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    indent: { left: 240 },
    children: [new TextRun({ text: "（" + text + "）", font: FONT_H, size: 18, color: GRAY })],
  });
}

function small(text) {
  return new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text, font: FONT_H, size: 18, color: GRAY })],
  });
}

/* ---------- 冒頭の表 ---------- */
function infoTable(rows) {
  return new Table({
    columnWidths: [1800, 7200],
    width: { size: 9000, type: WidthType.DXA },
    rows: rows.map(
      ([k, v]) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: 1800, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: "EAF6F2" },
              children: [
                new Paragraph({
                  spacing: { before: 40, after: 40 },
                  children: [new TextRun({ text: k, font: FONT_H, size: 19, bold: true })],
                }),
              ],
            }),
            new TableCell({
              width: { size: 7200, type: WidthType.DXA },
              children: [
                new Paragraph({
                  spacing: { before: 40, after: 40 },
                  children: [new TextRun({ text: v, font: FONT_H, size: 19 })],
                }),
              ],
            }),
          ],
        })
    ),
  });
}

/* ---------- 本文 ---------- */
const body = [];

body.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    children: [
      new TextRun({ text: "令和８年度　教育指導課訪問　５校時　指導・講評", font: FONT_H, size: 28, bold: true }),
    ],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 240 },
    children: [new TextRun({ text: "発表原稿【プランＢ】（１０分）", font: FONT_H, size: 22, color: TEAL, bold: true })],
  })
);

body.push(
  infoTable([
    ["日　時", "令和８年９月９日（水）５校時　13:20〜14:10"],
    ["訪問先", "練馬区立みらい青空学園（施設一体型小中一貫教育校・令和８年４月開校）"],
    ["参観授業", "塚本　瑞穂　先生（外国語・第７学年Ｂ組）／田口　和磨　先生（保健体育・第９学年ＡＢ組）／東海林　静江　先生（道徳・特別支援学級Ｄ組）"],
    ["講評者", "練馬区教育委員会　指導主事　紺多　章一郎"],
    ["構　成", "①アイスブレイク・現状／②３人の授業から／③今後に向けて（全15枚）"],
  ])
);

body.push(
  new Paragraph({ spacing: { before: 200, after: 60 }, children: [new TextRun({ text: "キーメッセージ", font: FONT_H, size: 22, bold: true, color: TEAL })] }),
  new Paragraph({
    spacing: { after: 60 },
    indent: { left: 240 },
    children: [new TextRun({ text: "「違い」を生かす授業が、９年間の学びをつなぐ。", font: FONT_H, size: 26, bold: true })],
  }),
  small("　表紙・スライド２・まとめの３か所で提示する。講評全体の目次を兼ねた一文。"),
  new Paragraph({ spacing: { before: 140, after: 60 }, children: [new TextRun({ text: "到達点としての一文", font: FONT_H, size: 22, bold: true, color: TEAL })] }),
  new Paragraph({
    spacing: { after: 60 },
    indent: { left: 240 },
    children: [new TextRun({ text: "教師が管理する学習から、子どもが管理する学習へ。", font: FONT_H, size: 24, bold: true })],
  }),
  small("　スライド12の締め。読み上げたあと、ひと呼吸おく。")
);

body.push(
  new Paragraph({ spacing: { before: 200, after: 60 }, children: [new TextRun({ text: "読み上げにあたって", font: FONT_H, size: 20, bold: true, color: TEAL })] }),
  small("・全文で約３，０００字。ふつうの速さで読んで、クイズの間を含め約１０分です。"),
  small("・見出しの右の時刻は、そこを読み始める目安です。２０秒以上ずれたら、太字以外を落として調整します。"),
  small("・（　）はト書きです。読み上げません。"),
  small("・数値は言い切らず、ひと呼吸おいてから言うと伝わります。"),
  small("・スライド７〜９では、写真に手で示しながら話すと伝わりやすくなります。")
);

body.push(new Paragraph({ children: [new PageBreak()] }));

/* ===== 導入 ===== */
body.push(head("スライド１", "0:00", "表紙"));
body.push(say("本日は、貴重な授業を参観させていただき、誠にありがとうございました。"));
body.push(say("練馬区教育委員会　指導課の紺多と申します。"));
body.push(say("授業を公開してくださった先生方に、まず御礼を申し上げます。"));
body.push(say("これから１０分間、お時間をいただきます。"));
body.push(say("本日申し上げたいことは、この一文に尽きます。「違い」を生かす授業が、９年間の学びをつなぐ。"));

body.push(head("スライド２", "0:29", "本日の講評"));
body.push(say("お話は三点です。"));
body.push(say("はじめにクイズと、みらい青空学園の現状。次に、３人の先生の授業から。最後に、今後に向けて。"));
body.push(say("３人の先生の指導案を手がかりに、本日の授業を読み解いてまいります。"));

/* ===== ① アイスブレイク・現状 ===== */
body.push(head("スライド３", "0:47", "①　アイスブレイク　クイズ"));
body.push(say("令和８年度　全国学力・学習状況調査の質問紙調査から、二問です。"));
body.push(say("全国では、小学校から中学校にかけて下がる項目が多くあります。「英語の勉強は好きですか」は、全国で１５．５ポイント下がります。"));
body.push(say("では、本校ではどうだったでしょうか。"));
body.push(say("第一問。「英語の勉強は好きですか」。本校の小学部６年は５１．８パーセントです。中学部９年は何パーセントでしょうか。"));
body.push(say("①　約４２パーセント。②　約５２パーセント。③　約６２パーセント。"));
body.push(say("第二問。「自分には、よいところがあると思いますか」。本校の小学部６年は８５．２パーセント。中学部９年はどうでしょうか。"));
body.push(say("①　約８５。②　約８８。③　約９２。"));
body.push(cue("数秒待つ。挙手を求めるか、近くの先生同士で一言交わしていただく"));

body.push(head("スライド４", "1:54", "①　アイスブレイク　答え"));
body.push(say("答えは、いずれも③です。"));
body.push(say("「英語の勉強は好きですか」は、５１．８から６１．７パーセントへ。９．９ポイント上がっています。全国が１５．５ポイント下がる項目です。"));
body.push(say("小学部では全国を下回っていたものが、中学部では全国を１０ポイント以上、上回りました。"));
body.push(say("「自分には、よいところがある」は、８５．２から９１．５パーセントへ。６．３ポイント上がっています。"));
body.push(say("全国も東京都も下がる項目で、本校は上がっている。これが、本日の出発点です。"));

body.push(head("スライド５", "2:35", "①　みらい青空学園の現状"));
body.push(say("対話と協働に関わる項目が、小学部から中学部にかけて伸びています。"));
body.push(say("「友達や周りの人の考えを大切にして、協力しながら課題の解決に取り組んでいますか」は、８５．１から１００パーセントへ。中学部は、全員が肯定的に答えています。"));
body.push(say("「話し合う活動を通じて、考えを深められている」は、９２．５から９７．９パーセントへ。"));
body.push(say("「自分と違う意見について考えるのは楽しい」は、８１．４から８７．２パーセントへ。"));
body.push(say("いずれも中学部は、練馬区や全国を上回っています。対話し、協働する力が、９年間を通じて確かに育っています。"));

/* ===== ② ３人の授業から ===== */
body.push(head("スライド６", "3:24", "②　３つの指導案に共通して書かれていたこと"));
body.push(say("では、なぜ育っているのか。３人の先生の指導案を読ませていただきました。"));
body.push(say("学級はまったく違います。水泳は泳力差が顕著、外国語は習熟度差が大きい、道徳は特別支援学級。"));
body.push(say("３つの学級は、いずれも一人一人の差が大きい集団です。"));
body.push(say("その差にどう向き合うか。指導案には、共通する手だてが書かれていました。"));
body.push(say("視点Ⅰ　見通しをもたせる。視点Ⅱ　違いに応じる。視点Ⅲ　対話で考えを深める。"));
body.push(say("この３つを手がかりに、３人の授業を見てまいります。"));

body.push(head("スライド７", "4:05", "②　田口　和磨　先生"));
body.push(say("田口先生の授業です。保健体育、第９学年ＡＢ組。水泳の第３時、バタフライのキックでした。"));
body.push(say("指導案には、コロナ禍と校舎改築で水泳の授業回数が少なく、苦手意識をもつ生徒とスイミングスクール経験者が混在し、泳力差が顕著、と書かれていました。"));
body.push(say("そのうえで、着替えの前に本時の目標と流れを確認する。泳力別に３段階でコースを分け、補助具の使用を認める。得意な生徒にアドバイス役を担わせる。バディで互いのキックを確認し合う。"));
body.push(say("参観して印象的だったのは、確認の視点を三点、明確に示されていたことです。子どもが自分の泳ぎを見る目をもてていました。"));

body.push(head("スライド８", "4:57", "②　東海林　静江　先生"));
body.push(say("東海林先生の授業です。道徳、特別支援学級Ｄ組。「ついていい嘘とついてはいけない嘘はどう違うのか」でした。"));
body.push(say("指導案には、将来の就労を見据えた「報告・連絡・相談」を大切にしている学級であること、そして人狼ゲームをめぐる生徒の発言が出発点だと書かれていました。"));
body.push(say("前時と２週続きで組み立てる。前時のワークシートを活用し、考えが変わってもよいと保障する。自己との対話で整理してから、グループで交流し議論へ。発表よりも議論の時間を重視する。"));
body.push(say("参観して、全員が自分の考えをもって議論に入れていました。"));

body.push(head("スライド９", "5:45", "②　塚本　瑞穂　先生"));
body.push(say("塚本先生の授業です。外国語、第７学年Ｂ組。Unit 4 の第２時でした。"));
body.push(say("指導案には、小学校時から苦手意識のある生徒、母語が英語の生徒、英検２級レベルの生徒まで混在し、習熟度の差が大きい、と書かれていました。"));
body.push(say("Today’s Goal と Plan を提示する。単元末はＡＬＴへの発表で、目的と相手が明確なゴールを置く。発表の前にペアで共有し、くじで指名して、１時間で全員に発話の機会をつくる。"));
body.push(say("参観して、板書を手がかりに、誰もが一文をつくれるようにされていました。"));

body.push(head("スライド１０", "6:31", "②　３つの授業に共通していたこと"));
body.push(say("３つの授業を並べると、このようになります。"));
body.push(say("見通しは、着替えの前の確認、２週続きの構成、Today’s Goal。"));
body.push(say("違いに応じる手だては、泳力別のコース、前時のワークシート、発表前のペア共有。"));
body.push(say("対話は、バディでの確認、自己内対話のあとの議論、ペアから全体へ。"));
body.push(say("教科も学年も違いますが、３人とも、差を埋めるのではなく、差を前提に授業を設計されていました。"));

body.push(head("スライド１１", "7:05", "②　この手だては、子どもに何を育てるか"));
body.push(say("では、この手だては、子どもに何を育てるのでしょうか。二つあると考えます。"));
body.push(say("一つは自己肯定感です。「できる子になる」ことではなく、「成長できる自分を信じること」。自分の段階で挑戦でき、考えが変わってもよいと保障されることで育ちます。"));
body.push(say("もう一つが、学習の自己調整です。学習指導要領が示す「学びに向かう力」の中核をなす力です。"));

body.push(head("スライド１２", "7:37", "②　学習の自己調整が育っていた場面"));
body.push(say("「主体的に学習に取り組む態度」は、粘り強く取り組む側面と、自らの学習を調整しようとする側面の二つで捉えます。"));
body.push(say("田口先生は、着替えの前に目標と流れを確認され、学習カードで要点を振り返らせていました。"));
body.push(say("塚本先生は、Today’s Goal と Plan を示し、振り返りをワークシートに記入させていました。"));
body.push(say("東海林先生は、議論に入る前に前時のワークシートで自分の考えを確かめさせていました。"));
body.push(say("いずれも、目標と振り返りが一対で置かれています。"));
body.push(say("教師が管理する学習から、子どもが管理する学習へ。"));
body.push(cue("ひと呼吸おく"));
body.push(say("本日の授業から見えた、いちばん大きな変化です。"));

/* ===== ③ 今後に向けて ===== */
body.push(head("スライド１３", "8:32", "③　この３つを、９年間の共通の言葉に"));
body.push(say("今後に向けてです。"));
body.push(say("本日見えた３つの視点は、学年や教科を越えて使えます。"));
body.push(say("学校経営計画に掲げられた「主体的・対話的で深い学びの実現に向けた授業改善」「個に応じたきめ細やかな指導・支援の充実」とも重なります。"));
body.push(say("学年ごと、教科ごとに閉じない。同じ言葉で語れることが、９年間をつなぎます。"));

body.push(head("スライド１４", "9:00", "③　その言葉で、互いの授業を見合う"));
body.push(say("そして、もう一点。"));
body.push(say("本日の３つの授業は、保健体育の９年、道徳のＤ組、外国語の７年。教科も学年も違います。それでも、同じ３つの視点で語ることができました。"));
body.push(say("だから、教科を越えて見合えます。"));
body.push(say("小学部の教員が中学部を、中学部の教員が小学部を見る。９年後の子どもの姿と、９年前の子どもの姿を、同じ校舎で見られる学校です。"));
body.push(say("学校経営計画の「乗り入れ授業」「相互の実践に生かす」を、この３つの観点で動かしていただきたい。"));
body.push(cue("時間に余裕があれば添える　子どもが学習を調整するように、教師も授業を調整する。その具体が、互いに見合うこと"));

body.push(head("スライド１５", "9:40", "まとめ"));
body.push(say("まとめます。"));
body.push(say("３人とも、一人一人の差を学びの出発点にされていました。"));
body.push(say("見通し、違いに応じること、対話。その先に育つのが、自己肯定感と、学習の自己調整です。"));
body.push(say("「違い」を生かす授業が、９年間の学びをつなぐ。"));
body.push(say("９年間を見通した学びの充実に期待しております。本日は、誠にありがとうございました。"));
body.push(cue("一礼して終了"));

/* ===== 巻末 ===== */
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(
  new Paragraph({
    spacing: { after: 160 },
    children: [new TextRun({ text: "巻末　読み上げる数値の一覧", font: FONT_H, size: 24, bold: true, color: TEAL })],
  }),
  small("出典：令和８年度　全国学力・学習状況調査　質問紙調査　肯定的回答の割合。本校は回答結果集計表による。")
);

function dataTable(header, rows) {
  const widths = [3600, 1800, 1800, 1800];
  const mk = (cells, isHead) =>
    new TableRow({
      children: cells.map(
        (c, i) =>
          new TableCell({
            width: { size: widths[i], type: WidthType.DXA },
            shading: isHead ? { type: ShadingType.CLEAR, fill: "EAF6F2" } : undefined,
            children: [
              new Paragraph({
                alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
                spacing: { before: 40, after: 40 },
                children: [new TextRun({ text: c, font: FONT_H, size: 18, bold: !!isHead })],
              }),
            ],
          })
      ),
    });
  return new Table({
    columnWidths: widths,
    width: { size: 9000, type: WidthType.DXA },
    rows: [mk(header, true), ...rows.map((r) => mk(r, false))],
  });
}

body.push(
  new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: "本校の小学部→中学部（スライド３〜５）", font: FONT_H, size: 20, bold: true })] })
);
body.push(
  dataTable(["質問項目", "小学部６年", "中学部９年", "差"], [
    ["英語の勉強は好きですか", "51.8", "61.7", "＋9.9"],
    ["自分には、よいところがあると思いますか", "85.2", "91.5", "＋6.3"],
    ["友達や周りの人の考えを大切にして、協力しながら課題の解決に取り組んでいる", "85.1", "100.0", "＋14.9"],
    ["話し合う活動を通じて、考えを深めたり新たな考え方に気付いたりできている", "92.5", "97.9", "＋5.4"],
    ["自分と違う意見について考えるのは楽しい", "81.4", "87.2", "＋5.8"],
  ])
);
body.push(small("※　参考（全国）：英語の勉強は好きですか　小66.7→中51.2（−15.5）／自分にはよいところがある　小85.6→中84.0（−1.6）"));
body.push(small("※　中学部の参照値：協力しながら課題解決　全国92.4／話し合う活動　練馬区86.6／違う意見　練馬区79.8"));
body.push(small("※　本校 小学部27名・中学部47名（令和８年４月実施）。"));

body.push(
  new Paragraph({ spacing: { before: 240, after: 80 }, children: [new TextRun({ text: "３人の手だて（スライド10の一覧）", font: FONT_H, size: 20, bold: true })] })
);
body.push(
  dataTable(["視点", "田口　先生", "東海林　先生", "塚本　先生"], [
    ["Ⅰ　見通し", "着替えの前に目標と流れを確認", "前時と２週続きで組み立てる", "Today’s Goal と Plan の提示"],
    ["Ⅱ　違いに応じる", "泳力別の３コース／補助具", "前時のワークシート／意見が変わってよい", "発表前のペア共有／くじ指名"],
    ["Ⅲ　対話", "バディで互いのキックを確認", "自己内対話のあと議論を重視", "ペアで確かめてから全体へ"],
  ])
);

/* ---------- 出力 ---------- */
const doc = new Document({
  creator: "練馬区教育委員会",
  title: "令和８年度 教育指導課訪問 ５校時 指導・講評 発表原稿（プランＢ）",
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [
    {
      properties: { page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
      children: body,
    },
  ],
});

const OUT = process.argv[2] || "発表原稿_５校時指導・講評_みらい青空学園_プランB.docx";
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote " + OUT);
});
