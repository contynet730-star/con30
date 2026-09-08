// 令和８年度 教育指導課訪問（練馬区立みらい青空学園）５校時 指導・講評　発表原稿
// 生成: node make_genkou.js  →  発表原稿_５校時指導・講評_みらい青空学園.docx
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
    children: [new TextRun({ text: "発表原稿（１０分）", font: FONT_H, size: 22, color: TEAL, bold: true })],
  })
);

body.push(
  infoTable([
    ["日　時", "令和８年９月９日（水）５校時　13:20〜14:10"],
    ["訪問先", "練馬区立みらい青空学園（施設一体型小中一貫教育校・令和８年４月開校）"],
    ["参観授業", "塚本　瑞穂　先生（外国語・第７学年Ｂ組）／田口　和磨　先生（保健体育・第９学年ＡＢ組）／東海林　静江　先生（道徳・特別支援学級Ｄ組）"],
    ["講評者", "練馬区教育委員会　指導主事　紺多　章一郎"],
    ["構　成", "０　アイスブレイク／１　練馬区の小中一貫教育／２　本時の授業について"],
  ])
);

body.push(
  new Paragraph({ spacing: { before: 240, after: 60 }, children: [new TextRun({ text: "読み上げにあたって", font: FONT_H, size: 20, bold: true, color: TEAL })] }),
  small("・全文で約３，０００字。ふつうの速さで読んでちょうど１０分です。クイズの間を含め約10分10秒。"),
  small("・見出しの右の時刻は、そこを読み始める目安です。２０秒以上ずれたら、太字以外を落として調整します。"),
  small("・（　）はト書きです。読み上げません。"),
  small("・数値は言い切らず、ひと呼吸おいてから言うと伝わります。"),
  small("・スライド９の本校の数値は、当日確定した値に読み替えてください。")
);

body.push(new Paragraph({ children: [new PageBreak()] }));

/* ===== 導入 ===== */
body.push(head("スライド１", "0:00", "表紙"));
body.push(say("本日は、貴重な授業を参観させていただき、誠にありがとうございました。"));
body.push(say("練馬区教育委員会　指導課の紺多と申します。"));
body.push(say("授業を公開してくださった先生方に、まず御礼を申し上げます。"));
body.push(say("これから１０分間、お時間をいただきます。"));

body.push(head("スライド２", "0:20", "本日の講評"));
body.push(say("本日は、二点に絞ってお話しします。"));
body.push(say("一点目は、練馬区の小中一貫教育について。"));
body.push(say("二点目は、本時の授業について。主体的・対話的で深い学びの視点からの授業改善です。"));
body.push(say("その前に、少しだけクイズにお付き合いください。"));

/* ===== ０ アイスブレイク ===== */
body.push(head("スライド３", "0:40", "０　アイスブレイク　クイズ"));
body.push(say("令和８年度の全国学力・学習状況調査、質問紙調査からの出題です。"));
body.push(say("全国では、小学校から中学校にかけて下がる項目が多くあります。「英語の勉強は好きですか」は、全国で１５．５ポイント下がります。"));
body.push(say("では、本校ではどうだったでしょうか。"));
body.push(say("第一問。「英語の勉強は好きですか」。本校の小学部６年は５１．８パーセントです。中学部９年は何パーセントでしょうか。"));
body.push(say("①　約４２パーセント。②　約５２パーセント。③　約６２パーセント。"));
body.push(say("第二問。「自分には、よいところがあると思いますか」。本校の小学部６年は８５．２パーセント。中学部９年はどうでしょうか。"));
body.push(say("①　約８５。②　約８８。③　約９２。"));
body.push(cue("数秒待つ。挙手を求めるか、近くの先生同士で一言交わしていただく"));

body.push(head("スライド４", "1:48", "０　アイスブレイク　答え"));
body.push(say("答えは、いずれも③です。"));
body.push(say("「英語の勉強は好きですか」は、５１．８から６１．７パーセントへ。９．９ポイント上がっています。"));
body.push(say("全国が１５．５ポイント下がる項目です。"));
body.push(say("小学部では全国を下回っていたものが、中学部では全国を１０ポイント以上、上回りました。"));
body.push(say("「自分には、よいところがある」は、８５．２から９１．５パーセントへ。６．３ポイント上がっています。全国は１．６ポイント下がります。"));
body.push(say("全国も東京都も下がる項目で、本校は上がっている。"));
body.push(say("これが、本日の出発点です。"));

body.push(head("スライド５", "2:32", "０　アイスブレイク　これが小中一貫教育校の強み"));
body.push(say("上がっているのは、この二項目だけではありません。"));
body.push(say("「友達や周りの人の考えを大切にして、協力しながら課題の解決に取り組んでいますか」は、８５．１から１００パーセントへ。中学部は、全員が肯定的に答えています。"));
body.push(say("「話し合う活動を通じて、自分の考えを深めたり、新たな考え方に気付いたりすることができていますか」は、９２．５から９７．９パーセントへ。"));
body.push(say("「自分と違う意見について考えるのは楽しい」は、８１．４から８７．２パーセントへ。"));
body.push(say("９年間を見通した教育の成果が、数値に表れています。ここから本題に入ります。"));

/* ===== １ 練馬区の小中一貫教育 ===== */
body.push(head("スライド６", "3:21", "１　練馬区の小中一貫教育"));
body.push(say("練馬区教育委員会は、「夢や希望をもち、困難を乗り越える力を備えた子どもたちの育成」を目標に掲げています。"));
body.push(say("その実現のための施策の一つが、小学校６年間と中学校３年間を合わせた、９年間を見通した教育です。"));
body.push(say("ねらいは、学力・体力の向上、豊かな人間性・社会性の育成、そして滑らかな接続による安定した学校生活の実現です。"));

body.push(head("スライド７", "3:52", "１　施設一体型の強み"));
body.push(say("施設一体型では、この効果がさらに高まります。教員間の連携強化、異学年交流の活性化、そして小中学校間の指導の統一化です。"));
body.push(say("本校は、区内２校目の施設一体型として、本年４月に開校しました。"));
body.push(say("１年生から９年生までが同じ学び舎で学ぶ強みを生かし、「目指す１５歳の姿」を９年間で描いていくことに期待しています。"));

/* ===== 学力調査から見える強み ===== */
body.push(head("スライド８", "4:21", "学力調査結果から見える強み"));
body.push(say("学力調査結果から見える子どもの姿を、もう少し見てまいります。"));
body.push(say("本校では、学習規律、学習習慣、そして学びへの主体性が、着実に育っています。"));

body.push(head("スライド９", "4:35", "本校と練馬区の比較"));
body.push(say("次に、練馬区平均との比較です。"));
body.push(say("本日お話しする三つの視点に対応する質問項目を並べました。"));
body.push(say("自己肯定感にあたる項目は、中学部で練馬区を８．２ポイント上回っています。"));
body.push(say("対話・協働は、小学部で５．５、中学部で１１．３ポイント。"));
body.push(say("学習の自己調整は、小学部で１３．６、中学部で１０．４ポイント上回っています。"));
body.push(say("六項目のうち五項目で、本校は練馬区平均を上回っています。"));
body.push(say("調査は一時点の結果ですが、その背景には、先生方の日々の授業の積み重ねがあると受け止めております。"));
body.push(cue("小学部の自己肯定感のみ、練馬区を1.2ポイント下回る。児童27名のため１人が3.7ポイントにあたることを踏まえ、深追いしない"));

body.push(head("スライド１０", "5:19", "なぜその成果が現れているのか"));
body.push(say("調査結果は、あくまで「結果」です。"));
body.push(say("大切なのは、その結果を生み出している要因です。"));
body.push(say("本日の三つの授業から、次の三つが見えてきました。"));
body.push(say("一つ、自己肯定感。二つ、対話・協働。三つ、学習の自己調整。順にお話しします。"));

/* ===== ① 自己肯定感 ===== */
body.push(head("スライド１１", "5:39", "２-①　自己肯定感"));
body.push(say("一点目は、自己肯定感です。"));
body.push(say("これからの時代、学力向上の基盤となるのは自己肯定感だと考えています。"));
body.push(say("ただし、ここでいう自己肯定感は、「できる子になる」ことではありません。"));
body.push(say("「成長できる自分を信じること」です。"));
body.push(say("この違いが、授業のつくり方を変えます。"));

body.push(head("スライド１２", "6:03", "２-①　自己肯定感を育む授業"));
body.push(say("田口先生の授業では、泳力差の大きい集団に対して、段階的な課題設定、泳力別のコース編成、バディでの確認活動が用意されていました。"));
body.push(say("キックの確認の視点を三点示されたことで、子どもが自分の泳ぎを見る目をもつことができていました。"));
body.push(say("東海林先生の授業では、安心して自分の考えを表現できる環境がつくられていました。意見が変わってもよいと保障されていた点が印象的でした。"));
body.push(say("塚本先生の授業では、発表の前にペアで共有する時間が確保され、全員に発話の機会が用意されていました。"));
body.push(say("いずれも、子どもが「認められる」よりも、「自分で成長を実感する」経験を積み重ねる授業でした。"));

/* ===== ② 対話・協働 ===== */
body.push(head("スライド１３", "6:57", "２-②　対話・協働"));
body.push(say("二点目は、対話・協働です。"));
body.push(say("調査結果に表れた主体的な学びの土台。その背景には、日常的に積み重ねられている「対話を通した学び」があると受け止めました。"));

body.push(head("スライド１４", "7:12", "２-②　対話によって学びは深まる"));
body.push(say("保健体育では、バディで互いの泳ぎを見合い、改善点を考えていました。得意な生徒が助言役を担っていました。"));
body.push(say("道徳では、「ついていい嘘と、ついてはいけない嘘」をテーマに、価値観の違いについて議論していました。発表よりも議論の時間を重視した構成でした。"));
body.push(say("外国語では、ペアでのやり取りを全体の学びへつないでいました。音読も相手とともに行い、表現を確かなものにしていました。"));
body.push(say("教科は異なりますが、共通していたのは「相手を通して自分を見つめる学び」です。これが対話・協働の本質だと考えます。"));

/* ===== ③ 学習の自己調整 ===== */
body.push(head("スライド１５", "7:59", "２-③　学習の自己調整"));
body.push(say("三点目は、学習の自己調整です。"));
body.push(say("これから求められる子どもは、教えられたことを学ぶ子どもではなく、自分で学び続ける子どもです。"));
body.push(say("そのために必要なのが、学習を自分で調整する力です。"));

body.push(head("スライド１６", "8:16", "２-③　自己調整が行われていた授業"));
body.push(say("本日の授業では、本時の目標、振り返り、自己評価、次時への見通し。この四点が大切にされていました。"));
body.push(say("田口先生は、着替えの前に目標と授業の流れを確認され、学習カードで要点を振り返らせていました。"));
body.push(say("塚本先生は、Today's Goal と Today's Plan を示し、振り返りをワークシートに記入させていました。"));
body.push(say("東海林先生は、前時のワークシートで自分の考えを確かめてから、議論に入っていました。"));
body.push(say("教師が管理する学習から、子どもが管理する学習へ。着実に転換が進んでいます。"));

/* ===== 今後・まとめ ===== */
body.push(head("スライド１７", "9:03", "今後に向けて"));
body.push(say("学力調査結果と、今日の授業をつなげて考えると、さらに伸ばしたいのは、「学びを自分事として捉える子ども」です。"));
body.push(say("そのためには、自己肯定感、対話・協働、学習の自己調整を、９年間で系統的に育てていくことが重要になります。"));
body.push(say("学年や教科で完結させず、９年間の系統として整理していただきたいと考えます。小竹小学校との校区別協議会も、その大切な場になります。"));

body.push(head("スライド１８", "9:37", "まとめ"));
body.push(say("まとめます。"));
body.push(say("学力向上は、知識の積み上げだけで実現するものではありません。"));
body.push(say("みらい青空学園では、自己肯定感、対話・協働、学習の自己調整が、着実に育成されています。"));
body.push(say("これこそが、施設一体型小中一貫教育校としての最大の強みであり、学力調査結果を支える土台です。"));
body.push(say("今後も、９年間を見通した学びの充実に期待しております。"));
body.push(say("本日は、誠にありがとうございました。"));
body.push(cue("一礼して終了　10:00"));

/* ===== 巻末：数値一覧 ===== */
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(
  new Paragraph({
    spacing: { after: 160 },
    children: [new TextRun({ text: "巻末　読み上げる数値の一覧", font: FONT_H, size: 24, bold: true, color: TEAL })],
  }),
  small("出典：令和８年度　全国学力・学習状況調査　質問紙調査　肯定的回答の割合。本校は回答結果集計表、練馬区は区の結果概要による。")
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
  new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: "アイスブレイク（スライド３〜５）　本校の小学部→中学部", font: FONT_H, size: 20, bold: true })] })
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
body.push(small("※　本校 小学部27名・中学部47名（令和８年４月実施）。"));

body.push(
  new Paragraph({ spacing: { before: 240, after: 80 }, children: [new TextRun({ text: "本校と練馬区の比較（スライド９）", font: FONT_H, size: 20, bold: true })] })
);
body.push(
  dataTable(["質問項目", "本校 小６／中３", "練馬区 小６／中３", "差"], [
    ["自分には、よいところがあると思いますか", "85.2 ／ 91.5", "86.4 ／ 83.3", "−1.2 ／ ＋8.2"],
    ["話し合う活動を通じて、考えを深められている", "92.5 ／ 97.9", "87.0 ／ 86.6", "＋5.5 ／ ＋11.3"],
    ["分かった点・分からない点を見直し、次につなげている", "92.6 ／ 89.4", "79.0 ／ 79.0", "＋13.6 ／ ＋10.4"],
  ])
);
body.push(small("※　６項目中５項目で本校が上回る。下回るのは小学部の自己肯定感のみ、差は1.2ポイント。"));

/* ---------- 出力 ---------- */
const doc = new Document({
  creator: "練馬区教育委員会",
  title: "令和８年度 教育指導課訪問 ５校時 指導・講評 発表原稿",
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [
    {
      properties: { page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
      children: body,
    },
  ],
});

const OUT = process.argv[2] || "発表原稿_５校時指導・講評_みらい青空学園.docx";
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote " + OUT);
});
