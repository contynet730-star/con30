// アツキのブリブリ・ワールドカップ 問題集 / 答えと解説 を Word で出力する
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeightRule, VerticalAlign,
  TableLayoutType, Footer, PageNumber,
} = require("docx");

const OUT_DIR = path.resolve(__dirname, "..");
const IMG = (n) => fs.readFileSync(path.join(__dirname, "img", n + ".png"));
const DAYS = [...require("./days1.js"), ...require("./days2.js")];

// ───────── 基本設定
const F_BODY = "UD デジタル 教科書体 N-R";   // Windows 10 以降に標準搭載（子ども向けの教科書体）
const F_HEAD = "BIZ UDPゴシック";            // Windows 10 以降に標準搭載
const font = (f) => ({ ascii: f, eastAsia: f, hAnsi: f, cs: f });
const W = 10200;                              // 本文の幅（DXA）
const BROWN = "6D4C41", BROWN_L = "EFEBE9";
const SUBJ = {
  国語: { c: "D32F2F", l: "FDECEA" },
  算数: { c: "1565C0", l: "E3F2FD" },
  理科: { c: "2E7D32", l: "E8F5E9" },
  社会: { c: "E65100", l: "FFF3E0" },
  英語: { c: "7B1FA2", l: "F3E5F5" },
};
const SZ = 22; // 11pt

const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NOB = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const TNOB = { ...NOB, insideHorizontal: NONE, insideVertical: NONE };
const line = (color, size = 8, style = BorderStyle.SINGLE) => ({ style, size, color });
const box = (color, size = 8, style = BorderStyle.SINGLE) =>
  ({ top: line(color, size, style), bottom: line(color, size, style), left: line(color, size, style), right: line(color, size, style) });
const shade = (fill) => ({ type: ShadingType.CLEAR, color: "auto", fill });

// ───────── テキスト（__下線__ / **太字** に対応）
function runs(text, o = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*|__[^_]+__)/).filter(Boolean);
  return parts.map((p) => {
    const r = { text: p, size: o.size || SZ, color: o.color, bold: o.bold, font: font(o.font || F_BODY) };
    if (p.startsWith("**")) { r.text = p.slice(2, -2); r.bold = true; }
    if (p.startsWith("__")) { r.text = p.slice(2, -2); r.underline = {}; }
    return new TextRun(r);
  });
}
function P(text, o = {}) {
  return new Paragraph({
    children: o.children || runs(text, o),
    alignment: o.align,
    spacing: { before: o.before || 0, after: o.after ?? 40, lineRule: "exact",
      // 行の高さを固定（フォントが変わってもページ割りが変わらないように）
      line: o.exact || Math.round(((o.size || SZ) / 2) * 20 * (o.line ? o.line / 210 : 1.36)) },
    indent: o.indent,
    keepNext: o.keepNext,
    keepLines: true,
    border: o.border,
    pageBreakBefore: o.pageBreak,
  });
}
const lines = (text, o = {}) => String(text).split("\n").map((t) => P(t, o));
const img = (n, w, h) => new ImageRun({ type: "png", data: IMG(n), transformation: { width: w, height: h || w } });
const spacer = (pt = 4) => new Paragraph({ spacing: { before: 0, after: 0, line: pt * 20, lineRule: "exact" }, children: [] });

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? shade(o.fill) : undefined,
    borders: o.borders || NOB,
    margins: o.margins || { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: o.valign || VerticalAlign.TOP,
    columnSpan: o.span,
  });
}
function table(cols, rows, o = {}) {
  return new Table({
    width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: o.borders || TNOB,
    alignment: o.align,
    rows: rows.map((r) => new TableRow({
      children: r.cells, cantSplit: true,
      height: r.h ? { value: r.h, rule: HeightRule.ATLEAST } : undefined,
    })),
  });
}
const oneBox = (children, o = {}) => table([o.w || W], [{ cells: [cell(children, { w: o.w || W, fill: o.fill, borders: o.borders, margins: o.margins })], h: o.h }], { align: o.align });

// ───────── 問題集のパーツ
function dayHeader(d) {
  const left = cell([
    P("DAY", { size: 20, bold: true, color: "FFFFFF", font: F_HEAD, align: AlignmentType.CENTER, after: 0, exact: 260 }),
    P(String(d.n), { size: 52, bold: true, color: "FFFFFF", font: F_HEAD, align: AlignmentType.CENTER, after: 0, exact: 640 }),
  ], { w: 1500, fill: BROWN, valign: VerticalAlign.CENTER });
  const mid = cell([
    ...(d.ch ? [P(d.ch, { size: 20, bold: true, color: "8D6E63", font: F_HEAD, after: 40, exact: 280 })] : []),
    P(d.t, { size: [...d.t].length > 15 ? 28 : 34, bold: true, font: F_HEAD, after: 0, exact: 460 }),
  ], { w: 5900, fill: BROWN_L, valign: VerticalAlign.CENTER, margins: { top: 60, bottom: 60, left: 200, right: 100 } });
  const right = cell([
    P("　　月　　日（　　）", { size: 20, after: 60 }),
    P("かかった時間　　　分", { size: 20, after: 0 }),
  ], { w: 2800, fill: BROWN_L, valign: VerticalAlign.CENTER });
  return table([1500, 5900, 2800], [{ cells: [left, mid, right], h: 900 }]);
}

function storyBox(d) {
  const head = new Paragraph({
    spacing: { after: 40 },
    children: [img("hakase", 26), new TextRun({ text: "　ストーリー", bold: true, size: 20, color: BROWN, font: font(F_HEAD) })],
  });
  return oneBox([head, ...d.story.map((s) => P(s, { size: SZ, after: 50, line: 290 }))], {
    fill: "FFF8E1", margins: { top: 80, bottom: 80, left: 200, right: 200 },
    borders: { ...NOB, left: line("A1887F", 36) },
  });
}

function choices(c, color) {
  const total = c.reduce((a, b) => a + b.length, 0);
  if (total <= 34) return [P(c.join("　　"), { indent: { left: 240 }, after: 30 })];
  return c.map((t) => P(t, { indent: { left: 240 }, after: 10 }));
}

function answerCell(w, color, h) {
  return cell([P("答え", { size: 15, color: "9E9E9E", after: 0, line: 200 })],
    { w, borders: box(color, 10), margins: { top: 30, bottom: 30, left: 80, right: 80 } });
}

function questionBlock(q, color) {
  const qParas = [...lines(q.q, { after: 30 }), ...(q.c ? choices(q.c, color) : [])];
  if (q.b === "S") {
    return [table([W - 2700, 140, 2560], [{
      cells: [cell(qParas, { w: W - 2700, margins: { top: 20, bottom: 20, left: 0, right: 60 } }),
        cell([P("")], { w: 140 }),
        answerCell(2560, color)],
      h: 620,
    }]), spacer(3)];
  }
  const h = q.b === "L" ? 1500 : 720;
  return [...qParas, table([W - 300], [{ cells: [answerCell(W - 300, color)], h }], { align: AlignmentType.RIGHT }), spacer(4)];
}

function problem(p, i) {
  const S = SUBJ[p.s];
  const head = new Paragraph({
    keepNext: true,
    spacing: { before: 120, after: 60 },
    border: { bottom: line(S.c, 8) },
    children: [
      new TextRun({ text: ` 問題${i + 1} `, bold: true, size: 22, color: "FFFFFF", font: font(F_HEAD), shading: shade("424242") }),
      new TextRun({ text: ` ${p.s} `, bold: true, size: 22, color: "FFFFFF", font: font(F_HEAD), shading: shade(S.c) }),
      new TextRun({ text: `　${p.t}`, bold: true, size: 24, color: S.c, font: font(F_HEAD) }),
    ],
  });
  const lead = p.lead.map((t) => P(t, { indent: { left: 120 }, after: 40, keepNext: true }));
  return [head, ...lead, ...p.q.flatMap((q) => questionBlock(q, S.c))];
}

function kanjiPK(d) {
  const head = new Paragraph({
    keepNext: true, spacing: { before: 160, after: 60 }, border: { bottom: line("263238", 8) },
    children: [
      new TextRun({ text: " ⚽ 漢字PK戦 ", bold: true, size: 24, color: "FFFFFF", font: font(F_HEAD), shading: shade("263238") }),
      new TextRun({ text: d.review ? "　【復習PK】今週の漢字をもう一度！" : "　【　】を漢字で書こう（書）／読みがなを書こう（読）", bold: true, size: 19, color: "263238", font: font(F_HEAD) }),
      new TextRun({ text: "　ゴール1本＝1点", size: 18, color: "616161", font: font(F_HEAD) }),
    ],
  });
  const colW = Math.floor(W / 3);
  const cells = d.pk.map((k, i) => {
    const kind = k.w === "書" ? "書き" : "読み";
    const sent = new Paragraph({
      spacing: { after: 80, line: 300 },
      children: [
        ...(k.pre ? runs(k.pre) : []),
        new TextRun({ text: "【" + k.y + "】", bold: true, size: SZ + 2, font: font(F_BODY), color: k.w === "書" ? "C62828" : "1565C0" }),
        ...(k.post ? runs(k.post) : []),
      ],
    });
    // 答えのマス
    const n = [...k.a].length;
    const sq = k.w === "書" ? 560 : 400;
    const grid = table(Array(n).fill(sq), [{ cells: Array(n).fill(0).map(() => cell([P("")], { w: sq, borders: box("78909C", 8) })), h: sq }],
      { align: AlignmentType.CENTER });
    return cell([
      P(`シュート${"①②③"[i]}（${kind}）`, { size: 18, bold: true, color: "455A64", font: F_HEAD, after: 40 }),
      sent, grid, spacer(2),
    ], { w: colW, borders: box("B0BEC5", 6), margins: { top: 80, bottom: 80, left: 120, right: 120 } });
  });
  return [head, table([colW, colW, W - 2 * colW], [{ cells }])];
}

function nextBox(d) {
  if (d.ending) {
    return oneBox([
      P("エンディング", { size: 22, bold: true, color: "F57F17", font: F_HEAD, after: 40 }),
      ...d.ending.map((t) => P(t, { after: 40 })),
    ], { fill: "FFFDE7", borders: box("F9A825", 16), margins: { top: 100, bottom: 100, left: 200, right: 200 } });
  }
  return oneBox([
    new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: "つづく…　", bold: true, size: 22, color: "37474F", font: font(F_HEAD) }), ...runs(d.next, { size: 21 })] }),
  ], { fill: "ECEFF1", borders: { ...NOB, left: line("37474F", 36) }, margins: { top: 80, bottom: 80, left: 200, right: 200 } });
}

function pointTable() {
  const heads = ["", "問題1", "問題2", "問題3", "問題4", "問題5", "漢字PK", "今日の合計", "これまでの合計", "おうちの人"];
  const cols = [1000, 760, 760, 760, 760, 760, 900, 1300, 1600, 1600];
  const hc = heads.map((h, i) => cell([P(h, { size: 17, bold: true, font: F_HEAD, align: AlignmentType.CENTER, after: 0, color: i ? "FFFFFF" : undefined })],
    { w: cols[i], fill: i ? BROWN : undefined, borders: box("8D6E63", 6), valign: VerticalAlign.CENTER, margins: { top: 30, bottom: 30, left: 30, right: 30 } }));
  const bc = heads.map((h, i) => cell([P(i === 0 ? "ポイント" : i === 6 ? "　／3" : i === 7 ? "　　／13" : "", { size: 17, bold: i === 0, font: F_HEAD, align: AlignmentType.CENTER, after: 0, color: "757575" })],
    { w: cols[i], borders: box("8D6E63", 6), valign: VerticalAlign.CENTER, fill: i === 0 ? BROWN_L : undefined }));
  return [
    table(cols, [{ cells: hc, h: 380 }, { cells: bc, h: 620 }]),
    P("◎一発で正解＝2点　○答えを見て直せた＝1点　漢字PK＝ゴール1本1点（1日の最高13点）　まちがえても、直せばポイントゲット！", { size: 16, color: "616161", after: 0, before: 30 }),
  ];
}

function trivia(d) {
  return new Paragraph({
    spacing: { before: 100, after: 0 },
    children: [img("unko", 22), new TextRun({ text: " うんこ豆知識　", bold: true, size: 19, color: BROWN, font: font(F_HEAD) }), ...runs(d.tri, { size: 19 })],
  });
}

function dayPages(d) {
  return [
    new Paragraph({ pageBreakBefore: true, spacing: { after: 0, line: 40, lineRule: "exact" }, children: [] }),
    dayHeader(d), spacer(6), storyBox(d),
    ...d.P.flatMap((p, i) => problem(p, i)),
    ...kanjiPK(d), spacer(8),
    nextBox(d),
    ...(d.chEnd ? [P("★　" + d.chEnd + "　★", { size: 26, bold: true, color: "F57F17", font: F_HEAD, align: AlignmentType.CENTER, before: 80, after: 40 })] : []),
    spacer(6), ...pointTable(), trivia(d),
  ];
}

// ───────── 表紙・使い方・登場人物・記録表・表彰状
const RANKS = [
  ["0〜39点", "うんこ見習い", "まずはここから！"],
  ["40〜89点", "うんこルーキー", "チームの一員になれたぞ"],
  ["90〜139点", "うんこレギュラー", "試合に出られるぞ"],
  ["140〜189点", "うんこエースストライカー", "チームの点取り屋だ！"],
  ["190〜239点", "うんこキャプテン", "みんなをひっぱるリーダー"],
  ["240〜260点", "伝説のうんこバロンドール", "世界一の選手！"],
];

function cover() {
  return [
    spacer(20),
    P("小学5年生　国語・算数・理科・社会・英語　20日間ミッション", { size: 22, bold: true, color: "8D6E63", font: F_HEAD, align: AlignmentType.CENTER, after: 200 }),
    P("アツキの", { size: 44, bold: true, font: F_HEAD, align: AlignmentType.CENTER, after: 80, exact: 600 }),
    P("ブリブリ・ワールドカップ", { size: 72, bold: true, color: BROWN, font: F_HEAD, align: AlignmentType.CENTER, after: 120, exact: 960 }),
    P("〜 うんこサッカー大ぼうけん問題集 〜", { exact: 440, size: 32, bold: true, color: "F57F17", font: F_HEAD, align: AlignmentType.CENTER, after: 200 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [img("cover", 620, 294)] }),
    oneBox([
      P("ミッション", { size: 24, bold: true, color: "C62828", font: F_HEAD, align: AlignmentType.CENTER, after: 60 }),
      P("ゲームの世界「ブリブリ・ワールド」にすいこまれたアツキ。", { align: AlignmentType.CENTER, after: 20 }),
      P("元の世界に帰るには、20日間で ブリブリ・ワールドカップに優勝するしかない！", { align: AlignmentType.CENTER, after: 20 }),
      P("毎日5つのナゾと漢字PK戦をクリアして、大魔王ベンピーをたおせ！", { align: AlignmentType.CENTER, after: 0 }),
    ], { fill: "FFF8E1", borders: box("F9A825", 16), margins: { top: 140, bottom: 140, left: 200, right: 200 } }),
    spacer(30),
    P("この問題集の持ち主：　　　　　　　　　　　　　キャプテン", { size: 30, bold: true, font: F_HEAD, align: AlignmentType.CENTER, border: { bottom: line(BROWN, 8) } }),
  ];
}

function howTo() {
  const steps = [
    ["1", "1日1回、ストーリーを読む", "まずは物語を読もう。声に出して読むと、もっとおもしろいぞ。国語の問題は、この物語から出る！"],
    ["2", "5つの問題にチャレンジ", "国語・算数・理科・社会・英語がまざって出てくる。だいたい15分でクリアできる。"],
    ["3", "漢字PK戦で3本シュート", "最後は漢字のPK戦。【　】の漢字を書いたり、読みがなを書いたりしよう。5日ごとに「復習PK」があるぞ。"],
    ["4", "おうちの人と答え合わせ", "まちがえても大丈夫。答えを見て直せば、ちゃんとポイントがもらえる！"],
    ["5", "ポイントを記録して、ランクアップ", "その日のページの下にポイントを書いて、「記録表」にも書きこもう。"],
  ];
  const stepRows = steps.map(([n, t, d]) => ({ h: 700, cells: [
    cell([P(n, { size: 40, bold: true, color: "FFFFFF", font: F_HEAD, align: AlignmentType.CENTER, after: 0, exact: 520 })], { w: 900, fill: BROWN, valign: VerticalAlign.CENTER }),
    cell([P(t, { size: 26, bold: true, font: F_HEAD, after: 30 }), P(d, { after: 0 })], { w: W - 900, fill: BROWN_L, valign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 200, right: 120 } }),
  ] }));
  const ptRows = [
    ["◎ 一発で正解", "2点"], ["○ 答えを見て直せた", "1点"], ["⚽ 漢字PK　ゴール1本", "1点"], ["1日の最高", "13点（20日で260点）"],
  ].map(([a, b], i) => ({ h: 420, cells: [
    cell([P(a, { bold: true, font: F_HEAD, after: 0 })], { w: 5000, borders: box("BCAAA4", 6), valign: VerticalAlign.CENTER, fill: i === 3 ? "FFF8E1" : undefined }),
    cell([P(b, { bold: true, font: F_HEAD, after: 0, align: AlignmentType.CENTER })], { w: 5200, borders: box("BCAAA4", 6), valign: VerticalAlign.CENTER, fill: i === 3 ? "FFF8E1" : undefined }),
  ] }));
  const rankRows = [["合計ポイント", "ランク", ""], ...RANKS].map((r, i) => ({ h: 420, cells: r.map((t, j) => cell([P(t, { bold: i === 0 || j === 1, font: F_HEAD, after: 0, color: i === 0 ? "FFFFFF" : undefined, align: j === 0 ? AlignmentType.CENTER : undefined })],
    { w: [2200, 4300, 3700][j], fill: i === 0 ? BROWN : (i % 2 ? undefined : BROWN_L), borders: box("BCAAA4", 6), valign: VerticalAlign.CENTER })) }));
  return [
    P("この問題集のあそび方", { size: 40, bold: true, font: F_HEAD, color: BROWN, align: AlignmentType.CENTER, pageBreak: true, after: 200 }),
    table([900, W - 900], stepRows, { borders: { ...TNOB, insideHorizontal: line("FFFFFF", 24) } }),
    P("ポイントのルール", { size: 28, bold: true, font: F_HEAD, color: BROWN, before: 260, after: 80 }),
    table([5000, 5200], ptRows),
    P("ランク表　〜合計ポイントでランクアップ！〜", { size: 28, bold: true, font: F_HEAD, color: BROWN, before: 260, after: 80 }),
    table([2200, 4300, 3700], rankRows),
  ];
}

function characters() {
  const list = [
    ["ball", "アツキ（主人公・背番号10）", "サッカーとゲームが大好きな小学5年生。ゲームの世界にすいこまれて、チーム・ブリブリのキャプテンになる。建築バトルも得意。"],
    ["hakase", "ウンチー博士（監督）", "うんこの形をした、物知りの博士。でも、ときどきとんでもないカンちがいをする。口ぐせは「〜なのじゃ」。"],
    ["butaro", "ブー太郎（FW）", "おならの力で空高く飛ぶ「ブーストジャンプ」が得意。ちょっと口が悪いけど、仲間思い。"],
    ["puko", "プゥ子（GK）", "ももの形をしたキーパー。おしりでボールをはじき返す「おしりセーブ」の名人。"],
    ["heppy", "ヘッピー（MF）", "パスの天才。でも、きんちょうすると「プッ、プッ」とおならが止まらなくなる。"],
    ["sukashi", "スカシ丸（DF）", "半とうめいの忍者。音のしないおならで気配を消して、ボールをうばう。「〜でござる」。"],
    ["benpi", "大魔王ベンピー（ラスボス）", "おなかがいつもカッチカチ。なんでもブロックでふさいでしまう。チーム「カチカチ・ブロッカーズ」をひきいる。"],
  ];
  const rows = list.map(([im, name, desc], i) => ({ h: 1500, cells: [
    cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [img(im, 84)] })], { w: 1700, valign: VerticalAlign.CENTER, fill: i % 2 ? undefined : BROWN_L }),
    cell([P(name, { size: 26, bold: true, font: F_HEAD, color: i === 6 ? "4E342E" : BROWN, after: 40 }), P(desc, { after: 0 })],
      { w: W - 1700, valign: VerticalAlign.CENTER, fill: i % 2 ? undefined : BROWN_L, margins: { top: 60, bottom: 60, left: 200, right: 160 } }),
  ] }));
  return [
    P("登場人物しょうかい", { size: 40, bold: true, font: F_HEAD, color: BROWN, align: AlignmentType.CENTER, pageBreak: true, after: 200 }),
    table([1700, W - 1700], rows),
  ];
}

function recordSheet() {
  const chapters = ["第1章　ブリブリ・ワールドへようこそ", "第2章　建築バトルの島", "第3章　ストームの谷", "第4章　決勝！大魔王ベンピー"];
  const cw = W / 5;
  const rows = [];
  chapters.forEach((c, w) => {
    rows.push({ h: 380, cells: [cell([P(c, { bold: true, font: F_HEAD, color: "FFFFFF", after: 0 })], { w: W, span: 5, fill: BROWN, valign: VerticalAlign.CENTER })] });
    rows.push({ h: 1500, cells: [0, 1, 2, 3, 4].map((k) => {
      const n = w * 5 + k + 1;
      return cell([
        P("DAY " + n, { size: 22, bold: true, font: F_HEAD, color: BROWN, align: AlignmentType.CENTER, after: 40 }),
        P("　　　点", { size: 22, align: AlignmentType.CENTER, after: 40, border: { bottom: line("BCAAA4", 6) } }),
        P("合計　　　点", { size: 17, color: "757575", align: AlignmentType.CENTER, after: 20 }),
        P("スタンプ", { size: 14, color: "BDBDBD", align: AlignmentType.CENTER, after: 0 }),
      ], { w: cw, borders: box("BCAAA4", 6) });
    }) });
  });
  return [
    P("ポイント記録表", { size: 40, bold: true, font: F_HEAD, color: BROWN, align: AlignmentType.CENTER, pageBreak: true, after: 60 }),
    P("毎日のポイントと、ここまでの合計を書こう。クリアしたら、スタンプやシールをはってもらおう！", { align: AlignmentType.CENTER, after: 160 }),
    table(Array(5).fill(cw), rows),
    P("いまのランク：", { size: 26, bold: true, font: F_HEAD, before: 300, after: 300, border: { bottom: line(BROWN, 8) } }),
    P("※ ランク表は「あそび方」のページにあるよ。", { size: 18, color: "757575" }),
  ];
}

function award() {
  const C = AlignmentType.CENTER;
  const inner = [
    spacer(40),
    P("表　しょう　状", { size: 72, bold: true, font: F_HEAD, align: C, after: 500, exact: 980 }),
    P("　　　　　　　　　　　　キャプテン　どの", { size: 36, bold: true, font: F_HEAD, align: C, after: 500, border: { bottom: line(BROWN, 8) } }),
    P("あなたは、20日間の大ぼうけんを最後までやりとげ、", { size: 30, align: C, after: 120, exact: 560 }),
    P("ブリブリ・ワールドカップで、みごと優勝しました。", { size: 30, align: C, after: 120, exact: 560 }),
    P("毎日コツコツ続けたそのがんばりを、ここにたたえます。", { size: 30, align: C, after: 400, exact: 560 }),
    new Paragraph({ alignment: C, spacing: { after: 400 }, children: [img("trophy", 210)] }),
    P("合計ポイント：　　　　　点", { size: 30, bold: true, font: F_HEAD, align: C, after: 200 }),
    P("最終ランク：　　　　　　　　　　　　　　", { size: 30, bold: true, font: F_HEAD, align: C, after: 500 }),
    P("　　　　年　　　月　　　日", { size: 28, align: C, after: 300 }),
    P("ブリブリ・ワールドカップ実行委員会", { size: 28, align: C, after: 60 }),
    P("会長　ウンチー博士", { size: 34, bold: true, font: F_HEAD, align: C, after: 300 }),
    P("おうちの人から：", { size: 24, before: 200, after: 0 }),
  ];
  return [
    new Paragraph({ pageBreakBefore: true, spacing: { after: 0, line: 40, lineRule: "exact" }, children: [] }),
    table([W], [{ cells: [cell(inner, { w: W, borders: box("C9A227", 36, BorderStyle.DOUBLE), margins: { top: 300, bottom: 300, left: 500, right: 500 } })], h: 14900 }]),
  ];
}

const footer = (label) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  new TextRun({ children: [label + "　－ ", PageNumber.CURRENT, " －"], size: 16, color: "9E9E9E", font: font(F_HEAD) }),
] })] });

const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 700, left: 853, right: 853, footer: 300 } };
const STYLES = { default: { document: { run: { font: font(F_BODY), size: SZ } } } };

async function buildWorkbook() {
  const doc = new Document({
    creator: "ブリブリ・ワールドカップ実行委員会",
    title: "アツキのブリブリ・ワールドカップ 問題集",
    styles: STYLES,
    sections: [{
      properties: { page: PAGE },
      footers: { default: footer("アツキのブリブリ・ワールドカップ") },
      children: [...cover(), ...howTo(), ...characters(), ...recordSheet(), ...DAYS.flatMap(dayPages), ...award()],
    }],
  });
  const f = path.join(OUT_DIR, "アツキのブリブリ・ワールドカップ_問題集.docx");
  fs.writeFileSync(f, await Packer.toBuffer(doc));
  return f;
}

// ───────── 答えと解説（おうちの方用）
function answerBook() {
  const H = (t, o = {}) => P(t, { size: o.size || 30, bold: true, font: F_HEAD, color: o.color || BROWN, before: o.before ?? 240, after: 100, keepNext: true, pageBreak: o.pageBreak, border: o.border });
  const B = (t, o = {}) => P(t, { size: 21, after: 60, line: 300, ...o });
  const bullet = (t) => P("・" + t, { size: 21, after: 40, line: 300, indent: { left: 360, hanging: 220 } });

  const intro = [
    spacer(60),
    P("アツキのブリブリ・ワールドカップ", { size: 44, bold: true, font: F_HEAD, color: BROWN, align: AlignmentType.CENTER, after: 120 }),
    P("答えと解説", { size: 60, bold: true, font: F_HEAD, align: AlignmentType.CENTER, after: 120, exact: 800 }),
    P("（おうちの方用）", { size: 30, font: F_HEAD, align: AlignmentType.CENTER, after: 400 }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [img("hakase", 180)] }),
    P("※ この冊子は、お子さんの手の届かないところに保管してください。", { size: 20, color: "757575", align: AlignmentType.CENTER, before: 400 }),

    H("保護者の方へ", { pageBreak: true, size: 36, before: 0 }),
    H("この問題集のねらい", { size: 26 }),
    B("「勉強はきらい。でもサッカーとゲームは大好き」というアツキくんが、毎日自分から机に向かえることを第一の目標にしています。20日間の続き物の物語（サッカー×ゲームの世界）の中に、5年生の国語・算数・理科・社会・英語の学習内容を問題として組みこみました。毎日「つづく…」で終わるので、次の日の続きが気になる作りになっています。"),
    H("1日の進め方（目安15分）", { size: 26 }),
    bullet("ストーリーを読む → 問題1〜5 → 漢字PK戦（3問）→ 答え合わせ → ポイント記入、の順です。"),
    bullet("1日2ページ（両面印刷なら1枚）です。最初の4ページ（表紙・あそび方・登場人物・記録表）は最初に一度だけ使います。"),
    bullet("答え合わせは、できればおうちの方と一緒に。まちがえた問題は、答えを見て直せばポイントになります。"),
    H("ポイントのつけ方", { size: 26 }),
    bullet("一発で正解＝2点、答えを見て直せた＝1点、漢字PK＝1本1点。1日の最高は13点、20日で260点です。"),
    bullet("(1)(2) がある問題は「半分以上できたら正解」くらいの、ゆるめの判定をおすすめします。続けることが一番の目的です。"),
    bullet("「答えが1つではない問題」（自分の考えを書く問題など）は、書けていれば正解にしてください。"),
    H("国語が苦手なアツキくんのための工夫", { size: 26 }),
    bullet("毎日の物語そのものが国語の読解文です。国語の(1)は、ほぼ毎回その日の物語から出題しています。"),
    bullet("前半は記号で選ぶ問題、後半になるほど「〜から。」「〜こと。」で書く問題、最後は「自分の考え」を書く問題へと、少しずつ書く量を増やしています。"),
    bullet("答え方の3つの型をくり返し練習します：理由をきかれたら「〜から。」、気持ちをきかれたら「〜気持ち」、書きぬきは「物語の中から探す」。"),
    bullet("漢字PK戦は5年生で習う漢字を中心に、全60本。5日目ごと（Day 5・10・15・20）はその週の漢字の復習です。答えのマスは文字数分あり、送りがなもマスに書きます。"),
    bullet("物語を声に出して読む（音読）のもおすすめです。登場人物のセリフを、おうちの方と役を分けて読むと盛り上がります。"),
    H("声かけのコツ", { size: 26 }),
    bullet("正解の数より「今日もやった」ことをほめてください。記録表のスタンプやシールが効果的です。"),
    bullet("まちがえたときは「おしい！」「どこを読めばわかったかな？」と、物語にもどる声かけを。"),
    bullet("各Dayの解説の最後に「声かけヒント」をつけています。"),
    H("下ネタについて", { size: 26 }),
    bullet("「うんこ・おなら・おしり・はなくそ」など、子どもが学校で話しても問題ない範囲にとどめています。人をからかったり、見た目を笑ったりする内容は入れていません。"),
    bullet("「人のたんじょう」（Day 19）と「公害」（Day 17）の単元は、ふざけずにまじめに扱っています。"),
    bullet("物語中の数字や設定（おならの長さなど）はゲームの世界の話です。社会・理科の知識は事実にもとづいています（統計は近年のデータ）。"),
    H("印刷について", { size: 26 }),
    bullet("A4で印刷してください。両面印刷にすると、1日分が1枚になります。"),
    bullet("Windowsでは「UDデジタル教科書体」（教科書と同じ形の文字）で表示されます。ほかの環境では別の文字で表示されることがありますが、内容は変わりません。"),
  ];

  // 学習内容一覧
  const subjOrder = ["国語", "算数", "理科", "社会", "英語"];
  const cols = [700, 1600, 1600, 1600, 1600, 1550, 1550];
  const head = ["DAY", ...subjOrder, "漢字PK"].map((h, i) => cell([P(h, { size: 16, bold: true, font: F_HEAD, color: "FFFFFF", align: AlignmentType.CENTER, after: 0 })],
    { w: cols[i], fill: i === 0 ? BROWN : i === 6 ? "263238" : SUBJ[h].c, borders: box("BCAAA4", 4), margins: { top: 30, bottom: 30, left: 40, right: 40 } }));
  const listRows = [{ cells: head, h: 300 }, ...DAYS.map((d) => ({ cells: [
    cell([P(String(d.n), { size: 16, bold: true, align: AlignmentType.CENTER, after: 0, font: F_HEAD })], { w: cols[0], borders: box("BCAAA4", 4), valign: VerticalAlign.CENTER, fill: BROWN_L }),
    ...subjOrder.map((s, i) => cell([P(d.P.find((p) => p.s === s).t, { size: 15, after: 0, line: 220 })], { w: cols[i + 1], borders: box("BCAAA4", 4), margins: { top: 30, bottom: 30, left: 50, right: 30 } })),
    cell([P(d.pk.map((k) => k.w === "書" ? k.a : k.y).join("・") + (d.review ? "（復習）" : ""), { size: 15, after: 0, line: 220 })], { w: cols[6], borders: box("BCAAA4", 4), margins: { top: 30, bottom: 30, left: 50, right: 30 } }),
  ] }))];
  const overview = [H("学習内容の一覧", { pageBreak: true, size: 36, before: 0 }), table(cols, listRows)];

  // 各日の答え
  const ac = [1000, 3300, 5900];
  const answers = [H("答えと解説", { pageBreak: true, size: 36, before: 0 })];
  DAYS.forEach((d) => {
    answers.push(P(`DAY ${d.n}　${d.t}`, { size: 26, bold: true, font: F_HEAD, color: "FFFFFF", before: 200, after: 80, keepNext: true,
      children: [new TextRun({ text: ` DAY ${d.n} `, bold: true, size: 26, color: "FFFFFF", font: font(F_HEAD), shading: shade(BROWN) }),
        new TextRun({ text: `　${d.t}`, bold: true, size: 26, color: BROWN, font: font(F_HEAD) })] }));
    const rows = d.P.map((p, i) => ({ cells: [
      cell([P(`問題${i + 1}`, { size: 16, after: 0, align: AlignmentType.CENTER, color: "FFFFFF", bold: true, font: F_HEAD }), P(p.s, { size: 18, after: 0, align: AlignmentType.CENTER, color: "FFFFFF", bold: true, font: F_HEAD })],
        { w: ac[0], fill: SUBJ[p.s].c, borders: box("BCAAA4", 4), valign: VerticalAlign.CENTER }),
      cell(p.q.map((q) => P(q.a, { size: 19, bold: true, after: 20, line: 260 })), { w: ac[1], borders: box("BCAAA4", 4), fill: SUBJ[p.s].l }),
      cell([P(p.ex, { size: 17, after: 0, line: 250 })], { w: ac[2], borders: box("BCAAA4", 4) }),
    ] }));
    rows.push({ cells: [
      cell([P("漢字PK", { size: 16, after: 0, align: AlignmentType.CENTER, color: "FFFFFF", bold: true, font: F_HEAD })], { w: ac[0], fill: "263238", borders: box("BCAAA4", 4), valign: VerticalAlign.CENTER }),
      cell([P(d.pk.map((k, i) => `${"①②③"[i]}${k.a}`).join("　"), { size: 19, bold: true, after: 0 })], { w: ac[1], borders: box("BCAAA4", 4), fill: "ECEFF1" }),
      cell([P(d.pk.map((k, i) => `${"①②③"[i]}${k.w === "書" ? "書き" : "読み"}「${k.y}」`).join("　") + (d.review ? "（復習）" : ""), { size: 17, after: 0 })], { w: ac[2], borders: box("BCAAA4", 4) }),
    ] });
    answers.push(table(ac, rows));
    answers.push(P("声かけヒント：" + d.tip, { size: 18, color: "5D4037", before: 40, after: 60, children: [img("unko", 16), ...runs("　声かけヒント：" + d.tip, { size: 18, color: "5D4037" })] }));
  });
  return [...intro, ...overview, ...answers];
}

async function buildAnswers() {
  const doc = new Document({
    creator: "ブリブリ・ワールドカップ実行委員会",
    title: "アツキのブリブリ・ワールドカップ 答えと解説",
    styles: STYLES,
    sections: [{ properties: { page: PAGE }, footers: { default: footer("答えと解説（おうちの方用）") }, children: answerBook() }],
  });
  const f = path.join(OUT_DIR, "アツキのブリブリ・ワールドカップ_答えと解説.docx");
  fs.writeFileSync(f, await Packer.toBuffer(doc));
  return f;
}

(async () => {
  console.log(await buildWorkbook());
  console.log(await buildAnswers());
})();
