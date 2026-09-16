// 校長会説明資料 ― 校長権限でつくる令和9年度の学校づくり
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";            // 13.333 x 7.5
pres.author = "教育委員会事務局";
pres.title = "校長権限でつくる令和9年度の学校づくり";

const W = 13.333, H = 7.5;
const M = 0.7;                           // 左右余白
const CW = W - M * 2;                    // コンテンツ幅 11.933

// ── パレット（育む・伸ばす・余白）──────────────────
const INK   = "14312A";   // 濃いパイン（濃色背景）
const PINE  = "1F5D4C";   // 主色
const MOSS  = "7FA86B";   // 若葉
const MOSSD = "5C8A4A";   // 若葉（濃）
const GOLD  = "B26B1B";   // 琥珀（文字可）
const GOLDL = "EBB84D";   // 琥珀（塗り）
const PAPER = "FFFFFF";
const MIST  = "F1F5F2";   // 淡い緑グレー（カード）
const MIST2 = "E4EDE6";   // やや濃いカード
const GREY  = "5B6660";
const TEXT  = "1C2320";
const LINEC = "D5E0D8";

const JP = "Meiryo";

let n = 0;

// ── ヘルパー ────────────────────────────────────────
function newSlide(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? INK : PAPER };
  n += 1;
  s.addText(String(n), {
    isTextBox: true, x: W - 1.0, y: H - 0.52, w: 0.5, h: 0.3,
    fontFace: JP, fontSize: 9, color: dark ? "6E8A7E" : "9AA8A0",
    align: "right", margin: 0,
  });
  return s;
}

function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, fontFace: JP, color: TEXT, margin: 0 }, o));
}

function card(s, x, y, w, h, fill, opts) {
  const o = opts || {};
  s.addShape(o.round === false ? pres.ShapeType.rect : pres.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: fill },
    line: o.line ? { color: o.line, width: o.lineW || 1 } : { type: "none" },
    rectRadius: 0.06,
  });
}

// 丸バッジ（本資料のモチーフ）
function badge(s, x, y, d, label, fill, fontColor, size) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" },
  });
  txt(s, label, {
    x, y: y + (d - 0.36) / 2, w: d, h: 0.36,
    fontSize: size || 15, bold: true, color: fontColor || PAPER, align: "center",
  });
}

// 見出し（コンテンツ用）
function head(s, kicker, title) {
  if (kicker) {
    txt(s, kicker, {
      x: M, y: 0.40, w: CW, h: 0.26,
      fontSize: 11, bold: true, color: MOSSD, charSpacing: 1.5,
    });
  }
  txt(s, title, {
    x: M, y: 0.68, w: CW, h: 0.62,
    fontSize: 28, bold: true, color: INK,
  });
}

// 濃色スライドの見出し
function headDark(s, kicker, title) {
  if (kicker) {
    txt(s, kicker, {
      x: M, y: 0.40, w: CW, h: 0.26,
      fontSize: 11, bold: true, color: MOSS, charSpacing: 1.5,
    });
  }
  txt(s, title, {
    x: M, y: 0.68, w: CW, h: 0.62,
    fontSize: 28, bold: true, color: PAPER,
  });
}

function bullets(s, items, o) {
  const runs = items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i !== items.length - 1 },
  }));
  s.addText(runs, Object.assign({
    isTextBox: true, fontFace: JP, color: TEXT, fontSize: 14,
    paraSpaceAfter: 8, margin: 0, valign: "top",
  }, o));
}

function foot(s, text) {
  txt(s, text, {
    x: M, y: H - 0.62, w: CW - 1.0, h: 0.34,
    fontSize: 9.5, color: GREY, valign: "top",
  });
}

// セクション扉
function section(num, title, lead, items) {
  const s = newSlide(true);
  badge(s, M, 1.62, 1.22, num, MOSS, INK, 40);
  txt(s, title, {
    x: M + 1.65, y: 1.70, w: CW - 1.65, h: 0.78,
    fontSize: 36, bold: true, color: PAPER,
  });
  txt(s, lead, {
    x: M + 1.65, y: 2.54, w: CW - 1.65, h: 0.5,
    fontSize: 15, color: MOSS,
  });
  let y = 3.62;
  items.forEach((it) => {
    badge(s, M + 0.08, y + 0.02, 0.3, "", MOSSD, PAPER, 10);
    txt(s, it, {
      x: M + 0.62, y, w: CW - 0.9, h: 0.42,
      fontSize: 14.5, color: "D6E4D8",
    });
    y += 0.56;
  });
  return s;
}

/* ══════════════════════════════════════════════════
   導入（1–5）
   ══════════════════════════════════════════════════ */

// 1 表紙
{
  const s = newSlide(true);
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.35, y: -1.35, w: 5.6, h: 5.6,
    fill: { color: PINE }, line: { type: "none" },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 10.85, y: 3.95, w: 3.1, h: 3.1,
    fill: { color: MOSSD }, line: { type: "none" },
  });
  txt(s, "中央教育審議会 教育課程企画特別部会「審議まとめ（素案）」を読む", {
    x: M, y: 1.55, w: 8.6, h: 0.34,
    fontSize: 13, bold: true, color: MOSS,
  });
  txt(s, "「素案」を、\n自分の学校の言葉にする。", {
    x: M, y: 2.05, w: 8.6, h: 1.85,
    fontSize: 40, bold: true, color: PAPER, lineSpacing: 52,
  });
  txt(s, "校長権限でつくる　令和9年度の学校づくり", {
    x: M, y: 4.08, w: 8.6, h: 0.5,
    fontSize: 21, bold: true, color: GOLDL,
  });
  s.addShape(pres.ShapeType.line, {
    x: M, y: 4.85, w: 3.2, h: 0,
    line: { color: "3D6B5A", width: 1 },
  });
  txt(s, "区立小・中学校 校長会\n令和8年　月　日", {
    x: M, y: 5.12, w: 6.0, h: 0.8,
    fontSize: 13, color: "AFC5B6", lineSpacing: 22,
  });
  s.addNotes(
    "本日は、8月31日に中教審の教育課程企画特別部会に示された「審議まとめ（素案）」を取り上げます。\n" +
    "ねらいは一つです。この素案を、国の動きの解説として聞いて終わるのではなく、先生方それぞれの学校で、令和9年度の経営にどう落とすかという形でお持ち帰りいただくことです。"
  );
}

// 2 全体像
{
  const s = newSlide();
  head(s, "本日の流れ", "5つのパートでお話しします");
  const parts = [
    ["Ⅰ", "なぜ今、校長が動くのか", "2030年の話ではありません"],
    ["Ⅱ", "素案は何を言っているか", "理念と、制度の中身"],
    ["Ⅲ", "校長の権限はどこまで届くか", "できること・できないこと"],
    ["Ⅳ", "明日から打てる手", "先行事例と10の打ち手"],
    ["Ⅴ", "令和9年度の学校経営に落とす", "経営計画への書き込み方"],
  ];
  let y = 1.62;
  parts.forEach((p, i) => {
    card(s, M, y, CW, 0.88, i === 4 ? MIST2 : MIST);
    badge(s, M + 0.30, y + 0.19, 0.50, p[0], i === 4 ? GOLD : PINE, PAPER, 16);
    txt(s, p[1], {
      x: M + 1.02, y: y + 0.15, w: 5.6, h: 0.36,
      fontSize: 17, bold: true, color: INK,
    });
    txt(s, p[2], {
      x: M + 1.02, y: y + 0.52, w: 5.6, h: 0.3,
      fontSize: 12, color: GREY,
    });
    y += 0.96;
  });
  card(s, M, 6.44, CW, 0.62, MIST);
  txt(s, "巻末に、想定質問集（Q1〜Q20）を付けています。質疑応答のときにお使いください。", {
    x: M + 0.45, y: 6.58, w: CW - 0.9, h: 0.36,
    fontSize: 13, bold: true, color: PINE,
  });
  s.addNotes("全体像です。ⅠとⅡで素案の中身を共有し、Ⅲで校長の権限を整理します。ⅣとⅤが本題で、自校で何をするかの部分です。");
}

// 3 結論
{
  const s = newSlide();
  head(s, "結論", "本日お持ち帰りいただきたいこと");
  const concl = [
    ["1", "2030年の話ではない", "調整授業時数制度は2028年度から先行導入できます。逆算すると、設計に使える年度は令和9年度と令和10年度しかありません。"],
    ["2", "「学校の裁量」を広げる改訂", "全国一律の横並びと決別する改訂です。制度が校長の判断に委ねる範囲が、これまでになく広がります。"],
    ["3", "令和9年度は「準備の年」", "確定を待つ必要はありません。今の法令のままで打てる手が、すでにいくつもあります。"],
  ];
  let y = 1.62;
  concl.forEach((c) => {
    card(s, M, y, CW, 1.52, MIST);
    badge(s, M + 0.34, y + 0.30, 0.64, c[0], PINE, PAPER, 22);
    txt(s, c[1], {
      x: M + 1.18, y: y + 0.24, w: CW - 1.6, h: 0.42,
      fontSize: 19, bold: true, color: INK,
    });
    txt(s, c[2], {
      x: M + 1.18, y: y + 0.72, w: CW - 1.6, h: 0.64,
      fontSize: 13.5, color: TEXT, lineSpacing: 20,
    });
    y += 1.66;
  });
  s.addNotes("先に結論を申し上げます。この3つだけ覚えて帰っていただければ、本日の目的は達成です。特に3つめ、「確定を待たなくていい」というところが今日の肝です。");
}

// 4 素案とは
{
  const s = newSlide();
  head(s, "取り上げる資料", "本日の土台となる「素案」");
  card(s, M, 1.60, CW, 2.05, MIST2);
  txt(s, "次期学習指導要領等に向けた　審議まとめ（素案）", {
    x: M + 0.5, y: 1.88, w: CW - 1.0, h: 0.55,
    fontSize: 26, bold: true, color: INK,
  });
  txt(s, "中央教育審議会　初等中等教育分科会　教育課程部会\n教育課程企画特別部会（第17回）　令和8年8月31日　資料1", {
    x: M + 0.5, y: 2.52, w: CW - 1.0, h: 0.8,
    fontSize: 14, color: PINE, lineSpacing: 23,
  });
  const meta = [
    ["位置づけ", "答申に向けた審議のまとめの素案。国の方向性を示す文書"],
    ["今後", "令和8年冬ごろ答申 → 令和8年度末に改訂学習指導要領を告示"],
    ["性格", "確定版ではない。今後の審議で表現・数値は変わりうる"],
  ];
  let y = 3.95;
  meta.forEach((m) => {
    card(s, M, y, 2.5, 0.78, PINE);
    txt(s, m[0], {
      x: M, y: y + 0.21, w: 2.5, h: 0.36,
      fontSize: 15, bold: true, color: PAPER, align: "center",
    });
    txt(s, m[1], {
      x: M + 2.85, y: y + 0.20, w: CW - 2.95, h: 0.42,
      fontSize: 14, color: TEXT,
    });
    y += 0.94;
  });
  s.addNotes("土台にするのはこの一本です。8月31日の部会に示された資料1、審議まとめの素案です。答申前の段階の文書ですが、方向性ははっきり出ています。");
}

// 5 素案であることの意味
{
  const s = newSlide();
  head(s, "前提の確認", "「素案だから待つ」が成り立たない理由");
  card(s, M, 1.62, 5.75, 2.25, MIST);
  badge(s, M + 0.36, 1.92, 0.5, "✓", PINE, PAPER, 16);
  txt(s, "確かに、確定ではありません", {
    x: M + 1.02, y: 1.94, w: 4.5, h: 0.4,
    fontSize: 17, bold: true, color: INK,
  });
  bullets(s, [
    "教科名、授業時数、評価の方法は今後変わりうる",
    "細部を今から作り込むのは、たしかに危険",
  ], { x: M + 0.42, y: 2.58, w: 5.0, h: 1.1, fontSize: 13 });

  card(s, M + 6.18, 1.62, 5.75, 2.25, INK);
  badge(s, M + 6.54, 1.92, 0.5, "!", GOLDL, INK, 18);
  txt(s, "しかし、待つと間に合いません", {
    x: M + 7.20, y: 1.94, w: 4.5, h: 0.4,
    fontSize: 17, bold: true, color: PAPER,
  });
  s.addText([
    { text: "先行導入は2028年度から", options: { bullet: true, breakLine: true } },
    { text: "告示を待つと、設計期間が1年を切る", options: { bullet: true } },
  ], {
    isTextBox: true, fontFace: JP, color: "D6E4D8", fontSize: 13,
    x: M + 6.60, y: 2.58, w: 5.0, h: 1.1, paraSpaceAfter: 8, margin: 0, valign: "top",
  });

  card(s, M, 4.12, CW, 1.62, MIST2);
  txt(s, "変わらないのは「理念」と「学校に委ねる」という方向。\nここは今から仕込めます。変わりうる「数字」は、確定してから詰めれば間に合います。", {
    x: M + 0.55, y: 4.42, w: CW - 1.1, h: 1.05,
    fontSize: 17, bold: true, color: INK, lineSpacing: 30,
  });
  s.addNotes("素案だから様子を見よう、という判断は自然です。ただ、先行導入が2028年度からである以上、告示を待つと設計に使える時間がほとんど残りません。\n分けて考えたいのは、理念と方向は今から仕込める、数字は後から詰めればよい、ということです。");
}

/* ══════════════════════════════════════════════════
   Ⅰ なぜ今、校長が動くのか（6–12）
   ══════════════════════════════════════════════════ */

// 6 扉
section("Ⅰ", "なぜ今、校長が動くのか", "2030年の話ではありません", [
  "逆算すると、設計に使えるのは令和9・10年度だけ",
  "今回の改訂は「全国一律」と決別する改訂",
  "裁量が増えることは、自動的に良くなることではない",
]).addNotes("まずは、なぜ今なのか、というところからです。");

// 7 逆算カレンダー
{
  const s = newSlide();
  head(s, "Ⅰ－1", "逆算カレンダー");
  const tl = [
    ["2026\n令和8", "冬ごろ\n答申", PINE],
    ["2027\n3月", "改訂指導要領\n告示", PINE],
    ["2028\n令和10", "幼稚園 全面実施\n調整授業時数制度 先行導入可", GOLD],
    ["2030\n令和12", "小学校\n全面実施", MOSSD],
    ["2031\n令和13", "中学校\n全面実施", MOSSD],
  ];
  const y0 = 2.55;
  s.addShape(pres.ShapeType.line, {
    x: M + 0.5, y: y0 + 0.36, w: CW - 1.4, h: 0,
    line: { color: LINEC, width: 2 },
  });
  const step = (CW - 1.4) / 4;
  tl.forEach((t, i) => {
    const cx = M + 0.5 + step * i;
    badge(s, cx - 0.36, y0, 0.72, "", t[2], PAPER, 12);
    txt(s, t[0], {
      x: cx - 1.05, y: y0 - 1.08, w: 2.1, h: 0.85,
      fontSize: 15, bold: true, color: t[2], align: "center", lineSpacing: 19,
    });
    txt(s, t[1], {
      x: cx - 1.15, y: y0 + 0.95, w: 2.3, h: 1.0,
      fontSize: 12, color: TEXT, align: "center", lineSpacing: 18,
    });
  });
  card(s, M, 4.75, CW, 1.28, MIST2);
  txt(s, "設計に使える年度は　令和9年度・令和10年度　の2年間", {
    x: M + 0.5, y: 5.02, w: CW - 1.0, h: 0.42,
    fontSize: 20, bold: true, color: INK,
  });
  txt(s, "2028年度（令和10年度）の先行導入に間に合わせるなら、令和9年度中に校内の合意形成と設計を終えておく必要があります。", {
    x: M + 0.5, y: 5.50, w: CW - 1.0, h: 0.4,
    fontSize: 13, color: TEXT,
  });
  foot(s, "※スケジュールは素案および報道に基づく。今後変更の可能性があります。");
  s.addNotes("横軸で見ていただくと分かりやすいと思います。注目していただきたいのは真ん中、2028年度です。全面実施は小学校が2030年度ですが、調整授業時数制度だけは2028年度から先行して導入できるとされています。\nそこから逆算すると、設計に使える年度は令和9年度と令和10年度の2年しかありません。");
}

// 8 2030年の話ではない
{
  const s = newSlide();
  head(s, "Ⅰ－2", "「2030年の話」ではない");
  card(s, M, 1.62, 5.75, 3.18, MIST);
  txt(s, "よくある受け止め", {
    x: M + 0.42, y: 1.92, w: 5.0, h: 0.36,
    fontSize: 15, bold: true, color: GREY,
  });
  txt(s, "小学校の全面実施は2030年度。\nまだ4年ある。", {
    x: M + 0.42, y: 2.40, w: 5.0, h: 1.0,
    fontSize: 22, bold: true, color: GREY, lineSpacing: 34,
  });
  bullets(s, [
    "告示を待ってから校内で検討",
    "区の説明会を待ってから動く",
    "教育課程の見直しは令和11年度から",
  ], { x: M + 0.42, y: 3.52, w: 5.0, h: 1.9, fontSize: 13.5, color: GREY });

  card(s, M + 6.18, 1.62, 5.75, 3.18, INK);
  txt(s, "実際のスケジュール", {
    x: M + 6.60, y: 1.92, w: 5.0, h: 0.36,
    fontSize: 15, bold: true, color: MOSS,
  });
  txt(s, "先行導入は2028年度。\n準備期間は実質2年。", {
    x: M + 6.60, y: 2.40, w: 5.0, h: 1.0,
    fontSize: 22, bold: true, color: PAPER, lineSpacing: 34,
  });
  s.addText([
    { text: "令和9年度：職員の理解と現状の棚卸し", options: { bullet: true, breakLine: true } },
    { text: "令和10年度：試行しながら設計を固める", options: { bullet: true, breakLine: true } },
    { text: "令和11年度以降：本格運用に接続", options: { bullet: true } },
  ], {
    isTextBox: true, fontFace: JP, color: "D6E4D8", fontSize: 13.5,
    x: M + 6.60, y: 3.52, w: 5.0, h: 1.9, paraSpaceAfter: 9, margin: 0, valign: "top",
  });
  card(s, M, 5.35, CW, 1.35, MIST2);
  txt(s, "違いは「意識」ではなく「持ち時間」です。", {
    x: M + 0.5, y: 5.62, w: CW - 1.0, h: 0.46,
    fontSize: 21, bold: true, color: INK,
  });
  txt(s, "先生方の熱意の問題ではありません。いつまでに何を決めるのか、その認識がずれているだけです。", {
    x: M + 0.5, y: 6.14, w: CW - 1.0, h: 0.40,
    fontSize: 13.5, color: TEXT,
  });
  s.addNotes("左が、よくある受け止め方だと思います。私も最初はそう思いました。\nですが右のように、先行導入から逆算すると景色が変わります。意識が低いという話ではなく、単純に持ち時間の認識がずれている、ということです。");
}

// 9 2年後の差
{
  const s = newSlide();
  head(s, "Ⅰ－3", "令和9年度の過ごし方で、2年後が決まります");
  const cols = [
    ["待つ学校", GREY, MIST, [
      "告示後に一斉に検討開始",
      "時数の組み替えは前例踏襲",
      "行事も研修もそのまま",
      "「やらされる改訂」になる",
      "職員は変更への対応に追われる",
    ]],
    ["つくる学校", PINE, MIST2, [
      "令和9年度に現状を棚卸し",
      "自校の重点に時間を寄せる設計",
      "行事と会議を先に整理",
      "「自分たちで決めた改訂」になる",
      "職員に余白が生まれる",
    ]],
  ];
  cols.forEach((c, i) => {
    const x = M + i * (CW / 2 + 0.16);
    const w = CW / 2 - 0.16;
    card(s, x, 1.62, w, 4.55, c[2]);
    badge(s, x + 0.4, 1.92, 0.5, i === 0 ? "A" : "B", c[1], PAPER, 16);
    txt(s, c[0], {
      x: x + 1.05, y: 1.96, w: w - 1.4, h: 0.42,
      fontSize: 21, bold: true, color: c[1],
    });
    bullets(s, c[3], {
      x: x + 0.45, y: 2.68, w: w - 0.9, h: 3.2,
      fontSize: 14, color: i === 0 ? GREY : TEXT,
    });
  });
  txt(s, "同じ制度でも、令和11年度に見える景色はまったく違います。", {
    x: M, y: 6.42, w: CW, h: 0.4,
    fontSize: 15, bold: true, color: INK, align: "center",
  });
  s.addNotes("同じ制度が来ても、令和9年度をどう過ごしたかで、2年後の学校の姿は大きく変わります。右側になるかどうかは、制度ではなく、校長の判断で決まる部分が大きいと考えています。");
}

// 10 全国一律との決別
{
  const s = newSlide();
  head(s, "Ⅰ－4", "今回の改訂の最大の特色");
  card(s, M, 1.60, CW, 1.42, INK);
  txt(s, "全国一律の「横並び」と決別し、各学校の裁量を大幅に広げる", {
    x: M + 0.55, y: 1.94, w: CW - 1.1, h: 0.6,
    fontSize: 24, bold: true, color: PAPER,
  });
  const shift = [
    ["これまで", "標準授業時数を全校ほぼ同じように運用", GREY, MIST],
    ["これから", "総時数は保ちつつ、配分は学校が決める", PINE, MIST2],
  ];
  let y = 3.25;
  shift.forEach((sh) => {
    card(s, M, y, CW, 1.15, sh[3]);
    txt(s, sh[0], {
      x: M + 0.45, y: y + 0.36, w: 1.9, h: 0.42,
      fontSize: 16, bold: true, color: sh[2],
    });
    txt(s, sh[1], {
      x: M + 2.55, y: y + 0.33, w: CW - 3.0, h: 0.5,
      fontSize: 18, color: TEXT,
    });
    y += 1.30;
  });
  card(s, M, 5.85, CW, 0.92, MIST);
  txt(s, "＝ 校長の判断が、これまでになく子供の学びに直結する改訂です。", {
    x: M + 0.5, y: 6.08, w: CW - 1.0, h: 0.46,
    fontSize: 17, bold: true, color: PINE,
  });
  s.addNotes("今回の改訂を一言でいうと、これです。全国一律の横並びをやめて、各学校の裁量を広げる。その核が、次の章で説明する調整授業時数制度です。\nつまり、校長の判断が子供の学びに直接効いてくる改訂だということです。");
}

// 11 裁量＝重荷にもなる
{
  const s = newSlide();
  head(s, "Ⅰ－5", "ただし、裁量が増えれば自動的に良くなるわけではない");
  card(s, M, 1.62, CW, 1.32, MIST2);
  txt(s, "「柔軟化は重荷にも？　問われるのは組織の成熟度」", {
    x: M + 0.55, y: 1.92, w: CW - 1.1, h: 0.5,
    fontSize: 21, bold: true, color: INK,
  });
  txt(s, "制度導入をめぐる報道の見出しより", {
    x: M + 0.55, y: 2.46, w: CW - 1.1, h: 0.3,
    fontSize: 11.5, color: GREY,
  });
  const risks = [
    ["決められない", "職員間で合意ができず、結局これまで通りになる"],
    ["過重になる", "削った分だけ別の活動を足してしまい、かえって負担増"],
    ["目的がずれる", "子供の学びではなく、消化のための時間割になる"],
  ];
  let y = 3.28;
  risks.forEach((r) => {
    card(s, M, y, CW, 0.92, MIST);
    badge(s, M + 0.34, y + 0.21, 0.5, "▲", GOLD, PAPER, 13);
    txt(s, r[0], {
      x: M + 1.02, y: y + 0.26, w: 2.6, h: 0.4,
      fontSize: 16, bold: true, color: INK,
    });
    txt(s, r[1], {
      x: M + 3.75, y: y + 0.27, w: CW - 4.2, h: 0.4,
      fontSize: 13.5, color: TEXT,
    });
    y += 1.05;
  });
  txt(s, "自由度が上がるほど、学校としての「決め方」が問われます。", {
    x: M, y: 6.48, w: CW, h: 0.4,
    fontSize: 15, bold: true, color: PINE,
  });
  s.addNotes("ここは正直にお伝えしておきたいところです。裁量が広がることは、そのまま重荷にもなりえます。\n決められない、かえって忙しくなる、目的がずれる。この3つは実際に起こりうると思います。だからこそ、決め方を先に用意しておく必要があります。");
}

// 12 だから校長が
{
  const s = newSlide(true);
  headDark(s, "Ⅰ　まとめ", "だから、校長が動きます");
  const pts = [
    ["時間の配分を決められるのは、学校です", "そして学校の教育課程に責任を持つのは校長です"],
    ["「決め方」を設計できるのも校長です", "誰と、いつ、何を根拠に決めるか。ここは権限の話です"],
    ["令和9年度の経営計画に書けば、動き出します", "書かなければ、日常業務が時間を埋めてしまいます"],
  ];
  let y = 1.85;
  pts.forEach((p, i) => {
    card(s, M, y, CW, 1.28, "1B3F35");
    badge(s, M + 0.36, y + 0.34, 0.6, String(i + 1), MOSS, INK, 20);
    txt(s, p[0], {
      x: M + 1.15, y: y + 0.26, w: CW - 1.6, h: 0.42,
      fontSize: 19, bold: true, color: PAPER,
    });
    txt(s, p[1], {
      x: M + 1.15, y: y + 0.73, w: CW - 1.6, h: 0.38,
      fontSize: 13, color: "AFC5B6",
    });
    y += 1.42;
  });
  s.addNotes("Ⅰ章のまとめです。時間の配分を決めるのは学校、その責任者は校長。そして決め方を設計できるのも校長です。\n最後の一行が実感としては一番大きくて、経営計画に書かないと、日常業務が時間を全部埋めてしまいます。");
}

/* ══════════════════════════════════════════════════
   Ⅱ 素案は何を言っているか（13–26）
   ══════════════════════════════════════════════════ */

// 13 扉
section("Ⅱ", "素案は何を言っているか", "理念と、制度の中身", [
  "基本理念は「多様な子供たちの『深い学び』を確かなものに」",
  "制度の核は「調整授業時数制度」",
  "情報活用能力を、すべての学びの基盤に位置づける",
]).addNotes("ここからは素案の中身です。理念と制度、両方おさえます。");

// 14 基本理念
{
  const s = newSlide();
  head(s, "Ⅱ－1", "基本理念");
  card(s, M, 1.70, CW, 1.95, INK);
  txt(s, "多様な子供たちの「深い学び」を\n確かなものに", {
    x: M + 0.6, y: 2.00, w: CW - 1.2, h: 1.35,
    fontSize: 32, bold: true, color: PAPER, lineSpacing: 46,
  });
  const reads = [
    ["「多様な」", "一部の子ではなく、目の前の全員が対象"],
    ["「深い学び」", "活動をこなすことではなく、考えが深まること"],
    ["「確かなものに」", "掛け声ではなく、実装しきるという宣言"],
  ];
  let y = 4.00;
  reads.forEach((r) => {
    card(s, M, y, CW, 0.78, MIST);
    txt(s, r[0], {
      x: M + 0.45, y: y + 0.20, w: 2.6, h: 0.4,
      fontSize: 16, bold: true, color: PINE,
    });
    txt(s, r[1], {
      x: M + 3.35, y: y + 0.21, w: CW - 3.8, h: 0.4,
      fontSize: 14, color: TEXT,
    });
    y += 0.90;
  });
  s.addNotes("素案の基本理念は、この一文に集約されています。読み解くポイントは3つです。多様な、つまり全員。深い学び、つまり活動の消化ではない。確かなものに、つまり今度こそ実装しきる、という意思表示です。");
}

// 15 3つの基盤
{
  const s = newSlide();
  head(s, "Ⅱ－2", "検討の基盤となる3つの考え方");
  const base = [
    ["主体的・対話的で\n深い学びの実装", "理念は現行と地続き。今回は「どう実装するか」に軸足", PINE],
    ["多様性の包摂", "特別な支援、日本語指導、不登校。例外ではなく前提として設計", MOSSD],
    ["実現可能性の確保", "現場が回らない改訂にしない。働き方改革と一体で考える", GOLD],
  ];
  const cw2 = (CW - 0.5) / 3;
  base.forEach((b, i) => {
    const x = M + i * (cw2 + 0.25);
    card(s, x, 1.70, cw2, 3.65, MIST);
    badge(s, x + cw2 / 2 - 0.35, 2.02, 0.7, String(i + 1), b[2], PAPER, 24);
    txt(s, b[0], {
      x: x + 0.28, y: 2.98, w: cw2 - 0.56, h: 1.0,
      fontSize: 18, bold: true, color: INK, align: "center", lineSpacing: 26,
    });
    txt(s, b[1], {
      x: x + 0.28, y: 4.08, w: cw2 - 0.56, h: 1.05,
      fontSize: 13, color: TEXT, align: "center", lineSpacing: 20,
    });
  });
  card(s, M, 5.62, CW, 1.05, MIST2);
  txt(s, "3つめの「実現可能性の確保」が、今回はっきり前面に出てきたところです。", {
    x: M + 0.5, y: 5.88, w: CW - 1.0, h: 0.5,
    fontSize: 16, bold: true, color: INK,
  });
  s.addNotes("検討の基盤が3つ示されています。1と2は従来からの流れですが、3つめの実現可能性の確保がはっきり前面に出てきたのが今回の特徴だと思います。現場が回らない改訂にはしない、という意思です。");
}

// 16 両輪
{
  const s = newSlide();
  head(s, "Ⅱ－3", "目指す姿は「両輪」で語られています");
  const wheels = [
    ["「好き」を育み、\n「得意」を伸ばす", "興味・関心を起点にする。\n全員を同じ地点に揃えることを目的にしない。", PINE],
    ["当事者意識をもって\n意見を形成し、対話と合意ができる", "自分の考えをもち、他者と折り合いをつける。\n民主的な社会の担い手を育てる。", MOSSD],
  ];
  wheels.forEach((w, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw3 = CW / 2 - 0.2;
    card(s, x, 1.72, cw3, 2.95, MIST);
    badge(s, x + 0.42, 2.02, 0.62, i === 0 ? "♡" : "◎", w[2], PAPER, 20);
    txt(s, w[0], {
      x: x + 0.42, y: 2.82, w: cw3 - 0.84, h: 0.95,
      fontSize: 19, bold: true, color: INK, lineSpacing: 27,
    });
    txt(s, w[1], {
      x: x + 0.42, y: 3.80, w: cw3 - 0.84, h: 0.72,
      fontSize: 13, color: TEXT, lineSpacing: 20,
    });
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: W / 2 - 0.34, y: 2.82, w: 0.68, h: 0.68,
    fill: { color: GOLDL }, line: { color: PAPER, width: 3 },
  });
  txt(s, "＋", {
    x: W / 2 - 0.34, y: 2.96, w: 0.68, h: 0.4,
    fontSize: 20, bold: true, color: INK, align: "center",
  });
  card(s, M, 4.95, CW, 1.62, INK);
  txt(s, "この2つは相互に有機的にかかわり合い、両輪として高まっていく", {
    x: M + 0.55, y: 5.22, w: CW - 1.1, h: 0.45,
    fontSize: 19, bold: true, color: PAPER,
  });
  txt(s, "どちらか一方ではありません。「好きなことだけ」でも「話し合いの形式だけ」でもない、という釘の刺し方がされています。", {
    x: M + 0.55, y: 5.78, w: CW - 1.1, h: 0.5,
    fontSize: 13.5, color: "AFC5B6",
  });
  s.addNotes("目指す姿は、この2つが両輪である、という書かれ方をしています。好きを伸ばすだけでも、話し合いの形だけでもだめで、この2つが絡み合って高まっていく教育課程を作ろう、ということです。");
}

// 17 余白
{
  const s = newSlide();
  head(s, "Ⅱ－4", "今回のキーワードは「余白」");
  card(s, M, 1.70, CW, 2.15, MIST2);
  txt(s, "余　白", {
    x: M + 0.6, y: 2.10, w: 3.4, h: 0.95,
    fontSize: 54, bold: true, color: PINE, charSpacing: 6,
  });
  txt(s, "子供が深く学ぶための、時間的なゆとり。\nそして、教師が教材研究と対話に使える時間。", {
    x: M + 4.5, y: 2.18, w: CW - 5.1, h: 1.1,
    fontSize: 18, color: INK, lineSpacing: 30,
  });
  const how = [
    ["学習内容の精選", "教える中身を絞る"],
    ["教科書の分量の見直し", "こなす量を減らす"],
    ["時数の柔軟化", "配分を学校が決める"],
  ];
  txt(s, "余白はどう生まれるか", {
    x: M, y: 4.15, w: CW, h: 0.36,
    fontSize: 15, bold: true, color: GREY,
  });
  const cw4 = (CW - 0.5) / 3;
  how.forEach((h, i) => {
    const x = M + i * (cw4 + 0.25);
    card(s, x, 4.62, cw4, 1.28, MIST);
    txt(s, h[0], {
      x: x + 0.25, y: 4.88, w: cw4 - 0.5, h: 0.42,
      fontSize: 16, bold: true, color: INK, align: "center",
    });
    txt(s, h[1], {
      x: x + 0.25, y: 5.34, w: cw4 - 0.5, h: 0.34,
      fontSize: 12.5, color: GREY, align: "center",
    });
  });
  txt(s, "余白づくりは、働き方改革と同じ方向を向いています。ここが今回、現場にとって一番大きい変化です。", {
    x: M, y: 6.18, w: CW, h: 0.44,
    fontSize: 15, bold: true, color: PINE,
  });
  s.addNotes("素案を通して出てくるキーワードが、この余白です。子供にとっての深い学びの時間であり、同時に教師が教材研究に使える時間でもあります。\n大事なのは、余白づくりが働き方改革と同じ方向を向いているという点です。");
}

// 18 オーバーロード
{
  const s = newSlide();
  head(s, "Ⅱ－5", "前提にあるのは「足し算の限界」");
  txt(s, "カリキュラム・オーバーロード", {
    x: M, y: 1.62, w: CW, h: 0.48,
    fontSize: 24, bold: true, color: INK,
  });
  txt(s, "求められる内容が増え続け、授業時数と現場の許容量を超えている状態", {
    x: M, y: 2.14, w: CW, h: 0.36,
    fontSize: 14, color: GREY,
  });
  const load = ["外国語", "プログラミング", "道徳の教科化", "主権者教育", "金融教育", "防災教育", "がん教育", "情報モラル", "キャリア教育", "食育"];
  let x = M, y = 2.72;
  load.forEach((l) => {
    const w = 0.42 + l.length * 0.185;
    if (x + w > M + CW) { x = M; y += 0.62; }
    card(s, x, y, w, 0.5, MIST);
    txt(s, l, {
      x, y: y + 0.13, w, h: 0.3,
      fontSize: 13, color: TEXT, align: "center",
    });
    x += w + 0.16;
  });
  card(s, M, 4.42, CW, 0.95, INK);
  txt(s, "これまでの改訂は、基本的に「足す」改訂でした。", {
    x: M + 0.5, y: 4.66, w: CW - 1.0, h: 0.46,
    fontSize: 19, bold: true, color: PAPER,
  });
  card(s, M, 5.55, CW, 1.15, MIST2);
  txt(s, "今回は「精選する」改訂です。減らした先に何を置くかを、学校が決めます。", {
    x: M + 0.5, y: 5.86, w: CW - 1.0, h: 0.52,
    fontSize: 20, bold: true, color: PINE,
  });
  s.addNotes("なぜ余白が必要かというと、前提にオーバーロードがあるからです。ここに並べたものは、いずれもこの20年ほどで学校に入ってきたものです。\nこれまでの改訂は足す改訂でした。今回は精選する改訂で、減らした先に何を置くかを学校が決める。ここが本日の中心テーマです。");
}

// 19 調整授業時数制度
{
  const s = newSlide();
  head(s, "Ⅱ－6", "制度の核 ―「調整授業時数制度」");
  card(s, M, 1.62, CW, 1.05, INK);
  txt(s, "年間の標準授業時数の総時数は保ったまま、各教科等への配分を学校が調整できる仕組み", {
    x: M + 0.5, y: 1.87, w: CW - 1.0, h: 0.55,
    fontSize: 18, bold: true, color: PAPER,
  });
  // 概念図
  const gy = 3.05;
  txt(s, "これまで", { x: M, y: gy, w: 2.0, h: 0.34, fontSize: 14, bold: true, color: GREY });
  card(s, M + 2.1, gy - 0.04, CW - 2.1, 0.62, MIST, { round: false });
  txt(s, "各教科等の時数は、ほぼ全校共通", {
    x: M + 2.4, y: gy + 0.10, w: CW - 2.7, h: 0.4, fontSize: 14, color: GREY,
  });

  txt(s, "これから", { x: M, y: gy + 1.05, w: 2.0, h: 0.34, fontSize: 14, bold: true, color: PINE });
  const seg = [
    ["各教科等\n（調整後）", 6.2, PINE, PAPER],
    ["調　整\n授業時数", 3.55, GOLDL, INK],
  ];
  let sx = M + 2.1;
  seg.forEach((g) => {
    card(s, sx, gy + 1.01, g[1], 0.72, g[2], { round: false });
    txt(s, g[0], {
      x: sx, y: gy + 1.12, w: g[1], h: 0.55,
      fontSize: 13, bold: true, color: g[3], align: "center", lineSpacing: 17,
    });
    sx += g[1] + 0.08;
  });
  s.addShape(pres.ShapeType.line, {
    x: M + 2.1, y: gy + 1.92, w: 9.83, h: 0,
    line: { color: LINEC, width: 1.5 },
  });
  txt(s, "総時数は変わらない", {
    x: M + 2.1, y: gy + 1.98, w: 9.83, h: 0.32,
    fontSize: 12, color: GREY, align: "center",
  });

  card(s, M, 5.60, CW, 1.12, MIST2);
  txt(s, "減らしてよい枠が制度として用意される ＝ 学校が「何に時間を寄せるか」を決められる", {
    x: M + 0.5, y: 5.88, w: CW - 1.0, h: 0.5,
    fontSize: 17, bold: true, color: INK,
  });
  s.addNotes("制度の核がこれです。総時数そのものは変えません。変えられるのは中の配分です。\n各教科等の時数を一定の範囲で減らし、その分を調整授業時数としてプールする。そのプールを何に使うかは学校が決める、という仕組みです。");
}

// 20 3つの使い道
{
  const s = newSlide();
  head(s, "Ⅱ－7", "調整した時間は、3つに使えます");
  const uses = [
    ["①", "既存教科の充実", "自校の課題に応じて、特定の教科等に時間を上乗せする", "例：読み書きに課題 → 国語に上乗せ", PINE],
    ["②", "学校設定教科の新設", "自校独自の教科等を設ける", "例：地域探究、キャリア、ICT活用", MOSSD],
    ["③", "「裁量的な時間」への充当", "授業以外の、学びの質を高める活動に充てる", "例：個別の学習支援、教員の研究・研修", GOLD],
  ];
  let y = 1.70;
  uses.forEach((u) => {
    card(s, M, y, CW, 1.48, MIST);
    badge(s, M + 0.36, y + 0.42, 0.62, u[0], u[4], PAPER, 20);
    txt(s, u[1], {
      x: M + 1.18, y: y + 0.22, w: 5.4, h: 0.42,
      fontSize: 19, bold: true, color: INK,
    });
    txt(s, u[2], {
      x: M + 1.18, y: y + 0.72, w: 6.2, h: 0.55,
      fontSize: 13.5, color: TEXT,
    });
    card(s, M + 7.8, y + 0.38, CW - 8.2, 0.72, PAPER);
    txt(s, u[3], {
      x: M + 8.0, y: y + 0.52, w: CW - 8.6, h: 0.48,
      fontSize: 12.5, color: PINE,
    });
    y += 1.62;
  });
  txt(s, "③が、これまでの制度になかった発想です。授業以外にも時数を充てられます。", {
    x: M, y: 6.58, w: CW, h: 0.38,
    fontSize: 14, bold: true, color: PINE,
  });
  s.addNotes("プールした時間の使い道が3つ示されています。①と②はイメージしやすいと思います。\n注目していただきたいのは③です。個別の学習支援や、教員の研究・研修といった、授業以外の活動にも時数を充てられる。これは従来の発想にはなかったものです。");
}

// 21 ルールと制約
{
  const s = newSlide();
  head(s, "Ⅱ－8", "ルールと制約");
  const rules = [
    ["総時数は減らさない", "年間の標準授業時数の総枠は維持されます", PINE],
    ["削減の上限が設けられる", "現行の授業時数特例校制度（1割程度）を上回る方向で検討中", GOLD],
    ["年35コマ以下の教科等は対象外", "道徳など、もともと時数の少ない教科等は削減の対象から除かれます", MOSSD],
    ["減らしても年35コマは下回らない", "すべての教科等で週1コマは確保される設計です", MOSSD],
    ["「裁量的な時間」にも上限", "際限なく授業以外に振り向けることはできません", GOLD],
  ];
  let y = 1.66;
  rules.forEach((r, i) => {
    card(s, M, y, CW, 0.92, i % 2 === 0 ? MIST : MIST2);
    badge(s, M + 0.32, y + 0.21, 0.5, String(i + 1), r[2], PAPER, 15);
    txt(s, r[0], {
      x: M + 1.0, y: y + 0.10, w: 4.6, h: 0.4,
      fontSize: 16, bold: true, color: INK,
    });
    txt(s, r[1], {
      x: M + 1.0, y: y + 0.50, w: CW - 1.5, h: 0.36,
      fontSize: 12.5, color: TEXT,
    });
    y += 1.02;
  });
  foot(s, "※具体的な上限値は審議中です。確定値は告示および区教育委員会の通知でご確認ください。");
  s.addNotes("制度である以上、当然ルールがあります。ここは正確におさえておきたいところです。\n特に2番目の上限については、まだ審議中です。現行の特例校制度が1割程度ですので、それを上回る方向という報道がありますが、数字は確定していません。");
}

// 22 裁量的な時間
{
  const s = newSlide();
  head(s, "Ⅱ－9", "「裁量的な時間」とは何か");
  card(s, M, 1.66, CW, 1.0, INK);
  txt(s, "授業のコマとしてではなく、学びの質を高める活動に充てられる時間", {
    x: M + 0.5, y: 1.90, w: CW - 1.0, h: 0.48,
    fontSize: 18, bold: true, color: PAPER,
  });
  const k = [
    ["子供に向けて", ["つまずいている子への個別の学習支援", "興味・関心を伸ばす個別の探究", "多様な背景のある子への支援"], PINE],
    ["教職員に向けて", ["授業改善に直結する研究・研修", "教材研究の時間", "学年・教科をまたいだ打ち合わせ"], MOSSD],
  ];
  k.forEach((g, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw5 = CW / 2 - 0.2;
    card(s, x, 2.88, cw5, 2.55, MIST);
    badge(s, x + 0.38, 3.16, 0.52, i === 0 ? "子" : "師", g[2], PAPER, 15);
    txt(s, g[0], {
      x: x + 1.05, y: 3.20, w: cw5 - 1.4, h: 0.4,
      fontSize: 18, bold: true, color: INK,
    });
    bullets(s, g[1], {
      x: x + 0.42, y: 3.88, w: cw5 - 0.84, h: 1.4, fontSize: 13.5,
    });
  });
  card(s, M, 5.70, CW, 1.05, MIST2);
  txt(s, "教員の研修を「時数」として位置づけられる。これは学校経営上、非常に大きい変化です。", {
    x: M + 0.5, y: 5.96, w: CW - 1.0, h: 0.5,
    fontSize: 17, bold: true, color: PINE,
  });
  s.addNotes("裁量的な時間の中身です。子供に向けたものと、教職員に向けたものの両方が想定されています。\n下に書いたとおり、教員の研修を時数として位置づけられるというのは、学校経営の観点ではかなり大きな変化だと思います。");
}

// 23 情報活用能力
{
  const s = newSlide();
  head(s, "Ⅱ－10", "情報活用能力を、すべての学びの基盤に");
  card(s, M, 1.68, CW, 1.05, INK);
  txt(s, "情報活用能力を、すべての教科等の学びの基盤として位置づける", {
    x: M + 0.5, y: 1.93, w: CW - 1.0, h: 0.5,
    fontSize: 19, bold: true, color: PAPER,
  });
  const flow = [
    ["基礎を計画的に育成", "小：総合の「情報の領域」\n中：「情報・技術科」", PINE],
    ["各教科等の学習で活用", "調べる・まとめる・伝える\n一連の学習過程で使う", MOSSD],
    ["学びの基盤として定着", "読み書き・計算と同じ位置づけ", GOLD],
  ];
  const cw6 = (CW - 1.3) / 3;
  flow.forEach((f, i) => {
    const x = M + i * (cw6 + 0.65);
    card(s, x, 3.05, cw6, 2.0, MIST);
    badge(s, x + cw6 / 2 - 0.3, 3.30, 0.6, String(i + 1), f[2], PAPER, 19);
    txt(s, f[0], {
      x: x + 0.2, y: 4.02, w: cw6 - 0.4, h: 0.42,
      fontSize: 15, bold: true, color: INK, align: "center",
    });
    txt(s, f[1], {
      x: x + 0.2, y: 4.46, w: cw6 - 0.4, h: 0.5,
      fontSize: 12, color: TEXT, align: "center", lineSpacing: 17,
    });
    if (i < 2) {
      txt(s, "▶", {
        x: x + cw6 + 0.08, y: 3.90, w: 0.5, h: 0.4,
        fontSize: 17, color: MOSS, align: "center",
      });
    }
  });
  card(s, M, 5.42, CW, 1.15, MIST2);
  txt(s, "「ICTを使う時間」ではなく「どの授業でも使える力」として育てる、という整理です。", {
    x: M + 0.5, y: 5.72, w: CW - 1.0, h: 0.5,
    fontSize: 17, bold: true, color: INK,
  });
  s.addNotes("情報教育の位置づけが変わります。特定の時間にICTを使うという発想から、読み書き計算と同じように、すべての学びの土台になる力として育てるという整理です。\n基礎を計画的に育てて、各教科で使う。この順番が示されています。");
}

// 24 情報の領域・情報技術科
{
  const s = newSlide();
  head(s, "Ⅱ－11", "小学校・中学校で新たに置かれるもの");
  const sch = [
    ["小学校", "総合的な学習の時間に\n「情報の領域」を設ける", [
      "体験的な活動を重視",
      "中核は情報の「活用」（収集・整理）",
      "情報モラルなど「適切な取扱い」",
      "コンピュータの仕組みなど「特性の理解」",
    ], PINE],
    ["中学校", "「情報・技術科」を新設\n（技術・家庭を分離）", [
      "情報分野を独立した教科として扱う",
      "技術分野と情報分野を体系的に",
      "小学校からの学びを接続",
      "高等学校「情報」へつなぐ",
    ], MOSSD],
  ];
  sch.forEach((c, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw7 = CW / 2 - 0.2;
    card(s, x, 1.70, cw7, 4.42, MIST);
    card(s, x, 1.70, cw7, 0.68, c[3]);
    txt(s, c[0], {
      x, y: 1.86, w: cw7, h: 0.38,
      fontSize: 18, bold: true, color: PAPER, align: "center",
    });
    txt(s, c[1], {
      x: x + 0.35, y: 2.60, w: cw7 - 0.7, h: 0.95,
      fontSize: 17, bold: true, color: INK, lineSpacing: 26,
    });
    bullets(s, c[2], {
      x: x + 0.38, y: 3.72, w: cw7 - 0.76, h: 2.2, fontSize: 13,
    });
  });
  foot(s, "※教科名・領域名および内容は素案段階のものです。");
  s.addNotes("具体的には、小学校は総合的な学習の時間に情報の領域を置きます。中学校は技術・家庭を分けて、情報・技術科を新設します。\n小学校の先生方にとっては、総合の設計に直接関わってくる話です。ここは後ほどⅣ章でも触れます。");
}

// 25 言語能力・外国語
{
  const s = newSlide();
  head(s, "Ⅱ－12", "言語能力と外国語科");
  card(s, M, 1.70, CW, 2.1, MIST);
  badge(s, M + 0.4, 2.02, 0.56, "言", PINE, PAPER, 17);
  txt(s, "学習の基盤としての言語活動の充実", {
    x: M + 1.1, y: 2.06, w: CW - 1.6, h: 0.42,
    fontSize: 20, bold: true, color: INK,
  });
  txt(s, "国語科だけの話ではありません。各教科等の中で、言葉を使って考え、まとめ、伝える活動を充実させます。", {
    x: M + 1.1, y: 2.58, w: CW - 1.6, h: 0.44,
    fontSize: 13.5, color: TEXT,
  });
  card(s, M + 1.1, 3.06, CW - 1.6, 0.6, PAPER);
  txt(s, "例：理科の実験レポートの作成　／　社会科の資料をもとにした説明　／　算数の考え方の記述", {
    x: M + 1.35, y: 3.20, w: CW - 2.1, h: 0.4,
    fontSize: 13, color: PINE,
  });

  card(s, M, 4.00, CW, 2.1, MIST2);
  badge(s, M + 0.4, 4.32, 0.56, "外", MOSSD, PAPER, 17);
  txt(s, "外国語科の改善", {
    x: M + 1.1, y: 4.36, w: CW - 1.6, h: 0.42,
    fontSize: 20, bold: true, color: INK,
  });
  txt(s, "新たに整理したコミュニケーション活動の中で、次の活動を充実させます。", {
    x: M + 1.1, y: 4.88, w: CW - 1.6, h: 0.4,
    fontSize: 13.5, color: TEXT,
  });
  card(s, M + 1.1, 5.32, CW - 1.6, 0.6, PAPER);
  txt(s, "自分の意見を形成・発信する活動　／　身近な地域のことを発信する活動", {
    x: M + 1.35, y: 5.46, w: CW - 2.1, h: 0.4,
    fontSize: 13, color: PINE,
  });
  txt(s, "いずれも「使って考える」方向。Ⅱ－3の両輪とつながっています。", {
    x: M, y: 6.34, w: CW, h: 0.38,
    fontSize: 14, bold: true, color: PINE,
  });
  s.addNotes("言語能力については、国語科だけの話ではないという点が重要です。理科のレポート、社会の説明、算数の記述。各教科で言葉を使う場面を充実させます。\n外国語は、自分の意見を発信する、地域のことを発信する、という方向が明示されました。");
}

// 26 学習評価
{
  const s = newSlide();
  head(s, "Ⅱ－13", "学習評価も見直されます");
  const ev = [
    ["観点別評価は継続、ただし深化", "知識・技能／思考・判断・表現／主体的に学習に取り組む態度 の3観点は維持"],
    ["「主体的に学習に取り組む態度」の扱いを見直し", "他の観点から切り離してA・B・Cで評定する仕組みが見直しの対象に"],
    ["所見による個別のフィードバックへ", "記号や数値ではなく、その子に向けた言葉で返す方向"],
    ["評価方法の多様化", "知識再生中心の筆記試験偏重から、多様な学びを評価できる仕組みへ"],
  ];
  let y = 1.62;
  ev.forEach((e, i) => {
    card(s, M, y, CW, 1.0, i % 2 === 0 ? MIST : MIST2);
    badge(s, M + 0.34, y + 0.25, 0.5, String(i + 1), PINE, PAPER, 15);
    txt(s, e[0], {
      x: M + 1.02, y: y + 0.14, w: CW - 1.5, h: 0.4,
      fontSize: 16.5, bold: true, color: INK,
    });
    txt(s, e[1], {
      x: M + 1.02, y: y + 0.56, w: CW - 1.5, h: 0.38,
      fontSize: 13, color: TEXT,
    });
    y += 1.10;
  });
  card(s, M, 6.08, CW, 0.78, INK);
  txt(s, "評価が変わると、授業が変わります。ここは校内研修の題材として最適です。", {
    x: M + 0.5, y: 6.28, w: CW - 1.0, h: 0.42,
    fontSize: 16, bold: true, color: PAPER,
  });
  s.addNotes("評価も見直されます。特に2つめ、主体的に学習に取り組む態度をA・B・Cで切り離して評価する仕組み。ここに手が入る見込みです。\n評価が変わると授業が変わりますので、校内研修の題材としては一番入りやすいところだと思います。");
}

/* ══════════════════════════════════════════════════
   Ⅲ 校長の権限はどこまで届くか（27–34）
   ══════════════════════════════════════════════════ */

// 27 扉
section("Ⅲ", "校長の権限はどこまで届くか", "できること・できないことを、はっきりさせます", [
  "教育課程を編成するのは「学校」です",
  "校長の判断だけで動かせる領域は、思っているより広い",
  "同時に、権限外のことも正確におさえておく",
]).addNotes("ここは整理の章です。できることとできないことを分けて考えます。");

// 28 大前提
{
  const s = newSlide();
  head(s, "Ⅲ－1", "大前提 ― 教育課程を編成するのは「学校」です");
  card(s, M, 1.70, CW, 1.55, INK);
  txt(s, "校長は、校務をつかさどり、所属職員を監督する。", {
    x: M + 0.55, y: 2.02, w: CW - 1.1, h: 0.52,
    fontSize: 24, bold: true, color: PAPER,
  });
  txt(s, "学校教育法 第37条第4項（中学校は第49条で準用）", {
    x: M + 0.55, y: 2.62, w: CW - 1.1, h: 0.34,
    fontSize: 12.5, color: MOSS,
  });
  const chain = [
    ["国", "学習指導要領で\n大綱的な基準を示す", GREY],
    ["区教育委員会", "管理運営の基準を定め、\n指導・助言する", MOSSD],
    ["学校（校長）", "自校の教育課程を\n編成し、実施する", PINE],
  ];
  const cw8 = (CW - 1.4) / 3;
  chain.forEach((c, i) => {
    const x = M + i * (cw8 + 0.7);
    card(s, x, 3.62, cw8, 1.82, i === 2 ? MIST2 : MIST);
    txt(s, c[0], {
      x: x + 0.2, y: 3.92, w: cw8 - 0.4, h: 0.44,
      fontSize: 19, bold: true, color: c[2], align: "center",
    });
    txt(s, c[1], {
      x: x + 0.2, y: 4.48, w: cw8 - 0.4, h: 0.75,
      fontSize: 13, color: TEXT, align: "center", lineSpacing: 19,
    });
    if (i < 2) {
      txt(s, "▶", {
        x: x + cw8 + 0.1, y: 4.38, w: 0.5, h: 0.4,
        fontSize: 18, color: MOSS, align: "center",
      });
    }
  });
  card(s, M, 5.72, CW, 1.0, MIST);
  txt(s, "国が決めるのは「大綱的な基準」。中身をどう組むかは、学校に委ねられています。", {
    x: M + 0.5, y: 5.97, w: CW - 1.0, h: 0.5,
    fontSize: 17, bold: true, color: PINE,
  });
  s.addNotes("当たり前のことですが、あらためて確認しておきたいところです。教育課程を編成するのは学校であって、国でも区でもありません。\n国が示すのは大綱的な基準です。その中身をどう組むかは、もともと学校に委ねられています。");
}

// 29 権限マップ
{
  const s = newSlide();
  head(s, "Ⅲ－2", "校長権限マップ");
  const map = [
    ["校長の判断で\nできる", PINE, MIST2, [
      "日課表・週時程の編成",
      "学校行事の精選・統合",
      "校内研修の内容と回数",
      "校務分掌の組み替え",
      "学校経営計画の重点設定",
      "授業改善の方針",
      "会議の持ち方・時間",
    ]],
    ["区教委との\n協議・届出", MOSSD, MIST, [
      "教育課程の届出・報告",
      "標準授業時数の特例的な運用",
      "学校設定教科等の新設",
      "長期休業日の変更",
      "特色ある教育課程の申請",
    ]],
    ["校長の\n権限外", GREY, MIST, [
      "教科書の採択（区教委の権限）",
      "学習指導要領の内容",
      "教職員定数・人事",
      "標準授業時数の総枠（法令）",
    ]],
  ];
  const cw9 = (CW - 0.5) / 3;
  map.forEach((m, i) => {
    const x = M + i * (cw9 + 0.25);
    card(s, x, 1.66, cw9, 4.46, m[2]);
    card(s, x, 1.66, cw9, 0.88, m[1]);
    txt(s, m[0], {
      x, y: 1.78, w: cw9, h: 0.66,
      fontSize: 16, bold: true, color: PAPER, align: "center", lineSpacing: 21,
    });
    bullets(s, m[3], {
      x: x + 0.3, y: 2.76, w: cw9 - 0.6, h: 3.2,
      fontSize: 12.5, color: i === 2 ? GREY : TEXT,
    });
  });
  foot(s, "※中央・右列の線引きは自治体の管理運営規則によって異なります。実施前に区教育委員会にご確認ください。");
  s.addNotes("左の列が、今日いちばん見ていただきたいところです。この7つは、原則として校長の判断だけで動かせます。\n真ん中は協議が要るもの、右は権限外のものです。中央と右の線引きは自治体で違いますので、区教委への確認は必要です。");
}

// 30 根拠法令
{
  const s = newSlide();
  head(s, "Ⅲ－3", "根拠となる主な法令");
  const laws = [
    ["校務の掌理", "学校教育法 第37条第4項", "校長は、校務をつかさどり、所属職員を監督する"],
    ["教育課程の編成", "学校教育法施行規則 第50条・第72条", "小・中学校の教育課程を構成する教科等"],
    ["授業時数", "同 第51条・第73条／別表第一・第二", "各学年の年間標準授業時数"],
    ["校務分掌", "同 第43条", "ふさわしい校務分掌の仕組みを整える"],
    ["研修", "教育公務員特例法 第21条・第22条", "研修の機会が与えられなければならない"],
    ["学校評価", "学校教育法施行規則 第66条～第68条", "自己評価・学校関係者評価・設置者への報告"],
    ["学校運営協議会", "地方教育行政法 第47条の5", "学校運営に関する協議・意見の申出"],
  ];
  let y = 1.62;
  laws.forEach((l, i) => {
    card(s, M, y, CW, 0.66, i % 2 === 0 ? MIST : PAPER);
    txt(s, l[0], {
      x: M + 0.32, y: y + 0.17, w: 2.2, h: 0.36,
      fontSize: 14, bold: true, color: PINE,
    });
    txt(s, l[1], {
      x: M + 2.65, y: y + 0.17, w: 3.9, h: 0.36,
      fontSize: 13, bold: true, color: INK,
    });
    txt(s, l[2], {
      x: M + 6.70, y: y + 0.18, w: CW - 7.0, h: 0.36,
      fontSize: 12.5, color: TEXT,
    });
    y += 0.72;
  });
  foot(s, "※令和8年9月時点。条文番号は改正により変わることがあります。");
  s.addNotes("根拠を並べておきました。お手元で確認されるときの索引としてお使いください。\n特に校務分掌の第43条と、研修の教特法21条・22条。この2つは、Ⅳ章の打ち手の裏づけになります。");
}

// 31 できること①
{
  const s = newSlide();
  head(s, "Ⅲ－4", "できること① 教育課程・時間割・日課表");
  card(s, M, 1.68, CW, 0.88, PINE);
  txt(s, "「何時間やるか」より先に、「どう並べるか」で学びは変わります", {
    x: M + 0.5, y: 1.90, w: CW - 1.0, h: 0.46,
    fontSize: 18, bold: true, color: PAPER,
  });
  const items = [
    ["単位時間の設定", "45分・40分など、単位時間の設定は学校の判断で検討できます（要協議）"],
    ["モジュールの活用", "15分×3などの短時間学習を、朝の時間や帯に組み込む"],
    ["週時程の組み替え", "探究や行事に使える、まとまった時間をあらかじめ確保する"],
    ["学年・教科の横断", "同じ時間帯に同一教科を置き、学年内で柔軟に動かせるようにする"],
    ["時間割の可変枠", "学期に数コマ、用途を決めない枠を先に空けておく"],
  ];
  let y = 2.76;
  items.forEach((it, i) => {
    card(s, M, y, CW, 0.74, i % 2 === 0 ? MIST : MIST2);
    badge(s, M + 0.30, y + 0.17, 0.4, "", MOSSD, PAPER, 11);
    txt(s, it[0], {
      x: M + 0.88, y: y + 0.19, w: 3.3, h: 0.38,
      fontSize: 15, bold: true, color: INK,
    });
    txt(s, it[1], {
      x: M + 4.35, y: y + 0.20, w: CW - 4.65, h: 0.38,
      fontSize: 13, color: TEXT,
    });
    y += 0.82;
  });
  s.addNotes("まず時間割です。総時数を変えなくても、並べ方を変えるだけでできることがあります。\n最後の可変枠は特におすすめです。学期に数コマ、用途を決めない枠を先に空けておく。これだけで年度途中の対応力が変わります。");
}

// 32 できること②
{
  const s = newSlide();
  head(s, "Ⅲ－5", "できること② 学校行事・特別活動の精選");
  card(s, M, 1.68, CW, 0.88, PINE);
  txt(s, "行事は、校長の判断で最も動かしやすい領域です", {
    x: M + 0.5, y: 1.90, w: CW - 1.0, h: 0.46,
    fontSize: 18, bold: true, color: PAPER,
  });
  const ax = [
    ["目的", "この行事は、どの資質・能力を育てるためのものか"],
    ["時間", "本番だけでなく、準備・練習にどれだけ使っているか"],
    ["負担", "担当者と子供に、どれだけの負荷がかかっているか"],
    ["代替", "他の行事や授業で、同じねらいを達成できないか"],
  ];
  txt(s, "4つの軸で棚卸しをします", {
    x: M, y: 2.78, w: CW, h: 0.36,
    fontSize: 15, bold: true, color: GREY,
  });
  const cw10 = (CW - 0.75) / 4;
  ax.forEach((a, i) => {
    const x = M + i * (cw10 + 0.25);
    card(s, x, 3.24, cw10, 1.72, MIST);
    badge(s, x + cw10 / 2 - 0.28, 3.48, 0.56, a[0], PINE, PAPER, 15);
    txt(s, a[1], {
      x: x + 0.22, y: 4.22, w: cw10 - 0.44, h: 0.62,
      fontSize: 12.5, color: TEXT, align: "center", lineSpacing: 18,
    });
  });
  card(s, M, 5.28, CW, 1.42, MIST2);
  txt(s, "「やめる」だけが精選ではありません", {
    x: M + 0.5, y: 5.52, w: CW - 1.0, h: 0.4,
    fontSize: 17, bold: true, color: INK,
  });
  txt(s, "統合する／規模を小さくする／準備時間を減らす／隔年開催にする／子供に任せる範囲を広げる。手段は5つ以上あります。", {
    x: M + 0.5, y: 5.96, w: CW - 1.0, h: 0.5,
    fontSize: 13.5, color: TEXT,
  });
  s.addNotes("行事は、校長の判断で最も動かしやすい領域です。ただ、いきなり減らすと必ず反発が出ます。\nこの4つの軸で棚卸しをして、事実を並べてから議論する。あと、やめる以外にも手段があるということを最初に共有しておくと、話が進みやすくなります。");
}

// 33 できること③
{
  const s = newSlide();
  head(s, "Ⅲ－6", "できること③ 校内研修の設計");
  card(s, M, 1.68, CW, 0.92, PINE);
  txt(s, "素案は、教員の研究・研修を「時数を充てられる活動」として位置づけています", {
    x: M + 0.5, y: 1.92, w: CW - 1.0, h: 0.48,
    fontSize: 18, bold: true, color: PAPER,
  });
  const ba = [
    ["これまでの校内研修", GREY, MIST, [
      "回数をこなすことが目的化",
      "全員が同じ内容を受ける",
      "放課後に押し込まれる",
      "研究発表がゴールになる",
    ]],
    ["これからの校内研修", PINE, MIST2, [
      "授業改善に直結する内容に絞る",
      "学年・教科ごとに必要なものを",
      "裁量的な時間として位置づける",
      "日々の授業がゴール",
    ]],
  ];
  ba.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw11 = CW / 2 - 0.2;
    card(s, x, 2.78, cw11, 2.95, b[2]);
    txt(s, b[0], {
      x: x + 0.38, y: 3.02, w: cw11 - 0.76, h: 0.42,
      fontSize: 18, bold: true, color: b[1],
    });
    bullets(s, b[3], {
      x: x + 0.4, y: 3.62, w: cw11 - 0.8, h: 1.95,
      fontSize: 13.5, color: i === 0 ? GREY : TEXT,
    });
  });
  card(s, M, 5.92, CW, 0.82, MIST);
  txt(s, "根拠：教育公務員特例法 第21条・第22条／校内研修の設計は校長の裁量の範囲です。", {
    x: M + 0.5, y: 6.14, w: CW - 1.0, h: 0.42,
    fontSize: 14, bold: true, color: PINE,
  });
  s.addNotes("研修は、今回の素案でかなり位置づけが上がりました。裁量的な時間として時数を充てられるという整理は、これまでになかったものです。\n逆に言えば、回数をこなすだけの研修は、これからは説明がつきにくくなります。");
}

// 34 できること④
{
  const s = newSlide();
  head(s, "Ⅲ－7", "できること④ 校務分掌と組織");
  card(s, M, 1.68, CW, 0.92, PINE);
  txt(s, "「調和のとれた学校運営が行われるためにふさわしい校務分掌の仕組みを整える」", {
    x: M + 0.5, y: 1.92, w: CW - 1.0, h: 0.48,
    fontSize: 17, bold: true, color: PAPER,
  });
  txt(s, "学校教育法施行規則 第43条", {
    x: M + 0.5, y: 2.34, w: CW - 1.0, h: 0.3,
    fontSize: 11, color: "C9DACD",
  });
  const org = [
    ["「教育課程検討」の担当を置く", "既存の分掌に埋もれさせず、独立した担当を明示する", PINE],
    ["学年主任・教科主任を設計の主体に", "管理職が作った案を下ろすのではなく、主任層に描かせる", MOSSD],
    ["若手を入れる", "2030年に中堅になる層を、設計段階から巻き込む", MOSSD],
    ["分掌の数そのものを見直す", "「余白」をつくるには、まず分掌の棚卸しから", GOLD],
  ];
  let y = 2.80;
  org.forEach((o) => {
    card(s, M, y, CW, 0.90, MIST);
    badge(s, M + 0.32, y + 0.20, 0.5, "◆", o[2], PAPER, 13);
    txt(s, o[0], {
      x: M + 1.0, y: y + 0.11, w: CW - 1.4, h: 0.4,
      fontSize: 16, bold: true, color: INK,
    });
    txt(s, o[1], {
      x: M + 1.0, y: y + 0.51, w: CW - 1.4, h: 0.36,
      fontSize: 13, color: TEXT,
    });
    y += 0.97;
  });
  txt(s, "誰に考えさせるかを決めるのは、校長にしかできない仕事です。", {
    x: M, y: 6.72, w: CW - 1.2, h: 0.36,
    fontSize: 15, bold: true, color: PINE,
  });
  s.addNotes("組織づくりです。ここは校長にしかできません。\n特に3つめ、若手を入れるという点です。2030年に全面実施を迎えるとき、今の若手が中堅になっています。その層を設計段階から入れておくかどうかで、定着の仕方が変わります。");
}

/* ══════════════════════════════════════════════════
   Ⅳ 明日から打てる手（35–46）
   ══════════════════════════════════════════════════ */

// 35 扉
section("Ⅳ", "明日から打てる手", "先行事例と、自校でできる10の打ち手", [
  "すでに先行して取り組んでいる自治体があります",
  "告示を待たずに、今の法令のままでできることがあります",
  "10の打ち手のうち、まず1つ選んでください",
]).addNotes("ここからが本題です。具体的に何をするか、という話に入ります。");

// 36 発想の転換
{
  const s = newSlide();
  head(s, "Ⅳ－1", "発想の転換 ―「増やす」から「組み替える」へ");
  const flow = [
    ["これまでの発想", "新しいことが求められた\n　↓\n時間を足す・行事を足す\n　↓\n現場が回らなくなる", GREY, MIST],
    ["これからの発想", "自校の重点を決める\n　↓\n重点以外を思い切って削る\n　↓\n削った時間を重点に寄せる", PINE, MIST2],
  ];
  flow.forEach((f, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw12 = CW / 2 - 0.2;
    card(s, x, 1.70, cw12, 3.1, f[3]);
    txt(s, f[0], {
      x: x + 0.4, y: 1.98, w: cw12 - 0.8, h: 0.42,
      fontSize: 18, bold: true, color: f[2],
    });
    txt(s, f[1], {
      x: x + 0.4, y: 2.58, w: cw12 - 0.8, h: 1.95,
      fontSize: 15, color: i === 0 ? GREY : TEXT, lineSpacing: 28,
    });
  });
  card(s, M, 5.00, CW, 1.72, INK);
  txt(s, "「何を減らすか」を決められるのは、校長だけです。", {
    x: M + 0.55, y: 5.28, w: CW - 1.1, h: 0.5,
    fontSize: 22, bold: true, color: PAPER,
  });
  txt(s, "職員は、自分の担当を自分から減らすとは言い出しにくいものです。減らす判断を引き受けることが、\n結果として職員に余白を渡すことになります。", {
    x: M + 0.55, y: 5.86, w: CW - 1.1, h: 0.7,
    fontSize: 13.5, color: "AFC5B6", lineSpacing: 21,
  });
  s.addNotes("発想の転換がいります。これまでは足す発想でした。これからは、重点を決めて、それ以外を削って、削った分を重点に寄せる。\n率直に申し上げて、減らす判断は職員からは出てきません。担当している人ほど言い出しにくい。ここを引き受けるのが校長の仕事だと思います。");
}

// 37 先行事例①
{
  const s = newSlide();
  head(s, "Ⅳ－2", "先行事例① 渋谷区 ― 授業時数特例校制度の活用");
  card(s, M, 1.68, CW, 1.28, MIST2);
  badge(s, M + 0.42, 1.98, 0.66, "渋", PINE, PAPER, 20);
  txt(s, "各教科の授業時数を1割程度削減し、探究的な学習に割り当てる", {
    x: M + 1.28, y: 2.08, w: CW - 1.8, h: 0.5,
    fontSize: 20, bold: true, color: INK,
  });
  const pts = [
    ["使った制度", "授業時数特例校制度（現行制度）"],
    ["削減の幅", "各教科 1割程度"],
    ["振り向け先", "探究的な学習"],
    ["示していること", "今の制度のままでも、すでに実行できる"],
  ];
  let y = 3.02;
  pts.forEach((p, i) => {
    card(s, M, y, CW, 0.70, i === 3 ? MIST2 : MIST);
    txt(s, p[0], {
      x: M + 0.35, y: y + 0.19, w: 3.0, h: 0.38,
      fontSize: 14, bold: true, color: i === 3 ? GOLD : PINE,
    });
    txt(s, p[1], {
      x: M + 3.55, y: y + 0.19, w: CW - 3.9, h: 0.38,
      fontSize: 15, bold: i === 3, color: INK,
    });
    y += 0.76;
  });
  card(s, M, 6.30, CW, 0.68, INK);
  txt(s, "調整授業時数制度は、この延長線上にあります。先行事例はすでにあります。", {
    x: M + 0.5, y: 6.47, w: CW - 1.0, h: 0.38,
    fontSize: 16, bold: true, color: PAPER,
  });
  s.addNotes("先行事例を2つ紹介します。まず渋谷区です。授業時数特例校制度という現行の制度を使って、各教科を1割程度削って、探究に回しています。\n大事なのは、これが今の制度のままでできているという点です。新しい制度を待たなくても、すでに道はあるということです。");
}

// 38 先行事例②
{
  const s = newSlide();
  head(s, "Ⅳ－3", "先行事例② 目黒区 ― 単位時間の短縮");
  card(s, M, 1.68, CW, 1.28, MIST2);
  badge(s, M + 0.42, 1.98, 0.66, "目", MOSSD, PAPER, 20);
  txt(s, "小学校の1単位時間を45分から40分に短縮し、積み上げた時間を活用", {
    x: M + 1.28, y: 2.08, w: CW - 1.8, h: 0.5,
    fontSize: 20, bold: true, color: INK,
  });
  // 5分の積み上げ
  card(s, M, 3.16, 5.6, 1.5, MIST);
  txt(s, "1コマ 5分 × 積み重ね", {
    x: M + 0.35, y: 3.40, w: 5.0, h: 0.38,
    fontSize: 15, bold: true, color: GREY,
  });
  txt(s, "日々のわずかな短縮が、年間ではまとまった時間になります。", {
    x: M + 0.35, y: 3.86, w: 5.0, h: 0.62,
    fontSize: 13, color: TEXT, lineSpacing: 19,
  });
  txt(s, "▶", {
    x: M + 5.75, y: 3.72, w: 0.6, h: 0.4,
    fontSize: 20, color: MOSS, align: "center",
  });
  const uses2 = [
    "特色ある活動",
    "教材研究",
    "教員の研修",
  ];
  let ux = M + 6.5;
  uses2.forEach((u) => {
    card(s, ux, 3.16, 1.72, 1.5, PINE);
    txt(s, u, {
      x: ux + 0.10, y: 3.70, w: 1.52, h: 0.44,
      fontSize: 14, bold: true, color: PAPER, align: "center",
    });
    ux += 1.84;
  });
  card(s, M, 4.92, CW, 0.9, MIST);
  txt(s, "使った制度：研究開発学校制度　／　振り向け先に「教材研究」「研修」が入っている点に注目", {
    x: M + 0.45, y: 5.14, w: CW - 0.9, h: 0.46,
    fontSize: 14, color: TEXT,
  });
  card(s, M, 6.00, CW, 0.82, INK);
  txt(s, "「子供のため」と「教職員の余白」は、両立させられます。", {
    x: M + 0.5, y: 6.22, w: CW - 1.0, h: 0.42,
    fontSize: 17, bold: true, color: PAPER,
  });
  s.addNotes("もう一つは目黒区です。1単位時間を45分から40分にして、積み上がった時間を特色ある活動と、教材研究、研修に充てています。\n注目していただきたいのは、振り向け先に教材研究と研修が入っていることです。子供のためと教職員の余白は、両立させられるということを示しています。");
}

// 39 10の打ち手
{
  const s = newSlide();
  head(s, "Ⅳ－4", "自校でできる10の打ち手");
  const moves = [
    ["①", "学校行事の棚卸し", "目的・時間・負担を一覧化する"],
    ["②", "統合・縮減の判断基準づくり", "感覚ではなく基準で決める"],
    ["③", "前年度踏襲の見直し会議", "年内に1回、必ず開く"],
    ["④", "日課表・単位時間の点検", "45分／40分・モジュール"],
    ["⑤", "週時程の組み替え", "まとまった時間を先に確保"],
    ["⑥", "総合的な学習の時間の再設計", "「情報の領域」を見据える"],
    ["⑦", "情報活用能力の系統表づくり", "自校版を1枚つくる"],
    ["⑧", "校内研修の絞り込み", "授業改善に直結するものへ"],
    ["⑨", "会議の精選", "回数・時間・参加者を見直す"],
    ["⑩", "評価と所見の見直し", "記号から言葉へ"],
  ];
  const cwm = (CW - 0.3) / 2;
  moves.forEach((m, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * (cwm + 0.3);
    const y = 1.62 + row * 1.02;
    card(s, x, y, cwm, 0.9, row % 2 === 0 ? MIST : MIST2);
    badge(s, x + 0.24, y + 0.21, 0.48, m[0], PINE, PAPER, 14);
    txt(s, m[1], {
      x: x + 0.86, y: y + 0.13, w: cwm - 1.1, h: 0.36,
      fontSize: 15, bold: true, color: INK,
    });
    txt(s, m[2], {
      x: x + 0.86, y: y + 0.50, w: cwm - 1.1, h: 0.32,
      fontSize: 12, color: GREY,
    });
  });
  txt(s, "全部やる必要はありません。まず1つ選んで、令和9年度の経営計画に書いてください。", {
    x: M, y: 6.82, w: CW, h: 0.4,
    fontSize: 15, bold: true, color: PINE,
  });
  s.addNotes("10並べましたが、全部やる必要はありません。むしろ全部やろうとすると失敗します。\nこの中から自校の状況に合うものを1つか2つ選んで、経営計画に書いていただく。それが今日のお願いです。");
}

// 40 打ち手①②③
{
  const s = newSlide();
  head(s, "Ⅳ－5", "打ち手①②③ ― 行事を棚卸しする");
  card(s, M, 1.62, CW, 0.78, PINE);
  txt(s, "いちばん着手しやすく、いちばん効果が見えやすい打ち手です", {
    x: M + 0.5, y: 1.80, w: CW - 1.0, h: 0.42,
    fontSize: 16, bold: true, color: PAPER,
  });
  const steps = [
    ["①", "一覧化する", "全行事を、本番の時数・準備の時数・担当者数まで書き出す。まず事実を揃えます。", "所要：1〜2週間"],
    ["②", "基準を決める", "「ねらいが他で代替できるものは統合」など、判断基準を先に共有する。個別の好き嫌いの話にしないための工夫です。", "所要：職員会議1回"],
    ["③", "年内に決める", "年明けでは翌年度の計画に間に合いません。12月までに結論を出します。", "期限：12月"],
  ];
  let y = 2.62;
  steps.forEach((st) => {
    card(s, M, y, CW, 1.28, MIST);
    badge(s, M + 0.36, y + 0.32, 0.62, st[0], PINE, PAPER, 20);
    txt(s, st[1], {
      x: M + 1.18, y: y + 0.18, w: 2.8, h: 0.4,
      fontSize: 17, bold: true, color: INK,
    });
    txt(s, st[2], {
      x: M + 1.18, y: y + 0.62, w: CW - 3.7, h: 0.55,
      fontSize: 13, color: TEXT,
    });
    card(s, M + CW - 2.25, y + 0.42, 1.95, 0.5, MIST2);
    txt(s, st[3], {
      x: M + CW - 2.25, y: y + 0.55, w: 1.95, h: 0.3,
      fontSize: 12, bold: true, color: PINE, align: "center",
    });
    y += 1.42;
  });
  s.addNotes("①から③は、行事の棚卸しです。ポイントは順番で、先に事実を揃えてから、基準を決めて、それから議論する。\n基準を先に決めるのは、個別の行事の好き嫌いの話にしないためです。ここを飛ばすと必ず紛糾します。");
}

// 41 打ち手④⑤
{
  const s = newSlide();
  head(s, "Ⅳ－6", "打ち手④⑤ ― 時間の器を組み替える");
  const box = [
    ["④", "日課表・単位時間の点検", [
      "45分を前提にしていないか確認する",
      "朝の時間・帯の時間の使い方を見直す",
      "モジュール（15分×3など）の可能性を検討",
      "※単位時間の変更は区教委と要協議",
    ], PINE],
    ["⑤", "週時程の組み替え", [
      "探究に使える2コマ続きを先に確保する",
      "同一学年・同一教科を同じ時間帯に置く",
      "学期に数コマ、用途未定の枠を空けておく",
      "学年で融通できる仕組みにする",
    ], MOSSD],
  ];
  box.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw13 = CW / 2 - 0.2;
    card(s, x, 1.68, cw13, 3.5, MIST);
    badge(s, x + 0.38, 1.98, 0.62, b[0], b[3], PAPER, 20);
    txt(s, b[1], {
      x: x + 1.16, y: 2.08, w: cw13 - 1.5, h: 0.44,
      fontSize: 17, bold: true, color: INK,
    });
    bullets(s, b[2], {
      x: x + 0.42, y: 2.82, w: cw13 - 0.84, h: 2.2,
      fontSize: 13, color: TEXT,
    });
  });
  card(s, M, 5.40, CW, 1.32, MIST2);
  txt(s, "総時数を変えなくても、ここまではできます。", {
    x: M + 0.5, y: 5.64, w: CW - 1.0, h: 0.44,
    fontSize: 19, bold: true, color: INK,
  });
  txt(s, "制度が変わるのを待たずに、令和9年度の時間割から着手できる領域です。", {
    x: M + 0.5, y: 6.12, w: CW - 1.0, h: 0.4,
    fontSize: 14, color: TEXT,
  });
  s.addNotes("④と⑤は、時間の器の話です。ここで強調したいのは、総時数を一切変えなくても、ここまではできるということです。\n制度を待つ必要がありません。令和9年度の時間割づくりから、すぐ着手できます。");
}

// 42 打ち手⑥⑦
{
  const s = newSlide();
  head(s, "Ⅳ－7", "打ち手⑥⑦ ― 総合と情報を先に描く");
  card(s, M, 1.62, CW, 0.78, PINE);
  txt(s, "小学校は総合に「情報の領域」、中学校は「情報・技術科」。影響が大きい領域です。", {
    x: M + 0.5, y: 1.80, w: CW - 1.0, h: 0.42,
    fontSize: 16, bold: true, color: PAPER,
  });
  const six = [
    ["⑥", "総合的な学習の時間の再設計", [
      "今の単元が「探究」になっているか点検する",
      "調べ学習の消化になっていないか",
      "情報の領域が入る前提で枠を空けておく",
      "地域の素材を洗い出しておく",
    ]],
    ["⑦", "情報活用能力の系統表づくり", [
      "6年間（3年間）で何を育てるか1枚にする",
      "各教科のどの単元で使うかを紐づける",
      "完成版でなくてよい。たたき台で十分",
      "つくる過程そのものが校内研修になる",
    ]],
  ];
  six.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw14 = CW / 2 - 0.2;
    card(s, x, 2.62, cw14, 3.15, i === 0 ? MIST : MIST2);
    badge(s, x + 0.38, 2.90, 0.62, b[0], PINE, PAPER, 20);
    txt(s, b[1], {
      x: x + 1.16, y: 2.98, w: cw14 - 1.5, h: 0.48,
      fontSize: 16.5, bold: true, color: INK,
    });
    bullets(s, b[2], {
      x: x + 0.42, y: 3.72, w: cw14 - 0.84, h: 1.9,
      fontSize: 13, color: TEXT,
    });
  });
  card(s, M, 5.96, CW, 0.82, INK);
  txt(s, "⑦は「つくる過程が研修になる」のが利点です。成果物より、対話が残ります。", {
    x: M + 0.5, y: 6.16, w: CW - 1.0, h: 0.44,
    fontSize: 16, bold: true, color: PAPER,
  });
  s.addNotes("⑥と⑦です。特に⑦の系統表づくりをおすすめします。\n完成度の高いものを作る必要はありません。たたき台で十分です。むしろ、作る過程で先生方が話し合うこと自体が研修になります。成果物より対話が残ります。");
}

// 43 打ち手⑧⑨
{
  const s = newSlide();
  head(s, "Ⅳ－8", "打ち手⑧⑨ ― 教職員の時間を取り戻す");
  const eight = [
    ["⑧", "校内研修の絞り込み", "年間の研修を洗い出し、「明日の授業が変わるか」で仕分ける。変わらないものは思い切ってやめる。", "素案の「裁量的な時間」に接続します"],
    ["⑨", "会議の精選", "職員会議・打合せ・委員会の回数と時間、参加者を見直す。文書共有で足りるものは会議にしない。", "生み出した時間を教材研究に回します"],
  ];
  let y = 1.68;
  eight.forEach((e) => {
    card(s, M, y, CW, 1.92, MIST);
    badge(s, M + 0.40, y + 0.42, 0.72, e[0], PINE, PAPER, 23);
    txt(s, e[1], {
      x: M + 1.32, y: y + 0.24, w: CW - 1.8, h: 0.46,
      fontSize: 19, bold: true, color: INK,
    });
    txt(s, e[2], {
      x: M + 1.32, y: y + 0.78, w: CW - 1.8, h: 0.55,
      fontSize: 14, color: TEXT,
    });
    card(s, M + 1.32, y + 1.36, 6.4, 0.42, MIST2);
    txt(s, e[3], {
      x: M + 1.52, y: y + 1.44, w: 6.1, h: 0.3,
      fontSize: 12.5, bold: true, color: PINE,
    });
    y += 2.08;
  });
  card(s, M, 5.96, CW, 0.85, INK);
  txt(s, "職員が「校長は本気で時間をつくろうとしている」と感じると、他の話が通りやすくなります。", {
    x: M + 0.5, y: 6.18, w: CW - 1.0, h: 0.44,
    fontSize: 16, bold: true, color: PAPER,
  });
  s.addNotes("⑧と⑨は、教職員の時間の話です。実務的な効果もありますが、それ以上に大きいのは信頼の面です。\n校長が本気で時間をつくろうとしている、と職員が感じると、その後の改革の話が通りやすくなります。順番としては、ここを先にやる手もあります。");
}

// 44 打ち手⑩
{
  const s = newSlide();
  head(s, "Ⅳ－9", "打ち手⑩ ― 評価を校内研修の題材にする");
  card(s, M, 1.68, CW, 1.0, PINE);
  txt(s, "評価の見直しは、素案の中で最も授業に近いところにあります", {
    x: M + 0.5, y: 1.94, w: CW - 1.0, h: 0.48,
    fontSize: 18, bold: true, color: PAPER,
  });
  const q = [
    "「主体的に学習に取り組む態度」を、自校ではどう見取っているか",
    "その見取り方は、子供にどう伝わっているか",
    "所見は、その子に向けた言葉になっているか",
    "テスト以外に、どんな評価の手立てを持っているか",
  ];
  txt(s, "校内研修で、この4つを問いにしてみてください", {
    x: M, y: 2.84, w: CW, h: 0.38,
    fontSize: 15, bold: true, color: GREY,
  });
  let y = 3.20;
  q.forEach((t, i) => {
    card(s, M, y, CW, 0.66, i % 2 === 0 ? MIST : MIST2);
    badge(s, M + 0.32, y + 0.13, 0.4, "Q", MOSSD, PAPER, 11);
    txt(s, t, {
      x: M + 0.90, y: y + 0.15, w: CW - 1.3, h: 0.38,
      fontSize: 14.5, color: TEXT,
    });
    y += 0.72;
  });
  card(s, M, 6.14, CW, 0.70, MIST);
  txt(s, "制度が固まる前でも、この問いは今すぐ立てられます。答えが出なくても構いません。", {
    x: M + 0.5, y: 6.31, w: CW - 1.0, h: 0.40,
    fontSize: 15, bold: true, color: PINE,
  });
  s.addNotes("⑩は評価です。評価は素案の中で最も授業に近いところにあるので、研修の題材としては入りやすい。\nこの4つの問いは、制度が固まる前でも立てられます。答えが出なくても構いません。問いを立てること自体に意味があります。");
}

// 45 職員を巻き込む
{
  const s = newSlide();
  head(s, "Ⅳ－10", "職員をどう巻き込むか");
  const st = [
    ["STEP 1", "事実を共有する", "制度の説明ではなく、自校の現状の数字（行事の時数、会議の時間）から始める。", PINE],
    ["STEP 2", "問いを渡す", "「どうすべきか」ではなく「何に時間を使いたいか」を問う。答えを先に用意しない。", MOSSD],
    ["STEP 3", "決めて、引き受ける", "議論は開く。決定は校長がする。反発の矢面には校長が立つ。", GOLD],
  ];
  let y = 1.68;
  st.forEach((x) => {
    card(s, M, y, CW, 1.32, MIST);
    card(s, M, y, 1.9, 1.32, x[3]);
    txt(s, x[0], {
      x: M, y: y + 0.48, w: 1.9, h: 0.36,
      fontSize: 15, bold: true, color: PAPER, align: "center",
    });
    txt(s, x[1], {
      x: M + 2.2, y: y + 0.26, w: CW - 2.6, h: 0.42,
      fontSize: 19, bold: true, color: INK,
    });
    txt(s, x[2], {
      x: M + 2.2, y: y + 0.74, w: CW - 2.6, h: 0.44,
      fontSize: 13.5, color: TEXT,
    });
    y += 1.46;
  });
  card(s, M, 6.06, CW, 0.82, INK);
  txt(s, "議論は開き、決定は引き受ける。これが校長の巻き込み方だと考えます。", {
    x: M + 0.5, y: 6.26, w: CW - 1.0, h: 0.44,
    fontSize: 17, bold: true, color: PAPER,
  });
  s.addNotes("巻き込み方です。制度の説明から入らないこと。自校の数字から入ると、自分たちの話として聞いてもらえます。\nそして最後、議論は開くけれど決定は校長がする。反発の矢面にも校長が立つ。ここを曖昧にすると、かえって職員が不安になります。");
}

// 46 反対と返し方
{
  const s = newSlide();
  head(s, "Ⅳ－11", "よくある反対と、その返し方");
  const qa = [
    ["まだ素案でしょう。確定してから動けばいい。", "確定を待つと設計期間が1年を切ります。今やるのは「決める準備」であって、決定ではありません。"],
    ["行事を減らすと、保護者から苦情が来る。", "減らす前に「何のための行事か」を説明できるようにします。目的とねらいを示せれば、大半は納得されます。"],
    ["削った時間で、結局新しい仕事が増えるのでは。", "だから、何に使うかを先に決めます。使い道を決めずに削ると、必ず別の何かで埋まります。"],
    ["うちの学校には余裕がない。", "余裕がないからこそ着手します。今のまま2028年を迎えるのが、いちばん負担が大きい選択です。"],
  ];
  let y = 1.62;
  qa.forEach((p) => {
    card(s, M, y, CW, 1.2, MIST);
    badge(s, M + 0.32, y + 0.16, 0.44, "Q", GREY, PAPER, 13);
    txt(s, p[0], {
      x: M + 0.92, y: y + 0.17, w: CW - 1.3, h: 0.38,
      fontSize: 15, bold: true, color: GREY,
    });
    badge(s, M + 0.32, y + 0.64, 0.44, "A", PINE, PAPER, 13);
    txt(s, p[1], {
      x: M + 0.92, y: y + 0.63, w: CW - 1.3, h: 0.44,
      fontSize: 13.5, color: TEXT,
    });
    y += 1.32;
  });
  txt(s, "反対意見は、設計を鍛えてくれます。潰すのではなく、拾って使ってください。", {
    x: M, y: 6.92, w: CW, h: 0.38,
    fontSize: 14, bold: true, color: PINE,
  });
  s.addNotes("実際に出てくるであろう反対を4つ挙げました。想定問答としてお使いください。\n3つめが実務的には一番大事です。使い道を決めずに削ると、必ず別の何かで埋まります。削ると決める前に、何に使うかを決めておく。");
}

/* ══════════════════════════════════════════════════
   Ⅴ 令和9年度の学校経営に落とす（47–50）
   ══════════════════════════════════════════════════ */

// 47 経営計画への書き込み例
{
  const s = newSlide();
  head(s, "Ⅴ－1", "学校経営計画に、こう書けます");
  const ba = [
    ["これまでの書き方", GREY, MIST, [
      "「主体的・対話的で深い学びを推進する」",
      "「ICTを効果的に活用する」",
      "「働き方改革を進める」",
    ], "何をするかが書かれていないため、日常業務に流されます。"],
    ["令和9年度の書き方", PINE, MIST2, [
      "「学校行事を棚卸しし、12月までに統合案を決定する」",
      "「情報活用能力の自校系統表を3月までに作成する」",
      "「校内研修を年間◯回に絞り、授業改善に直結する内容とする」",
    ], "主語・期限・成果物が入ると、動き出します。"],
  ];
  ba.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.2);
    const cw15 = CW / 2 - 0.2;
    card(s, x, 1.66, cw15, 4.05, b[2]);
    txt(s, b[0], {
      x: x + 0.38, y: 1.92, w: cw15 - 0.76, h: 0.44,
      fontSize: 19, bold: true, color: b[1],
    });
    let yy = 2.58;
    b[3].forEach((t) => {
      card(s, x + 0.34, yy, cw15 - 0.68, 0.82, PAPER);
      txt(s, t, {
        x: x + 0.54, y: yy + 0.13, w: cw15 - 1.08, h: 0.58,
        fontSize: 12.5, color: i === 0 ? GREY : INK, lineSpacing: 18,
      });
      yy += 0.94;
    });
    txt(s, b[4], {
      x: x + 0.38, y: 5.32, w: cw15 - 0.76, h: 0.34,
      fontSize: 12.5, bold: true, color: i === 0 ? GREY : PINE,
    });
  });
  card(s, M, 5.92, CW, 0.88, INK);
  txt(s, "「主語・期限・成果物」。この3つが入っているかだけ、確認してください。", {
    x: M + 0.5, y: 6.14, w: CW - 1.0, h: 0.46,
    fontSize: 17, bold: true, color: PAPER,
  });
  s.addNotes("経営計画への落とし方です。左のような書き方は、どの学校にもあると思います。私の学校にもありました。\n違いは、主語と期限と成果物が入っているかどうかだけです。この3つが入ると、年度末に振り返れる形になります。");
}

// 48 マイルストーン
{
  const s = newSlide();
  head(s, "Ⅴ－2", "令和9年度 マイルストーン");
  const ms = [
    ["4月", "共有する", ["職員会議で素案の要点を共有", "「何に時間を使いたいか」を問う", "担当を決める"], PINE],
    ["7月", "棚卸す", ["行事・会議・研修を一覧化", "時数と負担を数字にする", "1学期の実態を記録"], MOSSD],
    ["12月", "決める", ["統合・縮減の結論を出す", "令和10年度の時間割方針を決定", "保護者・地域へ説明"], GOLD],
    ["3月", "書く", ["令和10年度教育課程に反映", "経営計画に成果と次の目標", "振り返りを記録に残す"], PINE],
  ];
  const cw16 = (CW - 0.75) / 4;
  ms.forEach((m, i) => {
    const x = M + i * (cw16 + 0.25);
    card(s, x, 1.68, cw16, 4.15, MIST);
    card(s, x, 1.68, cw16, 1.05, m[3]);
    txt(s, m[0], {
      x, y: 1.82, w: cw16, h: 0.42,
      fontSize: 22, bold: true, color: PAPER, align: "center",
    });
    txt(s, m[1], {
      x, y: 2.28, w: cw16, h: 0.32,
      fontSize: 13, color: "D6E4D8", align: "center",
    });
    bullets(s, m[2], {
      x: x + 0.24, y: 2.98, w: cw16 - 0.48, h: 2.6,
      fontSize: 12.5, color: TEXT,
    });
  });
  card(s, M, 6.05, CW, 0.82, MIST2);
  txt(s, "12月が山場です。ここで決めておかないと、令和10年度の教育課程に間に合いません。", {
    x: M + 0.5, y: 6.26, w: CW - 1.0, h: 0.44,
    fontSize: 16, bold: true, color: INK,
  });
  s.addNotes("令和9年度の一年間を4つの節目で整理しました。\n山場は12月です。ここで結論を出しておかないと、令和10年度の教育課程編成に間に合いません。逆算すると、7月までに棚卸しを終えている必要があります。");
}

// 49 3段階とチェックリスト
{
  const s = newSlide();
  head(s, "Ⅴ－3", "無理のない3段階と、確認リスト");
  const lv = [
    ["最小", "まず1つだけ", "行事の棚卸しか、会議の精選。どちらか一方だけで十分です。", GREY],
    ["標準", "3つを組み合わせる", "棚卸し＋研修の絞り込み＋総合の点検。ここまでで令和10年度に接続できます。", MOSSD],
    ["挑戦", "時間割から変える", "単位時間・週時程の組み替えまで踏み込む。区教委との協議が前提です。", PINE],
  ];
  const cw17 = (CW - 0.5) / 3;
  lv.forEach((l, i) => {
    const x = M + i * (cw17 + 0.25);
    card(s, x, 1.66, cw17, 2.15, MIST);
    badge(s, x + 0.26, 1.90, 0.56, l[0], l[3], PAPER, 13);
    txt(s, l[1], {
      x: x + 0.26, y: 2.60, w: cw17 - 0.52, h: 0.42,
      fontSize: 17, bold: true, color: INK,
    });
    txt(s, l[2], {
      x: x + 0.26, y: 3.04, w: cw17 - 0.52, h: 0.68,
      fontSize: 12.5, color: TEXT, lineSpacing: 18,
    });
  });
  card(s, M, 4.06, CW, 2.65, INK);
  txt(s, "年度末に、この4つを確認してください", {
    x: M + 0.5, y: 4.30, w: CW - 1.0, h: 0.42,
    fontSize: 18, bold: true, color: MOSS,
  });
  const chk = [
    "自校の「何に時間を使いたいか」を、言葉にできたか",
    "減らすと決めたものが、1つ以上あるか",
    "生み出した時間の使い道を、先に決めたか",
    "令和10年度の教育課程に、その結果が入っているか",
  ];
  let y = 4.88;
  chk.forEach((c) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.55, y: y + 0.02, w: 0.32, h: 0.32,
      fill: { color: "1B3F35" }, line: { color: MOSS, width: 1.25 }, rectRadius: 0.04,
    });
    txt(s, c, {
      x: M + 1.05, y, w: CW - 1.6, h: 0.36,
      fontSize: 14.5, color: "D6E4D8",
    });
    y += 0.44;
  });
  s.addNotes("いきなり全部は無理ですので、3段階に分けました。最小の「まず1つだけ」で十分です。\n下のチェックリストは、年度末に見ていただくものです。特に3つめ、使い道を先に決めたか。ここが抜けると元に戻ります。");
}

// 50 結び
{
  const s = newSlide(true);
  s.addShape(pres.ShapeType.ellipse, {
    x: 10.1, y: -1.5, w: 4.8, h: 4.8,
    fill: { color: PINE }, line: { type: "none" },
  });
  txt(s, "むすびに", {
    x: M, y: 1.30, w: CW, h: 0.34,
    fontSize: 13, bold: true, color: MOSS, charSpacing: 1.5,
  });
  txt(s, "学校は「与えられるもの」から\n「つくるもの」へ。", {
    x: M, y: 1.78, w: 9.4, h: 1.5,
    fontSize: 34, bold: true, color: PAPER, lineSpacing: 47,
  });
  txt(s, "今回の改訂は、時間の配分という最も基本的な資源を、学校に返す改訂です。\n返されたものをどう使うかは、それぞれの学校が決めます。\n先生方が令和9年度に何を選ばれるか、その積み重ねが、2030年の子供たちの学びになります。", {
    x: M, y: 3.50, w: 10.6, h: 1.35,
    fontSize: 15, color: "C9DACD", lineSpacing: 27,
  });
  card(s, M, 5.15, CW, 1.45, "1B3F35");
  txt(s, "出典", {
    x: M + 0.45, y: 5.32, w: 2.0, h: 0.3,
    fontSize: 11, bold: true, color: MOSS,
  });
  txt(s,
    "中央教育審議会 初等中等教育分科会 教育課程部会 教育課程企画特別部会（第17回）令和8年8月31日 資料1\n" +
    "「次期学習指導要領等に向けた審議まとめ（素案）」（文部科学省）\n" +
    "ならびに同部会の審議経過に関する各種報道・解説記事", {
    x: M + 0.45, y: 5.62, w: CW - 0.9, h: 0.85,
    fontSize: 11.5, color: "AFC5B6", lineSpacing: 17,
  });
  s.addNotes("最後になります。今回の改訂は、時間の配分という最も基本的な資源を学校に返す改訂だと考えています。\n返されたものをどう使うかは、それぞれの学校が決めます。令和9年度に何を選ぶか。その積み重ねが2030年の子供たちの学びになります。\n本日はありがとうございました。");
}


/* ══════════════════════════════════════════════════
   巻末資料 想定質問集（51–56）
   ══════════════════════════════════════════════════ */

// 51 反発への向き合い方
{
  const s = newSlide(true);
  txt(s, "巻末資料", {
    x: M, y: 0.40, w: CW, h: 0.26,
    fontSize: 11, bold: true, color: MOSS, charSpacing: 1.5,
  });
  txt(s, "反発への向き合い方 ― そして想定質問 20", {
    x: M, y: 0.68, w: CW, h: 0.62,
    fontSize: 28, bold: true, color: PAPER,
  });
  card(s, M, 1.62, CW, 1.05, "1B3F35");
  txt(s, "新しいことを始めるとき、反発は必ず出ます。出ないほうが危険です。", {
    x: M + 0.5, y: 1.88, w: CW - 1.0, h: 0.5,
    fontSize: 19, bold: true, color: PAPER,
  });
  const pr = [
    ["反発は「抵抗」ではなく「情報」", "何が不安なのかを教えてくれます。潰すのではなく、聞いて設計に反映させます。反対した人ほど、決まった後は丁寧に動いてくださることが多いものです。"],
    ["「やめる」提案の前に、「減らした」実績を", "先に会議や研修を減らして、時間を返します。校長が本気だと伝わってから、行事の話に入ります。順番を逆にすると、まず通りません。"],
    ["議論は開く。決定と責任は校長が引き受ける", "「みんなで決めましょう」は、一見民主的ですが、反対する人に責任を負わせる形になります。決めるのは校長、と最初に宣言します。"],
  ];
  let y = 2.92;
  pr.forEach((p, i) => {
    card(s, M, y, CW, 1.28, "1B3F35");
    badge(s, M + 0.36, y + 0.34, 0.6, String(i + 1), MOSS, INK, 20);
    txt(s, p[0], {
      x: M + 1.15, y: y + 0.22, w: CW - 1.6, h: 0.42,
      fontSize: 18, bold: true, color: PAPER,
    });
    txt(s, p[1], {
      x: M + 1.15, y: y + 0.68, w: CW - 1.6, h: 0.48,
      fontSize: 12.5, color: "AFC5B6",
    });
    y += 1.42;
  });
  s.addNotes("質疑応答の前に、私の姿勢をお伝えしておきます。反発は必ず出ますし、出ないほうがむしろ危険です。\n特に2つめが実務的に大事です。先に減らした実績を作ってから、行事の話に入る。順番を逆にすると、まず通りません。");
}

// 52–56 Q&A
const QA = [
  ["A", "制度とスケジュールについて", [
    ["まだ素案でしょう。確定してから動けばいいのでは。",
     "今やるのは「決定」ではなく「決める準備」です。確定を待つと設計期間が1年を切ります。準備そのものは無駄になりません。"],
    ["告示されたら、結局やり直しになるのでは。",
     "やり直しになるのは数字の部分だけです。自校の重点を決める作業と行事の棚卸しは、告示の中身に関係なく残ります。"],
    ["指導要領はまた10年で変わる。いずれ元に戻るのでは。",
     "戻る可能性は否定できません。ただ、学校の裁量を広げる方向は20年続いている流れです。逆に狭まった改訂はありません。"],
    ["区教委から何も言われていない。先走らなくてよいのでは。",
     "通知が来てから動くと、全区の学校が同時に動きます。相談も研修も取り合いになります。先に考えておく分には損がありません。"],
  ]],
  ["B", "「今のままでよい」というご意見に", [
    ["今の教育課程で結果が出ている。変える必要があるのか。",
     "結果が出ている学校ほど、何が効いているかを言語化する価値があります。棚卸しは、やめるためではなく、残すものを選ぶ作業です。"],
    ["うちの子供たちには、今のやり方が合っている。",
     "でしたら、それを残す判断をしてください。この制度は減らすことを求めていません。配分を自分で決められるようになるだけです。"],
    ["前任校でうまくいった方法を、変えたくない。",
     "変えろとは申しません。ただ、それが「この学校の子」に合っているかどうかは、一度確かめる価値があると思います。"],
    ["行事は本校の伝統だ。減らすのは学校の文化を壊すことになる。",
     "伝統だからこそ、ねらいを言葉にして残しましょう。説明できない伝統は、次の代で理由もわからないまま消えていきます。"],
  ]],
  ["C", "負担と現実論について", [
    ["ただでさえ忙しい。これ以上仕事を増やさないでほしい。",
     "もっともです。ですから最初の打ち手は「会議の精選」にします。増やす話の前に、減らす話から始めます。"],
    ["削った時間で、結局新しい仕事が増えるのでは。",
     "使い道を先に決めます。決めずに削ると、必ず別の何かで埋まります。ここは校長の責任として明示します。"],
    ["検討する時間が、どこにあるのか。",
     "新しい会議は作りません。既存の職員会議と分掌会議の中で扱います。それで足りない分は、私が引き取ります。"],
    ["人が足りない。制度より人を増やしてほしい。",
     "同感です。人員は区に要望し続けます。同時に、今ある時間の配分は今の人数でも変えられます。両方やります。"],
  ]],
  ["D", "保護者・地域・子供への影響", [
    ["行事を減らすと、保護者から苦情が来る。",
     "減らす前に「何のための行事か」を示します。目的とねらいを説明できれば、大半は納得されます。矢面には私が立ちます。"],
    ["地域行事との関係はどうするのか。",
     "地域には、決定前の段階で早めに相談します。事後報告が一番こじれます。学校運営協議会を活用します。"],
    ["学力が下がったら、誰が責任を取るのか。",
     "私が取ります。そのうえで、下がっていないことを確かめる指標を、着手前に決めておきます。"],
    ["受験に不利になるのではないか。",
     "総時数は減りません。減るのは特定教科の一部で、上乗せ先も学校が選べます。むしろ自校の課題に時間を寄せられます。"],
  ]],
  ["E", "校内の合意形成と公平性", [
    ["自分の教科が削られるのは困る。",
     "どの教科を削るかは、私が一人で決めません。判断基準を先に共有し、全教科に同じものさしを使います。"],
    ["学年によって差が出るのは、不公平ではないか。",
     "学年の実態が違うのに同じにするほうが、不公平になることもあります。差をつける理由を説明できるかどうかで判断します。"],
    ["若手に任せると、質が落ちるのでは。",
     "任せきりにはしません。主任層が伴走します。2030年に中心になる層を、今から育てておく必要があります。"],
    ["校長が代わったら、元に戻るのでは。",
     "だから記録に残します。決めた理由と経緯を文書にしておけば、次の校長は判断のやり直しから始めずに済みます。"],
  ]],
];

let qn = 0;
QA.forEach((grp) => {
  const s = newSlide();
  head(s, "巻末資料　想定質問集", grp[0] + "．" + grp[1]);
  let y = 1.66;
  grp[2].forEach((qa) => {
    qn += 1;
    card(s, M, y, CW, 1.20, qn % 2 === 1 ? MIST : MIST2);
    badge(s, M + 0.30, y + 0.14, 0.44, "Q" + qn, GREY, PAPER, 10);
    txt(s, qa[0], {
      x: M + 0.90, y: y + 0.16, w: CW - 1.3, h: 0.36,
      fontSize: 15, bold: true, color: INK,
    });
    badge(s, M + 0.30, y + 0.64, 0.44, "A", PINE, PAPER, 12);
    txt(s, qa[1], {
      x: M + 0.90, y: y + 0.62, w: CW - 1.3, h: 0.52,
      fontSize: 12.5, color: TEXT, lineSpacing: 18,
    });
    y += 1.26;
  });
  s.addNotes("想定質問集 " + grp[0] + "．" + grp[1] + "（Q" + (qn - 3) + "〜Q" + qn + "）。質疑応答の際の手元資料としてお使いください。答えを読み上げるのではなく、自校の言葉に置き換えてお話しください。");
});

// ── 出力 ───────────────────────────────────────────
const OUT = process.argv[2] || "校長会資料.pptx";
pres.writeFile({ fileName: OUT }).then(() => {
  console.log("slides:", n);
  console.log("written:", OUT);
});
