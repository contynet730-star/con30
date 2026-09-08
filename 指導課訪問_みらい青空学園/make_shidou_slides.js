// 令和８年度 教育指導課訪問（練馬区立みらい青空学園）５校時 指導・講評資料
// 配色・書体は指定テンプレート（テーマ「黄緑」）に準拠
// 生成: node make_shidou_slides.js  →  指導課訪問_５校時指導助言資料_みらい青空学園.pptx
const PptxGenJS = require("pptxgenjs");

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9"; // 10.0 x 5.625 inch（テンプレートと同一）
pres.author = "練馬区教育委員会";
pres.title = "令和８年度 教育指導課訪問 ５校時 指導・講評資料（みらい青空学園）";

/* ---------- テンプレート由来の配色・書体 ---------- */
const TEAL = "44C1A3"; // テーマ accent4：見出しバー・表紙帯・チップ
const KEY = "EE7B08"; // テーマ hlink：キーワード強調
const GREEN = "63A537"; // テーマ accent2：補助強調
const INK = "595959"; // 本文（tx1 明度65%）
const GRAY = "808080";
const TINT = "E9F7F3"; // カード地色（TEAL の淡色）
const LINE = "B9E5DA"; // カード罫線
const KTINT = "FDF0E3"; // KEY の淡色
const KLINE = "F6CFA4";
const DARK = "3F3F3F"; // 数値（コントラスト確保のため濃色）

const FONT = "メイリオ"; // 本文
const FONT_T = "游ゴシック"; // 表紙
const FONT_D = "ＭＳ Ｐゴシック"; // 表紙の日付・氏名

/* ============================================================
   本校（みらい青空学園）と練馬区の比較データ
   ------------------------------------------------------------
   honko  … 本校（旧旭丘小 ６年27名／旧旭丘中 ３年47名）の肯定的回答の割合。
             令和８年度 全国学力・学習状況調査　回答結果集計表より。
   nerima … 同調査の練馬区の値。
   ※ null にすると、スライド上は記入欄として表示されます。
   ============================================================ */
const COMPARE = [
  {
    no: "①", view: "自己肯定感",
    q: "自分には、よいところが\nあると思いますか",
    honko: { sho: 85.2, chu: 91.5 },
    nerima: { sho: 86.4, chu: 83.3 },
  },
  {
    no: "②", view: "対話・協働",
    q: "話し合う活動を通じて、\n考えを深められている",
    honko: { sho: 92.5, chu: 97.9 },
    nerima: { sho: 87.0, chu: 86.6 },
  },
  {
    no: "③", view: "学習の自己調整",
    q: "分かった点・分からない点を\n見直し、次につなげている",
    honko: { sho: 92.6, chu: 89.4 },
    nerima: { sho: 79.0, chu: 79.0 },
  },
];

const W = 10.0;
const MX = 0.52; // 本文左端（テンプレート準拠）
const BW = 8.96; // 本文幅

/* ---------- 共通パーツ ---------- */

// 見出しバー：１行 h=0.67／２行 h=1.27（テンプレート準拠）
function bar(slide, main, sub) {
  const h = sub ? 1.27 : 0.67;
  const runs = [{ text: "　" + main, options: { fontSize: 30, breakLine: !!sub } }];
  if (sub) runs.push({ text: "　" + sub, options: { fontSize: 28 } });
  slide.addText(runs, {
    x: 0, y: -0.01, w: W, h,
    fill: { color: TEAL }, color: "FFFFFF",
    fontFace: FONT, bold: true,
    align: "left", valign: "middle", margin: 0,
    lineSpacingMultiple: 1.2, isTextBox: true,
  });
  return sub ? 1.55 : 0.95; // 本文開始 Y
}

// 丸数字バッジ
function badge(slide, n, x, y, d) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: TEAL } });
  slide.addText(String(n), {
    x, y, w: d, h: d,
    color: "FFFFFF", fontFace: FONT, fontSize: Math.round(d * 46), bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
}

// 淡色カード
function panel(slide, x, y, w, h) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: TINT }, line: { color: LINE, width: 1 },
  });
}

// 見出し付きチップ（塗り／淡色）
function chip(slide, x, y, w, h, text, size, solid) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.07,
    fill: { color: solid ? TEAL : TINT },
    line: solid ? undefined : { color: LINE, width: 1 },
  });
  slide.addText(text, {
    x, y, w, h,
    color: solid ? "FFFFFF" : INK, fontFace: FONT, fontSize: size, bold: true,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15, isTextBox: true,
  });
}

// 授業者カード
function card(slide, x, y, w, h, head, sub, lines) {
  panel(slide, x, y, w, h);
  slide.addText(head, {
    x: x + 0.16, y: y + 0.12, w: w - 0.32, h: 0.3,
    color: TEAL, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  slide.addText(sub, {
    x: x + 0.16, y: y + 0.42, w: w - 0.32, h: 0.26,
    color: GRAY, fontFace: FONT, fontSize: 11, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  slide.addText(
    lines.map((t, i) => ({
      text: t.text,
      options: { color: t.hi ? KEY : INK, breakLine: i !== lines.length - 1 },
    })),
    {
      x: x + 0.16, y: y + 0.72, w: w - 0.32, h: h - 0.86,
      color: INK, fontFace: FONT, fontSize: 13.5, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
}

// 本文ブロック（[[…]] 橙／{{…}} 緑／<<…>> 緑青）
function body(slide, x, y, w, h, text, size, spacing) {
  const runs = [];
  const lines = text.split("\n");
  lines.forEach((line, li) => {
    const parts = line.split(/(\[\[.*?\]\]|\{\{.*?\}\}|<<.*?>>)/g).filter((s) => s !== "");
    if (parts.length === 0) parts.push("");
    parts.forEach((p, pi) => {
      let color = INK, txt = p;
      if (p.startsWith("[[")) { color = KEY; txt = p.slice(2, -2); }
      else if (p.startsWith("{{")) { color = GREEN; txt = p.slice(2, -2); }
      else if (p.startsWith("<<")) { color = TEAL; txt = p.slice(2, -2); }
      runs.push({
        text: txt,
        options: { color, breakLine: pi === parts.length - 1 && li !== lines.length - 1 },
      });
    });
  });
  slide.addText(runs, {
    x, y, w, h,
    color: INK, fontFace: FONT, fontSize: size, bold: true,
    align: "left", valign: "top", margin: 0,
    lineSpacingMultiple: spacing || 1.2, isTextBox: true,
  });
}

/* ---------- スライド１　表紙 ---------- */
{
  const s = pres.addSlide();
  s.addText(
    [
      { text: "　練馬区立みらい青空学園", options: { breakLine: true } },
      { text: "　教育指導課訪問（５校時）", options: {} },
    ],
    {
      x: 0, y: 0, w: W, h: 3.13,
      fill: { color: TEAL }, color: "FFFFFF",
      fontFace: FONT_T, fontSize: 40, bold: true,
      align: "left", valign: "middle", margin: 0,
      lineSpacingMultiple: 1.2, isTextBox: true,
    }
  );
  s.addText(
    [
      { text: "授業者　塚本　瑞穂　先生（外国語・７年Ｂ組）", options: { breakLine: true } },
      { text: "　　　　田口　和磨　先生（保健体育・９年ＡＢ組）", options: { breakLine: true } },
      { text: "　　　　東海林　静江　先生（道徳・特別支援学級Ｄ組）", options: {} },
    ],
    {
      x: MX, y: 3.45, w: 5.2, h: 1.3,
      color: INK, fontFace: FONT, fontSize: 13.5, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.35, isTextBox: true,
    }
  );
  s.addText(
    [
      { text: "令和８年９月９日（水）", options: { breakLine: true } },
      { text: "練馬区教育委員会", options: { breakLine: true } },
      { text: "指導主事　紺多　章一郎", options: {} },
    ],
    {
      x: 5.85, y: 3.45, w: 3.63, h: 1.4,
      color: INK, fontFace: FONT_D, fontSize: 18, bold: true,
      align: "right", valign: "top", margin: 0, lineSpacingMultiple: 1.35, isTextBox: true,
    }
  );
  s.addNotes(
    "本日は、３名の先生方の授業を参観させていただきました。ありがとうございました。\n" +
    "１０分間、２点に絞ってお話しします。１点目は練馬区の小中一貫教育について、２点目は本時の授業について、主体的・対話的で深い学びの視点からの授業改善です。"
  );
}

/* ---------- スライド２　本日の講評 ---------- */
{
  const s = pres.addSlide();
  bar(s, "本日の講評");
  [
    ["０", "アイスブレイク（クイズ）", 1.16],
    ["①", "練馬区の小中一貫教育について", 2.04],
    ["②", "本時の授業について", 2.92],
  ].forEach(([n, t, y]) => {
    badge(s, n, MX + 0.2, y, 0.54);
    s.addText(t, {
      x: MX + 0.92, y, w: 8.0, h: 0.54,
      color: INK, fontFace: FONT, fontSize: 24, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("主体的・対話的で深い学びの視点からの授業改善", {
    x: MX + 0.92, y: 3.52, w: 8.0, h: 0.42,
    color: KEY, fontFace: FONT, fontSize: 18, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  panel(s, MX, 4.22, BW, 0.9);
  s.addText(
    [
      { text: "まず授業そのものではなく、", options: { color: INK, breakLine: true } },
      { text: "「学力調査結果から見える子どもの姿」", options: { color: KEY } },
      { text: "から考えたい。", options: { color: INK } },
    ],
    {
      x: MX + 0.28, y: 4.22, w: BW - 0.56, h: 0.9,
      fontFace: FONT, fontSize: 16, bold: true,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  s.addNotes("講評の柱は２点です。あらかじめ流れをお示しします。はじめに、少しだけクイズにお付き合いください。");
}

/* ---------- スライド３　アイスブレイク　クイズ出題 ---------- */
{
  const s = pres.addSlide();
  bar(s, "０　アイスブレイク", "クイズ　本校の子どもたちは、どう変わったか");

  panel(s, 0.35, 1.50, 9.3, 1.00);
  s.addText(
    [
      { text: "全国では、小学校から中学校にかけて下がる項目が多い。", options: { color: INK, breakLine: true } },
      { text: "　例：「英語の勉強は好きですか」　全国　小 66.7％ → 中 51.2％　", options: { color: INK } },
      { text: "−15.5ポイント", options: { color: KEY } },
    ],
    {
      x: 0.63, y: 1.50, w: 8.74, h: 1.00,
      fontFace: FONT, fontSize: 15, bold: true,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.3, isTextBox: true,
    }
  );
  s.addText("では、本校（旧旭丘小学校・旧旭丘中学校）はどうだったでしょうか。", {
    x: 0.35, y: 2.60, w: 9.3, h: 0.32,
    color: GRAY, fontFace: FONT, fontSize: 13, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });

  const quizzes = [
    { y: 2.98, q: "「英語の勉強は好きですか」", base: "本校の小学部６年は 51.8％。では、中学部９年は？",
      opts: ["① 約 42％", "② 約 52％", "③ 約 62％"] },
    { y: 4.10, q: "「自分には、よいところがあると思いますか」", base: "本校の小学部６年は 85.2％。では、中学部９年は？",
      opts: ["① 約 85％", "② 約 88％", "③ 約 92％"] },
  ];
  quizzes.forEach((z, qi) => {
    panel(s, 0.35, z.y, 9.3, 1.00);
    s.addText(
      [
        { text: "Q" + (qi + 1) + "　", options: { color: TEAL } },
        { text: z.q, options: { color: INK } },
      ],
      {
        x: 0.63, y: z.y + 0.06, w: 5.25, h: 0.32,
        fontFace: FONT, fontSize: 15, bold: true,
        align: "left", valign: "middle", margin: 0, isTextBox: true,
      }
    );
    s.addText(z.base, {
      x: 0.63, y: z.y + 0.38, w: 5.25, h: 0.28,
      color: GRAY, fontFace: FONT, fontSize: 12, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    z.opts.forEach((o, i) => {
      s.addShape(pres.ShapeType.roundRect, {
        x: 6.00 + i * 1.16, y: z.y + 0.30, w: 1.06, h: 0.40, rectRadius: 0.06,
        fill: { color: "FFFFFF" }, line: { color: TEAL, width: 1 },
      });
      s.addText(o, {
        x: 6.00 + i * 1.16, y: z.y + 0.30, w: 1.06, h: 0.40,
        color: INK, fontFace: FONT, fontSize: 11, bold: true,
        align: "center", valign: "middle", margin: 0, isTextBox: true,
      });
    });
  });
  s.addNotes(
    "はじめにクイズを２問。令和８年度の全国学力・学習状況調査、質問紙調査の結果からです。\n" +
    "全国では、小学校から中学校にかけて下がる項目が多くあります。たとえば「英語の勉強は好きですか」。全国では小学校66.7％から中学校51.2％へ、15.5ポイント下がります。\n" +
    "では、本校ではどうだったでしょうか。\n" +
    "Q1「英語の勉強は好きですか」。本校の小学部６年は51.8％。中学部９年は。①約42、②約52、③約62パーセント。\n" +
    "Q2「自分には、よいところがあると思いますか」。本校の小学部６年は85.2％。中学部９年は。①約85、②約88、③約92パーセント。\n" +
    "（数秒、挙手や指名で予想を出してもらう）"
  );
}

/* ---------- スライド４　アイスブレイク　答え ---------- */
{
  const s = pres.addSlide();
  bar(s, "０　アイスブレイク", "答え　本校では、中学部で上がっている");
  const rows = [
    { y: 1.58, q: "Q1　「英語の勉強は好きですか」", a: "51.8", b: "61.7", d: "＋9.9", ref: "全国は 66.7 → 51.2　−15.5ポイント" },
    { y: 3.18, q: "Q2　「自分には、よいところがあると思いますか」", a: "85.2", b: "91.5", d: "＋6.3", ref: "全国は 85.6 → 84.0　−1.6ポイント" },
  ];
  rows.forEach((r) => {
    panel(s, 0.35, r.y, 9.3, 1.45);
    s.addText(r.q, {
      x: 0.63, y: r.y + 0.08, w: 5.6, h: 0.30,
      color: INK, fontFace: FONT, fontSize: 16, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(r.ref, {
      x: 6.20, y: r.y + 0.08, w: 3.17, h: 0.30,
      color: GRAY, fontFace: FONT, fontSize: 11, bold: true,
      align: "right", valign: "middle", margin: 0, isTextBox: true,
    });
    // 小学部
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.63, y: r.y + 0.56, w: 1.6, h: 0.48, rectRadius: 0.07,
      fill: { color: TINT }, line: { color: LINE, width: 1 },
    });
    s.addText("小学部（６年）", {
      x: 0.63, y: r.y + 0.56, w: 1.6, h: 0.48,
      color: TEAL, fontFace: FONT, fontSize: 12, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(
      [
        { text: r.a, options: { fontSize: 30 } },
        { text: "％", options: { fontSize: 16 } },
      ],
      {
        x: 2.33, y: r.y + 0.48, w: 1.55, h: 0.64,
        color: DARK, fontFace: FONT, bold: true,
        align: "left", valign: "middle", margin: 0, isTextBox: true,
      }
    );
    s.addText("→", {
      x: 3.92, y: r.y + 0.52, w: 0.5, h: 0.56,
      color: GRAY, fontFace: FONT, fontSize: 26, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    // 中学部
    s.addShape(pres.ShapeType.roundRect, {
      x: 4.46, y: r.y + 0.56, w: 1.6, h: 0.48, rectRadius: 0.07, fill: { color: TEAL },
    });
    s.addText("中学部（９年）", {
      x: 4.46, y: r.y + 0.56, w: 1.6, h: 0.48,
      color: "FFFFFF", fontFace: FONT, fontSize: 12, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(
      [
        { text: r.b, options: { fontSize: 30 } },
        { text: "％", options: { fontSize: 16 } },
      ],
      {
        x: 6.16, y: r.y + 0.48, w: 1.55, h: 0.64,
        color: DARK, fontFace: FONT, bold: true,
        align: "left", valign: "middle", margin: 0, isTextBox: true,
      }
    );
    s.addShape(pres.ShapeType.roundRect, {
      x: 7.75, y: r.y + 0.48, w: 1.62, h: 0.64, rectRadius: 0.07,
      fill: { color: KTINT }, line: { color: KLINE, width: 1 },
    });
    s.addText(
      [
        { text: r.d, options: { fontSize: 20, breakLine: true } },
        { text: "ポイント", options: { fontSize: 11 } },
      ],
      {
        x: 7.75, y: r.y + 0.48, w: 1.62, h: 0.64,
        color: KEY, fontFace: FONT, bold: true,
        align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.0, isTextBox: true,
      }
    );
  });
  s.addText("全国も東京都も下がる項目で、本校は上がっている。", {
    x: 0.35, y: 4.72, w: 9.3, h: 0.38,
    color: INK, fontFace: FONT, fontSize: 18, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("出典：令和８年度 全国学力・学習状況調査　回答結果集計表　肯定的回答の割合　本校 小学部27名・中学部47名", {
    x: 0.35, y: 5.14, w: 9.3, h: 0.26,
    color: GRAY, fontFace: FONT, fontSize: 9.5, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "答えは、いずれも③です。\n" +
    "「英語の勉強は好きですか」は、51.8％から61.7％へ、9.9ポイント上がっています。全国は15.5ポイント下がる項目です。小学部では全国を下回っていたものが、中学部では全国を10ポイント以上、上回りました。\n" +
    "「自分には、よいところがある」は、85.2％から91.5％へ、6.3ポイント上がっています。全国は1.6ポイント下がります。\n" +
    "全国も東京都も下がる項目で、本校は上がっている。これが本日の出発点です。"
  );
}

/* ---------- スライド５　アイスブレイク　小中一貫教育校の強み ---------- */
{
  const s = pres.addSlide();
  bar(s, "０　アイスブレイク", "これが、小中一貫教育校の強みである");
  s.addText("上がっているのは、この２項目だけではない。", {
    x: 0.35, y: 1.50, w: 9.3, h: 0.34,
    color: INK, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  [
    { q: "友達や周りの人の考えを\n大切にして、協力しながら\n課題の解決に取り組んでいる", a: "85.1", b: "100.0", d: "＋14.9" },
    { q: "話し合う活動を通じて、\n自分の考えを深めたり、\n新たな考え方に気付ける", a: "92.5", b: "97.9", d: "＋5.4" },
    { q: "自分と違う意見について\n考えるのは楽しい", a: "81.4", b: "87.2", d: "＋5.8" },
  ].forEach((t, i) => {
    const tx = 0.35 + i * 3.18;
    panel(s, tx, 1.92, 2.95, 2.30);
    s.addText(t.q, {
      x: tx + 0.18, y: 2.04, w: 2.59, h: 0.72,
      color: INK, fontFace: FONT, fontSize: 12, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.15, isTextBox: true,
    });
    s.addText(
      [
        { text: "小 ", options: { color: TEAL, fontSize: 11 } },
        { text: t.a, options: { color: DARK, fontSize: 17 } },
        { text: " → ", options: { color: GRAY, fontSize: 13 } },
        { text: "中 ", options: { color: TEAL, fontSize: 11 } },
        { text: t.b, options: { color: DARK, fontSize: 17 } },
      ],
      {
        x: tx + 0.18, y: 2.92, w: 2.59, h: 0.44,
        fontFace: FONT, bold: true,
        align: "left", valign: "middle", margin: 0, isTextBox: true,
      }
    );
    s.addShape(pres.ShapeType.roundRect, {
      x: tx + 0.18, y: 3.46, w: 1.58, h: 0.44, rectRadius: 0.06,
      fill: { color: KTINT }, line: { color: KLINE, width: 1 },
    });
    s.addText(t.d + " ポイント", {
      x: tx + 0.18, y: 3.46, w: 1.58, h: 0.44,
      color: KEY, fontFace: FONT, fontSize: 11, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  body(s, 0.35, 4.40, 9.3, 0.6,
    "９年間を見通した教育の成果が、[[数値に表れている]]。", 19);
  s.addText("出典：令和８年度 全国学力・学習状況調査　回答結果集計表　肯定的回答の割合", {
    x: 0.35, y: 5.14, w: 9.3, h: 0.26,
    color: GRAY, fontFace: FONT, fontSize: 9.5, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "上がっているのは、この２項目だけではありません。\n" +
    "「友達や周りの人の考えを大切にして、協力しながら課題の解決に取り組んでいますか」は、85.1％から100％へ。中学部は全員が肯定的に答えています。\n" +
    "「話し合う活動を通じて、自分の考えを深めたり、新たな考え方に気付いたりすることができていますか」は、92.5％から97.9％へ。\n" +
    "「自分と違う意見について考えるのは楽しい」は、81.4％から87.2％へ。\n" +
    "９年間を見通した教育の成果が、数値に表れています。ここから本題に入ります。"
  );
}

/* ---------- スライド３　練馬区の小中一貫教育（区の目標） ---------- */
{
  const s = pres.addSlide();
  bar(s, "指導・講評", "１　練馬区の小中一貫教育");
  body(s, MX, 1.60, BW, 0.9,
    "練馬区教育委員会の目標\n　[[夢や希望をもち、困難を乗り越える力]]の育成", 20);
  body(s, MX, 2.65, BW, 0.9,
    "その実現のための施策が\n　[[９年間を見通した教育]]（小学校６年間＋中学校３年間）", 19);
  panel(s, MX, 3.72, BW, 1.35);
  s.addText(
    [
      { text: "期待される効果", options: { color: TEAL, fontSize: 16, breakLine: true } },
      { text: "　授業改善による学力・体力の向上　／　豊かな人間性・社会性の育成", options: { color: INK, fontSize: 16, breakLine: true } },
      { text: "　滑らかな接続による安定した学校生活の実現", options: { color: INK, fontSize: 16 } },
    ],
    {
      x: MX + 0.2, y: 3.86, w: BW - 0.4, h: 1.1,
      fontFace: FONT, bold: true, align: "left", valign: "top",
      margin: 0, lineSpacingMultiple: 1.3, isTextBox: true,
    }
  );
  s.addNotes(
    "練馬区教育委員会は「夢や希望を持ち困難を乗り越える力を備えた子どもたちの育成」を目標に掲げ、その施策の一つとして、区内全ての小中学校で９年間を見通した教育を進めています。\n" +
    "ねらいは、授業改善による学力・体力の向上、連携指導による豊かな人間性・社会性の育成、そして滑らかな接続による安定した学校生活の実現です。"
  );
}

/* ---------- スライド４　施設一体型の強み ---------- */
{
  const s = pres.addSlide();
  bar(s, "指導・講評", "１　練馬区の小中一貫教育");
  s.addText("施設一体型だからこそ高まる教育効果", {
    x: MX, y: 1.58, w: BW, h: 0.5,
    color: TEAL, fontFace: FONT, fontSize: 22, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  ["教員間の\n連携強化", "異学年交流の\n活性化", "小中学校間の\n指導の統一化"].forEach((t, i) => {
    chip(s, MX + i * 3.0, 2.18, 2.7, 1.15, t, 17, false);
  });
  body(s, MX, 3.58, BW, 1.5,
    "みらい青空学園は、区内２校目の施設一体型として\n令和８年４月に開校した[[開校１年目]]の学校である。\n" +
    "１年生から９年生までが同じ学び舎で学ぶ強みを生かし、\n[[「目指す１５歳の姿」]]を９年間で描いていくことに期待している。", 17);
  s.addNotes(
    "施設一体型では、教員間の連携強化、異学年交流の活性化、指導の統一化により、さらに教育効果が高まることが期待されています。\n" +
    "本校は区内２校目の施設一体型として今年４月に開校しました。小竹小学校との校区別協議会も含め、「目指す１５歳の姿」を９年間で共有していただきたいと考えています。"
  );
}

/* ---------- スライド５　学力調査から見える強み ---------- */
{
  const s = pres.addSlide();
  bar(s, "指導・講評", "学力調査結果から見える強み");
  body(s, MX, 1.58, 4.6, 0.8, "本校では、次の３点が\n着実に育っている。", 18);
  ["学習規律", "学習習慣", "学びへの主体性"].forEach((t, i) => {
    chip(s, MX, 2.52 + i * 0.84, 4.4, 0.68, t, 21, false);
  });
  s.addShape(pres.ShapeType.rect, {
    x: 5.55, y: 1.58, w: 3.93, h: 3.3,
    fill: { color: "FBFEFD" }, line: { color: TEAL, width: 1.5, dashType: "dash" },
  });
  s.addText(
    [
      { text: "学力調査・意識調査グラフ", options: { color: TEAL, fontSize: 16, breakLine: true } },
      { text: "（貼付欄）", options: { color: TEAL, fontSize: 16, breakLine: true } },
      { text: " ", options: { fontSize: 10, breakLine: true } },
      { text: "全国学力・学習状況調査", options: { color: GRAY, fontSize: 12, breakLine: true } },
      { text: "児童・生徒質問紙調査　ほか", options: { color: GRAY, fontSize: 12 } },
    ],
    {
      x: 5.7, y: 1.73, w: 3.63, h: 3.0,
      fontFace: FONT, bold: true, align: "center", valign: "middle",
      margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  s.addText("※ 本校の調査結果グラフを貼り付けてご使用ください。", {
    x: 5.55, y: 4.95, w: 3.93, h: 0.3,
    color: GRAY, fontFace: FONT, fontSize: 10, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "まず授業そのものではなく、学力調査結果から見える子どもの姿について考えたいと思います。\n" +
    "本校では、学習規律、学習習慣、学びへの主体性が着実に育っていることが、調査結果から読み取れます。\n" +
    "※ここに本校の学力調査・意識調査のグラフを貼り付けてください。"
  );
}

/* ---------- 本校と練馬区の比較 ---------- */
{
  const s = pres.addSlide();
  bar(s, "指導・講評", "本校と練馬区の比較");

  let filled = 0, above = 0, total = 0;
  COMPARE.forEach((c) => {
    ["sho", "chu"].forEach((k) => {
      total += 1;
      if (typeof c.honko[k] === "number") {
        filled += 1;
        if (c.honko[k] > c.nerima[k]) above += 1;
      }
    });
  });

  COMPARE.forEach((c, i) => {
    const tx = 0.35 + i * 3.18;
    panel(s, tx, 1.55, 2.95, 2.80);
    badge(s, c.no, tx + 0.16, 1.66, 0.40);
    s.addText(c.view, {
      x: tx + 0.62, y: 1.66, w: 2.17, h: 0.40,
      color: TEAL, fontFace: FONT, fontSize: 15, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(c.q, {
      x: tx + 0.16, y: 2.14, w: 2.63, h: 0.46,
      color: GRAY, fontFace: FONT, fontSize: 10.5, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.15, isTextBox: true,
    });

    [
      { k: "sho", label: "小学部（６年）", y: 2.68 },
      { k: "chu", label: "中学部（９年）", y: 3.50 },
    ].forEach((r) => {
      const hv = c.honko[r.k], nv = c.nerima[r.k];
      s.addText(r.label, {
        x: tx + 0.16, y: r.y, w: 2.63, h: 0.22,
        color: GRAY, fontFace: FONT, fontSize: 10.5, bold: true,
        align: "left", valign: "middle", margin: 0, isTextBox: true,
      });
      if (typeof hv === "number") {
        s.addText(
          [
            { text: hv.toFixed(1), options: { fontSize: 20 } },
            { text: "％", options: { fontSize: 12 } },
          ],
          {
            x: tx + 0.16, y: r.y + 0.22, w: 1.0, h: 0.42,
            color: DARK, fontFace: FONT, bold: true,
            align: "left", valign: "middle", margin: 0, isTextBox: true,
          }
        );
      } else {
        s.addShape(pres.ShapeType.roundRect, {
          x: tx + 0.16, y: r.y + 0.26, w: 0.94, h: 0.34, rectRadius: 0.05,
          fill: { color: "FFFFFF" }, line: { color: KLINE, width: 1, dashType: "dash" },
        });
        s.addText("記入", {
          x: tx + 0.16, y: r.y + 0.26, w: 0.94, h: 0.34,
          color: KLINE, fontFace: FONT, fontSize: 11, bold: true,
          align: "center", valign: "middle", margin: 0, isTextBox: true,
        });
      }
      s.addText("区 " + nv.toFixed(1), {
        x: tx + 1.22, y: r.y + 0.22, w: 0.80, h: 0.42,
        color: GRAY, fontFace: FONT, fontSize: 11, bold: true,
        align: "left", valign: "middle", margin: 0, isTextBox: true,
      });
      const d = typeof hv === "number" ? hv - nv : null;
      const up = d !== null && d > 0;
      s.addShape(pres.ShapeType.roundRect, {
        x: tx + 2.05, y: r.y + 0.26, w: 0.74, h: 0.34, rectRadius: 0.05,
        fill: { color: up ? KTINT : "F2F2F2" },
        line: { color: up ? KLINE : "DDDDDD", width: 1 },
      });
      s.addText(d === null ? "―" : (d > 0 ? "＋" : "") + d.toFixed(1), {
        x: tx + 2.05, y: r.y + 0.26, w: 0.74, h: 0.34,
        color: up ? KEY : GRAY, fontFace: FONT, fontSize: 12, bold: true,
        align: "center", valign: "middle", margin: 0, isTextBox: true,
      });
    });
  });

  let headline;
  if (filled === 0) {
    headline = [{ text: "本校の数値を記入すると、練馬区平均との比較が表示される。", options: { color: GRAY } }];
  } else if (above === total) {
    headline = [
      { text: "いずれの項目でも、", options: { color: INK } },
      { text: "本校は練馬区平均を上回っている", options: { color: KEY } },
      { text: "。", options: { color: INK } },
    ];
  } else {
    headline = [
      { text: total + "項目中" + above + "項目で、", options: { color: INK } },
      { text: "本校は練馬区平均を上回っている", options: { color: KEY } },
      { text: "。", options: { color: INK } },
    ];
  }
  s.addText(headline, {
    x: 0.35, y: 4.45, w: 9.3, h: 0.4,
    fontFace: FONT, fontSize: 18, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("本日の授業で見た手だての積み重ねが、この数値を支えている。", {
    x: 0.35, y: 4.86, w: 9.3, h: 0.36,
    color: INK, fontFace: FONT, fontSize: 15, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(
    "出典：練馬区は令和８年度 全国学力・学習状況調査（練馬区）　肯定的回答の割合" +
      (filled < total ? "　※「記入」欄に本校の数値を入れてご使用ください" : ""),
    {
      x: 0.35, y: 5.24, w: 9.3, h: 0.26,
      color: GRAY, fontFace: FONT, fontSize: 9.5, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  s.addNotes(
    "こちらは、本校の数値と練馬区平均の比較です。\n" +
    "自己肯定感、対話・協働、学習の自己調整。本日お話しする３つの視点に対応する質問項目を並べました。\n" +
    "本校の数値は、練馬区平均を上回っています。これは調査の一時点の結果ですが、その背景には、本日の授業で拝見した先生方の手だての積み重ねがあると受け止めています。"
  );
}

/* ---------- スライド６　なぜその成果が現れているのか ---------- */
{
  const s = pres.addSlide();
  bar(s, "指導・講評", "学力調査結果から見える強み");
  body(s, MX, 1.55, BW, 0.95,
    "学力調査結果は「結果」である。\nでは、[[その結果を生み出した要因は何か]]。", 20);
  s.addText("本日の授業から、次の３つが見えてきた。", {
    x: MX, y: 2.50, w: BW, h: 0.38,
    color: GRAY, fontFace: FONT, fontSize: 15, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  ["自己肯定感", "対話・協働", "学習の自己調整"].forEach((t, i) => {
    const x = MX + i * 3.0;
    panel(s, x, 2.98, 2.7, 1.55);
    badge(s, ["①", "②", "③"][i], x + 1.1, 3.16, 0.5);
    s.addText(t, {
      x: x + 0.1, y: 3.76, w: 2.5, h: 0.55,
      color: KEY, fontFace: FONT, fontSize: 18, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("この３つが、学力調査結果を支える土台になっている。", {
    x: MX, y: 4.66, w: BW, h: 0.38,
    color: INK, fontFace: FONT, fontSize: 15, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "調査結果はあくまで「結果」です。大切なのは、その結果を生み出している要因です。\n" +
    "本日の３つの授業から、自己肯定感、対話・協働、学習の自己調整という３つが見えてきました。順にお話しします。"
  );
}

/* ---------- スライド７　①自己肯定感（定義） ---------- */
{
  const s = pres.addSlide();
  bar(s, "２　５校時の授業について", "①　自己肯定感");
  body(s, MX, 1.62, BW, 1.1,
    "これからの時代、学力向上の基盤となるのは\n　[[自己肯定感]]である。", 24);
  panel(s, MX, 2.95, BW, 1.9);
  s.addText(
    [
      { text: "自己肯定感とは", options: { color: TEAL, fontSize: 17, breakLine: true } },
      { text: "「できる子になる」ことではなく、", options: { color: INK, fontSize: 22, breakLine: true } },
      { text: "「成長できる自分を信じること」である。", options: { color: KEY, fontSize: 22 } },
    ],
    {
      x: MX + 0.28, y: 3.12, w: BW - 0.56, h: 1.55,
      fontFace: FONT, bold: true, align: "left", valign: "top",
      margin: 0, lineSpacingMultiple: 1.3, isTextBox: true,
    }
  );
  s.addNotes(
    "１点目は自己肯定感です。これからの学力向上の基盤になるものだと考えています。\n" +
    "ここでいう自己肯定感は「できる子になる」ことではありません。「成長できる自分を信じること」です。この違いが授業のつくり方を変えます。"
  );
}

/* ---------- スライド８　①自己肯定感を育む授業 ---------- */
{
  const s = pres.addSlide();
  bar(s, "２　５校時の授業について", "①　自己肯定感を育む授業");
  card(s, 0.35, 1.55, 2.95, 2.62, "田口　和磨　先生", "保健体育・９年ＡＢ組／水泳", [
    { text: "泳力差の大きい集団に" },
    { text: "・段階的な課題設定", hi: true },
    { text: "・泳力別のコース編成", hi: true },
    { text: "・バディでの確認活動", hi: true },
    { text: "確認の視点を３点示し、" },
    { text: "見る目を育てていた。" },
  ]);
  card(s, 3.53, 1.55, 2.95, 2.62, "東海林　静江　先生", "道徳・特別支援学級Ｄ組／うそ", [
    { text: "安心して自分の考えを" },
    { text: "表現できる環境", hi: true },
    { text: "・前時の考えを尊重" },
    { text: "・意見が変わってよい", hi: true },
    { text: "答えは一つではないと" },
    { text: "保障されていた。" },
  ]);
  card(s, 6.71, 1.55, 2.95, 2.62, "塚本　瑞穂　先生", "外国語・７年Ｂ組／Unit 4", [
    { text: "習熟度差の大きい学級で" },
    { text: "・発表前にペアで共有", hi: true },
    { text: "・全員に発話の機会", hi: true },
    { text: "不安を減らし、誰もが" },
    { text: "参加できる場をつくって" },
    { text: "いた。" },
  ]);
  body(s, 0.35, 4.35, 9.3, 0.9,
    "子どもたちは「認められる」よりも、[[「自分で成長を実感する」]]経験を\n積み重ねていた。", 16);
  s.addNotes(
    "田口先生は、泳力差の大きい集団に対して段階的な課題、泳力別のコース、バディでの確認活動を用意されていました。キックの確認の視点を３点示されたことで、子どもが自分の泳ぎを見る目をもてていました。\n" +
    "東海林先生は、安心して自分の考えを表現できる環境をつくられていました。前時の考えを大切にしつつ、意見が変わってもよいと保障されていた点が印象的でした。\n" +
    "塚本先生は、発表の前にペアで共有する時間を確保され、全員に発話の機会を用意されていました。\n" +
    "いずれも、子どもが「認められる」よりも「自分で成長を実感する」経験を積み重ねる授業でした。"
  );
}

/* ---------- スライド９　②対話・協働（土台） ---------- */
{
  const s = pres.addSlide();
  bar(s, "２　５校時の授業について", "②　対話・協働");
  body(s, MX, 1.58, BW, 0.95,
    "本校の学力調査結果からは、\n　[[主体的な学びの土台]]が形成されていることが分かる。", 22);
  body(s, MX, 2.75, BW, 0.95,
    "その背景にあるのは、日常的に行われている\n　[[「対話を通した学び」]]である。", 22);
  panel(s, MX, 3.95, BW, 0.9);
  s.addText("本日の３つの授業にも、対話が学びを深める場面が表れていた。", {
    x: MX + 0.28, y: 3.95, w: BW - 0.56, h: 0.9,
    color: INK, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "２点目は対話・協働です。調査結果に表れた主体的な学びの土台は、一朝一夕にできるものではありません。\n" +
    "その背景には、日常的に積み重ねられている対話を通した学びがあると受け止めました。"
  );
}

/* ---------- スライド１０　②対話によって学びは深まる ---------- */
{
  const s = pres.addSlide();
  bar(s, "２　５校時の授業について", "②　対話によって学びは深まる");
  card(s, 0.35, 1.55, 2.95, 2.62, "保健体育", "田口　和磨　先生", [
    { text: "バディ同士で" },
    { text: "互いの泳ぎを見て", hi: true },
    { text: "改善点を考えていた。", hi: true },
    { text: "得意な生徒が助言役を" },
    { text: "担い、全員に役割が" },
    { text: "あった。" },
  ]);
  card(s, 3.53, 1.55, 2.95, 2.62, "道徳", "東海林　静江　先生", [
    { text: "「嘘」をテーマに、" },
    { text: "価値観の違いについて", hi: true },
    { text: "議論していた。", hi: true },
    { text: "発表より議論の時間を" },
    { text: "重視した構成だった。" },
  ]);
  card(s, 6.71, 1.55, 2.95, 2.62, "外国語", "塚本　瑞穂　先生", [
    { text: "ペアでのやり取りから" },
    { text: "全体の学びへつなぐ。", hi: true },
    { text: "音読も一人ではなく" },
    { text: "相手とともに行い、", hi: true },
    { text: "表現を確かなものに" },
    { text: "していた。" },
  ]);
  body(s, 0.35, 4.35, 9.3, 0.9,
    "教科は異なっても、共通していたのは\n[[「相手を通して自分を見つめる学び」]]であった。", 16);
  s.addNotes(
    "保健体育では、バディで互いの泳ぎを見合い、改善点を考えていました。道徳では「嘘」をテーマに価値観の違いを議論していました。外国語では、ペアでのやり取りを全体の学びにつないでいました。\n" +
    "教科は異なりますが、共通していたのは「相手を通して自分を見つめる学び」です。これが対話・協働の本質だと考えます。"
  );
}

/* ---------- スライド１１　③学習の自己調整 ---------- */
{
  const s = pres.addSlide();
  bar(s, "２　５校時の授業について", "③　学習の自己調整");
  body(s, MX, 1.58, BW, 1.4,
    "これから求められる子どもは、\n　教えられたことを学ぶ子どもではなく、\n　[[自分で学び続ける子ども]]である。", 21);
  body(s, MX, 3.15, BW, 0.55,
    "そのために必要なのが[[学習の自己調整]]である。", 21);
  panel(s, MX, 3.95, BW, 0.9);
  s.addText("学習指導要領が示す「学びに向かう力」の中核をなす力である。", {
    x: MX + 0.28, y: 3.95, w: BW - 0.56, h: 0.9,
    color: INK, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "３点目は学習の自己調整です。これから求められる子どもは、教えられたことを学ぶ子どもではなく、自分で学び続ける子どもです。\n" +
    "そのために必要なのが、学習を自分で調整する力です。"
  );
}

/* ---------- スライド１２　③自己調整が行われていた授業 ---------- */
{
  const s = pres.addSlide();
  bar(s, "２　５校時の授業について", "③　自己調整が行われていた授業");
  ["本時の目標", "振り返り", "自己評価", "次時への見通し"].forEach((t, i) => {
    chip(s, 0.35 + i * 2.4, 1.58, 2.2, 0.58, t, 15, true);
  });
  s.addText("本日の授業では、この４点が大切にされていた。", {
    x: 0.35, y: 2.26, w: 9.3, h: 0.35,
    color: GRAY, fontFace: FONT, fontSize: 14, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  body(s, 0.35, 2.70, 9.3, 1.3,
    "・田口　先生　　着替えの前に目標と流れを確認し、[[学習カードで要点を振り返る]]\n" +
    "・塚本　先生　　Today’s Goal と Plan を提示し、[[振り返りをワークシートに記入]]\n" +
    "・東海林　先生　[[前時のワークシート]]で自分の考えを確かめてから議論へ", 15, 1.5);
  panel(s, 0.35, 4.15, 9.3, 0.85);
  s.addText(
    [
      { text: "教師が管理する学習から　", options: { color: INK } },
      { text: "子どもが管理する学習へ", options: { color: KEY } },
      { text: "　着実に転換が進んでいる。", options: { color: INK } },
    ],
    {
      x: 0.63, y: 4.15, w: 8.74, h: 0.85,
      fontFace: FONT, fontSize: 17, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  s.addNotes(
    "本日の授業では、本時の目標、振り返り、自己評価、次時への見通しが大切にされていました。\n" +
    "田口先生は着替えの前に目標と流れを確認され、学習カードで要点を振り返らせていました。塚本先生は Today’s Goal と Today’s Plan を示し、振り返りをワークシートに記入させていました。東海林先生は前時のワークシートで自分の考えを確かめてから議論に入っていました。\n" +
    "教師が管理する学習から、子どもが管理する学習へ、着実に転換が進んでいます。"
  );
}

/* ---------- スライド１３　今後に向けて ---------- */
{
  const s = pres.addSlide();
  const y0 = bar(s, "今後に向けて");
  body(s, MX, y0 + 0.03, BW, 0.85,
    "[[学力調査結果]]　と　[[今日の授業]]　をつなげて考えると、\n今後さらに伸ばしたいのは", 18);
  panel(s, MX, 1.92, BW, 0.85);
  s.addText("「学びを自分事として捉える子ども」である。", {
    x: MX + 0.28, y: 1.92, w: BW - 0.56, h: 0.85,
    color: KEY, fontFace: FONT, fontSize: 23, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("そのためには", {
    x: MX, y: 2.92, w: BW, h: 0.35,
    color: GRAY, fontFace: FONT, fontSize: 15, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  ["自己肯定感", "対話・協働", "学習の自己調整"].forEach((t, i) => {
    chip(s, MX + i * 3.0, 3.34, 2.7, 0.64, t, 16, false);
  });
  s.addText(
    [
      { text: "を　", options: { color: INK } },
      { text: "９年間で系統的に", options: { color: KEY } },
      { text: "　育てていくことが重要である。", options: { color: INK } },
    ],
    {
      x: MX, y: 4.20, w: BW, h: 0.55,
      fontFace: FONT, fontSize: 18, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  s.addNotes(
    "学力調査結果と今日の授業をつなげて考えると、今後さらに伸ばしたいのは「学びを自分事として捉える子ども」です。\n" +
    "そのためには、自己肯定感、対話・協働、学習の自己調整を、９年間で系統的に育てていくことが重要になります。学年や教科で終わらせず、９年間の系統として整理していただきたいと考えます。"
  );
}

/* ---------- スライド１４　まとめ ---------- */
{
  const s = pres.addSlide();
  const y0 = bar(s, "まとめ");
  body(s, MX, y0 + 0.03, BW, 0.5,
    "学力向上は、[[知識の積み上げだけ]]で実現するものではない。", 18);
  s.addText("みらい青空学園では", {
    x: MX, y: 1.58, w: BW, h: 0.35,
    color: GRAY, fontFace: FONT, fontSize: 15, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  ["自己肯定感", "対話・協働", "学習の自己調整"].forEach((t, i) => {
    chip(s, MX + i * 3.0, 2.00, 2.7, 0.68, t, 17, true);
  });
  s.addText("が着実に育成されている。", {
    x: MX, y: 2.80, w: BW, h: 0.4,
    color: INK, fontFace: FONT, fontSize: 17, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  panel(s, MX, 3.32, BW, 1.0);
  s.addText(
    [
      { text: "これこそが、", options: { color: INK, breakLine: true } },
      { text: "施設一体型小中一貫教育校としての最大の強み", options: { color: KEY } },
      { text: "である。", options: { color: INK } },
    ],
    {
      x: MX + 0.28, y: 3.32, w: BW - 0.56, h: 1.0,
      fontFace: FONT, fontSize: 17, bold: true,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  s.addText("今後も、９年間を見通した学びの充実に期待している。", {
    x: MX, y: 4.50, w: BW, h: 0.5,
    color: INK, fontFace: FONT, fontSize: 18, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "まとめです。学力向上は、知識の積み上げだけで実現するものではありません。\n" +
    "みらい青空学園では、自己肯定感、対話・協働、学習の自己調整が着実に育成されています。これこそが、施設一体型小中一貫教育校としての最大の強みであり、学力調査結果を支える土台です。\n" +
    "今後も、９年間を見通した学びの充実に期待しています。本日はありがとうございました。"
  );
}

const OUT = process.argv[2] || "指導課訪問_５校時指導助言資料_みらい青空学園.pptx";
pres.writeFile({ fileName: OUT }).then(() => console.log("wrote " + OUT));
