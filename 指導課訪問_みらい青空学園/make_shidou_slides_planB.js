// 令和８年度 教育指導課訪問（練馬区立みらい青空学園）５校時 指導・講評資料【プランＢ】
// 指導案ベース構成：①アイスブレイク・現状 → ②３人の授業から → ③今後に向けて
// 配色・書体は指定テンプレート（テーマ「黄緑」）に準拠
// 生成: node make_shidou_slides_planB.js
const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9"; // 10.0 x 5.625 inch（テンプレートと同一）
pres.author = "練馬区教育委員会";
pres.title = "令和８年度 教育指導課訪問 ５校時 指導・講評資料（みらい青空学園）プランＢ";

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
   当日の授業写真
   ------------------------------------------------------------
   file に画像ファイル名（この JS と同じフォルダに置く）を書くと、
   スライド７〜９の各授業スライドに自動で埋め込まれます
   （枠に合わせて自動トリミング）。
   null のままだと「写真を貼付」の枠が表示されます。
   例： file: "taguchi.jpg"
   ============================================================ */
const PHOTOS = [
  { file: null, name: "田口　和磨　先生", sub: "保健体育・９年ＡＢ組／水泳" },
  { file: null, name: "東海林　静江　先生", sub: "道徳・特別支援学級Ｄ組／うそ" },
  { file: null, name: "塚本　瑞穂　先生", sub: "外国語・７年Ｂ組／Unit 4" },
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


/* ---------- プランＢ専用パーツ ---------- */

// キーメッセージ（全スライド共通の一文）
const KEYMSG_A = "「違い」を生かす授業";
const KEYMSG_B = "が、９年間の学びをつなぐ。";

// 写真枠（file があれば埋め込み、なければ点線の貼付欄）
function photo(slide, file, x, y, w, h) {
  const abs = file ? path.join(__dirname, file) : null;
  if (abs && fs.existsSync(abs)) {
    slide.addImage({ path: abs, x, y, w, h, sizing: { type: "cover", w, h } });
    slide.addShape(pres.ShapeType.rect, {
      x, y, w, h, fill: { type: "none" }, line: { color: LINE, width: 1 },
    });
    return false;
  }
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h,
    fill: { color: "FBFEFD" }, line: { color: TEAL, width: 1.5, dashType: "dash" },
  });
  slide.addText(
    [
      { text: "授業写真", options: { fontSize: 15, breakLine: true } },
      { text: "（貼付欄）", options: { fontSize: 15 } },
    ],
    {
      x, y, w, h,
      color: TEAL, fontFace: FONT, bold: true, align: "center", valign: "middle",
      margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  return true;
}

// 視点バッジ（Ⅰ・Ⅱ・Ⅲ）
function vbadge(slide, t, x, y, d) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w: d, h: d * 0.72, rectRadius: 0.05, fill: { color: TEAL },
  });
  slide.addText(t, {
    x, y, w: d, h: d * 0.72,
    color: "FFFFFF", fontFace: FONT, fontSize: 11, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
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
      x: 0, y: 0, w: W, h: 2.72,
      fill: { color: TEAL }, color: "FFFFFF",
      fontFace: FONT_T, fontSize: 40, bold: true,
      align: "left", valign: "middle", margin: 0,
      lineSpacingMultiple: 1.2, isTextBox: true,
    }
  );
  panel(s, 0, 2.72, W, 0.72);
  s.addText(
    [
      { text: KEYMSG_A, options: { color: KEY } },
      { text: KEYMSG_B, options: { color: INK } },
    ],
    {
      x: 0.52, y: 2.72, w: 8.96, h: 0.72,
      fontFace: FONT, fontSize: 22, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  s.addText(
    [
      { text: "授業者　塚本　瑞穂　先生（外国語・７年Ｂ組）", options: { breakLine: true } },
      { text: "　　　　田口　和磨　先生（保健体育・９年ＡＢ組）", options: { breakLine: true } },
      { text: "　　　　東海林　静江　先生（道徳・特別支援学級Ｄ組）", options: {} },
    ],
    {
      x: MX, y: 3.72, w: 5.2, h: 1.3,
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
      x: 5.85, y: 3.72, w: 3.63, h: 1.4,
      color: INK, fontFace: FONT_D, fontSize: 18, bold: true,
      align: "right", valign: "top", margin: 0, lineSpacingMultiple: 1.35, isTextBox: true,
    }
  );
  s.addNotes(
    "本日は、貴重な授業を参観させていただき、誠にありがとうございました。練馬区教育委員会指導課の紺多と申します。\n" +
    "１０分間、お時間をいただきます。本日申し上げたいことは、この一文に尽きます。「違い」を生かす授業が、９年間の学びをつなぐ。"
  );
}

/* ---------- スライド２　本日の講評 ---------- */
{
  const s = pres.addSlide();
  bar(s, "本日の講評");
  panel(s, MX, 1.00, BW, 0.95);
  s.addText(
    [
      { text: KEYMSG_A, options: { color: KEY } },
      { text: KEYMSG_B, options: { color: INK } },
    ],
    {
      x: MX + 0.28, y: 1.00, w: BW - 0.56, h: 0.95,
      fontFace: FONT, fontSize: 24, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  [
    ["①", "アイスブレイク　みらい青空学園の現状", 2.20],
    ["②", "３人の授業から", 3.05],
    ["③", "今後に向けて", 3.90],
  ].forEach(([n, t, y]) => {
    badge(s, n, MX + 0.2, y, 0.54);
    s.addText(t, {
      x: MX + 0.92, y, w: 8.0, h: 0.54,
      color: INK, fontFace: FONT, fontSize: 23, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("３人の先生の指導案を手がかりに、本日の授業を読み解きます。", {
    x: MX, y: 4.68, w: BW, h: 0.36,
    color: GRAY, fontFace: FONT, fontSize: 14, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "本日は三点でお話しします。はじめにクイズと、みらい青空学園の現状。次に、３人の先生の授業から。最後に、今後に向けてです。\n" +
    "３人の先生の指導案を手がかりに、本日の授業を読み解いてまいります。"
  );
}

/* ---------- スライド３　アイスブレイク　クイズ出題 ---------- */
{
  const s = pres.addSlide();
  bar(s, "①　アイスブレイク", "クイズ　本校の子どもたちは、どう変わったか");

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
  bar(s, "①　アイスブレイク", "答え　本校では、中学部で上がっている");
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

/* ---------- スライド５　みらい青空学園の現状 ---------- */
{
  const s = pres.addSlide();
  bar(s, "①　みらい青空学園の現状", "質問紙調査から見える強み");
  s.addText("対話と協働に関わる項目が、小学部から中学部にかけて伸びている。", {
    x: 0.35, y: 1.50, w: 9.3, h: 0.34,
    color: INK, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  [
    { q: "友達や周りの人の考えを\n大切にして、協力しながら\n課題の解決に取り組んでいる", a: "85.1", b: "100.0", d: "＋14.9", ref: "中学部は全国 92.4 を上回る" },
    { q: "話し合う活動を通じて、\n自分の考えを深めたり、\n新たな考え方に気付ける", a: "92.5", b: "97.9", d: "＋5.4", ref: "中学部は練馬区 86.6 を上回る" },
    { q: "自分と違う意見について\n考えるのは楽しい", a: "81.4", b: "87.2", d: "＋5.8", ref: "中学部は練馬区 79.8 を上回る" },
  ].forEach((t, i) => {
    const tx = 0.35 + i * 3.18;
    panel(s, tx, 1.92, 2.95, 2.62);
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
    s.addText(t.ref, {
      x: tx + 0.18, y: 3.98, w: 2.59, h: 0.26,
      color: GRAY, fontFace: FONT, fontSize: 10, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  body(s, 0.35, 4.72, 9.3, 0.5,
    "対話し、協働する力が、９年間を通じて[[確かに育っている]]。", 18);
  s.addText("出典：令和８年度 全国学力・学習状況調査　回答結果集計表　肯定的回答の割合", {
    x: 0.35, y: 5.26, w: 9.3, h: 0.24,
    color: GRAY, fontFace: FONT, fontSize: 9.5, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "みらい青空学園の現状です。対話と協働に関わる項目が、小学部から中学部にかけて伸びています。\n" +
    "「友達や周りの人の考えを大切にして、協力しながら課題の解決に取り組んでいますか」は、85.1％から100％へ。中学部は全員が肯定的に答えています。\n" +
    "「話し合う活動を通じて、自分の考えを深めたり、新たな考え方に気付いたりすることができていますか」は、92.5％から97.9％へ。\n" +
    "「自分と違う意見について考えるのは楽しい」は、81.4％から87.2％へ。\n" +
    "９年間を見通した教育の成果が、数値に表れています。ここから本題に入ります。"
  );
}


/* ---------- スライド６　３つの指導案に共通していた視点 ---------- */
{
  const s = pres.addSlide();
  bar(s, "②　３人の授業から", "３つの指導案に、共通して書かれていたこと");
  body(s, MX, 1.55, BW, 0.85,
    "３つの学級は、いずれも[[一人一人の差が大きい]]。\nその差にどう向き合うかが、３つの指導案に書かれていた。", 17);
  [
    { n: "Ⅰ", t: "見通しをもたせる", d: "目標と流れ、単元のゴールを\n子どもと共有する" },
    { n: "Ⅱ", t: "違いに応じる", d: "一人一人の課題・活動・支援に\n選べる幅をもたせる" },
    { n: "Ⅲ", t: "対話で考えを深める", d: "自分の考えをもってから\n他者と交わす" },
  ].forEach((v, i) => {
    const x = 0.35 + i * 3.18;
    panel(s, x, 2.62, 2.95, 1.85);
    vbadge(s, v.n, x + 0.18, 2.76, 0.46);
    s.addText(v.t, {
      x: x + 0.72, y: 2.74, w: 2.05, h: 0.36,
      color: TEAL, fontFace: FONT, fontSize: 14, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(v.d, {
      x: x + 0.18, y: 3.24, w: 2.59, h: 1.0,
      color: INK, fontFace: FONT, fontSize: 13, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2, isTextBox: true,
    });
  });
  body(s, 0.35, 4.62, 9.3, 0.5,
    "この３つを手がかりに、３人の授業を見ていく。", 17);
  s.addNotes(
    "３人の指導案を読ませていただきました。学級はまったく違いますが、共通して書かれていたことがあります。\n" +
    "水泳は泳力差が顕著、外国語は習熟度差が大きい、道徳は特別支援学級。３つの学級は、いずれも一人一人の差が大きい集団です。\n" +
    "その差にどう向き合うか。指導案には、視点Ⅰ 見通しをもたせる、視点Ⅱ 一人一人の違いに応じる、視点Ⅲ 対話で考えを深める、という手だてが書かれていました。この３つを手がかりに、３人の授業を見てまいります。"
  );
}

/* ---------- スライド７〜９　３人の授業 ---------- */
const LESSONS = [
  {
    photo: 0,
    who: "田口　和磨　先生",
    where: "保健体育・第９学年ＡＢ組／水泳（全10時間中の第３時）　バタフライのキック",
    real: "コロナ禍と校舎改築で水泳の授業回数が少ない。\n苦手意識や水への恐怖心をもつ生徒と、スイミングスクール経験者が混在し、泳力差が顕著。",
    plans: [
      { v: "Ⅰ", t: "着替えの前に、本時の目標と授業の流れを確認する" },
      { v: "Ⅱ", t: "泳力別に３段階でコースを分け、補助具の使用を認める" },
      { v: "Ⅱ", t: "得意な生徒に、他の生徒へのアドバイスを担当させる" },
      { v: "Ⅲ", t: "バディで互いのキックを確認し合う" },
    ],
    seen: "確認の視点を３点示し、子どもが[[自分の泳ぎを見る目]]をもてていた。",
  },
  {
    photo: 1,
    who: "東海林　静江　先生",
    where: "道徳・特別支援学級Ｄ組／「ついていい嘘とついてはいけない嘘はどう違うのか？」",
    real: "将来の就労を見据えた「報告・連絡・相談」を大切にしている学級。\n「嘘はつくなと言われているけど、これは嘘がうまくないと勝てない」という生徒の発言が出発点。",
    plans: [
      { v: "Ⅰ", t: "前時「いいこと・わるいこと」と２週続きで組み立てる" },
      { v: "Ⅱ", t: "前時のワークシートを活用し、考えが変わってもよいと保障する" },
      { v: "Ⅲ", t: "自己との対話で整理してから、グループで交流し議論へ" },
      { v: "Ⅲ", t: "発表よりも議論の時間を重視する" },
    ],
    seen: "全員が[[自分の考えをもって]]議論に入れていた。",
  },
  {
    photo: 2,
    who: "塚本　瑞穂　先生",
    where: "外国語・第７学年Ｂ組／Unit 4 Our New Friend（全９時間中の第２時）",
    real: "小学校時から苦手意識のある生徒、母語が英語の生徒、英検２級レベルの生徒まで混在。\n習熟度の差が大きい。",
    plans: [
      { v: "Ⅰ", t: "Today’s Goal と Today’s Plan を提示する" },
      { v: "Ⅰ", t: "単元末はＡＬＴへの発表。目的と相手が明確なゴールを置く" },
      { v: "Ⅱ", t: "発表の前にペアで共有し、発表時の不安を軽減する" },
      { v: "Ⅱ", t: "くじで指名し、１時間で全員に発話の機会をつくる" },
    ],
    seen: "板書を手がかりに、[[誰もが一文をつくれる]]ようにしていた。",
  },
];

LESSONS.forEach((L) => {
  const s = pres.addSlide();
  bar(s, "②　３人の授業から", L.who);
  s.addText(L.where, {
    x: MX, y: 1.42, w: BW, h: 0.30,
    color: GRAY, fontFace: FONT, fontSize: 12, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const blank = photo(s, PHOTOS[L.photo].file, 0.35, 1.80, 4.15, 2.72);
  s.addText("学級の実態（指導案より）", {
    x: 4.72, y: 1.80, w: 4.93, h: 0.28,
    color: TEAL, fontFace: FONT, fontSize: 13, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(L.real, {
    x: 4.72, y: 2.10, w: 4.93, h: 0.76,
    color: INK, fontFace: FONT, fontSize: 11.5, bold: true,
    align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.2, isTextBox: true,
  });
  s.addText("指導案に書かれていた手だて", {
    x: 4.72, y: 2.94, w: 4.93, h: 0.28,
    color: TEAL, fontFace: FONT, fontSize: 13, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  L.plans.forEach((pl, i) => {
    const y = 3.26 + i * 0.33;
    vbadge(s, pl.v, 4.72, y + 0.02, 0.34);
    s.addText(pl.t, {
      x: 5.16, y, w: 4.49, h: 0.30,
      color: INK, fontFace: FONT, fontSize: 11.5, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  panel(s, 0.35, 4.66, 9.3, 0.62);
  body(s, 0.63, 4.76, 8.74, 0.44, "参観して　" + L.seen, 15);
  if (blank) {
    s.addText("※ 点線の枠に当日の授業写真を貼り付けてご使用ください。", {
      x: 0.35, y: 5.34, w: 9.3, h: 0.22,
      color: GRAY, fontFace: FONT, fontSize: 9, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  s.addNotes(
    L.who + "の授業です。" + L.where + "。\n" +
    "指導案には、学級の実態がこう書かれていました。" + L.real.replace(/\n/g, "") + "\n" +
    "そのうえで打たれた手だてが、こちらです。" + L.plans.map((x) => x.t).join("。") + "。\n" +
    "参観して、" + L.seen.replace(/\[\[|\]\]/g, "") + "（写真に触れながら）"
  );
});

/* ---------- スライド１０　３つの授業に共通していたこと ---------- */
{
  const s = pres.addSlide();
  bar(s, "②　３人の授業から", "３つの授業に共通していたこと");
  const cols = ["田口　先生", "東海林　先生", "塚本　先生"];
  const rows = [
    { v: "Ⅰ", t: "見通し", c: ["着替えの前に\n目標と流れを確認", "前時と２週続きで\n組み立てる", "Today’s Goal と\nPlan の提示"] },
    { v: "Ⅱ", t: "違いに応じる", c: ["泳力別の３コース\n補助具を認める", "前時のワークシート\n意見が変わってよい", "発表前のペア共有\nくじ指名で全員発話"] },
    { v: "Ⅲ", t: "対話", c: ["バディで\n互いのキックを確認", "自己内対話のあと\n議論を重視", "ペアで確かめてから\n全体へ"] },
  ];
  const LX = 0.35, LW = 1.55, CW = 2.55, GAP = 0.08;
  cols.forEach((c, i) => {
    s.addText(c, {
      x: LX + LW + GAP + i * (CW + GAP), y: 1.42, w: CW, h: 0.30,
      color: TEAL, fontFace: FONT, fontSize: 13, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  rows.forEach((r, ri) => {
    const y = 1.78 + ri * 0.99;
    s.addShape(pres.ShapeType.roundRect, {
      x: LX, y, w: LW, h: 0.90, rectRadius: 0.05, fill: { color: TEAL },
    });
    s.addText(
      [
        { text: r.v, options: { fontSize: 15, breakLine: true } },
        { text: r.t, options: { fontSize: 12 } },
      ],
      {
        x: LX, y, w: LW, h: 0.90,
        color: "FFFFFF", fontFace: FONT, bold: true,
        align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15, isTextBox: true,
      }
    );
    r.c.forEach((t, ci) => {
      const x = LX + LW + GAP + ci * (CW + GAP);
      panel(s, x, y, CW, 0.90);
      s.addText(t, {
        x: x + 0.10, y, w: CW - 0.20, h: 0.90,
        color: INK, fontFace: FONT, fontSize: 11.5, bold: true,
        align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.2, isTextBox: true,
      });
    });
  });
  body(s, LX, 4.80, 9.3, 0.5,
    "３人とも、差を埋めるのではなく、[[差を前提に授業を設計]]していた。", 18);
  s.addNotes(
    "３つの授業を並べると、このようになります。\n" +
    "見通しは、着替えの前の確認、２週続きの構成、Today’s Goal。\n" +
    "違いに応じる手だては、泳力別のコース、前時のワークシート、発表前のペア共有。\n" +
    "対話は、バディでの確認、自己内対話のあとの議論、ペアから全体へ。\n" +
    "教科も学年も違いますが、３人とも、差を埋めるのではなく、差を前提に授業を設計されていました。これが本校の強みだと受け止めました。"
  );
}

/* ---------- スライド１１　手だては、子どもに何を育てるか ---------- */
{
  const s = pres.addSlide();
  bar(s, "②　３人の授業から", "この手だては、子どもに何を育てるか");
  ["Ⅰ　見通し", "Ⅱ　違いに応じる", "Ⅲ　対話"].forEach((t, i) => {
    chip(s, 0.35 + i * 3.18, 1.48, 2.95, 0.54, t, 15, true);
  });
  s.addText("↓", {
    x: 0.35, y: 2.06, w: 9.3, h: 0.34,
    color: GRAY, fontFace: FONT, fontSize: 18, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  panel(s, 0.35, 2.44, 4.55, 1.98);
  s.addText("自己肯定感", {
    x: 0.57, y: 2.56, w: 4.11, h: 0.36,
    color: TEAL, fontFace: FONT, fontSize: 17, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(
    [
      { text: "「できる子になる」ことではなく", options: { color: INK, breakLine: true } },
      { text: "「成長できる自分を信じること」", options: { color: KEY, breakLine: true } },
      { text: " ", options: { color: INK, fontSize: 8, breakLine: true } },
      { text: "自分の段階で挑戦でき、考えが変わってよいと", options: { color: INK, breakLine: true } },
      { text: "保障されることで育つ。", options: { color: INK } },
    ],
    {
      x: 0.57, y: 2.96, w: 4.11, h: 1.34,
      fontFace: FONT, fontSize: 13, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  panel(s, 5.10, 2.44, 4.55, 1.98);
  s.addText("学習の自己調整", {
    x: 5.32, y: 2.56, w: 4.11, h: 0.36,
    color: TEAL, fontFace: FONT, fontSize: 17, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(
    [
      { text: "学習指導要領が示す", options: { color: INK, breakLine: true } },
      { text: "「学びに向かう力」の中核をなす力", options: { color: KEY, breakLine: true } },
      { text: " ", options: { color: INK, fontSize: 8, breakLine: true } },
      { text: "見通しをもち、振り返り、次につなげる", options: { color: INK, breakLine: true } },
      { text: "経験の積み重ねで育つ。", options: { color: INK } },
    ],
    {
      x: 5.32, y: 2.96, w: 4.11, h: 1.34,
      fontFace: FONT, fontSize: 13, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  body(s, 0.35, 4.60, 9.3, 0.5,
    "本日の授業には、この二つが育つ場面があった。", 17);
  s.addNotes(
    "３つの手だては、子どもに何を育てるのでしょうか。二つあると考えます。\n" +
    "一つは自己肯定感です。「できる子になる」ことではなく、「成長できる自分を信じること」。自分の段階で挑戦でき、考えが変わってもよいと保障されることで育ちます。\n" +
    "もう一つが、学習の自己調整です。学習指導要領が示す「学びに向かう力」の中核をなす力です。見通しをもち、振り返り、次につなげる経験の積み重ねで育ちます。\n" +
    "本日の授業には、この二つが育つ場面がありました。"
  );
}

/* ---------- スライド１２　学習の自己調整が育っていた場面 ---------- */
{
  const s = pres.addSlide();
  bar(s, "②　３人の授業から", "学習の自己調整が育っていた場面");
  panel(s, 0.35, 1.48, 9.3, 0.86);
  s.addText(
    [
      { text: "「学びに向かう力」を支える「主体的に学習に取り組む態度」は、", options: { color: INK, breakLine: true } },
      { text: "　粘り強く取り組む側面　と　", options: { color: INK } },
      { text: "自らの学習を調整しようとする側面", options: { color: KEY } },
      { text: "　で捉える。", options: { color: INK } },
    ],
    {
      x: 0.63, y: 1.48, w: 8.74, h: 0.86,
      fontFace: FONT, fontSize: 14, bold: true,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  [
    { who: "田口　先生", a: "着替えの前に", b: "本時の目標と流れを確認", c: "学習カードで要点を振り返る" },
    { who: "塚本　先生", a: "授業のはじめに", b: "Today’s Goal と Plan を提示", c: "振り返りをワークシートに記入" },
    { who: "東海林　先生", a: "議論に入る前に", b: "前時のワークシートで自分の考えを確認", c: "本時の考えを書き残す" },
  ].forEach((r, i) => {
    const y = 2.46 + i * 0.70;
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.35, y, w: 1.55, h: 0.62, rectRadius: 0.05, fill: { color: TEAL },
    });
    s.addText(r.who, {
      x: 0.35, y, w: 1.55, h: 0.62,
      color: "FFFFFF", fontFace: FONT, fontSize: 12, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    panel(s, 1.98, y, 7.67, 0.62);
    s.addText(
      [
        { text: r.a + "　", options: { color: GRAY, fontSize: 11 } },
        { text: r.b, options: { color: INK, fontSize: 13 } },
        { text: "　→　", options: { color: GRAY, fontSize: 12 } },
        { text: r.c, options: { color: KEY, fontSize: 13 } },
      ],
      {
        x: 2.18, y, w: 7.27, h: 0.62,
        fontFace: FONT, bold: true,
        align: "left", valign: "middle", margin: 0, isTextBox: true,
      }
    );
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.35, y: 4.62, w: 9.3, h: 0.78, rectRadius: 0.07,
    fill: { color: KTINT }, line: { color: KLINE, width: 1.5 },
  });
  s.addText(
    [
      { text: "教師が管理する学習から、", options: { color: INK } },
      { text: "子どもが管理する学習へ。", options: { color: KEY } },
    ],
    {
      x: 0.35, y: 4.62, w: 9.3, h: 0.78,
      fontFace: FONT, fontSize: 22, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  s.addNotes(
    "とくに学習の自己調整について申し上げます。「学びに向かう力」を支える「主体的に学習に取り組む態度」は、粘り強く取り組む側面と、自らの学習を調整しようとする側面の二つで捉えます。\n" +
    "田口先生は、着替えの前に本時の目標と流れを確認され、学習カードで要点を振り返らせていました。\n" +
    "塚本先生は、授業のはじめに Today’s Goal と Plan を示し、振り返りをワークシートに記入させていました。\n" +
    "東海林先生は、議論に入る前に前時のワークシートで自分の考えを確かめさせ、本時の考えを書き残させていました。\n" +
    "いずれも、目標と振り返りが一対で置かれています。\n" +
    "教師が管理する学習から、子どもが管理する学習へ。（間をおく）本日の授業から見えた、いちばん大きな変化です。"
  );
}

/* ---------- スライド１１　今後に向けて　９年間で系統化する ---------- */
{
  const s = pres.addSlide();
  bar(s, "③　今後に向けて", "この３つを、９年間の共通の言葉に");
  body(s, MX, 1.55, BW, 0.6,
    "本日見えた３つの視点は、学年や教科を越えて使える。", 18);
  ["Ⅰ　見通し", "Ⅱ　違いに応じる", "Ⅲ　対話"].forEach((t, i) => {
    chip(s, MX + i * 3.0, 2.22, 2.7, 0.62, t, 15, true);
  });
  panel(s, MX, 3.06, BW, 1.30);
  s.addText(
    [
      { text: "学校経営計画とも重なる", options: { color: TEAL, fontSize: 14, breakLine: true } },
      { text: "　・主体的・対話的で深い学びの実現に向けた授業改善", options: { color: INK, fontSize: 14, breakLine: true } },
      { text: "　・個に応じたきめ細やかな指導・支援の充実", options: { color: INK, fontSize: 14 } },
    ],
    {
      x: MX + 0.28, y: 3.18, w: BW - 0.56, h: 1.1,
      fontFace: FONT, bold: true, align: "left", valign: "top",
      margin: 0, lineSpacingMultiple: 1.3, isTextBox: true,
    }
  );
  body(s, MX, 4.50, BW, 0.6,
    "小竹小学校との校区別協議会で、[[「目指す１５歳の姿」]]として共有したい。", 17);
  s.addNotes(
    "今後に向けてです。本日見えた３つの視点は、学年や教科を越えて使えます。\n" +
    "学校経営計画に掲げられた「主体的・対話的で深い学びの実現に向けた授業改善」「個に応じたきめ細やかな指導・支援の充実」とも重なります。\n" +
    "この３つを、９年間の共通の言葉にしていただきたい。小竹小学校との校区別協議会でも、「目指す１５歳の姿」として共有できるはずです。"
  );
}

/* ---------- 今後に向けて　その言葉で、互いの授業を見合う ---------- */
{
  const s = pres.addSlide();
  bar(s, "③　今後に向けて", "その言葉で、互いの授業を見合う");
  body(s, MX, 1.52, BW, 0.8,
    "本日の３つの授業は、保健体育の９年、道徳のＤ組、外国語の７年。\n教科も学年も違うのに、[[同じ３つの視点]]で語ることができた。", 17);
  s.addShape(pres.ShapeType.roundRect, {
    x: MX, y: 2.36, w: BW, h: 0.78, rectRadius: 0.07,
    fill: { color: KTINT }, line: { color: KLINE, width: 1.5 },
  });
  s.addText("だから、教科を越えて見合える。", {
    x: MX, y: 2.36, w: BW, h: 0.78,
    color: KEY, fontFace: FONT, fontSize: 22, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  panel(s, MX, 3.28, BW, 1.32);
  s.addText(
    [
      { text: "施設一体型だからこそできること", options: { color: TEAL, fontSize: 14, breakLine: true } },
      { text: "　・小学部の教員が中学部を、中学部の教員が小学部を見る", options: { color: INK, fontSize: 14, breakLine: true } },
      { text: "　・９年後の子どもの姿と、９年前の子どもの姿を、同じ校舎で見られる", options: { color: INK, fontSize: 14, breakLine: true } },
      { text: "　・見る観点は、Ⅰ見通し／Ⅱ違いに応じる／Ⅲ対話　の３つでよい", options: { color: INK, fontSize: 14 } },
    ],
    {
      x: MX + 0.28, y: 3.40, w: BW - 0.56, h: 1.12,
      fontFace: FONT, bold: true, align: "left", valign: "top",
      margin: 0, lineSpacingMultiple: 1.3, isTextBox: true,
    }
  );
  body(s, MX, 4.74, BW, 0.6,
    "学校経営計画の[[「乗り入れ授業」「相互の実践に生かす」]]を、この３つで動かしたい。", 16);
  s.addNotes(
    "もう一点、今後に向けてです。\n" +
    "本日の３つの授業は、保健体育の９年、道徳のＤ組、外国語の７年。教科も学年も違います。それでも、同じ３つの視点で語ることができました。\n" +
    "だから、教科を越えて見合えます。ここが施設一体型のいちばんの強みだと考えます。\n" +
    "小学部の教員が中学部を、中学部の教員が小学部を見る。９年後の子どもの姿と、９年前の子どもの姿を、同じ校舎で見られる学校です。見る観点は、見通し、違いに応じる、対話。この３つで足ります。\n" +
    "学校経営計画に掲げられた「乗り入れ授業」「相互の実践に生かす」を、この３つの観点で動かしていただきたいと思います。\n" +
    "子どもが学習を調整するように、教師も授業を調整する。その具体が、互いに見合うことだと考えます。"
  );
}

/* ---------- スライド１３　まとめ ---------- */
{
  const s = pres.addSlide();
  bar(s, "まとめ");
  panel(s, MX, 1.05, BW, 1.15);
  s.addText(
    [
      { text: KEYMSG_A, options: { color: KEY } },
      { text: KEYMSG_B, options: { color: INK } },
    ],
    {
      x: MX + 0.28, y: 1.05, w: BW - 0.56, h: 1.15,
      fontFace: FONT, fontSize: 26, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  body(s, MX, 2.40, BW, 0.5,
    "３人とも、一人一人の差を[[学びの出発点]]にしていた。", 17);
  ["Ⅰ　見通し", "Ⅱ　違いに応じる", "Ⅲ　対話"].forEach((t, i) => {
    chip(s, MX + i * 3.0, 2.96, 2.7, 0.56, t, 14, true);
  });
  s.addText("↓　その先に育つのが", {
    x: MX, y: 3.58, w: BW, h: 0.30,
    color: GRAY, fontFace: FONT, fontSize: 12, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  ["自己肯定感", "学習の自己調整"].forEach((t, i) => {
    chip(s, MX + 0.9 + i * 3.7, 3.94, 3.4, 0.56, t, 15, false);
  });
  body(s, MX, 4.68, BW, 0.5,
    "この積み重ねが、９年間を見通した学びの充実につながる。", 17);
  s.addNotes(
    "まとめます。３人とも、一人一人の差を学びの出発点にされていました。\n" +
    "見通し、違いに応じること、対話。その先に育つのが、自己肯定感と、学習の自己調整です。この積み重ねが、９年間を見通した学びの充実につながります。\n" +
    "「違い」を生かす授業が、９年間の学びをつなぐ。９年間を見通した学びの充実に期待しております。本日は誠にありがとうございました。"
  );
}

const OUT = process.argv[2] || "指導課訪問_５校時指導助言資料_みらい青空学園_プランB.pptx";
pres.writeFile({ fileName: OUT }).then(() => console.log("wrote " + OUT));
