// 令和８年度 教育指導課訪問（練馬区立みらい青空学園）５校時 指導・講評資料
// 生成: node make_shidou_slides.js  →  指導課訪問_５校時指導助言資料_みらい青空学園.pptx
const PptxGenJS = require("pptxgenjs");

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9"; // 10.0 x 5.625 inch（参考資料と同一）
pres.author = "練馬区教育委員会";
pres.title = "令和８年度 教育指導課訪問 ５校時 指導・講評資料（みらい青空学園）";

const FONT = "Meiryo";
const BLUE = "00B0F0"; // 見出しバー・小見出し
const ORANGE = "EB6C15"; // キーワード強調
const GREEN = "07A973"; // 補助強調
const INK = "1A1A1A";
const GRAY = "595959";
const TINT = "EAF6FD"; // カード地色
const W = 10.0;

/* ---------- 共通パーツ ---------- */

// 上部の見出しバー（参考資料と同じ：全幅・水色地・白抜き太字）
function bar(slide, label) {
  slide.addText("　" + label, {
    x: 0, y: 0, w: W, h: 0.63,
    fill: { color: BLUE }, color: "FFFFFF",
    fontFace: FONT, fontSize: 28, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
}

// 「①」などの丸数字バッジ
function badge(slide, n, x, y, d, color) {
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color: color || ORANGE },
  });
  slide.addText(String(n), {
    x, y, w: d, h: d,
    color: "FFFFFF", fontFace: FONT, fontSize: Math.round(d * 46), bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
}

// 授業者カード（薄い水色地・角丸／罫線ストライプは使わない）
function card(slide, x, y, w, h, head, sub, lines) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  slide.addText(head, {
    x: x + 0.16, y: y + 0.13, w: w - 0.32, h: 0.3,
    color: BLUE, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  slide.addText(sub, {
    x: x + 0.16, y: y + 0.43, w: w - 0.32, h: 0.26,
    color: GRAY, fontFace: FONT, fontSize: 11, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  slide.addText(
    lines.map((t, i) => ({
      text: t.text,
      options: { color: t.hi ? ORANGE : INK, breakLine: i !== lines.length - 1 },
    })),
    {
      x: x + 0.16, y: y + 0.74, w: w - 0.32, h: h - 0.88,
      color: INK, fontFace: FONT, fontSize: 13.5, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
      isTextBox: true,
    }
  );
}

// 本文の大きな太字ブロック（[[…]] で橙、{{…}} で緑、<<…>> で水色）
function body(slide, x, y, w, h, text, size, spacing) {
  const runs = [];
  const lines = text.split("\n");
  lines.forEach((line, li) => {
    const parts = line.split(/(\[\[.*?\]\]|\{\{.*?\}\}|<<.*?>>)/g).filter((s) => s !== "");
    if (parts.length === 0) parts.push("");
    parts.forEach((p, pi) => {
      let color = INK, txt = p;
      if (p.startsWith("[[")) { color = ORANGE; txt = p.slice(2, -2); }
      else if (p.startsWith("{{")) { color = GREEN; txt = p.slice(2, -2); }
      else if (p.startsWith("<<")) { color = BLUE; txt = p.slice(2, -2); }
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
    lineSpacingMultiple: spacing || 1.4, isTextBox: true,
  });
}

/* ---------- スライド１　表紙 ---------- */
{
  const s = pres.addSlide();
  s.addText("令和８年度　教育指導課訪問", {
    x: 0.6, y: 0.62, w: 8.8, h: 0.5,
    color: BLUE, fontFace: FONT, fontSize: 22, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("練馬区立みらい青空学園", {
    x: 0.6, y: 1.16, w: 8.8, h: 0.85,
    color: INK, fontFace: FONT, fontSize: 40, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("５校時　指導・講評", {
    x: 0.6, y: 2.02, w: 8.8, h: 0.62,
    color: INK, fontFace: FONT, fontSize: 30, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(
    [
      { text: "授業者　塚本　瑞穂　先生（外国語・７年Ｂ組）", options: { breakLine: true } },
      { text: "　　　　田口　和磨　先生（保健体育・９年ＡＢ組）", options: { breakLine: true } },
      { text: "　　　　東海林　静江　先生（道徳・特別支援学級Ｄ組）", options: {} },
    ],
    {
      x: 0.6, y: 2.86, w: 5.6, h: 1.3,
      color: GRAY, fontFace: FONT, fontSize: 14, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.35, isTextBox: true,
    }
  );
  s.addText(
    [
      { text: "令和８年９月９日（水）", options: { breakLine: true } },
      { text: "練馬区教育委員会", options: { breakLine: true } },
      { text: "指導主事　田口　暁之", options: {} },
    ],
    {
      x: 6.4, y: 3.55, w: 3.1, h: 1.4,
      color: GRAY, fontFace: FONT, fontSize: 15, bold: true,
      align: "right", valign: "middle", margin: 0, lineSpacingMultiple: 1.35, isTextBox: true,
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
  badge(s, "１", 0.7, 1.22, 0.62);
  s.addText("練馬区の小中一貫教育について", {
    x: 1.5, y: 1.22, w: 8.0, h: 0.62,
    color: INK, fontFace: FONT, fontSize: 30, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  badge(s, "２", 0.7, 2.42, 0.62);
  s.addText("本時の授業について", {
    x: 1.5, y: 2.42, w: 8.0, h: 0.62,
    color: INK, fontFace: FONT, fontSize: 30, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("主体的・対話的で深い学びの視点からの授業改善", {
    x: 1.5, y: 3.14, w: 8.0, h: 0.5,
    color: ORANGE, fontFace: FONT, fontSize: 22, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  body(s, 0.7, 4.02, 8.8, 1.0,
    "まず授業そのものではなく、\n[[「学力調査結果から見える子どもの姿」]]から考えたい。", 18, 1.35);
  s.addNotes("講評の柱は２点です。あらかじめ流れをお示しします。導入として、授業の前に学力調査から見える子どもの姿に触れます。");
}

/* ---------- スライド３　練馬区の小中一貫教育（区の目標） ---------- */
{
  const s = pres.addSlide();
  bar(s, "１　練馬区の小中一貫教育");
  body(s, 0.7, 0.95, 8.8, 1.0,
    "練馬区教育委員会の目標\n　[[夢や希望をもち、困難を乗り越える力]]の育成", 20, 1.35);
  body(s, 0.7, 2.15, 8.8, 1.0,
    "その実現のための施策が\n　[[９年間を見通した教育]]（小学校６年間＋中学校３年間）", 19, 1.35);
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.7, y: 3.35, w: 8.6, h: 1.5, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  s.addText(
    [
      { text: "期待される効果", options: { color: BLUE, breakLine: true } },
      { text: "　授業改善による学力・体力の向上　／　豊かな人間性・社会性の育成", options: { color: INK, breakLine: true } },
      { text: "　滑らかな接続による安定した学校生活の実現", options: { color: INK } },
    ],
    {
      x: 0.9, y: 3.5, w: 8.2, h: 1.2,
      fontFace: FONT, fontSize: 17, bold: true,
      align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3, isTextBox: true,
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
  bar(s, "１　練馬区の小中一貫教育");
  s.addText("施設一体型だからこそ高まる教育効果", {
    x: 0.7, y: 0.92, w: 8.8, h: 0.55,
    color: BLUE, fontFace: FONT, fontSize: 24, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const items = ["教員間の\n連携強化", "異学年交流の\n活性化", "小中学校間の\n指導の統一化"];
  items.forEach((t, i) => {
    const x = 0.7 + i * 2.95;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 1.62, w: 2.65, h: 1.28, rectRadius: 0.08,
      fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
    });
    s.addText(t, {
      x, y: 1.62, w: 2.65, h: 1.28,
      color: INK, fontFace: FONT, fontSize: 18, bold: true,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2, isTextBox: true,
    });
  });
  body(s, 0.7, 3.14, 8.8, 2.0,
    "みらい青空学園は、区内２校目の施設一体型として\n令和８年４月に開校した[[開校１年目]]の学校である。\n" +
    "１年生から９年生までが同じ学び舎で学ぶ強みを生かし、\n[[「目指す１５歳の姿」]]を９年間で描いていくことに期待している。", 19, 1.4);
  s.addNotes(
    "施設一体型では、教員間の連携強化、異学年交流の活性化、指導の統一化により、さらに教育効果が高まることが期待されています。\n" +
    "本校は区内２校目の施設一体型として今年４月に開校しました。小竹小学校との校区別協議会も含め、「目指す１５歳の姿」を９年間で共有していただきたいと考えています。"
  );
}

/* ---------- スライド５　学力調査から見える強み ---------- */
{
  const s = pres.addSlide();
  bar(s, "学力調査結果から見える強み");
  body(s, 0.7, 0.95, 4.9, 0.9,
    "本校では、次の３点が\n着実に育っている。", 20, 1.3);
  const strengths = ["学習規律", "学習習慣", "学びへの主体性"];
  strengths.forEach((t, i) => {
    const y = 2.05 + i * 0.92;
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.7, y, w: 4.55, h: 0.74, rectRadius: 0.08,
      fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
    });
    s.addText(t, {
      x: 0.7, y, w: 4.55, h: 0.74,
      color: ORANGE, fontFace: FONT, fontSize: 24, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.rect, {
    x: 5.75, y: 1.0, w: 3.6, h: 3.55,
    fill: { color: "FAFCFE" }, line: { color: BLUE, width: 1.5, dashType: "dash" },
  });
  s.addText(
    [
      { text: "学力調査・意識調査グラフ", options: { color: BLUE, fontSize: 17, breakLine: true } },
      { text: "（貼付欄）", options: { color: BLUE, fontSize: 17, breakLine: true } },
      { text: " ", options: { fontSize: 10, breakLine: true } },
      { text: "全国学力・学習状況調査", options: { color: GRAY, fontSize: 12, breakLine: true } },
      { text: "児童・生徒質問紙調査　ほか", options: { color: GRAY, fontSize: 12 } },
    ],
    {
      x: 5.9, y: 1.15, w: 3.3, h: 3.25,
      fontFace: FONT, bold: true, align: "center", valign: "middle",
      margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  s.addText("※ 本校の調査結果グラフを貼り付けてご使用ください。", {
    x: 5.75, y: 4.62, w: 3.6, h: 0.32,
    color: GRAY, fontFace: FONT, fontSize: 10, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(
    "まず授業そのものではなく、学力調査結果から見える子どもの姿について考えたいと思います。\n" +
    "本校では、学習規律、学習習慣、学びへの主体性が着実に育っていることが、調査結果から読み取れます。\n" +
    "※ここに本校の学力調査・意識調査のグラフを貼り付けてください。"
  );
}

/* ---------- スライド６　なぜその成果が現れているのか ---------- */
{
  const s = pres.addSlide();
  bar(s, "学力調査結果から見える強み");
  body(s, 0.7, 0.92, 8.8, 1.15,
    "学力調査結果は「結果」である。\nでは、[[その結果を生み出した要因は何か]]。", 22, 1.35);
  s.addText("本日の授業から、次の３つが見えてきた。", {
    x: 0.7, y: 2.12, w: 8.8, h: 0.42,
    color: GRAY, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const three = ["自己肯定感", "対話・協働", "学習の自己調整"];
  three.forEach((t, i) => {
    const x = 0.7 + i * 2.95;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.72, w: 2.65, h: 1.72, rectRadius: 0.08,
      fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
    });
    badge(s, ["①", "②", "③"][i], x + 1.055, 2.95, 0.54);
    s.addText(t, {
      x: x + 0.1, y: 3.62, w: 2.45, h: 0.6,
      color: ORANGE, fontFace: FONT, fontSize: 19, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("この３つが、学力調査結果を支える土台になっている。", {
    x: 0.7, y: 4.6, w: 8.8, h: 0.42,
    color: INK, fontFace: FONT, fontSize: 16, bold: true,
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
  bar(s, "２　５校時の授業について");
  badge(s, "①", 0.7, 0.9, 0.6);
  s.addText("今日の授業から見えたもの　自己肯定感", {
    x: 1.45, y: 0.9, w: 8.1, h: 0.6,
    color: BLUE, fontFace: FONT, fontSize: 24, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  body(s, 0.7, 1.78, 8.8, 1.1,
    "これからの時代、学力向上の基盤となるのは\n　[[自己肯定感]]である。", 26, 1.4);
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.7, y: 3.1, w: 8.6, h: 1.85, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  s.addText(
    [
      { text: "自己肯定感とは", options: { color: BLUE, fontSize: 18, breakLine: true } },
      { text: "「できる子になる」ことではなく、", options: { color: INK, fontSize: 22, breakLine: true } },
      { text: "「成長できる自分を信じること」である。", options: { color: ORANGE, fontSize: 22 } },
    ],
    {
      x: 1.0, y: 3.28, w: 8.0, h: 1.5,
      fontFace: FONT, bold: true, align: "left", valign: "top",
      margin: 0, lineSpacingMultiple: 1.35, isTextBox: true,
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
  bar(s, "２　５校時の授業について");
  s.addText("① 自己肯定感を育む授業", {
    x: 0.7, y: 0.78, w: 8.8, h: 0.45,
    color: BLUE, fontFace: FONT, fontSize: 21, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  card(s, 0.35, 1.32, 2.9, 2.6, "田口　和磨　先生", "保健体育・９年ＡＢ組／水泳", [
    { text: "泳力差の大きい集団に" },
    { text: "・段階的な課題設定", hi: true },
    { text: "・泳力別のコース編成", hi: true },
    { text: "・バディでの確認活動", hi: true },
    { text: "確認の視点を３点示し、" },
    { text: "見る目を育てていた。" },
  ]);
  card(s, 3.55, 1.32, 2.9, 2.6, "東海林　静江　先生", "道徳・特別支援学級Ｄ組／うそ", [
    { text: "安心して自分の考えを" },
    { text: "表現できる環境", hi: true },
    { text: "・前時の考えを尊重" },
    { text: "・意見が変わってよい", hi: true },
    { text: "答えは一つではないと" },
    { text: "保障されていた。" },
  ]);
  card(s, 6.75, 1.32, 2.9, 2.6, "塚本　瑞穂　先生", "外国語・７年Ｂ組／Unit 4", [
    { text: "習熟度差の大きい学級で" },
    { text: "・発表前にペアで共有", hi: true },
    { text: "・全員に発話の機会", hi: true },
    { text: "不安を減らし、誰もが" },
    { text: "参加できる場をつくって" },
    { text: "いた。" },
  ]);
  body(s, 0.35, 4.18, 9.3, 1.0,
    "子どもたちは「認められる」よりも、[[「自分で成長を実感する」]]経験を\n積み重ねていた。", 17, 1.3);
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
  bar(s, "２　５校時の授業について");
  badge(s, "②", 0.7, 0.9, 0.6);
  s.addText("今日の授業から見えたもの　対話・協働", {
    x: 1.45, y: 0.9, w: 8.1, h: 0.6,
    color: BLUE, fontFace: FONT, fontSize: 24, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  body(s, 0.7, 1.75, 8.8, 1.0,
    "本校の学力調査結果からは、\n　[[主体的な学びの土台]]が形成されていることが分かる。", 22, 1.4);
  body(s, 0.7, 3.0, 8.8, 1.0,
    "その背景にあるのは、日常的に行われている\n　[[「対話を通した学び」]]である。", 22, 1.4);
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.7, y: 4.2, w: 8.6, h: 0.9, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  s.addText("本日の３つの授業にも、対話が学びを深める場面が表れていた。", {
    x: 0.95, y: 4.2, w: 8.1, h: 0.9,
    color: INK, fontFace: FONT, fontSize: 17, bold: true,
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
  bar(s, "２　５校時の授業について");
  s.addText("② 対話によって学びは深まる", {
    x: 0.7, y: 0.78, w: 8.8, h: 0.45,
    color: BLUE, fontFace: FONT, fontSize: 21, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  card(s, 0.35, 1.32, 2.9, 2.6, "保健体育", "田口　和磨　先生", [
    { text: "バディ同士で" },
    { text: "互いの泳ぎを見て", hi: true },
    { text: "改善点を考えていた。", hi: true },
    { text: "得意な生徒が助言役を" },
    { text: "担い、全員に役割が" },
    { text: "あった。" },
  ]);
  card(s, 3.55, 1.32, 2.9, 2.6, "道徳", "東海林　静江　先生", [
    { text: "「嘘」をテーマに、" },
    { text: "価値観の違いについて", hi: true },
    { text: "議論していた。", hi: true },
    { text: "発表より議論の時間を" },
    { text: "重視した構成だった。" },
  ]);
  card(s, 6.75, 1.32, 2.9, 2.6, "外国語", "塚本　瑞穂　先生", [
    { text: "ペアでのやり取りから" },
    { text: "全体の学びへつなぐ。", hi: true },
    { text: "音読も一人ではなく" },
    { text: "相手とともに行い、", hi: true },
    { text: "表現を確かなものに" },
    { text: "していた。" },
  ]);
  body(s, 0.35, 4.18, 9.3, 1.0,
    "教科は異なっても、共通していたのは\n[[「相手を通して自分を見つめる学び」]]であった。", 17, 1.3);
  s.addNotes(
    "保健体育では、バディで互いの泳ぎを見合い、改善点を考えていました。道徳では「嘘」をテーマに価値観の違いを議論していました。外国語では、ペアでのやり取りを全体の学びにつないでいました。\n" +
    "教科は異なりますが、共通していたのは「相手を通して自分を見つめる学び」です。これが対話・協働の本質だと考えます。"
  );
}

/* ---------- スライド１１　③学習の自己調整 ---------- */
{
  const s = pres.addSlide();
  bar(s, "２　５校時の授業について");
  badge(s, "③", 0.7, 0.9, 0.6);
  s.addText("今日の授業から見えたもの　学習の自己調整", {
    x: 1.45, y: 0.9, w: 8.1, h: 0.6,
    color: BLUE, fontFace: FONT, fontSize: 24, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  body(s, 0.7, 1.7, 8.8, 1.35,
    "これから求められる子どもは、\n　教えられたことを学ぶ子どもではなく、\n　[[自分で学び続ける子ども]]である。", 22, 1.4);
  body(s, 0.7, 3.45, 8.8, 0.6,
    "そのために必要なのが[[学習の自己調整]]である。", 22, 1.35);
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.7, y: 4.2, w: 8.6, h: 0.9, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  s.addText("学習指導要領が示す「学びに向かう力」の中核をなす力である。", {
    x: 0.95, y: 4.2, w: 8.1, h: 0.9,
    color: INK, fontFace: FONT, fontSize: 17, bold: true,
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
  bar(s, "２　５校時の授業について");
  s.addText("③ 自己調整が行われていた授業", {
    x: 0.7, y: 0.78, w: 8.8, h: 0.45,
    color: BLUE, fontFace: FONT, fontSize: 21, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const chips = ["本時の目標", "振り返り", "自己評価", "次時への見通し"];
  chips.forEach((t, i) => {
    const x = 0.35 + i * 2.4;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 1.3, w: 2.2, h: 0.6, rectRadius: 0.1,
      fill: { color: ORANGE },
    });
    s.addText(t, {
      x, y: 1.3, w: 2.2, h: 0.6,
      color: "FFFFFF", fontFace: FONT, fontSize: 15, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("本日の授業では、この４点が大切にされていた。", {
    x: 0.35, y: 2.02, w: 9.3, h: 0.38,
    color: GRAY, fontFace: FONT, fontSize: 14, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  body(s, 0.35, 2.5, 9.3, 1.75,
    "・田口　先生　　着替えの前に目標と流れを確認し、[[学習カードで要点を振り返る]]\n" +
    "・塚本　先生　　Today’s Goal と Plan を提示し、[[振り返りをワークシートに記入]]\n" +
    "・東海林　先生　[[前時のワークシート]]で自分の考えを確かめてから議論へ", 16, 1.55);
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.35, y: 4.32, w: 9.3, h: 0.86, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  s.addText(
    [
      { text: "教師が管理する学習から　", options: { color: INK } },
      { text: "子どもが管理する学習へ", options: { color: ORANGE } },
      { text: "　着実に転換が進んでいる。", options: { color: INK } },
    ],
    {
      x: 0.55, y: 4.32, w: 8.9, h: 0.86,
      fontFace: FONT, fontSize: 18, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    }
  );
  s.addNotes(
    "本日の授業では、本時の目標、振り返り、自己評価、次時への見通しが大切にされていました。\n" +
    "田口先生は着替えの前に目標と流れを確認され、学習カードで要点を振り返らせていました。塚本先生は Today’s Goal と Today’s Plan を示し、振り返りをワークシートに記入させていました。東海林先生は前時のワークシートで自分の考えを確かめてから議論に入っていました。\n" +
    "教師が管理する学習から、子どもが管理する学習へ、着実に転換が進んでいます。"
  );
}

/* ---------- スライド１３　今後の方向性 ---------- */
{
  const s = pres.addSlide();
  bar(s, "今後に向けて");
  body(s, 0.7, 0.92, 8.8, 1.25,
    "[[学力調査結果]]　と　[[今日の授業]]　をつなげて考えると、\n今後さらに伸ばしたいのは", 19, 1.35);
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.7, y: 2.06, w: 8.6, h: 0.9, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  s.addText("「学びを自分事として捉える子ども」である。", {
    x: 0.9, y: 2.06, w: 8.2, h: 0.9,
    color: ORANGE, fontFace: FONT, fontSize: 24, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("そのためには", {
    x: 0.7, y: 3.12, w: 8.8, h: 0.38,
    color: INK, fontFace: FONT, fontSize: 16, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const three = ["自己肯定感", "対話・協働", "学習の自己調整"];
  three.forEach((t, i) => {
    const x = 0.7 + i * 2.95;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 3.56, w: 2.65, h: 0.66, rectRadius: 0.08,
      fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
    });
    s.addText(t, {
      x, y: 3.56, w: 2.65, h: 0.66,
      color: ORANGE, fontFace: FONT, fontSize: 17, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText(
    [
      { text: "を　", options: { color: INK } },
      { text: "９年間で系統的に", options: { color: ORANGE } },
      { text: "　育てていくことが重要である。", options: { color: INK } },
    ],
    {
      x: 0.7, y: 4.42, w: 8.8, h: 0.6,
      fontFace: FONT, fontSize: 19, bold: true,
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
  bar(s, "まとめ");
  body(s, 0.7, 0.92, 8.8, 0.8,
    "学力向上は、[[知識の積み上げだけ]]で実現するものではない。", 19, 1.3);
  s.addText("みらい青空学園では", {
    x: 0.7, y: 1.66, w: 8.8, h: 0.38,
    color: GRAY, fontFace: FONT, fontSize: 15, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const three = ["自己肯定感", "対話・協働", "学習の自己調整"];
  three.forEach((t, i) => {
    const x = 0.7 + i * 2.95;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.08, w: 2.65, h: 0.7, rectRadius: 0.08,
      fill: { color: ORANGE },
    });
    s.addText(t, {
      x, y: 2.08, w: 2.65, h: 0.7,
      color: "FFFFFF", fontFace: FONT, fontSize: 17, bold: true,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("が着実に育成されている。", {
    x: 0.7, y: 2.9, w: 8.8, h: 0.4,
    color: INK, fontFace: FONT, fontSize: 17, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.7, y: 3.42, w: 8.6, h: 1.02, rectRadius: 0.08,
    fill: { color: TINT }, line: { color: "C9E6F5", width: 1 },
  });
  s.addText(
    [
      { text: "これこそが、", options: { color: INK, breakLine: true } },
      { text: "施設一体型小中一貫教育校としての最大の強み", options: { color: ORANGE } },
      { text: "である。", options: { color: INK } },
    ],
    {
      x: 0.9, y: 3.42, w: 8.2, h: 1.02,
      fontFace: FONT, fontSize: 18, bold: true,
      align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.25, isTextBox: true,
    }
  );
  s.addText("今後も、９年間を見通した学びの充実に期待している。", {
    x: 0.7, y: 4.6, w: 8.8, h: 0.5,
    color: INK, fontFace: FONT, fontSize: 19, bold: true,
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
