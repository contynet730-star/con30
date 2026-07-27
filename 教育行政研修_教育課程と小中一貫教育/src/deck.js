// 令和8年度 教育行政研修「教育課程と練馬区の小中一貫教育」
const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.defineLayout({ name: "A4W", width: 13.33, height: 7.5 });
p.layout = "A4W";

// ---- palette ----
const DARK = "10352A";      // 深緑（表紙・扉・結び）
const DARK2 = "17402F";     // 装飾円
const PRIMARY = "1E5945";   // 主緑
const PRIMARY2 = "2E7059";
const SAGE = "5C9A82";      // 補助緑
const LIGHTBG = "EEF4F0";   // 淡緑カード
const CARDLINE = "CFE0D7";
const GOLD = "B67C12";      // 差し色（金）＝資料の黄色網掛けを踏襲
const GOLDDK = "9A6910";
const GOLDFILL = "FBF2DA";
const GOLDLINE = "E7CE93";
const INK = "223029";       // 本文
const MUTED = "6E8078";     // 補足
const WHITE = "FFFFFF";
const JP = "Meiryo";
const NUM = "Calibri";

const ML = 0.6, CW = 12.13, PAGES = 23;

function shadow() { return { type: "outer", color: "9CB3A8", blur: 5, offset: 2, angle: 90, opacity: 0.45 }; }
function dots(s) { // 装飾円モチーフ
  s.addShape(p.ShapeType.ellipse, { x: 11.2, y: -1.5, w: 3.6, h: 3.6, fill: { color: DARK2 }, line: { type: "none" } });
  s.addShape(p.ShapeType.ellipse, { x: 12.3, y: 5.2, w: 3.2, h: 3.2, fill: { color: DARK2 }, line: { type: "none" } });
  s.addShape(p.ShapeType.ellipse, { x: -1.1, y: 5.6, w: 2.6, h: 2.6, fill: { color: DARK2 }, line: { type: "none" } });
}
function head(s, part, title, page) {
  s.addText(String(part), { x: ML, y: 0.5, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: JP, fontSize: 20, bold: true, color: WHITE, fill: { color: PRIMARY }, shape: p.ShapeType.roundRect, rectRadius: 0.08, line: { type: "none" }, margin: 0 });
  s.addText([{ text: "第" + part + "部　", options: { color: GOLDDK, fontSize: 13, bold: true } }, { text: title, options: { color: INK, fontSize: 23, bold: true } }],
    { x: ML + 0.66, y: 0.48, w: CW - 0.66, h: 0.56, align: "left", valign: "middle", fontFace: JP, margin: 0 });
  s.addText("令和8年度 教育行政研修", { x: ML, y: 7.0, w: 6, h: 0.3, align: "left", fontFace: JP, fontSize: 9, color: MUTED });
  s.addText(page + " / " + PAGES, { x: 11.9, y: 7.0, w: 0.83, h: 0.3, align: "right", fontFace: NUM, fontSize: 10, color: MUTED });
}
function divider(part, title, sub) {
  const s = p.addSlide(); s.background = { color: DARK }; dots(s);
  s.addText("第" + part + "部", { x: ML + 0.2, y: 2.35, w: 6, h: 0.7, fontFace: JP, fontSize: 26, bold: true, color: GOLD, charSpacing: 2 });
  s.addText(title, { x: ML + 0.2, y: 3.0, w: 11.5, h: 1.6, fontFace: JP, fontSize: 46, bold: true, color: WHITE });
  s.addText(sub, { x: ML + 0.25, y: 4.7, w: 11.5, h: 0.6, fontFace: JP, fontSize: 17, color: SAGE });
  return s;
}
// numbered card: rounded card + circle badge + bold head + body
function numCard(s, n, x, y, w, h, header, body, o) {
  o = o || {};
  s.addShape(p.ShapeType.roundRect, { x, y, w, h, fill: { color: o.fill || LIGHTBG }, line: { color: o.line || CARDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText(String(n), { x: x + 0.2, y: y + 0.2, w: 0.44, h: 0.44, align: "center", valign: "middle", fontFace: JP, fontSize: 15, bold: true, color: WHITE, fill: { color: o.badge || PRIMARY }, shape: p.ShapeType.ellipse, line: { type: "none" }, margin: 0 });
  s.addText(header, { x: x + 0.78, y: y + 0.17, w: w - 0.95, h: 0.5, align: "left", valign: "middle", fontFace: JP, fontSize: 13.5, bold: true, color: o.headColor || INK, margin: 0 });
  if (body) s.addText(body, { x: x + 0.28, y: y + 0.72, w: w - 0.52, h: h - 0.9, align: "left", valign: "top", fontFace: JP, fontSize: 11.5, color: INK, margin: 0, lineSpacingMultiple: 1.06 });
}
function note(s, txt) { s.addNotes(txt); }

/* ============ SLIDE 1 : 表紙 ============ */
{
  const s = p.addSlide(); s.background = { color: DARK }; dots(s);
  s.addText("令和8年度　教育行政研修", { x: ML + 0.2, y: 1.5, w: 11, h: 0.5, fontFace: JP, fontSize: 18, color: SAGE, charSpacing: 2 });
  s.addText("教育課程と\n練馬区の小中一貫教育", { x: ML + 0.2, y: 2.15, w: 11.8, h: 2.2, fontFace: JP, fontSize: 48, bold: true, color: WHITE, lineSpacingMultiple: 1.05 });
  s.addText("～ 9年間の学びの連続性を、どう支えるか ～", { x: ML + 0.25, y: 4.45, w: 11, h: 0.6, fontFace: JP, fontSize: 19, color: GOLD });
  s.addShape(p.ShapeType.roundRect, { x: ML + 0.2, y: 5.55, w: 6.6, h: 1.15, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() });
  s.addText([
    { text: "練馬区教育委員会　教育指導課\n", options: { fontSize: 13, color: "CFE6DB" } },
    { text: "指導主事　紺多　章一郎", options: { fontSize: 18, bold: true, color: WHITE } }],
    { x: ML + 0.45, y: 5.62, w: 6.2, h: 1.0, valign: "middle", fontFace: JP, align: "left", margin: 0, lineSpacingMultiple: 1.1 });
  s.addText("令和8年7月28日（火）", { x: 8.2, y: 5.9, w: 4.5, h: 0.5, align: "right", fontFace: JP, fontSize: 15, color: SAGE });
  note(s, "本日はお集まりいただきありがとうございます。教育指導課・指導主事の紺多です。この30分で、『教育課程とその管理』、そして練馬区の『小中一貫教育』の全体像を、指導主事の視点からお話しします。説明25分、質疑5分の予定です。");
}

/* ============ SLIDE 2 : 本日のねらい・流れ ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE };
  s.addText("本日のねらいと流れ", { x: ML, y: 0.5, w: CW, h: 0.7, fontFace: JP, fontSize: 27, bold: true, color: INK });
  // ねらい
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.4, w: CW, h: 1.15, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.06, shadow: shadow() });
  s.addText("ねらい", { x: ML + 0.3, y: 1.58, w: 1.6, h: 0.4, fontFace: JP, fontSize: 13, bold: true, color: GOLDDK, margin: 0 });
  s.addText("「教育課程」と「その管理」における指導主事・教育委員会の役割を理解し、その具体である「練馬区の小中一貫教育」の全体像をつかむ。", { x: ML + 0.3, y: 1.95, w: CW - 0.6, h: 0.5, fontFace: JP, fontSize: 14.5, color: INK, margin: 0 });
  // 流れ 2枚
  numCard(s, 1, ML, 2.85, 5.9, 1.65, "教育課程とその管理", "教育課程とは何か／その「管理」で指導主事は何をするのか", { fill: LIGHTBG });
  s.addText("約10分", { x: ML + 4.5, y: 2.95, w: 1.3, h: 0.35, align: "right", fontFace: JP, fontSize: 11, color: MUTED, margin: 0 });
  numCard(s, 2, 6.83, 2.85, 5.9, 1.65, "練馬区が取り組む小中一貫教育", "背景・ねらい・練馬区の仕組みと令和8年度の取組", { fill: LIGHTBG });
  s.addText("約10分", { x: 11.33, y: 2.95, w: 1.3, h: 0.35, align: "right", fontFace: JP, fontSize: 11, color: MUTED, margin: 0 });
  // 質疑
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 4.75, w: CW, h: 0.7, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.06 });
  s.addText([{ text: "質疑応答　", options: { bold: true, color: PRIMARY, fontSize: 13.5 } }, { text: "約5分（説明のあとに）", options: { color: INK, fontSize: 13 } }], { x: ML + 0.3, y: 4.75, w: CW - 0.6, h: 0.7, valign: "middle", fontFace: JP, margin: 0 });
  // bridge
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 5.7, w: CW, h: 0.95, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.06, shadow: shadow() });
  s.addText([{ text: "この2つは別々の話ではありません。　", options: { color: "CFE6DB", fontSize: 13 } }, { text: "「9年間の学びの連続性」でひとつにつながります。", options: { color: WHITE, fontSize: 15, bold: true } }], { x: ML + 0.35, y: 5.7, w: CW - 0.7, h: 0.95, valign: "middle", fontFace: JP, margin: 0 });
  note(s, "本日のねらいと流れです。前半は教育課程とその管理、後半は練馬区の小中一貫教育。各10分、最後に5分質疑。ポイントは、この2つが『9年間の学びの連続性』という一本の軸でつながっている、ということです。ここを意識しながら聞いてください。");
}

/* ============ SLIDE 3 : 第1部 扉 ============ */
{
  const s = divider(1, "教育課程とその管理", "～ 指導主事は、教育課程をどう支えるのか ～");
  note(s, "それでは第1部、教育課程とその管理です。まず『教育課程とは何か』を確認し、指導主事が行う『管理』の中身、学校訪問での指導・助言のポイントまでを見ていきます。");
}

/* ============ SLIDE 4 : 導入の問い ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 1, "はじめに ― 2つの問い", 4);
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.7, w: 5.9, h: 2.5, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText("Q1", { x: ML + 0.35, y: 2.0, w: 1.3, h: 0.7, fontFace: JP, fontSize: 30, bold: true, color: PRIMARY, margin: 0 });
  s.addText("「教育課程」とは、\n何でしょうか？", { x: ML + 0.35, y: 2.75, w: 5.2, h: 1.2, fontFace: JP, fontSize: 22, bold: true, color: INK, margin: 0, lineSpacingMultiple: 1.05 });
  s.addShape(p.ShapeType.roundRect, { x: 6.83, y: 1.7, w: 5.9, h: 2.5, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText("Q2", { x: 7.18, y: 2.0, w: 1.3, h: 0.7, fontFace: JP, fontSize: 30, bold: true, color: GOLDDK, margin: 0 });
  s.addText("その「管理」とは、\n誰が・何をすること？", { x: 7.18, y: 2.75, w: 5.3, h: 1.2, fontFace: JP, fontSize: 22, bold: true, color: INK, margin: 0, lineSpacingMultiple: 1.05 });
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 4.55, w: CW, h: 1.5, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.07, shadow: shadow() });
  s.addText("今日は、この2つを「指導主事の視点」から一緒に考えます。", { x: ML + 0.4, y: 4.55, w: CW - 0.8, h: 1.5, valign: "middle", fontFace: JP, fontSize: 19, bold: true, color: WHITE, margin: 0 });
  note(s, "まず、みなさんに問いかけです。教育課程とは何か。そして、その『管理』とは、誰が、何をすることか。日々教えていても、あらためて言葉にすると難しいところです。5人と少人数ですので、途中で『こうかな』と思ったことがあれば遠慮なくどうぞ。今日はこれを指導主事の視点から整理します。");
}

/* ============ SLIDE 5 : 地教行法21条 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 1, "教育委員会の職務権限と指導主事", 5);
  s.addText("教育に関する事務は、教育委員会が管理・執行します（19の職務権限）。その一つが――", { x: ML, y: 1.55, w: CW, h: 0.5, fontFace: JP, fontSize: 14.5, color: INK });
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 2.2, w: CW, h: 1.7, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1.5 }, rectRadius: 0.07, shadow: shadow() });
  s.addText("第五号", { x: ML + 0.35, y: 2.45, w: 2.2, h: 0.55, fontFace: JP, fontSize: 18, bold: true, color: GOLDDK, margin: 0 });
  s.addText("学校の組織編制、教育課程、学習指導、生徒指導及び職業指導に関すること", { x: ML + 0.35, y: 3.0, w: CW - 0.7, h: 0.8, fontFace: JP, fontSize: 19, bold: true, color: INK, margin: 0, lineSpacingMultiple: 1.05 });
  // 指導主事の位置づけ
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 4.25, w: CW, h: 1.65, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.07, shadow: shadow() });
  s.addText([{ text: "指導主事とは　", options: { bold: true, color: GOLD, fontSize: 15 } }, { text: "この「教育課程」に関する事務を、学校現場に入って専門的に支える職。\n教育委員会の管理を、指導・助言というかたちで学校の教育活動に届ける役割を担います。", options: { color: WHITE, fontSize: 14.5 } }], { x: ML + 0.4, y: 4.25, w: CW - 0.8, h: 1.65, valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.1 });
  s.addText("資料1／地方教育行政の組織及び運営に関する法律 第21条（教育委員会の職務権限）", { x: ML, y: 6.05, w: CW, h: 0.35, fontFace: JP, fontSize: 10, italic: true, color: MUTED });
  note(s, "教育課程は、そもそも誰の仕事か。地方教育行政法21条は、教育委員会の職務権限を19号に定めています。学校の設置管理、人事、教科書…その第五号に『教育課程』があります。この教育課程に関する事務を、学校現場に入って専門的に支えるのが私たち指導主事です。お手元の資料1で、条文全体もご確認ください。");
}

/* ============ SLIDE 6 : 教育課程とは ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 1, "教育課程とは", 6);
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.6, w: CW, h: 1.15, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText([{ text: "各学校の使命　", options: { bold: true, color: PRIMARY, fontSize: 14 } }, { text: "教育活動を通して教育目標を実現し、児童・生徒に「生きる力」を身に付けさせること。", options: { color: INK, fontSize: 15 } }], { x: ML + 0.35, y: 1.6, w: CW - 0.7, h: 1.15, valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.1 });
  // flow: 教育目標 → 教育課程 → 生きる力
  const fy = 3.35, fw = 3.5, fh = 1.5;
  const box = (x, t, sub, fill, tc) => { s.addShape(p.ShapeType.roundRect, { x, y: fy, w: fw, h: fh, fill: { color: fill }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() }); s.addText([{ text: t + "\n", options: { fontSize: 17, bold: true, color: tc } }, { text: sub, options: { fontSize: 11, color: tc } }], { x: x + 0.15, y: fy, w: fw - 0.3, h: fh, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.05 }); };
  box(ML, "学校の教育目標", "地域・児童生徒の実態を踏まえて", SAGE, WHITE);
  s.addText("→", { x: ML + fw, y: fy, w: 0.71, h: fh, align: "center", valign: "middle", fontFace: JP, fontSize: 34, bold: true, color: PRIMARY, margin: 0 });
  box(ML + fw + 0.71, "教育課程", "＝目標実現のための全体計画", PRIMARY, WHITE);
  s.addText("→", { x: ML + 2 * fw + 0.71, y: fy, w: 0.71, h: fh, align: "center", valign: "middle", fontFace: JP, fontSize: 34, bold: true, color: PRIMARY, margin: 0 });
  box(ML + 2 * fw + 1.42, "「生きる力」", "の育成", GOLD, WHITE);
  s.addText("教育課程は、思いつきの寄せ集めではなく、目標から逆算して編成する「学校全体の設計図」です。", { x: ML, y: 5.35, w: CW, h: 0.6, fontFace: JP, fontSize: 13.5, color: INK });
  note(s, "教育課程とは何か。出発点は各学校の使命――教育活動を通して教育目標を実現し、子供たちに『生きる力』を身に付けさせることです。教育課程は、その目標を実現するための全体計画、いわば学校全体の設計図です。地域や児童・生徒の実態を踏まえ、目標から逆算して編成する。ここが大事な点です。");
}

/* ============ SLIDE 7 : 教育課程の管理＝PDCA ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 1, "教育課程の「管理」とは ― PDCAを回すこと", 7);
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.5, w: CW, h: 1.25, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText("「学校の教育目標を実現するために、児童・生徒や地域の実態を踏まえ、教育課程の編成・実施を充実したものにするよう、組織的かつ計画的にその状況を把握し、評価・改善すること」", { x: ML + 0.35, y: 1.5, w: CW - 0.7, h: 1.25, valign: "middle", fontFace: JP, fontSize: 13.5, italic: true, color: INK, margin: 0, lineSpacingMultiple: 1.12 });
  // PDCA boxes
  const py = 3.15, pw = 2.72, ph = 1.35, gap = 0.42;
  const labels = [["編成", "Plan（計画）"], ["実施", "Do"], ["評価", "Check"], ["改善", "Action"]];
  const cols = [PRIMARY, PRIMARY2, SAGE, GOLD];
  labels.forEach((L, i) => {
    const x = ML + i * (pw + gap);
    s.addShape(p.ShapeType.roundRect, { x, y: py, w: pw, h: ph, fill: { color: cols[i] }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() });
    s.addText([{ text: L[0] + "\n", options: { fontSize: 20, bold: true, color: WHITE } }, { text: L[1], options: { fontSize: 11, color: "E7F1EB" } }], { x, y: py, w: pw, h: ph, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.05 });
    if (i < 3) s.addText("→", { x: x + pw, y: py, w: gap, h: ph, align: "center", valign: "middle", fontFace: JP, fontSize: 24, bold: true, color: PRIMARY, margin: 0 });
  });
  s.addText("↻　改善を次の「編成」へ返し、サイクルとして回し続ける", { x: ML, y: 4.62, w: CW, h: 0.4, align: "center", fontFace: JP, fontSize: 12.5, color: PRIMARY });
  // 指導主事の役割
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 5.15, w: CW, h: 0.9, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText([{ text: "指導主事の役割　", options: { bold: true, color: GOLDDK, fontSize: 14 } }, { text: "このPDCAサイクルが各学校で有効に機能するよう、指導・監督する。", options: { color: INK, fontSize: 14.5 } }], { x: ML + 0.35, y: 5.15, w: CW - 0.7, h: 0.9, valign: "middle", fontFace: JP, margin: 0 });
  s.addText("資料2／指導主事実務の手引（令和5年3月 東京都教職員研修センター）", { x: ML, y: 6.12, w: CW, h: 0.32, fontFace: JP, fontSize: 10, italic: true, color: MUTED });
  note(s, "『管理』というと堅く聞こえますが、中身はPDCAを回すことです。定義はこの通り――目標実現のために、実態を踏まえ、編成・実施を充実させるよう、組織的・計画的に状況を把握し、評価・改善する。編成＝計画、実施、評価、改善。改善を次の編成へ返し、回し続ける。指導主事は、このサイクルが各学校で有効に機能するよう指導・監督します。詳細は資料2にあります。");
}

/* ============ SLIDE 8 : 指導主事による管理（4視点）============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 1, "指導主事による教育課程の管理 ― 4つの視点", 8);
  const x1 = ML, x2 = 6.83, y1 = 1.6, y2 = 3.9, w = 5.9, h = 2.15;
  numCard(s, 1, x1, y1, w, h, "全教職員の参画", "編成・実施の意義を全教職員が理解し、組織の一員として参加する。それが校長の経営方針・学校の教育目標の実現につながる。");
  numCard(s, 2, x2, y1, w, h, "体制の整備", "教職員の人的配置に配慮し、校内の「編成・実施・評価・改善」の体制を整える。");
  numCard(s, 3, x1, y2, w, h, "PDCAに即した指導・助言", "個人として、学年組織として教育課程を管理できるよう助言・調整し、教職員の意欲と専門職としての能力を高める。");
  numCard(s, 4, x2, y2, w, h, "進行管理", "教育目標の実現状況を把握し、必要に応じて前の段階へフィードバック・計画修正し、充実した教育活動につなげる。", { fill: GOLDFILL, line: GOLDLINE, badge: GOLD });
  note(s, "では指導主事は具体的に何をするか。学校を組織的に動かすための4つの視点です。1つ目、全教職員の参画。2つ目、編成から改善までを回す校内体制の整備。3つ目、PDCAに即した指導・助言で、先生方の意欲と専門性を高める。4つ目、進行管理。実現状況を見て、必要なら前の段階へ戻して計画を修正する。管理は縛ることではなく、学校が回る仕組みを支えることです。");
}

/* ============ SLIDE 9 : 学校訪問の指導助言 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 1, "学校訪問における指導・助言のポイント", 9);
  const x1 = ML, x2 = 6.83, y1 = 1.6, y2 = 3.9, w = 5.9, h = 2.15;
  numCard(s, 1, x1, y1, w, h, "ねらいの事前把握と評価", "行事や参観する授業のねらいを事前に十分把握し、児童・生徒の学習状況や指導内容を「届出教育課程」に照らして評価する。");
  numCard(s, 2, x2, y1, w, h, "承認と助言", "成果・課題を、児童生徒の学習状況・指導法改善の姿・組織的に取り組む教員の姿から具体的に捉え、認め、助言する。");
  numCard(s, 3, x1, y2, w, h, "具体例の提供", "各学校が教育活動を一層効果的に進められるよう、具体例や他校の実践例を提供する。");
  numCard(s, 4, x2, y2, w, h, "好事例の普及", "各学校が創意工夫した取組のよさや努力を、他校にも広めていく。", { fill: GOLDFILL, line: GOLDLINE, badge: GOLD });
  note(s, "管理が具体的に表れるのが学校訪問です。4つのポイント。1つ目、ねらいを事前に把握し、実際の授業を『届出教育課程』に照らして評価する。2つ目、良い点は具体的に認め、課題は助言する。3つ目、他校の実践例など具体を提供する。4つ目、良い取組を他校へ広める。指導主事は、評価者であると同時に、学校と学校をつなぐハブでもあります。");
}

/* ============ SLIDE 10 : ブリッジ ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 1, "第1部から第2部へ ― 管理が目指すもの", 10);
  s.addText("教育課程の管理が最終的に目指すのは――", { x: ML, y: 1.6, w: CW, h: 0.5, align: "center", fontFace: JP, fontSize: 16, color: INK });
  s.addShape(p.ShapeType.roundRect, { x: 2.4, y: 2.15, w: 8.5, h: 1.2, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1.5 }, rectRadius: 0.1, shadow: shadow() });
  s.addText("系統性・連続性のある「9年間の学び」", { x: 2.4, y: 2.15, w: 8.5, h: 1.2, align: "center", valign: "middle", fontFace: JP, fontSize: 26, bold: true, color: GOLDDK, margin: 0 });
  s.addText("▼", { x: 5.66, y: 3.5, w: 2, h: 0.6, align: "center", fontFace: JP, fontSize: 26, bold: true, color: PRIMARY, margin: 0 });
  s.addText("小学校と中学校は、9年間を通して学びの連続性を確保することが一層求められている", { x: ML, y: 4.15, w: CW, h: 0.5, align: "center", fontFace: JP, fontSize: 14, color: INK });
  s.addShape(p.ShapeType.roundRect, { x: 2.0, y: 4.8, w: 9.3, h: 1.35, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.1, shadow: shadow() });
  s.addText([{ text: "その連続性を、小・中の枠を越えて区全体で実現する取組が\n", options: { color: "CFE6DB", fontSize: 14 } }, { text: "→　練馬区の小中一貫教育", options: { color: WHITE, fontSize: 26, bold: true } }], { x: 2.0, y: 4.8, w: 9.3, h: 1.35, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.15 });
  note(s, "ここまでが第1部です。教育課程の管理が最終的に目指すのは、系統性・連続性のある『9年間の学び』です。小学校と中学校は、9年間を通して学びの連続性を確保することが、いま一層求められています。その連続性を、一つの学校の中だけでなく、小・中の枠を越えて区全体で実現しようという取組が、練馬区の小中一貫教育です。ここから第2部に入ります。");
}

/* ============ SLIDE 11 : 第2部 扉 ============ */
{
  const s = divider(2, "練馬区が取り組む小中一貫教育", "～ 9年間の連続性を、区全体でどう実現するか ～");
  note(s, "第2部、練馬区が取り組む小中一貫教育です。まず、なぜ必要とされるのか、その背景から見ていきます。");
}

/* ============ SLIDE 12 : なぜ小中一貫か（背景）============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "なぜ小中一貫教育が求められるのか", 12);
  s.addText("戦後の「6・3制」以来、小学校と中学校にそれぞれの学校文化が育ってきた", { x: ML, y: 1.55, w: CW, h: 0.5, fontFace: JP, fontSize: 14.5, color: INK });
  const by = 2.2, bw = 4.9, bh = 2.0;
  s.addShape(p.ShapeType.roundRect, { x: ML, y: by, w: bw, h: bh, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText([{ text: "小学校文化\n", options: { fontSize: 18, bold: true, color: PRIMARY } }, { text: "学級担任制／具体的な指導\n生活を丸ごと見る", options: { fontSize: 12.5, color: INK } }], { x: ML, y: by, w: bw, h: bh, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText("？？", { x: ML + bw, y: by, w: 1.53, h: bh, align: "center", valign: "middle", fontFace: JP, fontSize: 30, bold: true, color: GOLD, margin: 0 });
  s.addShape(p.ShapeType.roundRect, { x: ML + bw + 1.53, y: by, w: bw, h: bh, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText([{ text: "中学校文化\n", options: { fontSize: 18, bold: true, color: PRIMARY } }, { text: "教科担任制／抽象的な指導\n定期試験・部活動", options: { fontSize: 12.5, color: INK } }], { x: ML + bw + 1.53, y: by, w: bw, h: bh, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.15 });
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 4.5, w: CW, h: 1.55, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText([
    { text: "文化の違い　→　", options: { bold: true, color: GOLDDK, fontSize: 13 } },
    { text: "児童・生徒観や指導観の差異、小・中教員間の相互理解不足を生む。課題を「小学校は小学校だけ、中学校は中学校だけ」で抱えがちに。\n", options: { color: INK, fontSize: 13 } },
    { text: "少子化・核家族化　→　", options: { bold: true, color: GOLDDK, fontSize: 13 } },
    { text: "手本にしたい先輩を身近に見る機会も少なくなっている。", options: { color: INK, fontSize: 13 } }],
    { x: ML + 0.35, y: 4.5, w: CW - 0.7, h: 1.55, valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.2 });
  note(s, "なぜ必要か。戦後の6・3制以来、小学校文化・中学校文化と呼ばれる、それぞれの学校文化が育ちました。学級担任制か教科担任制か、具体的な指導か抽象的か。この違いが、児童・生徒観や指導観の差異、そして小・中の先生方の相互理解不足を生みます。すると課題を、小学校だけ・中学校だけで抱え込みがちになる。加えて少子化・核家族化で、手本になる先輩を身近に見る機会も減っています。");
}

/* ============ SLIDE 13 : 小中の主な差異6点 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "小・中学校段階の主な差異", 13);
  const data = [
    ["①", "指導体制", "学級担任制 ↔ 教科担任制"],
    ["②", "指導方法", "具体的 ↔ 抽象的"],
    ["③", "家庭学習", "進め方・分量の考え方の違い"],
    ["④", "評価方法", "定期試験の有無 など"],
    ["⑤", "生活指導の手法", "きめ細かさ・任せ方の違い"],
    ["⑥", "部活動の有無", "放課後の過ごし方の違い"]];
  const cw = 3.84, ch = 1.85, gx = 0.3, gy = 0.28, ox = ML, oy = 1.7;
  data.forEach((d, i) => {
    const c = i % 3, r = Math.floor(i / 3);
    const x = ox + c * (cw + gx), y = oy + r * (ch + gy);
    s.addShape(p.ShapeType.roundRect, { x, y, w: cw, h: ch, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
    s.addText(d[0], { x: x + 0.2, y: y + 0.18, w: 0.8, h: 0.6, fontFace: JP, fontSize: 24, bold: true, color: SAGE, margin: 0 });
    s.addText(d[1], { x: x + 0.95, y: y + 0.2, w: cw - 1.1, h: 0.6, valign: "middle", fontFace: JP, fontSize: 15.5, bold: true, color: INK, margin: 0 });
    s.addText(d[2], { x: x + 0.25, y: y + 0.95, w: cw - 0.5, h: 0.7, fontFace: JP, fontSize: 11.5, color: INK, margin: 0, lineSpacingMultiple: 1.05 });
  });
  s.addText("どちらが良い・悪いではなく、「差異の大きさ」への配慮が必要　（文部科学省「小中一貫した教育課程の編成・実施に関する手引き」H28 より）", { x: ML, y: 6.05, w: CW, h: 0.4, fontFace: JP, fontSize: 11.5, italic: true, color: MUTED });
  note(s, "差異を具体的に見ると、文部科学省の手引きは6点を挙げています。指導体制、指導方法、家庭学習、評価方法、生活指導の手法、部活動の有無。大事なのは、どちらが良い悪いではないということ。発達段階に応じた独自性は当然です。問題は『差異の大きさ』そのもの。ここへの配慮が要ります。");
}

/* ============ SLIDE 14 : 不登校データ ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "練馬区でも ― 令和6年度 不登校児童生徒数", 14);
  s.addImage({ path: __dirname + "/../assets/futoko_data.png", x: ML, y: 1.55, w: 9.4, h: 9.4 / 3.456, shadow: shadow() });
  // callout column
  s.addShape(p.ShapeType.roundRect, { x: 10.25, y: 1.55, w: 2.48, h: 4.2, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText([
    { text: "小6\n", options: { fontSize: 12, color: INK } },
    { text: "217人\n", options: { fontSize: 22, bold: true, color: PRIMARY } },
    { text: "▼\n", options: { fontSize: 16, color: GOLD, bold: true } },
    { text: "中1\n", options: { fontSize: 12, color: INK } },
    { text: "235人\n", options: { fontSize: 22, bold: true, color: GOLDDK } },
    { text: "→ 中2 322 → 中3 393", options: { fontSize: 12, bold: true, color: INK } }],
    { x: 10.25, y: 1.7, w: 2.48, h: 3.9, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.05 });
  s.addText("中学校進学後に大きく増加し、学年が上がるほど積み上がる。いわゆる「中1ギャップ」。環境の大きな変化への配慮が必要です。", { x: ML, y: 4.55, w: 9.4, h: 1.4, fontFace: JP, fontSize: 14, color: INK, valign: "top", lineSpacingMultiple: 1.15 });
  note(s, "これは他人事ではありません。令和6年度の練馬区の不登校児童生徒数です。小学校では学年とともに緩やかに増え、小6で217人。それが中1で235人、中2で322人、中3で393人と、中学校で大きく跳ね上がります。環境が大きく変わる中1でつまずく、いわゆる中1ギャップ。この段差をどう滑らかにするかが問われています。");
}

/* ============ SLIDE 15 : 中1ギャップ／努力の限界 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "個々の努力には限界 ― 中学校区単位へ", 15);
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.55, w: CW, h: 1.05, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText("発達段階に応じた独自性は当然。しかし、その差異が生徒に精神的・身体的な負担を生じさせることがある（＝中1ギャップ問題）。", { x: ML + 0.35, y: 1.55, w: CW - 0.7, h: 1.05, valign: "middle", fontFace: JP, fontSize: 14.5, color: INK, margin: 0, lineSpacingMultiple: 1.1 });
  // escalate
  const steps = [["教員一人一人の努力", SAGE], ["学年単位の努力", SAGE], ["学校単位の努力", SAGE], ["中学校区単位の取組", PRIMARY]];
  const sw = 2.72, sh = 1.5, sy = 3.1, gp = 0.42;
  steps.forEach((st, i) => {
    const x = ML + i * (sw + gp);
    s.addShape(p.ShapeType.roundRect, { x, y: sy, w: sw, h: sh, fill: { color: st[1] }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() });
    s.addText(st[0], { x: x + 0.1, y: sy, w: sw - 0.2, h: sh, align: "center", valign: "middle", fontFace: JP, fontSize: 14.5, bold: true, color: WHITE, margin: 0 });
    if (i < 3) s.addText("→", { x: x + sw, y: sy, w: gp, h: sh, align: "center", valign: "middle", fontFace: JP, fontSize: 22, bold: true, color: PRIMARY, margin: 0 });
  });
  s.addText("← 個々の対応には限界がある", { x: ML, y: 4.7, w: 8.5, h: 0.4, fontFace: JP, fontSize: 12.5, color: MUTED });
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 5.25, w: CW, h: 0.85, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText("だからこそ ― 差異の大きさに配慮し、小・中で学習方法の共通理解を図りながら、円滑に接続する取組が必要。", { x: ML + 0.35, y: 5.25, w: CW - 0.7, h: 0.85, valign: "middle", fontFace: JP, fontSize: 14, bold: true, color: GOLDDK, margin: 0 });
  note(s, "誤解のないように――発達段階に応じた独自性そのものは当然のものです。問題は、その差異が生徒に精神的・身体的な負担を生じさせてしまうこと。これが中1ギャップです。そして、教員一人一人、学年、学校、それぞれの努力だけでは対応に限界があります。だからこそ、中学校区という単位で、小・中が学習方法の共通理解を図りながら円滑に接続する。ここに小中一貫教育の必然性があります。");
}

/* ============ SLIDE 16 : 定義と3目的 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "練馬区の小中一貫教育 ― 定義と3つの目的", 16);
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.55, w: CW, h: 1.1, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.07, shadow: shadow() });
  s.addText([{ text: "定義　", options: { bold: true, color: GOLD, fontSize: 14 } }, { text: "義務教育の9年間を見通した教育課程のもとで実施する教育活動（施設一体型／学校施設が離れた連携型を含む）。", options: { color: WHITE, fontSize: 14.5 } }], { x: ML + 0.35, y: 1.55, w: CW - 0.7, h: 1.1, valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.1 });
  s.addText("目指すもの（3点）", { x: ML, y: 2.85, w: CW, h: 0.4, fontFace: JP, fontSize: 14, bold: true, color: INK });
  const gw = 3.84, gh = 2.55, gy = 3.35, gx = 0.3;
  const goals = [["1", "学力・体力の向上", "授業改善による"], ["2", "豊かな人間性・社会性の育成", "連携指導による"], ["3", "安定した学校生活", "滑らかな接続による"]];
  goals.forEach((g, i) => {
    const x = ML + i * (gw + gx);
    s.addShape(p.ShapeType.roundRect, { x, y: gy, w: gw, h: gh, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
    s.addText(g[0], { x: x + (gw - 0.7) / 2, y: gy + 0.3, w: 0.7, h: 0.7, align: "center", valign: "middle", fontFace: JP, fontSize: 20, bold: true, color: WHITE, fill: { color: i === 2 ? GOLD : PRIMARY }, shape: p.ShapeType.ellipse, line: { type: "none" }, margin: 0 });
    s.addText(g[2], { x: x + 0.2, y: gy + 1.15, w: gw - 0.4, h: 0.4, align: "center", fontFace: JP, fontSize: 12, color: MUTED, margin: 0 });
    s.addText(g[1], { x: x + 0.2, y: gy + 1.5, w: gw - 0.4, h: 0.9, align: "center", valign: "top", fontFace: JP, fontSize: 15.5, bold: true, color: INK, margin: 0, lineSpacingMultiple: 1.05 });
  });
  note(s, "練馬区の小中一貫教育の定義です。義務教育の9年間を見通した教育課程のもとで実施する教育活動。大泉桜学園のような施設一体型だけでなく、校舎が離れている学校どうしが連携するかたちも含みます。目指すものは3つ。授業改善による学力・体力の向上。連携指導による豊かな人間性・社会性の育成。そして滑らかな接続による安定した学校生活。第1部の『連続性』が、ここで具体的な目的になっています。");
}

/* ============ SLIDE 17 : 規模データ ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "練馬区の規模 ― 数字で見る", 17);
  // big pop stat
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.7, w: 5.75, h: 2.0, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() });
  s.addText("区の人口", { x: ML + 0.35, y: 1.85, w: 5, h: 0.4, fontFace: JP, fontSize: 13, color: "CFE6DB", margin: 0 });
  s.addText([{ text: "751,465", options: { fontSize: 46, bold: true, color: WHITE, fontFace: NUM } }, { text: " 人", options: { fontSize: 20, bold: true, color: WHITE } }], { x: ML + 0.3, y: 2.35, w: 5.2, h: 1.0, valign: "middle", fontFace: JP, align: "left", margin: 0 });
  s.addText("令和8年4月1日現在", { x: ML + 0.35, y: 3.28, w: 5, h: 0.35, fontFace: JP, fontSize: 11, color: "CFE6DB", margin: 0 });
  // schools
  s.addShape(p.ShapeType.roundRect, { x: 6.98, y: 1.7, w: 5.75, h: 2.0, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText("区立学校数", { x: 7.28, y: 1.85, w: 5, h: 0.4, fontFace: JP, fontSize: 13, color: MUTED, margin: 0 });
  s.addText([
    { text: "小学校 ", options: { fontSize: 15, color: INK } }, { text: "63", options: { fontSize: 30, bold: true, color: PRIMARY, fontFace: NUM } }, { text: "校　　", options: { fontSize: 15, color: INK } },
    { text: "中学校 ", options: { fontSize: 15, color: INK } }, { text: "31", options: { fontSize: 30, bold: true, color: PRIMARY, fontFace: NUM } }, { text: "校\n", options: { fontSize: 15, color: INK } },
    { text: "小中一貫教育校 ", options: { fontSize: 15, color: INK } }, { text: "2", options: { fontSize: 30, bold: true, color: GOLDDK, fontFace: NUM } }, { text: "校", options: { fontSize: 15, color: INK } }],
    { x: 7.28, y: 2.3, w: 5.2, h: 1.3, valign: "middle", fontFace: JP, align: "left", margin: 0, lineSpacingMultiple: 1.2 });
  // students
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 3.9, w: 12.13, h: 1.95, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText("児童・生徒数（令和7年5月1日現在）", { x: ML + 0.35, y: 4.05, w: 8, h: 0.4, fontFace: JP, fontSize: 13, color: GOLDDK, margin: 0 });
  s.addText([{ text: "小学生 ", options: { fontSize: 16, color: INK } }, { text: "33,426", options: { fontSize: 34, bold: true, color: PRIMARY, fontFace: NUM } }, { text: " 人", options: { fontSize: 16, color: INK } }], { x: ML + 0.35, y: 4.55, w: 6, h: 1.1, valign: "middle", fontFace: JP, align: "left", margin: 0 });
  s.addText([{ text: "中学生 ", options: { fontSize: 16, color: INK } }, { text: "13,226", options: { fontSize: 34, bold: true, color: PRIMARY, fontFace: NUM } }, { text: " 人", options: { fontSize: 16, color: INK } }], { x: 6.8, y: 4.55, w: 6, h: 1.1, valign: "middle", fontFace: JP, align: "left", margin: 0 });
  s.addText("小・中学生を合わせて約4万7千人の学びを、区全体で9年間つないでいます。", { x: ML, y: 6.05, w: CW, h: 0.4, fontFace: JP, fontSize: 12.5, color: MUTED });
  note(s, "練馬区の規模を数字で押さえます。人口は約75万人。23区で最も人口の多い区の一つです。区立の小学校が63校、中学校が31校、そして施設一体型の小中一貫教育校が2校。児童・生徒は小学生3万3千余、中学生1万3千余。合わせて約4万7千人の子供たちの学びを、区全体で9年間つないでいます。");
}

/* ============ SLIDE 18 : 推進の仕組み ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "推進の仕組み", 18);
  // 33 groups
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.6, w: 5.9, h: 1.75, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() });
  s.addText([{ text: "33", options: { fontSize: 40, bold: true, color: WHITE, fontFace: NUM } }, { text: " の中学校区グループ", options: { fontSize: 16, bold: true, color: WHITE } }], { x: ML + 0.35, y: 1.75, w: 5.3, h: 0.8, valign: "middle", fontFace: JP, align: "left", margin: 0 });
  s.addText("中学校を基盤に、区内の全小・中学校が連携。", { x: ML + 0.35, y: 2.6, w: 5.3, h: 0.6, fontFace: JP, fontSize: 12.5, color: "CFE6DB", margin: 0 });
  // 知的障害
  s.addShape(p.ShapeType.roundRect, { x: 6.83, y: 1.6, w: 5.9, h: 1.75, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText("知的障害学級 小中グループ", { x: 7.13, y: 1.78, w: 5.3, h: 0.5, fontFace: JP, fontSize: 16, bold: true, color: INK, margin: 0 });
  s.addText("4ブロックで情報交換・研究協議を行い、特別支援の視点でも9年間をつなぐ。", { x: 7.13, y: 2.3, w: 5.3, h: 0.9, fontFace: JP, fontSize: 12.5, color: INK, margin: 0, lineSpacingMultiple: 1.1 });
  // クリエーター
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 3.55, w: 12.13, h: 2.55, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText([{ text: "小中一貫教育クリエーターの選任　", options: { fontSize: 16, bold: true, color: GOLDDK } }, { text: "＝ 中学校区における連絡・調整を担う「推進の核」", options: { fontSize: 13.5, color: INK } }], { x: ML + 0.35, y: 3.72, w: 11.5, h: 0.5, fontFace: JP, margin: 0 });
  const roles = ["連続性・系統性のある教育課程", "児童・生徒の計画的・継続的な交流", "教員の計画的・継続的な交流", "小中一貫教育を進める学校運営"];
  const rw = 2.78, rh = 1.35, ry = 4.4, rx0 = ML + 0.3, rgap = 0.15;
  roles.forEach((r, i) => {
    const x = rx0 + i * (rw + rgap);
    s.addShape(p.ShapeType.roundRect, { x, y: ry, w: rw, h: rh, fill: { color: WHITE }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.07 });
    s.addText(String(i + 1), { x: x + 0.18, y: y_role(ry), w: 0.4, h: 0.4, align: "center", valign: "middle", fontFace: JP, fontSize: 13, bold: true, color: WHITE, fill: { color: GOLD }, shape: p.ShapeType.ellipse, line: { type: "none" }, margin: 0 });
    s.addText(r, { x: x + 0.15, y: ry + 0.6, w: rw - 0.3, h: rh - 0.7, align: "center", valign: "top", fontFace: JP, fontSize: 11.5, bold: true, color: INK, margin: 0, lineSpacingMultiple: 1.05 });
  });
  function y_role(v) { return v + 0.15; }
  note(s, "では区全体でどう進めるか。仕組みは3層です。まず、中学校を基盤に全小・中学校が連携する33の中学校区グループ。次に、知的障害学級の小中グループが4ブロックあり、特別支援の視点でも9年間をつなぎます。そして各中学校区に小中一貫教育クリエーターを選任。連続性・系統性のある教育課程、児童生徒の交流、教員の交流、学校運営――この4つを担う、推進の核となる存在です。");
}

/* ============ SLIDE 19 : 施設一体型 2校 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "施設一体型 小中一貫教育校", 19);
  const pw = 5.9, px2 = 6.83, py = 1.6, ph = 2.35;
  s.addImage({ path: __dirname + "/../assets/oizumisakura.png", x: ML, y: py, w: pw, h: ph, sizing: { type: "cover", w: pw, h: ph }, rounding: false, shadow: shadow() });
  s.addImage({ path: __dirname + "/../assets/miraiaozora.png", x: px2, y: py, w: pw, h: ph, sizing: { type: "cover", w: pw, h: ph }, shadow: shadow() });
  // captions
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 4.05, w: pw, h: 2.0, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText("大泉桜学園", { x: ML + 0.3, y: 4.18, w: pw - 0.6, h: 0.45, fontFace: JP, fontSize: 17, bold: true, color: PRIMARY, margin: 0 });
  s.addText("9年間を見通した学習・生活指導／小学校高学年からの一部教科担任制／合同学校行事／キャリア教育での異学年交流 など", { x: ML + 0.3, y: 4.62, w: pw - 0.6, h: 1.3, fontFace: JP, fontSize: 12.5, color: INK, margin: 0, lineSpacingMultiple: 1.15 });
  s.addShape(p.ShapeType.roundRect, { x: px2, y: 4.05, w: pw, h: 2.0, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText([{ text: "みらい青空学園　", options: { fontSize: 17, bold: true, color: GOLDDK } }, { text: "（旭丘小・旭丘中／令和8年4月開校）", options: { fontSize: 11.5, color: MUTED } }], { x: px2 + 0.3, y: 4.18, w: pw - 0.6, h: 0.45, fontFace: JP, margin: 0 });
  s.addText("施設一体型の強みを生かした日常的な児童・生徒交流／通常の学級と特別支援学級の交流学習／地域・近隣大学等との連携／9年間の段階的な学習指導と体育・健康教育 など", { x: px2 + 0.3, y: 4.62, w: pw - 0.6, h: 1.3, fontFace: JP, fontSize: 12.5, color: INK, margin: 0, lineSpacingMultiple: 1.15 });
  note(s, "その象徴が、施設一体型の小中一貫教育校です。左が大泉桜学園。9年間を見通した指導、小学校高学年からの一部教科担任制、合同行事、異学年交流などを行っています。右が、今年度・令和8年4月に開校したばかりの、みらい青空学園。旭丘小・旭丘中が母体です。日常的な児童・生徒交流、通常学級と特別支援学級の交流学習、地域や近隣大学との連携などが特色です。");
}

/* ============ SLIDE 20 : 令和8年度の重点取組 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "令和8年度の重点取組", 20);
  // R7 -> R8 flow
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 1.7, w: 5.55, h: 1.5, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.08, shadow: shadow() });
  s.addText([{ text: "令和7年度\n", options: { fontSize: 12, bold: true, color: MUTED } }, { text: "取組プログラムの\n", options: { fontSize: 15, color: INK } }, { text: "充実", options: { fontSize: 20, bold: true, color: PRIMARY } }], { x: ML, y: 1.7, w: 5.55, h: 1.5, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.05 });
  s.addText("→", { x: 6.15, y: 1.7, w: 1.0, h: 1.5, align: "center", valign: "middle", fontFace: JP, fontSize: 34, bold: true, color: GOLD, margin: 0 });
  s.addShape(p.ShapeType.roundRect, { x: 7.18, y: 1.7, w: 5.55, h: 1.5, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() });
  s.addText([{ text: "令和8年度\n", options: { fontSize: 12, bold: true, color: GOLD } }, { text: "取組プログラムの\n", options: { fontSize: 15, color: "CFE6DB" } }, { text: "見直し", options: { fontSize: 20, bold: true, color: WHITE } }], { x: 7.18, y: 1.7, w: 5.55, h: 1.5, align: "center", valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.05 });
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 3.5, w: CW, h: 1.25, fill: { color: GOLDFILL }, line: { color: GOLDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
  s.addText([{ text: "重点　", options: { bold: true, color: GOLDDK, fontSize: 15 } }, { text: "「目指す15歳の姿」の実現に向けた「小中一貫教育の取組プログラム」の見直し", options: { bold: true, color: INK, fontSize: 16 } }], { x: ML + 0.35, y: 3.5, w: CW - 0.7, h: 1.25, valign: "middle", fontFace: JP, margin: 0 });
  s.addText("全国学力・学習状況調査や新体力テスト等から見えた児童生徒の実態（論理的に考え説明する力の課題、体力の横ばい・減少傾向）を踏まえ、各グループが自校の実態に応じてプログラムを見直します。", { x: ML, y: 4.95, w: CW, h: 1.1, fontFace: JP, fontSize: 13, color: INK, valign: "top", lineSpacingMultiple: 1.2 });
  note(s, "今年度の重点取組です。これまで各グループは『目指す15歳の姿』に向けて取組プログラムを作り、昨年度はその充実を図ってきました。今年度・令和8年度は、それを『見直す』フェーズです。全国学力調査や新体力テストから、論理的に考え説明する力の課題、体力の横ばい・減少といった実態が見えています。それを踏まえ、各グループが自校の子供の実態に合わせてプログラムを磨き直します。");
}

/* ============ SLIDE 21 : 多様な取組 ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE }; head(s, 2, "多様な取組で「滑らかな接続」へ", 21);
  const tags = ["乗り入れ授業", "部活動体験", "合同行事", "児童会・生徒会交流", "あいさつ運動", "合同校内研究", "課題改善カリキュラム", "小学校での一部教科担任制"];
  const cols = 4, tw = 2.85, th = 1.15, gx = 0.24, gy = 0.28, ox = ML, oy = 1.75;
  const cyc = [PRIMARY, SAGE, PRIMARY2, GOLD];
  tags.forEach((t, i) => {
    const c = i % cols, r = Math.floor(i / cols);
    const x = ox + c * (tw + gx), y = oy + r * (th + gy);
    s.addShape(p.ShapeType.roundRect, { x, y, w: tw, h: th, fill: { color: LIGHTBG }, line: { color: CARDLINE, width: 1 }, rectRadius: 0.1, shadow: shadow() });
    s.addShape(p.ShapeType.ellipse, { x: x + 0.28, y: y + th / 2 - 0.11, w: 0.22, h: 0.22, fill: { color: cyc[i % 4] }, line: { type: "none" } });
    s.addText(t, { x: x + 0.6, y, w: tw - 0.75, h: th, valign: "middle", fontFace: JP, fontSize: 13.5, bold: true, color: INK, margin: 0, lineSpacingMultiple: 1.0 });
  });
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 4.9, w: CW, h: 1.15, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.08, shadow: shadow() });
  s.addText("多様な取組を継続することにより、小学校から中学校への「滑らかな接続」の実現を目指す。", { x: ML + 0.4, y: 4.9, w: CW - 0.8, h: 1.15, valign: "middle", fontFace: JP, fontSize: 16, bold: true, color: WHITE, margin: 0 });
  note(s, "プログラムの中身は、こうした多様な取組の組み合わせです。中学校の先生が小学校で教える乗り入れ授業、部活動体験、合同行事、児童会・生徒会の交流、あいさつ運動、合同での校内研究、そして課題改善カリキュラム。派手な一発ではなく、こうした取組を地道に継続することで、小学校から中学校への滑らかな接続を実現していきます。");
}

/* ============ SLIDE 22 : まとめ ============ */
{
  const s = p.addSlide(); s.background = { color: WHITE };
  s.addText("まとめ", { x: ML, y: 0.5, w: CW, h: 0.7, fontFace: JP, fontSize: 27, bold: true, color: INK });
  const items = [
    ["教育課程の管理", "PDCAを組織的に回し、9年間の系統性・連続性を確保する ― 指導主事の中核業務。"],
    ["練馬区の小中一貫教育", "その連続性を、中学校区・区全体で実現する具体の取組。定義・3つの目的・仕組み。"],
    ["つながる一本の軸", "「目指す15歳の姿」に向け、教育課程の視点で小学校と中学校をつなぐ。"]];
  let y = 1.4;
  items.forEach((it, i) => {
    s.addShape(p.ShapeType.roundRect, { x: ML, y, w: CW, h: 1.15, fill: { color: i === 2 ? GOLDFILL : LIGHTBG }, line: { color: i === 2 ? GOLDLINE : CARDLINE, width: 1 }, rectRadius: 0.07, shadow: shadow() });
    s.addText(String(i + 1), { x: ML + 0.28, y: y + 0.28, w: 0.6, h: 0.6, align: "center", valign: "middle", fontFace: JP, fontSize: 20, bold: true, color: WHITE, fill: { color: i === 2 ? GOLD : PRIMARY }, shape: p.ShapeType.ellipse, line: { type: "none" }, margin: 0 });
    s.addText([{ text: it[0] + "　", options: { fontSize: 16, bold: true, color: i === 2 ? GOLDDK : PRIMARY } }, { text: it[1], options: { fontSize: 13.5, color: INK } }], { x: ML + 1.1, y, w: CW - 1.4, h: 1.15, valign: "middle", fontFace: JP, margin: 0, lineSpacingMultiple: 1.1 });
    y += 1.32;
  });
  s.addShape(p.ShapeType.roundRect, { x: ML, y: 5.4, w: CW, h: 0.95, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.07 });
  s.addText([{ text: "配布資料　", options: { bold: true, color: GOLD, fontSize: 13 } }, { text: "①教育行政研修資料（教育課程の管理）　②令和8年度 小中一貫教育の取組について　― 詳細はお手元の資料をご参照ください。", options: { color: WHITE, fontSize: 13 } }], { x: ML + 0.35, y: 5.4, w: CW - 0.7, h: 0.95, valign: "middle", fontFace: JP, margin: 0 });
  note(s, "まとめます。1つ、教育課程の管理は、PDCAを組織的に回し、9年間の系統性・連続性を確保する、指導主事の中核業務。2つ、練馬区の小中一貫教育は、その連続性を中学校区・区全体で実現する具体の取組。3つ、両者は『目指す15歳の姿』に向けて、教育課程の視点で小・中をつなぐ一本の軸です。詳しくはお手元の2つの資料をご覧ください。");
}

/* ============ SLIDE 23 : 結び ============ */
{
  const s = p.addSlide(); s.background = { color: DARK }; dots(s);
  s.addText("ご清聴ありがとうございました", { x: ML, y: 2.5, w: 12.1, h: 1.2, align: "center", fontFace: JP, fontSize: 38, bold: true, color: WHITE });
  s.addShape(p.ShapeType.roundRect, { x: 4.66, y: 4.1, w: 4.0, h: 0.9, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.1, shadow: shadow() });
  s.addText("質疑応答（約5分）", { x: 4.66, y: 4.1, w: 4.0, h: 0.9, align: "center", valign: "middle", fontFace: JP, fontSize: 18, bold: true, color: WHITE, margin: 0 });
  s.addText("練馬区教育委員会 教育指導課　指導主事　紺多 章一郎", { x: ML, y: 5.6, w: 12.1, h: 0.4, align: "center", fontFace: JP, fontSize: 13, color: SAGE });
  note(s, "以上で説明を終わります。ご清聴ありがとうございました。ここから5分ほど、質疑応答の時間とします。教育課程の管理のこと、小中一貫の具体のこと、どんなことでも結構です。ご質問をどうぞ。");
}

p.writeFile({ fileName: __dirname + "/deck.pptx" }).then(f => console.log("WROTE", f));
