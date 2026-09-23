const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const C = {
  deep: "0B4F5C",   // 深いティール（主色）
  teal: "1B7F8C",
  mint: "DCEFEF",
  pale: "F3F8F8",
  sun: "F2A541",    // 朝日のアクセント
  coral: "D9583B",
  ink: "1E2B2F",
  mute: "5B6B70",
  white: "FFFFFF",
};
const F = "Meiryo";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// 円の中にアイコン（全スライド共通モチーフ）
async function iconCircle(slide, Comp, x, y, d, bg, fg) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const p = d * 0.26;
  slide.addImage({ data: await icon(Comp, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

function title(slide, text, sub) {
  slide.addText(text, {
    x: 0.6, y: 0.4, w: 12.1, h: 0.8, fontFace: F, fontSize: 30, bold: true,
    color: C.deep, margin: 0, isTextBox: true, valign: "middle",
  });
  if (sub) {
    slide.addText(sub, {
      x: 0.6, y: 1.2, w: 12.1, h: 0.45, fontFace: F, fontSize: 16,
      color: C.mute, margin: 0, isTextBox: true,
    });
  }
}

function pageNo(slide, n) {
  slide.addText(String(n), {
    x: 12.3, y: 6.95, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: C.mute,
    align: "right", margin: 0, isTextBox: true,
  });
}

// 根拠（文部科学省資料）の脚注
function source(slide, text, dark) {
  slide.addShape("roundRect", { x: 0.6, y: 6.86, w: 0.62, h: 0.3, rectRadius: 0.05, fill: { color: dark ? C.sun : C.deep }, line: { color: dark ? C.sun : C.deep } });
  slide.addText("根拠", { x: 0.6, y: 6.86, w: 0.62, h: 0.3, fontFace: F, fontSize: 10, bold: true, color: dark ? C.ink : C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
  slide.addText(text, { x: 1.32, y: 6.8, w: 10.9, h: 0.42, fontFace: F, fontSize: 9.5, color: dark ? C.mint : C.mute, valign: "middle", margin: 0, isTextBox: true });
}

function card(slide, x, y, w, h, fill) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.white }, line: { color: fill || C.white },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.12 },
  });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.title = "10年後の教育界を拓く視点と学校経営";

  // ---------- 1 タイトル ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.deep };
    s.addShape("ellipse", { x: 8.9, y: -1.6, w: 6.2, h: 6.2, fill: { color: C.teal, transparency: 55 }, line: { color: C.teal, transparency: 55 } });
    s.addShape("ellipse", { x: 10.4, y: 3.9, w: 4.2, h: 4.2, fill: { color: C.sun, transparency: 70 }, line: { color: C.sun, transparency: 70 } });
    await iconCircle(s, fa.FaCompass, 0.8, 0.9, 1.1, C.sun, C.deep);
    s.addText("指導講評・研修プログラム（45分）", {
      x: 0.8, y: 2.25, w: 9, h: 0.45, fontFace: F, fontSize: 16, color: C.sun, bold: true, margin: 0, isTextBox: true,
    });
    s.addText("10年後の教育界を拓く視点と学校経営", {
      x: 0.8, y: 2.8, w: 11, h: 1.2, fontFace: F, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true, valign: "middle",
    });
    s.addText("〜次期学習指導要領素案が描く「自律的学び」「マイ探究」「キャリア教育」への挑戦〜", {
      x: 0.8, y: 4.1, w: 11, h: 0.6, fontFace: F, fontSize: 18, color: C.mint, margin: 0, isTextBox: true,
    });
    s.addText("指導主事 講評", {
      x: 0.8, y: 6.3, w: 6, h: 0.4, fontFace: F, fontSize: 14, color: C.mint, margin: 0, isTextBox: true,
    });
    s.addNotes("校長先生、副校長先生、そして現場で子どもたちと向き合う先生方、本日もお疲れ様です。本日は講評を一方的にお聞きいただくのではなく、管理職と教諭が対等に語り合う45分間として設計しました。");
  }

  // ---------- 2 プログラム ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.pale };
    title(s, "本日のプログラム（対話型講評の設計）", "学校経営と授業改善を一体化させる45分間");
    const blocks = [
      { t: "00-05分", m: 5, h: "導入・アイスブレイク", d: "10年後の学校で子どもが歓喜する瞬間をペアで共有", ic: fa.FaSmile, c: C.sun },
      { t: "05-25分", m: 20, h: "指導主事講評・提案", d: "次期学習指導要領素案が描く10年後の未来\n自律的学び／マイ探究／キャリア教育／外国語教育", ic: fa.FaChalkboardTeacher, c: C.deep },
      { t: "25-40分", m: 15, h: "グループ協議・熟議", d: "校長・副校長・教諭のミックスグループで\n「我が校の明日への小さな一歩（出島づくり）」", ic: fa.FaUsers, c: C.teal },
      { t: "40-45分", m: 5, h: "まとめ・メッセージ", d: "教育課程は制限ではなく、未来を拓く出島", ic: fa.FaFlag, c: C.coral },
    ];
    // 時間配分バー（45分比例）
    const X0 = 0.6, W = 12.1, gap = 0.08;
    let x = X0;
    const unit = (W - gap * 3) / 45;
    blocks.forEach((b) => {
      const w = b.m * unit;
      s.addShape("rect", { x, y: 2.0, w, h: 0.55, fill: { color: b.c }, line: { color: b.c } });
      s.addText(b.m + "分", { x, y: 2.0, w, h: 0.55, fontFace: F, fontSize: 14, bold: true, color: b.c === C.sun ? C.ink : C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      x += w + gap;
    });
    // カード
    const cw = (W - 0.3 * 3) / 4;
    for (let i = 0; i < 4; i++) {
      const b = blocks[i];
      const cx = X0 + i * (cw + 0.3);
      card(s, cx, 3.0, cw, 3.4);
      await iconCircle(s, b.ic, cx + 0.3, 3.3, 0.8, b.c, b.c === C.sun ? C.ink : C.white);
      s.addText(b.t, { x: cx + 1.25, y: 3.45, w: cw - 1.4, h: 0.5, fontFace: F, fontSize: 14, color: C.mute, bold: true, margin: 0, isTextBox: true, valign: "middle" });
      s.addText(b.h, { x: cx + 0.3, y: 4.3, w: cw - 0.6, h: 0.5, fontFace: F, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(b.d, { x: cx + 0.3, y: 4.9, w: cw - 0.6, h: 1.6, fontFace: F, fontSize: 13, color: C.ink, margin: 0, isTextBox: true, valign: "top", lineSpacingMultiple: 1.2 });
    }
    pageNo(s, pres.slides.length);
    s.addNotes("講評を聞くだけの受動的な時間ではありません。前半で視点を共有し、後半は管理職も教諭も同じ輪に入って自校の実践を語り合います。学校経営と授業改善を一体化させる45分間です。");
  }

  // ---------- 3 アイスブレイク ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    title(s, "【導入】10年後の学校を想像する", "5分ワーク ／ ペアで1分半ずつ共有");
    card(s, 0.6, 1.95, 7.9, 4.8, C.mint);
    await iconCircle(s, fa.FaQuestion, 0.95, 2.3, 0.9, C.deep, C.white);
    s.addText("問い", { x: 2.05, y: 2.3, w: 3, h: 0.9, fontFace: F, fontSize: 20, bold: true, color: C.deep, valign: "middle", margin: 0, isTextBox: true });
    s.addText([
      { text: "生成AIや自動翻訳が日常となった", options: { breakLine: true } },
      { text: "10年後（2035〜2040年）、", options: { breakLine: true } },
      { text: "「それでも学校に来て仲間や先生と学んでよかった！」", options: { bold: true, color: C.coral, breakLine: true } },
      { text: "と子どもが歓喜する瞬間とは？" },
    ], { x: 0.95, y: 3.45, w: 7.2, h: 2.9, fontFace: F, fontSize: 19, color: C.ink, margin: 0, isTextBox: true, valign: "top", lineSpacingMultiple: 1.4 });

    // 進め方
    const steps = [
      ["1分", "一人で思い浮かべる"],
      ["1分半", "Aさんが語る"],
      ["1分半", "Bさんが語る"],
      ["1分", "共通点を見つける"],
    ];
    s.addText("進め方", { x: 9.0, y: 1.95, w: 3.7, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: C.deep, margin: 0, isTextBox: true });
    steps.forEach((st, i) => {
      const y = 2.6 + i * 0.85;
      s.addShape("roundRect", { x: 9.0, y, w: 1.25, h: 0.62, rectRadius: 0.1, fill: { color: i % 3 === 0 ? C.sun : C.teal }, line: { color: i % 3 === 0 ? C.sun : C.teal } });
      s.addText(st[0], { x: 9.0, y, w: 1.25, h: 0.62, fontFace: F, fontSize: 14, bold: true, color: i % 3 === 0 ? C.ink : C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(st[1], { x: 10.4, y, w: 2.4, h: 0.62, fontFace: F, fontSize: 15, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("ねらい：立場を超えて「教育のワクワク」を共有し、場をやわらかくほぐす", {
      x: 9.0, y: 6.1, w: 3.8, h: 0.65, fontFace: F, fontSize: 12, color: C.mute, margin: 0, isTextBox: true, italic: true,
    });
    pageNo(s, pres.slides.length);
    s.addNotes("まずはペアで、10年後の子どもの姿を思い描いてみましょう。AIが答えを出してくれる時代に、それでも学校に来てよかったと子どもが歓喜する瞬間とはどんな場面でしょうか。お一人1分半ずつ語ってください。管理職も教諭も、立場を超えてワクワクを共有する時間です。");
  }

  // ---------- 検討経過（文部科学省・中央教育審議会） ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    title(s, "次期学習指導要領の検討経過", "中央教育審議会における審議の流れ（文部科学省公表資料より）");
    const steps = [
      { d: "令和6年12月25日", h: "文部科学大臣 諮問", t: "「初等中等教育における教育課程の基準等の在り方について」", b: ["質の高い、深い学びの実現と分かりやすい学習指導要領", "多様な子供たちを包摂する柔軟な教育課程", "これからの時代に求められる資質・能力を踏まえた教科等の在り方", "教育課程の実施に伴う負担への対応"], ic: fa.FaScroll, c: C.teal },
      { d: "令和7年9月25日", h: "教育課程企画特別部会 論点整理", t: "目指す子供像と改訂の方向性を整理", b: ["「自らの人生を舵取りすることができる」子供の育成", "子供が学びを主体的に調整することを促す", "構造化・精選による「余白」の創出", "調整授業時数制度・裁量的な時間の検討"], ic: fa.FaClipboardList, c: C.deep },
      { d: "令和8年8月31日", h: "審議まとめ（素案）", t: "「次期学習指導要領等に向けた審議まとめ（素案）」（第17回特別部会 資料1）", b: ["各ワーキンググループの取りまとめを反映", "「主体的・対話的で深い学び」の一層の実現", "学びと実生活・実社会とのつながり"], ic: fa.FaFileAlt, c: C.coral },
    ];
    const w = 3.75, g = 0.425;
    for (let i = 0; i < steps.length; i++) {
      const st = steps[i], x = 0.6 + i * (w + g);
      card(s, x, 1.95, w, 4.65, C.pale);
      await iconCircle(s, st.ic, x + 0.25, 2.15, 0.75, st.c, C.white);
      s.addText(st.d, { x: x + 1.15, y: 2.15, w: w - 1.3, h: 0.35, fontFace: F, fontSize: 12, bold: true, color: st.c, margin: 0, isTextBox: true, valign: "middle" });
      s.addText(st.h, { x: x + 1.15, y: 2.5, w: w - 1.3, h: 0.45, fontFace: F, fontSize: 14, bold: true, color: C.ink, margin: 0, isTextBox: true, valign: "middle" });
      s.addText(st.t, { x: x + 0.25, y: 3.1, w: w - 0.5, h: 0.75, fontFace: F, fontSize: 12.5, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
      s.addText(st.b.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < st.b.length - 1 } })),
        { x: x + 0.25, y: 3.95, w: w - 0.5, h: 2.5, fontFace: F, fontSize: 13, color: C.ink, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 10 });
      if (i < steps.length - 1) {
        s.addShape("rightArrow", { x: x + w + 0.07, y: 2.35, w: 0.3, h: 0.36, fill: { color: C.sun }, line: { color: C.sun } });
      }
    }
    source(s, "文部科学省「初等中等教育における教育課程の基準等の在り方について（諮問）」（令和6年12月25日）／中央教育審議会 教育課程企画特別部会「論点整理」（令和7年9月25日）／同「次期学習指導要領等に向けた審議まとめ（素案）」（令和8年8月31日）※箇条は要旨");
    pageNo(s, pres.slides.length);
    s.addNotes("本日の講評の根拠となる文部科学省・中央教育審議会の資料を確認しておきます。令和6年12月25日に文部科学大臣から中央教育審議会へ諮問が行われ、令和7年9月25日に教育課程企画特別部会の論点整理、そして令和8年8月31日に審議まとめの素案が示されました。本日お話しする「舵取り」「主体的な学びの調整」「余白」といった考え方は、いずれもこの審議の流れの中で示されているものです。（各項目は資料の要旨です。原文は文部科学省ウェブサイトでご確認ください。）");
  }

  // ---------- 4 エージェンシー ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.pale };
    title(s, "10年後の教育界が目指す姿", "自らの人生を「舵取り」する子の育成");
    // 左：中核概念
    s.addShape("ellipse", { x: 0.9, y: 2.0, w: 4.5, h: 4.5, fill: { color: C.deep }, line: { color: C.deep } });
    s.addImage({ data: await icon(fa.FaCompass, C.sun), x: 2.6, y: 2.45, w: 1.1, h: 1.1 });
    s.addText("エージェンシー", { x: 0.9, y: 3.7, w: 4.5, h: 0.7, fontFace: F, fontSize: 26, bold: true, color: C.white, align: "center", margin: 0, isTextBox: true });
    s.addText("（舵取りする力）", { x: 0.9, y: 4.4, w: 4.5, h: 0.5, fontFace: F, fontSize: 16, color: C.mint, align: "center", margin: 0, isTextBox: true });
    s.addText("→ 人間の強みの再定義", { x: 0.9, y: 5.05, w: 4.5, h: 0.45, fontFace: F, fontSize: 13, color: C.sun, align: "center", margin: 0, isTextBox: true });

    const rows = [
      [fa.FaLightbulb, "自らの興味・問いを起点に", "学びと人生を主体的に進める"],
      [fa.FaHandsHelping, "身体性を伴う体験", "受動的な知識受容を超え、五感で学ぶ"],
      [fa.FaUsers, "他者との共創", "多様な仲間と対話し、新たな価値を創る"],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 1.95 + i * 1.05;
      card(s, 6.0, y, 6.7, 0.9);
      await iconCircle(s, rows[i][0], 6.2, y + 0.13, 0.64, C.teal, C.white);
      s.addText(rows[i][1], { x: 7.05, y: y + 0.08, w: 5.5, h: 0.42, fontFace: F, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(rows[i][2], { x: 7.05, y: y + 0.48, w: 5.5, h: 0.36, fontFace: F, fontSize: 13, color: C.mute, margin: 0, isTextBox: true });
    }
    s.addShape("roundRect", { x: 6.0, y: 5.1, w: 6.7, h: 1.5, rectRadius: 0.1, fill: { color: C.sun }, line: { color: C.sun } });
    s.addText([
      { text: "論点整理が示す目指す子供像", options: { fontSize: 12, bold: true, color: C.deep, breakLine: true } },
      { text: "「生涯にわたって主体的に学び続け、多様な他者と協働しながら、自らの人生を舵取りすることができる、民主的で持続可能な社会の創り手」", options: { fontSize: 14, bold: true, color: C.ink } },
    ], { x: 6.25, y: 5.15, w: 6.25, h: 1.4, fontFace: F, valign: "middle", margin: 0, isTextBox: true, paraSpaceAfter: 4 });
    source(s, "中央教育審議会 教育課程企画特別部会「論点整理」（令和7年9月25日）");
    pageNo(s, pres.slides.length);
    s.addNotes("次期学習指導要領の素案を紐解くと、そこに流れているのは『教員の負担を増やす細かなルール』ではなく、『10年後の不確実な社会を生きる子どもたちが、自らの人生を誇りを持って舵取りしていくための道しるべ』です。AIが答えを即座に出してくれる時代だからこそ、学校は『答えを覚える場所』から、『自ら問いを立て、自分のペースで学びを深め、多様な他者と対話しながら新たな価値を創り出す場所』へとアップデートする必要があります。【根拠】令和7年9月25日の教育課程企画特別部会「論点整理」では、目指す子供像として「生涯にわたって主体的に学び続け、多様な他者と協働しながら、自らの人生を舵取りすることができる、民主的で持続可能な社会の創り手」が掲げられています。");
  }

  // ---------- 5 自由進度学習 ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    title(s, "【自由進度学習の在り方】", "「放任」から「指導性と自律の往還」へ");
    // NG
    card(s, 0.6, 2.0, 4.3, 3.7, "FBEDE9");
    await iconCircle(s, fa.FaTimes, 0.9, 2.3, 0.75, C.coral, C.white);
    s.addText("陥ってはならない姿", { x: 1.85, y: 2.3, w: 2.9, h: 0.75, fontFace: F, fontSize: 17, bold: true, color: C.coral, valign: "middle", margin: 0, isTextBox: true });
    s.addText([
      { text: "プリントやドリル任せの「放任」", options: { bullet: true, breakLine: true } },
      { text: "できない子の「自己責任化」", options: { bullet: true, breakLine: true } },
      { text: "ゴールが見えないまま進む学習", options: { bullet: true } },
    ], { x: 0.9, y: 3.3, w: 3.8, h: 1.45, fontFace: F, fontSize: 14, color: C.ink, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 8 });
    s.addText([
      { text: "令和3年答申", options: { fontSize: 11, bold: true, color: C.coral, breakLine: true } },
      { text: "「個別最適な学び」が「孤立した学び」に陥らないよう…「協働的な学び」を充実", options: { fontSize: 12, color: C.ink } },
    ], { x: 0.9, y: 4.8, w: 3.8, h: 0.8, fontFace: F, margin: 0, isTextBox: true, valign: "top" });

    // 往還サイクル
    s.addText("本質は「往還」", { x: 5.4, y: 2.0, w: 7.3, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: C.deep, margin: 0, isTextBox: true });
    const L = { x: 5.4, y: 2.8, w: 3.1, h: 2.9 };
    const R = { x: 9.6, y: 2.8, w: 3.1, h: 2.9 };
    card(s, L.x, L.y, L.w, L.h, C.mint);
    card(s, R.x, R.y, R.w, R.h, C.mint);
    await iconCircle(s, fa.FaChild, L.x + 1.1, L.y + 0.25, 0.9, C.teal, C.white);
    await iconCircle(s, fa.FaChalkboardTeacher, R.x + 1.1, R.y + 0.25, 0.9, C.deep, C.white);
    s.addText("子ども", { x: L.x, y: L.y + 1.25, w: L.w, h: 0.4, fontFace: F, fontSize: 13, color: C.mute, align: "center", margin: 0, isTextBox: true });
    s.addText("自己調整学習\n（メタ認知）", { x: L.x + 0.2, y: L.y + 1.65, w: L.w - 0.4, h: 1.0, fontFace: F, fontSize: 17, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
    s.addText("教師", { x: R.x, y: R.y + 1.25, w: R.w, h: 0.4, fontFace: F, fontSize: 13, color: C.mute, align: "center", margin: 0, isTextBox: true });
    s.addText("適切な指導性の発揮\n（見守り・適時介入）", { x: R.x + 0.2, y: R.y + 1.65, w: R.w - 0.4, h: 1.0, fontFace: F, fontSize: 17, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
    // 双方向矢印
    s.addShape("rightArrow", { x: 8.65, y: 3.5, w: 0.8, h: 0.5, fill: { color: C.sun }, line: { color: C.sun } });
    s.addShape("leftArrow", { x: 8.65, y: 4.35, w: 0.8, h: 0.5, fill: { color: C.sun }, line: { color: C.sun } });
    s.addText("論点整理：「子供自らが自己の学習を主体的に調整することを促す」", {
      x: 5.4, y: 6.0, w: 7.3, h: 0.6, fontFace: F, fontSize: 14, color: C.deep, bold: true, margin: 0, isTextBox: true, valign: "middle",
    });
    source(s, "中央教育審議会答申「『令和の日本型学校教育』の構築を目指して」（令和3年1月26日）／教育課程企画特別部会「論点整理」（令和7年9月25日）");
    pageNo(s, pres.slides.length);
    s.addNotes("自由進度学習は、決して先生方の指導を放棄することではありません。単元のゴールを見通させ、子ども自身が学びを調整する自己調整学習と、それを見守り適時適切に介入する教師の指導性。この往還こそが本質です。先生方の温かい見取りと適切な投げかけがあってこそ、子どもは安心して試行錯誤できます。【根拠】令和3年1月の中教審答申「『令和の日本型学校教育』の構築を目指して」は、「個別最適な学び」が「孤立した学び」に陥らないよう、「協働的な学び」を充実することの重要性を示しています。また令和7年9月の「論点整理」は、子供自らが自己の学習を主体的に調整することを促すことにより、資質・能力の育成に資するとともに、一人一人の多様性に応じていくという視点を示しています。");
  }

  // ---------- 6 マイ探究 ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.pale };
    title(s, "【探究】「テーマ探究」から「マイ探究」への進化", "個人の「好き・得意」を起点とする探究へ");
    // 上段：進化
    card(s, 0.6, 2.0, 5.2, 1.3);
    s.addText([{ text: "テーマ探究", options: { bold: true, fontSize: 20, breakLine: true } }, { text: "教師が設定したテーマに沿って探究", options: { fontSize: 13, color: C.mute } }],
      { x: 0.9, y: 2.0, w: 4.8, h: 1.3, fontFace: F, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    s.addShape("rightArrow", { x: 6.05, y: 2.4, w: 1.2, h: 0.5, fill: { color: C.sun }, line: { color: C.sun } });
    card(s, 7.5, 2.0, 5.2, 1.3, C.deep);
    s.addText([{ text: "マイ探究", options: { bold: true, fontSize: 20, breakLine: true, color: C.white } }, { text: "自分の「好き・得意」から問いを立てる", options: { fontSize: 13, color: C.mint } }],
      { x: 7.8, y: 2.0, w: 4.8, h: 1.3, fontFace: F, valign: "middle", margin: 0, isTextBox: true });

    // 下段：課題の質の洗練プロセス
    s.addText("評価の視点：課題の質の洗練と自己変容", { x: 0.6, y: 3.65, w: 12.1, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: C.deep, margin: 0, isTextBox: true });
    const st = [
      [fa.FaSeedling, "素朴な問い", "「なぜ？」「好き」から出発"],
      [fa.FaSyncAlt, "試行錯誤", "調べ・つくり・語り合う"],
      [fa.FaSearchPlus, "解像度の高い課題", "問いが深まり、焦点化する"],
      [fa.FaStar, "自己変容", "自己や他者にとっての新たな意味・理解の構築"],
    ];
    const w = 2.8, g = 0.3;
    for (let i = 0; i < st.length; i++) {
      const x = 0.6 + i * (w + g);
      const last = i === st.length - 1;
      card(s, x, 4.35, w, 2.2, last ? C.sun : C.white);
      await iconCircle(s, st[i][0], x + 0.25, 4.6, 0.7, last ? C.deep : C.teal, C.white);
      s.addText(String(i + 1), { x: x + w - 0.75, y: 4.6, w: 0.5, h: 0.7, fontFace: F, fontSize: 24, bold: true, color: last ? C.deep : C.mint, align: "right", valign: "middle", margin: 0, isTextBox: true });
      s.addText(st[i][1], { x: x + 0.25, y: 5.45, w: w - 0.5, h: 0.45, fontFace: F, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(st[i][2], { x: x + 0.25, y: 5.9, w: w - 0.5, h: 0.7, fontFace: F, fontSize: 12, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
    }
    source(s, "教育課程部会 生活、総合的な学習・探究の時間WG「総合的な学習・探究の時間に関する現状・課題と検討事項」（令和7年10月15日）／同WG取りまとめ関連資料（令和8年7月）　※「マイ探究」は本講評での呼称");
    pageNo(s, pres.slides.length);
    s.addNotes("教師主導のテーマ探究から、個人の好き・得意を起点とするマイ探究への発展を評価したいと思います。試行錯誤の過程で素朴な問いが解像度の高い課題へと深まる変化を捉え、自己や他者にとっての新たな意味や理解の構築、すなわち自己変容そのものを成果として価値付けてください。【根拠】総合的な学習・探究の時間は、「課題の設定」「情報の収集」「整理・分析」「まとめ・表現」の探究の過程を通して、よりよく課題を解決し、自己の（在り方）生き方を考えていくための資質・能力の育成を目指すものです。生活、総合的な学習・探究の時間ワーキンググループでは、AI時代だからこそ子供の好奇心や思考を大切にし、対話・協働や自己調整を通じた質の高い探究を実現する方向で議論が進められています。なお「マイ探究」は本講評での呼称であり、文部科学省の用語ではありません。");
  }

  // ---------- 7 キャリア教育 ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    title(s, "【キャリア教育】「好き」を育み「得意」を伸ばす", "進路指導に閉じない、教育課程全体での構造化");
    const cols = [
      [fa.FaBook, "各教科", "好き・得意との出会い", "教科の学びの中で、自分の興味や強みに気づくきっかけをつくる", C.teal],
      [fa.FaCompass, "総合的な学習（探究）の時間", "マイ探究での育成", "好き・得意を起点に問いを深め、自分らしさを育てる", C.deep],
      [fa.FaPassport, "特別活動", "実践・見通しと振り返り", "キャリア・パスポート等で学びをつなぎ、将来を見通す", C.coral],
    ];
    const w = 3.8, g = 0.35;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (w + g);
      card(s, x, 2.0, w, 3.5, C.pale);
      await iconCircle(s, cols[i][0], x + (w - 1.0) / 2, 2.25, 1.0, cols[i][4], C.white);
      s.addText(cols[i][1], { x: x + 0.2, y: 3.4, w: w - 0.4, h: 0.5, fontFace: F, fontSize: 17, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
      s.addText(cols[i][2], { x: x + 0.2, y: 3.9, w: w - 0.4, h: 0.45, fontFace: F, fontSize: 15, bold: true, color: cols[i][4], align: "center", margin: 0, isTextBox: true });
      s.addText(cols[i][3], { x: x + 0.3, y: 4.45, w: w - 0.6, h: 1.0, fontFace: F, fontSize: 13, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
    }
    // 連携を示すベース
    s.addShape("roundRect", { x: 0.6, y: 5.75, w: 12.1, h: 0.9, rectRadius: 0.1, fill: { color: C.deep }, line: { color: C.deep } });
    s.addImage({ data: await icon(fa.FaLink, C.sun), x: 0.95, y: 5.97, w: 0.45, h: 0.45 });
    s.addText([
      { text: "学習指導要領 総則：", options: { fontSize: 12, color: C.sun, bold: true, breakLine: true } },
      { text: "「特別活動を要としつつ各教科等の特質に応じて、キャリア教育の充実を図ること」", options: { fontSize: 15, bold: true, color: C.white } },
    ], {
      x: 1.6, y: 5.75, w: 10.9, h: 0.9, fontFace: F, valign: "middle", margin: 0, isTextBox: true,
    });
    source(s, "小学校・中学校学習指導要領（平成29年告示）第1章 総則／教育課程企画特別部会「論点整理」（令和7年9月25日）");
    pageNo(s, pres.slides.length);
    s.addNotes("キャリア教育を単なる進路指導に閉じ込めてはいけません。各教科で好き・得意と出会い、総合のマイ探究で育て、特別活動でキャリア・パスポート等を用いて見通しと振り返りを行う。この三つが相互に連携する構造として教育課程を整理していきましょう。【根拠】現行の学習指導要領総則は、児童生徒が学ぶことと自己の将来とのつながりを見通しながら社会的・職業的自立に向けて必要な基盤となる資質・能力を身に付けていくことができるよう、「特別活動を要としつつ各教科等の特質に応じて、キャリア教育の充実を図ること」と定めています。キャリア・パスポートはこの特別活動を要とした取組の中で活用されています。");
  }

  // ---------- 8 外国語教育 ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.pale };
    title(s, "【外国語教育】AI時代に外国語を学ぶ本質的意義", "高精度な自動翻訳がある時代に、あえて学ぶ理由");
    const m = [
      [fa.FaGlobeAsia, "異文化への理解・包摂性", "多様な価値観を受け止め、共に生きる力"],
      [fa.FaMirror || fa.FaEye, "母語・自国文化のメタ認知", "他言語を通して自らの言葉と文化を見つめ直す"],
      [fa.FaComments, "生身のコミュニケーション意欲", "人と人とが直接つながりたいという思い"],
    ];
    const w = 3.8, g = 0.35;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (w + g);
      card(s, x, 1.95, w, 2.35);
      await iconCircle(s, m[i][0], x + 0.25, 2.15, 0.7, C.teal, C.white);
      s.addText(m[i][1], { x: x + 0.25, y: 2.98, w: w - 0.5, h: 0.45, fontFace: F, fontSize: 16, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(m[i][2], { x: x + 0.25, y: 3.48, w: w - 0.5, h: 0.7, fontFace: F, fontSize: 13, color: C.mute, margin: 0, isTextBox: true, valign: "top" });
    }
    // 往還
    s.addText("授業改善の方向", { x: 0.6, y: 4.6, w: 6, h: 0.45, fontFace: F, fontSize: 17, bold: true, color: C.deep, margin: 0, isTextBox: true });
    s.addShape("roundRect", { x: 0.6, y: 5.2, w: 3.6, h: 1.1, rectRadius: 0.1, fill: { color: C.deep }, line: { color: C.deep } });
    s.addText([{ text: "コミュニケーション活動", options: { bold: true, breakLine: true } }, { text: "思考力・判断力・表現力", options: { fontSize: 12, color: C.mint } }],
      { x: 0.6, y: 5.2, w: 3.6, h: 1.1, fontFace: F, fontSize: 15, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape("leftRightArrow", { x: 4.3, y: 5.5, w: 0.9, h: 0.5, fill: { color: C.sun }, line: { color: C.sun } });
    s.addShape("roundRect", { x: 5.3, y: 5.2, w: 3.0, h: 1.1, rectRadius: 0.1, fill: { color: C.teal }, line: { color: C.teal } });
    s.addText([{ text: "支える活動", options: { bold: true, breakLine: true } }, { text: "知識・技能", options: { fontSize: 12, color: C.mint } }],
      { x: 5.3, y: 5.2, w: 3.0, h: 1.1, fontFace: F, fontSize: 15, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });

    // 小中高接続
    s.addText("身近な話題 → 社会的な話題へ", { x: 8.8, y: 4.6, w: 3.9, h: 0.45, fontFace: F, fontSize: 17, bold: true, color: C.deep, margin: 0, isTextBox: true });
    ["小", "中", "高"].forEach((k, i) => {
      const x = 8.8 + i * 1.35;
      s.addShape("ellipse", { x, y: 5.2, w: 1.1, h: 1.1, fill: { color: [C.mint, C.teal, C.deep][i] }, line: { color: [C.mint, C.teal, C.deep][i] } });
      s.addText(k, { x, y: 5.2, w: 1.1, h: 1.1, fontFace: F, fontSize: 22, bold: true, color: i === 0 ? C.deep : C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("滑らかに接続", { x: 8.8, y: 6.3, w: 3.9, h: 0.4, fontFace: F, fontSize: 13, color: C.mute, align: "center", margin: 0, isTextBox: true });
    source(s, "教育課程部会 外国語WG 資料1「AI時代に外国語を学ぶ本質的意義」（令和7年10月30日）／教育課程企画特別部会「次期学習指導要領等に向けた審議まとめ（素案）」（令和8年8月31日）");
    pageNo(s, pres.slides.length);
    s.addNotes("高精度な自動翻訳が存在する時代にあえて外国語を学ぶ意義は、異文化への理解と包摂性、母語や自国文化のメタ認知、そして生身の人間同士のリアルなコミュニケーション意欲にあります。コミュニケーション活動と、それを支える知識・技能の活動を往還させ、身近な話題から社会的な話題へと発展させながら、小中高を滑らかに接続していきましょう。【根拠】外国語ワーキンググループでは「AI時代に外国語を学ぶ本質的意義」（令和7年10月30日 資料1）を議題として、言語を通じて伝え合う喜びや多様な他者との信頼関係の構築、外国の文化との比較を通じた自国の文化への理解の深まりなどが議論されています。また審議まとめ（素案）では、外国語科において自分の意見を形成・発信する活動や身近な地域のことを発信する活動等を充実させる方向が示されています。");
  }

  // ---------- 9 協議・まとめ ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.deep };
    s.addShape("ellipse", { x: 10.2, y: 4.6, w: 4.5, h: 4.5, fill: { color: C.teal, transparency: 55 }, line: { color: C.teal, transparency: 55 } });
    s.addText("【協議・まとめ】我が校の明日の一歩", { x: 0.6, y: 0.4, w: 12.1, h: 0.8, fontFace: F, fontSize: 30, bold: true, color: C.white, margin: 0, isTextBox: true, valign: "middle" });
    s.addText("グループ協議 15分 ＋ 指導主事からのエール 5分", { x: 0.6, y: 1.2, w: 12.1, h: 0.45, fontFace: F, fontSize: 16, color: C.mint, margin: 0, isTextBox: true });

    card(s, 0.6, 1.95, 12.1, 1.75, C.white);
    await iconCircle(s, fa.FaComments, 0.9, 2.4, 0.85, C.sun, C.deep);
    s.addText("10年後の子どもたちに「舵取りする力」を育むために、自校で「自由進度」「マイ探究」「キャリア教育」のどれから小さな挑戦（出島）を創めますか？", {
      x: 2.0, y: 2.05, w: 10.4, h: 1.55, fontFace: F, fontSize: 18, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true,
    });
    const opts = [["自由進度", fa.FaRoute], ["マイ探究", fa.FaSeedling], ["キャリア教育", fa.FaPassport]];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 2.9;
      s.addShape("roundRect", { x, y: 3.95, w: 2.6, h: 0.7, rectRadius: 0.35, fill: { color: C.teal }, line: { color: C.teal } });
      s.addImage({ data: await icon(opts[i][1], C.sun), x: x + 0.3, y: 4.1, w: 0.4, h: 0.4 });
      s.addText(opts[i][0], { x: x + 0.8, y: 3.95, w: 1.7, h: 0.7, fontFace: F, fontSize: 15, bold: true, color: C.white, valign: "middle", margin: 0, isTextBox: true });
    }
    s.addText("校長・副校長・教諭の\nミックスグループで", { x: 9.3, y: 3.95, w: 3.4, h: 0.7, fontFace: F, fontSize: 13, color: C.mint, valign: "middle", margin: 0, isTextBox: true });

    await iconCircle(s, fa.FaShip, 0.6, 5.1, 0.95, C.sun, C.deep);
    s.addText("学習指導要領は現場を縛る制限ではなく、学校と子どもたちが未来を自分たちの手で創るための「自由な出島」です。", {
      x: 1.8, y: 4.95, w: 9.0, h: 1.3, fontFace: F, fontSize: 19, bold: true, color: C.white, valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText("皆様の実践こそが、10年後の教育界の希望です。", { x: 1.8, y: 6.3, w: 9, h: 0.45, fontFace: F, fontSize: 15, color: C.sun, margin: 0, isTextBox: true });
    source(s, "教育課程企画特別部会「論点整理」（令和7年9月25日）：構造化・精選等による「余白」の創出、調整授業時数制度・裁量的な時間の検討（要旨）", true);
    s.addNotes("この後、管理職の先生方も若手の先生方も同じ輪に入り、『我が校なら明日からどんな小さなワクワク（出島）を創れるか』をぜひ語り合ってください。（協議後）学習指導要領は現場を縛る制限ではなく、学校と子どもたちが未来を自分たちの手で創るための自由な出島です。皆様の実践こそが、10年後の教育界の希望です。【根拠】論点整理では、学習指導要領の構造化・精選や標準授業時数の柔軟化等を通じて教師と子供の「余白」を生み出すこと、調整授業時数制度の下で裁量的な時間を設け、学校や子供の実態に応じた教育活動を行えるようにすることが検討されています。「出島」はこうした学校裁量の広がりを表す本講評での比喩です。");
  }

  // ---------- 参考資料 ----------
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    title(s, "参考資料（文部科学省・中央教育審議会）", "本講評の根拠資料。原文は文部科学省ウェブサイトでご確認ください");
    const refs = [
      ["諮問", "初等中等教育における教育課程の基準等の在り方について（諮問）", "令和6年12月25日", "mext.go.jp/b_menu/shingi/chukyo/chukyo0/toushin/mext_00003.html"],
      ["論点整理", "教育課程企画特別部会 論点整理", "令和7年9月25日", "mext.go.jp/b_menu/shingi/chukyo/chukyo3/004/gaiyou/mext_00010.html"],
      ["素案", "次期学習指導要領等に向けた審議まとめ（素案）（第17回 教育課程企画特別部会 資料1）", "令和8年8月31日", "mext.go.jp/b_menu/shingi/chukyo/chukyo3/101/siryo/mext_00056.html"],
      ["答申", "「令和の日本型学校教育」の構築を目指して（答申）【概要】", "令和3年1月26日", "mext.go.jp/content/20210126-mxt_syoto02-000012321_1-4.pdf"],
      ["外国語WG", "「AI時代に外国語を学ぶ本質的意義」（外国語WG 資料1）", "令和7年10月30日", "mext.go.jp/content/20251030-mxt_kyoiku01-000045617_003.pdf"],
      ["探究WG", "総合的な学習・探究の時間に関する現状・課題と検討事項", "令和7年10月15日", "mext.go.jp/content/20251015-mxt_kyoiku02-000045235_4.pdf"],
      ["告示", "小学校・中学校学習指導要領 第1章 総則（キャリア教育）", "平成29年3月告示", "mext.go.jp/a_menu/shotou/new-cs/"],
    ];
    const y0 = 1.95, rh = 0.66;
    refs.forEach((rf, i) => {
      const y = y0 + i * rh;
      if (i % 2 === 0) s.addShape("rect", { x: 0.6, y, w: 12.1, h: rh, fill: { color: C.pale }, line: { color: C.pale } });
      s.addShape("roundRect", { x: 0.75, y: y + 0.15, w: 1.25, h: 0.36, rectRadius: 0.06, fill: { color: C.deep }, line: { color: C.deep } });
      s.addText(rf[0], { x: 0.75, y: y + 0.15, w: 1.25, h: 0.36, fontFace: F, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(rf[1], { x: 2.2, y: y + 0.05, w: 7.9, h: 0.32, fontFace: F, fontSize: 13, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText("https://www." + rf[3], { x: 2.2, y: y + 0.36, w: 7.9, h: 0.26, fontFace: "Arial", fontSize: 9, color: C.teal, valign: "middle", margin: 0, isTextBox: true, hyperlink: { url: "https://www." + rf[3] } });
      s.addText(rf[2], { x: 10.2, y, w: 2.35, h: rh, fontFace: F, fontSize: 12, color: C.mute, align: "right", valign: "middle", margin: 0, isTextBox: true });
    });
    pageNo(s, pres.slides.length);
    s.addNotes("本講評で根拠とした文部科学省・中央教育審議会の資料一覧です。スライド中の引用のうち「」で示したものは原文、「要旨」と記したものは資料の趣旨を要約したものです。研修後に原文をご確認いただく際にご活用ください。");
  }

  await pres.writeFile({ fileName: require("path").join(__dirname, "..", "future_education_10years.pptx") });
  console.log("written");
})();
