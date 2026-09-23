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
    pageNo(s, 2);
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
    pageNo(s, 3);
    s.addNotes("まずはペアで、10年後の子どもの姿を思い描いてみましょう。AIが答えを出してくれる時代に、それでも学校に来てよかったと子どもが歓喜する瞬間とはどんな場面でしょうか。お一人1分半ずつ語ってください。管理職も教諭も、立場を超えてワクワクを共有する時間です。");
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
    s.addText("AI時代に不可欠な資質", { x: 0.9, y: 5.05, w: 4.5, h: 0.45, fontFace: F, fontSize: 13, color: C.sun, align: "center", margin: 0, isTextBox: true });

    const rows = [
      [fa.FaLightbulb, "自らの興味・問いを起点に", "学びと人生を主体的に進める"],
      [fa.FaHandsHelping, "身体性を伴う体験", "受動的な知識受容を超え、五感で学ぶ"],
      [fa.FaUsers, "他者との共創", "多様な仲間と対話し、新たな価値を創る"],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 1.95 + i * 1.3;
      card(s, 6.0, y, 6.7, 1.1);
      await iconCircle(s, rows[i][0], 6.2, y + 0.18, 0.74, C.teal, C.white);
      s.addText(rows[i][1], { x: 7.15, y: y + 0.14, w: 5.4, h: 0.45, fontFace: F, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(rows[i][2], { x: 7.15, y: y + 0.58, w: 5.4, h: 0.4, fontFace: F, fontSize: 14, color: C.mute, margin: 0, isTextBox: true });
    }
    s.addShape("roundRect", { x: 6.0, y: 5.9, w: 6.7, h: 0.8, rectRadius: 0.1, fill: { color: C.sun }, line: { color: C.sun } });
    s.addText("→ AI時代における「人間の強みの再定義」", { x: 6.2, y: 5.9, w: 6.4, h: 0.8, fontFace: F, fontSize: 17, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    pageNo(s, 4);
    s.addNotes("次期学習指導要領の素案を紐解くと、そこに流れているのは『教員の負担を増やす細かなルール』ではなく、『10年後の不確実な社会を生きる子どもたちが、自らの人生を誇りを持って舵取りしていくための道しるべ』です。AIが答えを即座に出してくれる時代だからこそ、学校は『答えを覚える場所』から、『自ら問いを立て、自分のペースで学びを深め、多様な他者と対話しながら新たな価値を創り出す場所』へとアップデートする必要があります。");
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
    ], { x: 0.9, y: 3.35, w: 3.8, h: 2.1, fontFace: F, fontSize: 15, color: C.ink, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 12 });

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
    s.addText("単元のゴールを見通させ、子ども自身が学びを調整する", {
      x: 5.4, y: 6.0, w: 7.3, h: 0.6, fontFace: F, fontSize: 14, color: C.deep, bold: true, margin: 0, isTextBox: true, valign: "middle",
    });
    pageNo(s, 5);
    s.addNotes("自由進度学習は、決して先生方の指導を放棄することではありません。単元のゴールを見通させ、子ども自身が学びを調整する自己調整学習と、それを見守り適時適切に介入する教師の指導性。この往還こそが本質です。先生方の温かい見取りと適切な投げかけがあってこそ、子どもは安心して試行錯誤できます。");
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
    pageNo(s, 6);
    s.addNotes("教師主導のテーマ探究から、個人の好き・得意を起点とするマイ探究への発展を評価したいと思います。試行錯誤の過程で素朴な問いが解像度の高い課題へと深まる変化を捉え、自己や他者にとっての新たな意味や理解の構築、すなわち自己変容そのものを成果として価値付けてください。");
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
      card(s, x, 2.0, w, 3.6, C.pale);
      await iconCircle(s, cols[i][0], x + (w - 1.0) / 2, 2.25, 1.0, cols[i][4], C.white);
      s.addText(cols[i][1], { x: x + 0.2, y: 3.4, w: w - 0.4, h: 0.5, fontFace: F, fontSize: 17, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
      s.addText(cols[i][2], { x: x + 0.2, y: 3.9, w: w - 0.4, h: 0.45, fontFace: F, fontSize: 15, bold: true, color: cols[i][4], align: "center", margin: 0, isTextBox: true });
      s.addText(cols[i][3], { x: x + 0.3, y: 4.45, w: w - 0.6, h: 1.0, fontFace: F, fontSize: 13, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
    }
    // 連携を示すベース
    s.addShape("roundRect", { x: 0.6, y: 5.9, w: 12.1, h: 0.85, rectRadius: 0.1, fill: { color: C.deep }, line: { color: C.deep } });
    s.addImage({ data: await icon(fa.FaLink, C.sun), x: 0.95, y: 6.1, w: 0.45, h: 0.45 });
    s.addText("三つが相互に連携し、「好き」を育み「得意」を伸ばす教育課程の全体像をつくる", {
      x: 1.6, y: 5.9, w: 10.9, h: 0.85, fontFace: F, fontSize: 16, bold: true, color: C.white, valign: "middle", margin: 0, isTextBox: true,
    });
    pageNo(s, 7);
    s.addNotes("キャリア教育を単なる進路指導に閉じ込めてはいけません。各教科で好き・得意と出会い、総合のマイ探究で育て、特別活動でキャリア・パスポート等を用いて見通しと振り返りを行う。この三つが相互に連携する構造として教育課程を整理していきましょう。");
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
    s.addText("滑らかに接続", { x: 8.8, y: 6.4, w: 3.9, h: 0.4, fontFace: F, fontSize: 13, color: C.mute, align: "center", margin: 0, isTextBox: true });
    pageNo(s, 8);
    s.addNotes("高精度な自動翻訳が存在する時代にあえて外国語を学ぶ意義は、異文化への理解と包摂性、母語や自国文化のメタ認知、そして生身の人間同士のリアルなコミュニケーション意欲にあります。コミュニケーション活動と、それを支える知識・技能の活動を往還させ、身近な話題から社会的な話題へと発展させながら、小中高を滑らかに接続していきましょう。");
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
    s.addNotes("この後、管理職の先生方も若手の先生方も同じ輪に入り、『我が校なら明日からどんな小さなワクワク（出島）を創れるか』をぜひ語り合ってください。（協議後）学習指導要領は現場を縛る制限ではなく、学校と子どもたちが未来を自分たちの手で創るための自由な出島です。皆様の実践こそが、10年後の教育界の希望です。");
  }

  await pres.writeFile({ fileName: require("path").join(__dirname, "..", "future_education_10years.pptx") });
  console.log("written");
})();
