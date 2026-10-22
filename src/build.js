const pptxgen = require("pptxgenjs");
const { applyTheme } = require("C:/Users/yokot/AppData/Local/Claude-Data/claude_pzs8sxrjxfjjc/local-agent-mode-sessions/skills-plugin/e4bbdc80-dad8-4a34-81b6-1ade1e249636/8bedad44-bfe6-49b0-8758-cde20c61d165/skills/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const TITLEF = "Noto Sans JP Print Black";
const HEAD = "Noto Sans JP Print";
const BODY = "Noto Sans JP Print";
const NUM = "Noto Sans JP Print Black";
const THEME = {
  name: "Marathon Lecture",
  headFontFace: TITLEF,
  bodyFontFace: BODY,
  colors: {
    dk1: "1E2533", lt1: "FFFFFF", dk2: "0B1F3A", lt2: "F0F2F5",
    accent1: "B7862B", accent2: "5B7083", accent3: "F2B33D", accent4: "727A86",
    accent5: "F7EFDF", accent6: "E7ECF1", hlink: "0B1F3A", folHlink: "727A86",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: TITLEF, bodyFontFace: BODY };
pres.title = "走るのが好きが、いちばんの体力づくり";
const C = pres.SchemeColor;
const NAVY = C.text2, ORANGE = C.accent1, GREEN = C.accent2, GRAY = C.accent4;
const TINT = C.background2, OTINT = C.accent5, GTINT = C.accent6, WHITE = C.background1, TEXT = C.text1;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: NAVY },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.6, w: 11.7, h: 2.4, fontSize: 48, bold: false, fontFace: TITLEF, color: WHITE, valign: "bottom", margin: 0, align: "left", lineSpacingMultiple: 1.1 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.3, w: 11.7, h: 1.6, fontSize: 20, color: OTINT, valign: "top", margin: 0 }, text: "" } },
  ],
});
const contentObjects = () => [
  { placeholder: { options: { name: "sec", type: "body", x: 0.6, y: 0.3, w: 8, h: 0.35, fontSize: 14, bold: true, color: ORANGE, valign: "middle", margin: 0 }, text: "" } },
  { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.65, w: 12.1, h: 1.45, fontSize: 36, bold: false, fontFace: TITLEF, color: NAVY, valign: "top", margin: 0, align: "left", lineSpacingMultiple: 0.9 }, text: "" } },
];
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: WHITE },
  objects: contentObjects(),
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 11, color: "727A86", fontFace: NUM, align: "right" },
});
pres.defineSlideMaster({
  title: "WORK",
  background: { color: OTINT },
  objects: contentObjects(),
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 11, color: "727A86", fontFace: NUM, align: "right" },
});

// ---------- helpers ----------
let sectionName = "";
function section(title) { sectionName = title; pres.addSection({ title }); }
function content(sec, title, notes, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master, sectionTitle: sectionName });
  s.addText(sec, { placeholder: "sec" });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function T(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 16, color: TEXT, valign: "top", fontFace: o && o.bold ? HEAD : BODY }, o));
}
function box(s, x, y, w, h, fill, extra) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.05, fill: { color: fill }, line: { color: fill, width: 0.5 } }, extra || {}));
}
function circle(s, x, y, d, label, fill, fontSize) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0.5 } });
  T(s, label, { x, y, w: d, h: d, align: "center", valign: "middle", bold: true, color: WHITE, fontSize: fontSize || 20, fontFace: /^[\x20-\x7E]+$/.test(label) ? NUM : HEAD });
}
function arrow(s, x, y, w, h, color) {
  s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h, fill: { color: color || GRAY }, line: { color: color || GRAY, width: 0.5 } });
}
// card with number circle, heading, body
function card(s, x, y, w, h, num, head, body, opt) {
  opt = opt || {};
  box(s, x, y, w, h, opt.fill || TINT);
  let tx = x + 0.3;
  if (num) { circle(s, x + 0.3, y + 0.3, 0.6, num, opt.numFill || ORANGE, 18); }
  const hy = num ? y + 1.05 : y + 0.3;
  T(s, head, { x: tx, y: hy, w: w - 0.6, h: 0.9, fontSize: opt.headSize || 20, bold: true, color: opt.headColor || NAVY });
  if (body) T(s, body, { x: tx, y: hy + (opt.bodyGap || 0.85), w: w - 0.6, h: h - (hy - y) - (opt.bodyGap || 0.85) - 0.2, fontSize: opt.bodySize || 15, color: TEXT });
}
function src(s, text) { T(s, text, { x: 0.6, y: 6.95, w: 11, h: 0.35, fontSize: 10, color: GRAY, valign: "middle" }); }
function bullets(items, o) {
  return items.map((t, i) => ({ text: t, options: Object.assign({ bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 8 }, o || {}) }));
}

// ================= 導入 =================
section("導入");
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: sectionName });
  s.addText("「走るのが好き」が、\nいちばんの体力づくり", { placeholder: "title" });
  s.addText("マラソン大会に向けた「持久走」の指導を考える\n校内研修　講師：〇〇〇〇", { placeholder: "body" });
  T(s, "RUN", { x: 0.8, y: 0.7, w: 4, h: 0.5, fontSize: 16, bold: true, color: ORANGE, fontFace: NUM, charSpacing: 6 });
  s.addNotes("【約30秒】\nよろしくお願いします。今日は、マラソン大会に向けた「持久走」の取り組みについてお話しします。\n体育が専門の先生だけでなく、全教職員の皆さんに聞いていただきたい内容です。大会当日は、全員がコースのどこかで子どもを見守ることになるからです。\n途中で3回、近くの方と話し合っていただく時間をとります。");
}
{
  const s = content("講師紹介", "大切にしてきたのは\n「納得できる、理由のある努力」",
    "【約1分】\n簡単に自己紹介をします。中学校で駅伝部・陸上競技部の指導をしてきました。駅伝では全国大会に3回出場し、8位・9位・20位。陸上競技部では市の総体で12年連続優勝し、獲得した優勝旗は48本になりました。\n（実績は短く。強調するのは次の一言）\nただ、私が一番大切にしてきたのは、勝つことそのものではなく、「なぜこの練習をするのか」を選手が納得して取り組むことです。納得できる、理由のある努力。今日の話は、全部この考え方につながっています。");
  const stats = [[["3", "回"], "駅伝 全国大会出場", "8位・9位・20位"], [["12", "年連続"], "市総体 優勝", "陸上競技部"], [["48", "本"], "獲得した優勝旗", ""]];
  stats.forEach((st, i) => {
    const y = 2.25 + i * 1.5;
    T(s, [{ text: st[0][0], options: { fontFace: NUM, fontSize: 48 } }, { text: st[0][1], options: { fontFace: HEAD, fontSize: 24 } }], { x: 0.6, y, w: 3.0, h: 0.8, bold: true, color: ORANGE, valign: "middle" });
    T(s, st[1], { x: 3.7, y: y + 0.05, w: 3.2, h: 0.4, fontSize: 18, bold: true, color: NAVY });
    if (st[2]) T(s, st[2], { x: 3.7, y: y + 0.45, w: 3.2, h: 0.35, fontSize: 14, color: GRAY });
  });
  box(s, 7.4, 2.25, 5.3, 4.3, NAVY);
  T(s, "指導で心がけてきたこと", { x: 7.8, y: 2.6, w: 4.5, h: 0.4, fontSize: 16, color: OTINT, bold: true });
  T(s, "納得できる、\n理由のある努力", { x: 7.8, y: 3.3, w: 4.5, h: 1.6, fontSize: 34, bold: true, color: WHITE });
  T(s, "「なぜこの練習をするのか」を\n選手自身がわかっていること", { x: 7.8, y: 5.2, w: 4.5, h: 1.0, fontSize: 16, color: WHITE });
}
{
  const s = content("問いかけ", "子どもたちは、\nマラソン大会を楽しみにしていますか？",
    "【約1分】\n最初に1つ質問です。皆さんのクラスの子どもたちは、マラソン大会を楽しみにしているでしょうか。\n「楽しみにしている子が多い」と思う方、手を挙げてください。……「いやがっている子が多い」と思う方は？\n（会場の反応を受けて）ありがとうございます。いやがる子がいるのは、子どもが悪いわけではありません。大会の位置づけや、練習のやり方に理由があるのかもしれません。今日はそこを一緒に考えていきます。");
  const opts = [["楽しみにしている子が多い", GREEN, GTINT], ["半々くらい", GRAY, TINT], ["いやがっている子が多い", ORANGE, OTINT]];
  opts.forEach((o, i) => {
    const x = 0.6 + i * 4.15;
    box(s, x, 2.6, 3.85, 3.0, o[2]);
    circle(s, x + 1.425, 2.95, 1.0, String(i + 1), o[1], 28);
    T(s, o[0], { x: x + 0.3, y: 4.25, w: 3.25, h: 1.0, fontSize: 18, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  T(s, "挙手でお答えください", { x: 0.6, y: 6.0, w: 12.1, h: 0.5, fontSize: 18, color: GRAY, align: "center" });
}

// ================= 1. 位置づけ =================
section("1. 位置づけ");
{
  const s = content("1. 長距離走の位置づけ", "小学校の体育に「長距離走」はない",
    "【約1分】\n意外に思われるかもしれませんが、小学校の学習指導要領の陸上運動には「長距離走」がありません。\n高学年の陸上運動は、短距離走・リレー、ハードル走、走り幅跳び、走り高跳びの4つです。\n長距離走が出てくるのは中学校からで、1000〜3000m程度とされています。\nでは、小学校で長く走る運動はどこにあるのでしょうか。");
  box(s, 0.6, 2.3, 5.8, 4.3, TINT);
  T(s, "小学校（高学年）陸上運動", { x: 0.95, y: 2.6, w: 5.1, h: 0.5, fontSize: 20, bold: true, color: NAVY });
  T(s, bullets(["短距離走・リレー", "ハードル走", "走り幅跳び", "走り高跳び"]), { x: 0.95, y: 3.3, w: 5.1, h: 2.4, fontSize: 18 });
  T(s, "長距離走は……ない", { x: 0.95, y: 5.75, w: 5.1, h: 0.5, fontSize: 18, bold: true, color: ORANGE });
  arrow(s, 6.6, 4.2, 0.6, 0.5);
  box(s, 7.4, 2.3, 5.3, 4.3, NAVY);
  T(s, "中学校 陸上競技", { x: 7.75, y: 2.6, w: 4.6, h: 0.5, fontSize: 20, bold: true, color: OTINT });
  T(s, "長距離走", { x: 7.75, y: 3.5, w: 4.6, h: 0.9, fontSize: 40, bold: true, color: WHITE });
  T(s, "1000〜3000m程度", { x: 7.75, y: 4.5, w: 4.6, h: 0.6, fontSize: 24, color: WHITE });
  T(s, "「長距離走」として扱うのは中学校から", { x: 7.75, y: 5.6, w: 4.6, h: 0.6, fontSize: 15, color: OTINT });
  src(s, "出典：小学校・中学校学習指導要領（平成29年告示）");
}
{
  const s = content("1. 長距離走の位置づけ", "あるのは\n「無理のない速さで5〜6分程度」の持久走",
    "【約1分30秒】\n小学校で長く走る運動は、「体つくり運動」の中にあります。体ほぐしの運動ではなく、「体の動きを高める運動」の、「動きを持続する能力を高めるための運動」です。\n高学年の解説の例示には、「無理のない速さで5〜6分程度の持久走」とあります。ほかにも、エアロビクスなどの全身運動を続けること、変化のある地形を移動することなどが挙げられています。\n（問いかけ）皆さんの学校の大会では、子どもたちは何分くらい走っているでしょうか。\n※低・中学年の時間の目安は二次資料で2〜3分、3〜4分とされることがある。使う場合は解説の原文で確認すること。");
  // hierarchy
  box(s, 0.6, 2.3, 3.6, 0.8, NAVY);
  T(s, "体つくり運動", { x: 0.6, y: 2.3, w: 3.6, h: 0.8, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
  box(s, 0.9, 3.4, 3.3, 0.7, TINT);
  T(s, "体ほぐしの運動", { x: 0.9, y: 3.4, w: 3.3, h: 0.7, fontSize: 16, color: GRAY, align: "center", valign: "middle" });
  box(s, 0.9, 4.3, 3.3, 0.7, TINT);
  T(s, "体の動きを高める運動", { x: 0.9, y: 4.3, w: 3.3, h: 0.7, fontSize: 16, bold: true, color: NAVY, align: "center", valign: "middle" });
  box(s, 1.2, 5.2, 3.0, 1.0, OTINT);
  T(s, "動きを持続する能力を\n高めるための運動", { x: 1.2, y: 5.2, w: 3.0, h: 1.0, fontSize: 15, bold: true, color: ORANGE, align: "center", valign: "middle" });
  arrow(s, 4.5, 5.45, 0.6, 0.5, ORANGE);
  // callout
  T(s, "高学年の例示", { x: 5.5, y: 2.3, w: 7.2, h: 0.4, fontSize: 16, bold: true, color: GRAY });
  T(s, [{ text: "5", options: { fontSize: 96, bold: true, color: ORANGE, fontFace: NUM } }, { text: "〜", options: { fontSize: 60, bold: true, color: ORANGE, fontFace: HEAD } }, { text: "6", options: { fontSize: 96, bold: true, color: ORANGE, fontFace: NUM } }, { text: " 分程度", options: { fontSize: 32, bold: true, color: NAVY, fontFace: HEAD } }], { x: 5.5, y: 2.8, w: 7.2, h: 1.5, valign: "middle" });
  T(s, "無理のない速さで", { x: 5.5, y: 4.3, w: 7.2, h: 0.5, fontSize: 22, bold: true, color: NAVY });
  T(s, bullets(["エアロビクスなどの全身運動を続ける", "変化のある地形などを移動する"]), { x: 5.5, y: 5.0, w: 7.2, h: 1.2, fontSize: 16 });
  src(s, "出典：小学校学習指導要領（平成29年告示）解説 体育編");
}
{
  const s = content("1. 長距離走の位置づけ", "大会は特別活動の「体育的行事」。\nねらいの一つは「運動に親しむ態度」",
    "【約1分】\n授業の持久走とは別に、マラソン大会は特別活動の学校行事、「健康安全・体育的行事」に当たります。\n指導要領が示すねらいには、心身の健全な発達や健康の保持増進、運動に親しむ態度の育成、責任感や連帯感の涵養、体力の向上などが並んでいます。\n「体力の向上」だけでなく、「運動に親しむ態度」もねらいに入っていることを、ぜひ心に留めておいてください。");
  const chips = [["心身の健全な発達", false], ["健康の保持増進", false], ["運動に親しむ態度の育成", true], ["責任感や連帯感の涵養", false], ["体力の向上", false]];
  T(s, "健康安全・体育的行事のねらい（要旨）", { x: 0.6, y: 2.3, w: 12, h: 0.45, fontSize: 18, bold: true, color: GRAY });
  chips.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.15, y = 3.0 + row * 1.55;
    box(s, x, y, 3.85, 1.25, c[1] ? ORANGE : TINT);
    T(s, c[0], { x: x + 0.2, y, w: 3.45, h: 1.25, fontSize: 19, bold: true, color: c[1] ? WHITE : NAVY, align: "center", valign: "middle" });
  });
  T(s, "「体力の向上」と並んで\n「親しむ」がある", { x: 8.9, y: 4.55, w: 3.8, h: 1.25, fontSize: 18, bold: true, color: ORANGE, valign: "middle" });
  src(s, "出典：小学校学習指導要領（平成29年告示）第6章 特別活動");
}
{
  const s = content("1. 長距離走の位置づけ（余談）", "運動会のルーツは、\n海軍の「隊の強さ比べ」",
    "【約1分・時間がなければ省略可】\n少し余談です。日本の運動会のルーツの一つは、1874年（明治7年）に海軍兵学寮で行われた「競闘遊戯」だと言われています。小隊ごとに訓練の成果を競い合う、いわば「どの隊が一番強いか」を比べる場でした。\nその後、明治期に学校の運動会が全国に広がります。目的は体育の普及でしたが、内容には軍事的な色合いもありました。\n学習指導要領ができたのは戦後です。最初は1947年の「試案」で、法的な拘束力を持つ告示になったのは1958年です。\nつまり、「競わせる行事」は指導要領より前からありました。私たちの大会にも、その名残があるかもしれません。");
  const tl = [["1874年", "海軍兵学寮\n「競闘遊戯」", "小隊ごとに訓練の成果を競う"], ["明治期", "学校運動会が\n全国に普及", "体育の普及が目的"], ["1947年", "学習指導要領\n（試案）", "戦後、初めての指導要領"], ["1958年", "学習指導要領\n（告示）", "法的な拘束力を持つ"]];
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 3.35, w: 11.5, h: 0, line: { color: GRAY, width: 2 } });
  tl.forEach((t, i) => {
    const x = 0.6 + i * 3.1;
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 3.2, w: 0.3, h: 0.3, fill: { color: i === 0 ? ORANGE : NAVY }, line: { color: WHITE, width: 1.5 } });
    T(s, t[0], { x, y: 2.4, w: 2.9, h: 0.6, fontSize: 28, bold: true, color: i === 0 ? ORANGE : NAVY, fontFace: /^\d/.test(t[0]) ? NUM : HEAD });
    box(s, x, 3.85, 2.85, 2.4, i === 0 ? OTINT : TINT);
    T(s, t[1], { x: x + 0.2, y: 4.05, w: 2.45, h: 1.0, fontSize: 18, bold: true, color: NAVY });
    T(s, t[2], { x: x + 0.2, y: 5.1, w: 2.45, h: 1.0, fontSize: 15, color: TEXT });
  });
  src(s, "「競わせる行事」は、学習指導要領より前からあった");
}
{
  const s = content("1. 長距離走の位置づけ", "私たちの大会は\n「行事」か、それとも「訓練」か",
    "【約30秒】\nここまでをまとめると、問いはこうなります。\n私たちのマラソン大会は、運動に親しみ、体力を高めるための「行事」でしょうか。それとも、苦しさに耐えることや競わせることが目的の「訓練」になっていないでしょうか。\n答えはここでは出しません。この問いを持ったまま、次へ進みます。");
  box(s, 0.6, 2.4, 5.6, 3.8, GTINT);
  T(s, "行事", { x: 0.6, y: 2.7, w: 5.6, h: 1.0, fontSize: 44, bold: true, color: GREEN, align: "center" });
  T(s, "運動に親しむ\n体力を高める", { x: 0.9, y: 4.0, w: 5.0, h: 1.6, fontSize: 22, color: NAVY, align: "center", bold: true });
  T(s, "?", { x: 6.2, y: 3.5, w: 0.95, h: 1.4, fontSize: 72, bold: true, color: ORANGE, align: "center", valign: "middle", fontFace: NUM });
  box(s, 7.15, 2.4, 5.6, 3.8, OTINT);
  T(s, "訓練", { x: 7.15, y: 2.7, w: 5.6, h: 1.0, fontSize: 44, bold: true, color: ORANGE, align: "center" });
  T(s, "苦しさに耐える\n競わせる", { x: 7.45, y: 4.0, w: 5.0, h: 1.6, fontSize: 22, color: NAVY, align: "center", bold: true });
}

// ================= 2. 目指す姿 =================
section("2. 目指す姿");
{
  const s = content("2. 目指す子どもの姿", "全国で戦った選手にも、\n「歯を食いしばれ」とは言わなかった",
    "【約1分30秒】\nここは私自身の経験からお話しします。\n全国大会で戦う駅伝の選手たちにも、私は「歯を食いしばって耐えろ」とは求めてきませんでした。\n（具体的なエピソードを1つ入れる：例）ある選手が伸びたのは、苦しい練習に耐えたからではなく、「この練習は何のためか」がわかって、自分でペースを考えるようになってからでした。\n全国レベルの中学生ですらそうです。まして小学生に、苦しさに耐える力を求める必要はありません。学習指導要領のどこにも、「耐える力」は書かれていません。");
  box(s, 0.6, 2.5, 5.8, 3.8, TINT);
  circle(s, 0.95, 2.85, 0.8, "✕", GRAY, 30);
  T(s, "歯を食いしばって\n耐える", { x: 0.95, y: 3.9, w: 5.1, h: 1.4, fontSize: 28, bold: true, color: GRAY });
  T(s, "どこからも求められていない", { x: 0.95, y: 5.45, w: 5.1, h: 0.5, fontSize: 16, color: GRAY });
  box(s, 6.9, 2.5, 5.8, 3.8, NAVY);
  circle(s, 7.25, 2.85, 0.8, "○", ORANGE, 30);
  T(s, "理由をわかって、\n自分で考えて走る", { x: 7.25, y: 3.9, w: 5.1, h: 1.4, fontSize: 28, bold: true, color: WHITE });
  T(s, "強くなった選手に共通していたこと", { x: 7.25, y: 5.45, w: 5.1, h: 0.5, fontSize: 16, color: OTINT });
}
{
  const s = content("2. 目指す子どもの姿", "目指すのは、苦しさをコントロールし、\n伸びを実感して、もっと走りたくなる子",
    "【約1分】\nでは、どんな姿を目指すのか。私は次の3段階だと考えています。\n1つ目は、苦しさをコントロールできること。自分でペースを調整し、苦しくなりすぎないように走れる。\n2つ目は、伸びを実感できること。「前より楽に走れた」「前より長く走れた」と自分でわかる。\n3つ目は、その結果「もっと走りたい」と意欲的になること。\nこれが理想の姿です。");
  const steps = [["1", "苦しさを\nコントロールする", "自分でペースを調整できる"], ["2", "伸びを\n実感する", "「前より楽」「前より長く」\nがわかる"], ["3", "もっと\n走りたくなる", "次への意欲が生まれる"]];
  steps.forEach((st, i) => {
    const x = 0.6 + i * 4.2;
    card(s, x, 2.5, 3.6, 3.9, st[0], st[1], st[2], { headSize: 24, bodyGap: 1.15, bodySize: 16, fill: i === 2 ? OTINT : TINT });
    if (i < 2) arrow(s, x + 3.65, 4.2, 0.5, 0.5, ORANGE);
  });
}
{
  const s = content("話し合い①　ペアで3分", "マラソン大会で、\n子どものどんな姿が見たいですか？",
    "【約3分＋共有1分】\nここで1回目の話し合いです。隣の方とペアになって、3分間話してください。\nテーマは「マラソン大会で、子どものどんな姿が見たいか」。練習中、大会当日、大会が終わった後、どの場面でもかまいません。\n（3分後、2〜3組に発表してもらう）\n（まとめ）「最後まであきらめない姿」「友達を応援する姿」などが出たと思います。その中に「苦しそうに歯を食いしばる姿」はあったでしょうか。たぶん、先ほどの3段階に近い姿が多かったはずです。", "WORK");
  const sc = [["練習中", "どんな顔で走っている？"], ["大会当日", "スタート前・\nゴールの瞬間は？"], ["大会後", "どんな言葉を口にしている？"]];
  sc.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    box(s, x, 2.5, 3.85, 2.6, WHITE);
    T(s, c[0], { x: x + 0.3, y: 2.8, w: 3.25, h: 0.6, fontSize: 24, bold: true, color: ORANGE });
    T(s, c[1], { x: x + 0.3, y: 3.6, w: 3.25, h: 1.2, fontSize: 18, color: NAVY });
  });
  circle(s, 0.6, 5.55, 1.0, "3分", ORANGE, 22);
  T(s, "隣の方とペアで話し合い、あとで2〜3組に紹介していただきます", { x: 1.9, y: 5.55, w: 10.8, h: 1.0, fontSize: 18, color: NAVY, valign: "middle", bold: true });
}

// ================= 3. 社会の中の「走る」 =================
section("3. 社会の中の走ること");
{
  const s = content("3. 社会の中の「走ること」", "大人にとって、走ることは\n「きつい運動」ではなく「気軽な運動」",
    "【約1分】\n学校の外に目を向けてみます。ジョギングは、大人がよく行う代表的な余暇のスポーツです。市民マラソンも各地で開かれ、多くの人が参加しています。\nでも、大人が走る理由は「健康のため」「気分転換」「自分のペースで」です。きつい運動として走っているわけではありません。\n一方で、学校の持久走には「きつい」「苦しい」「順位」というイメージがつきがちです。このずれが、子どもの「走るのが嫌い」につながっているのかもしれません。\n※レジャー白書の順位を示す場合は、最新版の数値を確認して差し替える。");
  box(s, 0.6, 2.5, 5.8, 3.9, GTINT);
  T(s, "大人が走る理由", { x: 0.95, y: 2.8, w: 5.1, h: 0.6, fontSize: 24, bold: true, color: GREEN });
  T(s, bullets(["健康のため", "気分転換・ストレス解消", "自分のペースで楽しむ"]), { x: 0.95, y: 3.6, w: 5.1, h: 2.4, fontSize: 20, color: NAVY });
  box(s, 6.9, 2.5, 5.8, 3.9, OTINT);
  T(s, "学校の持久走のイメージ", { x: 7.25, y: 2.8, w: 5.1, h: 0.6, fontSize: 24, bold: true, color: ORANGE });
  T(s, bullets(["きつい・苦しい", "止まってはいけない", "順位で比べられる"]), { x: 7.25, y: 3.6, w: 5.1, h: 2.4, fontSize: 20, color: NAVY });
}
{
  const s = content("3. 社会の中の「走ること」", "週1回以上運動する大人は、\n約半分しかいない",
    "【約1分・時間がなければ省略可】\nスポーツ庁の令和6年度の調査では、週1日以上運動やスポーツをしている20歳以上の人は52.5%でした。国の目標は70%なので、まだ大きな差があります。\n運動を「嫌いなもの」として覚えてしまうと、大人になってからも体を動かす習慣がつきにくくなります。\n小学校の持久走で「嫌い」にさせないことには、一生に関わる意味があります。");
  T(s, [{ text: "52.5", options: { fontSize: 110, bold: true, color: ORANGE, fontFace: NUM } }, { text: " %", options: { fontSize: 44, bold: true, color: ORANGE, fontFace: NUM } }], { x: 0.6, y: 2.4, w: 5.5, h: 1.8, valign: "middle" });
  T(s, "週1日以上 運動・スポーツをする\n20歳以上の割合（令和6年度）", { x: 0.6, y: 4.3, w: 5.5, h: 1.0, fontSize: 18, color: NAVY, bold: true });
  // bar
  const bx = 6.7, bw = 6.0;
  T(s, "実際", { x: bx, y: 2.6, w: 2, h: 0.4, fontSize: 16, color: GRAY, bold: true });
  box(s, bx, 3.05, bw, 0.6, TINT);
  box(s, bx, 3.05, bw * 0.525, 0.6, ORANGE);
  T(s, "52.5%", { x: bx + 0.15, y: 3.05, w: 2, h: 0.6, fontSize: 18, color: WHITE, bold: true, valign: "middle", fontFace: NUM });
  T(s, "国の目標", { x: bx, y: 4.0, w: 2, h: 0.4, fontSize: 16, color: GRAY, bold: true });
  box(s, bx, 4.45, bw, 0.6, TINT);
  box(s, bx, 4.45, bw * 0.7, 0.6, NAVY);
  T(s, "70%", { x: bx + 0.15, y: 4.45, w: 2, h: 0.6, fontSize: 18, color: WHITE, bold: true, valign: "middle", fontFace: NUM });
  box(s, 0.6, 5.65, 12.1, 0.95, TINT);
  T(s, "小学校で「嫌い」にさせないことは、一生の運動習慣に関わる", { x: 0.9, y: 5.65, w: 11.5, h: 0.95, fontSize: 20, bold: true, color: NAVY, valign: "middle" });
  src(s, "出典：スポーツ庁「スポーツの実施状況等に関する世論調査」（令和6年度）、第3期スポーツ基本計画");
}

// ================= 4. 苦しいの正体 =================
section("4. 苦しいの正体");
{
  const s = content("4. 「苦しい」の正体", "子どもが走れなくなるのは「苦しい」から。\nその正体は？",
    "【約30秒】\nでは、子どもはなぜ長く走り続けられないのでしょうか。子どもに聞くと、「疲れた」「きつい」「苦しい」という答えが返ってきます。\nこの「苦しい」の正体は何なのか。体の中で何が起きているのかを、簡単に見ていきます。");
  const v = [["疲れた…", 1.0, 2.6], ["きつい…", 4.9, 3.2], ["苦しい…", 8.8, 2.6]];
  v.forEach((b) => {
    s.addShape(pres.shapes.OVAL, { x: b[1], y: b[2], w: 3.4, h: 1.8, fill: { color: OTINT }, line: { color: OTINT, width: 0.5 } });
    T(s, b[0], { x: b[1], y: b[2], w: 3.4, h: 1.8, fontSize: 30, bold: true, color: ORANGE, align: "center", valign: "middle" });
  });
  T(s, "その正体は？", { x: 0.6, y: 5.5, w: 12.1, h: 0.9, fontSize: 32, bold: true, color: NAVY, align: "center", valign: "middle" });
}
{
  const s = content("4. 「苦しい」の正体", "長く動くには酸素がいる。\nだから息と心臓が速くなる",
    "【約1分30秒】\n体を動かすエネルギーのつくり方は、大きく3つあります。\n1つ目は、筋肉にためてあるエネルギーをすぐに使う仕組みで、10秒ほどしかもちません（ATPとクレアチンリン酸）。\n2つ目は、糖を素早く分解する仕組みで、2分くらいまでの運動の主役です。\n3つ目は、酸素を使って糖や脂肪を燃やす仕組みで、長く動くときの主役です。\n長く走るほど3つ目の仕組みが中心になり、たくさんの酸素が必要になります。だから呼吸が速くなり、心臓も速く打ちます。\n※補足：乳酸は昔「疲労物質」と言われていたが、現在はエネルギー源の一つとして扱われている。");
  const sys = [["〜10秒", "すぐ使える\nエネルギー", "筋肉にためてある分を使う", TINT, NAVY], ["〜2分", "糖を素早く\n分解する", "短い全力の運動の主役", TINT, NAVY], ["2分〜", "酸素を使って\n糖・脂肪を燃やす", "長く動くときの主役", OTINT, ORANGE]];
  sys.forEach((y, i) => {
    const x = 0.6 + i * 4.15;
    box(s, x, 2.3, 3.85, 3.05, y[3]);
    T(s, y[0], { x: x + 0.3, y: 2.5, w: 3.25, h: 0.6, fontSize: 26, bold: true, color: y[4] });
    T(s, y[1], { x: x + 0.3, y: 3.15, w: 3.25, h: 1.1, fontSize: 20, bold: true, color: NAVY });
    T(s, y[2], { x: x + 0.3, y: 4.4, w: 3.25, h: 0.7, fontSize: 15, color: TEXT });
  });
  arrow(s, 9.95, 5.45, 0.5, 0.45, ORANGE);
  box(s, 0.6, 6.0, 12.1, 0.85, NAVY);
  T(s, "酸素をたくさん取り込もうとする　→　呼吸が速くなる・心拍数が上がる", { x: 0.9, y: 6.0, w: 11.5, h: 0.85, fontSize: 20, bold: true, color: WHITE, valign: "middle" });
}
{
  const s = content("4. 「苦しい」の正体", "「苦しい」は、脳がかけるブレーキ。\n慣れると感じ方が変わる",
    "【約1分30秒】\n呼吸や心拍数が上がると、脳がそれを受け取って「苦しい」と感じます。これは、体を守るために脳がかけるブレーキのようなものです。\n大事なのは、このブレーキのかかり方は変えられるということです。同じ心拍数でも、慣れている人は「まだ大丈夫」と感じます。ゴールまでの見通しがあるときも、苦しさの感じ方は軽くなります。\nつまり、「苦しい」は体の限界そのものではなく、脳の感じ方です。練習で変えていくことができます。\nただし、本当に危険なサインもあります。それは最後の安全の話でお伝えします。");
  const fl = [["呼吸・心拍数が\n上がる", TINT, NAVY], ["脳が\n「苦しい」と感じる", TINT, NAVY], ["ブレーキが\nかかる", OTINT, ORANGE]];
  fl.forEach((f, i) => {
    const x = 0.6 + i * 4.15;
    box(s, x, 2.4, 3.6, 1.8, f[1]);
    T(s, f[0], { x, y: 2.4, w: 3.6, h: 1.8, fontSize: 22, bold: true, color: f[2], align: "center", valign: "middle" });
    if (i < 2) arrow(s, x + 3.65, 3.05, 0.45, 0.5);
  });
  box(s, 0.6, 4.6, 12.1, 1.7, GTINT);
  T(s, "感じ方は変えられる", { x: 0.95, y: 4.8, w: 4.2, h: 0.6, fontSize: 24, bold: true, color: GREEN });
  T(s, bullets(["同じ心拍数でも、慣れると「まだ大丈夫」と感じる", "ゴールまでの見通しがあると、苦しさは軽くなる"]), { x: 5.2, y: 4.8, w: 7.3, h: 1.4, fontSize: 17, color: NAVY });
  T(s, "※胸の痛み・めまい・顔色の悪さは「本当の危険サイン」（→ 8章）", { x: 0.6, y: 6.5, w: 12.1, h: 0.4, fontSize: 14, color: ORANGE, bold: true });
}
{
  const s = content("4. 「苦しい」の正体", "「もう走れない」には2種類ある",
    "【約1分】\n子どもが「もう走れない」と言うとき、原因は大きく2つです。\n1つは、エネルギー切れ。足が重くなって動かなくなる状態です。小学生の数分間の持久走では、まず起こりません。\nもう1つは、酸素不足による息切れ。呼吸と心拍数が上がりすぎて「苦しくてもうダメ」となる状態です。\n小学生の持久走で起きているのは、ほとんどが2つ目です。つまり、ペースが速すぎるのです。ペースを落とせば、また走れるようになります。");
  card(s, 0.6, 2.4, 5.8, 3.6, "1", "エネルギー切れ", "足が重くなり、動かなくなる\n\n小学生の数分間の持久走では、まず起こらない", { fill: TINT, numFill: NAVY, headSize: 26, bodySize: 17, bodyGap: 0.8 });
  card(s, 6.9, 2.4, 5.8, 3.6, "2", "酸素不足（息切れ）", "呼吸・心拍数が上がりすぎて\n「苦しくてもうダメ」となる\n\nペースを落とせば、また走れる", { fill: OTINT, headSize: 26, bodySize: 17, bodyGap: 0.8 });
  T(s, "小学生の「もう走れない」は、ほとんどが「ペースが速すぎる」", { x: 0.6, y: 6.25, w: 12.1, h: 0.55, fontSize: 20, bold: true, color: NAVY });
}
{
  const s = content("4. 「苦しい」の正体", "子どもは本来、\n大人より疲れにくく、回復も早い",
    "【約1分・時間がなければ省略可】\n実は、子どもは大人よりも疲れにくい体を持っています。思春期前の子どもは、酸素を使うエネルギーの仕組みに頼る割合が高く、運動の後の回復も早いことが研究でわかっています。\n休み時間の鬼ごっこを思い出してください。子どもたちは20分間、走ったり止まったりを繰り返しても平気です。\n子どもは本来、長く動ける体を持っています。持久走が苦手に見えるのは、子どもの体の問題ではなく、走らせ方の問題かもしれません。");
  box(s, 0.6, 2.4, 6.2, 4.1, TINT);
  T(s, "研究でわかっていること", { x: 0.95, y: 2.7, w: 5.5, h: 0.5, fontSize: 20, bold: true, color: NAVY });
  T(s, bullets(["酸素を使う仕組みに頼る割合が高い", "大人より疲れにくい", "運動後の回復が早い"]), { x: 0.95, y: 3.4, w: 5.5, h: 2.6, fontSize: 18 });
  box(s, 7.2, 2.4, 5.5, 4.1, NAVY);
  T(s, "たとえば休み時間の", { x: 7.55, y: 2.8, w: 4.8, h: 0.5, fontSize: 18, color: OTINT });
  T(s, "鬼ごっこ", { x: 7.55, y: 3.35, w: 4.8, h: 1.0, fontSize: 44, bold: true, color: WHITE });
  T(s, "20分走り回っても平気。\n苦手にしているのは、\n走らせ方かもしれない", { x: 7.55, y: 4.5, w: 4.8, h: 1.7, fontSize: 18, color: WHITE });
  src(s, "出典：Birat et al. (2018) Frontiers in Physiology ほか");
}

// ================= 5. 3つの方法 =================
section("5. 3つの方法");
{
  const s = content("5. 「持続する力」を高める方法", "持続する力を高める方法は、3つある",
    "【約1分30秒】\n動きを持続する力を高める方法は、大きく分けて3つです。\n1つ目は、運動しても心拍数が上がりにくい体をつくること。心臓や肺を鍛えるトレーニングで、インターバルトレーニングのように、心拍数を思い切り上げては下げることを繰り返します。強い負荷がかかります。\n2つ目は、心拍数が上がっても「苦しい」と感じにくくすること。少し心拍数が上がった状態で、長い時間動き続け、脳を慣れさせます。\n3つ目は、苦しさに耐えて走る力をつけること。タイムトライアルやレペティションのように、きつい状態で走る練習です。");
  const m = [["1", "心拍数を\n上がりにくくする", "心臓・肺を鍛える\n例：インターバル", "強"], ["2", "上がった心拍数に\n脳を慣らす", "ゆっくり長く動き続ける\n例：ジョギング", "弱"], ["3", "苦しさに\n耐えて走る", "きつい状態で走る\n例：タイムトライアル", "強"]];
  m.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    card(s, x, 2.35, 3.85, 4.3, c[0], c[1], c[2], { headSize: 22, bodyGap: 1.2, bodySize: 16, numFill: NAVY });
    box(s, x + 0.3, 5.75, 3.25, 0.6, c[3] === "強" ? OTINT : GTINT);
    T(s, "負荷：" + (c[3] === "強" ? "強い" : "弱い〜中くらい"), { x: x + 0.3, y: 5.75, w: 3.25, h: 0.6, fontSize: 16, bold: true, color: c[3] === "強" ? ORANGE : GREEN, align: "center", valign: "middle" });
  });
}
{
  const s = content("話し合い②　近くの方と3分", "小学生には、どれを選びますか？",
    "【約3分】\n2回目の話し合いです。この3つのうち、小学生にはどれを選ぶか、近くの方と3分間話してください。\n理由も一緒に考えてみてください。「今、自分のクラスでやっている練習はどれに近いか」という視点でもかまいません。\n（3分後）では、選んだ番号を指で示してください。せーの。\n（会場の結果を見て、次のスライドへ）", "WORK");
  const m = [["1", "心拍数を上がりにくくする", "インターバル"], ["2", "上がった心拍数に慣らす", "ゆっくり長く"], ["3", "苦しさに耐える", "タイムトライアル"]];
  m.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    box(s, x, 2.4, 3.85, 2.9, WHITE);
    circle(s, x + 1.375, 2.65, 1.1, c[0], NAVY, 32);
    T(s, c[1], { x: x + 0.2, y: 3.9, w: 3.45, h: 0.6, fontSize: 18, bold: true, color: NAVY, align: "center" });
    T(s, c[2], { x: x + 0.2, y: 4.5, w: 3.45, h: 0.5, fontSize: 16, color: GRAY, align: "center" });
  });
  circle(s, 0.6, 5.65, 1.0, "3分", ORANGE, 22);
  T(s, "理由も一緒に。今、自分のクラスの練習はどれに近いですか？", { x: 1.9, y: 5.65, w: 10.8, h: 1.0, fontSize: 18, color: NAVY, valign: "middle", bold: true });
}
{
  const s = content("5. 「持続する力」を高める方法", "主役は②。\n目安は「おしゃべりできる速さ」",
    "【約1分30秒】\n私のおすすめは②です。少し息が上がるけれど、隣の友達とおしゃべりができるくらいの速さで、長く動き続けることです。これが学習指導要領の「無理のない速さ」にも当たります。\n③の「苦しさに耐える」練習は、効果はあっても「走るのが嫌い」を生みやすい方法です。①も負荷が強く、小学生の段階で積極的に選ぶ必要はありません。\n補足です。大人なら心拍数で強さを管理できますが、子どもは安静時の心拍数も最大心拍数も大人より高く、「220−年齢」という大人向けの式は当てはまりません。だから、心拍数の数字より「おしゃべりできるか」を目安にするのが実用的です。");
  box(s, 0.6, 2.4, 7.4, 4.2, NAVY);
  circle(s, 0.95, 2.75, 1.0, "2", ORANGE, 32);
  T(s, "おしゃべりできる速さで、\nゆっくり長く", { x: 0.95, y: 3.95, w: 6.7, h: 1.4, fontSize: 30, bold: true, color: WHITE });
  T(s, "＝ 学習指導要領の「無理のない速さ」", { x: 0.95, y: 5.55, w: 6.7, h: 0.6, fontSize: 18, color: OTINT, bold: true });
  box(s, 8.4, 2.4, 4.3, 2.0, OTINT);
  T(s, "③は「嫌い」を生みやすい", { x: 8.7, y: 2.6, w: 3.7, h: 0.5, fontSize: 18, bold: true, color: ORANGE });
  T(s, "①も負荷が強く、\n小学生で積極的に選ぶ必要はない", { x: 8.7, y: 3.15, w: 3.7, h: 1.1, fontSize: 15, color: TEXT });
  box(s, 8.4, 4.6, 4.3, 2.0, TINT);
  T(s, "子どもの心拍数は別もの", { x: 8.7, y: 4.8, w: 3.7, h: 0.5, fontSize: 18, bold: true, color: NAVY });
  T(s, "大人向けの「220−年齢」の式は\n当てはまらない", { x: 8.7, y: 5.35, w: 3.7, h: 1.1, fontSize: 15, color: TEXT });
}

// ================= 6. 嫌いになる理由 =================
section("6. 嫌いになる理由");
{
  const s = content("6. 「走るのが嫌い」になる理由", "嫌いになる理由は\n「止まれない」と「目的がない」",
    "【約1分30秒】\n子どもが走ることを嫌いになる理由は、大きく2つあると考えています。\n1つ目は、「苦しくても走り続けなければいけない」こと。止まってはいけないと思うから、苦しさが怖くなります。だから、「きつくなったら、いつでも止まっていい（歩いていい）」と伝えます。止まってもいいとわかると、子どもは安心して走り出せます。\n2つ目は、目的がないこと。ただ同じコースを何周も走るだけでは、つまらないのは当然です。走る目的があれば、子どもの気持ちは変わります。\nでは、具体的にどうするか。ここからが今日の中心です。");
  const r = [["1", "苦しくても、止まれない", "きつくなったら、\nいつでも止まっていい"], ["2", "走る目的がない", "走る目的を\nつくる"]];
  r.forEach((c, i) => {
    const y = 2.4 + i * 2.2;
    box(s, 0.6, y, 5.6, 1.9, TINT);
    circle(s, 0.9, y + 0.5, 0.9, c[0], GRAY, 26);
    T(s, c[1], { x: 2.05, y, w: 4.0, h: 1.9, fontSize: 22, bold: true, color: GRAY, valign: "middle" });
    arrow(s, 6.4, y + 0.7, 0.7, 0.5, ORANGE);
    box(s, 7.3, y, 5.4, 1.9, OTINT);
    T(s, c[2], { x: 7.6, y, w: 4.8, h: 1.9, fontSize: 24, bold: true, color: ORANGE, valign: "middle" });
  });
}

// ================= 7. 大会までに何をするか =================
section("7. 大会までに何をするか");
{
  const s = content("7. 大会までに何をするか", "原則は\n「ゆっくり・短く・選べる・ためる」",
    "【約1分30秒】\nここからは、大会までの期間に何をするかです。私は4つの原則を提案します。\n「ゆっくり」：おしゃべりできる速さで走る。\n「短く」：最初は5分程度から始め、少しずつ伸ばす。\n「選べる」：ペースやコース、歩くかどうかを子ども自身が選べるようにする。\n「ためる」：走った距離や時間を記録して、積み上がっていくのが見えるようにする。\nこの4つがあれば、体力をつけながら、走ることを嫌いにさせずに済みます。");
  const p = [["ゆっくり", "おしゃべりできる\n速さで"], ["短く", "5分程度から始めて\n少しずつ伸ばす"], ["選べる", "ペース・コース・\n歩くかを自分で選ぶ"], ["ためる", "走った距離や時間を\n記録して積み上げる"]];
  p.forEach((c, i) => {
    const x = 0.6 + i * 3.1;
    box(s, x, 2.4, 2.85, 4.1, i % 2 === 0 ? TINT : OTINT);
    circle(s, x + 0.3, 2.7, 0.7, String(i + 1), ORANGE, 20);
    T(s, c[0], { x: x + 0.3, y: 3.65, w: 2.3, h: 0.9, fontSize: 32, bold: true, color: NAVY });
    T(s, c[1], { x: x + 0.3, y: 4.7, w: 2.3, h: 1.5, fontSize: 16, color: TEXT });
  });
}
{
  const s = content("7. 大会までに何をするか　活動例①", "自分の「ちょうどいいペース」を知る",
    "【約1分30秒】\n活動例の1つ目は、自分のちょうどいいペースを知る活動です。\n「おしゃべり走」：2〜3人で、話しながら走れる速さで走ります。話せなくなったら速すぎる合図です。\n「ペア走」：友達と並んで、同じペースで走ります。ペースを合わせることを目標にすると、競争になりません。\n「脈を測る」：走る前と後に、手首や首で15秒間脈を数えて4倍します。自分の体の変化を数字で知ることができ、保健や理科の学習ともつながります。");
  const a = [["おしゃべり走", "2〜3人で、話しながら\n走れる速さで。\n話せなくなったら\n速すぎる合図"], ["ペア走", "友達と並んで、\n同じペースで。\nペースを合わせるのが目標"], ["脈を測る", "走る前と後に\n15秒数えて4倍。\n自分の体の変化がわかる"]];
  a.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    card(s, x, 2.4, 3.85, 4.1, String(i + 1), c[0], c[1], { headSize: 24, bodyGap: 0.75, bodySize: 17 });
  });
}
{
  const s = content("7. 大会までに何をするか　活動例②", "走る目的と、楽しさを足す",
    "【約1分30秒】\n活動例の2つ目は、走る目的と楽しさを足す活動です。\n「ピタリ賞」：走る前に「何分何秒で走れるか」を予想し、実際のタイムとの差が小さい人が勝ち。速い子が勝つとは限らないので、どの子にもチャンスがあります。ペース感覚も身につきます。\n「日本一周マップ」：走った距離をクラスで合計し、地図の上でたどっていきます。「今日は〇〇県に着いた」と目に見えると、目的が生まれます。\n「変化のあるコース」：坂道、芝生、校舎の周りなど。学習指導要領の例示にもある「変化のある地形」です。\n「鬼遊び」：鬼ごっこやしっぽ取りも、立派な「動きを持続する運動」です。");
  const a = [["ピタリ賞", "予想タイムとの差が小さい人が勝ち。\n速くなくても勝てる"], ["日本一周マップ", "走った距離をクラスで合計し、\n地図でたどる"], ["変化のあるコース", "坂道・芝生・校舎の周りなど"], ["鬼遊び", "鬼ごっこ・しっぽ取りも持続する運動"]];
  a.forEach((c, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 2.35 + row * 2.2;
    box(s, x, y, 5.95, 1.95, row === 0 ? OTINT : TINT);
    circle(s, x + 0.3, y + 0.55, 0.85, String(i + 1), ORANGE, 22);
    T(s, c[0], { x: x + 1.4, y: y + 0.3, w: 4.3, h: 0.6, fontSize: 22, bold: true, color: NAVY });
    T(s, c[1], { x: x + 1.4, y: y + 0.95, w: 4.3, h: 0.8, fontSize: 16, color: TEXT });
  });
}
{
  const s = content("7. 大会までに何をするか　大会当日", "順位をつける部門にも、つけない部門にも\n「理由のある目標」を",
    "【約1分30秒】\n皆さんの学校の大会は、順位をつける部門と、順位をつけない部門に分かれていると聞いています。どちらの部門にも、子どもが納得できる「理由のある目標」を持たせることが大切です。\n順位をつける部門の子には、順位だけでなく、自己ベストやペース配分を課題にします。「前半を抑えて、後半に上げる」ことができれば、それは立派な成長です。\n順位をつけない部門の子には、予想タイムとの差、止まらずに走れた時間、自分で決めたペースを守れたか、などを目標にできます。\n大事なのは、ゴールした後に「自分はこれができた」と言えることです。");
  box(s, 0.6, 2.4, 5.8, 4.2, TINT);
  T(s, "順位をつける部門", { x: 0.95, y: 2.7, w: 5.1, h: 0.6, fontSize: 24, bold: true, color: NAVY });
  T(s, bullets(["自己ベストの更新", "ペース配分（後半に上げる）", "「自分の走り」を振り返る"]), { x: 0.95, y: 3.5, w: 5.1, h: 2.8, fontSize: 18 });
  box(s, 6.9, 2.4, 5.8, 4.2, OTINT);
  T(s, "順位をつけない部門", { x: 7.25, y: 2.7, w: 5.1, h: 0.6, fontSize: 24, bold: true, color: ORANGE });
  T(s, bullets(["予想タイムとの差（ピタリ賞）", "止まらずに走れた時間", "自分で決めたペースを守れたか"]), { x: 7.25, y: 3.5, w: 5.1, h: 2.8, fontSize: 18 });
}
{
  const s = content("話し合い③　グループで5分", "順位をつけない部門の子が\n「走ってよかった」と思えるには？",
    "【約5分＋共有2分】\n3回目、最後の話し合いです。3〜4人のグループで5分間話してください。\nテーマは「順位をつけない部門の子が、大会の後に『走ってよかった』と思えるには、どうしたらよいか」です。\n大会までの練習、当日の仕組み、終わった後の振り返り、どこからでもかまいません。担任以外の先生方も、ぜひご自分の立場からアイデアを出してください。\n（5分後、2〜3グループに発表してもらう。出たアイデアをまとめて、次のスライドへ）", "WORK");
  const sc = [["大会までの練習", "どんな活動・声かけ？"], ["大会当日", "どんな仕組み・目標？"], ["大会の後", "どんな振り返り？"]];
  sc.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    box(s, x, 2.5, 3.85, 2.6, WHITE);
    T(s, c[0], { x: x + 0.3, y: 2.8, w: 3.25, h: 0.6, fontSize: 22, bold: true, color: ORANGE });
    T(s, c[1], { x: x + 0.3, y: 3.6, w: 3.25, h: 1.2, fontSize: 18, color: NAVY });
  });
  circle(s, 0.6, 5.55, 1.0, "5分", ORANGE, 22);
  T(s, "3〜4人のグループで。担任以外の先生方も、ご自分の立場からアイデアをお願いします", { x: 1.9, y: 5.55, w: 10.8, h: 1.0, fontSize: 18, color: NAVY, valign: "middle", bold: true });
}
{
  const s = content("7. 大会までに何をするか　声かけ", "声かけは\n「がんばれ」より「いいペースだね」",
    "【約1分・時間がなければ省略可】\nアイデアをありがとうございました。最後に、全員がすぐに使えることを1つ。声かけです。\n「がんばれ」「抜かれるな」「止まるな」という声かけは、子どもに「もっと苦しくなれ」と伝えてしまうことがあります。\n代わりに、「いいペースだね」「さっきより楽そうだね」「自分で決めたペースを守れてるよ」と、子どもが自分の走りに気づける言葉をかけてみてください。\nこれは担任でなくても、コースに立つ先生全員ができることです。");
  box(s, 0.6, 2.4, 5.6, 4.1, TINT);
  T(s, "よく聞く声かけ", { x: 0.95, y: 2.7, w: 4.9, h: 0.5, fontSize: 18, bold: true, color: GRAY });
  T(s, "「がんばれ！」\n「抜かれるな！」\n「止まるな！」", { x: 0.95, y: 3.35, w: 4.9, h: 2.8, fontSize: 26, bold: true, color: GRAY });
  arrow(s, 6.35, 4.2, 0.6, 0.5, ORANGE);
  box(s, 7.1, 2.4, 5.6, 4.1, OTINT);
  T(s, "自分の走りに気づける声かけ", { x: 7.45, y: 2.7, w: 4.9, h: 0.5, fontSize: 18, bold: true, color: ORANGE });
  T(s, "「いいペースだね」\n「さっきより楽そうだね」\n「自分で決めたペースを\n　守れてるよ」", { x: 7.45, y: 3.35, w: 4.9, h: 2.9, fontSize: 24, bold: true, color: NAVY });
}

// ================= 8. 安全 =================
section("8. 安全管理");
{
  const s = content("8. 安全管理", "学校での心臓突然死は毎年10人前後。\n約3分の2は運動に関連して起きている",
    "【約1分30秒】\n最後に、必ずお伝えしたいのが安全のことです。\n学校の管理下での心臓突然死は、毎年10人前後起きています。そして、その約3分の2は運動に関連して起きています。\n持久走大会での事例もあります。日本スポーツ振興センターの資料には、2kmの持久走大会でゴール直後に意識を失い、翌日亡くなった児童の事例が載っています。この児童は持病があったものの、学校生活管理指導表では「強い運動も可」とされていました。\n心臓検診で異常がなかった子にも起こりえます。「苦しい」は脳のブレーキだとお話ししましたが、胸の痛み、めまい、顔色の悪さなどは本当の危険サインです。見逃さないでください。");
  T(s, [{ text: "10", options: { fontSize: 110, bold: true, color: ORANGE, fontFace: NUM } }, { text: " 人前後", options: { fontSize: 32, bold: true, color: NAVY, fontFace: HEAD } }], { x: 0.6, y: 2.4, w: 5.8, h: 1.8, valign: "middle" });
  T(s, "学校の管理下での心臓突然死（毎年）", { x: 0.6, y: 4.3, w: 5.8, h: 0.6, fontSize: 18, bold: true, color: NAVY });
  T(s, [{ text: "約 ", options: { fontSize: 40, bold: true, color: ORANGE, fontFace: HEAD } }, { text: "2/3", options: { fontSize: 110, bold: true, color: ORANGE, fontFace: NUM } }], { x: 6.9, y: 2.4, w: 5.8, h: 1.8, valign: "middle" });
  T(s, "そのうち運動に関連して起きた割合", { x: 6.9, y: 4.3, w: 5.8, h: 0.6, fontSize: 18, bold: true, color: NAVY });
  box(s, 0.6, 5.2, 12.1, 1.5, OTINT);
  T(s, "本当の危険サイン：胸の痛み・めまい・顔色の悪さ・急に座り込む", { x: 0.95, y: 5.2, w: 11.4, h: 0.75, fontSize: 20, bold: true, color: ORANGE, valign: "middle" });
  T(s, "心臓検診で異常がなかった子にも起こりうる", { x: 0.95, y: 5.9, w: 11.4, h: 0.6, fontSize: 16, color: TEXT, valign: "middle" });
  src(s, "出典：日本スポーツ振興センター「学校の管理下の災害」、山形県「学校における心臓突然死の防止」資料");
}
{
  const s = content("8. 安全管理", "ゴール前とゴール直後は、\n全職員で見守る",
    "【約1分30秒】\n特に注意が必要なのは、ゴール前のスパートと、ゴールした直後です。急に負荷が上がり、急に止まるタイミングで倒れる事例が多くあります。\nそこで、大会では次のことを確認しておいてください。\n・AEDの位置。コースのどこからでも数分で届く配置になっているか。\n・ゴール付近に、子どもの様子を見る担当の先生がいるか。\n・ゴール後に座り込んだ子や、様子がおかしい子に必ず声をかける。\n・事前に、心臓検診の結果や学校生活管理指導表を確認しておく。\nそして、倒れた子がいたときに誰が何をするかという緊急時の対応計画を、全職員で共有しておくことです。今日ここにいる全員が、その担い手です。");
  // course line
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 3.0, w: 11.5, h: 0, line: { color: GRAY, width: 4 } });
  T(s, "スタート", { x: 0.6, y: 2.25, w: 1.6, h: 0.5, fontSize: 16, bold: true, color: GRAY });
  circle(s, 0.75, 2.75, 0.5, "S", GRAY, 14);
  box(s, 7.6, 2.55, 2.6, 0.9, OTINT);
  T(s, "ゴール前\nスパート", { x: 7.6, y: 2.55, w: 2.6, h: 0.9, fontSize: 15, bold: true, color: ORANGE, align: "center", valign: "middle" });
  circle(s, 10.6, 2.6, 0.8, "G", ORANGE, 20);
  T(s, "ゴール直後", { x: 10.0, y: 3.5, w: 2.7, h: 0.5, fontSize: 16, bold: true, color: ORANGE, align: "center" });
  box(s, 0.6, 4.2, 12.1, 2.5, TINT);
  T(s, "大会前に確認すること", { x: 0.95, y: 4.4, w: 5, h: 0.5, fontSize: 18, bold: true, color: NAVY });
  T(s, bullets(["AEDの位置（どこからでも数分で届くか）", "ゴール付近に様子を見る担当を置く"]), { x: 0.95, y: 5.0, w: 5.6, h: 1.5, fontSize: 16 });
  T(s, bullets(["座り込んだ子には必ず声をかける", "心臓検診・管理指導表の確認", "緊急時の対応計画の共有"]), { x: 6.8, y: 5.0, w: 5.7, h: 1.5, fontSize: 16 });
}

// ================= まとめ =================
section("まとめ");
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: sectionName });
  s.addText("育てたいのは\n「理由を持って、\nまた走りたくなる子」", { placeholder: "title" });
  s.addText("「耐える子」ではなく。\nゆっくり・短く・選べる・ためる　／　声かけは「いいペースだね」　／　安全は全職員で", { placeholder: "body" });
  s.addNotes("【約1分30秒】\n今日のまとめです。\n小学校の持久走は「無理のない速さ」で行う運動で、マラソン大会は「運動に親しむ態度」を育てる行事です。\n「苦しい」は脳のブレーキで、ゆっくり長く走ることで感じ方を変えていけます。\n大会までは「ゆっくり・短く・選べる・ためる」。声かけは「がんばれ」より「いいペースだね」。そして安全は全職員で守る。\n私が中学生の選手たちと大切にしてきたのは、納得できる、理由のある努力でした。小学生も同じです。育てたいのは「耐える子」ではなく、「理由を持って、また走りたくなる子」です。\nご清聴ありがとうございました。");
}
{
  const s = content("参考資料", "参考資料",
    "（投影しなくてもよい。配布資料用）");
  T(s, bullets([
    "文部科学省『小学校学習指導要領（平成29年告示）』『同 解説 体育編』『同 解説 特別活動編』",
    "文部科学省『中学校学習指導要領（平成29年告示）』",
    "スポーツ庁『スポーツの実施状況等に関する世論調査』（令和6年度）",
    "日本スポーツ振興センター『学校の管理下の災害』、学校安全Web",
    "山形県教育委員会『学校における心臓突然死の防止』",
    "Birat A. et al. (2018) Metabolic and Fatigue Profiles Are Comparable Between Prepubertal Children and Well-Trained Adult Endurance Athletes. Frontiers in Physiology",
    "Verschuren O. et al. (2011) The 220-age equation does not predict maximum heart rate in children and adolescents. Dev Med Child Neurol",
  ]), { x: 0.6, y: 2.3, w: 12.1, h: 4.4, fontSize: 15 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote " + OUT);
})();
