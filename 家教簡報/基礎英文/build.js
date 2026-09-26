const pptxgen = require("pptxgenjs");
const path = require("path");

const OUT = process.argv[2] || path.join(__dirname, "基礎英文_Ch01-04_上課簡報.pptx");

const F = "Microsoft JhengHei";
const NAVY = "1F3A5F";
const NAVY2 = "2E5484";
const RED = "B8312F";
const INK = "1F2933";
const MUTED = "5F6B7A";
const TINT = "EEF2F8";
const REDTINT = "FBECEB";
const LINE = "CBD3DF";
const WHITE = "FFFFFF";
const PAPER = "F7F8FA";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "基礎英文 Chapter 01–04 上課簡報";
pres.lang = "zh-TW";

// ---------- helpers ----------
const R = (text, o = {}) => ({ text, options: { breakLine: true, ...o } });
function runs(arr) {
  const a = arr.flat().filter(Boolean);
  if (a.length) a[a.length - 1] = { ...a[a.length - 1], options: { ...a[a.length - 1].options, breakLine: false } };
  return a;
}
// English + Chinese example line
const EX = (en, zh, o = {}) => [
  { text: en, options: { bold: true, color: NAVY, breakLine: false, ...o } },
  { text: zh ? "　" + zh : "", options: { color: MUTED, fontSize: (o.fontSize || 13) - 1, breakLine: true } },
];
const H = (t, o = {}) => R(t, { bold: true, color: RED, ...o });
const P = (t, o = {}) => R(t, { color: INK, ...o });
const M = (t, o = {}) => R(t, { color: MUTED, ...o });
const GAP = (n = 0.5) => R("", { fontSize: 6 * n * 2 });

function text(s, arr, box) {
  s.addText(runs(arr), {
    fontFace: F, fontSize: 13, color: INK, margin: 0, isTextBox: true, valign: "top",
    paraSpaceAfter: 2, ...box,
  });
}

function pill(s, t, x, y, w, h, fill, color = WHITE, fs = 11) {
  s.addText(t, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: h / 2, x, y, w, h,
    fill: { color: fill }, line: { color: fill },
    fontFace: F, fontSize: fs, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
}

function base(title, page, notes) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText(title, {
    x: 0.5, y: 0.28, w: 7.7, h: 0.62, fontFace: F, fontSize: 24, bold: true, color: NAVY,
    margin: 0, isTextBox: true, valign: "middle",
  });
  if (page) pill(s, `課本 p.${page}`, 8.25, 0.42, 1.25, 0.34, NAVY, WHITE, 11);
  if (notes) s.addNotes(notes);
  return s;
}

function card(s, { x, y, w, h, fill = TINT, label, labelFill = NAVY, body, fs = 13, labelW = 0.9 }) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.1 });
  let ty = y + 0.18;
  if (label) {
    pill(s, label, x + 0.2, y + 0.16, labelW, 0.28, labelFill, WHITE, 10);
    ty = y + 0.55;
  }
  if (body) text(s, body, { x: x + 0.2, y: ty, w: w - 0.4, h: y + h - ty - 0.12, fontSize: fs });
}

// sentence-pattern formula: boxes joined by +
function formula(s, parts, x, y, w, h = 0.5, opts = {}) {
  const plusW = 0.28;
  const n = parts.length;
  const bw = (w - plusW * (n - 1)) / n;
  parts.forEach((p, i) => {
    const px = x + i * (bw + plusW);
    const isQ = /？|\?$/.test(p);
    s.addText(p, {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: px, y, w: bw, h,
      fill: { color: opts.fill || NAVY }, line: { color: opts.fill || NAVY },
      fontFace: F, fontSize: opts.fs || 14, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    if (i < n - 1) {
      s.addText("＋", {
        x: px + bw, y, w: plusW, h, fontFace: F, fontSize: 16, bold: true, color: MUTED,
        align: "center", valign: "middle", margin: 0, isTextBox: true,
      });
    }
  });
}

function table(s, rows, { x, y, w, colW, fs = 12, rowH, headFill = NAVY, align }) {
  const data = rows.map((r, ri) =>
    r.map((c, ci) => {
      const isHead = ri === 0;
      const cell = typeof c === "object" && c !== null && !Array.isArray(c) ? c : { text: String(c) };
      return {
        text: cell.text,
        options: {
          fontFace: F, fontSize: fs, margin: [2, 5, 2, 5],
          bold: isHead || !!cell.bold,
          color: isHead ? WHITE : cell.color || INK,
          fill: { color: isHead ? headFill : cell.fill || (ri % 2 === 0 ? PAPER : WHITE) },
          align: cell.align || (align ? align[ci] : "left"),
          valign: "middle",
          border: { type: "solid", pt: 0.75, color: LINE },
        },
      };
    })
  );
  s.addTable(data, { x, y, w, colW, rowH, autoPage: false });
}

function chapterSlide(num, en, zh, topics, pages, notes) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText(`Chapter ${num}`, { x: 0.7, y: 0.9, w: 5, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: "9FB4D6", margin: 0, isTextBox: true });
  s.addText(en, { x: 0.7, y: 1.4, w: 8.6, h: 1.0, fontFace: F, fontSize: en.length > 20 ? 30 : 40, bold: true, color: WHITE, margin: 0, isTextBox: true, valign: "middle" });
  s.addText(zh, { x: 0.7, y: 2.4, w: 8.6, h: 0.5, fontFace: F, fontSize: 20, color: "D6E0F0", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 3.15, w: 8.6, h: 1.75, fill: { color: NAVY2 }, line: { color: NAVY2 }, rectRadius: 0.1 });
  s.addText("這一章要學", { x: 0.95, y: 3.3, w: 3, h: 0.35, fontFace: F, fontSize: 12, bold: true, color: "9FB4D6", margin: 0, isTextBox: true });
  const half = Math.ceil(topics.length / 2);
  const col = (items, x) =>
    s.addText(runs(items.map((t) => R(t, { color: WHITE, bullet: { code: "25CF" } }))), {
      x, y: 3.68, w: 4.0, h: 1.15, fontFace: F, fontSize: 15, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 4,
    });
  col(topics.slice(0, half), 0.95);
  if (topics.length > half) col(topics.slice(half), 5.1);
  pill(s, `課本 p.${pages}`, 8.0, 0.95, 1.4, 0.34, RED, WHITE, 11);
  if (notes) s.addNotes(notes);
  return s;
}

function practiceSlide(chapter, items, answers, notes) {
  // question slide
  const mk = (showAns) => {
    const s = base(`你來試試 · Chapter ${chapter}`, null, showAns ? "對答案。答錯的地方回到前面的表格再看一次，不要直接跳過。" : notes);
    pill(s, showAns ? "解答" : "練習", 8.25, 0.42, 1.25, 0.34, RED, WHITE, 11);
    const colW = 4.35;
    const half = Math.ceil(items.length / 2);
    [items.slice(0, half), items.slice(half)].forEach((chunk, ci) => {
      const offset = ci === 0 ? 0 : half;
      const body = [];
      chunk.forEach((it, i) => {
        body.push(R(`${offset + i + 1}. ${it.q}`, { bold: true, color: NAVY, fontSize: 15 }));
        it.lines.forEach((ln, li) => {
          if (showAns) {
            const ans = answers[offset + i][li];
            body.push([
              { text: ln.replace("___", "") === ln ? ln : ln.split("___")[0], options: { color: INK, breakLine: false } },
              { text: ln.includes("___") ? ans : "　→ " + ans, options: { color: RED, bold: true, breakLine: false } },
              { text: ln.includes("___") ? ln.split("___")[1] : "", options: { color: INK, breakLine: true } },
            ]);
          } else body.push(P(ln));
        });
        body.push(GAP(0.4));
      });
      text(s, body, { x: 0.5 + ci * (colW + 0.3), y: 1.15, w: colW, h: 4.1, fontSize: 15, paraSpaceAfter: 5 });
    });
    return s;
  };
  mk(false);
  mk(true);
}

// ---------- vocab table helper ----------
function vocab(s, list, { x, y, w, cols = 2, fs = 12, rowH = 0.3 }) {
  // list: [[en, pos, zh], ...] -> split into cols side by side tables
  const per = Math.ceil(list.length / cols);
  const cw = (w - 0.2 * (cols - 1)) / cols;
  for (let c = 0; c < cols; c++) {
    const chunk = list.slice(c * per, (c + 1) * per);
    if (!chunk.length) continue;
    const rows = [["單字", "中譯"], ...chunk.map(([en, zh]) => [{ text: en, bold: true, color: NAVY }, zh])];
    table(s, rows, { x: x + c * (cw + 0.2), y, w: cw, colW: [cw * 0.42, cw * 0.58], fs, rowH });
  }
}

// =====================================================================
// 0. Title + agenda
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("基礎英文", { x: 0.7, y: 1.3, w: 8.6, h: 1.1, fontFace: F, fontSize: 48, bold: true, color: WHITE, margin: 0, isTextBox: true, valign: "middle" });
  s.addText("Chapter 01 – 04　上課簡報", { x: 0.7, y: 2.45, w: 8.6, h: 0.6, fontFace: F, fontSize: 24, color: "D6E0F0", margin: 0, isTextBox: true });
  s.addText("人稱代名詞 · be 動詞 · 疑問詞 · 指示代名詞 · 介係詞", { x: 0.7, y: 3.2, w: 8.6, h: 0.5, fontFace: F, fontSize: 15, color: "9FB4D6", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 4.2, w: 8.6, h: 0.75, fill: { color: NAVY2 }, line: { color: NAVY2 }, rectRadius: 0.1 });
  s.addText("課本翻到 p.33，跟著簡報右上角的頁碼走", { x: 0.95, y: 4.2, w: 8.2, h: 0.75, fontFace: F, fontSize: 15, color: WHITE, margin: 0, isTextBox: true, valign: "middle" });
  s.addNotes("開場：今天從 Chapter 1 開始，每一頁簡報右上角都有課本頁碼，讓學生同步翻課本。先講今天大概會上到哪裡（看進度，至少 Chapter 1 全部）。");
}
{
  const s = base("今天的路線圖", null, "四章的主線：第一章學「A 是 B」，第二、三章學怎麼「問問題」，第四章學怎麼說「在哪裡」。跟學生說：這四章其實都在用同一個東西，就是 be 動詞。");
  const boxes = [
    ["01", "Hi, I am Sooji.", "主格代名詞、be 動詞\n直述／疑問／否定句", "p.33–39"],
    ["02", "Who is this?", "who、how\n所有格、and / but", "p.43–49"],
    ["03", "What is this?", "what、this / that\n受格", "p.53–59"],
    ["04", "Where is the laundry basket?", "where、場所副詞\n介係詞 at / in / on…", "p.63–69"],
  ];
  boxes.forEach((b, i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.15 + Math.floor(i / 2) * 2.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.4, h: 1.85, fill: { color: TINT }, line: { color: TINT }, rectRadius: 0.1 });
    s.addText(b[0], { x: x + 0.2, y: y + 0.15, w: 0.9, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: RED, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(b[1], { x: x + 1.2, y: y + 0.18, w: 3.1, h: 0.45, fontFace: F, fontSize: b[1].length > 20 ? 12 : 15, bold: true, color: NAVY, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(b[2], { x: x + 1.2, y: y + 0.65, w: 3.0, h: 0.75, fontFace: F, fontSize: 12, color: INK, margin: 0, isTextBox: true, valign: "top" });
    s.addText(b[3], { x: x + 1.2, y: y + 1.42, w: 3.0, h: 0.3, fontFace: F, fontSize: 11, color: MUTED, margin: 0, isTextBox: true });
  });
}

// =====================================================================
// CHAPTER 01
// =====================================================================
chapterSlide("01", "Hi, I am Sooji.", "你好，我是守智。",
  ["人稱代名詞：主格", "be 動詞：是～", "be 動詞的直述句和疑問句", "be 動詞的否定句"], "33",
  "先讓學生唸一次標題 Hi, I am Sooji.。這一章的核心只有一句：「誰 ＋ 是 ＋ 什麼」，其他都是這句的變形。");

// 1.1 主格
{
  const s = base("人稱代名詞：主格", 34,
    "主格 = 當「主詞」用的代名詞，放在句子最前面，就是「誰」。先帶學生念完整張表，重點放在單數和複數對照：I → we、he/she/it → they。第二人稱 you 單複數長一樣，要看上下文。");
  text(s, [P("放在句子最前面、當「主詞」用的代名詞，依人稱和數量會有不同型態")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  table(s, [
    ["人稱", "單數", "中譯", "複數", "中譯"],
    ["第 1 人稱", { text: "I", bold: true, color: NAVY }, "我", { text: "we", bold: true, color: NAVY }, "我們"],
    ["第 2 人稱", { text: "you", bold: true, color: NAVY }, "你", { text: "you", bold: true, color: NAVY }, "你們"],
    ["第 3 人稱", { text: "he / she / it", bold: true, color: NAVY }, "他 / 她 / 它", { text: "they", bold: true, color: NAVY }, "他們"],
  ], { x: 0.5, y: 1.5, w: 5.6, colW: [1.2, 1.3, 1.2, 0.9, 1.0], fs: 13, rowH: 0.45 });
  card(s, { x: 6.4, y: 1.5, w: 3.1, h: 1.85, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("第二人稱 you 的單複數同形，要看上下文判斷是「你」還是「你們」"),
  ], fs: 12 });
  card(s, { x: 0.5, y: 3.65, w: 9.0, h: 1.4, label: "怎麼記", body: [
    P("第一人稱 = 說話的人（我、我們）　第二人稱 = 聽話的人（你、你們）　第三人稱 = 其他人／東西"),
    P("he 男生、she 女生、it 東西或動物；三個人以上不管男女全部用 they"),
  ], fs: 12 });
}

// 1.2 be 動詞
{
  const s = base("be 動詞：是～", 34,
    "be 動詞就是「是」。am / are / is 三個長得不一樣，但都是同一個字，看主詞決定用哪個。口訣：I 配 am，you / we / they 配 are，he / she / it 配 is。補語 = be 動詞後面接的東西，講「主詞是什麼、怎麼樣」。縮寫要讓學生自己唸一次。");
  text(s, [P("表「是～」，放在主詞後面。後面接的字叫「補語」：名字、特徵、外貌、國籍、職業、狀態⋯⋯")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  table(s, [
    ["", "主詞", "be 動詞", "補語", "中譯"],
    ["單數", "I", { text: "am", bold: true, color: RED }, "Alice.（名字）", "我是愛麗絲"],
    ["", "You", { text: "are", bold: true, color: RED }, "smart.（特徵）", "你很聰明"],
    ["", "He / She / It", { text: "is", bold: true, color: RED }, "cute.（外貌）", "他／她／它很可愛"],
    ["複數", "We", { text: "are", bold: true, color: RED }, "Korean.（國籍）", "我們是韓國人"],
    ["", "You", { text: "are", bold: true, color: RED }, "students.（身分）", "你們是學生"],
    ["", "They", { text: "are", bold: true, color: RED }, "hungry.（狀態）", "他們餓了"],
  ], { x: 0.5, y: 1.45, w: 6.1, colW: [0.65, 1.35, 0.85, 1.95, 1.3], fs: 12, rowH: 0.36 });
  card(s, { x: 6.8, y: 1.45, w: 2.7, h: 2.95, fill: REDTINT, label: "縮寫", labelFill: RED, body: [
    P("I am → I'm"), P("you are → you're"), P("he is → he's"), P("she is → she's"),
    P("it is → it's"), P("we are → we're"), P("they are → they're"),
    M("口說時幾乎都用縮寫", { fontSize: 11 }),
  ], fs: 13 });
  text(s, [[
    { text: "口訣：", options: { bold: true, color: RED, breakLine: false } },
    { text: "I → am　　he / she / it → is　　其他（you / we / they）→ are", options: { bold: true, color: NAVY, breakLine: false } },
  ]], { x: 0.5, y: 4.65, w: 9.0, h: 0.5, fontSize: 14 });
}

// 1.3 直述句 / 疑問句
{
  const s = base("be 動詞的直述句和疑問句", 35,
    "直述句就是「陳述一件事」，句尾句點。變成問句只要做一件事：把 be 動詞搬到最前面，句尾改問號，什麼字都不用加。讓學生把 She is pretty. 自己改成問句。回答用 Yes / No 開頭。");
  text(s, [P("直述句 → 疑問句：把 be 動詞搬到句首，句尾改成問號，其他不變")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  text(s, [H("直述句")], { x: 0.5, y: 1.45, w: 2, h: 0.35, fontSize: 13 });
  formula(s, ["主詞", "be 動詞", "補語 ."], 0.5, 1.8, 4.3, 0.48);
  text(s, [EX("She is pretty.", "她很漂亮"), EX("They are students.", "他們是學生")], { x: 0.5, y: 2.4, w: 4.3, h: 0.8, fontSize: 14 });
  text(s, [H("疑問句")], { x: 5.2, y: 1.45, w: 2, h: 0.35, fontSize: 13 });
  formula(s, ["be 動詞", "主詞", "補語 ?"], 5.2, 1.8, 4.3, 0.48, { fill: RED });
  text(s, [EX("Is she pretty?", "她漂亮嗎？"), EX("Are they students?", "他們是學生嗎？")], { x: 5.2, y: 2.4, w: 4.3, h: 0.8, fontSize: 14 });
  card(s, { x: 0.5, y: 3.35, w: 4.3, h: 1.75, label: "重點", body: [
    P("She is. 這樣不算完整的句子，be 動詞後面一定要接補語（名字、特徵、外貌、國籍、身分、職業、狀態）"),
  ], fs: 12 });
  card(s, { x: 5.2, y: 3.35, w: 4.3, h: 1.75, fill: REDTINT, label: "回答", labelFill: RED, body: [
    EX("Are you Taiwanese?", "你是臺灣人嗎？", { fontSize: 12 }),
    EX("Yes, I'm Taiwanese.", "肯定", { fontSize: 12 }),
    EX("No, I'm Japanese.", "否定", { fontSize: 12 }),
  ], fs: 12 });
}

// 1.4 否定句
{
  const s = base("be 動詞的否定句", 35,
    "否定句更簡單：be 動詞後面加 not，就結束了。縮寫有兩種，都可以，只有 am not 不能縮成 amn't。讓學生把 He is a doctor. 改成否定句，然後兩種縮寫各講一次。");
  text(s, [P("在 be 動詞後面加上 not 就是否定句")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  text(s, [H("肯定句")], { x: 0.5, y: 1.45, w: 2, h: 0.35, fontSize: 13 });
  formula(s, ["主詞", "be 動詞", "補語 ."], 0.5, 1.8, 4.3, 0.48);
  text(s, [EX("It is small.", "這個很小"), EX("I am a singer.", "我是歌手")], { x: 0.5, y: 2.4, w: 4.3, h: 0.8, fontSize: 14 });
  text(s, [H("否定句")], { x: 5.2, y: 1.45, w: 2, h: 0.35, fontSize: 13 });
  formula(s, ["主詞", "be 動詞", "not", "補語 ."], 5.2, 1.8, 4.3, 0.48, { fill: RED });
  text(s, [EX("It is not small.", "這個不小"), EX("I am not a singer.", "我不是歌手")], { x: 5.2, y: 2.4, w: 4.3, h: 0.8, fontSize: 14 });
  card(s, { x: 0.5, y: 3.3, w: 9.0, h: 1.9, fill: REDTINT, label: "兩種縮寫", labelFill: RED, labelW: 1.1, body: [] });
  table(s, [
    ["原句", "縮寫 ①", "縮寫 ②"],
    ["I am not", "I'm not", { text: "（沒有 amn't）", color: MUTED }],
    ["you are not", "you're not", "you aren't"],
    ["he / she / it is not", "he's not / she's not / it's not", "he isn't / she isn't / it isn't"],
    ["we / they are not", "we're not / they're not", "we aren't / they aren't"],
  ], { x: 0.7, y: 3.7, w: 8.6, colW: [2.0, 3.3, 3.3], fs: 11, rowH: 0.26 });
}

// 練習 Ch1
practiceSlide("01", [
  { q: "填入 am / is / are", lines: ["I ___ a student.", "She ___ pretty.", "They ___ hungry.", "It ___ small."] },
  { q: "改成疑問句", lines: ["He is a doctor.", "They are students."] },
  { q: "改成否定句（用縮寫）", lines: ["She is Korean.", "I am a singer."] },
  { q: "回答自己", lines: ["Where are you from?"] },
], [
  ["am", "is", "are", "is"],
  ["Is he a doctor?", "Are they students?"],
  ["She isn't Korean. / She's not Korean.", "I'm not a singer."],
  ["I'm from Taiwan. / I'm Taiwanese."],
], "讓學生口頭回答就好，不用寫。第 1 題卡住就回去看口訣：I → am、he/she/it → is、其他 → are。");

// 對話 1
{
  const s = base("對話 1　Nice to meet you.", 36,
    "先自己唸一遍給學生聽，再讓學生扮演 Sooji、你扮演 Justin，之後互換。重點三句：Nice to meet you. / How about you? / I'm a doctor.（職業前面要加 a）。too 放句尾。");
  const dlg = [
    ["Sooji", "Hi, I'm Sooji.", "你好！我是守智。"],
    ["Justin", "Hi, I'm Justin.", "你好！我是賈斯汀。"],
    ["Sooji", "Nice to meet you, Justin.", "很高興認識你，賈斯汀。"],
    ["Justin", "Nice to meet you, too. Are you a student, Sooji?", "我也很高興認識你。守智，你是學生嗎？"],
    ["Sooji", "Yes. How about you?", "是的。你呢？"],
    ["Justin", "I'm a doctor.", "我是醫生。"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 14 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 4.1, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 1.9, label: "新用法", body: [
    EX("Nice to meet you.", "很高興認識你", { fontSize: 12 }),
    EX("How about you?", "你呢？", { fontSize: 12 }),
    EX("Yes.", "是的", { fontSize: 12 }),
    M("Nice 可換成 Great / Happy；How about you? 可換成 What about you?", { fontSize: 10 }),
  ], fs: 12 });
  card(s, { x: 6.1, y: 3.1, w: 3.4, h: 2.05, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("說職業前面要加 a / an：I'm a doctor.", { fontSize: 12 }),
    P("too（也）放在句尾：I am a student, too.", { fontSize: 12 }),
    M("單字：nice 美好的 / meet 認識 / too 也 / student 學生 / doctor 醫生", { fontSize: 10 }),
  ], fs: 12 });
}

// 對話 2
{
  const s = base("對話 2　Where are you from?", 37,
    "重點句型：Where are you from? 回答兩種：I'm from ＋ 國家 或 I'm ＋ 國籍。No, I'm not. 是 No, I'm not American. 省略重複的字。國名和國籍第一個字母一定大寫。可以馬上問學生：Where are you from?");
  const dlg = [
    ["Elise", "Hi, I'm Elise.", "你好，我是伊莉絲。"],
    ["Carlo", "Nice to meet you, Elise. Are you American?", "很高興認識你，伊莉絲。你是美國人嗎？"],
    ["Elise", "Yes, I'm American. Where are you from?", "是的，我是美國人。你來自哪裡？"],
    ["Carlo", "I'm from Italy. Justin, are you American, too?", "我來自義大利。賈斯汀，你也是美國人嗎？"],
    ["Justin", "No, I'm not. I'm Canadian.", "不，我不是。我是加拿大人。"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 14 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 3.0, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 0.5, y: 4.0, w: 5.4, h: 1.15, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("國名、國籍第一個字母大寫：Italy / Italian、Japan / Japanese", { fontSize: 12 }),
    P("No, I'm not. = No, I'm not American.　重複的字省略", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 4.1, label: "問國籍", body: [
    EX("Where are you from?", "你來自哪裡？", { fontSize: 12 }),
    GAP(0.3),
    H("回答 ①　I'm from ＋ 國家", { fontSize: 12 }),
    EX("I'm from Korea.", "", { fontSize: 12 }),
    EX("I'm from Canada.", "", { fontSize: 12 }),
    GAP(0.3),
    H("回答 ②　I'm ＋ 國籍", { fontSize: 12 }),
    EX("I'm Korean.", "", { fontSize: 12 }),
    EX("I'm Canadian.", "", { fontSize: 12 }),
    GAP(0.3),
    M("where 哪裡 / from 來自 / Italy 義大利", { fontSize: 10 }),
  ], fs: 12 });
}

// 單字補給站：國名 / 國籍 / 職業
{
  const s = base("單字補給站　國名、國籍、職業", 38,
    "國籍的字尾有三種：-an（Korean, American）、-ese（Japanese, Chinese）、-ish/其他（French, German）。不用背規則，先會唸。職業帶過一次就好，Point 那句 I work for a company. 可以教學生說自己是上班族／學生。");
  table(s, [
    ["地區", "國名", "國籍"],
    ["亞洲", "Korea 韓國", "Korean 韓國人"],
    ["", "Japan 日本", "Japanese 日本人"],
    ["", "China 中國", "Chinese 中國人"],
    ["歐洲", "France 法國", "French 法國人"],
    ["", "Germany 德國", "German 德國人"],
    ["美洲", "the U.S.A. 美國", "American 美國人"],
    ["", "Canada 加拿大", "Canadian 加拿大人"],
    ["大洋洲／非洲", "Australia 澳洲", "Australian 澳洲人"],
    ["", "South Africa 南非", "South African 南非人"],
  ], { x: 0.5, y: 1.05, w: 5.0, colW: [1.3, 1.85, 1.85], fs: 11, rowH: 0.3 });
  text(s, [H("職業", { fontSize: 13 })], { x: 5.8, y: 1.05, w: 3.7, h: 0.3 });
  table(s, [
    ["單字", "中譯", "單字", "中譯"],
    ["cook", "廚師", "mail carrier", "郵差"],
    ["farmer", "農夫", "police officer", "警察"],
    ["painter", "畫家", "firefighter", "消防員"],
    ["businessman", "商人", "civil servant", "公務員"],
  ], { x: 5.8, y: 1.4, w: 3.7, colW: [1.2, 0.6, 1.3, 0.6], fs: 10.5, rowH: 0.3 });
  card(s, { x: 5.8, y: 3.35, w: 3.7, h: 1.8, label: "重點", body: [
    P("說自己是上班族，比 I'm a civil servant. 更自然的講法：", { fontSize: 11 }),
    EX("I work for a company.", "我是上班族", { fontSize: 12 }),
    EX("I work for the government.", "我是公務員", { fontSize: 12 }),
  ], fs: 11 });
}

// 實用表達法：打招呼 / 再見
{
  const s = base("實用表達法　見面和道別", 39,
    "Good morning / afternoon / evening 是見面用的；Good night 只用在道別或睡前。朋友之間直接 Hi、Hello。Take care 是「保重」，可以回 You too.。這頁可以直接跟學生練：你說 Good morning，他回。");
  card(s, { x: 0.5, y: 1.05, w: 4.3, h: 2.4, label: "見面", body: [
    EX("Good morning.", "早安（早上）", { fontSize: 13 }),
    EX("Good afternoon.", "午安（下午）", { fontSize: 13 }),
    EX("Good evening.", "晚上好（傍晚見面）", { fontSize: 13 }),
    EX("Hi. / Hello.", "朋友之間、一般場合", { fontSize: 13 }),
  ], fs: 13 });
  card(s, { x: 5.2, y: 1.05, w: 4.3, h: 2.4, label: "道別", body: [
    EX("Bye.", "再見", { fontSize: 13 }),
    EX("Take care.", "保重　→ 可回 You too.", { fontSize: 13 }),
    EX("See you later.", "下次再見", { fontSize: 13 }),
    EX("Good night.", "晚安（道別、睡前）", { fontSize: 13 }),
  ], fs: 13 });
  card(s, { x: 0.5, y: 3.6, w: 9.0, h: 1.5, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("Good evening. 是「晚上見面」時說的，常用在餐廳、演講等正式場合"),
    P("Good night. 是「晚上道別」或睡前說的，不是打招呼"),
  ], fs: 12 });
}

// =====================================================================
// CHAPTER 02
// =====================================================================
chapterSlide("02", "Who is this?", "這個人是誰？",
  ["疑問代名詞 who", "疑問副詞 how：狀態・程度", "人稱代名詞：所有格", "連接詞 and"], "43",
  "第二章開始學「問問題」。上一章的問句只能問 Yes / No，這一章開始用 who、how 問「誰」和「多～」。規則跟上一章一樣：be 動詞要跑到主詞前面。");

// 2.1 who
{
  const s = base("疑問代名詞 who：誰", 44,
    "who = 誰，放句首，後面接 be 動詞再接主詞。be 動詞還是看主詞：單數用 is，複數用 are。順便帶複數規則：名詞加 s。");
  formula(s, ["Who", "be 動詞", "（代）名詞 ?"], 0.5, 1.1, 5.4, 0.5, { fill: RED });
  table(s, [
    ["", "例句", "中譯"],
    [{ text: "單數 → is", bold: true, color: NAVY }, "Who is she?", "她是誰？"],
    ["", "Who is the girl?", "那女孩是誰？"],
    [{ text: "複數 → are", bold: true, color: NAVY }, "Who are they?", "他們是誰？"],
    ["", "Who are the girls?", "那些女孩是誰？"],
    [{ text: "you 都用 are", bold: true, color: NAVY }, "Who are you?", "你是誰？／你們是誰？"],
  ], { x: 0.5, y: 1.85, w: 5.4, colW: [1.5, 2.1, 1.8], fs: 12, rowH: 0.38 });
  card(s, { x: 6.2, y: 1.1, w: 3.3, h: 1.85, label: "重點", body: [
    P("who、how 開頭的問句，主詞和 be 動詞要對調，變成「be 動詞 ＋ 主詞」", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 6.2, y: 3.1, w: 3.3, h: 2.0, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("複數名詞 = 單數名詞 ＋ s", { fontSize: 12 }),
    EX("girl → girls", "女孩們", { fontSize: 12 }),
    EX("boy → boys", "男孩們", { fontSize: 12 }),
  ], fs: 12 });
}

// 2.2 how
{
  const s = base("疑問副詞 how：多～？", 45,
    "how 後面一定要接一個形容詞，才知道在問「多什麼」：多高、多重、多聰明。順序一樣：How ＋ 形容詞 ＋ be 動詞 ＋ 主詞。可以問學生 How tall are you?");
  text(s, [P("how 表「多少、如何」，後面接表狀態或程度的形容詞，用來問主詞「多～」")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  formula(s, ["How", "形容詞", "be 動詞", "主詞 ?"], 0.5, 1.5, 9.0, 0.5, { fill: RED });
  text(s, [
    EX("How tall is she?", "她有多高？", { fontSize: 16 }),
    EX("How heavy is he?", "他有多重？", { fontSize: 16 }),
    EX("How smart is the boy?", "那男孩有多聰明？", { fontSize: 16 }),
  ], { x: 0.5, y: 2.3, w: 5.5, h: 1.7, fontSize: 16, paraSpaceAfter: 8 });
  card(s, { x: 6.2, y: 2.3, w: 3.3, h: 2.7, label: "常用形容詞", body: [
    P("tall 高的"), P("heavy 重的"), P("old 老的、幾歲"), P("smart 聰明的"), P("long 長的"), P("big 大的"),
  ], fs: 13 });
  text(s, [M("下一頁對話會看到：How old is she?（幾歲）、How tall is she?（多高）")], { x: 0.5, y: 4.3, w: 5.5, h: 0.5, fontSize: 12 });
}

// 2.3 所有格
{
  const s = base("人稱代名詞：所有格", 45,
    "所有格 = 「誰的」。不能單獨用，後面一定要接名詞：my book、your sister。跟主格一一對照著背：I → my、you → your、he → his、she → her、it → its、we → our、they → their。普通名詞的所有格加 's。");
  text(s, [P("所有格表「～的」，不能單獨用，後面一定要接名詞")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  table(s, [
    ["人稱", "主格", "所有格", "主格（複數）", "所有格（複數）"],
    ["第 1 人稱", "I 我", { text: "my 我的", bold: true, color: RED }, "we 我們", { text: "our 我們的", bold: true, color: RED }],
    ["第 2 人稱", "you 你", { text: "your 你的", bold: true, color: RED }, "you 你們", { text: "your 你們的", bold: true, color: RED }],
    ["第 3 人稱", "he 他", { text: "his 他的", bold: true, color: RED }, "they 他們", { text: "their 他們的", bold: true, color: RED }],
    ["", "she 她", { text: "her 她的", bold: true, color: RED }, "", ""],
    ["", "it 它", { text: "its 它的", bold: true, color: RED }, "", ""],
  ], { x: 0.5, y: 1.45, w: 6.1, colW: [1.0, 1.0, 1.2, 1.4, 1.5], fs: 12, rowH: 0.34 });
  card(s, { x: 6.8, y: 1.45, w: 2.7, h: 1.75, label: "例句", body: [
    EX("He is my brother.", "他是我弟弟", { fontSize: 12 }),
    EX("Your sister is pretty.", "你姐姐很漂亮", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 6.8, y: 3.35, w: 2.7, h: 1.75, label: "名詞的所有格", labelW: 1.2, body: [
    P("名詞 ＋ 's", { fontSize: 12, bold: true }),
    EX("the baby's smile", "寶寶的笑容", { fontSize: 12 }),
    EX("the girl's book", "那女孩的書", { fontSize: 12 }),
  ], fs: 12 });
  text(s, [M("注意：his / her 前面不加 the，直接接名詞：his book（○）the his book（✕）")], { x: 0.5, y: 4.55, w: 6.1, h: 0.5, fontSize: 12 });
}

// 2.4 and / but
{
  const s = base("連接詞 and / but", 45,
    "and 把兩個東西連起來，意思相近或並列；but 是轉折，前後意思相反。可以連單字，也可以連整句。連整句時，and / but 前面加逗號。");
  card(s, { x: 0.5, y: 1.1, w: 4.3, h: 2.2, label: "and　和、而且", labelW: 1.4, body: [
    EX("the girl and the boy", "那女孩和那男孩（連單字）", { fontSize: 13 }),
    EX("He is handsome, and she is smart.", "他很英俊，而她很聰明（連句子）", { fontSize: 13 }),
  ], fs: 13 });
  card(s, { x: 5.2, y: 1.1, w: 4.3, h: 2.2, fill: REDTINT, label: "but　但是", labelFill: RED, labelW: 1.4, body: [
    P("前後意思相反、轉折時用 but", { fontSize: 13 }),
    EX("He is handsome, but he isn't smart.", "他很英俊，但是不聰明", { fontSize: 13 }),
  ], fs: 13 });
  card(s, { x: 0.5, y: 3.5, w: 9.0, h: 1.6, label: "三個以上", body: [
    P("連接兩個以上的形容詞，只要在最後一個前面加 and，中間用逗號"),
    EX("She is beautiful, slim, and cute.", "她漂亮、苗條又可愛", { fontSize: 14 }),
  ], fs: 13 });
}

practiceSlide("02", [
  { q: "填所有格", lines: ["This is (I) ___ book.", "That is (she) ___ brother.", "(they) ___ house is big."] },
  { q: "填 Who 或 How", lines: ["___ is she?（誰）", "___ old is he?", "___ tall is the boy?"] },
  { q: "填 and 或 but", lines: ["He is tall, ___ she is short.", "She is kind ___ smart."] },
  { q: "改成名詞所有格", lines: ["the smile of the baby"] },
], [
  ["my", "her", "Their"],
  ["Who", "How", "How"],
  ["but", "and"],
  ["the baby's smile"],
], "第 2 題提醒：who 後面直接接 be 動詞；how 後面要先接形容詞。");

// 對話 1
{
  const s = base("對話 1　Who is this? / How old is she?", 46,
    "情境：指著照片問「這是誰」。年紀句型：How old ＋ be ＋ 主詞？回答：主詞 ＋ be ＋ 數字 ＋ years old。兩歲以上 year 要加 s。really 放在形容詞前面加強語氣。");
  const dlg = [
    ["Carlo", "Who is this?", "這是誰？"],
    ["Elise", "He is Joshua. He is my brother.", "他是約書亞，是我的弟弟。"],
    ["Carlo", "He is really cute. Who is this?", "他真可愛。這是誰？"],
    ["Elise", "She is Anna. She is my sister.", "她是安娜，是我的妹妹。"],
    ["Carlo", "How old is she?", "她幾歲？"],
    ["Elise", "She is fifteen years old.", "她十五歲。"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 14 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 4.1, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 2.15, label: "問年紀", body: [
    EX("How old is she?", "她幾歲？", { fontSize: 12 }),
    EX("She is fifteen years old.", "她十五歲", { fontSize: 12 }),
    P("one year old　→　two years old", { fontSize: 12 }),
    M("兩歲以上 year 加 s；也可以只回答數字", { fontSize: 10 }),
  ], fs: 12 });
  card(s, { x: 6.1, y: 3.35, w: 3.4, h: 1.8, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("really（非常）放在形容詞前面加強：", { fontSize: 12 }),
    EX("He is cute. → He is really cute.", "", { fontSize: 12 }),
    M("brother 哥哥／弟弟 · sister 姐姐／妹妹 · cute 可愛的", { fontSize: 10 }),
  ], fs: 12 });
}

// 對話 2
{
  const s = base("對話 2　How tall is she?", 47,
    "身高句型：主詞 ＋ be ＋ 數字 ＋ centimeters tall。美國人用 feet，1 foot = 30.48 公分，小數點唸 point。最後一句是三個形容詞用逗號和 and 連起來。");
  const dlg = [
    ["Justin", "Who is that?", "那個人是誰？"],
    ["Sooji", "Yuri. She is a K-pop teen idol.", "俞利。她是韓國的偶像歌手。"],
    ["Justin", "She is beautiful. How tall is she?", "真漂亮。她有多高？"],
    ["Sooji", "She is 167 centimeters tall. So she is around 5.5 feet.", "167 公分。所以大約是 5.5 呎高。"],
    ["Justin", "She is beautiful, slim, and cute.", "她長得漂亮、苗條又可愛。"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 14 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 3.2, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 0.5, y: 4.15, w: 5.4, h: 1.0, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("this 這個（近）／ that 那個（遠）；5.5 唸 five point five", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 4.1, label: "問身高", body: [
    EX("How tall is she?", "她有多高？", { fontSize: 12 }),
    EX("She is 167 centimeters tall.", "她 167 公分", { fontSize: 12 }),
    EX("She is around 5.5 feet.", "她大約 5.5 呎", { fontSize: 12 }),
    GAP(0.3),
    M("teen idol 偶像 · beautiful 漂亮的 · tall 高的 · centimeter 公分（cm）· so 所以 · around 大約 · feet 呎（foot 複數）· slim 苗條的", { fontSize: 10 }),
  ], fs: 12 });
}

// 家庭成員
{
  const s = base("單字補給站　家庭成員名稱", 48,
    "先帶學生唸一次，然後讓他用 This is my ＋ 家人 介紹自己家人。英文的 brother 不分哥哥弟弟，uncle 不分伯叔舅姑姨丈。");
  const fam = (list) => [["單字", "中譯"], ...list.map(([en, zh]) => [{ text: en, bold: true, color: NAVY }, zh])];
  table(s, fam([
    ["grandfather / grandpa", "爺爺；外公"], ["grandmother / grandma", "奶奶；外婆"],
    ["father / dad", "父親；爸爸"], ["mother / mom", "母親；媽媽"],
    ["uncle", "伯伯、叔叔、舅舅、姑丈、姨丈"], ["aunt", "伯母、嬸嬸、姑姑、阿姨、舅媽"],
    ["brother", "哥哥；弟弟"], ["sister", "姐姐；妹妹"],
    ["cousin", "堂／表兄弟姐妹"],
  ]), { x: 0.5, y: 1.05, w: 5.2, colW: [2.3, 2.9], fs: 11, rowH: 0.3 });
  table(s, fam([
    ["grandparents", "（外）祖父母"], ["parents", "父母"],
    ["husband", "先生、丈夫"], ["wife", "太太、妻子"],
    ["son", "兒子"], ["daughter", "女兒"],
    ["grandson", "孫子"], ["granddaughter", "孫女"],
    ["nephew", "姪子"], ["niece", "姪女"],
  ]), { x: 5.9, y: 1.05, w: 3.6, colW: [1.7, 1.9], fs: 11, rowH: 0.3 });
  text(s, [M("練習：This is my ______.　用自己家人造句")], { x: 0.5, y: 4.7, w: 5.2, h: 0.4, fontSize: 12 });
}

// 實用表達法 讚嘆 / 關懷
{
  const s = base("實用表達法　讚嘆與關懷", 49,
    "讚嘆：不用主詞和 be 動詞，一個形容詞加驚嘆號就好。關懷：You look ＋ 形容詞 = 你看起來～。What's wrong? 怎麼了。這些句子直接念、直接記，不用分析文法。");
  card(s, { x: 0.5, y: 1.05, w: 4.3, h: 4.05, label: "表達讚嘆", body: [
    EX("Fantastic!", "太棒了！", { fontSize: 13 }),
    EX("Absolutely gorgeous!", "真是太美了！", { fontSize: 13 }),
    GAP(0.3),
    P("省略「主詞 ＋ be 動詞」，形容詞加驚嘆號就好：", { fontSize: 12 }),
    EX("Amazing!", "太驚人了", { fontSize: 12 }),
    EX("Excellent!", "太出色了", { fontSize: 12 }),
    EX("Great!", "非常好", { fontSize: 12 }),
    EX("Beautiful!", "太漂亮了", { fontSize: 12 }),
    EX("Wonderful!", "太棒了", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 5.2, y: 1.05, w: 4.3, h: 4.05, fill: REDTINT, label: "表達關懷", labelFill: RED, body: [
    EX("You look down today. What's wrong?", "你今天看起來心情不好。怎麼了？", { fontSize: 12 }),
    EX("I finally broke up with my boyfriend.", "我終於和男友分手了", { fontSize: 12 }),
    GAP(0.3),
    P("其他問法：", { fontSize: 12 }),
    EX("Is there something wrong?", "出了什麼事嗎？", { fontSize: 12 }),
    EX("What's the matter?", "發生了什麼事？", { fontSize: 12 }),
    EX("Are you okay?", "你還好嗎？", { fontSize: 12 }),
    GAP(0.3),
    P("You look ＋ 形容詞 = 你看起來～　You look happy.", { fontSize: 12 }),
  ], fs: 12 });
}

// =====================================================================
// CHAPTER 03
// =====================================================================
chapterSlide("03", "What is this?", "這是什麼？",
  ["疑問代名詞 what", "指示代名詞 this, that", "人稱代名詞：受格"], "53",
  "第三章：問「什麼」，還有「這個／那個」，最後是受格。受格對中文母語者是新東西，因為中文「我」不分主格受格，要多花一點時間。");

// 3.1 what
{
  const s = base("疑問代名詞 what：什麼", 54,
    "what 問東西的名稱或人的職業。句型跟 who 一模一樣，只是把 who 換成 what。整理一下：who / how / what 三個疑問詞都放句首，後面都是 be 動詞 ＋ 主詞。");
  formula(s, ["What", "be 動詞", "（代）名詞 ?"], 0.5, 1.1, 5.4, 0.5, { fill: RED });
  table(s, [
    ["", "例句", "中譯"],
    [{ text: "單數 → is", bold: true, color: NAVY }, "What is this?", "這是什麼？"],
    ["", "What is her name?", "她叫什麼名字？"],
    [{ text: "複數 → are", bold: true, color: NAVY }, "What are these?", "這些是什麼？"],
    ["", "What are those?", "那些是什麼？"],
  ], { x: 0.5, y: 1.85, w: 5.4, colW: [1.5, 2.1, 1.8], fs: 12, rowH: 0.4 });
  card(s, { x: 6.2, y: 1.1, w: 3.3, h: 4.0, fill: REDTINT, label: "整理", labelFill: RED, body: [
    P("who、how、what 三個疑問詞的共同點：", { fontSize: 12 }),
    P("① 放在句首", { fontSize: 12, bold: true }),
    P("② 「主詞 ＋ be 動詞」對調成「be 動詞 ＋ 主詞」", { fontSize: 12, bold: true }),
    GAP(0.3),
    EX("Who is she?", "誰", { fontSize: 12 }),
    EX("How tall is she?", "多～", { fontSize: 12 }),
    EX("What is this?", "什麼", { fontSize: 12 }),
  ], fs: 12 });
}

// 3.2 this / that / these / those
{
  const s = base("指示代名詞 this / that / these / those", 54,
    "近的用 this，遠的用 that；複數變 these / those。this / that 當第三人稱單數，配 is；these / those 配 are，後面的名詞要加 s。用手指教室裡的東西示範最快。");
  table(s, [
    ["", "單數（1 個）→ is", "複數（2 個以上）→ are"],
    [{ text: "近　靠近說話者", bold: true, color: NAVY }, { text: "this  這個", bold: true, color: RED }, { text: "these  這些", bold: true, color: RED }],
    ["", "This is my car.　這是我的車", "These are my cars.　這些是我的車"],
    [{ text: "遠　離說話者遠", bold: true, color: NAVY }, { text: "that  那個", bold: true, color: RED }, { text: "those  那些", bold: true, color: RED }],
    ["", "That is my car.　那是我的車", "Those are my cars.　那些是我的車"],
  ], { x: 0.5, y: 1.1, w: 9.0, colW: [2.0, 3.3, 3.7], fs: 13, rowH: 0.48 });
  card(s, { x: 0.5, y: 3.7, w: 4.3, h: 1.4, label: "重點", body: [
    P("this / that 算第三人稱單數 → 接 is"),
    P("these / those 是複數 → 接 are，名詞加 (e)s"),
  ], fs: 12 });
  card(s, { x: 5.2, y: 3.7, w: 4.3, h: 1.4, fill: REDTINT, label: "問句", labelFill: RED, body: [
    EX("What is this?", "這是什麼？", { fontSize: 12 }),
    EX("What are those?", "那些是什麼？", { fontSize: 12 }),
  ], fs: 12 });
}

// 3.3 受格
{
  const s = base("人稱代名詞：受格", 55,
    "受格 = 放在動詞後面、「被動作」的那個人。中文「我愛你」「你愛我」的「我」長一樣，英文不一樣：主詞位置用 I，受詞位置用 me。及物動詞 = 後面一定要接受詞的動詞，例如 love、hate。你 you 和它 it 主格受格同形。");
  formula(s, ["主詞（主格）", "及物動詞", "受詞（受格）"], 0.5, 1.05, 9.0, 0.48);
  text(s, [
    EX("I love you.", "我愛你　→ I 主格、you 受格", { fontSize: 14 }),
    EX("You love me.", "你愛我　→ You 主格、me 受格", { fontSize: 14 }),
  ], { x: 0.5, y: 1.65, w: 5.6, h: 0.85, fontSize: 14 });
  table(s, [
    ["人稱", "主格", "受格", "主格（複數）", "受格（複數）"],
    ["第 1 人稱", "I", { text: "me", bold: true, color: RED }, "we", { text: "us", bold: true, color: RED }],
    ["第 2 人稱", "you", { text: "you", bold: true, color: RED }, "you", { text: "you", bold: true, color: RED }],
    ["第 3 人稱", "he / she / it", { text: "him / her / it", bold: true, color: RED }, "they", { text: "them", bold: true, color: RED }],
  ], { x: 0.5, y: 2.6, w: 5.7, colW: [0.95, 1.15, 1.3, 1.15, 1.15], fs: 12, rowH: 0.36 });
  card(s, { x: 6.4, y: 1.65, w: 3.1, h: 1.6, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("中文的「我」沒有主格受格之分，英文要分 I 和 me", { fontSize: 11 }),
    P("受詞是普通名詞則不變：I love Lisa.", { fontSize: 11 }),
  ], fs: 11 });
  card(s, { x: 6.4, y: 3.4, w: 3.1, h: 1.7, label: "重點", body: [
    P("不及物動詞：後面不用接受詞", { fontSize: 11 }),
    EX("I sleep.", "我睡覺", { fontSize: 12 }),
    EX("I laugh.", "我笑", { fontSize: 12 }),
  ], fs: 11 });
  text(s, [M("介係詞（to / from / for）後面接代名詞也用受格：for him、from me")], { x: 0.5, y: 4.6, w: 5.7, h: 0.5, fontSize: 12 });
}

practiceSlide("03", [
  { q: "填 this / that / these / those", lines: ["___ is my car.（近，一台）", "___ are my cars.（遠，很多台）"] },
  { q: "填受格", lines: ["I love (she) ___.", "She loves (I) ___.", "We love (they) ___.", "This present is for (he) ___."] },
  { q: "填疑問詞", lines: ["___ is this?（什麼）", "___ is she?（誰）"] },
  { q: "改成正確大小寫", lines: ["korea", "christmas"] },
], [
  ["This", "Those"],
  ["her", "me", "them", "him"],
  ["What", "Who"],
  ["Korea", "Christmas"],
], "第 2 題最後一小題：for 是介係詞，後面也要用受格。");

// 對話 1
{
  const s = base("對話 1　What is this?", 56,
    "情境：送禮物。a present from me / a present for you 都是「我要送你的禮物」。awesome 是口語的「超棒」。順便帶大寫規則：人名、地名、國名、節日第一個字母大寫。");
  const dlg = [
    ["Eric", "What is this?", "這是什麼？"],
    ["Sooji", "This is \"Taegeuksun.\" It's a Korean fan. It is a present from me.", "這是太極扇，韓國的扇子。這是我要送你的禮物。"],
    ["Eric", "Thank you, Sooji. It is really beautiful. What is that?", "謝謝，守智。這個真漂亮。那是什麼？"],
    ["Sooji", "That is a Korean mask. It's my present for you, too.", "那是韓國的面具。那也是我要送你的禮物。"],
    ["Eric", "Really? Thank you. It's awesome.", "真的嗎？謝謝，好棒喔！"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 13 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 4.1, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 2.0, label: "新用法", body: [
    EX("It is a present from me.", "這是我送你的禮物", { fontSize: 12 }),
    EX("Really?", "真的嗎？", { fontSize: 12 }),
    EX("It's awesome.", "超棒的（口語）", { fontSize: 12 }),
    M("awesome ≈ great / cool / fantastic", { fontSize: 10 }),
  ], fs: 12 });
  card(s, { x: 6.1, y: 3.2, w: 3.4, h: 1.95, fill: REDTINT, label: "大寫", labelFill: RED, body: [
    P("第一個字母要大寫：人名、地名、國名、國籍、語言、節日", { fontSize: 11 }),
    P("Kevin · New York · Korea · Korean · Christmas", { fontSize: 11, bold: true, color: NAVY }),
    M("fan 扇子 · present 禮物 · mask 面具 · for 為了", { fontSize: 10 }),
  ], fs: 11 });
}

// 對話 2
{
  const s = base("對話 2　What are these?", 57,
    "these / those 的實際用法。How beautiful! 是感嘆句不是問句。可數名詞 vs 不可數名詞先有概念就好：能數的（apple、cat）可以加 s；抽象的（love、friendship）不行。");
  const dlg = [
    ["Carlo", "How beautiful! What are these?", "好漂亮喔！這些是什麼？"],
    ["Asako", "These are origami flowers.", "這些是摺紙花。"],
    ["Carlo", "What are those?", "那些是什麼？"],
    ["Asako", "Those are origami cranes.", "那些是紙鶴。"],
    ["Carlo", "What is origami?", "什麼是摺紙？"],
    ["Asako", "It's a traditional Japanese craft.", "是一種傳統的日本手工藝。"],
    ["Carlo", "That is amazing.", "真是令人驚奇。"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 13 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 4.1, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 1.8, label: "感嘆句", body: [
    P("How ＋ 形容詞！　不是問句，是「好～啊！」", { fontSize: 12 }),
    EX("How beautiful!", "好漂亮啊！", { fontSize: 12 }),
    EX("How awesome!", "好厲害啊！", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 6.1, y: 3.0, w: 3.4, h: 2.15, fill: REDTINT, label: "名詞分兩種", labelFill: RED, labelW: 1.2, body: [
    P("可數：數得出來，可加 s　apple、cat、flower", { fontSize: 11 }),
    P("不可數：抽象、專有名詞　love、friendship", { fontSize: 11 }),
    M("origami 摺紙 · crane 鶴 · traditional 傳統的 · craft 手工藝 · amazing 驚人的", { fontSize: 10 }),
  ], fs: 11 });
}

// 形容詞
{
  const s = base("單字補給站　形容詞", 58,
    "成對記：nice / bad、pretty / ugly、slim / fat、rich / poor⋯⋯。情緒的形容詞配 be 動詞就能造句：I am happy. 讓學生挑三個形容自己。");
  text(s, [H("外貌、個性、狀態（成對記）", { fontSize: 13 })], { x: 0.5, y: 1.0, w: 5, h: 0.3 });
  table(s, [
    ["單字", "中譯", "相反", "中譯"],
    ["nice", "親切的", "bad", "壞的"],
    ["pretty", "美麗的", "ugly", "醜的"],
    ["slim", "苗條的", "fat", "胖的"],
    ["kind", "和善的", "unkind", "不和善的"],
    ["rich", "富有的", "poor", "貧窮的"],
    ["tall", "高大的", "small", "矮小的"],
    ["young", "年輕的", "old", "年老的"],
    ["cheerful", "開朗的", "serious", "嚴肅的"],
    ["healthy", "健康的", "sick", "生病的"],
  ], { x: 0.5, y: 1.35, w: 4.6, colW: [1.15, 1.15, 1.15, 1.15], fs: 11, rowH: 0.3 });
  text(s, [H("情緒", { fontSize: 13 })], { x: 5.4, y: 1.0, w: 4, h: 0.3 });
  table(s, [
    ["正面", "中譯", "負面", "中譯"],
    ["happy", "快樂的", "sad", "難過的"],
    ["excited", "興奮的", "gloomy", "憂鬱的"],
    ["interested", "感興趣的", "worried", "憂心的"],
    ["proud", "自豪的", "anxious", "焦慮的"],
    ["confident", "有信心的", "disappointed", "失望的"],
    ["satisfied", "滿意的", "lonely", "孤單的"],
  ], { x: 5.4, y: 1.35, w: 4.1, colW: [1.05, 0.95, 1.2, 0.9], fs: 11, rowH: 0.3 });
  text(s, [M("練習：I am ______.　挑三個形容今天的自己")], { x: 5.4, y: 3.6, w: 4.1, h: 0.4, fontSize: 12 });
}

// 表達感謝
{
  const s = base("實用表達法　表達感謝", 59,
    "Thank you for ＋ 東西。回應三種：You're welcome. / No problem. / My pleasure. I appreciate 後面接事情不接人（appreciate your help ○，appreciate you ✕）。Here is / Here are 看後面名詞單複數。");
  card(s, { x: 0.5, y: 1.05, w: 4.3, h: 2.5, label: "說謝謝", body: [
    EX("Thank you for the present.", "謝謝你的禮物", { fontSize: 12 }),
    EX("I appreciate your help.", "謝謝你的幫忙（較正式）", { fontSize: 12 }),
    EX("Thank you very much.", "非常感謝", { fontSize: 12 }),
    M("appreciate 後面接事情（your help / your time），不接人", { fontSize: 10 }),
  ], fs: 12 });
  card(s, { x: 5.2, y: 1.05, w: 4.3, h: 2.5, fill: REDTINT, label: "回應", labelFill: RED, body: [
    EX("You're welcome.", "不客氣", { fontSize: 12 }),
    EX("No problem.", "沒什麼", { fontSize: 12 }),
    EX("(It's) my pleasure.", "我的榮幸 → 不客氣", { fontSize: 12 }),
    EX("I'm glad you like it.", "很高興你喜歡", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 0.5, y: 3.7, w: 9.0, h: 1.4, label: "Here is / Here are", labelW: 1.5, body: [
    P("把東西遞給別人時說。be 動詞看後面的名詞：單數 is、複數 are"),
    EX("Here is an apple.", "這裡有顆蘋果", { fontSize: 12 }),
    EX("Here are your books.", "這些是你的書", { fontSize: 12 }),
  ], fs: 12 });
}

// =====================================================================
// CHAPTER 04
// =====================================================================
chapterSlide("04", "Where is the laundry basket?", "洗衣籃在哪裡？",
  ["疑問副詞 where", "表達場所的副詞", "介係詞"], "63",
  "第四章：問「在哪裡」。be 動詞在這章的意思從「是」變成「在」。介係詞是重點，用教室裡的東西實際比給學生看。");

// 4.1 where
{
  const s = base("疑問副詞 where：在哪裡", 64,
    "where 放句首，後面 be 動詞 ＋ 主詞，跟前面三個疑問詞一樣。注意這裡 be 動詞的意思是「在」不是「是」。Where am I? 直譯「我在哪」，其實是問「這是哪裡」。");
  formula(s, ["Where", "be 動詞", "主詞 ?"], 0.5, 1.1, 5.4, 0.5, { fill: RED });
  table(s, [
    ["", "例句", "中譯"],
    [{ text: "單數", bold: true, color: NAVY }, "Where am I?", "這是哪裡？"],
    ["", "Where is he?", "他在哪裡？"],
    [{ text: "複數", bold: true, color: NAVY }, "Where are they?", "他們在哪裡？"],
    ["", "Where are the students?", "學生們在哪裡？"],
  ], { x: 0.5, y: 1.85, w: 5.4, colW: [1.0, 2.5, 1.9], fs: 12, rowH: 0.4 });
  card(s, { x: 6.2, y: 1.1, w: 3.3, h: 1.85, label: "重點", body: [
    P("前面學的 be 動詞是「是～」，這裡的 be 動詞是「在（某地）」", { fontSize: 12 }),
    EX("She is over there.", "她在那裡", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 6.2, y: 3.1, w: 3.3, h: 2.0, fill: REDTINT, label: "注意", labelFill: RED, body: [
    P("Where am I? 直譯「我在哪裡？」，實際是問「這是哪裡？」，迷路時用", { fontSize: 12 }),
  ], fs: 12 });
}

// 4.2 場所副詞
{
  const s = base("表達場所、地點的副詞", 64,
    "這六個字後面不用再接名詞，直接放在 be 動詞後面就是完整句子。here / there 最常用，來回比手勢。");
  text(s, [P("直接接在 be 動詞後面，不用再接名詞")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  const words = [["here", "這裡"], ["there", "那裡"], ["up", "往上"], ["down", "往下"], ["far", "遠"], ["near", "近"]];
  words.forEach(([en, zh], i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.5 + Math.floor(i / 3) * 1.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.85, h: 1.05, fill: { color: TINT }, line: { color: TINT }, rectRadius: 0.1 });
    s.addText(en, { x: x + 0.2, y: y + 0.1, w: 2.45, h: 0.5, fontFace: F, fontSize: 24, bold: true, color: NAVY, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(zh, { x: x + 0.2, y: y + 0.6, w: 2.45, h: 0.35, fontFace: F, fontSize: 13, color: MUTED, margin: 0, isTextBox: true });
  });
  text(s, [
    EX("I am here.", "我在這裡", { fontSize: 15 }),
    EX("He is there.", "他在那裡", { fontSize: 15 }),
    EX("She is over there.", "她在那邊（比 there 更遠一點）", { fontSize: 15 }),
  ], { x: 0.5, y: 4.15, w: 9, h: 1.0, fontSize: 15, paraSpaceAfter: 2 });
}

// 4.3 介係詞：場所
{
  const s = base("介係詞 ①　表示場所：at / in", 65,
    "介係詞後面一定要接名詞。at 用在小地方、一個「點」：at home、at school；in 用在大範圍、一個「區域」：in Taipei、in Asia。讓學生說 I am at home. / I am in Taiwan.");
  text(s, [P("介係詞後面接名詞，用來說明「在哪裡」。at 用於小地點，in 用於大範圍")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  card(s, { x: 0.5, y: 1.5, w: 4.3, h: 2.5, label: "at　小地點、一個「點」", labelW: 2.0, body: [
    EX("at work", "在工作、在上班", { fontSize: 14 }),
    EX("at school", "在學校", { fontSize: 14 }),
    EX("at home", "在家", { fontSize: 14 }),
  ], fs: 14 });
  card(s, { x: 5.2, y: 1.5, w: 4.3, h: 2.5, fill: REDTINT, label: "in　大範圍、一個「區域」", labelFill: RED, labelW: 2.2, body: [
    EX("in Asia", "在亞洲", { fontSize: 14 }),
    EX("in Korea", "在韓國", { fontSize: 14 }),
    EX("in Seoul", "在首爾", { fontSize: 14 }),
  ], fs: 14 });
  text(s, [
    EX("Where is Eric?　—　He's at home.", "艾瑞克在哪裡？他在家", { fontSize: 15 }),
    EX("I am in Taiwan. I am at school.", "我在臺灣。我在學校", { fontSize: 15 }),
  ], { x: 0.5, y: 4.2, w: 9, h: 0.9, fontSize: 15, paraSpaceAfter: 4 });
}

// 4.3 介係詞：位置
{
  const s = base("介係詞 ②　表示位置：on / under / behind…", 65,
    "拿一個杯子和一本書在鏡頭前示範：on the book、under the book、next to the book、in front of、behind、between。一個一個比，讓學生跟著說。the = 特定的、你我都知道的那個東西。");
  table(s, [
    ["介係詞", "意思", "例子", "中譯"],
    [{ text: "on", bold: true, color: RED }, "在～上面", "on the desk", "在書桌上"],
    [{ text: "under", bold: true, color: RED }, "在～下面", "under the desk", "在書桌下"],
    [{ text: "in", bold: true, color: RED }, "在～裡面", "in the refrigerator", "在冰箱裡"],
    [{ text: "in front of", bold: true, color: RED }, "在～前面", "in front of the chair", "在椅子前面"],
    [{ text: "behind", bold: true, color: RED }, "在～後面", "behind the chair", "在椅子後面"],
    [{ text: "next to / by", bold: true, color: RED }, "在～旁邊", "next to the clock", "在時鐘旁邊"],
    [{ text: "between A and B", bold: true, color: RED }, "在 A 和 B 之間", "between the cup and the clock", "在杯子和時鐘之間"],
  ], { x: 0.5, y: 1.05, w: 6.3, colW: [1.75, 1.15, 2.15, 1.25], fs: 11, rowH: 0.34 });
  card(s, { x: 7.0, y: 1.05, w: 2.5, h: 2.0, label: "還有", body: [
    P("over 在～上方", { fontSize: 12 }),
    P("below 在～下面", { fontSize: 12 }),
    M("課本 p.65 有圖，對照看", { fontSize: 10 }),
  ], fs: 12 });
  card(s, { x: 7.0, y: 3.2, w: 2.5, h: 1.9, fill: REDTINT, label: "the", labelFill: RED, body: [
    P("定冠詞，放名詞前面，指「特定的、你我都知道的」那個東西", { fontSize: 11 }),
    P("the desk = 那張桌子", { fontSize: 11, bold: true }),
  ], fs: 11 });
  text(s, [M("介係詞後面一定要接具體的名詞：on the desk（○）on（✕）")], { x: 0.5, y: 4.55, w: 6.3, h: 0.5, fontSize: 12 });
}

practiceSlide("04", [
  { q: "填 Where ＋ be 動詞", lines: ["___ ___ he?", "___ ___ the students?"] },
  { q: "填 at 或 in", lines: ["___ home", "___ Korea", "___ school", "___ Asia"] },
  { q: "用英文說位置", lines: ["書在桌子上", "蘋果在冰箱裡", "貓在椅子後面", "時鐘在杯子旁邊"] },
  { q: "樓層", lines: ["在三樓：on the ___ floor"] },
], [
  ["Where is he?", "Where are the students?"],
  ["at", "in", "at", "in"],
  ["on the desk", "in the refrigerator", "behind the chair", "next to the cup / by the cup"],
  ["third（3rd）"],
], "第 3 題可以延伸：讓學生講自己桌上東西的位置。");

// 對話 1
{
  const s = base("對話 1　Where is the laundry basket?", 66,
    "情境：住宿守則。四條規則是祈使句：省略 you，動詞原形開頭。對話本身練 Where is ＋ 東西？答 It's ＋ 介係詞 ＋ 地方。I see. / I got it. 都是「了解」。");
  const dlg = [
    ["Eric", "Sooji, these are the rules for this house.", "守智，這些是這家裡的住宿守則。"],
    ["Sooji", "Eric, where is the laundry basket?", "艾瑞克，洗衣籃在哪裡？"],
    ["Eric", "It's behind your closet door.", "在你的衣櫃門後面。"],
    ["Sooji", "I see. Is the dishwasher in the kitchen?", "了解。洗碗機在廚房嗎？"],
    ["Eric", "Yes, it's under the sink.", "對，在洗手槽下面。"],
    ["Sooji", "I got it.", "我知道了。"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 13 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 4.1, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 2.3, label: "住宿守則（祈使句）", labelW: 1.6, body: [
    P("1. Put your laundry in the basket in your room.", { fontSize: 10 }),
    P("2. Put your plates in the dishwasher after your meals.", { fontSize: 10 }),
    P("3. Come home by 11 p.m.", { fontSize: 10 }),
    P("4. Call Eric after you leave.", { fontSize: 10 }),
    M("祈使句：省略 you，動詞原形開頭　Come here. 來這裡", { fontSize: 10 }),
  ], fs: 10 });
  card(s, { x: 6.1, y: 3.5, w: 3.4, h: 1.65, fill: REDTINT, label: "新用法", labelFill: RED, body: [
    EX("I see.", "原來如此", { fontSize: 12 }),
    EX("I got it.", "我知道了（口語）", { fontSize: 12 }),
    M("laundry 換洗衣物 · basket 籃子 · closet 衣櫥 · sink 水槽", { fontSize: 10 }),
  ], fs: 12 });
}

// 對話 2
{
  const s = base("對話 2　Where is the frozen section?", 67,
    "情境：超市問路。section = 區。樓層一定用 on the ＋ 序數 ＋ floor：on the 2nd floor。Excuse me. 是問路開頭的「不好意思」。");
  const dlg = [
    ["Elise", "Where are the eggs in this supermarket?", "這間超市的雞蛋在哪裡？"],
    ["Justin", "I don't know. Excuse me. Where is the dairy section?", "我不知道。不好意思，請問乳製品區在哪裡？"],
    ["Clerk", "Oh, it's between the frozen section and the vegetable section.", "喔，在冷凍食品區和蔬菜區之間。"],
    ["Justin", "Where is the frozen section?", "冷凍食品區在哪裡？"],
    ["Clerk", "It's next to the meat section. There. It's on the 2nd floor.", "在肉品區旁邊。在那裡。就在二樓。"],
    ["Justin", "I see. Thank you.", "原來如此，謝謝你。"],
  ];
  const body = [];
  dlg.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: 12 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: 13 } },
    ]);
    body.push(M("　　　　" + zh, { fontSize: 11 }));
  });
  text(s, body, { x: 0.5, y: 1.05, w: 5.4, h: 4.1, fontSize: 13, paraSpaceAfter: 2 });
  card(s, { x: 6.1, y: 1.05, w: 3.4, h: 2.15, label: "說樓層", body: [
    P("on the ＋ 序數 ＋ floor", { fontSize: 12, bold: true }),
    EX("on the first (1st) floor", "在一樓", { fontSize: 12 }),
    EX("on the third (3rd) floor", "在三樓", { fontSize: 12 }),
    M("介係詞用 on，數字用序數", { fontSize: 10 }),
  ], fs: 12 });
  card(s, { x: 6.1, y: 3.35, w: 3.4, h: 1.8, fill: REDTINT, label: "新用法", labelFill: RED, body: [
    EX("Excuse me.", "不好意思，請問一下", { fontSize: 12 }),
    EX("I don't know.", "我不知道", { fontSize: 12 }),
    M("dairy 乳製品 · frozen 冷凍的 · vegetable 蔬菜 · meat 肉", { fontSize: 10 }),
  ], fs: 12 });
}

// 感嘆詞
{
  const s = base("單字補給站　各類感嘆詞", 68,
    "這頁輕鬆帶，用聲音和表情示範，不用背。挑最常用的：Wow、Oops、Ouch、Oh my gosh、Yay。");
  vocab(s, [
    ["A-ha!", "啊哈！"], ["Oh my gosh!", "噢，我的天啊！"],
    ["Wow!", "哇！（驚訝、佩服）"], ["Hurray!", "萬歲！（非常開心）"],
    ["Ssh!", "噓！（要人安靜）"], ["Oops!", "唉呀！（犯小錯）"],
    ["Ouch!", "哎呀！（突然痛）"], ["Peekaboo!", "躲貓貓（逗小孩）"],
    ["Tada!", "噠啦！（介紹驚喜）"], ["Tut tut!", "嘖嘖！（不認同）"],
    ["Whew!", "呼！（好險）"], ["Yay!", "呀呼！（歡欣）"],
    ["Yikes!", "唉呀！（噁心、厭惡）"], ["Yoo-hoo!", "喲呼！（大聲叫人）"],
    ["Gosh!", "天啊！（遇到壞事）"], ["Psst!", "喂！噓！（偷偷叫人）"],
    ["Shoot!", "唉！（憤怒、失望）"],
  ], { x: 0.5, y: 1.05, w: 9.0, cols: 2, fs: 11, rowH: 0.3 });
}

// 許願
{
  const s = base("實用表達法　許願、祝福", 69,
    "Good luck! 最常用，回 Thanks. Good luck to you, too.。旅行前說 Have a nice trip!。這頁一樣直接念、直接記。");
  card(s, { x: 0.5, y: 1.05, w: 2.9, h: 4.05, label: "祝好運", body: [
    EX("Good luck!", "祝你好運", { fontSize: 12 }),
    EX("Thanks. Good luck to you, too.", "謝謝，也祝你好運", { fontSize: 12 }),
    GAP(0.3),
    P("其他說法：", { fontSize: 11 }),
    EX("All the best.", "萬事順利", { fontSize: 11 }),
    EX("I wish you the best of luck.", "祝你好運", { fontSize: 11 }),
    EX("I'll keep my fingers crossed for you.", "我會為你祈求好運", { fontSize: 11 }),
  ], fs: 12 });
  card(s, { x: 3.55, y: 1.05, w: 2.9, h: 4.05, fill: REDTINT, label: "祝旅途愉快", labelFill: RED, labelW: 1.2, body: [
    EX("Have a nice trip!", "祝旅途愉快", { fontSize: 12 }),
    EX("Thanks, Mom. See you next Saturday.", "謝謝媽媽，下週六見", { fontSize: 12 }),
    GAP(0.3),
    P("其他說法：", { fontSize: 11 }),
    EX("Enjoy your travels!", "祝玩得開心", { fontSize: 11 }),
    EX("Bon voyage!", "一路順風", { fontSize: 11 }),
  ], fs: 12 });
  card(s, { x: 6.6, y: 1.05, w: 2.9, h: 4.05, label: "祝成功", body: [
    EX("I hope you succeed.", "祝你成功", { fontSize: 12 }),
    EX("Thank you very much.", "非常感謝", { fontSize: 12 }),
    GAP(0.3),
    P("其他說法：", { fontSize: 11 }),
    EX("I'll pray for your success.", "我會為你祈求成功", { fontSize: 11 }),
    EX("I wish you all the success.", "祝你成功", { fontSize: 11 }),
  ], fs: 12 });
}

// =====================================================================
// 總結
// =====================================================================
{
  const s = base("四章總整理　一張表看完", null,
    "收尾用。四章其實只學了一件事：be 動詞句型，加上四個疑問詞。讓學生看著這張表，每一列自己造一句。");
  table(s, [
    ["疑問詞", "問什麼", "句型", "例句", "回答"],
    [{ text: "（無）", bold: true, color: NAVY }, "是不是", "be ＋ 主詞 ＋ 補語？", "Are you a student?", "Yes, I am. / No, I'm not."],
    [{ text: "Who", bold: true, color: RED }, "誰", "Who ＋ be ＋ 主詞？", "Who is she?", "She is my sister."],
    [{ text: "How", bold: true, color: RED }, "多～", "How ＋ 形容詞 ＋ be ＋ 主詞？", "How old is she?", "She is 15 years old."],
    [{ text: "What", bold: true, color: RED }, "什麼", "What ＋ be ＋ 主詞？", "What is this?", "This is a fan."],
    [{ text: "Where", bold: true, color: RED }, "在哪裡", "Where ＋ be ＋ 主詞？", "Where is he?", "He is at home."],
  ], { x: 0.5, y: 1.05, w: 9.0, colW: [1.0, 1.0, 2.6, 2.0, 2.4], fs: 12, rowH: 0.36 });
  table(s, [
    ["人稱", "主格", "所有格", "受格"],
    ["我", "I", "my", "me"],
    ["你", "you", "your", "you"],
    ["他 / 她 / 它", "he / she / it", "his / her / its", "him / her / it"],
    ["我們", "we", "our", "us"],
    ["他們", "they", "their", "them"],
  ], { x: 0.5, y: 3.5, w: 5.4, colW: [1.3, 1.4, 1.4, 1.3], fs: 10.5, rowH: 0.24 });
  card(s, { x: 6.2, y: 3.5, w: 3.3, h: 1.65, fill: REDTINT, label: "be 動詞", labelFill: RED, body: [
    P("I → am", { fontSize: 12, bold: true }),
    P("he / she / it / this / that → is", { fontSize: 12, bold: true }),
    P("you / we / they / these / those → are", { fontSize: 12, bold: true }),
  ], fs: 12 });
}

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT));
