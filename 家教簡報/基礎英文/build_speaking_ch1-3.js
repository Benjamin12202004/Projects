const pptxgen = require("pptxgenjs");
const path = require("path");

const OUT = process.argv[2] || path.join(__dirname, "基礎英文_Ch01-03_口說練習.pptx");

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
pres.title = "基礎英文 Chapter 01–03 口說練習";
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


// =====================================================================
// 口說練習專用 helpers
// =====================================================================
function activity(num, title, chapter, minutes, notes) {
  const s = base(`${num}　${title}`, null, notes);
  pill(s, `Ch ${chapter}　${minutes} 分鐘`, 7.95, 0.42, 1.55, 0.34, RED, WHITE, 11);
  return s;
}
// 句型框：大字、留空
function frames(s, lines, box, fs = 18) {
  text(s, lines.map((l) => R(l, { bold: true, color: NAVY, fontSize: fs })), { paraSpaceAfter: 8, ...box });
}
function wordbank(s, title, words, box, fs = 12) {
  card(s, { ...box, label: title, labelW: Math.max(0.9, title.length * 0.22 + 0.3), body: [P(words.join("　"), { fontSize: fs })], fs });
}
function dialog(s, lines, box, fs = 14) {
  const body = [];
  lines.forEach(([who, en, zh]) => {
    body.push([
      { text: who + "　", options: { bold: true, color: RED, breakLine: false, fontSize: fs - 2 } },
      { text: en, options: { bold: true, color: NAVY, breakLine: true, fontSize: fs } },
    ]);
    if (zh) body.push(M("　　　" + zh, { fontSize: fs - 3 }));
  });
  text(s, body, { paraSpaceAfter: 2, ...box });
}

// =====================================================================
// 封面 + 使用說明
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("口說練習", { x: 0.7, y: 1.3, w: 8.6, h: 1.1, fontFace: F, fontSize: 48, bold: true, color: WHITE, margin: 0, isTextBox: true, valign: "middle" });
  s.addText("Chapter 01 – 03　把今天學的講出來", { x: 0.7, y: 2.45, w: 8.6, h: 0.6, fontFace: F, fontSize: 24, color: "D6E0F0", margin: 0, isTextBox: true });
  s.addText("自我介紹 · 快問快答 · 角色扮演 · 介紹家人 · 形容一個人", { x: 0.7, y: 3.2, w: 8.6, h: 0.5, fontFace: F, fontSize: 15, color: "9FB4D6", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 4.2, w: 8.6, h: 0.75, fill: { color: NAVY2 }, line: { color: NAVY2 }, rectRadius: 0.1 });
  s.addText("規則：先講再說，講錯沒關係，老師最後才糾正", { x: 0.95, y: 4.2, w: 8.2, h: 0.75, fontFace: F, fontSize: 15, color: WHITE, margin: 0, isTextBox: true, valign: "middle" });
  s.addNotes("口說課開場。跟學生說三件事：① 今天不寫，只講 ② 講錯沒關係，我最後才糾正 ③ 每個活動只用今天學過的句子。");
}
{
  const s = base("老師用　怎麼跑這一節", null,
    "這頁學生不用看。整節約 45 分鐘，時間不夠就跳活動 4、6、9（最難的三個）。每個活動的備註欄都有範例答案和常見錯誤。");
  pill(s, "老師用", 8.25, 0.42, 1.25, 0.34, RED, WHITE, 11);
  table(s, [
    ["#", "活動", "對應", "分鐘", "型態"],
    ["1", "自我介紹（暖身）", "Ch 1", "5", "獨白"],
    ["2", "快問快答：Yes / No", "Ch 1", "5", "一問一答"],
    ["3", "快問快答：Who / How / What", "Ch 1–3", "5", "一問一答"],
    ["4", "角色扮演：初次見面", "Ch 1", "5", "對話"],
    ["5", "介紹家人", "Ch 2", "5", "獨白"],
    ["6", "猜猜看 Who is this?", "Ch 2", "5", "描述＋猜"],
    ["7", "桌上的東西 What is this?", "Ch 3", "5", "一問一答"],
    ["8", "角色扮演：送禮物", "Ch 3", "4", "對話"],
    ["9", "形容一個人", "Ch 3", "5", "獨白"],
    ["10", "反應練習：這時候說什麼", "Ch 1–3", "4", "情境反應"],
    ["11", "一分鐘自我介紹（總結）", "Ch 1–3", "5", "獨白"],
  ], { x: 0.5, y: 1.05, w: 6.6, colW: [0.4, 3.2, 0.9, 0.7, 1.4], fs: 11, rowH: 0.3 });
  card(s, { x: 7.4, y: 1.05, w: 2.1, h: 4.1, fill: REDTINT, label: "原則", labelFill: RED, body: [
    P("學生講的量 ＞ 老師", { fontSize: 11 }),
    P("每個活動老師先示範一次", { fontSize: 11 }),
    P("學生卡住：給第一個字，不給整句", { fontSize: 11 }),
    P("糾錯留到活動結束，一次只糾一個", { fontSize: 11 }),
    P("時間不夠：跳 4、6、9", { fontSize: 11 }),
  ], fs: 11 });
}

// 1 自我介紹
{
  const s = activity("1", "自我介紹（暖身）", "1", 5,
    "老師先示範：Hi, I'm Ben. I'm from Taiwan. I'm Taiwanese. I'm a student. I'm 21 years old. Nice to meet you. 然後學生照框講一次，第二次不看簡報再講一次。常見錯誤：I'm come from（✗）→ I'm from；I am student（✗）→ I am a student。");
  text(s, [P("照著框講，一句一句來。講完第二次不看螢幕再講一次")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  frames(s, [
    "Hi, I'm ________ .",
    "I'm from ________ .　　I'm ________ .（國籍）",
    "I'm a ________ .（身分／職業）",
    "I'm ________ years old.",
    "Nice to meet you.",
  ], { x: 0.5, y: 1.5, w: 5.8, h: 3.3 });
  wordbank(s, "國家／國籍", ["Taiwan / Taiwanese", "Japan / Japanese", "Korea / Korean", "the U.S.A. / American"], { x: 6.6, y: 1.5, w: 2.9, h: 1.55 }, 11);
  wordbank(s, "身分", ["student", "teacher", "doctor", "cook", "police officer"], { x: 6.6, y: 3.25, w: 2.9, h: 1.55 }, 11);
}

// 2 快問快答 Yes/No
{
  const s = activity("2", "快問快答　Yes / No", "1", 5,
    "老師唸問題，學生只能用 Yes, I am. / No, I'm not. / Yes, it is. / No, it isn't. 回答，然後補一句。節奏要快。回答完可以互換：學生問老師。範例：Are you hungry? → No, I'm not. I'm full.（full 沒教過，學生說 No, I'm not. 就夠了）");
  text(s, [P("老師問，學生用 Yes / No 開頭回答，再多講一句")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  formula(s, ["be 動詞", "主詞", "補語 ?"], 0.5, 1.45, 4.3, 0.45, { fill: RED });
  text(s, [
    EX("Are you a student?", "", { fontSize: 14 }),
    EX("Are you hungry?", "", { fontSize: 14 }),
    EX("Are you from Taipei?", "", { fontSize: 14 }),
    EX("Are you tall?", "", { fontSize: 14 }),
    EX("Is your room big?", "", { fontSize: 14 }),
    EX("Is your phone new?", "", { fontSize: 14 }),
    EX("Is your mom a teacher?", "", { fontSize: 14 }),
    EX("Are your friends Taiwanese?", "", { fontSize: 14 }),
    EX("Is it cold today?", "", { fontSize: 14 }),
    EX("Are you happy today?", "", { fontSize: 14 }),
  ], { x: 0.5, y: 2.05, w: 4.3, h: 3.1, fontSize: 14, paraSpaceAfter: 1 });
  card(s, { x: 5.2, y: 1.45, w: 4.3, h: 2.1, label: "回答", body: [
    EX("Yes, I am.　/　No, I'm not.", "", { fontSize: 13 }),
    EX("Yes, it is.　/　No, it isn't.", "", { fontSize: 13 }),
    EX("Yes, she is.　/　No, she isn't.", "", { fontSize: 13 }),
    EX("Yes, they are.　/　No, they aren't.", "", { fontSize: 13 }),
  ], fs: 13 });
  card(s, { x: 5.2, y: 3.7, w: 4.3, h: 1.45, fill: REDTINT, label: "再多講一句", labelFill: RED, labelW: 1.1, body: [
    EX("No, I'm not. I'm from Taichung.", "", { fontSize: 12 }),
    EX("Yes, it is. It's really big.", "", { fontSize: 12 }),
  ], fs: 12 });
}

// 3 快問快答 Wh
{
  const s = activity("3", "快問快答　Who / How / What", "1–3", 5,
    "三個疑問詞加 this / these 的 Yes / No 問句混著問。學生回答要用完整句。問到 What is this? 時老師真的拿東西到鏡頭前。第二輪讓學生問老師，老師故意講錯一次看學生會不會發現。");
  text(s, [P("回答要用完整的句子，不能只講一個字")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  const qs = [
    ["Who", NAVY, ["Who is your best friend?", "Who is your English teacher?", "Who is this?（指照片）"]],
    ["How", RED, ["How old are you?", "How tall are you?", "How are you today?"]],
    ["What", NAVY, ["What is this?（老師拿東西）", "What is your name?", "What is your mom's name?"]],
    ["this / these", RED, ["Is this your pen?（老師指）", "Are these your books?", "Is that your sister?（指照片）"]],
  ];
  qs.forEach(([w, c, list], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.45 + Math.floor(i / 2) * 1.85;
    card(s, { x, y, w: 4.4, h: 1.7, fill: c === RED ? REDTINT : TINT, label: w, labelFill: c, labelW: w.length > 6 ? 1.4 : 0.9,
      body: list.map((q) => R(q, { bold: true, color: NAVY, fontSize: 13 })), fs: 13 });
  });
}

// 4 角色扮演 初次見面
{
  const s = activity("4", "角色扮演　初次見面", "1", 5,
    "老師先當 A、學生當 B 跑一次，再互換。第二輪把國家換掉（學生假裝是日本人／韓國人／加拿大人）。重點：Nice to meet you, too. 的 too；I'm a doctor. 的 a；No, I'm not. 後面要補正確答案。");
  text(s, [P("A 先問，B 回答。跑完一次互換，第二輪換一個國家")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  dialog(s, [
    ["A", "Hi, I'm ________ .", ""],
    ["B", "Hi, I'm ________ . Nice to meet you.", ""],
    ["A", "Nice to meet you, too. Are you a student?", ""],
    ["B", "Yes, I am. / No, I'm not. I'm a ________ . How about you?", ""],
    ["A", "I'm a ________ . Where are you from?", ""],
    ["B", "I'm from ________ . I'm ________ . Are you ________ , too?", ""],
    ["A", "Yes, I am. / No, I'm not. I'm ________ .", ""],
    ["B", "Good morning / afternoon / evening! See you later.", ""],
    ["A", "Bye. Take care.", ""],
  ], { x: 0.5, y: 1.45, w: 6.0, h: 3.7 }, 14);
  wordbank(s, "假裝你來自", ["Japan / Japanese", "Korea / Korean", "Canada / Canadian", "Italy / Italian", "France / French"], { x: 6.8, y: 1.45, w: 2.7, h: 1.8 }, 11);
  wordbank(s, "職業", ["student", "doctor", "cook", "painter", "farmer", "police officer"], { x: 6.8, y: 3.4, w: 2.7, h: 1.75 }, 11);
}

// 5 介紹家人
{
  const s = activity("5", "介紹家人", "2", 5,
    "學生選三個家人，每人講三句：是誰、幾歲（或 young / old）、怎麼樣。老師示範：This is my sister. She is 25 years old. She is tall and kind. 注意 he / she 不要混、brother 不分哥哥弟弟。進階：講到第三個人時加 and / but：She is kind, but she is serious.");
  text(s, [P("選三個家人，每個人講三句。可以真的拿照片給鏡頭看")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  frames(s, [
    "This is my ________ .",
    "He / She is ________ years old.",
    "He / She is ________ and ________ .",
    "His / Her name is ________ .",
  ], { x: 0.5, y: 1.5, w: 5.6, h: 2.4 });
  card(s, { x: 0.5, y: 3.95, w: 5.6, h: 1.2, fill: REDTINT, label: "進階", labelFill: RED, body: [
    EX("She is kind, but she is serious.", "用 but 轉折", { fontSize: 13 }),
    EX("My dad and my mom are teachers.", "兩個人用 are", { fontSize: 13 }),
  ], fs: 13 });
  wordbank(s, "家人", ["father / dad", "mother / mom", "brother", "sister", "grandpa", "grandma", "uncle", "aunt", "cousin"], { x: 6.4, y: 1.5, w: 3.1, h: 1.7 }, 11);
  wordbank(s, "形容詞", ["tall", "young", "old", "kind", "cute", "pretty", "smart", "cheerful", "serious", "healthy", "busy"], { x: 6.4, y: 3.35, w: 3.1, h: 1.8 }, 11);
}

// 6 猜猜看
{
  const s = activity("6", "猜猜看　Who is this?", "2", 5,
    "老師先出題：He is a man. He is from Taiwan. He is tall. He is a basketball player. He is really famous. Who is this? 學生猜。接著學生出題，老師猜。限制：只能用 be 動詞句。學生卡住就給他「職業、國家、外表、年紀」四個方向。");
  text(s, [P("一個人描述，一個人猜。只能用 be 動詞的句子，講 4–5 句")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  frames(s, [
    "He / She is a man / woman.",
    "He / She is from ________ .",
    "He / She is ________ .（外表）",
    "He / She is a ________ .（職業）",
    "He / She is ________ years old.",
    "Who is this?",
  ], { x: 0.5, y: 1.5, w: 5.4, h: 3.6 }, 17);
  card(s, { x: 6.2, y: 1.5, w: 3.3, h: 1.6, label: "老師示範", body: [
    P("He is a man. He is from Taiwan.", { fontSize: 11 }),
    P("He is tall. He is a singer.", { fontSize: 11 }),
    P("He is really famous. Who is this?", { fontSize: 11 }),
  ], fs: 11 });
  card(s, { x: 6.2, y: 3.25, w: 3.3, h: 1.9, fill: REDTINT, label: "猜的人可以問", labelFill: RED, labelW: 1.3, body: [
    EX("Is he a singer?", "", { fontSize: 12 }),
    EX("Is she from Korea?", "", { fontSize: 12 }),
    EX("How old is he?", "", { fontSize: 12 }),
    EX("Is it ________ ?", "", { fontSize: 12 }),
  ], fs: 12 });
}

// 7 What is this?
{
  const s = activity("7", "桌上的東西　What is this?", "3", 5,
    "學生拿桌上五樣東西到鏡頭前，自己問自己答：What is this? This is a pen. It is blue. 老師接著指遠的東西問 What is that? 逼學生用 that。複數：What are these? These are pens. 新單字不用查，不會講的東西直接說 I don't know.");
  text(s, [P("拿五樣東西到鏡頭前，近的用 this，遠的用 that，兩個以上用 these / those")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  table(s, [
    ["", "問", "答"],
    [{ text: "近・一個", bold: true, color: NAVY }, "What is this?", "This is a ________ . It is ________ ."],
    [{ text: "遠・一個", bold: true, color: NAVY }, "What is that?", "That is a ________ . It is ________ ."],
    [{ text: "近・很多", bold: true, color: NAVY }, "What are these?", "These are ________ s. They are ________ ."],
    [{ text: "遠・很多", bold: true, color: NAVY }, "What are those?", "Those are ________ s. They are ________ ."],
  ], { x: 0.5, y: 1.5, w: 9.0, colW: [1.4, 2.4, 5.2], fs: 13, rowH: 0.45 });
  wordbank(s, "桌上常見", ["pen", "book", "cup", "phone", "bag", "bottle", "notebook", "key", "glasses"], { x: 0.5, y: 3.9, w: 4.4, h: 1.25 }, 12);
  wordbank(s, "形容詞", ["big", "small", "new", "old", "pretty", "cute", "awesome", "really ________"], { x: 5.1, y: 3.9, w: 4.4, h: 1.25 }, 12);
}

// 8 送禮物
{
  const s = activity("8", "角色扮演　送禮物", "3", 4,
    "短對話，重點是受格 for you / from me 和回應感謝。跑兩次：第一次送一樣東西（It's a ... ），第二次送很多樣（These are ...）。學生當送的人時，提醒 It's a present for you. 的 for 後面是受格。");
  text(s, [P("A 送東西給 B。第一輪送一樣，第二輪送很多樣")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  dialog(s, [
    ["A", "This is a present for you.", "這是給你的禮物"],
    ["B", "Really? Thank you. What is it?", "真的嗎？謝謝。是什麼？"],
    ["A", "It's a ________ . It's from ________ .", "是＿＿。來自＿＿"],
    ["B", "Wow! It's really ________ . Thank you very much.", "哇！真的很＿＿。非常謝謝"],
    ["A", "You're welcome. I'm glad you like it.", "不客氣，很高興你喜歡"],
  ], { x: 0.5, y: 1.45, w: 5.8, h: 3.0 }, 15);
  card(s, { x: 6.6, y: 1.45, w: 2.9, h: 1.8, label: "受格提醒", labelW: 1.0, body: [
    EX("for you / for him / for her", "", { fontSize: 12 }),
    EX("from me / from us", "", { fontSize: 12 }),
    EX("I love it.", "it 當受詞", { fontSize: 12 }),
  ], fs: 12 });
  card(s, { x: 6.6, y: 3.4, w: 2.9, h: 1.75, fill: REDTINT, label: "回應感謝", labelFill: RED, body: [
    EX("You're welcome.", "", { fontSize: 12 }),
    EX("No problem.", "", { fontSize: 12 }),
    EX("My pleasure.", "", { fontSize: 12 }),
  ], fs: 12 });
}

// 9 形容一個人
{
  const s = activity("9", "形容一個人", "3", 5,
    "用 p.58 的形容詞。學生選一個朋友或老師，連講五句：外表兩句、個性兩句、今天的情緒一句。老師先示範。進階：最後一句用 but。常見錯誤：She is very kind person.（✗，kind 後面不用 person）；He is happy and sad.（語意矛盾，提醒用 but）。");
  text(s, [P("選一個朋友或家人，連講五句：外表、個性、今天的心情")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  frames(s, [
    "My friend is ________ .（名字）",
    "He / She is ________ and ________ .（外表）",
    "He / She is ________ .（個性）",
    "He / She is really ________ .",
    "Today he / she is ________ .（心情）",
    "He / She is ________ , but ________ .",
  ], { x: 0.5, y: 1.5, w: 5.4, h: 3.6 }, 17);
  wordbank(s, "外表", ["tall", "small", "slim", "pretty", "cute", "young", "old"], { x: 6.2, y: 1.5, w: 3.3, h: 1.1 }, 11);
  wordbank(s, "個性", ["nice", "kind", "cheerful", "serious", "smart", "rich", "healthy"], { x: 6.2, y: 2.75, w: 3.3, h: 1.1 }, 11);
  wordbank(s, "心情", ["happy", "excited", "proud", "confident", "sad", "worried", "lonely"], { x: 6.2, y: 4.0, w: 3.3, h: 1.15 }, 11);
}

// 11 反應練習
{
  const s = activity("10", "反應練習　這時候說什麼？", "1–3", 4,
    "老師唸情境（中文），學生馬上用英文反應，一句就好。答案不只一個。練到反射，不要想文法。可以第二輪加快速度。");
  text(s, [P("老師唸情境，學生馬上用一句英文反應。答案不只一個")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  table(s, [
    ["情境", "你可以說"],
    ["第一次見到新同學", "Hi, I'm ________ . Nice to meet you."],
    ["對方說 Nice to meet you.", "Nice to meet you, too."],
    ["朋友給你看他拍的照片，很漂亮", "Wow! How beautiful! / Fantastic!"],
    ["朋友送你禮物", "Really? Thank you. It's awesome."],
    ["朋友幫了你一個忙", "Thank you very much. / I appreciate your help."],
    ["朋友跟你說謝謝", "You're welcome. / No problem. / My pleasure."],
    ["把書遞給朋友", "Here is your book."],
    ["朋友今天看起來很難過", "You look down. What's wrong? / Are you okay?"],
    ["早上在走廊遇到老師", "Good morning!"],
    ["晚上要掛電話了", "Good night. See you later."],
    ["跟朋友道別", "Bye. Take care."],
  ], { x: 0.5, y: 1.5, w: 9.0, colW: [3.6, 5.4], fs: 12, rowH: 0.3 });
}

// 12 一分鐘自我介紹
{
  const s = activity("11", "一分鐘自我介紹（總結）", "1–3", 5,
    "把今天全部串起來。學生先看框講一次，老師計時。第二次把框收起來（切到下一頁）再講一次。講完老師只挑一個最常錯的地方糾正，其他稱讚。這段可以錄下來給學生當作業對照。");
  text(s, [P("八句話把今天學的全部用上。第一次看框，第二次不看")], { x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 13 });
  const lines = [
    ["Ch 1", "Hi, I'm ________ . I'm from ________ . I'm a ________ ."],
    ["Ch 1", "I'm ________ years old. I'm ________ and ________ ."],
    ["Ch 2", "This is my ________ . His / Her name is ________ ."],
    ["Ch 2", "He / She is ________ , but he / she is ________ ."],
    ["Ch 3", "This is my ________ . It's a present from ________ . I love it."],
    ["Ch 3", "These are my ________ s. They are really ________ ."],
    ["Ch 3", "Today I am ________ .（心情）"],
    ["Ch 1", "Nice to meet you. See you later!"],
  ];
  lines.forEach(([ch, l], i) => {
    const y = 1.5 + i * 0.44;
    pill(s, ch, 0.5, y + 0.06, 0.6, 0.3, i % 2 ? RED : NAVY, WHITE, 9);
    s.addText(l, { x: 1.25, y, w: 8.25, h: 0.42, fontFace: F, fontSize: 15, bold: true, color: NAVY, margin: 0, isTextBox: true, valign: "middle" });
  });
}
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("不看框，再講一次", { x: 0.7, y: 1.6, w: 8.6, h: 1.0, fontFace: F, fontSize: 40, bold: true, color: WHITE, margin: 0, isTextBox: true, valign: "middle" });
  s.addText("提示：名字 → 國家 → 身分 → 年紀 → 家人 → 禮物 → 心情 → 道別", { x: 0.7, y: 2.7, w: 8.6, h: 0.6, fontFace: F, fontSize: 18, color: "D6E0F0", margin: 0, isTextBox: true });
  s.addText("One minute. Go!", { x: 0.7, y: 3.6, w: 8.6, h: 0.6, fontFace: F, fontSize: 24, bold: true, color: "F2B8B5", margin: 0, isTextBox: true });
  s.addNotes("計時一分鐘。學生卡住只給提示詞（例如：family），不給整句。");
}

// 老師用：糾錯清單
{
  const s = base("老師用　今天最常見的錯誤", null,
    "學生講的時候在這張表上勾，活動結束一次講。每次只糾一到兩個，其他下次再說。");
  pill(s, "老師用", 8.25, 0.42, 1.25, 0.34, RED, WHITE, 11);
  table(s, [
    ["學生常講", "正確", "為什麼"],
    [{ text: "I am student.", color: RED }, "I am a student.", "職業、身分前面要 a / an"],
    [{ text: "She am / He are", color: RED }, "She is / He is", "he / she / it 配 is"],
    [{ text: "I'm come from Taiwan.", color: RED }, "I'm from Taiwan.", "be 動詞後面直接接 from"],
    [{ text: "He is my the brother.", color: RED }, "He is my brother.", "my 後面不加 the"],
    [{ text: "I love she.", color: RED }, "I love her.", "動詞後面用受格"],
    [{ text: "This is present for I.", color: RED }, "This is a present for me.", "介係詞後面用受格，別忘了 a"],
    [{ text: "Who she is?", color: RED }, "Who is she?", "疑問詞後面先 be 動詞再主詞"],
    [{ text: "How old is you?", color: RED }, "How old are you?", "you 配 are"],
    [{ text: "What this is?", color: RED }, "What is this?", "疑問詞後面先 be 動詞再主詞"],
    [{ text: "This is pens.", color: RED }, "These are pens.", "兩個以上用 these ＋ are"],
    [{ text: "These is my books.", color: RED }, "These are my books.", "these / those 配 are"],
    [{ text: "Nice to meet you too.（沒說）", color: RED }, "Nice to meet you, too.", "回應要加 too"],
  ], { x: 0.5, y: 1.05, w: 9.0, colW: [2.6, 2.6, 3.8], fs: 11, rowH: 0.3 });
}

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT));
