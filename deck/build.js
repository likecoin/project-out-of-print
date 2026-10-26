const pptxgen = require("pptxgenjs");
const fs = require("fs");
// jszip ships with pptxgenjs; resolve it from there so this script needs no extra dependency.
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));

// pptxgenjs cannot write theme colours, so patch ppt/theme/theme1.xml after writeFile().
// Until this runs, every scheme colour resolves to Office's stock palette.
const THEME_SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2",
  "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

async function applyTheme(deckPath, theme) {
  const part = "ppt/theme/theme1.xml";
  const zip = await JSZip.loadAsync(fs.readFileSync(deckPath));
  if (!zip.file(part)) throw new Error(deckPath + " has no " + part);
  for (const k of THEME_SLOTS) {
    if (!/^[0-9A-Fa-f]{6}$/.test(String(theme.colors[k]))) {
      throw new Error("theme.colors." + k + " must be six hex digits without '#'");
    }
  }
  const name = String(theme.name).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const scheme = '<a:clrScheme name="' + name + '">'
    + THEME_SLOTS.map((k) => "<a:" + k + '><a:srgbClr val="'
        + String(theme.colors[k]).toUpperCase() + '"/></a:' + k + ">").join("")
    + "</a:clrScheme>";
  const xml = (await zip.file(part).async("string"))
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => head + name + '"');
  if (!xml.includes(scheme)) throw new Error(part + " has no <a:clrScheme> to replace");
  zip.file(part, xml);
  fs.writeFileSync(deckPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// Palette taken from the site: 3ook.com brand mint, ink-dark paper, warm cream.
const THEME = {
  name: "POOP",
  headFontFace: "Noto Sans TC",
  bodyFontFace: "Noto Sans TC",
  colors: {
    dk1: "16130F", // dark paper
    lt1: "FFFFFF",
    dk2: "0A7059", // deep teal (brand mint darkened for text on light)
    lt2: "F5ECE0", // warm cream card tint
    accent1: "50E3C2", // 3ook.com brand mint
    accent2: "0A7059",
    accent3: "F2EADD",
    accent4: "7A6F64",
    accent5: "1C1813",
    accent6: "1D1A17",
    hlink: "0A7059",
    folHlink: "7A6F64",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "3ook.com 圖書館";
pres.title = "絕版重生計劃";
pres.subject = "POOP: Project Out-Of-Print";

const C = pres.SchemeColor;
const M = 0.75;            // side margin
const W = 13.33 - M * 2;   // content width

/* ---------------- layouts ---------------- */

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: THEME.colors.dk1 },
  objects: [
    { placeholder: { options: { align: "left", name: "kicker", type: "body", x: M, y: 2.15, w: W, h: 0.4,
      fontSize: 14, bold: true, charSpacing: 3, color: THEME.colors.accent1 }, text: "" } },
    { placeholder: { options: { align: "left", name: "title", type: "title", x: M, y: 2.6, w: W, h: 1.3,
      fontSize: 54, bold: true, color: THEME.colors.accent3 }, text: "" } },
    { placeholder: { options: { align: "left", name: "body", type: "body", x: M, y: 4.0, w: W, h: 1.0,
      fontSize: 22, color: "CFC3B2" }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT_LIGHT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { align: "left", name: "title", type: "title", x: M, y: 0.52, w: W, h: 0.85,
      fontSize: 36, bold: true, color: THEME.colors.accent6 }, text: "" } },
    { placeholder: { options: { align: "left", name: "body", type: "body", x: M, y: 1.34, w: W, h: 0.5,
      fontSize: 15, color: THEME.colors.accent4 }, text: "" } },
    { text: { text: "poop.3ook.com", options: { x: M, y: 6.85, w: 4, h: 0.3, margin: 0,
      fontSize: 10, color: "A89C8E" } } },
  ],
  slideNumber: { x: 12.5, y: 6.85, w: 0.5, h: 0.3, fontSize: 10, color: "A89C8E", align: "right" },
});

pres.defineSlideMaster({
  title: "CONTENT_DARK",
  background: { color: THEME.colors.dk1 },
  objects: [
    { placeholder: { options: { align: "left", name: "title", type: "title", x: M, y: 0.52, w: W, h: 0.85,
      fontSize: 36, bold: true, color: THEME.colors.accent3 }, text: "" } },
    { placeholder: { options: { align: "left", name: "body", type: "body", x: M, y: 1.34, w: W, h: 0.5,
      fontSize: 15, color: "9A8D7D" }, text: "" } },
    { text: { text: "poop.3ook.com", options: { x: M, y: 6.85, w: 4, h: 0.3, margin: 0,
      fontSize: 10, color: "6E6356" } } },
    { slideNumber: { x: 12.5, y: 6.85, w: 0.5, h: 0.3, fontSize: 10, color: "6E6356", align: "right" } },
  ],
});

pres.defineSlideMaster({
  title: "STATEMENT_DARK",
  background: { color: THEME.colors.dk1 },
  objects: [
    { placeholder: { options: { align: "left", name: "title", type: "title", x: M, y: 2.5, w: W, h: 2.0,
      fontSize: 40, bold: true, color: THEME.colors.accent3, lineSpacingMultiple: 1.5 }, text: "" } },
  ],
});

/* ---------------- helpers ---------------- */

function card(slide, { x, y, w, h, dark = false, name }) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12,
    fill: { color: dark ? THEME.colors.accent5 : THEME.colors.lt2 },
    line: { color: dark ? "332C24" : "E8DCCB", width: 1 },
    objectName: name,
  });
}

function badge(slide, { x, y, label, name }) {
  slide.addShape(pres.shapes.OVAL, {
    x, y, w: 0.52, h: 0.52,
    fill: { color: THEME.colors.accent1 }, line: { color: THEME.colors.accent1, width: 1 },
    objectName: name,
  });
  slide.addText(label, {
    x, y, w: 0.52, h: 0.52, margin: 0, align: "center", valign: "middle",
    fontSize: 15, bold: true, color: "06241E", isTextBox: true,
  });
}

/* ---------------- 1. title ---------------- */

pres.addSection({ title: "開場" });
let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "開場" });
s.addText("POOP · PROJECT OUT-OF-PRINT", { placeholder: "kicker" });
s.addText("絕版重生計劃", { placeholder: "title" });
s.addText("好書不該被遺忘，讓絕版重生，讓思想流通。", { placeholder: "body" });
s.addText("由 3ook.com 圖書館發起　|　現正招募社會大眾參與　|　poop.3ook.com", {
  x: M, y: 5.3, w: W, h: 0.4, margin: 0, fontSize: 14, color: "9A8D7D", isTextBox: true,
});
s.addNotes("絕版重生計劃，由 3ook.com 圖書館發起，現正招募社會大眾參與。網址 poop.3ook.com。");

/* ---------------- 2. 理念 ---------------- */

s = pres.addSlide({ masterName: "STATEMENT_DARK", sectionTitle: "開場" });
s.addText(
  [
    { text: "好書不該被遺忘，\n", options: { color: THEME.colors.accent3 } },
    { text: "讓絕版重生，讓思想流通。", options: { color: THEME.colors.accent1 } },
  ],
  { placeholder: "title" }
);
s.addText("計劃理念", {
  x: M, y: 1.9, w: W, h: 0.4, margin: 0, fontSize: 14, bold: true, charSpacing: 3,
  color: THEME.colors.accent1, isTextBox: true,
});
s.addNotes("這是整個計劃的核心理念。");

/* ---------------- 3. 為甚麼 ---------------- */

pres.addSection({ title: "為甚麼" });
s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "為甚麼" });
s.addText("為甚麼要做這件事", { placeholder: "title" });
s.addText("絕版，不只是書店缺貨，而是一本書逐漸從公共記憶中消失", { placeholder: "body" });

const reasons = [
  ["壹", "紙本書正在被銷毀", "Anthropic 等公司大量購入二手書餵飼 AI，破壞性掃描然後銷毀。書的內容進入了模型，實體卻不復存在。"],
  ["貳", "圖書館收藏有限", "公共圖書館只收藏部分書籍，很多作品在絕版後難以被公眾接觸到。"],
  ["參", "絕版是常態，不是例外", "絕大部分書籍，本來就會隨著需求下降，踏入絕版的階段。"],
];
reasons.forEach(([num, head, text], i) => {
  const x = M + i * (3.94 + 0.25);
  card(s, { x, y: 2.1, w: 3.94, h: 3.6, name: `reason-${i}` });
  badge(s, { x: x + 0.38, y: 2.48, label: num, name: `reason-badge-${i}` });
  s.addText(head, {
    x: x + 0.38, y: 3.12, w: 3.2, h: 0.55, margin: 0, fontSize: 19, bold: true,
    color: THEME.colors.accent6, isTextBox: true,
  });
  s.addText(text, {
    x: x + 0.38, y: 3.72, w: 3.2, h: 1.75, margin: 0, fontSize: 13.5, lineSpacingMultiple: 1.35,
    color: "4A423A", isTextBox: true,
  });
});
s.addNotes("三個背景：破壞性掃描後銷毀、圖書館收藏有限、絕版是常態。");

/* ---------------- 4. 誰發起 ---------------- */

pres.addSection({ title: "誰在做" });
s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "誰在做" });
s.addText("由 3ook.com 圖書館發起", { placeholder: "title" });
s.addText("民營的電子圖書館，現正招募社會大眾參與", { placeholder: "body" });

card(s, { x: M, y: 2.1, w: 6.2, h: 3.6, name: "lib-desc" });
s.addText("像電子書的 Netflix", {
  x: M + 0.45, y: 2.5, w: 5.3, h: 0.5, margin: 0, fontSize: 22, bold: true,
  color: THEME.colors.dk2, isTextBox: true,
});
s.addText(
  "3ook.com 圖書館是民營的電子圖書館，收藏數千本香港、台灣書籍，持續增加中。\n\n讀者可透過 iPhone、Android app 或網站使用，訂閱後無限閱聽。",
  { x: M + 0.45, y: 3.15, w: 5.3, h: 2.2, margin: 0, fontSize: 14.5, lineSpacingMultiple: 1.4,
    color: "4A423A", isTextBox: true }
);

const stats = [
  ["數千本", "香港、台灣書籍，持續增加中"],
  ["無限閱聽", "訂閱後可閱聽所有作品"],
  ["按時數回饋", "訂閱收益按閱聽時數回饋創作者"],
];
stats.forEach(([big, small], i) => {
  const y = 2.1 + i * 1.25;
  card(s, { x: M + 6.45, y, w: 5.38, h: 1.1, name: `stat-${i}` });
  s.addText(big, {
    x: M + 6.85, y: y + 0.17, w: 2.1, h: 0.45, margin: 0, fontSize: 20, bold: true,
    color: THEME.colors.dk2, isTextBox: true,
  });
  s.addText(small, {
    x: M + 6.85, y: y + 0.6, w: 4.5, h: 0.4, margin: 0, fontSize: 12.5,
    color: "4A423A", isTextBox: true,
  });
});
s.addNotes("3ook.com 圖書館是計劃的發起方，也是後續配套。");

/* ---------------- 5. 範圍 ---------------- */

pres.addSection({ title: "計劃內容" });
s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "計劃內容" });
s.addText("計劃涵蓋甚麼", { placeholder: "title" });
s.addText("不限語言及地域，與電子書店上架沒有必然關係", { placeholder: "body" });

const scopes = [
  ["不只香港", "雖然香港的確是絕版書的重災區，但計劃不限定語言及地域，而以華文書籍優先，尤其台灣及香港出版的作品。"],
  ["不一定要上架", "3ook.com 圖書館是「絕版重生」計劃的後續配套，但兩者沒有必然關係。版權持有人可自行處理，也可以交由 3ook.com 代勞。"],
];
scopes.forEach(([head, text], i) => {
  const x = M + i * (5.79 + 0.25);
  card(s, { x, y: 2.1, w: 5.79, h: 2.5, name: `scope-${i}` });
  s.addText(head, {
    x: x + 0.45, y: 2.45, w: 4.9, h: 0.5, margin: 0, fontSize: 21, bold: true,
    color: THEME.colors.accent6, isTextBox: true,
  });
  s.addText(text, {
    x: x + 0.45, y: 3.05, w: 4.9, h: 1.35, margin: 0, fontSize: 14, lineSpacingMultiple: 1.4,
    color: "4A423A", isTextBox: true,
  });
});

s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: M, y: 4.95, w: W, h: 0.95, rectRadius: 0.12,
  fill: { color: "E7F8F3" }, line: { color: "BDEBDF", width: 1 }, objectName: "sensitive-note",
});
s.addText(
  [
    { text: "敏感書？　", options: { bold: true, color: THEME.colors.dk2 } },
    { text: "計劃的重點在協助製作電子書。製作後是否上架由版權持有人決定；是否容許上架，由各電子書店按當地法規決定。", options: { color: "35544C" } },
  ],
  { x: M + 0.45, y: 5.12, w: W - 0.9, h: 0.6, margin: 0, fontSize: 14, isTextBox: true }
);
s.addNotes("計劃以華文書籍優先，尤其台灣及香港。製作與上架是兩件事。");

/* ---------------- 6. 版權 ---------------- */

s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "計劃內容" });
s.addText("版權仍然屬於版權持有人", { placeholder: "title" });
s.addText("參與計劃並不影響版權歸屬", { placeholder: "body" });

const rights = [
  ["版權誰屬", "原版權持有人保有電子書的版權，可自由使用，包括私下分發或自行上架各電子書平台。"],
  ["要出版社同意嗎", "視乎合約條文及是否過期。台灣及香港典型合約為期三至五年，如包括電子版，期間內應由出版社參與。"],
  ["合約期滿之後", "一般而言文本的出版權回歸作者；編輯、排版、封面設計等版權仍屬出版社，可協商取得，或由義工重新製作。"],
];
rights.forEach(([head, text], i) => {
  const y = 2.1 + i * 1.5;
  card(s, { x: M, y, w: W, h: 1.32, name: `right-${i}` });
  badge(s, { x: M + 0.42, y: y + 0.4, label: String(i + 1), name: `right-badge-${i}` });
  s.addText(head, {
    x: M + 1.2, y: y + 0.22, w: 3.0, h: 0.45, margin: 0, fontSize: 18, bold: true,
    color: THEME.colors.accent6, isTextBox: true,
  });
  s.addText(text, {
    x: M + 1.2, y: y + 0.68, w: 10.2, h: 0.5, margin: 0, fontSize: 13.5,
    color: "4A423A", isTextBox: true,
  });
});
s.addNotes("版權歸屬不變，這是版權持有人最常問的一點。");

/* ---------------- 7. 材料 ---------------- */

s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "計劃內容" });
s.addText("版權持有人需要提供甚麼", { placeholder: "title" });
s.addText("由最理想到最低限度，三種情況都能處理", { placeholder: "body" });

const mats = [
  ["最理想", "電子檔", "如 Word、InDesign、PDF，圖片另外提供。"],
  ["其次", "紙本書", "由義工逆向製作電子版。"],
  ["都沒有", "義工協助搜尋", "如果連紙本書都無法提供，義工會盡力協助搜尋。"],
];
mats.forEach(([tag, head, text], i) => {
  const x = M + i * (3.94 + 0.25);
  card(s, { x, y: 2.2, w: 3.94, h: 3.1, name: `mat-${i}` });
  s.addText(tag, {
    x: x + 0.38, y: 2.55, w: 3.2, h: 0.35, margin: 0, fontSize: 12, bold: true,
    charSpacing: 2, color: THEME.colors.dk2, isTextBox: true,
  });
  s.addText(head, {
    x: x + 0.38, y: 2.95, w: 3.2, h: 0.55, margin: 0, fontSize: 23, bold: true,
    color: THEME.colors.accent6, isTextBox: true,
  });
  s.addText(text, {
    x: x + 0.38, y: 3.6, w: 3.2, h: 1.3, margin: 0, fontSize: 14, lineSpacingMultiple: 1.35,
    color: "4A423A", isTextBox: true,
  });
  if (i < 2) {
    s.addText("→", {
      x: x + 3.94, y: 3.45, w: 0.25, h: 0.4, margin: 0, align: "center",
      fontSize: 18, bold: true, color: "C9BCA9", isTextBox: true,
    });
  }
});
s.addNotes("三種情況：有電子檔最好，只有紙本也可以，甚麼都沒有我們幫忙找。");

/* ---------------- 8. 費用與營運 ---------------- */

s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "計劃內容" });
s.addText("免費參加，那如何持續營運", { placeholder: "title" });
s.addText("為推動文化保育，計劃免費參加，但歡迎參加者自由定價支持", { placeholder: "body" });

s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: M, y: 2.2, w: 4.3, h: 3.1, rectRadius: 0.12,
  fill: { color: THEME.colors.dk2 }, line: { color: THEME.colors.dk2, width: 1 },
  objectName: "free-callout",
});
s.addText("免費", {
  x: M + 0.45, y: 2.75, w: 3.4, h: 1.0, margin: 0, fontSize: 60, bold: true,
  color: "FFFFFF", isTextBox: true,
});
s.addText("參加計劃不收費，歡迎自由定價支持。", {
  x: M + 0.45, y: 3.85, w: 3.4, h: 1.0, margin: 0, fontSize: 14.5, lineSpacingMultiple: 1.35,
  color: "CFEFE6", isTextBox: true,
});

const ops = [
  ["讓書籍持續流通", "我們認為最健康的模式是讓書籍持續流通，讓讀者接觸並付費支持。訂閱 3ook.com 圖書館，即可支持絕版重生，並獲得閱聽所有作品的權利。"],
  ["招募義工", "除了訂閱支持，若你願意付出時間，歡迎報名成為計劃的義工。"],
];
ops.forEach(([head, text], i) => {
  const y = 2.2 + i * 1.65;
  card(s, { x: M + 4.55, y, w: 7.28, h: 1.45, name: `ops-${i}` });
  s.addText(head, {
    x: M + 5.0, y: y + 0.2, w: 6.4, h: 0.4, margin: 0, fontSize: 18, bold: true,
    color: THEME.colors.accent6, isTextBox: true,
  });
  s.addText(text, {
    x: M + 5.0, y: y + 0.62, w: 6.4, h: 0.7, margin: 0, fontSize: 13.5, lineSpacingMultiple: 1.3,
    color: "4A423A", isTextBox: true,
  });
});
s.addNotes("免費參加；營運靠訂閱收益與義工。");

/* ---------------- 9. 書目 ---------------- */

pres.addSection({ title: "成果" });
s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "成果" });
s.addText("已經重生的作品", { placeholder: "title" });
s.addText("部分已經或正在重生的作品", { placeholder: "body" });

const books = [
  ["歐洲 12 國 16 天遊", "馬國明"],
  ["〇〇年代：社運結束，耕田繼續", "梁景鴻"],
  ["走過火紅的傘下銀髮", "楊寶熙"],
  ["七千零七夜：戀戀書廊", "趙良駿"],
  ["北韓迷宮", "薯伯伯"],
  ["大事件：記香港 WTO 採訪", "鄭依依等"],
  ["火紅芝大", "崔少明"],
  ["馬聞", "馬聞義工組"],
  ["Porcelain 白瓷（製作中）", "李智良"],
];
books.forEach(([t, a], i) => {
  const col = i % 3, row = Math.floor(i / 3);
  const x = M + col * (3.94 + 0.25);
  const y = 2.1 + row * 1.42;
  card(s, { x, y, w: 3.94, h: 1.22, name: `book-${i}` });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.3, y: y + 0.3, w: 0.09, h: 0.62, rectRadius: 0.04,
    fill: { color: THEME.colors.accent1 }, line: { color: THEME.colors.accent1, width: 1 },
    objectName: `book-mark-${i}`,
  });
  s.addText(t, {
    x: x + 0.55, y: y + 0.26, w: 3.1, h: 0.52, margin: 0, fontSize: 14, bold: true,
    color: THEME.colors.accent6, isTextBox: true,
  });
  s.addText(a, {
    x: x + 0.55, y: y + 0.76, w: 3.1, h: 0.3, margin: 0, fontSize: 12,
    color: "7A6F64", isTextBox: true,
  });
});
s.addNotes("這些只是其中一部分，完整書目見 poop.3ook.com。");

/* ---------------- 10. 參與 ---------------- */

pres.addSection({ title: "參與" });
s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "參與" });
s.addText("三種參與方式", { placeholder: "title" });
s.addText("免費參加，歡迎自由定價支持計劃", { placeholder: "body" });

const joins = [
  ["我是版權持有人", "作者、譯者、出版社、或已故作者的家屬——只要手上有一本值得重生的絕版書，都歡迎聯絡我們。", "poop@3ook.com"],
  ["我想當義工", "校對、排版、掃描、修圖、電子書製作、聯絡版權持有人——每一本重生的書，背後都是一班義工。", "poop@3ook.com"],
  ["我想支持計劃", "訂閱 3ook.com 圖書館，支持絕版重生，無限閱聽數千本香港、台灣書籍。", "3ook.com"],
];
joins.forEach(([head, text, cta], i) => {
  const x = M + i * (3.94 + 0.25);
  card(s, { x, y: 2.1, w: 3.94, h: 3.5, dark: true, name: `join-${i}` });
  s.addText(head, {
    x: x + 0.38, y: 2.45, w: 3.2, h: 0.5, margin: 0, fontSize: 20, bold: true,
    color: THEME.colors.accent1, isTextBox: true,
  });
  s.addText(text, {
    x: x + 0.38, y: 3.05, w: 3.2, h: 1.7, margin: 0, fontSize: 13.5, lineSpacingMultiple: 1.35,
    color: "CFC3B2", isTextBox: true,
  });
  s.addText(cta, {
    x: x + 0.38, y: 4.9, w: 3.2, h: 0.4, margin: 0, fontSize: 14, bold: true,
    color: "F2EADD", isTextBox: true,
  });
});
s.addNotes("版權持有人與義工都是 poop@3ook.com；支持計劃請訂閱 3ook.com。");

/* ---------------- 11. 結尾 ---------------- */

s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "參與" });
s.addText("POOP · PROJECT OUT-OF-PRINT", { placeholder: "kicker" });
s.addText("讓絕版重生", { placeholder: "title" });
s.addText("好書不該被遺忘，讓絕版重生，讓思想流通。", { placeholder: "body" });
s.addText("poop.3ook.com　|　poop@3ook.com", {
  x: M, y: 5.3, w: W, h: 0.45, margin: 0, fontSize: 18, bold: true,
  color: THEME.colors.accent1, isTextBox: true,
});
s.addNotes("歡迎聯絡 poop@3ook.com，或到 poop.3ook.com 了解更多。");

/* ---------------- write ---------------- */

(async () => {
  const out = __dirname + "/poop-deck.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote " + out);
})();
