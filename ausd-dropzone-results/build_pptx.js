// Builds AUSD-Dropzone-Results.pptx (native, editable shapes and text).
// Run: node build_pptx.js   (from this folder)
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "AUSD-Dropzone-Results.pptx");

const THEME = {
  name: "KUB Dark",
  headFontFace: "Poppins",
  bodyFontFace: "Poppins",
  colors: {
    dk1: "0E1210", lt1: "FFFFFF", dk2: "1A2B1F", lt2: "B9C2BC",
    accent1: "3FD46B", accent2: "E6B43A", accent3: "1F8F45", accent4: "7F8A83",
    accent5: "3E8A57", accent6: "5F6A63", hlink: "3FD46B", folHlink: "E6B43A",
  },
};

const GREEN = "3FD46B", GOLD = "E6B43A", WHITE = "FFFFFF", MUTED = "B9C2BC",
  DIM = "7F8A83", CARD = "17271C", LINE = "3E8A57", INNER = "111915", INNER_LINE = "2A342E", TRACK = "323A35";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "AUSD Dropzone Campaign Results";

pres.defineSlideMaster({
  title: "KUB_STATS",
  background: { path: A("bg.png") },
  objects: [
    { image: { path: A("kub-logo.png"), x: 11.93, y: 0.42, w: 0.8, h: 0.3 } },
    { text: { text: "Copyright © 2026 Bitkub Capital Group Holdings Co., Ltd. All Rights Reserved.",
      options: { x: 7.73, y: 7.0, w: 5.0, h: 0.28, fontSize: 8, color: "5F6A63", align: "right", valign: "top", margin: 0, isTextBox: true } } },
  ],
});

const s = pres.addSlide({ masterName: "KUB_STATS" });
const T = (text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true, valign: "top" }, o));
const box = (x, y, w, h, fill, line, r = 0.16, name) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: r, fill: { color: fill }, line: { color: line, width: 1.25 }, objectName: name });
const rect = (x, y, w, h, color, rounded) =>
  s.addShape(rounded ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE,
    Object.assign({ x, y, w, h, fill: { color }, line: { type: "none" } }, rounded ? { rectRadius: h / 2 } : {}));

// ---- Header
T([{ text: "AUSD ", options: { color: GOLD } }, { text: "Dropzone Campaign Results", options: { color: WHITE } }],
  { x: 0.6, y: 0.36, w: 11, h: 0.75, fontSize: 38, bold: true, objectName: "Title" });
T([{ text: "Agora AUSD Dropzone on KUB Wallet  ·  Reward drop: ", options: { color: MUTED } },
   { text: "29 October 2026", options: { color: GOLD, bold: true } },
   { text: "  ·  1 AUSD ≈ 33.7 THB", options: { color: MUTED } }],
  { x: 0.6, y: 1.12, w: 11, h: 0.36, fontSize: 14 });

// ---- Section titles
T("Participants & Rewards", { x: 0.6, y: 1.7, w: 6, h: 0.4, fontSize: 17, bold: true, color: GREEN });
T("AUSD Held · 30-Day Period", { x: 7.13, y: 1.7, w: 5.6, h: 0.4, fontSize: 17, bold: true, color: GREEN });

// ---- LEFT CARD: participants
const CB = 6.75; // cards bottom
box(0.6, 2.18, 6.27, CB - 2.18, CARD, LINE, 0.16, "Participants card");
const kpis = [
  { x: 0.85, n: "460", nc: WHITE, l: [{ text: "Total registered\nparticipants", options: { color: MUTED } }] },
  { x: 2.98, n: "326", nc: WHITE, l: [{ text: "Qualified & rewarded\n", options: { color: MUTED } }, { text: "70.9% success rate", options: { color: GREEN, bold: true } }] },
  { x: 4.9, n: "134", nc: MUTED, l: [{ text: "Did not meet\nholding criteria", options: { color: MUTED } }] },
];
kpis.forEach((k, i) => {
  rect(k.x, 2.42, 0.42, 0.04, GREEN);
  T(k.n, { x: k.x, y: 2.52, w: 1.8, h: 0.58, fontSize: 34, bold: true, color: k.nc });
  T(k.l, { x: k.x, y: 3.12, w: 1.9, h: 0.5, fontSize: 11, lineSpacingMultiple: 1.05 });
  if (i > 0) s.addShape(pres.shapes.LINE, { x: k.x - 0.15, y: 2.42, w: 0, h: 1.2, line: { color: "3A453E", width: 0.75 } });
});

// bar: qualified 326 vs not qualified 134 (of 460)
const bx = 0.85, bw = 5.77, by = 4.05, bh = 0.16;
rect(bx, by, bw, bh, TRACK, true);
rect(bx, by, bw * 326 / 460, bh, GREEN, true);
[["Qualified (326)", GREEN, 0.85], ["Not qualified (134)", "6B746E", 2.45]].forEach(([t, c, x]) => {
  s.addShape(pres.shapes.OVAL, { x, y: 4.36, w: 0.1, h: 0.1, fill: { color: c }, line: { type: "none" } });
  T(t, { x: x + 0.16, y: 4.29, w: 1.6, h: 0.25, fontSize: 10, color: MUTED });
});

// average metric boxes
const avgs = [
  { x: 0.85, k: "AVG. REWARD", n: "≈ 12", u: "AUSD", thb: "≈ ฿406", d: " per qualified user" },
  { x: 3.82, k: "AVG. AUSD HELD", n: "≈ 101", u: "AUSD", thb: "≈ ฿3,408", d: " per qualified user" },
];
avgs.forEach((t) => {
  box(t.x, 4.95, 2.8, 1.55, INNER, INNER_LINE, 0.12, t.k + " box");
  T(t.k, { x: t.x + 0.25, y: 5.15, w: 2.4, h: 0.25, fontSize: 10, bold: true, color: GOLD, charSpacing: 1 });
  T([{ text: t.n + " ", options: { fontSize: 29, bold: true, color: WHITE } }, { text: t.u, options: { fontSize: 13, bold: true, color: GOLD } }],
    { x: t.x + 0.25, y: 5.42, w: 2.4, h: 0.55, valign: "bottom" });
  T([{ text: t.thb, options: { color: GREEN, bold: true } }, { text: t.d, options: { color: MUTED } }],
    { x: t.x + 0.25, y: 6.02, w: 2.5, h: 0.25, fontSize: 11 });
});

// ---- RIGHT CARD 1: AUSD holdings
box(7.13, 2.18, 5.6, 2.15, CARD, LINE, 0.16, "Holdings card");
s.addImage({ path: A("ausd-coin.png"), x: 7.42, y: 2.73, w: 1.0, h: 1.0, altText: "AUSD" });
T([{ text: "32,972 ", options: { fontSize: 36, bold: true, color: WHITE } }, { text: "AUSD", options: { fontSize: 17, bold: true, color: GOLD } }],
  { x: 8.7, y: 2.45, w: 3.9, h: 0.65, valign: "bottom" });
T("≈ ฿1,111,156 THB", { x: 8.7, y: 3.13, w: 3.9, h: 0.32, fontSize: 15, bold: true, color: GREEN });
T([{ text: "Held by ", options: { color: MUTED } }, { text: "326", options: { color: WHITE, bold: true } }, { text: " qualified holders", options: { color: MUTED } }],
  { x: 8.7, y: 3.5, w: 3.9, h: 0.3, fontSize: 12 });

// ---- RIGHT CARD 2: budget
box(7.13, 4.55, 5.6, CB - 4.55, CARD, LINE, 0.16, "Budget card");
T("Rewards distributed", { x: 7.38, y: 4.76, w: 3, h: 0.26, fontSize: 11, color: MUTED });
T([{ text: "3,925 ", options: { fontSize: 29, bold: true, color: WHITE } }, { text: "/ 5,450 AUSD", options: { fontSize: 13, bold: true, color: MUTED } }],
  { x: 7.38, y: 5.0, w: 4, h: 0.5, valign: "bottom" });
T([{ text: "≈ ฿132,273", options: { color: GREEN, bold: true } }, { text: " of ฿183,665 THB budget", options: { color: MUTED } }],
  { x: 7.38, y: 5.55, w: 4.5, h: 0.26, fontSize: 11 });
T("72%", { x: 11.28, y: 5.02, w: 1.2, h: 0.46, fontSize: 24, bold: true, color: GREEN, align: "right", valign: "bottom" });
rect(7.38, 5.95, 5.1, 0.16, TRACK, true);
rect(7.38, 5.95, 5.1 * 0.72, 0.16, GREEN, true);
T("Budget utilised", { x: 7.38, y: 6.2, w: 2.5, h: 0.26, fontSize: 10, color: MUTED });
T([{ text: "Remaining: ", options: { color: MUTED } }, { text: "1,525 AUSD", options: { color: GOLD, bold: true } }, { text: " (≈ ฿51,393)", options: { color: MUTED } }],
  { x: 9.18, y: 6.2, w: 3.3, h: 0.26, fontSize: 10, align: "right" });

// ---- Thank-you footer
T([{ text: "Thank you to ", options: { color: "5F6A63" } }, { text: "BBT Tech", options: { color: DIM, bold: true } },
   { text: ", ", options: { color: "5F6A63" } }, { text: "BBT Product", options: { color: DIM, bold: true } },
   { text: " and ", options: { color: "5F6A63" } }, { text: "Coin Partnership", options: { color: DIM, bold: true } },
   { text: " teams", options: { color: "5F6A63" } }],
  { x: 0.6, y: 7.0, w: 6, h: 0.28, fontSize: 8 });

s.addNotes("AUSD Dropzone results (reward drop 29 Oct 2026). 460 registered; 326 qualified (70.9%); 134 did not qualify. 30-day holdings: 32,972 AUSD (≈ 1,111,156 THB) across 326 holders. Budget: 3,925 of 5,450 AUSD distributed (72%; ≈ 132,273 of 183,665 THB), 1,525 AUSD (≈ 51,393 THB) remaining. THB at 1 AUSD ≈ 33.7 THB (USD/THB close, 29 Oct 2026).");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
