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
   { text: "29 October 2026", options: { color: GOLD, bold: true } }],
  { x: 0.6, y: 1.12, w: 10, h: 0.36, fontSize: 14 });

// ---- Section titles
T("Participants & Rewards", { x: 0.6, y: 1.7, w: 6, h: 0.4, fontSize: 17, bold: true, color: GREEN });
T("AUSD Held · 30-Day Period", { x: 7.13, y: 1.7, w: 5.6, h: 0.4, fontSize: 17, bold: true, color: GREEN });

// ---- LEFT CARD: participants
box(0.6, 2.18, 6.27, 4.02, CARD, LINE, 0.16, "Participants card");
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

// stacked bar: Tier2 133, Tier1 193, Not qualified 134 (of 460)
const bx = 0.85, bw = 5.77, by = 3.95, bh = 0.16;
const g = bw * 133 / 460, o = bw * 193 / 460;
rect(bx, by, bw, bh, TRACK, true);
rect(bx + g, by, o, bh, GOLD);
rect(bx, by, 0.3, bh, GREEN, true);
rect(bx + 0.15, by, g - 0.15, bh, GREEN);
[["Tier 2", GREEN, 0.85], ["Tier 1", GOLD, 1.72], ["Not qualified", "6B746E", 2.6]].forEach(([t, c, x]) => {
  s.addShape(pres.shapes.OVAL, { x, y: 4.26, w: 0.1, h: 0.1, fill: { color: c }, line: { type: "none" } });
  T(t, { x: x + 0.16, y: 4.19, w: 1.4, h: 0.25, fontSize: 10, color: MUTED });
});

// tier boxes
const tiers = [
  { x: 0.85, pill: "TIER 2", filled: true, n: "133", p: "40.8% of rewarded users", pc: GREEN },
  { x: 3.82, pill: "TIER 1", filled: false, n: "193", p: "59.2% of rewarded users", pc: GOLD },
];
tiers.forEach((t) => {
  box(t.x, 4.78, 2.8, 1.18, INNER, INNER_LINE, 0.12, t.pill + " box");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: t.x + 0.18, y: 5.24, w: 0.62, h: 0.27, rectRadius: 0.135,
    fill: t.filled ? { color: GREEN } : { color: INNER }, line: t.filled ? { type: "none" } : { color: GOLD, width: 1 },
  });
  T(t.pill, { x: t.x + 0.18, y: 5.24, w: 0.62, h: 0.27, fontSize: 8, bold: true, align: "center", valign: "middle", color: t.filled ? "08200F" : GOLD });
  T(t.n, { x: t.x + 0.95, y: 4.98, w: 1.7, h: 0.5, fontSize: 29, bold: true, color: WHITE });
  T(t.p, { x: t.x + 0.95, y: 5.5, w: 1.8, h: 0.25, fontSize: 10, color: t.pc });
});

// ---- RIGHT CARD 1: AUSD holdings
box(7.13, 2.18, 5.6, 1.95, CARD, LINE, 0.16, "Holdings card");
s.addImage({ path: A("ausd-coin.png"), x: 7.42, y: 2.6, w: 0.95, h: 0.95, altText: "AUSD" });
T([{ text: "32,972 ", options: { fontSize: 36, bold: true, color: WHITE } }, { text: "AUSD", options: { fontSize: 17, bold: true, color: GOLD } }],
  { x: 8.65, y: 2.5, w: 3.9, h: 0.65, valign: "bottom" });
T([{ text: "Held by ", options: { color: MUTED } }, { text: "326", options: { color: WHITE, bold: true } }, { text: " qualified holders", options: { color: MUTED } }],
  { x: 8.65, y: 3.2, w: 3.9, h: 0.3, fontSize: 12 });

// ---- RIGHT CARD 2: budget
box(7.13, 4.35, 5.6, 1.85, CARD, LINE, 0.16, "Budget card");
T("Rewards distributed", { x: 7.38, y: 4.57, w: 3, h: 0.26, fontSize: 11, color: MUTED });
T([{ text: "3,925 ", options: { fontSize: 29, bold: true, color: WHITE } }, { text: "/ 5,450 AUSD", options: { fontSize: 13, bold: true, color: MUTED } }],
  { x: 7.38, y: 4.84, w: 4, h: 0.5, valign: "bottom" });
T("72%", { x: 11.28, y: 4.86, w: 1.2, h: 0.46, fontSize: 24, bold: true, color: GREEN, align: "right", valign: "bottom" });
rect(7.38, 5.5, 5.1, 0.16, TRACK, true);
rect(7.38, 5.5, 5.1 * 0.72, 0.16, GREEN, true);
T("Budget utilised", { x: 7.38, y: 5.76, w: 2.5, h: 0.26, fontSize: 10, color: MUTED });
T([{ text: "Remaining: ", options: { color: MUTED } }, { text: "1,525 AUSD", options: { color: GOLD, bold: true } }],
  { x: 9.98, y: 5.76, w: 2.5, h: 0.26, fontSize: 10, align: "right" });

// ---- Headline strip
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.4, w: 12.13, h: 0.46, rectRadius: 0.1,
  fill: { color: GOLD, transparency: 88 }, line: { color: GOLD, width: 1.25 }, objectName: "Headline metrics" });
T([{ text: "Avg. reward  ", options: { color: GOLD, bold: true } }, { text: "≈ 12 AUSD per qualified user", options: { color: WHITE, bold: true } },
   { text: "          Avg. AUSD held  ", options: { color: GOLD, bold: true } }, { text: "≈ 101 AUSD per qualified user", options: { color: WHITE, bold: true } }],
  { x: 0.8, y: 6.4, w: 11.7, h: 0.46, fontSize: 11, valign: "middle" });

// ---- Thank-you footer
T([{ text: "Thank you to ", options: { color: "5F6A63" } }, { text: "BBT Tech", options: { color: DIM, bold: true } },
   { text: ", ", options: { color: "5F6A63" } }, { text: "BBT Product", options: { color: DIM, bold: true } },
   { text: " and ", options: { color: "5F6A63" } }, { text: "Coin Partnership", options: { color: DIM, bold: true } },
   { text: " teams", options: { color: "5F6A63" } }],
  { x: 0.6, y: 7.0, w: 6, h: 0.28, fontSize: 8 });

s.addNotes("AUSD Dropzone results (reward drop 29 Oct 2026). 460 registered; 326 qualified (70.9%): Tier 2 = 133, Tier 1 = 193; 134 did not qualify. 30-day holdings: 32,972 AUSD across 326 holders; Budget: 3,925 of 5,450 AUSD distributed (72%), 1,525 AUSD remaining.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
