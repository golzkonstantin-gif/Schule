const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak, Header, Footer, TabStopType,
  PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", BOARD = "F3F6F1";
const C = { hv: "D9534F", mv: "2E9E6B", p2: "E8A33D", inf: "3F7CC4" };
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext, indent: o.indent });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h2 = (label, text, after = 80) => p([t(label + "  ", { font: "Cambria", size: 26, bold: true, color: C.hv }), t(text, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 200, after, keepNext: true });
const h3 = (text, col = NAVY) => p(t(text, { font: "Cambria", size: 23, bold: true, color: col }), { before: 120, after: 60, keepNext: true });
const gap = (n = 80) => p(t(""), { after: n });
const br = () => new Paragraph({ children: [new PageBreak()] });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };
const thin = allBorders(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || thin,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font, italics: o.it }), { after: 0, align: o.align }), { w, fill: o.fill });

// simple grid table: header + rows of strings/runs
function grid(widths, head, rows, o = {}) {
  return table(widths, [
    row(head.map((h, i) => hdr(h, widths[i], o.headFill ? o.headFill[i] : NAVY)), 420),
    ...rows.map((r) => row(r.map((v, i) => tc(v, widths[i], { font: o.fonts && o.fonts[i], align: o.align && o.align[i], bold: o.bold && o.bold[i], color: o.colors && o.colors[i] })), o.h || 460)),
  ]);
}

function box(children, fill = LIGHT2, border) {
  return table([W], [row([cell(children, { w: W, fill, borders: border || allBorders(none), m: 140, ml: 180 })])]);
}
// Merksatz box
function merk(nr, title, runs) {
  return box([
    p([t(`Merksatz ${nr} · `, { bold: true, color: C.hv, size: 20 }), t(title, { bold: true, color: NAVY, size: 22, font: "Cambria" })], { after: 60 }),
    ...runs.map((r) => p(r, { after: 40 })),
  ], LIGHT2);
}
// Tafel box
function tafel(title, children) {
  return table([W], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w: W, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)) })]),
    row([cell(children, { w: W, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 140, ml: 200 })]),
  ]);
}
const R = (s) => t(s, { bold: true, color: C.hv, font: "Cambria", size: 22 });
const Y = (s) => t(s, { bold: true, color: "C98A1E", font: "Cambria", size: 22 });
const B = (s) => t(s, { bold: true, color: C.inf, font: "Cambria", size: 22 });
const N = (s, o = {}) => t(s, Object.assign({ font: "Cambria", size: 22 }, o));
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));


// ================================================================
// Gemeinsame Bausteine für Tafelskripte
// ================================================================

const header = (label) => new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 7A · " + label, { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })],
  })],
});
const footer = (label) => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })],
});

// ---------- Was die Klasse gerade tut (nur zur Planung) ----------
const LABELS = {
  zuhoeren: "hört zu", gemeinsam: "macht mündlich mit", abschreiben: "schreibt ab", luecken: "arbeitet allein",
  tafel: "kommt an die Tafel", kontrolle: "kontrolliert", partner: "arbeitet zu zweit", fertig: "gibt ab",
};
function labelCell(m, w) {
  return cell([
    p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }),
    p(t(LABELS[m], { bold: true, color: NAVY, size: 20 }), { after: 0 }),
  ], { w, fill: LIGHT, borders: allBorders(solid("D5DDEE", 6)), m: 100, valign: VerticalAlign.TOP });
}

// Schritt: Modus + Inhalt
const BW = 1700, CW = W - BW;
function step(m, parts) {
  const children = [];
  if (parts.say) children.push(p([t("Ansage: ", { bold: true, size: 20, color: C.hv }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((d) => children.push(p(Array.isArray(d) ? d : S(d), { after: 50 })));
  if (parts.board) { children.push(...[].concat(parts.board)); children.push(p(t(""), { after: 20 })); }
  if (parts.sol) children.push(p([t("Lösung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  if (children.length && !parts.sol && !parts.board) children[children.length - 1] = children[children.length - 1];
  return table([BW, CW], [row([labelCell(m, BW), cell(children, { w: CW, borders: allBorders(solid("D5DDEE", 6)), m: 100, ml: 160, valign: VerticalAlign.TOP })])]);
}
const steps = (arr) => arr.flatMap((s) => [step(s[0], s[1]), p(t(""), { after: 40 })]);

// Tafel-Kasten (innerhalb eines Schritts)
const TW = CW - 340;
function board(title, children) {
  return table([TW], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 100, ml: 160 })]),
  ]);
}
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
// Merksatz steht auf der Folie (nicht an der Tafel)
function folie(title, children) {
  return table([TW], [
    row([cell(p(t("FOLIE · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: allBorders(solid(NAVY, 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: LIGHT, borders: allBorders(solid(NAVY, 12)), m: 100, ml: 160 })]),
  ]);
}
const merkBoard = (nr, title, runs) => folie(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  p(runs, { after: 0 }),
]);
// Übung an der Tafel – immer im gleichen Format
function auftrag(title, todo, lines, meta) {
  return board("rechts · " + title, [
    ...todo.map((x, i) => p([t(`${i + 1}. `, { bold: true, color: NAVY, size: 21 }), t(x, { bold: true, color: NAVY, size: 21 })], { after: 30 })),
    p(t(""), { after: 30 }),
    ...lines,
    p(t(""), { after: 30 }),
    p([t("Zeit: ", { bold: true, size: 19, color: "3D4A3F" }), t(meta, { size: 19, color: "3D4A3F" }), t("     Fertig? Stift hinlegen und noch einmal durchlesen.", { size: 19, italics: true, color: "3D4A3F" })], { after: 0 }),
  ]);
}
function phase(nr, title, min) {
  return p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
}


// ================================================================
// Druckvorlage „Welche Zeitform wann?“
// Seite 1: Schnipsel Zeitstrahl (2 pro A4) · Seite 2: Ü12 Fehlertext (2 pro A4)
// ab Seite 3: Satzstreifen und Fehlerstreifen für die Tafel (Großformat, 2 pro A4 quer)
// ================================================================
const { PageOrientation } = require("docx");
const PW = 11906 - 1200;
const IW = PW - 300; // Innenbreite im Schnipsel
const gray = "9AA3B5";
const sz = 22;
const s = (x, o = {}) => t(x, Object.assign({ size: sz }, o));
const sb = (x, o = {}) => s(x, Object.assign({ bold: true }, o));
const scissors = () => p(t("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", { color: gray, size: 16 }), { before: 80, after: 80, align: AlignmentType.CENTER });
const slip = (children) => table([PW], [row([cell(children, { w: PW, borders: allBorders(dashed), m: 140, ml: 150, valign: VerticalAlign.TOP })])]);
const slipHead = (nr, title) => p([t(nr ? nr + "  " : "", { font: "Cambria", bold: true, size: 26, color: C.hv }), t(title, { font: "Cambria", bold: true, size: 26, color: NAVY }), t("          Name: ______________________", { size: 18, color: MUTED })], { after: 80 });

// ---------- Schnipsel 1: Zeitstrahl ----------
const ZW = [2070, 2070, 2070, 2070, IW - 4 * 2070];
const zb = allBorders(solid("7A869E", 8));
const secCell = (x, w, span) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 22 }), { after: 0, align: AlignmentType.CENTER }), { w, fill: NAVY, span, borders: zb });
const emptyCell = (label, w) => cell(p(t(label, { size: 15, color: gray }), { after: 0 }), { w, borders: zb, valign: VerticalAlign.TOP, m: 50, ml: 80 });
function slipZeit() {
  const r = (label, h) => row(ZW.map((w) => emptyCell(label, w)), h, HeightRule.EXACT);
  return slip([
    slipHead("", "Die Zeitformen am Zeitstrahl"),
    p([s("Trage für jede Zeitform ihren Namen, den Bauplan und einen Beispielsatz von der Tafel ein. Markiere Hilfsverb "), sb("rot", { color: C.hv }), s(", Partizip II "), sb("gelb", { color: "C98A1E" }), s(", Infinitiv "), sb("blau", { color: C.inf }), s(" und unterstreiche ein finites Verb ohne Hilfsverb.")], { after: 100 }),
    table(ZW, [
      row([secCell("vorher", ZW[0]), secCell("Vergangenheit", ZW[1] + ZW[2], 2), secCell("jetzt", ZW[3]), secCell("Zukunft  ▶", ZW[4])], 420),
      r("Zeitform", 820),
      r("Bauplan", 1050),
      r("Beispielsatz", 3500),
    ]),
  ]);
}

// ---------- Schnipsel 2: Ü12 Fehlertext ----------
const BER = [
  "Am Samstag sind unsere Judoka nach Leipzig gefahren.",
  "Sie hatten sich wochenlang auf das Turnier vorbereitet.",
  "Im Halbfinale hat eine Judoka mit einem Haltegriff gewonnen.",
  "Unsere Gewichtheber hatten schon um sechs Uhr losgefahren.",
  "Im Reißen schaffte einer von ihnen einen neuen Vereinsrekord.",
  "Die Handballer werften am Sonntag 28 Tore.",
  "Vorher hatten sie den Siebenmeter lange geübt.",
];
function slipFehler() {
  const step = (i, x) => p([sb(`${i}. `, { color: NAVY, size: 20 }), s(x, { size: 20 })], { after: 20 });
  return slip([
    slipHead("Ü12", "Der Bericht für die Vereinszeitung"),
    step(1, "Lies den Bericht. Ein Teamkollege hat ihn so geschrieben, wie er spricht."),
    step(2, "Unterstreiche die vier Verbformen, die falsch sind."),
    step(3, "Schreibe die richtige Form darüber."),
    step(4, "Begründe einen Fehler im Heft: „… ist falsch, weil …“"),
    p(t("Starkes Wochenende für unseren Verein", { font: "Cambria", bold: true, size: 24, color: NAVY }), { before: 80, after: 0 }),
    ...BER.map((x, i) => p([t(`${i + 1}  `, { size: 18, bold: true, color: MUTED }), t(x, { font: "Cambria", size: 24 })], { before: 420, after: 0 })),
    p([sb("Zusatz: ", { size: 19, color: C.hv }), s("Was hattest du vor deinem Wettkampf gemacht? Schreibe einen Satz im Plusquamperfekt ins Heft.", { size: 19 }), s("   Fertig? Stift hinlegen.", { size: 19, italics: true, color: MUTED })], { before: 200, after: 0 }),
  ]);
}

// ---------- Streifen für die Tafel ----------
const LWS = 16838 - 1200;
const STRIPS = [
  ["VEREINSZEITUNG", "Die Judoka hatten wochenlang für das Turnier trainiert."],
  ["SPRACHNACHRICHT", "Ich habe im Finale Bronze geholt!"],
  ["VEREINSZEITUNG", "Im Stoßen hob unsere Gewichtheberin 75 Kilo."],
  ["POST", "Beim nächsten Turnier werde ich den Haltegriff besser verteidigen."],
  ["VEREINSZEITUNG", "Die Handballer waren schon am Freitag angereist."],
  ["SPRACHNACHRICHT", "Wir sind erst um zehn Uhr nach Hause gekommen."],
  ["POST", "Heute analysieren wir mit dem Trainer das Video."],
  ["VEREINSZEITUNG", "Unsere Handballer kämpften bis zur letzten Sekunde."],
  ["BERICHT · SATZ 1", "Am Samstag sind unsere Judoka nach Leipzig gefahren."],
  ["BERICHT · SATZ 3", "Im Halbfinale hat eine Judoka mit einem Haltegriff gewonnen."],
  ["BERICHT · SATZ 4", "Unsere Gewichtheber hatten schon um sechs Uhr losgefahren."],
  ["BERICHT · SATZ 6", "Die Handballer werften am Sonntag 28 Tore."],
];
const strip = ([lab, txt]) => table([LWS], [row([cell([
  p([t(lab, { size: 40, color: MUTED, bold: true })], { after: 120 }),
  p([t(txt, { font: "Cambria", size: 96, bold: true, color: "111111" })], { after: 0, align: AlignmentType.CENTER }),
], { w: LWS, borders: allBorders(dashed), m: 200, ml: 260, valign: VerticalAlign.CENTER })], 5050, HeightRule.EXACT)]);
const strips = [];
STRIPS.forEach((x, i) => {
  strips.push(strip(x));
  if (i % 2 === 0) strips.push(p(t(""), { after: 0 }));
  else if (i < STRIPS.length - 1) strips.push(new Paragraph({ children: [new PageBreak()] }));
});

const doc = [
  slipZeit(), scissors(), slipZeit(),
  new Paragraph({ children: [new PageBreak()] }),
  slipFehler(), scissors(), slipFehler(),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 400, left: 600, right: 600 } } };
const landProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 600, right: 600 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: doc }, { properties: landProps, children: strips }] })).then((buf) => fs.writeFileSync("Druckvorlage_Zeitformen_Zeitstrahl.docx", buf));
