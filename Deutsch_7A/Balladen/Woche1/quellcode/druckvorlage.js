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
// Druckvorlage Balladen Woche 1: Notizraster (Mi, 2 pro A4), Ü1 (Mi, 1 pro A4), Ü2 (Fr, 2 pro A4)
// ================================================================
const PW = 11906 - 1200;
const IW = PW - 300;
const gray = "9AA3B5";
const s = (x, o = {}) => t(x, Object.assign({ size: 20 }, o));
const sb = (x, o = {}) => s(x, Object.assign({ bold: true }, o));
const scissors = () => p(t("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", { color: gray, size: 16 }), { before: 40, after: 40, align: AlignmentType.CENTER });
const slip = (children) => table([PW], [row([cell(children, { w: PW, borders: allBorders(dashed), m: 110, ml: 150, valign: VerticalAlign.TOP })])]);
const head = (nr, title) => p([t(nr ? nr + "  " : "", { font: "Cambria", bold: true, size: 26, color: C.hv }), t(title, { font: "Cambria", bold: true, size: 26, color: NAVY }), t("          Name: ____________________", { size: 17, color: MUTED })], { after: 70 });
const bd = allBorders(solid("7A869E", 8));
const field = (label, w, h, hint) => cell([p(t(label, { bold: true, size: 19, color: NAVY }), { after: 0 }), ...(hint ? [p(t(hint, { size: 15, italics: true, color: gray }), { after: 0 })] : [])], { w, borders: bd, valign: VerticalAlign.TOP, m: 60, ml: 90 });
const lineRows = (n, h = 520) => table([IW], Array.from({ length: n }, () => row([cell(p(t(""), { after: 0 }), { w: IW, borders: { top: none, left: none, right: none, bottom: solid("9AA3B5", 6) } })], h, HeightRule.EXACT)));

// ---------- Notizraster ----------
function slipHoeren() {
  const H = IW / 2;
  return slip([
    head("", "John Maynard hören"),
    p([s("1. Hören: Notiere Stichpunkte. 2. Hören: Ergänze und beantworte die letzte Frage.")], { after: 80 }),
    table([H, H], [
      row([field("Wer?", H, 0, "Figuren"), field("Wo?", H, 0, "Ort")], 1300, HeightRule.EXACT),
      row([cell([p(t("Was passiert?", { bold: true, size: 19, color: NAVY }), { after: 0 }), p(t("3–4 Stichpunkte", { size: 15, italics: true, color: gray }), { after: 0 })], { w: IW, borders: bd, span: 2, valign: VerticalAlign.TOP, m: 60, ml: 90 })], 2000, HeightRule.EXACT),
      row([cell([p(t("Wie endet es?", { bold: true, size: 19, color: NAVY }), { after: 0 })], { w: IW, borders: bd, span: 2, valign: VerticalAlign.TOP, m: 60, ml: 90 })], 1100, HeightRule.EXACT),
      row([cell([p(t("Wer ist John Maynard? (nach dem 2. Hören, ein ganzer Satz)", { bold: true, size: 19, color: C.hv }), { after: 0 })], { w: IW, borders: bd, span: 2, valign: VerticalAlign.TOP, m: 60, ml: 90 })], 1100, HeightRule.EXACT),
    ]),
  ]);
}

// ---------- Ü1 ----------
function slipLesen() {
  const QW = [IW - 1500, 1500];
  const q = (nr, x) => row([cell([p([sb(nr + "  ", { color: C.hv, size: 19 }), s(x, { size: 19 })], { after: 0 })], { w: QW[0], borders: bd, m: 50, ml: 90, valign: VerticalAlign.TOP }), cell(p(t("Str. ___", { size: 18, color: gray }), { after: 0, align: AlignmentType.CENTER }), { w: QW[1], borders: bd, m: 50 })], 700, HeightRule.EXACT);
  const WW = [1900, (IW - 1900) / 2, (IW - 1900) / 2];
  const wh = (x, w) => cell(p(t(x, { bold: true, size: 17, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w, fill: NAVY, borders: bd, m: 30 });
  const wr = (x) => row([cell(p(t(x, { font: "Cambria", bold: true, size: 20 }), { after: 0 }), { w: WW[0], borders: bd, m: 40, ml: 90 }), cell(p(t(""), { after: 0 }), { w: WW[1], borders: bd }), cell(p(t(""), { after: 0 }), { w: WW[2], borders: bd })], 700, HeightRule.EXACT);
  return slip([
    head("Ü1", "Drei Arten zu lesen"),
    p([sb("a) überfliegend  ", { color: NAVY }), s("(1 Minute, dann Buch zu)  In der Ballade geht es um …", { size: 19 })], { after: 0 }),
    lineRows(3, 600),
    p([sb("b) orientierend  ", { color: NAVY }), s("Suche nur die Antworten und notiere die Strophe.", { size: 19 })], { before: 80, after: 30 }),
    table(QW, [
      q("1", "Wie heißt das Schiff?"),
      q("2", "Von wo nach wo fährt es?"),
      q("3", "Wie viele Minuten sind es noch bis Buffalo, als das Feuer ausbricht?"),
      q("4", "Was antwortet John Maynard dem Kapitän zuletzt?"),
      q("5*", "Wie viele Menschen folgen dem Sarg?"),
    ]),
    p([sb("c) genau  ", { color: NAVY }), s("Lies Strophe 5 und 6 Satz für Satz. Notiere Stichpunkte, was passiert:", { size: 19 })], { before: 80, after: 0 }),
    lineRows(3, 600),
    p([sb("* Wörter aus dem Kontext  ", { color: NAVY, size: 19 }), s("Was bedeutet das Wort? Welcher Hinweis im Text hilft?", { size: 19 })], { before: 70, after: 30 }),
    table(WW, [row([wh("Wort", WW[0]), wh("Meine Vermutung", WW[1]), wh("Hinweis im Text", WW[2])]), wr("Sprachrohr"), wr("Brandung"), wr("Bugspriet")]),
  ]);
}

// ---------- Ü2 ----------
function slipGliedern() {
  const GW2 = [1700, IW - 1700];
  const gh = (x, w) => cell(p(t(x, { bold: true, size: 18, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w, fill: NAVY, borders: bd, m: 30 });
  const gr = (a) => row([cell(p(t(a, { size: 18, color: gray }), { after: 0 }), { w: GW2[0], borders: bd, m: 40, ml: 90 }), cell(p(t(""), { after: 0 }), { w: GW2[1], borders: bd })], 820, HeightRule.EXACT);
  return slip([
    head("Ü2", "John Maynard gliedern"),
    p([s("Trage ein, welche Strophen zu einem Abschnitt gehören, und schreibe eine "), sb("Zwischenüberschrift"), s(" (kein ganzer Satz). Abschnitt 2 ist ein Beispiel.")], { after: 80 }),
    table(GW2, [
      row([gh("Strophe(n)", GW2[0]), gh("Zwischenüberschrift", GW2[1])]),
      gr("Str. ____"),
      row([cell(p(t("Str. 2", { size: 18, bold: true }), { after: 0 }), { w: GW2[0], borders: bd, m: 40, ml: 90, fill: LIGHT }), cell(p(t("Fröhliche Fahrt nach Buffalo", { font: "Cambria", size: 20, italics: true }), { after: 0 }), { w: GW2[1], borders: bd, m: 40, ml: 90, fill: LIGHT })], 620, HeightRule.EXACT),
      gr("Str. ____"), gr("Str. ____"), gr("Str. ____"), gr("Str. ____"), gr("Str. ____"),
    ]),
    p([sb("Test: ", { size: 19 }), s("Kann ich mit meinen Überschriften die Ballade nacherzählen?", { size: 19 })], { before: 80, after: 0 }),
  ]);
}

const doc = [
  slipHoeren(), scissors(), slipHoeren(),
  new Paragraph({ children: [new PageBreak()] }),
  slipLesen(),
  new Paragraph({ children: [new PageBreak()] }),
  slipGliedern(), scissors(), slipGliedern(),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 400, left: 600, right: 600 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: doc }] })).then((buf) => fs.writeFileSync("Druckvorlage_Balladen_Woche1.docx", buf));
