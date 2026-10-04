// Arbeitsblatt: Abschnitte A (Unterrichtsgespräch) und C (Einzelarbeit), Matrix zu C, Fließtext
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Header, Footer, TabStopType, PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const COL_A = "3F7CC4", COL_C = "D9534F";
const W = 10206;
const AUSZ = JSON.parse(fs.readFileSync(path.join(__dirname, "auszuege.json"), "utf8"));

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: o.ml ?? 100, right: o.ml ?? 100 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none), m: 100, ml: 140 })])]);
const kicker = (text, pb) => p(t(text.toUpperCase(), { size: 17, bold: true, color: MUTED }), { after: 20, pb });
const h1 = (text) => p(t(text, { font: "Cambria", size: 34, bold: true, color: NAVY }), { after: 60 });

// Text mit Zeilennummern (jede Zeile nummeriert, feste Umbrüche) + Worterklärungen rechts
function textBlock(lines, glossar, col) {
  const nW = 520, tW = W - nW;
  const rows = lines.map((l, i) => row([
    cell(p(t(String(i + 1), { size: 16, color: MUTED }), { align: AlignmentType.RIGHT, after: 0 }), { w: nW, borders: { top: none, bottom: none, left: none, right: solid(col, 10) }, m: 6, ml: 80 }),
    cell(p(t(l, { font: "Cambria", size: 21 }), { after: 0 }), { w: tW, borders: all(none), m: 6, ml: 140 }),
  ]));
  rows.push(row([cell(p(glossar.flatMap(([w, e], i) => [t((i ? "   ·   " : "") + w + " ", { size: 17, bold: true, color: col }), t("= " + e, { size: 17, color: MUTED })]), { after: 0 }), { w: W, span: 2, borders: { top: solid("D5DCEC", 6), bottom: none, left: none, right: none }, m: 60, ml: 140 })]));
  return table([nW, tW], rows);
}
const lines = (n, h = 470) => table([W], Array.from({ length: n }, () => row([cell(p(t("")), { w: W, borders: { top: none, left: none, right: none, bottom: solid("9AA6C4") } })], h)));

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Brief an den Vater“", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Kafka, Brief an den Vater · Seite ", PageNumber.CURRENT] })] })] });

// ================= Seite 1: Texte =================
const page1 = [
  kicker("Franz Kafka: „Brief an den Vater“ (1919) · Auszüge von S. 5–6"),
  h1("Wie der Vater spricht"),
  box([
    p([t("These: ", { bold: true, color: RED }), t("Kafkas „Furcht“ gründet auf Kommunikationsproblemen.", { bold: true, color: NAVY })], { after: 40 }),
    p([t("So analysiert man: ", { bold: true, color: NAVY, size: 20 }), t("1 Verstehen (Wer? An wen? Worum?)  →  2 Untersuchen (Handlung · Modell · Wirkung)  →  3 Schreiben (Behauptung → Beleg → Erklärung)", { size: 20 })], { after: 0 }),
  ]),
  p(t("Der Text steht in der Rechtschreibung vor 1996: daß = dass, mußte = musste, wußte = wusste.", { size: 17, italics: true, color: MUTED }), { before: 60, after: 100 }),
  p([t("A  ", { font: "Cambria", size: 26, bold: true, color: COL_A }), t("Redeverbot", { font: "Cambria", size: 26, bold: true, color: NAVY }), t("   gemeinsam im Unterrichtsgespräch – die Matrix übernimmst du in deinen Hefter", { size: 19, italics: true, color: MUTED })], { after: 60, keepNext: true }),
  textBlock(AUSZ.A, [["Verkehr", "Umgang miteinander"], ["Widerrede", "Widerspruch"], ["Gegenkräfte", "eigener Wille, Widerstand"], ["contra", "dagegen"], ["[…]", "Auslassung"]], COL_A),
  p(t(""), { after: 120 }),
  p([t("C  ", { font: "Cambria", size: 26, bold: true, color: COL_C }), t("Ironie", { font: "Cambria", size: 26, bold: true, color: NAVY }), t("   allein – Aufgabe 1 auf der Rückseite", { size: 19, italics: true, color: MUTED })], { after: 60, keepNext: true }),
  textBlock(AUSZ.C, [["Ermahnung", "Zurechtweisung"], ["gewissermaßen", "sozusagen"], ["formell", "der Form nach"], ["gewürdigt", "für wert befunden"], ["Gegenspiel", "Gegenstück, Folge"]], COL_C),
];

// ================= Seite 2: Aufgabe 1 – Matrix zu C =================
const mW = [900, 3000, 3100, 3206];
const mHead = row([
  cell(p(t("Z.", { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), { w: mW[0], fill: NAVY }),
  cell([p(t("Handlung des Vaters", { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), p(t("Was tut er? (Zitat)", { color: "FFFFFF", size: 15 }), { after: 0 })], { w: mW[1], fill: NAVY }),
  cell([p(t("Modell", { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), p(t("Seite der Nachricht / Axiom", { color: "FFFFFF", size: 15 }), { after: 0 })], { w: mW[2], fill: NAVY }),
  cell([p(t("Wirkung auf den Sohn", { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), p(t("Zitat + Zeile", { color: "FFFFFF", size: 15 }), { after: 0 })], { w: mW[3], fill: NAVY }),
], 520);
const mRows = [mHead, ...Array.from({ length: 5 }, () => row(mW.map((w) => cell(p(t("")), { w })), 1250))];

const page2 = [
  kicker("Aufgabe 1 · allein · ca. 13 Min.", true),
  h1("Abschnitt C untersuchen"),
  p([t("Untersuche Abschnitt C so, wie wir Abschnitt A gemeinsam untersucht haben. Fülle die Matrix aus: ", { size: 21 }), t("mindestens vier Zeilen", { bold: true, size: 21 }), t(", jede mit Zitat und Zeilenangabe.", { size: 21 })], { after: 80 }),
  box([
    p([t("Hilfe, wenn du nicht weiterkommst:  ", { bold: true, color: NAVY, size: 19 }), t("Wie klingen die Fragen des Vaters – und wie sind sie gemeint? (Z. 3–5)  ·  Was sagen Lachen und Gesicht? (Z. 5)  ·  Mit wem spricht der Vater eigentlich? (Z. 9–10)  ·  Was macht der Sohn daraus? (Z. 11–15)", { size: 19 })], { after: 0 }),
  ], "FFF8EC"),
  p(t(""), { after: 80 }),
  table(mW, mRows),
  p([t("Zitieren: ", { bold: true, size: 18, color: NAVY }), t("„Das ist Dir wohl schon zu viel?“ (Z. 3 f.)  ·  ein Wort angepasst: die „erhobene[n] Hand“  ·  ausgelassen: „schon bestraft, ehe man noch wußte […]“", { size: 18, italics: true, color: MUTED })], { before: 80, after: 0 }),
];

// ================= Seite 3: Aufgabe 2 – Fließtext =================
const page3 = [
  kicker("Aufgabe 2 · allein · ca. 15 Min.", true),
  h1("Vom Tafelbild zum Fließtext"),
  p([t("Überführe die Matrix von der Tafel (A und C) in einen ", { size: 21 }), t("zusammenhängenden Text", { bold: true, size: 21 }), t(". Schreibe so, wie du es in der Klausur tun würdest.", { size: 21 })], { after: 80 }),
  table([3100, W - 3100], [
    row([cell(p(t("Einleitung", { bold: true, color: NAVY, size: 20 }), { after: 0 }), { w: 3100, fill: LIGHT }), cell(p(t("Wer schreibt an wen, wann, worüber? Worum geht es in den beiden Abschnitten?", { size: 19 }), { after: 0 }), { w: W - 3100 })], 380),
    row([cell(p(t("Hauptteil", { bold: true, color: NAVY, size: 20 }), { after: 0 }), { w: 3100, fill: LIGHT }), cell(p(t("Je Zeile der Matrix: Behauptung → Beleg (Zitat + Z.) → Erklärung mit Fachbegriff. Erst A, dann C.", { size: 19 }), { after: 0 }), { w: W - 3100 })], 380),
    row([cell(p(t("Schluss", { bold: true, color: NAVY, size: 20 }), { after: 0 }), { w: 3100, fill: LIGHT }), cell(p(t("Stützen die Stellen die These? Was muss man ergänzen?", { size: 19 }), { after: 0 }), { w: W - 3100 })], 380),
  ]),
  p(t(""), { after: 60 }),
  box([
    p([t("Satzbausteine:  ", { bold: true, color: NAVY, size: 19 }), t("In seinem „Brief an den Vater“ (1919) …  ·  Indem der Vater …, sendet er …  ·  Dies zeigt sich in … (Z. …)  ·  Hinzu kommt, dass …  ·  Die Folge ist …  ·  Besonders deutlich wird dies, wenn …  ·  Insgesamt stützt die Stelle die These, weil …  ·  Allerdings …", { size: 19, italics: true })], { after: 0 }),
  ], "FFF8EC"),
  p(t(""), { after: 80 }),
  lines(20),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 750, left: 850, right: 850, header: 380, footer: 380 } } }, headers: { default: header }, footers: { default: footer }, children: [...page1, ...page2, ...page3] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Arbeitsblatt_Kafka_Brief_an_den_Vater.docx", b));
