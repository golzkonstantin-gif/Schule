// Arbeitsblatt „Vor dem Gesetz“ nach dem bekannten Muster (Kafka-Appell, „Gibs auf!“): Titel, Thema, Ziel, Info-Box,
// 1. Textgrundlage mit Zeilennummern, 2. Information „kafkaesk“, 3. Arbeitsaufträge mit Platz zum Schreiben (2 Seiten, doppelseitig)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Header, Footer, TabStopType,
} = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", BLUE = "3F7CC4";
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 80, line: o.line } });
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
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none), m: 90, ml: 160 })])]);
const h2 = (text, o = {}) => p(t(text, { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: o.before ?? 160, after: 70, pb: o.pb, keepNext: true });
const aufgabe = (a) => p([t(a.titel, { bold: true, color: NAVY, size: 22 }), t("   (" + a.form + ", " + a.zeit + ")", { italics: true, color: MUTED, size: 18 })], { before: 140, after: 40, keepNext: true });
const schritte = (arr) => arr.map((x, i) => p([t((i + 1) + ".  ", { bold: true, color: NAVY, size: 20 }), t(x, { size: 20 })], { after: 30 }));
const kopf = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w, fill: NAVY });

// Text mit Zeilennummern (jede 5. Zeile nummeriert)
function textBlock(lines) {
  const nW = 480, tW = W - nW;
  return table([nW, tW], lines.map((l, i) => row([
    cell(p(t((i + 1) % 5 === 0 || i === 0 ? String(i + 1) : "", { size: 16, color: MUTED }), { align: AlignmentType.RIGHT, after: 0 }), { w: nW, borders: { top: none, bottom: none, left: none, right: solid(BLUE, 10) }, m: 4, ml: 80, valign: VerticalAlign.CENTER }),
    cell(p(t(l, { font: "Cambria", size: 21 }), { after: 0, line: 290 }), { w: tW, borders: all(none), m: 4, ml: 140 }),
  ])));
}

// Aufgabe 1: gemischte Handlungsschritte mit Kästchen für die Nummer und Spalte für die Zeile
const a1W = [700, 800, W - 2900, 1400];
const a1 = table(a1W, [
  row([kopf("", a1W[0]), kopf("Nr.", a1W[1]), kopf("Handlungsschritt", a1W[2]), kopf("Zeile", a1W[3])]),
  ...D.REIHE_GEMISCHT.map((k, i) => row([
    cell(p(t(String.fromCharCode(65 + i), { bold: true, color: NAVY, size: 19 }), { after: 0, align: AlignmentType.CENTER }), { w: a1W[0], fill: LIGHT, valign: VerticalAlign.CENTER }),
    cell(p(t(""), { after: 0 }), { w: a1W[1] }),
    cell(p(t(D.SCHRITTE[k][0], { size: 19 }), { after: 0 }), { w: a1W[2], valign: VerticalAlign.CENTER }),
    cell(p(t("Z. ", { size: 19, color: MUTED }), { after: 0 }), { w: a1W[3], valign: VerticalAlign.CENTER }),
  ], 420)),
]);

// Aufgabe 2: Bildebene → Sachebene
const a2W = [2900, 1000, W - 3900];
const a2 = table(a2W, [
  row([kopf("Bildebene (achte auf …)", a2W[0]), kopf("Zeile", a2W[1]), kopf("Sachebene: Wofür steht es?", a2W[2])]),
  ...D.SYMBOLE.map(([b, achte, , , stern]) => row([
    cell([p(t(b + (stern ? " *" : ""), { bold: true, color: stern ? MUTED : RED, size: 19 }), { after: 10 }), p(t(achte, { size: 17, italics: true, color: BLUE }), { after: 0 })], { w: a2W[0], fill: LIGHT }),
    cell(p(t("Z. ", { size: 19, color: MUTED }), { after: 0 }), { w: a2W[1] }),
    cell(p(t(""), { after: 0 }), { w: a2W[2] }),
  ], 740)),
]);

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Vor dem Gesetz“", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t("Franz Kafka: „Vor dem Gesetz“ (1915), in heutiger Rechtschreibung", { size: 16, color: MUTED })] })] });

const seite1 = [
  p(t("Arbeitsblatt: Inhalt und Symbole in Kafkas „Vor dem Gesetz“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 50 }),
  p([t("Thema: ", { bold: true, size: 20 }), t("Ohnmacht vor der Macht – eine kafkaeske Parabel", { size: 20 })], { after: 20 }),
  p([t("Ziel: ", { bold: true, size: 20 }), t("Ich kann die Symbole der Parabel deuten und erklären, warum sie kafkaesk ist.", { size: 20 })], { after: 100 }),
  box([
    p(t("Info-Box: Kontext", { bold: true, color: RED, size: 21 }), { after: 30 }),
    p(t(D.KONTEXT.join(" "), { size: 19 }), { after: 0, line: 270 }),
  ]),
  h2("1. Textgrundlage: Franz Kafka, „Vor dem Gesetz“", { before: 130 }),
  textBlock(D.ZEILEN),
];

const seite2 = [
  h2("2. Information: „kafkaesk“", { pb: true, before: 0 }),
  box([
    p(t(D.KAFKAESK.duden, { bold: true, size: 20 }), { after: 40 }),
    ...D.KAFKAESK.merkmale.map((m) => p([t("▸ ", { color: RED, size: 19 }), t(m, { size: 19 })], { after: 15 })),
  ]),
  h2("3. Arbeitsaufträge"),
  p(t("Aufgaben 1 und 2 löst du direkt auf diesem Blatt. Den Zielcheck schreibst du ins Heft.", { size: 18, italics: true, color: MUTED }), { after: 20 }),
  aufgabe(D.AUFGABE1),
  ...schritte(D.AUFGABE1.schritte),
  a1,
  aufgabe(D.AUFGABE2),
  ...schritte(D.AUFGABE2.schritte),
  a2,
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 760, bottom: 620, left: 850, right: 850, header: 340, footer: 340 } } }, headers: { default: header }, footers: { default: footer }, children: [...seite1, ...seite2] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Arbeitsblatt_Vor_dem_Gesetz.docx", b));
