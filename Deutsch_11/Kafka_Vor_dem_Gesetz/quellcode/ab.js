// Arbeitsblatt „Vor dem Gesetz“ nach dem bekannten Muster (Kafka-Appell, „Gibs auf!“): Titel, Thema, Ziel, Info-Box,
// 1. Textgrundlage mit Zeilennummern, 2. Arbeitsaufträge 2–4 mit Platz zum Schreiben (Handlung über das Gespräch erschließen, Symbole) (2 Seiten, doppelseitig)
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
const aufgabe = (a) => p([t(a.titel, { bold: true, color: NAVY, size: 22 })], { before: 140, after: 40, keepNext: true });
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

const schreibLinien = (n, w, h = 420) => table([w], Array.from({ length: n }, () => row([cell(p(t(""), { after: 0 }), { w, borders: { top: none, left: none, right: none, bottom: solid("9AA6C4", 4) }, m: 0, ml: 0 })], h)));
const aufgabeKopf = (a) => p([t(a.titel, { bold: true, color: NAVY, size: 22 })], { before: 140, after: 40, keepNext: true });

// Aufgabe 3: das Gespräch Schritt für Schritt (Spalte „Der Mann“ vorgegeben)
const gW = [450, 3500, W - 3950];
const gespraech = table(gW, [
  row([kopf("", gW[0]), kopf("Der Mann vom Lande …", gW[1]), kopf("Der Türhüter … (Zitat, vgl. Z. …)", gW[2])]),
  ...D.GESPRAECH.map(([mann], i) => row([
    cell(p(t(String(i + 1), { bold: true, color: NAVY, size: 19 }), { after: 0, align: AlignmentType.CENTER }), { w: gW[0], fill: LIGHT }),
    cell([p(t(mann, { size: 19 }), { after: 10 }), p(t("vgl. " + D.MANN_Z[i], { size: 17, color: MUTED }), { after: 0 })], { w: gW[1], fill: LIGHT }),
    cell(p(t(""), { after: 0 }), { w: gW[2] }),
  ], 820)),
]);

// Aufgabe 4: Bildebene → Sachebene (das Tor macht die Lehrkraft vor)
const a4W = [2600, 1100, W - 3700];
const a4 = table(a4W, [
  row([kopf("Bildebene", a4W[0]), kopf("Zeile", a4W[1]), kopf("Sachebene: Wofür steht es? Begründe mit dem Text.", a4W[2])]),
  ...D.SYMBOLE.map(([b], i) => row([
    cell(p(t(b, { bold: true, color: RED, size: 20 }), { after: 0 }), { w: a4W[0], fill: LIGHT }),
    cell(p(t("vgl. Z. ", { size: 19, color: MUTED }), { after: 0 }), { w: a4W[1] }),
    cell(p(t(""), { after: 0 }), { w: a4W[2] }),
  ], 1000)),
]);

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Vor dem Gesetz“", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t("Franz Kafka: „Vor dem Gesetz“ (1915), in heutiger Rechtschreibung", { size: 16, color: MUTED })] })] });

const seite1 = [
  p(t("Arbeitsblatt: Gespräch und Symbole in Kafkas „Vor dem Gesetz“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 50 }),
  p([t("Thema: ", { bold: true, size: 20 }), t("Ohnmacht vor der Macht – eine kafkaeske Parabel", { size: 20 })], { after: 20 }),
  p([t("Ziel: ", { bold: true, size: 20 }), t("Ich kann die Handlung über das Gespräch erschließen, die Symbole deuten und erklären, warum die Parabel kafkaesk ist.", { size: 20 })], { after: 100 }),
  box([
    p(t("Info-Box: Kontext", { bold: true, color: RED, size: 21 }), { after: 30 }),
    p(t(D.KONTEXT[0] + " Im Dom erzählt ein Geistlicher sie Josef K., der verhaftet wurde, ohne zu erfahren, warum.", { size: 19 }), { after: 0, line: 270 }),
  ]),
  h2("1. Textgrundlage: Franz Kafka, „Vor dem Gesetz“", { before: 130 }),
  textBlock(D.ZEILEN),
];

const seite2 = [
  h2("2. Arbeitsaufträge", { pb: true, before: 0 }),
  aufgabeKopf(D.TATT),
  p([t("Bausteine: ", { bold: true, size: 19 }), t(D.TATT.bausteine, { size: 19 })], { after: 30 }),
  schreibLinien(4, W, 400),
  aufgabeKopf(D.AUFGABE3),
  ...schritte(D.AUFGABE3.schritte),
  gespraech,
  ...D.AUFGABE3.fragen.flatMap((q) => [p(t(q, { bold: true, size: 19, color: NAVY }), { before: 80, after: 0, keepNext: true }), schreibLinien(2, W, 380)]),
  aufgabeKopf(D.AUFGABE4),
  ...schritte(D.AUFGABE4.schritte),
  a4,
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 760, bottom: 620, left: 850, right: 850, header: 340, footer: 340 } } }, headers: { default: header }, footers: { default: footer }, children: [...seite1, ...seite2] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Arbeitsblatt_Vor_dem_Gesetz.docx", b));
