// Arbeitsblatt „Gibs auf!“ nach dem bekannten Muster (Kafka-Appell): Titel, Thema, Ziel, Info-Box Parabel,
// 1. Textgrundlage mit Zeilennummern, 2. Information Stilmittel, 3. Arbeitsaufträge mit Schreiblinien (2 Seiten, doppelseitig)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Header, Footer, TabStopType,
} = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GOLD = "E8A33D", BLUE = "3F7CC4";
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
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none), m: 100, ml: 160 })])]);
const h2 = (text, o = {}) => p(t(text, { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: o.before ?? 180, after: 70, pb: o.pb, keepNext: true });
const aufgabe = (a, nr, pb) => p([t(a.titel, { bold: true, color: NAVY, size: 22 }), t("   (" + a.form + ", " + a.zeit + ")", { italics: true, color: MUTED, size: 18 })], { before: pb ? 0 : 140, after: 40, keepNext: true, pb });
const schritte = (arr) => arr.map((x, i) => p([t((i + 1) + ".  ", { bold: true, color: NAVY, size: 20 }), t(x, { size: 20 })], { after: 30 }));
const schreibLinien = (n, w, h = 420) => table([w], Array.from({ length: n }, () => row([cell(p(t(""), { after: 0 }), { w, borders: { top: none, left: none, right: none, bottom: solid("9AA6C4", 4) }, m: 0, ml: 0 })], h)));

// Text mit Zeilennummern
function textBlock(lines) {
  const nW = 520, tW = W - nW;
  return table([nW, tW], lines.map((l, i) => row([
    cell(p(t(String(i + 1), { size: 16, color: MUTED }), { align: AlignmentType.RIGHT, after: 0 }), { w: nW, borders: { top: none, bottom: none, left: none, right: solid(BLUE, 10) }, m: 6, ml: 80, valign: VerticalAlign.CENTER }),
    cell(p(t(l, { font: "Cambria", size: 23 }), { after: 0, line: 340 }), { w: tW, borders: all(none), m: 6, ml: 140 }),
  ])));
}

// Stilmittel-Tabelle (bekannt + neu)
const mW = [1900, 4400, W - 6300];
const kopf = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w, fill: NAVY });
const bekannt = [
  ["Parataxe", "Reihung von Hauptsätzen", "wirkt nüchtern oder gehetzt"],
  ["Hypotaxe", "Satzgefüge mit Nebensätzen", "wirkt verschachtelt, oft verwirrt"],
  ["Metapher", "bildhafter Ausdruck im übertragenen Sinn", "z. B. „Rabenmutter“"],
  ["Symbol", "konkreter Gegenstand für etwas Abstraktes", "z. B. Uhr für Zeitdruck"],
];
const stilTabelle = table(mW, [
  row([kopf("Stilmittel", mW[0]), kopf("So erkennst du es", mW[1]), kopf("Wirkung / Test", mW[2])]),
  ...bekannt.map(([b, e, h]) => row([
    cell(p(t(b, { bold: true, color: NAVY, size: 19 }), { after: 0 }), { w: mW[0], fill: LIGHT }),
    cell(p(t(e, { size: 19 }), { after: 0 }), { w: mW[1] }),
    cell(p(t(h, { size: 18, color: MUTED }), { after: 0 }), { w: mW[2] }),
  ])),
  row([cell(p(t("neu", { bold: true, size: 16, color: RED }), { after: 0 }), { w: W, span: 3, fill: "FDECEA", m: 10 })]),
  ...D.STILMITTEL.map(([b, kern, , test, , bsp]) => row([
    cell(p(t(b, { bold: true, color: RED, size: 19 }), { after: 0 }), { w: mW[0], fill: LIGHT }),
    cell([p(t(kern, { size: 19 }), { after: 10 }), p(t("Beispiel: " + bsp, { size: 18, italics: true, color: BLUE }), { after: 0 })], { w: mW[1] }),
    cell(p(t("Test: " + test, { size: 18, color: MUTED }), { after: 0 }), { w: mW[2] }),
  ])),
]);

// Aufgabe 1: Tabelle zum Eintragen
const a1W = [2000, 3700, W - 5700];
const a1 = table(a1W, [
  row([kopf("Stilmittel", a1W[0]), kopf("Zitat (Zeile)", a1W[1]), kopf("Wirkung: „Dadurch wirkt …“", a1W[2])]),
  ...["Vergleich", "Rhetorische Frage", "Ellipse"].map((b) => row([
    cell(p(t(b, { bold: true, color: RED, size: 19 }), { after: 0 }), { w: a1W[0], fill: LIGHT, valign: VerticalAlign.CENTER }),
    cell(p(t(""), { after: 0 }), { w: a1W[1] }),
    cell(p(t(""), { after: 0 }), { w: a1W[2] }),
  ], 900)),
]);

// Schreibhilfe kompakt in drei Spalten
const sW = Math.floor(W / 3);
const schreibhilfe = table([sW, sW, W - 2 * sW], [row(D.SCHREIBHILFE.map(([k, arr], i) => cell([
  p(t(k, { bold: true, color: i === 2 ? GOLD : NAVY, size: 18 }), { after: 20 }),
  ...arr.map((x) => p(t(x, { size: 17 }), { after: 10 })),
], { w: i === 2 ? W - 2 * sW : sW, fill: i === 2 ? "FFF8EC" : LIGHT, borders: all(none), m: 60 })))]);

// Aufgabe 3*: Satz + Quadrat mit Linien
const qW = W / 2;
const quad = (label, focus) => cell([
  p(t("[ " + label + " ]", { bold: true, size: 18, color: focus ? RED : NAVY }), { after: 0 }),
  schreibLinien(2, qW - 260, 380),
  p(t(""), { after: 0 }),
], { w: qW, fill: focus ? "FDECEA" : "FFFFFF", borders: all(solid(focus ? RED : "B9C6E8", focus ? 12 : 6)), m: 50, ml: 120 });
const quadrat = table([qW, qW], [
  row([quad("SACHINHALT"), quad("SELBSTKUNDGABE")]),
  row([quad("BEZIEHUNG"), quad("APPELL", true)]),
]);

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Gibs auf!“", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t("Franz Kafka: „Gibs auf!“ (1922) · Wortlaut nach dem Arbeitsblatt vom 2.11.", { size: 16, color: MUTED })] })] });

const seite1 = [
  p(t("Arbeitsblatt: Stilmittel und Gespräch in Kafkas „Gibs auf!“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 50 }),
  p([t("Thema: ", { bold: true, size: 20 }), t("Stilmittel als Bedeutungsträger in einer Parabel – und ein eigenes „Gibs auf!“ aus dem Sport", { size: 20 })], { after: 20 }),
  p([t("Ziel: ", { bold: true, size: 20 }), t("Ich kann Stilmittel bestimmen und ihre Wirkung erklären und Bild- und Sachebene einer Parabel unterscheiden.", { size: 20 })], { after: 120 }),
  box([
    p(t("Info-Box: Die Parabel", { bold: true, color: RED, size: 21 }), { after: 40 }),
    p([t("Eine Parabel erzählt eine konkrete Geschichte (", { size: 20 }), t("Bildebene", { bold: true, size: 20 }), t("), die auf eine allgemeine menschliche Erfahrung (", { size: 20 }), t("Sachebene", { bold: true, size: 20 }), t(") übertragen werden muss. Die Übertragung leistet der Leser – der Text erklärt sich nicht selbst.", { size: 20 })], { after: 0, line: 290 }),
  ]),
  h2("1. Textgrundlage: Franz Kafka, „Gibs auf!“"),
  textBlock(D.ZEILEN),
  h2("2. Information: Stilmittel"),
  p(t("Stilmittel sind keine Dekoration, sondern Bedeutungsträger: Sie zeigen, wie der Erzähler die Situation erlebt.", { size: 19, italics: true, color: MUTED }), { after: 60 }),
  stilTabelle,
];

const seite2 = [
  h2("3. Arbeitsaufträge", { pb: true, before: 0 }),
  p(t("Aufgaben 1–3 löst du direkt auf diesem Blatt.", { size: 18, italics: true, color: MUTED }), { after: 20 }),
  aufgabe(D.AUFGABE1),
  ...schritte(D.AUFGABE1.schritte),
  a1,
  aufgabe(D.AUFGABE2),
  ...schritte(D.AUFGABE2.schritte),
  schreibhilfe,
  schreibLinien(9, W, 400),
  aufgabe(Object.assign({}, D.AUFGABE3, { titel: D.AUFGABE3.titel + " *" })),
  p([t("Abweisender Satz: ", { bold: true, size: 19 })], { after: 0 }),
  schreibLinien(1, W, 380),
  p(t(""), { after: 40 }),
  quadrat,
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 650, left: 850, right: 850, header: 360, footer: 360 } } }, headers: { default: header }, footers: { default: footer }, children: [...seite1, ...seite2] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Arbeitsblatt_Gibs_auf.docx", b));
