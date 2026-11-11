// Aufgabenblatt (2 × A5 auf A4 quer, eigenes Blatt neben der Quelle): Aufgaben in der Reihenfolge der Lehrkraft. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, PageOrientation } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const HALF = 7300, GAP = 640, W = HALF * 2 + GAP;
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, nob = { top: none, bottom: none, left: none, right: none };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, indent: o.indent, border: o.border, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const AUFGABEN = [
  ["Formuliere die Themenfrage und eine Einleitung (formale Analyse nach TATTEA).", "allein · Kontrolle: Rückseite der Tafel"],
  ["Arbeite die Position des Verfassers zur attischen Demokratie und zu ihren Bürgern heraus – in Stichpunkten mit Zeilenangaben.", "wir tragen zusammen"],
  ["Beschreibe die wichtigsten Verfassungsorgane der athenischen Demokratie unter Kleisthenes sowie die gesellschaftliche Gliederung der Bevölkerung mit ihren Rechten.", "Kontrolle: Aushang"],
  ["Beurteile, ob die Reformen des Kleisthenes der Vormachtstellung des Adels in der attischen Polis ein Ende gesetzt haben.", "Kontrolle: Aushang"],
];
function half() {
  return [
    P("GESCHICHTE · KLASSE 11 · PROBEKLAUSUR", { bold: true, size: 15, color: C.terra, after: 30 }),
    P([R(D.titel, { font: "Cambria", bold: true, size: 30 })], { after: 60 }),
    P("Aufgabenart: Interpretieren von Quellen. Nutze die Übersicht „So interpretierst du eine Textquelle“.", { size: 18, italics: true, color: C.muted, after: 200 }),
    ...AUFGABEN.flatMap(([a, k], i) => [
      P([R((i + 1) + "  ", { bold: true, color: C.terra, font: "Cambria", size: 28 }), R(a, { size: 23 })], { indent: { left: 380, hanging: 380 }, after: 40 }),
      P(k, { size: 17, italics: true, color: C.muted, indent: { left: 380 }, after: 220 }),
    ]),
  ];
}
const outer = new Table({ columnWidths: [HALF, GAP, HALF], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: nob }),
  new TableCell({ children: [P("")], width: { size: GAP, type: WidthType.DXA }, borders: { ...nob, right: { style: BorderStyle.DASHED, size: 6, color: "999999" } } }),
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: nob, margins: { left: 200 } }),
] })] });
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 400, left: 480, right: 480 } } }, children: [outer] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
