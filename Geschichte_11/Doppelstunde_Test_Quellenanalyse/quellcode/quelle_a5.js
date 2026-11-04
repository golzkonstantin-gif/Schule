// Schülerblatt nur mit dem Quellentext (Aristoteles): 2 × A5 auf A4 quer, tintensparend (keine Farbflächen).
// Zeilen fest gesetzt, damit die Zeilennummern mit Schreibplan und Erwartungshorizont übereinstimmen.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, TabStopType, PageOrientation } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const HALF = 7300, GAP = 640, W = HALF * 2 + GAP;
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || C.dark });
// **fett** im Text
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const P = (t, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0, line: o.line }, border: o.border, indent: o.indent, children: Array.isArray(t) ? t : runs(t, o) });
const zeile = (n, t) => new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 340 }, { type: TabStopType.LEFT, position: 520 }], spacing: { after: 0, line: 330 },
  children: [R("\t" + (n === 1 || n % 5 === 0 ? n : ""), { size: 15, color: C.soft }), R("\t" + t, { size: 20 })] });

function half() {
  return [
    P("GESCHICHTE · KLASSE 11 · QUELLE", { size: 15, bold: true, color: C.terra, after: 30 }),
    P([R("Aristoteles über die Reformen des Kleisthenes", { font: "Cambria", bold: true, size: 27 })], { after: 100 }),
    ...D.zeilen.map((t, i) => zeile(i + 1, t)),
    P(D.nachweis, { size: 14, color: C.muted, before: 90, after: 90 }),
    P("**Worterklärungen:** Alkmeoniden = mächtige Adelsfamilie · Isagoras = adliger Gegenspieler des Kleisthenes · Bünde der Adligen = Zusammenschlüsse adliger Freunde", { size: 16, after: 110 }),
    P("**Zum Autor und zur Schrift**", { size: 17, color: C.terra, after: 30, indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } }),
    ...D.info.map(t => P(t, { size: 16, after: 30, indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } })),
  ];
}
const outer = new Table({ columnWidths: [HALF, GAP, HALF], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none } }),
  new TableCell({ children: [P("")], width: { size: GAP, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: { style: BorderStyle.DASHED, size: 6, color: "999999" } } }),
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, right: none, left: none } }),
] })] });
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 400, left: 480, right: 480 } } }, children: [outer] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
