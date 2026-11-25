// Aufgabenblatt (A4, doppelseitig): vorne Fragentabelle (Frage | Zeilen | Antwort in Stichworten), hinten quer das stumme Schaubild Gesellschaft und Verfassung. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, HeightRule, ImageRun, PageOrientation } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 10206;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const cw = [3700, 900, 5606];
const cell = (children, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP, margins: { top: 100, bottom: 100, left: 120, right: 120 },
  borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children });
const head = new TableRow({ tableHeader: true, children: ["Frage", "Zeilen", "Antwort in Stichworten"].map((t, i) => cell([P(t, { bold: true })], i, { head: true })) });
const rows = D.fragen.map(([q], i) => new TableRow({ cantSplit: true, height: { value: 1640, rule: HeightRule.ATLEAST }, children: [
  cell([P([R((i + 1) + "  ", { bold: true, color: C.terra, font: "Cambria", size: 26 }), R(q, { size: 22 })])], 0),
  cell([P("Z.", { size: 20, color: C.soft })], 1), cell([P("")], 2)] }));
const vorne = [
  P("GESCHICHTE · KLASSE 11 · FRAGEN ZUM TEXT", { bold: true, size: 16, color: C.terra, after: 30 }),
  P([R(D.titel, { font: "Cambria", bold: true, size: 30 })], { after: 120 }),
  new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [head, ...rows] }),
];
const img = fs.readFileSync(__dirname + "/schaubild_stumm.png");
const hinten = [
  P("GESCHICHTE · KLASSE 11 · SCHAUBILD (WIRD GEMEINSAM AUSGEFÜLLT)", { bold: true, size: 16, color: C.terra, after: 60 }),
  new Paragraph({ children: [new ImageRun({ data: img, transformation: { width: 1000, height: 525 }, type: "png" })] }),
];
const doc = new Document({ sections: [
  { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 640, bottom: 500, left: 850, right: 850 } } }, children: vorne },
  { properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 600, right: 600 } } }, children: hinten },
] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
