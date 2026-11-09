// Antwortblatt zur Fragentabelle (A4 hoch): Frage, Zeilen, Antwort. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9922;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const cw = [3300, 1000, 5622];
const cell = (children, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 120, bottom: 120, left: 120, right: 120 },
  borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children });
const head = new TableRow({ tableHeader: true, children: ["Frage", "Zeilen", "Antwort"].map((t, i) => cell([P(t, { bold: true, size: 21 })], i, { head: true })) });
const rows = D.fragen.map(([q, z, e], i) => new TableRow({ cantSplit: true, children: [
  cell([P([R((i + 1) + "  ", { bold: true, color: C.terra, font: "Cambria", size: 24 }), R(q, { size: 21 })])], 0),
  cell([P(z, { bold: true, color: C.terra, size: 21 })], 1),
  cell([P(e, { size: 21 })], 2)] }));
const children = [
  P("GESCHICHTE · KLASSE 11 · ANTWORTEN ZUR FRAGENTABELLE", { bold: true, size: 16, color: C.terra, after: 30 }),
  P([R("Vom Perserkrieg zur Demokratie – Antworten", { font: "Cambria", bold: true, size: 30 })], { after: 60 }),
  P("Zeilenangaben beziehen sich auf das Textblatt. In der Tabelle genügen Stichworte; andere Formulierungen sind richtig, wenn der Kern stimmt.", { size: 18, italics: true, color: C.muted, after: 120 }),
  new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [head, ...rows] }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 992, right: 992 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
