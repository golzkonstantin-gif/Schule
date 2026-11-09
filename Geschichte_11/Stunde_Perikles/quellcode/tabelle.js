// Fragentabelle zum Ausfüllen (A4 hoch, eigenes Blatt neben dem Textblatt): Frage | Antwort in Stichworten. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, HeightRule } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 10206;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const cw = [3900, 6306];
const cell = (children, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.TOP, margins: { top: 100, bottom: 100, left: 120, right: 120 },
  borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children });
const head = new TableRow({ tableHeader: true, children: ["Frage", "Antwort in Stichworten"].map((t, i) => cell([P(t, { bold: true })], i, { head: true })) });
const rows = D.fragen.map(([q], i) => new TableRow({ cantSplit: true, height: { value: 1650, rule: HeightRule.ATLEAST }, children: [
  cell([P([R((i + 1) + "  ", { bold: true, color: C.terra, font: "Cambria", size: 26 }), R(q, { size: 22 })])], 0),
  cell([P("")], 1)] }));
const children = [
  P("GESCHICHTE · KLASSE 11 · FRAGEN ZUM TEXT", { bold: true, size: 16, color: C.terra, after: 30 }),
  P([R("Vom Perserkrieg zur Demokratie", { font: "Cambria", bold: true, size: 30 })], { after: 140 }),
  new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [head, ...rows] }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 640, bottom: 500, left: 850, right: 850 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
