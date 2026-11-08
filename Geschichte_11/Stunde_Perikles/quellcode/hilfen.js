// Hilfeblatt nur für die Lehrkraft (A4 quer): je Frage gestufte Hilfen zum individuellen Einsatz beim Umhergehen + Lösung. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, PageOrientation, VerticalAlign } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 15838 - 1000;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const cw = [3700, 1300, 3600, 3400, 2838];
const cell = (t, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 130, bottom: 130, left: 110, right: 110 },
  borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children: [P(t, o)] });
const head = new TableRow({ tableHeader: true, children: ["Frage", "Stufe 1: Fundstelle", "Stufe 2: Impuls", "Stufe 3: Satzanfang", "Lösung (Stichwort)"].map((t, i) => cell(t, i, { head: true, bold: true, size: 21 })) });
const rows = D.fragen.map(([q, z, , k, imp, satz], i) => new TableRow({ cantSplit: true, children: [
  cell([R((i + 1) + "  ", { bold: true, color: C.terra, font: "Cambria", size: 21 }), R(q, { size: 21 })], 0),
  cell(z, 1, { bold: true, size: 24, color: C.terra }), cell("„" + imp + "“", 2, { size: 21 }), cell(satz, 3, { size: 21, italics: true }), cell(k, 4, { size: 21, bold: true })] }));
const children = [
  P("NUR FÜR DIE LEHRKRAFT · BEIM UMHERGEHEN", { bold: true, size: 16, color: C.terra, after: 30 }),
  P([R("Gestufte Hilfen zu den Fragen", { font: "Cambria", bold: true, size: 30 })], { after: 40 }),
  P("Immer mit Stufe 1 beginnen; erst wenn es nicht reicht, Stufe 2, dann Stufe 3. Die Lösung nicht vorsagen – sie dient dir zur schnellen Kontrolle im Heft.", { size: 18, italics: true, color: C.muted, after: 100 }),
  new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [head, ...rows] }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 500, right: 500 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
