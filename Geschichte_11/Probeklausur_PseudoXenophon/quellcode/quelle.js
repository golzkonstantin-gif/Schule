// Quellenblatt (A4 hoch, einseitig): Pseudo-Xenophon im Blocksatz mit festen Zeilen und Zeilennummern (Word-Zeilennummerierung).
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, BorderStyle, AlignmentType, LineNumberRestartFormat } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ suppressLineNumbers: true, alignment: o.align, spacing: { after: o.after ?? 0, before: o.before ?? 0, line: o.line }, indent: o.indent, border: o.border, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const text = new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 0, before: 120, line: 360 },
  children: D.zeilen.map((t, i) => new TextRun({ text: t, font: "Calibri", size: 25, color: C.dark, break: i ? 1 : 0 })) });
const children = [
  P("GESCHICHTE · KLASSE 11 · PROBEKLAUSUR · QUELLE", { bold: true, size: 16, color: C.terra, after: 30 }),
  P([R(D.quelleTitel, { font: "Cambria", bold: true, size: 30 })], { after: 80 }),
  P(D.info, { size: 20, italics: true, color: C.muted, after: 60, indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } }),
  text,
  P(D.nachweis, { size: 16, color: C.muted, before: 160 }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 1250, right: 992 } },
  lineNumbers: { countBy: 5, restart: LineNumberRestartFormat.CONTINUOUS, distance: 300 } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
