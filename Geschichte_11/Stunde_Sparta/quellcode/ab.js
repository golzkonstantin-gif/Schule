// Textblatt (A4 hoch, einseitig): eigener Text mit festen Zeilen, Blocksatz und Zeilennummern (Word-Zeilennummerierung). Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, BorderStyle, AlignmentType, LineNumberRestartFormat } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ suppressLineNumbers: true, spacing: { after: o.after ?? 0, before: o.before ?? 0, line: o.line }, indent: o.indent, border: o.border, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const absatz = lines => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 0, before: 90, line: 318 },
  children: lines.map((t, i) => new TextRun({ text: t, font: "Calibri", size: 24, color: C.dark, break: i ? 1 : 0 })) });
const children = [
  P("GESCHICHTE · KLASSE 11 · TEXT", { bold: true, size: 16, color: C.terra, after: 30 }),
  P([R(D.titel, { font: "Cambria", bold: true, size: 30 })], { after: 40 }),
  P("Lies den Text. Beantworte die Fragen in Stichworten in der Fragentabelle.", { size: 19, italics: true, color: C.muted, after: 40 }),
  ...D.text.map(absatz),
  P(D.nachweis, { size: 15, color: C.muted, before: 120, after: 100 }),
  P([R("Spartiaten", { bold: true, size: 19 }), R(" = Vollbürger · ", { size: 19 }), R("Periöken", { bold: true, size: 19 }), R(" = freie Umwohnende ohne politische Rechte · ", { size: 19 }),
    R("Heloten", { bold: true, size: 19 }), R(" = staatseigene Unfreie · ", { size: 19 }), R("Ephoren", { bold: true, size: 19 }), R(" = „Aufseher“ · ", { size: 19 }), R("Gerusia", { bold: true, size: 19 }), R(" = Ältestenrat", { size: 19 })],
    { indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 650, bottom: 500, left: 1250, right: 992 } },
  lineNumbers: { countBy: 5, restart: LineNumberRestartFormat.CONTINUOUS, distance: 300 } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
