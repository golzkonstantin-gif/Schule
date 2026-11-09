// Textblatt (A4 hoch, einseitig): eigener Text mit festen Zeilen und Zeilennummern. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, TabStopType, BorderStyle } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0, line: o.line }, indent: o.indent, border: o.border, children: Array.isArray(runs) ? runs : [R(runs, o)] });
let n = 0;
const zeile = (t, first) => { n++; return new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 400 }, { type: TabStopType.LEFT, position: 640 }],
  spacing: { after: 0, before: first ? 110 : 0, line: 340 },
  children: [R("\t" + (n === 1 || n % 5 === 0 ? n : ""), { size: 17, color: C.soft }), R("\t" + t, { size: 23 })] }); };

const children = [
  P("GESCHICHTE · KLASSE 11 · TEXT", { bold: true, size: 16, color: C.terra, after: 30 }),
  P([R("Vom Perserkrieg zur Demokratie: Athen von Themistokles bis Perikles", { font: "Cambria", bold: true, size: 30 })], { after: 60 }),
  P("Lies den Text. Beantworte die Fragen in Stichworten in der Fragentabelle.", { size: 19, italics: true, color: C.muted, after: 60 }),
  ...D.text.flatMap(abs => abs.map((t, i) => zeile(t, i === 0))),
  P(D.nachweis, { size: 16, color: C.muted, before: 160, after: 140 }),
  P([R("Stratege", { bold: true, size: 20 }), R(" = einer der zehn Heerführer Athens, jedes Jahr gewählt · ", { size: 20 }), R("Diäten", { bold: true, size: 20 }), R(" = Geld für jeden Tag im Amt · ", { size: 20 }),
    R("Areopag", { bold: true, size: 20 }), R(" = Rat der ehemaligen Archonten · ", { size: 20 }), R("Theten", { bold: true, size: 20 }), R(" = unterste Vermögensklasse", { size: 20 })],
    { indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 992, right: 992 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
