// Fragenblatt (2 × A5 auf A4 quer, Antworten ins Heft) und Hilfekarten mit Fundstellen (Seite 2, am Pult). Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, PageOrientation, PageBreak, HeightRule, VerticalAlign } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const HALF = 7300, GAP = 640, W = HALF * 2 + GAP;
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, nob = { top: none, bottom: none, left: none, right: none };
const dash = { style: BorderStyle.DASHED, size: 6, color: "999999" };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, indent: o.indent, border: o.border, children: Array.isArray(runs) ? runs : [R(runs, o)] });

function half() {
  return [
    P("GESCHICHTE · KLASSE 11 · TEXTARBEIT", { bold: true, size: 15, color: C.terra, after: 30 }),
    P([R(D.titel, { font: "Cambria", bold: true, size: 32 })], { after: 80 }),
    P("Lehrbuch S. 17 ab „Die Perserkriege und ihre Folgen“ bis S. 18 oben. Antworte in Stichworten im Heft – in eigenen Worten.", { size: 21, italics: true, color: C.muted, after: 200 }),
    ...D.fragen.map(([q], i) => P([R((i + 1) + "  ", { bold: true, color: C.terra, font: "Cambria", size: 27 }), R(q + (i === 7 ? "  *" : ""), { size: 25 }),
      ...(i === 0 ? [R("  (Beispiel, gemeinsam)", { size: 17, italics: true, color: C.muted })] : [])], { after: 230, indent: { left: 380, hanging: 380 } })),
    P("* nur wenn du fertig bist (Randspalte zu Perikles)", { size: 19, italics: true, color: C.muted, before: 40, after: 160 }),
    P([R("Hilfe: ", { bold: true, size: 21, color: C.terra }), R("Fundstelle zu jeder Frage als Karte am Pult.", { size: 21 })], { indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } }),
  ];
}
const outer = new Table({ columnWidths: [HALF, GAP, HALF], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: nob }),
  new TableCell({ children: [P("")], width: { size: GAP, type: WidthType.DXA }, borders: { ...nob, right: dash } }),
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: nob, margins: { left: 200 } }),
] })] });

// Hilfekarten: Fundstellen (keine Lösungen – die bleiben am Pult bei der Lehrkraft)
const card = i => new TableCell({ width: { size: W / 2, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, verticalAlign: VerticalAlign.CENTER,
  margins: { top: 120, bottom: 120, left: 240, right: 240 }, children: [
    P([R("Frage " + (i + 1), { font: "Cambria", bold: true, size: 30, color: C.terra })], { after: 60 }),
    P(D.fragen[i][0], { size: 20, italics: true, color: C.muted, after: 80 }),
    P([R("Hier steht es: ", { bold: true, size: 24 }), R(D.fragen[i][1], { size: 24 })]),
  ] });
const karten = [
  new Paragraph({ children: [new PageBreak()] }),
  P("HILFEKARTEN · FUNDSTELLEN · AM PULT AUSLEGEN · AN DEN LINIEN SCHNEIDEN", { bold: true, size: 16, color: C.terra, after: 120 }),
  new Table({ columnWidths: [W / 2, W / 2], width: { size: W, type: WidthType.DXA }, rows: [0, 2, 4, 6].map(i => new TableRow({ cantSplit: true, height: { value: 2200, rule: HeightRule.ATLEAST }, children: [card(i), card(i + 1)] })) }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 400, left: 480, right: 480 } } }, children: [outer, ...karten] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
