// Fragenkärtchen zum Ausschneiden: 4 × A6 auf A4 hoch (neben den Text legen, kein Umdrehen). Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, HeightRule } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 10906, HW = W / 2;
const dash = { style: BorderStyle.DASHED, size: 6, color: "999999" };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, indent: o.indent, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const card = () => new TableCell({ width: { size: HW, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 200, bottom: 120, left: 260, right: 260 }, children: [
  P("FRAGEN ZUM TEXT · STICHWORTE INS HEFT", { bold: true, size: 16, color: C.terra, after: 140 }),
  ...D.fragen.map(([q], i) => P([R((i + 1) + "  ", { bold: true, color: C.terra, font: "Cambria", size: 26 }), R(q + (i === 7 ? " *" : ""), { size: 24 }),
    ...(i === 0 ? [R("  (Beispiel)", { size: 15, italics: true, color: C.muted })] : [])], { after: 200, indent: { left: 340, hanging: 340 } })),
  P("* nur wenn du fertig bist · Fest? Hand heben.", { size: 18, italics: true, color: C.muted, before: 20 }),
] });
const children = [new Table({ columnWidths: [HW, HW], width: { size: W, type: WidthType.DXA },
  rows: [0, 1].map(() => new TableRow({ cantSplit: true, height: { value: 7700, rule: HeightRule.EXACT }, children: [card(), card()] })) })];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 400, left: 500, right: 500 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
