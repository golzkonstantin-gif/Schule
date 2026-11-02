// Vorher-Nachher-Tabelle zur Reform des Kleisthenes: 2 × A5 auf A4 quer (Vorher vorgegeben, Nachher aus den Abschnitten ergänzen)
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign, HeightRule, PageOrientation } = require("docx");
const { D, C, NR } = require("./common.js");
const OUT = process.argv[2];
const HALF = 7300, GAP = 640, W = HALF * 2 + GAP;
const thin = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: o.after ?? 0, before: o.before ?? 0 }, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const cell = (children, w, fill) => new TableCell({ children, width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: { top: thin, bottom: thin, left: thin, right: thin },
  shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined, margins: { top: 50, bottom: 50, left: 90, right: 90 } });
const cw = [2000, 2400, 2650];

function half() {
  const head = new TableRow({ tableHeader: true, children: ["", "Vorher", "Nachher (ab 508/07)"].map((t, i) => cell([P(t, { bold: true, color: C.white, size: 19 })], cw[i], C.dark)) });
  const rows = D.vergleich.map(([m, v, n], i) => new TableRow({ cantSplit: true, height: { value: 1150, rule: HeightRule.ATLEAST }, children: [
    cell([P(m, { bold: true, size: 20 })], cw[0], C.sand), cell([P(v, { size: 20 })], cw[1]),
    cell([P(i === 0 ? n : "", { size: 20, italics: i === 0, color: i === 0 ? C.muted : C.dark })], cw[2])] }));
  return [
    P("GESCHICHTE · KLASSE 11 · KLEISTHENES", { bold: true, size: 15, color: C.terra, after: 30 }),
    P("Was hat sich für Lysias geändert?", { font: "Cambria", bold: true, size: 28, after: 50 }),
    P(`Ergänze die Spalte „Nachher“ mit den Abschnitten ${NR.E}, ${NR.F} und ${NR.G}. Die erste Zeile ist ein Beispiel.`, { size: 17, italics: true, color: C.muted, after: 80 }),
    new Table({ columnWidths: cw, width: { size: HALF, type: WidthType.DXA }, rows: [head, ...rows] }),
    P([R("Was bleibt gleich? ", { bold: true, size: 18, color: C.terra }), R("_______________________________________________", { size: 18, color: C.soft })], { before: 120, after: 40 }),
    P("________________________________________________________________", { size: 18, color: C.soft }),
  ];
}
const outer = new Table({ columnWidths: [HALF, GAP, HALF], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none } }),
  new TableCell({ children: [P("")], width: { size: GAP, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: { style: BorderStyle.DASHED, size: 6, color: "999999" } } }),
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, right: none, left: none }, margins: { left: 320 } }),
] })] });
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 480, right: 480 } } }, children: [outer] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
