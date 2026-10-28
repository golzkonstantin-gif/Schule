const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType,
  VerticalAlign, HeightRule, PageBreak } = require("docx");
const { D, C, ORDER, NR, card } = require("./common.js");
const OUT = process.argv[2];
const W = 10206, HALF = W / 2;
const dash = { style: BorderStyle.DASHED, size: 6, color: "999999" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ alignment: o.align, spacing: { after: o.after ?? 0, before: o.before ?? 0 }, children: runs });

function chip(id) {
  const k = card(id);
  return new TableCell({ width: { size: HALF, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 140, bottom: 140, left: 180, right: 180 },
    children: [P([R(String(NR[id]) + "  ", { bold: true, color: C.terra, size: 26, font: "Cambria" }), R(k.h, { bold: true, size: 23, font: "Cambria" })], { after: 80 }), P([R(k.t, { size: 21 })])] });
}
function set() {
  const rows = [];
  for (let i = 0; i < ORDER.length; i += 2)
    rows.push(new TableRow({ cantSplit: true, height: { value: 2600, rule: HeightRule.ATLEAST }, children: [chip(ORDER[i]), chip(ORDER[i + 1])] }));
  return [
    P([R("GESCHICHTE · KLASSE 11 · MYSTERY", { bold: true, size: 16, color: C.terra })], { after: 60 }),
    new Table({ columnWidths: [W], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [new TableCell({ width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: C.sand }, borders: { top: none, bottom: none, left: none, right: none }, margins: { top: 120, bottom: 120, left: 180, right: 180 },
      children: [P([R("Warum stimmt Lysias anders?", { font: "Cambria", bold: true, size: 30 })], { after: 60 }), P([R(D.frage, { size: 21, italics: true })])] })] })] }),
    P([R("Schneidet die acht Abschnitte an den gestrichelten Linien aus.", { size: 17, color: C.muted, italics: true })], { before: 80, after: 80 }),
    new Table({ columnWidths: [HALF, HALF], width: { size: W, type: WidthType.DXA }, rows }),
    P([R("Verbindungswörter: ", { bold: true, size: 21, color: C.terra }), R("Schreibt sie zwischen die Abschnitte und streicht sie hier ab.", { size: 19, italics: true, color: C.muted })], { before: 160, after: 80 }),
    new Table({ columnWidths: Array(4).fill(W / 4), width: { size: W, type: WidthType.DXA }, rows: [0, 4].map(o => new TableRow({ children: D.verbinder.slice(o, o + 4).map(v =>
      new TableCell({ width: { size: W / 4, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: { top: none, bottom: none, left: none, right: none },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: C.sand }, margins: { top: 90, bottom: 90, left: 60, right: 60 },
        children: [P([R(v, { bold: true, size: 28, font: "Cambria" })], { align: AlignmentType.CENTER })] })) })) }),
  ];
}
const HILFEN = [
  ["Hilfe 1", "Sortiert die Abschnitte zuerst in drei Stapel: VORHER (Lysias' Lage) – DIE WENDE (was passiert in Athen?) – NACHHER (was ist neu?)."],
  ["Hilfe 2", `Legt Abschnitt ${NR.B} und Abschnitt ${NR.F} nebeneinander. Worin unterscheiden sich die alten und die neuen Phylen?`],
  ["Hilfe 3", `Lest Abschnitt ${NR.F} (letzter Satz) und ${NR.G}: Wo landen die Anhänger des Nachbarn – und mit wem ist Lysias jetzt zusammen? Was bedeutet das für seine Abstimmung?`],
];
function hilfen() {
  return [
    P([R("GESTUFTE HILFEN · AM PULT AUSLEGEN · AN DEN LINIEN SCHNEIDEN", { bold: true, size: 16, color: C.terra })], { after: 120 }),
    new Table({ columnWidths: [W], width: { size: W, type: WidthType.DXA }, rows: HILFEN.map(([t, d]) => new TableRow({ cantSplit: true, height: { value: 2600, rule: HeightRule.ATLEAST },
      children: [new TableCell({ width: { size: W, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: { top: dash, bottom: dash, left: dash, right: dash },
        margins: { top: 200, bottom: 200, left: 400, right: 400 },
        children: [P([R(t, { font: "Cambria", bold: true, size: 40, color: C.terra })], { after: 120, align: AlignmentType.CENTER }), P([R(d, { size: 30 })], { align: AlignmentType.CENTER })] })] })) }),
  ];
}
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 500, left: 850, right: 850 } } },
  children: [...set(), new Paragraph({ children: [new PageBreak()] }), ...hilfen()] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
