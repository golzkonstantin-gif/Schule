const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign } = require("docx");
const { D, C, NR, card } = require("./common.js");
const OUT = process.argv[2];
const W = 10206;
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const dash = { style: BorderStyle.DASHED, size: 6, color: "999999" };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ alignment: o.align, spacing: { after: o.after ?? 0, before: o.before ?? 0 }, keepNext: o.keepNext, children: runs });
const one = (children, o = {}) => new Table({ columnWidths: [W], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ cantSplit: true, children: [new TableCell({
  width: { size: W, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
  borders: o.dash ? { top: dash, bottom: dash, left: dash, right: dash } : { top: none, bottom: none, left: none, right: none },
  margins: { top: o.pad ?? 90, bottom: o.pad ?? 90, left: 180, right: 180 }, children })] })] });
// Verbinder als schmale Zeile mit Pfeil
const link = v => P([R("↓  ", { size: 22, color: C.soft, bold: true }), R(v, { size: 24, bold: true, color: C.terra, font: "Cambria" })], { align: AlignmentType.CENTER, before: 30, after: 30, keepNext: true });

const children = [
  P([R("GESCHICHTE · KLASSE 11 · MYSTERY · LÖSUNG (LEHRKRAFT)", { bold: true, size: 16, color: C.terra })], { after: 60 }),
  P([R("Warum stimmt Lysias anders? – die Abschnitte in Verbindung", { font: "Cambria", bold: true, size: 30 })], { after: 60 }),
  P([R("Eine mögliche Lösung. Andere Verbindungen sind richtig, wenn sie den Zusammenhang sinnvoll begründen. Die Nummern entsprechen dem Schülerblatt.", { size: 17, italics: true, color: C.muted })], { after: 120 }),
];
D.loesung.forEach(([v, id], i) => {
  if (v) children.push(link(v));
  if (id === "Ergebnis") {
    children.push(one([P([R("Ergebnis: ", { bold: true, size: 22, color: C.gold, font: "Cambria" }), R("Lysias ist nicht mehr an seinen Adelsherrn gebunden – er kann sein Stimmrecht frei nutzen. Kleisthenes hat nicht die Rechte verändert (die gab schon Solon), sondern die Bindungen an die Adelsfamilien.", { size: 20, color: C.white, bold: true })])], { fill: C.dark, pad: 140 }));
    return;
  }
  const k = card(id);
  const zusatz = id === "H";
  children.push(one([
    P([R(String(NR[id]) + "  ", { bold: true, color: C.terra, size: 24, font: "Cambria" }), R(k.h, { bold: true, size: 22, font: "Cambria" }), ...(zusatz ? [R("   (Zusatz)", { size: 18, italics: true, color: C.muted })] : [])], { after: 40 }),
    P([R(k.t, { size: 19 })]),
  ], { dash: true, fill: zusatz ? "FFFFFF" : C.pale }));
});
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 500, left: 850, right: 850 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
