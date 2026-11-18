// Arbeitsblatt TATTEA + Einleitung, 2 × A5 auf A4 quer. Tabelle ohne Hinweise (nur die Fragen), Einleitung ins Heft mit Satzanfängen zum Abhaken. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, HeightRule, PageOrientation, TabStopType } = require("docx");
const { C } = require("../../Probeklausur_PseudoXenophon/quellcode/lib.js");
const OUT = process.argv[2];
const HALF = 7300, GAP = 640, W = HALF * 2 + GAP;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, terra = { style: BorderStyle.SINGLE, size: 10, color: C.terra };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const P = (t, o = {}) => new Paragraph({ spacing: { after: o.after ?? 40, before: o.before ?? 0, line: o.line ?? 250 }, indent: o.indent, tabStops: o.tabs, keepNext: o.keepNext, children: Array.isArray(t) ? t : runs(t, o) });
const cell = (children, w, o = {}) => new TableCell({ children, width: { size: w, type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.CENTER, borders: o.b || { top: line, bottom: line, left: line, right: line }, margins: { top: 40, bottom: 40, left: 90, right: 90 } });
const cw = [1700, HALF - 1700];
const BAU = [
  ["1", "Textsorte, Titel, Autor", ["„Bei der vorliegenden Quelle handelt es sich um einen Auszug aus …“", "„Als Verfasser gilt …; allerdings ist umstritten, …“"]],
  ["2", "Entstehung und Adressat", ["„Die Schrift entstand um … in …, also rund … Jahre nach …“", "„Sie richtet sich an …“"]],
  ["3", "Thema", ["„Sie behandelt …, vor allem die Frage, …“"]],
  ["4", "Themenfrage", ["„Die Untersuchung geht der Frage nach, ob …“"]],
];
function half() {
  return [
    P("GESCHICHTE · KLASSE 11 · ÜBUNGSKLAUSUR PEISISTRATOS", { bold: true, size: 15, color: C.terra, after: 20 }),
    P([R("Die Einleitung schreiben", { font: "Cambria", bold: true, size: 30 })], { after: 60 }),
    P([R("1  ", { font: "Cambria", bold: true, size: 24, color: C.terra }), R("TATTEA bestimmen", { font: "Cambria", bold: true, size: 24 }), R("   Stichworte · allein", { size: 17, italics: true, color: C.muted })], { after: 60 }),
    new Table({ columnWidths: cw, width: { size: HALF, type: WidthType.DXA }, rows: [
      new TableRow({ tableHeader: true, children: [cell([P("TATTEA", { bold: true, size: 19, after: 0 })], cw[0], { b: { top: line, left: line, right: line, bottom: terra } }), cell([P("Antwort in Stichworten", { bold: true, size: 19, after: 0 })], cw[1], { b: { top: line, left: line, right: line, bottom: terra } })] }),
      ...["Titel", "Autor", "Textsorte", "Thema", "Entstehung (Zeit, Ort)", "Adressat"].map(t => new TableRow({ cantSplit: true, height: { value: 830, rule: HeightRule.ATLEAST },
        children: [cell([P(t, { bold: true, size: 21, after: 0 })], cw[0]), cell([P("", { after: 0 })], cw[1])] })),
    ] }),
    P([R("2  ", { font: "Cambria", bold: true, size: 24, color: C.terra }), R("Einleitung schreiben", { font: "Cambria", bold: true, size: 24 }), R("   ins Heft · hake ab, was du nutzt", { size: 17, italics: true, color: C.muted })], { before: 160, after: 60 }),
    ...BAU.flatMap(([n, t, s]) => [
      P([R(n + "  " + t, { bold: true, size: 21, color: C.terra })], { after: 20, before: 40, keepNext: true }),
      ...s.map(x => P([R("☐  ", { size: 22, color: C.terra }), R(x, { size: 21 })], { indent: { left: 360, hanging: 360 }, after: 50 })),
    ]),
  ];
}
const outer = new Table({ columnWidths: [HALF, GAP, HALF], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none } }),
  new TableCell({ children: [P("")], width: { size: GAP, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: { style: BorderStyle.DASHED, size: 6, color: "999999" } } }),
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, right: none, left: none }, margins: { left: 320 } }),
] })] });
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 480, right: 480 } } }, children: [outer] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
