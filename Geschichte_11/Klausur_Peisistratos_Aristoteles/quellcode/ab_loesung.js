// Lösung zum Arbeitsblatt TATTEA + Einleitung – im Layout des Arbeitsblatts (2 × A5 auf A4 quer), damit Feld für Feld verglichen werden kann.
// Einleitung nach den vier Bausteinen des Blatts, benutzte Satzanfänge fett. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, HeightRule, PageOrientation, AlignmentType } = require("docx");
const { C } = require("../../Probeklausur_PseudoXenophon/quellcode/lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const HALF = 7300, GAP = 640, W = HALF * 2 + GAP;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, terra = { style: BorderStyle.SINGLE, size: 10, color: C.terra };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold, color: x.startsWith("**") ? C.terra : o.color }));
const P = (t, o = {}) => new Paragraph({ alignment: o.align, spacing: { after: o.after ?? 40, before: o.before ?? 0, line: o.line ?? 250 }, indent: o.indent, keepNext: o.keepNext, children: Array.isArray(t) ? t : runs(t, o) });
const cell = (children, w, o = {}) => new TableCell({ children, width: { size: w, type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.CENTER, borders: o.b || { top: line, bottom: line, left: line, right: line }, margins: { top: 40, bottom: 40, left: 90, right: 90 } });
const TW = HALF - 400, cw = [1700, TW - 1700];
const NAMEN = ["Titel", "Autor", "Textsorte", "Thema", "Entstehung (Zeit, Ort)", "Adressat"];
const EIN = [
  ["1  Textsorte, Titel, Autor", "**Bei der vorliegenden Quelle handelt es sich um einen Auszug aus** der Geschichtsschrift „Athenaion politeia“ („Der Staat der Athener“), Kapitel 16. **Als Verfasser gilt** der Philosoph Aristoteles; **allerdings ist umstritten,** ob er die Schrift wirklich verfasst hat."],
  ["2  Entstehung und Adressat", "**Die Schrift entstand um** 330 v. Chr. **in** Athen, **also rund** 200 **Jahre nach** der Herrschaft des Peisistratos. **Sie richtet sich an** gebildete Leser, die sich für die Geschichte der athenischen Verfassung interessieren."],
  ["3  Thema", "**Sie behandelt** die Herrschaft des Peisistratos, **vor allem die Frage,** wie er regierte und warum er sich so lange an der Macht halten konnte."],
  ["4  Themenfrage", "**Die Untersuchung geht der Frage nach, ob** Peisistratos’ Politik Ausdruck politischer Fürsorge oder gezielter Machtsicherung war."],
];
function half() {
  return [
    P([R("GESCHICHTE · KLASSE 11 · ÜBUNGSKLAUSUR PEISISTRATOS", { bold: true, size: 15, color: C.terra }), R("   LÖSUNG", { bold: true, size: 15, color: C.terra })], { after: 20 }),
    P([R("Die Einleitung schreiben – Lösung", { font: "Cambria", bold: true, size: 30 })], { after: 60 }),
    P([R("1  ", { font: "Cambria", bold: true, size: 24, color: C.terra }), R("TATTEA bestimmen", { font: "Cambria", bold: true, size: 24 })], { after: 60 }),
    new Table({ columnWidths: cw, width: { size: TW, type: WidthType.DXA }, rows: [
      new TableRow({ tableHeader: true, children: [cell([P("TATTEA", { bold: true, size: 19, after: 0 })], cw[0], { b: { top: line, left: line, right: line, bottom: terra } }), cell([P("Antwort in Stichworten", { bold: true, size: 19, after: 0 })], cw[1], { b: { top: line, left: line, right: line, bottom: terra } })] }),
      ...D.tattea.map(([, a], i) => new TableRow({ cantSplit: true, height: { value: 830, rule: HeightRule.ATLEAST },
        children: [cell([P(NAMEN[i], { bold: true, size: 21, after: 0 })], cw[0]), cell([P(a, { size: 20, after: 0, color: C.terra, bold: true, line: 240 })], cw[1])] })),
    ] }),
    P([R("2  ", { font: "Cambria", bold: true, size: 24, color: C.terra }), R("Einleitung schreiben", { font: "Cambria", bold: true, size: 24 }), R("   fett = Satzanfänge", { size: 17, italics: true, color: C.muted })], { before: 160, after: 60 }),
    ...EIN.flatMap(([t, x]) => [P([R(t, { bold: true, size: 19, color: C.terra })], { after: 10, before: 30, keepNext: true }), P(x, { size: 20, line: 255, after: 30, align: AlignmentType.JUSTIFIED })]),
  ];
}
const outer = new Table({ columnWidths: [HALF, GAP, HALF], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none } }),
  new TableCell({ children: [P("")], width: { size: GAP, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: { style: BorderStyle.DASHED, size: 6, color: "999999" } } }),
  new TableCell({ children: half(), width: { size: HALF, type: WidthType.DXA }, borders: { top: none, bottom: none, right: none, left: none }, margins: { left: 320 } }),
] })] });
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 480, right: 480 } } }, children: [outer] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
