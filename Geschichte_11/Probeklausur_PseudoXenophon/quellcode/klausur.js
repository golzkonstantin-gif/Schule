// Probeklausur im Klausurformat (A4, ein Blatt doppelseitig): Seite 1 Kopf und Aufgaben, Seite 2 Quelle mit Zeilennummern.
// Zeilen wie auf dem Quellenblatt (data.zeilen), damit Aushänge und Folien weiter stimmen. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, BorderStyle, AlignmentType, LineNumberRestartFormat, TabStopType, PageBreak } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 23, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ suppressLineNumbers: true, alignment: o.align, tabStops: o.tabs, keepNext: o.keepNext,
  spacing: { after: o.after ?? 0, before: o.before ?? 0, line: o.line }, indent: o.indent, border: o.border, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const TAB = [{ type: TabStopType.LEFT, position: 2600 }];
const zeile = (k, v, o = {}) => P([R(k, { bold: true }), R("\t" + v, { bold: o.bold, size: o.size })], { tabs: TAB, after: o.after ?? 160, indent: { left: 2600, hanging: 2600 } });
const AUFGABEN = [
  "Formulieren Sie die Themenfrage und eine Einleitung zur Quelle (formale Analyse).",
  "Arbeiten Sie die Positionen des Autors zur attischen Demokratie und zu deren Bürgern heraus.",
  "Beschreiben Sie die wichtigsten Verfassungsorgane der athenischen Demokratie unter Kleisthenes sowie die gesellschaftliche Gliederung der Bevölkerung mit ihren Rechten.",
  "Beurteilen Sie, ob die Reformen des Kleisthenes der Vormachtstellung des Adels in der attischen Polis ein Ende gesetzt haben.",
];
const text = new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 0, before: 120, line: 360 },
  children: D.zeilen.map((t, i) => new TextRun({ text: t, font: "Calibri", size: 25, color: C.dark, break: i ? 1 : 0 })) });

const children = [
  P([R("Klausur im Grundkurs Geschichte", { font: "Cambria", bold: true, size: 34 }), R("\tProbeklausur", { size: 20, italics: true, color: C.muted })],
    { tabs: [{ type: TabStopType.RIGHT, position: 9664 }], after: 300, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.dark, space: 6 } } }),
  P([R("Name: ", { bold: true }), R("______________________________", { color: C.soft }), R("\tDatum: ", { bold: true }), R("________________", { color: C.soft })],
    { tabs: [{ type: TabStopType.LEFT, position: 6200 }], after: 400 }),
  zeile("Aufgabenart:", "Interpretieren von Quellen"),
  zeile("Thema:", "Die Grundlegung der modernen Welt in der Antike – Demokratie am Beispiel der attischen Polis"),
  P([R(D.titel, { font: "Cambria", bold: true, size: 32 })], { align: AlignmentType.CENTER, before: 300, after: 360 }),
  P([R("Aufgaben:", { bold: true })], { after: 120 }),
  ...AUFGABEN.map((a, i) => P([R((i + 1) + ".", { bold: true }), R("\t" + a)], { tabs: [{ type: TabStopType.LEFT, position: 480 }], indent: { left: 480, hanging: 480 }, after: 180, line: 300 })),
  P("", { after: 200 }),
  zeile("Material:", D.quelleTitel + " (Rückseite)"),
  zeile("Hilfsmittel:", "Übersicht „So interpretierst du eine Textquelle“"),
  zeile("Bearbeitungszeit:", "135 Minuten"),

  // Seite 2: Material
  new Paragraph({ suppressLineNumbers: true, children: [new PageBreak()] }),
  P([R(D.quelleTitel, { font: "Cambria", bold: true, size: 30 })], { after: 80 }),
  P(D.info, { size: 20, italics: true, color: C.muted, after: 60, indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } }),
  text,
  P(D.nachweis, { size: 16, color: C.muted, before: 160 }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 700, left: 1250, right: 992 } },
  lineNumbers: { countBy: 5, restart: LineNumberRestartFormat.CONTINUOUS, distance: 300 } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
