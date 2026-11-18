// Übungsklausur im Klausurformat (A4, ein Blatt doppelseitig): Seite 1 Aufgabenstellung, Seite 2 Material M1 mit Zeilennummern (Zeilen wie im Original). Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, BorderStyle, AlignmentType, LineNumberRestartFormat, TabStopType, PageBreak } = require("docx");
const { C } = require("../../Probeklausur_PseudoXenophon/quellcode/lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 23, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ suppressLineNumbers: true, alignment: o.align, tabStops: o.tabs, keepNext: o.keepNext,
  spacing: { after: o.after ?? 0, before: o.before ?? 0, line: o.line }, indent: o.indent, border: o.border, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const TAB = [{ type: TabStopType.LEFT, position: 3000 }];
const zeile = (k, v) => P([R(k, { bold: true }), R("\t" + v)], { tabs: TAB, after: 160, indent: { left: 3000, hanging: 3000 } });
// Quelle: Absätze wie im Original, feste Zeilenumbrüche
const absaetze = [D.zeilen.slice(0, 8), D.zeilen.slice(8)];
const text = absaetze.map((z, k) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 0, before: k ? 0 : 120, line: 360 },
  children: z.map((t, i) => new TextRun({ text: t, font: "Calibri", size: 25, color: C.dark, break: i ? 1 : 0 })) }));

const children = [
  P([R("Klausur im Grundkurs Geschichte", { font: "Cambria", bold: true, size: 34 }), R("\tJahrgang 11 · Übungsklausur", { size: 20, italics: true, color: C.muted })],
    { tabs: [{ type: TabStopType.RIGHT, position: 9664 }], after: 300, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.dark, space: 6 } } }),
  P([R("Name: ", { bold: true }), R("______________________________", { color: C.soft }), R("\tDatum: ", { bold: true }), R("________________", { color: C.soft })],
    { tabs: [{ type: TabStopType.LEFT, position: 6200 }], after: 400 }),
  zeile("Aufgabenart:", "Interpretation einer historischen Quelle"),
  zeile("Hilfsmittel:", "Nachschlagewerk zur Rechtschreibung der deutschen Sprache"),
  zeile("Bearbeitungszeit:", "90 Minuten"),
  P([R("Aufgabe:", { bold: true })], { before: 300, after: 200 }),
  P([R(D.themenfrage, { font: "Cambria", bold: true, size: 32 })], { align: AlignmentType.CENTER, after: 300 }),
  P("Beurteilen Sie unter der oben genannten Themenfrage die Position des Verfassers, indem Sie …", { after: 160, line: 300 }),
  ...D.aufgabe.map(a => P([R("–\t"), R(a)], { tabs: [{ type: TabStopType.LEFT, position: 480 }], indent: { left: 480, hanging: 480 }, after: 140, line: 300 })),
  P("Viel Erfolg!", { align: AlignmentType.CENTER, before: 500, italics: true }),
  new Paragraph({ suppressLineNumbers: true, children: [new PageBreak()] }),
  P([R(D.quelleTitel, { font: "Cambria", bold: true, size: 30 })], { after: 80 }),
  P(D.info, { size: 20, italics: true, color: C.muted, after: 60, line: 280, indent: { left: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: C.terra, space: 6 } } }),
  ...text,
  P(D.nachweis, { size: 18, color: C.muted, before: 160 }),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 700, left: 1250, right: 992 } },
  lineNumbers: { countBy: 5, restart: LineNumberRestartFormat.CONTINUOUS, distance: 300 } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
