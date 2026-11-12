// Inhaltliche Zusammenfassung der Stunde vom 11.11. (Probeklausur bis zur Position des Verfassers) mit Ausblick auf Urteil und Klausur. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, AlignmentType } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9922, S = 22;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const P = (t, o = {}) => new Paragraph({ keepNext: o.keepNext, alignment: o.align, spacing: { after: o.after ?? 90, before: o.before ?? 0, line: o.line ?? 290 }, indent: o.indent, border: o.border,
  children: Array.isArray(t) ? t : runs(t, o) });
const H = (n, t) => P([R(n + "  ", { font: "Cambria", bold: true, size: 30, color: C.terra }), R(t, { font: "Cambria", bold: true, size: 28 })], { before: 260, after: 100, keepNext: true });
const li = t => P("–  " + t, { indent: { left: 300, hanging: 300 }, after: 50 });
const kasten = t => P(t, { italics: true, line: 310, align: AlignmentType.JUSTIFIED, indent: { left: 160 }, after: 120, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.terra, space: 8 } } });
function tab(cw, head, rows, size = 20) {
  const cell = (t, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 110, right: 110 },
    borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children: [P(t, { size, bold: o.bold, after: 0, line: 260 })] });
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [
    new TableRow({ tableHeader: true, children: head.map((t, i) => cell(t, i, { head: true, bold: true })) }),
    ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, i === 0 ? { bold: true } : {})) }))] });
}

const children = [
  P("GESCHICHTE · KLASSE 11 · ZUSAMMENFASSUNG DER STUNDE VOM 11.11.", { size: 16, bold: true, color: C.terra, after: 30 }),
  P([R("Probeklausur: Pseudo-Xenophon über die Demokratie", { font: "Cambria", bold: true, size: 36 })], { after: 40 }),
  P([R("Was wir erarbeitet haben – bis zur Position des Verfassers", { size: 22, italics: true, color: C.muted })], { after: 160 }),

  H("1", "Worum es ging"),
  P("Wir haben an einer Probeklausur geübt, wie man eine Textquelle interpretiert – nach den Arbeitsschritten der Übersicht „So interpretierst du eine Textquelle“. Die Quelle ist ein Auszug aus „Über den Staat der Athener“ von einem unbekannten Verfasser („Pseudo-Xenophon“), geschrieben um 430–424 v. Chr. Er beschreibt die attische Demokratie aus der Sicht eines Gegners."),
  P("**Themenfrage:** " + D.themenfrage),

  H("2", "Einleitung: die formale Analyse (TATTEA)"),
  tab([2000, 7922], ["TATTEA", "Angabe"], D.tattea.map(([a, b]) => [a, b])),
  P("**So kann die Einleitung lauten:**", { before: 140, after: 60 }),
  kasten("„" + D.einleitung + "“"),
  P("**Achtung:** Der Autor ist unbekannt – die Schrift wurde Xenophon nur irrtümlich zugeschrieben. Er schreibt um 430 v. Chr., also zur Zeit des Perikles, nicht zur Zeit des Kleisthenes.", { size: 20 }),

  H("3", "Die Position des Verfassers"),
  P("Wir haben den Text in Sinnabschnitte gegliedert und zu jedem Abschnitt die Haltung des Verfassers herausgearbeitet:", { after: 80 }),
  tab([1100, 2700, 6122], ["Zeilen", "Aspekt", "Position des Verfassers"], D.position.map(([a, b, z]) => [z, a, b])),
  P("**Gesamtposition:** " + D.positionGesamt, { before: 140 }),
  P("**Textsprache:** " + D.sprache.replace(/^Sprache /, "")),
  P("**So kann die Position ausformuliert werden:**", { before: 100, after: 60, keepNext: true }),
  kasten(D.muster2),

  H("4", "Ausblick"),
  li("**Nächste Stunde:** Urteilsbildung – Haben die Reformen die Vormachtstellung des Adels beendet? Dafür und dagegen abwägen, Fazit, Werturteil mit Maßstab."),
  li("**Klausur am 25.11. (90 Minuten):** Einleitung mit Themenfrage und TATTEA, Position des Verfassers herausarbeiten, Urteil. Eine eigene Aufgabe zur Beschreibung der Reformen des Kleisthenes gibt es nicht – das Wissen über die Reihe brauchst du aber als Beleg im Urteil."),
];
const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 992, right: 992 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
