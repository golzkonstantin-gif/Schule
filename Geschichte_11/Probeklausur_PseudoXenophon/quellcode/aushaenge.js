// Lösungsaushänge mit Formulierungsvorschlägen (A4 hoch, je Abschnitt eine neue Seite, große Schrift zum Aushängen). Tintensparend.
// 1 Themenfrage + Einleitung (Tafelrückseite) · 2 Position des Verfassers (Stichpunkte, für das Zusammentragen) · 3 Organe und Gliederung · 4 Urteil
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, PageBreak } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9922, S = 26;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const P = (t, o = {}) => new Paragraph({ spacing: { after: o.after ?? 80, before: o.before ?? 0, line: o.line }, indent: o.indent, border: o.border, children: Array.isArray(t) ? t : runs(t, o) });
const kopf = (nr, t, sub) => [P("AUSHANG " + nr + " · PROBEKLAUSUR „" + D.titel.toUpperCase() + "“", { size: 16, bold: true, color: C.terra, after: 30 }),
  P([R(t, { font: "Cambria", bold: true, size: 40 })], { after: 40 }), ...(sub ? [P(sub, { size: 20, italics: true, color: C.muted, after: 160 })] : [])];
const h = t => P([R(t, { font: "Cambria", bold: true, size: 30, color: C.terra })], { before: 200, after: 100 });
const vorschlag = t => P(t, { size: S, italics: true, line: 320, indent: { left: 160 }, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.terra, space: 8 } } });
const li = (t, o = {}) => P("–  " + t, { size: o.size || S, after: 60, indent: { left: 300, hanging: 300 } });
function tab(cw, head, rows, size = 23) {
  const cell = (t, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 90, bottom: 90, left: 120, right: 120 },
    borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children: [P(t, { size, bold: o.bold, color: o.color, after: 0 })] });
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [
    new TableRow({ tableHeader: true, children: head.map((t, i) => cell(t, i, { head: true, bold: true })) }),
    ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, i === 0 ? { bold: true } : {})) }))] });
}
const neu = () => new Paragraph({ children: [new PageBreak()] });

const children = [
  // 1 Einleitung
  ...kopf(1, "Themenfrage und Einleitung", "Aufgabe 1 · Formulierungsvorschlag (Rückseite der Tafel)"),
  h("Themenfrage"),
  P("**" + D.themenfrage + "**", { size: 28 }),
  h("Formale Analyse (TATTEA)"),
  tab([2300, 7622], ["", "Angabe"], D.tattea.map(([a, b]) => [a, b])),
  h("Formulierungsvorschlag"),
  vorschlag("„" + D.einleitung + "“"),

  // 2 Position
  neu(),
  ...kopf(2, "Position des Verfassers", "Aufgabe 2 · Zusammenfassung in Stichpunkten zum Zusammentragen"),
  tab([2600, 5922, 1400], ["Aspekt", "Position des Verfassers", "Zeilen"], D.position.map(([a, b, z]) => [a, b, z]), 22),
  h("Gesamtposition"),
  vorschlag(D.positionGesamt),
  P("**Textsprache:** " + D.sprache, { size: 22, before: 120 }),

  // 3 Organe und Gliederung
  neu(),
  ...kopf(3, "Verfassung und Gesellschaft unter Kleisthenes", "Aufgabe 3 · Beschreiben"),
  h("Die wichtigsten Verfassungsorgane"),
  tab([3000, 6922], ["Organ", "Zusammensetzung und Aufgaben"], D.organe, 22),
  h("Gesellschaftliche Gliederung und Rechte"),
  tab([3400, 6522], ["Gruppe", "Rechte"], D.gliederung, 22),
  h("Formulierungsvorschlag"),
  vorschlag(D.kontextText),

  // 4 Urteil
  neu(),
  ...kopf(4, "Urteil: Entmachtung des Adels?", "Aufgabe 4 · Beurteilen (Sachurteil), dazu ein mögliches Werturteil"),
  h("Dafür spricht"),
  ...D.urteilPro.map(t => li(t, { size: 23 })),
  h("Dagegen spricht"),
  ...D.urteilContra.map(t => li(t, { size: 23 })),
  h("Formulierungsvorschlag Sachurteil"),
  vorschlag("„" + D.sachurteil + "“"),
  h("Mögliches Werturteil"),
  vorschlag("„" + D.werturteil + "“"),
  P("Andere Urteile sind richtig, wenn sie eine klare Position beziehen und mit Belegen aus Quelle und Kontext begründet sind.", { size: 20, italics: true, color: C.muted, before: 120 }),
];
const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 992, right: 992 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
