const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line } });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function box(children, fill, border) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [new TableCell({ children, width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, borders: border || all(noB), margins: { top: 90, bottom: 90, left: 160, right: 160 }, verticalAlign: VerticalAlign.TOP })] })] });
}
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8 · Französische Revolution · Seite ", PageNumber.CURRENT] })] })] });

const ABSAETZE = [
  "Seit 1614 haben die Generalstände nicht mehr getagt, denn die Könige haben auch ohne sie regiert. Doch das Geld wird knapp. Kriege und Hofhaltung haben Frankreich hoch verschuldet. Unter Ludwig XVI. kommt der teure Krieg in Amerika dazu (1778 bis 1783). Schon die Zinsen für die Schulden verschlingen einen großen Teil der Staatseinnahmen.",
  "Der König braucht mehr Geld, doch die Privilegierten sind nicht bereit zu zahlen. 1787 scheitern Pläne, dass auch Adel und Klerus Steuern zahlen sollen. Als 1788 die Ernte schlecht ausfällt und das Brot teurer wird, droht der Staatsbankrott und Unruhe im Land. Im August 1788 gibt der König nach: Er kündigt die Generalstände für den Mai 1789 an.",
  "Im Frühjahr 1789 werden Abgeordnete gewählt. In jeder Gemeinde schreiben die Menschen ihre Klagen in Beschwerdehefte: über Steuern, Abgaben und Hunger. Am 5. Mai 1789 eröffnet der König die Generalstände in Versailles. Es kommen etwa 290 Vertreter des Klerus, 270 des Adels und 580 des Dritten Standes (Bauern und Bürger).",
];
const WOERTER = [["Generalstände", "Versammlung von Vertretern der drei Stände"], ["Privilegierte", "Menschen mit Vorrechten (Adel und Klerus)"], ["Staatsbankrott", "Der Staat kann seine Schulden nicht mehr bezahlen"], ["Beschwerdehefte", "Hefte, in die die Menschen ihre Klagen schrieben"]];
const EREIGNISSE = [
  "Die Generalstände tagen zum letzten Mal. Danach regieren die Könige 175 Jahre ohne sie.",
  "Der Krieg in Amerika kostet viel Geld. Frankreich macht hohe Schulden.",
  "Der Plan, dass auch Adel und Klerus Steuern zahlen, scheitert.",
  "Missernte, das Brot wird teurer, der Staatsbankrott droht.",
  "Der König kündigt die Generalstände an.",
  "Die Abgeordneten werden gewählt. Die Menschen schreiben ihre Beschwerdehefte.",
  "Der König eröffnet die Generalstände in Versailles.",
];

const c = [];
c.push(p(t("GESCHICHTE · KLASSE 8 · FRANZÖSISCHE REVOLUTION", { size: 16, bold: true, color: MUTED }), { after: 10 }));
c.push(p(t("Wie kam es zu den Generalständen?", { font: "Cambria", size: 34, bold: true, color: NAVY }), { after: 80 }));
ABSAETZE.forEach((a) => c.push(p(t(a, { font: "Cambria", size: 24 }), { after: 100, line: 300, align: AlignmentType.JUSTIFIED })));

// Wörter erklärt (zwei Spalten, kompakt)
c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W / 2, W / 2], rows: [0, 2].map((i) => new TableRow({ cantSplit: true, children: [WOERTER[i], WOERTER[i + 1]].map((x) => new TableCell({ width: { size: W / 2, type: WidthType.DXA }, shading: { fill: LIGHT2, type: ShadingType.CLEAR, color: "auto" }, borders: all(noB), margins: { top: 40, bottom: 40, left: 120, right: 120 }, children: [p([t(x[0] + ": ", { bold: true, size: 18 }), t(x[1], { size: 18 })], { after: 0 })] })) })) }));
c.push(p(t(""), { after: 40 }));
c.push(box([p([t("Deine Aufgabe  ", { bold: true, color: RED, size: 20 }), t("1  Lies den Text und unterstreiche alle Jahreszahlen.   2  Trage die Daten unten im Zeitstrahl ein.   3  Überlege: Warum ruft der König die Generalstände ein, obwohl er sie 175 Jahre nicht gebraucht hat?", { size: 20 })], { after: 0 })], "FFFFFF", all(solid("B9C6E8", 6))));

// Zeitstrahl (unten auf demselben Blatt)
c.push(p(t("Zeitstrahl: Der Weg zu den Generalständen", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 140, after: 40 }));
const ZW = [2000, 500, 7706];
c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ZW, rows: EREIGNISSE.map((e, i) => new TableRow({ cantSplit: true, height: { value: 680, rule: "atLeast" }, children: [
  new TableCell({ width: { size: ZW[0], type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM, borders: { top: noB, left: noB, right: noB, bottom: solid("8A96B5", 8) }, margins: { top: 20, bottom: 20, left: 80, right: 120 }, children: [p(t("", { size: 22 }), { after: 0 })] }),
  new TableCell({ width: { size: ZW[1], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: all(noB), margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [p(t(i === 0 ? "▼" : i === EREIGNISSE.length - 1 ? "▲" : "│", { size: 22, color: NAVY, bold: true }), { after: 0, align: AlignmentType.CENTER })] }),
  new TableCell({ width: { size: ZW[2], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: { fill: i % 2 ? LIGHT : LIGHT2, type: ShadingType.CLEAR, color: "auto" }, borders: all(solid("B9C6E8", 6)), margins: { top: 30, bottom: 30, left: 140, right: 140 }, children: [p(t(e, { size: 22 }), { after: 0 })] }),
] })) }));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 650, bottom: 600, left: 850, right: 850 } } }, footers: { default: footer }, children: c }] });
  fs.writeFileSync("../Text_Wie_kam_es_zu_den_Generalstaenden.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
