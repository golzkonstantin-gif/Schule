const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 26, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before ?? 0, after: o.after ?? 120, line: o.line } });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
function box(children, fill, border) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [new TableCell({ children, width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, borders: border || all({ style: BorderStyle.NONE, size: 0, color: "FFFFFF" }), margins: { top: 140, bottom: 140, left: 200, right: 200 }, verticalAlign: VerticalAlign.TOP })] })] });
}
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8 · Französische Revolution · Seite ", PageNumber.CURRENT] })] })] });

const ABSAETZE = [
  "Seit 1614 haben die Generalstände nicht mehr getagt, denn die Könige haben auch ohne sie regiert. Doch das Geld wird knapp. Kriege und Hofhaltung haben Frankreich hoch verschuldet. Unter Ludwig XVI. kommt der teure Krieg in Amerika dazu (1778 bis 1783). Schon die Zinsen für die Schulden verschlingen einen großen Teil der Staatseinnahmen.",
  "Der König braucht mehr Geld, doch die Privilegierten sind nicht bereit zu zahlen. 1787 scheitern Pläne, dass auch Adel und Klerus Steuern zahlen sollen. Als 1788 die Ernte schlecht ausfällt und das Brot teurer wird, droht der Staatsbankrott und Unruhe im Land. Im August 1788 gibt der König nach: Er kündigt die Generalstände für den Mai 1789 an.",
  "Im Frühjahr 1789 werden Abgeordnete gewählt. In jeder Gemeinde schreiben die Menschen ihre Klagen in Beschwerdehefte: über Steuern, Abgaben und Hunger. Am 5. Mai 1789 eröffnet der König die Generalstände in Versailles. Es kommen etwa 290 Vertreter des Klerus, 270 des Adels und 580 des Dritten Standes (Bauern und Bürger).",
];
const WOERTER = [["Generalstände", "Versammlung von Vertretern der drei Stände"], ["Privilegierte", "Menschen mit Vorrechten (Adel und Klerus)"], ["Staatsbankrott", "Der Staat kann seine Schulden nicht mehr bezahlen"], ["Beschwerdehefte", "Hefte, in die die Menschen ihre Klagen schrieben"]];

const c = [];
c.push(p(t("GESCHICHTE · KLASSE 8 · FRANZÖSISCHE REVOLUTION", { size: 18, bold: true, color: MUTED }), { after: 20 }));
c.push(p(t("Wie kam es zu den Generalständen?", { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 160 }));
ABSAETZE.forEach((a) => c.push(p(t(a, { font: "Cambria", size: 26 }), { after: 140, line: 340 })));
c.push(p(t(""), { after: 60 }));
c.push(box([p(t("Wörter erklärt", { bold: true, color: NAVY, size: 22 }), { after: 60 }), ...WOERTER.map(([w, e]) => p([t(w + ": ", { bold: true, size: 22 }), t(e, { size: 22 })], { after: 30 }))], LIGHT2));
c.push(p(t(""), { after: 60 }));
c.push(box([p(t("Deine Aufgabe", { bold: true, color: RED, size: 22 }), { after: 60 }), p(t("1  Lies den Text und unterstreiche alle Jahreszahlen.", { size: 22 }), { after: 30 }), p(t("2  Schreibe zu jeder Jahreszahl ein Stichwort an den Rand.", { size: 22 }), { after: 30 }), p(t("3  Überlege: Warum ruft der König die Generalstände ein, obwohl er sie 175 Jahre nicht gebraucht hat?", { size: 22 }), { after: 0 })], LIGHT, all(solid("B9C6E8", 6))));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 26 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer }, children: c }] });
  fs.writeFileSync("../Text_Wie_kam_es_zu_den_Generalstaenden.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
