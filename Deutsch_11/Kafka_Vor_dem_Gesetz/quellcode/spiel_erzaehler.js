// Spielvorlage mit Erzähler: Der Erzähler vermittelt das Geschehen, der Türhüter spricht, der Mann sagt fast nichts (3 Exemplare drucken)
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign } = require("docx");
const D = require("./daten_macht");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", BLUE = "3F7CC4", T_FILL = "FBD9D7", M_FILL = "D6E6F8", E_FILL = "EEEEEE";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const farbe = { "Erzähler": [E_FILL, NAVY], "Türhüter": [T_FILL, RED], "Mann": [M_FILL, BLUE] };
const cW = [1500, W - 1500];
const rows = D.SPIEL_ERZAEHLER.map(([rolle, text, regie]) => {
  const [fill, col] = farbe[rolle];
  const erz = rolle === "Erzähler";
  const inhalt = [];
  if (regie) inhalt.push(t("(" + regie + ") ", { italics: true, color: MUTED, size: 19 }));
  inhalt.push(t(erz ? text : "„" + text + "“", { size: erz ? 21 : 23, bold: rolle === "Mann" }));
  return new TableRow({ cantSplit: true, children: [
    new TableCell({ width: { size: cW[0], type: WidthType.DXA }, borders: all(none), shading: { fill, type: ShadingType.CLEAR, color: "auto" }, verticalAlign: VerticalAlign.TOP, margins: { top: 60, bottom: 60, left: 120, right: 80 }, children: [p(t(rolle, { bold: true, color: col, size: 21 }), { after: 0 })] }),
    new TableCell({ width: { size: cW[1], type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5DCEC" } }, margins: { top: 60, bottom: 60, left: 160, right: 80 }, children: [p(inhalt, { after: 0, line: 280 })] }),
  ] });
});
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 560, left: 850, right: 850 } } }, children: [
    p(t("Spielvorlage mit Erzähler: „Vor dem Gesetz“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 60 }),
    p(t(D.SPIEL_ERZAEHLER_REGIE, { size: 20, color: MUTED }), { after: 30 }),
    p(t("In Klammern = was ihr tut. Text nach Kafka, gekürzt.", { size: 18, italics: true, color: MUTED }), { after: 120 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cW, rows }),
  ] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Macht_Ohnmacht/Spielvorlage_mit_Erzaehler.docx", b));
