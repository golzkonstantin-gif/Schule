// Spielvorlage „Vor dem Gesetz“: nur die Sätze der beiden Figuren in der Reihenfolge des Textes (2 Exemplare drucken)
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign } = require("docx");
const D = require("./daten_macht");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", BLUE = "3F7CC4", T_FILL = "FBD9D7", M_FILL = "D6E6F8";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 23, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cW = [1700, W - 1700];
const rows = D.SPIEL.map(([fig, regie, satz, umg]) => {
  const T = fig === "Türhüter";
  const inhalt = [];
  if (regie) inhalt.push(t("(" + regie + ") ", { italics: true, color: MUTED, size: 21 }));
  if (satz) inhalt.push(t("„" + satz + "“", { italics: umg, color: umg ? "444444" : "000000" }));
  return new TableRow({ cantSplit: true, children: [
    new TableCell({ width: { size: cW[0], type: WidthType.DXA }, borders: all(none), shading: { fill: T ? T_FILL : M_FILL, type: ShadingType.CLEAR, color: "auto" }, verticalAlign: VerticalAlign.TOP, margins: { top: 70, bottom: 70, left: 120, right: 80 }, children: [p(t(fig, { bold: true, color: T ? RED : BLUE, size: 22 }), { after: 0 })] }),
    new TableCell({ width: { size: cW[1], type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5DCEC" } }, margins: { top: 70, bottom: 70, left: 160, right: 80 }, children: [p(inhalt, { after: 0, line: 290 })] }),
  ] });
});
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 23 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 760, bottom: 620, left: 850, right: 850 } } }, children: [
    p(t("Spielvorlage: „Vor dem Gesetz“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 60 }),
    p(t(D.SPIEL_REGIE, { size: 21, color: MUTED }), { after: 30 }),
    p(t("Kursiv in Grau = aus der indirekten Rede umgeformt. In Klammern = was ihr tut.", { size: 19, italics: true, color: MUTED }), { after: 140 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cW, rows }),
  ] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Macht_Ohnmacht/Spielvorlage_Vor_dem_Gesetz.docx", b));
