// Textblatt „Nervig oder kafkaesk?“ (Aufgabe 1): halbes A4, zweimal pro Seite zum Schneiden; nur Text, Antworten ins Heft
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, HeightRule, VerticalAlign } = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cut = { style: BorderStyle.DASHED, size: 6, color: "9AA6C4" };

function haelfte() {
  return [
    p([t("Aufgabe 1: Nervig oder kafkaesk?", { font: "Cambria", size: 26, bold: true, color: NAVY }), t("   (allein, Antworten ins Heft, 3 Min.)", { size: 18, italics: true, color: MUTED })], { after: 40 }),
    p(t("Notiere im Heft: Nummer und N (nervig) oder K (kafkaesk). Bei K: Schreibe den Baustein dazu.", { size: 19, italics: true, color: MUTED }), { after: 80 }),
    new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [500, W - 500],
      rows: D.SZENARIEN.map(([, text], i) => new TableRow({ cantSplit: true, children: [
        new TableCell({ width: { size: 500, type: WidthType.DXA }, borders: all(none), shading: { fill: LIGHT, type: ShadingType.CLEAR, color: "auto" }, verticalAlign: VerticalAlign.TOP, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [p(t(String(i + 1), { bold: true, color: NAVY, size: 22 }), { after: 0 })] }),
        new TableCell({ width: { size: W - 500, type: WidthType.DXA }, borders: all(none), margins: { top: 50, bottom: 50, left: 140, right: 80 }, children: [p(t(text, { size: 23 }), { after: 0, line: 290 })] }),
      ] })),
    }),
  ];
}
const trenner = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ height: { value: 500, rule: HeightRule.EXACT }, children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: cut }, children: [p(t("✂", { size: 18, color: MUTED }), { after: 0 })] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 850, right: 850 } } }, children: [...haelfte(), trenner, p(t(""), { after: 200 }), ...haelfte()] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Szenarien_Nervig_oder_kafkaesk.docx", b));
