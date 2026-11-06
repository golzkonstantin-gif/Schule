const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 50 }, border: o.line ? { bottom: { style: BorderStyle.SINGLE, size: 4, color: "8A96B5", space: 1 } } : undefined });
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const W = 10706;
const bul = (xs) => xs.map((x) => p([t("• ", { size: 19 }), t(x, { size: 19 })], { after: 20 }));

function schnipsel() {
  const hilfen = new Table({ width: { size: W - 340, type: WidthType.DXA }, columnWidths: [(W - 340) / 2, (W - 340) / 2], rows: [new TableRow({ children: [
    new TableCell({ width: { size: (W - 340) / 2, type: WidthType.DXA }, borders: all(none), margins: { top: 0, bottom: 0, left: 0, right: 60 }, children: [
      p(t("Meine Position", { font: "Cambria", size: 20, bold: true, color: NAVY }), { after: 20 }),
      ...bul(["Ich rate dir, dich anzuschließen, weil …", "Ich rate dir, noch abzuwarten, weil …"]),
      p(t("Zwischenlösung", { font: "Cambria", size: 20, bold: true, color: NAVY }), { before: 30, after: 20 }),
      ...bul(["Ich rate dir, …, aber gleichzeitig …"]),
    ] }),
    new TableCell({ width: { size: (W - 340) / 2, type: WidthType.DXA }, borders: all(none), margins: { top: 0, bottom: 0, left: 60, right: 0 }, children: [
      p(t("Mein Argument aus der Stunde", { font: "Cambria", size: 20, bold: true, color: NAVY }), { after: 20 }),
      ...bul(["Der Dritte Stand (Bauern und Bürger) …", "Die Bauern zahlen …, Klerus und Adel …", "Mit einer Verfassung …", "Wer abseits steht, …", "Ein Gegenargument, das ich ernst nehme: …"]),
    ] }),
  ] })] });
  const ch = [
    p([t("ABSCHLUSSFRAGE · GESCHICHTE 8C   ", { size: 15, bold: true, color: MUTED }), t("Erkläre es deinem Freund", { font: "Cambria", size: 22, bold: true, color: NAVY })], { after: 40 }),
    p(t("Dein Freund ist auch Bauer: „Die Abgeordneten des Dritten Standes (Bauern und Bürger) wollen eine Verfassung. Soll ich mich ihnen anschließen? Ich habe doch genug mit der Ernte zu tun.“", { size: 20, italics: true }), { after: 50 }),
    hilfen,
    p(t(" ", { size: 6 }), { after: 30 }),
    ...[0, 1, 2, 3, 4].map(() => p(t("_".repeat(86), { size: 22, color: "8A96B5" }), { before: 60, after: 40 })),
  ];
  return new TableCell({ children: ch, width: { size: W, type: WidthType.DXA }, borders: all(dash), margins: { top: 90, bottom: 60, left: 170, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
const rows = [0, 1, 2].map(() => new TableRow({ cantSplit: true, height: { value: 5100, rule: HeightRule.EXACT }, children: [schnipsel()] }));
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 450, left: 600, right: 600 } } }, children: [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows })] }] });
  fs.writeFileSync("../Schnipsel_Abschlussfrage_Rat_an_den_Freund.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
