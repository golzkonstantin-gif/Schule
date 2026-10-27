const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const STAND = {
  K: { name: "ERSTER STAND · KLERUS", fill: "EFE8F8", edge: "8E6BBE",
       rolle: "Du bist Geistlicher. Du zahlst keine Steuern, und das soll so bleiben.",
       tipp: "Du darfst dich mit dem Zweiten Stand absprechen. Ihr sitzt am Tisch vorn." },
  A: { name: "ZWEITER STAND · ADEL", fill: "FDF1D8", edge: "E8A33D",
       rolle: "Du bist Adliger. Du zahlst keine Steuern und hast Vorrechte (z. B. Jagdrecht). Das soll so bleiben.",
       tipp: "Du darfst dich mit dem Ersten Stand absprechen. Ihr sitzt am Tisch vorn." },
  D: { name: "DRITTER STAND (BAUERN UND BÜRGER)", fill: "E3EEFB", edge: "3F7CC4",
       rolle: "Du zahlst Steuern und Abgaben. Der Brotpreis ist gestiegen, die Familie hat Hunger.",
       tipp: "Du willst, dass auch Adel und Klerus zahlen. Du bleibst an deinem Platz." },
};
const W = 10706, CW = W / 2;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, border: o.border });
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function card(key, reserve) {
  const s = STAND[key];
  const ch = [
    p(t(reserve ? "STANDESKARTE · RESERVE" : "STANDESKARTE · GEHEIM", { size: 15, bold: true, color: MUTED }), { after: 10 }),
    p(t(s.name, { font: "Cambria", size: 26, bold: true, color: NAVY }), { after: 50 }),
    p(t(s.rolle, { size: 22 }), { after: 40 }),
    p(t(s.tipp, { size: 22, bold: true }), { after: 70 }),
    p([t("STIMMZETTEL  ", { size: 18, bold: true, color: RED }), t("Antrag: Auch Adel und Klerus zahlen künftig Steuern.", { size: 20 })], { after: 40, border: { top: { style: BorderStyle.SINGLE, size: 8, color: s.edge, space: 4 } } }),
    p([t("□ JA", { size: 28, bold: true }), t("          □ NEIN", { size: 28, bold: true }), t("        (Kreuz nach deiner Rolle)", { size: 16, color: MUTED })], { after: 0 }),
  ];
  return new TableCell({ children: ch, width: { size: CW, type: WidthType.DXA }, shading: { fill: s.fill, type: ShadingType.CLEAR, color: "auto" }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 90, bottom: 60, left: 170, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
function page(list) {
  const rows = [];
  for (let i = 0; i < list.length; i += 2) rows.push(new TableRow({ cantSplit: true, height: { value: 3090, rule: HeightRule.EXACT }, children: [card(...list[i]), card(...list[i + 1])] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [CW, CW], rows });
}
const pg1 = [["K"], ["K"], ["K"], ["A"], ["A"], ["A"], ["A"], ["D"], ["D"], ["D"]];
const pg2 = Array(10).fill(["D"]);
const pg3 = [...Array(5).fill(["D"]), ...Array(5).fill(["D", true])];

// Absprache-Zettel (zweimal, A5-Hälften)
function zettel() {
  const line = () => p(t("_".repeat(62), { size: 24, color: "8A96B5" }), { before: 60, after: 100 });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, height: { value: 7000, rule: HeightRule.ATLEAST }, children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 140, bottom: 100, left: 240, right: 240 }, children: [
    p(t("ABSPRACHE · GEHEIM · NUR ERSTER UND ZWEITER STAND", { size: 16, bold: true, color: MUTED }), { after: 20 }),
    p(t("Absprache vor der Abstimmung", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 120 }),
    p([t("1  Wir stimmen beide mit  ", { size: 24 }), t("□ JA     □ NEIN", { size: 26, bold: true })], { after: 140 }),
    p(t("2  Das versprechen wir uns gegenseitig, damit keiner abspringt:", { size: 24 }), { after: 40 }), line(), line(),
    p(t("3  Das sagen wir dem Dritten Stand, damit er nachgibt:", { size: 24 }), { before: 80, after: 40 }), line(), line(),
    p(t("Flüsterstufe. Schreibt gemeinsam auf diesen Zettel. Die Lehrkraft sammelt ihn danach ein.", { size: 20, italics: true, color: MUTED }), { before: 80, after: 0 }),
  ] })] })] });
}
const sp = () => new Paragraph({ children: [new TextRun({ text: " ", size: 8 })], spacing: { after: 0 } });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 450, left: 600, right: 600 } } }, children: [
    page(pg1), new Paragraph({ children: [new PageBreak()] }), page(pg2), new Paragraph({ children: [new PageBreak()] }), page(pg3),
    new Paragraph({ children: [new PageBreak()] }), zettel(), sp(), zettel(),
  ] }] });
  fs.writeFileSync("../Druckvorlage_Standeskarten_Stimmzettel.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
