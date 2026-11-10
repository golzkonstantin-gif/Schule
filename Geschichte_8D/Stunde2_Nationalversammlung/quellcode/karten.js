const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, PageBreak, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const STAND = {
  K: { name: "ERSTER STAND · KLERUS", edge: "8E6BBE", ink: "6B4A9E", rolle: "Du bist Geistlicher. Du zahlst keine Steuern und hast Vorrechte. Du willst sie behalten.", tipp: "Du stimmst nach deiner Rolle." },
  A: { name: "ZWEITER STAND · ADEL", edge: "E8A33D", ink: "B9791A", rolle: "Du bist Adliger. Du zahlst keine Steuern und hast Vorrechte (z. B. Jagdrecht, Ämter am Hof). Du willst sie behalten.", tipp: "Du stimmst nach deiner Rolle." },
  D: { name: "DRITTER STAND (BAUERN UND BÜRGER)", edge: "3F7CC4", ink: "2F6AAE", rolle: "Du zahlst Steuern und Abgaben. Der Brotpreis ist gestiegen, die Familie hat Hunger. Du willst, dass sich etwas ändert.", tipp: "Du stimmst nach deiner Überzeugung." },
};
const W = 10706, CW = W / 2;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, border: o.border });
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const thin = { style: BorderStyle.SINGLE, size: 6, color: "8A96B5" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

// ---- Standeskarten (12 pro A4, ohne Stimmzettel) ----
function card(key, reserve) {
  const s = STAND[key];
  return new TableCell({ children: [
    p(t(reserve ? "STANDESKARTE · RESERVE" : "STANDESKARTE · GEHEIM", { size: 15, bold: true, color: MUTED }), { after: 10 }),
    p(t(s.name, { font: "Cambria", size: 26, bold: true, color: s.ink }), { after: 50 }),
    p(t(s.rolle, { size: 22 }), { after: 40 }),
    p(t(s.tipp, { size: 22, bold: true }), { after: 0 }),
  ], width: { size: CW, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: { style: BorderStyle.SINGLE, size: 48, color: s.edge }, right: dash }, margins: { top: 90, bottom: 60, left: 200, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
function page(list) {
  const rows = [];
  for (let i = 0; i < list.length; i += 2) rows.push(new TableRow({ cantSplit: true, height: { value: 2220, rule: HeightRule.EXACT }, children: [card(...list[i]), card(...list[i + 1])] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [CW, CW], rows });
}
const pg1 = [["K"], ["K"], ["K"], ["A"], ["A"], ["A"], ["A"], ["D"], ["D"], ["D"], ["D"], ["D"], ["D"], ["D"]];
const pg2 = [...Array(11).fill(["D"]), ...Array(3).fill(["D", true])];

// ---- Stimmzettel (Schnipsel, acht pro A4), werden abgegeben ----
function stimm() {
  return [
    p([t("STIMMZETTEL · 8D", { size: 15, bold: true, color: MUTED })], { after: 20 }),
    p(t("Abstimmung der Stände", { font: "Cambria", size: 24, bold: true, color: NAVY }), { after: 40 }),
    p([t("Mein Stand:  ", { size: 20, bold: true }), t("□ Klerus   □ Adel   □ Dritter Stand", { size: 19 })], { after: 70 }),
    p([t("Antrag A", { size: 21, bold: true, color: RED })], { after: 10 }),
    p(t("Auch Adel und Klerus zahlen künftig Steuern.", { size: 20 }), { after: 20 }),
    p(t("□ Ja            □ Nein", { size: 22, bold: true }), { after: 70 }),
    p([t("Antrag B", { size: 21, bold: true, color: RED })], { after: 10 }),
    p(t("Die Stände werden abgeschafft. Alle haben gleiche Rechte.", { size: 20 }), { after: 20 }),
    p(t("□ Ja            □ Nein", { size: 22, bold: true }), { after: 0 }),
  ];
}
const sc = () => new TableCell({ children: stimm(), width: { size: CW, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 90, bottom: 60, left: 200, right: 170 }, verticalAlign: VerticalAlign.TOP });
const bogenPage = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [CW, CW], rows: [0, 1, 2, 3].map(() => new TableRow({ cantSplit: true, height: { value: 3700, rule: HeightRule.EXACT }, children: [sc(), sc()] })) });

const brk = () => new Paragraph({ children: [new PageBreak()] });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 450, left: 600, right: 600 } } }, children: [
    page(pg1), brk(), page(pg2), brk(), bogenPage,
  ] }] });
  fs.writeFileSync("../Druckvorlage_Karten_Stimmzettel_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
