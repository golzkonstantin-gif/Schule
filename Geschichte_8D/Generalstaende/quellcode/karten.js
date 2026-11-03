const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const STAND = {
  K: { name: "ERSTER STAND · KLERUS", fill: "FFFFFF", edge: "8E6BBE", ink: "6B4A9E",
       rolle: "Du bist Geistlicher. Du zahlst keine Steuern und hast Vorrechte. Du willst sie behalten.",
       tipp: "Du stimmst nach deiner Rolle." },
  A: { name: "ZWEITER STAND · ADEL", fill: "FFFFFF", edge: "E8A33D", ink: "B9791A",
       rolle: "Du bist Adliger. Du zahlst keine Steuern und hast Vorrechte (z. B. Jagdrecht, Ämter am Hof). Du willst sie behalten.",
       tipp: "Du stimmst nach deiner Rolle." },
  D: { name: "DRITTER STAND (BAUERN UND BÜRGER)", fill: "FFFFFF", edge: "3F7CC4", ink: "2F6AAE",
       rolle: "Du zahlst Steuern und Abgaben. Der Brotpreis ist gestiegen, die Familie hat Hunger. Du willst, dass sich etwas ändert.",
       tipp: "Du stimmst nach deiner Überzeugung." },
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
    p(t(s.name, { font: "Cambria", size: 26, bold: true, color: s.ink }), { after: 50 }),
    p(t(s.rolle, { size: 22 }), { after: 40 }),
    p(t(s.tipp, { size: 22, bold: true }), { after: 70 }),
    p([t("STIMMZETTEL  ", { size: 18, bold: true, color: RED }), t("Kreuze bei jedem Antrag an.", { size: 17, color: MUTED })], { after: 30, border: { top: { style: BorderStyle.SINGLE, size: 8, color: s.edge, space: 4 } } }),
    ...[["A", "Steuern auch für Adel und Klerus"], ["B", "Stände abschaffen, gleiche Rechte"], ["C", "Mehr Mitsprache für Bauern"]].map(([l, x]) => p([t(l + "  ", { size: 20, bold: true }), t(x, { size: 18 }), t("      □ Ja  □ Nein", { size: 20, bold: true })], { after: 20 })),
  ];
  return new TableCell({ children: ch, width: { size: CW, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: { style: BorderStyle.SINGLE, size: 48, color: s.edge }, right: dash }, margins: { top: 90, bottom: 60, left: 200, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
function page(list) {
  const rows = [];
  for (let i = 0; i < list.length; i += 2) rows.push(new TableRow({ cantSplit: true, height: { value: 3090, rule: HeightRule.EXACT }, children: [card(...list[i]), card(...list[i + 1])] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [CW, CW], rows });
}
const pg1 = [["K"], ["K"], ["K"], ["A"], ["A"], ["A"], ["A"], ["D"], ["D"], ["D"]];
const pg2 = Array(10).fill(["D"]);
const pg3 = [...Array(5).fill(["D"]), ...Array(5).fill(["D", true])];

const sp = () => new Paragraph({ children: [new TextRun({ text: " ", size: 8 })], spacing: { after: 0 } });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 450, left: 600, right: 600 } } }, children: [
    page(pg1), new Paragraph({ children: [new PageBreak()] }), page(pg2), new Paragraph({ children: [new PageBreak()] }), page(pg3),
  ] }] });
  fs.writeFileSync("../Druckvorlage_Standeskarten_Stimmzettel.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
