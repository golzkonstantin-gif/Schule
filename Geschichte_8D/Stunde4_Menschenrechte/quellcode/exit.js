const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, TabStopType, LeaderType, ImageRun, VerticalAlign } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const RW = 4700;
const lineTabs = [{ type: TabStopType.RIGHT, position: RW - 200, leader: LeaderType.UNDERSCORE }];
const p = (runs, o = {}) => new Paragraph({ children: runs, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, tabStops: o.tabs, alignment: o.align });
const line = (label) => p([t(label + " ", { bold: true, size: 18, color: NAVY }), t("\t")], { tabs: lineTabs, after: 130, before: 40 });
const more = () => p([t("\t")], { tabs: lineTabs, after: 130, before: 40 });
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const ic = new ImageRun({ type: "png", data: fs.readFileSync("../material/icon_waage.png"), transformation: { width: 30, height: 30 } });
function ticket() {
  const c = [];
  c.push(new Table({ width: { size: RW, type: WidthType.DXA }, columnWidths: [1560, 1560, 1580], rows: [new TableRow({ height: { value: 70, rule: "exact" }, children: [["1F4E9C", "1F4E9C"], ["FFFFFF", "B9C6E8"], ["D9534F", "D9534F"]].map(([f, l], i) => new TableCell({ width: { size: i === 2 ? 1580 : 1560, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: f, color: "auto" }, borders: { top: { style: BorderStyle.SINGLE, size: 4, color: l }, bottom: { style: BorderStyle.SINGLE, size: 4, color: l }, left: noB, right: noB }, children: [new Paragraph({ children: [] })] })) })] }));
  c.push(p([t("EXIT-TICKET · Geschichte 8D   ", { bold: true, size: 16, color: MUTED }), t("Name: ________________", { size: 16, color: MUTED }), t("   "), ic], { before: 50, after: 40 }));
  c.push(p([t("Ausgangsfrage. ", { bold: true, size: 19, color: NAVY }), t("Klerus und Adel sagen: „Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?“ Du bist der Dritte Stand (Bauern und Bürger). Antworte:", { size: 18 })], { after: 70 }));
  c.push(line("Wir wollen Menschenrechte, damit wir endlich …")); c.push(more()); c.push(more());
  c.push(p([t("Das steht heute auch im Grundgesetz. Kreuze an, was zu deinem Satz passt:", { bold: true, size: 17, color: NAVY })], { before: 60, after: 50 }));
  c.push(p([t("□ Art. 3: Gleichheit vor dem Gesetz     □ Art. 5: Meinungsfreiheit", { size: 17 })], { after: 40 }));
  c.push(p([t("□ Art. 14: Eigentum     □ Art. 20: Die Macht geht vom Volk aus", { size: 17 })], { after: 60 }));
  c.push(p([t("Das war mir unklar (freiwillig): ", { size: 16, color: MUTED, italics: true }), t("\t")], { tabs: lineTabs, after: 0, before: 40 }));
  return c;
}
const cell = () => new TableCell({ children: ticket(), width: { size: 5300, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 80, bottom: 60, left: 200, right: 200 }, verticalAlign: VerticalAlign.TOP });
(async () => {
  const tab = new Table({ width: { size: 10600, type: WidthType.DXA }, columnWidths: [5300, 5300], rows: [0, 1, 2].map(() => new TableRow({ cantSplit: true, height: { value: 4950, rule: "exact" }, children: [cell(), cell()] })) });
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 400, left: 650, right: 650 } } }, children: [tab] }] });
  fs.writeFileSync("../Exit-Ticket_Menschenrechte_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
