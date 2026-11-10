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

// ---- Beobachtungszettel (Schnipsel, vier pro A4): nur Ja-Stimmen zählen, Ergebnis selbst ausrechnen ----
const cc = (txt, o = {}) => new TableCell({ children: [p(t(txt, { size: o.size || 20, bold: o.bold, color: o.color }), { after: 0, align: o.align || AlignmentType.CENTER })], width: { size: o.w, type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 50, bottom: 50, left: 70, right: 70 }, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span });
function zettel() {
  const M = [800, 1250, 1250, 1500, 1200, 1350, 1200, 1650];
  const jn = (w, von) => cc("Ja ___ von " + von, { w, size: 18, color: "5B6B8C" });
  const matrix = new Table({ width: { size: M.reduce((x, y) => x + y, 0), type: WidthType.DXA }, columnWidths: M, rows: [
    new TableRow({ children: [cc("", { w: M[0] }), cc("Nach Ständen (jeder Stand eine Stimme)", { w: M[1] + M[2] + M[3] + M[4] + M[5], span: 5, bold: true, size: 17 }), cc("Nach Köpfen (jede Hand)", { w: M[6] + M[7], span: 2, bold: true, size: 17 })] }),
    new TableRow({ children: [cc("Antrag", { w: M[0], bold: true, size: 16 }), cc("Klerus", { w: M[1], bold: true, size: 17, color: "6B4A9E" }), cc("Adel", { w: M[2], bold: true, size: 17, color: "B9791A" }), cc("Dritter Stand", { w: M[3], bold: true, size: 17, color: "2F6AAE" }), cc("Stände mit Ja", { w: M[4], bold: true, size: 16 }), cc("Ergebnis", { w: M[5], bold: true, size: 17 }), cc("alle Ja", { w: M[6], bold: true, size: 16 }), cc("Ergebnis", { w: M[7], bold: true, size: 17 })] }),
    ...["A", "B"].map((x) => new TableRow({ height: { value: 640, rule: HeightRule.ATLEAST }, children: [cc(x, { w: M[0], bold: true, size: 24 }), jn(M[1], "___"), jn(M[2], "___"), jn(M[3], "___"), cc("___ von 3", { w: M[4], size: 17, color: "5B6B8C" }), cc("", { w: M[5] }), cc("___ von ___", { w: M[6], size: 17, color: "5B6B8C" }), cc("", { w: M[7] })] })),
  ] });
  return [
    p([t("BEOBACHTUNGSZETTEL · GESCHICHTE 8D   ", { size: 15, bold: true, color: MUTED }), t("Abstimmung der Stände", { font: "Cambria", size: 26, bold: true, color: NAVY })], { after: 40 }),
    p([t("Mein Stand:  ", { size: 19, bold: true }), t("□ Klerus  □ Adel  □ Dritter Stand", { size: 19 }), t("       Meine Stimme (vorab):  ", { size: 19, bold: true }), t("A □ Ja □ Nein    B □ Ja □ Nein", { size: 19 })], { after: 50 }),
    matrix,
    p(t("Schreibe mit, was die Lehrkraft an die Tafel schreibt. Nach Ständen: mindestens 2 von 3 Ständen Ja = angenommen. Nach Köpfen: mehr als die Hälfte Ja = angenommen.", { size: 16, italics: true, color: MUTED }), { before: 40, after: 0 }),
  ];
}
function zettelCell() {
  return new TableCell({ children: zettel(), width: { size: W, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 80, bottom: 60, left: 170, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
const bogenPage = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [0, 1, 2, 3].map(() => new TableRow({ cantSplit: true, height: { value: 3450, rule: HeightRule.EXACT }, children: [zettelCell()] })) });

const g = (txt, o = {}) => new TableCell({ children: [p(t(txt, { size: o.size || 19, bold: o.bold, color: o.color }), { after: 0, align: o.align || AlignmentType.CENTER })], width: { size: o.w, type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 30, bottom: 30, left: 50, right: 50 }, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span });
const brk = () => new Paragraph({ children: [new PageBreak()] });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 450, left: 600, right: 600 } } }, children: [
    page(pg1), brk(), page(pg2), brk(), bogenPage,
  ] }] });
  fs.writeFileSync("../Druckvorlage_Karten_Beobachtungszettel_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
