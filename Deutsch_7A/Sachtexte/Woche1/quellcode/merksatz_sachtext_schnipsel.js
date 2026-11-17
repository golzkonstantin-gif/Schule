// Merksatz-Schnipsel Mi 18.11.: „Merkmale eines Sachtexts“ + „Zeitformen im Sachtext“ – 2 Sätze pro A4 (zum Halbieren), ohne Farbflächen
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", YEL = "C98A1E", GRAY = "9AA3B5";
const PW = 11906 - 1100;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.it, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 50 }, indent: o.indent });
const b = (style, size, color) => ({ style, size, color });
const bs = (x) => ({ top: x, bottom: x, left: x, right: x });
const NB = bs(b(BorderStyle.NONE, 0, "FFFFFF"));
const cell = (ch, w, o = {}) => new TableCell({ children: ch, width: { size: w, type: WidthType.DXA }, borders: o.borders || NB, verticalAlign: VerticalAlign.TOP, margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 120, right: o.ml ?? 120 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, c) => a + c, 0), type: WidthType.DXA }, rows });
const M = [["Überschrift", "(oft mit einführender Information) nennt das Thema."], ["Teilüberschriften und Absätze", "gliedern den Text."], ["Fakten:", "Zahlen, Daten, Namen, die man überprüfen kann."], ["Fachbegriffe", "werden oft erklärt (Fußnote, „also …“)."], ["Beispiele und Erklärungen", "machen die Informationen verständlich."], ["Sachliche Sprache:", "keine Meinung, kein „ich“, keine Spannung."]];
function satz() {
  const inner = PW - 300;
  return table([PW], [new TableRow({ children: [cell([
    p([t("Merksatz · ", { bold: true, color: RED, size: 20 }), t("Merkmale eines Sachtexts", { font: "Cambria", bold: true, size: 26, color: NAVY })], { after: 80 }),
    table([inner], [new TableRow({ children: [cell([p(t("Ein Sachtext informiert über ein Thema. Er stellt Tatsachen sachlich dar.", { bold: true, size: 23 }), { after: 0 })], inner, { borders: bs(b(BorderStyle.SINGLE, 14, RED)), m: 80 })] })]),
    p(t(""), { after: 40 }),
    ...M.map(([a, c], i) => p([t(`${i + 1}  `, { bold: true, color: NAVY }), t(a + " ", { bold: true }), t(c)], { after: 30, indent: { left: 300, hanging: 300 } })),
    p([t("Wofür? ", { bold: true, color: RED }), t("Du erkennst, aus welchem Text du Informationen entnehmen kannst – für Referate und die Klassenarbeit.", { size: 20 })], { before: 60, after: 20 }),
    p([t("Test: ", { bold: true, color: RED }), t("Will der Text informieren – nicht unterhalten oder überzeugen? Kann man die Angaben überprüfen?", { size: 20 })], { after: 120 }),
    p([t("Merksatz · ", { bold: true, color: RED, size: 20 }), t("Zeitformen im Sachtext", { font: "Cambria", bold: true, size: 26, color: NAVY })], { after: 60 }),
    p([t("Präsens: ", { bold: true }), t("was heute noch gilt – „Beispielsweise "), t("weiß", { bold: true, u: true }), t(" man, …“")], { after: 20 }),
    p([t("Präteritum: ", { bold: true }), t("Bericht über früher – „Sie "), t("warnten", { bold: true, u: true }), t(" vor Gefahren.“")], { after: 20 }),
    p([t("Plusquamperfekt: ", { bold: true }), t("was davor geschah – „…, dass sie Wild "), t("aufgespürt", { bold: true, color: YEL }), t(" "), t("hatten", { bold: true, color: RED }), t("“")], { after: 0 }),
  ], PW, { borders: bs(b(BorderStyle.DASHED, 6, GRAY)), m: 140, ml: 200 })] })]);
}
const cut = p(t("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", { color: GRAY, size: 16 }), { align: AlignmentType.CENTER, before: 100, after: 100 });
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 400, left: 550, right: 550 } } };
Packer.toBuffer(new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: props, children: [satz(), cut, satz()] }] }))
  .then((buf) => fs.writeFileSync(__dirname + "/../Merksatz_Sachtext_Zeitformen_Schnipsel.docx", buf));
