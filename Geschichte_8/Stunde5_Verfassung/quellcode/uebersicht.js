const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageOrientation } = require("docx");
const E = require("./daten.js"); const MIX = E.MIX;
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT2 = "EAF0FA", GOLD = "E8A33D", GOLDL = "FFF6E0", GREEN = "1B5E3A", RED = "D9534F";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 0 }, tabStops: o.tabs });
const bd = (c, s = 6, st = BorderStyle.SINGLE) => ({ style: st, size: s, color: c });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const nb = { top: none, bottom: none, left: none, right: none };
const cell = (ch, o) => new TableCell({ children: Array.isArray(ch) ? ch : [ch], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.b || nb, verticalAlign: o.v || VerticalAlign.CENTER, margins: { top: 20, bottom: 20, left: o.ml ?? 100, right: 80 } });
const page = (children) => ({ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 600, right: 600 } } }, children });
const NUM = 500, BOXW = 7150, GAP = 338, RH = 1300;

function kopf(titel, sub) {
  const band = new Table({ width: { size: 15638, type: WidthType.DXA }, columnWidths: [15638], rows: [new TableRow({ height: { value: 90, rule: "exact" }, children: [cell(p(t("", { size: 2 })), { w: 15638, fill: GOLD })] })] });
  return [band, p([t("GESCHICHTE 8C · FRANZÖSISCHE REVOLUTION   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________________", { size: 18, color: MUTED })], { before: 60, after: 20 }),
    p([t(titel, { font: "Cambria", size: 32, bold: true, color: NAVY }), t("   " + sub, { size: 19, color: MUTED })], { after: 80 })];
}
function numCell(n, w) { return cell(p(t(String(n), { font: "Cambria", size: 26, bold: true, color: "FFFFFF" }), { align: AlignmentType.CENTER }), { w, fill: NAVY, b: nb, ml: 0 }); }

// Blatt mit leeren Feldern (loesung: Felder gefüllt)
function blatt(loesung) {
  const rows = [];
  for (let r = 0; r < 7; r++) {
    const a = r + 1, b = r + 8;
    const box = (n) => cell(loesung ? p([t(E[n - 1][0] + ": ", { bold: true, size: 20, color: GREEN }), t(E[n - 1][1], { size: 20, color: GREEN, italics: true })]) : p(t(r === 0 && n === 1 ? "" : "", { size: 10 })), { w: BOXW, b: { top: bd("8A97B8", 6, BorderStyle.DASHED), bottom: bd("8A97B8", 6, BorderStyle.DASHED), left: bd("8A97B8", 6, BorderStyle.DASHED), right: bd("8A97B8", 6, BorderStyle.DASHED) }, fill: "FFFFFF" });
    rows.push(new TableRow({ height: { value: RH, rule: "exact" }, children: [numCell(a, NUM), box(a), cell(p(t("")), { w: GAP }), numCell(b, NUM), box(b)] }));
  }
  const tb = new Table({ width: { size: NUM * 2 + BOXW * 2 + GAP, type: WidthType.DXA }, columnWidths: [NUM, BOXW, GAP, NUM, BOXW], rows });
  const foot = loesung
    ? p([t("deshalb-Pfeile (Beispiele): ", { bold: true, size: 18, color: RED }), t("14. Juli (11) → 4. August (13): Der König kann Aufstände nicht stoppen, deshalb geben Klerus und Adel nach.   17. Juni (8) → 20. Juni (9): Er erkennt die Versammlung nicht an, deshalb wird der Saal gesperrt.   1787 (3) → August 1788 (5): Das Geld fehlt, deshalb ruft er die Generalstände.", { size: 17, color: MUTED })], { before: 60 })
    : p([t("Klebe die Schnipsel erst nach der Kontrolle auf. Dann zeichnet ihr gemeinsam „deshalb“-Pfeile ein.", { size: 17, italics: true, color: MUTED })], { before: 60 });
  return [...kopf(loesung ? "Lösung: Der Weg zur Menschenrechtserklärung" : "Der Weg zur Menschenrechtserklärung", loesung ? "(für die Lehrkraft)" : "Meine Übersicht: von 1715 bis 1789"), tb, foot];
}
// Schnipselblatt (gemischt), Schnittlinien gestrichelt
function schnipsel() {
  const SW = 7000, SH = 1240;
  const rows = [];
  for (let r = 0; r < 7; r++) {
    const mk = (i) => { const e = E[MIX[i] - 1]; return cell(p([t(e[0] + ": ", { bold: true, size: 23, color: NAVY }), t(e[1], { size: 23 })]), { w: SW, b: { top: bd("5B6B8C", 6, BorderStyle.DASHED), bottom: bd("5B6B8C", 6, BorderStyle.DASHED), left: bd("5B6B8C", 6, BorderStyle.DASHED), right: bd("5B6B8C", 6, BorderStyle.DASHED) }, fill: "FFFFFF" }); };
    rows.push(new TableRow({ height: { value: SH, rule: "exact" }, children: [mk(r), mk(r + 7)] }));
  }
  const tb = new Table({ width: { size: SW * 2, type: WidthType.DXA }, columnWidths: [SW, SW], rows, alignment: AlignmentType.CENTER });
  return [...kopf("Schnipsel", "Ausschneiden an den gestrichelten Linien (Lehrkraft: Stapel schneiden, 14 Streifen pro Person)"), tb];
}
async function save(name, children) {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [page(children)] });
  fs.writeFileSync("../" + name, await Packer.toBuffer(doc)); console.log(name);
}
(async () => {
  await save("Uebersicht_Blatt_Vorgeschichte_8C.docx", blatt(false));
  await save("Uebersicht_Schnipsel_Vorgeschichte_8C.docx", schnipsel());
  await save("Loesung_Uebersicht_Vorgeschichte_8C.docx", blatt(true));
})();
