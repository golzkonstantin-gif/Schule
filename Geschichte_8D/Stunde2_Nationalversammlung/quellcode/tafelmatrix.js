const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, HeightRule, PageOrientation, PageBreak } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 } });
const line = { style: BorderStyle.SINGLE, size: 8, color: "5B6B8C" };
const W = 15000; // Querformat nutzbar
const c = (txt, w, o = {}) => new TableCell({ children: [p(t(txt, { size: o.size || 28, bold: o.bold, color: o.color }), { after: 0, align: o.align || AlignmentType.CENTER })], width: { size: w, type: WidthType.DXA }, borders: { top: line, bottom: line, left: line, right: line }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span, shading: o.fill ? { fill: o.fill, type: "clear", color: "auto" } : undefined });
const M = [1300, 1800, 1800, 2200, 1700, 2300, 1700, 2300];
const matrix = (filled) => new Table({ width: { size: M.reduce((x, y) => x + y, 0), type: WidthType.DXA }, columnWidths: M, rows: [
  new TableRow({ children: [c("", M[0]), c("Nach Ständen (jeder Stand hat eine Stimme)", M[1] + M[2] + M[3] + M[4] + M[5], { span: 5, bold: true, size: 26, fill: "EAF0FA" }), c("Nach Köpfen (jede Hand zählt)", M[6] + M[7], { span: 2, bold: true, size: 26, fill: "EAF0FA" })] }),
  new TableRow({ children: [c("Antrag", M[0], { bold: true }), c("Klerus (3)", M[1], { bold: true, color: "6B4A9E" }), c("Adel (4)", M[2], { bold: true, color: "B9791A" }), c("Dritter Stand (18)", M[3], { bold: true, color: "2F6AAE" }), c("Stände mit Ja", M[4], { bold: true, size: 24 }), c("Ergebnis nach Ständen", M[5], { bold: true, size: 24 }), c("alle Ja (von 25)", M[6], { bold: true, size: 24 }), c("Ergebnis nach Köpfen", M[7], { bold: true, size: 24 })] }),
  ...[["A", "Steuern auch für Adel und Klerus"], ["B", "Stände abschaffen, gleiche Rechte"]].map(([x, txt]) => new TableRow({ height: { value: 1300, rule: HeightRule.ATLEAST }, children: [c(x, M[0], { bold: true, size: 40 }), c(filled ? "Ja 0" : "Ja ___", M[1]), c(filled ? "Ja 0" : "Ja ___", M[2]), c(filled ? "Ja 18" : "Ja ___", M[3]), c(filled ? "1 von 3" : "___ von 3", M[4]), c(filled ? "abgelehnt" : "", M[5], { bold: true, color: RED }), c(filled ? "18 von 25" : "___ von 25", M[6]), c(filled ? "angenommen" : "", M[7], { bold: true, color: RED })] })),
] });
const SW = [3600, 2800, 2800, 2800, 2800];
const strich = new Table({ width: { size: SW.reduce((x, y) => x + y, 0), type: WidthType.DXA }, columnWidths: SW, rows: [
  new TableRow({ children: [c("Stand", SW[0], { bold: true }), c("A Ja", SW[1], { bold: true }), c("A Nein", SW[2], { bold: true }), c("B Ja", SW[3], { bold: true }), c("B Nein", SW[4], { bold: true })] }),
  ...[["Klerus (3)", "6B4A9E"], ["Adel (4)", "B9791A"], ["Dritter Stand (18)", "2F6AAE"]].map(([x, col]) => new TableRow({ height: { value: 1500, rule: HeightRule.ATLEAST }, children: [c(x, SW[0], { bold: true, color: col, align: AlignmentType.LEFT }), ...SW.slice(1).map((w) => c("", w))] })),
] });
const head = (kicker, title, sub) => [p(t(kicker, { size: 18, bold: true, color: MUTED }), { after: 20 }), p(t(title, { font: "Cambria", size: 44, bold: true, color: NAVY }), { after: 40 }), p(t(sub, { size: 24, italics: true, color: MUTED }), { after: 160 })];
const brk = () => new Paragraph({ children: [new PageBreak()] });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 700, bottom: 600, left: 700, right: 700 } } }, children: [
    ...head("GESCHICHTE 8D · VORLAGE FÜR DIE TAFEL", "Auswertungsmatrix: Abstimmung der Stände", "Leer. Diese Matrix zeichnest du vorab an die Tafel und füllst sie nach dem Zählen aus. Die Klasse rechnet danach die Spalten „Stände mit Ja“, „Ergebnis“ und „alle Ja“ selbst aus."),
    matrix(false),
    p(t("Regel: Nach Ständen: mindestens 2 von 3 Ständen Ja = angenommen. Nach Köpfen: mehr als die Hälfte aller Ja = angenommen.", { size: 24, italics: true, color: MUTED }), { before: 160 }),
    brk(),
    ...head("GESCHICHTE 8D · LÖSUNG FÜR DIE LEHRKRAFT", "Auswertungsmatrix: erwartetes Ergebnis", "Echte Zahlen gelten. Erwartet bei 25 Personen: Klerus und Adel stimmen Nein, der Dritte Stand (Bauern und Bürger) stimmt Ja."),
    matrix(true),
    p(t("Nach Ständen: Klerus Nein, Adel Nein, Dritter Stand Ja = 1 von 3, abgelehnt. Nach Köpfen: 18 von 25 Ja = angenommen. Dieselben Stimmen, ein anderes Ergebnis. (Bei A und B gleich, weil Klerus und Adel immer ihre Vorrechte verteidigen.)", { size: 24, color: NAVY }), { before: 160 }),
    brk(),
    ...head("GESCHICHTE 8D · AUSZÄHLHILFE FÜR DIE LEHRKRAFT", "Strichliste zum Auszählen der Karten", "Sortiere die Karten nach Farbe (Stand) und zähle Ja und Nein je Antrag. Kontrolle: Ja plus Nein ist je Stand die Personenzahl (3, 4, 18)."),
    strich,
    p(t("Danach trägst du nur die Ja-Zahlen in die Auswertungsmatrix an der Tafel ein.", { size: 24, italics: true, color: MUTED }), { before: 160 }),
  ] }] });
  fs.writeFileSync("../Tafelmatrix_Auswertung_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
