const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, ImageRun, TabStopType, LeaderType, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", ACC = "7A1F3D", ACCL = "F7ECF0", L1 = "F4F7FD", GREEN = "1B5E3A", W = 10700;
const loesung = process.argv[2] === "loesung";
const t = (x, o = {}) => new TextRun({ text: x, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (r, o = {}) => new Paragraph({ children: Array.isArray(r) ? r : [r], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 40 }, keepNext: o.keepNext, tabStops: o.tabs, indent: o.indent });
const sd = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const bx = (c, s) => ({ top: sd(c, s), bottom: sd(c, s), left: sd(c, s), right: sd(c, s) });
const NB = { top: nb, bottom: nb, left: nb, right: nb };
const cell = (ch, w, o = {}) => new TableCell({ children: Array.isArray(ch) ? ch : [ch], width: { size: w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.b || bx("C5CEE6", 6), margins: { top: 30, bottom: 30, left: 120, right: 120 }, verticalAlign: o.v || VerticalAlign.TOP });
const table = (ws, rows) => new Table({ width: { size: ws.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: ws, rows: rows.map((r) => new TableRow({ children: r, cantSplit: true })) });
const lineT = [{ type: TabStopType.RIGHT, position: W - 400, leader: LeaderType.UNDERSCORE }];
const IMG = new ImageRun({ type: "png", data: fs.readFileSync("../material/schaubild_verfassung.png"), transformation: { width: 570, height: 285 } });
// Auswahlzeile: Frage + Kästchen (Lösung: richtiges angekreuzt)
function opt(frage, optionen, ok) {
  const runs = [t(frage + "   ", { size: 21 })];
  optionen.forEach((o, i) => { const on = loesung && i === ok; runs.push(t((on ? "☒ " : "☐ ") + o + "     ", { size: 21, bold: on, color: on ? GREEN : "000000" })); });
  return p(runs, { after: 30 });
}
function multi(frage, optionen, oks) {
  const runs = [t(frage + "   ", { size: 21 })];
  optionen.forEach((o, i) => { const on = loesung && oks.includes(i); runs.push(t((on ? "☒ " : "☐ ") + o + "    ", { size: 21, bold: on, color: on ? GREEN : "000000" })); });
  return p(runs, { after: 30 });
}
function block(nr, titel, pt, hinweis, lines) {
  return table([W], [[cell([p([t(nr + "  ", { font: "Cambria", bold: true, size: 26, color: ACC }), t(titel, { bold: true, size: 22, color: NAVY }), t("   (" + pt + " Punkte, je Teilaufgabe 1 Punkt)", { size: 18, color: MUTED }), t("   " + hinweis, { size: 18, italics: true, color: MUTED })], { after: 40 }), ...lines], W, { fill: "FFFFFF" })]]);
}
const gap = () => p(t("", { size: 2 }), { after: 10 });
const d = [];
d.push(table([W], [[cell(p(t("", { size: 2 })), W, { fill: ACC, b: NB })]]));
d.push(p([t("GESCHICHTE 8C · FRANZÖSISCHE REVOLUTION   ", { size: 16, bold: true, color: MUTED }), t(loesung ? "Lösung (Lehrkraft)" : "Name: ______________________________", { size: 18, color: MUTED }), t("\tPunkte: ___ / 15 (+1)", { size: 18, color: MUTED })], { before: 40, after: 10, tabs: [{ type: TabStopType.RIGHT, position: W }] }));
d.push(p([t("Das Verfassungsschaubild von 1791", { font: "Cambria", size: 32, bold: true, color: ACC }), t("   AB 2 · Schaubild lesen und abhaken", { size: 19, color: MUTED })], { after: 50 }));
d.push(p(IMG, { align: AlignmentType.CENTER, after: 50 }));
d.push(p(t("Schaubild nach dem Schulbuch S. 115 (vereinfacht). Tipp: Lies in drei Schritten: Thema, Beschreibung, Beurteilung (S. 86). Kreuze an: ☐ = nicht gewählt, ☒ = gewählt.", { size: 17, italics: true, color: MUTED }), { after: 50 }));
d.push(block("1", "Thema", 3, "Kreuze an.", [opt("a) Die Verfassung gilt für", ["Frankreich", "England", "Preußen"], 0), opt("b) Beschlossen hat sie", ["der König", "die Nationalversammlung", "die Generalstände"], 1), opt("c) Regierungsform", ["absolute Monarchie", "konstitutionelle Monarchie", "Republik"], 1)]), gap());
d.push(block("2", "Wahlrecht", 3, "Wer durfte wählen, wer nicht? Sieh unten im Schaubild nach.", [
  multi("a) Das Wahlrecht hing ab von (3 Kreuze):", ["nur Männer", "über 25 Jahre", "Steuern gezahlt", "Religion", "Wohnort"], [0, 1, 2]),
  multi("b) Wer durfte nicht wählen? (3 Kreuze)", ["Frauen", "Männer unter 25", "arme Männer", "reiche Männer über 25"], [0, 1, 2]),
  opt("c) Die ca. 22 Millionen ohne Wahlrecht heißen:", ["Aktivbürger", "Passivbürger"], 1)]), gap());
d.push(block("3", "Gewalten", 3, "Wer macht was? Die Überschriften im Schaubild helfen.", [
  opt("a) Wer macht die Gesetze?", ["König und Regierung", "Nationalversammlung", "Gerichte"], 1),
  opt("b) Wer führt die Gesetze aus und regiert das Land?", ["König und Regierung", "Nationalversammlung", "Gerichte"], 0),
  opt("c) Wer urteilt nach den Gesetzen?", ["König und Regierung", "Nationalversammlung", "Gerichte"], 2)]), gap());
d.push(block("4", "Macht und Kontrolle", 3, "Wahr oder falsch?", [opt("a) Der König ernennt und entlässt die Regierung.", ["wahr", "falsch"], 0), opt("b) Sein Einspruch kann ein Gesetz für immer verhindern.", ["wahr", "falsch"], 1), opt("c) Die Nationalversammlung kontrolliert den König.", ["wahr", "falsch"], 0)]), gap());
const bel = loesung ? [p([t("c) Mein Beleg: ", { bold: true, size: 21 }), t("Mehrere gute Antworten: Wer „König“ wählt: Er ernennt die Regierung und hat ein Einspruchsrecht. Wer „Nationalversammlung“ wählt: Sie beschließt die Gesetze, wählt alle zwei Jahre und kontrolliert den König (Art. 3: Er regiert nur durch das Gesetz). 1 P für einen Beleg aus dem Schaubild, der zur Entscheidung passt.", { size: 19, color: GREEN, italics: true })], { after: 40 })] : [p([t("c) Mein Beleg: ", { bold: true, size: 21 }), t("Im Schaubild sieht man, dass …", { size: 19, italics: true, color: MUTED })], { after: 20 }), p(t("\t"), { tabs: lineT, after: 90, before: 50 })];
const zus = loesung ? [p([t("Zusatz (1 Punkt): ", { bold: true, size: 20 }), t("Gerecht? Ja oder nein mit Maßstab, zum Beispiel: Nein, weil über 80 von 100 nicht wählen durften.", { size: 19, color: GREEN, italics: true })], { after: 0 })] : [p([t("Zusatz (1 Punkt): ", { bold: true, size: 20 }), t("☐ gerecht   ☐ nicht gerecht, weil", { size: 21 })], { after: 20 }), p(t("\t"), { tabs: lineT, after: 20, before: 50 })];
d.push(block("5", "Anteil und Urteil", 3, "a) und b) ankreuzen, c) ein Satz.", [opt("a) Von 100 Menschen durften nur etwa 16 wählen. Das heißt:", ["fast alle durften", "die meisten durften nicht", "die Hälfte durfte"], 1), opt("b) Wer hatte mehr Macht?", ["der König", "die Nationalversammlung", "beide etwa gleich"], 1), ...bel, ...zus]));
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8C · Stunde 5 · AB 2"] })] })] });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 450, left: 600, right: 600 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync(loesung ? "../Loesung_AB2_Verfassungsschaubild_8C.docx" : "../AB2_Verfassungsschaubild_8C.docx", await Packer.toBuffer(doc)); console.log("ok");
})();
