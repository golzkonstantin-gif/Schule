const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber, PageBreak } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", L1 = "F4F7FD", L2 = "EAF0FA", GOLD = "C98A1E", GOLDL = "FFF6E0", RED = "D9534F", W = 10500;
const t = (x, o = {}) => new TextRun({ text: x, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (r, o = {}) => new Paragraph({ children: Array.isArray(r) ? r : [r], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, keepNext: o.keepNext, indent: o.indent });
const sd = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const bx = (c, s) => ({ top: sd(c, s), bottom: sd(c, s), left: sd(c, s), right: sd(c, s) });
const NB = { top: nb, bottom: nb, left: nb, right: nb };
const cell = (ch, w, o = {}) => new TableCell({ children: Array.isArray(ch) ? ch : [ch], width: { size: w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.b || bx("C5CEE6", 6), margins: { top: 70, bottom: 70, left: 120, right: 120 }, verticalAlign: o.v || VerticalAlign.TOP, columnSpan: o.span });
const table = (ws, rows) => new Table({ width: { size: ws.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: ws, rows });
const row = (cells, o = {}) => new TableRow({ children: cells, cantSplit: true, height: o.h ? { value: o.h, rule: "atLeast" } : undefined });
const band = () => table([W], [new TableRow({ height: { value: 100, rule: "exact" }, children: [cell(p(t("", { size: 2 })), W, { fill: GOLD, b: NB })] })]);
const kopf = (titel, sub) => [band(), p([t("GESCHICHTE · KLASSE 8   ", { size: 17, bold: true, color: MUTED }), t("Name: ______________________________", { size: 18, color: MUTED })], { before: 60, after: 10 }), p([t(titel, { font: "Cambria", size: 34, bold: true, color: NAVY }), t("   " + sub, { size: 22, color: MUTED })], { after: 80 })];
const h = (s) => p(t(s, { font: "Cambria", size: 25, bold: true, color: NAVY }), { before: 100, after: 50, keepNext: true });
const bullet = (parts, o = {}) => p(Array.isArray(parts) ? parts : t(parts, { size: o.size || 20 }), { indent: { left: 240 }, after: o.after ?? 30 });
const B = (s) => t(s, { bold: true, size: 22 }), T = (s) => t(s, { size: 22 });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Lernübersicht Absolutismus und Stände · Seite ", PageNumber.CURRENT] })] })] });
const d = [];
// ---------- Seite 1: Absolutismus ----------
d.push(...kopf("Lernübersicht: Absolutismus", "Was du für den Test können musst (1/2)"));
d.push(table([W], [row([cell([p([t("Absolutismus: ", { bold: true, color: NAVY, size: 22 }), t("Herrschaftsform in Europa im 17. und 18. Jahrhundert. Ein einzelner Monarch (König oder Fürst) herrscht allein und uneingeschränkt („absolut“). Niemand kann ihn stoppen: kein Parlament, kein Gericht.", { size: 22 })], { after: 40 }), p([t("Test-Tipp: ", { bold: true, size: 21, color: GOLD }), t("Wenn eine Aussage zum Absolutismus passt oder nicht passt, begründe immer mit der Macht des Königs („ein Parlament könnte ihn begrenzen, also passt es nicht“).", { size: 21, italics: true })], { after: 0 })], W, { fill: GOLDL, b: bx(GOLD, 10) })])]));
d.push(h("Die fünf Säulen der Macht Ludwigs XIV."));
d.push(p(t("Merke dir erst die fünf Namen, dann zu jeder Säule: Was war das? Warum sicherte es seine Macht?", { size: 21, color: MUTED, italics: true }), { after: 50 }));
const col = [2250, 3900, 4350];
const SAEULEN = [
  ["1", "Verwaltung und Justiz", "Beamte handeln in seinem Namen, in den Provinzen die Intendanten. Er ist oberster Richter: Mit „lettres de cachet“ lässt er Personen ohne Gerichtsverfahren verhaften.", "Er kontrolliert das ganze Land und kann Gegner ausschalten."],
  ["2", "Adel und Hof", "Der Adel behält Vorrechte (Steuerfreiheit, Jagdrecht) und wird an den Hof nach Versailles gebunden: Feste, Theater, Geschenke.", "Der Adel ist beschäftigt und abhängig, er leistet keinen Widerstand."],
  ["3", "Wirtschaft (Merkantilismus)", "Finanzminister Colbert: Der Staat lenkt die Wirtschaft (Manufakturen, hohe Zölle, Kolonien). Ziel: mehr fertige Waren verkaufen als einführen.", "Der Staat bekommt Geld für Hof, Beamte und Heer."],
  ["4", "Religion (Gottesgnadentum)", "Der König ist von Gott zum Herrschen ausgewählt. 1685 beendet er die religiöse Toleranz: ein König, ein Glaube.", "Seine Herrschaft gilt als gottgewollt, Einheit im Land."],
  ["5", "Heer", "Stehendes Heer statt Söldner: Berufssoldaten, einheitlich bewaffnet, in Kasernen, bis zu 400.000 Mann. Kostet in Friedenszeiten ein Drittel des Staatshaushalts.", "Er kann seine Macht nach innen und außen mit Gewalt durchsetzen."],
];
d.push(table(col, [row([cell(p(t("Säule", { bold: true, color: NAVY, size: 22 }), { after: 0 }), col[0], { fill: L2 }), cell(p(t("Was ist das?", { bold: true, color: NAVY, size: 22 }), { after: 0 }), col[1], { fill: L2 }), cell(p(t("Wie sichert es seine Macht?", { bold: true, color: NAVY, size: 22 }), { after: 0 }), col[2], { fill: L2 })]),
  ...SAEULEN.map(([n, name, was, wie]) => row([cell(p([t(n + "  ", { font: "Cambria", bold: true, size: 26, color: GOLD }), t(name, { bold: true, size: 22, color: NAVY })], { after: 0 }), col[0], { fill: L1, v: VerticalAlign.CENTER }), cell(p(t(was, { size: 21 }), { after: 0 }), col[1]), cell(p(t(wie, { size: 21 }), { after: 0 }), col[2])]))]));
d.push(h("Bonus (Zusatzpunkt): Merkantilismus"));
d.push(p(t("Der Staat lenkt die Wirtschaft, damit mehr fertige Waren (aus Manufakturen) ins Ausland verkauft als von dort eingeführt werden. Hohe Zölle machen ausländische Waren teuer.", { size: 22 }), { after: 40 }));
d.push(h("Typische Fehler vermeiden"));
d.push(bullet([B("Beamte: "), T("Sie gehorchen dem König, nicht dem Adel.")]));
d.push(bullet([B("Religion: "), T("1685 endet die Toleranz. Protestanten durften ihren Glauben nicht frei ausüben.")]));
d.push(bullet([B("Heer: "), T("Es ist ein stehendes Heer aus Berufssoldaten, keine Söldner.")]));

// ---------- Seite 2: Stände ----------
d.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }));
d.push(...kopf("Lernübersicht: Die drei Stände", "Was du für den Test können musst (2/2)"));
const c3 = [3500, 3500, 3500];
const stand = (name, anz, vor, pf, fill) => cell([p(t(name, { font: "Cambria", bold: true, size: 25, color: NAVY }), { after: 20 }), p([t("Anzahl: ", { bold: true, size: 21, color: MUTED }), t(anz, { size: 22 })], { after: 40 }), p(t("Vorrechte", { bold: true, size: 21, color: GOLD }), { after: 10 }), ...vor.map((x) => p(t("•  " + x, { size: 21 }), { indent: { left: 180 }, after: 15 })), p(t("Pflichten", { bold: true, size: 21, color: RED }), { before: 40, after: 10 }), ...pf.map((x) => p(t("•  " + x, { size: 21 }), { indent: { left: 180 }, after: 15 }))], 3500, { fill, b: bx("B8C2DB", 8) });
d.push(table(c3, [row([stand("Klerus", "ca. 150 000", ["keine Steuern zahlen", "eigene Gerichte haben"], ["für die Armen sorgen"], L1), stand("Adel", "ca. 500 000", ["keine Steuern zahlen", "alleiniges Jagdrecht"], ["verwaltende Aufgaben im Staat übernehmen"], L1), stand("Dritter Stand (Bauern und Bürger)", "ca. 20 000 000", ["keine"], ["Steuern und hohe Abgaben zahlen", "Pacht an den Grundherrn", "Kirchenzehnt an den Klerus (10 %)", "im Krieg Soldaten stellen"], GOLDL)])]));
d.push(h("Mitbestimmung"));
d.push(bullet([B("Politisch mitbestimmen durfte "), T("der Adel."), T("  "), B("Ausgeschlossen "), T("war der Dritte Stand (Bauern und Bürger), obwohl er fast alle Menschen umfasste.")]));
d.push(h("Merkregeln"));
d.push(bullet([B("Wer zahlt, wer nicht? "), T("Klerus und Adel zahlen keine Steuern (Vorrecht). Steuern und Abgaben zahlt der Dritte Stand (Bauern und Bürger).")]));
d.push(bullet([B("Wer zahlt an wen? "), T("Der Dritte Stand (Bauern und Bürger) zahlt den Kirchenzehnt an den Klerus, Pacht an die Grundherren und Steuern an den König. Nie umgekehrt.")]));
d.push(bullet([B("Heer: "), T("Hohe Posten in der Armee bekamen nur Adlige.")]));
d.push(bullet([B("Zahlenvergleich: "), T("Wenige haben Vorrechte (ca. 650 000), sehr viele keine (ca. 20 000 000).")]));
d.push(h("So übst du (35 Minuten Test)"));
d.push(table([5200, 5300], [row([cell([p(t("Aufgabentypen", { bold: true, color: NAVY, size: 22 }), { after: 30 }), bullet("Begriff „Absolutismus“ erklären (2 Sätze)", { after: 15 }), bullet("Stände zuordnen: Anzahl, Vorrechte, Pflichten", { after: 15 }), bullet("Wer durfte mitbestimmen? (zwei Aussagen)", { after: 15 }), bullet("Fünf Säulen nennen, eine erklären", { after: 15 }), bullet("Wahr oder falsch: bei „falsch“ begründen oder verbessern!", { after: 0 })], 5200, { fill: L1 }), cell([p(t("Übe so", { bold: true, color: NAVY, size: 22 }), { after: 30 }), bullet("Decke die rechte Spalte der Säulen ab und erkläre sie dir selbst.", { after: 15 }), bullet("Zeichne die drei Stände aus dem Kopf und trage Anzahl, Vorrechte, Pflichten ein.", { after: 15 }), bullet("Schreibe zu jedem Satz auf Seite 1 und 2 eine falsche Version und verbessere sie mit deinem Partner.", { after: 15 }), bullet("Erkläre einer Person, warum der König die Macht „absolut“ nennt.", { after: 0 })], 5300, { fill: GOLDL })])]));
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 650, left: 700, right: 700 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Lernuebersicht_Test_Absolutismus_Staende.docx", await Packer.toBuffer(doc)); console.log("ok");
})();
