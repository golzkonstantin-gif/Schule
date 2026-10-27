const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, keepNext: o.keepNext });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const LINE = solid("5B6B8C", 8);
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(LINE), verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: 110, right: 110 } });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: "atLeast" } : undefined, cantSplit: true });
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const h2 = (s) => p(t(s, { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 80, keepNext: true });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Materialliste Generalstände · Mi 04.11.2026 · Seite ", PageNumber.CURRENT] })] })] });

const hdr = (cols, widths) => row(cols.map((x, i) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), { w: widths[i], fill: NAVY, borders: all(solid(NAVY, 8)) })), 320);
const box = () => p(t("□", { size: 26 }), { after: 0, align: AlignmentType.CENTER });

// Druck
const DW = [550, 3500, 1700, 4456];
const druck = [
  ["Text „Wie kam es zu den Generalständen?“ (Textblatt mit Zeitstrahl)", "27 (25 + 2 Reserve), einseitig A4", "`Text_Wie_kam_es_zu_den_Generalstaenden.pdf`"],
  ["Druckvorlage Standeskarten und Stimmzettel, Seite 1 bis 3", "je 1 Seite, ausschneiden: 25 Karten + 5 Reserve", "`Druckvorlage_Standeskarten_Stimmzettel.pdf`: Seite 1 (3 Klerus, 4 Adel, 3 Dritter Stand), Seite 2 (10 Dritter Stand), Seite 3 (5 Dritter Stand, 5 Reserve)"],
  ["Absprache-Zettel (Seite 4 derselben Datei)", "1 Seite (zwei Zettel, einer genügt)", "für den Tisch vorn, geheim; wird nach der Absprache eingesammelt"],
  ["Tests der Klasse", "falls noch nicht zurückgegeben: verdeckt auf die Plätze", "keine Besprechung, Berichtigung mit Grün bis zur nächsten Stunde"],
];
// Vorbereitung
const prep = [
  ["Tafelskript ausdrucken oder auf dem Tablet bereithalten", "`Tafelskript_Generalstaende_8D.pdf`, 7 Seiten"],
  ["Folien öffnen und einmal durchklicken", "`Praesentation_Generalstaende_8D.pptx`, 17 Folien; Folie 1 steht beim Betreten da"],
  ["Rollen verteilen und Karten verdeckt auf die Plätze legen", "Klerus 3 und Adel 4 gezielt wählen (ruhige, zuverlässige Schülerinnen und Schüler), Rest Dritter Stand; Tabelle unten ausfüllen (ohne Namen, nur Platznummern)"],
  ["Tisch vorn (Hinterzimmer) mit 7 Stühlen aufstellen", "Absprache-Zettel dort ablegen"],
  ["Kreide und Grün zum Verbessern bereitlegen", "Kreide-Kette: 1 Stück Kreide, das weitergegeben wird"],
  ["Uhr oder Timer bereitstellen", "Übungsphase endet 5 Minuten vor Schluss, „Auswerten“ ist fest"],
  ["Stimmzettel sammeln und zählen vorbereiten", "3 Stapel (Klerus, Adel, Dritter Stand), Stift für die Tafelzahlen"],
];
// Tafel
const tafel = [
  ["Seitentafel", "Ziele („Ich kann …“, 3 Sätze) und Ablauf: A Ankommen · V Wiederholung · I Text und Zeitstrahl · V Ständeversammlung · A Auswerten (Ü3 Satz 3 mit *)"],
  ["Mitte, links", "Daten für die Kreide-Kette untereinander: 1614 · 1778–1783 · 1787 · 1788 · August 1788 · Frühjahr 1789 · 5. Mai 1789"],
  ["Mitte, rechts", "Leerer Rahmen für das Ergebnis der Abstimmung („nach Ständen: Klerus __ Adel __ Dritter Stand __ → __ : __“, „nach Köpfen: Ja __ Nein __“)"],
];

const d = [];
d.push(p(t("GESCHICHTE · KLASSE 8D · MITTWOCH 04.11.2026", { size: 18, bold: true, color: MUTED }), { after: 20 }));
d.push(p(t("Materialliste: Generalstände", { font: "Cambria", size: 44, bold: true, color: NAVY }), { after: 40 }));
d.push(p(t("Zum Abhaken. Alle Dateien liegen in Geschichte_8D/Generalstaende/. Ablauf: 3 Ankommen · 6 Wiederholung · 11 Text und Zeitstrahl · 17 Ständeversammlung · 5 Auswerten · 3 Puffer.", { size: 20, color: MUTED }), { after: 80 }));

d.push(h2("1 Drucken und schneiden"));
d.push(table(DW, [hdr(["", "Was", "Menge", "Datei und Hinweis"], DW), ...druck.map(([a, b, c]) => row([cell(box(), { w: DW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: DW[1] }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: DW[2] }), cell(p(t(c.replace(/`/g, ""), { size: 19 }), { after: 0 }), { w: DW[3] })], 560))]));

d.push(h2("2 Vorbereiten"));
const PW = [550, 4100, 5556];
d.push(table(PW, [hdr(["", "Aufgabe", "Hinweis"], PW), ...prep.map(([a, b]) => row([cell(box(), { w: PW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: PW[1] }), cell(p(t(b.replace(/`/g, ""), { size: 19 }), { after: 0 }), { w: PW[2] })], 520))]));

d.push(h2("3 Tafel vor der Stunde"));
const TW = [2000, 8206];
d.push(table(TW, [hdr(["Wo", "Was steht dort"], TW), ...tafel.map(([a, b]) => row([cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: TW[0], fill: LIGHT }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: TW[1] })], 520))]));

d.push(h2("4 Rollenverteilung (ohne Namen, nur Platznummern)"));
const RW = [2400, 1100, 6706];
const rolle = [["Erster Stand · Klerus", "3", "Plätze: ____  ____  ____"], ["Zweiter Stand · Adel", "4", "Plätze: ____  ____  ____  ____"], ["Dritter Stand (Bauern und Bürger)", "18", "alle übrigen Plätze; bei Fehlzeiten Reserve-Karten einsetzen"]];
d.push(table(RW, [hdr(["Stand", "Anzahl", "Plätze"], RW), ...rolle.map(([a, b, c]) => row([cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: RW[0], fill: LIGHT }), cell(p(t(b, { size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: RW[1] }), cell(p(t(c, { size: 20 }), { after: 0 }), { w: RW[2] })], 520))]));
d.push(p(t("Wenn Schülerinnen oder Schüler fehlen: Klerus bleibt 3, Adel bleibt 4. Klerus und Adel dürfen zusammen nie mehr als ein Drittel der Klasse sein.", { size: 19, italics: true, color: MUTED }), { before: 60 }));

d.push(h2("5 Kurz vor Beginn (5 Minuten)"));
["Folie 1 steht, Beamer an", "Karten liegen verdeckt auf den Plätzen", "Tisch vorn steht", "Seitentafel und Datumsliste an der Tafel", "Textblätter liegen bereit (nicht schon austeilen)", "Lautstärke-Ansagen im Kopf: „Stifte liegen. Ihr hört nur zu.“ · „Flüsterstufe. Nur dein Partner hört dich.“"].forEach((x) => d.push(p([t("□  ", { size: 24 }), t(x, { size: 21 })], { after: 50 })));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 850, right: 850 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Materialliste_Generalstaende_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
