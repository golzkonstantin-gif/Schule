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
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8C · Materialliste Nationalversammlung · Fr 06.11.2026 · Seite ", PageNumber.CURRENT] })] })] });

const hdr = (cols, widths) => row(cols.map((x, i) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), { w: widths[i], fill: NAVY, borders: all(solid(NAVY, 8)) })), 320);
const box = () => p(t("□", { size: 26 }), { after: 0, align: AlignmentType.CENTER });

// Druck
const DW = [550, 3500, 1700, 4456];
const druck = [
  ["Buchstabenrätsel „Generalstände und Abstimmung“ (Seite 1: zwei Rätsel pro A4)", "7 Blatt, in der Mitte schneiden (14 Rätsel, 12 + 2 Reserve)", "Buchstabenraetsel_Generalstaende_8C.pdf, Seite 2 = Lösung für dich, vorab auf die Tische legen"],
  ["Arbeitsblatt „Der Dritte Stand handelt“ (Text, Quelle von Sieyès, Zeitstrahl)", "14 (12 + 2 Reserve), einseitig A4", "Arbeitsblatt_Der_Dritte_Stand_handelt.pdf, vorab auf die Tische legen"],
  ["Schnipsel „Abschlussfrage: Erkläre es deinem Freund“ (drei pro A4)", "5 Blatt, in 3 Schnipsel schneiden (14 Schnipsel, 12 + 2 Reserve)", "Schnipsel_Abschlussfrage_Rat_an_den_Freund.pdf, erst in Phase 4 austeilen"],
  ["Schnipsel „Karikatur analysieren: Dreischritt“ (Seite 1)", "1 Blatt, 6 Schnipsel ausschneiden, je einer pro Paar", "Schnipsel_Karikatur_Dreischritt_und_Hilfekarten.pdf, Seite 1"],
  ["Hilfekarten Stufe 1 bis 3 (Seite 2)", "1 Blatt, ausschneiden, am Pult bereitlegen", "Gleiche Datei, Seite 2. Hilfe nur auf Nachfrage"],
];
// Vorbereitung
const prep = [
  ["Tafelskript ausdrucken oder auf dem Tablet bereithalten", "Tafelskript_Stunde2_Nationalversammlung.pdf, 7 Seiten"],
  ["Folien öffnen und einmal durchklicken", "Praesentation_Stunde2_Nationalversammlung.pptx, 17 Folien (27 Punkt); Folie 1 steht beim Betreten da"],
  ["Presenter (Fernbedienung) für die Wiederholung bereithalten", "Phase 1 (Folie 5 und 6): Schüler zeigen und erklären die Rätselwörter, sie führen am Presenter und bestimmen selbst den Nächsten. Vorher eine zuverlässige Person für den Anfang ansprechen"],
  ["Lehrbuch bereitlegen", "S. 111 („Für die Mühsal geboren“), S. 113 (Ballhausschwur), S. 116 sind auf den Folien; Buch nur zur Hand"],
  ["Kreide und Grün zum Verbessern bereitlegen", "Kreide-Kette beim Zeitstrahl nur, wenn Zeit ist"],
  ["Uhr oder Timer bereitstellen", "Vorwissen 8 Minuten; Übungsphase endet 5 Minuten vor Schluss, „Auswerten“ ist fest"],
  ["Absprache mit der hospitierenden Lehrkraft", "Sie hospitiert nur: Beobachtungsauftrag z. B. Zeitplan der Phase 1 und Erklärungen der Schüler am Presenter"],
];
// Tafel
const tafel = [
  ["Seitentafel", "Ziele („Ich kann …“, 3 Sätze, Satz 3 mit *) und Ablauf: A Ankommen · V Rätsel zur Wiederholung · I Text und Zeitstrahl · V Karikatur (Schritt 3 *) · A Berater-Frage"],
  ["Mitte", "leer lassen (Platz für Zeitstrahl und Ergebnisse der Karikatur-Analyse)"],
];

const d = [];
d.push(p(t("GESCHICHTE · KLASSE 8C · FREITAG 06.11.2026", { size: 18, bold: true, color: MUTED }), { after: 20 }));
d.push(p(t("Materialliste: Der Dritte Stand handelt", { font: "Cambria", size: 44, bold: true, color: NAVY }), { after: 40 }));
d.push(p(t("Zum Abhaken. Alle Dateien liegen in Geschichte_8/Stunde2_Nationalversammlung/. Ablauf: 3 Ankommen · 8 Rätsel zur Wiederholung · 12 Text und Zeitstrahl · 15 Karikatur · 6 Auswerten · 1 Puffer.", { size: 20, color: MUTED }), { after: 80 }));

d.push(h2("1 Drucken und schneiden"));
d.push(table(DW, [hdr(["", "Was", "Menge", "Datei und Hinweis"], DW), ...druck.map(([a, b, c]) => row([cell(box(), { w: DW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: DW[1] }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: DW[2] }), cell(p(t(c, { size: 19 }), { after: 0 }), { w: DW[3] })], 560))]));

d.push(h2("2 Vorbereiten"));
const PW = [550, 4100, 5556];
d.push(table(PW, [hdr(["", "Aufgabe", "Hinweis"], PW), ...prep.map(([a, b]) => row([cell(box(), { w: PW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: PW[1] }), cell(p(t(b, { size: 19 }), { after: 0 }), { w: PW[2] })], 520))]));

d.push(h2("3 Tafel vor der Stunde"));
const TW = [2000, 8206];
d.push(table(TW, [hdr(["Wo", "Was steht dort"], TW), ...tafel.map(([a, b]) => row([cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: TW[0], fill: LIGHT }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: TW[1] })], 520))]));

d.push(h2("4 Kurz vor Beginn (5 Minuten)"));
["Folie 1 steht (Eingangsfrage: Darf sich das Volk gegen einen König wehren, der allein regiert?), Beamer an", "Arbeitsblätter und Schnipsel liegen auf den Tischen (Blätter umgedreht lassen)", "Seitentafel beschrieben, Mitte leer", "Presenter funktioniert, Person für den Anfang von Phase 1 angesprochen", "Ansagen im Kopf: „Stifte liegen. Ihr hört nur zu.“ · „Flüsterstufe. Nur dein Partner hört dich.“"].forEach((x) => d.push(p([t("□  ", { size: 24 }), t(x, { size: 20 })], { after: 50 })));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 850, right: 850 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Materialliste_Stunde2_Nationalversammlung_8C.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
