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
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Materialliste Nationalversammlung · Mi 11.11.2026 · Seite ", PageNumber.CURRENT] })] })] });

const hdr = (cols, widths) => row(cols.map((x, i) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), { w: widths[i], fill: NAVY, borders: all(solid(NAVY, 8)) })), 320);
const box = () => p(t("□", { size: 26 }), { after: 0, align: AlignmentType.CENTER });

// Druck
const DW = [550, 3500, 1700, 4456];
const druck = [
  ["Druckvorlage Karten, Beobachtungsbogen, Rollenkarten: Seite 1 und 2 (Standeskarten)", "je 1 Seite: 24 Karten ausschneiden (3 Klerus, 4 Adel, 17 Dritter Stand, davon 1 Reserve)", "Druckvorlage_Karten_Beobachtungsbogen_8D.pdf, Seite 1 und 2"],
  ["Beobachtungsbogen (Seite 3, zwei pro A4)", "13 Blatt, in der Mitte schneiden (26 Bögen: 25 + 1 Reserve)", "Gleiche Datei, Seite 3, 13-mal drucken (Bogen liegt mit der Karte auf dem Platz)"],
  ["Rollenkarten Wahlleitung und Schriftführung (Seite 4)", "1 Blatt, in der Mitte teilen", "Gleiche Datei, Seite 4; vorab den zwei Personen geben"],
  ["Text mit Zeitstrahl „Der Dritte Stand handelt“", "27 (25 + 2 Reserve), einseitig A4", "Arbeitsblatt_Der_Dritte_Stand_handelt.pdf (unverändert wie 8C)"],
  ["Schnipsel „Abschlussfrage: Erkläre es deinem Freund“ (drei pro A4)", "9 Blatt, in 3 Schnipsel schneiden (27 Schnipsel)", "Schnipsel_Abschlussfrage_Rat_an_den_Freund.pdf, erst in Phase 4 austeilen"],
  ["Lösung Text mit Zeitstrahl (für dich)", "1 Blatt", "Loesung_Arbeitsblatt_Der_Dritte_Stand_handelt.pdf"],
];
// Vorbereitung
const prep = [
  ["Tafelskript ausdrucken oder auf dem Tablet bereithalten", "Tafelskript_Stunde2_Nationalversammlung_8D.pdf, 5 Seiten"],
  ["Folien öffnen und einmal durchklicken", "Praesentation_Stunde2_Nationalversammlung_8D.pptx, 15 Folien (27 Punkt); Folie 1 steht beim Betreten da"],
  ["Rollen verteilen", "Wahlleitung und Schriftführung zuerst ansprechen (zwei zuverlässige, rechensichere Personen, nicht stimmberechtigt), Rollenkarte geben. Klerus 3 und Adel 4 gezielt wählen, Rest Dritter Stand; Tabelle unten ausfüllen (ohne Namen, nur Platznummern)"],
  ["Karten und Bögen auf die Plätze legen", "Karte verdeckt, Bogen daneben; Reservekarten und -bögen am Pult"],
  ["Abstimmung: Sie sind nur Sicherheitsnetz", "Offene Abstimmung per Handzeichen: Die Wahlleitung ruft die Stände auf und zählt die Hände, die Schriftführung schreibt an die Tafel, alle anderen tragen auf dem Bogen ein. Nach 4 Minuten ohne Ergebnis helfen"],
  ["Kreide, Grün zum Verbessern, Uhr oder Timer", "Kurzstunde: 25 Minuten, kein Puffer; Phase 4 „Auswerten“ ist fest"],
];
// Tafel
const tafel = [
  ["Seitentafel", "Ziele („Ich kann …“, 3 Sätze, Satz 3 mit *) und Ablauf: A Ankommen · V Abstimmung · I Text · V Zeitstrahl · A Rat an den Freund"],
  ["Mitte", "Rahmen (Tabelle) für das Ergebnis, den die Schriftführung ausfüllt: Zeilen A und B; Spalten nach Ständen (Klerus, Adel, Dritter Stand, Ergebnis) und nach Köpfen (Ja, Nein, Ergebnis). Daneben Platz für den Zeitstrahl."],
];

const d = [];
d.push(p(t("GESCHICHTE · KLASSE 8D · MITTWOCH 11.11.2026", { size: 18, bold: true, color: MUTED }), { after: 20 }));
d.push(p(t("Materialliste: Nationalversammlung", { font: "Cambria", size: 44, bold: true, color: NAVY }), { after: 40 }));
d.push(p(t("Zum Abhaken. Alle Dateien liegen in Geschichte_8D/Stunde2_Nationalversammlung/. Ablauf (25 Minuten): 3 Ankommen · 9 Abstimmung · 6 Text · 3 Zeitstrahl · 4 Auswerten.", { size: 20, color: MUTED }), { after: 80 }));

d.push(h2("1 Drucken und schneiden"));
d.push(table(DW, [hdr(["", "Was", "Menge", "Datei und Hinweis"], DW), ...druck.map(([a, b, c]) => row([cell(box(), { w: DW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: DW[1] }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: DW[2] }), cell(p(t(c, { size: 19 }), { after: 0 }), { w: DW[3] })], 560))]));

d.push(h2("2 Vorbereiten"));
const PW = [550, 4100, 5556];
d.push(table(PW, [hdr(["", "Aufgabe", "Hinweis"], PW), ...prep.map(([a, b]) => row([cell(box(), { w: PW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: PW[1] }), cell(p(t(b, { size: 19 }), { after: 0 }), { w: PW[2] })], 520))]));

d.push(h2("3 Tafel vor der Stunde"));
const TW = [2000, 8206];
d.push(table(TW, [hdr(["Wo", "Was steht dort"], TW), ...tafel.map(([a, b]) => row([cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: TW[0], fill: LIGHT }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: TW[1] })], 520))]));

d.push(h2("4 Rollenverteilung (ohne Namen, nur Platznummern)"));
const RW = [2400, 1100, 6706];
const rolle = [["Erster Stand · Klerus", "3", "Plätze: ____  ____  ____"], ["Zweiter Stand · Adel", "4", "Plätze: ____  ____  ____  ____"], ["Dritter Stand (Bauern und Bürger)", "16", "alle übrigen Plätze; bei Fehlzeiten Reservekarte einsetzen"], ["Wahlleitung", "1", "Platz: ____  (nicht stimmberechtigt)"], ["Schriftführung", "1", "Platz: ____  (nicht stimmberechtigt)"]];
d.push(table(RW, [hdr(["Stand", "Anzahl", "Plätze"], RW), ...rolle.map(([a, b, c]) => row([cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: RW[0], fill: LIGHT }), cell(p(t(b, { size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: RW[1] }), cell(p(t(c, { size: 20 }), { after: 0 }), { w: RW[2] })], 520))]));
d.push(p(t("Wenn Schülerinnen oder Schüler fehlen: Klerus bleibt 3, Adel bleibt 4, Wahlleitung und Schriftführung bleiben. Klerus und Adel dürfen zusammen nie mehr als ein Drittel der Klasse sein.", { size: 19, italics: true, color: MUTED }), { before: 60 }));

d.push(h2("5 Kurz vor Beginn (5 Minuten)"));
["Folie 1 steht, Beamer an", "Karten liegen verdeckt auf den Plätzen, daneben der Bogen; Wahlleitung und Schriftführung haben ihre Rollenkarte", "Seitentafel, Rahmen für die Abstimmung und Datumsliste an der Tafel", "Textblätter liegen bereit (noch nicht austeilen), Schnipsel erst in Phase 4", "Ansagen im Kopf: „Stifte liegen. Ihr hört nur zu.“ · „Jetzt stimmt ihr offen ab. Die Wahlleitung führt, alle anderen beobachten und schreiben auf.“"].forEach((x) => d.push(p([t("□  ", { size: 24 }), t(x, { size: 20 })], { after: 50 })));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 850, right: 850 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Materialliste_Stunde2_Nationalversammlung_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
