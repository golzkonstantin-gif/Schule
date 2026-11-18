// Tafelskript „Vor dem Gesetz“ – Macht und Ohnmacht (Neustart, 45 Min.): Tafelbild, Planungskopf und Ablauf mit Ansagen, Erwartungen, Reflexion
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation } = require("docx");
const D = require("./daten_macht");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA";
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA", DIM = "C9C3A8", YEL = "F5D76E", APP = "F29A8E", SKY = "9CC3F0";
const W = 15338;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 50, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: 110, right: 110 }, columnSpan: o.span });
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const kicker = (text, pb) => p(t(text, { size: 17, bold: true, color: MUTED }), { after: 20, pb });
const h1 = (text) => p(t(text, { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 70 });
const h2 = (text) => p(t(text, { font: "Cambria", size: 23, bold: true, color: NAVY }), { before: 140, after: 50, keepNext: true });
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 17 }), { after: 0 }), { w, fill: NAVY });

// ===== Seite 1: Tafelbild =====
const c = (text, o = {}) => t(text, Object.assign({ color: CHALK, size: 20 }, o));
const line = solid("6E8C7A");
const bc = (children, o = {}) => cell(children, Object.assign({ fill: BOARD, borders: { top: line, bottom: line, left: none, right: none }, m: 70 }, o));
const hW = [W / 2, W / 2];
const board = table(hW, [
  row([bc([
    p(c("Franz Kafka: „Vor dem Gesetz“ – Macht und Ohnmacht", { font: "Cambria", size: 30, bold: true }), { after: 20 }),
    p(c("Seitentafel: Ziel „" + D.ZIEL + "“ · Leitfrage: " + D.LEITFRAGE, { size: 17, color: DIM, italics: true }), { after: 0 }),
  ], { w: W, span: 2, borders: { top: none, bottom: line, left: none, right: none } })]),
  row([
    bc([p(c("① Türhüter", { bold: true, color: APP, size: 22 }), { after: 20 }), p(c(D.TAFEL_SPALTEN[0][1])), p(c(D.TAFEL_SPALTEN[0][2], { italics: true, color: DIM })), p(c("→ MACHT", { bold: true, color: YEL, size: 24 }), { after: 0 })], { w: hW[0] }),
    bc([p(c("Mann vom Lande", { bold: true, color: SKY, size: 22 }), { after: 20 }), p(c(D.TAFEL_SPALTEN[1][1])), p(c(D.TAFEL_SPALTEN[1][2], { italics: true, color: DIM })), p(c("→ OHNMACHT", { bold: true, color: YEL, size: 24 }), { after: 0 })], { w: hW[1] }),
  ], 1400),
  row([bc([
    p(c("② Deutungshypothese", { bold: true, color: YEL, size: 22 }), { after: 20 }),
    p(c(D.EINLEITUNG.hypothese, { italics: true, size: 19 }), { after: 0 }),
  ], { w: W, span: 2 })], 800),
  row([
    bc([p(c("③ Bildebene (Rolle)", { bold: true, color: YEL, size: 21 }), { after: 20 }), ...D.BILD_SACH.slice(0, 2).map(([el, bild]) => p(c(el + ": " + bild.split(";")[0], { size: 18 }), { after: 10 }))], { w: hW[0], fill: BOARD2 }),
    bc([p(c("Sachebene (steht für)", { bold: true, color: YEL, size: 21 }), { after: 20 }), ...D.BILD_SACH.slice(0, 2).map(([el, , sach]) => p(c(el + ": " + sach, { size: 18 }), { after: 10 }))], { w: hW[1], fill: BOARD2 }),
  ], 900),
  row([bc([
    p(c("④ Stilmittel stützen die Hypothese: ", { bold: true, color: YEL, size: 21 }), { after: 10 }),
    p(c("Macht: Wiederholung („mächtig“), Vergleich („wie … große Herren“) · Ohnmacht: Symbol (Schemel) · beides: Metapher („Eintritt in das Gesetz“)", { size: 19 }), { after: 0 }),
  ], { w: W, span: 2, borders: { top: line, bottom: none, left: none, right: none } })], 650),
]);
const page1 = [
  p([t("TAFELBILD IM ENDSTAND · ", { size: 17, bold: true, color: MUTED }), t("① nach dem Markieren (Überschriften Macht/Ohnmacht zuletzt, aus der Klasse) – bleibt die ganze Stunde stehen · ② Einleitung · ③ Aufgabe 3 · ④ Aufgabe 4", { size: 17, color: MUTED })], { after: 60 }),
  board,
];

// ===== Seite 2: Planungskopf und Ablauf =====
const kopfRows = [
  ["Ziel (Ich kann …)", D.ZIEL],
  ["Prüfaufgabe", "Zielcheck im Heft: " + D.ZIELCHECK.auftrag],
  ["Bezug zur letzten Stunde", "Die Stunde am 12.11. ging nicht gut: Probleme mit dem Textverständnis. Deshalb Neustart direkt im Text: Die Klasse markiert, wer was und wie sagt – das reaktiviert die Handlung und zeigt Macht und Ohnmacht, bevor gedeutet wird. Aufbau wie das Klausurvorbereitungsblatt der Mentorin (Einleitung → Bild-/Sachebene → Stilmittel)."],
  ["Mindestziel", "Markieren und Auswertung, Einleitung mit Hypothese, Bild- und Sachebene, Zielcheck. Bei Zeitnot: Stilmittel nur Zeile 1 (vorgemacht) und Zeile 4 (Metapher), Rest Einstieg der nächsten Stunde."],
];
const kW = [2800, W - 2800];
const aW = [900, 2300, 7500, 2338, 2300];
const ablauf = [
  ["0–2'", "Ziel, Ablauf", "direkt", "Folie 1–2. Arbeitsblatt austeilen. „Ihr kennt den Text schon. Heute geht es direkt hinein.“", "Seitentafel", "Folie 1–2"],
  ["2–7'", "Aufgabe 1: Markieren", "allein", "Folie 3: „Jetzt arbeitet ihr allein. Zwei Farben: eine für den Türhüter, eine für den Mann. Einkreisen, wie sie es sagen.“ 5 Min. Herumgehen.", "–", "Folie 3, AB S. 1"],
  ["7–11'", "Auswertung", "Unterrichtsgespräch", "Folie 4: „Vergleicht und verbessert mit Grün.“ Folie 5: drei Fragen, je 2–3 Stichproben (Farbe – Wörter – Redeform). Zwei Spalten an die Tafel, die Überschriften Macht / Ohnmacht zuletzt aus der Klasse. Folie 6 als Sicherung.", "①", "Folie 4–6"],
  ["11–20'", "Aufgabe 2: Einleitung", "gemeinsam – allein – gemeinsam", "Folie 7: TATTE-Bausteine gemeinsam (Folie 8), daraus gemeinsam den TATTE-Satz (Folie 9). Inhalt allein: „Erzählt den Ablauf anhand eurer Farben. 1–3 Sätze, ein Satz genügt.“ (Folie 10). Hypothese: „Was will der Mann? Was wird thematisiert?“ – Antwort steht an der Tafel. Lehrkraft beginnt den Satz, Klasse vervollständigt (Folie 11).", "②", "Folie 7–11, AB S. 2"],
  ["20–25'", "Spiel", "zwei spielen, alle beobachten", "Folie 12: zwei Freiwillige mit der Spielvorlage; Türhüter steht, Mann sitzt seitwärts. Danach mündlich je Merkmal eine Stichprobe (Status, Absicht, Wie spricht er?, Verhalten am Ende). Brücke: „So sieht es auf der Bildebene aus. Wofür stehen die beiden?“", "–", "Folie 12, Spielvorlage"],
  ["25–32'", "Aufgabe 3: Bild- und Sachebene", "allein, dann Besprechung", "Folie 13: Merksatz (Figuren haben eine Rolle und stehen für etwas; Symbole nur Gegenstände). Folie 14: „Jetzt arbeitet ihr allein. Die Bildebene steht schon da – ihr ergänzt die Sachebene.“ 4 Min. Folie 15–16: aufdecken, mit Grün verbessern. Folie 17: Hypothese stimmt – Ergebnis notieren.", "③", "Folie 13–17, AB S. 2"],
  ["32–41'", "Aufgabe 4: Stilmittel", "Modell, dann allein", "Folie 18: Wiederholung (Vergleich mit „wie“, Metapher ohne „wie“). Folie 19–20: Zeile 1 laut denkend vormachen, alle schreiben mit. Zeilen 2–4 allein, 4 Min. Folie 21–22: Lösung, Folie 23: Ergebnis notieren.", "④", "Folie 18–23, AB S. 3"],
  ["41–45'", "Zielcheck", "allein, dann Kontrolle", "Folie 24: „Allein, im Heft. 3 Minuten.“ Folie 25: Musterlösung, mit Grün kontrollieren. Rückblick auf das Ziel. Nie streichen.", "–", "Folie 24–25"],
];
const page2 = [
  kicker("TAFELSKRIPT · DEUTSCH 11 · „VOR DEM GESETZ“ – NEUSTART · 45 MIN. · BEAMER · PLANUNGSMODELL AVIVA", true),
  h1("Planungskopf und Ablauf"),
  table(kW, kopfRows.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: kW[0], fill: LIGHT }), cell(p(t(v, { size: 18 }), { after: 0 }), { w: kW[1] })]))),
  p(t(""), { after: 60 }),
  table(aW, [
    row([hdr("Zeit", aW[0]), hdr("Phase · Form", aW[1]), hdr("Wörtliche Ansagen / Impulse / was passiert", aW[2]), hdr("Tafel", aW[3]), hdr("Material", aW[4])], 340),
    ...ablauf.map(([z, ph, form, imp, tafel, mat]) => row([
      cell(p(t(z, { bold: true, color: NAVY, size: 17 }), { after: 0 }), { w: aW[0], fill: LIGHT }),
      cell([p(t(ph, { bold: true, size: 17, color: NAVY }), { after: 10 }), p(t(form, { size: 15, italics: true, color: MUTED }), { after: 0 })], { w: aW[1], fill: LIGHT }),
      cell(p(t(imp, { size: 17 }), { after: 0 }), { w: aW[2] }),
      cell(p(t(tafel, { size: 17 }), { after: 0 }), { w: aW[3] }),
      cell(p(t(mat, { size: 16, color: MUTED }), { after: 0 }), { w: aW[4] }),
    ])),
  ]),
];

// ===== Seite 3: Erwartungen (ausführlich: Lösungsblatt) und Reflexion =====
const rW = [3000, W - 3000];
const erw = [
  ["Auswertung Aufgabe 1", D.AUSWERTUNG.map(([q, a]) => q + " → " + a).join(" · ") + " Zusatz: " + D.AUSWERTUNG_ZUSATZ.join(" → ")],
  ["Spiel (mündlich)", D.SPIEL_FRAGEN.map(([k, v]) => k + ": " + v).join(" · ")],
  ["Aufgaben 2–4", "siehe Arbeitsblatt_Macht_Ohnmacht_Loesung (Einleitung, Bild- und Sachebene, Stilmittel mit Wirkung und Bezug)."],
  ["Zielcheck", D.ZIELCHECK.muster + " Kriterien: " + D.ZIELCHECK.kriterien.join(" · ") + "."],
];
const refl = [
  ["AVIVA", "Kein eigener Einstieg und kein Vorlesen (Entscheidung der Lehrkraft): Die Klasse kennt den Text und wird durch das Markieren sofort aktiviert – Ankommen und Vorwissen fallen in Aufgabe 1 zusammen. Informieren: Merksatz Bild/Sach/Symbol und Wiederholung der Stilmittel, jeweils direkt vor der Anwendung. Verarbeiten: Einleitung, Spiel, Bild-/Sachebene, Stilmittel. Auswerten: Ergebnissätze und Zielcheck."],
  ["Direkte Instruktion", "Ziel und Ablauf zu Beginn. Erst das Werkzeug, dann die Anwendung (Merksatz vor Aufgabe 3, Wiederholung der Stilmittel vor Aufgabe 4). Ich – Wir – Ihr: Hypothese gemeinsam, Stilmittel-Zeile 1 vorgemacht. Rückmeldung: Abgleich mit Grün nach jeder Aufgabe, Zielcheck mit Kriterien."],
  ["Rückmeldungen umgesetzt", "Textverständnis zuerst (12.11.). Deutungshypothese fragt, was der Protagonist will und was thematisiert wird. Aufbau wie das AB der Mentorin. Spiel kriteriengeleitet, aber mündlich ausgewertet. Figuren sind keine Symbole. Vergleich mit „wie“, Metapher ohne „wie“. Inhalt in 1–3 Sätzen. Erarbeitung allein, 27 pt, keine Hausaufgabe."],
  ["Differenzierung", "Das Tafelbild aus Aufgabe 1 trägt die ganze Stunde (Verben → Rollen → Hypothese). Gerüst für die Hypothese, vorgemachte Zeile bei den Stilmitteln, Merksatz als Hilfe. Freiwillige für das Spiel."],
  ["Beobachtungsschwerpunkte", "Gelingt das Markieren ohne Vorlesen? Kommt die Klasse selbst auf Macht und Ohnmacht? Reicht ein Satz für den Inhalt? Unterscheiden die Schüler Rolle (Bildebene) und Übertragung (Sachebene)? Wie viele erkennen im Zielcheck den Pelzmantel als Symbol?"],
];
const page3 = [
  kicker("TAFELSKRIPT · ERWARTUNGEN UND REFLEXION", true),
  h1("Erwartungen"),
  table(rW, erw.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: rW[0], fill: LIGHT }), cell(p(t(v, { size: 18 }), { after: 0, line: 260 }), { w: rW[1] })]))),
  p(t(""), { after: 60 }),
  h1("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"),
  table(rW, refl.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: rW[0], fill: LIGHT }), cell(p(t(v, { size: 18 }), { after: 0, line: 260 }), { w: rW[1] })]))),
  p(t("Text aus dem Gedächtnis wiedergegeben (heutige Rechtschreibung) – vor dem Druck mit einer Ausgabe abgleichen.", { size: 16, italics: true, color: MUTED }), { before: 80 }),
];
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Kafka, „Vor dem Gesetz“ – Macht und Ohnmacht · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 19 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } } }, footers: { default: footer }, children: [...page1, ...page2, ...page3] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Macht_Ohnmacht/Tafelskript_Macht_Ohnmacht.docx", b));
