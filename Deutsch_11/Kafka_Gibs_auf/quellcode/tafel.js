// Tafelskript „Gibs auf!“ (Do 5.11., 45 Min., AVIVA): Tafelbild im Endstand, Planungskopf und Ablauf mit Ansagen,
// Lösungen und Heftübersicht, Umsetzung der Erkenntnisse zur Unterrichtsverbesserung
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GREEN = "2E9E6B";
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA", DIM = "C9C3A8", YEL = "F5D76E", APP = "F29A8E";
const W = 15338, H = W / 2;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 50, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: 110, right: 110 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const kicker = (text, pb) => p(t(text, { size: 17, bold: true, color: MUTED }), { after: 20, pb });
const h1 = (text) => p(t(text, { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 70 });
const h2 = (text) => p(t(text, { font: "Cambria", size: 23, bold: true, color: NAVY }), { before: 140, after: 50, keepNext: true });
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none), m: 80 })])]);

// ================= Seite 1: Tafelbild =================
const c = (text, o = {}) => t(text, Object.assign({ color: CHALK, size: 20 }, o));
const line = solid("6E8C7A");
const bc = (children, o = {}) => cell(children, Object.assign({ fill: BOARD, borders: { top: line, bottom: line, left: none, right: none }, m: 80 }, o));
const sW = [2600, 5000, W - 7600];
const board = table(sW, [
  row([bc([
    p(c("Franz Kafka: „Gibs auf!“ – Stilmittel und Parabel", { font: "Cambria", size: 30, bold: true }), { after: 20 }),
    p(c("Seitentafel: Ziel „Ich kann Stilmittel bestimmen und ihre Wirkung erklären.“ · Ablauf: Bild – Wiederholung – neue Stilmittel – eigenes Gespräch – Untersuchen* – Zielcheck", { size: 17, color: DIM, italics: true }), { after: 0 }),
  ], { w: W, span: 3, borders: { top: none, bottom: line, left: none, right: none } })]),
  row([bc([
    p(c("① Parabel", { bold: true, color: YEL, size: 22 }), { after: 30 }),
    p([c("Bildebene: ", { bold: true }), c("Ein Mann sucht in einer fremden Stadt den Weg zum Bahnhof, ist zu spät, fragt einen Schutzmann – der lacht und sagt „Gibs auf“.")]),
    p([c("Sachebene: ", { bold: true }), c("Ein Mensch sucht Orientierung im Leben und fragt eine Instanz, die helfen sollte (Staat, Autorität, Gott, Vater). Sie verweigert die Hilfe → Ohnmacht, Verlorenheit.")]),
  ], { w: W, span: 3 })], 1300),
  row([
    bc(p(c("② Stilmittel", { bold: true, color: YEL })), { w: sW[0], fill: BOARD2 }),
    bc(p(c("Beispiel", { bold: true, color: YEL })), { w: sW[1], fill: BOARD2 }),
    bc(p(c("Wirkung", { bold: true, color: YEL })), { w: sW[2], fill: BOARD2 }),
  ]),
  row([bc(p(c("Parataxe / Hypotaxe", { bold: true })), { w: sW[0] }), bc(p(c("Satz 1 kurz gereiht · Satz 2 erst Nebensätze, dann Reihe ohne Punkt", { italics: true })), { w: sW[1] }), bc(p(c("ruhig → überstürzt: Der Satz verliert die Ordnung wie der Erzähler")), { w: sW[2] })]),
  ...D.STILMITTEL.map(([b, , bsp, , wirk]) => row([
    bc(p(c(b, { bold: true, color: APP })), { w: sW[0] }),
    bc(p(c(bsp, { italics: true })), { w: sW[1] }),
    bc(p(c(wirk.split(".")[0] + ".")), { w: sW[2] }),
  ])),
  row([bc([
    p([c("③* Unser „Gibs auf!“ (Schülertext): ", { bold: true, color: YEL, size: 22 }), c("abweisender Satz → Sachinhalt · Selbstkundgabe · Beziehung · Appell (entsteht live) – Vergleich mit dem Schutzmann", { size: 19 })], { after: 0 }),
  ], { w: W, span: 3, fill: BOARD2, borders: { top: line, bottom: none, left: none, right: none } })], 700),
]);
const page1 = [
  p([t("TAFELBILD IM ENDSTAND · ", { size: 17, bold: true, color: MUTED }), t("① nach dem Vorwissen · ② während Informieren und Aufgabe 1 · ③ in der Auswertung (*)", { size: 17, color: MUTED })], { after: 60 }),
  board,
];

// ================= Seite 2: Planungskopf und Ablauf =================
const kopfRows = [
  ["Ziel (Ich kann …)", "Ich kann in „Gibs auf!“ Stilmittel bestimmen und ihre Wirkung erklären. Ich kann Bild- und Sachebene einer Parabel unterscheiden. (* Ich kann eine eigene Erfahrung als Gespräch aufschreiben und mit dem Kommunikationsquadrat untersuchen.)"],
  ["Prüfaufgabe", "Zielcheck: drei neue Sätze (Wiederholung, Vergleich, rhetorische Frage) bestimmen und die Wirkung in einem Satz erklären."],
  ["Beitrag zur Reihe", "Nach Kafkas „Brief an den Vater“ (Kommunikation) und der Stunde der Mentorin zu „Gibs auf!“: Die sprachliche Gestaltung wird zum Werkzeug der Deutung; der eigene Text verbindet Parabel und Kommunikationsmodell."],
  ["Mindestziel", "Wiederholung, neue Stilmittel, eigenes Gespräch geschrieben, Zielcheck. Mit * : gemeinsame Untersuchung eines Schülertextes (sonst Einstieg der nächsten Stunde)."],
  ["Bezug zur letzten Stunde", "Die Stunde der Mentorin (2.11.) lief gut. Offen geblieben: Die Frage „Parataxe oder Hypotaxe?“ war beim langen Satz nicht eindeutig – sie wird in der Wiederholung geklärt."],
];
const kW = [2800, W - 2800];
const aW = [900, 2300, 7500, 2338, 2300];
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 17 }), { after: 0 }), { w, fill: NAVY });
const ablauf = [
  ["0–3'", "Verlauf, Lernziele", "direkt", "Folie 1–2. Ziel und Ablauf an der Seitentafel, Untersuchung mit *. „Heute bleiben wir bei ‚Gibs auf!‘ – und dann schreibt ihr euer eigenes ‚Gibs auf!‘.“", "Seitentafel", "Folie 1–2"],
  ["3–7'", "A: Ankommen", "indirekt", "Folie 3 (Bildimpuls Wegweiser): „Stifte liegen. Schaut euch das Bild an.“ – „Was seht ihr? Wie fühlt sich jemand, der davorsteht?“ Erwartung: alle Schilder „Bahnhof“, verschiedene Richtungen, Uhr läuft → ratlos, Zeitdruck. Gelenk: „Woran erinnert euch das?“ → „Gibs auf!“", "–", "Folie 3"],
  ["7–15'", "V: Vorwissen aktivieren", "direkt", "Folie 4: „Ihr habt eine Minute. Schreibt Stichworte ins Heft.“ Dann Stichproben (4–6, unterschiedliche Niveaus), Antworten an die Tafel (①). Folie 5: Lehrkraft liest den Text vor, Klasse liest im Hefter mit. Folie 6–7: Satz 1 und 2 – Parataxe oder Hypotaxe? Erst 30 Sek. allein, dann Stichproben, dann Lösung (Satz 2: beides).", "① Parabel", "Folie 4–7, Hefter (Blatt 2.11.)"],
  ["15–20'", "I: Informieren", "direkt", "Folie 8–9: neue Stilmittel in zwei Portionen, je ein Beispiel; Warnung vor dem Fehler: Nicht jede Frage ist rhetorisch. Schnipsel einkleben. Folie 10: Wiederholung laut denkend vormachen (Was steht da? → Stilmittel? → Wirkung?). Tafel: erste Zeile in ②.", "② erste Zeilen", "Folie 8–10, Schnipsel"],
  ["20–26'", "V: Aufgabe 1", "indirekt (angeleitet)", "Folie 11: „Jetzt arbeitet ihr allein. Die Aufgabe steht auf der Folie.“ 5 Min. Herumgehen. Lösungen nacheinander aufdecken (Folie 12–14), erst eine Antwort vom Platz, dann Folie. „Vergleicht und verbessert mit Grün.“ Tafel: ② ergänzen.", "② vervollständigen", "Folie 11–14"],
  ["26–37'", "V: Aufgabe 2 (Schreiben)", "indirekt (selbstständig)", "Folie 15: „Jetzt schreibt ihr allein. Ihr müsst nichts vorlesen, was ihr nicht vorlesen wollt.“ Timer 10 Min. Schreibhilfe auf dem Schnipsel und Folie 16–17. Schnelle: ein zweites Stilmittel einbauen.", "–", "Folie 15–17, Heft"],
  ["37–42'", "A: Auswerten *", "direkt", "Folie 18: Ein freiwilliger Text wird vorgelesen. Alle schreiben den abweisenden Satz und seinen Appell ins Heft, dann gemeinsam das Quadrat an der Tafel (③). Gelenk: „Ist das auch eine Parabel? Was wäre die Sachebene?“ Entfällt bei Zeitnot – dann Einstieg der nächsten Stunde.", "③*", "Folie 18"],
  ["42–45'", "A: Zielcheck", "indirekt, dann Kontrolle", "Folie 19: „Allein, im Heft. 3 Minuten.“ Folie 20: Lösung, mit Grün kontrollieren. Herumgehen. Rückblick auf die Lernziele. Nie streichen.", "–", "Folie 19–20"],
];
const page2 = [
  kicker("TAFELSKRIPT · DEUTSCH 11 · DO 5.11. · 45 MIN. · PLANUNGSMODELL AVIVA", true),
  h1("Planungskopf und Ablauf"),
  table(kW, kopfRows.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: kW[0], fill: LIGHT }), cell(p(t(v, { size: 18 }), { after: 0 }), { w: kW[1] })]))),
  p(t(""), { after: 60 }),
  table(aW, [
    row([hdr("Zeit", aW[0]), hdr("Phase · Warum", aW[1]), hdr("Wörtliche Ansagen / Impulse / was passiert", aW[2]), hdr("Tafel", aW[3]), hdr("Material", aW[4])], 340),
    ...ablauf.map(([z, ph, warum, imp, tafel, mat]) => row([
      cell(p(t(z, { bold: true, color: NAVY, size: 17 }), { after: 0 }), { w: aW[0], fill: LIGHT }),
      cell([p(t(ph, { bold: true, size: 17, color: NAVY }), { after: 10 }), p(t(warum, { size: 15, italics: true, color: MUTED }), { after: 0 })], { w: aW[1], fill: LIGHT }),
      cell(p(t(imp, { size: 17 }), { after: 0 }), { w: aW[2] }),
      cell(p(t(tafel, { size: 17 }), { after: 0 }), { w: aW[3] }),
      cell(p(t(mat, { size: 16, color: MUTED }), { after: 0 }), { w: aW[4] }),
    ])),
  ]),
];

// ================= Seite 3: Lösungen, Text, Heft =================
const lW = [2600, 5200, W - 7800];
const page3 = [
  kicker("TAFELSKRIPT · LÖSUNGEN UND ERWARTUNGEN", true),
  h1("Lösungen"),
  h2("Wiederholung (Folie 4)"),
  p(t("Parabel: Bildebene (was konkret erzählt wird) → Übertragung auf eine Sachebene (allgemeine menschliche Erfahrung). Schutzmann: eine Instanz, die Orientierung geben sollte – Staat, Autorität, Gott; möglicher Bezug zum Vater aus dem „Brief an den Vater“. Bekannte Stilmittel: Parataxe, Hypotaxe, Metapher, Symbol (Uhr = Zeitdruck, Weg = Lebensweg).", { size: 18 }), { line: 280 }),
  h2("Satzbau (Folie 6–7)"),
  p([t("Satz 1: ", { bold: true, size: 18 }), t(D.SATZBAU.s1.loesung, { size: 18 })], { line: 280 }),
  p([t("Satz 2: ", { bold: true, size: 18 }), t(D.SATZBAU.s2.loesung, { size: 18 })], { line: 280 }),
  h2("Aufgabe 1 und Zielcheck"),
  table(lW, [
    row([hdr("Stilmittel", lW[0]), hdr("Zitat / Satz", lW[1]), hdr("Wirkung", lW[2])]),
    ...D.STILMITTEL.map(([b, , bsp, , wirk]) => row([cell(p(t(b, { bold: true, size: 18, color: RED }), { after: 0 }), { w: lW[0], fill: LIGHT }), cell(p(t(bsp, { italics: true, size: 18 }), { after: 0 }), { w: lW[1] }), cell(p(t(wirk, { size: 18 }), { after: 0 }), { w: lW[2] })])),
    ...D.ZIELCHECK.map(([satz, sm, wirk], i) => row([cell(p(t("Zielcheck " + (i + 1) + ": " + sm, { bold: true, size: 18, color: GREEN }), { after: 0 }), { w: lW[0], fill: LIGHT }), cell(p(t(satz, { italics: true, size: 18 }), { after: 0 }), { w: lW[1] }), cell(p(t(wirk, { size: 18 }), { after: 0 }), { w: lW[2] })])),
  ]),
  h2("Aufgabe 2: Erwartung an die Schülertexte"),
  p(t("6–8 Sätze, Ich-Form, Präteritum; eine Bitte um Hilfe oder Orientierung und eine Abweisung in wörtlicher Rede (Zeichensetzung!); mindestens ein Stilmittel von heute. Nicht bewertet. Persönliches nur freiwillig vorlesen; erfundene Situationen sind erlaubt.", { size: 18 }), { line: 280 }),
  h2("Hinweis zum Text"),
  box([p(t(D.ORIGINAL_HINWEIS, { size: 17 }), { after: 0 })], "FFF8EC"),
  h2("So sieht das Heft am Ende aus"),
  p(t("Überschrift „Kafka: ‚Gibs auf!‘ – Stilmittel“ · Stichworte zur Wiederholung · eingeklebter Merkkasten · Aufgabe 1 (drei Zitate + Wirkung, mit Grün verbessert) · eigenes „Gibs auf!“ (6–8 Sätze) · *Appell des Schülertextes · Zielcheck (mit Grün kontrolliert).", { size: 18 }), { line: 280 }),
];

// ================= Seite 4: Umsetzung der Erkenntnisse =================
const rW = [3000, W - 3000];
const refl = [
  ["AVIVA", "Alle fünf Phasen, das Auswerten (Zielcheck) wird nie gestrichen; der Stern trifft nur die Untersuchung des Schülertextes. Ankommen über einen Bildimpuls, der schon die Parabel-Struktur zeigt (Bild → Bedeutung)."],
  ["Direkte Instruktion", "Ziel und Ablauf zu Beginn (Seitentafel, Folie). Vorwissen aktiv abgeprüft (erst allein, dann Stichproben) – hat sich in der Kafka-Appell-Stunde bewährt. Informieren in zwei Portionen mit Fehlerwarnung. Ich – Wir – Ihr: Lehrkraft macht die Wiederholung laut denkend vor, die Klasse wendet es bei drei Stilmitteln an, dann selbstständig im eigenen Text. Rückmeldung: Lösungen nacheinander aufdecken, mit Grün verbessern; Zielcheck."],
  ["Rückmeldungen umgesetzt", "Aufgaben ohne Dopplungen: finden (Aufgabe 1) → selbst verwenden (Aufgabe 2) → ein Gespräch untersuchen (Aufgabe 3*). Erarbeitung allein, keine Partnerarbeit. Präsentation in 27 pt mit Verlauf und Lernzielen. Schreibhilfe nur mit inhaltsunabhängigen Satzmustern. Sportbezug nur beim eigenen Schreiben, wo er passt. Keine Hausaufgabe."],
  ["Differenzierung", "Schreibhilfe als gestufte Hilfe (Zeichensetzung, Stilmittel); Schnelle bauen ein zweites Stilmittel ein; erfundene statt persönlicher Situation erlaubt; Vorlesen freiwillig."],
  ["Beobachtungsschwerpunkte", "Klappt die Klärung Parataxe/Hypotaxe ohne Verwirrung? Erkennen die Schüler, dass der Erzähler die rhetorische Frage wörtlich nimmt? Reichen 10 Minuten für den eigenen Text? Wie viele schaffen den Zielcheck fehlerfrei (→ Wiederholung nächste Stunde)? Reicht die Zeit für die Untersuchung (*)?"],
];
const page4 = [
  kicker("TAFELSKRIPT · REFLEXION", true),
  h1("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"),
  table(rW, refl.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 19 }), { after: 0 }), { w: rW[0], fill: LIGHT }), cell(p(t(v, { size: 19 }), { after: 0, line: 280 }), { w: rW[1] })]))),
];

const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Kafka, „Gibs auf!“ · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 19 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } } }, footers: { default: footer }, children: [...page1, ...page2, ...page3, ...page4] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelskript_Gibs_auf.docx", b));
