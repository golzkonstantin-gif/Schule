// Tafelskript „Vor dem Gesetz“ (Do 12.11., 45 Min., AVIVA): Tafelbild im Endstand, Planungskopf und Ablauf mit Ansagen,
// Lösungen und Heftübersicht, Umsetzung der Erkenntnisse zur Unterrichtsverbesserung
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GREEN = "2E9E6B";
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA", DIM = "C9C3A8", YEL = "F5D76E", APP = "F29A8E";
const W = 15338;

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
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 17 }), { after: 0 }), { w, fill: NAVY });

// ================= Seite 1: Tafelbild =================
const c = (text, o = {}) => t(text, Object.assign({ color: CHALK, size: 20 }, o));
const line = solid("6E8C7A");
const bc = (children, o = {}) => cell(children, Object.assign({ fill: BOARD, borders: { top: line, bottom: line, left: none, right: none }, m: 70 }, o));
// Kurzfassungen für die Tafel
const TAFEL_SACH = [
  "Zugang scheint möglich – bleibt unerreichbar",
  "Autorität / Behörde; Macht durch unsichtbare Hierarchie",
  "Recht, Gerechtigkeit (offen: Wahrheit, Sinn, Gott, Vater)",
  "einfacher Mensch, kennt die Regeln nicht – jeder",
  "Leben vergeht; Gehorsam statt Handeln",
  "Macht ist nicht beeinflussbar, Schein von Hoffnung",
  "Erniedrigung, Ziel aus dem Blick",
  "Wahrheit / Hoffnung – zu spät erkannt",
  "die Chance war persönlich – vertan",
];
const bW = [3700, 1000, W - 4700];
const board = table(bW, [
  row([bc([
    p(c("Franz Kafka: „Vor dem Gesetz“ – eine kafkaeske Parabel", { font: "Cambria", size: 30, bold: true }), { after: 20 }),
    p(c("Seitentafel: Ziel „Ich kann die Symbole der Parabel deuten und erklären, warum sie kafkaesk ist.“ · Ablauf: Zitat – kafkaesk – Lesen – Handlungsschritte – Symbole – Deutung – Transfer* – Zielcheck", { size: 17, color: DIM, italics: true }), { after: 0 }),
  ], { w: W, span: 3, borders: { top: none, bottom: line, left: none, right: none } })]),
  row([bc([
    p([c("① kafkaesk", { bold: true, color: YEL, size: 22 }), c("   Zitat: „Mich brechen alle Hindernisse.“ → Ohnmacht · Resignation", { size: 19 })], { after: 30 }),
    p([c("Antworten der Klasse → ", { bold: true }), c("übermächtige Instanz · undurchschaubare Regeln · Ohnmacht → man fügt sich · nüchtern erzählt, wirkt wie ein Albtraum")]),
    p([c("Kafka ohnmächtig vor: ", { bold: true }), c("dem Vater (Brief) · dem Schutzmann („Gibs auf!“) · Behörde und Gericht („Der Process“)")], { after: 0 }),
  ], { w: W, span: 3 })], 1100),
  row([
    bc(p(c("② Bildebene", { bold: true, color: YEL })), { w: bW[0], fill: BOARD2 }),
    bc(p(c("Z.", { bold: true, color: YEL })), { w: bW[1], fill: BOARD2 }),
    bc(p(c("Sachebene", { bold: true, color: YEL })), { w: bW[2], fill: BOARD2 }),
  ]),
  ...D.SYMBOLE.map(([b, , anker, , stern], i) => row([
    bc(p(c(b + (stern ? " *" : ""), { bold: true, color: stern ? DIM : APP, size: 19 }), { after: 0 }), { w: bW[0], m: 30 }),
    bc(p(c(D.zeile(anker).replace("Z. ", ""), { size: 18 }), { after: 0 }), { w: bW[1], m: 30 }),
    bc(p(c(TAFEL_SACH[i], { size: 19 }), { after: 0 }), { w: bW[2], m: 30 }),
  ])),
  row([bc([
    p([c("③ Warum kommt er nicht hinein?  ", { bold: true, color: YEL, size: 22 }), c("von außen: Türhüter + Hierarchie  ·  von innen: nur „jetzt“ verboten – er wartet und fügt sich", { size: 19 })], { after: 20 }),
    p(c("→ Ohnmacht wird zur Resignation: „Mich brechen alle Hindernisse.“", { size: 19, italics: true }), { after: 0 }),
  ], { w: W, span: 3, fill: BOARD2, borders: { top: line, bottom: none, left: none, right: none } })], 650),
]);
const page1 = [
  p([t("TAFELBILD IM ENDSTAND · ", { size: 17, bold: true, color: MUTED }), t("① im Ankommen und Vorwissen · ② bei der Besprechung von Aufgabe 2 (erste Zeile beim Vormachen) · ③ in der Deutung", { size: 17, color: MUTED })], { after: 60 }),
  board,
];

// ================= Seite 2: Planungskopf und Ablauf =================
const kopfRows = [
  ["Ziel (Ich kann …)", D.LERNZIELE.map(([z]) => z).join(" ")],
  ["Prüfaufgabe", "Zielcheck im Heft: " + D.ZIELCHECK.auftrag],
  ["Beitrag zur Reihe", "Nach „Gibs auf!“ (Schutzmann als Instanz, die Orientierung verweigert) die zweite Parabel: Der Türhüter ist eine noch mächtigere Instanz. „kafkaesk“ verbindet Kafkas Biografie (Vater, Behörde), den „Process“ und die Parabeln."],
  ["Mindestziel", "Wiederholung „kafkaesk“, Text gelesen, Handlungsschritte, Symbole Tor, Türhüter, Gesetz gedeutet, Zielcheck. Mit *: Flöhe, Glanz, Schließen des Eingangs; Transfer (sonst Einstieg der nächsten Stunde)."],
  ["Bezug zur letzten Stunde", "„Gibs auf!“ (5.11.) lief sehr gut, mit gutem Transfer und guter Reflexion. Darauf baut die Stunde auf: Bild- und Sachebene sind sicher, der Schutzmann wird zum Türhüter. Wiederholung von „kafkaesk“ als Brücke."],
];
const kW = [2800, W - 2800];
const aW = [900, 2300, 7500, 2338, 2300];
const ablauf = [
  ["0–4'", "A: Ankommen", "indirekt (Impuls)", "Folie 1 (Zitat): „Stifte liegen. Lest den Satz still.“ – „" + D.ZITAT.impuls + "“ Nachfrage: „Was ist der Unterschied zwischen ‚Ich breche‘ und ‚Mich brechen‘?“ Stichworte Ohnmacht · Resignation an die Tafel. Gelenk: „Für dieses Gefühl gibt es ein eigenes Wort. Ihr kennt es schon.“", "① Zitat", "Folie 1"],
  ["4–6'", "Verlauf, Lernziele", "direkt", "Folie 2–3. Ziel und Ablauf an der Seitentafel, Transfer und Symbole mit *.", "Seitentafel", "Folie 2–3"],
  ["6–13'", "V: Vorwissen aktivieren", "direkt", "Folie 4: „Ihr habt eine Minute. Schreibt Stichworte ins Heft.“ Dann Stichproben (4–6, unterschiedliche Niveaus), Antworten an die Tafel (①). Folie 5: Definition aufdecken, mit der Tafel abgleichen. Folie 6: Kafka und die Macht kurz erzählen; Process-Satz: „Was ist daran kafkaesk?“", "① kafkaesk", "Folie 4–6"],
  ["13–20'", "I: Informieren", "direkt", "Folie 7: Kontext. Arbeitsblatt austeilen. „Stifte liegen. Ihr lest auf dem Arbeitsblatt mit.“ Lehrkraft liest vor (Modell-Lesen, ca. 4 Min.). Leseauftrag: „" + D.LESEAUFTRAG + "“ Folie 8: „Worum geht es – in einem Satz?“ 2–3 Schüler, dann Kernbotschaft aufdecken.", "–", "Folie 7–8, AB Nr. 1"],
  ["20–24'", "V: Aufgabe 1", "indirekt (angeleitet)", "Folie 9: „Jetzt arbeitet ihr allein. Die Aufgabe steht auf dem Arbeitsblatt, Nr. 3.“ 3 Min. Folie 10–11: Lösung, „Vergleicht und verbessert mit Grün.“", "–", "Folie 9–11, AB Aufgabe 1"],
  ["24–26'", "Ich mache vor", "direkt (Modell)", "Folie 12: das offene Tor laut denkend in drei Schritten (Was steht da? → Was fällt auf? → Wofür steht es?). Alle tragen die erste Zeile mit ein. Tafel: erste Zeile in ②.", "② erste Zeile", "Folie 12"],
  ["26–35'", "V: Aufgabe 2", "indirekt (selbstständig)", "Folie 13: „Jetzt arbeitet ihr allein. Die Tabelle steht auf dem Arbeitsblatt, Aufgabe 2.“ Timer 9 Min. Herumgehen; Hilfe nur als Frage: „Was fällt an dem Wort auf?“ Schnelle: Zeilen mit *.", "–", "Folie 13, AB Aufgabe 2"],
  ["35–39'", "V: Besprechung", "Unterrichtsgespräch", "Folie 14–16: je erst eine Antwort vom Platz, dann aufdecken; alle verbessern mit Grün. Gesetz ausdrücklich offen deuten (Recht, Wahrheit, Sinn, Gott, Vater). Folie 16 (*) nur bei Zeit, sonst das Schließen mündlich.", "② vervollständigen", "Folie 14–16"],
  ["39–41'", "A: Deutung (Transfer*)", "Unterrichtsgespräch", "Folie 17: „Warum kommt der Mann nicht hinein?“ Antworten sammeln, dann aufdecken: von außen / von innen; Rückgriff auf das Zitat. Folie 18 (Transfer*) nur, wenn noch mehr als 6 Min. bleiben.", "③", "Folie 17–18"],
  ["41–45'", "A: Zielcheck", "indirekt, dann Kontrolle", "Folie 19: „Allein, im Heft. 4 Minuten.“ Herumgehen. Folie 20: Musterlösung und Kriterien, mit Grün kontrollieren. Rückblick auf die Lernziele. Nie streichen.", "–", "Folie 19–20"],
];
const page2 = [
  kicker("TAFELSKRIPT · DEUTSCH 11 · DO 12.11. · 45 MIN. · BEAMER · PLANUNGSMODELL AVIVA", true),
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

// ================= Seite 3: Lösungen =================
const lW = [3000, 1100, W - 4100];
const page3 = [
  kicker("TAFELSKRIPT · LÖSUNGEN UND ERWARTUNGEN", true),
  h1("Lösungen"),
  h2("Zitat (Folie 1) und Vorwissen (Folie 4–6)"),
  p(t(D.ZITAT.erwartung + " – „kafkaesk“: " + D.KAFKAESK.merkmale.join(" ") + " Process-Satz: Josef K. hat nichts getan und erfährt keinen Grund – die Staatsmacht ist übermächtig und undurchschaubar.", { size: 18 }), { line: 270 }),
  h2("Kernbotschaft (Folie 8) und Aufgabe 1"),
  p(t(D.KERNBOTSCHAFT, { size: 18, italics: true }), { line: 270 }),
  p(t("Reihenfolge: " + D.SCHRITTE.map((s, i) => (i + 1) + " = " + String.fromCharCode(65 + D.REIHE_GEMISCHT.indexOf(i)) + " (" + D.zeile(s[1]) + ")").join(" · ") + ". Zeilen ± 1 gelten als richtig.", { size: 18 }), { line: 270 }),
  h2("Aufgabe 2: Symbole"),
  table(lW, [
    row([hdr("Bildebene", lW[0]), hdr("Zeile", lW[1]), hdr("Sachebene (Erwartung)", lW[2])]),
    ...D.SYMBOLE.map(([b, achte, anker, sach, stern]) => row([
      cell([p(t(b + (stern ? " *" : ""), { bold: true, size: 18, color: RED }), { after: 0 }), p(t(achte, { size: 16, italics: true, color: MUTED }), { after: 0 })], { w: lW[0], fill: LIGHT }),
      cell(p(t(D.zeile(anker), { size: 18 }), { after: 0 }), { w: lW[1] }),
      cell(p(t(sach, { size: 18 }), { after: 0 }), { w: lW[2] }),
    ])),
  ]),
  h2("Deutung (Folie 17) und Zielcheck"),
  p(t(D.DEUTUNG.aussen + " " + D.DEUTUNG.innen + " " + D.DEUTUNG.zitat, { size: 18 }), { line: 270 }),
  p([t("Musterlösung Zielcheck: ", { bold: true, size: 18, color: GREEN }), t(D.ZIELCHECK.muster, { size: 18 }), t("  Kriterien: " + D.ZIELCHECK.kriterien.join(" · ") + ".", { size: 18, color: MUTED })], { line: 270 }),
];

// ================= Seite 4: Umsetzung der Erkenntnisse =================
const rW = [3000, W - 3000];
const refl = [
  ["AVIVA", "Alle fünf Phasen; das Auswerten (Zielcheck) wird nie gestrichen, der Stern trifft nur drei Symbole und den Transfer. Ankommen mit einem Zitat als Impuls (Wunsch der Lehrkraft statt Bildimpuls): Es führt direkt zu Ohnmacht und Resignation, und in der Deutung wird darauf zurückgegriffen – der Bogen schließt sich."],
  ["Direkte Instruktion", "Ziel und Ablauf zu Beginn (Seitentafel, Folie). Vorwissen aktiv abgeprüft (erst allein, dann Stichproben, an die Tafel) – hat sich bewährt. Informieren kurz: Kontext, Modell-Lesen durch die Lehrkraft, Kernbotschaft vor der Analyse. Ich – Wir – Ihr: Das Tor wird laut denkend vorgemacht, die Klasse trägt mit ein, dann deutet sie allein. Rückmeldung: Lösungen nacheinander aufdecken, mit Grün verbessern; Zielcheck mit Kriterien."],
  ["Muster der Mentorin und Rückmeldungen", "Kontext zuerst, ein Schwerpunkt (Symbole), Kernbotschaft vor der Analyse, Leitwörter zu jedem Bildelement („achte auf …“), Transfer (*). Aufgaben ohne Dopplungen: Handlungsschritte ordnen (Inhalt) → Symbole deuten (Analyse) → Zielcheck (Urteil „kafkaesk“). Erarbeitung allein, keine Partnerarbeit. 27 pt mit Verlauf und Lernzielen. Satzgerüst im Zielcheck übt nebenbei das Komma vor „weil“. Keine Hausaufgabe."],
  ["Differenzierung", "Leitwörter in der Tabelle als eingebaute Hilfe; Mindestziel sechs Bildelemente, Schnelle bearbeiten die drei mit Stern; Hilfe beim Herumgehen nur als Frage („Was fällt an dem Wort auf?“); Satzgerüst für den Zielcheck."],
  ["Beobachtungsschwerpunkte", "Kommt „kafkaesk“ aus der Klasse oder muss es neu erklärt werden? Reichen 4 Minuten für das Vorlesen und 9 Minuten für die Tabelle? Wird die Deutung des Gesetzes offen gehalten und am Text begründet? Erkennen die Schüler die Selbstbeschränkung des Mannes („jetzt aber nicht“)? Wie viele nennen im Zielcheck ein Merkmal von „kafkaesk“ (→ Wiederholung nächste Stunde)?"],
];
const page4 = [
  kicker("TAFELSKRIPT · HEFT, HINWEISE, REFLEXION", true),
  h2("So sehen Arbeitsblatt und Heft am Ende aus"),
  p(t("Arbeitsblatt (in den Hefter): Aufgabe 1 nummeriert mit Zeilen, Aufgabe 2 Tabelle ausgefüllt (mindestens Tor, Türhüter, Gesetz, Mann, Warten, Bestechung), mit Grün verbessert. Heft: Stichworte zu „kafkaesk“ (Vorwissen), Zielcheck (2–3 Sätze, mit Grün kontrolliert).", { size: 18 }), { line: 270 }),
  h2("Hinweise zu Text und Zitat"),
  box([
    p(t(D.TEXT_HINWEIS, { size: 17 }), { after: 30 }),
    p(t("Zitat (Folie 1): aus Kafkas Nachlass (Notizen), Wortlaut bitte mit einer Ausgabe abgleichen. Alternative zum Thema Behörden: „Die Fesseln der gequälten Menschheit sind aus Kanzleipapier.“ – nur bei Gustav Janouch („Gespräche mit Kafka“) überliefert, Echtheit umstritten; deshalb nicht als Einstieg gewählt.", { size: 17 }), { after: 0 }),
  ], "FFF8EC"),
  p(t(""), { after: 40 }),
  h1("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"),
  table(rW, refl.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 19 }), { after: 0 }), { w: rW[0], fill: LIGHT }), cell(p(t(v, { size: 19 }), { after: 0, line: 280 }), { w: rW[1] })]))),
];

const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Kafka, „Vor dem Gesetz“ · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 19 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } } }, footers: { default: footer }, children: [...page1, ...page2, ...page3, ...page4] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelskript_Vor_dem_Gesetz.docx", b));
