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
  "die Chance war persönlich – durch Warten vertan",
];
const bW = [3700, 1300, W - 5000];
const board = table(bW, [
  row([bc([
    p(c("Franz Kafka: „Vor dem Gesetz“ – eine kafkaeske Parabel", { font: "Cambria", size: 30, bold: true }), { after: 20 }),
    p(c("Seitentafel: Ziele „Ich kann die zentralen Symbole der Parabel deuten.“ „Ich kann erklären, warum die Parabel kafkaesk ist.“ · Ablauf: Zitat (Kafka) – kafkaesk (Sprint, nervig oder kafkaesk?) – Lesen – Einleitungssatz – Symbole – Hypothese prüfen – Zielcheck", { size: 17, color: DIM, italics: true }), { after: 0 }),
  ], { w: W, span: 3, borders: { top: none, bottom: line, left: none, right: none } })]),
  row([bc([
    p([c("① Kafka", { bold: true, color: YEL, size: 22 }), c("   Zitat: „… von allen absperren …“ (Tagebuch 1913) → " + D.ZITAT.wissen.map(([k]) => k).join(" · ") + " (Stichworte der Klasse)", { size: 19 })], { after: 30 }),
    p([c("kafkaesk", { bold: true, color: YEL, size: 21 })], { after: 10 }),
    p([c("Bausteine: ", { bold: true }), c(D.KAFKAESK.bausteine.map(([b]) => b).join(" · "))]),
    p([c("Gefühle: ", { bold: true }), c(D.KAFKAESK.gefuehle)]),
    p([c("nervig ", { bold: true }), c("= Grund klar, ich kann etwas tun   ·   "), c("kafkaesk ", { bold: true, color: APP }), c("= übermächtig, undurchschaubar, kein Ausweg")], { after: 0 }),
  ], { w: W, span: 3 })], 1300),
  row([bc([
    p(c("② Einleitungssatz mit Hypothese", { bold: true, color: YEL, size: 22 }), { after: 30 }),
    p(c(D.TATT.satz + " " + D.TATT.hypothese, { size: 19, italics: true }), { after: 0 }),
  ], { w: W, span: 3 })], 1000),
  row([
    bc(p(c("③ Bildebene", { bold: true, color: YEL })), { w: bW[0], fill: BOARD2 }),
    bc(p(c("vgl. Z.", { bold: true, color: YEL })), { w: bW[1], fill: BOARD2 }),
    bc(p(c("Sachebene", { bold: true, color: YEL })), { w: bW[2], fill: BOARD2 }),
  ]),
  ...D.SYMBOLE.map(([b, anker], i) => row([
    bc(p(c(b, { bold: true, color: APP, size: 19 }), { after: 0 }), { w: bW[0], m: 30 }),
    bc(p(c(D.zeile(anker).replace("Z. ", ""), { size: 18 }), { after: 0 }), { w: bW[1], m: 30 }),
    bc(p(c(TAFEL_SACH[i], { size: 19 }), { after: 0 }), { w: bW[2], m: 30 }),
  ])),
  row([bc([
    p([c("④ Hypothese geprüft:  ", { bold: true, color: YEL, size: 22 }), c(D.HYPOTHESE.belege.map(([e, , anker]) => e + " (vgl. " + D.zeile(anker) + ")").join(" · "), { size: 19 })], { after: 20 }),
    p(c("→ komplementär (asymmetrisch): Der Mann redet mit niemandem außer dem Türhüter – am Ende ausgesperrt (Kafka: „von allen absperren“).", { size: 19, italics: true }), { after: 0 }),
  ], { w: W, span: 3, fill: BOARD2, borders: { top: line, bottom: none, left: none, right: none } })], 650),
]);
const page1 = [
  p([t("TAFELBILD IM ENDSTAND · ", { size: 17, bold: true, color: MUTED }), t("① Ankommen und Vorwissen (nach dem Sprint) · ② nach dem Lesen · ③ beim Vormachen und in der Besprechung · ④ Hypothese prüfen", { size: 17, color: MUTED })], { after: 60 }),
  board,
];

// ================= Seite 2: Planungskopf und Ablauf =================
const kopfRows = [
  ["Ziel (Ich kann …)", D.LERNZIELE.map(([z]) => z).join(" ")],
  ["Prüfaufgabe", "Zielcheck im Heft: " + D.ZIELCHECK.auftrag],
  ["Beitrag zur Reihe", "Zweite Parabel nach „Gibs auf!“: Der Türhüter ist eine noch mächtigere Instanz. Einleitungssatz mit Hypothese und „vgl. Z.“ üben für die Klausur; die Hypothese (asymmetrisches Gespräch) verbindet die Parabel mit Watzlawick."],
  ["Mindestziel", "Schlank geplant (45 Min.): Kafka und kafkaesk reaktiviert, Text gelesen, Einleitungssatz, drei Symbole, Zielcheck. Bei Zeitnot Hypothese nur an einem Beleg prüfen."],
  ["Bezug zur letzten Stunde", "„Gibs auf!“ lief sehr gut (guter Transfer, gute Reflexion). Die Mentorin hat am 9.11. „kafkaesk“ mit vier Bausteinen eingeführt – diese Begriffe werden wiederholt; ihre Szenarien (Zug, Amt) werden nicht wiederholt."],
];
const kW = [2800, W - 2800];
const aW = [900, 2300, 7500, 2338, 2300];
const ablauf = [
  ["0–4'", "A: Ankommen (Kafka reaktivieren)", "indirekt (Impuls)", "Folie 1: „Stifte liegen. Lest den Satz still. Wie wirkt er auf euch?“ – „Wie kann jemand so etwas über sich sagen? Was wisst ihr über Kafka?“ Stichworte an die Tafel (①). Auflösung: Tagebuch 1913, Heiratsfrage. Gelenk: „Was bedeutet ‚kafkaesk‘?“", "① Kafka", "Folie 1"],
  ["4–5'", "Verlauf, Lernziele", "direkt", "Folie 2–3. Ziele und Ablauf an der Seitentafel.", "Seitentafel", "Folie 2–3"],
  ["5–6'", "V: Wiederholung", "Unterrichtsgespräch", "Folie 4: „kafkaesk“ mündlich (1–2 Stichproben), noch nichts anschreiben.", "–", "Folie 4"],
  ["6–9'", "V: Begriffs-Sprint", "Einzelarbeit", "Folie 5: „Stift in die Hand. Ihr schreibt eine Minute ohne Pause. Der Stift bleibt auf dem Papier, bis ich ‚Stopp‘ sage.“ Folie 6: Begriff erscheint, Timer 1 Min. „Stopp.“ 2–3 lesen vor. Folie 7: Lösung (vier Bausteine, Gefühle) an die Tafel.", "① Bausteine", "Folie 5–7"],
  ["9–15'", "V: Nervig oder kafkaesk?", "allein, dann Besprechung", "Szenarien-Blatt austeilen. Folie 8: „Jetzt arbeitet ihr allein. Ihr schreibt nur Nummer, N oder K und den Baustein ins Heft.“ 3 Min. Folie 9–10: erst eine Begründung vom Platz („… ist kafkaesk, weil …“), dann aufdecken, mit Grün verbessern.", "① nervig / kafkaesk", "Folie 8–10, Szenarien-Blatt"],
  ["15–21'", "I: Kontext, Lesen", "direkt", "Folie 11: Kontext; Process-Satz: „Welcher Baustein ist das?“ (grundlose Schuld). Arbeitsblatt austeilen. „Stifte liegen. Ihr lest auf dem Arbeitsblatt mit.“ Lehrkraft liest vor (Modell-Lesen, ca. 4 Min.). Leseauftrag: „" + D.LESEAUFTRAG + "“", "–", "Folie 11, AB Nr. 1"],
  ["21–24'", "I: Einleitungssatz", "gemeinsam (Ich – Wir)", "Folie 12: Lehrkraft beginnt den Satz, Klasse ergänzt das Thema. Impuls: „Wer hat im Gespräch die Macht? Wie nennen wir so eine Beziehung?“ → asymmetrisch / komplementär. Alle schreiben auf dem AB mit (Aufgabe 2). Folie 13 nur bei Bedarf.", "②", "Folie 12–13, AB Aufgabe 2"],
  ["24–31'", "V: Symbole", "Modell, dann allein", "Folie 14: das offene Tor laut denkend vormachen, alle tragen mit ein. Folie 15: „Jetzt arbeitet ihr allein. Die Tabelle steht auf dem Arbeitsblatt, Aufgabe 3.“ Timer 6 Min. Hilfe nur als Frage: „Was fällt an dem Wort auf?“", "③ erste Zeile", "Folie 14–15, AB Aufgabe 3"],
  ["31–34'", "V: Besprechung", "Unterrichtsgespräch", "Folie 16: je erst eine Antwort vom Platz, dann aufdecken. „Vergleicht und verbessert mit Grün.“ Gesetz offen deuten.", "③", "Folie 16"],
  ["34–38'", "A: Hypothese prüfen", "allein, dann gemeinsam", "Folie 17: „Sucht eine Minute allein einen Beleg für die Asymmetrie.“ Stichproben, Folie 18: Belege aufdecken; Rückgriff auf das Zitat vom Anfang.", "④", "Folie 17–18, AB Aufgabe 4"],
  ["38–45'", "A: Zielcheck", "allein, dann Selbstkontrolle", "Folie 19: „Allein, im Heft. 4 Minuten.“ Herumgehen. Folie 20: Musterlösung und Kriterien, mit Grün kontrollieren. Rückblick auf die Lernziele. Nie streichen.", "–", "Folie 19–20"],
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
const lW = [3000, 1300, W - 4300];
const page3 = [
  kicker("TAFELSKRIPT · LÖSUNGEN UND ERWARTUNGEN", true),
  h1("Lösungen"),
  h2("Zitat (Folie 1) und „kafkaesk“ (Folie 4–7)"),
  p(t(D.ZITAT.erwartung + " Kontext: " + D.ZITAT.kontext + " – " + D.KAFKAESK.definition + " Bausteine: " + D.KAFKAESK.bausteine.map(([b, e]) => b + " (" + e + ")").join("; ") + ".", { size: 18 }), { line: 270 }),
  h2("Aufgabe 1: Nervig oder kafkaesk?"),
  table([3000, W - 3000], D.SZENARIEN.map(([stich, , nk, baustein, grund], i) => row([
    cell(p([t((i + 1) + " " + stich + ": ", { bold: true, size: 18 }), t(nk === "K" ? "kafkaesk" : "nervig", { bold: true, size: 18, color: nk === "K" ? RED : GREEN })], { after: 0 }), { w: 3000, fill: LIGHT }),
    cell(p(t((baustein ? baustein + " – " : "") + grund, { size: 18 }), { after: 0 }), { w: W - 3000 }),
  ]))),
  h2("Aufgabe 2: Einleitungssatz"),
  p(t(D.TATT.satz + " " + D.TATT.hypothese, { size: 18, italics: true }), { line: 270 }),
  h2("Aufgabe 3: Symbole"),
  table(lW, [
    row([hdr("Bildebene", lW[0]), hdr("Zeile", lW[1]), hdr("Sachebene (Erwartung)", lW[2])]),
    ...D.SYMBOLE.map(([b, anker, sach]) => row([
      cell(p(t(b, { bold: true, size: 18, color: RED }), { after: 0 }), { w: lW[0], fill: LIGHT }),
      cell(p(t("vgl. " + D.zeile(anker), { size: 18 }), { after: 0 }), { w: lW[1] }),
      cell(p(t(sach, { size: 18 }), { after: 0 }), { w: lW[2] }),
    ])),
  ]),
  h2("Aufgabe 4 und Zielcheck"),
  p(t(D.HYPOTHESE.belege.map(([e, b, anker, erkl]) => e + ": " + b + " (vgl. " + D.zeile(anker) + ") – " + erkl).join(" · ") + " " + D.HYPOTHESE.fazit, { size: 18 }), { line: 270 }),
  p([t("Musterlösung Zielcheck: ", { bold: true, size: 18, color: GREEN }), t(D.ZIELCHECK.muster, { size: 18 }), t("  Kriterien: " + D.ZIELCHECK.kriterien.join(" · ") + ".", { size: 18, color: MUTED })], { line: 270 }),
];

// ================= Seite 4: Umsetzung der Erkenntnisse =================
const rW = [3000, W - 3000];
const refl = [
  ["AVIVA", "Alle fünf Phasen; das Auswerten (Hypothese prüfen, Zielcheck) wird nie gestrichen. Ankommen mit einem Kafka-Zitat (Wunsch der Lehrkraft): Es reaktiviert das Wissen über Kafka (Vater, Schreiben, Beruf, Beziehungen), erst danach folgt „kafkaesk“. Beim Prüfen der Hypothese wird darauf zurückgegriffen („absperren“ – der geschlossene Eingang). Vorwissen mit drei Zugängen zum selben Begriff: mündlich wiederholen, Begriffs-Sprint (alle schreiben), Anwendung auf Szenarien."],
  ["Direkte Instruktion", "Ziel und Ablauf zu Beginn. Vorwissen aktiv abgeprüft. Informieren kurz: Kontext, Modell-Lesen, Einleitungssatz gemeinsam (Lehrkraft beginnt, Klasse ergänzt). Ich – Wir – Ihr bei den Symbolen: Das Tor wird laut denkend vorgemacht, die Klasse deutet drei weitere allein. Rückmeldung: Lösungen aufdecken, mit Grün verbessern; Zielcheck mit Kriterien."],
  ["Rückmeldungen umgesetzt", "45 Minuten realistisch: nur vier Symbole statt neun, keine Handlungsschritte, kein Transfer, kein Stern. Szenarien eindeutig und fast unheimlich, nah an der Erfahrung der Schüler (Schule), nur zwei aus dem Sport, keine Wiederholung der Szenarien der Mentorin. Begriffe der Mentorin („kafkaesk“-Bausteine). Muster der Mentorin: Kontext zuerst, Modellsatz gemeinsam, ein Schwerpunkt. Klausurbezug: Einleitungssatz mit Hypothese, „vgl. Z.“. Erarbeitung allein, 27 pt, keine Hausaufgabe."],
  ["Differenzierung", "Sprint ohne Richtig/Falsch – jeder kann etwas schreiben. Kriterium nervig/kafkaesk auf der Folie. Gerüst für Einleitungssatz und Zielcheck. Hilfe beim Herumgehen nur als Frage."],
  ["Beobachtungsschwerpunkte", "Wie viel wissen die Schüler nach dem Sprint noch von „kafkaesk“? Erkennen sie die Bausteine in den Szenarien sicher? Kommt die Klasse selbst auf „asymmetrisch“? Reichen 6 Minuten für drei Symbole? Wie viele nennen im Zielcheck einen Baustein (→ Wiederholung nächste Stunde)?"],
];
const page4 = [
  kicker("TAFELSKRIPT · HEFT, HINWEISE, REFLEXION", true),
  h2("So sehen Arbeitsblatt und Heft am Ende aus"),
  p(t("Heft: Sprint-Text, Aufgabe 1 (Nummer, N/K, Baustein, mit Grün verbessert), Zielcheck (2–3 Sätze, mit Grün kontrolliert). Arbeitsblatt (in den Hefter): Einleitungssatz mit Hypothese, Symboltabelle (vier Zeilen), zwei Belege für die Asymmetrie. Szenarien-Blatt einkleben.", { size: 18 }), { line: 270 }),
  h2("Hinweise zu Text, Zitat und Begriff"),
  box([
    p(t(D.TEXT_HINWEIS, { size: 17 }), { after: 30 }),
    p(t("Zitat (Folie 1): Kafka, Tagebuch, 21. Juli 1913 (Liste für und gegen die Heirat mit Felice Bauer) – Datum und Wortlaut bitte mit einer Ausgabe abgleichen.", { size: 17 }), { after: 30 }),
    p(t("„Komplementär“ (Watzlawick, 5. Axiom): Die Beziehung beruht auf Unterschiedlichkeit, die Rollen ergänzen sich als oben und unten. „Ergänzend“ heißt nicht harmonisch oder gelungen. Vorschlag: „komplementär (asymmetrisch)“ sagen – vorher mit der Mentorin abstimmen.", { size: 17 }), { after: 0 }),
  ], "FFF8EC"),
  p(t(""), { after: 40 }),
  h1("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"),
  table(rW, refl.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 19 }), { after: 0 }), { w: rW[0], fill: LIGHT }), cell(p(t(v, { size: 19 }), { after: 0, line: 270 }), { w: rW[1] })]))),
];

const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Kafka, „Vor dem Gesetz“ · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 19 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } } }, footers: { default: footer }, children: [...page1, ...page2, ...page3, ...page4] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelskript_Vor_dem_Gesetz.docx", b));
