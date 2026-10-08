const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", BAND = "5B6B8C", TOTAL = "D8E1F3", RED = "D9534F";
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 18, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: [].concat(runs), alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 0 }, keepNext: o.keepNext });
const side = (c = "B9C6E8") => ({ style: BorderStyle.SINGLE, size: 6, color: c });
const borders = { top: side(), bottom: side(), left: side(), right: side() };
const cell = (content, w, o = {}) => new TableCell({
  children: [].concat(content), width: { size: w, type: WidthType.DXA }, columnSpan: o.span, borders,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  verticalAlign: VerticalAlign.CENTER, margins: { top: 40, bottom: 40, left: 80, right: 80 },
});
const row = (cells) => new TableRow({ children: cells, cantSplit: true });
const head = (labels, cols) => row(labels.map((x, i) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { align: i === 0 || labels.length > 4 && i < 3 ? AlignmentType.LEFT : AlignmentType.LEFT }), cols[i], { fill: NAVY })));
const h1 = (text) => p(t(text, { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 40 });
const h2 = (text, brk) => new Paragraph({ children: [t(text, { font: "Cambria", size: 24, bold: true, color: NAVY })], spacing: { before: brk ? 0 : 200, after: 80 }, keepNext: true, pageBreakBefore: !!brk });
const bullet = (runs) => new Paragraph({ children: [].concat(runs), bullet: { level: 0 }, spacing: { after: 40 } });

// --- Daten (Einschätzung durch Claude nach dem Bewertungsraster; keine Noten) ---
const briefe = [
  [1, "„… ich finde, was in der Pressekonferenz passierte …“", 190, 13, 2],
  [2, "„… wie ihr gemerkt habt, bin ich … eskaliert“", 185, 17, 5],
  [3, "„… das heutige Spiel war leider eine Niederlage …“", 135, 12, 3],
  [4, "„Sehr geehrte Herr …“", 175, 17, 4],
  [5, "„… dass ich euch … schlecht gereet habe“", 205, 16, 6],
  [6, "„… dass ich nicht an eurer lestung zweifele …“", 98, 11, 8],
  [7, "„Wiedergut mach Brief“", 140, 14, 5],
  [8, "„… ist einiges schief gelaufen …“", 180, 15, 3],
  [9, "„… viele Dinge gesagt, die ich bereue“", 185, 16, 5],
  [10, "„damals war ich sehr wütend …“", 86, 12, 8],
  [11, "„bestimt abt ihr … mitbekommen“", 310, 16, 6],
  [12, "„Guten Tag ich möchte mich schriftlich …“", 200, 15, 2],
  [13, "„… über das Spiel mir ein paar Gedanken …“", 160, 15, 3],
  [14, "„Vorab möchte ich mich bei euch …“", 185, 14, 2],
];
const sum = (i) => briefe.reduce((s, b) => s + b[i], 0);
const avg = (i) => (sum(i) / briefe.length).toFixed(1).replace(".", ",");
const avgG = (sum(3) + sum(4)) / briefe.length;

const C1 = [600, 4400, 1000, 1400, 1400, 1406];
const t1 = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: C1, rows: [
  head(["Nr.", "Erkennbar am Anfang", "Wörter ca.", "Inhalt /20", "Sprache /10", "Gesamt /30"], C1),
  ...briefe.map((b, k) => row([b[0], b[1], b[2], b[3], b[4], b[3] + b[4]].map((v, i) =>
    cell(p(t(String(v), { size: 17, bold: i === 5, color: i === 5 ? NAVY : "000000" }), { align: i === 1 ? AlignmentType.LEFT : AlignmentType.CENTER }), C1[i], { fill: k % 2 ? undefined : LIGHT })))),
  row([cell(p(t("Durchschnitt", { bold: true, color: NAVY })), C1[0] + C1[1], { span: 2, fill: TOTAL }),
    cell(p(t(String(Math.round(sum(2) / briefe.length)), { bold: true }), { align: AlignmentType.CENTER }), C1[2], { fill: TOTAL }),
    cell(p(t(avg(3), { bold: true }), { align: AlignmentType.CENTER }), C1[3], { fill: TOTAL }),
    cell(p(t(avg(4), { bold: true }), { align: AlignmentType.CENTER }), C1[4], { fill: TOTAL }),
    cell(p(t(avgG.toFixed(1).replace(".", ","), { bold: true, color: NAVY }), { align: AlignmentType.CENTER }), C1[5], { fill: TOTAL })]),
] });

const kriterien = [
  ["I  Inhalt", null, null, null],
  ["Sachinhalt klären", "2,1", "4", "Meist nur „Leistung nicht gut“, selten konkret (z. B. Defensivarbeit, Training). In 9 von 14 Briefen höchstens 2 Punkte."],
  ["Selbstoffenbarung ehrlich", "3,1", "4", "Wut, Frust, Scham werden oft ehrlich benannt. Selten als Ich-Botschaft mit Anlass („enttäuscht, als …“)."],
  ["Beziehung reparieren", "3,4", "4", "Stärkstes Kriterium: Entschuldigung, Respekt, Einsicht. Schwächer, wenn Vorwürfe oder Belehrung folgen."],
  ["Appell an die Richtigen", "2,4", "4", "Meist nur Hoffnung („Ich hoffe, ihr …“), keine Bitte mit wer, was, bis wann. In 7 von 14 Briefen höchstens 2 Punkte."],
  ["Metakommunikation", "3,5", "4", "Gut: Die meisten benennen, dass sie intern statt öffentlich hätten sprechen müssen."],
  ["II  Sprache", null, null, null],
  ["Ausdruck und Ton", "1,4", "2", "Respektvoll, aber Umgangssprache (sauer, ausrasten, Job, rummeckern, runtermachen) und Bandwurmsätze."],
  ["Briefform", "0,9", "2", "Grußformel und Name fehlen in 8 von 14 Briefen. MFG/LG in 3, Ort/Datum meist nicht erkennbar."],
  ["Rechtschreibung", "0,4", "2", "Nur 6 von 28 möglichen Punkten. Meist 0 oder 1."],
  ["Grammatik", "1,0", "2", "Kasus, Satzbau, Verbstellung im Nebensatz."],
  ["Zeichensetzung", "0,6", "2", "Schwächstes Kriterium nach der Rechtschreibung. Fast immer fehlt das Komma vor dem Nebensatz."],
];
const C2 = [3000, 900, 900, 5406];
const t2 = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: C2, rows: [
  head(["Kriterium", "Ø", "Max.", "Befund"], C2),
  ...kriterien.map((k) => k[1] === null
    ? row([cell(p(t(k[0], { bold: true, color: "FFFFFF" })), W, { span: 4, fill: BAND })])
    : row([cell(p(t(k[0], { bold: true, color: NAVY })), C2[0], { fill: LIGHT }),
      cell(p(t(k[1], { bold: true }), { align: AlignmentType.CENTER }), C2[1]),
      cell(p(t(k[2]), { align: AlignmentType.CENTER }), C2[2]),
      cell(p(t(k[3], { size: 17 })), C2[3])])),
] });

const fehler = [
  ["Komma vor Nebensatz, Infinitivgruppe, „aber“", "Ich hoffe ihr … · sagt was … · die Chance euch zu … · gebe um … zu", "12 von 14", "Größter Hebel. Test: Steht das Verb am Ende, steht davor ein Komma."],
  ["Groß- und Kleinschreibung", "leistung, spiel · Satzanfang klein · aber auch Entschuldigen, Gute, Persönlich groß", "12 von 14", "Zwei Richtungen: Nomen zu klein, Verben und Adjektive zu groß (Überkorrektur). Test: Artikel davorsetzen."],
  ["Flüchtigkeitsfehler", "Leistng · nohcmal · Bruef · un dFür · bestimt", "13 von 14", "Fast alle Briefe. Fehlendes Gegenlesen, kein Wissensproblem."],
  ["Zusammen- und Getrenntschreibung", "Hier mit · vorallem · soetwas · fest gestellt · Top Spieler · loß gestellt", "10 von 14", "Wiederkehrend: hiermit, vor allem, so etwas, bloßgestellt, festgestellt."],
  ["Kasus und Präposition", "wegen dem · mit den Rest · von euch drei · an unsere Schwächen · schaden + Akkusativ", "9 von 14", "Nach mit, von, an, wegen und bei schaden, arbeiten an steht der Dativ."],
  ["das / dass", "dass statt das, das statt dass, oft mehrfach pro Brief (bis 4×)", "6 von 14", "Ersatzprobe dieses/welches/jenes. Das Muster wiederholt sich oft mehrfach im selben Brief."],
  ["Satzbau", "abgetrennte dass-Sätze · fehlendes Subjekt · Verb nicht am Ende des Nebensatzes", "10 von 14", "Satzfragmente und Subjektwechsel nach „und“. Mehr Gliederung statt Bandwurm."],
  ["Briefform", "Grußformel und Name fehlen · MFG/LG · Sie/ihr gemischt · Punkt statt Komma nach der Anrede", "8 von 14 (Gruß)", "Anredepronomen einheitlich, Grußformel in eigene Zeile."],
];
const C3 = [2300, 3300, 1100, 3506];
const t3 = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: C3, rows: [
  head(["Fehlerart", "Beispiele", "Briefe", "Hinweis"], C3),
  ...fehler.map((f, k) => row(f.map((v, i) => cell(p(t(v, { size: 17, bold: i === 0, color: i === 0 ? NAVY : "000000" }), { align: i === 2 ? AlignmentType.CENTER : AlignmentType.LEFT }), C3[i], { fill: i === 0 ? LIGHT : (k % 2 ? LIGHT : undefined) })))),
] });

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 850, right: 850 } } },
    children: [
      p(t("DEUTSCH 11 · KLASSENAUSWERTUNG", { size: 17, bold: true, color: MUTED }), { after: 20 }),
      h1("Brief „Entschuldigung nach der Pressekonferenz“"),
      p(t("Auswertung von 14 Briefen nach dem Bewertungsraster (Inhalt 20 + Sprache 10 Punkte). Die Punkte sind Einschätzungen von Claude, keine Noten. Die Skala für die Sprachzeilen (bis 1 Fehler pro 100 Wörter = 2 Punkte, bis 3 = 1 Punkt, mehr = 0) ist ein Vorschlag, der noch nicht mit der Fachschaft abgestimmt ist. Gezählt wurde jeder Fehler einzeln.", { size: 17, color: MUTED }), { after: 120 }),
      h2("Ergebnisse je Brief"),
      t1,
      p(t("Die Spanne reicht von 15 bis 22 von 30 Punkten. Inhaltlich liegen die Briefe eng beieinander (Ø 14,5 von 20), sprachlich weit auseinander (Ø 4,4 von 10). Briefe mit der besten Inhaltspunktzahl (Brief 2, 4) sind sprachlich nicht automatisch am besten.", { size: 17 }), { before: 80 }),
      h2("Auswertung nach Kriterien"),
      t2,
      h2("Häufigste Fehler", true),
      t3,
      h2("Inhaltlich aufgefallen"),
      bullet([t("Stärken: ", { bold: true, color: NAVY }), t("Die Schüler übernehmen Verantwortung (Beziehung 3,4 von 4, Metakommunikation 3,5 von 4). Mehrere benennen, dass sie intern hätten sprechen sollen („unter vier Augen … nicht vor Tausenden“).")]),
      bullet([t("Sachinhalt: ", { bold: true, color: NAVY }), t("Die ursprüngliche Kritik wird selten konkret benannt. Meist „Leistung nicht gut“ ohne Bezug zu Training, Defensivarbeit oder Einsatz.")]),
      bullet([t("Appell: ", { bold: true, color: NAVY }), t("Oft nur „Ich hoffe, ihr …“ statt einer Bitte mit wer, was, bis wann. Die besten Appelle nennen ein persönliches Gespräch, das nächste Training oder einen Termin.")]),
      bullet([t("Kippen in Vorwurf: ", { bold: true, color: NAVY }), t("In mehreren Briefen (z. B. 8, 11, 14) beginnt die Entschuldigung stark und endet mit Belehrung („nicht rummeckert“, „ihr nicht umsonst Profis geworden seid“). Das drückt Beziehung und Ton.")]),
      h2("Konsequenzen für den Unterricht"),
      p(t("In dieser Reihenfolge würde ich üben (Ideen, noch kein Material):", { size: 17 }), { after: 40 }),
      bullet([t("1. Komma vor Nebensätzen ", { bold: true }), t("(dass, was, wenn, weil, Relativsatz, um … zu, aber): an echten Sätzen aus den Briefen mit Kreide-Kette.")]),
      bullet([t("2. Groß- und Kleinschreibung ", { bold: true }), t("mit Artikelprobe, auch die Überkorrektur bei Verben und Adjektiven (Brief 13).")]),
      bullet([t("3. das / dass ", { bold: true }), t("mit Ersatzprobe, dazu Gegenlesen als feste Routine (13 von 14 Briefen mit Flüchtigkeitsfehlern).")]),
      bullet([t("4. Zusammenschreibung ", { bold: true }), t("der häufigsten Wörter (hiermit, vor allem, so etwas, festgestellt) und Dativ nach Präposition.")]),
      bullet([t("5. Brief aufbauen: ", { bold: true }), t("Kritik konkret benennen, dann eine Bitte mit wer, was, bis wann, dann Gruß und Name. Entschuldigung und Forderung bewusst trennen.")]),
      p(t("Passend dazu: Fehlertexte statt Diktat, mit Fehlern aus diesen Briefen (anonymisiert), Lösung an der Tafel und Verbessern mit Grün.", { size: 17, italics: true, color: MUTED }), { before: 60 }),
    ],
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync("../Auswertung_Brief.docx", b); });
