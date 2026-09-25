const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak, Header, Footer, TabStopType,
  PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA";
const RED = "D9534F";
// Farben der vier Strategien (Expertenkarten A–D)
const S = { A: "3F7CC4", B: "2E9E6B", C: "E8A33D", D: "8E5BB5" };
const W = 10206; // content width (A4, 1.5 cm margins)

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text, pb) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20, pb });
const h2 = (label, text, col = RED) => p([t(label + "  ", { font: "Cambria", size: 26, bold: true, color: col }), t(text, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 140, after: 80, keepNext: true });
const gap = (after = 100) => p(t(""), { after });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || allBorders(solid("B9C6E8", 6)),
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: o.ml ?? 100 },
    columnSpan: o.span,
  });
}
function table(widths, rows) {
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
}
function row(cells, h, rule = HeightRule.ATLEAST) {
  return new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
}
const hdr = (txt, w, fill = NAVY, sub) => cell([
  p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { align: AlignmentType.CENTER, after: 0 }),
  sub ? p(t(sub, { color: "FFFFFF", size: 15 }), { align: AlignmentType.CENTER, after: 0 }) : null,
].filter(Boolean), { w, fill });

function infoBox(paras, fill = LIGHT2, w = W) {
  return table([w], [row([cell(paras.map((r) => p(r, { after: 40 })), { w, fill, borders: allBorders(none), m: 120 })])]);
}
// Schreiblinien
function lines(n, h = 480, label) {
  const b = { top: none, left: none, right: none, bottom: solid("9AA6C4", 6) };
  return table([W], Array.from({ length: n }, (_, i) => row([cell(p(label && i === 0 ? t(label, { size: 18, color: MUTED, bold: true }) : t(""), { after: 0 }), { w: W, borders: b, valign: VerticalAlign.BOTTOM, ml: 0 })], h)));
}
const box = "☐";

// ---------- Kopf / Fuß ----------
const mkHeader = (title) => new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 11 · " + title, { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })],
  })],
});
const footer = (label) => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })],
});

// ================= Inhalt: Beispieldialog =================
const dialog = [
  ["Lena", "Hey, hast du die Folien fertig? Wir sind morgen dran.", null],
  ["Tom", "Fast. Den Rest mach ich heute Abend.", null],
  ["Lena", "Heute Abend? Du hattest eine ganze Woche Zeit.", null],
  ["Tom", "Ich hatte halt dreimal Training. Nicht jeder hat so viel Freizeit wie du.", null],
  ["Lena", "Freizeit? Ich hab die komplette Gliederung gemacht! Immer bleibt alles an mir hängen.", null],
  ["Tom", "Dann mach’s doch allein, wenn du eh alles besser kannst.", null],
  ["Lena", "Super. Genau das hab ich von dir erwartet.", "verdreht die Augen"],
  ["Tom", "Wie du meinst.", "zuckt mit den Schultern und schaut auf sein Handy"],
  ["Lena", "Ja, schau ruhig aufs Handy. Ist ja auch wichtiger.", null],
  ["Tom", "Boah, du nervst echt.", "steht auf und geht"],
];
const dW = [700, 1100, 6106, 2300];
function dialogTable(withNotes = true) {
  const rows = [row([hdr("Z.", dW[0]), hdr("Wer", dW[1]), hdr("Was gesagt oder getan wird", dW[2]), hdr(withNotes ? "↯ Notizen" : "", dW[3])], 420)];
  dialog.forEach(([who, text, stage], i) => rows.push(row([
    cell(p(t(String(i + 1), { bold: true, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: dW[0], fill: LIGHT }),
    cell(p(t(who, { bold: true, color: NAVY }), { after: 0 }), { w: dW[1], fill: LIGHT }),
    cell(p([stage ? t("(" + stage + ") ", { italics: true, color: MUTED, size: 21 }) : null, t(text, { font: "Cambria", size: 22 })].filter(Boolean), { after: 0 }), { w: dW[2] }),
    cell(p(t("")), { w: dW[3] }),
  ], 600, HeightRule.ATLEAST)));
  return table(dW, rows);
}

// ================= ARBEITSBLATT =================
const ab1 = [
  kicker("M1 · Beispieldialog"),
  h1("Das Referat"),
  infoBox([[t("Situation: ", { bold: true, color: NAVY }), t("Donnerstag, große Pause. Lena und Tom halten morgen in der ersten Stunde ein gemeinsames Referat. Die Gliederung hat Lena geschrieben, die Folien sollte Tom machen.", { size: 21 })]]),
  gap(120),
  dialogTable(),
  gap(120),
  h2("1", "Impuls · erst allein (3 Min.), dann zu zweit (2 Min.)"),
  p([t("Lies den Dialog still. Setze am Rand ein "), t("↯", { bold: true, color: RED }), t(" an jede Stelle, an der das Gespräch deiner Meinung nach "), t("kippt", { bold: true }), t(". Noch nicht reden!")], { after: 60 }),
  p([t("Danach mit deinem Sitznachbarn: ", { bold: true }), t("Habt ihr dieselben Stellen markiert? Wo seid ihr euch uneinig?")], { after: 0 }),
];

const kW = [800, 2900, 3606, 2900];
const kRows = [row([hdr("Zeile", kW[0]), hdr("Was passiert?", kW[1], NAVY, "Zitat oder Verhalten"), hdr("Warum kippt es hier?", kW[2], NAVY, "Fachbegriff aus den Kommunikationsmodellen"), hdr("Was hätte geholfen?", kW[3], NAVY, "erst nach Aufgabe 3 ausfüllen")], 620)];
for (let i = 0; i < 5; i++) kRows.push(row(kW.map((w, j) => cell(p(t("")), { w, fill: j === 3 ? LIGHT : undefined })), 1250));

const begriffe = ["Sachebene / Beziehungsebene", "Selbstkundgabe", "Appell", "Beziehungsohr", "Du-Botschaft (Vorwurf)", "Verallgemeinerung („immer“, „nie“)", "Ironie = inkongruente Botschaft", "symmetrische Eskalation", "Interpunktion", "„Man kann nicht nicht kommunizieren.“"];

const ab2 = [
  kicker("Aufgabe 2 · Partnerarbeit · 8 Min. + Unterrichtsgespräch 5 Min.", true),
  h1("Kipppunkte benennen"),
  p([t("Ein "), t("Kipppunkt", { bold: true }), t(" ist die Stelle, an der ein Gespräch von der Sache weg und in einen Konflikt rutscht. Oft gibt es mehrere, die aufeinander aufbauen.")], { after: 80 }),
  p([t("a) ", { bold: true }), t("Tragt "), t("mindestens drei", { bold: true }), t(" Kipppunkte aus M1 in die Tabelle ein. Die letzte Spalte bleibt vorerst leer.")], { after: 40 }),
  p([t("b) ", { bold: true }), t("Markiert den Kipppunkt, an dem man das Gespräch "), t("am leichtesten", { bold: true }), t(" hätte retten können, mit einem Stern ★.")], { after: 40 }),
  p([t("c) ", { bold: true }), t("Im anschließenden Unterrichtsgespräch vergleichen wir. Ergänzt, was euch fehlt, "), t("in einer anderen Farbe", { bold: true }), t(".")], { after: 120 }),
  table(kW, kRows),
  gap(100),
  infoBox([
    [t("Fachbegriffe zur Auswahl", { bold: true, color: NAVY })],
    [t(begriffe.join("  ·  "), { size: 20 })],
  ]),
  gap(60),
  p([t("Nicht weitergekommen? ", { bold: true, color: RED, size: 20 }), t("Erst Tippkarte 1, dann Tippkarte 2 vom Pult holen.", { size: 20, italics: true })], { after: 0 }),
];

const strat = [
  ["A", "Ich-Botschaft", S.A], ["B", "Empathie", S.B], ["C", "Wunsch / Bitte", S.C], ["D", "Metakommunikation", S.D],
];
const legendW = [2551, 2551, 2552, 2552];
const stratLegend = table(legendW, [row(strat.map(([k, n, c], i) => cell(p(t(k + " = " + n, { bold: true, color: "FFFFFF", size: 19 }), { align: AlignmentType.CENTER, after: 0 }), { w: legendW[i], fill: c, borders: allBorders(solid("FFFFFF", 12)) })), 520)]);

const rwW = [700, 1100, 7406, 1000];
const rwRows = [row([hdr("Z.", rwW[0]), hdr("Wer", rwW[1]), hdr("Version B: So hätten Lena und Tom reden können", rwW[2]), hdr("Strat.", rwW[3])], 440)];
rwRows.push(row([
  cell(p(t("1–2", { color: MUTED, bold: true }), { align: AlignmentType.CENTER, after: 0 }), { w: rwW[0], fill: LIGHT }),
  cell(p(t("", {})), { w: rwW[1], fill: LIGHT }),
  cell(p(t("bleiben wie im Original", { italics: true, color: MUTED }), { after: 0 }), { w: rwW[2], fill: LIGHT }),
  cell(p(t("")), { w: rwW[3], fill: LIGHT }),
], 440));
["Lena", "Tom", "Lena", "Tom", "Lena", "Tom", "Lena"].forEach((who, i) => rwRows.push(row([
  cell(p(t(String(i + 3), { bold: true, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: rwW[0], fill: LIGHT }),
  cell(p(t(who, { bold: true, color: NAVY }), { after: 0 }), { w: rwW[1], fill: LIGHT }),
  cell(p(t("")), { w: rwW[2] }),
  cell(p(t("")), { w: rwW[3] }),
], 900)));

const check = [
  "Niemand wird angegriffen oder abgewertet.",
  "Mindestens eine Ich-Botschaft statt einer Du-Botschaft.",
  "Einer geht auf die Gefühle des anderen ein.",
  "Es gibt eine konkrete, erfüllbare Bitte (wer? was? bis wann?).",
  "Kein „immer“, kein „nie“, keine Ironie.",
  "Am Ende steht eine Lösung für die Folien.",
];

const ab3 = [
  kicker("Aufgabe 3 · Gruppenpuzzle · 3 Gruppen à 7 · 17 Min.", true),
  h1("Das Gespräch retten"),
  p([t("0. ", { bold: true }), t("Setzt euch in eure "), t("Reihen-Gruppe", { bold: true }), t(" (Wandreihe, Mittelreihe, Fensterreihe) und öffnet euren Umschlag. Verteilt die vier Rollenkarten.")], { after: 40 }),
  p([t("1. ", { bold: true }), t("Verteilt die Expertenkarten: "), t("A", { bold: true, color: S.A }), t(" liest eine Person allein, "), t("B, C und D", { bold: true }), t(" lesen je zwei Personen im Tandem (4 Min.).")], { after: 40 }),
  p([t("2. ", { bold: true }), t("Reihum erklärt jede Expertin, jeder Experte bzw. jedes Tandem die Strategie in "), t("höchstens 90 Sekunden", { bold: true }), t(" – mit dem Beispiel von der Karte (7 Min.).")], { after: 40 }),
  p([t("3. ", { bold: true }), t("Schreibt den Dialog "), t("ab Zeile 3", { bold: true }), t(" gemeinsam neu – alle schreiben mit. Nutzt mindestens drei der vier Strategien und notiert rechts den Buchstaben (6 Min.).")], { after: 40 }),
  p([t("4. ", { bold: true }), t("Prüft eure Version mit der Checkliste. Im Unterrichtsgespräch liest die Protokoll-Person eure Version vor – dabei füllen alle den Merkkasten auf der nächsten Seite aus.")], { after: 100 }),
  stratLegend,
  gap(100),
  table(rwW, rwRows),
  gap(100),
  p(t("Checkliste Version B", { font: "Cambria", bold: true, size: 24, color: NAVY }), { after: 60, keepNext: true }),
  ...check.map((c) => p([t(box + "  ", { size: 24 }), t(c, { size: 21 })], { after: 30 })),
];


// Merkkasten Werkzeugkasten (Sicherung im Unterrichtsgespräch 2)
const mkW = [2700, 3906, 3600];
const mkRows = [row([hdr("Strategie", mkW[0]), hdr("Formel / So geht’s", mkW[1]), hdr("Beispiel", mkW[2], NAVY, "aus einer vorgelesenen Version")], 560)];
strat.forEach(([k, n, c]) => mkRows.push(row([
  cell([p(t(k, { font: "Cambria", size: 40, bold: true, color: "FFFFFF" }), { after: 0 }), p(t(n, { bold: true, color: "FFFFFF", size: 20 }), { after: 0 })], { w: mkW[0], fill: c, valign: VerticalAlign.TOP, m: 120 }),
  cell(p(t("")), { w: mkW[1] }),
  cell(p(t("")), { w: mkW[2] }),
], 2000)));

const abMerk = [
  kicker("Merkkasten · Unterrichtsgespräch 2 · 7 Min.", true),
  h1("Unser Werkzeugkasten"),
  p([t("Die Gruppen lesen ihre Version B vor. Hört genau hin: "), t("Welche Strategie steckt in welcher Formulierung?", { bold: true }), t(" Tragt für jede Strategie die Formel und das beste Beispiel ein.")], { after: 60 }),
  p([t("Diesen Kasten nutzt ihr gleich für eure eigene Szene.", { italics: true, color: MUTED })], { after: 120 }),
  table(mkW, mkRows),
  gap(160),
  p(t("Das macht Gespräche kaputt", { font: "Cambria", bold: true, size: 24, color: RED }), { after: 60, keepNext: true }),
  lines(3, 480),
];

// Planungsbogen
const pl = (label, h, hint) => row([
  cell([p(t(label, { bold: true, color: NAVY, size: 20 }), { after: 0 }), hint ? p(t(hint, { size: 16, italics: true, color: MUTED }), { after: 0 }) : null].filter(Boolean), { w: 2600, fill: LIGHT, valign: VerticalAlign.TOP }),
  cell(p(t("")), { w: W - 2600 }),
], h, HeightRule.ATLEAST);

const ab4 = [
  kicker("Aufgabe 4 · Reihen-Gruppe · 17 Min.", true),
  h1("Eure eigene Szene"),
  p([t("Eure Gruppe entwickelt "), t("eine", { bold: true }), t(" Alltagsszene in "), t("zwei Versionen", { bold: true }), t(". Beide beginnen gleich. "), t("Version A", { bold: true, color: RED }), t(" kippt und eskaliert. In "), t("Version B", { bold: true, color: S.B }), t(" wird genau am Kipppunkt anders reagiert.")], { after: 60 }),
  p([t("Denkt euch die Situation "), t("selbst", { bold: true }), t(" aus – am besten eine, die ihr so oder ähnlich schon erlebt habt: Familie, Freunde, Schule, Verein, Job, Chat …")], { after: 60 }),
  p([t("Nach 5 Minuten noch keine Idee? ", { bold: true }), t("Material-Person holt Tippkarte 1 für Aufgabe 4. Hilft auch die nicht, hat die Lehrkraft eine Notfall-Idee.", { italics: true })], { after: 60 }),
  infoBox([[t("Spielregeln: ", { bold: true, color: NAVY }), t("2–4 Rollen pro Version · jede Version höchstens 2 Minuten · ", { size: 20 }), t("Version A und B spielen verschiedene Personen", { size: 20, bold: true }), t(" – wer nicht spielt, führt Regie · Stichpunkte auf Karteikarten sind erlaubt · ", { size: 20 }), t("der Kipppunkt muss für das Publikum erkennbar sein.", { size: 20, bold: true })]]),
  gap(100),
  table([2600, W - 2600], [
    pl("Situation", 900, "Wo? Wer? Worum geht es?"),
    pl("Besetzung A", 600, "Figur → gespielt von"),
    pl("Besetzung B", 600, "Figur → gespielt von"),
    pl("Gemeinsamer Anfang", 1300, "2–3 Sätze, die in A und B gleich sind"),
  ]),
  gap(80),
  p(t("Version A – es kippt", { font: "Cambria", bold: true, size: 24, color: RED }), { after: 60, keepNext: true }),
  table([2600, W - 2600], [
    pl("Der Kipppunkt-Satz", 800, "wörtlich"),
    pl("Warum kippt es hier?", 800, "Fachbegriff!"),
    pl("Wie eskaliert es?", 900, "Stichpunkte bis zum Ende"),
  ]),
  gap(80),
  p(t("Version B – es wird gelöst", { font: "Cambria", bold: true, size: 24, color: S.B }), { after: 60, keepNext: true }),
  table([2600, W - 2600], [
    pl("Stattdessen sagt …", 1300, "wörtlich, ab dem Kipppunkt"),
    pl("Wie endet es?", 800, "Welche Lösung finden die Figuren?"),
    row([
      cell(p(t("Eingesetzte Strategien", { bold: true, color: NAVY, size: 20 }), { after: 0 }), { w: 2600, fill: LIGHT }),
      cell(p(strat.map(([k, n, c]) => t(box + " " + k + " " + n + "   ", { size: 18, bold: true, color: c })), { after: 0 }), { w: W - 2600 }),
    ], 520),
  ]),
];

const bW = [1800, 2500, 2500, 2206, 1200];
const bRows = [row([hdr("Gruppe", bW[0], NAVY, "★ = Patenszene"), hdr("Kipppunkt", bW[1], NAVY, "Wann genau? Zitat"), hdr("Ursache", bW[2], NAVY, "Fachbegriff"), hdr("Strategie in Version B", bW[3], NAVY, "A · B · C · D"), hdr("Wirkt?", bW[4], NAVY, "+  o  –")], 620)];
["Wandreihe", "Mittelreihe", "Fensterreihe"].forEach((g, i) => bRows.push(row([
  cell([p(t(String(i + 1), { bold: true, color: MUTED, size: 28 }), { align: AlignmentType.CENTER, after: 20 }), p(t(g, { bold: true, color: NAVY, size: 19 }), { align: AlignmentType.CENTER, after: 0 })], { w: bW[0], fill: LIGHT }),
  ...bW.slice(1).map((w) => cell(p(t("")), { w })),
], 2900)));

const ab5 = [
  kicker("Aufgabe 5 · Bühne frei · 13 Min. + Auswertung 11 Min.", true),
  h1("Beobachtungsbogen"),
  p([t("Spielen: ", { bold: true, color: NAVY }), t("Reihenfolge "), t("Wandreihe → Mittelreihe → Fensterreihe", { bold: true }), t(". Jede Gruppe spielt "), t("Version A", { bold: true, color: RED }), t(" und direkt danach "), t("Version B", { bold: true, color: S.B }), t(" – ohne Unterbrechung. Danach hast du "), t("30 Sekunden", { bold: true }), t(", um die Zeile auszufüllen.")], { after: 60 }),
  p([t("Patenszene: ", { bold: true, color: NAVY }), t("Die Wandreihe beobachtet besonders genau die Mittelreihe, die Mittelreihe die Fensterreihe, die Fensterreihe die Wandreihe. Markiere deine Patenszene mit ★.")], { after: 60 }),
  p([t("Auswertung: ", { bold: true, color: NAVY }), t("2 Min. still ergänzen → 3 Min. in der Gruppe die Patenszene besprechen → im Plenum ca. 2 Min. Rückmeldung pro Gruppe.")], { after: 120 }),
  table(bW, bRows),
];

// ================= KARTENSET =================
function cutCard(children, w = W, h = 7000, border = dashed) {
  return table([w], [row([cell(children, { w, borders: allBorders(border), valign: VerticalAlign.TOP, m: 200, ml: 260 })], h, HeightRule.ATLEAST)]);
}
const cardKicker = (txt, col) => p(t(txt.toUpperCase(), { size: 16, bold: true, color: col }), { after: 20 });
const cardTitle = (txt, col = NAVY, size = 34) => p(t(txt, { font: "Cambria", size, bold: true, color: col }), { after: 80 });
const lbl = (txt, col) => p(t(txt, { bold: true, color: col, size: 20 }), { before: 80, after: 30 });
const bl = (runs, size = 20) => p([t("•  ", { size, bold: true, color: MUTED }), ...(Array.isArray(runs) ? runs : [t(runs, { size })])], { after: 30 });

function formulaBox(txt, col) {
  return table([W - 520], [row([cell(p(t(txt, { font: "Cambria", size: 22, bold: true, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w: W - 520, fill: col, borders: allBorders(none), m: 100 })], 0)]);
}
function vsBox(from, to, col) {
  const w2 = (W - 520) / 2;
  return table([w2, w2], [row([
    cell([p(t("statt", { size: 16, bold: true, color: RED }), { after: 10 }), p(t(from, { font: "Cambria", size: 21, italics: true }), { after: 0 })], { w: w2, fill: "FBECEB", borders: allBorders(solid("FFFFFF", 12)), valign: VerticalAlign.TOP }),
    cell([p(t("besser", { size: 16, bold: true, color: col }), { after: 10 }), p(t(to, { font: "Cambria", size: 21, italics: true }), { after: 0 })], { w: w2, fill: LIGHT2, borders: allBorders(solid("FFFFFF", 12)), valign: VerticalAlign.TOP }),
  ], 0)]);
}

function expertCard(key, name, col, idea, formula, steps, z, from, to, trap) {
  return cutCard([
    cardKicker("Expertenkarte " + key, col),
    cardTitle(name, col),
    p(t(idea, { size: 21 }), { after: 80 }),
    formulaBox(formula, col),
    lbl("So geht’s", col),
    ...steps.map((s) => bl(s)),
    lbl("Beispiel aus M1 (Zeile " + z + ")", col),
    vsBox(from, to, col),
    lbl("Stolperfalle", RED),
    p(t(trap, { size: 20 }), { after: 0 }),
  ]);
}

const expA = expertCard("A", "Ich-Botschaft", S.A,
  "Du-Botschaften klingen wie ein Urteil über den anderen – der hört sie mit dem Beziehungsohr und verteidigt sich. Eine Ich-Botschaft sagt, was eine Situation mit mir macht. Das kann niemand bestreiten.",
  "Wenn … (Beobachtung),  fühle ich mich … (Gefühl),  weil … (Grund).",
  ["Beschreibe, was passiert ist, ohne es zu bewerten.", "Nenne dein Gefühl: gestresst, enttäuscht, unsicher, übergangen …", "Erkläre, warum dich das betrifft."],
  3, "„Du hattest eine ganze Woche Zeit.“",
  "„Wenn ich höre, dass die Folien erst heute Abend kommen, werde ich nervös, weil wir morgen schon dran sind.“",
  "„Ich finde, du bist unzuverlässig“ ist nur eine verkleidete Du-Botschaft. Hinter „Ich finde, du …“ steckt fast immer ein Urteil.");

const expB = expertCard("B", "Empathie & aktives Zuhören", S.B,
  "Wer sich angegriffen fühlt, will zuerst verstanden werden – erst dann ist er offen für Lösungen. Aktives Zuhören zeigt: Ich habe dich gehört, auch wenn ich nicht zustimme.",
  "Du meinst also … ?   ·   Das klingt, als ob du … bist.",
  ["Wiederhole mit eigenen Worten, was du verstanden hast (paraphrasieren).", "Sprich das Gefühl an, das du hinter der Aussage vermutest.", "Frage nach, statt sofort zu antworten oder dich zu rechtfertigen."],
  4, "„Nicht jeder hat so viel Freizeit wie du.“",
  "„Du bist gestresst, weil morgen alles klappen soll, oder? Das verstehe ich.“",
  "Zuhören ist kein „Ja, aber …“. Keine Ratschläge, keine Gegenrechnung – erst verstehen, dann antworten.");

const expC = expertCard("C", "Wunsch & Bitte", S.C,
  "Hinter jedem Vorwurf steckt ein unerfüllter Wunsch. Die gewaltfreie Kommunikation (Marshall Rosenberg) macht ihn sichtbar und verwandelt ihn in eine Bitte, auf die der andere reagieren kann.",
  "Beobachtung  →  Gefühl  →  Bedürfnis  →  Bitte",
  [[t("Formuliere ", { size: 20 }), t("positiv", { size: 20, bold: true }), t(": was du willst, nicht was der andere lassen soll.", { size: 20 })], [t("Formuliere ", { size: 20 }), t("konkret", { size: 20, bold: true }), t(": wer, was, bis wann?", { size: 20 })], [t("Eine Bitte ist ", { size: 20 }), t("verhandelbar", { size: 20, bold: true }), t(" – der andere darf ein Gegenangebot machen.", { size: 20 })]],
  5, "„Immer bleibt alles an mir hängen.“",
  "„Mir ist wichtig, dass wir morgen sicher sind. Kannst du mir die Folien bis 19 Uhr schicken, damit ich drüberschauen kann?“",
  "„Du musst endlich mal pünktlich liefern!“ ist eine Forderung, keine Bitte. Wer bei einem Nein sauer wird, hat nicht gebeten, sondern verlangt.");

const expD = expertCard("D", "Metakommunikation", S.D,
  "Manchmal ist das eigentliche Thema längst verloren – man streitet nur noch darüber, wie man miteinander redet. Dann hilft es, kurz aus dem Gespräch herauszutreten und über das Gespräch selbst zu sprechen.",
  "Stopp – wie reden wir hier eigentlich gerade miteinander?",
  ["Benenne, was gerade im Gespräch passiert (ohne Schuldzuweisung).", "Kläre Missverständnisse: „Wie hast du das gemeint?“", "Erinnere an das gemeinsame Ziel und schlage einen Neustart vor."],
  7, "„Super. Genau das hab ich von dir erwartet.“",
  "„Moment. Wir werfen uns gerade nur noch Sachen an den Kopf. Eigentlich wollen wir doch beide ein gutes Referat, oder?“",
  "Metakommunikation ist kein neuer Vorwurf. „Du redest immer so gemein mit mir“ ist wieder eine Du-Botschaft mit „immer“.");

const roles = [
  ["Zeit", "Du hast den Timer im Blick. Sag Bescheid, wenn die Hälfte der Zeit um ist und wenn noch 2 Minuten bleiben."],
  ["Protokoll", "Du schreibst die Ergebnisse der Gruppe auf. Am Ende muss euer Bogen vollständig sein."],
  ["Regie", "Du achtest darauf, dass alle zu Wort kommen. Bei der Szene verteilst du die Besetzung für A und B. Im Plenum sprichst du für die Gruppe."],
  ["Material", "Du holst Karten, Tipps und Lösungen. Nur du darfst die Lehrkraft fragen – und erst, wenn die Tippkarte nicht geholfen hat."],
];
const rW = [2551, 2551, 2552, 2552];
const roleGrid = table(rW, [row(roles.map(([n, d], i) => cell([
  p(t("ROLLENKARTE", { size: 14, bold: true, color: MUTED }), { after: 20 }),
  p(t(n, { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 60 }),
  p(t(d, { size: 18 }), { after: 0 }),
], { w: rW[i], borders: allBorders(dashed), valign: VerticalAlign.TOP, m: 140 })), 2600, HeightRule.ATLEAST)]);

const tipps = [
  ["Tipp 1 · Aufgabe 2", "Achte auf Sätze, die mit „Du …“ beginnen, auf Wörter wie „immer“ und „nie“ und darauf, was Lena und Tom tun, ohne etwas zu sagen."],
  ["Tipp 2 · Aufgabe 2", "Schau dir die Zeilen 3, 4, 5, 7 und 8 an. Frag jeweils: Was meint die Person eigentlich – und was kommt beim anderen an?"],
  ["Tipp 1 · Aufgabe 4", "Jede Person erzählt kurz einen Streit, den sie erlebt oder beobachtet hat – zu Hause, mit Freunden, im Verein, im Job, im Chat. Nehmt den, bei dem ihr am meisten nicken musstet."],
  ["Tipp 2 · Aufgabe 4", "Bauplan für Version A: harmloser Anfang → ein Satz mit Du-Vorwurf, „immer“ oder Ironie → Gegenangriff → Abbruch. Version B beginnt genauso, reagiert aber am Kipppunkt anders."],
];
const tW = [5103, 5103];
const tippGrid = table(tW, [0, 2].map((k) => row([0, 1].map((j) => {
  const [head, txt] = tipps[k + j];
  return cell([
    p(t("?  " + head.toUpperCase(), { size: 16, bold: true, color: S.C }), { after: 40 }),
    p(t(txt, { size: 20 }), { after: 0 }),
  ], { w: tW[j], borders: allBorders(dashed), valign: VerticalAlign.TOP, m: 160, fill: "FFF8EC" });
}), 1900, HeightRule.ATLEAST)));

const situations = [
  ["Gruppenarbeit", "Eine Person im Projekt hat bis zur Abgabe nichts beigetragen. Morgen ist Präsentation."],
  ["Zuhause", "Das Geschirr steht seit drei Tagen in der Spüle. Jeder behauptet, der andere sei dran."],
  ["Party", "Du willst am Samstag bis 2 Uhr auf eine Party. Deine Eltern sagen: um 23 Uhr bist du zuhause."],
  ["Chat", "Du schreibst deiner besten Freundin eine lange Nachricht. Die Antwort nach drei Stunden: „ok.“"],
  ["Verein", "Du wirst im wichtigen Spiel nicht aufgestellt, obwohl du bei jedem Training warst."],
  ["Nebenjob", "Deine Chefin trägt dich kurzfristig für Samstag ein – genau an dem Tag, an dem du frei haben wolltest."],
  ["Freundschaft", "Du hast einem Freund etwas Persönliches erzählt. Jetzt weiß es die halbe Stufe."],
  ["Klausur", "Du findest die Note deiner Klausur ungerecht und sprichst die Lehrkraft nach der Stunde darauf an."],
];
const sW = [5103, 5103];
const sitRows = [];
for (let i = 0; i < situations.length; i += 2) {
  sitRows.push(row([0, 1].map((j) => {
    const [n, d] = situations[i + j];
    return cell([
      p(t("NOTFALL-IDEE " + (i + j + 1), { size: 14, bold: true, color: MUTED }), { after: 20 }),
      p(t(n, { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 60 }),
      p(t(d, { size: 21 }), { after: 0 }),
    ], { w: sW[j], borders: allBorders(dashed), valign: VerticalAlign.TOP, m: 200 });
  }), 2500, HeightRule.ATLEAST));
}

const modLine = ([a, b]) => p([t(a + ":  ", { bold: true, color: NAVY, size: 20 }), t(b, { size: 20 })], { after: 50 });
const modCard = cutCard([
  cardKicker("Moderationskarte", NAVY),
  cardTitle("Bühne frei & Auswertung", NAVY, 30),
  p(t("Eine Person moderiert die ganze Phase. Lies deine Sätze einfach ab und halte den Timer im Blick.", { size: 19, italics: true, color: MUTED }), { after: 60 }),
  p(t("Beim Spielen · Wandreihe → Mittelreihe → Fensterreihe", { bold: true, color: RED, size: 21 }), { after: 30 }),
  ...[
    ["Ansage", "„Die Wandreihe, bitte auf die Bühne. Paten sind die Fensterreihe.“ (Mittelreihe: Paten = Wandreihe · Fensterreihe: Paten = Mittelreihe)"],
    ["Start", "„Version A – bitte!“ Danach: „Besetzung B – Version B, bitte!“ (keine Unterbrechung, höchstens 2 Min. je Version)"],
    ["Danach", "„Applaus! – 30 Sekunden Notizzeit.“ Dann die nächste Reihe aufrufen."],
  ].map(modLine),
  p(t("Bei der Auswertung", { bold: true, color: S.B, size: 21 }), { before: 60, after: 30 }),
  ...[
    ["Schritt 1", "„2 Minuten still: Ergänzt euren Beobachtungsbogen.“"],
    ["Schritt 2", "„3 Minuten in der Gruppe: Einigt euch zu eurer Patenszene auf Kipppunkt, Ursache und die Strategie, die am besten gewirkt hat.“"],
    ["Schritt 3", "„Wandreihe, was habt ihr bei der Mittelreihe beobachtet?“ – ca. 2 Minuten, dann: „Mittelreihe, möchtet ihr etwas ergänzen?“ Danach Mittelreihe über Fensterreihe, Fensterreihe über Wandreihe."],
  ].map(modLine),
], W, 5200);

const exitW = [5103, 5103];
const exitCell = (w) => cell([
  p(t("EXIT-TICKET", { size: 16, bold: true, color: RED }), { after: 60 }),
  p([t("1  ", { bold: true, color: RED, size: 20 }), t("Formuliere um – als Ich-Botschaft mit Bitte:", { size: 20 })], { after: 20 }),
  p(t("„Nie hörst du mir zu!“", { font: "Cambria", size: 24, bold: true, color: NAVY }), { after: 40 }),
  p(t("_____________________________________________", { color: "9AA6C4", size: 20 }), { after: 40 }),
  p(t("_____________________________________________", { color: "9AA6C4", size: 20 }), { after: 40 }),
  p(t("_____________________________________________", { color: "9AA6C4", size: 20 }), { after: 80 }),
  p([t("2  ", { bold: true, color: RED, size: 20 }), t("Woran erkennst du, dass ein Gespräch gerade kippt? Nenne zwei Signale.", { size: 20 })], { after: 40 }),
  p(t("_____________________________________________", { color: "9AA6C4", size: 20 }), { after: 40 }),
  p(t("_____________________________________________", { color: "9AA6C4", size: 20 }), { after: 0 }),
], { w, borders: allBorders(dashed), valign: VerticalAlign.TOP, m: 180 });
const exitGrid = table(exitW, [row([exitCell(exitW[0]), exitCell(exitW[1])], 4600, HeightRule.ATLEAST), row([exitCell(exitW[0]), exitCell(exitW[1])], 4600, HeightRule.ATLEAST)]);

const karten = [
  kicker("Kartenset · Expertenkarten · 2 Sätze pro Gruppe (A einzeln, B–D im Tandem)"),
  expA, gap(160), expB,
  kicker("Kartenset · Expertenkarten · 2 Sätze pro Gruppe (A einzeln, B–D im Tandem)", true),
  expC, gap(160), expD,
  kicker("Kartenset · Rollenkarten · 1 Satz pro Gruppe (3×)", true),
  roleGrid,
  gap(240),
  kicker("Kartenset · Tippkarten · 2–3 Sätze aufs Pult legen"),
  tippGrid,
  gap(240),
  infoBox([[t("Regel für Hilfe: ", { bold: true, color: NAVY }), t("1. Gruppe fragen  →  2. Tippkarte holen  →  3. Material-Person fragt die Lehrkraft (dort gibt es auch Notfall-Ideen für Aufgabe 4).", { size: 21 })]]),
  kicker("Kartenset · Notfall-Ideen · 1 Satz im Umschlag bei der Lehrkraft – nur ausgeben, wenn eine Gruppe nach Tippkarte 1 noch keine Idee hat", true),
  table(sW, sitRows),
  kicker("Kartenset · Moderationskarte · 2× (Moderation + Lehrkraft)", true),
  modCard,
  kicker("Kartenset · Exit-Tickets · je 1 pro Person (7 pro Gruppenumschlag)", true),
  exitGrid,
];

// ================= LEHRERMATERIAL & LÖSUNGEN =================
// ---------- Seite 1: Überblick & Vorbereitung ----------
const druck = [
  ["Arbeitsblatt (6 Seiten)", "21× (1 pro Person)"],
  ["Kartenset S. 1–2: Expertenkarten A–D", "6 Sätze (2 pro Gruppe; B–D werden im Tandem gelesen)"],
  ["Kartenset S. 3: Rollen- und Tippkarten", "Rollen: 3 Sätze · Tipps: 2–3 Sätze fürs Pult"],
  ["Kartenset S. 4: Notfall-Ideen", "1 Satz, in einen Umschlag – bleibt bei Ihnen"],
  ["Kartenset S. 5: Moderationskarte", "2× (Moderation + Sie)"],
  ["Kartenset S. 6: Exit-Tickets", "21× (6 Seiten), je 7 in die Gruppenumschläge"],
  ["Erwartungshorizonte (S. 3–4 dieses Dokuments)", "1× für Sie (Grundlage für die beiden Unterrichtsgespräche)"],
];
const dRows = druck.map(([a, b]) => row([
  cell(p([t(box + "  ", { size: 20 }), t(a, { size: 19, bold: true })], { after: 0 }), { w: 5000 }),
  cell(p(t(b, { size: 19 }), { after: 0 }), { w: W - 5000 }),
], 340));

const vorher = [
  ["Sitzordnung", "Die Klasse sitzt zu Beginn normal in den drei Reihen. Impuls und Partnerarbeit laufen mit dem Sitznachbarn. Ab Minute 15 bilden Wandreihe, Mittelreihe und Fensterreihe je eine Gruppe (7 Personen). Vorher überlegen, wie jede Reihe zusammenrückt; sitzen nicht genau 7 in einer Reihe, 1–2 Personen umsetzen."],
  ["3 Umschläge", "beschriftet „Wandreihe“, „Mittelreihe“, „Fensterreihe“. Inhalt: 4 Rollenkarten, 2 Sätze Expertenkarten A–D, 7 Exit-Tickets. Jeweils geschlossen auf den ersten Tisch der Reihe legen."],
  ["Auf jeden Platz", "das Arbeitsblatt, Seite 1 (M1) oben."],
  ["Aufs Pult", "Tippkarten (Tipp 1 und 2 zu Aufgabe 2 und 4) in beschrifteten Stapeln."],
  ["Bei Ihnen", "Umschlag „Notfall-Ideen“, 2 Moderationskarten, Zettel für Ihre Notizen (gelungene Formulierungen)."],
  ["Technik", "Präsentation auf Folie 1, Timer griffbereit (Handy unter der Dokumentenkamera oder Online-Timer)."],
  ["Tafel", "Rechts anschreiben: Reihenfolge Wand → Mitte → Fenster · Paten: Wand beobachtet Mitte, Mitte → Fenster, Fenster → Wand."],
];
const vhRows = vorher.map(([a, b]) => row([
  cell(p([t(box + "  ", { size: 20 }), t(a, { size: 19, bold: true, color: NAVY })], { after: 0 }), { w: 2900, fill: LIGHT, valign: VerticalAlign.TOP }),
  cell(p(t(b, { size: 19 }), { after: 0 }), { w: W - 2900 }),
], 340));

const lk1 = [
  kicker("Für die Lehrkraft"),
  h1("Kommunikationsstörungen lösen"),
  p([t("Doppelstunde (90 Min.) · Klasse 11 · Vorwissen: Kommunikationsmodelle (Schulz von Thun, Watzlawick) und Kommunikationsstörungen", { size: 20, color: MUTED, italics: true })], { after: 80 }),
  infoBox([
    [t("Lernziele  ", { bold: true, color: NAVY }), t("Die Schülerinnen und Schüler …", { size: 19 })],
    [t("•  benennen in einem Dialog den genauen Zeitpunkt, an dem die Kommunikation kippt, und begründen ihn mit Fachbegriffen,", { size: 19 })],
    [t("•  kennen vier Strategien zur Konfliktlösung (Ich-Botschaft, Empathie, Wunsch/Bitte, Metakommunikation) und wenden sie an,", { size: 19 })],
    [t("•  erfinden eigene Konfliktszenen, spielen sie und werten die Szenen der anderen kriteriengeleitet aus.", { size: 19 })],
    [t("Organisation  ", { bold: true, color: NAVY }), t("21 Schülerinnen und Schüler · Impuls einzeln, dann zu zweit · danach drei Gruppen à 7: Wandreihe, Mittelreihe, Fensterreihe.", { size: 19 })],
  ]),
  h2("1", "Druckliste · am Vortag", NAVY),
  table([5000, W - 5000], dRows),
  h2("2", "Vorbereitung · vor der Stunde bzw. in der Pause", NAVY),
  table([2900, W - 2900], vhRows),
  h2("3", "Grundregeln für Ihre Sprechzeit", NAVY),
  ...[
    ["Sie teilen während der Stunde nichts aus.", " Alles liegt vorher bereit. Ausnahmen: Notfall-Idee und Moderationskarte."],
    ["Folie wechseln, Timer starten, nicht erklären.", " Die Aufträge stehen auf Folie und Arbeitsblatt."],
    ["Zwei kurze Unterrichtsgespräche (5 und 7 Min.)", " sichern die Lösungen, im zweiten entsteht der Merkkasten. Tipp: Wer geantwortet hat, nimmt die nächste Person dran."],
    ["Fragen kommen nur von der Material-Person", " – und erst, nachdem Gruppe und Tippkarte nicht geholfen haben."],
    ["Beim Spielen sitzen Sie im Publikum.", " Eine Schülerin oder ein Schüler moderiert. Sie notieren gelungene Formulierungen."],
  ].map(([a, b]) => p([t("•  ", { bold: true, color: MUTED }), t(a, { bold: true, size: 19 }), t(b, { size: 19 })], { after: 20 })),
];

// ---------- Seite 2: Regieplan ----------
const rgW = [1050, 1900, 3150, 4106];
const regie = [
  ["0'", "Ankommen\nFolie 1 → 2", "setzt sich wie gewohnt in die Reihen, findet M1 auf dem Platz", "Einziger Satz der Einführung: „Heute steht alles auf den Folien – schaut einfach dorthin.“ Folie 2 zeigen, Timer 5 Min. starten."],
  ["0–5'", "Impuls EA → PA\nFolie 2", "liest M1 still und markiert Kipppunkte mit ↯ (3'), vergleicht dann mit dem Sitznachbarn (2')", "Nichts sagen. Nach 3 Min. Handzeichen/Gong für den Partnervergleich."],
  ["5–13'", "Kipppunkte benennen\nFolie 3", "Aufgabe 2 zu zweit", "Folie 3, Timer 8 Min. Herumgehen und zuhören: Welche Paare finden Zeile 3? Die merken Sie sich fürs Gespräch."],
  ["13–18'", "UG 1: Kipppunkte\nFolie 3 → 4", "Paare nennen Zeile + Fachbegriff, ergänzen ihre Tabelle in anderer Farbe", "Impuls: „In welcher Zeile kippt es zum ersten Mal – und warum?“ Wer geantwortet hat, nimmt die nächste Person dran. Ziel: Zeile 3 ★ + 2–3 weitere Kipppunkte (Erwartungshorizont S. 3). Zum Schluss Folie 4 als Sicherung zeigen."],
  ["18–21'", "Umsetzen\nFolie 5", "Wand-, Mittel- und Fensterreihe rücken zu je einer Gruppe zusammen, öffnen ihren Umschlag, verteilen die Rollen", "Folie 5, Timer 3 Min. Nur eingreifen, wenn eine Reihe nicht 7 Personen hat."],
  ["21–38'", "Gruppenpuzzle\nFolie 5", "Expertenkarten verteilen (A einzeln, B–D im Tandem); lesen 4', erklären 7', Dialog umschreiben 6'", "Timer 17 Min. Bei allen drei Gruppen kurz zuhören. Folie 6 (Werkzeugkasten) noch nicht zeigen – die Klasse soll die Strategien gleich selbst zusammentragen."],
  ["38–45'", "UG 2: Version B + Merkkasten\nFolie 5 → 6 → 7", "Die Protokoll-Person jeder Gruppe liest ihre Version ab Zeile 3 vor; alle füllen den Merkkasten (AB S. 4) aus: Formel und bestes Beispiel je Strategie", "Impuls: „Welche Strategie erkennt ihr – und an welcher Formulierung?“ Strategie für Strategie (A → D) sammeln, jeweils ein Beispiel aus den Versionen wählen lassen. Dann Folie 6 zum Vergleich, zuletzt Folie 7 (Musterlösung) als eine mögliche Version."],
  ["45–62'", "Eigene Szene\nFolie 8", "erfindet eine eigene Situation, verteilt Besetzung A und B, füllt den Planungsbogen aus (Merkkasten als Hilfe), probt in den letzten 5'", "Folie 8, Timer 17 Min. Nach 5 Min. Runde: Hat jede Gruppe eine Idee? Wenn nicht: erst Tippkarte, dann Notfall-Idee. Um ca. 57': eine Person als Moderation gewinnen und ihr die Moderationskarte geben."],
  ["62–75'", "Bühne frei\nFolie 9", "Wand → Mitte → Fenster spielen je Version A und B ohne Unterbrechung; das Publikum füllt den Beobachtungsbogen aus", "Folie 9. Ins Publikum setzen, die Moderation leitet. Pro Gruppe ca. 4 Min. Gelungene Formulierungen notieren."],
  ["75–86'", "Auswertung\nFolie 10", "2' still Bogen ergänzen → 3' Gruppe bespricht die Patenszene → Plenum: je Gruppe ca. 2 Min. Rückmeldung, die Spielgruppe ergänzt", "Folie 10. Die Moderation führt durch die Schritte (oder Sie übernehmen Schritt 3). Zum Schluss höchstens 1 Min.: zwei gelungene Formulierungen aus Ihren Notizen vorlesen."],
  ["86–90'", "Exit-Ticket\nFolie 11", "nimmt das Exit-Ticket aus dem Umschlag, bearbeitet es, gibt es an der Tür ab", "Folie 11, Timer 4 Min. An die Tür stellen und einsammeln."],
  ["danach", "Nachbereitung", "", "Exit-Tickets sichten: Wer kann schon eine Ich-Botschaft mit Bitte formulieren? 2–3 gute Beispiele als Einstieg für die nächste Stunde auswählen."],
];
const rgRows = [row([hdr("Zeit", rgW[0]), hdr("Phase · Folie", rgW[1]), hdr("Was die Klasse tut", rgW[2]), hdr("Was Sie tun", rgW[3])], 440)];
regie.forEach(([a, b, c, d]) => {
  const [ph, fo] = b.split("\n");
  rgRows.push(row([
    cell(p(t(a, { bold: true, color: NAVY, size: 19 }), { after: 0 }), { w: rgW[0], fill: LIGHT, valign: VerticalAlign.TOP }),
    cell([p(t(ph, { bold: true, size: 19 }), { after: 10 }), fo ? p(t(fo, { size: 16, color: MUTED }), { after: 0 }) : null].filter(Boolean), { w: rgW[1], fill: LIGHT, valign: VerticalAlign.TOP }),
    cell(p(t(c, { size: 18 }), { after: 0 }), { w: rgW[2], valign: VerticalAlign.TOP }),
    cell(p(t(d, { size: 18 }), { after: 0 }), { w: rgW[3], fill: LIGHT2, valign: VerticalAlign.TOP }),
  ], 500));
});

const lkRegie = [
  kicker("Für die Lehrkraft · Regieplan", true),
  h1("Wann passiert was?"),
  p(t("Minutengenauer Ablauf. Die Folien-Nummern beziehen sich auf die Präsentation.", { size: 20, italics: true, color: MUTED }), { after: 100 }),
  table(rgW, rgRows),
];

const lkW = [800, 2900, 4006, 2500];
const lkRows = [row([hdr("Zeile", lkW[0]), hdr("Was passiert?", lkW[1]), hdr("Warum kippt es hier?", lkW[2]), hdr("Was hätte geholfen?", lkW[3])], 440)];
const kipp = [
  ["3 ★", "„Du hattest eine ganze Woche Zeit.“", "Erster Kipppunkt. Lenas Selbstkundgabe (Ich bin nervös wegen morgen) versteckt sich in einem Vorwurf, also einer Du-Botschaft. Tom hört mit dem Beziehungsohr: „Du hältst mich für unzuverlässig.“", "A Ich-Botschaft, C Bitte mit Uhrzeit"],
  ["4", "„Nicht jeder hat so viel Freizeit wie du.“", "Tom verteidigt sich mit einem Gegenangriff auf der Beziehungsebene statt die Sache zu klären. Beginn der symmetrischen Eskalation: Jeder will „oben“ bleiben.", "B Empathie: Lenas Stress anerkennen"],
  ["5", "„Immer bleibt alles an mir hängen.“", "Verallgemeinerung („immer“). Lenas eigentliches Gefühl (überfordert, allein gelassen) kommt nur als Anklage an. Interpunktion: Beide sehen sich als diejenigen, die nur reagieren.", "A Ich-Botschaft, C Bedürfnis nennen"],
  ["7", "„Super. Genau das hab ich von dir erwartet.“ (verdreht die Augen)", "Ironie: Die Worte sind positiv, Tonfall und Mimik negativ. Das ist eine inkongruente Botschaft (digitale und analoge Ebene widersprechen sich). Zugleich wertet Lena Tom als Person ab.", "D Metakommunikation: Streit benennen, gemeinsames Ziel"],
  ["8–9", "Tom schaut aufs Handy. Lena: „Ist ja auch wichtiger.“", "„Man kann nicht nicht kommunizieren.“ Toms Rückzug ist auch eine Botschaft. Lena deutet ihn als Desinteresse, Tom meint vermutlich: „Ich will nicht mehr streiten.“ Danach ist das Gespräch kaum noch zu retten.", "D „Wie meinst du das?“, B Gefühl ansprechen"],
];
kipp.forEach(([a, b, c, d]) => lkRows.push(row([
  cell(p(t(a, { bold: true, color: RED, size: 22 }), { align: AlignmentType.CENTER, after: 0 }), { w: lkW[0], fill: LIGHT }),
  cell(p(t(b, { font: "Cambria", italics: true, size: 20 }), { after: 0 }), { w: lkW[1] }),
  cell(p(t(c, { size: 19 }), { after: 0 }), { w: lkW[2] }),
  cell(p(t(d, { size: 19, bold: true, color: NAVY }), { after: 0 }), { w: lkW[3], fill: LIGHT2 }),
], 1100, HeightRule.ATLEAST)));

const lk2 = [
  kicker("Erwartungshorizont · Unterrichtsgespräch 1 · Aufgabe 2 (auch Folie 4)", true),
  h1("Kipppunkte in „Das Referat“"),
  p([t("Erwartung: mindestens drei Kipppunkte, darunter Zeile 3. Andere Begründungen sind richtig, wenn sie mit dem Text und einem Modell belegt sind.", { size: 20, italics: true, color: MUTED })], { after: 120 }),
  table(lkW, lkRows),
  gap(100),
  infoBox([
    [t("★ Warum Zeile 3?  ", { bold: true, color: RED }), t("Hier ist noch nichts passiert, das man zurücknehmen müsste. Lena will eigentlich nur Sicherheit für morgen. Hätte sie das als Ich-Botschaft gesagt, wäre Tom gar nicht in die Verteidigung gegangen. Jeder spätere Kipppunkt baut auf dem vorherigen auf – je später man eingreift, desto schwerer wird es.", { size: 20 })],
  ]),
];

const mW = [700, 1100, 6906, 1500];
const muster = [
  ["3", "Lena", "Heute Abend? Ehrlich gesagt macht mich das nervös, weil wir morgen schon dran sind und ich die Folien vorher einmal sehen wollte.", "A"],
  ["4", "Tom", "Verstehe ich. Du willst sicher sein, dass alles passt, oder? Ich hatte diese Woche dreimal Training und hab’s unterschätzt.", "B"],
  ["5", "Lena", "Ja, genau. Ich hab schon die Gliederung gemacht und fühl mich gerade ein bisschen allein mit dem Ganzen.", "A"],
  ["6", "Tom", "Okay, das ist fair. Tut mir leid.", "B"],
  ["7", "Lena", "Kannst du mir die Folien bis 19 Uhr schicken? Dann schaue ich drüber, und wir proben morgen früh vor der ersten Stunde einmal.", "C"],
  ["8", "Tom", "19 Uhr schaffe ich. Und beim nächsten Mal sagen wir uns früher Bescheid, wenn’s eng wird – einverstanden?", "D"],
  ["9", "Lena", "Einverstanden.", ""],
];
const mRows = [row([hdr("Z.", mW[0]), hdr("Wer", mW[1]), hdr("Version B", mW[2]), hdr("Strategie", mW[3])], 440)];
muster.forEach(([z, w, txt, s]) => mRows.push(row([
  cell(p(t(z, { bold: true, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: mW[0], fill: LIGHT }),
  cell(p(t(w, { bold: true, color: NAVY }), { after: 0 }), { w: mW[1], fill: LIGHT }),
  cell(p(t(txt, { font: "Cambria", size: 21 }), { after: 0 }), { w: mW[2] }),
  cell(p(t(s ? s + " " + { A: "Ich-Botschaft", B: "Empathie", C: "Bitte", D: "Meta" }[s] : "", { bold: true, size: 18, color: s ? S[s] : MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: mW[3] }),
], 640, HeightRule.ATLEAST)));

const lk3 = [
  kicker("Erwartungshorizont · Unterrichtsgespräch 2 · Aufgabe 3 (auch Folie 7; Merkkasten = Folie 6)", true),
  h1("So hätte es laufen können"),
  p(t("Eine mögliche Version B. Die Versionen der Gruppen dürfen ganz anders klingen – entscheidend ist, dass sie die Checkliste erfüllen.", { size: 20, italics: true, color: MUTED }), { after: 120 }),
  table(mW, mRows),
  gap(100),
  infoBox([
    [t("Metakommunikation – wenn es schon eskaliert ist  ", { bold: true, color: S.D }), t("Auch nach Zeile 7 ist noch etwas zu retten, z. B. Tom: „Stopp. Wir streiten gerade nur noch. Ich will auch, dass das Referat gut wird – lass uns fünf Minuten überlegen, wie wir das bis morgen schaffen.“", { size: 20 })],
  ]),
  h2("✓", "Erwartungshorizont Exit-Ticket", NAVY),
  p([t("1  ", { bold: true, color: RED }), t("z. B. „Ich fühle mich gerade übergangen, wenn ich dir etwas erzähle und du dabei aufs Handy schaust. Kannst du es kurz weglegen, bis ich fertig bin?“ – Kriterien: Ich-Form, Gefühl, konkreter Anlass statt „nie“, erfüllbare Bitte.", { size: 20 })], { after: 60 }),
  p([t("2  ", { bold: true, color: RED }), t("z. B. Du-Botschaften und Vorwürfe · „immer“ / „nie“ · Gegenangriff statt Antwort · Ironie · es geht um die Person statt um die Sache · Körpersprache (Augenrollen, Wegdrehen, Handy) · lauter werden · Abbruch.", { size: 20 })], { after: 0 }),
];

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const TITLE = "Kommunikationsstörungen lösen";

const ab = new Document({ styles, sections: [{ properties: pageProps, headers: { default: mkHeader(TITLE) }, footers: { default: footer(TITLE) }, children: [...ab1, ...ab2, ...ab3, ...abMerk, ...ab4, ...ab5] }] });
const ks = new Document({ styles, sections: [{ properties: pageProps, footers: { default: footer("Kartenset · bitte an den gestrichelten Linien ausschneiden") }, children: karten }] });
const lk = new Document({ styles, sections: [{ properties: pageProps, footers: { default: footer("Lehrermaterial & Lösungen") }, children: [...lk1, ...lkRegie, ...lk2, ...lk3] }] });

Packer.toBuffer(ab).then((b) => fs.writeFileSync("Arbeitsblatt_Kommunikationsstoerungen_loesen.docx", b));
Packer.toBuffer(ks).then((b) => fs.writeFileSync("Kartenset_Kommunikationsstoerungen_loesen.docx", b));
Packer.toBuffer(lk).then((b) => fs.writeFileSync("Lehrermaterial_und_Loesungen.docx", b));
