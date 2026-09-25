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
  h2("1", "Stummer Einstieg · allein, 3 Min."),
  p([t("Lies den Dialog still. Setze am Rand ein "), t("↯", { bold: true, color: RED }), t(" an jede Stelle, an der das Gespräch deiner Meinung nach "), t("kippt", { bold: true }), t(". Noch nicht reden!")], { after: 60 }),
  p([t("Danach 2 Min. mit deinem Nachbarn: ", { bold: true }), t("Habt ihr dieselben Stellen markiert? Wo seid ihr euch uneinig?")], { after: 0 }),
];

const kW = [800, 2900, 3606, 2900];
const kRows = [row([hdr("Zeile", kW[0]), hdr("Was passiert?", kW[1], NAVY, "Zitat oder Verhalten"), hdr("Warum kippt es hier?", kW[2], NAVY, "Fachbegriff aus den Kommunikationsmodellen"), hdr("Was hätte geholfen?", kW[3], NAVY, "erst nach Aufgabe 3 ausfüllen")], 620)];
for (let i = 0; i < 5; i++) kRows.push(row(kW.map((w, j) => cell(p(t("")), { w, fill: j === 3 ? LIGHT : undefined })), 1250));

const begriffe = ["Sachebene / Beziehungsebene", "Selbstkundgabe", "Appell", "Beziehungsohr", "Du-Botschaft (Vorwurf)", "Verallgemeinerung („immer“, „nie“)", "Ironie = inkongruente Botschaft", "symmetrische Eskalation", "Interpunktion", "„Man kann nicht nicht kommunizieren.“"];

const ab2 = [
  kicker("Aufgabe 2 · Partnerarbeit · 10 Min.", true),
  h1("Kipppunkte benennen"),
  p([t("Ein "), t("Kipppunkt", { bold: true }), t(" ist die Stelle, an der ein Gespräch von der Sache weg und in einen Konflikt rutscht. Oft gibt es mehrere, die aufeinander aufbauen.")], { after: 80 }),
  p([t("a) ", { bold: true }), t("Tragt "), t("mindestens drei", { bold: true }), t(" Kipppunkte aus M1 in die Tabelle ein. Die letzte Spalte bleibt vorerst leer.")], { after: 40 }),
  p([t("b) ", { bold: true }), t("Markiert den Kipppunkt, an dem man das Gespräch "), t("am leichtesten", { bold: true }), t(" hätte retten können, mit einem Stern ★.")], { after: 40 }),
  p([t("c) ", { bold: true }), t("Vergleicht mit dem Lösungsblatt am Pult (Material-Person holt es). Ergänzt, was euch fehlt, "), t("in einer anderen Farbe", { bold: true }), t(".")], { after: 120 }),
  table(kW, kRows),
  gap(100),
  infoBox([
    [t("Fachbegriffe zur Auswahl", { bold: true, color: NAVY })],
    [t(begriffe.join("  ·  "), { size: 20 })],
  ]),
  gap(60),
  p([t("Nicht weitergekommen? ", { bold: true, color: RED, size: 20 }), t("Erst in der Gruppe fragen, dann Tippkarte 2 holen.", { size: 20, italics: true })], { after: 0 }),
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
  kicker("Aufgabe 3 · Gruppenpuzzle · 20 Min.", true),
  h1("Das Gespräch retten"),
  p([t("1. ", { bold: true }), t("Verteilt die Rollenkarten und die Expertenkarten "), t("A–D", { bold: true }), t(". Jede Person liest "), t("allein", { bold: true }), t(" ihre Expertenkarte (5 Min.).")], { after: 40 }),
  p([t("2. ", { bold: true }), t("Reihum erklärt jede Person ihre Strategie in "), t("höchstens 2 Minuten", { bold: true }), t(" – mit dem Beispiel von der Karte.")], { after: 40 }),
  p([t("3. ", { bold: true }), t("Schreibt den Dialog "), t("ab Zeile 3", { bold: true }), t(" neu. Nutzt mindestens drei der vier Strategien und notiert rechts den Buchstaben.")], { after: 40 }),
  p([t("4. ", { bold: true }), t("Prüft eure Version mit der Checkliste. Tragt dann in Aufgabe 2 die letzte Spalte nach.")], { after: 100 }),
  stratLegend,
  gap(100),
  table(rwW, rwRows),
  gap(100),
  p(t("Checkliste Version B", { font: "Cambria", bold: true, size: 24, color: NAVY }), { after: 60, keepNext: true }),
  ...check.map((c) => p([t(box + "  ", { size: 24 }), t(c, { size: 21 })], { after: 30 })),
];

// Planungsbogen
const pl = (label, h, hint) => row([
  cell([p(t(label, { bold: true, color: NAVY, size: 20 }), { after: 0 }), hint ? p(t(hint, { size: 16, italics: true, color: MUTED }), { after: 0 }) : null].filter(Boolean), { w: 2600, fill: LIGHT, valign: VerticalAlign.TOP }),
  cell(p(t("")), { w: W - 2600 }),
], h, HeightRule.ATLEAST);

const ab4 = [
  kicker("Aufgabe 4 · Gruppe · 20 Min.", true),
  h1("Eure eigene Szene"),
  p([t("Schreibt eine Alltagsszene in "), t("zwei Versionen", { bold: true }), t(". Beide beginnen gleich. "), t("Version A", { bold: true, color: RED }), t(" kippt und eskaliert. In "), t("Version B", { bold: true, color: S.B }), t(" wird genau am Kipppunkt anders reagiert.")], { after: 60 }),
  p([t("Keine Idee? ", { bold: true }), t("Die Material-Person zieht eine Situationskarte.", { italics: true })], { after: 60 }),
  infoBox([[t("Spielregeln: ", { bold: true, color: NAVY }), t("2–4 Rollen · jede Version höchstens 2 Minuten · Text darf auf Karteikarten mitgenommen werden · ", { size: 20 }), t("der Kipppunkt muss für das Publikum erkennbar sein.", { size: 20, bold: true })]]),
  gap(100),
  table([2600, W - 2600], [
    pl("Situation", 900, "Wo? Wer? Worum geht es?"),
    pl("Rollen", 700, "Figur → gespielt von"),
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

const bW = [1100, 2700, 2700, 2506, 1200];
const bRows = [row([hdr("Gruppe", bW[0]), hdr("Kipppunkt", bW[1], NAVY, "Wann genau? Zitat"), hdr("Ursache", bW[2], NAVY, "Fachbegriff"), hdr("Strategie in Version B", bW[3], NAVY, "A · B · C · D"), hdr("Wirkt?", bW[4], NAVY, "+  o  –")], 620)];
for (let i = 0; i < 7; i++) bRows.push(row(bW.map((w, j) => cell(p(t(j === 0 ? String(i + 1) : "", { bold: true, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w, fill: j === 0 ? LIGHT : undefined })), 1350));

const ab5 = [
  kicker("Aufgabe 5 · Bühne frei · 30 Min.", true),
  h1("Beobachtungsbogen"),
  p([t("So läuft es: ", { bold: true, color: NAVY }), t("Eine Gruppe spielt "), t("Version A", { bold: true, color: RED }), t(". Sobald du merkst, dass das Gespräch kippt, hältst du deine "), t("STOPP-Karte", { bold: true, color: RED }), t(" hoch. Die Moderation ruft "), t("„Freeze!“", { bold: true }), t(" – die Spielenden erstarren. Eine Person mit Karte sagt in einem Satz, was gerade passiert ist. Danach geht es weiter, direkt im Anschluss folgt "), t("Version B", { bold: true, color: S.B }), t(".")], { after: 80 }),
  p([t("Notiere während des Spiels in Stichpunkten. Die Moderation übernimmt immer die Gruppe, die als "), t("nächste", { bold: true }), t(" spielt.")], { after: 120 }),
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
  ["Regie", "Du achtest darauf, dass alle zu Wort kommen. Beim Spielen entscheidest du: Wer steht wo, wer beginnt?"],
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
  ["Tipp 1 · Aufgabe 4", "Denkt an Situationen aus eurem Alltag: Familie, Freunde, Schule, Verein, Job, Chat. Wer will was – und was stört wen?"],
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
      p(t("SITUATIONSKARTE " + (i + j + 1), { size: 14, bold: true, color: MUTED }), { after: 20 }),
      p(t(n, { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 60 }),
      p(t(d, { size: 21 }), { after: 0 }),
    ], { w: sW[j], borders: allBorders(dashed), valign: VerticalAlign.TOP, m: 200 });
  }), 2500, HeightRule.ATLEAST));
}

const modCard = cutCard([
  cardKicker("Moderationskarte", NAVY),
  cardTitle("So moderierst du", NAVY, 30),
  p(t("Die Gruppe, die als nächste spielt, übernimmt die Moderation. Lies deine Sätze einfach ab.", { size: 20, italics: true, color: MUTED }), { after: 80 }),
  ...[
    ["Ansage", "„Es spielt Gruppe …, Version A. Stopp-Karten bereit!“"],
    ["Bei Stopp-Karte", "„Freeze!“ – dann auf eine Person mit Karte zeigen: „Was ist gerade passiert?“ (höchstens ein Satz)"],
    ["Weiter", "„Und weiter!“ – Szene läuft bis zum Ende."],
    ["Wechsel", "„Jetzt Version B – achtet darauf, was am Kipppunkt anders ist.“"],
    ["Abschluss", "„Welche Strategie habt ihr erkannt?“ – zwei Meldungen, dann Applaus und nächste Gruppe."],
  ].map(([a, b]) => p([t(a + ":  ", { bold: true, color: NAVY, size: 21 }), t(b, { size: 21 })], { after: 60 })),
], W, 4200);

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

const stopW = [5103, 5103];
const stopCell = (w) => cell([
  p(t("✋", { size: 72, color: "FFFFFF" }), { align: AlignmentType.CENTER, after: 0 }),
  p(t("STOPP", { font: "Cambria", size: 96, bold: true, color: "FFFFFF" }), { align: AlignmentType.CENTER, after: 0 }),
  p(t("Hier kippt es!", { size: 28, bold: true, color: "FFFFFF" }), { align: AlignmentType.CENTER, after: 0 }),
], { w, fill: RED, borders: allBorders(solid("FFFFFF", 24)) });
const stopGrid = table(stopW, [row([stopCell(stopW[0]), stopCell(stopW[1])], 7000), row([stopCell(stopW[0]), stopCell(stopW[1])], 7000)]);

const karten = [
  kicker("Kartenset · Expertenkarten · 1 Satz pro Gruppe"),
  expA, gap(160), expB,
  kicker("Kartenset · Expertenkarten · 1 Satz pro Gruppe", true),
  expC, gap(160), expD,
  kicker("Kartenset · Rollenkarten · 1 Satz pro Gruppe", true),
  roleGrid,
  gap(240),
  kicker("Kartenset · Tippkarten · 2–3 Sätze aufs Pult legen"),
  tippGrid,
  gap(240),
  infoBox([[t("Regel für Hilfe: ", { bold: true, color: NAVY }), t("1. Gruppe fragen  →  2. Tippkarte holen  →  3. Material-Person fragt die Lehrkraft.", { size: 21 })]]),
  kicker("Kartenset · Situationskarten · 1 Satz aufs Pult legen", true),
  table(sW, sitRows),
  kicker("Kartenset · Moderationskarte · ca. 7×", true),
  modCard,
  gap(200),
  kicker("Kartenset · Exit-Tickets · je 1 pro Person"),
  exitGrid,
  kicker("Kartenset · Stopp-Karten · je 1 pro Person (farbig drucken, ggf. laminieren)", true),
  stopGrid,
];

// ================= LEHRERMATERIAL & LÖSUNGEN =================
const vW = [1300, 2500, 4406, 2000];
const verlauf = [
  ["0–5'", "Stummer Einstieg", "M1 liegt auf den Tischen, Folie 2 läuft. Allein lesen, Kipppunkte mit ↯ markieren, dann kurz mit dem Nachbarn vergleichen.", "Einzel → Partner"],
  ["5–15'", "Kipppunkte benennen", "Aufgabe 2 (Folie 3). Selbstkontrolle mit Lösungsblatt S. 2 am Pult. Ergänzungen in anderer Farbe.", "Partner"],
  ["15–35'", "Gruppenpuzzle Strategien", "Vierergruppen, Rollen- und Expertenkarten A–D. Lesen (5'), erklären (4 × 2'), Dialog ab Z. 3 umschreiben, Checkliste (Folie 4–5).", "Gruppe"],
  ["35–55'", "Eigene Szene", "Planungsbogen Aufgabe 4, Situationskarten bei Bedarf. Kurz proben (Folie 6).", "Gruppe"],
  ["55–85'", "Bühne frei", "Version A mit Stopp-Karten und Freeze, dann Version B. Schülermoderation, Beobachtungsbogen (Folie 7). Zeit knapp? Nur 3–4 Gruppen, Rest in der nächsten Stunde.", "Plenum, schülergeleitet"],
  ["85–90'", "Exit-Ticket", "Einsammeln an der Tür (Folie 8). Auswertung für die nächste Stunde.", "Einzel"],
];
const vRows = [row([hdr("Zeit", vW[0]), hdr("Phase", vW[1]), hdr("Was passiert", vW[2]), hdr("Sozialform", vW[3])], 440)];
verlauf.forEach(([a, b, c, d]) => vRows.push(row([
  cell(p(t(a, { bold: true, color: NAVY, size: 20 }), { after: 0 }), { w: vW[0], fill: LIGHT }),
  cell(p(t(b, { bold: true, size: 20 }), { after: 0 }), { w: vW[1], fill: LIGHT }),
  cell(p(t(c, { size: 19 }), { after: 0 }), { w: vW[2] }),
  cell(p(t(d, { size: 19, color: MUTED }), { after: 0 }), { w: vW[3] }),
], 560, HeightRule.ATLEAST)));

const druck = [
  ["Arbeitsblatt (5 Seiten)", "1× pro Person"],
  ["Kartenset S. 1–3: Experten-, Rollen-, Tippkarten", "1× pro Gruppe (Tippkarten: 2–3 Sätze fürs Pult)"],
  ["Kartenset S. 4: Situationskarten", "1× fürs Pult"],
  ["Kartenset S. 5: Moderationskarte und Exit-Tickets", "Moderation 7×, Exit-Tickets 1 pro Person"],
  ["Kartenset S. 6: Stopp-Karten", "1 pro Person, farbig, am besten laminiert (wiederverwendbar)"],
  ["Lösungsblatt S. 2–3 dieses Dokuments", "2–3× fürs Pult (Selbstkontrolle)"],
];
const dRows = druck.map(([a, b]) => row([
  cell(p([t(box + "  ", { size: 22 }), t(a, { size: 20, bold: true })], { after: 0 }), { w: 5600 }),
  cell(p(t(b, { size: 20 }), { after: 0 }), { w: W - 5600 }),
], 340, HeightRule.ATLEAST));

const lk1 = [
  kicker("Für die Lehrkraft"),
  h1("Kommunikationsstörungen lösen"),
  p([t("Doppelstunde (90 Min.) · Klasse 11 · Vorwissen: Kommunikationsmodelle (Schulz von Thun, Watzlawick) und Kommunikationsstörungen", { size: 21, color: MUTED, italics: true })], { after: 80 }),
  infoBox([
    [t("Lernziele  ", { bold: true, color: NAVY }), t("Die Schülerinnen und Schüler …", { size: 19 })],
    [t("•  benennen in einem Dialog den genauen Zeitpunkt, an dem die Kommunikation kippt, und begründen ihn mit Fachbegriffen,", { size: 19 })],
    [t("•  kennen vier Strategien zur Konfliktlösung (Ich-Botschaft, Empathie, Wunsch/Bitte, Metakommunikation) und wenden sie an,", { size: 19 })],
    [t("•  entwickeln eigene Konfliktszenen, spielen sie und beobachten die Szenen der anderen kriteriengeleitet.", { size: 19 })],
  ]),
  h2("1", "Verlaufsplan", NAVY),
  table(vW, vRows),
  h2("2", "Druckliste", NAVY),
  table([5600, W - 5600], dRows),
  h2("3", "So bleibt Ihre Sprechzeit gering", NAVY),
  ...[
    ["Aufträge stehen auf den Folien", " – mit Zeitangabe und sichtbarem Timer. Folie auflegen statt erklären."],
    ["Rollenkarten", " – die Gruppen organisieren sich selbst. Nur die Material-Person spricht Sie an."],
    ["Tippkarten und Lösungsblatt am Pult", " ersetzen das Nachfragen und die Besprechung im Plenum."],
    ["Stopp-Karten und Schülermoderation", " ersetzen das Auswertungsgespräch. Sie beobachten und notieren nur."],
    ["Ihre Rolle:", " beobachten, gezielt helfen, gelungene Formulierungen für die nächste Stunde notieren."],
  ].map(([a, b]) => p([t("•  ", { bold: true, color: MUTED }), t(a, { bold: true, size: 19 }), t(b, { size: 19 })], { after: 20 })),
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
  kicker("Lösungsblatt · Selbstkontrolle Aufgabe 2 · darf ans Pult", true),
  h1("Kipppunkte in „Das Referat“"),
  p([t("Mindestens drei davon solltet ihr gefunden haben. Andere Begründungen sind richtig, wenn sie mit dem Text und einem Modell belegt sind.", { size: 20, italics: true, color: MUTED })], { after: 120 }),
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
  kicker("Lösungsblatt · Musterlösung Aufgabe 3 · darf ans Pult", true),
  h1("So hätte es laufen können"),
  p(t("Eine mögliche Version B. Eure darf ganz anders klingen – wichtig ist, dass sie die Checkliste erfüllt.", { size: 20, italics: true, color: MUTED }), { after: 120 }),
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

const ab = new Document({ styles, sections: [{ properties: pageProps, headers: { default: mkHeader(TITLE) }, footers: { default: footer(TITLE) }, children: [...ab1, ...ab2, ...ab3, ...ab4, ...ab5] }] });
const ks = new Document({ styles, sections: [{ properties: pageProps, footers: { default: footer("Kartenset · bitte an den gestrichelten Linien ausschneiden") }, children: karten }] });
const lk = new Document({ styles, sections: [{ properties: pageProps, footers: { default: footer("Lehrermaterial & Lösungen") }, children: [...lk1, ...lk2, ...lk3] }] });

Packer.toBuffer(ab).then((b) => fs.writeFileSync("Arbeitsblatt_Kommunikationsstoerungen_loesen.docx", b));
Packer.toBuffer(ks).then((b) => fs.writeFileSync("Kartenset_Kommunikationsstoerungen_loesen.docx", b));
Packer.toBuffer(lk).then((b) => fs.writeFileSync("Lehrermaterial_und_Loesungen.docx", b));
