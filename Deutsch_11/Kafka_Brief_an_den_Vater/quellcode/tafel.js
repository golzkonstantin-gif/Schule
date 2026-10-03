const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
// Tafel: dunkelgrün; Kreidefarben je Reihe und für die Synthese
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA", DIM = "C9C3A8";
const K = { wand: "8EC3F0", mitte: "F5D76E", fenster: "F29A8E", komp: "9BDB9B", gen: "D7B4F0" };
const W = 15138; // Inhaltsbreite A4 quer, 1,5 cm Rand

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 100 }, keepNext: o.keepNext });
const h1 = (text) => p(t(text, { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 60 });
const kicker = (text, pb) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20, pb });
const h2 = (label, text) => p([t(label + "  ", { font: "Cambria", size: 24, bold: true, color: RED }), t(text, { font: "Cambria", size: 24, bold: true, color: NAVY })], { before: 140, after: 60, keepNext: true });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || allBorders(solid("B9C6E8", 6)),
    verticalAlign: o.valign || VerticalAlign.TOP,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 110, right: o.ml ?? 110 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 19 }), { align: AlignmentType.CENTER, after: 0 }), { w, fill, valign: VerticalAlign.CENTER });
const infoBox = (paras, fill = LIGHT2, w = W) => table([w], [row([cell(paras.map((r) => p(r, { after: 40 })), { w, fill, borders: allBorders(none), m: 110 })])]);

// Kreide-Helfer
const ck = (text, o = {}) => t(text, { font: o.font || "Calibri", size: o.size || 18, bold: o.bold, italics: o.italics, color: o.color || CHALK });
const phase = (n) => ck(["", "①", "②", "③", "④"][n] + " ", { size: 18, color: DIM, bold: true });
const chalkBorder = allBorders(solid("6E8C7A", 6));

// ================= Inhalte der Matrix (Tafel) =================
const auszuege = [
  {
    key: "wand", reihe: "Wandreihe", titel: "Pawlatsche und Badekabine", seite: "S. 3",
    vater: ["trägt das Kind nachts „im Hemd“ auf die Pawlatsche", "macht beim Baden „in guter Absicht“ Schwimmbewegungen vor"],
    modell: "Appell ohne Erklärung („Sei still!“) – beim Kind kommt nur die Beziehungsbotschaft an. Gute Absicht ≠ Wirkung (Sender ≠ Empfänger).",
    obenUnten: "„Ich mager, schwach, schmal, Du stark, groß, breit.“ – „die letzte Instanz“",
    wirkung: "„… daß ich also ein solches Nichts für ihn war.“ – „tiefe Beschämung“",
  },
  {
    key: "mitte", reihe: "Mittelreihe", titel: "Drohen und Ironie", seite: "S. 6",
    vater: ["„ich zerreiße Dich wie einen Fisch“", "„Das ist Dir wohl schon zu viel?“", "„Das kann man vom Herrn Sohn natürlich nicht haben“"],
    modell: "Drohung = Appell mit Beziehungsbotschaft (Macht). Ironie = inkongruente Botschaft (Worte ≠ Lachen/Gesicht). Kritik in der 3. Person: Empfänger wird übergangen.",
    obenUnten: "Vater straft, „ehe man noch wußte, daß man etwas Schlechtes getan hatte“",
    wirkung: "„Ich verlor das Vertrauen zu eigenem Tun.“ – fragt nur noch die Mutter: „Wie geht es dem Vater?“",
  },
  {
    key: "fenster", reihe: "Fensterreihe", titel: "Tischregeln und drei Welten", seite: "S. 5",
    vater: ["„Knochen durfte man nicht zerreißen, Du ja.“", "„Essig durfte man nicht schlürfen, Du ja.“", "schneidet sich bei Tisch die Nägel"],
    modell: "Regeln (digital) und Verhalten (analog) widersprechen sich → Vater als Sender unglaubwürdig; Regeln sind reine Beziehungsbotschaft: „Ich darf, du nicht.“",
    obenUnten: "drei Welten: „der Sklave“ – der Vater „mit der Regierung“ – die freien „übrigen Leute“",
    wirkung: "„Ich war immerfort in Schande“ – gehorchen, trotzen oder scheitern: alles Schande",
  },
];

// ================= Seite 1: Tafelbild im Endstand =================
const cw = [2500, 3300, 3700, 2900, 2738]; // Summe = W
const matrix = table(cw, [
  row([
    cell(p([phase(2), ck("Auszug", { bold: true })], { after: 0 }), { w: cw[0], fill: BOARD2, borders: chalkBorder }),
    cell(p(ck("Was tut / sagt der Vater? (Zitat)", { bold: true }), { after: 0 }), { w: cw[1], fill: BOARD2, borders: chalkBorder }),
    cell(p(ck("Modell (Seite / Axiom)", { bold: true }), { after: 0 }), { w: cw[2], fill: BOARD2, borders: chalkBorder }),
    cell(p(ck("oben ↔ unten (komplementär)", { bold: true }), { after: 0 }), { w: cw[3], fill: BOARD2, borders: chalkBorder }),
    cell(p(ck("Wirkung auf den Sohn", { bold: true }), { after: 0 }), { w: cw[4], fill: BOARD2, borders: chalkBorder }),
  ], 360),
  ...auszuege.map((a) => row([
    cell([p(ck(a.reihe, { bold: true, size: 20, color: K[a.key] }), { after: 10 }), p(ck(a.titel + " (" + a.seite + ")", { size: 16, italics: true, color: K[a.key] }), { after: 0 })], { w: cw[0], fill: BOARD, borders: chalkBorder }),
    cell(a.vater.map((v) => p(ck(v, { size: 16, italics: true, color: K[a.key] }), { after: 20 })), { w: cw[1], fill: BOARD, borders: chalkBorder }),
    cell(p(ck(a.modell, { size: 16 }), { after: 0 }), { w: cw[2], fill: BOARD, borders: chalkBorder }),
    cell(p(ck(a.obenUnten, { size: 16 }), { after: 0 }), { w: cw[3], fill: BOARD, borders: chalkBorder }),
    cell(p(ck(a.wirkung, { size: 16 }), { after: 0 }), { w: cw[4], fill: BOARD, borders: chalkBorder }),
  ], 1050)),
]);

const sw = [5000, 5000, W - 10000 - 400];
const synthese = table(sw, [row([
  cell([
    p([phase(3), ck("Starre Komplementarität", { bold: true, size: 19, color: K.komp })], { after: 30 }),
    p(ck("Vater immer oben: urteilt, droht, macht Regeln nur für den Sohn", { size: 16 }), { after: 10 }),
    p(ck("Sohn immer unten: gehorcht, schämt sich, verstummt", { size: 16 }), { after: 10 }),
    p(ck("→ kein Rollenwechsel, kein Gespräch auf Augenhöhe", { size: 16, bold: true, color: K.komp }), { after: 0 }),
  ], { w: sw[0], fill: BOARD2, borders: chalkBorder, m: 90 }),
  cell([
    p([phase(3), ck("Generationenkonflikt", { bold: true, size: 19, color: K.gen })], { after: 30 }),
    p(ck("„Schon mit sieben Jahren mußte ich mit dem Karren durch die Dörfer fahren.“", { size: 16, italics: true, color: K.gen }), { after: 10 }),
    p(ck("Vater: Not, Arbeit, Kraft, Dankbarkeit – Sohn: Wohlstand, Bücher, Empfindsamkeit", { size: 16 }), { after: 10 }),
    p(ck("→ Selbstbeklagung = versteckter Appell: „Sei dankbar!“", { size: 16, bold: true, color: K.gen }), { after: 0 }),
  ], { w: sw[1], fill: BOARD2, borders: chalkBorder, m: 90 }),
  cell([
    p([phase(4), ck("Umschreiben: Hätte das gereicht?", { bold: true, size: 19, color: "F5D76E" })], { after: 30 }),
    p(ck("Vater formuliert um → das Gespräch ändert sich sofort", { size: 16 }), { after: 10 }),
    p(ck("Sohn formuliert um → prallt am Machtgefälle ab", { size: 16 }), { after: 10 }),
    p(ck("⇒ Techniken brauchen Augenhöhe. Der Stärkere muss anfangen.", { size: 17, bold: true, color: "F5D76E" }), { after: 0 }),
  ], { w: sw[2], fill: BOARD2, borders: chalkBorder, m: 90 }),
])]);

const board = table([W], [
  row([cell([
    p([phase(1), ck("Franz Kafka: „Brief an den Vater“ (1919)", { font: "Cambria", size: 30, bold: true })], { after: 50 }),
    p([ck("„Du hast mich letzthin einmal gefragt, warum ich behaupte, ich hätte Furcht vor Dir. Ich wußte Dir, wie gewöhnlich, nichts zu antworten …“", { size: 19, italics: true })], { after: 30 }),
    p([ck("Leitfrage: ", { bold: true, size: 20, color: "F5D76E" }), ck("Warum schreibt Kafka seinem Vater, statt mit ihm zu reden?", { size: 20 }), ck("      ← ④ „… weil ich vor Dir weder denken noch reden konnte.“", { size: 17, italics: true, color: DIM })], { after: 0 }),
  ], { w: W, fill: BOARD, borders: allBorders(none), m: 120, ml: 200 })]),
  row([cell([matrix], { w: W, fill: BOARD, borders: allBorders(none), m: 40, ml: 200 })]),
  row([cell([synthese], { w: W, fill: BOARD, borders: allBorders(none), m: 80, ml: 200 })]),
]);

const page1 = [
  kicker("Tafelscript · Deutsch 11 · Kommunikation in Literatur · Klausurvorbereitung"),
  h1("Tafelbild im Endstand"),
  p([t("①–④ zeigen, ", { size: 19, color: MUTED }), t("wann", { size: 19, color: MUTED, bold: true }), t(" was an die Tafel kommt (Ablauf auf Seite 2). Kreidefarben: Wandreihe blau · Mittelreihe gelb · Fensterreihe rot · Komplementarität grün · Generationenkonflikt violett. Zitate in Originalschreibung (daß, mußte).", { size: 19, color: MUTED })], { after: 80 }),
  board,
];

// ================= Seite 2: Ablauf – wann kommt was an die Tafel? =================
const aW = [1100, 2300, 4200, 4838, 2700];
const ablauf = [
  ["0–15'", "", "Diktat + Selbstkontrolle", "– (Whiteboard: Diktattext)", "Diktat nach der Vorlesefassung. Danach Whiteboard-Folien zeigen, Fehler farbig anstreichen und zählen lassen. Keine Besprechung.", "schreiben mit, kontrollieren selbst"],
  ["15–20'", "①", "Einstieg: der Briefanfang", "Überschrift, Zitat des Briefanfangs, Leitfrage", "Briefanfang vorlesen oder zeigen. Impuls: „Warum schreibt ein 36-Jähriger seinem Vater einen Brief, statt mit ihm zu reden?“ Zwei, drei Vermutungen sammeln, dann die Leitfrage anschreiben. Die Antwort bleibt offen bis ④.", "äußern Vermutungen"],
  ["20–48'", "", "Einzelarbeit, dann Reihengruppe", "Leere Matrix vorbereiten: Spaltenköpfe, Reihen links in den Kreidefarben", "Jede Reihe bearbeitet ihren Auszug (Arbeitsblatt). Herumgehen, bei Verständnisfragen zum Text helfen. In der Gruppenphase: Sprecher bestimmen lassen, schwache Belege merken.", "füllen die Matrix für den eigenen Auszug; bündeln in der Reihe die drei stärksten Belege"],
  ["48–63'", "②", "Vorstellung an der Tafel", "Matrix zeilenweise füllen, Zitate in der Farbe der Reihe", "Je Reihe ca. 4 Min. Sie schreiben mit, korrigieren und präzisieren: „Welche Seite der Nachricht ist das genau?“ – „Wo stehen hier oben und unten?“ Die anderen Reihen füllen ihre Matrix mit.", "stellen vor bzw. schreiben mit"],
  ["63–68'", "③", "Synthese", "Kästen „Starre Komplementarität“ und „Generationenkonflikt“; Vater-Zitate zur Jugend auf dem Whiteboard", "Impuls 1: „Was haben alle drei Szenen gemeinsam?“ → starre Komplementarität. Impuls 2: Whiteboard mit den Sätzen des Vaters zu seiner Jugend: „Wie begründet der Vater seine Härte?“ → Generationenkonflikt, Selbstbeklagung als versteckter Appell.", "erkennen das Muster über alle Auszüge"],
  ["68–80'", "", "Umschreiben (allein)", "–", "Jede und jeder schreibt den Kipppunkt des eigenen Auszugs mit mindestens zwei Techniken um (Buchstaben A–D notieren). Wahl: als Vater oder als Sohn. Herumgehen, je eine Vater- und eine Sohn-Version vormerken.", "schreiben um"],
  ["80–87'", "④", "Vorlesen und Diskussion", "Kasten ④ rechts unten", "Vorgemerkte Versionen vorlesen lassen. Impuls: „Hätte das gereicht? Wer müsste sich ändern?“ – dann: „Generationenkonflikt oder Kommunikationsproblem?“ Ergebnis: Techniken brauchen Augenhöhe.", "lesen vor, diskutieren"],
  ["87–90'", "④", "Abschluss", "Pfeil von ④ zurück zur Leitfrage, Zitat ergänzen", "Die Stelle zur Widerrede vorlesen (S. 5): „… schließlich schwieg ich, zuerst vielleicht aus Trotz, dann, weil ich vor Dir weder denken noch reden konnte.“ Damit ist die Leitfrage beantwortet.", "hören zu"],
];
const aRows = [row([hdr("Zeit", aW[0]), hdr("Phase", aW[1]), hdr("Was an die Tafel kommt", aW[2]), hdr("Ihre Impulse", aW[3]), hdr("Klasse", aW[4])], 380)];
ablauf.forEach(([z, n, ph, tafel, imp, sus]) => aRows.push(row([
  cell(p(t(z, { bold: true, size: 19, color: NAVY }), { after: 0 }), { w: aW[0], fill: LIGHT }),
  cell([p([n ? t(n + " ", { bold: true, size: 22, color: RED }) : t(""), t(ph, { bold: true, size: 18, color: NAVY })], { after: 0 })], { w: aW[1], fill: LIGHT }),
  cell(p(t(tafel, { size: 17 }), { after: 0 }), { w: aW[2] }),
  cell(p(t(imp, { size: 17 }), { after: 0 }), { w: aW[3], fill: LIGHT2 }),
  cell(p(t(sus, { size: 17, color: MUTED }), { after: 0 }), { w: aW[4] }),
], 420)));

const page2 = [
  kicker("Tafelscript · Ablauf der Doppelstunde", true),
  h1("Wann kommt was an die Tafel?"),
  table(aW, aRows),
  p(t("Seitenangaben beziehen sich auf die PDF-Ausgabe von DigBib.Org. Gruppen = Sitzreihen; keine Partnerarbeit, kein Gruppenpuzzle (schwerer Text).", { size: 17, italics: true, color: MUTED }), { before: 80, after: 0 }),
];

// ================= Seite 3: Erwartungshorizont =================
const eW = [2400, 6369, 6369];
const erw = [
  ["wand", "Wandreihe\nPawlatsche und Badekabine",
    "Komplementär: Der Vater handelt als „letzte Instanz“, ohne zu erklären; das Kind versteht den Zusammenhang zwischen „sinnlosem Ums-Wasser-Bitten“ und dem „Hinausgetragenwerden“ nicht. Es empfängt nur die Beziehungsbotschaft: Ich bin ein Nichts. In der Kabine wird das Gefälle körperlich sichtbar. Missverständnis: Der Vater macht die Schwimmbewegungen „in guter Absicht“ vor – beim Sohn kommt „tiefe Beschämung“ an.",
    "Vater (nachts): „Du kannst nicht schlafen, oder? (B) Ich bin sehr müde und brauche Ruhe, weil ich früh ins Geschäft muss. (A) Ich bringe dir ein Glas Wasser, und dann schlafen wir beide – einverstanden? (C)“"],
  ["mitte", "Mittelreihe\nDrohen und Ironie",
    "Drohung („ich zerreiße Dich wie einen Fisch“): Appell mit maximaler Machtbotschaft. Ironie („Das ist Dir wohl schon zu viel?“): Wortlaut und „böses Lachen und böses Gesicht“ widersprechen sich (inkongruent); Strafe vor der Erklärung. Kritik über die Mutter in der 3. Person: Der eigentliche Empfänger wird übergangen – das Kind übernimmt das Muster („Wie geht es dem Vater?“).",
    "Vater: „Ich ärgere mich, wenn die Aufgabe liegen bleibt, weil ich mich auf dich verlassen will. (A) Kannst du sie bis heute Abend erledigen? (C) Wenn sie dir zu schwer ist, sag es mir direkt.“"],
  ["fenster", "Fensterreihe\nTischregeln und drei Welten",
    "Die Regeln (Knochen, Essig, Nägel) gelten nur für den Sohn; das Verhalten des Vaters widerspricht ihnen. Dadurch wird der Vater als Sender unglaubwürdig, und jede Regel wird zur reinen Beziehungsbotschaft („Ich darf, du nicht“). Die „drei Welten“ beschreiben starre Komplementarität als Weltordnung; der Sohn sitzt in jeder Reaktion in der „Schande“ fest.",
    "Sohn: „Vater, mir fällt auf, dass manche Regeln nur für mich gelten. (D) Das macht mich unsicher, weil ich nicht weiß, wonach ich mich richten soll. (A) Können wir die Regeln gemeinsam festlegen? (C)“ – Diskussion: Würde der Vater zuhören?"],
];
const eRows = [row([hdr("Auszug", eW[0]), hdr("Erwartete Analyse", eW[1]), hdr("Mögliche Umschreibung (Buchstaben = Techniken)", eW[2])], 380)];
erw.forEach(([k, a, b, c]) => {
  const [r, ti] = a.split("\n");
  eRows.push(row([
    cell([p(t(r, { bold: true, size: 19, color: NAVY }), { after: 10 }), p(t(ti, { size: 16, italics: true, color: MUTED }), { after: 0 })], { w: eW[0], fill: LIGHT }),
    cell(p(t(b, { size: 17 }), { after: 0 }), { w: eW[1] }),
    cell(p(t(c, { size: 17, italics: true }), { after: 0 }), { w: eW[2], fill: LIGHT2 }),
  ], 900));
});

const page3 = [
  kicker("Tafelscript · Erwartungshorizont", true),
  h1("Was in der Matrix und beim Umschreiben herauskommen sollte"),
  table(eW, eRows),
  h2("④", "Diskussion: Generationenkonflikt oder Kommunikationsproblem?"),
  infoBox([
    [t("Generationenkonflikt: ", { bold: true, color: NAVY, size: 18 }), t("unterschiedliche Lebenserfahrung (Not gegen Wohlstand) und Werte (Kraft, Arbeit, Dankbarkeit gegen Empfindsamkeit, Bücher); der Vater misst den Sohn an seiner eigenen Jugend.", { size: 18 })],
    [t("Kommunikationsproblem: ", { bold: true, color: NAVY, size: 18 }), t("Drohung, Ironie und widersprüchliche Regeln machen ein Gespräch unmöglich; Widerrede ist verboten. Der Konflikt könnte besprochen werden – wird es aber nicht.", { size: 18 })],
    [t("Zuspitzung: ", { bold: true, color: RED, size: 18 }), t("Beides hängt zusammen – der Generationenkonflikt wird erst durch die starre Komplementarität unlösbar. Der Brief ist Kafkas Versuch der Metakommunikation; er erreicht den Vater nie.", { size: 18 })],
  ]),
  p(t("Optional bei Zeit: Am Ende des Briefs lässt Kafka den Vater selbst antworten und wirft sich „Schmarotzertum“ vor (S. 20) – Perspektivwechsel als Diskussionsimpuls: Ist der Brief fair?", { size: 17, italics: true, color: MUTED }), { before: 80, after: 0 }),
];

// ================= Dokument =================
const pageProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 600, left: 850, right: 850, footer: 350 } } };
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelscript · Kafka, Brief an den Vater · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ properties: pageProps, footers: { default: footer }, children: [...page1, ...page2, ...page3] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelscript_Kafka_Brief_an_den_Vater.docx", b));
