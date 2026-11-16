// Hilfeblatt „Das Urteil schreiben“ zur Probeklausur (Aufgabe 4), A4 doppelseitig, tintensparend (keine Flächen, Farbe nur in Schrift und Linien).
// Vorderseite: Frage klären, erste Einschätzung, Argumente sortieren (Wissen über die Reformen anwenden). Rückseite: Bauplan mit Satzanfängen zum Abhaken.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, AlignmentType, PageBreak, TabStopType } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 10206, S = 23;
const line = { style: BorderStyle.SINGLE, size: 4, color: C.line }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const terra = { style: BorderStyle.SINGLE, size: 10, color: C.terra };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const P = (t, o = {}) => new Paragraph({ keepNext: o.keepNext, alignment: o.align, tabStops: o.tabs, spacing: { after: o.after ?? 60, before: o.before ?? 0, line: o.line ?? 270 },
  indent: o.indent, border: o.border, children: Array.isArray(t) ? t : runs(t, o) });
const H = (n, t, sub) => P([R("Schritt " + n + "  ", { font: "Cambria", bold: true, size: 26, color: C.terra }), R(t, { font: "Cambria", bold: true, size: 26 }),
  ...(sub ? [R("   " + sub, { size: 19, italics: true, color: C.muted })] : [])], { before: 200, after: 80, keepNext: true, border: { bottom: { ...terra, space: 2 } } });
const auftrag = (arr) => arr.map((t, i) => P([R((i + 1) + ".", { bold: true, color: C.terra }), R("\t"), ...runs(t)], { tabs: [{ type: TabStopType.LEFT, position: 340 }], indent: { left: 340, hanging: 340 }, after: 40 }));
const cell = (children, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.CENTER, columnSpan: o.span,
  margins: { top: o.mt ?? 50, bottom: o.mb ?? 50, left: 100, right: 100 }, borders: o.borders || { top: line, bottom: line, left: line, right: line }, children });
const zv = (a, b) => "Z. " + D.zeileVon(a) + "–" + D.zeileVon(b);

// ---------- Vorderseite ----------
const ARG = [
  ["Zehn neue Phylen: Jede besteht aus je einem Bezirk der Stadt, der Küste und des Binnenlandes. Die Gefolgsleute eines Adligen sind nun auf mehrere Phylen verteilt."],
  ["Bürger ist, wer in der Liste seines Dorfes (Demos) steht – nicht mehr die Familie entscheidet."],
  ["Solons Vermögensklassen bleiben: Archont kann nur werden, wer zu den Reichsten gehört."],
  ["Rat der 500: 50 Männer aus jeder Phyle bereiten die Beschlüsse der Volksversammlung vor."],
  ["Niemand verliert Land: Die Adligen bleiben reich und angesehen."],
  ["Ostrakismos: Die Volksversammlung kann einen zu mächtigen Mann für zehn Jahre verbannen."],
  ["Der Areopag aus ehemaligen Archonten beaufsichtigt weiterhin die Beamten."],
  ["Die Theten dürfen nur abstimmen und klagen, aber kein Amt übernehmen."],
  ["Adlige führen weiterhin die Politik – Kleisthenes selbst ist Adliger, später auch Perikles."],
  ["Quelle (um 430 v. Chr.): Der adlige Verfasser beklagt, dass es in Athen die „gemeinen Leute“ besser haben als die „Edlen“ (" + zv("Was die Staatsform", "die gemeinen Leute besser") + ")."],
  ["Quelle: Die militärische Führung verlangt das Volk nicht für sich, es überlässt sie anderen (" + zv("Jedoch: Alle Ämter", "militärischen Führung") + ")."],
  ["Erst Ephialtes und Perikles (462 v. Chr.) entmachten den Areopag; Perikles führt Diäten ein."],
];
const cw = [520, 6886, 800, 1000, 1000];
const argTab = new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [
  new TableRow({ tableHeader: true, children: [
    cell([P("Nr.", { bold: true, size: 19, after: 0, align: AlignmentType.CENTER })], cw[0], { borders: { top: line, left: line, right: line, bottom: terra } }),
    cell([P("Was wir über die Reformen und die Quelle wissen", { bold: true, size: 19, after: 0 })], cw[1], { borders: { top: line, left: line, right: line, bottom: terra } }),
    cell([P("A, B oder C", { bold: true, size: 17, after: 0, line: 220, align: AlignmentType.CENTER })], cw[2], { borders: { top: line, left: line, right: line, bottom: terra } }),
    cell([P("spricht für ein Ende", { bold: true, size: 17, after: 0, line: 220, align: AlignmentType.CENTER, color: C.green })], cw[3], { borders: { top: line, left: line, right: line, bottom: terra } }),
    cell([P("spricht dagegen", { bold: true, size: 17, after: 0, line: 220, align: AlignmentType.CENTER, color: C.terra })], cw[4], { borders: { top: line, left: line, right: line, bottom: terra } })] }),
  ...ARG.map(([t], i) => new TableRow({ cantSplit: true, children: [
    cell([P(String(i + 1), { bold: true, size: 20, after: 0, align: AlignmentType.CENTER })], cw[0]),
    cell([P(t, { size: 22, after: 0, line: 260 })], cw[1], { mt: 60, mb: 60 }),
    cell([P("", { after: 0 })], cw[2]),
    cell([P("☐", { size: 32, after: 0, align: AlignmentType.CENTER, color: C.green })], cw[3]),
    cell([P("☐", { size: 32, after: 0, align: AlignmentType.CENTER, color: C.terra })], cw[4])] })),
] });

const vorne = [
  P([R("GESCHICHTE · KLASSE 11 · PROBEKLAUSUR AUFGABE 4", { size: 16, bold: true, color: C.terra }), R("\tName: ______________________", { size: 18, color: C.muted })],
    { tabs: [{ type: TabStopType.RIGHT, position: W }], after: 40 }),
  P([R("Das Urteil schreiben", { font: "Cambria", bold: true, size: 38 })], { after: 80 }),
  P([R("Aufgabe 4: ", { bold: true, size: 23 }), R("Beurteilen Sie, ob die Reformen des Kleisthenes der Vormachtstellung des Adels in der attischen Polis ein Ende gesetzt haben.", { size: 23, italics: true })],
    { indent: { left: 160 }, after: 100, line: 290, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.terra, space: 8 } } }),
  P("**So gehst du vor:** erst die Frage klären und dich vorläufig festlegen (1) – dann Argumente sortieren (2) – dann abwägen und das **Sachurteil** schreiben (3) – zum Schluss das **Werturteil** (4).", { size: 20, after: 40 }),

  H(1, "Frage klären und vorläufig festlegen"),
  P("Worin bestand die Vormachtstellung des Adels **vor** Kleisthenes? Prüfe bei jedem Argument, welchen dieser drei Punkte es betrifft:", { size: 21, after: 50 }),
  new Table({ columnWidths: [3402, 3402, 3402], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
    ["A  Gefolgschaft", "Jeder Adlige hat seine Gegend als geschlossenen Block hinter sich (wie Peisistratos)."],
    ["B  Ämter", "Nur die Reichen können die höchsten Ämter übernehmen."],
    ["C  Reichtum und Ansehen", "Der Adel besitzt das meiste Land und führt die Politik."],
  ].map(([k, t]) => cell([P(k, { bold: true, size: 21, color: C.terra, after: 20 }), P(t, { size: 21, after: 0, line: 250 })], 3402, { v: VerticalAlign.TOP, mt: 70, mb: 70 })) })] }),
  P([R("Meine erste Einschätzung: ", { bold: true }), R("Die Reformen des Kleisthenes haben die Vormachtstellung des Adels beendet.", {})], { before: 120, after: 30 }),
  P([R("☐ ja            ☐ teilweise            ☐ nein", { size: 26, bold: true, color: C.terra })], { align: AlignmentType.CENTER, after: 40 }),

  H(2, "Argumente sortieren", "allein"),
  ...auftrag([
    "**Kreuze an:** Spricht die Aussage **für** oder **gegen** ein Ende der Vormachtstellung des Adels?",
    "**Notiere** in der Spalte A, B oder C: Welchen Punkt aus Schritt 1 betrifft die Aussage?",
    "**Kreise** auf jeder Seite das Argument **ein**, das du für das wichtigste hältst.",
  ]),
  P("", { after: 40 }),
  argTab,
];

// ---------- Rückseite ----------
// Satzanfänge zum Abhaken: zwei Spalten je Baustein
function baustein(nr, titel, hinweis, saetze) {
  const half = Math.ceil(saetze.length / 2), cols = [saetze.slice(0, half), saetze.slice(half)];
  const sp = arr => arr.map(t => P([R("☐  ", { size: 24, color: C.terra }), ...runs(t, { size: 23 })], { indent: { left: 380, hanging: 380 }, after: 70, line: 270 }));
  return [
    P([R(nr + "  ", { font: "Cambria", bold: true, size: 24, color: C.terra }), R(titel, { font: "Cambria", bold: true, size: 24 }), R("   " + hinweis, { size: 19, italics: true, color: C.muted })],
      { before: 140, after: 50, keepNext: true }),
    new Table({ columnWidths: [W / 2, W / 2], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ cantSplit: true, children: cols.map(c =>
      cell(sp(c), W / 2, { v: VerticalAlign.TOP, mt: 30, mb: 10, borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 12, color: C.terra } } })) })] }),
  ];
}
const hinten = [
  new Paragraph({ children: [new PageBreak()] }),
  H(3, "Abwägen und das Sachurteil schreiben", "ins Heft · aus der Sicht der damaligen Zeit"),
  ...auftrag([
    "**Schreibe** dein Sachurteil ins Heft – in der Reihenfolge der Bausteine a bis d.",
    "**Nutze** aus jedem Baustein mindestens einen Satzanfang und **hake** jeden Satzanfang **ab**, den du benutzt hast.",
    "**Belege** jedes Argument mit einem Beispiel aus Schritt 2 (bei der Quelle mit Zeilenangabe).",
  ]),
  ...baustein("a", "Einstieg: Frage aufgreifen und vorläufig Stellung nehmen", "1–2 Sätze", [
    "„Im Folgenden soll beurteilt werden, ob …“",
    "„Ob …, lässt sich nur mit … beantworten.“",
    "„Um diese Frage zu beantworten, muss man klären, worauf die Macht des Adels beruhte: …“",
    "„Vor Kleisthenes beruhte die Macht des Adels darauf, dass …“",
  ]),
  ...baustein("b", "Argumente dafür", "mindestens 2 mit Beleg", [
    "„Für ein Ende der Vormachtstellung spricht, dass …“",
    "„Dies zeigt sich daran, dass …“",
    "„Ein weiteres Argument ist …“",
    "„Dadurch konnte ein Adliger nicht mehr …“",
    "„Auch die Quelle bestätigt dies: … (Z. …)“",
    "„Hinzu kommt, dass …“",
  ]),
  ...baustein("c", "Argumente dagegen", "mindestens 2 mit Beleg", [
    "„Dagegen spricht jedoch, dass …“",
    "„Allerdings muss man bedenken, dass …“",
    "„Zwar …, aber …“",
    "„Einschränkend ist festzuhalten, dass …“",
    "„Weiterhin blieb …“",
    "„Erst … (Jahr) …“",
  ]),
  ...baustein("d", "Gewichten und Fazit", "Frage am Ende beantworten", [
    "„Schwerer wiegt meines Erachtens …, weil …“",
    "„Entscheidend ist, dass …“",
    "„Zugegebenermaßen …, allerdings …“",
    "„Insgesamt haben die Reformen die Vormachtstellung des Adels …, weil …“",
    "„Daher komme ich zu dem Ergebnis, dass …“",
    "„Kleisthenes hat … zwar nicht …, aber …“",
  ]),

  H(4, "Das Werturteil schreiben", "ins Heft · aus heutiger Sicht"),
  ...auftrag([
    "**Wähle** einen Maßstab und kreuze ihn an.",
    "**Bewerte** die Reformen aus heutiger Sicht nach diesem Maßstab. **Nenne** den Maßstab ausdrücklich und begründe mit einem Beispiel.",
  ]),
  P([R("☐ Gleichheit        ☐ Leistung        ☐ Schutz der Schwachen        ☐ gleiche Regeln für alle", { size: 24, bold: true, color: C.terra })], { align: AlignmentType.CENTER, before: 40, after: 20 }),
  ...baustein("e", "Werturteil", "Maßstab + Position + Begründung", [
    "„Gemessen am Maßstab der … halte ich die Reformen für …, weil …“",
    "„Aus heutiger Sicht sind die Reformen zu begrüßen / kritisch zu sehen, da …“",
    "„Diesem Urteil liegt der Maßstab … zugrunde.“",
    "„Positiv ist …, kritisch dagegen …“",
    "„Die Kritik des Verfassers überzeugt mich (nicht), weil …“",
    "„Ein Beispiel dafür ist …“",
  ]),
  P("**Achtung:** „Ich finde das gut“ ist noch kein Werturteil – erst Maßstab und Begründung machen es dazu.", { size: 20, before: 120, color: C.muted }),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 850, right: 850 } } }, children: [...vorne, ...hinten] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
