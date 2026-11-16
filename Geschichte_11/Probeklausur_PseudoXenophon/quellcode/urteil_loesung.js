// Lösung zum Hilfeblatt „Das Urteil schreiben“: Sortierung (Schritt 1–2) und Musterurteil aus den Satzanfängen (Schritt 3–4). Tintensparend.
// In den Mustertexten sind die benutzten Satzanfänge fett, die Bausteine a–e stehen am Rand – so sehen die Schüler, wie die Satzanfänge eingesetzt werden.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, AlignmentType, PageBreak, TabStopType } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 10206, S = 22;
const line = { style: BorderStyle.SINGLE, size: 4, color: C.line }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const terra = { style: BorderStyle.SINGLE, size: 10, color: C.terra };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold, color: x.startsWith("**") ? (o.hl || o.color) : o.color }));
const P = (t, o = {}) => new Paragraph({ keepNext: o.keepNext, alignment: o.align, tabStops: o.tabs, spacing: { after: o.after ?? 60, before: o.before ?? 0, line: o.line ?? 270 },
  indent: o.indent, border: o.border, children: Array.isArray(t) ? t : runs(t, o) });
const H = (n, t, sub) => P([R("Schritt " + n + "  ", { font: "Cambria", bold: true, size: 26, color: C.terra }), R(t, { font: "Cambria", bold: true, size: 26 }),
  ...(sub ? [R("   " + sub, { size: 19, italics: true, color: C.muted })] : [])], { before: 200, after: 80, keepNext: true, border: { bottom: { ...terra, space: 2 } } });
const cell = (children, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.CENTER,
  margins: { top: o.mt ?? 45, bottom: o.mb ?? 45, left: 100, right: 100 }, borders: o.borders || { top: line, bottom: line, left: line, right: line }, children });
const zv = (a, b) => "Z. " + D.zeileVon(a) + "–" + D.zeileVon(b);
const Z14 = zv("Was die Staatsform", "die gemeinen Leute besser"), Z1113 = zv("Jedoch: Alle Ämter", "militärischen Führung");

// Nr., Kurzfassung, A/B/C, dafür (true) / dagegen (false), Erklärung
const ARG = [
  [1, "Zehn neue Phylen, Gefolgsleute verteilt", "A", true, "Kern der Reform: Kein Adliger hat seine Gegend mehr als Block hinter sich."],
  [2, "Bürgerrecht über die Dorfliste", "A", true, "Nicht mehr die Familie (der Adelsherr) entscheidet, wer dazugehört."],
  [3, "Vermögensklassen bleiben, Archonten nur Reiche", "B", false, "Die höchsten Ämter bleiben dem Adel und den Reichen."],
  [4, "Rat der 500 aus allen Phylen", "A", true, "Bürger aus ganz Attika beraten gemeinsam – kein Block eines Adligen."],
  [5, "Niemand verliert Land", "C", false, "Reichtum und Ansehen des Adels bleiben unangetastet."],
  [6, "Ostrakismos", "A / C", true, "Das Volk kann einen zu mächtigen Adligen verbannen."],
  [7, "Areopag beaufsichtigt weiter die Beamten", "B", false, "Ein Gremium aus ehemaligen Archonten, also aus Reichen, kontrolliert weiter."],
  [8, "Theten ohne Ämter", "B", false, "Die Ärmsten dürfen nur abstimmen und klagen."],
  [9, "Adlige führen weiter die Politik", "C", false, "Kleisthenes selbst ist Adliger, später auch Perikles."],
  [10, "Quelle: „gemeine Leute“ besser als die „Edlen“ (" + Z14 + ")", "C", true, "Um 430 klagt ein Adliger, dass das Volk herrscht. Achtung: erst nach Ephialtes und Perikles."],
  [11, "Quelle: Volk überlässt die militärische Führung anderen (" + Z1113 + ")", "C", false, "Die Führung bleibt bei den Vornehmen und Reichen."],
  [12, "Erst Ephialtes und Perikles (462) entmachten den Areopag", "B", false, "Das Ende der Vormacht kommt erst später, nicht durch Kleisthenes."],
];
const cw = [520, 3500, 700, 900, 900, 3686];
const head = ["Nr.", "Aussage (kurz)", "A, B, C", "für ein Ende", "dagegen", "Warum?"];
const tab = new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [
  new TableRow({ tableHeader: true, children: head.map((t, i) => cell([P(t, { bold: true, size: 18, after: 0, line: 220, align: i === 1 || i === 5 ? undefined : AlignmentType.CENTER,
    color: i === 3 ? C.green : i === 4 ? C.terra : C.dark })], cw[i], { borders: { top: line, left: line, right: line, bottom: terra } })) }),
  ...ARG.map(([n, t, k, pro, w]) => new TableRow({ cantSplit: true, children: [
    cell([P(String(n), { bold: true, size: 20, after: 0, align: AlignmentType.CENTER })], cw[0]),
    cell([P(t, { size: 20, after: 0, line: 240 })], cw[1]),
    cell([P(k, { bold: true, size: 21, after: 0, align: AlignmentType.CENTER })], cw[2]),
    cell([P(pro ? "✔" : "", { bold: true, size: 26, after: 0, align: AlignmentType.CENTER, color: C.green })], cw[3]),
    cell([P(pro ? "" : "✔", { bold: true, size: 26, after: 0, align: AlignmentType.CENTER, color: C.terra })], cw[4]),
    cell([P(w, { size: 19, after: 0, line: 230, color: C.muted })], cw[5])] })),
] });

// Mustertext mit Baustein-Buchstabe am Rand
function muster(rows) {
  const cw2 = [560, W - 560];
  return new Table({ columnWidths: cw2, width: { size: W, type: WidthType.DXA }, rows: rows.map(([b, t]) => new TableRow({ cantSplit: true, children: [
    cell([P(b, { font: "Cambria", bold: true, size: 26, color: C.terra, after: 0, align: AlignmentType.CENTER })], cw2[0], { v: VerticalAlign.TOP, mt: 60, borders: { top: none, bottom: none, left: none, right: { style: BorderStyle.SINGLE, size: 12, color: C.terra } } }),
    cell([P(t, { size: 22, line: 285, after: 0, align: AlignmentType.JUSTIFIED, hl: C.terra })], cw2[1], { v: VerticalAlign.TOP, mt: 60, mb: 90, borders: { top: none, bottom: none, left: none, right: none } })] })) });
}

const children = [
  P([R("GESCHICHTE · KLASSE 11 · PROBEKLAUSUR AUFGABE 4", { size: 16, bold: true, color: C.terra }), R("\tLÖSUNG", { size: 18, bold: true, color: C.terra })],
    { tabs: [{ type: TabStopType.RIGHT, position: W }], after: 40 }),
  P([R("Das Urteil schreiben – Lösung", { font: "Cambria", bold: true, size: 38 })], { after: 80 }),
  P([R("Aufgabe 4: ", { bold: true }), R("Beurteilen Sie, ob die Reformen des Kleisthenes der Vormachtstellung des Adels in der attischen Polis ein Ende gesetzt haben.", { italics: true })],
    { indent: { left: 160 }, after: 80, line: 290, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.terra, space: 8 } } }),

  H("1", "Erste Einschätzung"),
  P("Am besten begründbar ist **„teilweise“**. „Ja“ oder „nein“ sind auch möglich – entscheidend ist, dass das Urteil beide Seiten abwägt und begründet.", { after: 40 }),

  H("2", "Argumente sortieren"),
  tab,
  P([R("Wichtigste Argumente: ", { bold: true }), R("dafür ", { bold: true, color: C.green }), R("Nr. 1 – die neuen Phylen zerschlagen die Gefolgschaft (A).   "),
    R("dagegen ", { bold: true, color: C.terra }), R("Nr. 3 – die höchsten Ämter bleiben den Reichen (B).")], { before: 100, after: 40 }),
  P([R("Das Muster: ", { bold: true }), R("Die Argumente dafür betreffen fast alle "), R("A", { bold: true, color: C.terra }), R(", die Argumente dagegen "), R("B und C", { bold: true, color: C.terra }),
    R(". Kleisthenes brach also die Gefolgschaft, nicht aber Ämter und Reichtum des Adels – genau das ist der Kern des Urteils „teilweise“.")], { after: 0 }),

  new Paragraph({ children: [new PageBreak()] }),
  H("3", "Musterlösung Sachurteil", "fett = benutzte Satzanfänge · am Rand = Baustein"),
  muster([
    ["a", "**Im Folgenden soll beurteilt werden, ob** die Reformen des Kleisthenes der Vormachtstellung des Adels ein Ende gesetzt haben. **Vor Kleisthenes beruhte die Macht des Adels darauf, dass** jeder Adlige seine Gegend als geschlossenen Block hinter sich hatte, nur die Reichen die höchsten Ämter übernehmen konnten und der Adel das meiste Land besaß."],
    ["b", "**Für ein Ende der Vormachtstellung spricht, dass** Kleisthenes zehn neue Phylen schuf, die jeweils aus einem Bezirk der Stadt, der Küste und des Binnenlandes bestanden. **Dadurch konnte ein Adliger nicht mehr** seine Gegend als Machtblock einsetzen, denn seine Gefolgsleute waren auf mehrere Phylen verteilt. **Dies zeigt sich daran, dass** im Rat der 500 nun Bürger aus allen Teilen Attikas gemeinsam die Beschlüsse vorbereiteten. **Hinzu kommt, dass** über das Bürgerrecht die Dorfliste entschied und nicht mehr die Familie und dass die Volksversammlung mit dem Ostrakismos einen zu mächtigen Mann verbannen konnte."],
    ["c", "**Dagegen spricht jedoch, dass** Solons Vermögensklassen bestehen blieben: Archont konnte weiterhin nur werden, wer zu den Reichsten gehörte, und der Areopag aus ehemaligen Archonten beaufsichtigte weiter die Beamten. Die Theten durften nur abstimmen und klagen. **Zwar** verloren die Adligen ihre Gefolgschaft als Block, **aber** sie behielten ihr Land und ihr Ansehen und führten weiterhin die Politik – Kleisthenes selbst war Adliger, später auch Perikles. Auch die Quelle zeigt, dass das Volk die militärische Führung anderen überlässt (" + Z1113 + "). **Erst** 462 v. Chr. entmachteten Ephialtes und Perikles den Areopag."],
    ["d", "**Schwerer wiegt meines Erachtens** die Neuordnung der Phylen, **weil** sie die wichtigste Grundlage der Adelsmacht, die geschlossene Gefolgschaft, zerschlug. **Zugegebenermaßen** blieben Ämter und Reichtum beim Adel, **allerdings** entschieden nun Bürger aus ganz Attika gemeinsam. Dass um 430 v. Chr. ein adliger Verfasser klagt, die „gemeinen Leute“ hätten es besser als die „Edlen“ (" + Z14 + "), zeigt, dass die Vormacht des Adels später tatsächlich gebrochen war. **Insgesamt haben die Reformen die Vormachtstellung des Adels** also nur teilweise beendet, **weil** Ämter und Reichtum beim Adel blieben. **Kleisthenes hat** sie **zwar nicht** beendet, **aber** ihre wichtigste Grundlage zerstört und damit den Weg zur Demokratie geöffnet."],
  ]),

  H("4", "Musterlösung Werturteil", "Maßstab: Gleichheit"),
  muster([
    ["e", "**Gemessen am Maßstab der** Gleichheit **halte ich die Reformen für** einen wichtigen, aber unvollständigen Schritt, **weil** sie vielen Bürgern echten Einfluss gaben, die Rechte aber weiter vom Besitz abhingen. **Positiv ist**, dass jeder Bürger in der Volksversammlung die gleiche Stimme hatte, **kritisch dagegen**, dass Frauen, Metöken und Sklaven ganz ausgeschlossen blieben. **Ein Beispiel dafür ist**, dass ein Thete zwar abstimmen, aber kein Amt übernehmen durfte. **Die Kritik des Verfassers überzeugt mich nicht, weil** er die Rechte des Volkes nur am Nutzen für den Adel misst."],
  ]),
  P("**Andere Maßstäbe sind ebenso richtig, wenn sie begründet werden:**", { before: 140, after: 50, keepNext: true }),
  ...[
    "**Leistung:** Die Theten rudern die Flotte und sichern Athens Macht, bekommen aber kein Amt → aus heutiger Sicht ungerecht.",
    "**Schutz der Schwachen:** Der Ostrakismos schützt vor einem neuen Tyrannen; die Armen bleiben aber von den Ämtern ausgeschlossen.",
    "**Gleiche Regeln für alle:** Abstimmung und Klage gelten für alle Bürger gleich, die Ämter nicht.",
  ].map(t => P("–  " + t, { size: 21, indent: { left: 300, hanging: 300 }, after: 40 })),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 850, right: 850 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
