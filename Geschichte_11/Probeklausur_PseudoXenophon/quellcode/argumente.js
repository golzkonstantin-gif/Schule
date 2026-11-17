// Blatt „Argumente und Beispielurteil“ (Rückmeldung 17.11.: Vorwissen zu Kleisthenes fehlte, Tafelbild nicht verstanden).
// Vorne: Argumente zur Ausgangsfrage schon sortiert (dafür/dagegen), je Argument „Was?“ und „Was bedeutet das für den Adel?“ in einfacher Sprache.
// Hinten: Beispiel-Sachurteil und -Werturteil zu einem bekannten Thema (Solon), damit die Antwort zu Kleisthenes nicht vorweggenommen wird.
// A4 doppelseitig, tintensparend (keine Flächen, Farbe nur in Schrift und Linien).
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, AlignmentType, PageBreak, TabStopType } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 10206, S = 23;
const line = { style: BorderStyle.SINGLE, size: 4, color: C.line }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold, color: x.startsWith("**") && o.hl ? o.hl : o.color }));
const P = (t, o = {}) => new Paragraph({ keepNext: o.keepNext, alignment: o.align, tabStops: o.tabs, spacing: { after: o.after ?? 60, before: o.before ?? 0, line: o.line ?? 270 },
  indent: o.indent, border: o.border, children: Array.isArray(t) ? t : runs(t, o) });
const cell = (children, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.TOP,
  margins: { top: o.mt ?? 60, bottom: o.mb ?? 60, left: 110, right: 110 }, borders: o.borders || { top: line, bottom: line, left: line, right: line }, children });
const auftrag = arr => arr.map((t, i) => P([R((i + 1) + ".", { bold: true, color: C.terra }), R("\t"), ...runs(t)], { tabs: [{ type: TabStopType.LEFT, position: 360 }], indent: { left: 360, hanging: 360 }, after: 40 }));
const zv = (a, b) => "Z. " + D.zeileVon(a) + "–" + D.zeileVon(b);

// ---------- Vorderseite: sortierte Argumente ----------
const DAFUER = [
  ["Neue Einteilung der Bürger", "Die Bürger werden in 10 neue Gruppen (Phylen) eingeteilt. Jede Gruppe wird aus Leuten aus Stadt, Küste und Binnenland zusammengelost.",
    "Die Bauern eines Adligen landen in verschiedenen Gruppen. Er kann sie nicht mehr geschlossen für sich abstimmen lassen."],
  ["Rat der 500", "Jede der 10 Gruppen schickt 50 Männer in den Rat. Der Rat bereitet die Beschlüsse der Volksversammlung vor.",
    "Im Rat sitzen Bürger aus ganz Attika gemischt. Kein Adliger hat dort mehr einen eigenen Block."],
  ["Bürgerrecht über das Dorf", "Wer Bürger ist, steht in der Liste seines Dorfes (Demos).",
    "Nicht mehr die adlige Familie entscheidet, wer dazugehört."],
  ["Scherbengericht (Ostrakismos)", "Die Volksversammlung kann einen zu mächtigen Mann für 10 Jahre verbannen.",
    "Ein Adliger, der zu viel Macht will (wie früher Peisistratos), kann entfernt werden."],
];
const DAGEGEN = [
  ["Ämter nur für Reiche", "Solons Vermögensklassen bleiben. Archont kann nur werden, wer zu den Reichsten gehört.",
    "Die wichtigsten Ämter bleiben beim Adel."],
  ["Areopag bleibt", "Der Rat aus ehemaligen Archonten beaufsichtigt weiter die Beamten.",
    "Ein Gremium der Reichen und Adligen behält die Kontrolle."],
  ["Kein Land weggenommen", "Niemand verliert Besitz. Kleisthenes selbst ist Adliger, später führt auch Perikles.",
    "Der Adel bleibt reich, angesehen und führt weiter die Politik."],
  ["Die Ärmsten ohne Ämter", "Die Theten dürfen nur abstimmen und klagen, aber kein Amt übernehmen.",
    "Mitbestimmen ja – regieren nein."],
];
const half = (W - 200) / 2;
function karte([titel, was, adel], farbe) {
  return [
    P([R("☐  ", { size: 28, color: farbe }), R(titel, { bold: true, size: 26, color: farbe })], { after: 20, keepNext: true }),
    P([R("Was? ", { bold: true, size: 24, color: C.muted }), R(was, { size: 24 })], { after: 30, line: 270, keepNext: true }),
    P([R("→ Für den Adel: ", { bold: true, size: 24, color: farbe }), R(adel, { size: 24, bold: true })], { after: 220, line: 270 }),
  ];
}
const spalte = (titel, arr, farbe) => cell([P(titel, { bold: true, size: 26, color: farbe, after: 100, font: "Cambria", border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: farbe, space: 3 } } }),
  ...arr.flatMap(a => karte(a, farbe))], half, { mt: 80, borders: { top: line, bottom: line, left: line, right: line } });

const vorne = [
  P([R("GESCHICHTE · KLASSE 11 · PROBEKLAUSUR AUFGABE 4", { size: 16, bold: true, color: C.terra }), R("\tSach- und Werturteil", { size: 18, color: C.muted })],
    { tabs: [{ type: TabStopType.RIGHT, position: W }], after: 40 }),
  P([R("Argumente zur Frage – schon sortiert", { font: "Cambria", bold: true, size: 36 })], { after: 80 }),
  P([R("Frage: ", { bold: true, size: 24 }), R("Haben die Reformen des Kleisthenes (508/07 v. Chr.) der Vormachtstellung des Adels in der attischen Polis ein Ende gesetzt?", { size: 24, italics: true })],
    { indent: { left: 160 }, after: 80, line: 290, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.terra, space: 8 } } }),
  P("**Vormachtstellung des Adels** heißt: Der Adel hat mehr Macht als alle anderen – durch seine abhängigen Bauern, durch die Ämter und durch seinen Reichtum.", { size: 21, after: 100, color: C.muted }),
  new Table({ columnWidths: [half, 200, half], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
    spalte("Dafür spricht …", DAFUER, C.green),
    cell([P("")], 200, { borders: { top: none, bottom: none, left: none, right: none } }),
    spalte("Dagegen spricht …", DAGEGEN, C.terra),
  ] })] }),
  P([R("Aus der Quelle: ", { bold: true }), R("dafür ", { bold: true, color: C.green }), R("– um 430 v. Chr. klagt ein Adliger, die „gemeinen Leute“ hätten es besser als die „Edlen“ (" + zv("Was die Staatsform", "die gemeinen Leute besser") + "; aber: erst nach weiteren Reformen 462 v. Chr.). "),
    R("dagegen ", { bold: true, color: C.terra }), R("– das Volk überlässt die militärische Führung anderen (" + zv("Jedoch: Alle Ämter", "militärischen Führung") + ").")], { size: 22, before: 140, after: 160, line: 280 }),
  P([R("Arbeitsauftrag", { font: "Cambria", bold: true, size: 26, color: C.terra })], { after: 60, keepNext: true }),
  ...auftrag([
    "**Lies** die Argumente. **Kreuze** in jeder Spalte die zwei Argumente an, die du am wichtigsten findest.",
    "**Lies** das Beispiel auf der Rückseite: So ist ein Sachurteil gebaut (Bausteine a–d).",
    "**Schreibe** dein Sachurteil ins Heft – genauso gebaut. Nutze die Satzanfänge vom Hilfeblatt und hake sie ab.",
    "**Schreibe** dein Werturteil: Wähle einen Maßstab (Gleichheit, Leistung, Schutz der Schwachen, gleiche Regeln für alle) und begründe.",
  ]),
];

// ---------- Rückseite: Beispielurteil zu Solon ----------
function muster(rows) {
  const cw = [1000, W - 1000];
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: rows.map(([b, name, t]) => new TableRow({ cantSplit: true, children: [
    cell([P(b, { font: "Cambria", bold: true, size: 30, color: C.terra, after: 0, align: AlignmentType.CENTER }), P(name, { size: 17, color: C.muted, after: 0, align: AlignmentType.CENTER, line: 200 })], cw[0],
      { mt: 70, borders: { top: none, bottom: none, left: none, right: { style: BorderStyle.SINGLE, size: 12, color: C.terra } } }),
    cell([P(t, { size: 28, line: 350, after: 0, align: AlignmentType.JUSTIFIED, hl: C.terra })], cw[1], { mt: 70, mb: 170, borders: { top: none, bottom: none, left: none, right: none } })] })) });
}
const hinten = [
  new Paragraph({ children: [new PageBreak()] }),
  P([R("So sieht ein Urteil aus – ein Beispiel zu Solon", { font: "Cambria", bold: true, size: 34 })], { after: 60 }),
  P("Das Thema kennst du schon aus dem Test. **Fett** sind die Satzanfänge, am Rand steht der Baustein. Baue dein Urteil zu Kleisthenes genauso.", { size: 21, color: C.muted, after: 80 }),
  P([R("Frage: ", { bold: true, size: 24 }), R("Haben die Reformen Solons (594/93 v. Chr.) die Krise in Athen beendet?", { size: 24, italics: true })],
    { indent: { left: 160 }, after: 120, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.terra, space: 8 } } }),
  P([R("Beispiel-Sachurteil", { font: "Cambria", bold: true, size: 26, color: C.terra }), R("   aus der Sicht der damaligen Zeit, mit Belegen", { size: 19, italics: true, color: C.muted })], { after: 60, keepNext: true }),
  muster([
    ["a", "Einstieg", "**Im Folgenden soll beurteilt werden, ob** die Reformen Solons die Krise in Athen beendet haben. **Um diese Frage zu beantworten, muss man klären, worin die Krise bestand:** Viele Bauern waren hoch verschuldet, manche wurden sogar versklavt, und Adel und Bauern standen sich feindlich gegenüber."],
    ["b", "dafür", "**Für ein Ende der Krise spricht, dass** Solon die Schulden strich und die Schuldknechtschaft verbot. **Dadurch konnte** kein Athener mehr wegen Schulden versklavt werden. **Hinzu kommt, dass** er versklavte Athener zurückkaufte und alle Bürger an der Volksversammlung teilnehmen durften."],
    ["c", "dagegen", "**Dagegen spricht jedoch, dass** Solon das Land nicht neu verteilte: Die Großgrundbesitzer behielten ihren Besitz, die armen Bauern blieben arm. **Zwar** richteten sich die Rechte nun nach dem Besitz statt nach der Herkunft, **aber** die höchsten Ämter blieben den Reichen vorbehalten. **Weiterhin blieb** die Unruhe, sodass Peisistratos wenige Jahrzehnte später als Tyrann die Macht ergreifen konnte."],
    ["d", "Fazit", "**Schwerer wiegt meines Erachtens** die Abschaffung der Schuldknechtschaft, **weil** sie die schlimmste Not beendete. **Zugegebenermaßen** blieben Armut und Streit, **allerdings** war die größte Gefahr gebannt. **Daher komme ich zu dem Ergebnis, dass** Solons Reformen die Krise nur teilweise beendet haben."],
  ]),
  P([R("Beispiel-Werturteil", { font: "Cambria", bold: true, size: 26, color: C.terra }), R("   aus heutiger Sicht, mit Maßstab", { size: 19, italics: true, color: C.muted })], { before: 160, after: 60, keepNext: true }),
  muster([
    ["e", "Maßstab", "**Gemessen am Maßstab** des Schutzes der Schwachen **halte ich** Solons Reformen **für** einen wichtigen Fortschritt, **weil** niemand mehr wegen Schulden seine Freiheit verlor. **Kritisch dagegen** ist, dass die Armen kein Land bekamen und keine Ämter übernehmen durften. **Ein Beispiel dafür ist**, dass ein Thete zwar abstimmen, aber nie Archont werden konnte."],
  ]),
  P("**Achtung:** Im Sachurteil zählt, was damals galt – im Werturteil sagst du, welchen Maßstab du heute anlegst.", { size: 21, before: 140, color: C.muted }),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 850, right: 850 } } }, children: [...vorne, ...hinten] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
