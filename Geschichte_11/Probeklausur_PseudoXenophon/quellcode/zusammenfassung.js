// Inhaltliche Zusammenfassung der Stunde vom 11.11. (Probeklausur bis zur Position des Verfassers) mit Ausblick auf Urteil und Klausur. Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, AlignmentType } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9922, S = 22;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const P = (t, o = {}) => new Paragraph({ keepNext: o.keepNext, alignment: o.align, spacing: { after: o.after ?? 90, before: o.before ?? 0, line: o.line ?? 290 }, indent: o.indent, border: o.border,
  children: Array.isArray(t) ? t : runs(t, o) });
const H = (n, t) => P([R(n + "  ", { font: "Cambria", bold: true, size: 30, color: C.terra }), R(t, { font: "Cambria", bold: true, size: 28 })], { before: 260, after: 100, keepNext: true });
const li = t => P("–  " + t, { indent: { left: 300, hanging: 300 }, after: 50 });
const kasten = t => P(t, { italics: true, line: 310, align: AlignmentType.JUSTIFIED, indent: { left: 160 }, after: 120, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.terra, space: 8 } } });
function tab(cw, head, rows, size = 20) {
  const cell = (t, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 110, right: 110 },
    borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children: [P(t, { size, bold: o.bold, after: 0, line: 260 })] });
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [
    new TableRow({ tableHeader: true, children: head.map((t, i) => cell(t, i, { head: true, bold: true })) }),
    ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, i === 0 ? { bold: true } : {})) }))] });
}

const children = [
  P("GESCHICHTE · KLASSE 11 · ZUSAMMENFASSUNG DER STUNDE VOM 11.11.", { size: 16, bold: true, color: C.terra, after: 30 }),
  P([R("Probeklausur: Pseudo-Xenophon über die Demokratie", { font: "Cambria", bold: true, size: 36 })], { after: 40 }),
  P([R("Was wir erarbeitet haben – bis zur Position des Verfassers – und die Reformen des Kleisthenes für das Urteil", { size: 22, italics: true, color: C.muted })], { after: 160 }),

  H("1", "Worum es ging"),
  P("Wir haben an einer Probeklausur geübt, wie man eine Textquelle interpretiert – nach den Arbeitsschritten der Übersicht „So interpretierst du eine Textquelle“. Die Quelle ist ein Auszug aus „Über den Staat der Athener“ von einem unbekannten Verfasser („Pseudo-Xenophon“), geschrieben um 430–424 v. Chr. Er beschreibt die attische Demokratie aus der Sicht eines Gegners."),
  P("**Themenfrage:** " + D.themenfrage),

  H("2", "Einleitung: die formale Analyse (TATTEA)"),
  tab([2000, 7922], ["TATTEA", "Angabe"], D.tattea.map(([a, b]) => [a, b])),
  P("**So kann die Einleitung lauten:**", { before: 140, after: 60 }),
  kasten("„" + D.einleitung + "“"),
  P("**Achtung:** Der Autor ist unbekannt – die Schrift wurde Xenophon nur irrtümlich zugeschrieben. Er schreibt um 430 v. Chr., also zur Zeit des Perikles, nicht zur Zeit des Kleisthenes.", { size: 20 }),

  H("3", "Die Position des Verfassers"),
  P("Wir haben den Text in Sinnabschnitte gegliedert und zu jedem Abschnitt die Haltung des Verfassers herausgearbeitet:", { after: 80 }),
  tab([1100, 2700, 6122], ["Zeilen", "Aspekt", "Position des Verfassers"], D.position.map(([a, b, z]) => [z, a, b])),
  P("**Gesamtposition:** " + D.positionGesamt, { before: 140 }),
  P("**Textsprache:** " + D.sprache.replace(/^Sprache /, "")),
  P("**So kann die Position ausformuliert werden:**", { before: 100, after: 60, keepNext: true }),
  kasten(D.muster2),

  H("4", "Die Reformen des Kleisthenes (508/07 v. Chr.)"),
  P("Dieses Wissen brauchst du, um die Themenfrage im Urteil zu beantworten.", { italics: true, color: C.muted, after: 80 }),
  P("**Ausgangslage:** Nach dem Sturz der Tyrannis (510 v. Chr.) kämpften die Adelsfamilien wieder um die Macht. Ihre Macht beruhte auf ihrer Gegend: Abhängige Bauern folgten ihrem Adligen als geschlossener Block – so war auch Peisistratos an die Macht gekommen. Kleisthenes setzte sich mit Hilfe des Volkes gegen seinen Rivalen Isagoras durch."),
  tab([2700, 7222], ["Maßnahme", "Was sich änderte"], [
    ["Demen (Dörfer)", "Grundeinheit; Bürger ist, wer in der Liste seines Dorfes steht – nicht mehr die Familie entscheidet. Man nennt sich nach dem Dorf („aus Paiania“)."],
    ["Trittyen (Bezirke)", "Attika in drei Gebiete geteilt (Stadt, Küste, Binnenland), jedes in zehn Bezirke: 30 Trittyen."],
    ["10 neue Phylen", "Jede Phyle bekommt per Los je einen Bezirk aus Stadt, Küste und Binnenland – die Bürger werden „vermischt“. Die vier alten Phylen nach Abstammung verlieren ihre Bedeutung."],
    ["Rat der 500", "50 Männer aus jeder Phyle (statt Rat der 400); bereitet die Beschlüsse der Volksversammlung vor."],
    ["Ostrakismos", "Die Volksversammlung kann einen zu mächtigen Mann für zehn Jahre verbannen."],
  ]),
  P("**Wirkung:** Niemand musste umziehen und niemand verlor Land. Aber die Gefolgsleute eines Adligen verteilten sich nun auf mehrere Phylen, und in jeder Phyle waren sie nur ein Drittel. Kein Adliger hatte mehr seine Gegend als Machtblock hinter sich – im Rat, im Heer und beim Bürgerrecht entschieden Bürger aus ganz Attika gemeinsam.", { before: 120 }),
  P("**Was gleich blieb:** Solons Vermögensklassen bestanden weiter. Archont konnte nur werden, wer zu den Reichsten gehörte; der Areopag aus ehemaligen Archonten beaufsichtigte die Beamten. Die Theten durften nur abstimmen und klagen. Frauen, Metöken und Sklaven hatten keine politischen Rechte. Adlige führten weiterhin die Politik."),
  P("**Danach:** Themistokles ließ die Archonten ab 487/86 v. Chr. auslosen, die Strategen blieben gewählt; durch die Flotte wurden die Theten wichtig. 462 v. Chr. entmachteten Ephialtes und Perikles den Areopag, Perikles führte Diäten ein. Erst jetzt sprach man von „Demokratie“ – und erst darüber klagt der Verfasser um 430 v. Chr."),

  H("5", "Ausblick"),
  li("**Nächste Stunde:** Urteilsbildung – Haben die Reformen die Vormachtstellung des Adels beendet? Wende dafür dein Wissen über die Reformen (Abschnitt 4) an: dafür und dagegen abwägen, Fazit, Werturteil mit Maßstab."),
  li("**Klausur am 25.11. (90 Minuten):** Einleitung mit Themenfrage und TATTEA, Position des Verfassers herausarbeiten, Urteil. Eine eigene Beschreibungsaufgabe zu den Reformen gibt es nicht – im Urteil musst du sie aber als Belege anwenden."),
];
const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 992, right: 992 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
