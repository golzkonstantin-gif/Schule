const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, PageBreak, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const STAND = {
  K: { name: "ERSTER STAND · KLERUS", edge: "8E6BBE", ink: "6B4A9E", rolle: "Du bist Geistlicher. Du zahlst keine Steuern und hast Vorrechte. Du willst sie behalten.", tipp: "Du stimmst nach deiner Rolle." },
  A: { name: "ZWEITER STAND · ADEL", edge: "E8A33D", ink: "B9791A", rolle: "Du bist Adliger. Du zahlst keine Steuern und hast Vorrechte (z. B. Jagdrecht, Ämter am Hof). Du willst sie behalten.", tipp: "Du stimmst nach deiner Rolle." },
  D: { name: "DRITTER STAND (BAUERN UND BÜRGER)", edge: "3F7CC4", ink: "2F6AAE", rolle: "Du zahlst Steuern und Abgaben. Der Brotpreis ist gestiegen, die Familie hat Hunger. Du willst, dass sich etwas ändert.", tipp: "Du stimmst nach deiner Überzeugung." },
};
const W = 10706, CW = W / 2;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, border: o.border });
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const thin = { style: BorderStyle.SINGLE, size: 6, color: "8A96B5" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

// ---- Standeskarten (12 pro A4, ohne Stimmzettel) ----
function card(key, reserve) {
  const s = STAND[key];
  return new TableCell({ children: [
    p(t(reserve ? "STANDESKARTE · RESERVE" : "STANDESKARTE · GEHEIM", { size: 15, bold: true, color: MUTED }), { after: 10 }),
    p(t(s.name, { font: "Cambria", size: 26, bold: true, color: s.ink }), { after: 50 }),
    p(t(s.rolle, { size: 22 }), { after: 40 }),
    p(t(s.tipp, { size: 22, bold: true }), { after: 0 }),
  ], width: { size: CW, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: { style: BorderStyle.SINGLE, size: 48, color: s.edge }, right: dash }, margins: { top: 90, bottom: 60, left: 200, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
function page(list) {
  const rows = [];
  for (let i = 0; i < list.length; i += 2) rows.push(new TableRow({ cantSplit: true, height: { value: 2560, rule: HeightRule.EXACT }, children: [card(...list[i]), card(...list[i + 1])] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [CW, CW], rows });
}
const pg1 = [["K"], ["K"], ["K"], ["A"], ["A"], ["A"], ["A"], ["D"], ["D"], ["D"], ["D"], ["D"]];
const pg2 = [...Array(11).fill(["D"]), ["D", true]];

// ---- Beobachtungszettel (Schnipsel, vier pro A4): nur Ja-Stimmen zählen, Ergebnis selbst ausrechnen ----
const cc = (txt, o = {}) => new TableCell({ children: [p(t(txt, { size: o.size || 20, bold: o.bold, color: o.color }), { after: 0, align: o.align || AlignmentType.CENTER })], width: { size: o.w, type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 50, bottom: 50, left: 70, right: 70 }, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span });
function zettel() {
  const Z = [3400, 3400, 3400];
  const zaehl = new Table({ width: { size: 10200, type: WidthType.DXA }, columnWidths: Z, rows: [
    new TableRow({ children: [cc("Klerus", { w: Z[0], bold: true, color: "6B4A9E" }), cc("Adel", { w: Z[1], bold: true, color: "B9791A" }), cc("Dritter Stand", { w: Z[2], bold: true, color: "2F6AAE" })] }),
    new TableRow({ height: { value: 560, rule: HeightRule.ATLEAST }, children: Z.map((w) => cc("Ja ____ von ____", { w, size: 22, color: "5B6B8C" })) }),
  ] });
  const E = [1900, 4150, 4150];
  const erg = new Table({ width: { size: 10200, type: WidthType.DXA }, columnWidths: E, rows: [
    new TableRow({ height: { value: 520, rule: HeightRule.ATLEAST }, children: [cc("Nach Ständen", { w: E[0], bold: true, size: 20, align: AlignmentType.LEFT }), cc("Stände mit Ja: ____ von 3", { w: E[1], size: 21 }), cc("□ angenommen    □ abgelehnt", { w: E[2], size: 21 })] }),
    new TableRow({ height: { value: 520, rule: HeightRule.ATLEAST }, children: [cc("Nach Köpfen", { w: E[0], bold: true, size: 20, align: AlignmentType.LEFT }), cc("alle Ja zusammen: ____ von ____", { w: E[1], size: 21 }), cc("□ angenommen    □ abgelehnt", { w: E[2], size: 21 })] }),
  ] });
  return [
    p([t("BEOBACHTUNGSZETTEL · GESCHICHTE 8D   ", { size: 15, bold: true, color: MUTED }), t("Abstimmung über Antrag A", { font: "Cambria", size: 26, bold: true, color: NAVY })], { after: 40 }),
    p([t("Mein Stand:  ", { size: 20, bold: true }), t("□ Klerus   □ Adel   □ Dritter Stand", { size: 20 }), t("          Meine Stimme (vorab):  ", { size: 20, bold: true }), t("□ Ja   □ Nein", { size: 20 })], { after: 50 }),
    p(t("1  Abstimmung: Trage die Ja-Stimmen ein, die die Wahlleitung ausruft.", { size: 18, bold: true, color: NAVY }), { after: 30 }),
    zaehl,
    p(t("2  Rechne danach selbst aus:", { size: 18, bold: true, color: NAVY }), { before: 60, after: 30 }),
    erg,
  ];
}
function zettelCell() {
  return new TableCell({ children: zettel(), width: { size: W, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 80, bottom: 60, left: 170, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
const bogenPage = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [0, 1, 2, 3].map(() => new TableRow({ cantSplit: true, height: { value: 3450, rule: HeightRule.EXACT }, children: [zettelCell()] })) });

const g = (txt, o = {}) => new TableCell({ children: [p(t(txt, { size: o.size || 19, bold: o.bold, color: o.color }), { after: 0, align: o.align || AlignmentType.CENTER })], width: { size: o.w, type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 30, bottom: 30, left: 50, right: 50 }, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span });
// ---- Rollenkarten Wahlleitung und Schriftführung ----
function rolle(titel, kicker, schritte, zusatz) {
  const ch = [
    p(t(kicker, { size: 16, bold: true, color: MUTED }), { after: 10 }),
    p(t(titel, { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 70 }),
    ...schritte.map((x, i) => p([t((i + 1) + "  ", { size: 22, bold: true, color: NAVY }), t(x, { size: 21 })], { after: 55 })),
    ...zusatz,
  ];
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, height: { value: 6900, rule: HeightRule.ATLEAST }, children: [new TableCell({ children: ch, width: { size: W, type: WidthType.DXA }, borders: { top: dash, bottom: dash, right: dash, left: { style: BorderStyle.SINGLE, size: 48, color: NAVY } }, margins: { top: 110, bottom: 90, left: 200, right: 170 }, verticalAlign: VerticalAlign.TOP })] })] });
}
const ZW = [3400, 2400, 2400];
const zaehl = new Table({ width: { size: ZW.reduce((x, y) => x + y, 0), type: WidthType.DXA }, columnWidths: ZW, rows: [
  new TableRow({ children: ["Zählhilfe", "Personen im Stand", "Ja-Stimmen"].map((x, i) => g(x, { w: ZW[i], bold: true })) }),
  ...["Klerus", "Adel", "Dritter Stand"].map((x) => new TableRow({ height: { value: 520, rule: HeightRule.ATLEAST }, children: [g(x, { w: ZW[0], bold: true, align: AlignmentType.LEFT }), ...ZW.slice(1).map((w) => g("", { w }))] })),
] });
const tb = new Table({ width: { size: ZW.reduce((x, y) => x + y, 0), type: WidthType.DXA }, columnWidths: ZW, rows: [
  new TableRow({ children: ["Stand", "Personen", "Ja"].map((x, i) => g(x, { w: ZW[i], bold: true })) }),
  ...[["Klerus", "3"], ["Adel", "4"], ["Dritter Stand", "16"]].map(([x, n]) => new TableRow({ height: { value: 520, rule: HeightRule.ATLEAST }, children: [g(x, { w: ZW[0], bold: true, align: AlignmentType.LEFT }), g(n, { w: ZW[1] }), g("", { w: ZW[2] })] })),
] });
const sp2 = () => p(t(" ", { size: 10 }), { after: 40 });
const seite4 = [
  rolle("Wahlleitung", "ROLLENKARTE · DU LEITEST DIE ABSTIMMUNG · DU STIMMST NICHT MIT", [
    "Frage zuerst: Wie viele Personen hat jeder Stand? Schreibe die Zahlen in die Zählhilfe.",
    "Rufe laut auf: „Klerus. Wer ist für Antrag A? Hand hoch.“ Zähle nur die Hände (nur die Ja-Stimmen).",
    "Rufe das Ergebnis laut aus: „Klerus: Ja …“ Die Schriftführung schreibt es an die Tafel, die Klasse trägt es ein.",
    "Wiederhole für Adel und Dritter Stand.",
    "Sage danach: „Die Abstimmung ist fertig. Rechnet jetzt auf eurem Zettel aus.“ Nein-Stimmen und das Ergebnis musst du nicht ausrechnen.",
    "Bei Problemen frage die Lehrkraft (nur im Notfall).",
  ], [sp2(), zaehl]),
  sp2(),
  rolle("Schriftführung", "ROLLENKARTE · DU SCHREIBST AN DIE TAFEL · DU STIMMST NICHT MIT", [
    "Der Rahmen (Tabelle) steht schon an der Tafel. Warte, bis die Wahlleitung ausruft.",
    "Schreibe die Ja-Zahl in das Feld des Standes.",
    "Schreibe groß und lesbar. Frage nach, wenn du eine Zahl nicht verstanden hast.",
    "Mehr musst du nicht tun. Das Ergebnis rechnet die Klasse selbst aus.",
  ], [sp2(), p(t("So sieht der Rahmen an der Tafel aus:", { size: 19, italics: true, color: MUTED }), { after: 40 }), tb]),
];
const brk = () => new Paragraph({ children: [new PageBreak()] });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 450, left: 600, right: 600 } } }, children: [
    page(pg1), brk(), page(pg2), brk(), bogenPage, brk(), ...seite4,
  ] }] });
  fs.writeFileSync("../Druckvorlage_Karten_Beobachtungszettel_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
