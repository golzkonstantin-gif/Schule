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

// ---- Beobachtungsbogen (zwei pro A4) ----
const g = (txt, o = {}) => new TableCell({ children: [p(t(txt, { size: o.size || 19, bold: o.bold, color: o.color }), { after: 0, align: o.align || AlignmentType.CENTER })], width: { size: o.w, type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 30, bottom: 30, left: 50, right: 50 }, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span });
function bogen() {
  const BW = [1100, 1500, 1500, 1700, 1500, 1700, 1700];
  const tbl = new Table({ width: { size: BW.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: BW, rows: [
    new TableRow({ children: [g("", { w: BW[0] }), g("nach Ständen", { w: BW[1] + BW[2] + BW[3] + BW[4], span: 4, bold: true }), g("nach Köpfen", { w: BW[5] + BW[6], span: 2, bold: true })] }),
    new TableRow({ children: [g("Antrag", { w: BW[0], bold: true, size: 17 }), g("Klerus Ja : Nein", { w: BW[1], size: 17 }), g("Adel Ja : Nein", { w: BW[2], size: 17 }), g("Dritter Stand Ja : Nein", { w: BW[3], size: 17 }), g("Ergebnis", { w: BW[4], size: 17, bold: true }), g("alle Ja : Nein", { w: BW[5], size: 17 }), g("Ergebnis", { w: BW[6], size: 17, bold: true })] }),
    ...["A", "B"].map((x) => new TableRow({ height: { value: 520, rule: HeightRule.ATLEAST }, children: [g(x, { w: BW[0], bold: true }), ...BW.slice(1).map((w) => g("", { w }))] })),
  ] });
  const lines = (txt) => [p([t(txt, { size: 19, bold: true })], { before: 50, after: 0 }), p(t("_".repeat(122), { size: 17, color: "8A96B5" }), { before: 90, after: 0 }), p(t("_".repeat(122), { size: 17, color: "8A96B5" }), { before: 90, after: 0 })];
  return [
    p([t("BEOBACHTUNGSBOGEN · GESCHICHTE 8D   ", { size: 15, bold: true, color: MUTED }), t("Abstimmung der Stände", { font: "Cambria", size: 26, bold: true, color: NAVY })], { after: 40 }),
    p([t("Mein Stand:  ", { size: 19, bold: true }), t("□ Klerus   □ Adel   □ Dritter Stand      ", { size: 19 }), t("Meine Stimme (vorab):  ", { size: 19, bold: true }), t("A □ Ja □ Nein      B □ Ja □ Nein", { size: 19 })], { after: 20 }),
    p([t("A  ", { size: 18, bold: true }), t("Auch Adel und Klerus zahlen künftig Steuern.     ", { size: 18 }), t("B  ", { size: 18, bold: true }), t("Die Stände werden abgeschafft. Alle haben gleiche Rechte.", { size: 18 })], { after: 50 }),
    p(t("Beobachte die Abstimmung. Trage die Zahlen ein, die die Wahlleitung ausruft, und das Ergebnis.", { size: 17, italics: true, color: MUTED }), { after: 30 }),
    tbl,
    p(t("Werte aus", { font: "Cambria", size: 22, bold: true, color: NAVY }), { before: 60, after: 0 }),
    ...lines("1  Klerus und Adel stimmten immer …, weil …"),
    ...lines("2  Nach Ständen wurde Antrag A …, nach Köpfen …"),
    ...lines("3  Das ist für den Dritten Stand (Bauern und Bürger) unfair, weil …* (nenne eine Zahl)"),
  ];
}
function bogenCell() {
  return new TableCell({ children: bogen(), width: { size: W, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: dash, right: dash }, margins: { top: 80, bottom: 60, left: 170, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
const bogenPage = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [0, 1].map(() => new TableRow({ cantSplit: true, height: { value: 7650, rule: HeightRule.EXACT }, children: [bogenCell()] })) });

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
const ZW = [1900, 1500, 1500, 1500, 1500];
const zaehl = new Table({ width: { size: ZW.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: ZW, rows: [
  new TableRow({ children: ["Zählhilfe", "A Ja", "A Nein", "B Ja", "B Nein"].map((x, i) => g(x, { w: ZW[i], bold: true })) }),
  ...["Klerus (3)", "Adel (4)", "Dritter Stand (Rest)", "Summe"].map((x) => new TableRow({ height: { value: 430, rule: HeightRule.ATLEAST }, children: [g(x, { w: ZW[0], bold: true, align: AlignmentType.LEFT }), ...ZW.slice(1).map((w) => g("", { w }))] })),
] });
const TW2 = [1200, 1000, 1000, 1400, 1100, 1000, 1000, 1100];
const tb = new Table({ width: { size: TW2.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: TW2, rows: [
  new TableRow({ children: [g("Antrag", { w: TW2[0], bold: true, size: 17 }), g("nach Ständen", { w: TW2[1] + TW2[2] + TW2[3] + TW2[4], span: 4, bold: true, size: 17 }), g("nach Köpfen", { w: TW2[5] + TW2[6] + TW2[7], span: 3, bold: true, size: 17 })] }),
  new TableRow({ children: ["", "Klerus", "Adel", "Dritter Stand", "Ergebnis", "Ja", "Nein", "Ergebnis"].map((x, i) => g(x, { w: TW2[i], size: 16 })) }),
  ...["A", "B"].map((x) => new TableRow({ height: { value: 430, rule: HeightRule.ATLEAST }, children: [g(x, { w: TW2[0], bold: true }), ...TW2.slice(1).map((w) => g("", { w }))] })),
] });
const sp2 = () => p(t(" ", { size: 10 }), { after: 40 });
const seite4 = [
  rolle("Wahlleitung", "ROLLENKARTE · DU LEITEST DIE ABSTIMMUNG · DU STIMMST NICHT MIT", [
    "Frage zuerst: Wie viele Personen hat jeder Stand? Schreibe die Zahlen in die Zählhilfe.",
    "Stimme offen ab. Rufe laut auf: „Antrag A. Klerus: Wer ist dafür? Hand hoch.“ Zähle die Hände (nur die Ja-Stimmen).",
    "Nein = Personenzahl des Standes minus Ja. Rufe laut aus: „Antrag A, Klerus: Ja …, Nein …“ Die Schriftführung schreibt, die Klasse trägt ein.",
    "Wiederhole für Adel und Dritter Stand, danach für Antrag B.",
    "Nach Ständen: Wie stimmt jeder Stand (Mehrheit im Stand)? Mindestens 2 von 3 Ständen Ja = angenommen.",
    "Nach Köpfen: Zähle alle Ja und alle Nein zusammen. Mehr Ja = angenommen. Lies jedes Ergebnis laut vor.",
    "Bei Problemen frage die Lehrkraft (nur im Notfall).",
  ], [sp2(), zaehl]),
  sp2(),
  rolle("Schriftführung", "ROLLENKARTE · DU SCHREIBST AN DIE TAFEL · DU STIMMST NICHT MIT", [
    "Der Rahmen (Tabelle) steht schon an der Tafel. Warte, bis die Wahlleitung ausruft.",
    "Schreibe die Zahlen in die Felder: Antrag A und B, nach Ständen und nach Köpfen.",
    "Schreibe das Ergebnis dazu: angenommen oder abgelehnt.",
    "Schreibe groß und lesbar. Frage nach, wenn du eine Zahl nicht verstanden hast.",
  ], [sp2(), p(t("So sieht der Rahmen an der Tafel aus:", { size: 19, italics: true, color: MUTED }), { after: 40 }), tb]),
];
const brk = () => new Paragraph({ children: [new PageBreak()] });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 450, left: 600, right: 600 } } }, children: [
    page(pg1), brk(), page(pg2), brk(), bogenPage, brk(), ...seite4,
  ] }] });
  fs.writeFileSync("../Druckvorlage_Karten_Beobachtungsbogen_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
