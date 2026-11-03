const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const STAND = {
  K: { name: "ERSTER STAND · KLERUS", fill: "FFFFFF", edge: "8E6BBE", ink: "6B4A9E",
       rolle: "Du bist Geistlicher. Du zahlst keine Steuern und hast Vorrechte. Du willst sie behalten.",
       tipp: "Du stimmst nach deiner Rolle." },
  A: { name: "ZWEITER STAND · ADEL", fill: "FFFFFF", edge: "E8A33D", ink: "B9791A",
       rolle: "Du bist Adliger. Du zahlst keine Steuern und hast Vorrechte (z. B. Jagdrecht, Ämter am Hof). Du willst sie behalten.",
       tipp: "Du stimmst nach deiner Rolle." },
  D: { name: "DRITTER STAND (BAUERN UND BÜRGER)", fill: "FFFFFF", edge: "3F7CC4", ink: "2F6AAE",
       rolle: "Du zahlst Steuern und Abgaben. Der Brotpreis ist gestiegen, die Familie hat Hunger. Du willst, dass sich etwas ändert.",
       tipp: "Du stimmst nach deiner Überzeugung." },
};
const W = 10706, CW = W / 2;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, border: o.border });
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function card(key, reserve) {
  const s = STAND[key];
  const ch = [
    p(t(reserve ? "STANDESKARTE · RESERVE" : "STANDESKARTE · GEHEIM", { size: 15, bold: true, color: MUTED }), { after: 10 }),
    p(t(s.name, { font: "Cambria", size: 26, bold: true, color: s.ink }), { after: 50 }),
    p(t(s.rolle, { size: 22 }), { after: 40 }),
    p(t(s.tipp, { size: 22, bold: true }), { after: 70 }),
    p([t("STIMMZETTEL  ", { size: 18, bold: true, color: RED }), t("Kreuze bei jedem Antrag an.", { size: 17, color: MUTED })], { after: 30, border: { top: { style: BorderStyle.SINGLE, size: 8, color: s.edge, space: 4 } } }),
    ...[["A", "Steuern auch für Adel und Klerus"], ["B", "Stände abschaffen, gleiche Rechte"], ["C", "Mehr Mitsprache für Bauern"]].map(([l, x]) => p([t(l + "  ", { size: 20, bold: true }), t(x, { size: 18 }), t("      □ Ja  □ Nein", { size: 20, bold: true })], { after: 20 })),
  ];
  return new TableCell({ children: ch, width: { size: CW, type: WidthType.DXA }, borders: { top: dash, bottom: dash, left: { style: BorderStyle.SINGLE, size: 48, color: s.edge }, right: dash }, margins: { top: 90, bottom: 60, left: 200, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
function page(list) {
  const rows = [];
  for (let i = 0; i < list.length; i += 2) rows.push(new TableRow({ cantSplit: true, height: { value: 3090, rule: HeightRule.EXACT }, children: [card(...list[i]), card(...list[i + 1])] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [CW, CW], rows });
}
const pg1 = [["K"], ["K"], ["K"], ["A"], ["A"], ["A"], ["A"], ["D"], ["D"], ["D"]];
const pg2 = Array(10).fill(["D"]);
const pg3 = [...Array(3).fill(["D"]), ...Array(7).fill(["D", true])];


// ---- Seite 4: Rollenkarten Wahlleitung und Schriftführung ----
const thin = { style: BorderStyle.SINGLE, size: 6, color: "8A96B5" };
const g = (txt, o = {}) => new TableCell({ children: [p(t(txt, { size: o.size || 19, bold: o.bold }), { after: 0, align: o.align || AlignmentType.CENTER })], width: { size: o.w, type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 40, bottom: 40, left: 60, right: 60 }, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span });
function rolle(titel, kicker, edge, ink, schritte, zusatz) {
  const ch = [
    p(t(kicker, { size: 16, bold: true, color: MUTED }), { after: 10 }),
    p(t(titel, { font: "Cambria", size: 36, bold: true, color: ink }), { after: 70 }),
    ...schritte.map((x, i) => p([t((i + 1) + "  ", { size: 22, bold: true, color: ink }), t(x, { size: 21 })], { after: 55 })),
    ...zusatz,
  ];
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, height: { value: 7300, rule: HeightRule.ATLEAST }, children: [new TableCell({ children: ch, width: { size: W, type: WidthType.DXA }, borders: { top: dash, bottom: dash, right: dash, left: { style: BorderStyle.SINGLE, size: 48, color: edge } }, margins: { top: 110, bottom: 90, left: 240, right: 200 }, verticalAlign: VerticalAlign.TOP })] })] });
}
const ZW = [1700, 1300, 1300, 1300, 1300, 1300, 1300];
const zaehl = new Table({ width: { size: ZW.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: ZW, rows: [
  new TableRow({ children: ["Zählhilfe", "A Ja", "A Nein", "B Ja", "B Nein", "C Ja", "C Nein"].map((x, i) => g(x, { w: ZW[i], bold: true })) }),
  ...["Klerus", "Adel", "Dritter Stand", "Summe"].map((x) => new TableRow({ height: { value: 430, rule: HeightRule.ATLEAST }, children: [g(x, { w: ZW[0], bold: true, align: AlignmentType.LEFT }), ...ZW.slice(1).map((w) => g("", { w }))] })),
] });
const TW2 = [1200, 900, 900, 1300, 900, 900, 1100, 1000];
const tb = new Table({ width: { size: TW2.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: TW2, rows: [
  new TableRow({ children: [g("Antrag", { w: TW2[0], bold: true, size: 17 }), g("nach Ständen", { w: TW2[1] + TW2[2] + TW2[3] + TW2[4], span: 4, bold: true, size: 17 }), g("nach Köpfen", { w: TW2[5] + TW2[6] + TW2[7], span: 3, bold: true, size: 17 })] }),
  new TableRow({ children: ["", "Klerus", "Adel", "Dritter Stand", "Ergebnis", "Ja", "Nein", "Ergebnis"].map((x, i) => g(x, { w: TW2[i], size: 16 })) }),
  ...["A", "B", "C"].map((x) => new TableRow({ height: { value: 430, rule: HeightRule.ATLEAST }, children: [g(x, { w: TW2[0], bold: true }), ...TW2.slice(1).map((w) => g("", { w }))] })),
] });
const sp2 = () => p(t(" ", { size: 10 }), { after: 40 });
const seite4 = [
  rolle("Wahlleitung", "ROLLENKARTE · DU LEITEST DIE ABSTIMMUNG · DU STIMMST NICHT MIT", "1E2761", "1E2761", [
    "Sage laut: „Stifte liegen. Gebt eure Karte verdeckt nach vorn, Reihe für Reihe.“",
    "Sortiere die Karten in drei Stapel: Klerus, Adel, Dritter Stand (Bauern und Bürger).",
    "Zähle Stapel für Stapel und trage die Zahlen in die Zählhilfe ein.",
    "Diktiere der Schriftführung die Zahlen, Antrag für Antrag: „Antrag A, Klerus: Ja …, Nein …“",
    "Nach Ständen: Wie stimmt jeder Stand (Mehrheit im Stand)? Mindestens 2 von 3 Ständen Ja = angenommen.",
    "Nach Köpfen: Zähle alle Ja und alle Nein zusammen. Mehr Ja = angenommen.",
    "Lies jedes Ergebnis laut vor. Bei Problemen frage die Lehrkraft (nur im Notfall).",
  ], [sp2(), zaehl]),
  sp2(), sp2(),
  rolle("Schriftführung", "ROLLENKARTE · DU SCHREIBST AN DIE TAFEL · DU STIMMST NICHT MIT", "1E2761", "1E2761", [
    "Der Rahmen (Tabelle) steht schon an der Tafel. Warte, bis die Wahlleitung diktiert.",
    "Schreibe die Zahlen in die Felder: Antrag A, B und C, nach Ständen und nach Köpfen.",
    "Schreibe das Ergebnis dazu: angenommen oder abgelehnt.",
    "Schreibe groß und lesbar. Frage nach, wenn du eine Zahl nicht verstanden hast.",
  ], [sp2(), p(t("So sieht der Rahmen an der Tafel aus:", { size: 19, italics: true, color: MUTED }), { after: 40 }), tb]),
];
const sp = () => new Paragraph({ children: [new TextRun({ text: " ", size: 8 })], spacing: { after: 0 } });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 450, left: 600, right: 600 } } }, children: [
    page(pg1), new Paragraph({ children: [new PageBreak()] }), page(pg2), new Paragraph({ children: [new PageBreak()] }), page(pg3), new Paragraph({ children: [new PageBreak()] }), ...seite4,
  ] }] });
  fs.writeFileSync("../Druckvorlage_Standeskarten_Stimmzettel.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
