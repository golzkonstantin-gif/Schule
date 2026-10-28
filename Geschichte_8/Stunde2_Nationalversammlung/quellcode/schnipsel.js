const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, HeightRule } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 50 } });
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const W = 10706;

// Seite 1: Dreischritt-Schnipsel (6 Stück, 2 x 3)
function schnipsel() {
  const bul = (xs) => xs.map((x) => p([t("• ", { size: 19 }), t(x, { size: 19 })], { after: 20 }));
  const ch = [
    p(t("KARIKATUR ANALYSIEREN", { size: 15, bold: true, color: MUTED }), { after: 10 }),
    p([t("1 Beschreiben ", { font: "Cambria", size: 22, bold: true, color: NAVY }), t("Was sehe ich?", { size: 19, italics: true })], { after: 20 }),
    ...bul(["Auf dem Bild sehe ich …", "Im Vordergrund / im Hintergrund …", "Die Person trägt / hält …", "Die Beischrift lautet …"]),
    p([t("2 Deuten ", { font: "Cambria", size: 22, bold: true, color: NAVY }), t("Was soll es bedeuten?", { size: 19, italics: true })], { before: 40, after: 20 }),
    ...bul(["… steht für …", "Der Zeichner will zeigen, dass …", "Die Beischrift sagt …", "Das Bild gehört zu den Ereignissen von …"]),
    p([t("3 Beurteilen ", { font: "Cambria", size: 22, bold: true, color: NAVY }), t("Stimmt das Bild? Was will er?", { size: 19, italics: true })], { before: 40, after: 20 }),
    ...bul(["Das stimmt (nicht) mit unserem Wissen überein, weil …", "Der Zeichner übertreibt, indem …", "Er steht auf der Seite von …", "Er will erreichen, dass …"]),
  ];
  return new TableCell({ children: ch, width: { size: W / 2, type: WidthType.DXA }, borders: all(dash), margins: { top: 90, bottom: 60, left: 170, right: 170 }, verticalAlign: VerticalAlign.TOP });
}
function seite1() {
  const rows = [];
  for (let i = 0; i < 3; i++) rows.push(new TableRow({ cantSplit: true, height: { value: 5100, rule: HeightRule.EXACT }, children: [schnipsel(), schnipsel()] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W / 2, W / 2], rows });
}
// Seite 2: Hilfekarten Stufe 1 bis 3 (je eine Karte, ausschneiden, am Pult)
function hilfe(titel, untertitel, zeilen, fill) {
  return new TableRow({ cantSplit: true, height: { value: 4900, rule: HeightRule.EXACT }, children: [new TableCell({ width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, borders: all(dash), margins: { top: 100, bottom: 80, left: 240, right: 240 }, verticalAlign: VerticalAlign.TOP, children: [
    p(t("HILFEKARTE AM PULT", { size: 15, bold: true, color: MUTED }), { after: 10 }),
    p(t(titel, { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 20 }),
    p(t(untertitel, { size: 20, italics: true, color: MUTED }), { after: 80 }),
    ...zeilen.map((z) => p(z, { after: 70 })),
  ] })] });
}
const q = (s) => [t("• ", { size: 22 }), t(s, { size: 22 })];
const seite2 = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [
  hilfe("Stufe 1: Fragen", "Du kommst nicht weiter? Beantworte der Reihe nach diese Fragen.", [
    q("Wer ist die größte Figur im Bild? Wie sieht sie aus (müde, kräftig, arm)?"),
    q("Was trägt oder hält sie? Was ist um sie herum zu sehen?"),
    q("Was steht über oder unter dem Bild? Lies es noch einmal genau."),
    q("Für wen oder was könnte die Figur stehen: für einen einzelnen Menschen oder für eine ganze Gruppe?"),
    q("Welcher Stand ist gemeint? Denke an Zahlen und Steuern."),
  ], "FFFFFF"),
  hilfe("Stufe 2: Satzanfänge", "Ergänze die Sätze mit deinen eigenen Beobachtungen.", [
    [t("1 ", { bold: true, size: 22 }), t("Auf dem Bild sehe ich …  Die große Figur trägt …  Im Hintergrund …", { size: 22 })],
    [t("2 ", { bold: true, size: 22 }), t("Die Figur steht für …  Die Beischrift sagt, dass …  Das passt zu den Steuern, denn …", { size: 22 })],
    [t("3 ", { bold: true, size: 22 }), t("Der Zeichner will, dass …  Er übertreibt, weil …  Das stimmt (nicht) mit unserem Wissen überein, denn der Dritte Stand (Bauern und Bürger) …", { size: 22 })],
  ], "F4F7FD"),
  hilfe("Stufe 3: Teilanalyse zum Ergänzen", "Eine halbe Analyse. Ergänze die Lücken in deinem Heft.", [
    [t("1 Beschreiben: ", { bold: true, size: 22 }), t("In der Mitte steht ein ______ mit Hut. Er trägt ______ und hält einen ______. Im Hintergrund sieht man ______.", { size: 22 })],
    [t("2 Deuten: ", { bold: true, size: 22 }), t("Der Bauer steht für den ______ Stand (Bauern und Bürger). Die Beischrift sagt, dass er das ganze Jahr arbeitet, um ______ zu bezahlen.", { size: 22 })],
    [t("3 Beurteilen: ", { bold: true, size: 22 }), t("Der Zeichner will, dass die Betrachter ______ empfinden. Das stimmt mit unserem Wissen überein, weil ______ und ______ keine Steuern zahlten. Er übertreibt, weil ______.", { size: 22 })],
  ], "EAF0FA"),
] });
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 450, left: 600, right: 600 } } }, children: [seite1(), new Paragraph({ children: [new PageBreak()] }), seite2] }] });
  fs.writeFileSync("../Schnipsel_Karikatur_Dreischritt_und_Hilfekarten.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
