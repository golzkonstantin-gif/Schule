const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GREEN = "2E9E6B";
const W = 10700;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line }, keepNext: o.keepNext });
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const brk = () => new Paragraph({ children: [new PageBreak()] });

const TEXT = {
  A: { titel: "Text A: Der Ballhausschwur (20. Juni 1789)", farbe: GREEN, sätze: [
    "Am 17. Juni 1789 erklärt sich der Dritte Stand (Bauern und Bürger) zur Nationalversammlung, weil er nicht länger überstimmt werden will.",
    "Der König ist dagegen und lässt den Sitzungssaal in Versailles sperren.",
    "Am 20. Juni treffen sich die Abgeordneten in einer Ballhalle in der Nähe.",
    "Dort schwören sie, nicht auseinanderzugehen, bis Frankreich eine Verfassung hat.",
    "Sie benutzen keine Waffen, nur Worte und ihren Eid.",
    "Es gibt keine Toten, aber die Abgeordneten riskieren, dass der König sie auflösen lässt.",
    "Der König kann sie nicht zum Gehen zwingen.",
    "Eine Woche später fordert er Klerus und Adel auf, sich der Nationalversammlung anzuschließen."],
    wörter: "Ballhalle: Halle, in der man Ballspiele spielte · Eid: feierliches Versprechen · Verfassung: Gesetze, die festlegen, wer wie regiert" },
  B: { titel: "Text B: Der Sturm auf die Bastille (14. Juli 1789)", farbe: RED, sätze: [
    "In Paris herrscht Hunger, das Brot ist teuer.",
    "Anfang Juli zieht der König Soldaten um Paris zusammen und entlässt den beliebten Finanzminister.",
    "Die Pariser haben Angst, dass die Soldaten gegen die Stadt eingesetzt werden.",
    "Am 14. Juli holen sich viele Menschen Gewehre und suchen Pulver.",
    "Das Pulver liegt in der Bastille, einer Festung und einem Gefängnis des Königs.",
    "Die Menge greift die Bastille an, es wird geschossen.",
    "Nach Stunden ergibt sich der Gouverneur, doch die Menge tötet ihn danach.",
    "Etwa 100 Angreifer sterben. In der Bastille sind nur sieben Gefangene.",
    "Der König zieht danach die Soldaten ab. Die Bastille wird zum Symbol der Revolution."],
    wörter: "Bastille: Festung und Gefängnis des Königs in Paris · Gouverneur: Befehlshaber der Festung · Pulver: Schießpulver für Waffen" },
};
const KRIT = [
  ["1  Vorgeschichte", "Warum kam es dazu?", { A: "Sätze 1 bis 2", B: "Sätze 1 bis 3" }],
  ["2  Verlauf", "Was geschah?", { A: "Sätze 3 bis 4", B: "Sätze 4 bis 7" }],
  ["3  Mittel", "Gewalt: ja oder nein?", { A: "Satz 5", B: "Satz 6" }],
  ["4  Opfer und Folgen", "Wer litt oder riskierte etwas?", { A: "Satz 6", B: "Sätze 7 bis 8" }],
  ["5  Wirkung", "Was änderte sich?", { A: "Sätze 7 bis 8", B: "Satz 9" }],
];
const LOES = {
  1: ["Der Dritte Stand erklärt sich zur Nationalversammlung, der König lässt den Saal sperren.", "Hunger, Brot teuer; Soldaten um Paris, Finanzminister entlassen, Angst vor dem Militär."],
  2: ["Die Abgeordneten treffen sich in einer Ballhalle und schwören, nicht auseinanderzugehen, bis Frankreich eine Verfassung hat.", "Menschen holen Gewehre, suchen Pulver, greifen die Bastille an, der Gouverneur ergibt sich."],
  3: ["Gewaltlos: Worte und Eid, keine Waffen.", "Gewaltsam: Waffen, Schüsse, Kampf."],
  4: ["Keine Toten; Risiko: Auflösung durch den König.", "Etwa 100 Tote unter den Angreifern, der Gouverneur wird getötet, nur sieben Gefangene."],
  5: ["Die Versammlung bleibt, der König kann sie nicht zum Gehen zwingen, später soll sich der Adel anschließen.", "Der König zieht die Soldaten ab, die Bastille wird zum Symbol der Revolution."],
};

const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("8A96B5", 6)), margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: 110, right: 110 }, verticalAlign: o.valign || VerticalAlign.TOP });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Stunde 3 Widerstand · Seite ", PageNumber.CURRENT] })] })] });

function blatt(me) {
  const other = me === "A" ? "B" : "A";
  const T = TEXT[me];
  const c = [];
  c.push(p([t("GESCHICHTE 8D · PARTNERPUZZLE   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________", { size: 18, color: MUTED })], { after: 30 }));
  c.push(p(t(T.titel, { font: "Cambria", size: 30, bold: true, color: T.farbe }), { after: 70 }));
  T.sätze.forEach((s, i) => c.push(p([t((i + 1) + "  ", { bold: true, size: 20, color: T.farbe }), t(s, { font: "Cambria", size: 24 })], { after: 60, line: 288 })));
  c.push(p(t(T.wörter, { size: 17, color: MUTED }), { before: 40, after: 80 }));
  c.push(p([t("Aufgabe  ", { bold: true, color: RED, size: 20 }), t("1  Lies deinen Text und fülle deine Spalte aus (Stichworte).   2  Höre deinem Partner zu und fülle seine Spalte aus.", { size: 20 })], { after: 90 }));
  const CW = [2500, 4100, 4100];
  const head = (txt, col, w) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w, fill: col, valign: VerticalAlign.CENTER, borders: all(solid(col, 8)) });
  const rows = [new TableRow({ children: [head("Kriterium", NAVY, CW[0]), head("Text A: Ballhausschwur" + (me === "A" ? " (dein Text)" : " (Partner)"), GREEN, CW[1]), head("Text B: Bastille" + (me === "B" ? " (dein Text)" : " (Partner)"), RED, CW[2])] })];
  KRIT.forEach(([n, q, hint], i) => rows.push(new TableRow({ cantSplit: true, height: { value: 1620, rule: "atLeast" }, children: [
    cell([p(t(n, { font: "Cambria", bold: true, size: 21, color: NAVY }), { after: 10 }), p(t(q, { size: 18 }), { after: 10 }), p(t("Hilfe: " + hint[me], { size: 16, color: MUTED, italics: true }), { after: 0 })], { w: CW[0], fill: i % 2 ? LIGHT : LIGHT2 }),
    cell(p(t(""), { after: 0 }), { w: CW[1], fill: me === "A" ? "FFFFFF" : "F7F7F7" }),
    cell(p(t(""), { after: 0 }), { w: CW[2], fill: me === "B" ? "FFFFFF" : "F7F7F7" }),
  ] })));
  c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: CW, rows }));
  c.push(p([t("□ Ich gebe das Blatt freiwillig ab.", { size: 20, bold: true }), t("          Punkte: ____ von 10", { size: 20 })], { before: 100, after: 0 }));
  return c;
}
function loesung() {
  const c = [];
  c.push(p(t("LÖSUNG FÜR DIE LEHRKRAFT · GESCHICHTE 8D · STUNDE 3", { size: 17, bold: true, color: RED }), { after: 20 }));
  c.push(p(t("Partnerpuzzle: Erwartungen und Punkte", { font: "Cambria", size: 34, bold: true, color: NAVY }), { after: 60 }));
  c.push(p(t("Je Feld ein Punkt (zwei Texte mal fünf Kriterien = 10 Punkte), wenn die Kernaussage im Feld steht (Stichworte genügen). Das Sachurteil und das Werturteil (Folie 17) zählen nicht zu den 10 Punkten. Die Zahlen zur Bastille bitte gegen das Lehrbuch prüfen. Notenschlüssel nach Ihrer Festlegung (Raster auf Folie 7: 0 bis 4, 5 bis 8, 9 bis 10 Punkte).", { size: 20, color: MUTED }), { after: 120 }));
  const CW = [2300, 4200, 4200];
  const head = (txt, col, w) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w, fill: col, valign: VerticalAlign.CENTER, borders: all(solid(col, 8)) });
  const rows = [new TableRow({ children: [head("Kriterium", NAVY, CW[0]), head("Ballhausschwur", GREEN, CW[1]), head("Bastille", RED, CW[2])] })];
  KRIT.forEach(([n, q], i) => rows.push(new TableRow({ cantSplit: true, children: [
    cell([p(t(n, { font: "Cambria", bold: true, size: 21, color: NAVY }), { after: 10 }), p(t(q, { size: 17 }), { after: 0 })], { w: CW[0], fill: i % 2 ? LIGHT : LIGHT2 }),
    cell(p(t(LOES[i + 1][0], { size: 20 }), { after: 0 }), { w: CW[1] }),
    cell(p(t(LOES[i + 1][1], { size: 20 }), { after: 0 }), { w: CW[2] }),
  ] })));
  c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: CW, rows }));
  c.push(p(t("Hinweise zur Haltung", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 60 }));
  ["Wirksam heißt nicht gerechtfertigt: Beide Wege bewirkten etwas, aber sie kosteten Unterschiedliches (Bilanz, Folie 16).", "Aha: Die Reihenfolge zeigt, dass der gewaltlose Schritt die Grundlage schuf und der Druck in Paris den König zum Rückzug zwang. Der Streit um die Frage, ob die Gewalt den König wirklich allein zum Nachgeben brachte, ist offen; die Klasse soll die Reihenfolge belegen, nicht Schuld verteilen.", "Gewalt war eine Reaktion auf die drohende Gewalt (Soldaten), das ist ein Anlass, keine Rechtfertigung. Die Aufgabe, ein Urteil mit einem offenen Maßstab (wirksam, gerecht, verantwortbar) zu begründen, lässt beide Antworten zu.", "Wirkung beim Ballhausschwur: Nach dem Schwur versuchte der König noch, die Beschlüsse zu kassieren (23. Juni); die Abgeordneten blieben. Am 27. Juni forderte er Klerus und Adel auf, sich anzuschließen. Zum Prüfen mit dem Lehrbuch."].forEach((x) => c.push(p([t("• ", { size: 20 }), t(x, { size: 20 })], { after: 50 })));
  return c;
}
(async () => {
  const sec = (children) => ({ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 650, bottom: 600, left: 600, right: 600 } } }, footers: { default: footer }, children });
  const d1 = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [sec([...blatt("A"), brk(), ...blatt("B")])] });
  fs.writeFileSync("../Partnerpuzzle_Blatt_A_und_B_8D.docx", await Packer.toBuffer(d1));
  const d2 = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [sec(loesung())] });
  fs.writeFileSync("../Loesung_Partnerpuzzle_8D.docx", await Packer.toBuffer(d2));
  console.log("ok");
})();
