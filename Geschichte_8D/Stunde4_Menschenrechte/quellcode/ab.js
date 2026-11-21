const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GOLD = "FFF6E0";
const W = 10700, TW = W - 200;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line }, keepNext: o.keepNext, pageBreakBefore: o.pb });
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("8A96B5", 6)), margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: 100, right: 100 }, verticalAlign: o.valign || VerticalAlign.TOP });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Stunde 4 Menschenrechte · Seite ", PageNumber.CURRENT] })] })] });

// Quelle: Erklärung der Menschen- und Bürgerrechte vom 26.8.1789, gekürzt und vereinfacht
const ART = {
  1: "Die Menschen werden frei und gleich an Rechten geboren und bleiben es.",
  2: "Der Zweck jeder politischen Vereinigung ist die Erhaltung der Menschenrechte. Diese sind das Recht auf Freiheit, das Recht auf Eigentum, das Recht auf Sicherheit und das Recht auf Widerstand gegen Unterdrückung.",
  3: "Der Ursprung jeder Herrschaft liegt bei der Nation, also beim Volk.",
  4: "Die Freiheit besteht darin, alles tun zu dürfen, was einem anderen nicht schadet.",
  6: "Das Gesetz muss für alle gleich sein, ob es beschützt oder bestraft. Alle Bürger dürfen nach ihren Fähigkeiten jedes Amt bekommen.",
  10: "Niemand soll wegen seiner Anschauungen, auch nicht wegen seiner religiösen, bestraft werden.",
  11: "Die freie Äußerung von Gedanken und Meinungen ist eines der kostbarsten Menschenrechte. Jeder Bürger kann also frei reden, schreiben und drucken.",
  17: "Das Eigentum ist ein unverletzliches und heiliges Recht. Es darf niemandem genommen werden.",
};
const QUELLE = "Quelle: Erklärung der Menschen- und Bürgerrechte, beschlossen von der Nationalversammlung am 26. August 1789 (gekürzt und vereinfacht). Vorbild: amerikanische Unabhängigkeitserklärung von 1776.";
const WORT = "Wortspeicher:  Gleichheit · Freiheit · Volkssouveränität (die Macht im Staat geht vom Volk aus) · Grundrechte · Religionsfreiheit · Meinungsfreiheit · Eigentumsrecht";
const head = (txt, w, sub) => cell([p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { after: sub ? 10 : 0, align: AlignmentType.CENTER }), ...(sub ? [p(t(sub, { color: "FFFFFF", size: 16 }), { after: 0, align: AlignmentType.CENTER })] : [])], { w, fill: NAVY, valign: VerticalAlign.CENTER, borders: all(solid(NAVY, 8)) });
const artCell = (n, w, fill) => cell([p(t("Artikel " + n, { font: "Cambria", bold: true, size: 21, color: NAVY }), { after: 15 }), p(t("„" + ART[n] + "“", { size: 18 }), { after: 0 })], { w, fill });
const empty = (w, fill) => cell(p(t(""), { after: 0 }), { w, fill });
const given = (txt, w) => cell(p(t(txt, { size: 18, color: "1B5E3A" }), { after: 0 }), { w, fill: GOLD });
const hint = (txt, w) => cell(p(t(txt, { size: 16, color: MUTED, italics: true }), { after: 0 }), { w });
const KOPF = (txt, pb) => p([t("GESCHICHTE 8D · QUELLENARBEIT · " + txt + "   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________", { size: 18, color: MUTED })], { after: 30, pb });

const LOES = {
  1: ["Gleichheit", "Die Geburt entscheidet über den Stand und die Vorrechte."],
  3: ["Volkssouveränität", "Der König bestimmt allein."],
  2: ["Grundrechte (Freiheit, Eigentum, Sicherheit, Widerstand)", "Ungehorsam gegen den König ist ein Verbrechen."],
  4: ["Freiheit", "Der König lässt Gegner ohne Urteil einsperren (Bastille)."],
  6: ["Gleichheit vor dem Gesetz", "Klerus und Adel zahlen keine Steuern, Ämter bekommen vor allem Adlige."],
  10: ["Religionsfreiheit", "Der katholische Glaube hat Vorrang, andere Gläubige sind benachteiligt."],
  11: ["Meinungsfreiheit", "Der König kontrolliert, was gedruckt wird (Zensur)."],
  17: ["Eigentumsrecht", "Bauern müssen einen Teil der Ernte und Dienste an Adel und Kirche abgeben, ihr Eigentum ist nicht gesichert."],
};
function seite1(loesung) {
  const c = [];
  c.push(KOPF("SEITE 1", false));
  c.push(p(t("Die Erklärung der Menschenrechte ins Heute übersetzen", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 50 }));
  c.push(p(t(QUELLE, { size: 17, color: MUTED }), { after: 60 }));
  c.push(p([t("Aufgabe  ", { bold: true, color: RED, size: 20 }), t("1  Lies den Artikel.   2  Welcher heutige Begriff passt? (Wortspeicher)   3  Was war vorher anders? Schreibe es bei den Artikeln 6, 11 und 17 auf.", { size: 20 })], { after: 50 }));
  c.push(p(t(WORT, { size: 17, color: NAVY }), { after: 70 }));
  const CW = [4300, 2200, 4000];
  const rows = [new TableRow({ children: [head("Quelle (Wortlaut)", CW[0]), head("Begriff heute", CW[1], "aus dem Wortspeicher"), head("Was war vorher anders?", CW[2], "Stichworte")] })];
  const order = [1, 3, 2, 4, 6, 10, 11, 17];
  const H = { 1: 1250, 3: 1350, 2: 1700, 4: 1250, 6: 1500, 10: 1250, 11: 1500, 17: 1250 };
  const must = [6, 11, 17];
  order.forEach((n, i) => {
    const fill = i % 2 ? LIGHT : LIGHT2;
    let cs;
    if (n === 1) cs = [given(LOES[1][0], CW[1]), given(LOES[1][1], CW[2])];
    else if (n === 3) cs = [given(LOES[3][0], CW[1]), loesung ? given(LOES[3][1], CW[2]) : cell(p(t("Hilfe:  Vorher bestimmt der … allein.", { size: 17, color: MUTED, italics: true }), { after: 0 }), { w: CW[2] })];
    else if (loesung) cs = [given(LOES[n][0], CW[1]), must.includes(n) ? given(LOES[n][1], CW[2]) : given("Zusatz*: " + LOES[n][1], CW[2])];
    else cs = [empty(CW[1]), must.includes(n) ? empty(CW[2]) : hint("Zusatz* (nur wenn du früh fertig bist)", CW[2])];
    rows.push(new TableRow({ cantSplit: true, height: { value: H[n], rule: "atLeast" }, children: [artCell(n, CW[0], fill), ...cs] }));
  });
  c.push(new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: CW, rows }));
  c.push(p(t("Gelb hinterlegt: Beispiel der Lehrkraft (Artikel 1) und Hilfe (Artikel 3). Artikel 2, 4, 6, 10, 11 und 17 arbeitest du selbst.", { size: 17, color: MUTED, italics: true }), { before: 50, after: 0 }));
  return c;
}

const AUSS = [
  ["Das Gesetz muss für alle gleich sein.", "1789", "Art. 6"],
  ["Alle Staatsgewalt geht vom Volke aus.", "heute", "Grundgesetz Art. 20 (2)"],
  ["Jeder Bürger kann frei reden, schreiben und drucken.", "1789", "Art. 11"],
  ["Das Eigentum und das Erbrecht werden gewährleistet.", "heute", "Grundgesetz Art. 14 (1)"],
  ["Alle Menschen sind vor dem Gesetz gleich.", "heute", "Grundgesetz Art. 3 (1)"],
  ["Der Ursprung jeder Herrschaft liegt bei der Nation, also beim Volk.", "1789", "Art. 3"],
  ["Das Eigentum ist ein unverletzliches und heiliges Recht.", "1789", "Art. 17"],
  ["Jeder hat das Recht, seine Meinung in Wort, Schrift und Bild frei zu äußern.", "heute", "Grundgesetz Art. 5 (1)"],
];
function seite2(loesung) {
  const c = [];
  c.push(KOPF("SEITE 2", true));
  c.push(p(t("Aus welchem Jahr? Und was bedeutet das?", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 60 }));
  c.push(p([t("Aufgabe A  ", { bold: true, color: RED, size: 20 }), t("Welche Sätze stammen aus der Erklärung von 1789, welche aus dem Grundgesetz von heute? Kreuze an. Rate, wenn du unsicher bist.", { size: 20 })], { after: 60 }));
  const AW = [500, 7300, 1350, 1350];
  const rows = [new TableRow({ children: [head("", AW[0]), head("Aussage", AW[1]), head("1789", AW[2]), head("heute", AW[3])] })];
  AUSS.forEach(([a, j, q], i) => rows.push(new TableRow({ cantSplit: true, height: { value: 620, rule: "atLeast" }, children: [
    cell(p(t(String.fromCharCode(65 + i), { bold: true, size: 20, color: NAVY }), { after: 0, align: AlignmentType.CENTER }), { w: AW[0], fill: i % 2 ? LIGHT : LIGHT2, valign: VerticalAlign.CENTER }),
    cell(p(t("„" + a + "“", { size: 19 }), { after: 0 }), { w: AW[1], valign: VerticalAlign.CENTER, fill: loesung ? GOLD : undefined }),
    cell(p(t(loesung ? (j === "1789" ? "☒ " + q : "□") : "□", { size: loesung ? 16 : 26, color: loesung ? "1B5E3A" : "000000" }), { after: 0, align: AlignmentType.CENTER }), { w: AW[2], valign: VerticalAlign.CENTER }),
    cell(p(t(loesung ? (j === "heute" ? "☒ " + q : "□") : "□", { size: loesung ? 16 : 26, color: loesung ? "1B5E3A" : "000000" }), { after: 0, align: AlignmentType.CENTER }), { w: AW[3], valign: VerticalAlign.CENTER }),
  ] })));
  c.push(new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: AW, rows }));
  c.push(p([t("Aufgabe B  ", { bold: true, color: RED, size: 20 }), t("Historische Bedeutung: Was haben die Menschenrechte für den Dritten Stand (Bauern und Bürger) verbessert? Schreibe in ganzen Sätzen. Nutze die Satzanfänge.", { size: 20 })], { before: 130, after: 50 }));
  const hilfen = [["Vorher", "Vor 1789 mussten die Menschen im Dritten Stand …"], ["Seit 1789", "Durch die Erklärung der Menschenrechte …  (Artikel …)"], ["Heute", "Noch heute steht im Grundgesetz …"]];
  const HW = [1500, 9000];
  c.push(new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: HW, rows: hilfen.map(([a, b]) => new TableRow({ cantSplit: true, height: { value: 800, rule: "atLeast" }, children: [cell(p(t(a, { bold: true, size: 18, color: NAVY }), { after: 0 }), { w: HW[0], fill: LIGHT2, valign: VerticalAlign.CENTER }), loesung ? given(a === "Vorher" ? "Vor 1789 mussten die Menschen im Dritten Stand Steuern und Abgaben zahlen, während Klerus und Adel Vorrechte hatten, und der König bestimmte allein." : a === "Seit 1789" ? "Durch die Erklärung der Menschenrechte gilt das Gesetz für alle gleich (Artikel 6), man darf frei reden und drucken (Artikel 11), das Eigentum ist geschützt (Artikel 17), und die Macht geht vom Volk aus (Artikel 3)." : "Noch heute steht im Grundgesetz: Alle Menschen sind vor dem Gesetz gleich (Artikel 3), jeder darf seine Meinung frei äußern (Artikel 5), die Staatsgewalt geht vom Volke aus (Artikel 20).", HW[1]) : cell(p(t(b, { size: 17, color: MUTED }), { after: 0 }), { w: HW[1] })] })) }));
  c.push(p(t("Rückmeldung der Lehrkraft (nur bei freiwilliger Abgabe)", { bold: true, size: 20, color: NAVY }), { before: 120, after: 50 }));
  const RW = [700, 9800];
  const stufen = [["Stufe 1", "Ich kann den Verzicht auf die Vorrechte und die Erklärung der Menschenrechte beschreiben und sagen, was sich für den Dritten Stand verbessert hat."], ["Stufe 2", "Ich kann Artikel in heutige Begriffe übersetzen (Gleichheit vor dem Gesetz, Eigentumsrecht, Meinungsfreiheit, Volkssouveränität)."], ["Stufe 3", "Ich kann an einem Grundgesetz-Artikel zeigen, dass die Erklärung bis heute nachwirkt.*"]];
  c.push(new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: RW, rows: stufen.map(([a, b], i) => new TableRow({ cantSplit: true, height: { value: 480, rule: "atLeast" }, children: [cell(p([t("□", { size: 26 })], { after: 0, align: AlignmentType.CENTER }), { w: RW[0], fill: i % 2 ? LIGHT : LIGHT2, valign: VerticalAlign.CENTER }), cell(p([t(a + "   ", { bold: true, size: 18, color: NAVY }), t(b, { size: 18 })], { after: 0 }), { w: RW[1], valign: VerticalAlign.CENTER })] })) }));
  c.push(p([t("□ Ich gebe das Blatt freiwillig ab.", { size: 22, bold: true }), t("          Rückmeldung: ________________________________", { size: 20, color: MUTED })], { before: 90, after: 0 }));
  return c;
}

(async () => {
  const sec = (children) => ({ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 400, left: 600, right: 600 } } }, footers: { default: footer }, children });
  const st = { default: { document: { run: { font: "Calibri", size: 21 } } } };
  fs.writeFileSync("../Arbeitsblatt_Menschenrechte_Quelle_8D.docx", await Packer.toBuffer(new Document({ styles: st, sections: [sec([...seite1(false), ...seite2(false)])] })));
  fs.writeFileSync("../Loesung_Arbeitsblatt_Menschenrechte_8D.docx", await Packer.toBuffer(new Document({ styles: st, sections: [sec([p([t("LÖSUNG FÜR DIE LEHRKRAFT · GESCHICHTE 8D · STUNDE 4", { size: 17, bold: true, color: RED }), t("   Musterantworten (Stichworte genügen, andere richtige Formulierungen zählen)", { size: 17, color: MUTED })], { after: 40 }), ...seite1(true), ...seite2(true)])] })));
  console.log("ok");
})();
