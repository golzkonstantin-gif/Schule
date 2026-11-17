const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GREEN = "2E9E6B", GOLD = "FFF6E0";
const W = 10700;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line }, keepNext: o.keepNext });
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const brk = () => new Paragraph({ children: [new PageBreak()] });
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("8A96B5", 6)), margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: 100, right: 100 }, verticalAlign: o.valign || VerticalAlign.TOP, columnSpan: o.span });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Stunde 4 Menschenrechte · Seite ", PageNumber.CURRENT] })] })] });

// Quelle (vereinfachte, gekürzte Übersetzung der Erklärung der Menschen- und Bürgerrechte vom 26.8.1789)
const ART = {
  1: "Die Menschen werden frei und gleich an Rechten geboren und bleiben es.",
  2: "Der Zweck jeder politischen Vereinigung ist die Erhaltung der Menschenrechte. Diese sind das Recht auf Freiheit, das Recht auf Eigentum, das Recht auf Sicherheit und das Recht auf Widerstand gegen Unterdrückung.",
  3: "Der Ursprung jeder Herrschaft liegt bei der Nation, also beim Volk.",
  4: "Die Freiheit besteht darin, alles tun zu dürfen, was einem anderen nicht schadet.",
  6: "Das Gesetz muss für alle gleich sein, ob es beschützt oder bestraft. Alle Bürger dürfen nach ihren Fähigkeiten jedes Amt bekommen.",
  10: "Niemand soll wegen seiner Anschauungen, auch nicht wegen seiner religiösen, bestraft werden.",
  11: "Die freie Äußerung von Gedanken und Meinungen ist eines der kostbarsten Menschenrechte. Jeder Bürger kann also frei reden, schreiben und drucken.",
};
const QUELLE = "Quelle: Erklärung der Menschen- und Bürgerrechte, beschlossen von der Nationalversammlung am 26. August 1789 (gekürzt und vereinfacht). Vorbild: amerikanische Unabhängigkeitserklärung von 1776.";
const KOPF = (txt) => p([t("GESCHICHTE 8D · QUELLENARBEIT · " + txt + "   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________", { size: 18, color: MUTED })], { after: 30 });
const artCell = (n, w, fill) => cell([p(t("Artikel " + n, { font: "Cambria", bold: true, size: 22, color: NAVY }), { after: 20 }), p(t("„" + ART[n] + "“", { size: 19 }), { after: 0 })], { w, fill });
const head = (txt, w, sub) => cell([p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { after: sub ? 10 : 0, align: AlignmentType.CENTER }), ...(sub ? [p(t(sub, { color: "FFFFFF", size: 16 }), { after: 0, align: AlignmentType.CENTER })] : [])], { w, fill: NAVY, valign: VerticalAlign.CENTER, borders: all(solid(NAVY, 8)) });
const empty = (w, fill) => cell(p(t(""), { after: 0 }), { w, fill });
const given = (txt, w, col = "1B5E3A") => cell(p(t(txt, { size: 18, color: col }), { after: 0 }), { w, fill: GOLD });

const WORT = "Wortspeicher:  Gleichheit · Freiheit · Volkssouveränität (die Macht im Staat geht vom Volk aus) · Grundrechte (Rechte, die jedem Menschen zustehen) · Religionsfreiheit · Meinungsfreiheit";

function seite1(loesung) {
  const c = [];
  c.push(KOPF("SEITE 1"));
  c.push(p(t("Eine Quelle auswerten: Die Erklärung der Menschenrechte", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 50 }));
  c.push(p(t(QUELLE, { size: 17, color: MUTED }), { after: 60 }));
  c.push(p([t("Aufgabe  ", { bold: true, color: RED, size: 20 }), t("Werte jeden Artikel in drei Schritten aus: ", { size: 20 }), t("1 Lesen", { size: 20, bold: true }), t(" (in eigenen Worten), ", { size: 20 }), t("2 Einordnen", { size: 20, bold: true }), t(" (Was war vorher anders?), ", { size: 20 }), t("3 Begriff", { size: 20, bold: true }), t(" (Wortspeicher).", { size: 20 })], { after: 50 }));
  c.push(p(t(WORT, { size: 17, color: NAVY }), { after: 70 }));
  const CW = [3600, 2700, 2700, 1700];
  const rows = [new TableRow({ children: [head("Quelle (Wortlaut)", CW[0]), head("1 Lesen", CW[1], "in eigenen Worten"), head("2 Einordnen", CW[2], "Vorher anders? Wer gewinnt?"), head("3 Begriff", CW[3])] })];
  const L = {
    1: ["Alle Menschen haben von Geburt an dieselben Rechte.", "Vorher: Geburt entscheidet, zu welchem Stand man gehört (Vorrechte für Klerus und Adel). Wer gewinnt: Dritter Stand.", "Gleichheit"],
    3: ["Die Herrschaft geht vom Volk aus, nicht vom König allein.", "Vorher: König bestimmt allein. Passt zum Ballhausschwur (Macht geht von der Nation aus).", "Volkssouveränität"],
    2: ["Der Staat muss Freiheit, Eigentum, Sicherheit und das Recht auf Widerstand gegen Unterdrückung schützen.", "Vorher: Ungehorsam gegen den König ist Verbrechen. Passt zum Ballhausschwur und zur Bastille.", "Grundrechte"],
    4: ["Jeder darf tun, was anderen nicht schadet. Die Freiheit endet, wo sie anderen schadet.", "Vorher: Der König ließ Gegner ohne Urteil einsperren (Bastille).", "Freiheit"],
    6: ["Gleiches Gesetz für alle. Alle dürfen nach ihren Fähigkeiten ein Amt bekommen.", "Vorher: Klerus und Adel zahlen keine Steuern und bekommen die Ämter.", "Gleichheit"],
  };
  const order = [1, 3, 2, 4, 6];
  order.forEach((n, i) => {
    const fill = i % 2 ? LIGHT : LIGHT2;
    const h = n === 1 ? 2000 : n === 3 ? 2100 : 2250;
    let cs;
    if (n === 1) cs = [given(L[1][0], CW[1]), given(L[1][1], CW[2]), given(L[1][2], CW[3])];
    else if (n === 3 && !loesung) cs = [cell([p(t("Hilfe:  Die Herrschaft geht von … aus, nicht vom …", { size: 17, color: MUTED, italics: true }), { after: 0 })], { w: CW[1] }), cell([p(t("Hilfe:  Vorher bestimmt der … allein. Passt zum …", { size: 17, color: MUTED, italics: true }), { after: 0 })], { w: CW[2] }), given("Volkssouveränität", CW[3])];
    else if (loesung) cs = [given(L[n][0], CW[1]), given(L[n][1], CW[2]), given(L[n][2], CW[3])];
    else cs = [empty(CW[1]), empty(CW[2]), empty(CW[3])];
    rows.push(new TableRow({ cantSplit: true, height: { value: h, rule: "atLeast" }, children: [artCell(n, CW[0], fill), ...cs] }));
  });
  c.push(new Table({ width: { size: W - 200, type: WidthType.DXA }, columnWidths: CW, rows }));
  c.push(p([t("Gelb hinterlegt: Beispiel von der Lehrkraft (Artikel 1) und Hilfe (Artikel 3). Artikel 2, 4 und 6 arbeitest du selbst.", { size: 17, color: MUTED, italics: true })], { before: 60, after: 0 }));
  return c;
}

function seite2(loesung) {
  const c = [];
  c.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 30 }, children: [t("GESCHICHTE 8D · QUELLENARBEIT · SEITE 2   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________", { size: 18, color: MUTED })] }));
  c.push(p(t("Zusatz, Urteil und Rückmeldung", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 70 }));
  // Zusatz
  c.push(p([t("Zusatz *  ", { bold: true, color: RED, size: 20 }), t("Nur wenn du früh fertig bist: Werte die Artikel 10 und 11 genauso aus (Wortspeicher auf Seite 1).", { size: 20 })], { after: 60 }));
  const CW = [4200, 3700, 2800];
  const rows = [new TableRow({ children: [head("Quelle (Wortlaut)", CW[0]), head("Lesen und Einordnen", CW[1]), head("Begriff", CW[2])] })];
  [10, 11].forEach((n, i) => rows.push(new TableRow({ cantSplit: true, height: { value: 1650, rule: "atLeast" }, children: [artCell(n, CW[0], i % 2 ? LIGHT : LIGHT2), loesung ? given(n === 10 ? "Niemand darf wegen seines Glaubens bestraft werden. Vorher: Der katholische Glaube hatte Vorrang, andere Gläubige waren benachteiligt." : "Jeder darf frei sagen, schreiben und drucken, was er denkt. Vorher: Der König kontrollierte, was gedruckt wurde.", CW[1]) : empty(CW[1]), loesung ? given(n === 10 ? "Religionsfreiheit" : "Meinungsfreiheit", CW[2]) : empty(CW[2])] })));
  c.push(new Table({ width: { size: W - 200, type: WidthType.DXA }, columnWidths: CW, rows }));
  // Urteil
  c.push(p([t("Urteil  ", { bold: true, color: RED, size: 22 }), t("Welcher Artikel ist für den Dritten Stand (Bauern und Bürger) am wichtigsten? Schreibe in ganzen Sätzen. Nutze die Satzanfänge.", { size: 20 })], { before: 120, after: 50 }));
  const hilfen = [["Beleg", "Artikel … sagt: „…“"], ["Wirkung", "Das ändert für den Dritten Stand, dass …"], ["Urteil mit Maßstab", "Ich halte Artikel … für am wichtigsten, weil …"], ["Gegenargument", "Dagegen spricht, dass …"]];
  const HW = [2000, 8500];
  c.push(new Table({ width: { size: W - 200, type: WidthType.DXA }, columnWidths: HW, rows: hilfen.map(([a, b], i) => new TableRow({ cantSplit: true, height: { value: 700, rule: "atLeast" }, children: [cell(p(t(a, { bold: true, size: 18, color: NAVY }), { after: 0 }), { w: HW[0], fill: LIGHT2, valign: VerticalAlign.CENTER }), cell(p(t(b, { size: 18, color: MUTED }), { after: 0 }), { w: HW[1], valign: VerticalAlign.TOP })] })) }));
  const lines = loesung
    ? [t("„Artikel 6 sagt: ‚Das Gesetz muss für alle gleich sein.‘ Das ändert für den Dritten Stand, dass Klerus und Adel keine Vorrechte mehr haben. Ich halte Artikel 6 für am wichtigsten, weil der Dritte Stand sonst weiter Steuern zahlt, während andere frei ausgehen. Dagegen spricht, dass Gleichheit auf dem Papier noch nicht heißt, dass alle gleich behandelt werden.“", { size: 18, italics: true, color: "1B5E3A" })]
    : [];
  if (loesung) c.push(new Table({ width: { size: W - 200, type: WidthType.DXA }, columnWidths: [W - 200], rows: [new TableRow({ children: [cell(p(lines, { after: 0 }), { w: W - 200, fill: GOLD })] })] }));
  else {
    c.push(new Table({ width: { size: W - 200, type: WidthType.DXA }, columnWidths: [W - 200], rows: [new TableRow({ cantSplit: true, height: { value: 2600, rule: "atLeast" }, children: [cell(p(t(""), { after: 0 }), { w: W - 200 })] })] }));
  }
  // Rückmeldung
  c.push(p(t("Rückmeldung der Lehrkraft (wird nur bei freiwilliger Abgabe ausgefüllt)", { bold: true, size: 20, color: NAVY }), { before: 140, after: 50 }));
  const RW = [700, 9800];
  const stufen = [["Stufe 1", "Ich kann einen Artikel in eigenen Worten wiedergeben."], ["Stufe 2", "Ich kann einen Artikel mit der Ständeordnung vergleichen (Was war vorher anders?)."], ["Stufe 3", "Ich kann mit einem Beleg aus der Quelle begründen, welcher Artikel am wichtigsten ist."]];
  c.push(new Table({ width: { size: W - 200, type: WidthType.DXA }, columnWidths: RW, rows: stufen.map(([a, b], i) => new TableRow({ cantSplit: true, height: { value: 480, rule: "atLeast" }, children: [cell(p([t("□", { size: 26 })], { after: 0, align: AlignmentType.CENTER }), { w: RW[0], fill: i % 2 ? LIGHT : LIGHT2, valign: VerticalAlign.CENTER }), cell(p([t(a + "   ", { bold: true, size: 19, color: NAVY }), t(b, { size: 19 })], { after: 0 }), { w: RW[1], valign: VerticalAlign.CENTER })] })) }));
  c.push(p([t("□ Ich gebe das Blatt freiwillig ab.", { size: 22, bold: true }), t("          Rückmeldung: ________________________________", { size: 20, color: MUTED })], { before: 100, after: 0 }));
  return c;
}

(async () => {
  const sec = (children) => ({ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 400, left: 600, right: 600 } } }, footers: { default: footer }, children });
  const st = { default: { document: { run: { font: "Calibri", size: 21 } } } };
  const d1 = new Document({ styles: st, sections: [sec([...seite1(false), ...seite2(false)])] });
  fs.writeFileSync("../Arbeitsblatt_Menschenrechte_Quelle_8D.docx", await Packer.toBuffer(d1));
  const d2 = new Document({ styles: st, sections: [sec([p([t("LÖSUNG FÜR DIE LEHRKRAFT · GESCHICHTE 8D · STUNDE 4", { size: 17, bold: true, color: RED }), t("   Musterantworten (Stichworte genügen, andere richtige Formulierungen zählen)", { size: 17, color: MUTED })], { after: 40 }), ...seite1(true), ...seite2(true)])] });
  fs.writeFileSync("../Loesung_Arbeitsblatt_Menschenrechte_8D.docx", await Packer.toBuffer(d2));
  console.log("ok");
})();
