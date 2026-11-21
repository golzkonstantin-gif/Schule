const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, BorderStyle, ShadingType, AlignmentType, TabStopType, LeaderType, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT2 = "EAF0FA", RED = "D9534F", GOLDF = "FFF6E0", GREEN = "1B5E3A";
const R = 10500; // Zeilenbreite (Tabstopp rechts)
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line }, keepNext: o.keepNext, pageBreakBefore: o.pb, indent: o.indent, border: o.border, shading: o.shading, tabStops: o.tabs, keepLines: true });
const bd = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c, space: 4 });
const box = (c) => ({ top: bd(c), bottom: bd(c), left: bd(c), right: bd(c) });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Stunde 4 Menschenrechte · Seite ", PageNumber.CURRENT] })] })] });
const KOPF = (txt, pb) => p([t("GESCHICHTE 8D · QUELLENARBEIT · " + txt + "   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________", { size: 18, color: MUTED })], { after: 40, pb });
// Schreiblinie mit Beschriftung
const lineTabs = [{ type: TabStopType.RIGHT, position: R, leader: LeaderType.UNDERSCORE }];
const wline = (label, text, loesung, hint) => p(loesung && text ? [t(label + " ", { bold: true, size: 20, color: NAVY }), t(text, { size: 20, color: GREEN, italics: true })] : [t(label + " ", { bold: true, size: 19, color: NAVY }), ...(hint ? [t(hint + " ", { size: 17, color: MUTED, italics: true })] : []), t("\t")], { tabs: lineTabs, after: 140, before: 70 });
const emptyLine = () => p([t("\t")], { tabs: lineTabs, after: 140, before: 70 });

const ART = {
  1: "Die Menschen werden frei und gleich an Rechten geboren und bleiben es.",
  3: "Der Ursprung jeder Herrschaft liegt bei der Nation, also beim Volk.",
  6: "Das Gesetz muss für alle gleich sein, ob es beschützt oder bestraft. Alle Bürger dürfen nach ihren Fähigkeiten jedes Amt bekommen.",
  11: "Die freie Äußerung von Gedanken und Meinungen ist eines der kostbarsten Menschenrechte. Jeder Bürger kann also frei reden, schreiben und drucken.",
  17: "Das Eigentum ist ein unverletzliches und heiliges Recht. Es darf niemandem genommen werden.",
};
const QUELLE = "Quelle: Erklärung der Menschen- und Bürgerrechte, 26. August 1789 (gekürzt und vereinfacht). Vorbild: amerikanische Unabhängigkeitserklärung von 1776.";
const WORT = [
  ["Gleichheit", "Alle haben dieselben Rechte, egal woher sie kommen."],
  ["Meinungsfreiheit", "Man darf frei sagen und schreiben, was man denkt."],
  ["Eigentumsrecht", "Was jemandem gehört, darf ihm niemand wegnehmen."],
  ["Volkssouveränität", "Die Macht im Staat geht vom Volk aus."],
];
const OPT = "□ Gleichheit     □ Meinungsfreiheit     □ Eigentumsrecht     □ Volkssouveränität";
const L = {
  1: ["Alle Menschen haben von Geburt an dieselben Rechte.", "Gleichheit", "Die Geburt entschied über den Stand und die Vorrechte."],
  3: ["Die Herrschaft geht vom Volk aus, nicht vom König allein.", "Volkssouveränität", "Der König bestimmte allein."],
  6: ["Für alle gilt dasselbe Gesetz. Wer geeignet ist, kann ein Amt bekommen.", "Gleichheit", "Klerus und Adel zahlten keine Steuern, Ämter bekamen vor allem Adlige."],
  11: ["Jeder darf frei sagen, schreiben und drucken, was er denkt.", "Meinungsfreiheit", "Der König kontrollierte, was gedruckt wurde (Zensur)."],
  17: ["Was jemandem gehört, darf ihm niemand wegnehmen.", "Eigentumsrecht", "Bauern mussten einen Teil der Ernte und Dienste an Adel und Kirche abgeben, ihr Eigentum war nicht gesichert."],
};
const optLine = (right, loesung, gold) => {
  const names = ["Gleichheit", "Meinungsfreiheit", "Eigentumsrecht", "Volkssouveränität"];
  const runs = [t("② Das heißt heute:  ", { bold: true, size: 19, color: NAVY })];
  names.forEach((n) => { const on = (loesung || gold) && n === right; runs.push(t((on ? "☒ " : "□ ") + n + "     ", { size: 19, bold: on, color: on ? GREEN : "000000" })); });
  return p(runs, { after: 130, before: 60 });
};
// Karte für einen Artikel
function karte(n, loesung) {
  const c = [];
  const gold = n === 1 || n === 3; // Beispiel und Hilfe
  c.push(p([t("Artikel " + n + "   ", { font: "Cambria", bold: true, size: 22, color: NAVY }), t("„" + ART[n] + "“", { size: 19 })], { before: 130, after: 60, border: box("8A96B5"), shading: { type: ShadingType.CLEAR, fill: LIGHT2, color: "auto" }, indent: { left: 80, right: 80 }, keepNext: true }));
  const show = loesung || n === 1;
  if (n === 3 && !loesung) c.push(wline("① In meinen Worten:", "", false, "Die Herrschaft geht von … aus, nicht vom … allein."));
  else c.push(wline("① In meinen Worten:", L[n][0], show));
  if (!(show || n === 3) || n === 6 || n === 11 || n === 17) { if (!loesung) c.push(emptyLine()); }
  c.push(optLine(L[n][1], loesung, n === 1 || n === 3));
  if (n === 3 && !loesung) c.push(wline("③ Vorher war es anders:", "", false, "Vorher bestimmte der … allein."));
  else c.push(wline("③ Vorher war es anders:", L[n][2], show));
  if (!show && n !== 3) c.push(emptyLine());
  return c;
}
function seite1(loesung) {
  const c = [];
  c.push(KOPF("SEITE 1", false));
  c.push(p(t("Die Erklärung der Menschenrechte ins Heute übersetzen", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 40 }));
  c.push(p(t(QUELLE, { size: 17, color: MUTED }), { after: 50 }));
  c.push(p([t("Aufgabe  ", { bold: true, color: RED, size: 20 }), t("Gehe bei jedem Artikel in drei Schritten vor:  ", { size: 20 }), t("① Schreibe in eigenen Worten, was dort steht.  ② Kreuze an, wie wir heute dazu sagen.  ③ Schreibe, was vorher anders war.", { size: 20, bold: true })], { after: 60 }));
  c.push(p([t("Wortspeicher", { bold: true, size: 19, color: NAVY })], { after: 20, border: { top: bd("8A96B5"), left: bd("8A96B5"), right: bd("8A96B5"), bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF", space: 0 } }, shading: { type: ShadingType.CLEAR, fill: GOLDF, color: "auto" }, indent: { left: 80, right: 80 } }));
  WORT.forEach(([a, b], i) => c.push(p([t(a + ":  ", { bold: true, size: 18, color: NAVY }), t(b, { size: 18 })], { after: i === WORT.length - 1 ? 60 : 10, border: { left: bd("8A96B5"), right: bd("8A96B5"), bottom: i === WORT.length - 1 ? bd("8A96B5") : { style: BorderStyle.NONE, size: 0, color: "FFFFFF", space: 0 }, top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF", space: 0 } }, shading: { type: ShadingType.CLEAR, fill: GOLDF, color: "auto" }, indent: { left: 80, right: 80 } })));
  [1, 3, 6, 11, 17].forEach((n) => c.push(...karte(n, loesung)));
  c.push(p(t("Beispiel (Artikel 1) und Hilfe (Artikel 3) sind vorgegeben. Artikel 6, 11 und 17 arbeitest du selbst.", { size: 16, color: MUTED, italics: true }), { before: 50, after: 0 }));
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
  c.push(p([t("Aufgabe A  ", { bold: true, color: RED, size: 20 }), t("Welche Sätze stammen aus der Erklärung von 1789, welche aus dem Grundgesetz von heute? Kreuze an. Rate, wenn du unsicher bist.", { size: 20 })], { after: 90 }));
  AUSS.forEach(([a, j, q], i) => {
    const x1 = loesung && j === "1789", x2 = loesung && j === "heute";
    c.push(p([t(String.fromCharCode(65 + i) + "   ", { bold: true, size: 20, color: NAVY }), t("„" + a + "“", { size: 20 }), t("\t"), t((x1 ? "☒" : "□") + " 1789", { size: 20, bold: x1, color: x1 ? GREEN : "000000" }), t("      "), t((x2 ? "☒" : "□") + " heute", { size: 20, bold: x2, color: x2 ? GREEN : "000000" }), ...(loesung ? [t("   (" + q + ")", { size: 15, color: MUTED })] : [])], { tabs: [{ type: TabStopType.LEFT, position: 7000 }], after: 200, border: { bottom: bd("D0D7E8", 4) } }));
  });
  c.push(p([t("Aufgabe B  ", { bold: true, color: RED, size: 20 }), t("Historische Bedeutung: Was haben die Menschenrechte für den Dritten Stand (Bauern und Bürger) verbessert? Schreibe in ganzen Sätzen. Nutze die Satzanfänge.", { size: 20 })], { before: 160, after: 60 }));
  const B = [["Vorher:", "Vor 1789 mussten die Menschen im Dritten Stand …", "Vor 1789 mussten die Menschen im Dritten Stand Steuern und Abgaben zahlen, während Klerus und Adel Vorrechte hatten, und der König bestimmte allein."], ["Seit 1789:", "Durch die Erklärung der Menschenrechte …  (Artikel …)", "Durch die Erklärung der Menschenrechte gilt das Gesetz für alle gleich (Artikel 6), man darf frei reden und drucken (Artikel 11), das Eigentum ist geschützt (Artikel 17), und die Macht geht vom Volk aus (Artikel 3)."], ["Heute:", "Noch heute steht im Grundgesetz …", "Noch heute steht im Grundgesetz: Alle Menschen sind vor dem Gesetz gleich (Artikel 3), jeder darf seine Meinung frei äußern (Artikel 5), die Staatsgewalt geht vom Volke aus (Artikel 20)."]];
  B.forEach(([a, h, s]) => { c.push(wline(a, s, loesung, h)); if (!loesung) c.push(emptyLine()); });
  c.push(p(t("Rückmeldung der Lehrkraft (nur bei freiwilliger Abgabe)", { bold: true, size: 20, color: NAVY }), { before: 140, after: 50 }));
  [["Stufe 1", "Ich kann erklären, was sich durch den Verzicht auf die Vorrechte und die Menschenrechte für den Dritten Stand verbessert hat."], ["Stufe 2", "Ich kann Artikel von 1789 in heutige Begriffe übersetzen (zum Beispiel Eigentumsrecht)."], ["Stufe 3", "Ich kann an einem Grundgesetz-Artikel zeigen, dass die Erklärung bis heute nachwirkt.*"]].forEach(([a, b]) => c.push(p([t("□  ", { size: 24 }), t(a + "   ", { bold: true, size: 18, color: NAVY }), t(b, { size: 18 })], { after: 50 })));
  c.push(p([t("□ Ich gebe das Blatt freiwillig ab.", { size: 22, bold: true }), t("          Rückmeldung: ________________________________", { size: 20, color: MUTED })], { before: 100, after: 0 }));
  return c;
}

(async () => {
  const sec = (children) => ({ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 400, left: 700, right: 700 } } }, footers: { default: footer }, children });
  const st = { default: { document: { run: { font: "Calibri", size: 21 } } } };
  fs.writeFileSync("../Arbeitsblatt_Menschenrechte_Quelle_8D.docx", await Packer.toBuffer(new Document({ styles: st, sections: [sec([...seite1(false), ...seite2(false)])] })));
  fs.writeFileSync("../Loesung_Arbeitsblatt_Menschenrechte_8D.docx", await Packer.toBuffer(new Document({ styles: st, sections: [sec([p([t("LÖSUNG FÜR DIE LEHRKRAFT · GESCHICHTE 8D · STUNDE 4", { size: 17, bold: true, color: RED }), t("   Musterantworten (andere richtige Formulierungen zählen)", { size: 17, color: MUTED })], { after: 40 }), ...seite1(true), ...seite2(true)])] })));
  console.log("ok");
})();
