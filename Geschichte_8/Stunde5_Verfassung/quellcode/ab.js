const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, ImageRun, PageBreak, TabStopType, LeaderType, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", L1 = "F4F7FD", L2 = "EAF0FA", ACC = "7A1F3D", ACCL = "F7ECF0", GREEN = "1B5E3A", W = 10500;
const loesung = process.argv[2] === "loesung";
const t = (x, o = {}) => new TextRun({ text: x, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (r, o = {}) => new Paragraph({ children: Array.isArray(r) ? r : [r], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, keepNext: o.keepNext, indent: o.indent, tabStops: o.tabs });
const sd = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const bx = (c, s) => ({ top: sd(c, s), bottom: sd(c, s), left: sd(c, s), right: sd(c, s) });
const NB = { top: nb, bottom: nb, left: nb, right: nb };
const cell = (ch, w, o = {}) => new TableCell({ children: Array.isArray(ch) ? ch : [ch], width: { size: w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.b || bx("C5CEE6", 6), margins: { top: 60, bottom: 60, left: 120, right: 120 }, verticalAlign: o.v || VerticalAlign.TOP });
const table = (ws, rows) => new Table({ width: { size: ws.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: ws, rows: rows.map((r) => new TableRow({ children: r, cantSplit: true })) });
const band = () => table([W], [[cell(p(t("", { size: 2 })), W, { fill: ACC, b: NB })]]);
const kopf = (sub, pt) => [band(), p([t("GESCHICHTE 8C · FRANZÖSISCHE REVOLUTION   ", { size: 16, bold: true, color: MUTED }), t(loesung ? "Lösung (Lehrkraft)" : "Name: ______________________________", { size: 18, color: MUTED }), t(pt ? "\t" + (loesung ? "Punkte: ___ / 15" : "Punkte: ___ / 15") : "", { size: 18, color: MUTED })], { before: 50, after: 10, tabs: [{ type: TabStopType.RIGHT, position: W }] }), p([t("Die Verfassung von 1791", { font: "Cambria", size: 34, bold: true, color: ACC }), t("   " + sub, { size: 20, color: MUTED })], { after: 70 })];
const lineT = [{ type: TabStopType.RIGHT, position: W, leader: LeaderType.UNDERSCORE }];
const wl = (n = 1) => Array.from({ length: n }, () => p(t("\t"), { tabs: lineT, after: 110, before: 55 }));
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8C · Stunde 5 · Seite ", PageNumber.CURRENT] })] })] });
const IMG = new ImageRun({ type: "png", data: fs.readFileSync("../material/schaubild_verfassung.png"), transformation: { width: 660, height: 330 } });
const d = [];
// Seite 1
d.push(...kopf("Schaubild M2 und Quelle Q1 (nach dem Schulbuch S. 114/115, vereinfacht)", false));
d.push(p(IMG, { align: AlignmentType.CENTER, after: 80 }));
d.push(table([6000, 4500], [[
  cell([p(t("Q1  Aus der Verfassung von 1791 (gekürzt)", { bold: true, color: ACC, size: 22 }), { after: 40 }), p([t("Artikel 1: ", { bold: true, size: 20 }), t("Das Königtum geht vom Vater auf den Sohn über. Frauen sind ausgeschlossen.", { size: 20 })], { after: 30 }), p([t("Artikel 2: ", { bold: true, size: 20 }), t("Die Person des Königs ist unverletzlich und heilig. Sein einziger Titel ist „König der Franzosen“.", { size: 20 })], { after: 30 }), p([t("Artikel 3: ", { bold: true, size: 20 }), t("Es gibt in Frankreich keine höhere Autorität als das Gesetz. Der König regiert nur durch das Gesetz, und nur im Namen des Gesetzes kann er Gehorsam verlangen.", { size: 20 })], { after: 0 })], 6000, { fill: ACCL, b: bx(ACC, 10) }),
  cell([p(t("Wortspeicher", { bold: true, color: NAVY, size: 22 }), { after: 40 }), p([t("konstitutionelle Monarchie: ", { bold: true, size: 19 }), t("Die Macht des Königs wird durch eine Verfassung begrenzt.", { size: 19 })], { after: 25 }), p([t("Gewaltenteilung: ", { bold: true, size: 19 }), t("Gesetzgebung, Regierung und Rechtsprechung liegen bei verschiedenen Stellen.", { size: 19 })], { after: 25 }), p([t("Aktivbürger / Passivbürger: ", { bold: true, size: 19 }), t("mit Wahlrecht / ohne Wahlrecht.", { size: 19 })], { after: 0 })], 4500, { fill: L1, b: bx("B8C2DB", 8) })]]));
d.push(p(t("Tipp: So liest du ein Verfassungsschaubild (Schulbuch S. 86): 1 Thema (Staat, Jahr, wer hat beschlossen?) · 2 Beschreibung (Wer wählt? Welche Ämter? Wer arbeitet mit wem?) · 3 Beurteilung (Wer hat wie viel Macht? Wer ist ausgeschlossen?)", { size: 19, italics: true, color: MUTED }), { before: 80, after: 0 }));
// Seite 2
d.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }));
d.push(...kopf("Aufgaben (15 Punkte, Abgabe freiwillig)", true));
const aufg = (nr, titel, pt, frage, hilfe, lines, sol) => {
  const c = [p([t(nr + "  ", { font: "Cambria", bold: true, size: 26, color: ACC }), t(titel, { bold: true, size: 22, color: NAVY }), t("   (" + pt + " Punkte)", { size: 19, color: MUTED })], { after: 30 }), p(t(frage, { size: 20 }), { after: 30 })];
  if (hilfe) c.push(p([t("Hilfe: ", { bold: true, size: 18, color: ACC }), t(hilfe, { size: 18, italics: true, color: MUTED })], { after: 20 }));
  if (loesung) c.push(p([t("Erwartung: ", { bold: true, size: 19, color: GREEN }), t(sol, { size: 19, color: GREEN, italics: true })], { after: 0 }));
  else c.push(...wl(lines));
  return table([W], [[cell(c, W, { fill: "FFFFFF", b: bx("C5CEE6", 6) })]]);
};
const gap = () => p(t("", { size: 6 }), { after: 30 });
d.push(aufg("1", "Thema", 3, "Für welchen Staat gilt die Verfassung, wer hat sie wann beschlossen, und welche Regierungsform entsteht? Erkläre die Regierungsform in einem Satz.", "Die Verfassung gilt für …, die … beschloss sie im … 1791. Es entsteht eine …", 2, "Frankreich; Nationalversammlung; September 1791. Konstitutionelle Monarchie: Der König bleibt, aber seine Macht wird durch die Verfassung (Gesetze) begrenzt. 3 P: Staat, Beschluss und Jahr, Erklärung. 2 P: eine Angabe fehlt. 1 P: nur Staat oder Begriff."), gap());
d.push(aufg("2", "Wahlrecht", 3, "Wer durfte wählen (Aktivbürger), wer nicht (Passivbürger)? Nenne zwei Merkmale der Aktivbürger und zwei Gruppen von Passivbürgern.", "Aktivbürger sind … Passivbürger sind … Sie haben Rechte, aber kein …", 3, "Aktivbürger: Männer, über 25 Jahre, zahlen Steuern (mindestens drei Arbeitstage), ca. 4,3 Mio. Passivbürger: alle Frauen und Männer ohne oder mit geringen Steuern, ca. 22 Mio. 3 P: zwei Merkmale und zwei Gruppen. 2 P: drei Angaben. 1 P: eine Seite."), gap());
d.push(aufg("3", "Gewalten", 3, "Nenne die drei Gewalten und ordne König, Nationalversammlung und Gerichte zu. Erkläre, welche Pfeile zeigen, dass König und Nationalversammlung sich gegenseitig kontrollieren.", "Vollziehende Gewalt: … Gesetzgebende Gewalt: … Der Pfeil … zeigt, dass …", 3, "Vollziehende (König, Regierung), gesetzgebende (Nationalversammlung), richterliche Gewalt (Gerichte). Pfeile: König legt Einspruch gegen Gesetze ein (aufschiebend), Nationalversammlung kontrolliert den König (Regierung). 3 P: Zuordnung und beide Pfeile. 2 P: Zuordnung oder ein Pfeil. 1 P: ein Teil."), gap());
d.push(aufg("4", "Anteil rechnen", 3, "Rechne: Wie viele von 100 Menschen in Frankreich waren Aktivbürger (4,3 Mio. von zusammen 26,3 Mio.)? Reicht das für „das Volk bestimmt mit“?", "4,3 von 26,3, das sind etwa … von 100. Das ist … weil …", 3, "Etwa 16 von 100 (ca. 16 %). Urteil offen mit Begründung, zum Beispiel: Das reicht nicht, weil über 80 von 100 (alle Frauen und Arme) nicht wählen dürfen. 3 P: richtig gerechnet, Urteil begründet. 2 P: richtig gerechnet, Urteil ohne Beleg. 1 P: Ansatz."), gap());
d.push(aufg("5", "Urteil: Wer hatte mehr Macht?", 3, "Beurteile: Wer hatte in der Verfassung die größere Macht, der König oder die Nationalversammlung? Belege mit dem Schaubild oder mit Q1. Sage dann, ob du das gerecht findest, und nenne deinen Maßstab.", "Im Schaubild sieht man, dass …, deshalb hatte … mehr Macht. Ich finde das (nicht) gerecht, weil …", 4, "Mehrere gute Antworten. Beleg zum Beispiel: Der König ernennt die Regierung und hat ein Einspruchsrecht, aber die Nationalversammlung beschließt die Gesetze und kontrolliert ihn; Art. 3: Der König regiert nur durch das Gesetz. Maßstab zum Beispiel: Gerechtigkeit, weil die meisten Menschen nicht wählen durften. 3 P: Sachurteil mit Beleg und Werturteil mit Maßstab. 2 P: Sachurteil mit Beleg. 1 P: nur Behauptung."));
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 600, left: 700, right: 700 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync(loesung ? "../Loesung_Arbeitsblatt_Verfassung_8C.docx" : "../Arbeitsblatt_Verfassung_Schaubild_8C.docx", await Packer.toBuffer(doc)); console.log("ok");
})();
