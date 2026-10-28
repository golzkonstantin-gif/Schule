const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line } });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function box(children, fill, border) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [new TableCell({ children, width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, borders: border || all(noB), margins: { top: 90, bottom: 90, left: 160, right: 160 }, verticalAlign: VerticalAlign.TOP })] })] });
}
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8 · Französische Revolution · Stunde 2 · Seite ", PageNumber.CURRENT] })] })] });

const ABS = [
  "Am 5. Mai 1789 eröffnet der König die Generalstände in Versailles. Gleich zu Beginn gibt es Streit: Der Dritte Stand (Bauern und Bürger) fordert, dass nach Köpfen abgestimmt wird. Dann hätte er etwa so viel Gewicht wie Klerus und Adel zusammen. Der König lehnt ab. Wochenlang geht nichts voran.",
  "Der Dritte Stand findet Unterstützung in einer Flugschrift, die im Winter 1788/89 erscheint. Ihr Verfasser, der Geistliche Sieyès, stellt darin drei Fragen und gibt drei kurze Antworten (siehe Quelle).",
  "Am 17. Juni 1789 erklären sich die Vertreter des Dritten Standes zur Nationalversammlung. Sie sagen: Wir vertreten die Nation, also fast alle Menschen in Frankreich. Einige Geistliche und Adlige schließen sich an. Der König lässt den Sitzungssaal sperren. Daraufhin treffen sich die Abgeordneten am 20. Juni in einer Ballhalle in Versailles. Sie schwören, nicht auseinanderzugehen, bis Frankreich eine Verfassung hat. Damit stellt der Dritte Stand die Macht des Königs infrage.",
  "Der König gibt nicht nach. Anfang Juli zieht er Soldaten um Paris und Versailles zusammen und entlässt seinen beliebten Finanzminister. In Paris wächst die Angst.",
];
const WOERTER = [["Flugschrift", "kurze Schrift, die eine Meinung verbreitet"], ["Nation", "alle Menschen eines Landes"], ["Verfassung", "Gesetze, die festlegen, wer wie regiert"], ["Nationalversammlung", "Versammlung der Vertreter der Nation"]];

const c = [];
c.push(p(t("GESCHICHTE · KLASSE 8 · FRANZÖSISCHE REVOLUTION", { size: 16, bold: true, color: MUTED }), { after: 10 }));
c.push(p(t("Der Dritte Stand handelt", { font: "Cambria", size: 34, bold: true, color: NAVY }), { after: 80 }));
ABS.forEach((a, i) => {
  c.push(p(t(a, { font: "Cambria", size: 23 }), { after: 90, line: 290, align: AlignmentType.JUSTIFIED }));
  if (i === 1) {
    c.push(box([
      p(t("Q  Sieyès: „Was ist der Dritte Stand?“ (Winter 1788/89, gekürzt)", { bold: true, size: 20, color: NAVY }), { after: 40 }),
      p(t("Wir haben uns drei Fragen vorzulegen:", { font: "Cambria", size: 22, italics: true }), { after: 30 }),
      p([t("1. Was ist der Dritte Stand? ", { font: "Cambria", size: 22 }), t("ALLES.", { font: "Cambria", size: 22, bold: true, color: RED })], { after: 20 }),
      p([t("2. Was war er bisher in der politischen Ordnung? ", { font: "Cambria", size: 22 }), t("NICHTS.", { font: "Cambria", size: 22, bold: true, color: RED })], { after: 20 }),
      p([t("3. Was verlangt er? ", { font: "Cambria", size: 22 }), t("ETWAS ZU SEIN.", { font: "Cambria", size: 22, bold: true, color: RED })], { after: 0 }),
    ], LIGHT2, all(solid("B9C6E8", 8))));
    c.push(p(t(""), { after: 60 }));
  }
});
c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W / 2, W / 2], rows: [0, 2].map((i) => new TableRow({ cantSplit: true, children: [WOERTER[i], WOERTER[i + 1]].map((x) => new TableCell({ width: { size: W / 2, type: WidthType.DXA }, shading: { fill: LIGHT2, type: ShadingType.CLEAR, color: "auto" }, borders: all(noB), margins: { top: 30, bottom: 30, left: 120, right: 120 }, children: [p([t(x[0] + ": ", { bold: true, size: 18 }), t(x[1], { size: 18 })], { after: 0 })] })) })) }));
c.push(p(t(""), { after: 40 }));
c.push(box([p([t("Deine Aufgabe  ", { bold: true, color: RED, size: 20 }), t("1  Lies den Text und unterstreiche alle Daten.   2  Schreibe zu jedem Datum das Ereignis in den Zeitstrahl (Stichworte).   3  Schreibe ins Heft: Sieyès sagt, der Dritte Stand sei ALLES. Erkläre das mit einer Zahl.", { size: 20 })], { after: 0 })], "FFFFFF", all(solid("B9C6E8", 6))));

// Zeitstrahl
c.push(p(t("Zeitstrahl: Der Dritte Stand handelt", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 140, after: 40 }));
const DATEN = ["Winter 1788/89", "5. Mai 1789", "17. Juni 1789", "20. Juni 1789", "Anfang Juli 1789"];
const ZW = [2300, 500, 7406];
c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ZW, rows: DATEN.map((dat, i) => new TableRow({ cantSplit: true, height: { value: 640, rule: "atLeast" }, children: [
  new TableCell({ width: { size: ZW[0], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: { fill: i % 2 ? LIGHT : LIGHT2, type: ShadingType.CLEAR, color: "auto" }, borders: all(solid("B9C6E8", 6)), margins: { top: 30, bottom: 30, left: 140, right: 100 }, children: [p(t(dat, { size: 22, bold: true, color: RED }), { after: 0 })] }),
  new TableCell({ width: { size: ZW[1], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: all(noB), margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [p(t(i === 0 ? "▼" : i === DATEN.length - 1 ? "▲" : "│", { size: 22, color: NAVY, bold: true }), { after: 0, align: AlignmentType.CENTER })] }),
  new TableCell({ width: { size: ZW[2], type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM, borders: { top: noB, left: noB, right: noB, bottom: solid("8A96B5", 8) }, margins: { top: 20, bottom: 20, left: 120, right: 80 }, children: [p(t("", { size: 22 }), { after: 0 })] }),
] })) }));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 560, left: 850, right: 850 } } }, footers: { default: footer }, children: c }] });
  fs.writeFileSync("../Arbeitsblatt_Der_Dritte_Stand_handelt.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
