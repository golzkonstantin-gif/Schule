const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10700;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };

const ABS = [
  "Am 5. Mai 1789 eröffnet der König die Generalstände in Versailles. Der Dritte Stand (Bauern und Bürger) fordert, dass nach Köpfen abgestimmt wird. Der König lehnt ab. Wochenlang geht nichts voran.",
  "Schon vorher hat der Geistliche Sieyès eine Flugschrift geschrieben. Er fragt: Was ist der Dritte Stand? Seine Antwort: ALLES. Was war er bisher in der Politik? NICHTS. Was verlangt er? ETWAS ZU SEIN.",
  "Am 17. Juni 1789 erklärt sich der Dritte Stand zur Nationalversammlung. Er sagt: Wir vertreten fast alle Menschen in Frankreich. Der König lässt den Saal sperren. Am 20. Juni schwören die Abgeordneten in einer Ballhalle, nicht auseinanderzugehen, bis Frankreich eine Verfassung hat.",
  "Anfang Juli zieht der König Soldaten um Paris und Versailles zusammen. In Paris wächst die Angst.",
];
const WORTER = "Flugschrift: kurze Schrift, die eine Meinung verbreitet · Verfassung: Gesetze, die festlegen, wer wie regiert · Nationalversammlung: Versammlung der Vertreter aller Menschen eines Landes · Ballhalle: Halle für Ballspiele";
const DATEN = ["5. Mai 1789", "17. Juni 1789", "20. Juni 1789", "Anfang Juli 1789"];
const LOES = ["Eröffnung der Generalstände, Streit um die Abstimmung", "Dritter Stand erklärt sich zur Nationalversammlung", "Ballhausschwur: keine Auflösung ohne Verfassung", "Soldaten um Paris, Angst in Paris"];
const parts = (a, ul) => ul ? a.split(/((?:\d{1,2}\. (?:Mai|Juni)(?: 1789)?)|(?:Anfang Juli))/).map((x, k) => t(x, { font: "Cambria", size: 20, ...(k % 2 === 1 ? {} : {}) })) : [t(a, { font: "Cambria", size: 22 })];
const partsUl = (a) => a.split(/((?:\d{1,2}\. (?:Mai|Juni)(?: 1789)?)|(?:Anfang Juli))/).map((x, k) => new TextRun({ text: x, font: "Cambria", size: 22, underline: k % 2 === 1 ? {} : undefined }));

// Ein Blatt (halbe A4-Seite)
function blatt(solution) {
  const c = [];
  c.push(p([t(solution ? "LÖSUNG FÜR DIE LEHRKRAFT · " : "GESCHICHTE 8D · FRANZÖSISCHE REVOLUTION   ", { size: 15, bold: true, color: solution ? RED : MUTED }), t("Der Dritte Stand wehrt sich", { font: "Cambria", size: 26, bold: true, color: NAVY })], { after: 50 }));
  ABS.forEach((a) => c.push(p(solution ? partsUl(a) : parts(a, false), { after: 50, line: 280, align: AlignmentType.JUSTIFIED })));
  c.push(p(t(WORTER, { size: 17, color: MUTED }), { after: 60, before: 20 }));
  c.push(p([t("Aufgabe  ", { bold: true, color: RED, size: 18 }), t("1  Lies den Text und unterstreiche alle Daten.   2  Schreibe zu jedem Datum das Ereignis in den Zeitstrahl.", { size: 20 })], { after: 60 }));
  const ZW = [2000, 400, 8100];
  c.push(new Table({ width: { size: ZW.reduce((x, y) => x + y, 0), type: WidthType.DXA }, columnWidths: ZW, rows: DATEN.map((dat, i) => new TableRow({ cantSplit: true, height: { value: 780, rule: "atLeast" }, children: [
    new TableCell({ width: { size: ZW[0], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: { fill: i % 2 ? LIGHT : LIGHT2, type: ShadingType.CLEAR, color: "auto" }, borders: all(solid("B9C6E8")), margins: { top: 20, bottom: 20, left: 100, right: 60 }, children: [p(t(dat, { size: 19, bold: true, color: RED }), { after: 0 })] }),
    new TableCell({ width: { size: ZW[1], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: all(noB), margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [p(t(i === 0 ? "▼" : i === DATEN.length - 1 ? "▲" : "│", { size: 18, color: NAVY, bold: true }), { after: 0, align: AlignmentType.CENTER })] }),
    new TableCell({ width: { size: ZW[2], type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM, borders: { top: noB, left: noB, right: noB, bottom: solid("8A96B5", 8) }, margins: { top: 10, bottom: 10, left: 100, right: 60 }, children: [p(t(solution ? LOES[i] : "", { size: 19, bold: solution, color: solution ? "2E7D4F" : "000000" }), { after: 0 })] }),
  ] })) }));
  return c;
}
const cell = (children) => new TableCell({ children, width: { size: W, type: WidthType.DXA }, borders: all(dash), margins: { top: 90, bottom: 60, left: 190, right: 190 }, verticalAlign: VerticalAlign.TOP });
const page = (solution) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [0, 1].map(() => new TableRow({ cantSplit: true, height: { value: 7700, rule: HeightRule.EXACT }, children: [cell(blatt(solution))] })) });
const single = (solution) => blatt(solution);

(async () => {
  const sect = (children) => ({ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 400, left: 600, right: 600 } } }, children });
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [sect([page(false)])] });
  fs.writeFileSync("../Arbeitsblatt_Der_Dritte_Stand_wehrt_sich.docx", await Packer.toBuffer(doc));
  const doc2 = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [sect(single(true))] });
  fs.writeFileSync("../Loesung_Arbeitsblatt_Der_Dritte_Stand_wehrt_sich.docx", await Packer.toBuffer(doc2));
  console.log("ok");
})();
