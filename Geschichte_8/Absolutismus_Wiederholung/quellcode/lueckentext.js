const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Header, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 23, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
function box(children, fill, border) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [new TableCell({ children, width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, borders: border || all(none), margins: { top: 100, bottom: 100, left: 180, right: 180 }, verticalAlign: VerticalAlign.CENTER })] })] });
}

// Text mit Lücken: {n} = Lücke n. Lösung zeigt das Wort in Rot.
const GAPS = {
  1: "absoluter", 2: "Beamte", 3: "Intendanten", 4: "Steuerfreiheit", 5: "Versailles", 6: "Merkantilismus",
  7: "Zölle", 8: "Gottesgnadentum", 9: "Glaube", 10: "stehendes", 11: "400.000", 12: "Drittel",
};
const WORDS = ["Beamte", "Glaube", "Merkantilismus", "Gottesgnadentum", "Versailles", "Steuerfreiheit", "Zölle", "stehendes", "400.000", "Drittel", "absoluter", "Intendanten"];
const SECTIONS = [
  ["Einleitung", "Ludwig XIV. regierte Frankreich als {1} Monarch. Er konzentrierte möglichst viel Macht bei sich und stützte sie auf fünf Säulen."],
  ["Säule 1 · Verwaltung und Justiz", "Der König setzte {2} ein, die in seinem Namen handelten. In den Provinzen vertraten ihn die {3}. Als oberster Richter konnte er Personen mit „lettres de cachet“ ohne Gerichtsverfahren verhaften lassen."],
  ["Säule 2 · Adel und Hof", "Dem Adel blieben Privilegien wie die {4} und das Jagdrecht. Viele Adelige band Ludwig an seinen Hof in {5} und hielt sie dort mit Festen und Geschenken bei Laune. Sie sollten keinen Widerstand leisten."],
  ["Säule 3 · Wirtschaft", "Finanzminister Colbert führte den {6} ein: Der Staat lenkte die Wirtschaft, damit mehr fertige Waren ins Ausland verkauft als von dort eingeführt wurden. Hohe {7} machten ausländische Produkte teuer."],
  ["Säule 4 · Religion", "Ludwig begründete seine Herrschaft mit dem {8}: Er sei von Gott zum Herrschen ausgewählt. 1685 beendete er die religiöse Toleranz. Wie nur einen König sollte es auch nur einen {9} geben."],
  ["Säule 5 · Heer", "Statt Söldner baute Ludwig ein {10} Heer auf. Es wuchs auf bis zu {11} Mann. In Friedenszeiten kostete es ein {12} des Staatshaushalts, was zu ständigen Steuererhöhungen führte."],
];

function runsFor(text, solution) {
  const out = [];
  text.split(/(\{\d+\})/).filter(Boolean).forEach((tok) => {
    const m = tok.match(/^\{(\d+)\}$/);
    if (!m) return out.push(t(tok, { size: 24 }));
    const n = m[1];
    out.push(t(`(${n}) `, { size: 18, bold: true, color: MUTED }));
    out.push(solution ? t(GAPS[n], { bold: true, color: RED, size: 24 }) : t("__________________", { size: 24, color: "7A869E" }));
  });
  return out;
}

function build(solution) {
  const c = [];
  c.push(p(t("GESCHICHTE · KLASSE 8 · ABSOLUTISMUS", { size: 18, bold: true, color: MUTED }), { after: 20 }));
  c.push(p(t(solution ? "Lösung: Die fünf Säulen der Herrschaft Ludwigs XIV." : "Die fünf Säulen der Herrschaft Ludwigs XIV.", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 60 }));
  if (!solution) {
    c.push(p([t("Aufgabe 1  ", { bold: true, color: RED, size: 22 }), t("Lies den Text einmal ganz. Ergänze dann die Lücken mit dem Wortspeicher und streiche jedes benutzte Wort durch.", { size: 22 })], { after: 80 }));
    c.push(box([p(t("Wortspeicher", { bold: true, color: NAVY, size: 20 }), { after: 40 }), p(t(WORDS.join("   ·   "), { font: "Cambria", size: 23, color: NAVY }), { after: 0 })], LIGHT2));
    c.push(p(t(""), { after: 60 }));
  }
  SECTIONS.forEach(([h, text]) => {
    c.push(p(t(h, { font: "Cambria", size: 25, bold: true, color: NAVY }), { before: 100, after: 40, keepNext: true }));
    c.push(p(runsFor(text, solution), { after: 60, line: 330 }));
  });
  if (!solution) {
    c.push(box([p([t("★ Schon fertig?  ", { bold: true, color: "C98A1E", size: 22 }), t("Welche Säule ist für Ludwig XIV. die wichtigste? Begründe in einem Satz. Schreibe ins Heft.", { size: 22 })], { after: 0 })], "FFFFFF", all(solid("E8A33D", 8))));
  } else {
    c.push(p([t("Zusatz ★: ", { bold: true, color: MUTED, size: 21 }), t("Es gibt keine einzig richtige Antwort. Erwartet wird eine Begründung, die zur gewählten Säule passt (z. B. Heer: Es setzt die Macht notfalls mit Gewalt durch; Wirtschaft: ohne Geld kann der König Heer und Beamte nicht bezahlen).", { size: 21, color: MUTED })], { before: 100, after: 40 }));
  }
  return new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 23 } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [(solution ? "Lösung · " : "") + "Geschichte 8 · Seite ", PageNumber.CURRENT] })] })] }) },
      children: c,
    }],
  });
}
(async () => {
  fs.writeFileSync("../Arbeitsblatt_Lueckentext_Fuenf_Saeulen.docx", await Packer.toBuffer(build(false)));
  fs.writeFileSync("../Loesung_Lueckentext_Fuenf_Saeulen.docx", await Packer.toBuffer(build(true)));
  console.log("ok");
})();
