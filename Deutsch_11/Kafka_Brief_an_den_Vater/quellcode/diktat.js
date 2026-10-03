// Diktat „Kontext und Fachbegriffe“: Vorlesefassung (docx) + Whiteboard zur Selbstkontrolle (pptx → pdf)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber,
} = require("docx");
const pptxgen = require("pptxgenjs");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GREEN = "2E9E6B", GOLD = "E8A33D";

// Diktattext: Abschnitte für das Diktieren durch „ / “ getrennt
const TEIL1 = [
  "Franz Kafka wurde 1883 in Prag geboren.",
  "Sein Vater Hermann war ein Kaufmann, / der sich aus ärmlichen Verhältnissen hochgearbeitet hatte.",
  "Er erwartete von seinem Sohn, / dass er stark, fleißig und dankbar ist.",
  "Im November 1919 schrieb der sechsunddreißigjährige Kafka seinem Vater einen Brief, / der über hundert handschriftliche Seiten lang ist.",
  "Darin versucht er zu erklären, / warum er Angst vor ihm hat.",
  "Das Besondere daran ist, / dass er schreibt, / weil er mit seinem Vater nicht reden kann.",
  "Kafka gab den Brief seiner Mutter, / damit sie ihn weitergibt.",
  "Sie tat das aber nicht, / sodass der Vater ihn nie gelesen hat.",
  "Veröffentlicht wurde der Brief erst nach Kafkas Tod.",
];
const TEIL2 = [
  "Nach Paul Watzlawick ist zwischenmenschliche Kommunikation entweder symmetrisch oder komplementär.",
  "Symmetrisch ist sie, / wenn die Beziehung auf Gleichheit beruht / und beide Partner auf derselben Stufe stehen.",
  "Komplementär ist sie, / wenn die Beziehung auf Unterschiedlichkeit beruht: / Einer ist überlegen, / der andere ist unterlegen.",
  "Wird dieses Verhältnis so starr, / dass niemand die Rolle wechseln kann, / ist kein Gespräch auf Augenhöhe mehr möglich.",
  "Von einem Generationenkonflikt spricht man, / wenn Eltern und Kinder unterschiedliche Werte, Erfahrungen und Erwartungen haben / und sich deshalb nicht verstehen.",
];
const clean = (arr) => arr.map((s) => s.replace(/ \/ /g, " ")).join(" ");

// Stolperstellen (für Lehrkraft und Whiteboard)
const STOLPER = [
  ["das oder dass?", "„dass“ leitet einen Nebensatz ein und lässt sich nicht durch „dieses“ oder „welches“ ersetzen.", ["… erwartete von seinem Sohn, dass er stark … ist.", "Das Besondere daran ist, dass er schreibt …", "Sie tat das aber nicht …  (das = dieses)"]],
  ["Komma vor Nebensätzen", "Vor dass, weil, wenn, damit, sodass, warum und vor Relativsätzen (der, die, das) steht ein Komma.", ["… einen Brief, der über hundert … Seiten lang ist.", "… zu erklären, warum er Angst vor ihm hat.", "… der Mutter, damit sie ihn weitergibt."]],
  ["Groß oder klein?", "Nominalisierte Adjektive groß, Zahlwörter und Zahladjektive klein und zusammen.", ["groß: das Besondere", "klein: der sechsunddreißigjährige Kafka", "klein: über hundert Seiten  ·  einer … der andere"]],
  ["Fachwörter", "Diese Wörter müssen sitzen – sie kommen in jeder Analyse vor.", ["symmetrisch  ·  komplementär", "Generationenkonflikt  ·  Watzlawick", "handschriftlich  ·  veröffentlicht"]],
];

// ================= Vorlesefassung (docx) =================
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const W = 10206;
const cell = (children, o = {}) => new TableCell({ children, width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: 80, bottom: 80, left: 120, right: 120 } });
const box = (paras, fill = LIGHT2) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [cell(paras, { w: W, fill, borders: all(none) })] })] });

function dictSentences(arr, startNo) {
  return arr.map((s, i) => {
    const parts = s.split(" / ");
    const runs = [t(String(startNo + i) + "  ", { bold: true, color: MUTED, size: 22 })];
    parts.forEach((part, j) => {
      runs.push(t(part, { font: "Cambria", size: 26 }));
      if (j < parts.length - 1) runs.push(t("  ‖  ", { bold: true, color: RED, size: 28 }));
    });
    return p(runs, { after: 110, line: 320 });
  });
}

const vorlesen = [
  p(t("DIKTAT · VORLESEFASSUNG FÜR DIE LEHRKRAFT", { size: 18, bold: true, color: MUTED }), { after: 20 }),
  p(t("Kafka und sein Vater – Kontext und Fachbegriffe", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 120 }),
  box([
    p([t("So läuft das Diktat (ca. 12 Min. + 3 Min. Selbstkontrolle)", { bold: true, color: NAVY })], { after: 60 }),
    p(t("1.  Den ganzen Text einmal vorlesen – die Klasse hört nur zu.", { size: 21 }), { after: 30 }),
    p(t("2.  Satz für Satz diktieren, an den roten Strichen ‖ kurz pausieren. Satzzeichen nicht ansagen – die Kommas gehören zur Übung.", { size: 21 }), { after: 30 }),
    p(t("3.  Den ganzen Text noch einmal vorlesen; die Klasse liest still mit und korrigiert.", { size: 21 }), { after: 30 }),
    p(t("4.  Selbstkontrolle über das Whiteboard: Fehler farbig anstreichen, Anzahl unten notieren (Kommafehler extra zählen). Danach die Folie mit den Stolperstellen zeigen.", { size: 21 }), { after: 0 }),
  ]),
  p(t("Teil 1 · Kontext", { font: "Cambria", size: 26, bold: true, color: RED }), { before: 200, after: 100 }),
  ...dictSentences(TEIL1, 1),
  p(t("Teil 2 · Fachbegriffe", { font: "Cambria", size: 26, bold: true, color: RED }), { before: 160, after: 100 }),
  ...dictSentences(TEIL2, TEIL1.length + 1),
  p(t("Stolperstellen im Überblick", { font: "Cambria", size: 26, bold: true, color: NAVY }), { pb: true, after: 100 }),
  ...STOLPER.flatMap(([h, regel, bsp]) => [
    p([t(h + ":  ", { bold: true, color: RED, size: 22 }), t(regel, { size: 21 })], { after: 30 }),
    ...bsp.map((b) => p(t("•  " + b, { size: 21, italics: true, color: MUTED }), { after: 20 })),
    p(t(""), { after: 60 }),
  ]),
  box([
    p([t("Zur Information: ", { bold: true, color: NAVY, size: 20 }), t(`Teil 1 hat ${clean(TEIL1).split(/\s+/).length} Wörter, Teil 2 hat ${clean(TEIL2).split(/\s+/).length} Wörter. Das Diktat deckt zugleich die vorgemerkten Fachbegriffe ab, die in der Stunde gebraucht werden (komplementär, symmetrisch, Generationenkonflikt).`, { size: 20 })], { after: 0 }),
  ]),
];

const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Diktat · Kafka und sein Vater · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, footer: 400 } } }, footers: { default: footer }, children: vorlesen }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Diktat_Vorlesefassung.docx", b));

// ================= Whiteboard zur Selbstkontrolle (pptx) =================
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Diktat – Selbstkontrolle";
const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: "Calibri", fontSize: 20, color: NAVY, margin: 0, valign: "top" }, o));
const rect = (s, x, y, w, h, fill, r = 0.08) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, rectRadius: r });
function header(s, kick, title) {
  rect(s, 0, 0, 13.333, 1.0, NAVY, 0);
  txt(s, kick, { x: 0.5, y: 0.12, w: 9, h: 0.3, fontSize: 13, bold: true, color: "CADCFC", charSpacing: 2 });
  txt(s, title, { x: 0.5, y: 0.4, w: 12.3, h: 0.55, fontFace: "Cambria", fontSize: 26, bold: true, color: "FFFFFF" });
}
[["TEIL 1 · KONTEXT", "Diktat – vergleiche Wort für Wort", TEIL1], ["TEIL 2 · FACHBEGRIFFE", "Diktat – vergleiche Wort für Wort", TEIL2]].forEach(([k, ti, arr]) => {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "SELBSTKONTROLLE · " + k, ti);
  txt(s, clean(arr), { x: 0.6, y: 1.3, w: 12.1, h: 5.1, fontFace: "Cambria", fontSize: k.includes("1") ? 23 : 24, lineSpacingMultiple: 1.25 });
  rect(s, 0.6, 6.55, 12.1, 0.6, LIGHT2);
  txt(s, "Fehler farbig anstreichen · Anzahl notieren · Kommafehler extra zählen", { x: 0.85, y: 6.55, w: 11.6, h: 0.6, fontSize: 17, bold: true, valign: "middle" });
});
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "SELBSTKONTROLLE · STOLPERSTELLEN", "Darauf achten – auch in jeder Klausur");
  const cols = [RED, GOLD, GREEN, "3F7CC4"];
  STOLPER.forEach(([h, regel, bsp], i) => {
    const x = 0.5 + (i % 2) * 6.2, y = 1.25 + Math.floor(i / 2) * 2.95;
    rect(s, x, y, 6.0, 2.8, LIGHT);
    rect(s, x, y, 0.14, 2.8, cols[i], 0.02);
    txt(s, h, { x: x + 0.35, y: y + 0.12, w: 5.5, h: 0.45, fontFace: "Cambria", fontSize: 20, bold: true, color: cols[i] });
    txt(s, regel, { x: x + 0.35, y: y + 0.6, w: 5.45, h: 0.85, fontSize: 16 });
    txt(s, bsp.join("\n"), { x: x + 0.35, y: y + 1.45, w: 5.45, h: 1.3, fontSize: 16, italic: true, color: MUTED, paraSpaceAfter: 3 });
  });
}
pres.writeFile({ fileName: process.argv[2] || "Diktat_Whiteboard_Selbstkontrolle.pptx" });
