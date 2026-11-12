// Lösung zum Arbeitsblatt „Vor dem Gesetz“ (Aufgaben 2–4) und zum Blatt „Nervig oder kafkaesk?“ (Aufgabe 1) – für die Lehrkraft
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Header, Footer, TabStopType,
} = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", RED = "D9534F", GREEN = "2E9E6B";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells) => new TableRow({ children: cells, cantSplit: true });
const kopf = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w, fill: NAVY });
const h2 = (text, o = {}) => p(t(text, { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: o.before ?? 160, after: 60, pb: o.pb, keepNext: true });
const aufg = (text) => p(t(text, { bold: true, color: NAVY, size: 22 }), { before: 140, after: 50, keepNext: true });
const loes = (text, o = {}) => t(text, Object.assign({ color: GREEN }, o));

// Aufgabe 2
const a2 = [
  p(loes(D.TATT.satz + " " + D.TATT.hypothese), { line: 290 }),
];

// Aufgabe 3
const gW = [450, 3500, W - 3950];
const a3 = table(gW, [
  row([kopf("", gW[0]), kopf("Der Mann vom Lande …", gW[1]), kopf("Der Türhüter … (Zitat, vgl. Z. …)", gW[2])]),
  ...D.GESPRAECH.map(([mann, th, anker], i) => row([
    cell(p(t(String(i + 1), { bold: true, color: NAVY }), { after: 0, align: AlignmentType.CENTER }), { w: gW[0], fill: LIGHT }),
    cell([p(t(mann, { size: 19 }), { after: 10 }), p(t("vgl. " + D.MANN_Z[i], { size: 17, color: MUTED }), { after: 0 })], { w: gW[1], fill: LIGHT }),
    cell(p([loes(th, { size: 19 }), loes("  (vgl. " + D.zeilen(anker) + ")", { size: 17 })], { after: 0 }), { w: gW[2] }),
  ])),
]);
const a3fragen = D.AUFGABE3.fragen.flatMap((q, i) => [
  p(t(q, { bold: true, size: 19, color: NAVY }), { before: 80, after: 20, keepNext: true }),
  p(loes(i === 0
    ? "Der Türhüter: Er entscheidet („jetzt aber nicht“), droht („Ich bin mächtig“), verhört und schließt am Ende den Eingang. Der Mann bittet, fragt, wartet und besticht – er nimmt die untere Rolle an. Bild dafür: der „Größenunterschied … zuungunsten des Mannes“ (vgl. " + D.zeile("Größenunterschied hat sich sehr") + ")."
    : "Der Türhüter spricht in wörtlicher Rede. Der Mann wird bis kurz vor seinem Tod nur indirekt wiedergegeben („fragt dann, ob er also später werde eintreten dürfen“, vgl. " + D.zeile("fragt dann, ob er also") + ") – er hat keine eigene Stimme. → Das Gespräch ist komplementär (asymmetrisch); die Hypothese stimmt.", { size: 19 }), { line: 280 }),
]);

// Aufgabe 4
const sW = [2600, 1100, W - 3700];
const a4 = table(sW, [
  row([kopf("Bildebene", sW[0]), kopf("Zeile", sW[1]), kopf("Sachebene: Wofür steht es? Begründe mit dem Text.", sW[2])]),
  ...D.SYMBOLE.map(([b, anker, sach]) => row([
    cell(p(t(b, { bold: true, color: RED, size: 20 }), { after: 0 }), { w: sW[0], fill: LIGHT }),
    cell(p(loes("vgl. " + D.zeile(anker), { size: 19 }), { after: 0 }), { w: sW[1] }),
    cell(p(loes(sach, { size: 19 }), { after: 0 }), { w: sW[2] }),
  ])),
]);
const gesetzOffen = p(t("Beim Gesetz jede Deutung gelten lassen, die am Text begründet ist (Recht, Wahrheit, Sinn, Gott, Vater). Wortbelege: Tor „offensteht wie immer“, Türhüter „Ich bin mächtig“, „nur der unterste“, Gesetz „jedem und immer zugänglich“.", { size: 17, italics: true, color: MUTED }), { before: 60 });

// Aufgabe 1: Nervig oder kafkaesk?
const nW = [520, W - 520];
const a1 = table(nW, D.SZENARIEN.map(([stich, text, nk, baustein, grund], i) => row([
  cell(p(t(String(i + 1), { bold: true, color: NAVY, size: 22 }), { after: 0 }), { w: nW[0], fill: LIGHT }),
  cell([
    p(t(text, { size: 18, color: MUTED }), { after: 40, line: 260 }),
    p([t(nk === "K" ? "☐ nervig   ☒ kafkaesk" : "☒ nervig   ☐ kafkaesk", { bold: true, size: 20, color: nk === "K" ? RED : GREEN })], { after: 30 }),
    p([t("Begründung: ", { bold: true, size: 19 }), loes((baustein ? baustein + " – " : "") + grund, { size: 19 })], { after: 0 }),
  ], { w: nW[1] }),
])));

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Vor dem Gesetz“", { size: 17, color: MUTED }), t("\tLÖSUNG (für die Lehrkraft)", { size: 17, bold: true, color: GREEN })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t("Zeilenangaben nach dem Arbeitsblatt; ± 1 Zeile gilt als richtig.", { size: 16, color: MUTED })] })] });

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 760, bottom: 620, left: 850, right: 850, header: 340, footer: 340 } } }, headers: { default: header }, footers: { default: footer }, children: [
    p(t("Lösung: Arbeitsblatt „Vor dem Gesetz“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 60 }),
    aufg(D.TATT.titel), ...a2,
    aufg(D.AUFGABE3.titel), a3, ...a3fragen,
    aufg(D.AUFGABE4.titel), a4, gesetzOffen,
    h2("Lösung: Aufgabe 1 – Nervig oder kafkaesk?", { pb: true, before: 0 }),
    a1,
  ] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Loesung_Vor_dem_Gesetz.docx", b));
