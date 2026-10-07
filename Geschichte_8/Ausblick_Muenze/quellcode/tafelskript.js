const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, VerticalMergeType, Header, Footer, TabStopType, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", BOARD = "F3F6F1", RED = "D9534F", DARK = "3D4A3F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext, indent: o.indent });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
function cell(children, o = {}) {
  return new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("B9C6E8", 6)), verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: 110, right: 110 }, rowSpan: o.rowSpan, verticalMerge: o.vMerge });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: "atLeast" } : undefined, cantSplit: true });
const th = (txt, w, fill = DARK, size = 19) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { after: 0, align: AlignmentType.CENTER }), { w, fill, borders: all(solid(DARK, 8)) });
const kicker = (s) => p(t(s.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h1 = (s) => p(t(s, { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 60 });
const h3 = (s) => p(t(s, { font: "Cambria", size: 24, bold: true, color: NAVY }), { before: 160, after: 60, keepNext: true });
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8 · Ausblick Französische Revolution", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Seite ", PageNumber.CURRENT] })] })] });

const BW = 1700, CW = W - BW, TW = CW - 340;
function step(klasse, parts) {
  const ch = [];
  if (parts.say) ch.push(p([t("Ansage: ", { bold: true, size: 20, color: RED }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((x) => ch.push(p(S(x), { after: 50 })));
  if (parts.extra) ch.push(...parts.extra);
  if (parts.sol) ch.push(p([t("Erwartung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  return table([BW, CW], [row([cell([p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }), p(t(klasse, { bold: true, color: NAVY, size: 20 }), { after: 0 })], { w: BW, fill: LIGHT, valign: VerticalAlign.TOP }), cell(ch, { w: CW, valign: VerticalAlign.TOP })])]);
}
const gapP = () => p(t(""), { after: 30 });
const phase = (nr, title, min) => p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: RED }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 180, after: 90, keepNext: true });

const TEXT = [
  "Eine Münze liegt in Maries Hand, die letzte aus diesem Herbst. Der Regen hat das Korn verfaulen lassen, das Brot kostet doppelt so viel wie im Frühjahr. Aber die Münze gehört nicht Marie. Drei Hände strecken sich nach ihr aus.",
  "Die erste gehört dem Verwalter des Grundherrn. Er verlangt die Pacht und fragt nicht, wie die Ernte war. Den Grundherrn selbst hat in Maries Dorf seit Jahren niemand gesehen. Er lebt am Hof in Versailles. Dort brennen tausend Kerzen, man spielt Theater, eine Gräfin gähnt. „Die Leute auf dem Land“, sagt sie, „immer haben sie etwas.“",
  "Die zweite Hand gehört dem Pfarrer. Er verlangt den Kirchenzehnt, und es fällt ihm schwer. Er soll im Winter für die Armen sorgen und nimmt doch, was ihnen fehlt. Er geht mit schlechtem Gewissen.",
  "Die dritte Hand gehört dem Steuereintreiber des Königs. Sein Geld wandert nach Paris in die Staatskasse. Aber die Kasse ist leer. Heer, Hof und Beamte verschlingen mehr, als hereinkommt. Reicht es nicht, steigen die Steuern, und wieder trifft es dieselben.",
  "Abends im Wirtshaus sitzt ein Anwalt aus Paris. Er ist gebildet, wohlhabend und zahlt Steuern. Er kennt die leere Kasse und wüsste, wie man es besser machen könnte. Aber niemand fragt ihn. Einer erzählt, am Hof habe man wieder drei Tage gefeiert. Früher hätte jemand gesagt: „So ist es eben.“ Heute sagt niemand etwas. Der Wirt legt langsam sein Tuch aus der Hand.",
];

const d = [];
d.push(kicker("Tafelskript für die Lehrkraft · Ausblick nach dem Test · 10 Minuten"), h1("Die Münze: Wohin wandert das Geld?"));
d.push(p(t("Ziel: Die Klasse ahnt, dass die Stimmung in Frankreich kippt, und erkennt die Probleme des Landes. Aufhänger ist eine Münze, nach der drei Hände greifen. Alle Stände kommen vor. Es gibt keine Folie und kein Heft: Das Ergebnis entsteht gemeinsam an der Tafel. Der Text ist fiktiv und aus den Gruppenpuzzle-Texten gebaut.", { size: 21 }), { after: 80 }));

d.push(table([800, 6200, 3206], [
  row([th("Min.", 800), th("Phase", 6200), th("Arbeitsform", 3206)], 320),
  ...[["2", "1  Text hören: Wohin wandert die Münze?", "zuhören"], ["5", "2  Tafelbild: Weg der Münze, Probleme, Stimmung", "gemeinsam, mündlich"], ["3", "3  Gespräch: Kippt die Stimmung? Ausblick 1789", "gemeinsam, mündlich"]].map((r) => row(r.map((v, i) => cell(p(t(v, { size: 20, bold: i < 2 }), { after: 0 }), { w: [800, 6200, 3206][i] })), 320)),
]));

// Phase 1
d.push(phase("1", "Text hören", 2));
d.push(step("hört zu", {
  say: "Stifte liegen. Ihr hört nur zu. Achtet darauf: Wohin wandert die Münze?",
  do: ["Langsam vorlesen. Nach „Drei Hände strecken sich nach ihr aus“ und nach dem Satz „Heute sagt niemand etwas“ jeweils kurz innehalten."],
  extra: [
    table([TW], [
      row([cell(p(t("TEXT ZUM VORLESEN", { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: all(solid(NAVY, 8)), m: 40 })]),
      row([cell(TEXT.map((x, i) => p(t(x, { font: "Cambria", size: 22 }), { after: i === TEXT.length - 1 ? 0 : 80 })), { w: TW, fill: LIGHT, borders: all(solid(NAVY, 8)), valign: VerticalAlign.TOP, m: 100 })]),
    ]),
    gapP(),
  ],
}));

// Phase 2
d.push(phase("2", "Tafelbild entwickeln", 5));
const widths = [1300, 1400, 1700, 2000, 1700];
const bc = (txt, w, o = {}) => cell(p(t(txt, { size: o.size || 19, bold: o.bold, font: "Cambria" }), { after: 0, align: o.align }), { w, fill: BOARD, borders: all(solid(DARK, 8)), rowSpan: o.rowSpan });
const cont = (w) => cell(p(t(""), { after: 0 }), { w, fill: BOARD, borders: all(solid(DARK, 8)), vMerge: VerticalMergeType.CONTINUE });
const tafelbild = table(widths, [
  row([th("MARIE", widths[0]), th("Pfeil", widths[1]), th("landet bei", widths[2]), th("Problem", widths[3]), th("Stimmung", widths[4])], 360),
  row([bc("MÜNZE", widths[0], { bold: true, size: 22, align: AlignmentType.CENTER, rowSpan: 3 }), bc("Pacht → Grundherr (Adel)", widths[1], { bold: true }), bc("Hof von Versailles: Feste, Theater", widths[2]), bc("Adel kassiert, ist weit weg und kennt die Not nicht", widths[3]), bc("Gräfin: gleichgültig", widths[4])], 640),
  row([bc("Zehnt → Kirche (Klerus)", widths[1], { bold: true }), bc("Kirche, Armenfürsorge", widths[2]), bc("Pfarrer im Zwiespalt: Armen helfen und Zehnt einziehen", widths[3]), bc("Pfarrer: schlechtes Gewissen", widths[4])], 640),
  row([bc("Steuer → König (Staat)", widths[1], { bold: true }), bc("Staatskasse → Heer, Beamte, Hof", widths[2]), bc("Kasse leer, Steuern steigen, immer trifft es dieselben", widths[3]), bc("Dorf: verbittert, still", widths[4])], 640),
  row([bc("Bürger: Anwalt", widths[0], { bold: true }), bc("zahlt Steuern", widths[1], { bold: true }), bc("Staatskasse", widths[2]), bc("gebildet und wohlhabend, aber ohne Mitsprache", widths[3]), bc("Anwalt: ungeduldig", widths[4])], 640),
]);
d.push(step("macht mündlich mit", {
  say: "Meldet euch. Ihr schreibt noch nichts auf. Wohin wandert die Münze?",
  do: ["Während die Klasse antwortet, entsteht das Tafelbild. Zuerst nur Marie mit der Münze und die drei Pfeile mit Namen („Pacht“, „Zehnt“, „Steuer“). Dann Spalte für Spalte ergänzen lassen: Wohin geht das Geld? Welches Problem? Welche Stimmung?"],
  extra: [tafelbild, gapP()],
  sol: "Die Klasse nennt zuerst „zum Hof von Versailles“, „zur Kirche“, „zum Staat“. Danach gezielt nachfragen: Was bleibt für Marie? (nichts) · Wer trägt die Last? (der Dritte Stand (Bauern und Bürger)) · Wer hat die Vorrechte? (Adel und Klerus).",
}));

// Phase 3
d.push(phase("3", "Gespräch und Ausblick", 3));
d.push(step("macht mündlich mit", {
  say: "Meldet euch. Ihr schreibt noch nichts auf.",
  do: [
    "1  Was bleibt für Marie?",
    "2  Welche Stelle im Text zeigt, dass die Stimmung kippen könnte? Begründe.",
    "3  Welches Problem ist größer: die leere Staatskasse oder die ungleiche Verteilung der Lasten? Begründe.",
    "Zum Schluss Ihr Hinweis: „1789 kippt es. Davon handelt die nächste Stunde: die Französische Revolution.“",
  ],
  sol: "1 nichts · 2 „Früher hätte jemand gesagt: ‚So ist es eben.‘ Heute sagt niemand etwas.“ Das Schweigen zeigt, dass die Menschen nicht mehr hinnehmen, sondern über Veränderung nachdenken. · 3 beide Antworten lassen sich begründen (leere Kasse: Der Staat braucht ständig mehr Geld, die Steuern steigen; ungleiche Lasten: Wer zahlen muss, hat keine Mitsprache).",
}));

d.push(h3("Zum Prüfen"));
[
  "Die Zuspitzung „drei Hände nach einer Münze“ ist erfunden. Historisch zahlten die Bauern Pacht, Zehnt und Steuern nebeneinander, teils in Geld, teils in Naturalien und Diensten. Das lässt sich im Gespräch richtigstellen.",
  "Dass der Grundherr nie im Dorf ist und der Anwalt im Wirtshaus sitzt, stammt nicht aus den Gruppenpuzzle-Texten, ist aber plausibel.",
  "Der Zeitabstand: Ludwig XIV. starb 1715, die Revolution begann 1789. Besser „die Spannungen wachsen über Jahrzehnte“ als „rund 100 Jahre später“.",
  "Wenn Zeit bleibt: die Karikatur „Drei Stände“ (1789) noch einmal zeigen. Die Unterschrift bedeutet sinngemäß „Hoffentlich ist das Spiel bald vorbei“.",
].forEach((x) => d.push(p([S("• "), S(x)], { after: 40, indent: { left: 240 } })));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Ausblick_Muenze.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
