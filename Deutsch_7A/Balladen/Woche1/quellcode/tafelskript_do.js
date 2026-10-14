const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak, Header, Footer, TabStopType,
  PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", BOARD = "F3F6F1";
const C = { hv: "D9534F", mv: "2E9E6B", p2: "E8A33D", inf: "3F7CC4" };
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext, indent: o.indent });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h2 = (label, text, after = 80) => p([t(label + "  ", { font: "Cambria", size: 26, bold: true, color: C.hv }), t(text, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 200, after, keepNext: true });
const h3 = (text, col = NAVY) => p(t(text, { font: "Cambria", size: 23, bold: true, color: col }), { before: 120, after: 60, keepNext: true });
const gap = (n = 80) => p(t(""), { after: n });
const br = () => new Paragraph({ children: [new PageBreak()] });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };
const thin = allBorders(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || thin,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font, italics: o.it }), { after: 0, align: o.align }), { w, fill: o.fill });

// simple grid table: header + rows of strings/runs
function grid(widths, head, rows, o = {}) {
  return table(widths, [
    row(head.map((h, i) => hdr(h, widths[i], o.headFill ? o.headFill[i] : NAVY)), 420),
    ...rows.map((r) => row(r.map((v, i) => tc(v, widths[i], { font: o.fonts && o.fonts[i], align: o.align && o.align[i], bold: o.bold && o.bold[i], color: o.colors && o.colors[i] })), o.h || 460)),
  ]);
}

function box(children, fill = LIGHT2, border) {
  return table([W], [row([cell(children, { w: W, fill, borders: border || allBorders(none), m: 140, ml: 180 })])]);
}
// Merksatz box
function merk(nr, title, runs) {
  return box([
    p([t(`Merksatz ${nr} · `, { bold: true, color: C.hv, size: 20 }), t(title, { bold: true, color: NAVY, size: 22, font: "Cambria" })], { after: 60 }),
    ...runs.map((r) => p(r, { after: 40 })),
  ], LIGHT2);
}
// Tafel box
function tafel(title, children) {
  return table([W], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w: W, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)) })]),
    row([cell(children, { w: W, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 140, ml: 200 })]),
  ]);
}
const R = (s) => t(s, { bold: true, color: C.hv, font: "Cambria", size: 22 });
const Y = (s) => t(s, { bold: true, color: "C98A1E", font: "Cambria", size: 22 });
const B = (s) => t(s, { bold: true, color: C.inf, font: "Cambria", size: 22 });
const N = (s, o = {}) => t(s, Object.assign({ font: "Cambria", size: 22 }, o));
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));


// ================================================================
// Gemeinsame Bausteine für Tafelskripte
// ================================================================

const header = (label) => new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 7A · " + label, { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })],
  })],
});
const footer = (label) => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })],
});

// ---------- Was die Klasse gerade tut (nur zur Planung) ----------
const LABELS = {
  zuhoeren: "hört zu", gemeinsam: "macht mündlich mit", abschreiben: "schreibt ab", luecken: "arbeitet allein",
  tafel: "kommt an die Tafel", kontrolle: "kontrolliert", partner: "arbeitet zu zweit", fertig: "gibt ab",
};
function labelCell(m, w) {
  return cell([
    p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }),
    p(t(LABELS[m], { bold: true, color: NAVY, size: 20 }), { after: 0 }),
  ], { w, fill: LIGHT, borders: allBorders(solid("D5DDEE", 6)), m: 100, valign: VerticalAlign.TOP });
}

// Schritt: Modus + Inhalt
const BW = 1700, CW = W - BW;
function step(m, parts) {
  const children = [];
  if (parts.say) children.push(p([t("Ansage: ", { bold: true, size: 20, color: C.hv }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((d) => children.push(p(Array.isArray(d) ? d : S(d), { after: 50 })));
  if (parts.board) { children.push(...[].concat(parts.board)); children.push(p(t(""), { after: 20 })); }
  if (parts.sol) children.push(p([t("Lösung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  if (children.length && !parts.sol && !parts.board) children[children.length - 1] = children[children.length - 1];
  return table([BW, CW], [row([labelCell(m, BW), cell(children, { w: CW, borders: allBorders(solid("D5DDEE", 6)), m: 100, ml: 160, valign: VerticalAlign.TOP })])]);
}
const steps = (arr) => arr.flatMap((s) => [step(s[0], s[1]), p(t(""), { after: 40 })]);

// Tafel-Kasten (innerhalb eines Schritts)
const TW = CW - 340;
function board(title, children) {
  return table([TW], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 100, ml: 160 })]),
  ]);
}
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
// Merksatz steht auf der Folie (nicht an der Tafel)
function folie(title, children) {
  return table([TW], [
    row([cell(p(t("FOLIE · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: allBorders(solid(NAVY, 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: LIGHT, borders: allBorders(solid(NAVY, 12)), m: 100, ml: 160 })]),
  ]);
}
const merkBoard = (nr, title, runs) => folie(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  p(runs, { after: 0 }),
]);
// Übung an der Tafel – immer im gleichen Format
function auftrag(title, todo, lines, meta) {
  return board("rechts · " + title, [
    ...todo.map((x, i) => p([t(`${i + 1}. `, { bold: true, color: NAVY, size: 21 }), t(x, { bold: true, color: NAVY, size: 21 })], { after: 30 })),
    p(t(""), { after: 30 }),
    ...lines,
    p(t(""), { after: 30 }),
    p([t("Zeit: ", { bold: true, size: 19, color: "3D4A3F" }), t(meta, { size: 19, color: "3D4A3F" }), t("     Fertig? Stift hinlegen und noch einmal durchlesen.", { size: 19, italics: true, color: "3D4A3F" })], { after: 0 }),
  ]);
}
function phase(nr, title, min) {
  return p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
}


// Tafelskript Do 15.10. „Paddy Fingal – wiederholen und nacherzählen“ – AVIVA
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Do 15.10. · Einzelstunde (45 Min.) · AVIVA · mit Präsentation"), h1("Paddy Fingal – wiederholen und nacherzählen"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Am Mittwoch kam die Klasse nach Leistungskontrolle und Mitarbeitsnoten nur bis zum Lesetraining (Vorlesen, Kehrreim im Chor); Ü1 wurde nicht geschafft. Wiederholung ist der Lehrkraft extrem wichtig: Deshalb beginnt die Stunde mit Ü1 als Wiederholung der Lesetechniken, danach folgt LB-Aufgabe 1. Merksätze werden nicht abgeschrieben – der Merksatz Lesetechniken liegt bei Bedarf als Schnipsel bereit.")], { after: 0 })], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S("Ich kann mit den Lesetechniken Informationen in der Ballade finden und den Inhalt mit eigenen Worten nacherzählen.", { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Wenn ich etwas Bestimmtes suche, lese ich orientierend. Beim Nacherzählen nenne ich Titel und Autor, halte die Reihenfolge ein, erfinde nichts dazu und erwähne die wichtigen Details.“")], { after: 100 }));
doc.push(h3("Ablauf"));
doc.push(grid([900, 1300, 3400, 4606], ["Zeit", "AVIVA", "Phase", "Alle Schüler …"], [
  ["2′", "A", "Ankommen: Ziel und Ablauf", "lesen Ziel und Ablauf auf der Folie"],
  ["13′", "V", "Wiederholung: Lesetechniken und Ü1", "ordnen drei Situationen zu, überfliegen (1 Min.), suchen orientierend"],
  ["8′", "V", "LB-Aufgabe 1 a", "lesen im Tandem, beantworten drei Fragen im Heft"],
  ["10′", "V", "LB-Aufgabe 1 b", "erzählen sich im Tandem nach und prüfen mit der Checkliste"],
  ["6′", "A", "Auswerten: Ziel erreicht?", "ein Paar erzählt vor, alle prüfen; „Heute habe ich gelernt, …“"],
], { bold: [true, true, true, false], colors: [C.hv, C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("6 Minuten. Mit * markiert: Frage 5 in Ü1 b, „spannend mit wörtlicher Rede“ in 1 b.")], { before: 80, after: 60 }));
doc.push(p([S("Heft: ", { bold: true }), S("Überschrift „Paddy Fingal – wiederholen und nacherzählen“ · A, B, C mit Lesetechnik · Ü1 a (Satz) · Ü1 b (1–5 mit Strophe) · LB 1 a (drei Antworten in Sätzen) · „Heute habe ich gelernt, …“ · ggf. Schnipsel Merksatz Lesetechniken eingeklebt.")], { after: 40 }));
doc.push(p([S("Material: ", { bold: true }), S("Lehrbuch S. 136 f. · Präsentation (9 Folien, Lösungen in den Notizen) · Schnipsel „Merksatz Lesetechniken“ (10 pro A4, nur bei Bedarf) · Tandems wie festgelegt.")], { after: 40 }));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 2));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Auf der Folie steht unser Ziel und der Ablauf für heute. Wir wiederholen zuerst, was wir am Mittwoch nicht mehr geschafft haben. Heft auf, Überschrift: Paddy Fingal – wiederholen und nacherzählen." }]]));
doc.push(avivaPhase("V", "Vorwissen aktivieren: Wiederholung Lesetechniken und Ü1", 13));
doc.push(...steps([
  ["luecken", { say: "Drei Situationen auf der Folie. Schreibt A, B, C ins Heft und die passende Lesetechnik. Zwei Minuten.", sol: "A orientierend (Worterklärung „Hüne“) · B überfliegend · C genau (Schayas List)", do: ["Kurz begründen lassen. Folie „Zur Erinnerung“ zeigen. Wer den Merksatz noch nicht im Heft hat, bekommt den Schnipsel zum Einkleben."] }],
  ["luecken", { say: "Ü1 a: Überfliegen. Genau eine Minute. Start! … Stopp. Buch zu. Schreibt: In der Ballade geht es um …", sol: "z. B.: … um den Riesen Paddy Fingal und einen noch größeren Schotten, vor dem er sich mit einer List rettet." }],
  ["luecken", { say: "Ü1 b: Orientierend lesen. Nummeriert 1 bis 5. Sucht nur die Antworten und notiert die Strophe. Sechs Minuten.", sol: "1 sechs Ellen (Str. 1) · 2 aus Schottland (Str. 2) · 3 Schaya (Str. 4) · 4 nur seine Nase (Str. 5) · *5 flache Stelle in einem Fluss oder Bach (Worterklärung 7)" }],
  ["partner", { say: "Vergleicht zu zweit und verbessert mit Grün. Zwei Minuten." }],
]));
doc.push(avivaPhase("V", "Verarbeiten: LB-Aufgabe 1 a", 8));
doc.push(...steps([["partner", { say: "Jetzt arbeitet ihr im Tandem. Lest die Ballade halblaut zusammen. Dann beantwortet ihr die drei Fragen von der Folie im Heft – in ganzen Sätzen, mit Strophe. Acht Minuten.", sol: "Aussehen: Riese, schwarzer Bart sechs Ellen lang, Fäuste so groß wie Tonnen, Mastbaum als Wurfgeschoss (Str. 1) · Der Schotte will herausfinden, wer stärker ist (Str. 2) · List: Schaya steckt Paddy ins Bett, türmt die Kissen hoch und gibt ihn als Baby aus; der Schotte sieht nur die riesige Nase, denkt an einen riesigen Vater und flieht (Str. 5–10).", do: ["Methodenkarte „Tandemlesen“."] }]]));
doc.push(avivaPhase("V", "Verarbeiten: LB-Aufgabe 1 b nacherzählen", 10));
doc.push(...steps([["partner", { say: "Im Tandem: A erzählt die Ballade nach, B hört zu und hakt die Checkliste auf der Folie ab. Dann Wechsel. Eure Antworten aus 1 a sind euer Spickzettel.", do: ["Checkliste: Titel und Autor · richtige Reihenfolge, nichts dazuerfunden · alle wichtigen Details · * spannend mit wörtlicher Rede.", "Herumgehen, ein gelungenes Paar für die Auswertung ansprechen."] }]]));
doc.push(avivaPhase("A", "Auswerten: Ziel erreicht?", 6));
doc.push(...steps([
  ["gemeinsam", { say: "Ein Paar erzählt vor. Alle anderen prüfen mit der Checkliste: Was ist gelungen, was fehlt?" }],
  ["luecken", { say: "Schaut auf unser Ziel. Schreibt ins Heft: Heute habe ich gelernt, …", do: ["Zwei Schüler lesen vor."] }],
]));
const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Do 15.10.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Do_15-10_Paddy_Fingal_Nacherzaehlen.docx", buf));
