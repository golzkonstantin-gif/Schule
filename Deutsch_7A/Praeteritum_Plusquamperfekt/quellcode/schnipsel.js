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


// ================================================================
// Kopiervorlage: Schnipsel A–D (je 2 Schüler pro A4-Seite)
// ================================================================
const PW = 11906 - 1200;          // nutzbare Breite (Rand 600)
const HW = Math.floor((PW - 200) / 2); // halbe Spalte
const sz = 19;
const s = (txt, o = {}) => t(txt, Object.assign({ size: sz }, o));
const sb = (txt, o = {}) => s(txt, Object.assign({ bold: true }, o));
const gray = "9AA3B5";
const para = (runs, o = {}) => p(runs, Object.assign({ after: 40 }, o));
function slipHead(nr, title) {
  return [
    p([t(nr + "  ", { font: "Cambria", bold: true, size: 22, color: C.hv }), t(title, { font: "Cambria", bold: true, size: 22, color: NAVY })], { after: 20 }),
    p([s("Name: ____________________", { color: MUTED, size: 16 })], { after: 60 }),
  ];
}
const slipCell = (children, w) => cell(children, { w, borders: allBorders(dashed), m: 120, ml: 140, valign: VerticalAlign.TOP });
const scissors = () => p(t("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", { color: gray, size: 16 }), { before: 100, after: 100, align: AlignmentType.CENTER });

// ---------- Schnipsel A ----------
const VERBS = ["werfen", "fallen", "greifen", "halten", "fangen", "treffen", "laufen", "heben", "reißen", "stoßen", "gewinnen", "verlieren", "kämpfen", "starten", "trainieren", "sein", "haben"];
function slipA() {
  const w = [1900, 2400, 650];
  return [
    ...slipHead("Ü9", "Präteritum der Sportverben"),
    para([s("Trage die Präteritumform ein (ich …). Markiere unregelmäßige Verben mit "), sb("*"), s(".")]),
    table(w, [
      row([hdr("Infinitiv", w[0], C.inf, 17), hdr("Präteritum (ich …)", w[1], NAVY, 17), hdr("*", w[2], MUTED, 17)], 300),
      ...VERBS.map((v) => row([tc(v, w[0], { size: 18, font: "Cambria" }), tc("", w[1]), tc("", w[2])], 330, HeightRule.EXACT)),
    ]),
  ];
}
// ---------- Schnipsel B ----------
const PQ = [["Ich war müde.", "Ich ______ die ganze Woche ______.", "trainieren"], ["Der Trainer war zufrieden.", "Wir ______ alles richtig ______.", "machen"], ["Sie gewann den Kampf.", "Sie ______ sich gut ______.", "aufwärmen"], ["Wir kamen pünktlich an.", "Wir ______ früh ______.", "losfahren"], ["Er stieß 80 Kilo.", "Das ______ er noch nie ______.", "schaffen"]];
function slipB() {
  return [
    ...slipHead("Ü10", "Was war vorher?"),
    para([s("Ergänze im zweiten Satz das "), sb("Plusquamperfekt"), s(". Markiere hatte/war "), sb("rot", { color: C.hv }), s(" und das Partizip II "), sb("gelb", { color: "C98A1E" }), s(".")]),
    para([s("Beispiel: ", { italics: true, color: MUTED }), s("Ich war nervös. Ich "), sb("hatte", { color: C.hv }), s(" kaum "), sb("geschlafen", { color: "C98A1E" }), s(". (schlafen)")], { after: 100 }),
    ...PQ.flatMap(([a, b, v], i) => [
      p([sb(`${i + 1}  `, { color: C.hv }), t(a, { font: "Cambria", size: sz })], { after: 20 }),
      p([s("    "), t(b, { font: "Cambria", size: sz }), s(`  (${v})`, { italics: true, color: MUTED })], { after: 120 }),
    ]),
  ];
}
// ---------- Schnipsel C ----------
function slipC() {
  const w3 = [1620, 1620, 1620];
  const col = (h, items) => cell([p(t(h, { bold: true, size: 17, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER })], { w: 1620, fill: NAVY, borders: allBorders(solid("FFFFFF", 6)) });
  const lst = (items) => cell(items.map((x) => p(s(x, { size: 17 }), { after: 10 })), { w: 1620, fill: LIGHT, borders: allBorders(solid("FFFFFF", 6)), valign: VerticalAlign.TOP });
  return [
    ...slipHead("Ü11", "Mein erster Wettkampf"),
    para([s("Schreibe "), sb("5–6 Sätze"), s(" ins Heft – im "), sb("Präteritum"), s(", mit mindestens einem Satz im "), sb("Plusquamperfekt"), s(" (Vorgeschichte).")], { after: 80 }),
    p(sb("Satzanfänge", { color: NAVY }), { after: 20 }),
    para([s("Vorher hatte ich …  ·  Am Wettkampftag …  ·  In der Halle …  ·  Beim ersten Kampf / In der ersten Halbzeit / Im ersten Versuch …  ·  Dann …  ·  Am Ende …")], { after: 80 }),
    p(sb("Bausteine für die Vorgeschichte", { color: NAVY }), { after: 20 }),
    para([s("hatte wochenlang trainiert  ·  hatte mich lange aufgewärmt  ·  hatte kaum geschlafen  ·  war früh losgefahren  ·  hatte noch nie vor so vielen Zuschauern gekämpft / gespielt")], { after: 80 }),
    p(sb("Wörter für deine Sportart", { color: NAVY }), { after: 30 }),
    table(w3, [
      row([col("Judo"), col("Handball"), col("Gewichtheben")]),
      row([lst(["die Matte", "der Gegner", "der Wurf", "der Haltegriff", "der Ippon"]), lst(["das Tor", "die Halbzeit", "der Siebenmeter", "die Abwehr", "der Schiedsrichter"]), lst(["die Hantel", "der Versuch", "das Reißen", "das Stoßen", "der Kampfrichter"])]),
    ]),
    p(s("Die Präteritumformen findest du in Ü9.", { italics: true, color: MUTED, size: 17 }), { before: 60, after: 0 }),
  ];
}
// ---------- Schnipsel D ----------
function slipD() {
  const box3 = (q) => [p([s("☐ ja   ☐ nein    ", { size: 18 }), s(q)], { after: 80 })];
  return [
    p([t("Partner-Check", { font: "Cambria", bold: true, size: 22, color: NAVY })], { after: 20 }),
    p([s("Text von: _______________   Geprüft von: _______________", { color: MUTED, size: 16 })], { after: 80 }),
    p(sb("So prüfst du:", { color: NAVY }), { after: 20 }),
    para([sb("1. "), s("Unterstreiche alle "), sb("Präteritum-Verben"), s(".")]),
    para([sb("2. "), s("Markiere im Plusquamperfekt hatte/war "), sb("rot", { color: C.hv }), s(" und das Partizip II "), sb("gelb", { color: "C98A1E" }), s(".")]),
    para([sb("3. "), s("Beantworte die Fragen. Verbessere nichts selbst!")], { after: 100 }),
    ...box3("Der Text steht durchgehend im Präteritum."),
    ...box3("Es gibt mindestens einen Satz zur Vorgeschichte im Plusquamperfekt."),
    ...box3("Die Verbformen stimmen."),
    p(s("Wenn nicht – wo? ____________________________", { size: 18 }), { after: 120 }),
    p(s("Das ist dir gut gelungen: ____________________", { size: 18 }), { after: 120 }),
    p(s("______________________________________________", { size: 18, color: gray }), { after: 0 }),
  ];
}

const pair = (left, right) => table([HW, 200, HW], [row([slipCell(left, HW), cell(p(t("")), { w: 200, borders: noBorders }), slipCell(right, HW)])]);
const doc = [
  pair(slipA(), slipB()), scissors(), pair(slipA(), slipB()),
  new Paragraph({ children: [new PageBreak()] }),
  pair(slipC(), slipD()), scissors(), pair(slipC(), slipD()),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 500, left: 600, right: 600 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: doc }] })).then((b) => fs.writeFileSync("Schnipsel_Praeteritum_Plusquamperfekt.docx", b));
