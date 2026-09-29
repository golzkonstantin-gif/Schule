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
// Zusatzblatt für schnelle Schüler – Vorderseite Aufgaben, Rückseite Lösungen
// ================================================================
const header = new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 7A · Verbformen · Zusatzblatt", { size: 17, color: MUTED }), t("\tName: ______________________", { size: 17, color: MUTED })],
  })],
});

const LINE = "_____________________________________________";
const ln = (txt = "", o = {}) => p([S(txt, { size: 20 }), S(LINE, { size: 20, color: "9AA3B5" })], Object.assign({ after: 70 }, o));

// Aufgabenkopf im Format der Stunde: Nummer + Titel, nummerierte Schritte, Arbeitsform
function aufgabe(nr, title, todo, star = "") {
  return [
    p([t(`${nr}  `, { font: "Cambria", size: 25, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 25, bold: true, color: NAVY }), t(star ? "   " + star : "", { size: 21, bold: true, color: "C98A1E" })], { before: 110, after: 30, keepNext: true }),
    ...todo.map((x, i) => p([t(`${i + 1}. `, { bold: true, size: 20, color: NAVY }), t(x, { bold: true, size: 20, color: NAVY })], { after: i === todo.length - 1 ? 80 : 20, keepNext: true })),
  ];
}

const doc = [];
doc.push(p([t("Schon fertig? ", { font: "Cambria", size: 34, bold: true, color: NAVY }), t("Profi-Aufgaben zu den Verbformen", { font: "Cambria", size: 26, color: MUTED })], { after: 40 }));
doc.push(box([p([S("So arbeitest du: ", { bold: true, size: 20 }), S("allein und leise. Kontrolle mit der Rückseite, verbessern mit Grün.", { size: 20 })], { after: 0 })], LIGHT2));

// ---------- Z1 Fehler-Profi ----------
doc.push(...aufgabe("Z1", "Fehler-Profi", ["Lies den Post und unterstreiche die fünf Fehler bei den Verbformen.", "Schreibe darunter: falsche Form → richtige Form."]));
doc.push(box([
  p([t("Klassenfahrt, Tag 2", { bold: true, size: 20, color: MUTED })], { after: 40 }),
  p(N("Heute haben wir schon um sieben aufgestanden. Danach haben wir mit dem Bus zum Strand gefahren. Dort haben wir Volleyball gespielen. Lea hat beim Hechtsprung ins Wasser gefallen, alle haben gelacht. Abends haben wir noch Musik gehört, und ich bin sofort eingeschlaft.", { size: 21 }), { after: 0, line: 276 }),
], LIGHT));
doc.push(p(t(""), { after: 30 }));
{
  const w = [3402, 3402, 3402];
  const r = () => row(w.map((x) => cell(p([S("___________ → ___________", { size: 20, color: "9AA3B5" })], { after: 0 }), { w: x, borders: allBorders(none), m: 70 })));
  doc.push(table(w, [r(), r()]));
}

// ---------- Z2 Perfekt bilden ----------
doc.push(...aufgabe("Z2", "Ab ins Perfekt", ["Setze jeden Satz ins Perfekt.", "Entscheide: haben oder sein? Achte auf das Partizip II."]));
[
  "Wir fahren am Samstag nach Hamburg.",
  "Ich schreibe meiner besten Freundin.",
  "Das Spiel beginnt um acht.",
  "Mein Bruder wacht mitten in der Nacht auf.",
].forEach((sentence, i) => {
  doc.push(p([S(`${i + 1}  `, { bold: true, color: C.hv, size: 20 }), N(sentence, { size: 21 })], { after: 20, keepNext: true }));
  doc.push(p(S("     ______________________________________________________________________", { size: 20, color: "9AA3B5" }), { after: 130 }));
});

// ---------- Z3 Eigene Sätze ----------
doc.push(...aufgabe("Z3", "Hilfsverb oder Vollverb – selbst gemacht", ["Schreibe in jede Zeile einen eigenen Satz, der dazu passt.", "Markiere Hilfsverben rot und Partizipien gelb."]));
{
  const w = [2600, 7606];
  doc.push(table(w, [["haben · Hilfsverb", "haben · Vollverb", "sein · Hilfsverb", "sein · Vollverb"].map((h) =>
    row([tc(h, w[0], { bold: true, size: 19, color: NAVY, fill: LIGHT }), tc("", w[1])], 520))].flat()));
}

// ---------- Z4 ★ Regel entdecken ----------
doc.push(...aufgabe("Z4", "Regel-Detektiv", ["Bilde das Partizip II und sortiere es in die richtige Spalte.", "Formuliere die Regel in einem Satz."], "★"));
doc.push(p([S("Verben: ", { bold: true, size: 20 }), N("verstehen · aufräumen · telefonieren · einkaufen · besuchen · mitspielen · erzählen · ausprobieren", { size: 21 })], { after: 60 }));
{
  const w = [5103, 5103];
  doc.push(table(w, [
    row([hdr("kein ge-", w[0], C.p2, 19), hdr("ge- in der Mitte", w[1], C.p2, 19)], 340),
    row(w.map((x) => tc("", x)), 1000),
  ]));
  doc.push(p([S("Regel: ", { bold: true, size: 20 }), S(LINE + "__________________________", { size: 20, color: "9AA3B5" })], { before: 80, after: 0 }));
}

// ================= Rückseite: Lösungen =================
doc.push(br(), p([t("Lösungen", { font: "Cambria", size: 34, bold: true, color: NAVY }), t("   Erst nachschauen, wenn du fertig bist!", { size: 21, italics: true, color: C.hv })], { after: 120 }));
const sol = (nr, title, lines) => [
  p([t(`${nr}  `, { font: "Cambria", size: 23, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 23, bold: true, color: NAVY })], { before: 120, after: 50, keepNext: true }),
  ...lines.map((l) => p(Array.isArray(l) ? l : S(l, { size: 20 }), { after: 40 })),
];
doc.push(...sol("Z1", "Fehler-Profi", [
  "haben … aufgestanden → sind aufgestanden (Veränderung: vom Liegen zum Stehen)",
  "haben … gefahren → sind gefahren (Bewegung von A nach B)",
  "gespielen → gespielt (regelmäßiges Verb: ge-…-t)",
  "hat … gefallen → ist gefallen (Bewegung: von oben nach unten)",
  "eingeschlaft → eingeschlafen (unregelmäßiges Verb: ge-…-en, ge- in der Mitte)",
]));
doc.push(...sol("Z2", "Ab ins Perfekt", [
  "1  Wir sind am Samstag nach Hamburg gefahren.",
  "2  Ich habe meiner besten Freundin geschrieben.",
  "3  Das Spiel hat um acht begonnen.  (kein ge-, weil das Verb mit be- beginnt)",
  "4  Mein Bruder ist mitten in der Nacht aufgewacht.  (Veränderung)",
]));
doc.push(...sol("Z3", "Hilfsverb oder Vollverb", [
  "Beispiele: Ich habe Kopfschmerzen. (Vollverb) · Ich habe lange geschlafen. (Hilfsverb) · Wir sind müde. (Vollverb) · Wir sind nach Hause gerannt. (Hilfsverb)",
  "Prüfe deine Sätze mit dem Test: Steht am Satzende ein Partizip II? Nur dann ist haben/sein ein Hilfsverb.",
]));
doc.push(...sol("Z4", "Regel-Detektiv", [
  "kein ge-: verstanden · telefoniert · besucht · erzählt · ausprobiert",
  "ge- in der Mitte: aufgeräumt · eingekauft · mitgespielt",
  "Regel: Verben auf be-, ver-, er- und -ieren bekommen kein ge-. Bei trennbaren Verben (auf-, ein-, mit-) steht ge- zwischen Vorsilbe und Stamm.",
  "Knifflig: ausprobiert hat beides – aus- ist trennbar, probieren endet auf -ieren, deshalb kein ge-.",
]));
const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 500, left: 850, right: 850, header: 340, footer: 280 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const d = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header }, children: doc }] });
Packer.toBuffer(d).then((b) => fs.writeFileSync("Zusatzblatt_Verbformen_Schnelle.docx", b));
