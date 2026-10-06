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
// Tafelskript „Präteritum und Plusquamperfekt – Mein erster Wettkampf“ (90 Min.)
// ================================================================
const merkTafel = (nr, title, children) => board(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  ...children,
]);
const U = (s) => t(s, { bold: true, font: "Cambria", size: 22, u: true }); // Präteritum-Verb: unterstrichen
const sl = (runs, o = {}) => p(runs, Object.assign({ after: 30 }, o));
function twoCols(h1, h2, rows, w = [3900, 3900]) {
  return table(w, [
    row([cell(p(N(h1, { bold: true, color: NAVY }), { after: 0 }), { w: w[0], borders: noBorders }), cell(p(N(h2, { bold: true, color: NAVY }), { after: 0 }), { w: w[1], borders: noBorders })]),
    ...rows.map(([a, b]) => row([cell(p(a, { after: 0 }), { w: w[0], borders: noBorders, m: 15 }), cell(p(b, { after: 0 }), { w: w[1], borders: noBorders, m: 15 })])),
  ]);
}
const n19 = (s, o = {}) => N(s, Object.assign({ size: 20 }, o));
const r19 = (s) => t(s, { bold: true, color: C.hv, font: "Cambria", size: 20 });
const y19 = (s) => t(s, { bold: true, color: "C98A1E", font: "Cambria", size: 20 });
const b19 = (s) => t(s, { bold: true, color: C.inf, font: "Cambria", size: 20 });
const u19 = (s) => t(s, { bold: true, font: "Cambria", size: 20, u: true });

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Doppelstunde (90 Min.)"), h1("Präteritum und Plusquamperfekt"));
doc.push(p(t("Rahmen: „Mein erster Wettkampf“", { font: "Cambria", size: 26, italics: true, color: MUTED }), { after: 100 }));
doc.push(box([
  p([S("Anknüpfung: ", { bold: true }), S("Infinitiv, finite Form und Partizip II sitzen, das Perfekt mit haben und sein ebenfalls. Merksatz 4 (hatte, war) steht im Heft. Der Rückbau Futur I → Perfekt lief gut und wird als Reaktivierung wiederholt. Die Klasse besteht aus Leistungssportlern (Judo, Handball, Gewichtheben) – deshalb trägt der eigene erste Wettkampf die ganze Doppelstunde: mündlich erzählt man im Perfekt, aufgeschrieben steht er im Präteritum, die Vorgeschichte im Plusquamperfekt.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Ziel: ", { bold: true }), S("Die Klasse bildet Präteritum (ohne Hilfsverb) und Plusquamperfekt (hatte/war + Partizip II), schreibt einen kurzen Wettkampfbericht und reflektiert, wann man welche Zeitform benutzt.")], { before: 120, after: 60 }));
doc.push(p([S("Markieren wie bisher, ohne Klammer-Bögen: ", { bold: true }), t("Hilfsverb rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · "), S("Präteritum-Verb (ohne Hilfsverb): unterstreichen", { bold: true })], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 3500, 5806], ["Zeit", "Phase", "Kern"], [
  ["10′", "0 Reaktivierung", "Ü8: Futur I → Perfekt per Kreide-Kette, alle schreiben mit"],
  ["8′", "1 Erzählt euch!", "Partner erzählen ihren ersten Wettkampf – Lehrkraft notiert Sätze (Perfekt)"],
  ["12′", "2 Präteritum entdecken", "Dieselben Sätze „in der Vereinszeitung“ – Merksatz 8"],
  ["12′", "3 Ü9 Präteritum der Sportverben", "Verben ins Heft schreiben und ergänzen, Vergleich per Kreide-Kette"],
  ["20′", "4 Plusquamperfekt", "„Ich war nervös. Ich hatte kaum geschlafen.“ – Zeitstrahl, haben/sein im Präteritum konjugieren, Merksatz 9, Ü10 ins Heft"],
  ["15′", "5 Ü11 Mein erster Wettkampf", "eigenen Bericht ins Heft schreiben (gedruckter Schreibrahmen)"],
  ["8′", "6 Partner-Check", "unterstreichen, markieren, drei Prüffragen (an der Tafel)"],
  ["5′", "7 Reflexion", "Tabelle „Wann nehme ich welche Zeitform?“"],
], { bold: [true, true, false], colors: [C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("Die 90 Minuten sind damit voll verplant. Wird es knapp, ist Phase 5 der Kern. Der Bericht kann als Hausaufgabe fertig geschrieben werden; Partner-Check und Reflexion (6, 7) sind dann der Einstieg der nächsten Stunde.")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Präteritum und Plusquamperfekt · Datum rechts"],
  ["Ü8", "zwei Spalten: Futur I | Perfekt (mitgeschrieben)"],
  ["Tafelbild", "zwei Spalten: So erzählen wir | So steht es in der Vereinszeitung"],
  ["Merksatz 8", "Das Präteritum"],
  ["Ü9", "Infinitiv → Präteritum der Sportverben (z. B. werfen → warf*)"],
  ["Tabelle", "haben und sein im Präteritum (ich hatte, ich war …)"],
  ["Merksatz 9", "Das Plusquamperfekt + kleiner Zeitstrahl"],
  ["Ü10", "nur Nummer + Plusquamperfekt: 1 hatte trainiert …"],
  ["Ü11", "eigener Bericht „Mein erster Wettkampf“ (5–6 Sätze)"],
  ["Tabelle", "Wann nehme ich welche Zeitform?"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : (a === "Überschrift" || a === "Tafelbild" || a === "Tabelle") ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Schreibrahmen Ü11 (4 Stück pro A4 – ein Viertel der Klassenstärke kopieren und zerschneiden) · Kreide Rot, Gelb · Schüler: Heft, Buntstifte Rot, Gelb, Grün")], { before: 100, after: 40 }));
doc.push(p([S("Tipp: ", { bold: true }), S("Merksatz 8 und 9 vorab auf die Seitentafel schreiben und zuklappen.")], { after: 0 }));

// ---------- Phase 0 ----------
const FP = [
  [["Am Samstag ", "werde", " ich gegen den Titelverteidiger ", "kämpfen", "."], ["Am Samstag ", "habe", " ich gegen den Titelverteidiger ", "gekämpft", "."]],
  [["Unsere Mannschaft ", "wird", " das Spiel ", "gewinnen", "."], ["Unsere Mannschaft ", "hat", " das Spiel ", "gewonnen", "."]],
  [["Ich ", "werde", " 60 Kilo ", "reißen", "."], ["Ich ", "habe", " 60 Kilo ", "gerissen", "."]],
  [["Wir ", "werden", " mit dem Bus zum Turnier ", "fahren", "."], ["Wir ", "sind", " mit dem Bus zum Turnier ", "gefahren", "."]],
];
doc.push(phase("0", "Reaktivierung: Futur I → Perfekt", 10));
doc.push(...steps([
  ["abschreiben", { say: "Heft auf. Datum nach rechts. Überschrift: Präteritum und Plusquamperfekt. Darunter Ü8 mit zwei Spalten: links Futur I, rechts Perfekt. Schreibt die Sätze links ab.", board: board("Mitte · Ü8", [twoCols("Futur I", "Perfekt", FP.map(([f]) => [[n19(f[0]), r19(f[1]), n19(f[2]), b19(f[3]), n19(f[4])], [n19("")]]))]) }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt einen Satz rechts ins Perfekt – Hilfsverb rot, Partizip II gelb. Alle anderen schreiben mit. Alle prüfen mit.", board: board("Mitte · Ü8 (danach)", [twoCols("Futur I", "Perfekt", FP.map(([f, pf]) => [[n19(f[0]), r19(f[1]), n19(f[2]), b19(f[3]), n19(f[4])], [n19(pf[0]), r19(pf[1]), n19(pf[2]), y19(pf[3]), n19(pf[4])]]))]), sol: "habe gekämpft · hat gewonnen · habe gerissen · sind gefahren (Bewegung → sein)" }],
]));

// ---------- Phase 1 ----------
doc.push(phase("1", "Erzählt euch: Euer erster Wettkampf", 8));
doc.push(...steps([
  ["partner", { say: "Jetzt arbeitet ihr zu zweit, ohne Heft. Erzählt euch gegenseitig von eurem ersten Wettkampf – oder von einem Wettkampf, an den ihr euch besonders gut erinnert. Jeder hat eine Minute. Ich höre zu.", do: ["Herumgehen und zuhören. Drei typische Sätze notieren – sie kommen automatisch im Perfekt („Ich bin total nervös gewesen“).", "Die Alternative „ein Wettkampf, an den du dich gut erinnerst“ bewusst anbieten: Nicht jeder hat gute Erinnerungen an den ersten Wettkampf."] }],
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Ich schreibe drei Sätze an, die ich gerade gehört habe.", board: board("Mitte · links: „So erzählen wir“ (Beispiele, falls nötig)", [
    sl([n19("Ich "), r19("bin"), n19(" total nervös "), y19("gewesen"), n19(".")]),
    sl([n19("Ich "), r19("habe"), n19(" meinen ersten Kampf "), y19("verloren"), n19(".")]),
    sl([n19("Wir "), r19("haben"), n19(" mit 25:20 "), y19("gewonnen"), n19(".")], { after: 0 }),
  ]), do: ["Hilfsverb rot und Partizip II gelb gleich mitmarkieren – oder per Kreide-Kette markieren lassen."] }],
]));

// ---------- Phase 2 ----------
const ZT = [
  [[n19("Ich "), r19("bin"), n19(" total nervös "), y19("gewesen"), n19(".")], [n19("Ich "), u19("war"), n19(" total nervös.")]],
  [[n19("Ich "), r19("habe"), n19(" meinen ersten Kampf "), y19("verloren"), n19(".")], [n19("Ich "), u19("verlor"), n19(" meinen ersten Kampf.")]],
  [[n19("Wir "), r19("haben"), n19(" mit 25:20 "), y19("gewonnen"), n19(".")], [n19("Wir "), u19("gewannen"), n19(" mit 25:20.")]],
];
doc.push(phase("2", "Präteritum entdecken", 12));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Stellt euch vor, über euren Wettkampf erscheint ein Bericht in der Vereinszeitung. Dort steht er so.", board: board("Mitte · zwei Spalten", [twoCols("So erzählen wir", "So steht es in der Vereinszeitung", ZT)]), do: ["Rechte Spalte neben die Schülersätze schreiben, die Verben unterstreichen."] }],
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf. Was ist rechts anders als links? Wo hört ihr die linke Form, wo lest ihr die rechte?", do: [
    "Ziel 1 (Form): Rechts gibt es kein Hilfsverb und kein Partizip II – nur ein einziges, finites Verb. Begriff: Präteritum.",
    "Ziel 2 (Sprachnutzung): Links erzählt man mündlich (Perfekt), rechts schreibt man – Zeitung, Bericht, Erzählung (Präteritum).",
    "Nachfragen: „Wie heißt der Infinitiv zu verlor, gewannen?“ – Stamm verändert sich (unregelmäßig). Zum Vergleich ein regelmäßiges: kämpfen → kämpfte.",
  ] }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 8 mit den Beispielen. Rot umrahmen.", board: merkTafel(8, "Das Präteritum", [
    p([S("Das "), S("Präteritum", { bold: true }), S(" ist eine einfache Zeitform: Es besteht nur aus "), S("einer finiten Verbform", { bold: true }), S(" – ohne Hilfsverb. Regelmäßige Verben bekommen "), S("-te", { bold: true }), S(" (kämpfte, startete). Unregelmäßige Verben ändern den Stamm (warf, fiel, gewann).")], { after: 60 }),
    p([S("Beispiel: ", { bold: true, color: MUTED }), n19("Ich "), u19("kämpfte"), n19(" gut, aber ich "), u19("verlor"), n19(" knapp.")], { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Mit dem Präteritum erzählt und berichtet man "), S("schriftlich", { bold: true }), S(" von Vergangenem – in Berichten und Erzählungen. Mündlich benutzt man meist das Perfekt. "), S("Test: ", { bold: true }), S("Steht nur ein Verb in einer Vergangenheitsform da, ohne Hilfsverb? Dann ist es Präteritum.")], { after: 0 }),
  ]) }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Ü9 Präteritum der Sportverben", 12));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts an der Tafel: Ü9. Schreibt die Verben untereinander ins Heft und ergänzt die Präteritumform. Fünf Minuten.", board: auftrag("Ü9 Präteritum der Sportverben", ["Schreibe Ü9 an den Rand.", "Schreibe jedes Verb ab und ergänze die Präteritumform: werfen → ich warf.", "Markiere unregelmäßige Verben mit einem Stern."], [p(S("werfen, fallen, greifen, halten, fangen, treffen, laufen, heben, reißen, stoßen, gewinnen, verlieren, kämpfen, starten, trainieren, sein, haben", { size: 20 }), { after: 0 })], "5 Minuten · allein · leise"),
    sol: "warf* · fiel* · griff* · hielt* · fing* · traf* · lief* · hob* · riss* · stieß* · gewann* · verlor* · kämpfte · startete · trainierte · war* · hatte*" }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt ein Verb mit Präteritumform an die Tafel. Alle prüfen mit und verbessern mit Grün.", do: ["Mit werfen beginnen – das kennen Judoka und Handballer. Danach gezielt Schüler der jeweiligen Sportart nach „ihren“ Verben fragen (Judo: greifen, halten; Handball: fangen, treffen; Gewichtheben: heben, reißen, stoßen).", "Erkenntnis: Die meisten Sportverben sind unregelmäßig – genau die müsst ihr für euren Bericht können."] }],
]));

// ---------- Phase 4 ----------
doc.push(phase("4", "Das Plusquamperfekt", 20));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Zwei Sätze aus einem Wettkampfbericht. Was ist zuerst passiert?", board: board("Mitte", [
    sl([n19("Ich "), u19("war"), n19(" vor dem Kampf sehr nervös. Ich "), r19("hatte"), n19(" die Nacht davor kaum "), y19("geschlafen"), n19(".")]),
    p(t(""), { after: 30 }),
    table([2600, 2600, 2600], [row([
      cell(p([t("hatte kaum geschlafen", { bold: true, size: 18, color: NAVY })], { after: 0, align: AlignmentType.CENTER }), { w: 2600, fill: "E4E9F5", borders: allBorders(solid(NAVY, 6)) }),
      cell(p([t("war nervös", { bold: true, size: 18, color: NAVY })], { after: 0, align: AlignmentType.CENTER }), { w: 2600, fill: "E4E9F5", borders: allBorders(solid(NAVY, 6)) }),
      cell(p([t("heute", { bold: true, size: 18, color: NAVY })], { after: 0, align: AlignmentType.CENTER }), { w: 2600, fill: "E4E9F5", borders: allBorders(solid(NAVY, 6)) }),
    ]), row([
      cell(p(t("Vorvergangenheit", { size: 16, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: 2600, borders: noBorders }),
      cell(p(t("Vergangenheit", { size: 16, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: 2600, borders: noBorders }),
      cell(p(t("Gegenwart  →", { size: 16, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: 2600, borders: noBorders }),
    ])]),
  ]), do: ["Zeitstrahl darunter zeichnen. Antwort: Das Schlafen war vorher – vor der Vergangenheit."] }],
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf. Schaut auf hatte und geschlafen: Was davon kennt ihr schon?", do: [
    "hatte ist haben im Präteritum, geschlafen ist ein Partizip II. Damit alle hatte und war sicher bilden können, konjugieren wir beide Hilfsverben gleich vollständig.",
    "Kontrast: Ich habe geschlafen (Perfekt) – Ich hatte geschlafen (Plusquamperfekt). Der einzige Unterschied steckt im Hilfsverb: Präsens oder Präteritum.",
    "Auch mit sein: Wir waren früh losgefahren.",
  ] }],
  ["tafel", { say: "Kreide-Kette: Wir konjugieren haben und sein im Präteritum. Wer die Kreide hat, trägt eine Form ein und gibt die Kreide weiter. Alle anderen schreiben die Tabelle ins Heft.", board: board("Mitte · haben und sein im Präteritum", [table([1500, 1900, 1900], [
    row([cell(p(N(""), { after: 0 }), { w: 1500, borders: allBorders(solid("9AA59C", 4)) }), cell(p(R("haben"), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)) }), cell(p(R("sein"), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)) })]),
    ...[["ich", "hatte", "war"], ["du", "hattest", "warst"], ["er/sie/es", "hatte", "war"], ["wir", "hatten", "waren"], ["ihr", "hattet", "wart"], ["sie/Sie", "hatten", "waren"]].map(([ps, h, w]) => row([cell(p(N(ps, { color: MUTED }), { after: 0 }), { w: 1500, borders: allBorders(solid("9AA59C", 4)), m: 15 }), cell(p(N(h), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)), m: 15 }), cell(p(N(w), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)), m: 15 })])),
  ])]), sol: "hatte · hattest · hatte · hatten · hattet · hatten | war · warst · war · waren · wart · waren", do: ["Die Personen und die Überschriften vorher anschreiben, die Formen trägt die Kreide-Kette ein.", "Stolperstellen: ihr wart (ohne e), du warst; die 1. und 3. Person Singular sind gleich (ich hatte – er hatte, ich war – sie war).", "Kurz nach Sportarten anwenden lassen: „Wir ___ müde.“ – „Ihr ___ gut trainiert.“"] }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 9 mit dem Beispiel und dem kleinen Zeitstrahl. Rot umrahmen.", board: merkTafel(9, "Das Plusquamperfekt", [
    p([S("Das "), S("Plusquamperfekt", { bold: true }), S(" bildet man mit "), S("haben oder sein im Präteritum", { bold: true }), S(" (hatte, war) und dem "), S("Partizip II", { bold: true }), S(".")], { after: 60 }),
    p([S("Beispiel: ", { bold: true, color: MUTED }), n19("Ich "), r19("hatte"), n19(" wochenlang "), y19("trainiert"), n19(".  ·  Wir "), r19("waren"), n19(" früh "), y19("losgefahren"), n19(".")], { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Damit erzählt man, was "), S("vorher", { bold: true }), S(" passiert war – die Vorgeschichte zu einem Ereignis in der Vergangenheit. "), S("Test: ", { bold: true }), S("Steht hatte oder war mit einem Partizip II? Dann ist es Plusquamperfekt. (Steht habe oder bin, ist es Perfekt.)")], { after: 0 }),
  ]) }],
  ["luecken", { say: "Jetzt arbeitet ihr allein: Ü10. Ihr schreibt die Sätze nicht ab – nur die Nummer und das Plusquamperfekt. Fünf Minuten.", board: auftrag("Ü10 Was war vorher?", ["Schreibe Ü10 an den Rand.", "Schreibe nur Nummer und Plusquamperfekt: 1 hatte … trainiert.", "Markiere hatte/war rot und das Partizip II gelb."], [
    bl("1  Ich war müde. Ich ___ die ganze Woche ___. (trainieren)"),
    bl("2  Der Trainer war zufrieden. Wir ___ alles richtig ___. (machen)"),
    bl("3  Sie gewann den Kampf. Sie ___ sich gut ___. (aufwärmen)"),
    bl("4  Wir kamen pünktlich an. Wir ___ früh ___. (losfahren)"),
    bl("5  Er stieß 80 Kilo. Das ___ er noch nie ___. (schaffen)", { after: 0 }),
  ], "5 Minuten · allein · leise"), sol: "1 hatte trainiert · 2 hatten gemacht · 3 hatte aufgewärmt · 4 waren losgefahren · 5 hatte geschafft" }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt eine Lösung an die Tafel und sagt, was zuerst passiert ist. Alle verbessern mit Grün.", do: ["Bei Nr. 4 nachfragen: Warum war und nicht hatte? → losfahren ist Bewegung, wie beim Perfekt (sind losgefahren)."] }],
]));

// ---------- Phase 5 ----------
doc.push(phase("5", "Ü11 Mein erster Wettkampf (gedruckter Schreibrahmen)", 15));
doc.push(...steps([
  ["luecken", { say: "Jetzt schreibt ihr allein euren eigenen Bericht für die Vereinszeitung – ins Heft, als Ü11. Der Schreibrahmen hilft euch mit Satzanfängen und Wörtern. Zwölf Minuten.", board: auftrag("Ü11 Mein erster Wettkampf", ["Schreibe 5–6 Sätze über deinen ersten Wettkampf (oder einen, an den du dich gut erinnerst) ins Heft.", "Schreibe im Präteritum – wie in der Vereinszeitung.", "Schreibe mindestens einen Satz zur Vorgeschichte im Plusquamperfekt („Vorher …“)."], [p(S("Hilfe: Satzanfänge und Bausteine auf dem Schreibrahmen, Präteritumformen in Ü9.", { size: 20, italics: true }), { after: 0 })], "12 Minuten · allein · leise"),
    do: ["Den Schreibrahmen neben das Heft legen (oder nach der Stunde einkleben) – geschrieben wird ins Heft.", "Typische Fehler: Wechsel ins Perfekt mitten im Text („Dann habe ich gewonnen“), Präteritum regelmäßig gebildet („werfte“, „fallte“), Plusquamperfekt mit habe statt hatte."] }],
]));

// ---------- Phase 6 ----------
doc.push(phase("6", "Partner-Check", 8));
doc.push(...steps([
  ["partner", { say: "Jetzt arbeitet ihr zu zweit. Tauscht die Hefte. Ihr prüft den Bericht eures Partners. Die Fragen stehen rechts an der Tafel.", board: auftrag("Partner-Check", ["Unterstreiche alle Präteritum-Verben.", "Markiere im Plusquamperfekt hatte/war rot und das Partizip II gelb.", "Schreibe unter den Text: die Antworten auf die drei Fragen (ja/nein) und einen Satz „Gut gelungen ist …“. Gib das Heft zurück."], [
    bl("Steht der Text durchgehend im Präteritum?   ja / nein"),
    bl("Gibt es mindestens einen Satz zur Vorgeschichte im Plusquamperfekt?   ja / nein"),
    bl("Stimmen die Formen? Wenn nicht: Wo?", { after: 0 }),
  ], "6 Minuten · zu zweit · Flüsterstimme"), do: ["Der Partner markiert mit Bleistift bzw. Rot/Gelb, verbessert aber nicht selbst – verbessert wird vom Verfasser mit Grün."] }],
  ["kontrolle", { say: "Heft zurück. Lest die Rückmeldung und verbessert mit Grün." }],
]));

// ---------- Phase 7 ----------
doc.push(phase("7", "Reflexion: Wann nehme ich welche Zeitform?", 5));
doc.push(...steps([
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, trägt eine Zeitform in die Tabelle ein. Alle anderen schreiben die Tabelle ins Heft.", board: board("Mitte", [table([3900, 3900], [
    ["Situation", "Zeitform"], ["Ich erzähle mündlich, im Chat", "Perfekt"], ["Ich schreibe einen Bericht, eine Erzählung", "Präteritum"], ["Ich erzähle, was vorher passiert war", "Plusquamperfekt"], ["Ich sage, was ich vorhabe", "Futur I"],
  ].map((r, i) => row(r.map((v, j) => cell(p(i === 0 ? N(v, { bold: true, color: NAVY }) : j === 1 ? N(v, { color: MUTED, italics: true }) : n19(v), { after: 0 }), { w: 3900, borders: allBorders(solid("9AA59C", 4)), m: 20 })))))]),
    do: ["Die rechte Spalte leer anschreiben, die Kreide-Kette trägt die Zeitformen ein.", "Abschlussfrage: „Was hat sich verändert, als ihr euren Wettkampf aufgeschrieben habt, statt ihn zu erzählen?“"] }],
]));
doc.push(p([S("Ausblick: ", { bold: true }), S("Noch offen: Futur II, Hilfsverb oder Vollverb, Modalverben (verschoben) und am Schluss der Einheit die Satzklammer.")], { before: 120, after: 0 }));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const dScript = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Präteritum und Plusquamperfekt") }, footers: { default: footer("Tafelskript") }, children: doc }] });
Packer.toBuffer(dScript).then((b) => fs.writeFileSync("Tafelskript_Praeteritum_Plusquamperfekt.docx", b));
