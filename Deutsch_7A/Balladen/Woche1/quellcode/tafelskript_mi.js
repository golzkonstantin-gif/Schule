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
// Tafelskript Mi 14.10. (Doppelstunde): Mitarbeit · John Maynard hören · Lesetechniken – AVIVA
// ================================================================
const merkTafel = (nr, title, children) => board(`Seitentafel · Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  ...children,
]);
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const zielTafel = (ziel, ablauf) => board("Seitentafel links · bleibt die ganze Stunde stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ziel, { size: 21, italics: true })], { after: 60 }),
  p(t("Ablauf", { bold: true, size: 20, color: NAVY }), { after: 20 }),
  ...ablauf.map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]);

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Mi 14.10. · Doppelstunde (90 Min.) · Planungsmodell AVIVA"), h1("John Maynard – hören und lesen"));
doc.push(box([
  p([S("Bezug: ", { bold: true }), S("Start der Balladen-Einheit (Grobplanung Std. 1–2). Aus der Rückmeldung zum Unterrichtsbesuch: Ziel und Ablauf stehen transparent an der Tafel, und in jeder Phase haben alle einen eigenen Auftrag – an der Tafel wird nur verglichen. Morgen schreibt die Klasse KA 1; die Stunde ist bewusst ruhig und klar getaktet.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S("Ich kann beim Zuhören Notizen zu einer Ballade machen und je nach Leseziel überfliegend, orientierend oder genau lesen.", { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Es gibt drei Lesetechniken: überfliegen, um zu wissen, worum es geht – orientierend lesen, um eine bestimmte Information zu finden – genau lesen, um alles zu verstehen. Ich wähle sie nach meinem Leseziel.“")], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 1300, 3400, 4606], ["Zeit", "AVIVA", "Phase", "Alle Schüler …"], [
  ["20′", "–", "Mitarbeit einschätzen", "füllen den Selbsteinschätzungsbogen aus"],
  ["12′", "A", "Ankommen: die Ballade hören", "notieren beim Hören im Notizraster"],
  ["7′", "V", "Vorwissen: Wie lest ihr?", "ordnen drei Alltagssituationen eine Leseweise zu"],
  ["8′", "I", "Merksatz 1: Lesetechniken", "schreiben ab"],
  ["30′", "V", "Ü1 Drei Arten zu lesen", "überfliegen, suchen, genau lesen – allein, dann zu zweit"],
  ["8′", "A", "Auswerten: Ziel erreicht?", "lösen drei Situationen, schreiben „Heute habe ich gelernt, …“"],
], { bold: [true, true, true, false], colors: [C.hv, C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("5 Minuten. Mit * markiert (kann entfallen): Ü1 c) Teil 2 (Wörter aus dem Kontext) und die Zusatzfrage 5 in Ü1 b).")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Einheit", "Leute, höret die Geschicht’ – Balladen (neue Seite)"],
  ["Überschrift", "John Maynard – hören und lesen · Datum rechts"],
  ["Schnipsel", "Notizraster „John Maynard hören“ (eingeklebt)"],
  ["Heft", "Wie lest ihr? – drei Situationen mit Leseweise"],
  ["Merksatz 1", "Lesetechniken (rot umrahmt)"],
  ["Ü1", "Schnipsel „Drei Arten zu lesen“ (eingeklebt und bearbeitet)"],
  ["Abschluss", "„Heute habe ich gelernt, …“"],
].map(([a, b2]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" || a === "Einheit" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b2, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Lehrbuch (John Maynard) · Hörfassung (z. B. Achim Reichel; sonst liest die Lehrkraft vor) · Selbsteinschätzungsbogen Mitarbeit · Druckvorlage: Notizraster (2 pro A4), Ü1 (1 pro A4, viel Schreibplatz).")], { before: 100, after: 40 }));
doc.push(p([S("Vorbereitung: ", { bold: true }), S("Ziel und Ablauf links an die Seitentafel, Merksatz 1 rechts auf die Seitentafel (zugeklappt). Seite und Zeilen von John Maynard im Lehrbuch nachschlagen und in Ü1 b) die Strophen ggf. durch Zeilen ersetzen.")], { after: 60 }));
doc.push(zielTafel("Ich kann beim Zuhören Notizen machen und je nach Leseziel überfliegend, orientierend oder genau lesen.", [
  "Mitarbeit einschätzen", "Eine Ballade hören", "Wie lest ihr?", "Merksatz 1: Lesetechniken", "Ü1 Drei Arten zu lesen (c Teil 2 *)", "Was habe ich heute gelernt?",
]));

// ---------- Mitarbeit ----------
doc.push(avivaPhase("0", "Mitarbeit einschätzen", 20));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Bevor wir mit dem neuen Thema anfangen, schaut ihr auf eure Mitarbeit seit Schuljahresbeginn. Ihr bekommt einen Bogen. Kreuzt ehrlich an – nur die linke Seite. Die rechte Seite fülle ich aus. Am Freitag bekommt ihr den Bogen mit meiner Einschätzung und eurer Note zurück.", do: ["Die Kriterien einmal laut vorlesen und kurz erklären, was z. B. „fachlich richtig, mit Fachbegriffen“ bedeutet."] }],
  ["luecken", { say: "Jetzt arbeitet ihr allein. Zwölf Minuten. Jeder für sich, ohne Absprache. Schreibt auch auf, was euch gut gelingt und woran ihr arbeiten wollt. Fertig? Bogen umdrehen.", do: ["Bögen einsammeln. Wer früher fertig ist, liest im Lehrbuch nur die Überschrift der Ballade und schreibt eine Vermutung ins Heft: „Wer ist John Maynard?“"] }],
]));

// ---------- A ----------
doc.push(avivaPhase("A", "Ankommen und einstimmen: die Ballade hören", 12));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ab heute geht es um Balladen. An der Tafel steht unser Ziel für heute und der Ablauf. Neue Heftseite: Leute, höret die Geschicht’ – Balladen. Darunter: John Maynard – hören und lesen, Datum rechts.", do: ["Auf Ziel und Ablauf an der Seitentafel zeigen und beides einmal vorlesen."] }],
  ["luecken", { say: "Ihr bekommt einen Schnipsel. Klebt ihn ein. Ihr hört jetzt eine Ballade. Beim ersten Hören notiert ihr nur Stichpunkte: Wer? Wo? Was passiert? Wie endet es?", do: ["Hörfassung einmal abspielen (ca. 3–4 Min.). Alle schreiben beim Hören."] }],
  ["partner", { say: "Vergleicht zu zweit eure Notizen. Ergänzt, was euch fehlt. Zwei Minuten. Dann hört ihr noch einmal und beantwortet die letzte Frage: Wer ist John Maynard?", do: ["Zweites Hören. Danach zwei Schüler: „Wer ist John Maynard?“ – erwartete Antwort: der Steuermann, der das brennende Schiff ans Ufer steuert, alle rettet und selbst stirbt."], sol: "Wer: John Maynard (Steuermann), Kapitän, Passagiere · Wo: auf dem Schiff „Schwalbe“ auf dem Eriesee, auf dem Weg nach Buffalo · Was passiert: Feuer bricht aus, Panik, Maynard hält das Steuer und fährt ans Ufer · Wie endet es: alle gerettet, nur Maynard stirbt; die Stadt ehrt ihn." }],
]));

// ---------- V ----------
doc.push(avivaPhase("V", "Vorwissen aktivieren: Wie lest ihr?", 7));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. An der Tafel stehen drei Situationen. Schreibt ins Heft zu jeder Situation: Wie lest ihr – schnell drüber, gezielt suchen oder langsam Wort für Wort? Zwei Minuten.", board: board("Mitte", [
    bl("A  Ihr bekommt die Ausschreibung für ein Turnier und sucht nur eure Startzeit."),
    bl("B  Ihr blättert in einem Buch und wollt wissen, worum es in einem Kapitel geht."),
    bl("C  Ihr bekommt eine neue Regel für euren Sport, die ihr genau verstehen müsst.", { after: 0 }),
  ]) }],
  ["partner", { say: "Vergleicht zu zweit. Gibt es Unterschiede? Dann sammeln wir.", do: ["Kurz im Plenum: Jede Situation verlangt eine andere Leseweise. Überleitung: „Diese drei Leseweisen haben Namen – das sind die Lesetechniken.“"], sol: "A gezielt suchen (orientierend) · B schnell drüber (überfliegend) · C langsam, Wort für Wort (genau)" }],
]));

// ---------- I ----------
doc.push(avivaPhase("I", "Informieren: Merksatz 1 Lesetechniken", 8));
doc.push(...steps([
  ["abschreiben", { say: "Jetzt schreibt ihr von der Seitentafel ab: Merksatz 1. Rot umrahmen.", board: merkTafel(1, "Lesetechniken", [
    p([S("Je nach "), S("Leseziel", { bold: true }), S(" lese ich unterschiedlich:")], { after: 40 }),
    p([S("überfliegend: ", { bold: true }), S("schnell über den Text (Überschrift, Anfang, Ende, auffällige Wörter) – ich will wissen, "), S("worum es geht", { bold: true }), S(".")], { after: 30 }),
    p([S("orientierend: ", { bold: true }), S("gezielt suchen, den Rest überspringen – ich will "), S("eine bestimmte Information", { bold: true }), S(" finden.")], { after: 30 }),
    p([S("genau: ", { bold: true }), S("Satz für Satz, Unbekanntes klären, Randnotizen machen – ich will "), S("alles verstehen", { bold: true }), S(".")], { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Ich spare Zeit und finde, was ich brauche – die Startzeit in der Ausschreibung genauso wie die wichtige Stelle in einer Ballade. "), S("Test: ", { bold: true }), S("Frag dich vor dem Lesen: Will ich einen Überblick, eine Information oder alles verstehen?")], { after: 0 }),
  ]), do: ["Rückbezug: A = orientierend, B = überfliegend, C = genau. Die Schüler ergänzen die Fachbegriffe in ihren Heftnotizen aus der V-Phase."] }],
]));

// ---------- V ----------
doc.push(avivaPhase("V", "Verarbeiten: Ü1 Drei Arten zu lesen", 30));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Klebt den Schnipsel Ü1 ein. Lehrbuch auf Seite … Wir machen die drei Lesetechniken nacheinander, jeweils mit Zeit. Teil a: überfliegen. Ihr habt genau eine Minute. Dann Buch zu und einen Satz schreiben.", board: auftrag("Ü1 Drei Arten zu lesen", ["Überfliege die Ballade (1 Minute). Schreibe einen Satz: „In der Ballade geht es um …“", "Lies orientierend und beantworte die Fragen. Notiere die Strophe.", "Lies Strophe 5 und 6 genau. Notiere Stichpunkte wie Randnotizen und kläre die Wörter aus dem Kontext."], [p(S("Die Fragen und Wörter stehen auf dem Schnipsel.", { size: 20, italics: true }), { after: 0 })], "a) 1 Min. + 2 Min. · b) 8 Min. · c) 10 Min. · allein · leise"),
    do: ["Zeit für a) laut stoppen: „Start … Stopp. Buch zu.“ Das Überfliegen funktioniert nur mit echter Zeitgrenze."] }],
  ["luecken", { say: "Teil b: orientierend lesen. Ihr sucht nur die Antworten – nicht alles lesen! Acht Minuten. Wer fertig ist, macht Frage 5.", sol: "1 „Schwalbe“ (Str. 2) · 2 von Detroit nach Buffalo, über den Eriesee (Str. 2) · 3 zwanzig Minuten (Str. 3) · 4 „Ja, Herr, ich halt’s!“ (Str. 6) · *5 zehntausend oder mehr (Str. 8)" }],
  ["luecken", { say: "Teil c: genau lesen. Strophe 5 und 6, Satz für Satz. Notiert auf dem Schnipsel zu jeder Strophe Stichpunkte, was dort passiert – wie Randnotizen, nur nicht im Buch. Dann klärt ihr die Wörter: Was vermutet ihr – und welcher Hinweis im Text hilft euch?", sol: "Stichpunkte z. B.: Kapitän fragt – Maynard antwortet – Befehl: auf den Strand – Maynard hält aus – steuert in die Brandung. Wörter: Sprachrohr = Rohr, durch das man über Entfernung spricht (Kapitän „fragt an“, sieht ihn nicht) · Brandung = Wellen, die am Ufer brechen („auf den Strand! In die Brandung!“) · Bugspriet = vorderster Teil des Schiffs („vorn ist noch Luft und Licht“, Str. 4).", do: ["Die Wörter (Teil 2) sind mit * markiert und können bei Zeitmangel entfallen."] }],
  ["partner", { say: "Vergleicht zu zweit a, b und c. Verbessert mit Grün. Fünf Minuten.", do: ["Herumgehen; bei b) nachfragen: „Wie hast du die Stelle so schnell gefunden?“ – Ziel: Schlüsselwörter suchen (Feuer, Minuten, Grab)."] }],
  ["tafel", { say: "Wer drankommt, schreibt die Antwort zu einer Frage aus b) an die Tafel – mit Strophe. Alle prüfen mit.", do: ["Nur b) an der Tafel, kurz: 4 Schüler, je eine Frage. Die übrigen haken in ihrem Schnipsel ab und verbessern mit Grün."] }],
]));

// ---------- A ----------
doc.push(avivaPhase("A", "Auswerten: Ziel erreicht?", 8));
doc.push(...steps([
  ["luecken", { say: "Stift in die Hand. Drei kurze Situationen an der Tafel. Schreibt ins Heft: Welche Lesetechnik passt? Eine Minute.", board: board("Mitte", [
    bl("1  Du willst wissen, ob ein Zeitungsartikel über deinen Verein berichtet."),
    bl("2  Du suchst im Lehrbuch, in welcher Strophe das Feuer ausbricht."),
    bl("3  Du sollst morgen erklären, warum John Maynard ein Held ist.", { after: 0 }),
  ]), sol: "1 überfliegend · 2 orientierend · 3 genau" }],
  ["gemeinsam", { say: "Schaut auf unser Ziel an der Tafel. Habt ihr es erreicht? Schreibt einen Satz ins Heft: Heute habe ich gelernt, …", do: ["Zwei Schüler lesen ihren Satz vor. Erwartet: die drei Lesetechniken und wann man sie benutzt.", "Ausblick: „Am Freitag lesen wir die Ballade genau und gliedern sie.“"] }],
]));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Mi 14.10.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Mi_14-10_Hoeren_Lesetechniken.docx", buf));
