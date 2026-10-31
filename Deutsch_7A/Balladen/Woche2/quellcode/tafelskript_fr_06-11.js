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



// Tafelskript Fr 6.11. (Einzelstunde): Wiederholung Reime · Die Ballade als Ur-Ei – AVIVA, mit Präsentation
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const ZIEL = "Ich kann an Paddy Fingal mit Belegen zeigen, dass eine Ballade Merkmale von Lyrik, Epik und Dramatik hat.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Fr 6.11. · Einzelstunde (45 Min.) · AVIVA · mit Präsentation (13 Folien)"), h1("Paddy Fingal – die Ballade als „Ur-Ei“"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Am Donnerstag hat die Klasse Reime und Reimschema an Strophe 1 entdeckt und an Texten aus der eigenen Sportart geübt. Heute wird das zuerst wiederholt (neue Strophe 11, Fachbegriffe). Dann kommt das „Ur-Ei“: Die Klasse lernt, dass eine Ballade Merkmale von Gedicht, Erzählung und Theater verbindet. Vergleiche aus ihrer Lebenswelt (Song-Text, Roman oder Spielbericht, Serienszene) sorgen für Abwechslung und Transfer. Danach weisen die Schüler die Merkmale mit dem Ur-Ei-Schnipsel an Paddy Fingal nach. Das Belegen kennen sie seit Mittwoch, die Lyrik-Merkmale (Nr. 1–3) seit Donnerstag.")], { after: 0 })], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S(ZIEL, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Eine Ballade ist wie ein Gedicht (Strophen, Reime, Kehrreim), wie eine Erzählung (Erzähler, Figuren, Handlung im Präteritum) und wie ein Theaterstück (wörtliche Rede, Spannung). Das habe ich an Paddy Fingal mit Strophen belegt.“ Überprüft wird das im Zielcheck.")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ZIEL, { size: 21, italics: true })], { after: 60 }),
  ...["Wiederholung: Ü1 Reimschema und Fachbegriffe", "Die Ballade als „Ur-Ei“", "Ü2 Merkmale an Paddy Fingal nachweisen (* Nr. 8)", "Zielcheck"].map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["2′", "A", "1–2", "Ankommen: Ziel und Ablauf", "lesen Ziel und Ablauf, schreiben die Überschrift"],
  ["7′", "V", "3–4", "Ü1 Reimschema Str. 11, Fachbegriffe", "lösen allein, verbessern mit der Lösungsfolie"],
  ["8′", "I", "5–7", "Die Ballade als „Ur-Ei“, Vergleich Song – Spielbericht – Serienszene", "hören zu, nennen Merkmale, vergleichen"],
  ["12′", "V", "8", "Ü2 Ur-Ei-Schnipsel", "weisen allein Merkmale nach, notieren Strophen"],
  ["6′", "A", "9–11", "Ü2 gemeinsam prüfen (je Gattung eine Folie)", "nennen Lösung mit Strophe, verbessern mit Grün"],
  ["6′", "A", "12–13", "Zielcheck", "ordnen drei Stellen einer Gattung zu und begründen; Lösungsfolie"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("41 Min., 4 Min. Puffer. Wenn es eng wird: * Nr. 8 (Spannung) beim Aufdecken überspringen.")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Die Ballade als „Ur-Ei“ · Datum rechts"],
  ["Wiederholung", "Ü1: Reimschema Str. 11 (aabbx, Paarreim) · drei Fachbegriffe in je einem Satz"],
  ["Ur-Ei", "Schnipsel „Die Ballade – ein Ur-Ei“ eingeklebt und ausgefüllt, mit Grün kontrolliert"],
  ["Zielcheck", "A, B, C mit Gattung und Begründung"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation `Freitag_06-11_Deutsch_7A.pptx` (Lösungen in den Notizen und auf Aufdeck-Folien) · Lehrbuch S. 136 f. · Schnipsel „Die Ballade – ein Ur-Ei“ aus `Woche1/` (unverändert, 2 pro A4) · Handreichung Polnisch für S-PL.")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 2));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Am Donnerstag ging es darum, wie die Ballade klingt. Heute zeigen wir: Eine Ballade ist mehr als ein Gedicht. Ziel und Ablauf stehen auf der Folie und links an der Tafel. Jetzt schreibt ihr die Überschrift ab: Die Ballade als Ur-Ei." }]]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Ü1 Reimschema und Fachbegriffe", 7));
doc.push(...steps([
  ["luecken", { say: "Wiederholung vom Donnerstag. Jetzt arbeitet ihr allein. a: Bestimmt das Reimschema von Strophe 11 und nennt die Reimart. b: Erklärt in je einem Satz: Strophe, Vers, Kehrreim. Vier Minuten. Fertig? Stift hinlegen.", sol: "a Siegesgeschrei – herbei (a a), an – dann (b b), Paddy Fingal (x) → Paarreim + Kehrreim · b Strophe: Abschnitt, Leerzeile dazwischen · Vers: eine Zeile · Kehrreim: Vers, der sich in jeder Strophe wiederholt" }],
  ["kontrolle", { say: "Vergleicht mit der Lösung auf der Folie und verbessert mit Grün.", do: ["Folie 4. Gezielt die Schüler drannehmen, die im Zielcheck am Donnerstag Fehler hatten."] }],
]));

doc.push(avivaPhase("I", "Informieren: Die Ballade als „Ur-Ei“", 8));
doc.push(board("Mitteltafel · Tafelbild (optional zusätzlich zur Folie, bleibt bis zum Zielcheck stehen)", [
  p([N("Die Ballade – ein „Ur-Ei“ (Goethe)", { bold: true, color: NAVY })], { after: 60, align: AlignmentType.CENTER }),
  bl([N("Lyrik", { bold: true, color: "2A8C8C" }), N(" – wie ein Song-Text: Strophen, Verse, Reime, Kehrreim")]),
  bl([N("Epik", { bold: true, color: "7A5BA6" }), N(" – wie ein Roman oder Spielbericht: Erzähler, Figuren, Handlung, Präteritum")]),
  bl([N("Dramatik", { bold: true, color: "C77C12" }), N(" – wie eine Serienszene: Figuren sprechen selbst, Konflikt, Spannung")]),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Goethe hat die Ballade ein Ur-Ei genannt. In einem Ei steckt schon alles, was später daraus wird. In der Ballade stecken alle drei großen Gattungen: Lyrik, Epik und Dramatik. Das braucht ihr für die erste Bewertung, die Analyse.", do: ["Folie 5 (Ich)."] }],
  ["gemeinsam", { say: "Woran erkennt ihr einen Song? Einen Spielbericht? Eine Serienszene?", sol: "Song: Strophen, Refrain, Reime · Spielbericht: jemand erzählt, was passiert ist, in der Vergangenheit · Serienszene: Figuren reden miteinander, Streit, man will wissen, wie es weitergeht", do: ["Folie 6 (Wir): Merkmale sammeln, dann Folie 7 aufdecken und vergleichen. Wer drankommt, nennt ein Merkmal – möglichst viele verschiedene Schüler drannehmen."] }],
]));

doc.push(avivaPhase("V", "Verarbeiten: Ü2 Merkmale nachweisen (Ur-Ei-Schnipsel)", 12));
doc.push(...steps([["luecken", { say: "Ich teile den Schnipsel aus. Klebt ihn ein. Jetzt arbeitet ihr allein. Nummer 1 bis 3 kennt ihr schon. Bei 4 bis 7 weist ihr die Merkmale an Paddy Fingal nach und notiert die Strophe. Zwölf Minuten. Wer fertig ist, macht Nummer 8 mit Stern. Fertig? Stift hinlegen.", sol: "1 11 Strophen mit je 5 Versen · 2 Paarreim (aabb), z. B. Riesenart / Bart (Str. 1) · 3 „… Paddy Fingal“ (Kehrreim) · 4 Paddy Fingal, Schaya, der Schotte (*die Nachbarn) · 5 Anfang Str. 1–3, Höhepunkt Str. 8–9, Ende Str. 10–11 (andere begründete Grenzen gelten) · 6 Präteritum, z. B. „Paddy Fingal war von Riesenart“ (Str. 1) · 7 Paddy (Str. 4), der Schotte (Str. 6, 9), Schaya (Str. 7) · *8 Paddy gegen den Schotten; Frage: Findet der Schotte Paddy?", do: ["Folie 8. Herumgehen. Typische Hürde: Höhepunkt (Nr. 5) – die Stelle, an der sich alles entscheidet (der Schotte steht vor dem Bett).", "Bei Nr. 6 Bezug zur Verb-Einheit: Woran erkennst du das Präteritum? (war, kroch, lief)"] }]]));

doc.push(avivaPhase("A", "Auswerten: Ü2 gemeinsam prüfen", 6));
doc.push(...steps([["gemeinsam", { say: "Wir prüfen gemeinsam, Gattung für Gattung. Wer drankommt, nennt seine Lösung mit Strophe. Dann decke ich auf. Alle verbessern mit Grün.", do: ["Folie 9 Lyrik, Folie 10 Epik, Folie 11 Dramatik. Verschiedene Schüler drannehmen, niemand geht an die Tafel."] }]]));

doc.push(avivaPhase("A", "Auswerten: Zielcheck", 6));
doc.push(...steps([
  ["luecken", { say: "Zielcheck. Jetzt arbeitet ihr allein im Heft. Ordnet jede Stelle einer Gattung zu und begründet mit dem Merkmal. Vier Minuten.", sol: "A Lyrik – Reim (Riesenart/Bart, Paarreim) · B Epik – ein Erzähler erzählt im Präteritum („lief“) · C Dramatik – wörtliche Rede von Paddy" }],
  ["kontrolle", { say: "Vergleicht mit der Lösung auf der Folie und verbessert mit Grün.", do: ["Folie 13. Herumgehen und schauen, wer noch Fehler hat – am Montag in der Wiederholung aufgreifen. Ausblick: Am Montag schreibt ihr selbst Reime."] }],
]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(new Paragraph({ pageBreakBefore: true, children: [t("FÜR DIE LEHRKRAFT · REFLEXION VOR UND NACH DER STUNDE", { size: 18, color: MUTED, bold: true })], spacing: { after: 20 } }), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Ziel und Ablauf auf Folie und Seitentafel, Rückgriff im Zielcheck."));
doc.push(bul("Vorwissen aktivieren:", "Echte Wiederholung vom Donnerstag an einer neuen Strophe (Transfer statt Abschreiben), dazu die Fachbegriffe."));
doc.push(bul("Informieren:", "Kurz und anschaulich: Ur-Ei-Bild, dann Vergleiche aus der Lebenswelt (Song, Spielbericht, Serienszene) – von der Lehrkraft ausdrücklich gewünscht für Abwechslung und Transfer."));
doc.push(bul("Verarbeiten und Auswerten:", "Alle arbeiten allein am Schnipsel, Auswertung über Aufdeck-Folien je Gattung, niemand muss an die Tafel. Sicherung als Zielcheck mit neuen Textstellen."));
doc.push(bul("Zeit:", "Ur-Ei bewusst auf eine eigene Stunde gelegt; 4 Min. Puffer, * Nr. 8 als Kürzungsoption."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Ein überprüfbares Ziel; Relevanz: Die Merkmale braucht man für die Analyse (Bewertung 1)."));
doc.push(bul("Ich – Wir – Ihr:", "Lehrkraft erklärt das Ur-Ei (Ich), Klasse sammelt Merkmale an Alltagsbeispielen (Wir), Schüler weisen sie allein an der Ballade nach (Ihr)."));
doc.push(bul("Rückmeldung:", "Lösungsfolie nach Ü1, Aufdeck-Folien nach Ü2, Zielcheck mit Lösung; typische Hürde benannt: Höhepunkt."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "Nr. 1–3 als sicherer Einstieg (bekannt), * Nr. 8 für Schnelle, Bezug zur Verb-Einheit bei Nr. 6."));
doc.push(bul("", "S-PL: Handreichung mit Ziel, Ablauf und Aufträgen auf Polnisch; Fachbegriffe (Lyrik, Epik, Dramatik …) bleiben deutsch, Polnisch erklärt nur die Bedeutung."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Sitzen Reimschema und Fachbegriffe nach einem Tag (Ü1)?"));
doc.push(bul("", "Tragen die Vergleiche aus der Lebenswelt – nennen die Schüler selbst Merkmale?"));
doc.push(bul("", "Gelingt das Belegen bei Epik und Dramatik (Nr. 4–7)? Wo hakt es (Höhepunkt, Präteritum)?"));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Fr 6.11.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Fr_06-11_Ur-Ei.docx", buf));
