const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Lesetechniken";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Paddy Fingal – Lesetechniken · Deutsch 7A";
const TECH = { ue: { name: "überfliegend", col: BLUE, sym: "»" }, or: { name: "orientierend", col: YEL, sym: "⌕" }, ge: { name: "genau", col: GREEN, sym: "✎" } };
let pageNo = 1;

function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 16, color: NAVY, margin: 0, valign: "top" }, o));
}
function base(kicker, title, notes) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.4, w: 9, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 2 });
  txt(s, title, { x: 0.6, y: 0.72, w: 12.1, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  txt(s, FOOT, { x: 0.6, y: 7.0, w: 7, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo), { x: 12.2, y: 7.0, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
  if (notes) s.addNotes(notes);
  return s;
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
}
function darkBg(s) {
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
}
function num(s, n, x, y, col = NAVY, d = 0.5) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: d * 32, bold: true, color: WHITE, align: "center", valign: "middle" });
}
// Übung im festen Format: Nummer + Titel, nummerierte Schritte, Zeit, Arbeitsform, „Fertig? Stift hinlegen“
function uebung(kicker, title, steps, meta, notes, extra) {
  const s = base(kicker, title, notes);
  steps.forEach((st, i) => {
    const y = 1.75 + i * 0.8;
    num(s, i + 1, 0.6, y);
    txt(s, st, { x: 1.3, y: y + 0.04, w: 11.3, h: 0.7, fontSize: 21, bold: true });
  });
  if (extra) extra(s, 1.75 + steps.length * 0.8 + 0.1);
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
  return s;
}

// ============ 1 Titel ============
{
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 2.6, 0.42, YEL, { r: 0.21 });
  txt(s, "BALLADEN", { x: 0.9, y: 2.05, w: 2.6, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Paddy Fingal – Lesetechniken", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, "Leute, höret die Geschicht’ – Wilhelm Brandes: Paddy Fingal (Lehrbuch S. 136 f.)", { x: 0.9, y: 3.7, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
  ["ue", "or", "ge"].forEach((k, i) => {
    box(s, 0.9 + i * 2.9, 4.7, 2.7, 0.6, TECH[k].col, { r: 0.3 });
    txt(s, TECH[k].name, { x: 0.9 + i * 2.9, y: 4.7, w: 2.7, h: 0.6, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}

// ============ 2 Ziel und Ablauf ============
{
  const s = base("Heute", "Ziel und Ablauf", "Ziel einmal vorlesen. Die Folie bleibt am Ende für die Auswertung wichtig (Folie „Ziel erreicht?“).");
  box(s, 0.6, 1.7, 12.1, 1.2, LIGHT2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 0.12, h: 1.2, fill: { color: RED }, line: { type: "none" } });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: "Ich kann je nach Leseziel überfliegend, orientierend oder genau lesen und meinen ersten Eindruck von einer Ballade begründen.", options: { italic: true } }], { x: 1.0, y: 1.9, w: 11.5, h: 0.9, fontSize: 22 });
  const ab = ["Wer ist Paddy Fingal?", "Wie lest ihr?", "Merksatz 1: Lesetechniken", "Ü1 Paddy Fingal lesen (Frage 5 *)", "Was habe ich heute gelernt?  ·  * LB-Aufgabe 1"];
  ab.forEach((a, i) => {
    num(s, i + 1, 0.9, 3.25 + i * 0.68, NAVY, 0.45);
    txt(s, a, { x: 1.55, y: 3.28 + i * 0.68, w: 10, h: 0.5, fontSize: 20 });
  });
}

// ============ 3 Ankommen ============
{
  const s = base("Ankommen", "Wer ist Paddy Fingal?", "Zwei Vermutungen vorlesen lassen (Riese, Seemann, Raufbold …). Noch nicht lesen – nur Titel und erstes Bild.");
  box(s, 0.6, 1.8, 12.1, 2.2, LIGHT);
  txt(s, "Paddy Fingal", { x: 0.6, y: 2.05, w: 12.1, h: 0.9, fontFace: HEAD, fontSize: 54, bold: true, color: RED, align: "center" });
  txt(s, "Wilhelm Brandes · Lehrbuch S. 136", { x: 0.6, y: 3.05, w: 12.1, h: 0.5, fontSize: 20, italic: true, color: MUTED, align: "center" });
  txt(s, [{ text: "Schau nur auf den Titel und das erste Bild. Schreibe einen Satz ins Heft: ", options: {} }, { text: "Wer ist Paddy Fingal?", options: { bold: true } }], { x: 0.6, y: 4.5, w: 12.1, h: 0.8, fontSize: 24, align: "center" });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "2 Minuten · allein      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}

// ============ 4 Wie lest ihr? ============
{
  const s = base("Vorwissen", "Wie lest ihr?", "Lösung: A gezielt suchen (orientierend) · B schnell drüber (überfliegend) · C langsam, Wort für Wort (genau). Überleitung: Diese drei Leseweisen haben Namen – die Lesetechniken.");
  txt(s, "Schreibe ins Heft zu jeder Situation: Wie liest du – schnell drüber, gezielt suchen oder langsam Wort für Wort?", { x: 0.6, y: 1.55, w: 12.1, h: 0.5, fontSize: 18, color: MUTED });
  const sit = [["A", "Du bekommst die Ausschreibung für ein Turnier und suchst nur deine Startzeit."], ["B", "Du blätterst in einem Buch und willst wissen, worum es in einem Kapitel geht."], ["C", "Du bekommst eine neue Regel für deinen Sport, die du genau verstehen musst."]];
  sit.forEach(([k, d], i) => {
    const y = 2.3 + i * 1.3;
    box(s, 0.6, y, 12.1, 1.1, LIGHT);
    txt(s, k, { x: 0.9, y: y + 0.18, w: 0.8, h: 0.7, fontFace: HEAD, fontSize: 36, bold: true, color: RED });
    txt(s, d, { x: 1.8, y: y + 0.3, w: 10.6, h: 0.6, fontSize: 22 });
  });
  txt(s, "2 Minuten · allein, dann zu zweit vergleichen", { x: 0.6, y: 6.35, w: 12.1, h: 0.4, fontSize: 16, bold: true, color: MUTED });
}

// ============ 5 Merksatz 1 ============
{
  const s = base("Merksatz 1 – abschreiben, rot umrahmen", "Lesetechniken", "Rückbezug: A = orientierend, B = überfliegend, C = genau. Die Schüler ergänzen die Fachbegriffe in ihren Notizen von vorhin.");
  box(s, 0.6, 1.6, 12.1, 4.15, WHITE, { line: RED, lw: 3, r: 0.05 });
  txt(s, [{ text: "Je nach " }, { text: "Leseziel", options: { bold: true } }, { text: " lese ich unterschiedlich:" }], { x: 0.95, y: 1.8, w: 11.4, h: 0.45, fontSize: 22 });
  const rows = [
    ["ue", "schnell über den Text (Überschrift, Anfang, Ende, auffällige Wörter)", "– ich will wissen, ", "worum es geht", "."],
    ["or", "gezielt suchen, den Rest überspringen", "– ich will ", "eine bestimmte Information", " finden."],
    ["ge", "Satz für Satz, Unbekanntes klären, Randnotizen machen", "– ich will ", "alles verstehen", "."],
  ];
  rows.forEach(([k, a, b, c, d], i) => {
    const y = 2.45 + i * 1.05;
    box(s, 0.95, y, 2.6, 0.75, TECH[k].col, { r: 0.37 });
    txt(s, TECH[k].sym + "  " + TECH[k].name, { x: 0.95, y, w: 2.6, h: 0.75, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [{ text: a + " " }, { text: b }, { text: c, options: { bold: true, color: TECH[k].col } }, { text: d }], { x: 3.8, y: y + 0.05, w: 8.6, h: 0.75, fontSize: 19, valign: "middle" });
  });
  box(s, 0.6, 5.95, 5.95, 0.95, LIGHT2);
  txt(s, [{ text: "Wofür? ", options: { bold: true } }, { text: "Ich spare Zeit und finde, was ich brauche – die Startzeit in der Ausschreibung genauso wie die wichtige Stelle in einer Ballade." }], { x: 0.8, y: 6.02, w: 5.6, h: 0.85, fontSize: 14 });
  box(s, 6.75, 5.95, 5.95, 0.95, LIGHT2);
  txt(s, [{ text: "Test: ", options: { bold: true } }, { text: "Frag dich vor dem Lesen: Will ich einen Überblick, eine Information oder alles verstehen?" }], { x: 6.95, y: 6.02, w: 5.6, h: 0.85, fontSize: 14 });
}

// ============ 6–8 Ü1 ============
uebung("Verarbeiten · Ü1 a", "Ü1 a) Überfliegen", [
  "Schreibe ins Heft: Ü1 Paddy Fingal lesen – a). Schlage S. 136 f. auf.",
  "Überfliege die Ballade – genau 1 Minute für beide Seiten.",
  "Klappe das Buch zu. Schreibe einen Satz: „In der Ballade geht es um …“",
], "1 Minute lesen + 2 Minuten schreiben · allein · leise",
"Zeit laut stoppen: „Start … Stopp. Buch zu.“ Lösung z. B.: In der Ballade geht es um den Riesen Paddy Fingal und einen noch größeren Schotten, vor dem er sich versteckt.",
(s, y) => {
  box(s, 0.6, y + 0.2, 12.1, 1.3, LIGHT);
  txt(s, "⏱  1 Minute", { x: 0.6, y: y + 0.35, w: 12.1, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: BLUE, align: "center", valign: "middle" });
});

// ============ Zuhören, Kehrreim, erster Eindruck ============
{
  const s = base("Verarbeiten · Zuhören", "Zuhören und mitsprechen", "Methodenkarte „Chorisches Lesen“: Lehrkraft liest die Strophen, die Klasse spricht die letzte Zeile jeder Strophe im Chor. Vorher Tempo und Lautstärke festlegen. Danach schreiben alle den ersten Eindruck in Ü1; zwei lesen vor.");
  ["Hör zu, während die Ballade vorgelesen wird. Lies im Buch mit.", "Die letzte Zeile jeder Strophe sprecht ihr alle im Chor mit.", "Schreibe ins Heft: Die Ballade wirkt lustig / spannend / unheimlich, weil …"].forEach((x, i) => {
    num(s, i + 1, 0.6, 1.85 + i * 0.85);
    txt(s, x, { x: 1.3, y: 1.89 + i * 0.85, w: 11.3, h: 0.7, fontSize: 21, bold: true });
  });
  box(s, 0.6, 4.6, 12.1, 1.3, LIGHT2);
  txt(s, "… so ein Kerl war Paddy Fingal!", { x: 0.6, y: 4.6, w: 12.1, h: 1.3, fontFace: HEAD, fontSize: 40, bold: true, color: RED, align: "center", valign: "middle" });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "6 Minuten · alle gemeinsam, dann allein" }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}

uebung("Verarbeiten · Ü1 b", "Ü1 b) Orientierend lesen", [
  "Suche nur die Antworten – lies nicht alles!",
  "Schreibe b) ins Heft. Notiere die Antwort und die Strophe.",
], "6 Minuten · allein · leise · Wer fertig ist, macht Frage 5*",
"Lösung: 1 sechs Ellen (Str. 1) · 2 aus Schottland (Str. 2) · 3 Schaya (Str. 4) · 4 nur seine Nase (Str. 5) · *5 eine flache Stelle in einem Fluss oder Bach (Worterklärung 7). Nachfragen: Wie hast du die Stelle so schnell gefunden? (Schlüsselwörter, Fußnoten)",
(s, y) => {
  const q = ["Wie lang ist Paddys Bart?", "Woher kommt der Fremde?", "Wie heißt Paddys Frau?", "Was guckt aus dem Bett heraus?", "* Was ist eine Furt? (Tipp: Worterklärungen)"];
  q.forEach((x, i) => txt(s, [{ text: `${i + 1}  `, options: { bold: true, color: YEL } }, { text: x }], { x: 1.3, y: y + i * 0.55, w: 11.3, h: 0.5, fontSize: 19 }));
});

uebung("Verarbeiten · Ü1 c", "Ü1 c) Genau lesen", [
  "Schreibe c) ins Heft. Lies Strophe 7 bis 9 Satz für Satz.",
  "Notiere zu jeder Strophe einen Stichpunkt.",
  "Beantworte: Warum läuft der Schotte weg?",
], "5 Minuten · allein · leise",
"Lösung: Str. 7 Schaya warnt: Paddys Kind schläft. · Str. 8 Der Schotte schaut ins Bett und sieht die riesige Nase. · Str. 9 Er denkt: Wenn schon das Baby so groß ist, wie groß ist dann der Vater? · Warum? Schayas List: Sie gibt Paddy als Baby aus – der Schotte bekommt Angst vor dem „Vater“.",
(s, y) => {
  box(s, 0.6, y + 0.25, 12.1, 1.0, LIGHT2);
  txt(s, "„Ist das ein Baby, wie Ihr sagt, – ein Narr, wer’s mit ihm selber wagt!“", { x: 0.6, y: y + 0.25, w: 12.1, h: 1.0, fontFace: HEAD, fontSize: 24, italic: true, color: GREEN, align: "center", valign: "middle" });
});

// ============ 9 Vergleich ============
{
  const s = base("Verarbeiten · Vergleich", "Vergleicht zu zweit", "Danach erklärt ein Schüler im Plenum Schayas List (Frage c).");
  ["Vergleicht a, b und c.", "Verbessert mit Grün.", "Erklärt euch gegenseitig Schayas List."].forEach((x, i) => {
    num(s, i + 1, 0.6, 1.85 + i * 0.9);
    txt(s, x, { x: 1.3, y: 1.89 + i * 0.9, w: 11, h: 0.6, fontSize: 24, bold: true });
  });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "3 Minuten · zu zweit" }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}

// ============ 10 Auswerten ============
{
  const s = base("Auswerten", "Welche Lesetechnik passt?", "Lösung: 1 überfliegend · 2 orientierend · 3 genau.");
  txt(s, "Schreibe ins Heft die passende Lesetechnik. 1 Minute.", { x: 0.6, y: 1.55, w: 12.1, h: 0.5, fontSize: 18, color: MUTED });
  ["Du willst wissen, ob ein Zeitungsartikel über deinen Verein berichtet.", "Du suchst, wie lang Paddys Bart ist.", "Du sollst erklären, warum der Schotte wegläuft."].forEach((d, i) => {
    const y = 2.3 + i * 1.3;
    box(s, 0.6, y, 12.1, 1.1, LIGHT);
    txt(s, String(i + 1), { x: 0.9, y: y + 0.18, w: 0.8, h: 0.7, fontFace: HEAD, fontSize: 36, bold: true, color: RED });
    txt(s, d, { x: 1.8, y: y + 0.3, w: 10.6, h: 0.6, fontSize: 22 });
  });
}

// ============ 11 Ziel erreicht? ============
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "Ziel erreicht?", { x: 0.9, y: 1.6, w: 11, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, "Ich kann je nach Leseziel überfliegend, orientierend oder genau lesen und meinen ersten Eindruck von einer Ballade begründen.", { x: 0.9, y: 2.6, w: 11.5, h: 0.9, fontSize: 22, italic: true, color: ICE });
  box(s, 0.9, 4.0, 11.5, 1.2, NAVY2);
  txt(s, "Schreibe ins Heft:  Heute habe ich gelernt, …", { x: 1.2, y: 4.0, w: 11, h: 1.2, fontFace: HEAD, fontSize: 30, bold: true, color: YEL, valign: "middle" });
  s.addNotes("Zwei Schüler lesen vor. Erwartet: die drei Lesetechniken und wann man sie benutzt. * Wenn Zeit bleibt: LB-Aufgabe 1 b – ein Schüler erzählt die Ballade mündlich nach (Titel, Autor, Reihenfolge, nichts dazuerfinden).");
}

pres.writeFile({ fileName: "Praesentation_Paddy_Fingal_Lesetechniken.pptx" });
