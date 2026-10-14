const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Paddy Fingal – nacherzählen";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Paddy Fingal – wiederholen und nacherzählen · Deutsch 7A";
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

const ZIEL = "Ich kann mit den Lesetechniken Informationen in der Ballade finden und den Inhalt mit eigenen Worten nacherzählen.";
const zeit = (s, meta) => { box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2); txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 }); };

// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 2.6, 0.42, YEL, { r: 0.21 });
  txt(s, "BALLADEN", { x: 0.9, y: 2.05, w: 2.6, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Paddy Fingal – wiederholen und nacherzählen", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE });
  txt(s, "Wilhelm Brandes · Lehrbuch S. 136 f.", { x: 0.9, y: 3.7, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
}
// 2 Ziel und Ablauf
{
  const s = base("Heute", "Ziel und Ablauf", "Ziel vorlesen. Am Ende darauf zurückkommen.");
  box(s, 0.6, 1.7, 12.1, 1.2, LIGHT2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 0.12, h: 1.2, fill: { color: RED }, line: { type: "none" } });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 1.0, y: 1.9, w: 11.5, h: 0.9, fontSize: 22 });
  ["Wiederholung: Welche Lesetechnik passt?", "Ü1 Lesetechniken anwenden (Frage 5 *)", "LB-Aufgabe 1 a: Fragen zur Ballade", "LB-Aufgabe 1 b: nacherzählen", "Was habe ich heute gelernt?"].forEach((a, i) => {
    num(s, i + 1, 0.9, 3.25 + i * 0.68, NAVY, 0.45);
    txt(s, a, { x: 1.55, y: 3.28 + i * 0.68, w: 10, h: 0.5, fontSize: 20 });
  });
}
// 3 Wiederholung: Situationen
{
  const s = base("Wiederholung", "Welche Lesetechnik passt?", "Lösung: A orientierend · B überfliegend · C genau. Kurz begründen lassen. Wer den Merksatz noch nicht hat, bekommt den Schnipsel „Lesetechniken“ zum Einkleben.");
  txt(s, "Schreibe ins Heft: A, B, C – und die passende Lesetechnik.", { x: 0.6, y: 1.55, w: 12.1, h: 0.5, fontSize: 18, color: MUTED });
  [["A", "Du suchst im Lehrbuch, was das Wort „Hüne“ bedeutet."], ["B", "Du willst schnell wissen, ob die Ballade lustig oder traurig endet."], ["C", "Du sollst erklären, welche List Paddys Frau anwendet."]].forEach(([k, d], i) => {
    const y = 2.3 + i * 1.3;
    box(s, 0.6, y, 12.1, 1.1, LIGHT);
    txt(s, k, { x: 0.9, y: y + 0.18, w: 0.8, h: 0.7, fontFace: HEAD, fontSize: 36, bold: true, color: RED });
    txt(s, d, { x: 1.8, y: y + 0.3, w: 10.6, h: 0.6, fontSize: 22 });
  });
  zeit(s, "2 Minuten · allein, dann kurz vergleichen");
}
// 4 Zur Erinnerung
{
  const s = base("Wiederholung", "Zur Erinnerung: drei Lesetechniken", "Nicht abschreiben – bei Bedarf den Schnipsel einkleben lassen.");
  [["ue", "schnell über den Text (Überschrift, Anfang, Ende, auffällige Wörter)", "worum es geht"], ["or", "gezielt suchen, den Rest überspringen", "eine bestimmte Information"], ["ge", "Satz für Satz, Unbekanntes klären, Notizen machen", "alles verstehen"]].forEach(([k, a, c], i) => {
    const y = 1.9 + i * 1.35;
    box(s, 0.6, y, 3.2, 0.9, TECH[k].col, { r: 0.45 });
    txt(s, TECH[k].sym + "  " + TECH[k].name, { x: 0.6, y, w: 3.2, h: 0.9, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [{ text: a + " – ich will " }, { text: c, options: { bold: true, color: TECH[k].col } }, { text: "." }], { x: 4.1, y: y + 0.1, w: 8.6, h: 0.8, fontSize: 21, valign: "middle" });
  });
}
// 5 Ü1 a
{
  const s = base("Wiederholung · Ü1 a", "Ü1 a) Überfliegen", "Zeit laut stoppen. Lösung z. B.: In der Ballade geht es um den Riesen Paddy Fingal und einen noch größeren Schotten, vor dem er sich mit einer List rettet.");
  ["Schreibe ins Heft: Ü1 Paddy Fingal – a). Schlage S. 136 f. auf.", "Überfliege die Ballade – genau 1 Minute.", "Klappe das Buch zu. Schreibe einen Satz: „In der Ballade geht es um …“"].forEach((x, i) => { num(s, i + 1, 0.6, 1.8 + i * 0.8); txt(s, x, { x: 1.3, y: 1.84 + i * 0.8, w: 11.3, h: 0.7, fontSize: 21, bold: true }); });
  box(s, 0.6, 4.4, 12.1, 1.3, LIGHT);
  txt(s, "⏱  1 Minute", { x: 0.6, y: 4.55, w: 12.1, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: BLUE, align: "center", valign: "middle" });
  zeit(s, "3 Minuten · allein · leise");
}
// 6 Ü1 b
{
  const s = base("Wiederholung · Ü1 b", "Ü1 b) Orientierend lesen", "Lösung: 1 sechs Ellen (Str. 1) · 2 aus Schottland (Str. 2) · 3 Schaya (Str. 4) · 4 nur seine Nase (Str. 5) · *5 flache Stelle in einem Fluss oder Bach (Worterklärung 7).");
  ["Schreibe b) ins Heft und nummeriere 1 bis 5.", "Suche nur die Antworten – lies nicht alles! Notiere die Strophe."].forEach((x, i) => { num(s, i + 1, 0.6, 1.8 + i * 0.8); txt(s, x, { x: 1.3, y: 1.84 + i * 0.8, w: 11.3, h: 0.7, fontSize: 21, bold: true }); });
  ["Wie lang ist Paddys Bart?", "Woher kommt der Fremde?", "Wie heißt Paddys Frau?", "Was guckt aus dem Bett heraus?", "* Was ist eine Furt? (Tipp: Worterklärungen)"].forEach((x, i) => txt(s, [{ text: `${i + 1}  `, options: { bold: true, color: YEL } }, { text: x }], { x: 1.3, y: 3.5 + i * 0.5, w: 11.3, h: 0.45, fontSize: 19 }));
  zeit(s, "6 Minuten · allein · leise · danach 2 Minuten zu zweit vergleichen");
}
// 7 LB 1 a
{
  const s = base("Verarbeiten · LB-Aufgabe 1 a", "Wovon handelt die Ballade?", "Lösung: Aussehen: Riese, schwarzer Bart sechs Ellen lang, Fäuste so groß wie Tonnen, Mastbaum als Wurfgeschoss (Str. 1) · Der Schotte will herausfinden, wer stärker ist (Str. 2) · List: Schaya steckt Paddy ins Bett, türmt die Kissen hoch und gibt ihn als Baby aus; der Schotte sieht nur die riesige Nase, denkt an einen riesigen Vater und flieht (Str. 5–10).");
  ["Lest die Ballade zu zweit aufmerksam (Tandem, halblaut).", "Beantwortet die drei Fragen im Heft – in ganzen Sätzen, mit Strophe."].forEach((x, i) => { num(s, i + 1, 0.6, 1.8 + i * 0.8); txt(s, x, { x: 1.3, y: 1.84 + i * 0.8, w: 11.3, h: 0.7, fontSize: 21, bold: true }); });
  ["Wie sieht Paddy Fingal aus?", "Was möchte der Riese aus Schottland von Paddy?", "Welche List wendet Paddys Frau an?"].forEach((x, i) => {
    box(s, 1.3, 3.5 + i * 0.85, 11.4, 0.7, LIGHT);
    txt(s, [{ text: `${i + 1}  `, options: { bold: true, color: GREEN } }, { text: x }], { x: 1.55, y: 3.62 + i * 0.85, w: 11, h: 0.5, fontSize: 20 });
  });
  zeit(s, "8 Minuten · zu zweit (Tandem)");
}
// 8 LB 1 b
{
  const s = base("Verarbeiten · LB-Aufgabe 1 b", "Erzählt die Ballade nach", "Partner hakt die vier Punkte ab und gibt eine Rückmeldung. Dann Wechsel. * Spannend mit wörtlicher Rede.");
  ["A erzählt, B hört zu und hakt die Checkliste ab. Dann Wechsel.", "Benutzt eure Antworten aus 1 a als Spickzettel."].forEach((x, i) => { num(s, i + 1, 0.6, 1.8 + i * 0.8); txt(s, x, { x: 1.3, y: 1.84 + i * 0.8, w: 11.3, h: 0.7, fontSize: 21, bold: true }); });
  box(s, 0.6, 3.45, 12.1, 2.6, LIGHT2);
  txt(s, "Checkliste", { x: 0.9, y: 3.55, w: 5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: NAVY });
  ["☐  Titel und Autor genannt", "☐  richtige Reihenfolge, nichts dazuerfunden", "☐  alle wichtigen Details (Aussehen, Schotte, Versteck, List, Flucht)", "☐  * spannend erzählt, z. B. mit wörtlicher Rede"].forEach((x, i) => txt(s, x, { x: 0.9, y: 4.0 + i * 0.48, w: 11.5, h: 0.45, fontSize: 19 }));
  zeit(s, "10 Minuten · im Tandem");
}
// 9 Auswerten
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "Ziel erreicht?", { x: 0.9, y: 1.4, w: 11, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, ZIEL, { x: 0.9, y: 2.4, w: 11.5, h: 0.9, fontSize: 22, italic: true, color: ICE });
  txt(s, "Ein Paar erzählt vor – alle prüfen mit der Checkliste.", { x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 22, color: WHITE });
  box(s, 0.9, 4.5, 11.5, 1.2, NAVY2);
  txt(s, "Schreibe ins Heft:  Heute habe ich gelernt, …", { x: 1.2, y: 4.5, w: 11, h: 1.2, fontFace: HEAD, fontSize: 30, bold: true, color: YEL, valign: "middle" });
  s.addNotes("Erwartet: Lesetechniken gezielt einsetzen; eine Ballade in der richtigen Reihenfolge mit allen wichtigen Details nacherzählen.");
}
pres.writeFile({ fileName: "Paddy_Fingal_Wiederholen_Nacherzaehlen.pptx" });
