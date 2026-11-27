const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Einleitung eines Sachtexts";
const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4", VIO = "7A5BA6", TEAL_ = "2A8C8C";
const HEAD = "Cambria", BODY = "Calibri";
let FOOT = "Mittwoch, 4.11. · Deutsch 7A";
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

function section(nr, title, sub, col) {
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 2.2, 0.42, col, { r: 0.21 });
  txt(s, "TEIL " + nr, { x: 0.9, y: 2.05, w: 2.2, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, title, { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, sub, { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
  pageNo++;
  return s;
}

// ================= Stunde 27.11. · AVIVA =================
// Leseband · Kontrolle Reportage-Übung (Aussagen prüfen, LB S. 183–184) · Merkmale eines Einleitungstextes · eigene Einleitung zur Sportart
const AUS = JSON.parse(require("fs").readFileSync(__dirname + "/aussagen_reportage.json", "utf8"));
FOOT = "Deutsch 7A · Sachtexte · Einleitung";
const ZIEL = "Ich kann die Einleitung eines Sachtexts zu meiner Sportart schreiben: Sie nennt das Thema, weckt Interesse und gibt einen Überblick.";
const zeit = (s, meta, y = 6.25, end = "Fertig? Stift hinlegen.") => {
  box(s, 0.6, y, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: end, options: { italic: true, color: MUTED } }], { x: 0.85, y: y + 0.08, w: 11.7, h: 0.4, fontSize: 16 });
};
// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "DEUTSCH 7A · SACHTEXTE VERSTEHEN", { x: 0.9, y: 2.0, w: 11, h: 0.4, fontSize: 14, color: ICE, charSpacing: 2 });
  txt(s, "Der erste Eindruck zählt", { x: 0.9, y: 2.5, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE });
  txt(s, "Die Einleitung eines Sachtexts – zu deiner Sportart", { x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 22, italic: true, color: ICE });
}
// 2 Ziel und Ablauf
{
  const s = base("Ankommen", "Ziel und Ablauf", "Ansage: „Stifte liegen. Ihr hört nur zu. Zuerst Leseband, dann kontrollieren wir die Übung zur Reportage. Danach werdet ihr selbst Autoren: Ihr schreibt die Einleitung für einen Sachtext über eure Sportart.“");
  box(s, 0.6, 1.55, 12.1, 1.2, WHITE, { line: RED, lw: 2.5 });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 0.9, y: 1.68, w: 11.6, h: 0.95, fontSize: 19, valign: "middle" });
  [["Leseband"], ["Kontrolle: Reportage – Aussagen prüfen"], ["Wie beginnt ein Sachtext? Merkmale einer Einleitung"], ["Schreiben: Einleitung zu meiner Sportart"], ["* Partnercheck mit der Checkliste"]].forEach(([a], i) => {
    const y = 3.0 + i * 0.62;
    num(s, i + 1, 0.6, y, NAVY, 0.45);
    txt(s, a, { x: 1.25, y: y + 0.04, w: 11.2, h: 0.45, fontSize: 19, bold: true });
  });
}
// 3 Leseband
{
  const s = base("Vorwissen aktivieren · Leseband", "Leseband", "Leseband mit dem eigenen Leseband-Text und der vereinbarten Methode (z. B. Ich-Du-Wir-Würfeln). Festes Stopp-Signal.");
  txt(s, "Lest euren Leseband-Text mit der Methode, die ich ansage.", { x: 0.6, y: 1.8, w: 12.1, h: 0.6, fontSize: 24, bold: true });
  txt(s, "Nach jedem Abschnitt: Fasst in einem Satz zusammen, worum es ging.", { x: 0.6, y: 2.6, w: 12.1, h: 0.6, fontSize: 20 });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Stopp-Signal: ", options: { bold: true } }, { text: "Stifte und Texte liegen, Blick nach vorn." }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}
// 4–13 Kontrolle Reportage
{
  const s = base("Vorwissen aktivieren · Kontrolle", "Kontrolle: Sehen mit den Fingerkuppen – Aussagen prüfen", "Ablauf: Wer drankommt, nennt Urteil, Zeile und Zitat. Bei falschen Aussagen: Verbesserung in eigenen Worten (nicht das Zitat noch einmal abschreiben – Rückmeldung 26.11.). Dann aufdecken, alle verbessern mit Grün.");
  box(s, 0.6, 1.6, 12.1, 2.1, LIGHT);
  txt(s, [{ text: "Beleg ", options: { bold: true, color: NAVY } }, { text: "= Zeile + Zitat aus dem Text (abschreiben).\n" }, { text: "Verbesserung ", options: { bold: true, color: RED } }, { text: "= die falsche Aussage in eigenen Worten richtigstellen (nicht abschreiben).\n\n" }, { text: "Beispiel: ", options: { bold: true, color: GREEN } }, { text: "„Die Ausstellung ist hell beleuchtet.“ – falsch · Beleg: Z. 1 „Herzlich willkommen in der Finsternis!“ · Verbesserung: Die Räume der Ausstellung sind ganz dunkel." }], { x: 0.9, y: 1.75, w: 11.6, h: 1.85, fontSize: 18 });
  txt(s, "Den Beleg schreibst du ab. Die Verbesserung schreibst du selbst.", { x: 0.6, y: 4.1, w: 12.1, h: 0.6, fontSize: 24, bold: true, color: RED, align: "center" });
}
AUS.forEach(([a, ok, z, zit, r], i) => {
  const s = base("Vorwissen aktivieren · Kontrolle", `Aussage ${i === 9 ? "* " : ""}${i + 1} von ${AUS.length}`, "Erst Urteil und Beleg nennen lassen, dann aufdecken. Mit Grün verbessern.");
  box(s, 0.6, 1.6, 12.1, 1.3, LIGHT);
  txt(s, "„" + a + "“", { x: 0.9, y: 1.75, w: 9.6, h: 1.0, fontSize: 24, bold: true, fontFace: HEAD, valign: "middle" });
  box(s, 10.7, 2.0, 1.8, 0.55, ok ? GREEN : RED, { r: 0.25 });
  txt(s, ok ? "✓ richtig" : "✗ falsch", { x: 10.7, y: 2.0, w: 1.8, h: 0.55, fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
  box(s, 0.6, 3.15, 12.1, 1.45, WHITE, { line: "C9D3EA" });
  txt(s, [{ text: "Beleg: ", options: { bold: true, color: NAVY } }, { text: z + "  " }, { text: zit, options: { italic: true } }], { x: 0.9, y: 3.3, w: 11.6, h: 1.2, fontSize: 20, valign: "middle" });
  if (r) {
    box(s, 0.6, 4.85, 12.1, 1.1, WHITE, { line: RED, lw: 2 });
    txt(s, [{ text: "Verbesserung (eigene Worte): ", options: { bold: true, color: RED } }, { text: r }], { x: 0.9, y: 5.0, w: 11.6, h: 0.8, fontSize: 20, valign: "middle" });
  }
});
// Merkmale einer Einleitung: entdecken
uebung("Informieren · Merkmale entdecken", "Wie beginnt ein Sachtext?", [
  "Lies zwei Einleitungen: LB S. 178 (kursiver Text bei „Die Rauchzeichen …“) und LB S. 180, Z. 1–7 (Gebärdensprache).",
  "Notiere zu zweit: Was macht der Anfang? Finde mindestens drei Dinge.",
  "Schreibe zu jedem Punkt ein Beispiel aus einer der Einleitungen dazu.",
], "6 Minuten · zu zweit · Heft", "Hilfsfragen: Woran erkennt man sofort das Thema? Was macht neugierig? Was erfährt man über den Rest des Textes? Wie lang ist die Einleitung?");
// Merksatz Einleitung
{
  const s = base("Informieren · Merksatz", "Merksatz: Die Einleitung eines Sachtexts", "Schnipsel (oberer Teil) austeilen und einkleben. Mit den Ergebnissen der Paare abgleichen.");
  box(s, 0.6, 1.5, 12.1, 0.85, WHITE, { line: RED, lw: 3 });
  txt(s, "Die Einleitung ist der Anfang eines Sachtexts. Sie nennt das Thema, weckt Interesse und gibt einen Überblick.", { x: 0.9, y: 1.58, w: 11.6, h: 0.7, fontSize: 19, bold: true, valign: "middle" });
  [["1", "Thema nennen", "Gleich am Anfang steht, worum es geht.", "„Mit den Händen sprechen – …“ (S. 180, Z. 1)"],
   ["2", "Interesse wecken", "mit einer überraschenden Tatsache, einer Frage oder direkter Ansprache", "„… das ist für gehörlose Menschen ganz selbstverständlich.“"],
   ["3", "Überblick geben", "kündigt an, was man im Text erfährt", "„… erklären wir euch hier. Außerdem könnt ihr …“ (Z. 5–7)"],
   ["4", "kurz und sachlich", "3–5 Sätze, Präsens, noch keine Einzelheiten", "Die Einzelheiten folgen im Hauptteil."]].forEach(([n, a, b, c], i) => {
    const y = 2.55 + i * 0.85;
    num(s, n, 0.6, y + 0.05, NAVY, 0.45);
    txt(s, a, { x: 1.25, y: y + 0.02, w: 3.1, h: 0.45, fontSize: 19, bold: true });
    txt(s, b, { x: 4.35, y: y + 0.02, w: 4.4, h: 0.75, fontSize: 15 });
    txt(s, c, { x: 8.8, y: y + 0.02, w: 3.9, h: 0.75, fontSize: 13, italic: true, color: MUTED });
  });
  box(s, 0.6, 6.0, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Wofür? ", options: { bold: true, color: RED } }, { text: "Ein guter Anfang entscheidet, ob jemand weiterliest.   " }, { text: "Test: ", options: { bold: true, color: RED } }, { text: "Weiß der Leser nach der Einleitung, worum es geht – und will er weiterlesen?" }], { x: 0.85, y: 6.1, w: 11.7, h: 0.4, fontSize: 15 });
}
// Schreibauftrag
uebung("Verarbeiten · Schreiben", "Schreibe die Einleitung: Sachtext über deine Sportart", [
  "Wähle eine Überschrift, z. B. „Judo – mehr als Kämpfen“.",
  "Schreibe 3–5 Sätze: Thema nennen, Interesse wecken, Überblick geben.",
  "Nutze die Satzanfänge auf dem Schnipsel. Schreibe im Präsens.",
  "* Tauscht: Prüft gegenseitig mit der Checkliste.",
], "12 Minuten · allein · auf dem Schnipsel", "Jeder schreibt über die eigene Sportart (Judo, Handball, Gewichtheben). Herumgehen, Satzanfänge zeigen, auf das Präsens achten. Wer fertig ist: Partnercheck mit der Checkliste (✓/✗), dann überarbeiten.");
// Beispiel (nicht Sportart der Klasse)
{
  const s = base("Verarbeiten · Beispiel", "So kann eine Einleitung aussehen (Beispiel: Schwimmen)", "Bewusst eine Sportart, die in der Klasse niemand macht – damit niemand abschreibt. Die drei Merkmale farbig zuordnen lassen.");
  txt(s, "Schwimmen – schneller als ein Delfin?", { x: 0.6, y: 1.55, w: 12.1, h: 0.6, fontFace: HEAD, fontSize: 26, bold: true });
  txt(s, [{ text: "Schwimmen ist eine der ältesten Sportarten der Welt. ", options: { color: NAVY, bold: true } }, { text: "Wusstest du, dass Spitzenschwimmer über 50 Meter mehr als acht Kilometer pro Stunde schnell sind? ", options: { color: RED } }, { text: "In diesem Text erfährst du, welche Schwimmarten es gibt, wie ein Training aussieht und warum Technik wichtiger ist als Kraft.", options: { color: GREEN } }], { x: 0.6, y: 2.3, w: 12.1, h: 2.0, fontSize: 21 });
  [["Thema nennen", NAVY], ["Interesse wecken", RED], ["Überblick geben", GREEN]].forEach(([a, c], i) => {
    box(s, 0.6 + i * 4.1, 4.6, 3.85, 0.6, WHITE, { line: c, lw: 2 });
    txt(s, a, { x: 0.6 + i * 4.1, y: 4.6, w: 3.85, h: 0.6, fontSize: 18, bold: true, color: c, align: "center", valign: "middle" });
  });
}
// Zielcheck
{
  const s = base("Auswerten · Zielcheck", "Zielcheck: Ziel erreicht?", "Zwei bis drei Schüler lesen ihre Einleitung vor. Die Klasse prüft mit den drei Merkmalen: Thema? Interesse? Überblick? Ausblick: Nächste Stunde schreibt ihr den Hauptteil.");
  txt(s, [{ text: "Ziel: ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 0.6, y: 1.5, w: 12.1, h: 0.8, fontSize: 18 });
  [["Thema genannt?", "Steht am Anfang, worum es geht?"], ["Interesse geweckt?", "Frage, Tatsache oder Ansprache?"], ["Überblick gegeben?", "Weiß ich, was im Text kommt?"]].forEach(([a, b], i) => {
    const y = 2.6 + i * 1.0;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    txt(s, "✓ / ✗", { x: 0.8, y: y + 0.2, w: 1.2, h: 0.45, fontSize: 18, bold: true, color: MUTED });
    txt(s, a, { x: 2.1, y: y + 0.18, w: 4.0, h: 0.45, fontSize: 20, bold: true });
    txt(s, b, { x: 6.2, y: y + 0.2, w: 6.4, h: 0.45, fontSize: 17, color: MUTED });
  });
  txt(s, "Wer liest seine Einleitung vor? Die anderen prüfen mit den drei Fragen.", { x: 0.6, y: 5.8, w: 12.1, h: 0.5, fontSize: 18, bold: true, align: "center" });
}
pres.writeFile({ fileName: __dirname + "/../Stunde_Einleitung_Sachtext.pptx" });
