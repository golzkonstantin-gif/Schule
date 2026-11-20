const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Funktionen von Sachtexten";
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

// ================= Einzelstunde (45 Min.) · AVIVA =================
// Funktionen von Sachtexten am Text „Wie Gebärdensprache funktioniert“ (LB S. 180–182) + Sprachtraining Imperativ
// Text nicht abgedruckt, nur Zeilenverweise (Zählung wie im LB, Z. 1–98).
FOOT = "Deutsch 7A · Sachtexte · Funktionen";
const ZIEL = "Ich kann erkennen, welche Absicht ein Sachtext verfolgt, und das mit Zeilenangabe belegen. Ich kann den Imperativ bilden.";
const zeit = (s, meta, y = 6.25, end = "Fertig? Stift hinlegen.") => {
  box(s, 0.6, y, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: end, options: { italic: true, color: MUTED } }], { x: 0.85, y: y + 0.08, w: 11.7, h: 0.4, fontSize: 16 });
};
const imp = (parts, size = 20) => parts.map(([x, k]) => ({ text: x, options: Object.assign({ fontSize: size }, k === "i" ? { bold: true, color: RED, underline: { style: "sng" } } : k === "m" ? { color: MUTED } : {}) }));

// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "DEUTSCH 7A · SACHTEXTE VERSTEHEN", { x: 0.9, y: 2.0, w: 11, h: 0.4, fontSize: 14, color: ICE, charSpacing: 2 });
  txt(s, "Was will der Text von mir?", { x: 0.9, y: 2.5, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE });
  txt(s, "Funktionen von Sachtexten · Sprachtraining: Imperativ", { x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 22, italic: true, color: ICE });
}
// 2 Ziel und Ablauf
{
  const s = base("Ankommen", "Ziel und Ablauf", "Ansage: „Stifte liegen. Ihr hört nur zu. Letzte Stunde habt ihr den Text über Gebärdensprache gelesen und euren Namen gebärdet. Heute fragen wir: Was will der Text eigentlich von uns? Am Ende prüft ihr im Zielcheck, ob ihr das Ziel erreicht habt.“");
  box(s, 0.6, 1.55, 12.1, 1.25, WHITE, { line: RED, lw: 2.5 });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 0.9, y: 1.7, w: 11.6, h: 1.0, fontSize: 20, valign: "middle" });
  [["Wiederholung: Aussagen prüfen · LB 6 Gebärden übersetzen", "7 Min."], ["LB 4a Zwischenüberschriften · Funktionen von Sachtexten", "10 Min."], ["LB 4b/c Absichten mit Zeile belegen", "12 Min."], ["Sprachtraining: Imperativ", "9 Min."], ["Zielcheck  (* LB 5 Internettext)", "5 Min."]].forEach(([a, m], i) => {
    const y = 3.05 + i * 0.62;
    num(s, i + 1, 0.6, y, NAVY, 0.45);
    txt(s, a, { x: 1.25, y: y + 0.04, w: 9.2, h: 0.45, fontSize: 19, bold: true });
    txt(s, m, { x: 10.6, y: y + 0.06, w: 2.0, h: 0.4, fontSize: 16, color: MUTED, align: "right" });
  });
}
// 3–4 Wiederholung: Aussagen prüfen
const AUS = [
  ["1", "Gebärdensprache ist eine Sprache, die man sieht und nicht hört.", true, "Z. 8–10", ""],
  ["2", "In Deutschland ist die Gebärdensprache seit Jahrhunderten als Sprache anerkannt.", false, "Z. 16–19", "Sie ist erst seit 2002 als eigene Sprache anerkannt."],
  ["3", "Auf der ganzen Welt gibt es nur eine einzige Gebärdensprache.", false, "Z. 23–27", "Es gibt viele Gebärdensprachen, z. B. die Chinesische und die Französische."],
  ["* 4", "Für Gebärden nutzt man immer die rechte Hand.", false, "Z. 28–32", "Man nutzt die dominante Hand – bei Linkshändern die linke."],
];
[false, true].forEach((show) => {
  const s = base("Vorwissen aktivieren · Wiederholung", show ? "Lösung – verbessert mit Grün" : "Richtig oder falsch? Belege mit Zeile (LB S. 180–181)", show
    ? "Wer drankommt, nennt Urteil und Zeile. Dann aufdecken. Alle verbessern mit Grün."
    : "Ansage: „Jetzt arbeitet ihr allein. Prüft jede Aussage: richtig oder falsch? Schreibt die Zeile dazu und verbessert falsche Aussagen. Drei Minuten.“");
  AUS.forEach(([n, a, ok, z, r], i) => {
    const y = 1.55 + i * (show ? 1.15 : 1.05);
    box(s, 0.6, y, 12.1, show ? 1.0 : 0.9, i === 3 ? WHITE : LIGHT, i === 3 ? { line: "C9D3EA" } : {});
    txt(s, n, { x: 0.8, y: y + 0.2, w: 0.7, h: 0.45, fontSize: 18, bold: true, color: VIO });
    txt(s, "„" + a + "“", { x: 1.5, y: y + 0.12, w: show ? 9.3 : 11.0, h: 0.45, fontSize: show ? 16 : 18 });
    if (show) {
      box(s, 11.0, y + 0.15, 1.5, 0.42, ok ? GREEN : RED, { r: 0.2 });
      txt(s, ok ? "richtig" : "falsch", { x: 11.0, y: y + 0.15, w: 1.5, h: 0.42, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, z + (r ? " · Richtig ist: " + r : ""), { x: 1.5, y: y + 0.6, w: 11.0, h: 0.35, fontSize: 14, color: GREEN, bold: true });
    }
  });
  if (!show) zeit(s, "3 Minuten · allein · Heft · Nr. 4 mit *");
});
// 5 LB 6
uebung("Vorwissen aktivieren · LB 6", "LB 6  Was sagen die Gebärden? (LB S. 182)", [
  "Sieh dir die Gebärden in LB 6 an.",
  "Übersetze sie mit dem Fingeralphabet (Blatt von letzter Stunde).",
  "Schreibe das Wort ins Heft. Wer es hat, gebärdet es leise zurück.",
], "2 Minuten · allein", "Kurzer, motivierender Einstieg. Lösung durch die Lehrkraft (Wort im LB prüfen).");
// 6–7 LB 4a
const ABS = [
  ["1", "Z. 8–15", "Was ist Gebärdensprache?"],
  ["2", "Z. 16–27", "Geschichte der Gebärdensprache"],
  ["3", "Z. 28–52", "Wie funktioniert die Deutsche Gebärdensprache?"],
  ["4", "Z. 53–62", "Gebärdensprache: Das Alphabet"],
  ["5", "Z. 63–98", "Wie unterhalte ich mich mit gehörlosen Menschen?"],
];
uebung("Informieren · LB 4a", "LB 4a  Zwischenüberschriften zuordnen", [
  "Lies die fünf Zwischenüberschriften in LB 4a (S. 182) und ordne sie den Absätzen 1–5 zu.",
  "Überfliege jeden Absatz noch einmal (Nummern 1–5 im Text).",
  "Notiere im Heft: Absatz 1 – …, Absatz 2 – … usw.",
], "3 Minuten · allein · dann vergleichen wir", "Erst Einzelarbeit, dann Unterrichtsgespräch: Wer drankommt, nennt Absatz und Überschrift und begründet mit einem Stichwort aus dem Absatz.");
{
  const s = base("Informieren · LB 4a Lösung", "LB 4a  Lösung", "Begründungen einfordern: z. B. Absatz 2 – „seit 2002 anerkannt“ (Geschichte). Mit Grün verbessern.");
  ABS.forEach(([a, z, u], i) => {
    const y = 1.6 + i * 0.85;
    box(s, 0.6, y, 12.1, 0.72, i % 2 ? WHITE : LIGHT, { line: "C9D3EA" });
    num(s, a, 0.75, y + 0.13, NAVY, 0.45);
    txt(s, z, { x: 1.4, y: y + 0.17, w: 1.8, h: 0.4, fontSize: 16, color: MUTED });
    txt(s, u, { x: 3.2, y: y + 0.15, w: 9.3, h: 0.45, fontSize: 20, bold: true });
  });
}
// 8 Merksatz Funktionen
const FK = [
  ["informieren", "über Tatsachen berichten, sachlich", "„erst seit 2002 … anerkannt“", NAVY],
  ["unterrichten / anleiten", "erklären, wie etwas geht", "„Für die Gebärden nutzt man …“", TEAL_],
  ["aufklären", "zeigen, was richtig ist – und was man vermeiden soll", "„Wichtig ist auch, dass du nie so tust …“", VIO],
  ["auffordern", "zu etwas bewegen: du-Ansprache, Imperativ, !", "„Also nimm doch einfach mal …!“", RED],
  ["überzeugen", "mit Argumenten eine Meinung begründen", "„…, weil …“", YEL],
];
{
  const s = base("Informieren · Merksatz", "Merksatz: Funktionen von Sachtexten", "Schnipsel austeilen und einkleben (nicht abschreiben). Infokasten LB S. 182 dazu. Ein Text kann mehrere Funktionen haben – oft wechselt sie von Absatz zu Absatz.");
  box(s, 0.6, 1.5, 12.1, 0.85, WHITE, { line: RED, lw: 3 });
  txt(s, "Sachtexte handeln von wirklichen Ereignissen und Sachverhalten. Sie können verschiedene Absichten (Funktionen) haben:", { x: 0.9, y: 1.58, w: 11.6, h: 0.7, fontSize: 18, bold: true, valign: "middle" });
  FK.forEach(([f, e, b, col], i) => {
    const y = 2.55 + i * 0.66;
    txt(s, f, { x: 0.8, y, w: 3.0, h: 0.45, fontSize: 19, bold: true, color: col });
    txt(s, e, { x: 3.8, y: y + 0.02, w: 4.7, h: 0.45, fontSize: 15 });
    txt(s, b, { x: 8.6, y: y + 0.02, w: 4.1, h: 0.45, fontSize: 14, italic: true, color: MUTED });
  });
  box(s, 0.6, 5.95, 12.1, 0.6, LIGHT2);
  txt(s, [{ text: "Test: ", options: { bold: true, color: RED } }, { text: "Was soll ich nach dem Lesen wissen, können oder tun?   " }, { text: "Wofür? ", options: { bold: true, color: RED } }, { text: "Du liest kritischer und erkennst, was ein Text will." }], { x: 0.85, y: 6.05, w: 11.7, h: 0.45, fontSize: 15 });
}
// 9–11 LB 4b/c
const AB = [
  ["1", "informieren über Gebärdensprache", "Z. 8–9: „Gebärdensprache ist eine visuelle Sprache.“"],
  ["2", "informieren über die Geschichte", "Z. 18–19: „erst seit 2002 in Deutschland als eigene Sprache anerkannt“"],
  ["3", "unterrichten / anleiten: wie man gebärdet", "Z. 28–29: „Für die Gebärden nutzt man immer seine dominante Hand.“"],
  ["4", "informieren über das Fingeralphabet", "Z. 57–61: „Das deutsche Gebärdensprachen-Alphabet wird gebraucht, wenn …“"],
  ["5", "aufklären + auffordern: richtig mit Gehörlosen umgehen", "Z. 80–81: „achte immer darauf …“ · Z. 86–88: „dass du nie so tust …“ · Z. 97–98: „Also nimm doch einfach mal die Sprache in die Hand!“"],
];
uebung("Verarbeiten · LB 4b/c", "LB 4b/c  Welche Absicht hat jeder Absatz?", [
  "Zeichne eine Tabelle ins Heft: Absatz | Absicht | Beleg (Z.).",
  "Bestimme für jeden Absatz die Hauptabsicht (Merksatz hilft).",
  "Schreibe zu jedem Absatz eine passende Textstelle mit Zeile auf.",
], "6 Minuten · allein · Heft · dann vergleichen wir", "Erst Einzelarbeit, dann Unterrichtsgespräch mit Aufdeck-Folien (Absatz 1–3, dann 4–5). Bei Absatz 5 fragen: Woran erkennt ihr das Auffordern? (du, Imperativ, Ausrufezeichen)");
[[0, 3], [0, 5]].forEach(([a, b], k) => {
  const s = base("Verarbeiten · LB 4b/c Lösung", "LB 4b/c  Lösung" + (k === 0 ? " (1/2)" : " (2/2)"), k === 0 ? "Wer drankommt, nennt Absicht und Beleg, dann aufdecken. Mit Grün verbessern." : "4d mündlich: Was wissen und können die Leser am Ende? (wissen, was Gebärdensprache ist; Fingeralphabet; wissen, wie man mit Gehörlosen Kontakt aufnimmt). Andere passende Belege gelten.");
  AB.slice(a, b).forEach(([n, f, z], i) => {
    const y = 1.55 + i * 0.95;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    num(s, n, 0.75, y + 0.2, NAVY, 0.45);
    txt(s, f, { x: 1.4, y: y + 0.08, w: 11.1, h: 0.4, fontSize: 18, bold: true });
    txt(s, z, { x: 1.4, y: y + 0.45, w: 11.1, h: 0.4, fontSize: 14, italic: true, color: MUTED });
  });
});
// 12 Imperativ entdecken
section("2", "Sprachtraining: Imperativ", "Ich kann den Imperativ bilden.", RED);
{
  const s = base("Sprachtraining · entdecken", "Wie fordert der Text auf? (Absatz 5)", "Gemeinsam: Was ist bei 2 und 3 anders als bei 1? (Verb steht vorn, kein „du“, Ausrufezeichen möglich). 1 ist auch eine Anweisung, aber ohne Imperativ: „du“ + Präsens.");
  [["1", [["Dazu ", ""], ["stampfst", "m"], [" du einfach auf den Boden …", ""]], "Z. 73–74", "Anweisung mit „du“ + Präsens"], ["2", [["… ", ""], ["achte", "i"], [" immer darauf, dass dein Mund nicht verdeckt ist …", ""]], "Z. 80–81", "Imperativ"], ["3", [["Also ", ""], ["nimm", "i"], [" doch einfach mal die Sprache in die Hand!", ""]], "Z. 97–98", "Imperativ"]].forEach(([n, r, z, l], i) => {
    const y = 1.6 + i * 1.1;
    box(s, 0.6, y, 12.1, 0.95, LIGHT);
    num(s, n, 0.75, y + 0.25, VIO, 0.45);
    txt(s, imp(r, 20), { x: 1.4, y: y + 0.12, w: 8.5, h: 0.5, fontFace: HEAD });
    txt(s, z, { x: 1.4, y: y + 0.58, w: 3, h: 0.3, fontSize: 13, color: MUTED });
    txt(s, l, { x: 9.9, y: y + 0.28, w: 2.7, h: 0.4, fontSize: 15, bold: true, color: i ? RED : MUTED, align: "right" });
  });
  txt(s, "Was ist bei 2 und 3 anders als bei 1?", { x: 0.6, y: 5.05, w: 12.1, h: 0.5, fontSize: 22, bold: true, align: "center" });
}
// 13 Merksatz Imperativ
{
  const s = base("Sprachtraining · Merksatz", "Merksatz: Der Imperativ (Befehlsform)", "Schnipsel (zweiter Teil) einkleben. Auf den e→i/ie-Wechsel hinweisen: nehmen → nimm, lesen → lies, geben → gib. Nicht: „Nehm!“, „Lese!“");
  box(s, 0.6, 1.5, 12.1, 0.85, WHITE, { line: RED, lw: 3 });
  txt(s, "Mit dem Imperativ fordert man jemanden direkt auf. Das Verb steht am Satzanfang, bei du und ihr ohne Pronomen.", { x: 0.9, y: 1.58, w: 11.6, h: 0.7, fontSize: 18, bold: true, valign: "middle" });
  [["du", "Verbstamm (+ e)", [["Achte", "i"], [" darauf! · ", ""], ["Lauf", "i"], [" schneller!", ""]]], ["du: e → i/ie", "wie in der du-Form Präsens", [["Nimm", "i"], [" die Hantel! · ", ""], ["Lies", "i"], [" genau! · ", ""], ["Gib", "i"], [" ab!", ""]]], ["ihr", "ihr-Form ohne „ihr“", [["Achtet", "i"], [" darauf! · ", ""], ["Nehmt", "i"], [" die Hantel!", ""]]], ["Sie", "Verb + Sie", [["Nehmen", "i"], [" Sie Platz! · ", ""], ["Achten", "i"], [" Sie darauf!", ""]]], ["sein", "unregelmäßig", [["Sei", "i"], [" leise! · ", ""], ["Seid", "i"], [" leise! · ", ""], ["Seien", "i"], [" Sie leise!", ""]]]].forEach(([a, b, r], i) => {
    const y = 2.55 + i * 0.62;
    txt(s, a, { x: 0.8, y, w: 2.2, h: 0.45, fontSize: 18, bold: true, color: NAVY });
    txt(s, b, { x: 3.0, y: y + 0.03, w: 3.4, h: 0.45, fontSize: 15, color: MUTED });
    txt(s, imp(r, 18), { x: 6.4, y: y + 0.02, w: 6.3, h: 0.45, fontFace: HEAD });
  });
  box(s, 0.6, 5.75, 12.1, 0.6, LIGHT2);
  txt(s, [{ text: "Wofür? ", options: { bold: true, color: RED } }, { text: "Anleitungen, Tipps, Trainer-Anweisungen.   " }, { text: "Test: ", options: { bold: true, color: RED } }, { text: "Steht das Verb vorn, ohne „du“/„ihr“? Dann ist es ein Imperativ." }], { x: 0.85, y: 5.85, w: 11.7, h: 0.45, fontSize: 15 });
}
// 14–15 Ü Imperativ
const UI = [
  ["a", "(laufen) schneller", "Lauf schneller!", "Lauft schneller!"],
  ["b", "(nehmen) die Hantel", "Nimm die Hantel!", "Nehmt die Hantel!"],
  ["c", "(geben) den Ball ab", "Gib den Ball ab!", "Gebt den Ball ab!"],
  ["d", "(sein) konzentriert", "Sei konzentriert!", "Seid konzentriert!"],
];
[false, true].forEach((show) => {
  const s = base("Sprachtraining · Ü", show ? "Ü  Lösung" : "Ü  Der Trainer gibt Anweisungen", show ? "Mit Grün verbessern. Häufige Fehler: „Nehm!“, „Geb!“ – e→i beachten." : "Ansage: „Jetzt arbeitet ihr allein. Bildet den Imperativ für eine Person (du) und für die Gruppe (ihr). Vier Minuten.“");
  if (!show) txt(s, "Bilde den Imperativ: einmal für eine Person (du), einmal für die Gruppe (ihr).", { x: 0.6, y: 1.5, w: 12.1, h: 0.45, fontSize: 19, bold: true });
  ["", "du", "ihr"].forEach((h, j) => { if (j) txt(s, h, { x: j === 1 ? 5.3 : 9.0, y: 2.05, w: 3.5, h: 0.4, fontSize: 17, bold: true, color: MUTED }); });
  UI.forEach(([n, x, du, ihr], i) => {
    const y = 2.5 + i * 0.85;
    box(s, 0.6, y, 12.1, 0.72, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.17, w: 0.5, h: 0.4, fontSize: 18, bold: true, color: VIO });
    txt(s, x, { x: 1.3, y: y + 0.17, w: 3.9, h: 0.4, fontSize: 18 });
    if (show) {
      txt(s, du, { x: 5.3, y: y + 0.17, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: RED, fontFace: HEAD });
      txt(s, ihr, { x: 9.0, y: y + 0.17, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: RED, fontFace: HEAD });
    }
  });
  if (!show) zeit(s, "4 Minuten · allein · Heft");
});
// 16–17 Zielcheck
const ZC = [
  ["1", "„Die Gebärdensprache ist seit 2002 in Deutschland als eigene Sprache anerkannt.“ – Welche Absicht?", "informieren (Tatsache, sachlich)"],
  ["2", "„Stampf auf den Boden, damit die Person dich bemerkt!“ – Welche Absicht? Woran erkennst du sie?", "auffordern / anleiten – Imperativ „Stampf“, Ausrufezeichen"],
  ["3", "Verbessere den Imperativ: „Lese den Text genau!“", "„Lies den Text genau!“ (e → ie)"],
];
[false, true].forEach((show) => {
  const s = base("Auswerten · Zielcheck", show ? "Zielcheck: Lösung – Ziel erreicht?" : "Zielcheck", show ? "Mit Grün verbessern. Haben wir das Ziel erreicht? Notieren, wer Fehler hat – Wiederholung nächste Stunde." : "Allein im Heft, 3 Minuten.");
  ZC.forEach(([n, a, l], i) => {
    const y = 1.55 + i * 1.35;
    box(s, 0.6, y, 12.1, 1.2, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.15, w: 0.5, h: 0.45, fontSize: 20, bold: true, color: VIO });
    txt(s, a, { x: 1.35, y: y + 0.12, w: 11.2, h: 0.6, fontSize: 17 });
    if (show) txt(s, l, { x: 1.35, y: y + 0.72, w: 11.2, h: 0.4, fontSize: 16, bold: true, color: GREEN });
  });
  if (show) txt(s, [{ text: "Ziel erreicht?  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 0.6, y: 5.75, w: 12.1, h: 0.6, fontSize: 15 });
  else zeit(s, "3 Minuten · allein im Heft", 5.85);
});
// 18 * LB 5
uebung("* Für Schnelle / Ausblick", "* LB 5  Internettext oder gedruckter Text?", [
  "Erkläre an Beispielen, was den Internettext vom gedruckten Text unterscheidet.",
  "Nenne Gemeinsamkeiten.",
  "Beschreibe, was die Leser durch interaktive Bilder und Links zusätzlich bekommen.",
], "nur wenn Zeit bleibt · sonst nächste Stunde", "Hinweise im Text: „Zur DVD-ROM hier“, „Zum Vergrößern klicken“, Quiz, Links am Ende.");

pres.writeFile({ fileName: __dirname + "/../Funktionen_Sachtexte_Imperativ.pptx" });
