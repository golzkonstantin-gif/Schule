const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Deutsch 7A – Mittwoch 4.11.";
const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4", VIO = "7A5BA6";
const HEAD = "Cambria", BODY = "Calibri";
let FOOT = "Mittwoch, 4.11. · Deutsch 7A";
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

const NC = { 1: "2E9E6B", 2: "5AAE62", 3: "9DB84A", 4: "E8A33D", 5: "E07B45", 6: "D9534F" };
function badge(s, n, x, y, d = 0.9) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: NC[n] }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontFace: HEAD, fontSize: d * 34, bold: true, color: WHITE, align: "center", valign: "middle" });
}

const STUFEN = {
  1: { ich: "Ich melde mich regelmäßig von mir aus. Ich erkenne das Problem, ordne es in einen größeren Zusammenhang ein, beurteile es sachlich und bringe eigene Gedanken ein. Ich drücke mich klar und mit Fachbegriffen aus.", fazit: "Die Leistung entspricht den Anforderungen in ganz besonderem Maße." },
  2: { ich: "Ich melde mich regelmäßig von mir aus. Ich verstehe auch schwierige Inhalte und ordne sie ins Thema ein. Ich unterscheide Wichtiges von Unwichtigem und weiß manchmal mehr, als im Unterricht dran war.", fazit: "Die Leistung entspricht in vollem Umfang den Anforderungen." },
  3: { ich: "Ich melde mich regelmäßig von mir aus. Was ich zum aktuellen Stoff sage, ist im Wesentlichen richtig, und ich verknüpfe es mit dem, was wir vorher in der Einheit gelernt haben.", fazit: "Die Leistung entspricht im Allgemeinen den Anforderungen." },
  4: { ich: "Ich melde mich nur gelegentlich. Ich gebe einfache Fakten aus dem gerade behandelten Stoff wieder – im Wesentlichen richtig.", fazit: "Die Leistung hat Mängel, entspricht im Ganzen aber noch den Anforderungen." },
  5: { ich: "Ich melde mich nicht von mir aus. Wenn ich drangenommen werde, ist meine Antwort nur teilweise richtig. Grundkenntnisse sind aber da.", fazit: "Die Leistung entspricht den Anforderungen nicht – die Lücken sind aber bald behebbar." },
  6: { ich: "Ich melde mich nicht von mir aus. Wenn ich drangenommen werde, ist meine Antwort falsch.", fazit: "Die Leistung entspricht den Anforderungen nicht; selbst Grundkenntnisse fehlen." },
};


// ============ Titel ============
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

// ================= Präsentation Mi 4.11.: Teil 1 Paddy Fingal – Inhalt · Teil 2 Nebensätze mit Subjunktion =================
const Z1 = "Ich kann mit Strophe und Vers belegen, ob Aussagen zur Ballade richtig oder falsch sind.";
const Z2 = "Ich kann Nebensätze mit Subjunktion erkennen und richtig bilden: Das finite Verb steht am Ende.";
const zeit = (s, meta) => { box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2); txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 }); };
// Satz mit markierten Teilen: [text, kind] kind: sub (Subjunktion, violett, Kasten-Optik), fin (finites Verb, unterstrichen), hv (Hilfsverb finit, rot unterstrichen), bold
const runs = (parts, size = 22) => parts.map(([x, k]) => ({ text: x, options: Object.assign({ fontSize: size }, k === "sub" ? { bold: true, color: VIO } : k === "fin" ? { bold: true, underline: { style: "sng" } } : k === "hv" ? { bold: true, color: RED, underline: { style: "sng" } } : k === "b" ? { bold: true } : k === "m" ? { color: MUTED } : {}) }));

// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 3.2, 0.42, YEL, { r: 0.21 });
  txt(s, "MITTWOCH · 4.11.", { x: 0.9, y: 2.05, w: 3.2, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Deutsch 7A – Paddy Fingal: Inhalt und Sprachtraining", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 38, bold: true, color: WHITE });
  [["1", "Den Inhalt erschließen", GREEN], ["2", "Sprachtraining an der Ballade", VIO]].forEach(([n, x, c], i) => {
    box(s, 0.9 + i * 5.2, 4.1, 4.9, 0.7, c, { r: 0.35 });
    txt(s, n + "  " + x, { x: 0.9 + i * 5.2, y: 4.1, w: 4.9, h: 0.7, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}
// 2 Überblick
{
  const s = base("Überblick", "Unsere Doppelstunde: Ablauf und Ziele", "Am Anfang zeigen und am Ende noch einmal: Was haben wir geschafft? Ziele und Ablauf stehen zusätzlich an der Seitentafel.");
  const P = [
    ["1", "Paddy Fingal – den Inhalt erschließen", "ca. 45 Min.", Z1, "Wiederholung Lesetechniken · Tandemlesen · Aufbau einer Ballade · Ü1 Belege finden · Ü2 Aussagen prüfen · * Ü3 nacherzählen", GREEN],
    ["2", "Sprachtraining: Nebensätze mit Subjunktion", "ca. 40 Min.", Z2, "Ü4 Entdecken in der Ballade · Merksatz · Ü5 Sätze verbinden · * Ü6 zurück zur Ballade", VIO],
  ];
  P.forEach(([n, t, z, ziel, ab, c], i) => {
    const y = 1.7 + i * 2.25;
    box(s, 0.6, y, 12.1, 2.0, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.12, h: 2.0, fill: { color: c }, line: { type: "none" } });
    num(s, n, 0.95, y + 0.3, c, 0.6);
    txt(s, t, { x: 1.8, y: y + 0.18, w: 8.6, h: 0.5, fontFace: HEAD, fontSize: 22, bold: true });
    txt(s, z, { x: 10.6, y: y + 0.2, w: 1.9, h: 0.4, fontSize: 16, bold: true, color: c, align: "right" });
    txt(s, [{ text: "Ziel: ", options: { bold: true, color: c } }, { text: ziel, options: { italic: true } }], { x: 1.8, y: y + 0.72, w: 10.7, h: 0.5, fontSize: 17 });
    txt(s, ab, { x: 1.8, y: y + 1.32, w: 10.7, h: 0.5, fontSize: 15, color: MUTED });
  });
  txt(s, "* = nur, wenn die Zeit reicht. Am Ende: Zielcheck – zwei kurze Aufgaben zu unseren Zielen.", { x: 0.6, y: 6.4, w: 12.1, h: 0.35, fontSize: 14, italic: true, color: MUTED });
}

// ================= TEIL 1 =================
FOOT = "Teil 1 · Paddy Fingal – den Inhalt erschließen · Deutsch 7A";
section("1", "Paddy Fingal – den Inhalt erschließen", "Wilhelm Brandes · Lehrbuch S. 136 f.", GREEN);
// Wiederholung
{
  const s = base("Vorwissen aktivieren · Wiederholung", "Welche Lesetechnik brauchst du? Begründe mit „weil“.", "Lösung: A orientierend, weil ich nur eine Information suche · B überfliegend, weil ich nur wissen will, wie es ausgeht · C genau, weil ich etwas erklären soll. Die „weil“-Sätze aufheben: In Teil 2 schauen wir sie uns genauer an. Falls am Montag E (Überfliegen) nicht geschafft wurde: jetzt 1 Minute überfliegen lassen, Satz „In der Ballade geht es um …“.");
  [["A", "Du willst wissen, woher der Fremde kommt."], ["B", "Du willst wissen, ob die Ballade gut ausgeht."], ["C", "Du sollst erklären, welche List Paddys Frau anwendet."]].forEach(([k, d], i) => {
    const y = 1.75 + i * 1.2;
    box(s, 0.6, y, 12.1, 1.0, LIGHT);
    txt(s, k, { x: 0.9, y: y + 0.13, w: 0.8, h: 0.7, fontFace: HEAD, fontSize: 34, bold: true, color: GREEN });
    txt(s, d, { x: 1.8, y: y + 0.27, w: 10.6, h: 0.6, fontSize: 22 });
  });
  txt(s, [{ text: "Schreibe ins Heft: ", options: { bold: true } }, { text: "A orientierend, weil …" }], { x: 0.6, y: 5.45, w: 12.1, h: 0.5, fontSize: 20 });
  zeit(s, "3 Minuten · allein, dann mit dem Nachbarn vergleichen");
}
// Ü1 Aussagen prüfen
const AUS = [
  ["Paddy ist so stark, dass er einen Schiffsmast wie einen Speer wirft.", true, "Str. 1, V. 4: „einen Mastbaum führt’ er als Wurfgeschoss“", ""],
  ["Der Schotte kommt zu Paddy, weil er ihn als Verbündeten gegen einen Feind gewinnen will.", false, "Str. 2, V. 7–9: „der hörte von Fingals starker Hand; da wollt’ er erproben alsogleich, wer fester sei …“", "Er hat von Paddys Kraft gehört und will herausfinden, wer von beiden stärker ist."],
  ["Als Paddy den Fremden von Weitem sieht, merkt er, dass beide ungefähr gleich groß sind.", false, "Str. 3, V. 14: „dass der Fremdling zehn Schuh größer war“", "Der Schotte ist viel größer als Paddy."],
  ["Paddy will sich dem Kampf stellen und bittet Schaya, ihm beim Vorbereiten zu helfen.", false, "Str. 4, V. 16–17: „Da lief er heim in jähem Schreck: O Schaya, birg mich im Versteck!“", "Paddy bekommt Angst und will sich verstecken."],
  ["Schaya bittet den Fremden, leise zu sein, damit das Kind nicht aufwacht.", true, "Str. 7, V. 31–32: „Tut leise, Fremder, und tretet sacht, dass Paddys Kindlein nicht erwacht!“", ""],
  ["Der Schotte hält sich an Schayas Warnung und bleibt vom Bett weg.", false, "Str. 8, V. 37: „der Schotte neugierig ans Lager ging“", "Er geht trotz der Warnung neugierig ans Bett."],
  ["Der Schotte erkennt Paddy im Bett und flieht vor ihm.", false, "Str. 9, V. 43–44: „Ist das ein Baby, wie Ihr sagt, ein Narr, wer’s mit ihm selber wagt!“", "Er hält Paddy für ein Baby und fürchtet sich vor dessen riesigem Vater."],
  ["Auf der Flucht hätte der Schotte beinahe sein Leben verloren.", true, "Str. 10, V. 48: „fast wär ertrunken der gute Held“", ""],
  ["Paddy wird am Ende zu Recht als tapferer Held gefeiert, weil er den Schotten besiegt hat.", false, "Str. 11, V. 55: „den tapfern Paddy Fingal“", "Paddy hat nie gekämpft und gewinnt nur durch Schayas List – „tapfer“ ist ironisch gemeint."],
];
// Tandemlesen
uebung("Verarbeiten · Lesen", "Paddy Fingal im Tandem lesen", [
  "Lest die Ballade im Tandem: abwechselnd Strophe für Strophe.",
  "Wer zuhört, liest leise mit und hilft bei Fehlern.",
  "Achtet darauf: Was passiert? Wer handelt?",
], "8 Minuten · Tandem", "Methodenkarte Tandemlesen (Tandems lesestärker/leseschwächer). Herumgehen und zuhören.");
// Aufbau einer Ballade
{
  const s = base("Informieren · Aufbau", "Aufbau einer Ballade – und wie ich belege", "Ich – Wir: Lehrkraft zeigt an Strophe 1, was eine Strophe und was ein Vers ist, und zählt mit der Klasse die Verse. Im Lehrbuch auf die Verszählung am Rand hinweisen (5, 10, 15 …). Dann das Beleg-Format zeigen: Str. + V. + Zitat in Anführungszeichen. Der Merkkasten steht auch oben auf dem Arbeitsblatt zu Ü2.");
  box(s, 0.6, 1.6, 6.6, 3.55, LIGHT);
  txt(s, "Strophe 1", { x: 0.85, y: 1.7, w: 3, h: 0.35, fontSize: 15, bold: true, color: GREEN });
  ["Paddy Fingal war von Riesenart:", "Sechs Ellen flog sein schwarzer Bart,", "seine Fäuste waren wie Tonnen groß,", "einen Mastbaum führt’ er als Wurfgeschoss –", "     so ein Kerl war Paddy Fingal!"].forEach((v, i) => {
    txt(s, String(i + 1), { x: 0.75, y: 2.15 + i * 0.55, w: 0.4, h: 0.4, fontSize: 15, color: MUTED, align: "right" });
    txt(s, v, { x: 1.3, y: 2.12 + i * 0.55, w: 5.8, h: 0.45, fontSize: 18, fontFace: HEAD });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 2.1, w: 0.06, h: 2.8, fill: { color: GREEN }, line: { type: "none" } });
  txt(s, "← Vers 2", { x: 6.0, y: 2.67, w: 1.0, h: 0.4, fontSize: 13, bold: true, color: RED });
  [["Strophe", "ein Abschnitt der Ballade; zwischen zwei Strophen steht eine Leerzeile. Paddy Fingal hat 11 Strophen."],
   ["Vers", "eine Zeile der Ballade. Im Lehrbuch sind die Verse am Rand gezählt: 5, 10, 15 …"]].forEach(([k, d], i) => {
    box(s, 7.45, 1.6 + i * 1.8, 5.25, 1.6, LIGHT2);
    txt(s, [{ text: k + ": ", options: { bold: true, color: GREEN } }, { text: d }], { x: 7.65, y: 1.75 + i * 1.8, w: 4.9, h: 1.35, fontSize: 17 });
  });
  box(s, 0.6, 5.4, 12.1, 1.3, WHITE, { line: RED, lw: 2.5, r: 0.05 });
  txt(s, [{ text: "So belege ich: ", options: { bold: true, color: RED } }, { text: "Strophe + Vers + Zitat in Anführungszeichen\n" }, { text: "Str. 1, V. 2: „Sechs Ellen flog sein schwarzer Bart“", options: { fontFace: HEAD, bold: true } }], { x: 0.9, y: 5.55, w: 11.6, h: 1.05, fontSize: 19 });
}
// Ü1 Belege finden
uebung("Verarbeiten · Aufbau", "Ü1  Belege finden", [
  "Zähle: Wie viele Strophen hat die Ballade? Wie viele Verse hat jede Strophe?",
  "Schlage nach: Was steht in Str. 3, V. 14? Schreibe das Zitat ab.",
  "Suche: Wo steht „birg mich im Versteck“? Gib Strophe und Vers an.",
], "3 Minuten · allein im Heft", "Lösung auf der nächsten Folie. Beim Vergleich darauf achten: Anführungszeichen beim Zitat, Strophe UND Vers.");
{
  const s = base("Verarbeiten · Aufbau", "Ü1  Lösung", "Mit Grün verbessern. Rückbezug zu den Lesetechniken: Bei 2 und 3 lest ihr orientierend – ihr sucht eine bestimmte Stelle.");
  [["1", "11 Strophen mit je 5 Versen"], ["2", "Str. 3, V. 14: „dass der Fremdling zehn Schuh größer war“"], ["3", "Str. 4, V. 17: „O Schaya, birg mich im Versteck!“"]].forEach(([n, x], i) => {
    const y = 1.8 + i * 1.2;
    box(s, 0.6, y, 12.1, 1.0, LIGHT);
    num(s, n, 0.9, y + 0.25, GREEN, 0.5);
    txt(s, x, { x: 1.7, y: y + 0.27, w: 10.8, h: 0.5, fontSize: 22, fontFace: HEAD });
  });
  txt(s, "Vergleiche und verbessere mit Grün.", { x: 0.6, y: 5.6, w: 12.1, h: 0.5, fontSize: 20, bold: true, align: "center" });
}
// Ü2 Aussagen prüfen
uebung("Verarbeiten · Ü2 (Arbeitsblatt)", "Ü2  Paddy Fingal – Aussagen prüfen", [
  "Prüfe jede Aussage auf dem Arbeitsblatt allein: richtig oder falsch?",
  "Belege mit Strophe, Vers und Zitat.",
  "Ist die Aussage falsch? Schreibe auf, was richtig ist.",
], "10 Minuten · allein · Nr. 9 mit *", "Arbeitsblatt austeilen. Die Aussagen sind bewusst in eigenen Worten formuliert: Die Schüler müssen die Stelle verstanden haben, um sie zu finden. Herumgehen; wer fertig ist, macht * Nr. 9. Lösungen auf den folgenden Aufdeck-Folien.", (s, y) => {
  box(s, 0.6, y, 12.1, 0.75, LIGHT2);
  txt(s, [{ text: "Beispiel: ", options: { bold: true, color: GREEN } }, { text: "„Paddy hat einen roten Bart.“ – falsch · Str. 1, V. 2, „Sechs Ellen flog sein schwarzer Bart“ · Richtig ist: Paddys Bart ist schwarz." }], { x: 0.85, y: y + 0.17, w: 11.6, h: 0.5, fontSize: 15 });
});
// Aufdeck-Folien Ü1: k Lösungen sichtbar
for (let k = 0; k <= AUS.length; k++) {
  const s = base("Auswerten · Teil 1 · gemeinsam prüfen", "Ü2  Richtig oder falsch?", k === 0
    ? "Ansage: Wir prüfen jetzt gemeinsam, Aussage für Aussage. Wer drankommt, sagt: richtig oder falsch und den Beleg. Erst dann decke ich auf. Alle verbessern auf dem Arbeitsblatt mit Grün."
    : `Aussage ${k}: erst einen Schüler Urteil und Beleg nennen lassen, bei falschen Aussagen auch „Richtig ist …“. Dann diese Folie zeigen. Alle verbessern mit Grün.` + (k === 9 ? " Nr. 9 ist eine Deutungsfrage: Die Ballade ist ironisch – der Feigling wird als Held gefeiert." : ""));
  AUS.forEach(([a, ok], i) => {
    const y = 1.45 + i * 0.43;
    const cur = i === k - 1;
    box(s, 0.6, y, 12.1, 0.38, cur ? "FFF4DE" : (i % 2 ? WHITE : LIGHT), { r: 0.04 });
    txt(s, (i === 8 ? "* " : "") + (i + 1), { x: 0.7, y: y + 0.06, w: 0.55, h: 0.3, fontSize: 14, bold: true, color: GREEN, align: "center" });
    txt(s, a, { x: 1.3, y: y + 0.07, w: 9.6, h: 0.3, fontSize: 14, bold: cur });
    if (i < k) {
      box(s, 11.05, y + 0.04, 1.55, 0.3, ok ? GREEN : RED, { r: 0.15 });
      txt(s, ok ? "✓ richtig" : "✗ falsch", { x: 11.05, y: y + 0.04, w: 1.55, h: 0.3, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    }
  });
  if (k > 0) {
    const [, ok, beleg, richtig] = AUS[k - 1];
    box(s, 0.6, 5.4, 12.1, 1.45, LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.4, w: 0.12, h: 1.45, fill: { color: ok ? GREEN : RED }, line: { type: "none" } });
    const r = [{ text: `Aussage ${k}: `, options: { bold: true } }, { text: ok ? "richtig" : "falsch", options: { bold: true, color: ok ? GREEN : RED } }, { text: "   ·   Beleg: ", options: { bold: true } }, { text: beleg }];
    if (richtig) r.push({ text: "\nRichtig ist: ", options: { bold: true, breakLine: false } }, { text: richtig });
    txt(s, r, { x: 0.9, y: 5.52, w: 11.6, h: 1.25, fontSize: 16 });
  } else {
    txt(s, "Wer drankommt, sagt: richtig oder falsch – und den Beleg. Erst dann wird aufgedeckt. Alle verbessern mit Grün.", { x: 0.6, y: 5.55, w: 12.1, h: 0.8, fontSize: 19, bold: true });
  }
}
// * Ü3 nacherzählen mit Stichwörtern
{
  const s = base("* Verarbeiten · LB-Aufgabe 1 b", "* Ü3  Erzählt die Ballade nach", "Nur, wenn die Zeit reicht; sonst Einstieg am Donnerstag. Die Stichwörter geben die Reihenfolge vor – keine fertigen Sätze. A erzählt Abschnitt 1–4, B Abschnitt 5–8, der andere hakt die Checkliste ab. Die Verknüpfungswörter (als, weil) knüpfen an Teil 2 an.");
  box(s, 0.6, 1.5, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Einstieg: ", options: { bold: true, color: GREEN } }, { text: "In der Ballade „Paddy Fingal“ von Wilhelm Brandes geht es um …" }], { x: 0.85, y: 1.6, w: 11.6, h: 0.4, fontSize: 17 });
  const AB = [
    ["1", "Paddy – Riese – Bart – Mastbaum", "Str. 1"], ["2", "Hüne aus Schottland – Wer ist stärker?", "Str. 2"],
    ["3", "Ufer – zehn Schuh größer – Schreck – Versteck", "Str. 3–4"], ["4", "Schaya – Bett – Kissen – nur die Nase", "Str. 5"],
    ["5", "Schotte sucht – Schayas Warnung – „Kindlein“", "Str. 6–7"], ["6", "Nase – Baby? – riesiger Vater", "Str. 8–9"],
    ["7", "Flucht – Sund – fast ertrunken", "Str. 10"], ["8", "Siegesgeschrei – Nachbarn – „tapfer“", "Str. 11"],
  ];
  AB.forEach(([n, w, st], i) => {
    const col = i < 4 ? 0 : 1, r = i % 4;
    const x = 0.6 + col * 6.15, y = 2.2 + r * 0.72;
    box(s, x, y, 5.95, 0.62, LIGHT);
    txt(s, n, { x: x + 0.12, y: y + 0.13, w: 0.35, h: 0.38, fontSize: 16, bold: true, color: GREEN });
    txt(s, w, { x: x + 0.5, y: y + 0.14, w: 4.4, h: 0.38, fontSize: 15 });
    txt(s, st, { x: x + 4.85, y: y + 0.15, w: 1.0, h: 0.35, fontSize: 12, color: MUTED, align: "right" });
  });
  txt(s, [{ text: "A: ", options: { bold: true } }, { text: "Abschnitt 1–4  ·  " }, { text: "B: ", options: { bold: true } }, { text: "Abschnitt 5–8  ·  " }, { text: "Verknüpfe: ", options: { bold: true, color: GREEN } }, { text: "zuerst · dann · als … · weil … · daraufhin · schließlich · am Ende" }], { x: 0.6, y: 5.15, w: 12.1, h: 0.4, fontSize: 15 });
  txt(s, [{ text: "Checkliste: ", options: { bold: true, color: GREEN } }, { text: "Titel und Autor · richtige Reihenfolge · nichts dazuerfinden · alle wichtigen Details · * spannend mit wörtlicher Rede" }], { x: 0.6, y: 5.6, w: 12.1, h: 0.4, fontSize: 14, color: MUTED });
  zeit(s, "7 Minuten · Tandem");
}

// ================= TEIL 2 =================
FOOT = "Teil 2 · Sprachtraining: Nebensätze mit Subjunktion · Deutsch 7A";
section("2", "Sprachtraining an der Ballade", "Nebensätze mit Subjunktion", VIO);
// Vorwissen finites Verb
{
  const s = base("Vorwissen aktivieren · Verb-Einheit", "Wo steht das finite Verb?", "Wiederholung aus der Verb-Einheit: Das finite Verb (die Personalform) richtet sich nach der Person. Lösung: war, lief – beide an Position 2 (nach dem ersten Satzglied). Merksatz der Verb-Einheit nicht wiederholen lassen, nur kurz fragen: Woran erkennst du das finite Verb? (ändert sich mit der Person: ich war, du warst).");
  txt(s, "Schreibe beide Sätze ab und unterstreiche das finite Verb.", { x: 0.6, y: 1.65, w: 12.1, h: 0.5, fontSize: 20, bold: true });
  [["Paddy Fingal war von Riesenart.", "Str. 1"], ["Da lief er heim in jähem Schreck.", "Str. 4"]].forEach(([x, st], i) => {
    const y = 2.45 + i * 1.25;
    box(s, 0.6, y, 12.1, 1.0, LIGHT);
    txt(s, x, { x: 0.95, y: y + 0.22, w: 9.8, h: 0.6, fontSize: 28, fontFace: HEAD });
    txt(s, st, { x: 10.9, y: y + 0.3, w: 1.6, h: 0.5, fontSize: 16, color: MUTED, align: "right" });
  });
  txt(s, [{ text: "Frage: ", options: { bold: true, color: VIO } }, { text: "An welcher Stelle steht das finite Verb?" }], { x: 0.6, y: 5.1, w: 12.1, h: 0.5, fontSize: 20 });
  zeit(s, "3 Minuten · allein");
}
// Entdecken
{
  const s = base("Informieren · Entdecken in der Ballade", "Ü4  Was passiert mit dem finiten Verb?", "Lösung: dass … war · dass … erwacht · als … sah. Erstes Wort: dass, als (Subjunktion). Das finite Verb steht am Ende. Vor der Subjunktion steht ein Komma (im Gedicht teils am Versende). Überleitung: Das ist eine Regel – Merksatz-Schnipsel austeilen.");
  const ex = [
    ["Str. 3", [["… und maß von ferne klipp und klar, ", "m"], ["dass der Fremdling zehn Schuh größer war", "b"]]],
    ["Str. 7", [["Tut leise, Fremder, und tretet sacht, ", "m"], ["dass Paddys Kindlein nicht erwacht", "b"], ["!", "m"]]],
    ["Str. 8", [["O heiliger Patrick, wie ward ihm da, ", "m"], ["als er die Nase ragen sah", "b"]]],
  ];
  ex.forEach(([st, parts], i) => {
    const y = 1.6 + i * 0.95;
    box(s, 0.6, y, 12.1, 0.8, LIGHT);
    txt(s, st, { x: 0.8, y: y + 0.22, w: 1.0, h: 0.4, fontSize: 15, color: MUTED });
    txt(s, runs(parts, 21), { x: 1.8, y: y + 0.18, w: 10.7, h: 0.5, fontFace: HEAD });
  });
  [["1", "Schreibe die fett gedruckten Teile ins Heft."], ["2", "Unterstreiche das finite Verb. Kreise das erste Wort ein."], ["3", "Vergleiche mit den Sätzen von eben: Was fällt dir auf?"]].forEach(([n, x], i) => {
    num(s, n, 0.6, 4.55 + i * 0.55, VIO, 0.42);
    txt(s, x, { x: 1.2, y: 4.58 + i * 0.55, w: 11.3, h: 0.45, fontSize: 19, bold: true });
  });
  zeit(s, "4 Minuten allein, 2 Minuten zu zweit");
}
// Merksatz
{
  const s = base("Merksatz – Schnipsel einkleben", "Nebensätze mit Subjunktion", "Schnipsel austeilen und einkleben lassen (nicht abschreiben). Gemeinsam lesen. Bezug Verb-Einheit: Im Perfekt ist das Hilfsverb finit – es wandert ans Ende. Bezug Englisch: because he is afraid – dort bleibt das Verb vorne.");
  box(s, 0.6, 1.55, 12.1, 3.3, WHITE, { line: RED, lw: 3, r: 0.05 });
  txt(s, [{ text: "Eine " }, { text: "Subjunktion", options: { bold: true, color: VIO } }, { text: " (dass, als, weil, wenn, ob, obwohl, während, nachdem) leitet einen " }, { text: "Nebensatz", options: { bold: true } }, { text: " ein. Im Nebensatz steht das " }, { text: "finite Verb am Ende", options: { bold: true, underline: { style: "sng" } } }, { text: ". Vor der Subjunktion steht ein " }, { text: "Komma", options: { bold: true } }, { text: "." }], { x: 0.95, y: 1.75, w: 11.4, h: 1.2, fontSize: 22 });
  txt(s, runs([["Der Schotte flieht, ", ""], ["weil", "sub"], [" er Angst vor dem Vater ", ""], ["hat", "fin"], [".", ""]], 22), { x: 0.95, y: 3.05, w: 11.4, h: 0.5, fontFace: HEAD });
  txt(s, runs([["Ich bin stolz, ", ""], ["weil", "sub"], [" ich 80 kg gestoßen ", ""], ["habe", "hv"], [".", ""], ["   (Perfekt: das Hilfsverb ist finit)", "m"]], 22), { x: 0.95, y: 3.65, w: 11.4, h: 0.5, fontFace: HEAD });
  txt(s, runs([["Hauptsatz", "b"], [": finites Verb an Position 2   ·   ", "m"], ["Nebensatz", "b"], [": finites Verb am Ende", "m"]], 18), { x: 0.95, y: 4.3, w: 11.4, h: 0.4 });
  box(s, 0.6, 5.05, 5.95, 1.05, LIGHT2);
  txt(s, [{ text: "Wofür? ", options: { bold: true } }, { text: "Ich kann begründen und Zusammenhänge zeigen – in jeder Analyse: „Die Ballade ist spannend, weil …“. Im Englischen bleibt das Verb vorne: because he is afraid." }], { x: 0.8, y: 5.12, w: 5.6, h: 0.95, fontSize: 14 });
  box(s, 6.75, 5.05, 5.95, 1.05, LIGHT2);
  txt(s, [{ text: "Test: ", options: { bold: true } }, { text: "Steht am Anfang dass, als, weil, wenn …? Dann muss das finite Verb ans Ende – und davor ein Komma." }], { x: 6.95, y: 5.12, w: 5.6, h: 0.95, fontSize: 14 });
  txt(s, "Markieren: Subjunktion einkreisen · finites Verb unterstreichen (Hilfsverb rot)", { x: 0.6, y: 6.3, w: 12.1, h: 0.4, fontSize: 15, italic: true, color: MUTED });
}
// Ü3 Sätze verbinden
{
  const s = base("Verarbeiten", "Ü5  Sätze verbinden", "Lösung: 1 Ich trainiere heute länger, weil am Samstag ein Wettkampf ist. 2 Der Trainer sagt, dass wir nach Frankfurt (Oder) fahren. 3 Ich war nervös, als ich auf der Matte stand. 4 Wir gewinnen, wenn wir als Team spielen. 5 Ich bin stolz, weil ich 80 kg gestoßen habe. *6 Der Trainer fragt, ob ich gut geschlafen habe. Danach gemeinsam auf den Aufdeck-Folien prüfen: Wer drankommt, liest seinen Satz vor und sagt, wo das finite Verb steht. Dann aufdecken, alle verbessern mit Grün.");
  [["1", "Verbinde die Sätze mit der Subjunktion in Klammern."], ["2", "Schreibe den ganzen Satz ins Heft. Denk an das Komma."], ["3", "Kreise die Subjunktion ein, unterstreiche das finite Verb."]].forEach(([n, x], i) => {
    num(s, n, 0.6, 1.55 + i * 0.5, VIO, 0.4);
    txt(s, x, { x: 1.15, y: 1.57 + i * 0.5, w: 11.4, h: 0.42, fontSize: 18, bold: true });
  });
  const it = [
    ["1", "Ich trainiere heute länger. Am Samstag ist ein Wettkampf.", "weil"],
    ["2", "Der Trainer sagt. Wir fahren nach Frankfurt (Oder).", "dass"],
    ["3", "Ich war nervös. Ich stand auf der Matte.", "als"],
    ["4", "Wir gewinnen. Wir spielen als Team.", "wenn"],
    ["5", "Ich bin stolz. Ich habe 80 kg gestoßen.", "weil"],
    ["* 6", "Der Trainer fragt. Hast du gut geschlafen?", "ob"],
  ];
  box(s, 0.6, 3.15, 12.1, 2.95, LIGHT);
  it.forEach(([n, x, sj], i) => {
    const y = 3.27 + i * 0.47;
    txt(s, n, { x: 0.85, y, w: 0.7, h: 0.4, fontSize: 18, bold: true, color: VIO });
    txt(s, x, { x: 1.55, y, w: 9.2, h: 0.4, fontSize: 18 });
    txt(s, "(" + sj + ")", { x: 10.8, y, w: 1.7, h: 0.4, fontSize: 18, bold: true, color: VIO, align: "right" });
  });
  zeit(s, "8 Minuten · allein");
}
// Aufdeck-Folien Ü3
const UE3 = [
  ["Ich trainiere heute länger. Am Samstag ist ein Wettkampf. (weil)", [["Ich trainiere heute länger, ", ""], ["weil", "sub"], [" am Samstag ein Wettkampf ", ""], ["ist", "fin"], [".", ""]]],
  ["Der Trainer sagt. Wir fahren nach Frankfurt (Oder). (dass)", [["Der Trainer sagt, ", ""], ["dass", "sub"], [" wir nach Frankfurt (Oder) ", ""], ["fahren", "fin"], [".", ""]]],
  ["Ich war nervös. Ich stand auf der Matte. (als)", [["Ich war nervös, ", ""], ["als", "sub"], [" ich auf der Matte ", ""], ["stand", "fin"], [".", ""]]],
  ["Wir gewinnen. Wir spielen als Team. (wenn)", [["Wir gewinnen, ", ""], ["wenn", "sub"], [" wir als Team ", ""], ["spielen", "fin"], [".", ""]]],
  ["Ich bin stolz. Ich habe 80 kg gestoßen. (weil)", [["Ich bin stolz, ", ""], ["weil", "sub"], [" ich 80 kg gestoßen ", ""], ["habe", "hv"], [".", ""]]],
  ["* Der Trainer fragt. Hast du gut geschlafen? (ob)", [["Der Trainer fragt, ", ""], ["ob", "sub"], [" ich gut geschlafen ", ""], ["habe", "hv"], [".", ""]]],
];
for (let k = 1; k <= UE3.length; k++) {
  const s = base("Verarbeiten · Ü5 gemeinsam prüfen", "Ü5  Lösung – Satz für Satz", `Satz ${k}: erst einen Schüler vorlesen lassen und fragen, wo das finite Verb steht und wo das Komma hin muss. Dann aufdecken. Alle verbessern mit Grün.` + (k === 5 ? " Nachfragen: Welches Verb ist finit? habe – das Hilfsverb (rot) wandert ans Ende." : "") + (k === 6 ? " Auch richtig: …, ob du gut geschlafen hast." : ""));
  UE3.forEach(([q, sol], i) => {
    const y = 1.5 + i * 0.78;
    box(s, 0.6, y, 12.1, 0.68, i === k - 1 ? "FFF4DE" : LIGHT, { r: 0.05 });
    txt(s, String(i + 1), { x: 0.75, y: y + 0.17, w: 0.5, h: 0.35, fontSize: 17, bold: true, color: VIO });
    if (i < k) txt(s, runs(sol, 20), { x: 1.3, y: y + 0.14, w: 11.2, h: 0.42, fontFace: HEAD });
    else txt(s, q, { x: 1.3, y: y + 0.17, w: 11.2, h: 0.4, fontSize: 16, color: MUTED });
  });
  txt(s, "Subjunktion eingekreist (violett) · finites Verb unterstrichen · Hilfsverb rot · Komma vor der Subjunktion", { x: 0.6, y: 6.35, w: 12.1, h: 0.35, fontSize: 14, italic: true, color: MUTED });
}
// * Ü5 zurück zur Ballade
{
  const s = base("* Verarbeiten · zurück zur Ballade", "* Ü6  Erkläre die Ballade mit Nebensätzen", "Nur, wenn Zeit bleibt. Satzanfänge, keine fertigen Sätze. Mögliche Ergänzungen: … weil er Angst vor dem Schotten hat · … dass Paddy ein Baby sei/ist · … als er die riesige Nase sieht · … weil er glaubt, dass der Vater noch größer ist.");
  txt(s, "Schreibe mindestens drei Sätze ins Heft. Kreise die Subjunktion ein und unterstreiche das finite Verb.", { x: 0.6, y: 1.65, w: 12.1, h: 0.6, fontSize: 20, bold: true });
  ["Paddy versteckt sich im Bett, weil …", "Schaya erzählt dem Schotten, dass …", "Der Schotte erschrickt, als …", "Der Schotte flieht, weil …", "Die Nachbarn staunen, als …"].forEach((x, i) => {
    box(s, 0.6, 2.45 + i * 0.68, 12.1, 0.56, LIGHT);
    txt(s, x, { x: 0.9, y: 2.53 + i * 0.68, w: 11.5, h: 0.42, fontSize: 20, fontFace: HEAD });
  });
  zeit(s, "5 Minuten · allein");
}
// Auswerten: Forscherfrage + Zielcheck
{
  const s = base("Auswerten", "Zielcheck: Hast du die Ziele erreicht?", "Erst (*) Forscherfrage mündlich: In Str. 3 steht „als er stapfte durch den Sund“ – das Verb steht nicht am Ende, wegen Reim und Rhythmus (Sund – stund). Dichter dürfen abweichen, im Aufsatz gilt die Regel. Dann Zielcheck allein im Heft (4 Min.), danach Lösungsfolie: Selbstkontrolle mit Grün. Herumgehen und schauen, wer noch Fehler hat – am Donnerstag in der Wiederholung aufgreifen.");
  box(s, 0.6, 1.55, 12.1, 0.85, LIGHT2);
  txt(s, [{ text: "* Forscherfrage: ", options: { bold: true, color: VIO } }, { text: "Str. 3: „Und als er stapfte durch den Sund“ – was ist hier anders, und warum?" }], { x: 0.9, y: 1.7, w: 11.6, h: 0.6, fontSize: 18 });
  [[GREEN, "Ziel 1", "Prüfe mit Strophe und Vers: „Der Schotte flieht, weil Paddy ihn im Bett angreift.“ – richtig oder falsch?"],
   [VIO, "Ziel 2 a", "Verbinde: Der Schotte erschrickt. Er sieht die riesige Nase. (als)"],
   [VIO, "Ziel 2 b", "Richtig oder falsch? Begründe:  „Paddy versteckt sich, weil er hat Angst.“"]].forEach(([c, n, x], i) => {
    const y = 2.7 + i * 1.0;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.12, h: 0.85, fill: { color: c }, line: { type: "none" } });
    txt(s, [{ text: n + ":  ", options: { bold: true, color: c } }, { text: x }], { x: 0.95, y: y + 0.2, w: 11.5, h: 0.55, fontSize: 19 });
  });
  zeit(s, "4 Minuten · allein im Heft, dann Lösung");
}
{
  const s = base("Auswerten · Selbstkontrolle", "Zielcheck – Lösung", "Mit Grün verbessern. Wer bei 2 a oder 2 b einen Fehler hat, schaut sich den Merksatz noch einmal an. Am Donnerstag beginnt die Wiederholung genau mit diesen Fehlern.");
  [[GREEN, "Ziel 1", "Falsch: Paddy greift nicht an. Der Schotte sieht die riesige Nase des „Babys“ und fürchtet den Vater (Str. 8–9, z. B. V. 44: „ein Narr, wer’s mit ihm selber wagt“)."],
   [VIO, "Ziel 2 a", "Der Schotte erschrickt, als er die riesige Nase sieht."],
   [VIO, "Ziel 2 b", "Falsch: Im Nebensatz steht das finite Verb am Ende → …, weil er Angst hat."]].forEach(([c, n, x], i) => {
    const y = 1.8 + i * 1.25;
    box(s, 0.6, y, 12.1, 1.05, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.12, h: 1.05, fill: { color: c }, line: { type: "none" } });
    txt(s, [{ text: n + ":  ", options: { bold: true, color: c } }, { text: x }], { x: 0.95, y: y + 0.25, w: 11.5, h: 0.65, fontSize: 20 });
  });
  txt(s, "Verbessere mit Grün. Alles richtig? Dann hast du beide Ziele erreicht.", { x: 0.6, y: 5.75, w: 12.1, h: 0.5, fontSize: 20, bold: true, align: "center" });
}
pres.writeFile({ fileName: "Mittwoch_04-11_Deutsch_7A.pptx" });
