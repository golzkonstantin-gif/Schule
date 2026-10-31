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
const Z1 = "Ich kann Fragen zum Inhalt der Ballade mit Textbeleg beantworten.";
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
  txt(s, "Deutsch 7A – Paddy Fingal: Inhalt und Grammatik", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 38, bold: true, color: WHITE });
  [["1", "Den Inhalt erschließen", GREEN], ["2", "Grammatik in der Ballade", VIO]].forEach(([n, x, c], i) => {
    box(s, 0.9 + i * 5.2, 4.1, 4.9, 0.7, c, { r: 0.35 });
    txt(s, n + "  " + x, { x: 0.9 + i * 5.2, y: 4.1, w: 4.9, h: 0.7, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}
// 2 Überblick
{
  const s = base("Überblick", "Unsere Doppelstunde: Ablauf und Ziele", "Am Anfang zeigen und am Ende noch einmal: Was haben wir geschafft? Ziele und Ablauf stehen zusätzlich an der Seitentafel.");
  const P = [
    ["1", "Paddy Fingal – den Inhalt erschließen", "ca. 40 Min.", Z1, "Wiederholung Lesetechniken · Tandemlesen · LB-Aufgabe 1 a · * LB-Aufgabe 1 b", GREEN],
    ["2", "Grammatik in der Ballade: Nebensätze mit Subjunktion", "ca. 45 Min.", Z2, "Ü2 Entdecken in der Ballade · Merksatz · Ü3 Sätze verbinden · Ü4 Fehler finden · * Ü5 zurück zur Ballade", VIO],
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
  txt(s, "* = nur, wenn die Zeit reicht. Am Ende: „Heute habe ich gelernt, …“", { x: 0.6, y: 6.4, w: 12.1, h: 0.35, fontSize: 14, italic: true, color: MUTED });
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
// LB 1a
uebung("Verarbeiten · LB-Aufgabe 1 a", "Ü1  Wovon handelt die Ballade?", [
  "Lest die Ballade im Tandem: abwechselnd Strophe für Strophe.",
  "Beantwortet die drei Fragen im Heft in ganzen Sätzen.",
  "Notiert hinter jeder Antwort die Strophe.",
], "8 Min. lesen (Tandem) + 8 Min. schreiben (allein)", "Methodenkarte Tandemlesen. Lösung: 1 Riese, schwarzer Bart sechs Ellen lang, Fäuste so groß wie Tonnen, wirft mit einem Mastbaum (Str. 1). 2 Der Schotte will herausfinden, wer von beiden stärker ist (Str. 2). 3 Schaya steckt Paddy ins Bett, türmt die Kissen hoch und gibt ihn als Baby aus; der Schotte sieht nur die riesige Nase, denkt an einen riesigen Vater und flieht (Str. 5–10).", (s, y) => {
  box(s, 0.6, y, 12.1, 1.75, LIGHT);
  ["1  Wie sieht Paddy Fingal aus?", "2  Was möchte der Riese aus Schottland von Paddy?", "3  Welche List wendet Paddys Frau an?"].forEach((q, i) => txt(s, q, { x: 0.95, y: y + 0.15 + i * 0.52, w: 11.4, h: 0.45, fontSize: 20 }));
});
// Vergleich
{
  const s = base("Auswerten · Teil 1", "Vergleich in der Kreide-Kette", "Drei Schüler schreiben je eine Antwort in Stichworten mit Strophe an die Tafel. Alle prüfen mit und verbessern mit Grün. Bei Frage 3 nachfragen: Welche Lesetechnik hast du gebraucht? (genau). Lösungen siehe Folie davor.");
  txt(s, "Wer drankommt, schreibt die Antwort in Stichworten mit Strophe an die Tafel.", { x: 0.6, y: 1.8, w: 12.1, h: 0.6, fontSize: 24, bold: true });
  txt(s, "Alle anderen prüfen mit und verbessern im Heft mit Grün.", { x: 0.6, y: 2.5, w: 12.1, h: 0.6, fontSize: 24 });
  box(s, 0.6, 3.5, 12.1, 1.6, LIGHT2);
  txt(s, [{ text: "Teil 1 – Ziel erreicht?  ", options: { bold: true, color: GREEN } }, { text: Z1, options: { italic: true } }], { x: 0.9, y: 3.75, w: 11.5, h: 1.2, fontSize: 20 });
}
// * 1b
uebung("* Verarbeiten · LB-Aufgabe 1 b", "* Ü1 b  Erzählt die Ballade nach", [
  "A erzählt Strophe 1–5 nach, B hört zu und hakt die Checkliste ab.",
  "Wechselt: B erzählt Strophe 6–11, A prüft.",
  "Nutzt eure Antworten aus Ü1 als Spickzettel.",
], "8 Minuten · Tandem", "Nur, wenn die Zeit reicht; sonst Einstieg am Donnerstag. Herumgehen, ein gelungenes Paar kurz vorerzählen lassen.", (s, y) => {
  box(s, 0.6, y, 12.1, 1.3, LIGHT);
  txt(s, [{ text: "Checkliste: ", options: { bold: true, color: GREEN } }, { text: "Titel und Autor  ·  richtige Reihenfolge  ·  nichts dazuerfinden  ·  alle wichtigen Details  ·  * spannend mit wörtlicher Rede" }], { x: 0.95, y: y + 0.2, w: 11.4, h: 0.9, fontSize: 19 });
});

// ================= TEIL 2 =================
FOOT = "Teil 2 · Grammatik in der Ballade: Nebensätze mit Subjunktion · Deutsch 7A";
section("2", "Grammatik in der Ballade", "Nebensätze mit Subjunktion", VIO);
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
  const s = base("Informieren · Entdecken in der Ballade", "Ü2  Was passiert mit dem finiten Verb?", "Lösung: dass … war · dass … erwacht · als … sah. Erstes Wort: dass, als (Subjunktion). Das finite Verb steht am Ende. Vor der Subjunktion steht ein Komma (im Gedicht teils am Versende). Überleitung: Das ist eine Regel – Merksatz-Schnipsel austeilen.");
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
  const s = base("Verarbeiten", "Ü3  Sätze verbinden", "Lösung: 1 Ich trainiere heute länger, weil am Samstag ein Wettkampf ist. 2 Der Trainer sagt, dass wir nach Frankfurt (Oder) fahren. 3 Ich war nervös, als ich auf der Matte stand. 4 Wir gewinnen, wenn wir als Team spielen. 5 Ich bin stolz, weil ich 80 kg gestoßen habe. *6 Der Trainer fragt, ob ich gut geschlafen habe. Vergleich in der Kreide-Kette: Wer drankommt, schreibt einen Satz an, alle prüfen mit.");
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
// Ü4 Fehler finden
{
  const s = base("Verarbeiten · Fehler finden", "Ü4  Was stimmt im Chat nicht?", "Lösung: 1 …, weil ich noch duschen muss. 2 …, dass wir morgen frei haben. 3 Weißt du, ob der Neue auch kommt? 4 Ich freue mich, weil wir frei haben. (Komma fehlt). Reflexion: Beim Sprechen hört man oft „weil ich muss …“ – geschrieben ist das falsch.");
  const msgs = [
    ["l", "Komme später, weil ich muss noch duschen 🙈"],
    ["r", "Trainer hat gesagt, dass wir haben morgen frei!"],
    ["l", "Weißt du, ob kommt der Neue auch?"],
    ["r", "Ich freue mich weil wir frei haben 💪"],
  ];
  box(s, 0.6, 1.55, 7.2, 4.5, "E9EEF6");
  msgs.forEach(([side, m], i) => {
    const y = 1.75 + i * 1.05;
    const x = side === "l" ? 0.85 : 2.35;
    box(s, x, y, 5.2, 0.8, side === "l" ? WHITE : "DCF3E4", { r: 0.18 });
    txt(s, (i + 1) + "  " + m, { x: x + 0.2, y: y + 0.17, w: 4.9, h: 0.5, fontSize: 17 });
  });
  [["1", "Finde in jeder Nachricht den Fehler."], ["2", "Schreibe den Satz richtig ins Heft."], ["3", "Notiere: Verbstellung oder Komma?"]].forEach(([n, x], i) => {
    num(s, n, 8.1, 1.8 + i * 1.0, VIO, 0.45);
    txt(s, x, { x: 8.7, y: 1.8 + i * 1.0, w: 4.0, h: 0.9, fontSize: 18, bold: true });
  });
  zeit(s, "5 Minuten · zu zweit");
}
// * Ü5 zurück zur Ballade
{
  const s = base("* Verarbeiten · zurück zur Ballade", "* Ü5  Erkläre die Ballade mit Nebensätzen", "Nur, wenn Zeit bleibt. Satzanfänge, keine fertigen Sätze. Mögliche Ergänzungen: … weil er Angst vor dem Schotten hat · … dass Paddy ein Baby sei/ist · … als er die riesige Nase sieht · … weil er glaubt, dass der Vater noch größer ist.");
  txt(s, "Schreibe mindestens drei Sätze ins Heft. Kreise die Subjunktion ein und unterstreiche das finite Verb.", { x: 0.6, y: 1.65, w: 12.1, h: 0.6, fontSize: 20, bold: true });
  ["Paddy versteckt sich im Bett, weil …", "Schaya erzählt dem Schotten, dass …", "Der Schotte erschrickt, als …", "Der Schotte flieht, weil …", "Die Nachbarn staunen, als …"].forEach((x, i) => {
    box(s, 0.6, 2.45 + i * 0.68, 12.1, 0.56, LIGHT);
    txt(s, x, { x: 0.9, y: 2.53 + i * 0.68, w: 11.5, h: 0.42, fontSize: 20, fontFace: HEAD });
  });
  zeit(s, "5 Minuten · allein");
}
// Auswerten: Sprachreflexion + Ziel
{
  const s = base("Auswerten", "Ziel erreicht?", "Kurz die Forscherfrage stellen (*): In Str. 3 steht „als er stapfte durch den Sund“ – das Verb steht nicht am Ende. Warum? Wegen Reim und Rhythmus (Sund – stund): Dichter dürfen von der Regel abweichen. Im Aufsatz gilt die Regel. Dann „Heute habe ich gelernt, …“ ins Heft, zwei Schüler lesen vor.");
  [[GREEN, "Teil 1", Z1], [VIO, "Teil 2", Z2]].forEach(([c, n, z], i) => {
    const y = 1.6 + i * 1.05;
    box(s, 0.6, y, 12.1, 0.9, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.12, h: 0.9, fill: { color: c }, line: { type: "none" } });
    txt(s, [{ text: n + ":  ", options: { bold: true, color: c } }, { text: z, options: { italic: true } }], { x: 0.95, y: y + 0.2, w: 11.5, h: 0.6, fontSize: 18 });
  });
  box(s, 0.6, 3.85, 12.1, 1.05, LIGHT2);
  txt(s, [{ text: "* Forscherfrage: ", options: { bold: true, color: VIO } }, { text: "In Strophe 3 steht: „Und als er stapfte durch den Sund“. Was ist hier anders – und warum macht der Dichter das?" }], { x: 0.9, y: 3.97, w: 11.6, h: 0.85, fontSize: 18 });
  txt(s, "Schreibe ins Heft:  Heute habe ich gelernt, …", { x: 0.6, y: 5.3, w: 12.1, h: 0.6, fontSize: 26, bold: true, fontFace: HEAD, align: "center" });
}
pres.writeFile({ fileName: "Mittwoch_04-11_Deutsch_7A.pptx" });
