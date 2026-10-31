const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Deutsch 7A – Freitag 6.11.";
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

// ================= Präsentation Fr 6.11.: Wiederholung Reime · Die Ballade als Ur-Ei =================
FOOT = "Freitag, 6.11. · Die Ballade als Ur-Ei · Deutsch 7A";
const ZIEL = "Ich kann an Paddy Fingal mit Belegen zeigen, dass eine Ballade Merkmale von Lyrik, Epik und Dramatik hat.";
const RC = { a: "2A8C8C", b: "C77C12", x: MUTED };
const GC = { ly: "2A8C8C", ep: "7A5BA6", dr: "C77C12" };
const zeit = (s, meta) => { box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2); txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 }); };
// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 3.2, 0.42, YEL, { r: 0.21 });
  txt(s, "FREITAG · 6.11.", { x: 0.9, y: 2.05, w: 3.2, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Paddy Fingal – die Ballade als „Ur-Ei“", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, "Gedicht, Erzählung und Theater in einem", { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
}
// 2 Ziel und Ablauf
{
  const s = base("Heute", "Ziel und Ablauf", "Ziel vorlesen; am Ende im Zielcheck darauf zurückkommen. Ziel und Ablauf stehen auch an der Seitentafel.");
  box(s, 0.6, 1.7, 12.1, 1.2, LIGHT2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 0.12, h: 1.2, fill: { color: RED }, line: { type: "none" } });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 1.0, y: 1.85, w: 11.5, h: 0.95, fontSize: 21 });
  ["Wiederholung: Ü1 Reimschema und Fachbegriffe", "Die Ballade als „Ur-Ei“", "Ü2 Merkmale an Paddy Fingal nachweisen (Schnipsel, * Nr. 8)", "Zielcheck"].forEach((a, i) => {
    num(s, i + 1, 0.9, 3.25 + i * 0.7, NAVY, 0.45);
    txt(s, a, { x: 1.55, y: 3.28 + i * 0.7, w: 10, h: 0.5, fontSize: 20 });
  });
}
// 3 Ü1 Wiederholung
{
  const s = base("Vorwissen aktivieren · Wiederholung", "Ü1  Reimschema und Fachbegriffe", "Lösung auf der nächsten Folie. Gezielt die Schüler drannehmen, die im Zielcheck am Donnerstag Fehler hatten.");
  txt(s, [{ text: "a  ", options: { bold: true, color: RED } }, { text: "Bestimme das Reimschema von Strophe 11 (V. 51–55) und nenne die Reimart." }], { x: 0.6, y: 1.55, w: 12.1, h: 0.5, fontSize: 19, bold: true });
  box(s, 0.6, 2.1, 12.1, 2.15, LIGHT);
  ["Der aber erhob ein Siegesgeschrei.", "Da kamen die Nachbarn rings herbei;", "die staunten den großen Fingal an,", "der den langen Schotten gejagt von dann,", "     den tapfern Paddy Fingal."].forEach((v, i) => txt(s, v, { x: 1.0, y: 2.2 + i * 0.4, w: 11, h: 0.38, fontSize: 19, fontFace: HEAD }));
  txt(s, [{ text: "b  ", options: { bold: true, color: RED } }, { text: "Erkläre in je einem Satz: Strophe · Vers · Kehrreim" }], { x: 0.6, y: 4.5, w: 12.1, h: 0.5, fontSize: 19, bold: true });
  zeit(s, "4 Minuten · allein im Heft");
}
{
  const s = base("Vorwissen aktivieren · Ü1 Lösung", "Ü1  Lösung", "Mit Grün verbessern.");
  [["Siegesgeschrei", "a"], ["herbei", "a"], ["an", "b"], ["dann", "b"], ["Paddy Fingal", "x"]].forEach(([w, l], i) => {
    const x = 0.6 + i * 2.45;
    box(s, x, 1.7, 2.3, 0.6, LIGHT);
    txt(s, w, { x, y: 1.75, w: 2.3, h: 0.5, fontSize: 17, fontFace: HEAD, align: "center", valign: "middle" });
    box(s, x + 0.8, 2.4, 0.7, 0.45, RC[l], { r: 0.22 });
    txt(s, l, { x: x + 0.8, y: 2.4, w: 0.7, h: 0.45, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  txt(s, [{ text: "a  ", options: { bold: true, color: RED } }, { text: "Paarreim (aabb) + Kehrreim (x)", options: { bold: true } }], { x: 0.6, y: 3.05, w: 12.1, h: 0.5, fontSize: 21 });
  box(s, 0.6, 3.7, 12.1, 2.3, LIGHT2);
  txt(s, [{ text: "b  Strophe: ", options: { bold: true, color: RED } }, { text: "ein Abschnitt der Ballade; zwischen zwei Strophen steht eine Leerzeile.\n" }, { text: "    Vers: ", options: { bold: true, color: RED } }, { text: "eine Zeile der Ballade.\n" }, { text: "    Kehrreim: ", options: { bold: true, color: RED } }, { text: "ein Vers, der sich in jeder Strophe wiederholt (hier: „… Paddy Fingal“)." }], { x: 0.9, y: 3.85, w: 11.6, h: 2.0, fontSize: 19 });
}
// Ur-Ei
{
  const s = base("Informieren · Die Ballade als „Ur-Ei“", "Goethe: Die Ballade ist ein „Ur-Ei“", "Ich-Phase: In einem Ei steckt schon alles, was später daraus wird. In der Ballade stecken alle drei großen Gattungen der Literatur. Relevanz: Diese Merkmale braucht ihr für die erste Bewertung (Analyse).");
  s.addShape(pres.shapes.OVAL, { x: 2.4, y: 1.55, w: 8.5, h: 4.9, fill: { color: LIGHT }, line: { color: NAVY, width: 2.5 } });
  s.addShape(pres.shapes.LINE, { x: 5.2, y: 1.75, w: 0, h: 4.5, line: { color: NAVY, width: 1.5, dashType: "dash" } });
  s.addShape(pres.shapes.LINE, { x: 8.1, y: 1.75, w: 0, h: 4.5, line: { color: NAVY, width: 1.5, dashType: "dash" } });
  [[GC.ly, "Lyrik", "wie ein Gedicht", 2.75], [GC.ep, "Epik", "wie eine Erzählung", 5.3], [GC.dr, "Dramatik", "wie ein Theaterstück", 8.2]].forEach(([c, n, d, x]) => {
    txt(s, n, { x, y: 3.25, w: 2.7, h: 0.6, fontSize: 28, bold: true, color: c, fontFace: HEAD, align: "center" });
    txt(s, d, { x, y: 3.9, w: 2.7, h: 0.5, fontSize: 16, color: NAVY, align: "center" });
  });
}
// Ein Training – drei Gattungen (Beispieltexte; werden am Montag wiederholt)
const BSP = [
  ["A", GC.ep, "Epik", [{ text: "Bei der letzten Runde waren meine Beine schwer wie Blei. Ich bekam kaum noch Luft und wollte aufhören. „Noch zwanzig Meter!“, rief mein Trainer vom Rand. Ich biss die Zähne zusammen und lief bis zur Linie." }], "Ein Erzähler erzählt (hier: Ich-Erzähler) · im Nachhinein, Präteritum · er gibt die Rede wieder und leitet sie ein („rief mein Trainer“)"],
  ["B", GC.dr, "Dramatik", [{ text: "TRAINER ", options: { bold: true } }, { text: "(klatscht in die Hände)", options: { italic: true } }, { text: ": Letzte Runde! Los!\n" }, { text: "SPORTLERIN ", options: { bold: true } }, { text: "(keucht)", options: { italic: true } }, { text: ": Ich kann nicht mehr …\n" }, { text: "TRAINER", options: { bold: true } }, { text: ": Noch zwanzig Meter! Du schaffst das!\n" }, { text: "(Die Sportlerin läuft bis zur Linie und lässt sich ins Gras fallen.)", options: { italic: true } }], "Kein Erzähler · die Figuren sprechen direkt miteinander, Name vor der Rede · es passiert jetzt (Präsens) · Regieanweisungen in Klammern"],
  ["C", GC.ly, "Lyrik", [{ text: "Die letzte Runde, die Beine schwer,\ndie Luft wird knapp, ich kann nicht mehr.\nEin Schritt noch, dann ist’s geschafft,\nich hol mir aus dem Bauch die Kraft." }], "Verse und Strophe · Reim (Paarreim: schwer – mehr, geschafft – Kraft) · Rhythmus"],
];
[false, true].forEach((show) => {
  const s = base("Informieren · Ein Training – drei Gattungen", show ? "So erkennst du die Gattung" : "Welcher Text ist Lyrik, Epik, Dramatik?", show
    ? "Merkmale aufdecken und mit den Antworten der Paare vergleichen. Entscheidend: Epik = ein Erzähler vermittelt (Er- oder Ich-Erzähler) und leitet Figurenrede ein; Dramatik = kein Erzähler, die Figuren sprechen unmittelbar. Wörtliche Rede gibt es in beiden – sie allein entscheidet nicht! Dieselbe Szene in drei Formen: Die Gattung hängt an der Form, nicht am Inhalt. Diese drei Texte werden am Montag wiederholt."
    : "Wir-Phase, zu zweit (3 Min.): Ordnet A, B, C zu und notiert je ein Merkmal. Dann einzelne Paare nennen lassen. Impuls: Wer erzählt hier? Gibt es überhaupt einen Erzähler?");
  BSP.forEach(([k, c, g, runs, m], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 1.55, 3.95, 0.55, show ? c : NAVY, { r: 0.08 });
    txt(s, show ? `${k}  ${g}` : `Text ${k}`, { x, y: 1.55, w: 3.95, h: 0.55, fontSize: 19, bold: true, color: WHITE, fontFace: HEAD, align: "center", valign: "middle" });
    box(s, x, 2.2, 3.95, show ? 2.45 : 3.75, LIGHT);
    txt(s, runs, { x: x + 0.15, y: 2.3, w: 3.65, h: show ? 2.25 : 3.55, fontSize: show ? 12.5 : 15, fontFace: HEAD, valign: "top" });
    if (show) {
      box(s, x, 4.75, 3.95, 1.4, WHITE, { line: c, lw: 2 });
      txt(s, m.split(" · ").map((z, j, arr) => ({ text: "• " + z + (j < arr.length - 1 ? "\n" : "") })), { x: x + 0.12, y: 4.82, w: 3.7, h: 1.3, fontSize: 12.5 });
    }
  });
  if (!show) txt(s, "Zu zweit: Ordnet A, B, C zu und notiert je ein Merkmal. Tipp: Wer erzählt hier – und gibt es überhaupt einen Erzähler?", { x: 0.6, y: 6.1, w: 12.1, h: 0.55, fontSize: 16, bold: true });
  else txt(s, "Mit Erzähler = Epik · ohne Erzähler = Dramatik · gebundene Form = Lyrik", { x: 0.6, y: 6.3, w: 12.1, h: 0.45, fontSize: 18, bold: true, align: "center" });
});
// Übertragen auf die Ballade
{
  const s = base("Informieren · Übertragen auf Paddy Fingal", "Und in der Ballade?", "Gemeinsam: Der Erzähler leitet Paddys Rede ein (Epik). In Str. 7 und 9 sprechen die Figuren eine ganze Strophe lang ohne Erzähler – wie auf der Bühne (Dramatik). Und alles in Versen mit Reim (Lyrik). Überleitung: Jetzt weist ihr das selbst nach – Schnipsel austeilen.");
  [[GC.ep, "Mit Erzähler (Epik)", "Da lief er heim in jähem Schreck:\n„O Schaya, birg mich im Versteck! …“", "Str. 4, V. 16–17 – der Erzähler leitet Paddys Rede ein"],
   [GC.dr, "Ohne Erzähler (Dramatik)", "„Tut leise, Fremder, und tretet sacht,\ndass Paddys Kindlein nicht erwacht! …“", "Str. 7 – die ganze Strophe ist Schayas Rede, kein Erzähler"]].forEach(([c, h, z, b], i) => {
    const x = 0.6 + i * 6.15;
    box(s, x, 1.6, 5.95, 0.55, c, { r: 0.08 });
    txt(s, h, { x, y: 1.6, w: 5.95, h: 0.55, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle" });
    box(s, x, 2.25, 5.95, 1.9, LIGHT);
    txt(s, z, { x: x + 0.2, y: 2.4, w: 5.55, h: 1.6, fontSize: 18, fontFace: HEAD });
    txt(s, b, { x, y: 4.25, w: 5.95, h: 0.5, fontSize: 15, italic: true, color: MUTED, align: "center" });
  });
  box(s, 0.6, 5.0, 12.1, 0.9, LIGHT2);
  txt(s, [{ text: "Und alles in Versen mit Reim: ", options: { bold: true, color: GC.ly } }, { text: "Lyrik. Die Ballade verbindet alle drei – ein „Ur-Ei“." }], { x: 0.9, y: 5.2, w: 11.6, h: 0.5, fontSize: 19 });
}
// Ü2 Schnipsel
uebung("Verarbeiten · Schnipsel", "Ü2  Merkmale an Paddy Fingal nachweisen", [
  "Fülle Nr. 1–3 (Lyrik) aus – das kennst du schon.",
  "Weise Nr. 4–7 (Epik, Dramatik) an Paddy Fingal nach.",
  "Notiere bei jedem Merkmal die Strophe.",
  "Fertig? Dann * Nr. 8.",
], "12 Minuten · allein", "Schnipsel „Die Ballade – ein Ur-Ei“ (geänderte Fassung in Woche 2: Nr. 6 und 7 geschärft) austeilen und einkleben lassen. Herumgehen. Typische Hürde: Höhepunkt (Nr. 5) – die Stelle, an der sich alles entscheidet (der Schotte steht vor dem Bett und sieht die Nase). Bei Nr. 6 Bezug zur Verb-Einheit: Woran erkennst du das Präteritum? (war, kroch, lief).");
// Aufdeck je Gattung
const AUF = [
  [GC.ly, "Lyrik – wie ein Gedicht", [["1", "11 Strophen mit je 5 Versen"], ["2", "Paarreim (aabb), z. B. Riesenart / Bart (Str. 1)"], ["3", "„… Paddy Fingal“ – der Kehrreim (wird leicht abgewandelt)"]]],
  [GC.ep, "Epik – wie eine Erzählung", [["4", "Figuren: Paddy Fingal, Schaya, der Schotte (* die Nachbarn)"], ["5", "Anfang Str. 1–3 · Höhepunkt Str. 8–9 (der Schotte sieht die Nase) · Ende Str. 10–11"], ["6", "Er-Erzähler · Präteritum („war“, „kroch“) · er leitet Figurenrede ein, z. B. Str. 4, V. 16 „Da lief er heim in jähem Schreck:“ oder Str. 6, V. 29"]]],
  [GC.dr, "Dramatik – wie ein Theaterstück", [["7", "Str. 7 (Schaya) oder Str. 9 (der Schotte): die ganze Strophe ist Figurenrede, kein Erzähler"], ["* 8", "Paddy gegen den Schotten – Frage: Findet der Schotte Paddy? Wer ist stärker?"]]],
];
AUF.forEach(([c, title, rows], k) => {
  const s = base("Auswerten · Ü2 gemeinsam prüfen", "Ü2  Lösung – " + title, "Erst einen Schüler seine Lösung mit Strophe nennen lassen, dann aufdecken. Alle verbessern mit Grün." + (k === 1 ? " Bei Nr. 5 sind andere begründete Grenzen auch richtig." : "") + (k === 2 ? " Nr. 7: Wörtliche Rede allein reicht nicht – entscheidend ist, dass kein Erzähler dazwischen ist." : ""));
  box(s, 0.6, 1.55, 12.1, 0.12, c);
  rows.forEach(([n, x], i) => {
    const y = 1.9 + i * 1.2;
    box(s, 0.6, y, 12.1, 1.0, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.25, w: 0.8, h: 0.5, fontSize: 22, bold: true, color: c });
    txt(s, x, { x: 1.6, y: y + 0.15, w: 10.9, h: 0.75, fontSize: 20, valign: "middle" });
  });
});
// Zielcheck
{
  const s = base("Auswerten", "Zielcheck: Hast du das Ziel erreicht?", "Allein im Heft, 4 Minuten, dann Lösungsfolie. D ist der Prüfstein: Hier steht zwar wörtliche Rede, aber der Erzähler leitet sie ein. Herumgehen: Wer noch Fehler hat, wird am Montag in der Wiederholung gezielt drangenommen.");
  txt(s, "Ordne jede Stelle einer Gattung zu: Lyrik, Epik oder Dramatik. Begründe mit dem Merkmal.", { x: 0.6, y: 1.5, w: 12.1, h: 0.45, fontSize: 18, bold: true });
  [["A", "Riesenart – Bart", "Str. 1"], ["B", "Da lief er heim in jähem Schreck.", "Str. 4, V. 16"], ["C", "„Ist das ein Baby, wie Ihr sagt, / ein Narr, wer’s mit ihm selber wagt!“", "Str. 9, V. 43–44"]].forEach(([k, x, st], i) => {
    const y = 2.05 + i * 0.9;
    box(s, 0.6, y, 12.1, 0.78, LIGHT);
    txt(s, k, { x: 0.85, y: y + 0.12, w: 0.6, h: 0.55, fontSize: 24, bold: true, color: RED, fontFace: HEAD });
    txt(s, x, { x: 1.6, y: y + 0.16, w: 8.9, h: 0.5, fontSize: 19, fontFace: HEAD });
    txt(s, st, { x: 10.4, y: y + 0.22, w: 2.1, h: 0.4, fontSize: 14, color: MUTED, align: "right" });
  });
  box(s, 0.6, 4.85, 12.1, 1.2, LIGHT2, { line: RED, lw: 1.5 });
  txt(s, [{ text: "D  Prüfstein: ", options: { bold: true, color: RED } }, { text: "„Da lief er heim in jähem Schreck: / O Schaya, birg mich im Versteck!“ (Str. 4, V. 16–17)\n", options: { fontFace: HEAD } }, { text: "Epik oder Dramatik? Begründe." , options: { bold: true } }], { x: 0.9, y: 4.97, w: 11.6, h: 1.0, fontSize: 18 });
  zeit(s, "4 Minuten · allein im Heft");
}
{
  const s = base("Auswerten · Selbstkontrolle", "Zielcheck – Lösung", "Mit Grün verbessern. Bei D gilt auch eine gut begründete andere Antwort, wenn das Merkmal „Erzähler“ richtig benutzt wird.");
  [["A", GC.ly, "Lyrik", "Reim: Riesenart und Bart klingen ab dem betonten Vokal gleich (Paarreim)."], ["B", GC.ep, "Epik", "Ein Erzähler erzählt, was passiert – im Präteritum („lief“)."], ["C", GC.dr, "Dramatik", "Die ganze Strophe ist Rede des Schotten – kein Erzähler dazwischen."], ["D", GC.ep, "Epik", "Zwar wörtliche Rede – aber der Erzähler leitet sie ein („Da lief er heim …:“). Er vermittelt."]].forEach(([k, c, g, x], i) => {
    const y = 1.6 + i * 1.05;
    box(s, 0.6, y, 12.1, 0.9, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.12, h: 0.9, fill: { color: c }, line: { type: "none" } });
    txt(s, [{ text: k + "  ", options: { bold: true, color: RED } }, { text: g + ": ", options: { bold: true, color: c } }, { text: x }], { x: 0.95, y: y + 0.2, w: 11.5, h: 0.55, fontSize: 19 });
  });
  txt(s, "Verbessere mit Grün. Alles richtig? Dann hast du das Ziel erreicht.", { x: 0.6, y: 5.95, w: 12.1, h: 0.45, fontSize: 19, bold: true, align: "center" });
}
pres.writeFile({ fileName: "Freitag_06-11_Deutsch_7A.pptx" });
