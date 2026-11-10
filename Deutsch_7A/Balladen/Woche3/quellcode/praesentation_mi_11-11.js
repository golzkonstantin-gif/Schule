const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Mi 11.11. · Test und John Maynard";
const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4", VIO = "7A5BA6";
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

// ================= Mi 11.11. · Doppelstunde · AVIVA =================
// Teil 1: Test „Paddy Fingal unter der Lupe“ (ca. 30 Min.) · Teil 2: John Maynard (LB S. 132–133), Aufgaben 1–3 gemeinsam
// Aufruf: node praesentation_mi_11-11.js (aus quellcode/)
FOOT = "Mittwoch, 11.11. · Deutsch 7A";
const Z1 = "Ich kann die Merkmale einer Ballade an „Paddy Fingal“ nachweisen.";
const Z2 = "Ich kann Strophe für Strophe wiedergeben, was auf der „Schwalbe“ geschieht, und mit Versen belegen.";
const zeit = (s, meta, y = 6.25, end = "Fertig? Stift hinlegen.") => {
  box(s, 0.6, y, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: end, options: { italic: true, color: MUTED } }], { x: 0.85, y: y + 0.08, w: 11.7, h: 0.4, fontSize: 16 });
};
const COL = { F: BLUE, P: GREEN, J: RED };

// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "MITTWOCH, 11.11. · DEUTSCH 7A · DOPPELSTUNDE", { x: 0.9, y: 2.0, w: 11, h: 0.4, fontSize: 14, color: ICE, charSpacing: 2 });
  txt(s, "Test · John Maynard", { x: 0.9, y: 2.5, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE });
  txt(s, "Paddy Fingal unter der Lupe – dann: Rettung in letzter Sekunde?!", { x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 22, italic: true, color: ICE });
  s.addNotes("Folie liegt beim Hereinkommen auf. Ziele und Ablauf stehen zusätzlich links an der Seitentafel.");
}

// 2 Überblick
{
  const s = base("Ankommen", "Ziele und Ablauf", "Ansage: „Stifte liegen. Ihr hört nur zu. Heute zwei Teile: Zuerst schreibt ihr den Test zu Paddy Fingal. Danach beginnen wir eine neue Ballade: John Maynard. Am Ende prüft ihr im Zielcheck, ob ihr das Ziel erreicht habt.“");
  [["1", "Test: Paddy Fingal unter der Lupe", "ca. 30 Min.", Z1, NAVY], ["2", "John Maynard (LB S. 132–133)", "ca. 55 Min.", Z2, RED]].forEach(([n, ti, m, z, col], i) => {
    const y = 1.6 + i * 2.05;
    box(s, 0.6, y, 12.1, 1.8, WHITE, { line: col, lw: 2.5 });
    num(s, n, 0.85, y + 0.22, col, 0.55);
    txt(s, ti, { x: 1.6, y: y + 0.22, w: 8.5, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true });
    txt(s, m, { x: 10.3, y: y + 0.27, w: 2.2, h: 0.45, fontSize: 17, color: MUTED, align: "right" });
    txt(s, [{ text: "Ziel: ", options: { bold: true, color: RED } }, { text: z, options: { italic: true } }], { x: 1.6, y: y + 0.85, w: 10.8, h: 0.75, fontSize: 18 });
  });
  txt(s, "Teil 2:  1 Erstes Lesen  ·  2 Lesen mit verteilten Rollen  ·  3 Was geschieht auf der „Schwalbe“?  ·  * 3d Zusammenfassen  ·  Zielcheck", { x: 0.6, y: 5.85, w: 12.1, h: 0.4, fontSize: 16, color: MUTED });
}

// 3–4 Teil 1 Test
section("1", "Test: Paddy Fingal unter der Lupe", Z1, NAVY);
{
  const s = base("Teil 1 · Test", "So läuft der Test", "Ansage: „Tische auseinander, nur ein Stift und das Lineal liegen auf dem Tisch. Ich teile Gruppe A und B aus, dazu den Text. Ihr arbeitet allein und still. Ca. 30 Minuten.“ S-PL bekommt die Fassung mit polnischer Übersetzung. Wer fertig ist: Test umdrehen, LB S. 132 aufschlagen und John Maynard still lesen (= Aufgabe 1a aus Teil 2).");
  [["1", "Lies jede Aufgabe genau. Der Text „Paddy Fingal“ liegt bei."], ["2", "Arbeite allein und still. Du hast ca. 30 Minuten."], ["3", "Fertig? Test umdrehen. LB S. 132 aufschlagen und „John Maynard“ still lesen."]].forEach(([n, x], i) => {
    const y = 1.75 + i * 0.95;
    num(s, n, 0.6, y, NAVY, 0.5);
    txt(s, x, { x: 1.3, y: y + 0.05, w: 11.3, h: 0.7, fontSize: 22, bold: true });
  });
  box(s, 0.6, 4.85, 12.1, 0.9, WHITE, { line: NAVY, lw: 2 });
  txt(s, [{ text: "Viel Erfolg!  ", options: { bold: true, color: RED } }, { text: "Denk an: Strophe, Vers und Zitat in Anführungszeichen „…“", options: { italic: true } }], { x: 0.9, y: 5.05, w: 11.6, h: 0.5, fontSize: 20 });
}

// 5 Teil 2
section("2", "John Maynard", Z2, RED);

// 6 Aufgabe 1
uebung("Teil 2 · Vorwissen aktivieren / Informieren", "Ü1  Erstes Lesen (LB S. 132–133)", [
  "Lies die Ballade „John Maynard“ ein erstes Mal still für dich.",
  "Notiere drei Textstellen, die du besonders spannend findest – jeweils mit Vers.",
  "Schreibe dazu ein Stichwort: Warum ist die Stelle spannend?",
], "8 Minuten · allein · Heft: Überschrift „John Maynard (Theodor Fontane)“", "Wer beim Test früher fertig war, hat schon gelesen und macht direkt Nr. 2. Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht auf der Folie.“ Worterklärungen stehen im LB unten (Eriesee, Gischt, Kajüte, Luke, Bugspriet).");

// 7 Ü1 gemeinsam sammeln: Countdown
{
  const s = base("Teil 2 · Ü1 gemeinsam", "Was macht die Ballade so spannend?", "Erst sammeln: Wer drankommt, nennt Vers und Stichwort. Alle ergänzen mit Grün. Dann aufdecken: Der Countdown – jede Strophe endet mit weniger Minuten bis Buffalo, gleichzeitig wird die Lage schlimmer. Weitere spannende Stellen: V. 20 „Feuer!“, V. 34–37 Funkspruch Kapitän – John Maynard, V. 41 „mit ersterbender Stimme“, V. 47 „Nur einer fehlt!“");
  const CD = [["V. 17", "Noch dreißig Minuten … halbe Stund."], ["V. 23", "Und noch zwanzig Minuten bis Buffalo."], ["V. 29", "Und noch fünfzehn Minuten bis Buffalo."], ["V. 39", "Und noch zehn Minuten bis Buffalo."], ["V. 45", "Rettung: der Strand von Buffalo!"]];
  CD.forEach(([v, x], i) => {
    const y = 1.6 + i * 0.78, w = 12.1 - i * 1.3;
    box(s, 0.6, y, w, 0.62, i === 4 ? WHITE : LIGHT, i === 4 ? { line: GREEN, lw: 2 } : {});
    txt(s, v, { x: 0.8, y: y + 0.13, w: 1.1, h: 0.4, fontSize: 16, bold: true, color: MUTED });
    txt(s, "„" + x + "“", { x: 1.9, y: y + 0.12, w: w - 1.5, h: 0.4, fontSize: 19, fontFace: HEAD, italic: true, color: i === 4 ? GREEN : NAVY });
  });
  box(s, 0.6, 5.65, 12.1, 0.75, WHITE, { line: RED, lw: 2 });
  txt(s, [{ text: "Countdown: ", options: { bold: true, color: RED } }, { text: "Die Zeit bis Buffalo wird immer kürzer – und die Lage immer gefährlicher. So entsteht Spannung." }], { x: 0.9, y: 5.8, w: 11.6, h: 0.5, fontSize: 18 });
}

// 8 Aufgabe 2 – Rollen finden
uebung("Teil 2 · Informieren", "Ü2  Wer spricht in der Ballade?", [
  "Suche alle Figuren, die in der Ballade sprechen.",
  "Notiere zu jeder Figur einen Vers, in dem sie spricht.",
  "Überlege: Welche Verse spricht der Erzähler?",
], "4 Minuten · zu zweit", "Bezug zum Test: Wo Figuren ohne Erzähler sprechen, ist die Ballade wie ein Theaterstück (Dramatik). Ansage: „Jetzt arbeitet ihr zu zweit. Vier Minuten.“");

// 9 Rollen – Lösung
{
  const s = base("Teil 2 · Ü2 Lösung", "Die Rollen", "Vergleichen und mit Grün ergänzen. Rollen verteilen (Plenum): Erzähler 2–3 Schüler im Wechsel nach Strophen, je 1–2 für die Figuren, Grabinschrift V. 58–62 liest die ganze Klasse im Chor.");
  const R = [
    ["Erzähler", "alle erzählenden Verse (z. B. V. 8–14)", NAVY],
    ["Stimme 1 / Stimme 2", "„John Maynard!“ (V. 1) – „Wer ist John Maynard?“ (V. 2) – Antwort (V. 3–7)", MUTED],
    ["Passagiere", "„Wie weit noch, Steuermann?“ (V. 15) · „Wo sind wir? wo?“ (V. 28)", GREEN],
    ["John Maynard", "„Noch dreißig Minuten …“ (V. 17) · „Ja, Herr. Ich bin.“ (V. 35) · „Ich halte drauf hin.“ (V. 37) · „Ja, Herr, ich halt’s!“ (V. 41)", RED],
    ["Kapitän", "„Noch da, John Maynard?“ (V. 34, 40) · „Auf den Strand! In die Brandung!“ (V. 36)", BLUE],
    ["Schiffsvolk", "„Feuer!“ (V. 20) · „Halt aus! Hallo!“ (V. 38)", YEL],
    ["Alle (Chor)", "Grabinschrift der Stadt (V. 58–62)", VIO],
  ];
  R.forEach(([r, v, col], i) => {
    const y = 1.55 + i * 0.66;
    box(s, 0.6, y, 12.1, 0.56, i % 2 ? WHITE : LIGHT);
    txt(s, r, { x: 0.8, y: y + 0.1, w: 2.6, h: 0.4, fontSize: 17, bold: true, color: col });
    txt(s, v, { x: 3.4, y: y + 0.1, w: 9.2, h: 0.4, fontSize: 15 });
  });
}

// 10 Rollenlesen
{
  const s = base("Teil 2 · Verarbeiten · Lesetraining", "Ü2  Lesen mit verteilten Rollen", "Methodenkarte: Lautlesetraining, Variante Rollenlesen. Ansage: „Wer eine Rolle hat, liest laut. Alle anderen lesen still mit und notieren eine Stelle, die besonders gut betont war.“ Danach kurz zwei, drei Rückmeldungen zur Betonung.");
  [["1", "Lies deine Rolle laut und betont."], ["2", "Wer keine Rolle hat: Lies still mit. Notiere eine Stelle, die besonders gut betont war."]].forEach(([n, x], i) => {
    const y = 1.7 + i * 0.85;
    num(s, n, 0.6, y, NAVY, 0.5);
    txt(s, x, { x: 1.3, y: y + 0.05, w: 11.3, h: 0.7, fontSize: 21, bold: true });
  });
  box(s, 0.6, 3.6, 12.1, 2.35, LIGHT);
  txt(s, [{ text: "Tipps zur Betonung\n", options: { bold: true, color: RED, fontSize: 19 } },
    { text: "• Der Countdown („Und noch … Minuten bis Buffalo“) wird von Strophe zu Strophe drängender.\n", options: { fontSize: 18 } },
    { text: "• John Maynard antwortet erst ruhig, in V. 41 „mit ersterbender Stimme“ – leise und schwach.\n", options: { fontSize: 18 } },
    { text: "• V. 47 „Nur einer fehlt!“ – kurze Pause davor.", options: { fontSize: 18 } }], { x: 0.9, y: 3.75, w: 11.6, h: 2.1 });
  zeit(s, "8 Minuten · Plenum", 6.25, "Methode: Rollenlesen");
}

// 11 Aufgabe 3 – Tabelle
{
  const s = base("Teil 2 · Verarbeiten", "Ü3  Was geschieht auf der „Schwalbe“?", "Gemeinsam: Tabelle ins Heft (quer über die Doppelseite oder Spalten schmal). Ich – Wir: Lehrkraft macht Strophe 2 vor (nächste Folie). Dann je Strophe: 1 Minute zu zweit Stichworte, jemand nennt, dann aufdecken, alle ergänzen mit Grün. Strophe 1 und 8 (Rahmen) zum Schluss.");
  [["1", "Zeichne die Tabelle ins Heft."], ["2", "Notiere Strophe für Strophe Stichworte in eigenen Worten."], ["3", "Steht in einer Spalte nichts? Schreibe „—“."]].forEach(([n, x], i) => {
    const y = 1.55 + i * 0.6;
    num(s, n, 0.6, y, NAVY, 0.42);
    txt(s, x, { x: 1.2, y: y + 0.02, w: 11.4, h: 0.45, fontSize: 19, bold: true });
  });
  const W = [1.4, 3.55, 3.55, 3.6], X = [0.6, 2.0, 5.55, 9.1];
  ["Strophe", "a  Fahrt", "b  Passagiere", "c  John Maynard"].forEach((h, i) => {
    box(s, X[i], 3.5, W[i], 0.55, WHITE, { line: i ? [COL.F, COL.P, COL.J][i - 1] : NAVY, lw: 2 });
    txt(s, h, { x: X[i] + 0.1, y: 3.58, w: W[i] - 0.2, h: 0.4, fontSize: 18, bold: true, color: i ? [COL.F, COL.P, COL.J][i - 1] : NAVY, align: "center" });
  });
  ["2 (V. 8–17)", "3 (V. 18–29)", "…"].forEach((r, j) => {
    const y = 4.15 + j * 0.6;
    X.forEach((x, i) => box(s, x, y, W[i], 0.5, i === 0 ? LIGHT : WHITE, { line: "C9D3EA" }));
    txt(s, r, { x: 0.65, y: y + 0.1, w: 1.3, h: 0.35, fontSize: 14, color: MUTED, align: "center" });
  });
  zeit(s, "ca. 25 Minuten · gemeinsam, Strophe für Strophe", 6.25, "Verbessern und ergänzen mit Grün.");
}

// 12–18 Aufdecken Strophe für Strophe
const STR = [
  ["2", "V. 8–17", "Ich – so geht es (Beispiel)",
    "Die „Schwalbe“ fährt über den Eriesee von Detroit nach Buffalo. Es dämmert, das Ufer ist schon zu sehen – noch 30 Minuten.",
    "Fröhlich und unbeschwert, Familien mit Kindern; sie schauen zum Ufer und fragen John Maynard, wie weit es noch ist.",
    "Er ist der Steuermann. Er schaut nach vorn und antwortet ruhig: „Noch dreißig Minuten“ (V. 17).",
    "Ansage: „Stifte liegen. Ich mache Strophe 2 vor.“ Laut denken: Was erfahre ich über die Fahrt? Ich schreibe es in eigenen Worten, nicht abschreiben. Dann: „Jetzt schreibt ihr die Zeile ab.“"],
  ["3", "V. 18–29", "Wir – zu zweit, dann aufdecken",
    "Im Schiffsraum bricht Feuer aus: Qualm und Flammen. Noch 20, dann 15 Minuten bis Buffalo.",
    "Erst noch froh; dann drängen sie sich vorne am Bugspriet zusammen, wo noch Luft ist, und jammern: „Wo sind wir? wo?“",
    "— (Indirekt: Am Steuer ist der Qualm am dichtesten, V. 27 – dort steht er.)",
    "Nachfragen: Warum ist John Maynard hier nicht zu sehen? (Er steht im dichten Qualm am Steuer.)"],
  ["4", "V. 30–39", "Wir – zu zweit, dann aufdecken",
    "Der Wind wird stärker, der Qualm bleibt. Der Kapitän befiehlt: Kurs auf den Strand, in die Brandung. Noch 10 Minuten.",
    "Alle an Bord (das Schiffsvolk) jubeln und feuern ihn an: „Halt aus! Hallo!“",
    "Man sieht ihn nicht mehr. Durchs Sprachrohr antwortet er: „Ja, Herr. Ich bin.“ – „Ich halte drauf hin.“",
    "Wörter: Sprachrohr (Rohr zum Rufen über Entfernung), Brandung (Wellen, die sich am Ufer brechen)."],
  ["5", "V. 40–45", "Wir – zu zweit, dann aufdecken",
    "Die „Schwalbe“ wird mitten in die Brandung gejagt, über Klippen und Steine – Rettung am Strand von Buffalo.",
    "— (sie warten auf die Rettung)",
    "Mit ersterbender Stimme: „Ja, Herr, ich halt’s!“ Er steuert das Schiff in die Brandung.",
    "„Mit ersterbender Stimme“: Er ist schon schwer verletzt bzw. stirbt gerade."],
  ["6", "V. 46–47", "Wir – zu zweit, dann aufdecken",
    "Das Schiff ist zerbrochen, das Feuer erlischt.",
    "Alle sind gerettet.",
    "Er fehlt als Einziger – er ist tot.",
    "Kurze Strophe = Pause, Schock. Nachfragen: Warum nur zwei Verse?"],
  ["7", "V. 48–57", "Wir – zu zweit, dann aufdecken",
    "— (die Fahrt ist vorbei)",
    "Die Stadt trauert: Glocken läuten, über 10 000 Menschen folgen dem Sarg, alle weinen. Der Sarg wird in Blumen begraben.",
    "Er wird beerdigt; die Stadt schreibt einen Dankspruch in goldener Schrift auf den Marmorstein.",
    "Wörter: Kapelle (kleine Kirche), Sarg, tränenleer (ohne Tränen) – „kein Aug’ … das tränenleer“ = alle weinen."],
  ["1 + 8", "V. 1–7 / V. 58–62", "Der Rahmen: Anfang = Ende",
    "— (Rückblick nach dem Unglück)",
    "Die Geretteten (V. 1–7) bzw. die Stadt (V. 58–62) erinnern an ihn und danken ihm.",
    "Er war der Steuermann, hielt im Qualm das Steuer fest, hat alle gerettet und ist dabei gestorben – „er trägt die Kron’“: Er ist ein Held.",
    "Erkenntnis: Strophe 1 und Strophe 8 sagen fast dasselbe (Rahmen). Am Anfang wissen wir schon das Ende – trotzdem bleibt es spannend, weil wir wissen wollen, wie es dazu kam."],
];
STR.forEach(([st, v, k, f, pa, j, note]) => {
  const s = base("Teil 2 · Ü3 · " + k, `Strophe ${st}  (${v})`, note);
  [["a  Fahrt", f, COL.F], ["b  Passagiere", pa, COL.P], ["c  John Maynard", j, COL.J]].forEach(([h, x, col], i) => {
    const y = 1.6 + i * 1.55;
    box(s, 0.6, y, 12.1, 1.38, WHITE, { line: col, lw: 2 });
    txt(s, h, { x: 0.85, y: y + 0.12, w: 3.0, h: 0.4, fontSize: 18, bold: true, color: col });
    txt(s, x, { x: 0.85, y: y + 0.52, w: 11.6, h: 0.8, fontSize: 18, color: x.startsWith("—") ? MUTED : NAVY });
  });
  txt(s, "Vergleicht und ergänzt mit Grün.", { x: 0.6, y: 6.35, w: 12.1, h: 0.35, fontSize: 15, italic: true, color: MUTED });
});

// 19 * 3d Zusammenfassen
{
  const s = base("Teil 2 · * Ü3 d", "* Fasse den Inhalt zusammen", "Gemeinsam, Satz für Satz: Lehrkraft schreibt mit (Folie/Tafel), alle schreiben ins Heft. Hinweis: Eine Inhaltszusammenfassung steht im Präsens. Wenn die Zeit nicht reicht: Einstieg in die nächste Stunde.");
  box(s, 0.6, 1.55, 12.1, 0.75, WHITE, { line: RED, lw: 2 });
  txt(s, [{ text: "Start: ", options: { bold: true, color: RED } }, { text: "Die Ballade „John Maynard“ von Theodor Fontane handelt von einem Schiffsunglück, bei dem der Steuermann …", options: { italic: true } }], { x: 0.85, y: 1.68, w: 11.6, h: 0.55, fontSize: 18 });
  const SW = [["Str. 1", "die Geretteten – erinnern – Steuermann"], ["Str. 2", "„Schwalbe“ – Eriesee – fröhlich – 30 Minuten"], ["Str. 3", "Feuer – Qualm – Bugspriet"], ["Str. 4–5", "Kapitän – Sprachrohr – Brandung – hält aus"], ["Str. 6", "alle gerettet – einer fehlt"], ["Str. 7–8", "Trauer – Begräbnis – Dankspruch"]];
  SW.forEach(([a, b], i) => {
    const y = 2.5 + i * 0.5;
    txt(s, a, { x: 0.8, y, w: 1.5, h: 0.4, fontSize: 17, bold: true, color: MUTED });
    txt(s, b, { x: 2.3, y, w: 10.3, h: 0.4, fontSize: 17 });
  });
  box(s, 0.6, 5.6, 12.1, 0.6, LIGHT2);
  txt(s, [{ text: "Verknüpfen: ", options: { bold: true } }, { text: "zuerst · dann · als · weil · obwohl · schließlich      " }, { text: "Zeitform: Präsens", options: { bold: true, color: RED } }], { x: 0.85, y: 5.7, w: 11.6, h: 0.4, fontSize: 17 });
}
// 20 * 3d Musterlösung
{
  const s = base("Teil 2 · * Ü3 d Lösung", "* So könnte es aussehen", "Vergleichen: Steht alles Wichtige drin, in der richtigen Reihenfolge, im Präsens? Andere gute Formulierungen gelten.");
  box(s, 0.6, 1.55, 12.1, 4.6, LIGHT);
  txt(s, "Die Ballade „John Maynard“ von Theodor Fontane handelt von einem Schiffsunglück, bei dem der Steuermann John Maynard alle Passagiere rettet und selbst stirbt. In der ersten Strophe wird berichtet, dass die Geretteten sich an ihn als ihren Retter erinnern. Die zweite Strophe erzählt vom Schiff „Schwalbe“, das über den Eriesee nach Buffalo fährt. Die Passagiere an Bord sind fröhlich, und John Maynard befindet sich am Steuer. Dann bricht ein Feuer aus, und die Passagiere drängen sich vorne am Schiff zusammen. Obwohl der Qualm immer dichter wird, hält John Maynard das Steuer fest und lenkt das Schiff in die Brandung. Schließlich sind alle gerettet, nur John Maynard ist tot. Die Stadt trauert um ihn und ehrt ihn mit einem Dankspruch auf seinem Grab.", { x: 0.9, y: 1.75, w: 11.5, h: 4.25, fontSize: 18, valign: "top", lineSpacingMultiple: 1.15 });
}

// 21 Zielcheck
const ZC = [
  ["1", "Das Feuer bricht aus, als die „Schwalbe“ noch etwa zwanzig Minuten von Buffalo entfernt ist.", true, "V. 20–23: „Feuer!“ … „Und noch zwanzig Minuten bis Buffalo.“"],
  ["2", "John Maynard verlässt das Steuer, als der Qualm zu dicht wird.", false, "V. 41: „Ja, Herr, ich halt’s!“ / V. 58–59: „In Qualm und Brand hielt er das Steuer fest in der Hand“ – richtig ist: Er bleibt am Steuer."],
  ["3", "Alle Menschen an Bord überleben das Unglück.", false, "V. 47: „Gerettet alle. Nur einer fehlt!“ – richtig ist: John Maynard stirbt."],
];
[false, true].forEach((show) => {
  const s = base("Auswerten · Zielcheck", show ? "Zielcheck: Lösung – Ziel erreicht?" : "Zielcheck", show
    ? "Mit Grün verbessern. Abschlussfrage: „Haben wir das Ziel erreicht?“ Herumgehen und notieren, wer noch Fehler hat – in der Wiederholung der nächsten Stunde aufgreifen."
    : "Allein im Heft, 4 Minuten. Ansage: „Zielcheck. Jetzt arbeitet ihr allein im Heft.“");
  txt(s, show ? [{ text: "Ziel erreicht?  ", options: { bold: true, color: RED } }, { text: Z2, options: { italic: true } }] : "Richtig oder falsch? Schreibe r oder f und den Vers als Beleg. Verbessere falsche Aussagen.", { x: 0.6, y: 1.5, w: 12.1, h: 0.5, fontSize: show ? 17 : 19, bold: !show });
  ZC.forEach(([n, a, ok, b], i) => {
    const y = 2.2 + i * 1.3;
    box(s, 0.6, y, 12.1, 1.15, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.15, w: 0.5, h: 0.45, fontSize: 20, bold: true, color: VIO });
    txt(s, "„" + a + "“", { x: 1.35, y: y + 0.12, w: 9.6, h: 0.45, fontSize: 18 });
    if (show) {
      box(s, 11.1, y + 0.15, 1.4, 0.42, ok ? GREEN : RED, { r: 0.2 });
      txt(s, ok ? "richtig" : "falsch", { x: 11.1, y: y + 0.15, w: 1.4, h: 0.42, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, b, { x: 1.35, y: y + 0.6, w: 11.2, h: 0.5, fontSize: 14, color: MUTED });
    }
  });
  if (!show) zeit(s, "4 Minuten · allein im Heft");
});

pres.writeFile({ fileName: __dirname + "/../Mittwoch_11-11_Deutsch_7A.pptx" });
