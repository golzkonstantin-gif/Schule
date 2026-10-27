const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Generalstände 8D";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Generalstände · Geschichte 8D";
let pageNo = 1;

function txt(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 18, color: NAVY, margin: 0, valign: "top" }, o)); }
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: o.r ?? 0.08, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" } });
}
function numCircle(s, n, x, y, d = 0.5, col = NAVY) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function footer(s) {
  txt(s, FOOT, { x: 0.6, y: 7.12, w: 6, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo), { x: 12.2, y: 7.12, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
}
function bar(s, runs) {
  box(s, 0.6, 6.2, 12.15, 0.85, NAVY);
  s.addText(runs, { isTextBox: true, x: 0.9, y: 6.2, w: 11.6, h: 0.85, fontFace: BODY, fontSize: 24, color: WHITE, valign: "middle", margin: 0 });
}

// Aufgabenfolie: gleiches Format wie immer (Nummer + Titel, nummerierte Schritte, Zeit, Arbeitsform)
function taskSlide(kicker, title, stepsArr, meta, notes, done = "Fertig? Stift hinlegen und noch einmal durchlesen.", stepW = 12.15) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.35, w: 8, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const per = stepW > 10 ? 56 : 38;
  const hs = stepsArr.map((st) => (st.length > per * 2 ? 1.4 : st.length > per ? 0.95 : 0.5));
  const total = hs.reduce((a, b) => a + b, 0) + (hs.length - 1) * 0.08;
  box(s, 0.6, 1.5, stepW, total + 0.24, LIGHT2);
  let yy = 1.62;
  stepsArr.forEach((st, i) => {
    numCircle(s, i + 1, 0.8, yy + 0.04, 0.42);
    txt(s, st, { x: 1.4, y: yy, w: stepW - 0.95, h: hs[i], fontSize: 24, bold: true });
    yy += hs[i] + 0.08;
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      ", options: {} }, { text: done, options: { italic: true, color: ICE } }]);
  footer(s);
  if (notes) s.addNotes(notes);
  return { s, top: 1.5 + total + 0.24 + 0.2 };
}
function sentences(s, list, top, o = {}) {
  const rows = list.length, avail = 6.1 - top, rh = Math.min(o.rh || 0.85, avail / rows);
  list.forEach((sent, i) => {
    const y = top + i * rh;
    box(s, 0.6, y, 12.15, rh - 0.08, LIGHT);
    numCircle(s, i + 1, 0.75, y + (rh - 0.08 - 0.4) / 2, 0.4, RED);
    s.addText(sent, { isTextBox: true, x: 1.4, y, w: 11.2, h: rh - 0.08, fontFace: HEAD, fontSize: o.size || 24, color: NAVY, valign: "middle", margin: 0 });
  });
}

// Kurzsyntax: **fett**, [r:rot], [g:grün], [b:blau]
function rich(str, base = {}) {
  const out = [];
  str.split(/(\*\*[^*]+\*\*|\[[rgb]:[^\]]+\])/).filter(Boolean).forEach((tok) => {
    let m;
    if ((m = tok.match(/^\*\*([^*]+)\*\*$/))) out.push({ text: m[1], options: Object.assign({}, base, { bold: true }) });
    else if ((m = tok.match(/^\[([rgb]):([^\]]+)\]$/))) out.push({ text: m[2], options: Object.assign({}, base, { bold: true, color: { r: RED, g: GREEN, b: BLUE }[m[1]] }) });
    else out.push({ text: tok, options: Object.assign({}, base) });
  });
  return out;
}
function merkSlide(nr, title, kern, beispiel, test, notes, kernSize = 24) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, `MERKSATZ ${nr} · ABSCHREIBEN UND ROT UMRAHMEN`, { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: RED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.5, w: 12.15, h: 2.1, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  s.addText(rich(kern), { isTextBox: true, x: 0.95, y: 1.55, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 24, color: NAVY, valign: "middle", margin: 0 });
  box(s, 0.6, 3.72, 12.15, 1.25, LIGHT2);
  s.addText([{ text: "Beispiel:   ", options: { bold: true, color: MUTED, fontFace: BODY, fontSize: 24 } }, ...rich(beispiel)], { isTextBox: true, x: 0.9, y: 3.72, w: 11.6, h: 1.25, fontFace: HEAD, fontSize: 24, color: NAVY, valign: "middle", margin: 0 });
  box(s, 0.6, 5.08, 12.15, 0.97, LIGHT);
  s.addText([{ text: "Test:   ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 24 } }, ...rich(test)], { isTextBox: true, x: 0.9, y: 5.08, w: 11.6, h: 0.97, fontFace: BODY, fontSize: 24, color: NAVY, valign: "middle", margin: 0 });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schreibe den Merksatz ab. Rahme ihn rot.      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
  footer(s);
  s.addNotes(notes || `Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz ${nr}. Rahmt ihn rot ein.“ Warten, bis alle den Stift hingelegt haben.`);
  return s;
}



// ---------- Hilfsfunktionen für diese Stunde ----------
const KLERUS = "8E6BBE", ADEL = "E8A33D", DRITTER = "3F7CC4", NEIN = "D9534F", JA = "2E9E6B";
const AVIVA = ["A", "V", "I", "V", "A"];
function phaseTag(s, idx) {
  AVIVA.forEach((l, i) => {
    const x = 10.0 + i * 0.55;
    s.addShape(pres.shapes.OVAL, { x, y: 0.3, w: 0.44, h: 0.44, fill: { color: i === idx ? RED : LIGHT2 }, line: { type: "none" } });
    txt(s, l, { x, y: 0.3, w: 0.44, h: 0.44, fontSize: 14, bold: true, color: i === idx ? WHITE : MUTED, align: "center", valign: "middle" });
  });
}
function blank(kicker, title, idx, notes) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.35, w: 9, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  if (idx !== undefined) phaseTag(s, idx);
  txt(s, title, { x: 0.6, y: 0.7, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  footer(s);
  if (notes) s.addNotes(notes);
  return s;
}
const legend = (s, x, y) => [["Klerus", KLERUS], ["Adel", ADEL], ["Dritter Stand (Bauern und Bürger)", DRITTER]].reduce((xx, [n, c]) => {
  s.addShape(pres.shapes.RECTANGLE, { x: xx, y: y + 0.08, w: 0.3, h: 0.3, fill: { color: c }, line: { type: "none" } });
  txt(s, n, { x: xx + 0.4, y, w: 8, h: 0.45, fontSize: 24 });
  return xx + 0.4 + (n.length > 10 ? 5.2 : 1.6);
}, x);
function stackBar(s, x, y, w, h, parts) {
  const tot = parts.reduce((a, b) => a + b[0], 0);
  let xx = x;
  parts.forEach(([v, col, label]) => {
    const ww = (w * v) / tot;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y, w: ww, h, fill: { color: col }, line: { color: WHITE, width: 1.5 } });
    if (label && ww > 1.0) txt(s, label, { x: xx, y, w: ww, h, fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle" });
    xx += ww;
  });
}
const cardRow = (s, y, list) => list.forEach(([n, sub, c], i) => {
  const xx = 0.6 + i * 4.1;
  box(s, xx, y, 3.95, 1.5, LIGHT);
  s.addShape(pres.shapes.RECTANGLE, { x: xx, y, w: 0.14, h: 1.5, fill: { color: c }, line: { type: "none" } });
  txt(s, n, { x: xx + 0.3, y: y + 0.08, w: 3.5, h: 0.75, fontFace: HEAD, fontSize: 24, bold: true, valign: "middle" });
  txt(s, sub, { x: xx + 0.3, y: y + 0.85, w: 3.5, h: 0.6, fontSize: 24, color: MUTED });
});
pageNo = 0;

// ============ 1 Titel / Ankommen ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "Generalstände 1789", { x: 0.9, y: 1.3, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  txt(s, "Wie können die Stände den Dritten Stand überstimmen?", { x: 0.9, y: 2.4, w: 11.5, h: 0.6, fontSize: 26, italic: true, color: ICE });
  box(s, 0.9, 3.7, 11.5, 2.5, NAVY2);
  txt(s, "Ankommen", { x: 1.2, y: 3.85, w: 6, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true, color: WHITE });
  ["Hole dein Heft heraus.", "Schreibe Datum und Überschrift „Generalstände 1789“.", "Auf deinem Platz liegt eine Karte. Lass sie verdeckt liegen."].forEach((x, i) => {
    numCircle(s, i + 1, 1.2, 4.55 + i * 0.52, 0.4, RED);
    txt(s, x, { x: 1.85, y: 4.52 + i * 0.52, w: 10.3, h: 0.46, fontSize: 24, color: WHITE, valign: "middle" });
  });
  txt(s, "Geschichte, Klasse 8D · Französische Revolution", { x: 0.9, y: 6.7, w: 8, h: 0.3, fontSize: 13, color: ICE });
  s.addNotes("Die Folie steht beim Betreten. Niemand muss etwas sagen. Tests liegen verdeckt auf den Plätzen (falls noch nicht zurückgegeben), keine Besprechung. Die Karten (Standeskarten) liegen verdeckt auf den Plätzen. Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2 und 3.");
}

// ============ 2 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Lernziele vorlesen (1 Minute, Klasse hört nur zu). Am Ende der Stunde (Folie 16) prüft die Klasse still, ob sie erreicht sind. Was nicht geschafft wird, hat einen Stern.");
  [["1", "Ich kann den Weg bis zu den Generalständen am Zeitstrahl erklären."], ["2", "Ich kann erklären, wie die Stände den Dritten Stand überstimmen konnten."], ["3", "Ich kann begründen, warum die Abstimmung nach Ständen unfair war."]].forEach(([n, x], i) => {
    const y = 1.75 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.37, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 26, valign: "middle" });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Zeitstrahl und Ständeversammlung (Spiel).", options: {} }]);
}

// ============ 3 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Ablauf zeigen. Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Ü3 Satz 3).");
  const ph = [["A", "Ankommen", "Ziele", "3 Min."], ["V", "Vorwissen", "Stände", "6 Min."], ["I", "Informieren", "Text und Zeitstrahl", "11 Min."], ["V", "Verarbeiten", "Spiel", "18 Min."], ["A", "Auswerten", "Merksatz, Abschlussfrage", "6 Min."]];
  ph.forEach(([l, n, x, m], i) => {
    const xx = 0.6 + i * 2.45;
    box(s, xx, 1.9, 2.3, 3.6, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fill: { color: RED }, line: { type: "none" } });
    txt(s, l, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, n, { x: xx, y: 3.1, w: 2.3, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true, align: "center" });
    txt(s, x, { x: xx + 0.1, y: 3.7, w: 2.1, h: 1.0, fontSize: 24, align: "center", color: MUTED });
    txt(s, m, { x: xx, y: 4.8, w: 2.3, h: 0.5, fontSize: 24, bold: true, color: RED, align: "center" });
  });
  bar(s, [{ text: "Wichtig: ", options: { bold: true } }, { text: "Wir arbeiten leise. Gesprochen wird nur, wenn es auf der Folie steht.", options: {} }]);
}

// ============ 4 Ü1 Fehlertext ============
{
  const s = blank("V · Vorwissen aktivieren · Übung 1", "Fehlertext: Absolutismus und Stände", 1, "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Kein Plenumsgespräch. Vier Sätze enthalten je einen Fehler, ein Satz ist richtig. Fehler: 1 (der König regiert allein, nicht mit den Ständen), 2 (die Macht kommt von Gott, Gottesgnadentum), 3 (der Klerus hat etwa 150 000 Mitglieder, nicht 20 Millionen), 4 (Klerus und Adel zahlen keine Steuern). Satz 5 ist richtig. Die Zahlen stammen aus dem Gruppenpuzzle (150 000 / 500 000 / 20 000 000). Wiederholung der letzten Stunde (Test). Die Generalstände kennt die Klasse noch nicht, sie kommen deshalb nicht vor. Nach 4 Minuten Folie 5 zeigen.");
  box(s, 0.6, 1.5, 12.15, 3.55, LIGHT2);
  const sat = ["Ludwig XIV. war der Sonnenkönig und regierte im Absolutismus gemeinsam mit den Ständen.", "Seine Macht komme vom Volk, sagte er (Gottesgnadentum).", "Zum Klerus gehörten etwa 20 Millionen Menschen, zum Adel etwa 500 000.", "Klerus und Adel zahlten die meisten Steuern.", "Der Dritte Stand (Bauern und Bürger) hatte kaum Vorrechte."];
  sat.forEach((x, i) => {
    numCircle(s, i + 1, 0.85, 1.7 + i * 0.66, 0.4, NAVY);
    txt(s, x, { x: 1.5, y: 1.62 + i * 0.66, w: 11.1, h: 0.58, fontFace: HEAD, fontSize: 24, valign: "middle" });
  });
  [["1", "Vier Sätze enthalten einen Fehler. Unterstreiche die Fehler."], ["2", "Schreibe diese Sätze richtig ins Heft."]].forEach(([n, x], i) => {
    numCircle(s, n, 0.85, 5.3 + i * 0.45, 0.38, RED);
    txt(s, x, { x: 1.5, y: 5.25 + i * 0.45, w: 11.1, h: 0.45, fontSize: 24, bold: true, valign: "middle" });
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "4 Minuten · allein · leise      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 5 Lösung Ü1 ============
{
  const s = blank("V · Vorwissen aktivieren · Kontrolle", "Fehlertext: Lösung", 1, "Ansage: „Vergleicht und verbessert mit Grün.“ 2 Minuten, nicht besprechen. Die korrigierten Stellen stehen grün. Satz 5 war richtig. Kurz zum Mitnehmen für die nächste Phase: Klerus und Adel zahlen keine Steuern, der Dritte Stand (Bauern und Bürger) zahlt sie. Bei Unklarheit eine gezielte Abfrage (2 bis 3 Personen), kein Gespräch.");
  box(s, 0.6, 1.5, 12.15, 4.45, LIGHT2);
  const G = { bold: true, color: GREEN };
  const lines = [
    [{ text: "Ludwig XIV. war der Sonnenkönig und regierte im Absolutismus " }, { text: "allein, ohne die Stände", options: G }, { text: "." }],
    [{ text: "Seine Macht komme " }, { text: "von Gott", options: G }, { text: ", sagte er (Gottesgnadentum)." }],
    [{ text: "Zum Klerus gehörten etwa " }, { text: "150 000", options: G }, { text: " Menschen, zum Adel etwa 500 000." }],
    [{ text: "Klerus und Adel zahlten " }, { text: "keine Steuern", options: G }, { text: "." }],
    [{ text: "Der Dritte Stand (Bauern und Bürger) hatte kaum Vorrechte. " }, { text: "(war richtig)", options: { italic: true, color: MUTED } }],
  ];
  lines.forEach((runs, i) => {
    numCircle(s, i + 1, 0.85, 1.75 + i * 0.85, 0.4, NAVY);
    s.addText(runs.map((r) => ({ text: r.text, options: Object.assign({ fontFace: HEAD, fontSize: 24, color: NAVY }, r.options || {}) })), { isTextBox: true, x: 1.5, y: 1.65 + i * 0.85, w: 11.0, h: 0.8, valign: "middle", margin: 0 });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 6 Ü2 ============
{
  const { s } = taskSlide("I · Informieren · Übung 2", "Text und Zeitstrahl", ["Lies den Text. Unterstreiche die Jahreszahlen.", "Schreibe zu jedem Datum das Ereignis in den Zeitstrahl.", "Prüfe: Passt jedes Ereignis zum Datum?"], "6 Minuten · allein · leise",
    "Ansage: „Ich teile den Text aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Oben der Text, unten der Zeitstrahl mit vorgegebenen Daten und leeren Ereignisfeldern. Danach Kreide-Kette: „Wer drankommt, schreibt an die Tafel. Alle prüfen mit. Gesprochen wird nicht.“ 7 Personen, Datum steht an der Tafel.");
  phaseTag(s, 2);
}

// ============ 7 Zeitstrahl ============
{
  const s = blank("I · Informieren · Kontrolle", "Der Weg zu den Generalständen", 2, "Erst nach der Kreide-Kette zeigen. „Vergleicht und verbessert mit Grün.“ 1 Minute. Die Frage „Warum ruft der König die Generalstände ein, obwohl er sie 175 Jahre nicht brauchte?“ steht im Text und wird nicht im Plenum besprochen. Antwort: Geldnot, Reformen scheitern, Krise (Hunger, drohender Bankrott).");
  const ev = [["1614", "Letzte Generalstände"], ["1778–1783", "Krieg in Amerika, Schulden"], ["1787", "Steuerplan scheitert"], ["1788", "Missernte, Brotpreis"], ["August 1788", "König kündigt an"], ["Frühjahr 1789", "Wahlen, Beschwerdehefte"], ["5. Mai 1789", "Eröffnung in Versailles"]];
  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 3.72, w: 11.7, h: 0.1, fill: { color: NAVY }, line: { type: "none" } });
  txt(s, "▶", { x: 12.4, y: 3.5, w: 0.5, h: 0.5, fontSize: 24, color: NAVY, align: "center", valign: "middle" });
  ev.forEach(([d, e], i) => {
    const cx = 1.85 + i * 1.6;
    const up = i % 2 === 0;
    s.addShape(pres.shapes.OVAL, { x: cx - 0.17, y: 3.6, w: 0.34, h: 0.34, fill: { color: i === ev.length - 1 ? RED : NAVY }, line: { color: WHITE, width: 2 } });
    s.addShape(pres.shapes.LINE, { x: cx, y: up ? 3.0 : 3.95, w: 0, h: 0.6, line: { color: MUTED, width: 1.5 } });
    const y = up ? 1.5 : 4.55;
    box(s, cx - 1.55, y, 3.1, 1.45, up ? LIGHT2 : LIGHT);
    txt(s, d, { x: cx - 1.5, y: y + 0.08, w: 3.0, h: 0.5, fontSize: 24, bold: true, color: RED, align: "center" });
    txt(s, e, { x: cx - 1.5, y: y + 0.6, w: 3.0, h: 0.85, fontFace: HEAD, fontSize: 24, align: "center" });
  });
}

// ============ 8 Ständeversammlung: Antrag ============
{
  const s = blank("V · Verarbeiten · Ständeversammlung", "Der Antrag", 3, "Ansage: „Stifte liegen. Ihr hört nur zu. Dreht jetzt eure Karte um. Lest sie leise. Zeigt sie niemandem.“ Dann Regel vorsprechen: Jeder Stand hat eine Stimme (siehe Text). Dann: „Erster und Zweiter Stand, ihr geht jetzt leise an den Tisch vorn. Dritter Stand, ihr bleibt sitzen.“ Absprache-Zettel auf den Tisch vorn legen. Das ist die einzige Bewegung der Stunde.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.65, w: 12.15, h: 1.7, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  txt(s, "Antrag: Auch Adel und Klerus zahlen künftig Steuern.", { x: 0.95, y: 1.7, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 32, bold: true, valign: "middle" });
  cardRow(s, 3.6, [["Klerus", "3 Personen · 1 Stimme", KLERUS], ["Adel", "4 Personen · 1 Stimme", ADEL], ["Dritter Stand (Bauern und Bürger)", "18 Personen · 1 Stimme", DRITTER]]);
  bar(s, [{ text: "Regel: ", options: { bold: true } }, { text: "Jeder Stand hat eine Stimme (Abstimmung nach Ständen).", options: {} }]);
}

// ============ 9 Aufträge ============
{
  const s = blank("V · Verarbeiten · Beratung", "Eure Aufträge (4 Minuten)", 3, "Ansage: „Flüsterstufe. Nur dein Partner hört dich. Ihr habt vier Minuten.“ Klerus und Adel füllen gemeinsam den Absprache-Zettel aus. Der Dritte Stand arbeitet zu zweit mit dem Sitznachbarn. Die Beobachtung ist der Zielpunkt: Dem Dritten Stand nicht sagen, dass es eine Absprache ist. Eine Minute vor Ende: „Noch eine Minute.“");
  box(s, 0.6, 1.65, 5.95, 4.35, LIGHT2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.65, w: 0.14, h: 4.35, fill: { color: ADEL }, line: { type: "none" } });
  txt(s, "Erster und Zweiter Stand", { x: 0.95, y: 1.75, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true });
  ["Füllt gemeinsam den Absprache-Zettel aus.", "Wie stimmt ihr ab?", "Was versprecht ihr euch?"].forEach((x, i) => {
    numCircle(s, i + 1, 0.95, 2.55 + i * 1.1, 0.42);
    txt(s, x, { x: 1.55, y: 2.5 + i * 1.1, w: 4.8, h: 0.95, fontSize: 24, bold: true });
  });
  box(s, 6.8, 1.65, 5.95, 4.35, LIGHT);
  s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 1.65, w: 0.14, h: 4.35, fill: { color: DRITTER }, line: { type: "none" } });
  txt(s, "Dritter Stand (Bauern und Bürger)", { x: 7.15, y: 1.75, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true });
  ["Schreibe zwei Argumente auf: Warum sollen auch Adel und Klerus Steuern zahlen?", "Beobachte den Tisch vorn. Schreibe eine Beobachtung auf: Was tun Klerus und Adel?"].forEach((x, i) => {
    numCircle(s, i + 1, 7.15, 2.55 + i * 1.7, 0.42);
    txt(s, x, { x: 7.75, y: 2.5 + i * 1.7, w: 4.8, h: 1.55, fontSize: 24, bold: true });
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "4 Minuten · Flüsterstufe      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 10 Abstimmung ============
{
  const { s } = taskSlide("V · Verarbeiten · Abstimmung", "Jetzt stimmt ihr ab", ["Kreuze auf dem Stimmzettel unten auf deiner Karte Ja oder Nein an.", "Lege die Karte verdeckt auf den Tisch.", "Schreibe auf: Wer gewinnt nach Ständen? Wer gewinnt nach Köpfen? Ein Satz mit Begründung."], "3 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Jetzt stimmt ihr ab. Kreuzt auf dem Stimmzettel unten auf eurer Karte Ja oder Nein an. Legt die Karte verdeckt auf den Tisch. Danach arbeitet ihr allein.“ Klerus und Adel stimmen nach Rolle, der Dritte Stand, wie er es für richtig hält. Karten einsammeln, in drei Stapel sortieren (Klerus, Adel, Dritter Stand), währenddessen schreibt die Klasse die Vermutung. Ergebnis (echte Zahlen!) an die Tafel schreiben.");
  phaseTag(s, 3);
}

// ============ 11 Zwei Regeln ============
{
  const s = blank("V · Verarbeiten · Auswertung", "Zwei Regeln für dieselben Stimmen", 3, "Ansage: „Stifte liegen. Ihr hört nur zu. Erster und Zweiter Stand, ihr geht leise an eure Plätze zurück.“ Die Zahlen schreiben Sie an die Tafel (Tafelbild links), die Folie zeigt nur die Regel. Je ein Satz: „Nach Ständen: abgelehnt, 2 zu 1.“ „Nach Köpfen: angenommen.“ Dann den Absprache-Zettel vorlesen (unter einer Minute).");
  box(s, 0.6, 1.65, 5.95, 4.35, LIGHT2);
  txt(s, "Nach Ständen", { x: 0.9, y: 1.75, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY });
  txt(s, "Jeder Stand hat eine Stimme.", { x: 0.9, y: 2.4, w: 5.4, h: 0.5, fontSize: 24 });
  [["Klerus", KLERUS], ["Adel", ADEL], ["Dritter Stand", DRITTER]].forEach(([n, c], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.9 + i * 1.8, y: 3.2, w: 1.65, h: 1.2, fill: { color: c }, line: { type: "none" } });
    txt(s, n, { x: 0.9 + i * 1.8, y: 3.2, w: 1.65, h: 1.2, fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  txt(s, "3 Stimmen insgesamt", { x: 0.9, y: 4.7, w: 5.4, h: 0.5, fontSize: 24, color: MUTED });
  box(s, 6.8, 1.65, 5.95, 4.35, LIGHT2);
  txt(s, "Nach Köpfen", { x: 7.1, y: 1.75, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY });
  txt(s, "Jede Person hat eine Stimme.", { x: 7.1, y: 2.4, w: 5.4, h: 0.5, fontSize: 24 });
  for (let i = 0; i < 25; i++) {
    const c = i < 3 ? KLERUS : i < 7 ? ADEL : DRITTER;
    s.addShape(pres.shapes.OVAL, { x: 7.1 + (i % 9) * 0.58, y: 3.2 + Math.floor(i / 9) * 0.58, w: 0.46, h: 0.46, fill: { color: c }, line: { type: "none" } });
  }
  txt(s, "25 Stimmen insgesamt", { x: 7.1, y: 4.9, w: 5.4, h: 0.5, fontSize: 24, color: MUTED });
  bar(s, [{ text: "Ergebnis: ", options: { bold: true } }, { text: "Zwei Regeln. Dieselben Stimmen. Zwei Ergebnisse.", options: {} }]);
}

// ============ 12 Und in Frankreich? ============
{
  const s = blank("V · Verarbeiten · Frankreich", "Und in Frankreich?", 3, "Verhältnisse in Frankreich: Klerus ca. 150 000, Adel ca. 500 000, Dritter Stand (Bauern und Bürger) ca. 20 000 000 Menschen. Zahlen aus dem Gruppenpuzzle-Material. Abgeordnete: Klerus ca. 290, Adel ca. 270, Dritter Stand ca. 580 (bitte gegen das Schulbuch prüfen). Das Spiel ist verzerrt: In der Klasse sind Klerus und Adel 28 %, in Frankreich ca. 3 %. Die Stimme des Klerus wiegt rund 133-mal so viel (20 000 000 : 150 000), die des Adels 40-mal. Zweiter kurzer Gesprächsmoment, unter einer Minute: „Wer will, schreibt jetzt auf, was ihn am meisten überrascht hat.“ Keine Antworten ins Plenum.");
  legend(s, 0.6, 1.5);
  [["Bevölkerung", [[0.73, KLERUS, ""], [2.42, ADEL, ""], [96.85, DRITTER, "ca. 97 %"]], "Klerus ca. 1 % · Adel ca. 2 %"], ["Abgeordnete", [[290, KLERUS, "ca. 290"], [270, ADEL, "ca. 270"], [580, DRITTER, "ca. 580"]], ""], ["Stimmen bei der Abstimmung nach Ständen", [[1, KLERUS, "1"], [1, ADEL, "1"], [1, DRITTER, "1"]], ""]].forEach(([l, parts, note], i) => {
    const y = 2.25 + i * 1.3;
    txt(s, l, { x: 0.6, y, w: 8, h: 0.45, fontFace: HEAD, fontSize: 24, bold: true });
    if (note) txt(s, note, { x: 5.0, y, w: 7.7, h: 0.45, fontSize: 24, color: MUTED, align: "right" });
    stackBar(s, 0.6, y + 0.5, 12.15, 0.65, parts);
  });
  bar(s, [{ text: "Gewicht: ", options: { bold: true } }, { text: "Eine Stimme des Klerus wiegt rund 133-mal so viel.", options: {} }]);
}

// ============ 13 Ü3 ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 3", "Was ist passiert?", ["Der Dritte Stand wurde überstimmt, weil …", "Klerus und Adel haben …", "* Die Abstimmung nach Ständen ist unfair, weil … (nenne eine Zahl)"], "4 Minuten · allein · leise",
    "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Nach 3 Minuten: „Noch eine Minute.“ Satz 3 hat einen Stern und darf entfallen. Erwartung: 1 … jeder Stand nur eine Stimme hat. 2 … sich abgesprochen und gemeinsam mit Nein gestimmt. 3 … der Dritte Stand ca. 97 % der Menschen stellt, aber nur 1 von 3 Stimmen hat.");
  phaseTag(s, 3);
}

// ============ 14 Merksatz ============
{
  const s = merkSlide(1, "Abstimmung nach Ständen",
    "Bei den Generalständen hat jeder Stand **eine Stimme**. Klerus und Adel (zusammen ca. 3 % der Bevölkerung) haben damit **zwei von drei Stimmen**. Sprechen sie sich ab, können sie den Dritten Stand (Bauern und Bürger, ca. 97 %) **immer überstimmen**.",
    "Im Spiel: 18 Personen im Dritten Stand hatten zusammen **eine** Stimme, 3 Personen im Klerus ebenfalls.",
    "Wie viele Menschen stehen hinter **einer Stimme**?",
    "Ansage: „Jetzt schreibt ihr ab: Merksatz 1. Rahmt ihn rot ein.“ Danach „Vergleicht und verbessert mit Grün.“", 24);
  phaseTag(s, 4);
}

// ============ 15 Abschlussfrage (Berater) ============
// Sprechblase: Rechteck mit Dreieck-Spitze (Spitze zeigt zur Figur)
function bubble(s, x, y, w, h, fill, line, side, tipY) {
  box(s, x, y, w, h, fill, { line, lw: 3, r: 0.18 });
  const L = 0.42, B = 0.5;
  if (side === "left") {
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x - L / 2 - B / 2 + 0.06, y: tipY - L / 2, w: B, h: L, rotate: 270, fill: { color: fill }, line: { color: line, width: 3 } });
    s.addShape(pres.shapes.RECTANGLE, { x: x - 0.02, y: tipY - B / 2 + 0.05, w: 0.1, h: B - 0.1, fill: { color: fill }, line: { type: "none" } });
  } else {
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + w + L / 2 - B / 2 - 0.06, y: tipY - L / 2, w: B, h: L, rotate: 90, fill: { color: fill }, line: { color: line, width: 3 } });
    s.addShape(pres.shapes.RECTANGLE, { x: x + w - 0.08, y: tipY - B / 2 + 0.05, w: 0.1, h: B - 0.1, fill: { color: fill }, line: { type: "none" } });
  }
}
{
  const s = blank("A · Auswerten", "Abschlussfrage: Du bist Berater", 4, "Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Zehn Sekunden denken, dann schreiben. Es gibt mehrere gute Antworten, geprüft wird die Begründung: eine Position, ein Sachargument aus der Stunde, ein „weil“. Erwartung (Auswahl, keine Musterlösung): FÜR Erlauben: Der Staat braucht Geld, Steuern gehen nur mit dem Dritten Stand (Bauern und Bürger); er stellt fast alle Menschen, sonst droht Unruhe; Hunger und Staatsbankrott sind nah. GEGEN Erlauben: Adel und Klerus sind wütend (der Steuerplan scheiterte 1787), der König verliert Rückhalt und Macht; der Dritte Stand wäre zu mächtig; seit 1614 wurde nach Ständen abgestimmt. Auch eine Zwischenlösung ist gut begründbar. Form nach Verlauf: Blitzlicht (je eine Stimme „erlauben“, eine „nicht erlauben“, gezielt aufgerufen), schriftlich mit Vorlesen oder Berater-Duell zu zweit. Stichprobe: bei 4 bis 6 Heften unterschiedlichen Niveaus mitlesen, nicht kommentieren (unbenotet). Die Figuren sind Zeichnungen, keine historischen Porträts.");
  s.addImage({ path: "../material/koenig.png", x: 0.55, y: 1.45, w: 3.1, h: 4.0 });
  txt(s, "König Ludwig XVI.", { x: 0.55, y: 5.5, w: 3.1, h: 0.4, fontFace: HEAD, fontSize: 24, bold: true, align: "center" });
  bubble(s, 4.2, 1.5, 8.55, 1.95, WHITE, NAVY, "left", 2.5);
  txt(s, "„Soll ich erlauben, dass der Dritte Stand (Bauern und Bürger) nach Köpfen abstimmt?“", { x: 4.5, y: 1.5, w: 8.0, h: 1.95, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle" });
  bubble(s, 4.2, 3.75, 6.9, 1.6, "FDF1D8", GOLD, "right", 4.55);
  txt(s, "Dein Rat:  „Ich rate dir, … , weil …“", { x: 4.5, y: 3.75, w: 6.4, h: 1.6, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle" });
  s.addImage({ path: "../material/berater.png", x: 11.75, y: 3.4, w: 1.3, h: 1.8 });
  txt(s, "Du", { x: 11.75, y: 5.2, w: 1.3, h: 0.4, fontFace: HEAD, fontSize: 24, bold: true, align: "center" });
  [["Position", 2.3], ["Sachargument", 3.3], ["„weil“", 2.1]].reduce((xx, [lab, ww], i) => {
    box(s, xx, 5.55, ww, 0.62, i % 2 ? LIGHT : LIGHT2);
    s.addText([{ text: lab, options: { bold: true, fontFace: HEAD, fontSize: 24, color: NAVY } }], { isTextBox: true, x: xx + 0.1, y: 5.55, w: ww - 0.2, h: 0.62, valign: "middle", align: "center", margin: 0 });
    return xx + ww + 0.2;
  }, 4.2);
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "3 Minuten · allein · leise      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 16 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Die Klasse prüft still: Kann ich das? Danach Folie 17. Ziel 3 hat einen Stern (*): Es kann sein, dass es nicht mehr geschafft wurde, dann kommt es nächste Stunde als Wiederholung. Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen.");
  [["Ich kann den Weg bis zu den Generalständen am Zeitstrahl erklären."], ["Ich kann erklären, wie die Stände den Dritten Stand überstimmen konnten."], ["Ich kann begründen, warum die Abstimmung nach Ständen unfair war.*"]].forEach(([x], i) => {
    const y = 1.7 + i * 1.4;
    box(s, 0.6, y, 12.15, 1.15, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.15, fontFace: HEAD, fontSize: 26, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Haben wir sie erreicht?", options: {} }]);
}

// ============ 16 Ausblick ============
{
  const s = blank("A · Auswerten · Ausblick", "Wie geht es weiter?", 4, "Folie nur lesen (1 Minute). Ausblick: Was passiert, wenn der Dritte Stand verlangt, dass nach Köpfen abgestimmt wird, und der König ablehnt? Hausaufgabe nur, wenn der Test zurückgegeben wurde: Falsche Aufgaben mit Grün berichtigen.");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Der Dritte Stand fordert:\n„Wir stimmen nach Köpfen ab!“", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Wie reagieren Klerus, Adel und der König? Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Hausaufgabe: ", options: { bold: true } }, { text: "Testaufgaben mit Grün berichtigen.", options: {} }]);
}

// ============ 17 Zusatz ============
{
  const s = blank("Zusatz · nur bei Zeit*", "Ein Sportverein stimmt ab", undefined, "Nur bei Zeit (Puffer). Einzelarbeit im Heft, 2 Minuten. Erwartung: Die Abteilung Fußball (200 Mitglieder) hat nur eine von drei Stimmen, Tischtennis (20) und Schach (10) können sie überstimmen. Das entspricht der Abstimmung nach Ständen. Fairer: nach Köpfen.");
  box(s, 0.6, 1.7, 12.15, 1.4, LIGHT2);
  txt(s, "Ein Verein hat drei Abteilungen: Fußball (200 Mitglieder), Tischtennis (20), Schach (10). Jede Abteilung hat bei der Abstimmung eine Stimme.", { x: 0.9, y: 1.7, w: 11.6, h: 1.4, fontFace: HEAD, fontSize: 24, valign: "middle" });
  ["Schreibe auf: Was ist an dieser Abstimmung unfair?", "Schreibe auf: Wie wäre die Abstimmung fairer?"].forEach((x, i) => {
    numCircle(s, i + 1, 0.9, 3.55 + i * 1.0, 0.42);
    txt(s, x, { x: 1.55, y: 3.5 + i * 1.0, w: 10.8, h: 0.55, fontSize: 24, bold: true });
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "2 Minuten · allein · leise      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

pres.writeFile({ fileName: "../Praesentation_Generalstaende_8D.pptx" }).then(() => console.log("ok"));
