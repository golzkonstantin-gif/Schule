const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Stunde 2: Der Dritte Stand handelt";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Der Dritte Stand handelt · Geschichte 8C";
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
  s.addText(runs, { isTextBox: true, x: 0.9, y: 6.2, w: 11.6, h: 0.85, fontFace: BODY, fontSize: 27, color: WHITE, valign: "middle", margin: 0 });
}

// Aufgabenfolie: gleiches Format wie immer (Nummer + Titel, nummerierte Schritte, Zeit, Arbeitsform)
function taskSlide(kicker, title, stepsArr, meta, notes, done = "Fertig? Stift hinlegen.", stepW = 12.15) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.35, w: 8, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const per = stepW > 10 ? 50 : 34;
  const hs = stepsArr.map((st) => (st.length > per * 2 ? 1.4 : st.length > per ? 0.95 : 0.5));
  const total = hs.reduce((a, b) => a + b, 0) + (hs.length - 1) * 0.08;
  box(s, 0.6, 1.5, stepW, total + 0.24, LIGHT2);
  let yy = 1.62;
  stepsArr.forEach((st, i) => {
    numCircle(s, i + 1, 0.8, yy + 0.04, 0.42);
    txt(s, st, { x: 1.4, y: yy, w: stepW - 0.95, h: hs[i], fontSize: 27, bold: true });
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
    s.addText(sent, { isTextBox: true, x: 1.4, y, w: 11.2, h: rh - 0.08, fontFace: HEAD, fontSize: o.size || 27, color: NAVY, valign: "middle", margin: 0 });
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
function merkSlide(nr, title, kern, beispiel, test, notes, kernSize = 27) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, `MERKSATZ ${nr} · ABSCHREIBEN UND ROT UMRAHMEN`, { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: RED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.5, w: 12.15, h: 2.1, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  s.addText(rich(kern), { isTextBox: true, x: 0.95, y: 1.55, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle", margin: 0 });
  box(s, 0.6, 3.72, 12.15, 1.25, LIGHT2);
  s.addText([{ text: "Beispiel:   ", options: { bold: true, color: MUTED, fontFace: BODY, fontSize: 27 } }, ...rich(beispiel)], { isTextBox: true, x: 0.9, y: 3.72, w: 11.6, h: 1.25, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle", margin: 0 });
  box(s, 0.6, 5.08, 12.15, 0.97, LIGHT);
  s.addText([{ text: "Test:   ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 27 } }, ...rich(test)], { isTextBox: true, x: 0.9, y: 5.08, w: 11.6, h: 0.97, fontFace: BODY, fontSize: 27, color: NAVY, valign: "middle", margin: 0 });
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
  txt(s, n, { x: xx + 0.4, y, w: 8, h: 0.45, fontSize: 27 });
  return xx + 0.4 + (n.length > 10 ? 5.2 : 1.6);
}, x);
function stackBar(s, x, y, w, h, parts) {
  const tot = parts.reduce((a, b) => a + b[0], 0);
  let xx = x;
  parts.forEach(([v, col, label]) => {
    const ww = (w * v) / tot;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y, w: ww, h, fill: { color: col }, line: { color: WHITE, width: 1.5 } });
    if (label && ww > 1.0) txt(s, label, { x: xx, y, w: ww, h, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
    xx += ww;
  });
}
const cardRow = (s, y, list) => list.forEach(([n, sub, c], i) => {
  const xx = 0.6 + i * 4.1;
  box(s, xx, y, 3.95, 1.5, LIGHT);
  s.addShape(pres.shapes.RECTANGLE, { x: xx, y, w: 0.14, h: 1.5, fill: { color: c }, line: { type: "none" } });
  txt(s, n, { x: xx + 0.3, y: y + 0.08, w: 3.5, h: 0.75, fontFace: HEAD, fontSize: 27, bold: true, valign: "middle" });
  txt(s, sub, { x: xx + 0.3, y: y + 0.85, w: 3.5, h: 0.6, fontSize: 27, color: MUTED });
});
pageNo = 0;


// Sprechblase (Rechteck mit Dreieck-Spitze)
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
const IMG = "../material/";

// ============ 1 Titel / Ankommen ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "Der Dritte Stand handelt", { x: 0.9, y: 1.1, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  txt(s, "Nationalversammlung und Ballhausschwur 1789", { x: 0.9, y: 2.2, w: 11.5, h: 0.6, fontSize: 26, italic: true, color: ICE });
  box(s, 0.9, 3.3, 11.5, 3.0, NAVY2);
  txt(s, "Ankommen", { x: 1.2, y: 3.4, w: 6, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true, color: WHITE });
  ["Hole dein Heft heraus.", "Schreibe Datum und Überschrift „Der Dritte Stand handelt“.", "Darf sich das Volk gegen einen König wehren, der allein regiert? Schreibe: Ich finde …, weil …"].forEach((x, i) => {
    numCircle(s, i + 1, 1.2, 4.1 + i * 0.7, 0.4, RED);
    txt(s, x, { x: 1.85, y: 4.05 + i * 0.7, w: 10.3, h: 0.5, fontSize: 27, color: WHITE, valign: "middle" });
  });
  txt(s, "Geschichte, Klasse 8C · Französische Revolution · Stunde 2", { x: 0.9, y: 6.8, w: 8, h: 0.3, fontSize: 13, color: ICE });
  s.addNotes("Die Folie steht beim Betreten. Niemand muss etwas sagen. Die Meinung zu „Darf sich das Volk gegen einen König wehren, der allein regiert?“ bleibt im Heft (Eingangsdiagnose ohne Fachbegriff, weil die Klasse „Revolution“ noch nicht kennt; Vergleich in Stunde 11). Es gibt keine richtige Antwort, geprüft wird nichts. Lehrkraft 2 hospitiert nur. Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2 und 3.");
}

// ============ 2 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Lernziele vorlesen (1 Minute, die Klasse hört nur zu). Am Ende (Folie 16) prüft die Klasse still, ob sie erreicht sind. Der Stern (*) bei Ziel 3 bedeutet: Das Beurteilen der Karikatur wird bei Zeitnot nur angeleitet gemeinsam gemacht.");
  [["1", "Ich kann erklären, warum sich der Dritte Stand (Bauern und Bürger) 1789 zur Nationalversammlung erklärte."], ["2", "Ich kann erklären, was die Abgeordneten im Ballhausschwur versprachen."], ["3", "Ich kann eine Karikatur analysieren: beschreiben, deuten, beurteilen.*"]].forEach(([n, x], i) => {
    const y = 1.7 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.3, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.4, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.3, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Nationalversammlung, Ballhausschwur und eine Karikatur.", options: {} }]);
}

// ============ 3 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Ablauf zeigen. Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Beurteilen der Karikatur).");
  const ph = [["A", "Ankommen", "Ziele", "3 Min."], ["V", "Vorwissen", "Abstimmung", "6 Min."], ["I", "Informieren", "Text, Zeitstrahl", "12 Min."], ["V", "Verarbeiten", "Karikatur", "15 Min."], ["A", "Auswerten", "Berater-Frage", "6 Min."]];
  ph.forEach(([l, n, x, m], i) => {
    const xx = 0.6 + i * 2.45;
    box(s, xx, 1.9, 2.3, 3.6, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fill: { color: RED }, line: { type: "none" } });
    txt(s, l, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, n, { x: xx, y: 3.1, w: 2.3, h: 0.5, fontFace: HEAD, fontSize: 27, bold: true, align: "center" });
    txt(s, x, { x: xx + 0.1, y: 3.7, w: 2.1, h: 1.0, fontSize: 27, align: "center", color: MUTED });
    txt(s, m, { x: xx, y: 4.8, w: 2.3, h: 0.5, fontSize: 27, bold: true, color: RED, align: "center" });
  });
  bar(s, [{ text: "Wichtig: ", options: { bold: true } }, { text: "Wir arbeiten konzentriert. Gesprochen wird nach Ansage.", options: {} }]);
}

// ============ 4 Ü1 Fehlertext ============
{
  const s = blank("V · Vorwissen aktivieren · Übung 1", "Fehlertext: Abstimmung bei den Generalständen", 1, "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Vier Sätze enthalten einen Fehler, ein Satz ist richtig. Fehler: 1 (jeder Stand hat eine Stimme, nicht jeder Abgeordnete), 3 (der Dritte Stand stellt etwa 97 %), 4 (nach Köpfen hätte er mehr Gewicht), 5 (der König lehnte ab). Satz 2 ist richtig. Wiederholung der letzten Stunde. Lehrkraft 2 hospitiert nur. Nach 3 Minuten allein: Partnercheck (1 Minute, leise), dann Folie 5. Hinweis: Die Abstimmungsmöglichkeiten nach Ständen und nach Köpfen sind laut Lehrkraft noch nicht sicher.");
  box(s, 0.6, 1.5, 12.15, 3.55, LIGHT2);
  const sat = ["Bei den Generalständen hat jeder Abgeordnete eine Stimme.", "Klerus und Adel haben zusammen zwei von drei Stimmen.", "Der Dritte Stand (Bauern und Bürger) stellt etwa 3 % der Menschen.", "Bei der Abstimmung nach Köpfen hätte er weniger Gewicht gehabt.", "Der König erlaubte die Abstimmung nach Köpfen."];
  sat.forEach((x, i) => {
    numCircle(s, i + 1, 0.85, 1.7 + i * 0.66, 0.4, NAVY);
    txt(s, x, { x: 1.5, y: 1.62 + i * 0.66, w: 11.1, h: 0.58, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  [["1", "Vier Sätze enthalten einen Fehler. Unterstreiche die Fehler."], ["2", "Schreibe diese Sätze richtig ins Heft."]].forEach(([n, x], i) => {
    numCircle(s, n, 0.85, 5.3 + i * 0.45, 0.38, RED);
    txt(s, x, { x: 1.5, y: 5.25 + i * 0.45, w: 11.1, h: 0.45, fontSize: 27, bold: true, valign: "middle" });
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "4 Minuten · allein, dann Partnercheck · leise      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 5 Lösung Fehlertext ============
{
  const s = blank("V · Vorwissen aktivieren · Kontrolle", "Fehlertext: Lösung", 1, "Ansage: „Vergleicht und verbessert mit Grün.“ Die korrigierten Stellen stehen grün. Satz 2 war richtig. Gezielte Abfrage (2 bis 3 Personen, kein Gespräch): „Wie viele Stimmen hat der Dritte Stand nach Ständen? Und nach Köpfen?“");
  box(s, 0.6, 1.5, 12.15, 4.45, LIGHT2);
  const G = { bold: true, color: GREEN };
  const lines = [
    [{ text: "Bei den Generalständen hat jeder " }, { text: "Stand", options: G }, { text: " eine Stimme." }],
    [{ text: "Klerus und Adel haben zusammen zwei von drei Stimmen. " }, { text: "(war richtig)", options: { italic: true, color: MUTED } }],
    [{ text: "Der Dritte Stand (Bauern und Bürger) stellt etwa " }, { text: "97 %", options: G }, { text: " der Menschen." }],
    [{ text: "Bei der Abstimmung nach Köpfen hätte er " }, { text: "mehr", options: G }, { text: " Gewicht gehabt." }],
    [{ text: "Der König " }, { text: "lehnte", options: G }, { text: " die Abstimmung nach Köpfen " }, { text: "ab", options: G }, { text: "." }],
  ];
  lines.forEach((runs, i) => {
    numCircle(s, i + 1, 0.85, 1.75 + i * 0.85, 0.4, NAVY);
    s.addText(runs.map((r) => ({ text: r.text, options: Object.assign({ fontFace: HEAD, fontSize: 27, color: NAVY }, r.options || {}) })), { isTextBox: true, x: 1.5, y: 1.65 + i * 0.85, w: 11.0, h: 0.8, valign: "middle", margin: 0 });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 6 Zwei Regeln ============
{
  const s = blank("V · Vorwissen aktivieren · Merkbild", "Nach Ständen oder nach Köpfen?", 1, "Kurz erklären (1 Minute), das Bild trägt die Wiederholung. Gedankenspiel aus Stunde 1 (Antrag: Auch Adel und Klerus zahlen Steuern): Nach Ständen stimmen Klerus Nein, Adel Nein, Dritter Stand Ja: 2 zu 1, abgelehnt. Nach Köpfen Ja 580, Nein 290 + 270 = 560: knapp angenommen. Wichtig: Es wurde 1789 nicht nach Köpfen abgestimmt, der Streit um die Regel blockierte die Generalstände. Gezielte Abfrage statt Gespräch.");
  box(s, 0.6, 1.5, 5.95, 4.5, LIGHT2);
  txt(s, "Nach Ständen", { x: 0.9, y: 1.6, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY });
  txt(s, "Jeder Stand hat eine Stimme.", { x: 0.9, y: 2.25, w: 5.4, h: 0.5, fontSize: 27 });
  [["Klerus", KLERUS, "Nein"], ["Adel", ADEL, "Nein"], ["Dritter Stand", DRITTER, "Ja"]].forEach(([n, c, v], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.9 + i * 1.8, y: 2.85, w: 1.65, h: 1.6, fill: { color: c }, line: { type: "none" } });
    txt(s, n + "\n" + v, { x: 0.9 + i * 1.8, y: 2.85, w: 1.65, h: 1.6, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  txt(s, "2 : 1  →  abgelehnt", { x: 0.9, y: 4.75, w: 5.4, h: 0.7, fontFace: HEAD, fontSize: 28, bold: true, color: RED });
  box(s, 6.8, 1.5, 5.95, 4.5, LIGHT2);
  txt(s, "Nach Köpfen", { x: 7.1, y: 1.6, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY });
  txt(s, "Jeder Abgeordnete: eine Stimme.", { x: 7.1, y: 2.25, w: 5.4, h: 0.5, fontSize: 27 });
  stackBar(s, 7.1, 2.85, 5.35, 1.6, [[580, JA, "Ja 580"], [560, NEIN, "Nein 560"]]);
  txt(s, "580 : 560  →  knapp angenommen", { x: 7.1, y: 4.75, w: 5.4, h: 1.1, fontFace: HEAD, fontSize: 28, bold: true, color: RED });
  bar(s, [{ text: "Gedankenspiel: ", options: { bold: true } }, { text: "1789 wurde nicht nach Köpfen abgestimmt.", options: {} }]);
}

// ============ 7 Ü2 ============
{
  const { s } = taskSlide("I · Informieren · Übung 2", "Text, Quelle und Zeitstrahl", ["Lies den Text. Unterstreiche alle Daten.", "Schreibe zu jedem Datum das Ereignis in den Zeitstrahl.", "Schreibe ins Heft: Sieyès sagt, der Dritte Stand sei ALLES. Erkläre das mit einer Zahl."], "8 Minuten · allein · leise",
    "Ansage: „Ich teile das Arbeitsblatt aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Oben der Text mit der Quelle von Sieyès, unten der Zeitstrahl mit vorgegebenen Daten und leeren Ereignisfeldern. Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung, historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Material vorab auf den Plätzen. Erwartung zu 3: Der Dritte Stand stellt etwa 97 % der Menschen (ca. 20 Mio. von rund 20,65 Mio.), deshalb sagt Sieyès, er sei ALLES.");
  phaseTag(s, 2);
}

// ============ 8 Zeitstrahl ============
{
  const s = blank("I · Informieren · Kontrolle", "Der Dritte Stand handelt", 2, "Erst nach dem Partnervergleich zeigen (1 Minute, „Vergleicht und verbessert mit Grün.“). Mögliche Kreide-Kette bei 12 Schülern, nur wenn Zeit ist. Quelle: Sieyès (verfasst 1788, erschienen Winter 1788/89).");
  const ev = [["Winter 1788/89", "Sieyès: Der Dritte Stand ist ALLES"], ["5. Mai 1789", "Eröffnung, Streit um die Abstimmung"], ["17. Juni 1789", "Nationalversammlung"], ["20. Juni 1789", "Ballhausschwur"], ["Anfang Juli 1789", "Soldaten um Paris, Angst in Paris"]];
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 1.7, w: 0.1, h: 4.35, fill: { color: NAVY }, line: { type: "none" } });
  ev.forEach(([d, e], i) => {
    const y = 1.5 + i * 0.9;
    box(s, 0.6, y, 2.8, 0.75, i % 2 ? LIGHT : LIGHT2);
    txt(s, d, { x: 0.7, y, w: 2.6, h: 0.75, fontSize: 27, bold: true, color: RED, align: "center", valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: y + 0.2, w: 0.34, h: 0.34, fill: { color: i === 3 ? RED : NAVY }, line: { color: WHITE, width: 2 } });
    txt(s, e, { x: 4.2, y, w: 8.5, h: 0.75, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
}

// ============ 9 Bild Ballhausschwur ============
{
  const s = blank("I · Informieren · Bild", "Der Ballhausschwur", 2, "Bild zeigen (1 Minute), keine Besprechung. Quelle: Zeichnung von Jacques-Louis David, 1791 (Lehrbuch S. 113, Q4). Das Buch weist darauf hin, dass die Zeichnung nicht den tatsächlichen Hergang wiedergibt: Im Vordergrund verbrüdern sich ein Mönch, ein katholischer Priester und ein protestantischer Pfarrer. Der Zeichner steht auf der Seite der Revolution. Hinweis für die Klasse: Hier sieht man, wie die Abgeordneten ihren Schwur dargestellt haben wollten.");
  s.addImage({ path: IMG + "ballhausschwur.png", x: 0.7, y: 1.5, w: 7.0, h: 4.81 });
  box(s, 8.6, 1.6, 4.15, 4.5, LIGHT2);
  txt(s, "Zeichnung von David, 1791", { x: 8.8, y: 1.7, w: 3.8, h: 0.9, fontFace: HEAD, fontSize: 27, bold: true });
  txt(s, "Die Abgeordneten schwören: Wir gehen nicht auseinander, bis Frankreich eine Verfassung hat.", { x: 8.8, y: 2.8, w: 3.8, h: 3.2, fontSize: 27 });
  txt(s, "Lehrbuch S. 113", { x: 0.7, y: 6.4, w: 7.0, h: 0.3, fontSize: 14, color: MUTED });
}

// ============ 10 Merksatz ============
{
  const s = merkSlide(1, "Nationalversammlung und Ballhausschwur",
    "Am **17. Juni 1789** erklärt sich der Dritte Stand (Bauern und Bürger) zur **Nationalversammlung**. Am **20. Juni** schwören die Abgeordneten im Ballhaus, nicht auseinanderzugehen, bis Frankreich eine **Verfassung** hat. Damit stellt der Dritte Stand die **Macht des Königs** infrage.",
    "Sieyès: Der Dritte Stand ist **ALLES**, bisher war er **NICHTS**, er will **ETWAS** sein.",
    "Wer erklärte sich zur Nationalversammlung? Was **schwor** sie?",
    "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“ Danach „Vergleicht und verbessert mit Grün.“ Warnung vor dem Fehler: 17. Juni (Nationalversammlung) und 20. Juni (Ballhausschwur) nicht verwechseln.", 24);
  phaseTag(s, 2);
}

// ============ 11 Methode Dreischritt ============
{
  const s = blank("V · Verarbeiten · Methode", "So analysiere ich eine Karikatur", 3, "Methode direkt (Portion 1, 1 Minute): Thema und Ziel nennen. Eine Karikatur ist ein Bild, das übertreibt, um eine Meinung zu zeigen. Wir gehen immer in drei Schritten vor: beschreiben, deuten, beurteilen. Die Schnipsel mit Satzanfängen liegen auf den Tischen (Buch S. 116: Historienbilder, gleiche Struktur). Warnung vor dem Fehler: Nicht sofort deuten, erst genau beschreiben.");
  [["1", "Beschreiben", "Was sehe ich? Personen, Dinge, Beischrift.", LIGHT2], ["2", "Deuten", "Was soll es bedeuten? Wofür steht die Figur?", LIGHT], ["3", "Beurteilen", "Stimmt das? Wo übertreibt es? Was will der Zeichner?", LIGHT2]].forEach(([n, a, b, f], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 1.7, 3.95, 3.9, f);
    numCircle(s, n, xx + 0.2, 1.9, 0.6, RED);
    txt(s, a, { x: xx + 0.2, y: 2.7, w: 3.55, h: 0.6, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY });
    txt(s, b, { x: xx + 0.2, y: 3.4, w: 3.55, h: 2.0, fontSize: 27 });
  });
  bar(s, [{ text: "Merke: ", options: { bold: true } }, { text: "Erst genau beschreiben, dann deuten, dann beurteilen.", options: {} }]);
}

// ============ 12 Modell ============
{
  const s = blank("V · Verarbeiten · Modell", "Ich zeige es euch: Karikatur 1789", 3, "Modell (3 Minuten, die Lehrkraft denkt laut, Klasse hört nur zu, „Stifte liegen“). Das Bild kennt die Klasse aus dem Absolutismus (Beischrift im Dialekt: „Man muss hoffen, dass dieses Spiel bald endet.“, 1789). 1 Beschreiben: „Ich sehe einen alten Bauern mit Sense, gebückt. Auf seinem Rücken sitzen ein Geistlicher mit Kreuz und ein Adliger mit Degen und Federhut. Unten fressen Hasen am Kohl, Hühner picken.“ 2 Deuten: „Der Bauer steht für den Dritten Stand (Bauern und Bürger). Er trägt Klerus und Adel: Er zahlt die Steuern, die beiden anderen nicht. Die Hasen am Kohl stehen für das Jagdrecht des Adels, das Wild frisst die Felder der Bauern. Die Beischrift sagt, dass das Spiel bald enden soll.“ 3 Beurteilen: „Der Zeichner steht auf der Seite des Dritten Standes. Er übertreibt, es gab auch arme Adlige und Geistliche. Aber er trifft etwas Richtiges: Klerus und Adel zahlten kaum Steuern. Er will, dass sich etwas ändert, 1789, vor den Generalständen.“ Warnung vor dem Fehler: nicht nur nennen, was man sieht, sondern deuten.");
  s.addImage({ path: IMG + "karikatur_drei_staende.jpg", x: 0.7, y: 1.45, w: 3.75, h: 4.65 });
  [["1", "Beschreiben", "Bauer, Geistlicher, Adliger, Hasen"], ["2", "Deuten", "Der Dritte Stand trägt die Last"], ["3", "Beurteilen", "Er übertreibt, trifft aber den Kern"]].forEach(([n, a, b], i) => {
    const y = 1.6 + i * 1.65;
    box(s, 5.1, y, 7.65, 1.45, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 5.3, y + 0.4, 0.55, RED);
    txt(s, a, { x: 6.1, y: y + 0.08, w: 6.5, h: 0.55, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY });
    txt(s, b, { x: 6.1, y: y + 0.65, w: 6.5, h: 0.75, fontSize: 27 });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Stifte liegen. Ihr hört nur zu.", options: {} }]);
}

// ============ 13 Ü3 Partnerarbeit ============
{
  const s = blank("V · Verarbeiten · Übung 3", "Ihr analysiert: „Für die Mühsal geboren“", 3, "Ansage: „Arbeitet zu zweit. Flüsterstufe. Die Aufgabe steht rechts.“ Partnerarbeit 7 Minuten, die Schnipsel mit dem Dreischritt liegen auf dem Tisch, die Hilfekarten (Stufe 1 bis 3) liegen am Pult und werden nur auf Nachfrage geholt. Lehrkraft geht durch die Reihen und liest bei 4 bis 6 Paaren mit (unbenotet). Bild: Nachkolorierter Kupferstich, Paris, 1789 (Lehrbuch S. 111, Q3). Bildtext: „Jeden Tag mitten auf dem Feld – in Hitze und Kälte – sieht man den armen Bauern – arbeitend das ganze Jahr über – um mit seiner Arbeit das Geld zu verdienen – und damit den Steuereintreiber zu bezahlen.“ Beischrift oben: „Ne pour la peine“ (Für die Mühsal geboren), unten: „L'Homme de Village“ (Der Mann des Dorfes). Schritt 3 hat einen Stern (*): Bei Zeitnot nur angeleitet gemeinsam. Beurteilen heißt hier: Stimmt das Bild mit dem überein, was wir über Steuern und Stände wissen, wo übertreibt es, was will der Zeichner erreichen?");
  s.addImage({ path: IMG + "karikatur_muehsal.png", x: 0.7, y: 1.4, w: 2.78, h: 4.65 });
  box(s, 3.9, 1.5, 8.85, 1.3, LIGHT);
  txt(s, "„… sieht man den armen Bauern … um mit seiner Arbeit das Geld zu verdienen und damit den Steuereintreiber zu bezahlen.“", { x: 4.1, y: 1.5, w: 8.45, h: 1.3, fontFace: HEAD, fontSize: 27, italic: true, valign: "middle" });
  [["1", "Beschreibe die Karikatur in drei Sätzen.", 3.0, 0.55], ["2", "Deute: Wofür steht der Bauer? Was sagt der Bildtext?", 3.6, 0.95], ["3", "Beurteile: Stimmt das Bild mit unserem Wissen über Steuern? Wo übertreibt es? *", 4.6, 1.2]].forEach(([n, x, y, h]) => {
    numCircle(s, n, 4.0, y + 0.07, 0.45, RED);
    txt(s, x, { x: 4.7, y, w: 8.0, h, fontSize: 27, bold: true, valign: "middle" });
  });
  txt(s, "Hilfe: Schnipsel auf dem Tisch · Hilfekarten am Pult", { x: 3.9, y: 5.75, w: 8.85, h: 0.45, fontSize: 27, color: MUTED });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "7 Minuten · zu zweit · Flüsterstufe      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 14 Ergebnis Karikatur 2 ============
{
  const s = blank("V · Verarbeiten · Sicherung", "Mögliche Ergebnisse", 3, "Sicherung (3 Minuten): gezielte Abfrage bei 3 Paaren (je ein Schritt), kein Gespräch. Erwartung (Auswahl): 1 Beschreiben: großer Bauer mit Hut, Stock, Sack und Werkzeug; im Hintergrund Dorf mit Kirche, Tiere (Schwein, Huhn), ein weiterer Arbeiter; Titel und Bildtext. 2 Deuten: Der Bauer steht für alle Bauern und den Dritten Stand (Bauern und Bürger); er arbeitet das ganze Jahr, um Steuern zu zahlen. Einzelne Bildelemente (z. B. der Hahn auf dem Stock) sind nicht sicher gedeutet: ehrlich offen lassen. 3 Beurteilen: Der Zeichner zeigt den Bauern als Opfer und will Mitleid und Veränderung; das passt zu den Beschwerden (Steuern), übertreibt aber, weil nicht alle Bauern gleich arm waren.");
  [["1 Beschreiben", "Ein großer Bauer mit Hut, Stock und Sack. Hinten ein Dorf, Tiere und ein Arbeiter."], ["2 Deuten", "Der Bauer steht für den Dritten Stand (Bauern und Bürger). Er arbeitet das ganze Jahr, um Steuern zu zahlen."], ["3 Beurteilen*", "Der Zeichner will Mitleid und Veränderung. Das stimmt mit den Beschwerden überein, ist aber auch übertrieben."]].forEach(([a, b], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 1.6, 3.95, 4.4, i % 2 ? LIGHT : LIGHT2);
    txt(s, a, { x: xx + 0.2, y: 1.7, w: 3.55, h: 0.6, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY });
    txt(s, b, { x: xx + 0.2, y: 2.4, w: 3.55, h: 3.4, fontSize: 27 });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche mit deiner Analyse und verbessere mit Grün.", options: {} }]);
}

// ============ 15 Abschlussfrage Berater ============
{
  const s = blank("A · Auswerten", "Abschlussfrage: Du bist Berater", 4, "Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Zehn Sekunden denken, dann schreiben. Es gibt mehrere gute Antworten, geprüft wird die Begründung: eine Position, ein Sachargument aus der Stunde, ein „weil“. Erwartung (Auswahl, keine Musterlösung): SCHWÖREN: Die Abgeordneten vertreten fast alle Menschen (Sieyès: ALLES); wenn alle schwören, kann der König nicht alle verhaften; wer jetzt nachgibt, dem bleibt alles beim Alten; ohne Verfassung entscheidet der König weiter allein; Hunger und Steuern der Menschen. NICHT SCHWÖREN: Die Familie ist in Gefahr; Gewalt vermeiden; allein ist man leicht zu verhaften; später verhandeln. Zwischenlösung (schwören, aber Familie in Sicherheit bringen) ist begründbar. Historisch schworen fast alle Abgeordneten. Das Risiko macht die Frage persönlich: Es steht etwas auf dem Spiel. Form nach Verlauf: Blitzlicht (je eine Stimme „schwören“, eine „nicht schwören“, gezielt aufgerufen), schriftlich mit Vorlesen oder Berater-Duell zu zweit. Stichprobe: bei 4 bis 6 Heften mitlesen, nicht kommentieren (unbenotet). Die Figuren sind Zeichnungen, keine historischen Porträts.");
  s.addImage({ path: IMG + "abgeordneter.png", x: 0.55, y: 1.45, w: 2.9, h: 4.19 });
  txt(s, "Abgeordneter", { x: 0.55, y: 5.6, w: 2.9, h: 0.4, fontFace: HEAD, fontSize: 27, bold: true, align: "center" });
  bubble(s, 4.0, 1.5, 8.75, 1.95, WHITE, NAVY, "left", 2.5);
  txt(s, "„Draußen stehen Soldaten. Wenn ich schwöre, kann ich verhaftet werden, auch meine Familie ist in Gefahr. Soll ich trotzdem schwören?“", { x: 4.3, y: 1.5, w: 8.2, h: 1.95, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle" });
  bubble(s, 4.0, 3.75, 7.1, 1.6, "FDF1D8", GOLD, "right", 4.55);
  txt(s, "Dein Rat:  „Ich rate dir, … , weil …“", { x: 4.3, y: 3.75, w: 6.6, h: 1.6, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle" });
  s.addImage({ path: IMG + "berater.png", x: 11.75, y: 3.4, w: 1.3, h: 1.8 });
  txt(s, "Du", { x: 11.75, y: 5.2, w: 1.3, h: 0.4, fontFace: HEAD, fontSize: 27, bold: true, align: "center" });
  [["Position", 2.1], ["Sachargument", 3.2], ["„weil“", 1.9]].reduce((xx, [lab, ww], i) => {
    box(s, xx, 5.55, ww, 0.62, i % 2 ? LIGHT : LIGHT2);
    s.addText([{ text: lab, options: { bold: true, fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: xx + 0.1, y: 5.55, w: ww - 0.2, h: 0.62, valign: "middle", align: "center", margin: 0 });
    return xx + ww + 0.2;
  }, 4.0);
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "3 Minuten · allein · leise      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 16 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Die Klasse prüft still: Kann ich das? Danach Folie 17. Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen. Wenn das Beurteilen der Karikatur (*) nicht mehr geschafft wurde, kommt es in Stunde 3 als Wiederholung.");
  [["Ich kann erklären, warum sich der Dritte Stand (Bauern und Bürger) 1789 zur Nationalversammlung erklärte."], ["Ich kann erklären, was die Abgeordneten im Ballhausschwur versprachen."], ["Ich kann eine Karikatur analysieren: beschreiben, deuten, beurteilen.*"]].forEach(([x], i) => {
    const y = 1.7 + i * 1.4;
    box(s, 0.6, y, 12.15, 1.15, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.15, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Haben wir sie erreicht?", options: {} }]);
}

// ============ 17 Ausblick ============
{
  const s = blank("A · Auswerten · Ausblick", "Wie geht es weiter?", 4, "Folie nur lesen (1 Minute). Ausblick: Anfang Juli zieht der König Soldaten um Paris zusammen und entlässt seinen Finanzminister. In Paris wächst die Angst. Das sehen wir in der nächsten Stunde (Bastille, Ende der Vorrechte, Menschenrechte).");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Der König zieht Soldaten zusammen.\nIn Paris wächst die Angst.", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Wie reagieren die Menschen in Paris? Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Nächste Stunde: ", options: { bold: true } }, { text: "Bastille, Ende der Vorrechte, Menschenrechte.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Stunde2_Nationalversammlung.pptx" }).then(() => console.log("ok"));
