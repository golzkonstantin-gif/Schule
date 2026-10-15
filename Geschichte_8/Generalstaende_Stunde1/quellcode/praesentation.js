const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Generalstände, Stunde 1";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Generalstände · Geschichte 8";
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
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schreibe den Merksatz mit Beispiel ab und rahme ihn rot ein.      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
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
function stackBar(s, x, y, w, h, parts, labelColors) {
  const tot = parts.reduce((a, b) => a + b[0], 0);
  let xx = x;
  parts.forEach(([v, col, label]) => {
    const ww = (w * v) / tot;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y, w: ww, h, fill: { color: col }, line: { color: WHITE, width: 1.5 } });
    if (label && ww > 1.0) txt(s, label, { x: xx, y, w: ww, h, fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle" });
    xx += ww;
  });
}

// ============ 1 Titel ============
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "Generalstände 1789", { x: 0.9, y: 2.4, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  txt(s, "Wie fair ist die Abstimmung?", { x: 0.9, y: 3.5, w: 11, h: 0.6, fontSize: 26, italic: true, color: ICE });
  txt(s, "Geschichte, Klasse 8 · Französische Revolution", { x: 0.9, y: 6.6, w: 8, h: 0.3, fontSize: 13, color: ICE });
  s.addNotes("Stundenaufbau nach AVIVA: Ankommen (Test zurück) · Vorwissen aktivieren · Informieren (Text und Zeitstrahl) · Verarbeiten (Abstimmung 1789, Schulbeispiel) · Auswerten. Zwei Lehrkräfte, 12 Schülerinnen und Schüler, 45 Minuten.");
}

// ============ 2 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Lernziele vorlesen lassen. Am Ende der Stunde (Folie 16) prüft die Klasse, ob sie erreicht sind.");
  [["1", "Du kannst den Weg bis zu den Generalständen am Zeitstrahl erklären."], ["2", "Du kannst die Abstimmung nach Ständen und nach Köpfen vergleichen."], ["3", "Du kannst begründen, warum die Abstimmung nach Ständen unfair war."]].forEach(([n, x], i) => {
    const y = 1.75 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.37, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 26, valign: "middle" });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Zeitstrahl, Abstimmung nach Ständen und nach Köpfen.", options: {} }]);
}

// ============ 3 Ablauf nach AVIVA ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Kurz den Ablauf zeigen. Die kleinen Kreise oben rechts auf den Folien zeigen die aktuelle Phase: A V I V A.");
  const ph = [["A", "Ankommen", "Test zurück", "8 Min."], ["V", "Vorwissen", "Probleme Frankreichs", "3 Min."], ["I", "Informieren", "Text und Zeitstrahl", "12 Min."], ["V", "Verarbeiten", "Abstimmen", "14 Min."], ["A", "Auswerten", "Ich kann …", "5 Min."]];
  ph.forEach(([l, n, x, m], i) => {
    const xx = 0.6 + i * 2.45;
    box(s, xx, 1.9, 2.3, 3.6, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fill: { color: RED }, line: { type: "none" } });
    txt(s, l, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, n, { x: xx, y: 3.1, w: 2.3, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true, align: "center" });
    txt(s, x, { x: xx + 0.1, y: 3.7, w: 2.1, h: 1.0, fontSize: 24, align: "center", color: MUTED });
    txt(s, m, { x: xx, y: 4.8, w: 2.3, h: 0.5, fontSize: 24, bold: true, color: RED, align: "center" });
  });
  bar(s, [{ text: "Zusammen: ", options: { bold: true } }, { text: "42 Minuten und 3 Minuten Puffer.", options: {} }]);
}

// ============ 4 Ankommen ============
{
  const s = blank("A · Ankommen", "Euer Test", 0, "Ansage: „Stifte liegen. Ihr bekommt jetzt euren Test zurück. Schaut ihn noch nicht an.“ Tests verdeckt austeilen, keine Noten laut. Verteilung und Erfolgsquote je Aufgabe vorher an der Tafel anschreiben (Strichliste ohne Namen). Rückmeldung (Vorschlag): Zwei Hälften, eine hat gelernt, eine nicht, das ist keine Frage von Begabung, sondern von Vorbereitung. Danach die beiden Fragen. Berichtigung mit Grün als Hausaufgabe.");
  [["1", "Was haben die gemacht, die gut abgeschnitten haben?"], ["2", "Was machst du beim nächsten Mal anders?"]].forEach(([n, q], i) => {
    const y = 1.75 + i * 1.7;
    box(s, 0.6, y, 12.15, 1.45, i ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.47, 0.5, RED);
    txt(s, q, { x: 1.7, y, w: 10.8, h: 1.45, fontFace: HEAD, fontSize: 28, valign: "middle" });
  });
  bar(s, [{ text: "Hausaufgabe: ", options: { bold: true } }, { text: "Falsche Aufgaben mit Grün verbessern.", options: {} }]);
}

// ============ 5 Vorwissen ============
{
  const s = blank("V · Vorwissen aktivieren", "Probleme Frankreichs", 1, "Ansage: „Meldet euch. Ihr schreibt noch nichts auf. Welche Probleme hatte Frankreich?“ Die Klasse nennt die Probleme aus der Münzgeschichte, danach die Kacheln zeigen (oder zuerst die Klasse sammeln lassen und die Folie als Abgleich nutzen). Leitfrage unten bleibt stehen.");
  const tiles = ["Leere Staatskasse", "Ungleiche Steuerlast", "Hunger und hoher Brotpreis", "Keine Mitsprache des Dritten Standes (Bauern und Bürger)", "Entfremdung zwischen Adel und Volk"];
  tiles.forEach((x, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const xx = 0.6 + col * 4.1, y = 1.7 + row * 2.1;
    box(s, xx, y, 3.95, 1.9, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y, w: 0.12, h: 1.9, fill: { color: RED }, line: { type: "none" } });
    txt(s, x, { x: xx + 0.3, y, w: 3.5, h: 1.9, fontFace: HEAD, fontSize: 24, valign: "middle" });
  });
  box(s, 8.8, 3.8, 3.95, 1.9, NAVY);
  txt(s, "Wie fair ist die Abstimmung?", { x: 9.0, y: 3.8, w: 3.6, h: 1.9, fontFace: HEAD, fontSize: 24, bold: true, color: WHITE, valign: "middle" });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Meldet euch. Ihr schreibt noch nichts auf.", options: {} }]);
}

// ============ 6 Informieren: Aufgabe ============
{
  const { s } = taskSlide("I · Informieren · Übung 1", "Text und Zeitstrahl", ["Lies den Text. Unterstreiche die Jahreszahlen.", "Trage die Daten in den Zeitstrahl ein.", "Prüfe: Passt jedes Datum zum Ereignis?"], "5 Minuten · allein · leise",
    "Ansage: „Ich teile das Blatt aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Oben der Text, unten der Zeitstrahl mit leeren Datenfeldern. Danach Kontrolle (Folie 7): „Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“ Dann „Vergleicht und verbessert mit Grün.“");
  phaseTag(s, 2);
}

// ============ 7 Zeitstrahl ============
{
  const s = blank("I · Informieren · Kontrolle", "Der Weg zu den Generalständen", 2, "Erst zeigen, wenn die Klasse die Daten genannt hat („Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“). Dann „Vergleicht und verbessert mit Grün.“ Leitfrage: Warum ruft der König die Generalstände ein, obwohl er sie 175 Jahre nicht brauchte? Antwort: Geldnot, Reformen scheitern, Krise (Hunger, drohender Bankrott). Der Zeitstrahl ist nicht maßstabsgetreu.");
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

// ============ 8 Zahlen ============
{
  const s = blank("V · Verarbeiten", "Wer hat wie viel Gewicht?", 3, "Ansage: „Meldet euch. Ihr schreibt noch nichts auf. Was fällt euch an diesen Balken auf?“ Erwartung: Der Dritte Stand (Bauern und Bürger) ist fast die ganze Bevölkerung (ca. 97 %) und stellt etwa die Hälfte der Abgeordneten (ca. 580 von ca. 1140), hat aber bei der Abstimmung nach Ständen nur eine von drei Stimmen. Klerus und Adel zusammen (ca. 3 %) haben zwei von drei Stimmen. Zahlen der Abgeordneten bitte gegen das Buch prüfen. Menschen: Klerus 150 000, Adel 500 000, Dritter Stand 20 000 000.");
  legend(s, 0.6, 1.5);
  [["Bevölkerung", [[0.73, KLERUS, ""], [2.42, ADEL, ""], [96.85, DRITTER, "ca. 97 %"]], "Klerus ca. 1 % · Adel ca. 2 %"], ["Abgeordnete", [[290, KLERUS, "ca. 290"], [270, ADEL, "ca. 270"], [580, DRITTER, "ca. 580"]], ""], ["Stimmen nach Ständen", [[1, KLERUS, "1"], [1, ADEL, "1"], [1, DRITTER, "1"]], ""]].forEach(([l, parts, note], i) => {
    const y = 2.25 + i * 1.3;
    txt(s, l, { x: 0.6, y, w: 8, h: 0.45, fontFace: HEAD, fontSize: 24, bold: true });
    if (note) txt(s, note, { x: 5.0, y, w: 7.7, h: 0.45, fontSize: 24, color: MUTED, align: "right" });
    stackBar(s, 0.6, y + 0.5, 12.15, 0.65, parts);
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Meldet euch. Was fällt euch auf?", options: {} }]);
}

// ============ 9 Antrag 1789 ============
{
  const s = blank("V · Verarbeiten · Historisches Beispiel", "Antrag in den Generalständen", 3, "Modellannahme: Klerus und Adel stimmen geschlossen Nein, der Dritte Stand Ja. Die Klasse rechnet: nach Ständen und nach Köpfen. Ergebnis erst auf Folie 10 zeigen. Hinweis: Die Annahme ist vereinfacht, tatsächlich unterstützten manche Geistliche und einzelne Adlige den Dritten Stand.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.65, w: 12.15, h: 1.9, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  txt(s, "Auch Adel und Klerus zahlen künftig Steuern.", { x: 0.95, y: 1.7, w: 11.45, h: 1.8, fontFace: HEAD, fontSize: 34, bold: true, valign: "middle" });
  [["Klerus", "Nein", KLERUS, NEIN], ["Adel", "Nein", ADEL, NEIN], ["Dritter Stand", "Ja", DRITTER, JA]].forEach(([n, v, c, vc], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 3.9, 3.95, 1.5, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 3.9, w: 0.14, h: 1.5, fill: { color: c }, line: { type: "none" } });
    txt(s, n, { x: xx + 0.3, y: 3.9, w: 2.2, h: 1.5, fontFace: HEAD, fontSize: 24, bold: true, valign: "middle" });
    txt(s, v, { x: xx + 2.3, y: 3.9, w: 1.5, h: 1.5, fontSize: 30, bold: true, color: vc, align: "center", valign: "middle" });
  });
  bar(s, [{ text: "Frage: ", options: { bold: true } }, { text: "Wer gewinnt nach Ständen? Wer gewinnt nach Köpfen?", options: {} }]);
}

// ============ 10 Ergebnis 1789 ============
{
  const s = blank("V · Verarbeiten · Ergebnis", "Zwei Regeln, zwei Ergebnisse", 3, "Nach Ständen: 2:1, abgelehnt. Nach Köpfen: Ja 580, Nein 290 + 270 = 560, knapp angenommen. Kernaussage: Dieselben Stimmen führen je nach Regel zu zwei verschiedenen Ergebnissen. Frage: Wie wichtig ist die Abstimmungsregel für den Dritten Stand (Bauern und Bürger)?");
  txt(s, "Nach Ständen", { x: 0.6, y: 1.6, w: 5.8, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY });
  [["Klerus", "Nein", KLERUS, NEIN], ["Adel", "Nein", ADEL, NEIN], ["Dritter Stand", "Ja", DRITTER, JA]].forEach(([n, v, c, vc], i) => {
    const y = 2.2 + i * 0.85;
    box(s, 0.6, y, 5.8, 0.75, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.14, h: 0.75, fill: { color: c }, line: { type: "none" } });
    txt(s, n, { x: 0.9, y, w: 3.3, h: 0.75, fontSize: 24, valign: "middle" });
    txt(s, v, { x: 4.4, y, w: 1.8, h: 0.75, fontSize: 26, bold: true, color: vc, align: "center", valign: "middle" });
  });
  box(s, 0.6, 4.85, 5.8, 1.05, NAVY);
  txt(s, "2 : 1  →  abgelehnt", { x: 0.6, y: 4.85, w: 5.8, h: 1.05, fontFace: HEAD, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Nach Köpfen", { x: 6.95, y: 1.6, w: 5.8, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY });
  stackBar(s, 6.95, 2.4, 5.8, 0.9, [[580, JA, "Ja 580"], [560, NEIN, "Nein 560"]]);
  txt(s, "Nein: Klerus 290 + Adel 270", { x: 6.95, y: 3.4, w: 5.8, h: 0.5, fontSize: 24, color: MUTED });
  box(s, 6.95, 4.85, 5.8, 1.05, NAVY);
  txt(s, "580 : 560  →  angenommen", { x: 6.95, y: 4.85, w: 5.8, h: 1.05, fontFace: HEAD, fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle" });
  bar(s, [{ text: "Merke: ", options: { bold: true } }, { text: "Die Regel entscheidet, nicht die Mehrheit der Menschen.", options: {} }]);
}

// ============ 11 Antrag Schule ============
{
  const s = blank("V · Verarbeiten · Schulbeispiel", "Dasselbe in der Schule", 3, "Ansage: „Jetzt spielen wir dasselbe in der Schule. Das ist nur ein Spiel, es gilt nicht.“ Rollen: Klerus = Lehrkraft 1, Adel = Lehrkraft 2, Dritter Stand (Bauern und Bürger) = die Schüler. Nach Ständen: Lehrkraft 1 Ja, Lehrkraft 2 Ja, Schüler Nein: 2:1, angenommen. Nach Köpfen: 2 Ja, 12 Nein: abgelehnt. Fragen: Wie fühlt sich das an? Was ist unfair? Sind die Lehrkräfte das Problem oder die Regeln? Sollbruchstelle: Bei Zeitnot entfällt das Schulbeispiel und beginnt Stunde 2.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.65, w: 12.15, h: 2.7, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  txt(s, "Nur Schüler putzen künftig das Klassenzimmer und räumen die Tische auf. Lehrkräfte sind davon befreit.", { x: 0.95, y: 1.7, w: 11.45, h: 2.6, fontFace: HEAD, fontSize: 32, bold: true, valign: "middle" });
  [["Klerus", "Lehrkraft 1", KLERUS], ["Adel", "Lehrkraft 2", ADEL], ["Dritter Stand (Bauern und Bürger)", "die Schüler", DRITTER]].forEach(([n, who, c], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 4.6, 3.95, 1.4, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 4.6, w: 0.14, h: 1.4, fill: { color: c }, line: { type: "none" } });
    txt(s, n, { x: xx + 0.3, y: 4.65, w: 3.5, h: 0.7, fontFace: HEAD, fontSize: 24, bold: true, valign: "middle" });
    txt(s, who, { x: xx + 0.3, y: 5.3, w: 3.5, h: 0.6, fontSize: 24, color: MUTED });
  });
  bar(s, [{ text: "Hinweis: ", options: { bold: true } }, { text: "Das ist nur ein Spiel. Der Antrag gilt nicht.", options: {} }]);
}

// ============ 12 Ergebnis Schule ============
{
  const s = blank("V · Verarbeiten · Ergebnis", "Schulspiel: Zwei Regeln, zwei Ergebnisse", 3, "Nach Ständen: Lehrkraft 1 Ja, Lehrkraft 2 Ja, Schüler Nein: 2:1, angenommen. Nach Köpfen: 2 Ja, 12 Nein: abgelehnt. Unterschied zu 1789: In der Schule ist das Ergebnis deutlich (12 gegen 2), 1789 knapp (580 gegen 560), weil der Dritte Stand schon mehr Abgeordnete bekommen hatte. Die Regeln sind das Problem, nicht die Personen.");
  txt(s, "Nach Ständen", { x: 0.6, y: 1.6, w: 5.8, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY });
  [["Lehrkraft 1 (Klerus)", "Ja", KLERUS, JA], ["Lehrkraft 2 (Adel)", "Ja", ADEL, JA], ["Schüler (Dritter Stand)", "Nein", DRITTER, NEIN]].forEach(([n, v, c, vc], i) => {
    const y = 2.2 + i * 0.85;
    box(s, 0.6, y, 5.8, 0.75, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.14, h: 0.75, fill: { color: c }, line: { type: "none" } });
    txt(s, n, { x: 0.9, y, w: 4.1, h: 0.75, fontSize: 24, valign: "middle" });
    txt(s, v, { x: 5.0, y, w: 1.3, h: 0.75, fontSize: 26, bold: true, color: vc, align: "center", valign: "middle" });
  });
  box(s, 0.6, 4.85, 5.8, 1.05, NAVY);
  txt(s, "2 : 1  →  angenommen", { x: 0.6, y: 4.85, w: 5.8, h: 1.05, fontFace: HEAD, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Nach Köpfen", { x: 6.95, y: 1.6, w: 5.8, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY });
  stackBar(s, 6.95, 2.4, 5.8, 0.9, [[2, JA, "Ja"], [12, NEIN, "Nein 12"]]);
  txt(s, "2 Ja · 12 Nein", { x: 6.95, y: 3.4, w: 5.8, h: 0.5, fontSize: 24, color: MUTED });
  box(s, 6.95, 4.85, 5.8, 1.05, NAVY);
  txt(s, "2 : 12  →  abgelehnt", { x: 6.95, y: 4.85, w: 5.8, h: 1.05, fontFace: HEAD, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
  bar(s, [{ text: "Frage: ", options: { bold: true } }, { text: "Was ist unfair? Die Lehrkräfte oder die Regeln?", options: {} }]);
}

// ============ 13 Gewicht einer Stimme ============
{
  const s = blank("V · Verarbeiten · Rechnung", "Wie viel zählt eine Stimme?", 3, "Rechnung an der Tafel: 20 000 000 : 150 000 ≈ 133 (Klerus), 20 000 000 : 500 000 = 40 (Adel). Im Klassenspiel wiegt eine Lehrerstimme zwölfmal so viel (1 Person gegen 12 Schüler). Frage: „Was würdet ihr als Dritter Stand in Frankreich fordern?“ Erwartung: Abstimmung nach Köpfen. Das ist genau, was die Abgeordneten 1789 fordern (Brücke zu Stunde 2). Die Balken sind maßstabsgetreu: Klerus und Adel sind kaum zu sehen.");
  const rows = [["Klerus", 150000, KLERUS, "pro Stimme: 150 000"], ["Adel", 500000, ADEL, "pro Stimme: 500 000"], ["Dritter Stand (Bauern und Bürger)", 20000000, DRITTER, "pro Stimme: 20 000 000"]];
  rows.forEach(([n, v, c, l], i) => {
    const y = 1.65 + i * 1.45;
    txt(s, n, { x: 0.6, y, w: 6.6, h: 0.45, fontFace: HEAD, fontSize: 24, bold: true });
    txt(s, l, { x: 7.25, y, w: 5.5, h: 0.45, fontSize: 24, color: MUTED, align: "right" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.55, w: 12.15, h: 0.6, fill: { color: LIGHT }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.55, w: Math.max(0.06, (12.15 * v) / 20000000), h: 0.6, fill: { color: c }, line: { type: "none" } });
  });
  bar(s, [{ text: "Gewicht: ", options: { bold: true } }, { text: "Eine Stimme des Klerus wiegt rund 133-mal so viel.", options: {} }]);
}

// ============ 14 Merksatz ============
{
  const s = merkSlide(1, "Abstimmung nach Ständen",
    "Bei den Generalständen hat jeder Stand **eine Stimme**. Klerus und Adel (zusammen ca. 3 % der Bevölkerung) haben damit **zwei von drei Stimmen** und können den Dritten Stand (Bauern und Bürger, ca. 97 %) **immer überstimmen**.",
    "Eine Stimme des Klerus wiegt rund **133-mal** so viel wie eine des Dritten Standes.",
    "Wie viele Menschen stehen hinter **einer Stimme**?",
    "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“", 24);
  phaseTag(s, 3);
}

// ============ 15 Auswerten ============
{
  const { s } = taskSlide("A · Auswerten · Übung 2", "Ist die Abstimmung fair?", ["Schreibe Ü2 ins Heft.", "Beende den Satz: „Die Abstimmung nach Ständen ist unfair, weil …“", "Begründe mit mindestens einer Zahl."], "4 Minuten · allein · leise",
    "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Beispiel: „… weil der Dritte Stand (Bauern und Bürger) ca. 97 % der Menschen stellt, aber nur eine von drei Stimmen hat und immer überstimmt wird.“ 1 bis 2 Sätze vorlesen lassen.");
  phaseTag(s, 4);
}

// ============ 16 Lernziele prüfen, Ausblick ============
{
  const s = blank("A · Auswerten", "Hast du die Lernziele erreicht?", 4, "Die Klasse prüft still: Kann ich das? Danach Ausblick: Was passiert, wenn der Dritte Stand „Nach Köpfen!“ fordert und der König ablehnt? Das sehen wir in der nächsten Stunde nach den Ferien. Die Berichtigung des Tests ist Hausaufgabe.");
  [["Ich kann den Weg bis zu den Generalständen am Zeitstrahl erklären."], ["Ich kann die Abstimmung nach Ständen und nach Köpfen vergleichen."], ["Ich kann begründen, warum die Abstimmung nach Ständen unfair war."]].forEach(([x], i) => {
    const y = 1.6 + i * 1.18;
    box(s, 0.6, y, 12.15, 1.05, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.85, y: y + 0.25, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.05, fontFace: HEAD, fontSize: 24, valign: "middle" });
  });
  box(s, 0.6, 5.2, 12.15, 0.9, NAVY);
  txt(s, "Nächste Stunde: Was passiert, wenn der Dritte Stand „Nach Köpfen!“ fordert?", { x: 0.9, y: 5.2, w: 11.6, h: 0.9, fontFace: HEAD, fontSize: 24, bold: true, color: WHITE, valign: "middle" });
  bar(s, [{ text: "Hausaufgabe: ", options: { bold: true } }, { text: "Testaufgaben mit Grün berichtigen.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Generalstaende_Stunde1.pptx" }).then(() => console.log("ok"));
