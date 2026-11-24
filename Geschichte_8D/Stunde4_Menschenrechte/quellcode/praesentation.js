const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Stunde 4: Menschenrechte (8D)";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Menschenrechte · Geschichte 8D";
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

const SAGT = "„Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?“";
const K = 0.0062, FW = 1280 * K, FH = 970 * K;
const crop = (s, which, x, y) => s.addImage({ path: IMG + "karikatur_rueckschau.png", x, y, w: FW, h: FH, sizing: { type: "crop", x: (which === "L" ? 30 : 670) * K, y: 22 * K, w: 580 * K, h: 336 * K } });
const PH = 336 * K; // Höhe eines Gemäldeausschnitts

// ============ 1 Karikatur / Ankommen ============
{
  const s = blank("Ankommen", "Ein Gespräch in Versailles", 0, "Dient Ziel 1 (Vermutung, Vorher-Nachher mit der Abschlussaufgabe). Die Folie steht beim Betreten. Niemand muss etwas sagen. Die Karikatur ist selbst gezeichnet, keine historische Quelle. An der Wand hängen zwei Bilder: der Eid im Ballhaus (links) und der Sturm auf die Bastille (rechts), beide unbeschriftet. Davor stehen Klerus und Adel und sagen: „Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?“ Der Dritte Stand (Bauern und Bürger) hört zu. Die Aussage ist bewusst überspitzt (erfundene Selbstdarstellung): Entgegenkommen heißt Gnade von oben. Gegenpol, der am Ende der Stunde auftaucht: Rechte, die für alle gelten. Die Klasse schreibt ihre Vermutung ins Heft (bleibt stehen, wird am Ende mit der Antwort verglichen): „Was will der Dritte Stand noch? Ich vermute …“ Keine Besprechung. Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2 bis 4 in zusammen etwa 2 Minuten.");
  s.addImage({ path: IMG + "karikatur_rueckschau.png", x: 0.6, y: 1.4, w: 6.8, h: (6.8 * 970) / 1280 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 3.35, w: 6.0, h: 1.5, rectRadius: 0.18, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: 3.7, y: 4.8, w: 0.6, h: 0.45, rotate: 180, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.72, y: 4.77, w: 0.56, h: 0.1, fill: { color: WHITE }, line: { type: "none" } });
  txt(s, SAGT, { x: 1.2, y: 3.35, w: 5.6, h: 1.5, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle" });
  [["Klerus", 0.95, 1.6], ["Adel", 3.3, 1.4], ["Dritter Stand", 5.0, 2.4]].forEach(([n, x, w]) => {
    box(s, x, 6.0, w, 0.5, WHITE, { line: NAVY, lw: 2 });
    txt(s, n, { x, y: 6.0, w, h: 0.5, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  [["Heft raus. Datum und Überschrift.", 1.6, 1.0], ["Sieh dir die Bilder an. Lies, was Klerus und Adel sagen.", 2.75, 1.5], ["Schreibe: Was will der Dritte Stand noch? Ich vermute …", 4.45, 1.6]].forEach(([x, y, h], i) => {
    numCircle(s, i + 1, 7.95, y + 0.05, 0.45, RED);
    txt(s, x, { x: 8.6, y, w: 4.15, h, fontSize: 27, bold: true, valign: "top" });
  });
}

// ============ 2 Stundenfrage ============
{
  const s = blank("Stundenfrage", "Das untersuchen wir heute", undefined, "Dient Ziel 1 bis 3 (Orientierung). Stundenfrage und Begriff (45 Sekunden, Klasse hört nur zu): „Klerus und Adel sagen: Wir kamen euch entgegen. Was wollt ihr noch? Wir prüfen: Was ändern der Verzicht auf die Vorrechte und die Erklärung der Menschenrechte von 1789 wirklich? Und: Was davon gilt noch heute? Die Erklärung ist eine Quelle, ein Text aus der Zeit selbst. Wir lesen sie nicht nur, wir übersetzen sie in heutige Begriffe. Das braucht ihr bei jeder Quelle und im Test.“ Relevanz benennen, nicht diskutieren. Die Frage lässt offen, was sich geändert hat. Keine Antworten ins Plenum, die Vermutungen stehen im Heft.");
  box(s, 0.6, 1.5, 12.15, 2.2, NAVY);
  txt(s, "Was will der Dritte Stand (Bauern und Bürger) noch? Was ändern die Vorrechte-Abschaffung und die Menschenrechte?", { x: 1.0, y: 1.5, w: 11.35, h: 2.2, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.95, 12.15, 2.0, LIGHT2);
  s.addText([{ text: "Quelle: ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 27 } }, { text: "Ein Text aus der Zeit selbst. Wir übersetzen ihn in heutige Begriffe. Das brauchst du bei jeder Quelle und im Test.", options: { fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: 1.0, y: 3.95, w: 11.35, h: 2.0, valign: "middle", margin: 0 });
}

// ============ 3 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Dient Ziel 1 bis 3 (Transparenz). Lernziele vorlesen (45 Sekunden). Sie sind zugleich das Raster für die Rückmeldung: Stufe 1 = Ziel 1, Stufe 2 = Ziel 2, Stufe 3 = Ziel 3. Am Ende beantwortet jede Person die Ausgangsfrage (Karikatur) auf einem Exit-Ticket; Sie lesen die Zettel und geben Rückmeldung nach den Stufen. Punkte oder Noten gibt es heute nicht. Keine Selbsteinschätzung der Schülerinnen und Schüler. Mindestziel: Ziel 1 und 2. Ziel 3 hat einen Stern: Der Transfer wird gekürzt, wenn die Zeit knapp ist. Die Phase „Auswerten“ bleibt immer. Am Ende (Folie 22) prüft die Klasse still, ob die Ziele erreicht sind.");
  [["1", "Ich kann erklären, was sich durch den Verzicht auf die Vorrechte und die Menschenrechte für den Dritten Stand verbessert hat."], ["2", "Ich kann Artikel von 1789 in heutige Begriffe übersetzen (zum Beispiel Eigentumsrecht)."], ["3", "Ich kann an einem Grundgesetz-Artikel zeigen, dass die Erklärung bis heute nachwirkt.*"]].forEach(([n, x], i) => {
    const y = 1.55 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.3, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.4, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.3, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Rückmeldung von mir: ", options: { bold: true } }, { text: "Ausgangsfrage als Exit-Ticket am Ende.", options: {} }]);
}

// ============ 4 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Dient Ziel 1 bis 3 (Transparenz). Ablauf zeigen (30 Sekunden). Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Ziel 3). Insgesamt 45 Minuten, die Stunde ist lehrergesteuert mit Einzelarbeit (nach dem Partnerpuzzle der letzten Stunde).");
  const ph = [["A", "Ankommen", "Karikatur", "4 Min."], ["V", "Vorwissen", "Rückschau", "4 Min."], ["I", "Informieren", "Vorrechte, Modell", "10 Min."], ["V", "Verarbeiten", "Blatt, Aha", "18 Min."], ["A", "Auswerten", "Merksatz, Antwort", "9 Min."]];
  ph.forEach(([l, n, x, m], i) => {
    const xx = 0.6 + i * 2.45;
    box(s, xx, 1.9, 2.3, 3.6, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fill: { color: RED }, line: { type: "none" } });
    txt(s, l, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, n, { x: xx, y: 3.1, w: 2.3, h: 0.5, fontFace: HEAD, fontSize: 27, bold: true, align: "center" });
    txt(s, x, { x: xx + 0.05, y: 3.65, w: 2.2, h: 1.15, fontSize: 27, align: "center", color: MUTED });
    txt(s, m, { x: xx, y: 4.85, w: 2.3, h: 0.5, fontSize: 27, bold: true, color: RED, align: "center" });
  });
  bar(s, [{ text: "Wichtig: ", options: { bold: true } }, { text: "Wir arbeiten leise. Gesprochen wird nur nach Ansage.", options: {} }]);
}

// ============ 5 Rückschau: Bilder ============
{
  const s = blank("V · Vorwissen aktivieren · Rückschau", "Rückschau: Was meinen Klerus und Adel?", 1, "Dient Ziel 1 (Ereignisse von 1789 als Voraussetzung für „vorher, nachher“). Rückschau (4 Minuten mit Folie 6, allein und Plenum, kein Partnercheck, weil die Stunde lehrergesteuert ist). Ansage: „Stifte liegen. Ihr hört nur zu. Schaut auf die zwei Bilder. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Die Bilder sind unbeschriftet: Die Klasse entschlüsselt „geschworen“ und „gestürmt“ selbst (links der Eid im Ballhaus, 20. Juni, rechts der Sturm auf die Bastille, 14. Juli). 2 Minuten allein ins Heft: zu jedem Bild „Was geschah? Was bewirkte es?“ Dann 1 Minute Plenum: Stichproben bei 3 bis 4 Personen unterschiedlichen Niveaus (kurze Antwort, unbenotet), dann Folie 6. Wurde in Stunde 3 etwas nicht geschafft, steht es hier als Wiederholung. Gelenk: „Die Bilder haben mit dem Entgegenkommen zu tun. Wie genau, seht ihr gleich.“");
  crop(s, "L", 0.6, 1.5);
  crop(s, "R", 0.6, 1.5 + PH + 0.12);
  [["Schaue auf die zwei Bilder. Was meinen „geschworen“ und „gestürmt“?", 1], ["Schreibe zu jedem Bild: Was geschah? Was bewirkte es?", 2], ["Überlege: Was haben die Bilder mit dem Entgegenkommen zu tun?", 3]].forEach(([x, n], i) => {
    const ys = [1.5, 2.85, 4.1], hs = [1.2, 1.1, 1.75];
    box(s, 4.6, ys[i], 8.15, hs[i], i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 4.8, ys[i] + (hs[i] - 0.45) / 2, 0.45, NAVY);
    txt(s, x, { x: 5.5, y: ys[i], w: 7.1, h: hs[i], fontSize: 27, bold: true, valign: "middle" });
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "2 Min. allein, 2 Min. Plenum      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 6 Kette ============
{
  const s = blank("V · Vorwissen aktivieren · Kontrolle", "So hängt es zusammen", 1, "Dient Ziel 1 (Rahmen). Kontrolle der Rückschau (1 Minute, Sie sprechen). Ansage: „Vergleicht und verbessert mit Grün.“ Zeilen 1 bis 3: 20. Juni Ballhausschwur (die Versammlung bleibt zusammen, gewaltlos), 14. Juli Bastille (der König zieht die Soldaten ab), im Sommer greifen Bauern Schlösser an. Zeilen 4 und 5 nur nennen (werden gleich erklärt): 4. August Vorrechte abgeschafft, 26. August Erklärung der Menschenrechte. Gelenk: „Klerus und Adel kamen entgegen, weil sie Angst hatten. Was genau haben sie aufgegeben? Hört zu.“");
  const ev = [["20. Juni 1789", "Ballhausschwur: Die Versammlung bleibt zusammen.", false], ["14. Juli 1789", "Bastille: Der König zieht die Soldaten ab.", false], ["Sommer 1789", "Bauern greifen Schlösser an.", false], ["4. August 1789", "Vorrechte werden abgeschafft.", true], ["26. August 1789", "Erklärung der Menschenrechte", true]];
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 1.7, w: 0.1, h: 4.4, fill: { color: NAVY }, line: { type: "none" } });
  ev.forEach(([d, e, hl], i) => {
    const y = 1.45 + i * 0.93;
    box(s, 0.6, y, 2.8, 0.78, i % 2 ? LIGHT : LIGHT2);
    txt(s, d, { x: 0.7, y, w: 2.6, h: 0.78, fontSize: 27, bold: true, color: RED, align: "center", valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: y + 0.22, w: 0.34, h: 0.34, fill: { color: hl ? RED : NAVY }, line: { color: WHITE, width: 2 } });
    txt(s, e, { x: 4.2, y, w: 8.55, h: 0.78, fontFace: HEAD, fontSize: 27, bold: hl, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 7 Verzicht auf die Vorrechte ============
{
  const s = blank("I · Informieren · Portion 1", "Der Verzicht auf die Vorrechte", 2, "Dient Ziel 1 (vorher, nachher). Portion 1 (3 Minuten, Sie sprechen, Stifte liegen). Ansage: „Stifte liegen. Ihr hört nur zu.“ In der Nacht zum 4. August 1789 beschließt die Nationalversammlung, die Vorrechte abzuschaffen. Klerus und Adel geben nach, weil sie Angst vor weiteren Aufständen der Bauern haben: Es war nicht freiwillig, das Entgegenkommen war Gnade von oben. Vorher: Klerus und Adel zahlen keine Steuern, die Bauern leisten Abgaben und Dienste an den Adel, Ämter und Jagd sind Vorrechte des Adels. Diese Folie ist (neben Stunde 1 bis 3) die Wissensbasis für den Vorher-Speicher auf dem Blatt. Danach (Beschluss): Die Steuerfreiheit fällt, viele Abgaben fallen weg oder müssen abgelöst werden, die Ämter stehen allen offen (erst am 26. August als Recht festgeschrieben). Warnung vor dem Fehler: Abgeschafft heißt nicht, dass sofort alle gleich leben (viele Abgaben müssen die Bauern noch ablösen, also bezahlen; im Lehrbuch gegenprüfen); und nicht der König, sondern die Nationalversammlung beschloss es. Prüffrage (Stichprobe, 2 Personen): „Warum geben Klerus und Adel nach?“ Erwartung: aus Angst vor den Aufständen. Tafelbild: Vorher → Nachher beginnt hier und wird in Folie 15 und im Merksatz weitergeführt.");
  [["Vorher", "Klerus und Adel zahlen keine Steuern. Bauern leisten Abgaben und Dienste. Ämter sind für Adlige.", "D9534F"], ["Nach dem 4. August", "Die Steuerfreiheit fällt. Viele Abgaben fallen weg oder werden abgelöst. Die Vorrechte gelten nicht mehr.", "2E9E6B"]].forEach(([h, x, col], i) => {
    const xx = 0.6 + i * 6.2;
    box(s, xx, 1.5, 5.95, 4.5, LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 1.5, w: 5.95, h: 0.7, fill: { color: col }, line: { type: "none" } });
    txt(s, h, { x: xx, y: 1.5, w: 5.95, h: 0.7, fontFace: HEAD, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, x, { x: xx + 0.25, y: 2.4, w: 5.45, h: 3.5, fontSize: 27 });
  });
  bar(s, [{ text: "Achtung: ", options: { bold: true } }, { text: "Abgeschafft heißt nicht, dass sofort alles anders ist.", options: {} }]);
}

// ============ 8 Drei Schritte ============
{
  const s = blank("I · Informieren · Portion 2", "Eine Quelle ins Heute übersetzen", 2, "Dient Ziel 2 (Methode). Portion 2, Teil 1 (1 Minute, die Erklärung vom 26. August ist die Quelle: Vorbild die amerikanische Unabhängigkeitserklärung von 1776). Ansage: „Stifte liegen. Ihr hört nur zu.“ Wir übersetzen die Erklärung in heutige Begriffe. Immer in dieser Reihenfolge: 1 Meine Worte: Was steht da? Sag es so, dass es ein Freund versteht. 2 Begriff heute: Wie sagen wir heute dazu? Der Wortspeicher auf dem Blatt erklärt jeden Begriff in einem Satz, die Schüler kreuzen an. 3 Vorher: Was war vorher anders? Der Vorher-Speicher auf dem Blatt nennt fünf Dinge über die Zeit vor 1789 (aus Stunde 1 bis 3 und Folie 7); die Schüler müssen nicht raten, sondern wählen aus. Warnung vor dem Fehler: Nicht den Wortlaut abschreiben, erst in eigenen Worten sagen, dann den Begriff wählen.");
  [["1", "Meine Worte", "Was steht da? Sag es einfach.", "2E9E6B"], ["2", "Begriff heute", "Wie sagen wir heute dazu? Wortspeicher.", "3F7CC4"], ["3", "Vorher", "Was war vorher anders?", "E8A33D"]].forEach(([n, h, x, col], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 1.6, 3.95, 3.9, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 1.6, w: 3.95, h: 0.16, fill: { color: col }, line: { type: "none" } });
    numCircle(s, n, xx + 0.3, 1.95, 0.6, RED);
    txt(s, h, { x: xx + 1.1, y: 1.95, w: 2.7, h: 0.6, fontFace: HEAD, fontSize: 28, bold: true, valign: "middle" });
    txt(s, x, { x: xx + 0.3, y: 2.9, w: 3.4, h: 2.4, fontSize: 27 });
  });
  bar(s, [{ text: "Drei Schritte, ", options: { bold: true } }, { text: "immer in dieser Reihenfolge.", options: {} }]);
}

// ============ 9 Modell Artikel 1 ============
{
  const s = blank("I · Informieren · Portion 2 · Modell", "Ich mache es vor: Artikel 1", 2, "Dient Ziel 2 (Modell des Zielprodukts). Portion 2, Teil 2 (3 Minuten, Sie denken laut). Ansage: „Stifte liegen. Ihr hört nur zu. Ich mache es vor, ihr schaut zu.“ Zeigen Sie zuerst nur den Artikel im roten Rahmen (die drei Zeilen abdecken). Denken Sie laut: „Schritt 1, meine Worte: ‚Frei und gleich an Rechten geboren.‘ Das heißt: Alle Menschen haben von Geburt an dieselben Rechte. Schritt 2: Wie sagen wir heute dazu? Im Wortspeicher steht: Gleichheit, alle haben dieselben Rechte. Das passt, ich kreuze Gleichheit an. Schritt 3: Was war vorher anders? Die Geburt entschied über den Stand und die Vorrechte. Wer gewinnt? Der Dritte Stand.“ Zeigen Sie dann die drei Zeilen. Auf dem Blatt (Seite 1) steht dasselbe gelb hinterlegt. Warnung vor dem Fehler: ‚Gleich an Rechten‘ heißt nicht gleich reich oder gleich stark. Reserve (1 Minute, nur bei Bedarf): Schritt 2 noch einmal erklären.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.5, w: 12.15, h: 1.35, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  s.addText([{ text: "Artikel 1:  ", options: { bold: true, color: RED } }, { text: "„Die Menschen werden frei und gleich an Rechten geboren und bleiben es.“", options: {} }], { isTextBox: true, x: 0.95, y: 1.5, w: 11.45, h: 1.35, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle", margin: 0 });
  [["1  Meine Worte", "Alle Menschen haben von Geburt an dieselben Rechte.", "2E9E6B"], ["2  Begriff heute", "Gleichheit", "3F7CC4"], ["3  Vorher", "Die Geburt entschied über Stand und Vorrechte. Gewinner: Dritter Stand.", "E8A33D"]].forEach(([h, x, col], i) => {
    const y = 3.0 + i * 1.05;
    box(s, 0.6, y, 12.15, 0.92, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.16, h: 0.92, fill: { color: col }, line: { type: "none" } });
    txt(s, h, { x: 0.95, y, w: 3.0, h: 0.92, fontFace: HEAD, fontSize: 27, bold: true, valign: "middle" });
    txt(s, x, { x: 4.0, y, w: 8.6, h: 0.92, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Du schaust zu. Stifte liegen.", options: {} }]);
}

// ============ 10 Prüffrage ============
{
  const s = blank("I · Informieren · Prüffrage", "Prüffrage", 2, "Dient Ziel 2 (Verständnisprüfung vor dem Üben). Prüffrage (1 Minute, unbenotet, getrennt von den mündlichen Noten). Ansage: „Stifte liegen. Ich rufe euch auf. Antwortet in einem Satz.“ Rufen Sie 4 bis 6 Personen unterschiedlichen Niveaus auf, kurze Antworten, kein Gespräch. Frage 1: Was machst du im Schritt 2 „Begriff heute“? Erwartung: Ich lese im Wortspeicher, welcher Satz zu meinen eigenen Worten passt, und kreuze den Begriff an. Frage 2: Warum steht „Gleichheit“ nicht wörtlich im Artikel 1? Erwartung: Die Erklärung ist von 1789, wir sagen heute ein anderes Wort für dieselbe Sache. Wird die Prüffrage nicht richtig beantwortet, wiederholen Sie Schritt 2 am Artikel 1 (hier liegt die Reserve). Es wird nicht weitergegangen, wenn etwas nicht verstanden ist. Gelenk: „Jetzt macht ihr es selbst, zuerst mit Hilfe.“");
  box(s, 0.6, 1.6, 12.15, 1.7, NAVY);
  txt(s, "Was machst du im Schritt 2 „Begriff heute“?", { x: 1.0, y: 1.6, w: 11.35, h: 1.7, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.6, 12.15, 1.7, LIGHT2);
  txt(s, "Warum steht „Gleichheit“ nicht wörtlich im Artikel?", { x: 1.0, y: 3.6, w: 11.35, h: 1.7, fontFace: HEAD, fontSize: 32, bold: true, valign: "middle" });
  bar(s, [{ text: "Ich rufe auf. ", options: { bold: true } }, { text: "Antworte in einem Satz.", options: {} }]);
}

// ============ 11 Ü1 Artikel 3 ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 1", "Artikel 3 mit Hilfe", ["Arbeitsblatt, Seite 1. Lies Artikel 3.", "Schreibe bei „In meinen Worten“. Nutze die Hilfe.", "Der Begriff ist schon angekreuzt.", "Ergänze „Vorher“. Nutze den Vorher-Speicher."], "3 Min. · allein · leise",
    "Dient Ziel 1 und 2 (erste Anwendung mit Hilfe). Angeleitetes Üben (3 Minuten mit Folie 12). Ansage: „Stifte liegen. Ihr hört nur zu. Jetzt nehmt ihr das Arbeitsblatt. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Das Arbeitsblatt liegt vorab auf den Tischen (umgedreht, jede Person ein Blatt, doppelseitig). Durch die Reihen gehen (Heft-Stichprobe, nicht kommentieren). Warnung vor dem Fehler: Hilfe ausfüllen, nicht den Satz aus dem Artikel kopieren. Danach Folie 12.");
  phaseTag(s, 3);
}

// ============ 12 Kontrolle Artikel 3 ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "So sieht Artikel 3 aus", 3, "Dient Ziel 1 und 2 (Kontrolle der ersten Anwendung). Kontrolle (30 Sekunden bis 1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ Stichprobe bei 2 Personen (kurz vorlesen, kein Gespräch). Gelenk: „Das war mit Hilfe. Jetzt macht ihr es allein, mit drei Artikeln: 6, 11 und 17, jeweils in den drei Schritten.“ Danach Folie 13.");
  box(s, 0.6, 1.5, 12.15, 1.15, WHITE, { line: RED, lw: 3 });
  s.addText([{ text: "Artikel 3:  ", options: { bold: true, color: RED } }, { text: "„Der Ursprung jeder Herrschaft liegt bei der Nation, also beim Volk.“", options: {} }], { isTextBox: true, x: 0.95, y: 1.5, w: 11.45, h: 1.15, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle", margin: 0 });
  [["1  Meine Worte", "Die Herrschaft geht vom Volk aus, nicht vom König allein.", "2E9E6B"], ["2  Begriff heute", "Volkssouveränität", "3F7CC4"], ["3  Vorher", "Der König bestimmte allein.", "E8A33D"]].forEach(([h, x, col], i) => {
    const y = 2.85 + i * 1.1;
    box(s, 0.6, y, 12.15, 0.95, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.16, h: 0.95, fill: { color: col }, line: { type: "none" } });
    txt(s, h, { x: 0.95, y, w: 3.0, h: 0.95, fontFace: HEAD, fontSize: 27, bold: true, valign: "middle" });
    txt(s, x, { x: 4.0, y, w: 8.6, h: 0.95, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 13 Ü2 selbst ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 2", "Drei Artikel selbst", ["Lies Artikel 6, 11 und 17 auf Seite 1.", "Schreibe in eigenen Worten, was dort steht.", "Kreuze an, wie wir heute dazu sagen.", "Schreibe, was vorher anders war. Nutze den Vorher-Speicher."], "6 Min. · allein · leise",
    "Dient Ziel 1 und 2 (selbstständiges Üben). Selbstständiges Üben (6 Minuten). Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Keine Hilfen mehr, außer dem Wortspeicher (er erklärt jeden Begriff in einem Satz). Durch die Reihen gehen (Heft-Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus, mitlesen, nicht kommentieren; unbenotet). Differenzierung: Wer unsicher ist, beginnt mit dem kurzen Artikel 17 und schreibt bei Schritt 1 und 3 nur einen Satzteil. Wer früh fertig ist, erklärt seinem Nachbarn einen Artikel. Nach 6 Minuten: Folie 14. Die Übungsphase endet 9 Minuten vor Schluss (Auswerten läuft immer; Sollbruchstelle: nur Artikel 6 und 11).");
  phaseTag(s, 3);
}

// ============ 14 Kontrolle Begriffe ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "So heißt es heute", 3, "Dient Ziel 2 (Kontrolle der Übersetzung). Kontrolle (1,5 Minuten). Ansage: „Vergleicht und verbessert mit Grün.“ Die Folie zeigt die Begriffe zu allen Artikeln. Artikel 2, 4 und 10 stehen nicht auf dem Blatt; nennen Sie sie kurz mündlich (Begriffe für den Test: Grundrechte, Freiheit, Religionsfreiheit). Stichprobe: Rufen Sie 2 Personen auf (Artikel 6, 11, 17). Hinweis: Artikel 2 nennt das Recht auf Widerstand (das, was in Stunde 3 vorkam); der Begriff „Grundrechte“ fasst Freiheit, Eigentum, Sicherheit und Widerstand zusammen. Artikel 2 nennt aber nicht Gewalt; Artikel 4 begrenzt die Freiheit (sie endet, wo sie anderen schadet): wirksam heißt nicht gerechtfertigt (Stunde 3). Gelenk: „Und was war vorher anders? Seht auf die nächste Folie.“");
  [["Art. 1", "Gleichheit"], ["Art. 2", "Grundrechte (auch Widerstand)"], ["Art. 3", "Volkssouveränität"], ["Art. 4", "Freiheit"], ["Art. 6", "Gleichheit vor dem Gesetz"], ["Art. 10", "Religionsfreiheit"], ["Art. 11", "Meinungsfreiheit"], ["Art. 17", "Eigentumsrecht"]].forEach(([a, b], i) => {
    const col = i < 4 ? 0 : 1, row = i % 4;
    const x = 0.6 + col * 6.2, y = 1.5 + row * 1.1;
    box(s, x, y, 5.95, 0.95, row % 2 ? LIGHT : LIGHT2);
    txt(s, a, { x: x + 0.15, y, w: 1.4, h: 0.95, fontFace: HEAD, fontSize: 27, bold: true, color: RED, valign: "middle" });
    txt(s, b, { x: x + 1.6, y, w: 4.3, h: 0.95, fontSize: 27, valign: "middle", bold: true });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 15 Vorher und seitdem ============
{
  const s = blank("V · Verarbeiten · Bilanz", "Was hat sich verbessert?", 3, "Dient Ziel 1 (vorher, nachher, Bedeutung). Bilanz (1,5 Minuten, Sie sprechen, die Klasse vergleicht mit ihrem Blatt). Ansage: „Stifte liegen. Ihr hört nur zu. Vergleicht mit euren Einträgen bei Artikel 6, 11 und 17.“ Die Bedeutung: Die Menschenrechte haben Vorrechte nach Geburt durch gleiche Rechte ersetzt, Willkür des Königs durch Schutz (Freiheit, Meinung, Eigentum), und die Macht liegt beim Volk. Haltung: Das war ein großer Schritt auf dem Papier. Ob die Rechte für alle galten, ist die Frage der nächsten Stunde (Aktivbürger und Passivbürger). Danach Folie 16.");
  const cols = [["Artikel", 1.9, NAVY], ["Vorher", 5.1, "D9534F"], ["Seitdem", 5.15, "2E9E6B"]];
  let x0 = 0.6; const xs = cols.map(([, w]) => { const v = x0; x0 += w; return v; });
  cols.forEach(([n, w, c], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y: 1.45, w, h: 0.65, fill: { color: c }, line: { color: WHITE, width: 1.5 } });
    txt(s, n, { x: xs[i], y: 1.45, w, h: 0.65, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  [["Art. 3", "Der König bestimmt allein.", "Die Macht geht vom Volk aus."], ["Art. 6", "Vorrechte für Klerus und Adel, keine Steuern.", "Das Gesetz gilt für alle gleich."], ["Art. 11", "Der König kontrolliert, was gedruckt wird.", "Man darf frei reden und schreiben."], ["Art. 17", "Bauern geben Ernte und Dienste ab.", "Das Eigentum ist geschützt."]].forEach((row, r) => {
    row.forEach((v, i) => {
      const y = 2.1 + r * 1.0;
      s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y, w: cols[i][1], h: 1.0, fill: { color: r % 2 ? LIGHT : LIGHT2 }, line: { color: WHITE, width: 1.5 } });
      txt(s, v, { x: xs[i] + 0.1, y, w: cols[i][1] - 0.2, h: 1.0, fontSize: 27, bold: i === 0, align: i === 0 ? "center" : "left", valign: "middle", fontFace: i === 0 ? HEAD : BODY });
    });
  });
  bar(s, [{ text: "Bedeutung: ", options: { bold: true } }, { text: "Gleiche Rechte statt Vorrechte.", options: {} }]);
}

// ============ 16 Das Grundgesetz ============
{
  const s = blank("V · Verarbeiten · Transfer", "Und heute? Das Grundgesetz", 3, "Dient Ziel 3 (Transfer). Kurze Erklärung (30 Sekunden, Sie sprechen, Stifte liegen). Ansage: „Stifte liegen. Ihr hört nur zu.“ Das Grundgesetz ist die Verfassung der Bundesrepublik Deutschland, es gilt seit 1949. Am Anfang stehen die Grundrechte (Artikel 1 bis 19). Diese Rechte gelten heute für alle, die in Deutschland leben. Überleitung zur Frage eines Schülers (bitte nur anbringen, wenn es sich anbietet): „Neulich hat jemand gefragt, wofür man Geschichte braucht. Hier ist ein Beispiel.“ Gelenk: „Gleich seht ihr drei Sätze. Überlegt: Aus welchem Jahr?“ Danach Folie 17.");
  s.addImage({ path: IMG + "icon_buch.png", x: 0.7, y: 1.6, w: 2.6, h: 2.6 });
  box(s, 3.7, 1.55, 9.05, 3.0, LIGHT2);
  txt(s, "Das Grundgesetz ist die Verfassung der Bundesrepublik Deutschland. Es gilt seit 1949. Am Anfang stehen die Grundrechte.", { x: 4.0, y: 1.55, w: 8.5, h: 3.0, fontFace: HEAD, fontSize: 30, bold: true, valign: "middle" });
  box(s, 0.6, 4.85, 12.15, 1.15, LIGHT);
  txt(s, "Die Grundrechte gelten heute für alle, die in Deutschland leben.", { x: 0.9, y: 4.85, w: 11.6, h: 1.15, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Gleich: ", options: { bold: true } }, { text: "drei Sätze, aus welchem Jahr?", options: {} }]);
}

// ============ 17 1789 oder Grundgesetz ============
{
  const s = blank("V · Verarbeiten · Aha", "1789 oder Grundgesetz?", 3, "Dient Ziel 3 (Transfer, Aha-Moment). Denkphase (1 Minute, allein, still). Ansage: „Stifte liegen. Ihr hört nur zu. Auf der Folie stehen drei Sätze. Überlegt allein: Stammt der Satz aus der Erklärung von 1789 oder aus dem Grundgesetz? Notiert A, B und C mit 1789 oder GG ins Heft. Raten ist erlaubt.“ Zeigen Sie nicht, dass es Paare gibt. Die Sätze: A „Das Gesetz muss für alle gleich sein.“ (Artikel 6 von 1789), B „Alle Staatsgewalt geht vom Volke aus.“ (Grundgesetz Artikel 20 Absatz 2), C „Das Eigentum und das Erbrecht werden gewährleistet.“ (Grundgesetz Artikel 14 Absatz 1). Stichprobe bei 2 bis 3 Personen: Was habt ihr und woran habt ihr es erkannt? (kurz, kein Gespräch, nichts auflösen). Danach Folie 18.");
  [["A", "Das Gesetz muss für alle gleich sein."], ["B", "Alle Staatsgewalt geht vom Volke aus."], ["C", "Das Eigentum und das Erbrecht werden gewährleistet."]].forEach(([l, x], i) => {
    const y = 1.5 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.3, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, l, 0.9, y + 0.4, 0.5, RED);
    txt(s, "„" + x + "“", { x: 1.7, y, w: 10.8, h: 1.3, fontFace: HEAD, fontSize: 30, valign: "middle" });
  });
  bar(s, [{ text: "Heft: ", options: { bold: true } }, { text: "A, B, C: 1789 oder Grundgesetz?      ", options: {} }, { text: "Raten ist erlaubt.", options: { italic: true, color: ICE } }]);
}

// ============ 18 Das Echo von 1789 ============
{
  const s = blank("V · Verarbeiten · Aha", "Das Echo von 1789", 3, "Dient Ziel 3 (Aha-Moment). Auflösung (1,5 Minuten, Sie sprechen). Ansage: „Stifte liegen. Ihr hört nur zu. Vergleicht mit euren Notizen.“ Zu jedem Satz zeigt die Folie das Gegenstück: links 1789, rechts das Grundgesetz. Lösung: A 1789 (Artikel 6), B Grundgesetz (Artikel 20 Absatz 2), C Grundgesetz (Artikel 14 Absatz 1). Aussage: „Fast niemand konnte es sicher sagen, weil die Sätze so ähnlich sind. Was ihr als Recht aus 1789 gelesen habt, steht fast wörtlich im Grundgesetz.“ Den Wortlaut des Grundgesetzes (hier leicht gekürzt) mit dem Original abgleichen. Danach Folie 19.");
  const pairs = [["1789 · Art. 6", "Das Gesetz muss für alle gleich sein.", "Grundgesetz Art. 3 (1)", "Alle Menschen sind vor dem Gesetz gleich."], ["1789 · Art. 3", "Der Ursprung jeder Herrschaft liegt beim Volk.", "Grundgesetz Art. 20 (2)", "Alle Staatsgewalt geht vom Volke aus."], ["1789 · Art. 17", "Das Eigentum ist ein unverletzliches und heiliges Recht.", "Grundgesetz Art. 14 (1)", "Das Eigentum und das Erbrecht werden gewährleistet."]];
  pairs.forEach(([a1, t1, a2, t2], i) => {
    const y = 1.4 + i * 1.6;
    box(s, 0.6, y, 5.6, 1.5, i % 2 ? LIGHT : LIGHT2);
    box(s, 6.95, y, 5.8, 1.5, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 6.27, y: y + 0.5, w: 0.6, h: 0.5, fill: { color: RED }, line: { type: "none" } });
    s.addText([{ text: a1 + "\n", options: { bold: true, color: RED, fontSize: 20, fontFace: BODY } }, { text: t1, options: { fontSize: 27, fontFace: HEAD, color: NAVY } }], { isTextBox: true, x: 0.8, y, w: 5.25, h: 1.5, valign: "middle", margin: 0 });
    s.addText([{ text: a2 + "\n", options: { bold: true, color: "2E9E6B", fontSize: 20, fontFace: BODY } }, { text: t2, options: { fontSize: 27, fontFace: HEAD, color: NAVY } }], { isTextBox: true, x: 7.1, y, w: 5.45, h: 1.5, valign: "middle", margin: 0 });
  });
  bar(s, [{ text: "Aha: ", options: { bold: true } }, { text: "Links 1789, rechts Grundgesetz.", options: {} }]);
}

// ============ 19 Bedeutung von 1789 für das Grundgesetz ============
{
  const s = blank("V · Verarbeiten · Bedeutung", "Was bedeutet 1789 für unser Grundgesetz?", 3, "Dient Ziel 3 (Bedeutung erkennen). Kurzes Gespräch mit Stichproben (1 Minute, nur 2 bis 3 Antworten, kein Gespräch). Ansage: „Stifte liegen. Ich rufe euch auf. Was fällt euch auf? Was bedeutet 1789 für unser Grundgesetz?“ Erst die Frage zeigen, die Antwort unten abdecken. Erwartung (Auswahl): Die Erklärung von 1789 war ein Vorbild. Viele Rechte, die wir heute selbstverständlich haben, sind damals zum ersten Mal aufgeschrieben worden. Danach die Antwort zeigen: „1789 war ein Vorbild: Rechte aus der Erklärung stehen noch heute im Grundgesetz.“ Haltung: Heute gelten sie für alle, 1789 galten sie nicht für alle (Stunde 5). Gelenk: „Das schreibt ihr jetzt ab.“ Danach Folie 20.");
  box(s, 0.6, 1.5, 12.15, 2.2, NAVY);
  txt(s, "Was fällt euch auf? Was bedeutet 1789 für unser Grundgesetz?", { x: 1.0, y: 1.5, w: 11.35, h: 2.2, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.95, 12.15, 2.0, LIGHT2);
  s.addText([{ text: "Antwort: ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 27 } }, { text: "1789 war ein Vorbild. Rechte aus der Erklärung stehen noch heute im Grundgesetz.", options: { fontFace: HEAD, fontSize: 28, color: NAVY, bold: true } }], { isTextBox: true, x: 1.0, y: 3.95, w: 11.35, h: 2.0, valign: "middle", margin: 0 });
}

// ============ 20 Merksatz ============
{
  const s = merkSlide(1, "Vorrechte weg, Menschenrechte da",
    "Am **4. August 1789** fallen die Vorrechte, am **26. August** folgt die **Erklärung der Menschen- und Bürgerrechte**: **Gleichheit vor dem Gesetz**, **Freiheit**, **Eigentumsrecht**, **Volkssouveränität**. Vieles steht noch heute im **Grundgesetz**.",
    "Artikel 6 (1789): Das Gesetz muss für alle gleich sein. Grundgesetz Artikel 3: Alle Menschen sind vor dem Gesetz gleich.",
    "**Recht** gilt für alle, **Vorrecht** nur für wenige (nach Geburt).",
    "Dient Ziel 1 bis 3 (Ergebnissicherung, Wissen für den Test). Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“ Warten, bis alle den Stift hingelegt haben (2 Minuten). Der Merksatz sichert Daten, heutige Begriffe, die Bedeutung (Recht statt Vorrecht) und das Echo im Grundgesetz. Gelenk: „Jetzt beantwortet ihr die Frage vom Anfang.“ Danach Folie 21.");
  phaseTag(s, 4);
}

// ============ 21 Exit-Ticket ============
{
  const s = blank("A · Auswerten", "Exit-Ticket: Die Ausgangsfrage", 4, "Dient Ziel 1, 2 und 3 (Abschlussaufgabe, prüft alle Ziele; Sicherung). Ablauf (5 Minuten, ein Satz und ein Kreuz). Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Klerus und Adel haben gefragt: ‚Was wollt ihr noch?‘ Ihr seid der Dritte Stand. Ihr bekommt einen Zettel. Jetzt arbeitet ihr allein. Beendet den Satz auf dem Zettel und kreuzt unten an, welcher Grundgesetz-Artikel zu eurem Satz passt. Den Zettel gebt ihr beim Hinausgehen ab.“ Die Zettel liegen bis dahin umgedreht (6 pro A4, ausgeschnitten, 26 Stück). 4 Minuten schreiben, Heft-Stichprobe beim Durchgehen. Danach 1 Minute: 2 bis 3 Personen lesen ihre Antwort vor (freiwillig, kein Kommentar). Wer hat umgedacht? Vergleich mit der Vermutung vom Anfang. Erwartung (Auswahl, keine Musterlösung): „Wir wollen Menschenrechte, damit wir endlich alle vor dem Gesetz gleich sind / mitbestimmen dürfen / frei reden dürfen / unser Eigentum behalten.“ Der Zettel prüft Ziel 1 (Satz: Was hat sich verbessert?) und Ziel 3 (Ankreuzen: Welcher Grundgesetz-Artikel passt zu meinem Satz? Art. 3 Gleichheit vor dem Gesetz, Art. 5 Meinungsfreiheit, Art. 14 Eigentum, Art. 20 Macht vom Volk). Ziel 2 zeigt sich im Arbeitsblatt. Passt das Kreuz zum Satz (zum Beispiel „alle gleich“ und Art. 3), ist Stufe 3 erreicht. Rückmeldung: Sie lesen die Zettel nach der Stunde. Die freiwillige Zeile „Das war mir unklar“ ist die Rückmeldung der Klasse (ein Satz). Selbst gezeichnete Figuren sind keine historischen Quellen. Unbenotet.");
  s.addImage({ path: IMG + "karikatur_rueckschau.png", x: 0.6, y: 1.5, w: 4.2, h: 3.183 });
  bubble(s, 5.4, 1.45, 7.35, 1.95, WHITE, NAVY, "left", 2.4);
  txt(s, "Klerus und Adel: " + SAGT, { x: 5.7, y: 1.45, w: 6.85, h: 1.95, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  box(s, 5.4, 3.55, 7.35, 1.5, LIGHT2);
  txt(s, "Du, Dritter Stand:  „Wir wollen Menschenrechte, damit wir endlich …“", { x: 5.65, y: 3.55, w: 6.95, h: 1.5, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  box(s, 5.4, 5.15, 7.35, 0.85, LIGHT);
  txt(s, "Kreuze an: Grundgesetz Art. 3, 5, 14 oder 20", { x: 5.55, y: 5.15, w: 7.15, h: 0.85, fontFace: HEAD, fontSize: 27, valign: "middle" });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "4 Min. · allein · leise      ", options: {} }, { text: "Zettel beim Hinausgehen abgeben.", options: { italic: true, color: ICE } }]);
}

// ============ 22 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Dient Ziel 1 bis 3 (Zielprüfung). Die Klasse prüft still: Kann ich das? Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen (1 Minute). Wenn Ziel 3 (*) nicht mehr geschafft wurde, kommt es in Stunde 5 als Wiederholung. Danach Folie 23.");
  [["Ich kann erklären, was sich durch den Verzicht auf die Vorrechte und die Menschenrechte für den Dritten Stand verbessert hat."], ["Ich kann Artikel von 1789 in heutige Begriffe übersetzen (zum Beispiel Eigentumsrecht)."], ["Ich kann an einem Grundgesetz-Artikel zeigen, dass die Erklärung bis heute nachwirkt.*"]].forEach(([x], i) => {
    const y = 1.6 + i * 1.45;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.35, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Exit-Ticket abgeben.", options: {} }]);
}

// ============ 23 Ausblick ============
{
  const s = blank("A · Auswerten · Brücke zu Stunde 5", "Wie geht es weiter?", 4, "Brücke zur Reihe (kein Ziel der Stunde). Folie nur lesen (30 Sekunden). Ausblick: In Artikel 1 steht „die Menschen“, in Artikel 6 „alle Bürger“. Heute gilt das für alle. Galt es 1789 auch für alle: für Frauen, für Arme, für Sklaven in den Kolonien? Das sehen wir in der nächsten Stunde bei der Verfassung von 1791 (Wahlrecht, Aktivbürger und Passivbürger). Vormerken für Stunde 5: Der König weigerte sich zunächst, die Erklärung zu unterschreiben. Erst nach dem Marsch der Frauen nach Versailles im Oktober 1789 unterschrieb er (zum Beispiel als Einstieg in Stunde 5). Danach folgt der erste Test (Stunden 1 bis 5 am 9.12.).");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Für wen galten diese Rechte 1789 wirklich?", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Nächste Stunde: ", options: { bold: true } }, { text: "Verfassung von 1791.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Stunde4_Menschenrechte_8D.pptx" }).then(() => console.log("ok"));
