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
  const s = blank("Ankommen", "Ein Gespräch in Versailles", 0, "Dient Ziel: Stundenfrage und Vermutung (Vorher-Nachher). Die Folie steht beim Betreten. Niemand muss etwas sagen. Die Karikatur ist selbst gezeichnet, keine historische Quelle. An der Wand hängen zwei Bilder: der Eid im Ballhaus (links) und der Sturm auf die Bastille (rechts), beide unbeschriftet. Davor stehen Klerus und Adel und sagen: „Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?“ Der Dritte Stand (Bauern und Bürger) hört zu. Die Aussage ist bewusst überspitzt (erfundene Selbstdarstellung): Entgegenkommen heißt Gnade von oben. Gegenpol, der am Ende der Stunde auftaucht: Rechte, die für alle gelten. Die Klasse schreibt ihre Vermutung ins Heft (bleibt stehen, wird am Ende mit der Antwort verglichen): „Was will der Dritte Stand noch? Ich vermute …“ Keine Besprechung. Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2 bis 4 in zusammen etwa 2 Minuten.");
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
  const s = blank("Stundenfrage", "Das untersuchen wir heute", undefined, "Dient Ziel: alle drei (Orientierung). Stundenfrage und Begriff (45 Sekunden, Klasse hört nur zu): „Klerus und Adel sagen: Wir kamen euch entgegen. Was wollt ihr noch? Die Antwort steht in einer Quelle: der Erklärung der Menschenrechte von 1789. Eine Quelle ist ein Text aus der Zeit selbst. Wir lesen sie nicht einfach, wir werten sie aus: Was steht da? Was war vorher anders? Das braucht ihr bei jeder Quelle und im Test.“ Relevanz benennen, nicht diskutieren. Die Frage lässt offen, was die Erklärung für den Dritten Stand bedeutet. Keine Antworten ins Plenum, die Vermutungen stehen im Heft.");
  box(s, 0.6, 1.5, 12.15, 2.2, NAVY);
  txt(s, "Was will der Dritte Stand (Bauern und Bürger) noch? Was steht dazu in der Erklärung der Menschenrechte?", { x: 1.0, y: 1.5, w: 11.35, h: 2.2, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.95, 12.15, 2.0, LIGHT2);
  s.addText([{ text: "Quelle: ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 27 } }, { text: "Ein Text aus der Zeit selbst. Wir werten ihn aus. Das brauchst du bei jeder Quelle und im Test.", options: { fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: 1.0, y: 3.95, w: 11.35, h: 2.0, valign: "middle", margin: 0 });
}

// ============ 3 Lernziele (zugleich Raster) ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Dient Ziel: alle drei (Transparenz). Lernziele vorlesen (45 Sekunden). Sie sind zugleich das Raster für die Rückmeldung: Stufe 1 = Ziel 1, Stufe 2 = Ziel 2, Stufe 3 = Ziel 3. Die Abgabe des Blattes ist freiwillig, Sie geben Rückmeldung nach den Stufen, Punkte oder Noten gibt es heute nicht. Keine Selbsteinschätzung der Schülerinnen und Schüler. Mindestziel: Ziel 1 und 2. Ziel 3 hat einen Stern: Das Urteil wird gekürzt, wenn die Zeit knapp ist. Die Phase „Auswerten“ bleibt immer. Am Ende (Folie 19) prüft die Klasse still, ob die Ziele erreicht sind.");
  [["1", "Ich kann Artikel der Erklärung der Menschenrechte in eigenen Worten wiedergeben und den passenden Begriff nennen."], ["2", "Ich kann einen Artikel mit der Ständeordnung vergleichen und einem Ereignis von 1789 zuordnen."], ["3", "Ich kann mit Beleg und Maßstab begründen, welcher Artikel für den Dritten Stand am wichtigsten ist.*"]].forEach(([n, x], i) => {
    const y = 1.55 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.3, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.4, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.3, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Rückmeldung von mir: ", options: { bold: true } }, { text: "Stufe 1, 2 oder 3. Blatt freiwillig abgeben.", options: {} }]);
}

// ============ 4 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Dient Ziel: alle drei (Transparenz). Ablauf zeigen (30 Sekunden). Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Ziel 3). Insgesamt 45 Minuten, davon eine kleine Reserve in „Informieren“ zum erneuten Erklären.");
  const ph = [["A", "Ankommen", "Karikatur", "4 Min."], ["V", "Vorwissen", "Rückschau", "5 Min."], ["I", "Informieren", "Quelle auswerten", "8 Min."], ["V", "Verarbeiten", "Blatt, Urteil", "20 Min."], ["A", "Auswerten", "Merksatz, Antwort", "8 Min."]];
  ph.forEach(([l, n, x, m], i) => {
    const xx = 0.6 + i * 2.45;
    box(s, xx, 1.9, 2.3, 3.6, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fill: { color: RED }, line: { type: "none" } });
    txt(s, l, { x: xx + 0.75, y: 2.1, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, n, { x: xx, y: 3.1, w: 2.3, h: 0.5, fontFace: HEAD, fontSize: 27, bold: true, align: "center" });
    txt(s, x, { x: xx + 0.1, y: 3.7, w: 2.1, h: 1.0, fontSize: 27, align: "center", color: MUTED });
    txt(s, m, { x: xx, y: 4.8, w: 2.3, h: 0.5, fontSize: 27, bold: true, color: RED, align: "center" });
  });
  bar(s, [{ text: "Wichtig: ", options: { bold: true } }, { text: "Wir arbeiten leise. Gesprochen wird nur nach Ansage.", options: {} }]);
}

// ============ 5 Rückschau: Bilder ============
{
  const s = blank("V · Vorwissen aktivieren · Rückschau", "Rückschau: Was meinen Klerus und Adel?", 1, "Dient Ziel 2 (Ereignisse von 1789 kennen, um sie später Artikeln zuzuordnen). Rückschau (5 Minuten mit Folie 6, allein, Partnercheck, Plenum). Ansage: „Stifte liegen. Ihr hört nur zu. Schaut auf die zwei Bilder. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Die Bilder sind unbeschriftet: Die Klasse entschlüsselt „geschworen“ und „gestürmt“ selbst (links der Eid im Ballhaus, 20. Juni, rechts der Sturm auf die Bastille, 14. Juli). 2 Minuten allein ins Heft: zu jedem Bild „Was geschah? Was bewirkte es?“ Dann 1 Minute Partnercheck: „Flüsterstufe. Nur dein Partner hört dich. Erklärt euch: Was haben die Bilder mit dem Entgegenkommen zu tun?“ Dann 1 Minute Plenum: Stichproben bei 3 bis 4 Personen unterschiedlichen Niveaus (kurze Antwort, unbenotet), dann Folie 6. Wurde in Stunde 3 etwas nicht geschafft (zum Beispiel die Reihenfolge am Zeitstrahl), steht es hier als Wiederholung. Gelenk: „Die Bilder haben mit dem Entgegenkommen zu tun. Wie genau, seht ihr gleich.“");
  crop(s, "L", 0.6, 1.5);
  crop(s, "R", 0.6, 1.5 + PH + 0.12);
  [["Schaue auf die zwei Bilder. Was meinen „geschworen“ und „gestürmt“?", 1], ["Schreibe zu jedem Bild: Was geschah? Was bewirkte es?", 2], ["Erkläre deinem Partner: Was haben die Bilder mit dem Entgegenkommen zu tun?", 3]].forEach(([x, n], i) => {
    const ys = [1.5, 2.85, 4.1], hs = [1.2, 1.1, 1.75];
    box(s, 4.6, ys[i], 8.15, hs[i], i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 4.8, ys[i] + (hs[i] - 0.45) / 2, 0.45, NAVY);
    txt(s, x, { x: 5.5, y: ys[i], w: 7.1, h: hs[i], fontSize: 27, bold: true, valign: "middle" });
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "2 Min. allein, 1 Min. Partner, 1 Min. Plenum      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}

// ============ 6 Kette und Portion 1 ============
{
  const s = blank("V · I · Kontrolle und Portion 1", "So hängt es zusammen", 1, "Dient Ziel 2 (Ereignisse 1789 und Rahmen für das „vorher anders“). Zwei Teile auf einer Folie (3 Minuten). Teil 1, Kontrolle der Rückschau (1 Minute): Ansage: „Vergleicht und verbessert mit Grün.“ Zeilen 1 bis 3: 20. Juni Ballhausschwur (die Versammlung bleibt zusammen, gewaltlos), 14. Juli Bastille (der König zieht die Soldaten ab), im Sommer greifen Bauern Schlösser an. Teil 2, Informieren, Portion 1 (2 Minuten, Sie sprechen, Stifte liegen): Ansage: „Stifte liegen. Ihr hört nur zu.“ Zeilen 4 und 5: In der Nacht zum 4. August 1789 beschließt die Nationalversammlung, die Vorrechte abzuschaffen (zum Beispiel die Steuerfreiheit von Klerus und Adel, Abgaben und Dienste der Bauern an den Adel). Klerus und Adel kamen der Klasse also entgegen, weil sie Angst vor den Aufständen hatten: Das war nicht freiwillig. Entgegenkommen ist Gnade von oben. Am 26. August 1789 folgt die Erklärung der Menschen- und Bürgerrechte, Vorbild ist die amerikanische Unabhängigkeitserklärung von 1776. Warnung vor dem Fehler: Abgeschafft heißt nicht, dass sofort alle gleich leben (viele Abgaben müssen die Bauern noch ablösen, im Lehrbuch gegenprüfen); und nicht der König, sondern die Nationalversammlung schaffte die Vorrechte ab. Prüffrage (Stichprobe, 2 Personen): „Warum kommen Klerus und Adel am 4. August entgegen?“ Erwartung: aus Angst vor den Aufständen der Bauern. Gelenk: „Entgegenkommen ist Gnade. Was ist der Gegenpol? Das lest ihr jetzt in der Quelle.“");
  const ev = [["20. Juni 1789", "Ballhausschwur: Die Versammlung bleibt zusammen.", false], ["14. Juli 1789", "Bastille: Der König zieht die Soldaten ab.", false], ["Sommer 1789", "Bauern greifen Schlösser an.", false], ["4. August 1789", "Vorrechte werden abgeschafft.", true], ["26. August 1789", "Erklärung der Menschenrechte", true]];
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 1.7, w: 0.1, h: 4.4, fill: { color: NAVY }, line: { type: "none" } });
  ev.forEach(([d, e, hl], i) => {
    const y = 1.45 + i * 0.93;
    box(s, 0.6, y, 2.8, 0.78, i % 2 ? LIGHT : LIGHT2);
    txt(s, d, { x: 0.7, y, w: 2.6, h: 0.78, fontSize: 27, bold: true, color: RED, align: "center", valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: y + 0.22, w: 0.34, h: 0.34, fill: { color: hl ? RED : NAVY }, line: { color: WHITE, width: 2 } });
    txt(s, e, { x: 4.2, y, w: 8.55, h: 0.78, fontFace: HEAD, fontSize: 27, bold: hl, valign: "middle" });
  });
  bar(s, [{ text: "Achtung: ", options: { bold: true } }, { text: "Abgeschafft heißt nicht, dass sofort alles anders ist.", options: {} }]);
}

// ============ 7 Drei Schritte ============
{
  const s = blank("I · Informieren · Portion 2", "Eine Quelle auswerten: drei Schritte", 2, "Dient Ziel 1 und 2 (Methode der Auswertung). Portion 2, Teil 1 (1 Minute). Ansage: „Stifte liegen. Ihr hört nur zu.“ Wir lesen die Erklärung der Menschenrechte nicht nur, wir werten sie aus. Immer in dieser Reihenfolge: 1 Lesen: Was steht da? In eigenen Worten, so, dass es ein Freund versteht. 2 Einordnen: Was war vorher anders? Wer gewinnt, wer verliert dabei? Passt ein Bild von 1789 dazu? Das wissen wir aus den Stunden 1 bis 3 (Ständeordnung, Vorrechte, König bestimmt allein). 3 Begriff: Welcher Fachbegriff passt? (Wortspeicher auf dem Blatt). Warnung vor dem Fehler: Nicht den Wortlaut abschreiben, sondern in eigene Worte fassen.");
  [["1", "Lesen", "Was steht da? In eigenen Worten.", "2E9E6B"], ["2", "Einordnen", "Was war vorher anders? Passt ein Bild?", "E8A33D"], ["3", "Begriff", "Welcher Fachbegriff passt?", "3F7CC4"]].forEach(([n, h, x, col], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 1.6, 3.95, 3.9, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 1.6, w: 3.95, h: 0.16, fill: { color: col }, line: { type: "none" } });
    numCircle(s, n, xx + 0.3, 1.95, 0.6, RED);
    txt(s, h, { x: xx + 1.1, y: 1.95, w: 2.7, h: 0.6, fontFace: HEAD, fontSize: 28, bold: true, valign: "middle" });
    txt(s, x, { x: xx + 0.3, y: 2.9, w: 3.4, h: 2.4, fontSize: 27 });
  });
  bar(s, [{ text: "Drei Schritte, ", options: { bold: true } }, { text: "immer in dieser Reihenfolge.", options: {} }]);
}

// ============ 8 Modell Artikel 1 ============
{
  const s = blank("I · Informieren · Portion 2 · Modell", "Ich mache es vor: Artikel 1", 2, "Dient Ziel 1 und 2 (Modell des Zielprodukts). Portion 2, Teil 2 (3 Minuten, Sie denken laut). Ansage: „Stifte liegen. Ihr hört nur zu. Ich mache es vor, ihr schaut zu.“ Zeigen Sie zuerst nur den Artikel im roten Rahmen (die drei Zeilen abdecken oder erst später aufdecken). Denken Sie laut: „Schritt 1: Was steht da? ‚Frei und gleich an Rechten geboren.‘ Das heißt: Alle Menschen haben von Geburt an dieselben Rechte. Schritt 2: Was war vorher anders? In der Ständeordnung entschied die Geburt über den Stand und die Vorrechte. Wer gewinnt? Der Dritte Stand. Schritt 3: Der Begriff dafür heißt Gleichheit.“ Zeigen Sie dann die drei Zeilen. Auf dem Blatt (Seite 1) steht dasselbe gelb hinterlegt. Warnung vor dem Fehler: ‚Gleich an Rechten‘ heißt nicht gleich reich oder gleich stark. Reserve (1 Minute, nur bei Bedarf): Schritt 2 noch einmal erklären.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.5, w: 12.15, h: 1.35, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  s.addText([{ text: "Artikel 1:  ", options: { bold: true, color: RED } }, { text: "„Die Menschen werden frei und gleich an Rechten geboren und bleiben es.“", options: {} }], { isTextBox: true, x: 0.95, y: 1.5, w: 11.45, h: 1.35, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle", margin: 0 });
  [["1  Lesen", "Alle Menschen haben von Geburt an dieselben Rechte.", "2E9E6B"], ["2  Einordnen", "Vorher entschied die Geburt über Stand und Vorrechte. Gewinner: Dritter Stand.", "E8A33D"], ["3  Begriff", "Gleichheit", "3F7CC4"]].forEach(([h, x, col], i) => {
    const y = 3.0 + i * 1.05;
    box(s, 0.6, y, 12.15, 0.92, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.16, h: 0.92, fill: { color: col }, line: { type: "none" } });
    txt(s, h, { x: 0.95, y, w: 2.6, h: 0.92, fontFace: HEAD, fontSize: 27, bold: true, valign: "middle" });
    txt(s, x, { x: 3.6, y, w: 9.0, h: 0.92, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Du schaust zu. Stifte liegen.", options: {} }]);
}

// ============ 9 Prüffrage ============
{
  const s = blank("I · Informieren · Prüffrage", "Prüffrage", 2, "Dient Ziel 1 und 2 (Verständnisprüfung vor dem Üben). Prüffrage (1 Minute, unbenotet, getrennt von den mündlichen Noten). Ansage: „Stifte liegen. Ich rufe euch auf. Antwortet in einem Satz.“ Rufen Sie 4 bis 6 Personen unterschiedlichen Niveaus auf, kurze Antworten, kein Gespräch. Frage 1: Was machst du im Schritt 2? Erwartung: Ich frage, was vorher anders war, wer gewinnt und ob ein Bild passt. Frage 2: Warum schreibe ich den Artikel nicht einfach ab? Erwartung: Ich muss ihn in eigene Worte fassen, so zeigt sich, ob ich ihn verstanden habe. Wird die Prüffrage nicht richtig beantwortet, wiederholen Sie Schritt 2 am Artikel 1 (hier liegt die Reserve), bevor die Klasse beginnt. Es wird nicht weitergegangen, wenn etwas nicht verstanden ist. Gelenk: „Jetzt macht ihr es selbst, zuerst mit Hilfe.“");
  box(s, 0.6, 1.6, 12.15, 1.7, NAVY);
  txt(s, "Was machst du im Schritt 2 „Einordnen“?", { x: 1.0, y: 1.6, w: 11.35, h: 1.7, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.6, 12.15, 1.7, LIGHT2);
  txt(s, "Warum schreibst du den Artikel nicht einfach ab?", { x: 1.0, y: 3.6, w: 11.35, h: 1.7, fontFace: HEAD, fontSize: 32, bold: true, valign: "middle" });
  bar(s, [{ text: "Ich rufe auf. ", options: { bold: true } }, { text: "Antworte in einem Satz.", options: {} }]);
}

// ============ 10 Ü1 Artikel 3 ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 1", "Artikel 3 mit Hilfe", ["Arbeitsblatt, Seite 1. Lies Artikel 3.", "Fülle „Lesen“ und „Einordnen“ aus. Nutze die Hilfe.", "Der Begriff steht schon da.", "Vergleiche kurz mit deinem Partner."], "2 Min. allein, 1 Min. Partner · leise",
    "Dient Ziel 1 und 2 (erste Anwendung mit Hilfe). Angeleitetes Üben (3 Minuten mit Folie 11). Ansage: „Stifte liegen. Ihr hört nur zu. Jetzt nehmt ihr das Arbeitsblatt. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Das Arbeitsblatt liegt vorab auf den Tischen (umgedreht, jede Person ein Blatt, doppelseitig). Nach 2 Minuten: „Flüsterstufe. Nur dein Partner hört dich. Vergleicht.“ Durch die Reihen gehen (Heft-Stichprobe, nicht kommentieren). Warnung vor dem Fehler: Hilfen ausfüllen, nicht den Satz aus dem Artikel kopieren. Danach Folie 11.");
  phaseTag(s, 3);
}

// ============ 11 Kontrolle Artikel 3 ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "So sieht Artikel 3 aus", 3, "Dient Ziel 1 und 2 (Kontrolle der ersten Anwendung). Kontrolle (30 Sekunden bis 1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ Stichprobe bei 2 Personen (kurz vorlesen, was sie geschrieben haben, kein Gespräch). Gelenk: „Das war mit Hilfe. Jetzt macht ihr es allein, mit drei Artikeln.“ Artikel 3 passt zum Ballhausschwur (Macht geht von der Nation aus). Danach Folie 12.");
  box(s, 0.6, 1.5, 12.15, 1.15, WHITE, { line: RED, lw: 3 });
  s.addText([{ text: "Artikel 3:  ", options: { bold: true, color: RED } }, { text: "„Der Ursprung jeder Herrschaft liegt bei der Nation, also beim Volk.“", options: {} }], { isTextBox: true, x: 0.95, y: 1.5, w: 11.45, h: 1.15, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle", margin: 0 });
  [["1  Lesen", "Die Herrschaft geht vom Volk aus, nicht vom König allein.", "2E9E6B"], ["2  Einordnen", "Vorher bestimmte der König allein. Passt zum Ballhausschwur.", "E8A33D"], ["3  Begriff", "Volkssouveränität", "3F7CC4"]].forEach(([h, x, col], i) => {
    const y = 2.85 + i * 1.1;
    box(s, 0.6, y, 12.15, 0.95, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.16, h: 0.95, fill: { color: col }, line: { type: "none" } });
    txt(s, h, { x: 0.95, y, w: 2.6, h: 0.95, fontFace: HEAD, fontSize: 27, bold: true, valign: "middle" });
    txt(s, x, { x: 3.6, y, w: 9.0, h: 0.95, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 12 Ü2 Artikel 2, 4, 6 ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 2", "Artikel 2, 4 und 6 selbst", ["Lies Artikel 2, 4 und 6 auf Seite 1.", "Fülle „Lesen“, „Einordnen“ und „Begriff“ aus.", "Nutze den Wortspeicher.", "Fertig? Zusatz* auf Seite 2: Artikel 10 und 11."], "6 Min. · allein · leise",
    "Dient Ziel 1 und 2 (selbstständiges Üben). Selbstständiges Üben (6 Minuten). Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Keine Hilfen mehr, außer Wortspeicher. Durch die Reihen gehen (Heft-Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus, mitlesen, nicht kommentieren; unbenotet). Differenzierung: Wer unsicher ist, beginnt mit Artikel 4 (kurzer Satz). Wer früh fertig ist, bearbeitet das Zusatz* (Artikel 10 und 11). Nach 6 Minuten: Folie 13. Ende der Übungsphase ist fest (Auswerten läuft immer; Sollbruchstelle: Urteil auf zwei Sätze kürzen).");
  phaseTag(s, 3);
}

// ============ 13 Partnercheck und Zuordnung ============
{
  const { s } = taskSlide("V · Verarbeiten · Partnercheck", "Vergleicht und ordnet zu", ["Lest euch Artikel 2, 4 und 6 vor. Verbessert mit Grün.", "Welches Bild passt zu Artikel 3, welches zu Artikel 2?", "Begründet mit einem Wort aus dem Artikel.", "Prüft: Was sagt Artikel 4 dazu, wie weit man gehen darf?"], "4 Min. · zu zweit · Flüsterstufe",
    "Dient Ziel 2 (Einordnen: einem Ereignis zuordnen). Partnercheck mit Zuordnung (4 Minuten). Ansage: „Stifte liegen. Flüsterstufe. Nur dein Partner hört dich. Vergleicht Artikel 2, 4 und 6. Dann ordnet ihr zu: Welches Bild an der Wand der Karikatur passt zu welchem Artikel? Begründet mit einem Wort aus dem Artikel.“ Die Bilder sind auf Folie 5 und 14 zu sehen (hier nicht gezeigt, damit die Klasse sich erinnern muss; bei Bedarf Folie 5 einblenden). Sie gehen durch die Reihen und hören bei 4 bis 6 Paaren kurz zu, ohne zu helfen. Sollbruchstelle: Reicht die Zeit nicht, genügt der Vergleich und eine Zuordnung. Nach 4 Minuten: Folie 14.");
  phaseTag(s, 3);
}

// ============ 14 Kontrolle Artikel 2, 4, 6 ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "So sehen Artikel 2, 4 und 6 aus", 3, "Dient Ziel 1 und 2 (Kontrolle). Kontrolle (2 Minuten). Ansage: „Vergleicht und verbessert mit Grün.“ Die Folie zeigt Stichworte, andere richtige Formulierungen zählen. Stichprobe: Rufen Sie 2 Personen auf, die ihren Eintrag vorlesen. Gelenk: „Und welches Bild passt zu welchem Artikel? Seht auf die nächste Folie.“ Danach Folie 15.");
  const cols = [["", 1.7, NAVY], ["Lesen", 3.8, "2E9E6B"], ["Einordnen", 4.35, "E8A33D"], ["Begriff", 2.3, "3F7CC4"]];
  let x0 = 0.6; const xs = cols.map(([, w]) => { const v = x0; x0 += w; return v; });
  cols.forEach(([n, w, c], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y: 1.45, w, h: 0.65, fill: { color: c }, line: { color: WHITE, width: 1.5 } });
    txt(s, n, { x: xs[i], y: 1.45, w, h: 0.65, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  [["Art. 2", "Der Staat muss Rechte schützen, auch das Recht auf Widerstand.", "Vorher: Ungehorsam war Verbrechen.", "Grundrechte"], ["Art. 4", "Jeder darf tun, was anderen nicht schadet.", "Vorher: Einsperren ohne Urteil.", "Freiheit"], ["Art. 6", "Gleiches Gesetz für alle.", "Vorher: Steuerfreiheit, Ämter nur für Adel.", "Gleichheit"]].forEach((row, r) => {
    row.forEach((v, i) => {
      const y = 2.1 + r * 1.35;
      s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y, w: cols[i][1], h: 1.35, fill: { color: r % 2 ? LIGHT : LIGHT2 }, line: { color: WHITE, width: 1.5 } });
      txt(s, v, { x: xs[i] + 0.1, y, w: cols[i][1] - 0.2, h: 1.35, fontSize: 27, bold: i === 0 || i === 3, align: i === 0 || i === 3 ? "center" : "left", valign: "middle", fontFace: i === 0 ? HEAD : BODY });
    });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 15 Aha: Bild und Artikel ============
{
  const s = blank("V · Verarbeiten · Aha", "Was sie taten, wird zum Recht", 3, "Dient Ziel 2 (Ereignis zuordnen) und bereitet Ziel 3 vor (Beleg, Maßstab). Aha-Moment (1 Minute, Sie sprechen, Klasse vergleicht mit ihrer Zuordnung aus dem Partnercheck). Ansage: „Stifte liegen. Ihr hört nur zu. Vergleicht mit eurer Zuordnung.“ Ballhaus passt zu Artikel 3: Die Herrschaft geht vom Volk aus (Volkssouveränität); die Abgeordneten handelten gegen den Willen des Königs. Bastille passt zu Artikel 2: Das Recht auf Widerstand gegen Unterdrückung. Artikel 2 nennt aber nicht Gewalt. Haltung: Wir beschreiben, was dort steht, wir erklären nicht, dass Gewalt gut ist. Artikel 4 begrenzt das Ganze: Die Freiheit endet, wo sie anderen schadet. Das knüpft an Stunde 3 an: wirksam heißt nicht gerechtfertigt. Die Ereignisse von 1789 stehen ab dem 26. August als Recht in der Erklärung. Gelenk: „Welcher Artikel ist für den Dritten Stand am wichtigsten? Das entscheidet ihr jetzt.“ Danach Folie 16.");
  crop(s, "L", 0.6, 1.5);
  crop(s, "R", 0.6, 1.5 + PH + 0.12);
  [[0, "Artikel 3", "Die Herrschaft geht vom Volk aus."], [1, "Artikel 2", "Recht auf Widerstand gegen Unterdrückung."]].forEach(([i, h, x]) => {
    const y = 1.5 + i * (PH + 0.12);
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 4.3, y: y + PH / 2 - 0.25, w: 0.6, h: 0.5, fill: { color: NAVY }, line: { type: "none" } });
    box(s, 5.1, y, 7.65, PH, i % 2 ? LIGHT : LIGHT2);
    txt(s, h, { x: 5.35, y: y + 0.12, w: 7.2, h: 0.55, fontFace: HEAD, fontSize: 27, bold: true, color: RED });
    txt(s, x, { x: 5.35, y: y + 0.75, w: 7.2, h: PH - 0.85, fontSize: 27 });
  });
  bar(s, [{ text: "Artikel 4: ", options: { bold: true } }, { text: "Die Freiheit endet, wo sie anderen schadet.", options: {} }]);
}

// ============ 16 Ü3 Urteil ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 3", "Dein Urteil", ["Wähle einen Artikel, der für den Dritten Stand am wichtigsten ist.", "Schreibe auf Seite 2 mit den Satzanfängen.", "Gib Beleg und Maßstab an. Was heißt „wichtig“?", "Nenne auch ein Gegenargument."], "4 Min. · allein · leise",
    "Dient Ziel 3 (Beleg und Maßstab). Urteil (4 Minuten). Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Es gibt kein richtiges Urteil. Sachurteil mit Beleg (Zitat aus dem Artikel), dann Werturteil mit offengelegtem Maßstab (Satzanfänge auf Seite 2: „Artikel … sagt: …“, „Das ändert für den Dritten Stand, dass …“, „Mein Maßstab: Wichtig ist für mich, was …“, „Ich halte Artikel … für am wichtigsten, weil …“, „Dagegen spricht, dass …“). Mögliche Maßstäbe stehen als Bausteine auf dem Blatt (betrifft die meisten Menschen, ändert am meisten, schützt vor Willkür und Gewalt). Das Gegenargument ist die zweite Perspektive; es kann die Sicht von Klerus und Adel sein. Heft- oder Blatt-Stichprobe bei 4 bis 6 Personen (mitlesen, nicht kommentieren, unbenotet). Differenzierung: Wer unsicher ist, schreibt nur die ersten drei Satzanfänge. Wird die Zeit knapp: Urteil auf zwei Sätze kürzen (Ziel 3 hat einen Stern). Nach 4 Minuten: Folie 17.");
  phaseTag(s, 3);
}

// ============ 17 Merksatz ============
{
  const s = merkSlide(1, "Erklärung der Menschen- und Bürgerrechte",
    "Am **4. August 1789** schafft die Nationalversammlung die Vorrechte ab, am **26. August** beschließt sie die **Erklärung der Menschen- und Bürgerrechte**: **Freiheit**, **Gleichheit**, Recht auf **Widerstand**, Macht beim Volk (**Volkssouveränität**).",
    "Artikel 6: Das Gesetz gilt für alle gleich. Früher zahlten Klerus und Adel keine Steuern.",
    "**Recht** gilt für alle, **Vorrecht** nur für wenige (nach Geburt).",
    "Dient Ziel 1 und 2 (Ergebnissicherung, Wissen für den Test). Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“ Warten, bis alle den Stift hingelegt haben (2 Minuten). Der Merksatz sichert Datum, Begriffe und den Gegensatz Recht und Vorrecht. Gelenk: „Jetzt antwortet ihr Klerus und Adel.“ Danach Folie 18.");
  phaseTag(s, 4);
}

// ============ 18 Abschlussfrage ============
{
  const s = blank("A · Auswerten", "Abschlussfrage: Du bist der Dritte Stand", 4, "Dient Ziel 1, 2 und 3 (Abschlussaufgabe, prüft alle Ziele). Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Ihr seid der Dritte Stand (Bauern und Bürger). Klerus und Adel haben gesagt: ‚Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?‘ Denkt zehn Sekunden nach. Ich rufe euch auf.“ Die Antwort kommt mündlich, in der Ich-Form, ohne Schnipsel: „Ich antworte euch: Wir wollen …, weil Artikel … sagt …, früher war …“ Mehrere gute Antworten. Geprüft werden Position, Beleg aus der Quelle (Artikel mit eigenen Worten), der Vergleich mit „früher“ und ein „weil“. Erwartung (Auswahl, keine Musterlösung): „Wir wollen keine Gnade, sondern Rechte, weil Artikel 6 sagt, dass das Gesetz für alle gleich ist, früher hatten Klerus und Adel Vorrechte.“ „Wir wollen mitbestimmen, weil Artikel 3 sagt, dass die Herrschaft vom Volk ausgeht, früher bestimmte der König allein.“ „Wir wollen frei sein, weil Artikel 4 sagt, dass man alles tun darf, was anderen nicht schadet, früher konnte der König jemanden ohne Urteil einsperren.“ Form nach Verlauf: Blitzlicht (4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar) oder Duell zu zweit (leise: eine Person Klerus und Adel, eine Dritter Stand). Vergleich mit der Vermutung vom Anfang: Wer hat umgedacht? Die Antwort zeigt den Lernzuwachs (Vorher-Nachher). Selbst gezeichnete Figuren sind keine historischen Quellen. Unbenotet.");
  s.addImage({ path: IMG + "karikatur_rueckschau.png", x: 0.6, y: 1.5, w: 5.0, h: 3.789 });
  box(s, 0.6, 5.4, 5.0, 0.7, WHITE, { line: NAVY, lw: 2 });
  txt(s, "Klerus · Adel · Dritter Stand", { x: 0.6, y: 5.4, w: 5.0, h: 0.7, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  bubble(s, 6.4, 1.5, 6.35, 2.1, WHITE, NAVY, "left", 2.5);
  txt(s, "Klerus und Adel: " + SAGT, { x: 6.7, y: 1.5, w: 5.85, h: 2.1, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  bubble(s, 6.4, 3.75, 6.35, 1.55, "FDF1D8", GOLD, "left", 4.5);
  txt(s, "Du, Dritter Stand:  „Ich antworte euch: Wir wollen …, weil Artikel … sagt …“", { x: 6.7, y: 3.75, w: 5.85, h: 1.55, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  [["Position", 1.8], ["Beleg (Artikel)", 2.8], ["„weil“", 1.3]].reduce((xx, [lab, ww], i) => {
    box(s, xx, 5.4, ww, 0.62, i % 2 ? LIGHT : LIGHT2);
    s.addText([{ text: lab, options: { bold: true, fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: xx + 0.05, y: 5.4, w: ww - 0.1, h: 0.62, valign: "middle", align: "center", margin: 0 });
    return xx + ww + 0.15;
  }, 6.4);
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "4 Min. · erst denken, dann mündlich      ", options: {} }, { text: "Alle denken mit.", options: { italic: true, color: ICE } }]);
}

// ============ 19 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Dient Ziel 1, 2 und 3 (Zielprüfung). Die Klasse prüft still: Kann ich das? Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen (30 Sekunden). Wenn Ziel 3 (*) nicht mehr geschafft wurde, kommt es in Stunde 5 als Wiederholung. Wer das Blatt abgeben möchte, legt es vorn ab. Danach Folie 20.");
  [["Ich kann Artikel der Erklärung der Menschenrechte in eigenen Worten wiedergeben und den passenden Begriff nennen."], ["Ich kann einen Artikel mit der Ständeordnung vergleichen und einem Ereignis von 1789 zuordnen."], ["Ich kann mit Beleg und Maßstab begründen, welcher Artikel für den Dritten Stand am wichtigsten ist.*"]].forEach(([x], i) => {
    const y = 1.6 + i * 1.45;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.35, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Blatt freiwillig abgeben.", options: {} }]);
}

// ============ 20 Ausblick ============
{
  const s = blank("A · Auswerten · Brücke zu Stunde 5", "Wie geht es weiter?", 4, "Brücke zur Reihe (kein Ziel der Stunde). Folie nur lesen (30 Sekunden). Ausblick: In Artikel 1 steht „die Menschen“, in Artikel 6 „alle Bürger“. Gilt das wirklich für alle: für Frauen, für Arme, für Sklaven in den Kolonien? Das sehen wir in der nächsten Stunde bei der Verfassung von 1791 (Wahlrecht, Aktivbürger und Passivbürger). Danach folgt der erste Test (Stunden 1 bis 5 am 9.12.).");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Für wen gelten diese Rechte wirklich?", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Nächste Stunde: ", options: { bold: true } }, { text: "Verfassung von 1791.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Stunde4_Menschenrechte_8D.pptx" }).then(() => console.log("ok"));
