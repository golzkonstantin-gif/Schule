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

const bubbleTop = (s, text) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.85, y: 1.55, w: 6.5, h: 1.75, rectRadius: 0.18, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: 3.75, y: 3.25, w: 0.6, h: 0.5, rotate: 180, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.77, y: 3.22, w: 0.56, h: 0.1, fill: { color: WHITE }, line: { type: "none" } });
  txt(s, text, { x: 1.05, y: 1.55, w: 6.1, h: 1.75, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle" });
};
const SAGT = "„Wir sind euch entgegengekommen. Jetzt sind alle gleich. Was wollt ihr noch?“";

// ============ 1 Karikatur / Ankommen ============
{
  const s = blank("Ankommen", "Ein Gespräch in Versailles", 0, "Die Folie steht beim Betreten. Niemand muss etwas sagen. Die Karikatur ist selbst gezeichnet, keine historische Quelle. Die Aussage ist bewusst überspitzt: eine erfundene Selbstdarstellung von Klerus und Adel (großzügig, alles erledigt), die die Klasse am Ende selbst widerlegt (nicht freiwillig, sondern aus Angst vor den Aufständen; nicht gleich). Klerus und Adel sprechen, der Dritte Stand (Bauern und Bürger) hört zu. Sie sagen: „Wir sind euch entgegengekommen. Jetzt sind alle gleich. Was wollt ihr noch?“ Hintergrund: In der Nacht zum 4. August 1789 beschloss die Nationalversammlung, die Vorrechte abzuschaffen (Sie erklären das gleich in Phase 2). Die Klasse schreibt ihre Vermutung ins Heft (bleibt stehen, wird am Ende mit dem Ergebnis verglichen): „Reicht das dem Dritten Stand? Ich vermute …“ Keine Besprechung. Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2 bis 5 in zusammen etwa 2 Minuten.");
  s.addImage({ path: IMG + "karikatur_drei_staende.png", x: 0.6, y: 1.45, w: 7.0, h: 4.375 });
  bubbleTop(s, SAGT);
  [["Klerus", 0.75, 1.75], ["Adel", 3.3, 1.4], ["Dritter Stand", 5.05, 2.45]].forEach(([n, x, w]) => {
    box(s, x, 5.2, w, 0.52, WHITE, { line: NAVY, lw: 2 });
    txt(s, n, { x, y: 5.2, w, h: 0.52, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  [["Heft raus. Datum und Überschrift.", 1.6, 1.0], ["Lies, was Klerus und Adel sagen. Der Dritte Stand hört zu.", 2.75, 1.4], ["Schreibe: Reicht das? Ich vermute …", 4.5, 1.5]].forEach(([x, y, h], i) => {
    numCircle(s, i + 1, 7.95, y + 0.05, 0.45, RED);
    txt(s, x, { x: 8.6, y, w: 4.15, h, fontSize: 27, bold: true, valign: "top" });
  });
}

// ============ 2 Erwartungen ============
{
  const s = blank("A · Ankommen · Zusammenarbeit", "Das ist mir wichtig", undefined, "Dauer: 20 Sekunden (bewährt), Sie sprechen, die Klasse hört nur zu, kein Gespräch. Ansage: „Stifte liegen. Ihr hört nur zu.“ Haltung (Neue Autorität): ruhig, präsent, wertschätzend. Sinngemäß: „Ich bin euer Lehrer und euer Coach. Mir ist wichtig, dass ihr Spaß habt und etwas lernt. Beides gelingt nur gemeinsam. Wir hören einander zu. Wir arbeiten leise, wenn Arbeitszeit ist. Wir machen mit, auch wenn es mal schwer ist. Ich bin da, ich helfe euch, und ich bleibe dran.“ Bei Hospitation unverändert lassen, damit die Stunde vertraut wirkt.");
  box(s, 0.6, 1.5, 12.15, 1.15, NAVY2);
  txt(s, "Ich bin euer Lehrer und Coach. Mir ist wichtig, dass ihr Spaß habt und lernt.", { x: 0.9, y: 1.5, w: 11.6, h: 1.15, fontFace: HEAD, fontSize: 27, bold: true, color: WHITE, valign: "middle" });
  txt(s, "Das gelingt, wenn wir gemeinsam arbeiten:", { x: 0.6, y: 2.8, w: 12.1, h: 0.55, fontSize: 27, bold: true, color: NAVY });
  ["Wir hören einander zu, wenn jemand spricht.", "Wir arbeiten leise, wenn Arbeitszeit ist.", "Wir machen mit, auch wenn es mal schwer ist."].forEach((x, i) => {
    const y = 3.45 + i * 0.88;
    box(s, 0.6, y, 12.15, 0.76, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, i + 1, 0.8, y + 0.17, 0.42, RED);
    txt(s, x, { x: 1.4, y, w: 11.2, h: 0.76, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Ich bin da und bleibe dran. ", options: { bold: true } }, { text: "Ich helfe euch.", options: {} }]);
}

// ============ 3 Stundenfrage ============
{
  const s = blank("Stundenfrage", "Das untersuchen wir heute", undefined, "Stundenfrage und Begriff (45 Sekunden, Klasse hört nur zu): „Klerus und Adel sagen: Wir sind euch entgegengekommen, jetzt sind alle gleich. Stimmt das? Wir prüfen es an einer Quelle: der Erklärung der Menschenrechte von 1789. Eine Quelle ist ein Text aus der Zeit selbst. Wir lesen sie nicht einfach, wir werten sie aus: Was steht da? Was war vorher anders?“ Wichtig: Die Frage lässt offen, was die Erklärung für den Dritten Stand (Bauern und Bürger) bedeutet. Keine Antworten ins Plenum, die Vermutungen stehen im Heft.");
  box(s, 0.6, 1.5, 12.15, 2.2, NAVY);
  txt(s, "Was steht in der Erklärung der Menschenrechte, und was ändert sie für den Dritten Stand (Bauern und Bürger)?", { x: 1.0, y: 1.5, w: 11.35, h: 2.2, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.95, 12.15, 2.0, LIGHT2);
  s.addText([{ text: "Quelle: ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 27 } }, { text: "Ein Text aus der Zeit selbst. Wir lesen ihn nicht nur, wir werten ihn aus.", options: { fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: 1.0, y: 3.95, w: 11.35, h: 2.0, valign: "middle", margin: 0 });
}

// ============ 4 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Lernziele vorlesen (30 Sekunden). Mindestziel: Ziel 1 und 2. Ziel 3 hat einen Stern: Das Urteil wird verkürzt, wenn die Zeit knapp ist. Die Phase „Auswerten“ bleibt immer. Am Ende (Folie 20) prüft die Klasse still, ob die Ziele erreicht sind.");
  [["1", "Ich kann Artikel der Erklärung der Menschenrechte in eigenen Worten wiedergeben."], ["2", "Ich kann einen Artikel mit der Ständeordnung vergleichen: Was war vorher anders?"], ["3", "Ich kann mit einem Beleg aus der Quelle begründen, welcher Artikel für den Dritten Stand am wichtigsten ist.*"]].forEach(([n, x], i) => {
    const y = 1.75 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.37, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Rückblick, Quelle auswerten, Urteil.", options: {} }]);
}

// ============ 5 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Ablauf zeigen (30 Sekunden). Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Ziel 3). Insgesamt 45 Minuten.");
  const ph = [["A", "Ankommen", "Karikatur", "4 Min."], ["V", "Vorwissen", "Rückblick 1789", "5 Min."], ["I", "Informieren", "Quelle auswerten", "9 Min."], ["V", "Verarbeiten", "Arbeitsblatt, Urteil", "20 Min."], ["A", "Auswerten", "Antwort als Dritter Stand", "7 Min."]];
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

// ============ 6 Rückblick: Wahr oder falsch ============
{
  const { s } = taskSlide("V · Vorwissen aktivieren · Rückblick", "Das wissen wir schon: richtig oder falsch?", [
    "Der Dritte Stand verlor die Abstimmung nach Ständen, obwohl er fast alle Menschen vertritt.",
    "Beim Ballhausschwur schworen die Abgeordneten, nicht auseinanderzugehen, bis es eine Verfassung gibt.",
    "Der Sturm auf die Bastille war ein gewaltloser Protest.",
    "Der Ballhausschwur fand am 14. Juli 1789 statt.",
  ], "2 Min. allein, 1 Min. Partner, 1 Min. Plenum", "Rückblick (5 Minuten, allein, Partnercheck, Plenum). Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Aufgabe: Schreibe W oder F ins Heft, bei falschen Aussagen die Korrektur. Dann: „Flüsterstufe. Nur dein Partner hört dich. Vergleicht.“ Dann Stichproben bei 3 bis 4 Personen unterschiedlichen Niveaus (kurze Antwort, kein Gespräch, unbenotet), dann Folie 7. Hinweis: Aussage 3 und 4 sind falsch. Aussage 4 ist ein typischer Fehler (Datum). Wurde in Stunde 3 etwas nicht geschafft (Ziel 2 mit Stern), steht es hier als Wiederholung: Passen Sie Aussage 3 an. Gelenk: „1789 hat der Dritte Stand gehandelt. Aber was haben Klerus und Adel danach getan? Das lest ihr jetzt.“", "Fertig? Stift hinlegen.", 12.15);
  phaseTag(s, 1);
}

// ============ 7 Auflösung Rückblick ============
{
  const s = blank("V · Vorwissen aktivieren · Kontrolle", "Das stimmt", 1, "Auflösung (1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ 1 richtig (Klerus und Adel haben zusammen zwei von drei Stimmen), 2 richtig, 3 falsch (gewaltsam, etwa 100 Tote unter den Angreifern, siehe Stunde 3), 4 falsch (20. Juni 1789; 14. Juli ist die Bastille). Gelenk zu Phase 2: „Der Dritte Stand hat gehandelt. Und was passiert danach? Hört zu.“");
  [["1", "richtig", JA_(), "Klerus und Adel haben zusammen zwei von drei Stimmen."], ["2", "richtig", JA_(), "Eid im Ballhaus, 20. Juni 1789."], ["3", "falsch", NEIN_(), "Die Bastille war gewaltsam: Schüsse, etwa 100 Tote."], ["4", "falsch", NEIN_(), "Ballhausschwur: 20. Juni. Bastille: 14. Juli."]].forEach(([n, v, col, x], i) => {
    const y = 1.5 + i * 1.1;
    box(s, 0.6, y, 12.15, 0.95, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.8, y + 0.22, 0.5, NAVY);
    box(s, 1.55, y + 0.15, 1.9, 0.65, col);
    txt(s, v, { x: 1.55, y: y + 0.15, w: 1.9, h: 0.65, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, x, { x: 3.65, y, w: 9.0, h: 0.95, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Verbessere falsche Aussagen mit Grün.", options: {} }]);
}
function JA_() { return "2E9E6B"; }
function NEIN_() { return "D9534F"; }

// ============ 8 Zeitstrahl Sommer 1789 (Portion 1) ============
{
  const s = blank("I · Informieren · Portion 1", "Was geschah nach dem 14. Juli?", 2, "Portion 1 (3 Minuten, Sie sprechen, Klasse hört nur zu; Stifte liegen). Ansage: „Stifte liegen. Ihr hört nur zu.“ Inhalt: Nach der Bastille wächst die Angst vor Aufständen. Im ganzen Land greifen Bauern Schlösser an und verbrennen Urkunden über ihre Abgaben. In der Nacht zum 4. August 1789 beschließt die Nationalversammlung, Vorrechte abzuschaffen (zum Beispiel die Steuerfreiheit von Klerus und Adel, Abgaben und Dienste der Bauern an den Adel, das alleinige Jagdrecht). Am 26. August 1789 beschließt sie die Erklärung der Menschen- und Bürgerrechte. Vorbild ist die amerikanische Unabhängigkeitserklärung von 1776. Warnung vor dem Fehler: Abgeschafft heißt nicht, dass sofort alle gleich leben. Viele Abgaben müssen die Bauern erst ablösen, also bezahlen (Lehrbuch gegenprüfen). Prüffrage (nach der Portion, Stichprobe bei 2 Personen): „Warum verzichten Klerus und Adel am 4. August?“ Erwartung: aus Angst vor den Aufständen der Bauern.");
  const ev = [["14. Juli 1789", "Sturm auf die Bastille"], ["Sommer 1789", "Bauern greifen Schlösser an"], ["4. August 1789", "Vorrechte werden abgeschafft"], ["26. August 1789", "Erklärung der Menschenrechte"]];
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 1.7, w: 0.1, h: 3.7, fill: { color: NAVY }, line: { type: "none" } });
  ev.forEach(([d, e], i) => {
    const y = 1.5 + i * 1.0;
    box(s, 0.6, y, 2.8, 0.8, i % 2 ? LIGHT : LIGHT2);
    txt(s, d, { x: 0.7, y, w: 2.6, h: 0.8, fontSize: 27, bold: true, color: RED, align: "center", valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: y + 0.23, w: 0.34, h: 0.34, fill: { color: i >= 2 ? RED : NAVY }, line: { color: WHITE, width: 2 } });
    txt(s, e, { x: 4.2, y, w: 8.5, h: 0.8, fontFace: HEAD, fontSize: 27, bold: i >= 2, valign: "middle" });
  });
  bar(s, [{ text: "Achtung: ", options: { bold: true } }, { text: "Abgeschafft heißt nicht, dass sofort alles anders ist.", options: {} }]);
}

// ============ 9 Drei Schritte ============
{
  const s = blank("I · Informieren · Portion 2", "Eine Quelle auswerten: drei Schritte", 2, "Portion 2, Teil 1 (1 Minute). Ansage: „Stifte liegen. Ihr hört nur zu.“ Wir lesen die Erklärung der Menschenrechte nicht nur, wir werten sie aus. Immer in dieser Reihenfolge: 1 Lesen: Was steht da? In eigenen Worten, so, dass es ein Freund versteht. 2 Einordnen: Was war vorher anders? Wer gewinnt, wer verliert dabei? Das wissen wir aus den Stunden 1 bis 3 (Ständeordnung, Vorrechte, König bestimmt allein). 3 Begriff: Welcher Fachbegriff passt? (Wortspeicher auf dem Blatt). Warnung vor dem Fehler: Nicht den Wortlaut abschreiben, sondern in eigene Worte fassen.");
  [["1", "Lesen", "Was steht da? In eigenen Worten.", "2E9E6B"], ["2", "Einordnen", "Was war vorher anders? Wer gewinnt?", "E8A33D"], ["3", "Begriff", "Welcher Fachbegriff passt?", "3F7CC4"]].forEach(([n, h, x, col], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 1.6, 3.95, 3.9, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 1.6, w: 3.95, h: 0.16, fill: { color: col }, line: { type: "none" } });
    numCircle(s, n, xx + 0.3, 1.95, 0.6, RED);
    txt(s, h, { x: xx + 1.1, y: 1.95, w: 2.7, h: 0.6, fontFace: HEAD, fontSize: 28, bold: true, valign: "middle" });
    txt(s, x, { x: xx + 0.3, y: 2.9, w: 3.4, h: 2.4, fontSize: 27 });
  });
  bar(s, [{ text: "Drei Schritte, ", options: { bold: true } }, { text: "immer in dieser Reihenfolge.", options: {} }]);
}

// ============ 10 Modell Artikel 1 ============
{
  const s = blank("I · Informieren · Portion 2 · Modell", "Ich mache es vor: Artikel 1", 2, "Portion 2, Teil 2 (3 Minuten, Sie denken laut). Ansage: „Stifte liegen. Ihr hört nur zu. Ich mache es vor, ihr schaut zu.“ Zeigen Sie zuerst nur den Artikel im roten Rahmen (decken Sie die drei Zeilen ab, zum Beispiel mit einem Blatt oder indem Sie die Folie erst später weiter zeigen). Denken Sie laut: „Schritt 1: Was steht da? ‚Frei und gleich an Rechten geboren.‘ Das heißt: Alle Menschen haben von Geburt an dieselben Rechte. Schritt 2: Was war vorher anders? In der Ständeordnung entschied die Geburt über den Stand und die Vorrechte. Wer gewinnt? Der Dritte Stand. Schritt 3: Der Begriff dafür heißt Gleichheit.“ Zeigen Sie dann die drei Zeilen. Auf dem Blatt (Seite 1) steht dasselbe gelb hinterlegt. Warnung vor dem Fehler: ‚Gleich an Rechten‘ heißt nicht, dass alle gleich reich oder gleich stark sind.");
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

// ============ 11 Prüffrage ============
{
  const s = blank("I · Informieren · Prüffrage", "Prüffrage", 2, "Prüffrage (1 Minute, unbenotet, getrennt von den mündlichen Noten). Ansage: „Stifte liegen. Ich rufe euch auf. Antwortet in einem Satz.“ Rufen Sie 4 bis 6 Personen unterschiedlichen Niveaus auf, kurze Antworten, kein Gespräch. Frage 1: Was machst du im Schritt 2? Erwartung: Ich frage, was vorher anders war und wer gewinnt. Frage 2: Warum schreibe ich den Artikel nicht einfach ab? Erwartung: Ich muss ihn in eigene Worte fassen, so zeigt sich, ob ich ihn verstanden habe. Wird die Prüffrage nicht richtig beantwortet, wiederholen Sie Schritt 2 am Artikel 1, bevor die Klasse beginnt. Es wird nicht weitergegangen, wenn etwas nicht verstanden ist.");
  box(s, 0.6, 1.6, 12.15, 1.7, NAVY);
  txt(s, "Was machst du im Schritt 2 „Einordnen“?", { x: 1.0, y: 1.6, w: 11.35, h: 1.7, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.6, 12.15, 1.7, LIGHT2);
  txt(s, "Warum schreibst du den Artikel nicht einfach ab?", { x: 1.0, y: 3.6, w: 11.35, h: 1.7, fontFace: HEAD, fontSize: 32, bold: true, valign: "middle" });
  bar(s, [{ text: "Ich rufe auf. ", options: { bold: true } }, { text: "Antworte in einem Satz.", options: {} }]);
}

// ============ 12 Raster ============
{
  const s = blank("V · Verarbeiten · Rückmeldung", "So bekommst du Rückmeldung", 3, "Raster vorstellen (45 Sekunden), nicht diskutieren. Ansage: „Stifte liegen. Ihr hört nur zu.“ Die Abgabe des Blattes ist freiwillig. Wer abgibt, erhält Rückmeldung nach den drei Stufen. Punkte oder Noten gibt es heute nicht (kommt später). Keine Selbsteinschätzung der Schülerinnen und Schüler. Auf dem Blatt (Seite 2) kreuzen Sie die Stufe an.");
  [["Stufe 1", "Ich kann einen Artikel in eigenen Worten wiedergeben."], ["Stufe 2", "Ich kann einen Artikel mit der Ständeordnung vergleichen."], ["Stufe 3", "Ich kann mit einem Beleg aus der Quelle begründen, welcher Artikel am wichtigsten ist."]].forEach(([n, x], i) => {
    const y = 1.5 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.3, i % 2 ? LIGHT : LIGHT2);
    txt(s, n, { x: 0.8, y, w: 2.3, h: 1.3, fontFace: HEAD, fontSize: 27, bold: true, color: RED, valign: "middle" });
    txt(s, x, { x: 3.2, y, w: 9.4, h: 1.3, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Abgabe freiwillig: ", options: { bold: true } }, { text: "Du bekommst Rückmeldung von mir.", options: {} }]);
}

// ============ 13 Ü1 Artikel 3 ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 1", "Artikel 3 mit Hilfe", ["Arbeitsblatt, Seite 1. Lies Artikel 3.", "Fülle „Lesen“ und „Einordnen“ aus. Nutze die Hilfe.", "Der Begriff steht schon da.", "Vergleiche kurz mit deinem Partner."], "3 Min. · allein, dann Partner · leise",
    "Angeleitetes Üben (3 Minuten). Ansage: „Stifte liegen. Ihr hört nur zu. Jetzt nehmt ihr das Arbeitsblatt. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Das Arbeitsblatt liegt vorab auf den Tischen (umgedreht, jede Person ein Blatt, doppelseitig). Nach 2 Minuten: „Flüsterstufe. Nur dein Partner hört dich. Vergleicht.“ Durch die Reihen gehen (Heft-Stichprobe, nicht kommentieren). Warnung vor dem Fehler: Hilfen ausfüllen, nicht den Satz aus dem Artikel kopieren. Danach Folie 14.");
  phaseTag(s, 3);
}

// ============ 14 Kontrolle Artikel 3 ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "So sieht Artikel 3 aus", 3, "Kontrolle (1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ Stichprobe bei 2 Personen (kurz vorlesen, was sie geschrieben haben, kein Gespräch). Gelenk zu Übung 2: „Das war mit Hilfe. Jetzt macht ihr es allein, mit drei Artikeln.“ Warnung: Artikel 3 passt zum Ballhausschwur (Macht geht von der Nation aus). Das knüpft an die letzte Stunde an. Danach Folie 15.");
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

// ============ 15 Ü2 Artikel 2, 4, 6 ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 2", "Artikel 2, 4 und 6 selbst", ["Lies Artikel 2, 4 und 6 auf Seite 1.", "Fülle „Lesen“, „Einordnen“ und „Begriff“ aus.", "Nutze den Wortspeicher.", "Fertig? Zusatz* auf Seite 2: Artikel 10 und 11."], "6 Min. · allein · leise",
    "Selbstständiges Üben (6 Minuten). Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Keine Hilfen mehr. Durch die Reihen gehen (Heft-Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus, mitlesen, nicht kommentieren; unbenotet). Differenzierung: Wer unsicher ist, beginnt mit Artikel 4 (kurzer Satz). Wer früh fertig ist, bearbeitet das Zusatz* (Artikel 10 und 11). Nach 6 Minuten: Folie 16. Ende der Übungsphase ist fest (Auswerten läuft immer, Sollbruchstelle: Urteil auf zwei Sätze kürzen).");
  phaseTag(s, 3);
}

// ============ 16 Ü partnercheck ============
{
  const { s } = taskSlide("V · Verarbeiten · Partnercheck", "Vergleicht zu zweit", ["Lest euch Artikel 2, 4 und 6 vor.", "Wo habt ihr etwas anderes geschrieben?", "Entscheidet gemeinsam, was besser passt.", "Verbessert mit Grün."], "3 Min. · zu zweit · Flüsterstufe",
    "Partnercheck (3 Minuten). Ansage: „Stifte liegen. Flüsterstufe. Nur dein Partner hört dich. Vergleicht Artikel 2, 4 und 6. Verbessert mit Grün.“ Jede Person liest vor, nicht nur abschreiben. Sie gehen durch die Reihen und hören bei 4 bis 6 Paaren kurz zu, ohne zu helfen. Sollbruchstelle: Schaffen Paare nur zwei Artikel, genügt das. Nach 3 Minuten: Folie 17.");
  phaseTag(s, 3);
}

// ============ 17 Kontrolle Artikel 2, 4, 6 ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "So sehen Artikel 2, 4 und 6 aus", 3, "Kontrolle (3 Minuten). Ansage: „Vergleicht und verbessert mit Grün.“ Die Folie zeigt Stichworte, andere richtige Formulierungen zählen. Stichprobe: Rufen Sie 2 Personen auf, die ihren Eintrag vorlesen. Wichtig (Haltung, Warnung vor dem Fehler): Artikel 2 nennt das Recht auf Widerstand, aber nicht Gewalt. Das Recht auf Freiheit endet nach Artikel 4, wo man anderen schadet. Das knüpft an Stunde 3 an: Wirksam heißt nicht gerechtfertigt. Sagen Sie: „Wir beschreiben, was dort steht. Wir erklären nicht, dass Gewalt gut ist.“ Gelenk: „Welcher Artikel ist für den Dritten Stand am wichtigsten? Das entscheidet ihr jetzt.“");
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
  bar(s, [{ text: "Achtung: ", options: { bold: true } }, { text: "Artikel 2 nennt Widerstand, nicht Gewalt.", options: {} }]);
}

// ============ 18 Ü3 Urteil ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 3", "Dein Urteil", ["Wähle einen Artikel, der für den Dritten Stand am wichtigsten ist.", "Schreibe auf Seite 2 mit den Satzanfängen.", "Gib einen Beleg aus der Quelle an (Zitat).", "Nenne auch ein Gegenargument."], "5 Min. · allein · leise",
    "Urteil (5 Minuten). Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Es gibt kein richtiges Urteil. Sachurteil mit Beleg (Zitat aus dem Artikel), dann Werturteil mit offengelegtem Maßstab (Satzanfänge auf Seite 2: „Artikel … sagt: …“, „Das ändert für den Dritten Stand, dass …“, „Ich halte Artikel … für am wichtigsten, weil …“, „Dagegen spricht, dass …“). Heft- oder Blatt-Stichprobe bei 4 bis 6 Personen (mitlesen, nicht kommentieren, unbenotet). Differenzierung: Wer unsicher ist, schreibt nur die ersten beiden Satzanfänge. Wird die Zeit knapp: Urteil auf zwei Sätze kürzen (Ziel 3 hat einen Stern). Danach: „Wer das Blatt abgeben möchte, legt es jetzt vorn ab.“ Nach 5 Minuten: Folie 19.");
  phaseTag(s, 3);
}

// ============ 19 Abschlussfrage ============
{
  const s = blank("A · Auswerten", "Abschlussfrage: Du bist der Dritte Stand", 4, "Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Ihr seid der Dritte Stand (Bauern und Bürger). Klerus und Adel haben gesagt: ‚Wir sind euch entgegengekommen. Jetzt sind alle gleich. Was wollt ihr noch?‘ Denkt zehn Sekunden nach. Ich rufe euch auf.“ Die Antwort kommt mündlich, in der Ich-Form, ohne Schnipsel: „Ich antworte euch: …, weil Artikel … sagt …“ Mehrere gute Antworten. Geprüft werden Position, ein Beleg aus der Quelle (Artikel) und ein „weil“. Erwartung (Auswahl, keine Musterlösung): „Das reicht uns nicht, weil Artikel 3 sagt, dass die Herrschaft vom Volk ausgeht, und wir wollen mitbestimmen.“ „Das ist ein Anfang, weil Artikel 6 sagt, dass das Gesetz für alle gleich ist. Aber dann müssen auch alle gleich zahlen.“ „Wir geben uns nicht zufrieden, weil auf dem Papier alles gleich ist, aber die Bauern noch Abgaben ablösen müssen.“ Form nach Verlauf: Blitzlicht (4 bis 6 Personen gezielt aufrufen, kein Kommentar) oder Duell zu zweit (leise). Vergleich mit der Vermutung vom Anfang: Wer hat umgedacht? Selbst gezeichnete Figuren sind keine historischen Quellen. Unbenotet.");
  s.addImage({ path: IMG + "karikatur_drei_staende.png", x: 0.6, y: 1.5, w: 5.4, h: 3.375 });
  box(s, 0.6, 5.0, 5.4, 0.75, WHITE, { line: NAVY, lw: 2 });
  txt(s, "Klerus · Adel · Dritter Stand", { x: 0.6, y: 5.0, w: 5.4, h: 0.75, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  bubble(s, 6.4, 1.5, 6.35, 2.2, WHITE, NAVY, "left", 2.6);
  txt(s, "Klerus und Adel: " + SAGT, { x: 6.7, y: 1.5, w: 5.85, h: 2.2, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  bubble(s, 6.4, 3.95, 6.35, 1.3, "FDF1D8", GOLD, "left", 4.6);
  txt(s, "Du, Dritter Stand:  „Ich antworte euch: …, weil Artikel … sagt …“", { x: 6.7, y: 3.95, w: 5.85, h: 1.3, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  [["Position", 1.8], ["Beleg (Artikel)", 2.8], ["„weil“", 1.3]].reduce((xx, [lab, ww], i) => {
    box(s, xx, 5.4, ww, 0.62, i % 2 ? LIGHT : LIGHT2);
    s.addText([{ text: lab, options: { bold: true, fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: xx + 0.05, y: 5.4, w: ww - 0.1, h: 0.62, valign: "middle", align: "center", margin: 0 });
    return xx + ww + 0.15;
  }, 6.4);
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "4 Min. · erst denken, dann mündlich      ", options: {} }, { text: "Alle denken mit.", options: { italic: true, color: ICE } }]);
}

// ============ 20 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Die Klasse prüft still: Kann ich das? Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen (30 Sekunden). Wenn Ziel 3 (*) nicht mehr geschafft wurde, kommt es in Stunde 5 als Wiederholung. Wer das Blatt abgeben möchte, legt es vorn ab. Danach Folie 21.");
  [["Ich kann Artikel der Erklärung der Menschenrechte in eigenen Worten wiedergeben."], ["Ich kann einen Artikel mit der Ständeordnung vergleichen: Was war vorher anders?"], ["Ich kann mit einem Beleg aus der Quelle begründen, welcher Artikel für den Dritten Stand am wichtigsten ist.*"]].forEach(([x], i) => {
    const y = 1.7 + i * 1.4;
    box(s, 0.6, y, 12.15, 1.15, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.15, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Blatt freiwillig abgeben.", options: {} }]);
}

// ============ 21 Ausblick ============
{
  const s = blank("A · Auswerten · Ausblick", "Wie geht es weiter?", 4, "Folie nur lesen (30 Sekunden). Ausblick: In Artikel 1 steht „die Menschen“, in Artikel 6 „alle Bürger“. Gilt das wirklich für alle: für Frauen, für Arme, für Sklaven in den Kolonien? Das sehen wir in der nächsten Stunde bei der Verfassung von 1791 (Wahlrecht, Aktivbürger und Passivbürger). Danach folgt der erste Test (Stunden 1 bis 5 am 9.12.).");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Für wen gelten diese Rechte wirklich?", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Nächste Stunde: ", options: { bold: true } }, { text: "Verfassung von 1791.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Stunde4_Menschenrechte_8D.pptx" }).then(() => console.log("ok"));
