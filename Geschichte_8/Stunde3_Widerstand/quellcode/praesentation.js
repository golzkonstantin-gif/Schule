const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Stunde 3: Widerstand (8C)";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Widerstand · Geschichte 8C";
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



// ============ 1 Karikatur / Ankommen ============
{
  const s = blank("Ankommen", "Ein Gespräch in Versailles", 0, "Die Folie steht beim Betreten. Niemand muss etwas sagen. Die Karikatur ist selbst gezeichnet, keine historische Quelle. Drei Figuren: Klerus, Adel und der Dritte Stand (Bauern und Bürger), der zuhört. Der Adlige sagt: „Der Dritte Stand hat abgestimmt und verloren. Jetzt kann er sich nur noch fügen.“ Klerus und Adel sprechen, der Dritte Stand hört zu. Die Klasse schreibt ihre Vermutung ins Heft (bleibt stehen, wird am Ende der Stunde mit dem Ergebnis verglichen): „Was kann der Dritte Stand noch tun? Ich vermute …“ Keine Besprechung, kein Urteil. Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2 bis 5 in zusammen etwa 2 Minuten.");
  s.addImage({ path: IMG + "karikatur_drei_staende.png", x: 0.6, y: 1.45, w: 7.0, h: 4.375 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.85, y: 1.55, w: 6.5, h: 1.55, rectRadius: 0.18, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: 3.75, y: 3.05, w: 0.6, h: 0.5, rotate: 180, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.77, y: 3.02, w: 0.56, h: 0.1, fill: { color: WHITE }, line: { type: "none" } });
  txt(s, "„Der Dritte Stand hat abgestimmt und verloren. Jetzt kann er sich nur noch fügen.“", { x: 1.05, y: 1.55, w: 6.1, h: 1.55, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle" });
  [["Klerus", 0.75, 1.75], ["Adel", 3.3, 1.4], ["Dritter Stand", 5.05, 2.45]].forEach(([n, x, w]) => {
    box(s, x, 5.2, w, 0.52, WHITE, { line: NAVY, lw: 2 });
    txt(s, n, { x, y: 5.2, w, h: 0.52, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  [["Heft raus. Datum und Überschrift.", 1.6, 1.0], ["Lies, was der Adlige sagt. Der Dritte Stand hört zu.", 2.75, 1.4], ["Schreibe: Was kann er noch tun? Ich vermute …", 4.3, 1.5]].forEach(([x, y, h], i) => {
    numCircle(s, i + 1, 7.95, y + 0.05, 0.45, RED);
    txt(s, x, { x: 8.6, y, w: 4.15, h, fontSize: 27, bold: true, valign: "top" });
  });
}

// ============ 2 Erwartungen (Neue Autorität) ============
{
  const s = blank("A · Ankommen · Zusammenarbeit", "Das ist mir wichtig", undefined, "Dauer: 20 Sekunden (bewährt in Stunde 1 vom 04.11., hier nur kurz wiederholen), Sie sprechen, die Klasse hört nur zu, kein Gespräch, keine Abfrage. Ansage: „Stifte liegen. Ihr hört nur zu.“ Haltung (Neue Autorität): ruhig, präsent, wertschätzend, ohne Drohung und ohne Vorwurf. Sie sagen, was Ihnen wichtig ist und wofür Sie selbst einstehen (Präsenz, Beziehung, Beharrlichkeit), und laden zum gemeinsamen „Wir“ ein, statt Forderungen von oben nach unten zu stellen. Sinngemäß: „Ich bin euer Lehrer und euer Coach. Mir ist wichtig, dass ihr hier Spaß habt und etwas lernt. Beides gelingt nur, wenn wir zusammenarbeiten. Dafür sorgen wir gemeinsam: Wir hören einander zu. Wir arbeiten leise, wenn Arbeitszeit ist. Wir machen mit, auch wenn es mal schwer ist. Ich bin da, ich helfe euch, und ich bleibe dran.“ Im Sport kennen sie das: Eine Mannschaft funktioniert nur, wenn alle mitziehen. Die drei Sätze sind ein Vorschlag, bitte nach Ihrer Klasse anpassen. Später in der Stunde knüpfen Sie bei Bedarf kurz und ruhig daran an („Wir hatten vereinbart: …“) statt zu ermahnen.");
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
  const s = blank("Stundenfrage", "Das untersuchen wir heute", undefined, "Stundenfrage und Begriff (45 Sekunden, Klasse hört nur zu): „Der Dritte Stand hat gesehen: Abstimmen bringt nichts. Welche Wege bleiben ihm? Heute geht es um Widerstand. Widerstand heißt: Man wehrt sich gegen die Herrschenden, obwohl man sich fügen soll. Wir vergleichen zwei Wege: den Ballhausschwur und die Bastille.“ Wichtig: Der Begriff ist der Maßstab, an dem beide Texte gelesen werden (Gegen wen? Wie? Mit welchem Preis?). Keine Antworten ins Plenum, die Vermutungen stehen im Heft.");
  box(s, 0.6, 1.5, 12.15, 2.2, NAVY);
  txt(s, "Wie kann sich der Dritte Stand (Bauern und Bürger) wehren, wenn Wählen nichts bringt?", { x: 1.0, y: 1.5, w: 11.35, h: 2.2, fontFace: HEAD, fontSize: 34, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 3.95, 12.15, 2.0, LIGHT2);
  s.addText([{ text: "Widerstand: ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 27 } }, { text: "Man wehrt sich gegen die Herrschenden, obwohl man sich fügen soll.", options: { fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: 1.0, y: 3.95, w: 11.35, h: 1.1, valign: "middle", margin: 0 });
  txt(s, "Wir vergleichen zwei Wege: Ballhausschwur und Bastille.", { x: 1.0, y: 5.0, w: 11.35, h: 0.9, fontFace: HEAD, fontSize: 27, valign: "middle" });
}

// ============ 4 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Lernziele vorlesen (30 Sekunden). Mindestziel: Ziel 1 und 3. Ziel 2 hat einen Stern. Am Ende (Folie 19) prüft die Klasse still, ob die Ziele erreicht sind.");
  [["1", "Ich kann den Ballhausschwur und den Sturm auf die Bastille mit fünf Kriterien vergleichen."], ["2", "Ich kann an der Reihenfolge der Ereignisse erklären, wie die Wege zusammenwirkten.*"], ["3", "Ich kann beurteilen, welcher Weg berechtigt war, und das mit einem Maßstab begründen."]].forEach(([n, x], i) => {
    const y = 1.75 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.37, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Karikatur, Partnerpuzzle, Zeitstrahl, Urteil.", options: {} }]);
}

// ============ 5 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Ablauf zeigen (30 Sekunden). Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Ziel 2). Insgesamt 45 Minuten.");
  const ph = [["A", "Ankommen", "Karikatur", "4 Min."], ["V", "Vorwissen", "Rückblick", "2 Min."], ["I", "Informieren", "Text lesen", "10 Min."], ["V", "Verarbeiten", "Puzzle, Urteil", "24 Min."], ["A", "Auswerten", "Antwort", "5 Min."]];
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

// ============ 6 Wiederholung ============
{
  const s = blank("V · Vorwissen aktivieren · Wiederholung", "Das wissen wir schon", 1, "Wiederholung (2 Minuten). Ansage: „Stifte liegen. Ihr hört nur zu.“ Kurz an den Zeitstrahl und das Ergebnis der Abstimmung erinnern: Nach Ständen verlor der Dritte Stand (Bauern und Bürger) immer, obwohl er fast alle Menschen vertritt; 1789 gab es keine Abstimmung nach Köpfen. Dann: „Erklärt eurem Partner in einem Satz, warum der Dritte Stand die Abstimmung verlor.“ (Partnercheck, 1 Minute, Flüsterstufe: „Nur dein Partner hört dich.“) Danach gezielte Abfrage bei 2 Personen, kein Gespräch. Gelenk zum Puzzle: „Wählen bringt nichts. Welche Wege bleiben? Das lest ihr jetzt.“");
  const ev = [["5. Mai 1789", "Eröffnung, Streit um die Abstimmung"], ["17. Juni 1789", "Nationalversammlung"], ["20. Juni 1789", "Ballhausschwur"], ["Anfang Juli 1789", "Soldaten um Paris, Angst"]];
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 1.7, w: 0.1, h: 3.5, fill: { color: NAVY }, line: { type: "none" } });
  ev.forEach(([d, e], i) => {
    const y = 1.5 + i * 0.95;
    box(s, 0.6, y, 2.8, 0.75, i % 2 ? LIGHT : LIGHT2);
    txt(s, d, { x: 0.7, y, w: 2.6, h: 0.75, fontSize: 27, bold: true, color: RED, align: "center", valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: y + 0.2, w: 0.34, h: 0.34, fill: { color: NAVY }, line: { color: WHITE, width: 2 } });
    txt(s, e, { x: 4.2, y, w: 8.5, h: 0.75, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Partner: ", options: { bold: true } }, { text: "Warum verlor der Dritte Stand die Abstimmung?", options: {} }]);
}

// ============ 7 Bewertung ============
{
  const s = blank("V · Verarbeiten · Partnerpuzzle", "So wird dein Blatt bewertet", 3, "Vor dem Puzzle (1 Minute): Raster vorstellen, nicht diskutieren. Je Feld der Matrix ein Punkt (zwei Texte mal fünf Kriterien = 10 Punkte). Die Abgabe ist freiwillig; wer abgibt, erhält Punkte und Rückmeldung. Notenschlüssel: nach Ihrer Festlegung. Keine Selbsteinschätzung der Schülerinnen und Schüler. Ansage: „Stifte liegen. Ihr hört nur zu.“");
  [["0 bis 4 Punkte", "Ich kann einzelne Angaben zu einem Text nennen."], ["5 bis 8 Punkte", "Ich kann beide Texte mit den fünf Kriterien beschreiben."], ["9 bis 10 Punkte", "Ich kann beide Texte genau und vollständig beschreiben und vergleichen."]].forEach(([n, x], i) => {
    const y = 1.5 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.3, i % 2 ? LIGHT : LIGHT2);
    txt(s, n, { x: 0.8, y, w: 3.3, h: 1.3, fontFace: HEAD, fontSize: 27, bold: true, color: RED, valign: "middle" });
    txt(s, x, { x: 4.2, y, w: 8.4, h: 1.3, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Abgabe freiwillig: ", options: { bold: true } }, { text: "Jedes Feld zählt einen Punkt.", options: {} }]);
}

// ============ 8 Kriterien ============
{
  const s = blank("V · Verarbeiten · Partnerpuzzle", "Die fünf Kriterien", 3, "Die Folie steht während der Einzelarbeit und des Austauschs. Sie nennt die fünf Kriterien mit Leitfragen. Alle Felder der Matrix auf dem Blatt entsprechen diesen Kriterien. Ansage: „Stifte liegen. Ihr hört nur zu.“ Dann Folie 9.");
  [["1  Vorgeschichte", "Warum kam es dazu?"], ["2  Verlauf", "Was geschah? Gegen wen?"], ["3  Mittel", "Gewalt: ja oder nein?"], ["4  Opfer und Folgen", "Wer litt, was riskierte jemand?"], ["5  Wirkung", "Was änderte sich?"]].forEach(([n, x], i) => {
    const y = 1.5 + i * 0.9;
    box(s, 0.6, y, 12.15, 0.78, i % 2 ? LIGHT : LIGHT2);
    txt(s, n, { x: 0.8, y, w: 4.8, h: 0.78, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
    txt(s, x, { x: 5.7, y, w: 6.9, h: 0.78, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Fünf Kriterien, zwei Texte: ", options: { bold: true } }, { text: "zehn Punkte.", options: {} }]);
}

// ============ 9 Ü1 Einzelarbeit ============
{
  const { s } = taskSlide("I · Informieren · Übung 1", "Partnerpuzzle: Dein Text", ["Links sitzt Text A, rechts Text B.", "Lies deinen Text genau. Unterstreiche.", "Fülle deine Spalte in der Tabelle aus.", "Nutze die fünf Kriterien."], "10 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Ihr hört nur zu. Vor euch liegen ein Text und eine Tabelle. Wer links sitzt, hat Text A, wer rechts sitzt, Text B. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Die Texte (umgedreht) und die Tabellen liegen vorab auf den Tischen (Text A links, Text B rechts, pro Person eine Tabelle). 6 Paare, 12 Personen. Beide Tabellen sind gleich, jede Person füllt zuerst die eigene Spalte aus. Die Texte sind länger (fünf Absätze) und haben keine Hilfen: Die Schülerinnen und Schüler suchen selbst im Text, was zu den Kriterien passt, und unterstreichen. Sie gehen durch die Reihen (Heft-Stichprobe), nicht kommentieren. Wer früh fertig ist, markiert im Text, woran man erkennt, ob Gewalt im Spiel war. Nach 10 Minuten: Folie 10.");
  phaseTag(s, 2);
}

// ============ 10 Ü2 Austausch ============
{
  const { s } = taskSlide("V · Verarbeiten · Partnerpuzzle", "Austausch: Erkläre deinem Partner", ["Erkläre deinem Partner deinen Text.", "Er schreibt deine Spalte in sein Blatt.", "Wechselt: Jetzt erklärt dein Partner.", "Prüft zusammen: Sind alle zehn Felder voll?"], "8 Minuten · zu zweit · Flüsterstufe",
    "Ansage: „Stifte liegen. Flüsterstufe. Nur dein Partner hört dich. Erst erklärt die Person links ihren Text, dann die Person rechts. Wer zuhört, schreibt mit.“ Nach 4 Minuten: „Wechselt.“ Jede Person muss für ihren Text sprechen, niemand liest vor. Sie gehen durch die Reihen und hören bei 4 bis 6 Paaren kurz zu, ohne zu helfen. Nach 8 Minuten: Folie 11. Sollbruchstelle: Ist ein Paar nach 6 Minuten noch nicht beim Wechsel, kürzen Sie die zweite Runde.");
  phaseTag(s, 3);
}

// ============ 11 Kontrolle Matrix ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "So sieht es aus", 3, "Kontrolle (3 Minuten). Ansage: „Vergleicht und verbessert mit Grün.“ Die Folie zeigt Stichworte, die Punkte vergeben Sie nach dem Lösungsblatt (Datei Lösung_Partnerpuzzle). Gelenk: „Zwei Wege, ein Ziel: Der Dritte Stand will mitbestimmen. Aber was hat bei der Reihenfolge gewirkt? Seht auf die Zeitleiste.“ Nach der Kontrolle: Folie 12.");
  const cols = [["", 3.0, NAVY], ["Ballhausschwur", 4.55, "2E9E6B"], ["Bastille", 4.6, "D9534F"]];
  let x0 = 0.6; const xs = cols.map(([, w]) => { const v = x0; x0 += w; return v; });
  cols.forEach(([n, w, c], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y: 1.45, w, h: 0.75, fill: { color: c }, line: { color: WHITE, width: 1.5 } });
    txt(s, n, { x: xs[i], y: 1.45, w, h: 0.75, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  [["Vorgeschichte", "Saal gesperrt", "Hunger, Soldaten"], ["Verlauf", "Eid im Ballhaus", "Sturm auf Festung"], ["Mittel", "gewaltlos", "gewaltsam"], ["Opfer, Folgen", "keine Toten", "ca. 100 Tote"], ["Wirkung", "Versammlung bleibt", "Truppen abgezogen"]].forEach((row, r) => {
    row.forEach((v, i) => {
      const y = 2.2 + r * 0.78;
      s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y, w: cols[i][1], h: 0.78, fill: { color: r % 2 ? LIGHT : LIGHT2 }, line: { color: WHITE, width: 1.5 } });
      txt(s, v, { x: xs[i] + 0.1, y, w: cols[i][1] - 0.2, h: 0.78, fontSize: 27, bold: i === 0 || r === 2, color: r === 2 && i === 1 ? "2E9E6B" : r === 2 && i === 2 ? RED : NAVY, align: i === 0 ? "left" : "center", valign: "middle", fontFace: i === 0 ? HEAD : BODY });
    });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 12-15 Zeitstrahl-Auflösung ============
const AUF = [
  ["20. Juni 1789", "Ballhausschwur", "gewaltlos", "2E9E6B", "Gegen wen richtet sich der Eid?", "Aha-Moment, Schritt 1 (1 Minute). Ansage: „Stifte liegen. Ihr hört nur zu. Wir bringen die Ereignisse in die Reihenfolge.“ Zeigen Sie nur den ersten Eintrag. Frage an die Klasse (still denken, keine Antworten): „Gegen wen richtet sich der Eid?“ Erwartung: gegen den König und sein Verbot, die Abgeordneten handeln gegen seinen Willen (Ungehorsam). Danach ergänzend: „Wie reagiert der König?“ Dann Folie 13."],
  ["Anfang Juli 1789", "König zieht Soldaten zusammen", "Drohung", "E8A33D", "Was passiert in Paris?", "Schritt 2 (1 Minute). Der König zieht Soldaten um Paris und Versailles zusammen und entlässt den beliebten Finanzminister. Das war die Antwort auf den gewaltlosen Widerstand. Frage (still denken): „Wie reagieren die Menschen in Paris?“ Dann Folie 14."],
  ["14. Juli 1789", "Sturm auf die Bastille", "gewaltsam", "D9534F", "Wer hat die Gewalt angefangen?", "Schritt 3 (1 Minute). Ansage: „Schreibt ins Heft: Wer hat die Gewalt angefangen? Ein Satz.“ Antwort nicht im Plenum, nur Heft (Stichprobe bei 3 bis 4 Personen). Mögliche Antworten: Die Drohung des Königs (Soldaten) und die Angst der Menschen sind der Anlass; die Menge griff dann selbst zur Gewalt. Es geht um Reihenfolge und Anlass, nicht um Schuld. Dann Folie 15."],
  ["Mitte Juli 1789", "König zieht Soldaten ab", "Wirkung", "3F7CC4", "Was hat gewirkt?", "Schritt 4 (1 Minute). Ansage: „Seht euch die Reihenfolge an. Was hat den König zum Nachgeben gebracht?“ Antwort mündlich nicht nötig, 10 Sekunden denken. Aha: Der Druck wirkte erst durch beides zusammen: Der gewaltlose Eid schuf die Grundlage (Verfassung), die Gewalt in Paris zwang den König zum Rückzug. Wichtig für die Haltung: Wirksam heißt nicht gerechtfertigt (siehe Bilanz auf Folie 16). Dann Folie 16."],
];
AUF.forEach((_, k) => {
  const s = blank("V · Verarbeiten · Aha", "Die Reihenfolge der Ereignisse", 3, AUF[k][5]);
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 1.7, w: 0.1, h: 4.1, fill: { color: NAVY }, line: { type: "none" } });
  AUF.slice(0, k + 1).forEach(([d, e, tag, col], i) => {
    const y = 1.5 + i * 1.1;
    box(s, 0.6, y, 2.8, 0.85, i % 2 ? LIGHT : LIGHT2);
    txt(s, d, { x: 0.7, y, w: 2.6, h: 0.85, fontSize: 27, bold: true, color: RED, align: "center", valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: y + 0.26, w: 0.34, h: 0.34, fill: { color: i === k ? RED : NAVY }, line: { color: WHITE, width: 2 } });
    txt(s, e, { x: 4.15, y, w: 5.9, h: 0.85, fontFace: HEAD, fontSize: 27, bold: i === k, valign: "middle" });
    box(s, 10.2, y + 0.08, 2.55, 0.69, col);
    txt(s, tag, { x: 10.2, y: y + 0.08, w: 2.55, h: 0.69, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  bar(s, [{ text: "Frage: ", options: { bold: true } }, { text: AUF[k][4], options: {} }]);
});

// ============ 16 Bilanz ============
{
  const s = blank("V · Verarbeiten · Bilanz", "Was bewirkte es? Was kostete es?", 3, "Bilanz (1 Minute, Sie sprechen, kein Gespräch). Ansage: „Stifte liegen. Ihr hört nur zu.“ Wirksam heißt nicht gerechtfertigt: Beide Wege bewirkten etwas, aber sie kosteten Unterschiedliches. Haltung: Beschreiben, nicht gutheißen. Ihr Satz: „Wir erklären, warum es dazu kam. Ob es richtig war, beurteilt ihr gleich mit einem Maßstab.“ Bastille: Zahlen aus dem Lehrbuch gegenprüfen (etwa 100 Tote unter den Angreifern, der Gouverneur wurde nach der Übergabe getötet, sieben Gefangene). Der Eid brachte keine Toten, aber das Risiko, vom König aufgelöst zu werden. Dann Folie 17.");
  [["Ballhausschwur", "2E9E6B", "Versammlung bleibt. König gibt nach.", "Keine Toten. Risiko: Auflösung."], ["Bastille", "D9534F", "Truppen abgezogen. Symbol der Revolution.", "Ca. 100 Tote. Gouverneur getötet. Nur 7 Gefangene."]].forEach(([n, col, win, cost], i) => {
    const x = 0.6 + i * 6.2;
    box(s, x, 1.5, 5.95, 4.5, LIGHT2);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 5.95, h: 0.7, fill: { color: col }, line: { type: "none" } });
    txt(s, n, { x, y: 1.5, w: 5.95, h: 0.7, fontFace: HEAD, fontSize: 28, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, "Was bewirkte es?", { x: x + 0.25, y: 2.3, w: 5.5, h: 0.5, fontSize: 27, bold: true, color: NAVY });
    txt(s, win, { x: x + 0.25, y: 2.8, w: 5.5, h: 1.0, fontSize: 27 });
    txt(s, "Was kostete es?", { x: x + 0.25, y: 3.9, w: 5.5, h: 0.5, fontSize: 27, bold: true, color: RED });
    txt(s, cost, { x: x + 0.25, y: 4.4, w: 5.5, h: 1.5, fontSize: 27 });
  });
  bar(s, [{ text: "Wirksam ", options: { bold: true } }, { text: "heißt nicht gerechtfertigt.", options: {} }]);
}

// ============ 17 Ü3 Urteil ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 3", "Dein Urteil", ["Wähle einen Maßstab: wirksam, gerecht oder verantwortbar.", "Schreibe ins Heft: Ich halte … für …, weil …", "Nenne dabei einen Beleg und die Opfer.", "Vergleiche mit deiner Vermutung vom Anfang."], "8 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Hilfen (nur Satzanfänge, keine fertigen Sätze): „Wirksam, weil …“ „Gerecht, weil …“ „Verantwortbar, weil … (die Opfer: …)“ „Ich hatte vermutet …, jetzt sehe ich …“. Es gibt kein richtiges Urteil. Ein Sachurteil mit Beleg (zum Beispiel die Reihenfolge oder die Zahl der Toten), dann ein Werturteil mit offengelegtem Maßstab. Heft-Stichprobe: durch die Reihen gehen und bei 4 bis 6 Personen mitlesen, nicht kommentieren (unbenotet; das Urteil zählt nicht zu den 10 Punkten). Zusatz: Wer möchte, gibt das Blatt (Matrix) zur Bewertung und das Urteil freiwillig ab. Nach 8 Minuten: Folie 18.");
  phaseTag(s, 3);
}

// ============ 18 Abschlussfrage: Du bist der Dritte Stand ============
{
  const s = blank("A · Auswerten", "Abschlussfrage: Du bist der Dritte Stand", 4, "Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Jetzt seid ihr der Dritte Stand (Bauern und Bürger). Klerus und Adel haben gesagt: ‚Der Dritte Stand hat abgestimmt und verloren. Jetzt kann er sich nur noch fügen.‘ Denkt zehn Sekunden nach. Ich rufe euch auf.“ Die Antwort kommt mündlich, in der Ich-Form, ohne Schnipsel: „Ich antworte euch: …, weil …“ Die Klasse antwortet als Dritter Stand auf die Aussage vom Anfang. Mehrere gute Antworten, geprüft werden Position, ein Sachargument aus der Stunde (Eid, Bastille, Reihenfolge, Opfer) und ein „weil“. Erwartung (Auswahl, keine Musterlösung): „Wir fügen uns nicht, weil wir fast alle Menschen vertreten und uns zur Nationalversammlung erklärt haben.“ „Wir wehren uns, aber ohne Gewalt, weil die Bastille viele Tote gekostet hat.“ „Erst der Druck von beiden Seiten hat den König zum Rückzug gebracht, also fügen wir uns nicht.“ Form nach Verlauf: Blitzlicht (4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar) oder Duell zu zweit (leise: eine Person ist Klerus und Adel, eine antwortet als Dritter Stand). Für die 8C: Die Klasse kennt die Menschenrechte (13.11.) und darf sich auf Artikel 2 (Recht auf Widerstand) oder Artikel 3 (Macht geht vom Volk aus) berufen. Vergleich mit der Vermutung vom Anfang: Wer hat umgedacht? Abgabe-Hinweis: Wer das Blatt abgeben möchte, legt es jetzt vorn ab. Unbenotet.");
  s.addImage({ path: IMG + "karikatur_drei_staende.png", x: 0.6, y: 1.5, w: 5.4, h: 3.375 });
  box(s, 0.6, 5.0, 5.4, 0.75, WHITE, { line: NAVY, lw: 2 });
  txt(s, "Klerus · Adel · Dritter Stand", { x: 0.6, y: 5.0, w: 5.4, h: 0.75, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  bubble(s, 6.4, 1.5, 6.35, 2.2, WHITE, NAVY, "left", 2.6);
  txt(s, "Klerus und Adel: „Der Dritte Stand hat abgestimmt und verloren. Jetzt kann er sich nur noch fügen.“", { x: 6.7, y: 1.5, w: 5.85, h: 2.2, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  bubble(s, 6.4, 4.0, 6.35, 1.2, "FDF1D8", GOLD, "left", 4.6);
  txt(s, "Du, Dritter Stand:  „Ich antworte euch: …, weil …“", { x: 6.7, y: 4.0, w: 5.85, h: 1.2, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  [["Position", 1.8], ["Sachargument", 2.9], ["„weil“", 1.3]].reduce((xx, [lab, ww], i) => {
    box(s, xx, 5.4, ww, 0.62, i % 2 ? LIGHT : LIGHT2);
    s.addText([{ text: lab, options: { bold: true, fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: xx + 0.05, y: 5.4, w: ww - 0.1, h: 0.62, valign: "middle", align: "center", margin: 0 });
    return xx + ww + 0.15;
  }, 6.4);
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "3 Minuten · erst denken, dann mündlich      ", options: {} }, { text: "Alle denken mit.", options: { italic: true, color: ICE } }]);
}

// ============ 19 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Die Klasse prüft still: Kann ich das? Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen (30 Sekunden). Wenn Ziel 2 (*) nicht mehr geschafft wurde, kommt es in der Verfassungsstunde als Wiederholung. Danach Folie 20. Wer das Blatt abgeben möchte, legt es vorn ab.");
  [["Ich kann den Ballhausschwur und den Sturm auf die Bastille mit fünf Kriterien vergleichen."], ["Ich kann an der Reihenfolge der Ereignisse erklären, wie die Wege zusammenwirkten.*"], ["Ich kann beurteilen, welcher Weg berechtigt war, und das mit einem Maßstab begründen."]].forEach(([x], i) => {
    const y = 1.7 + i * 1.4;
    box(s, 0.6, y, 12.15, 1.15, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.15, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Blatt freiwillig abgeben.", options: {} }]);
}

// ============ 20 Ausblick ============
{
  const s = blank("A · Auswerten · Ausblick", "Wie geht es weiter?", 4, "Folie nur lesen (30 Sekunden). Die 8C kennt die Erklärung der Menschenrechte schon (13.11.). Ausblick: Artikel 2 nennt das Recht auf Widerstand gegen Unterdrückung, Artikel 3 die Macht des Volkes. Das Handeln der Abgeordneten und der Pariser stand also später als Recht in der Erklärung. Offen bleibt: Wie wird ein solches Recht im Staat gesichert? Das sehen wir in der nächsten Stunde bei der Verfassung von 1791 (Gewaltenteilung, Wahlrecht). Der zweite Satz zielt auf die Haltung: Gewalt war nicht das Ende der Revolution. Was passiert, wenn Gewalt zum normalen Mittel wird, sehen wir später bei den Jakobinern. Danach folgt der Test (4.12.).");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Die Menschenrechte geben dem Widerstand ein Recht. Wie wird dieses Recht gesichert?", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Nächste Stunde: ", options: { bold: true } }, { text: "Verfassung von 1791.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Stunde3_Widerstand_8C.pptx" }).then(() => console.log("ok"));
