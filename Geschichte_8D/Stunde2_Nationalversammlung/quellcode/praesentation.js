const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Stunde 2: Abstimmung und Nationalversammlung (8D)";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Nationalversammlung · Geschichte 8D";
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
  const s = blank("Ankommen", "Ein Gespräch in Versailles", 0, "Die Folie steht beim Betreten. Niemand muss etwas sagen. Die Karikatur ist selbst gezeichnet, keine historische Quelle. Der Adlige sagt: „Wenn der Dritte Stand unzufrieden ist, soll er einfach gegen die Politik des Königs stimmen.“ Klerus und Adel unterhalten sich. Die Klasse schreibt ihre Vermutung ins Heft (Eingangsdiagnose, bleibt stehen, wird in Phase 4 mit dem Ergebnis verglichen): „Stimmt das? Ich vermute …, weil …“ Keine Besprechung, kein Urteil. Die farbigen Standeskarten (Rolle und Stimmzettel in einem) liegen verdeckt auf den Plätzen. Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2 bis 5 in zusammen etwa 2 Minuten.");
  s.addImage({ path: IMG + "karikatur_klerus_adel.png", x: 0.6, y: 1.45, w: 7.0, h: 4.375 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.85, y: 1.55, w: 6.5, h: 1.55, rectRadius: 0.18, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: 5.5, y: 3.05, w: 0.6, h: 0.5, rotate: 180, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.52, y: 3.02, w: 0.56, h: 0.1, fill: { color: WHITE }, line: { type: "none" } });
  txt(s, "„Wenn der Dritte Stand unzufrieden ist, soll er einfach gegen die Politik des Königs stimmen.“", { x: 1.05, y: 1.55, w: 6.1, h: 1.55, fontFace: HEAD, fontSize: 27, color: NAVY, valign: "middle" });
  [["Klerus", 1.2], ["Adel", 4.7]].forEach(([n, x]) => {
    box(s, x, 5.2, n === "Klerus" ? 1.9 : 1.6, 0.52, WHITE, { line: NAVY, lw: 2 });
    txt(s, n, { x, y: 5.2, w: n === "Klerus" ? 1.9 : 1.6, h: 0.52, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  [["Heft raus. Datum und Überschrift.", 1.6, 1.0], ["Lies, was der Adlige sagt.", 2.85, 1.0], ["Schreibe: Stimmt das? Ich vermute …, weil …", 4.1, 1.7]].forEach(([x, y, h], i) => {
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
  const s = blank("Stundenfrage", "Das untersuchen wir heute", undefined, "Die Behauptung von Klerus und Adel wird zur Stundenfrage (30 Sekunden, Klasse hört nur zu): „Klerus und Adel sagen: Der Dritte Stand soll einfach abstimmen. Stimmt das? Das prüfen wir heute mit einer Abstimmung.“ Keine Antworten ins Plenum, die Vermutungen stehen im Heft.");
  box(s, 0.6, 1.6, 12.15, 2.6, NAVY);
  txt(s, "Kann der Dritte Stand (Bauern und Bürger) die Politik des Königs einfach wegstimmen?", { x: 1.0, y: 1.6, w: 11.35, h: 2.6, fontFace: HEAD, fontSize: 34, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.5, 12.15, 1.4, LIGHT2);
  txt(s, "Wir prüfen es mit einer Abstimmung der Stände.", { x: 1.0, y: 4.5, w: 11.35, h: 1.4, fontFace: HEAD, fontSize: 30, valign: "middle" });
}

// ============ 4 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Lernziele vorlesen (30 Sekunden, Klasse hört nur zu). Ziel 1 und 2 sind das Mindestziel. Ziel 3 hat einen Stern: Was nicht mehr geschafft wird, hat einen Stern. Am Ende (Folie 15) prüft die Klasse still, ob die Ziele erreicht sind.");
  [["1", "Ich kann erklären, wie die Stände den Dritten Stand (Bauern und Bürger) überstimmen können."], ["2", "Ich kann beurteilen, ob der Dritte Stand die Politik durch Abstimmen ändern kann, und das mit einer Zahl begründen."], ["3", "Ich kann am Zeitstrahl erklären, wie sich der Dritte Stand wehrt.*"]].forEach(([n, x], i) => {
    const y = 1.75 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.37, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Karikatur, Abstimmung, Text, Zeitstrahl.", options: {} }]);
}

// ============ 5 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Ablauf zeigen (30 Sekunden). Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Ziel 3). Insgesamt 25 Minuten.");
  const ph = [["A", "Ankommen", "Karikatur", "4 Min."], ["V", "Vorwissen", "Abstimmung", "4 Min."], ["I", "Informieren", "Text", "4 Min."], ["V", "Verarbeiten", "Zeitstrahl", "3 Min."], ["A", "Auswerten", "Ergebnis, Rat", "10 Min."]];
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

// ============ 6 Zwei Anträge ============
{
  const s = blank("V · Vorwissen aktivieren · Abstimmung", "Zwei Anträge, eine Regel", 1, "Gelenk: „Klerus und Adel sagen: Einfach abstimmen. Das testen wir jetzt.“ Die beiden Anträge sind die Politik, über die der Dritte Stand abstimmen möchte. Ansage: „Stifte liegen. Ihr hört nur zu.“ Anträge und Regel vorlesen (1 Minute). Die beiden Anträge sind Spielanträge, historisch ging der Streit um die Abstimmungsregel. Klerus und Adel haben in beiden Fällen dieselben Interessen: Sie wollen ihre Vorrechte behalten. Klerus 3, Adel 4, Dritter Stand (Bauern und Bürger) 18 Personen (bei 25 Personen, ohne Wahlhelfer). Die Klasse stimmt vorab auf der Standeskarte ab (Stimmzettel unten auf der Karte), die Karten werden eingesammelt. Sie werten sie während Text und Zeitstrahl aus und schreiben die Ergebnisse in die Matrix an der Tafel. Regel zum Vorlesen: Jeder Stand hat eine Stimme (Abstimmung nach Ständen). Gezählt werden nur die Ja-Stimmen je Stand. Das Ergebnis nach Ständen und nach Köpfen rechnet die Klasse später selbst aus (Folie 11).");
  [["A", "Auch Adel und Klerus zahlen künftig Steuern."], ["B", "Die Stände werden abgeschafft. Alle haben gleiche Rechte."]].forEach(([n, x], i) => {
    const y = 1.5 + i * 1.15;
    box(s, 0.6, y, 12.15, 0.95, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.85, y + 0.25, 0.46, RED);
    txt(s, x, { x: 1.55, y, w: 11.0, h: 0.95, fontFace: HEAD, fontSize: 27, bold: true, valign: "middle" });
  });
  txt(s, "Klerus 3 · Adel 4 · Dritter Stand (Bauern und Bürger) 18 Personen", { x: 0.7, y: 4.0, w: 12.0, h: 0.65, fontSize: 27, color: NAVY, valign: "middle" });
  txt(s, "Ihr stimmt auf eurer Karte ab. Ich sammle die Karten ein und werte sie aus.", { x: 0.7, y: 4.75, w: 12.0, h: 0.9, fontSize: 27, color: NAVY, valign: "middle" });
  bar(s, [{ text: "Regel: ", options: { bold: true } }, { text: "Jeder Stand hat eine Stimme (nach Ständen).", options: {} }]);
}

// ============ 7 Ü1 Vorab ============
{
  const { s } = taskSlide("V · Vorwissen aktivieren · Übung 1", "Wir stimmen ab: Deine Karte", ["Drehe deine Karte um. Lies sie leise.", "Kreuze unten deine Stimme bei Antrag A an.", "Kreuze unten deine Stimme bei Antrag B an.", "Halte die Karte bereit. Ich sammle sie ein."], "3 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Klerus und Adel stimmen nach ihrer Rolle (Vorrechte behalten), der Dritte Stand (Bauern und Bürger) nach seiner Überzeugung. Keine Absprache, keine Bewegung. Nach etwa zwei Minuten: „Stifte hinlegen. Ich sammle die Karten ein.“ Sie gehen durch die Reihen und sammeln die Zettel ein (etwa 1 Minute, ruhig). Danach sofort Folie 8. Sie werten die Karten während Text und Zeitstrahl aus: Karten nach Farbe (Stand) sortieren, Ja und Nein je Antrag in der Strichliste (Tafelmatrix-Datei, Seite 3), Ja-Zahlen in die Matrix an der Tafel schreiben. Notfall: Wer keine Karte hat, nimmt eine Reservekarte (Dritter Stand, blau).");
  phaseTag(s, 1);
}

// ============ 8 Ü2 Text ============
{
  const { s } = taskSlide("I · Informieren · Übung 2", "Der Dritte Stand wehrt sich: Text", ["Lies den Text auf dem Blatt.", "Unterstreiche alle Daten.", "Lies die Wörter unter dem Text."], "4 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Ihr hört nur zu. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Das Blatt (Arbeitsblatt_Der_Dritte_Stand_wehrt_sich.pdf, ein Blatt pro Person) liegt vorab auf den Tischen, umgedreht, oder wird jetzt ausgeteilt. Der Text ist gegenüber der 8C gekürzt (ein Drittel einer Seite, vier Daten). Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung. Historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Während die Klasse liest, sortieren und zählen Sie die Karten (Strichliste, Tafelmatrix-Datei Seite 3) und tragen die Ja-Zahlen in die Matrix an der Tafel ein. Die Heft-Stichprobe entfällt in dieser Phase. Wer schneller fertig ist, liest den Text noch einmal und überlegt, wie der Dritte Stand die Macht des Königs infrage stellt. Nach 4 Minuten: Folie 9.");
  phaseTag(s, 2);
}

// ============ 9 Ü3 Zeitstrahl ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 3", "Zeitstrahl: Was geschah wann?", ["Schreibe zu jedem Datum das Ereignis in den Zeitstrahl.", "Nutze Stichworte aus dem Text.", "Prüfe: Passt jedes Ereignis zum Datum?"], "2 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Sie schreiben die Ja-Zahlen an die Tafel (Strichliste, Matrix), während die Klasse arbeitet. Der Zeitstrahl steht unten auf demselben Blatt, die Daten sind vorgegeben. 2 Minuten, dann Folie 10 (Lösung, 1 Minute). Aufgabe 3 auf dem Blatt (Sieyès und die Zahl) entfällt oder ist Zusatz (*). Wenn die Zeit knapp ist, genügen die drei letzten Einträge (17. Juni, 20. Juni, Anfang Juli).");
  phaseTag(s, 3);
}

// ============ 10 Zeitstrahl ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "Der Dritte Stand wehrt sich: Zeitstrahl", 3, "Nach dem Ausfüllen zeigen (1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ Der letzte Eintrag („Anfang Juli: Soldaten um Paris“) ist das Ziel der Stunde. Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung. Historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Quelle Sieyès: verfasst 1788, erschienen Winter 1788/89.");
  const ev = [["5. Mai 1789", "Eröffnung, Streit um die Abstimmung"], ["17. Juni 1789", "Nationalversammlung"], ["20. Juni 1789", "Ballhausschwur"], ["Anfang Juli 1789", "Soldaten um Paris, Angst in Paris"]];
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 1.7, w: 0.1, h: 4.35, fill: { color: NAVY }, line: { type: "none" } });
  ev.forEach(([d, e], i) => {
    const y = 1.55 + i * 1.1;
    box(s, 0.6, y, 2.8, 0.75, i % 2 ? LIGHT : LIGHT2);
    txt(s, d, { x: 0.7, y, w: 2.6, h: 0.75, fontSize: 27, bold: true, color: RED, align: "center", valign: "middle" });
    s.addShape(pres.shapes.OVAL, { x: 3.5, y: y + 0.2, w: 0.34, h: 0.34, fill: { color: i === 3 ? RED : NAVY }, line: { color: WHITE, width: 2 } });
    txt(s, e, { x: 4.2, y, w: 8.5, h: 0.75, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
}


// ============ 11 Ü4 Rechne und werte aus ============
{
  const { s } = taskSlide("A · Auswerten · Übung 4", "Rechne und werte aus", ["Lies die Ja-Zahlen an der Tafel ab. Rechne: nach Ständen und nach Köpfen.", "Schreibe: Nach Ständen …, nach Köpfen … Das zeigt: … (nenne Zahlen).", "Vergleiche mit deiner Vermutung vom Anfang."], "4 Minuten · allein · leise",
    "Gelenk: „Ihr habt abgestimmt, Klerus, Adel und Dritter Stand (Bauern und Bürger). Jetzt seht ihr, was herausgekommen ist.“ Ansage: „Stifte liegen. Ihr hört nur zu. Hier ist das Ergebnis.“ Die Matrix steht an der Tafel (Ja-Zahlen je Stand für A und B). Dann: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Zuerst rechnen (etwa 1 Minute): Nach Ständen: Wie viele der drei Stände sagen Ja? Mindestens 2 von 3 = angenommen. Nach Köpfen: Alle Ja zusammenzählen (von allen 25), mehr als die Hälfte = angenommen. Dann ins Heft (etwa 3 Minuten): zuerst die beiden Ergebnisse, dann was die Abstimmung zeigt (Antwort auf die Stundenfrage „Kann der Dritte Stand die Politik einfach wegstimmen?“, mit Zahlen) und der Vergleich mit der Vermutung vom Anfang. Die Antwort an den Adligen kommt danach mündlich (Folie 14), vorher zeigt Folie 13, dass es 1789 keine Abstimmung nach Köpfen gab. Wenn die Zeit knapp ist, genügen die Ergebnisse und ein Urteilssatz. Heft-Stichprobe: durch die Reihen gehen, bei 4 bis 6 Personen mitlesen, nicht kommentieren (unbenotet). Erwartung (bei A und B gleich): Nach Ständen 1 von 3 Ständen Ja, abgelehnt. Nach Köpfen 18 von 25 Ja, angenommen (echte Zahlen gelten). Dieselben Stimmen, ein anderes Ergebnis. Warum: Klerus und Adel stimmten gleich, jeder Stand hat nur eine Stimme. Antwort mit Zahl, z. B. „Nein, das stimmt nicht, weil Klerus und Adel zusammen 2 von 3 Stimmen haben, obwohl sie nur etwa 3 % der Menschen sind“ oder „Nur nach Köpfen hätte der Dritte Stand (Bauern und Bürger) gewonnen: 18 von 25“. Die Antwort wird nicht vorgegeben. Einordnung (ein Satz): Im Spiel stimmen Klerus und Adel immer gleich. Historisch waren die Stände nicht immer einheitlich (viele einfache Pfarrer unterstützten den Dritten Stand), bei der Abstimmungsregel hielten die Privilegierten aber oft zusammen. In Frankreich sind Klerus und Adel zusammen ca. 3 % der Menschen (ca. 150 000 und ca. 500 000 von ca. 20,65 Millionen).");
  phaseTag(s, 4);
}

// ============ 12 Matrix ausgefüllt ============
{
  const s = blank("A · Auswerten · Kontrolle", "Dieselben Stimmen, zwei Regeln", 4, "Kontrolle (1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ Die Zahlen aus der Klasse gelten, die Folie zeigt die erwartete Matrix (Klerus 3, Adel 4, Dritter Stand 18). Gelenk: „Klerus und Adel sagen, der Dritte Stand soll einfach abstimmen. Nach Ständen verliert er immer, obwohl er fast alle Menschen vertritt. Was rät man ihm jetzt?“ Vereinfachung im Spiel: Klerus und Adel stimmen immer gleich; die Anträge sind Spielanträge.");
  const cols = [["Antrag", 1.25, NAVY], ["Klerus", 1.8, KLERUS], ["Adel", 1.8, ADEL], ["Dritter Stand", 2.1, DRITTER], ["Ergebnis nach Ständen", 2.6, NAVY], ["Ergebnis nach Köpfen", 2.6, NAVY]];
  let x0 = 0.6;
  const xs = cols.map(([, w]) => { const v = x0; x0 += w; return v; });
  cols.forEach(([n, w, c], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y: 1.5, w, h: 0.95, fill: { color: c }, line: { color: WHITE, width: 1.5 } });
    txt(s, n, { x: xs[i], y: 1.5, w, h: 0.95, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  [["A", "Ja 0\nvon 3", "Ja 0\nvon 4", "Ja 18\nvon 18", "1 von 3 Ja: abgelehnt", "18 von 25 Ja: angenommen"], ["B", "Ja 0\nvon 3", "Ja 0\nvon 4", "Ja 18\nvon 18", "1 von 3 Ja: abgelehnt", "18 von 25 Ja: angenommen"]].forEach((row, r) => {
    row.forEach((v, i) => {
      const y = 2.45 + r * 1.55;
      s.addShape(pres.shapes.RECTANGLE, { x: xs[i], y, w: cols[i][1], h: 1.55, fill: { color: r % 2 ? LIGHT : LIGHT2 }, line: { color: WHITE, width: 1.5 } });
      txt(s, v, { x: xs[i] + 0.05, y, w: cols[i][1] - 0.1, h: 1.55, fontSize: i === 0 ? 30 : 27, bold: i === 0 || i > 3, color: i > 3 ? RED : NAVY, align: "center", valign: "middle" });
    });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
}

// ============ 13 Keine Abstimmung nach Köpfen ============
{
  const s = blank("A · Auswerten · Geschichte", "1789: Wie wurde wirklich abgestimmt?", 4, "Dauer: etwa 20 Sekunden, Sie sprechen, kein Gespräch. Ansage: „Stifte liegen. Ihr hört nur zu.“ Wichtig für die Einordnung: Im Spiel haben wir zwei Regeln verglichen. 1789 gab es keine Abstimmung nach Köpfen. Der Dritte Stand (Bauern und Bürger) forderte sie, der König lehnte ab, bei den Generalständen wurde nach Ständen abgestimmt. Der Streit um diese Regel blockierte die Generalstände wochenlang (siehe Text). Nach Köpfen gerechnet hätte der Dritte Stand Mehrheiten gehabt, deshalb war die Regel so wichtig. Stichwort für die nächste Stunde und das Ende der Folie: Erst als sich der Dritte Stand am 17. Juni zur Nationalversammlung erklärte, zählten die Abgeordneten nach Köpfen. Vergleich: Die Folie zeigt, was in unserem Spiel nur zur Probe berechnet wurde.");
  box(s, 0.6, 1.5, 5.95, 4.5, LIGHT2);
  txt(s, "Nach Ständen", { x: 0.9, y: 1.6, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY });
  txt(s, "So wurde 1789 abgestimmt.", { x: 0.9, y: 2.3, w: 5.4, h: 0.6, fontSize: 27 });
  [["Klerus", KLERUS], ["Adel", ADEL], ["Dritter Stand", DRITTER]].forEach(([n, c], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.9 + i * 1.8, y: 3.1, w: 1.65, h: 1.3, fill: { color: c }, line: { type: "none" } });
    txt(s, n + "\n1 Stimme", { x: 0.9 + i * 1.8, y: 3.1, w: 1.65, h: 1.3, fontSize: 27, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  txt(s, "Jeder Stand hat eine Stimme.", { x: 0.9, y: 4.75, w: 5.4, h: 0.9, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY });
  // durchgestrichen
  box(s, 6.8, 1.5, 5.95, 4.5, "FBE9E8", { line: RED, lw: 3 });
  s.addText("Nach Köpfen", { isTextBox: true, x: 7.1, y: 1.6, w: 5.4, h: 0.55, fontFace: HEAD, fontSize: 28, bold: true, color: RED, strike: "sngStrike", margin: 0 });
  s.addText("Jede Hand zählt einzeln.", { isTextBox: true, x: 7.1, y: 2.3, w: 5.4, h: 0.6, fontFace: BODY, fontSize: 27, color: RED, strike: "sngStrike", margin: 0 });
  stackBar(s, 7.1, 3.1, 5.35, 1.3, [[18, "E3A3A0", "Ja"], [7, "E3A3A0", "Nein"]]);
  s.addText("Gab es 1789 nicht.", { isTextBox: true, x: 7.1, y: 4.75, w: 5.4, h: 0.9, fontFace: HEAD, fontSize: 28, bold: true, color: RED, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 7.1, y: 3.1, w: 5.35, h: 1.3, line: { color: RED, width: 10 } });
  s.addShape(pres.shapes.LINE, { x: 12.45, y: 3.1, w: -5.35, h: 1.3, line: { color: RED, width: 10 } });
  bar(s, [{ text: "Wichtig: ", options: { bold: true } }, { text: "1789 gab es keine Abstimmung nach Köpfen.", options: {} }]);
}

// ============ 14 Abschlussfrage: Antwort an den Adligen ============
{
  const s = blank("A · Auswerten", "Abschlussfrage: Was antwortest du ihm?", 4, "Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Denkt zehn Sekunden nach. Ich rufe euch auf.“ Die Klasse nimmt Bezug auf die Aussage des Adligen vom Anfang: „Wenn der Dritte Stand unzufrieden ist, soll er einfach gegen die Politik des Königs stimmen.“ Die Antwort kommt mündlich, ohne Schnipsel: „Ich antworte dir, …, weil …“. Mehrere gute Antworten, geprüft werden Position, ein Sachargument aus der Stunde (Zahl aus der Matrix, 97 % der Menschen, 1 von 3 Stimmen, Nationalversammlung) und ein „weil“. Erwartung (Auswahl, keine Musterlösung): „Das funktioniert nicht, weil nach Ständen Klerus und Adel 2 von 3 Stimmen haben, obwohl der Dritte Stand (Bauern und Bürger) fast alle Menschen vertritt.“ „Das stimmt nur, wenn nach Köpfen abgestimmt wird, dort hätte der Dritte Stand 18 von 25 Stimmen.“ „Ihr habt Angst vor der Mehrheit, deshalb wollt ihr nicht nach Köpfen abstimmen.“ „Wenn Abstimmen nichts bringt, wehrt sich der Dritte Stand anders, er erklärt sich zur Nationalversammlung.“ Form nach Verlauf: Blitzlicht (4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar) oder Duell zu zweit (leise: eine Person ist der Adlige, eine antwortet). Unbenotet. Vergleich mit der Vermutung vom Anfang: Wer hat umgedacht?");
  s.addImage({ path: IMG + "karikatur_klerus_adel.png", x: 0.6, y: 1.5, w: 5.4, h: 3.375 });
  box(s, 0.6, 5.0, 5.4, 0.75, WHITE, { line: NAVY, lw: 2 });
  txt(s, "Klerus und Adel", { x: 0.6, y: 5.0, w: 5.4, h: 0.75, fontSize: 27, bold: true, color: NAVY, align: "center", valign: "middle" });
  bubble(s, 6.4, 1.5, 6.35, 2.2, WHITE, NAVY, "left", 2.6);
  txt(s, "„Wenn der Dritte Stand unzufrieden ist, soll er einfach gegen die Politik des Königs stimmen.“", { x: 6.7, y: 1.5, w: 5.85, h: 2.2, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  bubble(s, 6.4, 4.0, 6.35, 1.2, "FDF1D8", GOLD, "left", 4.6);
  txt(s, "Du antwortest:  „Ich antworte dir, …, weil …“", { x: 6.7, y: 4.0, w: 5.85, h: 1.2, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, valign: "middle" });
  [["Position", 1.8], ["Sachargument", 2.9], ["„weil“", 1.3]].reduce((xx, [lab, ww], i) => {
    box(s, xx, 5.4, ww, 0.62, i % 2 ? LIGHT : LIGHT2);
    s.addText([{ text: lab, options: { bold: true, fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: xx + 0.05, y: 5.4, w: ww - 0.1, h: 0.62, valign: "middle", align: "center", margin: 0 });
    return xx + ww + 0.15;
  }, 6.4);
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "3 Minuten · erst denken, dann mündlich      ", options: {} }, { text: "Alle denken mit.", options: { italic: true, color: ICE } }]);
}

// ============ 15 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Die Klasse prüft still: Kann ich das? Danach Folie 16. Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen (30 Sekunden). Wenn Ziel 3 (*) nicht mehr geschafft wurde, kommt es als Wiederholung in Stunde 3.");
  [["Ich kann erklären, wie die Stände den Dritten Stand (Bauern und Bürger) überstimmen können."], ["Ich kann beurteilen, ob der Dritte Stand die Politik durch Abstimmen ändern kann, und das mit einer Zahl begründen."], ["Ich kann am Zeitstrahl erklären, wie sich der Dritte Stand wehrt.*"]].forEach(([x], i) => {
    const y = 1.7 + i * 1.4;
    box(s, 0.6, y, 12.15, 1.15, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.15, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Haben wir sie erreicht?", options: {} }]);
}

// ============ 16 Ausblick ============
{
  const s = blank("A · Auswerten · Ausblick", "Wie geht es weiter?", 4, "Folie nur lesen (30 Sekunden). Ausblick: Anfang Juli zieht der König Soldaten um Paris zusammen und entlässt seinen Finanzminister. In Paris wächst die Angst. Das sehen wir in der nächsten Stunde (Bastille, Ende der Vorrechte, Menschenrechte).");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Der König zieht Soldaten zusammen.\nIn Paris wächst die Angst.", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Wie reagieren die Menschen in Paris? Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Nächste Stunde: ", options: { bold: true } }, { text: "Bastille, Ende der Vorrechte, Menschenrechte.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Stunde2_Nationalversammlung_8D.pptx" }).then(() => console.log("ok"));
