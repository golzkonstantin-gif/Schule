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


// ============ 1 Titel / Ankommen ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "Nationalversammlung 1789", { x: 0.9, y: 1.2, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  txt(s, "Wie wehrt sich der Dritte Stand gegen die Abstimmungsregel?", { x: 0.9, y: 2.3, w: 11.5, h: 0.6, fontSize: 27, italic: true, color: ICE });
  box(s, 0.9, 3.6, 11.5, 2.6, NAVY2);
  txt(s, "Ankommen", { x: 1.2, y: 3.75, w: 6, h: 0.5, fontFace: HEAD, fontSize: 27, bold: true, color: WHITE });
  ["Hole dein Heft heraus.", "Schreibe Datum und Überschrift ins Heft.", "Lass deine Karte verdeckt liegen."].forEach((x, i) => {
    numCircle(s, i + 1, 1.2, 4.45 + i * 0.52, 0.4, RED);
    txt(s, x, { x: 1.85, y: 4.42 + i * 0.52, w: 10.3, h: 0.46, fontSize: 27, color: WHITE, valign: "middle" });
  });
  txt(s, "Geschichte, Klasse 8D · Französische Revolution · Stunde 2", { x: 0.9, y: 6.7, w: 8, h: 0.3, fontSize: 13, color: ICE });
  s.addNotes("Die Folie steht beim Betreten. Niemand muss etwas sagen. Auf jedem Platz liegen eine Standeskarte (verdeckt) und ein Beobachtungszettel. An der Tafel steht die Matrix für die Abstimmung (siehe Tafelskript). Ansage: „Stifte liegen. Ihr hört nur zu.“ Danach Folie 2, 3 und 4 in zusammen etwa 2 Minuten.");
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


// ============ 3 Lernziele ============
{
  const s = blank("Lernziele", "Das kannst du am Ende der Stunde", undefined, "Lernziele vorlesen (30 Sekunden, Klasse hört nur zu). Ziel 1 und 2 sind das Mindestziel. Ziel 3 hat einen Stern: Was nicht mehr geschafft wird, hat einen Stern. Am Ende (Folie 14) prüft die Klasse still, ob die Ziele erreicht sind.");
  [["1", "Ich kann erklären, wie die Stände den Dritten Stand (Bauern und Bürger) überstimmen können."], ["2", "Ich kann die Abstimmung nach Ständen beurteilen und mein Urteil begründen."], ["3", "Ich kann am Zeitstrahl erklären, wie sich der Dritte Stand zur Nationalversammlung erklärte.*"]].forEach(([n, x], i) => {
    const y = 1.75 + i * 1.5;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.9, y + 0.37, 0.5, RED);
    txt(s, x, { x: 1.7, y, w: 10.8, h: 1.25, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Abstimmung beobachten, Text, Zeitstrahl.", options: {} }]);
}

// ============ 4 Ablauf ============
{
  const s = blank("Ablauf", "So arbeiten wir heute", undefined, "Ablauf zeigen (30 Sekunden). Die Kreise oben rechts zeigen die Phase: A V I V A. Der Stern * bedeutet: wird eventuell nicht mehr geschafft (Ziel 3). Insgesamt 25 Minuten.");
  const ph = [["A", "Ankommen", "Ziele", "3 Min."], ["V", "Vorwissen", "Abstimmung", "10 Min."], ["I", "Informieren", "Text", "5 Min."], ["V", "Verarbeiten", "Zeitstrahl", "3 Min."], ["A", "Auswerten", "Rat an Freund", "4 Min."]];
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

// ============ 5 Zwei Anträge ============
{
  const s = blank("V · Vorwissen aktivieren · Abstimmung", "Zwei Anträge, eine Regel", 1, "Gelenk: „Auf dem Zeitstrahl endete der Streit um die Abstimmung. Heute erlebt ihr diesen Streit selbst.“ Ansage: „Stifte liegen. Ihr hört nur zu.“ Anträge und Regel vorlesen (1 Minute). Die beiden Anträge sind Spielanträge, historisch ging der Streit um die Abstimmungsregel. Klerus und Adel haben in beiden Fällen dieselben Interessen: Sie wollen ihre Vorrechte behalten. Klerus 3, Adel 4, Dritter Stand (Bauern und Bürger) 18 Personen (bei 25 Personen, ohne Wahlhelfer). Sie zählen und füllen die Matrix an der Tafel aus, die Klasse schreibt auf dem Zettel mit. Regel zum Vorlesen: Jeder Stand hat eine Stimme (Abstimmung nach Ständen). Gezählt werden nur die Ja-Stimmen je Stand. Das Ergebnis nach Ständen und nach Köpfen wird danach ausgerechnet.");
  [["A", "Auch Adel und Klerus zahlen künftig Steuern."], ["B", "Die Stände werden abgeschafft. Alle haben gleiche Rechte."]].forEach(([n, x], i) => {
    const y = 1.5 + i * 1.15;
    box(s, 0.6, y, 12.15, 0.95, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, n, 0.85, y + 0.25, 0.46, RED);
    txt(s, x, { x: 1.55, y, w: 11.0, h: 0.95, fontFace: HEAD, fontSize: 27, bold: true, valign: "middle" });
  });
  txt(s, "Klerus 3 · Adel 4 · Dritter Stand (Bauern und Bürger) 18 Personen", { x: 0.7, y: 4.0, w: 12.0, h: 0.65, fontSize: 27, color: NAVY, valign: "middle" });
  txt(s, "Ich zähle die Hände und schreibe an die Tafel. Ihr schreibt mit.", { x: 0.7, y: 4.75, w: 12.0, h: 0.9, fontSize: 27, color: NAVY, valign: "middle" });
  bar(s, [{ text: "Regel: ", options: { bold: true } }, { text: "Jeder Stand hat eine Stimme (nach Ständen).", options: {} }]);
}

// ============ 6 Ü1 Vorab ============
{
  const { s } = taskSlide("V · Vorwissen aktivieren · Übung 1", "Vorab: Deine Stimme", ["Drehe deine Karte um. Lies sie leise.", "Kreuze auf dem Zettel deinen Stand an.", "Kreuze deine Stimme bei A und B an.", "Zeige niemandem deine Karte."], "1 Minute · allein · leise",
    "Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Klerus und Adel stimmen nach ihrer Rolle (Vorrechte behalten), der Dritte Stand (Bauern und Bürger) nach seiner Überzeugung. Keine Absprache, keine Bewegung. Nach etwa einer Minute: „Stifte hinlegen. Wir stimmen jetzt ab.“ Notfall: Wer keine Karte hat, nimmt eine Reservekarte (Dritter Stand).");
  phaseTag(s, 1);
}

// ============ 7 Ü2 Abstimmung ============
{
  const { s } = taskSlide("V · Vorwissen aktivieren · Übung 2", "Wir stimmen ab: Du schreibst mit", ["Ich rufe einen Stand und einen Antrag auf.", "Wer dafür ist, hebt die Hand.", "Ich zähle und schreibe die Ja-Zahl an die Tafel.", "Trage die Ja-Zahl in deinen Zettel ein."], "4 Minuten · alle beobachten · leise",
    "Ansage: „Stifte liegen. Jetzt stimmt ihr offen ab. Ich zähle und schreibe an die Tafel. Ihr beobachtet und schreibt auf eurem Zettel mit.“ Ablauf (Zeile für Zeile der Matrix): „Antrag A. Klerus: Wer ist dafür? Hand hoch.“ Hände zählen, die Zahl laut sagen und an die Tafel schreiben („Klerus: Ja 0 von 3“), dann Adel und Dritter Stand, danach Antrag B. Nur die Ja-Stimmen werden gezählt; die Ergebnisse folgen danach (Folie 8). Erwartet: Klerus Ja 0 von 3, Adel Ja 0 von 4, Dritter Stand Ja 18 von 18, bei A und B (echte Zahlen gelten). Sechs Zählungen in etwa 4 Minuten; Notfallplan: Wird es unruhig, zählen Sie nach Stand ab und schreiben die Zahlen ohne Nachfrage an.");
  phaseTag(s, 1);
}

// ============ 8 Ü3 Rechnen und auswerten ============
{
  const { s } = taskSlide("V · Vorwissen aktivieren · Übung 3", "Rechne und werte aus", ["Rechne in der Matrix: Stände mit Ja und alle Ja zusammen.", "Schreibe das Ergebnis dazu: angenommen oder abgelehnt.", "Schreibe ins Heft: Nach Ständen …, nach Köpfen … Beurteile die Regel „Jeder Stand hat eine Stimme“: Ich finde sie …, weil … (nenne eine Zahl)."], "3 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. Rechnet erst in der Matrix, dann schreibt ihr ins Heft.“ Zuerst rechnen (etwa 1 Minute): Nach Ständen: Wie viele der drei Stände sagen Ja? Mindestens 2 von 3 = angenommen. Nach Köpfen: Alle Ja zusammenzählen (von allen 25), mehr als die Hälfte = angenommen. Dann ins Heft (etwa 2 Minuten): zuerst die Ergebnisse, dann das Urteil. Wenn die Zeit knapp ist, genügen die Ergebnisse und ein Urteilssatz. Erwartung (bei A und B gleich): Nach Ständen 1 von 3 Ständen Ja, abgelehnt. Nach Köpfen 18 von 25 Ja, angenommen (echte Zahlen gelten). Dieselben Stimmen, ein anderes Ergebnis. Warum: Klerus und Adel stimmten gleich, jeder Stand hat nur eine Stimme. Urteil: offen mit Zahl (mehrere gute Antworten), z. B. „Ich finde die Regel ungerecht, weil der Dritte Stand (Bauern und Bürger) ca. 97 % der Menschen stellt, aber nur 1 von 3 Stimmen hat“ oder „Ich finde sie nachvollziehbar, weil jeder Stand als Gruppe zählt, aber die Zahl 97 % zeigt ein Problem“. Das Urteil wird nicht vorgegeben. Einordnung (ein Satz): Im Spiel stimmen Klerus und Adel immer gleich. Historisch waren die Stände nicht immer einheitlich (viele einfache Pfarrer unterstützten den Dritten Stand), bei der Abstimmungsregel hielten die Privilegierten aber oft zusammen. In Frankreich sind Klerus und Adel zusammen ca. 3 % der Menschen (ca. 150 000 und ca. 500 000 von ca. 20,65 Millionen). Gezielte Abfrage nur bei Bedarf (2 bis 3 Personen, kurz): „Wie haben Klerus und Adel abgestimmt? Warum?“");
  phaseTag(s, 1);
}

// ============ 9 Matrix ausgefüllt ============
{
  const s = blank("V · Vorwissen aktivieren · Kontrolle", "Dieselben Stimmen, zwei Regeln", 1, "Kontrolle (1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ Die Zahlen aus der Klasse gelten, die Folie zeigt die erwartete Matrix (Klerus 3, Adel 4, Dritter Stand 18). Gelenk zum Text: „Der Dritte Stand hat gegen die Regel verloren, aber er gibt nicht auf. Wie reagiert er? Das lest ihr jetzt.“ Vereinfachung im Spiel: Klerus und Adel stimmen immer gleich; die Anträge sind Spielanträge.");
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

// ============ 10 Ü4 Text ============
{
  const { s } = taskSlide("I · Informieren · Übung 4", "Der Dritte Stand handelt: Text", ["Lies den Text auf dem Blatt.", "Unterstreiche alle Daten.", "Lies die Quelle von Sieyès: Was verlangt der Dritte Stand?"], "5 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Ihr hört nur zu. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Das Blatt (Arbeitsblatt_Der_Dritte_Stand_handelt.pdf, ein Blatt pro Person) liegt vorab auf den Tischen, umgedreht, oder wird jetzt ausgeteilt. Text unverändert wie in der 8C. Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung. Historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Heft-Stichprobe: durch die Reihen gehen und bei 4 bis 6 Personen unterschiedlichen Niveaus mitlesen (unterstrichene Daten), nicht kommentieren (unbenotet). Wer schneller fertig ist, liest den Text noch einmal und überlegt, wie der Dritte Stand die Macht des Königs infrage stellt. Nach 5 Minuten: Folie 11.");
  phaseTag(s, 2);
}

// ============ 11 Ü5 Zeitstrahl ============
{
  const { s } = taskSlide("V · Verarbeiten · Übung 5", "Zeitstrahl: Was geschah wann?", ["Schreibe zu jedem Datum das Ereignis in den Zeitstrahl.", "Nutze Stichworte aus dem Text.", "Prüfe: Passt jedes Ereignis zum Datum?"], "2 Minuten · allein · leise",
    "Ansage: „Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Der Zeitstrahl steht unten auf demselben Blatt, die Daten sind vorgegeben. 2 Minuten, dann Folie 12 (Lösung, 1 Minute). Aufgabe 3 auf dem Blatt (Sieyès und die Zahl) entfällt oder ist Zusatz (*). Wenn die Zeit knapp ist, genügen die drei letzten Einträge (17. Juni, 20. Juni, Anfang Juli).");
  phaseTag(s, 3);
}

// ============ 12 Zeitstrahl ============
{
  const s = blank("V · Verarbeiten · Kontrolle", "Der Dritte Stand handelt: Zeitstrahl", 3, "Nach dem Ausfüllen zeigen (1 Minute). Ansage: „Vergleicht und verbessert mit Grün.“ Der letzte Eintrag („Anfang Juli: Soldaten um Paris“) ist das Ziel der Stunde. Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung. Historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Quelle Sieyès: verfasst 1788, erschienen Winter 1788/89.");
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


// ============ 13 Abschlussfrage Freund (Bauer) ============
{
  const s = blank("A · Auswerten", "Abschlussfrage: Erkläre es deinem Freund", 4,"Ansage: „Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Denkt zehn Sekunden nach. Ich rufe euch auf.“ Die Antwort kommt mündlich (Blitzlicht, 4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar), kein Schnipsel, kein Schreiben. Es gibt mehrere gute Antworten, geprüft wird die Begründung: eine Position, ein Sachargument aus der Stunde, ein „weil“. Der Freund ist selbst Bauer, es geht um einen Rat unter Freunden, nicht um Gefahr. Erwartung (Auswahl, keine Musterlösung): ANSCHLIESSEN: Der Dritte Stand (Bauern und Bürger) sind fast alle Menschen (Sieyès: ALLES); die Bauern tragen Steuern und Abgaben, Klerus und Adel zahlen kaum; gemeinsam sind sie stark, ihre Vertreter bilden die Nationalversammlung; mit einer Verfassung entscheidet der König nicht mehr allein; wer abseits steht, hat keine Stimme. ABWARTEN: Die Ernte und die Familie brauchen Zeit und Kraft; es ist unklar, ob der König nachgibt; Veränderungen dauern lange; man kann die Vertreter erst beobachten. Zwischenlösung (die Beschwerden mitunterschreiben, aber weiterarbeiten) ist begründbar. Form nach Verlauf: Blitzlicht oder Rat-Duell zu zweit (leise). Transfer später: Das Muster „Erkläre deinem Freund, warum …“ funktioniert in jeder Epoche (Wahlen, Widerstand, Beitritt zu einer Bewegung). Gezielt aufrufen, nicht kommentieren (unbenotet). Die Figuren sind Zeichnungen, keine historischen Porträts.");
  s.addImage({ path: IMG + "bauer.png", x: 0.55, y: 1.45, w: 2.9, h: 4.19 });
  txt(s, "Dein Freund", { x: 0.55, y: 5.6, w: 2.9, h: 0.4, fontFace: HEAD, fontSize: 27, bold: true, align: "center" });
  bubble(s, 4.0, 1.5, 8.75, 1.95, WHITE, NAVY, "left", 2.5);
  txt(s, "„Die Abgeordneten des Dritten Standes (Bauern und Bürger) wollen eine Verfassung. Soll ich mich ihnen anschließen? Ich habe doch genug mit der Ernte zu tun.“", { x: 4.3, y: 1.5, w: 8.2, h: 1.95, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle" });
  bubble(s, 4.0, 3.75, 7.1, 1.6, "FDF1D8", GOLD, "right", 4.55);
  txt(s, "Du erklärst:  „Ich rate dir, …, weil …“", { x: 4.3, y: 3.75, w: 6.6, h: 1.6, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle" });
  s.addImage({ path: IMG + "berater.png", x: 11.75, y: 3.4, w: 1.3, h: 1.8 });
  txt(s, "Du", { x: 11.75, y: 5.2, w: 1.3, h: 0.4, fontFace: HEAD, fontSize: 27, bold: true, align: "center" });
  [["Position", 2.1], ["Sachargument", 3.2], ["„weil“", 1.9]].reduce((xx, [lab, ww], i) => {
    box(s, xx, 5.55, ww, 0.62, i % 2 ? LIGHT : LIGHT2);
    s.addText([{ text: lab, options: { bold: true, fontFace: HEAD, fontSize: 27, color: NAVY } }], { isTextBox: true, x: xx + 0.1, y: 5.55, w: ww - 0.2, h: 0.62, valign: "middle", align: "center", margin: 0 });
    return xx + ww + 0.2;
  }, 4.0);
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "2 Minuten · erst denken, dann mündlich      ", options: {} }, { text: "Alle denken mit.", options: { italic: true, color: ICE } }]);
}

// ============ 14 Ziele prüfen ============
{
  const s = blank("A · Auswerten", "Haben wir die Ziele erreicht?", 4, "Die Klasse prüft still: Kann ich das? Danach Folie 15. Keine Antworten ins Plenum, keine Abfrage von Selbsteinschätzungen (30 Sekunden). Wenn Ziel 3 (*) nicht mehr geschafft wurde, kommt es als Wiederholung in Stunde 3.");
  [["Ich kann erklären, wie die Stände den Dritten Stand (Bauern und Bürger) überstimmen können."], ["Ich kann die Abstimmung nach Ständen beurteilen und mein Urteil begründen."], ["Ich kann am Zeitstrahl erklären, wie sich der Dritte Stand zur Nationalversammlung erklärte.*"]].forEach(([x], i) => {
    const y = 1.7 + i * 1.4;
    box(s, 0.6, y, 12.15, 1.15, i % 2 ? LIGHT : LIGHT2);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    txt(s, x, { x: 1.75, y, w: 10.8, h: 1.15, fontFace: HEAD, fontSize: 27, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schau auf die Ziele. Haben wir sie erreicht?", options: {} }]);
}

// ============ 15 Ausblick ============
{
  const s = blank("A · Auswerten · Ausblick", "Wie geht es weiter?", 4, "Folie nur lesen (30 Sekunden). Ausblick: Anfang Juli zieht der König Soldaten um Paris zusammen und entlässt seinen Finanzminister. In Paris wächst die Angst. Das sehen wir in der nächsten Stunde (Bastille, Ende der Vorrechte, Menschenrechte).");
  box(s, 0.6, 1.8, 12.15, 2.0, NAVY);
  txt(s, "Der König zieht Soldaten zusammen.\nIn Paris wächst die Angst.", { x: 0.95, y: 1.8, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  box(s, 0.6, 4.1, 12.15, 1.6, LIGHT2);
  txt(s, "Wie reagieren die Menschen in Paris? Das sehen wir in der nächsten Stunde.", { x: 0.95, y: 4.1, w: 11.45, h: 1.6, fontFace: HEAD, fontSize: 28, valign: "middle" });
  bar(s, [{ text: "Nächste Stunde: ", options: { bold: true } }, { text: "Bastille, Ende der Vorrechte, Menschenrechte.", options: {} }]);
}

pres.writeFile({ fileName: "../Praesentation_Stunde2_Nationalversammlung_8D.pptx" }).then(() => console.log("ok"));
