const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Hilfsverben erkennen – Merksätze und Übungen";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Hilfsverben erkennen · Deutsch 7A";
let pageNo = 1;

function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 18, color: NAVY, margin: 0, valign: "top" }, o));
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: o.r ?? 0.08, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" } });
}
function numCircle(s, n, x, y, d = 0.5, col = NAVY) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
}

// Grundgerüst jeder Aufgabenfolie: Kopf, Arbeitsschritte, Fußleiste mit Zeit
function taskSlide(kicker, title, stepsArr, meta, notes, done = "Fertig? Stift hinlegen und noch einmal durchlesen.") {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.35, w: 8, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  // Arbeitsschritte
  const hs = stepsArr.map((st) => (st.length > 68 ? 0.8 : 0.45));
  const total = hs.reduce((a, b) => a + b, 0) + (hs.length - 1) * 0.08;
  box(s, 0.6, 1.55, 12.15, total + 0.3, LIGHT2);
  let yy = 1.7;
  stepsArr.forEach((st, i) => {
    numCircle(s, i + 1, 0.8, yy + 0.015, 0.42);
    txt(s, st, { x: 1.4, y: yy, w: 11.2, h: hs[i], fontSize: 21, bold: true, valign: "top" });
    yy += hs[i] + 0.08;
  });
  // Fußleiste
  box(s, 0.6, 6.45, 12.15, 0.6, NAVY);
  s.addText([
    { text: "Zeit: ", options: { bold: true } }, { text: meta + "      ", options: {} },
    { text: done, options: { italic: true, color: ICE } },
  ], { isTextBox: true, x: 0.9, y: 6.45, w: 11.6, h: 0.6, fontFace: BODY, fontSize: 17, color: WHITE, valign: "middle", margin: 0 });
  txt(s, FOOT, { x: 0.6, y: 7.12, w: 6, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo), { x: 12.2, y: 7.12, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
  if (notes) s.addNotes(notes);
  return { s, top: 1.55 + total + 0.3 + 0.3 };
}
// Liste von Lückensätzen
function sentences(s, list, top, o = {}) {
  const cols = o.cols || 1, rows = Math.ceil(list.length / cols);
  const avail = 6.3 - top, rh = Math.min(0.72, avail / rows);
  const cw = (12.15 - (cols - 1) * 0.3) / cols;
  list.forEach((sent, i) => {
    const c = Math.floor(i / rows), r = i % rows;
    const x = 0.6 + c * (cw + 0.3), y = top + r * rh;
    box(s, x, y, cw, rh - 0.07, LIGHT);
    numCircle(s, i + 1, x + 0.15, y + (rh - 0.07 - 0.4) / 2, 0.4, RED);
    const parts = Array.isArray(sent) ? sent : [{ text: sent, options: {} }];
    s.addText(parts, { isTextBox: true, x: x + 0.75, y, w: cw - 0.9, h: rh - 0.07, fontFace: HEAD, fontSize: o.size || 24, color: NAVY, valign: "middle", margin: 0 });
  });
}
const gapSent = (a, b, hint) => [{ text: a, options: {} }, { text: "_____", options: { bold: true, color: RED } }, { text: b, options: {} }].concat(hint ? [{ text: "  (" + hint + ")", options: { color: MUTED, italic: true, fontSize: 20 } }] : []);


// ---------- Merksatz-Folien ----------
// Kurzsyntax: **fett**, [r:rot] Hilfsverb, [y:gelb] Partizip II, [b:blau] Infinitiv, [g:grün] Modalverb
function rich(str, base = {}) {
  const out = [];
  str.split(/(\*\*[^*]+\*\*|\[[rybg]:[^\]]+\])/).filter(Boolean).forEach((tok) => {
    let m;
    if ((m = tok.match(/^\*\*([^*]+)\*\*$/))) out.push({ text: m[1], options: Object.assign({}, base, { bold: true }) });
    else if ((m = tok.match(/^\[([rybg]):([^\]]+)\]$/))) out.push({ text: m[2], options: Object.assign({}, base, { bold: true, color: { r: RED, y: GOLD, b: BLUE, g: "2E9E6B" }[m[1]] }) });
    else out.push({ text: tok, options: Object.assign({}, base) });
  });
  return out;
}
function merkSlide(nr, title, kern, beispiel, wofuer, test, notes) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, `MERKSATZ ${nr} · ABSCHREIBEN UND ROT UMRAHMEN`, { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: RED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  // Kernsatz im roten Rahmen
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.55, w: 12.15, h: 2.0, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  s.addText(rich(kern), { isTextBox: true, x: 0.95, y: 1.65, w: 11.45, h: 1.8, fontFace: HEAD, fontSize: 22, color: NAVY, valign: "middle", margin: 0 });
  // Beispiel
  box(s, 0.6, 3.7, 12.15, 0.6, LIGHT2);
  s.addText([{ text: "Beispiel:   ", options: { bold: true, color: MUTED, fontFace: BODY, fontSize: 17 } }, ...rich(beispiel)], { isTextBox: true, x: 0.9, y: 3.7, w: 11.6, h: 0.6, fontFace: HEAD, fontSize: 20, color: NAVY, valign: "middle", margin: 0 });
  // Wofür? / Test
  [["Wofür brauche ich das?", wofuer, 0.6], ["Test", test, 6.8]].forEach(([h, body, x]) => {
    box(s, x, 4.45, 5.95, 1.8, LIGHT);
    txt(s, h, { x: x + 0.25, y: 4.58, w: 5.45, h: 0.4, fontFace: HEAD, fontSize: 19, bold: true, color: RED });
    s.addText(rich(body), { isTextBox: true, x: x + 0.25, y: 5.0, w: 5.45, h: 1.15, fontFace: BODY, fontSize: 19, color: NAVY, valign: "top", margin: 0 });
  });
  box(s, 0.6, 6.45, 12.15, 0.6, NAVY);
  s.addText([{ text: "Jetzt: ", options: { bold: true } }, { text: "Schreibe den Merksatz mit Beispiel ab und rahme ihn rot ein.      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }], { isTextBox: true, x: 0.9, y: 6.45, w: 11.6, h: 0.6, fontFace: BODY, fontSize: 17, color: WHITE, valign: "middle", margin: 0 });
  txt(s, FOOT, { x: 0.6, y: 7.12, w: 6, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo), { x: 12.2, y: 7.12, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
  s.addNotes(notes || "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz " + nr + ". Rahmt ihn rot ein.“ Warten, bis alle den Stift hingelegt haben.");
}

// ============ Titel ============
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "Hilfsverben erkennen", { x: 0.9, y: 2.5, w: 11, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  txt(s, "Merksätze und Übungen der Stunde", { x: 0.9, y: 3.55, w: 11, h: 0.5, fontSize: 22, italic: true, color: ICE });
  txt(s, "Deutsch, Klasse 7A", { x: 0.9, y: 6.6, w: 5, h: 0.3, fontSize: 13, color: ICE });
  s.addNotes("Merksätze zum Abschreiben und Übungen. Die Erarbeitung (Sätze ohne Hilfsverb, Klammer) passiert an der Tafel – siehe Tafelskript. Lösungen stehen in den Notizen.");
}

// ============ Reaktivierung ============
{
  const { s, top } = taskSlide("Wiederholung · Kreide-Kette", "Welche Verbform ist das?", ["Nimm ein Wort von der Liste.", "Schreibe es an der Tafel in die richtige Spalte.", "Gib die Kreide weiter."], "8 Minuten · Kreide-Kette",
    "Lösung: Infinitiv: spielen, gehen, lesen · finit: spielst, war, hat, bist, hatten · Partizip II: gespielt, gegangen, gelesen, gekauft. Danach fragen: Woran habt ihr das Partizip II erkannt?", "Alle anderen prüfen mit.");
  const words = ["spielen", "gespielt", "spielst", "war", "gegangen", "gehen", "hat", "gelesen", "lesen", "bist", "hatten", "gekauft"];
  const cw = 1.85, ch = 0.62;
  words.forEach((w, i) => {
    const x = 0.6 + (i % 6) * (cw + 0.21), y = top + Math.floor(i / 6) * (ch + 0.15);
    box(s, x, y, cw, ch, WHITE, { line: NAVY, lw: 1.5 });
    txt(s, w, { x, y, w: cw, h: ch, fontFace: HEAD, fontSize: 21, bold: true, align: "center", valign: "middle" });
  });
  const y2 = top + 2 * (ch + 0.15) + 0.1;
  [["Infinitiv", BLUE], ["finite Form", NAVY], ["Partizip II", GOLD]].forEach(([h, c], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, y2, 3.95, 0.5, c, { r: 0.05 });
    txt(s, h, { x, y: y2, w: 3.95, h: 0.5, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}

// ============ Merksatz 5 ============
merkSlide(5, "Partizip II", "Das **Partizip II** ist infinit und verändert sich **nie**. Man bildet es meist mit **ge-…-t** oder **ge-…-en**. Verben auf be-, ver-, -ieren bekommen kein ge-.", "[y:gespielt]  ·  [y:gelaufen]  ·  [y:bestellt]  ·  [y:verstanden]  ·  [y:telefoniert]", "Mit dem Partizip II erzählt man, was schon passiert ist – aber **nie allein**.", "Passt **„ich habe …“** oder **„ich bin …“** davor?",
  "Ansage: „Heft auf. Datum nach rechts. Neue Überschrift: Hilfsverben erkennen. Dann schreibt ihr von der Folie ab: Merksatz 5.“ „Aber nie allein“ ist die Brücke zur Erarbeitung des Hilfsverbs an der Tafel.");

// ============ Merksatz 6 ============
merkSlide(6, "Das Hilfsverb", "Hilfsverben sind **haben, sein und werden**. Das Hilfsverb ist **finit** und steht an **Position 2**. Mit dem Partizip II am Satzende bildet es eine **Klammer**.", "Ich [r:habe] Fußball [y:gespielt].   ·   Wir [r:sind] ins Kino [y:gegangen].", "Das Partizip II kann nicht zeigen, **wer** etwas tut – das übernimmt das Hilfsverb. Zusammen bilden sie **Zeitformen**.", "Steht am Satzende ein **Partizip II**? Dann ist das finite haben oder sein davor ein Hilfsverb.",
  "Vorher an der Tafel: vier Sätze ohne Hilfsverb, Kreide-Kette setzt die roten Wörter ein, Klammern einzeichnen. Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 6 mit dem Beispielsatz. Zeichnet die Klammer ein.“ werden nur erwähnen – kommt beim Futur.");

// ============ Ü6 ============
{
  const { s, top } = taskSlide("Übung 6", "Hilfsverb gesucht", ["Schreibe Ü6 an den Rand.", "Finde in jedem Satz das Hilfsverb und das Partizip II.", "Schreibe nur: Nummer, Hilfsverb – Partizip II (rot und gelb)."], "4 Minuten · allein · leise",
    "Lösung: 1 hat – gezockt · 2 sind – gefahren · 3 habt – vergessen · 4 ist – hingefallen · 5 habe – gekauft · 6 hat – begonnen (kein ge-). Vergleich per Kreide-Kette: unterstreichen und Klammer zeichnen.");
  sentences(s, ["Mein Bruder hat gestern bis Mitternacht gezockt.", "Wir sind mit dem Bus zur Schule gefahren.", "Ihr habt die Hausaufgaben vergessen.", "Lea ist beim Training hingefallen.", "Ich habe mir ein neues Handy gekauft.", "Die Party hat um acht begonnen."], top, { size: 20 });
}

// ============ Merksatz 7 ============
merkSlide(7, "Hilfsverb oder Vollverb?", "haben und sein sind nur dann **Hilfsverben**, wenn am Satzende ein **Partizip II** steht. Sonst sind sie **Vollverben** und tragen selbst die Bedeutung.", "Ich **habe** Hunger. (Vollverb)   ·   Ich [r:habe] [y:gegessen]. (Hilfsverb)", "Wer Hilfsverben erkennt, kann später die **Zeitform** bestimmen.", "Steht am Satzende ein **Partizip II**?",
  "Vorher an der Tafel: „Ich habe Hunger.“ / „Ich habe gegessen.“ – Ist es beide Male ein Hilfsverb? Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 7 mit den beiden Beispielsätzen.“");

// ============ Ü7 ============
{
  const { s, top } = taskSlide("Übung 7", "Hilfsverb oder Vollverb?", ["Schreibe Ü7 an den Rand.", "Schreibe zu jeder Nummer H (Hilfsverb) oder V (Vollverb).", "Begründe bei Nr. 2 in einem Satz."], "3 Minuten · allein · leise",
    "Lösung: 1 V · 2 H · 3 V · 4 H · 5 V · 6 H. Begründung 2: Am Satzende steht das Partizip II runtergefallen. Vergleich per Kreide-Kette: H oder V hinter den Satz schreiben und Grund sagen.");
  sentences(s, ["Mein Handy ist kaputt.", "Mein Handy ist runtergefallen.", "Wir hatten keine Hausaufgaben.", "Ihr habt die Hausaufgaben vergessen.", "Ich bin total müde.", "Ich bin um elf eingeschlafen."], top, { size: 20 });
}

// ============ Exit-Ticket ============
{
  const { s, top } = taskSlide("Zum Schluss · ohne Heft", "Exit-Ticket", ["Nimm einen Zettel und schreibe deinen Namen oben hin.", "Schreibe zu jedem Satz: Hilfsverb ja oder nein – und wenn ja, das Partizip II.", "Erkläre in einem Satz: Was ist ein Hilfsverb?"], "4 Minuten · allein · ohne Heft",
    "Lösung: 1 nein (Vollverb) · 2 ja – gegangen · 3 nein (Vollverb). Erklärung z. B.: haben, sein oder werden; finit; bildet mit dem Partizip II eine Zeitform. Auswertung: „ja“ bei 1 oder 3 → Merksatz 7 nächste Stunde wiederholen.");
  sentences(s, ["Wir waren gestern im Kino.", "Wir sind gestern ins Kino gegangen.", "Sie hat einen Hund."], top, { size: 24 });
}

pres.writeFile({ fileName: "Praesentation_Hilfsverben_erkennen.pptx" }).then(() => console.log("ok"));
