const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Verbformen verstehen – Merksätze und Übungen";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Verbformen verstehen · Deutsch 7A";
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
  txt(s, "Verbformen verstehen", { x: 0.9, y: 2.5, w: 11, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  txt(s, "Merksätze und Übungen der Stunde", { x: 0.9, y: 3.55, w: 11, h: 0.5, fontSize: 22, italic: true, color: ICE });
  txt(s, "Deutsch, Klasse 7A", { x: 0.9, y: 6.6, w: 5, h: 0.3, fontSize: 13, color: ICE });
  s.addNotes("Diese Präsentation enthält die Merksätze (zum Abschreiben) und die Übungen. Tafelbilder und Tabellen entstehen weiterhin an der Tafel (siehe Tafelskript). Lösungen stehen jeweils in den Notizen.");
}

// Chat-Verlauf als Sprechblasen
const CHAT = [["Mo", "Hast du gestern das Spiel gesehen?"], ["Jona", "Nee, ich habe erst um zehn nach Hause gegangen."], ["Mo", "Schade. Wir sind echt stark gespielt."], ["Jona", "Wer hat die Tore geschießt?"], ["Mo", "Ich hatte zwei Treffer. Das dritte hat Ali gemacht."]];
function chatBubbles(s, top) {
  const rh = Math.min(0.72, (6.3 - top) / CHAT.length);
  CHAT.forEach(([who, msg], i) => {
    const left = who === "Mo", w = 10.2, x = left ? 0.6 : 12.75 - w, y = top + i * rh;
    box(s, x, y, w, rh - 0.1, left ? LIGHT : LIGHT2, { r: 0.15 });
    s.addText([{ text: who + ":  ", options: { bold: true, color: MUTED, fontFace: BODY, fontSize: 18 } }, { text: msg, options: {} }], { isTextBox: true, x: x + 0.25, y, w: w - 0.4, h: rh - 0.1, fontFace: HEAD, fontSize: 20, color: NAVY, valign: "middle", margin: 0 });
  });
}

// ============ Einstieg ============
{
  const { s, top } = taskSlide("Einstieg · ohne Heft", "Was stimmt in diesem Chat nicht?", ["Lies den Chat.", "Finde die drei Fehler bei den Verben. Melde dich."], "3 Minuten · gemeinsam · mündlich",
    "Lösung: habe … gegangen → bin gegangen · sind … gespielt → haben gespielt · geschießt → geschossen. Satz 5 ist richtig (hatte = Vollverb). Noch nicht erklären lassen – das passiert am Ende der Stunde.", "Noch nicht erklären – das machen wir am Ende.");
  chatBubbles(s, top);
}

merkSlide(1, "Infinitiv", "Der **Infinitiv** ist die **Grundform** eines Verbs. Er endet auf -en oder -n.", "ging → [b:gehen]   ·   isst → [b:essen]   ·   basteln", "So steht das Verb im Wörterbuch. Aus dem Infinitiv werden alle anderen Formen gebildet.", "Lässt sich die Form so im Wörterbuch finden?");
merkSlide(2, "Finite Verbform (Personalform)", "Die **finite** Verbform passt sich der Person und der Zahl an. Infinitiv und Partizip II sind **infinit** – sie verändern sich nicht.", "ich spiel**e**  →  wir spiel**en**", "Die finite Form zeigt, **wer** etwas tut.", "Ersetze **ich** durch **wir** – die Form, die sich ändert, ist finit.");

// ============ Ü1 ============
{
  const { s, top } = taskSlide("Übung 1", "Zurück zur Grundform", ["Schreibe Ü1 an den Rand.", "Schreibe jedes Wort ab und ergänze den Infinitiv."], "3 Minuten · allein · leise",
    "Lösung: konnte → können · liest → lesen · bin → sein · fuhr → fahren · wusste → wissen · schläft → schlafen");
  const words = ["konnte", "liest", "bin", "fuhr", "wusste", "schläft"];
  const cw = 3.85, ch = 1.15;
  words.forEach((w, i) => {
    const x = 0.6 + (i % 3) * (cw + 0.3), y = top + 0.1 + Math.floor(i / 3) * (ch + 0.3);
    box(s, x, y, cw, ch, LIGHT);
    s.addText([{ text: w + "  →  ", options: {} }, { text: "______", options: { color: BLUE, bold: true } }], { isTextBox: true, x: x + 0.3, y, w: cw - 0.4, h: ch, fontFace: HEAD, fontSize: 24, color: NAVY, valign: "middle", margin: 0 });
  });
}

merkSlide(3, "Formen zurückführen", "Jede finite Form gehört zu einem **Infinitiv**. Bei unregelmäßigen Verben ändert sich der Stamm stark.", "bin, ist, war → [b:sein]   ·   habe, hat, hatte → [b:haben]", "Nur wer den Infinitiv kennt, erkennt das Verb, kann es nachschlagen und seine Zeitform bestimmen.", "Frage: **Wie heißt der Infinitiv?**", "Vorher an der Tafel: Tabelle haben/sein im Präsens gemeinsam entwickeln (Präteritum-Spalten leer). Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 3. Dann zeichnet ihr die Tabelle von der Tafel mit Lineal ab – auch die leeren Spalten!“");

// ============ Ü2 ============
{
  const { s, top } = taskSlide("Übung 2", "Finite Form finden", ["Schreibe Ü2 an den Rand.", "Finde in jedem Satz die finite Verbform.", "Schreibe nur Nummer, finite Form und Infinitiv: 1 … → …"], "4 Minuten · allein · leise",
    "Lösung: 1 Bist → sein · 2 hat → haben · 3 sind → sein · 4 Habt → haben · 5 ist → sein · 6 schaut → schauen (bewusst kein haben/sein). Bei 4 nachfragen: gesehen ist infinit.");
  sentences(s, ["Bist du morgen beim Training?", "Meine Schwester hat ein neues Handy.", "Wir sind am Freitag im Kino.", "Habt ihr die neue Staffel schon gesehen?", "Mein Akku ist fast leer.", "Leon schaut jeden Abend Videos."], top, { size: 21 });
}

merkSlide(4, "haben und sein im Präteritum", "Im Präteritum heißt es: ich **hatte** (haben) und ich **war** (sein). Auch „war“ gehört zum Infinitiv sein.", "Gestern [r:war] ich müde.   ·   Gestern [r:hatte] ich keine Zeit.", "Im Präteritum erzählt man **schriftlich** von Vergangenem – in Erzählungen und Berichten.", "Setze **„gestern“** davor.", "Vorher an der Tafel: Präteritum-Spalten der Tabelle gemeinsam füllen. Ansage: „Füllt in eurem Heft die leeren Spalten aus. Dann schreibt ihr von der Folie ab: Merksatz 4.“");

// ============ Speed-Duell ============
{
  const { s, top } = taskSlide("Partnerarbeit · ohne Heft", "Speed-Duell zu zweit", ["Nenne Person, Verb und Zeitform: „wir – sein – Präteritum“.", "Dein Partner antwortet in drei Sekunden. Richtig = 1 Punkt.", "Nach 2 Minuten wechselt ihr. Wer hat mehr Punkte?"], "4 Minuten · zu zweit · Flüsterstimme",
    "Einmal mit einem Schüler vormachen. Tabelle im Heft abdecken lassen. Am Ende fragen: Welche Form war am schwierigsten? (meist wart, hattet)", "Tabelle im Heft abdecken!");
  const cols = [["Person", ["ich", "du", "er/sie/es", "wir", "ihr", "sie"], NAVY], ["Verb", ["haben", "sein"], RED], ["Zeitform", ["Präsens", "Präteritum"], MUTED]];
  const cw = [5.0, 2.9, 3.85];
  let x = 0.6;
  cols.forEach(([h, items, col], i) => {
    box(s, x, top, cw[i], 0.5, col, { r: 0.05 });
    txt(s, h, { x, y: top, w: cw[i], h: 0.5, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, items.join(" · "), { x, y: top + 0.6, w: cw[i], h: 0.55, fontFace: HEAD, fontSize: 19, align: "center", valign: "middle" });
    x += cw[i] + 0.2;
  });
  const y2 = top + 1.45;
  box(s, 0.6, y2, 12.15, 0.9, LIGHT2);
  s.addText([
    { text: "Beispiel:  ", options: { bold: true, color: MUTED, fontFace: BODY } },
    { text: "„ihr – haben – Präteritum“  →  ", options: {} },
    { text: "ihr hattet", options: { bold: true, color: RED } },
  ], { isTextBox: true, x: 0.9, y: y2, w: 11.6, h: 0.9, fontFace: HEAD, fontSize: 26, color: NAVY, valign: "middle", margin: 0 });
}

// ============ Ü3 ============
{
  const { s, top } = taskSlide("Übung 3", "Letzte Woche", ["Schreibe Ü3 an den Rand.", "Schreibe untereinander nur die Nummer und das fehlende Wort."], "3 Minuten · allein · leise",
    "Lösung: 1 war · 2 hattest · 3 waren · 4 hattet · 5 Warst");
  sentences(s, [gapSent("Gestern ", " ich beim Zahnarzt.", "sein"), gapSent("Du ", " am Wochenende Geburtstag, oder?", "haben"), gapSent("Wir ", " letzte Woche auf Klassenfahrt.", "sein"), gapSent("Ihr ", " gestern kein WLAN.", "haben"), gapSent("", " du schon einmal in einem Escape Room?", "sein")], top, { size: 22 });
}

merkSlide(5, "Partizip II", "Das **Partizip II** ist infinit und verändert sich **nie**. Man bildet es meist mit **ge-…-t** oder **ge-…-en**. Verben auf be-, ver-, -ieren bekommen kein ge-.", "[y:gespielt]  ·  [y:gelaufen]  ·  [y:bestellt]  ·  [y:verstanden]  ·  [y:telefoniert]", "Mit haben oder sein erzähle ich, was schon passiert ist (Perfekt): Ich [r:habe] [y:gespielt]. Ich [r:bin] [y:gelaufen].", "Passt **„ich habe …“** oder **„ich bin …“** davor?");

// ============ Ü4 ============
{
  const { s, top } = taskSlide("Übung 4", "Das Partizip II", ["Schreibe Ü4 an den Rand.", "Schreibe ab: Verb → Partizip II. Die ersten zwei stehen schon da.", "Markiere ge- und die Endung gelb."], "3 Minuten · allein · leise",
    "Lösung: gekauft · geschrieben · gegessen · gefahren · ★ verstanden · telefoniert · aufgeräumt (ge- in der Mitte)");
  const items = [["spielen", "gespielt"], ["gehen", "gegangen"], ["kaufen", null], ["schreiben", null], ["essen", null], ["fahren", null]];
  const cw = 3.85, ch = 0.8;
  const ys = top + 2 * (ch + 0.15) + 0.05;
  box(s, 0.6, ys, 12.15, 0.6, WHITE, { line: GOLD, lw: 1.5 });
  s.addText([{ text: "★ Für Schnelle:   ", options: { bold: true, color: GOLD, fontFace: BODY, fontSize: 18 } }, { text: "verstehen → _____    telefonieren → _____    aufräumen → _____", options: {} }], { isTextBox: true, x: 0.85, y: ys, w: 11.7, h: 0.6, fontFace: HEAD, fontSize: 18, color: NAVY, valign: "middle", margin: 0 });
  items.forEach(([inf, p2], i) => {
    const x = 0.6 + (i % 3) * (cw + 0.3), y = top + Math.floor(i / 3) * (ch + 0.15);
    box(s, x, y, cw, ch, p2 ? LIGHT2 : LIGHT);
    const right = p2 ? [{ text: "ge", options: { color: GOLD, bold: true } }, { text: p2.slice(2, p2.length - (p2.endsWith("en") ? 2 : 1)), options: {} }, { text: p2.endsWith("en") ? "en" : "t", options: { color: GOLD, bold: true } }] : [{ text: "______", options: { color: GOLD, bold: true } }];
    s.addText([{ text: inf + "  →  ", options: {} }, ...right], { isTextBox: true, x: x + 0.3, y, w: cw - 0.4, h: ch, fontFace: HEAD, fontSize: 23, color: NAVY, valign: "middle", margin: 0 });
  });
}

merkSlide(6, "Perfekt mit haben oder sein", "**Perfekt = haben oder sein im Präsens + Partizip II.** Mit sein: Verben der **Bewegung** von A nach B (gehen, fahren) und der **Veränderung** (einschlafen, aufwachen). Die meisten anderen Verben: **haben**.", "Ich [r:habe] Fußball [y:gespielt].   ·   Ich [r:bin] zum Spiel [y:gefahren].", "Mit dem Perfekt erzählt man **mündlich** von Vergangenem – im Gespräch, im Chat. Wie im Englischen: I have played.", "Bewegung von A nach B oder Veränderung? → **sein**", "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 6 mit den beiden Beispielsätzen – in Farbe.“");

// ============ Ü5 ============
{
  const { s, top } = taskSlide("Übung 5", "haben oder sein?", ["Schreibe Ü5 an den Rand.", "Schreibe nur Nummer + Wort: habe, hat, bin, ist oder sind.", "Begründe bei Nr. 3 und 6 in einem Satz, warum dort sein steht."], "5 Minuten · allein · leise",
    "Lösung: 1 bin · 2 habe · 3 sind · 4 hat · 5 sind · 6 bin. Begründung 3: fahren = Bewegung von A nach B. Begründung 6: einschlafen = Veränderung.");
  sentences(s, [gapSent("Am Samstag ", " ich erst um elf aufgewacht."), gapSent("Dann ", " ich mit meinem Bruder gezockt."), gapSent("Am Nachmittag ", " wir mit dem Rad zum See gefahren."), gapSent("Meine Freundin ", " mir ein Video geschickt."), gapSent("Abends ", " wir noch ins Kino gegangen."), gapSent("Um Mitternacht ", " ich endlich eingeschlafen.")], top, { size: 20 });
}

merkSlide(7, "Hilfsverb oder Vollverb?", "haben und sein sind nur dann **Hilfsverben**, wenn am Satzende ein **Partizip II** steht. Sonst sind sie **Vollverben** und tragen selbst die Bedeutung.", "Ich **habe** Hunger. (Vollverb)   ·   Ich [r:habe] [y:gegessen]. (Hilfsverb)", "Hilfsverben bilden Zeitformen – wer sie erkennt, kann die Zeitform bestimmen.", "Steht am Satzende ein **Partizip II**?");

// ============ Hilfsverb oder Vollverb ============
{
  const { s, top } = taskSlide("Gemeinsam · ohne Heft", "Hilfsverb oder Vollverb?", ["Melde dich.", "Sage: Hilfsverb oder Vollverb – und begründe mit dem Test: Steht am Satzende ein Partizip II?"], "gemeinsam · mündlich",
    "Lösung: 1 Vollverb · 2 Hilfsverb (runtergefallen) · 3 Vollverb · 4 Hilfsverb (vergessen)", "Wer begründet, bekommt das Wort.");
  sentences(s, ["Mein Handy ist kaputt.", "Mein Handy ist runtergefallen.", "Wir hatten keine Hausaufgaben.", "Ihr habt die Hausaufgaben vergessen."], top, { size: 24 });
}

// ============ Chat korrigieren ============
{
  const { s, top } = taskSlide("Abschluss", "Jetzt könnt ihr es erklären", ["Erkläre einen Fehler im Chat – mit den Fachbegriffen.", "Schreibe die drei korrigierten Sätze ab: Hilfsverb rot, Partizip II gelb."], "gemeinsam, dann abschreiben",
    "Erwartung: gegangen = Bewegung → bin · spielen = keine Bewegung → haben · Partizip II von schießen = geschossen. Satz 5: hatte ist Vollverb (kein Partizip II am Ende). Die korrigierten Sätze an die Tafel schreiben.");
  chatBubbles(s, top);
}

// ============ Exit-Ticket ============
{
  const { s, top } = taskSlide("Zum Schluss · ohne Heft", "Exit-Ticket", ["Nimm einen Zettel und schreibe deinen Namen oben hin.", "Schreibe zu jedem Satz: Infinitiv (haben/sein) · Präsens oder Präteritum · Hilfsverb ja oder nein."], "4 Minuten · allein · ohne Heft",
    "Lösung: 1 sein, Präteritum, nein · 2 haben, Präsens, ja · 3 haben, Präteritum, nein · 4 sein, Präsens, ja. Auswertung: „ja“ bei 1 oder 3 → Hilfsverb/Vollverb verwechselt; „Präteritum“ bei 2 oder 4 → Form des Hilfsverbs mit der Zeit des Geschehens verwechselt.");
  sentences(s, ["Wir waren im Zoo.", "Du hast gewonnen.", "Sie hatten keine Zeit.", "Ich bin nach Hause gerannt."], top, { cols: 2, size: 24 });
  const y = top + 2 * Math.min(0.72, (6.25 - top) / 2) + 0.05;
  if (y < 6.0) txt(s, "So schreibst du es auf – Beispiel „Ich habe gegessen.“:   haben · Präsens · ja", { x: 0.6, y, w: 12, h: 0.35, fontSize: 15, italic: true, color: MUTED });
}

pres.writeFile({ fileName: "Aufgaben_Verbformen_verstehen.pptx" }).then(() => console.log("ok"));
