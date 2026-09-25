const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Zeitformen bauen";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const C = { hv: "D9534F", mv: "2E9E6B", p2: "E8A33D", inf: "3F7CC4" };
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Zeitformen bauen · Deutsch 7A";
let pageNo = 0;

function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 16, color: NAVY, margin: 0, valign: "top" }, o));
}
function base(kicker, title, sub) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 2 });
  txt(s, title, { x: 0.6, y: 0.72, w: 12.1, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  if (sub) txt(s, sub, { x: 0.6, y: 1.42, w: 12, h: 0.35, fontSize: 14, italic: true, color: MUTED });
  txt(s, FOOT, { x: 0.6, y: 7.0, w: 6, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo + 1), { x: 12.2, y: 7.0, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
  return s;
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(o.round === false ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE, Object.assign({
    x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5, dashType: o.dash || "solid" } : { type: "none" },
  }, o.round === false ? {} : { rectRadius: o.r ?? 0.08 }));
}
// word block ("Baustein")
function block(s, word, x, y, w, h, col, size = 22) {
  if (col) box(s, x, y, w, h, col);
  else box(s, x, y, w, h, WHITE, { line: NAVY, lw: 1.5 });
  txt(s, word, { x, y, w, h, fontFace: HEAD, fontSize: size, bold: true, color: col ? WHITE : NAVY, align: "center", valign: "middle" });
}
function pill(s, label, x, y, w, col, size = 13) {
  box(s, x, y, w, 0.42, col, { r: 0.21 });
  txt(s, label, { x, y, w, h: 0.42, fontSize: size, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function numCircle(s, n, x, y, col = NAVY) {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color: col }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function hline(s, x, y, w, col = NAVY, wt = 2.5, o = {}) { s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: Object.assign({ color: col, width: wt }, o) }); }
function vline(s, x, y, h, col = NAVY, wt = 2.5, o = {}) { s.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: Object.assign({ color: col, width: wt }, o) }); }
function darkBg(s) {
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
}

// ============ 1 Titel ============
{
  const s = pres.addSlide();
  darkBg(s);
  pill(s, "BASTELSTUNDE", 0.9, 2.05, 2.4, C.p2, 12);
  txt(s, "Zeitformen bauen", { x: 0.9, y: 2.65, w: 10, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  txt(s, "Hilfsverben, Modalverben und Partizip II – wir basteln uns einen Tempus-Baukasten", { x: 0.9, y: 3.7, w: 11.5, h: 0.5, fontSize: 18, italic: true, color: ICE });
  const chips = [["Hilfsverb", C.hv], ["Modalverb", C.mv], ["Partizip II", C.p2], ["Infinitiv", C.inf]];
  chips.forEach(([l, c], i) => pill(s, l, 0.9 + i * 1.9, 4.55, 1.75, c, 12));
  txt(s, "Deutsch, Klasse 7A", { x: 0.9, y: 6.6, w: 5, h: 0.3, fontSize: 12, color: ICE });
  s.addNotes("Ziel der Stunde: Die Schülerinnen und Schüler sehen, dass zusammengesetzte Zeitformen aus Bausteinen bestehen (Hilfsverb/Modalverb + Partizip II/Infinitiv). Material vorab austeilen: Arbeitsblatt (Seite 1 einseitig drucken), Schere, Kleber, Buntstifte in Rot, Grün, Gelb, Blau.");
}

// ============ 2 Einstieg ============
{
  const s = base("Einstieg", "Was ist gleich – was ist anders?", "Lies die vier Sätze genau. Achte nur auf die Verben.");
  const rows = [
    [["Ich", null], ["backe", "v"], ["einen Kuchen.", null]],
    [["Ich", null], ["habe", "v"], ["einen Kuchen", null], ["gebacken.", "v"]],
    [["Ich", null], ["werde", "v"], ["einen Kuchen", null], ["backen.", "v"]],
    [["Ich", null], ["muss", "v"], ["einen Kuchen", null], ["backen.", "v"]],
  ];
  rows.forEach((r, i) => {
    const y = 2.0 + i * 1.0;
    box(s, 0.6, y, 7.6, 0.8, LIGHT);
    numCircle(s, i + 1, 0.8, y + 0.19);
    const parts = r.flatMap(([w, v]) => { const dot = w.endsWith(".") ? "." : ""; const core = dot ? w.slice(0, -1) : w; return [{ text: core, options: { bold: !!v, underline: v ? { style: "sng" } : undefined, color: NAVY } }, { text: dot + " ", options: { color: NAVY } }]; });
    s.addText(parts, { isTextBox: true, x: 1.45, y, w: 6.6, h: 0.8, fontFace: HEAD, fontSize: 22, valign: "middle", margin: 0 });
  });
  box(s, 8.7, 2.0, 4.0, 3.8, NAVY);
  txt(s, "Denkfragen", { x: 9.0, y: 2.25, w: 3.5, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true, color: WHITE });
  s.addText([
    { text: "Wie viele Verbteile hat jeder Satz?", options: { bullet: true, breakLine: true } },
    { text: "Wo stehen die Verbteile im Satz?", options: { bullet: true, breakLine: true } },
    { text: "Welche Zeit ist gemeint: jetzt, früher oder später?", options: { bullet: true } },
  ], { isTextBox: true, x: 9.0, y: 2.8, w: 3.5, h: 2.8, fontFace: BODY, fontSize: 16, color: WHITE, paraSpaceAfter: 10, valign: "top", margin: 0 });
  s.addNotes("Kurzes Unterrichtsgespräch (ca. 3 Min.). Erwartung: Satz 1 hat nur ein Verb, die anderen haben zwei Verbteile – eins auf Position 2, eins am Satzende. Satz 2 = Vergangenheit (Perfekt), Satz 3 = Zukunft (Futur I), Satz 4 = Gegenwart mit Modalverb. Noch keine Fachbegriffe verlangen – die kommen auf der nächsten Folie.");
}

// ============ 3 Bausteine ============
{
  const s = base("Kurz erinnert", "Unsere Bausteine – und ihre Farben");
  const cards = [
    ["Hilfsverb", C.hv, "rot", "haben · sein · werden", "Hilft beim Bilden der Zeitform. Steht gebeugt auf Position 2.", "habe, ist, wird, hatten, war"],
    ["Modalverb", C.mv, "grün", "können · müssen · dürfen · wollen · sollen · mögen", "Sagt, wie etwas gemeint ist (Pflicht, Erlaubnis, Wunsch …).", "kann, musst, darf, sollte"],
    ["Partizip II", C.p2, "gelb", "ge- … -t / ge- … -en", "Steht am Satzende. Achtung: manchmal ohne ge-!", "gespielt, gelaufen, verstanden, aufgeräumt"],
    ["Infinitiv", C.inf, "blau", "Grundform auf -en / -n", "Steht am Satzende nach werden oder einem Modalverb.", "backen, gehen, spielen"],
  ];
  const w = 2.9, gap = 0.2;
  cards.forEach(([name, col, farbe, forms, desc, ex], i) => {
    const x = 0.6 + i * (w + gap);
    box(s, x, 1.7, w, 4.1, LIGHT);
    pill(s, name, x + 0.25, 1.95, w - 0.5, col, 14);
    txt(s, "Farbe: " + farbe, { x: x + 0.25, y: 2.5, w: w - 0.5, h: 0.3, fontSize: 12, bold: true, color: col, align: "center" });
    txt(s, forms, { x: x + 0.25, y: 2.9, w: w - 0.5, h: 0.75, fontFace: HEAD, fontSize: 15, bold: true, color: NAVY });
    txt(s, desc, { x: x + 0.25, y: 3.7, w: w - 0.5, h: 1.1, fontSize: 14, color: NAVY });
    txt(s, "z. B. " + ex, { x: x + 0.25, y: 4.85, w: w - 0.5, h: 0.8, fontSize: 13, italic: true, color: MUTED });
  });
  box(s, 0.6, 6.05, 12.2, 0.7, NAVY);
  txt(s, "Merke: Hilfs- und Modalverb stehen auf Position 2 – ihr Partner (Partizip II oder Infinitiv) wartet am Satzende.", { x: 0.9, y: 6.05, w: 11.7, h: 0.7, fontSize: 16, bold: true, color: WHITE, valign: "middle" });
  s.addNotes("Farbcode festlegen – er gilt für Arbeitsblatt und Tafel: Hilfsverb rot, Modalverb grün, Partizip II gelb, Infinitiv blau. Partizip-Test für die Schüler: Passt „ich habe …“ oder „ich bin …“ davor? Dann ist es ein Partizip II. Sonderfälle ansprechen: Verben auf be-/ver-/er- und -ieren haben kein ge- (verstanden, telefoniert); bei trennbaren Verben steht ge- in der Mitte (auf-ge-räumt).");
}

// ============ 4 Ablauf ============
{
  const s = base("So läuft die Stunde", "Schneiden, sortieren, bauen, sichern");
  const steps = [
    ["1", "Ausschneiden", "Alle Karten von Seite 1 ausschneiden. A- und B-Karten getrennt legen.", "5 Min."],
    ["2", "Sortieren", "A-Karten in die Tabelle legen, vergleichen, aufkleben, anmalen.", "10 Min."],
    ["3", "Bauen", "B-Karten in den Tempus-Bauplan kleben. Verbklammer einzeichnen.", "15 Min."],
    ["4", "Sichern", "Bauformeln ergänzen und Tempus-Detektiv lösen.", "10 Min."],
  ];
  const w = 2.75, gap = 0.4;
  steps.forEach(([n, head, desc, time], i) => {
    const x = 0.6 + i * (w + gap);
    box(s, x, 1.8, w, 3.1, LIGHT);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 2.05, w: 0.7, h: 0.7, fill: { color: NAVY }, line: { type: "none" } });
    txt(s, n, { x: x + 0.25, y: 2.05, w: 0.7, h: 0.7, fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: HEAD });
    txt(s, time, { x: x + 1.1, y: 2.25, w: w - 1.3, h: 0.3, fontSize: 13, bold: true, color: C.p2, align: "right" });
    txt(s, head, { x: x + 0.25, y: 2.95, w: w - 0.5, h: 0.45, fontFace: HEAD, fontSize: 19, bold: true });
    txt(s, desc, { x: x + 0.25, y: 3.45, w: w - 0.5, h: 1.3, fontSize: 14 });
    if (i < 3) txt(s, "›", { x: x + w + 0.02, y: 3.0, w: 0.36, h: 0.6, fontSize: 36, bold: true, color: MUTED, align: "center", valign: "middle" });
  });
  txt(s, "Das brauchst du", { x: 0.6, y: 5.3, w: 4, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true });
  const mats = [["Schere", NAVY], ["Kleber", NAVY], ["Rot", C.hv], ["Grün", C.mv], ["Gelb", C.p2], ["Blau", C.inf]];
  mats.forEach(([m, c], i) => pill(s, m, 0.6 + i * 1.6, 5.85, 1.45, c, 13));
  s.addNotes("Zeitplan für 45 Minuten inkl. Einstieg (ca. 5 Min.). Tipp: Beim Sortieren erst legen lassen, dann mit dem Nachbarn vergleichen, erst danach kleben – so werden Fehler nicht festgeklebt. Schnelle Schüler: Profi-Aufgabe auf Seite 4.");
}

// ============ 5 Aufgabe 1 ============
{
  const s = base("Aufgabe 1", "Sortiere die A-Karten", "Arbeitsblatt Seite 2 · erst legen, dann vergleichen, dann kleben");
  const heads = [["Hilfsverben", C.hv], ["Modalverben", C.mv], ["Partizip II", C.p2]];
  heads.forEach(([h, c], i) => {
    const x = 0.6 + i * 2.55;
    box(s, x, 2.0, 2.4, 0.55, c, { r: 0.05 });
    txt(s, h, { x, y: 2.0, w: 2.4, h: 0.55, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    for (let r = 0; r < 4; r++) box(s, x, 2.7 + r * 0.72, 2.4, 0.6, WHITE, { line: "B9C6E8", lw: 1, dash: "dash", r: 0.05 });
  });
  // tips
  const tx = 8.4, tw = 4.35;
  box(s, tx, 2.0, tw, 2.05, LIGHT);
  txt(s, "Partizip-Test", { x: tx + 0.25, y: 2.15, w: tw - 0.5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: C.p2 });
  s.addText([
    { text: "Passt ", options: {} }, { text: "„ich habe …“", options: { bold: true } }, { text: " oder ", options: {} }, { text: "„ich bin …“", options: { bold: true } },
    { text: " davor? Dann ist es ein Partizip II: ich habe ", options: {} }, { text: "gespielt", options: { bold: true, color: C.p2 } }, { text: ", ich bin ", options: {} }, { text: "gelaufen", options: { bold: true, color: C.p2 } }, { text: ".", options: {} },
  ], { isTextBox: true, x: tx + 0.25, y: 2.6, w: tw - 0.5, h: 1.35, fontFace: BODY, fontSize: 15, color: NAVY, valign: "top", margin: 0 });
  box(s, tx, 4.25, tw, 1.7, C.hv);
  txt(s, "Achtung: 4 Stolperkarten!", { x: tx + 0.25, y: 4.4, w: tw - 0.5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: WHITE });
  txt(s, "Vier Karten passen in keine Spalte. Klebe sie unten in das Stolper-Feld.", { x: tx + 0.25, y: 4.85, w: tw - 0.5, h: 1.0, fontSize: 15, color: WHITE });
  box(s, 0.6, 6.0, 12.15, 0.7, LIGHT2);
  txt(s, "Danach: Jede Karte in der Farbe ihrer Spalte anmalen.", { x: 0.9, y: 6.0, w: 11.6, h: 0.7, fontSize: 16, bold: true, valign: "middle" });
  s.addNotes("Einzelarbeit, dann Partnervergleich. Typische Fehler: „spielte“ wird für ein Partizip gehalten (ist aber Präteritum), „verstanden“ wird nicht erkannt, weil das ge- fehlt. Hier gezielt den Partizip-Test anwenden lassen.");
}

// ============ 6 Lösung 1 ============
{
  const s = base("Lösung · Aufgabe 1", "So sehen die Spalten aus");
  const cols = [
    ["Hilfsverben", C.hv, ["habe", "hast", "bist", "ist", "wird", "werden", "hatten", "war"]],
    ["Modalverben", C.mv, ["kann", "musst", "darf", "könnt", "wollen", "müssen", "sollte", "mochte"]],
    ["Partizip II", C.p2, ["gespielt", "gelaufen", "gegessen", "geschrieben", "gefahren", "gekauft", "verstanden", "aufgeräumt"]],
    ["Stolperkarten", MUTED, ["spielte", "ging", "laufen", "gehst"]],
  ];
  const w = 2.9, gap = 0.2;
  cols.forEach(([h, c, words], i) => {
    const x = 0.6 + i * (w + gap);
    box(s, x, 1.65, w, 0.5, c, { r: 0.05 });
    txt(s, h, { x, y: 1.65, w, h: 0.5, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    words.forEach((wd, j) => {
      const col = j % 2, rw = (w - 0.1) / 2;
      const y = 2.3 + Math.floor(j / 2) * 0.62;
      box(s, x + col * (rw + 0.1), y, rw, 0.5, i === 3 ? LIGHT2 : LIGHT, { r: 0.05 });
      txt(s, wd, { x: x + col * (rw + 0.1), y, w: rw, h: 0.5, fontFace: HEAD, fontSize: 14, bold: true, color: i === 3 ? MUTED : c, align: "center", valign: "middle" });
    });
  });
  txt(s, "Präteritum · Präteritum · Infinitiv · Präsens", { x: 9.9, y: 3.6, w: 2.9, h: 0.5, fontSize: 12, italic: true, color: MUTED, align: "center" });
  box(s, 0.6, 5.0, 12.2, 1.6, LIGHT2);
  txt(s, "Aufgepasst", { x: 0.9, y: 5.15, w: 4, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: C.hv });
  s.addText([
    { text: "verstanden", options: { bold: true, color: C.p2 } }, { text: " hat kein ge- (Vorsilbe ver-).   ", options: {} },
    { text: "aufgeräumt", options: { bold: true, color: C.p2 } }, { text: " hat das ge- in der Mitte.   ", options: {} },
    { text: "spielte", options: { bold: true, color: MUTED } }, { text: " ist kein Partizip, sondern Präteritum – dafür braucht man keinen Helfer.", options: {} },
  ], { isTextBox: true, x: 0.9, y: 5.6, w: 11.6, h: 0.9, fontFace: BODY, fontSize: 15, color: NAVY, valign: "top", margin: 0 });
  s.addNotes("Ergebnis vergleichen lassen, Fehler korrigieren BEVOR geklebt wird (falls noch nicht geschehen). Frage an die Klasse: Warum brauchen die Stolperkarten keinen Partner? → Es sind einfache Verbformen (Präsens/Präteritum) bzw. der Infinitiv allein.");
}

// ============ 7 Verbklammer ============
{
  const s = base("Neu: Die Verbklammer", "Zwei Verbteile – eine Klammer");
  const cols = [["Vorfeld", 0.9, 1.4], ["Position 2", 2.5, 2.1], ["Mittelfeld", 4.8, 3.0], ["Satzende", 8.0, 2.6]];
  cols.forEach(([l, x, w]) => txt(s, l, { x, y: 1.65, w, h: 0.3, fontSize: 12, bold: true, color: MUTED, align: "center" }));
  const ex = (y, w2, c2, w4, c4, label) => {
    block(s, "Ich", 0.9, y, 1.4, 0.85, null);
    block(s, w2, 2.5, y, 2.1, 0.85, c2);
    block(s, "einen Kuchen", 4.8, y, 3.0, 0.85, null);
    block(s, w4, 8.0, y, 2.6, 0.85, c4);
    const cx1 = 3.55, cx2 = 9.3, yb = y + 0.85;
    vline(s, cx1, yb, 0.35); vline(s, cx2, yb, 0.35); hline(s, cx1, yb + 0.35, cx2 - cx1);
    txt(s, label, { x: cx1, y: yb + 0.42, w: cx2 - cx1, h: 0.35, fontSize: 14, bold: true, italic: true, color: NAVY, align: "center" });
  };
  ex(2.0, "habe", C.hv, "gebacken", C.p2, "Verbklammer: Hilfsverb + Partizip II");
  ex(4.0, "muss", C.mv, "backen", C.inf, "Verbklammer: Modalverb + Infinitiv");
  box(s, 10.95, 2.0, 1.85, 2.85, LIGHT);
  txt(s, "Position 2 = gebeugtes Verb (passt zu „ich“)", { x: 11.1, y: 2.15, w: 1.55, h: 2.6, fontSize: 13, color: NAVY });
  box(s, 0.6, 6.0, 12.2, 0.7, LIGHT2);
  s.addText([
    { text: "Ohne Helfer keine Klammer: ", options: { bold: true } },
    { text: "Ich backe einen Kuchen.  ·  Ich backte einen Kuchen.", options: { fontFace: HEAD } },
  ], { isTextBox: true, x: 0.9, y: 6.0, w: 11.6, h: 0.7, fontFace: BODY, fontSize: 16, color: NAVY, valign: "middle", margin: 0 });
  s.addNotes("Kernidee der Stunde: Das Hilfs- oder Modalverb steht gebeugt auf Position 2, der Partner (Partizip II oder Infinitiv) steht ganz am Ende. Beide zusammen umklammern den Rest des Satzes. An der Tafel mit Bogen einzeichnen – genau das sollen die Schüler im Bauplan auch tun.");
}

// ============ 8 Aufgabe 2 ============
{
  const s = base("Aufgabe 2", "Baue den Kuchen-Satz in allen Zeitformen", "Arbeitsblatt Seite 3 · B-Karten in die dicken Kästen kleben");
  const heads = [["Zeitform", 0.6, 2.2], ["Vorfeld", 2.9, 1.0], ["Position 2", 4.0, 1.9], ["Mittelfeld", 6.0, 2.0], ["Satzende", 8.1, 2.6]];
  heads.forEach(([h, x, w]) => { box(s, x, 1.9, w, 0.45, NAVY, { r: 0.04 }); txt(s, h, { x, y: 1.9, w, h: 0.45, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" }); });
  const zs = ["Präsens", "Perfekt", "Futur I"];
  zs.forEach((z, i) => {
    const y = 2.5 + i * 0.75;
    box(s, 0.6, y, 2.2, 0.62, LIGHT, { r: 0.04 });
    txt(s, z, { x: 0.75, y, w: 2.0, h: 0.62, fontSize: 15, bold: true, valign: "middle" });
    txt(s, "Ich", { x: 2.9, y, w: 1.0, h: 0.62, fontFace: HEAD, fontSize: 17, align: "center", valign: "middle" });
    box(s, 4.0, y, 1.9, 0.62, WHITE, { line: NAVY, lw: 2, r: 0.04 });
    txt(s, "einen Kuchen", { x: 6.0, y, w: 2.0, h: 0.62, fontFace: HEAD, fontSize: 17, align: "center", valign: "middle" });
    box(s, 8.1, y, 2.6, 0.62, WHITE, { line: NAVY, lw: 2, r: 0.04 });
  });
  txt(s, "… und so weiter für alle 7 Zeilen", { x: 0.6, y: 4.8, w: 6, h: 0.35, fontSize: 13, italic: true, color: MUTED });
  box(s, 11.0, 1.9, 1.8, 3.25, LIGHT2);
  txt(s, "B-Karten: backe, backte, habe, hatte, werde (2×), muss, gebacken (3×), backen (2×), haben", { x: 11.15, y: 2.05, w: 1.5, h: 3.0, fontSize: 12, color: NAVY });
  const tips = [
    ["1", "Klebe erst die Karte für Position 2, dann das Satzende."],
    ["2", "Manche Zeilen brauchen am Satzende gar nichts – welche?"],
    ["3", "Anmalen im Farbcode und die Verbklammer als Bogen einzeichnen."],
  ];
  tips.forEach(([n, t], i) => {
    const x = 0.6 + i * 4.1;
    numCircle(s, n, x, 5.55, C.hv);
    txt(s, t, { x: x + 0.55, y: 5.45, w: 3.4, h: 0.9, fontSize: 14, valign: "middle" });
  });
  s.addNotes("Partnerarbeit möglich. Hinweis: Futur II braucht zwei Karten am Satzende (gebacken + haben). Präsens und Präteritum bleiben am Satzende leer – das ist der Aha-Moment: einfache Zeitformen haben keine Verbklammer. Schnelle Paare dürfen die Fragen a)–d) unter der Tabelle schon beantworten.");
}

// ============ 9 Lösung 2 ============
{
  const s = base("Lösung · Aufgabe 2", "Der fertige Tempus-Bauplan");
  const rows = [
    ["Präsens", ["backe", null], [], "Verb im Präsens"],
    ["Präteritum", ["backte", null], [], "Verb im Präteritum"],
    ["Perfekt", ["habe", C.hv], [["gebacken", C.p2]], "haben/sein (Präsens) + Partizip II"],
    ["Plusquamperfekt", ["hatte", C.hv], [["gebacken", C.p2]], "haben/sein (Präteritum) + Partizip II"],
    ["Futur I", ["werde", C.hv], [["backen", C.inf]], "werden + Infinitiv"],
    ["Futur II", ["werde", C.hv], [["gebacken", C.p2], ["haben", C.hv]], "werden + Partizip II + haben/sein"],
    ["mit Modalverb", ["muss", C.mv], [["backen", C.inf]], "Modalverb + Infinitiv"],
  ];
  const X = { z: 0.6, zw: 2.3, ich: 3.0, iw: 0.55, p2: 3.6, pw: 1.3, mf: 5.0, mw: 1.7, se: 6.8, sw: 2.75, f: 9.7, fw: 3.1 };
  const y0 = 1.6, rh = 0.62;
  rows.forEach(([z, pos2, end, formel], i) => {
    const y = y0 + i * rh;
    const simple = i < 2;
    box(s, X.z, y, X.zw, rh - 0.1, simple ? LIGHT2 : LIGHT, { r: 0.04 });
    txt(s, z, { x: X.z + 0.12, y, w: X.zw - 0.2, h: rh - 0.1, fontSize: 15, bold: true, valign: "middle", color: i === 6 ? C.mv : NAVY });
    txt(s, "Ich", { x: X.ich, y, w: X.iw, h: rh - 0.1, fontFace: HEAD, fontSize: 16, align: "center", valign: "middle" });
    block(s, pos2[0], X.p2, y, X.pw, rh - 0.1, pos2[1], 16);
    txt(s, "einen Kuchen", { x: X.mf, y, w: X.mw, h: rh - 0.1, fontFace: HEAD, fontSize: 16, align: "center", valign: "middle" });
    if (end.length === 0) box(s, X.se, y, X.sw, rh - 0.1, WHITE, { line: "B9C6E8", lw: 1, dash: "dash", r: 0.04 });
    else if (end.length === 1) block(s, end[0][0], X.se, y, X.sw, rh - 0.1, end[0][1], 16);
    else { block(s, end[0][0], X.se, y, 1.55, rh - 0.1, end[0][1], 16); block(s, end[1][0], X.se + 1.65, y, X.sw - 1.65, rh - 0.1, end[1][1], 16); }
    txt(s, formel, { x: X.f, y, w: X.fw, h: rh - 0.1, fontSize: 12, italic: true, color: MUTED, valign: "middle" });
  });
  const yb = y0 + 7 * rh + 0.12;
  box(s, 0.6, yb, 5.9, 0.62, LIGHT2);
  txt(s, "Präsens & Präteritum: ein Verbteil, keine Klammer", { x: 0.8, y: yb, w: 5.6, h: 0.62, fontSize: 14, bold: true, valign: "middle" });
  box(s, 6.8, yb, 5.95, 0.62, NAVY);
  txt(s, "Alle anderen: Helfer auf Pos. 2 + Partner am Ende", { x: 7.0, y: yb, w: 5.6, h: 0.62, fontSize: 14, bold: true, color: WHITE, valign: "middle" });
  s.addNotes("Vergleich über Dokumentenkamera oder diese Folie. Besonders herausstellen: Perfekt und Plusquamperfekt sehen fast gleich aus – der einzige Unterschied ist die Zeitform des Hilfsverbs (habe ↔ hatte). Genauso Futur I und Futur II: Bei Futur II steht am Ende Partizip II + haben/sein.");
}

// ============ 10 Zeitstrahl ============
{
  const s = base("Wann passiert es?", "Die Zeitformen auf dem Zeitstrahl");
  const y = 3.35;
  s.addShape(pres.shapes.LINE, { x: 0.7, y, w: 12.0, h: 0, line: { color: NAVY, width: 3, endArrowType: "triangle" } });
  vline(s, 6.7, y - 0.55, 1.1, C.hv, 2, { dashType: "dash" });
  txt(s, "JETZT", { x: 6.2, y: y - 0.9, w: 1.0, h: 0.3, fontSize: 12, bold: true, color: C.hv, align: "center" });
  const pts = [
    [1.9, "Plusquamperfekt", "vor der Vergangenheit", [["Ich ", null], ["hatte", C.hv], [" den Kuchen ", null], ["gebacken", C.p2], [".", null]]],
    [4.3, "Präteritum / Perfekt", "Vergangenheit", [["Ich backte den Kuchen. / Ich ", null], ["habe", C.hv], [" ihn ", null], ["gebacken", C.p2], [".", null]]],
    [6.7, "Präsens", "Gegenwart", [["Ich backe den Kuchen.", null]]],
    [9.1, "Futur II", "in der Zukunft schon fertig", [["Ich ", null], ["werde", C.hv], [" ihn ", null], ["gebacken", C.p2], [" ", null], ["haben", C.hv], [".", null]]],
    [11.5, "Futur I", "Zukunft", [["Ich ", null], ["werde", C.hv], [" den Kuchen ", null], ["backen", C.inf], [".", null]]],
  ];
  pts.forEach(([cx, name, when, sent]) => {
    s.addShape(pres.shapes.OVAL, { x: cx - 0.16, y: y - 0.16, w: 0.32, h: 0.32, fill: { color: WHITE }, line: { color: NAVY, width: 3 } });
    txt(s, name, { x: cx - 1.18, y: 1.85, w: 2.36, h: 0.4, fontFace: HEAD, fontSize: 14, bold: true, align: "center" });
    txt(s, when, { x: cx - 1.15, y: 2.25, w: 2.3, h: 0.3, fontSize: 12, italic: true, color: MUTED, align: "center" });
    box(s, cx - 1.1, 3.85, 2.2, 1.55, LIGHT);
    s.addText(sent.map(([t, c]) => ({ text: t, options: { bold: !!c, color: c || NAVY } })), { isTextBox: true, x: cx - 1.0, y: 3.95, w: 2.0, h: 1.35, fontFace: HEAD, fontSize: 14, align: "center", valign: "middle", margin: 0 });
  });
  box(s, 0.6, 5.85, 12.2, 0.85, NAVY);
  txt(s, "Perfekt und Präteritum meinen beide die Vergangenheit. Das Perfekt benutzen wir meist beim Sprechen, das Präteritum beim Erzählen und Schreiben.", { x: 0.9, y: 5.85, w: 11.6, h: 0.85, fontSize: 15, color: WHITE, valign: "middle" });
  s.addNotes("Optionaler Exkurs, falls die Zeitbedeutung noch unsicher ist. Plusquamperfekt = etwas war schon vorbei, bevor etwas anderes in der Vergangenheit passiert ist („Als die Gäste kamen, hatte ich den Kuchen schon gebacken.“). Futur II = etwas wird zu einem Zeitpunkt in der Zukunft schon abgeschlossen sein („Bis heute Abend werde ich den Kuchen gebacken haben.“).");
}

// ============ 11 Tempus-Detektiv (Entscheidungsbaum) ============
{
  const s = base("Tempus-Detektiv", "So findest du jede Zeitform heraus");
  // Frage 1
  box(s, 0.6, 1.65, 5.4, 0.85, NAVY);
  numCircle(s, 1, 0.8, 1.86, C.p2);
  txt(s, "Steht am Satzende ein Partizip II oder ein Infinitiv?", { x: 1.4, y: 1.65, w: 4.5, h: 0.85, fontSize: 15, bold: true, color: WHITE, valign: "middle" });
  // Nein -> einfach
  s.addShape(pres.shapes.LINE, { x: 6.0, y: 2.07, w: 1.0, h: 0, line: { color: NAVY, width: 2.5, endArrowType: "triangle" } });
  txt(s, "nein", { x: 6.0, y: 1.72, w: 1.0, h: 0.3, fontSize: 13, bold: true, color: C.hv, align: "center" });
  box(s, 7.1, 1.65, 5.65, 0.85, LIGHT2);
  s.addText([{ text: "Einfache Zeitform: ", options: { bold: true } }, { text: "Präsens (backe) oder Präteritum (backte)", options: {} }], { isTextBox: true, x: 7.3, y: 1.65, w: 5.3, h: 0.85, fontFace: BODY, fontSize: 15, color: NAVY, valign: "middle", margin: 0 });
  // Ja -> Frage 2
  s.addShape(pres.shapes.LINE, { x: 3.3, y: 2.5, w: 0, h: 0.55, line: { color: NAVY, width: 2.5, endArrowType: "triangle" } });
  txt(s, "ja", { x: 3.45, y: 2.58, w: 0.6, h: 0.3, fontSize: 13, bold: true, color: C.mv });
  box(s, 0.6, 3.1, 5.4, 0.8, NAVY);
  numCircle(s, 2, 0.8, 3.29, C.p2);
  txt(s, "Was steht auf Position 2?", { x: 1.4, y: 3.1, w: 4.5, h: 0.8, fontSize: 15, bold: true, color: WHITE, valign: "middle" });
  // Verteiler
  const w = 2.9, gap = 0.2, xs = [0, 1, 2, 3].map((i) => 0.6 + i * (w + gap));
  vline(s, 3.3, 3.9, 0.3, NAVY, 2.5);
  hline(s, xs[0] + w / 2, 4.2, xs[3] - xs[0], NAVY, 2.5);
  const outs = [
    ["haben/sein · Präsens", "habe, hast, ist, sind …", C.hv, "Perfekt", "+ Partizip II"],
    ["haben/sein · Präteritum", "hatte, hatten, war …", C.hv, "Plusquamperfekt", "+ Partizip II"],
    ["werden", "werde, wirst, wird …", C.hv, "Futur I / Futur II", "+ Infinitiv = Futur I\n+ Partizip II + haben/sein = Futur II"],
    ["ein Modalverb", "kann, muss, darf …", C.mv, "Modalverb-Satz", "+ Infinitiv"],
  ];
  outs.forEach(([h, ex, col, res, plus], i) => {
    const x = xs[i];
    s.addShape(pres.shapes.LINE, { x: x + w / 2, y: 4.2, w: 0, h: 0.3, line: { color: NAVY, width: 2.5, endArrowType: "triangle" } });
    box(s, x, 4.55, w, 2.2, LIGHT);
    txt(s, h, { x: x + 0.2, y: 4.65, w: w - 0.4, h: 0.35, fontSize: 13, bold: true, color: col });
    txt(s, ex, { x: x + 0.2, y: 5.0, w: w - 0.4, h: 0.3, fontSize: 12, italic: true, color: MUTED });
    pill(s, res, x + 0.2, 5.4, w - 0.4, NAVY, 13);
    txt(s, plus, { x: x + 0.2, y: 5.9, w: w - 0.4, h: 0.8, fontSize: 12, color: NAVY });
  });
  s.addNotes("Diese Folie ist die Denkhilfe für Aufgabe 3 – gern ausgedruckt an die Tafel hängen. Zwei Fragen reichen: 1. Gibt es einen zweiten Verbteil am Satzende? 2. Welcher Helfer steht auf Position 2 und in welcher Zeitform? Wichtig für die Unterscheidung Perfekt/Plusquamperfekt: habe/ist = Präsens → Perfekt; hatte/war = Präteritum → Plusquamperfekt.");
}

// ============ 12 Aufgabe 3 ============
const saetze = [
  "Lena ist ins Kino gegangen.",
  "Wir hatten die Hausaufgaben vergessen.",
  "Du wirst die Prüfung schaffen.",
  "Die Kinder spielten im Garten.",
  "Ich darf heute länger aufbleiben.",
  "Bis morgen werdet ihr das Buch gelesen haben.",
  "Er war zu spät gekommen.",
];
{
  const s = base("Aufgabe 3", "Tempus-Detektiv: Welche Zeitform ist das?", "Markiere die Verbteile im Farbcode und bestimme die Zeitform (Arbeitsblatt Seite 4).");
  saetze.forEach((st, i) => {
    const col = i < 4 ? 0 : 1, r = i < 4 ? i : i - 4;
    const x = 0.6 + col * 6.2, y = 1.95 + r * 0.95;
    box(s, x, y, 5.95, 0.78, LIGHT);
    numCircle(s, i + 1, x + 0.2, y + 0.18);
    txt(s, st, { x: x + 0.8, y, w: 5.0, h: 0.78, fontFace: HEAD, fontSize: 18, bold: true, valign: "middle" });
  });
  box(s, 6.8, 4.8, 5.95, 0.78, C.p2);
  txt(s, "★ Profi: „Ich gehe nach Hause.“ in Perfekt, Plusquamperfekt und Futur I", { x: 7.0, y: 4.8, w: 5.6, h: 0.78, fontSize: 14, bold: true, color: WHITE, valign: "middle" });
  s.addNotes("Einzelarbeit mit dem Entscheidungsbaum von der vorherigen Folie. Die Profi-Aufgabe trainiert das Perfekt mit „sein“ (Bewegungsverb): Ich bin … gegangen / Ich war … gegangen / Ich werde … gehen.");
}

// ============ 13 Lösung 3 ============
{
  const s = base("Lösung · Aufgabe 3", "Die Zeitformen im Überblick");
  const sol = [
    ["ist", C.hv, "gegangen", C.p2, "Perfekt"],
    ["hatten", C.hv, "vergessen", C.p2, "Plusquamperfekt"],
    ["wirst", C.hv, "schaffen", C.inf, "Futur I"],
    ["spielten", null, "—", null, "Präteritum"],
    ["darf", C.mv, "aufbleiben", C.inf, "Präsens (mit Modalverb)"],
    ["werdet", C.hv, "gelesen haben", C.p2, "Futur II"],
    ["war", C.hv, "gekommen", C.p2, "Plusquamperfekt"],
  ];
  const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 14 } });
  const rows = [[hdr("Nr."), hdr("Satz"), hdr("Position 2"), hdr("Satzende"), hdr("Zeitform")]];
  sol.forEach(([p2, c2, se, c3, z], i) => {
    const fill = { color: i % 2 ? LIGHT2 : LIGHT };
    rows.push([
      { text: String(i + 1), options: { fill, bold: true, align: "center" } },
      { text: saetze[i], options: { fill, fontFace: HEAD } },
      { text: p2, options: { fill, bold: true, color: c2 || NAVY } },
      { text: se, options: { fill, bold: true, color: c3 || MUTED } },
      { text: z, options: { fill, bold: true } },
    ]);
  });
  s.addTable(rows, { x: 0.6, y: 1.65, w: 12.15, colW: [0.7, 5.3, 1.7, 1.95, 2.5], rowH: 0.56, fontFace: BODY, fontSize: 15, color: NAVY, valign: "middle", border: { type: "solid", pt: 1, color: WHITE } });
  box(s, 0.6, 6.25, 12.15, 0.5, LIGHT2);
  s.addText([{ text: "★ Profi: ", options: { bold: true, color: C.p2 } }, { text: "Ich bin nach Hause gegangen. · Ich war nach Hause gegangen. · Ich werde nach Hause gehen.", options: {} }], { isTextBox: true, x: 0.85, y: 6.25, w: 11.7, h: 0.5, fontFace: BODY, fontSize: 14, color: NAVY, valign: "middle", margin: 0 });
  s.addNotes("Satz 5: Das Modalverb „darf“ steht im Präsens, deshalb ist der Satz Präsens. Satz 6: Futur II erkennt man an den zwei Wörtern am Satzende (gelesen haben).");
}

// ============ 14 Merke (dunkel) ============
{
  const s = pres.addSlide();
  pageNo++;
  darkBg(s);
  txt(s, "ZUM MERKEN", { x: 0.9, y: 0.6, w: 6, h: 0.3, fontSize: 12, color: ICE, charSpacing: 2 });
  txt(s, "Die Bauformeln", { x: 0.9, y: 0.95, w: 10, h: 0.8, fontFace: HEAD, fontSize: 36, bold: true, color: WHITE });
  const f = [
    ["Perfekt", [["haben/sein · Präsens", C.hv], ["Partizip II", C.p2]]],
    ["Plusquamperfekt", [["haben/sein · Präteritum", C.hv], ["Partizip II", C.p2]]],
    ["Futur I", [["werden · Präsens", C.hv], ["Infinitiv", C.inf]]],
    ["Futur II", [["werden · Präsens", C.hv], ["Partizip II", C.p2], ["haben/sein", C.hv]]],
    ["Modalverb-Satz", [["Modalverb", C.mv], ["Infinitiv", C.inf]]],
  ];
  f.forEach(([name, parts], i) => {
    const y = 2.0 + i * 0.85;
    txt(s, name, { x: 0.9, y, w: 2.8, h: 0.6, fontFace: HEAD, fontSize: 20, bold: true, color: WHITE, valign: "middle" });
    let x = 3.8;
    parts.forEach(([p, c], j) => {
      const w = p.length > 14 ? 3.0 : 2.0;
      box(s, x, y, w, 0.6, c, { r: 0.3 });
      txt(s, p, { x, y, w, h: 0.6, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
      x += w;
      if (j < parts.length - 1) { txt(s, "+", { x, y, w: 0.5, h: 0.6, fontSize: 22, bold: true, color: ICE, align: "center", valign: "middle" }); x += 0.5; }
    });
  });
  txt(s, "Präsens und Präteritum brauchen keinen Helfer – sie sind einfache Zeitformen.", { x: 0.9, y: 6.4, w: 11, h: 0.4, fontSize: 15, italic: true, color: ICE });
  s.addNotes("Abschluss: Die Schüler vergleichen ihre Bauformeln auf Seite 4 mit dieser Folie. Hausaufgabe/Übung möglich: Einen eigenen Satz in allen Zeitformen in den Bauplan schreiben.");
}

pres.writeFile({ fileName: "Praesentation_Zeitformen_bauen.pptx" }).then(() => console.log("ok"));
