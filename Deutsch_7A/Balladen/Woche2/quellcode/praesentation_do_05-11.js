const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Deutsch 7A – Donnerstag 5.11.";
const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4", VIO = "7A5BA6";
const HEAD = "Cambria", BODY = "Calibri";
let FOOT = "Mittwoch, 4.11. · Deutsch 7A";
const TECH = { ue: { name: "überfliegend", col: BLUE, sym: "»" }, or: { name: "orientierend", col: YEL, sym: "⌕" }, ge: { name: "genau", col: GREEN, sym: "✎" } };
let pageNo = 1;
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 16, color: NAVY, margin: 0, valign: "top" }, o));
}
function base(kicker, title, notes) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.4, w: 9, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 2 });
  txt(s, title, { x: 0.6, y: 0.72, w: 12.1, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  txt(s, FOOT, { x: 0.6, y: 7.0, w: 7, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo), { x: 12.2, y: 7.0, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
  if (notes) s.addNotes(notes);
  return s;
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
}
function darkBg(s) {
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
}
function num(s, n, x, y, col = NAVY, d = 0.5) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: d * 32, bold: true, color: WHITE, align: "center", valign: "middle" });
}
// Übung im festen Format: Nummer + Titel, nummerierte Schritte, Zeit, Arbeitsform, „Fertig? Stift hinlegen“
function uebung(kicker, title, steps, meta, notes, extra) {
  const s = base(kicker, title, notes);
  steps.forEach((st, i) => {
    const y = 1.75 + i * 0.8;
    num(s, i + 1, 0.6, y);
    txt(s, st, { x: 1.3, y: y + 0.04, w: 11.3, h: 0.7, fontSize: 21, bold: true });
  });
  if (extra) extra(s, 1.75 + steps.length * 0.8 + 0.1);
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
  return s;
}

const NC = { 1: "2E9E6B", 2: "5AAE62", 3: "9DB84A", 4: "E8A33D", 5: "E07B45", 6: "D9534F" };
function badge(s, n, x, y, d = 0.9) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: NC[n] }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontFace: HEAD, fontSize: d * 34, bold: true, color: WHITE, align: "center", valign: "middle" });
}

const STUFEN = {
  1: { ich: "Ich melde mich regelmäßig von mir aus. Ich erkenne das Problem, ordne es in einen größeren Zusammenhang ein, beurteile es sachlich und bringe eigene Gedanken ein. Ich drücke mich klar und mit Fachbegriffen aus.", fazit: "Die Leistung entspricht den Anforderungen in ganz besonderem Maße." },
  2: { ich: "Ich melde mich regelmäßig von mir aus. Ich verstehe auch schwierige Inhalte und ordne sie ins Thema ein. Ich unterscheide Wichtiges von Unwichtigem und weiß manchmal mehr, als im Unterricht dran war.", fazit: "Die Leistung entspricht in vollem Umfang den Anforderungen." },
  3: { ich: "Ich melde mich regelmäßig von mir aus. Was ich zum aktuellen Stoff sage, ist im Wesentlichen richtig, und ich verknüpfe es mit dem, was wir vorher in der Einheit gelernt haben.", fazit: "Die Leistung entspricht im Allgemeinen den Anforderungen." },
  4: { ich: "Ich melde mich nur gelegentlich. Ich gebe einfache Fakten aus dem gerade behandelten Stoff wieder – im Wesentlichen richtig.", fazit: "Die Leistung hat Mängel, entspricht im Ganzen aber noch den Anforderungen." },
  5: { ich: "Ich melde mich nicht von mir aus. Wenn ich drangenommen werde, ist meine Antwort nur teilweise richtig. Grundkenntnisse sind aber da.", fazit: "Die Leistung entspricht den Anforderungen nicht – die Lücken sind aber bald behebbar." },
  6: { ich: "Ich melde mich nicht von mir aus. Wenn ich drangenommen werde, ist meine Antwort falsch.", fazit: "Die Leistung entspricht den Anforderungen nicht; selbst Grundkenntnisse fehlen." },
};


// ============ Titel ============
function section(nr, title, sub, col) {
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 2.2, 0.42, col, { r: 0.21 });
  txt(s, "TEIL " + nr, { x: 0.9, y: 2.05, w: 2.2, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, title, { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, sub, { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
  pageNo++;
  return s;
}

// ================= Präsentation Do 5.11.: Wiederholung · Reime erkennen =================
FOOT = "Donnerstag, 5.11. · Reime erkennen · Deutsch 7A";
const ZIEL = "Ich kann Reime erkennen und das Reimschema einer Strophe bestimmen.";
const RC = { a: "2A8C8C", b: "C77C12", x: MUTED };
const zeit = (s, meta) => { box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2); txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 }); };
const runs = (parts, size = 22) => parts.map(([x, k]) => ({ text: x, options: Object.assign({ fontSize: size }, k === "sub" ? { bold: true, color: VIO } : k === "fin" ? { bold: true, underline: { style: "sng" } } : k === "hv" ? { bold: true, color: RED, underline: { style: "sng" } } : k === "b" ? { bold: true } : k === "m" ? { color: MUTED } : {}) }));

// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 3.4, 0.42, YEL, { r: 0.21 });
  txt(s, "DONNERSTAG · 5.11.", { x: 0.9, y: 2.05, w: 3.4, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Paddy Fingal – Reime erkennen", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, "Wie eine Ballade klingt: Reim und Reimschema", { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
}
// 2 Ziel und Ablauf
{
  const s = base("Heute", "Ziel und Ablauf", "Ziel vorlesen; am Ende im Zielcheck darauf zurückkommen. Ziel und Ablauf stehen auch an der Seitentafel.");
  box(s, 0.6, 1.7, 12.1, 1.0, LIGHT2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 0.12, h: 1.0, fill: { color: RED }, line: { type: "none" } });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 1.0, y: 1.95, w: 11.5, h: 0.6, fontSize: 22 });
  ["Wiederholung: Ü1 Fehler im Chat · Ü2 Belegen", "Reime in Strophe 1 entdecken", "Ü3 Reime erkennen (Schnipsel)", "Zielcheck", "* Nacherzählen mit Stichwörtern"].forEach((a, i) => {
    num(s, i + 1, 0.9, 3.05 + i * 0.68, NAVY, 0.45);
    txt(s, a, { x: 1.55, y: 3.08 + i * 0.68, w: 10, h: 0.5, fontSize: 20 });
  });
}
// Ü4 Fehler finden
{
  const s = base("Vorwissen aktivieren · Wiederholung Sprachtraining", "Ü1  Was stimmt im Chat nicht?", "Lösung: 1 …, weil ich noch duschen muss. 2 …, dass wir morgen frei haben. 3 Weißt du, ob der Neue auch kommt? 4 Ich freue mich, weil wir frei haben. (Komma fehlt). Reflexion: Beim Sprechen hört man oft „weil ich muss …“ – geschrieben ist das falsch.");
  const msgs = [
    ["l", "Komme später, weil ich muss noch duschen 🙈"],
    ["r", "Trainer hat gesagt, dass wir haben morgen frei!"],
    ["l", "Weißt du, ob kommt der Neue auch?"],
    ["r", "Ich freue mich weil wir frei haben 💪"],
  ];
  box(s, 0.6, 1.55, 7.2, 4.5, "E9EEF6");
  msgs.forEach(([side, m], i) => {
    const y = 1.75 + i * 1.05;
    const x = side === "l" ? 0.85 : 2.35;
    box(s, x, y, 5.2, 0.8, side === "l" ? WHITE : "DCF3E4", { r: 0.18 });
    txt(s, (i + 1) + "  " + m, { x: x + 0.2, y: y + 0.17, w: 4.9, h: 0.5, fontSize: 17 });
  });
  [["1", "Finde in jeder Nachricht den Fehler."], ["2", "Schreibe den Satz richtig ins Heft."], ["3", "Notiere: Verbstellung oder Komma?"]].forEach(([n, x], i) => {
    num(s, n, 8.1, 1.8 + i * 1.0, VIO, 0.45);
    txt(s, x, { x: 8.7, y: 1.8 + i * 1.0, w: 4.0, h: 0.9, fontSize: 18, bold: true });
  });
  zeit(s, "4 Minuten · allein, dann Lösung");
}
// Ü4 Lösung
{
  const s = base("Vorwissen aktivieren · Ü1 Lösung", "Ü1  So ist es richtig", "Vergleichen und mit Grün verbessern. Kurz reflektieren: Beim Sprechen hört man oft „weil ich muss …“ – geschrieben ist das falsch.");
  [[[["Komme später, ", ""], ["weil", "sub"], [" ich noch duschen ", ""], ["muss", "fin"], [".", ""]], "Verbstellung"],
   [[["Trainer hat gesagt, ", ""], ["dass", "sub"], [" wir morgen frei ", ""], ["haben", "fin"], ["!", ""]], "Verbstellung"],
   [[["Weißt du, ", ""], ["ob", "sub"], [" der Neue auch ", ""], ["kommt", "fin"], ["?", ""]], "Verbstellung"],
   [[["Ich freue mich", ""], [",", "b"], [" ", ""], ["weil", "sub"], [" wir frei ", ""], ["haben", "fin"], [".", ""]], "Komma"]].forEach(([r, typ], i) => {
    const y = 1.6 + i * 1.05;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    txt(s, String(i + 1), { x: 0.8, y: y + 0.22, w: 0.5, h: 0.4, fontSize: 18, bold: true, color: VIO });
    txt(s, runs(r, 22), { x: 1.35, y: y + 0.2, w: 8.8, h: 0.5, fontFace: HEAD });
    box(s, 10.4, y + 0.2, 2.1, 0.45, typ === "Komma" ? YEL : VIO, { r: 0.22 });
    txt(s, typ, { x: 10.4, y: y + 0.2, w: 2.1, h: 0.45, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}

// Ü2 Belegen
uebung("Vorwissen aktivieren · Wiederholung Belegen", "Ü2  Wo steht das?", [
  "Wo steht, dass Paddy ins Bett kriecht?",
  "Wo bittet Schaya den Fremden, leise zu sein?",
  "Gib jeweils Strophe, Vers und Zitat an.",
], "3 Minuten · allein im Heft", "Wiederholung vom Mittwoch: Beleg = Str. + V. + Zitat in Anführungszeichen. Lösung auf der nächsten Folie.");
{
  const s = base("Vorwissen aktivieren · Ü2 Lösung", "Ü2  Lösung", "Mit Grün verbessern. Auf Anführungszeichen und die Angabe von Strophe UND Vers achten.");
  [["1", "Str. 5, V. 21: „Ins Bette Paddy Fingal kroch“"], ["2", "Str. 7, V. 31: „Tut leise, Fremder, und tretet sacht“"]].forEach(([n, x], i) => {
    const y = 2.0 + i * 1.3;
    box(s, 0.6, y, 12.1, 1.05, LIGHT);
    num(s, n, 0.9, y + 0.27, GREEN, 0.5);
    txt(s, x, { x: 1.7, y: y + 0.3, w: 10.8, h: 0.5, fontSize: 24, fontFace: HEAD });
  });
  txt(s, "Vergleiche und verbessere mit Grün.", { x: 0.6, y: 5.2, w: 12.1, h: 0.5, fontSize: 20, bold: true, align: "center" });
}
// Reime entdecken (2 Folien)
const STR1 = [["Paddy Fingal war von Riesen", "art", ":", "a"], ["Sechs Ellen flog sein schwarzer B", "art", ",", "a"], ["seine Fäuste waren wie Tonnen gr", "oß", ",", "b"], ["einen Mastbaum führt’ er als Wurfgesch", "oss", " –", "b"], ["     so ein Kerl war Paddy Fingal!", "", "", "x"]];
[false, true].forEach((show) => {
  const s = base("Informieren · Reime entdecken", show ? "Strophe 1: aabb + Kehrreim" : "Welche Wörter klingen gleich?", show
    ? "Erst nach dem Gespräch zeigen. Reim = gleicher Klang ab dem letzten betonten Vokal (-art, -oss). Gleiche Reime bekommen denselben Buchstaben. Vers 5 reimt sich nicht (x), kommt aber in jeder Strophe wieder: Kehrreim. aabb heißt Paarreim. Überleitung: Schnipsel austeilen, Merkkasten gemeinsam lesen."
    : "Strophe 1 im Chor lesen (Methodenkarte „Chorisches Lesen“). Dann fragen: Hört auf die Enden der Verse – welche klingen gleich? Ab wo klingen Bart und Riesenart gleich?");
  STR1.forEach(([a, r, e, l], i) => {
    const y = 1.75 + i * 0.75;
    box(s, 0.6, y, 9.6, 0.6, LIGHT);
    txt(s, String(i + 1), { x: 0.7, y: y + 0.12, w: 0.4, h: 0.4, fontSize: 15, color: MUTED, align: "right" });
    txt(s, [{ text: a }, { text: r, options: { bold: show, color: show && l !== "x" ? RC[l] : NAVY } }, { text: e }], { x: 1.3, y: y + 0.1, w: 8.8, h: 0.45, fontSize: 22, fontFace: HEAD });
    if (show) {
      box(s, 10.45, y + 0.05, 0.9, 0.5, RC[l], { r: 0.25 });
      txt(s, l, { x: 10.45, y: y + 0.05, w: 0.9, h: 0.5, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle" });
    }
  });
  if (show) {
    box(s, 0.6, 5.65, 12.1, 0.9, LIGHT2);
    txt(s, [{ text: "aabb = Paarreim", options: { bold: true, color: NAVY } }, { text: "   ·   " }, { text: "x = Vers ohne Reim, kehrt in jeder Strophe wieder: ", options: {} }, { text: "Kehrreim (Refrain)", options: { bold: true } }], { x: 0.9, y: 5.85, w: 11.6, h: 0.5, fontSize: 19 });
  } else {
    txt(s, "Hört auf die Enden der Verse. Welche klingen gleich – und ab welchem Laut?", { x: 0.6, y: 5.75, w: 12.1, h: 0.6, fontSize: 21, bold: true });
  }
});
// Merkkasten
{
  const s = base("Merkkasten – Schnipsel einkleben", "Reim und Reimschema", "Schnipsel austeilen, einkleben lassen, gemeinsam lesen. Kontrollfragen: Wie heißt das Muster, wenn sich Vers 1 und 3 reimen? (abab, Kreuzreim) – und Vers 1 und 4? (abba, umarmender Reim).");
  box(s, 0.6, 1.55, 12.1, 2.1, WHITE, { line: RED, lw: 3, r: 0.05 });
  txt(s, [{ text: "Reim: ", options: { bold: true, color: RED } }, { text: "Zwei Wörter reimen sich, wenn sie " }, { text: "ab dem letzten betonten Vokal gleich klingen", options: { bold: true } }, { text: ": Schreck – Versteck, daher – sehr.\n" }, { text: "Reimschema: ", options: { bold: true, color: RED } }, { text: "Versenden, die sich reimen, bekommen den gleichen Buchstaben (a, b, c …). Ein Vers ohne Reim bekommt ein x." }], { x: 0.9, y: 1.7, w: 11.5, h: 1.85, fontSize: 20 });
  [["Paarreim", "aabb"], ["Kreuzreim", "abab"], ["umarmender Reim", "abba"]].forEach(([n, l], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 3.85, 3.9, 1.15, LIGHT);
    txt(s, n, { x, y: 3.92, w: 3.9, h: 0.45, fontSize: 19, bold: true, align: "center" });
    txt(s, l.split("").map((c, k) => ({ text: c + (k < 3 ? "  " : ""), options: { bold: true, color: RC[c] } })), { x, y: 4.4, w: 3.9, h: 0.5, fontSize: 24, align: "center" });
  });
  box(s, 0.6, 5.2, 5.95, 0.95, LIGHT2);
  txt(s, [{ text: "Wofür? ", options: { bold: true } }, { text: "Reim und Reimschema sind Merkmale der Lyrik – du brauchst sie für die Analyse einer Ballade." }], { x: 0.8, y: 5.28, w: 5.6, h: 0.85, fontSize: 14 });
  box(s, 6.75, 5.2, 5.95, 0.95, LIGHT2);
  txt(s, [{ text: "Test: ", options: { bold: true } }, { text: "Lies die Versenden laut. Klingen sie ab dem betonten Vokal gleich? Dann bekommen sie denselben Buchstaben." }], { x: 6.95, y: 5.28, w: 5.6, h: 0.85, fontSize: 14 });
}
// Ü3 Aufgabe
uebung("Verarbeiten · Schnipsel", "Ü3  Reime erkennen", [
  "Lies die Versenden halblaut.",
  "Schreibe hinter jeden Vers den Buchstaben (a, b … oder x).",
  "Nenne die Reimart: Paarreim, Kreuzreim oder umarmender Reim.",
], "6 Minuten · allein · Nr. 1 Paddy Fingal, Nr. 2–4 aus deiner Sportart", "Drei Fassungen des Schnipsels: Gewichtheben, Judo, Handball – jeder bekommt seine Sportart. Aufbau überall gleich: 1 Paddy Str. 4, 2 Kreuzreim, 3 umarmender Reim, 4 Paarreim. Lösungen auf den Aufdeck-Folien (je Nummer alle drei Sportarten). Wer fertig ist: * bei Nr. 1 den Kehrreim bestimmen.");
// Ü3 Aufdeck-Folien: Nr. 1 gemeinsam, Nr. 2–4 je Nummer alle drei Sportarten
{
  const s = base("Verarbeiten · Ü3 gemeinsam prüfen", "Ü3  Lösung – 1  Paddy Fingal, Str. 4", "Erst einen Schüler die Buchstaben und die Reimart nennen lassen, dann diese Folie zeigen. Alle verbessern mit Grün. Der Kehrreim reimt sich nicht (x), wiederholt sich aber in jeder Strophe.");
  [["… Schreck", "a"], ["… Versteck", "a"], ["… daher", "b"], ["… sehr", "b"], ["* … Paddy Fingal", "x"]].forEach(([v, l], i) => {
    const y = 1.7 + i * 0.72;
    box(s, 0.6, y, 9.6, 0.6, LIGHT);
    txt(s, v, { x: 0.9, y: y + 0.1, w: 9.2, h: 0.45, fontSize: 22, fontFace: HEAD });
    box(s, 10.45, y + 0.05, 0.9, 0.5, RC[l], { r: 0.25 });
    txt(s, l, { x: 10.45, y: y + 0.05, w: 0.9, h: 0.5, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  box(s, 0.6, 5.55, 12.1, 0.85, LIGHT2);
  txt(s, [{ text: "Reimart: ", options: { bold: true, color: RED } }, { text: "Paarreim (aabb) + Kehrreim (x)", options: { bold: true } }], { x: 0.9, y: 5.72, w: 11.6, h: 0.55, fontSize: 24 });
}
const SP3 = {
  2: ["Kreuzreim (abab)", "abab", [["Gewichtheben", ["Die Halle ist noch kalt und leer,", "die Hantel liegt schon da.", "Der erste Versuch fällt mir schwer,", "doch dann ruft jemand: „Ja!“"]], ["Judo", ["Wir grüßen uns, der Kampf beginnt,", "am Ärmel spür ich seinen Griff.", "Ich weiß genau, wer heut gewinnt:", "wer schneller ist mit seinem Kniff."]], ["Handball", ["Der Pass kommt flach, ich bin jetzt dran,", "die Abwehr steht ganz dicht.", "Ich täusche links, der Gegner kann", "mich halten – heute nicht!"]]]],
  3: ["umarmender Reim (abba)", "abba", [["Gewichtheben", ["Der letzte Versuch, die Halle schweigt,", "ich atme tief, ich bin bereit,", "der Trainer nickt, jetzt ist es Zeit –", "die Stange über meinem Kopf, sie steigt."]], ["Judo", ["Der Kampf beginnt, ich greife zu,", "mein Gegner hält mich am Kragen,", "ich muss jetzt alles wagen –", "ein Wurf, und dann ist Ruh."]], ["Handball", ["Die Halle bebt beim letzten Spiel,", "wir liegen hinten, nur ein Tor.", "Der Trainer ruft: „Jetzt alle vor!“", "Wir kämpfen bis zum Schluss ums Ziel."]]]],
  4: ["Paarreim (aabb)", "aabb", [["Gewichtheben", ["Ich kreide meine Hände ein,", "heut soll der Rekord der meine sein.", "Ich reiße hoch und steh ganz still –", "drei weiße Lampen, genau, was ich will!"]], ["Judo", ["Ich binde meinen Gürtel fest,", "jetzt kommt der Kampf, jetzt kommt der Test.", "Ich greife, drehe, ziehe an –", "Ippon! Gewonnen hab ich dann."]], ["Handball", ["Zehn Sekunden noch, der Ball ist frei,", "ich zieh nach innen, schnell vorbei,", "ich steige hoch, der Wurf geht rein –", "die ganze Halle fängt an zu schrein!"]]]],
};
[2, 3, 4].forEach((n) => {
  const [art, letters, cols] = SP3[n];
  const s = base("Verarbeiten · Ü3 gemeinsam prüfen", `Ü3  Lösung – Nr. ${n}: ${art}`, `Je einen Schüler aus jeder Sportart Buchstaben und Reimart nennen lassen, dann zeigen. Alle verbessern mit Grün. Gemeinsam festhalten: Nr. ${n} ist in allen drei Fassungen ein ${art}.` + (n === 3 ? " Typische Hürde: Vers 1 und 4 vergleichen." : ""));
  cols.forEach(([sp, vs], c) => {
    const x = 0.6 + c * 4.1;
    box(s, x, 1.55, 3.95, 0.45, NAVY, { r: 0.05 });
    txt(s, sp, { x, y: 1.55, w: 3.95, h: 0.45, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    vs.forEach((v, i) => {
      const y = 2.15 + i * 0.78;
      box(s, x, y, 3.95, 0.68, LIGHT, { r: 0.05 });
      txt(s, v, { x: x + 0.1, y: y + 0.08, w: 3.25, h: 0.55, fontSize: 12.5, fontFace: HEAD, valign: "middle" });
      const l = letters[i];
      box(s, x + 3.42, y + 0.14, 0.45, 0.4, RC[l], { r: 0.2 });
      txt(s, l, { x: x + 3.42, y: y + 0.14, w: 0.45, h: 0.4, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
  });
  box(s, 0.6, 5.45, 12.1, 0.85, LIGHT2);
  txt(s, [{ text: "Reimart: ", options: { bold: true, color: RED } }, { text: art, options: { bold: true } }, { text: "  – in allen drei Fassungen" }], { x: 0.9, y: 5.62, w: 11.6, h: 0.55, fontSize: 24 });
});
// Zielcheck
{
  const s = base("Auswerten", "Zielcheck: Hast du das Ziel erreicht?", "Allein im Heft, 4 Minuten, dann Lösungsfolie. Herumgehen: Wer bei a noch Fehler hat, wird am Freitag in der Wiederholung gezielt drangenommen.");
  box(s, 0.6, 1.6, 12.1, 3.05, LIGHT);
  txt(s, [{ text: "a  ", options: { bold: true, color: RED } }, { text: "Bestimme das Reimschema von Strophe 6 (V. 26–30) und nenne die Reimart. Welcher Vers ist der Kehrreim?" }], { x: 0.9, y: 1.75, w: 11.6, h: 0.75, fontSize: 19 });
  ["Indem so schob der Schotte herein,", "an den Balken rührte sein Scheitelbein,", "und er schnob und wischte sich den Schweiß", "und rollte die Augen wild im Kreis:", "     „Wo steckt der Paddy Fingal?“"].forEach((v, i) => txt(s, v, { x: 1.4, y: 2.5 + i * 0.4, w: 10.5, h: 0.38, fontSize: 18, fontFace: HEAD }));
  box(s, 0.6, 4.85, 12.1, 1.15, LIGHT);
  txt(s, [{ text: "b  ", options: { bold: true, color: RED } }, { text: "Welche Reimart ist das?   " }, { text: "Ball – Tor – Fall – vor", options: { bold: true, fontFace: HEAD } }], { x: 0.9, y: 5.15, w: 11.6, h: 0.6, fontSize: 21 });
  zeit(s, "4 Minuten · allein im Heft");
}
{
  const s = base("Auswerten · Selbstkontrolle", "Zielcheck – Lösung", "Mit Grün verbessern. Alles richtig? Dann ist das Ziel erreicht.");
  box(s, 0.6, 1.7, 12.1, 1.6, LIGHT);
  txt(s, [{ text: "a  ", options: { bold: true, color: RED } }, { text: "herein – Scheitelbein (a a) · Schweiß – Kreis (b b) · „Wo steckt der Paddy Fingal?“ (x)\n" }, { text: "Paarreim (aabb) – V. 30 ist der Kehrreim", options: { bold: true } }], { x: 0.9, y: 1.9, w: 11.6, h: 1.3, fontSize: 21 });
  box(s, 0.6, 3.55, 12.1, 1.0, LIGHT);
  txt(s, [{ text: "b  ", options: { bold: true, color: RED } }, { text: "Ball – Fall (a a), Tor – vor (b b) im Wechsel: abab = " }, { text: "Kreuzreim", options: { bold: true } }], { x: 0.9, y: 3.8, w: 11.6, h: 0.6, fontSize: 21 });
  txt(s, "Verbessere mit Grün. Alles richtig? Dann hast du das Ziel erreicht.", { x: 0.6, y: 5.1, w: 12.1, h: 0.5, fontSize: 20, bold: true, align: "center" });
}
// * Ü3 nacherzählen mit Stichwörtern
{
  const s = base("* Puffer · Nacherzählen vom Mittwoch", "* Erzählt die Ballade nach", "Puffer: Nur, wenn die Stunde vorher fertig ist (oder falls Ü3 am Mittwoch nicht geschafft wurde). Die Stichwörter geben die Reihenfolge vor – keine fertigen Sätze. A erzählt Abschnitt 1–4, B Abschnitt 5–8, der andere hakt die Checkliste ab. Die Verknüpfungswörter (als, weil) knüpfen an Teil 2 an.");
  box(s, 0.6, 1.5, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Einstieg: ", options: { bold: true, color: GREEN } }, { text: "In der Ballade „Paddy Fingal“ von Wilhelm Brandes geht es um …" }], { x: 0.85, y: 1.6, w: 11.6, h: 0.4, fontSize: 17 });
  const AB = [
    ["1", "Paddy – Riese – Bart – Mastbaum", "Str. 1"], ["2", "Hüne aus Schottland – Wer ist stärker?", "Str. 2"],
    ["3", "Ufer – zehn Schuh größer – Schreck – Versteck", "Str. 3–4"], ["4", "Schaya – Bett – Kissen – nur die Nase", "Str. 5"],
    ["5", "Schotte sucht – Schayas Warnung – „Kindlein“", "Str. 6–7"], ["6", "Nase – Baby? – riesiger Vater", "Str. 8–9"],
    ["7", "Flucht – Sund – fast ertrunken", "Str. 10"], ["8", "Siegesgeschrei – Nachbarn – „tapfer“", "Str. 11"],
  ];
  AB.forEach(([n, w, st], i) => {
    const col = i < 4 ? 0 : 1, r = i % 4;
    const x = 0.6 + col * 6.15, y = 2.2 + r * 0.72;
    box(s, x, y, 5.95, 0.62, LIGHT);
    txt(s, n, { x: x + 0.12, y: y + 0.13, w: 0.35, h: 0.38, fontSize: 16, bold: true, color: GREEN });
    txt(s, w, { x: x + 0.5, y: y + 0.14, w: 4.4, h: 0.38, fontSize: 15 });
    txt(s, st, { x: x + 4.85, y: y + 0.15, w: 1.0, h: 0.35, fontSize: 12, color: MUTED, align: "right" });
  });
  txt(s, [{ text: "A: ", options: { bold: true } }, { text: "Abschnitt 1–4  ·  " }, { text: "B: ", options: { bold: true } }, { text: "Abschnitt 5–8  ·  " }, { text: "Verknüpfe: ", options: { bold: true, color: GREEN } }, { text: "zuerst · dann · als … · weil … · daraufhin · schließlich · am Ende" }], { x: 0.6, y: 5.15, w: 12.1, h: 0.4, fontSize: 15 });
  txt(s, [{ text: "Checkliste: ", options: { bold: true, color: GREEN } }, { text: "Titel und Autor · richtige Reihenfolge · nichts dazuerfinden · alle wichtigen Details · * spannend mit wörtlicher Rede" }], { x: 0.6, y: 5.6, w: 12.1, h: 0.4, fontSize: 14, color: MUTED });
  zeit(s, "7 Minuten · Tandem");
}


pres.writeFile({ fileName: "Donnerstag_05-11_Deutsch_7A.pptx" });
