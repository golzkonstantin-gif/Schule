const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Mi 18.11. · Sachtexte verstehen";
const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4", VIO = "7A5BA6", TEAL_ = "2A8C8C";
const HEAD = "Cambria", BODY = "Calibri";
let FOOT = "Mittwoch, 4.11. · Deutsch 7A";
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

// ================= Mi 18.11. · Doppelstunde (ca. 75 Min.) · AVIVA =================
// Einführung Sachtexte: LB S. 178–179 „Die Rauchzeichen der Native Americans“ (Text nicht abgedruckt, nur Zeilenverweise)
// Teil 1: Merkmale eines Sachtexts · Teil 2: Sprachtraining Zeitformen im Sachtext + Wiederholung
FOOT = "Mittwoch, 18.11. · Deutsch 7A";
const Z1 = "Ich kann an einem Text zeigen, woran ich einen Sachtext erkenne, und Merkmale mit Zeilenangabe belegen.";
const Z2 = "Ich kann Zeitformen in einem Sachtext bestimmen und erklären, warum sie verwendet werden.";
const zeit = (s, meta, y = 6.25, end = "Fertig? Stift hinlegen.") => {
  box(s, 0.6, y, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: end, options: { italic: true, color: MUTED } }], { x: 0.85, y: y + 0.08, w: 11.7, h: 0.4, fontSize: 16 });
};
const vr = (parts, size = 22) => parts.map(([x, k]) => ({ text: x, options: Object.assign({ fontSize: size }, k === "hv" ? { bold: true, color: RED } : k === "p2" ? { bold: true, color: YEL } : k === "fin" ? { bold: true, underline: { style: "sng" } } : k === "m" ? { color: MUTED } : {}) }));

// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "MITTWOCH, 18.11. · DEUTSCH 7A · NEUES THEMA", { x: 0.9, y: 2.0, w: 11, h: 0.4, fontSize: 14, color: ICE, charSpacing: 2 });
  txt(s, "Sachtexte verstehen", { x: 0.9, y: 2.5, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 46, bold: true, color: WHITE });
  txt(s, "Lesen – verstehen – zusammenfassen", { x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 22, italic: true, color: ICE });
}

// 2 Überblick
{
  const s = base("Ankommen", "Ziele und Ablauf", "Ansage: „Stifte liegen. Ihr hört nur zu. Mit John Maynard sind wir bei den Balladen fertig. Heute beginnt ein neues Thema: Sachtexte. In vier Wochen schreibt ihr dazu eine Klassenarbeit: Ihr entnehmt Informationen aus einem Sachtext und fasst sie zusammen. Heute klären wir, woran man einen Sachtext erkennt. Im zweiten Teil gibt es wie jeden Mittwoch Sprachtraining.“");
  [["1", "Was ist ein Sachtext?", "ca. 50 Min.", Z1, NAVY], ["2", "Sprachtraining: Zeitformen im Sachtext", "ca. 17 Min.", Z2, RED]].forEach(([n, ti, m, z, col], i) => {
    const y = 1.6 + i * 1.9;
    box(s, 0.6, y, 12.1, 1.7, WHITE, { line: col, lw: 2.5 });
    num(s, n, 0.85, y + 0.2, col, 0.55);
    txt(s, ti, { x: 1.6, y: y + 0.2, w: 8.6, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true });
    txt(s, m, { x: 10.3, y: y + 0.25, w: 2.2, h: 0.45, fontSize: 17, color: MUTED, align: "right" });
    txt(s, [{ text: "Ziel: ", options: { bold: true, color: RED } }, { text: z, options: { italic: true } }], { x: 1.6, y: y + 0.82, w: 10.8, h: 0.7, fontSize: 18 });
  });
  txt(s, "Teil 1:  Beispiele · LB 1 Vor dem Lesen · LB 2 Beim Lesen · Merkmale · Ü1 Merkmale belegen      Teil 2:  Ü2 Zeitformen bestimmen · * Ü3 Wiederholung · Zielcheck", { x: 0.6, y: 5.55, w: 12.1, h: 0.6, fontSize: 15, color: MUTED });
}

// 3 Leselotse – der Weg der Einheit
{
  const s = base("Ankommen · Überblick über die Einheit", "Der Weg zur Klassenarbeit: lesen wie ein Lotse", "Kurz zeigen: Die drei Schritte begleiten die ganze Einheit (Lesestrategien des Leselotsen). Heute: vor und beim Lesen. Nächste Stunde: genau lesen, Notizen, zusammenfassen (LB 3 und 4).");
  [["Vor dem Lesen", "Überblick verschaffen: Überschrift, Bilder, Teilüberschriften · Vermutungen · Leseziel", NAVY], ["Beim Lesen", "zügig lesen · dann genau lesen · Wichtiges markieren · Notizen mit Zeilenangabe", TEAL_], ["Nach dem Lesen", "Informationen ordnen · zusammenfassen · in eigenen Worten wiedergeben", RED]].forEach(([h, x, col], i) => {
    const xx = 0.6 + i * 4.1;
    box(s, xx, 1.7, 3.85, 3.4, WHITE, { line: col, lw: 2.5 });
    num(s, i + 1, xx + 0.25, 1.9, col, 0.55);
    txt(s, h, { x: xx + 0.95, y: 1.95, w: 2.8, h: 0.5, fontFace: HEAD, fontSize: 22, bold: true, color: col });
    txt(s, x, { x: xx + 0.25, y: 2.75, w: 3.4, h: 2.2, fontSize: 17 });
    if (i < 2) txt(s, "→", { x: xx + 3.82, y: 3.1, w: 0.3, h: 0.5, fontSize: 24, bold: true, color: MUTED });
  });
  box(s, 0.6, 5.4, 12.1, 0.7, LIGHT2);
  txt(s, [{ text: "Ziel der Einheit: ", options: { bold: true, color: RED } }, { text: "Klassenarbeit – Informationen aus einem Sachtext entnehmen und schriftlich zusammenfassen." }], { x: 0.85, y: 5.52, w: 11.7, h: 0.45, fontSize: 18 });
}

// 4 Vorwissen: Beispiele
const BSP = [
  ["A", "„Die Handball-Weltmeisterschaft der Männer 2027 findet in Deutschland statt.“", "Nachricht"],
  ["B", "„Judo (japanisch für ‚sanfter Weg‘) ist eine Kampfsportart aus Japan. Seit 1964 ist Judo olympisch.“", "Lexikonartikel"],
  ["C", "„Beim Reißen wird die Hantel in einer einzigen Bewegung vom Boden über den Kopf gehoben.“", "Fachtext (Gewichtheben)"],
  ["D", "„Aufwärmen: 10 Minuten lockeres Laufen, danach 5 Minuten Mobilisation.“", "Trainingsplan / Anleitung"],
  ["E", "Empfohlene Schlafdauer für 13- bis 17-Jährige: 8–10 Stunden pro Nacht", "Tabelle / Grafik"],
];
[false, true].forEach((show) => {
  const s = base("Vorwissen aktivieren", show ? "Worauf liegt der Fokus? – Lösung" : "Worauf liegt bei diesen Texten der Fokus?", show
    ? "Erwartung: Alle Texte wollen informieren – über Tatsachen, die man überprüfen kann. Keine Geschichte, keine Meinung. Das sind Sachtexte. Textsorten nennen (Nachricht, Lexikon, Fachtext, Anleitung, Grafik). Wo lest ihr so etwas? (Lehrbuch, Nachrichten-App, Trainingsplan, Beipackzettel …)"
    : "Ansage: „Lest die fünf Texte. Überlegt zu zweit: Was wollen alle diese Texte? Worauf liegt der Fokus? Zwei Minuten.“ Dann sammeln.");
  BSP.forEach(([l, x, typ], i) => {
    const y = 1.55 + i * 0.82;
    box(s, 0.6, y, 12.1, 0.7, i % 2 ? WHITE : LIGHT, { line: "C9D3EA" });
    txt(s, l, { x: 0.8, y: y + 0.15, w: 0.5, h: 0.4, fontSize: 20, bold: true, color: VIO });
    txt(s, x, { x: 1.35, y: y + 0.13, w: show ? 8.3 : 11.2, h: 0.45, fontSize: 16, italic: i !== 4 });
    if (show) txt(s, typ, { x: 9.8, y: y + 0.15, w: 2.8, h: 0.4, fontSize: 14, bold: true, color: TEAL_, align: "right" });
  });
  if (show) {
    box(s, 0.6, 5.75, 12.1, 0.65, WHITE, { line: RED, lw: 2.5 });
    txt(s, [{ text: "Fokus: ", options: { bold: true, color: RED } }, { text: "informieren – über Tatsachen, sachlich, ohne Geschichte und ohne Meinung. Das sind Sachtexte." }], { x: 0.85, y: 5.86, w: 11.7, h: 0.45, fontSize: 18 });
  } else zeit(s, "2 Minuten · zu zweit", 5.85, "Dann sammeln wir.");
});

// 6 LB 1 Vor dem Lesen
uebung("Teil 1 · Informieren · Vor dem Lesen", "LB 1  Verschafft euch einen Überblick (LB S. 178)", [
  "Lies nur die Überschrift und die einführende Information (kursiv).",
  "Sieh dir die Abbildung an. Überfliege die drei Teilüberschriften und die ersten drei Textzeilen.",
  "Besprich mit deinem Partner: Worum könnte es gehen? Was wisst ihr schon über das Thema?",
], "6 Minuten · erst allein, dann zu zweit", "Lesetechnik: überfliegendes bzw. orientierendes Lesen (bekannt aus der Balladen-Einheit). Danach 2–3 Vermutungen sammeln, an die Seitentafel schreiben – nach dem Lesen prüfen.");

// 7 LB 2 Beim Lesen
uebung("Teil 1 · Verarbeiten · Beim Lesen", "LB 2  Lest den Text zügig (LB S. 178–179)", [
  "Lies den ganzen Text zügig durch.",
  "Erkläre deinem Partner, was du über die Verständigung der Native Americans erfahren hast.",
  "Wechselt: Der Partner ergänzt, was noch fehlt.",
], "10 Minuten · allein lesen, dann Tandem", "Methodenkarte: Tandem (Erklären im Wechsel). Danach kurz: Stimmen unsere Vermutungen von vorhin? Worterklärungen stehen unten im LB (Fußnoten 1 und 2).");

// 8 Merkmale entdecken
uebung("Teil 1 · Informieren · Merkmale entdecken", "Woran merkt ihr, dass das ein Sachtext ist?", [
  "Sieh dir den Text noch einmal an: Aufbau, Inhalt, Sprache.",
  "Notiere zu zweit mindestens vier Merkmale.",
  "Schreibe zu jedem Merkmal eine Zeile aus dem Text dazu.",
], "6 Minuten · zu zweit · Heft", "Danach sammeln, dann Merksatz-Folie aufdecken und Schnipsel austeilen. Hilfsfragen, falls nötig: Wie ist der Text gegliedert? Was für Wörter stehen in den Fußnoten? Gibt es eine Meinung oder ein „ich“?");

// 9 Merksatz Sachtext
{
  const s = base("Teil 1 · Merksatz", "Merksatz: Merkmale eines Sachtexts", "Schnipsel austeilen und einkleben lassen (nicht abschreiben). Mit den gesammelten Merkmalen der Klasse vergleichen.");
  box(s, 0.6, 1.5, 12.1, 0.95, WHITE, { line: RED, lw: 3 });
  txt(s, "Ein Sachtext informiert über ein Thema. Er stellt Tatsachen sachlich dar.", { x: 0.9, y: 1.62, w: 11.6, h: 0.7, fontSize: 22, bold: true, valign: "middle" });
  const M = [["Überschrift", "(oft mit einführender Information) nennt das Thema"], ["Teilüberschriften und Absätze", "gliedern den Text"], ["Fakten", "Zahlen, Daten, Namen, die man überprüfen kann"], ["Fachbegriffe", "werden oft erklärt (Fußnote, „also …“)"], ["Beispiele und Erklärungen", "machen die Informationen verständlich"], ["sachliche Sprache", "keine Meinung, kein „ich“, keine Spannung"]];
  M.forEach(([a, b], i) => {
    const col = i % 2, r = Math.floor(i / 2);
    const x = 0.6 + col * 6.1, y = 2.65 + r * 0.72;
    num(s, i + 1, x, y + 0.05, NAVY, 0.42);
    txt(s, [{ text: a + " ", options: { bold: true } }, { text: b, options: { color: MUTED } }], { x: x + 0.55, y: y + 0.03, w: 5.4, h: 0.6, fontSize: 16 });
  });
  box(s, 0.6, 4.95, 5.95, 1.35, LIGHT);
  txt(s, [{ text: "Wofür?\n", options: { bold: true, color: RED, fontSize: 17 } }, { text: "Du erkennst, aus welchem Text du Informationen entnehmen kannst – für Referate und die Klassenarbeit.", options: { fontSize: 15 } }], { x: 0.8, y: 5.05, w: 5.6, h: 1.2 });
  box(s, 6.75, 4.95, 5.95, 1.35, LIGHT);
  txt(s, [{ text: "Test\n", options: { bold: true, color: RED, fontSize: 17 } }, { text: "Will der Text informieren – nicht unterhalten oder überzeugen? Kann man die Angaben überprüfen?", options: { fontSize: 15 } }], { x: 6.95, y: 5.05, w: 5.6, h: 1.2 });
}

// 10 Ü1 Merkmale belegen
const UE1 = [
  ["Überschrift", "„Die Rauchzeichen der Native Americans“ (mit einführender Information)"],
  ["Teilüberschriften", "Z. 1–2 „Nachrichtenübermittlung vor 400 Jahren“ · Z. 30–31 „Erzeugung von Rauch …“ · Z. 46 „Bedeutung der Zeichen“"],
  ["Fakten", "Z. 3–5: „vor ungefähr 400 Jahren“, Apachen, Hopis, Navajos, Sioux · Z. 49–53: eine, zwei, drei Wolken"],
  ["Fachbegriff mit Erklärung", "Z. 5: „indigenen Stämmen“ (Fußnote 1) · Z. 56–57: „allgemeingültigen Code“ (Fußnote 2)"],
  ["Beispiel", "Z. 20–22: „zum Beispiel, wenn Feinde im Anmarsch waren oder ein Fluss Hochwasser führte“"],
  ["sachliche Sprache", "z. B. Z. 15–17: „Rauchzeichen dienten dazu, Nachrichten über große Entfernungen weiterzugeben“ – keine Meinung, kein „ich“"],
];
uebung("Teil 1 · Verarbeiten", "Ü1  Merkmale mit Zeilenangabe belegen", [
  "Zeichne eine Tabelle ins Heft: Merkmal | Beispiel aus dem Text | Zeile.",
  "Trage die sechs Merkmale vom Merksatz ein.",
  "Finde zu jedem Merkmal ein Beispiel im Text und notiere die Zeile (Z.).",
], "10 Minuten · allein · Heft", "Zeilenzählung im LB: Z. 1 = erste Teilüberschrift; die Überschrift und die einführende Information haben keine Zeilennummer. Lösung auf der nächsten Folie.");
{
  const s = base("Teil 1 · Ü1 Lösung", "Ü1  Lösung – Vergleiche und verbessere mit Grün", "Andere passende Stellen gelten. Bei „sachliche Sprache“ nachfragen: Woran merkt man das? (keine Meinung, kein ich, keine Spannung). Kurz zurück zu den Vermutungen: Was hat gestimmt?");
  UE1.forEach(([m, b], i) => {
    const y = 1.5 + i * 0.78;
    box(s, 0.6, y, 12.1, 0.68, i % 2 ? WHITE : LIGHT, { line: "C9D3EA" });
    num(s, i + 1, 0.75, y + 0.13, NAVY, 0.42);
    txt(s, m, { x: 1.3, y: y + 0.15, w: 2.9, h: 0.4, fontSize: 16, bold: true });
    txt(s, b, { x: 4.2, y: y + 0.08, w: 8.4, h: 0.55, fontSize: 14 });
  });
}

// Teil 2 Sprachtraining
section("2", "Sprachtraining: Zeitformen im Sachtext", Z2, RED);
const ZF = [
  ["Z. 13", [["Beispielsweise ", ""], ["weiß", "fin"], [" man, …", ""]], "Präsens", "gilt heute noch / allgemein"],
  ["Z. 13–14", [["… wie sie miteinander kommuniziert ", ""], ["haben", "hv"]], "Perfekt", "abgeschlossen, Ergebnis wichtig"],
  ["Z. 20", [["Sie ", ""], ["warnten", "fin"], [" vor Gefahren.", ""]], "Präteritum", "Vergangenheit (Bericht)"],
  ["Z. 23–24", [["… dass sie Wild ", ""], ["aufgespürt", "p2"], [" ", ""], ["hatten", "hv"]], "Plusquamperfekt", "Vorvergangenheit: vor dem Melden"],
];
[false, true].forEach((show) => {
  const s = base("Teil 2 · Sprachtraining", show ? "Ü2  Lösung" : "Ü2  Welche Zeitform steht hier – und warum?", show
    ? "Mit Grün verbessern. Farbcode: Hilfsverb rot, Partizip II gelb. Nachfragen: Warum erzählt der Text überwiegend im Präteritum? (Es geht um früher.) Warum steht „weiß man“ im Präsens? (Das wissen wir heute.)"
    : "Ansage: „Jetzt arbeitet ihr allein. Bestimmt die Zeitform der markierten Verben und überlegt, warum sie hier steht. Fünf Minuten.“");
  if (!show) [["1", "Bestimme die Zeitform des Verbs in jeder Zeile."], ["2", "Begründe in Stichworten: Warum steht hier diese Zeitform?"]].forEach(([n, x], i) => {
    num(s, n, 0.6, 1.5 + i * 0.55, RED, 0.42);
    txt(s, x, { x: 1.2, y: 1.52 + i * 0.55, w: 11.4, h: 0.45, fontSize: 18, bold: true });
  });
  ZF.forEach(([z, r, tf, why], i) => {
    const y = (show ? 1.55 : 2.75) + i * 0.85;
    box(s, 0.6, y, 12.1, 0.72, LIGHT);
    txt(s, z, { x: 0.8, y: y + 0.17, w: 1.3, h: 0.4, fontSize: 15, bold: true, color: MUTED });
    txt(s, vr(r, 19), { x: 2.1, y: y + 0.15, w: 5.6, h: 0.45, fontFace: HEAD });
    if (show) {
      txt(s, tf, { x: 7.8, y: y + 0.15, w: 2.1, h: 0.45, fontSize: 17, bold: true, color: RED });
      txt(s, why, { x: 9.9, y: y + 0.17, w: 2.7, h: 0.45, fontSize: 14, color: MUTED });
    }
  });
  if (show) {
    box(s, 0.6, 5.1, 12.1, 1.15, WHITE, { line: RED, lw: 3 });
    txt(s, [{ text: "Merke: ", options: { bold: true, color: RED } }, { text: "Sachtexte über früher berichten meist im Präteritum. Was heute noch gilt, steht im Präsens. Was vorher geschah, steht im Plusquamperfekt." }], { x: 0.9, y: 5.2, w: 11.6, h: 0.95, fontSize: 18, valign: "middle" });
  } else zeit(s, "5 Minuten · allein · Heft");
});

// * Ü3 Wiederholung Zeitformen
const UE3 = [
  ["a", "Die Stämme benutzen Rauchzeichen.", "Präteritum", [["Die Stämme ", ""], ["benutzten", "fin"], [" Rauchzeichen.", ""]]],
  ["b", "Der Späher meldet die Gefahr.", "Perfekt", [["Der Späher ", ""], ["hat", "hv"], [" die Gefahr ", ""], ["gemeldet", "p2"], [".", ""]]],
  ["c", "Die Männer decken das Feuer zu.", "Präteritum", [["Die Männer ", ""], ["deckten", "fin"], [" das Feuer zu.", ""]]],
  ["d", "Der Trainer erklärt die Übung.", "Plusquamperfekt", [["Der Trainer ", ""], ["hatte", "hv"], [" die Übung ", ""], ["erklärt", "p2"], [".", ""]]],
];
[false, true].forEach((show) => {
  const s = base("Teil 2 · Sprachwiederholung", show ? "* Ü3  Lösung" : "* Ü3  Setze in die angegebene Zeitform", show
    ? "Mit Grün verbessern. Bei c: trennbares Verb (zudecken) – „zu“ steht am Ende. Bei b und d: Hilfsverb rot, Partizip II gelb."
    : "Wiederholung aus der Verb-Einheit. Wer fertig ist, kontrolliert mit der nächsten Folie.");
  UE3.forEach(([n, x, tf, sol], i) => {
    const y = 1.6 + i * 1.0;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.2, w: 0.5, h: 0.45, fontSize: 20, bold: true, color: VIO });
    txt(s, show ? vr(sol, 20) : x, { x: 1.35, y: y + 0.2, w: 8.2, h: 0.5, fontSize: 20, fontFace: HEAD });
    txt(s, "→ " + tf, { x: 9.6, y: y + 0.22, w: 3.0, h: 0.45, fontSize: 17, bold: true, color: RED, align: "right" });
  });
  if (!show) zeit(s, "6 Minuten · allein · Heft");
});

// Zielcheck
const ZC = [
  ["1", "„Laut JIM-Studie besitzen über 90 Prozent der 12- bis 19-Jährigen ein eigenes Smartphone.“ – Sachtext? Nenne ein Merkmal.", "Ja – Fakt mit Zahl und Quelle, überprüfbar; sachlich, informiert."],
  ["2", "„Ich finde, Handys sollten in der Schule komplett verboten werden!“ – Sachtext? Begründe.", "Nein – Meinung („ich finde“), will überzeugen, nicht informieren."],
  ["3", "„Sie entfachten auf einer sichtbaren Anhöhe ein Feuer.“ (Z. 34–35) – Bestimme die Zeitform.", "Präteritum (entfachten) – Bericht über früher."],
];
[false, true].forEach((show) => {
  const s = base("Auswerten · Zielcheck", show ? "Zielcheck: Lösung – Ziele erreicht?" : "Zielcheck", show
    ? "Mit Grün verbessern. Abschlussfrage: Haben wir die Ziele erreicht? Ausblick: Nächste Stunde lesen wir den Text genau, machen Notizen und fassen ihn zusammen (LB 3 und 4)."
    : "Allein im Heft, 4 Minuten. Herumgehen.");
  ZC.forEach(([n, a, l], i) => {
    const y = 1.55 + i * 1.4;
    box(s, 0.6, y, 12.1, 1.25, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.15, w: 0.5, h: 0.45, fontSize: 20, bold: true, color: VIO });
    txt(s, a, { x: 1.35, y: y + 0.12, w: 11.2, h: 0.6, fontSize: 17 });
    if (show) txt(s, l, { x: 1.35, y: y + 0.72, w: 11.2, h: 0.45, fontSize: 16, bold: true, color: GREEN });
  });
  if (show) txt(s, [{ text: "Ziele erreicht?  ", options: { bold: true, color: RED } }, { text: "Sachtext erkennen und Merkmale belegen · Zeitformen bestimmen und begründen", options: { italic: true } }], { x: 0.6, y: 5.85, w: 12.1, h: 0.45, fontSize: 16 });
  else zeit(s, "4 Minuten · allein im Heft", 5.85);
});

pres.writeFile({ fileName: __dirname + "/../Mittwoch_18-11_Deutsch_7A.pptx" });
