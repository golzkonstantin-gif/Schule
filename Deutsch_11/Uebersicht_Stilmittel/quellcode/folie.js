// Eine Übersichtsfolie: alle bisher bekannten Stilmittel (Deutsch 11, Stand 19.11.: „Gibs auf!“ und „Vor dem Gesetz“) mit Beschreibung und Effekt.
// Stil wie die bisherigen Deutsch-11-Präsentationen (16:9 wide, Navy, Cambria/Calibri), Schrift durchgehend 27 pt.
const pptxgen = require("pptxgenjs");
const OUT = process.argv[2];
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", WHITE = "FFFFFF", RED = "D9534F";
const HEAD = "Cambria", BODY = "Calibri", FS = 27;
const pres = new pptxgen(); pres.layout = "LAYOUT_WIDE"; pres.title = "Übersicht Stilmittel";
const S = [
  ["Parataxe", "Hauptsätze gereiht", "knapp, ruhig, nüchtern"],
  ["Hypotaxe", "Satzgefüge mit Nebensätzen", "verschachtelt, gedrängt"],
  ["Wiederholung", "dasselbe Wort mehrmals", "verstärkt, betont"],
  ["Vergleich", "zwei Bereiche mit „wie“", "anschaulich, bildhaft"],
  ["Metapher", "Vergleich ohne „wie“", "bildhaft, eindringlich"],
  ["Symbol", "Gegenstand für Abstraktes", "Abstraktes wird sichtbar"],
  ["Rhetorische Frage", "Frage, eigentlich Aussage", "provoziert, lenkt Leser"],
  ["Ellipse", "unvollständiger Satz", "knapp, schnell, spannend"],
];
const s = pres.addSlide(); s.background = { color: WHITE };
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 1.0, fill: { color: NAVY }, line: { type: "none" } });
s.addText("Stilmittel – unsere Übersicht", { x: 0.5, y: 0, w: 12.3, h: 1.0, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, valign: "middle", margin: 0 });
const cell = (t, o = {}) => ({ text: t, options: { fontFace: BODY, fontSize: FS, color: o.color || NAVY, bold: o.bold, fill: { color: o.fill || WHITE }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.14] } });
const rows = [
  [cell("Stilmittel", { bold: true, color: WHITE, fill: NAVY }), cell("Beschreibung", { bold: true, color: WHITE, fill: NAVY }), cell("Effekt", { bold: true, color: WHITE, fill: NAVY })],
  ...S.map(([a, b, c], i) => [cell(a, { bold: true, fill: i % 2 ? LIGHT2 : LIGHT }), cell(b, { fill: i % 2 ? LIGHT2 : LIGHT }), cell(c, { fill: i % 2 ? LIGHT2 : LIGHT, color: RED, bold: true })]),
];
s.addTable(rows, { x: 0.4, y: 1.2, w: 12.53, colW: [3.35, 4.6, 4.58], rowH: 0.64, border: { type: "solid", pt: 1, color: WHITE } });
s.slideNumber = { x: 12.6, y: 7.1, w: 0.5, h: 0.3, fontFace: BODY, fontSize: 12, color: MUTED, align: "right" };
s.addNotes("Bisher bekannt (Gibs auf!, Vor dem Gesetz): Parataxe, Hypotaxe, Metapher, Symbol (alt) – Wiederholung, Vergleich, rhetorische Frage, Ellipse (neu).\nTests: Vergleich – steht ein „wie“? · Metapher – lässt sich ein „wie“ einsetzen? · Rhetorische Frage – lässt sie sich in eine Aussage umformen? · Ellipse – muss ich ein Wort ergänzen?\nEffekt immer auf den Text beziehen: „Dadurch wirkt …“ – z. B. Wiederholung „Ich bin mächtig … einer mächtiger als der andere“: Die Macht wirkt endlos gesteigert.");
pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
