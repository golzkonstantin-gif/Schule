// Präsentation zur Kafka-Stunde (16:9): Einstieg, Diktat-Kontrolle, Texte, Aufträge, These, Musteranfang, Ausblick
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { TEIL1, TEIL2, STOLPER, clean } = require("./diktat");
const AUSZ = JSON.parse(fs.readFileSync(path.join(__dirname, "auszuege.json"), "utf8"));

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Kafka – Brief an den Vater";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";

const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 20, color: NAVY, margin: 0, valign: "top" }, o));
const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
const pill = (s, label, x, y, w, col, size = 14) => { rect(s, x, y, w, 0.44, col, { r: 0.22 }); txt(s, label, { x, y, w, h: 0.44, fontSize: size, bold: true, color: WHITE, align: "center", valign: "middle" }); };
function base(kicker, title, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  rect(s, 0, 0, 13.333, 1.05, NAVY, { r: 0 });
  txt(s, kicker.toUpperCase(), { x: 0.5, y: 0.12, w: 9.5, h: 0.3, fontSize: 13, bold: true, color: ICE, charSpacing: 2 });
  txt(s, title, { x: 0.5, y: 0.42, w: 10.3, h: 0.58, fontFace: HEAD, fontSize: 27, bold: true, color: WHITE });
  if (opts.time) { rect(s, 11.0, 0.18, 1.9, 0.7, opts.timeCol || RED, { r: 0.12 }); txt(s, opts.time, { x: 11.0, y: 0.18, w: 1.9, h: 0.7, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" }); }
  return s;
}
function textWithLines(s, lines, x, y, w, size, col) {
  const lh = size / 72 * 1.32;
  lines.forEach((l, i) => {
    txt(s, String(i + 1), { x, y: y + i * lh, w: 0.45, h: lh, fontSize: size - 5, color: MUTED, align: "right", valign: "middle" });
    txt(s, l, { x: x + 0.6, y: y + i * lh, w: w - 0.6, h: lh, fontFace: HEAD, fontSize: size, valign: "middle" });
  });
  s.addShape(pres.shapes.LINE, { x: x + 0.52, y, w: 0, h: lines.length * lh, line: { color: col, width: 2 } });
}

// 1 Einstieg
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "FRANZ KAFKA · „BRIEF AN DEN VATER“ · 1919", { x: 0.9, y: 0.8, w: 10, h: 0.4, fontSize: 15, bold: true, color: ICE, charSpacing: 2 });
  txt(s, "„Liebster Vater,\nDu hast mich letzthin einmal gefragt, warum ich behaupte, ich hätte Furcht vor Dir. Ich wußte Dir, wie gewöhnlich, nichts zu antworten, zum Teil eben aus der Furcht, die ich vor Dir habe …“", { x: 0.9, y: 1.5, w: 11.4, h: 3.2, fontFace: HEAD, fontSize: 30, italic: true, color: WHITE, paraSpaceAfter: 8 });
  rect(s, 0.9, 5.25, 11.5, 1.1, RED);
  txt(s, "Wovor hat ein 36-Jähriger Furcht – und warum schreibt er, statt zu reden?", { x: 1.2, y: 5.25, w: 11, h: 1.1, fontFace: HEAD, fontSize: 26, bold: true, color: WHITE, valign: "middle" });
  s.addNotes("Briefanfang vorlesen, die Klasse liest mit. Impulsfrage stellen, zwei, drei Vermutungen sammeln. Dann die These an die Tafel: „Kafkas ‚Furcht‘ gründet auf Kommunikationsproblemen.“ Überleitung zum Diktat: „Bevor wir das prüfen, brauchen wir den Hintergrund – den schreibt ihr jetzt mit.“");
}

// 2–3 Diktat Selbstkontrolle
[["Teil 1 · Kontext", TEIL1, 23], ["Teil 2 · Fachbegriffe", TEIL2, 24]].forEach(([k, arr, fs]) => {
  const s = base("Diktat · Selbstkontrolle · " + k, "Vergleiche Wort für Wort");
  txt(s, clean(arr), { x: 0.6, y: 1.3, w: 12.1, h: 5.1, fontFace: HEAD, fontSize: fs, lineSpacingMultiple: 1.25 });
  rect(s, 0.6, 6.55, 12.1, 0.6, LIGHT2);
  txt(s, "Fehler farbig anstreichen · Anzahl notieren · Kommafehler extra zählen", { x: 0.85, y: 6.55, w: 11.6, h: 0.6, fontSize: 17, bold: true, valign: "middle" });
  s.addNotes("Text stehen lassen, bis alle verglichen haben. Keine Besprechung – nur Selbstkontrolle.");
});
// 4 Stolperstellen
{
  const s = base("Diktat · Selbstkontrolle", "Darauf achten – auch in jeder Klausur");
  const cols = [RED, GOLD, GREEN, BLUE];
  STOLPER.forEach(([h, regel, bsp], i) => {
    const x = 0.5 + (i % 2) * 6.2, y = 1.25 + Math.floor(i / 2) * 2.95;
    rect(s, x, y, 6.0, 2.8, LIGHT);
    rect(s, x, y, 0.14, 2.8, cols[i], { r: 0.02 });
    txt(s, h, { x: x + 0.35, y: y + 0.12, w: 5.5, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: cols[i] });
    txt(s, regel, { x: x + 0.35, y: y + 0.6, w: 5.45, h: 0.85, fontSize: 16 });
    txt(s, bsp.join("\n"), { x: x + 0.35, y: y + 1.45, w: 5.45, h: 1.3, fontSize: 16, italic: true, color: MUTED, paraSpaceAfter: 3 });
  });
  s.addNotes("Kurz zeigen. Danach den Fahrplan an die Tafel und das Arbeitsblatt austeilen.");
}

// 5 Abschnitt A
{
  const s = base("Abschnitt A · gemeinsam · Matrix in den Hefter", "„ich verlernte das Reden“ (S. 5 f.)");
  textWithLines(s, AUSZ.A, 0.4, 1.25, 12.6, 20, BLUE);
  s.addNotes("Erst 3 Minuten still lesen lassen. Dann im Gespräch Zeile für Zeile die Matrix entwickeln (Impulse im Tafelscript). Stellen am Whiteboard markieren, über die gerade gesprochen wird.");
}

// 6 Auftrag 1
{
  const s = base("Auftrag 1 · allein", "Untersuche Abschnitt C wie Abschnitt A", { time: "13 Min." });
  txt(s, "Lege in deinem Hefter die Matrix an und fülle sie aus:", { x: 0.6, y: 1.35, w: 12, h: 0.45, fontSize: 21 });
  const heads = [["Z.", 1.2], ["Handlung des Vaters", 3.7], ["Modell (Seite / Axiom)", 3.6], ["Wirkung auf den Sohn (Zitat)", 3.6]];
  let x = 0.6;
  heads.forEach(([h, w]) => { rect(s, x, 1.95, w - 0.08, 0.65, NAVY, { r: 0.04 }); txt(s, h, { x: x + 0.12, y: 1.95, w: w - 0.3, h: 0.65, fontSize: 17, bold: true, color: WHITE, valign: "middle" }); x += w; });
  txt(s, "mindestens vier Zeilen · jede mit Zitat und Zeilenangabe", { x: 0.6, y: 2.75, w: 12, h: 0.45, fontSize: 19, bold: true, color: RED });
  rect(s, 0.6, 3.5, 12.1, 3.55, "FFF8EC");
  txt(s, "Wenn du nicht weiterkommst", { x: 0.9, y: 3.65, w: 11, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: GOLD });
  txt(s, [
    { text: "Wie klingen die Fragen des Vaters – und wie sind sie gemeint? (Z. 3–5)", options: { bullet: true, breakLine: true } },
    { text: "Was sagen Lachen und Gesicht? (Z. 5)", options: { bullet: true, breakLine: true } },
    { text: "Mit wem spricht der Vater eigentlich? (Z. 9–10)", options: { bullet: true, breakLine: true } },
    { text: "Was macht der Sohn daraus? (Z. 11–15)", options: { bullet: true } },
  ], { x: 0.9, y: 4.2, w: 11.5, h: 2.7, fontSize: 19, paraSpaceAfter: 8 });
  s.addNotes("Timer 13 Minuten. Herumgehen, zwei gute Zeilen für das Zusammentragen vormerken.");
}

// 7 Abschnitt C (zum Zusammentragen)
{
  const s = base("Abschnitt C · zusammentragen", "Ironie (S. 6)");
  textWithLines(s, AUSZ.C, 0.4, 1.25, 12.6, 20, RED);
  s.addNotes("Beim Zusammentragen stehen lassen, damit alle die genannten Zeilen finden. Matrix C an der Tafel ergänzen.");
}

// 8 These prüfen
{
  const s = base("These prüfen", "Stimmt die These?");
  rect(s, 0.6, 1.4, 12.1, 1.0, LIGHT2);
  txt(s, "These: Kafkas „Furcht“ gründet auf Kommunikationsproblemen.", { x: 0.9, y: 1.4, w: 11.6, h: 1.0, fontFace: HEAD, fontSize: 23, bold: true, valign: "middle" });
  txt(s, "Kafka schreibt aber auch:", { x: 0.6, y: 2.75, w: 12, h: 0.45, fontSize: 20, bold: true, color: MUTED });
  [["„Wieder hüte ich mich zu behaupten, daß ich nur durch Dich so wurde; Du verstärktest nur, was war, aber Du verstärktest es sehr …“", "S. 6"],
    ["„… während es nur selbstverständliche Folge Deiner Stärke und meiner Schwäche war.“", "A, Z. 16"]].forEach(([q, src], i) => {
    const y = 3.3 + i * 1.75;
    rect(s, 0.6, y, 12.1, 1.55, "FBECEB");
    txt(s, q, { x: 0.9, y: y + 0.12, w: 10.3, h: 1.3, fontFace: HEAD, fontSize: 21, italic: true, valign: "middle" });
    txt(s, src, { x: 11.2, y: y + 0.12, w: 1.3, h: 1.3, fontSize: 16, bold: true, color: RED, align: "right", valign: "middle" });
  });
  s.addNotes("Erst fragen: „Bestätigen A und C die These?“ Diese Folie nur zeigen, wenn die Gegenstimme nicht von selbst kommt. Ergebnis: Die These trägt, braucht aber eine Ergänzung (Machtgefälle).");
}

// 9 Auftrag 2
{
  const s = base("Auftrag 2 · allein", "Vom Tafelbild zum Fließtext", { time: "15 Min." });
  txt(s, "Schreibe deine Analyse von A und C als zusammenhängenden Text in deinen Hefter – so wie in der Klausur.", { x: 0.6, y: 1.35, w: 12, h: 0.75, fontSize: 21 });
  [["Einleitung", "Wer schreibt an wen, wann, worüber?"], ["Hauptteil", "je Zeile der Matrix: Behauptung → Beleg (Zitat + Z.) → Erklärung · erst A, dann C"], ["Schluss", "Stützen die Stellen die These? Was muss man ergänzen?"]].forEach(([h, d], i) => {
    const y = 2.3 + i * 0.8;
    pill(s, h, 0.6, y + 0.1, 2.2, NAVY, 16);
    txt(s, d, { x: 3.05, y, w: 9.6, h: 0.65, fontSize: 19, valign: "middle" });
  });
  rect(s, 0.6, 4.85, 12.1, 2.25, "FFF8EC");
  txt(s, "Satzbausteine", { x: 0.9, y: 4.98, w: 6, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: GOLD });
  txt(s, "In seinem „Brief an den Vater“ (1919) …  ·  Indem der Vater …, sendet er …  ·  Dies zeigt sich in … (Z. …)  ·  Hinzu kommt, dass …  ·  Die Folge ist …  ·  Besonders deutlich wird dies, wenn …  ·  Insgesamt stützt die Stelle die These, weil …  ·  Allerdings …", { x: 0.9, y: 5.5, w: 11.5, h: 1.5, fontSize: 18, italic: true, paraSpaceAfter: 4 });
  s.addNotes("Timer 15 Minuten. Wer nach 5 Minuten keinen Anfang hat: nächste Folie (Musteranfang). Wer nicht fertig wird, schreibt zu Beginn der nächsten Stunde weiter – keine Hausaufgabe.");
}

// 10 Musteranfang
{
  const s = base("Hilfe zum Start", "So könnte dein Text beginnen");
  rect(s, 0.6, 1.4, 12.1, 5.6, LIGHT);
  txt(s, [
    { text: "In seinem „Brief an den Vater“ (1919) versucht Franz Kafka, seinem Vater zu erklären, warum er Furcht vor ihm hat. In den beiden Abschnitten beschreibt er, wie der Vater mit ihm gesprochen hat: durch Redeverbote (A) und durch Ironie (C).", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 10 } },
    { text: "Indem der Vater jede Widerrede verbietet, sendet er kaum eine Sachinformation, sondern einen Appell – „Schweig!“ – und vor allem eine Beziehungsbotschaft: Er allein bestimmt. Seine Drohung „kein Wort der Widerrede!“ (A, Z. 5) unterstreicht er nonverbal mit der „erhobene[n] Hand“ (Z. 5). …" },
  ], { x: 0.95, y: 1.6, w: 11.4, h: 5.2, fontFace: HEAD, fontSize: 21, lineSpacingMultiple: 1.2 });
  s.addNotes("Optional, frühestens nach 5 Minuten Schreibzeit zeigen. Nur Einleitung und erster Satz zu A – den Rest schreiben die Schülerinnen und Schüler selbst.");
}

// 11 Ausblick
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  pill(s, "AUSBLICK", 0.9, 1.6, 2.2, RED, 13);
  txt(s, "Übungsklausur", { x: 0.9, y: 2.3, w: 11, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, "Dieselben Schritte an einem neuen Ausschnitt aus dem „Brief an den Vater“:\nverstehen → untersuchen → schreiben.", { x: 0.9, y: 3.4, w: 11.4, h: 1.6, fontSize: 24, color: ICE, paraSpaceAfter: 6 });
  s.addNotes("Ein, zwei Einleitungen vorlesen lassen, dann Ausblick auf die Übungsklausur (Tischregeln, S. 5).");
}

pres.writeFile({ fileName: process.argv[2] || "Praesentation_Kafka_Brief_an_den_Vater.pptx" });
