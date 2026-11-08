const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Mo 9.11. · Balladen wiederholen";
const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4", VIO = "7A5BA6";
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

// ================= Mo 9.11. · Einzelstunde · AVIVA =================
// Leseband (Ich-Du-Wir-Würfeln, eigener Text) → Kreuzworträtsel Balladen → Sprachtraining: Fehler im Chat → Zielcheck
// Aufruf: node praesentation_mo_09-11.js (aus quellcode/, liest kw.json)
const fs = require("fs");
const KW = JSON.parse(fs.readFileSync(__dirname + "/kw.json", "utf8"));
FOOT = "Montag, 9.11. · Deutsch 7A";
const ZIEL = "Ich kann die Merkmale einer Ballade benennen und Nebensätze mit Subjunktion richtig schreiben.";
const runs = (parts, size = 22) => parts.map(([x, k]) => ({ text: x, options: Object.assign({ fontSize: size }, k === "sub" ? { bold: true, color: VIO } : k === "fin" ? { bold: true, underline: { style: "sng" } } : k === "mv" ? { bold: true, color: GREEN, underline: { style: "sng" } } : k === "hv" ? { bold: true, color: RED, underline: { style: "sng" } } : k === "k" ? { bold: true, color: RED } : k === "m" ? { color: MUTED } : {}) }));
const zeit = (s, meta, y = 6.25) => {
  box(s, 0.6, y, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: y + 0.08, w: 11.7, h: 0.4, fontSize: 16 });
};

// 1 Titel
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "MONTAG, 9.11. · DEUTSCH 7A", { x: 0.9, y: 2.0, w: 9, h: 0.4, fontSize: 14, color: ICE, charSpacing: 2 });
  txt(s, "Balladen – alles im Griff?", { x: 0.9, y: 2.5, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE });
  txt(s, "Leseband · Kreuzworträtsel · Sprachtraining: Fehler im Chat", { x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 22, italic: true, color: ICE });
  s.addNotes("Folie liegt beim Hereinkommen auf. Ziel und Ablauf stehen zusätzlich links an der Seitentafel.");
}

// 2 Ziel und Ablauf
{
  const s = base("Ankommen", "Ziel und Ablauf", "Ansage: „Stifte liegen. Ihr hört nur zu. Heute holen wir alles zu den Balladen zurück ins Gedächtnis – und am Ende schreibt ihr Sätze mit Subjunktion ohne Fehler. Am Schluss prüft ihr im Zielcheck, ob ihr das Ziel erreicht habt.“");
  box(s, 0.6, 1.6, 12.1, 1.05, WHITE, { line: RED, lw: 2.5 });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 0.9, y: 1.78, w: 11.6, h: 0.75, fontSize: 22 });
  [["Leseband: Ich-Du-Wir-Würfeln", "10 Min."], ["Kreuzworträtsel Balladen + nächster Schritt gemeinsam", "13 Min."], ["Sprachtraining: Fehler im Chat (* Nachricht 5)", "13 Min."], ["Zielcheck", "6 Min."]].forEach(([a, m], i) => {
    const y = 3.0 + i * 0.75;
    num(s, i + 1, 0.6, y, NAVY, 0.5);
    txt(s, a, { x: 1.3, y: y + 0.06, w: 9.0, h: 0.45, fontSize: 22, bold: true });
    txt(s, m, { x: 10.6, y: y + 0.08, w: 2.0, h: 0.45, fontSize: 18, color: MUTED, align: "right" });
  });
}

// 3 Leseband
{
  const s = base("Vorwissen aktivieren · Leseband", "Ich-Du-Wir-Würfeln", "Eigener Text aus dem Leseband. Tischgruppen zu 3–4, ein Würfel pro Gruppe. Abschnitt = ein Absatz. Ansage: „Jetzt arbeitet ihr in der Tischgruppe. Würfelt. Wer liest, liest laut. Alle anderen lesen still mit. Nach jedem Abschnitt sagt die Person, die gewürfelt hat, in einem Satz, worum es ging. Dann würfelt die nächste Person.“ Stopp-Signal ansagen: „Stopp. Würfel in die Mitte, Stifte liegen.“ Herumgehen, auf Betonung und Tempo achten.");
  [["1–2", "ICH", "Wer gewürfelt hat, liest."], ["3–4", "DU", "Die Person links liest."], ["5–6", "WIR", "Alle lesen gemeinsam."]].forEach(([n, w, d], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 1.65, 3.85, 1.9, WHITE, { line: NAVY, lw: 2 });
    txt(s, "⚄ " + n, { x: x + 0.25, y: 1.8, w: 3.4, h: 0.45, fontSize: 18, color: MUTED });
    txt(s, w, { x: x + 0.25, y: 2.2, w: 3.4, h: 0.65, fontFace: HEAD, fontSize: 34, bold: true, color: [RED, VIO, GREEN][i] });
    txt(s, d, { x: x + 0.25, y: 2.95, w: 3.4, h: 0.45, fontSize: 18 });
  });
  [["1", "Würfelt reihum. Die Zahl bestimmt, wer den nächsten Abschnitt laut liest."], ["2", "Alle anderen lesen still mit."], ["3", "Nach jedem Abschnitt: Wer gewürfelt hat, sagt in einem Satz, worum es ging."]].forEach(([n, x], i) => {
    const y = 3.9 + i * 0.7;
    num(s, n, 0.6, y, NAVY, 0.45);
    txt(s, x, { x: 1.25, y: y + 0.03, w: 11.4, h: 0.5, fontSize: 20, bold: true });
  });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "10 Minuten · Tischgruppe · ein Würfel      " }, { text: "Stopp-Signal: Würfel in die Mitte.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}

// 4 Kreuzworträtsel – Aufgabe
uebung("Vorwissen aktivieren · Wiederholung Balladen", "Ü1  Kreuzworträtsel: Balladen", [
  "Klebe den Schnipsel ins Heft.",
  "Trage die Wörter ein – der Reihe nach, von 1 bis 13.",
  "Die Nummern gehen vom Allgemeinen ins Detail: Ballade → Gattungen → Merkmale.",
  "Schreibe noch nichts bei Nr. 2 – das ist der nächste Schritt, den wir gemeinsam machen.",
], "8 Minuten · allein", "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht auf der Folie. Acht Minuten.“ Schnipsel austeilen (A4 quer, 2 pro Blatt). Herumgehen; typische Hürden: UMARMENDER (nur das Adjektiv), ERZÄHLER mit Ä in einem Kästchen.");

// 5 Lösung Kreuzworträtsel: Gitter + Begriffsbaum
{
  const s = base("Vorwissen aktivieren · Ü1 Lösung", "Vom Allgemeinen ins Detail", "Ansage: „Vergleicht und verbessert mit Grün.“ Dann am Baum rechts kurz nachfragen: Was ist der Oberbegriff? Welche drei Gattungen stecken in der Ballade? Was unterscheidet Epik und Dramatik? (Erzähler – ja oder nein).");
  const letters = {};
  KW.placed.forEach(([w, r, c, d]) => [...w].forEach((ch, i) => (letters[d === "H" ? `${r},${c + i}` : `${r + i},${c}`] = ch)));
  const cs = 0.33, gx = 0.6, gy = 1.75;
  for (let r = 0; r < KW.h; r++) for (let c = 0; c < KW.w; c++) {
    const ch = letters[`${r},${c}`];
    if (!ch) continue;
    s.addShape(pres.shapes.RECTANGLE, { x: gx + c * cs, y: gy + r * cs, w: cs, h: cs, fill: { color: WHITE }, line: { color: NAVY, width: 0.75 } });
    txt(s, ch, { x: gx + c * cs, y: gy + r * cs, w: cs, h: cs, fontSize: 12, bold: true, color: RED, align: "center", valign: "middle" });
  }
  const T = [
    [0, "1 BALLADE", NAVY], [1, "2 LYRIK", "2A8C8C"], [2, "3 STROPHE · 4 VERS", null], [2, "5 REIM", null], [3, "6 PAARREIM · 7 KREUZREIM · 8 UMARMENDER", MUTED], [2, "9 KEHRREIM", null],
    [1, "10 EPIK", VIO], [2, "11 ERZÄHLER", null], [1, "12 DRAMATIK", "C77C12"], [2, "13 DIALOG", null],
  ];
  const tx = 6.25;
  box(s, tx - 0.15, 1.6, 6.6, 4.35, LIGHT);
  T.forEach(([lvl, w, col], i) => {
    txt(s, (lvl ? "└ " : "") + w, { x: tx + lvl * 0.45, y: 1.75 + i * 0.41, w: 6.2 - lvl * 0.45, h: 0.38, fontSize: lvl < 2 ? 18 : 15, bold: lvl < 2, color: col || NAVY });
  });
  txt(s, "Ä = ein Kästchen", { x: 0.6, y: 5.85, w: 5, h: 0.3, fontSize: 12, color: MUTED });
}

// 6 Nächster Schritt gemeinsam
{
  const s = base("Vorwissen aktivieren · nächster Schritt gemeinsam", "Ü1 Nr. 2  Warum ist „Paddy Fingal“ eine Ballade?", "Gemeinsam (Ich – Wir): Lehrkraft macht das erste Merkmal vor (Lyrik: Strophen mit Reimen), die Klasse nennt Epik und Dramatik, alle schreiben mit. Ansage: „Jetzt machen wir den nächsten Schritt gemeinsam. Nehmt Lösungswörter aus dem Rätsel. Ich fange an, ihr ergänzt. Alle schreiben auf den Schnipsel mit.“ Lösung: … weil sie Strophen mit Reimen hat, weil ein Erzähler die Geschichte erzählt und weil in Str. 7 Schaya ohne Erzähler spricht.");
  [["Lyrik", "2A8C8C", "Strophen, Verse, Reim, Kehrreim"], ["Epik", VIO, "Erzähler"], ["Dramatik", "C77C12", "Dialog, Figurenrede ohne Erzähler"]].forEach(([g, col, w], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 1.65, 3.85, 1.2, WHITE, { line: col, lw: 2 });
    txt(s, g, { x: x + 0.25, y: 1.75, w: 3.4, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true, color: col });
    txt(s, w, { x: x + 0.25, y: 2.22, w: 3.4, h: 0.5, fontSize: 16, color: MUTED });
  });
  box(s, 0.6, 3.2, 12.1, 1.0, LIGHT);
  txt(s, runs([["„Paddy Fingal“ ist eine Ballade, ", ""], ["weil", "sub"], [" sie …", ""]], 26), { x: 0.9, y: 3.42, w: 11.6, h: 0.6, fontFace: HEAD });
  txt(s, "… weil ein Erzähler …  … weil in Str. 7 …", { x: 0.9, y: 4.45, w: 11.6, h: 0.5, fontSize: 20, color: MUTED, italic: true });
  zeit(s, "5 Minuten · gemeinsam, alle schreiben mit");
}

// 7 Lösung + Brücke zum Sprachtraining
{
  const s = base("Informieren · Sprachtraining", "Was fällt euch an diesem Satz auf?", "Erst die Lösung zeigen, alle verbessern mit Grün. Dann fragen: „Was fällt euch an den weil-Sätzen auf?“ Erwartung: weil = Subjunktion, finites Verb am Ende (hat, erzählt, spricht), vor weil ein Komma. Überleitung: „Genau darauf achtet ihr gleich im Chat.“");
  box(s, 0.6, 1.7, 12.1, 2.6, LIGHT);
  txt(s, runs([["„Paddy Fingal“ ist eine Ballade, ", ""], ["weil", "sub"], [" sie Strophen mit Reimen ", ""], ["hat", "fin"], [", ", "k"], ["weil", "sub"], [" ein Erzähler die Geschichte ", ""], ["erzählt", "fin"], [" und ", ""], ["weil", "sub"], [" in Str. 7 Schaya ohne Erzähler ", ""], ["spricht", "fin"], [".", ""]], 26), { x: 0.95, y: 1.95, w: 11.5, h: 2.2, fontFace: HEAD, valign: "middle" });
  [["weil", VIO, "Subjunktion – leitet den Nebensatz ein"], ["hat", NAVY, "finites Verb – steht am Ende (unterstrichen)"], [",", RED, "Komma vor der Subjunktion"]].forEach(([w, col, d], i) => {
    const y = 4.6 + i * 0.55;
    txt(s, w, { x: 0.9, y, w: 1.0, h: 0.45, fontSize: 22, bold: true, color: col, fontFace: HEAD, underline: i === 1 ? { style: "sng" } : undefined });
    txt(s, d, { x: 2.0, y: y + 0.03, w: 10, h: 0.45, fontSize: 19 });
  });
}

// 8 Merksatz (Wiederholung Mi 4.11.)
{
  const s = base("Informieren · Sprachtraining", "Merksatz: Nebensatz mit Subjunktion", "Merksatz vom Mittwoch (klebt schon im Heft) – nicht abschreiben, nur kurz zeigen. Ansage: „Stifte liegen. Ihr hört nur zu.“ Den Test einmal laut am Satz von Folie 7 durchgehen.");
  box(s, 0.6, 1.65, 12.1, 1.35, WHITE, { line: RED, lw: 3 });
  txt(s, "Eine Subjunktion (weil, dass, als, nachdem, ob …) leitet einen Nebensatz ein. Das finite Verb steht im Nebensatz am Ende. Vor der Subjunktion steht ein Komma.", { x: 0.9, y: 1.8, w: 11.6, h: 1.1, fontSize: 21, bold: true, valign: "middle" });
  txt(s, runs([["Ich komme später", ""], [",", "k"], [" ", ""], ["weil", "sub"], [" ich noch duschen ", ""], ["muss", "mv"], [".", ""]], 24), { x: 0.9, y: 3.25, w: 11.6, h: 0.55, fontFace: HEAD });
  box(s, 0.6, 4.05, 5.9, 1.9, LIGHT);
  txt(s, [{ text: "Wofür?\n", options: { bold: true, color: RED, fontSize: 18 } }, { text: "Mit Nebensätzen begründest du (weil) und gibst wieder, was jemand sagt (dass) – im Aufsatz und in der Analyse einer Ballade.", options: { fontSize: 17 } }], { x: 0.85, y: 4.2, w: 5.45, h: 1.65 });
  box(s, 6.8, 4.05, 5.9, 1.9, LIGHT);
  txt(s, [{ text: "Test\n", options: { bold: true, color: RED, fontSize: 18 } }, { text: "1  Subjunktion gefunden?\n2  Steht das finite Verb ganz am Ende?\n3  Steht vor der Subjunktion ein Komma?", options: { fontSize: 17 } }], { x: 7.05, y: 4.2, w: 5.45, h: 1.65 });
}

// 9 Chat – Aufgabe
const CHAT = [
  ["r", "", "Und? Hast du gegen Paddy gewonnen? 💪"],
  ["l", "1", "Nein. Ich bin sofort weggerannt weil ich das Baby gesehen habe."],
  ["r", "2", "Welches Baby? Ich weiß nur, dass Paddy hat einen schwarzen Bart."],
  ["l", "3", "Seine Frau hat gesagt, dass ich muss leise sein."],
  ["l", "4", "Ich bin ans Bett gegangen, weil ich war neugierig."],
  ["r", "* 5", "Ich glaube dass Paddy selbst im Bett lag 😂"],
];
const SOL = [
  ["1", [["Ich bin sofort weggerannt", ""], [",", "k"], [" ", ""], ["weil", "sub"], [" ich das Baby gesehen ", ""], ["habe", "hv"], [".", ""]], "K"],
  ["2", [["Ich weiß nur, ", ""], ["dass", "sub"], [" Paddy einen schwarzen Bart ", ""], ["hat", "fin"], [".", ""]], "V"],
  ["3", [["Seine Frau hat gesagt, ", ""], ["dass", "sub"], [" ich leise sein ", ""], ["muss", "mv"], [".", ""]], "V"],
  ["4", [["Ich bin ans Bett gegangen, ", ""], ["weil", "sub"], [" ich neugierig ", ""], ["war", "fin"], [".", ""]], "V"],
  ["* 5", [["Ich glaube", ""], [",", "k"], [" ", ""], ["dass", "sub"], [" Paddy selbst im Bett ", ""], ["lag", "fin"], [".", ""]], "K"],
];
{
  const s = base("Verarbeiten · Sprachtraining", "Ü2  Fehler im Chat: der Schotte und sein Bruder", "Zu zweit, jede/r schreibt ins eigene Heft. Ansage: „Jetzt arbeitet ihr zu zweit. In jeder Nachricht mit Nummer steckt genau ein Fehler: Komma oder Verbstellung. Schreibt die Sätze richtig ins Heft. Acht Minuten. Nachricht 5 mit Stern macht, wer fertig ist.“ Herumgehen. Typische Hürde: Nr. 1 und 5 klingen richtig – der Fehler ist das fehlende Komma.");
  box(s, 0.6, 1.55, 7.4, 4.55, "E9EEF6");
  txt(s, "Der Schotte  ↔  sein Bruder", { x: 0.8, y: 1.6, w: 7.0, h: 0.3, fontSize: 12, color: MUTED, align: "center" });
  CHAT.forEach(([side, n, m], i) => {
    const y = 1.95 + i * 0.69;
    const x = side === "l" ? 0.8 : 2.0;
    box(s, x, y, 5.8, 0.6, side === "l" ? WHITE : "DCF3E4", { r: 0.15 });
    txt(s, [{ text: n ? n + "  " : "", options: { bold: true, color: VIO } }, { text: m }], { x: x + 0.15, y: y + 0.06, w: 5.55, h: 0.5, fontSize: 14, valign: "middle" });
  });
  [["1", "Finde in jeder Nachricht mit Nummer genau einen Fehler: Komma (K) oder Verbstellung (V)?"], ["2", "Schreibe den Satz richtig ins Heft. Notiere dahinter K oder V."], ["3", "Kreise die Subjunktion ein. Unterstreiche das finite Verb."]].forEach(([n, x], i) => {
    num(s, n, 8.3, 1.65 + i * 1.45, VIO, 0.45);
    txt(s, x, { x: 8.9, y: 1.65 + i * 1.45, w: 3.85, h: 1.35, fontSize: 17, bold: true });
  });
  zeit(s, "8 Minuten · zu zweit, jede/r ins eigene Heft · Nr. 5 mit *");
}

// 10–12 Chat – Lösung zum Aufdecken
[[0, 2], [0, 4], [0, 5]].forEach(([a, b], k) => {
  const s = base("Verarbeiten · Ü2 Lösung", "Ü2  So ist es richtig" + (k < 2 ? "  (" + (k + 1) + "/3)" : ""), k === 0
    ? "Aufdecken Schritt für Schritt. Wer drankommt, liest seinen Satz vor und nennt K oder V. Dann aufdecken, alle verbessern mit Grün. Niemand muss an die Tafel."
    : k === 1 ? "Bei 3 nachfragen: Welches Verb ist finit? (muss – Modalverb, grün). Bei 4: Im Gespräch hört man oft „weil ich war …“ – geschrieben ist das falsch."
    : "Sprachreflexion: Warum ist das beim Sprechen kein Problem, beim Schreiben aber schon? (Beim Schreiben fehlen Pausen und Betonung – das Komma und die Verbstellung zeigen, wo der Nebensatz beginnt und endet.)");
  SOL.slice(a, b).forEach(([n, r, typ], i) => {
    const y = 1.6 + i * 0.92;
    box(s, 0.6, y, 12.1, 0.78, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.17, w: 0.7, h: 0.45, fontSize: 18, bold: true, color: VIO });
    txt(s, runs(r, 21), { x: 1.5, y: y + 0.16, w: 9.4, h: 0.5, fontFace: HEAD });
    box(s, 11.1, y + 0.17, 1.4, 0.45, typ === "K" ? RED : NAVY, { r: 0.22 });
    txt(s, typ === "K" ? "Komma" : "Verbst.", { x: 11.1, y: y + 0.17, w: 1.4, h: 0.45, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
  txt(s, [{ text: "Farben: ", options: { bold: true } }, { text: "Subjunktion", options: { color: VIO, bold: true } }, { text: " · finites Verb unterstrichen · " }, { text: "Hilfsverb rot", options: { color: RED, bold: true } }, { text: " · " }, { text: "Modalverb grün", options: { color: GREEN, bold: true } }], { x: 0.6, y: 6.4, w: 12.1, h: 0.35, fontSize: 14, color: MUTED });
});

// 13 Zielcheck
const ZC = [
  ["1", "Ergänze:  abab heißt ________ .  Ein Gespräch von Figuren ohne Erzähler heißt ________ .", null],
  ["2", "Schaya versteckt Paddy im Bett. Der Schotte soll ihn nicht finden.", "weil"],
  ["3", "Der Schotte rennt sofort weg. Er sieht das riesige Baby.", "als"],
];
{
  const s = base("Auswerten · Zielcheck", "Zielcheck", "Allein im Heft, 3 Minuten, dann Lösungsfolie, verbessern mit Grün. Herumgehen: Wer noch Fehler hat, wird in der Wiederholung der nächsten Stunde gezielt aufgegriffen.");
  txt(s, [{ text: "Ziel: ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 0.6, y: 1.5, w: 12.1, h: 0.5, fontSize: 17 });
  txt(s, "Bei 2 und 3: Verbinde mit der Subjunktion in Klammern. Unterstreiche das finite Verb.", { x: 0.6, y: 2.05, w: 12.1, h: 0.45, fontSize: 19, bold: true });
  ZC.forEach(([n, q, sj], i) => {
    const y = 2.65 + i * 1.05;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.2, w: 0.5, h: 0.45, fontSize: 20, bold: true, color: VIO });
    txt(s, q, { x: 1.35, y: y + 0.12, w: sj ? 9.2 : 11.2, h: 0.62, fontSize: 19, valign: "middle" });
    if (sj) txt(s, "(" + sj + ")", { x: 10.6, y: y + 0.2, w: 1.9, h: 0.45, fontSize: 20, bold: true, color: VIO, align: "right" });
  });
  zeit(s, "3 Minuten · allein im Heft");
}

// 14 Zielcheck – Lösung
{
  const s = base("Auswerten · Zielcheck Lösung", "Zielcheck: Lösung – Ziel erreicht?", "Mit Grün verbessern. Abschlussfrage an die Klasse: „Haben wir das Ziel erreicht? Was habt ihr heute wiederholt?“ (Gattungen und Merkmale der Ballade; Nebensatz: Komma vor der Subjunktion, finites Verb am Ende.)");
  box(s, 0.6, 1.6, 12.1, 0.85, LIGHT);
  txt(s, "1", { x: 0.8, y: 1.8, w: 0.5, h: 0.45, fontSize: 20, bold: true, color: VIO });
  txt(s, [{ text: "abab = " }, { text: "Kreuzreim", options: { bold: true, color: RED } }, { text: "  ·  Gespräch ohne Erzähler = " }, { text: "Dialog", options: { bold: true, color: RED } }, { text: " (Dramatik)", options: { color: MUTED } }], { x: 1.35, y: 1.8, w: 11.2, h: 0.5, fontSize: 21 });
  [["2", [["Schaya versteckt Paddy im Bett, ", ""], ["weil", "sub"], [" der Schotte ihn nicht finden ", ""], ["soll", "mv"], [".", ""]]],
   ["3", [["Der Schotte rennt sofort weg, ", ""], ["als", "sub"], [" er das riesige Baby ", ""], ["sieht", "fin"], [".", ""]]]].forEach(([n, r], i) => {
    const y = 2.65 + i * 1.05;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    txt(s, n, { x: 0.8, y: y + 0.2, w: 0.5, h: 0.45, fontSize: 20, bold: true, color: VIO });
    txt(s, runs(r, 22), { x: 1.35, y: y + 0.2, w: 11.2, h: 0.5, fontFace: HEAD });
  });
  box(s, 0.6, 5.0, 12.1, 1.0, WHITE, { line: RED, lw: 2.5 });
  txt(s, [{ text: "Ziel erreicht?  ", options: { bold: true, color: RED } }, { text: ZIEL, options: { italic: true } }], { x: 0.9, y: 5.15, w: 11.6, h: 0.7, fontSize: 19, valign: "middle" });
}

pres.writeFile({ fileName: __dirname + "/../Montag_09-11_Deutsch_7A.pptx" });
