const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Französische Revolution 1789 – Gruppenpuzzle";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Französische Revolution · Geschichte 8";
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
  s.addText(runs, { isTextBox: true, x: 0.9, y: 6.2, w: 11.6, h: 0.85, fontFace: BODY, fontSize: 24, color: WHITE, valign: "middle", margin: 0 });
}

// Aufgabenfolie: gleiches Format wie immer (Nummer + Titel, nummerierte Schritte, Zeit, Arbeitsform)
function taskSlide(kicker, title, stepsArr, meta, notes, done = "Fertig? Stift hinlegen und noch einmal durchlesen.", stepW = 12.15) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.35, w: 8, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const per = stepW > 10 ? 56 : 38;
  const hs = stepsArr.map((st) => (st.length > per * 2 ? 1.4 : st.length > per ? 0.95 : 0.5));
  const total = hs.reduce((a, b) => a + b, 0) + (hs.length - 1) * 0.08;
  box(s, 0.6, 1.5, stepW, total + 0.24, LIGHT2);
  let yy = 1.62;
  stepsArr.forEach((st, i) => {
    numCircle(s, i + 1, 0.8, yy + 0.04, 0.42);
    txt(s, st, { x: 1.4, y: yy, w: stepW - 0.95, h: hs[i], fontSize: 24, bold: true });
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
    s.addText(sent, { isTextBox: true, x: 1.4, y, w: 11.2, h: rh - 0.08, fontFace: HEAD, fontSize: o.size || 24, color: NAVY, valign: "middle", margin: 0 });
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
function merkSlide(nr, title, kern, beispiel, test, notes, kernSize = 24) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, `MERKSATZ ${nr} · ABSCHREIBEN UND ROT UMRAHMEN`, { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: RED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.5, w: 12.15, h: 2.1, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  s.addText(rich(kern), { isTextBox: true, x: 0.95, y: 1.55, w: 11.45, h: 2.0, fontFace: HEAD, fontSize: 24, color: NAVY, valign: "middle", margin: 0 });
  box(s, 0.6, 3.72, 12.15, 1.25, LIGHT2);
  s.addText([{ text: "Beispiel:   ", options: { bold: true, color: MUTED, fontFace: BODY, fontSize: 24 } }, ...rich(beispiel)], { isTextBox: true, x: 0.9, y: 3.72, w: 11.6, h: 1.25, fontFace: HEAD, fontSize: 24, color: NAVY, valign: "middle", margin: 0 });
  box(s, 0.6, 5.08, 12.15, 0.97, LIGHT);
  s.addText([{ text: "Test:   ", options: { bold: true, color: RED, fontFace: HEAD, fontSize: 24 } }, ...rich(test)], { isTextBox: true, x: 0.9, y: 5.08, w: 11.6, h: 0.97, fontFace: BODY, fontSize: 24, color: NAVY, valign: "middle", margin: 0 });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schreibe den Merksatz mit Beispiel ab und rahme ihn rot ein.      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
  footer(s);
  s.addNotes(notes || `Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz ${nr}. Rahmt ihn rot ein.“ Warten, bis alle den Stift hingelegt haben.`);
}


const fs = require("fs");
// ============ Bildimpuls ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "BILDIMPULS · STIFTE LIEGEN", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, "Paris, Sommer 1789", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const img = "../material/bastille.jpg";
  const hasImg = fs.existsSync(img);
  if (hasImg) s.addImage({ path: img, x: 6.9, y: 1.55, w: 5.85, h: 4.5, sizing: { type: "contain", w: 5.85, h: 4.5 } });
  const w = hasImg ? 5.9 : 12.15;
  [["1", "Was ist zu sehen?"], ["2", "Was ist vorher passiert?"]].forEach(([n, q], i) => {
    const y = 1.7 + i * 1.8;
    box(s, 0.6, y, w, 1.5, i ? LIGHT : LIGHT2);
    numCircle(s, n, 0.85, y + 0.5, 0.5, RED);
    txt(s, q, { x: 1.6, y, w: w - 1.2, h: 1.5, fontFace: HEAD, fontSize: 28, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Stifte liegen. Ihr schaut nur. Dann meldet ihr euch.", options: {} }]);
  footer(s);
  s.addNotes("Bild vom Sturm auf die Bastille einfügen (Gemälde von 1789, z. B. von Jean-Pierre Houël, gemeinfrei) oder als Datei unter ../material/bastille.jpg ablegen und das Skript neu ausführen. Ansage: „Stifte liegen. Ihr schaut nur. Dann meldet ihr euch.“ Anbindung an die letzte Stunde (Münzgeschichte, „Heute sagt niemand etwas“). Leitfrage an die Tafel: „Wie kippt die Stimmung 1789?“");
}

// ============ Ü1 Expertengruppen ============
taskSlide("Übung 1 · Expertengruppe", "Dein Thema verstehen", ["Zeichne das Raster von der Tafel ab.", "Lies deinen Text und unterstreiche Stichworte.", "Ergänze deine Zeile. Vergleicht in der Gruppe."], "10 Minuten · erst allein, dann in der Gruppe",
  "Ansage: „Setzt euch in eure Expertengruppen. Zeichnet zuerst das Raster von der Tafel ab. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Nach 3 Minuten Einzelarbeit: „Jetzt arbeitet ihr in eurer Gruppe.“ Raster an der Tafel: Wann? · Was geschah? · Welches Problem steckt dahinter? Zeilen A, B, C. Lösung siehe Tafelskript.");

// ============ Ü2 Stammgruppe ============
taskSlide("Übung 2 · Stammgruppe", "Dein Thema erklären", ["Erkläre dein Thema in 3 Minuten.", "Ergänze das Raster mit den anderen Themen.", "Stellt Nachfragen, bis alles klar ist."], "12 Minuten · in der Gruppe · leise",
  "Ansage: „Jetzt arbeitet ihr in der Stammgruppe. Setzt euch zu A, B und C zusammen. Erst erklärt A, dann B, dann C.“ Beim Wechsel ansagen: „A hat fertig, jetzt erklärt B.“");

// ============ Ü3 Abschluss ============
taskSlide("Übung 3 · Abschluss", "Warum begann die Revolution?", ["Schreibe Ü3 ins Heft.", "Nenne zwei Gründe für die Revolution.", "Begründe, welcher Grund wichtiger war."], "4 Minuten · allein · leise",
  "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Mögliche Gründe: leere Staatskasse, ungleiche Steuerlast, keine Mitsprache des Dritten Standes (Bauern und Bürger), Hunger und Brotpreis, Angst vor der Willkür des Königs. 1 bis 2 Antworten vorlesen lassen.");

pres.writeFile({ fileName: "../Praesentation_Franzoesische_Revolution.pptx" }).then(() => console.log("ok"));
