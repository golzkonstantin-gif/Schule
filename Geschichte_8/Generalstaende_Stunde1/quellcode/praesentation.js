const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Generalstände, Stunde 1";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Generalstände · Geschichte 8";
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


// ============ 1 Text lesen ============
taskSlide("Übung 1 · Text und Zeitstrahl", "Wie kam es zu den Generalständen?", ["Lies den Text. Unterstreiche die Jahreszahlen.", "Trage die Daten in den Zeitstrahl ein.", "Prüfe: Passt jedes Datum zum Ereignis?"], "5 Minuten · allein · leise",
  "Ansage: „Ich teile Text und Zeitstrahl aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Danach Kontrolle an der Tafel (siehe Tafelskript): 1614 · 1778–1783 · 1787 · 1788 · August 1788 · Frühjahr 1789 · 5. Mai 1789. Dann „Vergleicht und verbessert mit Grün.“");

// ============ 2 Zahlen ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "ZAHLEN · GEMEINSAM · OHNE HEFT", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, "Wer hat wie viel Gewicht?", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const cw = [3.4, 1.9, 1.9, 4.95];
  const heads = ["", "Klerus", "Adel", "Dritter Stand (Bauern und Bürger)"];
  const rows = [["Menschen", "150 000", "500 000", "20 000 000"], ["Anteil", "ca. 1 %", "ca. 2 %", "ca. 97 %"], ["Abgeordnete", "ca. 290", "ca. 270", "ca. 580"], ["Stimmen nach Ständen", "1", "1", "1"]];
  let x = 0.6;
  heads.forEach((h, i) => {
    box(s, x, 1.5, cw[i] - 0.08, 1.0, NAVY, { r: 0.05 });
    txt(s, h, { x: x + 0.08, y: 1.5, w: cw[i] - 0.24, h: 1.0, fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle" });
    x += cw[i];
  });
  rows.forEach((r, j) => {
    let xx = 0.6; const y = 2.6 + j * 0.88;
    r.forEach((v, i) => {
      box(s, xx, y, cw[i] - 0.08, 0.82, i === 0 ? LIGHT2 : LIGHT);
      txt(s, v, { x: xx + 0.08, y, w: cw[i] - 0.24, h: 0.82, fontFace: i === 0 ? BODY : HEAD, fontSize: 24, bold: i === 0, align: i === 0 ? "left" : "center", valign: "middle" });
      xx += cw[i];
    });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Meldet euch. Was fällt euch auf?", options: {} }]);
  footer(s);
  s.addNotes("Ansage: „Meldet euch. Ihr schreibt noch nichts auf. Was fällt euch an diesen Zahlen auf?“ Erwartung: Der Dritte Stand (Bauern und Bürger) ist fast die ganze Bevölkerung und stellt etwa die Hälfte der Abgeordneten, hat aber nur eine von drei Stimmen. Klerus und Adel (ca. 3 %) haben zusammen zwei von drei Stimmen. Zahlen der Abgeordneten bitte gegen das Buch prüfen. Die Anzahl der Menschen stammt aus dem Gruppenpuzzle-Material.");
}

// ============ 3 Antrag ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "ABSTIMMUNG · NUR EIN SPIEL", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: RED, charSpacing: 2, bold: true });
  txt(s, "Antrag", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.55, w: 12.15, h: 3.0, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  txt(s, "Nur Schüler putzen künftig das Klassenzimmer und räumen die Tische auf. Lehrkräfte sind davon befreit.", { x: 0.95, y: 1.6, w: 11.45, h: 2.9, fontFace: HEAD, fontSize: 32, bold: true, valign: "middle" });
  txt(s, "Klerus: Lehrkraft 1 · Adel: Lehrkraft 2 · Dritter Stand (Bauern und Bürger): die Schüler", { x: 0.6, y: 4.8, w: 12.15, h: 0.9, fontFace: HEAD, fontSize: 24, color: MUTED });
  bar(s, [{ text: "Hinweis: ", options: { bold: true } }, { text: "Das ist nur ein Spiel. Der Antrag gilt nicht.", options: {} }]);
  footer(s);
  s.addNotes("Ansage: „Wir spielen jetzt eine Abstimmung. Das ist nur ein Spiel, es gilt nicht.“ Schritt 1, nach Ständen: Jede Gruppe eine Stimme. Klerus (L1) Ja, Adel (L2) Ja, Dritter Stand (Schüler) Nein: 2:1, angenommen. Schritt 2, nach Köpfen: 2 Ja, 12 Nein: abgelehnt. Fragen: Wie fühlt sich das an? Was ist unfair? Sind die Lehrkräfte das Problem oder die Regeln? Sollbruchstelle: Bei Zeitnot entfällt die Abstimmung und beginnt Stunde 2.");
}

// ============ 4 Merksatz ============
merkSlide(1, "Abstimmung nach Ständen",
  "Bei den Generalständen hat jeder Stand **eine Stimme**. Klerus und Adel (zusammen ca. 3 % der Bevölkerung) haben damit **zwei von drei Stimmen** und können den Dritten Stand (Bauern und Bürger, ca. 97 %) **immer überstimmen**.",
  "Eine Stimme des Klerus wiegt rund **133-mal** so viel wie eine des Dritten Standes.",
  "Wie viele Menschen stehen hinter **einer Stimme**?",
  "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“ Rechnung dazu an der Tafel: 20 000 000 : 150 000 ≈ 133 (Klerus), 20 000 000 : 500 000 = 40 (Adel).", 24);

// ============ 5 Auswerten ============
taskSlide("Übung 2 · Auswerten", "Ist die Abstimmung fair?", ["Schreibe Ü2 ins Heft.", "Beende den Satz: „Die Abstimmung nach Ständen ist unfair, weil …“", "Begründe mit mindestens einer Zahl."], "4 Minuten · allein · leise",
  "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Beispiel: „… weil der Dritte Stand (Bauern und Bürger) ca. 97 % der Menschen stellt, aber nur eine von drei Stimmen hat und immer überstimmt wird.“ Ausblick: Was passiert, wenn der Dritte Stand nach Köpfen fordert und der König ablehnt? (Stunde 2)");

pres.writeFile({ fileName: "../Praesentation_Generalstaende_Stunde1.pptx" }).then(() => console.log("ok"));
