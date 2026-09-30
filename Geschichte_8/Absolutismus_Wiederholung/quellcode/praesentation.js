const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Absolutismus und Ständegesellschaft – Wiederholung";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "C98A1E", BLUE = "3F7CC4", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Absolutismus · Wiederholung · Geschichte 8";
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
  box(s, 0.6, 6.45, 12.15, 0.6, NAVY);
  s.addText(runs, { isTextBox: true, x: 0.9, y: 6.45, w: 11.6, h: 0.6, fontFace: BODY, fontSize: 17, color: WHITE, valign: "middle", margin: 0 });
}

// Aufgabenfolie: gleiches Format wie immer (Nummer + Titel, nummerierte Schritte, Zeit, Arbeitsform)
function taskSlide(kicker, title, stepsArr, meta, notes, done = "Fertig? Stift hinlegen und noch einmal durchlesen.") {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.35, w: 8, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const hs = stepsArr.map((st) => (st.length > 68 ? 0.8 : 0.45));
  const total = hs.reduce((a, b) => a + b, 0) + (hs.length - 1) * 0.08;
  box(s, 0.6, 1.55, 12.15, total + 0.3, LIGHT2);
  let yy = 1.7;
  stepsArr.forEach((st, i) => {
    numCircle(s, i + 1, 0.8, yy + 0.015, 0.42);
    txt(s, st, { x: 1.4, y: yy, w: 11.2, h: hs[i], fontSize: 21, bold: true });
    yy += hs[i] + 0.08;
  });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      ", options: {} }, { text: done, options: { italic: true, color: ICE } }]);
  footer(s);
  if (notes) s.addNotes(notes);
  return { s, top: 1.55 + total + 0.3 + 0.3 };
}
function sentences(s, list, top, o = {}) {
  const rows = list.length, avail = 6.3 - top, rh = Math.min(o.rh || 0.85, avail / rows);
  list.forEach((sent, i) => {
    const y = top + i * rh;
    box(s, 0.6, y, 12.15, rh - 0.08, LIGHT);
    numCircle(s, i + 1, 0.75, y + (rh - 0.08 - 0.4) / 2, 0.4, RED);
    s.addText(sent, { isTextBox: true, x: 1.4, y, w: 11.2, h: rh - 0.08, fontFace: HEAD, fontSize: o.size || 22, color: NAVY, valign: "middle", margin: 0 });
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
function merkSlide(nr, title, kern, beispiel, wofuer, test, notes, kernSize = 22) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, `MERKSATZ ${nr} · ABSCHREIBEN UND ROT UMRAHMEN`, { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: RED, charSpacing: 2, bold: true });
  txt(s, title, { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.55, w: 12.15, h: 2.0, rectRadius: 0.08, fill: { color: WHITE }, line: { color: RED, width: 3 } });
  s.addText(rich(kern), { isTextBox: true, x: 0.95, y: 1.65, w: 11.45, h: 1.8, fontFace: HEAD, fontSize: kernSize, color: NAVY, valign: "middle", margin: 0 });
  box(s, 0.6, 3.7, 12.15, 0.6, LIGHT2);
  s.addText([{ text: "Beispiel:   ", options: { bold: true, color: MUTED, fontFace: BODY, fontSize: 17 } }, ...rich(beispiel)], { isTextBox: true, x: 0.9, y: 3.7, w: 11.6, h: 0.6, fontFace: HEAD, fontSize: 19, color: NAVY, valign: "middle", margin: 0 });
  [["Wofür brauche ich das?", wofuer, 0.6], ["Test", test, 6.8]].forEach(([h, body, x]) => {
    box(s, x, 4.45, 5.95, 1.8, LIGHT);
    txt(s, h, { x: x + 0.25, y: 4.58, w: 5.45, h: 0.4, fontFace: HEAD, fontSize: 19, bold: true, color: RED });
    s.addText(rich(body), { isTextBox: true, x: x + 0.25, y: 5.0, w: 5.45, h: 1.15, fontFace: BODY, fontSize: 19, color: NAVY, valign: "top", margin: 0 });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Schreibe den Merksatz mit Beispiel ab und rahme ihn rot ein.      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
  footer(s);
  s.addNotes(notes || `Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz ${nr}. Rahmt ihn rot ein.“ Warten, bis alle den Stift hingelegt haben.`);
}

// ============ Titel ============
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "Absolutismus und Ständegesellschaft", { x: 0.9, y: 2.5, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE });
  txt(s, "Wiederholung vor dem Test", { x: 0.9, y: 3.55, w: 11, h: 0.5, fontSize: 22, italic: true, color: ICE });
  txt(s, "Geschichte, Klasse 8", { x: 0.9, y: 6.6, w: 5, h: 0.3, fontSize: 13, color: ICE });
  s.addNotes("Diese Präsentation enthält die Merksätze (zum Abschreiben) und die Aufgaben. Tafelbilder und die Tabelle zur Ständegesellschaft entstehen an der Tafel (siehe Tafelskript). Lösungen stehen in den Notizen.");
}

// ============ Testankündigung ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "ANSAGE ZU BEGINN · STIFTE LIEGEN", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, "Nächste Woche schreiben wir einen Test", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const items = [
    ["Ständegesellschaft", "Vorrechte · Pflichten · Anzahl · Mitbestimmung (dazu: Bildung)"],
    ["Begriff Absolutismus", "Was bedeutet er? Woran erkenne ich ihn?"],
    ["Die fünf Säulen des Absolutismus", "Nennen und erklären, wie der König damit seine Macht sichert"],
  ];
  items.forEach(([h, b], i) => {
    const y = 1.75 + i * 1.45;
    box(s, 0.6, y, 12.15, 1.25, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, i + 1, 0.9, y + 0.375, 0.5, RED);
    txt(s, h, { x: 1.7, y: y + 0.14, w: 10.8, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true });
    txt(s, b, { x: 1.7, y: y + 0.68, w: 10.8, h: 0.45, fontSize: 19, color: MUTED });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Wir wiederholen genau das und ergänzen, was in euren Unterlagen noch fehlt.", options: {} }]);
  footer(s);
  s.addNotes("Ansage: „Stifte liegen. Ihr hört nur zu. Nächste Woche schreiben wir einen Test zu diesen drei Themen. Die Stunde heute ist eure Vorbereitung: Was hier nicht klappt, wisst ihr danach.“ Noch keinen Termin im Detail diskutieren; Fragen zum Test am Stundenende.");
}

// ============ Einstieg Wahr oder falsch ============
{
  const { s, top } = taskSlide("Einstieg · ohne Heft", "Wahr oder falsch?", ["Lies die vier Behauptungen. Entscheide still: wahr oder falsch?", "Melde dich: Nenne deine Entscheidung und begründe sie.", "Bei „falsch“: Verbessere die Aussage."], "4 Minuten · gemeinsam · mündlich",
    "Ansage: „Stifte liegen. Ihr schreibt noch nichts auf. Entscheidet still und meldet euch.“ Lösung: 1 wahr (ca. 20 Mio. von gut 20,6 Mio.) · 2 falsch: Der Adel zahlte keine Steuern · 3 wahr (Säule Heer) · 4 falsch: 1685 beendete Ludwig die religiöse Toleranz (Säule Religion). Nicht lange erklären, sondern notieren, was unsicher ist. Es kommt in den folgenden Phasen.", "Wer begründet, bekommt das Wort.");
  sentences(s, [
    "Der Dritte Stand war mit Abstand der größte Stand.",
    "Der Adel musste hohe Steuern an den König zahlen.",
    "Ludwig XIV. verzichtete auf Söldner und baute ein stehendes Heer auf.",
    "Ludwig XIV. erlaubte den Protestanten die freie Religionsausübung.",
  ], top, { size: 22, rh: 0.8 });
}

// ============ Absolutismus ============
merkSlide(1, "Absolutismus",
  "**Absolutismus** war eine Herrschaftsform in Europa im **17. und 18. Jahrhundert**, in der ein einzelner Monarch (König oder Fürst) **uneingeschränkt (absolut)** herrschte.",
  "Ludwig XIV. verkündet 1661: „Ich regiere jetzt selbst.“",
  "Der Begriff ordnet eine ganze Epoche ein und ist im Test die Grundlage für alle weiteren Fragen.",
  "Kann jemand den König stoppen (Parlament, Verfassung, Stände)? **Nein** → Absolutismus.",
  "Vorher: Wort „Absolutismus“ an die Tafel, gemeinsam klären: absolut = uneingeschränkt. Dann Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“", 21);

{
  const { s, top } = taskSlide("Übung 1", "Absolutismus – ja oder nein?", ["Schreibe Ü1 an den Rand.", "Schreibe Nummer + „Absolutismus“ oder „kein Absolutismus“.", "Begründe mit dem Test: Kann jemand den König stoppen?"], "4 Minuten · allein · leise",
    "Lösung: 1 Absolutismus (keine Kontrolle) · 2 kein Absolutismus (Parlament kann den König stoppen) · 3 Absolutismus (Beamte sind dem König zum Gehorsam verpflichtet, er kontrolliert sie) · 4 kein Absolutismus (Wahl = Demokratie, Macht ist begrenzt). Nach der Einzelarbeit: „Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“ Dann: „Vergleicht und verbessert mit Grün.“");
  sentences(s, [
    "Der König entscheidet allein über Gesetze, Krieg und Steuern.",
    "Ein Parlament kann Steuererhöhungen des Königs verbieten.",
    "Beamte handeln im Namen des Königs und müssen ihm gehorchen.",
    "Die Bürger wählen alle vier Jahre ihre Regierung.",
  ], top, { size: 22, rh: 0.8 });
}

// ============ Fünf Säulen ============
merkSlide(2, "Die fünf Säulen des Absolutismus",
  "Ludwig XIV. sicherte seine Macht auf **fünf Säulen**: **1** Verwaltung und Justiz · **2** Adel und Hof von Versailles · **3** Wirtschaft (Merkantilismus) · **4** Religion (Gottesgnadentum) · **5** Heer (stehendes Heer).",
  "[r:Verwaltung] kontrolliert das Land · [r:Hof] bindet den Adel · [r:Wirtschaft] liefert Geld · [r:Religion] begründet die Macht · [r:Heer] setzt sie durch",
  "Im Test sollst du die Säulen nennen **und** erklären, wie der König damit seine Macht sichert.",
  "Bei jeder Säule fragen: **Was sichert der König damit?**",
  "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 2. Rahmt ihn rot ein.“ Danach Arbeitsblatt Lückentext verteilen. Tipp: Die Säulen sind aus dem Gruppenpuzzle bekannt – kurz fragen, wer welche Säule bearbeitet hat.", 20);

{
  const { s, top } = taskSlide("Übung 2 · Arbeitsblatt", "Lückentext: Die fünf Säulen", ["Schreibe Ü2 oben auf das Arbeitsblatt.", "Lies den Text einmal ganz. Ergänze dann die 12 Lücken mit dem Wortspeicher.", "Streiche jedes benutzte Wort durch. Zwei Wörter passen nicht."], "6 Minuten · allein · leise",
    "Lösung: 1 absoluter · 2 Beamte · 3 Intendanten · 4 Steuerfreiheit · 5 Versailles · 6 Merkantilismus · 7 Zölle · 8 Gottesgnadentum · 9 Glaube · 10 stehendes · 11 400.000 · 12 Drittel. Nicht passend: Söldner, Parlament. Kontrolle: „Wer drankommt, liest einen Absatz vor. Alle prüfen mit.“ Dann „Vergleicht und verbessert mit Grün.“ Schnelle machen die ★-Aufgabe (wichtigste Säule begründen).",
    "Fertig? Stift hinlegen. Schnelle: ★-Aufgabe.");
  box(s, 0.6, top + 0.1, 12.15, 1.6, LIGHT);
  txt(s, "Wortspeicher", { x: 0.9, y: top + 0.22, w: 4, h: 0.35, fontFace: HEAD, fontSize: 17, bold: true, color: RED });
  txt(s, "Beamte · Glaube · Söldner · Merkantilismus · Gottesgnadentum · Versailles · Steuerfreiheit · Zölle · stehendes · 400.000 · Drittel · absoluter · Intendanten · Parlament", { x: 0.9, y: top + 0.62, w: 11.6, h: 1.0, fontFace: HEAD, fontSize: 20 });
}

// ============ Ständegesellschaft ============
{
  const { s, top } = taskSlide("Übung 3 · Tafel", "Die Ständegesellschaft aus dem Kopf", ["Klappe das Heft zu. Ihr meldet euch und nennt, was ihr wisst.", "Wir füllen die Tabelle an der Tafel gemeinsam aus.", "Jetzt schreibst du die fertige Tabelle mit Lineal ab."], "8 Minuten · gemeinsam, dann abschreiben",
    "Ansage: „Hefte zu. Ihr schreibt noch nichts auf. Meldet euch.“ Raster aus dem Gruppenpuzzle (Stand | Anzahl | Bildung | Vorrechte | Pflichten) plus Spalte Mitbestimmung bereits leer an der Tafel. Inhalte siehe Tafelskript (aus den Gruppenpuzzle-Texten). Nach dem gemeinsamen Ausfüllen: „Jetzt schreibt ihr ab: die Tabelle. Mit Lineal.“ Lösungen siehe Tafelskript. Hinweis: Der Text nennt zur Mitbestimmung des Klerus nichts.  Der Dritte Stand ist ausgeschlossen, der Adel bestimmt das politische Geschehen.",
    "Fertig? Stift hinlegen.");
  const cols = ["Stand", "Anzahl", "Bildung", "Vorrechte", "Pflichten", "Mitbestimmung"];
  const cw = [1.8, 1.3, 1.9, 2.4, 2.4, 2.35];
  let x = 0.6;
  cols.forEach((c, i) => {
    box(s, x, top - 0.15, cw[i] - 0.1, 0.45, NAVY, { r: 0.05 });
    txt(s, c, { x, y: top - 0.15, w: cw[i] - 0.1, h: 0.45, fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
    x += cw[i];
  });
  ["1. Stand", "2. Stand", "3. Stand"].forEach((r, j) => {
    let xx = 0.6; const y = top + 0.38 + j * 0.72;
    cols.forEach((c, i) => {
      box(s, xx, y, cw[i] - 0.1, 0.64, i === 0 ? LIGHT2 : LIGHT);
      if (i === 0) txt(s, r, { x: xx, y, w: cw[i] - 0.1, h: 0.64, fontFace: HEAD, fontSize: 19, bold: true, align: "center", valign: "middle" });
      xx += cw[i];
    });
  });
}

merkSlide(3, "Die Ständegesellschaft",
  "In der Ständegesellschaft gehört jeder Mensch **durch Geburt** zu einem von **drei Ständen**: **Klerus (ca. 150 000)**, **Adel (ca. 500 000)** und **Dritter Stand (ca. 20 000 000)**. Klerus und Adel haben **Vorrechte**. Der Dritte Stand zahlt **Steuern und Abgaben** und ist von der **politischen Mitbestimmung ausgeschlossen**.",
  "Adel: **Jagdrecht** · Klerus: **Kirchenzehnt** · 3. Stand: **Abgaben**",
  "Du erklärst, warum Ludwig den Adel bei Laune hält und warum der Dritte Stand die Last trägt.",
  "Wer zahlt, wer ist ausgeschlossen? **Wenige Menschen haben die Vorrechte, viele tragen die Last.**",
  "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 3. Rahmt ihn rot ein.“ Alle Zahlen stammen aus den Gruppenpuzzle-Texten. Der Text nennt zur Mitbestimmung des Klerus nichts.", 20);

{
  const { s, top } = taskSlide("Übung 4", "Im Absolutismus fühle ich mich …", ["Schreibe Ü4 an den Rand und wähle einen Stand: Klerus, Adel oder Dritter Stand.", "Schreibe den Satz ab und beende ihn: „Im Absolutismus fühle ich mich … , weil …“", "Nenne in deiner Begründung mindestens ein Recht und eine Pflicht deines Standes."], "6 Minuten · allein · leise",
    "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Danach 3 Schüler vorlesen lassen (verschiedene Stände): „Wer drankommt, liest vor. Alle prüfen mit.“ Kriterien stehen auf der nächsten Folie. Mögliche Antwort: „Im Absolutismus fühle ich mich als Adeliger privilegiert, weil ich keine Steuern zahlen muss und allein jagen darf. Dafür muss ich mich ehrenhaft verhalten und Verwaltungsaufgaben übernehmen.“ / „… als Bauer ausgenutzt, weil ich Steuern, hohe Abgaben und Pachtgebühren zahlen muss und von der Mitbestimmung ausgeschlossen bin.“");
  box(s, 0.6, top - 0.1, 5.95, 1.5, LIGHT);
  txt(s, "Gefühle", { x: 0.85, y: top, w: 5.4, h: 0.35, fontFace: HEAD, fontSize: 18, bold: true, color: RED });
  txt(s, "privilegiert · sicher · zufrieden · ausgenutzt · benachteiligt · machtlos · überlastet", { x: 0.85, y: top + 0.4, w: 5.45, h: 1.0, fontFace: HEAD, fontSize: 18 });
  box(s, 6.8, top - 0.1, 5.95, 1.5, LIGHT);
  txt(s, "Fachbegriffe für die Begründung", { x: 7.05, y: top, w: 5.4, h: 0.35, fontFace: HEAD, fontSize: 18, bold: true, color: RED });
  txt(s, "Vorrechte · Steuerfreiheit · Jagdrecht · Kirchenzehnt · Pachtgebühren · Mitbestimmung", { x: 7.05, y: top + 0.4, w: 5.45, h: 1.0, fontFace: HEAD, fontSize: 18 });
}

{
  const { s, top } = taskSlide("Kontrolle zu Ü4", "Prüft die Sätze der anderen", ["Hört zu: Welcher Stand wurde gewählt?", "Prüft: Wird ein Recht und eine Pflicht genannt? Passen sie zum Stand?", "Prüft: Passt das Gefühl zur Begründung?"], "3 Minuten · gemeinsam · mündlich",
    "Ansage: „Wer drankommt, liest vor. Alle prüfen mit.“ Typische Fehler: Adel als Steuerzahler, Klerus mit Wehrdienst (befreit), Dritter Stand mit Vorrechten. Dann: „Verbessert mit Grün.“", "Wer korrigiert, nennt die Begründung mit Fachbegriff.");
}

// ============ Schluss ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "ZUM SCHLUSS · STIFTE LIEGEN", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, "So lernst du bis zum Test", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const items = ["Merksatz 1 auswendig: Was bedeutet Absolutismus?", "Lückentext noch einmal ohne Wortspeicher ausfüllen", "Tabelle Ständegesellschaft aus dem Kopf aufschreiben", "Deinen Ü4-Satz zu einem zweiten Stand noch einmal schreiben"];
  items.forEach((x, i) => {
    const y = 1.7 + i * 1.05;
    box(s, 0.6, y, 12.15, 0.9, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, i + 1, 0.85, y + 0.2, 0.5, RED);
    txt(s, x, { x: 1.6, y, w: 10.9, h: 0.9, fontFace: HEAD, fontSize: 22, valign: "middle" });
  });
  bar(s, [{ text: "Test: ", options: { bold: true } }, { text: "nächste Woche · Ständegesellschaft · Absolutismus · fünf Säulen", options: {} }]);
  footer(s);
  s.addNotes("Ansage: „Stifte liegen. Das ist eure Lernliste für den Test. Fragen zum Test?“ Die Liste bewusst kurz halten; Testtermin noch einmal nennen.");
}

pres.writeFile({ fileName: "../Praesentation_Absolutismus_Wiederholung.pptx" }).then(() => console.log("ok"));
