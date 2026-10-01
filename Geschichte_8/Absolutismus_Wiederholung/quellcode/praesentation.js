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
    txt(s, b, { x: 1.7, y: y + 0.66, w: 10.8, h: 0.5, fontSize: 24, color: MUTED });
  });
  bar(s, [{ text: "Heute: ", options: { bold: true } }, { text: "Wir wiederholen genau das und ergänzen, was in euren Unterlagen noch fehlt.", options: {} }]);
  footer(s);
  s.addNotes("Ansage: „Stifte liegen. Ihr hört nur zu. Nächste Woche schreiben wir einen Test zu diesen drei Themen. Die Stunde heute ist eure Vorbereitung: Was hier nicht klappt, wisst ihr danach.“ Noch keinen Termin im Detail diskutieren; Fragen zum Test am Stundenende.");
}

// ============ Fünf Säulen ============
// Leitfrage: Wie kann ein König ein Land regieren?
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "EINSTIEG · STIFTE LIEGEN", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, "Wie kann ein König ein ganzes Land regieren?", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 30, bold: true });
  box(s, 0.6, 1.75, 12.15, 3.4, LIGHT2);
  txt(s, "Stellt euch vor: Ein Land mit über 20 Millionen Menschen. Ein König entscheidet allein.", { x: 1.0, y: 2.1, w: 11.4, h: 0.9, fontFace: HEAD, fontSize: 26 });
  txt(s, "Was braucht er, damit seine Befehle im ganzen Land befolgt werden?", { x: 1.0, y: 3.3, w: 11.4, h: 0.9, fontFace: HEAD, fontSize: 26, bold: true, color: RED });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Meldet euch. Ihr schreibt noch nichts auf.      ", options: {} }, { text: "Sammelt Ideen im Gespräch.", options: { italic: true, color: ICE } }]);
  footer(s);
  s.addNotes("Ansage: „Stifte liegen. Ihr schreibt noch nichts auf. Meldet euch.“ Ideen der Klasse sammeln, noch nicht bewerten. Erwartbar: Beamte, Soldaten, Geld, Gesetze, Religion, Adel bei Laune halten. Dann überleiten: „Ihr erinnert euch: Es gab fünf Säulen.“");
}
// Erinnerung: Welche fünf Säulen waren das?
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "ERINNERN · GEMEINSAM · TAFELBILD", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, "Es gab fünf Säulen. Welche waren das?", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  // Dach
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.75, w: 12.15, h: 0.9, fill: { color: NAVY }, line: { type: "none" } });
  txt(s, "Absolute Herrschaft Ludwigs XIV.", { x: 0.6, y: 1.75, w: 12.15, h: 0.9, fontFace: HEAD, fontSize: 26, bold: true, color: WHITE, align: "center", valign: "middle" });
  const cw = 2.25, gap = 0.225;
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * (cw + gap);
    box(s, x, 2.85, cw, 3.3, LIGHT2, { line: NAVY, lw: 1.5 });
    numCircle(s, i + 1, x + cw / 2 - 0.25, 3.05, 0.5, RED);
    txt(s, "?", { x, y: 3.85, w: cw, h: 1.2, fontFace: HEAD, fontSize: 60, bold: true, color: MUTED, align: "center", valign: "middle" });
  }
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Meldet euch. Nennt die Säule und wie sie dem König hilft.      ", options: {} }, { text: "Noch nichts aufschreiben.", options: { italic: true, color: ICE } }]);
  footer(s);
  s.addNotes("Tafelbild gemeinsam entwickeln (nicht abzeichnen): Dach und fünf Säulen an die Tafel, die Klasse füllt sie. Stichworte: 1 Verwaltung und Justiz → Kontrolle über das Land · 2 Adel und Hof → kein adeliger Widerstand · 3 Wirtschaft → Geld für Hof, Beamte und Heer · 4 Religion → Herrschaft von Gott gewollt, ein Glaube · 5 Heer → Macht durchsetzen. Was fehlt, wird im Lückentext ergänzt. Ansage: „Meldet euch. Ihr schreibt noch nichts auf.“ Dann: „Jetzt schreibt ihr von der Folie ab: Merksatz 1.“");
}
merkSlide(1, "Die fünf Säulen des Absolutismus",
  "Ludwig XIV. sicherte seine Macht auf **fünf Säulen**: **1** Verwaltung und Justiz · **2** Adel und Hof von Versailles · **3** Wirtschaft (Merkantilismus) · **4** Religion (Gottesgnadentum) · **5** Heer (stehendes Heer).",
  "[r:Verwaltung] kontrolliert das Land · [r:Hof] bindet den Adel · [r:Wirtschaft] liefert Geld · [r:Religion] begründet die Macht · [r:Heer] setzt sie durch",
  "Bei jeder Säule fragen: **Was sichert der König damit?**",
  "Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“ Danach Arbeitsblatt Lückentext verteilen.", 20);

{
  const { s, top } = taskSlide("Übung 1 · Arbeitsblatt", "Lückentext: Die fünf Säulen", ["Schreibe Ü1 oben auf das Arbeitsblatt.", "Lies den Text einmal ganz. Ergänze dann die 12 Lücken mit dem Wortspeicher.", "Streiche jedes benutzte Wort durch."], "5 Minuten · allein · leise",
    "Lösung: 1 absoluter · 2 Beamte · 3 Intendanten · 4 Steuerfreiheit · 5 Versailles · 6 Merkantilismus · 7 Zölle · 8 Gottesgnadentum · 9 Glaube · 10 stehendes · 11 400.000 · 12 Drittel. Kontrolle: „Wer drankommt, liest einen Absatz vor. Alle prüfen mit.“ Dann „Vergleicht und verbessert mit Grün.“ Schnelle machen die ★-Aufgabe (wichtigste Säule begründen).",
    "Fertig? Stift hinlegen. Schnelle: ★-Aufgabe.");
  box(s, 0.6, top - 0.1, 12.15, 6.1 - top + 0.1, LIGHT);
  txt(s, "Wortspeicher", { x: 0.9, y: top - 0.02, w: 4, h: 0.4, fontFace: HEAD, fontSize: 24, bold: true, color: RED });
  txt(s, "Beamte · Glaube · Merkantilismus · Gottesgnadentum · Versailles · Steuerfreiheit · Zölle · stehendes · 400.000 · Drittel · absoluter · Intendanten", { x: 0.9, y: top + 0.45, w: 11.6, h: 6.1 - top - 0.5, fontFace: HEAD, fontSize: 24 });
}


// ============ Exkurs: Steuern in Deutschland (nach Antwort „Adel steuerfrei“) ============
function exkursSlides() {
  // 1 Brutto und Netto
  {
    const s = pres.addSlide();
    pageNo++;
    s.background = { color: WHITE };
    txt(s, "EXKURS · GEMEINSAM · OHNE HEFT", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
    txt(s, "Exkurs: Steuern in Deutschland", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
    [["Brutto", "Lohn vor allen Abzügen", 0.6], ["Netto", "Das, was auf dem Konto ankommt", 6.8]].forEach(([h, b, x]) => {
      box(s, x, 1.6, 5.95, 2.1, LIGHT2);
      txt(s, h, { x: x + 0.3, y: 1.75, w: 5.4, h: 0.6, fontFace: HEAD, fontSize: 34, bold: true, color: RED });
      txt(s, b, { x: x + 0.3, y: 2.55, w: 5.4, h: 1.0, fontFace: HEAD, fontSize: 24 });
    });
    box(s, 0.6, 3.95, 12.15, 1.9, LIGHT);
    txt(s, "Durchschnittsgehalt in Deutschland (gerundet):", { x: 0.9, y: 4.1, w: 11.6, h: 0.5, fontFace: HEAD, fontSize: 24, color: MUTED });
    txt(s, "ca. 50 000 € brutto im Jahr", { x: 0.9, y: 4.75, w: 11.6, h: 0.8, fontFace: HEAD, fontSize: 34, bold: true });
    bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Meldet euch. Wer erklärt brutto und netto?      ", options: {} }, { text: "Noch nichts aufschreiben.", options: { italic: true, color: ICE } }]);
    footer(s);
    s.addNotes("Kurzer Exkurs zur Verdeutlichung (ca. 5 Minuten, bei Zeitnot weglassen). Ansage: „Stifte liegen. Ihr schreibt noch nichts auf. Meldet euch.“ Brutto und Netto von der Klasse erklären lassen, dann die Zahl nennen. Hinweis: 50 000 € ist ein runder Wert. Der Durchschnitt für Vollzeitbeschäftigte liegt nach meinem Kenntnisstand eher höher (etwa 55 000 bis 60 000 € im Jahr). Aktuellen Wert beim Statistischen Bundesamt prüfen und ggf. anpassen.");
  }
  // 2 Was bleibt netto
  {
    const s = pres.addSlide();
    pageNo++;
    s.background = { color: WHITE };
    txt(s, "EXKURS · GEMEINSAM · OHNE HEFT", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
    txt(s, "Was bleibt von 50 000 € brutto?", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
    const rows = [
      ["Brutto", "50 000 €", NAVY, WHITE],
      ["− Lohnsteuer", "ca. 7 000 €", LIGHT2, NAVY],
      ["− Sozialabgaben (Rente, Kranken-, Pflege-, Arbeitslosenversicherung)", "ca. 10 500 €", LIGHT2, NAVY],
      ["= Netto", "ca. 32 500 €", RED, WHITE],
    ];
    rows.forEach(([l, r, fill, col], i) => {
      const y = 1.6 + i * 1.12, h = 1.0;
      box(s, 0.6, y, 12.15, h, fill);
      txt(s, l, { x: 0.9, y, w: 8.4, h, fontFace: HEAD, fontSize: 24, bold: i === 0 || i === 3, color: col, valign: "middle" });
      txt(s, r, { x: 9.3, y, w: 3.2, h, fontFace: HEAD, fontSize: 28, bold: true, color: col, align: "right", valign: "middle" });
    });
    bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Meldet euch. Wofür braucht der Staat die Steuern?      ", options: {} }, { text: "Noch nichts aufschreiben.", options: { italic: true, color: ICE } }]);
    footer(s);
    s.addNotes("Gerundete Beispielrechnung: ledig, ohne Kinder, Lohnsteuerklasse I, Stand etwa 2025. Mit einem Brutto-Netto-Rechner (z. B. des Bundesfinanzministeriums) gegenprüfen und die Zahlen ggf. anpassen. Impuls: Wofür braucht der Staat Steuern? (Schulen, Straßen, Polizei, Bundeswehr, Beamte). Das ist die Brücke zu Versailles, Heer und Beamten im Absolutismus.");
  }
  // 3 Übung: keine Steuern
  {
    const { s } = taskSlide("Exkurs · Übung", "Stellt euch vor: keine Steuern", ["Du verdienst 50 000 € brutto und zahlst keine Steuern.", "Rechne: Wie viel mehr hast du im Monat? (Lohnsteuer: ca. 7 000 € im Jahr)", "Antworte: Wäre das gerecht, wenn dein Nachbar mit demselben Gehalt Steuern zahlt? Begründe."], "3 Minuten · zu zweit · Flüsterstimme",
      "Ansage: „Jetzt arbeitet ihr zu zweit. Die Aufgabe steht rechts.“ Lösung: 7 000 € : 12 ≈ 580 € mehr pro Monat. Begründung offen, erwartbar: ungerecht, weil beide gleich viel verdienen und die Straßen, Schulen und die Polizei gleich nutzen. Danach 2 bis 3 Antworten sammeln. Leitfrage zur Überleitung: „Und jetzt stellt euch vor, das hängt nur davon ab, in welche Familie ihr geboren seid.“",
      "Fertig? Stift hinlegen.");
  }
  // 4 Brücke
  {
    const s = pres.addSlide();
    pageNo++;
    s.background = { color: WHITE };
    txt(s, "EXKURS · ZURÜCK ZUM ABSOLUTISMUS", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
    txt(s, "Genau das war im Absolutismus Realität", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
    box(s, 0.6, 1.6, 12.15, 1.7, LIGHT2);
    txt(s, [{ text: "Der Adel", options: { bold: true, color: RED } }, { text: " (ca. 500 000 Menschen) zahlte keine Steuern – nur wegen seiner Geburt.", options: {} }], { x: 0.9, y: 1.6, w: 11.6, h: 1.7, fontFace: HEAD, fontSize: 24, valign: "middle" });
    box(s, 0.6, 3.45, 12.15, 1.7, LIGHT);
    txt(s, [{ text: "Der Dritte Stand (Bauern und Bürger)", options: { bold: true, color: RED } }, { text: " (ca. 20 000 000 Menschen) zahlte Steuern und Abgaben für fast alles.", options: {} }], { x: 0.9, y: 3.45, w: 11.6, h: 1.7, fontFace: HEAD, fontSize: 24, valign: "middle" });
    bar(s, [{ text: "Frage: ", options: { bold: true } }, { text: "Wer bezahlte dann Versailles, Heer und Beamte?", options: {} }]);
    footer(s);
    s.addNotes("Ansage: „Meldet euch. Ihr schreibt noch nichts auf.“ Antwort: der Dritte Stand (Bauern und Bürger). Überleitung zurück zu Wahr oder falsch, Aussage 3 (stehendes Heer) und 4: Was kostet das alles? (Heer in Friedenszeiten ein Drittel des Staatshaushalts.) Der Adel zahlte keine Steuern, hatte aber ranghohe Posten in der Armee und alleiniges Jagdrecht.");
  }
}

// ============ Wahr oder falsch (einzeln nacheinander) ============
{
  const WF = [
    ["Der Dritte Stand (Bauern und Bürger) war mit Abstand der größte Stand.", "WAHR", "Etwa 20 Millionen Bauern und Bürger – gegenüber ca. 500 000 Adligen und ca. 150 000 Geistlichen (Klerus).", "1 wahr: ca. 20 Mio. von gut 20,6 Mio. Menschen. Leitfrage im Gespräch: Wer gehört zum Dritten Stand? Wie groß sind die anderen Stände?"],
    ["Der Adel musste hohe Steuern an den König zahlen.", "FALSCH", "Der Adel war steuerfrei – das war eines seiner Vorrechte (neben dem alleinigen Jagdrecht). Steuern zahlte der Dritte Stand (Bauern und Bürger).", "2 falsch: Adel zahlte keine Steuern. Im Gespräch: Welche weiteren Vorrechte kennt ihr? Wer zahlte dann die Steuern?"],
    ["Ludwig XIV. verzichtete auf Söldner und baute ein stehendes Heer auf.", "WAHR", "Säule Heer: einheitlich bewaffnet und uniformiert, in Kasernen, direkt dem König unterstellt – bis zu 400 000 Mann.", "3 wahr (Säule Heer). Im Gespräch: Was kostet ein solches Heer? (ein Drittel des Staatshaushalts in Friedenszeiten)"],
    ["Ludwig XIV. erlaubte den Protestanten die freie Religionsausübung.", "FALSCH", "1685 beendete er die religiöse Toleranz: Wie nur einen König sollte es auch nur einen Glauben geben (Säule Religion, Gottesgnadentum).", "4 falsch: 1685 Ende der Toleranz. Im Gespräch: Wie passt das zum Gottesgnadentum?"],
  ];
  WF.forEach(([aussage, urteil, erkl, note], i) => {
    // Aussage
    {
      const s = pres.addSlide();
      pageNo++;
      s.background = { color: WHITE };
      txt(s, `WAHR ODER FALSCH? · AUSSAGE ${i + 1} VON 4 · OHNE HEFT`, { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
      txt(s, "Wahr oder falsch?", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
      box(s, 0.6, 1.75, 12.15, 3.0, LIGHT2);
      numCircle(s, i + 1, 0.95, 2.1, 0.6, RED);
      txt(s, aussage, { x: 1.9, y: 1.95, w: 10.5, h: 2.6, fontFace: HEAD, fontSize: 30, valign: "middle" });
      bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Entscheide still, melde dich, begründe.      ", options: {} }, { text: "Bei „falsch“: Aussage verbessern.", options: { italic: true, color: ICE } }]);
      footer(s);
      s.addNotes(`Ansage: „Stifte liegen. Ihr schreibt noch nichts auf. Entscheidet still und meldet euch.“ Unterrichtsgespräch: Entscheidung und Begründung sammeln, erst dann die nächste Folie (Antwort) zeigen. Lösung: ${note}`);
    }
    // Antwort
    {
      const s = pres.addSlide();
      pageNo++;
      s.background = { color: WHITE };
      txt(s, `WAHR ODER FALSCH? · ANTWORT ${i + 1} VON 4`, { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
      txt(s, urteil === "WAHR" ? "Wahr" : "Falsch", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true, color: urteil === "WAHR" ? GREEN : RED });
      box(s, 0.6, 1.75, 12.15, 1.5, LIGHT);
      txt(s, aussage, { x: 1.0, y: 1.75, w: 11.4, h: 1.5, fontFace: HEAD, fontSize: 24, valign: "middle", color: MUTED });
      box(s, 0.6, 3.45, 12.15, 2.6, WHITE, { line: urteil === "WAHR" ? GREEN : RED, lw: 3 });
      txt(s, "Begründung", { x: 1.0, y: 3.6, w: 6, h: 0.4, fontFace: HEAD, fontSize: 24, bold: true, color: urteil === "WAHR" ? GREEN : RED });
      txt(s, erkl, { x: 1.0, y: 4.1, w: 11.4, h: 1.9, fontFace: HEAD, fontSize: 26 });
      bar(s, [{ text: "Vergleiche: ", options: { bold: true } }, { text: "Stimmt deine Begründung mit der Folie überein?", options: {} }]);
      footer(s);
      s.addNotes("Erst zeigen, wenn die Klasse begründet hat. Nicht ausführlich erklären: Was unsicher bleibt, kommt in den folgenden Phasen.");
    }
    if (i === 1) exkursSlides();
  });
}

// ============ Absolutismus ============
merkSlide(2, "Absolutismus",
  "**Absolutismus** war eine Herrschaftsform in Europa im **17. und 18. Jahrhundert**, in der ein einzelner Monarch (König oder Fürst) **uneingeschränkt (absolut)** herrschte.",
  "Ludwig XIV. verkündet 1661: „Ich regiere jetzt selbst.“",
  "Kann jemand den König stoppen (Parlament, Verfassung, Stände)? **Nein** → Absolutismus.",
  "Vorher: Wort „Absolutismus“ an die Tafel, gemeinsam klären: absolut = uneingeschränkt. Dann Ansage: „Jetzt schreibt ihr von der Folie ab: Merksatz 2. Rahmt ihn rot ein.“", 21);

{
  const { s, top } = taskSlide("Übung 2", "Absolutismus – ja oder nein?", ["Schreibe Ü2 an den Rand.", "Schreibe Nummer + „Absolutismus“ oder „kein Absolutismus“.", "Begründe mit dem Test: Kann jemand den König stoppen?"], "3 Minuten · allein · leise",
    "Lösung: 1 Absolutismus (keine Kontrolle) · 2 kein Absolutismus (Parlament kann den König stoppen) · 3 Absolutismus (Beamte sind dem König zum Gehorsam verpflichtet, er kontrolliert sie) · 4 kein Absolutismus (Wahl = Demokratie, Macht ist begrenzt). Kontrolle kurz: „Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“ Dann: „Vergleicht und verbessert mit Grün.“");
  sentences(s, [
    "Der König entscheidet allein über Gesetze, Krieg und Steuern.",
    "Ein Parlament kann Steuererhöhungen des Königs verbieten.",
    "Beamte handeln im Namen des Königs und müssen ihm gehorchen.",
    "Die Bürger wählen alle vier Jahre ihre Regierung.",
  ], top, { size: 22, rh: 0.8 });
}

// ============ Im Absolutismus fühle ich mich ============
{
  const { s, top } = taskSlide("Übung 3", "Im Absolutismus fühle ich mich …", ["Schreibe Ü3 an den Rand. Betrachte die Karikatur.", "Wähle einen Stand und beende den Satz: „Im Absolutismus fühle ich mich …, weil …“", "Nenne ein Recht und eine Pflicht deines Standes."], "5 Minuten · allein · leise",
    "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ Die Karikatur gibt Hinweise auf Vorrechte und Lasten der Stände (hier: Karikatur Drei Stände, 1789; die Karikatur mit Sprechblasen kann die Bilddatei in quellcode/praesentation.js ersetzen). Danach 3 Schüler vorlesen lassen (verschiedene Stände): „Wer drankommt, liest vor. Alle prüfen mit.“ Mögliche Antwort: „Im Absolutismus fühle ich mich als Adeliger privilegiert, weil ich keine Steuern zahlen muss und allein jagen darf. Dafür muss ich mich ehrenhaft verhalten und Verwaltungsaufgaben übernehmen.“ / „… als Bauer ausgenutzt, weil ich Steuern, hohe Abgaben und Pachtgebühren zahlen muss und von der Mitbestimmung ausgeschlossen bin.“ Prüfen: Wird ein Vorrecht und eine Pflicht genannt? Passen sie zum Stand? Typische Fehler: Adel als Steuerzahler, Klerus mit Wehrdienst (befreit), Dritter Stand (Bauern und Bürger) mit Vorrechten. Dann: „Verbessert mit Grün.“",
    "Fertig? Stift hinlegen.", 8.9);
  s.addImage({ path: "../material/karikatur_drei_staende.jpg", x: 9.7, y: 1.5, w: 3.05, h: 3.78 });
  box(s, 0.6, top + 0.1, 12.15, 0.8, LIGHT);
  txt(s, "Klerus (Geistliche) · Adel · Dritter Stand (Bauern und Bürger)", { x: 0.9, y: top + 0.1, w: 11.6, h: 0.8, fontFace: HEAD, fontSize: 24, valign: "middle" });
}

// ============ Schluss ============
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "ZUM SCHLUSS · STIFTE LIEGEN", { x: 0.6, y: 0.35, w: 10, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  txt(s, "So lernst du bis zum Test", { x: 0.6, y: 0.65, w: 12.1, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true });
  const items = ["Merksatz 2 auswendig: Was bedeutet Absolutismus?", "Lückentext noch einmal ohne Wortspeicher ausfüllen", "Tabelle Ständegesellschaft aus dem Kopf aufschreiben", "Deinen Ü3-Satz zu einem zweiten Stand noch einmal schreiben"];
  items.forEach((x, i) => {
    const y = 1.7 + i * 1.05;
    box(s, 0.6, y, 12.15, 0.9, i % 2 ? LIGHT : LIGHT2);
    numCircle(s, i + 1, 0.85, y + 0.2, 0.5, RED);
    txt(s, x, { x: 1.6, y, w: 10.9, h: 0.9, fontFace: HEAD, fontSize: 24, valign: "middle" });
  });
  bar(s, [{ text: "Test: ", options: { bold: true } }, { text: "nächste Woche · Ständegesellschaft · Absolutismus · fünf Säulen", options: {} }]);
  footer(s);
  s.addNotes("Ansage: „Stifte liegen. Das ist eure Lernliste für den Test. Fragen zum Test?“ Die Liste bewusst kurz halten; Testtermin noch einmal nennen.");
}

pres.writeFile({ fileName: "../Praesentation_Absolutismus_Wiederholung.pptx" }).then(() => console.log("ok"));
