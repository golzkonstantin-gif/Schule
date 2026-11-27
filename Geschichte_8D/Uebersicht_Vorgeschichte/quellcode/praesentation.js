const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Übersicht Vorgeschichte (8D)";
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF", RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B";
const HEAD = "Cambria", BODY = "Calibri", AVIVA = ["A", "V", "I", "V", "A"];
let pageNo = 0;
const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 18, color: NAVY, margin: 0, valign: "top" }, o));
const box = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: o.line ? { color: o.line, width: 1.5 } : { type: "none" } });
const circ = (s, n, x, y, d, col) => { s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" } }); txt(s, String(n), { x, y, w: d, h: d, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" }); };
function slide(kicker, title, notes) {
  const s = pres.addSlide(); pageNo++; s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.35, w: 9, h: 0.3, fontSize: 13, color: MUTED, charSpacing: 2, bold: true });
  AVIVA.forEach((l, i) => { const x = 10.0 + i * 0.55; s.addShape(pres.shapes.OVAL, { x, y: 0.3, w: 0.44, h: 0.44, fill: { color: i === 1 ? RED : LIGHT2 }, line: { type: "none" } }); txt(s, l, { x, y: 0.3, w: 0.44, h: 0.44, fontSize: 14, bold: true, color: i === 1 ? WHITE : MUTED, align: "center", valign: "middle" }); });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.72, w: 1.2, h: 0.05, fill: { color: GOLD }, line: { type: "none" } });
  txt(s, title, { x: 0.6, y: 0.85, w: 12.1, h: 0.7, fontFace: HEAD, fontSize: 34, bold: true });
  txt(s, "Übersicht · Geschichte 8D", { x: 0.6, y: 7.12, w: 6, h: 0.25, fontSize: 10, color: MUTED });
  if (notes) s.addNotes(notes);
  return s;
}
const bar = (s, runs) => { box(s, 0.6, 6.25, 12.15, 0.8, NAVY); s.addText(runs, { isTextBox: true, x: 0.9, y: 6.25, w: 11.6, h: 0.8, fontFace: BODY, fontSize: 27, color: WHITE, valign: "middle", margin: 0 }); };
const LET = require("./daten.js").LETTER;
const SHORT = [["1715", "Ludwig XIV. stirbt, hohe Schulden"], ["1778 bis 1783", "Krieg in Amerika, noch mehr Schulden"], ["1787", "Steuerplan scheitert"], ["1788", "Missernte, Staatsbankrott droht"], ["August 1788", "König kündigt Generalstände an"], ["Frühjahr 1789", "Wahl, Beschwerdehefte"], ["5. Mai 1789", "Generalstände werden eröffnet"], ["17. Juni 1789", "Nationalversammlung"], ["20. Juni 1789", "Ballhausschwur"], ["Anfang Juli 1789", "Soldaten um Paris"], ["14. Juli 1789", "Sturm auf die Bastille"], ["Sommer 1789", "Bauern greifen Schlösser an"], ["4. August 1789", "Vorrechte abgeschafft"], ["26. August 1789", "Erklärung der Menschenrechte"]];
// Fortschrittsleiste: 14 Punkte, bis "upto" gefüllt
function dots(s, upto, from) { for (let i = 1; i <= 14; i++) { const x = 0.6 + (i - 1) * 0.87; s.addShape(pres.shapes.OVAL, { x, y: 1.62, w: 0.3, h: 0.3, fill: { color: i >= from && i <= upto ? RED : i < from ? NAVY : LIGHT2 }, line: { type: "none" } }); } s.addShape(pres.shapes.LINE, { x: 0.75, y: 1.77, w: 11.3, h: 0, line: { color: ICE, width: 1 } }); }

// 1 Auftrag
{
  const s = slide("Vorwissen aktivieren · Übung 1", "Die Vorgeschichte in der richtigen Reihenfolge",
    "Dient dem Ziel „Ich kann die Ereignisse von 1715 bis zur Menschenrechtserklärung in die richtige Reihenfolge bringen“. Ansage: „Stifte liegen. Ihr hört nur zu. Jetzt arbeitet ihr zu zweit. Flüsterstufe. Nur dein Partner hört dich. Die Aufgabe steht an der Tafel.“ Jede Person hat Blatt und Schnipsel. Legt, nicht kleben. Durch die Reihen gehen, nichts verraten. Wer fertig ist: Stift hinlegen.");
  box(s, 0.6, 1.7, 12.15, 3.3, LIGHT2);
  ["Lege die 14 Schnipsel in die richtige Reihenfolge.", "Sprich mit deinem Partner: Was war zuerst, was danach?", "Klebe noch nicht auf. Wir kontrollieren gemeinsam."].forEach((st, i) => { circ(s, i + 1, 0.85, 1.95 + i * 1.0, 0.5, NAVY); txt(s, st, { x: 1.65, y: 1.95 + i * 1.0, w: 10.8, h: 0.9, fontSize: 27, bold: true }); });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "3 Min. · zu zweit · Flüsterstufe      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}
// 2 bis 4 Kontrolle
const groups = [[1, 5, "Kontrolle 1: Felder 1 bis 5", "Aufdecken, Feld für Feld. Frage bei jedem Feld: „Warum steht es vor dem nächsten?“ Stichproben 2 bis 3 Personen. Feld 2 und 3 stehen oft vertauscht: 1787 (Steuerplan) kommt nach dem Krieg in Amerika. Verbessern mit Grün (nur Zahlen am Blatt, noch nichts kleben)."], [6, 10, "Kontrolle 2: Felder 6 bis 10", "Feld 7 (5. Mai) vor Feld 8 (17. Juni), Feld 9 (20. Juni) direkt danach: Warum? Der König erkennt die Versammlung nicht an. Feld 10 (Anfang Juli): Soldaten, weil er Angst vor der Versammlung hat."], [11, 14, "Kontrolle 3: Felder 11 bis 14", "Feld 11 Bastille, Feld 12 Aufstände im Sommer, Feld 13 (4. August) und 14 (26. August). Frage: „Was haben Bastille und Aufstände mit dem 4. August zu tun?“ Danach: Alles gleich? Dann jetzt aufkleben."]];
groups.forEach(([a, b, title, note]) => {
  const s = slide("Vorwissen aktivieren · Kontrolle", title, "Ansage: „Stifte liegen. Ihr hört nur zu. Vergleicht und verbessert mit Grün.“ " + note);
  dots(s, b, a);
  const n = b - a + 1, rh = 0.78;
  for (let i = 0; i < n; i++) {
    const k = a + i - 1, y = 2.1 + i * rh;
    box(s, 0.6, y, 12.15, rh - 0.1, i % 2 ? WHITE : LIGHT, { line: LIGHT2 });
    circ(s, k + 1, 0.75, y + 0.12, 0.5, NAVY);
    txt(s, SHORT[k][0], { x: 1.5, y, w: 3.6, h: rh - 0.1, fontSize: 27, bold: true, color: RED, valign: "middle" });
    circ(s, LET[k + 1], 12.1, y + 0.09, 0.5, RED); txt(s, SHORT[k][1], { x: 5.2, y, w: 6.8, h: rh - 0.1, fontSize: 27, fontFace: HEAD, valign: "middle" });
  }
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Vergleiche und verbessere mit Grün.", options: {} }]);
});
// 5 Kleben und deshalb
{
  const s = slide("Vorwissen aktivieren · Zusammenhang", "Jetzt kleben, dann: Wie hängt das zusammen?",
    "Ansage: „Jetzt klebt ihr auf. Stifte liegen erst danach.“ (1 Min.) Danach: Ihr zeichnet drei Pfeile ein, wir sprechen sie gemeinsam. Erwartung: 3 zu 5: Das Geld fehlt und die Privilegierten zahlen nicht, deshalb ruft der König die Generalstände. 8 zu 9: Der König erkennt die Nationalversammlung nicht an, deshalb sperrt er den Saal, deshalb der Ballhausschwur. 11 zu 13: Die Bastille und die Aufstände zeigen, dass der König nicht mehr alles kontrolliert, deshalb geben Klerus und Adel nach. Pfeile gemeinsam an die Folie zeichnen, die Klasse zeichnet mit (1 Minute).");
  [["3  1787 →  5  August 1788", "Warum ruft der König die Generalstände?"], ["8  17. Juni →  9  20. Juni", "Warum der Ballhausschwur?"], ["11  Bastille →  13  4. August", "Warum geben Klerus und Adel nach?"]].forEach(([l, q], i) => {
    const y = 1.8 + i * 1.45;
    box(s, 0.6, y, 12.15, 1.25, LIGHT2); txt(s, l, { x: 0.9, y, w: 5.4, h: 1.25, fontSize: 27, bold: true, color: RED, valign: "middle" }); txt(s, "deshalb: " + q, { x: 6.4, y, w: 6.2, h: 1.25, fontSize: 27, fontFace: HEAD, valign: "middle" });
  });
  bar(s, [{ text: "Jetzt: ", options: { bold: true } }, { text: "Zeichne die drei Pfeile mit „deshalb“.", options: {} }]);
}
// 6 Pruefaufgabe
{
  const s = slide("Vorwissen aktivieren · Prüfaufgabe", "Warum gaben Klerus und Adel am 4. August nach?",
    "Dient dem Ziel „erklären, warum die Ereignisse aufeinander folgen“. Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht an der Tafel.“ 1 Minute ins Heft. Stichprobe bei 2 bis 3 Personen. Erwartung (Auswahl): Klerus und Adel gaben nach, weil die Bauern Schlösser angriffen und der König nach der Bastille nicht mehr alles kontrollierte. Die Bauern griffen Schlösser an, deshalb … Bei Zeitnot nur mündlich. Danach: Gelenk zur Stunde: „Das war der Weg bis 1789. Heute: für wen galten diese Rechte?“");
  box(s, 0.6, 1.7, 12.15, 3.6, LIGHT2);
  ["Schreibe zwei Sätze ins Heft.", "Klerus und Adel gaben nach, weil …", "Die Bauern griffen Schlösser an, deshalb …"].forEach((st, i) => { circ(s, i + 1, 0.85, 1.95 + i * 1.05, 0.5, i === 0 ? NAVY : RED); txt(s, st, { x: 1.65, y: 1.95 + i * 1.05, w: 10.8, h: 0.9, fontSize: 27, bold: i === 0, fontFace: i === 0 ? BODY : HEAD }); });
  bar(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "1 Min. · allein · leise      ", options: {} }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: ICE } }]);
}
pres.writeFile({ fileName: "../Praesentation_Uebersicht_Vorgeschichte_8D.pptx" }).then(() => console.log("ok"));
