const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Wiederholung Subjunktionen";
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

// Wiederholungsfolie Subjunktionen (Einstieg/Puffer): Aufgabe + Lösung
FOOT = "Sprachtraining · Wiederholung Subjunktionen · Deutsch 7A";
const runs = (parts, size = 22) => parts.map(([x, k]) => ({ text: x, options: Object.assign({ fontSize: size }, k === "sub" ? { bold: true, color: VIO } : k === "fin" ? { bold: true, underline: { style: "sng" } } : k === "hv" ? { bold: true, color: RED, underline: { style: "sng" } } : k === "mv" ? { bold: true, color: GREEN, underline: { style: "sng" } } : k === "m" ? { color: MUTED } : {}) }));
const ITEMS = [
  ["1", "Paddy versteckt sich im Bett. Er hat Angst vor dem Schotten.", "weil", [["Paddy versteckt sich im Bett, ", ""], ["weil", "sub"], [" er Angst vor dem Schotten ", ""], ["hat", "fin"], [".", ""]]],
  ["2", "Der Trainer sagt. Wir müssen heute länger laufen.", "dass", [["Der Trainer sagt, ", ""], ["dass", "sub"], [" wir heute länger laufen ", ""], ["müssen", "mv"], [".", ""]]],
  ["* 3", "Der Schotte flieht. Er hat die riesige Nase gesehen.", "nachdem", [["Der Schotte flieht, ", ""], ["nachdem", "sub"], [" er die riesige Nase gesehen ", ""], ["hat", "hv"], [".", ""]]],
];
[false, true].forEach((show) => {
  const s = base("Vorwissen aktivieren · Wiederholung Sprachtraining", show ? "Lösung" : "Verbinde die Sätze mit der Subjunktion", show
    ? "Mit Grün verbessern. Nachfragen: Wo steht das finite Verb? (am Ende) – Was steht vor der Subjunktion? (Komma). Bei 2 ist das Modalverb finit (grün), bei 3 das Hilfsverb (rot) – beide wandern ans Ende."
    : "Allein im Heft, 4 Minuten. Erinnerung an den Merksatz vom Mittwoch: Subjunktion leitet den Nebensatz ein, finites Verb am Ende, Komma davor. Lösung auf der nächsten Folie.");
  if (!show) {
    [["1", "Verbinde die Sätze mit der Subjunktion in Klammern. Schreibe den ganzen Satz ins Heft."], ["2", "Unterstreiche das finite Verb. Kreise die Subjunktion ein."]].forEach(([n, x], i) => {
      num(s, n, 0.6, 1.55 + i * 0.55, VIO, 0.42);
      txt(s, x, { x: 1.2, y: 1.58 + i * 0.55, w: 11.4, h: 0.45, fontSize: 18, bold: true });
    });
  }
  ITEMS.forEach(([n, q, sj, sol], i) => {
    const y = (show ? 1.6 : 2.85) + i * (show ? 1.25 : 0.95);
    box(s, 0.6, y, 12.1, show ? 1.05 : 0.8, i === 2 ? WHITE : LIGHT, i === 2 ? { line: "C9D3EA" } : {});
    txt(s, n, { x: 0.8, y: y + (show ? 0.3 : 0.2), w: 0.8, h: 0.45, fontSize: 20, bold: true, color: VIO });
    if (show) txt(s, runs(sol, 24), { x: 1.6, y: y + 0.27, w: 10.9, h: 0.55, fontFace: HEAD });
    else {
      txt(s, q, { x: 1.6, y: y + 0.2, w: 9.0, h: 0.45, fontSize: 20 });
      txt(s, "(" + sj + ")", { x: 10.6, y: y + 0.2, w: 1.9, h: 0.45, fontSize: 20, bold: true, color: VIO, align: "right" });
    }
  });
  if (show) {
    box(s, 0.6, 5.5, 12.1, 0.75, LIGHT2);
    txt(s, [{ text: "Regel: ", options: { bold: true, color: RED } }, { text: "Subjunktion → Nebensatz → finites Verb am Ende, davor ein Komma." }, { text: "   (Modalverb grün, Hilfsverb rot)", options: { color: MUTED } }], { x: 0.9, y: 5.65, w: 11.6, h: 0.5, fontSize: 18 });
  } else {
    box(s, 0.6, 5.85, 12.1, 0.55, LIGHT2);
    txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "4 Minuten · allein im Heft · Nr. 3 mit *      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 5.93, w: 11.7, h: 0.4, fontSize: 16 });
  }
});
pres.writeFile({ fileName: "Folie_Wiederholung_Subjunktionen.pptx" });
