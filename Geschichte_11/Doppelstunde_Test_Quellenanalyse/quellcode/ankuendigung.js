// Folien Testankündigung (Ende der Montagsstunde) – alle Texte für die Klasse mindestens 27 pt (Vorgabe der Lehrkraft).
const pptxgen = require("pptxgenjs");
const D = require("./data.js");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", pale: "FBF8F2", green: "3C783C", blue: "3E5C76" };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Testankündigung";
const H = "Cambria", B = "Calibri";
const MIN = 27, TITLE = 36;
const PH = ["A", "V", "I", "V", "A"];

function txt(s, text, o) { s.addText(text, { isTextBox: true, fontFace: B, color: C.dark, margin: 0, valign: "top", fontSize: MIN, ...o }); }
function num(s) { s.slideNumber = { x: 9.3, y: 5.3, w: 0.4, h: 0.25, fontFace: B, fontSize: 12, color: C.soft, align: "right" }; }
function aviva(s, k, dark = false) {
  PH.forEach((l, i) => s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 7.15 + i * 0.5, y: 0.12, w: 0.44, h: 0.44,
    fill: { color: i === k ? C.terra : (dark ? "3A3024" : C.sand) }, color: i === k ? C.white : (dark ? "8C8170" : C.soft), fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 }));
}
function tag(s, t, fill = C.dark, color = C.white) {
  s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 7.15, y: 0.66, w: 2.4, h: 0.5, fill: { color: fill }, color, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
}
function title(s, t, k, o = {}) { aviva(s, k, o.dark); txt(s, t, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle", color: o.dark ? C.white : C.dark }); }
function card(s, x, y, w, h, fill = C.sand) { s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } }); }
function boxText(s, t, x, y, w, h, fill, color, o = {}) { card(s, x, y, w, h, fill); txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color, valign: "middle", ...o }); }
function steps(s, items, y, gap, o = {}) {
  items.forEach((t, i) => {
    s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + i * gap, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: o.h || 0.55, valign: "middle" });
  });
}
const warn = (s, t, y = 4.15, h = 1.0) => boxText(s, "⚠ " + t, 0.45, y, 9.1, h, C.dark, C.gold, { bold: true });
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };
const ZIEL = "Ich kann eine Textquelle mit dem Schreibplan analysieren: Einleitung schreiben und Inhalt wiedergeben.";


// 1 Was kommt dran?
{ const s = slide(); txt(s, "Test am Mittwoch · 45 Min.", { x: 0.45, y: 0.15, w: 9.1, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle" });
  steps(s, ["Die Krise Athens um 600 v. Chr.", "Drakon, Solon, Peisistratos: Maßnahmen", "Herrschaftsformen mit Übersetzung", "Solons Verfassung: das Schaubild", "Sach- und Werturteil erkennen und schreiben"], 1.3, 0.78);
  num(s); s.addNotes("Am Ende der Montagsstunde zeigen (ca. 1 Min.). Testinhalte: Probleme um 600 und Schuldknechtschaft; richtig/falsch zu Drakon, Solon, Peisistratos und zu Herrschaftsformen; stummes Verfassungsschaubild (Ämter, Klassen, Rechte, Pfeile); S/W erkennen, schwaches Werturteil verbessern; Sach- und Werturteil zu Solon selbst schreiben."); }

// 2 Nicht dran / So lernst du
{ const s = slide(); txt(s, "So bereitest du dich vor", { x: 0.45, y: 0.15, w: 9.1, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle" });
  boxText(s, "Kleisthenes kommt noch nicht dran.", 0.45, 1.3, 9.1, 0.8, C.dark, C.gold, { bold: true });
  ["Zusammenfassung lesen (Athen bis Peisistratos).", "Solons Schaubild aus dem Gedächtnis zeichnen.", "Maßstäbe für Werturteile wiederholen."].forEach((t, i) =>
    boxText(s, t, 0.45, 2.3 + i * 0.95, 9.1, 0.8, i % 2 ? C.white : C.sand, C.dark));
  num(s); s.addNotes("Zusammenfassung: Klausurvorbereitung/Zusammenfassung_Athen_bis_Peisistratos.pdf (ggf. mit der Audiodatei). Maßstäbe: Gleichheit, Leistung, Schutz der Schwachen, gleiche Regeln."); }

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
