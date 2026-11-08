// Folien „Perikles – Herrschaft des Volkes oder eines Mannes?“ – alle Texte für die Klasse mindestens 27 pt (Vorgabe der Lehrkraft).
const pptxgen = require("pptxgenjs");
const D = require("./data.js");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", pale: "FBF8F2", green: "3C783C", blue: "3E5C76" };
const COL = { stadt: "C9962A", kueste: C.blue, binnen: C.green };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = D.titel;
const H = "Cambria", B = "Calibri", MIN = 27, TITLE = 36;
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
// Kasten mit Innenabstand: Form + eingerückter Text
function panel(s, t, x, y, w, h, fill, color, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: fill }, line: { color: o.line || fill, width: o.lw || 1 } });
  txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color, bold: o.bold, italic: o.italic, valign: "middle", align: o.align });
}
function label(s, t, x, y, w, h, fill, color = C.white) {
  s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x, y, w, h, fill: { color: fill }, color, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
}
function dot(s, n, x, y) {
  s.addText(String(n), { isTextBox: true, shape: pres.shapes.OVAL, x, y, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
}
function steps(s, items, y, gap) { items.forEach((t, i) => { dot(s, i + 1, 0.45, y + i * gap); txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: 0.55, valign: "middle" }); }); }
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };

// 1 Titel
{ const s = slide(C.dark);
  txt(s, "GESCHICHTE · KLASSE 11", { x: 0.7, y: 1.0, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, "Perikles", { x: 0.7, y: 1.55, w: 8.6, h: 0.8, fontSize: 44, bold: true, fontFace: H, color: C.white });
  txt(s, "Herrschaft des Volkes oder eines Mannes?", { x: 0.7, y: 2.5, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Seitentafel: Ziel und Ablauf (Folie 2). Buch und Hefter bleiben für die Wiederholung zu."); }

// 2 Ziel und Ablauf
{ const s = slide(); title(s, "Ziel und Ablauf", 0); tag(s, "2 Min.");
  panel(s, D.ziel, 0.45, 1.3, 9.1, 1.6, C.dark, C.white, { bold: true });
  ["Wiederholung: Kleisthenes", "Zwei neue Begriffe", "Text lesen, Fragen beantworten", "Lösung und Zeitleiste", "Abschlussaufgabe"].forEach((t, i) =>
    txt(s, (i + 1) + "  " + t, { x: 0.6, y: 3.0 + i * 0.47, w: 9, h: 0.47, valign: "middle" }));
  num(s); s.addNotes("0–2 Min. Ziel vorlesen (steht auch an der Seitentafel). „Heute ist das letzte Stück auf dem Weg zur Demokratie dran – und am Ende entscheidet ihr: Hat 430 v. Chr. wirklich das Volk regiert?“ Frage 8 bei Zeitnot weglassen."); }

// 3–5 Wiederholung Kleisthenes
[false, true].forEach(solved => { const s = slide(); title(s, solved ? "Lösung" : "Wiederholung: Kleisthenes", 1); tag(s, solved ? "Lösung" : "3 Min.", solved ? C.gold : C.dark, solved ? C.dark : C.white);
  D.wiederholung.forEach(([q, a], i) => {
    const y = 1.3 + i * 0.78;
    dot(s, i + 1, 0.45, y + 0.04);
    txt(s, q, { x: 1.1, y, w: solved ? 4.7 : 8.4, h: 0.6, valign: "middle" });
    if (solved) txt(s, a, { x: 5.85, y, w: 3.7, h: 0.6, valign: "middle", bold: true, color: C.terra });
  });
  num(s); s.addNotes(solved ? "5–6 Min. Vergleichen, Fehler grün verbessern. Bei Unsicherheit: Folie „Die Folge für Thrasymachos“ aus der Wiederholungspräsentation zeigen." : "2–5 Min. Buch und Hefter zu. Allein ins Heft, nur Stichworte (2 Min.), dann Stichproben bei 4–5 Schülern unterschiedlichen Niveaus (1 Min.). Abruf aus dem Gedächtnis – die Reform war beim letzten Mal schwer."); });
{ const s = slide(); title(s, "Kleisthenes: die Treppe", 1); tag(s, "Merken", C.gold, C.dark);
  [["Dorf\n(Demos)", "139", COL.binnen], ["Bezirk\n(Trittys)", "30", COL.kueste], ["Phyle", "10", C.terra], ["Rat der 500", "50 je Phyle", C.dark]].forEach(([n, z, col], i) => {
    const x = 0.45 + i * 2.32, y = 2.75 - i * 0.45;
    label(s, n, x, y, 2.2, 0.95, col);
    txt(s, z, { x, y: y + 1.0, w: 2.2, h: 0.5, align: "center", valign: "middle", color: C.muted });
  });
  panel(s, "Kein Adliger hat mehr seine Gegend als Block.", 0.45, 4.5, 9.1, 0.8, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("6–8 Min. Treppe von unten nach oben erklären lassen (Stichprobe, ein Schüler pro Stufe). Gelenkstelle: Folie 6."); }

// 6 Gelenkstelle und Leitfrage (Aha vorbereiten)
{ const s = slide(); title(s, "Und danach?", 2); tag(s, "Daumen");
  panel(s, "Nach Kleisthenes: Die höchsten Ämter haben noch die Reichen, der Areopag kontrolliert alles.", 0.45, 1.3, 9.1, 1.4, C.sand, C.dark);
  panel(s, "Athen 430 v. Chr.: Wer regiert – das Volk oder ein Mann?", 0.45, 2.9, 9.1, 1.0, C.dark, C.white, { bold: true });
  txt(s, "Daumen hoch = das Volk · Daumen runter = ein Mann", { x: 0.45, y: 4.1, w: 9.1, h: 0.7, valign: "middle", color: C.terra, bold: true });
  num(s); s.addNotes("8–10 Min. Gelenkstelle: „Kleisthenes hat den Adel als Block entmachtet. Aber noch immer konnten nur die Reichen Archonten werden, und der Areopag aus ehemaligen Archonten hatte die Aufsicht über alles. Was änderte sich danach – und wer regierte am Ende wirklich?“ Daumenabfrage (Vermutung, erlaubt in Kl. 11): Ergebnis grob an der Tafel notieren – wird am Ende mit Thukydides verglichen (Aha)."); }

// 7–8 Begriffe
[["Stratege", "griech. stratēgós = Heerführer", "Einer von zehn Heerführern Athens. Sie wurden jedes Jahr von der Volksversammlung gewählt – nicht ausgelost.", "Wer Stratege ist, hat echte Macht."],
 ["Diäten", "lat. diaeta = Tagesablauf, Tagegeld", "Geld für einen Tag im Amt, damit man keinen Verdienst verliert.", "Ohne Diäten konnten nur Reiche Ämter übernehmen."]].forEach(([w, e, d, m], i) => {
  const s = slide(); aviva(s, 2);
  txt(s, w, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: 44, bold: true, fontFace: H, valign: "middle" });
  txt(s, e, { x: 0.45, y: 1.2, w: 9, h: 0.5, italic: true, color: C.muted });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 1.85, w: 0.08, h: 1.8, fill: { color: C.terra }, line: { color: C.terra } });
  txt(s, d, { x: 0.75, y: 1.85, w: 8.8, h: 1.8, fontSize: 30 });
  panel(s, m, 0.45, 3.95, 9.1, 0.9, C.dark, C.gold, { bold: true });
  num(s); s.addNotes(i === 0 ? "10–12 Min. Begriffs-Check: Strategen und Diäten kommen im Text vor und sind noch nicht gesichert. Bekannt: Archonten, Areopag, Theten, Ostrakismos (Mystery)." : "Prüffrage an 2 Schüler: „Warum braucht ein armer Bauer Diäten, um Richter zu sein?“ (sonst fehlt ihm der Tagesverdienst)."); });

// 9 Vormachen
{ const s = slide(); title(s, "So geht's – Frage 1", 2); tag(s, "Lehrkraft");
  panel(s, "1  " + D.fragen[0][0], 0.45, 1.3, 9.1, 0.8, C.sand, C.dark, { bold: true });
  txt(s, "Im Text (Z. 4f.): „… wollten nun auch Griechenland erobern – zur Strafe für die Hilfe der Athener.“", { x: 0.45, y: 2.25, w: 9.1, h: 1.0, italic: true, color: C.muted });
  panel(s, "Heft: " + D.fragen[0][3], 0.45, 3.35, 9.1, 0.75, C.white, C.dark, { line: C.terra, lw: 2, bold: true });
  panel(s, "Nicht abschreiben – Stichworte in eigenen Worten.", 0.45, 4.35, 9.1, 0.8, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("12–14 Min. Vormachen, laut denken: „Frage 1 fragt nach dem Grund – ich suche ein Signalwort wie ‚weil‘, ‚da‘ oder ‚zur Strafe für‘. In Zeile 4 und 5 steht: Die Perser wollten Griechenland erobern – zur Strafe für die Hilfe der Athener. Welche Hilfe? Zeile 3: Athen unterstützte den Aufstand mit Schiffen. Ich schreibe nicht den Satz ab, sondern kurz: Athen half dem Aufstand in Kleinasien.“ Warnung vor dem Fehler vorlesen."); }

// 10 Auftrag
{ const s = slide(); title(s, "Lesen und beantworten", 3); tag(s, "16 Min.");
  steps(s, ["Lies den Text (Zeile 1–34).", "Beantworte Fragen 2–8 in Stichworten im Heft.", "Du kommst nicht weiter? Hand heben – ich komme zu dir.", "Letzte 3 Minuten: mit dem Nachbarn vergleichen."], 1.35, 0.85);
  num(s); s.addNotes("14–30 Min. Indirekt: Einzelarbeit mit dem Textblatt, Fragen auf dem Fragenkärtchen . Fragen auch auf den Folien 11–13 an der Wand. Umhergehen, im Heft mitlesen und individuell helfen: Hilfeblatt (Fundstelle → Impuls → Satzanfang); Stolpersteine in der Handreichung. Nach 13 Min.: „Noch 3 Minuten – jetzt vergleicht ihr zu zweit.“"); }

// 11–12 Fragen an der Wand
[[1, 4], [4, 6], [6, 8]].forEach(([a, b]) => { const s = slide(); title(s, "Die Fragen", 3); tag(s, (a + 1) + "–" + b);
  D.fragen.slice(a, b).forEach(([q], i) => { const y = 1.3 + i * 1.25;
    dot(s, a + i + 1, 0.45, y + 0.05); txt(s, q, { x: 1.1, y, w: 8.45, h: 1.1, valign: "middle" }); });
  num(s); s.addNotes("Arbeitsgrundlage an die Wand, solange gearbeitet wird."); });

// 13–14 Lösung
[[0, 4], [4, 8]].forEach(([a, b], n) => { const s = slide(); title(s, `Lösung (${n + 1}/2)`, 3); tag(s, "Lösung", C.gold, C.dark);
  D.fragen.slice(a, b).forEach(([, , , k], i) => { const y = 1.3 + i * 0.95;
    dot(s, a + i + 1, 0.45, y + 0.12); txt(s, k, { x: 1.1, y, w: 8.45, h: 0.75, valign: "middle", bold: true, color: C.terra }); });
  num(s); s.addNotes(n === 0 ? "30–34 Min. Je Frage ein Schüler vom Platz (Stichprobe, unterschiedliche Niveaus), dann aufdecken; alle ergänzen grün. Ausführliche Erwartung in der Handreichung." : "Frage 8 nur, wenn bearbeitet. Gelenkstelle: „Fragen 3, 4, 6 und 7 zeigen, wie das Volk Schritt für Schritt mehr Macht bekam.“ → Zeitleiste."); });

// 15 Zeitleiste
{ const s = slide(); title(s, "Der Weg zur Demokratie", 3); tag(s, "Sicherung");
  const Z = [["508/07", "Kleisthenes", "Bürger gemischt"], ["487/86", "Themistokles", "Archonten gelost"], ["462", "Ephialtes, Perikles", "Areopag entmachtet"], ["ab 462", "Perikles", "Diäten für Ämter"]];
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 1.75, w: 8.8, h: 0, line: { color: C.dark, width: 3, endArrowType: "triangle" } });
  Z.forEach(([j, w, m], i) => { const x = 0.45 + i * 2.32;
    label(s, j, x, 1.45, 2.15, 0.6, [C.muted, C.blue, C.terra, C.green][i]);
    txt(s, w, { x, y: 2.2, w: 2.15, h: 0.95, bold: true, align: "center", valign: "middle" });
    txt(s, m, { x, y: 3.15, w: 2.15, h: 0.95, align: "center", valign: "middle", color: C.muted }); });
  panel(s, "Jetzt spricht man von „Demokratie“.", 0.45, 4.4, 9.1, 0.85, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("34–36 Min. Tafelbild (alle übernehmen ins Heft): Zeitstrahl mit vier Stationen. Prüffrage an 2 Schüler: „Welche Maßnahme half den Ärmsten am meisten?“ (Diäten, auch die Flotte/Theten)."); }

// 16 Aha: Thukydides
{ const s = slide(C.dark); aviva(s, 4, true); tag(s, "Aha", C.gold, C.dark);
  txt(s, "Thukydides über Athen", { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, color: C.white, valign: "middle" });
  txt(s, D.zitat, { x: 0.45, y: 1.4, w: 9.1, h: 1.6, italic: true, color: C.gold, fontSize: 30 });
  txt(s, "Ein Zeitgenosse des Perikles. Und eure Daumen?", { x: 0.45, y: 3.3, w: 9.1, h: 0.8, color: "E8DFD0", valign: "middle" });
  num(s); s.addNotes("36–37 Min. Auflösung zur Daumenabfrage: Der Geschichtsschreiber Thukydides (Zeitgenosse) sah in Perikles den eigentlichen Herrscher (Lehrbuch S. 17, Randspalte). Kurz zurück zur Abstimmung vom Anfang – dann Abschlussaufgabe."); }

// 17 Abschlussaufgabe
{ const s = slide(); title(s, "Abschlussaufgabe", 4); tag(s, "4 Min.");
  panel(s, D.abschluss, 0.45, 1.3, 9.1, 1.5, C.dark, C.white, { bold: true });
  ["„Ich stimme zu / nur teilweise / nicht zu, …“", "„Zwar …, aber …“", "„… weil …“"].forEach((t, i) =>
    txt(s, t, { x: 0.6, y: 3.0 + i * 0.6, w: 9, h: 0.55, italic: true, valign: "middle" }));
  num(s); s.addNotes("37–41 Min. Prüft das Ziel: Position + Belege aus den Fragen 3–8 + „weil“. Allein ins Heft (2–3 Sätze), danach gezielt 3 Schüler unterschiedlichen Niveaus vorlesen lassen. Mögliche Argumente in der Handreichung. Diese Phase nie kürzen."); }

// 18 Beispiel
{ const s = slide(); title(s, "So könnte es lauten", 4); tag(s, "Lösung", C.gold, C.dark);
  txt(s, D.beispiel, { x: 0.45, y: 1.3, w: 9.1, h: 3.9, italic: true });
  num(s); s.addNotes("41–43 Min. Andere Positionen sind richtig, wenn sie mit Belegen begründet sind. Klausur-Hinweis: Das ist ein Sachurteil aus damaliger Sicht."); }

// 19 Ziel erreicht?
{ const s = slide(C.dark); aviva(s, 4, true);
  txt(s, "Haben wir das Ziel erreicht?", { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, color: C.white, valign: "middle" });
  txt(s, D.ziel, { x: 0.45, y: 1.5, w: 9.1, h: 1.7, color: C.gold, bold: true });
  txt(s, "Klausur 25.11.: Peisistratos – Kleisthenes – Perikles", { x: 0.45, y: 3.6, w: 9.1, h: 0.8, color: "E8DFD0" });
  num(s); s.addNotes("43–45 Min. Rückbezug aufs Ziel, Hinweis auf die Klausur."); }

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
