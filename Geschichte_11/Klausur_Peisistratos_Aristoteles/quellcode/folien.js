// Folien Teil 2 der Stunde: neue Übungsklausur „Peisistratos“ – TATTEA bestimmen und Einleitung formulieren. AVIVA, alle Texte ≥ 27 pt, Foliennummern automatisch.
const pptxgen = require("pptxgenjs");
const D = require("./data.js");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", green: "3C783C", grey: "B5AC9E" };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Übungsklausur Peisistratos – TATTEA und Einleitung";
const H = "Cambria", B = "Calibri", MIN = 27, TITLE = 36, PH = ["A", "V", "I", "V", "A"];
const txt = (s, text, o) => s.addText(text, { isTextBox: true, fontFace: B, color: C.dark, margin: 0, valign: "top", fontSize: MIN, ...o });
const num = s => { s.slideNumber = { x: 9.3, y: 5.3, w: 0.4, h: 0.25, fontFace: B, fontSize: 12, color: C.soft, align: "right" }; };
function aviva(s, k) { PH.forEach((l, i) => s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 7.15 + i * 0.5, y: 0.12, w: 0.44, h: 0.44,
  fill: { color: i === k ? C.terra : C.sand }, color: i === k ? C.white : C.soft, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 })); }
const tag = (s, t, fill = C.dark, color = C.white) => s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 7.15, y: 0.66, w: 2.4, h: 0.5,
  fill: { color: fill }, color, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
const loesung = s => tag(s, "Lösung", C.gold, C.dark);
const title = (s, t, k) => { aviva(s, k); txt(s, t, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle" }); };
function panel(s, t, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: o.fill || C.white }, line: { color: o.line || C.terra, width: o.lw || 2 } });
  txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color: o.color || C.dark, bold: o.bold, valign: "middle", align: o.align, italic: o.italic });
}
function steps(s, items, y, gap) { items.forEach((t, i) => {
  s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + i * gap, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
  txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: 0.55, valign: "middle" }); }); }
function rows(s, arr, o = {}) { const lw = o.lw || 2.4, y0 = o.y0 || 1.35, rh = o.rh || 1.15, gap = 0.12;
  arr.forEach(([l, r], i) => { const y = y0 + i * (rh + gap);
    s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 0.45, y, w: lw, h: rh, fill: { color: C.sand }, color: C.dark, fontFace: B, fontSize: MIN, bold: true, valign: "middle", margin: [0, 6, 0, 10] });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.55 + lw, y, w: 9.0 - lw, h: rh, rectRadius: 0.06, fill: { color: C.white }, line: { color: r ? C.terra : C.soft, width: 1.5, dashType: r ? "solid" : "dash" } });
    if (r) txt(s, r, { x: 0.75 + lw, y, w: 8.6 - lw, h: rh, valign: "middle", color: o.color || C.terra, bold: o.bold !== false }); }); }
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };
const rich = t => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => x.startsWith("**") ? { text: x.slice(2, -2), options: { bold: true, color: C.terra } } : { text: x, options: { color: C.dark } });

// 1 Titel
{ const s = slide(C.dark);
  txt(s, "TEIL 2 · NEUE ÜBUNGSKLAUSUR", { x: 0.7, y: 0.9, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 2 });
  txt(s, D.themenfrage, { x: 0.7, y: 1.45, w: 8.6, h: 1.9, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Heute: TATTEA und Einleitung", { x: 0.7, y: 3.75, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Material: Uebungsklausur_Peisistratos (A4 doppelseitig, 1 pro Schüler), TATTEA_Einleitung_A5 (2 × A5 auf A4 quer, schneiden). Übersicht „So interpretierst du eine Textquelle“ darf auf dem Tisch liegen.\nGelenk zu Teil 1: „Eben habt ihr das Ende einer Interpretation geübt – das Urteil. Jetzt nehmen wir uns den Anfang vor: die Einleitung, an einer neuen Klausur.“"); }

// 2 Ziel und Ablauf
{ const s = slide(); title(s, "Ziel und Ablauf", 0); tag(s, "Ankommen");
  txt(s, "Ich kann zu einer neuen Quelle TATTEA bestimmen und die Einleitung schreiben.", { x: 0.45, y: 1.3, w: 9.1, h: 0.95, bold: true, color: C.terra, valign: "middle" });
  steps(s, ["TATTEA wiederholen", "Aufgabe und Quelle lesen", "TATTEA bestimmen", "Einleitung schreiben", "Fehler finden"], 2.4, 0.58);
  num(s); s.addNotes("1 Min. Ziel vorlesen. Mindestziel: TATTEA vollständig in Stichworten. Die Einleitung bekommt ihren Platz, ‚Fehler finden‘ (Auswerten) wird nie gestrichen."); }

// 3–4 Vorwissen TATTEA
{ const s = slide(); title(s, "Wofür steht TATTEA?", 1); tag(s, "allein · 1 Min.");
  "TATTEA".split("").forEach((l, i) => s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 0.6 + i * 1.5, y: 1.6, w: 1.2, h: 1.2,
    fill: { color: C.white }, line: { color: C.terra, width: 2.5 }, color: C.terra, fontFace: H, fontSize: 48, bold: true, align: "center", valign: "middle", margin: 0 }));
  steps(s, ["Schreibe die sechs Wörter ins Heft – ohne Übersicht."], 3.4, 0.7);
  num(s); s.addNotes("2–4 Min. Wissens-Check (CLAUDE.md, 17.11.): erst abrufen, dann prüfen. Nach 1 Min. Stichprobe bei 4 Schülern unterschiedlichen Niveaus (je ein Buchstabe). Danach Folie 4. Fehlt viel, Folie 4 länger stehen lassen und abschreiben lassen."); }
{ const s = slide(); title(s, "TATTEA", 1); loesung(s);
  [["Titel", "Wie heißt der Text?"], ["Autor", "Wer schrieb ihn?"], ["Textsorte", "Was für ein Text?"], ["Thema", "Worum geht es?"], ["Entstehung", "Wann und wo?"], ["Adressat", "Für wen?"]].forEach(([a, b], i) => {
    const y = 1.3 + i * 0.66;
    txt(s, a, { x: 0.45, y, w: 3.0, h: 0.6, bold: true, color: C.terra, valign: "middle" });
    txt(s, b, { x: 3.5, y, w: 6.0, h: 0.6, valign: "middle" }); });
  num(s); s.addNotes("Kurz vergleichen lassen, Fehlendes ergänzen."); }

// 5 Begriffe der Themenfrage
{ const s = slide(); title(s, "Die Themenfrage verstehen", 2); tag(s, "Begriffe");
  panel(s, "Fürsorge = sich um das Wohl der Menschen kümmern", 0.45, 1.4, 9.1, 1.0, { line: C.green });
  panel(s, "Machtsicherung = dafür sorgen, dass man an der Macht bleibt", 0.45, 2.6, 9.1, 1.0, { line: C.terra });
  panel(s, "Die Frage: Half Peisistratos den Menschen – oder half er vor allem sich selbst?", 0.45, 3.85, 9.1, 1.2, { fill: C.dark, line: C.dark, color: C.gold, bold: true });
  num(s); s.addNotes("4–6 Min. Direkt erklären. Die Themenfrage ist in dieser Klausur VORGEGEBEN – sie muss nicht selbst formuliert werden, kommt aber an das Ende der Einleitung.\nPrüffrage (2 Schüler): „Erkläre mit eigenen Worten: Was wäre Machtsicherung?“"); }

// 6 Fundstellen
{ const s = slide(); title(s, "Wo finde ich TATTEA?", 2); tag(s, "zuhören");
  rows(s, [["Infokasten", "Titel · Autor · Textsorte"], ["Nachweis", "Entstehung (unter dem Text)"], ["Text lesen", "Thema"], ["erschließen", "Adressat: Wer liest so etwas?"]], { rh: 0.82, y0: 1.35 });
  num(s); s.addNotes("6–8 Min. Auf dem Klausurblatt zeigen: kursiver Infokasten über M1, Nachweis klein unter dem Text. Der Adressat steht nirgends – er wird erschlossen."); }

// 7 Warnung vor dem Fehler
{ const s = slide(); title(s, "Achtung: zwei Fallen", 2); tag(s, "Warnung", C.terra);
  panel(s, "Autor: umstritten! → „Als Verfasser gilt Aristoteles …“", 0.45, 1.4, 9.1, 1.2, { line: C.terra });
  panel(s, "Entstehung: um 330 v. Chr. – Peisistratos herrschte 561–527 v. Chr.", 0.45, 2.8, 9.1, 1.2, { line: C.terra });
  panel(s, "→ rund 200 Jahre später: kein Augenzeuge!", 0.45, 4.15, 9.1, 0.9, { fill: C.dark, line: C.dark, color: C.gold, bold: true });
  num(s); s.addNotes("8–9 Min. Beide Fallen sind später wichtig für die Glaubwürdigkeit im Urteil.\nPrüffrage (2–3 Schüler): „Warum ist ‚Aristoteles berichtet als Zeitzeuge‘ falsch?“"); }

// 8 Vormachen
{ const s = slide(); title(s, "So geht's – ein Beispiel", 2); tag(s, "Vormachen");
  rows(s, [["Titel", "„Athenaion politeia“ (Der Staat der Athener), Kap. 16"], ["Autor", "umstritten; in der Antike Aristoteles zugeschrieben"]], { rh: 1.4, y0: 1.5 });
  num(s); s.addNotes("9–11 Min. Laut denken: „Ich suche den Titel. Über dem Text steht ‚M1: Aristoteles über …‘ – das ist nur die Überschrift der Klausur. Den echten Titel finde ich im Infokasten: Athenaion politeia. Im Nachweis steht noch ‚Kapitel 16‘ – das gehört dazu, weil es ein Auszug ist.\nAutor: Im Infokasten steht ‚umstritten‘ und ‚in der Antike Aristoteles zugeschrieben‘. Also schreibe ich nicht einfach ‚Aristoteles‘, sondern genau das.“"); }

// 9 Auftrag 1
{ const s = slide(); title(s, "TATTEA bestimmen", 3); tag(s, "allein · 8 Min.");
  steps(s, ["Lies den Infokasten und den Nachweis.", "Lies den Text (Z. 1–16).", "Notiere Stichworte in der Tabelle (Teil 1).", "Titel und Autor hast du schon."], 1.4, 0.8);
  num(s); s.addNotes("11–19 Min. A5-Blatt austeilen. Umhergehen. Impulse gestuft:\n1. „Wo steht es? Infokasten oder Nachweis?“\n2. Auf die Zeile zeigen.\n3. Satzanfang: „Der Text entstand um …“\nThema: Wer nur ‚Peisistratos‘ schreibt → „Was genau über ihn?“ (wie er regierte, warum er sich halten konnte)."); }

// 10 Grundlage an die Wand
{ const s = slide(); title(s, "Infokasten und Nachweis", 3); tag(s, "Grundlage");
  txt(s, "„Athenaion politeia“ (dt.: „Der Staat der Athener“): historischer Überblick über Athen von Solon bis 403 v. Chr. Autor umstritten; in der Antike meist Aristoteles zugeschrieben.", { x: 0.45, y: 1.35, w: 9.1, h: 2.4, italic: true });
  txt(s, "Nachweis: Kapitel 16, verfasst um 330 v. Chr.", { x: 0.45, y: 4.0, w: 9.1, h: 0.6, bold: true, color: C.terra });
  num(s); s.addNotes("Kurzfassung des Infokastens; bleibt während der Arbeitsphase stehen. Den Text selbst lesen die Schüler auf dem Klausurblatt."); }

// 11–12 Lösung TATTEA
{ const s = slide(); title(s, "TATTEA (1)", 3); loesung(s);
  rows(s, [["Titel", "„Athenaion politeia“, Kapitel 16"], ["Autor", "umstritten; Aristoteles zugeschrieben"], ["Textsorte", "Geschichtsschrift (historischer Überblick)"]], { rh: 1.05 });
  num(s); s.addNotes("19–22 Min. Vergleichen, Fehlendes ergänzen. Stichprobe: 3 Schüler lesen je eine Zeile vor, bevor die Lösung aufgedeckt wird."); }
{ const s = slide(); title(s, "TATTEA (2)", 3); loesung(s);
  rows(s, [["Thema", "Herrschaft des Peisistratos: wie er regierte, warum er sich halten konnte"], ["Entstehung", "um 330 v. Chr., Athen – 200 Jahre später"], ["Adressat", "gebildete Leser in Athen"]], { rh: 1.05 });
  num(s); s.addNotes("Thema: Auch richtig ist „wie Peisistratos seine Herrschaft führte und sicherte“. Falsch: eine Wertung („wie gut Peisistratos war“) – die gehört ins Urteil.\nAdressat: auch richtig „Schüler und Gelehrte (Philosophenschule)“, wenn begründet."); }

// 13 Auftrag 2
{ const s = slide(); title(s, "Einleitung schreiben", 3); tag(s, "allein · 10 Min.");
  steps(s, ["Schreibe vier Sätze ins Heft (Teil 2).", "Nutze die Satzanfänge und hake sie ab.", "Achte auf die zwei Fallen.", "Ende mit der Themenfrage."], 1.4, 0.8);
  num(s); s.addNotes("22–32 Min. Umhergehen, zuerst zu den Schwächeren. Stichproben im Heft: Steht ‚gilt als‘ beim Autor? Steht ‚200 Jahre später‘? Endet die Einleitung mit der Themenfrage?\nImpuls bei Satz 4: „Die Untersuchung geht der Frage nach, ob Peisistratos’ Politik …“"); }

// 14–15 Muster-Einleitung
{ const s = slide(); title(s, "Einleitung – Muster (1)", 3); loesung(s);
  txt(s, rich("**Bei der vorliegenden Quelle handelt es sich um einen Auszug aus** der Geschichtsschrift „Athenaion politeia“, Kapitel 16. **Als Verfasser gilt** der Philosoph Aristoteles; **allerdings ist umstritten,** ob er sie wirklich schrieb."),
    { x: 0.45, y: 1.35, w: 9.1, h: 3.9, lineSpacingMultiple: 1.05 });
  num(s); s.addNotes("32–35 Min. Vollständige Muster-Einleitung in der Handreichung. Schüler vergleichen mit ihrem Heft und ergänzen in einer anderen Farbe."); }
{ const s = slide(); title(s, "Einleitung – Muster (2)", 3); loesung(s);
  txt(s, rich("**Die Schrift entstand um** 330 v. Chr. in Athen, **also rund** 200 Jahre **nach** Peisistratos. **Sie richtet sich an** gebildete Leser. **Sie behandelt** die Herrschaft des Peisistratos. **Die Untersuchung geht der Frage nach, ob** seine Politik Fürsorge oder Machtsicherung war."),
    { x: 0.45, y: 1.35, w: 9.1, h: 3.9, lineSpacingMultiple: 1.05 });
  num(s); }

// 16–17 Auswerten
{ const s = slide(); title(s, "Was rätst du ihm?", 4); tag(s, "zu zweit");
  panel(s, "„Aristoteles berichtet als Augenzeuge, dass Peisistratos ein guter Herrscher war.“", 0.45, 1.35, 9.1, 1.5, { line: C.terra, italic: true });
  txt(s, "So beginnt eine Einleitung. Findet zwei Fehler.", { x: 0.45, y: 3.05, w: 9.1, h: 0.6 });
  panel(s, "„Ich rate dir, …, weil …“", 0.45, 3.95, 9.1, 0.9, { fill: C.dark, line: C.dark, color: C.gold, bold: true, align: "center" });
  num(s); s.addNotes("35–41 Min. Nie streichen. Zu zweit 2 Min., dann Blitzlicht: 3–4 Paare gezielt aufrufen, je ein Fehler mit ‚weil‘."); }
{ const s = slide(); title(s, "Die Fehler", 4); loesung(s);
  rows(s, [["Autor", "umstritten → „Als Verfasser gilt …“"], ["Zeit", "330 v. Chr. → kein Augenzeuge"], ["Thema", "keine Wertung → „Sie behandelt …“"]], { rh: 0.95 });
  panel(s, "Ziel erreicht: TATTEA und Einleitung?", 0.45, 4.6, 9.1, 0.7, { line: C.terra, color: C.terra, bold: true });
  num(s); s.addNotes("41–45 Min. Außerdem fehlen Titel und Textsorte. Zurück zum Ziel: Kurze Stichprobe – „Welcher Teil der Einleitung fiel dir am schwersten?“ (Sachfrage, keine Selbsteinschätzung).\nAusblick: Nächste Stunde geht es mit der Position des Verfassers weiter."); }

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
