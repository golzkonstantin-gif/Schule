// Folien zur Probeklausur: je Abschnitt Vorgehen → leere Vorlage → ausgefüllte Lösung. Alle Texte für die Klasse mindestens 27 pt.
const pptxgen = require("pptxgenjs");
const D = require("./data.js");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", pale: "FBF8F2", green: "3C783C", blue: "3E5C76" };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Probeklausur – " + D.titel;
const H = "Cambria", B = "Calibri", MIN = 27, TITLE = 36;
const PH = ["A", "V", "I", "V", "A"];
const txt = (s, text, o) => s.addText(text, { isTextBox: true, fontFace: B, color: C.dark, margin: 0, valign: "top", fontSize: MIN, ...o });
const num = s => { s.slideNumber = { x: 9.3, y: 5.3, w: 0.4, h: 0.25, fontFace: B, fontSize: 12, color: C.soft, align: "right" }; };
function aviva(s, k, dark = false) {
  PH.forEach((l, i) => s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 7.15 + i * 0.5, y: 0.12, w: 0.44, h: 0.44,
    fill: { color: i === k ? C.terra : (dark ? "3A3024" : C.sand) }, color: i === k ? C.white : (dark ? "8C8170" : C.soft), fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 }));
}
const tag = (s, t, fill = C.dark, color = C.white) => s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 7.15, y: 0.66, w: 2.4, h: 0.5, fill: { color: fill }, color, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
const title = (s, t, k) => { aviva(s, k); txt(s, t, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle" }); };
function panel(s, t, x, y, w, h, fill, color, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: fill }, line: { color: o.line || fill, width: o.lw || 1 } });
  txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color, bold: o.bold, italic: o.italic, valign: "middle" });
}
function steps(s, items, y, gap) {
  items.forEach((t, i) => {
    s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + i * gap, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: 0.55, valign: "middle" });
  });
}
// Raster: linke Spalte Beschriftung, rechte Spalte leer oder Lösung
function raster(s, rows, o = {}) {
  const lw = o.lw || 2.6, y0 = o.y0 || 1.3, rh = o.rh || 0.62, gap = o.gap || 0.08;
  rows.forEach(([l, r], i) => {
    const y = y0 + i * (rh + gap);
    s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 0.45, y, w: lw, h: rh, fill: { color: C.sand }, color: C.dark, fontFace: B, fontSize: MIN, bold: true, align: "left", valign: "middle", margin: [0, 6, 0, 8] });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.45 + lw + 0.1, y, w: 9.1 - lw - 0.1, h: rh, rectRadius: 0.06, fill: { color: C.white }, line: { color: o.solved ? C.terra : C.soft, width: 1.5, dashType: o.solved ? "solid" : "dash" } });
    if (r) txt(s, r, { x: 0.45 + lw + 0.25, y, w: 9.1 - lw - 0.4, h: rh, valign: "middle", color: C.terra, bold: !o.plain });
  });
}
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };
const ab = (s, n, art) => tag(s, art, art === "Lösung" ? C.gold : C.dark, art === "Lösung" ? C.dark : C.white);

// 1 Titel
{ const s = slide(C.dark);
  txt(s, "GESCHICHTE · KLASSE 11 · PROBEKLAUSUR", { x: 0.7, y: 1.0, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, "Die Volksherrschaft in Athen – Entmachtung des Adels?", { x: 0.7, y: 1.55, w: 8.6, h: 1.5, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Pseudo-Xenophon: Über den Staat der Athener", { x: 0.7, y: 3.2, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Quellenblatt, Aufgabenblatt und Übersicht „So interpretierst du eine Textquelle“ liegen bereit. Lösungen: Teil 1 an der Tafelrückseite, Teile 3 und 4 als Aushang (Aushaenge_Loesungen), Teil 2 wird zusammengetragen."); }

// 2 Ablauf
{ const s = slide(); title(s, "So gehen wir vor", 0); tag(s, "Ablauf");
  steps(s, ["Themenfrage und Einleitung – allein", "Position des Verfassers – zusammentragen", "Verfassung und Gesellschaft – Aushang", "Urteil – Aushang"], 1.35, 0.8);
  panel(s, "Jeder Teil wird gleich kontrolliert.", 0.45, 4.55, 9.1, 0.75, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("Die Reihenfolge weicht bewusst von der Klausuraufgabe ab: erst Einleitung, dann Position, dann Kontext, zuletzt Urteil – so wie in den Arbeitsschritten der Übersicht. Jeder Abschnitt wird sofort kontrolliert, bevor es weitergeht."); }

// ---------- Teil 1 ----------
{ const s = slide(); title(s, "Teil 1: Einleitung", 1); ab(s, 1, "Vorgehen");
  steps(s, ["Lies die Quelle und den Infokasten.", "Notiere die TATTEA-Angaben.", "Formuliere die Themenfrage.", "Schreibe die Einleitung (2–3 Sätze).", "Kontrolliere an der Tafelrückseite."], 1.35, 0.75);
  num(s); s.addNotes("Allein. Hilfe: Übersicht, Schritt 2a. Die Themenfrage steckt im Titel der Klausur: „Entmachtung des Adels?“ – als Frage formulieren."); }
{ const s = slide(); title(s, "TATTEA", 1); ab(s, 1, "leer");
  raster(s, ["Titel", "Autor", "Textsorte", "Thema", "Entstehung", "Adressat"].map(l => [l, ""]));
  num(s); s.addNotes("Vorlage zum Mitschreiben."); }
{ const s = slide(); title(s, "TATTEA", 1); ab(s, 1, "Lösung");
  raster(s, [["Titel", "„Über den Staat der Athener“"], ["Autor", "unbekannt, ein Adliger"], ["Textsorte", "politische Streitschrift"], ["Thema", "Kritik an der Demokratie"], ["Entstehung", "um 430–424 v. Chr., Athen"], ["Adressat", "gebildete, adlige Leser"]], { solved: true });
  num(s); s.addNotes("Ausführlich auf Aushang 1. Stolperstein: Xenophon ist NICHT der Autor – die Schrift wurde ihm nur irrtümlich zugeschrieben."); }
{ const s = slide(); title(s, "Themenfrage", 1); ab(s, 1, "Lösung");
  panel(s, D.themenfrage, 0.45, 1.35, 9.1, 1.4, C.dark, C.white, { bold: true });
  txt(s, "„Die Untersuchung geht der Frage nach, ob …“", { x: 0.45, y: 3.0, w: 9.1, h: 0.7, italic: true, valign: "middle" });
  num(s); s.addNotes("Ganzer Einleitungstext als Formulierungsvorschlag auf Aushang 1 (Tafelrückseite)."); }

// ---------- Teil 2 ----------
// ausformulierte Position je Sinnabschnitt (Wunsch der Lehrkraft: nicht zu stark abgekürzt)
const LANG = ["Er lehnt die Demokratie ab: Dort haben es die einfachen Leute besser als die Adligen.",
  "Er gibt zu: Das Volk rudert die Schiffe und sichert Athens Macht – daher sein Anspruch.",
  "Gleiche Ämter und Redefreiheit für alle wirken nur auf den ersten Blick gerecht.",
  "Das Volk überlässt die Heerführung anderen und will nur Ämter, die Geld bringen.",
  "Teure Aufgaben wie Festchor oder Kriegsschiff zahlen die Reichen – das Volk verdient.",
  "Vor Gericht urteilt das Volk nach seinem eigenen Vorteil, nicht nach dem Recht.",
  "Adlige seien gerecht und gut, das Volk ungebildet und gemein – aus Armut.",
  "Das Volk will frei sein und herrschen; eine gute Ordnung ist ihm egal.",
  "Sein Ideal: Nur die Edlen regieren, das Volk wird in Zucht gehalten."];
const POS = D.position.map(([, , z], i) => [z, LANG[i]]);
{ const s = slide(); title(s, "Teil 2: Position", 3); ab(s, 2, "Vorgehen");
  steps(s, ["Teile den Text in Sinnabschnitte.", "Notiere je Abschnitt die Position in Stichpunkten.", "Gib die Zeilen an.", "Fasse die Gesamtposition in einem Satz zusammen."], 1.35, 0.85);
  num(s); s.addNotes("Allein oder zu zweit, danach im Unterrichtsgespräch zusammentragen: Ist alles inhaltlich abgedeckt? Raster auf den nächsten Folien; Lösung danach. Stichpunkt-Zusammenfassung für die Lehrkraft: Aushang 2."); }
[[0, 3], [3, 6], [6, 9]].forEach(([a, b], n) => { const s = slide(); title(s, `Position (${n + 1}/3)`, 3); ab(s, 2, "leer");
  raster(s, POS.slice(a, b).map(([z]) => [z, ""]), { lw: 1.45, rh: 1.27, gap: 0.06 });
  num(s); s.addNotes("Raster mit den Zeilenabschnitten – zum Zusammentragen: Was sagt der Verfasser hier?"); });
[[0, 3], [3, 6], [6, 9]].forEach(([a, b], n) => { const s = slide(); title(s, `Position (${n + 1}/3)`, 3); ab(s, 2, "Lösung");
  raster(s, POS.slice(a, b), { lw: 1.45, rh: 1.27, gap: 0.06, solved: true, plain: true });
  num(s); s.addNotes("Abgleich mit den gesammelten Punkten. Ausführlich mit Zitaten: Aushang 2."); });
{ const s = slide(); title(s, "Gesamtposition", 3); ab(s, 2, "Lösung");
  panel(s, "Der Verfasser ist ein Gegner der Demokratie und spricht für den Adel. Er erkennt den Flottendienst des Volkes an, wirft ihm aber Eigennutz und Unbildung vor.", 0.45, 1.3, 9.1, 2.0, C.sand, C.dark);
  panel(s, "Sprache: wertend – „Edle“ gegen „gemeine Leute“.", 0.45, 3.5, 9.1, 0.9, C.dark, C.gold, { bold: true });
  num(s); s.addNotes(D.positionGesamt + " " + D.sprache); }

// ---------- Teil 3 ----------
const ORG = [["Volksversammlung", "alle Bürger ab 20; Gesetze, Krieg"], ["Rat der 500", "50 je Phyle; bereitet alles vor"], ["Archonten", "9 höchste Beamte, nur Reiche"], ["Areopag", "Ex-Archonten; Aufsicht"], ["Strategen", "10 Heerführer, gewählt"], ["Volksgericht", "Geschworene; jeder kann klagen"]];
const GLI = [["1. Klasse", "alle Rechte, auch Archont"], ["2. und 3. Klasse", "Rat und Ämter, nicht Archont"], ["Theten", "nur abstimmen und klagen"], ["Frauen", "keine politischen Rechte"], ["Metöken", "Steuern, Kriegsdienst, keine Rechte"], ["Sklaven", "keine Rechte, Eigentum"]];
{ const s = slide(); title(s, "Teil 3: Beschreiben", 3); ab(s, 3, "Vorgehen");
  steps(s, ["Nenne die Organe: Wer gehört dazu? Was tun sie?", "Ordne die Bevölkerung: Bürger nach Klassen, Nichtbürger.", "Nenne zu jeder Gruppe ihre Rechte.", "Schreibe einen zusammenhängenden Text.", "Kontrolliere am Aushang."], 1.35, 0.75);
  num(s); s.addNotes("Beschreiben = sachlich darstellen, noch kein Urteil. Grundlage: Kleisthenes-Stunde, Treppe Dorf – Bezirk – Phyle – Rat, Solons Klassen."); }
[["Verfassungsorgane", ORG, 3.3], ["Bevölkerung und Rechte", GLI, 2.8]].forEach(([t, rows, lw]) => {
  [false, true].forEach(solved => { const s = slide(); title(s, t, 3); ab(s, 3, solved ? "Lösung" : "leer");
    raster(s, rows.map(([l, r]) => [l, solved ? r : ""]), { lw, solved });
    num(s); s.addNotes(solved ? "Ausführlich mit Formulierungsvorschlag: Aushang 3." : "Vorlage: Die Namen sind vorgegeben, die Schüler ergänzen Zusammensetzung, Aufgaben bzw. Rechte."); }); });

// ---------- Teil 4 ----------
{ const s = slide(); title(s, "Teil 4: Urteilen", 4); ab(s, 4, "Vorgehen");
  steps(s, ["Beziehe Stellung zur Themenfrage.", "Nenne Argumente dafür – mit Belegen.", "Nenne Argumente dagegen – mit Belegen.", "Ziehe ein Fazit.", "Werturteil: Maßstab nennen."], 1.35, 0.75);
  num(s); s.addNotes("Sachurteil aus damaliger Sicht mit Belegen aus Quelle und Kontext, dann ein Werturteil mit offengelegtem Maßstab (Übersicht, Schritt 4)."); }
{ const s = slide(); title(s, "Urteil", 4); ab(s, 4, "leer");
  ["„Die Reformen des Kleisthenes beendeten … , weil …“", "„Einerseits … (Beleg) …“", "„Andererseits … (Beleg) …“", "„Insgesamt …“", "„Gemessen an … halte ich …“"].forEach((t, i) =>
    txt(s, t, { x: 0.6, y: 1.35 + i * 0.75, w: 9, h: 0.6, italic: true, valign: "middle" }));
  num(s); s.addNotes("Schreibgerüst."); }
[["Dafür", ["Phylen: Adelsblöcke zerschlagen", "Rat und Versammlung: alle Gebiete", "Bürgerrecht im Dorf, Ostrakismos", "Quelle 430: Adliger klagt über Volk"]],
 ["Dagegen", ["Klassen bleiben: Archonten nur Reiche", "Theten ohne Ämter, viele ausgeschlossen", "Adlige führen weiter (Z. 11–13)", "erst 462: Areopag entmachtet, Diäten"]]].forEach(([t, items]) => {
  const s = slide(); title(s, "Urteil: " + t, 4); ab(s, 4, "Lösung");
  items.forEach((it, i) => panel(s, it, 0.45, 1.3 + i * 0.9, 9.1, 0.78, i % 2 ? C.white : C.sand, C.dark, { line: C.sand }));
  num(s); s.addNotes(t === "Dafür" ? D.urteilPro.join(" ") : D.urteilContra.join(" ")); });
{ const s = slide(); title(s, "Fazit", 4); ab(s, 4, "Lösung");
  panel(s, "Nur teilweise: Kleisthenes brach die Macht der Adelsfamilien über ihre Gegenden.", 0.45, 1.3, 9.1, 1.45, C.dark, C.white, { bold: true });
  panel(s, "Ganz endete die Vorherrschaft erst mit Ephialtes und Perikles.", 0.45, 2.95, 9.1, 1.25, C.sand, C.dark);
  num(s); s.addNotes("Formulierungsvorschlag Sachurteil und Werturteil: Aushang 4. Andere Urteile sind richtig, wenn sie klar Stellung beziehen und belegt sind."); }

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
