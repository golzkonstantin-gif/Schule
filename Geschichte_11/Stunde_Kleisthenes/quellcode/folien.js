const pptxgen = require("pptxgenjs");
const { D, C, NR, nr } = require("./common.js");
const OUT = process.argv[2];
const PEI = "/home/user/Schule/Geschichte_11/Stunde_Peisistratos/";
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Kleisthenes – Warum stimmt Lysias anders?";
const H = "Cambria", B = "Calibri";
const FOOT = "Kleisthenes – Warum stimmt Lysias anders?";
const PH = ["A", "V", "I", "V", "A"], PHN = ["Ankommen", "Vorwissen", "Informieren", "Verarbeiten", "Auswerten"];

function txt(s, text, o) { s.addText(text, { isTextBox: true, fontFace: B, color: C.dark, margin: 0, valign: "top", ...o }); }
function aviva(s, k, dark = false) {
  PH.forEach((l, i) => {
    const on = i === k;
    s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.05, x: 7.75 + i * 0.36, y: 0.1, w: 0.32, h: 0.28,
      fill: { color: on ? C.terra : (dark ? "3A3024" : C.sand) }, color: on ? C.white : (dark ? "8C8170" : C.soft), fontFace: B, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  txt(s, PHN[k], { x: 7.75, y: 0.4, w: 1.76, h: 0.2, fontSize: 9, color: dark ? C.gold : C.terra, bold: true, align: "center" });
}
function head(s, k, kicker, title) {
  aviva(s, k);
  txt(s, kicker.toUpperCase(), { x: 0.5, y: 0.3, w: 7.0, h: 0.25, fontSize: 11, bold: true, color: C.terra, charSpacing: 2 });
  txt(s, title, { x: 0.5, y: 0.55, w: 7.2, h: 0.65, fontSize: 26, bold: true, fontFace: H });
}
function footer(s, dark = false) {
  if (!dark) txt(s, FOOT, { x: 0.5, y: 5.22, w: 6, h: 0.2, fontSize: 9, color: C.soft });
  s.slideNumber = { x: 9.0, y: 5.17, w: 0.5, h: 0.25, fontFace: B, fontSize: 9, color: C.soft, align: "right" };
}
function card(s, x, y, w, h, fill = C.sand) { s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } }); }
function circle(s, label, x, y, d = 0.42, fill = C.terra, fs = 14) {
  s.addText(label, { isTextBox: true, shape: pres.shapes.OVAL, x, y, w: d, h: d, fill: { color: fill }, color: C.white, fontFace: H, fontSize: fs, bold: true, align: "center", valign: "middle", margin: 0 });
}
function chip(s, t, x, y, w, h, fill, color, fs = 12) {
  s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x, y, w, h, fill: { color: fill }, color, fontFace: B, fontSize: fs, bold: true, align: "center", valign: "middle", margin: 0 });
}
function timeChip(s, t) { chip(s, t, 6.35, 0.12, 1.25, 0.28, C.dark, C.white, 11); }
function solChip(s, t = "Lösung") { chip(s, t, 6.35, 0.12, 1.25, 0.28, C.gold, C.dark, 11); }
function steps(s, items, x, y, w, fs = 14, gap = 0.62) {
  items.forEach((t, i) => { circle(s, String(i + 1), x, y + i * gap, 0.4); txt(s, t, { x: x + 0.55, y: y + i * gap - 0.04, w: w - 0.55, h: gap - 0.04, fontSize: fs, valign: "middle" }); });
}

// 1 Titel
{
  const s = pres.addSlide(); s.background = { color: C.dark };
  txt(s, "GESCHICHTE · KLASSE 11", { x: 0.7, y: 1.1, w: 8.6, h: 0.3, fontSize: 12, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, "Warum stimmt Lysias anders?", { x: 0.7, y: 1.45, w: 8.6, h: 0.8, fontSize: 36, bold: true, fontFace: H, color: C.white });
  txt(s, "Die Reformen des Kleisthenes (508/07 v. Chr.)", { x: 0.7, y: 2.35, w: 8.6, h: 0.5, fontSize: 18, color: "E8DFD0" });
  s.addNotes("Titelfolie. Seitentafel vorbereitet: Ziel und Ablauf (Folie 2), daneben das leere Peisistratos-Gerüst (Skizze in der Handreichung).");
}

// 2 Ziel und Ablauf
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 0, "Ankommen", "Ziel und Ablauf"); timeChip(s, "2 Min.");
  card(s, 0.5, 1.35, 9, 1.05, C.dark);
  txt(s, [{ text: "Ich kann erklären, ", options: {} }, { text: "wie Kleisthenes mit der neuen Phylenordnung die Macht der Adelsfamilien brach.", options: { bold: true, color: C.gold } }],
    { x: 0.75, y: 1.35, w: 8.5, h: 1.05, fontSize: 19, color: C.white, valign: "middle" });
  const A = [["V", "Wiederholung: Peisistratos – und was danach geschah"], ["I", "Eine rätselhafte Frage und zwei neue Begriffe"], ["V", "Mystery: Warum stimmt Lysias anders?"], ["V", "Die Lösung und die neue Ordnung Attikas (Ostrakismos*)"], ["A", "Abschlussaufgabe: Kleisthenes und Solon"]];
  A.forEach(([l, t], i) => { chip(s, l, 0.5, 2.62 + i * 0.5, 0.42, 0.38, C.sand, C.terra, 13); txt(s, t, { x: 1.1, y: 2.62 + i * 0.5, w: 8.2, h: 0.38, fontSize: 15, valign: "middle" }); });
  footer(s);
  s.addNotes("0–2 Min. Ziel vorlesen, steht auch an der Seitentafel. * = wird bei Zeitnot nur kurz gestreift. Im Auswerten auf das Ziel zurückkommen.");
}

// 3 Wiederholung Peisistratos
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 1, "Wiederholung · allein, dann zu zweit", "Was wisst ihr noch über Peisistratos?"); timeChip(s, "5 Min.");
  card(s, 0.5, 1.35, 9, 0.5, C.terra);
  txt(s, "Buch und Hefter bleiben zu.", { x: 0.75, y: 1.35, w: 8.5, h: 0.5, fontSize: 16, bold: true, color: C.white, valign: "middle" });
  steps(s, ["Schau dir das leere Gerüst an der Tafel an. Notiere im Heft zu jedem der vier Felder ein Stichwort: Wer profitiert, und wovon?",
    "Vergleiche mit deinem Nachbarn und ergänze, was dir fehlt.",
    "Wer aufgerufen wird, nennt ein Feld und eine Maßnahme. Die Lehrkraft trägt ein."], 0.5, 2.1, 5.3, 13.5, 0.85);
  s.addImage({ path: PEI + "Tafelbild_stumm.png", x: 6.1, y: 2.05, w: 3.2, h: 3.18 });
  footer(s);
  s.addNotes("2–7 Min. Direkt: Wiederholung allein (2) – Partnercheck (1) – Stichproben bei 4 Schülern (2), du trägst ein. Warum: Der Abruf aus dem Gedächtnis fehlte letzte Stunde; die Phylen-Idee baut darauf auf (Peisistratos schwächte den Adel nur teilweise).");
}

// 4 Lösung + Leitfrage nachholen
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 1, "Wiederholung · Lösung", "Peisistratos: Wegbereiter der Demokratie?"); solChip(s);
  s.addImage({ path: PEI + "Tafelbild_Loesung.png", x: 0.4, y: 1.3, w: 5.2, h: 3.45 });
  card(s, 5.85, 1.3, 3.65, 3.6, C.sand);
  txt(s, "Ja – ungewollt:", { x: 6.05, y: 1.4, w: 3.3, h: 0.35, fontSize: 15, bold: true, fontFace: H, color: C.terra });
  txt(s, "Bauern und Mittelstand werden unabhängiger, das Volk wächst zusammen, die mächtigsten Gegner aus dem Adel verlieren.", { x: 6.05, y: 1.8, w: 3.3, h: 1.4, fontSize: 13 });
  txt(s, "Aber:", { x: 6.05, y: 3.2, w: 3.3, h: 0.3, fontSize: 15, bold: true, fontFace: H, color: C.terra });
  txt(s, "Die Macht hängt an einem Mann – und die Adelsfamilien bestimmen weiter über ihre Anhänger.", { x: 6.05, y: 3.55, w: 3.3, h: 1.2, fontSize: 13 });
  footer(s);
  s.addNotes("7–9 Min. Die offene Leitfrage der letzten Stunde: zwei Schüler antworten in einem Satz, dann die Folie. Das „Aber“ ist die Gelenkstelle zur heutigen Stunde.");
}

// 5 Gelenkstelle: Sturz
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 1, "Und dann?", "Das Ende der Tyrannis");
  const T = [["528/27 v. Chr.", "Peisistratos stirbt. Seine Söhne Hippias und Hipparchos herrschen weiter."], ["514 v. Chr.", "Hipparchos wird ermordet. Es kommt zu blutigen Aufständen."], ["510 v. Chr.", "Hippias wird verbannt. Athen hat keinen Herrscher mehr."]];
  s.addShape(pres.shapes.LINE, { x: 0.7, y: 1.75, w: 8.6, h: 0, line: { color: C.light, width: 3 } });
  T.forEach(([t, d], i) => {
    const x = 0.5 + i * 3.05;
    circle(s, String(i + 1), x + 0.05, 1.53, 0.45);
    txt(s, t, { x, y: 2.1, w: 2.85, h: 0.4, fontSize: 18, bold: true, fontFace: H, color: C.terra });
    txt(s, d, { x, y: 2.55, w: 2.85, h: 1.2, fontSize: 15 });
  });
  card(s, 0.5, 4.0, 9, 0.9, C.dark);
  txt(s, "Jetzt kämpfen wieder die Adelsfamilien um die Macht. Wie kann man ihre Macht brechen – ohne ihnen das Land wegzunehmen?", { x: 0.75, y: 4.0, w: 8.5, h: 0.9, fontSize: 16, bold: true, color: C.gold, valign: "middle" });
  footer(s);
  s.addNotes("9–10 Min. Kurz selbst erzählen (direkt). Quelle: Lehrbuch S. 16 (Hipparchos ermordet 514, Hippias verbannt 510).");
}

// 6 Mystery-Frage + Vermutung
{
  const s = pres.addSlide(); s.background = { color: C.dark };
  aviva(s, 2, true);
  txt(s, "DAS RÄTSEL · ALLEIN", { x: 0.5, y: 0.3, w: 7, h: 0.25, fontSize: 11, bold: true, color: C.gold, charSpacing: 2 });
  chip(s, "1 Min.", 6.35, 0.12, 1.25, 0.28, C.gold, C.dark, 11);
  txt(s, "Warum stimmt Lysias anders?", { x: 0.5, y: 0.65, w: 7, h: 0.6, fontSize: 28, bold: true, fontFace: H, color: C.white });
  txt(s, D.frage, { x: 0.5, y: 1.45, w: 9, h: 1.9, fontSize: 19, italic: true, color: "E8DFD0" });
  card(s, 0.5, 3.6, 9, 0.95, C.terra);
  txt(s, "Schreib eine erste Vermutung ins Heft: „Ich vermute, Lysias stimmt anders, weil …“", { x: 0.75, y: 3.6, w: 8.5, h: 0.95, fontSize: 17, bold: true, color: C.white, valign: "middle" });
  footer(s, true);
  s.addNotes("10–11 Min. Jeder legt sich fest (Aha vorbereiten). Nicht besprechen – die Vermutung wird am Ende mit der Lösung verglichen. Lysias ist eine erfundene Figur, Paiania und die Phyle Pandionis gab es wirklich.");
}

// 7 Begriffe
{
  const s = pres.addSlide(); s.background = { color: C.white };
  aviva(s, 2);
  txt(s, "BEGRIFFE", { x: 0.5, y: 0.25, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.terra, charSpacing: 3 });
  const Bg = [["Demos", "griech. dēmos = Volk, auch: Dorfgemeinde", "Eine Dorf- oder Stadtteilgemeinde in Attika. Kleisthenes gliedert Attika in rund 139 Demen. Wer in der Liste seines Demos steht, ist Bürger."],
    ["Trittys", "griech. „Drittel“ (Mehrzahl: Trittyen)", "Ein Bezirk aus mehreren Demen. Jede neue Phyle besteht aus drei Trittyen – daher der Name."]];
  Bg.forEach(([t, e, d], i) => {
    const x = 0.5 + i * 4.6;
    txt(s, t, { x, y: 0.6, w: 4.3, h: 0.6, fontFace: H, fontSize: 34, bold: true });
    txt(s, e, { x, y: 1.2, w: 4.3, h: 0.3, fontSize: 14, italic: true, color: C.muted });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w: 0.07, h: 2.3, fill: { color: C.terra }, line: { color: C.terra } });
    txt(s, d, { x: x + 0.25, y: 1.6, w: 4.05, h: 2.4, fontSize: 21 });
  });
  card(s, 0.5, 4.2, 9, 0.8, C.dark);
  s.addText([{ text: "Schon bekannt: ", options: { bold: true, color: C.gold } }, { text: "Phyle = Stammesverband. Bisher gab es vier alte Phylen, in denen die Adelsfamilien den Ton angaben.", options: { color: C.white } }],
    { x: 0.8, y: 4.2, w: 8.5, h: 0.8, fontFace: B, fontSize: 16, valign: "middle", margin: 0 });
  footer(s);
  s.addNotes("11–13 Min. Direkt: Begriffe vor dem Mystery klären (Begriffs-Check). Prüffrage an zwei Schüler: „Was ist größer – Demos oder Trittys?“ (Trittys).");
}

// 8 Vormachen
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 2, "So geht's – ein Beispiel", "Zwei Schnipsel verbinden"); chip(s, "Lehrkraft", 6.35, 0.12, 1.25, 0.28, C.dark, C.white, 11);
  const ex = [[NR.B, "Seit Solon darf jeder Bürger über 20 in der Volksversammlung abstimmen – auch ein armer Bauer wie Lysias."],
    [NR.C, "Lysias' Nachbar ist ein reicher Adliger. In schlechten Jahren hat er Lysias Saatgut geliehen …"]];
  ex.forEach(([n, t], i) => {
    const x = 0.5 + i * 4.9;
    card(s, x, 1.35, 4.1, 1.35, C.pale);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.35, w: 4.1, h: 1.35, rectRadius: 0.08, fill: { type: "none" }, line: { color: C.light, width: 1, dashType: "dash" } });
    txt(s, [{ text: n + "  ", options: { bold: true, color: C.terra, fontFace: H } }, { text: t }], { x: x + 0.15, y: 1.45, w: 3.8, h: 1.2, fontSize: 13.5 });
  });
  chip(s, "aber", 4.68, 1.82, 0.64, 0.4, C.terra, C.white, 13);
  card(s, 0.5, 2.95, 9, 0.95, C.sand);
  txt(s, [{ text: "Laut gedacht: ", options: { bold: true, color: C.terra } }, { text: "„Schnipsel " + NR.B + " sagt: Er hat das Recht. Schnipsel " + NR.C + " sagt: Er ist abhängig. Zusammen: Er hat das Stimmrecht, aber er nutzt es nicht frei. Die beiden lege ich nebeneinander und schreibe ‚aber‘ dazwischen.“", options: { italic: true } }],
    { x: 0.75, y: 2.95, w: 8.5, h: 0.95, fontSize: 14, valign: "middle" });
  card(s, 0.5, 4.1, 9, 0.85, C.dark);
  txt(s, [{ text: "Achtung, typischer Fehler: ", options: { bold: true, color: C.gold } }, { text: "Kleisthenes nimmt niemandem Land weg und ändert Solons Vermögensklassen nicht. Sucht die Lösung woanders!", options: { color: C.white } }],
    { x: 0.75, y: 4.1, w: 8.5, h: 0.85, fontSize: 15, valign: "middle" });
  footer(s);
  s.addNotes("13–15 Min. Direkt: modellieren und laut denken. Warnung vor dem Fehler vorlesen. Dann Gruppen zu dritt bilden und Sets austeilen (vorher bereitlegen).");
}

// 9 Auftrag Mystery
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 3, "Mystery · Gruppen zu dritt · Heft", "Warum stimmt Lysias anders?"); timeChip(s, "14 Min.");
  steps(s, ["Schneidet die Schnipsel aus und verteilt sie. Jeder liest seine vor.",
    "Sortiert aus: Was hilft bei der Frage – was nicht?",
    "Legt die Schnipsel zu einer Erklärungskette. Schreibt Pfeile oder „weil“, „deshalb“, „aber“ dazwischen.",
    "Jeder schreibt die Antwort in 3–4 Sätzen ins Heft.",
    "Ihr kommt nicht weiter? Holt euch Hilfe 1, 2 oder 3 vom Pult – der Reihe nach."], 0.5, 1.35, 5.6, 13.5, 0.68);
  card(s, 6.3, 1.35, 3.2, 3.55, C.sand);
  txt(s, "SATZANFÄNGE", { x: 6.5, y: 1.45, w: 2.9, h: 0.25, fontSize: 10, bold: true, color: C.terra, charSpacing: 2 });
  txt(s, "„Vorher war Lysias …, weil …“\n\n„Nach dem Sturz des Hippias …“\n\n„Kleisthenes teilte …“\n\n„Dadurch …, deshalb kann Lysias jetzt …“", { x: 6.5, y: 1.8, w: 2.9, h: 3.0, fontSize: 13, italic: true });
  footer(s);
  s.addNotes("15–29 Min. Indirekt: Die Gruppen erschließen die neuen Informationen selbst (Text in Schnipseln). Umhergehen, zuhören. Stolpersteine und Hilfen in der Handreichung. Nach 10 Min. ansagen: „Noch 4 Minuten – jetzt schreibt jeder die Antwort.“ Zwei Schnipsel helfen nicht (Oliven, Großvater) – sie dürfen aussortiert werden.");
}

// 10 Lösung: Erklärungskette
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 3, "Mystery · Lösung", "Die Erklärungskette"); solChip(s);
  const cols = [C.muted, C.terra, C.blue, C.green, C.dark];
  D.kette.slice(0, 5).forEach(([t, ids, d], i) => {
    const y = 1.3 + i * 0.7;
    chip(s, t, 0.5, y, 1.55, 0.58, cols[i], C.white, 12.5);
    txt(s, ids ? "Schnipsel " + nr(ids) : "", { x: 2.2, y, w: 1.6, h: 0.58, fontSize: 10.5, color: C.soft, valign: "middle" });
    txt(s, d, { x: 3.7, y, w: 5.8, h: 0.58, fontSize: 12.5, valign: "middle", bold: i === 4 });
    if (i < 4) s.addShape(pres.shapes.LINE, { x: 1.27, y: y + 0.58, w: 0, h: 0.12, line: { color: C.soft, width: 1.5, endArrowType: "triangle" } });
  });
  txt(s, "* Schnipsel " + nr("P") + ": Gegen einen neuen Tyrannen hat das Volk jetzt den Ostrakismos (Scherbengericht). Unwichtig: Schnipsel " + nr("Q, R") + ".", { x: 0.5, y: 4.85, w: 9, h: 0.3, fontSize: 11, color: C.muted });
  footer(s);
  s.addNotes("29–33 Min. Zwei Gruppen lesen ihre Antwort vom Platz vor, du trägst die Kette an der Tafel ein. Dann diese Folie. Impuls: Vergleicht mit eurer Vermutung von vorhin – wie nah wart ihr dran?");
}

// 11 Schaubild Phylenordnung
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 3, "Die neue Ordnung Attikas", "Mischen statt enteignen"); solChip(s, "Ergebnis");
  const R = [["Stadt", "C9962A"], ["Küste", C.blue], ["Binnenland", C.green]];
  R.forEach(([n, col], r) => {
    const y = 1.45 + r * 1.05;
    txt(s, n, { x: 0.5, y, w: 1.3, h: 0.3, fontSize: 13, bold: true, color: col });
    for (let i = 0; i < 10; i++) s.addShape(pres.shapes.RECTANGLE, { x: 0.5 + i * 0.29, y: y + 0.33, w: 0.24, h: 0.42, fill: { color: col }, line: { color: C.white, width: 1 } });
  });
  txt(s, "3 Gebiete × 10 Trittyen = 30 Trittyen", { x: 0.5, y: 4.55, w: 3.2, h: 0.3, fontSize: 12, bold: true });
  s.addShape(pres.shapes.LINE, { x: 3.65, y: 2.95, w: 0.75, h: 0, line: { color: C.dark, width: 3, endArrowType: "triangle" } });
  txt(s, "per Los", { x: 3.6, y: 2.6, w: 0.9, h: 0.3, fontSize: 11, color: C.muted, align: "center" });
  for (let p = 0; p < 10; p++) {
    const x = 4.6 + (p % 5) * 0.98, y = 1.45 + Math.floor(p / 5) * 1.62;
    R.forEach(([, col], r) => s.addShape(pres.shapes.RECTANGLE, { x, y: y + r * 0.36, w: 0.8, h: 0.34, fill: { color: col }, line: { color: C.white, width: 1 } }));
    txt(s, "Phyle " + (p + 1), { x: x - 0.05, y: y + 1.1, w: 0.9, h: 0.25, fontSize: 9.5, align: "center", color: C.muted });
  }
  txt(s, "10 neue Phylen – jede mit Stadt, Küste und Binnenland", { x: 4.6, y: 4.55, w: 4.9, h: 0.3, fontSize: 12, bold: true });
  card(s, 0.5, 4.88, 9, 0.32, C.dark);
  txt(s, "Jede Phyle schickt 50 Männer → Rat der 500 · Die Gefolgschaft eines Adligen ist über mehrere Phylen verteilt.", { x: 0.65, y: 4.88, w: 8.7, h: 0.32, fontSize: 12, bold: true, color: C.gold, valign: "middle" });
  s.addNotes("33–38 Min. Schaubild erklären (direkt, 2 Min.), dann Prüffragen an 3–4 Schüler: „Warum mischt Kleisthenes gerade drei Gebiete?“ (damit keine Familie in einer Phyle die Mehrheit hat) – „Was verliert der Adlige, was behält er?“ (verliert die geschlossene Gefolgschaft, behält Land und Vermögensklasse). Tafelskizze in der Handreichung.");
  s.slideNumber = { x: 9.0, y: 5.25, w: 0.5, h: 0.25, fontFace: B, fontSize: 9, color: C.soft, align: "right" };
}

// 12 Abschlussaufgabe
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 4, "Abschlussaufgabe · allein · Heft", "Kleisthenes und Solon"); timeChip(s, "4 Min.");
  card(s, 0.5, 1.35, 9, 1.35, C.dark);
  txt(s, "Kleisthenes änderte an Solons Vermögensklassen nichts – und trotzdem gilt er als „Wegbereiter der Demokratie“. Erkläre das in 3 Sätzen.", { x: 0.8, y: 1.35, w: 8.4, h: 1.35, fontSize: 19, bold: true, fontFace: H, color: C.white, valign: "middle" });
  card(s, 0.5, 2.95, 9, 1.55, C.sand);
  txt(s, "SATZANFÄNGE", { x: 0.75, y: 3.05, w: 4, h: 0.25, fontSize: 10, bold: true, color: C.terra, charSpacing: 2 });
  txt(s, "„Solon gab den Bürgern …, aber …“     „Kleisthenes änderte nicht …, sondern …“     „Dadurch konnten die Bürger …“", { x: 0.75, y: 3.4, w: 8.5, h: 1.0, fontSize: 15, italic: true });
  txt(s, "Danach: Zwei lesen vor. Haben wir das Ziel erreicht?", { x: 0.5, y: 4.65, w: 9, h: 0.35, fontSize: 14, bold: true, color: C.terra });
  footer(s);
  s.addNotes("38–43 Min. Prüfaufgabe zum Ziel. Gezielt zwei Schüler unterschiedlichen Niveaus vorlesen lassen, dann Folie 13 und Rückbezug aufs Ziel (Seitentafel). Diese Phase nie kürzen – bei Zeitnot Folie 11 nur kurz zeigen.");
}

// 13 Lösung Abschluss
{
  const s = pres.addSlide(); s.background = { color: C.white };
  head(s, 4, "Abschlussaufgabe · Beispiel", "So könnte die Antwort lauten"); solChip(s);
  card(s, 0.5, 1.35, 9, 2.0, C.sand);
  txt(s, "„Solon gab allen Bürgern das Stimmrecht, aber viele Bauern stimmten weiter so, wie ihr Adelsherr es wollte. Kleisthenes änderte nicht die Rechte, sondern die Bindungen: Die neuen Phylen mischten Stadt, Küste und Binnenland und verteilten die Gefolgschaft der Adligen. Dadurch konnten die Bürger ihr Stimmrecht zum ersten Mal frei nutzen.“", { x: 0.8, y: 1.45, w: 8.4, h: 1.8, fontSize: 15, italic: true, valign: "middle" });
  card(s, 0.5, 3.6, 9, 1.3, C.dark);
  txt(s, [{ text: "Haben wir das Ziel erreicht?\n", options: { bold: true, color: C.gold, fontSize: 15 } }, { text: "Ich kann erklären, wie Kleisthenes mit der neuen Phylenordnung die Macht der Adelsfamilien brach.", options: { color: C.white, fontSize: 16 } }],
    { x: 0.8, y: 3.6, w: 8.4, h: 1.3, valign: "middle" });
  footer(s);
  s.addNotes("43–45 Min. Rückbezug zum Ziel. Ausblick: Am Mittwoch lesen wir, wie Aristoteles diese Reform beschreibt – und lernen, eine Quelle zu analysieren.");
}

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
