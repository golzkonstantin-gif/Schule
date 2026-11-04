// Wiederholung: Die Reform des Kleisthenes Schritt für Schritt (Wunsch der Lehrkraft nach der Stunde: Klasse hatte Probleme, die Reform nachzuvollziehen).
// Alle Texte für die Klasse mindestens 27 pt; Sprechtext in den Notizen. Figur Thrasymachos wie im Mystery (erfunden).
const pptxgen = require("pptxgenjs");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", pale: "FBF8F2", green: "3C783C", blue: "3E5C76", light: "E3DCCF" };
const COL = { stadt: "C9962A", kueste: C.blue, binnen: C.green };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Die Reform des Kleisthenes – Schritt für Schritt";
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
function title(s, t, k = 1, o = {}) { aviva(s, k, o.dark); txt(s, t, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle", color: o.dark ? C.white : C.dark }); }
function box(s, t, x, y, w, h, fill, color = C.white, o = {}) {
  s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x, y, w, h, fill: { color: fill }, line: { color: o.line || fill, width: o.lw || 1 },
    color, fontFace: B, fontSize: MIN, bold: o.bold, align: o.align || "center", valign: "middle", margin: o.margin ?? 0 });
}
// Kasten mit Innenabstand: Form + eingerückter Text (Ränder in pptxgenjs-Formen unzuverlässig)
function panel(s, t, x, y, w, h, fill, color, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: fill }, line: { color: o.line || fill, width: o.lw || 1 } });
  txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color, bold: o.bold, valign: "middle" });
}
const bottom = (s, t, y = 4.4, h = 0.9) => panel(s, t, 0.45, y, 9.1, h, C.dark, C.gold, { bold: true });
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };
const REG = [["Stadt", COL.stadt], ["Küste", COL.kueste], ["Binnenland", COL.binnen]];

// 1 Titel
{ const s = slide(C.dark);
  txt(s, "GESCHICHTE · KLASSE 11 · WIEDERHOLUNG", { x: 0.7, y: 1.0, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, "Die Reform des Kleisthenes", { x: 0.7, y: 1.55, w: 8.6, h: 0.8, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Schritt für Schritt (508/07 v. Chr.)", { x: 0.7, y: 2.5, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Wiederholung, weil die Klasse die Reform noch nicht sicher nachvollziehen konnte. Ablauf: Ausgangslage – fünf Schritte – Überblick – Prüffragen. Direkt (DI): jede Folie kurz erklären, nach jedem Schritt eine Prüffrage an 2–3 Schüler unterschiedlichen Niveaus. Erst weitergehen, wenn der Schritt sitzt."); }

// 2 Leitfrage
{ const s = slide(); title(s, "Die Frage"); tag(s, "Leitfrage");
  panel(s, "Wie brach Kleisthenes die Macht der Adelsfamilien – ohne ihnen etwas wegzunehmen?", 0.45, 1.35, 9.1, 1.7, C.dark, C.white, { bold: true });
  txt(s, "Heute Schritt für Schritt: vom Dorf bis zum Rat der 500.", { x: 0.45, y: 3.35, w: 9.1, h: 0.9, valign: "middle" });
  num(s); s.addNotes("Ziel nennen: „Am Ende kann jeder die Reform in fünf Schritten erklären.“ Bezug zum Mystery: Thrasymachos verlor kein Land, kein Geld, kein Recht – und trotzdem seine Macht. Heute schauen wir genau hin, wie das funktioniert."); }

// 3 Ausgangslage
{ const s = slide(); title(s, "Vorher: Gegenden als Blöcke"); tag(s, "Ausgangslage");
  REG.forEach(([n, col], i) => {
    const x = 0.45 + i * 3.05;
    box(s, n, x, 1.3, 2.9, 0.6, col, C.white, { bold: true });
    box(s, "Adliger + seine Leute", x, 1.95, 2.9, 1.2, C.sand, C.dark, { line: col, lw: 2 });
  });
  txt(s, "Jede Gegend hält zu ihrem Adligen.", { x: 0.45, y: 3.35, w: 9.1, h: 0.8, valign: "middle" });
  bottom(s, "So kam Peisistratos an die Macht.");
  num(s); s.addNotes("Vorher (direkt): Attika besteht aus Gegenden – Stadt, Küste, Binnenland. In jeder Gegend sind viele Bauern von einem Adligen abhängig (Pacht, Schulden) und halten zu ihm. Die Adligen kämpfen mit ihren Gegenden als Block um Athen – so kam Peisistratos mit seinen Leuten an die Macht. Die vier alten Phylen sind nach Familie gebildet; im Rat der 400 geben die Adelsfamilien den Ton an. Prüffrage: „Worauf beruht die Macht eines Adligen?“ (auf seiner Gegend, den abhängigen Leuten)."); }

// 4 Schritt 1: Dörfer
{ const s = slide(); title(s, "Schritt 1: das Dorf"); tag(s, "Demos");
  REG.forEach(([n, col], i) => {
    const x = 0.45 + i * 3.05;
    box(s, n, x, 1.3, 2.9, 0.6, col, C.white, { bold: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.95, w: 2.9, h: 1.3, rectRadius: 0.06, fill: { color: C.white }, line: { color: col, width: 2 } });
    for (let d = 0; d < 12; d++) s.addShape(pres.shapes.OVAL, { x: x + 0.2 + (d % 6) * 0.45, y: 2.15 + Math.floor(d / 6) * 0.5, w: 0.3, h: 0.3, fill: { color: col }, line: { color: col } });
  });
  txt(s, "Jeder Punkt ist ein Dorf (Demos) – rund 139.", { x: 0.45, y: 3.35, w: 9.1, h: 0.8, valign: "middle" });
  bottom(s, "Bürger ist, wer in der Dorfliste steht.");
  num(s); s.addNotes("Schritt 1 (direkt): Grundeinheit ist das Dorf, der Demos (Mehrzahl Demen), in ganz Attika rund 139. Wer in der Liste seines Dorfes steht, ist Bürger – darüber entscheidet die Dorfversammlung, nicht mehr die Familie. Man nennt sich nach dem Dorf. Warnung vor dem Fehler: Niemand zieht um – jeder bleibt in seinem Dorf. Prüffrage: „Wer entscheidet jetzt, wer Bürger ist?“ (die Dorfversammlung)."); }

// 5 Schritt 2: Bezirke
{ const s = slide(); title(s, "Schritt 2: der Bezirk"); tag(s, "Trittys");
  REG.forEach(([n, col], r) => {
    const y = 1.3 + r * 0.72;
    box(s, n, 0.45, y, 2.5, 0.6, col, C.white, { bold: true });
    for (let i = 0; i < 10; i++) box(s, String(i + 1), 3.1 + i * 0.645, y, 0.58, 0.6, C.white, col, { line: col, lw: 2, bold: true });
  });
  txt(s, "Mehrere Dörfer bilden einen Bezirk (Trittys).", { x: 0.45, y: 3.55, w: 9.1, h: 0.7, valign: "middle" });
  bottom(s, "Jede Gegend: 10 Bezirke. 3 × 10 = 30 Bezirke.");
  num(s); s.addNotes("Schritt 2 (direkt): Die Dörfer werden zu Bezirken zusammengefasst, den Trittyen (Einzahl Trittys, „Drittel“). Jede der drei Gegenden besteht aus 10 Bezirken – zusammen 30. Paiania gehört zu einem der 10 Bezirke im Binnenland. Prüffrage: „Was ist größer – Dorf oder Bezirk?“ (Bezirk) – „Wie viele Bezirke gibt es?“ (30)."); }

// 6 Schritt 3: das Los
{ const s = slide(); title(s, "Schritt 3: das Los"); tag(s, "Mischen");
  REG.forEach(([n, col], r) => {
    const y = 1.3 + r * 0.72;
    box(s, n, 0.45, y, 2.5, 0.6, col, C.white, { bold: true });
    for (let i = 0; i < 6; i++) { const pick = i === [3, 1, 4][r];
      box(s, "", 3.1 + i * 0.5, y + 0.08, 0.42, 0.44, pick ? col : C.white, C.white, { line: col, lw: pick ? 1 : 1.5 }); }
  });
  s.addShape(pres.shapes.LINE, { x: 6.2, y: 2.02, w: 0.7, h: 0, line: { color: C.dark, width: 3, endArrowType: "triangle" } });
  txt(s, "Los", { x: 6.15, y: 1.45, w: 0.8, h: 0.5, bold: true, align: "center", valign: "middle" });
  REG.forEach(([, col], r) => box(s, "", 7.15, 1.3 + r * 0.72, 2.4, 0.6, col));
  txt(s, "= 1 Phyle", { x: 7.15, y: 3.45, w: 2.4, h: 0.6, bold: true, align: "center", valign: "middle" });
  bottom(s, "Je 1 Bezirk aus Stadt, Küste, Binnenland = 1 Phyle.");
  num(s); s.addNotes("Schritt 3 (direkt, laut denken): „Jetzt kommt der entscheidende Trick. Das Los zieht aus jeder Gegend einen Bezirk – einen aus der Stadt, einen von der Küste, einen aus dem Binnenland. Diese drei Bezirke bilden zusammen eine Phyle. Die drei Bezirke liegen weit auseinander – trotzdem gehören ihre Bürger jetzt zusammen.“ Warnung vor dem Fehler: Die Menschen ziehen nicht um – nur die Zugehörigkeit ihres Bezirks wird festgelegt. Prüffrage: „Woraus besteht eine Phyle?“ (aus je einem Bezirk aus Stadt, Küste, Binnenland)."); }

// 7 Schritt 4: zehn Phylen
{ const s = slide(); title(s, "Schritt 4: zehn Phylen"); tag(s, "Phylen");
  for (let p = 0; p < 10; p++) {
    const x = 0.45 + p * 0.915;
    REG.forEach(([, col], r) => box(s, "", x, 1.3 + r * 0.55, 0.8, 0.5, col));
    txt(s, String(p + 1), { x, y: 2.95, w: 0.8, h: 0.5, bold: true, align: "center", valign: "middle" });
  }
  txt(s, "30 Bezirke → 10 Phylen, jede gemischt.", { x: 0.45, y: 3.55, w: 9.1, h: 0.7, valign: "middle" });
  bottom(s, "Die 4 alten Familien-Phylen zählen nicht mehr.");
  num(s); s.addNotes("Schritt 4 (direkt): So entstehen zehn neue Phylen – in jeder sind Stadt, Küste und Binnenland je ein Drittel (gelb, blau, grün). Die vier alten Phylen, die nach Familie und Abstammung gebildet waren, verlieren ihre politische Bedeutung. Prüffrage: „Gibt es noch eine Phyle, die nur aus dem Binnenland besteht?“ (nein – jede ist gemischt)."); }

// 8 Schritt 5: wofür die Phyle zählt
{ const s = slide(); title(s, "Schritt 5: Rat und Heer"); tag(s, "Wofür?");
  box(s, "Rat der 500", 0.45, 1.3, 4.45, 0.65, C.terra, C.white, { bold: true });
  panel(s, "50 Männer je Phyle bereiten die Abstimmungen vor.", 0.45, 2.0, 4.45, 1.4, C.white, C.dark, { line: C.terra, lw: 2 });
  box(s, "Heer", 5.1, 1.3, 4.45, 0.65, C.blue, C.white, { bold: true });
  panel(s, "Gekämpft wird in Abteilungen nach Phylen.", 5.1, 2.0, 4.45, 1.4, C.white, C.dark, { line: C.blue, lw: 2 });
  txt(s, "In der Volksversammlung stimmt jeder einzeln.", { x: 0.45, y: 3.55, w: 9.1, h: 0.7, valign: "middle" });
  bottom(s, "Die Phyle bestimmt, mit wem man berät und kämpft.");
  num(s); s.addNotes("Schritt 5 (direkt): Die Phyle ist nicht nur ein Name. Jede Phyle schickt 50 Männer in den Rat der 500 – der Rat bereitet alles vor, worüber die Volksversammlung abstimmt. Im Heer kämpft man in der Abteilung seiner Phyle. Warnung vor dem Fehler: In der Volksversammlung selbst stimmt jeder einzeln, nicht die Phyle als Block. Prüffrage: „Wie viele Männer schickt jede Phyle in den Rat?“ (50) – „Was macht der Rat?“ (bereitet die Abstimmungen vor)."); }

// 9 Folge für den Adligen
{ const s = slide(); title(s, "Die Folge für Thrasymachos"); tag(s, "Ergebnis", C.gold, C.dark);
  [["Pandionis", "Paiania · B"], ["Phyle 2", "Dorf C · D"], ["Phyle 3", "Dorf E · F"]].forEach(([ph, dd], c) => {
    const x = 0.45 + c * 3.05;
    txt(s, ph, { x, y: 1.25, w: 2.9, h: 0.5, bold: true, align: "center", valign: "middle", fontFace: H });
    box(s, "Stadt", x, 1.8, 2.9, 0.55, COL.stadt);
    box(s, "Küste", x, 2.4, 2.9, 0.55, COL.kueste);
    box(s, dd, x, 3.0, 2.9, 0.55, COL.binnen, C.white, { bold: true });
  });
  txt(s, "Seine 6 Dörfer liegen in 3 Bezirken → 3 Phylen.", { x: 0.45, y: 3.65, w: 9.1, h: 0.65, valign: "middle" });
  bottom(s, "Überall nur ⅓ – kein Adliger hat mehr einen Block.");
  num(s); s.addNotes("Die Folge (laut denken): „Thrasymachos' Leute wohnen in sechs Dörfern. Diese liegen in drei verschiedenen Bezirken – und das Los hat die drei Bezirke drei verschiedenen Phylen zugeteilt. In jeder Phyle sind seine Leute nur das Binnenland-Drittel. Zwei Drittel kommen aus Stadt und Küste und schulden ihm nichts. Im Rat und im Heer kann er nichts mehr allein bestimmen.“ Er verliert kein Land und kein Recht – nur seinen Block. Prüffrage: „Was verliert Thrasymachos?“ (seinen geschlossenen Block, sonst nichts)."); }

// 10 Überblick: die Treppe
{ const s = slide(); title(s, "Der Überblick"); tag(s, "Merken", C.gold, C.dark);
  const T = [["Dorf\n(Demos)", "139"], ["Bezirk\n(Trittys)", "30"], ["Phyle", "10"], ["Rat der 500", "50 je Phyle"]];
  T.forEach(([n, z], i) => {
    const x = 0.45 + i * 2.32, y = 2.75 - i * 0.45;
    box(s, n, x, y, 2.2, 0.95, [COL.binnen, COL.kueste, C.terra, C.dark][i], C.white, { bold: true });
    txt(s, z, { x, y: y + 1.0, w: 2.2, h: 0.5, align: "center", valign: "middle", color: C.muted });
  });
  bottom(s, "Niemand zieht um – gemischt wird die Zugehörigkeit.", 4.6, 0.75);
  num(s); s.addNotes("Überblick als Treppe von unten nach oben: Dorf (Demos, rund 139) → Bezirk (Trittys, 30) → Phyle (10, je ein Bezirk aus Stadt, Küste, Binnenland) → Rat der 500 (50 je Phyle). Tafelskizze: dieselbe Treppe, darunter der Merksatz. Alle schreiben die Treppe ins Heft."); }

// 11–12 Prüffragen und Lösung
const Q = [["Wie viele Phylen gibt es?", "10"], ["Woraus besteht eine Phyle?", "Stadt + Küste + Binnenland"], ["Größer: Dorf oder Bezirk?", "der Bezirk"], ["Zieht jemand um?", "nein"], ["Was verliert der Adlige?", "nur seinen Block"]];
[false, true].forEach(solved => { const s = slide(); title(s, solved ? "Lösung" : "Prüf dich selbst", 4); tag(s, solved ? "Lösung" : "3 Min.", solved ? C.gold : C.dark, solved ? C.dark : C.white);
  Q.forEach(([q, a], i) => {
    const y = 1.3 + i * 0.78;
    s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + 0.04, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, q, { x: 1.1, y, w: solved ? 4.7 : 8.4, h: 0.6, valign: "middle" });
    if (solved) txt(s, a, { x: 5.85, y, w: 3.7, h: 0.6, valign: "middle", bold: true, color: C.terra });
  });
  num(s); s.addNotes(solved ? "Vergleichen. Wer eine Frage falsch hatte, schreibt die richtige Antwort grün daneben. Bei Unsicherheit zurück zur passenden Schritt-Folie, nicht einfach weitergehen." : "Buch und Hefter zu. Allein ins Heft, nur Stichworte (2 Min.), dann Stichproben bei 4–5 Schülern unterschiedlichen Niveaus (1 Min.)."); });

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
