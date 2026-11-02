// Folien Kleisthenes – alle Texte für die Klasse mindestens 27 pt (Vorgabe der Lehrkraft).
const pptxgen = require("pptxgenjs");
const { D, C, NR } = require("./common.js");
const OUT = process.argv[2];
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Kleisthenes – " + D.titel;
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
function title(s, t, k, o = {}) { aviva(s, k, o.dark); txt(s, t, { x: 0.45, y: 0.15, w: o.w || 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle", color: o.dark ? C.white : C.dark }); }
function card(s, x, y, w, h, fill = C.sand) { s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } }); }
function boxText(s, t, x, y, w, h, fill, color, o = {}) { card(s, x, y, w, h, fill); txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color, valign: "middle", ...o }); }
function steps(s, items, y, gap) {
  items.forEach((t, i) => {
    s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + i * gap, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: 0.55, valign: "middle" });
  });
}
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };

// 1 Titel
{ const s = slide(C.dark);
  txt(s, "GESCHICHTE · KLASSE 11", { x: 0.7, y: 1.0, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, D.titel, { x: 0.7, y: 1.55, w: 8.6, h: 1.4, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Die Reformen des Kleisthenes (508/07 v. Chr.)", { x: 0.7, y: 3.1, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Titelfolie. Seitentafel: Ziel und Ablauf (Folie 2)."); }

// 2 Ziel und Ablauf
{ const s = slide(); title(s, "Ziel und Ablauf", 0); tag(s, "2 Min.");
  boxText(s, "Ich kann erklären, wie Kleisthenes mit den neuen Phylen die Macht der Adelsfamilien brach.", 0.45, 1.3, 9.1, 1.25, C.dark, C.white, { bold: true });
  ["Fehler finden: Wer regierte Athen?", "Rätsel und zwei neue Begriffe", "Mystery: Warum verliert ein Adliger seine Macht?", "Lösung und neue Ordnung*", "Abschlussaufgabe"].forEach((t, i) =>
    txt(s, (i + 1) + "  " + t, { x: 0.6, y: 2.7 + i * 0.5, w: 9, h: 0.5, valign: "middle" }));
  num(s); s.addNotes("0–2 Min. Ziel vorlesen (steht auch an der Seitentafel). * = bei Zeitnot nur kurz. Im Auswerten auf das Ziel zurückkommen."); }

// 3–6 Fehler-Übersicht in zwei Teilen + Lösung
const UEB = [
  ["vor 600", "Adelsfamilien", "Aristokratie = Herrschaft der Besten", [["Wenige Adelsfamilien haben die Macht.", null], ["Recht sprechen gewählte Richter aus dem Volk.", "Recht spricht der Adel."]]],
  ["624/621", "Drakon", "Aristokratie", [["Er schreibt das Recht erstmals auf.", null], ["Er schafft die Schuldknechtschaft ab.", "Die Schuldknechtschaft bleibt."]]],
  ["594/93", "Solon", "Timokratie = Herrschaft nach dem Wert (Besitz)", [["Er streicht die Schulden und verteilt das Land neu.", "Schulden ja – kein neues Land."], ["Die Rechte richten sich nach der Herkunft.", "Die Rechte richten sich nach dem Besitz."]]],
  ["546", "Peisistratos", "Tyrannis = Herrschaft eines Alleinherrschers", [["Die Volksversammlung wählt ihn.", "Er erobert die Macht mit Söldnern."], ["Er schafft Solons Verfassung ab.", "Die Verfassung bleibt."]]],
];
function ueb(s, rows, solved) {
  rows.forEach(([z, pr, form, st], r) => {
    const y = 1.25 + r * 2.05;
    card(s, 0.45, y, 9.1, 1.95, r % 2 ? C.white : C.pale);
    txt(s, [{ text: pr + "  ", options: { bold: true, fontFace: H } }, { text: z, options: { color: C.muted } }], { x: 0.65, y: y + 0.03, w: 8.8, h: 0.45, valign: "middle" });
    txt(s, form, { x: 0.65, y: y + 0.48, w: 8.8, h: 0.45, valign: "middle", color: C.terra, bold: true });
    st.forEach(([t, fix], k) => {
      const yy = y + 0.95 + k * 0.48;
      if (!solved) txt(s, "• " + t, { x: 0.65, y: yy, w: 8.8, h: 0.46, valign: "middle" });
      else if (!fix) txt(s, "✓ " + t, { x: 0.65, y: yy, w: 8.8, h: 0.46, valign: "middle", color: C.green });
      else txt(s, "✗ " + fix, { x: 0.65, y: yy, w: 8.8, h: 0.46, valign: "middle", color: C.terra, bold: true });
    });
  });
}
[[0, 2], [2, 4]].forEach(([a, b], i) => { const s = slide(); title(s, `Finde die Fehler (${i + 1}/2)`, 1); tag(s, i === 0 ? "5 Min." : "6 Fehler");
  ueb(s, UEB.slice(a, b), false); num(s);
  s.addNotes(i === 0 ? "2–7 Min. Buch und Hefter zu, nur das Heft. Ansage: „Auf zwei Folien stehen 6 Fehler. Schreibt jede Berichtigung ins Heft.“ Allein (2), Partnercheck (1), Stichproben bei 4–6 Schülern unterschiedlichen Niveaus (2). Zwischen Folie 3 und 4 nach Bedarf wechseln." : "Weiter allein bzw. im Partnercheck."); });
[[0, 2], [2, 4]].forEach(([a, b], i) => { const s = slide(); title(s, `Lösung (${i + 1}/2)`, 1); tag(s, "Lösung", C.gold, C.dark);
  ueb(s, UEB.slice(a, b), true); num(s);
  s.addNotes("7–8 Min. Vergleichen, Fehler im Heft grün verbessern."); });

// 7 Leitfrage Peisistratos
{ const s = slide(); title(s, "Peisistratos – Wegbereiter der Demokratie?", 1, { w: 6.6 }); tag(s, "Lösung", C.gold, C.dark);
  boxText(s, "Ja – ungewollt: Bauern und Mittelstand werden unabhängiger, der mächtigste Adel verliert.", 0.45, 1.35, 9.1, 1.6, C.sand, C.dark);
  boxText(s, "Aber: Die Adelsfamilien bestimmen weiter über ihre Anhänger.", 0.45, 3.15, 9.1, 1.3, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("8–9 Min. Offene Leitfrage der letzten Stunde: zwei Schüler antworten in einem Satz, dann aufdecken. Das „Aber“ ist die Gelenkstelle zu heute."); }

// 8 Sturz
{ const s = slide(); title(s, "Das Ende der Tyrannis", 1);
  [["528/27", "Peisistratos stirbt, seine Söhne herrschen."], ["514", "Sein Sohn Hipparchos wird ermordet."], ["510", "Sein Sohn Hippias wird verbannt."]].forEach(([j, t], i) =>
    txt(s, [{ text: j + " v. Chr.   ", options: { bold: true, color: C.terra, fontFace: H } }, { text: t }], { x: 0.6, y: 1.35 + i * 0.6, w: 9, h: 0.55, valign: "middle" }));
  boxText(s, "Jetzt kämpfen wieder die Adelsfamilien. Wie bricht man ihre Macht – ohne Land wegzunehmen?", 0.45, 3.35, 9.1, 1.6, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("9–10 Min. Kurz selbst erzählen (direkt). Quelle: Lehrbuch S. 16."); }

// 9 Rätsel
{ const s = slide(C.dark); aviva(s, 2, true); tag(s, "1 Min.", C.gold, C.dark);
  txt(s, "Das Rätsel", { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, color: C.white, valign: "middle" });
  txt(s, D.frage, { x: 0.45, y: 1.3, w: 9.1, h: 2.9, italic: true, color: "E8DFD0" });
  boxText(s, "Schreib ins Heft: „Ich vermute, Thrasymachos verliert seine Macht, weil …“", 0.45, 4.3, 9.1, 0.95, C.terra, C.white, { bold: true });
  num(s); s.addNotes("10–11 Min. Jeder legt sich fest (Aha vorbereiten). Nicht besprechen – die Vermutung wird am Ende mit der Lösung verglichen. Thrasymachos und Lysias sind erfunden, Paiania und die Phyle Pandionis gab es wirklich. Erwartbare Vermutungen: enteignet, verbannt, Rechte verloren – alles falsch, das ist der Aha-Moment."); }

// 10–11 Begriffe
[["Demos", "griech. dēmos = Volk, Dorfgemeinde", "Eine Dorfgemeinde in Attika. Wer in der Liste seines Demos steht, ist Bürger.", "Mehrzahl: Demen (rund 139)"],
 ["Trittys", "griech. „Drittel“", "Ein Bezirk aus mehreren Demen. Jede neue Phyle besteht aus drei Trittyen.", "Mehrzahl: Trittyen. Bekannt: Phyle = Stammesverband"]].forEach(([w, e, d, m], i) => {
  const s = slide(); aviva(s, 2);
  txt(s, w, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: 44, bold: true, fontFace: H, valign: "middle" });
  txt(s, e, { x: 0.45, y: 1.2, w: 9, h: 0.5, italic: true, color: C.muted });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 1.85, w: 0.08, h: 1.7, fill: { color: C.terra }, line: { color: C.terra } });
  txt(s, d, { x: 0.75, y: 1.85, w: 8.8, h: 1.7, fontSize: 30 });
  boxText(s, m, 0.45, 3.9, 9.1, 0.9, C.dark, C.gold, { bold: true });
  num(s); s.addNotes(i === 0 ? "11–13 Min. Direkt: Begriffe vor dem Mystery klären." : "Prüffrage an zwei Schüler: „Was ist größer – Demos oder Trittys?“ (Trittys)."); });

// 12 Vormachen
{ const s = slide(); title(s, "So geht's – ein Beispiel", 2); tag(s, "Lehrkraft");
  boxText(s, NR.A + "  … Wenn Thrasymachos ruft, folgen ihm seine Leute geschlossen.", 0.45, 1.3, 3.8, 1.6, C.pale, C.dark);
  boxText(s, NR.B + "  ______ ist Thrasymachos auch in Athen mächtig …", 5.75, 1.3, 3.8, 1.6, C.pale, C.dark);
  s.addText("Dadurch", { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 4.3, y: 1.85, w: 1.4, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
  boxText(s, "Achtung: Thrasymachos verliert weder Land noch Rechte. Die Lösung liegt woanders.", 0.45, 3.2, 9.1, 1.4, C.dark, C.gold, { bold: true });
  num(s); s.addNotes(`13–15 Min. Direkt: laut denken. „Abschnitt ${NR.A} sagt: Hunderte Bauern folgen Thrasymachos geschlossen. Abschnitt ${NR.B} sagt: Er ist auch in Athen mächtig. Wie hängt das zusammen? Seine Leute sind der Grund für seine Macht in Athen – also passt ‚Dadurch‘ in die Lücke: ‚Dadurch ist Thrasymachos auch in Athen mächtig.‘ Der Satz stimmt, ich streiche ‚dadurch‘ unten auf dem Blatt ab.“ Dann die Warnung vorlesen. Gruppen zu dritt bilden, Sets liegen bereit.`); }

// 13 Auftrag
{ const s = slide(); title(s, "Mystery · zu dritt", 3); tag(s, "12 Min.");
  steps(s, ["Ausschneiden – jeder liest seine Abschnitte vor.", "Kette legen: Was gehört zusammen?", "Wörter von unten in die Lücken, abstreichen.", "Jeder schreibt die Antwort ins Heft.", "Fest? Hilfe 1, 2, 3 vom Pult holen."], 1.35, 0.72);
  num(s); s.addNotes("15–27 Min. Indirekt: Die Gruppen erschließen die neuen Informationen selbst. Umhergehen, zuhören; Stolpersteine und Hilfen in der Handreichung. Nach 9 Min. Folie 14 zeigen: „Noch 3 Minuten – jetzt schreibt jeder die Antwort.“"); }

// 14 Satzanfänge
{ const s = slide(); title(s, "Eure Antwort ins Heft", 3); tag(s, "3 Min.");
  ["„Vorher beruhte seine Macht auf …“", "„Nach dem Sturz des Hippias …“", "„Kleisthenes teilte …“", "„Dadurch sind seine Leute …, deshalb …“"].forEach((t, i) =>
    boxText(s, t, 0.45, 1.35 + i * 0.95, 9.1, 0.8, i % 2 ? C.white : C.sand, C.dark, { italic: true }));
  num(s); s.addNotes("24–27 Min. Während die Gruppen schreiben, stehen lassen."); }

// 15 Lösungskette
{ const s = slide(); title(s, "Die Lösung", 3); tag(s, "Lösung", C.gold, C.dark);
  const K = [["Vorher", "Seine Gegend folgt ihm als Block", C.muted], ["Wende", "Das Volk setzt sich durch", C.terra], ["Reform", "10 gemischte Phylen", C.blue], ["Folge", "Seine Leute: überall nur ⅓", C.green], ["Ergebnis", "Seine Macht ist gebrochen", C.dark]];
  K.forEach(([l, t, col], i) => {
    const y = 1.3 + i * 0.78;
    s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 0.45, y, w: 2.3, h: 0.62, fill: { color: col }, color: C.white, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: 2.95, y, w: 6.6, h: 0.62, valign: "middle", bold: i === 4 });
  });
  num(s); s.addNotes(`27–29 Min. Zwei Gruppen lesen vom Platz vor, du trägst die Kette an der Tafel ein (zügig – die Einzelheiten klärt danach die Vorher-Nachher-Grafik). Abschnitte: Vorher ${NR.A}+${NR.B}, Wende ${NR.C}+${NR.D}, Reform ${NR.E}+${NR.F}, Folge ${NR.G}; Zusatz ${NR.H} (Ostrakismos). Verbindungswörter: dadurch – doch dann – daraufhin – danach – daraus – aber – (deshalb) – außerdem. Vergleich mit der Vermutung: Wie nah wart ihr dran?`); }

// 16 Vorher: seine Gegend als Block
const COL = { stadt: "C9962A", kueste: C.blue, binnen: C.green };
{ const s = slide(); title(s, "Vorher: ein Block", 3); tag(s, "gemeinsam");
  card(s, 0.45, 1.3, 5.6, 3.2, C.sand);
  txt(s, "Seine Gegend im Binnenland", { x: 0.65, y: 1.35, w: 5.3, h: 0.5, bold: true, valign: "middle" });
  boxText(s, "Thrasymachos", 0.65, 1.95, 5.1, 0.6, C.dark, C.white, { bold: true, align: "center", margin: 0 });
  ["Paiania", "Dorf B", "Dorf C", "Dorf D", "Dorf E", "Dorf F"].forEach((d, k) =>
    boxText(s, d, 0.65 + (k % 3) * 1.8, 2.75 + Math.floor(k / 3) * 0.8, 1.65, 0.65, COL.binnen, C.white, { align: "center", margin: 0 }));
  txt(s, "Hunderte abhängige Bauern folgen ihm geschlossen.", { x: 6.25, y: 1.4, w: 3.3, h: 1.6, bold: true, color: C.terra });
  txt(s, "Adel bestimmt Rat und Bürgerrecht.", { x: 6.25, y: 3.1, w: 3.3, h: 1.2 });
  boxText(s, "So kam auch Peisistratos an die Macht.", 0.45, 4.65, 9.1, 0.7, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("29–32 Min. Direkt erklären, an der Grafik zeigen: Die Macht eines Adligen beruht auf seiner Gegend. Die Bauern in sechs Dörfern sind von ihm abhängig (Pacht, Schulden) und folgen ihm geschlossen. Mit solchen Blöcken kämpften die Adligen um Athen – Peisistratos kam mit seinen Leuten aus dem Bergland an die Macht. Dazu bestimmen die Adelsfamilien im Rat der 400 und in den Familienverbänden, die über das Bürgerrecht entscheiden. Prüffrage an 2 Schüler: „Worauf beruht die Macht des Thrasymachos?“ (auf seiner Gegend, den abhängigen Bauern als Block)."); }

// 17 Nachher: drei neue Phylen
{ const s = slide(); title(s, "Nachher: aufgeteilt", 3); tag(s, "gemeinsam");
  [["Pandionis", "Paiania · B"], ["Phyle 2", "Dorf C · D"], ["Phyle 3", "Dorf E · F"]].forEach(([ph, dd], c) => {
    const x = 0.45 + c * 3.05;
    txt(s, ph, { x, y: 1.25, w: 2.9, h: 0.5, bold: true, align: "center", valign: "middle", fontFace: H });
    boxText(s, "Stadt", x, 1.8, 2.9, 0.6, COL.stadt, C.white, { align: "center", margin: 0 });
    boxText(s, "Küste", x, 2.45, 2.9, 0.6, COL.kueste, C.white, { align: "center", margin: 0 });
    boxText(s, dd, x, 3.1, 2.9, 0.6, COL.binnen, C.white, { align: "center", bold: true, margin: 0 });
  });
  txt(s, "+ 7 weitere Phylen · jede: ⅓ Stadt, ⅓ Küste, ⅓ Binnenland", { x: 0.45, y: 3.8, w: 9.1, h: 0.5, valign: "middle", color: C.muted });
  boxText(s, "Niemand zieht um – aber seine Leute sind in jeder Phyle nur ⅓.", 0.45, 4.4, 9.1, 0.95, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("32–35 Min. Laut denken, an der Grafik zeigen: „Alle wohnen weiter in ihrem Dorf, Thrasymachos behält sein Land. Neu ist nur, zu welcher Phyle ein Dorf gehört. Seine sechs Dörfer liegen in drei verschiedenen Bezirken – das Los teilt sie drei verschiedenen Phylen zu. In jeder dieser Phylen sind seine Leute nur das Binnenland-Drittel; zwei Drittel sind Fischer und Handwerker aus Stadt und Küste, die ihm nichts schulden. Im Rat der 500 (50 je Phyle), im Heer und bei der Bürgerliste kann er deshalb nichts mehr allein bestimmen. Und mit seiner Gegend die Macht ergreifen wie Peisistratos – das geht nicht mehr.“ Warnung vor dem Fehler: Thrasymachos wird nicht versetzt und verliert nichts; verteilt werden die Dörfer. Prüffragen an 3–4 Schüler: „Zieht jemand um?“ (nein) – „Was verliert Thrasymachos?“ (nichts an Besitz und Rechten – nur seinen Block) – „Wer bestimmt in seiner Phyle mit?“ (zwei Drittel aus Stadt und Küste)."); }

// 18 Auftrag Vorher-Nachher-Tabelle
{ const s = slide(); title(s, "Vorher – Nachher", 3); tag(s, "3 Min.");
  steps(s, ["Nimm das A5-Blatt „Was hat sich geändert?“.", `Ergänze „Nachher“ mit den Abschnitten ${NR.E}, ${NR.F}, ${NR.G}.`, "Notiere unten, was gleich bleibt.", "Vergleiche gleich mit der Lösung."], 1.35, 0.85);
  num(s); s.addNotes("35–38 Min. Allein, dann mit dem Nachbarn vergleichen. Durch die Reihen gehen und mitlesen: Steht bei „Machtbasis“ wirklich „auf drei Phylen verteilt“ – und nicht „er wird versetzt“ oder „enteignet“?"); }

// 19–20 Lösung Vorher-Nachher
[[0, 3], [3, 6]].forEach(([a, b], n) => { const s = slide(); title(s, `Lösung (${n + 1}/2)`, 3); tag(s, "Lösung", C.gold, C.dark);
  const xs = [0.45, 2.85, 6.05], ws = [2.3, 3.1, 3.5];
  ["", "Vorher", "Nachher"].forEach((t, i) => t && boxText(s, t, xs[i], 1.3, ws[i], 0.55, C.dark, C.white, { bold: true, align: "center", margin: 0 }));
  D.vergleich.slice(a, b).forEach(([, , , m, v, nn], r) => {
    const y = 1.95 + r * 0.85;
    boxText(s, m, xs[0], y, ws[0], 0.75, C.sand, C.dark, { bold: true, margin: 0 });
    txt(s, v, { x: xs[1] + 0.1, y, w: ws[1] - 0.1, h: 0.75, valign: "middle" });
    txt(s, nn, { x: xs[2] + 0.1, y, w: ws[2] - 0.1, h: 0.75, valign: "middle", bold: true, color: C.terra });
  });
  if (n === 1) boxText(s, "Gleich bleiben: Land, Vermögen, Klasse, Rechte.", 0.45, 4.6, 9.1, 0.7, C.dark, C.gold, { bold: true });
  num(s); s.addNotes(n === 0 ? "38–39 Min. Zeile für Zeile aufdecken, Stichproben vom Platz, alle verbessern im Blatt." : "Ausführlich im Blatt (A5) und in der Handreichung. Gelenkstelle: „Kleisthenes hat dem Adel nichts weggenommen, sondern seine Machtbasis aufgeteilt – das ist der Kern der Abschlussaufgabe.“"); });

// 17 Abschlussaufgabe
{ const s = slide(); title(s, "Abschlussaufgabe", 4); tag(s, "4 Min.");
  boxText(s, "Kleisthenes änderte Solons Vermögensklassen nicht – trotzdem gilt er als „Wegbereiter der Demokratie“. Erkläre in 3 Sätzen.", 0.45, 1.3, 9.1, 1.75, C.dark, C.white, { bold: true });
  ["„Solon gab den Bürgern …, aber …“", "„Kleisthenes änderte nicht …, sondern …“", "„Dadurch …“"].forEach((t, i) =>
    txt(s, t, { x: 0.6, y: 3.25 + i * 0.6, w: 9, h: 0.55, italic: true, valign: "middle" }));
  num(s); s.addNotes("39–43 Min. Prüfaufgabe zum Ziel, allein ins Heft. Danach gezielt zwei Schüler unterschiedlichen Niveaus vorlesen lassen. Diese Phase nie kürzen."); }

// 18 Beispielantwort
{ const s = slide(); title(s, "So könnte es lauten", 4); tag(s, "Lösung", C.gold, C.dark);
  txt(s, "„Solon gab allen Bürgern Rechte, aber die Adligen beherrschten ihre Gegend als Block. Kleisthenes änderte nicht die Rechte, sondern die Ordnung: Er mischte die Bürger in zehn neuen Phylen. Dadurch verloren die Adligen ihre Machtbasis, und die Bürger aller Gebiete entschieden gemeinsam.“", { x: 0.45, y: 1.3, w: 9.1, h: 3.6, italic: true });
  num(s); s.addNotes("43–44 Min. Vergleichen."); }

// 19 Ziel erreicht?
{ const s = slide(C.dark); aviva(s, 4, true);
  txt(s, "Haben wir das Ziel erreicht?", { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, color: C.white, valign: "middle" });
  txt(s, "Ich kann erklären, wie Kleisthenes mit den neuen Phylen die Macht der Adelsfamilien brach.", { x: 0.45, y: 1.5, w: 9.1, h: 1.6, color: C.gold, bold: true });
  txt(s, "Mittwoch: Wie beschreibt Aristoteles diese Reform?", { x: 0.45, y: 3.6, w: 9.1, h: 0.8, color: "E8DFD0" });
  num(s); s.addNotes("44–45 Min. Rückbezug aufs Ziel, Ausblick auf Mittwoch (Quellenanalyse)."); }

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
