const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Deutsch 7A – Mittwoch 14.10.";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";
let FOOT = "Mittwoch, 14.10. · Deutsch 7A";
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
{
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 3.2, 0.42, YEL, { r: 0.21 });
  txt(s, "MITTWOCH · 14.10.", { x: 0.9, y: 2.05, w: 3.2, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Deutsch 7A – heute in drei Teilen", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  [["1", "Leistungskontrolle", RED], ["2", "Mündliche Noten", BLUE], ["3", "Paddy Fingal", GREEN]].forEach(([n, x, c], i) => {
    box(s, 0.9 + i * 3.9, 4.1, 3.6, 0.7, c, { r: 0.35 });
    txt(s, n + "  " + x, { x: 0.9 + i * 3.9, y: 4.1, w: 3.6, h: 0.7, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}

// ============ Überblick und Lernziele ============
{
  const s = base("Überblick", "Unsere Doppelstunde: Ablauf und Ziele", "Diese Folie am Anfang zeigen und am Ende noch einmal: Was haben wir geschafft?");
  const P = [
    ["1", "Leistungskontrolle Verben", "30 Min.", "Ich zeige, was ich über Verbformen und Zeitformen kann.", RED],
    ["2", "Wie entsteht meine mündliche Note?", "ca. 15 Min.", "Ich kenne die Kriterien für meine mündliche Note und weiß, wie ich mich verbessern kann.", BLUE],
    ["3", "Paddy Fingal – Lesetechniken", "45 Min.", "Ich kann je nach Leseziel überfliegend, orientierend oder genau lesen und meinen ersten Eindruck von einer Ballade begründen.", GREEN],
  ];
  P.forEach(([n, t, z, ziel, c], i) => {
    const y = 1.7 + i * 1.6;
    box(s, 0.6, y, 12.1, 1.4, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.12, h: 1.4, fill: { color: c }, line: { type: "none" } });
    num(s, n, 0.95, y + 0.3, c, 0.6);
    txt(s, t, { x: 1.8, y: y + 0.18, w: 8.5, h: 0.5, fontFace: HEAD, fontSize: 22, bold: true });
    txt(s, z, { x: 10.6, y: y + 0.2, w: 1.9, h: 0.4, fontSize: 16, bold: true, color: c, align: "right" });
    txt(s, [{ text: "Ziel: ", options: { bold: true, color: c } }, { text: ziel, options: { italic: true } }], { x: 1.8, y: y + 0.7, w: 10.7, h: 0.6, fontSize: 16 });
  });
  txt(s, "* In Teil 3 kann LB-Aufgabe 1 entfallen, wenn die Zeit nicht reicht.", { x: 0.6, y: 6.5, w: 12.1, h: 0.35, fontSize: 14, italic: true, color: MUTED });
}

// ============ TEIL 1 Leistungskontrolle ============
FOOT = "Teil 1 · Leistungskontrolle Verben · Deutsch 7A";
section("1", "Leistungskontrolle Verben", "Verben und Zeitformen · 30 Minuten · Gruppe A und B", RED);
{
  const s = base("Teil 1 · Leistungskontrolle", "So läuft die Leistungskontrolle", "Gruppen A und B im Wechsel austeilen. Erst austeilen, dann gemeinsam starten.");
  const R1 = [
    "Schreibe deinen Namen auf das Blatt. Die Gruppe (A oder B) steht schon drauf.",
    "Vorderseite: Aufgabe 1 und 2 · Rückseite: Aufgabe 3, 4 und Zusatz.",
    "Für Aufgabe 4 brauchst du Rot, Gelb und Blau (ohne Farbstifte: einfach und doppelt unterstreichen).",
    "Du arbeitest allein und leise. Zeit: 30 Minuten.",
    "Fertig? Lies alles noch einmal durch und prüfe die Hilfsverben. Dann Blatt umdrehen und still lesen.",
  ];
  R1.forEach((x, i) => {
    num(s, i + 1, 0.6, 1.8 + i * 0.95);
    txt(s, x, { x: 1.3, y: 1.84 + i * 0.95, w: 11.3, h: 0.8, fontSize: 21 });
  });
}
{
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, "⏱  30 Minuten", { x: 0.6, y: 2.2, w: 12.1, h: 1.6, fontFace: HEAD, fontSize: 72, bold: true, color: RED, align: "center", valign: "middle" });
  txt(s, "Viel Erfolg!", { x: 0.6, y: 4.0, w: 12.1, h: 0.8, fontFace: HEAD, fontSize: 36, color: NAVY, align: "center" });
  txt(s, "Start: ______   Ende: ______", { x: 0.6, y: 5.2, w: 12.1, h: 0.6, fontSize: 24, color: MUTED, align: "center" });
  s.addNotes("Start- und Endzeit an die Tafel schreiben. Fünf Minuten vor Schluss ansagen.");
}
{
  const s = base("Teil 1 · Leistungskontrolle", "Stifte hinlegen", "Blätter einsammeln (Gruppen getrennt). Kurze Bewegungspause, dann Teil 2.");
  txt(s, "Stifte liegen. Ich sammle ein.", { x: 0.6, y: 2.6, w: 12.1, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: RED, align: "center" });
  txt(s, "Danach: Wie entsteht meine mündliche Note?", { x: 0.6, y: 4.0, w: 12.1, h: 0.6, fontSize: 24, color: MUTED, align: "center" });
}

// ============ TEIL 2 Mündliche Noten ============
FOOT = "Teil 2 · Mündliche Noten · Deutsch 7A";
section("2", "Wie entsteht meine mündliche Note?", "Die Kriterien – und wie du dich verbessern kannst", BLUE);

// ============ 2 Zwei Fragen ============
{
  const s = base("Worauf ich achte", "Zwei Fragen entscheiden über deine Note", "Beide Fragen zählen. Wer sich oft meldet, aber nur Bekanntes wiederholt, landet nicht bei 1 – und wer viel weiß, aber nie etwas sagt, auch nicht.");
  const col = (x, nr, q, sub, items, c) => {
    box(s, x, 1.75, 5.85, 4.9, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.75, w: 0.12, h: 4.9, fill: { color: c }, line: { type: "none" } });
    txt(s, nr, { x: x + 0.4, y: 1.95, w: 1, h: 0.6, fontFace: HEAD, fontSize: 36, bold: true, color: c });
    txt(s, q, { x: x + 1.2, y: 2.0, w: 4.4, h: 0.6, fontFace: HEAD, fontSize: 24, bold: true });
    txt(s, sub, { x: x + 1.2, y: 2.6, w: 4.4, h: 0.4, fontSize: 14, italic: true, color: MUTED });
    items.forEach(([lab, d], i) => {
      txt(s, [{ text: lab + "  ", options: { bold: true, color: c } }, { text: d }], { x: x + 0.45, y: 3.3 + i * 0.85, w: 5.2, h: 0.8, fontSize: 18 });
    });
  };
  col(0.6, "1", "Wie oft?", "Freiwillige Mitarbeit", [["selten", "nur, wenn ich drangenommen werde"], ["gelegentlich", "ab und zu von mir aus"], ["regelmäßig", "in jeder Stunde von mir aus"]], BLUE);
  col(6.85, "2", "Wie gut?", "Qualität der Beiträge", [["wiedergeben", "einfache Fakten aus der Stunde"], ["verknüpfen", "mit dem, was wir vorher gelernt haben"], ["verstehen und einordnen", "auch Schwieriges, Wichtiges erkennen"], ["eigene Gedanken", "begründet, klar, mit Fachbegriffen"]], GREEN);
}

// ============ 3 Treppe ============
{
  const s = base("Überblick", "Die Notentreppe", "Von unten nach oben: Mit jeder Stufe kommt etwas dazu. Die Stufe darunter muss erfüllt sein.");
  const labels = { 6: "keine freiwillige Mitarbeit · Antworten falsch", 5: "keine freiwillige Mitarbeit · teilweise richtig", 4: "gelegentlich · Fakten wiedergeben", 3: "regelmäßig · richtig und verknüpft", 2: "verstehen, einordnen, Wichtiges erkennen", 1: "eigene Gedanken · klar und begründet" };
  [6, 5, 4, 3, 2, 1].forEach((n, i) => {
    const w = 2.0, h = 1.5 + i * 0.5;
    const x = 0.6 + i * 2.05, y = 6.7 - h;
    box(s, x, y, w, h, NC[n], { r: 0.05 });
    txt(s, String(n), { x, y: y + 0.08, w, h: 0.6, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, align: "center" });
    txt(s, labels[n], { x: x + 0.1, y: y + 0.7, w: w - 0.2, h: Math.max(h - 0.75, 0.3), fontSize: 12, color: WHITE, align: "center" });
  });
}

// ============ 4–6 Stufen im Detail ============
const pair = (a, b, title, notes) => {
  const s = base("Die Stufen im Detail", title, notes);
  [a, b].forEach((n, i) => {
    const y = 1.75 + i * 2.55;
    box(s, 0.6, y, 12.1, 2.35, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.12, h: 2.35, fill: { color: NC[n] }, line: { type: "none" } });
    badge(s, n, 1.0, y + 0.35, 1.1);
    txt(s, "So arbeitest du mit:", { x: 2.4, y: y + 0.25, w: 10, h: 0.3, fontSize: 13, bold: true, color: MUTED });
    txt(s, STUFEN[n].ich, { x: 2.4, y: y + 0.55, w: 10.1, h: 1.1, fontSize: 18 });
    txt(s, [{ text: "Bedeutung: ", options: { bold: true, color: NC[n] } }, { text: STUFEN[n].fazit, options: { italic: true } }], { x: 2.4, y: y + 1.75, w: 10.1, h: 0.4, fontSize: 15 });
  });
};
pair(6, 5, "Note 6 und 5: keine freiwillige Mitarbeit", "Wichtig: Der Unterschied zwischen 5 und 6 liegt in der Richtigkeit der Antworten, wenn ich jemanden drannehme.");
pair(4, 3, "Note 4 und 3: gelegentlich oder regelmäßig", "Der Sprung von 4 zu 3: regelmäßig von sich aus melden UND Neues mit Früherem verknüpfen.");
pair(2, 1, "Note 2 und 1: verstehen und selbst denken", "Für die 1 braucht es eigenständige Gedanken: ein Problem erkennen, sachlich beurteilen und klar formulieren.");

// ============ 7 Beispiel ============
{
  const s = base("Ein Beispiel aus unserem Unterricht", "Warum steht im Bericht das Präteritum?", "Beispiel aus der Verb-Einheit. Die Antworten zeigen, wie dieselbe Frage auf verschiedenen Stufen beantwortet werden kann.");
  const ans = [
    [4, "„Weil das Vergangenheit ist.“", "richtig, aber nur Wiedergabe"],
    [3, "„Weil der Bericht etwas Vergangenes erzählt. Das haben wir beim Zeitstrahl gelernt: Was man aufschreibt, steht im Präteritum.“", "richtig und mit früherem Stoff verknüpft"],
    [2, "„Ein Bericht ist geschrieben, deshalb Präteritum. Wenn ich es mündlich erzähle, nehme ich das Perfekt – der Unterschied liegt also nicht in der Zeit, sondern darin, ob ich schreibe oder spreche.“", "verstanden, eingeordnet, das Wesentliche erkannt"],
    [1, "„… und das Plusquamperfekt brauche ich, wenn ich im Bericht die Vorgeschichte erzähle. Ich finde, dadurch wird ein Bericht spannender, weil man versteht, warum etwas passiert ist.“", "eigener Gedanke, begründet und klar formuliert"],
  ];
  ans.forEach(([n, a, k], i) => {
    const y = 1.7 + i * 1.3;
    badge(s, n, 0.6, y + 0.15, 0.8);
    box(s, 1.6, y, 11.1, 1.15, LIGHT);
    txt(s, a, { x: 1.85, y: y + 0.12, w: 10.7, h: 0.7, fontSize: 15 });
    txt(s, k, { x: 1.85, y: y + 0.8, w: 10.7, h: 0.3, fontSize: 12, italic: true, color: NC[n], bold: true });
  });
}

// ============ Unterrichtsstörungen: was zählt als Störung ============
{
  const s = base("Unterrichtsstörungen", "Was stört den Unterricht?", "Kurz mit der Klasse sammeln, dann die Folie zeigen. Wichtig: Es geht darum, dass alle lernen können, nicht um Strafe.");
  const items = [
    ["Reinrufen", "ohne Melden dazwischenreden"],
    ["Nebengespräche", "während jemand anderes spricht oder alle arbeiten"],
    ["Handy", "benutzen oder sichtbar auf dem Tisch"],
    ["Nicht arbeiten", "Aufgabe liegen lassen, andere ablenken"],
    ["Herumlaufen", "ohne Auftrag den Platz verlassen"],
    ["Material fehlt", "Heft, Buch oder Stifte nicht dabei"],
  ];
  items.forEach(([h, d], i) => {
    const x = 0.6 + (i % 2) * 6.15, y = 1.75 + Math.floor(i / 2) * 1.55;
    box(s, x, y, 5.95, 1.35, LIGHT);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.12, h: 1.35, fill: { color: RED }, line: { type: "none" } });
    txt(s, h, { x: x + 0.4, y: y + 0.22, w: 5.4, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true });
    txt(s, d, { x: x + 0.4, y: y + 0.72, w: 5.4, h: 0.45, fontSize: 16, color: MUTED });
  });
  txt(s, "Wer stört, hält nicht nur sich selbst vom Lernen ab, sondern auch alle anderen.", { x: 0.6, y: 6.5, w: 12.1, h: 0.4, fontSize: 16, italic: true, color: NAVY, align: "center" });
}

// ============ Unterrichtsstörungen: Stufen ============
{
  const s = base("Unterrichtsstörungen", "So reagiere ich – Stufe für Stufe", "Stufen an die Schulordnung bzw. Absprachen der Klassenleitung anpassen. Die Stufen gelten pro Stunde; wer sich wieder fängt, startet in der nächsten Stunde neu. Wiederholte Störungen fließen in die Einschätzung der Mitarbeit ein, weil sachliche, regelmäßige Mitarbeit dann nicht möglich ist.");
  const ST = [
    ["1", "Hinweis", "Blick, Zeichen oder Name – ohne Unterbrechung", "9DB84A"],
    ["2", "Ermahnung", "deutliche Ansage vor der Klasse", "E8A33D"],
    ["3", "Notiz", "Eintrag in meiner Liste, kurzes Gespräch nach der Stunde", "E07B45"],
    ["4", "Konsequenz", "Umsetzen, Zusatzaufgabe oder Nacharbeit", "D9534F"],
    ["5", "Eltern und Klassenleitung", "Information, Eintrag, gemeinsames Gespräch", "A93A36"],
  ];
  ST.forEach(([n, h, d, c], i) => {
    const x = 0.6 + i * 2.45, h0 = 2.3 + i * 0.42, y = 6.6 - h0;
    box(s, x, y, 2.3, h0, c, { r: 0.05 });
    txt(s, "Stufe " + n, { x, y: y + 0.15, w: 2.3, h: 0.4, fontSize: 14, bold: true, color: WHITE, align: "center" });
    txt(s, h, { x: x + 0.1, y: y + 0.55, w: 2.1, h: 0.75, fontFace: HEAD, fontSize: 18, bold: true, color: WHITE, align: "center" });
    txt(s, d, { x: x + 0.12, y: y + 1.3, w: 2.06, h: h0 - 1.35, fontSize: 13, color: WHITE, align: "center" });
  });
  txt(s, "Mit jeder Störung in derselben Stunde geht es eine Stufe höher. Die nächste Stunde beginnt wieder bei null.", { x: 0.6, y: 1.65, w: 7.0, h: 0.8, fontSize: 16, italic: true, color: MUTED });
}

// ============ 8 So verbesserst du dich ============
{
  const s = base("Nach oben kommen", "So verbesserst du deine mündliche Note", "Konkrete Tipps; gerne gemeinsam an die Tafel schreiben lassen, was sich jeder bis Weihnachten vornimmt.");
  const tips = [
    ["Melde dich auch, wenn du unsicher bist.", "Ein Fehler zählt nicht gegen dich – keine Mitarbeit schon."],
    ["Benutze Fachbegriffe.", "Finite Form, Partizip II, Strophe, Vers … statt „das Wort da“."],
    ["Begründe mit „weil“.", "Eine Antwort mit Begründung ist mehr wert als ein einzelnes Wort."],
    ["Verknüpfe mit früheren Stunden.", "„Das ist wie beim …“ – genau das macht den Unterschied zwischen 4 und 3."],
    ["Hör zu und knüpf an.", "„Ich ergänze zu …“ oder „Ich sehe das anders, weil …“"],
  ];
  tips.forEach(([h, d], i) => {
    const y = 1.75 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.08, w: 0.6, h: 0.6, fill: { color: NAVY }, line: { type: "none" } });
    txt(s, String(i + 1), { x: 0.6, y: y + 0.08, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, h, { x: 1.45, y: y, w: 11, h: 0.45, fontFace: HEAD, fontSize: 21, bold: true });
    txt(s, d, { x: 1.45, y: y + 0.45, w: 11, h: 0.4, fontSize: 16, color: MUTED });
  });
}

// ============ 9 Abschluss ============
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "Und jetzt?", { x: 0.9, y: 1.9, w: 11, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  const pts = ["Ich teile dir deine aktuelle mündliche Note einzeln mit – mit einer kurzen Begründung.", "Überleg dir: Welche Stufe willst du bis Weihnachten erreichen?", "Schreib einen Tipp ins Heft, an dem du arbeitest – und was du gegen Störungen tust."];
  pts.forEach((x, i) => {
    txt(s, [{ text: `${i + 1}  `, options: { bold: true, color: YEL } }, { text: x }], { x: 0.9, y: 3.1 + i * 0.85, w: 11.5, h: 0.7, fontSize: 22, color: WHITE });
  });
  txt(s, "Grundlage: Kriterien zur Beurteilung der mündlichen Leistung (lehrerfreund.de)", { x: 0.9, y: 6.8, w: 11, h: 0.3, fontSize: 11, color: ICE });
  s.addNotes("Noten am besten einzeln mitteilen (z. B. während einer Arbeitsphase), nicht im Plenum.");
}


// ============ TEIL 3 Paddy Fingal ============
FOOT = "Teil 3 · Paddy Fingal – Lesetechniken · Deutsch 7A";
section("3", "Paddy Fingal – Lesetechniken", "Leute, höret die Geschicht’ – Wilhelm Brandes, Lehrbuch S. 136 f.", GREEN);

// ============ 2 Ziel und Ablauf ============
{
  const s = base("Heute", "Ziel und Ablauf", "Ziel einmal vorlesen. Die Folie bleibt am Ende für die Auswertung wichtig (Folie „Ziel erreicht?“).");
  box(s, 0.6, 1.7, 12.1, 1.2, LIGHT2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 0.12, h: 1.2, fill: { color: RED }, line: { type: "none" } });
  txt(s, [{ text: "Ziel:  ", options: { bold: true, color: RED } }, { text: "Ich kann je nach Leseziel überfliegend, orientierend oder genau lesen und meinen ersten Eindruck von einer Ballade begründen.", options: { italic: true } }], { x: 1.0, y: 1.9, w: 11.5, h: 0.9, fontSize: 22 });
  const ab = ["Wer ist Paddy Fingal?", "Wie lest ihr?", "Merksatz 1: Lesetechniken", "Ü1 Paddy Fingal lesen (Frage 5 *)", "Was habe ich heute gelernt?  ·  * LB-Aufgabe 1"];
  ab.forEach((a, i) => {
    num(s, i + 1, 0.9, 3.25 + i * 0.68, NAVY, 0.45);
    txt(s, a, { x: 1.55, y: 3.28 + i * 0.68, w: 10, h: 0.5, fontSize: 20 });
  });
}

// ============ 3 Ankommen ============
{
  const s = base("Ankommen", "Wer ist Paddy Fingal?", "Zwei Vermutungen vorlesen lassen (Riese, Seemann, Raufbold …). Noch nicht lesen – nur Titel und erstes Bild.");
  box(s, 0.6, 1.8, 12.1, 2.2, LIGHT);
  txt(s, "Paddy Fingal", { x: 0.6, y: 2.05, w: 12.1, h: 0.9, fontFace: HEAD, fontSize: 54, bold: true, color: RED, align: "center" });
  txt(s, "Wilhelm Brandes · Lehrbuch S. 136", { x: 0.6, y: 3.05, w: 12.1, h: 0.5, fontSize: 20, italic: true, color: MUTED, align: "center" });
  txt(s, [{ text: "Schau nur auf den Titel und das erste Bild. Schreibe einen Satz ins Heft: ", options: {} }, { text: "Wer ist Paddy Fingal?", options: { bold: true } }], { x: 0.6, y: 4.5, w: 12.1, h: 0.8, fontSize: 24, align: "center" });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "2 Minuten · allein      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}

// ============ 4 Wie lest ihr? ============
{
  const s = base("Vorwissen", "Wie lest ihr?", "Lösung: A gezielt suchen (orientierend) · B schnell drüber (überfliegend) · C langsam, Wort für Wort (genau). Überleitung: Diese drei Leseweisen haben Namen – die Lesetechniken.");
  txt(s, "Schreibe ins Heft zu jeder Situation: Wie liest du – schnell drüber, gezielt suchen oder langsam Wort für Wort?", { x: 0.6, y: 1.55, w: 12.1, h: 0.5, fontSize: 18, color: MUTED });
  const sit = [["A", "Du bekommst die Ausschreibung für ein Turnier und suchst nur deine Startzeit."], ["B", "Du blätterst in einem Buch und willst wissen, worum es in einem Kapitel geht."], ["C", "Du bekommst eine neue Regel für deinen Sport, die du genau verstehen musst."]];
  sit.forEach(([k, d], i) => {
    const y = 2.3 + i * 1.3;
    box(s, 0.6, y, 12.1, 1.1, LIGHT);
    txt(s, k, { x: 0.9, y: y + 0.18, w: 0.8, h: 0.7, fontFace: HEAD, fontSize: 36, bold: true, color: RED });
    txt(s, d, { x: 1.8, y: y + 0.3, w: 10.6, h: 0.6, fontSize: 22 });
  });
  txt(s, "2 Minuten · allein, dann zu zweit vergleichen", { x: 0.6, y: 6.35, w: 12.1, h: 0.4, fontSize: 16, bold: true, color: MUTED });
}

// ============ 5 Merksatz 1 ============
{
  const s = base("Merksatz 1 – abschreiben, rot umrahmen", "Lesetechniken", "Rückbezug: A = orientierend, B = überfliegend, C = genau. Die Schüler ergänzen die Fachbegriffe in ihren Notizen von vorhin.");
  box(s, 0.6, 1.6, 12.1, 4.15, WHITE, { line: RED, lw: 3, r: 0.05 });
  txt(s, [{ text: "Je nach " }, { text: "Leseziel", options: { bold: true } }, { text: " lese ich unterschiedlich:" }], { x: 0.95, y: 1.8, w: 11.4, h: 0.45, fontSize: 22 });
  const rows = [
    ["ue", "schnell über den Text (Überschrift, Anfang, Ende, auffällige Wörter)", "– ich will wissen, ", "worum es geht", "."],
    ["or", "gezielt suchen, den Rest überspringen", "– ich will ", "eine bestimmte Information", " finden."],
    ["ge", "Satz für Satz, Unbekanntes klären, Randnotizen machen", "– ich will ", "alles verstehen", "."],
  ];
  rows.forEach(([k, a, b, c, d], i) => {
    const y = 2.45 + i * 1.05;
    box(s, 0.95, y, 2.6, 0.75, TECH[k].col, { r: 0.37 });
    txt(s, TECH[k].sym + "  " + TECH[k].name, { x: 0.95, y, w: 2.6, h: 0.75, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [{ text: a + " " }, { text: b }, { text: c, options: { bold: true, color: TECH[k].col } }, { text: d }], { x: 3.8, y: y + 0.05, w: 8.6, h: 0.75, fontSize: 19, valign: "middle" });
  });
  box(s, 0.6, 5.95, 5.95, 0.95, LIGHT2);
  txt(s, [{ text: "Wofür? ", options: { bold: true } }, { text: "Ich spare Zeit und finde, was ich brauche – die Startzeit in der Ausschreibung genauso wie die wichtige Stelle in einer Ballade." }], { x: 0.8, y: 6.02, w: 5.6, h: 0.85, fontSize: 14 });
  box(s, 6.75, 5.95, 5.95, 0.95, LIGHT2);
  txt(s, [{ text: "Test: ", options: { bold: true } }, { text: "Frag dich vor dem Lesen: Will ich einen Überblick, eine Information oder alles verstehen?" }], { x: 6.95, y: 6.02, w: 5.6, h: 0.85, fontSize: 14 });
}

// ============ 6–8 Ü1 ============
uebung("Verarbeiten · Ü1 a", "Ü1 a) Überfliegen", [
  "Schreibe ins Heft: Ü1 Paddy Fingal lesen – a). Schlage S. 136 f. auf.",
  "Überfliege die Ballade – genau 1 Minute für beide Seiten.",
  "Klappe das Buch zu. Schreibe einen Satz: „In der Ballade geht es um …“",
], "1 Minute lesen + 2 Minuten schreiben · allein · leise",
"Zeit laut stoppen: „Start … Stopp. Buch zu.“ Lösung z. B.: In der Ballade geht es um den Riesen Paddy Fingal und einen noch größeren Schotten, vor dem er sich versteckt.",
(s, y) => {
  box(s, 0.6, y + 0.2, 12.1, 1.3, LIGHT);
  txt(s, "⏱  1 Minute", { x: 0.6, y: y + 0.35, w: 12.1, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: BLUE, align: "center", valign: "middle" });
});

// ============ Zuhören, Kehrreim, erster Eindruck ============
{
  const s = base("Verarbeiten · Zuhören", "Zuhören und mitsprechen", "Methodenkarte „Chorisches Lesen“: Lehrkraft liest die Strophen, die Klasse spricht die letzte Zeile jeder Strophe im Chor. Vorher Tempo und Lautstärke festlegen. Danach schreiben alle den ersten Eindruck in Ü1; zwei lesen vor.");
  ["Hör zu, während die Ballade vorgelesen wird. Lies im Buch mit.", "Die letzte Zeile jeder Strophe sprecht ihr alle im Chor mit.", "Schreibe ins Heft: Die Ballade wirkt lustig / spannend / unheimlich, weil …"].forEach((x, i) => {
    num(s, i + 1, 0.6, 1.85 + i * 0.85);
    txt(s, x, { x: 1.3, y: 1.89 + i * 0.85, w: 11.3, h: 0.7, fontSize: 21, bold: true });
  });
  box(s, 0.6, 4.6, 12.1, 1.3, LIGHT2);
  txt(s, "… so ein Kerl war Paddy Fingal!", { x: 0.6, y: 4.6, w: 12.1, h: 1.3, fontFace: HEAD, fontSize: 40, bold: true, color: RED, align: "center", valign: "middle" });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "6 Minuten · alle gemeinsam, dann allein" }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}

uebung("Verarbeiten · Ü1 b", "Ü1 b) Orientierend lesen", [
  "Suche nur die Antworten – lies nicht alles!",
  "Schreibe b) ins Heft. Notiere die Antwort und die Strophe.",
], "6 Minuten · allein · leise · Wer fertig ist, macht Frage 5*",
"Lösung: 1 sechs Ellen (Str. 1) · 2 aus Schottland (Str. 2) · 3 Schaya (Str. 4) · 4 nur seine Nase (Str. 5) · *5 eine flache Stelle in einem Fluss oder Bach (Worterklärung 7). Nachfragen: Wie hast du die Stelle so schnell gefunden? (Schlüsselwörter, Fußnoten)",
(s, y) => {
  const q = ["Wie lang ist Paddys Bart?", "Woher kommt der Fremde?", "Wie heißt Paddys Frau?", "Was guckt aus dem Bett heraus?", "* Was ist eine Furt? (Tipp: Worterklärungen)"];
  q.forEach((x, i) => txt(s, [{ text: `${i + 1}  `, options: { bold: true, color: YEL } }, { text: x }], { x: 1.3, y: y + i * 0.55, w: 11.3, h: 0.5, fontSize: 19 }));
});

uebung("Verarbeiten · Ü1 c", "Ü1 c) Genau lesen", [
  "Schreibe c) ins Heft. Lies Strophe 7 bis 9 Satz für Satz.",
  "Notiere zu jeder Strophe einen Stichpunkt.",
  "Beantworte: Warum läuft der Schotte weg?",
], "5 Minuten · allein · leise",
"Lösung: Str. 7 Schaya warnt: Paddys Kind schläft. · Str. 8 Der Schotte schaut ins Bett und sieht die riesige Nase. · Str. 9 Er denkt: Wenn schon das Baby so groß ist, wie groß ist dann der Vater? · Warum? Schayas List: Sie gibt Paddy als Baby aus – der Schotte bekommt Angst vor dem „Vater“.",
(s, y) => {
  box(s, 0.6, y + 0.25, 12.1, 1.0, LIGHT2);
  txt(s, "„Ist das ein Baby, wie Ihr sagt, – ein Narr, wer’s mit ihm selber wagt!“", { x: 0.6, y: y + 0.25, w: 12.1, h: 1.0, fontFace: HEAD, fontSize: 24, italic: true, color: GREEN, align: "center", valign: "middle" });
});

// ============ 9 Vergleich ============
{
  const s = base("Verarbeiten · Vergleich", "Vergleicht zu zweit", "Danach erklärt ein Schüler im Plenum Schayas List (Frage c).");
  ["Vergleicht a, b und c.", "Verbessert mit Grün.", "Erklärt euch gegenseitig Schayas List."].forEach((x, i) => {
    num(s, i + 1, 0.6, 1.85 + i * 0.9);
    txt(s, x, { x: 1.3, y: 1.89 + i * 0.9, w: 11, h: 0.6, fontSize: 24, bold: true });
  });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "3 Minuten · zu zweit" }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}

// ============ 10 Auswerten ============
{
  const s = base("Auswerten", "Welche Lesetechnik passt?", "Lösung: 1 überfliegend · 2 orientierend · 3 genau.");
  txt(s, "Schreibe ins Heft die passende Lesetechnik. 1 Minute.", { x: 0.6, y: 1.55, w: 12.1, h: 0.5, fontSize: 18, color: MUTED });
  ["Du willst wissen, ob ein Zeitungsartikel über deinen Verein berichtet.", "Du suchst, wie lang Paddys Bart ist.", "Du sollst erklären, warum der Schotte wegläuft."].forEach((d, i) => {
    const y = 2.3 + i * 1.3;
    box(s, 0.6, y, 12.1, 1.1, LIGHT);
    txt(s, String(i + 1), { x: 0.9, y: y + 0.18, w: 0.8, h: 0.7, fontFace: HEAD, fontSize: 36, bold: true, color: RED });
    txt(s, d, { x: 1.8, y: y + 0.3, w: 10.6, h: 0.6, fontSize: 22 });
  });
}

// ============ 11 Ziel erreicht? ============
{
  const s = pres.addSlide();
  darkBg(s);
  txt(s, "Ziel erreicht?", { x: 0.9, y: 1.6, w: 11, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, "Ich kann je nach Leseziel überfliegend, orientierend oder genau lesen und meinen ersten Eindruck von einer Ballade begründen.", { x: 0.9, y: 2.6, w: 11.5, h: 0.9, fontSize: 22, italic: true, color: ICE });
  box(s, 0.9, 4.0, 11.5, 1.2, NAVY2);
  txt(s, "Schreibe ins Heft:  Heute habe ich gelernt, …", { x: 1.2, y: 4.0, w: 11, h: 1.2, fontFace: HEAD, fontSize: 30, bold: true, color: YEL, valign: "middle" });
  s.addNotes("Zwei Schüler lesen vor. Erwartet: die drei Lesetechniken und wann man sie benutzt. * Wenn Zeit bleibt: LB-Aufgabe 1 b – ein Schüler erzählt die Ballade mündlich nach (Titel, Autor, Reihenfolge, nichts dazuerfinden).");
}


pres.writeFile({ fileName: "Mittwoch_14-10_Deutsch_7A.pptx" });
