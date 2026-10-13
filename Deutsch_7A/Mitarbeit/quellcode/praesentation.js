const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Mündliche Noten";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Mündliche Noten · Deutsch 7A";
// Notenfarben von 6 (rot) bis 1 (grün)
const NC = { 1: "2E9E6B", 2: "5AAE62", 3: "9DB84A", 4: "E8A33D", 5: "E07B45", 6: "D9534F" };
let pageNo = 1;

function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 16, color: NAVY, margin: 0, valign: "top" }, o));
}
function base(kicker, title, notes) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 2 });
  txt(s, title, { x: 0.6, y: 0.72, w: 12.1, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  txt(s, FOOT, { x: 0.6, y: 7.0, w: 6, h: 0.25, fontSize: 10, color: MUTED });
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

// ============ 1 Titel ============
{
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 2.6, 0.42, YEL, { r: 0.21 });
  txt(s, "MITARBEIT", { x: 0.9, y: 2.05, w: 2.6, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Wie entsteht meine mündliche Note?", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 38, bold: true, color: WHITE });
  txt(s, "Die Kriterien – und wie du dich verbessern kannst", { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
  [6, 5, 4, 3, 2, 1].forEach((n, i) => badge(s, n, 0.9 + i * 1.05, 4.8, 0.75));
  s.addNotes("Einstieg: Die mündliche Note ist kein Bauchgefühl. Sie beruht auf festen Kriterien, die für alle gleich sind. Heute erkläre ich sie; danach teile ich euch eure Note mit.");
}

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
  const pts = ["Ich teile dir deine aktuelle mündliche Note mit – mit einer kurzen Begründung.", "Überleg dir: Welche Stufe willst du bis Weihnachten erreichen?", "Schreib einen Tipp von der letzten Folie ins Heft, an dem du arbeitest."];
  pts.forEach((x, i) => {
    txt(s, [{ text: `${i + 1}  `, options: { bold: true, color: YEL } }, { text: x }], { x: 0.9, y: 3.1 + i * 0.85, w: 11.5, h: 0.7, fontSize: 22, color: WHITE });
  });
  txt(s, "Grundlage: Kriterien zur Beurteilung der mündlichen Leistung (lehrerfreund.de)", { x: 0.9, y: 6.8, w: 11, h: 0.3, fontSize: 11, color: ICE });
  s.addNotes("Noten am besten einzeln mitteilen (z. B. während einer Arbeitsphase), nicht im Plenum.");
}

pres.writeFile({ fileName: "Muendliche_Noten_Kriterien.pptx" });
