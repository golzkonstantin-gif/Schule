const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Kommunikationsstörungen lösen";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F";
const S = { A: "3F7CC4", B: "2E9E6B", C: "E8A33D", D: "8E5BB5" };
const HEAD = "Cambria", BODY = "Calibri";
const FOOT = "Kommunikationsstörungen lösen · Deutsch 11";
let pageNo = 0;

function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 16, color: NAVY, margin: 0, valign: "top" }, o));
}
function base(kicker, title, sub) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 2 });
  txt(s, title, { x: 0.6, y: 0.72, w: 10.2, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  if (sub) txt(s, sub, { x: 0.6, y: 1.42, w: 10.2, h: 0.35, fontSize: 14, italic: true, color: MUTED });
  txt(s, FOOT, { x: 0.6, y: 7.0, w: 6, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo + 1), { x: 12.2, y: 7.0, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
  return s;
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08,
  });
}
function pill(s, label, x, y, w, col, size = 13) {
  box(s, x, y, w, 0.42, col, { r: 0.21 });
  txt(s, label, { x, y, w, h: 0.42, fontSize: size, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function numCircle(s, n, x, y, col = NAVY, d = 0.42, size = 14) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: size, bold: true, color: WHITE, align: "center", valign: "middle" });
}
// Zeit-Plakette oben rechts: ersetzt die mündliche Zeitansage
function timeBadge(s, min, sozial) {
  box(s, 10.9, 0.45, 1.85, 1.0, NAVY);
  txt(s, min + " Min.", { x: 10.9, y: 0.5, w: 1.85, h: 0.55, fontFace: HEAD, fontSize: 26, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, sozial, { x: 10.9, y: 1.02, w: 1.85, h: 0.35, fontSize: 12, color: ICE, align: "center", valign: "middle" });
}
function steps(s, list, x, y, w, rowH = 0.72, size = 17) {
  list.forEach((st, i) => {
    numCircle(s, i + 1, x, y + i * rowH + 0.02);
    s.addText(Array.isArray(st) ? st : [{ text: st }], { isTextBox: true, x: x + 0.6, y: y + i * rowH, w: w - 0.6, h: rowH, fontFace: BODY, fontSize: size, color: NAVY, valign: "top", margin: 0 });
  });
}
const b = (text) => ({ text, options: { bold: true } });
function darkBg(s) {
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
}

const dialog = [
  ["Lena", "Hey, hast du die Folien fertig? Wir sind morgen dran."],
  ["Tom", "Fast. Den Rest mach ich heute Abend."],
  ["Lena", "Heute Abend? Du hattest eine ganze Woche Zeit."],
  ["Tom", "Ich hatte halt dreimal Training. Nicht jeder hat so viel Freizeit wie du."],
  ["Lena", "Freizeit? Ich hab die komplette Gliederung gemacht! Immer bleibt alles an mir hängen."],
  ["Tom", "Dann mach’s doch allein, wenn du eh alles besser kannst."],
  ["Lena", "(verdreht die Augen) Super. Genau das hab ich von dir erwartet."],
  ["Tom", "(zuckt mit den Schultern, schaut aufs Handy) Wie du meinst."],
  ["Lena", "Ja, schau ruhig aufs Handy. Ist ja auch wichtiger."],
  ["Tom", "(steht auf und geht) Boah, du nervst echt."],
];

// ============ 1 Titel ============
{
  const s = pres.addSlide();
  darkBg(s);
  pill(s, "DOPPELSTUNDE", 0.9, 2.05, 2.4, RED, 12);
  txt(s, "Kommunikationsstörungen lösen", { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 38, bold: true, color: WHITE });
  txt(s, "Wo kippt ein Gespräch – und wie rettet man es?", { x: 0.9, y: 3.7, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
  [["A Ich-Botschaft", S.A], ["B Empathie", S.B], ["C Wunsch & Bitte", S.C], ["D Metakommunikation", S.D]]
    .forEach(([l, c], i) => pill(s, l, 0.9 + i * 2.55, 4.55, 2.4, c, 12));
  txt(s, "Deutsch, Klasse 11", { x: 0.9, y: 6.6, w: 5, h: 0.3, fontSize: 12, color: ICE });
  s.addNotes("Diese Folie läuft beim Ankommen. Die Klasse sitzt normal in den drei Reihen. Alles liegt bereits: Arbeitsblatt auf jedem Platz, je ein geschlossener Umschlag auf dem ersten Tisch jeder Reihe (4 Rollenkarten, 7 Exit-Tickets) – geöffnet wird er erst in Minute 43. Aufs Pult: Tippkarten. Bei Ihnen: Notfall-Ideen und Moderationskarten. Einziger Satz: „Heute steht alles auf den Folien – schaut einfach dorthin.“ Dann Folie 2 und Timer 5 Min.");
}

// ============ 2 Stummer Einstieg ============
{
  const s = base("Impuls · Einzelarbeit → Partnerarbeit", "Ab welcher Zeile geht es schief?", "M1 auf deinem Arbeitsblatt · erst allein, dann mit deinem Sitznachbarn");
  timeBadge(s, 5, "allein → zu zweit");
  dialog.forEach(([who, line], i) => {
    const y = 1.95 + i * 0.47;
    box(s, 0.6, y, 8.9, 0.42, i % 2 ? WHITE : LIGHT, { r: 0.04 });
    txt(s, String(i + 1), { x: 0.7, y, w: 0.4, h: 0.42, fontSize: 13, bold: true, color: MUTED, align: "center", valign: "middle" });
    txt(s, who, { x: 1.15, y, w: 0.8, h: 0.42, fontSize: 14, bold: true, valign: "middle" });
    txt(s, line, { x: 1.95, y, w: 7.5, h: 0.42, fontFace: HEAD, fontSize: 14, valign: "middle" });
  });
  box(s, 9.9, 1.95, 2.85, 4.6, RED);
  txt(s, "↯", { x: 9.9, y: 2.1, w: 2.85, h: 0.9, fontSize: 48, bold: true, color: WHITE, align: "center", valign: "middle" });
  s.addText([
    { text: "Markiere jede Stelle, an der das Gespräch kippt.", options: { breakLine: true, bold: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "3 Min. still.", options: { breakLine: true } },
    { text: "Dann 2 Min. vergleichen: Wo seid ihr euch uneinig?" },
  ], { isTextBox: true, x: 10.1, y: 3.1, w: 2.45, h: 3.3, fontFace: BODY, fontSize: 16, color: WHITE, valign: "top", margin: 0 });
  s.addNotes("Nicht erklären, nur Timer starten. Nach 3 Minuten ein Handzeichen oder Gong, dann 2 Minuten Partnervergleich. Es gibt hier keine Plenumsrunde – direkt weiter mit Folie 3.");
}

// ============ 3 Aufgabe 2 ============
{
  const s = base("Aufgabe 2 · Partnerarbeit", "Kipppunkte benennen", "Arbeitsblatt S. 2");
  timeBadge(s, 8, "zu zweit");
  steps(s, [
    [{ text: "Tragt " }, b("mindestens drei"), { text: " Kipppunkte in die Tabelle ein: Zeile, Zitat, Fachbegriff." }],
    [{ text: "Markiert mit ★ den Kipppunkt, an dem man das Gespräch " }, b("am leichtesten"), { text: " hätte retten können." }],
    [{ text: "Danach vergleichen wir im " }, b("Unterrichtsgespräch"), { text: ". Ergänzt in einer anderen Farbe." }],
  ], 0.6, 2.0, 7.2, 0.95);
  box(s, 8.3, 1.95, 4.45, 3.1, LIGHT2);
  txt(s, "Werkzeuge aus den Modellen", { x: 8.55, y: 2.1, w: 4, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true });
  s.addText([
    "Sach- / Beziehungsebene", "Selbstkundgabe · Appell", "Du-Botschaft · „immer“ / „nie“", "Ironie = inkongruente Botschaft", "symmetrische Eskalation", "Interpunktion",
  ].map((x, i, a) => ({ text: x, options: { bullet: true, breakLine: i < a.length - 1 } })), { isTextBox: true, x: 8.55, y: 2.6, w: 4.0, h: 2.4, fontFace: BODY, fontSize: 15, color: NAVY, paraSpaceAfter: 4, margin: 0 });
  box(s, 0.6, 5.4, 12.15, 1.2, "FFF8EC");
  txt(s, "Nicht weitergekommen?", { x: 0.9, y: 5.52, w: 4, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: S.C });
  txt(s, "1. Tippkarte 1 vom Pult   →   2. Tippkarte 2   →   3. erst dann die Lehrkraft", { x: 0.9, y: 5.95, w: 11.5, h: 0.45, fontSize: 17 });
  s.addNotes("Während der Partnerarbeit herumgehen und zuhören: Welche Paare finden Zeile 3? Danach 5 Minuten Unterrichtsgespräch: „In welcher Zeile kippt es zum ersten Mal – und warum?“ Wer geantwortet hat, nimmt die nächste Person dran. Zum Schluss Folie 4 als Sicherung zeigen.");
}

// ============ 4 Sicherung Kipppunkte ============
{
  const s = base("Sicherung · Unterrichtsgespräch", "Hier kippt das Gespräch", "Erst besprechen, dann diese Folie zeigen");
  timeBadge(s, 5, "Plenum");
  const rows = [
    ["3 ★", "„Du hattest eine ganze Woche Zeit.“", "Du-Botschaft statt Selbstkundgabe · Tom hört mit dem Beziehungsohr"],
    ["4", "„Nicht jeder hat so viel Freizeit wie du.“", "Gegenangriff · symmetrische Eskalation beginnt"],
    ["5", "„Immer bleibt alles an mir hängen.“", "Verallgemeinerung · Interpunktion"],
    ["7", "„Super. Genau das hab ich von dir erwartet.“", "Ironie = inkongruente Botschaft · Abwertung"],
    ["8–9", "Tom schaut aufs Handy – „Ist ja auch wichtiger.“", "Man kann nicht nicht kommunizieren · Rückzug wird als Desinteresse gedeutet"],
  ];
  rows.forEach(([z, q, why], i) => {
    const y = 1.95 + i * 0.9;
    box(s, 0.6, y, 12.15, 0.8, i === 0 ? "FBECEB" : LIGHT, { r: 0.05 });
    txt(s, z, { x: 0.75, y, w: 1.0, h: 0.8, fontFace: HEAD, fontSize: 20, bold: true, color: RED, align: "center", valign: "middle" });
    txt(s, q, { x: 1.9, y, w: 5.2, h: 0.8, fontFace: HEAD, fontSize: 15, italic: true, valign: "middle" });
    txt(s, why, { x: 7.3, y, w: 5.3, h: 0.8, fontSize: 15, bold: true, valign: "middle" });
  });
  s.addNotes("Erst nach dem Unterrichtsgespräch zeigen. ★ Zeile 3 ist der Kipppunkt, an dem man am leichtesten hätte eingreifen können. Überleitung: Genau diese Zeilen 3, 4, 5 und 7 schreibt gleich jede und jeder allein um – mit je einer Technik.");
}

// ============ 5 Technik-Training ============
{
  const s = base("Aufgabe 3 · Einzelarbeit", "Technik-Training: Schritt für Schritt", "Arbeitsblatt S. 3–4 · jede Technik an einem Kipppunkt aus M1");
  timeBadge(s, 18, "4 × 4' + 2' PA");
  const tech = [
    ["A", "Ich-Botschaft", "Zeile 3", "18'–22'"],
    ["B", "Empathie", "Zeile 4", "22'–26'"],
    ["C", "Wunsch & Bitte", "Zeile 5", "26'–30'"],
    ["D", "Metakommunikation", "Zeile 7", "30'–34'"],
  ];
  const w = 2.9, gap = 0.18;
  tech.forEach(([k, n, z, tm], i) => {
    const x = 0.6 + i * (w + gap);
    box(s, x, 1.95, w, 2.0, LIGHT);
    numCircle(s, k, x + 0.25, 2.12, S[k], 0.62, 22);
    txt(s, n, { x: x + 0.95, y: 2.12, w: w - 1.0, h: 0.62, fontFace: HEAD, fontSize: 13, bold: true, color: S[k], valign: "middle" });
    txt(s, "→ " + z + " umschreiben", { x: x + 0.25, y: 2.95, w: w - 0.5, h: 0.4, fontSize: 16, bold: true });
    txt(s, "4 Min.", { x: x + 0.25, y: 3.4, w: w - 0.5, h: 0.35, fontSize: 13, color: MUTED });
  });
  txt(s, "Bei jeder Technik", { x: 0.6, y: 4.25, w: 5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true });
  [["1", "Lesen", "Formel, So geht’s, WG-Beispiel"], ["2", "Anwenden", "den Kipppunkt aus M1 umschreiben"], ["3", "Prüfen", "Mini-Checkliste abhaken"]].forEach(([n, h, d], i) => {
    const x = 0.6 + i * 4.1;
    numCircle(s, n, x, 4.8, NAVY);
    txt(s, h, { x: x + 0.55, y: 4.78, w: 3.3, h: 0.45, fontFace: HEAD, fontSize: 18, bold: true });
    txt(s, d, { x: x + 0.55, y: 5.2, w: 3.3, h: 0.45, fontSize: 14, color: MUTED });
  });
  box(s, 0.6, 6.0, 12.15, 0.65, NAVY);
  txt(s, "Gong = nächste Technik.   Zum Schluss 2 Minuten: Vergleich mit dem Sitznachbarn.", { x: 0.9, y: 6.0, w: 11.6, h: 0.65, fontSize: 16, bold: true, color: WHITE, valign: "middle" });
  s.addNotes("Timer 4 Minuten, bei jedem Gong neu starten (22', 26', 30', 34'), dann 2 Minuten Partnervergleich. Herumgehen und mitlesen: Wer schreibt noch Du-Botschaften oder ‚immer‘? Pro Technik ein bis zwei gelungene Lösungen für das Unterrichtsgespräch vormerken. Schnelle Schülerinnen und Schüler: eine zweite Technik am selben Kipppunkt ausprobieren.");
}

// ============ 5 Werkzeugkasten ============
{
  const s = base("Sicherung · Unterrichtsgespräch 2", "Der Werkzeugkasten", "Vergleicht mit euren eigenen Lösungen auf S. 3–4");
  timeBadge(s, 7, "Plenum");
  const tools = [
    ["A", "Ich-Botschaft", S.A, "Wenn … fühle ich mich …, weil …", "„Ich werde nervös, weil wir morgen schon dran sind.“"],
    ["B", "Empathie", S.B, "Du meinst also … ? · Das klingt, als ob du … bist.", "„Du bist gestresst, weil alles klappen soll, oder?“"],
    ["C", "Wunsch & Bitte", S.C, "Beobachtung → Gefühl → Bedürfnis → Bitte", "„Kannst du mir die Folien bis 19 Uhr schicken?“"],
    ["D", "Metakommunikation", S.D, "Stopp – wie reden wir gerade miteinander?", "„Wir werfen uns nur noch Sachen an den Kopf.“"],
  ];
  const w = 2.9, gap = 0.18;
  tools.forEach(([k, n, c, f, ex], i) => {
    const x = 0.6 + i * (w + gap);
    box(s, x, 1.95, w, 4.1, LIGHT);
    numCircle(s, k, x + 0.25, 2.15, c, 0.7, 24);
    txt(s, n, { x: x + 0.2, y: 2.95, w: w - 0.3, h: 0.5, fontFace: HEAD, fontSize: 15, bold: true, color: c });
    txt(s, f, { x: x + 0.25, y: 3.5, w: w - 0.5, h: 1.1, fontSize: 15, bold: true });
    txt(s, ex, { x: x + 0.25, y: 4.7, w: w - 0.5, h: 1.2, fontFace: HEAD, fontSize: 14, italic: true, color: MUTED });
  });
  box(s, 0.6, 6.2, 12.15, 0.55, "FBECEB");
  txt(s, "Tabu: Du-Vorwürfe · „immer“ und „nie“ · Ironie · Gegenangriff · „Ja, aber …“", { x: 0.9, y: 6.2, w: 11.6, h: 0.55, fontSize: 15, bold: true, color: RED, valign: "middle" });
  s.addNotes("Unterrichtsgespräch 2 (7 Min.): Technik für Technik lesen ein bis zwei Personen ihre Lösung vor (gezielt die vorgemerkten drannehmen). Impuls: „Was ist daran gelungen – was sagt die Checkliste?“ Erst danach diese Folie zum Vergleich zeigen. Folie 7 (ganzer Dialog) nur, wenn Zeit bleibt.");
}

// ============ 7 Sicherung Version B ============
{
  const s = base("Sicherung · Unterrichtsgespräch", "So hätte es laufen können", "Eine mögliche Version B – eure darf ganz anders klingen");
  pill(s, "OPTIONAL", 10.9, 0.75, 1.85, MUTED, 12);
  const rows = [
    ["Lena", "Heute Abend? Ehrlich gesagt macht mich das nervös, weil wir morgen schon dran sind.", "A"],
    ["Tom", "Verstehe ich. Du willst sicher sein, dass alles passt, oder? Ich hab’s unterschätzt.", "B"],
    ["Lena", "Ich hab schon die Gliederung gemacht und fühl mich gerade ein bisschen allein damit.", "A"],
    ["Lena", "Kannst du mir die Folien bis 19 Uhr schicken? Dann proben wir morgen früh einmal.", "C"],
    ["Tom", "19 Uhr schaffe ich. Und nächstes Mal sagen wir uns früher Bescheid, wenn’s eng wird?", "D"],
  ];
  const names = { A: "Ich-Botschaft", B: "Empathie", C: "Bitte", D: "Metakomm." };
  rows.forEach(([who, line, k], i) => {
    const y = 1.95 + i * 0.9;
    box(s, 0.6, y, 12.15, 0.8, i % 2 ? WHITE : LIGHT, { r: 0.05 });
    txt(s, who, { x: 0.8, y, w: 0.9, h: 0.8, fontSize: 15, bold: true, valign: "middle" });
    txt(s, line, { x: 1.75, y, w: 8.4, h: 0.8, fontFace: HEAD, fontSize: 15, valign: "middle" });
    pill(s, k + " " + names[k], 10.35, y + 0.19, 2.2, S[k], 12);
  });
  s.addNotes("Optional, wenn Zeit bleibt: Hier sieht man, wie die Techniken im ganzen Gespräch zusammenspielen. Als eine mögliche Lösung zeigen, nicht als die richtige.");
}

// ============ 6 Eigene Szene ============
{
  const s = base("Aufgabe 4 · Reihen-Gruppe", "Eure eigene Szene", "Arbeitsblatt S. 5 · Zuerst: als Reihe zusammenrücken, Umschlag öffnen, Rollen verteilen");
  timeBadge(s, 16, "+ 3 Min. Umsetzen");
  box(s, 0.6, 1.95, 5.9, 2.35, "FBECEB");
  txt(s, "Version A – es kippt", { x: 0.9, y: 2.08, w: 5.3, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: RED });
  txt(s, "harmloser Anfang  →  ein Satz mit Du-Vorwurf, „immer“ oder Ironie  →  Gegenangriff  →  Abbruch", { x: 0.9, y: 2.65, w: 5.3, h: 1.5, fontSize: 16 });
  box(s, 6.85, 1.95, 5.9, 2.35, "E6F4EC");
  txt(s, "Version B – es wird gelöst", { x: 7.15, y: 2.08, w: 5.3, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: S.B });
  txt(s, "gleicher Anfang  →  am Kipppunkt reagiert jemand anders (A · B · C · D)  →  Lösung", { x: 7.15, y: 2.65, w: 5.3, h: 1.5, fontSize: 16 });
  txt(s, "Spielregeln", { x: 0.6, y: 4.6, w: 5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true });
  s.addText([
    "2–4 Rollen pro Version, jede Version höchstens 2 Minuten",
    "A und B spielen verschiedene Personen – wer nicht spielt, führt Regie",
    "Der Kipppunkt muss für das Publikum erkennbar sein",
    "Letzte 5 Minuten: einmal durchproben",
  ].map((x, i, a) => ({ text: x, options: { bullet: true, breakLine: i < a.length - 1 } })), { isTextBox: true, x: 0.6, y: 5.05, w: 7.5, h: 1.6, fontFace: BODY, fontSize: 16, color: NAVY, paraSpaceAfter: 6, margin: 0 });
  box(s, 8.6, 4.6, 4.15, 2.05, LIGHT2);
  txt(s, "Nach 5 Min. keine Idee?", { x: 8.85, y: 4.72, w: 3.7, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true });
  txt(s, "1. Tippkarte 1 zu Aufgabe 4 vom Pult\n2. Hilft nicht? Material-Person holt eine Notfall-Idee bei der Lehrkraft.", { x: 8.85, y: 5.15, w: 3.7, h: 1.4, fontSize: 14 });
  s.addNotes("Erst 3 Minuten Umsetzen: Wand-, Mittel- und Fensterreihe rücken zu je einer Gruppe zusammen, öffnen den Umschlag und verteilen die Rollen. Dann Timer 16 Minuten. Nach 5 Minuten eine Runde: Hat jede Gruppe eine eigene Situation? Notfall-Ideen nur ausgeben, wenn Tippkarte 1 nicht geholfen hat. Besetzung A (2–4), Besetzung B (2–4), Rest Regie. Um ca. Minute 57 eine Person als Moderation gewinnen und ihr die Moderationskarte geben.");
}

// ============ 7 Bühne frei ============
{
  const s = base("Aufgabe 5 · Bühne frei", "Spielen – ohne Unterbrechung", "Arbeitsblatt S. 6 · Beobachtungsbogen");
  timeBadge(s, 13, "Plenum");
  steps(s, [
    [{ text: "Reihenfolge: " }, b("Wandreihe → Mittelreihe → Fensterreihe"), { text: "." }],
    [b("Version A"), { text: ", direkt danach " }, b("Version B"), { text: " mit anderer Besetzung – je höchstens 2 Minuten." }],
    [b("Niemand unterbricht."), { text: " Das Publikum schaut genau hin: Wo kippt es? Was ist in B anders?" }],
    [{ text: "Applaus – dann " }, b("30 Sekunden"), { text: ", um die Zeile im Beobachtungsbogen auszufüllen." }],
  ], 0.6, 2.0, 8.2, 0.95);
  box(s, 9.3, 1.95, 3.45, 2.6, NAVY);
  txt(s, "Patenszene ★", { x: 9.5, y: 2.1, w: 3.05, h: 0.5, fontFace: HEAD, fontSize: 20, bold: true, color: WHITE });
  txt(s, "Wand beobachtet Mitte\nMitte beobachtet Fenster\nFenster beobachtet Wand", { x: 9.5, y: 2.65, w: 3.05, h: 1.8, fontSize: 15, color: WHITE });
  box(s, 9.3, 4.8, 3.45, 1.85, LIGHT2);
  txt(s, "Im Bogen notieren", { x: 9.5, y: 4.92, w: 3.05, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true });
  txt(s, "Kipppunkt · Ursache · Strategie in B · Wirkt es?", { x: 9.5, y: 5.35, w: 3.05, h: 1.2, fontSize: 15 });
  s.addNotes("Sie sitzen im Publikum und notieren gelungene Formulierungen. Die Moderation leitet mit der Moderationskarte. Pro Gruppe ca. 4 Minuten (2 × höchstens 2 Min. + Notizzeit). Achten Sie darauf, dass die Versionen kurz bleiben.");
}

// ============ 8 Auswertung ============
{
  const s = base("Auswertung", "Was haben wir gesehen?", "Beobachtungsbogen · eure Patenszene");
  timeBadge(s, 9, "still → Gruppe → Plenum");
  const st = [
    ["2'", "Still", "Ergänze deinen Beobachtungsbogen. Was ist dir bei deiner Patenszene aufgefallen?", LIGHT],
    ["2'", "Gruppe", "Einigt euch zu eurer Patenszene: Wo genau kippt es? Warum (Fachbegriff)? Welche Strategie hat in Version B am besten gewirkt?", LIGHT],
    ["5'", "Plenum", "Jede Gruppe: ca. 1,5 Minuten Rückmeldung zu ihrer Patenszene. Die Spielgruppe darf ergänzen.", LIGHT2],
  ];
  st.forEach(([m, h, d, f], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 1.95, 3.9, 3.6, f);
    txt(s, m, { x: x + 0.25, y: 2.1, w: 1.2, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true, color: RED });
    txt(s, h, { x: x + 0.25, y: 2.85, w: 3.4, h: 0.5, fontFace: HEAD, fontSize: 20, bold: true });
    txt(s, d, { x: x + 0.25, y: 3.45, w: 3.4, h: 2.0, fontSize: 15 });
  });
  box(s, 0.6, 5.85, 12.15, 0.8, NAVY);
  txt(s, "Sprechende Person im Plenum: die Regie-Person der Gruppe.   Die Moderation ruft auf.", { x: 0.9, y: 5.85, w: 11.6, h: 0.8, fontSize: 16, bold: true, color: WHITE, valign: "middle" });
  s.addNotes("Die Moderation führt durch die drei Schritte; Schritt 3 können Sie auch selbst moderieren. Nur zum Schluss (höchstens 1 Minute): zwei gelungene Formulierungen aus Ihren Notizen vorlesen. Wenn eine Gruppe den Kipppunkt anders sieht als die Spielgruppe, ist das ein produktiver Moment – kurz beide Sichtweisen stehen lassen.");
}

// ============ 11 Exit-Ticket: Mini-Test ============
{
  const s = pres.addSlide();
  pageNo++;
  darkBg(s);
  pill(s, "EXIT-TICKET · MINI-TEST · 6 MIN. · ALLEIN", 0.9, 0.8, 5.0, RED, 12);
  txt(s, "Jetzt allein: Kipppunkt finden und retten", { x: 0.9, y: 1.4, w: 11, h: 0.8, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE });
  const d = [["Mia", "Papa, kann ich am Samstag auf Leons Party?"], ["Vater", "Wie lange geht die denn?"], ["Mia", "Keine Ahnung, bis eins vielleicht."], ["Vater", "Du bist doch sowieso immer die Letzte, die nach Hause kommt."], ["Mia", "Boah, du vertraust mir echt nie!"]];
  d.forEach(([who, line], i) => {
    const y = 2.45 + i * 0.5;
    txt(s, String(i + 1), { x: 0.9, y, w: 0.4, h: 0.45, fontSize: 14, bold: true, color: ICE, valign: "middle" });
    txt(s, who, { x: 1.35, y, w: 1.0, h: 0.45, fontSize: 15, bold: true, color: WHITE, valign: "middle" });
    txt(s, line, { x: 2.4, y, w: 9.5, h: 0.45, fontFace: HEAD, fontSize: 16, color: WHITE, valign: "middle" });
  });
  numCircle(s, 1, 0.9, 5.15, RED, 0.5, 16);
  txt(s, "In welcher Zeile kippt es? Warum? (Fachbegriff)", { x: 1.6, y: 5.15, w: 10.5, h: 0.5, fontSize: 18, color: ICE, valign: "middle" });
  numCircle(s, 2, 0.9, 5.8, RED, 0.5, 16);
  txt(s, "Schreibe diese Zeile mit zwei Techniken um und notiere die Buchstaben.", { x: 1.6, y: 5.8, w: 10.5, h: 0.5, fontSize: 18, color: ICE, valign: "middle" });
  txt(s, "Ticket liegt im Gruppenumschlag · Abgabe an der Tür.", { x: 0.9, y: 6.7, w: 8, h: 0.4, fontSize: 14, italic: true, color: ICE });
  s.addNotes("Einzelarbeit, 6 Minuten, ohne Hilfe – das Format entspricht der späteren Testaufgabe. An der Tür einsammeln. Erwartungshorizont und Bewertungskriterien: Lehrermaterial S. 4. Die Auswertung zeigt, welche Technik in der nächsten Stunde noch einmal geübt werden muss.");
}

pres.writeFile({ fileName: "Praesentation_Kommunikationsstoerungen_loesen.pptx" }).then(() => console.log("ok"));
