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
  s.addNotes("Diese Folie läuft beim Ankommen. Vorher austeilen: Arbeitsblatt (1 pro Person), Stopp-Karten (1 pro Person). Pro Gruppentisch: Rollen- und Expertenkarten. Aufs Pult: Tippkarten, Situationskarten, Lösungsblätter (Lehrermaterial S. 2–3). Timer sichtbar einblenden (z. B. Handy unter der Dokumentenkamera oder ein Online-Timer). Möglichst ohne Einführung direkt zu Folie 2 wechseln.");
}

// ============ 2 Stummer Einstieg ============
{
  const s = base("Einstieg · still", "Ab welcher Zeile geht es schief?", "M1 auf deinem Arbeitsblatt · erst allein, dann mit dem Nachbarn");
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
  timeBadge(s, 10, "zu zweit");
  steps(s, [
    [{ text: "Tragt " }, b("mindestens drei"), { text: " Kipppunkte in die Tabelle ein: Zeile, Zitat, Fachbegriff." }],
    [{ text: "Markiert mit ★ den Kipppunkt, an dem man das Gespräch " }, b("am leichtesten"), { text: " hätte retten können." }],
    [{ text: "Vergleicht mit dem " }, b("Lösungsblatt am Pult"), { text: ". Ergänzt in einer anderen Farbe." }],
  ], 0.6, 2.0, 7.2, 0.95);
  box(s, 8.3, 1.95, 4.45, 3.1, LIGHT2);
  txt(s, "Werkzeuge aus den Modellen", { x: 8.55, y: 2.1, w: 4, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true });
  s.addText([
    "Sach- / Beziehungsebene", "Selbstkundgabe · Appell", "Du-Botschaft · „immer“ / „nie“", "Ironie = inkongruente Botschaft", "symmetrische Eskalation", "Interpunktion",
  ].map((x, i, a) => ({ text: x, options: { bullet: true, breakLine: i < a.length - 1 } })), { isTextBox: true, x: 8.55, y: 2.6, w: 4.0, h: 2.4, fontFace: BODY, fontSize: 15, color: NAVY, paraSpaceAfter: 4, margin: 0 });
  box(s, 0.6, 5.4, 12.15, 1.2, "FFF8EC");
  txt(s, "Nicht weitergekommen?", { x: 0.9, y: 5.52, w: 4, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: S.C });
  txt(s, "1. Nachbarn fragen   →   2. Tippkarte vom Pult holen   →   3. erst dann die Lehrkraft", { x: 0.9, y: 5.95, w: 11.5, h: 0.45, fontSize: 17 });
  s.addNotes("Während der Partnerarbeit herumgehen und nur zuhören. Wer fertig ist, holt sich das Lösungsblatt selbst (Lehrermaterial S. 2). Erwartung: Z. 3 (★ Du-Botschaft/Beziehungsohr), Z. 4 (Gegenangriff, symmetrische Eskalation), Z. 5 (Verallgemeinerung, Interpunktion), Z. 7 (Ironie, inkongruent), Z. 8–9 (nonverbaler Rückzug, man kann nicht nicht kommunizieren). Jetzt Vierergruppen bilden lassen (z. B. zwei Paare zusammen).");
}

// ============ 4 Gruppenpuzzle ============
{
  const s = base("Aufgabe 3 · Gruppenpuzzle", "Das Gespräch retten", "Arbeitsblatt S. 3 · Vierergruppen");
  timeBadge(s, 20, "Gruppe");
  txt(s, "Zuerst: Rollen verteilen", { x: 0.6, y: 1.95, w: 5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true });
  [["Zeit", "hat den Timer im Blick"], ["Protokoll", "schreibt auf"], ["Regie", "alle kommen zu Wort"], ["Material", "holt Karten · fragt als Einzige die Lehrkraft"]].forEach(([n, d], i) => {
    const x = 0.6 + i * 3.08;
    box(s, x, 2.45, 2.9, 1.05, LIGHT);
    txt(s, n, { x: x + 0.2, y: 2.52, w: 2.5, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true });
    txt(s, d, { x: x + 0.2, y: 2.92, w: 2.5, h: 0.55, fontSize: 13, color: MUTED });
  });
  const phases = [
    ["5'", "Lesen", "Jede Person liest allein eine Expertenkarte A–D."],
    ["8'", "Erklären", "Reihum: jede Strategie in höchstens 2 Minuten – mit Beispiel."],
    ["7'", "Umschreiben", "Dialog ab Zeile 3 neu schreiben, mindestens 3 Strategien. Checkliste prüfen."],
  ];
  phases.forEach(([m, h, d], i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 3.85, 3.9, 2.0, WHITE, { line: "B9C6E8", lw: 1.2 });
    txt(s, m, { x: x + 0.25, y: 4.0, w: 1.0, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true, color: RED });
    txt(s, h, { x: x + 1.2, y: 4.05, w: 2.5, h: 0.45, fontFace: HEAD, fontSize: 19, bold: true });
    txt(s, d, { x: x + 0.25, y: 4.6, w: 3.45, h: 1.2, fontSize: 15 });
  });
  box(s, 0.6, 6.1, 12.15, 0.6, NAVY);
  txt(s, "Wer fertig ist: Musterlösung am Pult vergleichen und die letzte Spalte von Aufgabe 2 ausfüllen.", { x: 0.9, y: 6.1, w: 11.6, h: 0.6, fontSize: 15, bold: true, color: WHITE, valign: "middle" });
  s.addNotes("Timer auf 20 Minuten, die Zeit-Person der Gruppe achtet auf die Teilphasen. Folie 5 (Werkzeugkasten) kann parallel als Überblick eingeblendet werden, sobald die Gruppen mit dem Umschreiben beginnen. Musterlösung: Lehrermaterial S. 3.");
}

// ============ 5 Werkzeugkasten ============
{
  const s = base("Überblick", "Der Werkzeugkasten", "Vier Strategien, um ein kippendes Gespräch zu retten");
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
  s.addNotes("Überblicksfolie, kein Vortrag. Sie kann während Aufgabe 3 und 4 stehen bleiben und dient als Gedächtnisstütze beim Schreiben der eigenen Szene.");
}

// ============ 6 Eigene Szene ============
{
  const s = base("Aufgabe 4 · Gruppe", "Eure eigene Szene", "Arbeitsblatt S. 4 · Planungsbogen");
  timeBadge(s, 20, "Gruppe");
  box(s, 0.6, 1.95, 5.9, 2.35, "FBECEB");
  txt(s, "Version A – es kippt", { x: 0.9, y: 2.08, w: 5.3, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: RED });
  txt(s, "harmloser Anfang  →  ein Satz mit Du-Vorwurf, „immer“ oder Ironie  →  Gegenangriff  →  Abbruch", { x: 0.9, y: 2.65, w: 5.3, h: 1.5, fontSize: 16 });
  box(s, 6.85, 1.95, 5.9, 2.35, "E6F4EC");
  txt(s, "Version B – es wird gelöst", { x: 7.15, y: 2.08, w: 5.3, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: S.B });
  txt(s, "gleicher Anfang  →  am Kipppunkt reagiert jemand anders (A · B · C · D)  →  Lösung", { x: 7.15, y: 2.65, w: 5.3, h: 1.5, fontSize: 16 });
  txt(s, "Spielregeln", { x: 0.6, y: 4.6, w: 5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true });
  s.addText([
    "2–4 Rollen, jede Version höchstens 2 Minuten",
    "Der Kipppunkt muss für das Publikum erkennbar sein",
    "Letzte 5 Minuten: einmal durchproben",
  ].map((x, i, a) => ({ text: x, options: { bullet: true, breakLine: i < a.length - 1 } })), { isTextBox: true, x: 0.6, y: 5.05, w: 7.5, h: 1.6, fontFace: BODY, fontSize: 16, color: NAVY, paraSpaceAfter: 6, margin: 0 });
  box(s, 8.6, 4.6, 4.15, 2.05, LIGHT2);
  txt(s, "Keine Idee?", { x: 8.85, y: 4.72, w: 3.7, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true });
  txt(s, "Material-Person zieht eine Situationskarte: Gruppenarbeit · Party · Chat · Nebenjob · Verein …", { x: 8.85, y: 5.15, w: 3.7, h: 1.4, fontSize: 14 });
  s.addNotes("Gruppen verteilen sich zum Proben ggf. auf Flur oder Ecken. Reihenfolge der Auftritte schon jetzt festlegen (z. B. Zettel an der Tafel), damit jede Gruppe weiß, wann sie moderiert: immer die Gruppe, die als nächste spielt.");
}

// ============ 7 Bühne frei ============
{
  const s = base("Aufgabe 5 · Bühne frei", "Freeze! Hier kippt es", "Arbeitsblatt S. 5 · Beobachtungsbogen");
  timeBadge(s, 30, "Plenum");
  steps(s, [
    [b("Version A"), { text: " wird gespielt." }],
    [{ text: "Du erkennst den Kipppunkt? " }, b("STOPP-Karte hoch!")],
    [{ text: "Moderation ruft " }, b("„Freeze!“"), { text: " – eine Person erklärt in einem Satz, was passiert ist." }],
    [b("Version B"), { text: " folgt direkt. Welche Strategie erkennst du?" }],
    [{ text: "Notiere alles im " }, b("Beobachtungsbogen"), { text: "." }],
  ], 0.6, 2.0, 8.2, 0.85);
  box(s, 9.3, 1.95, 3.45, 3.45, RED);
  txt(s, "✋", { x: 9.3, y: 2.15, w: 3.45, h: 1.0, fontSize: 48, color: WHITE, align: "center", valign: "middle" });
  txt(s, "STOPP", { x: 9.3, y: 3.15, w: 3.45, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "Hier kippt es!", { x: 9.3, y: 4.2, w: 3.45, h: 0.5, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
  box(s, 9.3, 5.6, 3.45, 1.05, NAVY);
  txt(s, "Moderation = die Gruppe, die als nächste spielt", { x: 9.5, y: 5.6, w: 3.05, h: 1.05, fontSize: 14, bold: true, color: WHITE, valign: "middle", align: "center" });
  s.addNotes("Sie setzen sich ins Publikum und notieren nur. Pro Gruppe etwa 4 Minuten. Wird zu früh oder zu spät gestoppt, ist das ein guter Anlass: Die Moderation fragt „Wer sieht es anders?“. Zeit knapp? Nur 3–4 Gruppen spielen, die übrigen zu Beginn der nächsten Stunde. Alternative: Alle spielen nur Version A mit Freeze, und das Publikum schlägt am Kipppunkt spontan eine bessere Reaktion vor.");
}

// ============ 8 Exit-Ticket ============
{
  const s = pres.addSlide();
  pageNo++;
  darkBg(s);
  pill(s, "EXIT-TICKET · 5 MIN.", 0.9, 1.2, 3.2, RED, 12);
  txt(s, "Bevor du gehst …", { x: 0.9, y: 1.8, w: 10, h: 0.8, fontFace: HEAD, fontSize: 36, bold: true, color: WHITE });
  numCircle(s, 1, 0.9, 3.0, RED, 0.55, 18);
  txt(s, "Formuliere als Ich-Botschaft mit Bitte:", { x: 1.7, y: 3.02, w: 10, h: 0.5, fontSize: 20, color: ICE });
  txt(s, "„Nie hörst du mir zu!“", { x: 1.7, y: 3.55, w: 10, h: 0.8, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE });
  numCircle(s, 2, 0.9, 4.75, RED, 0.55, 18);
  txt(s, "Woran erkennst du, dass ein Gespräch gerade kippt? Nenne zwei Signale.", { x: 1.7, y: 4.77, w: 10.5, h: 0.6, fontSize: 20, color: ICE });
  txt(s, "Abgabe an der Tür.", { x: 0.9, y: 6.4, w: 6, h: 0.4, fontSize: 16, italic: true, color: ICE });
  s.addNotes("Exit-Tickets an der Tür einsammeln. Erwartungshorizont: Lehrermaterial S. 3. Zwei oder drei gelungene Formulierungen können die nächste Stunde eröffnen.");
}

pres.writeFile({ fileName: "Praesentation_Kommunikationsstoerungen_loesen.pptx" }).then(() => console.log("ok"));
