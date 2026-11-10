// Leistungskontrolle „Merkmale einer Ballade“ (Mi 11.11., ca. 30 Min., kleiner Test) – Deutsch 7A
// Gruppe A und B, je Schülerfassung, Lösung und Fassung mit polnischer Übersetzung der Aufgaben (S-PL),
// dazu Erwartungshorizont mit Notenschlüssel. Text: Kopie LB S. 136 f. (Paddy Fingal, mit Verszählung) als Anlage.
// Druck ohne Farbflächen (Tinte sparen). Optik überarbeitet 10.11.: Aufgabenkästen, Gattungsfarben, Antwortfelder. Aufruf: node leistungskontrolle_balladen.js
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", VIO = "7A5BA6", SOLC = "2E7D4F", PLC = "5B4A8A", GRAY = "9AA3B5";
const PW = 11906 - 1700;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.it, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 80, line: o.line }, indent: o.indent, keepNext: o.keepNext });
const B = (style = BorderStyle.SINGLE, size = 6, color = "B9C6E8") => ({ style, size, color });
const borders = (b) => ({ top: b, bottom: b, left: b, right: b });
const NONE = borders(B(BorderStyle.NONE, 0, "FFFFFF"));
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, borders: o.borders || borders(B()), verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: o.ml ?? 100, right: o.ml ?? 100 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });

// ---------- Inhalte ----------
const G = {
  A: {
    wer: [["… will herausfinden, wer stärker ist.", "S"], ["… versteckt sich aus Angst.", "P"], ["… bittet den Fremden, leise zu sein.", "Sch"]],
    aus: [
      ["Der Schotte kommt, weil er Paddy um Hilfe bitten will.", false, "Str. 2, V. 8–9: „da wollt’ er erproben alsogleich, wer fester sei“", "Er will herausfinden, wer von beiden stärker ist."],
      ["Auf der Flucht wäre der Schotte beinahe ertrunken.", true, "Str. 10, V. 48: „fast wär ertrunken der gute Held“", ""],
    ],
    merk: [["Strophen mit Reimschema", "L", "strofy z układem rymów"], ["Ein Erzähler erzählt die Handlung.", "E", "narrator opowiada akcję"], ["Figuren sprechen unmittelbar, ohne Erzähler.", "D", "postacie mówią bezpośrednio, bez narratora"], ["Kehrreim am Ende jeder Strophe", "L", "refren na końcu każdej strofy"], ["Ein Erzähler leitet die Rede der Figuren ein („…“, rief er.)", "E", "narrator wprowadza słowa postaci („…“, zawołał.)"], ["Regieanweisungen in Klammern", "D", "wskazówki dla aktorów (didaskalia) w nawiasach"]],
    rStr: 1, rSol: "a a b b x · Paarreim",
    dStr: 7, dWho: "Schaya (Paddys Frau)",
    dWhy: "Die ganze Strophe ist Schayas Rede. Kein Erzähler leitet sie ein – die Figur spricht unmittelbar wie auf der Bühne.",
  },
  B: {
    wer: [["… sieht, dass der Fremde viel größer ist.", "P"], ["… versteckt Paddy im Bett.", "Sch"], ["… geht trotz der Warnung ans Bett.", "S"]],
    aus: [
      ["Paddy merkt, dass der Schotte viel größer ist als er.", true, "Str. 3, V. 14: „dass der Fremdling zehn Schuh größer war“", ""],
      ["Der Schotte bleibt vom Bett weg, weil Schaya ihn warnt.", false, "Str. 8, V. 37: „der Schotte neugierig ans Lager ging“", "Er geht trotz der Warnung neugierig ans Bett."],
    ],
    merk: [["Reim am Versende", "L", "rym na końcu wersu"], ["Ein Er-Erzähler berichtet, was die Figuren tun.", "E", "narrator w 3. osobie opowiada, co robią postacie"], ["Rollenname vor jeder Rede (SCHAYA: …)", "D", "imię postaci przed każdą wypowiedzią (SCHAYA: …)"], ["Paarreim (aabb)", "L", "rym parzysty (aabb)"], ["Ein Erzähler leitet die Rede der Figuren ein.", "E", "narrator wprowadza słowa postaci"], ["Nur Figurenrede, kein Erzähler", "D", "tylko słowa postaci, bez narratora"]],
    rStr: 11, rSol: "a a b b x · Paarreim",
    dStr: 9, dWho: "der Schotte",
    dWhy: "Die ganze Strophe ist die Rede des Schotten. Kein Erzähler leitet sie ein – die Figur spricht unmittelbar wie auf der Bühne.",
  },
};
const FIG = [["P", "Paddy"], ["Sch", "Schaya"], ["S", "der Schotte"]];
const FLUCHT = "Er hält Paddy für ein Baby und denkt: Wenn schon das Baby so riesig ist, muss der Vater ein noch größerer Riese sein.";
const GN = { L: "Lyrik", E: "Epik", D: "Dramatik" };

// Polnische Übersetzung der Aufgabenstellungen (Fachbegriffe bleiben deutsch, keine Lösungen)
const PL = {
  intro: "Przeczytaj dokładnie każde zadanie. Pisz starannie. Masz ok. 30 minut. Zadania 1, 2, 4 i 5 dotyczą ballady „Paddy Fingal“ (tekst w załączniku, numeracja wersów na marginesie). Odpowiadasz po niemiecku.",
  a1a: "Kto to robi? Przy każdym zdaniu zaznacz krzyżykiem jedną postać: Paddy, Schaya albo der Schotte (Szkot).",
  a1b: "Dlaczego Szkot ucieka? Uzasadnij jednym zdaniem po niemiecku.",
  a2: "Czy zdanie jest prawdziwe (richtig) czy fałszywe (falsch)? Zaznacz. Udowodnij cytatem: strofa (Str.), wers (V.) i dokładne słowa z tekstu w cudzysłowie „…“. Jeśli zdanie jest fałszywe, napisz, jak jest naprawdę (Richtig ist: …).",
  a3: "Przy każdej cesze zaznacz krzyżykiem jeden rodzaj literacki: ten, dla którego ta cecha jest typowa (Lyrik, Epik albo Dramatik).",
  a4a: "Ile strof ma ballada? Ile wersów ma każda strofa?",
  a4b: (n) => `Określ układ rymów w strofie ${n}: zapisz literę przy każdym wersie (a, b …; wers bez rymu = x). Podaj nazwę rymu.`,
  a4c: "Jakimi dwoma słowami kończy się każda strofa (Kehrreim)?",
  a5a: "Kto opowiada? Zaznacz: Er-Erzähler (narrator w 3. osobie) albo Ich-Erzähler (narrator w 1. osobie).",
  a5b: (n) => `Kto mówi w strofie ${n}?`,
  a5c: (n) => `Uzasadnij jednym zdaniem po niemiecku, dlaczego strofa ${n} to cecha dramatu (Dramatik).`,
  z: "Zadanie dodatkowe: Uzasadnij jednym zdaniem z „weil“, dlaczego „Paddy Fingal“ to ballada. Podkreśl czasownik odmieniony (das finite Verb).",
};
const s = (x, o = {}) => t(x, Object.assign({ size: 26 }, o));
const pl = (x) => p([t("PL: ", { size: 20, bold: true, color: PLC }), t(x, { size: 20, it: true, color: PLC })], { after: 40 });
const solR = (x) => t(x, { size: 24, it: true, color: SOLC, bold: true });
const lineP = (sol, txt, o = {}) => sol && txt ? p([solR(txt)], { before: 140, after: 0 }) : p(t("_".repeat(o.n || 86), { color: GRAY }), { before: o.before ?? 220, after: 0 });
const box = "☐", boxX = "☒";
function head(nr, title, pts) {
  return p([t(`Aufgabe ${nr}  `, { font: "Cambria", size: 30, bold: true, color: RED }), t(title, { font: "Cambria", size: 30, bold: true, color: NAVY }), t(`     ___ / ${pts} P.`, { size: 22, color: MUTED })], { before: 220, after: 60, keepNext: true });
}
const sub = (l, runs, pts) => p([t(l + "  ", { bold: true, color: NAVY }), ...runs, ...(pts ? [t(`  (${pts} P.)`, { size: 20, color: MUTED })] : [])], { after: 40, keepNext: true });
const kreuz = (opts, right, sol) => p(opts.map((o, i) => t(`${sol && o === right ? boxX : box} ${o}${i < opts.length - 1 ? "        " : ""}`, { size: 26, bold: sol && o === right, color: sol && o === right ? SOLC : "000000" })), { before: 60, after: 60 });

// ---------- Gestaltung (Stand 10.11.: Optik verbessert) ----------
// Jede Aufgabe in einem eigenen Kasten mit farbigem Rahmen (keine Farbflächen), Punkte rechts im Kästchen,
// Gattungen mit Farbe, Symbol und Vergleich, Antwortfelder als Kästen statt Strichlinien.
const GC = { L: "2A8C8C", E: VIO, D: "C77C12" };
const GS = { L: "♪", E: "✎", D: "◆" };
const GV = { L: "wie ein Song-Text", E: "wie ein Spielbericht", D: "wie eine Serienszene" };
const FB = borders(B(BorderStyle.SINGLE, 6, "9AA3B5"));
// Antwortfeld (leeres Kästchen bzw. Lösung)
const field = (w, h, sol, solTxt, o = {}) => cell(p(sol && solTxt ? solR(solTxt) : t(""), { after: 0, align: o.center ? AlignmentType.CENTER : undefined }), { w, borders: FB, ml: 100, m: 30 });
const lbl = (x, w, o = {}) => cell(p(s(x, o), { after: 0, align: o.right ? AlignmentType.RIGHT : undefined }), { w, borders: NONE, ml: 60, m: 30 });
const gapC = (w) => cell(p(t(""), { after: 0 }), { w, borders: NONE, m: 0 });
const line = (cells, h = 480) => table(cells.map((c) => c.w), [row(cells.map((c) => c.c), h, HeightRule.ATLEAST)]);
const L = (x, w, o) => ({ c: lbl(x, w, o), w });
const F = (w, sol, solTxt, o) => ({ c: field(w, 0, sol, solTxt, o), w });
const G0 = (w) => ({ c: gapC(w), w });
const bigField = (w, h, sol, solTxt) => table([w], [row([cell(p(sol && solTxt ? solR(solTxt) : t(""), { after: 0 }), { w, borders: FB, ml: 140, m: 80, valign: VerticalAlign.TOP })], h)]);
const spacer = (n = 80) => p(t("", { size: 8 }), { after: n });
const IW = PW - 300; // Innenbreite der Aufgabenkästen
function aufgabe(nr, title, pts, col, children, o = {}) {
  const BW = o.w || PW, iw = BW - 300;
  const head = table([iw - 1450, 1450], [row([
    cell(p([t(nr ? `Aufgabe ${nr}   ` : "Zusatz   ", { font: "Cambria", size: 28, bold: true, color: col }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), ...(o.sub ? [t("   " + o.sub, { size: 20, it: true, color: col })] : [])], { after: 0 }), { w: iw - 1450, borders: NONE, ml: 0, m: 0 }),
    cell(p([t("___ / " + pts + " P.", { size: 21, bold: true, color: NAVY })], { after: 0, align: AlignmentType.CENTER }), { w: 1450, borders: borders(B(BorderStyle.SINGLE, 10, col)), m: 30 }),
  ])]);
  const box = table([BW], [new TableRow({ cantSplit: true, children: [cell([head, spacer(20), ...children], { w: BW, borders: borders(B(o.dashed ? BorderStyle.DASHED : BorderStyle.SINGLE, 14, col)), m: 90, ml: 150, valign: VerticalAlign.TOP })] })]);
  return o.raw ? box : [box, spacer(70)];
}
const subA = (l, txt, pts, col = NAVY) => p([t(l + "   ", { bold: true, size: 26, color: col }), s(txt), ...(pts ? [t(`  (${pts} P.)`, { size: 20, color: MUTED })] : [])], { before: 40, after: 40, keepNext: true });
const kreuzL = (opts, right, sol) => p(opts.map((o, i) => t(`${sol && o === right ? "☒" : "☐"}  ${o}${i < opts.length - 1 ? "          " : ""}`, { size: 26, bold: sol && o === right, color: sol && o === right ? SOLC : "000000" })), { before: 20, after: 50 });

function test(g, mode) { // mode: "s" Schüler, "l" Lösung, "pl" polnisch
  const D = G[g], sol = mode === "l", PLm = mode === "pl";
  const d = [];
  // Kopf
  d.push(table([PW - 2700, 2700], [row([
    cell([
      p(t("TEST · DEUTSCH · KLASSE 7A", { size: 18, color: MUTED, bold: true }), { after: 20 }),
      p(t("Paddy Fingal unter der Lupe", { font: "Cambria", bold: true, size: 38, color: NAVY }), { after: 0 }),
      p([t("Merkmale einer Ballade  ·  ", { font: "Cambria", size: 24, color: MUTED }), t("♪ Lyrik", { font: "Cambria", size: 24, bold: true, color: GC.L }), t("  ✎ Epik", { font: "Cambria", size: 24, bold: true, color: GC.E }), t("  ◆ Dramatik", { font: "Cambria", size: 24, bold: true, color: GC.D })], { after: 80 }),
      line([L("Name:", 900), F(4100, false), G0(200), L("Datum: 11.11.", 2000)]),
    ], { w: PW - 2700, borders: NONE, m: 0, ml: 0, valign: VerticalAlign.TOP }),
    cell([p(t(`Gruppe ${g}`, { font: "Cambria", bold: true, size: 40, color: RED }), { after: 40, align: AlignmentType.CENTER }), ...(sol ? [p(t("LÖSUNG", { bold: true, size: 24, color: SOLC }), { after: 30, align: AlignmentType.CENTER })] : []), p(s("Punkte: ____ / 26", { size: 22 }), { after: 40, align: AlignmentType.CENTER }), p(s("Zusatz: ____ / 2", { size: 22 }), { after: 40, align: AlignmentType.CENTER }), p(s("Note: ______", { size: 22 }), { after: 0, align: AlignmentType.CENTER })], { w: 2700, borders: borders(B(BorderStyle.SINGLE, 14, RED)), m: 100 }),
  ])]));
  d.push(p(s("Lies jede Aufgabe genau und schreibe sauber. Ca. 30 Minuten. Der Text „Paddy Fingal“ liegt bei (Verszählung am Rand).", { it: true, color: MUTED, size: 21 }), { before: 60, after: PLm ? 10 : 40 }));
  if (PLm) { d.push(pl(PL.intro)); d.push(spacer(20)); }

  d.push(p([s("Die Aufgaben 1, 2, 4 und 5 beziehen sich auf die Ballade „Paddy Fingal“.", { bold: true, size: 23, color: NAVY })], { after: 90 }));

  // Aufgabe 1 · Inhalt
  const a1 = [subA("a", "Wer tut das in „Paddy Fingal“? Kreuze bei jeder Zeile genau eine Figur an.", 3)];
  if (PLm) a1.push(pl(PL.a1a));
  const WW = [IW - 3 * 1450, 1450, 1450, 1450];
  const fh = (x, w) => cell(p(t(x, { bold: true, size: 24, color: NAVY, font: "Cambria" }), { after: 0, align: AlignmentType.CENTER }), { w, borders: borders(B(BorderStyle.SINGLE, 10, NAVY)), m: 30 });
  a1.push(table(WW, [
    row([fh("Wer …", WW[0]), ...FIG.map(([, n], j) => fh(n, WW[j + 1]))], 400),
    ...D.wer.map(([x, k], i) => row([cell(p([t(`${i + 1}   `, { bold: true, size: 24, color: NAVY }), s(x, { size: 23 })], { after: 0 }), { w: WW[0], ml: 120, m: 20 }), ...FIG.map(([c], j) => cell(p(t(sol && c === k ? "✗" : "☐", { size: 26, bold: sol && c === k, color: sol && c === k ? SOLC : "B9C6E8" }), { after: 0, align: AlignmentType.CENTER }), { w: WW[j + 1], m: 20 }))], 380)),
  ]));
  a1.push(subA("b", "Warum flieht der Schotte am Ende? Begründe in einem Satz.", 2));
  if (PLm) a1.push(pl(PL.a1b));
  a1.push(bigField(IW, 700, sol, FLUCHT));
  d.push(...aufgabe(1, "Worum geht es in „Paddy Fingal“?", 5, NAVY, a1));

  // Aufgabe 2 · Aussagen prüfen
  const a2 = [p(s("Ist die Aussage richtig oder falsch? Kreuze an. Belege mit Strophe, Vers und Zitat. Ist die Aussage falsch, schreibe auf, was richtig ist."), { after: 30, keepNext: true })];
  if (PLm) a2.push(pl(PL.a2));
  D.aus.forEach(([a, ok, z, r], i) => {
    a2.push(p([t(`${["a", "b"][i]}   `, { bold: true, size: 26, color: RED }), s("„" + a + "“"), t("  (3 P.)", { size: 20, color: MUTED })], { before: 50, after: 20, keepNext: true }));
    a2.push(kreuzL(["richtig", "falsch"], ok ? "richtig" : "falsch", sol));
    a2.push(line([L("Beleg:", 1100), F(IW - 1100, sol, z)]));
    a2.push(spacer(20));
    a2.push(line([L("Falls falsch – richtig ist:", 2900, { size: 22 }), F(IW - 2900, sol, ok ? "–" : r)]));
  });
  d.push(...aufgabe(2, "Aussagen zu „Paddy Fingal“ prüfen", 6, RED, a2));

  // Aufgabe 3 · Gattungen
  const W = [IW - 3 * 1450, 1450, 1450, 1450];
  const GPL = { L: "liryka", E: "epika", D: "dramat" };
  const gh = (k, w) => cell([p(t(`${GS[k]} ${GN[k]}`, { bold: true, size: 24, color: GC[k], font: "Cambria" }), { after: 0, align: AlignmentType.CENTER }), p(t(PLm ? GPL[k] : GV[k], { size: 17, it: true, color: PLm ? PLC : MUTED }), { after: 0, align: AlignmentType.CENTER })], { w, borders: borders(B(BorderStyle.SINGLE, 10, GC[k])), m: 30 });
  const a3 = [p(s("Kreuze bei jedem Merkmal genau eine Gattung an: die Gattung, für die das Merkmal typisch ist."), { after: 40, keepNext: true })];
  if (PLm) a3.push(pl(PL.a3));
  a3.push(table(W, [
    row([cell(p(t("Merkmal", { bold: true, size: 24, color: NAVY }), { after: 0 }), { w: W[0], borders: borders(B(BorderStyle.SINGLE, 10, NAVY)), ml: 120 }), gh("L", W[1]), gh("E", W[2]), gh("D", W[3])], 420),
    ...D.merk.map(([m, k, mpl], i) => row([cell([p([t(`${i + 1}   `, { bold: true, size: 24, color: NAVY }), s(m, { size: 23 })], { after: 0 }), ...(PLm ? [p(t("PL: " + mpl, { size: 19, it: true, color: PLC }), { after: 0, indent: { left: 300 } })] : [])], { w: W[0], ml: 120, m: 20 }), ...["L", "E", "D"].map((x, j) => cell(p(t(sol && x === k ? "✗" : "☐", { size: 26, bold: sol && x === k, color: sol && x === k ? SOLC : "B9C6E8" }), { after: 0, align: AlignmentType.CENTER }), { w: W[j + 1], m: 20 }))], 360)),
  ]));
  d.push(...aufgabe(3, "Merkmale den Gattungen zuordnen", 6, NAVY, a3));

  // Aufgabe 4 · Lyrik und Aufgabe 5 · Epik/Dramatik nebeneinander
  const HB = Math.floor((PW - 160) / 2), hw = HB - 300;
  const sch = D.rSol.split(" · "), lt = sch[0].split(" ");
  const a4 = [subA("a", "Wie viele Strophen hat „Paddy Fingal“? Wie viele Verse hat jede Strophe?", 2, GC.L)];
  if (PLm) a4.push(pl(PL.a4a));
  a4.push(line([L("Strophen:", 1200), F(800, sol, "11", { center: true }), G0(200), L("Verse:", 900), F(800, sol, "5", { center: true })]));
  a4.push(subA("b", `Bestimme das Reimschema von Strophe ${D.rStr} (Vers ohne Reim = x) und nenne die Reimart.`, 2, GC.L));
  if (PLm) a4.push(pl(PL.a4b(D.rStr)));
  a4.push(line([...[1, 2, 3, 4, 5].flatMap((n) => [L(String(n), 280, { right: true }), F(520, sol, lt[n - 1], { center: true })])]));
  a4.push(spacer(20));
  a4.push(line([L("Reimart:", 1200), F(hw - 1200, sol, sch[1])]));
  a4.push(subA("c", "Mit welchen zwei Wörtern endet jede Strophe? (Kehrreim)", 1, GC.L));
  if (PLm) a4.push(pl(PL.a4c));
  a4.push(line([F(hw, sol, "„Paddy Fingal“")]));
  const a5 = [subA("a", "Wer erzählt „Paddy Fingal“?", 1, GC.E)];
  if (PLm) a5.push(pl(PL.a5a));
  a5.push(kreuzL(["Er-Erzähler", "Ich-Erzähler"], "Er-Erzähler", sol));
  a5.push(subA("b", `Wer spricht in Strophe ${D.dStr}?`, 1, GC.D));
  if (PLm) a5.push(pl(PL.a5b(D.dStr)));
  a5.push(line([F(hw, sol, D.dWho)]));
  a5.push(subA("c", `Begründe in einem Satz, warum Strophe ${D.dStr} ein Merkmal der Dramatik zeigt.`, 2, GC.D));
  if (PLm) a5.push(pl(PL.a5c(D.dStr)));
  a5.push(bigField(hw, 1100, sol, D.dWhy));
  d.push(table([HB, 160, HB], [new TableRow({ cantSplit: true, children: [
    cell([aufgabe(4, "Paddy Fingal als Gedicht", 5, GC.L, a4, { sub: "♪", w: HB, raw: true })], { w: HB, borders: NONE, m: 0, ml: 0, valign: VerticalAlign.TOP }),
    cell(p(t(""), { after: 0 }), { w: 160, borders: NONE, m: 0, ml: 0 }),
    cell([aufgabe(5, "Erzähler oder nicht?", 4, GC.E, a5, { sub: "✎ ◆", w: HB, raw: true })], { w: HB, borders: NONE, m: 0, ml: 0, valign: VerticalAlign.TOP }),
  ] })]));
  d.push(spacer(70));

  // Zusatz
  const az = [p(s("Begründe in einem Satz mit „weil“, warum „Paddy Fingal“ eine Ballade ist. Unterstreiche das finite Verb im weil-Satz."), { after: 60 })];
  if (PLm) az.push(pl(PL.z));
  az.push(sol ? p([solR("z. B. „Paddy Fingal“ ist eine Ballade, weil ein Erzähler eine spannende Geschichte in Strophen mit Reimen "), t("erzählt", { size: 24, it: true, bold: true, color: SOLC, u: true }), solR(".")], { before: 80, after: 0 }) : bigField(IW, 700, false));
  d.push(...aufgabe(0, "Für Profis: weil-Satz", 2, MUTED, az, { dashed: true }));
  d.pop(); // kein Leerabsatz am Ende
  return d;
}

function erwartung() {
  const d = [];
  d.push(p(t("Test „Merkmale einer Ballade“ · Erwartungshorizont und Notenschlüssel", { font: "Cambria", bold: true, size: 30, color: NAVY }), { after: 60 }));
  d.push(p(s("Mi 11.11. · ca. 30 Min. · kleiner Test (nicht die Analyse-Bewertung) · Inhalt (11 P.) und Merkmale (15 P.) · Gruppe A und B gleich schwer, nur Handlungen, Aussagen, Strophen und Merkmale getauscht · Text: Kopie LB S. 136 f. mit Verszählung als Anlage.", { size: 20, color: MUTED }), { after: 120 }));
  const W = [1300, 4400, 4506 - 0];
  const th = (x, w) => cell(p(t(x, { bold: true, color: NAVY, size: 20 }), { after: 0 }), { w, borders: borders(B(BorderStyle.SINGLE, 8, NAVY)) });
  const tc = (x, w, o = {}) => cell(p(t(x, { size: 19, bold: o.b, color: o.c }), { after: 0 }), { w });
  const rows = [
    ["1 a (3)", "A: 1 der Schotte · 2 Paddy · 3 Schaya", "B: 1 Paddy · 2 Schaya · 3 der Schotte · je Zeile 1 P."],
    ["1 b (2)", "Er hält Paddy für ein Baby und schließt daraus, dass der Vater ein noch größerer Riese sein muss (Str. 9).", "2 P. für Baby + Schluss auf den riesigen Vater; 1 P. nur für „Er hat Angst“ oder nur „Baby“"],
    ["2 (6)", "A: a falsch – Str. 2, V. 8–9 „… wer fester sei“ – richtig ist: Er will herausfinden, wer stärker ist · b richtig – Str. 10, V. 48 „fast wär ertrunken der gute Held“", "B: a richtig – Str. 3, V. 14 „dass der Fremdling zehn Schuh größer war“ · b falsch – Str. 8, V. 37 „der Schotte neugierig ans Lager ging“ – richtig ist: Er geht trotz Warnung ans Bett · je Aussage: Urteil (bei falschen mit Verbesserung) 1 P., Stelle 1 P., Zitat 1 P.; andere passende Stellen gelten"],
    ["3 (6)", "A: 1 L · 2 E · 3 D · 4 L · 5 E · 6 D", "B: 1 L · 2 E · 3 D · 4 L · 5 E · 6 D · je richtiges Kreuz 1 P."],
    ["4 (5)", "11 Strophen · 5 Verse · a a b b x · Paarreim · Kehrreim „Paddy Fingal“", "je 1 P. (Schema ohne x, aber aabb richtig: 1 P.)"],
    ["5 (4)", "Er-Erzähler · A Str. 7: Schaya / B Str. 9: der Schotte · Begründung: Die Figur spricht unmittelbar, ohne Erzähler (wie auf der Bühne)", "a 1 P., b 1 P., c 2 P. (1 P. nur für „wörtliche Rede“ – allein kein Merkmal der Dramatik)"],
    ["Zusatz (2)", "weil-Satz mit Merkmal(en), finites Verb am Ende unterstrichen", "1 P. Inhalt, 1 P. Satzbau (Komma, Verb am Ende, unterstrichen)"],
  ];
  d.push(table(W, [row([th("Aufgabe", W[0]), th("Erwartung", W[1]), th("Bewertung", W[2])], 380), ...rows.map(([a, b, c]) => row([tc(a, W[0], { b: true }), tc(b, W[1]), tc(c, W[2])], 360))]));
  d.push(p(t("Notenschlüssel (Vorschlag, gleiche Prozentstufen wie beim Verben-Test – an den Schulschlüssel anpassen)", { font: "Cambria", bold: true, size: 24, color: NAVY }), { before: 240, after: 80 }));
  const NK = [["1", "26 – 24,5"], ["2", "24 – 21"], ["3", "20,5 – 17"], ["4", "16,5 – 13"], ["5", "12,5 – 6,5"], ["6", "6 – 0"]];
  d.push(table([1800, 3000], [row([th("Note", 1800), th("Punkte (von 26)", 3000)]), ...NK.map(([n, x]) => row([tc(n, 1800, { b: true }), tc(x, 3000)], 340))]));
  d.push(p(s("Zusatzpunkte werden zur Gesamtpunktzahl addiert (höchstens bis 26).", { color: MUTED, size: 19 }), { before: 80, after: 0 }));
  d.push(p(s("S-PL: gleicher Test (Gruppe A oder B) mit polnischer Übersetzung der Aufgabenstellungen; Antworten auf Deutsch, Fachbegriffe deutsch, gleiche Bewertung.", { size: 19 }), { before: 120, after: 0 }));
  return d;
}

const footer = (label) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })] });
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 450, left: 850, right: 850, footer: 250 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const save = (name, children, label) => Packer.toBuffer(new Document({ styles, sections: [{ properties: props, footers: { default: footer(label) }, children }] })).then((buf) => fs.writeFileSync(__dirname + "/../" + name + ".docx", buf));
for (const g of ["A", "B"]) {
  save(`LK_Balladen_Gruppe_${g}`, test(g, "s"), `Test Balladen · Gruppe ${g}`);
  save(`LK_Balladen_Gruppe_${g}_Loesung`, test(g, "l"), `Test Balladen · Gruppe ${g} · Lösung`);
  save(`LK_Balladen_Gruppe_${g}_Polnisch`, test(g, "pl"), `Test Balladen · Gruppe ${g} · z polskim tłumaczeniem`);
}
save("LK_Balladen_Erwartungshorizont_Notenschluessel", erwartung(), "Test Balladen · Erwartungshorizont");
