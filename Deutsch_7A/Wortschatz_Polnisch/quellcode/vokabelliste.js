const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak, Header, Footer, TabStopType,
  PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", BOARD = "F3F6F1";
const C = { hv: "D9534F", mv: "2E9E6B", p2: "E8A33D", inf: "3F7CC4" };
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext, indent: o.indent });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h2 = (label, text, after = 80) => p([t(label + "  ", { font: "Cambria", size: 26, bold: true, color: C.hv }), t(text, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 200, after, keepNext: true });
const h3 = (text, col = NAVY) => p(t(text, { font: "Cambria", size: 23, bold: true, color: col }), { before: 120, after: 60, keepNext: true });
const gap = (n = 80) => p(t(""), { after: n });
const br = () => new Paragraph({ children: [new PageBreak()] });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };
const thin = allBorders(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || thin,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font, italics: o.it }), { after: 0, align: o.align }), { w, fill: o.fill });

// simple grid table: header + rows of strings/runs
function grid(widths, head, rows, o = {}) {
  return table(widths, [
    row(head.map((h, i) => hdr(h, widths[i], o.headFill ? o.headFill[i] : NAVY)), 420),
    ...rows.map((r) => row(r.map((v, i) => tc(v, widths[i], { font: o.fonts && o.fonts[i], align: o.align && o.align[i], bold: o.bold && o.bold[i], color: o.colors && o.colors[i] })), o.h || 460)),
  ]);
}

function box(children, fill = LIGHT2, border) {
  return table([W], [row([cell(children, { w: W, fill, borders: border || allBorders(none), m: 140, ml: 180 })])]);
}
// Merksatz box
function merk(nr, title, runs) {
  return box([
    p([t(`Merksatz ${nr} · `, { bold: true, color: C.hv, size: 20 }), t(title, { bold: true, color: NAVY, size: 22, font: "Cambria" })], { after: 60 }),
    ...runs.map((r) => p(r, { after: 40 })),
  ], LIGHT2);
}
// Tafel box
function tafel(title, children) {
  return table([W], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w: W, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)) })]),
    row([cell(children, { w: W, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 140, ml: 200 })]),
  ]);
}
const R = (s) => t(s, { bold: true, color: C.hv, font: "Cambria", size: 22 });
const Y = (s) => t(s, { bold: true, color: "C98A1E", font: "Cambria", size: 22 });
const B = (s) => t(s, { bold: true, color: C.inf, font: "Cambria", size: 22 });
const N = (s, o = {}) => t(s, Object.assign({ font: "Cambria", size: 22 }, o));
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));


// ================================================================
// Gemeinsame Bausteine für Tafelskripte
// ================================================================

const header = (label) => new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 7A · " + label, { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })],
  })],
});
const footer = (label) => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })],
});

// ---------- Was die Klasse gerade tut (nur zur Planung) ----------
const LABELS = {
  zuhoeren: "hört zu", gemeinsam: "macht mündlich mit", abschreiben: "schreibt ab", luecken: "arbeitet allein",
  tafel: "kommt an die Tafel", kontrolle: "kontrolliert", partner: "arbeitet zu zweit", fertig: "gibt ab",
};
function labelCell(m, w) {
  return cell([
    p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }),
    p(t(LABELS[m], { bold: true, color: NAVY, size: 20 }), { after: 0 }),
  ], { w, fill: LIGHT, borders: allBorders(solid("D5DDEE", 6)), m: 100, valign: VerticalAlign.TOP });
}

// Schritt: Modus + Inhalt
const BW = 1700, CW = W - BW;
function step(m, parts) {
  const children = [];
  if (parts.say) children.push(p([t("Ansage: ", { bold: true, size: 20, color: C.hv }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((d) => children.push(p(Array.isArray(d) ? d : S(d), { after: 50 })));
  if (parts.board) { children.push(...[].concat(parts.board)); children.push(p(t(""), { after: 20 })); }
  if (parts.sol) children.push(p([t("Lösung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  if (children.length && !parts.sol && !parts.board) children[children.length - 1] = children[children.length - 1];
  return table([BW, CW], [row([labelCell(m, BW), cell(children, { w: CW, borders: allBorders(solid("D5DDEE", 6)), m: 100, ml: 160, valign: VerticalAlign.TOP })])]);
}
const steps = (arr) => arr.flatMap((s) => [step(s[0], s[1]), p(t(""), { after: 40 })]);

// Tafel-Kasten (innerhalb eines Schritts)
const TW = CW - 340;
function board(title, children) {
  return table([TW], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 100, ml: 160 })]),
  ]);
}
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
// Merksatz steht auf der Folie (nicht an der Tafel)
function folie(title, children) {
  return table([TW], [
    row([cell(p(t("FOLIE · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: allBorders(solid(NAVY, 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: LIGHT, borders: allBorders(solid(NAVY, 12)), m: 100, ml: 160 })]),
  ]);
}
const merkBoard = (nr, title, runs) => folie(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  p(runs, { after: 0 }),
]);
// Übung an der Tafel – immer im gleichen Format
function auftrag(title, todo, lines, meta) {
  return board("rechts · " + title, [
    ...todo.map((x, i) => p([t(`${i + 1}. `, { bold: true, color: NAVY, size: 21 }), t(x, { bold: true, color: NAVY, size: 21 })], { after: 30 })),
    p(t(""), { after: 30 }),
    ...lines,
    p(t(""), { after: 30 }),
    p([t("Zeit: ", { bold: true, size: 19, color: "3D4A3F" }), t(meta, { size: 19, color: "3D4A3F" }), t("     Fertig? Stift hinlegen und noch einmal durchlesen.", { size: 19, italics: true, color: "3D4A3F" })], { after: 0 }),
  ]);
}
function phase(nr, title, min) {
  return p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
}


// Vokabelliste Deutsch – Polnisch: Unterricht, Arbeitsaufträge, Verben/Zeitformen, Balladen
const PWD = 11906 - 1100;
const DW = 3050, PLW = 3350, BW2 = PWD - DW - PLW;
const bd = allBorders(solid("C9D3EA", 6));
const f = 18;
const COL = { r: C.hv, y: "C98A1E", b: C.inf, g: C.mv };
const rich = (txt, base) => txt.split(/(\{[rybg]:[^}]*\})/).filter(Boolean).map((x) => {
  const m = x.match(/^\{([rybg]):([^}]*)\}$/);
  return m ? t(m[2], Object.assign({}, base, { bold: true, color: COL[m[1]] })) : t(x, base);
});
const de = (word, gram) => cell([p(rich(word, { bold: true, size: f + 1, color: NAVY }), { after: 0 }), ...(gram ? [p(t(gram, { size: f - 3, color: MUTED }), { after: 0 })] : [])], { w: DW, borders: bd, m: 25, ml: 80 });
const pl = (txt) => cell(p(txt.split("**").map((x, i) => t(x, { size: f, bold: i % 2 === 1 })), { after: 0 }), { w: PLW, borders: bd, m: 25, ml: 80, fill: "FBFCFF" });
const bsp = (txt) => cell(p(rich(txt, { size: f - 1, color: "33405F" }), { after: 0 }), { w: BW2, borders: bd, m: 25, ml: 80 });
const sec = (a, b, col) => row([cell(p([t(a, { bold: true, size: f + 2, color: "FFFFFF", font: "Cambria" }), t("   " + b, { size: f, color: "FFFFFF", italics: true })], { after: 0 }), { w: PWD, fill: col, borders: bd, m: 40, ml: 80, span: 3 })]);
const r = (w, g, x, b = "") => row([de(w, g), pl(x), bsp(b)], 0);
const head = () => row([hdr("Deutsch", DW), hdr("Polski", PLW), hdr("Beispiel", BW2)]);

const rows = [
  head(),
  sec("1  Im Unterricht", "na lekcji", NAVY),
  r("das Heft", "die Hefte", "**zeszyt**", "Schreibt ins Heft."),
  r("die Tafel", "die Tafeln", "**tablica**", "Wer drankommt, schreibt an die Tafel."),
  r("das Lehrbuch", "die Lehrbücher", "**podręcznik**", "Lehrbuch S. 136"),
  r("die Überschrift", "die Überschriften", "**tytuł, nagłówek**", "Schreibt die Überschrift ab."),
  r("das Lernziel · der Ablauf", "", "**cel lekcji** · **przebieg (plan) lekcji**", "Ziel: Ich kann Reime bestimmen."),
  r("der Merksatz", "die Merksätze", "**reguła do zapamiętania**", ""),
  r("der Schnipsel · einkleben", "die Schnipsel", "**karteczka** · **wkleić**", "Klebt den Schnipsel ins Heft."),
  r("die Aufgabe · die Übung", "die Aufgaben · die Übungen", "**zadanie** · **ćwiczenie**", "Ü1, LB-Aufgabe 1 a"),
  r("die Lösung", "die Lösungen", "**rozwiązanie**", ""),
  r("der Fehler", "die Fehler", "**błąd**", "Sucht einen Fehler."),
  r("die Leistungskontrolle", "die Leistungskontrollen", "**sprawdzian**", ""),
  r("die Note · die Mitarbeit", "", "**ocena** · **aktywność na lekcji**", "mündliche Note"),
  r("sich melden", "", "**zgłaszać się (podnieść rękę)**", "Ich melde mich."),
  r("allein · zu zweit · im Tandem", "", "**samodzielnie** · **w parach** · **w parze (czytanie)**", "Jetzt arbeitet ihr allein."),

  sec("2  Arbeitsaufträge", "polecenia w zadaniach", "5B6B8C"),
  r("lesen", "Lies …!", "**czytać** · Przeczytaj …!", "Lies Strophe 1."),
  r("abschreiben", "Schreibe ab!", "**przepisać**", "Schreibt die Überschrift ab."),
  r("nennen", "Nenne …!", "**wymienić, podać** (krótko)", "Nenne die Figuren."),
  r("beschreiben", "Beschreibe …!", "**opisać**", "Beschreibe, wie Paddy aussieht."),
  r("erklären", "Erkläre …!", "**wyjaśnić**", "Erkläre die List von Schaya."),
  r("begründen", "Begründe …!", "**uzasadnić** (dlaczego? „…, weil …“)", "Ich lese genau, weil ich etwas erklären soll."),
  r("belegen", "Belege …!", "**udowodnić** (podać strofę lub cytat)", "Belege mit der Strophe: (Str. 4)"),
  r("bestimmen", "Bestimme …!", "**określić**", "Bestimme Person, Numerus und Tempus."),
  r("bilden", "Bilde …!", "**utworzyć**", "Bilde das Perfekt von „laufen“."),
  r("einsetzen", "Setze … ein!", "**wstawić**", "Setze das richtige Hilfsverb ein."),
  r("ergänzen · eintragen", "Ergänze! · Trage ein!", "**uzupełnić** · **wpisać**", "Ergänze die Tabelle."),
  r("markieren", "Markiere …!", "**zaznaczyć (kolorem)**", "Markiere das Hilfsverb rot."),
  r("unterstreichen", "Unterstreiche …!", "**podkreślić**", "Unterstreiche das Partizip II."),
  r("ankreuzen", "Kreuze an!", "**zaznaczyć krzyżykiem**", ""),
  r("ordnen · zuordnen", "Ordne … zu!", "**uporządkować** · **przyporządkować**", "Ordne die Sätze dem Zeitstrahl zu."),
  r("die Aussage prüfen", "die Aussagen", "**sprawdzić stwierdzenie** – czy jest prawdziwe?", "Prüfe die Aussage."),
  r("richtig · falsch", "", "**prawdziwy, poprawny** · **fałszywy, błędny**", "☐ richtig  ☐ falsch"),
  r("das Zitat", "die Zitate", "**cytat** – dokładne słowa z tekstu w cudzysłowie", "„fast wär ertrunken der gute Held“"),
  r("verbinden", "Verbinde …!", "**połączyć**", "Verbinde die Sätze mit „weil“."),
  r("einkreisen", "Kreise … ein!", "**zakreślić (kółkiem)**", "Kreise die Subjunktion ein."),
  r("vergleichen", "Vergleiche …!", "**porównać**", "Vergleicht mit eurem Nachbarn."),
  r("prüfen · verbessern", "Prüfe! · Verbessere!", "**sprawdzić** · **poprawić**", "Verbessere mit Grün."),
  r("notieren", "Notiere …!", "**zanotować**", "Notiere die Strophe."),
  r("nacherzählen", "Erzähle nach!", "**opowiedzieć własnymi słowami**", "Erzählt die Ballade nach."),
  r("zuerst · dann · schließlich", "daraufhin · am Ende", "**najpierw** · **potem** · **w końcu** (do łączenia zdań)", "Zuerst …, dann …, schließlich …"),
  r("Fertig? Stift hinlegen.", "", "**Gotowe? Odłóż długopis.**", ""),

  sec("3  Verben und Zeitformen", "czasowniki i czasy", C.hv),
  r("das Verb", "die Verben", "**czasownik**", "trainieren, laufen, heben"),
  r("der Infinitiv", "die Infinitive", "**bezokolicznik**", "{b:gewinnen}, {b:laufen}, {b:heben}"),
  r("die finite Form (Personalform)", "finit", "**forma osobowa** (odmieniona przez osoby)", "Ich trainiere. Er läuft."),
  r("infinit", "", "**nieodmieniony** (bezokolicznik, Partizip II)", "{b:gewinnen}, {y:gewonnen}"),
  r("konjugieren", "die Konjugation", "**odmieniać przez osoby** · **odmiana**", "ich laufe, du läufst, er läuft"),
  r("der Stamm · die Endung", "die Endungen", "**temat** · **końcówka**", "lauf-en: Stamm „lauf“, Endung „-en“"),
  r("die Person", "1., 2., 3. Person", "**osoba**", "ich (1.), du (2.), er/sie/es (3.)"),
  r("der Numerus: Singular · Plural", "", "**liczba: pojedyncza** · **mnoga**", "ich laufe (Sg.) · wir laufen (Pl.)"),
  r("das Personalpronomen", "die Personalpronomen", "**zaimek osobowy**", "ich, du, er, sie, es, wir, ihr, sie"),
  r("das Tempus · die Zeitform", "die Tempora · die Zeitformen", "**czas gramatyczny**", ""),
  r("der Zeitstrahl", "", "**oś czasu**", ""),
  r("die Vorvergangenheit · die Vergangenheit", "", "**zaprzeszłość** · **przeszłość**", ""),
  r("die Gegenwart · die Zukunft", "", "**teraźniejszość** · **przyszłość**", ""),
  r("das Präsens", "", "**czas teraźniejszy**", "Ich trainiere heute."),
  r("das Präteritum", "", "**czas przeszły prosty** – w tekstach pisanych i opowiadaniach", "Ich trainierte. Er lief."),
  r("das Perfekt", "", "**czas przeszły złożony** – gdy mówimy o przeszłości", "Ich {r:habe} {y:gewonnen}. Ich {r:bin} {y:gelaufen}."),
  r("das Plusquamperfekt", "", "**czas zaprzeszły** – coś stało się wcześniej niż inna rzecz w przeszłości", "Ich {r:hatte} vorher {y:trainiert}."),
  r("das Futur I", "", "**czas przyszły**", "Ich {r:werde} {b:gewinnen}."),
  r("das Futur II", "", "**czas przyszły dokonany** – coś będzie już skończone", "Ich {r:werde} {y:gewonnen} {r:haben}."),
  r("das {r:Hilfsverb}", "haben, sein, werden", "**czasownik posiłkowy**", "Sie {r:hat} {y:gewonnen}."),
  r("das Vollverb", "die Vollverben", "**czasownik główny (samodzielny)**", "Ich hebe 80 kg."),
  r("das {g:Modalverb}", "können, müssen, dürfen, sollen, wollen, mögen", "**czasownik modalny**", "Ich {g:muss} {b:trainieren}."),
  r("das {y:Partizip II}", "", "**imiesłów czasu przeszłego** („trzecia forma“)", "{y:gewonnen}, {y:trainiert}, {y:gelaufen}"),
  r("regelmäßig · unregelmäßig", "", "**regularny** · **nieregularny**", "trainieren – trainierte · laufen – lief – gelaufen"),
  r("der Satz · das Subjekt", "die Sätze", "**zdanie** · **podmiot**", "Der Trainer (Subjekt) lobt mich."),

  sec("4  Satzbau", "budowa zdania", "7A5BA6"),
  r("der Hauptsatz · der Nebensatz", "die Hauptsätze · die Nebensätze", "**zdanie główne** · **zdanie podrzędne**", "Der Schotte flieht, weil er Angst hat."),
  r("die Subjunktion", "die Subjunktionen", "**spójnik podrzędny** – zaczyna zdanie podrzędne", "dass, als, weil, wenn, ob, obwohl, während, nachdem"),
  r("dass · als", "", "**że** · **kiedy** (raz, w przeszłości)", "…, als er die Nase ragen sah."),
  r("weil · wenn", "", "**ponieważ, bo** · **jeśli, kiedy** (zawsze)", "…, wenn wir als Team spielen."),
  r("ob · obwohl", "", "**czy** · **chociaż**", "Der Trainer fragt, ob …"),
  r("während · nachdem", "", "**podczas gdy** · **po tym, jak**", ""),
  r("die Verbstellung", "", "**szyk (miejsce) czasownika** – w zdaniu podrzędnym forma osobowa stoi na końcu", "…, weil ich 80 kg gestoßen {r:habe}."),
  r("das Komma", "die Kommas", "**przecinek** – stoi przed spójnikiem", "Ich freue mich, weil …"),

  r("die Ballade", "die Balladen", "**ballada**", "„Paddy Fingal“ von Wilhelm Brandes"),
  r("die Lesetechnik", "überfliegend · orientierend · genau", "**technika czytania**: pobieżnie · wybiórczo · dokładnie", "Wie heißt Paddys Frau? → orientierend"),
  r("die Strophe · der Vers", "die Strophen · die Verse", "**strofa** · **wers**", "11 Strophen mit je 5 Versen"),
  r("der Reim · sich reimen", "die Reime", "**rym** · **rymować się**", "Schreck – Versteck"),
  r("das Reimschema", "", "**układ rymów**", "a a b b x"),
  r("der Paarreim · der Kreuzreim · der umarmende Reim", "aabb · abab · abba", "**rym parzysty** · **krzyżowy** · **okalający**", ""),
  r("der Kehrreim (Refrain)", "", "**refren**", "„… Paddy Fingal!“"),
  r("die Lyrik · die Epik · die Dramatik", "", "**liryka** · **epika** · **dramat**", "Die Ballade ist ein „Ur-Ei“."),
  r("der Erzähler · die Figur", "die Figuren", "**narrator** · **postać**", "Paddy, Schaya, der Schotte"),
  r("die Handlung · der Höhepunkt", "", "**akcja (fabuła)** · **punkt kulminacyjny**", "Der Schotte sieht die Nase."),
  r("die wörtliche Rede", "", "**mowa niezależna**", "„O Schaya, birg mich im Versteck!“"),
  r("die Spannung · der Konflikt", "", "**napięcie** · **konflikt**", "Paddy gegen den Schotten"),
  r("die Zwischenüberschrift · gliedern", "", "**śródtytuł** · **podzielić tekst na części**", "Str. 10: Flucht"),
  sec("6  Meine neuen Wörter", "moje nowe słowa – dopisuj sama", MUTED),
  ...Array.from({ length: 10 }, () => row([cell(p(t("")), { w: DW, borders: bd }), cell(p(t("")), { w: PLW, borders: bd, fill: "FBFCFF" }), cell(p(t("")), { w: BW2, borders: bd })], 520)),
];

const doc = [
  p([t("Vokabelliste Deutsch 7A", { font: "Cambria", bold: true, size: 30, color: NAVY }), t("   Słowniczek: niemiecki – polski", { font: "Cambria", size: 24, color: MUTED, italics: true })], { after: 40 }),
  p([t("Tu są ważne słowa z naszych lekcji: na lekcji, polecenia, czasowniki i czasy, budowa zdania, ballady. Po lewej słowo niemieckie (z rodzajnikiem i liczbą mnogą), w środku po polsku, po prawej przykład. ", { size: f }), t("Wskazówka: zakryj kolumnę „Polski“ i sprawdź, czy pamiętasz znaczenie. Ucz się codziennie kilku słów.", { size: f, italics: true })], { after: 40 }),
  p([t("Kolory: ", { size: f, bold: true }), t("czasownik posiłkowy", { size: f, bold: true, color: COL.r }), t(" · ", { size: f }), t("Partizip II", { size: f, bold: true, color: COL.y }), t(" · ", { size: f }), t("bezokolicznik", { size: f, bold: true, color: COL.b }), t(" · ", { size: f }), t("czasownik modalny", { size: f, bold: true, color: COL.g })], { after: 100 }),
  table([DW, PLW, BW2], rows),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 600, left: 550, right: 550, footer: 300 } } };
const styles = { default: { document: { run: { font: "Calibri", size: f } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, footers: { default: footer("Deutsch 7A · Vokabelliste Deutsch – Polnisch") }, children: doc }] })).then((buf) => fs.writeFileSync("Vokabelliste_Deutsch_Polnisch.docx", buf));
