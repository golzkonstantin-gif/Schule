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


// Lernwortschatz Balladen: Begriffe Deutsch, Erklärungen Polnisch, Beispiele Deutsch
const GC = { ly: "2A8C8C", ep: "7A5BA6", dr: "C77C12", gr: NAVY, le: "3F7CC4", op: "5B6B8C", pf: "8A6D3B" };
const PWD = 11906 - 1200;
const DW = 2900, PLW = 4500, BW2 = PWD - DW - PLW;
const bd = allBorders(solid("C9D3EA", 6));
const f = 18;
const de = (word, gram) => cell([p([t(word, { bold: true, size: f + 1, color: NAVY })], { after: 0 }), ...(gram ? [p(t(gram, { size: f - 3, color: MUTED }), { after: 0 })] : [])], { w: DW, borders: bd, m: 30, ml: 80 });
const pl = (txt) => {
  const parts = txt.split("**");
  return cell(p(parts.map((x, i) => t(x, { size: f, bold: i % 2 === 1 })), { after: 0 }), { w: PLW, borders: bd, m: 30, ml: 80, fill: "FBFCFF" });
};
const bsp = (txt) => cell(p(t(txt, { size: f - 1, italics: true, color: "33405F" }), { after: 0 }), { w: BW2, borders: bd, m: 30, ml: 80 });
const sec = (de1, pl1, col) => row([cell(p([t(de1, { bold: true, size: f + 2, color: "FFFFFF", font: "Cambria" }), t("   " + pl1, { size: f, color: "FFFFFF", italics: true })], { after: 0 }), { w: PWD, fill: col, borders: bd, m: 40, ml: 80, span: 3 })]);
const r = (w, g, x, b = "") => row([de(w, g), pl(x), bsp(b)]);
const SEC = (de1, pl1, col) => { const c = sec(de1, pl1, col); return c; };

const rows = [
  row([hdr("Deutsch", DW), hdr("Polski – wyjaśnienie", PLW), hdr("Beispiel (Paddy Fingal)", BW2)]),
  SEC("Die Ballade – Grundbegriffe", "pojęcia podstawowe", GC.gr),
  r("die Ballade", "die Balladen", "**ballada** – wiersz, który opowiada (często dramatyczną) historię.", "„Paddy Fingal“ von Wilhelm Brandes ist eine Ballade."),
  r("die Gattung", "die Gattungen", "**rodzaj literacki**. Są trzy: liryka, epika i dramat.", ""),
  r("das „Ur-Ei“", "", "**„prajajo“** – tak Goethe nazwał balladę, bo są w niej wszystkie trzy rodzaje literackie.", "Die Ballade ist ein „Ur-Ei“."),
  r("das Merkmal", "die Merkmale", "**cecha** – coś typowego dla tekstu.", "Ein Merkmal der Lyrik ist der Reim."),
  r("der Autor / die Autorin", "", "**autor / autorka** – osoba, która napisała tekst.", "Autor: Wilhelm Brandes"),

  SEC("Lyrik – wie ein Gedicht", "liryka – jak wiersz", GC.ly),
  r("das Gedicht", "die Gedichte", "**wiersz**.", ""),
  r("die Strophe", "die Strophen", "**strofa (zwrotka)** – grupa wersów. Między strofami jest pusta linijka.", "Die Ballade hat 11 Strophen."),
  r("der Vers", "die Verse", "**wers** – jedna linijka wiersza.", "Jede Strophe hat 5 Verse."),
  r("der Reim", "die Reime · sich reimen", "**rym** · **rymować się** – słowa brzmią tak samo od ostatniej akcentowanej samogłoski.", "Schreck – Versteck (Str. 4)"),
  r("betont · der Vokal", "", "**akcentowany** · **samogłoska** (a, e, i, o, u, ä, ö, ü).", "B-a-rt: das a ist betont."),
  r("das Reimschema", "", "**układ rymów** – rymy oznaczamy literami (a, b, c …). Wers bez rymu = x.", "a a b b x"),
  r("der Paarreim", "aabb", "**rym parzysty** – rymują się wersy 1 i 2 oraz 3 i 4.", "Riesenart – Bart, groß – Wurfgeschoss"),
  r("der Kreuzreim", "abab", "**rym krzyżowy (przeplatany)** – rymują się wersy 1 i 3 oraz 2 i 4.", "leer – da – schwer – Ja"),
  r("der umarmende Reim", "abba", "**rym okalający** – rymują się wersy 1 i 4 oraz 2 i 3.", "zu – Kragen – wagen – Ruh"),
  r("der Kehrreim / der Refrain", "", "**refren** – wers, który powtarza się w każdej strofie (czasem trochę zmieniony).", "„… Paddy Fingal!“"),

  SEC("Epik – wie eine Erzählung", "epika – jak opowiadanie", GC.ep),
  r("die Erzählung", "die Erzählungen · erzählen", "**opowiadanie** · **opowiadać**.", ""),
  r("der Erzähler", "der Er-Erzähler · der Ich-Erzähler", "**narrator** – głos, który opowiada historię (w 3. albo w 1. osobie). To nie jest autor! Najważniejsza cecha epiki: jest narrator.", "„Paddy Fingal war von Riesenart …“ (Er-Erzähler)"),
  r("die Figur", "die Figuren", "**postać (bohater)** w tekście. Uwaga: nie „figura“!", "Paddy Fingal, Schaya, der Schotte"),
  r("die Handlung", "", "**akcja, fabuła** – to, co dzieje się w tekście.", ""),
  r("der Anfang · der Höhepunkt · das Ende", "", "**początek** · **punkt kulminacyjny** (moment, w którym wszystko się rozstrzyga) · **zakończenie**.", "Höhepunkt: Der Schotte sieht die Nase (Str. 8–9)."),
  r("das Präteritum", "", "**czas przeszły** (Präteritum) – w nim zwykle opowiada się historie.", "war, kroch, lief"),

  SEC("Dramatik – wie ein Theaterstück", "dramat – jak sztuka teatralna", GC.dr),
  r("das Theaterstück", "die Theaterstücke", "**sztuka teatralna** – postacie mówią same, nie ma narratora.", ""),
  r("ohne Erzähler", "die Figurenrede", "**bez narratora** – postacie mówią same, jak na scenie. Najważniejsza cecha dramatu. Uwaga: mowa niezależna (wörtliche Rede) jest też w epice – tam wprowadza ją narrator.", "Str. 7: cała strofa to słowa Schai"),
  r("der Dialog", "die Dialoge", "**dialog, rozmowa** między postaciami.", ""),
  r("die Spannung", "spannend", "**napięcie** · **trzymający w napięciu** – chcemy wiedzieć, co będzie dalej.", "Findet der Schotte Paddy?"),
  r("der Konflikt", "die Konflikte", "**konflikt** – postacie są przeciwko sobie.", "Paddy gegen den Schotten"),

  SEC("Lesen und mit Texten arbeiten", "czytanie i praca z tekstem", GC.le),
  r("die Lesetechnik", "die Lesetechniken", "**technika czytania**.", ""),
  r("überfliegend lesen", "", "**czytać pobieżnie** – szybko, żeby poznać ogólny sens.", "1 Minute: Worum geht es?"),
  r("orientierend lesen", "", "**czytać wybiórczo** – szukać konkretnej informacji.", "Wie heißt Paddys Frau?"),
  r("genau lesen", "", "**czytać dokładnie** – słowo po słowie, żeby zrozumieć szczegóły.", "Welche List wendet Schaya an?"),
  r("die Worterklärung", "die Worterklärungen", "**objaśnienie słowa** – pod tekstem w podręczniku.", "die Furt (Fußnote 7)"),
  r("der Abschnitt · gliedern", "die Abschnitte", "**fragment, część tekstu** · **podzielić tekst na części**.", ""),
  r("die Zwischenüberschrift", "die Zwischenüberschriften", "**śródtytuł** – krótki tytuł jednej części tekstu.", "Str. 10: Flucht"),
  r("nacherzählen", "", "**opowiedzieć własnymi słowami** – po kolei, niczego nie wymyślać.", ""),
  r("der Textbeleg · belegen", "die Textbelege", "**dowód z tekstu** · **udowodnić** – podać strofę albo cytat.", "Paarreim (Str. 1)"),
  r("analysieren · die Analyse", "", "**analizować** · **analiza** – badać tekst i opisywać jego cechy.", "Bewertung 1"),
  r("die Perspektive einer Figur", "", "**perspektywa postaci** – pisać tak, jakby się było tą postacią (np. pamiętnik).", "Bewertung 2"),
  r("vortragen · die Rezitation", "", "**recytować** · **recytacja** – mówić wiersz na głos, z pamięci.", "Bewertung 3"),

  SEC("Arbeitsaufträge", "polecenia w zadaniach", GC.op),
  r("Lies … · Nenne …", "", "**Przeczytaj …** · **Wymień, podaj** (krótko).", ""),
  r("Beschreibe … · Erkläre …", "", "**Opisz …** · **Wyjaśnij …**", ""),
  r("Begründe …", "", "**Uzasadnij …** – napisz, dlaczego („…, weil …“).", ""),
  r("Belege mit der Strophe.", "", "**Udowodnij, podając numer strofy.**", "(Str. 4)"),
  r("Markiere … · Unterstreiche …", "", "**Zaznacz …** · **Podkreśl …**", ""),
  r("Vergleiche … · Ergänze …", "", "**Porównaj …** · **Uzupełnij …**", ""),
  r("Trage ein. · Notiere …", "", "**Wpisz.** · **Zanotuj …**", ""),
  r("Arbeitet zu zweit / im Tandem.", "", "**Pracujcie w parach.**", ""),
  r("Verbessere mit Grün.", "", "**Popraw na zielono.**", ""),
  r("Fertig? Stift hinlegen.", "", "**Gotowe? Odłóż długopis.**", ""),

  SEC("Wörter aus „Paddy Fingal“", "słowa z ballady", GC.pf),
  r("der Riese · der Hüne", "die Riesen · die Hünen", "**olbrzym** · **bardzo duży, silny mężczyzna**.", ""),
  r("die Elle", "die Ellen", "**łokieć** – dawna miara długości.", "sechs Ellen"),
  r("der Sund · die Furt", "", "**cieśnina** · **bród** (płytkie miejsce w rzece).", ""),
  r("die List · das Versteck", "", "**podstęp** · **kryjówka** (sich verstecken = **chować się**).", ""),
  r("fliehen", "floh, ist geflohen", "**uciekać**.", "Der Schotte flieht."),
  r("der Schotte · Schottland", "", "**Szkot** · **Szkocja**.", ""),
];

const doc = [
  p([t("Balladen – Wörter für den Unterricht", { font: "Cambria", bold: true, size: 30, color: NAVY }), t("   Ballady – słowniczek", { font: "Cambria", size: 24, color: MUTED, italics: true })], { after: 40 }),
  p([t("Ten słowniczek pomoże Ci na lekcjach niemieckiego o balladach. Po lewej jest niemieckie słowo (z rodzajnikiem i liczbą mnogą), w środku wyjaśnienie po polsku, po prawej przykład z ballady „Paddy Fingal“. ", { size: f }), t("Ucz się niemieckich słów – będą na lekcjach i na sprawdzianach. ", { size: f, bold: true }), t("Wskazówka: zakryj środkową kolumnę i sprawdź, czy pamiętasz znaczenie.", { size: f, italics: true })], { after: 100 }),
  table([DW, PLW, BW2], rows),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 600, left: 600, right: 600, footer: 300 } } };
const styles = { default: { document: { run: { font: "Calibri", size: f } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, footers: { default: footer("Deutsch 7A · Balladen · Lernwortschatz Deutsch – Polnisch") }, children: doc }] })).then((buf) => fs.writeFileSync("Lernwortschatz_Balladen_Deutsch_Polnisch.docx", buf));
