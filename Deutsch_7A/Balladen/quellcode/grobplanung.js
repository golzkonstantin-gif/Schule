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


// ================================================================
// Grobplanung „Leute, höret die Geschicht’ – Umgang mit Balladen“ (7A, ca. 9 Wochen, 5 Std./Woche)
// ================================================================
const { PageOrientation } = require("docx");
const LW = 16838 - 1400;
const s19 = (x, o = {}) => t(x, Object.assign({ size: 19 }, o));
const GW = [700, 2700, 3600, 5300, 1238, 1900];
const ghdr = ["Std.", "Thema", "Lernziel: „Ich kann …“", "Inhalt und Methode (alle aktiv)", "RLP", "Bewertung"];
const gb = allBorders(solid("B9C6E8", 6));
const gcell = (x, w, o = {}) => cell(p(Array.isArray(x) ? x : s19(x, o), { after: 0 }), { w, fill: o.fill, borders: gb, m: 50, ml: 80, valign: VerticalAlign.TOP });
const BEW = "C0392B";
const grow = ([nr, thema, ziel, inhalt, rlp, bew]) => row([
  gcell(String(nr), GW[0], { bold: true, color: C.hv, fill: LIGHT }),
  gcell(thema, GW[1], { bold: true, color: NAVY }),
  gcell(ziel, GW[2], { italics: true }),
  gcell(inhalt, GW[3]),
  gcell(rlp, GW[4], { color: MUTED, size: 17 }),
  gcell(bew, GW[5], { bold: true, color: BEW }),
]);
const week = (title, sub) => row([cell(p([t(title, { bold: true, color: "FFFFFF", size: 21, font: "Cambria" }), t(sub ? "   " + sub : "", { color: "DCE3F2", size: 18 })], { after: 0 }), { w: LW, fill: NAVY, borders: gb, span: 6, m: 50, ml: 100 })]);

const WEEKS = [
  ["Woche 1 · Was ist eine Ballade?", "4 Stunden · John Maynard (Fontane)", [
    [1, "Eine Ballade hören", "… beim Zuhören Notizen zu Figuren, Ort und Handlung machen.", "Hörfassung „John Maynard“ (z. B. Achim Reichel). Hörauftrag für alle: Notizraster wer – wo – was passiert – wie endet es. Zweites Hören zu zweit ergänzen, dann Abgleich.", "2.3 D · 2.11 D", ""],
    [2, "Eine Ballade gliedern", "… eine Ballade in Abschnitte gliedern und Zwischenüberschriften formulieren.", "Text lesen, in PA Abschnitte markieren und Zwischenüberschriften notieren; Vergleich an der Tafel (Streifen mit Überschriften ordnen – erst am Platz, dann Tafel).", "2.8 D", ""],
    [3, "Das „Ur-Ei“: episch, lyrisch, dramatisch", "… an einer Ballade erzählende, lyrische und dramatische Elemente nachweisen.", "Merksatz Ballade (Goethe: „Ur-Ei“). Belegstellen in John Maynard in drei Farben markieren (allein), dann in PA vergleichen.", "2.9 D", ""],
    [4, "Strophe, Vers, Reim", "… Strophe und Vers unterscheiden und Paarreim, Kreuzreim und umarmenden Reim bestimmen.", "Merksatz Reimformen; Reimschema an Strophen bestimmen (Schnipsel); Speed-Duell zu Reimformen.", "2.9 D", ""],
  ]],
  ["Woche 2 · Eine Ballade erschließen", "5 Stunden · John Maynard", [
    [5, "Unbekannte Wörter erschließen", "… unbekannte Wörter aus dem Kontext erschließen und Randnotizen anlegen.", "Wörter wie Steuermann, Bug, Qualm: erst aus dem Kontext raten, dann nachschlagen; Randnotizen-Technik einüben.", "2.8 E · 2.12", ""],
    [6, "Der Spannungsbogen", "… den Handlungsverlauf einer Ballade als Spannungskurve darstellen und begründen.", "Spannungskurve zeichnen (allein), Höhepunkt mit Textstelle belegen; Vergleich in PA. Vergleich mit dem Spielverlauf eines knappen Wettkampfs.", "2.8 E · 2.9 F", ""],
    [7, "Erzählzeit und erzählte Zeit", "… Erzählzeit und erzählte Zeit unterscheiden und ihre Wirkung erklären.", "Countdown „Noch zwanzig Minuten …“: Zeitangaben markieren, Zeitraffung/-dehnung erkennen, Wirkung auf die Spannung beschreiben.", "2.9 F", ""],
    [8, "Fakten hinter der Ballade", "… einen Sachtext orientierend und genau lesen und Informationen auf Verlässlichkeit prüfen.", "Sachtext zum Unglück der „Erie“ (1841) und zu Fontane; Lesetechniken; zwei Quellen vergleichen: Was ist belegt, was erfunden?", "2.8 D · 2.10 D/E", ""],
    [9, "Ein Held? John Maynard", "… eine Figur beschreiben und meine Einschätzung am Text begründen.", "Figurenbeschreibung mit Textbelegen; Positionslinie „Held oder Pflichterfüllung?“; Bezug Verantwortung im Team.", "2.8 D · 2.9 E", ""],
  ]],
  ["Woche 3 · Figuren und Konflikte", "5 Stunden · Die Bürgschaft (Schiller)", [
    [10, "Ein Konflikt beginnt", "… den Ausgangskonflikt einer Ballade benennen und Vermutungen zum Verlauf begründen.", "Strophen 1–7 lesen/hören; Konflikt Möros – Dionys herausarbeiten; Vermutungen schriftlich, dann Weiterlesen.", "2.8 D · 2.9 F", ""],
    [11, "Ein Lauf gegen die Zeit", "… die Hindernisse der Handlung in einer Grafik darstellen.", "Hindernisparcours als Grafik (Fluss, Räuber, Hitze); Parallele zum Ausdauerwettkampf: Was treibt Möros an?", "2.8 E · 2.9 F", ""],
    [12, "Figurenkonstellation", "… die Beziehungen der Figuren in einem Schaubild darstellen.", "Schaubild (Dionys, Möros, Freund, Philostratus) allein, dann in PA ergänzen; Beziehungspfeile beschriften.", "2.9 E", ""],
    [13, "Eine Figur verändert sich", "… die Entwicklung einer Figur beschreiben und am Text belegen.", "Dionys am Anfang und am Ende: Zitate sammeln, Wandel begründen; einfache Zitiertechnik (Strophe/Vers).", "2.9 E/F", ""],
    [14, "Sprache und Wirkung", "… sprachliche Mittel erkennen und ihre Wirkung fachsprachlich beschreiben.", "Wortwahl, Wiederholung, sprachliche Bilder in ausgewählten Strophen; Wirkung mit Formulierungshilfen beschreiben.", "2.9 E", ""],
  ]],
  ["Woche 4 · Balladen analysieren", "5 Stunden · Der Handschuh (Schiller)", [
    [15, "Der Handschuh", "… den Inhalt einer Ballade erfassen und gliedern (binnendifferenziert).", "Lesen mit gestuften Hilfen (Worterklärungen, Leitfragen); Gliederung mit Zwischenüberschriften.", "2.8 D/E", ""],
    [16, "Mutprobe oder Demütigung?", "… das Verhalten zweier Figuren bewerten und meine Deutung am Text belegen.", "Kunigunde und Delorges: Positionen sammeln, mit Zitaten belegen; Bezug Mut, Risiko und Respekt.", "2.9 E", ""],
    [17, "Wie analysiere ich eine Ballade?", "… eine Ballade nach Inhalt, Form und Sprache untersuchen.", "Checkliste Inhalt – Form – Sprache gemeinsam erarbeiten; Musteranalyse einer Strophe; Formulierungshilfen.", "2.9 D/E", ""],
    [18, "Analyse üben", "… eine Ballade mithilfe der Checkliste analysieren und überprüfen.", "Analyse allein schreiben, in PA mit Checkliste prüfen, mit Grün verbessern.", "2.9 D/E", ""],
    [19, "Leistung: Analyse", "… eine Ballade hinsichtlich ihrer Merkmale analysieren.", "Schriftliche Analyse (behandelte Ballade oder kurze neue Strophenfolge).", "2.9", "Bewertung 1: Analyse"],
  ]],
  ["Woche 5 · Aus der Sicht einer Figur schreiben", "5 Stunden · Bürgschaft / Handschuh / John Maynard", [
    [20, "Perspektive wechseln", "… Erzähler und Figur unterscheiden und Gefühle einer Figur aus dem Text erschließen.", "Erzählperspektive bestimmen; Gefühlskurve einer Figur (z. B. der Freund als Bürge) mit Textbelegen.", "2.9 E", ""],
    [21, "Textformen und Schreibplan", "… Merkmale von Tagebuch, Brief und Chatverlauf nennen und einen Schreibplan anlegen.", "Merkmale vergleichen; Figur und Textform wählen; Schreibplan (Cluster/Stichpunkte).", "2.6 D/F", ""],
    [22, "Entwurf schreiben", "… einen Text aus der Sicht einer Figur schreiben.", "Entwurf mit Satzanfängen als Formulierungshilfe (keine fertigen Sätze).", "2.6 F", ""],
    [23, "Schreibkonferenz", "… einen fremden Text nach Kriterien einschätzen und Verbesserungsvorschläge machen.", "Kriterienbogen; Schreibkonferenz in Dreiergruppen mit festen Rollen.", "2.6 D · 2.3", ""],
    [24, "Überarbeiten", "… meinen Text nach Kriterien und eigenen Schwerpunkten überarbeiten.", "Überarbeitung mit Grün; eigene Überarbeitungsschwerpunkte notieren.", "2.6 D/E", ""],
  ]],
  ["Woche 6 · Balladen in Ton und Musik", "5 Stunden · Erlkönig (Goethe)", [
    [25, "Leistung: Textüberarbeitung", "… einen Text aus der Figurenperspektive kriteriengeleitet überarbeiten.", "Überarbeitung und Abgabe (Bewertung).", "2.6", "Bewertung 2: Textüberarbeitung"],
    [26, "Wer spricht im Erlkönig?", "… die Sprecher einer Ballade unterscheiden und den Dialog als dramatisches Element erklären.", "Sprecher farbig markieren (Erzähler, Vater, Sohn, Erlkönig); mit verteilten Rollen lesen.", "2.9 D", ""],
    [27, "Wie entsteht das Unheimliche?", "… beschreiben, wie Sprache und Aufbau Spannung erzeugen.", "Steigerung von Strophe zu Strophe; Wortwahl des Erlkönigs; Wirkung belegen.", "2.9 E", ""],
    [28, "Schuberts Erlkönig", "… mediale Gestaltungsmittel einer Vertonung beschreiben.", "Hören: Galopp im Klavier, Stimmlagen der Figuren; Höraufträge für alle, Notizraster.", "2.11 D · 2.3", ""],
    [29, "Moderne Fassungen", "… zwei Fassungen vergleichen und ihre Wirkung auf mich begründen.", "Rap-/Popversion vs. Schubert; eigene Medienerfahrung austauschen.", "2.11 F · C/D", ""],
  ]],
  ["Woche 7 · Vom Text zum Bild – und zur Bühne", "5 Stunden · Der Zauberlehrling (Goethe)", [
    [30, "Der Zauberlehrling", "… Aufbau und Kehrreim einer Ballade beschreiben und ihre Wirkung erklären.", "Lesen, Kehrreim markieren; Bezug: Was passiert, wenn man eine Technik einsetzt, die man noch nicht beherrscht?", "2.9 D/E", ""],
    [31, "Ballade und Film", "… Text und Verfilmung vergleichen.", "Micky-Maus-Fassung (Fantasia): Was übernimmt der Film, was verändert er? Vergleichsraster allein, dann PA.", "2.11 D", ""],
    [32, "Ballade als Graphic Novel", "… mein Textverständnis in einer anderen Darstellungsform ausdrücken.", "Eine Strophe als Comic-Panel mit Sprechblase gestalten; Galeriegang mit Kurzbegründung.", "2.8 E", ""],
    [33, "Was macht einen guten Vortrag aus?", "… Kriterien für einen gelungenen Balladenvortrag nennen.", "Zwei Vorträge derselben Strophe vergleichen; Kriterien sammeln (Tempo, Pausen, Betonung, Lautstärke, Mimik, Gestik, Blickkontakt).", "2.7 D · 2.3", ""],
    [34, "Vortragszeichen setzen", "… einen Text mit Vortragszeichen für meinen Vortrag vorbereiten.", "Vortragszeichen an einer Strophe; Ballade für die Rezitation wählen.", "2.7 D", ""],
  ]],
  ["Woche 8 · Rezitation vorbereiten und vortragen", "5 Stunden", [
    [35, "Auswendig lernen", "… Strategien zum Auswendiglernen anwenden.", "Strophenweise lernen (Training in Etappen), Stichwortkarten, Partnerabfrage.", "2.7", ""],
    [36, "Mimik, Gestik, Stimme", "… nonverbale Mittel und Intonation gezielt einsetzen.", "Übungen zu Stimme und Haltung; eine Strophe mit unterschiedlicher Wirkung sprechen.", "2.1 · 2.7 E/F", ""],
    [37, "Üben mit Feedback", "… einen Vortrag nach Kriterien einschätzen und Feedback geben.", "PA: Vortrag – Feedbackbogen mit Notizen beim Zuhören – zweiter Durchgang.", "2.7 D · 2.3", ""],
    [38, "Generalprobe", "… meinen Vortrag im geschützten Rahmen halten und verbessern.", "Kleingruppen; Selbst- und Fremdeinschätzung.", "2.7", ""],
    [39, "Rezitation I", "… eine Ballade kriterienorientiert vortragen.", "Vorträge; Zuhörer notieren Feedback.", "2.7 · 2.3", "Bewertung 3: Rezitation"],
  ]],
  ["Woche 9 · Abschluss", "5 Stunden", [
    [40, "Rezitation II", "… eine Ballade kriterienorientiert vortragen.", "Vorträge; Zuhörer notieren Feedback.", "2.7 · 2.3", "Bewertung 3: Rezitation"],
    [41, "Rezitation III", "… eine Ballade kriterienorientiert vortragen.", "Restliche Vorträge (ggf. im geschützten Raum).", "2.7", "Bewertung 3: Rezitation"],
    [42, "Feedback und Selbsteinschätzung", "… meinen Lernstand in der Einheit einschätzen.", "Selbsteinschätzungsbogen; Rückmeldung zu den drei Bewertungen.", "2.7 D", ""],
    [43, "Meine Ballade", "… begründen, welche Ballade mich am meisten überzeugt.", "Kurze begründete Stellungnahme; Austausch in PA, dann Plenum.", "2.8 D · 2.11 C/D", ""],
    [44, "Puffer", "", "Reserve für Verschiebungen.", "", ""],
  ]],
];

const doc = [];
doc.push(kicker("Grobplanung · Deutsch · Klasse 7A · ca. 9 Wochen · 5 Stunden pro Woche"));
doc.push(h1("„Leute, höret die Geschicht’“ – Umgang mit Balladen"));
doc.push(table([3000, LW - 3000], [
  ["Balladen", "Fontane: John Maynard · Schiller: Die Bürgschaft, Der Handschuh · Goethe: Erlkönig, Der Zauberlehrling · zur Rezitation frei wählbar"],
  ["Leistungsbewertung (SchiC)", "Bewertung 1: Analyse der Merkmale (Std. 19) · Bewertung 2: Textüberarbeitung aus der Figurenperspektive (Std. 25) · Bewertung 3: Rezitation einer selbst gewählten Ballade (Std. 39–41)"],
  ["Planungsprinzipien", "Jede Stunde nach AVIVA · Lernziel als „Ich kann …“ und Ablauf an der Tafel (Streichbares mit *) · alle Schüler in jeder Phase aktiv: erst allein oder zu zweit, dann Vergleich an der Tafel · am Ende „Heute habe ich gelernt, …“ · Bezüge zum Leistungssport, wo sie sich anbieten"],
  ["Noch offen", "KA 1 (Verben-Test) ist nicht eingeplant – Termin legt die Lehrkraft fest; ggf. Puffer nutzen und Woche 7 straffen. Ferien und Feiertage sind nicht berücksichtigt."],
].map(([a, b]) => row([tc(a, 3000, { bold: true, color: NAVY, fill: LIGHT, size: 19 }), cell(p(s19(b), { after: 0 }), { w: LW - 3000, m: 50, ml: 80 })]))));
doc.push(p(t(""), { after: 120 }));
const rows = [new TableRow({ tableHeader: true, children: ghdr.map((h, i) => hdr(h, GW[i], NAVY, 18)) })];
WEEKS.forEach(([title, sub, list]) => { rows.push(week(title, sub)); list.forEach((x) => rows.push(grow(x))); });
doc.push(table(GW, rows));

const landProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 700, bottom: 600, left: 700, right: 700, header: 350, footer: 350 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
const ghead = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: LW }], children: [t("Deutsch · Klasse 7A · Balladen", { size: 17, color: MUTED }), t("\tGrobplanung", { size: 17, color: MUTED })] })] });
Packer.toBuffer(new Document({ styles, sections: [{ properties: landProps, headers: { default: ghead }, footers: { default: footer("Grobplanung") }, children: doc }] })).then((buf) => fs.writeFileSync("Grobplanung_Balladen.docx", buf));
