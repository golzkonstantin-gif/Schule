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

const none = { style: BorderStyle.NONE, size: 0, color: NAVY };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };
const thin = allBorders(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: undefined, // Tinte sparen: keine Farbflächen im Druck
    borders: o.borders || thin,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: NAVY, size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font, italics: o.it }), { after: 0, align: o.align }), { w, fill: o.fill });

// simple grid table: header + rows of strings/runs
function grid(widths, head, rows, o = {}) {
  return table(widths, [
    row(head.map((h, i) => hdr(h, widths[i], o.headFill ? o.headFill[i] : NAVY)), 420),
    ...rows.map((r) => row(r.map((v, i) => tc(v, widths[i], { font: o.fonts && o.fonts[i], align: o.align && o.align[i], bold: o.bold && o.bold[i], color: o.colors && o.colors[i] })), o.h || 460)),
  ]);
}

function box(children, fill = LIGHT2, border) {
  return table([W], [row([cell(children, { w: W, fill, borders: border || allBorders(solid("B9C6E8", 8)), m: 140, ml: 180 })])]);
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
    row([cell(p(t("TAFEL · " + title, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: W, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)) })]),
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
  tafel: "kommt an die Tafel", kontrolle: "kontrolliert", partner: "arbeitet zu zweit", fertig: "gibt ab", gruppe: "arbeitet in der Gruppe",
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
    row([cell(p(t("TAFEL · " + title, { bold: true, color: NAVY, size: 16 }), { after: 0 }), { w: TW, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 100, ml: 160 })]),
  ]);
}
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
// Merksatz steht auf der Folie (nicht an der Tafel)
function folie(title, children) {
  return table([TW], [
    row([cell(p(t("FOLIE · " + title, { bold: true, color: NAVY, size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: allBorders(solid(NAVY, 12)), m: 30 })]),
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



// Tafelskript 27.11. (Einzelstunde, 45 Min.): Leseband · Kontrolle Reportage-Aussagen · Einleitung eines Sachtexts – AVIVA, mit Präsentation. Ohne Farbflächen.
const AUS = JSON.parse(fs.readFileSync(__dirname + "/aussagen_reportage.json", "utf8"));
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const Z = "Ich kann die Einleitung eines Sachtexts zu meiner Sportart schreiben: Sie nennt das Thema, weckt Interesse und gibt einen Überblick.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · 27.11. · Einzelstunde (45 Min.) · AVIVA · mit Präsentation (19 Folien)"), h1("Der erste Eindruck zählt – die Einleitung"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Letzte Stunde: Reportage „Sehen mit den Fingerkuppen“ (LB S. 183–184) gelesen, Aussagen prüfen und belegen als Übung (AB). Rückmeldung 26.11.: „Richtig ist“ hat verwirrt, einige haben das Zitat noch einmal abgeschrieben. Deshalb heute vor der Kontrolle ausdrücklich: Beleg = Zitat (abschreiben), Verbesserung = eigene Worte. Danach der Schritt vom Lesen zum Schreiben: Die Klasse schreibt selbst einen Sachtext zur eigenen Sportart – heute nur die Einleitung. Die Merkmale werden an zwei bekannten LB-Texten wiederholt bzw. erarbeitet (Vorspann S. 178, Anfang S. 180).")], { after: 0 })]));
doc.push(p([S("Ziel: ", { bold: true }), S(Z, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Eine Einleitung nennt das Thema, macht neugierig und sagt, was im Text kommt – kurz, sachlich, im Präsens. Ich habe so eine Einleitung für einen Text über meine Sportart geschrieben.“")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Ziel: ", { bold: true, size: 20, color: NAVY }), t(Z, { size: 20, italics: true })], { after: 40 }),
  ...["Leseband", "Kontrolle: Reportage – Aussagen prüfen", "Wie beginnt ein Sachtext? Merkmale einer Einleitung", "Schreiben: Einleitung zu meiner Sportart", "* Partnercheck mit der Checkliste", "Zielcheck"].map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["2′", "A", "1–2", "Ziel und Ablauf", "hören zu, lesen Ziel und Ablauf"],
  ["8′", "V", "3", "Leseband", "lesen mit der angesagten Methode"],
  ["10′", "V", "4–14", "Kontrolle Reportage – Aussagen prüfen", "vergleichen ihr AB, verbessern mit Grün"],
  ["7′", "I", "15–16", "Wie beginnt ein Sachtext? → Merksatz", "untersuchen zwei Einleitungen zu zweit, gleichen mit dem Merksatz ab"],
  ["13′", "V", "17–18", "Einleitung zur eigenen Sportart schreiben", "schreiben allein auf dem Schnipsel; * Partnercheck"],
  ["5′", "A", "19", "Zielcheck: prüfen und vorlesen", "prüfen die eigene und eine vorgelesene Einleitung mit drei Fragen"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("45 Min. Wenn es eng wird: Aussage * 10 in der Kontrolle überspringen (Lösung zeigen), Partnercheck zu Beginn der nächsten Stunde. Die Einleitung wird nächste Stunde weitergeschrieben (Hauptteil).")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["AB Reportage", "mit Grün verbessert: Urteil, Beleg (Zeile + Zitat), bei falschen Aussagen Verbesserung in eigenen Worten"],
  ["Überschrift", "Einen Sachtext schreiben: die Einleitung · Datum"],
  ["Merkmale", "Notizen zu zweit (drei Merkmale mit Beispiel aus LB S. 178 / S. 180)"],
  ["Schnipsel", "Merksatz, Satzanfänge, eigene Überschrift und Einleitung (3–5 Sätze), Checkliste – eingeklebt"],
].map(([a, b]) => row([tc(a, 2300, { bold: true }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation „Stunde_Einleitung_Sachtext.pptx“ · Lehrbuch S. 178, 180, 183–184 · AB „Aussagen prüfen“ (ausgefüllt von letzter Stunde) · Schnipsel „Schnipsel_Schreibhilfe_Einleitung_Sachtext“ (2 pro A4 quer, halbieren) · Leseband-Texte · Handreichung Polnisch für S-PL („Handreichung_Polnisch_27-11_Einleitung“, zweisprachige Fassung des Schnipsels – ersetzt bei ihr den Schnipsel).")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 2));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Zuerst Leseband, dann kontrollieren wir die Übung zur Reportage. Danach werdet ihr selbst Autoren: Ihr schreibt die Einleitung für einen Sachtext über eure Sportart.", do: ["Folie 1–2. Auf das Ziel an der Seitentafel zeigen."] }]]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Leseband", 8));
doc.push(...steps([["partner", { say: "Nehmt euren Leseband-Text. Heute lest ihr mit der Methode … (ansagen). Nach jedem Abschnitt fasst ihr in einem Satz zusammen, worum es ging. Stopp-Signal: Stifte und Texte liegen, Blick nach vorn.", do: ["Folie 3. Methode nach Methodenkarte wählen, z. B. Tandemlesen oder Ich-Du-Wir-Würfeln; die Methode an der Tafel benennen.", "Der Ein-Satz-Rückblick übt schon das Zusammenfassen für die KA 3."] }]]));

doc.push(avivaPhase("V", "Kontrolle: Reportage – Aussagen prüfen", 10));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Der Beleg ist die Zeile und das Zitat – das schreibt ihr aus dem Text ab. Die Verbesserung schreibt ihr selbst, in eigenen Worten. Das Zitat schreibt ihr dort nicht noch einmal ab.", do: ["Folie 4: Unterschied Beleg – Verbesserung am Beispiel."] }],
  ["kontrolle", { say: "Nehmt euer AB und einen grünen Stift. Wer drankommt, nennt: richtig oder falsch, Zeile und Zitat. Dann decke ich auf. Alle verbessern mit Grün.", sol: AUS.map(([a, ok, z, , r], i) => `${i + 1} ${ok ? "richtig" : "falsch"} (${z})${r ? " → " + r : ""}`).join(" · "), do: ["Folie 5–14, eine Aussage pro Folie. Erst Urteil und Beleg nennen lassen, dann aufdecken.", "Bei falschen Aussagen zwei, drei Verbesserungen vorlesen lassen: Sind es wirklich eigene Worte? Andere richtige Formulierungen gelten.", "Aussage * 10 nur, wenn Zeit ist."] }],
]));

doc.push(avivaPhase("I", "Informieren: Wie beginnt ein Sachtext?", 7));
doc.push(...steps([
  ["partner", { say: "Jetzt schreibt ihr die Überschrift ab: Einen Sachtext schreiben: die Einleitung. Dann arbeitet ihr zu zweit. Lest zwei Einleitungen: LB S. 178, der kursive Text bei den Rauchzeichen, und LB S. 180, Zeile 1 bis 7. Notiert: Was macht der Anfang? Findet mindestens drei Dinge mit Beispiel. Fertig? Stift hinlegen.", sol: "Thema sofort klar (Überschrift + erster Satz) · macht neugierig (überraschende Tatsache, Frage, Ansprache „euch“) · kündigt an, was kommt („erklären wir euch hier …“) · kurz, sachlich, Präsens, noch keine Einzelheiten", do: ["Folie 15. Hilfsfragen in den Notizen: Woran erkennt man sofort das Thema? Was macht neugierig? Was erfährt man über den Rest? Wie lang ist der Anfang?", "Beide Texte sind bekannt (18.11. und Gebärdensprache-Stunde) – es geht nur um den Anfang."] }],
  ["gemeinsam", { say: "Wer drankommt, nennt ein Merkmal und das Beispiel. Dann vergleichen wir mit dem Merksatz. Ich teile den Schnipsel aus – klebt ihn ein.", do: ["Folie 16: Kernsatz im roten Rahmen, vier Merkmale, Wofür? und Test. Der Merksatz steht oben auf dem Schnipsel (nicht abschreiben)."] }],
]));

doc.push(avivaPhase("V", "Verarbeiten: Einleitung zur eigenen Sportart schreiben", 13));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht vorne. Jeder schreibt über seine eigene Sportart. Erst die Überschrift, dann drei bis fünf Sätze: Thema nennen, Interesse wecken, Überblick geben. Nutzt die Satzanfänge auf dem Schnipsel. Schreibt im Präsens. Fertig? Stift hinlegen und noch einmal durchlesen.", do: ["Folie 17. Herumgehen: Wer nicht anfängt, bekommt den Tipp „Beginne mit einem Satzanfang aus der ersten Spalte“.", "Häufige Fehler: Einzelheiten schon in der Einleitung (Regeln, Gewichtsklassen) · Meinung („Judo ist der beste Sport“) · Erzählung im Präteritum („Letztes Jahr bin ich …“) – auf die Checkliste zeigen."] }],
  ["gemeinsam", { say: "Stifte liegen. Seht euch ein Beispiel an – Schwimmen, das macht hier niemand. Welcher Satz nennt das Thema, welcher weckt Interesse, welcher gibt den Überblick?", sol: "blau: Schwimmen ist eine der ältesten Sportarten der Welt. · rot: Wusstest du, dass …? · grün: In diesem Text erfährst du, …", do: ["Folie 18 – erst zeigen, wenn die meisten mindestens zwei Sätze haben (sonst wird nachgeahmt). Danach weiterschreiben bzw. überarbeiten."] }],
  ["partner", { say: "* Wer fertig ist, tauscht mit dem Nachbarn. Prüft mit der Checkliste: Haken oder Kreuz. Dann gebt ihr den Schnipsel zurück und überarbeitet mit Grün.", do: ["Partnercheck ist optional (*). Wenn die Zeit fehlt: zu Beginn der nächsten Stunde."] }],
]));

doc.push(avivaPhase("A", "Auswerten: Zielcheck", 5));
doc.push(...steps([["kontrolle", { say: "Zielcheck. Prüft zuerst eure eigene Einleitung mit den drei Fragen: Thema genannt? Interesse geweckt? Überblick gegeben? Setzt am Rand einen Haken oder ein Kreuz. Dann liest jemand vor, alle prüfen mit.", sol: "Pro Frage ✓/✗ mit Begründung aus dem Text („Der erste Satz nennt das Thema“).", do: ["Folie 19. Zwei bis drei Schüler lesen vor (möglichst verschiedene Sportarten). Die Klasse nennt zu jeder Frage ✓ oder ✗ mit Begründung.", "Schnipsel einkleben lassen; notieren, wem noch ein Merkmal fehlt. Ausblick: Nächste Stunde schreibt ihr den Hauptteil."] }]]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(new Paragraph({ pageBreakBefore: true, children: [t("FÜR DIE LEHRKRAFT · REFLEXION VOR UND NACH DER STUNDE", { size: 18, color: MUTED, bold: true })], spacing: { after: 20 } }), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Ziel als „Ich kann …“ und Ablauf an der Seitentafel, Partnercheck mit * markiert."));
doc.push(bul("Vorwissen aktivieren:", "Leseband (Methodenkarte) und echte Wiederholung: Kontrolle der Übung der letzten Stunde, Aussage für Aussage aufgedeckt, alle verbessern mit Grün – niemand muss an die Tafel."));
doc.push(bul("Informieren:", "Merkmale werden an zwei bekannten Einleitungen entdeckt, erst danach der Merksatz (Schnipsel, nicht abschreiben)."));
doc.push(bul("Verarbeiten:", "Erst das Werkzeug (Merksatz + Satzanfänge), dann die Anwendung im eigenen Text; das Beispiel kommt erst nach dem Start, damit nicht nachgeahmt wird."));
doc.push(bul("Auswerten:", "Zielcheck mit genau den drei Zielmerkmalen; das Ergebnis entscheidet, was die nächste Stunde wiederholt."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Ein überprüfbares Ziel; Relevanz: KA 3 (zusammenfassen) und eigene Texte über den Sport."));
doc.push(bul("Ich – Wir – Ihr:", "Merkmale gemeinsam bündeln (Wir), Beispiel Schwimmen (Ich), eigene Einleitung (Ihr)."));
doc.push(bul("Rückmeldung:", "Aufdeckfolien bei der Kontrolle, Checkliste im Partnercheck, Vorlesen im Zielcheck."));
doc.push(bul("Klarheit:", "Rückmeldung 26.11. umgesetzt: Beleg und Verbesserung werden vorab ausdrücklich unterschieden, auf den Folien heißt das Feld „Verbesserung (eigene Worte)“. Keine Zeitangaben auf dem Schnipsel."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "Satzanfänge (je vier pro Merkmal, keine fertigen Sätze); Sportbezug individuell – jeder schreibt über die eigene Sportart; Partnercheck * für Schnelle."));
doc.push(bul("", "S-PL: Handreichung als ganzes A4-Blatt (zweisprachige Fassung des Schnipsels): Ziel, Ablauf, Merksatz und Checkliste auch auf Polnisch, Satzanfänge deutsch. Sie schreibt auf Deutsch; keine Lösungen."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Schreiben die Schüler bei falschen Aussagen jetzt wirklich eigene Worte statt des Zitats?"));
doc.push(bul("", "Bleibt die Einleitung allgemein, oder rutschen Einzelheiten, Meinungen oder Erzählungen hinein?"));
doc.push(bul("", "Werden die Satzanfänge genutzt und sinnvoll ergänzt? Reicht die Zeit für 3–5 Sätze?"));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Sachtexte · 27.11.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync(__dirname + "/../Tafelskript_27-11_Einleitung_Sachtext.docx", buf));
