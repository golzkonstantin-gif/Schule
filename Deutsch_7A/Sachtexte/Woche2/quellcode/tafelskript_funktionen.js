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



// Tafelskript Einzelstunde: Funktionen von Sachtexten (LB S. 180–182) + Sprachtraining Imperativ – AVIVA, mit Präsentation. Ohne Farbflächen.
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const ZIEL = "Ich kann erkennen, welche Absicht ein Sachtext verfolgt, und das mit Zeilenangabe belegen. Ich kann den Imperativ bilden.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · AVIVA · mit Präsentation (19 Folien)"), h1("Was will der Text von mir? – Funktionen von Sachtexten"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Letzte Stunde: Sachtext „Wie Gebärdensprache funktioniert“ (LB S. 180–181) gelesen, über Kommunikation gesprochen, Namen mit dem Fingeralphabet buchstabiert. Heute: Wiederholung des Inhalts über Aussagen-Prüfen (bewährtes Format), dann LB 4 a–d gemeinsam – erst Einzelarbeit, dann Unterrichtsgespräch mit Aufdeck-Folien (Entscheidung der Lehrkraft). Der Infokasten „Funktionen von Sachtexten“ wird als Merksatz-Schnipsel gesichert. Sprachtraining: Imperativ, entdeckt an Absatz 5 und geübt an Trainer-Anweisungen. LB 6 übersetzen die Schüler mit dem Fingeralphabet. Zeilenzählung wie im LB (Z. 1–98).")], { after: 0 })]));
doc.push(p([S("Lernziel: ", { bold: true }), S(ZIEL, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Ein Sachtext kann informieren, anleiten, aufklären, auffordern oder überzeugen. Im Gebärdensprache-Text informieren die ersten Absätze, der letzte fordert auf – das erkennt man an ‚du‘ und am Imperativ, z. B. ‚Nimm!‘.“")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Ziel: ", { bold: true, size: 20, color: NAVY }), t(ZIEL, { size: 20, italics: true })], { after: 40 }),
  ...["Wiederholung: Aussagen prüfen · LB 6", "LB 4a Zwischenüberschriften · Merksatz Funktionen", "LB 4b/c Absichten mit Zeile belegen", "Sprachtraining: Imperativ", "Zielcheck  (* LB 5)"].map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["2′", "A", "1–2", "Ziel und Ablauf", "lesen Ziel und Ablauf"],
  ["7′", "V", "3–5", "Wiederholung: Aussagen prüfen · LB 6 Gebärden übersetzen", "prüfen allein mit Zeile, verbessern mit Grün, übersetzen mit dem Fingeralphabet"],
  ["10′", "I", "6–8", "LB 4a (EA → UG), Merksatz Funktionen", "ordnen allein zu, vergleichen im UG, kleben Schnipsel ein"],
  ["12′", "V", "9–11", "LB 4b/c (EA → UG), 4d mündlich", "Tabelle Absatz | Absicht | Beleg, vergleichen mit Aufdeck-Folien"],
  ["9′", "I/V", "12–16", "Sprachtraining Imperativ: entdecken, Merksatz, Ü", "entdecken in Absatz 5, bilden du/ihr, verbessern mit Grün"],
  ["5′", "A", "17–18", "Zielcheck", "lösen allein, verbessern mit Grün"],
  ["*", "", "19", "* LB 5 Internettext", "nur wenn Zeit bleibt"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("45 Min. ohne Puffer. Kürzen: * Aussage 4, LB 6 nur kurz, LB 5 in die nächste Stunde.")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Funktionen von Sachtexten · Datum"],
  ["Wiederholung", "Aussagen 1–3 (*4): r/f, Zeile, Verbesserung · LB 6: Lösungswort"],
  ["LB 4", "4a Absatz 1–5 mit Zwischenüberschrift · Merksatz-Schnipsel · 4b/c Tabelle Absatz | Absicht | Beleg (Z.)"],
  ["Sprachtraining", "Ü Imperativ a–d (du/ihr) · Zielcheck"],
].map(([a, b]) => row([tc(a, 2300, { bold: true }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation `Funktionen_Sachtexte_Imperativ.pptx` · LB S. 180–182 · Fingeralphabet-Blatt (letzte Stunde) · Schnipsel `Merksatz_Funktionen_Imperativ_Schnipsel` (2 pro A4), für S-PL die zweisprachige Fassung · Handreichung Polnisch für S-PL (`Handreichung_Polnisch_Funktionen_Sachtexte`).")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 2));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Letzte Stunde habt ihr den Text über Gebärdensprache gelesen und euren Namen gebärdet. Heute fragen wir: Was will der Text eigentlich von uns? Am Ende prüft ihr im Zielcheck, ob ihr das Ziel erreicht habt.", do: ["Folie 1–2."] }]]));

doc.push(avivaPhase("V", "Wiederholung: Aussagen prüfen · LB 6", 7));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Prüft jede Aussage: richtig oder falsch? Schreibt die Zeile dazu und verbessert falsche Aussagen. Drei Minuten.", sol: "1 richtig (Z. 8–10) · 2 falsch (Z. 16–19) – erst seit 2002 · 3 falsch (Z. 23–27) – viele Gebärdensprachen · *4 falsch (Z. 28–32) – die dominante Hand", do: ["Folie 3."] }],
  ["kontrolle", { say: "Wer drankommt, nennt Urteil und Zeile. Dann decke ich auf. Alle verbessern mit Grün.", do: ["Folie 4."] }],
  ["luecken", { say: "LB Seite 182, Aufgabe 6: Was sagen die Gebärden? Übersetzt mit dem Fingeralphabet. Wer es hat, gebärdet es leise zurück.", do: ["Folie 5. Lösungswort im LB prüfen."] }],
]));

doc.push(avivaPhase("I", "LB 4a Zwischenüberschriften · Merksatz Funktionen", 10));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Ordnet den Absätzen 1 bis 5 die Zwischenüberschriften aus Aufgabe 4a zu. Drei Minuten.", sol: "1 Was ist Gebärdensprache? · 2 Geschichte der Gebärdensprache · 3 Wie funktioniert die Deutsche Gebärdensprache? · 4 Gebärdensprache: Das Alphabet · 5 Wie unterhalte ich mich mit gehörlosen Menschen?", do: ["Folie 6."] }],
  ["gemeinsam", { say: "Wer drankommt, nennt Absatz, Überschrift und ein Stichwort aus dem Absatz als Begründung.", do: ["Folie 7 aufdecken. Mit Grün verbessern."] }],
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Ein Sachtext will etwas von uns. Was genau, steht im Infokasten auf Seite 182. Ich teile den Schnipsel aus – klebt ihn ein.", do: ["Folie 8: Merksatz Funktionen (informieren, unterrichten/anleiten, aufklären, auffordern, überzeugen), Test, Wofür?"] }],
]));

doc.push(avivaPhase("V", "LB 4b/c Absichten mit Zeile belegen", 12));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Zeichnet die Tabelle: Absatz, Absicht, Beleg mit Zeile. Bestimmt für jeden Absatz die Hauptabsicht und schreibt eine passende Stelle mit Zeile dazu. Sechs Minuten.", sol: "1 informieren (Z. 8–9) · 2 informieren über Geschichte (Z. 18–19) · 3 unterrichten/anleiten (Z. 28–29) · 4 informieren über das Fingeralphabet (Z. 57–61) · 5 aufklären + auffordern (Z. 80–81 „achte“, Z. 86–88, Z. 97–98 „nimm“)", do: ["Folie 9."] }],
  ["gemeinsam", { say: "Wer drankommt, nennt Absicht und Beleg. Dann decke ich auf. Alle verbessern mit Grün.", do: ["Folie 10–11. Bei Absatz 5: Woran erkennt ihr das Auffordern? (du, Imperativ, Ausrufezeichen)", "4d mündlich: Was wissen und können die Leser am Ende?"] }],
]));

doc.push(avivaPhase("I/V", "Sprachtraining: Imperativ", 9));
doc.push(...steps([
  ["gemeinsam", { say: "Seht euch die drei Stellen aus Absatz 5 an. Was ist bei 2 und 3 anders als bei 1?", sol: "1 „stampfst du“: Anweisung mit du + Präsens · 2 „achte“, 3 „nimm“: Imperativ – Verb vorn, ohne „du“, oft mit Ausrufezeichen", do: ["Folie 12–13."] }],
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. So bildet man den Imperativ. Der zweite Teil des Schnipsels.", do: ["Folie 14: du, e→i/ie, ihr, Sie, sein. Achtung: nicht „Nehm!“, „Lese!“"] }],
  ["luecken", { say: "Jetzt arbeitet ihr allein. Bildet den Imperativ für eine Person und für die Gruppe. Vier Minuten.", sol: "a Lauf! / Lauft! · b Nimm! / Nehmt! · c Gib ab! / Gebt ab! · d Sei! / Seid!", do: ["Folie 15, Lösung Folie 16. Mit Grün verbessern."] }],
]));

doc.push(avivaPhase("A", "Auswerten: Zielcheck", 5));
doc.push(...steps([
  ["luecken", { say: "Zielcheck. Jetzt arbeitet ihr allein im Heft. Drei Minuten.", sol: "1 informieren · 2 auffordern/anleiten – Imperativ „Stampf“, Ausrufezeichen · 3 „Lies den Text genau!“", do: ["Folie 17."] }],
  ["kontrolle", { say: "Vergleicht mit der Lösung und verbessert mit Grün. Haben wir das Ziel erreicht?", do: ["Folie 18. Notieren, wer Fehler hat. * Folie 19 (LB 5) nur bei Restzeit."] }],
]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(new Paragraph({ pageBreakBefore: true, children: [t("FÜR DIE LEHRKRAFT · REFLEXION VOR UND NACH DER STUNDE", { size: 18, color: MUTED, bold: true })], spacing: { after: 20 } }), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Ziel als „Ich kann …“ und Ablauf auf Folie und Seitentafel, * LB 5 markiert."));
doc.push(bul("Vorwissen aktivieren:", "Echte Wiederholung des Textes der letzten Stunde über Aussagen-Prüfen mit Zeile (bewährt), dazu LB 6 als motivierender Anschluss ans Fingeralphabet."));
doc.push(bul("Informieren:", "Funktionen werden erst an den Absätzen erfahren (4a), dann im Merksatz gebündelt (Schnipsel, nicht abschreiben)."));
doc.push(bul("Verarbeiten:", "LB 4b/c gemeinsam: erst Einzelarbeit, dann Unterrichtsgespräch mit Aufdeck-Folien (Entscheidung der Lehrkraft) – alle sind in der EA-Phase aktiv."));
doc.push(bul("Auswerten:", "Zielcheck mit drei kurzen Aufgaben zu beiden Teilen des Ziels."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Funktionen erkennen = kritisch lesen (KA 3: Informationen entnehmen); Imperativ für Anleitungen und Trainer-Anweisungen."));
doc.push(bul("Ich – Wir – Ihr:", "Merksätze durch die Lehrkraft (Ich), Entdecken des Imperativs im UG (Wir), Ü und Zielcheck allein (Ihr)."));
doc.push(bul("Rückmeldung:", "Lösungsfolien nach jeder Phase, Zielcheck mit Lösung."));
doc.push(bul("Klarheit:", "Feste Ansagen, gleiches Übungsformat, Zeilenzählung wie im LB."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "* Aussage 4, * LB 5 für Schnelle; Merksatz als Stütze bei 4b/c."));
doc.push(bul("", "S-PL: zweisprachiger Merksatz-Schnipsel und Handreichung mit Erklärungen und Worthilfen auf Polnisch; Fachbegriffe deutsch, keine Lösungen."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Unterscheiden die Schüler informieren / anleiten / aufklären / auffordern – oder sagen sie überall „informieren“?"));
doc.push(bul("", "Finden sie passende Belege mit Zeile? Sitzt der e→i-Wechsel beim Imperativ (nimm, lies, gib)?"));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Sachtexte · Funktionen") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync(__dirname + "/../Tafelskript_Funktionen_Sachtexte_Imperativ.docx", buf));
