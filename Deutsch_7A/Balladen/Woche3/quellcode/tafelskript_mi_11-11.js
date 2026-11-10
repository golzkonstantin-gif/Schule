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



// Tafelskript Mi 11.11. (Doppelstunde): Teil 1 Test „Paddy Fingal unter der Lupe“ · Teil 2 John Maynard (LB S. 132–133) – AVIVA, mit Präsentation
// Druckfassung ohne Farbflächen (Tinte sparen)
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const Z1 = "Ich kann die Merkmale einer Ballade an „Paddy Fingal“ nachweisen.";
const Z2 = "Ich kann Strophe für Strophe wiedergeben, was auf der „Schwalbe“ geschieht, und mit Versen belegen.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Mi 11.11. · Doppelstunde (90 Min.) · AVIVA · mit Präsentation (22 Folien)"), h1("Test · John Maynard"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Paddy Fingal ist abgeschlossen; die Hospitationsstunde am Mo 9.11. (Leseband, Kreuzworträtsel, Fehler-Chat) hat „gut funktioniert“. Teil 1 ist der kleine Test „Paddy Fingal unter der Lupe“ (Inhalt 11 P., Merkmale 15 P.). Teil 2 beginnt mit „John Maynard“ (LB S. 132–133, 8 Strophen, V. 1–62). Die LB-Aufgaben 1–3 werden bewusst gemeinsam bearbeitet, weil die Klasse mit der neuen Ballade am Anfang steht (Entscheidung der Lehrkraft). Damit trotzdem alle aktiv sind: Ich – Wir – Ihr, je Strophe erst zu zweit Stichworte, dann aufdecken. Das Rollenlesen knüpft an „Dramatik = Figuren sprechen ohne Erzähler“ aus dem Test an. Kein eigenes Sprachtraining in dieser Doppelstunde (Test); der weil-Satz im Zusatz des Tests greift die Subjunktionen auf.")], { after: 0 })]));
doc.push(p([S("Ziel Teil 1: ", { bold: true }), S(Z1, { italics: true })], { before: 120, after: 20 }));
doc.push(p([S("Ziel Teil 2: ", { bold: true }), S(Z2, { italics: true })], { after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„John Maynard ist Steuermann auf der ‚Schwalbe‘. Als ein Feuer ausbricht, hält er das Steuer fest, bis das Schiff den Strand erreicht. Alle werden gerettet, nur er stirbt. Das kann ich mit Versen belegen.“ Überprüft im Zielcheck.")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Teil 1 · Test – Ziel: ", { bold: true, size: 20, color: NAVY }), t(Z1, { size: 20, italics: true })], { after: 40 }),
  p([t("Teil 2 · John Maynard – Ziel: ", { bold: true, size: 20, color: NAVY }), t(Z2, { size: 20, italics: true })], { after: 40 }),
  ...["Ü1 Erstes Lesen, spannende Stellen", "Ü2 Wer spricht? Lesen mit verteilten Rollen", "Ü3 Was geschieht auf der „Schwalbe“? (* 3d Zusammenfassen)", "Zielcheck"].map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["3′", "A", "1–2", "Ankommen: Ziele und Ablauf beider Teile", "lesen Ziele und Ablauf"],
  ["30′", "–", "3–4", "Teil 1: Test (Gruppe A/B, S-PL mit polnischer Fassung)", "schreiben den Test; wer fertig ist, liest John Maynard still (= Ü1a)"],
  ["2′", "A", "5", "Übergang Teil 2: neue Ballade", "legen LB S. 132 bereit, Überschrift ins Heft"],
  ["8′", "V/I", "6–7", "Ü1 Erstes Lesen, drei spannende Stellen, Countdown", "lesen still, notieren Verse, sammeln"],
  ["12′", "I/V", "8–10", "Ü2 Rollen finden, Rollenlesen (Methodenkarte)", "suchen zu zweit Rollen, lesen laut oder still mit und notieren eine gut betonte Stelle"],
  ["25′", "V", "11–18", "Ü3 Tabelle Strophe für Strophe (Ich – Wir)", "notieren je Strophe zu zweit Stichworte, vergleichen, ergänzen mit Grün"],
  ["*", "V", "19–20", "* Ü3d Zusammenfassen", "fassen gemeinsam im Präsens zusammen"],
  ["5′", "A", "21–22", "Zielcheck", "lösen allein, verbessern mit Grün"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("85 Min. + 5 Min. Puffer. Wenn es eng wird: Ü3 nach Strophe 6 abbrechen, Strophe 7 und Rahmen (1 + 8) sowie * 3d als Einstieg der nächsten Stunde.")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "John Maynard (Theodor Fontane) · LB S. 132–133 · Datum rechts"],
  ["Ü1", "Drei spannende Stellen mit Vers und Stichwort · Countdown V. 17, 23, 29, 39, 45"],
  ["Ü2", "Rollen mit je einem Vers"],
  ["Ü3", "Tabelle: Strophe | a Fahrt | b Passagiere | c John Maynard – Strophe 1–8, mit Grün ergänzt · * Zusammenfassung"],
  ["Zielcheck", "Drei Aussagen: r/f, Vers, Verbesserung, mit Grün kontrolliert"],
].map(([a, b]) => row([tc(a, 2300, { bold: true }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation `Mittwoch_11-11_Deutsch_7A.pptx` (Lösungen auf Aufdeck-Folien und in den Notizen) · Test `Leistungskontrolle_Balladen/` Gruppe A/B + Text Paddy Fingal (Kopie LB S. 136 f.) · Lehrbuch S. 132–133 · Handreichung Polnisch für S-PL (`Handreichung_Polnisch_11-11`) und polnische Testfassung.")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziele und Ablauf", 3));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Heute zwei Teile: Zuerst schreibt ihr den Test zu Paddy Fingal. Danach beginnen wir eine neue Ballade: John Maynard. Am Ende prüft ihr im Zielcheck, ob ihr das Ziel erreicht habt.", do: ["Folie 1–2. Überblick mit Zeiten und je einem Ich-Ziel pro Teil."] }]]));

doc.push(avivaPhase("1", "Teil 1: Test „Paddy Fingal unter der Lupe“", 30));
doc.push(...steps([
  ["luecken", { say: "Tische auseinander, nur Stift und Lineal liegen auf dem Tisch. Ich teile Gruppe A und B aus, dazu den Text. Ihr arbeitet allein und still. Ca. 30 Minuten. Wer fertig ist: Test umdrehen, Lehrbuch Seite 132 aufschlagen und John Maynard still lesen.", do: ["Folie 3–4. S-PL bekommt die Fassung mit polnischer Übersetzung der Aufgaben.", "Die Frühfertigen erledigen so schon Ü1a."] }],
  ["fertig", { say: "Stifte hinlegen. Ich sammle ein.", do: ["Tests einsammeln (A und B getrennt)."] }],
]));

doc.push(avivaPhase("V/I", "Teil 2 · Ü1 Erstes Lesen", 10));
doc.push(...steps([
  ["luecken", { say: "Neue Ballade: John Maynard von Theodor Fontane, Lehrbuch Seite 132. Schreibt die Überschrift ins Heft. Jetzt arbeitet ihr allein: Lest die Ballade still. Notiert drei Textstellen, die ihr besonders spannend findet, jeweils mit Vers und einem Stichwort, warum. Acht Minuten.", do: ["Folie 5–6. Methode: stilles Lesen (erstes, orientierendes Lesen). Wer schon gelesen hat, beginnt direkt mit den drei Stellen."] }],
  ["gemeinsam", { say: "Wer drankommt, nennt einen Vers und sagt, warum die Stelle spannend ist. Alle ergänzen mit Grün.", sol: "Countdown V. 17 (dreißig Minuten), V. 23 (zwanzig), V. 29 (fünfzehn), V. 39 (zehn), V. 45 (Rettung) · weitere: V. 20 „Feuer!“, V. 34–37 Kapitän – John Maynard, V. 41 „mit ersterbender Stimme“, V. 47 „Nur einer fehlt!“", do: ["Folie 7 nach dem Sammeln aufdecken: Countdown – die Zeit wird kürzer, die Lage gefährlicher."] }],
]));

doc.push(avivaPhase("I/V", "Ü2 Wer spricht? Lesen mit verteilten Rollen", 12));
doc.push(...steps([
  ["partner", { say: "Jetzt arbeitet ihr zu zweit. Sucht alle Figuren, die in der Ballade sprechen, und notiert zu jeder einen Vers. Überlegt auch: Welche Verse spricht der Erzähler? Vier Minuten.", sol: "Erzähler · Stimme 1/2 (V. 1–7) · Passagiere (V. 15, 28) · John Maynard (V. 17, 35, 37, 41) · Kapitän (V. 34, 36, 40) · Schiffsvolk (V. 20, 38) · Chor: Grabinschrift (V. 58–62)", do: ["Folie 8, dann Folie 9 (Lösung). Bezug zum Test: Wo Figuren ohne Erzähler sprechen, wirkt die Ballade wie ein Theaterstück."] }],
  ["gemeinsam", { say: "Wer eine Rolle hat, liest laut und betont. Alle anderen lesen still mit und notieren eine Stelle, die besonders gut betont war.", do: ["Folie 10. Methodenkarte: Lautlesetraining, Variante Rollenlesen (Plenum). Rollen verteilen: Erzähler 2–3 Schüler im Wechsel, je 1–2 pro Figur, V. 58–62 alle im Chor.", "Danach 2–3 Rückmeldungen zur Betonung (Countdown drängender, V. 41 leise)."] }],
]));

doc.push(avivaPhase("V", "Ü3 Was geschieht auf der „Schwalbe“? (gemeinsam)", 25));
doc.push(...steps([
  ["abschreiben", { say: "Jetzt schreibt ihr ab: Zeichnet die Tabelle ins Heft – Strophe, a Fahrt, b Passagiere, c John Maynard.", do: ["Folie 11. Tipp: Tabelle quer über die Doppelseite oder schmale Spalte für die Strophe."] }],
  ["zuhoeren", { say: "Stifte liegen. Ich mache Strophe 2 vor. Was erfahre ich über die Fahrt? … über die Passagiere? … über John Maynard? Ich schreibe in eigenen Worten, nicht abschreiben. Jetzt schreibt ihr die Zeile ab.", do: ["Folie 12 (Ich): laut denken."] }],
  ["partner", { say: "Strophe 3. Eine Minute zu zweit: Notiert Stichworte in allen drei Spalten. Wer drankommt, nennt sie. Dann decke ich auf. Vergleicht und ergänzt mit Grün.", sol: "siehe Folien 13–18 (Str. 3, 4, 5, 6, 7, Rahmen 1 + 8)", do: ["Folie 13–17 (Wir) im gleichen Ablauf: Strophe 3 bis 7. Niemand geht an die Tafel.", "Bei Str. 3 nachfragen: Warum ist John Maynard nicht zu sehen? (dichter Qualm am Steuer, V. 27). Bei Str. 6: Warum nur zwei Verse? (Schock, Pause).", "Folie 18: Strophe 1 und 8 als Rahmen – Anfang = Ende."] }],
  ["gemeinsam", { say: "* Wir fassen den Inhalt gemeinsam zusammen. Eine Zusammenfassung steht im Präsens. Der Anfang steht auf der Folie.", sol: "siehe Folie 20 (Musterzusammenfassung)", do: ["Folie 19–20 nur, wenn Zeit bleibt; sonst Einstieg der nächsten Stunde. Stichwörter je Abschnitt und Verknüpfungswörter (zuerst, dann, als, weil, obwohl, schließlich) auf der Folie."] }],
]));

doc.push(avivaPhase("A", "Auswerten: Zielcheck", 5));
doc.push(...steps([
  ["luecken", { say: "Zielcheck. Jetzt arbeitet ihr allein im Heft. Richtig oder falsch? Schreibt r oder f und den Vers als Beleg. Verbessert falsche Aussagen. Vier Minuten.", sol: "1 richtig (V. 20–23) · 2 falsch (V. 41 / V. 58–59) – er bleibt am Steuer · 3 falsch (V. 47) – John Maynard stirbt", do: ["Folie 21."] }],
  ["kontrolle", { say: "Vergleicht mit der Lösung und verbessert mit Grün. Haben wir das Ziel erreicht?", do: ["Folie 22. Herumgehen und notieren, wer noch Fehler hat – in der Wiederholung der nächsten Stunde gezielt aufgreifen."] }],
]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(new Paragraph({ pageBreakBefore: true, children: [t("FÜR DIE LEHRKRAFT · REFLEXION VOR UND NACH DER STUNDE", { size: 18, color: MUTED, bold: true })], spacing: { after: 20 } }), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Überblicksfolie für die Doppelstunde mit Ablauf, Zeiten und je einem Ich-Ziel pro Teil (Rückmeldung 14.10.); Ziele auch an der Seitentafel."));
doc.push(bul("Vorwissen aktivieren / Informieren:", "Erstes Lesen mit klarem Auftrag (drei Stellen mit Vers) statt freiem Leseeindruck – das Belegen aus der Paddy-Fingal-Einheit wird weitergeführt. Der Countdown macht die Spannung sichtbar."));
doc.push(bul("Verarbeiten:", "Ü3 gemeinsam (Entscheidung der Lehrkraft), aber mit Ich – Wir und je Strophe einer Minute Partnerarbeit, damit alle aktiv sind; Aufdeck-Folien Strophe für Strophe, niemand an der Tafel."));
doc.push(bul("Auswerten:", "Zielcheck mit drei Aussagen (richtig/falsch + Vers) – das bewährte Format „Aussagen prüfen“."));
doc.push(bul("Zeit:", "Nach dem Test bleibt wenig Zeit; * 3d und klare Abbruchstelle (nach Strophe 6) eingeplant."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Ein überprüfbares Ziel pro Teil; Inhalt sicher wiedergeben ist die Grundlage für Analyse und Schreiben aus der Figurenperspektive (Bewertungen der Einheit)."));
doc.push(bul("Ich – Wir – Ihr:", "Lehrkraft macht Strophe 2 vor (Ich), Strophen 3–8 gemeinsam mit Partnerphase (Wir), Zielcheck allein (Ihr)."));
doc.push(bul("Rückmeldung:", "Lösungsfolien nach Ü1, Ü2 und jeder Strophe in Ü3; Zielcheck mit Lösung."));
doc.push(bul("Klarheit:", "Feste wörtliche Ansagen; Aufgaben im gleichen Format; Tabelle mit festen Spalten."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "Frühfertige im Test lesen John Maynard schon still (Ü1a). Rollenlesen: unsichere Leser bekommen kurze Rollen (Passagiere, Schiffsvolk) oder lesen im Chor mit."));
doc.push(bul("", "S-PL: Test mit polnischer Übersetzung; Handreichung mit Zielen, Ablauf, Aufgaben auf Polnisch und Worterklärungen zu John Maynard (Eriesee, Gischt, Bugspriet, Qualm, Sprachrohr, Brandung …); Fachbegriffe bleiben deutsch, keine Lösungen."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Test: Welche Aufgaben fallen schwer (Belegen/Zitieren, Epik vs. Dramatik)? → Wiederholung in der nächsten Stunde."));
doc.push(bul("", "Ü3: Schreiben die Schüler in eigenen Worten oder schreiben sie ab? Erkennen sie den Rahmen (Str. 1 = Str. 8)?"));
doc.push(bul("", "Zeit: Wie weit kommt die Klasse in Ü3? Wo setzt die nächste Stunde an?"));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Mi 11.11.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync(__dirname + "/../Tafelskript_Mi_11-11_Test_John_Maynard.docx", buf));
