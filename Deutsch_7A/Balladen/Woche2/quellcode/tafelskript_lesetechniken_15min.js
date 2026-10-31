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



// Tafelskript „Lesetechniken wiederholen“ – Kurzeinheit 15 Min. nach den Herbstferien – AVIVA
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const zielTafel = (ziel, ablauf) => board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ziel, { size: 21, italics: true })], { after: 60 }),
  p(t("Ablauf", { bold: true, size: 20, color: NAVY }), { after: 20 }),
  ...ablauf.map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]);
const ZIEL = "Ich kann die drei Lesetechniken unterscheiden und begründen, welche zu meinem Leseziel passt.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · erste Stunde nach den Herbstferien · nur 15 Min. · AVIVA · ohne Präsentation"), h1("Lesetechniken wiederholen"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Die letzte Stunde vor den Ferien (Fr 16.10.) ist ausgefallen. Offen sind LB-Aufgabe 1 a zu Paddy Fingal, die Reime und die Merkmale der Ballade (Ur-Ei) – darauf konzentriert sich diese Woche. Nach zweieinhalb Wochen Pause werden in den 15 Minuten die Lesetechniken aufgefrischt (eingeführt am 14.10., wiederholt am 15.10.). Die Kurzeinheit endet mit einem einminütigen Überfliegen von Paddy Fingal und baut so die Brücke zur nächsten Stunde.")], { after: 0 })], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S(ZIEL, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Ich wähle die Lesetechnik nach meinem Leseziel: Will ich wissen, worum es geht, lese ich überfliegend. Suche ich eine bestimmte Information, lese ich orientierend. Will ich alles verstehen, lese ich genau.“")], { after: 80 }));
doc.push(zielTafel(ZIEL, ["Lesetechniken aus dem Gedächtnis", "Welche Technik passt? (A–D)", "* Paddy Fingal überfliegen", "Ziel erreicht? Heute habe ich gelernt, …"]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([900, 1300, 3400, 4606], ["Zeit", "AVIVA", "Phase", "Alle Schüler …"], [
  ["1′", "A", "Ankommen: Ziel und Ablauf", "lesen Ziel und Ablauf, schreiben die Überschrift"],
  ["4′", "V", "Lesetechniken aus dem Gedächtnis", "notieren allein die drei Techniken, vergleichen mit dem Nachbarn"],
  ["1′", "I", "Leitfrage: Was ist mein Leseziel?", "hören zu"],
  ["6′", "V", "Situationen A–D zuordnen und begründen; * Paddy Fingal überfliegen", "ordnen allein zu, begründen mit „weil“; * überfliegen 1 Min."],
  ["3′", "A", "Auswerten: Ziel erreicht?", "vergleichen in der Kreide-Kette, schreiben „Heute habe ich gelernt, …“"],
], { bold: [true, true, true, false], colors: [C.hv, C.hv, NAVY, null], h: 340 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("keiner. Mit * markiert: Paddy Fingal überfliegen – fällt es weg, ist es der Einstieg der nächsten Stunde.")], { before: 80, after: 60 }));
doc.push(p([S("Heft: ", { bold: true }), S("Überschrift „Lesetechniken wiederholen“ · drei Lesetechniken mit „Ich lese so, wenn …“ · A–D mit Technik und Begründung · * „In der Ballade geht es um …“ · „Heute habe ich gelernt, …“")], { after: 40 }));
doc.push(p([S("Material: ", { bold: true }), S("Lehrbuch S. 136 f. · Schnipsel „Merksatz Lesetechniken“ (Woche 1) für alle, die ihn nicht mehr im Heft haben · Handreichung Polnisch für S-PL.")], { after: 40 }));
doc.push(p([S("Tipp: ", { bold: true, color: C.hv }), S("Ziel, Ablauf und die Situationen A–D vor der Stunde anschreiben, A–D zunächst zugeklappt.")], { after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 1));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Willkommen zurück. Wir beschäftigen uns weiter mit Balladen. Heute frischen wir in einer Viertelstunde die Lesetechniken auf – links stehen Ziel und Ablauf. Jetzt schreibt ihr die Überschrift ab: Lesetechniken wiederholen." }]]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Lesetechniken aus dem Gedächtnis", 4));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein, ohne Heft von vorher und ohne Buch. Schreibt die drei Lesetechniken untereinander und ergänzt zu jeder: Ich lese so, wenn … Zwei Minuten. Fertig? Stift hinlegen.", sol: "überfliegend – wenn ich schnell wissen will, worum es geht · orientierend – wenn ich eine bestimmte Information suche · genau – wenn ich alles verstehen will (Satz für Satz, Unbekanntes klären)" }],
  ["kontrolle", { say: "Vergleicht mit eurem Nachbarn und verbessert mit Grün. Schaut dazu in euren Merksatz von vor den Ferien.", do: ["Wer den Merksatz nicht im Heft hat, bekommt den Schnipsel „Lesetechniken“ zum Einkleben."] }],
]));

doc.push(avivaPhase("I", "Informieren: Leitfrage", 1));
doc.push(board("Mitteltafel", [bl([N("Was ist mein Leseziel?  →  Dann wähle ich die Lesetechnik.", { bold: true, color: NAVY })])]));
doc.push(p(t(""), { after: 40 }));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Es gibt keine beste Lesetechnik. Entscheidend ist euer Leseziel: Was will ich aus dem Text herausholen? Danach wählt ihr die Technik." }]]));

doc.push(avivaPhase("V", "Verarbeiten: Welche Lesetechnik passt?", 6));
doc.push(board("Mitteltafel · Welche Lesetechnik passt? Begründe mit „weil“.", [
  bl("A  Du willst im Trainingsplan nachsehen, wann am Mittwoch das Training beginnt."),
  bl("B  Du blätterst eine Sportzeitschrift durch und willst wissen, ob ein Artikel dich interessiert."),
  bl("C  Vor dem Wettkampf liest du die neuen Wettkampfregeln, damit du nicht disqualifiziert wirst."),
  bl("D  Du sollst erklären, warum der Schotte in „Paddy Fingal“ am Ende flieht."),
  bl([N("* E  ", { bold: true }), N("Überfliege Paddy Fingal (1 Min.). Schreibe: In der Ballade geht es um …")], { before: 60 }),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht an der Tafel. Schreibt A bis D, die Lesetechnik und eine Begründung mit weil. Vier Minuten. Wer fertig ist, macht E mit Stern. Fertig? Stift hinlegen.", sol: "A orientierend, weil ich nur eine Uhrzeit suche · B überfliegend, weil ich nur wissen will, worum es geht · C genau, weil ich jede Regel verstehen muss · D genau, weil ich etwas erklären soll · *E z. B. … um den Riesen Paddy Fingal und einen noch größeren Schotten, vor dem er sich mit einer List seiner Frau rettet.", do: ["Für E: Lehrbuch S. 136 auf, Zeit laut stoppen. E ist zugleich der Einstieg in die Arbeit an Paddy Fingal in dieser Woche."] }],
]));

doc.push(avivaPhase("A", "Auswerten: Ziel erreicht?", 3));
doc.push(...steps([
  ["tafel", { say: "Wer drankommt, schreibt den Buchstaben und die Lesetechnik an die Tafel und sagt die Begründung. Alle prüfen mit und verbessern mit Grün.", do: ["Kreide-Kette: vier Schüler, je eine Situation. Bei D nachfragen: „Warum reicht orientierend nicht?“"] }],
  ["luecken", { say: "Schaut auf unser Ziel links. Schreibt ins Heft: Heute habe ich gelernt, …", do: ["Ein bis zwei Schüler lesen vor. Ausblick: „In den nächsten Stunden arbeiten wir mit Paddy Fingal weiter – dafür braucht ihr die Lesetechniken.“"] }],
]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(p(new PageBreak()));
doc.push(kicker("Für die Lehrkraft · Reflexion vor und nach der Stunde"), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Auch in 15 Minuten vollständig:", "Ziel und Ablauf an der Seitentafel, jede Phase dient dem Ziel, Rückgriff in der Auswertung."));
doc.push(bul("Vorwissen aktivieren:", "Echte Wiederholung nach der Ferienpause – erst aus dem Gedächtnis abrufen, dann mit dem Merksatz prüfen."));
doc.push(bul("Zeit:", "Bewusst wenig Inhalt; der einzige Zusatz (E) ist mit * markiert und dient als Brücke zur nächsten Stunde."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Das Leseziel bestimmt die Technik – mit Situationen aus Training und Wettkampf und aus der Ballade."));
doc.push(bul("Abruf statt Wiederholung durch die Lehrkraft:", "Die Schüler rufen das Wissen selbst ab (Abrufübung), der Merksatz dient zur Selbstkontrolle."));
doc.push(bul("Ihr und Rückmeldung:", "Zuordnung mit Begründung allein, Vergleich in der Kreide-Kette mit Auftrag für alle."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "* E für Schnelle; Merksatz-Schnipsel für alle, denen er fehlt; Handreichung Polnisch für S-PL."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Wissen die Schüler die drei Techniken nach den Ferien noch, oder müssen sie in der nächsten Stunde erneut wiederholt werden?"));
doc.push(bul("", "Können sie ihre Wahl begründen (D: genau statt orientierend)?"));
doc.push(bul("", "Wurde E geschafft? Sonst ist E der Einstieg der nächsten Stunde."));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · nach den Herbstferien") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Lesetechniken_wiederholen_15min.docx", buf));
