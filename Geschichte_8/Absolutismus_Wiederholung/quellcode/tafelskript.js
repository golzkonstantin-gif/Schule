const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak, Header, Footer, TabStopType, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", BOARD = "F3F6F1", RED = "D9534F", GREEN = "2E9E6B";
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100 }, keepNext: o.keepNext });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h3 = (text) => p(t(text, { font: "Cambria", size: 23, bold: true, color: NAVY }), { before: 120, after: 60, keepNext: true });
const br = () => new Paragraph({ children: [new PageBreak()] });
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));
const N = (s, o = {}) => t(s, Object.assign({ font: "Cambria", size: 22 }, o));

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const thin = all(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || thin, verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font }), { after: 0 }), { w, fill: o.fill });

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Geschichte · Klasse 8 · Absolutismus Wiederholung", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Seite ", PageNumber.CURRENT] })] })] });

const LABELS = { zuhoeren: "hört zu", gemeinsam: "macht mündlich mit", abschreiben: "schreibt ab", allein: "arbeitet allein", tafel: "kommt an die Tafel", kontrolle: "kontrolliert", };
const BW = 1700, CW = W - BW, TW = CW - 340;
function labelCell(m) {
  return cell([p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }), p(t(LABELS[m], { bold: true, color: NAVY, size: 20 }), { after: 0 })], { w: BW, fill: LIGHT, borders: all(solid("D5DDEE", 6)), m: 100, valign: VerticalAlign.TOP });
}
function step(m, parts) {
  const ch = [];
  if (parts.say) ch.push(p([t("Ansage: ", { bold: true, size: 20, color: RED }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((d) => ch.push(p(Array.isArray(d) ? d : S(d), { after: 50 })));
  if (parts.board) { ch.push(...[].concat(parts.board)); ch.push(p(t(""), { after: 20 })); }
  if (parts.sol) ch.push(p([t("Lösung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  return table([BW, CW], [row([labelCell(m), cell(ch, { w: CW, borders: all(solid("D5DDEE", 6)), m: 100, ml: 160, valign: VerticalAlign.TOP })])]);
}
const steps = (arr) => arr.flatMap((s) => [step(s[0], s[1]), p(t(""), { after: 40 })]);
function boxed(label, fill, col, children) {
  return table([TW], [
    row([cell(p(t(label, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: col, borders: all(solid(col, 12)), m: 30 })]),
    row([cell(children, { w: TW, fill, borders: all(solid(col, 12)), m: 100, ml: 160 })]),
  ]);
}
const board = (title, ch) => boxed("TAFEL · " + title, BOARD, "3D4A3F", ch);
const folie = (title, ch) => boxed("FOLIE · " + title, LIGHT, NAVY, ch);
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
const merkBoard = (nr, title, text) => folie(`Merksatz ${nr} – abschreiben, rot umrahmen`, [p(t(title, { bold: true, color: RED, font: "Cambria", size: 22 }), { after: 40 }), ...text.split("  ").map((x, i, a) => p(N(x, i ? { size: 20, color: MUTED } : {}), { after: i < a.length - 1 ? 50 : 0 }))]);
function auftrag(title, todo, lines, meta) {
  return board("rechts · " + title, [
    ...todo.map((x, i) => p([t(`${i + 1}. `, { bold: true, color: NAVY, size: 21 }), t(x, { bold: true, color: NAVY, size: 21 })], { after: 30 })),
    ...(lines.length ? [p(t(""), { after: 30 }), ...lines] : []),
    p(t(""), { after: 30 }),
    p([t("Zeit: ", { bold: true, size: 19, color: "3D4A3F" }), t(meta, { size: 19, color: "3D4A3F" }), t("     Fertig? Stift hinlegen und noch einmal durchlesen.", { size: 19, italics: true, color: "3D4A3F" })], { after: 0 }),
  ]);
}
const phase = (nr, title, min) => p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: RED }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });

// Tabelle Ständegesellschaft (Tafelbild)
const TBW = [1000, 800, 1350, 1800, 1750, 1466];
const stTable = (filled) => {
  const rows = [
    ["1. Stand\nKlerus", "ca. 150 000", "meist sehr gebildet, fast alle können lesen und schreiben", "keine Steuern · eigene Gerichte · vom Wehrdienst befreit · Kirchenzehnt (10 % vom Einkommen der Bauern und Bürger)", "Fürsorge für die Armen", "im Text nicht genannt"],
    ["2. Stand\nAdel", "ca. 500 000", "sehr gebildet, Kinder besuchen die Schule", "keine Steuern · alleiniges Jagdrecht · ranghohe Posten in der Armee · bestimmt das politische Geschehen", "verwaltende Aufgaben · ehrenhaftes Verhalten", "ja, bestimmt das politische Geschehen"],
    ["3. Stand\nBauern, Bürger", "ca. 20 000 000", "nur wenige können lesen und schreiben", "keine; kaum Aufstiegsmöglichkeit, man wird in den Stand hineingeboren", "Steuern und hohe Abgaben · Pachtgebühren an Grundherren · im Krieg Soldaten stellen", "ausgeschlossen"],
  ];
  return table(TBW, [
    row(["Stand", "Anzahl", "Bildung", "Vorrechte", "Pflichten", "Mitbestimmung"].map((h, i) => hdr(h, TBW[i], "3D4A3F", 15)), 330),
    ...rows.map((r) => row(r.map((v, i) => tc(filled ? v : "", TBW[i], { size: 17, bold: i === 0, fill: i === 0 ? BOARD : undefined })), filled ? 700 : 600)),
  ]);
};

// ================================================================
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.)"), h1("Absolutismus und Ständegesellschaft – Wiederholung vor dem Test"));
doc.push(p(t("Ziel: Die Klasse ist auf den Test nächste Woche vorbereitet (Ständegesellschaft, Begriff Absolutismus, fünf Säulen) und erkennt selbst, wo noch Lücken sind. Klarheit entsteht durch eindeutige Ansagen und ein festes Aufgabenformat. Einziges Arbeitsblatt: der Lückentext zu den fünf Säulen (mit Lösung). Merksätze stehen auf den Folien, an der Tafel entstehen nur das Säulen-Tafelbild und die Tabelle.", { size: 21 }), { after: 100 }));

doc.push(h3("Anbindung an die letzte Stunde"));
doc.push(p(S("Die fünf Säulen wurden im Gruppenpuzzle erarbeitet (Text „Herrschaft ohne Stände“), jede Gruppe kannte nur ihre Säule. Deshalb hat vermutlich jede und jeder nur einen Teil gesichert. Der Lückentext bringt alle fünf Säulen zusammen. Die Ständegesellschaft liegt schon länger zurück, deshalb wird sie zuerst aus dem Gedächtnis abgerufen (Hefte zu), danach erst verglichen.")), { after: 60 });

doc.push(h3("Ablauf im Überblick"));
doc.push(table([900, 4400, 2100, 2806], [
  row(["Min.", "Phase", "Material", "Arbeitsform"].map((h, i) => hdr(h, [900, 4400, 2100, 2806][i], NAVY, 19)), 340),
  ...[
    ["2", "0  Testankündigung", "Folie 2", "zuhören"],
    ["4", "1  Einstieg: Wahr oder falsch?", "Folie 3", "mündlich, gemeinsam"],
    ["7", "2  Absolutismus: Merksatz 1 + Ü1", "Folie 4–5", "abschreiben, allein, Tafel"],
    ["11", "3  Fünf Säulen: Merksatz 2 + Lückentext (Ü2)", "Folie 6–7, Arbeitsblatt", "abschreiben, allein, Kontrolle"],
    ["10", "4  Ständegesellschaft: Tabelle (Ü3) + Merksatz 3", "Folie 8–9, Tafelbild", "gemeinsam, abschreiben"],
    ["9", "5  „Im Absolutismus fühle ich mich …“ (Ü4)", "Folie 10–11", "allein, Kontrolle"],
    ["2", "6  Lernliste und Test", "Folie 12", "zuhören"],
  ].map((r) => row(r.map((v, i) => tc(v, [900, 4400, 2100, 2806][i], { size: 20, bold: i < 2 })), 340)),
]));
doc.push(p(S("Zeitpuffer: Wird es knapp, die Kontrolle zu Ü1 auf zwei Sätze kürzen (schon eingeplant) und bei Ü4 nur zwei Schüler vorlesen lassen. Gibt es eine Doppelstunde, Ü3 als Partnerquiz verlängern.", { italics: true, color: MUTED }), { before: 80, after: 60 }));

doc.push(h3("Eindeutige Ansagen (bei jedem Wechsel, wörtlich)"));
doc.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären oder fragen, niemand schreibt."],
  ["„Meldet euch. Ihr schreibt noch nichts auf.“", "Gemeinsam an der Tafel erarbeiten."],
  ["„Jetzt schreibt ihr ab: …“", "Merksatz oder Tabelle wird ins Heft übernommen."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Übung im Heft, leise."],
  ["„Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“", "Lösungen sammeln."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Heft."],
].map(([a, b]) => row([tc([t(a, { italics: true, bold: true, size: 20, color: NAVY })], 4700, { fill: LIGHT }), tc(b, 5506, { size: 20 })], 380))));
doc.push(p([S("Gleicher Ablauf je Phase: ", { bold: true }), S("gemeinsam entdecken → Merksatz abschreiben → allein üben → an der Tafel vergleichen → mit Grün verbessern.")], { before: 100, after: 40 }));
doc.push(p([S("Material: ", { bold: true }), S("Beamer, Kreide/Stifte in Rot und Grün · Arbeitsblatt Lückentext (Klassensatz) und Lösung für die Lehrkraft · Schüler: Heft, Lineal, Stifte Rot und Grün")], { after: 40 }));

doc.push(br());

// Phase 0
doc.push(phase("0", "Testankündigung", 2));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Nächste Woche schreiben wir einen Test zu diesen drei Themen.", do: ["Folie 2 zeigen. Die drei Themen vorlesen. Kernbotschaft: Die Stunde heute ist die Vorbereitung."] }],
]));

// Phase 1
doc.push(phase("1", "Einstieg: Wahr oder falsch?", 4));
doc.push(...steps([
  ["gemeinsam", { say: "Stifte liegen. Ihr schreibt noch nichts auf. Entscheidet still: wahr oder falsch? Dann meldet ihr euch und begründet.", do: ["Die vier Behauptungen stehen auf Folie 3 (oder vorher an der Tafel). Pro Behauptung eine Antwort mit Begründung, bei „falsch“ wird die Aussage verbessert.", "Die Antworten zeigen, was schon sitzt. Nicht ausführlich erklären: Was unsicher ist, kommt in den folgenden Phasen."], board: board("Mitte · Behauptungen", [bl("1  Der Dritte Stand war mit Abstand der größte Stand."), bl("2  Der Adel musste hohe Steuern an den König zahlen."), bl("3  Ludwig XIV. verzichtete auf Söldner und baute ein stehendes Heer auf."), bl("4  Ludwig XIV. erlaubte den Protestanten die freie Religionsausübung.", { after: 0 })]), sol: "1 wahr (ca. 20 Millionen von gut 20,6 Millionen) · 2 falsch: Der Adel musste keine Steuern zahlen (Vorrecht) · 3 wahr (Säule Heer) · 4 falsch: 1685 beendete er die religiöse Toleranz, ein König, ein Glaube (Säule Religion)." }],
]));

doc.push(phase("2", "Absolutismus", 7));
doc.push(...steps([
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf.", do: ["Frage: Was bedeutet das Wort „absolut“? Wo habt ihr es schon gehört (absolut richtig, absoluter Nullpunkt)?", "Erwartung: uneingeschränkt, ohne Ausnahme."], board: board("Mitte", [bl([N("Absolut", { bold: true, color: RED }), N("-ismus:  absolut = uneingeschränkt")], {}), bl("Herrschaft ohne Kontrolle durch andere")]) }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.", board: merkBoard(1, "Absolutismus", "Absolutismus war eine Herrschaftsform in Europa im 17. und 18. Jahrhundert, in der ein einzelner Monarch (König oder Fürst) uneingeschränkt (absolut) herrschte.  Beispiel: Ludwig XIV. verkündet 1661: „Ich regiere jetzt selbst.“  Wofür? Den Begriff einordnen – Grundlage für den Test.  Test: Kann jemand den König stoppen? Nein → Absolutismus.") }],
  ["allein", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.", board: auftrag("Ü1 Absolutismus – ja oder nein?", ["Schreibe Ü1 an den Rand.", "Schreibe Nummer + „Absolutismus“ oder „kein Absolutismus“.", "Begründe mit dem Test: Kann jemand den König stoppen?"], [bl("1  Der König entscheidet allein über Gesetze, Krieg und Steuern."), bl("2  Ein Parlament kann Steuererhöhungen des Königs verbieten."), bl("3  Beamte handeln im Namen des Königs und müssen ihm gehorchen."), bl("4  Die Bürger wählen alle vier Jahre ihre Regierung.", { after: 0 })], "4 Minuten · allein · leise") }],
  ["tafel", { say: "Wer drankommt, schreibt an die Tafel. Alle prüfen mit. … Vergleicht und verbessert mit Grün.", sol: "1 Absolutismus (niemand kontrolliert ihn) · 2 kein Absolutismus (das Parlament kann den König stoppen) · 3 Absolutismus (Beamte sind ihm zum Gehorsam verpflichtet, er kontrolliert sie) · 4 kein Absolutismus (Wahl = Macht ist begrenzt und auf Zeit)." }],
]));

// Phase 2
doc.push(phase("3", "Die fünf Säulen", 11));
doc.push(...steps([
  ["abschreiben", { say: "Jetzt schreibt ihr von der Folie ab: Merksatz 2. Rahmt ihn rot ein.", do: ["Kurz fragen: Wer hat welche Säule im Gruppenpuzzle bearbeitet? (Stärkt das Wiedererkennen.)"], board: merkBoard(2, "Die fünf Säulen des Absolutismus", "Ludwig XIV. sicherte seine Macht auf fünf Säulen: 1 Verwaltung und Justiz · 2 Adel und Hof von Versailles · 3 Wirtschaft (Merkantilismus) · 4 Religion (Gottesgnadentum) · 5 Heer (stehendes Heer).  Wofür? Im Test nennen und erklären.  Test: Was sichert der König damit?") }],
  ["allein", { say: "Ich teile das Arbeitsblatt aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.", board: auftrag("Ü2 Lückentext: Die fünf Säulen", ["Schreibe Ü2 oben auf das Arbeitsblatt.", "Lies den Text einmal ganz. Ergänze die 12 Lücken mit dem Wortspeicher.", "Streiche jedes benutzte Wort durch. Zwei Wörter passen nicht."], [], "6 Minuten · allein · leise"), do: ["Schnelle: ★-Aufgabe auf dem Arbeitsblatt (wichtigste Säule in einem Satz begründen)."] }],
  ["tafel", { say: "Wer drankommt, liest einen Absatz vor. Alle prüfen mit. … Vergleicht und verbessert mit Grün.", do: ["Pro Säule eine Schülerin oder einen Schüler den Absatz vorlesen lassen. Bei Fehlern nach der Begründung fragen: „Welche Stelle im Text spricht dagegen?“"], sol: "1 absoluter · 2 Beamte · 3 Intendanten · 4 Steuerfreiheit · 5 Versailles · 6 Merkantilismus · 7 Zölle · 8 Gottesgnadentum · 9 Glaube · 10 stehendes · 11 400.000 · 12 Drittel. Nicht passend: Söldner (Ludwig verzichtete auf sie), Parlament (im Absolutismus kann niemand den König stoppen)." }],
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf. Was sichert der König mit jeder Säule?", do: ["Das Tafelbild entsteht gemeinsam, nicht abzeichnen (steht schon im Merksatz und im Lückentext):"], board: board("Mitte · Säulen-Tafelbild", [
    bl([N("Dach:  Absolute Herrschaft Ludwigs XIV.", { bold: true, color: NAVY })]),
    bl("Säule 1 Verwaltung/Justiz  →  Kontrolle über das Land"),
    bl("Säule 2 Adel/Hof  →  kein adeliger Widerstand"),
    bl("Säule 3 Wirtschaft  →  Geld für Hof, Beamte und Heer"),
    bl("Säule 4 Religion  →  Herrschaft von Gott gewollt, ein Glaube"),
    bl("Säule 5 Heer  →  Macht nach innen und außen durchsetzen", { after: 0 })]), sol: "Schüler formulieren die Antworten selbst, die Lehrkraft notiert Stichworte. Abschlussfrage aus dem Gruppenpuzzle: Wie sicherte Ludwig XIV. seine Macht? (Er konzentriert möglichst viel Macht bei sich und stützt sie auf Verwaltung, Hof, Wirtschaft, Religion und Heer.)" }],
]));

// Phase 3
doc.push(phase("4", "Ständegesellschaft", 10));
doc.push(...steps([
  ["gemeinsam", { say: "Hefte zu. Ihr schreibt noch nichts auf. Meldet euch.", do: ["Die leere Tabelle (Kopfzeile und Stände) ist schon an der Tafel. Unten steht die ausgefüllte Lösung.", "Fragen der Reihe nach: Wie heißen die drei Stände? Wie viele Menschen, ungefähr wie viel Prozent? Welche Bildung, welche Vorrechte, welche Pflichten? Durften die Stände mitbestimmen?", "Lücken bewusst stehen lassen und später mit Heft oder Buch klären."], board: board("Mitte · Tabelle", [stTable(true)]), sol: "Alle Angaben stammen aus den Gruppenpuzzle-Texten (Klerus, Adel, Bauern und Bürger). Die Texte nennen zur Mitbestimmung des Klerus nichts; hier ggf. ergänzen. Umrechnung der Zahlen: Klerus unter 1 %, Adel ca. 2–3 %, Dritter Stand ca. 97 %." }],
  ["abschreiben", { say: "Jetzt schreibt ihr ab: die Tabelle. Mit Lineal. Dann: Merksatz 3 von der Folie. Rahmt ihn rot ein.", board: merkBoard(3, "Die Ständegesellschaft", "In der Ständegesellschaft gehört jeder Mensch durch Geburt zu einem von drei Ständen: Klerus (ca. 150 000), Adel (ca. 500 000) und Dritter Stand (ca. 20 000 000). Klerus und Adel haben Vorrechte. Der Dritte Stand zahlt Steuern und Abgaben und ist von der politischen Mitbestimmung ausgeschlossen.  Beispiel: Adel: alleiniges Jagdrecht · Klerus: Kirchenzehnt · Dritter Stand: hohe Abgaben.  Wofür? Erklären, warum Ludwig den Adel bei Laune hält und warum der Dritte Stand die Last trägt.  Test: Wer zahlt, wer ist ausgeschlossen? Wenige haben die Vorrechte, viele tragen die Last.") }],
]));

// Phase 4
doc.push(phase("5", "Im Absolutismus fühle ich mich …", 9));
doc.push(...steps([
  ["allein", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.", board: auftrag("Ü4 Im Absolutismus fühle ich mich …", ["Schreibe Ü4 an den Rand und wähle einen Stand.", "Schreibe den Satz ab und beende ihn: „Im Absolutismus fühle ich mich …, weil …“", "Nenne mindestens ein Recht und eine Pflicht deines Standes."], [bl([N("Hilfe:  ", { bold: true }), N("privilegiert · sicher · zufrieden · ausgenutzt · benachteiligt · machtlos · überlastet")], { after: 0 })], "6 Minuten · allein · leise"), sol: "Beispiele: „… als Adeliger privilegiert, weil ich keine Steuern zahlen muss und allein jagen darf. Dafür muss ich mich ehrenhaft verhalten und Verwaltungsaufgaben übernehmen.“ · „… als Bauer ausgenutzt, weil ich Steuern, hohe Abgaben und Pachtgebühren zahlen muss und von der Mitbestimmung ausgeschlossen bin.“" }],
  ["tafel", { say: "Wer drankommt, liest vor. Alle prüfen mit.", do: ["Drei Schüler vorlesen lassen, möglichst aus drei verschiedenen Ständen.", "Die Klasse prüft: Recht und Pflicht genannt? Passen sie zum Stand? Passt das Gefühl zur Begründung? (Folie 10)", "Typische Fehler: Adel zahlt Steuern · Klerus leistet Wehrdienst (ist befreit) · Dritter Stand hat Vorrechte."], sol: "Begründung ist wichtiger als das Gefühl: Jedes Gefühl ist richtig, wenn es mit Rechten und Pflichten begründet wird." }],
  ["kontrolle", { say: "Vergleicht und verbessert mit Grün.", do: ["Wer möchte, ergänzt seine Begründung um einen Fachbegriff aus der Liste."] }],
]));

// Phase 5
doc.push(phase("6", "Lernliste und Test", 2));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Das ist eure Lernliste für den Test.", do: ["Folie 12 zeigen, Testtermin noch einmal nennen, Fragen zum Test beantworten."] }],
]));

// Hefteintrag
doc.push(h3("So sieht das Heft am Ende der Stunde aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Absolutismus und Ständegesellschaft – Wiederholung (mit Lineal unterstrichen) · Datum rechts"],
  ["Merksatz 1", "Absolutismus: Definition, Beispiel Ludwig XIV., Wofür, Test – rot umrahmt"],
  ["Ü1", "4 Zeilen: Nummer + Absolutismus/kein Absolutismus + Begründung"],
  ["Merksatz 2", "Die fünf Säulen – rot umrahmt (Lückentext liegt als Arbeitsblatt dabei, eingeklebt)"],
  ["Ü3 Tabelle", "Ständegesellschaft: Stand, Anzahl, Bildung, Vorrechte, Pflichten, Mitbestimmung"],
  ["Merksatz 3", "Die Ständegesellschaft – rot umrahmt"],
  ["Ü4", "Satz „Im Absolutismus fühle ich mich …, weil …“ – Verbesserungen grün"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? RED : NAVY, fill: LIGHT, size: 20 }), tc(b, 7906, { size: 20 })], 360))));
doc.push(p([S("Ton für die 8. Klasse: ", { bold: true }), S("Fachbegriffe konsequent nutzen (Privilegien, Mitbestimmung, Merkantilismus, Gottesgnadentum), Begründungen einfordern und die Klasse in die Rolle der Expertinnen und Experten setzen, die einschätzen, ob eine Begründung passt.")], { before: 120, after: 0 }));

(async () => {
  const d = new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 900, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: doc }],
  });
  fs.writeFileSync("../Tafelskript_Absolutismus_Wiederholung.docx", await Packer.toBuffer(d));
  console.log("ok");
})();
