// Lernskript für Schüler: So bearbeitest du die Probeklausur (Vorgehen, Mindestanforderungen, Musterlösung, Fehler je Aufgabe). Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, Footer, PageNumber, AlignmentType } = require("docx");
const { C } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9922, S = 22;
const line = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, thick = { style: BorderStyle.SINGLE, size: 16, color: C.dark };
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => R(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const P = (t, o = {}) => new Paragraph({ keepNext: o.keepNext, alignment: o.align, spacing: { after: o.after ?? 90, before: o.before ?? 0, line: o.line ?? 290 }, indent: o.indent, border: o.border,
  children: Array.isArray(t) ? t : runs(t, o) });
const H1 = t => P([R(t, { font: "Cambria", bold: true, size: 34 })], { before: 360, after: 120, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: C.dark, space: 4 } } });
const H2 = t => P([R(t, { font: "Cambria", bold: true, size: 26, color: C.terra })], { before: 220, after: 90, keepNext: true });
const li = t => P("–  " + t, { indent: { left: 300, hanging: 300 }, after: 60 });
const nr = (n, t) => P([R(n + "  ", { bold: true, color: C.terra, font: "Cambria" }), ...runs(t)], { indent: { left: 340, hanging: 340 }, after: 60 });
const box = (titel, items, color = C.terra) => [P([R(titel, { bold: true, color })], { keepNext: true, after: 50, indent: { left: 160 }, border: { left: { style: BorderStyle.SINGLE, size: 24, color, space: 8 } } }),
  ...items.map(t => P("✓  " + t, { indent: { left: 460, hanging: 300 }, after: 40, border: { left: { style: BorderStyle.SINGLE, size: 24, color, space: 8 } } }))];
const muster = t => P(t, { italics: true, line: 320, indent: { left: 160 }, after: 120, align: AlignmentType.JUSTIFIED, border: { left: { style: BorderStyle.SINGLE, size: 24, color: C.soft, space: 8 } } });
const fehler = items => [P([R("Typische Fehler", { bold: true, color: C.muted })], { keepNext: true, before: 100, after: 50 }), ...items.map(([f, b]) => P([R("✗  " + f + " ", { bold: true }), R("→ " + b)], { indent: { left: 300, hanging: 300 }, after: 50 }))];
function tab(cw, head, rows, size = 20) {
  const cell = (t, i, o = {}) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 110, right: 110 },
    borders: { top: line, left: line, right: line, bottom: o.head ? thick : line }, children: [P(t, { size, bold: o.bold, after: 0, line: 260 })] });
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [
    new TableRow({ tableHeader: true, children: head.map((t, i) => cell(t, i, { head: true, bold: true })) }),
    ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, i === 0 ? { bold: true } : {})) }))] });
}
const sp = () => P("", { after: 60 });

const MUSTER2 = D.muster2;

const MUSTER3 = D.kontextText + " Die Phylenreform sorgte dafür, dass in jeder Phyle Bürger aus allen drei Gebieten zusammenkamen; dadurch konnte kein Adliger mehr seine Gegend als geschlossenen Block einsetzen. Mit dem Ostrakismos konnte die Volksversammlung zudem einen Mann, der zu mächtig wurde, für zehn Jahre verbannen.";

const MUSTER4 = "Ob die Reformen des Kleisthenes der Vormachtstellung des Adels ein Ende gesetzt haben, lässt sich nur mit „teilweise“ beantworten. Für ein Ende spricht zunächst die Neuordnung der Phylen: Weil jede Phyle aus einem Bezirk der Stadt, der Küste und des Binnenlandes bestand, verloren die Adelsfamilien ihre Gegenden als geschlossene Machtblöcke. Im Rat der 500 und in der Volksversammlung entschieden nun Bürger aus ganz Attika gemeinsam, das Bürgerrecht hing an der Dorfliste und nicht mehr an der Familie, und mit dem Ostrakismos konnte das Volk zu mächtige Männer verbannen. Auch die Quelle bestätigt diese Entwicklung: Um 430 v. Chr. beklagt ein adliger Verfasser, dass in Athen die „gemeinen Leute“ herrschten (Z. 1–4) – der Adel hatte seine Vorherrschaft also verloren. Gegen ein vollständiges Ende spricht jedoch, dass Solons Vermögensklassen bestehen blieben: Archonten und damit auch der Areopag kamen weiterhin nur aus den Reichen, die Theten hatten keine Ämter, und Frauen, Metöken und Sklaven blieben ganz ausgeschlossen. Zudem führten weiterhin Adlige die Politik – Kleisthenes selbst, später Perikles. Selbst der Verfasser räumt ein, dass das Volk die militärische Führung anderen überlasse (Z. 11–13). Entscheidend geschwächt wurde der Adel erst durch Ephialtes und Perikles, die 462 v. Chr. den Areopag entmachteten und Diäten einführten. Insgesamt hat Kleisthenes die Vormachtstellung des Adels also nicht beendet, aber ihre wichtigste Grundlage – die Macht über die eigene Gegend – zerschlagen und damit den Weg zur Demokratie geöffnet.";

const children = [
  P("GESCHICHTE · KLASSE 11 · LERNSKRIPT ZUR PROBEKLAUSUR", { size: 16, bold: true, color: C.terra, after: 30 }),
  P([R("So bearbeitest du die Klausur", { font: "Cambria", bold: true, size: 40 })], { after: 40 }),
  P([R(D.titel + " · Quelle: " + D.quelleTitel, { size: 20, italics: true, color: C.muted })], { after: 160 }),
  P("Dieses Skript zeigt dir Schritt für Schritt, wie du die Probeklausur bearbeitest. Zu jeder Aufgabe findest du: **was verlangt ist**, **wie du vorgehst**, **was du mindestens brauchst, um zu bestehen**, eine **Musterlösung** und **typische Fehler**. Die Zeilenangaben beziehen sich auf die Quelle der Klausur. Nutze dazu die Übersicht „So interpretierst du eine Textquelle“."),

  // ---------- Vor dem Schreiben ----------
  H1("Vor dem Schreiben"),
  H2("Zeit einteilen (135 Minuten)"),
  tab([3400, 1500, 5022], ["Schritt", "Zeit", "Was du tust"], [
    ["Aufgaben und Quelle lesen", "15 Min.", "Aufgaben zuerst lesen, dann die Quelle zweimal: einmal im Ganzen, einmal mit Stift"],
    ["Aufgabe 1: Einleitung", "15 Min.", "Themenfrage und TATTEA"],
    ["Aufgabe 2: Position", "35 Min.", "Sinnabschnitte, indirekte Rede, Zeilen"],
    ["Aufgabe 3: Beschreiben", "30 Min.", "Organe und Bevölkerung mit Rechten"],
    ["Aufgabe 4: Urteil", "30 Min.", "Dafür – dagegen – Fazit – Werturteil"],
    ["Kontrolle", "10 Min.", "Checkliste am Ende dieses Skripts"],
  ]),
  H2("Die Operatoren verstehen"),
  tab([2400, 7522], ["Operator", "Das bedeutet"], [
    ["formulieren", "Etwas knapp und genau in eigenen Worten ausdrücken (Themenfrage, Einleitung)."],
    ["herausarbeiten", "Aus dem Text Aussagen und Haltungen erschließen, die nicht alle wörtlich dastehen – mit Belegen (Zeilen)."],
    ["beschreiben", "Sachlich und geordnet darstellen, was du weißt – ohne eigene Meinung."],
    ["beurteilen", "Ein begründetes Urteil aus damaliger Sicht fällen: Argumente dafür und dagegen abwägen und zu einem Ergebnis kommen (Sachurteil)."],
  ]),
  H2("Die Quelle markieren"),
  li("**Sinnabschnitte** mit einem Strich am Rand trennen und mit einem Stichwort überschreiben."),
  li("**Wertende Wörter** unterstreichen („gemeine Leute“, „Edle“, „Gemeinheit“, „wahnsinnige Leute“) – du brauchst sie für die Position und die Sprache."),
  li("**Zugeständnisse** markieren („muss ich zugeben“, Z. 4) – sie zeigen, dass der Verfasser nicht nur schimpft."),
  li("Am Rand notieren, was zu **Aufgabe 3 und 4** passt (z. B. Flotte, Ämter, Gerichte)."),

  // ---------- Aufgabe 1 ----------
  H1("Aufgabe 1: Themenfrage und Einleitung"),
  P("**Verlangt:** „Formulieren Sie die Themenfrage und eine Einleitung zur Quelle (formale Analyse).“"),
  H2("So gehst du vor"),
  nr(1, "Die Themenfrage steckt im Titel der Klausur: „Entmachtung des Adels?“ – formuliere daraus eine vollständige Frage."),
  nr(2, "Sammle die TATTEA-Angaben aus dem Kasten über der Quelle."),
  nr(3, "Schreibe zwei bis drei Sätze, in denen alle sechs Angaben vorkommen, und nenne zum Schluss die Themenfrage."),
  sp(),
  tab([2000, 7922], ["TATTEA", "Angabe"], D.tattea.map(([a, b]) => [a, b])),
  sp(),
  ...box("Das brauchst du mindestens", ["Themenfrage als Frage formuliert", "Titel, Autor (unbekannt), Textsorte, Zeit (430–424 v. Chr.)", "Thema in einem Satz"]),
  H2("Musterlösung"),
  muster(D.einleitung),
  ...fehler([["„Xenophon schrieb …“", "Der Autor ist unbekannt; die Schrift wurde Xenophon nur irrtümlich zugeschrieben."], ["Inhalt schon hier nacherzählen", "Die Einleitung bleibt kurz: nur die Rahmendaten und das Thema."], ["Themenfrage vergessen", "Sie gibt dem ganzen Aufsatz die Richtung – am Ende beantwortest du sie im Urteil."]]),

  // ---------- Aufgabe 2 ----------
  H1("Aufgabe 2: Position des Verfassers herausarbeiten"),
  P("**Verlangt:** „Arbeiten Sie die Positionen des Autors zur attischen Demokratie und zu deren Bürgern heraus.“"),
  H2("So gehst du vor"),
  nr(1, "Teile den Text in Sinnabschnitte (siehe Tabelle)."),
  nr(2, "Fasse jeden Abschnitt in eigenen Worten zusammen – in **indirekter Rede** („er räumt ein, das Volk rudere …“)."),
  nr(3, "Gib zu jeder Aussage die **Zeilen** an; wichtige Begriffe darfst du kurz zitieren."),
  nr(4, "Arbeite die **Gesamtposition** heraus: Wofür steht der Verfasser? Wie urteilt er über das Volk?"),
  nr(5, "Beschreibe kurz die **Sprache** und belege sie."),
  sp(),
  tab([1100, 2600, 6222], ["Zeilen", "Aspekt", "Position des Verfassers"], D.position.map(([a, b, z]) => [z, a, b])),
  sp(),
  ...box("Das brauchst du mindestens", ["Ablehnung der Demokratie und ihr Grund (Z. 1–4)", "das Zugeständnis zur Flotte (Z. 4–8)", "mindestens drei Vorwürfe ans Volk mit Zeilen (Ämter, Lasten, Gerichte, Menschenbild)", "sein Ideal: Herrschaft der Edlen (Z. 26–31)", "Gesamtposition in einem Satz: Gegner der Demokratie, Standpunkt des Adels"]),
  H2("Musterlösung"),
  muster(MUSTER2),
  ...fehler([["Nur nacherzählen, Satz für Satz", "Herausarbeiten heißt: die Haltung des Verfassers zeigen – „er kritisiert“, „er wirft vor“, „er räumt ein“."], ["Eigene Meinung einbauen („Das ist falsch, weil …“)", "Deine Meinung gehört in Aufgabe 4."], ["Ohne Zeilenangaben", "Jede Aussage braucht einen Beleg, z. B. (Z. 11–15)."], ["Das Zugeständnis übersehen", "Z. 4–8 zeigt, dass der Verfasser die Flotte anerkennt – das macht die Position genauer."]]),

  // ---------- Aufgabe 3 ----------
  H1("Aufgabe 3: Verfassung und Gesellschaft beschreiben"),
  P("**Verlangt:** „Beschreiben Sie die wichtigsten Verfassungsorgane der athenischen Demokratie unter Kleisthenes sowie die gesellschaftliche Gliederung der Bevölkerung mit ihren Rechten.“"),
  H2("So gehst du vor"),
  nr(1, "Beginne mit der Reform: zehn Phylen aus Stadt, Küste und Binnenland."),
  nr(2, "Beschreibe die Organe – jeweils: **Wer gehört dazu? Was tut es?**"),
  nr(3, "Beschreibe die Bevölkerung: Bürger nach Solons Vermögensklassen, dann die Nichtbürger – jeweils mit ihren Rechten."),
  nr(4, "Schreibe einen zusammenhängenden Text, keine Liste – und ohne Urteil."),
  sp(),
  tab([3000, 6922], ["Organ", "Zusammensetzung und Aufgaben"], D.organe),
  sp(),
  tab([3400, 6522], ["Gruppe", "Rechte"], D.gliederung),
  sp(),
  ...box("Das brauchst du mindestens", ["Volksversammlung, Rat der 500 (50 je Phyle), Archonten", "die vier Vermögensklassen mit ihren Rechten", "Frauen, Metöken und Sklaven ohne politische Rechte"]),
  H2("Musterlösung"),
  muster(MUSTER3),
  ...fehler([["Rat der 400 statt Rat der 500", "Der Rat der 400 war Solons Rat (4 × 100); Kleisthenes schuf den Rat der 500 (10 × 50)."], ["Diäten oder Entmachtung des Areopags unter Kleisthenes", "Das kam erst 462 v. Chr. mit Ephialtes und Perikles."], ["Theten konnten Ämter übernehmen", "Sie durften nur abstimmen und klagen."], ["Schon urteilen („das war ungerecht“)", "Beschreiben heißt sachlich darstellen."]]),

  // ---------- Aufgabe 4 ----------
  H1("Aufgabe 4: Beurteilen"),
  P("**Verlangt:** „Beurteilen Sie, ob die Reformen des Kleisthenes der Vormachtstellung des Adels in der attischen Polis ein Ende gesetzt haben.“"),
  H2("So gehst du vor"),
  nr(1, "**These:** Beziehe zu Beginn Stellung (ja / nein / teilweise)."),
  nr(2, "**Dafür:** Argumente mit Belegen aus Kontext und Quelle."),
  nr(3, "**Dagegen:** Argumente mit Belegen aus Kontext und Quelle."),
  nr(4, "**Abwägung und Fazit:** Was wiegt schwerer? Beantworte damit die Themenfrage."),
  nr(5, "**Werturteil (Zusatz):** Bewerte aus heutiger Sicht – mit offengelegtem Maßstab."),
  sp(),
  tab([4961, 4961], ["Dafür spricht", "Dagegen spricht"], D.urteilPro.map((p, i) => [p, D.urteilContra[i]]), 20),
  sp(),
  ...box("Das brauchst du mindestens", ["eine klare Position zur Frage", "je mindestens zwei Argumente dafür und dagegen", "mindestens ein Beleg aus der Quelle (mit Zeile)", "ein Fazit, das die Themenfrage beantwortet"]),
  H2("Musterlösung Sachurteil"),
  muster(MUSTER4),
  H2("Mögliches Werturteil"),
  muster(D.werturteil),
  P("Andere Urteile sind richtig, wenn sie klar Stellung beziehen und mit Belegen aus Quelle und Kontext begründet sind.", { size: 20, italics: true, color: C.muted }),
  ...fehler([["Nur eine Seite darstellen", "Beurteilen heißt abwägen – immer dafür und dagegen."], ["Urteil ohne Belege", "Jedes Argument braucht ein Beispiel aus Unterricht oder Quelle."], ["„Meiner Meinung nach …“ im Sachurteil", "Das Sachurteil argumentiert aus damaliger Sicht; die eigene Meinung mit Maßstab gehört ins Werturteil."], ["Die Quelle im Urteil vergessen", "Nutze sie als Beleg: Ein Adliger beklagt um 430 die Herrschaft des Volkes."]]),

  // ---------- Checkliste ----------
  H1("Checkliste vor der Abgabe"),
  ...box("Habe ich …", ["die Themenfrage formuliert und am Ende beantwortet?", "alle TATTEA-Angaben genannt und den Autor richtig (unbekannt) angegeben?", "die Position in indirekter Rede und mit Zeilen herausgearbeitet?", "das Zugeständnis und das Ideal des Verfassers genannt?", "Organe und Bevölkerung sachlich beschrieben – ohne Urteil?", "im Urteil dafür und dagegen abgewogen und ein Fazit gezogen?", "im Werturteil meinen Maßstab offengelegt?", "Absätze für jede Aufgabe gemacht und Rechtschreibung geprüft?"], C.dark),
];
const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 800, left: 992, right: 992 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Seite ", PageNumber.CURRENT], size: 16, color: "998C77" })] })] }) },
    children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
