// Podcast-Skript „Wie man ‚Vor dem Gesetz‘ knackt“ – Grundlage für einen Podcast der Lehrkraft (ca. 10 Min.), Deutsch 11
// Aufbau entlang der Stunde „Macht und Ohnmacht“: verstehen → Einleitung → Hypothese → Bild-/Sachebene → Stilmittel → Klausur-Check
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Footer, PageNumber } = require("docx");
const D = require("./daten_macht");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", GREEN = "2E9E6B";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 23, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 120, line: o.line ?? 320 } });
const regie = (x) => p(t("[" + x + "]", { italics: true, color: MUTED, size: 20 }), { after: 80 });
const teil = (zeit, titel) => p([t(zeit + "  ", { color: RED, bold: true, size: 22 }), t(titel, { font: "Cambria", bold: true, size: 28, color: NAVY })], { before: 280, after: 100, keepNext: true });
const sprich = (x) => p(t(x), { after: 140 });
const BLUE = "3F7CC4";
// Transfer: Wie erkennst du das an einer anderen Parabel? (Klausur: unbekannte Kafka-Parabel, 90 Minuten)
const transfer = (fragen) => [
  p(t("Und bei einer anderen Parabel?", { bold: true, color: BLUE }), { after: 40, keepNext: true }),
  ...fragen.map((x) => p([t("▸ ", { color: BLUE }), t(x)], { after: 40 })),
  p(t(""), { after: 80 }),
];
const merk = (x) => p([t("Merke: ", { bold: true, color: GREEN }), t(x, { bold: true })], { after: 160 });

const inhalt = [
  p(t("Podcast-Skript", { size: 20, color: MUTED, bold: true }), { after: 0 }),
  p(t("Wie man „Vor dem Gesetz“ knackt – Kafkas Parabel in fünf Schritten", { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 80 }),
  p(t("Deutsch · Klasse 11 · Länge ca. 13 Minuten · Grundlage zum Einsprechen, frei anpassbar · Ziel: Vorbereitung auf die 90-Minuten-Klausur mit einer unbekannten Kafka-Parabel", { size: 20, color: MUTED }), { after: 60 }),
  p(t("Sprechhinweis: ruhig und klar, kurze Pausen bei [Pause]. Zitate langsamer und etwas betont. Zeilenangaben nach dem Arbeitsblatt „Macht und Ohnmacht“.", { size: 20, italics: true, color: MUTED }), { after: 200 }),

  teil("0:00", "Intro"),
  regie("kurzer Jingle"),
  sprich("Hey, willkommen! Heute knacken wir einen Text, bei dem viele erst mal denken: Hä? Was soll das? – „Vor dem Gesetz“ von Franz Kafka."),
  sprich("Die gute Nachricht: Diese Parabel ist kein Rätsel, das man nur mit Glück löst. Es gibt einen Weg, wie man sie Schritt für Schritt aufschließt. Ich erkläre ihn dir am Beispiel von „Vor dem Gesetz“."),
  sprich("Aber Achtung: In der Klausur bekommst du nicht diesen Text, sondern eine andere Parabel von Kafka – und du hast 90 Minuten Zeit. Deshalb geht es heute nicht nur um diese eine Geschichte, sondern um das Werkzeug. Nach jedem Schritt frage ich dich deshalb: Wie würdest du das an einer anderen Parabel erkennen?"),
  sprich("Fünf Schritte – los geht’s."),
  regie("Pause"),

  teil("0:40", "Worum geht es überhaupt?"),
  sprich("Kurz zur Geschichte: " + D.EINLEITUNG.inhalt),
  sprich("Klingt simpel. Aber genau da fängt Kafka an, uns zu verunsichern: Warum geht der Mann nicht einfach hinein? Das Tor steht doch offen!"),

  teil("1:20", "Schritt 1: Erst verstehen, dann deuten – Wer sagt was, und wie?"),
  sprich("Der häufigste Fehler: Man will sofort deuten und hat den Text noch gar nicht richtig verstanden. Deshalb schaust du zuerst ganz genau hin: Wer sagt was – und wie sagt er es?"),
  sprich("Nimm zwei Farben. Alles, was der Türhüter sagt, in der einen Farbe. Alles, was der Mann sagt, in der anderen. Und dann kreist du die Wörter ein, die zeigen, wie gesprochen wird."),
  sprich("Was du dann siehst, ist verblüffend. Der Türhüter sagt, lacht, verhört, brüllt – und zwar in wörtlicher Rede. Zum Beispiel: „Es ist möglich, jetzt aber nicht.“ Oder: „Merke aber: Ich bin mächtig.“"),
  sprich("Und der Mann? Der bittet, fragt, verflucht, brummt. Fast alles, was er sagt, erfahren wir nur indirekt: Er „fragt, ob er also später werde eintreten dürfen“. Eine eigene Stimme bekommt er erst ganz am Ende, kurz vor seinem Tod."),
  merk("Schon beim Markieren siehst du die Machtverhältnisse: Der eine spricht, der andere bittet."),
  ...transfer([
    "Gibt es ein Gespräch? Wer spricht in wörtlicher Rede, wer wird nur indirekt wiedergegeben?",
    "Welche Verben leiten die Rede ein – sagt, befiehlt, lacht, bittet, fragt?",
    "Gibt es kein Gespräch: Wer handelt – und wer erleidet nur?",
  ]),

  teil("3:00", "Schritt 2: Die Einleitung – TATTE"),
  sprich("Jede Analyse beginnt mit einer Einleitung. Dafür gibt es eine Eselsbrücke: TATTE. Textsorte, Autor, Titel, Thema, Erscheinungsjahr."),
  sprich("Daraus baust du einen Satz: " + D.EINLEITUNG.satz1),
  sprich("Danach der Inhalt – in ein bis drei Sätzen, im Präsens. Ein Satz kann völlig reichen. Wichtig ist nur: Wer will was, was passiert, wie endet es?"),
  ...transfer([
    "TATTE funktioniert bei jedem Text gleich. Das Erscheinungsjahr findest du in der Quellenangabe unter dem Text.",
    "Für den Inhalt immer dieselben drei Fragen: Wer will was? Was passiert? Wie endet es?",
  ]),

  teil("4:00", "Schritt 3: Die Deutungshypothese – und warum sie nicht das Thema ist"),
  sprich("Jetzt kommt der wichtigste Satz deiner Einleitung: die Deutungshypothese. Und hier passiert oft ein Missverständnis."),
  sprich("Das Thema sagt, worum es geht – neutral, da sind sich alle einig. Bei uns: das Verhältnis zwischen dem Einzelnen und einer übermächtigen Instanz."),
  sprich("Die Deutungshypothese sagt, was der Text darüber aussagt. Sie ist eine Behauptung – und die musst du später beweisen."),
  sprich("Ein einfacher Test: Kann man widersprechen? Wenn nein, ist es das Thema. Wenn ja, ist es eine Hypothese."),
  sprich("Zwei Fragen helfen dir beim Formulieren: Was will der Protagonist? Und was wird dabei thematisiert? Der Mann will Einlass ins Gesetz. Und was wir in Schritt 1 gesehen haben, ist: Macht auf der einen Seite, Ohnmacht auf der anderen."),
  sprich("Daraus wird: " + D.EINLEITUNG.hypothese),
  merk("„Es lässt sich vermuten, dass …“ – eine Hypothese ist vorläufig. Sie wird im Hauptteil geprüft."),
  ...transfer([
    "Frag immer: Was will die Hauptfigur? Was steht ihr im Weg? Was wird dabei thematisiert?",
    "Bei Kafka geht es oft um Macht und Ohnmacht, um Orientierungslosigkeit oder um gescheiterte Verständigung – das ist ein Hinweis, keine Schablone. Deine Hypothese muss zum Text passen.",
  ]),

  teil("6:00", "Schritt 4: Bildebene und Sachebene – das Herz der Parabel"),
  sprich("Eine Parabel ist ein Gleichnis. Sie erzählt eine konkrete Geschichte – das ist die Bildebene. Und diese Geschichte steht für etwas Allgemeines – das ist die Sachebene. Deine Aufgabe ist die Übertragung."),
  sprich("Ein kleiner, aber wichtiger Unterschied: Figuren sind keine Symbole. Auf der Bildebene haben sie eine Rolle. Auf der Sachebene stehen sie für etwas. Symbole sind Gegenstände – wie das Tor, der Schemel oder der Pelzmantel."),
  sprich("Gehen wir es durch."),
  sprich("Der Türhüter ist auf der Bildebene ein Wächter am Tor. Er kontrolliert den Zugang – und ist nach eigener Aussage „nur der unterste“ einer ganzen Hierarchie. Auf der Sachebene steht er für eine übermächtige, undurchschaubare Instanz: eine Behörde, den Staat, den Vater, vielleicht Gott."),
  sprich("Der Mann vom Lande ist auf der Bildebene ein Bittsteller. Er hat keinen Namen, wir kennen nur seine Herkunft. Gerade deshalb kann er für jeden stehen: Auf der Sachebene ist er der Einzelne, der dieser Macht ohnmächtig gegenübersteht – jeder von uns."),
  sprich("Das Gesetz ist auf der Bildebene ein Gebäude mit Sälen hinter dem Tor, angeblich „jedem und immer zugänglich“. Auf der Sachebene steht es für das, was der Mann sich am meisten wünscht."),
  sprich("Und das Tor? Es steht offen – und wird am Ende geschlossen, mit dem Satz: „dieser Eingang war nur für dich bestimmt“. Auf der Sachebene ist das der Zugang, der ihm bis zuletzt verwehrt bleibt."),
  sprich("Übrigens: Was genau das Gesetz ist, legt Kafka nicht fest. Recht, Wahrheit, der Sinn des Lebens, vielleicht sogar Gott – alles ist möglich. Entscheidend ist, dass du deine Deutung am Text begründest."),
  regie("Pause"),
  sprich("Kleiner Vergleich aus dem Sport: Stell dir vor, der Trainer schickt dich zum Einwechseln. An der Wechselzone hält dich der Vierte Offizielle auf: „Jetzt nicht. Bei der nächsten Unterbrechung vielleicht.“ Und dann: „Ich bin nur der Vierte Offizielle. Über mir stehen die Linienrichter, darüber der Schiedsrichter, darüber der Videoassistent.“ Du wartest. 60. Minute, 75. Minute. Kurz vor Schluss fragst du endlich nach – und hörst: „Diese Wechselzone war heute nur für dich offen. Abpfiff.“"),
  sprich("Genau so funktioniert die Parabel. Eine Macht, die man nie ganz sieht. Ein Warten, das zu lange dauert. Und eine Chance, die nicht wiederkommt."),
  merk("Bildebene: Was ist die Figur in der Geschichte? Sachebene: Wofür steht sie?"),
  ...transfer([
    "Haben die Figuren keinen Namen, sondern nur eine Funktion oder Herkunft – Türhüter, Schutzmann, Mann vom Lande? Dann sind sie Stellvertreter. Frag: Wofür steht diese Rolle?",
    "Gibt es Orte oder Gegenstände, die auffällig viel Gewicht bekommen – ein Tor, ein Weg, eine Uhr, ein Stuhl? Das sind oft Symbole.",
    "Ist das Ende überraschend oder offen? Dann liegt dort meist der Schlüssel zur Sachebene.",
  ]),

  teil("8:30", "Schritt 5: Die Stilmittel stützen die Hypothese"),
  sprich("Zum Schluss zeigst du, dass auch die Sprache deine Hypothese stützt. Und zwar immer in vier Schritten: Stilmittel nennen, Textstelle zitieren, Wirkung erklären, Bezug zur Hypothese herstellen."),
  sprich("Erstens: die Wiederholung. Der Türhüter sagt „Ich bin mächtig“ – und gleich danach „einer mächtiger als der andere“. Die Macht wird immer wieder betont und wirkt unüberwindbar. Bezug: Macht."),
  sprich("Zweitens: der Vergleich. Den erkennst du am Signalwort „wie“. Der Türhüter stellt „teilnahmslose Fragen, wie sie große Herren stellen“. Er wirkt herablassend – das ist ein Verhältnis von Herr und Untertan. Bezug: Macht."),
  sprich("Drittens: die Metapher. Das ist ein Vergleich ohne „wie“. Der Mann bittet um „Eintritt in das Gesetz“. Aber in ein Gesetz kann man doch gar nicht eintreten! Probe: Das Gesetz ist wie ein Gebäude. Funktioniert – also Metapher. Das Recht wird zu einem bewachten Ort, zu dem der Einzelne keinen Zugang hat. Bezug: Macht und Ohnmacht zugleich."),
  sprich("Viertens: das Symbol. Der Türhüter gibt dem Mann einen Schemel und lässt ihn „seitwärts von der Tür“ sitzen. Unten, am Rand, abseits vom Eingang. Ein Gegenstand, der für etwas Abstraktes steht: seine untergeordnete Stellung. Bezug: Ohnmacht."),
  merk("Stilmittel – Zitat – Wirkung – Bezug zur Hypothese. Immer diese vier."),
  ...transfer([
    "Steht ein „wie“? Dann prüf, ob es ein Vergleich ist.",
    "Wird etwas Abstraktes behandelt wie ein Ort oder ein Gegenstand? Dann ist es vermutlich eine Metapher – mach die „wie“-Probe.",
    "Tauchen Wörter mehrfach auf? Wiederholung. Sind die Sätze kurz gereiht oder lang verschachtelt? Parataxe oder Hypotaxe.",
    "Und immer zurück zur Frage: Stützt das meine Hypothese?",
  ]),

  teil("10:30", "Und warum ist das kafkaesk?"),
  sprich("Ein Wort zum Schluss, das du kennst: kafkaesk. Eine übermächtige, undurchschaubare Macht. Regeln, die niemand erklärt. Ein Mensch, der ohnmächtig wartet. Und alles wird ganz nüchtern erzählt, als wäre es normal. „Vor dem Gesetz“ ist dafür das perfekte Beispiel."),

  teil("11:00", "Dein Klausur-Check: 90 Minuten, eine neue Parabel"),
  sprich("In der Klausur gehst du genau diese fünf Schritte – nur an einem Text, den du noch nicht kennst. Ein Vorschlag, wie du deine 90 Minuten einteilst:"),
  ...[["15 Minuten", "lesen und markieren: Wer sagt was, und wie?"], ["10 Minuten", "Bild- und Sachebene notieren, Hypothese formulieren, Gliederung"], ["10 Minuten", "Einleitung schreiben: TATTE-Satz, Inhalt, Hypothese"], ["40 Minuten", "Hauptteil: Bild- und Sachebene, Gespräch, Stilmittel – immer mit Beleg"], ["5 Minuten", "Schluss: Stimmt die Hypothese?"], ["10 Minuten", "überarbeiten: Präsens, Zitate, Kommas"]].map(([z, w]) => p([t(z + ": ", { bold: true }), t(w)], { after: 40 })),
  p(t(""), { after: 80 }),
  sprich("Bevor du abgibst, prüf vier Dinge: Hast du im Präsens geschrieben? Hast du deine Zitate mit Zeilenangaben belegt – zum Beispiel „vergleiche Zeile 4“? Hast du die Fachbegriffe genutzt – Parabel, Bildebene, Sachebene, die Stilmittel? Und hast du nicht nur nacherzählt, sondern die Wirkung der Sprache gedeutet?"),
  sprich("Das war’s! Fünf Schritte: verstehen, einleiten, Hypothese aufstellen, Bild- und Sachebene unterscheiden, mit den Stilmitteln belegen. Erklärt habe ich sie dir an „Vor dem Gesetz“ – aber sie funktionieren bei jeder Parabel von Kafka. Probier es gleich aus, zum Beispiel noch einmal an „Gibs auf!“: Wer sagt was? Wofür steht der Schutzmann? Welche Stilmittel stützen deine Hypothese? Bis zum nächsten Mal!"),
  regie("Jingle, Ende"),

  p(t("Hinweis für die Lehrkraft: Zitate nach dem Text auf dem Arbeitsblatt (aus dem Gedächtnis wiedergegeben, heutige Rechtschreibung) – vor der Aufnahme mit einer Ausgabe abgleichen. Zeitmarken sind Richtwerte (ca. 130 Wörter pro Minute). Die Kästen „Und bei einer anderen Parabel?“ können als Fragen an die Hörer gesprochen werden – mit kurzer Pause zum Mitdenken.", { size: 18, italics: true, color: MUTED }), { before: 300 }),
];

const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Podcast-Skript · „Vor dem Gesetz“ knacken · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 23 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 900, left: 1100, right: 1100 } } }, footers: { default: footer }, children: inhalt }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Macht_Ohnmacht/Podcast_Skript_Vor_dem_Gesetz.docx", b));
