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
const merk = (x) => p([t("Merke: ", { bold: true, color: GREEN }), t(x, { bold: true })], { after: 160 });

const inhalt = [
  p(t("Podcast-Skript", { size: 20, color: MUTED, bold: true }), { after: 0 }),
  p(t("Wie man „Vor dem Gesetz“ knackt – Kafkas Parabel in fünf Schritten", { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 80 }),
  p(t("Deutsch · Klasse 11 · Länge ca. 10 Minuten · Grundlage zum Einsprechen, frei anpassbar", { size: 20, color: MUTED }), { after: 60 }),
  p(t("Sprechhinweis: ruhig und klar, kurze Pausen bei [Pause]. Zitate langsamer und etwas betont. Zeilenangaben nach dem Arbeitsblatt „Macht und Ohnmacht“.", { size: 20, italics: true, color: MUTED }), { after: 200 }),

  teil("0:00", "Intro"),
  regie("kurzer Jingle"),
  sprich("Hey, willkommen! Heute knacken wir einen Text, bei dem viele erst mal denken: Hä? Was soll das? – „Vor dem Gesetz“ von Franz Kafka."),
  sprich("Die gute Nachricht: Diese Parabel ist kein Rätsel, das man nur mit Glück löst. Es gibt einen Weg, wie man sie Schritt für Schritt aufschließt. Und genau diesen Weg brauchst du auch in der Klausur. Fünf Schritte – los geht’s."),
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

  teil("2:40", "Schritt 2: Die Einleitung – TATTE"),
  sprich("Jede Analyse beginnt mit einer Einleitung. Dafür gibt es eine Eselsbrücke: TATTE. Textsorte, Autor, Titel, Thema, Erscheinungsjahr."),
  sprich("Daraus baust du einen Satz: " + D.EINLEITUNG.satz1),
  sprich("Danach der Inhalt – in ein bis drei Sätzen, im Präsens. Ein Satz kann völlig reichen. Wichtig ist nur: Wer will was, was passiert, wie endet es?"),

  teil("3:30", "Schritt 3: Die Deutungshypothese – und warum sie nicht das Thema ist"),
  sprich("Jetzt kommt der wichtigste Satz deiner Einleitung: die Deutungshypothese. Und hier passiert oft ein Missverständnis."),
  sprich("Das Thema sagt, worum es geht – neutral, da sind sich alle einig. Bei uns: das Verhältnis zwischen dem Einzelnen und einer übermächtigen Instanz."),
  sprich("Die Deutungshypothese sagt, was der Text darüber aussagt. Sie ist eine Behauptung – und die musst du später beweisen."),
  sprich("Ein einfacher Test: Kann man widersprechen? Wenn nein, ist es das Thema. Wenn ja, ist es eine Hypothese."),
  sprich("Zwei Fragen helfen dir beim Formulieren: Was will der Protagonist? Und was wird dabei thematisiert? Der Mann will Einlass ins Gesetz. Und was wir in Schritt 1 gesehen haben, ist: Macht auf der einen Seite, Ohnmacht auf der anderen."),
  sprich("Daraus wird: " + D.EINLEITUNG.hypothese),
  merk("„Es lässt sich vermuten, dass …“ – eine Hypothese ist vorläufig. Sie wird im Hauptteil geprüft."),

  teil("5:00", "Schritt 4: Bildebene und Sachebene – das Herz der Parabel"),
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

  teil("7:00", "Schritt 5: Die Stilmittel stützen die Hypothese"),
  sprich("Zum Schluss zeigst du, dass auch die Sprache deine Hypothese stützt. Und zwar immer in vier Schritten: Stilmittel nennen, Textstelle zitieren, Wirkung erklären, Bezug zur Hypothese herstellen."),
  sprich("Erstens: die Wiederholung. Der Türhüter sagt „Ich bin mächtig“ – und gleich danach „einer mächtiger als der andere“. Die Macht wird immer wieder betont und wirkt unüberwindbar. Bezug: Macht."),
  sprich("Zweitens: der Vergleich. Den erkennst du am Signalwort „wie“. Der Türhüter stellt „teilnahmslose Fragen, wie sie große Herren stellen“. Er wirkt herablassend – das ist ein Verhältnis von Herr und Untertan. Bezug: Macht."),
  sprich("Drittens: die Metapher. Das ist ein Vergleich ohne „wie“. Der Mann bittet um „Eintritt in das Gesetz“. Aber in ein Gesetz kann man doch gar nicht eintreten! Probe: Das Gesetz ist wie ein Gebäude. Funktioniert – also Metapher. Das Recht wird zu einem bewachten Ort, zu dem der Einzelne keinen Zugang hat. Bezug: Macht und Ohnmacht zugleich."),
  sprich("Viertens: das Symbol. Der Türhüter gibt dem Mann einen Schemel und lässt ihn „seitwärts von der Tür“ sitzen. Unten, am Rand, abseits vom Eingang. Ein Gegenstand, der für etwas Abstraktes steht: seine untergeordnete Stellung. Bezug: Ohnmacht."),
  merk("Stilmittel – Zitat – Wirkung – Bezug zur Hypothese. Immer diese vier."),

  teil("8:40", "Und warum ist das kafkaesk?"),
  sprich("Ein Wort zum Schluss, das du kennst: kafkaesk. Eine übermächtige, undurchschaubare Macht. Regeln, die niemand erklärt. Ein Mensch, der ohnmächtig wartet. Und alles wird ganz nüchtern erzählt, als wäre es normal. „Vor dem Gesetz“ ist dafür das perfekte Beispiel."),

  teil("9:10", "Dein Klausur-Check"),
  sprich("Bevor du abgibst, prüf vier Dinge: Hast du im Präsens geschrieben? Hast du deine Zitate mit Zeilenangaben belegt – zum Beispiel „vergleiche Zeile 4“? Hast du die Fachbegriffe genutzt – Parabel, Bildebene, Sachebene, die Stilmittel? Und hast du nicht nur nacherzählt, sondern die Wirkung der Sprache gedeutet?"),
  sprich("Das war’s! Fünf Schritte: verstehen, einleiten, Hypothese aufstellen, Bild- und Sachebene unterscheiden, mit den Stilmitteln belegen. Probier es gleich aus – zum Beispiel noch einmal an „Gibs auf!“. Bis zum nächsten Mal!"),
  regie("Jingle, Ende"),

  p(t("Hinweis für die Lehrkraft: Zitate nach dem Text auf dem Arbeitsblatt (aus dem Gedächtnis wiedergegeben, heutige Rechtschreibung) – vor der Aufnahme mit einer Ausgabe abgleichen. Zeitmarken sind Richtwerte (ca. 130 Wörter pro Minute).", { size: 18, italics: true, color: MUTED }), { before: 300 }),
];

const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Podcast-Skript · „Vor dem Gesetz“ knacken · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 23 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 900, left: 1100, right: 1100 } } }, footers: { default: footer }, children: inhalt }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Macht_Ohnmacht/Podcast_Skript_Vor_dem_Gesetz.docx", b));
