// Gemeinsame Daten der Stunde „Gibs auf!“ – Stilmittel und eigene Gesprächssituationen (Deutsch 11, Do 5.11., 45 Min.)
// Text im Wortlaut des Arbeitsblatts der Mentorin (2.11.), weil die Klasse ihn im Hefter hat.
// Vom Original vermutlich abweichend – bitte mit der Quelle abgleichen (siehe ORIGINAL_HINWEIS).
const TITEL = "Gibs auf!";
const AUTOR = "Franz Kafka";
const TEXT = [
  "Es war sehr früh am Morgen, die Straßen rein und leer, ich ging zum Bahnhof.",
  "Als ich eine Turmuhr mit meiner Uhr verglich, sah ich, dass es schon viel später war, als ich geglaubt hatte, ich musste mich sehr beeilen, der Schrecken über diese Entdeckung ließ mich im Weg irrewerden, ich kannte mich in dieser Stadt noch nicht sehr gut aus.",
  "Zum Glück war ein Schutzmann in der Nähe, ich lief zu ihm und fragte ihn atemlos nach dem Weg.",
  "Er lächelte und sagte: „Von mir willst du den Weg erfahren?“ „Ja“, sagte ich, „da ich ihn selbst nicht finden kann.“ „Gibs auf, gibs auf“, sagte er und wandte sich mit einem großen Schwung ab, so wie Leute, die mit ihrem Lachen allein sein wollen.",
];
const ORIGINAL_HINWEIS = "Vermutete Abweichungen vom Original (aus dem Gedächtnis, nicht geprüft – bitte mit einer Ausgabe abgleichen): Im Original läuft der lange Satz weiter bis „… fragte ihn atemlos nach dem Weg“ („…, ich kannte mich in dieser Stadt noch nicht sehr gut aus, glücklicherweise war ein Schutzmann in der Nähe, …“); statt „irrewerden“ vermutlich „unsicher werden“; „mit einem großen Schwunge“; alte Schreibung „daß“. Für die Stunde unerheblich: Alle geübten Stilmittel stehen in beiden Fassungen.";

// Lernziele (Ich kann …), das dritte ist optional (*)
const LERNZIELE = [
  ["Ich kann in „Gibs auf!“ Stilmittel bestimmen und ihre Wirkung erklären.", false],
  ["Ich kann Bild- und Sachebene einer Parabel unterscheiden.", false],
  ["Ich kann eine eigene Erfahrung als Gespräch aufschreiben und mit dem Kommunikationsquadrat untersuchen.", true],
];
const VERLAUF = [
  ["Bild: die Geschichte nacherzählen", "Ankommen"],
  ["Wiederholung: Parabel und Stilmittel", "Vorwissen"],
  ["Neue Stilmittel", "Informieren"],
  ["Stilmittel finden · eigenes Gespräch schreiben", "Verarbeiten"],
  ["Ein Gespräch untersuchen* · Zielcheck", "Auswerten"],
];

// Wiederholung: Vorwissen-Fragen (Antworten an die Tafel)
const VORWISSEN = [
  ["Das ist die Bildebene – was ist die Sachebene?", "Parabel: Was wird übertragen?"],
  ["Wofür könnte der Schutzmann stehen?", "eine Instanz, die Orientierung geben sollte"],
  ["Welche Stilmittel kennt ihr schon?", "4 aus der letzten Stunde"],
];

// Satzbau: Satz 1 und Satz 2 im Vergleich
const SATZBAU = {
  frage: "Parataxe oder Hypotaxe?",
  s1: { text: TEXT[0], loesung: "Parataxe: drei kurze Hauptsätze, nur durch Kommas gereiht („die Straßen rein und leer“ ohne Verb). Wirkung: ruhig, nüchtern – alles scheint in Ordnung." },
  s2: { text: TEXT[1], loesung: "Beides: Am Anfang Nebensätze (Als ich …, dass es …, als ich …), danach eine lange Reihe von Hauptsätzen ohne Punkt. Wirkung: Der Satz überstürzt sich wie der Erzähler – Hektik, Schrecken, Verwirrung." },
};

// Neue Stilmittel: [Begriff, Kernsatz, Beispiel aus dem Text, Test, Wirkung im Text]
const STILMITTEL = [
  ["Wiederholung", "Ein Wort oder eine Wortgruppe wird unmittelbar wiederholt.", "„Gibs auf, gibs auf“", "Steht dasselbe direkt zweimal hintereinander?", "Die Aussage wird verstärkt und endgültig. Es klingt wie ein Abwimmeln – Widerspruch zwecklos."],
  ["Vergleich", "Zwei Dinge werden mit „wie“ oder „als ob“ verbunden.", "„so wie Leute, die mit ihrem Lachen allein sein wollen“", "Steht ein „wie“ zwischen zwei Bereichen?", "Der Schutzmann wird anschaulich: Er lacht über den Erzähler und will ihn los sein – Spott und Distanz."],
  ["Rhetorische Frage", "Eine Frage, auf die keine Antwort erwartet wird; sie ist eigentlich eine Aussage.", "„Von mir willst du den Weg erfahren?“", "Lässt sich die Frage in eine Aussage umformen?", "Gemeint ist: „Ausgerechnet von mir? Das ist lächerlich.“ Der Erzähler antwortet trotzdem mit „Ja“ – er hört nur die Sache, nicht den Spott."],
  ["Ellipse", "Ein Satz ist unvollständig, ein Wort (oft das Verb) fehlt.", "„die Straßen rein und leer“", "Muss ich ein Wort ergänzen, damit der Satz vollständig ist?", "Es fehlt „waren“. Knapp wie ein schneller Blick – leer und still, niemand außer ihm."],
];
const ALTE_STILMITTEL = [
  ["Parataxe", "Reihung von Hauptsätzen"],
  ["Hypotaxe", "Satzgefüge mit Nebensätzen"],
  ["Metapher", "bildhafter Ausdruck im übertragenen Sinn"],
  ["Symbol", "konkreter Gegenstand für etwas Abstraktes (Uhr, Weg)"],
];

// Aufgabe 1: drei Stilmittel finden (Wiederholung macht die Lehrkraft vor)
const AUFGABE1 = {
  titel: "Aufgabe 1: Stilmittel finden",
  zeit: "5 Min.", form: "allein, im Heft",
  schritte: [
    "Suche im Text einen Vergleich, eine rhetorische Frage und eine Ellipse.",
    "Schreibe jeweils das Zitat ab.",
    "Erkläre die Wirkung in einem Satz: „Dadurch wirkt …“",
  ],
};

// Aufgabe 2: eigenes Gespräch schreiben
const AUFGABE2 = {
  titel: "Aufgabe 2: Dein „Gibs auf!“",
  zeit: "10 Min.", form: "allein, im Heft",
  schritte: [
    "Denke an eine Situation aus deinem Sport, in der du Hilfe oder Orientierung gesucht hast – und abgewiesen oder nicht ernst genommen wurdest (Trainer, Kampfrichter, Mannschaft …). Du darfst die Situation auch erfinden.",
    "Schreibe sie in 6–8 Sätzen auf: Ich-Form, Präteritum.",
    "Baue einen kurzen Wortwechsel in wörtlicher Rede ein.",
    "Verwende mindestens ein Stilmittel von heute.",
  ],
};
// Schreibhilfe: nur inhaltsunabhängige Satzmuster (Grundsatz Deutsch 11)
const SCHREIBHILFE = [
  ["Wörtliche Rede", ["Er sagte: „…“", "„…“, sagte sie.", "„…?“, fragte ich.", "„…“, sagte er, „…“"]],
  ["Komma", ["Als ich …, …", "Ich hoffte, …", "…, weil …", "…, aber …"]],
  ["Stilmittel", ["… wie … (Vergleich)", "„…, …“ – zweimal dasselbe (Wiederholung)", "„Und du willst …?“ (rhetorische Frage)", "Kurz ohne Verb: „Die Halle leer.“ (Ellipse)"]],
];

// Aufgabe 3 (*): ein Gespräch untersuchen
const AUFGABE3 = {
  titel: "Aufgabe 3: Ein Gespräch untersuchen",
  zeit: "5 Min.", form: "erst allein, dann gemeinsam",
  schritte: [
    "Hört den Text zu. Schreibt den Satz auf, mit dem die Person abweist.",
    "Notiert in eurem Heft: Was ist der Appell dieses Satzes?",
    "Gemeinsam an der Tafel: alle vier Seiten – und der Vergleich mit dem Schutzmann.",
  ],
};

// Zielcheck: Stilmittel bestimmen + Wirkung
const ZIELCHECK = [
  ["„Ich lief, lief, lief – und kam doch zu spät.“", "Wiederholung", "Das Laufen wirkt endlos und vergeblich."],
  ["„Der Trainer stand da wie eine Mauer.“", "Vergleich", "Er wirkt unnahbar, man kommt nicht an ihm vorbei."],
  ["„Und du willst Kapitän sein?“", "Rhetorische Frage", "Gemeint ist: Du bist es nicht wert – Abwertung."],
];

module.exports = {
  TITEL, AUTOR, TEXT, ORIGINAL_HINWEIS, LERNZIELE, VERLAUF, VORWISSEN, SATZBAU,
  STILMITTEL, ALTE_STILMITTEL, AUFGABE1, AUFGABE2, SCHREIBHILFE, AUFGABE3, ZIELCHECK,
};
