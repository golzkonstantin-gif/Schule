// Stunde „Vor dem Gesetz“ – Macht und Ohnmacht (Neustart nach der Stunde vom 12.11., 45 Min.)
// Aufbau wie das Klausurvorbereitungsblatt der Mentorin: Einleitung → Kommunikationssituation → Stilmittel → Schreiben
const B = require("./daten"); // Text, Zeilen, zeile()

const ZIEL = "Ich kann die Einleitung einer Parabelanalyse schreiben und erklären, warum die Kommunikation scheitert.";
const LEITFRAGE = "Wer hat die Macht – und warum kommt der Mann nicht hinein?";

// Aufgabe 1: Wer sagt was – und wie? Markierungen für die Lösung: [Textstelle, T/M = was gesagt, TW/MW = wie gesagt (einkreisen)]
const MARKIERUNG = [
  ["bittet", "MW"], ["um Eintritt in das Gesetz", "M"],
  ["sagt, dass", "TW"], ["er ihm jetzt den Eintritt nicht gewähren könne", "T"],
  ["überlegt und fragt", "MW"], ["ob er also später werde eintreten dürfen", "M"],
  ["„Es ist möglich“", "T"], ["sagt der Türhüter", "TW"], ["„jetzt aber nicht.“", "T"],
  ["lacht er und sagt", "TW"], ["„Wenn es dich so lockt, versuche es doch, trotz meines Verbotes hineinzugehn.", "T"], ["Merke aber", "TW"],
  [": Ich bin mächtig. Und ich bin nur der unterste Türhüter. Von Saal zu Saal stehn aber Türhüter, einer mächtiger als der andere. Schon den Anblick des dritten kann nicht einmal ich mehr ertragen.“", "T"],
  ["das Gesetz soll doch jedem und immer zugänglich sein", "M"], ["denkt er", "MW"],
  ["eingelassen zu werden", "M"], ["ermüdet den Türhüter durch seine Bitten", "MW"],
  ["stellt öfters kleine Verhöre mit ihm an", "TW"], ["fragt ihn über seine Heimat aus und nach vielem andern", "T"],
  ["teilnahmslose Fragen, wie sie große Herren stellen", "TW"], ["sagt er ihm immer wieder", "TW"], ["dass er ihn noch nicht einlassen könne", "T"],
  ["sagt dabei", "TW"], ["„Ich nehme es nur an, damit du nicht glaubst, etwas versäumt zu haben.“", "T"],
  ["verflucht", "MW"], ["den unglücklichen Zufall", "M"], ["rücksichtslos und laut", "MW"], ["brummt er nur noch vor sich hin", "MW"],
  ["bittet er auch die Flöhe", "MW"], ["ihm zu helfen und den Türhüter umzustimmen", "M"],
  ["winkt ihm zu", "MW"],
  ["„Was willst du denn jetzt noch wissen?“", "T"], ["fragt der Türhüter", "TW"], ["„du bist unersättlich.“", "T"],
  ["„Alle streben doch nach dem Gesetz“", "M"], ["sagt der Mann", "MW"], ["„wieso kommt es, dass in den vielen Jahren niemand außer mir Einlass verlangt hat?“", "M"],
  ["brüllt er ihn an", "TW"], ["„Hier konnte niemand sonst Einlass erhalten, denn dieser Eingang war nur für dich bestimmt. Ich gehe jetzt und schließe ihn.“", "T"],
];
const AUFGABE1 = {
  titel: "Aufgabe 1: Wer sagt was – und wie?",
  schritte: [
    "Markiere, was der Türhüter sagt, in einer Farbe – und was der Mann sagt, in einer zweiten Farbe (auch indirekte Rede: „er fragt, ob …“).",
    "Kreise ein, wie sie es sagen (z. B. sagt, bittet, lacht, brüllt).",
  ],
};

// Aufgabe 2: Einleitung (TATTE + Inhalt + Deutungshypothese)
const TATTE = [
  ["Textsorte", "Parabel"],
  ["Autor", "Franz Kafka"],
  ["Titel", "„Vor dem Gesetz“"],
  ["Thema", "das Verhältnis zwischen dem Einzelnen und einer übermächtigen Instanz"],
  ["Erscheinungsjahr", "1915"],
];
const EINLEITUNG = {
  satz1: "In der Parabel „Vor dem Gesetz“ von Franz Kafka, die 1915 veröffentlicht wurde, geht es um das Verhältnis zwischen dem Einzelnen und einer übermächtigen Instanz.",
  inhalt: "Ein Mann vom Lande bittet einen Türhüter um Einlass in das Gesetz. Der Türhüter verweigert ihm den Eintritt, und der Mann wartet sein Leben lang vergeblich vor dem offenen Tor. Erst kurz vor seinem Tod erfährt er, dass der Eingang nur für ihn bestimmt war.",
  hypothese: "Es lässt sich vermuten, dass Kafka die Ohnmacht des Einzelnen gegenüber einer übermächtigen, undurchschaubaren Macht thematisiert, die ihm den Zugang zu dem verwehrt, was er sich am meisten wünscht.",
  leitfragen: "Was will der Mann? Was wird dabei thematisiert?",
  geruest: "Es lässt sich vermuten, dass …",
};
const AUFGABE2 = {
  titel: "Aufgabe 2: Die Einleitung",
  schritte: [
    "Ergänze die Bausteine (TATTE).",
    "Fasse den Inhalt in drei Sätzen zusammen (Präsens).",
    "Formuliere die Deutungshypothese: " + EINLEITUNG.leitfragen,
  ],
};

// Aufgabe 3: Beobachtungsbogen zum Spiel (Format der Mentorin: Merkmal | Mann | Türhüter)
const BEOBACHTUNG = [
  ["Status / Gefühlslage", "unterlegen; hofft, wartet – am Ende resigniert und erschöpft", "überlegen, sicher, gleichgültig („teilnahmslose Fragen“)"],
  ["Absicht im Gespräch", "will Einlass in das Gesetz; bittet und versucht zu bestechen", "verweigert den Eintritt, ohne Gründe zu nennen; hält ihn hin"],
  ["Wie spricht er?", "bittet, fragt, verflucht, brummt – fast nur indirekt wiedergegeben", "sagt, lacht, verhört, brüllt – in wörtlicher Rede"],
  ["Verhalten am Ende", "erstarrt; stellt erst jetzt seine einzige echte Frage", "brüllt ihn an und schließt den Eingang"],
];
const AUFGABE3 = {
  titel: "Aufgabe 3: Das Gespräch im Spiel",
  schritte: ["Beobachte das Spiel und fülle die Tabelle aus."],
};

// Aufgabe 4: Warum scheitert die Kommunikation?
const GRUENDE = [
  ["Der Türhüter widerspricht sich:", "„Es ist möglich, jetzt aber nicht“; „versuche es doch, trotz meines Verbotes … Merke aber: Ich bin mächtig“ – Hoffnung und Verbot, Einladung und Drohung zugleich.", ["Es ist möglich", "Wenn es dich so lockt"]],
  ["Der Mann fragt nicht nach:", "Er bittet nur, fragt, ob er „später“ eintreten darf, und besticht. Die entscheidende Frage stellt er erst im Sterben.", ["ob er also später werde", "wieso kommt es, dass"]],
  ["Grund für beides – das Machtgefälle:", "Der Mächtige muss nichts erklären, der Ohnmächtige wagt nicht zu fragen. Die Beziehung ist komplementär (asymmetrisch).", []],
];
const ERGEBNIS = "Die Macht zeigt sich im widersprüchlichen Sprechen des Türhüters, die Ohnmacht im Schweigen und Warten des Mannes – die Deutungshypothese stimmt.";
const AUFGABE4 = {
  titel: "Aufgabe 4: Warum scheitert die Kommunikation?",
  schritte: ["Nenne drei Gründe und belege sie mit dem Text (vgl. Z. …)."],
};

// Aufgabe 5*: Stilmittel zuordnen (Format der Mentorin: Buchstaben eintragen) – [Stilmittel, Textstelle, Anker, richtige Deutung (Buchstabe), Macht/Ohnmacht]
const DEUTUNGEN = {
  A: "Das lebenslange Warten wird nüchtern und beiläufig erzählt – es wirkt ausweglos.",
  B: "Die Macht wird immer wieder betont und wirkt unüberwindbar.",
  C: "Der Platz unten und am Rand zeigt die Stellung des Wartenden ohne Macht.",
  D: "Das Gespräch wird zum Verhältnis von Herr und Untertan; echtes Interesse fehlt.",
};
const STILMITTEL = [
  ["Wiederholung", "„Ich bin mächtig“ … „einer mächtiger als der andere“", "Merke aber: Ich bin mächtig", "B", "Macht"],
  ["Vergleich", "„teilnahmslose Fragen, wie sie große Herren stellen“", "teilnahmslose Fragen", "D", "Macht"],
  ["Symbol", "der Schemel „seitwärts von der Tür“", "Der Türhüter gibt ihm einen Schemel", "C", "Ohnmacht"],
  ["Parataxe", "„Dort sitzt er Tage und Jahre. Er macht viele Versuche …“", "Dort sitzt er Tage", "A", "Ohnmacht"],
];
const STIL_ERGEBNIS = "Auch die Sprache stützt die Hypothese: Wiederholung und Vergleich zeigen die Macht des Türhüters, Symbol und Parataxe die Ohnmacht des Mannes.";
const AUFGABE5 = {
  titel: "Aufgabe 5*: Wie stützt die Sprache die Hypothese?",
  schritte: [
    "Ordne jedem Stilmittel die passende Deutung zu (Buchstabe A–D).",
    "Entscheide: Zeigt es Macht oder Ohnmacht?",
  ],
};

module.exports = Object.assign({}, B, {
  ZIEL, LEITFRAGE, MARKIERUNG, AUFGABE1, TATTE, EINLEITUNG, AUFGABE2, BEOBACHTUNG, AUFGABE3,
  GRUENDE, ERGEBNIS, AUFGABE4, DEUTUNGEN, STILMITTEL, STIL_ERGEBNIS, AUFGABE5,
});
