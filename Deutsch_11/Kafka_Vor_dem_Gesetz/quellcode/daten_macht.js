// Stunde „Vor dem Gesetz“ – Macht und Ohnmacht (Neustart nach der Stunde vom 12.11., 45 Min.)
// Aufbau wie das Klausurvorbereitungsblatt der Mentorin: Einleitung → Kommunikationssituation → Stilmittel → Schreiben
const B = require("./daten"); // Text, Zeilen, zeile()

const ZIEL = "Ich kann die Einleitung einer Parabelanalyse schreiben und Bild- und Sachebene unterscheiden.";
const LEITFRAGE = "Wer hat die Macht – und wofür stehen die Figuren?";

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
  inhalt: "Ein Mann vom Lande bittet einen Türhüter um Einlass in das Gesetz, wartet sein Leben lang vergeblich und erfährt erst im Sterben, dass der Eingang nur für ihn bestimmt war.",
  hypothese: "Es lässt sich vermuten, dass Kafka die Ohnmacht des Einzelnen gegenüber einer übermächtigen, undurchschaubaren Macht thematisiert, die ihm den Zugang zu dem verwehrt, was er sich am meisten wünscht.",
  leitfragen: "Was will der Mann? Was wird dabei thematisiert?",
  geruest: "Es lässt sich vermuten, dass …",
};
const AUFGABE2 = {
  titel: "Aufgabe 2: Die Einleitung",
  schritte: [
    "Ergänze die Bausteine (TATTE).",
    "Schreibe aus den Bausteinen den TATTE-Satz (Einleitungssatz).",
    "Schreibe den Inhalt in 1–3 Sätzen auf (Präsens) – auch ein einziger Satz genügt.",
    "Formuliere die Deutungshypothese: " + EINLEITUNG.leitfragen,
  ],
};

// Aufgabe 3: Bild- und Sachebene (Figuren haben auf der Bildebene eine Rolle und stehen auf der Sachebene für etwas)
const BILD_SACH = [
  ["der Türhüter", "Wächter am Tor, kontrolliert den Zugang; „nur der unterste“ einer Hierarchie", "eine übermächtige Instanz (Behörde, Staat, Vater, Gott)", "nur der unterste Türhüter"],
  ["der Mann vom Lande", "Bittsteller, will hinein; namenlos, nur über seine Herkunft bestimmt", "der Einzelne, der Zugang sucht – jeder von uns", "kommt ein Mann vom Lande"],
  ["das Gesetz", "Gebäude mit Sälen hinter dem Tor, angeblich „jedem und immer zugänglich“", "Recht und Gerechtigkeit – offen auch Wahrheit, Sinn, Gott", "das Gesetz soll doch jedem"],
  ["das Tor / der Eingang", "steht offen, wird am Ende geschlossen; „nur für dich bestimmt“", "die persönliche Chance, die er nie nutzt", "Hier konnte niemand sonst"],
];
const MERKSATZ = [
  ["Bildebene", "Was ist die Figur in der Geschichte? → ihre Rolle"],
  ["Sachebene", "Wofür steht sie? → die Übertragung"],
  ["Symbol", "nur Gegenstände (Tor, Schemel, Pelzmantel)"],
];
const ERGEBNIS = "Auf der Bildebene sind die Rollen ungleich verteilt: Der Türhüter kontrolliert, der Mann bittet. Auf der Sachebene zeigt sich darin die Ohnmacht des Einzelnen gegenüber einer übermächtigen Instanz – die Deutungshypothese stimmt.";
const AUFGABE3 = {
  titel: "Aufgabe 3: Bild- und Sachebene",
  schritte: [
    "Lies die Bildebene: Welche Rolle hat die Figur oder das Element in der Geschichte?",
    "Ergänze die Sachebene: Wofür steht sie oder es?",
  ],
};

// Aufgabe 4: Stilmittel – Wirkung beschreiben und Bezug zur Deutungshypothese (Macht/Ohnmacht); Wiederholung macht die Lehrkraft vor
// [Stilmittel, Textstelle, Anker für die Zeile, Wirkung (Lösung), Bezug zur Hypothese (Lösung)]
const STILMITTEL = [
  ["Wiederholung", "„Ich bin mächtig“ … „einer mächtiger als der andere“", "Merke aber: Ich bin mächtig", "Die Macht wird immer wieder betont und wirkt endlos gesteigert.", "Macht: Die Instanz erscheint unüberwindbar."],
  ["Vergleich", "„teilnahmslose Fragen, wie sie große Herren stellen“", "teilnahmslose Fragen", "Der Türhüter wirkt herablassend und desinteressiert.", "Macht: Das Verhältnis ist das von Herr und Untertan."],
  ["Symbol", "der Schemel „seitwärts von der Tür“", "Der Türhüter gibt ihm einen Schemel", "Der Mann sitzt unten und am Rand, abseits vom Eingang.", "Ohnmacht: Sein Platz zeigt seine untergeordnete Stellung."],
  ["Metapher", "„Eintritt in das Gesetz“", "um Eintritt in das Gesetz", "Das Gesetz, etwas Abstraktes, wird zu einem Gebäude mit Toren, Sälen und Wächtern (Probe: Das Gesetz ist wie ein Gebäude).", "Macht und Ohnmacht: Das Recht ist nicht frei zugänglich, sondern bewacht – der Einzelne steht davor."],
];
const STIL_ERGEBNIS = "Die Stilmittel stützen die Hypothese: Wiederholung und Vergleich zeigen die Macht des Türhüters, das Symbol des Schemels die Ohnmacht des Mannes; die Metapher macht das Gesetz zu einem bewachten Ort, zu dem der Einzelne keinen Zugang hat.";
// Wiederholung der Stilmittel (Präsentation) – Beispiele bewusst nicht aus dem Text
const STIL_WIEDERHOLUNG = [
  ["Vergleich", "mit „wie“", "„Sie kämpfte wie eine Löwin.“"],
  ["Metapher", "ohne „wie“ – direkter Vergleich. Probe: Ein „wie“ lässt sich einsetzen.", "„Sie ist eine Löwin.“ → „Sie ist (wie) eine Löwin.“"],
  ["Wiederholung", "dasselbe Wort mehrmals", "„Weiter, weiter, nicht stehen bleiben!“"],
  ["Symbol", "ein Gegenstand steht für etwas Abstraktes", "die Uhr für Zeitdruck"],
];
const AUFGABE5 = {
  titel: "Aufgabe 4: Wie stützt die Sprache die Hypothese?",
  schritte: [
    "Beschreibe die Wirkung des Stilmittels.",
    "Stelle den Bezug zur Deutungshypothese her: Zeigt es Macht oder Ohnmacht?",
  ],
};


// ---------- Verlauf (45 Min.) ----------
const VERLAUF = [
  ["Wer sagt was – und wie? (markieren)", "Verarbeiten"],
  ["Die Einleitung: TATTE, Inhalt, Deutungshypothese", "Verarbeiten"],
  ["Die Szene spielen", "Verarbeiten"],
  ["Bild- und Sachebene", "Verarbeiten"],
  ["Stilmittel: Wie stützt die Sprache die Hypothese?", "Auswerten"],
  ["Zielcheck", "Auswerten"],
];

// ---------- Auswertung der Markierungen ----------
const EINGEKREIST = {
  tuerhueter: ["sagt", "lacht und sagt", "„Merke aber“", "stellt Verhöre an, fragt ihn aus", "„teilnahmslose Fragen, wie sie große Herren stellen“", "sagt immer wieder", "brüllt ihn an"],
  mann: ["bittet", "überlegt und fragt", "denkt", "ermüdet ihn durch seine Bitten", "verflucht, „rücksichtslos und laut“", "brummt nur noch", "bittet auch die Flöhe", "winkt ihm zu"],
};
const AUSWERTUNG = [
  ["Schaut aufs Blatt: Welche Farbe überwiegt? Wer spricht mehr?", "Der Türhüter – das sieht man, ohne zu lesen."],
  ["Vergleicht die eingekreisten Wörter: Wie sprechen die beiden?", "Der Türhüter bestimmt (sagt, lacht, verhört, brüllt), der Mann bittet und klagt (bittet, fragt, verflucht, brummt)."],
  ["In welcher Form wird ihre Rede wiedergegeben?", "Der Türhüter spricht in wörtlicher Rede, der Mann fast nur indirekt („fragt, ob er … werde eintreten dürfen“) – wörtlich erst kurz vor dem Tod."],
];
const AUSWERTUNG_ZUSATZ = ["Wie verändert sich der Mann?", "Er verflucht erst „laut“, dann „brummt“ er nur noch, am Ende „winkt“ er nur – der Verfall der Ohnmacht."];
const TAFEL_SPALTEN = [
  ["Türhüter", "sagt, lacht, „Merke aber“, verhört, brüllt", "wörtliche Rede", "Macht"],
  ["Mann vom Lande", "bittet, fragt, verflucht, brummt, winkt", "meist indirekte Rede", "Ohnmacht"],
];

// ---------- Spiel: Spielvorlage (nur die Sätze der Figuren; umgeformt = aus indirekter Rede) und mündliche Auswertung ----------
// [Figur, Regieanweisung, Satz, umgeformt?]
const SPIEL = [
  ["Mann", "steht vor dem Tor, bittet", "Lass mich bitte in das Gesetz eintreten.", true],
  ["Türhüter", "", "Jetzt kann ich dir den Eintritt nicht gewähren.", true],
  ["Mann", "überlegt", "Darf ich also später eintreten?", true],
  ["Türhüter", "", "Es ist möglich, jetzt aber nicht.", false],
  ["Mann", "bückt sich, um durch das Tor zu sehen", "", false],
  ["Türhüter", "lacht", "Wenn es dich so lockt, versuche es doch, trotz meines Verbotes hineinzugehn. Merke aber: Ich bin mächtig. Und ich bin nur der unterste Türhüter. Von Saal zu Saal stehn aber Türhüter, einer mächtiger als der andere. Schon den Anblick des dritten kann nicht einmal ich mehr ertragen.", false],
  ["Türhüter", "gibt ihm einen Schemel; der Mann setzt sich seitwärts", "", false],
  ["Mann", "bittet immer wieder", "Bitte, lass mich doch hinein.", true],
  ["Türhüter", "teilnahmslos, fragt ihn aus", "Woher kommst du? … Ich kann dich noch nicht einlassen.", true],
  ["Mann", "gibt ihm Geschenke", "", false],
  ["Türhüter", "nimmt alles an", "Ich nehme es nur an, damit du nicht glaubst, etwas versäumt zu haben.", false],
  ["Mann", "verflucht laut den Zufall, später brummt er nur noch vor sich hin", "", false],
  ["Mann", "wird schwach, winkt den Türhüter heran; der Türhüter beugt sich tief hinunter", "", false],
  ["Türhüter", "", "Was willst du denn jetzt noch wissen? Du bist unersättlich.", false],
  ["Mann", "", "Alle streben doch nach dem Gesetz. Wieso kommt es, dass in den vielen Jahren niemand außer mir Einlass verlangt hat?", false],
  ["Türhüter", "brüllt", "Hier konnte niemand sonst Einlass erhalten, denn dieser Eingang war nur für dich bestimmt. Ich gehe jetzt und schließe ihn.", false],
];
const SPIEL_REGIE = "Der Türhüter steht und spricht laut. Der Mann sitzt auf einem Stuhl (Schemel) und spricht leise.";
const SPIEL_FRAGEN = [
  ["Status / Gefühlslage", "Mann: unterlegen, hoffend, am Ende resigniert · Türhüter: überlegen, sicher, gleichgültig"],
  ["Absicht im Gespräch", "Mann: will Einlass, bittet, besticht · Türhüter: verweigert, ohne Gründe zu nennen"],
  ["Wie spricht er?", "Mann: leise, bittend, kurze Sätze · Türhüter: laut, bestimmend, lange Sätze"],
  ["Verhalten am Ende", "Mann: erstarrt, stellt seine einzige echte Frage · Türhüter: brüllt, schließt den Eingang"],
];


// ---------- Trainer-Beispiel nach dem Spiel: Brücke von der Bildebene zur Sachebene (Beispiel bewusst nicht aus dem Text) ----------
const TRAINER = {
  zitat: "„Wenn du meinst, dass du's besser weißt, dann lauf doch deine eigene Taktik. Aber denk dran, wer die Aufstellung macht.“",
  bank: "Am Samstag sitzt der Spieler auf der Bank. Spiel für Spiel. Der Trainer sagt nichts.",
  frage: "Was würdet ihr als Spieler tun?",
  reaktionen: ["nachfragen: „Was muss ich tun, um wieder zu spielen?“", "im Training Leistung zeigen", "mit Co-Trainer oder Sportdirektor sprechen", "den Verein wechseln", "… oder sich fügen und warten"],
  vergleich: [
    ["", "Der Spieler", "Der Mann vom Lande"],
    ["Was tut er?", "fragt nach, zeigt Leistung, sucht Hilfe", "wartet, bittet, besticht – fragt nicht nach"],
    ["Wie lange?", "Wochen, vielleicht eine Saison", "sein ganzes Leben („Tage und Jahre“)"],
    ["Danach?", "nächstes Spiel, nächste Saison – eine neue Chance", "keine: „Ich gehe jetzt und schließe ihn.“ – er kommt nicht wieder"],
  ],
  ergebnis: "Der Spieler hat Möglichkeiten und eine zweite Chance. Der Mann vom Lande hätte Möglichkeiten gehabt, nutzt sie aber nicht – und seine Chance kommt nie wieder. So wird aus Ohnmacht Resignation.",
};

// ---------- Zielcheck (im Heft) ----------
const ZIELCHECK = {
  auftrag: "Der Türhüter trägt einen Pelzmantel. Erkläre in 1–2 Sätzen, wofür der Pelzmantel steht und wie er die Deutungshypothese stützt.",
  anker: "in seinem Pelzmantel genauer ansieht",
  muster: "",
  kriterien: ["Symbol erkannt (Gegenstand für etwas Abstraktes)", "Deutung: Zeichen von Macht und Rang", "Bezug zur Hypothese mit Beleg (vgl. Z. …)"],
};
ZIELCHECK.muster = "Der Pelzmantel ist ein Symbol für die Macht und den Rang des Türhüters: Als der Mann ihn „in seinem Pelzmantel genauer ansieht“, „entschließt er sich, doch lieber zu warten“ (vgl. Z. " + parseInt(B.zeile("in seinem Pelzmantel genauer ansieht").slice(3), 10) + "–" + (parseInt(B.zeile("entschließt er sich, doch lieber").slice(3), 10) + 1) + "). Die äußeren Zeichen der Macht schüchtern ihn ein – das stützt die Hypothese von der Ohnmacht des Einzelnen.";

module.exports = Object.assign({}, B, {
  ZIEL, LEITFRAGE, MARKIERUNG, AUFGABE1, TATTE, EINLEITUNG, AUFGABE2,
  BILD_SACH, MERKSATZ, ERGEBNIS, AUFGABE3, STILMITTEL, STIL_ERGEBNIS, STIL_WIEDERHOLUNG, AUFGABE4: AUFGABE5,
  VERLAUF, EINGEKREIST, AUSWERTUNG, AUSWERTUNG_ZUSATZ, TAFEL_SPALTEN, SPIEL, SPIEL_REGIE, SPIEL_FRAGEN, TRAINER, ZIELCHECK,
});
