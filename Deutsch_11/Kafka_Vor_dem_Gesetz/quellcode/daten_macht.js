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
  ["der Türhüter", "Wächter am Tor, kontrolliert den Zugang; „nur der unterste“ einer Hierarchie", "eine übermächtige, undurchschaubare Instanz (Behörde, Staat, Vater, Gott)", "nur der unterste Türhüter"],
  ["der Mann vom Lande", "Bittsteller, will hinein; namenlos, nur über seine Herkunft bestimmt", "der Einzelne, der ihr ohnmächtig gegenübersteht – jeder von uns", "kommt ein Mann vom Lande"],
  ["das Gesetz", "Gebäude mit Sälen hinter dem Tor, angeblich „jedem und immer zugänglich“", "Recht und Gerechtigkeit – offen auch Wahrheit, Sinn, Gott: das, was er sich am meisten wünscht", "das Gesetz soll doch jedem"],
  ["das Tor / der Eingang", "steht offen, wird am Ende geschlossen; „nur für dich bestimmt“", "der Zugang, der ihm bis zuletzt verwehrt bleibt", "Hier konnte niemand sonst"],
];
const MERKSATZ = [
  ["Bildebene", "Was ist die Figur in der Geschichte? → ihre Rolle"],
  ["Sachebene", "Wofür steht sie? → die Übertragung"],
  ["Symbol", "nur Gegenstände (Tor, Schemel, Pelzmantel)"],
];
const ERGEBNIS = "Die Bild- und Sachebene bestätigt die Hypothese: Der Türhüter steht für eine übermächtige, undurchschaubare Instanz, der Mann für den Einzelnen, der ihr ohnmächtig gegenübersteht. Das Gesetz ist das, was er sich am meisten wünscht – und der Eingang bleibt ihm bis zuletzt verwehrt.";
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

// Abgleich nach Aufgabe 1 (Präsentation, 2 Folien): das Wichtigste – wie (eingekreist) und was (markiert)
const ABGLEICH = {
  tuerhueter: [
    ["sagt", "„Es ist möglich, jetzt aber nicht.“"],
    ["lacht und sagt", "„versuche es doch, trotz meines Verbotes … Merke aber: Ich bin mächtig.“"],
    ["verhört teilnahmslos", "fragt ihn über seine Heimat aus; kann ihn „noch nicht einlassen“"],
    ["sagt dabei", "„Ich nehme es nur an, damit du nicht glaubst, etwas versäumt zu haben.“"],
    ["brüllt", "„nur für dich bestimmt. Ich gehe jetzt und schließe ihn.“"],
  ],
  mann: [
    ["bittet", "um Eintritt in das Gesetz"],
    ["überlegt und fragt", "ob er später eintreten darf"],
    ["ermüdet ihn durch seine Bitten", "eingelassen zu werden"],
    ["verflucht, brummt", "den unglücklichen Zufall"],
    ["sagt – das einzige Mal wörtlich", "„wieso kommt es, dass … niemand außer mir Einlass verlangt hat?“"],
  ],
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


// ---------- Fußball-Szene nach dem Spiel (Impuls der Lehrkraft, nur im Tafelskript): Brücke von der Bildebene zur Sachebene ----------
// Ein einziges Spiel als Zeitrahmen: Abpfiff = Schließen des Eingangs, die Chance kommt nicht wieder.
const FUSSBALL = {
  szene: [
    ["", "Der Trainer hat dich zum Einwechseln geschickt. An der Wechselzone hält dich der Vierte Offizielle auf:"],
    ["Vierter Offizieller", "„Jetzt nicht. Bei der nächsten Unterbrechung vielleicht.“"],
    ["", "(Unterbrechung. Du stehst bereit.)"],
    ["Vierter Offizieller", "„Jetzt aber nicht. Und merk dir: Ich bin nur der Vierte Offizielle. Über mir stehen die Linienrichter, darüber der Schiedsrichter, darüber der Videoassistent – einer mächtiger als der andere.“"],
    ["", "(Du wartest. 60. Minute, 75. Minute. Du sprichst niemanden sonst an.)"],
    ["Spieler, 89. Minute", "„Warum will eigentlich kein anderer hier rein?“"],
    ["Vierter Offizieller", "„Diese Wechselzone war heute nur für dich offen. Abpfiff.“"],
  ],
  offen: "Offene Variante ohne Pointe: Die Szene endet nach der 75. Minute mit „Abpfiff.“ – dann entdeckt die Klasse die Parallele zum Schluss der Parabel selbst.",
  vergleich: [
    ["der Vierte Offizielle an der Wechselzone", "der Türhüter am Tor"],
    ["„Jetzt nicht. … vielleicht.“", "„Es ist möglich, jetzt aber nicht.“"],
    ["Linienrichter → Schiedsrichter → Videoassistent", "„von Saal zu Saal … einer mächtiger als der andere“"],
    ["warten an der Seitenlinie bis zur 89. Minute", "„Tage und Jahre“ auf dem Schemel"],
    ["die Frage erst kurz vor Schluss", "die Frage erst vor dem Tod"],
    ["„nur für dich offen“ – Abpfiff", "„nur für dich bestimmt … Ich gehe jetzt und schließe ihn.“"],
  ],
  fragen: [
    ["Was hätte der Spieler tun können?", "zum Schiedsrichter gehen, den Trainer rufen, einfach aufs Feld laufen"],
    ["Hat er den Schiedsrichter je gesehen?", "Nein – die Macht wird nur behauptet; er lässt sich von der Drohung abhalten, nicht von der Macht selbst."],
    ["Wie ist das beim Mann vom Lande?", "Er sieht nie mehr als den untersten Türhüter; den dritten „kann nicht einmal“ der Türhüter selbst ertragen. Er könnte es „versuchen“ – tut es aber nicht."],
    ["Wofür stehen also Türhüter und Mann?", "→ Aufgabe 3: Sachebene"],
  ],
  ergebnis: "Beide werden von einer behaupteten, unsichtbaren Hierarchie abgehalten und fragen zu spät. Beim Spieler endet es mit dem Abpfiff, beim Mann mit dem Tod – die Chance kommt nicht wieder. Aus Ohnmacht wird Resignation.",
};


// ---------- Spielvorlage mit Erzähler: Der Erzähler vermittelt das Geschehen, der Mann sagt fast nichts ----------
// Erzählertext und Figurenrede nach dem Text (gekürzt). Der Mann spricht nur seine einzige wörtliche Rede aus dem Original.
// [Rolle, Text, Regieanweisung]
const SPIEL_ERZAEHLER = [
  ["Erzähler", "Vor dem Gesetz steht ein Türhüter. Zu diesem Türhüter kommt ein Mann vom Lande und bittet um Eintritt in das Gesetz. Aber der Türhüter sagt, dass er ihm jetzt den Eintritt nicht gewähren könne. Der Mann überlegt und fragt dann, ob er also später werde eintreten dürfen.", "Türhüter steht vor der Tür"],
  ["Mann", { macht: "tritt vor den Türhüter, zeigt auf das Tor; überlegt nach dem Verbot", sagt: "leise: „Später?“" }],
  ["Türhüter", "Es ist möglich, jetzt aber nicht.", ""],
  ["Erzähler", "Da das Tor zum Gesetz offensteht wie immer und der Türhüter beiseite tritt, bückt sich der Mann, um durch das Tor in das Innere zu sehn. Als der Türhüter das merkt, lacht er und sagt:", ""],
  ["Mann", { macht: "bückt sich und späht durch das Tor; zuckt zurück, als der Türhüter lacht", sagt: "nichts" }],
  ["Türhüter", "Wenn es dich so lockt, versuche es doch, trotz meines Verbotes hineinzugehn. Merke aber: Ich bin mächtig. Und ich bin nur der unterste Türhüter. Von Saal zu Saal stehn aber Türhüter, einer mächtiger als der andere. Schon den Anblick des dritten kann nicht einmal ich mehr ertragen.", "lacht"],
  ["Erzähler", "Solche Schwierigkeiten hat der Mann vom Lande nicht erwartet; er entschließt sich, doch lieber zu warten. Der Türhüter gibt ihm einen Schemel und lässt ihn seitwärts von der Tür sich niedersetzen. Dort sitzt er Tage und Jahre. Er macht viele Versuche, eingelassen zu werden, und ermüdet den Türhüter durch seine Bitten.", "Türhüter stellt den Stuhl seitwärts hin"],
  ["Mann", { macht: "setzt sich seitwärts auf den Stuhl; steht immer wieder auf und bittet mit den Händen", sagt: "leise: „Bitte …“" }],
  ["Erzähler", "Der Türhüter stellt öfters kleine Verhöre mit ihm an – teilnahmslose Fragen, wie sie große Herren stellen – und sagt ihm immer wieder, dass er ihn noch nicht einlassen könne. Der Mann verwendet alles, um den Türhüter zu bestechen. Dieser nimmt zwar alles an, aber sagt dabei:", ""],
  ["Mann", { macht: "reicht dem Türhüter nacheinander Geschenke (z. B. Uhr, Mäppchen, Jacke)", sagt: "nichts" }],
  ["Türhüter", "Ich nehme es nur an, damit du nicht glaubst, etwas versäumt zu haben.", "nimmt alles, ohne hinzusehen"],
  ["Erzähler", "Der Mann verflucht den unglücklichen Zufall, in den ersten Jahren laut, später brummt er nur noch vor sich hin. Er wird kindisch, und er bittet sogar die Flöhe im Pelzkragen des Türhüters, ihm zu helfen. Schließlich wird sein Augenlicht schwach. Wohl aber erkennt er jetzt im Dunkel einen Glanz, der unverlöschlich aus der Türe des Gesetzes bricht.", ""],
  ["Mann", { macht: "ballt die Faust, dann murmelt er nur noch; sinkt auf dem Stuhl zusammen; starrt zur Tür", sagt: "murmelt unverständlich" }],
  ["Erzähler", "Nun lebt er nicht mehr lange. Er winkt den Türhüter heran. Der muss sich tief zu ihm hinunterneigen.", ""],
  ["Mann", { macht: "winkt den Türhüter mit letzter Kraft heran", sagt: "nichts" }],
  ["Türhüter", "Was willst du denn jetzt noch wissen? Du bist unersättlich.", "beugt sich tief herab"],
  ["Mann", { macht: "hebt den Kopf", sagt: "„Alle streben doch nach dem Gesetz. Wieso kommt es, dass in den vielen Jahren niemand außer mir Einlass verlangt hat?“" }],
  ["Erzähler", "Der Türhüter erkennt, dass der Mann schon an seinem Ende ist, und brüllt ihn an:", ""],
  ["Türhüter", "Hier konnte niemand sonst Einlass erhalten, denn dieser Eingang war nur für dich bestimmt. Ich gehe jetzt und schließe ihn.", "brüllt; geht ab"],
];
const SPIEL_ERZAEHLER_REGIE = "Drei Rollen: Erzähler (liest), Türhüter (steht, spricht laut), Mann (sitzt auf einem Stuhl seitwärts; spielt vor allem mit dem Körper und spricht fast nichts – sein einziger ganzer Satz kommt am Ende).";

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
  VERLAUF, EINGEKREIST, ABGLEICH, AUSWERTUNG, AUSWERTUNG_ZUSATZ, TAFEL_SPALTEN, SPIEL, SPIEL_REGIE, SPIEL_ERZAEHLER, SPIEL_ERZAEHLER_REGIE, SPIEL_FRAGEN, FUSSBALL, ZIELCHECK,
});
