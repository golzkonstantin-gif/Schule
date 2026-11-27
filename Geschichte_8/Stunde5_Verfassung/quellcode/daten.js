// 14 Ereignisse, Lösungsreihenfolge. Daten wie in Stunde 1 bis 4 (8C und 8D gleich).
module.exports = [
  ["1715", "Ludwig XIV. stirbt und hinterlässt hohe Schulden."],
  ["1778 bis 1783", "Frankreich beteiligt sich am Krieg in Amerika, der viel Geld kostet."],
  ["1787", "Der Plan scheitert, dass auch Adel und Klerus Steuern zahlen."],
  ["1788", "Die Ernte fällt schlecht aus, das Brot wird teurer, der Staatsbankrott droht."],
  ["August 1788", "Der König kündigt die Generalstände für den Mai 1789 an."],
  ["Frühjahr 1789", "Abgeordnete werden gewählt, die Menschen schreiben ihre Klagen in Beschwerdehefte."],
  ["5. Mai 1789", "Der König eröffnet die Generalstände, wie abgestimmt wird, bleibt offen."],
  ["17. Juni 1789", "Der Dritte Stand (Bauern und Bürger) erklärt sich zur Nationalversammlung."],
  ["20. Juni 1789", "Der König lässt den Saal sperren, die Abgeordneten schwören im Ballhaus, nicht auseinanderzugehen, bis es eine Verfassung gibt."],
  ["Anfang Juli 1789", "Der König zieht Soldaten um Paris zusammen."],
  ["14. Juli 1789", "Die Menschen in Paris stürmen die Bastille, die Festung und das Gefängnis des Königs."],
  ["Sommer 1789", "Bauern greifen im ganzen Land Schlösser an."],
  ["4. August 1789", "Die Nationalversammlung schafft die Vorrechte von Klerus und Adel ab."],
  ["26. August 1789", "Die Nationalversammlung beschließt die Erklärung der Menschen- und Bürgerrechte."],
];
// Mischung für das Schnipselblatt (Position im Blatt -> Nummer der Lösung)
module.exports.MIX = [9, 3, 13, 6, 1, 11, 4, 14, 7, 2, 12, 8, 5, 10];

// Buchstabe je Ereignis (Position im gemischten Schnipselblatt, zeilenweise A bis N)
module.exports.LETTER = module.exports.MIX.reduce((o, k, i) => { o[k] = "ABCDEFGHIJKLMN"[2 * (i % 7) + (i < 7 ? 0 : 1)]; return o; }, {});
