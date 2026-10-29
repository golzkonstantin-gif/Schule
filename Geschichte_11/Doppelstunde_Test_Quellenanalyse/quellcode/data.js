// Gemeinsame Inhalte: Quelle (Aristoteles), Schreibplan und Lösungen
// Eigene, vereinfachte Übertragung nach Aristoteles, Athenaion politeia 20,1 und 21,1–4 (gekürzt).
// Jede Zeile wird als eigene Zeile gesetzt, damit die Zeilennummern auf Blatt und Folie übereinstimmen.
const zeilen = [
  "Nach dem Sturz der Tyrannen kämpften Isagoras und Kleisthenes um die Macht.",
  "Kleisthenes stammte aus der Familie der Alkmeoniden. Weil er in den Bünden",
  "der Adligen unterlag, gewann er das Volk für sich, indem er der Menge die",
  "Macht im Staat übergab. […] Da das Volk ihm vertraute, wurde er sein Anführer.",
  "Im vierten Jahr nach dem Sturz der Tyrannen teilte er zuerst alle Bürger",
  "in zehn Phylen statt in vier ein. Er wollte sie miteinander vermischen,",
  "damit mehr Menschen an der Leitung des Staates teilhätten. Dann richtete",
  "er einen Rat der 500 statt der 400 ein, fünfzig Männer aus jeder Phyle.",
  "Das Land teilte er nach den Demen in dreißig Teile: zehn um die Stadt,",
  "zehn an der Küste und zehn im Binnenland. Diese Teile nannte er Trittyen.",
  "Durch das Los teilte er jeder Phyle drei Trittyen zu, damit jede Phyle",
  "an allen Gebieten Anteil hatte. Die Bewohner eines Demos ließ er nach",
  "ihrem Demos benennen und nicht mehr nach ihrem Vater, damit man die",
  "neuen Bürger nicht an ihrem Namen erkennen konnte.",
];
const nachweis = "Eigene, vereinfachte Übertragung nach: Aristoteles, Der Staat der Athener (Athenaion politeia) 20,1 und 21,1–4, gekürzt.";
const info = [
  "**Aristoteles** (384–322 v. Chr.) war ein griechischer Philosoph. Er lebte lange in Athen und leitete dort eine Philosophenschule, war aber kein athenischer Bürger.",
  "„Der Staat der Athener“ entstand **um 330 v. Chr.** in seiner Schule. Die Schrift beschreibt, wie sich die Verfassung Athens entwickelte. Sie richtete sich an seine Schüler und gebildete Leser.",
  "Die Reformen des Kleisthenes (508/07 v. Chr.) lagen damals **rund 180 Jahre** zurück. Aristoteles stützte sich auf ältere Schriften.",
];
// Sinnabschnitte: [Zeilen, Überschrift, Vorgabe auf dem Blatt, Lösung]
const abschnitte = [
  ["Z. 1–4", "Machtkampf nach der Tyrannis",
    "Zunächst berichtet Aristoteles, nach dem Sturz der Tyrannen hätten Isagoras und Kleisthenes um die Macht gekämpft (Z. 1). Kleisthenes habe das Volk für sich gewonnen, indem er ihm die Macht im Staat übergeben habe (Z. 2–4).",
    "Zunächst berichtet Aristoteles, nach dem Sturz der Tyrannen hätten Isagoras und Kleisthenes um die Macht gekämpft (Z. 1). Kleisthenes habe das Volk für sich gewonnen, indem er ihm die Macht im Staat übergeben habe (Z. 2–4)."],
  ["Z. 5–8", "Zehn Phylen und der Rat der 500",
    "Anschließend erklärt er, Kleisthenes habe … (Stichworte: zehn Phylen – vermischen – Rat der 500)",
    "Anschließend erklärt er, Kleisthenes habe die Bürger in zehn statt vier Phylen eingeteilt, um sie zu vermischen und mehr Menschen an der Politik zu beteiligen (Z. 5–7). Außerdem habe er einen Rat der 500 mit fünfzig Männern aus jeder Phyle geschaffen (Z. 7f.)."],
  ["Z. 9–12", "Die Trittyen",
    "Außerdem habe Kleisthenes …",
    "Außerdem habe Kleisthenes das Land in dreißig Trittyen geteilt, je zehn um die Stadt, an der Küste und im Binnenland (Z. 9f.). Jeder Phyle seien durch das Los drei Trittyen zugeteilt worden, damit sie an allen Gebieten Anteil habe (Z. 11f.)."],
  ["Z. 12–14", "", "",
    "Schließlich habe er bestimmt, dass man sich nach seinem Demos und nicht mehr nach dem Vater benennt, damit man die neuen Bürger nicht erkennen konnte (Z. 12–14)."],
];
// Themenfrage (heute vorgegeben)
const themenfrage = "Wie veränderte Kleisthenes laut Aristoteles die Ordnung Athens – und mit welchem Ziel?";
// 2a Formale Aspekte – Lösung (Reihenfolge nach der TATTE(A)-Formel)
const einleitung = {
  felder: [["Titel", "„Der Staat der Athener“ (Athenaion politeia), Auszug"], ["Autor", "Aristoteles, griechischer Philosoph (384–322 v. Chr.), lebte in Athen, kein Bürger"], ["Textsorte", "Sachtext: Schrift über die Geschichte der Verfassung Athens"], ["Thema", "die Reformen des Kleisthenes (508/07 v. Chr.)"], ["Entstehung", "um 330 v. Chr. in Athen (rund 180 Jahre nach den Reformen)"], ["Adressat", "seine Schüler und gebildete Leser"]],
  satz: "Bei der Quelle handelt es sich um einen Auszug aus der Schrift „Der Staat der Athener“ des griechischen Philosophen Aristoteles, die um 330 v. Chr. in Athen entstand. Sie richtet sich an seine Schüler und gebildete Leser und behandelt die Reformen des Kleisthenes 508/07 v. Chr.",
};
// 2b Inhaltliche Aspekte – Schlüsselbegriffe und Textsprache (Lösung)
const begriffe = "Phylen – vermischen – Rat der 500 – Trittyen – Demos";
const sprachen = ["sachlich", "emotional", "erklärend", "wertend", "auffordernd"];
const sprache = "sachlich und erklärend: Aristoteles zählt die Maßnahmen geordnet auf und nennt mit „damit …“ jeweils ihren Zweck (Z. 7, 11, 13).";
// Modell an einer anderen Quelle (wird nur vorgemacht)
const modell = {
  felder: [["Titel", "ohne Titel (Fragment)"], ["Autor", "Solon, Gesetzgeber Athens"], ["Textsorte", "Gedicht"], ["Thema", "Rechtfertigung seiner Reformen"], ["Entstehung", "um 590 v. Chr. in Athen"], ["Adressat", "die Athener"]],
  satz: "Bei der Quelle handelt es sich um ein Gedicht ohne Titel des athenischen Gesetzgebers Solon, das um 590 v. Chr. in Athen entstand. Er wendet sich an die Athener und rechtfertigt seine Reformen.",
};
module.exports = { zeilen, nachweis, info, abschnitte, themenfrage, einleitung, begriffe, sprachen, sprache, modell };
