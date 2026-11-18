// Übungsklausur „Peisistratos – politische Fürsorge oder gezielte Machtsicherung?“ (Vorlage der Lehrkraft, Scan 18.11.), Quelle M1: Athenaion politeia, Kap. 16.
// Zeilen wie im Original (16 Zeilen), Tippfehler des Scans berichtigt.
const zeilen = [
  "Peisistratos führte die Staatsgeschäfte maßvoll und eher wie ein Bürger als wie ein",
  "Tyrann. Er war ein freundlicher und milder Mann, der auch gegenüber Menschen, die",
  "Fehler begangen hatten, Nachsicht zeigte. Den Armen gewährte er Geld für ihre",
  "landwirtschaftliche Arbeit, damit sie ihren Lebensunterhalt bestreiten konnten. Dabei",
  "verfolgte er zwei Ziele: Einerseits sollten die Menschen auf dem Land bleiben und nicht",
  "in die Stadt ziehen. Andererseits sollten sie durch ihre Arbeit einen ausreichenden",
  "Lebensunterhalt haben und weder die Zeit noch den Wunsch verspüren, sich mit den",
  "politischen Angelegenheiten zu beschäftigen.",
  "Auch sonst bemühte er sich darum, das Volk nicht zu bedrängen. Er sorgte für Frieden",
  "und Ruhe. Deshalb wurde seine Herrschaft später sogar mit dem goldenen Zeitalter",
  "des Kronos verglichen. Besonders wichtig war, dass er sich volksfreundlich und",
  "menschlich verhielt. Er wollte die Angelegenheiten des Staates nach den bestehenden",
  "Gesetzen regeln und sich keine besonderen Vorteile verschaffen. Peisistratos konnte",
  "sich lange an der Macht halten, weil er sowohl bei vielen Angehörigen des Adels als",
  "auch beim einfachen Volk Unterstützung fand. Die einen gewann er durch persönliche",
  "Beziehungen, die anderen durch seine Hilfeleistungen.",
];
const absatzEnde = [8]; // nach Zeile 8 beginnt ein neuer Absatz
const QUELLE = zeilen.join(" ");
const zeileVon = s => { const i = zeilen.findIndex(z => z.includes(s)); if (i < 0) throw new Error("nicht gefunden: " + s); return i + 1; };
const Z = (a, b) => { const x = zeileVon(a), y = zeileVon(b); return x === y ? `Z. ${x}` : y === x + 1 ? `Z. ${x}f.` : `Z. ${x}–${y}`; };
module.exports = {
  zeilen, absatzEnde, zeileVon, Z,
  themenfrage: "Peisistratos – politische Fürsorge oder gezielte Machtsicherung?",
  quelleTitel: "M1: Aristoteles über die Herrschaft des Peisistratos",
  info: "In der Schrift Athēnaion politeia (dt.: „Der Staat der Athener“) wird in einem historischen Überblick die Zeit im antiken Athen von Solon bis zur Herrschaft der Dreißig (Ende des 7. Jahrhunderts bis 403 v. Chr.) behandelt. Es ist bis heute umstritten, wer der Autor der Schrift ist. In der Antike wurde sie in der Regel dem Philosophen Aristoteles zugeschrieben.",
  nachweis: "aus: Aristoteles, Athenaion Politeia, Kapitel 16, verfasst um 330 v. Chr. Sinngemäße deutsche Übertragung, gekürzt und sprachlich vereinfacht.",
  aufgabe: [
    "das Material analysieren und die Position des Verfassers herausarbeiten,",
    "die Entstehung und Sicherung von Peisistratos’ Herrschaft erläutern,",
    "auf dieser Basis ein begründetes Urteil zur Themenfrage formulieren.",
  ],
  // Lösung TATTEA (Stichworte) und Muster-Einleitung
  tattea: [
    ["Titel", "„Athenaion politeia“ (Der Staat der Athener), Kapitel 16 (Auszug)", "Infokasten, Nachweis"],
    ["Autor", "umstritten; in der Antike dem Philosophen Aristoteles zugeschrieben", "Infokasten"],
    ["Textsorte", "Geschichtsschrift: historischer Überblick über die Verfassung Athens", "Infokasten"],
    ["Thema", "die Herrschaft des Peisistratos: wie er regierte und warum er sich lange an der Macht halten konnte", "ganzer Text"],
    ["Entstehung", "um 330 v. Chr., in Athen – rund 200 Jahre nach Peisistratos (Herrschaft 561–527 v. Chr.)", "Nachweis"],
    ["Adressat", "gebildete Leser in Athen, die sich für die Geschichte der Verfassung interessieren", "erschließen"],
  ],
  einleitung: "Bei der vorliegenden Quelle handelt es sich um einen Auszug aus der Geschichtsschrift „Athenaion politeia“ („Der Staat der Athener“), Kapitel 16. Als Verfasser gilt der Philosoph Aristoteles; ob er sie wirklich geschrieben hat, ist allerdings umstritten. Die Schrift entstand um 330 v. Chr. in Athen, also rund 200 Jahre nach der Herrschaft des Peisistratos, und richtet sich an gebildete Leser, die sich für die Geschichte der athenischen Verfassung interessieren. Sie behandelt die Herrschaft des Peisistratos, vor allem die Frage, wie er regierte und warum er sich so lange an der Macht halten konnte. Die Untersuchung geht der Frage nach, ob Peisistratos’ Politik Ausdruck politischer Fürsorge oder gezielter Machtsicherung war.",
};
