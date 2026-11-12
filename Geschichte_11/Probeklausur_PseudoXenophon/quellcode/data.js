// Probeklausur „Die Volksherrschaft in Athen – Entmachtung des Adels?“ (Quelle: Pseudo-Xenophon, Über den Staat der Athener)
// Quellentext unverändert aus der Vorlage der Lehrkraft (Übersetzung nach: Geschichte in Quellen, Leipzig 1913).
const QUELLE = "Was die Staatsform der Athener belangt, kann ich es freilich nicht billigen, dass die gerade für diese Art der Staatsform [gemeint ist die Demokratie] sich entschieden haben. Denn hiermit haben sie sich zugleich dafür entschieden, dass es die gemeinen Leute besser haben als die Edlen. [...] Einleitend muss ich zugeben, dass die Armen und das Volk mit einigem Recht beanspruchen können, den Vorzug vor den Reichen und den Adligen zu haben, und zwar deshalb, weil nur das Volk es ist, das die Schiffe treibt und dadurch der Stadt ihre Machtstellung verschafft [...], wenigstens viel eher als das schwere Fußvolk und die Adligen und Vornehmen. Unter diesen Umständen sieht es bei vordergründiger Betrachtung so aus, als sei es nur gerecht, dass allen die Ämter offenstehen und dass es jedem von den Bürgern, wer da will, freisteht, öffentlich zu reden. Jedoch: Alle Ämter, die der Gesamtheit des Volkes Segen bringen, wenn sie in guten Händen sind, und Gefahr, wenn in schlechten, verlangt das Volk keineswegs für sich, etwa die Ämter der militärischen Führung. Alle Ämter hingegen, die dazu da sind, Sold einzutragen und Nutzen ins Haus zu bringen, um die bewirbt sich das Volk. [...] Die Aufgaben, einen Festchor aufzustellen, ein Gymnasium oder ein Schiff zu unterhalten, weist das Volk den Reichen zu. Das Volk hält es für angemessen, Geld zu verdienen, mit Singen, Rennen, Tanzen und Schiffsdienst, damit es selber den Gewinn habe, während die Reichen zugleich verarmen. Und ebenso kümmert es sich in den Gerichtshöfen nicht so sehr um das Recht als vielmehr um den eigenen Vorteil. [...] Es gilt aber auch wirklich für jedes Land, dass das bessere Element Gegner der Volksherrschaft ist; denn bei den Besseren ist Zuchtlosigkeit und Ungerechtigkeit am geringsten, gewissenhafter Eifer für das Gute und Edle am größten, beim Volke aber Mangel an Bildung und Selbstzucht am größten und Gemeinheit; denn sowohl die Armut verleitet sie viel eher zur Schlechtigkeit als auch der Mangel an Erziehung und Bildung. [...] Das Volk ist darauf aus, frei zu sein und zu herrschen – die Missordnung kümmert es wenig. Wenn du aber nach einer guten gesetzlichen Ordnung ausschaust, so wirst du zum Ersten zu sehen bekommen, dass nur rechte Männer die Gesetze geben, ferner werden die Vornehmen das gemeine Volk in Zucht halten, und die Edlen werden über den Staat beraten und werden nicht wahnsinnige Leute, zum Worte oder auch nur zu einer Versammlung zulassen.";
// feste Zeilen (gleichmäßig umbrochen), damit Zeilenangaben auf Quellenblatt und Aushängen übereinstimmen
const BREITE = 84;
function umbrechen(text, w) {
  const lines = []; let cur = "";
  for (const wort of text.split(" ")) { if ((cur + " " + wort).trim().length > w) { lines.push(cur); cur = wort; } else cur = (cur ? cur + " " : "") + wort; }
  lines.push(cur); return lines;
}
const zeilen = umbrechen(QUELLE, BREITE);
// Zeile, in der eine Textstelle beginnt bzw. endet
const zeileVon = s => { const i = QUELLE.indexOf(s); if (i < 0) throw new Error("nicht gefunden: " + s); let pos = 0;
  for (let z = 0; z < zeilen.length; z++) { pos += zeilen[z].length + 1; if (i < pos) return z + 1; } };
const Z = (von, bis) => { const a = zeileVon(von), b = zeileVon(bis); return a === b ? `Z. ${a}` : b === a + 1 ? `Z. ${a}f.` : `Z. ${a}–${b}`; };

module.exports = {
  zeilen,
  titel: "Die Volksherrschaft in Athen – Entmachtung des Adels?",
  quelleTitel: "Pseudo-Xenophon: Über den Staat der Athener",
  info: "Schrift eines unbekannten Verfassers, die irrtümlich lange Zeit dem Dichter Xenophon zugeschrieben wurde; verfasst wurde die Schrift vermutlich zwischen 430 und 424 v. Chr.",
  nachweis: "Aus: Die Pseudo-Xenophontische Schrift vom Staate der Athener. Zitiert nach Geschichte in Quellen. Leipzig 1913, S. 232 ff.",
  // Aushang 1: Themenfrage und formale Analyse (TATTEA)
  themenfrage: "Haben die Reformen des Kleisthenes die Vormachtstellung des Adels in Athen beendet?",
  tattea: [
    ["Titel", "„Über den Staat der Athener“"],
    ["Autor", "unbekannt („Pseudo-Xenophon“), lange irrtümlich Xenophon zugeschrieben; nach seiner Haltung ein Adliger, Gegner der Demokratie"],
    ["Textsorte", "politische Streitschrift"],
    ["Thema", "Kritik an der athenischen Demokratie und am „Volk“"],
    ["Entstehung", "vermutlich zwischen 430 und 424 v. Chr., wohl in Athen – zur Zeit des Perikles bzw. kurz danach"],
    ["Adressat", "gebildete, wohl adlige Leser, die die Demokratie ebenfalls ablehnen (das „du“ in Z. " + zeileVon("Wenn du aber") + ")"],
  ],
  einleitung: "Bei der vorliegenden Quelle handelt es sich um einen Auszug aus der politischen Streitschrift „Über den Staat der Athener“. Ihr Verfasser ist unbekannt; die Schrift wurde lange irrtümlich dem Dichter Xenophon zugeschrieben, weshalb man ihn „Pseudo-Xenophon“ nennt. Sie entstand vermutlich zwischen 430 und 424 v. Chr., also zur Zeit des Perikles bzw. kurz danach, und richtet sich an gebildete, wohl adlige Leser. Thema ist die Kritik des Verfassers an der athenischen Demokratie und am einfachen Volk. Die Untersuchung geht der Frage nach, ob die Reformen des Kleisthenes die Vormachtstellung des Adels in Athen beendet haben.",
  // Aushang 2 / Lehrkraft: Position des Verfassers in Stichpunkten mit Zeilen
  position: [
    ["Ablehnung der Demokratie", "billigt die Staatsform nicht, weil es dort die „gemeinen Leute“ besser haben als die „Edlen“", Z("Was die Staatsform", "als die Edlen")],
    ["Zugeständnis", "das Volk hat „mit einigem Recht“ Vorrang, weil es die Schiffe rudert und Athen seine Machtstellung verschafft – mehr als Hopliten und Adel", Z("Einleitend muss ich", "Adligen und Vornehmen")],
    ["Gleichheit nur scheinbar gerecht", "dass allen die Ämter offenstehen und jeder reden darf, wirkt nur „bei vordergründiger Betrachtung“ gerecht", Z("Unter diesen Umständen", "öffentlich zu reden")],
    ["Eigennutz bei den Ämtern", "das Volk überlässt verantwortungsvolle Ämter (militärische Führung, Strategen) anderen und will nur Ämter, die Sold bringen", Z("Jedoch: Alle Ämter", "bewirbt sich das Volk")],
    ["Lasten für die Reichen", "die Kosten für Festchor, Gymnasium und Kriegsschiff tragen die Reichen; das Volk verdient Geld, die Reichen verarmen", Z("Die Aufgaben, einen Festchor", "zugleich verarmen")],
    ["Parteiische Gerichte", "in den Gerichten geht es dem Volk um den eigenen Vorteil, nicht um das Recht", Z("Und ebenso kümmert", "eigenen Vorteil")],
    ["Menschenbild", "die „Besseren“ (Adel) seien gerecht und tugendhaft, das Volk ungebildet, zuchtlos und gemein – wegen Armut und fehlender Erziehung", Z("Es gilt aber auch", "Erziehung und Bildung")],
    ["Freiheit statt Ordnung", "das Volk will frei sein und herrschen, die „Missordnung“ ist ihm gleichgültig", Z("Das Volk ist darauf aus", "kümmert es wenig")],
    ["Ideal: Herrschaft der Edlen", "„gute gesetzliche Ordnung“: nur „rechte Männer“ geben Gesetze, die Vornehmen halten das Volk in Zucht, die Edlen beraten – das Volk bleibt von Versammlungen ausgeschlossen", Z("Wenn du aber", "Versammlung zulassen")],
  ],
  positionGesamt: "Der Verfasser ist ein entschiedener Gegner der Demokratie und vertritt den Standpunkt des Adels (oligarchisch-aristokratisch). Er erkennt zwar an, dass das Volk durch den Flottendienst Anspruch auf Mitsprache hat, wirft ihm aber Eigennutz, Unbildung und Ungerechtigkeit vor. Sein Ideal ist eine Herrschaft der „Edlen“, die das Volk von der Politik fernhält.",
  sprache: "Sprache stark wertend und abwertend: Gegensatzpaare „Edle – gemeine Leute“, „die Besseren – das Volk“; Abwertungen wie „Gemeinheit“, „Zuchtlosigkeit“, „wahnsinnige Leute“ (Z. " + zeileVon("wahnsinnige Leute") + ").",
  // Aushang 3: Verfassungsorgane unter Kleisthenes und gesellschaftliche Gliederung
  organe: [
    ["Volksversammlung (Ekklesia)", "alle männlichen Bürger ab 20 Jahren; entscheidet über Gesetze, Krieg und Frieden; seit Kleisthenes auch über den Ostrakismos"],
    ["Rat der 500", "50 Männer aus jeder der 10 Phylen, aus allen Gebieten gemischt; bereitet die Beschlüsse der Volksversammlung vor und führt die laufenden Geschäfte"],
    ["Archonten", "9 höchste Beamte, jährlich gewählt (ab 487/86 gelost); nur aus den oberen Vermögensklassen"],
    ["Areopag", "Rat der ehemaligen Archonten; Aufsicht über Beamte und Gesetze (bis 462 v. Chr.)"],
    ["Strategen", "10 Heerführer, je einer pro Phyle, jährlich gewählt (ab 501 v. Chr.)"],
    ["Volksgericht", "Geschworene aus der Bürgerschaft (seit Solon); jeder Bürger kann klagen"],
  ],
  gliederung: [
    ["1. Klasse (Fünfhundertscheffler)", "alle Rechte: Volksversammlung, Rat, alle Ämter einschließlich Archont"],
    ["2. und 3. Klasse (Ritter, Hopliten/Zeugiten)", "Volksversammlung, Gericht, Rat und weitere Ämter – nicht Archont"],
    ["4. Klasse (Theten)", "nur Volksversammlung und Gericht (abstimmen, klagen), keine Ämter"],
    ["Frauen", "keine politischen Rechte"],
    ["Metöken (Fremde)", "frei, zahlen Steuern, leisten Kriegsdienst – keine politischen Rechte"],
    ["Sklaven", "keine Rechte, gelten als Eigentum"],
  ],
  kontextText: "Nach der Reform des Kleisthenes (508/07 v. Chr.) war die Bürgerschaft in zehn Phylen gegliedert, die jeweils aus einem Bezirk der Stadt, der Küste und des Binnenlandes bestanden. Wichtigstes Organ war die Volksversammlung, in der alle männlichen Bürger ab 20 Jahren abstimmten. Ihre Beschlüsse bereitete der Rat der 500 vor, in den jede Phyle 50 Männer schickte. Die höchsten Ämter hatten die neun Archonten, die nur aus den oberen Vermögensklassen kamen; aus ehemaligen Archonten bestand der Areopag, der die Beamten beaufsichtigte. Ab 501 v. Chr. führten zehn gewählte Strategen das Heer. Die Rechte der Bürger richteten sich weiter nach Solons Vermögensklassen: Nur die erste Klasse konnte Archont werden, die Theten durften lediglich abstimmen und klagen. Frauen, Metöken und Sklaven hatten keine politischen Rechte.",
  // Aushang 4: Urteil
  urteilPro: ["Phylenreform: Die Adligen verlieren ihre Gegend als geschlossenen Machtblock.", "Rat der 500 aus allen Gebieten gemischt; Bürgerrecht über die Dorfliste statt über die Familie.", "Volksversammlung entscheidet, Ostrakismos gegen zu mächtige Männer.", "Die Quelle (um 430) zeigt: Ein Adliger beklagt, dass das „Volk“ herrscht – der Adel hat seine Vorherrschaft verloren."],
  urteilContra: ["Solons Vermögensklassen bleiben: Archonten und Areopag nur aus den Reichen.", "Theten ohne Ämter; Frauen, Metöken, Sklaven ausgeschlossen.", "Adlige führen weiter: Kleisthenes, später Perikles; laut Quelle überlässt das Volk die militärische Führung den Fähigen (Z. " + zeileVon("Jedoch: Alle Ämter") + "–" + zeileVon("militärischen Führung") + ").", "Erst Ephialtes und Perikles (462, Diäten) brechen die Macht von Areopag und Reichtum."],
  sachurteil: "Die Reformen des Kleisthenes haben die Vormachtstellung des Adels nur teilweise beendet. Einerseits brach die Neuordnung der Phylen die Macht der Adelsfamilien über ihre Gegenden: Im Rat der 500 und in der Volksversammlung entschieden nun Bürger aus allen Teilen Attikas gemeinsam, das Bürgerrecht hing am Demos und nicht mehr an der Familie. Andererseits blieben Solons Vermögensklassen bestehen; die höchsten Ämter und der Areopag waren weiterhin den Reichen vorbehalten, und Adlige wie Kleisthenes und später Perikles führten die Politik. Die Quelle zeigt, dass ein Adliger um 430 v. Chr. die Herrschaft des Volkes beklagt – die Vorherrschaft des Adels war also gebrochen, allerdings erst nach den weiteren Reformen von Ephialtes und Perikles. Kleisthenes legte dafür die entscheidende Grundlage.",
  werturteil: "Gemessen am Maßstab der Gleichheit halte ich die Reformen für einen wichtigen, aber unvollständigen Schritt: Sie gaben vielen Bürgern echten Einfluss, schlossen aber Frauen, Metöken und Sklaven weiterhin aus. Die Kritik des Verfassers überzeugt mich nicht, weil er die Rechte des Volkes nur am Nutzen für den Adel misst.",
  // ausformulierte Position des Verfassers (indirekte Rede, mit Zeilen) – Musterlösung für Skript und Zusammenfassung
  muster2: "Der Verfasser lehnt die attische Demokratie entschieden ab. Gleich zu Beginn erklärt er, er könne nicht billigen, dass sich die Athener für diese Staatsform entschieden hätten, weil es dort die „gemeinen Leute“ besser hätten als die „Edlen“ (Z. 1–4). Zwar räumt er ein, das Volk könne mit einigem Recht den Vorrang beanspruchen, da es die Schiffe rudere und Athen dadurch seine Machtstellung verschaffe (Z. 4–8). Doch sei es nur „bei vordergründiger Betrachtung“ gerecht, dass allen die Ämter offenstünden und jeder öffentlich reden dürfe (Z. 8–11). Denn das Volk überlasse verantwortungsvolle Ämter wie die militärische Führung anderen und bewerbe sich nur um Ämter, die Sold einbrächten (Z. 11–15). Die kostspieligen Aufgaben – Festchöre, Gymnasien und Kriegsschiffe – weise es den Reichen zu, während es selbst daran verdiene und die Reichen verarmten (Z. 15–19). Auch in den Gerichten gehe es dem Volk um den eigenen Vorteil statt um das Recht (Z. 19f.). Hinter dieser Kritik steht ein klares Menschenbild: Die „Besseren“ seien gerecht und tugendhaft, das Volk dagegen ungebildet, zuchtlos und gemein, was der Verfasser auf Armut und fehlende Erziehung zurückführt (Z. 20–25). Das Volk wolle vor allem frei sein und herrschen, die „Missordnung“ kümmere es wenig (Z. 25f.). Als Gegenmodell entwirft er eine „gute gesetzliche Ordnung“, in der nur „rechte Männer“ Gesetze gäben, die Vornehmen das Volk in Zucht hielten und die Edlen über den Staat berieten, während das Volk von den Versammlungen ausgeschlossen bliebe (Z. 26–31). Insgesamt vertritt der Verfasser den Standpunkt des Adels und einer Herrschaft weniger (Oligarchie). Er erkennt zwar die Bedeutung des Volkes für die Flotte an, spricht ihm aber die Fähigkeit ab, verantwortungsvoll zu regieren. Seine Sprache ist stark wertend: Gegensatzpaare wie „Edle“ und „gemeine Leute“ sowie Abwertungen wie „Gemeinheit“ oder „wahnsinnige Leute“ (Z. 30) zeigen seine Verachtung für das einfache Volk.",
  zeileVon,
};
