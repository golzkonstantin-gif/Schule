const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber, PageBreak, PageOrientation } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GOLD = "FFF6E0";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 80, line: o.line }, keepNext: o.keepNext, indent: o.indent });
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("B9C6E8", 6)), margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: o.valign || VerticalAlign.TOP, columnSpan: o.span });
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, o = {}) => new TableRow({ children: cells, cantSplit: true, tableHeader: o.header });
const H1 = (s) => p(t(s, { font: "Cambria", size: 30, bold: true, color: NAVY }), { before: 240, after: 100, keepNext: true });
const H2 = (s) => p(t(s, { font: "Cambria", size: 24, bold: true, color: NAVY }), { before: 160, after: 70, keepNext: true });
const rich = (s, o = {}) => s.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((x) => x.startsWith("**") ? t(x.slice(2, -2), { size: o.size || 21, bold: true, italics: o.italics, color: o.color }) : t(x, { size: o.size || 21, italics: o.italics, color: o.color }));
const P = (s, o = {}) => p(rich(s, o), { after: o.after ?? 90, align: AlignmentType.LEFT });
const B = (arr, size = 21) => arr.map((x) => p([t("•  ", { size }), ...rich(x, { size })], { after: 50, indent: { left: 280, hanging: 0 } }));
const kv = (rows, w0 = 2500, w1 = 7700) => table([w0, w1], rows.map(([a, b]) => row([cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: w0, fill: LIGHT }), cell(Array.isArray(b) ? b : p(t(b, { size: 20 }), { after: 0 }), { w: w1 })])));
const hdr = (cols, widths) => row(cols.map((c, i) => cell(p(t(c, { bold: true, size: 19, color: "FFFFFF" }), { after: 0 }), { w: widths[i], fill: NAVY, borders: all(solid(NAVY, 8)), valign: VerticalAlign.CENTER })), { header: true });
const tc = (txt, w, o = {}) => cell(Array.isArray(txt) ? txt : p(t(txt, { size: o.size || 18, bold: o.bold }), { after: 0 }), { w, fill: o.fill });
const footer = (txt) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [txt + " · Seite ", PageNumber.CURRENT] })] })] });
const PH = (s) => t(s, { size: 20, color: RED, italics: true });

// ---------- Seite 1: Kopf ----------
const s1 = [];
s1.push(p(t("UNTERRICHTSENTWURF", { size: 20, bold: true, color: MUTED }), { after: 30 }));
s1.push(p(t("Vorrechte weg, Menschenrechte da: Die Erklärung der Menschen- und Bürgerrechte als Quelle ins Heute übersetzen", { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 120 }));
s1.push(kv([
  ["Fach / Klasse", "Geschichte · Klasse 8D (24 Schülerinnen und Schüler)"],
  ["Datum, Zeit, Raum", [p([PH("Mittwoch, 25.11.2026, [Stunde, Uhrzeit], [Raum]  (bitte ergänzen)"), t("  · 45 Minuten, Hospitation", { size: 20 })], { after: 0 })]],
  ["Lehrkraft / Hospitation", [p([PH("[Name der Lehrkraft, Hospitierende (bitte ergänzen)]")], { after: 0 })]],
  ["Reihe", "Französische Revolution (SchiC Klasse 8, Modul 2), Stunde 4 von 11"],
  ["Leitfrage der Reihe", "Wie kam es zur Französischen Revolution, und war sie ein Erfolg?"],
  ["Thema der Stunde", "Verzicht auf die Vorrechte (4. August 1789) und Erklärung der Menschen- und Bürgerrechte (26. August 1789): Was hat sich für den Dritten Stand (Bauern und Bürger) verbessert, und was davon gilt noch heute?"],
  ["Stundenfrage für die Klasse", "Was will der Dritte Stand noch? Was ändern die Abschaffung der Vorrechte und die Menschenrechte?"],
  ["Hospitationsschwerpunkt", "Aktivierung mit Fokus auf dem Kompetenzerwerb (Quelle auswerten); lehrergesteuerte Stunde mit Einzelarbeit nach dem Partnerpuzzle der Vorwoche; Sicherung am Ende"],
  ["Planungsmodell", "AVIVA (Ankommen, Vorwissen aktivieren, Informieren, Verarbeiten, Auswerten); direktes Vorgehen in Informieren und im Üben, offene Antwort in der Auswertung"],
  ["Unterlagen im Anhang", "Folien (23), Arbeitsblatt mit Lösung, Exit-Ticket, Tafelskript, Materialliste (siehe Abschnitt 10)"],
]));
s1.push(H1("1  Lehr- und Lernbedingungen"));
s1.push(H2("1.1  Lerngruppe"));
s1.push(...B([
  "24 Schülerinnen und Schüler, alle Leistungssportlerinnen und Leistungssportler (Sportschule); Beispiele aus Training, Wettkampf und Mannschaft sind anschlussfähig, werden aber nur eingesetzt, wo sie passen.",
  "Die Klasse ist laut und lebhaft. Ruhe entsteht, wenn alle beim Ankommen und in Einzel- und Partnerphasen einen Auftrag haben. Bewährt sind feste, wörtliche Ansagen („Stifte liegen. Ihr hört nur zu.“, „Flüsterstufe. Nur dein Partner hört dich.“), feste Zeiten, wenig Bewegung und wenig freies Unterrichtsgespräch. Die Lehrkraft arbeitet mit Neuer Autorität (Präsenz, Beziehung, Wir-Form).",
  "Mündliche Beiträge werden über gezielte Stichproben (3 bis 6 Personen unterschiedlichen Niveaus, unbenotet, getrennt von den mündlichen Noten) und Heft-Stichproben gesichert, nicht über freiwillige Meldungen allein.",
]));
s1.push(H2("1.2  Lernstand und Vorerfahrungen"));
s1.push(...B([
  "Stunde 1 (04.11.): Generalstände, Ständeordnung, Abstimmung nach Ständen. Die Klasse arbeitete sehr gut mit (Schüler als Presenter beim Buchstabenrätsel), die Abstimmung wurde nicht mehr geschafft und in Stunde 2 aufgegriffen.",
  "Stunde 2 (11.11., Kurzstunde): Abstimmungsspiel (Stimmen nach Ständen und nach Köpfen), Nationalversammlung, Ballhausschwur, Soldaten um Paris. Die Klasse kennt Vorrechte und die Benachteiligung des Dritten Standes (Bauern und Bürger).",
  "Stunde 3 (18.11.): Ballhausschwur und Sturm auf die Bastille als Wege des Widerstands (Partnerpuzzle in 2er-Gruppen, fünf Kriterien, Zeitstrahl, Bilanz „wirksam heißt nicht gerechtfertigt“). Ergebnis und Rückmeldung zur Stunde: ",
]));
s1.push(p([PH("[Rückmeldung zu Stunde 3 ergänzen: Zeit, Aktivierung, Lautstärke, Abgaben]")], { after: 80, indent: { left: 280 } }));
s1.push(...B([
  "Methodenkompetenz: Die Klasse hat Karikatur als Impuls, Texterschließung im Partnerpuzzle, Zeitstrahl und Stichproben kennengelernt. Quellenarbeit mit einer Textquelle (Erklärung der Menschenrechte) ist **neu** und wird in dieser Stunde modelliert (Vormachen, angeleitet, selbstständig).",
  "Schreiben: Die Klasse schreibt kurze Sätze mit Satzanfängen sicher, längere freie Texte fallen schwerer. Deshalb gibt es Satzanfänge statt fertiger Sätze, Wortspeicher mit Erklärungen und einen Vorher-Speicher.",
]));
s1.push(H2("1.3  Rahmenbedingungen"));
s1.push(...B([
  "Beamer und Dokumentenkamera nicht nötig; Folien mit Schrift ab 27 Punkt, Phasenanzeige A V I V A. Tafel (Seitentafel) mit Stundenfrage, Zielen und Ablauf vor Stundenbeginn beschriftet, Mitte frei.",
  "Material liegt vorab umgedreht auf den Tischen (Arbeitsblatt), das Exit-Ticket wird erst in der Auswertung ausgegeben, damit kein Austeilen Zeit kostet.",
  "Unbenotete Stunde: Das Arbeitsblatt ist freiwillig abzugeben, der Zettel (Exit-Ticket) wird abgegeben. Punkte oder Noten gibt es nicht (später möglich).",
]));
s1.push(H1("2  Sachanalyse"));
s1.push(P("Nach dem Sturm auf die Bastille (14. Juli 1789) greifen im Sommer 1789 Bauern im ganzen Land Schlösser an; die Nationalversammlung beschließt in der Nacht zum 4. August die Abschaffung der Vorrechte (unter anderem Steuerfreiheit von Klerus und Adel, Abgaben und Dienste der Bauern an den Adel, Vorrechte bei Ämtern und Jagd). Die Aufhebung ist für die Bauern teils mit Ablösezahlungen verbunden; sie führt nicht sofort zu gleichen Lebensverhältnissen. Klerus und Adel geben nicht freiwillig nach, sondern unter dem Druck der Aufstände. Am 26. August 1789 beschließt die Nationalversammlung die Erklärung der Menschen- und Bürgerrechte (Vorbild: amerikanische Unabhängigkeitserklärung von 1776)."));
s1.push(P("Für die Stunde sind acht Artikel der Erklärung ausgewählt (gekürzte und vereinfachte Fassung, Nummerierung wie im Original): Artikel 1 (frei und gleich an Rechten), 2 (Grundrechte: Freiheit, Eigentum, Sicherheit, Widerstand gegen Unterdrückung), 3 (Ursprung der Herrschaft beim Volk, Volkssouveränität), 4 (Freiheit endet, wo sie anderen schadet), 6 (Gesetz für alle gleich, Ämter nach Fähigkeiten), 10 (Religionsfreiheit), 11 (Meinungsfreiheit), 17 (Eigentum). Auf dem Arbeitsblatt arbeiten die Schüler an den Artikeln 1, 3, 6, 11 und 17; die Artikel 2, 4 und 10 werden kurz im Plenum besprochen."));
s1.push(P("Die Erklärung gilt 1789 in der Praxis nicht für alle (Frauen, Besitzlose, Sklaven in den Kolonien); das ist Thema der Verfassung von 1791 (Aktiv- und Passivbürger) in Stunde 5 und wird in der Stunde nur als Ausblick angesprochen. Der Transfer in die Gegenwart erfolgt über drei Artikel des Grundgesetzes (Art. 3 Abs. 1: „Alle Menschen sind vor dem Gesetz gleich.“, Art. 14 Abs. 1: Eigentum und Erbrecht werden gewährleistet, Art. 20 Abs. 2: „Alle Staatsgewalt geht vom Volke aus.“); zum Kreuzen auf dem Exit-Ticket kommt Art. 5 (Meinungsfreiheit) hinzu. Die Formulierungen des Grundgesetzes sind auf den Folien leicht gekürzt und vor Gebrauch mit dem Original abzugleichen."));
s1.push(P("Vereinfachungen und Deutungen, die im Unterricht offen benannt werden: (1) Die Zuordnung „Vorher“ (Zensur, Abgaben, Steuerfreiheit) ist verkürzt. (2) Die Karikatur ist selbst gezeichnet und keine historische Quelle; die Aussage der Figuren („Wir kamen euch entgegen“) ist eine bewusst überspitzte Selbstdarstellung, historisch war der Verzicht nicht freiwillig.", { italics: true }));

// ---------- 3 Didaktische Analyse ----------
s1.push(H1("3  Didaktische Analyse"));
s1.push(H2("3.1  Bedeutung und Relevanz"));
s1.push(...B([
  "**Gegenwartsbezug:** Die Rechte, die die Klasse als selbstverständlich erlebt (Gleichheit vor dem Gesetz, Meinungsfreiheit, Eigentum), sind historisch erkämpft. Der Aha-Moment („1789 oder Grundgesetz?“) zeigt die Kontinuität und die Vorbildwirkung der Erklärung für das Grundgesetz.",
  "**Zukunftsbezug:** Die Kompetenz, eine Quelle in eigenen Worten wiederzugeben, in heutige Begriffe zu übersetzen und mit der Zeit davor zu vergleichen, trägt die Reihe (Verfassungsschaubild, Urteil) und den Test.",
  "**Lebenswelt:** Gleiche Regeln für alle und Leistung statt Herkunft sind für Leistungssportlerinnen und -sportler nachvollziehbar (Artikel 6: Ämter nach Fähigkeiten); der Bezug wird angeboten, nicht erzwungen.",
]));
s1.push(H2("3.2  Exemplarität und Zugänglichkeit"));
s1.push(...B([
  "Die Erklärung ist die zentrale Quelle zur Bedeutung der Revolution und im SchiC als Inhalt des Moduls „Französische Revolution“ verankert (Bürger- und Menschenrechtserklärung). Die Quellenauswertung ergänzt die im SchiC genannten Methoden (Karikatur, Verfassungsschaubild).",
  "Zugänglichkeit: Die Quelle ist gekürzt und vereinfacht, die Begriffe werden mit kurzen Erklärungen angeboten (Wortspeicher), das Vorwissen zur Zeit vor 1789 steht als Vorher-Speicher auf dem Blatt, weil die Zensur nicht behandelt wurde. Die Klasse muss nicht raten, sie wählt aus.",
  "Das Einstiegsbild (Karikatur mit zwei Wandbildern von Ballhaus und Bastille) knüpft an Stunde 3 an und macht die Frage „Was wollt ihr noch?“ zur Stundenfrage.",
]));
s1.push(H2("3.3  Kompetenzen (SchiC Klasse 8, Modul 2)"));
s1.push(table([2400, 7800], [
  hdr(["Kompetenzbereich", "Beitrag dieser Stunde"], [2400, 7800]),
  row([tc("Methodenkompetenz", 2400, { bold: true, fill: LIGHT }), tc("Eine Textquelle in drei Schritten auswerten: in eigenen Worten wiedergeben, den heutigen Begriff zuordnen, das Vorher benennen (Vormachen, angeleitet mit Hilfe, selbstständig).", 7800)]),
  row([tc("Rekonstruktionskompetenz", 2400, { bold: true, fill: LIGHT }), tc("Zusammenhänge zwischen Ereignissen (Ballhausschwur, Bastille, Aufstände, 4. August, 26. August) herstellen; Ursache und Folge des Entgegenkommens.", 7800)]),
  row([tc("Urteilskompetenz (anbahnend)", 2400, { bold: true, fill: LIGHT }), tc("Die Bedeutung der Erklärung erkennen („Was hat sich verbessert?“, „Was bedeutet 1789 für unser Grundgesetz?“) als Vorstufe zum Sach- und Werturteil; eine ausgebaute Urteilsaufgabe folgt ab Stunde 8.", 7800)]),
]));

// ---------- 4 Sequenz ----------
s1.push(H1("4  Einordnung in die Reihe (angepasste Sequenz)"));
s1.push(P("Die Reihe umfasst für die 8D elf wöchentliche Stunden (Mittwoch) bis zum Halbjahresende, ohne Reserve, mit zwei Tests. Die Sequenz wurde am 17.11. angepasst (Hospitation in Stunde 4, Test 1 verschoben, Napoleon und Ende der Revolution zusammengelegt)."));
const SQ = [
  ["1", "04.11.", "Generalstände, Abstimmung nach Ständen", "gelaufen"],
  ["2", "11.11.", "Dritter Stand handelt: Nationalversammlung, Ballhausschwur, Abstimmungsspiel (Kurzstunde)", "gelaufen"],
  ["3", "18.11.", "Wege des Widerstands: Ballhausschwur und Bastille (Partnerpuzzle, Urteil)", "gelaufen"],
  ["4", "25.11.", "Vorrechte weg, Menschenrechte da (Quelle ins Heute übersetzen), HOSPITATION", "diese Stunde"],
  ["5", "02.12.", "Verfassung 1791: konstitutionelle Monarchie, Aktiv- und Passivbürger (Einstieg: Der König unterschreibt die Erklärung erst nach dem Marsch der Frauen im Oktober 1789)", "geplant"],
  ["6", "09.12.", "Test 1 (Stunden 1 bis 5)", "geplant"],
  ["7", "16.12.", "Vom König zur Republik: Flucht, Krieg, Hinrichtung", "geplant"],
  ["8", "06.01.", "Jakobinerherrschaft und Terror; Sach- und Werturteil (Modell)", "geplant"],
  ["9", "13.01.", "Ende der Revolution und Napoleon; Urteil angeleitet", "geplant"],
  ["10", "20.01.", "Test 2 (Stunden 7 bis 9) mit Urteilsaufgabe", "geplant"],
  ["11", "27.01.", "Gesamturteil zur Revolution (Vergleich mit der Eingangsdiagnose); Wiener Kongress nur als Zusatz", "geplant"],
];
s1.push(table([700, 1100, 6900, 1500], [hdr(["Nr.", "Datum", "Thema", "Stand"], [700, 1100, 6900, 1500]), ...SQ.map((r, i) => row(r.map((v, j) => tc(v, [700, 1100, 6900, 1500][j], { bold: i === 3, fill: i === 3 ? GOLD : undefined })))) ]));
s1.push(P(""));
s1.push(...B([
  "**Vorstunde (3):** Widerstand von Ballhausschwur und Bastille; die Rückschau an den zwei Bildern greift sie auf. **Folgestunde (5):** Wem gelten die Rechte 1789 wirklich? (Verfassung 1791). Die Frage am Ende der Stunde 4 („Für wen galten diese Rechte 1789 wirklich?“) ist das Gelenk.",
  "**Methodenaufbau:** Karikatur (Stunde 2), Quellenauswertung (Stunde 4: Modell, angeleitet, selbstständig), Verfassungsschaubild (Stunde 5), Ursache-Folge-Kette (6, 7), Sach- und Werturteil (8, 9), Test (10), Gesamturteil (11).",
  "**Parallelklasse 8C:** Die 8C hat die Menschenrechte bereits am 13.11. behandelt (Epochenjahr 1789 wiederholt, Arbeitsblatt zu den Artikeln 1, 2, 3, 4, 6, 10, 11); Artikel und Begriffe sind abgestimmt (Vergleichbarkeit für den Test), ergänzt um das Eigentumsrecht (Artikel 17).",
]));
const sec1 = { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer("Unterrichtsentwurf Geschichte 8D · 25.11.2026") }, children: s1 };

// ---------- 5 bis 7 ----------
const s2 = [];
s2.push(H1("5  Lernziele und Prüfaufgabe"));
s2.push(table([1300, 4300, 4600], [
  hdr(["Ziel", "Ich kann …", "Woran erkenne ich es?"], [1300, 4300, 4600]),
  row([tc("1 (Mindest)", 1300, { bold: true, fill: LIGHT }), tc("erklären, was sich durch den Verzicht auf die Vorrechte und die Menschenrechte für den Dritten Stand verbessert hat.", 4300), tc("Bilanz „Was hat sich verbessert?“ (Folie 15); Exit-Ticket: Satz „Wir wollen Menschenrechte, damit wir endlich …“.", 4600)]),
  row([tc("2 (Mindest)", 1300, { bold: true, fill: LIGHT }), tc("Artikel von 1789 in heutige Begriffe übersetzen (zum Beispiel Eigentumsrecht).", 4300), tc("Arbeitsblatt: Artikel 6, 11 und 17 in den drei Schritten (in meinen Worten, Begriff ankreuzen, vorher), Stichprobe im Heft.", 4600)]),
  row([tc("3 (*)", 1300, { bold: true, fill: LIGHT }), tc("an einem Grundgesetz-Artikel zeigen, dass die Erklärung bis heute nachwirkt.", 4300), tc("„1789 oder Grundgesetz?“ (Heft, Folien 16 bis 19); Exit-Ticket: Kreuz beim passenden Grundgesetz-Artikel.", 4600)]),
]));
s2.push(P(""));
s2.push(P("Die Stundenziele stehen als „Ich kann …“-Sätze zu Beginn an der Seitentafel und auf Folie 3. Das Mindestziel sind Ziel 1 und 2; Ziel 3 trägt den Stern und wird bei Zeitnot verkürzt. Die Phase „Auswerten“ wird nie gestrichen. Die Rückmeldung an die Klasse erfolgt nach drei Stufen, die den Zielen entsprechen (Stufe 1 bis 3); es gibt keine Selbsteinschätzung der Schülerinnen und Schüler."));

s2.push(H1("6  Didaktisch-methodisches Konzept"));
s2.push(H2("6.1  Struktur nach AVIVA"));
s2.push(table([1700, 1100, 7400], [
  hdr(["Phase", "Zeit", "Didaktische Funktion und Begründung"], [1700, 1100, 7400]),
  row([tc("Ankommen", 1700, { bold: true, fill: LIGHT }), tc("4 Min.", 1100), tc("Die Karikatur (Klerus und Adel: „Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?“) wird zur Stundenfrage. Die Vermutung im Heft („Was will der Dritte Stand noch? Ich vermute …“) ist die Eingangsdiagnose und wird am Ende mit der Antwort verglichen (Vorher-Nachher). Alle haben einen Auftrag; Ziel und Ablauf werden transparent gemacht (indirekt/offen für die Vermutung, direkt für Ziel und Ablauf).", 7400)]),
  row([tc("Vorwissen aktivieren", 1700, { bold: true, fill: LIGHT }), tc("4 Min.", 1100), tc("Rückschau an zwei unbeschrifteten Bildern (Ballhaus, Bastille): Was geschah, was bewirkte es? Allein, dann kurzes Plenum mit Stichproben. Die Klasse entschlüsselt „geschworen“ und „gestürmt“ selbst und verbindet die Ereignisse mit dem Entgegenkommen (Ursache-Folge). Echte Wiederholung statt Anriss; Voraussetzung für „vorher, nachher“.", 7400)]),
  row([tc("Informieren", 1700, { bold: true, fill: LIGHT }), tc("10 Min.", 1100), tc("Direktes Vorgehen, weil das Vorwissen zur Quellenarbeit gering ist. Portion 1: Verzicht auf die Vorrechte (4. August, vorher und nachher, nicht freiwillig). Portion 2: Erklärung vom 26. August und die drei Schritte; die Lehrkraft denkt an Artikel 1 laut. Prüffrage mit Stichproben (4 bis 6 Personen), 2 Minuten Reserve zum erneuten Erklären. Es wird nicht weitergegangen, wenn etwas nicht verstanden ist.", 7400)]),
  row([tc("Verarbeiten", 1700, { bold: true, fill: LIGHT }), tc("18 Min.", 1100), tc("Progression der direkten Instruktion in Einzelarbeit: erste Anwendung mit Hilfe (Artikel 3), selbstständiges Üben (Artikel 6, 11, 17), Kontrolle und Bilanz „Was hat sich verbessert?“. Dann der Aha-Moment: Grundgesetz kurz erklärt, drei Sätze „1789 oder Grundgesetz?“ (Denkphase), Auflösung als Paare, Frage nach der Bedeutung von 1789 für das Grundgesetz. Der Transfer läuft nur auf den Folien, nicht auf dem Blatt.", 7400)]),
  row([tc("Auswerten", 1700, { bold: true, fill: LIGHT }), tc("9 Min.", 1100), tc("Sicherung: Merksatz (abschreiben, rot umrahmen), Exit-Ticket mit der Ausgangsfrage (alle schreiben: „Wir wollen Menschenrechte, damit wir endlich …“, dazu Kreuz beim passenden Grundgesetz-Artikel), Zielprüfung, Ausblick (Brücke zu Stunde 5). Das Exit-Ticket enthält eine freiwillige Zeile „Das war mir unklar“ als Rückmeldung der Klasse.", 7400)]),
]));
s2.push(H2("6.2  Entscheidungen zur Methodenwahl"));
s2.push(...B([
  "**Lehrergesteuert mit Einzelarbeit:** Nach dem Partnerpuzzle der Vorwoche soll die Klasse nicht auf ständigen Sozialformwechsel konditioniert werden. Partnerarbeit entfällt weitgehend; Einzelarbeit mit Stichproben und kurzen Plenumsphasen trägt die Stunde.",
  "**Höchstens drei Methoden:** (1) Quellenauswertung in drei Schritten (Vormachen, angeleitet, selbstständig), (2) Aha-Moment „1789 oder Grundgesetz?“, (3) Exit-Ticket. Karikatur und Rückschau sind Impuls beziehungsweise Routine.",
  "**Quellenauswertung ohne Tabelle:** Das Arbeitsblatt arbeitet mit Artikel-Blöcken (Wortlaut, drei Schritte, Schreiblinien); Wortspeicher mit Erklärungen und Vorher-Speicher geben die Wissensbasis, statt reine Wortzuordnung zu verlangen.",
  "**Sprache:** Satzanfänge statt fertiger Sätze, ausschließlich Fachbegriffe der Reihe (Vorrechte, Volkssouveränität, Eigentumsrecht), keine kindlichen Metaphern.",
  "**Offene Auswertung:** Das Exit-Ticket lässt mehrere gute Antworten zu (Gleichheit, Mitbestimmung, Meinungsfreiheit, Eigentum). Geprüft werden eine Verbesserung für den Dritten Stand und ein passender Grundgesetz-Artikel.",
]));
s2.push(H2("6.3  Differenzierung und Unterstützung"));
s2.push(...B([
  "Gestufte Hilfen auf dem Blatt: Artikel 1 als gelbes Beispiel, Artikel 3 mit Hilfe, Artikel 6, 11 und 17 ohne Hilfe; Wortspeicher mit Erklärungen; Vorher-Speicher; Satzanfang auf dem Exit-Ticket.",
  "Zusätzliche Wege: Wer unsicher ist, beginnt mit dem kurzen Artikel 17 und schreibt nur einen Satzteil; wer früh fertig ist, erklärt seinem Nachbarn einen Artikel oder begründet beim Raten, woran er das Jahr erkannt hat.",
  "Alle Aufgaben mit Stichproben unterschiedlichen Niveaus; keine Selbsteinschätzung; Rückmeldung an die Lehrkraft über das Exit-Ticket.",
]));
s2.push(H2("6.4  Transparenz, Klassenführung, Fehlerkultur"));
s2.push(...B([
  "Ziel, Ablauf (A V I V A, mit Minuten) und Stern bei verkürzbarem Teil stehen zu Beginn an der Tafel und auf den Folien 3 und 4.",
  "Feste, wörtliche Ansagen bei jedem Wechsel; Übungen im immer gleichen Format (Nummer und Titel, nummerierte Schritte mit Verb am Anfang, Zeit, Arbeitsform, „Fertig? Stift hinlegen“).",
  "Fehlerkultur: Verbessern mit Grün; neutrale Reaktion bei Stichproben, erst danach Kommentar; Diagnosefragen sind unbenotet.",
]));
s2.push(H2("6.5  Medien"));
s2.push(P("Präsentation mit 23 Folien (Cambria/Calibri, mindestens 27 Punkt), selbst gezeichnete Karikatur mit zwei Wandbildern (keine historische Quelle), Arbeitsblatt (einseitig, Tricolore-Kopf mit kleinen Bildern), Exit-Ticket (6 pro A4), Tafel."));

// ---------- 7 Verlauf (Querformat) ----------
const s3 = [];
s3.push(H1("7  Verlaufsplanung"));
const VW = [800, 1300, 3700, 3000, 1700, 1500, 2400];
const VH = ["Min.", "Phase", "Inhalt und Lehrerimpuls (wörtlich, wo wichtig)", "Erwartetes Schülerverhalten", "Sozialform, Methode", "Medien", "Ziel, Begründung, Reserve"];
const VR = [
  ["4", "Ankommen", "Folie 1 beim Betreten. Ansage: „Stifte liegen. Ihr hört nur zu. Lest, was Klerus und Adel sagen. Schreibt ins Heft: Was will der Dritte Stand noch? Ich vermute …“ Danach Folie 2 bis 4 (Stundenfrage, Begriff Quelle, Ziele als „Ich kann …“, Ablauf), je 30 bis 45 Sekunden.", "Heft raus, Datum und Überschrift; schreiben ein bis zwei Sätze Vermutung. Hören zu.", "Einzelarbeit, still; Lehrervortrag", "Folien 1 bis 4, Seitentafel, Heft", "Rahmen; Vermutung als Vorher-Nachher für Ziel 1 bis 3. Folie „Das ist mir wichtig“ entfällt (eingeführt)."],
  ["4", "Vorwissen aktivieren", "Folie 5: „Schaut auf die zwei Bilder. Schreibt zu jedem: Was geschah? Was bewirkte es?“ (2 Min. allein), Stichproben bei 3 bis 4 Personen (1 Min.). Folie 6: Kette Ballhausschwur, Bastille, Aufstände; Zeilen 4. und 26. August nur nennen. „Vergleicht und verbessert mit Grün.“", "Erkennen Ballhaus und Bastille, schreiben Stichworte; nennen Ursache-Folge (Angst von Klerus und Adel). Verbessern mit Grün.", "Einzelarbeit, kurzes Plenum", "Folien 5 und 6, Heft", "Ziel 1 (Voraussetzung). Kein Partnercheck wegen lehrergesteuerter Stunde. Plenum auf 30 Sekunden kürzbar."],
  ["10", "Informieren", "Folie 7 (3 Min.): Verzicht auf die Vorrechte (4. August), nicht freiwillig, vorher und nachher; Warnung: abgeschafft heißt nicht sofort gleich; Prüffrage „Warum geben Klerus und Adel nach?“. Folie 8 (1 Min.): 26. August, drei Schritte (Meine Worte, Begriff heute, Vorher). Folie 9 (3 Min.): Modell Artikel 1, Lehrkraft denkt laut. Folie 10 (1 Min.): Prüffragen („Was machst du im Schritt 2?“, „Warum schreibst du nicht ab?“). Reserve 2 Min.", "Hören zu, geben kurze Antworten (aus Angst vor den Aufständen; ich suche den Begriff im Wortspeicher; ich fasse es in eigene Worte).", "Lehrervortrag in Portionen, Modell, Stichproben (4 bis 6)", "Folien 7 bis 10, Arbeitsblatt (noch umgedreht)", "Direktes Vorgehen; Ziel 1 und 2. Nicht weitergehen, wenn die Prüffrage nicht sitzt (Reserve nutzen)."],
  ["3", "Verarbeiten 1", "Folie 11: „Jetzt nehmt ihr das Arbeitsblatt. Jetzt arbeitet ihr allein.“ Artikel 3: in meinen Worten (mit Hilfe), Begriff ist angekreuzt, Vorher mit Hilfe. Folie 12 (30 Sek.): Lösung, „Vergleicht und verbessert mit Grün.“", "Schreiben in eigenen Worten, nutzen die Hilfe; vergleichen mit Grün.", "Einzelarbeit, leise", "Folien 11 und 12, Arbeitsblatt", "Ziel 1 und 2 (erste Anwendung mit Hilfe). Heft-Stichprobe."],
  ["6", "Verarbeiten 2", "Folie 13: Artikel 6, 11 und 17 allein in den drei Schritten (in meinen Worten, Begriff ankreuzen, Vorher mit dem Vorher-Speicher). Lehrkraft geht durch die Reihen, Stichprobe bei 4 bis 6 Personen, kein Kommentar.", "Schreiben in eigenen Worten, kreuzen Begriffe an, wählen Vorher aus dem Speicher.", "Einzelarbeit, leise", "Folie 13, Arbeitsblatt", "Ziel 1 und 2 (selbstständig). Differenzierung: kurzer Artikel 17 zuerst. Sollbruchstelle: nur Artikel 6 und 11."],
  ["3", "Verarbeiten 3", "Folie 14 (1,5 Min.): Begriffe zu allen Artikeln, auch 2, 4, 10 mündlich; Hinweis: Artikel 2 nennt Widerstand, nicht Gewalt (Stunde 3). Folie 15 (1,5 Min.): Bilanz „Was hat sich verbessert?“ (Artikel 3, 6, 11, 17: vorher, seitdem).", "Vergleichen mit Grün, nennen Verbesserungen (Gleichheit vor dem Gesetz, freie Rede, Eigentum, Macht beim Volk).", "Plenum kurz, Stichproben", "Folien 14, 15", "Ziel 1 (Bedeutung). Kontrolle der Übersetzung (Ziel 2)."],
  ["4", "Verarbeiten 4 (Aha)", "Folie 16 (30 Sek.): Grundgesetz kurz erklärt. Folie 17 (1 Min. allein): drei Sätze A, B, C, „1789 oder Grundgesetz?“, ins Heft. Folie 18 (1,5 Min.): Auflösung als Paare. Folie 19 (1 Min.): „Was bedeutet 1789 für unser Grundgesetz?“, Antwort: Vorbild.", "Notieren A, B, C; raten und begründen kurz; erkennen die Ähnlichkeit; formulieren: 1789 war ein Vorbild.", "Einzelarbeit (Denkphase), Lehrervortrag, Stichproben", "Folien 16 bis 19, Heft", "Ziel 3 (Transfer, Aha). Sollbruchstelle: nur Folie 17 und 18, Folie 19 in einem Satz."],
  ["2", "Auswerten 1", "Folie 20: „Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“ (4. August, 26. August, Gleichheit vor dem Gesetz, Freiheit, Eigentumsrecht, Volkssouveränität, Echo im Grundgesetz; Test: Recht gilt für alle, Vorrecht nur für wenige).", "Schreiben ab, rahmen rot.", "Einzelarbeit", "Folie 20, Heft", "Ergebnissicherung für den Test, Ziel 1 bis 3."],
  ["5", "Auswerten 2", "Folie 21: „Schaut auf die Karikatur. Ihr bekommt einen Zettel. Beendet den Satz ‚Wir wollen Menschenrechte, damit wir endlich …‘ und kreuzt den passenden Grundgesetz-Artikel an.“ 4 Min. schreiben, danach 2 bis 3 freiwillige Vorleser. Vergleich mit der Vermutung vom Anfang.", "Beenden den Satz (zum Beispiel „… alle vor dem Gesetz gleich sind“) und kreuzen Art. 3, 5, 14 oder 20 an.", "Einzelarbeit, leise; kurzes Vorlesen", "Folie 21, Exit-Ticket", "Abschlussaufgabe, prüft Ziel 1 und 3 bei allen. Zettel werden abgegeben."],
  ["2", "Auswerten 3", "Folie 22 (1 Min., still): „Haben wir die Ziele erreicht?“ Folie 23 (Ausblick, 1 Min.): „Für wen galten diese Rechte 1789 wirklich?“ Hinweis auf Stunde 5.", "Prüfen still die Ziele; hören den Ausblick.", "still, Lehrervortrag", "Folien 22, 23", "Zielprüfung (ohne Selbsteinschätzung); Brücke zu Stunde 5."],
];
s3.push(table(VW, [hdr(VH, VW), ...VR.map((r) => row(r.map((v, j) => tc(v, VW[j], { bold: j === 1, fill: j === 1 ? LIGHT : undefined, size: 17 }))))]));
s3.push(P("Gesamt: 4 + 4 + 10 + 3 + 6 + 3 + 4 + 2 + 5 + 2 = 43 Minuten plus 2 Minuten Reserve (45 Minuten).", { size: 19, after: 60 }));
s3.push(H2("7.1  Gelenkstellen"));
s3.push(...B([
  "Ankommen → Vorwissen: „Was haben Ballhaus und Bastille mit dem Entgegenkommen zu tun?“",
  "Vorwissen → Informieren: „Klerus und Adel kamen entgegen, weil sie Angst hatten. Was genau haben sie aufgegeben?“",
  "Informieren → Verarbeiten: „Entgegenkommen ist Gnade von oben. Was ist der Gegenpol? Das lest ihr jetzt in der Quelle. Jetzt macht ihr es selbst.“",
  "Verarbeiten → Aha: „Das war vor 237 Jahren. Was davon gilt noch heute?“",
  "Aha → Auswerten: „Jetzt beantwortet ihr die Frage vom Anfang.“ Auswerten → nächste Stunde: „Gilt das für alle? Das sehen wir nächste Woche.“",
], 20));
s3.push(H2("7.2  Didaktische Reserve und Sollbruchstellen"));
s3.push(...B([
  "Zeit hat Vorrang; die Stunde ist dicht. Reserve: 2 Minuten in „Informieren“ (erneutes Erklären).",
  "Kürzung in dieser Reihenfolge: Plenum der Rückschau (auf 30 Sekunden), Artikel 11 und 17 (nur Artikel 6 und 11), Aha (nur Folie 17 und 18), Vorlesen der Exit-Tickets. „Auswerten“ wird nicht gestrichen; die Übungsphase endet 9 Minuten vor Schluss.",
], 20));
const sec3 = { properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 800, bottom: 700, left: 700, right: 700 } } }, footers: { default: footer("Unterrichtsentwurf Geschichte 8D · 25.11.2026") }, children: s3 };

// ---------- 8-11 ----------
const s4 = [];
s4.push(H1("8  Erwartungshorizont"));
s4.push(table([2300, 7900], [
  hdr(["Aufgabe", "Erwartete Lösung (Stichworte, andere richtige Formulierungen zählen)"], [2300, 7900]),
  row([tc("Rückschau (Folie 5)", 2300, { bold: true, fill: LIGHT }), tc("Links: Eid im Ballhaus am 20. Juni, die Abgeordneten wollen nicht auseinandergehen (gewaltlos), die Versammlung bleibt. Rechts: Sturm auf die Bastille am 14. Juli (Gewalt), der König zieht die Soldaten ab, Aufstände im Land. Zusammenhang: Das Entgegenkommen von Klerus und Adel war nicht freiwillig, sie hatten Angst vor weiteren Aufständen.", 7900)]),
  row([tc("Prüffrage Folie 7", 2300, { bold: true, fill: LIGHT }), tc("Aus Angst vor den Aufständen der Bauern (und weil die Nationalversammlung es beschloss).", 7900)]),
  row([tc("Artikel 1 (Modell)", 2300, { bold: true, fill: LIGHT }), tc("In meinen Worten: Alle Menschen haben von Geburt an dieselben Rechte. Begriff: Gleichheit. Vorher: Die Geburt entschied über Stand und Vorrechte.", 7900)]),
  row([tc("Artikel 3 (Hilfe)", 2300, { bold: true, fill: LIGHT }), tc("Die Herrschaft geht vom Volk aus, nicht vom König allein. Volkssouveränität. Vorher: Der König bestimmte allein.", 7900)]),
  row([tc("Artikel 6, 11, 17 (selbst)", 2300, { bold: true, fill: LIGHT }), tc("6: Für alle gilt dasselbe Gesetz, Ämter nach Fähigkeit. Gleichheit (vor dem Gesetz). Vorher: Klerus und Adel zahlten keine Steuern, Ämter bekamen vor allem Adlige. 11: Jeder darf frei sagen, schreiben, drucken. Meinungsfreiheit. Vorher: Der König kontrollierte, was gedruckt wurde (Zensur). 17: Was jemandem gehört, darf ihm niemand wegnehmen. Eigentumsrecht. Vorher: Bauern gaben einen Teil der Ernte und Dienste an Adel und Kirche ab, ihr Eigentum war nicht gesichert.", 7900)]),
  row([tc("Bilanz (Folie 15)", 2300, { bold: true, fill: LIGHT }), tc("Gleiche Rechte statt Vorrechte nach Geburt, Schutz statt Willkür des Königs, Macht beim Volk.", 7900)]),
  row([tc("1789 oder Grundgesetz? (Folie 17)", 2300, { bold: true, fill: LIGHT }), tc("A „Das Gesetz muss für alle gleich sein.“ = 1789 (Artikel 6). B „Alle Staatsgewalt geht vom Volke aus.“ = Grundgesetz (Art. 20 Abs. 2). C „Das Eigentum und das Erbrecht werden gewährleistet.“ = Grundgesetz (Art. 14 Abs. 1). Paare: Art. 6 und GG Art. 3 (1); Art. 3 und GG Art. 20 (2); Art. 17 und GG Art. 14 (1). Bedeutung: 1789 war ein Vorbild, Rechte aus der Erklärung stehen noch heute im Grundgesetz.", 7900)]),
  row([tc("Exit-Ticket (Folie 21)", 2300, { bold: true, fill: LIGHT }), tc("Beispiele: „Wir wollen Menschenrechte, damit wir endlich alle vor dem Gesetz gleich sind“ (Kreuz: Art. 3); „… mitbestimmen dürfen“ (Art. 20); „… frei reden und schreiben dürfen“ (Art. 5); „… unser Eigentum behalten“ (Art. 14). Mehrere gute Antworten; Stufe 1: nennt eine Verbesserung; Stufe 3: Kreuz passt zum Satz.", 7900)]),
]));
s4.push(H1("9  Mögliche Schwierigkeiten und Reaktionen"));
s4.push(table([3400, 6800], [
  hdr(["Mögliche Schwierigkeit", "Reaktion"], [3400, 6800]),
  row([tc("Die Klasse versteht „Quelle ins Heute übersetzen“ nicht.", 3400), tc("Modell an Artikel 1 wiederholen (Reserve), Wortspeicher mit Erklärungen zeigen; Stichproben, bis die Prüffrage sitzt.", 6800)]),
  row([tc("Es wird laut, weil alle gleichzeitig schreiben oder raten.", 3400), tc("„Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“ Kein Gespräch; Aufträge bleiben schriftlich.", 6800)]),
  row([tc("Zeit reicht nicht (Übungsphase dauert länger).", 3400), tc("Sollbruchstellen (Abschnitt 7.2); Auswerten bleibt, Vorlesen der Zettel entfällt.", 6800)]),
  row([tc("Schülerinnen und Schüler halten „abgeschafft“ für „sofort umgesetzt“.", 3400), tc("Warnung vor dem Fehler (Folie 7): Abgaben mussten teils abgelöst werden; die Nationalversammlung beschloss, nicht der König.", 6800)]),
  row([tc("Die Klasse hat die Zensur nicht im Wissen.", 3400), tc("Der Vorher-Speicher auf dem Blatt nennt sie; die Schüler wählen aus.", 6800)]),
  row([tc("Aha bleibt aus (die Sätze werden erkannt).", 3400), tc("Aussage: Genau das zeigt die Vorbildwirkung; Folie 19 knüpft die Bedeutung an. Notfalls vierten Satz ergänzen (Meinungsfreiheit).", 6800)]),
]));
s4.push(H1("10  Anhang: Unterlagen"));
s4.push(...B([
  "Präsentation_Stunde4_Menschenrechte_8D (23 Folien; Notizen mit Ansagen, Zielbezug und Lösungen)",
  "Arbeitsblatt_Menschenrechte_Quelle_8D (einseitig, 26-mal) und Loesung_Arbeitsblatt_Menschenrechte_8D",
  "Exit-Ticket_Menschenrechte_8D (6 Zettel pro A4, 5-mal drucken und schneiden)",
  "Tafelskript_Stunde4_Menschenrechte_8D (wörtliche Ansagen, Warnungen vor Fehlern, Erwartungen, Heftbild) und Materialliste_Stunde4_Menschenrechte_8D",
]));
s4.push(H2("Heftbild am Ende der Stunde"));
s4.push(P("Überschrift mit Datum und Stundenfrage; Vermutung („Was will der Dritte Stand noch? Ich vermute …“); Rückschau zu den zwei Bildern (verbessert mit Grün); Notiz zu Folie 17 (A, B, C: 1789 oder Grundgesetz); Merksatz 1 (abgeschrieben, rot umrahmt). Arbeitsblatt eingeklebt, Exit-Ticket abgegeben."));
s4.push(H1("11  Quellen und Grundlagen"));
s4.push(...B([
  "Erklärung der Menschen- und Bürgerrechte vom 26. August 1789 (gekürzte und vereinfachte Fassung für die Klasse, Nummerierung nach dem Original); Grundgesetz für die Bundesrepublik Deutschland, Art. 3, 5, 14, 20.",
  "Schulinternes Curriculum (SchiC) Geschichte Klasse 8, Modul 2 „Französische Revolution“; Rahmenlehrplan Teil C.",
  "Lehrbuch (nur als Inspiration, keine Seiten im Unterricht); Material und Planungsmodell AVIVA aus dem Studienseminar; Hospitationskriterien (kognitive Aktivierung, konstruktive Unterstützung, Klassenführung); Literatur zur Direkten Instruktion, zu Kompetenzrastern und zur Überprüfbarkeit des Lernzuwachses (siehe Methoden/Literatur_Unterrichtsqualitaet.md im Planungsordner).",
]));
s4.push(H2("Offene Punkte vor der Hospitation"));
s4.push(...B([
  "Kopfdaten ergänzen (Uhrzeit, Raum, Namen); Rückmeldung zu Stunde 3 in Abschnitt 1.2 eintragen.",
  "Wortlaut der Artikel und des Grundgesetzes mit den Originalen abgleichen; Angaben zum 4. August (abgeschaffte Vorrechte, Ablösung der Abgaben) im Lehrbuch prüfen.",
  "Der Wortspeicher der Parallelklasse enthält das Eigentumsrecht nicht; Absprache zur Vergleichbarkeit im Test.",
], 20));
const sec4 = { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer("Unterrichtsentwurf Geschichte 8D · 25.11.2026") }, children: s2.concat(s4) };
// s2 (5 bis 6) und s4 (8 bis 11) liegen im Hochformat; der Verlauf (7) liegt im Querformat dazwischen.
const secA = { properties: sec4.properties, footers: sec4.footers, children: s2 };
const secB = sec3;
const secC = { properties: sec4.properties, footers: sec4.footers, children: s4 };

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [sec1, secA, secB, secC] });
  fs.writeFileSync("../Unterrichtsentwurf_Menschenrechte_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
