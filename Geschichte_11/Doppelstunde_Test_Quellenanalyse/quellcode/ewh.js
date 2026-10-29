// Erwartungshorizont und Notenschlüssel zum Test
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Header } = require("docx");
const { C, p, h, table, box, schluessel } = require("./lib.js");
const OUT = process.argv[2];
const W = 9638, MAX = 37;
const K = schluessel(MAX);
const fmt = x => String(x).replace(".", ",");
const content = [
  p("GESCHICHTE · KLASSE 11 · TEST · LEHRKRAFT", { size: 16, bold: true, color: C.terra, after: 30 }),
  new Paragraph({ children: [new TextRun({ text: "Erwartungshorizont und Notenschlüssel", font: "Cambria", bold: true, size: 32 })], spacing: { after: 60 } }),
  p("Test „Athen um 600 v. Chr. – Drakon, Solon, Peisistratos“ · 45 Minuten · **37 Punkte + 1 Zusatzpunkt** (der Zusatzpunkt kann fehlende Punkte ausgleichen, die Höchstpunktzahl bleibt 37).", { size: 19 }),

  h("Aufgabe 1 · Krise um 600 (6 P.)"),
  table(W, [2300, 6338, 1000], ["Teil", "Erwartung", "Punkte"], [
    ["a) vier Probleme", "Schuldknechtschaft · Land in den Händen weniger (Adel/Großgrundbesitzer) · Recht nur beim Adel, Blutrache/keine geschriebenen Gesetze · politische Macht nur beim Adel. Auch akzeptieren: Bevölkerungswachstum/Landmangel, Verarmung der Bauern, drohender Aufruhr.", "je 1 = 4"],
    ["b) Weg in die Schuldknechtschaft", "Kleinbauer kann von seinem Land kaum leben → leiht Saatgut/Getreide beim Adligen → kann nicht zurückzahlen → verliert sein Land und muss mit Leib und Leben haften: wird abhängig oder Sklave, evtl. ins Ausland verkauft.", "2 (Ursache 1, Folge 1)"],
  ], { firstFill: true }),

  h("Aufgabe 2 · Richtig oder falsch (9 P.)"),
  table(W, [700, 1300, 6638, 1000], ["", "Lösung", "Berichtigung", "Punkte"], [
    ["a", "richtig", "–", "1"],
    ["b", "falsch", "Die Schuldknechtschaft blieb unter Drakon bestehen (erst Solon hob sie auf).", "1 + 1"],
    ["c", "richtig", "–", "1"],
    ["d", "falsch", "Solon strich die Schulden, verteilte aber kein Land neu; die Großgrundbesitzer behielten ihr Land.", "1 + 1"],
    ["e", "falsch", "Peisistratos eroberte die Macht mit einem Söldnerheer und besetzte die Akropolis (546).", "1 + 1"],
    ["f", "richtig", "–", "1"],
  ], { firstFill: true }),

  h("Aufgabe 3 · Solons Verfassung (10 P.)"),
  table(W, [2300, 6338, 1000], ["Teil", "Erwartung", "Punkte"], [
    ["a) Ämter und Volksversammlung", "von links: Areopag – Archonten (Pfeil zum Areopag) – Rat der 400 – Geschworenengericht; unten: Volksversammlung", "je 1 = 5"],
    ["b) Rechte", "1. Klasse: alles · 2. und 3. Klasse: abstimmen, klagen, Rat – nicht Archont · Theten: nur abstimmen und klagen", "je Klasse 1 = 4"],
    ["c) Pfeile", "1. Klasse → Archonten und Rat der 400; 2. und 3. Klasse → Rat der 400; Theten → kein Amt", "1"],
  ], { firstFill: true }),

  h("Aufgabe 4 · Urteile erkennen (6 P.)"),
  table(W, [2300, 6338, 1000], ["Teil", "Erwartung", "Punkte"], [
    ["a)", "1 = B (überprüfbare Tatsache) · 2 = S (Urteil aus der damaligen Zeit, begründet) · 3 = W (heutiger Maßstab genannt) · 4 = W (wertet, aber ohne Maßstab und Begründung)", "je 1 = 4"],
    ["b) verbessertes Werturteil", "Maßstab genannt (z. B. gleiche Regeln: Macht mit Gewalt erobert → nicht gerechtfertigt; Schutz der Schwachen: Land für Besitzlose → teilweise gerecht) und begründet.", "Maßstab 1, Begründung 1"],
  ], { firstFill: true }),

  h("Aufgabe 5 · Selbst urteilen (6 P.)"),
  table(W, [2300, 6338, 1000], ["Teil", "Erwartung", "Punkte"], [
    ["a) Sachurteil", "Klare Position („weitgehend“/„teilweise“) aus der Sicht der damaligen Zeit. Belege z. B.: Schulden gestrichen, Schuldknechtschaft verboten, Versklavte zurückgekauft · Rechte nach Besitz statt Herkunft, Volksversammlung für alle Bürger · aber: kein neues Land, Ämter nur für Reiche, Theten ohne Amt, Unzufriedenheit danach (Peisistratos).", "Position 1, je Beleg 1"],
    ["b) Werturteil", "Maßstab ausdrücklich genannt (Gleichheit, Leistung, Schutz der Schwachen, gleiche Regeln), Position, Begründung mit Beispiel (z. B. Theten dürfen kein Amt übernehmen; Aufstieg durch Besitz möglich). Unterschiedliche Urteile sind richtig, wenn begründet.", "Maßstab 1, Position mit Begründung 1, Beispiel 1"],
  ], { firstFill: true }),
  p("**Zusatz (+1):** Timokratie = Staatsform, in der die politischen Rechte vom Besitz abhängen („Wer mehr hat, darf mehr“).", { size: 19, before: 80 }),

  h("Notenschlüssel (37 Punkte)"),
  table(W, [1800, 1800, 6038], ["Notenpunkte", "ab Prozent", "ab Punkten"], K.map(([np, pr, pt]) => [String(np), pr + " %", fmt(pt)]), { size: 18 }),
  p("Üblicher Oberstufen-Schlüssel nach Prozent; Mindestpunktzahl auf halbe Punkte aufgerundet. Bei Bedarf an den Schlüssel der Schule anpassen.", { size: 16, color: C.muted, before: 60 }),
];
const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Geschichte · Klasse 11 · Test – Erwartungshorizont", size: 16, color: "998C77" })] })] }) },
    children: content }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
