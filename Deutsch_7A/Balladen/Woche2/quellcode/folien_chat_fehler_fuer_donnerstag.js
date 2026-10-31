// Ü4 Fehler finden
{
  const s = base("Verarbeiten · Fehler finden", "Ü4  Was stimmt im Chat nicht?", "Lösung: 1 …, weil ich noch duschen muss. 2 …, dass wir morgen frei haben. 3 Weißt du, ob der Neue auch kommt? 4 Ich freue mich, weil wir frei haben. (Komma fehlt). Reflexion: Beim Sprechen hört man oft „weil ich muss …“ – geschrieben ist das falsch.");
  const msgs = [
    ["l", "Komme später, weil ich muss noch duschen 🙈"],
    ["r", "Trainer hat gesagt, dass wir haben morgen frei!"],
    ["l", "Weißt du, ob kommt der Neue auch?"],
    ["r", "Ich freue mich weil wir frei haben 💪"],
  ];
  box(s, 0.6, 1.55, 7.2, 4.5, "E9EEF6");
  msgs.forEach(([side, m], i) => {
    const y = 1.75 + i * 1.05;
    const x = side === "l" ? 0.85 : 2.35;
    box(s, x, y, 5.2, 0.8, side === "l" ? WHITE : "DCF3E4", { r: 0.18 });
    txt(s, (i + 1) + "  " + m, { x: x + 0.2, y: y + 0.17, w: 4.9, h: 0.5, fontSize: 17 });
  });
  [["1", "Finde in jeder Nachricht den Fehler."], ["2", "Schreibe den Satz richtig ins Heft."], ["3", "Notiere: Verbstellung oder Komma?"]].forEach(([n, x], i) => {
    num(s, n, 8.1, 1.8 + i * 1.0, VIO, 0.45);
    txt(s, x, { x: 8.7, y: 1.8 + i * 1.0, w: 4.0, h: 0.9, fontSize: 18, bold: true });
  });
  zeit(s, "5 Minuten · zu zweit");
}
// Ü4 Lösung
{
  const s = base("Verarbeiten · Ü4 Lösung", "Ü4  So ist es richtig", "Vergleichen und mit Grün verbessern. Kurz reflektieren: Beim Sprechen hört man oft „weil ich muss …“ – geschrieben ist das falsch.");
  [[[["Komme später, ", ""], ["weil", "sub"], [" ich noch duschen ", ""], ["muss", "fin"], [".", ""]], "Verbstellung"],
   [[["Trainer hat gesagt, ", ""], ["dass", "sub"], [" wir morgen frei ", ""], ["haben", "fin"], ["!", ""]], "Verbstellung"],
   [[["Weißt du, ", ""], ["ob", "sub"], [" der Neue auch ", ""], ["kommt", "fin"], ["?", ""]], "Verbstellung"],
   [[["Ich freue mich", ""], [",", "b"], [" ", ""], ["weil", "sub"], [" wir frei ", ""], ["haben", "fin"], [".", ""]], "Komma"]].forEach(([r, typ], i) => {
    const y = 1.6 + i * 1.05;
    box(s, 0.6, y, 12.1, 0.85, LIGHT);
    txt(s, String(i + 1), { x: 0.8, y: y + 0.22, w: 0.5, h: 0.4, fontSize: 18, bold: true, color: VIO });
    txt(s, runs(r, 22), { x: 1.35, y: y + 0.2, w: 8.8, h: 0.5, fontFace: HEAD });
    box(s, 10.4, y + 0.2, 2.1, 0.45, typ === "Komma" ? YEL : VIO, { r: 0.22 });
    txt(s, typ, { x: 10.4, y: y + 0.2, w: 2.1, h: 0.45, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}
