const { chromium } = require("playwright");
(async () => { const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1600, height: 800 }, deviceScaleFactor: 2 }); await p.goto("file://" + process.cwd() + "/schaubild.svg.html"); await p.locator("#s").screenshot({ path: "../material/schaubild_verfassung.png" }); await b.close(); })();
