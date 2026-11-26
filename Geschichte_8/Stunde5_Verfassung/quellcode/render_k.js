const { chromium } = require("playwright");
(async () => { const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1280, height: 970 }, deviceScaleFactor: 2 }); await p.goto("file://" + process.cwd() + "/karikatur_papier.svg.html"); await p.locator("#k").screenshot({ path: "../material/karikatur_papier.png" }); await b.close(); })();
