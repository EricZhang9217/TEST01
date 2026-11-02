const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const files = process.argv.slice(2);
  const b = await chromium.launch({ args: ['--no-sandbox'] });
  const p = await b.newPage({ viewportSize: { width: 1600, height: 400 }, deviceScaleFactor: 2 });
  for (const f of files) {
    await p.goto('file://' + path.resolve(f));
    await p.waitForTimeout(120);
    // the page's own body width decides the canvas; the viewport follows it
    const w = Math.ceil(await p.evaluate(
      () => document.body.getBoundingClientRect().width));
    await p.setViewportSize({ width: w, height: 400 });
    await p.waitForTimeout(80);
    const h = Math.ceil(await p.evaluate(
      () => Math.max(document.body.scrollHeight,
                     document.body.getBoundingClientRect().height)));
    await p.setViewportSize({ width: w, height: Math.max(120, h) });
    await p.waitForTimeout(80);
    const out = f.replace(/\.html$/, '.png');
    await p.screenshot({ path: out });
    console.log(`${out}  ${w}x${h}  aspect ${(w / h).toFixed(3)}`);
  }
  await b.close();
})();
