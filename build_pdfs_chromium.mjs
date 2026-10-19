// Headless Chromium renderer used by build_pdfs.py on Windows.
// Usage: node build_pdfs_chromium.mjs '[["in.html", "out.pdf"], ...]'
import puppeteer from 'puppeteer';
import { pathToFileURL } from 'node:url';

const jobs = JSON.parse(process.argv[2] ?? '[]');
const browser = await puppeteer.launch({ headless: true });
try {
  for (const [html, pdf] of jobs) {
    const page = await browser.newPage();
    await page.goto(pathToFileURL(html).href, { waitUntil: 'networkidle0' });
    await page.evaluate(() => document.fonts.ready);
    await page.pdf({
      path: pdf,
      printBackground: true,
      preferCSSPageSize: true, // honour @page { size: A4 } from the CV stylesheet
      tagged: true,
    });
    await page.close();
  }
} finally {
  await browser.close();
}
