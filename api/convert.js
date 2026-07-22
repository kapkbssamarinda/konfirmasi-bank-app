import chromium from '@sparticuz/chromium';
import puppeteer from 'puppeteer-core';

export const config = {
  maxDuration: 30,
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
  },
};

const PAGE_HTML = `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <style>
      * { margin: 0; padding: 0; box-sizing: border-box; }
      body { background: white; }
      .docx-wrapper { background: white !important; padding: 0 !important; }
    </style>
  </head>
  <body>
    <div id="root"></div>
  </body>
</html>`;

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { docxBase64 } = req.body;
  if (!docxBase64) {
    return res.status(400).json({ error: 'Missing docxBase64' });
  }

  let browser;
  try {
    browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: { width: 794, height: 1123 },
      executablePath: await chromium.executablePath(),
      headless: chromium.headless,
    });

    const page = await browser.newPage();
    await page.setContent(PAGE_HTML, { waitUntil: 'domcontentloaded' });

    // Load docx-preview dari CDN ke dalam headless browser
    await page.addScriptTag({
      url: 'https://unpkg.com/docx-preview@0.3.7/dist/docx-preview.js',
    });

    // Render DOCX → DOM di dalam headless browser
    await page.evaluate(async (base64) => {
      const binary = atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      await docx.renderAsync(bytes.buffer, document.getElementById('root'), null, {
        inWrapper: true,
        breakPages: true,
        useBase64URL: true,
      });
    }, docxBase64);

    // Chrome print engine → PDF asli (vector, teks bisa diselect)
    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.send(Buffer.from(pdfBuffer));
  } catch (err) {
    console.error('PDF conversion error:', err);
    res.status(500).json({ error: err.message });
  } finally {
    if (browser) await browser.close();
  }
}
