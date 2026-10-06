import chromium from '@sparticuz/chromium';
import puppeteer from 'puppeteer-core';
import { existsSync, readFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

export const config = {
  maxDuration: 30,
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
  },
};

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

// Dibaca sekali saat cold start — di-inject inline ke headless browser
// JSZip harus dimuat lebih dulu karena docx-preview UMD bergantung pada window.JSZip
const jszipScript = readFileSync(join(root, 'node_modules/jszip/dist/jszip.min.js'), 'utf-8');
const docxPreviewScript = readFileSync(join(root, 'node_modules/docx-preview/dist/docx-preview.js'), 'utf-8');

const PAGE_HTML = `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <style>
      * { margin: 0; padding: 0; box-sizing: border-box; }
      body { background: white; }
      .docx-wrapper { background: white !important; padding: 0 !important; }
      /* Cegah halaman kosong di akhir: hapus break-after dari section/article terakhir */
      section:last-of-type,
      article:last-of-type,
      .docx:last-of-type {
        break-after: auto !important;
        page-break-after: auto !important;
      }
    </style>
  </head>
  <body>
    <div id="root"></div>
  </body>
</html>`;

async function getLaunchOptions() {
  if (process.env.AWS_LAMBDA_FUNCTION_NAME) {
    return {
      args: chromium.args,
      executablePath: await chromium.executablePath(),
      headless: chromium.headless,
      defaultViewport: { width: 794, height: 1123 },
    };
  }

  const username = process.env.USERNAME || process.env.USER || '';
  const candidates = [
    `C:\\Users\\${username}\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe`,
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    `C:\\Users\\${username}\\AppData\\Local\\BraveSoftware\\Brave-Browser\\Application\\brave.exe`,
    'C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe',
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium-browser',
  ];

  const executablePath = candidates.find(p => existsSync(p));
  if (!executablePath) {
    throw new Error('Browser tidak ditemukan. Install Google Chrome di komputer ini.');
  }

  return {
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    executablePath,
    headless: true,
    defaultViewport: { width: 794, height: 1123 },
  };
}

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
    browser = await puppeteer.launch(await getLaunchOptions());

    const page = await browser.newPage();
    await page.setContent(PAGE_HTML, { waitUntil: 'domcontentloaded' });

    // 1. Load JSZip dulu → window.JSZip
    await page.addScriptTag({ content: jszipScript });
    // 2. Load docx-preview raw (UMD branch browser) → window.docx
    await page.addScriptTag({ content: docxPreviewScript });

    // Render DOCX → DOM di dalam headless browser
    await page.evaluate(async (base64) => {
      const binary = atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      await window.docx.renderAsync(
        bytes.buffer,
        document.getElementById('root'),
        null,
        { inWrapper: true, breakPages: true, useBase64URL: true }
      );

      // docx-preview membungkus tiap "page"/section Word sebagai child
      // langsung dari .docx-wrapper. root.lastElementChild adalah wrapper
      // itu sendiri (berisi semua page), jadi harus dicek children di
      // dalamnya, bukan root.
      const wrapper = document.querySelector('.docx-wrapper');
      if (wrapper) {
        let pages = Array.from(wrapper.children);
        // Hapus page kosong di ujung (biasanya sisa sectPr trailing di docx
        // yang tidak dirender Word tapi dirender docx-preview)
        while (pages.length > 1) {
          const lastPage = pages[pages.length - 1];
          const hasText = lastPage.textContent.replace(/\s/g, '').length > 0;
          const hasVisual = lastPage.querySelector('img, svg, table');
          if (hasText || hasVisual) break;
          lastPage.remove();
          pages.pop();
        }

        // Hilangkan efek "kertas" (box-shadow/border) bawaan docx-preview
        // untuk preview di layar — kalau tidak dihapus, ikut kecetak ke PDF
        pages.forEach((p) => {
          p.style.setProperty('box-shadow', 'none', 'important');
          p.style.setProperty('border', 'none', 'important');
          p.style.setProperty('margin', '0', 'important');
        });

        const finalPage = pages[pages.length - 1];
        if (finalPage) {
          finalPage.style.setProperty('break-after', 'auto', 'important');
          finalPage.style.setProperty('page-break-after', 'auto', 'important');
        }
      }
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
