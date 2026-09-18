// Runs entirely on-device: mammoth turns the .docx body into HTML, then the
// browser's own Print dialog ("Save as PDF") produces the file. No server,
// no upload - but the layout is an approximation of Word, not byte-exact.

const $ = (id) => document.getElementById(id);

let lastGenerated = null; // { blob, filename } of the most recently generated MC/OS/Notice document

/** Called by app.js right after a decree/notice is generated, so it can be reused here. */
export function setLastGenerated(blob, filename) {
  lastGenerated = { blob, filename };
  const btn = $('pdf-use-last');
  if (btn) {
    btn.disabled = false;
    btn.textContent = `Use "${filename}"`;
  }
}

function printableDocument(title, bodyHtml) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title>
<style>
  body { font-family: "Times New Roman", Georgia, serif; font-size: 13pt; line-height: 1.5; margin: 1in; color: #000; }
  h1, h2, h3 { font-size: 1.05em; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0; }
  td, th { border: 1px solid #333; padding: 4px 8px; vertical-align: top; }
  p { margin: 0 0 6px; }
  @media print { body { margin: 0.75in; } }
</style>
</head><body>${bodyHtml}</body></html>`;
}

async function convertToPdfWindow(blob, title) {
  const status = $('pdf-status');
  status.textContent = 'Converting\u2026';
  status.classList.remove('is-warn');

  // Must open synchronously, before any await, or mobile browsers treat the
  // later window.open as not tied to the click and silently block it.
  const win = window.open('', '_blank');
  if (!win) {
    status.textContent = 'Pop-up blocked \u2014 allow pop-ups for this site and try again.';
    status.classList.add('is-warn');
    return;
  }
  win.document.write('<p style="font-family:sans-serif">Converting\u2026</p>');

  try {
    const arrayBuffer = await blob.arrayBuffer();
    const result = await window.mammoth.convertToHtml({ arrayBuffer });
    win.document.open();
    win.document.write(printableDocument(title, result.value));
    win.document.close();
    win.focus();
    // Give the new tab a moment to lay out before opening the print sheet.
    setTimeout(() => win.print(), 300);
    status.textContent =
      'Opened in a new tab \u2014 use its Print option and choose "Save as PDF".' +
      (result.messages.length ? ' (Some formatting could not be converted exactly.)' : '');
  } catch (err) {
    win.close();
    status.textContent = `Could not convert: ${err.message}`;
    status.classList.add('is-warn');
  }
}

export function setupWordToPdf() {
  const fileInput = $('pdf-file');
  const useLastBtn = $('pdf-use-last');
  const convertBtn = $('pdf-convert');
  if (!fileInput || !useLastBtn || !convertBtn) return;

  convertBtn.addEventListener('click', async () => {
    const status = $('pdf-status');
    if (fileInput.files && fileInput.files[0]) {
      const file = fileInput.files[0];
      await convertToPdfWindow(file, file.name.replace(/\.docx$/i, ''));
    } else if (lastGenerated) {
      await convertToPdfWindow(lastGenerated.blob, lastGenerated.filename.replace(/\.docx$/i, ''));
    } else {
      status.textContent =
        'Choose a .docx file, or generate one first and use "Use the last generated document".';
      status.classList.add('is-warn');
    }
  });

  useLastBtn.addEventListener('click', async () => {
    if (!lastGenerated) return;
    fileInput.value = '';
    await convertToPdfWindow(lastGenerated.blob, lastGenerated.filename.replace(/\.docx$/i, ''));
  });
}
