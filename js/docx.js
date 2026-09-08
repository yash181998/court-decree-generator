import { BASE_PARTS } from './base-parts.js';

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** Zip the shared package parts together with the generated document + header. */
export async function packDocx({ documentXml, headerXml }) {
  const zip = new window.JSZip();
  for (const [name, content] of Object.entries(BASE_PARTS)) {
    zip.file(name, content);
  }
  zip.file('word/document.xml', documentXml);
  zip.file('word/header1.xml', headerXml);
  return zip.generateAsync({
    type: 'blob',
    mimeType: DOCX_MIME,
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

export function safeFileName(caseTitle) {
  return caseTitle.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '') || 'decree';
}
