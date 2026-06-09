import { PDFDocument } from 'pdf-lib';

/**
 * Merge uploaded files (images + PDFs) into a single PDF buffer.
 * Images (jpeg/png) are embedded as full pages.
 * PDFs are merged page-by-page.
 */
export async function imagesToPdf(files) {
  const merged = await PDFDocument.create();

  for (const file of files) {
    const { buffer, mimetype } = file;

    if (mimetype === 'application/pdf') {
      // Copy all pages from the uploaded PDF
      const srcDoc = await PDFDocument.load(buffer);
      const pages  = await merged.copyPages(srcDoc, srcDoc.getPageIndices());
      pages.forEach(p => merged.addPage(p));

    } else if (mimetype === 'image/jpeg' || mimetype === 'image/jpg') {
      const img  = await merged.embedJpg(buffer);
      const page = merged.addPage([img.width, img.height]);
      page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });

    } else if (mimetype === 'image/png') {
      const img  = await merged.embedPng(buffer);
      const page = merged.addPage([img.width, img.height]);
      page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
    }
    // Skip unsupported types silently
  }

  return Buffer.from(await merged.save());
}
