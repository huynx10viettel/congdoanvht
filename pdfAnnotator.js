import { PDFDocument, PDFName, PDFString, PDFArray } from 'pdf-lib';

/**
 * Thêm 3 FreeText annotation comment vào trang cuối của PDF.
 * Numbered 1, 2, 3 từ trái qua phải.
 */
export async function addCommentsToPdf(pdfBuffer) {
  const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
  const pages  = pdfDoc.getPages();
  const page   = pages[pages.length - 1];
  const { width } = page.getSize();

  const boxW = Math.floor((width - 80) / 3);
  const boxH = 60;
  const y    = 30; // từ dưới lên

  for (let i = 0; i < 3; i++) {
    const x = 40 + i * (boxW + 10);

    const annotDict = pdfDoc.context.obj({
      Type:    'Annot',
      Subtype: 'FreeText',
      Rect:    [x, y, x + boxW, y + boxH],
      Contents: PDFString.of(`${i + 1}.`),
      DA:      PDFString.of('/Helvetica 10 Tf 0 0 0 rg'),
      C:       [1, 1, 0],          // màu vàng
      F:       4,                  // Print flag
      BS:      pdfDoc.context.obj({ Type: 'Border', W: 1, S: 'S' }),
      T:       PDFString.of(`Ghi chú ${i + 1}`),
    });

    const annotRef = pdfDoc.context.register(annotDict);

    const existingAnnots = page.node.get(PDFName.of('Annots'));
    if (existingAnnots instanceof PDFArray) {
      existingAnnots.push(annotRef);
    } else {
      page.node.set(PDFName.of('Annots'), pdfDoc.context.obj([annotRef]));
    }
  }

  return Buffer.from(await pdfDoc.save());
}
