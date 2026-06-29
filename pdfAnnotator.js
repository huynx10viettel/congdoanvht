import { PDFDocument, PDFName, PDFString, PDFArray } from 'pdf-lib';

/**
 * Thêm 3 sticky note annotation (Text/Note icon) vào trang cuối PDF
 * tại 3 vị trí ký: Công đoàn bộ phận | Chỉ huy đơn vị | Công đoàn cơ sở
 */
export async function addCommentsToPdf(pdfBuffer) {
  const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
  const pages  = pdfDoc.getPages();
  const page   = pages[pages.length - 1];
  const { width, height } = page.getSize();

  const iconSize = 18;                          // kích thước icon sticky note (pt)
  const yPos     = Math.round(height * 0.18);  // ~18% từ dưới = dòng ký tên
  const colWidth = width / 3;                   // chia 3 cột đều

  for (let i = 0; i < 3; i++) {
    const x = Math.round(colWidth * i + colWidth / 2 - iconSize / 2);

    const annotDict = pdfDoc.context.obj({
      Type:     'Annot',
      Subtype:  'Text',           // sticky note icon (không phải FreeText box)
      Rect:     [x, yPos, x + iconSize, yPos + iconSize],
      Contents: PDFString.of(''),
      Name:     PDFName.of('Note'),
      C:        [1, 1, 0],        // vàng
      F:        4,                // Print flag
      Open:     false,
    });

    const annotRef     = pdfDoc.context.register(annotDict);
    const existingAnnots = page.node.get(PDFName.of('Annots'));
    if (existingAnnots instanceof PDFArray) {
      existingAnnots.push(annotRef);
    } else {
      page.node.set(PDFName.of('Annots'), pdfDoc.context.obj([annotRef]));
    }
  }

  return Buffer.from(await pdfDoc.save());
}
