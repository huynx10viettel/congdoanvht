// Gộp nhiều ảnh (JPG/PNG) thành 1 file PDF bằng pdf-lib (mã nguồn mở).
// Mỗi ảnh là 1 trang A4, căn giữa, giữ tỉ lệ.
import { PDFDocument } from "pdf-lib";

const A4 = { width: 595.28, height: 841.89 }; // điểm (points), khổ A4 dọc
const MARGIN = 28; // ~1cm

/**
 * @param {{buffer: Buffer, mimetype: string}[]} images danh sách ảnh
 * @returns {Promise<Buffer>} buffer PDF
 */
export async function imagesToPdf(images) {
  const pdf = await PDFDocument.create();

  for (const img of images) {
    let embedded;
    if (img.mimetype === "image/png") {
      embedded = await pdf.embedPng(img.buffer);
    } else if (img.mimetype === "image/jpeg" || img.mimetype === "image/jpg") {
      embedded = await pdf.embedJpg(img.buffer);
    } else {
      // Bỏ qua định dạng không hỗ trợ
      continue;
    }

    const page = pdf.addPage([A4.width, A4.height]);
    const maxW = A4.width - MARGIN * 2;
    const maxH = A4.height - MARGIN * 2;

    const scale = Math.min(maxW / embedded.width, maxH / embedded.height, 1);
    const w = embedded.width * scale;
    const h = embedded.height * scale;

    page.drawImage(embedded, {
      x: (A4.width - w) / 2,
      y: (A4.height - h) / 2,
      width: w,
      height: h,
    });
  }

  const bytes = await pdf.save();
  return Buffer.from(bytes);
}
