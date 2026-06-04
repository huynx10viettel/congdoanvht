// Điền dữ liệu form vào template Word (.docx) bằng docxtemplater (mã nguồn mở).
// Template dùng cú pháp placeholder {ten_truong}. Xem templates/mau-don-phuc-loi.docx.
import fs from "node:fs";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

/**
 * @param {string} templatePath đường dẫn file .docx template
 * @param {Object} data dữ liệu điền vào, key trùng tên placeholder
 * @returns {Buffer} buffer của file .docx đã điền
 */
export function fillDocx(templatePath, data) {
  const content = fs.readFileSync(templatePath, "binary");
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    // Nếu thiếu trường nào thì để trống thay vì báo lỗi
    nullGetter: () => "",
  });

  doc.render(data);

  return doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}
