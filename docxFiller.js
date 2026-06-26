import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import fs from 'fs';

/**
 * Fill a .docx template with data using single-brace {field} placeholders.
 * Returns a Buffer of the filled document.
 */
export function fillDocx(templatePath, data) {
  const content = fs.readFileSync(templatePath, 'binary');
  const zip     = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks:    true,
    nullGetter:    () => '',   // trả '' cho tag thiếu thay vì throw
  });

  doc.render(data);

  return doc.getZip().generate({ type: 'nodebuffer' });
}
