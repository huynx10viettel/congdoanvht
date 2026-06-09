import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { PDFDocument } from 'pdf-lib';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { soThanhChu } from './soThanhChu.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// ──────────────────────────────────────────────
// Template map based on loai_phuc_loi
// ──────────────────────────────────────────────
const TEMPLATE_MAP = {
  'tham-hoi-om':            'De nghi thanh toan tien cong doan, phuc loi_temp_om.docx',
  'ket-hon':                'De nghi thanh toan tien cong doan, phuc loi_temp_chucmung.docx',
  'sinh-con':               'De nghi thanh toan tien cong doan, phuc loi_temp_chucmung.docx',
  'tham-vieng-cbcnv':       'De nghi thanh toan tien cong doan, phuc loi_temp_CBNVtutran.docx',
  'tham-vieng-thannhan':    'De nghi thanh toan tien cong doan, phuc loi_temp_thannhanCBNVtutran.docx',
};

// su_kien label for chúc mừng template
const SU_KIEN_LABEL = {
  'ket-hon': 'kết hôn',
  'sinh-con': 'sinh con',
};

// ──────────────────────────────────────────────
// Microsoft Graph client
// ──────────────────────────────────────────────
function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });
  return Client.initWithMiddleware({ authProvider });
}

// ──────────────────────────────────────────────
// Fill docx template with docxtemplater
// ──────────────────────────────────────────────
function fillDocx(templatePath, data) {
  const content = fs.readFileSync(templatePath, 'binary');
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });
  doc.render(data);
  return doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ──────────────────────────────────────────────
// Convert uploaded images to PDF
// ──────────────────────────────────────────────
async function imagesToPdf(imageBuffers) {
  const pdfDoc = await PDFDocument.create();
  for (const buf of imageBuffers) {
    let img;
    // Detect PNG or JPEG by magic bytes
    if (buf[0] === 0x89 && buf[1] === 0x50) {
      img = await pdfDoc.embedPng(buf);
    } else {
      img = await pdfDoc.embedJpg(buf);
    }
    const page = pdfDoc.addPage([img.width, img.height]);
    page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
  }
  return pdfDoc.save();
}

// ──────────────────────────────────────────────
// Upload files to SharePoint via Graph API
// ──────────────────────────────────────────────
async function luuHoSo(ten_cbcnv, ma_nv, ngay, thang, nam, docxBuffer, pdfBuffer) {
  const client = getGraphClient();
  const driveId = process.env.DRIVE_ID;
  const folder = `Shared Documents/PhucLoiCongDoan`;
  const baseName = `${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}`;

  await client
    .api(`/drives/${driveId}/root:/${folder}/${baseName}.docx:/content`)
    .put(docxBuffer);

  if (pdfBuffer) {
    await client
      .api(`/drives/${driveId}/root:/${folder}/${baseName}_chungtu.pdf:/content`)
      .put(pdfBuffer);
  }
}

// ──────────────────────────────────────────────
// Static files
// ──────────────────────────────────────────────
app.use(express.static(__dirname));

// ──────────────────────────────────────────────
// POST /api/submit
// ──────────────────────────────────────────────
app.post('/api/submit', upload.array('images'), async (req, res) => {
  try {
    const {
      nguoi_de_nghi,
      chuc_vu_de_nghi,
      ten_cbcnv,
      ma_nv,
      chuc_danh,
      don_vi,
      loai_phuc_loi,
    } = req.body;

    // Validate required fields
    if (!nguoi_de_nghi || !ten_cbcnv || !ma_nv || !loai_phuc_loi) {
      return res.status(400).json({ error: 'Thiếu thông tin bắt buộc.' });
    }

    // Date
    const now = new Date();
    const ngay  = String(now.getDate()).padStart(2, '0');
    const thang = String(now.getMonth() + 1).padStart(2, '0');
    const nam   = String(now.getFullYear());

    // Pick template
    const templateFile = TEMPLATE_MAP[loai_phuc_loi];
    if (!templateFile) {
      return res.status(400).json({ error: 'Loại phúc lợi không hợp lệ.' });
    }
    const templatePath = path.join(__dirname, templateFile);

    // Build template data
    const data = {
      nguoi_de_nghi,
      chuc_vu_de_nghi: chuc_vu_de_nghi || '',
      ten_cbcnv,
      ma_nv,
      chuc_danh:  chuc_danh  || '',
      don_vi:     don_vi     || '',
      ngay,
      thang,
      nam,
      su_kien: SU_KIEN_LABEL[loai_phuc_loi] || '',
    };

    // Fill docx
    const docxBuffer = fillDocx(templatePath, data);

    // Convert images to PDF (if any)
    let pdfBuffer = null;
    if (req.files && req.files.length > 0) {
      const imageBuffers = req.files.map(f => f.buffer);
      pdfBuffer = await imagesToPdf(imageBuffers);
    }

    // Upload to SharePoint
    await luuHoSo(ten_cbcnv, ma_nv, ngay, thang, nam, docxBuffer, pdfBuffer);

    res.json({ success: true });
  } catch (err) {
    console.error('Submit error:', err);
    res.status(500).json({ error: 'Có lỗi xảy ra. Vui lòng thử lại.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
