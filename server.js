import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { fillDocx } from './docxFiller.js';
import { imagesToPdf } from './imagesToPdf.js';
import { luuHoSo, convertDocxToPdf, uploadPdf } from './graph.js';
import { addCommentsToPdf } from './pdfAnnotator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// ──────────────────────────────────────────────
// Template map theo loai_phuc_loi
// ──────────────────────────────────────────────
const TEMPLATE_MAP = {
  'tham-hoi-om-cbnv':    'De nghi thanh toan tien cong doan, phuc loi_temp_om.docx',
  'tham-hoi-om-thannhan':'De nghi thanh toan tien cong doan, phuc loi_temp_om.docx',
  'ket-hon':             'De nghi thanh toan tien cong doan, phuc loi_temp_chucmung.docx',
  'sinh-con':            'De nghi thanh toan tien cong doan, phuc loi_temp_chucmung.docx',
  'tham-vieng-cbcnv':    'De nghi thanh toan tien cong doan, phuc loi_temp_CBNVtutran.docx',
  'tham-vieng-thannhan': 'De nghi thanh toan tien cong doan, phuc loi_temp_thannhanCBNVtutran.docx',
};

const SU_KIEN_LABEL = {
  'ket-hon':  'kết hôn',
  'sinh-con': 'sinh con',
};

// ──────────────────────────────────────────────
// Static
// ──────────────────────────────────────────────
app.use(express.static(__dirname));

// ──────────────────────────────────────────────
// POST /api/submit
// ──────────────────────────────────────────────
app.post('/api/submit', upload.array('images'), async (req, res) => {
  try {
    const {
      chuc_vu_de_nghi,
      ten_cbcnv,
      ma_nv,
      chuc_danh,
      don_vi,
      loai_phuc_loi,
      ten_nguoi_than,
      benh_vien,
      ngay_cuoi,
      so_giay_ket_hon,
      ten_con,
      so_giay_khai_sinh,
    } = req.body;

    // nguoi_de_nghi falls back to chuc_vu_de_nghi
    const nguoi_de_nghi = req.body.nguoi_de_nghi || chuc_vu_de_nghi || '';

    if (!ten_cbcnv || !ma_nv || !loai_phuc_loi) {
      return res.status(400).json({ error: 'Thiếu thông tin bắt buộc.' });
    }

    const templateFile = TEMPLATE_MAP[loai_phuc_loi];
    if (!templateFile) {
      return res.status(400).json({ error: 'Loại phúc lợi không hợp lệ.' });
    }

    const now   = new Date();
    const ngay  = String(now.getDate()).padStart(2, '0');
    const thang = String(now.getMonth() + 1).padStart(2, '0');
    const nam   = String(now.getFullYear());

    const templatePath = path.join(__dirname, templateFile);

    const data = {
      nguoi_de_nghi,
      chuc_vu_de_nghi:   chuc_vu_de_nghi   || '',
      ten_cbcnv,
      ma_nv,
      chuc_danh:         chuc_danh         || '',
      don_vi:            don_vi            || '',
      ngay,
      thang,
      nam,
      su_kien:           SU_KIEN_LABEL[loai_phuc_loi] || '',
      ten_nguoi_than:    ten_nguoi_than    || '',
      benh_vien:         benh_vien         || '',
      ngay_cuoi:         ngay_cuoi         || '',
      so_giay_ket_hon:   so_giay_ket_hon   || '',
      ten_con:           ten_con           || '',
      so_giay_khai_sinh: so_giay_khai_sinh || '',
      ghi_chu_1:         '',
      ghi_chu_2:         '',
      ghi_chu_3:         '',
    };

    // Fill docx
    const docxBuffer = fillDocx(templatePath, data);
    const folderName = `${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}`;
    const docxName   = `De-nghi-phuc-loi_${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}.docx`;

    // Ảnh / PDF đính kèm → merge PDF
    let imagesPdfBuffer = null;
    if (req.files && req.files.length > 0) {
      imagesPdfBuffer = await imagesToPdf(req.files);
    }

    // 1) Upload docx (+ ảnh PDF) lên SharePoint
    const { webUrl, driveId, docxItemId } = await luuHoSo({
      folderName,
      docxBuffer,
      docxName,
      imagesPdfBuffer,
    });

    // 2) Graph API convert docx → PDF
    let pdfWebUrl = null;
    try {
      const rawPdf       = await convertDocxToPdf(driveId, docxItemId);
      // 3) Add 3 FreeText comment annotations
      const annotatedPdf = await addCommentsToPdf(rawPdf);
      // 4) Upload PDF đã annotate vào cùng folder
      const pdfName      = `De-nghi-phuc-loi_${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}.pdf`;
      const pdfResult    = await uploadPdf({ folderName, pdfBuffer: annotatedPdf, pdfName });
      pdfWebUrl = pdfResult.webUrl;
    } catch (pdfErr) {
      // Không chặn luồng chính — docx vẫn đã được lưu thành công
      console.error('PDF convert/annotate error (non-fatal):', pdfErr.message);
    }

    res.json({ success: true, webUrl, pdfWebUrl });
  } catch (err) {
    console.error('Submit error:', err);
    res.status(500).json({ error: 'Có lỗi xảy ra. Vui lòng thử lại.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server chạy tại http://localhost:${PORT}`));
