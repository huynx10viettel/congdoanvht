import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { fillDocx } from './docxFiller.js';
import { imagesToPdf } from './imagesToPdf.js';
import { luuHoSo } from './graph.js';

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
      ngay_ra_vien,
      ngay_cuoi,
      so_giay_ket_hon,
      ten_con,
      so_giay_khai_sinh,
      ghi_chu_1,
      ghi_chu_2,
      ghi_chu_3,
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
      ngay_ra_vien:      ngay_ra_vien      || '',
      ngay_cuoi:         ngay_cuoi         || '',
      so_giay_ket_hon:   so_giay_ket_hon   || '',
      ten_con:           ten_con           || '',
      so_giay_khai_sinh: so_giay_khai_sinh || '',
      ghi_chu_1:         ghi_chu_1         || '',
      ghi_chu_2:         ghi_chu_2         || '',
      ghi_chu_3:         ghi_chu_3         || '',
    };

    // Fill docx
    const docxBuffer = fillDocx(templatePath, data);
    const docxName   = `De-nghi-phuc-loi_${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}.docx`;

    // Ảnh → PDF
    let imagesPdfBuffer = null;
    if (req.files && req.files.length > 0) {
      imagesPdfBuffer = await imagesToPdf(req.files);
    }

    // Lưu lên SharePoint
    const folderName = `${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}`;
    const result = await luuHoSo({ folderName, docxBuffer, docxName, imagesPdfBuffer });

    res.json({ success: true, webUrl: result.webUrl });
  } catch (err) {
    console.error('Submit error:', err);
    res.status(500).json({ error: 'Có lỗi xảy ra. Vui lòng thử lại.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server chạy tại http://localhost:${PORT}`));
