import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { fillDocx } from './docxFiller.js';
import { imagesToPdf } from './imagesToPdf.js';
import { luuHoSo, convertDocxToPdf, uploadPdf, uploadJson, timHoSo, listSubmissionsByMonth } from './graph.js';
import { generateExcel } from './exportExcel.js';
import { addCommentsToPdf } from './pdfAnnotator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// ──────────────────────────────────────────────
// Template map theo loai_phuc_loi
// ──────────────────────────────────────────────
const TEMPLATE_MAP = {
  'tham-hoi-om-cbnv':    'De nghi thanh toan tien cong doan, phuc loi_CBVN_Om.docx',
  'tham-hoi-om-thannhan':'De nghi thanh toan tien cong doan, phuc loi_TN_CBNV_om.docx',
  'ket-hon':             'De nghi thanh toan tien cong doan, phuc loi_CBNV_KetHon.docx',
  'sinh-con':            'De nghi thanh toan tien cong doan, phuc loi_CBNV_SinhCon.docx',
  'tham-vieng-cbnv':    'De nghi thanh toan tien cong doan, phuc loi_CBNV_Tt.docx',
  'tham-vieng-thannhan': 'De nghi thanh toan tien cong doan, phuc loi_TN_CBNV_Tt.docx',
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
      qh_than_nhan,
      benh_vien,
      ngay_cuoi,
      so_giay_ket_hon,
      ten_con,
      so_giay_khai_sinh,
      loai_giay_to,
      so_giay_to,
      ngay_giay_to,
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

    // Cấu trúc thư mục: Thang-MM-YYYY / TenNV_MaNV_dd-mm-yyyy
    const monthFolder = `Thang-${thang}-${nam}`;

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
      qh_than_nhan:      qh_than_nhan      || '',
      ten_nguoi_than:    ten_nguoi_than    || '',
      benh_vien:         benh_vien         || '',
      ngay_cuoi:         ngay_cuoi         || '',
      so_giay_ket_hon:   (loai_phuc_loi === 'ket-hon'  ? so_giay_to : so_giay_ket_hon)  || '',
      ten_con:           ten_con           || '',
      so_giay_khai_sinh: (loai_phuc_loi === 'sinh-con' ? so_giay_to : so_giay_khai_sinh) || '',
      ghi_chu_1:         '',
      ghi_chu_2:         '',
      ghi_chu_3:         '',
    };

    // Fill docx
    const docxBuffer = fillDocx(templatePath, data);
    const subFolder  = `${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}`;
    const folderName = `${monthFolder}/${subFolder}`;   // Thang-MM-YYYY/TenNV_MaNV_dd-mm-yyyy
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

    // Upload meta.json (fire-and-forget, không chặn luồng chính)
    uploadJson({
      folderName,
      data: {
        ten_cbcnv, ma_nv, chuc_danh, don_vi, loai_phuc_loi,
        qh_than_nhan:      qh_than_nhan      || '',
        ten_nguoi_than:    ten_nguoi_than    || '',
        benh_vien:         benh_vien         || '',
        ngay_cuoi:         ngay_cuoi         || '',
        so_giay_ket_hon:   so_giay_ket_hon   || '',
        ten_con:           ten_con           || '',
        so_giay_khai_sinh: so_giay_khai_sinh || '',
        loai_giay_to:      loai_giay_to      || '',
        so_giay_to:        so_giay_to        || '',
        ngay_giay_to:      ngay_giay_to      || '',
        ngay_nop: `${ngay}-${thang}-${nam}`,
      },
    }).catch(e => console.error('meta.json upload failed (non-fatal):', e.message));

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
    res.status(500).json({ error: err.message || 'Có lỗi xảy ra. Vui lòng thử lại.' });
  }
});

// ──────────────────────────────────────────────
// GET /api/lookup?ma_nv=123456
// ──────────────────────────────────────────────
app.get('/api/lookup', async (req, res) => {
  const { ma_nv } = req.query;
  if (!ma_nv || !/^\d{6}$/.test(ma_nv)) {
    return res.status(400).json({ error: 'Mã nhân viên phải gồm đúng 6 chữ số.' });
  }
  try {
    const results = await timHoSo(ma_nv);
    res.json({ success: true, results });
  } catch (err) {
    console.error('Lookup error:', err);
    res.status(500).json({ error: 'Không thể tra cứu. Vui lòng thử lại.' });
  }
});

// ──────────────────────────────────────────────
// GET /api/admin/verify — Kiểm tra mã admin
// ──────────────────────────────────────────────
const ADMIN_CODE = process.env.ADMIN_CODE || 'Hnmuanao';

app.get('/api/admin/verify', (req, res) => {
  const code = req.headers['x-admin-code'];
  if (!code || code !== ADMIN_CODE) {
    return res.status(401).json({ error: 'Mã admin không đúng.' });
  }
  res.json({ ok: true });
});

// ──────────────────────────────────────────────
// GET /api/admin/export?month=MM&year=YYYY
// Header: X-Admin-Code
// ──────────────────────────────────────────────
app.get('/api/admin/export', async (req, res) => {
  const code = req.headers['x-admin-code'];
  if (!code || code !== ADMIN_CODE) {
    return res.status(401).json({ error: 'Mã admin không đúng.' });
  }

  const { month, year } = req.query;
  if (!month || !year || !/^\d{1,2}$/.test(month) || !/^\d{4}$/.test(year)) {
    return res.status(400).json({ error: 'Tháng/năm không hợp lệ.' });
  }

  try {
    const submissions  = await listSubmissionsByMonth(month, year);
    const excelBuffer  = await generateExcel(submissions, month, year);
    const mm = String(month).padStart(2, '0');
    res.setHeader('Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',
      `attachment; filename="danh-sach-yeu-cau-T${mm}-${year}.xlsx"`);
    res.send(excelBuffer);
  } catch (err) {
    console.error('Admin export error:', err);
    res.status(500).json({ error: 'Có lỗi khi kết xuất. Vui lòng thử lại.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server chạy tại http://localhost:${PORT}`);

  // ── Keep-alive: tự ping mỗi 10 phút để tránh Render free tier ngủ ──
  const SELF_URL = process.env.RENDER_EXTERNAL_URL;
  if (SELF_URL) {
    setInterval(() => {
      fetch(`${SELF_URL}/api/ping`)
        .then(() => console.log('[keep-alive] ping OK'))
        .catch(e => console.warn('[keep-alive] ping failed:', e.message));
    }, 10 * 60 * 1000); // 10 phút
  }
});

// Endpoint ping (nhẹ, không có logic)
app.get('/api/ping', (_req, res) => res.json({ ok: true }));
