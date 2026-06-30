import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { fillDocx } from './docxFiller.js';
import { imagesToPdf } from './imagesToPdf.js';
import { luuHoSo, convertDocxToPdf, uploadPdf, uploadJson, uploadExcel, uploadToFolder, downloadFile, listFilesInFolder, timHoSo, listSubmissionsByMonth, listSubmissionsByDateRange } from './graph.js';
import { generateExcel, generateExcelFromTemplate } from './exportExcel.js';
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

    // 2) Trả response ngay — client không cần chờ PDF
    res.json({ success: true, webUrl });

    // 3) Convert docx → PDF + annotate + upload (fire-and-forget, chạy ngầm)
    const pdfName = `De-nghi-phuc-loi_${ten_cbcnv}_${ma_nv}_${ngay}-${thang}-${nam}.pdf`;
    convertDocxToPdf(driveId, docxItemId)
      .then(rawPdf  => addCommentsToPdf(rawPdf))
      .then(annPdf  => uploadPdf({ folderName, pdfBuffer: annPdf, pdfName }))
      .then(r       => console.log('[pdf] uploaded:', r.webUrl))
      .catch(e      => console.error('[pdf] error (non-fatal):', e.message));
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
// GET /api/admin/export?startDate=YYYY-MM-DD&endDate=YYYY-MM-DD
// Header: X-Admin-Code
// Folder lưu lấy theo tháng của endDate.
// Trả về JSON { success, webUrl, folderWebUrl, filename, count }
// ──────────────────────────────────────────────
app.get('/api/admin/export', async (req, res) => {
  res.setHeader('Content-Type', 'application/json');   // luôn trả JSON

  const code = req.headers['x-admin-code'];
  if (!code || code !== ADMIN_CODE) {
    return res.status(401).json({ error: 'Mã admin không đúng.' });
  }

  const { startDate, endDate } = req.query;
  const dateRe = /^\d{4}-\d{2}-\d{2}$/;
  if (!startDate || !endDate || !dateRe.test(startDate) || !dateRe.test(endDate)) {
    return res.status(400).json({ error: 'Ngày không hợp lệ. Định dạng: YYYY-MM-DD.' });
  }
  if (new Date(startDate) > new Date(endDate)) {
    return res.status(400).json({ error: 'Ngày bắt đầu phải ≤ ngày kết thúc.' });
  }

  try {
    // Folder & tiêu đề lấy tháng/năm của ngày kết thúc
    const [yyyy, mm] = endDate.split('-');

    const submissions = await listSubmissionsByDateRange(startDate, endDate);

    const templatePath = path.join(__dirname, 'template_export.xlsx');
    const excelBuffer  = await generateExcelFromTemplate(
      submissions, parseInt(mm), parseInt(yyyy), templatePath
    );

    // Tên file + export folder (subfolder riêng chứa tất cả file)
    const now = new Date();
    const ts  = `${String(now.getDate()).padStart(2,'0')}${String(now.getMonth()+1).padStart(2,'0')}${now.getFullYear()}-${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
    const baseName     = `Tong-hop-phuc-loi-T${mm}-${yyyy}_${ts}`;
    const filename     = `${baseName}.xlsx`;
    const monthFolder  = `Thang-${mm}-${yyyy}`;
    const exportFolder = `${monthFolder}/Export_${startDate}_${endDate}_${ts}`;

    // 1. Upload Excel vào export folder (cần driveId + itemId để convert)
    const excelResult = await uploadToFolder({
      folderPath:  exportFolder,
      filename,
      buffer:      excelBuffer,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const { webUrl } = excelResult;

    // 2. Convert Excel → PDF (Graph API hỗ trợ xlsx → pdf)
    let pdfUploaded = false;
    try {
      const pdfBuffer     = await convertDocxToPdf(excelResult.driveId, excelResult.itemId);
      const annotatedPdf  = await addCommentsToPdf(pdfBuffer);
      await uploadToFolder({
        folderPath:  exportFolder,
        filename:    `${baseName}.pdf`,
        buffer:      annotatedPdf,
        contentType: 'application/pdf',
      });
      pdfUploaded = true;
    } catch (pdfErr) {
      console.warn('[export] Excel→PDF failed (non-fatal):', pdfErr.message);
    }

    // 3. Copy file PDF hồ sơ từng submission vào export folder
    let copied = 0;
    for (const s of submissions) {
      if (!s._folderPath) continue;
      try {
        const files = await listFilesInFolder(s._folderPath);
        for (const fname of files) {
          if (!fname.toLowerCase().endsWith('.pdf')) continue;
          const buf = await downloadFile(`${s._folderPath}/${fname}`);
          if (!buf) continue;
          // Đặt tên: HoSo_TenNV_MaNV_TenFile.pdf
          const safe    = (s.ten_cbcnv || 'unknown').replace(/[/\\:*?"<>|]/g, '_');
          const newName = `HoSo_${safe}_${s.ma_nv || ''}_${fname}`;
          await uploadToFolder({
            folderPath:  exportFolder,
            filename:    newName,
            buffer:      buf,
            contentType: 'application/pdf',
          });
          copied++;
        }
      } catch (copyErr) {
        console.warn(`[export] copy ${s._folderPath} failed:`, copyErr.message);
      }
    }

    console.log(`[export] ${submissions.length} submissions, ${copied} PDFs copied, pdfUploaded=${pdfUploaded}`);

    res.json({ success: true, webUrl, folderWebUrl: 'https://w0tks.sharepoint.com/sites/PLCDVHT', filename, count: submissions.length });
  } catch (err) {
    console.error('Admin export error:', err);
    res.status(500).json({ error: err.message || 'Có lỗi khi kết xuất. Vui lòng thử lại.' });
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
