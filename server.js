import "dotenv/config";
import express from "express";
import multer from "multer";
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

import { fillDocx } from "./docxFiller.js";
import { imagesToPdf } from "./imagesToPdf.js";
import { soThanhChu, chuanHoaTien, dinhDangTien } from "./soThanhChu.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATE_PATH =
  process.env.TEMPLATE_PATH || path.join(__dirname, "mau-don-phuc-loi.docx");

// Chế độ chạy thử cục bộ: nếu chưa cấu hình Graph, lưu file ra thư mục ./output
const GRAPH_CONFIGURED =
  process.env.TENANT_ID &&
  process.env.CLIENT_ID &&
  process.env.CLIENT_SECRET &&
  (process.env.GRAPH_DRIVE_ID || process.env.GRAPH_SITE_URL);

const app = express();
// Phục vụ trang form (index.html nằm cùng thư mục, mọi CSS/JS đã nhúng sẵn bên trong)
app.get("/", (_req, res) => res.sendFile(path.join(__dirname, "index.html")));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024, files: 20 }, // tối đa 10MB/ảnh, 20 ảnh
});

const slug = (s) =>
  (s || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);

app.post("/api/submit", upload.array("images", 20), async (req, res) => {
  try {
    const f = req.body;
    const now = new Date();

    // Tính tiền + đọc số thành chữ (server tự tính tổng, người dùng chỉ nhập số)
    const tienPL = chuanHoaTien(f.tien_phuc_loi);
    const tienCD = chuanHoaTien(f.tien_cong_doan);
    const tong = tienPL + tienCD;

    // Chuẩn hóa dữ liệu cho template (key trùng placeholder trong .docx)
    const data = {
      nguoi_de_nghi: f.nguoi_de_nghi || "",
      chuc_vu_de_nghi: f.chuc_vu_de_nghi || "",
      tien_phuc_loi: dinhDangTien(tienPL),
      tien_phuc_loi_chu: soThanhChu(tienPL),
      tien_cong_doan: dinhDangTien(tienCD),
      tien_cong_doan_chu: soThanhChu(tienCD),
      tong_cong: dinhDangTien(tong),
      tong_cong_chu: soThanhChu(tong),
      noi_dung_chi: f.noi_dung_chi || "",
      ten_cbcnv: f.ten_cbcnv || "",
      ma_nv: f.ma_nv || "",
      chuc_danh: f.chuc_danh || "",
      don_vi: f.don_vi || "",
      ngay: f.ngay || String(now.getDate()),
      thang: f.thang || String(now.getMonth() + 1),
      nam: f.nam || String(now.getFullYear()),
    };

    if (!data.nguoi_de_nghi) return res.status(400).json({ ok: false, error: "Thiếu họ tên người đề nghị" });

    // 1) Điền template Word
    const docxBuffer = fillDocx(TEMPLATE_PATH, data);

    // 2) Gộp ảnh thành PDF
    const images = (req.files || []).map((x) => ({ buffer: x.buffer, mimetype: x.mimetype }));
    const imagesPdfBuffer = images.length ? await imagesToPdf(images) : null;

    const folderName = `${slug(data.nguoi_de_nghi)}-${now
      .toISOString()
      .slice(0, 10)}-${Date.now().toString().slice(-4)}`;
    const docxName = `Giay-de-nghi-chi-${slug(data.nguoi_de_nghi)}.docx`;

    if (!GRAPH_CONFIGURED) {
      // Chế độ thử cục bộ: lưu ra ./output (không convert PDF vì Graph làm việc đó)
      const outDir = path.join(__dirname, "output", folderName);
      fs.mkdirSync(outDir, { recursive: true });
      fs.writeFileSync(path.join(outDir, docxName), docxBuffer);
      if (imagesPdfBuffer) fs.writeFileSync(path.join(outDir, "anh-dinh-kem.pdf"), imagesPdfBuffer);
      return res.json({
        ok: true,
        mode: "local",
        message: "Đã lưu cục bộ (chưa cấu hình Microsoft Graph nên chưa convert PDF/lưu OneDrive).",
        folder: outDir,
        files: [docxName, imagesPdfBuffer ? "anh-dinh-kem.pdf" : null].filter(Boolean),
      });
    }

    // 3) Convert PDF + lưu OneDrive/SharePoint qua Microsoft Graph
    const { luuHoSo } = await import("./graph.js");
    const result = await luuHoSo({ folderName, docxBuffer, docxName, imagesPdfBuffer });

    res.json({ ok: true, mode: "graph", message: "Đã nộp và lưu hồ sơ thành công.", ...result });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: String(err.message || err) });
  }
});

app.get("/api/health", (_req, res) =>
  res.json({ ok: true, graph: !!GRAPH_CONFIGURED, template: fs.existsSync(TEMPLATE_PATH) })
);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server chạy tại http://localhost:${PORT}  (Graph: ${GRAPH_CONFIGURED ? "ON" : "OFF - chế độ thử cục bộ"})`);
});
