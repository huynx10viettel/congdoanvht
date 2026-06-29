import { ConfidentialClientApplication } from '@azure/msal-node';

// ──────────────────────────────────────────────
// MSAL setup (app-only auth)
// ──────────────────────────────────────────────
const cca = new ConfidentialClientApplication({
  auth: {
    clientId:     process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    authority:    `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
  },
});

async function getToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  return result.accessToken;
}

// ──────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────
const SITE_ID  = process.env.SHAREPOINT_SITE_ID; // optional
const DRIVE_ID = process.env.GRAPH_DRIVE_ID;      // matches Render env var
const BASE     = 'https://graph.microsoft.com/v1.0';

// Site-based drive URL — không phụ thuộc vào DRIVE_ID format.
// Graph API chấp nhận /sites/{siteId}/drive thay cho /drives/{driveId}.
const DRIVE_BASE = SITE_ID
  ? `${BASE}/sites/${SITE_ID}/drive`
  : `${BASE}/drives/${DRIVE_ID}`;

/**
 * Upload một file buffer vào SharePoint.
 * Trả về { webUrl, driveId, itemId }.
 */
async function uploadBuffer({ folderPath, filename, buffer, contentType = 'application/octet-stream' }) {
  const token = await getToken();
  const url   = `${DRIVE_BASE}/root:/${folderPath}/${filename}:/content`;

  const res = await fetch(url, {
    method:  'PUT',
    headers: {
      Authorization:  `Bearer ${token}`,
      'Content-Type': contentType,
    },
    body: buffer,
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph upload failed ${res.status}: ${text}`);
  }

  const item = await res.json();
  return {
    webUrl:  item.webUrl,
    driveId: item.parentReference?.driveId ?? DRIVE_ID,
    itemId:  item.id,
  };
}

// ──────────────────────────────────────────────
// luuHoSo — Lưu docx + ảnh PDF lên SharePoint
// Trả về { webUrl, driveId, docxItemId }
// ──────────────────────────────────────────────
export async function luuHoSo({ folderName, docxBuffer, docxName, imagesPdfBuffer }) {
  const uploadResults = await Promise.all([
    uploadBuffer({
      folderPath:  folderName,
      filename:    docxName,
      buffer:      docxBuffer,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
    imagesPdfBuffer
      ? uploadBuffer({
          folderPath:  folderName,
          filename:    'chung-tu.pdf',
          buffer:      imagesPdfBuffer,
          contentType: 'application/pdf',
        })
      : null,
  ]);

  const docxResult = uploadResults[0];
  return {
    webUrl:      docxResult.webUrl,
    driveId:     docxResult.driveId,
    docxItemId:  docxResult.itemId,
  };
}

// ──────────────────────────────────────────────
// convertDocxToPdf — Dùng Graph API convert file
// Word đã upload sang PDF. Trả về PDF Buffer.
// ──────────────────────────────────────────────
export async function convertDocxToPdf(driveId, itemId) {
  const token = await getToken();
  const url   = `${BASE}/drives/${driveId}/items/${itemId}/content?format=pdf`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    redirect: 'follow',
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph PDF convert failed ${res.status}: ${text}`);
  }

  return Buffer.from(await res.arrayBuffer());
}

// ──────────────────────────────────────────────
// timHoSo — Tra cứu hồ sơ theo mã nhân viên
// Tìm folder có tên dạng: TenNV_maNV_ngay-thang-nam
// Trả về mảng { tenNV, maNV, ngayNop, webUrl }
// ──────────────────────────────────────────────
export async function timHoSo(maNV) {
  const token = await getToken();

  // 1. Lấy danh sách folder tháng ở root (dạng Thang-MM-YYYY)
  const rootUrl = `${DRIVE_BASE}/root/children?$select=name,folder&$top=200`;
  const rootRes = await fetch(rootUrl, { headers: { Authorization: `Bearer ${token}` } });
  if (!rootRes.ok) {
    const text = await rootRes.text();
    throw new Error(`Graph list root failed ${rootRes.status}: ${text}`);
  }
  const rootData  = await rootRes.json();
  const monthFolders = (rootData.value || []).filter(
    item => item.folder && /^Thang-\d{2}-\d{4}$/.test(item.name)
  );

  // 2. Song song: tìm submission folder khớp _maNV_ trong mỗi folder tháng
  const pattern   = new RegExp(`_${maNV}_`);
  const allMatches = [];

  await Promise.all(monthFolders.map(async mf => {
    const url = `${DRIVE_BASE}/root:/${mf.name}:/children?$select=name,webUrl,folder,createdDateTime&$top=200`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) return;
    const data = await res.json();
    (data.value || [])
      .filter(item => item.folder && pattern.test(item.name))
      .forEach(item => {
        const parts = item.name.split('_');
        const tenNV = parts.slice(0, parts.length - 2).join('_');
        const date  = parts[parts.length - 1];
        allMatches.push({
          tenNV,
          maNV,
          ngayNop:         date,
          tenFolder:       `${mf.name}/${item.name}`,
          webUrl:          item.webUrl,
          createdDateTime: item.createdDateTime,
        });
      });
  }));

  return allMatches.sort((a, b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));
}

// ──────────────────────────────────────────────
// uploadPdf — Upload file PDF đã annotate
// ──────────────────────────────────────────────
export async function uploadPdf({ folderName, pdfBuffer, pdfName }) {
  const result = await uploadBuffer({
    folderPath:  folderName,
    filename:    pdfName,
    buffer:      pdfBuffer,
    contentType: 'application/pdf',
  });
  return { webUrl: result.webUrl };
}

// ──────────────────────────────────────────────
// uploadJson — Upload JSON metadata vào folder
// ──────────────────────────────────────────────
export async function uploadJson({ folderName, data, filename = 'meta.json' }) {
  await uploadBuffer({
    folderPath:  folderName,
    filename,
    buffer:      Buffer.from(JSON.stringify(data, null, 2)),
    contentType: 'application/json',
  });
}

// ──────────────────────────────────────────────
// readMetaFromFolder — Đọc meta.json từ folder
// Trả về object hoặc null nếu không tìm thấy
// ──────────────────────────────────────────────
async function readMetaFromFolder(folderName) {
  const token = await getToken();
  const url = `${DRIVE_BASE}/root:/${folderName}/meta.json:/content`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) return null;
  try {
    const text = await res.text();
    return JSON.parse(text);
  } catch {
    return null;
  }
}

// ──────────────────────────────────────────────
// listSubmissionsByMonth — Liệt kê & đọc metadata
// tất cả hồ sơ trong tháng/năm chỉ định.
// Trả về mảng meta objects (có thể thiếu trường
// nếu submission cũ không có meta.json)
// ──────────────────────────────────────────────
export async function listSubmissionsByMonth(thang, nam) {
  const token      = await getToken();
  const mm         = String(thang).padStart(2, '0');
  const yyyy       = String(nam);
  const monthFolder = `Thang-${mm}-${yyyy}`;

  // List children của folder tháng (Thang-MM-YYYY)
  const url = `${DRIVE_BASE}/root:/${monthFolder}:/children?$select=name,webUrl,folder,createdDateTime&$top=500`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });

  if (!res.ok) {
    if (res.status === 404) return []; // folder tháng chưa tồn tại
    const text = await res.text();
    throw new Error(`Graph list failed ${res.status}: ${text}`);
  }

  const body     = await res.json();
  const matching = (body.value || []).filter(item => item.folder);

  if (matching.length === 0) return [];

  // Đọc meta.json từng folder song song
  const submissions = await Promise.all(
    matching.map(async folder => {
      const meta = await readMetaFromFolder(`${monthFolder}/${folder.name}`);
      if (meta) return meta;
      // Fallback: tái tạo từ tên folder
      const parts = folder.name.split('_');
      return {
        ten_cbcnv:     parts.slice(0, parts.length - 2).join('_'),
        ma_nv:         parts[parts.length - 2],
        ngay_nop:      parts[parts.length - 1],
        loai_phuc_loi: '',
        don_vi:        '',
        chuc_danh:     '',
        qh_than_nhan:  '',
        ten_nguoi_than:'',
        benh_vien:     '',
      };
    })
  );

  return submissions.sort((a, b) => (a.ten_cbcnv || '').localeCompare(b.ten_cbcnv || '', 'vi'));
}
