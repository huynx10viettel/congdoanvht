// Tích hợp Microsoft Graph (app-only / client credentials) để:
//  - Tạo folder theo từng yêu cầu
//  - Upload file
//  - Convert .docx -> PDF (Graph hỗ trợ sẵn ?format=pdf, miễn phí với license M365)
//  - Lưu vào thư viện tài liệu SharePoint
//
// Lưu ý: luồng app-only chỉ hoạt động với SharePoint / OneDrive for Business.
//
// Cấu hình (env): TENANT_ID, CLIENT_ID, CLIENT_SECRET và MỘT trong hai cách trỏ kho lưu:
//   Cách dễ:  GRAPH_SITE_URL = https://congty.sharepoint.com/sites/CongDoan
//             (tùy chọn GRAPH_LIBRARY_NAME nếu muốn 1 thư viện khác mặc định)
//   Cách thủ công: GRAPH_DRIVE_ID = <id drive lấy sẵn>
import { ConfidentialClientApplication } from "@azure/msal-node";

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const DRIVE_ID_ENV = process.env.GRAPH_DRIVE_ID; // (tùy chọn) chỉ định trực tiếp
const SITE_URL = process.env.GRAPH_SITE_URL; // (khuyên dùng) link site SharePoint
const LIBRARY_NAME = process.env.GRAPH_LIBRARY_NAME; // (tùy chọn) tên thư viện tài liệu
const BASE_FOLDER = (process.env.TARGET_FOLDER || "PhucLoiCongDoan").replace(/^\/+|\/+$/g, "");

const GRAPH = "https://graph.microsoft.com/v1.0";

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET,
  },
});

async function getToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  if (!result?.accessToken) throw new Error("Không lấy được access token từ Microsoft Graph");
  return result.accessToken;
}

async function graph(path, { method = "GET", token, body, headers = {}, raw = false } = {}) {
  const res = await fetch(path.startsWith("http") ? path : `${GRAPH}${path}`, {
    method,
    headers: { Authorization: `Bearer ${token}`, ...headers },
    body,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${method} ${path} -> ${res.status}: ${text}`);
  }
  if (raw) return res;
  if (res.status === 204) return null;
  return res.json();
}

function encodePath(p) {
  return p.split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

// Tự xác định Drive ID: ưu tiên GRAPH_DRIVE_ID; nếu không có thì phân giải từ GRAPH_SITE_URL.
let _driveIdCache = null;
async function resolveDriveId(token) {
  if (_driveIdCache) return _driveIdCache;
  if (DRIVE_ID_ENV) {
    _driveIdCache = DRIVE_ID_ENV;
    return _driveIdCache;
  }
  if (!SITE_URL) throw new Error("Thiếu cấu hình: cần GRAPH_SITE_URL hoặc GRAPH_DRIVE_ID");

  const u = new URL(SITE_URL);
  const hostname = u.hostname; // vd: congty.sharepoint.com
  const sitePath = u.pathname.replace(/\/+$/g, ""); // vd: /sites/CongDoan
  const site = await graph(`/sites/${hostname}:${sitePath}`, { token });

  if (LIBRARY_NAME) {
    const drives = await graph(`/sites/${site.id}/drives`, { token });
    const found = (drives.value || []).find((d) => d.name === LIBRARY_NAME);
    if (!found) throw new Error(`Không tìm thấy thư viện "${LIBRARY_NAME}" trong site`);
    _driveIdCache = found.id;
  } else {
    const drive = await graph(`/sites/${site.id}/drive`, { token }); // thư viện mặc định
    _driveIdCache = drive.id;
  }
  return _driveIdCache;
}

async function ensureFolder(token, driveId, name, parentPath = "") {
  const parent = parentPath
    ? `/drives/${driveId}/root:/${encodePath(parentPath)}:/children`
    : `/drives/${driveId}/root/children`;
  return graph(parent, {
    method: "POST",
    token,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "replace" }),
  });
}

async function uploadFile(token, driveId, folderPath, fileName, buffer, contentType) {
  const path = `/drives/${driveId}/root:/${encodePath(folderPath)}/${encodeURIComponent(fileName)}:/content`;
  return graph(path, { method: "PUT", token, headers: { "Content-Type": contentType }, body: buffer });
}

async function downloadAsPdf(token, driveId, itemId) {
  const res = await graph(`/drives/${driveId}/items/${itemId}/content?format=pdf`, { token, raw: true });
  const arr = await res.arrayBuffer();
  return Buffer.from(arr);
}

/**
 * Lưu trọn bộ hồ sơ 1 yêu cầu phúc lợi vào SharePoint.
 * @returns {Promise<{folder: string, webUrl: string, files: string[]}>}
 */
export async function luuHoSo({ folderName, docxBuffer, docxName, imagesPdfBuffer }) {
  const token = await getToken();
  const driveId = await resolveDriveId(token);

  await ensureFolder(token, driveId, BASE_FOLDER, "");
  const folder = await ensureFolder(token, driveId, folderName, BASE_FOLDER);
  const folderPath = `${BASE_FOLDER}/${folderName}`;
  const files = [];

  // 1) Upload .docx đã điền
  const docxItem = await uploadFile(
    token, driveId, folderPath, docxName, docxBuffer,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );
  files.push(docxName);

  // 2) Convert .docx -> PDF (qua Graph) rồi upload PDF
  const pdfBuffer = await downloadAsPdf(token, driveId, docxItem.id);
  const pdfName = docxName.replace(/\.docx$/i, ".pdf");
  await uploadFile(token, driveId, folderPath, pdfName, pdfBuffer, "application/pdf");
  files.push(pdfName);

  // 3) Upload PDF ảnh (nếu có)
  if (imagesPdfBuffer && imagesPdfBuffer.length) {
    await uploadFile(token, driveId, folderPath, "anh-dinh-kem.pdf", imagesPdfBuffer, "application/pdf");
    files.push("anh-dinh-kem.pdf");
  }

  return { folder: folderPath, webUrl: folder.webUrl, files };
}
