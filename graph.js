// Tích hợp Microsoft Graph (app-only / client credentials) để:
//  - Tạo folder theo từng yêu cầu
//  - Upload file
//  - Convert .docx -> PDF (Graph hỗ trợ sẵn ?format=pdf, miễn phí với license M365)
//  - Lưu vào thư viện tài liệu SharePoint hoặc OneDrive
//
// Lưu ý: luồng app-only chỉ hoạt động với SharePoint/OneDrive for Business.
import { ConfidentialClientApplication } from "@azure/msal-node";

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const DRIVE_ID = process.env.GRAPH_DRIVE_ID; // ID của Document Library (drive) trên SharePoint
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
    headers: {
      Authorization: `Bearer ${token}`,
      ...headers,
    },
    body,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${method} ${path} -> ${res.status}: ${text}`);
  }
  if (raw) return res; // trả response thô (để đọc binary)
  if (res.status === 204) return null;
  return res.json();
}

function encodePath(p) {
  return p
    .split("/")
    .filter(Boolean)
    .map((s) => encodeURIComponent(s))
    .join("/");
}

// Tạo folder con (idempotent). parentPath rỗng = gốc thư viện.
async function ensureFolder(token, name, parentPath = "") {
  const parent = parentPath
    ? `/drives/${DRIVE_ID}/root:/${encodePath(parentPath)}:/children`
    : `/drives/${DRIVE_ID}/root/children`;
  return graph(parent, {
    method: "POST",
    token,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "replace",
    }),
  });
}

// Upload file (<4MB là đủ cho đơn + ảnh nén; dùng simple upload)
async function uploadFile(token, folderPath, fileName, buffer, contentType) {
  const path = `/drives/${DRIVE_ID}/root:/${encodePath(folderPath)}/${encodeURIComponent(fileName)}:/content`;
  return graph(path, {
    method: "PUT",
    token,
    headers: { "Content-Type": contentType },
    body: buffer,
  });
}

// Convert item .docx sang PDF bằng Graph, trả Buffer PDF
async function downloadAsPdf(token, itemId) {
  const res = await graph(`/drives/${DRIVE_ID}/items/${itemId}/content?format=pdf`, {
    token,
    raw: true,
  });
  const arr = await res.arrayBuffer();
  return Buffer.from(arr);
}

/**
 * Lưu trọn bộ hồ sơ 1 yêu cầu phúc lợi vào SharePoint/OneDrive.
 * @returns {Promise<{folder: string, webUrl: string, files: string[]}>}
 */
export async function luuHoSo({ folderName, docxBuffer, docxName, imagesPdfBuffer }) {
  const token = await getToken();

  // Tạo folder cha (nếu chưa có) và folder riêng cho yêu cầu này
  await ensureFolder(token, BASE_FOLDER, "");
  const folder = await ensureFolder(token, folderName, BASE_FOLDER);
  const folderPath = `${BASE_FOLDER}/${folderName}`;
  const files = [];

  // 1) Upload .docx đã điền
  const docxItem = await uploadFile(
    token,
    folderPath,
    docxName,
    docxBuffer,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );
  files.push(docxName);

  // 2) Convert .docx -> PDF (qua Graph) rồi upload PDF
  const pdfBuffer = await downloadAsPdf(token, docxItem.id);
  const pdfName = docxName.replace(/\.docx$/i, ".pdf");
  await uploadFile(token, folderPath, pdfName, pdfBuffer, "application/pdf");
  files.push(pdfName);

  // 3) Upload PDF ảnh (nếu có)
  if (imagesPdfBuffer && imagesPdfBuffer.length) {
    const imgPdfName = "anh-dinh-kem.pdf";
    await uploadFile(token, folderPath, imgPdfName, imagesPdfBuffer, "application/pdf");
    files.push(imgPdfName);
  }

  return { folder: folderPath, webUrl: folder.webUrl, files };
}
