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
const SITE_ID  = process.env.SHAREPOINT_SITE_ID;
const DRIVE_ID = process.env.SHAREPOINT_DRIVE_ID;
const BASE     = 'https://graph.microsoft.com/v1.0';

/**
 * Upload một file buffer vào SharePoint.
 * Trả về { webUrl, driveId, itemId }.
 */
async function uploadBuffer({ folderPath, filename, buffer, contentType = 'application/octet-stream' }) {
  const token = await getToken();
  const url   = `${BASE}/drives/${DRIVE_ID}/root:/${folderPath}/${filename}:/content`;

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
