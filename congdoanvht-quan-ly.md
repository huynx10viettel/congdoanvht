# 📋 Thông tin quản lý Web Phúc Lợi Công Đoàn VHT

> File này chứa toàn bộ thông tin cần thiết để quản lý, vận hành, và bảo trì hệ thống.

---

## 🌐 URL Truy Cập

| Mục | URL |
|-----|-----|
| **Web chính (domain)** | https://congdoanvht.dpdns.org |
| **URL Render trực tiếp** | https://congdoanvht-swq7.onrender.com |

---

## 📦 GitHub

| Mục | Thông tin |
|-----|-----------|
| **Account** | huynx10viettel |
| **Repo** | https://github.com/huynx10viettel/congdoanvht |
| **Branch chính** | main |
| **Auto Deploy** | Có (push lên main → Render tự deploy) |

---

## 🚀 Render (Hosting)

| Mục | Thông tin |
|-----|-----------|
| **Account email** | huynx10.viettel@gmail.com |
| **Dashboard** | https://dashboard.render.com |
| **Service name** | congdoanvht |
| **Service ID** | srv-d8johhk2m8qs739bccog |
| **Plan** | Free (spin down sau 15 phút không dùng, cold start ~50s) |
| **Runtime** | Node.js |
| **Build command** | `npm install` |
| **Start command** | `npm start` |

### Environment Variables (Render)

| Key | Value | Ghi chú |
|-----|-------|---------|
| `TENANT_ID` | `f380bf27-17b2-4901-8f9f-84d2dd0410ee` | Azure AD Tenant |
| `CLIENT_ID` | `481eb345-41e0-47cd-9624-31fd5e1647aa` | App Registration |
| `CLIENT_SECRET` | *(secret - xem Azure Entra)* | Secret name: **render-new**, hết hạn **12/6/2026** |
| `GRAPH_DRIVE_ID` | `b!EAZQ9_bOYkuEwiMNOCrvcL2qxqJ__lVBmQzE1MfcgwIFlLAQCo-gQaNDa1hd-I1J` | SharePoint Document Library PLCDVHT |
| `TARGET_FOLDER` | `PhucLoiCongDoan` | Thư mục lưu file trên SharePoint |

> ⚠️ **Lưu ý**: CLIENT_SECRET hết hạn **12/6/2026**. Cần tạo secret mới trước ngày đó tại Azure Entra.

---

## ☁️ Microsoft Azure (Graph API)

| Mục | Thông tin |
|-----|-----------|
| **Portal** | https://entra.microsoft.com |
| **App name** | congdoanvht-phucloi |
| **App (Client) ID** | `481eb345-41e0-47cd-9624-31fd5e1647aa` |
| **Tenant ID** | `f380bf27-17b2-4901-8f9f-84d2dd0410ee` |
| **Account login** | huynx268@w0tks.onmicrosoft.com |
| **Tenant domain** | w0tks.onmicrosoft.com |

### Client Secrets

| Tên | Trạng thái | Hết hạn |
|-----|-----------|---------|
| render | Cũ (không dùng) | 6/3/2028 |
| render-new | Cũ (không dùng) | 12/6/2026 |
| **render-2yr** | **Đang dùng** | **6/8/2028** |

### Quyền API (App Permissions)
- `Files.ReadWrite.All` — đọc/ghi file OneDrive/SharePoint
- `Sites.ReadWrite.All` — truy cập SharePoint Sites

---

## 📁 SharePoint

| Mục | Thông tin |
|-----|-----------|
| **Site name** | PLCDVHT (Phúc Lợi Công Đoàn VHT) |
| **Site URL** | https://w0tks.sharepoint.com/sites/PLCDVHT |
| **Site ID** | `w0tks.sharepoint.com,f7500610-cef6-4b62-84c2-230d382aef70,a2c6aabd-fe7f-4155-990c-c4d4c7dc8302` |
| **Drive (Document Library)** | Documents |
| **Drive ID** | `b!EAZQ9_bOYkuEwiMNOCrvcL2qxqJ__lVBmQzE1MfcgwIFlLAQCo-gQaNDa1hd-I1J` |
| **Thư mục lưu file** | `Shared Documents/PhucLoiCongDoan/` |

---

## 🌍 Domain & DNS

### Domain: congdoanvht.dpdns.org

| Mục | Thông tin |
|-----|-----------|
| **Domain provider** | DigitalPlat (miễn phí) |
| **Dashboard** | https://dash.domain.digitalplat.org |
| **DNS provider** | deSEC (miễn phí) |
| **deSEC dashboard** | https://desec.io |
| **deSEC account** | huynx10.viettel@gmail.com |

### DNS Records (deSEC)

| Type | Subname | Value | TTL |
|------|---------|-------|-----|
| A | (root) | `216.24.57.7` | 3600 |
| NS | (root) | ns1.desec.io, ns2.desec.org | 3600 |

> 💡 Nếu Render thay đổi IP, cần vào deSEC cập nhật lại A record.

---

## 🏗️ Kiến Trúc Hệ Thống

```
User → congdoanvht.dpdns.org
         ↓ (DNS A record)
       216.24.57.7 (Render CDN/Cloudflare)
         ↓
       congdoanvht-swq7.onrender.com (Node.js Express)
         ↓ (Microsoft Graph API)
       SharePoint PLCDVHT/Shared Documents/PhucLoiCongDoan/
```

### Stack kỹ thuật
- **Backend**: Node.js + Express (ESM modules)
- **Frontend**: HTML/CSS/JS thuần (single file `index.html`)
- **File processing**: docxtemplater (Word), pdf-lib (PDF)
- **Cloud storage**: Microsoft Graph API → SharePoint/OneDrive
- **Auth**: MSAL app-only (client credentials flow)

---

## 🔄 Quy Trình Khi Cần Bảo Trì

### Cập nhật code
1. Sửa code → push lên `main` của `huynx10viettel/congdoanvht`
2. Render tự động deploy (autoDeploy: true)

### Gia hạn CLIENT_SECRET (trước 12/6/2026)
1. Vào https://entra.microsoft.com → App registrations → congdoanvht-phucloi
2. Certificates & secrets → New client secret
3. Copy giá trị mới → Cập nhật env var `CLIENT_SECRET` trên Render
4. Dashboard Render → Service congdoanvht → Environment → Edit

### Nếu web không truy cập được
1. Kiểm tra Render service: https://dashboard.render.com
2. Kiểm tra DNS: https://desec.io → domain congdoanvht.dpdns.org
3. Nếu IP thay đổi: cập nhật A record trong deSEC

### Xem logs
- Render Dashboard → congdoanvht → Logs

---

## 📝 Tính Năng Web

Form nhập liệu phúc lợi công đoàn với các trường:
- Người đề nghị, chức vụ
- Tiền phúc lợi, tiền công đoàn, nội dung chi
- Thông tin CBCNV (tên, mã NV, chức danh, đơn vị)
- Ngày tháng năm
- Upload ảnh chứng từ (drag-drop, tối đa 20 ảnh, 10MB/ảnh)

**Output**: File Word (.docx) + PDF ảnh chứng từ → tự động upload lên SharePoint

---

*Cập nhật lần cuối: 09/06/2026*
