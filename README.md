# Web quản lý yêu cầu phúc lợi công đoàn

Web nội bộ giúp CBNV nộp đề nghị hưởng phúc lợi công đoàn. Luồng nghiệp vụ:

1. Người dùng điền form + upload ảnh chứng từ trên web.
2. Hệ thống trích xuất thông tin, **điền vào template Word** (`.docx`) bằng `docxtemplater`.
3. **Convert Word đã điền sang PDF** bằng Microsoft Graph (`?format=pdf`).
4. **Gộp ảnh đính kèm thành 1 file PDF** bằng `pdf-lib`.
5. **Lưu tất cả vào 1 thư mục** trên OneDrive / SharePoint (mỗi đơn 1 folder riêng).

Toàn bộ dùng thư viện **mã nguồn mở** và hạ tầng **miễn phí**.

---

## Kiến trúc & vì sao chọn

| Thành phần | Công nghệ | Lý do |
|---|---|---|
| Backend | Node.js + Express (open source) | Nhẹ, miễn phí, dễ deploy |
| Điền Word | [docxtemplater](https://github.com/open-xml-templating/docxtemplater) | Mã nguồn mở, dùng cú pháp `{placeholder}`, người không lập trình cũng sửa template được |
| Ảnh → PDF | [pdf-lib](https://github.com/Hopding/pdf-lib) | Mã nguồn mở, chạy thuần JS |
| Word → PDF | [Microsoft Graph](https://learn.microsoft.com/graph) `GET /items/{id}/content?format=pdf` | **Miễn phí** với license M365 sẵn có, convert chuẩn Office, không cần cài LibreOffice |
| Lưu trữ | OneDrive / SharePoint qua Graph | Đúng nơi công ty đang dùng, phân quyền sẵn |
| Auth | MSAL app-only (client credentials) | Không cần người dùng đăng nhập, form mở cho mọi CBNV |
| Hosting | [Render](https://render.com) free tier | Cho phép dùng nội bộ doanh nghiệp, có URL công khai miễn phí |

> Lưu ý hosting: **Vercel Hobby (free) cấm dùng cho mục đích thương mại/doanh nghiệp**, nên với web nội bộ công ty hãy dùng **Render free** (hoặc Cloudflare/Railway). Render free "ngủ" sau 15 phút không dùng và mất ~30–60s để thức dậy ở request đầu — chấp nhận được với lượng đơn phúc lợi không lớn.

---

## Chạy thử nhanh trên máy (chưa cần Microsoft Graph)

```bash
cd phuc-loi-cong-doan
npm install
npm start
```

Mở http://localhost:3000 → điền form → nộp. Vì chưa cấu hình Graph, hệ thống chạy **chế độ thử cục bộ**: lưu file `.docx` + PDF ảnh vào thư mục `./output/<tên-đơn>/` (chưa convert Word→PDF vì việc đó do Graph làm).

---

## Cấu hình Microsoft Graph (để convert PDF + lưu OneDrive/SharePoint)

### 1. Đăng ký ứng dụng trong Entra ID (Azure AD)
1. Vào https://entra.microsoft.com → **App registrations** → **New registration**. Đặt tên, chọn "Single tenant".
2. Ghi lại **Application (client) ID** và **Directory (tenant) ID**.
3. **Certificates & secrets** → **New client secret** → ghi lại giá trị (chỉ hiện 1 lần).
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → thêm `Sites.ReadWrite.All` (hoặc `Files.ReadWrite.All`) → **Grant admin consent** (cần quyền admin).

### 2. Lấy `GRAPH_DRIVE_ID` (thư viện tài liệu SharePoint cần lưu)
Dùng [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) hoặc curl với token:
```
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}
  → lấy site id
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
  → lấy "id" của Document Library muốn dùng
```

### 3. Điền `.env`
```bash
cp .env.example .env
# rồi điền TENANT_ID, CLIENT_ID, CLIENT_SECRET, GRAPH_DRIVE_ID
```

Khởi động lại `npm start`. Giờ mỗi đơn sẽ được lưu vào `TARGET_FOLDER/<tên-đơn>/` gồm: đơn `.docx`, đơn `.pdf`, và `anh-dinh-kem.pdf`.

> App-only chỉ hoạt động với **SharePoint / OneDrive for Business** (không dùng được OneDrive cá nhân). Đây cũng là lý do nên trỏ vào một thư viện tài liệu SharePoint của công ty.

---

## Deploy miễn phí lên Render

1. Đẩy thư mục này lên một repo Git (GitHub/GitLab).
2. Trên Render: **New** → **Web Service** → kết nối repo.
3. Build command: `npm install` — Start command: `npm start`.
4. Mục **Environment**: thêm các biến trong `.env` (TENANT_ID, CLIENT_ID, CLIENT_SECRET, GRAPH_DRIVE_ID, TARGET_FOLDER).
5. Deploy → nhận URL dạng `https://<ten-app>.onrender.com` để phát cho CBNV.

---

## Thay template bằng mẫu đơn của công ty

File mẫu hiện tại: `templates/mau-don-phuc-loi.docx` (đã dựng từ "Giấy đề nghị chi Quỹ phúc lợi & Quỹ công đoàn" của công ty, giữ nguyên layout/bảng chữ ký). Mở bằng Word và sửa nội dung tùy ý, **giữ nguyên các placeholder dạng `{ten_truong}`**. Các trường đang dùng:

| Placeholder | Ý nghĩa | Nguồn |
|---|---|---|
| `{nguoi_de_nghi}` | Họ tên người đề nghị | Form |
| `{chuc_vu_de_nghi}` | Chức vụ người đề nghị | Form |
| `{tien_phuc_loi}` / `{tien_phuc_loi_chu}` | Tiền Quỹ phúc lợi (số / bằng chữ) | Form + tự đọc |
| `{tien_cong_doan}` / `{tien_cong_doan_chu}` | Tiền Quỹ công đoàn (số / bằng chữ) | Form + tự đọc |
| `{tong_cong}` / `{tong_cong_chu}` | Tổng cộng (số / bằng chữ) | **Server tự tính** |
| `{noi_dung_chi}` | Nội dung chi | Form |
| `{ten_cbcnv}` `{ma_nv}` `{chuc_danh}` `{don_vi}` | CBCNV được thụ hưởng | Form |
| `{ngay}` `{thang}` `{nam}` | Ngày lập đơn (mặc định hôm nay) | Form |

Số tiền: người dùng chỉ nhập **số**, hệ thống tự tính tổng và đọc thành chữ (module `src/soThanhChu.js`, đã test các mốc thường gặp).

Muốn thêm trường mới: thêm `{ten_moi}` vào Word, thêm `<input name="ten_moi">` trong `public/index.html`, và thêm `ten_moi: f.ten_moi || ""` trong `server.js`.

---

## Cấu trúc thư mục
```
phuc-loi-cong-doan/
├─ server.js              # Express server + luồng nghiệp vụ
├─ src/
│  ├─ docxFiller.js       # Điền Word bằng docxtemplater
│  ├─ imagesToPdf.js      # Gộp ảnh -> PDF bằng pdf-lib
│  └─ graph.js            # Microsoft Graph: convert PDF + lưu OneDrive/SharePoint
├─ public/index.html      # Form web
├─ templates/mau-don-phuc-loi.docx   # Template mẫu (thay bằng mẫu công ty)
├─ .env.example
└─ package.json
```
