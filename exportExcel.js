import ExcelJS from 'exceljs';

// ── Mức tiền theo template gốc (cột F=CÔNG ĐOÀN=cd, cột G=PHÚC LỢI=pl) ─────
// Giá trị khớp với template_export.xlsx (C6=CĐ, C7=PL)
const LOAI_AMOUNTS = {
  'tham-hoi-om-cbnv':    { cd: 1_000_000, pl:   500_000 },
  'tham-hoi-om-thannhan':{ cd: 1_000_000, pl:   500_000 },
  'ket-hon':             { cd: 1_500_000, pl: 1_000_000 },
  'sinh-con':            { cd: 1_500_000, pl: 1_000_000 },
  'tham-vieng-cbnv':     { cd: 2_500_000, pl: 1_500_000 },
  'tham-vieng-thannhan': { cd: 2_000_000, pl: 1_500_000 },
};

// ── Chuyển số thành chữ tiếng Việt ──────────────────────────────────────────
function _doc3so(n, coTram) {
  const dv  = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
  const t   = Math.floor(n / 100);
  const ch  = Math.floor((n % 100) / 10);
  const dvi = n % 10;
  let s = '';
  if (t > 0 || coTram) {
    s += dv[t] + ' trăm';
    if (ch === 0 && dvi > 0) { s += ' linh ' + dv[dvi]; }
    else if (ch === 1) { s += ' mười' + (dvi === 5 ? ' lăm' : dvi > 0 ? ' ' + dv[dvi] : ''); }
    else if (ch > 1)  { s += ' ' + dv[ch] + ' mươi' + (dvi === 1 ? ' mốt' : dvi === 5 ? ' lăm' : dvi > 0 ? ' ' + dv[dvi] : ''); }
  } else if (ch > 0) {
    if (ch === 1) s += 'mười' + (dvi === 5 ? ' lăm' : dvi > 0 ? ' ' + dv[dvi] : '');
    else          s += dv[ch] + ' mươi' + (dvi === 1 ? ' mốt' : dvi === 5 ? ' lăm' : dvi > 0 ? ' ' + dv[dvi] : '');
  } else if (dvi > 0) {
    s += dv[dvi];
  }
  return s.trim();
}

function soThanhChu(so) {
  if (!so || so === 0) return 'Không đồng./.';
  const ty    = Math.floor(so / 1_000_000_000);
  const trieu = Math.floor((so % 1_000_000_000) / 1_000_000);
  const nghin = Math.floor((so % 1_000_000) / 1_000);
  const le    = so % 1_000;
  const parts = [];
  if (ty    > 0) parts.push(_doc3so(ty,    false)          + ' tỷ');
  if (trieu > 0) parts.push(_doc3so(trieu, ty > 0)         + ' triệu');
  if (nghin > 0) parts.push(_doc3so(nghin, trieu > 0 || ty > 0) + ' nghìn');
  if (le    > 0) parts.push(_doc3so(le,    nghin > 0 || trieu > 0 || ty > 0));
  if (!parts.length) return 'Không đồng./.';
  const str = parts.join(' ');
  return str.charAt(0).toUpperCase() + str.slice(1) + ' đồng./.';
}

// ── NỘI DUNG theo từng loại (khớp template) ─────────────────────────────────
function buildNoiDung(s) {
  const name = s.ten_cbcnv    || '';
  const qh   = s.qh_than_nhan || '';
  switch (s.loai_phuc_loi) {
    case 'tham-hoi-om-cbnv':
      return `Chi kinh phí CĐ, PL thăm hỏi đ/c ${name} ốm`;
    case 'tham-hoi-om-thannhan':
      return `Chi kinh phí CĐ, PL thăm hỏi ${qh} đ/c ${name} ốm`;
    case 'ket-hon':
      return `Chi kinh phí CĐ, PL chúc mừng đ/c ${name} kết hôn`;
    case 'sinh-con':
      return `Chi kinh phí CĐ, PL chúc mừng đ/c ${name} sinh con`;
    case 'tham-vieng-cbnv':
      return `Chi kinh phí CĐ, PL thăm viếng đ/c ${name} từ trần`;
    case 'tham-vieng-thannhan':
      return `Chi kinh phí CĐ, PL thăm viếng ${qh} đ/c ${name} từ trần`;
    default:
      return '';
  }
}

// ── HỒ SƠ KÈM THEO ──────────────────────────────────────────────────────────
// Label mapping cho loại giấy tờ
const GIAY_TO_LABEL = {
  'giay-ra-vien':           'Giấy ra viện',
  'giay-xac-nhan-nam-vien': 'Giấy xác nhận nằm viện',
  'giay-khai-sinh':         'Giấy khai sinh',
  'giay-chung-sinh':        'Giấy chứng sinh',
  'thiep-cuoi':             'Thiệp cưới',
  'giay-dang-ky-ket-hon':   'Giấy đăng ký kết hôn',
  'cao-pho':                'Cáo phó',
};

function buildHoSo(s) {
  // Nếu đã có giá trị rõ ràng từ meta cũ
  if (s.ho_so_kem_theo) return s.ho_so_kem_theo;

  const loai = s.loai_giay_to || '';
  const so   = s.so_giay_to   || '';
  const ngay = s.ngay_giay_to || '';

  switch (s.loai_phuc_loi) {
    case 'tham-hoi-om-cbnv':
    case 'tham-hoi-om-thannhan': {
      const label = GIAY_TO_LABEL[loai] || 'Giấy ra viện';
      return ngay ? `${label} ngày ${ngay}` : label;
    }
    case 'sinh-con': {
      const label  = GIAY_TO_LABEL[loai] || 'Giấy khai sinh';
      const soStr  = so   ? ` số ${so}`    : '';
      const ngayStr = ngay ? ` ngày ${ngay}` : '';
      return `${label}${soStr}${ngayStr}`;
    }
    case 'ket-hon': {
      if (loai === 'giay-dang-ky-ket-hon') {
        return so ? `Giấy đăng ký kết hôn số ${so}` : 'Giấy đăng ký kết hôn';
      }
      return ngay ? `Thiệp cưới ngày ${ngay}` : 'Thiệp cưới';
    }
    case 'tham-vieng-cbnv':
    case 'tham-vieng-thannhan':
      return ngay ? `Cáo phó ngày ${ngay}` : 'Cáo phó';
    default:
      return '';
  }
}

// ── HỌ TÊN người nhận (cột 10) ──────────────────────────────────────────────
function buildHoTen(s) {
  switch (s.loai_phuc_loi) {
    case 'tham-hoi-om-thannhan':
    case 'tham-vieng-thannhan':
      return s.ten_nguoi_than || '';
    default:
      return s.ten_cbcnv || '';
  }
}

// ── QUAN HỆ GIA ĐÌNH (cột 11) ───────────────────────────────────────────────
function buildQuanHe(s) {
  switch (s.loai_phuc_loi) {
    case 'tham-hoi-om-thannhan':
    case 'tham-vieng-thannhan':
      return s.qh_than_nhan || '';
    default:
      return 'Bản thân';
  }
}

// ── Shared border style ──────────────────────────────────────────────────────
const BORDER_THIN = {
  top:    { style: 'thin' },
  left:   { style: 'thin' },
  bottom: { style: 'thin' },
  right:  { style: 'thin' },
};

// ── generateExcelFromTemplate ────────────────────────────────────────────────
// Clone template_export.xlsx, điền data submissions vào sheet
// "CD - PL in cho ký nhận" bằng giá trị tĩnh (không dùng formula).
// Template có 4 data rows (8–11) → splice ra, splice n rows mới vào.
// ─────────────────────────────────────────────────────────────────────────────
export async function generateExcelFromTemplate(submissions, thang, nam, templatePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);

  const ws = wb.getWorksheet('CD - PL in cho ký nhận');
  if (!ws) throw new Error('Không tìm thấy sheet "CD - PL in cho ký nhận" trong template');

  const yyyy = String(nam);

  // 1. Ghi đè tiêu đề (A4 và A5 là merged cell có formula → ghi value tĩnh)
  ws.getCell('A4').value = `DANH SÁCH CHI TIỀN CÔNG ĐOÀN PHÚC LỢI THÁNG ${parseInt(thang)} ${yyyy}`;
  ws.getCell('A5').value = `(Kèm theo phiếu chi số PC_CĐCS số              ngày    /    /${yyyy})`;

  // 2. Xoá 7 data rows của template (rows 8–14)
  // Template có 7 mẫu: R8 (placeholder), R9-R14 (sample data)
  const DATA_START           = 8;
  const TEMPLATE_DATA_COUNT  = 7;
  ws.spliceRows(DATA_START, TEMPLATE_DATA_COUNT);
  // Sau bước này: TỔNG CỘNG ở row 8, BẰNG CHỮ ở row 9

  // 3. Chèn n data rows tại row 8
  const n = submissions.length;
  if (n > 0) {
    const dataArrays = submissions.map((s, i) => {
      const amounts = LOAI_AMOUNTS[s.loai_phuc_loi] || { cd: 0, pl: 0 };
      return [
        i + 1,
        s.ma_nv      || '',
        s.ten_cbcnv  || '',
        buildNoiDung(s),
        s.don_vi     || '',
        amounts.cd,
        amounts.pl,
        '',                         // THUẾ TNCN
        amounts.cd + amounts.pl,    // THỰC NHẬN
        buildHoTen(s),
        buildQuanHe(s),
        buildHoSo(s),
        '',                         // KÝ NHẬN
      ];
    });
    ws.spliceRows(DATA_START, 0, ...dataArrays);
  }

  // 4. Áp dụng format cho các data row mới chèn vào
  let tongCD = 0, tongPL = 0;
  for (let i = 0; i < n; i++) {
    const s       = submissions[i];
    const amounts = LOAI_AMOUNTS[s.loai_phuc_loi] || { cd: 0, pl: 0 };
    tongCD += amounts.cd;
    tongPL += amounts.pl;

    const row = ws.getRow(DATA_START + i);
    row.height = 22;

    for (let c = 1; c <= 13; c++) {
      const cell = row.getCell(c);
      cell.border    = BORDER_THIN;
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
    // Số tiền — căn phải + format
    [6, 7, 9].forEach(col => {
      const cell = row.getCell(col);
      cell.numFmt    = '#,##0';
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
    });
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(4).alignment = { wrapText: true, vertical: 'middle' };
  }

  // 5. Cập nhật TỔNG CỘNG (bây giờ ở row DATA_START + n)
  const totalRowIdx = DATA_START + n;
  const totalRow    = ws.getRow(totalRowIdx);
  totalRow.getCell(6).value = tongCD;
  totalRow.getCell(7).value = tongPL;
  totalRow.getCell(8).value = '';
  totalRow.getCell(9).value = tongCD + tongPL;
  [6, 7, 9].forEach(col => {
    const c = totalRow.getCell(col);
    c.numFmt    = '#,##0';
    c.alignment = { horizontal: 'right', vertical: 'middle' };
  });

  // 6. Cập nhật BẰNG CHỮ (row DATA_START + n + 1)
  // Template: C1='BẰNG CHỮ', C6=giá trị chữ (cột F, không phải B)
  const bangChuRow = ws.getRow(totalRowIdx + 1);
  bangChuRow.getCell(6).value = soThanhChu(tongCD + tongPL);

  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ── generateExcel ────────────────────────────────────────────────────────────
export async function generateExcel(submissions, thang, nam) {
  const wb = new ExcelJS.Workbook();

  // ══════════════════════════════════════════════════════════════════════════
  // Sheet 1: Danh sách yêu cầu (raw data)
  // ══════════════════════════════════════════════════════════════════════════
  const s1 = wb.addWorksheet('Danh sách yêu cầu');

  const s1Headers = [
    'STT', 'Ngày nộp', 'Họ và tên CBNV', 'Mã NV', 'Chức danh', 'Đơn vị',
    'Loại phúc lợi',
    'Quan hệ thân nhân', 'Họ tên người thân',
    'Loại giấy tờ', 'Số giấy tờ', 'Ngày giấy tờ',
    'Hồ sơ kèm theo (tổng hợp)',
  ];

  const hRow1 = s1.addRow(s1Headers);
  hRow1.height = 30;
  hRow1.eachCell(c => {
    c.font      = { bold: true };
    c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    c.border    = BORDER_THIN;
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  submissions.forEach((s, i) => {
    const row = s1.addRow([
      i + 1,
      s.ngay_nop       || '',
      s.ten_cbcnv      || '',
      s.ma_nv          || '',
      s.chuc_danh      || '',
      s.don_vi         || '',
      s.loai_phuc_loi  || '',
      s.qh_than_nhan   || '',
      s.ten_nguoi_than || '',
      GIAY_TO_LABEL[s.loai_giay_to] || s.loai_giay_to || '',
      s.so_giay_to     || '',
      s.ngay_giay_to   || '',
      buildHoSo(s),
    ]);
    row.eachCell(c => { c.border = BORDER_THIN; c.alignment = { vertical: 'middle' }; });
  });

  const s1Widths = [5, 12, 24, 10, 18, 16, 26, 18, 24, 26, 14, 14, 34];
  s1.columns.forEach((col, i) => { col.width = s1Widths[i] ?? 14; });

  // ══════════════════════════════════════════════════════════════════════════
  // Sheet 2: CD - PL in cho ký nhận (bảng in ký)
  // ══════════════════════════════════════════════════════════════════════════
  const s2 = wb.addWorksheet('CD - PL in cho ký nhận');
  const mm  = String(thang).padStart(2, '0');
  const TOTAL_COLS = 13; // A→M

  // ── Hàng 1: Tên tổ chức ─────────────────────────────────────────────────
  s2.mergeCells(`A1:M1`);
  const r1 = s2.getRow(1);
  r1.getCell(1).value     = 'TỔNG CÔNG TY CÔNG NGHIỆP CÔNG NGHỆ CAO VIETTEL';
  r1.getCell(1).font      = { bold: true };
  r1.getCell(1).alignment = { horizontal: 'center' };

  // ── Hàng 2: Tên công đoàn ───────────────────────────────────────────────
  s2.mergeCells('A2:M2');
  const r2 = s2.getRow(2);
  r2.getCell(1).value     = 'CÔNG ĐOÀN CƠ SỞ';
  r2.getCell(1).font      = { bold: true };
  r2.getCell(1).alignment = { horizontal: 'center' };

  // ── Hàng 3: Tiêu đề ─────────────────────────────────────────────────────
  s2.mergeCells('A3:M3');
  const r3      = s2.getRow(3);
  r3.height     = 26;
  r3.getCell(1).value     = `DANH SÁCH CHI TIỀN CÔNG ĐOÀN PHÚC LỢI THÁNG ${parseInt(thang)} ${nam}`;
  r3.getCell(1).font      = { bold: true, size: 13 };
  r3.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

  // ── Hàng 4: Ghi chú phiếu chi ──────────────────────────────────────────
  s2.mergeCells('A4:M4');
  const r4 = s2.getRow(4);
  r4.getCell(1).value     = `(Kèm theo phiếu chi số PC_CĐCS số              ngày    /    /${nam})`;
  r4.getCell(1).alignment = { horizontal: 'center' };

  // ── Hàng 5: Header cột ──────────────────────────────────────────────────
  const colHeaders = [
    'TT', 'MÃ NV', 'HỌ VÀ TÊN', 'NỘI DUNG', 'ĐƠN VỊ',
    'CÔNG ĐOÀN', 'PHÚC LỢI', 'THUẾ TNCN', 'THỰC NHẬN\n(VNĐ)',
    'HỌ TÊN', 'QUAN HỆ GIA ĐÌNH', 'HỒ SƠ KÈM THEO', 'KÝ NHẬN\n(Ký và ghi rõ họ tên)',
  ];
  const hRow2   = s2.addRow(colHeaders);  // row 5
  hRow2.height  = 36;
  hRow2.eachCell(c => {
    c.font      = { bold: true };
    c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
    c.border    = BORDER_THIN;
    c.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
  });

  // ── Data rows (bắt đầu từ row 6) ────────────────────────────────────────
  const DATA_START = 6;
  let tongCD = 0, tongPL = 0;

  submissions.forEach((s, i) => {
    const amounts  = LOAI_AMOUNTS[s.loai_phuc_loi] || { cd: 0, pl: 0 };
    tongCD += amounts.cd;
    tongPL += amounts.pl;

    const row = s2.addRow([
      i + 1,
      s.ma_nv     || '',
      s.ten_cbcnv || '',
      buildNoiDung(s),
      s.don_vi    || '',
      amounts.cd,
      amounts.pl,
      '',                  // THUẾ TNCN
      amounts.cd + amounts.pl,
      buildHoTen(s),
      buildQuanHe(s),
      buildHoSo(s),
      '',                  // KÝ NHẬN
    ]);

    row.height = 22;
    row.eachCell(c => { c.border = BORDER_THIN; c.alignment = { vertical: 'middle', wrapText: true }; });

    // Format số tiền
    [6, 7, 9].forEach(col => {
      const c = row.getCell(col);
      c.numFmt    = '#,##0';
      c.alignment = { horizontal: 'right', vertical: 'middle' };
    });
    row.getCell(4).alignment = { wrapText: true, vertical: 'middle' };
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // ── Hàng TỔNG CỘNG ──────────────────────────────────────────────────────
  const DATA_END = DATA_START + submissions.length - 1;

  const totalRow = s2.addRow([
    '', '', 'TỔNG CỘNG', '', '',
    submissions.length ? `=SUM(F${DATA_START}:F${DATA_END})` : 0,
    submissions.length ? `=SUM(G${DATA_START}:G${DATA_END})` : 0,
    '',
    submissions.length ? `=SUM(I${DATA_START}:I${DATA_END})` : 0,
    '', '', '', '',
  ]);
  totalRow.height = 22;
  totalRow.font   = { bold: true };
  totalRow.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
  totalRow.eachCell(c => { c.border = BORDER_THIN; });
  [6, 7, 9].forEach(col => {
    const c = totalRow.getCell(col);
    c.numFmt    = '#,##0';
    c.alignment = { horizontal: 'right', vertical: 'middle' };
  });
  totalRow.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };

  // ── Hàng BẰNG CHỮ ───────────────────────────────────────────────────────
  const tongTatCa = tongCD + tongPL;
  const bangChuRow = s2.addRow(['BẰNG CHỮ', soThanhChu(tongTatCa), '', '', '', '', '', '', '', '', '', '', '']);
  bangChuRow.height = 20;
  bangChuRow.getCell(1).font = { bold: true, italic: true };
  bangChuRow.getCell(2).font = { italic: true };
  bangChuRow.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };

  // ── Hàng chữ ký ─────────────────────────────────────────────────────────
  s2.addRow([]); // blank spacer
  const sigRow = s2.addRow(['', 'NGƯỜI NHẬN TIỀN', '', '', 'NGƯỜI LẬP PHIẾU', '', '', '', '', 'CHỦ TỊCH CĐCS', '', '', '']);
  sigRow.height = 20;
  sigRow.getCell(2).font  = { bold: true };
  sigRow.getCell(5).font  = { bold: true };
  sigRow.getCell(10).font = { bold: true };
  [2, 5, 10].forEach(col => {
    sigRow.getCell(col).alignment = { horizontal: 'center' };
  });

  // ── Column widths ────────────────────────────────────────────────────────
  const colWidths = [5, 10, 22, 46, 16, 14, 14, 10, 14, 22, 18, 26, 20];
  s2.columns.forEach((col, i) => { col.width = colWidths[i] ?? 12; });

  return Buffer.from(await wb.xlsx.writeBuffer());
}
