import ExcelJS from 'exceljs';

// ──────────────────────────────────────────────
// Mapping loai_phuc_loi → label (khớp với template Excel)
// ──────────────────────────────────────────────
const LOAI_LABEL = {
  'tham-hoi-om-cbnv':    'thăm hỏi CBNV',
  'tham-hoi-om-thannhan':'thăm hỏi thân nhân CBNV',
  'ket-hon':             'chúc mừng CBNV',
  'sinh-con':            'chúc mừng CBNV',
  'tham-vieng-cbcnv':    'viếng CBNV từ trần',
  'tham-vieng-thannhan': 'viếng thân nhân CBNV từ trần',
};

// Mức tiền theo loại phúc lợi (VNĐ)
const LOAI_AMOUNTS = {
  'tham-hoi-om-cbnv':    { cd: 500_000,   pl: 1_000_000 },
  'tham-hoi-om-thannhan':{ cd: 500_000,   pl: 1_000_000 },
  'ket-hon':             { cd: 1_000_000, pl: 1_500_000 },
  'sinh-con':            { cd: 1_000_000, pl: 1_500_000 },
  'tham-vieng-cbcnv':    { cd: 1_500_000, pl: 2_500_000 },
  'tham-vieng-thannhan': { cd: 1_500_000, pl: 2_000_000 },
};

function buildNoiDung(s) {
  const name = s.ten_cbcnv || '';
  const qh   = s.qh_than_nhan || '';
  const bv   = s.benh_vien || '';
  switch (s.loai_phuc_loi) {
    case 'tham-hoi-om-cbnv':
      return `Chi kinh phí CĐ, PL thăm hỏi đ/c ${name} ốm điều trị tại ${bv}`.trimEnd();
    case 'tham-hoi-om-thannhan':
      return `Chi kinh phí CĐ, PL thăm hỏi ${qh} đ/c ${name} ốm điều trị tại ${bv}`.trimEnd();
    case 'ket-hon':
      return `Chi kinh phí CĐ, PL chúc mừng đ/c ${name} kết hôn`;
    case 'sinh-con':
      return `Chi kinh phí CĐ, PL chúc mừng đ/c ${name} sinh con`;
    case 'tham-vieng-cbcnv':
      return `Chi kinh phí CĐ, PL viếng đ/c ${name} từ trần`;
    case 'tham-vieng-thannhan':
      return `Chi kinh phí CĐ, PL viếng ${qh} đ/c ${name} từ trần`.trimEnd();
    default:
      return '';
  }
}

// Format số tiền: 1500000 → "1.500.000"
function fmtVND(n) {
  return Number(n).toLocaleString('vi-VN');
}

// ──────────────────────────────────────────────
// generateExcel — tạo buffer xlsx từ mảng submissions
// submissions: [{ ten_cbcnv, ma_nv, chuc_danh, don_vi, loai_phuc_loi,
//                 qh_than_nhan, ten_nguoi_than, benh_vien, ngay_nop, ... }]
// ──────────────────────────────────────────────
export async function generateExcel(submissions, thang, nam) {
  const wb = new ExcelJS.Workbook();

  // ── Sheet 1: Danh sách yêu cầu thanh toán ──────────────────────────────────
  const s1 = wb.addWorksheet('Danh sách yêu cầu thanh toán');

  const s1Headers = [
    'ID', 'Start time', 'Completion time', 'Email', 'Name', 'Last modified time',
    'Tên CBNV', 'Mã NV', 'Chức vụ', 'Đơn vị',
    'Chọn phúc lợi yêu cầu', 'Chúc mừng CBNV', 'Thân nhân của bạn đồng chí là',
    'Nơi điều trị', 'Loại hồ sơ, giấy tờ kèm theo', 'Số giấy tờ', 'Xác nhận thông tin',
    'Họ và tên thân nhân', 'Họ và tên thân nhân:', 'Thân nhân của bạn đồng chí là:', 'Ngày mất',
  ];
  const headerRow1 = s1.addRow(s1Headers);
  headerRow1.font = { bold: true };
  headerRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };

  submissions.forEach((s, i) => {
    const chuMungLoai = s.loai_phuc_loi === 'ket-hon'
      ? 'kết hôn' : s.loai_phuc_loi === 'sinh-con' ? 'sinh con' : '';
    const soGiayTo = s.so_giay_ket_hon || s.so_giay_khai_sinh || '';
    s1.addRow([
      i + 1,
      s.ngay_nop || '',
      s.ngay_nop || '',
      '', '', '',
      s.ten_cbcnv     || '',
      s.ma_nv         || '',
      s.chuc_danh     || '',
      s.don_vi        || '',
      LOAI_LABEL[s.loai_phuc_loi] || s.loai_phuc_loi || '',
      chuMungLoai,
      s.qh_than_nhan  || '',
      s.benh_vien     || '',
      '',
      soGiayTo,
      '',
      s.ten_nguoi_than || '',
      '',
      s.qh_than_nhan  || '',
      '',
    ]);
  });

  // Auto-width for key columns
  s1.columns.forEach((col, idx) => {
    col.width = idx < 6 ? 16 : idx === 6 ? 22 : 18;
  });

  // ── Sheet 2: CD - PL in cho ký nhận ────────────────────────────────────────
  const s2 = wb.addWorksheet('CD - PL in cho ký nhận');

  // Tiêu đề bảng
  s2.mergeCells('A1:M1');
  const titleRow = s2.getRow(1);
  titleRow.getCell(1).value =
    `DANH SÁCH CHI PHÍ CÔNG ĐOÀN VÀ PHÚC LỢI – THÁNG ${parseInt(thang)}/${nam}`;
  titleRow.getCell(1).font      = { bold: true, size: 13 };
  titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
  titleRow.height = 28;

  const s2Headers = [
    'STT', 'MÃ NV', 'HỌ VÀ TÊN', 'NỘI DUNG', 'ĐƠN VỊ',
    'CÔNG ĐOÀN\n(đồng)', 'PHÚC LỢI\n(đồng)', 'THUẾ TNCN', 'THỰC NHẬN\n(đồng)',
    'HỌ TÊN', 'QUAN HỆ GIA ĐÌNH', 'HỒ SƠ KÈM THEO', 'KÝ NHẬN',
  ];
  const headerRow2 = s2.addRow(s2Headers);
  headerRow2.font      = { bold: true };
  headerRow2.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
  headerRow2.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
  headerRow2.height = 36;

  const borderThin = {
    top:    { style: 'thin' },
    left:   { style: 'thin' },
    bottom: { style: 'thin' },
    right:  { style: 'thin' },
  };

  // Apply border to header
  headerRow2.eachCell(cell => {
    cell.border    = borderThin;
    cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
  });

  let tongCD = 0, tongPL = 0, tongNhan = 0;

  submissions.forEach((s, i) => {
    const amounts = LOAI_AMOUNTS[s.loai_phuc_loi] || { cd: 0, pl: 0 };
    const thueTNCN = 0;
    const thucNhan = amounts.cd + amounts.pl - thueTNCN;
    tongCD   += amounts.cd;
    tongPL   += amounts.pl;
    tongNhan += thucNhan;

    const row = s2.addRow([
      i + 1,
      s.ma_nv          || '',
      s.ten_cbcnv      || '',
      buildNoiDung(s),
      s.don_vi         || '',
      amounts.cd,
      amounts.pl,
      thueTNCN || '',
      thucNhan,
      s.ten_nguoi_than || '',
      s.qh_than_nhan   || '',
      '',
      '',
    ]);

    // Format money cells
    [6, 7, 8, 9].forEach(colIdx => {
      const cell = row.getCell(colIdx);
      if (cell.value !== '') {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'right' };
      }
    });

    row.eachCell(cell => { cell.border = borderThin; });
  });

  // Tổng cộng
  if (submissions.length > 0) {
    const totalRow = s2.addRow([
      '', '', 'TỔNG CỘNG', '', '',
      tongCD, tongPL, '', tongNhan,
      '', '', '', '',
    ]);
    totalRow.font = { bold: true };
    totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
    [6, 7, 9].forEach(colIdx => {
      const cell = totalRow.getCell(colIdx);
      cell.numFmt = '#,##0';
      cell.alignment = { horizontal: 'right' };
    });
    totalRow.eachCell(cell => { cell.border = borderThin; });
  }

  // Column widths for sheet 2
  const s2ColWidths = [5, 10, 22, 42, 18, 14, 14, 10, 14, 20, 18, 16, 12];
  s2.columns.forEach((col, i) => { col.width = s2ColWidths[i] ?? 12; });

  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}
