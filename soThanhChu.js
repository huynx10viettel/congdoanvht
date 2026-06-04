// Đọc số tiền (VNĐ) thành chữ tiếng Việt. Vd: 1500000 -> "Một triệu năm trăm nghìn đồng chẵn".
const DV = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];

function docBaChuSo(so, dayDu) {
  const tram = Math.floor(so / 100);
  const chuc = Math.floor((so % 100) / 10);
  const donVi = so % 10;
  let kq = "";
  if (dayDu || tram > 0) {
    kq += DV[tram] + " trăm";
  }
  if (chuc === 0) {
    if (donVi > 0 && (dayDu || tram > 0)) kq += " lẻ";
  } else if (chuc === 1) {
    kq += " mười";
  } else {
    kq += " " + DV[chuc] + " mươi";
  }
  if (donVi > 0) {
    if (donVi === 1 && chuc >= 2) kq += " mốt";
    else if (donVi === 5 && chuc >= 1) kq += " lăm";
    else kq += " " + DV[donVi];
  }
  return kq.trim();
}

/** @param {number|string} input số tiền @returns {string} chữ */
export function soThanhChu(input) {
  let n = parseInt(String(input).replace(/[^\d]/g, ""), 10);
  if (!Number.isFinite(n)) return "";
  if (n === 0) return "Không đồng";

  const nhom = ["", " nghìn", " triệu", " tỷ"]; // mỗi nhóm 3 chữ số
  const phan = [];
  while (n > 0) {
    phan.push(n % 1000);
    n = Math.floor(n / 1000);
  }

  let kq = "";
  for (let i = phan.length - 1; i >= 0; i--) {
    if (phan[i] === 0) continue;
    const dayDu = i < phan.length - 1; // các nhóm sau nhóm cao nhất đọc đầy đủ
    kq += docBaChuSo(phan[i], dayDu) + nhom[i] + " ";
  }
  kq = kq.replace(/\s+/g, " ").trim();
  kq = kq.charAt(0).toUpperCase() + kq.slice(1);
  return kq + " đồng chẵn";
}

/** Chuẩn hóa chuỗi tiền về số nguyên (bỏ . , đ VNĐ...) */
export function chuanHoaTien(input) {
  const n = parseInt(String(input || "").replace(/[^\d]/g, ""), 10);
  return Number.isFinite(n) ? n : 0;
}

/** Định dạng số có dấu chấm phân nhóm kiểu VN: 1500000 -> "1.500.000" */
export function dinhDangTien(n) {
  return Number(n).toLocaleString("vi-VN");
}
