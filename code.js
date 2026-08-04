/**
 * ============================================
 * TOYOTAKI NVL Dashboard - Google Apps Script
 * ============================================
 * 
 * V2: Hỗ trợ sheet "Tồn Đầu Kỳ"
 * Logic: Tồn thực tế = Tồn Đầu Kỳ + Σ(Nhập - Xuất) từ ngày đầu kỳ
 * 
 * ============================================
 */

const SHEET_ID = "1sFRigbmMAKdKX2spRrq6IPjPeNA3hDaOECobqaXMa-8";

const SHEETS = {
  OPENING: 'Tồn Đầu Kỳ',
  PLAN: 'Kế Hoạch',
  TRANSACTION: 'Giao Dịch',
  TEM: 'Tem NVL',
  SETTINGS: 'Cài Đặt',
  STOCK: 'Tồn Hiện Tại'
};

// ============================================
// MATERIALS - CHỈ dùng cho các sheet có DÒNG CỐ ĐỊNH:
//   • Tồn Đầu Kỳ, Kế Hoạch, Cài Đặt, Tồn Hiện Tại
//   → 4 sheet này có dòng thứ N ứng với mã thứ N trong list này.
//   → Nếu bạn thêm mã mới, phải:
//     1. Thêm mã vào cuối list này
//     2. Thêm 1 dòng tương ứng ở cuối mỗi sheet trên
//
// Riêng tem NVL (in tem, quét QR xuất/nhập): KHÔNG cần list này
//   → Luôn lấy mã và tên từ Master (file KHSX cột V:W).
// ============================================
const MATERIALS = [
  'ADC12-DAK', 'ADC12-CHT', 'ADC12-TAQ', 'ADC6-CHT',
  'HD2-HDT', 'EZDA3-HDT', 'EZDA3-MET'
];

const PLAN_START_ROW = 5;
const PLAN_START_COL = 4;
const PLAN_NUM_DAYS = 30;

/**
 * GET HANDLER
 */
function doGet(e) {
  const action = (e && e.parameter) ? (e.parameter.action || 'getAllData') : 'getAllData';
  
  try {
    let result;
    switch(action) {
      case 'getAllData':      result = getAllData(); break;
      case 'getOpening':      result = getOpeningStockData(); break;
      case 'getPlanDaily':    result = getPlanDailyData(); break;
      case 'getSettings':     result = getSettingsData(); break;
      case 'getCurrentStock': result = getCurrentStockData(); break;
      case 'getTransactions': result = getTransactionsData(); break;
      case 'addTransaction':  result = addTransaction(e.parameter); break;
      case 'updateOpening':   result = updateOpening(e.parameter); break;
      case 'updateStock':     result = updateStock(e.parameter); break;
      case 'updateSettings':  result = updateSettings(e.parameter); break;
      case 'updatePlanNhap':  result = updatePlanNhap(e.parameter); break;
      case 'recalcAllPlanStock': result = recalcAllPlanStock(); break;
      case 'createTem':       result = createTem(e.parameter); break;
      case 'getTemList':      result = getTemList(); break;
      case 'getTemByTagNo':   result = getTemByTagNo(e.parameter); break;
      case 'checkFIFO':       result = checkFIFO(e.parameter); break;
      case 'processMultiTransaction': result = processMultiTransaction(e.parameter); break;
      case 'getMasterMaterials': result = getMasterMaterials(); break;
      case 'test':            result = { status: 'ok', message: 'API v2 hoạt động!', timestamp: new Date().toISOString() }; break;
      default: return returnError('Invalid action: ' + action);
    }
    return returnJSON(result);
  } catch (error) {
    return returnError(error.toString());
  }
}

/**
 * POST HANDLER
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    let result;
    switch(action) {
      case 'addTransaction': result = addTransaction(params); break;
      case 'updateOpening':  result = updateOpening(params); break;
      case 'updateStock':    result = updateStock(params); break;
      case 'updateSettings': result = updateSettings(params); break;
      case 'updatePlanNhap': result = updatePlanNhap(params); break;
      case 'createTem':      result = createTem(params); break;
      case 'processMultiTransaction': result = processMultiTransaction(params); break;
      default: return returnError('Invalid action: ' + action);
    }
    return returnJSON(result);
  } catch (error) {
    return returnError(error.toString());
  }
}

/**
 * READ FUNCTIONS
 */

// Lấy TẤT CẢ dữ liệu
function getAllData() {
  return {
    opening: getOpeningStockData(),      // Tồn đầu kỳ
    planDaily: getPlanDailyData(),       // Kế hoạch daily 30 ngày
    transactions: getTransactionsData(), // Toàn bộ giao dịch
    settings: getSettingsData(),         // Bề rộng quản lý
    currentStock: getCurrentStockData(), // Tồn hiện tại
    timestamp: new Date().toISOString()
  };
}

// Lấy tồn đầu kỳ (NEW)
// Trả về: { "ADC12-DAK": { openingDate: "2026-07-01", openingStock: 21.113 }, ... }
function getOpeningStockData() {
  const sheet = getSheet(SHEETS.OPENING);
  const data = sheet.getRange('A4:C100').getValues();
  const result = {};
  
  data.forEach(row => {
    const code = row[0];
    if (code && code.toString().trim()) {
      result[code.toString().trim()] = {
        openingDate: formatDate(row[1]),
        openingStock: parseFloat(row[2]) || 0
      };
    }
  });
  
  return result;
}

// Lấy kế hoạch daily 30 ngày
function getPlanDailyData() {
  const sheet = getSheet(SHEETS.PLAN);
  
  const dateHeader = sheet.getRange(3, PLAN_START_COL, 1, PLAN_NUM_DAYS).getValues()[0];
  const dates = dateHeader.map(d => formatDate(d));
  
  const numRows = MATERIALS.length * 3;
  const dataRange = sheet.getRange(PLAN_START_ROW, 1, numRows, 3 + PLAN_NUM_DAYS).getValues();
  
  const result = {};
  for (let i = 0; i < MATERIALS.length; i++) {
    const code = MATERIALS[i];
    const nhapRow = dataRange[i * 3];
    const tieuHaoRow = dataRange[i * 3 + 1];
    const tonRow = dataRange[i * 3 + 2];
    
    const nhapDaily = [];
    const tieuHaoDaily = [];
    const tonDuKienDaily = [];
    
    for (let d = 0; d < PLAN_NUM_DAYS; d++) {
      const colIdx = 3 + d;
      nhapDaily.push(parseFloat(nhapRow[colIdx]) || 0);
      tieuHaoDaily.push(parseFloat(tieuHaoRow[colIdx]) || 0);
      tonDuKienDaily.push(parseFloat(tonRow[colIdx]) || 0);
    }
    
    result[code] = {
      dates: dates,
      nhap: nhapDaily,          // kg
      tieuHao: tieuHaoDaily,    // kg
      tonDuKien: tonDuKienDaily // kg
    };
  }
  return result;
}

// Lấy cài đặt bề rộng
function getSettingsData() {
  const sheet = getSheet(SHEETS.SETTINGS);
  const data = sheet.getRange('A4:B100').getValues();
  const result = {};
  data.forEach(row => {
    const code = row[0];
    if (code && code.toString().trim()) {
      result[code.toString().trim()] = parseFloat(row[1]) || 5;
    }
  });
  return result;
}

// Lấy tồn hiện tại
function getCurrentStockData() {
  const sheet = getSheet(SHEETS.STOCK);
  const data = sheet.getRange('A4:C100').getValues();
  const result = {};
  data.forEach(row => {
    const code = row[0];
    if (code && code.toString().trim()) {
      result[code.toString().trim()] = parseFloat(row[1]) || 0;
    }
  });
  return result;
}

// Lấy tất cả giao dịch
function getTransactionsData() {
  const sheet = getSheet(SHEETS.TRANSACTION);
  const data = sheet.getRange('A4:G10000').getValues();
  const result = [];
  data.forEach(row => {
    if (row[0]) {
      result.push({
        date: formatDate(row[0]),
        type: row[1] ? row[1].toString() : '',
        material: row[2] ? row[2].toString() : '',
        quantity: parseFloat(row[3]) || 0,
        comment: row[4] ? row[4].toString() : '',
        afterStock: parseFloat(row[5]) || 0,
        timestamp: row[6] ? row[6].toString() : ''
      });
    }
  });
  return result;
}

/**
 * WRITE FUNCTIONS
 */

// Thêm giao dịch mới
function addTransaction(params) {
  const sheet = getSheet(SHEETS.TRANSACTION);
  const stockSheet = getSheet(SHEETS.STOCK);
  
  const date = params.date || formatDate(new Date());
  const type = params.type || 'IN';
  const material = params.material || '';
  const quantity = parseFloat(params.quantity) || 0;
  const comment = params.comment || '';
  const afterStock = parseFloat(params.afterStock) || 0;
  const timestamp = new Date().toISOString();
  
  // Check mã theo Master (không dùng list nội bộ)
  const masterCheck = checkMasterMaterial(material);
  if (!masterCheck.valid) {
    return { status: 'error', message: 'Mã NVL không tồn tại trong Master: ' + material };
  }
  if (quantity <= 0) {
    return { status: 'error', message: 'Khối lượng phải > 0' };
  }
  
  const lastRow = findLastDataRow(sheet, 4);
  const newRow = lastRow + 1;
  
  sheet.getRange(newRow, 1, 1, 7).setValues([[
    date, type, material, quantity, comment, afterStock, timestamp
  ]]);
  
  updateStockForMaterial(stockSheet, material, afterStock);
  
  return { status: 'ok', message: 'OK', row: newRow };
}

// Cập nhật tồn đầu kỳ (NEW)
function updateOpening(params) {
  const sheet = getSheet(SHEETS.OPENING);
  const material = params.material;
  const openingDate = params.openingDate;
  const openingStock = parseFloat(params.openingStock);
  
  if (!material || !MATERIALS.includes(material)) {
    return { status: 'error', message: 'Mã NVL không hợp lệ' };
  }
  
  const rowIndex = findRowByMaterial(sheet, material, 4);
  if (rowIndex === -1) {
    return { status: 'error', message: 'Không tìm thấy mã NVL trong sheet Tồn Đầu Kỳ' };
  }
  
  if (openingDate) sheet.getRange(rowIndex, 2).setValue(openingDate);
  if (!isNaN(openingStock)) sheet.getRange(rowIndex, 3).setValue(openingStock);
  
  // Đồng bộ tồn hiện tại về = tồn đầu kỳ (nếu chưa có giao dịch nào)
  const stockSheet = getSheet(SHEETS.STOCK);
  if (!isNaN(openingStock)) {
    updateStockForMaterial(stockSheet, material, openingStock);
  }
  
  return { status: 'ok', message: 'Cập nhật tồn đầu kỳ thành công' };
}

// Cập nhật tồn hiện tại
function updateStock(params) {
  const stockSheet = getSheet(SHEETS.STOCK);
  const material = params.material;
  const stock = parseFloat(params.stock);
  
  if (!material || !MATERIALS.includes(material)) {
    return { status: 'error', message: 'Mã NVL không hợp lệ' };
  }
  
  updateStockForMaterial(stockSheet, material, stock);
  return { status: 'ok', message: 'OK' };
}

// Cập nhật bề rộng
function updateSettings(params) {
  const sheet = getSheet(SHEETS.SETTINGS);
  const material = params.material;
  const width = parseFloat(params.width);
  
  if (!material || !MATERIALS.includes(material)) {
    return { status: 'error', message: 'Mã NVL không hợp lệ' };
  }
  
  const rowIndex = findRowByMaterial(sheet, material, 4);
  if (rowIndex === -1) {
    return { status: 'error', message: 'Không tìm thấy mã NVL' };
  }
  
  if (!isNaN(width)) sheet.getRange(rowIndex, 2).setValue(width);
  return { status: 'ok', message: 'OK' };
}

// Cập nhật kế hoạch nhập nhôm cho một mã NVL (NEW)
// Params: material, planNhap (JSON string: mảng 30 số kg cho 30 ngày)
function updatePlanNhap(params) {
  const sheet = getSheet(SHEETS.PLAN);
  const material = params.material;
  
  if (!material || !MATERIALS.includes(material)) {
    return { status: 'error', message: 'Mã NVL không hợp lệ: ' + material };
  }
  
  // Parse mảng số từ JSON
  let planNhapArray;
  try {
    planNhapArray = JSON.parse(params.planNhap);
  } catch (e) {
    return { status: 'error', message: 'Dữ liệu kế hoạch nhập không hợp lệ' };
  }
  
  if (!Array.isArray(planNhapArray) || planNhapArray.length !== PLAN_NUM_DAYS) {
    return { status: 'error', message: 'Kế hoạch nhập phải có ' + PLAN_NUM_DAYS + ' ngày' };
  }
  
  // Mỗi mã có 3 dòng liên tiếp trong sheet Kế Hoạch:
  //   dòng 0: Kế hoạch nhập (kg)
  //   dòng 1: Kế hoạch tiêu hao (kg)
  //   dòng 2: Tồn kho dự kiến (kg) — được TÍNH bằng: tồn cũ + nhập - tiêu hao
  const matIndex = MATERIALS.indexOf(material);
  const nhapRow = PLAN_START_ROW + matIndex * 3;
  const tieuHaoRow = nhapRow + 1;
  const tonRow = nhapRow + 2;
  
  // Ghi mảng nhập nhôm vào dòng "Kế hoạch nhập"
  const nhapValues = planNhapArray.map(v => parseFloat(v) || 0);
  sheet.getRange(nhapRow, PLAN_START_COL, 1, PLAN_NUM_DAYS).setValues([nhapValues]);
  
  // Đọc dòng tiêu hao hiện tại (để tính tồn)
  const tieuHaoValues = sheet.getRange(tieuHaoRow, PLAN_START_COL, 1, PLAN_NUM_DAYS).getValues()[0]
    .map(v => parseFloat(v) || 0);
  
  // Đọc tồn đầu kỳ (ngày đầu tiên đã có sẵn trong ô đầu của dòng tồn)
  // Convention: ô đầu tiên (cột D) chứa tồn đầu kỳ nhập tay
  const tonDauKy = parseFloat(sheet.getRange(tonRow, PLAN_START_COL).getValue()) || 0;
  
  // Tính lại tồn dự kiến cho cả 30 ngày:
  //   Ngày 0: giữ nguyên tồn đầu kỳ (đã có)
  //   Ngày N (N >= 1): tồn(N) = tồn(N-1) + nhập(N) - tiêu hao(N)
  const tonValues = new Array(PLAN_NUM_DAYS);
  tonValues[0] = tonDauKy;
  for (let i = 1; i < PLAN_NUM_DAYS; i++) {
    tonValues[i] = tonValues[i-1] + nhapValues[i] - tieuHaoValues[i];
  }
  
  // Ghi mảng tồn dự kiến (giá trị số, KHÔNG dùng formula để tránh vấn đề locale VN)
  sheet.getRange(tonRow, PLAN_START_COL, 1, PLAN_NUM_DAYS).setValues([tonValues]);
  
  return { 
    status: 'ok', 
    message: 'Đã cập nhật kế hoạch nhập ' + material + ' và tính lại tồn dự kiến',
    row: nhapRow,
    daysUpdated: PLAN_NUM_DAYS,
    tonDauKy: tonDauKy,
    tonCuoiKy: tonValues[PLAN_NUM_DAYS - 1]
  };
}

// Tính lại tồn kế hoạch cho TẤT CẢ 7 mã NVL
// Hữu ích khi user sửa tiêu hao hoặc tồn đầu kỳ trực tiếp trong Sheet
function recalcAllPlanStock() {
  const sheet = getSheet(SHEETS.PLAN);
  const results = [];
  
  MATERIALS.forEach((material, matIndex) => {
    const nhapRow = PLAN_START_ROW + matIndex * 3;
    const tieuHaoRow = nhapRow + 1;
    const tonRow = nhapRow + 2;
    
    const nhapValues = sheet.getRange(nhapRow, PLAN_START_COL, 1, PLAN_NUM_DAYS).getValues()[0]
      .map(v => parseFloat(v) || 0);
    const tieuHaoValues = sheet.getRange(tieuHaoRow, PLAN_START_COL, 1, PLAN_NUM_DAYS).getValues()[0]
      .map(v => parseFloat(v) || 0);
    const tonDauKy = parseFloat(sheet.getRange(tonRow, PLAN_START_COL).getValue()) || 0;
    
    const tonValues = new Array(PLAN_NUM_DAYS);
    tonValues[0] = tonDauKy;
    for (let i = 1; i < PLAN_NUM_DAYS; i++) {
      tonValues[i] = tonValues[i-1] + nhapValues[i] - tieuHaoValues[i];
    }
    
    sheet.getRange(tonRow, PLAN_START_COL, 1, PLAN_NUM_DAYS).setValues([tonValues]);
    
    results.push({
      material: material,
      tonDauKy: tonDauKy,
      tonCuoiKy: tonValues[PLAN_NUM_DAYS - 1]
    });
  });
  
  return { status: 'ok', message: 'Đã tính lại tồn kế hoạch cho ' + MATERIALS.length + ' mã', results: results };
}

// ==========================================
// TEM NVL - QR CODE LABELS
// ==========================================

// Tạo tem mới với Tag No tự động
// Params: material, tenNVL, ngayNhap, donVi, ghiChu
//   + soLuongList (JSON: mảng số lượng cho từng tem)  → ưu tiên nếu có
//     HOẶC soLuong (1 số) + quantity (số tem)         → mọi tem cùng số lượng
// Trả về: mảng các tem đã tạo (mỗi tem có tagNo riêng, có thể có số lượng khác nhau)
function createTem(params) {
  const sheet = getSheet(SHEETS.TEM);
  
  const material = params.material || '';
  const ngayNhap = params.ngayNhap || formatDate(new Date());
  const donVi = params.donVi || 'kg';
  const ghiChu = params.ghiChu || '';
  
  // Xác định soLuongList
  let soLuongList = [];
  if (params.soLuongList) {
    try {
      const parsed = JSON.parse(params.soLuongList);
      if (Array.isArray(parsed)) {
        soLuongList = parsed.map(v => parseFloat(v) || 0);
      }
    } catch (e) {
      return { status: 'error', message: 'soLuongList không phải JSON hợp lệ' };
    }
  } else {
    // Fallback: dùng soLuong + quantity → mảng đồng đều
    const soLuong = parseFloat(params.soLuong) || 0;
    const quantity = parseInt(params.quantity) || 1;
    soLuongList = new Array(quantity).fill(soLuong);
  }
  
  // Validate mã theo Master (không dùng list nội bộ)
  const masterCheck = checkMasterMaterial(material);
  if (!masterCheck.valid) {
    return { status: 'error', message: 'Mã NVL không tồn tại trong Master: ' + material };
  }
  if (soLuongList.length < 1 || soLuongList.length > 100) {
    return { status: 'error', message: 'Số tem phải từ 1-100' };
  }
  if (soLuongList.some(v => v <= 0)) {
    return { status: 'error', message: 'Tất cả số lượng phải > 0' };
  }
  if (!['kg', 'pcs'].includes(donVi)) {
    return { status: 'error', message: 'Đơn vị phải là kg hoặc pcs' };
  }
  
  // Ưu tiên tên từ Master, fallback về mã
  const tenNVL = params.tenNVL || masterCheck.ten || material;
  const timestamp = new Date().toISOString();
  
  // Sinh Tag No: TYK-YYYYMMDD-XXXX (XXXX là số thứ tự trong ngày, 4 chữ số)
  const datePart = ngayNhap.replace(/-/g, ''); // 20260710
  
  // Tìm số cuối cùng đã dùng trong ngày ngayNhap
  const allTems = sheet.getRange(4, 1, sheet.getLastRow() - 3 || 1, 1).getValues();
  let maxSeq = 0;
  const prefix = 'TYK-' + datePart + '-';
  allTems.forEach(row => {
    if (row[0] && row[0].toString().startsWith(prefix)) {
      const seq = parseInt(row[0].toString().substring(prefix.length));
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  
  // Tạo N tem, mỗi tem 1 Tag No liên tiếp
  const createdTems = [];
  const rowsToAppend = [];
  
  soLuongList.forEach(soLuong => {
    maxSeq++;
    const seqStr = ('0000' + maxSeq).slice(-4);  // 0001, 0002, ...
    const tagNo = prefix + seqStr;
    
    rowsToAppend.push([
      tagNo,
      material,
      tenNVL,
      ngayNhap,
      soLuong,
      donVi,
      ghiChu,
      timestamp,
      'Đã nhập kho'  // Status ngay khi tạo tem = đã nhập
    ]);
    
    createdTems.push({
      tagNo: tagNo,
      material: material,
      tenNVL: tenNVL,
      ngayNhap: ngayNhap,
      soLuong: soLuong,
      donVi: donVi,
      ghiChu: ghiChu,
      timestamp: timestamp,
      qrContent: [tagNo, material, tenNVL, ngayNhap, soLuong, donVi].join('>')
    });
  });
  
  // Ghi tất cả vào sheet Tem NVL 1 lần
  const startRow = findLastDataRow(sheet, 4) + 1;
  sheet.getRange(startRow, 1, rowsToAppend.length, 9).setValues(rowsToAppend);
  
  // TỰ ĐỘNG GHI GIAO DỊCH NHẬP KHO cho từng pallet
  const txSheet = getSheet(SHEETS.TRANSACTION);
  const stockSheet = getSheet(SHEETS.STOCK);
  const txStartRow = findLastDataRow(txSheet, 4) + 1;
  const txRows = [];
  
  createdTems.forEach(tem => {
    // Chuyển kg → tấn cho tồn kho (giao dịch dùng tấn theo convention hiện tại)
    const qtyInTons = tem.donVi === 'kg' ? tem.soLuong / 1000 : tem.soLuong;
    
    txRows.push([
      tem.ngayNhap,
      'IN',
      tem.material,
      qtyInTons,
      '[Tem ' + tem.tagNo + '] ' + (tem.ghiChu ? tem.ghiChu : 'Nhập kho từ in tem'),
      0, // afterStock sẽ được tính sau
      timestamp
    ]);
  });
  
  // Ghi giao dịch
  if (txRows.length > 0) {
    txSheet.getRange(txStartRow, 1, txRows.length, 7).setValues(txRows);
    
    // Cập nhật tồn hiện tại cho từng mã (cộng dồn tổng nhập)
    const stockDelta = {}; // {material: tổng tấn nhập}
    createdTems.forEach(tem => {
      const qtyInTons = tem.donVi === 'kg' ? tem.soLuong / 1000 : tem.soLuong;
      stockDelta[tem.material] = (stockDelta[tem.material] || 0) + qtyInTons;
    });
    
    Object.keys(stockDelta).forEach(mat => {
      const currentRow = findRowByMaterial(stockSheet, mat, 4);
      if (currentRow !== -1) {
        const currentStock = parseFloat(stockSheet.getRange(currentRow, 2).getValue()) || 0;
        const newStock = currentStock + stockDelta[mat];
        stockSheet.getRange(currentRow, 2).setValue(newStock);
        stockSheet.getRange(currentRow, 3).setValue(new Date().toISOString());
        
        // Cập nhật afterStock cho các giao dịch vừa ghi (cùng mã)
        let runningStock = currentStock;
        for (let i = 0; i < createdTems.length; i++) {
          if (createdTems[i].material === mat) {
            const qtyInTons = createdTems[i].donVi === 'kg' ? createdTems[i].soLuong / 1000 : createdTems[i].soLuong;
            runningStock += qtyInTons;
            txSheet.getRange(txStartRow + i, 6).setValue(runningStock);
          }
        }
      }
    });
  }
  
  return {
    status: 'ok',
    message: 'Đã tạo ' + soLuongList.length + ' tem và ghi nhận nhập kho',
    tems: createdTems,
    count: soLuongList.length,
    transactionsCreated: txRows.length
  };
}

// Lấy danh sách tem đã tạo (100 tem mới nhất)
function getTemList() {
  const sheet = getSheet(SHEETS.TEM);
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  
  const data = sheet.getRange(4, 1, lastRow - 3, 9).getValues();
  const result = [];
  
  data.forEach(row => {
    if (row[0]) {
      result.push({
        tagNo: row[0].toString(),
        material: row[1] ? row[1].toString() : '',
        tenNVL: row[2] ? row[2].toString() : '',
        ngayNhap: formatDate(row[3]),
        soLuong: parseFloat(row[4]) || 0,
        donVi: row[5] ? row[5].toString() : '',
        ghiChu: row[6] ? row[6].toString() : '',
        timestamp: row[7] ? row[7].toString() : '',
        status: row[8] ? row[8].toString() : 'Active'
      });
    }
  });
  
  // Trả về 100 tem mới nhất (đảo ngược mảng, lấy 100 đầu)
  return result.reverse().slice(0, 100);
}

// ==========================================
// TEM LOOKUP & FIFO CHECK
// ==========================================

// Tìm 1 tem theo Tag No (đọc từ sheet Tem NVL)
// Trả về thông tin tem + row index (để update status sau)
function getTemByTagNo(params) {
  const tagNo = params.tagNo || '';
  if (!tagNo) return { status: 'error', message: 'Thiếu tagNo' };
  
  const sheet = getSheet(SHEETS.TEM);
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return { status: 'error', message: 'Chưa có tem nào' };
  
  const data = sheet.getRange(4, 1, lastRow - 3, 9).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === tagNo) {
      return {
        status: 'ok',
        tem: {
          tagNo: data[i][0].toString(),
          material: data[i][1] ? data[i][1].toString() : '',
          tenNVL: data[i][2] ? data[i][2].toString() : '',
          ngayNhap: formatDate(data[i][3]),
          soLuong: parseFloat(data[i][4]) || 0,
          donVi: data[i][5] ? data[i][5].toString() : '',
          ghiChu: data[i][6] ? data[i][6].toString() : '',
          timestamp: data[i][7] ? data[i][7].toString() : '',
          currentStatus: data[i][8] ? data[i][8].toString() : ''
        },
        rowIndex: 4 + i  // Cho update sau
      };
    }
  }
  
  return { status: 'error', message: 'Không tìm thấy tem: ' + tagNo };
}

// Kiểm tra FIFO cho 1 tem sắp xuất
// Trả về danh sách các tem cùng mã, cũ hơn, còn "Đã nhập kho"
function checkFIFO(params) {
  const tagNo = params.tagNo || '';
  if (!tagNo) return { status: 'error', message: 'Thiếu tagNo' };
  
  // Lấy tem đang được scan
  const targetResult = getTemByTagNo({ tagNo: tagNo });
  if (targetResult.status !== 'ok') return targetResult;
  
  const target = targetResult.tem;
  const material = target.material;
  const ngayNhapTarget = target.ngayNhap;
  
  // Đọc toàn bộ tem, filter các tem cùng mã, cũ hơn, còn "Đã nhập kho"
  const sheet = getSheet(SHEETS.TEM);
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return { status: 'ok', olderPallets: [] };
  
  const data = sheet.getRange(4, 1, lastRow - 3, 9).getValues();
  const olderPallets = [];
  
  data.forEach(row => {
    if (!row[0]) return;
    const rTagNo = row[0].toString();
    const rMaterial = row[1] ? row[1].toString() : '';
    const rNgayNhap = formatDate(row[3]);
    const rStatus = row[8] ? row[8].toString() : '';
    const rSoLuong = parseFloat(row[4]) || 0;
    const rDonVi = row[5] ? row[5].toString() : '';
    
    // Bỏ chính nó
    if (rTagNo === tagNo) return;
    
    // Chỉ lấy cùng mã + Đã nhập kho + Ngày nhập < ngày nhập của target
    if (rMaterial === material && rStatus === 'Đã nhập kho' && rNgayNhap < ngayNhapTarget) {
      olderPallets.push({
        tagNo: rTagNo,
        material: rMaterial,
        ngayNhap: rNgayNhap,
        soLuong: rSoLuong,
        donVi: rDonVi
      });
    }
  });
  
  // Sort theo ngày nhập tăng dần (cũ nhất lên đầu)
  olderPallets.sort((a, b) => a.ngayNhap.localeCompare(b.ngayNhap));
  
  return {
    status: 'ok',
    target: {
      tagNo: target.tagNo,
      material: material,
      ngayNhap: ngayNhapTarget
    },
    olderPallets: olderPallets,
    olderCount: olderPallets.length
  };
}

// Xử lý nhiều giao dịch xuất/nhập trong 1 lần (từ giỏ hàng scan QR)
// Params: type (IN/OUT), items (JSON: mảng [{tagNo, rowIndex}, ...])
function processMultiTransaction(params) {
  const type = params.type || 'OUT';
  if (!['IN', 'OUT'].includes(type)) {
    return { status: 'error', message: 'Loại giao dịch phải là IN hoặc OUT' };
  }
  
  let items;
  try {
    items = JSON.parse(params.items);
  } catch (e) {
    return { status: 'error', message: 'Danh sách items không hợp lệ' };
  }
  
  if (!Array.isArray(items) || items.length === 0) {
    return { status: 'error', message: 'Chưa có pallet nào để ghi nhận' };
  }
  
  const txSheet = getSheet(SHEETS.TRANSACTION);
  const stockSheet = getSheet(SHEETS.STOCK);
  const temSheet = getSheet(SHEETS.TEM);
  
  const today = formatDate(new Date());
  const timestamp = new Date().toISOString();
  const results = [];
  const errors = [];
  
  // Cập nhật status tem trước (nếu có tagNo)
  items.forEach(item => {
    if (!item.tagNo) return;
    const lookupResult = getTemByTagNo({ tagNo: item.tagNo });
    if (lookupResult.status !== 'ok') {
      errors.push('Không tìm thấy tem: ' + item.tagNo);
      return;
    }
    
    const tem = lookupResult.tem;
    const rowIndex = lookupResult.rowIndex;
    
    // Validate: khi xuất thì tem phải "Đã nhập kho"; khi nhập thì tem phải chưa "Đã nhập kho" HOẶC đang "Đã xuất kho"
    if (type === 'OUT' && tem.currentStatus !== 'Đã nhập kho') {
      errors.push('Tem ' + item.tagNo + ' không ở trạng thái "Đã nhập kho" (hiện: ' + tem.currentStatus + ')');
      return;
    }
    if (type === 'IN' && tem.currentStatus === 'Đã nhập kho') {
      errors.push('Tem ' + item.tagNo + ' đã ở trạng thái "Đã nhập kho" rồi');
      return;
    }
    
    // Cập nhật status
    const newStatus = type === 'IN' ? 'Đã nhập kho' : 'Đã xuất kho';
    temSheet.getRange(rowIndex, 9).setValue(newStatus);
    
    // Số lượng theo tấn (giao dịch tồn theo tấn)
    const qtyInTons = tem.donVi === 'kg' ? tem.soLuong / 1000 : tem.soLuong;
    
    results.push({
      tagNo: tem.tagNo,
      material: tem.material,
      qtyInTons: qtyInTons,
      soLuong: tem.soLuong,
      donVi: tem.donVi,
      ghiChu: '[Tem ' + tem.tagNo + '] ' + (item.ghiChu || tem.ghiChu || '')
    });
  });
  
  if (results.length === 0) {
    return { 
      status: 'error', 
      message: 'Không xử lý được pallet nào',
      errors: errors 
    };
  }
  
  // Cập nhật tồn hiện tại + ghi giao dịch cho từng pallet
  const stockDelta = {}; // {material: delta tấn}
  results.forEach(r => {
    const delta = type === 'IN' ? r.qtyInTons : -r.qtyInTons;
    stockDelta[r.material] = (stockDelta[r.material] || 0) + delta;
  });
  
  // Cập nhật tồn hiện tại cho từng mã và tính afterStock
  const stockAfter = {}; // {material: tồn sau khi làm hết}
  Object.keys(stockDelta).forEach(mat => {
    const currentRow = findRowByMaterial(stockSheet, mat, 4);
    if (currentRow !== -1) {
      const currentStock = parseFloat(stockSheet.getRange(currentRow, 2).getValue()) || 0;
      const newStock = Math.max(0, currentStock + stockDelta[mat]);
      stockSheet.getRange(currentRow, 2).setValue(newStock);
      stockSheet.getRange(currentRow, 3).setValue(timestamp);
      stockAfter[mat] = { start: currentStock, current: currentStock };
    } else {
      stockAfter[mat] = { start: 0, current: 0 };
    }
  });
  
  // Ghi từng giao dịch với afterStock đúng thứ tự
  const txStartRow = findLastDataRow(txSheet, 4) + 1;
  const txRows = [];
  
  results.forEach(r => {
    const delta = type === 'IN' ? r.qtyInTons : -r.qtyInTons;
    stockAfter[r.material].current += delta;
    const afterStock = Math.max(0, stockAfter[r.material].current);
    
    txRows.push([
      today,
      type,
      r.material,
      r.qtyInTons,
      r.ghiChu,
      afterStock,
      timestamp
    ]);
  });
  
  if (txRows.length > 0) {
    txSheet.getRange(txStartRow, 1, txRows.length, 7).setValues(txRows);
  }
  
  return {
    status: 'ok',
    message: 'Đã ghi ' + results.length + ' giao dịch ' + (type === 'IN' ? 'nhập' : 'xuất'),
    processedCount: results.length,
    processed: results.map(r => ({ tagNo: r.tagNo, material: r.material })),
    errors: errors.length > 0 ? errors : undefined
  };
}

// ==========================================
// MASTER DATA - Đọc Mã & Tên NVL từ file KHSX external
// ==========================================
const MASTER_KHSX_SHEET_ID = "1WMF1EoGsmKNVaYIQwEe9i9tw6k_dZC1gSbHJp7RNsC0";
const MASTER_KHSX_SHEET_NAME = "Master";
const MASTER_KHSX_COL_TEN = "V";  // Cột V = Tên NVL
const MASTER_KHSX_COL_MA = "W";   // Cột W = Mã NVL

// Lấy danh sách Mã & Tên NVL từ file KHSX Master (cột V:W từ dòng 7)
// Cache 10 phút để tránh gọi external sheet quá nhiều lần
function getMasterMaterials() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('masterMaterials_v2');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // cache hỏng, đọc lại
    }
  }
  
  try {
    const ss = SpreadsheetApp.openById(MASTER_KHSX_SHEET_ID);
    const sheet = ss.getSheetByName(MASTER_KHSX_SHEET_NAME);
    if (!sheet) {
      return { status: 'error', message: 'Sheet "Master" không tồn tại trong file KHSX' };
    }
    
    // Đọc cột V:W từ dòng 7 trở đi (bỏ 6 dòng header)
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return { status: 'ok', materials: [] };
    
    const range = sheet.getRange(MASTER_KHSX_COL_TEN + '7:' + MASTER_KHSX_COL_MA + lastRow).getValues();
    
    // Lọc dòng có mã và bỏ dòng trống
    const materials = [];
    const seenCodes = {};  // dedupe theo mã
    
    range.forEach(row => {
      // V (index 0) = Tên NVL, W (index 1) = Mã NVL
      const ten = row[0] ? row[0].toString().trim() : '';
      const ma = row[1] ? row[1].toString().trim() : '';
      
      if (!ma) return;
      
      // Dedupe
      if (seenCodes[ma]) return;
      seenCodes[ma] = true;
      
      materials.push({ ma: ma, ten: ten || ma });
    });
    
    const result = { 
      status: 'ok', 
      materials: materials,
      count: materials.length,
      cached: false
    };
    
    // Cache 10 phút (600 giây)
    cache.put('masterMaterials_v2', JSON.stringify(result), 600);
    
    return result;
  } catch (error) {
    return { status: 'error', message: 'Lỗi đọc file KHSX Master: ' + error.toString() };
  }
}

// Helper: check 1 mã có tồn tại trong Master không (có cache)
// Trả về { valid: bool, ten: string }
function checkMasterMaterial(code) {
  if (!code) return { valid: false, ten: '' };
  
  const masterData = getMasterMaterials();
  if (masterData.status !== 'ok' || !masterData.materials) {
    // Nếu không đọc được Master → fallback về MATERIALS list nội bộ
    return { 
      valid: MATERIALS.indexOf(code) !== -1, 
      ten: code
    };
  }
  
  const found = masterData.materials.find(m => m.ma === code);
  if (found) {
    return { valid: true, ten: found.ten || code };
  }
  return { valid: false, ten: '' };
}

/**
 * HELPER FUNCTIONS
 */

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" không tồn tại');
  }
  return sheet;
}

function formatDate(date) {
  if (!date) return '';
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return date.toString();
}

function findLastDataRow(sheet, startRow) {
  const data = sheet.getRange(startRow, 1, 10000, 1).getValues();
  let lastRow = startRow - 1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      lastRow = startRow + i;
    }
  }
  return lastRow;
}

function findRowByMaterial(sheet, material, startRow) {
  const data = sheet.getRange(startRow, 1, 100, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === material) {
      return startRow + i;
    }
  }
  return -1;
}

function updateStockForMaterial(sheet, material, stock) {
  const rowIndex = findRowByMaterial(sheet, material, 4);
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, 2).setValue(stock);
    sheet.getRange(rowIndex, 3).setValue(new Date().toISOString());
  }
}

function returnJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function returnError(message) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'error', error: message
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * TEST FUNCTIONS
 */
function testConnection() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('✅ Kết nối OK: ' + ss.getName());
    ss.getSheets().forEach(s => Logger.log('  - ' + s.getName()));
  } catch (e) {
    Logger.log('❌ Lỗi: ' + e.toString());
  }
}

function testGetAllData() {
  const result = getAllData();
  Logger.log('Opening: ' + JSON.stringify(result.opening));
  Logger.log('Transactions: ' + result.transactions.length);
  Logger.log('Settings: ' + JSON.stringify(result.settings));
  Logger.log('Current Stock: ' + JSON.stringify(result.currentStock));
}