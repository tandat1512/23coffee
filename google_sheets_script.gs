/**
 * =========================================================================
 * 23 COFFEE — GOOGLE APPS SCRIPT ĐỒNG BỘ CÔNG THỨC CHUẨN 100% VỚI EXCEL
 * =========================================================================
 * ĐỒNG BỘ HÓA HOÀN TOÀN GIỮA FILE EXCEL (.XLSX) VÀ GOOGLE SHEETS
 * - Sử dụng cùng hệ thống công thức liên kết động giữa 5 Sheet
 * - Tự động nhận diện dấu chấm phẩy (;) cho Google Sheets Việt Nam
 * - Sử dụng hàm ROUNDUP đồng nhất thay cho CEILING (loại bỏ 100% lỗi cú pháp)
 * - Tự động liên kết Sheet 1 -> Sheet 3 -> Sheet 4 -> Sheet 5
 * =========================================================================
 * HƯỚNG DẪN CẬP NHẬT TRÊN GOOGLE SHEETS (30 GIÂY):
 * 1. Mở file Google Sheets: https://docs.google.com/spreadsheets/d/1uGRxJEPZljoVKlbBXRatjPiIpk2NBnXldwWSfVYVy_Y/edit
 * 2. Chọn menu "Tiện ích mở rộng" (Extensions) -> "Apps Script".
 * 3. Xóa toàn bộ code cũ trong file Code.gs, DÁN TOÀN BỘ file này vào.
 * 4. Nhấn biểu tượng Đĩa mềm (Save / Lưu) hoặc Ctrl+S.
 * 5. Bấm nút xanh "Triển khai" (Deploy) góc trên bên phải -> "Quản lý bản triển khai" (Manage deployments).
 * 6. Bấm biểu tượng cây bút chì (Chỉnh sửa) -> Ở mục "Phiên bản" chọn "Phiên bản mới" -> Bấm "Triển khai".
 * 7. Bấm "Xong". Sau đó quay lại Web bấm "Đẩy Lên Google Sheet" để trải nghiệm kết quả mượt mà!
 * =========================================================================
 */

function doPost(e) {
  try {
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 0. Lưu bản sao lưu hệ thống toàn vẹn vào sheet ẩn _SYSTEM_DATA_
    // (Bảo đảm khi Render hay server reset, 100% Menu, Công thức, Nguyên liệu, Chi phí được phục hồi)
    saveSystemBackup(ss, contents);

    // Tự động nhận diện dấu phân cách công thức (; cho Tiếng Việt / , cho Tiếng Anh)
    var sep = getFormulaSeparator(ss);

    // 1. Cập nhật Sheet 1: Database Nguyên Liệu
    var matNameToRow = {};
    if (data.materials && Array.isArray(data.materials)) {
      matNameToRow = updateMaterialsSheet(ss, data.materials, sep);
    }

    // 2. Cập nhật Sheet 2: Vật Tư Bao Bì
    var pkgCodeToRow = {};
    if (data.packaging && Array.isArray(data.packaging)) {
      pkgCodeToRow = updatePackagingSheet(ss, data.packaging, sep);
    }

    // 3. Cập nhật Sheet 3: Recipe SOP Chi Tiết (Có dòng Tổng Cost từng món)
    var drinkRecipeRanges = {};
    if (data.drinks && Array.isArray(data.drinks)) {
      drinkRecipeRanges = updateRecipeSheet(ss, data.drinks, data.materials, sep, matNameToRow);
    }

    // 4. Cập nhật Sheet 4: Bảng Cost & Định Giá Menu (Liên kết động Sheet 2 & Sheet 3)
    var avgRow = 5;
    if (data.drinks && Array.isArray(data.drinks)) {
      avgRow = updateMenuSheet(ss, data.drinks, data.packaging, data.materials, sep, drinkRecipeRanges, pkgCodeToRow);
    }

    // 5. Cập nhật Sheet 5: Điểm Hòa Vốn & Dòng Tiền (Liên kết động Sheet 4)
    if (data.fixed_costs || data.target_profit) {
      updateBreakEvenSheet(ss, data.fixed_costs, data.target_profit, data.drinks, sep, avgRow);
    }

    var result = {
      status: "success",
      message: "Đồng bộ dữ liệu thành công! Toàn bộ 5 sheet đã khớp 100% với file Excel!",
      separator_used: sep,
      timestamp: new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" })
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Ưu tiên lấy từ bản sao lưu hoàn chỉnh _SYSTEM_DATA_ (Đầy đủ cả món, công thức, bao bì, chi phí)
    var backupData = loadSystemBackup(ss);
    if (backupData && backupData.materials && backupData.materials.length > 0) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        data: backupData,
        source: "system_backup",
        timestamp: new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" })
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Nếu chưa có sheet ẩn, tự động xuất từ các sheet trực tiếp
    var exportedData = exportDataFromSheets(ss);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: exportedData,
      source: "sheets_export",
      timestamp: new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" })
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------
// SAO LƯU TOÀN DIỆN VÀO SHEET ẨN _SYSTEM_DATA_
// -------------------------------------------------------------
function saveSystemBackup(ss, contents) {
  try {
    var sheet = ss.getSheetByName("_SYSTEM_DATA_");
    if (!sheet) {
      sheet = ss.insertSheet("_SYSTEM_DATA_");
      sheet.hideSheet();
    } else {
      sheet.clear();
    }
    // Chia nhỏ dữ liệu phòng trường hợp chuỗi vượt giới hạn ô (50.000 ký tự)
    var chunkSize = 40000;
    var chunks = [];
    for (var i = 0; i < contents.length; i += chunkSize) {
      chunks.push([contents.substring(i, i + chunkSize)]);
    }
    sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
  } catch (err) {
    Logger.log("Lỗi lưu _SYSTEM_DATA_: " + err);
  }
}

function loadSystemBackup(ss) {
  try {
    var sheet = ss.getSheetByName("_SYSTEM_DATA_");
    if (!sheet || sheet.getLastRow() < 1) return null;
    var values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    var fullStr = "";
    for (var i = 0; i < values.length; i++) {
      fullStr += values[i][0];
    }
    if (fullStr && fullStr.trim().startsWith("{")) {
      return JSON.parse(fullStr);
    }
  } catch (err) {
    Logger.log("Lỗi đọc _SYSTEM_DATA_: " + err);
  }
  return null;
}

// -------------------------------------------------------------
// HELPER: NHẬN DIỆN DẤU PHÂN CÁCH CÔNG THỨC (; HOẶC ,)
// -------------------------------------------------------------
function getFormulaSeparator(ss) {
  try {
    var locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
    // Tiếng Anh (Mỹ, Anh) dùng dấu phẩy
    if (locale.indexOf("en") === 0) {
      return ",";
    }
  } catch(e) {}
  // Mặc định cho Việt Nam (vi_VN) và các định dạng chuẩn là dấu chấm phẩy ;
  return ";";
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// -------------------------------------------------------------
// 1. UPDATE SHEET 1: DATABASE NGUYÊN LIỆU
// -------------------------------------------------------------
function updateMaterialsSheet(ss, materials, sep) {
  var sheet = getOrCreateSheet(ss, "1. Database Nguyên Liệu");
  sheet.clear();

  sheet.getRange("A1").setValue("DATABASE NGUYÊN LIỆU GỐC — 23 COFFEE (MASTER RAW MATERIALS)")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2").setValue("Nhập giá mua sỉ và quy cách thực tế. Cột 'Đơn Giá Cơ Sở' tự động tính toán để đồng bộ toàn bộ Menu.")
    .setFontSize(10).setFontStyle("italic").setFontColor("#555555");

  var headers = [
    "Mã NL", "Tên Nguyên Liệu", "Quy Cách Mua Sỉ", "Đơn Vị Mua", 
    "Khối Lượng / Dung Tích", "Giá Mua Sỉ (VNĐ)", "Đơn Giá Cơ Sở (VNĐ/đơn vị)", 
    "Đơn Vị Cơ Sở", "Ghi Chú Sử Dụng"
  ];

  var headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  var matNameToRow = {};

  if (materials.length > 0) {
    var rows = [];
    for (var i = 0; i < materials.length; i++) {
      var m = materials[i];
      var r = 5 + i;
      var mName = (m.name || "").trim();
      matNameToRow[mName.toLowerCase()] = r;

      rows.push([
        m.id || ("NL" + (i < 9 ? "0" : "") + (i + 1)),
        mName,
        m.pack || "Gói 1kg",
        m.unit_buy || "kg",
        Number(m.qty || 1000),
        Number(m.price || 0),
        "=IF(E" + r + ">0" + sep + " ROUND(F" + r + "/E" + r + sep + " 2)" + sep + " 0)",
        m.unit_base || "g",
        m.note || ""
      ]);
    }

    var dataRange = sheet.getRange(5, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(5, 5, rows.length, 2).setNumberFormat("#,##0");
    sheet.getRange(5, 7, rows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");
    sheet.getRange(5, 1, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(5, 4, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(5, 8, rows.length, 1).setHorizontalAlignment("center");
  }

  sheet.autoResizeColumns(1, headers.length);
  return matNameToRow;
}

// -------------------------------------------------------------
// 2. UPDATE SHEET 2: VẬT TƯ BAO BÌ
// -------------------------------------------------------------
function updatePackagingSheet(ss, packaging, sep) {
  var sheet = getOrCreateSheet(ss, "2. Vật Tư Bao Bì");
  sheet.clear();

  sheet.getRange("A1").setValue("VẬT TƯ BAO BÌ & ĐÓNG GÓI TIÊU HAO (PACKAGING)")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2").setValue("Chuẩn hóa chi phí ly, nắp, túi, muỗng, ống hút cho đồ uống mang đi và tại quán.")
    .setFontSize(10).setFontStyle("italic").setFontColor("#555555");

  var headers = [
    "Mã Gói", "Quy Cách Phục Vụ / Size", "Ly (Cốc)", "Nắp / Màng Ép", 
    "Ống Hút", "Túi Quai Xách", "Muỗng / Khăn Giấy", "Tổng Chi Phí Bao Bì/Ly (VNĐ)", "Ứng Dụng Thực Tế"
  ];

  var headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  var pkgCodeToRow = {};

  if (packaging.length > 0) {
    var rows = [];
    for (var i = 0; i < packaging.length; i++) {
      var p = packaging[i];
      var r = 5 + i;
      var code = p.id || ("BB" + (i < 9 ? "0" : "") + (i + 1));
      pkgCodeToRow[code] = r;

      rows.push([
        code,
        p.name || "",
        Number(p.cup || 0),
        Number(p.lid || 0),
        Number(p.straw || 0),
        Number(p.bag || 0),
        Number(p.spoon || 0),
        "=SUM(C" + r + ":G" + r + ")",
        p.note || ""
      ]);
    }

    var dataRange = sheet.getRange(5, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(5, 3, rows.length, 6).setNumberFormat("#,##0");
    sheet.getRange(5, 8, rows.length, 1).setFontWeight("bold").setBackground("#E8F4F8");
    sheet.getRange(5, 1, rows.length, 1).setHorizontalAlignment("center");
  }

  sheet.autoResizeColumns(1, headers.length);
  return pkgCodeToRow;
}

// -------------------------------------------------------------
// 3. UPDATE SHEET 3: RECIPE SOP CHI TIẾT (CHUẨN 100% EXCEL)
// -------------------------------------------------------------
function updateRecipeSheet(ss, drinks, materials, sep, matNameToRow) {
  var sheet = getOrCreateSheet(ss, "3. Recipe SOP Chi Tiết");
  sheet.clear();

  sheet.getRange("A1").setValue("RECIPE ĐỊNH LƯỢNG CHI TIẾT MENU — 23 COFFEE")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2").setValue("Định lượng chính xác từng thành phần giúp kiểm soát chất lượng đồng đều và cấp giá vốn chuẩn xác cho Menu.")
    .setFontSize(10).setFontStyle("italic").setFontColor("#555555");

  var headers = [
    "STT", "Tên Thức Uống", "Tên Thành Phần / Nguyên Liệu", "Định Lượng Chuẩn", 
    "Đơn Vị", "Đơn Giá Cơ Sở (VNĐ)", "Cost Thành Phần (VNĐ)", "Ghi Chú Pha Chế"
  ];

  var headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#335C67").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  var drinkRecipeRanges = {};
  var currentR = 5;

  for (var i = 0; i < drinks.length; i++) {
    var d = drinks[i];
    var recipe = d.recipe || [];
    if (recipe.length === 0) {
      recipe = [{ name: "Cà phê", qty: 18, unit: "g" }];
    }
    var startR = currentR;

    var drinkRows = [];
    for (var j = 0; j < recipe.length; j++) {
      var ing = recipe[j];
      var ingName = (ing.name || "").trim();
      var key = ingName.toLowerCase();
      var matRow = matNameToRow ? matNameToRow[key] : null;
      var refFormula = matRow ? "='1. Database Nguyên Liệu'!G" + matRow : Number(ing.unit_cost || 0);
      var qty = Number(ing.qty || 0);

      drinkRows.push([
        (j === 0 ? (i + 1) : ""),
        (j === 0 ? d.name : ""),
        ingName,
        qty,
        ing.unit || "g",
        refFormula,
        "=ROUND(D" + currentR + "*F" + currentR + sep + " 0)",
        ing.note || ""
      ]);
      currentR++;
    }

    var dataRange = sheet.getRange(startR, 1, drinkRows.length, headers.length);
    dataRange.setValues(drinkRows);
    dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(startR, 4, drinkRows.length, 1).setNumberFormat("#,##0.##");
    sheet.getRange(startR, 6, drinkRows.length, 2).setNumberFormat("#,##0");
    sheet.getRange(startR, 7, drinkRows.length, 1).setFontWeight("bold");
    sheet.getRange(startR, 1, drinkRows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(startR, 5, drinkRows.length, 1).setHorizontalAlignment("center");

    // Dòng TỔNG COST NGUYÊN LIỆU cho món này (Y hệt file Excel)
    var endR = currentR - 1;
    var subTotRange = sheet.getRange(currentR, 1, 1, headers.length);
    subTotRange.setValues([[
      "",
      "TỔNG COST NGUYÊN LIỆU — " + d.name,
      "",
      "",
      "",
      "",
      "=SUM(G" + startR + ":G" + endR + ")",
      ""
    ]]);
    subTotRange.setFontWeight("bold").setBackground("#E8F4F8");
    sheet.getRange(currentR, 7).setNumberFormat("#,##0");
    subTotRange.setBorder(true, true, true, true, true, true, "#1B365D", SpreadsheetApp.BorderStyle.DOUBLE);

    drinkRecipeRanges[(d.name || "").toLowerCase().trim()] = currentR;
    currentR += 2; // Cách 1 dòng trống giữa các món
  }

  sheet.autoResizeColumns(1, headers.length);
  return drinkRecipeRanges;
}

// -------------------------------------------------------------
// 4. UPDATE SHEET 4: BẢNG COST & ĐỊNH GIÁ MENU (CHUẨN 100% EXCEL)
// -------------------------------------------------------------
function updateMenuSheet(ss, drinks, packaging, materials, sep, drinkRecipeRanges, pkgCodeToRow) {
  var sheet = getOrCreateSheet(ss, "4. Bảng Cost & Định Giá Menu");
  sheet.clear();

  sheet.getRange("A1").setValue("BẢNG TÍNH COST THỨC UỐNG & CHIẾN LƯỢC ĐỊNH GIÁ TOÀN DIỆN — 23 COFFEE")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2").setValue("Tự động liên kết: Cost Nguyên Liệu từ SOP + Cost Bao Bì chuẩn | Mục tiêu 30% Food Cost | Bạn chỉ cần đổi 'Giá Bán Lựa Chọn'")
    .setFontSize(10).setFontStyle("italic").setFontColor("#555555");

  var headers = [
    "STT", "Nhóm Món", "Tên Thức Uống", "Cost Nguyên Liệu (VNĐ)", "Cost Bao Bì (VNĐ)", 
    "% Hao Hụt", "TỔNG COST THỰC TẾ (VNĐ)", "GIÁ GỢI Ý (30% COST)", "GIÁ BẠN CHỌN BÁN (VNĐ)", 
    "% FOOD COST THỰC TẾ", "LÃI GỘP / LY (VNĐ)", "CẢNH BÁO TÀI CHÍNH"
  ];

  var headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  if (drinks.length > 0) {
    var rows = [];
    for (var i = 0; i < drinks.length; i++) {
      var d = drinks[i];
      var r = 5 + i;
      var dNameKey = (d.name || "").toLowerCase().trim();
      
      // Liên kết động sang Sheet 3 (SOP) và Sheet 2 (Bao Bì) y hệt file Excel!
      var recipeTotRow = (drinkRecipeRanges && drinkRecipeRanges[dNameKey]) ? drinkRecipeRanges[dNameKey] : (5 + i);
      var pkgRow = (pkgCodeToRow && pkgCodeToRow[d.pkg_id]) ? pkgCodeToRow[d.pkg_id] : 5;

      var formulaNl = "='3. Recipe SOP Chi Tiết'!G" + recipeTotRow;
      var formulaBb = "='2. Vật Tư Bao Bì'!H" + pkgRow;
      var wastePct = Number(d.waste_pct || 0.05);
      var sellPrice = Number(d.selling_price || 30000);

      // Công thức ROUNDUP((G*10/3); -3) làm tròn lên nghìn đồng - 100% chuẩn cả Excel lẫn Google Sheets!
      rows.push([
        (i + 1),
        d.category || "Món Khác",
        d.name || "",
        formulaNl,
        formulaBb,
        wastePct,
        "=ROUND(D" + r + "*(1+F" + r + ")+E" + r + sep + " 0)",
        "=ROUNDUP(G" + r + "*10/3" + sep + " -3)",
        sellPrice,
        "=IF(I" + r + ">0" + sep + " G" + r + "/I" + r + sep + " 0)",
        "=I" + r + "-G" + r + "",
        '=IF(J' + r + '>35%' + sep + ' "⚠️ Cost Cao (>35%) - Cần Tối Ưu"' + sep + ' IF(J' + r + '<25%' + sep + ' "💎 Lãi Rất Dày (<25%)"' + sep + ' "✅ Chuẩn Ngành (25-35%)"))'
      ]);
    }

    var dataRange = sheet.getRange(5, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
    
    // Formatting
    sheet.getRange(5, 4, rows.length, 2).setNumberFormat("#,##0");
    sheet.getRange(5, 6, rows.length, 1).setNumberFormat("0.0%");
    sheet.getRange(5, 7, rows.length, 2).setNumberFormat("#,##0").setFontWeight("bold").setBackground("#E8F4F8");
    sheet.getRange(5, 9, rows.length, 1).setNumberFormat("#,##0").setFontWeight("bold").setBackground("#E2F0D9"); // Input
    sheet.getRange(5, 10, rows.length, 1).setNumberFormat("0.0%").setFontWeight("bold");
    sheet.getRange(5, 11, rows.length, 1).setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(5, 1, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(5, 6, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(5, 10, rows.length, 1).setHorizontalAlignment("center");

    // Dòng TRUNG BÌNH TOÀN BỘ MENU (Khớp 100% Excel)
    var avgRow = 5 + drinks.length;
    var avgRowRange = sheet.getRange(avgRow, 1, 1, headers.length);
    sheet.getRange(avgRow, 1, 1, 3).merge().setValue("TRUNG BÌNH TOÀN BỘ MENU (" + drinks.length + " MÓN)").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(avgRow, 4).setFormula("=AVERAGE(D5:D" + (avgRow - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRow, 5).setFormula("=AVERAGE(E5:E" + (avgRow - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRow, 6).setFormula("=AVERAGE(F5:F" + (avgRow - 1) + ")").setNumberFormat("0.0%").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(avgRow, 7).setFormula("=AVERAGE(G5:G" + (avgRow - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRow, 8).setFormula("=AVERAGE(H5:H" + (avgRow - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRow, 9).setFormula("=AVERAGE(I5:I" + (avgRow - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRow, 10).setFormula("=AVERAGE(J5:J" + (avgRow - 1) + ")").setNumberFormat("0.0%").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(avgRow, 11).setFormula("=AVERAGE(K5:K" + (avgRow - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRow, 12).setValue("Chỉ số trung bình");
    avgRowRange.setBackground("#F3F4F6").setBorder(true, true, true, true, true, true, "#1B365D", SpreadsheetApp.BorderStyle.DOUBLE);
  }

  sheet.autoResizeColumns(1, headers.length);
  return (5 + drinks.length);
}

// -------------------------------------------------------------
// 5. UPDATE SHEET 5: ĐIỂM HÒA VỐN & DÒNG TIỀN (CHUẨN 100% EXCEL)
// -------------------------------------------------------------
function updateBreakEvenSheet(ss, fixedCosts, targetProfit, drinks, sep, avgRow) {
  var sheet = getOrCreateSheet(ss, "5. Điểm Hòa Vốn & Dòng Tiền");
  sheet.clear();

  sheet.getRange("A1").setValue("MÔ HÌNH DỰ TOÁN ĐIỂM HÒA VỐN & KẾ HOẠCH DÒNG TIỀN — 23 COFFEE")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2").setValue("Tính toán chuẩn xác: Mỗi ngày quán phải bán tối thiểu bao nhiêu ly để không lỗ và đạt lợi nhuận kỳ vọng.")
    .setFontSize(10).setFontStyle("italic").setFontColor("#555555");

  // 1. DỰ TOÁN CHI PHÍ CỐ ĐỊNH HÀNG THÁNG
  sheet.getRange("B4").setValue("1. DỰ TOÁN CHI PHÍ CỐ ĐỊNH HÀNG THÁNG (FIXED OPERATING COSTS)")
    .setFontSize(11).setFontWeight("bold").setFontColor("#1B365D");

  var fcHeaders = ["STT", "Hạng Mục Chi Phí Cố Định", "Chi Phí Hàng Tháng (VNĐ)", "Tỷ Lệ / Tổng Chi Phí", "Ghi Chú Vận Hành"];
  var fcHeaderRange = sheet.getRange(5, 2, 1, fcHeaders.length);
  fcHeaderRange.setValues([fcHeaders]);
  fcHeaderRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  var fcList = fixedCosts || [
    { name: "Tiền thuê mặt bằng kinh doanh", cost: 15400000, note: "Mức thuê quán ki-ốt / take-away" },
    { name: "Lương nhân viên pha chế & phục vụ (Part-time)", cost: 11600000, note: "2 ca x 4-5 tiếng/ngày x 22.000đ/giờ" },
    { name: "Tiền điện (Máy pha, tủ mát, máy xay, máy lạnh)", cost: 3500000, note: "Biến phí điện năng thiết bị quầy bar" },
    { name: "Tiền nước máy sinh hoạt & lọc", cost: 500000, note: "Rửa dụng cụ, vệ sinh" },
    { name: "Khấu hao máy móc & dụng cụ pha chế", cost: 1000000, note: "Phân bổ 24-36 tháng máy pha & dụng cụ" },
    { name: "Marketing, biển bảng, POS bán hàng", cost: 800000, note: "Phần mềm bán hàng & quảng bá" },
    { name: "Chi phí dự phòng sửa chữa nhỏ quầy bar", cost: 500000, note: "Bảo trì định kỳ gioăng, lưỡi xay" }
  ];

  var sumFixedR = 6 + fcList.length;
  var fcRows = [];

  for (var i = 0; i < fcList.length; i++) {
    var item = fcList[i];
    var costVal = Number(item.cost || 0);
    var r = 6 + i;
    fcRows.push([
      (i + 1),
      item.name || "",
      costVal,
      '=IF($D$' + sumFixedR + '>0' + sep + ' D' + r + '/$D$' + sumFixedR + sep + ' 0)',
      item.note || ""
    ]);
  }

  var fcDataRange = sheet.getRange(6, 2, fcRows.length, fcHeaders.length);
  fcDataRange.setValues(fcRows);
  fcDataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(6, 4, fcRows.length, 1).setNumberFormat("#,##0");
  sheet.getRange(6, 5, fcRows.length, 1).setNumberFormat("0.0%").setHorizontalAlignment("center");

  // Dòng Tổng Cộng Chi Phí Cố Định
  sheet.getRange(sumFixedR, 2, 1, 2).setValues([["", "TỔNG CHI PHÍ CỐ ĐỊNH / THÁNG"]]).setFontWeight("bold").setBackground("#F2F4F8");
  sheet.getRange(sumFixedR, 4).setFormula("=SUM(D6:D" + (sumFixedR - 1) + ")")
    .setFontWeight("bold").setNumberFormat("#,##0").setBackground("#E8F4F8");
  sheet.getRange(sumFixedR, 5).setFormula("=SUM(E6:E" + (sumFixedR - 1) + ")")
    .setFontWeight("bold").setNumberFormat("0.0%").setHorizontalAlignment("center");
  sheet.getRange(sumFixedR, 2, 1, fcHeaders.length).setBorder(true, true, true, true, true, true, "#1B365D", SpreadsheetApp.BorderStyle.DOUBLE);

  // 2. CÔNG THỨC TÍNH ĐIỂM HÒA VỐN THỰC CHIẾN (Khớp 100% Excel)
  var beStartR = sumFixedR + 2;
  sheet.getRange("B" + beStartR).setValue("2. CÔNG THỨC TÍNH ĐIỂM HÒA VỐN THỰC CHIẾN (BREAK-EVEN POINT)")
    .setFontSize(11).setFontWeight("bold").setFontColor("#1B365D");

  var beHdrR = beStartR + 1;
  var headers5_2 = ["Chỉ Số Tài Chính", "Giá Trị Tính Toán", "Đơn Vị", "Giải Thích Ý Nghĩa & Ứng Dụng", "Hành Động Của Chủ Quán"];
  var beHdrRange = sheet.getRange(beHdrR, 2, 1, headers5_2.length);
  beHdrRange.setValues([headers5_2]);
  beHdrRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  var r_aov = beHdrR + 1;
  var r_cm = beHdrR + 2;
  var r_bemonth = beHdrR + 3;
  var r_beday = beHdrR + 4;
  var r_revday = beHdrR + 5;
  var r_revmonth = beHdrR + 6;
  var r_target = beHdrR + 7;
  var r_targetday = beHdrR + 8;

  var targetProf = Number(targetProfit || 15000000);

  var beItems = [
    [r_aov, "Giá bán trung bình 1 ly (AOV)", "='4. Bảng Cost & Định Giá Menu'!I" + avgRow, "VNĐ/ly", "Giá trung bình khách chi tiêu tại quán"],
    [r_cm, "Lãi gộp trung bình trên 1 ly (CM)", "='4. Bảng Cost & Định Giá Menu'!K" + avgRow, "VNĐ/ly", "Số tiền lời trên mỗi ly để bù tiền mặt bằng & nhân sự"],
    [r_bemonth, "Số ly hòa vốn / tháng (Break-even Monthly)", "=ROUNDUP(D" + sumFixedR + "/C" + r_cm + sep + " 0)", "Ly/tháng", "Tổng số ly phải bán trong tháng để đạt điểm 0 đồng lãi/lỗ"],
    [r_beday, "SỐ LY HÒA VỐN / NGÀY (Break-even Daily)", "=ROUNDUP(C" + r_bemonth + "/30" + sep + " 0)", "Ly/ngày", "MỤC TIÊU BẮT BUỘC: Mỗi ngày mở mắt ra phải bán đủ số này!"],
    [r_revday, "Doanh thu hòa vốn / ngày", "=C" + r_beday + "*C" + r_aov, "VNĐ/ngày", "Doanh thu tối thiểu cần quẹt máy mỗi ngày"],
    [r_revmonth, "Doanh thu hòa vốn / tháng", "=C" + r_bemonth + "*C" + r_aov, "VNĐ/tháng", "Mức doanh thu chạm ngưỡng an toàn"],
    [r_target, "Mục tiêu Lợi Nhuận Ròng kỳ vọng", targetProf, "VNĐ/tháng", "Tiền chủ quán muốn bỏ túi mang về sau khi trừ hết chi phí"],
    [r_targetday, "SỐ LY CẦN BÁN ĐỂ ĐẠT LÃI KỲ VỌNG / NGÀY", "=ROUNDUP((D" + sumFixedR + "+C" + r_target + ")/(C" + r_cm + "*30)" + sep + " 0)", "Ly/ngày", "Mục tiêu kinh doanh thực tế để quán phát triển bền vững"]
  ];

  for (var k = 0; k < beItems.length; k++) {
    var item = beItems[k];
    var curR = item[0];
    var metric = item[1];
    var val = item[2];
    var unit = item[3];
    var expl = item[4];

    sheet.getRange("B" + curR).setValue(metric).setFontWeight(curR === r_beday || curR === r_targetday ? "bold" : "normal");
    var valCell = sheet.getRange("C" + curR);
    if (typeof val === "string" && val.indexOf("=") === 0) {
      valCell.setFormula(val);
    } else {
      valCell.setValue(val);
    }
    valCell.setFontWeight("bold").setNumberFormat("#,##0");
    if (curR === r_beday || curR === r_targetday) {
      valCell.setBackground("#FEF3C7").setFontColor("#92400E"); // Amber highlight
    }
    sheet.getRange("D" + curR).setValue(unit).setHorizontalAlignment("center");
    sheet.getRange("E" + curR).setValue(expl).setFontStyle("italic").setFontColor("#666666");
    sheet.getRange(curR, 2, 1, 4).setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.autoResizeColumns(1, 8);
}

// -------------------------------------------------------------
// XUẤT DỮ LIỆU TỪ SHEET VỀ JSON
// -------------------------------------------------------------
function exportDataFromSheets(ss) {
  var s1 = ss.getSheetByName("1. Database Nguyên Liệu");
  var materials = [];
  if (s1 && s1.getLastRow() >= 5) {
    var data1 = s1.getRange(5, 1, s1.getLastRow() - 4, 9).getValues();
    for (var i = 0; i < data1.length; i++) {
      if (data1[i][1]) {
        materials.push({
          id: data1[i][0] || ("NL" + (i < 9 ? "0" : "") + (i + 1)),
          name: data1[i][1],
          pack: data1[i][2],
          unit_buy: data1[i][3],
          qty: Number(data1[i][4]),
          price: Number(data1[i][5]),
          unit_cost: Number(data1[i][6]),
          unit_base: data1[i][7] || "g",
          note: data1[i][8] || ""
        });
      }
    }
  }

  // Đọc Sheet 2: Vật Tư Bao Bì
  var s2 = ss.getSheetByName("2. Vật Tư Bao Bì");
  var packaging = [];
  if (s2 && s2.getLastRow() >= 5) {
    var data2 = s2.getRange(5, 1, s2.getLastRow() - 4, 9).getValues();
    for (var j = 0; j < data2.length; j++) {
      if (data2[j][1]) {
        packaging.push({
          id: data2[j][0] || ("BB0" + (j + 1)),
          name: data2[j][1],
          cup: Number(data2[j][2] || 0),
          lid: Number(data2[j][3] || 0),
          straw: Number(data2[j][4] || 0),
          bag: Number(data2[j][5] || 0),
          spoon: Number(data2[j][6] || 0),
          total: Number(data2[j][7] || 0),
          note: data2[j][8] || ""
        });
      }
    }
  }

  // Đọc Sheet 5: Chi Phí Cố Định
  var s5 = ss.getSheetByName("5. Điểm Hòa Vốn & Dòng Tiền");
  var fixedCosts = [];
  var targetProfit = 15000000;
  if (s5 && s5.getLastRow() >= 6) {
    var data5 = s5.getRange(6, 2, Math.min(15, s5.getLastRow() - 5), 5).getValues();
    for (var k = 0; k < data5.length; k++) {
      var itemName = data5[k][1];
      if (itemName && itemName.indexOf("TỔNG CHI PHÍ") === -1) {
        fixedCosts.push({
          name: itemName,
          cost: Number(data5[k][2] || 0),
          note: data5[k][4] || ""
        });
      }
    }
  }

  return {
    materials: materials,
    packaging: packaging.length > 0 ? packaging : undefined,
    fixed_costs: fixedCosts.length > 0 ? fixedCosts : undefined,
    target_profit: targetProfit
  };
}
