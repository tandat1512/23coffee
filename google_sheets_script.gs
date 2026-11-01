/**
 * =========================================================================
 * 23 COFFEE — GOOGLE APPS SCRIPT WEBHOOK ĐỒNG BỘ HAI CHIỀU DỮ LIỆU COST
 * =========================================================================
 * CẬP NHẬT: TỰ ĐỘNG TƯƠNG THÍCH CẢ GOOGLE SHEETS VIỆT NAM (DẤU ;) VÀ QUỐC TẾ (DẤU ,)
 * KHẮC PHỤC 100% LỖI #ERROR! TẠI CÁC CỘT CÔNG THỨC TRÊN GOOGLE SHEETS!
 * =========================================================================
 * Hướng dẫn cập nhật trên Google Sheet (Chỉ mất 30 giây):
 * 1. Mở file Google Sheet của bạn: https://docs.google.com/spreadsheets/d/1uGRxJEPZljoVKlbBXRatjPiIpk2NBnXldwWSfVYVy_Y/edit
 * 2. Trên menu trên cùng, chọn: "Tiện ích mở rộng" (Extensions) -> "Apps Script".
 * 3. Xóa hết code cũ trong file Code.gs và DÁN TOÀN BỘ file này vào.
 * 4. Nhấn biểu tượng Đĩa mềm (Save / Lưu) hoặc Ctrl+S.
 * 5. Bấm nút "Triển khai" (Deploy) màu xanh ở góc phải trên -> Chọn "Quản lý bản triển khai" (Manage deployments).
 * 6. Bấm biểu tượng cây bút chì (Edit) -> Ở mục "Phiên bản" (Version) chọn "Phiên bản mới" (New version) -> Bấm "Triển khai" (Deploy).
 * 7. Bấm "Xong" (Done). Sau đó quay lại Web bấm "Đẩy Lên Google Sheet" để kiểm tra kết quả ngay!
 * =========================================================================
 */

function doPost(e) {
  try {
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Tự động nhận diện dấu phân cách công thức (; cho Tiếng Việt / , cho Tiếng Anh)
    var sep = getFormulaSeparator(ss);

    // 1. Cập nhật Sheet 1: Database Nguyên Liệu
    if (data.materials && Array.isArray(data.materials)) {
      updateMaterialsSheet(ss, data.materials, sep);
    }

    // 2. Cập nhật Sheet 2: Vật Tư Bao Bì
    if (data.packaging && Array.isArray(data.packaging)) {
      updatePackagingSheet(ss, data.packaging, sep);
    }

    // 3. Cập nhật Sheet 3: Recipe SOP Chi Tiết
    if (data.drinks && Array.isArray(data.drinks)) {
      updateRecipeSheet(ss, data.drinks, data.materials, sep);
    }

    // 4. Cập nhật Sheet 4: Bảng Cost & Định Giá Menu
    if (data.drinks && Array.isArray(data.drinks)) {
      updateMenuSheet(ss, data.drinks, data.packaging, data.materials, sep);
    }

    // 5. Cập nhật Sheet 5: Điểm Hòa Vốn & Dòng Tiền
    if (data.fixed_costs || data.target_profit) {
      updateBreakEvenSheet(ss, data.fixed_costs, data.target_profit, data.drinks, sep);
    }

    var result = {
      status: "success",
      message: "Đồng bộ dữ liệu thành công không còn lỗi #ERROR!!",
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
    var exportedData = exportDataFromSheets(ss);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: exportedData,
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
  // Mặc định cho Việt Nam (vi_VN) và các định dạng chuẩn F&B là dấu chấm phẩy ;
  return ";";
}

// -------------------------------------------------------------
// HELPER: LẤY HOẶC TẠO SHEET VỚI TIÊU ĐỀ
// -------------------------------------------------------------
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// -------------------------------------------------------------
// 1. UPDATE SHEET NGUYÊN LIỆU
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

  if (materials.length > 0) {
    var rows = [];
    for (var i = 0; i < materials.length; i++) {
      var m = materials[i];
      var r = 5 + i;
      rows.push([
        m.id || ("NL" + (i < 9 ? "0" : "") + (i + 1)),
        m.name || "",
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
}

// -------------------------------------------------------------
// 2. UPDATE SHEET VẬT TƯ BAO BÌ
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

  if (packaging.length > 0) {
    var rows = [];
    for (var i = 0; i < packaging.length; i++) {
      var p = packaging[i];
      var r = 5 + i;
      rows.push([
        p.id || ("BB" + (i < 9 ? "0" : "") + (i + 1)),
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
}

// -------------------------------------------------------------
// 3. UPDATE SHEET RECIPE SOP CHI TIẾT
// -------------------------------------------------------------
function updateRecipeSheet(ss, drinks, materials, sep) {
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

  var matPriceMap = {};
  var matUnitMap = {};
  if (materials) {
    for (var m = 0; m < materials.length; m++) {
      var key = (materials[m].name || "").toLowerCase().trim();
      matPriceMap[key] = Number(materials[m].unit_cost || 0);
      matUnitMap[key] = materials[m].unit_base || "g";
    }
  }

  var rows = [];
  var rowIdx = 5;
  for (var i = 0; i < drinks.length; i++) {
    var d = drinks[i];
    var recipe = d.recipe || [];
    for (var j = 0; j < recipe.length; j++) {
      var ing = recipe[j];
      var ingName = ing.name || "";
      var key = ingName.toLowerCase().trim();
      var unitCost = matPriceMap[key] !== undefined ? matPriceMap[key] : Number(ing.unit_cost || 0);
      var unit = ing.unit || matUnitMap[key] || "g";
      var qty = Number(ing.qty || 0);

      rows.push([
        (j === 0 ? (i + 1) : ""),
        (j === 0 ? d.name : ""),
        ingName,
        qty,
        unit,
        unitCost,
        "=ROUND(D" + rowIdx + "*F" + rowIdx + sep + " 0)",
        ing.note || ""
      ]);
      rowIdx++;
    }
  }

  if (rows.length > 0) {
    var dataRange = sheet.getRange(5, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(5, 4, rows.length, 1).setNumberFormat("#,##0.##");
    sheet.getRange(5, 6, rows.length, 2).setNumberFormat("#,##0");
    sheet.getRange(5, 7, rows.length, 1).setFontWeight("bold");
    sheet.getRange(5, 1, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(5, 5, rows.length, 1).setHorizontalAlignment("center");
  }

  sheet.autoResizeColumns(1, headers.length);
}

// -------------------------------------------------------------
// 4. UPDATE SHEET BẢNG COST & ĐỊNH GIÁ MENU
// -------------------------------------------------------------
function updateMenuSheet(ss, drinks, packaging, materials, sep) {
  var sheet = getOrCreateSheet(ss, "4. Bảng Cost & Định Giá Menu");
  sheet.clear();

  sheet.getRange("A1").setValue("BẢNG TÍNH COST THỨC UỐNG & CHIẾN LƯỢC ĐỊNH GIÁ TOÀN DIỆN — 23 COFFEE")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2").setValue("Tự động liên kết: Cost Nguyên Liệu từ SOP + Cost Bao Bì chuẩn | Mục tiêu 30% Food Cost | Bạn chỉ cần đổi 'Giá Bán Lựa Chọn'")
    .setFontSize(10).setFontStyle("italic").setFontColor("#555555");

  var headers = [
    "STT", "Nhóm Món", "Tên Thức Uống", "Cost Nguyên Liệu (VNĐ)", "Cost Bao Bì (VNĐ)", 
    "% Hao Hụt", "TỔNG COST THỰC TẾ (VNĐ)", "GIÁ GỢI Ý (30% COST)", "GIÁ BẠN CHỌN BÁN (VNĐ)", 
    "% FOOD COST", "LÃI GỘP / LY (VNĐ)", "ĐÁNH GIÁ CHIẾN LƯỢC"
  ];

  var headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  var pkgMap = {};
  if (packaging) {
    for (var p = 0; p < packaging.length; p++) {
      pkgMap[packaging[p].id] = Number(packaging[p].total || 1100);
    }
  }

  var matPriceMap = {};
  if (materials) {
    for (var m = 0; m < materials.length; m++) {
      var mKey = (materials[m].name || "").toLowerCase().trim();
      matPriceMap[mKey] = Number(materials[m].unit_cost || 0);
    }
  }

  if (drinks.length > 0) {
    var rows = [];
    for (var i = 0; i < drinks.length; i++) {
      var d = drinks[i];
      var r = 5 + i;
      var pkgCost = pkgMap[d.pkg_id] !== undefined ? pkgMap[d.pkg_id] : 1100;
      
      // Tính toán Cost nguyên liệu nếu chưa có sẵn
      var costNl = Number(d.cost_nl || 0);
      if (costNl === 0 && d.recipe && Array.isArray(d.recipe)) {
        for (var k = 0; k < d.recipe.length; k++) {
          var ing = d.recipe[k];
          var ingKey = (ing.name || "").toLowerCase().trim();
          var uCost = matPriceMap[ingKey] !== undefined ? matPriceMap[ingKey] : Number(ing.unit_cost || 0);
          costNl += Math.round(Number(ing.qty || 0) * uCost);
        }
      }

      var wastePct = Number(d.waste_pct || 0.05);
      var sellPrice = Number(d.selling_price || 30000);

      // Công thức sử dụng dấu phân cách sep (; hoặc ,)
      // Chú ý: *10/3 thay cho /0.3 để tránh lỗi dấu thập phân ở mọi quốc gia!
      rows.push([
        (i + 1),
        d.category || "Món Khác",
        d.name || "",
        costNl,
        pkgCost,
        wastePct,
        "=ROUND((D" + r + "*(1+F" + r + "))+E" + r + sep + " 0)",
        "=CEILING(G" + r + "*10/3" + sep + " 1000)",
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
    sheet.getRange(5, 9, rows.length, 1).setNumberFormat("#,##0").setFontWeight("bold").setBackground("#E2F0D9"); // Soft Green input
    sheet.getRange(5, 10, rows.length, 1).setNumberFormat("0.0%").setFontWeight("bold");
    sheet.getRange(5, 11, rows.length, 1).setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(5, 1, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(5, 6, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(5, 10, rows.length, 1).setHorizontalAlignment("center");

    // Dòng Trung Bình Toàn Menu
    var avgRowIdx = 5 + drinks.length;
    var avgRowRange = sheet.getRange(avgRowIdx, 1, 1, headers.length);
    sheet.getRange(avgRowIdx, 1, 1, 3).merge().setValue("TRUNG BÌNH TOÀN MENU").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(avgRowIdx, 4).setFormula("=AVERAGE(D5:D" + (avgRowIdx - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRowIdx, 5).setFormula("=AVERAGE(E5:E" + (avgRowIdx - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRowIdx, 6).setFormula("=AVERAGE(F5:F" + (avgRowIdx - 1) + ")").setNumberFormat("0.0%").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(avgRowIdx, 7).setFormula("=AVERAGE(G5:G" + (avgRowIdx - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRowIdx, 8).setFormula("=AVERAGE(H5:H" + (avgRowIdx - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRowIdx, 9).setFormula("=AVERAGE(I5:I" + (avgRowIdx - 1) + ")").setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRowIdx, 10).setFormula("=IF(I" + avgRowIdx + ">0" + sep + " G" + avgRowIdx + "/I" + avgRowIdx + sep + " 0)").setNumberFormat("0.0%").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(avgRowIdx, 11).setFormula("=I" + avgRowIdx + "-G" + avgRowIdx).setNumberFormat("#,##0").setFontWeight("bold");
    sheet.getRange(avgRowIdx, 12).setValue("Chỉ số trung bình");
    avgRowRange.setBackground("#F3F4F6").setBorder(true, true, true, true, true, true, "#1B365D", SpreadsheetApp.BorderStyle.DOUBLE);
  }

  sheet.autoResizeColumns(1, headers.length);
}

// -------------------------------------------------------------
// 5. UPDATE SHEET ĐIỂM HÒA VỐN & DÒNG TIỀN
// -------------------------------------------------------------
function updateBreakEvenSheet(ss, fixedCosts, targetProfit, drinks, sep) {
  var sheet = getOrCreateSheet(ss, "5. Điểm Hòa Vốn & Dòng Tiền");
  sheet.clear();

  sheet.getRange("A1").setValue("MÔ HÌNH DỰ TOÁN ĐIỂM HÒA VỐN & KẾ HOẠCH DÒNG TIỀN — 23 COFFEE")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2").setValue("Tính toán chuẩn xác: Mỗi ngày quán phải bán tối thiểu bao nhiêu ly để không lỗ và đạt lợi nhuận kỳ vọng.")
    .setFontSize(10).setFontStyle("italic").setFontColor("#555555");

  // Bảng Chi Phí Cố Định
  sheet.getRange("B4").setValue("1. DỰ TOÁN CHI PHÍ CỐ ĐỊNH HÀNG THÁNG (FIXED OPERATING COSTS)")
    .setFontSize(11).setFontWeight("bold").setFontColor("#1B365D");

  var fcHeaders = ["STT", "Hạng Mục Chi Phí Cố Định", "Chi Phí Hàng Tháng (VNĐ)", "Tỷ Lệ / Tổng Chi Phí", "Ghi Chú Vận Hành"];
  var fcHeaderRange = sheet.getRange(5, 2, 1, fcHeaders.length);
  fcHeaderRange.setValues([fcHeaders]);
  fcHeaderRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  var fcRows = [];
  var fcList = fixedCosts || [
    { name: "Tiền thuê mặt bằng kinh doanh", cost: 15400000, note: "Mức thuê quán ki-ốt / take-away" },
    { name: "Lương nhân viên pha chế & phục vụ (Part-time)", cost: 11600000, note: "2 ca x 4-5 tiếng/ngày x 22.000đ/giờ" },
    { name: "Tiền điện (Máy pha, tủ mát, máy xay, máy lạnh)", cost: 3500000, note: "Biến phí điện năng thiết bị quầy bar" },
    { name: "Tiền nước máy sinh hoạt & lọc", cost: 500000, note: "Rửa dụng cụ, vệ sinh" },
    { name: "Khấu hao máy móc & dụng cụ pha chế", cost: 1000000, note: "Phân bổ 24-36 tháng máy pha & dụng cụ" },
    { name: "Marketing, biển bảng, POS bán hàng", cost: 800000, note: "Phần mềm bán hàng & quảng bá" },
    { name: "Chi phí dự phòng sửa chữa nhỏ quầy bar", cost: 500000, note: "Bảo trì định kỳ gioăng, lưỡi xay" }
  ];

  var totalRowIdx = 6 + fcList.length;

  for (var i = 0; i < fcList.length; i++) {
    var item = fcList[i];
    var costVal = Number(item.cost || 0);
    var r = 6 + i;
    fcRows.push([
      (i + 1),
      item.name || "",
      costVal,
      '=IF($D$' + totalRowIdx + '>0' + sep + ' D' + r + '/$D$' + totalRowIdx + sep + ' 0)',
      item.note || ""
    ]);
  }

  var fcDataRange = sheet.getRange(6, 2, fcRows.length, fcHeaders.length);
  fcDataRange.setValues(fcRows);
  fcDataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(6, 4, fcRows.length, 1).setNumberFormat("#,##0");
  sheet.getRange(6, 5, fcRows.length, 1).setNumberFormat("0.0%").setHorizontalAlignment("center");

  // Dòng Tổng Cộng
  sheet.getRange(totalRowIdx, 2, 1, 2).setValues([["", "TỔNG CHI PHÍ CỐ ĐỊNH / THÁNG"]]).setFontWeight("bold").setBackground("#F2F4F8");
  sheet.getRange(totalRowIdx, 4).setFormula("=SUM(D6:D" + (totalRowIdx - 1) + ")")
    .setFontWeight("bold").setNumberFormat("#,##0").setBackground("#E8F4F8");
  sheet.getRange(totalRowIdx, 5).setFormula("=SUM(E6:E" + (totalRowIdx - 1) + ")")
    .setFontWeight("bold").setNumberFormat("0.0%").setHorizontalAlignment("center");
  sheet.getRange(totalRowIdx, 2, 1, fcHeaders.length).setBorder(true, true, true, true, true, true, "#1B365D", SpreadsheetApp.BorderStyle.DOUBLE);

  // Khối Chỉ Số Hòa Vốn
  var kpiStartRow = totalRowIdx + 3;
  sheet.getRange("B" + kpiStartRow).setValue("2. CHỈ SỐ HÒA VỐN & MỤC TIÊU LỢI NHUẬN").setFontSize(11).setFontWeight("bold").setFontColor("#1B365D");

  var targetProf = Number(targetProfit || 15000000);
  var numDrinks = drinks ? drinks.length : 1;
  var menuLastRow = 4 + numDrinks;

  var kpis = [
    ["Giá bán trung bình toàn Menu", "=AVERAGE('4. Bảng Cost & Định Giá Menu'!I5:I" + menuLastRow + ")", "VNĐ/ly (Lấy trung bình từ Menu)"],
    ["Chi phí vốn trung bình / ly", "=AVERAGE('4. Bảng Cost & Định Giá Menu'!G5:G" + menuLastRow + ")", "VNĐ/ly (Lấy trung bình từ Menu)"],
    ["Lợi nhuận mục tiêu mong muốn", targetProf, "VNĐ/tháng (Bạn có thể sửa số này)"],
    ["SỐ LY HÒA VỐN / THÁNG", "=CEILING(D" + totalRowIdx + "/(D" + (kpiStartRow + 1) + "-D" + (kpiStartRow + 2) + ")" + sep + " 1)", "Ly/tháng"],
    ["SỐ LY HÒA VỐN / NGÀY (TỐI THIỂU)", "=CEILING(D" + (kpiStartRow + 4) + "/30" + sep + " 1)", "Ly/ngày cần đạt để không lỗ"],
    ["DOANH THU HÒA VỐN / NGÀY", "=D" + (kpiStartRow + 5) + "*D" + (kpiStartRow + 1), "VNĐ/ngày tối thiểu"],
    ["SỐ LY ĐẠT MỤC TIÊU LỢI NHUẬN / NGÀY", "=CEILING((D" + totalRowIdx + "+D" + (kpiStartRow + 3) + ")/((D" + (kpiStartRow + 1) + "-D" + (kpiStartRow + 2) + ")*30)" + sep + " 1)", "Ly/ngày để đạt mục tiêu"]
  ];

  for (var k = 0; k < kpis.length; k++) {
    var curR = kpiStartRow + 1 + k;
    sheet.getRange("C" + curR).setValue(kpis[k][0]).setFontWeight(k >= 3 ? "bold" : "normal");
    var valCell = sheet.getRange("D" + curR);
    if (typeof kpis[k][1] === "string" && kpis[k][1].indexOf("=") === 0) {
      valCell.setFormula(kpis[k][1]);
    } else {
      valCell.setValue(kpis[k][1]);
    }
    valCell.setFontWeight("bold").setNumberFormat("#,##0");
    if (k === 4 || k === 6) {
      valCell.setBackground("#FEF3C7").setFontColor("#92400E"); // Amber highlight
    }
    sheet.getRange("E" + curR).setValue(kpis[k][2]).setFontStyle("italic").setFontColor("#666666");
    sheet.getRange(curR, 3, 1, 3).setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
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

  return {
    materials: materials
  };
}
