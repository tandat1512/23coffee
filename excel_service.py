import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_excel_from_data(data, output_filename="Bang_Cost_Va_Chien_Luoc_Dinh_Gia_23_Coffee.xlsx"):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # Fonts & Styles
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Arial", size=14, bold=True, color="1B365D")
    subtitle_font = Font(name="Arial", size=10, italic=True, color="555555")
    bold_font = Font(name="Arial", size=10, bold=True)
    regular_font = Font(name="Arial", size=10)
    italic_font = Font(name="Arial", size=9, italic=True, color="666666")

    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # Navy Blue
    recipe_header_fill = PatternFill(start_color="335C67", end_color="335C67", fill_type="solid")
    highlight_fill = PatternFill(start_color="E8F4F8", end_color="E8F4F8", fill_type="solid") # Soft Ice Blue
    user_input_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid") # Soft Green

    thin = Side(border_style="thin", color="D9D9D9")
    double_bottom = Side(border_style="double", color="1B365D")
    box_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    total_border = Border(left=thin, right=thin, top=thin, bottom=double_bottom)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    materials = data.get("materials", [])
    packaging = data.get("packaging", [])
    drinks = data.get("drinks", [])
    fixed_costs = data.get("fixed_costs", [])
    target_profit = data.get("target_profit", 15000000)

    # ----------------------------------------------------
    # SHEET 1: DATABASE NGUYÊN LIỆU
    # ----------------------------------------------------
    ws1 = wb.create_sheet(title="1. Database Nguyên Liệu")
    ws1.views.sheetView[0].showGridLines = True
    ws1["A1"] = "DATABASE NGUYÊN LIỆU GỐC — 23 COFFEE (MASTER RAW MATERIALS)"
    ws1["A1"].font = title_font
    ws1["A2"] = "Nhập giá mua sỉ và quy cách thực tế. Cột 'Đơn Giá Cơ Sở' tự động tính toán để đồng bộ toàn bộ Menu."
    ws1["A2"].font = subtitle_font

    headers1 = [
        "Mã NL", "Tên Nguyên Liệu", "Quy Cách Mua Sỉ", "Đơn Vị Mua", 
        "Khối Lượng / Dung Tích", "Giá Mua Sỉ (VNĐ)", "Đơn Giá Cơ Sở (VNĐ/đơn vị)", 
        "Đơn Vị Cơ Sở", "Ghi Chú Sử Dụng"
    ]
    for col_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = box_border

    mat_name_to_row = {}
    for row_idx, item in enumerate(materials, 5):
        code = item.get("id", f"NL{row_idx-4:02d}")
        name = item.get("name", "")
        pack = item.get("pack", "")
        u_buy = item.get("unit_buy", "")
        qty = item.get("qty", 1)
        price = item.get("price", 0)
        u_base = item.get("unit_base", "g")
        note = item.get("note", "")

        mat_name_to_row[name.lower().strip()] = row_idx

        ws1.cell(row=row_idx, column=1, value=code).alignment = align_center
        ws1.cell(row=row_idx, column=2, value=name).alignment = align_left
        ws1.cell(row=row_idx, column=3, value=pack).alignment = align_left
        ws1.cell(row=row_idx, column=4, value=u_buy).alignment = align_center
        
        c_qty = ws1.cell(row=row_idx, column=5, value=qty)
        c_qty.alignment = align_right
        c_qty.number_format = "#,##0"
        
        c_price = ws1.cell(row=row_idx, column=6, value=price)
        c_price.alignment = align_right
        c_price.number_format = "#,##0"
        
        c_unit_cost = ws1.cell(row=row_idx, column=7, value=f"=ROUND(F{row_idx}/E{row_idx}, 2)")
        c_unit_cost.alignment = align_right
        c_unit_cost.font = bold_font
        c_unit_cost.number_format = "#,##0.00"
        
        ws1.cell(row=row_idx, column=8, value=u_base).alignment = align_center
        ws1.cell(row=row_idx, column=9, value=note).alignment = align_left
        for c in range(1, 10):
            ws1.cell(row=row_idx, column=c).border = box_border

    # ----------------------------------------------------
    # SHEET 2: VẬT TƯ BAO BÌ
    # ----------------------------------------------------
    ws2 = wb.create_sheet(title="2. Vật Tư Bao Bì")
    ws2.views.sheetView[0].showGridLines = True
    ws2["A1"] = "VẬT TƯ BAO BÌ & ĐÓNG GÓI TIÊU HAO (PACKAGING)"
    ws2["A1"].font = title_font
    ws2["A2"] = "Chuẩn hóa chi phí ly, nắp, túi, muỗng, ống hút cho đồ uống mang đi và tại quán."
    ws2["A2"].font = subtitle_font

    headers2 = [
        "Mã Gói", "Quy Cách Phục Vụ / Size", "Ly (Cốc)", "Nắp / Màng Ép", 
        "Ống Hút", "Túi Quai Xách", "Muỗng / Khăn Giấy", "Tổng Chi Phí Bao Bì/Ly (VNĐ)", "Ứng Dụng Thực Tế"
    ]
    for col_idx, h in enumerate(headers2, 1):
        cell = ws2.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = box_border

    pkg_code_to_row = {}
    for row_idx, item in enumerate(packaging, 5):
        code = item.get("id", f"BB{row_idx-4:02d}")
        name = item.get("name", "")
        cup = item.get("cup", 0)
        lid = item.get("lid", 0)
        straw = item.get("straw", 0)
        bag = item.get("bag", 0)
        spoon = item.get("spoon", 0)
        tot = item.get("total", cup+lid+straw+bag+spoon)
        note = item.get("note", "")

        pkg_code_to_row[code] = row_idx

        ws2.cell(row=row_idx, column=1, value=code).alignment = align_center
        ws2.cell(row=row_idx, column=2, value=name).alignment = align_left
        for c_idx, val in enumerate([cup, lid, straw, bag, spoon], 3):
            cell = ws2.cell(row=row_idx, column=c_idx, value=val)
            cell.alignment = align_right
            cell.number_format = "#,##0"

        c_tot = ws2.cell(row=row_idx, column=8, value=f"=SUM(C{row_idx}:G{row_idx})")
        c_tot.alignment = align_right
        c_tot.font = bold_font
        c_tot.number_format = "#,##0"
        ws2.cell(row=row_idx, column=9, value=note).alignment = align_left
        for c in range(1, 10):
            ws2.cell(row=row_idx, column=c).border = box_border

    # ----------------------------------------------------
    # SHEET 3: RECIPE CHI TIẾT (SOP BARISTA)
    # ----------------------------------------------------
    ws3 = wb.create_sheet(title="3. Recipe SOP Chi Tiết")
    ws3.views.sheetView[0].showGridLines = True
    ws3["A1"] = "RECIPE ĐỊNH LƯỢNG CHI TIẾT MENU — 23 COFFEE"
    ws3["A1"].font = title_font
    ws3["A2"] = "Định lượng chính xác từng thành phần giúp kiểm soát chất lượng đồng đều và cấp giá vốn chuẩn xác cho Menu."
    ws3["A2"].font = subtitle_font

    headers3 = [
        "STT", "Tên Thức Uống", "Tên Thành Phần / Nguyên Liệu", "Định Lượng Chuẩn", 
        "Đơn Vị", "Đơn Giá Cơ Sở (VNĐ)", "Cost Thành Phần (VNĐ)", "Ghi Chú Pha Chế"
    ]
    for col_idx, h in enumerate(headers3, 1):
        cell = ws3.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = recipe_header_fill
        cell.alignment = align_center
        cell.border = box_border

    current_r3 = 5
    drink_recipe_ranges = {}

    for idx, d in enumerate(drinks, 1):
        d_name = d.get("name", "")
        recipe = d.get("recipe", [])
        start_r = current_r3

        if not recipe:
            recipe = [{"name": "Cà phê", "qty": 18, "unit": "g"}]

        for ing_idx, ing in enumerate(recipe):
            ing_name = ing.get("name", "").strip()
            qty = ing.get("qty", 0)
            unit = ing.get("unit", "g")

            mat_row = mat_name_to_row.get(ing_name.lower(), None)
            ref_formula = f"='1. Database Nguyên Liệu'!G{mat_row}" if mat_row else "0"

            ws3.cell(row=current_r3, column=1, value=idx if ing_idx == 0 else "").alignment = align_center
            ws3.cell(row=current_r3, column=2, value=d_name if ing_idx == 0 else "").alignment = align_left
            ws3.cell(row=current_r3, column=3, value=ing_name).alignment = align_left
            
            c_q = ws3.cell(row=current_r3, column=4, value=qty)
            c_q.alignment = align_right
            c_q.number_format = "#,##0.0" if isinstance(qty, float) and qty < 1 else "#,##0"
            
            ws3.cell(row=current_r3, column=5, value=unit).alignment = align_center
            
            c_ref = ws3.cell(row=current_r3, column=6, value=ref_formula)
            c_ref.alignment = align_right
            c_ref.number_format = "#,##0.00"
            
            c_cost = ws3.cell(row=current_r3, column=7, value=f"=ROUND(D{current_r3}*F{current_r3}, 0)")
            c_cost.alignment = align_right
            c_cost.font = bold_font
            c_cost.number_format = "#,##0"
            
            for c in range(1, 9):
                ws3.cell(row=current_r3, column=c).border = box_border
            current_r3 += 1

        end_r = current_r3 - 1
        ws3.cell(row=current_r3, column=2, value=f"TỔNG COST NGUYÊN LIỆU — {d_name}").font = bold_font
        ws3.cell(row=current_r3, column=2).fill = highlight_fill
        c_subtot = ws3.cell(row=current_r3, column=7, value=f"=SUM(G{start_r}:G{end_r})")
        c_subtot.font = bold_font
        c_subtot.alignment = align_right
        c_subtot.fill = highlight_fill
        c_subtot.number_format = "#,##0"
        for c in range(1, 9):
            ws3.cell(row=current_r3, column=c).border = total_border

        drink_recipe_ranges[d_name] = current_r3
        current_r3 += 2

    # ----------------------------------------------------
    # SHEET 4: BẢNG COST MENU & CHIẾN LƯỢC ĐỊNH GIÁ
    # ----------------------------------------------------
    ws4 = wb.create_sheet(title="4. Bảng Cost & Định Giá Menu")
    ws4.views.sheetView[0].showGridLines = True
    ws4["A1"] = "BẢNG TÍNH COST THỨC UỐNG & CHIẾN LƯỢC ĐỊNH GIÁ TOÀN DIỆN — 23 COFFEE"
    ws4["A1"].font = title_font
    ws4["A2"] = "Tự động liên kết: Cost Nguyên Liệu từ SOP + Cost Bao Bì chuẩn | Mục tiêu 30% Food Cost | Bạn chỉ cần đổi 'Giá Bán Lựa Chọn'"
    ws4["A2"].font = subtitle_font

    headers4 = [
        "STT", "Nhóm Món", "Tên Thức Uống", "Cost Nguyên Liệu (VNĐ)", 
        "Cost Bao Bì (VNĐ)", "% Hao Hụt", "TỔNG COST THỰC TẾ (VNĐ)", 
        "GIÁ GỢI Ý (30% COST)", "GIÁ BẠN CHỌN BÁN (VNĐ)", "% FOOD COST THỰC TẾ", 
        "LÃI GỘP / LY (VNĐ)", "CẢNH BÁO TÀI CHÍNH"
    ]
    for col_idx, h in enumerate(headers4, 1):
        cell = ws4.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = box_border

    for idx, d in enumerate(drinks, 5):
        stt = idx - 4
        category = d.get("category", "Menu Chính")
        d_name = d.get("name", "")
        price = d.get("selling_price", 30000)
        waste = d.get("waste_pct", 0.05)
        pkg_id = d.get("pkg_id", "BB01")
        pkg_row = pkg_code_to_row.get(pkg_id, 5)

        recipe_tot_row = drink_recipe_ranges.get(d_name, 5)

        ws4.cell(row=idx, column=1, value=stt).alignment = align_center
        ws4.cell(row=idx, column=2, value=category).alignment = align_left
        ws4.cell(row=idx, column=3, value=d_name).alignment = align_left

        c_nl = ws4.cell(row=idx, column=4, value=f"='3. Recipe SOP Chi Tiết'!G{recipe_tot_row}")
        c_nl.alignment = align_right
        c_nl.number_format = "#,##0"

        c_bb = ws4.cell(row=idx, column=5, value=f"='2. Vật Tư Bao Bì'!H{pkg_row}")
        c_bb.alignment = align_right
        c_bb.number_format = "#,##0"

        c_w = ws4.cell(row=idx, column=6, value=waste)
        c_w.alignment = align_right
        c_w.number_format = "0.0%"

        c_tot = ws4.cell(row=idx, column=7, value=f"=ROUND(D{idx}*(1+F{idx})+E{idx}, 0)")
        c_tot.alignment = align_right
        c_tot.font = bold_font
        c_tot.fill = highlight_fill
        c_tot.number_format = "#,##0"

        c_sug = ws4.cell(row=idx, column=8, value=f"=ROUNDUP(G{idx}*10/3, -3)")
        c_sug.alignment = align_right
        c_sug.number_format = "#,##0"

        c_price = ws4.cell(row=idx, column=9, value=price)
        c_price.alignment = align_right
        c_price.font = bold_font
        c_price.fill = user_input_fill
        c_price.number_format = "#,##0"

        c_fc = ws4.cell(row=idx, column=10, value=f"=IF(I{idx}>0, G{idx}/I{idx}, 0)")
        c_fc.alignment = align_right
        c_fc.font = bold_font
        c_fc.number_format = "0.0%"

        c_prof = ws4.cell(row=idx, column=11, value=f"=I{idx}-G{idx}")
        c_prof.alignment = align_right
        c_prof.font = bold_font
        c_prof.number_format = "#,##0"

        c_warn = ws4.cell(row=idx, column=12, value=f'=IF(J{idx}>35%, "⚠️ Cost Cao (>35%) - Cần Tối Ưu", IF(J{idx}<25%, "💎 Lãi Rất Dày (<25%)", "✅ Chuẩn Ngành (25-35%)"))')
        c_warn.alignment = align_left

        for c in range(1, 13):
            ws4.cell(row=idx, column=c).border = box_border

    avg_row = len(drinks) + 5
    ws4.cell(row=avg_row, column=2, value=f"TRUNG BÌNH TOÀN BỘ MENU ({len(drinks)} MÓN)").font = bold_font
    ws4.cell(row=avg_row, column=4, value=f"=AVERAGE(D5:D{avg_row-1})").number_format = "#,##0"
    ws4.cell(row=avg_row, column=5, value=f"=AVERAGE(E5:E{avg_row-1})").number_format = "#,##0"
    ws4.cell(row=avg_row, column=7, value=f"=AVERAGE(G5:G{avg_row-1})").number_format = "#,##0"
    ws4.cell(row=avg_row, column=8, value=f"=AVERAGE(H5:H{avg_row-1})").number_format = "#,##0"
    ws4.cell(row=avg_row, column=9, value=f"=AVERAGE(I5:I{avg_row-1})").number_format = "#,##0"
    ws4.cell(row=avg_row, column=10, value=f"=AVERAGE(J5:J{avg_row-1})").number_format = "0.0%"
    ws4.cell(row=avg_row, column=11, value=f"=AVERAGE(K5:K{avg_row-1})").number_format = "#,##0"

    for c in range(1, 13):
        cell = ws4.cell(row=avg_row, column=c)
        cell.font = bold_font
        cell.fill = highlight_fill
        cell.border = total_border

    # ----------------------------------------------------
    # SHEET 5: DỰ TOÁN ĐIỂM HÒA VỐN & DÒNG TIỀN
    # ----------------------------------------------------
    ws5 = wb.create_sheet(title="5. Điểm Hòa Vốn & Dòng Tiền")
    ws5.views.sheetView[0].showGridLines = True
    ws5["A1"] = "MÔ HÌNH DỰ TOÁN ĐIỂM HÒA VỐN & KẾ HOẠCH DÒNG TIỀN — 23 COFFEE"
    ws5["A1"].font = title_font
    ws5["A2"] = "Tính toán chuẩn xác: Mỗi ngày quán phải bán tối thiểu bao nhiêu ly để không lỗ và đạt lợi nhuận kỳ vọng."
    ws5["A2"].font = subtitle_font

    ws5["B4"] = "1. DỰ TOÁN CHI PHÍ CỐ ĐỊNH HÀNG THÁNG (FIXED OPERATING COSTS)"
    ws5["B4"].font = bold_font

    headers5_1 = ["STT", "Hạng Mục Chi Phí Cố Định", "Chi Phí Hàng Tháng (VNĐ)", "Tỷ Lệ / Tổng Chi Phí", "Ghi Chú Vận Hành"]
    for col_idx, h in enumerate(headers5_1, 2):
        cell = ws5.cell(row=5, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = box_border

    tot_fixed_rows = len(fixed_costs)
    for idx, item in enumerate(fixed_costs, 6):
        stt = item.get("id", idx-5)
        name = item.get("name", "")
        cost = item.get("cost", 0)
        note = item.get("note", "")

        ws5.cell(row=idx, column=2, value=stt).alignment = align_center
        ws5.cell(row=idx, column=3, value=name).alignment = align_left

        c_c = ws5.cell(row=idx, column=4, value=cost)
        c_c.alignment = align_right
        c_c.number_format = "#,##0"

        c_pct = ws5.cell(row=idx, column=5, value=f"=IF($D${tot_fixed_rows+6}>0, D{idx}/$D${tot_fixed_rows+6}, 0)")
        c_pct.alignment = align_right
        c_pct.number_format = "0.0%"

        ws5.cell(row=idx, column=6, value=note).alignment = align_left
        for c in range(2, 7):
            ws5.cell(row=idx, column=c).border = box_border

    sum_fixed_r = tot_fixed_rows + 6
    ws5.cell(row=sum_fixed_r, column=3, value="TỔNG CHI PHÍ CỐ ĐỊNH / THÁNG").font = bold_font
    c_tot_fixed = ws5.cell(row=sum_fixed_r, column=4, value=f"=SUM(D6:D{sum_fixed_r-1})")
    c_tot_fixed.font = bold_font
    c_tot_fixed.alignment = align_right
    c_tot_fixed.number_format = "#,##0"
    ws5.cell(row=sum_fixed_r, column=5, value=f"=SUM(E6:E{sum_fixed_r-1})").number_format = "0.0%"
    ws5.cell(row=sum_fixed_r, column=6, value="Mức chi phí tối thiểu bắt buộc phải chi hàng tháng").font = italic_font
    for c in range(2, 7):
        ws5.cell(row=sum_fixed_r, column=c).border = total_border
        ws5.cell(row=sum_fixed_r, column=c).fill = highlight_fill

    be_start_r = sum_fixed_r + 2
    ws5.cell(row=be_start_r, column=2, value="2. CÔNG THỨC TÍNH ĐIỂM HÒA VỐN THỰC CHIẾN (BREAK-EVEN POINT)").font = bold_font
    
    be_hdr_r = be_start_r + 1
    headers5_2 = ["Chỉ Số Tài Chính", "Giá Trị Tính Toán", "Đơn Vị", "Giải Thích Ý Nghĩa & Ứng Dụng", "Hành Động Của Chủ Quán"]
    for col_idx, h in enumerate(headers5_2, 2):
        cell = ws5.cell(row=be_hdr_r, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = box_border

    r_aov = be_hdr_r + 1
    r_cm = be_hdr_r + 2
    r_bemonth = be_hdr_r + 3
    r_beday = be_hdr_r + 4
    r_revday = be_hdr_r + 5
    r_revmonth = be_hdr_r + 6
    r_target = be_hdr_r + 7
    r_targetday = be_hdr_r + 8

    be_items = [
        (r_aov, "Giá bán trung bình 1 ly (AOV)", f"='4. Bảng Cost & Định Giá Menu'!I{avg_row}", "VNĐ/ly", "Giá trung bình khách chi tiêu tại quán"),
        (r_cm, "Lãi gộp trung bình trên 1 ly (CM)", f"='4. Bảng Cost & Định Giá Menu'!K{avg_row}", "VNĐ/ly", "Số tiền lời trên mỗi ly để bù tiền mặt bằng & nhân sự"),
        (r_bemonth, "Số ly hòa vốn / tháng (Break-even Monthly)", f"=ROUNDUP(D{sum_fixed_r}/C{r_cm}, 0)", "Ly/tháng", "Tổng số ly phải bán trong tháng để đạt điểm 0 đồng lãi/lỗ"),
        (r_beday, "SỐ LY HÒA VỐN / NGÀY (Break-even Daily)", f"=ROUNDUP(C{r_bemonth}/30, 0)", "Ly/ngày", "MỤC TIÊU BẮT BUỘC: Mỗi ngày mở mắt ra phải bán đủ số này!"),
        (r_revday, "Doanh thu hòa vốn / ngày", f"=C{r_beday}*C{r_aov}", "VNĐ/ngày", "Doanh thu tối thiểu cần quẹt máy mỗi ngày"),
        (r_revmonth, "Doanh thu hòa vốn / tháng", f"=C{r_bemonth}*C{r_aov}", "VNĐ/tháng", "Mức doanh thu chạm ngưỡng an toàn"),
        (r_target, "Mục tiêu Lợi Nhuận Ròng kỳ vọng", target_profit, "VNĐ/tháng", "Tiền chủ quán muốn bỏ túi mang về sau khi trừ hết chi phí"),
        (r_targetday, "SỐ LY CẦN BÁN ĐỂ ĐẠT LÃI KỲ VỌNG / NGÀY", f"=ROUNDUP((D{sum_fixed_r}+C{r_target})/(C{r_cm}*30), 0)", "Ly/ngày", "Mục tiêu kinh doanh thực tế để quán phát triển bền vững")
    ]

    for r_num, metric, val, unit, expl in be_items:
        ws5.cell(row=r_num, column=2, value=metric).alignment = align_left
        c_v = ws5.cell(row=r_num, column=3, value=val)
        c_v.alignment = align_right
        c_v.font = bold_font
        c_v.number_format = "#,##0"
        if r_num in [r_beday, r_targetday]:
            c_v.fill = user_input_fill
            ws5.cell(row=r_num, column=2).font = bold_font

        ws5.cell(row=r_num, column=4, value=unit).alignment = align_center
        ws5.cell(row=r_num, column=5, value=expl).alignment = align_left
        for c in range(2, 6):
            ws5.cell(row=r_num, column=c).border = box_border

    # Adjust column widths
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2]:
                    continue
                val_str = str(cell.value or "")
                if val_str.startswith("="):
                    val_str = "123,456"
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_filename)
    return output_filename
