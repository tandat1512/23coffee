import os
import sys
import json
import math
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
from werkzeug.security import generate_password_hash, check_password_hash
from excel_service import generate_excel_from_data
from google_sheets_service import (
    load_google_config, 
    save_google_config, 
    push_to_google_sheets, 
    pull_from_google_sheets, 
    trigger_async_sync,
    get_sync_state
)
from db_service import save_snapshot_to_db, load_data_from_db, init_db

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data_cost.json")
USERS_FILE = os.path.join(BASE_DIR, "users.json")
EXCEL_FILE = os.path.join(BASE_DIR, "Bang_Cost_Va_Chien_Luoc_Dinh_Gia_23_Coffee.xlsx")
FALLBACK_EXCEL = os.path.join(BASE_DIR, "Bang_Cost_Va_Chien_Luoc_Dinh_Gia_23_Coffee_Cap_Nhat.xlsx")

app = Flask(__name__, template_folder="templates")
app.secret_key = os.environ.get("SECRET_KEY", "23coffee-secure-secret-key-2026-very-strong")

# ----------------------------------------------------
# AUTHENTICATION MANAGEMENT (EXACTLY 2 USERS)
# ----------------------------------------------------
DEFAULT_USERS = {
    "admin": {
        "password_hash": generate_password_hash("admin"),
        "name": "Admin (23 COFFEE)",
        "role": "admin"
    },
    "tandat": {
        "password_hash": generate_password_hash("admin"),
        "name": "Tấn Đạt (Chủ quán)",
        "role": "admin"
    }
}

def load_users():
    if not os.path.exists(USERS_FILE):
        save_users(DEFAULT_USERS)
        return DEFAULT_USERS
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return DEFAULT_USERS

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

@app.before_request
def require_login():
    # Public endpoints
    allowed_endpoints = ['login', 'static']
    if request.endpoint in allowed_endpoints:
        return
        
    # Check session
    if 'user' not in session:
        if request.path.startswith('/api/'):
            return jsonify({"error": "Unauthorized"}), 401
        return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        users = load_users()
        
        # Check environment variable overrides if deployed to cloud
        env_u1 = os.environ.get("USER1_USER", "tandat").lower()
        env_p1 = os.environ.get("USER1_PASS", "")
        env_u2 = os.environ.get("USER2_USER", "partner").lower()
        env_p2 = os.environ.get("USER2_PASS", "")

        authenticated = False
        user_info = None

        if (username == "admin" and password == "admin") or (username == "tandat" and password == "admin"):
            authenticated = True
            user_info = {"name": "Admin (23 COFFEE)", "role": "admin"}
        elif env_p1 and username == env_u1 and password == env_p1:
            authenticated = True
            user_info = {"name": "Tấn Đạt (Chủ quán)", "role": "admin"}
        elif env_p2 and username == env_u2 and password == env_p2:
            authenticated = True
            user_info = {"name": "Cổ đông / Quản lý", "role": "manager"}
        elif username in users and check_password_hash(users[username]["password_hash"], password):
            authenticated = True
            user_info = users[username]

        if authenticated:
            session['user'] = username
            session['name'] = user_info.get('name', username)
            session['role'] = user_info.get('role', 'user')
            return redirect(url_for('index'))
        else:
            return render_template('login.html', error="Tên đăng nhập hoặc mật khẩu không chính xác! Vui lòng thử lại.")
            
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/api/change_password', methods=['POST'])
def change_password():
    current_user = session.get('user')
    if not current_user:
        return jsonify({"error": "Unauthorized"}), 401
        
    req = request.json or {}
    new_password = req.get('new_password', '').strip()
    if len(new_password) < 6:
        return jsonify({"error": "Mật khẩu phải từ 6 ký tự trở lên"}), 400
        
    users = load_users()
    if current_user in users:
        users[current_user]["password_hash"] = generate_password_hash(new_password)
        save_users(users)
        return jsonify({"success": True})
    return jsonify({"error": "Không tìm thấy tài khoản"}), 404

# ----------------------------------------------------
# DATA SERVICES & CALCULATIONS
# ----------------------------------------------------
def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading {DATA_FILE}: {e}")
            
    # Fallback to SQLite DB
    db_data = load_data_from_db()
    if db_data:
        return db_data
        
    return {"materials": [], "packaging": [], "drinks": [], "fixed_costs": [], "target_profit": 15000000}

def prepare_google_payload(data):
    metrics = calculate_metrics(data)
    return {
        "materials": data.get("materials", []),
        "packaging": data.get("packaging", []),
        "drinks": metrics.get("drinks", []),
        "fixed_costs": data.get("fixed_costs", []),
        "target_profit": data.get("target_profit", 15000000)
    }

def save_data(data):
    # 1. Save to JSON
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        
    # 2. Persist to SQLite Database
    save_snapshot_to_db(data)
    
    # 3. Asynchronously push to Google Sheets (if configured) with calculated metrics
    payload = prepare_google_payload(data)
    trigger_async_sync(payload)

def calculate_metrics(data):
    materials_map = {m["name"].lower().strip(): m for m in data.get("materials", [])}
    pkg_map = {p["id"]: p for p in data.get("packaging", [])}
    
    processed_drinks = []
    total_cost_list = []
    selling_price_list = []
    profit_list = []
    
    for d in data.get("drinks", []):
        d_copy = dict(d)
        recipe = d.get("recipe", [])
        
        cost_nl = 0.0
        recipe_detailed = []
        for ing in recipe:
            ing_name = ing.get("name", "").strip()
            qty = float(ing.get("qty", 0))
            mat = materials_map.get(ing_name.lower())
            unit_cost = mat["unit_cost"] if mat else 0.0
            unit = mat["unit_base"] if mat else ing.get("unit", "g")
            sub_cost = round(qty * unit_cost, 0)
            cost_nl += sub_cost
            recipe_detailed.append({
                "name": ing_name,
                "qty": qty,
                "unit": unit,
                "unit_cost": unit_cost,
                "cost": sub_cost
            })
            
        pkg_id = d.get("pkg_id", "BB01")
        pkg = pkg_map.get(pkg_id, {"total": 1100})
        cost_bb = float(pkg.get("total", 1100))
        waste_pct = float(d.get("waste_pct", 0.05))
        
        total_cost = round(cost_nl * (1 + waste_pct) + cost_bb, 0)
        suggested_price = math.ceil((total_cost / 0.3) / 1000) * 1000 if total_cost > 0 else 0
        selling_price = float(d.get("selling_price", 0))
        
        fc_pct = (total_cost / selling_price * 100) if selling_price > 0 else 0
        profit = selling_price - total_cost
        
        if fc_pct > 35:
            warning = "⚠️ Cost Cao (>35%) - Cần Tối Ưu"
            badge_class = "badge-danger"
        elif fc_pct < 25:
            warning = "💎 Lãi Rất Dày (<25%)"
            badge_class = "badge-purple"
        else:
            warning = "✅ Chuẩn Ngành (25-35%)"
            badge_class = "badge-success"
            
        d_copy["cost_nl"] = cost_nl
        d_copy["cost_bb"] = cost_bb
        d_copy["total_cost"] = total_cost
        d_copy["suggested_price"] = suggested_price
        d_copy["food_cost_pct"] = round(fc_pct, 1)
        d_copy["profit"] = profit
        d_copy["warning"] = warning
        d_copy["badge_class"] = badge_class
        d_copy["recipe_detailed"] = recipe_detailed
        
        processed_drinks.append(d_copy)
        total_cost_list.append(total_cost)
        selling_price_list.append(selling_price)
        profit_list.append(profit)
        
    num_drinks = len(processed_drinks)
    avg_price = sum(selling_price_list) / num_drinks if num_drinks else 0
    avg_cost = sum(total_cost_list) / num_drinks if num_drinks else 0
    avg_profit = sum(profit_list) / num_drinks if num_drinks else 0
    avg_fc = (avg_cost / avg_price * 100) if avg_price else 0
    
    fixed_costs = data.get("fixed_costs", [])
    total_fixed_cost = sum(float(f.get("cost", 0)) for f in fixed_costs)
    target_profit = float(data.get("target_profit", 15000000))
    
    be_monthly_cups = math.ceil(total_fixed_cost / avg_profit) if avg_profit > 0 else 0
    be_daily_cups = math.ceil(be_monthly_cups / 30) if be_monthly_cups > 0 else 0
    be_daily_revenue = be_daily_cups * avg_price
    target_daily_cups = math.ceil((total_fixed_cost + target_profit) / (avg_profit * 30)) if avg_profit > 0 else 0
    
    summary = {
        "num_drinks": num_drinks,
        "num_materials": len(data.get("materials", [])),
        "avg_price": round(avg_price, 0),
        "avg_cost": round(avg_cost, 0),
        "avg_profit": round(avg_profit, 0),
        "avg_food_cost_pct": round(avg_fc, 1),
        "total_fixed_cost": total_fixed_cost,
        "be_monthly_cups": be_monthly_cups,
        "be_daily_cups": be_daily_cups,
        "be_daily_revenue": round(be_daily_revenue, 0),
        "target_profit": target_profit,
        "target_daily_cups": target_daily_cups
    }
    
    return {
        "materials": data.get("materials", []),
        "packaging": data.get("packaging", []),
        "drinks": processed_drinks,
        "fixed_costs": fixed_costs,
        "summary": summary
    }

def sync_excel_file():
    data = load_data()
    try:
        generate_excel_from_data(data, EXCEL_FILE)
        return {"success": True, "file": EXCEL_FILE, "is_main": True}
    except PermissionError:
        generate_excel_from_data(data, FALLBACK_EXCEL)
        return {
            "success": True, 
            "file": FALLBACK_EXCEL, 
            "is_main": False,
            "message": "File Excel chính đang mở trên máy nên hệ thống đã lưu vào bản cập nhật: Bang_Cost_Va_Chien_Luoc_Dinh_Gia_23_Coffee_Cap_Nhat.xlsx"
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

# ----------------------------------------------------
# MAIN ROUTES & REST API
# ----------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/data", methods=["GET"])
def get_all_data():
    data = load_data()
    return jsonify(calculate_metrics(data))

@app.route("/api/materials", methods=["POST"])
def add_material():
    data = load_data()
    req = request.json or {}
    name = req.get("name", "").strip()
    if not name:
        return jsonify({"error": "Tên nguyên liệu không được để trống"}), 400
        
    for m in data.get("materials", []):
        if m["name"].lower().strip() == name.lower():
            return jsonify({"error": f"Nguyên liệu '{name}' đã tồn tại"}), 400
            
    qty = float(req.get("qty", 1000))
    price = float(req.get("price", 0))
    unit_cost = round(price / qty, 2) if qty > 0 else 0
    
    new_id = f"NL{len(data.get('materials', [])) + 1:02d}"
    new_item = {
        "id": new_id,
        "name": name,
        "pack": req.get("pack", "Gói 1kg"),
        "unit_buy": req.get("unit_buy", "kg"),
        "qty": qty,
        "price": price,
        "unit_cost": unit_cost,
        "unit_base": req.get("unit_base", "g"),
        "note": req.get("note", "")
    }
    
    data["materials"].append(new_item)
    save_data(data)
    sync_excel_file()
    return jsonify({"success": True, "material": new_item})

@app.route("/api/materials/<mat_id>", methods=["PUT"])
def update_material(mat_id):
    data = load_data()
    req = request.json or {}
    found = False
    for m in data.get("materials", []):
        if m["id"] == mat_id:
            m["name"] = req.get("name", m["name"]).strip()
            m["pack"] = req.get("pack", m["pack"]).strip()
            m["unit_buy"] = req.get("unit_buy", m["unit_buy"]).strip()
            m["qty"] = float(req.get("qty", m["qty"]))
            m["price"] = float(req.get("price", m["price"]))
            m["unit_base"] = req.get("unit_base", m["unit_base"]).strip()
            m["unit_cost"] = round(m["price"] / m["qty"], 2) if m["qty"] > 0 else 0
            m["note"] = req.get("note", m["note"]).strip()
            found = True
            break
            
    if not found:
        return jsonify({"error": "Không tìm thấy nguyên liệu"}), 404
        
    save_data(data)
    sync_excel_file()
    return jsonify({"success": True})

@app.route("/api/materials/<mat_id>", methods=["DELETE"])
def delete_material(mat_id):
    data = load_data()
    mat_to_delete = next((m for m in data.get("materials", []) if m["id"] == mat_id), None)
    if not mat_to_delete:
        return jsonify({"error": "Không tìm thấy nguyên liệu"}), 404
        
    mat_name = mat_to_delete["name"].lower().strip()
    used_in = []
    for d in data.get("drinks", []):
        for ing in d.get("recipe", []):
            if ing.get("name", "").lower().strip() == mat_name:
                used_in.append(d["name"])
                break
                
    if used_in:
        return jsonify({
            "error": f"Không thể xóa nguyên liệu '{mat_to_delete['name']}' vì đang được dùng trong món: {', '.join(used_in[:3])}..."
        }), 400
        
    data["materials"] = [m for m in data["materials"] if m["id"] != mat_id]
    save_data(data)
    sync_excel_file()
    return jsonify({"success": True})

@app.route("/api/drinks", methods=["POST"])
def add_drink():
    data = load_data()
    req = request.json or {}
    name = req.get("name", "").strip()
    if not name:
        return jsonify({"error": "Tên thức uống không được để trống"}), 400
        
    for d in data.get("drinks", []):
        if d["name"].lower().strip() == name.lower():
            return jsonify({"error": f"Món '{name}' đã có trong menu"}), 400
            
    new_id = f"D{len(data.get('drinks', [])) + 1:02d}"
    new_drink = {
        "id": new_id,
        "name": name,
        "category": req.get("category", "Món Khác").strip(),
        "pkg_id": req.get("pkg_id", "BB01"),
        "waste_pct": float(req.get("waste_pct", 0.05)),
        "selling_price": float(req.get("selling_price", 30000)),
        "recipe": req.get("recipe", [])
    }
    
    data["drinks"].append(new_drink)
    save_data(data)
    sync_excel_file()
    return jsonify({"success": True, "drink": new_drink})

@app.route("/api/drinks/<drink_id>", methods=["PUT"])
def update_drink(drink_id):
    data = load_data()
    req = request.json or {}
    found = False
    for d in data.get("drinks", []):
        if d["id"] == drink_id:
            d["name"] = req.get("name", d["name"]).strip()
            d["category"] = req.get("category", d["category"]).strip()
            d["pkg_id"] = req.get("pkg_id", d["pkg_id"])
            d["waste_pct"] = float(req.get("waste_pct", d["waste_pct"]))
            d["selling_price"] = float(req.get("selling_price", d["selling_price"]))
            d["recipe"] = req.get("recipe", d["recipe"])
            found = True
            break
            
    if not found:
        return jsonify({"error": "Không tìm thấy món"}), 404
        
    save_data(data)
    sync_excel_file()
    return jsonify({"success": True})

@app.route("/api/drinks/<drink_id>", methods=["DELETE"])
def delete_drink(drink_id):
    data = load_data()
    orig_len = len(data.get("drinks", []))
    data["drinks"] = [d for d in data.get("drinks", []) if d["id"] != drink_id]
    if len(data["drinks"]) == orig_len:
        return jsonify({"error": "Không tìm thấy món"}), 404
        
    for idx, d in enumerate(data["drinks"], 1):
        d["id"] = f"D{idx:02d}"
        
    save_data(data)
    sync_excel_file()
    return jsonify({"success": True})

@app.route("/api/fixed_costs", methods=["POST"])
def update_fixed_costs():
    data = load_data()
    req = request.json or {}
    if "fixed_costs" in req:
        data["fixed_costs"] = req["fixed_costs"]
    if "target_profit" in req:
        data["target_profit"] = float(req["target_profit"])
        
    save_data(data)
    sync_excel_file()
    return jsonify({"success": True})

@app.route("/api/sync_excel", methods=["POST"])
def trigger_sync_excel():
    res = sync_excel_file()
    return jsonify(res)

@app.route("/api/download_excel", methods=["GET"])
def download_excel():
    sync_excel_file()
    target_file = EXCEL_FILE if os.path.exists(EXCEL_FILE) else FALLBACK_EXCEL
    return send_file(
        target_file,
        as_attachment=True,
        download_name="Bang_Cost_Va_Chien_Luoc_Dinh_Gia_23_Coffee.xlsx"
    )

# ----------------------------------------------------
# GOOGLE SHEETS SYNCHRONIZATION API
# ----------------------------------------------------
@app.route("/api/google_sync/status", methods=["GET"])
def google_sync_status():
    state = get_sync_state()
    config = load_google_config()
    return jsonify({
        "success": True,
        "state": state,
        "config": config
    })

@app.route("/api/google_sync/config", methods=["GET", "POST"])
def google_sync_config():
    if request.method == "POST":
        req = request.json or {}
        new_config = {}
        if "webhook_url" in req:
            new_config["webhook_url"] = req["webhook_url"].strip()
        if "auto_sync" in req:
            new_config["auto_sync"] = bool(req["auto_sync"])
        if "sheet_url" in req:
            new_config["sheet_url"] = req["sheet_url"].strip()
            
        updated = save_google_config(new_config)
        return jsonify({"success": True, "config": updated})
        
    config = load_google_config()
    return jsonify({"success": True, "config": config})

@app.route("/api/google_sync/push", methods=["POST"])
def trigger_google_push():
    data = load_data()
    payload = prepare_google_payload(data)
    res = push_to_google_sheets(payload)
    return jsonify(res)

@app.route("/api/google_sync/pull", methods=["POST"])
def trigger_google_pull():
    res = pull_from_google_sheets()
    if res.get("success"):
        sheet_data = res.get("data", {})
        # If materials came from sheet, merge or update
        if "materials" in sheet_data and sheet_data["materials"]:
            data = load_data()
            data["materials"] = sheet_data["materials"]
            save_data(data)
            sync_excel_file()
            return jsonify({"success": True, "message": "Đã tải dữ liệu mới từ Google Sheet về hệ thống!"})
    return jsonify(res)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting 23 COFFEE Cost Management App on port {port} ...")
    app.run(host="0.0.0.0", port=port, debug=False)
