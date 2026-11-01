import os
import json
import threading
import datetime
import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "google_config.json")
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1uGRxJEPZljoVKlbBXRatjPiIpk2NBnXldwWSfVYVy_Y/edit?usp=sharing"
DEFAULT_SHEET_ID = "1uGRxJEPZljoVKlbBXRatjPiIpk2NBnXldwWSfVYVy_Y"

def load_google_config():
    defaults = {
        "sheet_url": DEFAULT_SHEET_URL,
        "sheet_id": DEFAULT_SHEET_ID,
        "webhook_url": os.environ.get("GOOGLE_SHEET_WEBHOOK", ""),
        "auto_sync": True,
        "last_synced": None,
        "last_status": "Chưa kết nối Webhook",
        "is_configured": False
    }
    
    # Environment variable has priority
    env_webhook = os.environ.get("GOOGLE_SHEET_WEBHOOK", "").strip()
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                defaults.update(saved)
        except Exception:
            pass
            
    if env_webhook:
        defaults["webhook_url"] = env_webhook
        
    defaults["is_configured"] = bool(defaults.get("webhook_url", "").strip())
    return defaults

def save_google_config(config_data):
    current = load_google_config()
    current.update(config_data)
    current["is_configured"] = bool(current.get("webhook_url", "").strip())
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(current, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error saving google_config.json: {e}")
    return current

def push_to_google_sheets(data):
    """
    Synchronously push data payload to Google Sheet via Google Apps Script Webhook.
    """
    config = load_google_config()
    webhook_url = config.get("webhook_url", "").strip()
    
    if not webhook_url:
        return {
            "success": False, 
            "error": "Chưa cấu hình URL Webhook Google Sheets. Vui lòng bấm Cấu hình Google Sheets để nhập URL."
        }
        
    try:
        headers = {"Content-Type": "application/json"}
        # Google Apps Script redirects with 302, requests follows automatically
        response = requests.post(webhook_url, json=data, headers=headers, timeout=25, allow_redirects=True)
        
        now_str = datetime.datetime.now().strftime("%H:%M:%S %d/%m/%Y")
        
        if response.status_code == 200:
            try:
                res_json = response.json()
                if res_json.get("status") == "success":
                    save_google_config({
                        "last_synced": now_str,
                        "last_status": "Thành công (" + now_str + ")"
                    })
                    return {
                        "success": True,
                        "message": "Đã đồng bộ thành công vào Google Sheet!",
                        "timestamp": now_str,
                        "google_response": res_json
                    }
                else:
                    err_msg = res_json.get("message", "Lỗi từ Apps Script")
                    save_google_config({"last_status": f"Lỗi: {err_msg}"})
                    return {"success": False, "error": err_msg}
            except Exception:
                # Text response fallback
                save_google_config({
                    "last_synced": now_str,
                    "last_status": "Thành công (" + now_str + ")"
                })
                return {
                    "success": True,
                    "message": "Đã gửi dữ liệu sang Google Sheets!",
                    "timestamp": now_str
                }
        else:
            err_msg = f"HTTP {response.status_code}: {response.text[:200]}"
            save_google_config({"last_status": f"Lỗi HTTP: {err_msg}"})
            return {"success": False, "error": err_msg}
            
    except requests.exceptions.Timeout:
        save_google_config({"last_status": "Hết thời gian chờ kết nối (Timeout)"})
        return {"success": False, "error": "Kết nối tới Google quá thời gian chờ (Timeout). Vui lòng thử lại."}
    except Exception as e:
        save_google_config({"last_status": f"Lỗi kết nối: {str(e)}"})
        return {"success": False, "error": str(e)}

def pull_from_google_sheets():
    """
    Pull data from Google Sheet via Apps Script doGet.
    """
    config = load_google_config()
    webhook_url = config.get("webhook_url", "").strip()
    
    if not webhook_url:
        return {"success": False, "error": "Chưa cấu hình URL Webhook Google Sheets."}
        
    try:
        response = requests.get(webhook_url, timeout=25, allow_redirects=True)
        if response.status_code == 200:
            res_json = response.json()
            if res_json.get("status") == "success":
                return {"success": True, "data": res_json.get("data", {})}
            else:
                return {"success": False, "error": res_json.get("message", "Lỗi không xác định")}
        else:
            return {"success": False, "error": f"HTTP {response.status_code}"}
    except Exception as e:
        return {"success": False, "error": str(e)}

def trigger_async_sync(data):
    """
    Background worker to push updates to Google Sheets asynchronously
    so API calls return immediately without blocking UI.
    """
    config = load_google_config()
    if config.get("auto_sync", True) and config.get("is_configured", False):
        t = threading.Thread(target=push_to_google_sheets, args=(data,), daemon=True)
        t.start()
        return True
    return False
