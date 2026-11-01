import os
import json
import threading
import datetime
import urllib.request
import urllib.error
import urllib.parse

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

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
    Uses requests if available, or native urllib.request.
    """
    config = load_google_config()
    webhook_url = config.get("webhook_url", "").strip()
    
    if not webhook_url:
        return {
            "success": False, 
            "error": "Chưa cấu hình URL Webhook Google Sheets. Vui lòng bấm Google Sheets trên thanh menu để dán URL."
        }
        
    now_str = datetime.datetime.now().strftime("%H:%M:%S %d/%m/%Y")
    payload_bytes = json.dumps(data, ensure_ascii=False).encode("utf-8")

    # Method 1: requests library
    if HAS_REQUESTS:
        try:
            headers = {"Content-Type": "application/json; charset=utf-8"}
            response = requests.post(webhook_url, data=payload_bytes, headers=headers, timeout=25, allow_redirects=True)
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
                    save_google_config({"last_synced": now_str, "last_status": "Thành công (" + now_str + ")"})
                    return {"success": True, "message": "Đã gửi dữ liệu sang Google Sheets!", "timestamp": now_str}
        except Exception as e:
            print(f"Requests post error, trying urllib fallback: {e}")

    # Method 2: Standard library urllib.request
    try:
        req = urllib.request.Request(
            webhook_url,
            data=payload_bytes,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST"
        )
        
        with urllib.request.urlopen(req, timeout=25) as resp:
            resp_data = resp.read().decode("utf-8")
            try:
                res_json = json.loads(resp_data)
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
                save_google_config({"last_synced": now_str, "last_status": "Thành công (" + now_str + ")"})
                return {"success": True, "message": "Đã gửi dữ liệu sang Google Sheets!", "timestamp": now_str}
                
    except urllib.error.URLError as e:
        save_google_config({"last_status": f"Lỗi kết nối: {e.reason}"})
        return {"success": False, "error": f"Không thể kết nối đến Google Sheets: {e.reason}"}
    except Exception as e:
        save_google_config({"last_status": f"Lỗi: {str(e)}"})
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
        req = urllib.request.Request(webhook_url, method="GET")
        with urllib.request.urlopen(req, timeout=25) as resp:
            resp_data = resp.read().decode("utf-8")
            res_json = json.loads(resp_data)
            if res_json.get("status") == "success":
                return {"success": True, "data": res_json.get("data", {})}
            else:
                return {"success": False, "error": res_json.get("message", "Lỗi không xác định")}
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
