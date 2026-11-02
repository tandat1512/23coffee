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
DEFAULT_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbz3XF_9TVUAA5XPXnFq9nYakh0qm3Kuc78vl50nIYxodfQI0iMDhSJC78hdy4aU8yAFJg/exec"

# Thread-safe global sync state for live user notifications
_sync_lock = threading.Lock()
_sync_state = {
    "is_syncing": False,
    "last_status": "Đã kết nối Google Sheets",
    "last_synced": None,
    "has_error": False,
    "error_message": "",
    "timestamp": None
}

def get_sync_state():
    with _sync_lock:
        return dict(_sync_state)

def update_sync_state(is_syncing=None, last_status=None, last_synced=None, has_error=None, error_message=None):
    with _sync_lock:
        now_time = datetime.datetime.now().strftime("%H:%M:%S")
        if is_syncing is not None:
            _sync_state["is_syncing"] = is_syncing
        if last_status is not None:
            _sync_state["last_status"] = last_status
        if last_synced is not None:
            _sync_state["last_synced"] = last_synced
        if has_error is not None:
            _sync_state["has_error"] = has_error
        if error_message is not None:
            _sync_state["error_message"] = error_message
        _sync_state["timestamp"] = now_time
    return dict(_sync_state)

def load_google_config():
    defaults = {
        "sheet_url": DEFAULT_SHEET_URL,
        "sheet_id": DEFAULT_SHEET_ID,
        "webhook_url": DEFAULT_WEBHOOK_URL,
        "auto_sync": True,
        "last_synced": None,
        "last_status": "Sẵn sàng",
        "is_configured": True
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                defaults.update(saved)
        except Exception:
            pass
            
    env_webhook = os.environ.get("GOOGLE_SHEET_WEBHOOK", "").strip()
    if env_webhook:
        defaults["webhook_url"] = env_webhook
    elif not defaults.get("webhook_url"):
        defaults["webhook_url"] = DEFAULT_WEBHOOK_URL
        
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
    Push data payload to Google Sheet via Google Apps Script Webhook.
    Updates thread-safe sync state so user sees live feedback.
    """
    config = load_google_config()
    webhook_url = config.get("webhook_url", "").strip() or DEFAULT_WEBHOOK_URL
    
    update_sync_state(
        is_syncing=True,
        last_status="Đang tự động lưu lên Google Sheets...",
        has_error=False,
        error_message=""
    )
    
    now_full = datetime.datetime.now().strftime("%H:%M:%S %d/%m/%Y")
    now_time = datetime.datetime.now().strftime("%H:%M:%S")
    payload_bytes = json.dumps(data, ensure_ascii=False).encode("utf-8")

    # Method 1: requests library
    if HAS_REQUESTS:
        try:
            headers = {"Content-Type": "application/json; charset=utf-8"}
            response = requests.post(webhook_url, data=payload_bytes, headers=headers, timeout=30, allow_redirects=True)
            if response.status_code == 200:
                try:
                    res_json = response.json()
                    if res_json.get("status") == "success":
                        save_google_config({
                            "last_synced": now_full,
                            "last_status": f"Đã lưu thành công ({now_time})"
                        })
                        update_sync_state(
                            is_syncing=False,
                            last_status=f"Đã tự động lưu thành công ({now_time})",
                            last_synced=now_time,
                            has_error=False,
                            error_message=""
                        )
                        return {
                            "success": True,
                            "message": "Đã tự động lưu thành công lên Google Sheet!",
                            "timestamp": now_time,
                            "google_response": res_json
                        }
                    else:
                        err_msg = res_json.get("message", "Lỗi từ Google Apps Script")
                        update_sync_state(
                            is_syncing=False,
                            last_status="Hiện tại chưa lưu được lên Google Sheets",
                            has_error=True,
                            error_message=err_msg
                        )
                        return {"success": False, "error": err_msg}
                except Exception:
                    update_sync_state(
                        is_syncing=False,
                        last_status=f"Đã tự động lưu thành công ({now_time})",
                        last_synced=now_time,
                        has_error=False
                    )
                    return {"success": True, "message": "Đã gửi dữ liệu sang Google Sheets!", "timestamp": now_time}
            elif response.status_code in [401, 403]:
                err_auth = "Lỗi quyền truy cập (Unauthorized): Mục 'Ai có quyền truy cập' cần chọn là 'Bất kỳ ai' (Anyone)."
                update_sync_state(
                    is_syncing=False,
                    last_status="Hiện tại chưa lưu được lên Google Sheets (Lỗi quyền 401)",
                    has_error=True,
                    error_message=err_auth
                )
                return {"success": False, "error": err_auth}
        except Exception as e:
            print(f"Requests error, trying urllib fallback: {e}")

    # Method 2: Standard library urllib.request
    try:
        req = urllib.request.Request(
            webhook_url,
            data=payload_bytes,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST"
        )
        
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp_data = resp.read().decode("utf-8")
            try:
                res_json = json.loads(resp_data)
                if res_json.get("status") == "success":
                    save_google_config({
                        "last_synced": now_full,
                        "last_status": f"Đã lưu thành công ({now_time})"
                    })
                    update_sync_state(
                        is_syncing=False,
                        last_status=f"Đã tự động lưu thành công ({now_time})",
                        last_synced=now_time,
                        has_error=False,
                        error_message=""
                    )
                    return {
                        "success": True,
                        "message": "Đã tự động lưu thành công lên Google Sheet!",
                        "timestamp": now_time,
                        "google_response": res_json
                    }
                else:
                    err_msg = res_json.get("message", "Lỗi từ Apps Script")
                    update_sync_state(
                        is_syncing=False,
                        last_status="Hiện tại chưa lưu được lên Google Sheets",
                        has_error=True,
                        error_message=err_msg
                    )
                    return {"success": False, "error": err_msg}
            except Exception:
                update_sync_state(
                    is_syncing=False,
                    last_status=f"Đã tự động lưu thành công ({now_time})",
                    last_synced=now_time,
                    has_error=False
                )
                return {"success": True, "message": "Đã gửi dữ liệu sang Google Sheets!", "timestamp": now_time}
                
    except urllib.error.HTTPError as e:
        if e.code in [401, 403]:
            friendly_err = "Lỗi quyền truy cập (Unauthorized): Mục 'Ai có quyền truy cập' cần chọn là 'Bất kỳ ai' (Anyone)."
        else:
            friendly_err = f"Lỗi máy chủ Google HTTP {e.code}: {e.reason}"
        update_sync_state(
            is_syncing=False,
            last_status="Hiện tại chưa lưu được lên Google Sheets",
            has_error=True,
            error_message=friendly_err
        )
        return {"success": False, "error": friendly_err}
    except urllib.error.URLError as e:
        err_msg = f"Không thể kết nối đến Google Sheets: {e.reason}"
        update_sync_state(
            is_syncing=False,
            last_status="Hiện tại chưa lưu được lên Google Sheets (Mất kết nối)",
            has_error=True,
            error_message=err_msg
        )
        return {"success": False, "error": err_msg}
    except Exception as e:
        err_msg = str(e)
        update_sync_state(
            is_syncing=False,
            last_status="Hiện tại chưa lưu được lên Google Sheets",
            has_error=True,
            error_message=err_msg
        )
        return {"success": False, "error": err_msg}

def pull_from_google_sheets(timeout=25):
    """
    Pull data from Google Sheet via Apps Script doGet.
    Returns complete dataset (materials, packaging, drinks, fixed_costs, target_profit).
    """
    config = load_google_config()
    webhook_url = config.get("webhook_url", "").strip() or DEFAULT_WEBHOOK_URL
    
    try:
        req = urllib.request.Request(
            webhook_url, 
            headers={"User-Agent": "Mozilla/5.0 (23CoffeeSync/1.0)"},
            method="GET"
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            resp_data = resp.read().decode("utf-8")
            res_json = json.loads(resp_data)
            if res_json.get("status") == "success":
                return {
                    "success": True, 
                    "data": res_json.get("data", {}),
                    "source": res_json.get("source", "unknown"),
                    "timestamp": res_json.get("timestamp")
                }
            else:
                return {"success": False, "error": res_json.get("message", "Lỗi không xác định từ Google Sheets")}
    except Exception as e:
        return {"success": False, "error": str(e)}

def trigger_async_sync(data):
    """
    Background worker to push updates to Google Sheets asynchronously
    so API calls return immediately without blocking UI.
    """
    config = load_google_config()
    if config.get("auto_sync", True):
        update_sync_state(
            is_syncing=True,
            last_status="Đang tự động lưu lên Google Sheets...",
            has_error=False,
            error_message=""
        )
        t = threading.Thread(target=push_to_google_sheets, args=(data,), daemon=True)
        t.start()
        return True
    return False
