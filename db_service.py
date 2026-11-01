import os
import sqlite3
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "coffee23.db")
DATA_FILE = os.path.join(BASE_DIR, "data_cost.json")

def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # Store settings & key-values
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    
    # Store backup snapshots
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS data_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            data_json TEXT
        )
    """)
    
    conn.commit()
    conn.close()

def save_snapshot_to_db(data):
    try:
        init_db()
        conn = get_connection()
        cursor = conn.cursor()
        
        # Save latest full payload as current state
        data_str = json.dumps(data, ensure_ascii=False)
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ("current_cost_data", data_str))
        
        # Keep latest 20 snapshots history for safety
        cursor.execute("INSERT INTO data_snapshots (data_json) VALUES (?)", (data_str,))
        cursor.execute("""
            DELETE FROM data_snapshots 
            WHERE id NOT IN (
                SELECT id FROM data_snapshots ORDER BY id DESC LIMIT 20
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error persisting to SQLite DB: {e}")

def load_data_from_db():
    try:
        if not os.path.exists(DB_FILE):
            return None
        init_db()
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = 'current_cost_data'")
        row = cursor.fetchone()
        conn.close()
        if row and row["value"]:
            return json.loads(row["value"])
    except Exception as e:
        print(f"Error reading from SQLite DB: {e}")
    return None
