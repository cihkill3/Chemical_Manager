import json
import os

import sys

def get_app_root():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CONFIG_FILE = os.path.join(get_app_root(), "config.json")

DEFAULT_CONFIG = {
    "source_file": "",
    "source_sheet": "",
    "header_row": 1,
    "source_headers": [
        "번호", "날짜", "주문자", "회사", "품목명", 
        "수령확인", "CAS 번호", "품번", "용량", "수량", "보관온도"
    ],
    "target_headers": [
        "Order No.", "Order Date", "Ordered By", "Product Name", 
        "Manufacturer", "Package Size", "CAS No.", "Catalog No.", 
        "Room", "Storage Temp.", "Cabinet", "Quantity", "Used", "Status", "Remarks"
    ],
    "mapping": {
        "Order No.": "번호",
        "Order Date": "날짜",
        "Ordered By": "주문자",
        "Product Name": "품목명",
        "Manufacturer": "회사",
        "Package Size": "용량",
        "CAS No.": "CAS 번호",
        "Catalog No.": "품번",
        "Room": "",
        "Storage Temp.": "보관온도",
        "Cabinet": "",
        "Quantity": "수량",
        "Used": "",
        "Status": "",
        "Remarks": ""
    },
    "colors": {
        "done_bg": "#e6f7e6",
        "done_font": "#000000",
        "warn_cas_bg": "#fffbe6",
        "warn_cas_font": "#000000",
        "warn_num_bg": "#fff2e6",
        "warn_num_font": "#000000",
        "warn_both_bg": "#ffe6e6",
        "warn_both_font": "#000000",
        "warn_loc_bg": "#e6f2ff",
        "warn_loc_font": "#000000",
        "normal_bg": "#ffffff",
        "normal_font": "#000000",
        "header_bg": "#d9d9d9",
        "header_font": "#000000",
        "search_failed_bg": "#ffff00",
        "search_failed_font": "#ff0000",
        "manual_input_bg": "#ffff00",
        "manual_input_font": "#ff0000",
        "status_x_bg": "#ffe6e6",
        "status_x_font": "#ff0000"
    },
    "sync_interval_minutes": 0,
    "watch_enabled": False,
    "run_on_startup": False
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            # Merge with defaults for missing keys
            for key, val in DEFAULT_CONFIG.items():
                if key not in data:
                    data[key] = val
                elif isinstance(val, dict):
                    # For nested dictionaries like mapping and colors
                    for sub_key, sub_val in val.items():
                        if sub_key not in data[key]:
                            data[key][sub_key] = sub_val
            return data
    except Exception as e:
        print(f"Error loading config: {e}")
        return DEFAULT_CONFIG.copy()

def save_config(config_data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")
