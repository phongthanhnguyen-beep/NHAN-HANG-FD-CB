import os
import sys
import json
import time
import queue
import threading
from datetime import datetime
import customtkinter as ctk
from tkinter import ttk
from tkinter import messagebox
from pygame import mixer
import pandas as pd
import gspread
import win32print
import win32ui
import win32con
#phong

# Thử import thư viện qrcode hỗ trợ tạo mã QR
try:
    import qrcode
    HAS_QRCODE = True
except ImportError:
    HAS_QRCODE = False
    print("=" * 70)
    print("❌ CHƯA CÀI THƯ VIỆN 'qrcode' -> TEM IN RA SẼ KHÔNG CÓ MÃ QR!")
    print("👉 Cách khắc phục: mở CMD/Terminal rồi chạy lệnh: pip install qrcode")
    print("=" * 70)

# ========================================================================================
# CẤU HÌNH ĐƯỜNG DẪN VÀ THÔNG SỐ HỆ THỐNG (HỖ TRỢ CẢ KHI ĐÓNG GÓI PYINSTALLER .EXE)
# ========================================================================================
if getattr(sys, 'frozen', False):
    RESOURCE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    DATA_DIR = os.path.dirname(sys.executable)
else:
    RESOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = RESOURCE_DIR

BACKUP_JSON_PATH = os.path.join(DATA_DIR, 'backup_data_fdcb.json')
SHEET_ID = '1zUvOI_9iARGQgf2EFfedZLPnFCMTQzEUVcCqDHwS6Vw'

CREDENTIAL_PATHS = [
    os.path.join(RESOURCE_DIR, 'file2', 'vns-rt-ft-102.json'),
    os.path.join(DATA_DIR, 'file2', 'vns-rt-ft-102.json'),
    os.path.join(DATA_DIR, 'vns-rt-ft-102.json'),
    os.path.join(RESOURCE_DIR, 'vns-rt-ft-102.json'),
    os.path.join(RESOURCE_DIR, 'file2', 'phong-thanhnguyen.json'),
    os.path.join(DATA_DIR, 'file2', 'phong-thanhnguyen.json'),
    os.path.join(DATA_DIR, 'phong-thanhnguyen.json')
]

SHEET_NAME_TYPE = 'Type'
SHEET_NAME_TRAM1 = 'SPX TRẠM 1'
SHEET_NAME_TRAM2 = 'SPX TRẠM 2'
SHEET_NAME_DBT = 'ĐƠN BẤT THƯỜNG'

# --- CẤU HÌNH DANH SÁCH RỔ & GỘP PHÂN LOẠI THEO YÊU CẦU MỚI ---
PREDEFINED_BASKETS = [
    'Resell',
    'LQ',
    'QC Tay',
    'Return To Seller',
    'LQ JP-KR-TH',
    'KO PHÂN LOẠI'
]

# Quy tắc gộp: LQ JP, LQ KR, LQ TH gom chung vào 1 rổ 'LQ JP-KR-TH'
CATEGORY_MAPPING = {
    'LQ JP': 'LQ JP-KR-TH',
    'LQ KR': 'LQ JP-KR-TH',
    'LQ TH': 'LQ JP-KR-TH',
    'Resell': 'Resell',
    'LQ': 'LQ',
    'QC Tay': 'QC Tay',
    'Return To Seller': 'Return To Seller'
}

def get_basket_name(category):
    if not category or pd.isna(category) or str(category).strip().lower() in ['none', 'nan', '', 'null']:
        return "KO PHÂN LOẠI"
    cat_str = str(category).strip()
    return CATEGORY_MAPPING.get(cat_str, cat_str)

# ========================================================================================
# QUẢN LÝ DỮ LIỆU SAO LƯU OFFLINE BẰNG FILE JSON (CHỐNG MẤT DỮ LIỆU KHI TẮT MÁY / MẤT MẠNG)
# ========================================================================================
backup_lock = threading.Lock()

def load_backup_json():
    with backup_lock:
        default_data = {
            "tram1": {"unsubmitted_items": [], "pending_batches": [], "history": []},
            "tram2": {"unsubmitted_items": [], "pending_batches": [], "history": []},
            "tram_bat_thuong": {"unsubmitted_items": [], "pending_batches": [], "history": []}
        }
        if not os.path.exists(BACKUP_JSON_PATH):
            try:
                with open(BACKUP_JSON_PATH, 'w', encoding='utf-8') as f:
                    json.dump(default_data, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"⚠️ Không thể tạo file backup JSON ban đầu: {e}")
            return default_data
        try:
            with open(BACKUP_JSON_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
                for k in ["tram1", "tram2", "tram_bat_thuong"]:
                    if k not in data:
                        data[k] = {"unsubmitted_items": [], "pending_batches": [], "history": []}
                return data
        except Exception as e:
            print(f"⚠️ Lỗi khi đọc file backup JSON: {e}")
            return default_data

def save_backup_json(data):
    with backup_lock:
        temp_path = BACKUP_JSON_PATH + ".tmp"
        try:
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if os.path.exists(BACKUP_JSON_PATH):
                os.replace(temp_path, BACKUP_JSON_PATH)
            else:
                os.rename(temp_path, BACKUP_JSON_PATH)
        except Exception as e:
            print(f"⚠️ Lỗi lưu file backup JSON: {e}")
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass

def json_add_unsubmitted(tram_key, item):
    data = load_backup_json()
    if tram_key not in data:
        data[tram_key] = {"unsubmitted_items": [], "pending_batches": [], "history": []}
    data[tram_key]["unsubmitted_items"].append(item)
    save_backup_json(data)

def json_add_pending_batch(tram_key, batch):
    data = load_backup_json()
    if tram_key not in data:
        data[tram_key] = {"unsubmitted_items": [], "pending_batches": [], "history": []}
    
    tns_set = set(batch.get("tns", []))
    data[tram_key]["unsubmitted_items"] = [
        item for item in data[tram_key]["unsubmitted_items"]
        if item.get("tn") not in tns_set
    ]
    data[tram_key]["pending_batches"].append(batch)
    save_backup_json(data)

def json_mark_batch_completed(tram_key, batch_id, basket_id, rows):
    data = load_backup_json()
    if tram_key not in data:
        data[tram_key] = {"unsubmitted_items": [], "pending_batches": [], "history": []}
    
    data[tram_key]["pending_batches"] = [
        b for b in data[tram_key]["pending_batches"] if b.get("batch_id") != batch_id
    ]
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for r in rows:
        data[tram_key]["history"].append({
            "tn": r[0],
            "phan_loai": r[1],
            "gio_nhan": r[2],
            "basket_id": basket_id,
            "submitted_at": now_str
        })
    if len(data[tram_key]["history"]) > 5000:
        data[tram_key]["history"] = data[tram_key]["history"][-5000:]
    save_backup_json(data)

# ========================================================================================
# KẾT NỐI GOOGLE SHEETS
# ========================================================================================
def get_gspread_client():
    for p in CREDENTIAL_PATHS:
        if os.path.exists(p):
            try:
                gc = gspread.service_account(filename=p)
                return gc
            except Exception as e:
                print(f"Lỗi thử kết nối bằng {p}: {e}")
    raise FileNotFoundError("Không tìm thấy file chứng thực Service Account JSON (vns-rt-ft-102.json)!")

try:
    gc_main = get_gspread_client()
    spreadsheet = gc_main.open_by_key(SHEET_ID)
    ws_type = spreadsheet.worksheet(SHEET_NAME_TYPE)
    ws_tram1 = spreadsheet.worksheet(SHEET_NAME_TRAM1)
    ws_tram2 = spreadsheet.worksheet(SHEET_NAME_TRAM2)
    ws_dbt = spreadsheet.worksheet(SHEET_NAME_DBT)
except Exception as e:
    messagebox.showerror("Lỗi Kết Nối Google Sheets", f"Không thể kết nối đến Google Sheets:\n{e}\nVui lòng kiểm tra file json và kết nối mạng!")
    sys.exit(1)

def ensure_sheet_headers(ws):
    try:
        header = ws.row_values(1)
        if not header:
            ws.append_row(['Mã Vận Chuyển', 'Phân Loại', 'Giờ Nhận', 'Mã Rổ'], value_input_option='USER_ENTERED')
            print(f"✅ Đã tạo tiêu đề cột cho sheet: {ws.title}")
    except Exception as e:
        print(f"⚠️ Lỗi kiểm tra header sheet {ws.title}: {e}")

ensure_sheet_headers(ws_tram1)
ensure_sheet_headers(ws_tram2)
ensure_sheet_headers(ws_dbt)

# ========================================================================================
# NẠP DỮ LIỆU TYPE (CỘT A: ĐƠN HÀNG, CỘT J: PHÂN LOẠI)
# ========================================================================================
source_data_dict = {}  # {return_tn: phan_loai}
scanned_set_tram1 = set()
scanned_set_tram2 = set()
scanned_set_dbt = set()

def fetch_source_type_data():
    global source_data_dict
    raw_data = ws_type.get("A2:J")
    new_dict = {}
    for row in raw_data:
        if row and len(row) >= 1:
            tn = str(row[0]).strip()
            cat = str(row[9]).strip() if len(row) >= 10 else ""
            if tn:
                new_dict[tn] = cat
    source_data_dict = new_dict
    return len(source_data_dict)

def fetch_scanned_tn_sets():
    global scanned_set_tram1, scanned_set_tram2, scanned_set_dbt
    try:
        col_t1 = ws_tram1.col_values(1)[1:]
        scanned_set_tram1 = set(str(x).strip() for x in col_t1 if str(x).strip())
    except Exception as e:
        print(f"⚠️ Lỗi nạp đơn đã quét Trạm 1: {e}")
        scanned_set_tram1 = set()
    try:
        col_t2 = ws_tram2.col_values(1)[1:]
        scanned_set_tram2 = set(str(x).strip() for x in col_t2 if str(x).strip())
    except Exception as e:
        print(f"⚠️ Lỗi nạp đơn đã quét Trạm 2: {e}")
        scanned_set_tram2 = set()
    try:
        col_dbt = ws_dbt.col_values(1)[1:]
        scanned_set_dbt = set(str(x).strip() for x in col_dbt if str(x).strip())
    except Exception as e:
        print(f"⚠️ Lỗi nạp đơn đã quét Đơn Bất Thường: {e}")
        scanned_set_dbt = set()

print("🔄 Đang tải dữ liệu ban đầu từ Google Sheet...")
fetch_source_type_data()
fetch_scanned_tn_sets()
print(f"✅ Đã nạp thành công {len(source_data_dict)} mã đơn hàng từ Sheet 'Type'.")

def reload_source_data():
    def _worker():
        root.after(0, lambda: btn_reload_source.configure(state="disabled", text="⏳ Đang nạp..."))
        try:
            count = fetch_source_type_data()
            fetch_scanned_tn_sets()
            now_str = datetime.now().strftime("%H:%M:%S")
            print(f"✅ Đã nạp lại source_data thành công: {count} dòng lúc {now_str}.")
            root.after(0, lambda: label_source_status.configure(
                text=f"📄 Source Data: {count} dòng | Cập nhật lúc {now_str}"))
            root.after(0, lambda: messagebox.showinfo(
                "Thành Công", f"Đã nạp lại {count} dòng từ Sheet 'Type' thành công!"))
        except Exception as e:
            print(f"⚠️ Lỗi khi nạp lại source_data: {e}")
            root.after(0, lambda err=e: messagebox.showerror(
                "Lỗi Nạp Data", f"Không thể nạp lại dữ liệu từ sheet 'Type':\n{err}"))
        finally:
            root.after(0, lambda: btn_reload_source.configure(state="normal", text="🔄 Nạp Lại Source Data"))
    threading.Thread(target=_worker, daemon=True).start()

# ========================================================================================
# QUẢN LÝ BỘ NHỚ VÀ BIẾN ĐẾM (3 TRẠM)
# ========================================================================================
baskets_data_tram1 = {b: [] for b in PREDEFINED_BASKETS}
baskets_data_tram2 = {b: [] for b in PREDEFINED_BASKETS}
baskets_data_dbt = {b: [] for b in PREDEFINED_BASKETS}

baskets_ui_tram1 = {}
baskets_ui_tram2 = {}
baskets_ui_dbt = {}

active_basket_ids = {
    "tram1": {},
    "tram2": {},
    "dbt": {}
}

count_pending_tram1 = 0
count_total_tram1 = 0
count_pending_tram2 = 0
count_total_tram2 = 0
count_pending_dbt = 0
count_total_dbt = 0

session_scans_tram1 = []
session_scans_tram2 = []
session_scans_dbt = []

# ========================================================================================
# TẠO MÃ RỔ THEO THỜI GIAN NGAY LÚC ĐÓ (SPX1 Resell 02/10-0121)
# ========================================================================================
def generate_basket_id(station_tag, basket_name):
    now = datetime.now()
    date_str = now.strftime("%d/%m")
    time_str = now.strftime("%H%M")
    return f"{station_tag} {basket_name} {date_str}-{time_str}"

def get_or_create_active_basket_id(station_type, basket_name):
    if station_type == "tram1":
        station_tag = "SPX1"
    elif station_type == "tram2":
        station_tag = "SPX2"
    else:
        station_tag = "DBT"

    if basket_name not in active_basket_ids[station_type] or not active_basket_ids[station_type][basket_name]:
        active_basket_ids[station_type][basket_name] = generate_basket_id(station_tag, basket_name)
    
    return active_basket_ids[station_type][basket_name]

# ========================================================================================
# ÂM THANH PHẢN HỒI (VOICE)
# ========================================================================================
def voice(voice_name):
    try:
        mixer.init()
        sound_path = os.path.join(RESOURCE_DIR, 'file2', f"{voice_name}.mp3")
        if not os.path.exists(sound_path):
            sound_path = os.path.join(DATA_DIR, 'file2', f"{voice_name}.mp3")
        if os.path.exists(sound_path):
            mixer.music.load(sound_path)
            mixer.music.play()
        else:
            print(f"⚠️ Không tìm thấy file âm thanh: {sound_path}")
    except Exception as e:
        print(f"Lỗi âm thanh: {e}")

# ========================================================================================
# IN NHÃN NHIỆT KHỔ LỚN 10cm x 5cm (KÈM MÃ QR & TRÌNH BÀY XUỐNG DÒNG RÕ RÀNG)
# ========================================================================================
def print_basket_label(basket_id, qty=0):
    try:
        printer_name = win32print.GetDefaultPrinter()
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        
        hdc.StartDoc("Basket Label QR")
        hdc.StartPage()
        hdc.SetMapMode(win32con.MM_TEXT)
        
        dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)
        
        width_px = int(3.937 * dpi_x)
        height_px = int(1.968 * dpi_y)
        
        margin = int(0.12 * dpi_x)
        pen = win32ui.CreatePen(win32con.PS_SOLID, int(dpi_x * 0.02), 0)
        old_pen = hdc.SelectObject(pen)
        hdc.Rectangle((margin, margin, width_px - margin, height_px - margin))
        hdc.SelectObject(old_pen)
        
        qr_size_px = int(1.45 * dpi_x)
        qr_x = width_px - margin - qr_size_px - int(0.08 * dpi_x)
        qr_y = (height_px - qr_size_px) // 2
        
        if HAS_QRCODE:
            try:
                qr = qrcode.QRCode(version=1, box_size=1, border=4, error_correction=qrcode.constants.ERROR_CORRECT_L)
                qr.add_data(basket_id)
                qr.make(fit=True)
                matrix = qr.get_matrix()
                
                num_blocks = len(matrix)
                block_size = qr_size_px // num_blocks
                actual_qr_size = block_size * num_blocks
                
                qr_x_aligned = qr_x + (qr_size_px - actual_qr_size) // 2
                qr_y_aligned = qr_y + (qr_size_px - actual_qr_size) // 2
                
                brush = win32ui.CreateBrush(win32con.BS_SOLID, 0, 0)
                old_brush = hdc.SelectObject(brush)
                pen_null = win32ui.CreatePen(win32con.PS_SOLID, 0, 0)
                old_pen2 = hdc.SelectObject(pen_null)
                
                for r in range(num_blocks):
                    for c in range(num_blocks):
                        if matrix[r][c]:
                            bx1 = qr_x_aligned + c * block_size
                            by1 = qr_y_aligned + r * block_size
                            bx2 = bx1 + block_size
                            by2 = by1 + block_size
                            hdc.Rectangle((bx1, by1, bx2, by2))
                            
                hdc.SelectObject(old_brush)
                hdc.SelectObject(old_pen2)
            except Exception as qr_err:
                print(f"⚠️ Lỗi khởi dựng mã QR: {qr_err}")
        
        font_title_h = int(dpi_y * 0.10)
        font_id_h_max = int(dpi_y * 0.35)
        font_id_h_min = int(dpi_y * 0.22)
        font_seq_h = int(dpi_y * 0.38)
        font_qty_h = int(dpi_y * 0.24)
        font_footer_h = int(dpi_y * 0.10)
        
        font_title = win32ui.CreateFont({"name": "Arial", "height": font_title_h, "weight": win32con.FW_BOLD})
        font_seq = win32ui.CreateFont({"name": "Arial", "height": font_seq_h, "weight": win32con.FW_BOLD})
        font_qty = win32ui.CreateFont({"name": "Arial", "height": font_qty_h, "weight": win32con.FW_BOLD})
        font_footer = win32ui.CreateFont({"name": "Arial", "height": font_footer_h, "weight": win32con.FW_NORMAL})
        
        text_margin = int(margin * 1.3)
        max_text_width = qr_x - text_margin - int(0.05 * dpi_x)
        
        hdc.SelectObject(font_title)
        hdc.TextOut(text_margin, int(margin * 1.2), "CHỐT RỔ / COMPLETED")
        
        if ' ' in basket_id:
            line_name, line_seq = basket_id.rsplit(' ', 1)
        else:
            line_name, line_seq = basket_id, ""
        
        fit_h = font_id_h_max
        while True:
            font_id = win32ui.CreateFont({"name": "Arial", "height": fit_h, "weight": win32con.FW_BOLD})
            hdc.SelectObject(font_id)
            text_w, _ = hdc.GetTextExtent(line_name)
            if text_w <= max_text_width or fit_h <= font_id_h_min:
                break
            fit_h -= 2
        
        line1_y = int(margin * 1.2 + font_title_h * 1.1)
        hdc.TextOut(text_margin, line1_y, line_name)
        
        hdc.SelectObject(font_seq)
        line2_y = int(line1_y + fit_h * 1.05)
        hdc.TextOut(text_margin, line2_y, line_seq)

        hdc.SelectObject(font_qty)
        line3_y = int(line2_y + font_seq_h * 1.05)
        hdc.TextOut(text_margin, line3_y, f"SL: {qty}")
        
        hdc.SelectObject(font_footer)
        now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        hdc.TextOut(text_margin, int(height_px - margin * 1.3 - font_footer_h * 2.0), f"Thời gian: {now_str}")
        hdc.TextOut(text_margin, int(height_px - margin * 1.3 - font_footer_h * 1.0), "Shopee VN Reverse Logistics")
        
        hdc.EndPage()
        hdc.EndDoc()
        print(f"🖨️ Đã in tem rổ: {basket_id} | SL: {qty} thành công!")
    except Exception as e:
        print(f"⚠️ Lỗi in tem: {e}")

# ========================================================================================
# CƠ CHẾ ĐẨY GOOGLE SHEETS AN TOÀN, TỰ MỞ THÊM DÒNG & HÀNG ĐỢI SUBMIT ĐA LUỒNG
# ========================================================================================
upload_queue = queue.Queue()
print_queue = queue.Queue()
sheet_write_lock = threading.Lock()


def print_worker_thread():
    """Print labels independently so a slow Sheets request cannot delay printing."""
    while True:
        task = print_queue.get()
        if task is None:
            print_queue.task_done()
            break
        basket_id, qty = task
        try:
            print_basket_label(basket_id, qty=qty)
        except Exception as err:
            print(f"Lỗi worker in tem rổ [{basket_id}]: {err}")
        finally:
            print_queue.task_done()

def safe_append_to_google_sheet(ws_sheet, rows):
    with sheet_write_lock:
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                try:
                    current_total_rows = ws_sheet.row_count
                    col1_vals = ws_sheet.col_values(1)
                    filled_rows = len(col1_vals)
                    needed_rows = filled_rows + len(rows)
                    if needed_rows >= current_total_rows - 20:
                        rows_to_add = max(1000, len(rows) + 500)
                        ws_sheet.add_rows(rows_to_add)
                        print(f"➕ Đã tự động mở thêm {rows_to_add} dòng cho sheet [{ws_sheet.title}].")
                except Exception as ex_chk:
                    print(f"⚠️ Cảnh báo kiểm tra mở rộng dòng: {ex_chk}")

                ws_sheet.append_rows(rows, value_input_option='USER_ENTERED')
                return True
            except Exception as e:
                print(f"🚨 [Lần thử {attempt}/{max_retries}] Lỗi đẩy Sheet [{ws_sheet.title}]: {e}")
                if attempt < max_retries:
                    time.sleep(2 * attempt)
                else:
                    raise e

def upload_worker_thread():
    while True:
        try:
            task = upload_queue.get()
            if task is None:
                break
            
            station_type = task["station_type"]
            batch_id = task["batch_id"]
            basket_id = task["basket_id"]
            rows = task["rows"]
            qty = task["qty"]
            tns = task["tns"]

            if station_type == "tram1":
                ws_target = ws_tram1
                scanned_target = scanned_set_tram1
                tram_key = "tram1"
            elif station_type == "tram2":
                ws_target = ws_tram2
                scanned_target = scanned_set_tram2
                tram_key = "tram2"
            else:
                ws_target = ws_dbt
                scanned_target = scanned_set_dbt
                tram_key = "tram_bat_thuong"

            print(f"🔄 [Submit Worker] Bắt đầu đẩy {qty} đơn rổ [{basket_id}] lên sheet '{ws_target.title}'...")
            
            try:
                safe_append_to_google_sheet(ws_target, rows)
                print(f"✅ [Submit Worker] Chốt thành công rổ [{basket_id}] lên Google Sheets!")

                json_mark_batch_completed(tram_key, batch_id, basket_id, rows)

                for tn in tns:
                    scanned_target.add(tn)

            except Exception as err:
                print(f"❌ [Submit Worker] Không thể đẩy rổ [{basket_id}] do lỗi mạng: {err}")
                root.after(0, lambda b_id=basket_id, q=qty, tr_title=ws_target.title: messagebox.showwarning(
                    "Lỗi Đẩy Google Sheet",
                    f"Mất mạng hoặc lỗi kết nối! Toàn bộ {q} đơn rổ [{b_id}] đã được lưu an toàn trong file JSON.\n"
                    f"Dữ liệu sẽ tự động đẩy lên sheet '{tr_title}' khi có mạng trở lại!"
                ))

            upload_queue.task_done()
        except Exception as q_err:
            print(f"Lỗi worker: {q_err}")

threading.Thread(target=upload_worker_thread, daemon=True).start()
threading.Thread(target=print_worker_thread, daemon=True).start()

def auto_sync_pending_batches():
    def _sync_worker():
        try:
            data = load_backup_json()
            mapping = [
                ("tram1", "tram1"),
                ("tram2", "tram2"),
                ("tram_bat_thuong", "dbt")
            ]
            for json_key, st_type in mapping:
                pending_list = data.get(json_key, {}).get("pending_batches", [])
                for batch in pending_list:
                    batch_id = batch.get("batch_id")
                    upload_queue.put({
                        "station_type": st_type,
                        "batch_id": batch_id,
                        "basket_id": batch.get("basket_id"),
                        "rows": batch.get("rows", []),
                        "qty": batch.get("qty", 0),
                        "tns": batch.get("tns", [])
                    })
        except Exception as e:
            print(f"Lỗi auto sync: {e}")
    threading.Thread(target=_sync_worker, daemon=True).start()

# ========================================================================================
# LOGIC CHỐT RỔ (SUBMIT BASKET) CHO CẢ 3 TRẠM
# ========================================================================================
def submit_basket(station_type, basket_name):
    global count_pending_tram1, count_total_tram1
    global count_pending_tram2, count_total_tram2
    global count_pending_dbt, count_total_dbt

    if station_type == "tram1":
        baskets_data = baskets_data_tram1
        baskets_ui = baskets_ui_tram1
        tree_widget = tree_tram1
        tram_key = "tram1"
        station_tag = "SPX1"
    elif station_type == "tram2":
        baskets_data = baskets_data_tram2
        baskets_ui = baskets_ui_tram2
        tree_widget = tree_tram2
        tram_key = "tram2"
        station_tag = "SPX2"
    else:
        baskets_data = baskets_data_dbt
        baskets_ui = baskets_ui_dbt
        tree_widget = tree_dbt
        tram_key = "tram_bat_thuong"
        station_tag = "DBT"

    tns_in_basket = baskets_data.get(basket_name, [])

    if not tns_in_basket:
        messagebox.showwarning("Trống", f"Rổ '{basket_name}' hiện không có đơn nào cần chốt!")
        return

    # 1. LẤY MÃ RỔ ĐANG MỞ HOẶC SINH MỚI
    basket_id = active_basket_ids[station_type].get(basket_name)
    if not basket_id:
        basket_id = generate_basket_id(station_tag, basket_name)

    # 2. ĐÓNG GÓI DỮ LIỆU VỚI MÃ RỔ ĐÃ GÁN
    rows_to_upload = []
    tns_set = set(tns_in_basket)

    for child in tree_widget.get_children():
        values = tree_widget.item(child, 'values')
        if values[0] in tns_set:
            assigned_b_id = values[3] if len(values) >= 4 and values[3] else basket_id
            rows_to_upload.append([values[0], values[1], values[2], assigned_b_id])

    if not rows_to_upload:
        for tn in tns_in_basket:
            rows_to_upload.append([tn, basket_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), basket_id])

    total_qty = len(rows_to_upload)
    batch_id = f"{tram_key}_{basket_name}_{int(time.time()*1000)}"

    # 3. LƯU VÀO JSON DỰ PHÒNG NGAY LẬP TỨC
    batch_record = {
        "batch_id": batch_id,
        "basket_id": basket_id,
        "basket_name": basket_name,
        "qty": total_qty,
        "rows": rows_to_upload,
        "tns": tns_in_basket,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    json_add_pending_batch(tram_key, batch_record)

    # 4. GIẢI PHÓNG GIAO DIỆN TỨC THÌ ĐỂ NGƯỜI DÙNG QUÉT TIẾP
    children_to_delete = []
    for child in tree_widget.get_children():
        values = tree_widget.item(child, 'values')
        if values[0] in tns_set:
            children_to_delete.append(child)
    for child in children_to_delete:
        tree_widget.delete(child)

    baskets_data[basket_name] = []
    baskets_ui[basket_name].configure(text="0")
    
    active_basket_ids[station_type][basket_name] = None

    if station_type == "tram1":
        count_pending_tram1 = max(0, count_pending_tram1 - total_qty)
        count_total_tram1 += total_qty
        qty_pending_tram1.configure(text=str(count_pending_tram1))
        qty_total_tram1.configure(text=str(count_total_tram1))
        entry_tram1.focus()
    elif station_type == "tram2":
        count_pending_tram2 = max(0, count_pending_tram2 - total_qty)
        count_total_tram2 += total_qty
        qty_pending_tram2.configure(text=str(count_pending_tram2))
        qty_total_tram2.configure(text=str(count_total_tram2))
        entry_tram2.focus()
    else:
        count_pending_dbt = max(0, count_pending_dbt - total_qty)
        count_total_dbt += total_qty
        qty_pending_dbt.configure(text=str(count_pending_dbt))
        qty_total_dbt.configure(text=str(count_total_dbt))
        entry_dbt.focus()

    # 5. ĐƯA VÀO HÀNG ĐỢI SUBMIT NỀN
    upload_queue.put({
        "station_type": station_type,
        "batch_id": batch_id,
        "basket_id": basket_id,
        "rows": rows_to_upload,
        "qty": total_qty,
        "tns": tns_in_basket
    })
    # Print immediately on the dedicated printer worker, in parallel with upload.
    print_queue.put((basket_id, total_qty))

def retry_all_pending_baskets(station_type):
    auto_sync_pending_batches()
    messagebox.showinfo("Đồng Bộ", "Đang tiến hành đồng bộ các lô đơn tồn đọng lên Google Sheets...")

# ========================================================================================
# TẠO GIAO DIỆN THẺ RỔ PHÂN LOẠI
# ========================================================================================
def create_basket_ui(station_type, basket_name):
    if station_type == "tram1":
        parent_frame = baskets_frame_tram1
        baskets_ui = baskets_ui_tram1
    elif station_type == "tram2":
        parent_frame = baskets_frame_tram2
        baskets_ui = baskets_ui_tram2
    else:
        parent_frame = baskets_frame_dbt
        baskets_ui = baskets_ui_dbt
    
    frame = ctk.CTkFrame(parent_frame, fg_color="#2b2b2b", corner_radius=8, border_width=1, border_color="#3d3d3d")
    frame.pack(fill="x", pady=4, padx=5)
    
    lbl_name = ctk.CTkLabel(frame, text=basket_name, font=('Arial', 13, 'bold'), text_color="#ff9f00", width=130, anchor="w")
    lbl_name.grid(row=0, column=0, padx=8, pady=6)
    
    lbl_count = ctk.CTkLabel(frame, text="0", font=('Arial', 15, 'bold'), text_color="#10b981", width=35)
    lbl_count.grid(row=0, column=1, padx=6, pady=6)
    baskets_ui[basket_name] = lbl_count 
    
    btn_submit = ctk.CTkButton(frame, text="Chốt & In 🖨️", font=('Arial', 11, 'bold'), width=85,
                               fg_color="#0284c7", hover_color="#0369a1",
                               command=lambda b=basket_name: submit_basket(station_type, b))
    btn_submit.grid(row=0, column=2, padx=8, pady=6)

# ========================================================================================
# XỬ LÝ QUÉT ĐƠN HÀNG TRẠM 1
# ========================================================================================
def save_tram1(event):
    global count_pending_tram1
    return_tn = entry_tram1.get().strip()
    if not return_tn:
        return
    
    entry_tram1.delete(0, 'end')
    
    if return_tn in scanned_set_tram1:
        voice("duplicate_voice")
        label_loaihang_tram1.configure(text="Trùng Đơn (Sheets)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    if return_tn not in source_data_dict:
        voice("nodata_voice")
        label_loaihang_tram1.configure(text="No_data", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    raw_fn = source_data_dict.get(return_tn, "")
    if not raw_fn or str(raw_fn).strip().lower() in ['nan', 'none', '', 'null']:
        fn = "KO PHÂN LOẠI"
    else:
        fn = str(raw_fn).strip()

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if return_tn in session_scans_tram1:
        if fn == "KO PHÂN LOẠI":
            voice("NR_voice")
        else:
            voice(fn)
        label_loaihang_tram1.configure(text=f"{fn} (Đã quét)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    session_scans_tram1.append(return_tn)
    
    if fn == "KO PHÂN LOẠI":
        voice("NR_voice")
        label_loaihang_tram1.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#ff9f00")
    else:
        voice(fn)
        label_loaihang_tram1.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#10b981")

    basket_name = get_basket_name(fn)
    if basket_name not in baskets_data_tram1:
        baskets_data_tram1[basket_name] = []
        create_basket_ui(station_type="tram1", basket_name=basket_name)

    baskets_data_tram1[basket_name].append(return_tn)
    baskets_ui_tram1[basket_name].configure(text=str(len(baskets_data_tram1[basket_name])))

    count_pending_tram1 += 1
    qty_pending_tram1.configure(text=str(count_pending_tram1))

    basket_id = get_or_create_active_basket_id(station_type="tram1", basket_name=basket_name)

    json_add_unsubmitted("tram1", {
        "tn": return_tn,
        "phan_loai": fn,
        "gio_nhan": now_str,
        "basket_name": basket_name,
        "basket_id": basket_id
    })

    tree_tram1.insert('', 'end', values=(return_tn, fn, now_str, basket_id))
    tree_tram1.see(tree_tram1.get_children()[-1])

# ========================================================================================
# XỬ LÝ QUÉT ĐƠN HÀNG TRẠM 2
# ========================================================================================
def save_tram2(event):
    global count_pending_tram2
    return_tn = entry_tram2.get().strip()
    if not return_tn:
        return
    
    entry_tram2.delete(0, 'end')
    
    if return_tn in scanned_set_tram2:
        voice("duplicate_voice")
        label_loaihang_tram2.configure(text="Trùng Đơn (Sheets)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    if return_tn not in source_data_dict:
        voice("nodata_voice")
        label_loaihang_tram2.configure(text="No_data", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    raw_fn = source_data_dict.get(return_tn, "")
    if not raw_fn or str(raw_fn).strip().lower() in ['nan', 'none', '', 'null']:
        fn = "KO PHÂN LOẠI"
    else:
        fn = str(raw_fn).strip()

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if return_tn in session_scans_tram2:
        if fn == "KO PHÂN LOẠI":
            voice("NR_voice")
        else:
            voice(fn)
        label_loaihang_tram2.configure(text=f"{fn} (Đã quét)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    session_scans_tram2.append(return_tn)
    
    if fn == "KO PHÂN LOẠI":
        voice("NR_voice")
        label_loaihang_tram2.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#ff9f00")
    else:
        voice(fn)
        label_loaihang_tram2.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#10b981")

    basket_name = get_basket_name(fn)
    if basket_name not in baskets_data_tram2:
        baskets_data_tram2[basket_name] = []
        create_basket_ui(station_type="tram2", basket_name=basket_name)

    baskets_data_tram2[basket_name].append(return_tn)
    baskets_ui_tram2[basket_name].configure(text=str(len(baskets_data_tram2[basket_name])))

    count_pending_tram2 += 1
    qty_pending_tram2.configure(text=str(count_pending_tram2))

    basket_id = get_or_create_active_basket_id(station_type="tram2", basket_name=basket_name)

    json_add_unsubmitted("tram2", {
        "tn": return_tn,
        "phan_loai": fn,
        "gio_nhan": now_str,
        "basket_name": basket_name,
        "basket_id": basket_id
    })

    tree_tram2.insert('', 'end', values=(return_tn, fn, now_str, basket_id))
    tree_tram2.see(tree_tram2.get_children()[-1])

# ========================================================================================
# XỬ LÝ QUÉT ĐƠN HÀNG TRẠM ĐƠN BẤT THƯỜNG
# ========================================================================================
def save_dbt(event):
    global count_pending_dbt
    return_tn = entry_dbt.get().strip()
    if not return_tn:
        return
    
    entry_dbt.delete(0, 'end')
    
    if return_tn in scanned_set_dbt:
        voice("duplicate_voice")
        label_loaihang_dbt.configure(text="Trùng Đơn (Sheets)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    if return_tn not in source_data_dict:
        voice("nodata_voice")
        label_loaihang_dbt.configure(text="No_data", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    raw_fn = source_data_dict.get(return_tn, "")
    if not raw_fn or str(raw_fn).strip().lower() in ['nan', 'none', '', 'null']:
        fn = "KO PHÂN LOẠI"
    else:
        fn = str(raw_fn).strip()

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if return_tn in session_scans_dbt:
        if fn == "KO PHÂN LOẠI":
            voice("NR_voice")
        else:
            voice(fn)
        label_loaihang_dbt.configure(text=f"{fn} (Đã quét)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return

    session_scans_dbt.append(return_tn)
    
    if fn == "KO PHÂN LOẠI":
        voice("NR_voice")
        label_loaihang_dbt.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#ff9f00")
    else:
        voice(fn)
        label_loaihang_dbt.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#10b981")

    basket_name = get_basket_name(fn)
    if basket_name not in baskets_data_dbt:
        baskets_data_dbt[basket_name] = []
        create_basket_ui(station_type="dbt", basket_name=basket_name)

    baskets_data_dbt[basket_name].append(return_tn)
    baskets_ui_dbt[basket_name].configure(text=str(len(baskets_data_dbt[basket_name])))

    count_pending_dbt += 1
    qty_pending_dbt.configure(text=str(count_pending_dbt))

    basket_id = get_or_create_active_basket_id(station_type="dbt", basket_name=basket_name)

    json_add_unsubmitted("tram_bat_thuong", {
        "tn": return_tn,
        "phan_loai": fn,
        "gio_nhan": now_str,
        "basket_name": basket_name,
        "basket_id": basket_id
    })

    tree_dbt.insert('', 'end', values=(return_tn, fn, now_str, basket_id))
    tree_dbt.see(tree_dbt.get_children()[-1])

# ========================================================================================
# KHÔI PHỤC DỮ LIỆU TỪ FILE JSON DỰ PHÒNG KHI KHỞI ĐỘNG (CẢ 3 TRẠM)
# ========================================================================================
def restore_offline_backup():
    global count_pending_tram1, count_pending_tram2, count_pending_dbt
    data = load_backup_json()
    
    # Khôi phục Trạm 1
    unsub_t1 = data.get("tram1", {}).get("unsubmitted_items", [])
    for item in unsub_t1:
        tn = item.get("tn")
        fn = item.get("phan_loai")
        gio = item.get("gio_nhan")
        b_name = item.get("basket_name") or get_basket_name(fn)
        b_id = item.get("basket_id") or get_or_create_active_basket_id("tram1", b_name)
        active_basket_ids["tram1"][b_name] = b_id
        
        session_scans_tram1.append(tn)
        tree_tram1.insert('', 'end', values=(tn, fn, gio, b_id))
        
        if b_name not in baskets_data_tram1:
            baskets_data_tram1[b_name] = []
        baskets_data_tram1[b_name].append(tn)
        count_pending_tram1 += 1

    for b_name in baskets_data_tram1:
        if b_name in baskets_ui_tram1:
            baskets_ui_tram1[b_name].configure(text=str(len(baskets_data_tram1[b_name])))
    qty_pending_tram1.configure(text=str(count_pending_tram1))

    # Khôi phục Trạm 2
    unsub_t2 = data.get("tram2", {}).get("unsubmitted_items", [])
    for item in unsub_t2:
        tn = item.get("tn")
        fn = item.get("phan_loai")
        gio = item.get("gio_nhan")
        b_name = item.get("basket_name") or get_basket_name(fn)
        b_id = item.get("basket_id") or get_or_create_active_basket_id("tram2", b_name)
        active_basket_ids["tram2"][b_name] = b_id
        
        session_scans_tram2.append(tn)
        tree_tram2.insert('', 'end', values=(tn, fn, gio, b_id))
        
        if b_name not in baskets_data_tram2:
            baskets_data_tram2[b_name] = []
        baskets_data_tram2[b_name].append(tn)
        count_pending_tram2 += 1

    for b_name in baskets_data_tram2:
        if b_name in baskets_ui_tram2:
            baskets_ui_tram2[b_name].configure(text=str(len(baskets_data_tram2[b_name])))
    qty_pending_tram2.configure(text=str(count_pending_tram2))

    # Khôi phục Đơn Bất Thường
    unsub_dbt = data.get("tram_bat_thuong", {}).get("unsubmitted_items", [])
    for item in unsub_dbt:
        tn = item.get("tn")
        fn = item.get("phan_loai")
        gio = item.get("gio_nhan")
        b_name = item.get("basket_name") or get_basket_name(fn)
        b_id = item.get("basket_id") or get_or_create_active_basket_id("dbt", b_name)
        active_basket_ids["dbt"][b_name] = b_id
        
        session_scans_dbt.append(tn)
        tree_dbt.insert('', 'end', values=(tn, fn, gio, b_id))
        
        if b_name not in baskets_data_dbt:
            baskets_data_dbt[b_name] = []
        baskets_data_dbt[b_name].append(tn)
        count_pending_dbt += 1

    for b_name in baskets_data_dbt:
        if b_name in baskets_ui_dbt:
            baskets_ui_dbt[b_name].configure(text=str(len(baskets_data_dbt[b_name])))
    qty_pending_dbt.configure(text=str(count_pending_dbt))

    total_restored = len(unsub_t1) + len(unsub_t2) + len(unsub_dbt)
    if total_restored > 0:
        print(f"♻️ Đã khôi phục thành công {total_restored} đơn tồn đọng từ file JSON dự phòng!")

    auto_sync_pending_batches()

# ========================================================================================
# GIAO DIỆN CHÍNH (CUSTOMTKINTER & TREEVIEW CHO 3 TRẠM)
# ========================================================================================
ctk.set_appearance_mode("dark")  
ctk.set_default_color_theme("blue")  
root = ctk.CTk()
root.title("Tool BY PHONG Nhận Hàng FD-CB GITHUB")
root.geometry("1400x780") 

style = ttk.Style()
style.theme_use('clam')  
style.configure("Treeview", background="#333333", foreground="white", fieldbackground="#333333", borderwidth=0, font=('Arial', 12)) 
style.configure("Treeview.Heading", background="#ff9f00", foreground="white", font=('Arial', 13, 'bold')) 
style.map('Treeview', background=[('selected', '#144870')], foreground=[('selected', 'white')])

# THANH TRÊN CÙNG
top_bar = ctk.CTkFrame(root, fg_color="#1f1f1f", corner_radius=8)
top_bar.pack(fill="x", padx=15, pady=(10, 0))

label_source_status = ctk.CTkLabel(
    top_bar,
    text=f"📄 Source Data (Sheet Type): {len(source_data_dict)} dòng | Cập nhật lúc {datetime.now().strftime('%H:%M:%S')}",
    font=('Arial', 12), text_color="#9ca3af")
label_source_status.pack(side="left", padx=15, pady=8)

btn_reload_source = ctk.CTkButton(
    top_bar, text="🔄 Nạp Lại Source Data", font=('Arial', 12, 'bold'),
    width=190, height=30, fg_color="#15803d", hover_color="#166534",
    command=reload_source_data)
btn_reload_source.pack(side="right", padx=15, pady=8)

tabview = ctk.CTkTabview(root)
tabview.pack(fill="both", expand=True, padx=15, pady=5)

tabview.add("SPX TRẠM 1")
tabview.add("SPX TRẠM 2")
tabview.add("ĐƠN BẤT THƯỜNG")

columns_tree = ("Mã Vận Chuyển", "Phân Loại", "Giờ Nhận", "Mã Rổ")

# ================= TAB 1: SPX TRẠM 1 =================
tab1 = tabview.tab("SPX TRẠM 1")
tab1.grid_columnconfigure(0, weight=6)
tab1.grid_columnconfigure(1, weight=4)
tab1.grid_rowconfigure(0, weight=1)

main_frame_tram1 = ctk.CTkFrame(tab1, fg_color="transparent")
main_frame_tram1.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=5)

baskets_frame_tram1 = ctk.CTkScrollableFrame(tab1, label_text="HỆ THỐNG RỔ PHÂN LOẠI TRẠM 1", label_font=('Arial', 14, 'bold'), label_text_color="#ff9f00")
baskets_frame_tram1.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

main_frame_tram1.grid_columnconfigure(0, weight=1)
main_frame_tram1.grid_columnconfigure(1, weight=2)
main_frame_tram1.grid_rowconfigure(4, weight=1)

ctk.CTkLabel(main_frame_tram1, text="Scan Return TN:", font=('Arial', 14, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_tram1 = ctk.CTkEntry(main_frame_tram1, width=280, font=('Arial', 14))
entry_tram1.grid(row=0, column=1, padx=10, pady=10, sticky="w")
entry_tram1.bind("<Return>", save_tram1)

ctk.CTkLabel(main_frame_tram1, text='Loại Hàng:', font=('Arial', 14, 'bold')).grid(row=1, column=0, padx=10, pady=5, sticky="e")
label_loaihang_tram1 = ctk.CTkLabel(main_frame_tram1, text='No_data', font=('Arial', 24, 'bold'), text_color="#10b981")
label_loaihang_tram1.grid(row=1, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_tram1, text='Số đơn chưa chốt:', font=('Arial', 13)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
qty_pending_tram1 = ctk.CTkLabel(main_frame_tram1, text=count_pending_tram1, font=('Arial', 14, 'bold'), text_color="#f59e0b")
qty_pending_tram1.grid(row=2, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_tram1, text='Số đơn đã chốt rổ:', font=('Arial', 13)).grid(row=3, column=0, padx=10, pady=5, sticky="e")
qty_total_tram1 = ctk.CTkLabel(main_frame_tram1, text=count_total_tram1, font=('Arial', 14, 'bold'), text_color="#3b82f6")
qty_total_tram1.grid(row=3, column=1, padx=10, pady=5, sticky="w")

tree_container_tram1 = ctk.CTkFrame(main_frame_tram1, fg_color="transparent")
tree_container_tram1.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
tree_container_tram1.grid_columnconfigure(0, weight=1)
tree_container_tram1.grid_rowconfigure(0, weight=1)

tree_tram1 = ttk.Treeview(tree_container_tram1, columns=columns_tree, show="headings", style="Treeview")
for col in columns_tree:
    tree_tram1.heading(col, text=col)
tree_tram1.column("Mã Vận Chuyển", width=170, anchor="center")
tree_tram1.column("Phân Loại", width=110, anchor="center")
tree_tram1.column("Giờ Nhận", width=140, anchor="center")
tree_tram1.column("Mã Rổ", width=220, anchor="center")
tree_tram1.pack(side="left", fill="both", expand=True)

v_scrollbar_tram1 = ttk.Scrollbar(tree_container_tram1, orient="vertical", command=tree_tram1.yview)
tree_tram1.configure(yscrollcommand=v_scrollbar_tram1.set)
v_scrollbar_tram1.pack(side="right", fill="y")

button_sync_tram1 = ctk.CTkButton(main_frame_tram1, text="THỬ LẠI / ĐỒNG BỘ CÁC ĐƠN LỖI MẠNG LÊN GOOGLE SHEET", font=('Arial', 13, 'bold'),
                                  height=38, fg_color="#1e3a8a", hover_color="#1e40af", command=lambda: retry_all_pending_baskets(station_type="tram1"))
button_sync_tram1.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# ================= TAB 2: SPX TRẠM 2 =================
tab2 = tabview.tab("SPX TRẠM 2")
tab2.grid_columnconfigure(0, weight=6)
tab2.grid_columnconfigure(1, weight=4)
tab2.grid_rowconfigure(0, weight=1)

main_frame_tram2 = ctk.CTkFrame(tab2, fg_color="transparent")
main_frame_tram2.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=5)

baskets_frame_tram2 = ctk.CTkScrollableFrame(tab2, label_text="HỆ THỐNG RỔ PHÂN LOẠI TRẠM 2", label_font=('Arial', 14, 'bold'), label_text_color="#ff9f00")
baskets_frame_tram2.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

main_frame_tram2.grid_columnconfigure(0, weight=1)
main_frame_tram2.grid_columnconfigure(1, weight=2)
main_frame_tram2.grid_rowconfigure(4, weight=1)

ctk.CTkLabel(main_frame_tram2, text="Scan Return TN:", font=('Arial', 14, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_tram2 = ctk.CTkEntry(main_frame_tram2, width=280, font=('Arial', 14))
entry_tram2.grid(row=0, column=1, padx=10, pady=10, sticky="w")
entry_tram2.bind("<Return>", save_tram2)

ctk.CTkLabel(main_frame_tram2, text='Loại Hàng:', font=('Arial', 14, 'bold')).grid(row=1, column=0, padx=10, pady=5, sticky="e")
label_loaihang_tram2 = ctk.CTkLabel(main_frame_tram2, text='No_data', font=("Arial", 22, 'bold'), text_color="#10b981")
label_loaihang_tram2.grid(row=1, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_tram2, text='Số đơn chưa chốt:', font=('Arial', 13)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
qty_pending_tram2 = ctk.CTkLabel(main_frame_tram2, text=count_pending_tram2, font=('Arial', 14, 'bold'), text_color="#f59e0b")
qty_pending_tram2.grid(row=2, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_tram2, text='Số đơn đã chốt rổ:', font=('Arial', 13)).grid(row=3, column=0, padx=10, pady=5, sticky="e")
qty_total_tram2 = ctk.CTkLabel(main_frame_tram2, text=count_total_tram2, font=('Arial', 14, 'bold'), text_color="#3b82f6")
qty_total_tram2.grid(row=3, column=1, padx=10, pady=5, sticky="w")

tree_container_tram2 = ctk.CTkFrame(main_frame_tram2, fg_color="transparent")
tree_container_tram2.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
tree_container_tram2.grid_columnconfigure(0, weight=1)
tree_container_tram2.grid_rowconfigure(0, weight=1)

tree_tram2 = ttk.Treeview(tree_container_tram2, columns=columns_tree, show="headings", style="Treeview")
for col in columns_tree:
    tree_tram2.heading(col, text=col)
tree_tram2.column("Mã Vận Chuyển", width=170, anchor="center")
tree_tram2.column("Phân Loại", width=110, anchor="center")
tree_tram2.column("Giờ Nhận", width=140, anchor="center")
tree_tram2.column("Mã Rổ", width=220, anchor="center")
tree_tram2.pack(side="left", fill="both", expand=True)

v_scrollbar_tram2 = ttk.Scrollbar(tree_container_tram2, orient="vertical", command=tree_tram2.yview)
tree_tram2.configure(yscrollcommand=v_scrollbar_tram2.set)
v_scrollbar_tram2.pack(side="right", fill="y")

button_sync_tram2 = ctk.CTkButton(main_frame_tram2, text="THỬ LẠI / ĐỒNG BỘ CÁC ĐƠN LỖI MẠNG LÊN GOOGLE SHEET", font=('Arial', 13, 'bold'),
                                  height=38, fg_color="#1e3a8a", hover_color="#1e40af", command=lambda: retry_all_pending_baskets(station_type="tram2"))
button_sync_tram2.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# ================= TAB 3: ĐƠN BẤT THƯỜNG =================
tab3 = tabview.tab("ĐƠN BẤT THƯỜNG")
tab3.grid_columnconfigure(0, weight=6)
tab3.grid_columnconfigure(1, weight=4)
tab3.grid_rowconfigure(0, weight=1)

main_frame_dbt = ctk.CTkFrame(tab3, fg_color="transparent")
main_frame_dbt.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=5)

baskets_frame_dbt = ctk.CTkScrollableFrame(tab3, label_text="HỆ THỐNG RỔ ĐƠN BẤT THƯỜNG", label_font=('Arial', 14, 'bold'), label_text_color="#ff9f00")
baskets_frame_dbt.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

main_frame_dbt.grid_columnconfigure(0, weight=1)
main_frame_dbt.grid_columnconfigure(1, weight=2)
main_frame_dbt.grid_rowconfigure(4, weight=1)

ctk.CTkLabel(main_frame_dbt, text="Scan Return TN:", font=('Arial', 14, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_dbt = ctk.CTkEntry(main_frame_dbt, width=280, font=('Arial', 14))
entry_dbt.grid(row=0, column=1, padx=10, pady=10, sticky="w")
entry_dbt.bind("<Return>", save_dbt)

ctk.CTkLabel(main_frame_dbt, text='Loại Hàng:', font=('Arial', 14, 'bold')).grid(row=1, column=0, padx=10, pady=5, sticky="e")
label_loaihang_dbt = ctk.CTkLabel(main_frame_dbt, text='No_data', font=("Arial", 22, 'bold'), text_color="#10b981")
label_loaihang_dbt.grid(row=1, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_dbt, text='Số đơn chưa chốt:', font=('Arial', 13)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
qty_pending_dbt = ctk.CTkLabel(main_frame_dbt, text=count_pending_dbt, font=('Arial', 14, 'bold'), text_color="#f59e0b")
qty_pending_dbt.grid(row=2, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_dbt, text='Số đơn đã chốt rổ:', font=('Arial', 13)).grid(row=3, column=0, padx=10, pady=5, sticky="e")
qty_total_dbt = ctk.CTkLabel(main_frame_dbt, text=count_total_dbt, font=('Arial', 14, 'bold'), text_color="#3b82f6")
qty_total_dbt.grid(row=3, column=1, padx=10, pady=5, sticky="w")

tree_container_dbt = ctk.CTkFrame(main_frame_dbt, fg_color="transparent")
tree_container_dbt.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
tree_container_dbt.grid_columnconfigure(0, weight=1)
tree_container_dbt.grid_rowconfigure(0, weight=1)

tree_dbt = ttk.Treeview(tree_container_dbt, columns=columns_tree, show="headings", style="Treeview")
for col in columns_tree:
    tree_dbt.heading(col, text=col)
tree_dbt.column("Mã Vận Chuyển", width=170, anchor="center")
tree_dbt.column("Phân Loại", width=110, anchor="center")
tree_dbt.column("Giờ Nhận", width=140, anchor="center")
tree_dbt.column("Mã Rổ", width=220, anchor="center")
tree_dbt.pack(side="left", fill="both", expand=True)

v_scrollbar_dbt = ttk.Scrollbar(tree_container_dbt, orient="vertical", command=tree_dbt.yview)
tree_dbt.configure(yscrollcommand=v_scrollbar_dbt.set)
v_scrollbar_dbt.pack(side="right", fill="y")

button_sync_dbt = ctk.CTkButton(main_frame_dbt, text="THỬ LẠI / ĐỒNG BỘ CÁC ĐƠN LỖI MẠNG LÊN GOOGLE SHEET", font=('Arial', 13, 'bold'),
                                height=38, fg_color="#1e3a8a", hover_color="#1e40af", command=lambda: retry_all_pending_baskets(station_type="dbt"))
button_sync_dbt.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# ================= KHỞI TẠO CÁC RỔ PHÂN LOẠI & KHÔI PHỤC DỮ LIỆU (3 TRẠM) =================
for b_name in PREDEFINED_BASKETS:
    create_basket_ui(station_type="tram1", basket_name=b_name)
    create_basket_ui(station_type="tram2", basket_name=b_name)
    create_basket_ui(station_type="dbt", basket_name=b_name)

restore_offline_backup()

status = ctk.CTkLabel(root, text="FD-CB Reverse Logistics System v4.2 - Ready", font=('Arial', 11), text_color="#9ca3af")
status.pack(side=ctk.BOTTOM, fill=ctk.X, pady=2)

# Keep all Tkinter focus changes on the UI thread; submit_basket refocuses
# the active station after each submit.
root.after(250, entry_tram1.focus_set)

if not HAS_QRCODE:
    root.after(500, lambda: messagebox.showwarning(
        "Thiếu thư viện mã QR",
        "Chưa cài thư viện 'qrcode' nên tem in ra sẽ KHÔNG có mã QR!\n\n"
        "Cách khắc phục: Mở CMD gõ: pip install qrcode"
    ))

root.mainloop()
