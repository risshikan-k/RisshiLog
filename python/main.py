import gspread
from oauth2client.service_account import ServiceAccountCredentials
from smartcard.System import readers
from smartcard.util import toHexString
from datetime import datetime
import time
import winsound
from gspread.exceptions import APIError
import config

# ==========================================
# 設定エリア
# ==========================================
SPREADSHEET_KEY = config.SPREADSHEET_KEY
JSON_FILE = config.JSON_FILE
STUDENT_URL_BASE = config.STUDENT_URL_BASE
USER_SHEET_NAME = '名簿'
STATS_SHEET_NAME = '統計'
MONITOR_SHEET_NAME = 'モニター'

# ★メモリキャッシュを完全に廃止しました（毎回シートを確認します）

# ==========================================
# 0. 便利関数
# ==========================================
def normalize_id(val):
    return str(val).strip()

def safe_api_call(func, *args, **kwargs):
    max_retries = 3
    for i in range(max_retries):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            if e.response.status_code == 429:
                wait_time = 10 * (i + 1)
                print(f"⏳ API制限検知。{wait_time}秒待機して再接続します... ({i+1}/{max_retries})")
                time.sleep(wait_time)
            else:
                raise e
    print("❌ API接続に失敗しました。処理をスキップします。")
    return None

# ==========================================
# 1. 基本機能
# ==========================================
def get_workbook():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_KEY)

def get_sheet_safe(workbook, sheet_name, header_row):
    def _get():
        try:
            return workbook.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            sheet = workbook.add_worksheet(title=sheet_name, rows=100, cols=10)
            sheet.append_row(header_row)
            return sheet
    return safe_api_call(_get)

def get_yearly_sheet(workbook):
    now = datetime.now()
    
    # 1月〜3月の場合は「前年度」になるよう計算
    if now.month <= 3:
        year = now.year - 1
    else:
        year = now.year
        
    sheet_name = f"{year}年度"
    return get_sheet_safe(workbook, sheet_name, ['IDm', '名前', '日付', '入室時刻', '退出時刻', '滞在時間(分)'])

# ==========================================
# 2. モニター用シート更新
# ==========================================
def update_monitor_sheet(workbook, name, status, date_str, time_str):
    try:
        sheet = get_sheet_safe(workbook, MONITOR_SHEET_NAME, ['名前', 'ステータス', '日付', '時刻'])
        if sheet:
            safe_api_call(sheet.append_row, [name, status, date_str, time_str])
    except Exception:
        pass


# ==========================================
# 3. 統計データの更新（月ごとの記録対応版）
# ==========================================
def update_statistics(workbook, idm, name, duration_min, date_str):
    # 4月始まりの固定ヘッダー（全29列）
    stats_header = [
        'IDm', '名前', '累計入室日数', '累計時間(分)', '最終入室日',
        '4月日数', '4月時間', '5月日数', '5月時間', '6月日数', '6月時間',
        '7月日数', '7月時間', '8月日数', '8月時間', '9月日数', '9月時間',
        '10月日数', '10月時間', '11月日数', '11月時間', '12月日数', '12月時間',
        '1月日数', '1月時間', '2月日数', '2月時間', '3月日数', '3月時間'
    ]
    stats_sheet = get_sheet_safe(workbook, STATS_SHEET_NAME, stats_header)
    if not stats_sheet: return

    all_rows = safe_api_call(stats_sheet.get_all_values)
    if not all_rows: return

    target_row_index = -1
    current_days = 0
    current_total_time = 0.0
    last_visit_date_str = ""

    if len(all_rows) < 1: all_rows = [stats_header]

    for i in range(1, len(all_rows)):
        row = all_rows[i]
        if len(row) > 0 and normalize_id(row[0]) == idm:
            target_row_index = i + 1
            try: current_days = int(row[2])
            except: current_days = 0
            try: current_total_time = float(row[3])
            except: current_total_time = 0.0
            if len(row) > 4: last_visit_date_str = str(row[4])
            break

    # 今月の列の位置を計算 (4月=6列目, 5月=8列目...)
    try:
        visit_date = datetime.strptime(date_str, '%Y-%m-%d')
        m = visit_date.month
    except:
        m = datetime.now().month

    if m >= 4: offset = m - 4
    else: offset = m + 8

    days_col = 6 + offset * 2
    time_col = 7 + offset * 2

    current_monthly_days = 0
    current_monthly_time = 0.0

    if target_row_index != -1:
        row = all_rows[target_row_index - 1]
        try:
            if len(row) >= days_col and str(row[days_col - 1]).strip() != "":
                current_monthly_days = int(row[days_col - 1])
        except: current_monthly_days = 0
        try:
            if len(row) >= time_col and str(row[time_col - 1]).strip() != "":
                current_monthly_time = float(row[time_col - 1])
        except: current_monthly_time = 0.0

        # 同じ日に複数回入室した場合は日数を増やさない
        if last_visit_date_str != date_str:
            new_days = current_days + 1
            new_monthly_days = current_monthly_days + 1
        else:
            new_days = current_days
            new_monthly_days = current_monthly_days

        new_total_time = current_total_time + duration_min
        new_monthly_time = current_monthly_time + duration_min

        safe_api_call(stats_sheet.update_cell, target_row_index, 2, name)
        safe_api_call(stats_sheet.update_cell, target_row_index, 3, new_days)
        safe_api_call(stats_sheet.update_cell, target_row_index, 4, round(new_total_time, 1))
        safe_api_call(stats_sheet.update_cell, target_row_index, 5, date_str)
        # 月ごとのデータを更新
        safe_api_call(stats_sheet.update_cell, target_row_index, days_col, new_monthly_days)
        safe_api_call(stats_sheet.update_cell, target_row_index, time_col, round(new_monthly_time, 1))

        print(f"   📈 統計更新: 計{new_days}日 ({m}月: {new_monthly_days}日 / {round(new_monthly_time, 1)}分)")
    else:
        new_row = [idm, name, 1, round(duration_min, 1), date_str]
        new_row.extend([""] * 24) # 後ろの列を空欄で埋める
        new_row[days_col - 1] = 1
        new_row[time_col - 1] = round(duration_min, 1)

        safe_api_call(stats_sheet.append_row, new_row)
        print(f"   📈 統計新規作成: {name}")

# ==========================================
# 4. 入退室処理（毎回確認版）
# ==========================================
def handle_tap(idm, workbook):
    safe_idm = normalize_id(idm)
    
    user_name = "未登録(新規)"
    personal_url = STUDENT_URL_BASE + safe_idm
    is_new_user = True
    
    # ★変更: タッチされるたびに必ずスプレッドシートを見に行く
    user_sheet = get_sheet_safe(workbook, USER_SHEET_NAME, ['IDm', '名前', '学年', 'ふりがな', '生徒用URL'])
    if user_sheet:
        user_rows = safe_api_call(user_sheet.get_all_values)
        if user_rows:
            if len(user_rows) > 1:
                for i in range(1, len(user_rows)):
                    row = user_rows[i]
                    if len(row) > 0 and normalize_id(row[0]) == safe_idm:
                        user_name = row[1] if row[1] else "未登録"
                        is_new_user = False # 名簿に存在した
                        
                        # URLが空欄なら追記
                        if len(row) < 5 or row[4] == "":
                            safe_api_call(user_sheet.update_cell, i+1, 5, personal_url)
                        break

    # 名簿に無ければ（消されていれば）新規登録をやり直す
    if is_new_user:
        if user_sheet:
            print(f"🆕 名簿に無いIDを検出（新規登録します）: {safe_idm}")
            safe_api_call(user_sheet.append_row, [safe_idm, '', '', '', personal_url])
            user_name = "未登録(新規)"

    # --- 2. 入退室記録 ---
    sheet = get_yearly_sheet(workbook)
    if not sheet: return

    now = datetime.now()
    date_str = now.strftime('%Y-%m-%d')
    time_str = now.strftime('%H:%M:%S')
    
    print(f"処理中... {user_name} (ID: {safe_idm})")
    
    monthly_rows = safe_api_call(sheet.get_all_values)
    if not monthly_rows: return

    target_row_index = -1
    last_entry_time = None
    
    if len(monthly_rows) > 1:
        for i in range(len(monthly_rows) - 1, 0, -1):
            row = monthly_rows[i]
            if len(row) > 0 and normalize_id(row[0]) == safe_idm:
                exit_time = ""
                if len(row) > 4: exit_time = row[4].strip()
                if exit_time == "":
                    target_row_index = i + 1
                    entry_str = f"{row[2]} {row[3]}"
                    try: last_entry_time = datetime.strptime(entry_str, '%Y-%m-%d %H:%M:%S')
                    except: target_row_index = -1
                break
    
    if target_row_index != -1 and last_entry_time:
        # --- 退出 ---
        duration_min = (now - last_entry_time).total_seconds() / 60
        safe_api_call(sheet.update_cell, target_row_index, 5, time_str)
        safe_api_call(sheet.update_cell, target_row_index, 6, round(duration_min, 1))
        
        print(f"👋 【退出】 {user_name} ({round(duration_min, 1)}分)")
        
        update_statistics(workbook, safe_idm, user_name, duration_min, date_str)
        update_monitor_sheet(workbook, user_name, "退出", date_str, time_str)
        
    else:
        # --- 入室 ---
        new_row = [safe_idm, user_name, date_str, time_str, "", ""]
        safe_api_call(sheet.append_row, new_row)
        print(f"🔔 【入室】 {user_name}")
        
        update_monitor_sheet(workbook, user_name, "入室", date_str, time_str)

def main():
    print("システム起動中...")
    try:
        workbook = get_workbook()
        print("✅ スプレッドシート接続OK")
    except Exception as e:
        print(f"❌ 接続エラー: {e}")
        return

    r = readers()
    if not r:
        print("❌ エラー: リーダーが見つかりません。")
        return

    print("💳 カードリーダー待機中... (Ctrl+Cで終了)")
    connection = r[0].createConnection()
    holding_card_id = None
    
    while True:
        try:
            connection.connect()
            data, sw1, sw2 = connection.transmit([0xFF, 0xCA, 0x00, 0x00, 0x00])
            raw_idm = toHexString(data).replace(" ", "")
            
            if raw_idm == holding_card_id:
                time.sleep(0.5)
                continue
            
            winsound.Beep(2000, 200)
            handle_tap(raw_idm, workbook)
            holding_card_id = raw_idm
            
            time.sleep(1.5)
                
        except Exception:
            holding_card_id = None
        
        time.sleep(0.5)

if __name__ == '__main__':
    main()