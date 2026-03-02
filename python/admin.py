import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from smartcard.System import readers
from smartcard.util import toHexString
import qrcode
from PIL import Image
import time
import io

# ==========================================
# 設定エリア
# ==========================================
JSON_FILE = 'credentials.json'
SPREADSHEET_KEY = '1eYvBli5lOdl991ZSkinb3Bvhrrw3KqGeQLzfedHyITY' # スプレッドシートID
NETLIFY_URL = 'https://astounding-kleicha-4dfef3.netlify.app/' # NetlifyのURL
USER_SHEET_NAME = '名簿'

# ==========================================
# 関数定義
# ==========================================
@st.cache_resource
def get_worksheet():
    """スプレッドシートへの接続（キャッシュして高速化）"""
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
    client = gspread.authorize(creds)
    workbook = client.open_by_key(SPREADSHEET_KEY)
    
    # シートがなければ作る（列: ID, 名前, 学年, URL）
    try:
        sheet = workbook.worksheet(USER_SHEET_NAME)
    except:
        sheet = workbook.add_worksheet(title=USER_SHEET_NAME, rows=100, cols=5)
        sheet.append_row(['IDm', '名前', '学年', '生徒用URL'])
    return sheet

def scan_nfc():
    """PaSoRiからIDを読み取る（見つからなくても待機しない）"""
    try:
        r = readers()
        if not r: return None
        connection = r[0].createConnection()
        connection.connect()
        data, sw1, sw2 = connection.transmit([0xFF, 0xCA, 0x00, 0x00, 0x00])
        return toHexString(data).replace(" ", "")
    except:
        return None

def generate_qr(data):
    """URLからQRコード画像を生成"""
    qr = qrcode.QRCode(box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

# ==========================================
# アプリ画面 (UI)
# ==========================================
st.set_page_config(page_title="タグ管理モニター", layout="centered")

# セッション状態の初期化
if 'last_scanned_id' not in st.session_state:
    st.session_state['last_scanned_id'] = None # 最後に読み取ったID（離しても残る）
if 'is_editing' not in st.session_state:
    st.session_state['is_editing'] = False # 編集モードかどうか

st.title("🏷️ 学習記録タグ管理")

# ------------------------------------------
# モードA: 編集モード (読み取り停止中)
# ------------------------------------------
if st.session_state['is_editing']:
    idm = st.session_state['last_scanned_id']
    st.warning(f"⏸️ 編集中のため読み取りを停止しています (対象ID: {idm})")
    
    # 既存データを取得
    sheet = get_worksheet()
    try:
        records = sheet.get_all_records()
    except:
        records = []
        
    existing_name = ""
    existing_grade = ""
    
    # スプレッドシートから現在の情報を探す
    for row in records:
        if str(row['IDm']) == idm:
            existing_name = row['名前']
            if '学年' in row:
                existing_grade = str(row['学年'])
            break

    # --- 編集フォーム ---
    with st.form("edit_form"):
        st.subheader("📝 登録内容の変更")
        col1, col2 = st.columns(2)
        with col1:
            new_name = st.text_input("名前", value=existing_name, placeholder="氏名を入力")
        with col2:
            new_grade = st.text_input("学年 (任意)", value=existing_grade, placeholder="例: 1年")
            
        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            submit = st.form_submit_button("保存して再開", type="primary")
        with col_btn2:
            cancel = st.form_submit_button("キャンセル")

    # 保存処理
    if submit:
        target_url = f"{NETLIFY_URL}?id={idm}"
        try:
            cell = sheet.find(idm)
            if cell:
                # 更新
                sheet.update_cell(cell.row, 2, new_name)
                sheet.update_cell(cell.row, 3, new_grade)
                sheet.update_cell(cell.row, 4, target_url)
                st.toast(f"✅ 更新しました: {new_name}")
            else:
                # 新規
                sheet.append_row([idm, new_name, new_grade, target_url])
                st.toast(f"🎉 新規登録しました: {new_name}")
            
            # 編集モード終了
            st.session_state['is_editing'] = False
            st.rerun()
            
        except Exception as e:
            st.error(f"保存エラー: {e}")

    # キャンセル処理
    if cancel:
        st.session_state['is_editing'] = False
        st.rerun()

# ------------------------------------------
# モードB: モニタリングモード (常時スキャン)
# ------------------------------------------
else:
    # 1. 常に最新をスキャンする
    scanned_now = scan_nfc()
    
    # もし新しいタグが見つかったら、記憶を更新する
    if scanned_now is not None:
        if st.session_state['last_scanned_id'] != scanned_now:
            st.session_state['last_scanned_id'] = scanned_now
            st.rerun() # 画面を即座に更新

    # 2. 最後に見たIDを表示する（タグを離していてもここが表示される）
    current_id = st.session_state['last_scanned_id']
    
    st.markdown("### 📡 最新のタグ情報")
    
    if current_id:
        # IDがある場合の表示
        st.success(f"ID: {current_id}")
        
        # 既存データを検索
        sheet = get_worksheet()
        try:
            records = sheet.get_all_records()
        except:
            records = []
            
        disp_name = "未登録"
        disp_grade = "-"
        
        for row in records:
            if str(row['IDm']) == current_id:
                disp_name = row['名前']
                if '学年' in row:
                    disp_grade = row['学年']
                break
        
        # 情報表示エリア
        with st.container():
            col1, col2, col3 = st.columns([2, 1, 2])
            with col1:
                st.metric("名前", disp_name)
            with col2:
                st.metric("学年", disp_grade)
            with col3:
                st.write("") # レイアウト調整
                # ★ここを押すと編集モードに入り、スキャンが止まる
                if st.button("✏️ 内容を変更する", type="secondary"):
                    st.session_state['is_editing'] = True
                    st.rerun()

        st.divider()
        st.caption("👇 書き込み用QRコード")
        target_url = f"{NETLIFY_URL}?id={current_id}"
        st.image(generate_qr(target_url), width=120)

    else:
        # まだ一度も読み取っていない場合
        st.info("リーダーにタグをかざしてください...")

    # 3. 自動更新（編集モードでない限りループし続ける）
    time.sleep(1) # 1秒ごとにスキャン
    st.rerun()