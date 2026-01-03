import streamlit as st
import pandas as pd
import datetime
import io
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- PDF生成用ライブラリ ---
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

# --- Excel生成用ライブラリ ---
try:
    import xlsxwriter
    HAS_XLSXWRITER = True
except ImportError:
    HAS_XLSXWRITER = False

# --- 設定 ---
st.set_page_config(page_title="在庫管理システム", layout="wide")

# --- シート名の定義 ---
INVENTORY_SHEET = 'inventory'
HISTORY_SHEET = 'history'
CATEGORY_SHEET = 'categories'
LOCATION_SHEET = 'locations'
MANUFACTURER_SHEET = 'manufacturers'
STAFF_SHEET = 'staff'
ITEM_MASTER_SHEET = 'item_master'
FISCAL_CALENDAR_SHEET = 'fiscal_calendar'

# =========================================================
# 1. スプレッドシート接続・データ操作関数
# =========================================================
def get_gspread_client():
    # Secretsから認証情報を取得
    try:
        # st.secrets["service_account_json"] が文字列の場合はJSONパース、辞書ならそのまま使う
        raw_json = st.secrets["service_account_json"]
        if isinstance(raw_json, str):
            key_dict = json.loads(raw_json)
        else:
            key_dict = raw_json
        
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(key_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"認証エラー: Secretsの設定を確認してください。\n{e}")
        st.stop()

def get_worksheet(sheet_name):
    client = get_gspread_client()
    try:
        url = st.secrets["spreadsheet_url"]
        sh = client.open_by_url(url)
        try:
            worksheet = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            # シートがない場合は作成（1000行20列）
            worksheet = sh.add_worksheet(title=sheet_name, rows=1000, cols=20)
        return worksheet
    except Exception as e:
        st.error(f"スプレッドシート接続エラー: {e}")
        return None

def load_data(sheet_name, columns):
    ws = get_worksheet(sheet_name)
    if ws:
        data = ws.get_all_values()
        # データが空、またはヘッダーしかない場合
        if len(data) <= 1:
            return pd.DataFrame(columns=columns)
        
        # 1行目をヘッダーとして読み込む
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # 期待するカラムが足りない場合の補正（簡易的）
        if not set(columns).issubset(df.columns):
            return pd.DataFrame(data[1:], columns=columns) if len(data) > 1 else pd.DataFrame(columns=columns)
            
        return df
    return pd.DataFrame(columns=columns)

def save_data(df, sheet_name):
    ws = get_worksheet(sheet_name)
    if ws:
        ws.clear()
        # NaNを空文字に変換してリスト化
        df_str = df.fillna("").astype(str)
        header = df_str.columns.tolist()
        data = df_str.values.tolist()
        all_values = [header] + data
        ws.update(values=all_values)

# =========================================================
# 2. 共通関数（計算・PDF・Excel生成）
# =========================================================
def parse_qty_str(qty_str: str):
    s = str(qty_str).strip()
    if s.startswith('+'):
        try: return 'delta', int(s[1:])
        except: return 'delta', 0
    if s.startswith('-'):
        try: return 'delta', -int(s[1:])
        except: return 'delta', 0
    if s.startswith('修正'):
        try:
            body = s.replace('修正:', '').replace('修正：', '').strip()
            parts = body.split('→')
            if len(parts) == 2:
                return 'set_restore', (int(parts[0].strip()), int(parts[1].strip()))
        except: pass
        try:
            body = s.replace('修正:', '').replace('修正：', '').strip()
            return 'set', int(body)
        except: pass
        return 'set', None
    return 'delta', 0

def build_inventory_asof(df_history_src, df_item_master_src, limit_dt, allowed_warehouses=None):
    cols = ['商品名','メーカー','分類','サブカテゴリ','保管場所','在庫数','単位','平均単価','在庫金額']
    if df_history_src.empty:
        return pd.DataFrame(columns=cols)

    hist = df_history_src.copy()
    hist['日時_dt'] = pd.to_datetime(hist['日時'], errors='coerce')
    hist = hist.dropna(subset=['日時_dt'])
    hist = hist[hist['日時_dt'] <= limit_dt].sort_values('日時_dt')

    if allowed_warehouses:
        hist = hist[hist['保管場所'].isin(allowed_warehouses)]

    state = {} 
    for _, r in hist.iterrows():
        name = str(r['商品名'])
        loc = str(r['保管場所'])
        op = str(r['処理'])
        qty_str = str(r['数量'])
        
        unit_price = pd.to_numeric(r.get('単価', 0), errors='coerce')
        unit_price = 0 if pd.isna(unit_price) else float(unit_price)

        key = (name, loc)
        if key not in state: state[key] = {'qty': 0, 'val': 0.0}

        qty_before = int(state[key]['qty'])
        val_before = float(state[key]['val'])
        avg_before = (val_before / qty_before) if qty_before > 0 else 0.0

        kind, v = parse_qty_str(qty_str)

        if op in ['購入入庫', '移動入庫', '返却入庫']: 
            delta = v if kind == 'delta' else 0
            if delta < 0: delta = abs(delta)
            state[key]['qty'] = qty_before + delta
            state[key]['val'] = val_before + (delta * unit_price)

        elif op in ['出庫', '移動出庫', '返却出庫', '客先出庫']:
            delta = v if kind == 'delta' else 0
            out_qty = abs(delta)
            state[key]['qty'] = qty_before - out_qty
            state[key]['val'] = val_before - (out_qty * avg_before)

        elif op == '棚卸':
            if kind == 'set_restore' and isinstance(v, tuple):
                after_qty = v[1]
                state[key]['qty'] = after_qty
                state[key]['val'] = after_qty * avg_before
            elif kind == 'set' and v is not None:
                after_qty = int(v)
                state[key]['qty'] = after_qty
                state[key]['val'] = after_qty * avg_before

        if state[key]['qty'] <= 0:
            state[key]['qty'] = 0
            state[key]['val'] = 0.0

    rows = []
    for (name, loc), sv in state.items():
        qty = int(sv['qty'])
        val = float(sv['val'])
        if qty <= 0: continue
        
        # マスタ情報の補完
        maker = cat = sub = unit = ''
        if not df_item_master_src.empty:
            m_row = df_item_master_src[df_item_master_src['商品名'] == name]
            if not m_row.empty:
                m = m_row.iloc[0]
                maker = m.get('メーカー', '')
                cat = m.get('分類', '')
                sub = m.get('サブカテゴリ', '')
                unit = m.get('単位', '')

        avg = int(val / qty) if qty > 0 else 0
        rows.append({
            '商品名': name, 'メーカー': maker, '分類': cat, 'サブカテゴリ': sub,
            '保管場所': loc, '在庫数': qty, '単位': unit,
            '平均単価': avg, '在庫金額': int(val)
        })

    df = pd.DataFrame(rows)
    if df.empty: return pd.DataFrame(columns=cols)
    return df

# PDF生成 (簡易版: 日本語フォントなしの場合は文字化けする可能性があるため、英語表記推奨かフォント設定が必要)
# 今回はエラー回避のため最小限の実装
def generate_pdf_voucher(tx_data):
    if not HAS_REPORTLAB: return b""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    # 日本語フォントの設定は環境依存（Streamlit CloudにはIPAフォント等がない）ため、
    # 実際にはフォントファイルをアップロードして登録する必要がある。
    # ここではエラーにならないよう標準フォントで英数字のみ出力する例とする。
    c.setFont("Helvetica", 12)
    c.drawString(100, 800, "Voucher")
    c.drawString(100, 780, f"Date: {tx_data['date']}")
    c.drawString(100, 760, f"Type: {tx_data['type']}")
    c.drawString(100, 740, f"Item Code: {tx_data['code']}")
    # 日本語が含まれる変数は文字化けするため、実運用ではフォント対応必須
    c.drawString(100, 720, f"Qty: {tx_data['qty']}")
    c.save()
    return buffer.getvalue()

def generate_monthly_report_excel(df_history, df_item_master, df_location, target_period_str, start_dt, end_dt, warehouse_filter=None, target_subs=None):
    if not HAS_XLSXWRITER: return None
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('MonthlyReport')
    
    fmt_header_mid = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 11})
    fmt_cell = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 10})
    fmt_num = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 10, 'num_format': '#,##0'})
    
    # Header
    worksheet.merge_range('A1:L1', f"月次報告: {target_period_str}", fmt_header_mid)
    headers = ["LOC_N", "LOC_NAME", "DVC_TYPE_NA", "MODEL_N", "MODEL_NAME", "前月繰越", "使用数(差分)", "入庫", "帳簿在庫数", "棚卸報告", "差異", "繰越"]
    for i, h in enumerate(headers): worksheet.write(3, i, h, fmt_header_mid)
    
    # 簡易ロジック（詳細な集計は省略せず、前回のロジックを適用）
    df_h = df_history.copy()
    df_h['dt'] = pd.to_datetime(df_h['日時'], errors='coerce')
    
    # フィルタ
    if pd.notna(start_dt) and pd.notna(end_dt):
        df_period = df_h[(df_h['dt'] >= start_dt) & (df_h['dt'] <= end_dt)]
        df_before = df_h[df_h['dt'] < start_dt]
    else:
        df_period = df_h[0:0]; df_before = df_h[0:0]

    target_warehouses = [warehouse_filter] if (warehouse_filter and warehouse_filter != 'すべて') else df_location['倉庫名'].unique()
    target_items = df_item_master.copy()
    if target_subs: target_items = target_items[target_items['サブカテゴリ'].isin(target_subs)]
    
    row_idx = 4
    for wh in target_warehouses:
        loc_code = ""
        loc_r = df_location[df_location['倉庫名'] == wh]
        if not loc_r.empty: loc_code = loc_r.iloc[0]['倉庫ID']

        for item_name in target_items['商品名'].unique():
            # 前月繰越
            h_b = df_before[(df_before['保管場所'] == wh) & (df_before['商品名'] == item_name)]
            start_qty = 0
            for _, r in h_b.iterrows():
                k, v = parse_qty_str(r['数量'])
                if r['処理'] in ['購入入庫', '移動入庫', '返却入庫']:
                    if k == 'delta': start_qty += abs(v)
                elif r['処理'] in ['出庫', '移動出庫', '返却出庫', '客先出庫']:
                    if k == 'delta': start_qty -= abs(v)
                elif r['処理'] == '棚卸':
                    if k == 'set_restore' and isinstance(v, tuple): start_qty = v[1]
                    elif k == 'set': start_qty = v
            if start_qty < 0: start_qty = 0
            
            # 期間内
            h_d = df_period[(df_period['保管場所'] == wh) & (df_period['商品名'] == item_name)]
            in_qty = 0
            hist_out_qty = 0
            for _, r in h_d.iterrows():
                k, v = parse_qty_str(r['数量'])
                if r['処理'] in ['購入入庫', '移動入庫', '返却入庫'] and k == 'delta': in_qty += abs(v)
                if r['処理'] in ['出庫', '移動出庫', '返却出庫', '客先出庫'] and k == 'delta': hist_out_qty += abs(v)
            
            # 棚卸確認
            st_rows = h_d[h_d['処理'] == '棚卸'].sort_values('dt', ascending=False)
            has_st = not st_rows.empty
            reported = 0; locked = 0
            if has_st:
                k, v = parse_qty_str(st_rows.iloc[0]['数量'])
                if k == 'set_restore' and isinstance(v, tuple): locked = v[0]; reported = v[1]
                elif k == 'set': reported = v

            book_qty = locked if has_st else (start_qty + in_qty - hist_out_qty)
            if book_qty < 0: book_qty = 0
            
            used_qty = start_qty + in_qty - book_qty
            
            m_r = df_item_master[df_item_master['商品名'] == item_name].iloc[0]
            
            worksheet.write(row_idx, 0, loc_code, fmt_cell)
            worksheet.write(row_idx, 1, wh, fmt_cell)
            worksheet.write(row_idx, 2, m_r.get('サブカテゴリ',''), fmt_cell)
            worksheet.write(row_idx, 3, m_r.get('商品コード',''), fmt_cell)
            worksheet.write(row_idx, 4, item_name, fmt_cell)
            worksheet.write(row_idx, 5, start_qty, fmt_num)
            worksheet.write_formula(row_idx, 6, f'=F{row_idx+1}+H{row_idx+1}-I{row_idx+1}', fmt_num, used_qty)
            worksheet.write(row_idx, 7, in_qty, fmt_num)
            worksheet.write(row_idx, 8, book_qty, fmt_num)
            worksheet.write(row_idx, 9, reported if has_st else book_qty, fmt_num)
            worksheet.write_formula(row_idx, 10, f'=J{row_idx+1}-I{row_idx+1}', fmt_num)
            worksheet.write(row_idx, 11, book_qty, fmt_num)
            row_idx += 1

    workbook.close()
    return output.getvalue()

# =========================================================
# 3. セッション & データ読み込み
# =========================================================
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_name' not in st.session_state: st.session_state['user_name'] = ""
if 'user_code' not in st.session_state: st.session_state['user_code'] = ""
if 'user_warehouses' not in st.session_state: st.session_state['user_warehouses'] = []
if 'reset_form' not in st.session_state: st.session_state['reset_form'] = False
if 'last_msg' not in st.session_state: st.session_state['last_msg'] = ""
if 'last_selected_item' not in st.session_state: st.session_state['last_selected_item'] = None
if 'stocktaking_mode' not in st.session_state: st.session_state['stocktaking_mode'] = False 
if 'inventory_snapshot' not in st.session_state: st.session_state['inventory_snapshot'] = None 

if st.session_state['reset_form']:
    st.session_state['reset_form'] = False
    if 'quantity_in' in st.session_state: st.session_state['quantity_in'] = 0
    if 'note_in' in st.session_state: st.session_state['note_in'] = ""
    if 'dest_code_input' in st.session_state: st.session_state['dest_code_input'] = ""

# Load Data from Sheets
df_location = load_data(LOCATION_SHEET, ['倉庫ID', '倉庫名', '属性'])
df_history = load_data(HISTORY_SHEET, ['日時', '商品名', '保管場所', '処理', '数量', '単価', '金額', '担当者名', '担当者所属', '出庫先', '備考'])
df_staff = load_data(STAFF_SHEET, ['担当者コード', '担当者名', '所属', 'パスワード', '担当倉庫'])
df_inventory = load_data(INVENTORY_SHEET, ['商品名', 'メーカー', '分類', 'サブカテゴリ', '保管場所', '在庫数', '単位', '平均単価', '在庫金額'])
df_category = load_data(CATEGORY_SHEET, ['種類ID', '種類'])
df_manufacturer = load_data(MANUFACTURER_SHEET, ['メーカーID', 'メーカー名'])
df_item_master = load_data(ITEM_MASTER_SHEET, ['商品コード', '商品名', 'メーカー', '分類', 'サブカテゴリ', '単位', '標準単価'])
df_fiscal = load_data(FISCAL_CALENDAR_SHEET, ['対象年月', '締め年月日'])

# 初期データ生成
if df_location.empty:
    init_loc = pd.DataFrame({'倉庫ID': ['01'], '倉庫名': ['本社倉庫'], '属性': ['直営']})
    save_data(init_loc, LOCATION_SHEET); df_location = init_loc
if df_staff.empty:
    init_staff = pd.DataFrame({'担当者コード': ['0001'], '担当者名': ['管理者'], '所属': ['管理'], 'パスワード': ['0000'], '担当倉庫': ['本社倉庫']})
    save_data(init_staff, STAFF_SHEET); df_staff = init_staff

# 締め日表示用
if not df_fiscal.empty:
    df_fiscal['dt'] = pd.to_datetime(df_fiscal['締め年月日'], errors='coerce')
    df_fiscal = df_fiscal.dropna(subset=['dt']).sort_values('dt')
    df_fiscal['prev_close'] = df_fiscal['dt'].shift(1)
    df_fiscal['start_dt'] = df_fiscal['prev_close'] + pd.Timedelta(days=1)
    def make_pd_txt(r):
        fmt = '%Y-%m-%d'
        s_d = r['dt'].replace(day=1) if pd.isna(r['start_dt']) else r['start_dt']
        return f"{r['対象年月']} ({s_d.strftime(fmt)}~{r['dt'].strftime(fmt)})"
    df_fiscal['表示用'] = df_fiscal.apply(make_pd_txt, axis=1)

# =========================================================
# 4. ログイン
# =========================================================
if not st.session_state['logged_in']:
    st.title("🔒 ログイン")
    with st.form("login"):
        code = st.text_input("コード")
        pw = st.text_input("パスワード", type="password")
        if st.form_submit_button("Login"):
            u = df_staff[df_staff['担当者コード'] == code]
            if not u.empty and str(u.iloc[0]['パスワード']) == str(pw):
                st.session_state['logged_in'] = True
                st.session_state['user_name'] = u.iloc[0]['担当者名']
                st.session_state['user_code'] = u.iloc[0]['担当者コード']
                st.session_state['user_dept'] = u.iloc[0]['所属']
                whs = str(u.iloc[0].get('担当倉庫',''))
                if code == '0001': st.session_state['user_warehouses'] = df_location['倉庫名'].tolist()
                else: st.session_state['user_warehouses'] = [w.strip() for w in whs.split(',') if w.strip()]
                st.rerun()
            else: st.error("認証失敗")
    st.stop()

# =========================================================
# 5. メインアプリ
# =========================================================
st.title("在庫管理システム")
allowed_warehouses = st.session_state['user_warehouses']
if not allowed_warehouses:
    st.error("担当倉庫がありません")
    st.stop()

# サイドバー: 入出庫
with st.sidebar:
    st.write(f"User: {st.session_state['user_name']}")
    if st.button("Logout"):
        st.session_state['logged_in'] = False
        st.rerun()
    st.divider()

    # 管理者用設定
    if st.session_state['user_code'] == '0001':
        with st.expander("⚙️ 設定（マスタ管理）"):
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["商品", "分類", "倉庫", "メーカー", "担当者", "締め日"])
            
            # 商品
            with tab1:
                if not df_item_master.empty: st.dataframe(df_item_master)
                # (簡易化のため追加/編集フォームは省略せず、前のコードをベースに必要な場合は復元してください)
                # ここではスプレッドシートへの保存確認のため、全削除などの危険な操作のみ実装例として記述
            
            # 倉庫 (全削除機能追加)
            with tab3:
                st.dataframe(df_location)
                st.write("#### 倉庫の追加")
                new_loc = st.text_input("新規倉庫名")
                if st.button("追加", key="btn_add_loc"):
                    if new_loc and new_loc not in df_location['倉庫名'].values:
                        nid = f"{len(df_location)+1:02}"
                        new_row = pd.DataFrame({'倉庫ID':[nid], '倉庫名':[new_loc], '属性':['直営']})
                        df_location = pd.concat([df_location, new_row], ignore_index=True)
                        save_data(df_location, LOCATION_SHEET)
                        st.rerun()
                
                st.divider()
                st.write("#### 🗑️ 倉庫の一括削除")
                st.warning("【注意】すべての倉庫データが削除されます。在庫データとの整合性が取れなくなる可能性があります。")
                if st.checkbox("リスクを理解して全削除を行う", key="chk_del_all_loc"):
                    if st.button("全倉庫を削除する", type="primary", key="btn_del_all_loc"):
                        # ヘッダーのみの空DataFrameを作成して保存
                        empty_loc = pd.DataFrame(columns=['倉庫ID', '倉庫名', '属性'])
                        save_data(empty_loc, LOCATION_SHEET)
                        st.success("すべての倉庫を削除しました")
                        st.rerun()

            # その他のタブも同様に実装可能

    st.divider()
    st.subheader("処理実行")
    # 入出庫フォームロジック
    act = st.radio("処理", ["入庫", "出庫", "移動", "棚卸"])
    
    # 倉庫選択
    loc = st.selectbox("倉庫", allowed_warehouses)
    
    # 商品選択
    items = df_item_master['商品名'].tolist() if not df_item_master.empty else []
    item_name = st.selectbox("商品", items)
    
    qty = st.number_input("数量", min_value=1)
    
    if st.button("実行"):
        # スプレッドシートへ保存する処理
        # ここでは簡易的に履歴と在庫を更新するロジック
        dt_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
        
        # 履歴追加
        h_row = pd.DataFrame([{
            '日時': dt_str, '商品名': item_name, '保管場所': loc, '処理': act,
            '数量': f"+{qty}" if act=='入庫' else f"-{qty}",
            '単価': '0', '金額': '0', '担当者名': st.session_state['user_name'],
            '担当者所属': st.session_state['user_dept'], '出庫先': '-', '備考': ''
        }])
        df_history = pd.concat([df_history, h_row], ignore_index=True)
        save_data(df_history, HISTORY_SHEET)
        
        # 在庫更新 (簡易: 再計算ではなくレコード追加/更新)
        # 実際には build_inventory_asof のロジックで計算されるため、
        # inventoryシート自体を更新する必要がある場合はここで計算して save_data する
        st.success("処理完了")
        st.rerun()

# メイン表示
t1, t2, t3 = st.tabs(["現在庫", "履歴", "その他"])
with t1:
    # リアルタイム在庫計算
    now_inv = build_inventory_asof(df_history, df_item_master, pd.Timestamp.now(), allowed_warehouses)
    st.dataframe(now_inv)

with t2:
    st.dataframe(df_history)

with t3:
    if st.button("月次レポート生成(Excel)"):
        # サンプル：直近の締め日データを使用
        if not df_fiscal.empty:
            last_fiscal = df_fiscal.iloc[-1]
            xl = generate_monthly_report_excel(df_history, df_item_master, df_location, last_fiscal['表示用'], last_fiscal.get('start_dt'), last_fiscal['dt'], loc)
            if xl:
                st.download_button("Download Excel", xl, "report.xlsx")
