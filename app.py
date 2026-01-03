import streamlit as st
import pandas as pd
import os
import datetime
import io

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

# --- ファイル名の設定 ---
INVENTORY_FILE = 'inventory.csv'
HISTORY_FILE = 'history.csv'
CATEGORY_FILE = 'categories.csv'
LOCATION_FILE = 'locations.csv'
MANUFACTURER_FILE = 'manufacturers.csv'
STAFF_FILE = 'staff.csv'
ITEM_MASTER_FILE = 'item_master.csv'
FISCAL_CALENDAR_FILE = 'fiscal_calendar.csv'

st.set_page_config(page_title="在庫管理システム", layout="wide")

# =========================================================
# 1. セッション状態の初期化
# =========================================================
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_name' not in st.session_state:
    st.session_state['user_name'] = ""
if 'user_code' not in st.session_state:
    st.session_state['user_code'] = ""
if 'user_dept' not in st.session_state:
    st.session_state['user_dept'] = ""
if 'user_warehouses' not in st.session_state:
    st.session_state['user_warehouses'] = []
if 'latest_voucher' not in st.session_state:
    st.session_state['latest_voucher'] = None
if 'latest_voucher_name' not in st.session_state:
    st.session_state['latest_voucher_name'] = ""

# フォームクリア＆メッセージ保持用フラグ
if 'reset_form' not in st.session_state:
    st.session_state['reset_form'] = False
if 'last_msg' not in st.session_state:
    st.session_state['last_msg'] = ""

# 商品変更検知用
if 'last_selected_item' not in st.session_state:
    st.session_state['last_selected_item'] = None

# 棚卸モード管理
if 'stocktaking_mode' not in st.session_state:
    st.session_state['stocktaking_mode'] = False 
if 'inventory_snapshot' not in st.session_state:
    st.session_state['inventory_snapshot'] = None 

# 【重要】ウィジェット描画前に値をリセットする
if st.session_state['reset_form']:
    st.session_state['reset_form'] = False
    if 'dest_code_input' in st.session_state:
        st.session_state['dest_code_input'] = ""
    if 'note_in' in st.session_state:
        st.session_state['note_in'] = ""
    if 'quantity_in' in st.session_state:
        st.session_state['quantity_in'] = 0

# =========================================================
# 2. データ読み込み・保存関数
# =========================================================
def load_data(file, columns):
    if os.path.exists(file):
        df = pd.read_csv(file, dtype=str)
        return df.fillna("")
    return pd.DataFrame(columns=columns)

def save_data(df, file):
    df.to_csv(file, index=False)

# =========================================================
# 3. 共通関数（計算・PDF・Excel生成）
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
        if key not in state:
            state[key] = {'qty': 0, 'val': 0.0}

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

        master_row = df_item_master_src[df_item_master_src['商品名'] == name]
        if not master_row.empty:
            m = master_row.iloc[0]
            maker = m.get('メーカー', '')
            cat = m.get('分類', '')
            sub = m.get('サブカテゴリ', '')
            unit = m.get('単位', '')
        else:
            maker = cat = sub = unit = ''

        avg = int(val / qty) if qty > 0 else 0
        rows.append({
            '商品名': name, 'メーカー': maker, '分類': cat, 'サブカテゴリ': sub,
            '保管場所': loc, '在庫数': qty, '単位': unit,
            '平均単価': avg, '在庫金額': int(val)
        })

    df = pd.DataFrame(rows)
    if df.empty: return pd.DataFrame(columns=cols)
    return df

# PDF生成関数
def generate_pdf_voucher(tx_data):
    if not HAS_REPORTLAB:
        raise ImportError("reportlabがインストールされていません。pip install reportlab を実行してください。")

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4 
    
    font_name = "Helvetica"
    font_candidates = [
        "C:\\Windows\\Fonts\\msgothic.ttc",
        "C:\\Windows\\Fonts\\meiryo.ttc",
        "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
        "/System/Library/Fonts/Heiti.ttc",
        "/Library/Fonts/Arial Unicode.ttf"
    ]
    for fpath in font_candidates:
        if os.path.exists(fpath):
            try:
                pdfmetrics.registerFont(TTFont('JpFont', fpath))
                font_name = 'JpFont'
                break
            except: continue

    def draw_half(y_offset, title, is_receipt=False):
        c.setFont(font_name, 18)
        c.drawCentredString(width / 2, y_offset + 370, title)
        
        c.setFont(font_name, 10)
        c.drawString(400, y_offset + 390, f"発行日: {tx_data['date']}")
        c.drawString(400, y_offset + 375, f"担当者: {tx_data['operator']}")
        
        c.setFont(font_name, 12)
        c.drawString(50, y_offset + 345, f"納入先: {tx_data['to']}  御中")
        
        c.setFont(font_name, 10)
        from_val = str(tx_data['from'])
        if from_val == "nan" or from_val == "-" or not from_val:
            from_disp = "(記録なし)"
        else:
            from_disp = from_val
        c.drawString(50, y_offset + 325, f"出荷元: {from_disp}")
        
        table_top = y_offset + 290
        c.setLineWidth(1)
        c.line(40, table_top, 550, table_top)
        c.drawString(50, table_top - 15, "商品コード")
        c.drawString(130, table_top - 15, "商品名 / 規格")
        c.drawString(380, table_top - 15, "数量")
        c.drawString(480, table_top - 15, "単位")
        c.line(40, table_top - 25, 550, table_top - 25)
        
        c.drawString(50, table_top - 45, str(tx_data['code']))
        c.drawString(130, table_top - 45, f"{tx_data['name']}")
        c.setFont(font_name, 8)
        c.drawString(130, table_top - 58, f"({tx_data['maker']} / {tx_data['sub']})")
        c.setFont(font_name, 10)
        c.drawString(380, table_top - 45, str(tx_data['qty']))
        c.drawString(480, table_top - 45, str(tx_data['unit']))
        c.line(40, table_top - 70, 550, table_top - 70)

        note_str = str(tx_data.get('note', ''))
        c.drawString(50, table_top - 90, f"備考: {note_str}")

        if is_receipt:
            c.drawString(380, y_offset + 50, "受領印:")
            c.line(420, y_offset + 50, 530, y_offset + 50)
            c.drawString(40, y_offset + 50, "上記正に受領いたしました。")

    title_upper = "納 品 伝 票"
    if tx_data.get('type') == 'return':
        title_upper = "返 却 伝 票"
    elif tx_data.get('type') == 'transfer':
        title_upper = "出 庫 伝 票"

    draw_half(height / 2, title_upper, is_receipt=False)
    c.setDash(1, 2)
    c.line(20, height / 2, width - 20, height / 2)
    c.setDash([]) 
    draw_half(0, "受 領 書", is_receipt=True)
    c.showPage()
    c.save()
    return buffer.getvalue()

# --- Excel (xlsxwriter) 月次報告生成関数（複雑レイアウト版） ---
def generate_monthly_report_excel(df_history, df_item_master, df_location, target_period_str, start_dt, end_dt, warehouse_filter=None, target_subs=None):
    if not HAS_XLSXWRITER:
        return None
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('MonthlyReport')
    
    # スタイル定義
    fmt_header_top = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9', 'font_size': 10})
    fmt_header_mid = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 11})
    fmt_header_sub = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 9, 'text_wrap': True})
    fmt_cell = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 10})
    fmt_num = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 10, 'num_format': '#,##0'})
    fmt_cell_calc = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 10, 'num_format': '#,##0', 'bg_color': '#FFFFCC'})
    fmt_gray = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 10, 'bg_color': '#808080'}) 
    
    # ヘッダー構築
    title_ym = target_period_str.split(' ')[0] if ' ' in target_period_str else target_period_str
    
    worksheet.merge_range('A1:L1', f"月次報告: {title_ym}", fmt_header_mid)
    
    s_str = start_dt.strftime('%Y/%m/%d') if pd.notna(start_dt) else ""
    e_str = end_dt.strftime('%Y/%m/%d') if pd.notna(end_dt) else ""
    period_str = f"{s_str}～{e_str}"
    worksheet.merge_range('A2:L2', period_str, fmt_header_mid)
    
    worksheet.merge_range('A3:E3', '商品情報', fmt_header_top)
    worksheet.merge_range('F3:I3', '帳簿', fmt_header_top)
    worksheet.merge_range('J3:L3', 'SMS在庫', fmt_header_top)
    worksheet.merge_range('M3:O3', '業者報告', fmt_header_top)
    worksheet.write('P3', 'DCBEE', fmt_header_top)
    worksheet.write('Q3', '', fmt_header_top)
    
    headers = [
        "LOC_N", "LOC_NAME", "DVC_TYPE_NA", "MODEL_N", "MODEL_NAME",
        "前月繰越", "使用数(差分)", "入庫", "帳簿在庫数",
        "新品", "中古", "その他",
        "出庫報告", "棚卸報告", "簿在庫との差",
        "工事件数", "繰越"
    ]
    
    for col_num, header in enumerate(headers):
        worksheet.write(3, col_num, header, fmt_header_sub)
        
    worksheet.set_column('A:A', 8)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 10)
    worksheet.set_column('D:D', 10)
    worksheet.set_column('E:E', 25)
    worksheet.set_column('F:Q', 9)
    
    # データ集計
    df_h = df_history.copy()
    df_h['dt'] = pd.to_datetime(df_h['日時'], errors='coerce')
    
    # 日付フィルタ
    if pd.notna(start_dt) and pd.notna(end_dt):
        mask_period = (df_h['dt'] >= start_dt) & (df_h['dt'] <= end_dt)
        df_period = df_h[mask_period]
        # 前月繰越用
        mask_before = (df_h['dt'] < start_dt)
        df_before = df_h[mask_before]
    else:
        df_period = df_h[0:0] 
        df_before = df_h[0:0]

    if warehouse_filter and warehouse_filter != 'すべて':
        target_warehouses = [warehouse_filter]
    elif allowed_warehouses:
        target_warehouses = allowed_warehouses
    else:
        target_warehouses = []
    
    target_items_df = df_item_master.copy()
    if target_subs:
        target_items_df = target_items_df[target_items_df['サブカテゴリ'].isin(target_subs)]
    all_items = target_items_df['商品名'].unique() if not target_items_df.empty else []
    
    row_idx = 4
    
    for wh in target_warehouses:
        loc_code = ""
        if not df_location.empty:
             loc_row = df_location[df_location['倉庫名'] == wh]
             if not loc_row.empty:
                 loc_code = loc_row.iloc[0]['倉庫ID']
        
        for item_name in all_items:
            m_row = df_item_master[df_item_master['商品名'] == item_name].iloc[0]
            dvc_type = m_row.get('サブカテゴリ', '') 
            model_code = m_row.get('商品コード', '')
            model_name = item_name
            
            # --- 前月繰越計算 ---
            h_b = df_before[(df_before['保管場所'] == wh) & (df_before['商品名'] == item_name)].sort_values('dt')
            start_qty = 0
            for _, r in h_b.iterrows():
                op = r['処理']
                k, v = parse_qty_str(r['数量'])
                if op in ['購入入庫', '移動入庫', '返却入庫']:
                    if k == 'delta': start_qty += abs(v)
                elif op in ['出庫', '客先出庫', '移動出庫', '返却出庫']:
                    if k == 'delta': start_qty -= abs(v)
                elif op == '棚卸':
                    if k == 'set_restore' and isinstance(v, tuple): start_qty = v[1]
                    elif k == 'set' and v is not None: start_qty = v
                if start_qty < 0: start_qty = 0
            
            # --- 期間内集計 (入庫・出庫) ---
            h_data = df_period[(df_period['保管場所'] == wh) & (df_period['商品名'] == item_name)]
            
            in_qty = 0
            hist_out_qty = 0 # 履歴上の出庫数（計算チェック用）
            
            for _, r in h_data.iterrows():
                op = r['処理']
                k, v = parse_qty_str(r['数量'])
                
                if op in ['出庫', '客先出庫', '移動出庫', '返却出庫']:
                    if k == 'delta': hist_out_qty += abs(v)
                elif op in ['購入入庫', '移動入庫', '返却入庫']:
                    if k == 'delta': in_qty += abs(v)
            
            # 棚卸情報 (最新のみ)
            stocktake_rows = h_data[h_data['処理'] == '棚卸'].sort_values('dt', ascending=False)
            reported_qty = 0
            locked_qty_val = 0
            
            has_stocktake = False
            if not stocktake_rows.empty:
                has_stocktake = True
                latest_st = stocktake_rows.iloc[0]
                k, v = parse_qty_str(latest_st['数量'])
                if k == 'set_restore' and isinstance(v, tuple):
                    locked_qty_val = v[0] # ロック数
                    reported_qty = v[1]   # 実棚
                elif k == 'set' and v is not None:
                    reported_qty = v
                    locked_qty_val = 0
            
            # --- 帳簿在庫 (Book Qty) と 使用数 (Used Qty) の決定 ---
            if has_stocktake:
                book_qty = locked_qty_val
            else:
                # 棚卸がない場合: 期首 + 入庫 - 履歴出庫
                book_qty = start_qty + in_qty - hist_out_qty
                if book_qty < 0: book_qty = 0
            
            # 使用数(差分) = 前月繰越 + 入庫 - 帳簿在庫 (逆算)
            used_qty = start_qty + in_qty - book_qty
            
            # 書き込み
            worksheet.write(row_idx, 0, loc_code, fmt_cell)
            worksheet.write(row_idx, 1, wh, fmt_cell)
            worksheet.write(row_idx, 2, dvc_type, fmt_cell) 
            worksheet.write(row_idx, 3, model_code, fmt_cell)
            worksheet.write(row_idx, 4, model_name, fmt_cell)
            
            worksheet.write(row_idx, 5, start_qty, fmt_num) # F: 前月繰越
            
            # G: 使用数 (計算式も埋め込む: F+H-I)
            idx = row_idx + 1
            worksheet.write_formula(row_idx, 6, f'=F{idx}+H{idx}-I{idx}', fmt_num, used_qty) 
            
            worksheet.write(row_idx, 7, in_qty, fmt_num)    # H: 入庫
            worksheet.write(row_idx, 8, book_qty, fmt_num)  # I: 帳簿在庫
            
            # SMS在庫 (新品/中古 判定)
            if '(再)' in model_name or '中古' in model_name:
                worksheet.write(row_idx, 9, '', fmt_gray) 
                worksheet.write(row_idx, 10, 0, fmt_num)   
            else:
                worksheet.write(row_idx, 9, 0, fmt_num)   
                worksheet.write(row_idx, 10, '', fmt_gray) 
            worksheet.write(row_idx, 11, 0, fmt_num) 
            
            # M: 出庫報告 (使用数と同じにする)
            worksheet.write(row_idx, 12, used_qty, fmt_num) 
            
            # N: 棚卸報告
            if has_stocktake:
                worksheet.write(row_idx, 13, reported_qty, fmt_num)
            else:
                worksheet.write(row_idx, 13, book_qty, fmt_num)
            
            # O: 差異 (計算式: 棚卸報告(N) - 帳簿在庫(I))
            worksheet.write_formula(row_idx, 14, f'=N{idx}-I{idx}', fmt_num)
            
            worksheet.write(row_idx, 15, '', fmt_cell)
            
            # Q列: 繰越 -> 帳簿在庫と同じ値を表示
            worksheet.write(row_idx, 16, book_qty, fmt_num)
            
            row_idx += 1

    workbook.close()
    return output.getvalue()


# =========================================================
# 4. データ読み込み & 前処理
# =========================================================
df_location = load_data(LOCATION_FILE, ['倉庫ID', '倉庫名', '属性'])
df_history = load_data(HISTORY_FILE, ['日時', '商品名', '保管場所', '処理', '数量', '単価', '金額', '担当者名', '担当者所属', '出庫先', '備考'])
df_staff = load_data(STAFF_FILE, ['担当者コード', '担当者名', '所属', 'パスワード', '担当倉庫'])
df_inventory = load_data(INVENTORY_FILE, ['商品名', 'メーカー', '分類', 'サブカテゴリ', '保管場所', '在庫数', '単位', '平均単価', '在庫金額'])
df_category = load_data(CATEGORY_FILE, ['種類ID', '種類'])
df_manufacturer = load_data(MANUFACTURER_FILE, ['メーカーID', 'メーカー名'])
df_item_master = load_data(ITEM_MASTER_FILE, ['商品コード', '商品名', 'メーカー', '分類', 'サブカテゴリ', '単位', '標準単価'])
df_fiscal = load_data(FISCAL_CALENDAR_FILE, ['対象年月', '締め年月日'])

# --- 締め日データのクリーンアップ ---
if not df_fiscal.empty:
    df_fiscal = df_fiscal[['対象年月', '締め年月日']]

# --- 各種自動修復処理 ---
if not df_staff.empty and '担当倉庫' not in df_staff.columns:
    df_staff['担当倉庫'] = ""
    save_data(df_staff, STAFF_FILE)
    df_staff = load_data(STAFF_FILE, ['担当者コード', '担当者名', '所属', 'パスワード', '担当倉庫'])

if not df_history.empty:
    changed = False
    if '出庫先' not in df_history.columns:
        df_history['出庫先'] = "-"
        changed = True
    if '備考' not in df_history.columns:
        df_history['備考'] = ""
        changed = True
    if changed:
        save_data(df_history, HISTORY_FILE)
        df_history = load_data(HISTORY_FILE, ['日時', '商品名', '保管場所', '処理', '数量', '単価', '金額', '担当者名', '担当者所属', '出庫先', '備考'])

if not df_location.empty:
    loc_changed = False
    if '倉庫ID' not in df_location.columns:
        ids = [f"{i+1:02}" for i in range(len(df_location))]
        df_location.insert(0, '倉庫ID', ids)
        loc_changed = True
    if '属性' not in df_location.columns:
        df_location['属性'] = '直営'
        loc_changed = True
    
    if loc_changed:
        save_data(df_location, LOCATION_FILE)
        df_location = load_data(LOCATION_FILE, ['倉庫ID', '倉庫名', '属性'])

if not df_manufacturer.empty and 'メーカーID' not in df_manufacturer.columns:
    ids = [f"{i+1:02}" for i in range(len(df_manufacturer))]
    df_manufacturer.insert(0, 'メーカーID', ids)
    save_data(df_manufacturer, MANUFACTURER_FILE)
    df_manufacturer = load_data(MANUFACTURER_FILE, ['メーカーID', 'メーカー名'])

if not df_category.empty and '種類ID' not in df_category.columns:
    ids = [f"{i+1:02}" for i in range(len(df_category))]
    df_category.insert(0, '種類ID', ids)
    save_data(df_category, CATEGORY_FILE)
    df_category = load_data(CATEGORY_FILE, ['種類ID', '種類'])

if not df_item_master.empty and '商品コード' not in df_item_master.columns:
    codes = [f"{i+1:04}" for i in range(len(df_item_master))]
    df_item_master.insert(0, '商品コード', codes)
    save_data(df_item_master, ITEM_MASTER_FILE)
    df_item_master = load_data(ITEM_MASTER_FILE, ['商品コード', '商品名', 'メーカー', '分類', 'サブカテゴリ', '単位', '標準単価'])

# --- 初期データ生成 ---
if df_location.empty:
    default_locs = pd.DataFrame({
        '倉庫ID': ['01', '02', '99'], 
        '倉庫名': ['高木2ビル１F倉庫', '本社倉庫', '返却倉庫'],
        '属性': ['直営', '直営', '直営']
    })
    save_data(default_locs, LOCATION_FILE)
    df_location = load_data(LOCATION_FILE, ['倉庫ID', '倉庫名', '属性'])

if df_staff.empty:
    all_locs_str = ",".join(df_location['倉庫名'].tolist())
    df_staff = pd.DataFrame({
        '担当者コード': ['0001'], '担当者名': ['管理者'], '所属': ['システム管理'], 
        'パスワード': ['0000'], '担当倉庫': [all_locs_str]
    })
    save_data(df_staff, STAFF_FILE)
    df_staff = load_data(STAFF_FILE, ['担当者コード', '担当者名', '所属', 'パスワード', '担当倉庫'])

if df_category.empty:
    default_cats = pd.DataFrame({'種類ID': ['01', '02', '03'], '種類': ['PC', 'モニター', 'ケーブル']})
    save_data(default_cats, CATEGORY_FILE)
    df_category = load_data(CATEGORY_FILE, ['種類ID', '種類'])

if df_manufacturer.empty:
    default_makers = pd.DataFrame({'メーカーID': ['01', '02'], 'メーカー名': ['自社', 'メーカーA']})
    save_data(default_makers, MANUFACTURER_FILE)
    df_manufacturer = load_data(MANUFACTURER_FILE, ['メーカーID', 'メーカー名'])

# 締め日期間の計算処理
if not df_fiscal.empty:
    df_fiscal['dt'] = pd.to_datetime(df_fiscal['締め年月日'], errors='coerce')
    df_fiscal = df_fiscal.dropna(subset=['dt']).sort_values('dt')
    df_fiscal['prev_close'] = df_fiscal['dt'].shift(1)
    df_fiscal['start_dt'] = df_fiscal['prev_close'] + pd.Timedelta(days=1)
    
    def make_period_text(row):
        date_fmt = '%Y-%m-%d'
        end_str = row['dt'].strftime(date_fmt)
        if pd.isna(row['start_dt']):
            start_str = row['dt'].replace(day=1).strftime(date_fmt)
        else:
            start_str = row['start_dt'].strftime(date_fmt)
        return f"{row['対象年月']} 期間{start_str}～{end_str}"

    df_fiscal['表示用'] = df_fiscal.apply(make_period_text, axis=1)

# =========================================================
# 5. ログイン画面
# =========================================================
if not st.session_state['logged_in']:
    st.title("🔒 ログイン")
    st.caption("担当者コード（4桁）とパスワードを入力してください")

    with st.form("login_form"):
        login_code = st.text_input("担当者コード", placeholder="例: 0001")
        login_pass = st.text_input("パスワード", type="password")
        submit_login = st.form_submit_button("ログイン")

        if submit_login:
            user_row = df_staff[df_staff['担当者コード'] == login_code]
            if not user_row.empty:
                user_data = user_row.iloc[0]
                if str(user_data['パスワード']) == str(login_pass):
                    st.session_state['logged_in'] = True
                    st.session_state['user_name'] = user_data['担当者名']
                    st.session_state['user_code'] = user_data['担当者コード']
                    st.session_state['user_dept'] = user_data['所属']

                    if user_data['担当者コード'] == '0001':
                        st.session_state['user_warehouses'] = df_location['倉庫名'].tolist()
                    else:
                        warehouses_str = ""
                        if '担当倉庫' in user_data and pd.notna(user_data['担当倉庫']):
                            warehouses_str = str(user_data['担当倉庫'])
                        
                        if warehouses_str == '' or warehouses_str == 'nan':
                            st.session_state['user_warehouses'] = []
                        else:
                            st.session_state['user_warehouses'] = warehouses_str.split(',')

                    st.success("ログインしました")
                    st.rerun()
                else:
                    st.error("パスワードが違います")
            else:
                st.error("担当者コードが見つかりません")
    st.stop()

# =========================================================
# メインアプリ
# =========================================================
allowed_warehouses = st.session_state['user_warehouses']

st.title('🚚 在庫管理システム')

with st.sidebar:
    st.info(f"ログイン中:\n{st.session_state['user_name']} ({st.session_state['user_code']})")
    
    # 管理者用：操作モード切替
    if st.session_state['user_code'] == '0001':
        st.subheader("👑 管理者メニュー")
        admin_mode = st.radio("操作モード切替", ["全倉庫 (管理者)", "倉庫指定 (担当者)"], horizontal=True, key="admin_mode_select")

        if admin_mode == "倉庫指定 (担当者)":
            all_locs = df_location['倉庫名'].tolist()
            selected_sim_locs = st.multiselect("操作する倉庫を選択", all_locs, default=[], key="admin_sim_locs")
            
            if selected_sim_locs:
                allowed_warehouses = selected_sim_locs
                st.caption(f"選択中: {', '.join(selected_sim_locs)}")
            else:
                st.warning("倉庫を選択してください")
                allowed_warehouses = []
        else:
            allowed_warehouses = df_location['倉庫名'].tolist()
        
        st.divider()

    # 伝票ダウンロード
    if st.session_state.get('latest_voucher') is not None:
        st.download_button(
            label="📥 直近の伝票DL (PDF)",
            data=st.session_state['latest_voucher'],
            file_name=st.session_state['latest_voucher_name'],
            mime="application/pdf",
            key="btn_download_voucher"
        )
        st.divider()

    # 再読込後の成功メッセージ表示
    if st.session_state['last_msg']:
        st.success(st.session_state['last_msg'])
        st.session_state['last_msg'] = "" 

    if not allowed_warehouses:
        st.error("操作可能な倉庫がありません。")

    if st.button("ログアウト"):
        st.session_state['logged_in'] = False
        st.session_state['user_name'] = ""
        st.session_state['user_code'] = ""
        st.session_state['user_dept'] = ""
        st.session_state['user_warehouses'] = []
        st.session_state['latest_voucher'] = None
        st.session_state['reset_form'] = False
        st.session_state['last_msg'] = ""
        # 棚卸モードも解除
        st.session_state['stocktaking_mode'] = False
        st.session_state['inventory_snapshot'] = None
        st.rerun()
    st.divider()

# =========================================================
# 設定（マスタ管理）
# =========================================================
with st.sidebar.expander("⚙️ 設定（マスタ管理）"):
    if st.session_state['user_code'] != '0001':
        st.error("⛔️ この機能は管理者（コード: 0001）のみ使用可能です。")
    else:
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["商品", "分類", "倉庫", "メーカー", "担当者", "📅 締め日"])

        # 1. 商品
        with tab1:
            item_mode = st.radio("操作モード", ["🆕 新規登録", "✏️ 編集・削除"], horizontal=True, key="item_mode_select")
            st.divider()

            if item_mode == "🆕 新規登録":
                st.write("#### 商品の新規登録")
                m_name = st.text_input("商品名", key="m_name_in")
                
                col_m1, col_m2 = st.columns(2)
                with col_m1:
                    m_maker_opts = df_manufacturer['メーカー名'].tolist() if not df_manufacturer.empty else []
                    m_maker = st.selectbox("メーカー", m_maker_opts, key="m_maker_in") if m_maker_opts else ""
                    m_cat = st.radio("分類", ['機器', '部材', 'その他'], key="m_cat_in")
                with col_m2:
                    m_sub_cat = st.selectbox("機器種類", df_category['種類'], key="m_sub_in") if (m_cat == '機器' and not df_category.empty) else '-'
                    m_unit = st.selectbox("単位", ['個', '本', '枚', 'kg', 'セット'], key="m_unit_in")
                    m_price = st.number_input("標準単価", min_value=0, step=10, key="m_price_in")

                maker_id = "00"
                if m_maker:
                    m_row = df_manufacturer[df_manufacturer['メーカー名'] == m_maker]
                    if not m_row.empty: maker_id = m_row.iloc[0]['メーカーID']
                
                cat_id = "00"
                if m_cat == '機器' and m_sub_cat != '-':
                    c_row = df_category[df_category['種類'] == m_sub_cat]
                    if not c_row.empty: cat_id = c_row.iloc[0]['種類ID']
                
                code_prefix = maker_id + cat_id
                next_seq = 1
                if not df_item_master.empty:
                    existing_codes = df_item_master[df_item_master['商品コード'].str.startswith(code_prefix, na=False)]['商品コード']
                    if not existing_codes.empty:
                        max_seq = 0
                        for c in existing_codes:
                            try:
                                suffix = c[len(code_prefix):]
                                if suffix.isdigit(): max_seq = max(max_seq, int(suffix))
                            except: pass
                        next_seq = max_seq + 1
                
                auto_code = f"{code_prefix}{next_seq:03}"
                st.info(f"🆕 発行予定コード: **{auto_code}**")

                if st.button("追加（自動コード発行）", key="btn_item_add"):
                    if m_name and m_maker:
                        if m_name in df_item_master['商品名'].values:
                            st.error("その商品名は既に登録されています")
                        elif auto_code in df_item_master['商品コード'].values:
                            st.error("コード生成エラー: 既に存在するコードです")
                        else:
                            new_item = pd.DataFrame({
                                '商品コード': [auto_code],
                                '商品名': [m_name], 'メーカー': [m_maker], '分類': [m_cat],
                                'サブカテゴリ': [m_sub_cat], '単位': [m_unit], '標準単価': [m_price]
                            })
                            df_item_master = pd.concat([df_item_master, new_item], ignore_index=True)
                            save_data(df_item_master, ITEM_MASTER_FILE)
                            st.success(f"「{m_name} (コード:{auto_code})」を登録しました")
                            st.rerun()
                    else:
                        st.error("商品名、メーカーは必須です")

            else:
                st.write("#### 既存商品の編集・削除")
                if not df_item_master.empty:
                    del_opts = [f"{r['商品コード']} : {r['サブカテゴリ']} : {r['商品名']}" for i, r in df_item_master.iterrows()]
                    edit_target_str = st.selectbox("編集する商品を選択", del_opts, key="sel_item_edit")
                    
                    if edit_target_str:
                        target_code = edit_target_str.split(':')[0].strip()
                        if target_code in df_item_master['商品コード'].values:
                            target_row = df_item_master[df_item_master['商品コード'] == target_code].iloc[0]

                            with st.form("edit_item_form"):
                                st.write(f"**商品コード: {target_code}** (変更不可)")
                                e_name = st.text_input("商品名", value=target_row['商品名'])
                                col_e1, col_e2 = st.columns(2)
                                with col_e1:
                                    m_maker_opts = df_manufacturer['メーカー名'].tolist()
                                    curr_maker_idx = m_maker_opts.index(target_row['メーカー']) if target_row['メーカー'] in m_maker_opts else 0
                                    e_maker = st.selectbox("メーカー", m_maker_opts, index=curr_maker_idx)
                                    cat_opts = ['機器', '部材', 'その他']
                                    curr_cat_idx = cat_opts.index(target_row['分類']) if target_row['分類'] in cat_opts else 0
                                    e_cat = st.radio("分類", cat_opts, index=curr_cat_idx)
                                with col_e2:
                                    sub_opts = df_category['種類'].tolist()
                                    curr_sub_idx = sub_opts.index(target_row['サブカテゴリ']) if target_row['サブカテゴリ'] in sub_opts else 0
                                    e_sub_cat = st.selectbox("機器種類", sub_opts, index=curr_sub_idx)
                                    if e_cat != '機器': e_sub_cat = '-'
                                    unit_opts = ['個', '本', '枚', 'kg', 'セット']
                                    curr_unit_idx = unit_opts.index(target_row['単位']) if target_row['単位'] in unit_opts else 0
                                    e_unit = st.selectbox("単位", unit_opts, index=curr_unit_idx)
                                    val_price = int(float(target_row['標準単価'])) if target_row['標準単価'] else 0
                                    e_price = st.number_input("標準単価", min_value=0, step=10, value=val_price)

                                col_btn1, col_btn2 = st.columns(2)
                                with col_btn1:
                                    update = st.form_submit_button("情報を更新")
                                with col_btn2:
                                    delete = st.form_submit_button("この商品を削除", type="primary")
                                
                                if update:
                                    if e_name and e_maker:
                                        idx = df_item_master[df_item_master['商品コード'] == target_code].index
                                        df_item_master.loc[idx, '商品名'] = e_name
                                        df_item_master.loc[idx, 'メーカー'] = e_maker
                                        df_item_master.loc[idx, '分類'] = e_cat
                                        df_item_master.loc[idx, 'サブカテゴリ'] = e_sub_cat
                                        df_item_master.loc[idx, '単位'] = e_unit
                                        df_item_master.loc[idx, '標準単価'] = e_price
                                        save_data(df_item_master, ITEM_MASTER_FILE)
                                        st.success("商品情報を更新しました")
                                        st.rerun()
                                    else:
                                        st.error("商品名とメーカーは必須です")

                                if delete:
                                    df_item_master = df_item_master[df_item_master['商品コード'] != target_code]
                                    save_data(df_item_master, ITEM_MASTER_FILE)
                                    st.success(f"商品コード {target_code} を削除しました")
                                    st.rerun()
                else:
                    st.info("登録されている商品がありません")

        # 2. 分類
        with tab2:
            col_cat1, col_cat2 = st.columns(2)
            with col_cat1:
                new_cat_id = st.text_input("種類ID (2桁)", key="cat_id_in", max_chars=2, placeholder="例: 01")
            with col_cat2:
                new_cat = st.text_input("種類名", key="cat_in")
            
            if st.button("追加", key="cat_btn"):
                if new_cat_id and new_cat:
                    if new_cat_id not in df_category['種類ID'].values and new_cat not in df_category['種類'].values:
                        df_category = pd.concat([df_category, pd.DataFrame({'種類ID': [new_cat_id], '種類': [new_cat]})], ignore_index=True)
                        save_data(df_category, CATEGORY_FILE)
                        st.rerun()
                    else:
                        st.error("IDまたは種類名が重複しています")
                else:
                    st.error("IDと種類名を入力してください")

            if not df_category.empty:
                st.divider()
                cat_opts = [f"{row['種類ID']}: {row['種類']}" for idx, row in df_category.iterrows()]
                del_cat_str = st.selectbox("削除種類", cat_opts, key="sel_cat_del")
                
                if st.button("削除実行", key="btn_cat_del", disabled=not st.checkbox("確認", key="chk_cat")):
                    if del_cat_str:
                        target_id = del_cat_str.split(':')[0]
                        df_category = df_category[df_category['種類ID'] != target_id]
                        save_data(df_category, CATEGORY_FILE)
                        st.rerun()

        # 3. 倉庫
        with tab3:
            loc_mode = st.radio("操作モード", ["🆕 新規登録", "✏️ 編集・削除"], horizontal=True, key="loc_mode_select")
            st.divider()

            if loc_mode == "🆕 新規登録":
                st.caption("倉庫登録")
                col_loc1, col_loc2 = st.columns(2)
                with col_loc1:
                    new_loc_name = st.text_input("倉庫名", key="loc_in")
                with col_loc2:
                    new_loc_type = st.radio("属性", ['直営', '委託先'], horizontal=True, key="loc_type_in")
                
                next_loc_num = 1
                if not df_location.empty:
                    current_ids = []
                    for x in df_location['倉庫ID']:
                        try: current_ids.append(int(x))
                        except: pass
                    if current_ids:
                        next_loc_num = max(current_ids) + 1
                
                auto_loc_id = f"{next_loc_num:02}"
                st.info(f"🆕 次に発行される倉庫ID: **{auto_loc_id}**")

                if st.button("追加（自動ID発行）", key="loc_btn"):
                    if new_loc_name:
                        if new_loc_name not in df_location['倉庫名'].values:
                            new_row = pd.DataFrame({'倉庫ID': [auto_loc_id], '倉庫名': [new_loc_name], '属性': [new_loc_type]})
                            df_location = pd.concat([df_location, new_row], ignore_index=True)
                            save_data(df_location, LOCATION_FILE)
                            st.rerun()
                        else:
                            st.error("倉庫名が重複しています")
                    else:
                        st.error("倉庫名を入力してください")
            
            else:
                if not df_location.empty:
                    loc_opts = [f"{row['倉庫ID']}: {row['倉庫名']} ({row['属性']})" for idx, row in df_location.iterrows()]
                    edit_target_str = st.selectbox("編集/削除する倉庫を選択", loc_opts, key="sel_loc_edit")
                    
                    if edit_target_str:
                        target_id = edit_target_str.split(':')[0].strip()
                        if target_id in df_location['倉庫ID'].values:
                            target_row = df_location[df_location['倉庫ID'] == target_id].iloc[0]
                            
                            with st.form("edit_loc_form"):
                                st.write(f"**倉庫ID: {target_id}**")
                                edit_loc_name = st.text_input("倉庫名", value=target_row['倉庫名'])
                                
                                current_type = target_row['属性']
                                type_opts = ['直営', '委託先']
                                try:
                                    idx_type = type_opts.index(current_type)
                                except:
                                    idx_type = 0
                                edit_loc_type = st.radio("属性", type_opts, index=idx_type, horizontal=True)
                                
                                col_btn1, col_btn2 = st.columns(2)
                                with col_btn1:
                                    update = st.form_submit_button("情報を更新")
                                with col_btn2:
                                    delete = st.form_submit_button("この倉庫を削除", type="primary")
                                
                                if update:
                                    if edit_loc_name:
                                        other_locs = df_location[df_location['倉庫ID'] != target_id]['倉庫名'].values
                                        if edit_loc_name in other_locs:
                                            st.error("その倉庫名は既に使われています")
                                        else:
                                            idx = df_location[df_location['倉庫ID'] == target_id].index
                                            df_location.loc[idx, '倉庫名'] = edit_loc_name
                                            df_location.loc[idx, '属性'] = edit_loc_type
                                            save_data(df_location, LOCATION_FILE)
                                            st.success("倉庫情報を更新しました")
                                            st.rerun()
                                    else:
                                        st.error("倉庫名は必須です")
                                
                                if delete:
                                    df_location = df_location[df_location['倉庫ID'] != target_id]
                                    save_data(df_location, LOCATION_FILE)
                                    st.success(f"倉庫ID {target_id} を削除しました")
                                    st.rerun()
                        else:
                            st.warning("指定された倉庫データが見つかりません。")
                else:
                    st.info("登録されている倉庫がありません")

        # 4. メーカー
        with tab4:
            col_mak1, col_mak2 = st.columns(2)
            with col_mak1:
                new_maker_id = st.text_input("メーカーID (2桁)", key="maker_id_in", max_chars=2, placeholder="例: 01")
            with col_mak2:
                new_maker_name = st.text_input("メーカー名", key="maker_in")
            
            if st.button("追加", key="maker_btn"):
                if new_maker_id and new_maker_name:
                    if new_maker_id not in df_manufacturer['メーカーID'].values and new_maker_name not in df_manufacturer['メーカー名'].values:
                        df_manufacturer = pd.concat([df_manufacturer, pd.DataFrame({'メーカーID': [new_maker_id], 'メーカー名': [new_maker_name]})], ignore_index=True)
                        save_data(df_manufacturer, MANUFACTURER_FILE)
                        st.rerun()
                    else:
                        st.error("IDまたはメーカー名が重複しています")
                else:
                    st.error("IDとメーカー名を入力してください")

            if not df_manufacturer.empty:
                st.divider()
                maker_opts = [f"{row['メーカーID']}: {row['メーカー名']}" for idx, row in df_manufacturer.iterrows()]
                del_maker_str = st.selectbox("削除メーカー", maker_opts, key="sel_maker_del")
                
                if st.button("削除実行", key="btn_maker_del", disabled=not st.checkbox("確認", key="chk_maker")):
                    if del_maker_str:
                        target_id = del_maker_str.split(':')[0]
                        df_manufacturer = df_manufacturer[df_manufacturer['メーカーID'] != target_id]
                        save_data(df_manufacturer, MANUFACTURER_FILE)
                        st.rerun()

        # 5. 担当者
        with tab5:
            st.write("### ➕ 新規登録")
            col_s1, col_s2 = st.columns(2)
            with col_s1: new_staff_name = st.text_input("氏名", key="staff_name_in")
            with col_s2: new_staff_dept = st.text_input("所属", key="staff_dept_in") 
            new_staff_pass = st.text_input("パスワード設定", key="staff_pass_in", type="password")
            
            all_warehouses = df_location['倉庫名'].tolist() if not df_location.empty else []
            new_staff_locs = st.multiselect("担当する倉庫", all_warehouses, key="staff_locs_in")

            next_code = f"{len(df_staff) + 1:04}"
            st.info(f"次に発行されるコード: {next_code}")

            if st.button("担当者を追加（コード発番）", key="staff_btn"):
                if new_staff_name and new_staff_dept and new_staff_pass and new_staff_locs:
                    locs_str = ",".join(new_staff_locs)
                    new_staff_row = pd.DataFrame({
                        '担当者コード': [next_code], '担当者名': [new_staff_name], 
                        '所属': [new_staff_dept], 'パスワード': [str(new_staff_pass)],
                        '担当倉庫': [locs_str]
                    })
                    df_staff = pd.concat([df_staff, new_staff_row], ignore_index=True)
                    save_data(df_staff, STAFF_FILE)
                    st.success(f"登録完了！コード「{next_code}」")
                    st.rerun()
                else:
                    st.error("担当倉庫を含む全ての項目を入力してください")

            st.divider()
            st.write("### ✏️ 登録情報の編集・削除")
            if not df_staff.empty:
                staff_display_list = [f"{row['担当者コード']}: {row['担当者名']}" for index, row in df_staff.iterrows()]
                if staff_display_list:
                    edit_target_str = st.selectbox("編集/削除する担当者を選択", staff_display_list, key="sel_staff_edit")
                    
                    if edit_target_str:
                        target_code_edit = edit_target_str.split(':')[0].strip()
                        if target_code_edit in df_staff['担当者コード'].values:
                            target_row = df_staff[df_staff['担当者コード'] == target_code_edit].iloc[0]
                            
                            with st.form(key="edit_staff_form"):
                                col_e1, col_e2 = st.columns(2)
                                with col_e1: edit_name = st.text_input("氏名", value=target_row['担当者名'])
                                with col_e2: edit_dept = st.text_input("所属", value=target_row['所属'])
                                edit_pass = st.text_input("パスワード", value=str(target_row['パスワード']), type="password")
                                
                                current_locs_str = str(target_row.get('担当倉庫', '') or '')
                                default_locs = current_locs_str.split(',') if current_locs_str and current_locs_str != 'nan' else []
                                default_locs = [x for x in default_locs if x in all_warehouses]
                                edit_locs = st.multiselect("担当倉庫", all_warehouses, default=default_locs)
                                
                                col_btn1, col_btn2 = st.columns(2)
                                with col_btn1: update_btn = st.form_submit_button("情報を更新")
                                with col_btn2: delete_btn = st.form_submit_button("この担当者を削除", type="primary")

                                if update_btn:
                                    if edit_name and edit_dept and edit_pass and edit_locs:
                                        df_staff.loc[df_staff['担当者コード'] == target_code_edit, '担当者名'] = edit_name
                                        df_staff.loc[df_staff['担当者コード'] == target_code_edit, '所属'] = edit_dept
                                        df_staff.loc[df_staff['担当者コード'] == target_code_edit, 'パスワード'] = str(edit_pass)
                                        df_staff.loc[df_staff['担当者コード'] == target_code_edit, '担当倉庫'] = ",".join(edit_locs)
                                        save_data(df_staff, STAFF_FILE)
                                        st.success(f"{edit_name} さんの情報を更新しました")
                                        st.rerun()
                                    else: st.error("全ての項目を入力してください")
                                
                                if delete_btn:
                                    if target_code_edit == '0001': st.error("管理者は削除できません")
                                    else:
                                        df_staff = df_staff[df_staff['担当者コード'] != target_code_edit]
                                        save_data(df_staff, STAFF_FILE)
                                        st.success("削除しました")
                                        st.rerun()
                        else:
                            st.warning("選択された担当者コードのデータが見つかりません。")
                else:
                    st.info("表示可能な担当者がいません。")
            else:
                st.info("登録されている担当者がいません。")

        # 6. 締め日 (Tab6)
        with tab6:
            st.caption("月ごとの締め日を登録します")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                this_month = datetime.date.today().strftime("%Y-%m")
                fiscal_ym = st.text_input("対象年月 (YYYY-MM)", value=this_month, key="fiscal_ym_in")
            with col_f2:
                fiscal_date = st.date_input("締め年月日", datetime.date.today(), key="fiscal_date_in")
            
            if st.button("締め日を登録/更新", key="btn_fiscal_add"):
                if fiscal_ym and fiscal_date:
                    fiscal_date_str = fiscal_date.strftime('%Y-%m-%d')
                    
                    if not df_fiscal.empty:
                        # 対象年月の行を一旦削除
                        df_fiscal = df_fiscal[df_fiscal['対象年月'] != fiscal_ym]
                    
                    # 新しい行を作成（表示用列などは含めない）
                    new_fiscal_row = pd.DataFrame({'対象年月': [fiscal_ym], '締め年月日': [fiscal_date_str]})
                    
                    # 既存データから必要な2列だけ抽出して結合
                    if not df_fiscal.empty:
                         df_fiscal_clean = df_fiscal[['対象年月', '締め年月日']]
                    else:
                         df_fiscal_clean = pd.DataFrame(columns=['対象年月', '締め年月日'])
                    
                    df_fiscal = pd.concat([df_fiscal_clean, new_fiscal_row], ignore_index=True)
                    df_fiscal = df_fiscal.sort_values('対象年月')
                    
                    # 【重要】保存時は必要な2列のみに絞る
                    save_data(df_fiscal[['対象年月', '締め年月日']], FISCAL_CALENDAR_FILE)
                    
                    st.success(f"{fiscal_ym} の締め日を {fiscal_date_str} に設定しました")
                    st.rerun()
            
            if not df_fiscal.empty:
                st.divider()
                # 画面表示には「表示用」列も含める
                if '表示用' in df_fiscal.columns:
                    st.dataframe(df_fiscal[['対象年月', '締め年月日', '表示用']], use_container_width=True)
                else:
                    st.dataframe(df_fiscal[['対象年月', '締め年月日']], use_container_width=True)
                
                del_fiscal_ym = st.selectbox("削除する年月", df_fiscal['対象年月'], key="sel_fiscal_del")
                if st.button("締め日設定を削除", key="btn_fiscal_del", disabled=not st.checkbox("確認", key="chk_fiscal")):
                    df_fiscal = df_fiscal[df_fiscal['対象年月'] != del_fiscal_ym]
                    # 【重要】保存時は必要な2列のみにする
                    save_data(df_fiscal[['対象年月', '締め年月日']], FISCAL_CALENDAR_FILE)
                    st.rerun()

        # 管理者用：データリセット機能
        if st.session_state['user_code'] == '0001':
            with st.sidebar.expander("🔥 データ初期化"):
                st.error("【注意】\n在庫データと入出庫履歴を\n全て消去します。\n復元はできません。")
                if st.checkbox("理解してリセットする", key="ack_reset"):
                    if st.button("実行 (全データ消去)", type="primary"):
                        cols_inv = ['商品名', 'メーカー', '分類', 'サブカテゴリ', '保管場所', '在庫数', '単位', '平均単価', '在庫金額']
                        df_empty_inv = pd.DataFrame(columns=cols_inv)
                        save_data(df_empty_inv, INVENTORY_FILE)

                        cols_hist = ['日時', '商品名', '保管場所', '処理', '数量', '単価', '金額', '担当者名', '担当者所属', '出庫先', '備考']
                        df_empty_hist = pd.DataFrame(columns=cols_hist)
                        save_data(df_empty_hist, HISTORY_FILE)

                        st.success("データをリセットしました")
                        st.rerun()

st.sidebar.divider()

# =========================================================
# 入出庫フォーム (サイドバー)
# =========================================================
st.sidebar.header('🚚 入出庫処理')

if not allowed_warehouses:
    st.sidebar.warning("担当倉庫がないため、操作できません。")
else:
    action_opts = ['客先出庫', '機器返却', '棚卸']
    if st.session_state['user_code'] == '0001':
        current_mode = st.session_state.get('admin_mode_select', '全倉庫 (管理者)')
        if current_mode == '全倉庫 (管理者)':
             action_opts = ['購入入庫', '在庫移動', '客先出庫', '棚卸']

    action_type = st.sidebar.radio("処理区分", action_opts, help="購入：外部からの仕入れ（単価必須）")

    if df_item_master.empty:
        st.sidebar.warning("商品マスタがありません。")
        st.stop()

    default_index = 0
    target_default_name = "高木2ビル１F倉庫"
    if action_type == '機器返却': target_default_name = "返却倉庫"
    elif st.session_state['user_code'] == '0001' and action_type == '購入入庫': target_default_name = "高木2ビル１F倉庫"
    
    if target_default_name in allowed_warehouses:
        default_index = allowed_warehouses.index(target_default_name)
    
    current_opts = allowed_warehouses
    if action_type == '購入入庫':
        if not df_location.empty:
            direct_locs = df_location[df_location['属性'] == '直営']['倉庫名'].tolist()
            current_opts = [x for x in allowed_warehouses if x in direct_locs]
            if not current_opts:
                st.sidebar.error("購入入庫ができる直営倉庫の権限がありません。")
                st.stop()
    
    if target_default_name in current_opts:
        default_index = current_opts.index(target_default_name)
    else: default_index = 0

    location = st.sidebar.selectbox('対象倉庫（保管場所）', current_opts, index=default_index)
    
    if action_type == '棚卸':
        if not st.session_state['stocktaking_mode']:
            st.info("棚卸を開始すると、現在の在庫数が「ロック（帳簿在庫）」されます。")
            if st.button("棚卸を開始する"):
                st.session_state['inventory_snapshot'] = df_inventory.copy()
                st.session_state['stocktaking_mode'] = True
                st.rerun()
        else:
            st.warning("現在、棚卸モード中です。実数を入力してください。")
            if st.button("棚卸を終了する（ロック解除）", type="primary"):
                st.session_state['stocktaking_mode'] = False
                st.session_state['inventory_snapshot'] = None
                st.rerun()

    all_classes = ['すべて'] + sorted(df_item_master['分類'].dropna().unique().tolist())
    filter_class = st.sidebar.selectbox("分類絞り込み", all_classes, key="sb_class")
    df_step1 = df_item_master.copy()
    if filter_class != 'すべて': df_step1 = df_step1[df_step1['分類'] == filter_class]

    all_subs = ['すべて'] + sorted(df_step1['サブカテゴリ'].dropna().unique().tolist())
    filter_sub = st.sidebar.selectbox("機器種類絞り込み", all_subs, key="sb_sub")
    df_step2 = df_step1.copy()
    if filter_sub != 'すべて': df_step2 = df_step2[df_step2['サブカテゴリ'] == filter_sub]

    all_makers = ['すべて'] + sorted(df_step2['メーカー'].dropna().unique().tolist())
    filter_maker = st.sidebar.selectbox("メーカー絞り込み", all_makers, key="sb_maker")
    df_filtered_items = df_step2.copy()
    if filter_maker != 'すべて': df_filtered_items = df_filtered_items[df_filtered_items['メーカー'] == filter_maker]

    if action_type == '購入入庫': pass
    elif action_type == '機器返却':
        current_inv = df_inventory.copy()
        current_inv['在庫数'] = pd.to_numeric(current_inv['在庫数'], errors='coerce')
        exist_items = current_inv[current_inv['在庫数'] > 0]['商品名'].unique()
        clean_items = [x for x in exist_items if '(返却品)' not in str(x)]
        df_filtered_items = df_filtered_items[df_filtered_items['商品名'].isin(clean_items)]
    else:
        current_inv = df_inventory[df_inventory['保管場所'] == location].copy()
        current_inv['在庫数'] = pd.to_numeric(current_inv['在庫数'], errors='coerce')
        exist_items_in_loc = current_inv[current_inv['在庫数'] > 0]['商品名'].unique()
        df_filtered_items = df_filtered_items[df_filtered_items['商品名'].isin(exist_items_in_loc)]

    if df_filtered_items.empty:
        st.sidebar.warning("選択可能な商品がありません")
        st.stop()
    else:
        item_list = df_filtered_items['商品名'].tolist()
        selected_item_name = st.sidebar.selectbox('商品を選択', item_list, index=None, placeholder="商品を選択してください")

    if selected_item_name != st.session_state['last_selected_item']:
        st.session_state['last_selected_item'] = selected_item_name
        if 'quantity_in' in st.session_state and st.session_state['quantity_in'] != 0:
            st.session_state['quantity_in'] = 0
            st.rerun()

    if selected_item_name:
        item_data = df_item_master[df_item_master['商品名'] == selected_item_name].iloc[0]
        st.sidebar.info(f"{item_data['メーカー']} / {item_data['分類']} / {item_data['単位']}")

        if action_type == '棚卸' and st.session_state['stocktaking_mode']:
            snapshot = st.session_state['inventory_snapshot']
            target_row = snapshot[(snapshot['商品名'] == selected_item_name) & (snapshot['保管場所'] == location)]
            locked_qty = 0
            if not target_row.empty: locked_qty = int(pd.to_numeric(target_row.iloc[0]['在庫数'], errors='coerce') or 0)
            st.sidebar.markdown(f"**帳簿在庫（ロック数）:** {locked_qty}")
            st.sidebar.markdown("👇 **実数（数えた数）**を入力してください")

        location_from = None
        location_to = None
        destination_code = "-"

        if action_type == '在庫移動':
            location_from = location
            st.sidebar.markdown(f"**移動元:** {location_from}")
            location_to = st.sidebar.selectbox('移動先倉庫', allowed_warehouses, key='loc_to')
            if location_from and location_to and location_from == location_to:
                st.sidebar.warning("移動元と移動先は別の倉庫にしてください")
        
        qty_label = '数量'
        if action_type == '棚卸': qty_label = '棚卸数 (実数)'
        quantity = st.sidebar.number_input(qty_label, min_value=0, step=1, key='quantity_in')

        input_price = 0
        if action_type == '購入入庫':
            default_price = int(pd.to_numeric(item_data['標準単価'], errors='coerce') or 0)
            input_price = st.sidebar.number_input('購入単価 (円)', min_value=0, step=10, value=default_price)
        elif action_type == '在庫移動': st.sidebar.caption("※移動元の在庫評価額（平均単価）で移動します")
        elif action_type == '機器返却':
            st.sidebar.caption("※返却された機器を【在庫として入庫】します")
            direct_opts = df_location[df_location['属性'] == '直営']['倉庫名'].tolist()
            destination_code = st.sidebar.selectbox("返却先（直営倉庫）", ["-"] + direct_opts)
        elif action_type == '客先出庫':
            st.sidebar.caption("※出庫・棚卸時は、現在の平均単価が自動適用されます")
            destination_code = st.sidebar.text_input("出庫先コード (7桁)", max_chars=7, help="数字7桁で入力してください", key="dest_code_input")
        else: st.sidebar.caption("※出庫・棚卸時は、現在の平均単価が自動適用されます")

        input_note = ""
        if action_type != '購入入庫':
            lbl_note = "備考 (返却理由など)"
            if action_type == '機器返却': lbl_note += " ※必須"
            input_note = st.sidebar.text_input(lbl_note, key="note_in")

        st.sidebar.caption("処理日時（デフォルトは現在）")
        col_date, col_time = st.sidebar.columns(2)
        with col_date: input_date = st.date_input("日付", datetime.date.today())
        with col_time: input_time = st.time_input("時間", datetime.datetime.now().time())

        if st.sidebar.button('処理を実行'):
            operator_name = st.session_state['user_name']
            operator_dept = st.session_state['user_dept']
            name = selected_item_name
            manufacturer = item_data['メーカー']
            category = item_data['分類']
            sub_category = item_data['サブカテゴリ']
            unit = item_data['単位']
            
            if action_type == '機器返却':
                if not input_note:
                    st.sidebar.error("返却理由（備考）は必須です")
                    st.stop()
                if destination_code == '-':
                    st.sidebar.error("返却先を選択してください")
                    st.stop()

            input_dt = datetime.datetime.combine(input_date, input_time)
            record_str = input_dt.strftime('%Y-%m-%d %H:%M')
            record_str_filename = input_dt.strftime('%Y%m%d_%H%M%S')

            # --- 処理ロジック ---
            # 1. 在庫移動
            if action_type == '在庫移動':
                if not location_from or not location_to or location_from == location_to:
                    st.sidebar.error("移動元と移動先を正しく選択してください")
                    st.stop()
                row_from = df_inventory[(df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location_from)]
                qty_from = 0
                val_from = 0.0
                if not row_from.empty:
                    qty_from = int(pd.to_numeric(row_from.iloc[0]['在庫数'], errors='coerce') or 0)
                    val_from = float(pd.to_numeric(row_from.iloc[0]['在庫金額'], errors='coerce') or 0)
                if qty_from < quantity:
                    st.sidebar.error(f"移動元の在庫が不足しています（在庫: {qty_from}）")
                    st.stop()
                avg_price_from = (val_from / qty_from) if qty_from > 0 else 0
                move_value = quantity * avg_price_from
                
                # 元
                new_qty_from = qty_from - quantity
                new_val_from = val_from - move_value
                new_avg_from = int(new_val_from / new_qty_from) if new_qty_from > 0 else 0
                # 先
                row_to = df_inventory[(df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location_to)]
                qty_to = 0
                val_to = 0.0
                if not row_to.empty:
                    qty_to = int(pd.to_numeric(row_to.iloc[0]['在庫数'], errors='coerce') or 0)
                    val_to = float(pd.to_numeric(row_to.iloc[0]['在庫金額'], errors='coerce') or 0)
                new_qty_to = qty_to + quantity
                new_val_to = val_to + move_value
                new_avg_to = int(new_val_to / new_qty_to) if new_qty_to > 0 else 0
                
                # 更新
                df_inventory = df_inventory[~((df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location_from))]
                if new_qty_from > 0:
                    df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                        '商品名': name, 'メーカー': manufacturer, '分類': category, 'サブカテゴリ': sub_category,
                        '保管場所': location_from, '在庫数': new_qty_from, '単位': unit,
                        '平均単価': new_avg_from, '在庫金額': int(new_val_from)
                    }])], ignore_index=True)
                else:
                    df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                        '商品名': name, 'メーカー': manufacturer, '分類': category, 'サブカテゴリ': sub_category,
                        '保管場所': location_from, '在庫数': 0, '単位': unit,
                        '平均単価': 0, '在庫金額': 0
                    }])], ignore_index=True)
                
                df_inventory = df_inventory[~((df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location_to))]
                df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                    '商品名': name, 'メーカー': manufacturer, '分類': category, 'サブカテゴリ': sub_category,
                    '保管場所': location_to, '在庫数': new_qty_to, '単位': unit,
                    '平均単価': new_avg_to, '在庫金額': int(new_val_to)
                }])], ignore_index=True)

                hist_out = pd.DataFrame([{
                    '日時': record_str, '商品名': name, '保管場所': location_from, '処理': '移動出庫',
                    '数量': f"-{quantity}", '単価': int(avg_price_from), '金額': int(move_value),
                    '担当者名': operator_name, '担当者所属': operator_dept, '出庫先': location_to, '備考': input_note
                }])
                hist_in = pd.DataFrame([{
                    '日時': record_str, '商品名': name, '保管場所': location_to, '処理': '移動入庫',
                    '数量': f"+{quantity}", '単価': int(avg_price_from), '金額': int(move_value),
                    '担当者名': operator_name, '担当者所属': operator_dept, '出庫先': location_from, '備考': input_note
                }])
                df_history = pd.concat([df_history, hist_out, hist_in], ignore_index=True)
                
                save_data(df_inventory, INVENTORY_FILE)
                save_data(df_history, HISTORY_FILE)

                # PDF
                tx_data = {'type': 'transfer', 'date': record_str, 'operator': operator_name, 'from': location_from, 'to': location_to, 'code': item_data.get('商品コード', '-'), 'name': name, 'maker': manufacturer, 'sub': sub_category, 'qty': quantity, 'unit': unit, 'note': input_note}
                try:
                    st.session_state['latest_voucher'] = generate_pdf_voucher(tx_data)
                    st.session_state['latest_voucher_name'] = f"transfer_{record_str_filename}.pdf"
                except Exception as e: st.sidebar.error(f"PDF生成エラー: {e}")
                
                st.session_state['last_msg'] = f"移動完了: {location_from} -> {location_to}"
                st.session_state['reset_form'] = True
                st.rerun()

            # 2. 機器返却
            elif action_type == '機器返却':
                return_name = f"{name} (返却品)"
                if return_name not in df_item_master['商品名'].values:
                    new_master_row = item_data.copy()
                    new_master_row['商品名'] = return_name
                    new_master_row['商品コード'] = f"{item_data['商品コード']}-R"
                    df_item_master = pd.concat([df_item_master, pd.DataFrame([new_master_row])], ignore_index=True)
                    save_data(df_item_master, ITEM_MASTER_FILE)
                
                current_src = df_inventory[(df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location)]
                src_qty = 0; src_val = 0.0
                if not current_src.empty:
                    src_qty = int(pd.to_numeric(current_src.iloc[0]['在庫数'], errors='coerce') or 0)
                    src_val = float(pd.to_numeric(current_src.iloc[0]['在庫金額'], errors='coerce') or 0)
                
                if src_qty < quantity:
                    st.sidebar.error(f"在庫不足です (現在: {src_qty})")
                    st.stop()
                
                avg_price = (src_val / src_qty) if src_qty > 0 else 0
                move_val = quantity * avg_price
                
                # 元を減らす
                df_inventory = df_inventory[~((df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location))]
                new_src_qty = src_qty - quantity
                new_src_val = src_val - move_val
                new_src_avg = int(new_src_val / new_src_qty) if new_src_qty > 0 else 0
                df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                    '商品名': name, 'メーカー': manufacturer, '分類': category, 'サブカテゴリ': sub_category,
                    '保管場所': location, '在庫数': new_src_qty, '単位': unit,
                    '平均単価': new_src_avg, '在庫金額': int(new_src_val)
                }])], ignore_index=True)

                # 先を増やす (返却品名)
                target_loc = destination_code
                current_dest = df_inventory[(df_inventory['商品名'] == return_name) & (df_inventory['保管場所'] == target_loc)]
                dest_qty = 0; dest_val = 0.0
                if not current_dest.empty:
                    dest_qty = int(pd.to_numeric(current_dest.iloc[0]['在庫数'], errors='coerce') or 0)
                    dest_val = float(pd.to_numeric(current_dest.iloc[0]['在庫金額'], errors='coerce') or 0)
                
                new_dest_qty = dest_qty + quantity
                new_dest_val = dest_val + move_val
                new_dest_avg = int(new_dest_val / new_dest_qty) if new_dest_qty > 0 else 0
                
                df_inventory = df_inventory[~((df_inventory['商品名'] == return_name) & (df_inventory['保管場所'] == target_loc))]
                df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                    '商品名': return_name, 'メーカー': manufacturer, '分類': category, 'サブカテゴリ': sub_category,
                    '保管場所': target_loc, '在庫数': new_dest_qty, '単位': unit,
                    '平均単価': new_dest_avg, '在庫金額': int(new_dest_val)
                }])], ignore_index=True)

                save_data(df_inventory, INVENTORY_FILE)
                
                hist_out = pd.DataFrame([{
                    '日時': record_str, '商品名': name, '保管場所': location, '処理': '返却出庫',
                    '数量': f"-{quantity}", '単価': int(avg_price), '金額': int(move_val),
                    '担当者名': operator_name, '担当者所属': operator_dept, '出庫先': target_loc, '備考': input_note
                }])
                hist_in = pd.DataFrame([{
                    '日時': record_str, '商品名': return_name, '保管場所': target_loc, '処理': '返却入庫',
                    '数量': f"+{quantity}", '単価': int(avg_price), '金額': int(move_val),
                    '担当者名': operator_name, '担当者所属': operator_dept, '出庫先': location, '備考': input_note
                }])
                df_history = pd.concat([df_history, hist_out, hist_in], ignore_index=True)
                save_data(df_history, HISTORY_FILE)

                # PDF
                tx_data = {'type': 'return', 'date': record_str, 'operator': operator_name, 'from': location, 'to': target_loc, 'code': item_data.get('商品コード', '-'), 'name': return_name, 'maker': manufacturer, 'sub': sub_category, 'qty': quantity, 'unit': unit, 'note': input_note}
                try:
                    st.session_state['latest_voucher'] = generate_pdf_voucher(tx_data)
                    st.session_state['latest_voucher_name'] = f"return_{record_str_filename}.pdf"
                except Exception as e: st.sidebar.error(f"PDF生成エラー: {e}")

                st.session_state['last_msg'] = f'{return_name} の返却処理（移動）完了'
                st.session_state['reset_form'] = True
                st.rerun()

            # 3. 棚卸
            elif action_type == '棚卸':
                current_row = df_inventory[(df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location)]
                locked_qty = 0
                if st.session_state['stocktaking_mode']:
                    snap = st.session_state['inventory_snapshot']
                    snap_row = snap[(snap['商品名'] == name) & (snap['保管場所'] == location)]
                    if not snap_row.empty: locked_qty = int(pd.to_numeric(snap_row.iloc[0]['在庫数'], errors='coerce') or 0)
                else:
                    if not current_row.empty: locked_qty = int(pd.to_numeric(current_row.iloc[0]['在庫数'], errors='coerce') or 0)
                
                actual_qty = quantity
                
                # 金額計算
                cur_val = 0.0; cur_qty = 0
                if not current_row.empty:
                    cur_val = float(pd.to_numeric(current_row.iloc[0]['在庫金額'], errors='coerce') or 0)
                    cur_qty = int(pd.to_numeric(current_row.iloc[0]['在庫数'], errors='coerce') or 0)
                
                avg_price = (cur_val / cur_qty) if cur_qty > 0 else 0
                new_val = actual_qty * avg_price
                diff_amount = new_val - cur_val
                
                df_inventory = df_inventory[~((df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location))]
                df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                    '商品名': name, 'メーカー': manufacturer, '分類': category, 'サブカテゴリ': sub_category,
                    '保管場所': location, '在庫数': actual_qty, '単位': unit,
                    '平均単価': int(avg_price), '在庫金額': int(new_val)
                }])], ignore_index=True)
                save_data(df_inventory, INVENTORY_FILE)
                
                hist_row = pd.DataFrame([{
                    '日時': record_str, '商品名': name, '保管場所': location, '処理': '棚卸',
                    '数量': f"修正: {locked_qty}→{actual_qty}", '単価': int(avg_price), '金額': int(diff_amount),
                    '担当者名': operator_name, '担当者所属': operator_dept, '出庫先': '-', '備考': input_note
                }])
                df_history = pd.concat([df_history, hist_row], ignore_index=True)
                save_data(df_history, HISTORY_FILE)
                
                st.session_state['last_msg'] = f'{name} の棚卸完了 ({locked_qty}→{actual_qty})'
                st.session_state['reset_form'] = True
                st.rerun()

            # 4. その他（購入、客先出庫）
            elif location:
                if action_type == '客先出庫':
                    if not destination_code or len(destination_code) != 7 or not destination_code.isdigit():
                        st.sidebar.error("出庫先コードは7桁の数字で入力してください")
                        st.stop()

                current_row = df_inventory[(df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location)]
                curr_qty = 0; curr_val = 0.0
                if not current_row.empty:
                    curr_qty = int(pd.to_numeric(current_row.iloc[0]['在庫数'], errors='coerce') or 0)
                    curr_val = float(pd.to_numeric(current_row.iloc[0]['在庫金額'], errors='coerce') or 0)
                
                log_qty = ""; log_price = 0; log_amount = 0
                new_qty = 0; new_val = 0.0; new_avg = 0

                if action_type == '購入入庫':
                    if input_price <= 0:
                        st.sidebar.error("購入単価を入力してください")
                        st.stop()
                    move_amount = quantity * input_price
                    new_qty = curr_qty + quantity
                    new_val = curr_val + move_amount
                    log_qty = f"+{quantity}"
                    log_price = int(input_price)
                    log_amount = int(move_amount)
                
                elif action_type == '客先出庫':
                    if curr_qty < quantity:
                        st.sidebar.error("在庫不足です")
                        st.stop()
                    avg_price = (curr_val / curr_qty) if curr_qty > 0 else 0
                    move_amount = quantity * avg_price
                    new_qty = curr_qty - quantity
                    new_val = curr_val - move_amount
                    log_qty = f"-{quantity}"
                    log_price = int(avg_price)
                    log_amount = int(move_amount)
                
                new_avg = int(new_val / new_qty) if new_qty > 0 else 0
                
                df_inventory = df_inventory[~((df_inventory['商品名'] == name) & (df_inventory['保管場所'] == location))]
                df_inventory = pd.concat([df_inventory, pd.DataFrame([{
                    '商品名': name, 'メーカー': manufacturer, '分類': category, 'サブカテゴリ': sub_category,
                    '保管場所': location, '在庫数': new_qty, '単位': unit,
                    '平均単価': new_avg, '在庫金額': int(new_val)
                }])], ignore_index=True)
                save_data(df_inventory, INVENTORY_FILE)
                
                dest_val = destination_code if action_type == '客先出庫' else '-'
                hist_row = pd.DataFrame([{
                    '日時': record_str, '商品名': name, '保管場所': location, '処理': action_type,
                    '数量': log_qty, '単価': log_price, '金額': log_amount,
                    '担当者名': operator_name, '担当者所属': operator_dept, '出庫先': dest_val, '備考': input_note
                }])
                df_history = pd.concat([df_history, hist_row], ignore_index=True)
                save_data(df_history, HISTORY_FILE)
                
                st.session_state['latest_voucher'] = None
                st.session_state['last_msg'] = f'{name} の処理完了'
                st.session_state['reset_form'] = True
                st.rerun()

# =========================================================
# 画面表示 (メインコンテンツ)
# =========================================================
tab_titles = ["📦 現在庫一覧", "📜 入出庫履歴", "📝 棚卸結果", "📒 商品マスタ一覧", "📅 締め日一覧"]
if st.session_state['user_code'] == '0001':
    tab_titles.append("👥 ユーザー一覧")
    tab_titles.append("🏭 倉庫一覧")
    tab_titles.append("🏭 メーカー一覧")
    tab_titles.append("🔌 機器種類一覧")

tabs = st.tabs(tab_titles)

# -----------------------------
# 1. 在庫一覧 (Tab1)
# -----------------------------
with tabs[0]:
    view_mode = st.radio("表示基準", ["現在（リアルタイム）", "月次締め（過去時点）"], horizontal=True)
    display_date_str = "現在"

    if view_mode == "現在（リアルタイム）":
        target_inventory_df = df_inventory.copy()
    else:
        if df_fiscal.empty:
            st.warning("締め日が設定されていません。")
            target_inventory_df = pd.DataFrame(columns=df_inventory.columns)
        else:
            fiscal_opts = df_fiscal['表示用'].tolist()
            selected_display_text = st.selectbox("対象期間を選択", fiscal_opts, index=len(fiscal_opts)-1)
            selected_row = df_fiscal[df_fiscal['表示用'] == selected_display_text].iloc[0]
            closing_date_str = selected_row['締め年月日']
            display_date_str = f"{selected_display_text} 時点"
            st.info(f"📅 {closing_date_str} 時点の在庫を計算")
            limit_dt = pd.to_datetime(f"{closing_date_str} 23:59:59")
            target_inventory_df = build_inventory_asof(df_history, df_item_master, limit_dt, allowed_warehouses)

    if not target_inventory_df.empty:
        target_inventory_df['削除用表示'] = target_inventory_df['商品名'].astype(str) + ' (' + target_inventory_df['保管場所'].astype(str) + ')'

    view_df = target_inventory_df.copy()
    if allowed_warehouses: view_df = view_df[view_df['保管場所'].isin(allowed_warehouses)]
    else: view_df = view_df[0:0]

    col1, col2, col3 = st.columns(3)
    with col1: f_loc = st.selectbox('倉庫', ['すべて'] + allowed_warehouses)
    with col2: f_cat = st.selectbox('分類', ['すべて', '機器', '部材', 'その他'])
    with col3: f_maker = st.selectbox('メーカー', ['すべて'] + (df_manufacturer['メーカー名'].tolist() if not df_manufacturer.empty else []))

    if f_loc != 'すべて': view_df = view_df[view_df['保管場所'] == f_loc]
    if f_cat != 'すべて': view_df = view_df[view_df['分類'] == f_cat]
    if f_maker != 'すべて': view_df = view_df[view_df['メーカー'] == f_maker]

    if not df_item_master.empty and not view_df.empty:
        if 'サブカテゴリ' in view_df.columns: view_df = view_df.drop(columns=['サブカテゴリ'])
        view_df = pd.merge(view_df, df_item_master[['商品名', 'サブカテゴリ', '標準単価']], on='商品名', how='left')
    elif '標準単価' not in view_df.columns: view_df['標準単価'] = 0

    if not df_history.empty and not view_df.empty:
        df_buy = df_history[df_history['処理'] == '購入入庫'].copy()
        if not df_buy.empty:
            df_buy['日時_dt'] = pd.to_datetime(df_buy['日時'], errors='coerce')
            df_buy = df_buy.sort_values('日時_dt', ascending=False)
            df_last = df_buy.drop_duplicates(subset=['商品名', '保管場所'])[['商品名', '保管場所', '単価']]
            df_last = df_last.rename(columns={'単価': '最終購入単価'})
            view_df = pd.merge(view_df, df_last, on=['商品名', '保管場所'], how='left')
        else: view_df['最終購入単価'] = 0
    elif '最終購入単価' not in view_df.columns: view_df['最終購入単価'] = 0

    num_cols = ['標準単価', '平均単価', '最終購入単価', '在庫数', '在庫金額']
    for c in num_cols:
        if c in view_df.columns: view_df[c] = pd.to_numeric(view_df[c], errors='coerce').fillna(0)
    view_df['在庫金額'] = view_df['在庫数'] * view_df['平均単価']

    st.write(f"▼ 在庫一覧（基準: **{display_date_str}**）")
    st.dataframe(view_df, use_container_width=True, column_order=['商品名', 'メーカー', '分類', 'サブカテゴリ', '保管場所', '在庫数', '単位', '標準単価', '平均単価', '最終購入単価', '在庫金額'])

    # Excel出力 (Tab1)
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            view_df.to_excel(writer, index=False, sheet_name='現在庫')
        st.download_button(label="📥 現在庫一覧をExcel出力", data=buffer.getvalue(), file_name=f"inventory_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dl_inv")
    except Exception as e: st.error(f"Excel出力エラー: {e}")

    if view_mode == "現在（リアルタイム）" and not view_df.empty:
        st.divider()
        st.write("▼ データの削除（修正用）")
        del_target = st.selectbox('削除するデータ', view_df['削除用表示'].unique())
        if st.button('削除実行', key='btn_inv_del', disabled=not st.checkbox("確認", key="chk_inv_del")):
            tmp = df_inventory.copy()
            tmp['削除用表示'] = tmp['商品名'].astype(str) + ' (' + tmp['保管場所'].astype(str) + ')'
            tmp = tmp[tmp['削除用表示'] != del_target]
            tmp = tmp.drop(columns=['削除用表示'], errors='ignore')
            save_data(tmp, INVENTORY_FILE)
            st.rerun()

# -----------------------------
# 2. 履歴 (Tab2)
# -----------------------------
with tabs[1]:
    st.write("過去の動き（最新順）")
    
    # 期間指定フィルタ
    hist_period_mode = st.radio("表示期間", ["全期間", "期間指定"], horizontal=True)
    selected_hist_period = None
    if hist_period_mode == "期間指定":
        if not df_fiscal.empty:
            period_opts = df_fiscal['表示用'].tolist()
            selected_hist_period = st.selectbox("対象期間を選択", period_opts, index=len(period_opts)-1, key="hist_period_sel")
        else: st.warning("締め日設定がありません")
    
    # 倉庫フィルタ (UI追加)
    hist_loc_opts = ['すべて'] + allowed_warehouses
    hist_loc_filter = st.selectbox("倉庫絞り込み", hist_loc_opts, key="hist_loc_filter")
    
    # データ準備 (全履歴)
    view_hist = df_history.copy()
    view_hist['dt_obj'] = pd.to_datetime(view_hist['日時'], errors='coerce')
    
    # 倉庫フィルタ適用
    if hist_loc_filter != 'すべて':
        view_hist = view_hist[view_hist['保管場所'] == hist_loc_filter]
    elif allowed_warehouses:
        view_hist = view_hist[view_hist['保管場所'].isin(allowed_warehouses)]
    
    # --- 処理後在庫(Running Balance)計算ロジック ---
    # 計算のために時系列昇順にソート
    view_hist = view_hist.sort_values('dt_obj', ascending=True)
    
    # 各商品・倉庫ごとの現在庫を追跡する辞書
    # key: (商品名, 保管場所), value: int
    inventory_map = {}
    balance_list = []
    
    for _, row in view_hist.iterrows():
        key = (row['商品名'], row['保管場所'])
        current_val = inventory_map.get(key, 0)
        
        op = row['処理']
        k, v = parse_qty_str(row['数量'])
        
        if k == 'delta':
            if op in ['購入入庫', '移動入庫', '返却入庫']:
                current_val += abs(v)
            elif op in ['出庫', '移動出庫', '返却出庫', '客先出庫']:
                current_val -= abs(v)
        elif k == 'set_restore' and isinstance(v, tuple):
            current_val = v[1]
        elif k == 'set' and v is not None:
            current_val = v
            
        if current_val < 0: current_val = 0
        inventory_map[key] = current_val
        balance_list.append(current_val)
        
    view_hist['処理後在庫'] = balance_list
    
    # --- 期間フィルタ適用 ---
    if hist_period_mode == "期間指定" and selected_hist_period:
         f_row = df_fiscal[df_fiscal['表示用'] == selected_hist_period].iloc[0]
         start_ts = pd.Timestamp(f_row['start_dt']).replace(hour=0, minute=0, second=0)
         end_ts = pd.Timestamp(f_row['dt']).replace(hour=23, minute=59, second=59)
         view_hist = view_hist[(view_hist['dt_obj'] >= start_ts) & (view_hist['dt_obj'] <= end_ts)]

    # 表示用に降順に戻す
    view_hist = view_hist.sort_values('dt_obj', ascending=False)

    if not view_hist.empty:
        view_hist['削除用表示'] = view_hist['日時'].astype(str) + ' | ' + view_hist['商品名'].astype(str) + ' | ' + view_hist['処理'].astype(str) + ' | ' + view_hist['数量'].astype(str)
    for c in ['単価', '金額']:
        if c in view_hist.columns: view_hist[c] = pd.to_numeric(view_hist[c], errors='coerce').fillna(0)

    st.dataframe(
        view_hist, 
        use_container_width=True,
        column_order=['日時', '商品名', '保管場所', '処理', '数量', '処理後在庫', '単価', '金額', '担当者名', '出庫先', '備考']
    )

    # Excel出力 (Tab2)
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            view_hist.to_excel(writer, index=False, sheet_name='入出庫履歴')
        st.download_button(label="📥 入出庫履歴をExcel出力", data=buffer.getvalue(), file_name=f"history_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dl_hist")
    except Exception as e: st.error(f"Excel出力エラー: {e}")

    st.divider()
    st.write("#### 🖨️ 伝票発行")
    if not view_hist.empty:
        hist_opts = [f"{r['削除用表示']}" for i, r in view_hist.iloc[::-1].iterrows() if r['処理'] in ['移動出庫', '移動入庫', '出庫', '客先出庫', '機器返却', '返却出庫', '返却入庫']]
        target_hist_str = st.selectbox("伝票を発行する履歴を選択", hist_opts, key="sel_hist_voucher") if hist_opts else None
        if target_hist_str and st.button("伝票生成", key="btn_hist_gen"):
            target_hist_row = view_hist[view_hist['削除用表示'] == target_hist_str].iloc[0]
            m_row = df_item_master[df_item_master['商品名'] == target_hist_row['商品名']]
            if m_row.empty and '(返却品)' in target_hist_row['商品名']:
                orig_name = target_hist_row['商品名'].replace(' (返却品)', '')
                m_row = df_item_master[df_item_master['商品名'] == orig_name]

            if not m_row.empty:
                m_data = m_row.iloc[0]
                tx_type = 'sales'
                if '移動' in target_hist_row['処理']: tx_type = 'transfer'
                elif '返却' in target_hist_row['処理']: tx_type = 'return'
                
                # 簡易的な場所推定
                if target_hist_row['処理'] in ['移動出庫', '返却出庫']:
                    loc_from = target_hist_row['保管場所']; loc_to = target_hist_row['出庫先']
                elif target_hist_row['処理'] in ['移動入庫', '返却入庫']:
                    loc_from = target_hist_row['出庫先']; loc_to = target_hist_row['保管場所']
                else: 
                    loc_from = target_hist_row['保管場所']; loc_to = target_hist_row['出庫先']

                tx_data = {'type': tx_type, 'date': str(target_hist_row['日時']), 'operator': str(target_hist_row['担当者名']), 'from': loc_from, 'to': loc_to, 'code': m_data['商品コード'], 'name': str(target_hist_row['商品名']), 'maker': m_data['メーカー'], 'sub': m_data['サブカテゴリ'], 'qty': str(target_hist_row['数量']).replace('+','').replace('-',''), 'unit': m_data['単位'], 'note': str(target_hist_row.get('備考', ''))}
                try:
                    pdf_data = generate_pdf_voucher(tx_data)
                    st.download_button(label="📥 ダウンロード開始", data=pdf_data, file_name=f"voucher.pdf", mime="application/pdf")
                except Exception as e: st.error(f"エラー: {e}")
            else: st.error("商品マスタが見つかりません")

    st.divider()
    st.write("▼ 履歴データの削除")
    if not view_hist.empty:
        del_hist_target = st.selectbox("削除する履歴を選択", view_hist['削除用表示'].unique(), key="sel_hist_del")
        if st.button("履歴削除実行", key="btn_hist_del", disabled=not st.checkbox("本当に削除しますか？", key="chk_hist_del")):
            target_data = df_history[df_history['削除用表示'] == del_hist_target]
            if not target_data.empty:
                t_row = target_data.iloc[0]
                t_name = t_row['商品名']; t_loc = t_row['保管場所']; t_qty_str = t_row['数量']
                t_amount = float(pd.to_numeric(t_row['金額'], errors='coerce') or 0)
                revert_qty = 0; revert_amount = 0
                kind, val = parse_qty_str(t_qty_str)
                if kind == 'delta':
                    revert_qty = -1 * val
                    if val > 0: revert_amount = -1 * abs(t_amount)
                    else: revert_amount = abs(t_amount)
                elif kind == 'set_restore':
                    if isinstance(val, tuple): revert_qty = val[0] - val[1]
                    else: revert_qty = 0
                    revert_amount = 0 
                mask = (df_inventory['商品名'] == t_name) & (df_inventory['保管場所'] == t_loc)
                if not df_inventory[mask].empty:
                    curr_qty = float(pd.to_numeric(df_inventory.loc[mask, '在庫数'], errors='coerce'))
                    curr_val = float(pd.to_numeric(df_inventory.loc[mask, '在庫金額'], errors='coerce'))
                    new_qty = max(0, curr_qty + revert_qty)
                    new_val = max(0, curr_val + revert_amount)
                    new_avg = int(new_val / new_qty) if new_qty > 0 else 0
                    df_inventory.loc[mask, '在庫数'] = int(new_qty)
                    df_inventory.loc[mask, '在庫金額'] = int(new_val)
                    df_inventory.loc[mask, '平均単価'] = int(new_avg)
                    save_data(df_inventory, INVENTORY_FILE)
            tmp = df_history.copy()
            tmp['削除用表示'] = tmp['日時'].astype(str) + ' | ' + tmp['商品名'].astype(str) + ' | ' + tmp['処理'].astype(str) + ' | ' + tmp['数量'].astype(str)
            tmp = tmp[tmp['削除用表示'] != del_hist_target]
            tmp = tmp.drop(columns=['削除用表示'], errors='ignore')
            save_data(tmp, HISTORY_FILE)
            st.success("履歴を削除し、在庫数を元に戻しました")
            st.rerun()

# -----------------------------
# 3. 棚卸結果 (Tab3)
# -----------------------------
with tabs[2]:
    st.subheader("📝 棚卸実施結果")
    
    with st.expander("📊 月次報告書 (Excel) の出力"):
        st.caption("指定した期間の入出庫・棚卸結果を集計してExcel出力します。")
        if not df_fiscal.empty:
            rep_opts = df_fiscal['表示用'].tolist()
            rep_period_txt = st.selectbox("報告対象期間", rep_opts, index=len(rep_opts)-1, key="rep_period_sel")
            rep_wh_opts = ['すべて'] + (allowed_warehouses if allowed_warehouses else [])
            rep_wh = st.selectbox("対象倉庫", rep_wh_opts, key="rep_wh_sel")
            
            # 【追加】サブカテゴリの選択
            all_subs = sorted(df_item_master['サブカテゴリ'].dropna().unique().tolist())
            target_subs = st.multiselect("対象の機器種類 (指定なしで全種類)", all_subs, key="rep_sub_sel")
            
            if st.button("Excel生成"):
                sel_row = df_fiscal[df_fiscal['表示用'] == rep_period_txt].iloc[0]
                
                # --- NaT check fix ---
                raw_start = sel_row['start_dt']
                if pd.isna(raw_start):
                    # 前の締め日がない場合は、その月の1日を開始日とする
                    raw_start = sel_row['dt'].replace(day=1)
                
                s_dt = pd.Timestamp(raw_start).replace(hour=0, minute=0, second=0)
                e_dt = pd.Timestamp(sel_row['dt']).replace(hour=23, minute=59, second=59)
                
                # pass target_subs to function
                excel_data = generate_monthly_report_excel(df_history, df_item_master, df_location, rep_period_txt, s_dt, e_dt, rep_wh, target_subs)
                if excel_data:
                    st.download_button(
                        label="📥 ダウンロード",
                        data=excel_data,
                        file_name=f"monthly_report_{rep_period_txt[:7]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Excel生成に失敗しました (xlsxwriterがインストールされていない可能性があります)")
        else:
            st.warning("締め日が設定されていません。")

    st.divider()

    hist_stock = df_history[df_history['処理'] == '棚卸'].copy()
    if not hist_stock.empty:
        hist_stock['dt_obj'] = pd.to_datetime(hist_stock['日時'], errors='coerce')
        hist_stock = hist_stock.sort_values('dt_obj', ascending=False)
        hist_stock = hist_stock.drop_duplicates(subset=['商品名', '保管場所'], keep='first')
        
        target_locs = ['すべて'] + (allowed_warehouses if allowed_warehouses else [])
        selected_loc = st.selectbox("倉庫で絞り込み", target_locs, key="stocktake_loc_filter")
        if selected_loc != 'すべて': hist_stock = hist_stock[hist_stock['保管場所'] == selected_loc]
        elif allowed_warehouses: hist_stock = hist_stock[hist_stock['保管場所'].isin(allowed_warehouses)]

        display_data = []
        for _, row in hist_stock.iterrows():
            kind, val = parse_qty_str(row['数量'])
            if kind == 'set_restore' and isinstance(val, tuple):
                old_val, new_val = val
                diff = new_val - old_val
                diff_str = f"+{diff}" if diff > 0 else str(diff)
            else:
                new_val = row['数量']; diff_str = "-"; old_val = "-"

            m_row = df_item_master[df_item_master['商品名'] == row['商品名']]
            maker = ""; cat = ""; sub = ""
            if not m_row.empty:
                m = m_row.iloc[0]
                maker = m['メーカー']; cat = m['分類']; sub = m['サブカテゴリ']
            
            unit_price = int(float(row.get('単価', 0) or 0))
            stock_amount = 0
            if isinstance(new_val, int): stock_amount = new_val * unit_price
            
            display_data.append({'実施日時': row['日時'], '商品名': row['商品名'], 'メーカー': maker, '分類': cat, '機器種類': sub, '保管場所': row['保管場所'], 'ロック数(帳簿)': old_val, '棚卸数(実数)': new_val, '差分': diff_str, '平均単価': unit_price, '在庫金額': stock_amount, '担当者': row['担当者名']})

        if display_data:
            df_display = pd.DataFrame(display_data)
            st.dataframe(df_display, use_container_width=True)
            try:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_display.to_excel(writer, index=False, sheet_name='棚卸結果')
                st.download_button(label="📥 棚卸結果をExcel出力", data=buffer.getvalue(), file_name=f"stocktaking_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e: st.error(f"Excel出力エラー: {e}")
        else: st.info("表示対象の棚卸データがありません。")
    else: st.info("棚卸の実施履歴がありません。")

# -----------------------------
# 4. マスタ (Tab4) ~
# -----------------------------
with tabs[3]:
    st.write("商品マスタ")
    if not df_item_master.empty:
        df_item_master['標準単価'] = pd.to_numeric(df_item_master['標準単価'], errors='coerce').fillna(0)
    st.dataframe(df_item_master, use_container_width=True)

with tabs[4]:
    st.subheader("📅 締め日スケジュール")
    if not df_fiscal.empty:
        today_str = datetime.date.today().strftime('%Y-%m-%d')
        future_dates = df_fiscal[df_fiscal['締め年月日'] >= today_str].sort_values('締め年月日')
        if not future_dates.empty:
            next_row = future_dates.iloc[0]
            st.info(f"🔔 **次回の締め日: {next_row['締め年月日']}** （{next_row['表示用']}）")
        else: st.info("これ以降の締め日設定はありません。")
        st.write("▼ 全リスト")
        st.dataframe(df_fiscal[['対象年月', '締め年月日', '表示用']], use_container_width=True)
    else: st.warning("締め日データがありません。")

if st.session_state['user_code'] == '0001':
    with tabs[5]:
        st.subheader("👥 登録ユーザー一覧")
        view_staff = df_staff.drop(columns=['パスワード'], errors='ignore')
        st.dataframe(view_staff, use_container_width=True)
    with tabs[6]:
        st.subheader("🏭 登録倉庫一覧")
        st.dataframe(df_location, use_container_width=True)
    with tabs[7]:
        st.subheader("🏭 登録メーカー一覧")
        st.dataframe(df_manufacturer, use_container_width=True)
    with tabs[8]:
        st.subheader("🔌 登録機器種類一覧")
        st.dataframe(df_category, use_container_width=True)