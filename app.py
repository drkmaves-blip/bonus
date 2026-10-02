import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import io
import re
import pdfplumber

# ==========================================
# 1. 基本設定與 30 個品牌清單
# ==========================================
st.set_page_config(page_title="雪茄批發報價自動整理系統", layout="wide")
st.title("🍂 雪茄批發報價單自動清洗與彙整工具")
st.markdown("上傳各家 Excel 報價單 ➔ 自動分類 30 個品牌 ➔ 產出 5 欄位精簡版 Excel (含篩選器與總表)")

TARGET_BRANDS = [
    "千里達", "大衛杜夫", "比雅達", "世界之王", "卡諾之花", "古巴榮耀", "外交官", "多明尼加之花", "多爾賽碼頭", "好友",
    "我的父親", "帕特加斯", "拉斐爾", "金特羅", "威古洛", "玻利瓦", "胡安洛佩斯", "庫阿巴", "烏普曼", "高希霸",
    "富恩特", "獅子王", "雷蒙阿隆尼", "蒙特", "潑辣", "潘趣", "潮牌CAO", "豐塞卡", "羅密歐與茱麗葉", "羅賓納"
]

# ==========================================
# 1.5 供應商名稱標準化
# ==========================================
VENDOR_RULES = {
    "天空": "古巴天空",
    "千源": "千源古巴行",
    "古巴之家": "古巴之家",
    "818": "818同行批發",
    "壹茄": "壹茄批發",
    "寰宇": "寰宇之家",
    "维嘉": "維嘉批發",
    "維嘉": "維嘉批發",
    "行货": "行貨雪茄報價",
    "行貨": "行貨雪茄報價",
    "2026": "澳門批發庫存表",
    "澳门批发": "澳門批發庫存表",
    "澳門批發": "澳門批發庫存表",
    "庫存": "庫存報價(09.29)",
    "库存": "庫存報價(09.29)",
}

def get_standard_vendor(filename):
    """根據檔名關鍵字自動標準化供應商名稱"""
    for key, standard_name in VENDOR_RULES.items():
        if key in filename:
            return standard_name
    return filename.split('.')[0].replace('报价', '').replace('报價', '')

# ==========================================
# 2. 品牌分類核心邏輯
# ==========================================
def classify_brand(pname):
    """根據品名關鍵字，嚴格匹配 30 個指定品牌"""
    text = str(pname).lower()
    
    if any(k in text for k in ['富恩特', 'fuente', '海明威']): return '富恩特'
    if any(k in text for k in ['我的父亲', '我的父親', 'my father']): return '我的父親'
    if any(k in text for k in ['多米尼加之花', '多明尼加之花', 'flor dominicana']): return '多明尼加之花'
    if any(k in text for k in ['狮子王', '獅子王', '拉奥罗拉', 'aurora']): return '獅子王'
    
    # 避免「好友蒙特利」誤判為「蒙特」
    if any(k in text for k in ['好友', 'hoyo', 'epicure', '逍遥', '赛科', '聖胡安', '帕尔马斯']):
        if not any(k in text for k in ['蒙特克里斯托', 'montecristo', '蒙特2号', '蒙特4号']): 
            return '好友'
            
    if any(k in text for k in ['泼辣', '潑辣', '波尔', 'larranaga']): return '潑辣'
    if any(k in text for k in ['高希霸', 'cohiba', 'bhk', '世纪', '半世纪', '魔术师', '导师']): return '高希霸'
    if any(k in text for k in ['千里达', '千里達', '特立尼达', 'trinidad', '3t', '雷耶斯', '暗礁']): return '千里達'
    if any(k in text for k in ['蒙特', 'montecristo', '艾蒙多', '1935', '大仲马']): return '蒙特'
    if any(k in text for k in ['帕特加斯', 'partagas', '路西塔尼亚', 'd4', 'p2', 'e2']): return '帕特加斯'
    if any(k in text for k in ['罗密欧', '羅密歐', 'romeo', '短丘', '宽丘', '丘比特']): return '羅密歐與茱麗葉'
    if any(k in text for k in ['乌普曼', '烏普曼', 'upmann', '玛瑙', '鉴赏家', '半皇冠']): return '烏普曼'
    if any(k in text for k in ['玻利瓦', 'bolivar']): return '玻利瓦'
    if any(k in text for k in ['雷蒙', 'allones']): return '雷蒙阿隆尼'
    if any(k in text for k in ['多尔塞', 'quai', "d'orsay", '码头']): return '多爾賽碼頭'
    if any(k in text for k in ['潘趣', 'punch']): return '潘趣'
    if any(k in text for k in ['库阿巴', 'cuaba']): return '庫阿巴'
    if any(k in text for k in ['威古洛', 'vegueros']): return '威古洛'
    if any(k in text for k in ['胡安', 'juan lopez']): return '胡安洛佩斯'
    if any(k in text for k in ['外交官', 'diplomaticos']): return '外交官'
    if any(k in text for k in ['比亚达', '比雅達', 'piedra', '猎人']): return '比雅達'
    if any(k in text for k in ['金特罗', 'quintero', '挚爱']): return '金特羅'
    if any(k in text for k in ['世界之王', 'rey del mundo']): return '世界之王'
    if any(k in text for k in ['卡诺之花', 'flor de cano']): return '卡諾之花'
    if any(k in text for k in ['古巴荣耀', 'gloria cubana']): return '古巴榮耀'
    if any(k in text for k in ['拉斐尔', 'rafael gonzalez']): return '拉斐爾'
    if any(k in text for k in ['丰塞卡', 'fonseca']): return '豐塞卡'
    if any(k in text for k in ['罗宾纳', 'robaina']): return '羅賓納'
    if any(k in text for k in ['大卫杜夫', 'davidoff']): return '大衛杜夫'
    if any(k in text for k in ['潮牌', 'cao']): return '潮牌CAO'
    
    return '其他品牌'

# ==========================================
# 3. 產出精美 Excel 檔案邏輯
# ==========================================
def generate_excel(df_list):
    """將清洗後的 DataFrame 轉為帶有格式的 Excel 二進位資料"""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # 建立樣式
    header_font = Font(name='Microsoft JhengHei', size=11, bold=True, color='FFFFFF')
    body_font = Font(name='Microsoft JhengHei', size=10)
    bold_font = Font(name='Microsoft JhengHei', size=10, bold=True)
    
    primary_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    accent_fill = PatternFill(start_color='2C5E3B', end_color='2C5E3B', fill_type='solid')
    zebra_fill = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
    
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    # --- 工作表 1：全品項清單 ---
    ws_list = wb.create_sheet(title='全品項清單')
    headers_list = ['供應商', '品牌分類', '品名規格', '價格(HKD)', '產地/版本']
    ws_list.append(headers_list)
    
    for _, row in df_list.iterrows():
        ws_list.append([row['供應商'], row['品牌分類'], row['品名規格'], row['價格'], row.get('產地', '')])
        
    for col_num in range(1, 6):
        c = ws_list.cell(row=1, column=col_num)
        c.font, c.alignment = header_font, Alignment(horizontal='center', vertical='center')
        
    for r in range(2, ws_list.max_row + 1):
        for c in range(1, 6):
            cell = ws_list.cell(row=r, column=c)
            cell.font = body_font
            if c in [1, 2, 5]: 
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if c == 2: cell.font = bold_font
            elif c == 4:
                cell.alignment, cell.number_format = Alignment(horizontal='right', vertical='center'), '#,##0'

    # 建立正式表格 (等同於在 Excel 中按下 Ctrl+T)，方便後續插入篩選器
    tab_list = Table(displayName="DataList", ref=f"A1:E{ws_list.max_row}")
    tab_list.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_list.add_table(tab_list)
    ws_list.freeze_panes = 'A2'

    # --- 工作表 2：品牌分類總覽統計 ---
    ws_sum = wb.create_sheet(title='品牌分類總覽統計')
    headers_sum = ['品牌名稱', '品項總報價數', '最低報價(HKD)', '最高報價(HKD)', '平均報價(HKD)']
    ws_sum.append(headers_sum)
    
    # 統計邏輯
    for brand in TARGET_BRANDS + ["其他品牌"]:
        subset = df_list[df_list['品牌分類'] == brand]
        if not subset.empty:
            ws_sum.append([
                brand, len(subset), 
                int(subset['價格'].min()), int(subset['价格'].max()) if '价格' in subset else int(subset['價格'].max()), round(subset['價格'].mean(), 1)
            ])

    for col_num in range(1, 6):
        c = ws_sum.cell(row=1, column=col_num)
        c.font, c.fill, c.alignment = header_font, accent_fill, Alignment(horizontal='center', vertical='center')
        
    for r in range(2, ws_sum.max_row + 1):
        for c in range(1, 6):
            cell = ws_sum.cell(row=r, column=c)
            cell.font, cell.border = body_font, thin_border
            if c == 1: cell.font = bold_font
            elif c in [3, 4, 5]: cell.number_format = '#,##0'
            
    ws_sum.auto_filter.ref = f"A1:E{ws_sum.max_row}"
    ws_sum.freeze_panes = 'A2'

    # --- 調整自動欄寬 ---
    for ws in [ws_list, ws_sum]:
        for col in ws.columns:
            max_len = max((sum(2 if ord(ch) > 127 else 1 for ch in str(c.value or '')) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 12), 48)

    wb.save(output)
    return output.getvalue()

# ==========================================
# 4. 網頁前端與自動解析執行
# ==========================================
uploaded_files = st.file_uploader("上傳報價單（可多選 Excel / PDF）", type=["xlsx", "xls", "pdf"], accept_multiple_files=True)

if uploaded_files and st.button("🚀 開始解析並合併轉檔"):
    all_records = []
    
    for file in uploaded_files:
        vendor = get_standard_vendor(file.name)
        try:
            raw_rows = []
            file_ext = file.name.split('.')[-1].lower()
            
            def is_price(val):
                cv = re.sub(r'[,$\s]|HKD|RMB|USD|EUR|¥|￥', '', val, flags=re.IGNORECASE)
                return cv.replace('.', '', 1).isdigit() and float(cv) > 50

            if file_ext in ['xlsx', 'xls']:
                df_raw = pd.read_excel(file, header=None)
                for _, row in df_raw.iterrows():
                    current_chunk = []
                    for x in row:
                        val = str(x).strip()
                        if pd.notna(x) and val != "" and val.lower() != 'nan':
                            current_chunk.append(val)
                        else:
                            if current_chunk:
                                p_count = sum(1 for v in current_chunk if is_price(v))
                                if p_count > 1:
                                    size = max(1, len(current_chunk) // p_count)
                                    for i in range(p_count):
                                        raw_rows.append(current_chunk[i*size : (i+1)*size])
                                elif len(current_chunk) >= 2:
                                    raw_rows.append(current_chunk)
                                current_chunk = []
                    if current_chunk:
                        p_count = sum(1 for v in current_chunk if is_price(v))
                        if p_count > 1:
                            size = max(1, len(current_chunk) // p_count)
                            for i in range(p_count):
                                raw_rows.append(current_chunk[i*size : (i+1)*size])
                        elif len(current_chunk) >= 2:
                            raw_rows.append(current_chunk)
            elif file_ext == 'pdf':
                with pdfplumber.open(file) as pdf:
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        if tables:
                            for table in tables:
                                for row in table:
                                    if row:
                                        raw_rows.append([str(x).strip() for x in row if x and str(x).strip()])
                        else:
                            text = page.extract_text()
                            if text:
                                for line in text.split('\n'):
                                    cols = [x.strip() for x in re.split(r'\s{2,}|\t', line) if x.strip()]
                                    if cols:
                                        p_count = sum(1 for v in cols if is_price(v))
                                        if p_count > 1:
                                            size = max(1, len(cols) // p_count)
                                            for i in range(p_count):
                                                raw_rows.append(cols[i*size : (i+1)*size])
                                        else:
                                            raw_rows.append(cols)
                                        
            for vals in raw_rows:
                if len(vals) == 1:
                    vals = [x.strip() for x in re.split(r'\s{2,}|\t', vals[0]) if x.strip()]
                if len(vals) < 2: continue
                
                price, name, origin = None, None, ""
                
                for v in reversed(vals):
                    if is_price(v):
                        cv = re.sub(r'[,$\s]|HKD|RMB|USD|EUR|¥|￥', '', v, flags=re.IGNORECASE)
                        price = int(float(cv))
                        vals.remove(v)
                        break
                        
                if price:
                    candidates = [v for v in vals if not v.isdigit()]
                    if candidates:
                        name = max(candidates, key=len)
                        vals.remove(name)
                        
                if name and price:
                    full_text = name + " " + " ".join(vals)
                    brand = classify_brand(full_text)
                    
                    # 過濾純數字(數量)或品牌名稱，保留真正的產地版本
                    origin_candidates = [v for v in vals if len(v) < 15 and not v.isdigit() and not any(b in v for b in TARGET_BRANDS)]
                    if origin_candidates: 
                        origin = " ".join(origin_candidates)
                    
                    all_records.append({
                        "供應商": vendor, "品牌分類": brand, "品名規格": name, 
                        "價格": price, "產地": origin
                    })
        except Exception as e:
            st.error(f"❌ 解析 {file.name} 失敗: {e}")

    if all_records:
        df_result = pd.DataFrame(all_records)
        
        # 依 30 個品牌排序並去重複
        brand_order_map = {b: i for i, b in enumerate(TARGET_BRANDS)}
        brand_order_map['其他品牌'] = 999
        df_result['排序權重'] = df_result['品牌分類'].map(brand_order_map)
        
        # 執行排序與去重
        df_result = df_result.sort_values(by=['排序權重', '品名規格', '價格'])
        df_result = df_result.drop_duplicates(subset=['供應商', '品牌分類', '品名規格']).drop(columns=['排序權重'])

        # 把結果存入 session_state，這樣點擊其他按鈕時才不會重置網頁
        st.session_state['df_result'] = df_result

# 只要有暫存資料，就顯示 UI (移出 st.button 區塊)
if 'df_result' in st.session_state:
    df_result = st.session_state['df_result']

    # --- 側邊欄：內建網頁版交叉篩選器 (Slicer) ---
    st.sidebar.header("🔍 篩選器 (Slicer)")
    st.sidebar.markdown("點擊下方選單即可即時篩選表格")
    
    selected_brands = st.sidebar.multiselect("📌 品牌分類", df_result['品牌分類'].unique())
    selected_vendors = st.sidebar.multiselect("🏬 供應商", df_result['供應商'].unique())
    
    # 執行篩選
    df_display = df_result.copy()
    if selected_brands:
        df_display = df_display[df_display['品牌分類'].isin(selected_brands)]
    if selected_vendors:
        df_display = df_display[df_display['供應商'].isin(selected_vendors)]
        
    st.success(f"✅ 成功清洗並去重，共取得 {len(df_result)} 筆精簡報價資料！(當前篩選顯示 {len(df_display)} 筆)")
    
    # 產生 Excel 並提供下載
    excel_bytes = generate_excel(df_display)  # 修改為下載「篩選後」的結果
    st.download_button(
        label="📥 點擊下載彙總精簡版 Excel (已內建正式表格)",
        data=excel_bytes,
        file_name="雪茄批發報價彙總_精簡版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.markdown("---")
    
    # --- 網頁版跨店比價矩陣 (Pivot Table) ---
    st.subheader("📊 跨店比價矩陣 (橫向對比)")
    st.markdown("將同一款雪茄在不同店家的報價「橫向展開」，一眼看出哪家有貨與最低價！")
    
    # 製作樞紐分析表
    pivot_df = df_display.pivot_table(
        index=['品牌分類', '品名規格'], 
        columns='供應商', 
        values='價格', 
        aggfunc='min'
    )
    
    # 轉換為支援空值的整數格式，方便網頁乾淨顯示
    pivot_df = pivot_df.astype('Int64')
    st.dataframe(pivot_df, use_container_width=True)
    
    st.markdown("---")
    
    # --- 全品項清單展示 ---
    st.subheader("📋 篩選後全品項清單")
    st.table(df_display)
