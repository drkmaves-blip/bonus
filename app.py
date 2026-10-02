import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import io
import re
import pdfplumber
from datetime import datetime

# ==========================================
# 1. 基本設定與 30 個品牌清單
# ==========================================
st.set_page_config(
    page_title="雪茄批發報價自動整理系統", 
    page_icon="🍂",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自訂 CSS 美化介面
st.markdown("""
<style>
    /* 頂部標題區塊 */
    .main-header {
        background: linear-gradient(135deg, #1F4E78 0%, #2C5E3B 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .main-header p { color: #E0E0E0; margin: 0.3rem 0 0 0; font-size: 0.95rem; }
    
    /* KPI 卡片 */
    [data-testid="stMetric"] {
        background: #F8F9FA;
        border: 1px solid #E9ECEF;
        border-radius: 10px;
        padding: 0.8rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    [data-testid="stMetricLabel"] { font-size: 0.85rem !important; }
    
    /* 側邊欄美化 */
    section[data-testid="stSidebar"] > div {
        background: linear-gradient(180deg, #F8F9FA 0%, #FFFFFF 100%);
    }
    
    /* 表格美化 */
    .stDataFrame { border-radius: 8px; overflow: hidden; }
    
    /* 按鈕美化 */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1F4E78, #2C5E3B) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.5rem 1.5rem !important;
    }
</style>
""", unsafe_allow_html=True)

# 頂部標題
st.markdown("""
<div class="main-header">
    <h1>🍂 雪茄批發報價單自動清洗與彙整工具</h1>
    <p>上傳各家 Excel / PDF 報價單 ➔ 自動辨識 30 個品牌 + 10 大供應商 ➔ 產出精簡版 Excel・即時比價矩陣・數據儀表板</p>
</div>
""", unsafe_allow_html=True)

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
    # 清理檔名中常見的雜訊
    name = filename.split('.')[0]
    # 移除日期格式 (例如 9.20、09.24、2024)
    name = re.sub(r'\d{1,2}\.\d{1,2}', '', name)
    name = re.sub(r'20\d{2}', '', name)
    # 移除常見贅字
    for noise in ['报价', '报價', '港币', '港幣', '最新', '澳门', '澳門', '批发', '批發', '价格', '價格', '号', '號', '(1)', '（1）']:
        name = name.replace(noise, '')
    return name.strip() or filename.split('.')[0]

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
    
    accent_fill = PatternFill(start_color='2C5E3B', end_color='2C5E3B', fill_type='solid')
    
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
    if ws_list.max_row > 1:
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
                int(subset['價格'].min()), int(subset['價格'].max()), round(subset['價格'].mean(), 1)
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
# 3.5 價格判斷輔助函數
# ==========================================
def is_price(val):
    """判斷一個字串是否為有效的價格數值"""
    try:
        cv = re.sub(r'[,$\s]|HKD|RMB|USD|EUR|¥|￥', '', str(val), flags=re.IGNORECASE)
        return cv.replace('.', '', 1).isdigit() and float(cv) > 50
    except:
        return False

def extract_price(val):
    """從字串中萃取價格數值"""
    cv = re.sub(r'[,$\s]|HKD|RMB|USD|EUR|¥|￥', '', str(val), flags=re.IGNORECASE)
    return int(float(cv))

# ==========================================
# 4. 網頁前端與自動解析執行
# ==========================================

# 側邊欄上方：檔案上傳區
st.sidebar.markdown("### 📂 檔案上傳區")
uploaded_files = st.sidebar.file_uploader(
    "拖曳或點擊上傳報價單", 
    type=["xlsx", "xls", "pdf"], 
    accept_multiple_files=True,
    help="支援 Excel (.xlsx, .xls) 和 PDF 格式，可一次上傳多個檔案"
)

parse_button = st.sidebar.button("🚀 開始解析並合併轉檔", use_container_width=True, type="primary")

# 側邊欄下方：顯示解析進度 / 檔案資訊
if uploaded_files:
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**已選取 {len(uploaded_files)} 個檔案：**")
    for f in uploaded_files:
        size_kb = f.size / 1024
        icon = "📗" if f.name.endswith(('.xlsx', '.xls')) else "📕"
        st.sidebar.caption(f"{icon} {f.name} ({size_kb:.0f} KB)")

# 主畫面上方的操作提示
if not uploaded_files:
    st.info("👈 請在左側上傳報價單檔案，支援 Excel 和 PDF 格式。上傳後點擊「🚀 開始解析」即可開始！")

if parse_button and uploaded_files:
    all_records = []
    file_stats = []  # 記錄每個檔案的解析統計
    
    progress_bar = st.progress(0, text="正在解析檔案...")
    
    for idx, file in enumerate(uploaded_files):
        vendor = get_standard_vendor(file.name)
        file_record_count = 0
        try:
            raw_rows = []
            file_ext = file.name.split('.')[-1].lower()

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
                        price = extract_price(v)
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
                    file_record_count += 1
                    
            file_stats.append({"檔名": file.name, "供應商": vendor, "抓取筆數": file_record_count, "狀態": "✅ 成功"})
        except Exception as e:
            file_stats.append({"檔名": file.name, "供應商": vendor, "抓取筆數": 0, "狀態": f"❌ 失敗: {e}"})
            
        # 更新進度條
        progress_bar.progress((idx + 1) / len(uploaded_files), text=f"正在解析：{file.name} ({idx+1}/{len(uploaded_files)})")
    
    progress_bar.empty()

    # 顯示各檔案解析結果報告
    st.session_state['file_stats'] = file_stats

    if all_records:
        df_result = pd.DataFrame(all_records)
        
        # 依 30 個品牌排序並去重複
        brand_order_map = {b: i for i, b in enumerate(TARGET_BRANDS)}
        brand_order_map['其他品牌'] = 999
        df_result['排序權重'] = df_result['品牌分類'].map(brand_order_map)
        
        # 執行排序與去重
        df_result = df_result.sort_values(by=['排序權重', '品名規格', '價格'])
        df_result = df_result.drop_duplicates(subset=['供應商', '品牌分類', '品名規格']).drop(columns=['排序權重'])

        # 把結果存入 session_state
        st.session_state['df_result'] = df_result
    else:
        st.warning("⚠️ 所有檔案均未成功解析出有效資料，請檢查檔案格式。")

# ==========================================
# 5. 結果展示區（使用 session_state 持久化）
# ==========================================
if 'df_result' in st.session_state:
    df_result = st.session_state['df_result']

    # --- 顯示解析報告 (可收合) ---
    if 'file_stats' in st.session_state:
        with st.expander("📋 各檔案解析結果報告", expanded=False):
            stats_df = pd.DataFrame(st.session_state['file_stats'])
            st.dataframe(stats_df, use_container_width=True, hide_index=True)
            total_parsed = stats_df['抓取筆數'].sum()
            success_count = stats_df[stats_df['狀態'] == '✅ 成功'].shape[0]
            st.caption(f"共解析 {len(stats_df)} 個檔案，{success_count} 個成功，原始抓取 {total_parsed} 筆 → 去重後 {len(df_result)} 筆")

    # --- 側邊欄：篩選器 ---
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 篩選器 (Slicer)")
    
    # 品牌篩選
    all_brands = sorted(df_result['品牌分類'].unique())
    selected_brands = st.sidebar.multiselect("📌 品牌分類", all_brands)
    
    # 供應商篩選
    all_vendors = sorted(df_result['供應商'].unique())
    selected_vendors = st.sidebar.multiselect("🏬 供應商", all_vendors)
    
    # 價格區間篩選
    st.sidebar.markdown("**💰 價格區間 (HKD)**")
    price_min = int(df_result['價格'].min())
    price_max = int(df_result['價格'].max())
    price_range = st.sidebar.slider(
        "拖曳調整價格範圍",
        min_value=price_min, max_value=price_max, 
        value=(price_min, price_max),
        step=50,
        format="$%d"
    )
    
    # 品名關鍵字搜尋
    keyword = st.sidebar.text_input("🔎 品名關鍵字搜尋", placeholder="例如：短丘、D4、BHK、鋁管")
    
    # 一鍵清除篩選
    if st.sidebar.button("🗑️ 清除所有篩選條件", use_container_width=True):
        st.rerun()
    
    # 執行篩選
    df_display = df_result.copy()
    if selected_brands:
        df_display = df_display[df_display['品牌分類'].isin(selected_brands)]
    if selected_vendors:
        df_display = df_display[df_display['供應商'].isin(selected_vendors)]
    df_display = df_display[(df_display['價格'] >= price_range[0]) & (df_display['價格'] <= price_range[1])]
    if keyword:
        df_display = df_display[df_display['品名規格'].str.contains(keyword, case=False, na=False)]
        
    # --- KPI 數據儀表板 ---
    st.markdown("### 📈 數據總覽 Dashboard")
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("📦 篩選品項數", f"{len(df_display)} 筆")
    col2.metric("🏢 涵蓋供應商", f"{df_display['供應商'].nunique()} 家")
    col3.metric("🏷️ 涵蓋品牌數", f"{df_display['品牌分類'].nunique()} 個")
    
    if not df_display.empty:
        col4.metric("⬇️ 最低報價", f"${int(df_display['價格'].min()):,}")
        col5.metric("⬆️ 最高報價", f"${int(df_display['價格'].max()):,}")
    else:
        col4.metric("⬇️ 最低報價", "$0")
        col5.metric("⬆️ 最高報價", "$0")
    
    st.caption(f"📊 總庫存 {len(df_result)} 筆 → 篩選顯示 {len(df_display)} 筆 | 平均報價 ${int(df_display['價格'].mean()) if not df_display.empty else 0:,} HKD")
        
    # --- 下載按鈕區 ---
    dl_col1, dl_col2 = st.columns(2)
    
    today_str = datetime.now().strftime("%Y%m%d")
    
    with dl_col1:
        excel_bytes = generate_excel(df_display)
        st.download_button(
            label="📥 下載彙總精簡版 Excel",
            data=excel_bytes,
            file_name=f"雪茄批發報價彙總_精簡版_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with dl_col2:
        # 樞紐分析表下載
        if not df_display.empty:
            pivot_dl = df_display.pivot_table(
                index=['品牌分類', '品名規格'], columns='供應商', values='價格', aggfunc='min'
            ).astype('Int64')
            output_pivot = io.BytesIO()
            with pd.ExcelWriter(output_pivot, engine='openpyxl') as writer:
                pivot_dl.to_excel(writer, sheet_name="跨店比價矩陣")
            st.download_button(
                label="📥 下載跨店比價矩陣 Excel",
                data=output_pivot.getvalue(),
                file_name=f"跨店比價矩陣_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_pivot",
                use_container_width=True
            )
    
    # --- 圖表分析 ---
    with st.expander("📊 查看品牌報價數量分佈圖 & 供應商品項數對比"):
        chart_col1, chart_col2 = st.columns(2)
        with chart_col1:
            st.markdown("**各品牌報價數量**")
            brand_counts = df_display['品牌分類'].value_counts()
            st.bar_chart(brand_counts)
        with chart_col2:
            st.markdown("**各供應商品項數**")
            vendor_counts = df_display['供應商'].value_counts()
            st.bar_chart(vendor_counts)
    
    st.markdown("---")
    
    # --- 分頁標籤切換不同視圖 ---
    tab1, tab2 = st.tabs(["📊 跨店比價矩陣", "📋 全品項清單"])
    
    with tab1:
        st.markdown("將同一款雪茄在不同店家的報價「橫向展開」，一眼看出哪家有貨與最低價！")
        
        if not df_display.empty:
            pivot_df = df_display.pivot_table(
                index=['品牌分類', '品名規格'], 
                columns='供應商', 
                values='價格', 
                aggfunc='min'
            )
            pivot_df = pivot_df.astype('Int64')
            
            # 用 Styler 標記每行最低價 (高亮綠色)
            def highlight_min(row):
                numeric_vals = row.dropna()
                if numeric_vals.empty:
                    return ['' for _ in row]
                min_val = numeric_vals.min()
                return ['background-color: #D4EDDA; font-weight: bold' if pd.notna(v) and v == min_val else '' for v in row]
            
            styled_pivot = pivot_df.style.apply(highlight_min, axis=1)
            st.dataframe(styled_pivot, use_container_width=True, height=600)
        else:
            st.warning("⚠️ 目前篩選條件下無資料，請調整左側篩選器。")
    
    with tab2:
        if not df_display.empty:
            st.dataframe(
                df_display.reset_index(drop=True), 
                use_container_width=True, 
                height=600,
                hide_index=True
            )
        else:
            st.warning("⚠️ 目前篩選條件下無資料，請調整左側篩選器。")

# --- 頁腳 ---
st.sidebar.markdown("---")
st.sidebar.caption(f"🍂 雪茄報價系統 v2.0\n\n更新時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
