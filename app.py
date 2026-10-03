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

# ==========================================
# 全域 CSS 樣式
# ==========================================
st.markdown("""
<style>
    /* ===== 全域字體與背景 ===== */
    .block-container { padding-top: 1rem; }
    
    /* ===== 頂部 Hero 橫幅 ===== */
    .hero {
        background: linear-gradient(135deg, #0F2027 0%, #203A43 40%, #2C5364 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 1.8rem;
        position: relative;
        overflow: hidden;
    }
    .hero::before {
        content: '';
        position: absolute;
        top: -50%; right: -20%;
        width: 400px; height: 400px;
        background: radial-gradient(circle, rgba(255,255,255,0.05) 0%, transparent 70%);
        border-radius: 50%;
    }
    .hero h1 {
        color: #FFFFFF;
        font-size: 2rem;
        font-weight: 700;
        margin: 0 0 0.4rem 0;
        letter-spacing: 0.5px;
    }
    .hero .subtitle {
        color: rgba(255,255,255,0.75);
        font-size: 0.95rem;
        margin: 0;
        line-height: 1.6;
    }
    .hero .badge {
        display: inline-block;
        background: rgba(255,255,255,0.15);
        border: 1px solid rgba(255,255,255,0.25);
        color: #FFFFFF;
        padding: 0.2rem 0.7rem;
        border-radius: 20px;
        font-size: 0.75rem;
        margin-top: 0.8rem;
        backdrop-filter: blur(4px);
    }
    
    /* ===== KPI 卡片 ===== */
    [data-testid="stMetric"] {
        background: linear-gradient(145deg, #FFFFFF, #F8FAFE);
        border: 1px solid #E3E8F0;
        border-radius: 12px;
        padding: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        transition: transform 0.2s, box-shadow 0.2s;
    }
    [data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 16px rgba(0,0,0,0.08);
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.8rem !important;
        color: #6B7280 !important;
        font-weight: 500 !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.4rem !important;
        font-weight: 700 !important;
        color: #1F2937 !important;
    }
    
    /* ===== 側邊欄 ===== */
    section[data-testid="stSidebar"] > div:first-child {
        background: linear-gradient(180deg, #F0F4F8 0%, #FFFFFF 100%);
    }
    .sidebar-title {
        background: linear-gradient(135deg, #0F2027, #2C5364);
        color: white;
        padding: 0.6rem 1rem;
        border-radius: 8px;
        font-size: 0.9rem;
        font-weight: 600;
        margin-bottom: 0.8rem;
        text-align: center;
    }
    .sidebar-section {
        background: white;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 0.8rem;
        margin-bottom: 0.8rem;
    }
    
    /* ===== 下載按鈕 ===== */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1F4E78, #2C5E3B) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.6rem 1.5rem !important;
        font-weight: 600 !important;
        transition: all 0.3s !important;
        box-shadow: 0 2px 8px rgba(31,78,120,0.3) !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 16px rgba(31,78,120,0.4) !important;
    }
    
    /* ===== 分區標題 ===== */
    .section-header {
        display: flex;
        align-items: center;
        gap: 0.6rem;
        padding: 0.8rem 0;
        margin: 0.5rem 0;
        border-bottom: 2px solid #E5E7EB;
    }
    .section-header .icon {
        font-size: 1.5rem;
    }
    .section-header .text {
        font-size: 1.15rem;
        font-weight: 700;
        color: #1F2937;
    }
    .section-header .desc {
        font-size: 0.82rem;
        color: #6B7280;
        margin-left: auto;
    }
    
    /* ===== 資訊卡 (空狀態) ===== */
    .empty-state {
        text-align: center;
        padding: 4rem 2rem;
        color: #9CA3AF;
    }
    .empty-state .icon { font-size: 4rem; margin-bottom: 1rem; }
    .empty-state .title { font-size: 1.2rem; font-weight: 600; color: #6B7280; }
    .empty-state .desc { font-size: 0.9rem; margin-top: 0.5rem; }
    
    /* ===== 狀態標籤 ===== */
    .status-bar {
        background: linear-gradient(90deg, #ECFDF5, #F0FDF4);
        border: 1px solid #A7F3D0;
        border-radius: 10px;
        padding: 0.7rem 1.2rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        margin: 0.8rem 0;
    }
    .status-bar .dot {
        width: 8px; height: 8px;
        background: #10B981;
        border-radius: 50%;
        animation: pulse 2s infinite;
    }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
    .status-bar .text {
        font-size: 0.88rem;
        color: #065F46;
        font-weight: 500;
    }
    
    /* ===== Tabs 標籤美化 ===== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
        background: #F3F4F6;
        border-radius: 12px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: white !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.1);
    }
    
    /* ===== 頁腳 ===== */
    .footer {
        text-align: center;
        padding: 1.5rem 0;
        color: #9CA3AF;
        font-size: 0.8rem;
        border-top: 1px solid #F3F4F6;
        margin-top: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ===== Hero 標題橫幅 =====
st.markdown("""
<div class="hero">
    <h1>🍂 雪茄批發報價自動彙整系統</h1>
    <p class="subtitle">
        上傳各家 Excel / PDF 報價單 → 自動辨識 30 個品牌 + 10 大供應商 → 即時比價矩陣 · 數據儀表板 · 一鍵下載精簡版 Excel
    </p>
    <span class="badge">⚡ v2.1 — 支援多檔上傳 · 品名搜尋 · 價格區間 · 最低價高亮</span>
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
    name = filename.split('.')[0]
    name = re.sub(r'\d{1,2}\.\d{1,2}', '', name)
    name = re.sub(r'20\d{2}', '', name)
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

    if ws_list.max_row > 1:
        tab_list = Table(displayName="DataList", ref=f"A1:E{ws_list.max_row}")
        tab_list.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws_list.add_table(tab_list)
    ws_list.freeze_panes = 'A2'

    # --- 工作表 2：品牌分類總覽統計 ---
    ws_sum = wb.create_sheet(title='品牌分類總覽統計')
    headers_sum = ['品牌名稱', '品項總報價數', '最低報價(HKD)', '最高報價(HKD)', '平均報價(HKD)']
    ws_sum.append(headers_sum)
    
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
# 4. 側邊欄：上傳 + 篩選
# ==========================================

# --- 側邊欄品牌 Logo ---
st.sidebar.markdown('<div class="sidebar-title">🍂 雪茄報價系統</div>', unsafe_allow_html=True)

# --- 上傳區 ---
st.sidebar.markdown("##### 📂 上傳報價單")
uploaded_files = st.sidebar.file_uploader(
    "拖曳或點擊上傳", 
    type=["xlsx", "xls", "pdf"], 
    accept_multiple_files=True,
    help="支援 Excel (.xlsx, .xls) 和 PDF 格式，可一次上傳多個檔案",
    label_visibility="collapsed"
)

if uploaded_files:
    st.sidebar.caption(f"📎 已選取 **{len(uploaded_files)}** 個檔案")
    with st.sidebar.expander("查看檔案清單", expanded=False):
        for f in uploaded_files:
            size_kb = f.size / 1024
            icon = "📗" if f.name.endswith(('.xlsx', '.xls')) else "📕"
            st.caption(f"{icon} {f.name} ({size_kb:.0f} KB)")

parse_button = st.sidebar.button("🚀 開始解析並合併轉檔", use_container_width=True, type="primary")

# ==========================================
# 5. 主畫面：空狀態 / 解析 / 結果
# ==========================================

# --- 空狀態引導 ---
if not uploaded_files and 'df_result' not in st.session_state:
    st.markdown("""
    <div class="empty-state">
        <div class="icon">📂</div>
        <div class="title">尚未上傳任何報價單</div>
        <div class="desc">👈 請在左側面板上傳 Excel 或 PDF 報價單，支援一次多檔上傳</div>
    </div>
    """, unsafe_allow_html=True)
    
    # 功能特色介紹
    st.markdown("---")
    feat_col1, feat_col2, feat_col3 = st.columns(3)
    with feat_col1:
        st.markdown("#### 🤖 智慧解析引擎")
        st.markdown("自動辨識 **30 個品牌** + **10 大供應商**，無論報價單格式如何混亂，都能精準萃取品名、價格與產地。")
    with feat_col2:
        st.markdown("#### 📊 即時比價矩陣")
        st.markdown("同一款雪茄在不同店家的報價橫向展開，**最低價自動高亮**，一眼看出誰家最便宜。")
    with feat_col3:
        st.markdown("#### 📥 一鍵匯出 Excel")
        st.markdown("下載的 Excel 已內建正式表格格式，直接就能使用 **交叉分析篩選器 (Slicer)**。")

# --- 解析執行 ---
if parse_button and uploaded_files:
    all_records = []
    file_stats = []
    
    progress_bar = st.progress(0, text="⏳ 正在解析檔案...")
    
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
            file_stats.append({"檔名": file.name, "供應商": vendor, "抓取筆數": 0, "狀態": f"❌ {e}"})
            
        progress_bar.progress((idx + 1) / len(uploaded_files), text=f"正在解析：{file.name} ({idx+1}/{len(uploaded_files)})")
    
    progress_bar.empty()
    st.session_state['file_stats'] = file_stats

    if all_records:
        df_result = pd.DataFrame(all_records)
        brand_order_map = {b: i for i, b in enumerate(TARGET_BRANDS)}
        brand_order_map['其他品牌'] = 999
        df_result['排序權重'] = df_result['品牌分類'].map(brand_order_map)
        df_result = df_result.sort_values(by=['排序權重', '品名規格', '價格'])
        df_result = df_result.drop_duplicates(subset=['供應商', '品牌分類', '品名規格']).drop(columns=['排序權重'])
        st.session_state['df_result'] = df_result
    else:
        st.warning("⚠️ 所有檔案均未成功解析出有效資料，請檢查檔案格式。")

# ==========================================
# 6. 結果展示區
# ==========================================
if 'df_result' in st.session_state:
    df_result = st.session_state['df_result']

    # --- 解析報告 ---
    if 'file_stats' in st.session_state:
        with st.expander("📋 解析報告：各檔案處理明細", expanded=False):
            stats_df = pd.DataFrame(st.session_state['file_stats'])
            st.dataframe(stats_df, use_container_width=True, hide_index=True)
            total_parsed = stats_df['抓取筆數'].sum()
            success_count = stats_df[stats_df['狀態'] == '✅ 成功'].shape[0]
            st.caption(f"共處理 {len(stats_df)} 個檔案 · {success_count} 個成功 · 原始抓取 {total_parsed} 筆 → 去重後 {len(df_result)} 筆")

    # --- 側邊欄篩選器 ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("##### 🔍 篩選器")
    
    all_brands = sorted(df_result['品牌分類'].unique())
    selected_brands = st.sidebar.multiselect("品牌分類", all_brands, placeholder="選擇品牌...")
    
    all_vendors = sorted(df_result['供應商'].unique())
    selected_vendors = st.sidebar.multiselect("供應商", all_vendors, placeholder="選擇供應商...")
    
    price_min = int(df_result['價格'].min())
    price_max = int(df_result['價格'].max())
    price_range = st.sidebar.slider(
        "💰 價格區間 (HKD)",
        min_value=price_min, max_value=price_max, 
        value=(price_min, price_max),
        step=50, format="$%d"
    )
    
    keyword = st.sidebar.text_input("🔎 品名搜尋", placeholder="短丘、D4、BHK...")
    
    if st.sidebar.button("🗑️ 清除篩選", use_container_width=True):
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
    
    # --- 狀態列 ---
    filter_tags = []
    if selected_brands: filter_tags.append(f"品牌: {', '.join(selected_brands)}")
    if selected_vendors: filter_tags.append(f"供應商: {', '.join(selected_vendors)}")
    if price_range != (price_min, price_max): filter_tags.append(f"價格: ${price_range[0]:,}~${price_range[1]:,}")
    if keyword: filter_tags.append(f"搜尋: {keyword}")
    
    filter_text = f" · 篩選條件：{'、'.join(filter_tags)}" if filter_tags else ""
    st.markdown(f"""
    <div class="status-bar">
        <div class="dot"></div>
        <span class="text">已載入 {len(df_result)} 筆資料 → 顯示 {len(df_display)} 筆{filter_text}</span>
    </div>
    """, unsafe_allow_html=True)
        
    # --- KPI 儀表板 ---
    st.markdown("""
    <div class="section-header">
        <span class="icon">📈</span>
        <span class="text">數據總覽</span>
        <span class="desc">即時更新 · 隨篩選條件連動</span>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("篩選品項", f"{len(df_display)} 筆")
    col2.metric("供應商數", f"{df_display['供應商'].nunique()} 家")
    col3.metric("品牌數", f"{df_display['品牌分類'].nunique()} 個")
    
    if not df_display.empty:
        col4.metric("最低價", f"${int(df_display['價格'].min()):,}")
        col5.metric("最高價", f"${int(df_display['價格'].max()):,}")
    else:
        col4.metric("最低價", "—")
        col5.metric("最高價", "—")
        
    # --- 下載區 ---
    today_str = datetime.now().strftime("%Y%m%d")
    dl_col1, dl_col2, dl_col3 = st.columns([2, 2, 3])
    
    with dl_col1:
        excel_bytes = generate_excel(df_display)
        st.download_button(
            label="📥 下載精簡版 Excel",
            data=excel_bytes,
            file_name=f"雪茄報價彙總_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with dl_col2:
        if not df_display.empty:
            pivot_dl = df_display.pivot_table(
                index=['品牌分類', '品名規格'], columns='供應商', values='價格', aggfunc='min'
            ).astype('Int64')
            output_pivot = io.BytesIO()
            with pd.ExcelWriter(output_pivot, engine='openpyxl') as writer:
                pivot_dl.to_excel(writer, sheet_name="跨店比價矩陣")
            st.download_button(
                label="📥 下載比價矩陣 Excel",
                data=output_pivot.getvalue(),
                file_name=f"跨店比價矩陣_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_pivot",
                use_container_width=True
            )
    
    with dl_col3:
        if not df_display.empty:
            avg_p = int(df_display['價格'].mean())
            st.markdown(f"<div style='text-align:center; padding:0.5rem; color:#6B7280; font-size:0.85rem;'>📊 平均報價 <b style=\"color:#1F2937; font-size:1.1rem;\">${avg_p:,} HKD</b></div>", unsafe_allow_html=True)
    
    # --- 圖表分析 ---
    with st.expander("📊 數據視覺化分析"):
        chart_col1, chart_col2 = st.columns(2)
        with chart_col1:
            st.markdown("**各品牌報價數量**")
            brand_counts = df_display['品牌分類'].value_counts()
            st.bar_chart(brand_counts)
        with chart_col2:
            st.markdown("**各供應商品項數**")
            vendor_counts = df_display['供應商'].value_counts()
            st.bar_chart(vendor_counts)
    
    # --- 主要資料區：分頁標籤 ---
    st.markdown("""
    <div class="section-header">
        <span class="icon">📋</span>
        <span class="text">資料檢視</span>
        <span class="desc">切換標籤查看不同視圖</span>
    </div>
    """, unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["📊 跨店比價矩陣", "📋 全品項清單"])
    
    with tab1:
        if not df_display.empty:
            st.caption("💡 綠底粗體 = 該品項各供應商中的最低報價")
            
            pivot_df = df_display.pivot_table(
                index=['品牌分類', '品名規格'], 
                columns='供應商', 
                values='價格', 
                aggfunc='min'
            )
            pivot_df = pivot_df.astype('Int64')
            
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
            st.caption(f"共 {len(df_display)} 筆資料")
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
st.sidebar.caption(f"🍂 v2.1 | {datetime.now().strftime('%Y-%m-%d %H:%M')}")
st.markdown(f"""
<div class="footer">
    🍂 雪茄批發報價自動彙整系統 v2.1 · Built with Streamlit · {datetime.now().strftime('%Y-%m-%d')}
</div>
""", unsafe_allow_html=True)
