import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher

EXPORT_SUM_COLS = ['Số tiền CK', 'Thành tiền chưa VAT', 'Thành tiền có VAT']
SORT_DATE_COL = '__sort_date'
EXTRA_FORMAT_COLS = ['Đơn giá', 'Thành tiền']

# ==========================================
# 1. CẤU HÌNH TRANG & BIẾN TOÀN CỤC
# ==========================================
st.set_page_config(page_title="Xử lý dữ liệu Super SM2057", layout="wide")
st.title("📊 CPC1 - Super SM2057 (NT)")
st.markdown(
    """
    <style>
    .st-key-filter_bar {
        position: fixed !important;
        bottom: 0px !important;
        left: 0px;
        right: 0px;
        z-index: 999 !important;
        background: white;
        padding: 1rem 3rem 1rem 3rem;
        box-shadow: 0px -5px 10px rgba(0,0,0,0.05);
        border-top: 1px solid #e0e0e0;
    }
    .st-key-filter_bar [data-testid="stHorizontalBlock"] {
        align-items: end;
        max-width: 1200px;
        margin: 0 auto;
    }
    /* Tăng khoảng không bên dưới để bảng không bị che khuất bởi thanh fixed */
    .main .block-container, [data-testid="stMainBlockContainer"] {
        padding-bottom: 15rem !important;
    }
    .st-key-filter_reset_btn button {
        color: #16a34a;
        border: 1px solid #86efac;
        background: #f0fdf4;
        font-size: 1.2rem;
        font-weight: 700;
        min-height: 2.5rem;
    }
    .st-key-filter_reset_btn button:hover {
        border-color: #22c55e;
        background: #dcfce7;
        color: #15803d;
    }
    .st-key-clear_upload_btn_filter,
    .st-key-clear_upload_btn_merge {
        margin-top: 2.3rem;
    }
    .st-key-clear_upload_btn_filter button,
    .st-key-clear_upload_btn_merge button {
        color: #dc2626;
        border: 1px solid #fecaca;
        background: #fef2f2;
        font-size: 1.1rem;
        font-weight: 700;
        min-height: 3rem;
        padding: 0 0.5rem;
        border-radius: 0.5rem;
    }
    .st-key-clear_upload_btn_filter button:hover,
    .st-key-clear_upload_btn_merge button:hover {
        border-color: #f87171;
        background: #fee2e2;
        color: #b91c1c;
    }
    button[data-baseweb="tab"] {
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True
)
# Danh sách các cột chứa dữ liệu số cần tính toán
GLOBAL_CALC_COLS = ['Số lượng', 'Đơn giá chưa VAT', 'Đơn giá có VAT', 'Số tiền CK', 'Thành tiền chưa VAT', 'Thành tiền có VAT']

# ==========================================
# 2. CÁC HÀM XỬ LÝ & RENDER DỮ LIỆU
# ==========================================
def doc_so_thanh_chu(n):
    """Hàm đọc số tiền ra chữ tiếng Việt chuẩn xác"""
    if pd.isna(n) or n == "": return ""
    try:
        n = int(n)
    except:
        return ""
    if n == 0: return "Không đồng"
    if n < 0: return "Âm " + doc_so_thanh_chu(-n).lower()
    
    units = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"]
    words = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    
    def read_3_digits(num, read_zero_hundred=False):
        h = num // 100
        t = (num % 100) // 10
        u = num % 10
        res = []
        if h > 0 or read_zero_hundred:
            res.extend([words[h], "trăm"])
        if t > 1:
            res.extend([words[t], "mươi"])
            if u == 1: res.append("mốt")
            elif u == 4: res.append("tư")
            elif u == 5: res.append("lăm")
            elif u > 0: res.append(words[u])
        elif t == 1:
            res.append("mười")
            if u == 5: res.append("lăm")
            elif u > 0: res.append(words[u])
        else:
            if u > 0:
                if h > 0 or read_zero_hundred:
                    res.append("lẻ")
                res.append(words[u])
        return " ".join(res)

    chunks = []
    temp = n
    while temp > 0:
        chunks.append(temp % 1000)
        temp //= 1000
        
    res = []
    for i, chunk in enumerate(chunks):
        if chunk > 0:
            read_zero = False
            if i < len(chunks) - 1 and sum(chunks[i+1:]) > 0 and chunk < 100:
                read_zero = True
            
            chunk_str = read_3_digits(chunk, read_zero_hundred=read_zero)
            unit = units[i] if i < len(units) else ""
            res.insert(0, (chunk_str + " " + unit).strip())
            
    final_str = " ".join(res).strip()
    final_str = final_str.replace("  ", " ")
    return final_str.capitalize() + " đồng"

def parse_invoice_dates(date_series):
    parsed_dates = pd.to_datetime(date_series, format='%d/%m/%Y', errors='coerce')
    return parsed_dates.fillna(pd.to_datetime(date_series, dayfirst=True, errors='coerce'))

def get_file_month_sort_key(df, date_column, sample_size=10):
    if not date_column or date_column not in df.columns:
        return (9999, 12)
    sampled_dates = df[date_column].dropna().head(sample_size)
    parsed_dates = parse_invoice_dates(sampled_dates).dropna()
    if parsed_dates.empty:
        return (9999, 12)
    min_date = parsed_dates.min()
    return (int(min_date.year), int(min_date.month))

def normalize_text(value):
    text = "" if pd.isna(value) else str(value)
    text = re.sub(r'\s+', ' ', text).strip()
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(char for char in text if not unicodedata.combining(char))
    return text.lower()

def reset_filter_state():
    st.session_state["filter_selected_kh"] = "Tất cả|||tat ca"
    st.session_state["filter_selected_hd"] = "Tất cả"
    st.session_state["filter_selected_months"] = []

def reset_month_filter():
    st.session_state["filter_selected_months"] = []

def clear_uploaded_filter_data():
    st.session_state["filter_selected_kh"] = "Tất cả|||tat ca"
    st.session_state["filter_selected_hd"] = "Tất cả"
    st.session_state["filter_selected_months"] = []
    st.session_state["filter_uploader_nonce"] = st.session_state.get("filter_uploader_nonce", 0) + 1

def clear_uploaded_merge_data():
    st.session_state["merge_uploader_nonce"] = st.session_state.get("merge_uploader_nonce", 0) + 1
    st.session_state.pop("merge_cache_fingerprint", None)
    st.session_state.pop("merge_cached_df", None)
    st.session_state.pop("merge_cached_errors", None)
    st.session_state.pop("merge_cached_date_col", None)

def get_uploaded_files_fingerprint(uploaded_files):
    return tuple((file.name, file.size) for file in uploaded_files)

def sanitize_filename(filename):
    filename = re.sub(r'[<>:"/\\|?*]+', '_', str(filename).strip())
    filename = re.sub(r'\s+', '_', filename)
    return filename or "du_lieu"

def ensure_xlsx_extension(filename):
    filename = sanitize_filename(filename)
    return filename if filename.lower().endswith('.xlsx') else f"{filename}.xlsx"

def standardize_header_name(value, fallback_prefix="Cột"):
    text = re.sub(r'\s+', ' ', str(value).strip())
    if not text:
        text = fallback_prefix
    return ' '.join(word[:1].upper() + word[1:].lower() for word in text.split())

def make_unique_columns(columns):
    seen = {}
    unique_columns = []
    for index, col in enumerate(columns, start=1):
        base_name = standardize_header_name(col, fallback_prefix=f"Cột {index}")
        count = seen.get(base_name, 0)
        unique_name = base_name if count == 0 else f"{base_name}_{count + 1}"
        seen[base_name] = count + 1
        unique_columns.append(unique_name)
    return unique_columns

def find_header_row(preview_df, keywords=None):
    if preview_df.empty:
        return 0
    if keywords:
        preview_text = preview_df.fillna('').astype(str)
        for idx, row in preview_text.iterrows():
            row_values = row.str.lower().str.strip().tolist()
            if any(any(kw in cell for kw in keywords) for cell in row_values):
                return idx
    non_null_counts = preview_df.notna().sum(axis=1)
    return int(non_null_counts.idxmax())

def safe_read_excel(file, header=0, nrows=None):
    file.seek(0)
    try:
        return pd.read_excel(file, header=header, nrows=nrows, engine='calamine')
    except Exception:
        file.seek(0)
        return pd.read_excel(file, header=header, nrows=nrows)

def collapse_duplicate_columns(df):
    collapsed_data = {}
    for col in pd.unique(df.columns):
        same_name_columns = df.loc[:, df.columns == col]
        if same_name_columns.shape[1] == 1:
            collapsed_data[col] = same_name_columns.iloc[:, 0]
        else:
            collapsed_data[col] = same_name_columns.bfill(axis=1).iloc[:, 0]
    return pd.DataFrame(collapsed_data)

def find_best_matching_header(column_name, base_headers, threshold=0.85):
    normalized_column = normalize_text(column_name)
    best_match = None
    best_score = 0
    for base_header in base_headers:
        score = SequenceMatcher(None, normalized_column, normalize_text(base_header)).ratio()
        if score > best_score:
            best_score = score
            best_match = base_header
    if best_score >= threshold:
        return best_match
    return None

def find_invoice_date_column(columns):
    keywords = ['ngay hoa don', 'ngay hoa don gtgt']
    best_match = None
    best_score = 0
    for column in columns:
        normalized_column = normalize_text(column)
        if 'ngay' in normalized_column and 'hoa' in normalized_column and 'don' in normalized_column:
            return column
        for keyword in keywords:
            score = SequenceMatcher(None, normalized_column, keyword).ratio()
            if score > best_score:
                best_score = score
                best_match = column
    if best_score >= 0.6:
        return best_match
    return None

def read_merge_source_file(file):
    preview_df = safe_read_excel(file, header=None, nrows=30)
    header_idx = find_header_row(preview_df)
    df = safe_read_excel(file, header=header_idx)
    df = df.dropna(axis=0, how='all')
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', case=False, regex=True)]
    df.columns = make_unique_columns(df.columns)
    df = collapse_duplicate_columns(df)
    return df.reset_index(drop=True)

def align_dataframe_columns(df, base_headers, threshold=0.85):
    rename_map = {}
    for column in df.columns:
        matched_header = find_best_matching_header(column, base_headers, threshold=threshold)
        rename_map[column] = matched_header or column

    df = df.rename(columns=rename_map)
    df = collapse_duplicate_columns(df)

    for column in df.columns:
        if column not in base_headers:
            base_headers.append(column)

    return df.reindex(columns=base_headers, fill_value=pd.NA), base_headers

def merge_sm2057_files(uploaded_files, progress_bar=None, status_placeholder=None):
    merged_frames = []
    base_headers = []
    errors = []

    for index, file in enumerate(uploaded_files, start=1):
        if progress_bar is not None:
            progress_bar.progress(index / len(uploaded_files), text=f"Đang xử lý {index}/{len(uploaded_files)} file")
        if status_placeholder is not None:
            status_placeholder.info(f"Đang đọc: {file.name}")

        try:
            current_df = read_merge_source_file(file)
            current_df['Nguồn_File_Gốc'] = file.name
            date_column = find_invoice_date_column(current_df.columns)
            month_sort_key = get_file_month_sort_key(current_df, date_column)

            if not base_headers:
                base_headers = list(current_df.columns)
            else:
                current_df, base_headers = align_dataframe_columns(current_df, base_headers)

            merged_frames.append((month_sort_key, index, current_df.reindex(columns=base_headers, fill_value=pd.NA)))
        except Exception as exc:
            errors.append(f"{file.name}: {exc}")

    if not merged_frames:
        return None, errors, None

    merged_frames.sort(key=lambda item: (item[0][0], item[0][1], item[1]))
    merged_df = pd.concat([item[2] for item in merged_frames], ignore_index=True)
    merged_df = merged_df.dropna(axis=0, how='all').reset_index(drop=True)

    date_column = find_invoice_date_column(merged_df.columns)
    return merged_df, errors, date_column

def build_agg_rules(columns, group_col):
    agg_rules = {}
    for col in columns:
        if col in GLOBAL_CALC_COLS:
            agg_rules[col] = 'sum'
        elif col != group_col and col != 'KH_Display' and col != SORT_DATE_COL:
            agg_rules[col] = 'first'
    return agg_rules

def build_bang_08a(source_df):
    group_col = 'Số chứng từ ngoại'
    date_col = 'Ngày hóa đơn'
    amount_candidates = ['Thành tiền có VAT', 'Thành tiền chưa VAT']

    if group_col not in source_df.columns or date_col not in source_df.columns:
        return pd.DataFrame(columns=[
            'Số chứng từ ngoại', 'Ngày hóa đơn', 'Nội dung công việc',
            'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền'
        ])

    amount_col = next((col for col in amount_candidates if col in source_df.columns), None)
    grouped = source_df.groupby(group_col, as_index=False, sort=False).agg({
        date_col: 'first',
        SORT_DATE_COL: 'first' if SORT_DATE_COL in source_df.columns else (lambda x: pd.NaT)
    })

    if amount_col:
        amount_df = source_df.groupby(group_col, as_index=False, sort=False)[amount_col].sum()
        grouped = grouped.merge(amount_df, on=group_col, how='left')
        grouped['Thành tiền'] = grouped[amount_col]
        grouped = grouped.drop(columns=[amount_col])
    else:
        grouped['Thành tiền'] = 0

    grouped['Nội dung công việc'] = (
        "Hoá đơn số " + grouped[group_col].astype(str) + " ngày " + grouped[date_col].astype(str)
    )
    grouped['Đơn vị tính'] = ""
    grouped['Số lượng'] = ""
    grouped['Đơn giá'] = ""

    return grouped[[
        'Số chứng từ ngoại', 'Ngày hóa đơn', 'Nội dung công việc',
        'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền'
    ]]

@st.cache_data(show_spinner=False)
def aggregate_dataframe(source_df, group_col, keep_sort_order=False):
    agg_rules = build_agg_rules(source_df.columns, group_col)
    return source_df.groupby(group_col, as_index=False, sort=not keep_sort_order).agg(agg_rules)

@st.cache_data(show_spinner=False)
def build_export_excel(export_df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Ket_qua')
    return buffer.getvalue()

@st.cache_data(show_spinner=False)
def process_uploaded_files(uploaded_files):
    """Đọc và gộp nhiều file Excel/CSV thành 1 DataFrame duy nhất"""
    dfs = []
    
    for file in uploaded_files:
        try:
            # --- BƯỚC 1 & BƯỚC 2: Dò tìm dòng Header và đọc file ---
            if file.name.endswith('.csv'):
                file.seek(0)
                df_preview = pd.read_csv(file, header=None, nrows=30)
                header_idx = find_header_row(df_preview, keywords=["mã kh hóa đơn", "mã kh hoá đơn"])
                file.seek(0)
                df = pd.read_csv(file, header=header_idx)
            elif file.name.endswith(('.xls', '.xlsx')):
                df_preview = safe_read_excel(file, header=None, nrows=30)
                header_idx = find_header_row(df_preview, keywords=["mã kh hóa đơn", "mã kh hoá đơn"])
                df = safe_read_excel(file, header=header_idx)
            else:
                continue
                
            # Loại bỏ cột rác
            df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
            month_sort_key = get_file_month_sort_key(df, 'Ngày hóa đơn' if 'Ngày hóa đơn' in df.columns else None)
            dfs.append((month_sort_key, len(dfs), df))
            
        except Exception as e:
            st.error(f"Lỗi khi đọc file {file.name}: {e}")
            
    if dfs:
        dfs.sort(key=lambda item: (item[0][0], item[0][1], item[1]))
        master_df = pd.concat([item[2] for item in dfs], ignore_index=True)
        
        # Chuẩn hóa tên cột để xử lý lỗi gõ dấu (Hóa vs Hoá)
        master_df.columns = (
            master_df.columns.astype(str)
            .str.strip()
            .str.replace('oá', 'óa', regex=False)
            .str.replace('Oá', 'Óa', regex=False)
        )
        
        # Chỉ parse ngày để dùng cho các bộ lọc/bảng khác, không sort lại toàn bộ dòng
        if 'Ngày hóa đơn' in master_df.columns:
            master_df[SORT_DATE_COL] = parse_invoice_dates(master_df['Ngày hóa đơn'])
        
        # Cột hiển thị khách hàng
        if 'Mã KH hóa đơn' in master_df.columns and 'Tên KH hóa đơn' in master_df.columns:
            master_df['KH_Code'] = master_df['Mã KH hóa đơn'].astype(str).str.strip()
            master_df['KH_Search'] = master_df['Tên KH hóa đơn'].astype(str)
            master_df['KH_Display'] = master_df['Mã KH hóa đơn'].astype(str) + " - " + master_df['Tên KH hóa đơn'].astype(str)
        else:
            master_df['KH_Code'] = "Không xác định"
            master_df['KH_Search'] = "Không xác định"
            master_df['KH_Display'] = "Không xác định - Không xác định"
            
        if 'Số hợp đồng' in master_df.columns:
            master_df['Số hợp đồng'] = master_df['Số hợp đồng'].fillna("Không xác định")
        else:
            master_df['Số hợp đồng'] = "Không xác định"
            
        # Xử lý Nhà SX (Cắt phần đuôi sau dấu _)
        if 'Nhà SX' in master_df.columns:
            master_df['Nhà SX'] = master_df['Nhà SX'].fillna('')
            master_df['Nhà SX'] = master_df['Nhà SX'].astype(str).apply(
                lambda x: x.split('_', 1)[-1].strip() if '_' in x else x
            )
            
        # Xử lý định dạng số VN khi nhập liệu
        for col in GLOBAL_CALC_COLS:
            if col in master_df.columns:
                if master_df[col].dtype == 'object':
                    master_df[col] = master_df[col].astype(str).str.replace(r'\.', '', regex=True).str.replace(',', '.', regex=False)
                master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0)
                
        return master_df
    return None

def render_result_table(final_df, selected_columns, tab_key, show_total_text=True, default_file_name=None):
    """Hàm tạo bảng kết quả, đánh STT, tính tổng và xuất Excel"""
    if not selected_columns:
        st.warning("Vui lòng chọn ít nhất 1 cột để hiển thị kết quả.")
        return

    available_columns = [col for col in selected_columns if col in final_df.columns and col != SORT_DATE_COL]
    display_df = final_df.loc[:, available_columns].copy()
    
    if 'STT' in display_df.columns:
        display_df = display_df.drop(columns=['STT'])
    display_df.insert(0, 'STT', range(1, len(display_df) + 1))
    
    # Tính dòng TỔNG CỘNG
    total_row = {col: "" for col in display_df.columns}
    text_col = next((col for col in display_df.columns if col not in GLOBAL_CALC_COLS and col != 'STT'), None)
    
    if text_col: total_row[text_col] = "TỔNG CỘNG"
    else: total_row['STT'] = "TỔNG"
        
    total_display_value = 0
    total_display_label = None
    # Chỉ tính tổng dòng cuối cho các cột Thành tiền và Tiền Chiết Khấu (Bỏ qua Số lượng, Đơn giá)
    for col in EXPORT_SUM_COLS:
        if col in display_df.columns:
            val = display_df[col].sum()
            total_row[col] = val
            if col == 'Thành tiền có VAT':
                total_display_value = val
                total_display_label = 'Tổng Thành tiền có VAT'

    if total_display_label is None and 'Thành tiền' in display_df.columns:
        total_display_value = display_df['Thành tiền'].sum()
        total_display_label = 'Tổng Thành tiền'
    
    export_df = display_df.copy()
    export_df = export_df.fillna("")
    
    # --- ĐỊNH DẠNG UI (Visual DataFrame) ---
    def format_vn_number(x):
        if pd.isna(x) or x == "": return ""
        try:
            return f"{float(x):,.0f}".replace(",", ".")
        except:
            return x
            
    visual_df = export_df.copy()
    for col in GLOBAL_CALC_COLS + EXTRA_FORMAT_COLS:
        if col in visual_df.columns:
            visual_df[col] = visual_df[col].apply(format_vn_number)
    
    # Dùng Pandas Styler để căn lề phải cho cột số trên Web
    
    # Bảng hiển thị Web (autosize width)
    st.dataframe(visual_df, width='stretch', hide_index=True)
    
    # Hiển thị số tiền bằng chữ
    if show_total_text and total_display_label is not None:
        def fmt_vn(x): return f"{float(x):,.0f}".replace(",", ".") if pd.notna(x) else "0"
        st.markdown(f"**💰 {total_display_label}:** <span style='color:red; font-size: 1.1em;'>{fmt_vn(total_display_value)}</span> VNĐ", unsafe_allow_html=True)
        st.markdown(f"**✍️ Bằng chữ:** *{doc_so_thanh_chu(total_display_value)}*")
    
    # --- XUẤT FILE EXCEL (.XLSX) ---
    file_name_value = st.text_input(
        "Tên file tải về",
        value=ensure_xlsx_extension(default_file_name or f'ket_qua_loc_{tab_key}'),
        key=f"download_name_{tab_key}"
    )
    st.download_button(
        label="📥 Tải kết quả xuống (Excel - .xlsx)",
        data=build_export_excel(export_df),
        file_name=ensure_xlsx_extension(file_name_value),
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key=f"dl_btn_{tab_key}"
    )

main_tab_filter, main_tab_merge = st.tabs(["📊 Lọc Super SM2057 (NT)", "🧩 Gộp file SM2057"])

with main_tab_filter:
    # ==========================================
    # 3. GIAO DIỆN TẢI FILE & LỌC DỮ LIỆU CHUNG
    # ==========================================
    st.header("1. Tải dữ liệu lên")
    if "filter_uploader_nonce" not in st.session_state:
        st.session_state["filter_uploader_nonce"] = 0
    upload_col, clear_col = st.columns([1, 0.06])
    with upload_col:
        uploaded_files = st.file_uploader(
            "Kéo thả, thêm nhiều file SM2057 (NT)",
            type=['csv', 'xlsx', 'xls'],
            accept_multiple_files=True,
            key=f"filter_uploader_{st.session_state['filter_uploader_nonce']}"
        )
    with clear_col:
        clear_upload_container = st.container(key="clear_upload_btn_filter")
        with clear_upload_container:
            st.button(
                "↻",
                key="clear_uploaded_files_btn",
                width="stretch",
                help="Xóa dữ liệu đã tải",
                on_click=clear_uploaded_filter_data
            )

    if uploaded_files:
        with st.spinner('Đang đọc và gộp dữ liệu...'):
            df = process_uploaded_files(uploaded_files)
            
        if df is not None:
            st.success(f"✅ Đã gộp thành công {len(uploaded_files)} file. Tổng số dòng dữ liệu: {len(df)}")

            st.header("2. Lọc khách hàng & Hợp đồng")

            filter_bar = st.container(key="filter_bar")
            with filter_bar:
                col1, col2, col3, col4 = st.columns([1.2, 1.2, 1, 0.22])
                with col1:
                    kh_lookup = (
                        df[['KH_Code', 'KH_Display', 'KH_Search']]
                        .dropna(subset=['KH_Code'])
                        .groupby('KH_Code', as_index=False)
                        .agg({'KH_Display': 'first', 'KH_Search': 'first'})
                        .sort_values(by='KH_Display')
                    )
                    kh_options = ["Tất cả|||tat ca"] + [
                        f"{row.KH_Code}|||{row.KH_Display}|||{normalize_text(row.KH_Search)}"
                        for row in kh_lookup.itertuples(index=False)
                    ]
                    selected_kh_option = st.selectbox(
                        "1️⃣ Chọn khách hàng:",
                        kh_options,
                        format_func=lambda value: "Tất cả" if value == "Tất cả|||tat ca" else value.split("|||", 2)[1],
                        key="filter_selected_kh",
                        on_change=reset_month_filter
                    )
                    selected_kh = "Tất cả" if selected_kh_option == "Tất cả|||tat ca" else selected_kh_option.split("|||", 2)[0]

                filtered_df = df[df['KH_Code'] == selected_kh] if selected_kh != "Tất cả" else df

                with col2:
                    hd_list = ["Tất cả"] + list(filtered_df['Số hợp đồng'].dropna().unique())
                    selected_hd = st.selectbox("2️⃣ Chọn Số Hợp Đồng:", hd_list, key="filter_selected_hd")

                if selected_hd != "Tất cả":
                    filtered_df = filtered_df[filtered_df['Số hợp đồng'] == selected_hd]

                month_options = []
                if SORT_DATE_COL in filtered_df.columns:
                    month_values = filtered_df[SORT_DATE_COL].dropna().dt.strftime('%m/%Y').unique().tolist()
                    month_options = sorted(month_values, key=lambda x: (int(x[3:]), int(x[:2])))

                with col3:
                    selected_months = st.multiselect(
                        "3️⃣ Chọn tháng: (Mặc định tất cả)",
                        month_options,
                        placeholder="Chọn Tháng nếu muốn",
                        key="filter_selected_months"
                    )

                with col4:
                    st.write("")
                    st.button(
                        "↻",
                        key="filter_reset_btn",
                        width='stretch',
                        help="Reset bộ lọc",
                        on_click=reset_filter_state
                    )

                if selected_months and SORT_DATE_COL in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df[SORT_DATE_COL].dt.strftime('%m/%Y').isin(selected_months)]

            st.caption(f"🔍 Dữ liệu hiện tại: **{len(filtered_df)}** dòng khớp điều kiện.")

            tab1, tab2, tab3, tab4 = st.tabs(["📦 Bảng Chi tiết Hoá đơn - Hàng", "🛒 Bảng Hàng hoá", "🧾 Bảng Hóa đơn", "📄 Bảng 08a"])
            
            with tab1:
                available_columns_t1 = [col for col in df.columns if col not in ('KH_Display', 'KH_Search', 'KH_Code', SORT_DATE_COL) and str(col).upper() != 'STT']
                desired_defaults_t1 = ['Ngày hóa đơn', 'Số chứng từ ngoại', 'Tên HH', 'ĐVT', 'Số lượng', 'Đơn giá có VAT', 'Thành tiền có VAT', 'Nhà SX']
                default_cols_t1 = [col for col in desired_defaults_t1 if col in available_columns_t1]
                if not default_cols_t1:
                    default_cols_t1 = available_columns_t1[:8]

                with st.expander("Cột hiển thị", expanded=False):
                    selected_columns_t1 = st.multiselect(
                        "Cột hiển thị (có thể chuyển vị trí cột ở bảng)",
                        options=available_columns_t1,
                        default=default_cols_t1,
                        key="multi_t1"
                    )

                final_df_t1 = filtered_df
                st.write("")
                default_file_name_t1 = f"chi_tiet_hoa_don_hang_{selected_kh}_{selected_hd or 'tat_ca'}"
                render_result_table(final_df_t1, selected_columns_t1, "chi_tiet_hoa_don_hang", default_file_name=default_file_name_t1)

            with tab2:
                available_columns_t2 = [col for col in df.columns if col not in ('KH_Display', 'KH_Search', 'KH_Code', SORT_DATE_COL) and str(col).upper() != 'STT']
                desired_defaults_t2 = ['Tên HH', 'ĐVT', 'Số lượng', 'Đơn giá có VAT', 'Thành tiền có VAT', 'Nhà SX']
                default_cols_t2 = [col for col in desired_defaults_t2 if col in available_columns_t2]
                if not default_cols_t2:
                    default_cols_t2 = available_columns_t2[:6]

                with st.expander("Cột hiển thị", expanded=False):
                    selected_columns_t2 = st.multiselect(
                        "Cột hiển thị (có thể chuyển vị trí cột ở bảng)",
                        options=available_columns_t2,
                        default=default_cols_t2,
                        key="multi_t2"
                    )

                final_df_t2 = filtered_df
                group_col_t2 = 'Mã HH'
                if group_col_t2 in final_df_t2.columns:
                    final_df_t2 = aggregate_dataframe(final_df_t2, group_col_t2)

                st.write("")
                default_file_name_t2 = f"bang_hang_hoa_{selected_kh}_{selected_hd or 'tat_ca'}"
                render_result_table(final_df_t2, selected_columns_t2, "bang_hang_hoa", default_file_name=default_file_name_t2)

            with tab3:
                available_columns_t3 = [col for col in df.columns if col not in ('KH_Display', 'KH_Search', 'KH_Code', SORT_DATE_COL) and str(col).upper() != 'STT']
                desired_defaults_t3 = ['Ngày hóa đơn', 'Số chứng từ ngoại', 'Thành tiền có VAT']
                default_cols_t3 = [col for col in desired_defaults_t3 if col in available_columns_t3]
                if not default_cols_t3:
                    default_cols_t3 = available_columns_t3[:3]

                with st.expander("Cột hiển thị", expanded=False):
                    selected_columns_t3 = st.multiselect(
                        "Cột hiển thị (có thể chuyển vị trí cột ở bảng)",
                        options=available_columns_t3,
                        default=default_cols_t3,
                        key="multi_t3"
                    )

                final_df_t3 = filtered_df
                group_col_t3 = 'Số chứng từ ngoại'
                if group_col_t3 in final_df_t3.columns:
                    final_df_t3 = aggregate_dataframe(final_df_t3, group_col_t3, keep_sort_order=True)

                st.write("")
                default_file_name_t3 = f"danh_sach_hoa_don_{selected_kh}_{selected_hd or 'tat_ca'}"
                render_result_table(final_df_t3, selected_columns_t3, "danh_sach_hoa_don", default_file_name=default_file_name_t3)

            with tab4:
                final_df_t4 = build_bang_08a(filtered_df)
                available_columns_t4 = [
                    'Số chứng từ ngoại', 'Ngày hóa đơn', 'Nội dung công việc',
                    'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền'
                ]

                with st.expander("Cột hiển thị", expanded=False):
                    selected_columns_t4 = st.multiselect(
                        "Cột hiển thị (có thể chuyển vị trí cột ở bảng)",
                        options=available_columns_t4,
                        default=available_columns_t4,
                        key="multi_t4"
                    )
                st.write("")
                default_file_name_t4 = f"bang_08a_{selected_kh}_{selected_hd or 'tat_ca'}"
                render_result_table(final_df_t4, selected_columns_t4, "bang_08a", default_file_name=default_file_name_t4)

with main_tab_merge:
    st.header("Làm sạch & Gộp file Excel giống nhau")
    if "merge_uploader_nonce" not in st.session_state:
        st.session_state["merge_uploader_nonce"] = 0
    merge_upload_col, merge_clear_col = st.columns([1, 0.06])
    with merge_upload_col:
        merge_files = st.file_uploader(
            "Kéo thả nhiều file Excel vào đây (.xlsx, .xls)",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            key=f"merge_uploader_{st.session_state['merge_uploader_nonce']}"
        )
    with merge_clear_col:
        clear_upload_container = st.container(key="clear_upload_btn_merge")
        with clear_upload_container:
            st.button(
                "↻",
                key="clear_uploaded_merge_files_btn",
                width="stretch",
                help="Xóa dữ liệu đã tải",
                on_click=clear_uploaded_merge_data
            )

    if merge_files:
        merge_fingerprint = get_uploaded_files_fingerprint(merge_files)
        cache_key = "merge_cache_fingerprint"
        if st.session_state.get(cache_key) != merge_fingerprint:
            progress_bar = st.progress(0, text="Sẵn sàng xử lý")
            status_placeholder = st.empty()
            merged_df, merge_errors, detected_date_col = merge_sm2057_files(
                merge_files,
                progress_bar=progress_bar,
                status_placeholder=status_placeholder
            )

            progress_bar.progress(1.0, text="Đã hoàn tất xử lý")
            status_placeholder.empty()
            st.session_state[cache_key] = merge_fingerprint
            st.session_state["merge_cached_df"] = merged_df
            st.session_state["merge_cached_errors"] = merge_errors
            st.session_state["merge_cached_date_col"] = detected_date_col
        else:
            merged_df = st.session_state.get("merge_cached_df")
            merge_errors = st.session_state.get("merge_cached_errors", [])
            detected_date_col = st.session_state.get("merge_cached_date_col")

        if merged_df is not None:
            st.success(f"✅ Đã gộp thành công {len(merge_files)} file.")
            st.caption(f"📌 Tổng số dòng dữ liệu thu thập được: **{len(merged_df)}**")
            if detected_date_col:
                st.caption(f"🗓️ Đã phát hiện cột ngày hóa đơn để sắp xếp: **{detected_date_col}**")

            if merge_errors:
                st.warning("Một số file không đọc được:")
                for error in merge_errors:
                    st.write(f"- {error}")

            merge_download_name = st.text_input(
                "Tên file tải về",
                value=ensure_xlsx_extension(f"DuLieu_DaGop_{datetime.now().strftime('%Y-%m-%d')}"),
                key="download_name_merged_sm2057"
            )
            st.download_button(
                label="📥 Tải về DuLieu_DaGop.xlsx",
                data=build_export_excel(merged_df),
                file_name=ensure_xlsx_extension(merge_download_name),
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key="download_merged_sm2057"
            )

            st.subheader("Xem trước dữ liệu")
            st.dataframe(merged_df.head(50), width=1500, hide_index=True)
        else:
            st.error("Không thể gộp dữ liệu từ các file đã tải lên.")
            if merge_errors:
                for error in merge_errors:
                    st.write(f"- {error}")
    else:
        st.session_state.pop("merge_cache_fingerprint", None)
        st.session_state.pop("merge_cached_df", None)
        st.session_state.pop("merge_cached_errors", None)
        st.session_state.pop("merge_cached_date_col", None)

# ==========================================
# 5. FOOTER (CREDIT)
# ==========================================
st.markdown(
    """
    <div style="text-align: center; margin-top: 60px; padding-bottom: 20px; color: #888888; font-size: 14px;">
        Thiết kế: Nguyễn Văn Dũng ❤️ 0978.777.191
    </div>
    """, 
    unsafe_allow_html=True
)
