import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher

# ==========================================
# 1. CẤU HÌNH TRANG, BIẾN TOÀN CỤC & MAP TÊN CỘT
# ==========================================

# Gom nhóm tên cột để chống Hardcode. 
# Nếu file Excel nguồn đổi tên cột, chỉ cần đổi ở đây!
class COL:
    MA_KH = 'Mã KH hóa đơn'
    TEN_KH = 'Tên KH hóa đơn'
    SO_HD = 'Số hợp đồng'
    NGAY_HD = 'Ngày hóa đơn'
    SO_CT = 'Số chứng từ ngoại'
    NHA_SX = 'Nhà SX'
    MA_HH = 'Mã HH'
    TEN_HH = 'Tên HH'
    DVT = 'ĐVT'
    SO_LUONG = 'Số lượng'
    DON_GIA = 'Đơn giá'
    DON_GIA_CHUA_VAT = 'Đơn giá chưa VAT'
    DON_GIA_CO_VAT = 'Đơn giá có VAT'
    TIEN_CK = 'Số tiền CK'
    THANH_TIEN = 'Thành tiền'
    THANH_TIEN_CHUA_VAT = 'Thành tiền chưa VAT'
    THANH_TIEN_CO_VAT = 'Thành tiền có VAT'
    NOI_DUNG = 'Nội dung công việc'
    
    # Các cột nội bộ sinh ra trong quá trình xử lý
    KH_CODE = 'KH_Code'
    KH_SEARCH = 'KH_Search'
    KH_DISPLAY = 'KH_Display'
    SORT_DATE = '__sort_date'
    NGUON_FILE = 'Nguồn_File_Gốc'
    STT = 'STT'

GLOBAL_CALC_COLS = [COL.SO_LUONG, COL.DON_GIA_CHUA_VAT, COL.DON_GIA_CO_VAT, COL.TIEN_CK, COL.THANH_TIEN_CHUA_VAT, COL.THANH_TIEN_CO_VAT]
EXPORT_SUM_COLS = [COL.TIEN_CK, COL.THANH_TIEN_CHUA_VAT, COL.THANH_TIEN_CO_VAT]
EXTRA_FORMAT_COLS = [COL.DON_GIA, COL.THANH_TIEN]

def setup_page_config():
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
        button[data-baseweb="tab"] { font-weight: 700; }
        pre {
            max-height: 48px !important;
            background-color: #4e4e4e !important;
            color: #d4d4d4 !important;
        }
        code {
            color: #ffcc00 !important;
            background-color: transparent !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# ==========================================
# 2. CÁC HÀM XỬ LÝ DỮ LIỆU & TIỆN ÍCH
# ==========================================
def doc_so_thanh_chu(n):
    if pd.isna(n) or n == "": return ""
    try: n = int(n)
    except: return ""
    if n == 0: return "Không đồng"
    if n < 0: return "Âm " + doc_so_thanh_chu(-n).lower()
    
    units = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"]
    words = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    
    def read_3_digits(num, read_zero_hundred=False):
        h, t, u = num // 100, (num % 100) // 10, num % 10
        res = []
        if h > 0 or read_zero_hundred: res.extend([words[h], "trăm"])
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
                if h > 0 or read_zero_hundred: res.append("lẻ")
                res.append(words[u])
        return " ".join(res)

    chunks, temp = [], n
    while temp > 0:
        chunks.append(temp % 1000)
        temp //= 1000
        
    res = []
    for i, chunk in enumerate(chunks):
        if chunk > 0:
            read_zero = (i < len(chunks) - 1 and sum(chunks[i+1:]) > 0 and chunk < 100)
            chunk_str = read_3_digits(chunk, read_zero_hundred=read_zero)
            unit = units[i] if i < len(units) else ""
            res.insert(0, (chunk_str + " " + unit).strip())
            
    final_str = " ".join(res).strip().replace("  ", " ")
    return final_str.capitalize() + " đồng"

def parse_invoice_dates(date_series):
    parsed_dates = pd.to_datetime(date_series, format='%d/%m/%Y', errors='coerce')
    return parsed_dates.fillna(pd.to_datetime(date_series, dayfirst=True, errors='coerce'))

def get_file_month_sort_key(df, date_column, sample_size=10):
    if not date_column or date_column not in df.columns: return (9999, 12)
    sampled_dates = df[date_column].dropna().head(sample_size)
    parsed_dates = parse_invoice_dates(sampled_dates).dropna()
    if parsed_dates.empty: return (9999, 12)
    min_date = parsed_dates.min()
    return (int(min_date.year), int(min_date.month))

def normalize_text(value):
    text = "" if pd.isna(value) else str(value)
    text = re.sub(r'\s+', ' ', text).strip()
    text = unicodedata.normalize('NFKD', text)
    return ''.join(char for char in text if not unicodedata.combining(char)).lower()

def reset_filter_state():
    st.session_state["filter_selected_kh"] = "Tất cả|||tat ca"
    st.session_state["filter_selected_hd"] = "Tất cả"
    st.session_state["filter_selected_months"] = []

def reset_month_filter():
    st.session_state["filter_selected_months"] = []

def clear_uploaded_filter_data():
    reset_filter_state()
    st.session_state["filter_uploader_nonce"] = st.session_state.get("filter_uploader_nonce", 0) + 1

def clear_uploaded_merge_data():
    st.session_state["merge_uploader_nonce"] = st.session_state.get("merge_uploader_nonce", 0) + 1
    for key in ["merge_cache_fingerprint", "merge_cached_df", "merge_cached_errors", "merge_cached_date_col"]:
        st.session_state.pop(key, None)

def get_uploaded_files_fingerprint(uploaded_files):
    return tuple((file.name, file.size) for file in uploaded_files)

def sanitize_filename(filename):
    filename = re.sub(r'[<>:"/\\|?*]+', '_', str(filename).strip())
    return re.sub(r'\s+', '_', filename) or "du_lieu"

def ensure_xlsx_extension(filename):
    filename = sanitize_filename(filename)
    return filename if filename.lower().endswith('.xlsx') else f"{filename}.xlsx"

def standardize_header_name(value, fallback_prefix="Cột"):
    text = re.sub(r'\s+', ' ', str(value).strip())
    if not text: text = fallback_prefix
    return ' '.join(word[:1].upper() + word[1:].lower() for word in text.split())

def make_unique_columns(columns):
    seen, unique_columns = {}, []
    for index, col in enumerate(columns, start=1):
        base_name = standardize_header_name(col, fallback_prefix=f"Cột {index}")
        count = seen.get(base_name, 0)
        unique_name = base_name if count == 0 else f"{base_name}_{count + 1}"
        seen[base_name] = count + 1
        unique_columns.append(unique_name)
    return unique_columns

def find_header_row(preview_df, keywords=None):
    if preview_df.empty: return 0
    if keywords:
        preview_text = preview_df.fillna('').astype(str)
        for idx, row in preview_text.iterrows():
            row_values = row.str.lower().str.strip().tolist()
            if any(any(kw in cell for kw in keywords) for cell in row_values): return idx
    return int(preview_df.notna().sum(axis=1).idxmax())

def safe_read_excel(file, header=0, nrows=None):
    file.seek(0)
    try: return pd.read_excel(file, header=header, nrows=nrows, engine='calamine')
    except Exception:
        file.seek(0)
        return pd.read_excel(file, header=header, nrows=nrows)

def collapse_duplicate_columns(df):
    collapsed_data = {}
    for col in pd.unique(df.columns):
        same_name_columns = df.loc[:, df.columns == col]
        collapsed_data[col] = same_name_columns.iloc[:, 0] if same_name_columns.shape[1] == 1 else same_name_columns.bfill(axis=1).iloc[:, 0]
    return pd.DataFrame(collapsed_data)

def find_best_matching_header(column_name, base_headers, threshold=0.85):
    normalized_column = normalize_text(column_name)
    best_match, best_score = None, 0
    for base_header in base_headers:
        score = SequenceMatcher(None, normalized_column, normalize_text(base_header)).ratio()
        if score > best_score:
            best_score, best_match = score, base_header
    return best_match if best_score >= threshold else None

def find_invoice_date_column(columns):
    keywords = ['ngay hoa don', 'ngay hoa don gtgt']
    best_match, best_score = None, 0
    for column in columns:
        normalized_column = normalize_text(column)
        if 'ngay' in normalized_column and 'hoa' in normalized_column and 'don' in normalized_column: return column
        for keyword in keywords:
            score = SequenceMatcher(None, normalized_column, keyword).ratio()
            if score > best_score: best_score, best_match = score, column
    return best_match if best_score >= 0.6 else None

def read_merge_source_file(file):
    preview_df = safe_read_excel(file, header=None, nrows=30)
    df = safe_read_excel(file, header=find_header_row(preview_df))
    df = df.dropna(axis=0, how='all')
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', case=False, regex=True)]
    df.columns = make_unique_columns(df.columns)
    return collapse_duplicate_columns(df).reset_index(drop=True)

def align_dataframe_columns(df, base_headers, threshold=0.85):
    rename_map = {col: find_best_matching_header(col, base_headers, threshold) or col for col in df.columns}
    df = collapse_duplicate_columns(df.rename(columns=rename_map))
    for column in df.columns:
        if column not in base_headers: base_headers.append(column)
    return df.reindex(columns=base_headers, fill_value=pd.NA), base_headers

def merge_sm2057_files(uploaded_files, progress_bar=None, status_placeholder=None):
    merged_frames, base_headers, errors = [], [], []

    for index, file in enumerate(uploaded_files, start=1):
        if progress_bar: progress_bar.progress(index / len(uploaded_files), text=f"Đang xử lý {index}/{len(uploaded_files)} file")
        if status_placeholder: status_placeholder.info(f"Đang đọc: {file.name}")

        try:
            current_df = read_merge_source_file(file)
            current_df[COL.NGUON_FILE] = file.name
            date_column = find_invoice_date_column(current_df.columns)
            month_sort_key = get_file_month_sort_key(current_df, date_column)

            if not base_headers:
                base_headers = list(current_df.columns)
            else:
                current_df, base_headers = align_dataframe_columns(current_df, base_headers)

            merged_frames.append((month_sort_key, index, current_df.reindex(columns=base_headers, fill_value=pd.NA)))
        except Exception as exc:
            errors.append(f"{file.name}: {exc}")

    if not merged_frames: return None, errors, None

    merged_frames.sort(key=lambda item: (item[0][0], item[0][1], item[1]))
    merged_df = pd.concat([item[2] for item in merged_frames], ignore_index=True).dropna(axis=0, how='all').reset_index(drop=True)
    return merged_df, errors, find_invoice_date_column(merged_df.columns)

def build_agg_rules(columns, group_col):
    agg_rules = {}
    for col in columns:
        if col in GLOBAL_CALC_COLS: agg_rules[col] = 'sum'
        elif col not in (group_col, COL.KH_DISPLAY, COL.SORT_DATE): agg_rules[col] = 'first'
    return agg_rules

def build_bang_08a(source_df):
    if COL.SO_CT not in source_df.columns or COL.NGAY_HD not in source_df.columns:
        return pd.DataFrame(columns=[COL.SO_CT, COL.NGAY_HD, COL.NOI_DUNG, COL.DVT, COL.SO_LUONG, COL.DON_GIA, COL.THANH_TIEN])

    amount_col = next((col for col in [COL.THANH_TIEN_CO_VAT, COL.THANH_TIEN_CHUA_VAT] if col in source_df.columns), None)
    
    grouped = source_df.groupby(COL.SO_CT, as_index=False, sort=False).agg({
        COL.NGAY_HD: 'first',
        COL.SORT_DATE: 'first' if COL.SORT_DATE in source_df.columns else (lambda x: pd.NaT)
    })

    if amount_col:
        amount_df = source_df.groupby(COL.SO_CT, as_index=False, sort=False)[amount_col].sum()
        grouped = grouped.merge(amount_df, on=COL.SO_CT, how='left')
        grouped[COL.THANH_TIEN] = grouped[amount_col]
        grouped = grouped.drop(columns=[amount_col])
    else:
        grouped[COL.THANH_TIEN] = 0

    grouped[COL.NOI_DUNG] = "Hoá đơn số " + grouped[COL.SO_CT].astype(str) + " ngày " + grouped[COL.NGAY_HD].astype(str)
    grouped[COL.DVT] = ""
    grouped[COL.SO_LUONG] = ""
    grouped[COL.DON_GIA] = ""

    return grouped[[COL.SO_CT, COL.NGAY_HD, COL.NOI_DUNG, COL.DVT, COL.SO_LUONG, COL.DON_GIA, COL.THANH_TIEN]]

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
    dfs = []
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                file.seek(0)
                header_idx = find_header_row(pd.read_csv(file, header=None, nrows=30), keywords=["mã kh hóa đơn", "mã kh hoá đơn"])
                file.seek(0)
                df = pd.read_csv(file, header=header_idx)
            elif file.name.endswith(('.xls', '.xlsx')):
                header_idx = find_header_row(safe_read_excel(file, header=None, nrows=30), keywords=["mã kh hóa đơn", "mã kh hoá đơn"])
                df = safe_read_excel(file, header=header_idx)
            else: continue
                
            df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
            dfs.append((get_file_month_sort_key(df, COL.NGAY_HD if COL.NGAY_HD in df.columns else None), len(dfs), df))
        except Exception as e:
            st.error(f"Lỗi khi đọc file {file.name}: {e}")
            
    if not dfs: return None

    dfs.sort(key=lambda item: (item[0][0], item[0][1], item[1]))
    master_df = pd.concat([item[2] for item in dfs], ignore_index=True)
    
    # Chuẩn hóa tên cột để xử lý lỗi gõ dấu (Hóa vs Hoá)
    master_df.columns = master_df.columns.astype(str).str.strip().str.replace('oá', 'óa', regex=False).str.replace('Oá', 'Óa', regex=False)
    
    if COL.NGAY_HD in master_df.columns:
        master_df[COL.SORT_DATE] = parse_invoice_dates(master_df[COL.NGAY_HD])
    
    if COL.MA_KH in master_df.columns and COL.TEN_KH in master_df.columns:
        master_df[COL.KH_CODE] = master_df[COL.MA_KH].astype(str).str.strip()
        master_df[COL.KH_SEARCH] = master_df[COL.TEN_KH].astype(str)
        master_df[COL.KH_DISPLAY] = master_df[COL.MA_KH].astype(str) + " - " + master_df[COL.TEN_KH].astype(str)
    else:
        master_df[COL.KH_CODE] = master_df[COL.KH_SEARCH] = "Không xác định"
        master_df[COL.KH_DISPLAY] = "Không xác định - Không xác định"
        
    master_df[COL.SO_HD] = master_df[COL.SO_HD].fillna("Không xác định") if COL.SO_HD in master_df.columns else "Không xác định"
        
    if COL.NHA_SX in master_df.columns:
        master_df[COL.NHA_SX] = master_df[COL.NHA_SX].fillna('').astype(str).apply(lambda x: x.split('_', 1)[-1].strip() if '_' in x else x)
        
    for col in GLOBAL_CALC_COLS:
        if col in master_df.columns:
            if master_df[col].dtype == 'object':
                master_df[col] = master_df[col].astype(str).str.replace(r'\.', '', regex=True).str.replace(',', '.', regex=False)
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0)
            
    return master_df


# ==========================================
# 3. CÁC HÀM RENDER GIAO DIỆN (UI COMPONENTS)
# ==========================================
def render_result_table(final_df, selected_columns, tab_key, show_total_text=True, default_file_name=None):
    if not selected_columns:
        st.warning("Vui lòng chọn ít nhất 1 cột để hiển thị kết quả.")
        return

    available_columns = [col for col in selected_columns if col in final_df.columns and col != COL.SORT_DATE]
    display_df = final_df.loc[:, available_columns].copy()
    
    if COL.STT in display_df.columns: display_df = display_df.drop(columns=[COL.STT])
    display_df.insert(0, COL.STT, range(1, len(display_df) + 1))
    
    total_display_value, total_display_label = 0, None
    if COL.THANH_TIEN_CO_VAT in display_df.columns:
        total_display_value = display_df[COL.THANH_TIEN_CO_VAT].sum()
        total_display_label = f'Tổng {COL.THANH_TIEN_CO_VAT}'
    elif COL.THANH_TIEN in display_df.columns:
        total_display_value = display_df[COL.THANH_TIEN].sum()
        total_display_label = f'Tổng {COL.THANH_TIEN}'
    
    export_df = display_df.fillna("").copy()
    
    def format_vn_number(x):
        try: return f"{float(x):,.0f}".replace(",", ".") if pd.notna(x) and x != "" else ""
        except: return x
            
    visual_df = export_df.copy()
    format_cols = GLOBAL_CALC_COLS + EXTRA_FORMAT_COLS
    for col in format_cols:
        if col in visual_df.columns: visual_df[col] = visual_df[col].apply(format_vn_number)
    
    num_cols = [c for c in format_cols if c in visual_df.columns]
    st.dataframe(
        visual_df.style.set_properties(subset=num_cols, **{'text-align': 'right'}),
        use_container_width=True, hide_index=True,
        column_config={COL.STT: st.column_config.Column(width="small")}
    )
    
    st.write("") # Tạo khoảng trống
    
    # CHIA CỘT CHÍNH [8, 2]
    col_main, col_copy = st.columns([8, 2])
    
    # Tính trước format tiền để dùng cho cả 2 cột
    fmt_val = ""
    if show_total_text and total_display_label:
        fmt_val = f"{float(total_display_value):,.0f}".replace(",", ".")
        
    with col_main:
        # --- 1. KHU VỰC HIỂN THỊ TỔNG TIỀN (Hiển thị nổi bật dạng Card) ---
        if show_total_text and total_display_label:
            st.markdown(f"""
            <div style="background-color: #f8fafc; padding: 12px 20px; border-radius: 8px; border-left: 5px solid #0ea5e9; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <span style="font-size: 1.1em;"><strong>💰 {total_display_label}:</strong></span>
                <span style="color:#dc2626; font-size: 1.3em; font-weight: bold; margin-left: 8px;">{fmt_val}</span> VNĐ <br>
                <span style="font-size: 1em; color: #475569; margin-top: 5px; display: inline-block;"><strong>✍️ Bằng chữ:</strong> <i>{doc_so_thanh_chu(total_display_value)}</i></span>
            </div>
            """, unsafe_allow_html=True)

        # --- 2. KHU VỰC TẢI VỀ ---
        dl_col1, dl_col2 = st.columns([4, 2])
        with dl_col1:
            file_name_value = st.text_input(
                "Tên file tải về:",
                value=ensure_xlsx_extension(default_file_name or f'ket_qua_{tab_key}'),
                key=f"download_name_{tab_key}"
            )
            
        with dl_col2:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Tải Excel (.xlsx)",
                data=build_export_excel(export_df),
                file_name=ensure_xlsx_extension(file_name_value),
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f"dl_btn_{tab_key}",
                use_container_width=True
            )
            
    with col_copy:
        # --- 3. KHU VỰC COPY ---
        st.code(export_df.to_csv(sep='\t', index=False), language='plaintext')
        
        if show_total_text and total_display_label:
            st.code(fmt_val, language='plaintext')
            st.code(doc_so_thanh_chu(total_display_value), language='plaintext')
            
        st.markdown("<div style='margin-top: 5px; font-size: 0.9em; color: #64748b; text-align: right;'><b>📋 Copy nhanh Bảng, Số tiền, Bằng chữ</b></div>", unsafe_allow_html=True)

def get_available_columns_for_tabs(df):
    return [col for col in df.columns if col not in (COL.KH_DISPLAY, COL.KH_SEARCH, COL.KH_CODE, COL.SORT_DATE) and str(col).upper() != COL.STT]

def render_tab_chi_tiet(filtered_df, available_columns, selected_kh, selected_hd):
    desired_defaults = [COL.NGAY_HD, COL.SO_CT, COL.TEN_HH, COL.DVT, COL.SO_LUONG, COL.DON_GIA_CO_VAT, COL.THANH_TIEN_CO_VAT, COL.NHA_SX]
    default_cols = [col for col in desired_defaults if col in available_columns] or available_columns[:8]

    with st.expander("Cột hiển thị", expanded=False):
        selected_columns = st.multiselect("Cột hiển thị (có thể chuyển vị trí cột)", options=available_columns, default=default_cols, key="multi_t1")

    st.write("")
    render_result_table(filtered_df, selected_columns, "chi_tiet_hoa_don_hang", default_file_name=f"chi_tiet_hoa_don_hang_{selected_kh}_{selected_hd or 'tat_ca'}")

def render_tab_hang_hoa(filtered_df, available_columns, selected_kh, selected_hd):
    desired_defaults = [COL.TEN_HH, COL.DVT, COL.SO_LUONG, COL.DON_GIA_CO_VAT, COL.THANH_TIEN_CO_VAT, COL.NHA_SX]
    default_cols = [col for col in desired_defaults if col in available_columns] or available_columns[:6]

    with st.expander("Cột hiển thị", expanded=False):
        selected_columns = st.multiselect("Cột hiển thị (có thể chuyển vị trí cột)", options=available_columns, default=default_cols, key="multi_t2")

    final_df = aggregate_dataframe(filtered_df, COL.MA_HH) if COL.MA_HH in filtered_df.columns else filtered_df
    st.write("")
    render_result_table(final_df, selected_columns, "bang_hang_hoa", default_file_name=f"bang_hang_hoa_{selected_kh}_{selected_hd or 'tat_ca'}")

def render_tab_hoa_don(filtered_df, available_columns, selected_kh, selected_hd):
    desired_defaults = [COL.NGAY_HD, COL.SO_CT, COL.THANH_TIEN_CO_VAT]
    default_cols = [col for col in desired_defaults if col in available_columns] or available_columns[:3]

    with st.expander("Cột hiển thị", expanded=False):
        selected_columns = st.multiselect("Cột hiển thị (có thể chuyển vị trí cột)", options=available_columns, default=default_cols, key="multi_t3")

    final_df = aggregate_dataframe(filtered_df, COL.SO_CT, keep_sort_order=True) if COL.SO_CT in filtered_df.columns else filtered_df
    st.write("")
    render_result_table(final_df, selected_columns, "danh_sach_hoa_don", default_file_name=f"danh_sach_hoa_don_{selected_kh}_{selected_hd or 'tat_ca'}")

def render_tab_bang_08a(filtered_df, selected_kh, selected_hd):
    final_df = build_bang_08a(filtered_df)
    available_columns = [COL.SO_CT, COL.NGAY_HD, COL.NOI_DUNG, COL.DVT, COL.SO_LUONG, COL.DON_GIA, COL.THANH_TIEN]

    with st.expander("Cột hiển thị", expanded=False):
        selected_columns = st.multiselect("Cột hiển thị (có thể chuyển vị trí cột)", options=available_columns, default=available_columns, key="multi_t4")
        
    st.write("")
    render_result_table(final_df, selected_columns, "bang_08a", default_file_name=f"bang_08a_{selected_kh}_{selected_hd or 'tat_ca'}")

# ==========================================
# 4. GIAO DIỆN CHÍNH (MAIN BLOCKS)
# ==========================================
def render_filter_section():
    st.header("1. Tải dữ liệu lên")
    if "filter_uploader_nonce" not in st.session_state: st.session_state["filter_uploader_nonce"] = 0
    
    upload_col, clear_col = st.columns([1, 0.06])
    with upload_col:
        uploaded_files = st.file_uploader("Kéo thả, thêm nhiều file SM2057 (NT)", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True, key=f"filter_uploader_{st.session_state['filter_uploader_nonce']}")
    with clear_col:
        with st.container(key="clear_upload_btn_filter"):
            st.button("↻", key="clear_uploaded_files_btn", use_container_width=True, help="Xóa dữ liệu đã tải", on_click=clear_uploaded_filter_data)

    if uploaded_files:
        with st.spinner('Đang đọc và gộp dữ liệu...'):
            df = process_uploaded_files(uploaded_files)
            
        if df is not None:
            st.success(f"✅ Đã gộp thành công {len(uploaded_files)} file. Tổng số dòng dữ liệu: {len(df)}")
            st.header("2. Lọc khách hàng & Hợp đồng")

            with st.container(key="filter_bar"):
                col1, col2, col3, col4 = st.columns([1.2, 1.2, 1, 0.22])
                with col1:
                    kh_lookup = df[[COL.KH_CODE, COL.KH_DISPLAY, COL.KH_SEARCH]].dropna(subset=[COL.KH_CODE]).groupby(COL.KH_CODE, as_index=False).agg({COL.KH_DISPLAY: 'first', COL.KH_SEARCH: 'first'}).sort_values(by=COL.KH_DISPLAY)
                    kh_options = ["Tất cả|||tat ca"] + [f"{row.KH_Code}|||{row.KH_Display}|||{normalize_text(row.KH_Search)}" for row in kh_lookup.itertuples(index=False)]
                    selected_kh_option = st.selectbox("1️⃣ Chọn khách hàng:", kh_options, format_func=lambda v: "Tất cả" if v == "Tất cả|||tat ca" else v.split("|||", 2)[1], key="filter_selected_kh", on_change=reset_month_filter)
                    selected_kh = "Tất cả" if selected_kh_option == "Tất cả|||tat ca" else selected_kh_option.split("|||", 2)[0]

                filtered_df = df[df[COL.KH_CODE] == selected_kh] if selected_kh != "Tất cả" else df

                with col2:
                    hd_list = ["Tất cả"] + list(filtered_df[COL.SO_HD].dropna().unique())
                    selected_hd = st.selectbox("2️⃣ Chọn Số Hợp Đồng:", hd_list, key="filter_selected_hd")

                if selected_hd != "Tất cả": filtered_df = filtered_df[filtered_df[COL.SO_HD] == selected_hd]

                month_options = []
                if COL.SORT_DATE in filtered_df.columns:
                    month_options = sorted(filtered_df[COL.SORT_DATE].dropna().dt.strftime('%m/%Y').unique().tolist(), key=lambda x: (int(x[3:]), int(x[:2])))

                with col3:
                    selected_months = st.multiselect("3️⃣ Chọn tháng: (Mặc định tất cả)", month_options, placeholder="Chọn Tháng nếu muốn", key="filter_selected_months")

                with col4:
                    st.write("")
                    st.button("↻", key="filter_reset_btn", use_container_width=True, help="Reset bộ lọc", on_click=reset_filter_state)

                if selected_months and COL.SORT_DATE in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df[COL.SORT_DATE].dt.strftime('%m/%Y').isin(selected_months)]

            st.caption(f"🔍 Dữ liệu hiện tại: **{len(filtered_df)}** dòng khớp điều kiện.")
            
            tab1, tab2, tab3, tab4 = st.tabs(["📦 Bảng Chi tiết Hoá đơn - Hàng", "🛒 Bảng Hàng hoá", "🧾 Bảng Hóa đơn", "📄 Bảng 08a"])
            available_cols = get_available_columns_for_tabs(df)
            
            with tab1: render_tab_chi_tiet(filtered_df, available_cols, selected_kh, selected_hd)
            with tab2: render_tab_hang_hoa(filtered_df, available_cols, selected_kh, selected_hd)
            with tab3: render_tab_hoa_don(filtered_df, available_cols, selected_kh, selected_hd)
            with tab4: render_tab_bang_08a(filtered_df, selected_kh, selected_hd)


def render_merge_section():
    st.header("Làm sạch & Gộp file Excel giống nhau")
    if "merge_uploader_nonce" not in st.session_state: st.session_state["merge_uploader_nonce"] = 0
    
    merge_upload_col, merge_clear_col = st.columns([1, 0.06])
    with merge_upload_col:
        merge_files = st.file_uploader("Kéo thả nhiều file Excel vào đây (.xlsx, .xls)", type=['xlsx', 'xls'], accept_multiple_files=True, key=f"merge_uploader_{st.session_state['merge_uploader_nonce']}")
    with merge_clear_col:
        with st.container(key="clear_upload_btn_merge"):
            st.button("↻", key="clear_uploaded_merge_files_btn", use_container_width=True, help="Xóa dữ liệu đã tải", on_click=clear_uploaded_merge_data)

    if merge_files:
        merge_fingerprint = get_uploaded_files_fingerprint(merge_files)
        cache_key = "merge_cache_fingerprint"
        
        if st.session_state.get(cache_key) != merge_fingerprint:
            progress_bar = st.progress(0, text="Sẵn sàng xử lý")
            status_placeholder = st.empty()
            merged_df, merge_errors, detected_date_col = merge_sm2057_files(merge_files, progress_bar, status_placeholder)

            progress_bar.progress(1.0, text="Đã hoàn tất xử lý")
            status_placeholder.empty()
            st.session_state.update({cache_key: merge_fingerprint, "merge_cached_df": merged_df, "merge_cached_errors": merge_errors, "merge_cached_date_col": detected_date_col})
        else:
            merged_df, merge_errors, detected_date_col = st.session_state.get("merge_cached_df"), st.session_state.get("merge_cached_errors", []), st.session_state.get("merge_cached_date_col")

        if merged_df is not None:
            st.success(f"✅ Đã gộp thành công {len(merge_files)} file.")
            st.caption(f"📌 Tổng số dòng dữ liệu thu thập được: **{len(merged_df)}**")
            if detected_date_col: st.caption(f"🗓️ Đã phát hiện cột ngày hóa đơn để sắp xếp: **{detected_date_col}**")
            if merge_errors:
                st.warning("Một số file không đọc được:")
                for error in merge_errors: st.write(f"- {error}")

            merge_download_name = st.text_input("Tên file tải về", value=ensure_xlsx_extension(f"DuLieu_DaGop_{datetime.now().strftime('%Y-%m-%d')}"), key="download_name_merged_sm2057")
            st.download_button(label="📥 Tải về DuLieu_DaGop.xlsx", data=build_export_excel(merged_df), file_name=ensure_xlsx_extension(merge_download_name), mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key="download_merged_sm2057")
            st.subheader("Xem trước dữ liệu")
            st.dataframe(merged_df.head(50), width=1500, hide_index=True)
        else:
            st.error("Không thể gộp dữ liệu từ các file đã tải lên.")
            if merge_errors:
                for error in merge_errors: st.write(f"- {error}")
    else:
        clear_uploaded_merge_data() # Reset nếu xóa file

# ==========================================
# 5. ĐIỂM BẮT ĐẦU (ENTRY POINT)
# ==========================================
def main():
    setup_page_config()
    
    main_tab_filter, main_tab_merge = st.tabs(["📊 Lọc Super SM2057 (NT)", "🧩 Gộp file SM2057"])
    
    with main_tab_filter:
        render_filter_section()
        
    with main_tab_merge:
        render_merge_section()

    st.markdown(
        """
        <div style="text-align: center; margin-top: 60px; padding-bottom: 20px; color: #888888; font-size: 14px;">
            Thiết kế: Nguyễn Văn Dũng ❤️ 0978.777.191
        </div>
        """, 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()