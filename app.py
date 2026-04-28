import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import unicodedata
import re
import io
import time
from datetime import date

# ==========================================
# CẤU HÌNH GIAO DIỆN & SESSION STATE
# ==========================================
st.set_page_config(page_title="Hệ thống Đối chiếu CPC1", page_icon="🏥", layout="wide")

# Khởi tạo bộ nhớ tạm (Session State)
if 'alias_hc' not in st.session_state: st.session_state.alias_hc = {}
if 'alias_hl' not in st.session_state: st.session_state.alias_hl = {}
if 'list_aliases' not in st.session_state: st.session_state.list_aliases = [] 
if 'new_aliases' not in st.session_state: st.session_state.new_aliases = [] 
if 'master_file_name' not in st.session_state: st.session_state.master_file_name = ""
if 'target_file_name' not in st.session_state: st.session_state.target_file_name = "" # BIẾN MỚI CHO TARGET
if 'processed' not in st.session_state: st.session_state.processed = False
if 'uploader_key' not in st.session_state: st.session_state.uploader_key = 0

# ==========================================
# CONSTANTS & DICTIONARIES
# ==========================================
REQUIRED_COLUMNS = ["Nhóm", "Hoạt_chất", "Hàm_lượng"]

# Gom nhóm từ khóa theo từng cột chuẩn để dễ nhìn và dễ thêm mới
_HEADER_CONFIG = {
    "Mã_phần": [
        "ma phan", "so phan", "ma phan thau", "ma phan lo", "phan lo", "so phan lo", "ma lo", "phan thau"
    ],
    "Nhóm": [
        "nhom", "nhom thau", "nhom thuoc", "phan nhom", "group", "nhom hang", "nhom sp"
    ],
    "Hoạt_chất": [
        "hoat chat", "ten hoat chat", "thanh phan", "thanh phan chinh", "hoatchat", "ten hc"
    ],
    "Tên_thuốc": [
        "ten thuoc", "ten mat hang", "ten thuong mai", "mat hang", 
        "ten sp", "ten san pham", "biet duoc", "ten biet duoc", "san pham", "hang hoa", "ten hang", "ten day du"
    ],
    "Hàm_lượng": [
        "ham luong", "nong do", "ham luong nong do", "nong do ham luong", "nong doham luong", "hl nd", "hamluong", "nongdo"
    ],
    "Hãng": [
        "hang", "brand", "thuong hieu"
    ],
    "Mã_sản_phẩm": [
        "ma sp", "ma san pham", "ma thuoc", "ma hang", "ma cpc1", "ma vtyt", "ma bfo"
    ]
}

# Tự động bung danh sách trên thành Dictionary phẳng cho hệ thống tra cứu (variant: standard)
HEADER_ALIASES = {variant: std_col for std_col, variants in _HEADER_CONFIG.items() for variant in variants}

# Đơn vị hàm lượng → quy đổi về mg (mass) hoặc ml (volume)
_UNIT_FACTOR = {
    "mcg": 0.001, "ug": 0.001,
    "mg":  1.0,
    "g":   1000.0,
    "kg":  1000000.0,
    "ml":  1.0,
    "l":   1000.0,
    "ui":  1.0, "iu": 1.0, "don vi": 1.0, "donvi": 1.0,
    "%":   1.0,
}

_UNIT_GROUP = {
    "mcg": "mass", "ug": "mass", "mg": "mass", "g": "mass", "kg": "mass",
    "ml": "volume", "l": "volume",
    "ui": "other", "iu": "other", "don vi": "other", "donvi": "other",
    "%": "pct",
}

# ==========================================
# LÕI XỬ LÝ DỮ LIỆU (DATA ENGINE)
# ==========================================
def remove_accents(input_str):
    if pd.isna(input_str): return ""
    s = str(input_str).strip().lower()
    s = s.replace('/', ' ').replace('\n', ' ').replace('-', ' ').replace('_', ' ').replace('(', ' ').replace(')', ' ')
    s = ' '.join(s.split()) 
    return unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')

def read_excel_auto_header(excel_file_bytes, sheet_name=0):
    file_stream = io.BytesIO(excel_file_bytes)
    df_temp = pd.read_excel(file_stream, sheet_name=sheet_name, nrows=20, header=None)
    header_idx = 0
    max_matches = 0
    
    for idx, row in df_temp.iterrows():
        row_strs = [remove_accents(str(x)) for x in row.values]
        matches = sum(1 for val in row_strs if val in HEADER_ALIASES)
        if matches > max_matches:
            max_matches = matches
            header_idx = idx
            
    file_stream.seek(0)
    return pd.read_excel(file_stream, sheet_name=sheet_name, header=header_idx)

def auto_map_headers(df):
    mapped_columns = {}
    for col in df.columns:
        normalized_col = remove_accents(col)
        if normalized_col in HEADER_ALIASES:
            mapped_columns[col] = HEADER_ALIASES[normalized_col]
            continue
            
        best_score = 0
        best_target = None
        for alias_key, target_col in HEADER_ALIASES.items():
            score = fuzz.token_sort_ratio(normalized_col, alias_key)
            if score > best_score:
                best_score = score
                best_target = target_col
                
        if best_score >= 85:
            mapped_columns[col] = best_target
        else:
            mapped_columns[col] = col

    df = df.rename(columns=mapped_columns)
    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    return df, missing_cols

def load_alias_from_excel(excel_file_bytes):
    """Tìm sheet chứa chữ 'alias' trong file gốc để nạp từ điển"""
    try:
        xls = pd.ExcelFile(io.BytesIO(excel_file_bytes))
        # Tìm sheet có tên chứa 'alias'
        alias_sheet = next((s for s in xls.sheet_names if 'alias' in s.lower()), None)
        
        if alias_sheet:
            df_alias = pd.read_excel(xls, sheet_name=alias_sheet)
            
            # CHUẨN HÓA TÊN CỘT: Xóa khoảng trắng và đưa về chữ thường
            df_alias.columns = [str(c).strip().lower() for c in df_alias.columns]
            
            # Đổi tên cột tự động để hỗ trợ cả định dạng cũ (field) và mới (Trường)
            col_mapping = {
                'field': 'Trường', 'trường': 'Trường',
                'variant': 'Biến thể (đầu vào)', 'biến thể (đầu vào)': 'Biến thể (đầu vào)',
                'standard': 'Chuẩn (CPC1)', 'chuẩn (cpc1)': 'Chuẩn (CPC1)'
            }
            df_alias.rename(columns=col_mapping, inplace=True)
            
            req_cols = ['Trường', 'Biến thể (đầu vào)', 'Chuẩn (CPC1)']
            if all(col in df_alias.columns for col in req_cols):
                count_hc, count_hl = 0, 0
                for _, row in df_alias.iterrows():
                    field_raw = str(row['Trường']).strip().lower()
                    raw_variant = str(row['Biến thể (đầu vào)'])
                    variant = remove_accents(raw_variant)
                    standard = str(row['Chuẩn (CPC1)']).strip()
                    
                    if str(raw_variant).lower() == 'nan' or not raw_variant.strip():
                        continue
                        
                    # Phân loại Alias và nạp vào Session State
                    if field_raw in ['hoạt chất', 'hoat_chat', 'hoat chat']:
                        st.session_state.alias_hc[variant] = standard
                        count_hc += 1
                        display_field = "Hoạt chất"
                    elif field_raw in ['hàm lượng', 'ham_luong', 'ham luong']:
                        st.session_state.alias_hl[variant] = standard
                        count_hl += 1
                        display_field = "Hàm lượng"
                    else:
                        continue # Bỏ qua nếu field không hợp lệ
                        
                    def get_val(col_name):
                        val = str(row.get(col_name, ''))
                        return "" if val.lower() == 'nan' else val.strip()

                    # Lưu vào danh sách lịch sử để xuất ra file sau này
                    st.session_state.list_aliases.append({
                        "Trường": display_field, 
                        "Biến thể (đầu vào)": raw_variant, 
                        "Chuẩn (CPC1)": standard,
                        "Mã sản phẩm": get_val("mã sản phẩm"), 
                        "Tên sản phẩm": get_val("tên sản phẩm"), 
                        "Ghi chú": get_val("ghi chú"), 
                        "Ngày thêm": get_val("ngày thêm"), 
                        "Người thêm": get_val("người thêm")
                    })
                return True, count_hc, count_hl
        return False, 0, 0
    except Exception as e:
        return False, 0, 0

def parse_groups(val):
    if pd.isna(val): return []
    if isinstance(val, float) and val.is_integer(): val = int(val)
    s = str(val).strip().upper()
    if s.endswith('.0') and s.replace('.0', '').isdigit(): s = s[:-2]
    s = remove_accents(s).upper()
    tokens = re.findall(r'[A-Z0-9]+', s)
    ignore_words = {"NHOM", "GROUP", "THAU", "THUOC"}
    return [t for t in tokens if t not in ignore_words]

def clean_dosage_string(val):
    if pd.isna(val): return ""
    s = str(val).lower().replace(',', '.')
    s = re.sub(r'[^a-z0-9\.]', ' ', s)
    s = re.sub(r'([0-9])([a-z])', r'\1 \2', s)
    s = re.sub(r'([a-z])([0-9])', r'\1 \2', s)
    return ' '.join(s.split())

def extract_all_dosages(dosage_str):
    if pd.isna(dosage_str): return []
    cleaned_str = str(dosage_str).lower().replace(',', '.')
    
    # Mẹo nhỏ: Xử lý trước chữ "don vi" vì nó có dấu cách, regex bình thường hay bị bắt hụt
    cleaned_str = cleaned_str.replace('don vi', 'iu')
    
    matches = re.findall(r"([\d\.]+)\s*([a-z%]+)", cleaned_str)
    
    results = []
    for val_str, unit in matches:
        try: 
            val = float(val_str)
        except ValueError: 
            continue
            
        # Sử dụng Dictionary quy đổi hệ số và nhóm
        factor = _UNIT_FACTOR.get(unit, 1.0)
        group = _UNIT_GROUP.get(unit, unit)
        
        std_val = val * factor
        results.append((std_val, group))
        
    return results

def check_dosage_match(input_raw, cpc1_raw):
    if pd.isna(input_raw) or pd.isna(cpc1_raw): return False
    
    cpc1_str = str(cpc1_raw).lower()
    cpc1_options = re.split(r'\s+hoặc\s+|\s+or\s+', cpc1_str)
    
    s1 = clean_dosage_string(input_raw)
    dosages1 = extract_all_dosages(input_raw)
    
    for option in cpc1_options:
        if not option.strip(): continue
        s2 = clean_dosage_string(option)
        if fuzz.token_sort_ratio(s1, s2) >= 90: return True
            
        dosages2 = extract_all_dosages(option)
        if not dosages1 or not dosages2: continue
            
        if dosages1[0] == dosages2[0]:
            if len(dosages1) == len(dosages2):
                if dosages1 == dosages2: return True
            else:
                return True
    return False

def standardize_combo(hc_raw, hl_raw):
    """Hàm trói cặp (Zipping) và sắp xếp A-Z cho thuốc phối hợp"""
    hc_str = str(hc_raw).lower().strip()
    hl_str = str(hl_raw).lower().strip()
    
    # Tách cụm theo dấu +, /, hoặc ;
    hc_parts = [p.strip() for p in re.split(r'\+|/|;', hc_str) if p.strip()]
    hl_parts = [p.strip() for p in re.split(r'\+|/|;', hl_str) if p.strip()]
    
    # Nếu là thuốc phối hợp và số lượng hoạt chất == số lượng hàm lượng
    if len(hc_parts) > 1 and len(hc_parts) == len(hl_parts):
        pairs = list(zip(hc_parts, hl_parts))
        # Sort A-Z theo tên hoạt chất (phần tử [0] của mỗi cặp)
        pairs.sort(key=lambda x: x[0])
        # Ghép lại thành chuỗi tiêu chuẩn
        sorted_hc = " + ".join([p[0] for p in pairs])
        sorted_hl = " + ".join([p[1] for p in pairs])
        return sorted_hc, sorted_hl
        
    return hc_str, hl_str

def process_single_row(input_row, cpc1_db, wratio_threshold=80):
    # Lấy thông tin Mã phần (nếu có), dọn dẹp đuôi .0 nếu Excel tự động ép kiểu số
    ma_phan = str(input_row.get("Mã_phần", "")).strip()
    if ma_phan.lower() == 'nan': ma_phan = ""
    # Nếu bị dính đuôi .0 (vd: 1.0, 2.0), gọt bỏ để trả về đúng số phần nguyên thủy
    if ma_phan.endswith('.0') and ma_phan.replace('.0', '').isdigit(): 
        ma_phan = ma_phan[:-2]

    for field in REQUIRED_COLUMNS:
        if pd.isna(input_row.get(field)) or str(input_row.get(field)).strip() == "":
            return {
                "Mã phần": ma_phan,
                "Dữ liệu Danh mục cần đối chiếu": "Lỗi Dữ Liệu",
                "Dữ liệu Danh mục gốc CPC1": f"Vui lòng bổ sung cột [{field}]",
                "Tên mặt hàng (CPC1)": "",
                "Hãng chào": "",
                "Tỉ lệ khớp": "N/A",
                "Ghi chú / Giải trình": "Thiếu dữ liệu đầu vào"
            }

    alias_hc = st.session_state.alias_hc
    alias_hl = st.session_state.alias_hl

    raw_hc = str(input_row.get("Hoạt_chất")).strip()
    norm_hc = remove_accents(raw_hc)
    input_hc = alias_hc.get(norm_hc, raw_hc) 
    
    raw_hl = str(input_row.get("Hàm_lượng"))
    norm_hl = remove_accents(raw_hl)
    input_dosage_raw = alias_hl.get(norm_hl, raw_hl)
    
    input_groups = parse_groups(input_row.get("Nhóm"))
    input_display = f"{raw_hc} - {raw_hl} - Nhóm {str(input_row.get('Nhóm')).strip()}"
    used_alias = (input_hc != raw_hc) or (input_dosage_raw != raw_hl)

    # Trói cặp (Zipping) dữ liệu đầu vào
    input_hc_std, input_hl_std = standardize_combo(input_hc, input_dosage_raw)

    best_match = None
    highest_score = 0
    highest_weight = -1 # <-- BIẾN MỚI: Trọng số ưu tiên (3: 100%, 2: 80-99%, 1: 50-79%)
    match_category = "Dưới 50%"
    note = "Không có sản phẩm nào thỏa mãn điều kiện."

    for _, cpc1_row in cpc1_db.iterrows():
        cpc1_groups = parse_groups(cpc1_row.get("Nhóm"))
        group_match = ("BD" in input_groups) or ("BD" in cpc1_groups) or bool(set(input_groups) & set(cpc1_groups))
        if not group_match: continue

        cpc1_hc_raw = str(cpc1_row.get("Hoạt_chất")).strip()
        cpc1_hl_raw = str(cpc1_row.get("Hàm_lượng")).strip()
        
        # Trói cặp (Zipping) dữ liệu CPC1
        cpc1_hc_std, cpc1_hl_std = standardize_combo(cpc1_hc_raw, cpc1_hl_raw)

        # --- BỘ LỌC 3 TẦNG CHO HOẠT CHẤT ---
        if input_hc_std == cpc1_hc_std:
            # Tầng 1: Khớp tuyệt đối (Exact Match)
            hc_score = 100
        else:
            # Tầng 2: Đặc trị đảo từ (token_sort_ratio)
            score_sort = fuzz.token_sort_ratio(input_hc_std, cpc1_hc_std)
            if score_sort == 100:
                hc_score = 100
            else:
                # Tầng 3: Cứu cánh lỗi chính tả (WRatio)
                score_wratio = fuzz.WRatio(input_hc_std, cpc1_hc_std)
                hc_score = max(score_sort, score_wratio)

        if hc_score < 50: continue

        # --- ĐỐI CHIẾU HÀM LƯỢNG TRÊN CHUỖI ĐÃ ZIPPING ---
        dosage_match = check_dosage_match(input_hl_std, cpc1_hl_std)

        # Đánh giá mức độ và gán trọng số
        if hc_score == 100 and dosage_match:
            current_category = "100%"
            current_note = "Khớp chính xác."
            weight = 3
        elif hc_score >= wratio_threshold and dosage_match:
            current_category = "80% - 99%"
            current_note = f"Sai khác chính tả. Chuỗi gốc: {cpc1_row['Hoạt_chất']}"
            weight = 2
        elif hc_score >= wratio_threshold and not dosage_match:
            current_category = "50% - 79%"
            current_note = f"Lệch hàm lượng. Gợi ý gần nhất: {cpc1_row['Hàm_lượng']} (Nhóm {cpc1_row['Nhóm']})"
            weight = 1
        else:
            continue

        # FIX BUG: Cập nhật nếu Trọng số TỐT HƠN, hoặc Trọng số bằng nhau nhưng Điểm Hoạt chất cao hơn
        if weight > highest_weight or (weight == highest_weight and hc_score > highest_score):
            highest_weight = weight
            highest_score = hc_score
            best_match = cpc1_row
            match_category = current_category
            note = current_note
            if highest_weight == 3: break # Tối ưu hiệu năng: Tìm thấy 100% thì dừng luôn

    if used_alias and best_match is not None:
        note += f" (Đã ánh xạ Alias: {input_hc} - {input_dosage_raw})"

    if best_match is not None:
        cpc1_display = f"{best_match.get('Hoạt_chất')} - {best_match.get('Hàm_lượng')} - Nhóm {best_match.get('Nhóm')}"
        
        # FIX LỖI "Không có thông tin": Lấy dữ liệu và kiểm tra xem có bị rỗng (NaN) không
        ten_mat_hang = str(best_match.get("Tên_thuốc", "")).strip()
        if ten_mat_hang.lower() == 'nan' or not ten_mat_hang:
            ten_mat_hang = "Không có thông tin"
            
        hang_sx = str(best_match.get("Hãng", "")).strip()
        if hang_sx.lower() == 'nan' or not hang_sx:
            hang_sx = "Không có thông tin"
            
    else:
        cpc1_display = "Không tìm thấy"
        ten_mat_hang = ""
        hang_sx = ""

    return {
        "Mã phần": ma_phan,
        "Dữ liệu Danh mục cần đối chiếu": input_display,
        "Dữ liệu Danh mục gốc CPC1": cpc1_display,
        "Tên mặt hàng (CPC1)": ten_mat_hang,
        "Hãng chào": hang_sx,
        "Tỉ lệ khớp": match_category,
        "Ghi chú / Giải trình": note
    }

def export_to_excel(main_df, main_sheet_name, list_aliases=None, alias_sheet_name="Alias"):
    """Hàm chung để xuất Excel kèm theo Sheet Alias"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        main_df.to_excel(writer, index=False, sheet_name=main_sheet_name)
        
        # LUÔN TẠO SHEET ALIAS: Dù list có rỗng thì cũng tạo bảng mẫu
        if list_aliases and len(list_aliases) > 0:
            df_alias = pd.DataFrame(list_aliases).drop_duplicates(subset=['Trường', 'Biến thể (đầu vào)'])
        else:
            df_alias = pd.DataFrame(columns=['Trường', 'Biến thể (đầu vào)', 'Chuẩn (CPC1)', 'Mã sản phẩm', 'Tên sản phẩm', 'Ghi chú', 'Ngày thêm', 'Người thêm'])
            
        df_alias.to_excel(writer, index=False, sheet_name=alias_sheet_name)
    return output.getvalue()


# ==========================================
# GIAO DIỆN NGƯỜI DÙNG (UI)
# ==========================================
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>🏥 HỆ THỐNG ĐỐI CHIẾU DANH MỤC ĐẤU THẦU CPC1</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #64748b; font-size: 1.1rem; margin-top: -10px; margin-bottom: 25px;'>Đối chiếu Hoạt chất · Hàm lượng · Nhóm → Ánh xạ danh mục</p>", unsafe_allow_html=True)
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📋 1. Danh mục gốc CPC1")
    st.caption("File chuẩn, có thể chứa sheet 'Alias'")
    master_file = st.file_uploader("Tải tệp Excel gốc tại đây", type=["xlsx", "xls"], key=f"master_{st.session_state.uploader_key}")
    master_msg = st.empty() 

with col2:
    st.subheader("📄 2. Danh mục cần đối chiếu")
    st.caption("File chào thầu ghi tên cột: Hoạt chất, Hàm lượng, Nhóm")
    target_file = st.file_uploader("Tải tệp cần kiểm tra tại đây", type=["xlsx", "xls"], key=f"target_{st.session_state.uploader_key}")
    target_msg = st.empty() 

# ------------------------------------------
# XỬ LÝ FILE GỐC (MASTER) NGAY KHI UPLOAD
# ------------------------------------------
if master_file:
    file_id = f"{master_file.name}_{master_file.size}"
    if file_id != st.session_state.master_file_name:
        master_bytes = master_file.read()
        df_cpc1 = read_excel_auto_header(master_bytes, sheet_name=0)
        df_cpc1, missing_cpc1 = auto_map_headers(df_cpc1)
        
        if missing_cpc1:
            st.session_state.master_status = ("error", f"❌ File gốc CPC1 thiếu cột: {missing_cpc1}.")
            st.session_state.df_cpc1 = None
        else:
            st.session_state.alias_hc = {}
            st.session_state.alias_hl = {}
            st.session_state.list_aliases = []
            
            has_alias, chc, chl = load_alias_from_excel(master_bytes)
            
            st.session_state.df_cpc1 = df_cpc1
            st.session_state.master_file_name = file_id
            
            if has_alias and (chc > 0 or chl > 0):
                st.session_state.master_status = ("success", f"✅ Đã nạp CPC1 và tìm thấy Alias: {chc} Hoạt chất, {chl} Hàm lượng.")
            else:
                st.session_state.master_status = ("success", "✅ Đã nạp danh mục gốc CPC1.")
                
    if 'master_status' in st.session_state:
        m_type, m_text = st.session_state.master_status
        if m_type == "error": master_msg.error(m_text)
        else: master_msg.success(m_text)

# ------------------------------------------
# XỬ LÝ FILE ĐỐI CHIẾU (TARGET) NGAY KHI UPLOAD
# ------------------------------------------
if target_file:
    t_file_id = f"{target_file.name}_{target_file.size}"
    if t_file_id != st.session_state.target_file_name:
        target_bytes = target_file.read()
        df_target = read_excel_auto_header(target_bytes, sheet_name=0)
        df_target, missing_target = auto_map_headers(df_target)
        df_target = df_target.reset_index(drop=True)
        
        if missing_target:
            st.session_state.target_status = ("error", f"❌ File đối chiếu thiếu cột: {missing_target}.")
            st.session_state.df_target = None
        else:
            st.session_state.df_target = df_target
            st.session_state.target_status = ("success", f"✅ Đã nạp danh mục cần đối chiếu: **{target_file.name}**")
            st.session_state.target_file_name = t_file_id

    if 'target_status' in st.session_state:
        t_type, t_text = st.session_state.target_status
        if t_type == "error": target_msg.error(t_text)
        else: target_msg.success(t_text)

# Cảnh báo được dời lên ngay dưới khung tải file
if not master_file or not target_file:
    st.warning("⚠️ Vui lòng tải lên ĐẦY ĐỦ cả 2 tệp (Danh mục gốc và Danh mục đối chiếu) để bắt đầu.")

if master_file and target_file:
    st.markdown("---")
    
    # Logic kiểm tra xem có được bấm nút Start không
    can_start = True
    if 'master_status' in st.session_state and st.session_state.master_status[0] == "error": can_start = False
    if 'target_status' in st.session_state and st.session_state.target_status[0] == "error": can_start = False

    col_opt, col_btn_start, col_btn_reset = st.columns([2, 1, 1])
    
    with col_opt:
        wratio_threshold = st.slider("🎯 Ngưỡng chấp nhận sai khác chính tả/gốc muối (%)", min_value=60, max_value=95, value=90, step=5)

    with col_btn_start:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        # Nút BẮT ĐẦU sẽ bị mờ (disabled) nếu can_start = False
        start_btn = st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU", use_container_width=True, type="primary", disabled=not can_start)
        
    with col_btn_reset:
        st.markdown("""
        <style>
        div[data-testid="stButton"] button[kind="secondary"] {
            background-color: #16a34a !important;
            color: white !important;
            border-color: #16a34a !important;
        }
        div[data-testid="stButton"] button[kind="secondary"]:hover {
            background-color: #15803d !important;
            border-color: #15803d !important;
            color: white !important;
        }
        </style>
        <div style='margin-top: 28px;'></div>
        """, unsafe_allow_html=True)
        reset_btn = st.button("🔄 LÀM MỚI (RESET)", use_container_width=True, type="secondary")

    if reset_btn:
        # Bổ sung target_file_name và target_status vào danh sách xóa
        keys_to_clear = ['alias_hc', 'alias_hl', 'list_aliases', 'new_aliases', 'master_file_name', 'target_file_name', 'processed', 'df_results', 'df_target', 'df_cpc1', 'master_status', 'target_status']
        for k in keys_to_clear:
            if k in st.session_state:
                del st.session_state[k]
                
        st.session_state.uploader_key += 1
        
        if hasattr(st, 'rerun'): st.rerun()
        else: st.experimental_rerun()

    if start_btn:
        try:
            # Lấy thẳng dữ liệu đã đọc từ Session State, không tốn thời gian đọc lại file
            df_cpc1_current = st.session_state.df_cpc1 
            df_target_current = st.session_state.df_target

            # Chạy logic đối chiếu
            progress_bar = st.progress(0)
            status_text = st.empty()
            total_rows = len(df_target_current)
            results = []

            for index, row in df_target_current.iterrows():
                if index % max(1, total_rows // 100) == 0:
                    progress_bar.progress(min(index / total_rows, 1.0))
                    status_text.text(f"Đang xử lý: {index}/{total_rows} dòng...")

                res = process_single_row(row, df_cpc1_current, wratio_threshold)
                res_with_stt = {"STT": index + 1}
                res_with_stt.update(res)
                results.append(res_with_stt)
            
            progress_bar.empty()
            status_text.empty()
            
            df_results = pd.DataFrame(results)
            cat_order = ["100%", "80% - 99%", "50% - 79%", "Dưới 50%", "N/A"]
            df_results['Tỉ lệ khớp'] = pd.Categorical(df_results['Tỉ lệ khớp'], categories=cat_order, ordered=True)
            st.session_state.df_results = df_results
            st.session_state.processed = True

        except Exception as e:
            st.error(f"Đã xảy ra lỗi hệ thống: {str(e)}")

# ==========================================
# KHU VỰC HIỂN THỊ KẾT QUẢ VÀ HỌC MÁY (ALIAS)
# ==========================================
if st.session_state.processed:
    df_results = st.session_state.df_results
    df_target = st.session_state.df_target
    df_cpc1 = st.session_state.df_cpc1
    
    # ------------------------------------------
    # KHU VỰC THỐNG KÊ (METRICS)
    # ------------------------------------------
    st.markdown("### 📈 THỐNG KÊ TỔNG QUAN")
    
    # Thêm CSS custom để ép giao diện các Metric thành dạng Card
    st.markdown("""
    <style>
    [data-testid="stMetric"] {
        background: white; 
        border-radius: 10px; 
        padding: 14px 10px;
        border: 1px solid #e2e8f0; 
        box-shadow: 0 1px 3px rgba(0,0,0,0.05); /* Thêm đổ bóng nhẹ cho đẹp */
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
    }
    [data-testid="stMetricLabel"] {
        display: flex !important;
        justify-content: center !important;
        width: 100% !important;
    }
    [data-testid="stMetricLabel"] > div {
        text-align: center !important;
    }
    [data-testid="stMetricValue"] {
        text-align: center !important;
        width: 100% !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    counts = df_results['Tỉ lệ khớp'].value_counts()
    
    # Hiển thị trên 1 hàng 7 cột, rút gọn text để không bị tràn chữ
    metric_cols = st.columns(7)
    metric_cols[0].metric("📚 Gốc CPC1", len(df_cpc1))
    metric_cols[1].metric("🎯 Cần đối chiếu", len(df_target))
    metric_cols[2].metric("✅ 100%", counts.get("100%", 0))
    metric_cols[3].metric("🟢 80% - 99%", counts.get("80% - 99%", 0))
    metric_cols[4].metric("🟡 50% - 79%", counts.get("50% - 79%", 0))
    metric_cols[5].metric("🔴 Dưới 50%", counts.get("Dưới 50%", 0))
    metric_cols[6].metric("⚪ N/A (Lỗi)", counts.get("N/A", 0))

    # --- Thống kê theo Hãng ---
    # Chỉ lấy các dòng có tỉ lệ khớp từ 80% trở lên
    high_match_df = df_results[df_results['Tỉ lệ khớp'].isin(['100%', '80% - 99%'])]
    brand_counts = high_match_df['Hãng chào'].value_counts()
    
    # Lọc bỏ các giá trị trống hoặc "Không có thông tin" để đếm cho chuẩn
    brand_texts = [f"{brand}-{count}" for brand, count in brand_counts.items() if str(brand).strip() and brand != "Không có thông tin"]
    brand_display = ", ".join(brand_texts) if brand_texts else "Không có dữ liệu"
    
    st.markdown(f"<div style='margin-top: 15px; padding: 12px 18px; background-color: #f8fafc; border-radius: 10px; border: 1px solid #e2e8f0; color: #334155; font-size: 15px;'><b>🏢 Kết quả theo hãng (Khớp &ge; 80%):</b> {brand_display}</div>", unsafe_allow_html=True)

    # ------------------------------------------
    # HIỂN THỊ BẢNG KẾT QUẢ
    # ------------------------------------------
    st.markdown("### 📊 KẾT QUẢ ĐỐI CHIẾU")
    
    # NÚT XUẤT FILE & BỘ LỌC (Dời lên trên bảng để lọc dữ liệu trực tiếp)
    excel_res = export_to_excel(df_results, "Ket_Qua", st.session_state.list_aliases, "Alias_Cap_Nhat")
    
    col_filter, col_download = st.columns([5, 2])
    with col_filter:
        filter_options = ["100%", "80% - 99%", "50% - 79%", "Dưới 50%", "N/A"]
        # Đổi thành multiselect, mặc định hiển thị tất cả
        selected_filters = st.multiselect(
            "🔍 Lọc bảng theo Tỉ lệ khớp (Có thể chọn nhiều):", 
            options=filter_options,
            default=filter_options
        )

    with col_download:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        # Đổi type="primary" thành "secondary" để nút có nền trắng, viền nhạt nhẹ nhàng
        st.download_button("📥 TẢI XUỐNG KẾT QUẢ ĐỐI CHIẾU", excel_res, "Ket_Qua_Doi_Chieu.xlsx", type="secondary", use_container_width=True)

    # Sắp xếp 2 lớp: Theo Tỉ lệ khớp trước, sau đó theo STT tăng dần
    df_display = df_results.sort_values(['Tỉ lệ khớp', 'STT']).reset_index(drop=True)
    df_display['Tỉ lệ khớp'] = df_display['Tỉ lệ khớp'].astype(str)

    # Áp dụng bộ lọc đa luồng
    if selected_filters:
        df_display = df_display[df_display['Tỉ lệ khớp'].isin(selected_filters)]
    else:
        # Nếu người dùng xóa hết các tag lọc thì hiển thị bảng trống
        df_display = pd.DataFrame(columns=df_display.columns)

    def highlight_similarity(row):
        val = row['Tỉ lệ khớp']
        styles = [''] * len(row)
        if val == '100%': 
            styles = ['background-color: #dcfce7; color: #166534;'] * len(row)
        elif val == '80% - 99%': 
            styles = ['background-color: #fef08a; color: #854d0e;'] * len(row)
        elif val == '50% - 79%': 
            styles = ['background-color: #ffedd5; color: #c2410c;'] * len(row)
        elif val == 'Dưới 50%': 
            # Giữ nguyên nền trắng cho dòng bị lỗi, chỉ bôi đỏ riêng ô Tỉ lệ khớp
            col_idx = row.index.get_loc('Tỉ lệ khớp')
            styles[col_idx] = 'background-color: #fee2e2; color: #b91c1c;'
        return styles

    # Thay map() bằng apply(axis=1) để áp dụng định dạng cho toàn bộ dòng ngang
    st.dataframe(
        df_display.style.apply(highlight_similarity, axis=1), 
        width='stretch', 
        height=400, 
        hide_index=True,
        column_config={
            "STT": st.column_config.NumberColumn(width="small"),
            "Mã phần": st.column_config.TextColumn(width="small"),
            "Hãng chào": st.column_config.TextColumn(width="small"),
            "Tỉ lệ khớp": st.column_config.TextColumn(width="small")
        }
    )

    # ------------------------------------------
    # KHU VỰC DẠY HỆ THỐNG
    # ------------------------------------------
    st.markdown("---")
    st.markdown("### 🎓 DẠY HỆ THỐNG (TẠO ALIAS ÁNH XẠ)")
    st.markdown("Chọn dòng bị lệch và ghép nối với sản phẩm CPC1 đúng. Dữ liệu sẽ được lưu tự động xuống bảng bên dưới.")

    failed_mask = df_results['Tỉ lệ khớp'].isin(['Dưới 50%', '50% - 79%', '80% - 99%'])
    failed_indices = df_results[failed_mask].index.tolist()

    if not failed_indices:
        st.success("🎉 Tuyệt vời! Toàn bộ dữ liệu đều khớp 100%, không có dòng nào cần ghép nối thủ công.")
    else:
        def format_input(idx):
            res_row = df_results.iloc[idx]
            return f"Dòng {idx+1}: {res_row['Dữ liệu Danh mục cần đối chiếu']} ➔ {res_row['Tỉ lệ khớp']}"
            
        def format_cpc1(idx):
            row = df_cpc1.iloc[idx]
            return f"{row.get('Hoạt_chất')} | {row.get('Hàm_lượng')} | Nhóm {row.get('Nhóm')}"

        with st.container(border=True):
            # Chia làm 2 cột với tỉ lệ 5:2
            col_left, col_right = st.columns([5, 2])
            
            with col_left:
                input_idx = st.selectbox("🔎 1. Dòng dữ liệu đối chiếu (bị lệch):", failed_indices, format_func=format_input)
                cpc1_idx = st.selectbox("✅ 2. Sản phẩm chuẩn CPC1 tương ứng:", df_cpc1.index.tolist(), format_func=format_cpc1)
                
            with col_right:
                nguoi_them_input = st.text_input("👤 3. Người thêm:", value="", placeholder="Ví dụ: N.V.A")
                
                # Thêm khoảng đệm để đẩy nút bấm xuống ngang hàng với ô chọn thứ 2 bên trái
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                clicked_ghep_noi = st.button("🔗 GHÉP NỐI & TẠO ALIAS", type="primary", use_container_width=True)

            if clicked_ghep_noi:
                target_row = df_target.iloc[input_idx]
                cpc1_row = df_cpc1.iloc[cpc1_idx]
                res_row = df_results.iloc[input_idx] 
                
                raw_hc_input = str(target_row.get('Hoạt_chất', '')).strip()
                raw_hl_input = str(target_row.get('Hàm_lượng', '')).strip()
                
                std_hc = str(cpc1_row.get('Hoạt_chất', '')).strip()
                std_hl = str(cpc1_row.get('Hàm_lượng', '')).strip()
                
                ma_sp = str(cpc1_row.get('Mã_sản_phẩm', '')).strip()
                if ma_sp.lower() == 'nan': ma_sp = ""
                
                ten_sp = str(cpc1_row.get('Tên_thuốc', '')).strip()
                if ten_sp.lower() == 'nan': ten_sp = ""
                
                actual_note = str(res_row.get('Ghi chú / Giải trình', '')).strip()
                today_str = date.today().strftime("%Y-%m-%d")
                nguoi_them_val = nguoi_them_input.strip() # Lấy giá trị Người thêm từ ô input
                
                added = []
                
                # Cập nhật Alias cho Hoạt Chất
                norm_hc_in = remove_accents(raw_hc_input)
                norm_hc_std = remove_accents(std_hc)
                if norm_hc_in != norm_hc_std and raw_hc_input and std_hc:
                    if norm_hc_in not in st.session_state.alias_hc:
                        st.session_state.alias_hc[norm_hc_in] = std_hc
                        new_data = {
                            "Trường": "Hoạt chất", "Biến thể (đầu vào)": raw_hc_input, "Chuẩn (CPC1)": std_hc,
                            "Mã sản phẩm": ma_sp, "Tên sản phẩm": ten_sp, "Ghi chú": actual_note,
                            "Ngày thêm": today_str, "Người thêm": nguoi_them_val
                        }
                        st.session_state.list_aliases.append(new_data)
                        st.session_state.new_aliases.append(new_data) # Nạp vào bộ nhớ Alias Mới
                        added.append("Hoạt chất")
                        
                # Cập nhật Alias cho Hàm Lượng
                norm_hl_in = remove_accents(raw_hl_input)
                norm_hl_std = remove_accents(std_hl)
                if norm_hl_in != norm_hl_std and raw_hl_input and std_hl:
                    if norm_hl_in not in st.session_state.alias_hl:
                        st.session_state.alias_hl[norm_hl_in] = std_hl
                        new_data = {
                            "Trường": "Hàm lượng", "Biến thể (đầu vào)": raw_hl_input, "Chuẩn (CPC1)": std_hl,
                            "Mã sản phẩm": ma_sp, "Tên sản phẩm": ten_sp, "Ghi chú": actual_note,
                            "Ngày thêm": today_str, "Người thêm": nguoi_them_val
                        }
                        st.session_state.list_aliases.append(new_data)
                        st.session_state.new_aliases.append(new_data) # Nạp vào bộ nhớ Alias Mới
                        added.append("Hàm lượng")
                        
                if added:
                    st.success(f"✅ Đã ghi nhận Alias cho [{', '.join(added)}] vào bảng chờ bên dưới.")
                else:
                    st.warning("⚠️ Alias này đã có sẵn trong danh sách hoặc hai sản phẩm đã giống hệt nhau.")

    # HIỂN THỊ BẢNG ALIAS ĐỂ COPY (CHỈ HIỆN KHI CÓ ALIAS MỚI)
    if st.session_state.new_aliases:
        st.markdown("### 📋 ÁNH XẠ: COPY THÊM VÀO ALIAS GỐC")
        st.markdown("*(Hệ thống đã tự động loại bỏ tiêu đề cột để bạn dán nối tiếp dễ dàng hơn)*")
        
        # 1. Khung COPY dữ liệu thô (TSV) chống lỗi định dạng Excel
        df_new_alias = pd.DataFrame(st.session_state.new_aliases).drop_duplicates(subset=['Trường', 'Biến thể (đầu vào)'])
        cols_to_show = ['Trường', 'Biến thể (đầu vào)', 'Chuẩn (CPC1)', 'Mã sản phẩm', 'Tên sản phẩm', 'Ghi chú', 'Ngày thêm', 'Người thêm']
        for c in cols_to_show:
            if c not in df_new_alias.columns: df_new_alias[c] = ""
        df_new_alias = df_new_alias[cols_to_show]
        
        # Chuyển Dataframe thành văn bản phân tách bằng phím Tab (chuẩn Excel) & Ẩn đi Header gốc
        tsv_data = df_new_alias.to_csv(index=False, header=False, sep='\t')
        
        st.info("👉 **BƯỚC 1:** Nhấn vào biểu tượng **Copy** (hình 2 ô vuông) ở góc trên bên phải khung xám dưới đây:")
        st.code(tsv_data, language="text")
        
        st.info("👉 **BƯỚC 2:** Mở file Excel danh mục gốc CPC1, click vào ô trống đầu tiên ở dưới cùng của sheet `Alias` và nhấn **Ctrl+V**.")
        st.info("👉 **BƯỚC 3:** Lưu file Excel lại, quay về đây và nhấn nút **🔄 LÀM MỚI (RESET)** ở bên trên.")
        
        # 2. Hiển thị bảng để người dùng nhìn cho trực quan
        st.markdown("<br><i>📚 Giao diện hiển thị trước các dòng dữ liệu bạn vừa copy:</i>", unsafe_allow_html=True)
        st.dataframe(df_new_alias, width='stretch', hide_index=True)

# ==========================================
# FOOTER / CREDIT
# ==========================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #64748b; padding: 20px 0; font-size: 15px;'>
        Thiết kế: <b>Nguyễn Văn Dũng</b> ❤️ 0978.777.191
    </div>
    """, 
    unsafe_allow_html=True
)