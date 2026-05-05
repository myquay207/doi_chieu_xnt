import io, re, warnings
import streamlit as st
import pandas as pd

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Đối Chiếu Dược – BV Đà Nẵng", page_icon="🏥", layout="wide")

# --- HÀM CHUẨN HÓA TÊN ---
def norm(s):
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    return re.sub(r'\s+', ' ', s).strip()

# --- HÀM ĐỌC FILE CỰC KỲ LINH HOẠT ---
def flexible_read_excel(file_content):
    try:
        # Đọc hết các sheet, ưu tiên sheet có dữ liệu
        all_sheets = pd.read_excel(io.BytesIO(file_content), header=None, sheet_name=None)
        df_raw = list(all_sheets.values())[0] # Lấy sheet đầu tiên
        
        header_idx = -1
        # Từ khóa mở rộng để không bị sót file nào của bệnh viện
        keywords = ["tên", "đơn giá", "thành tiền", "sổ sách", "thực tế", "số lượng"]
        
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
            # Chỉ cần khớp 2 từ khóa bất kỳ là nhận diện được dòng tiêu đề
            if sum(1 for k in keywords if k in row_str) >= 2:
                header_idx = i
                break
        
        if header_idx == -1: return pd.DataFrame()

        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = [str(c).strip().lower() for c in df_raw.iloc[header_idx]]
        
        # Loại bỏ dòng rỗng và dòng không có STT (số)
        df = df[pd.to_numeric(df.iloc[:, 0], errors='coerce').notna()]
        return df
    except:
        return pd.DataFrame()

# --- HÀM ĐỐI CHIẾU ---
def run_reconciliation(df_src, df_tgt):
    # Tự tìm cột số lượng (vì mỗi file đặt tên mỗi kiểu: Tồn, Thực tế, Số lượng...)
    def find_qty_col(df):
        cols = df.columns
        for c in cols:
            if any(k in str(c) for k in ['thực tế', 'sổ sách', 'số lượng', 'tồn', 'nhập']): return c
        return cols[-1]

    col_s = find_qty_col(df_src)
    col_t = find_qty_col(df_tgt)

    # Ép kiểu số và lọc bỏ số 0 (như bạn yêu cầu)
    df_src[col_s] = pd.to_numeric(df_src[col_s], errors='coerce').fillna(0)
    df_tgt[col_t] = pd.to_numeric(df_tgt[col_t], errors='coerce').fillna(0)
    
    df_src = df_src[df_src[col_s] != 0].copy()
    df_tgt = df_tgt[df_tgt[col_t] != 0].copy()

    # Tiến hành so khớp
    res = []
    used_tgt = set()

    for _, rs in df_src.iterrows():
        name_s = norm(rs.iloc[1]) # Thường tên thuốc ở cột 2
        qty_s = float(rs[col_s])
        
        found = False
        for it, rt in df_tgt.iterrows():
            if it in used_tgt: continue
            if name_s == norm(rt.iloc[1]):
                qty_t = float(rt[col_t])
                res.append({'Tên thuốc': rs.iloc[1], 'Số lượng bên gửi': qty_s, 'Số lượng Thống kê': qty_t, 'Chênh lệch': qty_s - qty_t})
                used_tgt.add(it)
                found = True
                break
        if not found:
            res.append({'Tên thuốc': rs.iloc[1], 'Số lượng bên gửi': qty_s, 'Số lượng Thống kê': 0, 'Chênh lệch': qty_s})

    for it, rt in df_tgt.iterrows():
        if it not in used_tgt:
            res.append({'Tên thuốc': rt.iloc[1], 'Số lượng bên gửi': 0, 'Số lượng Thống kê': rt[col_t], 'Chênh lệch': -rt[col_t]})
            
    return pd.DataFrame(res)

# --- GIAO DIỆN ---
st.title("🏥 Đối chiếu Dược - BV Đà Nẵng (Bản Fix)")

f1 = st.file_uploader("1. Tải file BBKK / BBKN / XNT thô", type=["xlsx", "xls"])
f2 = st.file_uploader("2. Tải file Thống kê (Chuẩn)", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("Bắt đầu đối chiếu"):
        d1 = flexible_read_excel(f1.read())
        d2 = flexible_read_excel(f2.read())
        
        if d1.empty or d2.empty:
            st.error("Vẫn không tìm thấy tiêu đề. Bạn kiểm tra lại sheet chứa dữ liệu có phải là sheet đầu tiên không?")
        else:
            final = run_reconciliation(d1, d2)
            st.success(f"Đã đối chiếu xong! Tìm thấy {len(final)} mặt hàng có phát sinh tồn/nhập.")
            st.dataframe(final, use_container_width=True)
