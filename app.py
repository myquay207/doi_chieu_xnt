import io, re, warnings
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
#  CẤU HÌNH GIAO DIỆN
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Đối Chiếu Dược – BV Đà Nẵng", page_icon="🏥", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;600;700&display=swap');
    html,body,[class*="css"]{font-family:'Be Vietnam Pro',sans-serif;}
    .hero{background:linear-gradient(135deg,#1a3a5c 0%,#2563a8 60%,#1e7fcb 100%);
      border-radius:16px;padding:32px;margin-bottom:24px;color:white;text-align:center;}
    .stat-card{background:white;border:1px solid #e2e8f0;border-radius:12px;padding:16px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.05);}
    .num{font-size:1.8rem;font-weight:700;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  LOGIC XỬ LÝ FILE LINH HOẠT (KHÔNG PHỤ THUỘC SỐ DÒNG)
# ══════════════════════════════════════════════════════════════════════════════

def norm(s):
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    return re.sub(r'\s+', ' ', s).strip()

def flexible_read_excel(file_content, keywords=["tên thuốc", "số lượng", "sổ sách"]):
    """Tìm dòng tiêu đề dựa trên từ khóa và lọc dữ liệu số lượng != 0"""
    try:
        df_raw = pd.read_excel(io.BytesIO(file_content), header=None)
        header_idx = -1
        # Tìm dòng chứa ít nhất 2 từ khóa
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
            matches = sum(1 for k in keywords if k in row_str)
            if matches >= 2:
                header_idx = i
                break
        
        if header_idx == -1: return pd.DataFrame()

        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = [str(c).strip().lower() for c in df_raw.iloc[header_idx]]
        
        # Chỉ lấy dòng có STT là số
        df = df[pd.to_numeric(df.iloc[:, 0], errors='coerce').notna()]
        return df
    except:
        return pd.DataFrame()

# ══════════════════════════════════════════════════════════════════════════════
#  HÀM ĐỐI CHIẾU CHÍNH (ĐÃ FIX LỖI SỐ 0 VÀ TYPEERROR)
# ══════════════════════════════════════════════════════════════════════════════

def run_reconciliation(df_source, df_target, type_mode="XNT"):
    """
    df_source: Dữ liệu cần đối chiếu (XNT thô, BBKK, hoặc BBKN)
    df_target: Dữ liệu chuẩn (Thống kê)
    """
    results = []
    used_target = set()

    # Lọc bỏ các dòng có số lượng = 0 ngay từ đầu để làm sạch báo cáo
    # Giả sử cột số lượng là 'tồn cuối', 'thực tế' hoặc cột có chứa chữ 'số lượng'
    def get_qty_col(df):
        for c in df.columns:
            if any(k in str(c) for k in ['tồn', 'thực tế', 'số lượng', 'quantity']): return c
        return df.columns[-1]

    col_src = get_qty_col(df_source)
    col_tgt = get_qty_col(df_target)

    # Chuyển đổi số lượng về numeric để so sánh
    df_source[col_src] = pd.to_numeric(df_source[col_src], errors='coerce').fillna(0)
    df_target[col_tgt] = pd.to_numeric(df_target[col_tgt], errors='coerce').fillna(0)

    # Lọc != 0
    df_source = df_source[df_source[col_src] != 0].copy()
    df_target_active = df_target[df_target[col_tgt] != 0].copy()

    # Logic khớp theo Tên + Đơn giá (vì thô thường không có mã)
    for idx_s, row_s in df_source.iterrows():
        ten_s = norm(row_s.get('tên thuốc', row_s.iloc[1]))
        gia_s = float(row_s.get('đơn giá', 0))
        qty_s = float(row_s[col_src])

        matched = False
        for idx_t, row_t in df_target_active.iterrows():
            if idx_t in used_target: continue
            
            ten_t = norm(row_t.get('tên thuốc', row_t.iloc[1]))
            gia_t = float(row_t.get('đơn giá', 0))
            
            # Khớp tên và giá (cho phép sai lệch giá nhỏ do làm tròn)
            if ten_s == ten_t and abs(gia_s - gia_t) < 10:
                qty_t = float(row_t[col_tgt])
                results.append({
                    'ten': row_s.get('tên thuốc', row_s.iloc[1]),
                    'gia': gia_s,
                    'qty_src': qty_s,
                    'qty_tgt': qty_t,
                    'cl': qty_s - qty_t,
                    'status': 'Khớp' if abs(qty_s - qty_t) < 0.01 else 'Lệch'
                })
                used_target.add(idx_t)
                matched = True
                break
        
        if not matched:
            results.append({
                'ten': row_s.get('tên thuốc', row_s.iloc[1]),
                'gia': gia_s, 'qty_src': qty_s, 'qty_tgt': 0, 'cl': qty_s, 'status': 'Chỉ có bên gửi'
            })

    # Các dòng chỉ có bên Thống kê
    for idx_t, row_t in df_target_active.iterrows():
        if idx_t not in used_target:
            results.append({
                'ten': row_t.get('tên thuốc', row_t.iloc[1]),
                'gia': row_t.get('đơn giá', 0),
                'qty_src': 0, 'qty_tgt': row_t[col_tgt], 'cl': -row_t[col_tgt], 'status': 'Chỉ có Thống kê'
            })

    return pd.DataFrame(results)

# ══════════════════════════════════════════════════════════════════════════════
#  GIAO DIỆN STREAMLIT
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="hero"><h1>🏥 HỆ THỐNG ĐỐI CHIẾU DƯỢC LINH HOẠT</h1><p>Tự động tìm cấu trúc file | Lọc bỏ tồn bằng 0 | Hỗ trợ BBKK & BBKN</p></div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📊 Đối chiếu dữ liệu", "⚙️ Hướng dẫn"])

with tab1:
    col_l, col_r = st.columns(2)
    with col_l:
        f_src = st.file_uploader("📂 Tải lên File cần đối chiếu (XNT thô / BBKK / BBKN)", type=["xlsx", "xls"])
    with col_r:
        f_tgt = st.file_uploader("📋 Tải lên File Thống kê (Chuẩn)", type=["xlsx", "xls"])

    if f_src and f_tgt:
        if st.button("🚀 Bắt đầu đối chiếu"):
            df_s = flexible_read_excel(f_src.read())
            df_t = flexible_read_excel(f_tgt.read())

            if df_s.empty or df_t.empty:
                st.error("❌ Không tìm thấy dòng tiêu đề phù hợp trong file. Hãy kiểm tra lại file có chứa các từ khóa như 'Tên thuốc', 'Số lượng', 'Sổ sách' không?")
            else:
                res = run_reconciliation(df_s, df_t)
                
                # Hiển thị thống kê
                c1, c2, c3 = st.columns(3)
                c1.metric("Khớp hoàn toàn", len(res[res['status']=='Khớp']))
                c2.metric("Chênh lệch", len(res[res['status']=='Lệch']))
                c3.metric("Chỉ có một bên", len(res[res['status'].str.contains('Chỉ có')]))

                st.dataframe(res, use_container_width=True)

                # Download Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    res.to_excel(writer, index=False, sheet_name='KetQua')
                st.download_button("📥 Tải kết quả về Excel", output.getvalue(), "ket_qua_doi_chieu.xlsx")
