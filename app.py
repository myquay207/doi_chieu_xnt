import io, re, warnings
import streamlit as st
import pandas as pd

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Đối Chiếu Dược – BV Đà Nẵng", page_icon="🏥", layout="wide")

def norm(s):
    if not isinstance(s, str): return ""
    return re.sub(r'\s+', ' ', s.strip().lower())

def find_header_and_read(file_bytes):
    try:
        # Đọc file, nạp tất cả các sheet
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        # Ưu tiên tìm trong sheet có tên 'Sheet1' hoặc sheet đầu tiên
        df_raw = pd.read_excel(xls, sheet_name=0, header=None)
        
        keywords = ['tên thuốc', 'tên hoạt chất', 'đơn giá', 'sổ sách', 'thực tế', 'số lượng']
        header_idx = -1
        
        # Quét từng dòng để tìm dòng tiêu đề
        for i, row in df_raw.iterrows():
            # Chuyển cả dòng thành danh sách các chuỗi đã chuẩn hóa
            row_vals = [norm(str(x)) for x in row if pd.notna(x)]
            # Nếu dòng này chứa ít nhất 1 từ khóa quan trọng
            if any(k in " ".join(row_vals) for k in keywords):
                header_idx = i
                break
        
        if header_idx == -1: return pd.DataFrame()

        # Lấy dữ liệu từ dòng tìm được
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = df_raw.iloc[header_idx]
        
        # Làm sạch tên cột: bỏ xuống dòng, bỏ khoảng trắng thừa
        df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
        
        # Chỉ giữ lại dòng có STT hoặc Tên thuốc (loại bỏ dòng ký tên, dòng tiêu đề phụ)
        df = df[df.iloc[:, 0].notna() | df.iloc[:, 1].notna()]
        return df
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return pd.DataFrame()

def run_reconciliation(df1, df2):
    # Tự động nhận diện cột số lượng (Tồn cuối / Thực tế / Sổ sách)
    def get_col(df, options):
        for c in df.columns:
            if any(o in str(c).lower() for o in options): return c
        return df.columns[-1]

    col1 = get_col(df1, ['thực tế', 'số lượng', 'tồn', 'nhập'])
    col2 = get_col(df2, ['sổ sách', 'thực tế', 'số lượng', 'tồn'])

    # Ép kiểu số
    df1[col1] = pd.to_numeric(df1[col1], errors='coerce').fillna(0)
    df2[col2] = pd.to_numeric(df2[col2], errors='coerce').fillna(0)

    # Lọc bỏ các dòng tồn bằng 0 như anh yêu cầu
    df1 = df1[df1[col1] != 0]
    df2 = df2[df2[col2] != 0]

    # So khớp theo Tên thuốc (thường nằm ở cột 1 hoặc 2)
    # Ở đây tôi dùng logic: nếu không tìm thấy cột 'Tên thuốc', lấy cột thứ 2
    name_col1 = next((c for c in df1.columns if 'tên' in str(c).lower()), df1.columns[1])
    name_col2 = next((c for c in df2.columns if 'tên' in str(c).lower()), df2.columns[1])

    results = []
    matched_idx2 = set()

    for i1, r1 in df1.iterrows():
        n1 = norm(str(r1[name_col1]))
        q1 = float(r1[col1])
        found = False
        
        for i2, r2 in df2.iterrows():
            if i2 in matched_idx2: continue
            if n1 == norm(str(r2[name_col2])):
                q2 = float(r2[col2])
                results.append({'Tên thuốc': r1[name_col1], 'Số lượng Gửi': q1, 'Số lượng Chuẩn': q2, 'Chênh lệch': q1 - q2})
                matched_idx2.add(i2)
                found = True
                break
        if not found:
            results.append({'Tên thuốc': r1[name_col1], 'Số lượng Gửi': q1, 'Số lượng Chuẩn': 0, 'Chênh lệch': q1})

    for i2, r2 in df2.iterrows():
        if i2 not in matched_idx2:
            results.append({'Tên thuốc': r2[name_col2], 'Số lượng Gửi': 0, 'Số lượng Chuẩn': r2[col2], 'Chênh lệch': -r2[col2]})

    return pd.DataFrame(results)

# --- GIAO DIỆN ---
st.markdown("### 🏥 Đối chiếu dữ liệu Dược (Bản ổn định)")

f1 = st.file_uploader("1. Tải file Thô (BBKK / BBKN / XNT)", type=["xlsx", "xls"])
f2 = st.file_uploader("2. Tải file Thống kê (Chuẩn)", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Bắt đầu đối chiếu"):
        d1 = find_header_and_read(f1.read())
        d2 = find_header_and_read(f2.read())
        
        if d1.empty or d2.empty:
            st.warning("⚠️ Không tìm thấy bảng dữ liệu. Hãy kiểm tra lại file Excel của anh.")
        else:
            final = run_reconciliation(d1, d2)
            st.success(f"Xử lý xong {len(final)} mặt hàng!")
            st.dataframe(final, use_container_width=True)
