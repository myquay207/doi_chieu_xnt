"""
ĐỐI CHIẾU XUẤT NHẬP TỒN – Bệnh viện Đà Nẵng
Ứng dụng Streamlit độc lập – chạy thử nghiệm trước khi merge vào app chính
"""

import io, re, warnings
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Đối Chiếu XNT – BV Đà Nẵng",
    page_icon="🏥",
    layout="centered",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;600;700&display=swap');
html,body,[class*="css"]{font-family:'Be Vietnam Pro',sans-serif;}
.hero{background:linear-gradient(135deg,#1a3a5c 0%,#2563a8 60%,#1e7fcb 100%);
  border-radius:16px;padding:32px 36px 24px;margin-bottom:24px;color:white;
  box-shadow:0 8px 32px rgba(37,99,168,.25);}
.hero h1{font-size:1.5rem;font-weight:700;margin:0 0 6px;line-height:1.3;}
.hero .sub{font-size:.88rem;font-weight:300;opacity:.85;margin:0;}
.hero .badge{display:inline-block;background:rgba(255,255,255,.18);border-radius:20px;
  padding:3px 12px;font-size:.75rem;font-weight:600;letter-spacing:1px;
  margin-bottom:12px;text-transform:uppercase;}
.info-box{background:#eff6ff;border-left:4px solid #2563a8;border-radius:0 10px 10px 0;
  padding:12px 16px;margin:12px 0 18px;font-size:.86rem;color:#1e3a5f;line-height:1.7;}
.warn-box{background:#fff7ed;border-left:4px solid #f59e0b;border-radius:0 10px 10px 0;
  padding:12px 16px;margin:12px 0;font-size:.85rem;color:#92400e;line-height:1.6;}
.stat-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:16px 0;}
.stat-card{background:white;border:1px solid #e2e8f0;border-radius:12px;
  padding:16px 12px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.06);}
.stat-card .num{font-size:1.6rem;font-weight:700;line-height:1;}
.stat-card .lbl{font-size:.72rem;color:#64748b;margin-top:4px;}
.ok-box{background:#f0fdf4;border:1.5px solid #86efac;border-radius:12px;
  padding:16px 20px;margin:16px 0;text-align:center;}
.ok-box h3{color:#166534;margin:6px 0 4px;font-size:1rem;}
.section-label{font-size:.72rem;font-weight:700;letter-spacing:1.5px;
  text-transform:uppercase;color:#94a3b8;margin-bottom:8px;}
.stButton>button{background:linear-gradient(135deg,#1a3a5c,#2563a8)!important;
  color:white!important;font-weight:600!important;font-size:.95rem!important;
  border:none!important;border-radius:10px!important;padding:13px 0!important;
  width:100%!important;box-shadow:0 4px 14px rgba(37,99,168,.3)!important;}
[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#166534,#16a34a)!important;
  color:white!important;font-weight:700!important;font-size:1rem!important;
  border:none!important;border-radius:10px!important;padding:15px 0!important;
  width:100%!important;}
[data-testid="stFileUploader"]{border:2px dashed #2563a8!important;
  border-radius:12px!important;background:#f0f6ff!important;}
hr{border:none;border-top:1px solid #e2e8f0;margin:20px 0;}
#MainMenu,footer{visibility:hidden;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  NORMALIZE
# ══════════════════════════════════════════════════════════════════════════════
def norm(s):
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = re.sub(r'_x000a_', ' ', s)
    s = re.sub(r'\n', ' ', s)
    s = re.sub(r'[#]', '', s)
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    s = re.sub(r'(\d)\s+(mg|ml|mcg|μg|g|iu|%|meq|l)', r'\1\2', s)
    s = re.sub(r'\s*\+\s*', '+', s)
    return re.sub(r'\s+', ' ', s).strip()

# ══════════════════════════════════════════════════════════════════════════════
#  EXTRACT MÃ TỪ FILE NHẬP/XUẤT
# ══════════════════════════════════════════════════════════════════════════════
def extract_ma(df):
    rows = []
    for _, row in df.iterrows():
        try: int(str(row[0]).strip())
        except: continue
        if pd.isna(row[1]) or pd.isna(row[2]): continue
        gia = row[5] if not pd.isna(row[5]) else 0
        rows.append({
            'ma':   str(row[1]).strip(),
            'ten':  str(row[2]).strip(),
            'nd':   str(row[3]).strip() if not pd.isna(row[3]) else '',
            'gia':  gia,
            'kten': norm(str(row[2])),
            'knd':  norm(str(row[3]) if not pd.isna(row[3]) else ''),
            'kgia': int(round(float(gia))) if gia else 0,
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()

# ══════════════════════════════════════════════════════════════════════════════
#  LOGIC ĐỐI CHIẾU CHÍNH
# ══════════════════════════════════════════════════════════════════════════════
def run_reconciliation(files_nhap_xuat, df_xnt_raw, df_tk_raw, confirmed_override):
    """
    files_nhap_xuat: list of DataFrames từ file nhập/xuất có mã
    df_xnt_raw:      DataFrame file XNT HPT (dữ liệu thô)
    df_tk_raw:       DataFrame file thống kê
    confirmed_override: dict {ten_thuoc: ma_chon} từ file xác nhận mã
    """

    # ── 1. Build bảng map mã ────────────────────────────────────────────
    df_all = pd.concat([extract_ma(df) for df in files_nhap_xuat
                        if not df.empty], ignore_index=True)
    if df_all.empty:
        return None, "Không đọc được dữ liệu từ file nhập/xuất"

    df_map = (df_all.groupby(['kten','knd','kgia'], as_index=False)
              .agg(ten     = ('ten',  'first'),
                   nd      = ('nd',   'first'),
                   gia     = ('gia',  'first'),
                   ma      = ('ma',   lambda x: x.value_counts().index[0]),
                   so_ma   = ('ma',   'nunique'),
                   ma_list = ('ma',   lambda x: '/'.join(x.unique()))))

    # Thêm Augmentin manual (tồn từ trước 2025)
    aug = pd.DataFrame([{
        'kten': norm('Augmentin 1g'), 'knd': norm('875mg + 125mg'),
        'kgia': 16680, 'ten': 'Augmentin 1g', 'nd': '875mg + 125mg',
        'gia': 16680, 'ma': '0005301225', 'so_ma': 1, 'ma_list': '0005301225'
    }])
    df_map = pd.concat([df_map, aug], ignore_index=True).drop_duplicates(
        subset=['kten','knd','kgia'], keep='last')

    # Override mã từ file xác nhận
    for i, row in df_map.iterrows():
        if row['ten'] in confirmed_override:
            df_map.at[i, 'ma'] = confirmed_override[row['ten']]

    # ── 2. Parse XNT (giữ nguyên từng dòng) ────────────────────────────
    xnt_rows = []
    for _, row in df_xnt_raw.iterrows():
        try: stt = int(str(row[0]).strip())
        except: continue
        if pd.isna(row[2]) or not isinstance(row[2],str) or row[2].strip().isdigit(): continue
        xnt_rows.append({
            'stt':     stt,
            'ten':     str(row[2]).strip(),
            'nd':      str(row[3]).strip() if not pd.isna(row[3]) else '',
            'gia':     row[8]  if not pd.isna(row[8])  else 0,
            'ton_xnt': float(row[12]) if not pd.isna(row[12]) else 0,
            'kten':    norm(str(row[2])),
            'knd':     norm(str(row[3]) if not pd.isna(row[3]) else ''),
            'kgia':    int(round(float(row[8]))) if not pd.isna(row[8]) else 0,
        })
    df_xnt = pd.DataFrame(xnt_rows)

    # ── 3. Parse TK ─────────────────────────────────────────────────────
    tk_rows = []
    for _, row in df_tk_raw.iloc[5:].iterrows():
        ma  = str(row[4]).strip() if not pd.isna(row[4]) else ''
        ten = str(row[5]).strip() if not pd.isna(row[5]) else ''
        if not ma or not ten: continue
        tk_rows.append({
            'ma':     ma,
            'ten_tk': ten,
            'nd_tk':  str(row[8]).strip() if not pd.isna(row[8]) else '',
            'gia_tk': row[11] if not pd.isna(row[11]) else 0,
            'ton_tk': float(row[24]) if not pd.isna(row[24]) else 0,
            'kten':   norm(ten),
            'knd':    norm(str(row[8]) if not pd.isna(row[8]) else ''),
            'kgia':   int(round(float(row[11]))) if not pd.isna(row[11]) else 0,
        })
    df_tk = pd.DataFrame(tk_rows)

    # ── 4. Logic ghép 3 bước ────────────────────────────────────────────
    results    = []
    used_tk    = set()

    for (kten, knd, kgia), grp_x in df_xnt.groupby(['kten','knd','kgia'], sort=False):
        # Tìm TK tương ứng chưa dùng
        mask   = (df_tk['kten']==kten) & (df_tk['knd']==knd) & (df_tk['kgia']==kgia)
        grp_t  = df_tk[mask & ~df_tk.index.isin(used_tk)].copy()

        def make_row(xr, tr, method):
            return {
                'ma': tr['ma'] if tr is not None else '',
                'ten': xr['ten'], 'nd': xr['nd'], 'gia': xr['gia'],
                'ton_xnt': xr['ton_xnt'],
                'ten_tk': tr['ten_tk'] if tr is not None else '',
                'nd_tk':  tr['nd_tk']  if tr is not None else '',
                'ton_tk': tr['ton_tk'] if tr is not None else None,
                'method': method,
                'cl': (xr['ton_xnt'] - tr['ton_tk']) if tr is not None else None,
            }

        if len(grp_x) == 1 and len(grp_t) == 1:
            xr = grp_x.iloc[0]; tr = grp_t.iloc[0]
            results.append(make_row(xr, tr, '1-1'))
            used_tk.add(grp_t.index[0])

        elif len(grp_x) >= 1 and len(grp_t) > 0:
            xl = grp_x.reset_index(drop=True)
            tl = grp_t.reset_index(drop=True)
            matched_x, matched_t = set(), set()

            # Bước A: khớp tồn chính xác
            for xi, xr in xl.iterrows():
                for ti, tr in tl.iterrows():
                    if ti in matched_t: continue
                    if abs(xr['ton_xnt'] - tr['ton_tk']) < 0.01:
                        results.append(make_row(xr, tr, 'exact_ton'))
                        used_tk.add(grp_t.index[ti])
                        matched_x.add(xi); matched_t.add(ti); break

            # Bước B: ghép gần nhất
            for xi, xr in xl.iterrows():
                if xi in matched_x: continue
                best_diff, best_ti, best_tr = float('inf'), None, None
                for ti, tr in tl.iterrows():
                    if ti in matched_t: continue
                    d = abs(xr['ton_xnt'] - tr['ton_tk'])
                    if d < best_diff:
                        best_diff, best_ti, best_tr = d, ti, tr
                if best_tr is not None:
                    results.append(make_row(xr, best_tr, 'nearest_ton'))
                    used_tk.add(grp_t.index[best_ti]); matched_t.add(best_ti)
                else:
                    results.append(make_row(xr, None, 'no_tk'))

        else:  # Không có TK
            for _, xr in grp_x.iterrows():
                results.append(make_row(xr, None, 'no_tk'))

   # TK có, XNT không (Chỉ lấy những dòng có tồn > 0)
    for idx, tr in df_tk.iterrows():
        # Thêm điều kiện: chưa dùng AND tồn phải khác 0 (hoặc > 0)
        if idx not in used_tk and abs(tr['ton_tk']) >= 0.01:
            results.append({
                'ma': tr['ma'], 'ten': '', 'nd': '', 'gia': '',
                'ton_xnt': None, 'ten_tk': tr['ten_tk'],
                'nd_tk': tr['nd_tk'], 'ton_tk': tr['ton_tk'],
                'method': 'no_xnt', 'cl': None,
            })

    return pd.DataFrame(results), None


# ══════════════════════════════════════════════════════════════════════════════
#  XUẤT EXCEL KẾT QUẢ
# ══════════════════════════════════════════════════════════════════════════════
def export_excel(df_res, thang_nam=""):
    THIN = Side(style='thin'); MED = Side(style='medium')
    def B():  return Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
    def BM(): return Border(left=THIN,right=THIN,top=MED,bottom=MED)

    F_H   = PatternFill('solid', fgColor='1F3864')
    F_OK  = PatternFill('solid', fgColor='E2EFDA')
    F_L   = PatternFill('solid', fgColor='FFD7D7')
    F_W   = PatternFill('solid', fgColor='FFF9C4')
    F_XT  = PatternFill('solid', fgColor='DEEBF7')
    F_TK  = PatternFill('solid', fgColor='FCE4D6')

    wb = Workbook()

    # ── Sheet 1: Kết quả đối chiếu ──────────────────────────────────────
    ws1 = wb.active; ws1.title = f"DC XNT {thang_nam.replace('/', '_')}"
    hdrs = [('Mã HPT',16),('Tên thuốc (XNT)',32),('Tên thuốc (TK)',28),
            ('Nồng độ',22),('Đơn giá',12),('Tồn cuối XNT',13),
            ('Tồn cuối TK',13),('Chênh lệch',13),('Trạng thái',25),('Phương pháp',18)]
    for ci,(h,w) in enumerate(hdrs,1):
        c = ws1.cell(row=1,column=ci,value=h)
        c.font = Font(name='Times New Roman',bold=True,size=11,color='FFFFFF')
        c.fill = F_H; c.border = BM()
        c.alignment = Alignment(horizontal='center',vertical='center',wrap_text=True)
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.row_dimensions[1].height = 36

    df_matched = df_res[df_res['cl'].notna()].copy()
    df_no_tk   = df_res[df_res['method']=='no_tk'].copy()
    df_no_xnt  = df_res[df_res['method']=='no_xnt'].copy()

    df_lech = df_matched[df_matched['cl'].abs()>=0.01].sort_values('cl')
    df_khop = df_matched[df_matched['cl'].abs()<0.01]
    df_s1   = pd.concat([df_lech, df_khop], ignore_index=True)

    method_map = {'1-1':'Chính xác 1-1','exact_ton':'Khớp tồn chính xác',
                  'nearest_ton':'Gần nhất theo tồn ⚠️'}

    for ri,(_,r) in enumerate(df_s1.iterrows(),2):
        cl = r['cl']
        if abs(cl) < 0.01:   fill, st = F_OK, '✅ Khớp'
        elif cl > 0:          fill, st = F_L,  f'⬆️ HPT cao hơn {cl:+.0f}'
        else:                  fill, st = F_L,  f'⬇️ HPT thấp hơn {cl:+.0f}'
        if r['method']=='nearest_ton' and abs(cl)>=0.01: fill = F_W

        vals = [r['ma'],r['ten'],r.get('ten_tk',''),r['nd'],r['gia'],
                r['ton_xnt'],r['ton_tk'],cl,st,method_map.get(r['method'],'')]
        for ci,v in enumerate(vals,1):
            cell = ws1.cell(row=ri,column=ci,value=v if not pd.isna(v) else '')
            cell.font = Font(name='Times New Roman',size=11)
            cell.fill = fill; cell.border = B()
            cell.alignment = Alignment(vertical='center',
                horizontal='right' if ci in (5,6,7,8) else
                'center' if ci in (1,9,10) else 'left')
            if ci in (6,7,8) and isinstance(v,(int,float)): cell.number_format='#,##0.##'
            if ci==5 and isinstance(v,(int,float)): cell.number_format='#,##0'
        ws1.row_dimensions[ri].height = 18
    ws1.freeze_panes = 'A2'

    # ── Sheet 2: XNT có / TK không ──────────────────────────────────────
    ws2 = wb.create_sheet("XNT có – TK không")
    for ci,(h,w) in enumerate([('Tên thuốc',35),('Nồng độ',25),
                                 ('Đơn giá',12),('Tồn XNT',12),('Ghi chú',35)],1):
        c=ws2.cell(row=1,column=ci,value=h)
        c.font=Font(name='Times New Roman',bold=True,size=11,color='FFFFFF')
        c.fill=F_H; c.border=BM(); c.alignment=Alignment(horizontal='center',vertical='center')
        ws2.column_dimensions[get_column_letter(ci)].width=w
    for ri,(_,r) in enumerate(df_no_tk.iterrows(),2):
        for ci,v in enumerate([r['ten'],r['nd'],r['gia'],r['ton_xnt'],
                                'HPT có nhưng TK không theo dõi'],1):
            c=ws2.cell(row=ri,column=ci,value=v if not pd.isna(v) else '')
            c.font=Font(name='Times New Roman',size=11); c.fill=F_XT; c.border=B()
            c.alignment=Alignment(vertical='center')
        ws2.row_dimensions[ri].height=18

    # ── Sheet 3: TK có / XNT không ──────────────────────────────────────
    ws3 = wb.create_sheet("TK có – XNT không")
    for ci,(h,w) in enumerate([('Mã HPT',16),('Tên thuốc TK',35),
                                 ('Nồng độ TK',25),('Tồn TK',12),('Ghi chú',35)],1):
        c=ws3.cell(row=1,column=ci,value=h)
        c.font=Font(name='Times New Roman',bold=True,size=11,color='FFFFFF')
        c.fill=F_H; c.border=BM(); c.alignment=Alignment(horizontal='center',vertical='center')
        ws3.column_dimensions[get_column_letter(ci)].width=w
    for ri,(_,r) in enumerate(df_no_xnt.iterrows(),2):
        for ci,v in enumerate([r['ma'],r['ten_tk'],r['nd_tk'],r['ton_tk'],
                                'TK có nhưng XNT không phát sinh'],1):
            c=ws3.cell(row=ri,column=ci,value=v if not pd.isna(v) else '')
            c.font=Font(name='Times New Roman',size=11); c.fill=F_TK; c.border=B()
            c.alignment=Alignment(vertical='center')
        ws3.row_dimensions[ri].height=18

    # ── Sheet 4: Tóm tắt ────────────────────────────────────────────────
    ws4 = wb.create_sheet("Tóm tắt")
    n_khop = (df_res['cl'].abs()<0.01).sum()
    n_lech = ((df_res['cl'].notna())&(df_res['cl'].abs()>=0.01)).sum()
    for ri,(k,v) in enumerate([
        (f'ĐỐI CHIẾU XNT {thang_nam}',''),('',''),
        ('✅ Khớp hoàn toàn',f'{n_khop} dòng'),
        ('⚠️  Chênh lệch',f'{n_lech} dòng'),
        ('📋 XNT có / TK không',f'{len(df_no_tk)} dòng'),
        ('📋 TK có / XNT không',f'{len(df_no_xnt)} dòng'),('',''),
        ('Tỷ lệ khớp tồn',
         f'{n_khop/(n_khop+n_lech)*100:.1f}%' if (n_khop+n_lech)>0 else 'N/A'),
    ],1):
        ws4.cell(row=ri,column=1,value=k).font=Font(
            name='Times New Roman',bold=(ri==1),size=12 if ri==1 else 11)
        ws4.cell(row=ri,column=2,value=v).font=Font(name='Times New Roman',size=11)
    ws4.column_dimensions['A'].width=35; ws4.column_dimensions['B'].width=20

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  GIAO DIỆN
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
  <div class="badge">🏥 Bệnh viện Đà Nẵng · Khoa Dược</div>
  <h1>ĐỐI CHIẾU XUẤT NHẬP TỒN<br>HPT vs THỐNG KÊ</h1>
  <p class="sub">Tự động ghép mã hàng · Đối chiếu tồn cuối · Phát hiện chênh lệch</p>
</div>
""", unsafe_allow_html=True)

# Hướng dẫn
with st.expander("📋 Hướng dẫn sử dụng hàng tháng", expanded=False):
    st.markdown("""
    **Mỗi tháng thực hiện theo 5 bước:**

    **Bước 1 – Xuất từ HPT 4 file:**
    - File **Xuất kho theo tháng** (có cột Mã hàng)
    - File **Nhập hàng theo tháng** (có cột Mã hàng)
    - File **Dữ liệu thô XNT** tháng hiện tại (không có mã)
    - File **XNT Thống kê** từ bộ phận thống kê (có mã, có tồn chuẩn)

    **Bước 2 – Upload lên app** theo đúng từng ô bên dưới

    **Bước 3 – Nhấn "Bắt đầu đối chiếu"**

    **Bước 4 – Xem kết quả:**
    - 🟢 Khớp = tồn đúng
    - 🔴 Lệch = cần kiểm tra thực tế
    - 🟡 Nền vàng = ghép tự động gần nhất, cần xác nhận

    **Bước 5 – Tải file Excel kết quả về lưu hồ sơ**

    > **Lưu ý:** Nếu có thuốc mới chưa từng nhập/xuất từ 2025 → tồn xuất hiện trong
    > "XNT có – TK không theo dõi". Bổ sung vào file nhập/xuất tháng trước là tự động
    > có mã tháng sau.
    """)

st.markdown("<hr>", unsafe_allow_html=True)

# ── Tháng / Năm ─────────────────────────────────────────────────────────────
col_t, col_n = st.columns(2)
with col_t: thang = st.selectbox("Tháng báo cáo", range(1,13), index=2, format_func=lambda x: f"Tháng {x}")
with col_n: nam   = st.number_input("Năm", min_value=2024, max_value=2030, value=2026)
thang_nam = f"T{thang}/{nam}"

st.markdown("<hr>", unsafe_allow_html=True)

# ── Upload files ─────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">📂 Upload file dữ liệu</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
<b>File nhập/xuất có Mã hàng</b> (bắt buộc) — dùng để xây bảng map mã.
Có thể upload nhiều tháng để tăng độ phủ mã.
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    f_nhap = st.file_uploader(
        "📥 File Nhập hàng (có Mã HPT)",
        type=["xlsx","xls"],
        accept_multiple_files=True,
        help="Có thể chọn nhiều file nhập: tháng này, tháng trước, cả năm..."
    )
with col2:
    f_xuat = st.file_uploader(
        "📤 File Xuất kho (có Mã HPT)",
        type=["xlsx","xls"],
        accept_multiple_files=True,
        help="Có thể chọn nhiều file xuất: tháng này, tháng trước, cả năm..."
    )

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
col3, col4 = st.columns(2)
with col3:
    f_xnt = st.file_uploader(
        "📊 File XNT thô HPT (không có Mã)",
        type=["xlsx","xls"],
        help="File xuất nhập tồn từ HPT, không có cột mã hàng"
    )
with col4:
    f_tk = st.file_uploader(
        "📋 File XNT Thống kê (có Mã, tồn chuẩn)",
        type=["xlsx","xls"],
        help="File từ bộ phận thống kê, có mã hàng và tồn cuối chuẩn"
    )

# File xác nhận mã (tùy chọn)
st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
f_confirm = st.file_uploader(
    "✅ File xác nhận mã (tùy chọn) — file 'chon_ma_hang_nhieu_ma.xlsx'",
    type=["xlsx"],
    help="Nếu có thuốc nhiều mã và bạn đã chọn mã đúng, upload file xác nhận này"
)

st.markdown("<hr>", unsafe_allow_html=True)

# ── Nút xử lý ────────────────────────────────────────────────────────────────
ready = (f_nhap or f_xuat) and f_xnt and f_tk
if st.button("🔍  Bắt đầu đối chiếu", disabled=not ready):
    with st.spinner("Đang xử lý..."):
        try:
            # Đọc files nhập/xuất
            dfs_nx = []
            for f in (f_nhap or []):
                dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))
            for f in (f_xuat or []):
                dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))

            # Đọc XNT thô
            df_xnt_raw = pd.read_excel(io.BytesIO(f_xnt.read()), sheet_name=0, header=None)

            # Đọc TK
            df_tk_raw = pd.read_excel(io.BytesIO(f_tk.read()), sheet_name=0, header=None)

            # Đọc file xác nhận mã (nếu có)
            confirmed = {}
            if f_confirm:
                from openpyxl import load_workbook
                wb_c = load_workbook(io.BytesIO(f_confirm.read()), data_only=True)
                ws_c = wb_c.active
                cur_ten = ''
                for r in range(2, ws_c.max_row + 1):
                    row = [ws_c.cell(row=r, column=c).value for c in range(1,8)]
                    ten,nd,gia,ma,_,_,chon = row
                    if ma and '→' in str(ma): cur_ten=str(ten or '').strip(); continue
                    if chon and str(chon).strip().lower() in ('yes','y','có','co','x') and ma:
                        confirmed[cur_ten] = str(ma).strip()

            # Chạy đối chiếu
            df_res, err = run_reconciliation(dfs_nx, df_xnt_raw, df_tk_raw, confirmed)
            if err:
                st.error(f"❌ {err}")
                st.stop()

            st.session_state['result']    = df_res
            st.session_state['thang_nam'] = thang_nam
            st.session_state['done']      = True

        except Exception as e:
            st.error(f"❌ Lỗi: {e}")
            st.exception(e)

# ── Hiển thị kết quả ─────────────────────────────────────────────────────────
if st.session_state.get('done'):
    df_res    = st.session_state['result']
    thang_nam = st.session_state['thang_nam']

    df_matched = df_res[df_res['cl'].notna()]
    df_no_tk   = df_res[df_res['method']=='no_tk']
    df_no_xnt  = df_res[df_res['method']=='no_xnt']

    n_khop = (df_matched['cl'].abs()<0.01).sum()
    n_lech = (df_matched['cl'].abs()>=0.01).sum()
    n_noTK = len(df_no_tk)
    n_noXNT= len(df_no_xnt)
    pct    = n_khop/(n_khop+n_lech)*100 if (n_khop+n_lech)>0 else 0

    st.markdown(f"""
    <div class="ok-box">
      <div style="font-size:2rem">✅</div>
      <h3>Đối chiếu hoàn tất – {thang_nam}</h3>
    </div>
    <div class="stat-grid">
      <div class="stat-card">
        <div class="num" style="color:#166534">{n_khop}</div>
        <div class="lbl">✅ Khớp hoàn toàn</div>
      </div>
      <div class="stat-card">
        <div class="num" style="color:#dc2626">{n_lech}</div>
        <div class="lbl">⚠️ Chênh lệch</div>
      </div>
      <div class="stat-card">
        <div class="num" style="color:#2563a8">{n_noTK}</div>
        <div class="lbl">📋 XNT có / TK không</div>
      </div>
      <div class="stat-card">
        <div class="num" style="color:#d97706">{n_noXNT}</div>
        <div class="lbl">📋 TK có / XNT không</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Hiển thị dòng lệch
    df_lech = df_matched[df_matched['cl'].abs()>=0.01].sort_values('cl')
    if len(df_lech):
        st.markdown(f"**⚠️ {n_lech} dòng chênh lệch:**")
        st.dataframe(
            df_lech[['ma','ten','nd','ton_xnt','ton_tk','cl','method']].rename(columns={
                'ma':'Mã','ten':'Tên thuốc','nd':'Nồng độ',
                'ton_xnt':'Tồn XNT','ton_tk':'Tồn TK','cl':'Chênh lệch','method':'Phương pháp'
            }),
            use_container_width=True, hide_index=True,
        )

    if n_noTK > 0:
        st.markdown(f"**📋 {n_noTK} dòng XNT có mà TK không theo dõi:**")
        st.dataframe(
            df_no_tk[['ten','nd','gia','ton_xnt']].rename(columns={
                'ten':'Tên thuốc','nd':'Nồng độ','gia':'Đơn giá','ton_xnt':'Tồn XNT'}),
            use_container_width=True, hide_index=True,
        )

    # Cảnh báo nearest_ton
    df_warn = df_lech[df_lech['method']=='nearest_ton']
    if len(df_warn):
        st.markdown(f"""
        <div class="warn-box">
        ⚠️ <b>{len(df_warn)} dòng ghép theo "tồn gần nhất"</b> — đây là trường hợp cùng tên+nồng độ+giá
        có nhiều mã. App tự ghép theo tồn gần nhất. Hãy kiểm tra lại thực tế.
        </div>""", unsafe_allow_html=True)

    # Tải file
    st.markdown("<hr>", unsafe_allow_html=True)
    excel_bytes = export_excel(df_res, thang_nam)
    st.download_button(
        label=f"⬇️  Tải Kết Quả Đối Chiếu XNT {thang_nam} (.xlsx)",
        data=excel_bytes,
        file_name=f"doi_chieu_XNT_{thang_nam.replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
