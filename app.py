"""
ĐỐI CHIẾU DƯỢC – Bệnh viện Đà Nẵng
v5 – Sửa used_tk index, tách active/inactive TK, khớp đa mã chính xác
"""

import io, re, warnings
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Đối Chiếu Dược – BV Đà Nẵng", page_icon="🏥", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;600;700&display=swap');
html,body,[class*="css"]{font-family:'Be Vietnam Pro',sans-serif;}
.hero{background:linear-gradient(135deg,#1a3a5c 0%,#2563a8 60%,#1e7fcb 100%);
  border-radius:16px;padding:28px 36px 22px;margin-bottom:20px;color:white;
  box-shadow:0 8px 32px rgba(37,99,168,.25);}
.hero h1{font-size:1.45rem;font-weight:700;margin:0 0 6px;}
.hero .sub{font-size:.86rem;font-weight:300;opacity:.85;margin:0;}
.hero .badge{display:inline-block;background:rgba(255,255,255,.18);border-radius:20px;
  padding:3px 12px;font-size:.73rem;font-weight:600;letter-spacing:1px;margin-bottom:10px;text-transform:uppercase;}
.upload-section{background:#f8faff;border:1.5px solid #c7d9f5;border-radius:14px;
  padding:18px 20px;margin-bottom:18px;}
.upload-section h4{color:#1a3a5c;font-size:.82rem;font-weight:700;letter-spacing:1.2px;
  text-transform:uppercase;margin:0 0 14px;}
.map-box{background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;
  padding:10px 16px;margin:10px 0;font-size:.85rem;color:#166534;}
.warn-box{background:#fff7ed;border-left:4px solid #f59e0b;border-radius:0 10px 10px 0;
  padding:11px 15px;margin:10px 0;font-size:.84rem;color:#92400e;line-height:1.6;}
.info-box{background:#eff6ff;border-left:4px solid #2563a8;border-radius:0 10px 10px 0;
  padding:11px 15px;margin:10px 0 16px;font-size:.85rem;color:#1e3a5f;line-height:1.7;}
.stat-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0;}
.stat-card{background:white;border:1px solid #e2e8f0;border-radius:12px;
  padding:14px 10px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.06);}
.stat-card .num{font-size:1.5rem;font-weight:700;line-height:1;}
.stat-card .lbl{font-size:.7rem;color:#64748b;margin-top:4px;}
.ok-box{background:#f0fdf4;border:1.5px solid #86efac;border-radius:12px;
  padding:14px 18px;margin:14px 0;text-align:center;}
.ok-box h3{color:#166534;margin:4px 0;font-size:.97rem;}
.stButton>button{background:linear-gradient(135deg,#1a3a5c,#2563a8)!important;
  color:white!important;font-weight:600!important;font-size:.93rem!important;
  border:none!important;border-radius:10px!important;padding:12px 0!important;
  width:100%!important;box-shadow:0 4px 14px rgba(37,99,168,.3)!important;}
[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#166534,#16a34a)!important;
  color:white!important;font-weight:700!important;font-size:.97rem!important;
  border:none!important;border-radius:10px!important;padding:14px 0!important;width:100%!important;}
[data-testid="stFileUploader"]{border:2px dashed #2563a8!important;border-radius:10px!important;background:#f0f6ff!important;}
hr{border:none;border-top:1px solid #e2e8f0;margin:18px 0;}
#MainMenu,footer{visibility:hidden;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  HÀM TIỆN ÍCH
# ══════════════════════════════════════════════════════════════════════════════

def norm(s):
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = re.sub(r'_x000a_', ' ', s); s = re.sub(r'\n', ' ', s)
    s = re.sub(r'[#]', '', s)
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    s = re.sub(r'(\d)\s+(mg|ml|mcg|μg|g|iu|%|meq|l)', r'\1\2', s)
    s = re.sub(r'\s*\+\s*', '+', s)
    return re.sub(r'\s+', ' ', s).strip()


def is_drug(s):
    """Lọc tên thuốc thực sự – bỏ số thứ tự cột (3,4,5...)."""
    if not isinstance(s, str): return False
    s = s.strip()
    if len(s) < 2: return False
    try: float(s); return False
    except: return True


def safe_float(v):
    try: return float(v)
    except: return 0.0


def parse_tk(df_raw):
    """Parse file XNT Thống kê. Dùng reset_index để index là 0,1,2,... liên tục."""
    rows = []
    for _, row in df_raw.iloc[5:].iterrows():
        ma  = str(row[4]).strip() if not pd.isna(row[4]) else ''
        ten = str(row[5]).strip() if not pd.isna(row[5]) else ''
        if not ma or not ten: continue
        gia = safe_float(row[11])
        rows.append({
            'ma': ma, 'ten_tk': ten,
            'nd_tk':   str(row[8]).strip() if not pd.isna(row[8]) else '',
            'gia_tk':  gia,
            'nhap_tk': safe_float(row[14]),
            'ton_tk':  safe_float(row[24]),
            'kten': norm(ten),
            'knd':  norm(str(row[8]) if not pd.isna(row[8]) else ''),
            'kgia': int(round(gia)) if gia else 0,
        })
    return pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()


def extract_ma_map(dfs_nhap_xuat):
    """Bảng mã HPT từ file nhập/xuất kho (col0=STT, col1=Mã, col2=Tên, col3=ND, col5=Giá)."""
    all_rows = []
    for df in dfs_nhap_xuat:
        if df is None or df.empty: continue
        for _, row in df.iterrows():
            try: int(str(row[0]).strip())
            except: continue
            if pd.isna(row[1]) or pd.isna(row[2]): continue
            ten = str(row[2]).strip()
            if not is_drug(ten): continue
            gia = safe_float(row[5])
            all_rows.append({
                'ma': str(row[1]).strip(), 'ten': ten,
                'nd': str(row[3]).strip() if not pd.isna(row[3]) else '',
                'gia': gia,
                'kten': norm(ten),
                'knd':  norm(str(row[3]) if not pd.isna(row[3]) else ''),
                'kgia': int(round(gia)) if gia else 0,
            })
    if not all_rows: return pd.DataFrame()
    return (pd.DataFrame(all_rows)
              .drop_duplicates(subset=['ma', 'kten', 'knd', 'kgia'])
              .reset_index(drop=True))


def find_ma(kten, knd, kgia, global_map):
    if global_map is None or global_map.empty: return ''
    mask = (global_map['kten']==kten)&(global_map['knd']==knd)&(global_map['kgia']==kgia)
    matched = global_map[mask]['ma'].unique()
    return ', '.join(matched) if len(matched) > 0 else ''


def parse_raw_lines_bbkn(df_raw, sl_col=9):
    """
    Parse BBKN: col0=STT, col1=Mã HĐ, col2=Tên, col3=ND, col8=Giá, col9=SL
    Giữ TỪNG dòng – KHÔNG gộp. Thuật toán match sẽ quyết định khi nào gộp.
    """
    rows = []
    for _, row in df_raw.iterrows():
        try: int(str(row[0]).strip())
        except: continue
        if pd.isna(row[2]): continue
        ten = str(row[2]).strip()
        if not is_drug(ten): continue
        nd    = str(row[3]).strip() if not pd.isna(row[3]) else ''
        gia   = safe_float(row[8])
        sl    = safe_float(row[sl_col])
        ma_hd = str(row[1]).strip() if not pd.isna(row[1]) else ''
        rows.append({'ten': ten, 'nd': nd, 'gia': gia, 'sl': sl, 'ma_hd': ma_hd,
                     'kten': norm(ten), 'knd': norm(nd),
                     'kgia': int(round(gia)) if gia else 0})
    return pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()


def parse_raw_lines_bbkk(df_raw, sl_col=8):
    """
    Parse BBKK: col0=STT, col1=Tên, col2=ND, col4=Giá, col8=SL thực tế
    Giữ TỪNG dòng – KHÔNG gộp.
    """
    rows = []
    for _, row in df_raw.iterrows():
        try: int(str(row[0]).strip())
        except: continue
        if pd.isna(row[1]): continue
        ten = str(row[1]).strip()
        if not is_drug(ten): continue
        nd  = str(row[2]).strip() if not pd.isna(row[2]) else ''
        gia = safe_float(row[4])
        sl  = safe_float(row[sl_col])
        rows.append({'ten': ten, 'nd': nd, 'gia': gia, 'sl': sl, 'ma_hd': '',
                     'kten': norm(ten), 'knd': norm(nd),
                     'kgia': int(round(gia)) if gia else 0})
    return pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
#  THUẬT TOÁN KHỚP THỐNG NHẤT (dùng cho cả KN và KK)
# ══════════════════════════════════════════════════════════════════════════════

def match_hpt_to_tk(df_hpt, df_tk, sl_col_tk):
    """
    Thuật toán khớp đa mã, xử lý đúng mọi trường hợp:

    Với mỗi nhóm HPT cùng kten+knd+kgia:
      - Lọc TK available theo key, tách thành active (sl > 0) và inactive (sl = 0)
      - Inactive TK: đánh dấu used nhưng không tạo kết quả (không báo lỗi giả)
      - 0 active TK  → hpt_no_tk
      - 1 active TK  → gộp tổng HPT rồi so với TK đó (nhiều HĐ → 1 mã)
      - N active TK  → khớp 1-1 từng dòng HPT với TK gần nhất theo SL
                       (Xatral 6000↔6000, 18000↔18000)
    """
    results = []
    used_tk_idx = set()   # lưu df_tk.index thực sự (integer)

    def is_active(tk_val):
        return (sl_col_tk == 'nhap_tk' and tk_val > 0) or \
               (sl_col_tk == 'ton_tk'  and abs(tk_val) >= 0.01)

    for (kten, knd, kgia), grp_hpt in df_hpt.groupby(['kten', 'knd', 'kgia'], sort=False):
        mask     = (df_tk['kten']==kten)&(df_tk['knd']==knd)&(df_tk['kgia']==kgia)
        avail_tk = df_tk[mask & ~df_tk.index.isin(used_tk_idx)]
        hpt_list = grp_hpt.reset_index(drop=True)
        sl_sum   = hpt_list['sl'].sum()
        hoa_don  = ', '.join([v for v in hpt_list['ma_hd'] if v])
        ten0, nd0, gia0 = hpt_list.iloc[0]['ten'], hpt_list.iloc[0]['nd'], hpt_list.iloc[0]['gia']

        def row_hpt_no_tk():
            return {'ma': '', 'ten_hpt': ten0, 'nd': nd0, 'gia': gia0,
                    'sl_hpt': sl_sum, 'hoa_don': hoa_don,
                    'ten_tk': '', 'sl_tk': None, 'cl': None, 'status': 'hpt_no_tk'}

        def row_matched(sl_hpt, hd, tr):
            tk_val = tr[sl_col_tk]
            return {'ma': tr['ma'], 'ten_hpt': ten0, 'nd': nd0, 'gia': gia0,
                    'sl_hpt': sl_hpt, 'hoa_don': hd,
                    'ten_tk': tr['ten_tk'], 'sl_tk': tk_val,
                    'cl': sl_hpt - tk_val, 'status': 'matched'}

        def row_tk_no_hpt(tr):
            return {'ma': tr['ma'], 'ten_hpt': '', 'nd': tr['nd_tk'], 'gia': tr['gia_tk'],
                    'sl_hpt': None, 'hoa_don': '',
                    'ten_tk': tr['ten_tk'], 'sl_tk': tr[sl_col_tk],
                    'cl': None, 'status': 'tk_no_hpt'}

        if len(avail_tk) == 0:
            results.append(row_hpt_no_tk()); continue

        # Tách active / inactive
        active_idx   = [i for i in avail_tk.index if is_active(avail_tk.loc[i, sl_col_tk])]
        inactive_idx = [i for i in avail_tk.index if not is_active(avail_tk.loc[i, sl_col_tk])]

        # Inactive: đánh dấu used, không báo lỗi
        for i in inactive_idx:
            used_tk_idx.add(i)

        if len(active_idx) == 0:
            results.append(row_hpt_no_tk()); continue

        if len(active_idx) == 1:
            # 1 active TK → gộp tổng HPT
            ti = active_idx[0]
            results.append(row_matched(sl_sum, hoa_don, df_tk.loc[ti]))
            used_tk_idx.add(ti); continue

        # N active TK → khớp 1-1 từng dòng HPT với TK gần nhất
        tk_pool = list(active_idx)
        matched_tk = set(); matched_hpt = set()

        # Pass 1: exact match
        for hi, hr in hpt_list.iterrows():
            for ti in tk_pool:
                if ti in matched_tk: continue
                if abs(hr['sl'] - df_tk.loc[ti, sl_col_tk]) < 0.01:
                    results.append(row_matched(hr['sl'], hr['ma_hd'], df_tk.loc[ti]))
                    matched_tk.add(ti); matched_hpt.add(hi); used_tk_idx.add(ti); break

        # Pass 2: nearest for remaining
        for hi, hr in hpt_list.iterrows():
            if hi in matched_hpt: continue
            best_ti, best_d = None, float('inf')
            for ti in tk_pool:
                if ti in matched_tk: continue
                d = abs(hr['sl'] - df_tk.loc[ti, sl_col_tk])
                if d < best_d: best_d, best_ti = d, ti
            if best_ti is not None:
                results.append(row_matched(hr['sl'], hr['ma_hd'], df_tk.loc[best_ti]))
                matched_tk.add(best_ti); used_tk_idx.add(best_ti)
            else:
                results.append({'ma': '', 'ten_hpt': hr['ten'], 'nd': hr['nd'], 'gia': hr['gia'],
                                 'sl_hpt': hr['sl'], 'hoa_don': hr['ma_hd'],
                                 'ten_tk': '', 'sl_tk': None, 'cl': None, 'status': 'hpt_no_tk'})

        # Pass 3: unmatched active TK → tk_no_hpt
        for ti in tk_pool:
            if ti not in matched_tk:
                if is_active(df_tk.loc[ti, sl_col_tk]):
                    results.append(row_tk_no_hpt(df_tk.loc[ti]))
                used_tk_idx.add(ti)

    # TK còn lại (không có HPT nào khớp)
    for ti, tr in df_tk.iterrows():
        if ti not in used_tk_idx and is_active(tr[sl_col_tk]):
            results.append(row_tk_no_hpt(tr))

    return pd.DataFrame(results)


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE 1 – ĐỐI CHIẾU XNT
# ══════════════════════════════════════════════════════════════════════════════

def run_xnt(dfs_nx, df_xnt_raw, df_tk_raw, confirmed_override):
    global_map = extract_ma_map(dfs_nx)
    if global_map.empty:
        return None, "Không đọc được dữ liệu từ file nhập/xuất"

    aug = pd.DataFrame([{'ma': '0005301225', 'ten': 'Augmentin 1g', 'nd': '875mg + 125mg',
                          'gia': 16680, 'kten': norm('Augmentin 1g'),
                          'knd': norm('875mg + 125mg'), 'kgia': 16680}])
    global_map = pd.concat([global_map, aug], ignore_index=True).drop_duplicates(
        subset=['ma', 'kten', 'knd', 'kgia'])

    for i, row in global_map.iterrows():
        if row['ten'] in confirmed_override:
            global_map.at[i, 'ma'] = confirmed_override[row['ten']]

    xnt_rows = []
    for _, row in df_xnt_raw.iterrows():
        try: stt = int(str(row[0]).strip())
        except: continue
        if pd.isna(row[2]) or not isinstance(row[2], str): continue
        if row[2].strip().isdigit(): continue
        xnt_rows.append({
            'stt': stt, 'ten': str(row[2]).strip(),
            'nd':  str(row[3]).strip() if not pd.isna(row[3]) else '',
            'gia': row[8] if not pd.isna(row[8]) else 0,
            'ton_xnt': safe_float(row[12]),
            'kten': norm(str(row[2])),
            'knd':  norm(str(row[3]) if not pd.isna(row[3]) else ''),
            'kgia': int(round(float(row[8]))) if not pd.isna(row[8]) else 0,
        })
    df_xnt = pd.DataFrame(xnt_rows)
    df_tk  = parse_tk(df_tk_raw)

    results = []; used_tk = set()

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

    for (kten, knd, kgia), grp_x in df_xnt.groupby(['kten', 'knd', 'kgia'], sort=False):
        mask  = (df_tk['kten']==kten)&(df_tk['knd']==knd)&(df_tk['kgia']==kgia)
        grp_t = df_tk[mask & ~df_tk.index.isin(used_tk)].copy()

        if len(grp_x) == 1 and len(grp_t) == 1:
            xr = grp_x.iloc[0]; tr = grp_t.iloc[0]
            results.append(make_row(xr, tr, '1-1'))
            used_tk.add(grp_t.index[0])
        elif len(grp_x) >= 1 and len(grp_t) > 0:
            xl = grp_x.reset_index(drop=True); tl = grp_t.reset_index(drop=True)
            mx, mt = set(), set()
            for xi, xr in xl.iterrows():
                for ti, tr in tl.iterrows():
                    if ti in mt: continue
                    if abs(xr['ton_xnt'] - tr['ton_tk']) < 0.01:
                        results.append(make_row(xr, tr, 'exact_ton'))
                        used_tk.add(grp_t.index[ti]); mx.add(xi); mt.add(ti); break
            for xi, xr in xl.iterrows():
                if xi in mx: continue
                best_d, best_ti, best_tr = float('inf'), None, None
                for ti, tr in tl.iterrows():
                    if ti in mt: continue
                    d = abs(xr['ton_xnt'] - tr['ton_tk'])
                    if d < best_d: best_d, best_ti, best_tr = d, ti, tr
                if best_tr is not None:
                    results.append(make_row(xr, best_tr, 'nearest_ton'))
                    used_tk.add(grp_t.index[best_ti]); mt.add(best_ti)
                else:
                    results.append(make_row(xr, None, 'no_tk'))
        else:
            for _, xr in grp_x.iterrows():
                if xr['ton_xnt'] != 0:
                    results.append(make_row(xr, None, 'no_tk'))

    for idx, tr in df_tk.iterrows():
        if idx not in used_tk and abs(tr['ton_tk']) >= 0.01:
            results.append({
                'ma': tr['ma'], 'ten': '', 'nd': '', 'gia': '',
                'ton_xnt': None, 'ten_tk': tr['ten_tk'],
                'nd_tk': tr['nd_tk'], 'ton_tk': tr['ton_tk'],
                'method': 'no_xnt', 'cl': None,
            })

    return pd.DataFrame(results), None


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE 2 – ĐỐI CHIẾU KIỂM NHẬP
# ══════════════════════════════════════════════════════════════════════════════

def run_kn(df_bbkn_raw, df_tk_raw, global_map, sl_col=9):
    df_kn = parse_raw_lines_bbkn(df_bbkn_raw, sl_col)
    if df_kn.empty:
        return None, "Không đọc được dữ liệu từ file BBKN"
    df_tk = parse_tk(df_tk_raw)
    if df_tk.empty:
        return None, "Không đọc được dữ liệu từ file Thống kê"
    df_r = match_hpt_to_tk(df_kn, df_tk, 'nhap_tk')
    # Đổi tên cột cho giao diện / Excel
    df_r = df_r.rename(columns={'sl_hpt': 'nhap_hpt', 'sl_tk': 'nhap_tk'})
    return df_r, None


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE 3 – ĐỐI CHIẾU KIỂM KÊ
# ══════════════════════════════════════════════════════════════════════════════

def run_kk(df_bbkk_raw, df_tk_raw, global_map, sl_col=8):
    df_kk = parse_raw_lines_bbkk(df_bbkk_raw, sl_col)
    if df_kk.empty:
        return None, "Không đọc được dữ liệu từ file Biên bản kiểm kê"
    df_tk = parse_tk(df_tk_raw)
    if df_tk.empty:
        return None, "Không đọc được dữ liệu từ file Thống kê"
    df_r = match_hpt_to_tk(df_kk, df_tk, 'ton_tk')
    # Đổi tên cột cho giao diện / Excel
    df_r = df_r.rename(columns={'sl_hpt': 'sl_kk', 'sl_tk': 'ton_tk'})
    return df_r, None


# ══════════════════════════════════════════════════════════════════════════════
#  XUẤT EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def _hdr(ws, headers, fill_hex='1F3864'):
    TH = Side(style='thin'); MH = Side(style='medium')
    FH = PatternFill('solid', fgColor=fill_hex)
    for ci, (h, w) in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = Font(name='Times New Roman', bold=True, size=11, color='FFFFFF')
        c.fill = FH
        c.border = Border(left=TH, right=TH, top=MH, bottom=MH)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 36


def _row(ws, ri, vals, fill, right_cols=(), center_cols=()):
    TH = Side(style='thin')
    for ci, v in enumerate(vals, 1):
        safe = '' if (v is None or (isinstance(v, float) and pd.isna(v))) else v
        c = ws.cell(row=ri, column=ci, value=safe)
        c.font = Font(name='Times New Roman', size=11)
        c.fill = fill
        c.border = Border(left=TH, right=TH, top=TH, bottom=TH)
        if ci in right_cols:
            c.alignment = Alignment(vertical='center', horizontal='right')
            if ci in right_cols and isinstance(safe, (int, float)):
                c.number_format = '#,##0.##'
        elif ci in center_cols:
            c.alignment = Alignment(vertical='center', horizontal='center')
        else:
            c.alignment = Alignment(vertical='center', horizontal='left')
    ws.row_dimensions[ri].height = 18


FOK = PatternFill('solid', fgColor='E2EFDA')
FLE = PatternFill('solid', fgColor='FFD7D7')
FWA = PatternFill('solid', fgColor='FFF9C4')
FBL = PatternFill('solid', fgColor='DEEBF7')
FOR = PatternFill('solid', fgColor='FCE4D6')


def build_xnt_sheets(wb, df_res, tn):
    ws = wb.create_sheet(f"DC XNT {tn.replace('/','_')}")
    _hdr(ws, [('Mã HPT',16),('Tên thuốc (HPT)',32),('Tên thuốc (TK)',28),
              ('Nồng độ',22),('Đơn giá',12),('Tồn cuối HPT',13),
              ('Tồn cuối TK',13),('Chênh lệch',13),('Trạng thái',25),('Phương pháp',18)])
    mm = {'1-1':'Chính xác','exact_ton':'Khớp tồn','nearest_ton':'Gần nhất ⚠️'}
    df_m = df_res[df_res['cl'].notna()]
    df_l = df_m[df_m['cl'].abs()>=0.01].sort_values('cl')
    df_k = df_m[df_m['cl'].abs()<0.01]
    for ri,(_, r) in enumerate(pd.concat([df_l,df_k],ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl)<0.01:    fill,st = FOK,'✅ Khớp'
        elif cl>0:           fill,st = FLE,f'⬆️ HPT cao hơn {cl:+.0f}'
        else:                fill,st = FLE,f'⬇️ HPT thấp hơn {cl:+.0f}'
        if r['method']=='nearest_ton' and abs(cl)>=0.01: fill=FWA
        _row(ws,ri,[r['ma'],r['ten'],r.get('ten_tk',''),r['nd'],r['gia'],
                    r['ton_xnt'],r['ton_tk'],cl,st,mm.get(r['method'],'')],
             fill,right_cols=(5,6,7,8),center_cols=(1,9,10))
    ws.freeze_panes='A2'

    # HPT có – TK không (thêm cột mã)
    ws2 = wb.create_sheet("XNT – HPT có, TK không")
    _hdr(ws2,[('Mã HPT',16),('Tên thuốc HPT',35),('Nồng độ',25),('Đơn giá',12),('Tồn HPT',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['method']=='no_tk'].iterrows(), 2):
        _row(ws2,ri,[r['ma'],r['ten'],r['nd'],r['gia'],r['ton_xnt'],'HPT có nhưng TK không theo dõi'],
             FBL,right_cols=(3,4,5))

    # TK có – HPT không
    ws3 = wb.create_sheet("XNT – TK có, HPT không")
    _hdr(ws3,[('Mã HPT',16),('Tên thuốc TK',35),('Nồng độ TK',25),('Tồn TK',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['method']=='no_xnt'].iterrows(), 2):
        _row(ws3,ri,[r['ma'],r['ten_tk'],r['nd_tk'],r['ton_tk'],'TK có nhưng HPT không phát sinh'],
             FOR,right_cols=(4,))


def build_kn_sheets(wb, df_res, tn):
    ws = wb.create_sheet(f"DC Kiểm nhập {tn.replace('/','_')}")
    _hdr(ws,[('Mã HPT',16),('Tên thuốc (HPT)',32),('Tên thuốc (TK)',28),
             ('Nồng độ',22),('Đơn giá',12),('Nhập HPT (tổng)',13),
             ('Số HĐ',8),('Danh sách HĐ',22),('Nhập TK',12),('Chênh lệch',13),('Trạng thái',28)])
    df_m = df_res[df_res['status']=='matched']
    df_l = df_m[df_m['cl'].abs()>=0.01].sort_values('cl')
    df_k = df_m[df_m['cl'].abs()<0.01]
    for ri,(_, r) in enumerate(pd.concat([df_l,df_k],ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl)<0.01:    fill,st = FOK,'✅ Khớp'
        elif cl>0:           fill,st = FLE,f'⬆️ HPT cao hơn {cl:+.0f}'
        else:                fill,st = FLE,f'⬇️ HPT thấp hơn {cl:+.0f}'
        _row(ws,ri,[r['ma'],r['ten_hpt'],r['ten_tk'],r['nd'],r['gia'],
                    r['nhap_hpt'],r.get('n_hoadon',''),r.get('hoa_don',''),
                    r['nhap_tk'],cl,st],
             fill,right_cols=(5,6,8,9,10),center_cols=(1,11))
    ws.freeze_panes='A2'

    ws2 = wb.create_sheet("KN – HPT có, TK không")
    _hdr(ws2,[('Mã HPT',16),('Tên thuốc HPT',35),('Nồng độ',25),('Đơn giá',12),
              ('Nhập HPT (tổng)',13),('Danh sách HĐ',22),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='hpt_no_tk'].iterrows(), 2):
        _row(ws2,ri,[r['ma'],r['ten_hpt'],r['nd'],r['gia'],r['nhap_hpt'],
                     r.get('hoa_don',''),'HPT có nhưng TK không có số nhập'],
             FBL,right_cols=(3,4,5))

    ws3 = wb.create_sheet("KN – TK có, HPT không")
    _hdr(ws3,[('Mã HPT',16),('Tên thuốc TK',35),('Nồng độ TK',25),('Nhập TK',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='tk_no_hpt'].iterrows(), 2):
        _row(ws3,ri,[r['ma'],r['ten_tk'],r['nd'],r['nhap_tk'],'TK có nhập nhưng HPT không có'],
             FOR,right_cols=(4,))


def build_kk_sheets(wb, df_res, tn):
    ws = wb.create_sheet(f"DC Kiểm kê {tn.replace('/','_')}")
    _hdr(ws,[('Mã HPT',16),('Tên thuốc (HPT)',32),('Tên thuốc (TK)',28),
             ('Nồng độ',22),('Đơn giá',12),('SL Kiểm kê (tổng)',14),
             ('Tồn cuối TK',13),('Chênh lệch',13),('Trạng thái',28)])
    df_m = df_res[df_res['status']=='matched']
    df_l = df_m[df_m['cl'].abs()>=0.01].sort_values('cl')
    df_k = df_m[df_m['cl'].abs()<0.01]
    for ri,(_, r) in enumerate(pd.concat([df_l,df_k],ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl)<0.01:    fill,st = FOK,'✅ Khớp'
        elif cl>0:           fill,st = FLE,f'⬆️ HPT cao hơn {cl:+.0f}'
        else:                fill,st = FLE,f'⬇️ HPT thấp hơn {cl:+.0f}'
        _row(ws,ri,[r['ma'],r['ten_hpt'],r['ten_tk'],r['nd'],r['gia'],
                    r['sl_kk'],r['ton_tk'],cl,st],
             fill,right_cols=(5,6,7,8),center_cols=(1,9))
    ws.freeze_panes='A2'

    ws2 = wb.create_sheet("KK – HPT có, TK không")
    _hdr(ws2,[('Mã HPT',16),('Tên thuốc HPT',35),('Nồng độ',25),('Đơn giá',12),
              ('SL Kiểm kê',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='hpt_no_tk'].iterrows(), 2):
        _row(ws2,ri,[r['ma'],r['ten_hpt'],r['nd'],r['gia'],r['sl_kk'],'HPT có nhưng TK không theo dõi tồn'],
             FBL,right_cols=(3,4,5))

    ws3 = wb.create_sheet("KK – TK có, HPT không")
    _hdr(ws3,[('Mã HPT',16),('Tên thuốc TK',35),('Nồng độ TK',25),('Tồn TK',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='tk_no_hpt'].iterrows(), 2):
        _row(ws3,ri,[r['ma'],r['ten_tk'],r['nd'],r['ton_tk'],'TK có tồn nhưng không có trong kiểm kê'],
             FOR,right_cols=(4,))


def build_summary(wb, results_map, tn):
    ws = wb.create_sheet("📊 Tóm tắt", 0)
    ws.column_dimensions['A'].width = 40; ws.column_dimensions['B'].width = 18
    def wr(ri,k,v,bold=False):
        c1=ws.cell(row=ri,column=1,value=k); c2=ws.cell(row=ri,column=2,value=v)
        c1.font=Font(name='Times New Roman',bold=bold,size=12 if bold else 11)
        c2.font=Font(name='Times New Roman',size=11)
    ri=1; wr(ri,f'BÁO CÁO ĐỐI CHIẾU DƯỢC – {tn}','',bold=True); ri+=2
    for label,df_r,col_cl,col_st,no_vals,no_labels in results_map:
        wr(ri,f'── {label} ──','',bold=True); ri+=1
        if df_r is not None and col_cl:
            m=df_r[df_r[col_cl].notna()]
            nk=(m[col_cl].abs()<0.01).sum(); nl=(m[col_cl].abs()>=0.01).sum()
            pct=nk/(nk+nl)*100 if (nk+nl)>0 else 0
            wr(ri,'  ✅ Khớp hoàn toàn',f'{nk} dòng'); ri+=1
            wr(ri,'  ⚠️  Chênh lệch',f'{nl} dòng'); ri+=1
            wr(ri,'  📊 Tỷ lệ khớp',f'{pct:.1f}%'); ri+=1
        if df_r is not None and col_st:
            for nv,nl2 in zip(no_vals,no_labels):
                cnt=(df_r[col_st]==nv).sum()
                wr(ri,f'  📋 {nl2}',f'{cnt} dòng'); ri+=1
        ri+=1


def export_excel(res_xnt, res_kn, res_kk, tn):
    wb=Workbook(); wb.remove(wb.active)
    if res_xnt is not None: build_xnt_sheets(wb, res_xnt, tn)
    if res_kn  is not None: build_kn_sheets(wb, res_kn, tn)
    if res_kk  is not None: build_kk_sheets(wb, res_kk, tn)
    rm=[]
    if res_xnt is not None:
        rm.append(('Đối chiếu XNT',res_xnt,'cl','method',
                   ['no_tk','no_xnt'],['HPT có – TK không','TK có – HPT không']))
    if res_kn is not None:
        rm.append(('Đối chiếu Kiểm nhập',res_kn,'cl','status',
                   ['hpt_no_tk','tk_no_hpt'],['HPT có – TK không','TK có – HPT không']))
    if res_kk is not None:
        rm.append(('Đối chiếu Kiểm kê',res_kk,'cl','status',
                   ['hpt_no_tk','tk_no_hpt'],['HPT có – TK không','TK có – HPT không']))
    build_summary(wb, rm, tn)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  GIAO DIỆN – Upload 1 lần, dùng chung
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="hero">
  <div class="badge">🏥 Bệnh viện Đà Nẵng · Khoa Dược</div>
  <h1>ĐỐI CHIẾU DƯỢC – HPT vs THỐNG KÊ</h1>
  <p class="sub">XNT · Kiểm nhập · Kiểm kê · Upload 1 lần · Tách đa mã thầu</p>
</div>
""", unsafe_allow_html=True)

# ── Tháng / Năm ──────────────────────────────────────────────────────────────
ca, cb = st.columns(2)
with ca: thang = st.selectbox("Tháng báo cáo", range(1,13), index=2, format_func=lambda x: f"Tháng {x}")
with cb: nam   = st.number_input("Năm", min_value=2024, max_value=2030, value=2026)
thang_nam = f"T{thang}/{nam}"

st.markdown("<hr>", unsafe_allow_html=True)

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  ZONE UPLOAD DUY NHẤT – dùng chung cho tất cả 3 module                 ║
# ╚══════════════════════════════════════════════════════════════════════════╝
st.markdown("""
<div class="upload-section">
<h4>📂 Khu vực Upload File – Tải lên 1 lần, dùng cho cả 3 module</h4>
</div>
""", unsafe_allow_html=True)

with st.container():
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        f_nhap = st.file_uploader(
            "📥 Báo cáo Nhập hàng trong tháng (HPT – có Mã)",
            type=["xlsx","xls"], accept_multiple_files=True, key="u_nhap",
            help="Dùng để lấy mã HPT. Áp dụng cho cả XNT, Kiểm nhập, Kiểm kê.")
    with col_u2:
        f_xuat = st.file_uploader(
            "📤 Báo cáo Xuất kho trong tháng (HPT – có Mã)",
            type=["xlsx","xls"], accept_multiple_files=True, key="u_xuat",
            help="Dùng để lấy mã HPT. Áp dụng cho cả XNT, Kiểm nhập, Kiểm kê.")

    col_u3, col_u4 = st.columns(2)
    with col_u3:
        f_tk = st.file_uploader(
            "📋 File XNT Thống kê – số chuẩn (dùng chung)",
            type=["xlsx","xls"], key="u_tk",
            help="File số liệu chuẩn của TK. Dùng chung cho XNT, Kiểm nhập và Kiểm kê.")
    with col_u4:
        f_xnt_tho = st.file_uploader(
            "📊 File XNT thô HPT (dành riêng cho Tab Đối chiếu XNT)",
            type=["xlsx","xls"], key="u_xnt_tho",
            help="Chỉ cần cho Tab Đối chiếu XNT.")

    col_u5, col_u6 = st.columns(2)
    with col_u5:
        f_bbkn = st.file_uploader(
            "📄 Biên bản Kiểm nhập – BBKN (dành riêng cho Tab Kiểm nhập)",
            type=["xlsx","xls"], key="u_bbkn")
    with col_u6:
        f_bbkk = st.file_uploader(
            "📄 Biên bản Kiểm kê – BBKK (dành riêng cho Tab Kiểm kê)",
            type=["xlsx","xls"], key="u_bbkk")

# Hiển thị trạng thái bảng mã
gmap = st.session_state.get('global_map')
if gmap is not None and not gmap.empty:
    n_ma = gmap['ma'].nunique()
    n_multi = gmap.groupby(['kten','knd','kgia']).filter(lambda x: len(x)>1)['kten'].nunique()
    st.markdown(f"""
    <div class="map-box">✅ <b>Bảng mã hàng sẵn sàng:</b> {n_ma} mã HPT
    {f'— <b>{n_multi} thuốc</b> có ≥2 mã (thầu cũ/mới)' if n_multi>0 else ''}
    </div>""", unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ── Chọn cột SL cho BBKN / BBKK ──────────────────────────────────────────────
sl_col_kn = 9; sl_col_kk = 8
with st.expander("⚙️ Tùy chọn cột số lượng (nếu cấu trúc file khác mặc định)", expanded=False):
    cc1, cc2 = st.columns(2)
    with cc1:
        if f_bbkn:
            try:
                f_bbkn.seek(0)
                pv = pd.read_excel(io.BytesIO(f_bbkn.read()), sheet_name=0, header=None, nrows=12)
                f_bbkn.seek(0)
                opts = {f"Cột {i} | {str(pv.iloc[9,i] if len(pv)>9 else '')[:28]}": i
                        for i in range(len(pv.columns))}
                def_kn = next((k for k,v in opts.items() if v==9), list(opts.keys())[0])
                sl_col_kn = opts[st.selectbox("Cột SL nhập trong BBKN:", list(opts.keys()),
                    index=list(opts.keys()).index(def_kn), key="sel_kn")]
            except: pass
    with cc2:
        if f_bbkk:
            try:
                f_bbkk.seek(0)
                pv2 = pd.read_excel(io.BytesIO(f_bbkk.read()), sheet_name=0, header=None, nrows=12)
                f_bbkk.seek(0)
                opts2 = {f"Cột {i} | {str(pv2.iloc[9,i] if len(pv2)>9 else '')[:28]}": i
                         for i in range(len(pv2.columns))}
                def_kk = next((k for k,v in opts2.items() if v==8), list(opts2.keys())[0])
                sl_col_kk = opts2[st.selectbox("Cột SL thực tế trong BBKK:", list(opts2.keys()),
                    index=list(opts2.keys()).index(def_kk), key="sel_kk")]
            except: pass

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  3 TABS KẾT QUẢ                                                         ║
# ╚══════════════════════════════════════════════════════════════════════════╝
tab_xnt, tab_kn, tab_kk = st.tabs(["📊 Đối chiếu XNT", "📥 Đối chiếu Kiểm nhập", "🔍 Đối chiếu Kiểm kê"])

# ─── TAB 1: XNT ───────────────────────────────────────────────────────────────
with tab_xnt:
    st.markdown("""<div class="info-box">
    Upload file <b>XNT thô HPT</b> + file <b>XNT Thống kê</b> + file <b>Nhập/Xuất kho</b> ở trên.
    Nhấn nút bên dưới để chạy đối chiếu tồn cuối.
    </div>""", unsafe_allow_html=True)

    ready_xnt = bool((f_nhap or f_xuat) and f_xnt_tho and f_tk)
    if st.button("🔍 Chạy Đối chiếu XNT", key="btn_xnt", disabled=not ready_xnt):
        with st.spinner("Đang xử lý XNT..."):
            try:
                dfs_nx = []
                for f in (f_nhap or []):
                    f.seek(0); dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))
                for f in (f_xuat or []):
                    f.seek(0); dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))
                # Lưu global_map
                st.session_state['global_map'] = extract_ma_map(dfs_nx)
                f_xnt_tho.seek(0); f_tk.seek(0)
                df_xnt_raw = pd.read_excel(io.BytesIO(f_xnt_tho.read()), sheet_name=0, header=None)
                df_tk_raw  = pd.read_excel(io.BytesIO(f_tk.read()), sheet_name=0, header=None)
                df_res, err = run_xnt(dfs_nx, df_xnt_raw, df_tk_raw, {})
                if err: st.error(f"❌ {err}"); st.stop()
                st.session_state.update({'xnt_result': df_res, 'xnt_done': True, 'shared_tn': thang_nam})
                st.rerun()
            except Exception as e:
                st.error(f"❌ Lỗi: {e}"); st.exception(e)

    if st.session_state.get('xnt_done'):
        df_r = st.session_state['xnt_result']
        dm   = df_r[df_r['cl'].notna()]
        nk   = (dm['cl'].abs()<0.01).sum(); nl = (dm['cl'].abs()>=0.01).sum()
        dn_tk  = df_r[df_r['method']=='no_tk'];  dn_xnt = df_r[df_r['method']=='no_xnt']
        st.markdown(f"""<div class="ok-box"><div style="font-size:2rem">✅</div>
          <h3>Đối chiếu XNT hoàn tất – {st.session_state.get('shared_tn','')}</h3></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="num" style="color:#166534">{nk}</div><div class="lbl">✅ Khớp</div></div>
          <div class="stat-card"><div class="num" style="color:#dc2626">{nl}</div><div class="lbl">⚠️ Chênh lệch</div></div>
          <div class="stat-card"><div class="num" style="color:#2563a8">{len(dn_tk)}</div><div class="lbl">📋 HPT có – TK không</div></div>
          <div class="stat-card"><div class="num" style="color:#d97706">{len(dn_xnt)}</div><div class="lbl">📋 TK có – HPT không</div></div>
        </div>""", unsafe_allow_html=True)
        dl = dm[dm['cl'].abs()>=0.01].sort_values('cl')
        if len(dl):
            st.markdown(f"**⚠️ {nl} dòng chênh lệch:**")
            st.dataframe(dl[['ma','ten','nd','ton_xnt','ton_tk','cl']].rename(columns={
                'ma':'Mã HPT','ten':'Tên thuốc','nd':'Nồng độ',
                'ton_xnt':'Tồn HPT','ton_tk':'Tồn TK','cl':'Chênh lệch'}),
                use_container_width=True, hide_index=True)

# ─── TAB 2: KIỂM NHẬP ────────────────────────────────────────────────────────
with tab_kn:
    st.markdown("""<div class="info-box">
    Upload file <b>BBKN</b> ở trên. File TK Thống kê và bảng mã HPT dùng chung từ khu vực upload.
    <br>App tự động <b>cộng tổng</b> tất cả hóa đơn cùng tên+nồng độ+giá trước khi đối chiếu.
    </div>""", unsafe_allow_html=True)

    ready_kn = bool(f_bbkn and f_tk)
    if st.button("🔍 Chạy Đối chiếu Kiểm nhập", key="btn_kn", disabled=not ready_kn):
        with st.spinner("Đang xử lý Kiểm nhập..."):
            try:
                f_bbkn.seek(0); f_tk.seek(0)
                df_bbkn_raw = pd.read_excel(io.BytesIO(f_bbkn.read()), sheet_name=0, header=None)
                df_tk_raw   = pd.read_excel(io.BytesIO(f_tk.read()), sheet_name=0, header=None)
                gmap_cur    = st.session_state.get('global_map', pd.DataFrame())
                df_res, err = run_kn(df_bbkn_raw, df_tk_raw, gmap_cur, sl_col_kn)
                if err: st.error(f"❌ {err}"); st.stop()
                st.session_state.update({'kn_result': df_res, 'kn_done': True, 'shared_tn': thang_nam})
            except Exception as e:
                st.error(f"❌ Lỗi: {e}"); st.exception(e)

    if st.session_state.get('kn_done'):
        df_kn = st.session_state['kn_result']
        dm_kn = df_kn[df_kn['status']=='matched']
        nk_kn = (dm_kn['cl'].abs()<0.01).sum(); nl_kn = (dm_kn['cl'].abs()>=0.01).sum()
        no_tk = df_kn[df_kn['status']=='hpt_no_tk']; no_hpt = df_kn[df_kn['status']=='tk_no_hpt']
        st.markdown(f"""<div class="ok-box"><div style="font-size:2rem">✅</div>
          <h3>Đối chiếu Kiểm nhập hoàn tất – {st.session_state.get('shared_tn','')}</h3></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="num" style="color:#166534">{nk_kn}</div><div class="lbl">✅ Khớp</div></div>
          <div class="stat-card"><div class="num" style="color:#dc2626">{nl_kn}</div><div class="lbl">⚠️ Chênh lệch</div></div>
          <div class="stat-card"><div class="num" style="color:#2563a8">{len(no_tk)}</div><div class="lbl">📋 HPT có – TK không</div></div>
          <div class="stat-card"><div class="num" style="color:#d97706">{len(no_hpt)}</div><div class="lbl">📋 TK có – HPT không</div></div>
        </div>""", unsafe_allow_html=True)
        dl_kn = dm_kn[dm_kn['cl'].abs()>=0.01].sort_values('cl')
        if len(dl_kn):
            st.markdown(f"**⚠️ {nl_kn} dòng chênh lệch:**")
            st.dataframe(dl_kn[['ma','ten_hpt','nd','nhap_hpt','nhap_tk','cl']].rename(columns={
                'ma':'Mã HPT','ten_hpt':'Tên thuốc (HPT)','nd':'Nồng độ',
                'nhap_hpt':'Nhập HPT (tổng)','nhap_tk':'Nhập TK','cl':'Chênh lệch'}),
                use_container_width=True, hide_index=True)
        if len(no_hpt):
            st.markdown(f"**📋 {len(no_hpt)} mã – TK có nhập, HPT không có:**")
            st.dataframe(no_hpt[['ma','ten_tk','nd','nhap_tk']].rename(columns={
                'ma':'Mã HPT','ten_tk':'Tên thuốc TK','nd':'Nồng độ','nhap_tk':'Nhập TK'}),
                use_container_width=True, hide_index=True)

# ─── TAB 3: KIỂM KÊ ──────────────────────────────────────────────────────────
with tab_kk:
    st.markdown("""<div class="info-box">
    Upload file <b>BBKK</b> ở trên. File TK Thống kê và bảng mã HPT dùng chung.
    <br>App tự động <b>cộng tổng</b> các dòng cùng tên+nồng độ+giá (kể cả dòng điều chỉnh âm).
    </div>""", unsafe_allow_html=True)

    ready_kk = bool(f_bbkk and f_tk)
    if st.button("🔍 Chạy Đối chiếu Kiểm kê", key="btn_kk", disabled=not ready_kk):
        with st.spinner("Đang xử lý Kiểm kê..."):
            try:
                f_bbkk.seek(0); f_tk.seek(0)
                df_bbkk_raw = pd.read_excel(io.BytesIO(f_bbkk.read()), sheet_name=0, header=None)
                df_tk_raw   = pd.read_excel(io.BytesIO(f_tk.read()), sheet_name=0, header=None)
                gmap_cur    = st.session_state.get('global_map', pd.DataFrame())
                df_res, err = run_kk(df_bbkk_raw, df_tk_raw, gmap_cur, sl_col_kk)
                if err: st.error(f"❌ {err}"); st.stop()
                st.session_state.update({'kk_result': df_res, 'kk_done': True, 'shared_tn': thang_nam})
            except Exception as e:
                st.error(f"❌ Lỗi: {e}"); st.exception(e)

    if st.session_state.get('kk_done'):
        df_kk2 = st.session_state['kk_result']
        dm_kk  = df_kk2[df_kk2['status']=='matched']
        nk_kk  = (dm_kk['cl'].abs()<0.01).sum(); nl_kk = (dm_kk['cl'].abs()>=0.01).sum()
        no_tk2 = df_kk2[df_kk2['status']=='hpt_no_tk']; no_hpt2 = df_kk2[df_kk2['status']=='tk_no_hpt']
        st.markdown(f"""<div class="ok-box"><div style="font-size:2rem">✅</div>
          <h3>Đối chiếu Kiểm kê hoàn tất – {st.session_state.get('shared_tn','')}</h3></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="num" style="color:#166534">{nk_kk}</div><div class="lbl">✅ Khớp</div></div>
          <div class="stat-card"><div class="num" style="color:#dc2626">{nl_kk}</div><div class="lbl">⚠️ Chênh lệch</div></div>
          <div class="stat-card"><div class="num" style="color:#2563a8">{len(no_tk2)}</div><div class="lbl">📋 HPT có – TK không</div></div>
          <div class="stat-card"><div class="num" style="color:#d97706">{len(no_hpt2)}</div><div class="lbl">📋 TK có – HPT không</div></div>
        </div>""", unsafe_allow_html=True)
        dl_kk = dm_kk[dm_kk['cl'].abs()>=0.01].sort_values('cl')
        if len(dl_kk):
            st.markdown(f"**⚠️ {nl_kk} dòng chênh lệch:**")
            st.dataframe(dl_kk[['ma','ten_hpt','nd','sl_kk','ton_tk','cl']].rename(columns={
                'ma':'Mã HPT','ten_hpt':'Tên thuốc (HPT)','nd':'Nồng độ',
                'sl_kk':'SL Kiểm kê','ton_tk':'Tồn TK','cl':'Chênh lệch'}),
                use_container_width=True, hide_index=True)
        if len(no_hpt2):
            st.markdown(f"**📋 {len(no_hpt2)} mã – TK có tồn, HPT không có:**")
            st.dataframe(no_hpt2[['ma','ten_tk','nd','ton_tk']].rename(columns={
                'ma':'Mã HPT','ten_tk':'Tên thuốc TK','nd':'Nồng độ','ton_tk':'Tồn TK'}),
                use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
#  XUẤT EXCEL TỔNG HỢP
# ══════════════════════════════════════════════════════════════════════════════
has_any = any(st.session_state.get(k) for k in ['xnt_done','kn_done','kk_done'])
if has_any:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### 📥 Tải kết quả tổng hợp")
    st.markdown("""<div class="info-box">
    File Excel gộp tất cả kết quả vào các Sheet riêng biệt với tên đồng nhất:
    <b>HPT có – TK không</b> / <b>TK có – HPT không</b>.
    Sheet KN và KK có thêm cột <b>Mã HPT</b> và <b>Danh sách Hóa đơn</b>.
    Màu <span style="background:#E2EFDA;padding:1px 6px;border-radius:3px">🟢 Xanh</span> = Khớp |
    <span style="background:#FFD7D7;padding:1px 6px;border-radius:3px">🔴 Đỏ</span> = Lệch.
    </div>""", unsafe_allow_html=True)
    tn = st.session_state.get('shared_tn', thang_nam)
    excel_bytes = export_excel(
        st.session_state.get('xnt_result'),
        st.session_state.get('kn_result'),
        st.session_state.get('kk_result'),
        tn)
    st.download_button(
        label=f"⬇️  Tải Kết Quả Tổng Hợp {tn} (.xlsx)",
        data=excel_bytes,
        file_name=f"doi_chieu_tong_hop_{tn.replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
