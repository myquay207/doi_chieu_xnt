"""
ĐỐI CHIẾU DƯỢC – Bệnh viện Đà Nẵng
3 module: XNT · Kiểm nhập · Kiểm kê
v3 – Global Map, multi-mã, bỏ dòng rác
"""

import io, re, warnings
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Đối Chiếu Dược – BV Đà Nẵng",
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
.map-box{background:#f0fdf4;border:1.5px solid #86efac;border-radius:12px;
  padding:12px 18px;margin:10px 0;font-size:.86rem;color:#166534;}
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
#  HÀM TIỆN ÍCH DÙNG CHUNG
# ══════════════════════════════════════════════════════════════════════════════

def norm(s):
    """Chuẩn hóa tên thuốc / nồng độ để so khớp."""
    if not isinstance(s, str):
        return ""
    s = s.strip().lower()
    s = re.sub(r'_x000a_', ' ', s)
    s = re.sub(r'\n', ' ', s)
    s = re.sub(r'[#]', '', s)
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    s = re.sub(r'(\d)\s+(mg|ml|mcg|μg|g|iu|%|meq|l)', r'\1\2', s)
    s = re.sub(r'\s*\+\s*', '+', s)
    return re.sub(r'\s+', ' ', s).strip()


def is_valid_drug_name(s):
    """Kiểm tra chuỗi có phải tên thuốc thực sự không (loại bỏ số thứ tự cột)."""
    if not isinstance(s, str):
        return False
    s = s.strip()
    if len(s) < 2:
        return False
    try:
        float(s)
        return False     # Là số → không phải tên thuốc
    except ValueError:
        return True


def parse_tk(df_tk_raw):
    """
    Parse file Thống kê XNT.
    Trả về DataFrame với:
      ma, ten_tk, nd_tk, gia_tk, ton_tk (col 24), nhap_tk (col 14)
      kten, knd, kgia
    Giữ NGUYÊN từng dòng (không gộp) để hỗ trợ nhiều mã cùng tên+giá.
    """
    tk_rows = []
    for _, row in df_tk_raw.iloc[5:].iterrows():
        ma  = str(row[4]).strip() if not pd.isna(row[4]) else ''
        ten = str(row[5]).strip() if not pd.isna(row[5]) else ''
        if not ma or not ten:
            continue
        gia = row[11] if not pd.isna(row[11]) else 0
        tk_rows.append({
            'ma':      ma,
            'ten_tk':  ten,
            'nd_tk':   str(row[8]).strip() if not pd.isna(row[8]) else '',
            'gia_tk':  gia,
            'ton_tk':  float(row[24]) if not pd.isna(row[24]) else 0,
            'nhap_tk': float(row[14]) if not pd.isna(row[14]) else 0,
            'kten':    norm(ten),
            'knd':     norm(str(row[8]) if not pd.isna(row[8]) else ''),
            'kgia':    int(round(float(gia))) if gia else 0,
        })
    return pd.DataFrame(tk_rows) if tk_rows else pd.DataFrame()


def extract_ma_map(dfs_nhap_xuat):
    """
    Xây bảng map mã HPT từ file nhập/xuất.
    Cấu trúc: col0=STT, col1=Mã HPT, col2=Tên, col3=Nồng độ, col5=Đơn giá
    Trả về DataFrame với: ma, ten, nd, gia, kten, knd, kgia
    Giữ từng dòng riêng (kể cả cùng tên+giá khác mã).
    """
    all_rows = []
    for df in dfs_nhap_xuat:
        if df.empty:
            continue
        for _, row in df.iterrows():
            try:
                int(str(row[0]).strip())
            except:
                continue
            if pd.isna(row[1]) or pd.isna(row[2]):
                continue
            ten = str(row[2]).strip()
            if not is_valid_drug_name(ten):
                continue
            gia = row[5] if not pd.isna(row[5]) else 0
            all_rows.append({
                'ma':   str(row[1]).strip(),
                'ten':  ten,
                'nd':   str(row[3]).strip() if not pd.isna(row[3]) else '',
                'gia':  gia,
                'kten': norm(ten),
                'knd':  norm(str(row[3]) if not pd.isna(row[3]) else ''),
                'kgia': int(round(float(gia))) if gia else 0,
            })
    if not all_rows:
        return pd.DataFrame()
    df_map = pd.DataFrame(all_rows).drop_duplicates(subset=['ma', 'kten', 'knd', 'kgia'])
    return df_map.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
#  PARSER CHO BBKN  (Biên bản kiểm nhập)
# ══════════════════════════════════════════════════════════════════════════════

def parse_bbkn(df_raw, sl_col=9):
    """
    Parse file Biên bản kiểm nhập.
    Cấu trúc: col0=STT, col1=Mã HĐ, col2=Tên, col3=Nồng độ,
              col4=DVT, col5=Lô, col6=Hãng SX, col7=Hạn SD,
              col8=Đơn giá, col9=Số lượng, col10=Thành tiền

    Hàng rác (row 7 trong file gốc): col0='1', col2='3', ... → bị lọc bởi is_valid_drug_name.
    Hàng công ty (giữa các mặt hàng): col0=int nhưng col2=NaN/tên công ty → bị lọc.
    Trả về DataFrame: ten, nd, gia, sl, ma_hd, kten, knd, kgia
    KHÔNG gộp mã — mỗi dòng dữ liệu giữ nguyên.
    """
    rows = []
    for _, row in df_raw.iterrows():
        try:
            int(str(row[0]).strip())
        except:
            continue
        if pd.isna(row[2]):
            continue
        ten = str(row[2]).strip()
        if not is_valid_drug_name(ten):
            continue
        nd  = str(row[3]).strip() if not pd.isna(row[3]) else ''
        gia = float(row[8]) if not pd.isna(row[8]) else 0
        sl_val = sl_col if isinstance(sl_col, int) else 9
        sl  = float(row[sl_val]) if not pd.isna(row[sl_val]) else 0
        ma_hd = str(row[1]).strip() if not pd.isna(row[1]) else ''
        rows.append({
            'ten':   ten,
            'nd':    nd,
            'gia':   gia,
            'sl':    sl,
            'ma_hd': ma_hd,
            'kten':  norm(ten),
            'knd':   norm(nd),
            'kgia':  int(round(gia)) if gia else 0,
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
#  PARSER CHO BBKK  (Biên bản kiểm kê)
# ══════════════════════════════════════════════════════════════════════════════

def parse_bbkk(df_raw, sl_col=8):
    """
    Parse file Biên bản kiểm kê.
    Cấu trúc (từ row 10): col0=STT, col1=Tên, col2=Nồng độ, col3=DVT,
                           col4=Đơn giá, col5=Lô, col6=Hãng SX, col7=Hạn SD,
                           col8=Thực tế, col9=Thành tiền, col10=Sổ sách, ...
    Trả về DataFrame: ten, nd, gia, sl, kten, knd, kgia
    KHÔNG gộp — giữ từng dòng.
    """
    rows = []
    for _, row in df_raw.iterrows():
        try:
            int(str(row[0]).strip())
        except:
            continue
        if pd.isna(row[1]):
            continue
        ten = str(row[1]).strip()
        if not is_valid_drug_name(ten):
            continue
        nd  = str(row[2]).strip() if not pd.isna(row[2]) else ''
        gia = float(row[4]) if not pd.isna(row[4]) else 0
        sl_val = sl_col if isinstance(sl_col, int) else 8
        sl  = float(row[sl_val]) if not pd.isna(row[sl_val]) else 0
        rows.append({
            'ten':  ten,
            'nd':   nd,
            'gia':  gia,
            'sl':   sl,
            'kten': norm(ten),
            'knd':  norm(nd),
            'kgia': int(round(gia)) if gia else 0,
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE 1 – ĐỐI CHIẾU XNT (giữ nguyên logic gốc)
# ══════════════════════════════════════════════════════════════════════════════

def run_reconciliation_xnt(files_nhap_xuat, df_xnt_raw, df_tk_raw, confirmed_override):
    df_map = extract_ma_map(files_nhap_xuat)
    if df_map.empty:
        return None, "Không đọc được dữ liệu từ file nhập/xuất"

    # Augmentin fallback
    aug = pd.DataFrame([{
        'ma': '0005301225', 'ten': 'Augmentin 1g', 'nd': '875mg + 125mg',
        'gia': 16680, 'kten': norm('Augmentin 1g'),
        'knd': norm('875mg + 125mg'), 'kgia': 16680
    }])
    df_map = pd.concat([df_map, aug], ignore_index=True).drop_duplicates(
        subset=['ma', 'kten', 'knd', 'kgia'])

    for i, row in df_map.iterrows():
        if row['ten'] in confirmed_override:
            df_map.at[i, 'ma'] = confirmed_override[row['ten']]

    # XNT thô
    xnt_rows = []
    for _, row in df_xnt_raw.iterrows():
        try:
            stt = int(str(row[0]).strip())
        except:
            continue
        if pd.isna(row[2]) or not isinstance(row[2], str):
            continue
        if row[2].strip().isdigit():
            continue
        xnt_rows.append({
            'stt':     stt,
            'ten':     str(row[2]).strip(),
            'nd':      str(row[3]).strip() if not pd.isna(row[3]) else '',
            'gia':     row[8]  if not pd.isna(row[8]) else 0,
            'ton_xnt': float(row[12]) if not pd.isna(row[12]) else 0,
            'kten':    norm(str(row[2])),
            'knd':     norm(str(row[3]) if not pd.isna(row[3]) else ''),
            'kgia':    int(round(float(row[8]))) if not pd.isna(row[8]) else 0,
        })
    df_xnt = pd.DataFrame(xnt_rows)
    df_tk  = parse_tk(df_tk_raw)

    results = []
    used_tk = set()

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
        mask  = ((df_tk['kten'] == kten) & (df_tk['knd'] == knd) & (df_tk['kgia'] == kgia))
        grp_t = df_tk[mask & ~df_tk.index.isin(used_tk)].copy()

        if len(grp_x) == 1 and len(grp_t) == 1:
            xr = grp_x.iloc[0]; tr = grp_t.iloc[0]
            results.append(make_row(xr, tr, '1-1'))
            used_tk.add(grp_t.index[0])
        elif len(grp_x) >= 1 and len(grp_t) > 0:
            xl = grp_x.reset_index(drop=True)
            tl = grp_t.reset_index(drop=True)
            matched_x, matched_t = set(), set()
            for xi, xr in xl.iterrows():
                for ti, tr in tl.iterrows():
                    if ti in matched_t:
                        continue
                    if abs(xr['ton_xnt'] - tr['ton_tk']) < 0.01:
                        results.append(make_row(xr, tr, 'exact_ton'))
                        used_tk.add(grp_t.index[ti])
                        matched_x.add(xi); matched_t.add(ti); break
            for xi, xr in xl.iterrows():
                if xi in matched_x:
                    continue
                best_diff, best_ti, best_tr = float('inf'), None, None
                for ti, tr in tl.iterrows():
                    if ti in matched_t:
                        continue
                    d = abs(xr['ton_xnt'] - tr['ton_tk'])
                    if d < best_diff:
                        best_diff, best_ti, best_tr = d, ti, tr
                if best_tr is not None:
                    results.append(make_row(xr, best_tr, 'nearest_ton'))
                    used_tk.add(grp_t.index[best_ti]); matched_t.add(best_ti)
                else:
                    results.append(make_row(xr, None, 'no_tk'))
        else:
            for _, xr in grp_x.iterrows():
                for _, xr in grp_x.iterrows():
    results.append(make_row(xr, None, 'no_tk'))

    for idx, tr in df_tk.iterrows():
        if idx not in used_tk:
            results.append({
                'ma': tr['ma'], 'ten': '', 'nd': '', 'gia': '',
                'ton_xnt': None, 'ten_tk': tr['ten_tk'],
                'nd_tk': tr['nd_tk'], 'ton_tk': tr['ton_tk'],
                'method': 'no_xnt', 'cl': None,
            })

    return pd.DataFrame(results), None


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE 2 – ĐỐI CHIẾU KIỂM NHẬP  (v3: dùng global_map, tách mã)
# ══════════════════════════════════════════════════════════════════════════════

def run_reconciliation_kn(df_bbkn_raw, df_tk_raw, global_map, sl_col=9):
    """
    Đối chiếu số nhập BBKN vs số nhập Thống kê, tách theo từng mã hàng.

    Quy trình:
    1. Parse BBKN → danh sách dòng thuốc (KHÔNG gộp mã).
    2. Với mỗi dòng BBKN:
       a. Tra global_map theo kten+knd+kgia → lấy ra TẤT CẢ mã hàng tương ứng.
       b. Với mỗi mã, tìm dòng trong TK có cùng mã → lấy nhap_tk.
       c. Nếu BBKN có nhiều dòng cùng tên+nồng độ+giá, khớp theo thứ tự STT.
    3. Dòng TK có nhập nhưng không có trong BBKN → "no_kn".
    """
    df_kn = parse_bbkn(df_bbkn_raw, sl_col)
    if df_kn.empty:
        return None, "Không đọc được dữ liệu từ file BBKN"

    df_tk = parse_tk(df_tk_raw)
    if df_tk.empty:
        return None, "Không đọc được dữ liệu từ file Thống kê"

    results = []
    used_tk = set()

    # Nhóm BBKN theo kten+knd+kgia (giữ nguyên từng dòng)
    for (kten, knd, kgia), grp_kn in df_kn.groupby(['kten', 'knd', 'kgia'], sort=False):
        # Tất cả dòng TK khớp key này
        mask_tk = ((df_tk['kten'] == kten) & (df_tk['knd'] == knd) & (df_tk['kgia'] == kgia))
        grp_tk  = df_tk[mask_tk & ~df_tk.index.isin(used_tk)].reset_index()

        kn_list = grp_kn.reset_index(drop=True)
        n_kn    = len(kn_list)
        n_tk    = len(grp_tk)

        if n_tk == 0:
            # BBKN có nhưng TK không
            for _, kr in kn_list.iterrows():
                results.append({
                    'ma': '', 'ten_kn': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                    'nhap_kn': kr['sl'], 'ten_tk': '', 'nhap_tk': None,
                    'cl': None, 'status': 'no_tk'
                })
            continue

        if n_kn == 1 and n_tk == 1:
            # Đơn giản 1-1
            kr = kn_list.iloc[0]; tr = grp_tk.iloc[0]
            cl = kr['sl'] - tr['nhap_tk']
            results.append({
                'ma': tr['ma'], 'ten_kn': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                'nhap_kn': kr['sl'], 'ten_tk': tr['ten_tk'], 'nhap_tk': tr['nhap_tk'],
                'cl': cl, 'status': 'matched'
            })
            used_tk.add(grp_tk.iloc[0]['index'])

        elif n_kn >= 1 and n_tk >= 1:
            # Nhiều-nhiều: khớp theo thứ tự xuất hiện (vị trí tương ứng)
            matched_tk = set()
            for i, kr in kn_list.iterrows():
                # Tìm TK chưa dùng, khớp gần nhất về số lượng
                best = None; best_diff = float('inf')
                for j, tr in grp_tk.iterrows():
                    if j in matched_tk:
                        continue
                    d = abs(kr['sl'] - tr['nhap_tk'])
                    if d < best_diff:
                        best_diff, best = d, j
                if best is not None:
                    tr = grp_tk.loc[best]
                    cl = kr['sl'] - tr['nhap_tk']
                    results.append({
                        'ma': tr['ma'], 'ten_kn': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                        'nhap_kn': kr['sl'], 'ten_tk': tr['ten_tk'], 'nhap_tk': tr['nhap_tk'],
                        'cl': cl, 'status': 'matched'
                    })
                    used_tk.add(tr['index']); matched_tk.add(best)
                else:
                    results.append({
                        'ma': '', 'ten_kn': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                        'nhap_kn': kr['sl'], 'ten_tk': '', 'nhap_tk': None,
                        'cl': None, 'status': 'no_tk'
                    })
            # TK dư (không có trong BBKN)
            for j, tr in grp_tk.iterrows():
                if j not in matched_tk:
                    results.append({
                        'ma': tr['ma'], 'ten_kn': '', 'nd': tr['nd_tk'], 'gia': tr['gia_tk'],
                        'nhap_kn': None, 'ten_tk': tr['ten_tk'], 'nhap_tk': tr['nhap_tk'],
                        'cl': None, 'status': 'no_kn'
                    })
                    used_tk.add(tr['index'])

    # TK còn lại có nhập, BBKN không có
    for idx, tr in df_tk.iterrows():
        if idx not in used_tk:
            results.append({
                'ma': tr['ma'], 'ten_kn': '', 'nd': tr['nd_tk'], 'gia': tr['gia_tk'],
                'nhap_kn': None, 'ten_tk': tr['ten_tk'], 'nhap_tk': tr['nhap_tk'],
                'cl': None, 'status': 'no_kn'
            })

    return pd.DataFrame(results), None


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE 3 – ĐỐI CHIẾU KIỂM KÊ  (v3: dùng global_map, tách mã, không gộp)
# ══════════════════════════════════════════════════════════════════════════════

def run_reconciliation_kk(df_bbkk_raw, df_tk_raw, global_map, sl_col=8):
    """
    Đối chiếu số lượng kiểm kê thực tế (BBKK) vs Tồn cuối TK.

    Quy trình:
    1. Parse BBKK → danh sách dòng thuốc (KHÔNG gộp).
    2. Nhóm BBKK theo kten+knd+kgia.
    3. Với mỗi nhóm, tìm tất cả dòng TK cùng key.
    4. Khớp từng cặp (BBKK_i, TK_j) theo thứ tự/số lượng gần nhất.
    5. TK có tồn nhưng không xuất hiện trong BBKK → "no_kk".
    """
    df_kk = parse_bbkk(df_bbkk_raw, sl_col)
    if df_kk.empty:
        return None, "Không đọc được dữ liệu từ file Biên bản kiểm kê"

    df_tk = parse_tk(df_tk_raw)
    if df_tk.empty:
        return None, "Không đọc được dữ liệu từ file Thống kê"

    results = []
    used_tk = set()

    for (kten, knd, kgia), grp_kk in df_kk.groupby(['kten', 'knd', 'kgia'], sort=False):
        mask_tk = ((df_tk['kten'] == kten) & (df_tk['knd'] == knd) & (df_tk['kgia'] == kgia))
        grp_tk  = df_tk[mask_tk & ~df_tk.index.isin(used_tk)].reset_index()

        kk_list = grp_kk.reset_index(drop=True)
        n_kk    = len(kk_list)
        n_tk    = len(grp_tk)

        if n_tk == 0:
            for _, kr in kk_list.iterrows():
                results.append({
                    'ma': '', 'ten_kk': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                    'sl_kk': kr['sl'], 'ten_tk': '', 'ton_tk': None,
                    'cl': None, 'status': 'no_tk'
                })
            continue

        if n_kk == 1 and n_tk == 1:
            kr = kk_list.iloc[0]; tr = grp_tk.iloc[0]
            cl = kr['sl'] - tr['ton_tk']
            results.append({
                'ma': tr['ma'], 'ten_kk': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                'sl_kk': kr['sl'], 'ten_tk': tr['ten_tk'], 'ton_tk': tr['ton_tk'],
                'cl': cl, 'status': 'matched'
            })
            used_tk.add(grp_tk.iloc[0]['index'])

        else:
            # Nhiều-nhiều: khớp theo số lượng gần nhất
            matched_tk = set()
            for i, kr in kk_list.iterrows():
                best = None; best_diff = float('inf')
                for j, tr in grp_tk.iterrows():
                    if j in matched_tk:
                        continue
                    d = abs(kr['sl'] - tr['ton_tk'])
                    if d < best_diff:
                        best_diff, best = d, j
                if best is not None:
                    tr = grp_tk.loc[best]
                    cl = kr['sl'] - tr['ton_tk']
                    results.append({
                        'ma': tr['ma'], 'ten_kk': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                        'sl_kk': kr['sl'], 'ten_tk': tr['ten_tk'], 'ton_tk': tr['ton_tk'],
                        'cl': cl, 'status': 'matched'
                    })
                    used_tk.add(tr['index']); matched_tk.add(best)
                else:
                    results.append({
                        'ma': '', 'ten_kk': kr['ten'], 'nd': kr['nd'], 'gia': kr['gia'],
                        'sl_kk': kr['sl'], 'ten_tk': '', 'ton_tk': None,
                        'cl': None, 'status': 'no_tk'
                    })
            # TK dư
            for j, tr in grp_tk.iterrows():
                if j not in matched_tk:
                    results.append({
                        'ma': tr['ma'], 'ten_kk': '', 'nd': tr['nd_tk'], 'gia': tr['gia_tk'],
                        'sl_kk': None, 'ten_tk': tr['ten_tk'], 'ton_tk': tr['ton_tk'],
                        'cl': None, 'status': 'no_kk'
                    })
                    used_tk.add(tr['index'])

    # TK còn lại có tồn, KK không có
    for idx, tr in df_tk.iterrows():
        if idx not in used_tk:
            results.append({
                'ma': tr['ma'], 'ten_kk': '', 'nd': tr['nd_tk'], 'gia': tr['gia_tk'],
                'sl_kk': None, 'ten_tk': tr['ten_tk'], 'ton_tk': tr['ton_tk'],
                'cl': None, 'status': 'no_kk'
            })

    return pd.DataFrame(results), None


# ══════════════════════════════════════════════════════════════════════════════
#  XUẤT EXCEL – GỘP TẤT CẢ SHEET
# ══════════════════════════════════════════════════════════════════════════════

def _hdr(ws, headers, fill_hex='1F3864'):
    THIN = Side(style='thin'); MED = Side(style='medium')
    F_H  = PatternFill('solid', fgColor=fill_hex)
    for ci, (h, w) in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = Font(name='Times New Roman', bold=True, size=11, color='FFFFFF')
        c.fill = F_H
        c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 36


def _row(ws, ri, vals, fill, right_cols=(), center_cols=()):
    THIN = Side(style='thin')
    for ci, v in enumerate(vals, 1):
        safe = '' if (v is None or (isinstance(v, float) and pd.isna(v))) else v
        cell = ws.cell(row=ri, column=ci, value=safe)
        cell.font = Font(name='Times New Roman', size=11)
        cell.fill = fill
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        if ci in right_cols:
            cell.alignment = Alignment(vertical='center', horizontal='right')
        elif ci in center_cols:
            cell.alignment = Alignment(vertical='center', horizontal='center')
        else:
            cell.alignment = Alignment(vertical='center', horizontal='left')
        if ci in right_cols and isinstance(safe, (int, float)):
            cell.number_format = '#,##0.##'
    ws.row_dimensions[ri].height = 18


def build_excel_xnt(wb, df_res, thang_nam):
    F_OK = PatternFill('solid', fgColor='E2EFDA')
    F_L  = PatternFill('solid', fgColor='FFD7D7')
    F_W  = PatternFill('solid', fgColor='FFF9C4')
    F_XT = PatternFill('solid', fgColor='DEEBF7')
    F_TK = PatternFill('solid', fgColor='FCE4D6')

    ws1 = wb.create_sheet(f"DC XNT {thang_nam.replace('/', '_')}")
    _hdr(ws1, [('Mã HPT',16),('Tên thuốc (XNT)',32),('Tên thuốc (TK)',28),
               ('Nồng độ',22),('Đơn giá',12),('Tồn cuối XNT',13),
               ('Tồn cuối TK',13),('Chênh lệch',13),('Trạng thái',25),('Phương pháp',18)])

    df_matched = df_res[df_res['cl'].notna()].copy()
    df_lech = df_matched[df_matched['cl'].abs() >= 0.01].sort_values('cl')
    df_khop = df_matched[df_matched['cl'].abs() < 0.01]
    mm = {'1-1': 'Chính xác 1-1', 'exact_ton': 'Khớp tồn chính xác', 'nearest_ton': 'Gần nhất ⚠️'}
    for ri, (_, r) in enumerate(pd.concat([df_lech, df_khop], ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl) < 0.01:   fill, st = F_OK, '✅ Khớp'
        elif cl > 0:          fill, st = F_L,  f'⬆️ HPT cao hơn {cl:+.0f}'
        else:                 fill, st = F_L,  f'⬇️ HPT thấp hơn {cl:+.0f}'
        if r['method'] == 'nearest_ton' and abs(cl) >= 0.01: fill = F_W
        _row(ws1, ri, [r['ma'], r['ten'], r.get('ten_tk', ''), r['nd'], r['gia'],
                       r['ton_xnt'], r['ton_tk'], cl, st, mm.get(r['method'], '')],
             fill, right_cols=(5, 6, 7, 8), center_cols=(1, 9, 10))
    ws1.freeze_panes = 'A2'

    df_no_tk  = df_res[df_res['method'] == 'no_tk']
    df_no_xnt = df_res[df_res['method'] == 'no_xnt']
    ws2 = wb.create_sheet("XNT - HPT có TK không")
    _hdr(ws2, [('Tên thuốc', 35), ('Nồng độ', 25), ('Đơn giá', 12), ('Tồn XNT', 12), ('Ghi chú', 35)])
    for ri, (_, r) in enumerate(df_no_tk.iterrows(), 2):
        _row(ws2, ri, [r['ten'], r['nd'], r['gia'], r['ton_xnt'], 'HPT có nhưng TK không theo dõi'],
             F_XT, right_cols=(3, 4))

    ws3 = wb.create_sheet("XNT - TK có HPT không")
    _hdr(ws3, [('Mã HPT', 16), ('Tên thuốc TK', 35), ('Nồng độ TK', 25), ('Tồn TK', 12), ('Ghi chú', 35)])
    for ri, (_, r) in enumerate(df_no_xnt.iterrows(), 2):
        _row(ws3, ri, [r['ma'], r['ten_tk'], r['nd_tk'], r['ton_tk'],
                       'TK có nhưng XNT không phát sinh'], F_TK, right_cols=(4,))


def build_excel_kn(wb, df_res, thang_nam):
    F_OK = PatternFill('solid', fgColor='E2EFDA')
    F_L  = PatternFill('solid', fgColor='FFD7D7')
    F_XT = PatternFill('solid', fgColor='DEEBF7')
    F_TK = PatternFill('solid', fgColor='FCE4D6')

    ws = wb.create_sheet(f"DC Kiểm nhập {thang_nam.replace('/', '_')}")
    _hdr(ws, [('Mã HPT', 16), ('Tên thuốc (BBKN)', 32), ('Tên thuốc (TK)', 28),
              ('Nồng độ', 22), ('Đơn giá', 12), ('Nhập BBKN', 12),
              ('Nhập TK', 12), ('Chênh lệch', 13), ('Trạng thái', 25)])

    df_m  = df_res[df_res['status'] == 'matched']
    df_l  = df_m[df_m['cl'].abs() >= 0.01].sort_values('cl')
    df_k  = df_m[df_m['cl'].abs() < 0.01]
    for ri, (_, r) in enumerate(pd.concat([df_l, df_k], ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl) < 0.01:  fill, st = F_OK, '✅ Khớp'
        elif cl > 0:         fill, st = F_L,  f'⬆️ BBKN cao hơn {cl:+.0f}'
        else:                fill, st = F_L,  f'⬇️ BBKN thấp hơn {cl:+.0f}'
        _row(ws, ri, [r['ma'], r['ten_kn'], r['ten_tk'], r['nd'], r['gia'],
                      r['nhap_kn'], r['nhap_tk'], cl, st],
             fill, right_cols=(5, 6, 7, 8), center_cols=(1, 9))
    ws.freeze_panes = 'A2'

    ws2 = wb.create_sheet("KN - BBKN có TK không")
    _hdr(ws2, [('Tên thuốc', 35), ('Nồng độ', 25), ('Đơn giá', 12), ('Nhập BBKN', 12), ('Ghi chú', 35)])
    for ri, (_, r) in enumerate(df_res[df_res['status'] == 'no_tk'].iterrows(), 2):
        _row(ws2, ri, [r['ten_kn'], r['nd'], r['gia'], r['nhap_kn'],
                       'BBKN có nhưng TK không có số nhập'], F_XT, right_cols=(3, 4))

    ws3 = wb.create_sheet("KN - TK có BBKN không")
    _hdr(ws3, [('Mã HPT', 16), ('Tên thuốc TK', 35), ('Nồng độ TK', 25), ('Nhập TK', 12), ('Ghi chú', 35)])
    for ri, (_, r) in enumerate(df_res[df_res['status'] == 'no_kn'].iterrows(), 2):
        _row(ws3, ri, [r['ma'], r['ten_tk'], r['nd'], r['nhap_tk'],
                       'TK có nhập nhưng không có trong BBKN'], F_TK, right_cols=(4,))


def build_excel_kk(wb, df_res, thang_nam):
    F_OK = PatternFill('solid', fgColor='E2EFDA')
    F_L  = PatternFill('solid', fgColor='FFD7D7')
    F_XT = PatternFill('solid', fgColor='DEEBF7')
    F_TK = PatternFill('solid', fgColor='FCE4D6')

    ws = wb.create_sheet(f"DC Kiểm kê {thang_nam.replace('/', '_')}")
    _hdr(ws, [('Mã HPT', 16), ('Tên thuốc (Kiểm kê)', 32), ('Tên thuốc (TK)', 28),
              ('Nồng độ', 22), ('Đơn giá', 12), ('SL Kiểm kê thực tế', 14),
              ('Tồn cuối TK', 13), ('Chênh lệch', 13), ('Trạng thái', 25)])

    df_m  = df_res[df_res['status'] == 'matched']
    df_l  = df_m[df_m['cl'].abs() >= 0.01].sort_values('cl')
    df_k  = df_m[df_m['cl'].abs() < 0.01]
    for ri, (_, r) in enumerate(pd.concat([df_l, df_k], ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl) < 0.01:  fill, st = F_OK, '✅ Khớp'
        elif cl > 0:         fill, st = F_L,  f'⬆️ Thực tế cao hơn {cl:+.0f}'
        else:                fill, st = F_L,  f'⬇️ Thực tế thấp hơn {cl:+.0f}'
        _row(ws, ri, [r['ma'], r['ten_kk'], r['ten_tk'], r['nd'], r['gia'],
                      r['sl_kk'], r['ton_tk'], cl, st],
             fill, right_cols=(5, 6, 7, 8), center_cols=(1, 9))
    ws.freeze_panes = 'A2'

    ws2 = wb.create_sheet("KK - Kiểm kê có TK không")
    _hdr(ws2, [('Tên thuốc', 35), ('Nồng độ', 25), ('Đơn giá', 12), ('SL KK', 12), ('Ghi chú', 35)])
    for ri, (_, r) in enumerate(df_res[df_res['status'] == 'no_tk'].iterrows(), 2):
        _row(ws2, ri, [r['ten_kk'], r['nd'], r['gia'], r['sl_kk'],
                       'KK có nhưng TK không theo dõi tồn'], F_XT, right_cols=(3, 4))

    ws3 = wb.create_sheet("KK - TK có kiểm kê không")
    _hdr(ws3, [('Mã HPT', 16), ('Tên thuốc TK', 35), ('Nồng độ TK', 25), ('Tồn TK', 12), ('Ghi chú', 35)])
    for ri, (_, r) in enumerate(df_res[df_res['status'] == 'no_kk'].iterrows(), 2):
        _row(ws3, ri, [r['ma'], r['ten_tk'], r['nd'], r['ton_tk'],
                       'TK có tồn nhưng không xuất hiện trong KK'], F_TK, right_cols=(4,))


def build_summary_sheet(wb, results_map, thang_nam):
    ws = wb.create_sheet("📊 Tóm tắt", 0)
    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 20

    def wr(ri, k, v, bold=False):
        c1 = ws.cell(row=ri, column=1, value=k)
        c2 = ws.cell(row=ri, column=2, value=v)
        c1.font = Font(name='Times New Roman', bold=bold, size=12 if bold else 11)
        c2.font = Font(name='Times New Roman', size=11)

    ri = 1
    wr(ri, f'BÁO CÁO ĐỐI CHIẾU DƯỢC – {thang_nam}', '', bold=True); ri += 2
    for label, df_r, col_cl, col_status, no_vals, no_labels in results_map:
        wr(ri, f'── {label} ──', '', bold=True); ri += 1
        if df_r is not None and col_cl:
            m = df_r[df_r[col_cl].notna()]
            nk = (m[col_cl].abs() < 0.01).sum()
            nl = (m[col_cl].abs() >= 0.01).sum()
            pct = nk / (nk + nl) * 100 if (nk + nl) > 0 else 0
            wr(ri, '  ✅ Khớp hoàn toàn', f'{nk} dòng'); ri += 1
            wr(ri, '  ⚠️  Chênh lệch', f'{nl} dòng'); ri += 1
            wr(ri, '  📊 Tỷ lệ khớp', f'{pct:.1f}%'); ri += 1
        if df_r is not None and col_status:
            for nv, nl2 in zip(no_vals, no_labels):
                cnt = (df_r[col_status] == nv).sum()
                wr(ri, f'  📋 {nl2}', f'{cnt} dòng'); ri += 1
        ri += 1


def export_all_excel(res_xnt, res_kn, res_kk, thang_nam):
    wb = Workbook()
    wb.remove(wb.active)
    if res_xnt is not None: build_excel_xnt(wb, res_xnt, thang_nam)
    if res_kn  is not None: build_excel_kn(wb, res_kn, thang_nam)
    if res_kk  is not None: build_excel_kk(wb, res_kk, thang_nam)
    rm = []
    if res_xnt is not None:
        rm.append(('Đối chiếu XNT', res_xnt, 'cl', 'method',
                   ['no_tk', 'no_xnt'], ['HPT có – TK không', 'TK có – HPT không']))
    if res_kn  is not None:
        rm.append(('Đối chiếu Kiểm nhập', res_kn, 'cl', 'status',
                   ['no_tk', 'no_kn'], ['BBKN có – TK không', 'TK có – BBKN không']))
    if res_kk  is not None:
        rm.append(('Đối chiếu Kiểm kê', res_kk, 'cl', 'status',
                   ['no_tk', 'no_kk'], ['KK có – TK không', 'TK có – KK không']))
    build_summary_sheet(wb, rm, thang_nam)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  GIAO DIỆN CHÍNH
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="hero">
  <div class="badge">🏥 Bệnh viện Đà Nẵng · Khoa Dược</div>
  <h1>ĐỐI CHIẾU DƯỢC – HPT vs THỐNG KÊ</h1>
  <p class="sub">XNT · Kiểm nhập · Kiểm kê · Phân tách đa mã · Phát hiện chênh lệch</p>
</div>
""", unsafe_allow_html=True)

col_t, col_n = st.columns(2)
with col_t: thang = st.selectbox("Tháng báo cáo", range(1, 13), index=2, format_func=lambda x: f"Tháng {x}")
with col_n: nam   = st.number_input("Năm", min_value=2024, max_value=2030, value=2026)
thang_nam = f"T{thang}/{nam}"

# ── Hiển thị global_map status ───────────────────────────────────────────────
gmap = st.session_state.get('global_map')
if gmap is not None and not gmap.empty:
    n_unique_ma  = gmap['ma'].nunique()
    n_multi      = gmap.groupby(['kten', 'knd', 'kgia']).filter(lambda x: len(x) > 1)['kten'].nunique()
    st.markdown(f"""
    <div class="map-box">
    ✅ <b>Bảng Mã Hàng đã tải:</b> {n_unique_ma} mã HPT duy nhất
    {f'— <b>{n_multi} thuốc</b> có từ 2 mã trở lên (thầu cũ/mới)' if n_multi > 0 else ''}
    <br><small>Bảng mã này tự động áp dụng cho cả Tab Kiểm nhập và Tab Kiểm kê.</small>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="warn-box">
    ⚠️ <b>Chưa có bảng mã hàng.</b> Hãy upload file Nhập/Xuất kho ở Tab <b>Đối chiếu XNT</b> trước
    để hệ thống có thể phân tách các mã thầu cũ/mới khi đối chiếu Kiểm nhập và Kiểm kê.
    </div>
    """, unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

tab_xnt, tab_kn, tab_kk = st.tabs(["📊 Đối chiếu XNT", "📥 Đối chiếu Kiểm nhập", "🔍 Đối chiếu Kiểm kê"])


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  TAB 1 – XNT                                                            ║
# ╚══════════════════════════════════════════════════════════════════════════╝
with tab_xnt:
    with st.expander("📋 Hướng dẫn Đối chiếu XNT", expanded=False):
        st.markdown("""
        **Bước 1:** Upload file Nhập hàng + Xuất kho (có Mã HPT) → hệ thống lưu bảng mã dùng cho tất cả tab.
        **Bước 2:** Upload file XNT thô và XNT Thống kê.
        **Bước 3:** Nhấn **Bắt đầu đối chiếu XNT**.
        """)

    st.markdown('<div class="section-label">📂 File nguồn cho XNT</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        f_nhap = st.file_uploader("📥 File Nhập hàng (có Mã HPT)", type=["xlsx", "xls"],
                                   accept_multiple_files=True, key="xnt_nhap")
    with col2:
        f_xuat = st.file_uploader("📤 File Xuất kho (có Mã HPT)", type=["xlsx", "xls"],
                                   accept_multiple_files=True, key="xnt_xuat")
    col3, col4 = st.columns(2)
    with col3:
        f_xnt = st.file_uploader("📊 File XNT thô HPT (không Mã)", type=["xlsx", "xls"], key="xnt_tho")
    with col4:
        f_tk_xnt = st.file_uploader("📋 File XNT Thống kê (có Mã, tồn chuẩn)", type=["xlsx", "xls"], key="xnt_tk")

    f_confirm = st.file_uploader("✅ File xác nhận mã (tùy chọn)", type=["xlsx"], key="xnt_confirm")

    if st.button("🔍 Bắt đầu đối chiếu XNT", key="btn_xnt",
                  disabled=not ((f_nhap or f_xuat) and f_xnt and f_tk_xnt)):
        with st.spinner("Đang xử lý XNT..."):
            try:
                dfs_nx = []
                for f in (f_nhap or []):
                    f.seek(0)
                    dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))
                for f in (f_xuat or []):
                    f.seek(0)
                    dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))

                # ── Lưu global_map vào session_state ──
                gmap_new = extract_ma_map(dfs_nx)
                st.session_state['global_map'] = gmap_new

                f_xnt.seek(0)
                f_tk_xnt.seek(0)
                df_xnt_raw = pd.read_excel(io.BytesIO(f_xnt.read()), sheet_name=0, header=None)
                df_tk_raw  = pd.read_excel(io.BytesIO(f_tk_xnt.read()), sheet_name=0, header=None)

                confirmed = {}
                if f_confirm:
                    from openpyxl import load_workbook
                    wb_c = load_workbook(io.BytesIO(f_confirm.read()), data_only=True)
                    ws_c = wb_c.active
                    cur_ten = ''
                    for r in range(2, ws_c.max_row + 1):
                        row = [ws_c.cell(row=r, column=c).value for c in range(1, 8)]
                        ten, nd, gia, ma, _, _, chon = row
                        if ma and '→' in str(ma):
                            cur_ten = str(ten or '').strip()
                            continue
                        if chon and str(chon).strip().lower() in ('yes', 'y', 'có', 'co', 'x') and ma:
                            confirmed[cur_ten] = str(ma).strip()

                df_res, err = run_reconciliation_xnt(dfs_nx, df_xnt_raw, df_tk_raw, confirmed)
                if err:
                    st.error(f"❌ {err}"); st.stop()

                st.session_state['xnt_result']    = df_res
                st.session_state['xnt_thang_nam'] = thang_nam
                st.session_state['xnt_done']      = True
                st.rerun()
            except Exception as e:
                st.error(f"❌ Lỗi: {e}"); st.exception(e)

    if st.session_state.get('xnt_done'):
        df_res     = st.session_state['xnt_result']
        tn_s       = st.session_state['xnt_thang_nam']
        df_matched = df_res[df_res['cl'].notna()]
        df_no_tk   = df_res[df_res['method'] == 'no_tk']
        df_no_xnt  = df_res[df_res['method'] == 'no_xnt']
        n_khop = (df_matched['cl'].abs() < 0.01).sum()
        n_lech = (df_matched['cl'].abs() >= 0.01).sum()

        st.markdown(f"""
        <div class="ok-box"><div style="font-size:2rem">✅</div>
          <h3>Đối chiếu XNT hoàn tất – {tn_s}</h3></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="num" style="color:#166534">{n_khop}</div>
            <div class="lbl">✅ Khớp hoàn toàn</div></div>
          <div class="stat-card"><div class="num" style="color:#dc2626">{n_lech}</div>
            <div class="lbl">⚠️ Chênh lệch</div></div>
          <div class="stat-card"><div class="num" style="color:#2563a8">{len(df_no_tk)}</div>
            <div class="lbl">📋 XNT có / TK không</div></div>
          <div class="stat-card"><div class="num" style="color:#d97706">{len(df_no_xnt)}</div>
            <div class="lbl">📋 TK có / XNT không</div></div>
        </div>""", unsafe_allow_html=True)

        df_lech = df_matched[df_matched['cl'].abs() >= 0.01].sort_values('cl')
        if len(df_lech):
            st.markdown(f"**⚠️ {n_lech} dòng chênh lệch:**")
            st.dataframe(df_lech[['ma', 'ten', 'nd', 'ton_xnt', 'ton_tk', 'cl', 'method']].rename(
                columns={'ma': 'Mã', 'ten': 'Tên thuốc', 'nd': 'Nồng độ',
                          'ton_xnt': 'Tồn XNT', 'ton_tk': 'Tồn TK',
                          'cl': 'Chênh lệch', 'method': 'PP'}),
                use_container_width=True, hide_index=True)


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  TAB 2 – KIỂM NHẬP                                                      ║
# ╚══════════════════════════════════════════════════════════════════════════╝
with tab_kn:
    with st.expander("📋 Hướng dẫn Đối chiếu Kiểm nhập", expanded=False):
        st.markdown("""
        **Mục đích:** So sánh **số lượng nhập** trên BBKN với **số nhập chuẩn** từ file Thống kê.

        **Lưu ý quan trọng (v3):**
        - App sẽ hiển thị **tách biệt từng mã hàng** (thầu cũ / thầu mới) thay vì gộp chung.
        - Cần upload file **Nhập/Xuất kho** ở Tab XNT trước để có bảng mã.
        - Nếu BBKN có 2 dòng cùng tên+giá → App khớp từng dòng với từng mã TK.
        """)

    if gmap is None or gmap.empty:
        st.markdown("""
        <div class="warn-box">⚠️ Chưa có bảng mã hàng. Hãy upload và xử lý Tab XNT trước.</div>
        """, unsafe_allow_html=True)

    col_kn1, col_kn2 = st.columns(2)
    with col_kn1:
        f_bbkn = st.file_uploader("📄 File Biên bản kiểm nhập (BBKN)", type=["xlsx", "xls"], key="kn_bbkn")
    with col_kn2:
        f_tk_kn = st.file_uploader("📋 File XNT Thống kê (số nhập chuẩn)", type=["xlsx", "xls"], key="kn_tk")

    sl_col_kn_idx = 9
    if f_bbkn:
        try:
            f_bbkn.seek(0)
            df_prev = pd.read_excel(io.BytesIO(f_bbkn.read()), sheet_name=0, header=None, nrows=15)
            f_bbkn.seek(0)
            hdr_row = df_prev.iloc[7] if len(df_prev) > 7 else df_prev.iloc[0]
            col_opts = {f"Cột {i} | {str(df_prev.iloc[9, i] if len(df_prev) > 9 else '')[:30]}": i
                        for i in range(len(df_prev.columns))}
            default_lbl = next((k for k, v in col_opts.items() if v == 9), list(col_opts.keys())[0])
            chosen = st.selectbox("📌 Chọn cột Số lượng nhập trong BBKN:", list(col_opts.keys()),
                                   index=list(col_opts.keys()).index(default_lbl), key="kn_sl_col")
            sl_col_kn_idx = col_opts[chosen]
        except:
            pass

    if st.button("🔍 Bắt đầu đối chiếu Kiểm nhập", key="btn_kn", disabled=not (f_bbkn and f_tk_kn)):
        with st.spinner("Đang xử lý Kiểm nhập..."):
            try:
                f_bbkn.seek(0); f_tk_kn.seek(0)
                df_bbkn_raw = pd.read_excel(io.BytesIO(f_bbkn.read()), sheet_name=0, header=None)
                df_tk_raw   = pd.read_excel(io.BytesIO(f_tk_kn.read()), sheet_name=0, header=None)
                gmap_cur    = st.session_state.get('global_map', pd.DataFrame())
                df_res_kn, err = run_reconciliation_kn(df_bbkn_raw, df_tk_raw, gmap_cur, sl_col_kn_idx)
                if err:
                    st.error(f"❌ {err}"); st.stop()
                st.session_state['kn_result']    = df_res_kn
                st.session_state['kn_thang_nam'] = thang_nam
                st.session_state['kn_done']      = True
            except Exception as e:
                st.error(f"❌ Lỗi: {e}"); st.exception(e)

    if st.session_state.get('kn_done'):
        df_kn    = st.session_state['kn_result']
        tn_kn    = st.session_state['kn_thang_nam']
        matched  = df_kn[df_kn['status'] == 'matched']
        no_tk    = df_kn[df_kn['status'] == 'no_tk']
        no_kn    = df_kn[df_kn['status'] == 'no_kn']
        n_khop   = (matched['cl'].abs() < 0.01).sum()
        n_lech   = (matched['cl'].abs() >= 0.01).sum()

        st.markdown(f"""
        <div class="ok-box"><div style="font-size:2rem">✅</div>
          <h3>Đối chiếu Kiểm nhập hoàn tất – {tn_kn}</h3></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="num" style="color:#166534">{n_khop}</div>
            <div class="lbl">✅ Khớp hoàn toàn</div></div>
          <div class="stat-card"><div class="num" style="color:#dc2626">{n_lech}</div>
            <div class="lbl">⚠️ Chênh lệch</div></div>
          <div class="stat-card"><div class="num" style="color:#2563a8">{len(no_tk)}</div>
            <div class="lbl">📋 BBKN có / TK không</div></div>
          <div class="stat-card"><div class="num" style="color:#d97706">{len(no_kn)}</div>
            <div class="lbl">📋 TK có / BBKN không</div></div>
        </div>""", unsafe_allow_html=True)

        df_lech_kn = matched[matched['cl'].abs() >= 0.01].sort_values('cl')
        if len(df_lech_kn):
            st.markdown(f"**⚠️ {n_lech} dòng chênh lệch (theo từng mã):**")
            st.dataframe(df_lech_kn[['ma', 'ten_kn', 'nd', 'nhap_kn', 'nhap_tk', 'cl']].rename(columns={
                'ma': 'Mã HPT', 'ten_kn': 'Tên thuốc (BBKN)', 'nd': 'Nồng độ',
                'nhap_kn': 'Nhập BBKN', 'nhap_tk': 'Nhập TK', 'cl': 'Chênh lệch'}),
                use_container_width=True, hide_index=True)

        if len(no_kn):
            st.markdown(f"**📋 {len(no_kn)} mã hàng TK có nhập / BBKN không có:**")
            st.dataframe(no_kn[['ma', 'ten_tk', 'nd', 'nhap_tk']].rename(columns={
                'ma': 'Mã HPT', 'ten_tk': 'Tên thuốc TK', 'nd': 'Nồng độ', 'nhap_tk': 'Nhập TK'}),
                use_container_width=True, hide_index=True)

        if len(no_tk):
            st.markdown(f"**📋 {len(no_tk)} dòng BBKN có / TK không ghi nhận:**")
            st.dataframe(no_tk[['ten_kn', 'nd', 'gia', 'nhap_kn']].rename(columns={
                'ten_kn': 'Tên thuốc', 'nd': 'Nồng độ', 'gia': 'Đơn giá', 'nhap_kn': 'Nhập BBKN'}),
                use_container_width=True, hide_index=True)


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  TAB 3 – KIỂM KÊ                                                        ║
# ╚══════════════════════════════════════════════════════════════════════════╝
with tab_kk:
    with st.expander("📋 Hướng dẫn Đối chiếu Kiểm kê", expanded=False):
        st.markdown("""
        **Mục đích:** So sánh **số lượng kiểm kê thực tế** với **Tồn cuối** trong file XNT Thống kê.

        **Lưu ý quan trọng (v3):**
        - BBKK không có mã, nhưng vì BBKK đã liệt kê **từng dòng cho từng mã thầu riêng**
          (ví dụ: Nolvadex-D tồn 9 và 1407 là hai dòng riêng biệt),
          App sẽ ghép từng dòng BBKK với từng mã TK tương ứng theo số lượng gần nhất.
        - Cần upload file Nhập/Xuất kho ở Tab XNT trước để hệ thống biết có bao nhiêu mã.
        """)

    if gmap is None or gmap.empty:
        st.markdown("""
        <div class="warn-box">⚠️ Chưa có bảng mã hàng. Hãy upload và xử lý Tab XNT trước.</div>
        """, unsafe_allow_html=True)

    col_kk1, col_kk2 = st.columns(2)
    with col_kk1:
        f_bbkk = st.file_uploader("📄 File Biên bản kiểm kê (thực tế)", type=["xlsx", "xls"], key="kk_bbkk")
    with col_kk2:
        f_tk_kk = st.file_uploader("📋 File XNT Thống kê (tồn cuối chuẩn)", type=["xlsx", "xls"], key="kk_tk")

    sl_col_kk_idx = 8
    if f_bbkk:
        try:
            f_bbkk.seek(0)
            df_prev_kk = pd.read_excel(io.BytesIO(f_bbkk.read()), sheet_name=0, header=None, nrows=12)
            f_bbkk.seek(0)
            col_opts_kk = {f"Cột {i} | {str(df_prev_kk.iloc[9, i] if len(df_prev_kk) > 9 else '')[:30]}": i
                           for i in range(len(df_prev_kk.columns))}
            default_kk = next((k for k, v in col_opts_kk.items() if v == 8), list(col_opts_kk.keys())[0])
            chosen_kk = st.selectbox("📌 Chọn cột Số lượng kiểm kê thực tế:",
                                      list(col_opts_kk.keys()),
                                      index=list(col_opts_kk.keys()).index(default_kk), key="kk_sl_col")
            sl_col_kk_idx = col_opts_kk[chosen_kk]
        except:
            pass

    if st.button("🔍 Bắt đầu đối chiếu Kiểm kê", key="btn_kk", disabled=not (f_bbkk and f_tk_kk)):
        with st.spinner("Đang xử lý Kiểm kê..."):
            try:
                f_bbkk.seek(0); f_tk_kk.seek(0)
                df_bbkk_raw = pd.read_excel(io.BytesIO(f_bbkk.read()), sheet_name=0, header=None)
                df_tk_raw   = pd.read_excel(io.BytesIO(f_tk_kk.read()), sheet_name=0, header=None)
                gmap_cur    = st.session_state.get('global_map', pd.DataFrame())
                df_res_kk, err = run_reconciliation_kk(df_bbkk_raw, df_tk_raw, gmap_cur, sl_col_kk_idx)
                if err:
                    st.error(f"❌ {err}"); st.stop()
                st.session_state['kk_result']    = df_res_kk
                st.session_state['kk_thang_nam'] = thang_nam
                st.session_state['kk_done']      = True
            except Exception as e:
                st.error(f"❌ Lỗi: {e}"); st.exception(e)

    if st.session_state.get('kk_done'):
        df_kk     = st.session_state['kk_result']
        tn_kk     = st.session_state['kk_thang_nam']
        matched_kk = df_kk[df_kk['status'] == 'matched']
        no_tk_kk   = df_kk[df_kk['status'] == 'no_tk']
        no_kk_kk   = df_kk[df_kk['status'] == 'no_kk']
        n_khop_kk  = (matched_kk['cl'].abs() < 0.01).sum()
        n_lech_kk  = (matched_kk['cl'].abs() >= 0.01).sum()

        st.markdown(f"""
        <div class="ok-box"><div style="font-size:2rem">✅</div>
          <h3>Đối chiếu Kiểm kê hoàn tất – {tn_kk}</h3></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="num" style="color:#166534">{n_khop_kk}</div>
            <div class="lbl">✅ Khớp hoàn toàn</div></div>
          <div class="stat-card"><div class="num" style="color:#dc2626">{n_lech_kk}</div>
            <div class="lbl">⚠️ Chênh lệch</div></div>
          <div class="stat-card"><div class="num" style="color:#2563a8">{len(no_tk_kk)}</div>
            <div class="lbl">📋 KK có / TK không</div></div>
          <div class="stat-card"><div class="num" style="color:#d97706">{len(no_kk_kk)}</div>
            <div class="lbl">📋 TK có / KK không</div></div>
        </div>""", unsafe_allow_html=True)

        df_lech_kk = matched_kk[matched_kk['cl'].abs() >= 0.01].sort_values('cl')
        if len(df_lech_kk):
            st.markdown(f"**⚠️ {n_lech_kk} dòng chênh lệch (theo từng mã):**")
            st.dataframe(df_lech_kk[['ma', 'ten_kk', 'nd', 'sl_kk', 'ton_tk', 'cl']].rename(columns={
                'ma': 'Mã HPT', 'ten_kk': 'Tên thuốc (KK)', 'nd': 'Nồng độ',
                'sl_kk': 'SL Thực tế', 'ton_tk': 'Tồn TK', 'cl': 'Chênh lệch'}),
                use_container_width=True, hide_index=True)

        if len(no_kk_kk):
            st.markdown(f"**📋 {len(no_kk_kk)} mã hàng TK có tồn / KK không có:**")
            st.dataframe(no_kk_kk[['ma', 'ten_tk', 'nd', 'ton_tk']].rename(columns={
                'ma': 'Mã HPT', 'ten_tk': 'Tên thuốc TK', 'nd': 'Nồng độ', 'ton_tk': 'Tồn TK'}),
                use_container_width=True, hide_index=True)

        if len(no_tk_kk):
            st.markdown(f"**📋 {len(no_tk_kk)} dòng KK có / TK không theo dõi tồn:**")
            st.dataframe(no_tk_kk[['ten_kk', 'nd', 'gia', 'sl_kk']].rename(columns={
                'ten_kk': 'Tên thuốc', 'nd': 'Nồng độ', 'gia': 'Đơn giá', 'sl_kk': 'SL Thực tế'}),
                use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  NÚT TẢI FILE EXCEL GỘP
# ══════════════════════════════════════════════════════════════════════════════
has_any = any(st.session_state.get(k) for k in ['xnt_done', 'kn_done', 'kk_done'])
if has_any:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### 📥 Tải kết quả tổng hợp")
    st.markdown("""
    <div class="info-box">
    File Excel xuất ra gộp <b>tất cả kết quả</b> (XNT + Kiểm nhập + Kiểm kê) vào các Sheet riêng biệt,
    kèm Sheet <b>📊 Tóm tắt</b>. Màu <span style="background:#E2EFDA;padding:2px 6px;border-radius:3px">🟢 Xanh</span> = Khớp |
    <span style="background:#FFD7D7;padding:2px 6px;border-radius:3px">🔴 Đỏ</span> = Lệch.
    Kết quả Kiểm nhập và Kiểm kê được <b>tách theo từng mã hàng</b>.
    </div>
    """, unsafe_allow_html=True)

    res_xnt = st.session_state.get('xnt_result')
    res_kn  = st.session_state.get('kn_result')
    res_kk  = st.session_state.get('kk_result')
    tn      = (st.session_state.get('xnt_thang_nam') or
               st.session_state.get('kn_thang_nam') or
               st.session_state.get('kk_thang_nam') or thang_nam)

    excel_bytes = export_all_excel(res_xnt, res_kn, res_kk, tn)
    st.download_button(
        label=f"⬇️  Tải Kết Quả Tổng Hợp {tn} (.xlsx)",
        data=excel_bytes,
        file_name=f"doi_chieu_tong_hop_{tn.replace('/', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
