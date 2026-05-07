"""
HỆ THỐNG TỰ ĐỘNG HÓA BIÊN BẢN DƯỢC – Bệnh viện Đà Nẵng
Hỗ trợ: BBKN · BBKK · XNT · Đối Chiếu Dược (XNT, Kiểm nhập, Kiểm kê)
v7 – Thêm module Biên Bản Kiểm Kê (BBKK) + chọn tháng/năm báo cáo tự động
"""

import io, math, copy, datetime, re, warnings
import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Biên Bản Dược – BV Đà Nẵng",
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
.hero h1{font-size:1.55rem;font-weight:700;margin:0 0 6px;line-height:1.3;}
.hero .sub{font-size:.88rem;font-weight:300;opacity:.85;margin:0;}
.hero .badge{display:inline-block;background:rgba(255,255,255,.18);border-radius:20px;
  padding:3px 12px;font-size:.75rem;font-weight:600;letter-spacing:1px;
  margin-bottom:12px;text-transform:uppercase;}
.tab-desc{background:#eff6ff;border-left:4px solid #2563a8;border-radius:0 10px 10px 0;
  padding:12px 16px;margin:12px 0 18px;font-size:.86rem;color:#1e3a5f;line-height:1.6;}
.stat-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:16px 0;}
.stat-grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0;}
.stat-card{background:white;border:1px solid #e2e8f0;border-radius:12px;
  padding:16px 12px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.06);}
.stat-card .num{font-size:1.7rem;font-weight:700;color:#1a3a5c;line-height:1;}
.stat-card .lbl{font-size:.75rem;color:#64748b;margin-top:4px;}
.ok-box{background:#f0fdf4;border:1.5px solid #86efac;border-radius:12px;
  padding:18px 20px;margin:16px 0;text-align:center;}
.ok-box .icon{font-size:2rem;}.ok-box h3{color:#166534;margin:6px 0 4px;font-size:1rem;}
.ok-box p{color:#15803d;font-size:.84rem;margin:0;}
.note{background:#fff7ed;border-left:4px solid #f59e0b;border-radius:0 10px 10px 0;
  padding:10px 14px;font-size:.82rem;color:#92400e;margin-top:14px;line-height:1.6;}
.warn-box{background:#fff7ed;border-left:4px solid #f59e0b;border-radius:0 10px 10px 0;
  padding:11px 15px;margin:10px 0;font-size:.84rem;color:#92400e;line-height:1.6;}
.info-box{background:#eff6ff;border-left:4px solid #2563a8;border-radius:0 10px 10px 0;
  padding:11px 15px;margin:10px 0 16px;font-size:.85rem;color:#1e3a5f;line-height:1.7;}
.map-box{background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;
  padding:10px 16px;margin:10px 0;font-size:.85rem;color:#166534;}
.upload-section{background:#f8faff;border:1.5px solid #c7d9f5;border-radius:14px;
  padding:18px 20px;margin-bottom:18px;}
.upload-section h4{color:#1a3a5c;font-size:.82rem;font-weight:700;letter-spacing:1.2px;
  text-transform:uppercase;margin:0 0 14px;}
.stButton>button{background:linear-gradient(135deg,#1a3a5c,#2563a8)!important;
  color:white!important;font-weight:600!important;font-size:.95rem!important;
  border:none!important;border-radius:10px!important;padding:13px 0!important;
  width:100%!important;box-shadow:0 4px 14px rgba(37,99,168,.3)!important;}
[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#166534,#16a34a)!important;
  color:white!important;font-weight:700!important;font-size:1rem!important;
  border:none!important;border-radius:10px!important;padding:15px 0!important;
  width:100%!important;box-shadow:0 4px 14px rgba(22,163,74,.3)!important;}
[data-testid="stFileUploader"]{border:2px dashed #2563a8!important;
  border-radius:12px!important;background:#f0f6ff!important;}
hr{border:none;border-top:1px solid #e2e8f0;margin:20px 0;}
#MainMenu,footer{visibility:hidden;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  SHARED HELPERS (dùng chung BBKN + XNT)
# ══════════════════════════════════════════════════════════════════════════════
SKIP_KW = ['Tổng cộng','Hội đồng','Trưởng','Trang','Đã kiểm nhập',
           'Ông/bà','kiểm nhập những','Trang 1']
THIN = Side(style='thin')
MED  = Side(style='medium')
NO_FILL = PatternFill(fill_type=None)

def b_thin(): return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
def b_med():  return Border(left=THIN, right=THIN, top=MED,  bottom=MED)

def safe_set(cell, **kwargs):
    for k, v in kwargs.items():
        try: setattr(cell, k, v)
        except AttributeError: pass

def is_co_row(v0, row):
    return (isinstance(v0, str) and pd.isna(row[1]) and pd.isna(row[2])
            and not any(kw in str(v0) for kw in SKIP_KW))

def is_drug_row(v0, col2):
    try: int(str(v0).strip())
    except: return False
    return not pd.isna(col2) and isinstance(col2, str) and not col2.strip().isdigit()

def parse_companies(raw_df, qty_col):
    result, cur, rows, skipped = [], None, [], 0
    for _, row in raw_df.iterrows():
        v0 = row[0]
        if is_co_row(v0, row):
            if cur and rows: result.append((cur, rows))
            cur, rows = str(v0).strip(), []
        elif is_drug_row(v0, row[2]):
            try:    qty = float(row[qty_col]) if not pd.isna(row[qty_col]) else 0
            except: qty = 0
            if qty != 0: rows.append(row)
            else: skipped += 1
    if cur and rows: result.append((cur, rows))
    return result, {'companies': len(result),
                    'drugs': sum(len(d) for _, d in result),
                    'skipped': skipped}

def gs(ws, r, c):
    cl = ws.cell(row=r, column=c)
    return {k: copy.copy(getattr(cl, k))
            for k in ('font','border','alignment','fill','number_format')}

def ap(cell, s):
    for k, v in s.items():
        try: setattr(cell, k, copy.copy(v))
        except AttributeError: pass


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE BBKN
# ══════════════════════════════════════════════════════════════════════════════
BBKN_W = {1:5.5,2:10,3:38,4:17,5:7.5,6:11,7:24,8:11,9:12,10:9.5,11:14,12:8}
BBKN_A = {1:('center','center'),2:('center','center'),3:('left','center'),
          4:('left','center'),5:('center','center'),6:('center','center'),
          7:('left','center'),8:('center','center'),9:('right','center'),
          10:('right','center'),11:('right','center'),12:('center','center')}
BBKN_WRAP = {3,4,7}
BBKN_NUM  = {9,10,11}

def bbkn_h(ws, r):
    ml = 1
    for c in BBKN_WRAP:
        v = ws.cell(row=r, column=c).value
        if not v or not isinstance(v, str): continue
        cw = max(BBKN_W.get(c,15)*1.1, 1)
        ml = max(ml, sum(max(1,math.ceil(len(ln)/cw)) for ln in v.split('\n')))
    return max(22, min(ml*15.6+4, 120))

def build_bbkn(tmpl_bytes, companies):
    wb = load_workbook(io.BytesIO(tmpl_bytes))
    ws = wb.active
    cs = {c: gs(ws,15,c) for c in range(1,13)}
    ds = {c: gs(ws,16,c) for c in range(1,13)}
    tks = gs(ws,213,11)

    fs = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 'HỘI ĐỒNG KIỂM NHẬP': fs = cell.row; break
        if fs: break
    if not fs: fs = 215

    DS = 15
    need = sum(1+len(d) for _,d in companies)+1
    ins = (DS+need-1) - fs + 1
    if ins > 0: ws.insert_rows(fs, ins); fs += ins

    for m in [str(mr) for mr in ws.merged_cells.ranges if DS <= mr.min_row < fs]:
        ws.merged_cells.remove(m)
    for r in range(DS, fs):
        for c in range(1,13):
            try: ws.cell(row=r,column=c).value = None
            except: pass

    def wco(rn, name):
        cl = ws.cell(row=rn, column=1, value=name); ap(cl, cs[1])
        cl.font = Font(name='Times New Roman', bold=True, size=12); cl.fill = NO_FILL
        for c in range(2,13):
            cc = ws.cell(row=rn, column=c); ap(cc, cs[c]); cc.fill = NO_FILL
        ws.row_dimensions[rn].height = 20

    def wdr(rn, stt, dr):
        cols = [
            (1, stt,                                               'center', False, None),
            (2, ''if pd.isna(dr[1])else str(dr[1]).strip(),       'center', False, None),
            (3, ''if pd.isna(dr[2])else str(dr[2]).strip(),       'left',   True,  None),
            (4, ''if pd.isna(dr[3])else str(dr[3]).strip(),       'left',   True,  None),
            (5, ''if pd.isna(dr[4])else str(dr[4]).strip(),       'center', False, None),
            (6, ''if pd.isna(dr[5])else str(dr[5]).strip(),       'center', False, None),
            (7, ''if pd.isna(dr[6])else str(dr[6]).strip(),       'left',   True,  None),
            (8, dr[7] if isinstance(dr[7],datetime.datetime)
                else(''if pd.isna(dr[7])else dr[7]),              'center', False, 'DD/MM/YYYY'),
            (9, dr[8] if not pd.isna(dr[8]) else 0,              'right',  False, '#,##0'),
            (10,int(dr[9]) if not pd.isna(dr[9]) else 0,         'right',  False, '#,##0'),
        ]
        for col,val,ha,wrap,fmt in cols:
            cl = ws.cell(row=rn,column=col,value=val); ap(cl,ds[col])
            cl.font = Font(name='Times New Roman',size=12)
            cl.alignment = Alignment(horizontal=ha,vertical='center',wrap_text=wrap)
            if fmt and val!='': cl.number_format = fmt
        ck = ws.cell(row=rn,column=11,value=f'=I{rn}*J{rn}'); ap(ck,ds[11])
        ck.font=Font(name='Times New Roman',size=12)
        ck.alignment=Alignment(horizontal='right',vertical='center')
        ck.number_format='#,##0'
        cl12=ws.cell(row=rn,column=12,value=''); ap(cl12,ds[12])

    cr = DS; drn = []
    for name,drugs in companies:
        wco(cr,name); cr+=1
        for i,dr in enumerate(drugs,1): wdr(cr,i,dr); drn.append(cr); cr+=1

    tr = cr
    lbl=ws.cell(row=tr,column=1,value='Tổng cộng: ')
    lbl.font=Font(name='Times New Roman',bold=True,size=12)
    lbl.alignment=Alignment(horizontal='left',vertical='center'); lbl.border=b_med()
    for c in range(2,11):
        try: ws.cell(row=tr,column=c).border=b_med()
        except: pass
    ck=ws.cell(row=tr,column=11,value=f'=SUM({",".join(f"K{r}"for r in drn)})')
    ap(ck,tks); ck.font=Font(name='Times New Roman',bold=True,size=12)
    ck.alignment=Alignment(horizontal='right',vertical='center')
    ck.number_format='#,##0'; ck.border=b_med()
    ws.row_dimensions[tr].height=22

    ws.cell(row=13,column=3).value='Tên thuốc'
    for col in range(1,13):
        for r in (13,14):
            safe_set(ws.cell(row=r,column=col),fill=NO_FILL,
                     font=Font(name='Times New Roman',bold=True,size=12),
                     border=b_med(),
                     alignment=Alignment(horizontal='center',vertical='center',wrap_text=True))
    ws.row_dimensions[13].height=42; ws.row_dimensions[14].height=18

    for r in range(DS,tr+1):
        av=ws.cell(row=r,column=1).value; cv=ws.cell(row=r,column=3).value
        is_co=isinstance(av,str) and not str(av).strip().lstrip('-').isdigit() and not cv
        if r==tr: pass
        elif is_co:
            ws.row_dimensions[r].height=20
            for col in range(1,13):
                safe_set(ws.cell(row=r,column=col),fill=NO_FILL,border=b_thin(),
                         font=Font(name='Times New Roman',bold=True,size=12),
                         alignment=Alignment(horizontal='left',vertical='center'))
        else:
            ws.row_dimensions[r].height=bbkn_h(ws,r)
            for col in range(1,13):
                cl=ws.cell(row=r,column=col)
                ha,va=BBKN_A.get(col,('left','center'))
                safe_set(cl,fill=NO_FILL,border=b_thin(),
                         font=Font(name='Times New Roman',size=12),
                         alignment=Alignment(horizontal=ha,vertical=va,wrap_text=col in BBKN_WRAP))
                if col in BBKN_NUM and cl.value is not None: cl.number_format='#,##0'
                if col==8 and isinstance(cl.value,datetime.datetime): cl.number_format='DD/MM/YYYY'

    for col,w in BBKN_W.items(): ws.column_dimensions[get_column_letter(col)].width=w
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    for a,v in [('left',.4),('right',.4),('top',.5),('bottom',.5),('header',.2),('footer',.2)]:
        setattr(ws.page_margins,a,v)
    ws.print_title_rows='1:14'; ws.freeze_panes=ws.cell(row=DS,column=1)

    out=io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE XNT
# ══════════════════════════════════════════════════════════════════════════════
XNT_W = {1:5,2:28,3:16,4:7,5:10,6:30,7:11,8:12,9:9,10:9,11:9,12:9,13:14,14:8}
XNT_A = {1:('center','center'),2:('left','center'),3:('left','center'),
         4:('center','center'),5:('center','center'),6:('left','center'),
         7:('center','center'),8:('right','center'),9:('right','center'),
         10:('right','center'),11:('right','center'),12:('right','center'),
         13:('right','center'),14:('center','center')}
XNT_WRAP = {2,3,6}
XNT_NUM  = {8,9,10,11,12,13}

def xnt_h(ws,r):
    ml=1
    for c in XNT_WRAP:
        v=ws.cell(row=r,column=c).value
        if not v or not isinstance(v,str): continue
        cw=max(XNT_W.get(c,15)*1.1,1)
        ml=max(ml,sum(max(1,math.ceil(len(ln)/cw)) for ln in v.split('\n')))
    return max(20,min(ml*14.3+4,120))

def build_xnt(tmpl_bytes, companies):
    wb=load_workbook(io.BytesIO(tmpl_bytes))
    ws=wb.active
    cs={c:gs(ws,12,c) for c in range(1,15)}
    ds={c:gs(ws,13,c) for c in range(1,15)}

    fs=None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value=='Tổng cộng': fs=cell.row; break
        if fs: break
    if not fs: fs=279

    ws.delete_rows(fs, 1)

    DS=12
    need=sum(1+len(d) for _,d in companies)+1
    data_end=DS+need-1
    ins=data_end-fs+1
    if ins>0: ws.insert_rows(fs,ins); fs+=ins

    for m in [str(mr) for mr in ws.merged_cells.ranges if DS<=mr.min_row<fs]:
        ws.merged_cells.remove(m)
    for r in range(DS,fs):
        for c in range(1,15):
            try: ws.cell(row=r,column=c).value=None
            except: pass

    def wco(rn,name):
        cl=ws.cell(row=rn,column=1,value=name); ap(cl,cs[1])
        cl.font=Font(name='Times New Roman',bold=True,size=11); cl.fill=NO_FILL
        for c in range(2,15):
            cc=ws.cell(row=rn,column=c); ap(cc,cs[c]); cc.fill=NO_FILL
        ws.row_dimensions[rn].height=18

    def wdr(rn,stt,dr):
        cols=[
            (1, stt,                                                    'center',False,None),
            (2, ''if pd.isna(dr[2])else str(dr[2]).strip(),            'left',  True, None),
            (3, ''if pd.isna(dr[3])else str(dr[3]).strip(),            'left',  True, None),
            (4, ''if pd.isna(dr[4])else str(dr[4]).strip(),            'center',False,None),
            (5, ''if pd.isna(dr[5])else str(dr[5]).strip(),            'center',False,None),
            (6, ''if pd.isna(dr[6])else str(dr[6]).strip(),            'left',  True, None),
            (7, dr[7] if isinstance(dr[7],datetime.datetime)
                else(''if pd.isna(dr[7])else dr[7]),                   'center',False,'DD/MM/YYYY'),
            (8, dr[8] if not pd.isna(dr[8])else 0,                    'right', False,'#,##0'),
            (9, dr[9] if not pd.isna(dr[9])else 0,                    'right', False,'#,##0'),
            (10,dr[10]if not pd.isna(dr[10])else 0,                   'right', False,'#,##0'),
            (11,dr[11]if not pd.isna(dr[11])else 0,                   'right', False,'#,##0'),
            (12,dr[12]if not pd.isna(dr[12])else 0,                   'right', False,'#,##0'),
        ]
        for col,val,ha,wrap,fmt in cols:
            cl=ws.cell(row=rn,column=col,value=val); ap(cl,ds[col])
            cl.font=Font(name='Times New Roman',size=11)
            cl.alignment=Alignment(horizontal=ha,vertical='center',wrap_text=wrap)
            if fmt and val!='': cl.number_format=fmt
        cm=ws.cell(row=rn,column=13,value=f'=H{rn}*L{rn}'); ap(cm,ds[13])
        cm.font=Font(name='Times New Roman',size=11)
        cm.alignment=Alignment(horizontal='right',vertical='center')
        cm.number_format='#,##0'
        cn=ws.cell(row=rn,column=14,value=''); ap(cn,ds[14])

    cr=DS; drn=[]
    for name,drugs in companies:
        wco(cr,name); cr+=1
        for i,dr in enumerate(drugs,1): wdr(cr,i,dr); drn.append(cr); cr+=1

    tr=cr
    total_val = sum(
        ws.cell(row=r,column=8).value * ws.cell(row=r,column=12).value
        for r in drn
        if isinstance(ws.cell(row=r,column=8).value,(int,float))
        and isinstance(ws.cell(row=r,column=12).value,(int,float))
    )
    lbl=ws.cell(row=tr,column=1,value='Tổng cộng')
    lbl.font=Font(name='Times New Roman',bold=True,size=11)
    lbl.alignment=Alignment(horizontal='left',vertical='center'); lbl.border=b_med()
    for c in range(2,13):
        try: ws.cell(row=tr,column=c).border=b_med()
        except: pass
    cm=ws.cell(row=tr,column=13,value=round(total_val))
    cm.font=Font(name='Times New Roman',bold=True,size=11)
    cm.alignment=Alignment(horizontal='right',vertical='center')
    cm.number_format='#,##0'; cm.border=b_med()
    ws.row_dimensions[tr].height=20

    for r in range(tr+1, ws.max_row+1):
        ws.row_dimensions[r].height=22
        for col in range(1,15):
            cl=ws.cell(row=r,column=col)
            if cl.value and isinstance(cl.value,str):
                cl.value=cl.value.strip()
                cl.alignment=Alignment(horizontal='center',vertical='center',wrap_text=False)
                cl.font=Font(name='Times New Roman',size=11)

    for col in range(1,15):
        for r in range(9,12):
            safe_set(ws.cell(row=r,column=col),fill=NO_FILL,
                     font=Font(name='Times New Roman',bold=True,size=11),
                     border=b_med(),
                     alignment=Alignment(horizontal='center',vertical='center',wrap_text=True))
        ws.row_dimensions[r].height=20

    for r in range(DS,tr+1):
        av=ws.cell(row=r,column=1).value; bv=ws.cell(row=r,column=2).value
        is_co=isinstance(av,str) and not str(av).strip().lstrip('-').isdigit() and not bv
        if r==tr: pass
        elif is_co:
            ws.row_dimensions[r].height=18
            for col in range(1,15):
                safe_set(ws.cell(row=r,column=col),fill=NO_FILL,border=b_thin(),
                         font=Font(name='Times New Roman',bold=True,size=11),
                         alignment=Alignment(horizontal='left',vertical='center'))
        else:
            ws.row_dimensions[r].height=xnt_h(ws,r)
            for col in range(1,15):
                cl=ws.cell(row=r,column=col)
                ha,va=XNT_A.get(col,('left','center'))
                safe_set(cl,fill=NO_FILL,border=b_thin(),
                         font=Font(name='Times New Roman',size=11),
                         alignment=Alignment(horizontal=ha,vertical=va,wrap_text=col in XNT_WRAP))
                if col in XNT_NUM and cl.value is not None and cl.value!='': cl.number_format='#,##0'
                if col==7 and isinstance(cl.value,datetime.datetime): cl.number_format='DD/MM/YYYY'

    for col,w in XNT_W.items(): ws.column_dimensions[get_column_letter(col)].width=w
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    for a,v in [('left',.35),('right',.35),('top',.5),('bottom',.5),('header',.2),('footer',.2)]:
        setattr(ws.page_margins,a,v)
    ws.print_title_rows='1:11'; ws.freeze_panes=ws.cell(row=DS,column=1)

    out=io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE ĐỐI CHIẾU DƯỢC – toàn bộ logic từ app_doi_chieu_v5
# ══════════════════════════════════════════════════════════════════════════════

def dc_norm(s):
    if not isinstance(s, str): return ""
    s = s.strip().lower()
    s = re.sub(r'_x000a_', ' ', s); s = re.sub(r'\n', ' ', s)
    s = re.sub(r'[#]', '', s)
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    s = re.sub(r'(\d)\s+(mg|ml|mcg|μg|g|iu|%|meq|l)', r'\1\2', s)
    s = re.sub(r'\s*\+\s*', '+', s)
    return re.sub(r'\s+', ' ', s).strip()

def dc_is_drug(s):
    if not isinstance(s, str): return False
    s = s.strip()
    if len(s) < 2: return False
    try: float(s); return False
    except: return True

def dc_safe_float(v):
    try: return float(v)
    except: return 0.0

def dc_parse_tk(df_raw):
    rows = []
    for _, row in df_raw.iloc[5:].iterrows():
        ma  = str(row[4]).strip() if not pd.isna(row[4]) else ''
        ten = str(row[5]).strip() if not pd.isna(row[5]) else ''
        if not ma or not ten: continue
        gia = dc_safe_float(row[11])
        rows.append({
            'ma': ma, 'ten_tk': ten,
            'nd_tk':   str(row[8]).strip() if not pd.isna(row[8]) else '',
            'gia_tk':  gia,
            'nhap_tk': dc_safe_float(row[14]),
            'ton_tk':  dc_safe_float(row[24]),
            'kten': dc_norm(ten),
            'knd':  dc_norm(str(row[8]) if not pd.isna(row[8]) else ''),
            'kgia': int(round(gia)) if gia else 0,
        })
    return pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()

def dc_extract_ma_map(dfs_nhap_xuat):
    all_rows = []
    for df in dfs_nhap_xuat:
        if df is None or df.empty: continue
        for _, row in df.iterrows():
            try: int(str(row[0]).strip())
            except: continue
            if pd.isna(row[1]) or pd.isna(row[2]): continue
            ten = str(row[2]).strip()
            if not dc_is_drug(ten): continue
            gia = dc_safe_float(row[5])
            all_rows.append({
                'ma': str(row[1]).strip(), 'ten': ten,
                'nd': str(row[3]).strip() if not pd.isna(row[3]) else '',
                'gia': gia,
                'kten': dc_norm(ten),
                'knd':  dc_norm(str(row[3]) if not pd.isna(row[3]) else ''),
                'kgia': int(round(gia)) if gia else 0,
            })
    if not all_rows: return pd.DataFrame()
    return (pd.DataFrame(all_rows)
              .drop_duplicates(subset=['ma', 'kten', 'knd', 'kgia'])
              .reset_index(drop=True))

def dc_parse_raw_lines_bbkn(df_raw, sl_col=9):
    rows = []
    for _, row in df_raw.iterrows():
        try: int(str(row[0]).strip())
        except: continue
        if pd.isna(row[2]): continue
        ten = str(row[2]).strip()
        if not dc_is_drug(ten): continue
        nd    = str(row[3]).strip() if not pd.isna(row[3]) else ''
        gia   = dc_safe_float(row[8])
        sl    = dc_safe_float(row[sl_col])
        ma_hd = str(row[1]).strip() if not pd.isna(row[1]) else ''
        rows.append({'ten': ten, 'nd': nd, 'gia': gia, 'sl': sl, 'ma_hd': ma_hd,
                     'kten': dc_norm(ten), 'knd': dc_norm(nd),
                     'kgia': int(round(gia)) if gia else 0})
    return pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()

def dc_parse_raw_lines_bbkk(df_raw, sl_col=8):
    rows = []
    for _, row in df_raw.iterrows():
        try: int(str(row[0]).strip())
        except: continue
        if pd.isna(row[1]): continue
        ten = str(row[1]).strip()
        if not dc_is_drug(ten): continue
        nd  = str(row[2]).strip() if not pd.isna(row[2]) else ''
        gia = dc_safe_float(row[4])
        sl  = dc_safe_float(row[sl_col])
        rows.append({'ten': ten, 'nd': nd, 'gia': gia, 'sl': sl, 'ma_hd': '',
                     'kten': dc_norm(ten), 'knd': dc_norm(nd),
                     'kgia': int(round(gia)) if gia else 0})
    return pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()

def dc_match_hpt_to_tk(df_hpt, df_tk, sl_col_tk):
    results = []
    used_tk_idx = set()

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

        active_idx   = [i for i in avail_tk.index if is_active(avail_tk.loc[i, sl_col_tk])]
        inactive_idx = [i for i in avail_tk.index if not is_active(avail_tk.loc[i, sl_col_tk])]

        for i in inactive_idx:
            used_tk_idx.add(i)

        if len(active_idx) == 0:
            results.append(row_hpt_no_tk()); continue

        if len(active_idx) == 1:
            ti = active_idx[0]
            results.append(row_matched(sl_sum, hoa_don, df_tk.loc[ti]))
            used_tk_idx.add(ti); continue

        tk_pool = list(active_idx)
        matched_tk = set(); matched_hpt = set()

        for hi, hr in hpt_list.iterrows():
            for ti in tk_pool:
                if ti in matched_tk: continue
                if abs(hr['sl'] - df_tk.loc[ti, sl_col_tk]) < 0.01:
                    results.append(row_matched(hr['sl'], hr['ma_hd'], df_tk.loc[ti]))
                    matched_tk.add(ti); matched_hpt.add(hi); used_tk_idx.add(ti); break

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

        for ti in tk_pool:
            if ti not in matched_tk:
                if is_active(df_tk.loc[ti, sl_col_tk]):
                    results.append(row_tk_no_hpt(df_tk.loc[ti]))
                used_tk_idx.add(ti)

    for ti, tr in df_tk.iterrows():
        if ti not in used_tk_idx and is_active(tr[sl_col_tk]):
            results.append(row_tk_no_hpt(tr))

    return pd.DataFrame(results)

def dc_run_xnt(dfs_nx, df_xnt_raw, df_tk_raw):
    global_map = dc_extract_ma_map(dfs_nx)
    if global_map.empty:
        return None, "Không đọc được dữ liệu từ file nhập/xuất"

    aug = pd.DataFrame([{'ma': '0005301225', 'ten': 'Augmentin 1g', 'nd': '875mg + 125mg',
                          'gia': 16680, 'kten': dc_norm('Augmentin 1g'),
                          'knd': dc_norm('875mg + 125mg'), 'kgia': 16680}])
    global_map = pd.concat([global_map, aug], ignore_index=True).drop_duplicates(
        subset=['ma', 'kten', 'knd', 'kgia'])

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
            'ton_xnt': dc_safe_float(row[12]),
            'kten': dc_norm(str(row[2])),
            'knd':  dc_norm(str(row[3]) if not pd.isna(row[3]) else ''),
            'kgia': int(round(float(row[8]))) if not pd.isna(row[8]) else 0,
        })
    df_xnt = pd.DataFrame(xnt_rows)
    df_tk  = dc_parse_tk(df_tk_raw)

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

def dc_run_kn(df_bbkn_raw, df_tk_raw, global_map, sl_col=9):
    df_kn = dc_parse_raw_lines_bbkn(df_bbkn_raw, sl_col)
    if df_kn.empty:
        return None, "Không đọc được dữ liệu từ file BBKN"
    df_tk = dc_parse_tk(df_tk_raw)
    if df_tk.empty:
        return None, "Không đọc được dữ liệu từ file Thống kê"
    df_r = dc_match_hpt_to_tk(df_kn, df_tk, 'nhap_tk')
    df_r = df_r.rename(columns={'sl_hpt': 'nhap_hpt', 'sl_tk': 'nhap_tk'})
    return df_r, None

def dc_run_kk(df_bbkk_raw, df_tk_raw, global_map, sl_col=8):
    df_kk = dc_parse_raw_lines_bbkk(df_bbkk_raw, sl_col)
    if df_kk.empty:
        return None, "Không đọc được dữ liệu từ file Biên bản kiểm kê"
    df_tk = dc_parse_tk(df_tk_raw)
    if df_tk.empty:
        return None, "Không đọc được dữ liệu từ file Thống kê"
    df_r = dc_match_hpt_to_tk(df_kk, df_tk, 'ton_tk')
    df_r = df_r.rename(columns={'sl_hpt': 'sl_kk', 'sl_tk': 'ton_tk'})
    return df_r, None


# ── Excel export helpers ──────────────────────────────────────────────────────
def _dc_hdr(ws, headers, fill_hex='1F3864'):
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

def _dc_row(ws, ri, vals, fill, right_cols=(), center_cols=()):
    TH = Side(style='thin')
    for ci, v in enumerate(vals, 1):
        safe = '' if (v is None or (isinstance(v, float) and pd.isna(v))) else v
        c = ws.cell(row=ri, column=ci, value=safe)
        c.font = Font(name='Times New Roman', size=11)
        c.fill = fill
        c.border = Border(left=TH, right=TH, top=TH, bottom=TH)
        if ci in right_cols:
            c.alignment = Alignment(vertical='center', horizontal='right')
            if isinstance(safe, (int, float)): c.number_format = '#,##0.##'
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

def dc_build_xnt_sheets(wb, df_res, tn):
    ws = wb.create_sheet(f"DC XNT {tn.replace('/','_')}")
    _dc_hdr(ws, [('Mã HPT',16),('Tên thuốc (HPT)',32),('Tên thuốc (TK)',28),
              ('Nồng độ',22),('Đơn giá',12),('Tồn cuối HPT',13),
              ('Tồn cuối TK',13),('Chênh lệch',13),('Trạng thái',25),('Phương pháp',18)])
    mm = {'1-1':'Chính xác','exact_ton':'Khớp tồn','nearest_ton':'Gần nhất ⚠️'}
    df_m = df_res[df_res['cl'].notna()]
    df_l = df_m[df_m['cl'].abs()>=0.01].sort_values('cl')
    df_k = df_m[df_m['cl'].abs()<0.01]
    for ri,(_, r) in enumerate(pd.concat([df_l,df_k],ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl)<0.01:  fill,st = FOK,'✅ Khớp'
        elif cl>0:         fill,st = FLE,f'⬆️ HPT cao hơn {cl:+.0f}'
        else:              fill,st = FLE,f'⬇️ HPT thấp hơn {cl:+.0f}'
        if r['method']=='nearest_ton' and abs(cl)>=0.01: fill=FWA
        _dc_row(ws,ri,[r['ma'],r['ten'],r.get('ten_tk',''),r['nd'],r['gia'],
                    r['ton_xnt'],r['ton_tk'],cl,st,mm.get(r['method'],'')],
             fill,right_cols=(5,6,7,8),center_cols=(1,9,10))
    ws.freeze_panes='A2'
    ws2 = wb.create_sheet("XNT – HPT có, TK không")
    _dc_hdr(ws2,[('Mã HPT',16),('Tên thuốc HPT',35),('Nồng độ',25),('Đơn giá',12),('Tồn HPT',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['method']=='no_tk'].iterrows(), 2):
        _dc_row(ws2,ri,[r['ma'],r['ten'],r['nd'],r['gia'],r['ton_xnt'],'HPT có nhưng TK không theo dõi'],
             FBL,right_cols=(3,4,5))
    ws3 = wb.create_sheet("XNT – TK có, HPT không")
    _dc_hdr(ws3,[('Mã HPT',16),('Tên thuốc TK',35),('Nồng độ TK',25),('Tồn TK',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['method']=='no_xnt'].iterrows(), 2):
        _dc_row(ws3,ri,[r['ma'],r['ten_tk'],r['nd_tk'],r['ton_tk'],'TK có nhưng HPT không phát sinh'],
             FOR,right_cols=(4,))

def dc_build_kn_sheets(wb, df_res, tn):
    ws = wb.create_sheet(f"DC Kiểm nhập {tn.replace('/','_')}")
    _dc_hdr(ws,[('Mã HPT',16),('Tên thuốc (HPT)',32),('Tên thuốc (TK)',28),
             ('Nồng độ',22),('Đơn giá',12),('Nhập HPT (tổng)',13),
             ('Số HĐ',8),('Danh sách HĐ',22),('Nhập TK',12),('Chênh lệch',13),('Trạng thái',28)])
    df_m = df_res[df_res['status']=='matched']
    df_l = df_m[df_m['cl'].abs()>=0.01].sort_values('cl')
    df_k = df_m[df_m['cl'].abs()<0.01]
    for ri,(_, r) in enumerate(pd.concat([df_l,df_k],ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl)<0.01:  fill,st = FOK,'✅ Khớp'
        elif cl>0:         fill,st = FLE,f'⬆️ HPT cao hơn {cl:+.0f}'
        else:              fill,st = FLE,f'⬇️ HPT thấp hơn {cl:+.0f}'
        _dc_row(ws,ri,[r['ma'],r['ten_hpt'],r['ten_tk'],r['nd'],r['gia'],
                    r['nhap_hpt'],r.get('n_hoadon',''),r.get('hoa_don',''),
                    r['nhap_tk'],cl,st],
             fill,right_cols=(5,6,8,9,10),center_cols=(1,11))
    ws.freeze_panes='A2'
    ws2 = wb.create_sheet("KN – HPT có, TK không")
    _dc_hdr(ws2,[('Mã HPT',16),('Tên thuốc HPT',35),('Nồng độ',25),('Đơn giá',12),
              ('Nhập HPT (tổng)',13),('Danh sách HĐ',22),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='hpt_no_tk'].iterrows(), 2):
        _dc_row(ws2,ri,[r['ma'],r['ten_hpt'],r['nd'],r['gia'],r['nhap_hpt'],
                     r.get('hoa_don',''),'HPT có nhưng TK không có số nhập'],
             FBL,right_cols=(3,4,5))
    ws3 = wb.create_sheet("KN – TK có, HPT không")
    _dc_hdr(ws3,[('Mã HPT',16),('Tên thuốc TK',35),('Nồng độ TK',25),('Nhập TK',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='tk_no_hpt'].iterrows(), 2):
        _dc_row(ws3,ri,[r['ma'],r['ten_tk'],r['nd'],r['nhap_tk'],'TK có nhập nhưng HPT không có'],
             FOR,right_cols=(4,))

def dc_build_kk_sheets(wb, df_res, tn):
    ws = wb.create_sheet(f"DC Kiểm kê {tn.replace('/','_')}")
    _dc_hdr(ws,[('Mã HPT',16),('Tên thuốc (HPT)',32),('Tên thuốc (TK)',28),
             ('Nồng độ',22),('Đơn giá',12),('SL Kiểm kê (tổng)',14),
             ('Tồn cuối TK',13),('Chênh lệch',13),('Trạng thái',28)])
    df_m = df_res[df_res['status']=='matched']
    df_l = df_m[df_m['cl'].abs()>=0.01].sort_values('cl')
    df_k = df_m[df_m['cl'].abs()<0.01]
    for ri,(_, r) in enumerate(pd.concat([df_l,df_k],ignore_index=True).iterrows(), 2):
        cl = r['cl']
        if abs(cl)<0.01:  fill,st = FOK,'✅ Khớp'
        elif cl>0:         fill,st = FLE,f'⬆️ HPT cao hơn {cl:+.0f}'
        else:              fill,st = FLE,f'⬇️ HPT thấp hơn {cl:+.0f}'
        _dc_row(ws,ri,[r['ma'],r['ten_hpt'],r['ten_tk'],r['nd'],r['gia'],
                    r['sl_kk'],r['ton_tk'],cl,st],
             fill,right_cols=(5,6,7,8),center_cols=(1,9))
    ws.freeze_panes='A2'
    ws2 = wb.create_sheet("KK – HPT có, TK không")
    _dc_hdr(ws2,[('Mã HPT',16),('Tên thuốc HPT',35),('Nồng độ',25),('Đơn giá',12),
              ('SL Kiểm kê',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='hpt_no_tk'].iterrows(), 2):
        _dc_row(ws2,ri,[r['ma'],r['ten_hpt'],r['nd'],r['gia'],r['sl_kk'],'HPT có nhưng TK không theo dõi tồn'],
             FBL,right_cols=(3,4,5))
    ws3 = wb.create_sheet("KK – TK có, HPT không")
    _dc_hdr(ws3,[('Mã HPT',16),('Tên thuốc TK',35),('Nồng độ TK',25),('Tồn TK',12),('Ghi chú',35)])
    for ri,(_, r) in enumerate(df_res[df_res['status']=='tk_no_hpt'].iterrows(), 2):
        _dc_row(ws3,ri,[r['ma'],r['ten_tk'],r['nd'],r['ton_tk'],'TK có tồn nhưng không có trong kiểm kê'],
             FOR,right_cols=(4,))

def dc_build_summary(wb, results_map, tn):
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

def dc_export_excel(res_xnt, res_kn, res_kk, tn):
    wb=Workbook(); wb.remove(wb.active)
    if res_xnt is not None: dc_build_xnt_sheets(wb, res_xnt, tn)
    if res_kn  is not None: dc_build_kn_sheets(wb, res_kn, tn)
    if res_kk  is not None: dc_build_kk_sheets(wb, res_kk, tn)
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
    dc_build_summary(wb, rm, tn)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE BBKK – Biên Bản Kiểm Kê
# ══════════════════════════════════════════════════════════════════════════════
import calendar

def get_last_day(thang, nam):
    """Trả về ngày cuối tháng."""
    return calendar.monthrange(nam, thang)[1]

def ten_thang_viet(thang):
    return f"Tháng {thang}"

def format_ngay_viet(ngay, thang, nam):
    return f"ngày {ngay} tháng {thang} năm {nam}"

BBKK_W = {1:5, 2:35, 3:14, 4:7.5, 5:10, 6:32, 7:11, 8:10, 9:10, 10:8, 11:10}
BBKK_A = {1:('center','center'), 2:('left','center'), 3:('left','center'),
          4:('center','center'), 5:('center','center'), 6:('left','center'),
          7:('center','center'), 8:('right','center'), 9:('right','center'),
          10:('right','center'), 11:('center','center')}
BBKK_WRAP = {2, 3, 6}
BBKK_NUM  = {8, 9}

def bbkk_h(ws, r):
    ml = 1
    for c in BBKK_WRAP:
        v = ws.cell(row=r, column=c).value
        if not v or not isinstance(v, str): continue
        cw = max(BBKK_W.get(c, 15) * 1.1, 1)
        ml = max(ml, sum(max(1, math.ceil(len(ln)/cw)) for ln in v.split('\n')))
    return max(20, min(ml * 14.3 + 4, 120))

def parse_bbkk_raw(raw_df):
    """
    Parse dữ liệu thô BBKK từ HPT.
    Columns in raw: 0=STT, 1=TenThuoc, 2=NongDo, 3=DVT, 4=DonGia,
                    5=SoLo, 6=HangSX, 7=HanDung, 8=SLSoSach, 9=ThanhTien, 10=SLThucTe
    Lọc dòng SL Thực tế = 0 (col 10); nếu col10 NaN dùng col8.
    Trả về list of row (pandas Series)
    """
    drugs, skipped = [], 0
    for _, row in raw_df.iterrows():
        try: int(str(row[0]).strip())
        except: continue
        if pd.isna(row[1]) or not isinstance(row[1], str): continue
        # Số lượng thực tế: col 10 nếu có, else col 8
        sl_tt = row[10] if not pd.isna(row.get(10, float('nan'))) else (row[8] if not pd.isna(row[8]) else 0)
        try: sl_tt = float(sl_tt)
        except: sl_tt = 0
        if sl_tt == 0:
            skipped += 1
            continue
        drugs.append(row)
    return drugs, {'drugs': len(drugs), 'skipped': skipped}

def build_bbkk(tmpl_bytes, drugs, thang, nam):
    """
    Điền dữ liệu vào form BBKK chuẩn.
    Tự động cập nhật tháng/năm trong tiêu đề và ngày kiểm kê.
    """
    wb = load_workbook(io.BytesIO(tmpl_bytes))
    ws = wb.active

    last_day = get_last_day(thang, nam)

    # ── Cập nhật tiêu đề tháng/năm ──────────────────────────────────────────
    # R3C5: "Tháng X năm YYYY"
    ws.cell(row=3, column=5).value = f"Tháng {thang} năm {nam}"

    # R10C1: dòng kiểm kê tại...
    old_r10 = ws.cell(row=10, column=1).value or ''
    # Thay thế toàn bộ ngày giờ trong dòng R10
    import re as _re
    new_r10 = _re.sub(
        r'ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+',
        f'ngày {last_day} tháng {thang} năm {nam}',
        old_r10
    )
    ws.cell(row=10, column=1).value = new_r10

    # ── Lấy style từ template ────────────────────────────────────────────────
    DS = 13  # data start row (row 13 = first data row in template, rows 11-12 = headers)
    cs = {c: gs(ws, DS, c) for c in range(1, 12)}
    ds = {c: gs(ws, DS+1 if ws.max_row > DS else DS, c) for c in range(1, 12)}

    # Lấy style từ row đầu tiên có số liệu
    first_data_row = None
    for r in range(DS, min(DS+5, ws.max_row+1)):
        if ws.cell(row=r, column=1).value is not None:
            first_data_row = r
            break
    if first_data_row:
        ds = {c: gs(ws, first_data_row, c) for c in range(1, 12)}

    # ── Tìm vị trí "Tổng khoản" để insert rows ──────────────────────────────
    fs = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and 'Tổng khoản' in str(cell.value):
                fs = cell.row; break
        if fs: break
    if not fs: fs = DS + 185  # fallback

    # Xóa dữ liệu cũ từ DS đến fs-1
    for m in [str(mr) for mr in ws.merged_cells.ranges if DS <= mr.min_row < fs]:
        ws.merged_cells.remove(m)
    for r in range(DS, fs):
        for c in range(1, 12):
            try: ws.cell(row=r, column=c).value = None
            except: pass

    # ── Insert rows nếu cần ─────────────────────────────────────────────────
    need = len(drugs) + 1  # +1 cho dòng Tổng khoản
    current_space = fs - DS
    if need > current_space:
        ins = need - current_space
        ws.insert_rows(fs, ins)
        fs += ins

    # ── Ghi dữ liệu ──────────────────────────────────────────────────────────
    def wdr_kk(rn, stt, dr):
        ten = str(dr[1]).strip() if not pd.isna(dr[1]) else ''
        nd  = str(dr[2]).strip() if not pd.isna(dr[2]) else ''
        # Ghép tên + nồng độ vào cột 2 (như template gốc)
        ten_full = f"{ten} {nd}".strip() if nd else ten
        cols = [
            (1,  stt,                                               'center', False, None),
            (2,  ten_full,                                          'left',   True,  None),
            (3,  nd,                                                'left',   False, None),
            (4,  str(dr[3]).strip() if not pd.isna(dr[3]) else '',  'center', False, None),
            (5,  str(dr[5]).strip() if not pd.isna(dr[5]) else '',  'center', False, None),
            (6,  str(dr[6]).strip() if not pd.isna(dr[6]) else '',  'left',   True,  None),
            (7,  dr[7] if isinstance(dr[7], datetime.datetime)
                 else ('' if pd.isna(dr[7]) else dr[7]),            'center', False, 'DD/MM/YYYY'),
            (8,  float(dr[8]) if not pd.isna(dr[8]) else 0,        'right',  False, '#,##0'),
        ]
        sl_tt = dr[10] if not pd.isna(dr.get(10, float('nan'))) else (dr[8] if not pd.isna(dr[8]) else 0)
        try: sl_tt = float(sl_tt)
        except: sl_tt = 0
        cols.append((9, sl_tt, 'right', False, '#,##0'))
        for col, val, ha, wrap, fmt in cols:
            cl = ws.cell(row=rn, column=col, value=val)
            ap(cl, ds[col])
            cl.font = Font(name='Times New Roman', size=11)
            cl.alignment = Alignment(horizontal=ha, vertical='center', wrap_text=wrap)
            if fmt and val != '': cl.number_format = fmt
        # Cột 10 (Hỏng) và 11 (Ghi chú) để trống
        for c in (10, 11):
            cc = ws.cell(row=rn, column=c, value='')
            ap(cc, ds[c])

    stt = 1
    for dr in drugs:
        wdr_kk(DS + stt - 1, stt, dr)
        stt += 1

    # ── Dòng Tổng khoản ──────────────────────────────────────────────────────
    tr = DS + len(drugs)
    lbl = ws.cell(row=tr, column=2, value=f'Tổng khoản: {len(drugs)} khoản')
    lbl.font = Font(name='Times New Roman', bold=True, size=11)
    lbl.alignment = Alignment(horizontal='left', vertical='center')
    lbl.border = b_med()
    ws.cell(row=tr, column=1).border = b_med()
    for c in range(3, 12):
        ws.cell(row=tr, column=c).border = b_med()
    ws.row_dimensions[tr].height = 20

    # ── Format vùng data ─────────────────────────────────────────────────────
    for r in range(DS, tr):
        ws.row_dimensions[r].height = bbkk_h(ws, r)
        for col in range(1, 12):
            cl = ws.cell(row=r, column=col)
            ha, va = BBKK_A.get(col, ('left', 'center'))
            safe_set(cl, fill=NO_FILL, border=b_thin(),
                     font=Font(name='Times New Roman', size=11),
                     alignment=Alignment(horizontal=ha, vertical=va, wrap_text=col in BBKK_WRAP))
            if col in BBKK_NUM and cl.value is not None and cl.value != '':
                cl.number_format = '#,##0'
            if col == 7 and isinstance(cl.value, datetime.datetime):
                cl.number_format = 'DD/MM/YYYY'

    # ── Cập nhật footer: ngày ký biên bản (dòng cuối) ────────────────────────
    # Tìm dòng có "Ngày" trong phần cuối
    for r in range(tr+1, min(tr+15, ws.max_row+1)):
        v = ws.cell(row=r, column=1).value
        if v and isinstance(v, str) and 'Ngày' in v:
            import re as _re2
            ws.cell(row=r, column=1).value = _re2.sub(
                r'Ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+',
                f'Ngày {last_day} tháng {thang} năm {nam}',
                v
            )
            break

    for col, w in BBKK_W.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    for a, v in [('left', .5), ('right', .5), ('top', .7), ('bottom', .7),
                 ('header', .3), ('footer', .3)]:
        setattr(ws.page_margins, a, v)
    ws.print_title_rows = '1:12'
    ws.freeze_panes = ws.cell(row=DS, column=1)

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()


def update_xnt_dates(tmpl_bytes, thang, nam):
    """Cập nhật tháng/năm trong template XNT."""
    wb = load_workbook(io.BytesIO(tmpl_bytes))
    ws = wb.active
    last_day = get_last_day(thang, nam)
    import re as _re
    for r in range(1, min(15, ws.max_row+1)):
        for c in range(1, ws.max_column+1):
            v = ws.cell(row=r, column=c).value
            if not v or not isinstance(v, str): continue
            # Thay "Tháng X năm YYYY"
            new_v = _re.sub(r'Tháng\s+\d+\s+năm\s+\d+', f'Tháng {thang} năm {nam}', v)
            # Thay ngày cuối tháng
            new_v = _re.sub(r'ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+',
                            f'ngày {last_day} tháng {thang} năm {nam}', new_v)
            if new_v != v:
                ws.cell(row=r, column=c).value = new_v
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()


def update_bbkn_dates(tmpl_bytes, thang, nam):
    """Cập nhật tháng/năm trong template BBKN."""
    wb = load_workbook(io.BytesIO(tmpl_bytes))
    ws = wb.active
    last_day = get_last_day(thang, nam)
    import re as _re
    for r in range(1, min(20, ws.max_row+1)):
        for c in range(1, ws.max_column+1):
            v = ws.cell(row=r, column=c).value
            if not v or not isinstance(v, str): continue
            new_v = _re.sub(r'Tháng\s+\d+\s+năm\s+\d+', f'Tháng {thang} năm {nam}', v)
            new_v = _re.sub(r'ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+',
                            f'ngày {last_day} tháng {thang} năm {nam}', new_v)
            if new_v != v:
                ws.cell(row=r, column=c).value = new_v
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  GIAO DIỆN CHÍNH
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
  <div class="badge">🏥 Bệnh viện Đà Nẵng · Khoa Dược</div>
  <h1>HỆ THỐNG TỰ ĐỘNG HÓA<br>BIÊN BẢN DƯỢC</h1>
  <p class="sub">Biên Bản Kiểm (BBKN · BBKK) &nbsp;·&nbsp; Xuất Nhập Tồn (XNT) &nbsp;·&nbsp; Đối Chiếu Dược</p>
</div>
""", unsafe_allow_html=True)

# ── 3 Tab chính ───────────────────────────────────────────────────────────────
tab_bienban, tab_xnt_main, tab_dc = st.tabs([
    "📋 Biên Bản Kiểm",
    "📊 Báo Cáo XNT",
    "🔍 Đối Chiếu Dược",
])


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  TAB 1 – BIÊN BẢN KIỂM (gộp BBKN + BBKK)                              ║
# ╚══════════════════════════════════════════════════════════════════════════╝
with tab_bienban:
    st.markdown("""<div class="info-box">
    Chọn loại biên bản cần xử lý. Mỗi biên bản có bộ chọn <b>Tháng / Năm</b> riêng —
    hệ thống sẽ tự động điền tháng, năm và ngày cuối tháng vào đúng vị trí trong form.
    </div>""", unsafe_allow_html=True)

    sub_bbkn, sub_bbkk = st.tabs([
        "📥 Kiểm Nhập (BBKN)",
        "🔎 Kiểm Kê (BBKK)",
    ])

    # ── SUB-TAB BBKN ──────────────────────────────────────────────────────────
    with sub_bbkn:
        st.markdown("""<div class="tab-desc">
        Upload <b>file dữ liệu thô BBKN từ HPT</b> (.xls/.xlsx) và <b>file form chuẩn kiểm nhập</b>.<br>
        Logic: Lọc bỏ dòng Số lượng nhập = 0 · Phân nhóm theo công ty · Tính thành tiền tự động.
        </div>""", unsafe_allow_html=True)

        bbkn_ta, bbkn_tb = st.columns(2)
        with bbkn_ta:
            bbkn_thang = st.selectbox("📅 Tháng báo cáo", range(1, 13), index=2,
                format_func=lambda x: f"Tháng {x}", key="bbkn_thang")
        with bbkn_tb:
            bbkn_nam = st.number_input("📅 Năm", min_value=2024, max_value=2030,
                value=2026, key="bbkn_nam")

        col1, col2 = st.columns(2)
        with col1: raw_file_bbkn = st.file_uploader("📂 File dữ liệu thô HPT (BBKN)", type=["xls","xlsx"], key="bbkn_raw")
        with col2: tpl_file_bbkn = st.file_uploader("📄 File form chuẩn Kiểm Nhập", type=["xlsx"], key="bbkn_tpl")
        st.markdown("<hr>", unsafe_allow_html=True)

        ready_bbkn = raw_file_bbkn is not None and tpl_file_bbkn is not None
        if st.button("⚡ Bắt đầu xử lý BBKN", disabled=not ready_bbkn, key="btn_bbkn"):
            with st.spinner("Đang xử lý BBKN..."):
                try:
                    raw_b = raw_file_bbkn.read()
                    if raw_file_bbkn.name.endswith('.xls'):
                        try:    raw_df=pd.read_excel(io.BytesIO(raw_b),sheet_name=0,header=None,engine='xlrd')
                        except: raw_df=pd.read_excel(io.BytesIO(raw_b),sheet_name=0,header=None)
                    else:
                        raw_df=pd.read_excel(io.BytesIO(raw_b),sheet_name=0,header=None)
                    tpl_b = tpl_file_bbkn.read()
                    # Cập nhật tháng/năm trong form trước khi build
                    tpl_b = update_bbkn_dates(tpl_b, bbkn_thang, bbkn_nam)
                    companies, stats = parse_companies(raw_df, 9)
                    if not companies:
                        st.error("❌ Không tìm thấy dữ liệu hợp lệ. Kiểm tra lại file HPT.")
                        st.stop()
                    result = build_bbkn(tpl_b, companies)
                    st.session_state.update(bbkn_result=result, bbkn_stats=stats,
                                            bbkn_done=True, bbkn_thang=bbkn_thang, bbkn_nam=bbkn_nam)
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}"); st.exception(e)

        if st.session_state.get("bbkn_done"):
            stats  = st.session_state["bbkn_stats"]
            result = st.session_state["bbkn_result"]
            _t = st.session_state.get("bbkn_thang", bbkn_thang)
            _n = st.session_state.get("bbkn_nam", bbkn_nam)
            fname  = f"BBKN_T{_t}_{_n}_HoanChinh.xlsx"
            st.markdown(f"""
            <div class="ok-box">
              <div class="icon">✅</div>
              <h3>Xử lý BBKN hoàn tất – Tháng {_t}/{_n}!</h3>
              <p>File sẵn sàng — tải về và in ký hội đồng.</p>
            </div>
            <div class="stat-grid">
              <div class="stat-card"><div class="num">{stats['companies']}</div><div class="lbl">Công ty cung cấp</div></div>
              <div class="stat-card"><div class="num">{stats['drugs']}</div><div class="lbl">Mặt hàng có nhập</div></div>
              <div class="stat-card"><div class="num">{stats['skipped']}</div><div class="lbl">Dòng SL=0 đã lọc</div></div>
            </div>""", unsafe_allow_html=True)
            st.download_button(label=f"⬇️ Tải File BBKN – {fname}", data=result, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_bbkn")
            st.markdown("""<div class="note">💡 <b>Khi in:</b> File đã thiết lập sẵn <b>A4 Ngang · Fit All Columns on One Page</b>.
            Mở Excel → Ctrl+P → in ngay.</div>""", unsafe_allow_html=True)

    # ── SUB-TAB BBKK ──────────────────────────────────────────────────────────
    with sub_bbkk:
        st.markdown("""<div class="tab-desc">
        Upload <b>file dữ liệu thô BBKK từ HPT</b> (.xlsx) và <b>file form chuẩn kiểm kê</b>.<br>
        Logic: Lọc bỏ dòng Số lượng thực tế = 0 · Điền đủ tên, lô, hạn dùng, SL sổ sách & thực tế.
        </div>""", unsafe_allow_html=True)

        bbkk_ta, bbkk_tb = st.columns(2)
        with bbkk_ta:
            bbkk_thang = st.selectbox("📅 Tháng báo cáo", range(1, 13), index=2,
                format_func=lambda x: f"Tháng {x}", key="bbkk_thang")
        with bbkk_tb:
            bbkk_nam = st.number_input("📅 Năm", min_value=2024, max_value=2030,
                value=2026, key="bbkk_nam")

        col1k, col2k = st.columns(2)
        with col1k: raw_file_bbkk = st.file_uploader("📂 File dữ liệu thô HPT (BBKK)", type=["xls","xlsx"], key="bbkk_raw")
        with col2k: tpl_file_bbkk = st.file_uploader("📄 File form chuẩn Kiểm Kê", type=["xlsx"], key="bbkk_tpl")
        st.markdown("<hr>", unsafe_allow_html=True)

        ready_bbkk = raw_file_bbkk is not None and tpl_file_bbkk is not None
        if st.button("⚡ Bắt đầu xử lý BBKK", disabled=not ready_bbkk, key="btn_bbkk"):
            with st.spinner("Đang xử lý BBKK..."):
                try:
                    raw_b_kk = raw_file_bbkk.read()
                    if raw_file_bbkk.name.endswith('.xls'):
                        try:    raw_df_kk=pd.read_excel(io.BytesIO(raw_b_kk),sheet_name=0,header=None,engine='xlrd')
                        except: raw_df_kk=pd.read_excel(io.BytesIO(raw_b_kk),sheet_name=0,header=None)
                    else:
                        raw_df_kk=pd.read_excel(io.BytesIO(raw_b_kk),sheet_name=0,header=None)
                    tpl_b_kk = tpl_file_bbkk.read()
                    drugs_kk, stats_kk = parse_bbkk_raw(raw_df_kk)
                    if not drugs_kk:
                        st.error("❌ Không tìm thấy dữ liệu hợp lệ. Kiểm tra lại file HPT.")
                        st.stop()
                    result_kk = build_bbkk(tpl_b_kk, drugs_kk, bbkk_thang, bbkk_nam)
                    st.session_state.update(bbkk_result=result_kk, bbkk_stats=stats_kk,
                                            bbkk_done=True, bbkk_thang=bbkk_thang, bbkk_nam=bbkk_nam)
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}"); st.exception(e)

        if st.session_state.get("bbkk_done"):
            stats_kk2  = st.session_state["bbkk_stats"]
            result_kk2 = st.session_state["bbkk_result"]
            _tk = st.session_state.get("bbkk_thang", bbkk_thang)
            _nk = st.session_state.get("bbkk_nam", bbkk_nam)
            fname_kk = f"BBKK_T{_tk}_{_nk}_HoanChinh.xlsx"
            st.markdown(f"""
            <div class="ok-box">
              <div class="icon">✅</div>
              <h3>Xử lý BBKK hoàn tất – Tháng {_tk}/{_nk}!</h3>
              <p>File sẵn sàng — tải về, in và ký hội đồng kiểm kê.</p>
            </div>
            <div class="stat-grid">
              <div class="stat-card"><div class="num">{stats_kk2['drugs']}</div><div class="lbl">Mặt hàng có tồn</div></div>
              <div class="stat-card"><div class="num">{stats_kk2['skipped']}</div><div class="lbl">Dòng SL=0 đã lọc</div></div>
            </div>""", unsafe_allow_html=True)
            st.download_button(label=f"⬇️ Tải File BBKK – {fname_kk}", data=result_kk2, file_name=fname_kk,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_bbkk")
            st.markdown("""<div class="note">💡 <b>Khi in:</b> File đã thiết lập sẵn <b>A4 Đứng</b>.
            Mở Excel → Ctrl+P → in ngay.</div>""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  TAB 2 – XNT                                                            ║
# ╚══════════════════════════════════════════════════════════════════════════╝
with tab_xnt_main:
    st.markdown("""<div class="tab-desc">
    Upload <b>file dữ liệu thô XNT từ HPT</b> (.xlsx) và <b>file BBXNT form chuẩn</b>.<br>
    Logic: Giữ dòng có Tồn cuối ≠ 0 · Phân nhóm theo công ty · Thành tiền = Đơn giá × Tồn cuối.
    </div>""", unsafe_allow_html=True)

    xnt_ta, xnt_tb = st.columns(2)
    with xnt_ta:
        xnt_thang = st.selectbox("📅 Tháng báo cáo", range(1, 13), index=2,
            format_func=lambda x: f"Tháng {x}", key="xnt_thang")
    with xnt_tb:
        xnt_nam = st.number_input("📅 Năm", min_value=2024, max_value=2030,
            value=2026, key="xnt_nam")

    col1x, col2x = st.columns(2)
    with col1x: raw_file_xnt = st.file_uploader("📂 File dữ liệu thô HPT (XNT)", type=["xlsx"], key="xnt_raw")
    with col2x: tpl_file_xnt = st.file_uploader("📄 File form chuẩn XNT", type=["xlsx"], key="xnt_tpl")
    st.markdown("<hr>", unsafe_allow_html=True)

    ready_xnt_main = raw_file_xnt is not None and tpl_file_xnt is not None
    if st.button("⚡ Bắt đầu xử lý XNT", disabled=not ready_xnt_main, key="btn_xnt_main"):
        with st.spinner("Đang xử lý XNT..."):
            try:
                raw_df2 = pd.read_excel(io.BytesIO(raw_file_xnt.read()), sheet_name=0, header=None)
                tpl_b2  = tpl_file_xnt.read()
                # Cập nhật tháng/năm trong form
                tpl_b2 = update_xnt_dates(tpl_b2, xnt_thang, xnt_nam)
                companies2, stats2 = parse_companies(raw_df2, 12)
                if not companies2:
                    st.error("❌ Không tìm thấy dữ liệu hợp lệ."); st.stop()
                result2 = build_xnt(tpl_b2, companies2)
                st.session_state.update(xnt_main_result=result2, xnt_main_stats=stats2,
                                        xnt_main_done=True, xnt_main_thang=xnt_thang, xnt_main_nam=xnt_nam)
            except Exception as e:
                st.error(f"❌ Lỗi: {e}"); st.exception(e)

    if st.session_state.get("xnt_main_done"):
        stats2  = st.session_state["xnt_main_stats"]
        result2 = st.session_state["xnt_main_result"]
        _tx = st.session_state.get("xnt_main_thang", xnt_thang)
        _nx = st.session_state.get("xnt_main_nam", xnt_nam)
        fname2  = f"XNT_T{_tx}_{_nx}_HoanChinh.xlsx"
        st.markdown(f"""
        <div class="ok-box">
          <div class="icon">✅</div>
          <h3>Xử lý XNT hoàn tất – Tháng {_tx}/{_nx}!</h3>
          <p>File sẵn sàng — tải về và kiểm tra.</p>
        </div>
        <div class="stat-grid">
          <div class="stat-card"><div class="num">{stats2['companies']}</div><div class="lbl">Công ty cung cấp</div></div>
          <div class="stat-card"><div class="num">{stats2['drugs']}</div><div class="lbl">Mặt hàng phát sinh</div></div>
          <div class="stat-card"><div class="num">{stats2['skipped']}</div><div class="lbl">Dòng tồn cuối=0 đã lọc</div></div>
        </div>""", unsafe_allow_html=True)
        st.download_button(label=f"⬇️ Tải File XNT – {fname2}", data=result2, file_name=fname2,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xnt_main")
        st.markdown("""<div class="note">💡 <b>Khi in:</b> A4 Ngang · Fit All Columns on One Page.
        Mở Excel → Ctrl+P → in ngay.</div>""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  TAB 3 – ĐỐI CHIẾU DƯỢC                                                ║
# ╚══════════════════════════════════════════════════════════════════════════╝
with tab_dc:
    st.markdown("""<div class="info-box">
    Module đối chiếu số liệu HPT vs Thống kê. Upload file 1 lần, dùng cho cả 3 loại đối chiếu:
    <b>XNT · Kiểm nhập · Kiểm kê</b>.
    </div>""", unsafe_allow_html=True)

    # Tháng / Năm
    dca, dcb = st.columns(2)
    with dca: dc_thang = st.selectbox("Tháng báo cáo", range(1,13), index=2, format_func=lambda x: f"Tháng {x}", key="dc_thang")
    with dcb: dc_nam   = st.number_input("Năm", min_value=2024, max_value=2030, value=2026, key="dc_nam")
    dc_tn = f"T{dc_thang}/{dc_nam}"

    st.markdown("""
    <div class="upload-section">
    <h4>📂 Upload File – dùng chung cho cả 3 module đối chiếu</h4>
    </div>""", unsafe_allow_html=True)

    dcu1, dcu2 = st.columns(2)
    with dcu1:
        dc_f_nhap = st.file_uploader("📥 Báo cáo Nhập hàng trong tháng (có Mã HPT)",
            type=["xlsx","xls"], accept_multiple_files=True, key="dc_nhap")
    with dcu2:
        dc_f_xuat = st.file_uploader("📤 Báo cáo Xuất kho trong tháng (có Mã HPT)",
            type=["xlsx","xls"], accept_multiple_files=True, key="dc_xuat")

    dcu3, dcu4 = st.columns(2)
    with dcu3:
        dc_f_tk = st.file_uploader("📋 File XNT Thống kê – số chuẩn (dùng chung)",
            type=["xlsx","xls"], key="dc_tk")
    with dcu4:
        dc_f_xnt_tho = st.file_uploader("📊 File XNT thô HPT (chỉ cho tab Đối chiếu XNT)",
            type=["xlsx","xls"], key="dc_xnt_tho")

    dcu5, dcu6 = st.columns(2)
    with dcu5:
        dc_f_bbkn = st.file_uploader("📄 Biên bản Kiểm nhập – BBKN",
            type=["xlsx","xls"], key="dc_bbkn")
    with dcu6:
        dc_f_bbkk = st.file_uploader("📄 Biên bản Kiểm kê – BBKK",
            type=["xlsx","xls"], key="dc_bbkk")

    # Hiển thị trạng thái bảng mã
    gmap_dc = st.session_state.get('dc_global_map')
    if gmap_dc is not None and not gmap_dc.empty:
        n_ma = gmap_dc['ma'].nunique()
        st.markdown(f'<div class="map-box">✅ <b>Bảng mã hàng sẵn sàng:</b> {n_ma} mã HPT</div>',
                    unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # Tùy chọn cột SL
    dc_sl_col_kn = 9; dc_sl_col_kk = 8
    with st.expander("⚙️ Tùy chọn cột số lượng (nếu cấu trúc file khác mặc định)", expanded=False):
        dcc1, dcc2 = st.columns(2)
        with dcc1:
            if dc_f_bbkn:
                try:
                    dc_f_bbkn.seek(0)
                    pv = pd.read_excel(io.BytesIO(dc_f_bbkn.read()), sheet_name=0, header=None, nrows=12)
                    dc_f_bbkn.seek(0)
                    opts = {f"Cột {i} | {str(pv.iloc[9,i] if len(pv)>9 else '')[:28]}": i
                            for i in range(len(pv.columns))}
                    def_kn = next((k for k,v in opts.items() if v==9), list(opts.keys())[0])
                    dc_sl_col_kn = opts[st.selectbox("Cột SL nhập trong BBKN:", list(opts.keys()),
                        index=list(opts.keys()).index(def_kn), key="dc_sel_kn")]
                except: pass
        with dcc2:
            if dc_f_bbkk:
                try:
                    dc_f_bbkk.seek(0)
                    pv2 = pd.read_excel(io.BytesIO(dc_f_bbkk.read()), sheet_name=0, header=None, nrows=12)
                    dc_f_bbkk.seek(0)
                    opts2 = {f"Cột {i} | {str(pv2.iloc[9,i] if len(pv2)>9 else '')[:28]}": i
                             for i in range(len(pv2.columns))}
                    def_kk = next((k for k,v in opts2.items() if v==8), list(opts2.keys())[0])
                    dc_sl_col_kk = opts2[st.selectbox("Cột SL thực tế trong BBKK:", list(opts2.keys()),
                        index=list(opts2.keys()).index(def_kk), key="dc_sel_kk")]
                except: pass

    # 3 sub-tabs bên trong tab Đối Chiếu
    sub_xnt, sub_kn, sub_kk = st.tabs([
        "📊 Đối chiếu XNT",
        "📥 Đối chiếu Kiểm nhập",
        "🔍 Đối chiếu Kiểm kê",
    ])

    # ── SUB-TAB: XNT ──────────────────────────────────────────────────────────
    with sub_xnt:
        st.markdown("""<div class="info-box">
        Cần file <b>XNT thô HPT</b> + <b>XNT Thống kê</b> + file <b>Nhập/Xuất kho</b>.
        </div>""", unsafe_allow_html=True)
        ready_dc_xnt = bool((dc_f_nhap or dc_f_xuat) and dc_f_xnt_tho and dc_f_tk)
        if st.button("🔍 Chạy Đối chiếu XNT", key="dc_btn_xnt", disabled=not ready_dc_xnt):
            with st.spinner("Đang xử lý đối chiếu XNT..."):
                try:
                    dfs_nx = []
                    for f in (dc_f_nhap or []):
                        f.seek(0); dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))
                    for f in (dc_f_xuat or []):
                        f.seek(0); dfs_nx.append(pd.read_excel(io.BytesIO(f.read()), sheet_name=0, header=None))
                    st.session_state['dc_global_map'] = dc_extract_ma_map(dfs_nx)
                    dc_f_xnt_tho.seek(0); dc_f_tk.seek(0)
                    df_xnt_raw = pd.read_excel(io.BytesIO(dc_f_xnt_tho.read()), sheet_name=0, header=None)
                    df_tk_raw  = pd.read_excel(io.BytesIO(dc_f_tk.read()), sheet_name=0, header=None)
                    df_res, err = dc_run_xnt(dfs_nx, df_xnt_raw, df_tk_raw)
                    if err: st.error(f"❌ {err}"); st.stop()
                    st.session_state.update({'dc_xnt_result': df_res, 'dc_xnt_done': True, 'dc_tn': dc_tn})
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}"); st.exception(e)

        if st.session_state.get('dc_xnt_done'):
            df_r = st.session_state['dc_xnt_result']
            dm   = df_r[df_r['cl'].notna()]
            nk   = (dm['cl'].abs()<0.01).sum(); nl = (dm['cl'].abs()>=0.01).sum()
            dn_tk  = df_r[df_r['method']=='no_tk'];  dn_xnt = df_r[df_r['method']=='no_xnt']
            st.markdown(f"""<div class="ok-box"><div class="icon">✅</div>
              <h3>Đối chiếu XNT hoàn tất – {st.session_state.get('dc_tn','')}</h3></div>
            <div class="stat-grid-4">
              <div class="stat-card"><div class="num" style="color:#166534">{nk}</div><div class="lbl">✅ Khớp</div></div>
              <div class="stat-card"><div class="num" style="color:#dc2626">{nl}</div><div class="lbl">⚠️ Chênh lệch</div></div>
              <div class="stat-card"><div class="num" style="color:#2563a8">{len(dn_tk)}</div><div class="lbl">HPT có – TK không</div></div>
              <div class="stat-card"><div class="num" style="color:#d97706">{len(dn_xnt)}</div><div class="lbl">TK có – HPT không</div></div>
            </div>""", unsafe_allow_html=True)
            dl = dm[dm['cl'].abs()>=0.01].sort_values('cl')
            if len(dl):
                st.markdown(f"**⚠️ {nl} dòng chênh lệch:**")
                st.dataframe(dl[['ma','ten','nd','ton_xnt','ton_tk','cl']].rename(columns={
                    'ma':'Mã HPT','ten':'Tên thuốc','nd':'Nồng độ',
                    'ton_xnt':'Tồn HPT','ton_tk':'Tồn TK','cl':'Chênh lệch'}),
                    use_container_width=True, hide_index=True)

    # ── SUB-TAB: KIỂM NHẬP ────────────────────────────────────────────────────
    with sub_kn:
        st.markdown("""<div class="info-box">
        Cần file <b>BBKN</b> + <b>XNT Thống kê</b>. App tự động cộng tổng tất cả hóa đơn cùng tên+nồng độ+giá.
        </div>""", unsafe_allow_html=True)
        ready_dc_kn = bool(dc_f_bbkn and dc_f_tk)
        if st.button("🔍 Chạy Đối chiếu Kiểm nhập", key="dc_btn_kn", disabled=not ready_dc_kn):
            with st.spinner("Đang xử lý Kiểm nhập..."):
                try:
                    dc_f_bbkn.seek(0); dc_f_tk.seek(0)
                    df_bbkn_raw = pd.read_excel(io.BytesIO(dc_f_bbkn.read()), sheet_name=0, header=None)
                    df_tk_raw   = pd.read_excel(io.BytesIO(dc_f_tk.read()), sheet_name=0, header=None)
                    gmap_cur    = st.session_state.get('dc_global_map', pd.DataFrame())
                    df_res, err = dc_run_kn(df_bbkn_raw, df_tk_raw, gmap_cur, dc_sl_col_kn)
                    if err: st.error(f"❌ {err}"); st.stop()
                    st.session_state.update({'dc_kn_result': df_res, 'dc_kn_done': True, 'dc_tn': dc_tn})
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}"); st.exception(e)

        if st.session_state.get('dc_kn_done'):
            df_kn = st.session_state['dc_kn_result']
            dm_kn = df_kn[df_kn['status']=='matched']
            nk_kn = (dm_kn['cl'].abs()<0.01).sum(); nl_kn = (dm_kn['cl'].abs()>=0.01).sum()
            no_tk_kn = df_kn[df_kn['status']=='hpt_no_tk']; no_hpt_kn = df_kn[df_kn['status']=='tk_no_hpt']
            st.markdown(f"""<div class="ok-box"><div class="icon">✅</div>
              <h3>Đối chiếu Kiểm nhập hoàn tất – {st.session_state.get('dc_tn','')}</h3></div>
            <div class="stat-grid-4">
              <div class="stat-card"><div class="num" style="color:#166534">{nk_kn}</div><div class="lbl">✅ Khớp</div></div>
              <div class="stat-card"><div class="num" style="color:#dc2626">{nl_kn}</div><div class="lbl">⚠️ Chênh lệch</div></div>
              <div class="stat-card"><div class="num" style="color:#2563a8">{len(no_tk_kn)}</div><div class="lbl">HPT có – TK không</div></div>
              <div class="stat-card"><div class="num" style="color:#d97706">{len(no_hpt_kn)}</div><div class="lbl">TK có – HPT không</div></div>
            </div>""", unsafe_allow_html=True)
            dl_kn = dm_kn[dm_kn['cl'].abs()>=0.01].sort_values('cl')
            if len(dl_kn):
                st.markdown(f"**⚠️ {nl_kn} dòng chênh lệch:**")
                st.dataframe(dl_kn[['ma','ten_hpt','nd','nhap_hpt','nhap_tk','cl']].rename(columns={
                    'ma':'Mã HPT','ten_hpt':'Tên thuốc (HPT)','nd':'Nồng độ',
                    'nhap_hpt':'Nhập HPT (tổng)','nhap_tk':'Nhập TK','cl':'Chênh lệch'}),
                    use_container_width=True, hide_index=True)
            if len(no_hpt_kn):
                st.markdown(f"**📋 {len(no_hpt_kn)} mã – TK có nhập, HPT không có:**")
                st.dataframe(no_hpt_kn[['ma','ten_tk','nd','nhap_tk']].rename(columns={
                    'ma':'Mã HPT','ten_tk':'Tên thuốc TK','nd':'Nồng độ','nhap_tk':'Nhập TK'}),
                    use_container_width=True, hide_index=True)

    # ── SUB-TAB: KIỂM KÊ ──────────────────────────────────────────────────────
    with sub_kk:
        st.markdown("""<div class="info-box">
        Cần file <b>BBKK</b> + <b>XNT Thống kê</b>. App tự động cộng tổng các dòng cùng tên+nồng độ+giá.
        </div>""", unsafe_allow_html=True)
        ready_dc_kk = bool(dc_f_bbkk and dc_f_tk)
        if st.button("🔍 Chạy Đối chiếu Kiểm kê", key="dc_btn_kk", disabled=not ready_dc_kk):
            with st.spinner("Đang xử lý Kiểm kê..."):
                try:
                    dc_f_bbkk.seek(0); dc_f_tk.seek(0)
                    df_bbkk_raw = pd.read_excel(io.BytesIO(dc_f_bbkk.read()), sheet_name=0, header=None)
                    df_tk_raw   = pd.read_excel(io.BytesIO(dc_f_tk.read()), sheet_name=0, header=None)
                    gmap_cur    = st.session_state.get('dc_global_map', pd.DataFrame())
                    df_res, err = dc_run_kk(df_bbkk_raw, df_tk_raw, gmap_cur, dc_sl_col_kk)
                    if err: st.error(f"❌ {err}"); st.stop()
                    st.session_state.update({'dc_kk_result': df_res, 'dc_kk_done': True, 'dc_tn': dc_tn})
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}"); st.exception(e)

        if st.session_state.get('dc_kk_done'):
            df_kk2 = st.session_state['dc_kk_result']
            dm_kk  = df_kk2[df_kk2['status']=='matched']
            nk_kk  = (dm_kk['cl'].abs()<0.01).sum(); nl_kk = (dm_kk['cl'].abs()>=0.01).sum()
            no_tk2 = df_kk2[df_kk2['status']=='hpt_no_tk']; no_hpt2 = df_kk2[df_kk2['status']=='tk_no_hpt']
            st.markdown(f"""<div class="ok-box"><div class="icon">✅</div>
              <h3>Đối chiếu Kiểm kê hoàn tất – {st.session_state.get('dc_tn','')}</h3></div>
            <div class="stat-grid-4">
              <div class="stat-card"><div class="num" style="color:#166534">{nk_kk}</div><div class="lbl">✅ Khớp</div></div>
              <div class="stat-card"><div class="num" style="color:#dc2626">{nl_kk}</div><div class="lbl">⚠️ Chênh lệch</div></div>
              <div class="stat-card"><div class="num" style="color:#2563a8">{len(no_tk2)}</div><div class="lbl">HPT có – TK không</div></div>
              <div class="stat-card"><div class="num" style="color:#d97706">{len(no_hpt2)}</div><div class="lbl">TK có – HPT không</div></div>
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

    # ── Xuất Excel tổng hợp đối chiếu ─────────────────────────────────────────
    dc_has_any = any(st.session_state.get(k) for k in ['dc_xnt_done','dc_kn_done','dc_kk_done'])
    if dc_has_any:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("### 📥 Tải kết quả đối chiếu tổng hợp")
        st.markdown("""<div class="info-box">
        File Excel gộp tất cả kết quả vào các Sheet riêng biệt.
        Màu <span style="background:#E2EFDA;padding:1px 6px;border-radius:3px">🟢 Xanh</span> = Khớp |
        <span style="background:#FFD7D7;padding:1px 6px;border-radius:3px">🔴 Đỏ</span> = Lệch.
        </div>""", unsafe_allow_html=True)
        tn_export = st.session_state.get('dc_tn', dc_tn)
        excel_bytes = dc_export_excel(
            st.session_state.get('dc_xnt_result'),
            st.session_state.get('dc_kn_result'),
            st.session_state.get('dc_kk_result'),
            tn_export)
        st.download_button(
            label=f"⬇️ Tải Kết Quả Đối Chiếu Tổng Hợp {tn_export} (.xlsx)",
            data=excel_bytes,
            file_name=f"doi_chieu_tong_hop_{tn_export.replace('/','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dc_dl_all")
