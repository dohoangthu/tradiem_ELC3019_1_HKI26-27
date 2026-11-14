import html
import re
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import streamlit as st

# ============================================================
# 1. CẤU HÌNH
# ============================================================
FILE_PATH = '271_ELC3019_ELC3019_1.xlsx'   # tên file Excel (phải trùng tên file trên GitHub)
COURSE = 'ELC3019 – Thanh toán điện tử'
LECTURER = 'Đỗ Hoàng Thu'

# File Excel gồm 13 cột theo thứ tự:
# ClassID | UID | FullName
# | A (trắc nghiệm ch.1-2) | B (thảo luận) | C (điểm cộng TP1) | Thành phần 1
# | E (tự luận ch.3-4) | F (trắc nghiệm ch.3-6) | G (điểm cộng TP2) | Thành phần 2
# | Bài tập nhóm lớn | Vấn đáp (không hiển thị)
COL_NAMES = ['Lop', 'MSV', 'Ho_ten',
             'A', 'B', 'C', 'TP1',
             'E', 'F', 'G', 'TP2',
             'BTN', 'VD']

st.set_page_config(page_title='Tra Cứu Điểm ELC3019', page_icon='🎓', layout='centered')

# ============================================================
# 2. GIAO DIỆN (CSS)
# ============================================================
st.markdown(
    """
    <style>
    .block-container {max-width: 760px; padding-top: 2rem; padding-bottom: 2rem;}
    #MainMenu, footer {visibility: hidden;}

    .hero {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 55%, #38bdf8 100%);
        color: #ffffff; padding: 28px 30px; border-radius: 18px;
        box-shadow: 0 8px 24px rgba(37, 99, 235, 0.28); margin-bottom: 22px;
    }
    .hero h1 {margin: 0; font-size: 1.9rem; color: #ffffff; padding: 0;}
    .hero .course {margin-top: 6px; font-size: 1.05rem; opacity: 0.95;}
    .hero .lecturer {margin-top: 10px; font-size: 0.95rem; opacity: 0.9;}

    .student-card {
        background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
        border-left: 6px solid #10b981; border-radius: 14px;
        padding: 16px 20px; margin: 18px 0 14px 0; color: #064e3b;
    }
    .student-card .name {font-size: 1.3rem; font-weight: 700;}
    .student-card .meta {font-size: 0.95rem; margin-top: 2px;}

    table.score-table {
        width: 100%; border-collapse: collapse; margin-top: 8px;
        border: 1px solid #d1d5db; font-size: 0.95rem;
    }
    .score-table th, .score-table td {border: 1px solid #d1d5db; padding: 10px 12px;}
    .score-table th {background: #f3f4f6; text-align: left;}
    .score-table td.score {text-align: center; width: 90px; font-weight: 600;}
    .score-table th.score {text-align: center; width: 90px;}
    .score-table tr.total td {background: #fff7ed; font-weight: 700;}
    .score-table tr.total td.score {font-size: 1.2rem; color: #c2410c;}
    .score-table .formula {display: block; font-size: 0.8rem; font-weight: 400; color: #6b7280;}
    .score-table .note {display: block; font-size: 0.8rem; font-weight: 400; color: #6b7280; font-style: italic;}

    .footer {text-align: center; color: #6b7280; font-size: 0.85rem; margin-top: 36px;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 3. ĐỌC DỮ LIỆU
# ============================================================
@st.cache_data
def load_data(path):
    df = pd.read_excel(path)
    df = df.iloc[:, :len(COL_NAMES)]     # đọc theo vị trí cột, không phụ thuộc tên cột
    df.columns = COL_NAMES
    df['MSV'] = df['MSV'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return df


try:
    df = load_data(FILE_PATH)
except FileNotFoundError:
    st.error(f"Lỗi: Không tìm thấy file '{FILE_PATH}'. Vui lòng kiểm tra lại tên file trên GitHub.")
    st.stop()
except Exception as e:
    st.error(f'Lỗi khi đọc file Excel: {e}')
    st.stop()


# ============================================================
# 4. HÀM XỬ LÝ
# ============================================================
def fmt(v):
    """Điểm thường: bỏ số 0 thừa, tối đa 3 chữ số thập phân. Ô trống hiện dấu '-'."""
    if pd.isna(v):
        return '-'
    v = round(float(v), 3)
    return str(int(v)) if v.is_integer() else str(v)


def fmt_tp1(v):
    """Điểm thành phần: luôn 1 chữ số thập phân, làm tròn lên khi .x5 (giống Excel)."""
    if pd.isna(v):
        return '-'
    d = Decimal(str(round(float(v), 6))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
    return str(d)


def fmt_cell(v):
    """Ô điểm có thể chứa số hoặc chữ ghi chú, ví dụ '1 (đã cộng bù 0,1đ ...)'.
    Trả về (giá trị hiển thị, ghi chú)."""
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return '-', ''
        m = re.match(r'^(\d+(?:[.,]\d+)?)\s*(.*)$', s)
        if m:
            num = float(m.group(1).replace(',', '.'))
            return fmt(num), m.group(2).strip()
        return '-', s                      # chữ thuần: không có điểm, chỉ có ghi chú
    return fmt(v), ''


def lookup_scores(mssv_input):
    mssv_input = str(mssv_input).strip()
    result = df[df['MSV'] == mssv_input]
    if result.empty:
        return None
    row = result.iloc[0]
    data = {'ten': row['Ho_ten'], 'lop': row['Lop']}
    for k in ['A', 'B', 'C', 'E', 'F', 'G', 'BTN']:
        data[k], data[k + '_note'] = fmt_cell(row[k])
    data['TP1'] = fmt_tp1(row['TP1'])
    data['TP2'] = fmt_tp1(row['TP2'])
    return data


def table_row(label, value, note='', formula='', total=False):
    extra = ''
    if formula:
        extra += f'<span class="formula">{formula}</span>'
    if note:
        extra += f'<span class="note">{html.escape(note)}</span>'
    css = ' class="total"' if total else ''
    return f'<tr{css}><td>{label}{extra}</td><td class="score">{value}</td></tr>'


def build_result_html(data, mssv):
    # Không để dòng trống trong chuỗi HTML (Markdown sẽ hiểu sai và hiển thị thành code)
    head = (
        '<div class="student-card">'
        f'<div class="name">✅ {html.escape(str(data["ten"]))}</div>'
        f'<div class="meta">MSSV: {html.escape(mssv)} &nbsp;|&nbsp; Lớp: {html.escape(str(data["lop"]))}</div>'
        '</div>'
    )
    rows = [
        table_row('Kiểm tra trắc nghiệm Chương 1 và 2 (A)', data['A'], data['A_note']),
        table_row('Thảo luận các chủ đề về E-Payment (B)', data['B'], data['B_note']),
        table_row('Điểm cộng thành phần 1 (C)', data['C'], data['C_note']),
        table_row('ĐIỂM THÀNH PHẦN 1', data['TP1'], formula='= A × 50% + B × 50% + C × 5%', total=True),
        table_row('Tự luận chương 3-4 (E)', data['E'], data['E_note']),
        table_row('Kiểm tra trắc nghiệm chương 3,4,5,6 (F)', data['F'], data['F_note']),
        table_row('Điểm cộng thành phần 2 (G)', data['G'], data['G_note']),
        table_row('ĐIỂM THÀNH PHẦN 2', data['TP2'], formula='= E × 50% + F × 50% + G × 5%', total=True),
        table_row('Bài tập nhóm lớn', data['BTN'], data['BTN_note']),
    ]
    table = (
        '<table class="score-table">'
        '<thead><tr><th>Nội dung</th><th class="score">Điểm</th></tr></thead>'
        '<tbody>' + ''.join(rows) + '</tbody></table>'
    )
    return head + table


# ============================================================
# 5. GIAO DIỆN CHÍNH
# ============================================================
st.markdown(
    f"""
    <div class="hero">
        <h1>🎓 Tra Cứu Điểm</h1>
        <div class="course">Học phần: {COURSE}</div>
        <div class="lecturer">👩‍🏫 Giảng viên: <b>{LECTURER}</b></div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.form('form_tra_cuu'):
    mssv_input = st.text_input('📝 Nhập Mã Số Sinh Viên (MSSV)', placeholder='Ví dụ: 221121302202')
    submitted = st.form_submit_button('🔍 Tra cứu điểm', type='primary', use_container_width=True)

if submitted:
    if not mssv_input.strip():
        st.warning('⚠️ Vui lòng nhập Mã Số Sinh Viên.')
    else:
        data = lookup_scores(mssv_input)
        if data is None:
            st.error(f'❌ Không tìm thấy dữ liệu cho MSSV: **{mssv_input.strip()}**. '
                     'Vui lòng kiểm tra lại.')
        else:
            st.markdown(build_result_html(data, mssv_input.strip()), unsafe_allow_html=True)

st.markdown(
    f'<div class="footer">© Giảng viên {LECTURER} – Trường Đại học Kinh tế, Đại học Đà Nẵng<br>'
    'Nếu có thắc mắc về điểm, vui lòng liên hệ trực tiếp giảng viên.</div>',
    unsafe_allow_html=True,
)
