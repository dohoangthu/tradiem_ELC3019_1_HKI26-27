from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import streamlit as st

# ============================================================
# 1. CẤU HÌNH
# ============================================================
FILE_PATH = '271_ELC3019_ELC3019_1.xlsx'   # tên file Excel (phải trùng tên file trên GitHub)
COURSE = 'ELC3019 – Thanh toán điện tử'
LECTURER = 'Đỗ Hoàng Thu'

# File Excel gồm 7 cột theo thứ tự:
# ClassID | UID | FullName | A (trắc nghiệm) | B (thảo luận) | C (điểm cộng) | Thành phần 1
COL_NAMES = ['Lop', 'MSV', 'Ho_ten', 'A', 'B', 'C', 'TP1']

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

    .score-grid {display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 8px;}
    .score-card {
        background: #ffffff; border: 1px solid #e5e7eb; border-radius: 14px;
        padding: 16px 12px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .score-card .label {font-size: 0.85rem; color: #6b7280; min-height: 2.6em;}
    .score-card .value {font-size: 2rem; font-weight: 800; color: #1d4ed8; margin-top: 4px;}

    .final-card {
        background: linear-gradient(135deg, #f59e0b 0%, #f97316 100%);
        color: #ffffff; border-radius: 16px; padding: 20px; text-align: center;
        margin-top: 14px; box-shadow: 0 6px 18px rgba(249, 115, 22, 0.3);
    }
    .final-card .label {font-size: 1rem; opacity: 0.95;}
    .final-card .value {font-size: 3rem; font-weight: 800; line-height: 1.1; margin-top: 4px;}
    .final-card .formula {font-size: 0.8rem; opacity: 0.9; margin-top: 6px;}

    .footer {text-align: center; color: #6b7280; font-size: 0.85rem; margin-top: 36px;}
    @media (max-width: 600px) {.score-grid {grid-template-columns: 1fr;}}
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
    df['MSV'] = df['MSV'].astype(str).str.strip()
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
    """Thành phần 1: luôn 1 chữ số thập phân, làm tròn lên khi .x5 (giống Excel)."""
    if pd.isna(v):
        return '-'
    d = Decimal(str(round(float(v), 6))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
    return str(d)


def lookup_scores(mssv_input):
    mssv_input = str(mssv_input).strip()
    result = df[df['MSV'] == mssv_input]
    if result.empty:
        return None
    row = result.iloc[0]
    return {
        'ten': row['Ho_ten'],
        'lop': row['Lop'],
        'A': fmt(row['A']),
        'B': fmt(row['B']),
        'C': fmt(row['C']),
        'TP1': fmt_tp1(row['TP1']),
        'TP1_raw': float(row['TP1']),
    }


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
            st.markdown(
                f"""
                <div class="student-card">
                    <div class="name">✅ {data['ten']}</div>
                    <div class="meta">MSSV: {mssv_input.strip()} &nbsp;|&nbsp; Lớp: {data['lop']}</div>
                </div>
                <div class="score-grid">
                    <div class="score-card">
                        <div class="label">Kiểm tra trắc nghiệm<br>Chương 1 và 2 (A)</div>
                        <div class="value">{data['A']}</div>
                    </div>
                    <div class="score-card">
                        <div class="label">Thảo luận các chủ đề<br>về E-Payment (B)</div>
                        <div class="value">{data['B']}</div>
                    </div>
                    <div class="score-card">
                        <div class="label">Điểm cộng<br>thành phần 1 (C)</div>
                        <div class="value">{data['C']}</div>
                    </div>
                </div>
                <div class="final-card">
                    <div class="label">⭐ ĐIỂM THÀNH PHẦN 1</div>
                    <div class="value">{data['TP1']}</div>
                    <div class="formula">= A × 50% + B × 50% + C × 5%</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.progress(min(max(data['TP1_raw'] / 10, 0.0), 1.0))

st.markdown(
    f'<div class="footer">© Giảng viên {LECTURER} – Trường Đại học Kinh tế, Đại học Đà Nẵng<br>'
    'Nếu có thắc mắc về điểm, vui lòng liên hệ trực tiếp giảng viên.</div>',
    unsafe_allow_html=True,
)