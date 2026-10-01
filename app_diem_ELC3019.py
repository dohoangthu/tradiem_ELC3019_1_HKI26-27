from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import streamlit as st

# 1. Tên file Excel (phải trùng chính xác với tên file trên GitHub)
FILE_PATH = '271_ELC3019_ELC3019_1.xlsx'
LECTURER = 'Đỗ Hoàng Thu'

# 2. Các thành phần điểm: (nhãn hiển thị, từ khóa nhận diện trong tên cột Excel)
#    Code tìm cột theo TÊN nên không bị lệch khi file Excel thêm/bớt cột.
SCORE_SPEC = [
    ('Điểm danh', 'điểm danh'),                                   # nếu file không có cột này thì tự bỏ qua
    ('Kiểm tra trắc nghiệm Chương 1 và 2 (A)', '(a)'),
    ('Thảo luận các chủ đề về E-Payment (B)', '(b)'),
    ('Điểm cộng thành phần 1 (C)', '(c)'),
    ('Thành phần 1 (A×50% + B×50% + C×5%)', '^thành phần 1'),                            # '^' = tên cột bắt đầu bằng cụm này
]
TP1_LABEL = 'Thành phần 1 (A×50% + B×50% + C×5%)'

st.set_page_config(page_title='Tra Cứu Điểm ELC3019', page_icon='🎓')


# 3. Đọc file Excel
@st.cache_data
def load_data(path):
    df = pd.read_excel(path)
    df.columns = [' '.join(str(c).split()) for c in df.columns]   # bỏ xuống dòng/khoảng trắng thừa
    return df


def find_col(df, keyword):
    """Tìm tên cột chứa từ khóa (không phân biệt hoa thường). Từ khóa bắt đầu bằng '^' = tên cột phải bắt đầu bằng cụm đó."""
    for c in df.columns:
        low = c.lower()
        if keyword.startswith('^'):
            if low.startswith(keyword[1:]):
                return c
        elif keyword in low:
            return c
    return None


try:
    df = load_data(FILE_PATH)
except FileNotFoundError:
    st.error(f"Lỗi: Không tìm thấy file '{FILE_PATH}'. Vui lòng kiểm tra lại tên file trên GitHub.")
    st.stop()
except Exception as e:
    st.error(f'Lỗi khi đọc file Excel: {e}')
    st.stop()

# Cột định danh: ClassID | UID | FullName
COL_LOP, COL_MSV, COL_TEN = df.columns[0], df.columns[1], df.columns[2]
df[COL_MSV] = df[COL_MSV].astype(str).str.strip()

# Ghép nhãn hiển thị với cột thực tế trong file
SCORE_COLS = []
for label, key in SCORE_SPEC:
    col = find_col(df, key)
    if col is not None:
        SCORE_COLS.append((label, col))


# 4. Hàm định dạng điểm
def fmt(v):
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


# 5. Hàm tra cứu
def lookup_scores(mssv_input):
    mssv_input = str(mssv_input).strip()
    result = df[df[COL_MSV] == mssv_input]
    if result.empty:
        return None
    row = result.iloc[0]
    scores = []
    for label, col in SCORE_COLS:
        scores.append((label, fmt_tp1(row[col]) if label == TP1_LABEL else fmt(row[col])))
    return {'Họ và Tên': row[COL_TEN], 'Lớp': row[COL_LOP], 'Điểm số': scores}


# 6. Giao diện
st.title('🤖 Tra Cứu Điểm ELC3019')
st.markdown('---')

st.header('Nhập Mã Số Sinh Viên (MSSV)')
mssv_input = st.text_input('MSSV của bạn:', placeholder='Ví dụ: 221121302202')

if st.button('Tra Cứu Điểm', type='primary'):
    if mssv_input:
        with st.spinner('Đang tìm kiếm...'):
            data = lookup_scores(mssv_input)

        if data:
            st.success(f'✅ Tìm thấy: **{data["Họ và Tên"]}** - Lớp **{data["Lớp"]}**')
            st.subheader('Bảng Điểm Chi Tiết')

            # Bảng điểm, in đậm dòng Thành phần 1
            lines = ['| Thành Phần | Điểm |', '|:--|:--:|']
            for label, value in data['Điểm số']:
                if label == TP1_LABEL:
                    lines.append(f'| **{label}** | **{value}** |')
                else:
                    lines.append(f'| {label} | {value} |')
            st.markdown('\n'.join(lines))
        else:
            st.error(f'❌ Không tìm thấy dữ liệu cho MSSV: **{mssv_input}**.')
    else:
        st.warning('⚠️ Vui lòng nhập Mã Số Sinh Viên.')

st.markdown('---')
st.caption(f'Giảng viên: {LECTURER}')
