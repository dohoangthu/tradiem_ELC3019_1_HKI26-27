from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import streamlit as st

# 1. Tên file Excel (phải trùng chính xác với tên file upload lên GitHub)
FILE_PATH = '271_ELC3019_ELC3019_1.xlsx'

# 2. Các cột trong file Excel (đọc theo vị trí cột, nên không bị lỗi khi tên cột có xuống dòng/khoảng trắng)
#    Thứ tự: ClassID | UID | FullName | Điểm danh | A | B | C | Thành phần 1
COL_NAMES = ['Lop', 'MSV', 'Ho_ten', 'Diem_danh', 'A', 'B', 'C', 'TP1']

# 3. Tên hiển thị cho sinh viên (theo thứ tự hiển thị trong bảng điểm)
SCORE_LABELS = {
    'Diem_danh': 'Điểm danh',
    'A': 'Kiểm tra trắc nghiệm Chương 1 và 2 (A)',
    'B': 'Thảo luận các chủ đề về E-Payment (B)',
    'C': 'Điểm cộng thành phần 1 (C)',
    'TP1': 'Thành phần 1 (= A×50% + B×50% + C×5%)',
}


# 4. Đọc file Excel (cache lại để không đọc lại mỗi lần tra cứu)
@st.cache_data
def load_data(path):
    df = pd.read_excel(path)
    df = df.iloc[:, :len(COL_NAMES)]   # chỉ lấy 8 cột đầu
    df.columns = COL_NAMES
    df['MSV'] = df['MSV'].astype(str).str.strip()
    return df


try:
    df = load_data(FILE_PATH)
except FileNotFoundError:
    st.error(f"Lỗi: Không tìm thấy file '{FILE_PATH}'.")
    st.error("Vui lòng kiểm tra lại tên file và đảm bảo file nằm cùng thư mục với app_diem.py.")
    st.stop()
except Exception as e:
    st.error(f"Lỗi khi đọc file Excel: {e}")
    st.stop()


# 5. Hàm định dạng giá trị điểm
def format_tp1(v):
    """Thành phần 1: luôn hiện 1 chữ số thập phân (làm tròn lên khi .x5, giống Excel)."""
    if pd.isna(v):
        return '-'
    d = Decimal(str(round(float(v), 6))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
    return str(d)


def format_value(v):
    if pd.isna(v):
        return '-'          # ô trống (ví dụ không có điểm cộng)
    if isinstance(v, (int, float)):
        v = round(float(v), 3)          # làm tròn tối đa 3 chữ số thập phân
        return str(int(v)) if v.is_integer() else str(v)
    return str(v)


# 6. Hàm tra cứu
def lookup_scores(mssv_input):
    mssv_input = str(mssv_input).strip()
    result = df[df['MSV'] == mssv_input]

    if result.empty:
        return None

    row = result.iloc[0]
    scores = {
        label: (format_tp1(row[col]) if col == 'TP1' else format_value(row[col]))
        for col, label in SCORE_LABELS.items()
    }

    return {
        'MSSV': mssv_input,
        'Họ và Tên': row['Ho_ten'],
        'Lớp': row['Lop'],
        'Điểm số': scores,
    }


# 7. Giao diện Streamlit
st.set_page_config(page_title="Tra Cứu Điểm ELC3019", page_icon="🎓")

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

            score_df = pd.DataFrame(list(data['Điểm số'].items()), columns=['Thành Phần', 'Điểm'])
            st.dataframe(score_df, hide_index=True, use_container_width=True)
        else:
            st.error(f'❌ Không tìm thấy dữ liệu cho MSSV: **{mssv_input}**.')
    else:
        st.warning('⚠️ Vui lòng nhập Mã Số Sinh Viên.')