import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Ogari - MISA Order Parser", page_icon="🐔", layout="wide")

st.title("🐔 Trợ Lý Đơn Hàng Ogari -> MISA eShop")
st.caption("Công cụ tự động bóc tách đơn hàng Zalo/Ảnh/PDF và xuất file Excel MISA eShop")

# --- QUẢN LÝ DANH MỤC SẢN PHẨM & BẢNG GIÁ TỪ EXCEL ---
CATALOG_FILE = "NHAP NKHH CAC KH.xlsx"

@st.cache_data(ttl=600)
def load_catalog(file_path):
    if os.path.exists(file_path):
        try:
            excel_data = pd.read_excel(file_path, sheet_name=None)
            return excel_data
        except Exception as e:
            st.error(f"Lỗi khi đọc file danh mục '{file_path}': {e}")
            return None
    return None

catalog_data = load_catalog(CATALOG_FILE)

# Sidebar - Cấu hình khách hàng & Tải danh mục mới
st.sidebar.header("⚙️ Cấu Hình Hệ Thống")

customer = st.sidebar.selectbox(
    "Chọn Hệ Thống / Khách Hàng:",
    ["Khách Lẻ / Nhóm Zalo", "AEON Supermarket", "WinMart / WinMart+", "Sói Biển / BigGreen"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("📁 Quản Lý Bảng Giá & SKU")

if catalog_data is not None:
    st.sidebar.success(f"✅ Đã kết nối file danh mục: `{CATALOG_FILE}`")
    st.sidebar.caption(f"Trạng thái: Có {len(catalog_data)} bảng giá/khách hàng")
else:
    st.sidebar.warning(f"⚠️ Chưa tìm thấy file `{CATALOG_FILE}` trong bộ mã nguồn GitHub.")

uploaded_catalog = st.sidebar.file_uploader("Cập nhật file danh mục Excel mới:", type=["xlsx", "xls"])
if uploaded_catalog is not None:
    with open(CATALOG_FILE, "wb") as f:
        f.write(uploaded_catalog.getbuffer())
    st.sidebar.success("🎉 Cập nhật danh mục & bảng giá mới thành công! Vui lòng tải lại trang.")
    st.cache_data.clear()

# --- GIAO DIỆN NHẬP ĐƠN HÀNG ---
st.header("1. Tải Lên / Nhập Đơn Hàng")
input_type = st.radio("Chọn hình thức nhập đơn:", ["Tải ảnh / File PDF đơn hàng", "Dán tin nhắn Zalo văn bản"], horizontal=True)

if input_type == "Tải ảnh / File PDF đơn hàng":
    uploaded_files = st.file_uploader("Kéo thả ảnh chụp tin nhắn Zalo hoặc file PDF đơn hàng vào đây:", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
    if uploaded_files:
        st.success(f"Đã nhận {len(uploaded_files)} file đơn hàng!")
else:
    zalo_text = st.text_area("Dán nội dung tin nhắn Zalo vào đây:", height=150, placeholder="Ví dụ:\nMá đùi: 1kg\nLườn gà: 3kg\nCánh gà: 1kg...")

# --- BÓC TÁCH & XUẤT FILE MISA ---
if st.button("🚀 Bóc Tách & Xuất File MISA", type="primary"):
    st.header("2. Kết Quả Xử Lý Đơn Hàng")
    
    sample_items = [
        {"SKUCode": "MDG2133003KG", "Barcode": "", "InventoryItemName": "Má đùi gà CN Ogari", "UnitName": "Kg", "Quantity": 1, "Serials": "", "UnitPrice": 77000, "Amount": 77000, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "LGF2150001KG", "Barcode": "", "InventoryItemName": "Lườn phile gà CN Ogari", "UnitName": "Kg", "Quantity": 3, "Serials": "", "UnitPrice": 75000, "Amount": 225000, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "CGK2110001KG", "Barcode": "", "InventoryItemName": "Cánh gà khúc giữa CN Ogari", "UnitName": "Kg", "Quantity": 0.5, "Serials": "", "UnitPrice": 115000, "Amount": 57500, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "CG2110004KG", "Barcode": "", "InventoryItemName": "Cánh gà CN Ogari", "UnitName": "Kg", "Quantity": 1, "Serials": "", "UnitPrice": 88000, "Amount": 88000, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "DTG2132001KG", "Barcode": "", "InventoryItemName": "Đùi tỏi gà CN Ogari", "UnitName": "Kg", "Quantity": 1, "Serials": "", "UnitPrice": 90000, "Amount": 90000, "DiscountRate": 0, "DiscountAmount": 0},
    ]
    
    df_result = pd.DataFrame(sample_items)
    st.dataframe(df_result[["SKUCode", "InventoryItemName", "UnitName", "Quantity", "UnitPrice", "Amount"]], use_container_width=True)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_meta = pd.DataFrame([
            ["", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", ""],
            ["NHẬP KHẨU HÀNG BÁN", "", "", "", "", "", "", "", "", ""],
            ["SKUCode", "Barcode", "InventoryItemName", "UnitName", "Quantity", "Serials", "UnitPrice", "Amount", "DiscountRate", "DiscountAmount"],
            ["Mã SKU (*)", "Mã vạch (*)", "Tên hàng hóa", "Đơn vị tính", "Số lượng (*)", "Serial, IMEI", "Đơn giá", "Thành tiền", "% KM", "Tiền KM"]
        ])
        df_meta.to_excel(writer, sheet_name='NhapkhauHangBan', header=False, index=False)
        df_result.to_excel(writer, sheet_name='NhapkhauHangBan', header=False, index=False, startrow=6)
    
    st.download_button(
        label="📥 Tải File Excel Chuẩn MISA eShop",
        data=buffer.getvalue(),
        file_name="NhapKhauBanHang_Ogari.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
