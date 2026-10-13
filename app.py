import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Ogari - MISA Order Parser", page_icon="🐔", layout="wide")

st.title("🐔 Trợ Lý Đơn Hàng Ogari -> MISA eShop")
st.caption("Công cụ tự động bóc tách đơn hàng Zalo/Ảnh/PDF và xuất file Excel MISA eShop")

# Sidebar chọn khách hàng
st.sidebar.header("⚙️ Cấu Hình Khách Hàng")
customer = st.sidebar.selectbox(
    "Chọn Hệ Thống / Khách Hàng:",
    ["Khách Lẻ / Nhóm Zalo", "AEON Supermarket", "WinMart / WinMart+", "Sói Biển / BigGreen"]
)

# Upload đơn hàng
st.header("1. Tải Lên / Nhập Đơn Hàng")
input_type = st.radio("Chọn hình thức nhập đơn:", ["Tải ảnh / File PDF đơn hàng", "Dán tin nhắn Zalo văn bản"], horizontal=True)

order_data = []

if input_type == "Tải ảnh / File PDF đơn hàng":
    uploaded_files = st.file_uploader("Kéo thả ảnh chụp tin nhắn Zalo hoặc file PDF đơn hàng vào đây:", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
    if uploaded_files:
        st.success(f"Đã nhận {len(uploaded_files)} file đơn hàng!")
else:
    zalo_text = st.text_area("Dán nội dung tin nhắn Zalo vào đây:", height=150, placeholder="Ví dụ:\nMá đùi: 1kg\nLườn gà: 3kg\nCánh gà: 1kg...")

# Nút xử lý
if st.button("🚀 Bóc Tách & Xuất File MISA", type="primary"):
    st.header("2. Kết Quả Xử Lý Đơn Hàng")
    
    # Dữ liệu mẫu mô phỏng logic AI
    sample_items = [
        {"SKUCode": "MDG2133003KG", "Barcode": "", "InventoryItemName": "Má đùi gà CN Ogari", "UnitName": "Kg", "Quantity": 1, "Serials": "", "UnitPrice": 77000, "Amount": 77000, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "LGF2150001KG", "Barcode": "", "InventoryItemName": "Lườn phile gà CN Ogari", "UnitName": "Kg", "Quantity": 3, "Serials": "", "UnitPrice": 75000, "Amount": 225000, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "CGK2110001KG", "Barcode": "", "InventoryItemName": "Cánh gà khúc giữa CN Ogari", "UnitName": "Kg", "Quantity": 0.5, "Serials": "", "UnitPrice": 115000, "Amount": 57500, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "CG2110004KG", "Barcode": "", "InventoryItemName": "Cánh gà CN Ogari", "UnitName": "Kg", "Quantity": 1, "Serials": "", "UnitPrice": 88000, "Amount": 88000, "DiscountRate": 0, "DiscountAmount": 0},
        {"SKUCode": "DTG2132001KG", "Barcode": "", "InventoryItemName": "Đùi tỏi gà CN Ogari", "UnitName": "Kg", "Quantity": 1, "Serials": "", "UnitPrice": 90000, "Amount": 90000, "DiscountRate": 0, "DiscountAmount": 0},
    ]
    
    df_result = pd.DataFrame(sample_items)
    st.dataframe(df_result[["SKUCode", "InventoryItemName", "UnitName", "Quantity", "UnitPrice", "Amount"]], use_container_width=True)
    
    # Tạo file Excel MISA đúng format
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
