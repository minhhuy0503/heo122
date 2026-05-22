import os
import sys
import subprocess

# Ép hệ thống phải tự cài unidecode và openpyxl ngay khi vừa mở web lên
try:
    import unidecode
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "unidecode", "openpyxl"])
    import unidecode
    import openpyxl

import streamlit as st
import pandas as pd
import re
import io

# Cấu hình trang web giao diện điện thoại cho đẹp
st.set_page_config(page_title="Bộ Lọc Hồ Sơ SASPA", page_icon="📊", layout="centered")

st.title("📊 Bộ Lọc & Tổng Hợp Hồ Sơ")
st.write("Được thiết kế riêng để xử lý báo cáo từ DiscordChatExporter.")
st.markdown("---")

def chuan_hoa_ten(text):
    if not text: return ""
    text = unidecode.unidecode(str(text).lower().strip())
    text = re.sub(r'\b(v\d+|g)\b', '', text)
    return ' '.join(text.split())

# Nút upload file ngay trên màn hình
uploaded_file = st.file_uploader("Ném file CSV từ Discord vào đây nha mày:", type=["csv"])

if uploaded_file is not None:
    try:
        # Đọc file CSV
        df_goc = pd.read_csv(uploaded_file)
        
        # Kiểm tra xem có đúng cột Content (nội dung tin nhắn) không
        if 'Content' not in df_goc.columns:
            st.error("File CSV này lạ quá, không thấy cột 'Content' chứa tin nhắn đâu hết!")
        else:
            ket_qua_chi_tiet = []

            for index, row in df_goc.iterrows():
                noi_dung = str(row['Content']).strip()
                # Lấy tên người báo, nếu không có thì để N/A
                nguoi_bao = row['Author'] if 'Author' in df_goc.columns else "Ẩn danh"
                
                # Tìm đoạn văn bản sau chữ "Hành vi tàng trữ :"
                match = re.search(r'(?:Hành vi tàng trữ|hanh vi tang tru)\s*:\s*(.*)', noi_dung, re.IGNORECASE | re.DOTALL)
                
                if match:
                    doan_vat_dung = match.group(1).strip()
                    cac_dong = doan_vat_dung.split('\n')
                    
                    for dong in cac_dong:
                        dong = dong.strip()
                        if not dong: continue
                        
                        # Tách số lượng (x2, x1, 2) và tên vật dụng
                        match_item = re.match(r'^(?:x|X)?\s*(\d+)?\s*(.*)$', dong)
                        
                        if match_item:
                            so_luong = match_item.group(1)
                            so_luong = int(so_luong) if so_luong else 1
                            
                            ten_vat_dung_goc = match_item.group(2).strip()
                            ten_vat_dung_chuan = chuan_hoa_ten(ten_vat_dung_goc)
                            
                            if ten_vat_dung_chuan:
                                ket_qua_chi_tiet.append({
                                    "Người báo": nguoi_bao,
                                    "Vật dụng gốc": ten_vat_dung_goc,
                                    "Vật dụng chuẩn hóa": ten_vat_dung_chuan,
                                    "Số lượng": so_luong
                                })

            if ket_qua_chi_tiet:
                df_chi_tiet = pd.DataFrame(ket_qua_chi_tiet)
                
                # Gom nhóm tính tổng số lượng
                df_tong_hop = df_chi_tiet.groupby("Vật dụng chuẩn hóa")["Số lượng"].sum().reset_index(name="Tổng số lượng")
                
                st.success(f"🔥 Đã xử lý xong xuôi! Tìm thấy {len(df_chi_tiet)} vật dụng.")
                
                # Hiển thị bản xem trước nhanh ngay trên giao diện web
                st.subheader("📋 Bảng tổng cộng nhanh cuối ngày:")
                st.dataframe(df_tong_hop, use_container_width=True)
                
                # Xuất dữ liệu ra file Excel lưu vào bộ nhớ đệm (buffer)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_chi_tiet.to_excel(writer, sheet_name="Chi tiết vật dụng", index=False)
                    df_tong_hop.to_excel(writer, sheet_name="Tổng cộng cuối ngày", index=False)
                
                st.markdown("---")
                # Nút bấm tải file về máy
                st.download_button(
                    label="📥 TẢI FILE EXCEL TỔNG HỢP",
                    data=buffer.getvalue(),
                    file_name="Tong_Hop_Ho_So_Cuoi_Ngay.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Ủa quét hết file rồi mà không thấy đứa nào ghi đúng cú pháp 'Hành vi tàng trữ :' hết mày ơi!")
    except Exception as e:
        st.error(f"Có lỗi xảy ra rồi: {e}")