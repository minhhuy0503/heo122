import os
import sys
import subprocess

# Ép hệ thống cài openpyxl và unidecode khi mở web
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
import json

# Giao diện rộng rãi, hiện đại
st.set_page_config(
    page_title="Hệ Thống Bộ Lọc SASPA", 
    page_icon="🛡️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Thêm chút CSS để làm đẹp giao diện, đổi màu bảng và các nút
st.markdown("""
    <style>
        .main { background-color: #f5f7f9; }
        .stTable { background-color: #ffffff; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        h1 { color: #1e3d59; font-family: 'Helvetica Neue', Arial, sans-serif; }
        h3 { color: #17b890; }
        .stButton>button { background-color: #17b890; color: white; border-radius: 6px; }
    </style>
""", unsafe_allow_html=True)

# Thanh bên trái (Sidebar) chứa thông tin hướng dẫn
with st.sidebar:
    st.markdown("## 🛡️ **SASPA MANAGEMENT**")
    st.info("Hệ thống xử lý tự động dữ liệu tách riêng từ DiscordChatExporter định dạng **JSON**.")
    st.markdown("---")
    st.markdown("### 💡 Hướng dẫn nhanh:")
    st.write("Mày có file nào thì ném vào ô đó, không nhất thiết phải ném cả 2 cùng lúc nha mày!")

# Khu vực tiêu đề chính giữa màn hình
st.title("🛡️ BỘ LỌC TỔNG HỢP HỒ SƠ & GIỜ TRỰC SASPA")
st.caption("Phiên bản nâng cấp: Hỗ trợ Tách Riêng File Giờ Trực & File Vật Dụng")
st.markdown("---")

def chuan_hoa_ten_vat_dung(text):
    if not text: return ""
    text = unidecode.unidecode(str(text).lower().strip())
    text = re.sub(r'\b(v\d+|g)\b', '', text)
    return ' '.join(text.split())

def tinh_phut_tu_chuoi(time_str):
    time_str = time_str.lower().strip()
    phut = 0
    match_ca_hai = re.search(r'(\d+)\s*h\s*(\d+)', time_str)
    if match_ca_hai:
        return int(match_ca_hai.group(1)) * 60 + int(match_ca_hai.group(2))
    match_gio = re.search(r'(\d+)\s*h', time_str)
    if match_gio:
        return int(match_gio.group(1)) * 60
    return phut

def doi_phut_thanh_chuoi(tong_phut):
    if tong_phut <= 0:
        return "0h"
    gio = tong_phut // 60
    phut = tong_phut % 60
    if phut == 0:
        return f"{gio}h"
    return f"{gio}h{phut:02d}"

# --- KHU VỰC UPLOAD FILE CHIA LÀM 2 Ô RIÊNG BIỆT ---
st.subheader("📥 Bước 1: Tải các file JSON dữ liệu lên")
col_file1, col_file2 = st.columns(2)

with col_file1:
    file_duty = st.file_uploader("⏳ Ném file JSON GIỜ TRỰC (DUTY) vào đây:", type=["json"])
with col_file2:
    file_vat_dung = st.file_uploader("📦 Ném file JSON VẬT DỤNG (HỒ SƠ) vào đây:", type=["json"])

ket_qua_duty = []
ket_qua_vat_dung = []

# Xử lý file Giờ trực nếu có
if file_duty is not None:
    try:
        data_duty = json.load(file_duty)
        messages_duty = data_duty.get('messages', [])
        for msg in messages_duty:
            noi_dung = str(msg.get('content', '')).strip()
            author_data = msg.get('author', {})
            nguoi_bao = author_data.get('nickname') or author_data.get('name') or "Ẩn danh"
            nguoi_bao = str(nguoi_bao).strip()
            
            match_duty = re.search(r'(?:tổng|tong)\s*:\s*([^\n]+)', noi_dung, re.IGNORECASE)
            if match_duty:
                chuoi_thoi_gian = match_duty.group(1).strip()
                chuoi_thoi_gian = re.sub(r'\(.*\)', '', chuoi_thoi_gian).strip()
                phut_tinh_duoc = tinh_phut_tu_chuoi(chuoi_thoi_gian)
                if phut_tinh_duoc > 0:
                    ket_qua_duty.append({
                        "Tên Nhân Viên (Tên Máy Chủ)": nguoi_bao,
                        "Số phút": phut_tinh_duoc
                    })
    except Exception as e:
        st.error(f"Lỗi đọc file Giờ Trực rồi mày ơi: {e}")

# Xử lý file Vật dụng nếu có
if file_vat_dung is not None:
    try:
        data_vd = json.load(file_vat_dung)
        messages_vd = data_vd.get('messages', [])
        for msg in messages_vd:
            noi_dung = str(msg.get('content', '')).strip()
            author_data = msg.get('author', {})
            nguoi_bao = author_data.get('nickname') or author_data.get('name') or "Ẩn danh"
            nguoi_bao = str(nguoi_bao).strip()
            
            match_ho_so = re.search(r'(?:Hành vi tàng trữ|hanh vi tang tru)\s*:\s*(.*)', noi_dung, re.IGNORECASE | re.DOTALL)
            if match_ho_so:
                doan_vat_dung = match_ho_so.group(1).strip()
                cac_dong = doan_vat_dung.split('\n')
                for dong in cac_dong:
                    dong = dong.strip()
                    if not dong: continue
                    match_item = re.match(r'^(?:x|X)?\s*(\d+)?\s*(.*)$', dong)
                    if match_item:
                        so_luong = match_item.group(1)
                        so_luong = int(so_luong) if so_luong else 1
                        ten_vat_dung_goc = match_item.group(2).strip()
                        ten_vat_dung_chuan = chuan_hoa_ten_vat_dung(ten_vat_dung_goc)
                        if ten_vat_dung_chuan:
                            ket_qua_vat_dung.append({
                                "Người báo (Tên Máy Chủ)": nguoi_bao,
                                "Vật dụng gốc": ten_vat_dung_goc,
                                "Vật dụng chuẩn hóa": ten_vat_dung_chuan,
                                "Số lượng": so_luong
                            })
    except Exception as e:
        st.error(f"Lỗi đọc file Vật Dụng rồi mày ơi: {e}")

# --- HIỂN THỊ KẾT QUẢ ---
if file_duty is not None or file_vat_dung is not None:
    st.markdown("---")
    st.subheader("📊 Bước 2: Kết quả phân tích hệ thống")

    # Tính toán số liệu tổng quan (Metrics)
    tong_nhan_vien = 0
    tong_phut_ca_phe = 0
    if ket_qua_duty:
        df_temp_duty = pd.DataFrame(ket_qua_duty)
        tong_nhan_vien = df_temp_duty["Tên Nhân Viên (Tên Máy Chủ)"].nunique()
        tong_phut_ca_phe = df_temp_duty["Số phút"].sum()
    
    tong_items = 0
    if ket_qua_vat_dung:
        df_temp_vd = pd.DataFrame(ket_qua_vat_dung)
        tong_items = df_temp_vd["Số lượng"].sum()

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(label="👥 Nhân Viên Đã Trực", value=f"{tong_nhan_vien} Đồng chí")
    with col_m2:
        st.metric(label="⏳ Tổng Thời Gian Phục Vụ", value=doi_phut_thanh_chuoi(tong_phut_ca_phe))
    with col_m3:
        st.metric(label="📦 Tổng Vật Dụng Thu Giữ", value=f"{tong_items} Tấm/Cơ giáp")

    st.markdown("<br>", unsafe_with_html=True)

    # Chia 2 cột hiển thị bảng dữ liệu
    col_b1, col_b2 = st.columns(2)

    df_duty_tong = pd.DataFrame()
    df_tong_hop = pd.DataFrame()

    with col_b1:
        if ket_qua_duty:
            st.markdown("### 🕒 BẢNG TỔNG HỢP GIỜ TRỰC")
            df_duty_Goc = pd.DataFrame(ket_qua_duty)
            df_duty_tong = df_duty_Goc.groupby("Tên Nhân Viên (Tên Máy Chủ)")["Số phút"].sum().reset_index()
            df_duty_tong["Tổng Thời Gian Trực"] = df_duty_tong["Số phút"].apply(doi_phut_thanh_chuoi)
            
            df_bnd_duty = df_duty_tong[["Tên Nhân Viên (Tên Máy Chủ)", "Tổng Thời Gian Trực"]]
            st.dataframe(df_bnd_duty, use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có file Giờ Trực hoặc không tìm thấy dữ liệu ca trực.")

    with col_b2:
        if ket_qua_vat_dung:
            st.markdown("### 📋 BẢNG TỔNG HỢP VẬT DỤNG / CƠ GIÁP")
            df_chi_tiet = pd.DataFrame(ket_qua_vat_dung)
            df_tong_hop = df_chi_tiet.groupby("Vật dụng chuẩn hóa")["Số lượng"].sum().reset_index()
            df_tong_hop.columns = ["Tên Vật Dụng / Cơ Giáp", "Tổng Số Lượng"]
            st.dataframe(df_tong_hop, use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có file Vật Dụng hoặc không tìm thấy dữ liệu hồ sơ.")

    # Nút xuất Excel chung
    if ket_qua_vat_dung or ket_qua_duty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if ket_qua_duty:
                pd.DataFrame(ket_qua_duty).to_excel(writer, sheet_name="Chi tiết ca trực", index=False)
                df_duty_tong.to_excel(writer, sheet_name="Tổng cộng Duty", index=False)
            if ket_qua_vat_dung:
                df_chi_tiet.to_excel(writer, sheet_name="Chi tiết vật dụng", index=False)
                df_tong_hop.to_excel(writer, sheet_name="Tổng cộng Vật dụng", index=False)
        
        st.markdown("---")
        st.markdown("### 📥 Bước 3: Xuất báo cáo lưu trữ")
        st.download_button(
            label="🟢 BẤM ĐỂ TẢI FILE EXCEL BÁO CÁO TỔNG HỢP (.XLSX)",
            data=buffer.getvalue(),
            file_name="Bao_Cao_Tong_Hop_SASPA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
