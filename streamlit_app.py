import os
import sys
import subprocess

# Ép hệ thống cài openpyxl và unidecode khi mở web lên
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

# Cấu hình giao diện rộng rãi, sạch sẽ
st.set_page_config(
    page_title="Hệ Thống Bộ Lọc SASPA", 
    page_icon="🛡️", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Thêm CSS custom giống hệt bản cũ của mày cho đẹp
st.markdown("""
    <style>
        .main { background-color: #ffffff; }
        h1, h2, h3 { color: #1e3d59; font-family: 'Helvetica Neue', Arial, sans-serif; }
        .stButton>button { background-color: #f0f2f6; color: #31333f; border-radius: 4px; border: 1px solid #d3d3d3; }
    </style>
""", unsafe_allow_html=True)

# Tiêu đề chính của web
st.title("🛡️ BỘ LỌC TỔNG HỢP HỒ SƠ & GIỜ TRỰC SASPA")
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

# --- KHU VỰC UPLOAD FILE TÁCH BIỆT (JSON) ---
st.markdown("### 📥 Tải các file JSON dữ liệu lên")
file_duty = st.file_uploader("⏳ Ném file JSON GIỜ TRỰC (DUTY) vào đây:", type=["json"])
file_vat_dung = st.file_uploader("📦 Ném file JSON VẬT DỤNG (HỒ SƠ) vào đây:", type=["json"])

ket_qua_duty = []
ket_qua_vat_dung = []

# Xử lý đọc file Giờ trực
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
        st.error(f"Lỗi đọc file Giờ Trực: {e}")

# Xử lý đọc file Vật dụng
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
        st.error(f"Lỗi đọc file Vật Dụng: {e}")

# --- HIỂN THỊ KẾT QUẢ ĐÚNG THEO GIAO DIỆN CŨ ---
if file_duty is not None or file_vat_dung is not None:
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 1. HỘP THÔNG BÁO MÀU XANH LÁ QUEN THUỘC ĐÂY RỒI MÀY ƠI
    tong_so_luong_do = 0
    if ket_qua_vat_dung:
        df_temp_vd = pd.DataFrame(ket_qua_vat_dung)
        tong_so_luong_do = df_temp_vd["Số lượng"].sum()
        
    if file_vat_dung is not None:
        st.success(f"🔥 Đã xử lý xong xuôi! Tìm thấy {tong_so_luong_do} vật dụng.")
    elif file_duty is not None:
        st.success(f"🔥 Đã xử lý xong xuôi file Giờ Trực!")

    df_duty_tong = pd.DataFrame()
    df_tong_hop = pd.DataFrame()

    # 2. HIỂN THỊ BẢNG VẬT DỤNG TO RÕ NHƯ CŨ
    if ket_qua_vat_dung:
        st.markdown("## 📋 Bảng tổng cộng nhanh cuối ngày:")
        df_chi_tiet = pd.DataFrame(ket_qua_vat_dung)
        df_tong_hop = df_chi_tiet.groupby("Vật dụng chuẩn hóa")["Số lượng"].sum().reset_index()
        df_tong_hop.columns = ["Vật dụng chuẩn hóa", "Tổng số lượng"]
        st.dataframe(df_tong_hop, use_container_width=True, hide_index=False)

    # 3. HIỂN THỊ BẢNG GIỜ TRỰC (NẰM DƯỚI)
    if ket_qua_duty:
        st.markdown("## 🕒 Bảng tổng hợp giờ trực nhân viên:")
        df_duty_Goc = pd.DataFrame(ket_qua_duty)
        df_duty_tong = df_duty_Goc.groupby("Tên Nhân Viên (Tên Máy Chủ)")["Số phút"].sum().reset_index()
        df_duty_tong["Tổng Thời Gian Trực"] = df_duty_tong["Số phút"].apply(doi_phut_thanh_chuoi)
        
        df_bnd_duty = df_duty_tong[["Tên Nhân Viên (Tên Máy Chủ)", "Tổng Thời Gian Trực"]]
        st.dataframe(df_bnd_duty, use_container_width=True, hide_index=False)

    # 4. NÚT TẢI FILE EXCEL TỔNG HỢP Ở CUỐI TRANG
    if ket_qua_vat_dung or ket_qua_duty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if ket_qua_duty:
                pd.DataFrame(ket_qua_duty).to_excel(writer, sheet_name="Chi tiết ca trực", index=False)
                df_duty_tong.to_excel(writer, sheet_name="Tổng cộng Duty", index=False)
            if ket_qua_vat_dung:
                df_chi_tiet.to_excel(writer, sheet_name="Chi tiết vật dụng", index=False)
                df_tong_hop.to_excel(writer, sheet_name="Tổng cộng Vật dụng", index=False)
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📥 TẢI FILE EXCEL TỔNG HỢP",
            data=buffer.getvalue(),
            file_name="Bao_Cao_Tong_Hop_SASPA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
