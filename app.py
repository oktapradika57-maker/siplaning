import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import plotly.express as px
import io

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SiPLANING", page_icon="📝", layout="wide")
st.title("📝 SiPLANING - System Planning & Activity Tracking")

# --- KONEKSI GOOGLE SHEETS ---
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1HvgVicTWwO4RMQI6ZR3Mu3IgGicwjcLZl9mDN1auvJU/edit"
SHEET_NAME = "Record Activity"

@st.cache_resource
def init_connection():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # Pastikan file credentials.json ada di folder yang sama
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(SPREADSHEET_URL).worksheet(SHEET_NAME)
    return sheet

try:
    sheet = init_connection()
    koneksi_sukses = True
except Exception as e:
    st.error(f"Gagal terhubung ke Google Sheets. Pastikan credentials.json sudah benar dan sheet sudah di-share ke email service account. Error: {e}")
    koneksi_sukses = False

# --- TABS NAVIGASI ---
tab1, tab2 = st.tabs(["📋 Input Activity & Checklist", "📊 Dashboard Project Plan"])

# ==========================================
# TAB 1: FORM INPUT & CHECKLIST
# ==========================================
with tab1:
    st.header("Form Rencana & Checklist Pekerjaan")
    
    with st.form("form_siplaning"):
        col1, col2 = st.columns(2)
        
        with col1:
            site_id = st.text_input("Site ID")
            plan_date = st.date_input("Plan Visit Date")
            pic_visit = st.text_input("PIC Visit")
            
        with col2:
            sow_visit = st.selectbox("SOW Visit", ["CME", "TE", "MBP", "PM", "Lainnya"])
            detail_sow = st.text_area("Detail Pekerjaan / Catatan Tambahan")
            
        st.markdown("### ✅ Checklist Aktivitas Lapangan")
        st.caption("Centang pekerjaan yang sudah diselesaikan (progres akan dihitung otomatis)")
        
        col_chk1, col_chk2 = st.columns(2)
        with col_chk1:
            chk_persiapan = st.checkbox("1. Persiapan & Briefing K3")
            chk_material = st.checkbox("2. Pengecekan Material/Tools")
            chk_eksekusi = st.checkbox("3. Eksekusi Pekerjaan Utama")
        with col_chk2:
            chk_testing = st.checkbox("4. Testing & Commisioning")
            chk_clearing = st.checkbox("5. Site Clearing & Housekeeping")
            chk_dokumentasi = st.checkbox("6. Dokumentasi & Bast")
            
        submit_btn = st.form_submit_button("Simpan Activity")
        
        if submit_btn:
            if not site_id or not pic_visit:
                st.warning("Mohon isi Site ID dan PIC Visit!")
            elif koneksi_sukses:
                # Hitung Progres
                checklists = [chk_persiapan, chk_material, chk_eksekusi, chk_testing, chk_clearing, chk_dokumentasi]
                jml_selesai = sum(checklists)
                prosentase = int((jml_selesai / len(checklists)) * 100)
                
                # Format tanggal dan timestamp
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                tanggal_visit = plan_date.strftime("%Y-%m-%d")
                
                # Susun data untuk dikirim ke GSheets
                row_data = [
                    timestamp,
                    site_id,
                    tanggal_visit,
                    pic_visit,
                    sow_visit,
                    detail_sow,
                    "Selesai" if chk_persiapan else "Belum",
                    "Selesai" if chk_eksekusi else "Belum",
                    "Selesai" if chk_dokumentasi else "Belum",
                    prosentase
                ]
                
                try:
                    sheet.append_row(row_data)
                    st.success(f"Data Site {site_id} berhasil disimpan ke Spreadsheet! Progres: {prosentase}%")
                except Exception as e:
                    st.error(f"Gagal menyimpan data: {e}")

# ==========================================
# TAB 2: DASHBOARD & EXPORT PROJECT PLAN
# ==========================================
with tab2:
    st.header("Dashboard Planning & Progress Pekerjaan")
    
    if koneksi_sukses:
        # Ambil data dari GSheets
        data = sheet.get_all_records()
        
        if data:
            df = pd.DataFrame(data)
            
            # Jika header spreadsheet belum rapi, pastikan nama kolom sesuai. 
            # Asumsi kolom di GSheets: Timestamp, Site ID, Plan Visit, PIC, SOW, Detail, Persiapan, Eksekusi, Dokumentasi, Progres (%)
            
            # Metrik Cepat
            col_met1, col_met2, col_met3 = st.columns(3)
            col_met1.metric("Total Site Visit", len(df))
            col_met2.metric("Rata-rata Progres", f"{df['Progres (%)'].mean():.1f}%" if 'Progres (%)' in df.columns else "0%")
            
            site_selesai = len(df[df['Progres (%)'] == 100]) if 'Progres (%)' in df.columns else 0
            col_met3.metric("Site 100% Selesai", site_selesai)
            
            st.divider()
            
            # Tampilkan Tabel Data
            st.subheader("Data Rangkuman Aktivitas")
            st.dataframe(df, use_container_width=True)
            
            # Tombol Download ke Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Record_Activity')
            buffer.seek(0)
            
            st.download_button(
                label="📥 Download Data as Excel",
                data=buffer,
                file_name=f"SiPLANING_Export_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.divider()
            
            # Visualisasi Progres per Site menggunakan Plotly
            if 'Progres (%)' in df.columns and 'Site ID' in df.columns:
                st.subheader("Grafik Progres Penyelesaian per Site")
                
                # Mengurutkan berdasarkan Plan Visit jika ada
                if 'Plan Visit' in df.columns:
                    df = df.sort_values(by='Plan Visit')
                
                fig = px.bar(
                    df, 
                    x='Site ID', 
                    y='Progres (%)',
                    color='SOW',
                    hover_data=['PIC', 'Plan Visit'],
                    title="Persentase Penyelesaian Proyek (SOW)",
                    labels={'Progres (%)': 'Progres (%)', 'Site ID': 'Site'},
                    text='Progres (%)'
                )
                fig.update_traces(texttemplate='%{text}%', textposition='outside')
                fig.update_layout(yaxis_range=[0, 110])
                st.plotly_chart(fig, use_container_width=True)
                
        else:
            st.info("Belum ada data di Spreadsheet. Silakan input data di Tab 1.")
