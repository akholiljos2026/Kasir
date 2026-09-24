import streamlit as st
import pandas as pd
import os
from datetime import datetime

# === KONFIGURASI HALAMAN ===
st.set_page_config(page_title="Menu & Pemesanan", page_icon="🍽️", layout="wide")
FILE_MENU = "Menu_Makanan.xlsx"
FILE_PESANAN = "pesanan_konsumen.xlsx"

# === ATUR TAMPILAN & UKURAN FONT ===
st.markdown("""
    <style>
    div[class*="stRadio"] > label {
        font-size: 16px !important;
        font-weight: bold;
    }
    div[class*="stRadio"] p {
        font-size: 15px !important;
    }
    h1 { font-size: 26px !important; }
    h2 { font-size: 20px !important; }
    hr { border-color: #ddd !important; }
    </style>
""", unsafe_allow_html=True)

# === FUNGSI BACA DATA MENU ===
def load_menu():
    kolom = ["Tanggal", "Menu Makanan", "Porsi", "Stock", "Harga", "Jenis Makanan", "Gambar"]
    if os.path.exists(FILE_MENU):
        try:
            df = pd.read_excel(FILE_MENU, engine="openpyxl")
            for k in kolom:
                if k not in df.columns:
                    df[k] = ""
            return df
        except Exception as e:
            st.error(f"Gagal baca Menu: {e}")
            return pd.DataFrame(columns=kolom)
    return pd.DataFrame(columns=kolom)

# === FUNGSI SIMPAN MENU ===
def save_menu(df):
    try:
        df.to_excel(FILE_MENU, index=False, engine="openpyxl")
        return True, "✅ Tersimpan!"
    except PermissionError:
        return False, "❌ Tutup file Menu_Makanan.xlsx terlebih dahulu!"
    except Exception as e:
        return False, f"❌ {str(e)}"

# === FUNGSI BACA DATA PESANAN ===
def load_pesanan():
    kolom = ["Tanggal Pesanan", "Nama Pelanggan", "Nomor Meja", "Menu Dipesan", "Jumlah", "Harga Satuan", "Subtotal", "Total Bayar"]
    if os.path.exists(FILE_PESANAN):
        try:
            df = pd.read_excel(FILE_PESANAN, engine="openpyxl")
            return df
        except Exception as e:
            st.error(f"Gagal baca Pesanan: {e}")
            return pd.DataFrame(columns=kolom)
    return pd.DataFrame(columns=kolom)

# === FUNGSI SIMPAN PESANAN ===
def save_pesanan(data_list):
    try:
        df_lama = load_pesanan()
        df_baru = pd.DataFrame(data_list)
        df_gabung = pd.concat([df_lama, df_baru], ignore_index=True)
        df_gabung.to_excel(FILE_PESANAN, index=False, engine="openpyxl")
        return True, "✅ Pesanan tersimpan!"
    except PermissionError:
        return False, "❌ Tutup file pesanan_konsumen.xlsx terlebih dahulu!"
    except Exception as e:
        return False, f"❌ {str(e)}"

# === JUDUL UTAMA ===
st.title("🍽️ Daftar Menu & Pemesanan Makanan")

# ==================================================
# MENU NAVIGASI
# ==================================================
#st.sidebar.markdown("## 📋 Kelola Menu")
#st.sidebar.markdown("---")
#st.sidebar.markdown("## 📝 Pemesanan Makanan")

st.sidebar.markdown("## 📋 Kelola Menu")
st.sidebar.markdown("---")

pilihan = st.sidebar.radio("", [
    "📋 Tampilkan Data Menu",
    "➕ Tambah Menu",
    "✏️ Edit Menu",
    "🗑️ Hapus Menu",
    "──────────────────────",  # ← Garis pemisah di daftar
    "🛒 Pesanan Pelanggan",
    "🧾 Cetak Nota Pembayaran"
])

# Opsional: sembunyikan opsi pemisah agar tidak bisa dipilih
if pilihan == "──────────────────────":
    st.stop()


# ==================================================
# 1. TAMPILKAN DATA MENU
# ==================================================
if pilihan == "📋 Tampilkan Data Menu":
    df = load_menu()
    st.subheader("Daftar Menu")
    if df.empty:
        st.info("Belum ada data menu.")
    else:
        df_tampil = df.copy()
        df_tampil.insert(0, "No", range(1, len(df_tampil) + 1))
        st.dataframe(df_tampil, use_container_width=True, hide_index=True)
        with open(FILE_MENU, "rb") as f:
            st.download_button(
                "📥 Unduh Data Menu", f, file_name=FILE_MENU,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==================================================
# 2. TAMBAH MENU
# ==================================================
elif pilihan == "➕ Tambah Menu":
    df = load_menu()
    st.subheader("Tambah Menu Baru")
    with st.form("form_tambah", clear_on_submit=True):
        tanggal = st.date_input("Tanggal")
        menu = st.text_input("Menu Makanan")
        porsi = st.text_input("Porsi (contoh: Sedang, Gelas, Mangkok)")
        stock = st.number_input("Stock", min_value=0, step=1)
        harga = st.number_input("Harga (Rp)", min_value=0, step=500)
        jenis = st.text_input("Jenis Makanan (Makanan/Minuman)")
        gambar = st.text_input("URL Gambar (Opsional)")
        
        if st.form_submit_button("Simpan"):
            if not menu:
                st.warning("Nama menu wajib diisi!")
            else:
                baru = {
                    "Tanggal": tanggal, "Menu Makanan": menu, "Porsi": porsi,
                    "Stock": stock, "Harga": harga, "Jenis Makanan": jenis, "Gambar": gambar
                }
                df = pd.concat([df, pd.DataFrame([baru])], ignore_index=True)
                ok, pesan = save_menu(df)
                if ok:
                    st.success(pesan)
                    st.rerun()
                else:
                    st.error(pesan)

# ==================================================
# 3. EDIT MENU
# ==================================================
elif pilihan == "✏️ Edit Menu":
    df = load_menu()
    if df.empty:
        st.info("Belum ada data untuk diedit.")
    else:
        st.subheader("Edit Menu")
        pilih_menu = st.selectbox("Pilih Menu", ["-- Pilih --"] + df["Menu Makanan"].tolist())
        
        if pilih_menu != "-- Pilih --":
            idx = df[df["Menu Makanan"] == pilih_menu].index[0]
            with st.form(f"form_edit_{idx}", clear_on_submit=True):
                tgl_def = pd.to_datetime(df.loc[idx, "Tanggal"]) if pd.notna(df.loc[idx, "Tanggal"]) else None
                porsi_def = df.loc[idx, "Porsi"] if pd.notna(df.loc[idx, "Porsi"]) else ""
                stock_def = int(df.loc[idx, "Stock"]) if pd.notna(df.loc[idx, "Stock"]) else 0
                harga_def = int(df.loc[idx, "Harga"]) if pd.notna(df.loc[idx, "Harga"]) else 0
                
                tanggal = st.date_input("Tanggal", value=tgl_def)
                menu = st.text_input("Menu Makanan", value=df.loc[idx, "Menu Makanan"])
                porsi = st.text_input("Porsi", value=porsi_def)
                stock = st.number_input("Stock", min_value=0, value=stock_def, step=1)
                harga = st.number_input("Harga", min_value=0, value=harga_def, step=500)
                jenis = st.text_input("Jenis Makanan", value=df.loc[idx, "Jenis Makanan"])
                gambar = st.text_input("URL Gambar", value=df.loc[idx, "Gambar"])
                
                if st.form_submit_button("Perbarui"):
                    if not menu:
                        st.warning("Nama menu wajib diisi!")
                    else:
                        df.loc[idx] = [tanggal, menu, porsi, stock, harga, jenis, gambar]
                        ok, pesan = save_menu(df)
                        if ok:
                            st.success(pesan)
                            st.rerun()
                        else:
                            st.error(pesan)

# ==================================================
# 4. HAPUS MENU
# ==================================================
elif pilihan == "🗑️ Hapus Menu":
    df = load_menu()
    if df.empty:
        st.info("Belum ada data untuk dihapus.")
    else:
        st.subheader("Hapus Menu")
        pilih_hapus = st.selectbox("Pilih Menu", ["-- Pilih --"] + df["Menu Makanan"].tolist())
        
        if pilih_hapus != "-- Pilih --":
            if st.button(f"⚠️ Hapus {pilih_hapus}"):
                df = df[df["Menu Makanan"] != pilih_hapus].reset_index(drop=True)
                ok, pesan = save_menu(df)
                if ok:
                    st.success(pesan)
                    st.rerun()
                else:
                    st.error(pesan)

# ==================================================
# 5. PESANAN PELANGGAN
# ==================================================
elif pilihan == "🛒 Pesanan Pelanggan":
    st.subheader("🛒 Form Pemesanan Pelanggan")
    
    df_menu = load_menu()
    if df_menu.empty:
        st.warning("⚠️ Data menu belum tersedia! Silakan tambah menu terlebih dahulu.")
        st.stop()
    
    col1, col2 = st.columns(2)
    with col1:
        nama_pelanggan = st.text_input("Nama Pelanggan")
    with col2:
        nomor_meja = st.text_input("Nomor Meja")
    
    st.markdown("---")
    st.subheader("Pilih Menu yang Dipesan")
    
    daftar_menu = df_menu["Menu Makanan"].dropna().tolist()
    harga_menu = dict(zip(df_menu["Menu Makanan"], df_menu["Harga"]))
    
    if "jumlah_baris" not in st.session_state:
        st.session_state.jumlah_baris = 1
    
    baris = st.session_state.jumlah_baris
    pesanan_list = []
    total_bayar = 0
    
    for i in range(baris):
        st.markdown(f"**Pesanan {i+1}**")
        c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
        with c1:
            pilih = c1.selectbox(f"Menu {i+1}", ["-- Pilih Menu --"] + daftar_menu, key=f"menu_{i}")
        with c2:
            jumlah = c2.number_input(f"Jumlah {i+1}", min_value=1, value=1, step=1, key=f"jumlah_{i}")
        with c3:
            harga_sat = harga_menu.get(pilih, 0)
            st.write(f"Rp {harga_sat:,}")
        with c4:
            subtotal = harga_sat * jumlah
            st.write(f"Rp {subtotal:,}")
        
        if pilih != "-- Pilih Menu --":
            pesanan_list.append({
                "Menu": pilih,
                "Jumlah": jumlah,
                "Harga Satuan": harga_sat,
                "Subtotal": subtotal
            })
            total_bayar += subtotal
        st.markdown("---")
    
    if st.button("➕ Tambah Baris Pesanan"):
        st.session_state.jumlah_baris += 1
        st.rerun()
    
    st.subheader(f"💰 Total Bayar: Rp {total_bayar:,}")
    
    if st.button("✅ Simpan Pesanan"):
        if not nama_pelanggan or not nomor_meja:
            st.warning("Nama Pelanggan dan Nomor Meja wajib diisi!")
        elif not pesanan_list:
            st.warning("Silakan pilih menu yang dipesan!")
        else:
            tgl_sekarang = datetime.now().strftime("%d %B %Y, %H:%M:%S")
            data_untuk_simpan = []
            for p in pesanan_list:
                data_untuk_simpan.append({
                    "Tanggal Pesanan": tgl_sekarang,
                    "Nama Pelanggan": nama_pelanggan,
                    "Nomor Meja": nomor_meja,
                    "Menu Dipesan": p["Menu"],
                    "Jumlah": p["Jumlah"],
                    "Harga Satuan": p["Harga Satuan"],
                    "Subtotal": p["Subtotal"],
                    "Total Bayar": total_bayar
                })
            
            ok, pesan = save_pesanan(data_untuk_simpan)
            if ok:
                st.success(pesan)
                st.balloons()
                st.session_state.jumlah_baris = 1
            else:
                st.error(pesan)

# ==================================================
# 6. CETAK NOTA PEMBAYARAN — DIPERBAIKI
# ==================================================
# ==================================================
# 6. CETAK NOTA PEMBAYARAN — Versi Paling Sederhana
# ==================================================
elif pilihan == "🧾 Cetak Nota Pembayaran":
    st.subheader("🧾 Nota Pembayaran")
    
    df_pesanan = load_pesanan()
    if df_pesanan.empty:
        st.info("Belum ada data pesanan.")
        st.stop()
    
    daftar_pelanggan = df_pesanan["Nama Pelanggan"].dropna().unique().tolist()
    nama_pilih = st.selectbox("Pilih Nama Pelanggan", ["-- Pilih --"] + daftar_pelanggan)
    
    if nama_pilih != "-- Pilih --":
        data_pilih = df_pesanan[df_pesanan["Nama Pelanggan"] == nama_pilih]
        tgl_pilih = st.selectbox("Pilih Tanggal Pesanan", data_pilih["Tanggal Pesanan"].unique().tolist())
        
        nota = data_pilih[data_pilih["Tanggal Pesanan"] == tgl_pilih]
        if not nota.empty:
            meja = nota["Nomor Meja"].iloc[0]
            total = nota["Total Bayar"].iloc[0]


            total_bayar = nota['Subtotal'].sum() # menjumlahkan subtotal ====== tambahan
            
            st.markdown("---")
            
            # TAMPILAN NOTA — SEMUA DALAM SATU BLOK
            html = f"""
<div style="max-width: 450px; margin: 0 auto; border: 2px solid #333; padding: 20px; background: #fff;">
<h2 style="text-align: center; margin: 0;">NOTA PEMBAYARAN</h2>
<hr style="border: none; border-bottom: 1px dashed #999; margin: 10px 0;">

<table style="width: 100%; margin-bottom: 10px;">
<tr><td style="width: 90px; font-weight: bold;">Nama</td><td>: {nama_pilih}</td></tr>
<tr><td style="font-weight: bold;">No. Meja</td><td>: {meja}</td></tr>
<tr><td style="font-weight: bold;">Tanggal</td><td>: {tgl_pilih}</td></tr>
</table>

<hr style="border: none; border-bottom: 1px dashed #999; margin: 10px 0;">

<table style="width: 100%; border-collapse: collapse;">
<tr style="border-bottom: 2px solid #333;">
<th style="text-align: left; padding: 6px;">Menu</th>
<th style="text-align: center; padding: 6px; width: 50px;">Jml</th>
<th style="text-align: right; padding: 6px; width: 120px;">Subtotal</th>
</tr>
"""
            # Tambahkan setiap baris pesanan
            for _, row in nota.iterrows():
                html += f"""
<tr style="border-bottom: 1px dotted #ccc;">
<td style="padding: 6px;">{row['Menu Dipesan']}</td>
<td style="padding: 6px; text-align: center;">{int(row['Jumlah'])}</td>
<td style="padding: 6px; text-align: right;">Rp {row['Subtotal']:,}</td>
</tr>
"""
            # Tutup tabel dan kotak
            html += f"""
<tr style="font-weight: bold; font-size: 16px;">
<td colspan="2" style="text-align: right; padding: 10px; border-top: 2px solid #333;">TOTAL</td>
<td style="text-align: right; padding: 10px; border-top: 2px solid #333; color: red;">Rp {total_bayar:,}</td>
</tr>
</table>

<hr style="border: none; border-bottom: 1px dashed #999; margin: 15px 0;">
<p style="text-align: center;">Terima kasih atas kunjungan Anda 🙏</p>
</div>
"""
            # === INI KUNCINYA — PASTIKAN TIDAK ADA YANG TERTINGGAL ===
            st.markdown(html, unsafe_allow_html=True)
            
            st.markdown("<br>")
            
            with open(FILE_PESANAN, "rb") as f:
                st.download_button(
                    "📥 Unduh Semua Data Pesanan", f, file_name=FILE_PESANAN,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("---")
st.caption("🍽️ Program Manajemen Menu & Pemesanan Makanan")