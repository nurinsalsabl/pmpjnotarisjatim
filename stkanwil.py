import streamlit as st
import pandas as pd
import json
from datetime import datetime
import os
import openpyxl
import pytesseract
import gspread 
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from PyPDF2 import PdfReader
import difflib
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import pdfplumber
import io
from PIL import Image


scope = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets"
]

# --- Autentikasi Google (hybrid: lokal + Streamlit Cloud) ---
creds = None

# ✅ 1. Kalau ada token lokal (buat development di laptop)
if os.path.exists("token.json"):
    # st.info("🔑 Menggunakan token lokal (token.json)")
    creds = Credentials.from_authorized_user_file("token.json", scopes=scope)

# ✅ 2. Kalau deploy di Streamlit Cloud (pakai secrets)
elif "google" in st.secrets and "token" in st.secrets["google"]:
    # st.info("☁️ Menggunakan token dari Streamlit Secrets")
    creds_data = json.loads(st.secrets["google"]["token"])
    creds = Credentials.from_authorized_user_info(creds_data, scopes=scope)

# ✅ 3. Kalau token ada tapi expired, refresh otomatis
if creds and creds.expired and creds.refresh_token:
    creds.refresh(Request())

# ✅ 4. Kalau belum ada sama sekali (misalnya pertama kali run lokal)
if not creds:
    st.warning("🔐 Token belum ada, buka login Google OAuth untuk membuat token.json (hanya di lokal).")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes=scope)
    creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token_file:
        token_file.write(creds.to_json())
    # st.success("✅ Token baru disimpan sebagai token.json!")

# --- Koneksi ke Google Sheets ---
try:
    client = gspread.authorize(creds)
    # st.success("✅ Autentikasi Google Sheets berhasil!")
except Exception as e:
    # st.error(f"❌ Gagal autentikasi Google Sheets: {e}")
    client = None

# --- Koneksi ke Google Drive ---
try:
    drive_service = build('drive', 'v3', credentials=creds)
    # st.info("✅ Terhubung ke Google Drive API.")
except Exception as e:
    # st.error(f"❌ Gagal menghubungkan ke Google Drive API: {e}")
    drive_service = None

# --- Buka Spreadsheet ---
if client:
    try:
        sh = client.open("Kuisioner PMPJ Notaris FINAL 2025")  # Ganti dengan nama sheet kamu
        worksheet = sh.sheet1
        st.success("📄 Mohon lengkapi kuisioner berikut sesuai format!")
    except gspread.SpreadsheetNotFound:
        st.error("❌ Spreadsheet tidak ditemukan. Pastikan sudah dibagikan ke akun Google.")
        sh = None
        worksheet = None
else:
    sh = None
    worksheet = None
 
# --- Dictionary Mapping ---
profil = {
    "a. Pengusaha/wiraswasta" : 9,
    "b.  PNS (termasuk pensiunan)": 4,
    "c.  Ibu Rumah Tangga" : 2,
    "d.  Pelajar/Mahasiswa" : 2,
    "e.  Pegawai Swasta" : 7,
    "f.  Pejabat Lembaga Legislatif dan Pemerintah" : 4,
    "g.  TNI/POLRI (termasuk Pensiunan)": 3,
    "h. Pegawai BI/BUMN/BUMD (termasuk Pensiunan)": 2,
    "i.  Profesional dan Konsultan" : 6,
    "j.  Pedagang" : 5,
    "k.  Pegawai Bank" : 2,
    "l. Pegawai Money Changer" : 1,
    "m. Pengajar dan Dosen" : 2,
    "n. Petani" : 1,
    "o.  Korporasi Perseroan Terbatas" : 7,
    "p.  Korporasi Koperasi" : 2,
    "q.  Korporasi Yayasan": 2,
    "r.  Korporasi CV, Firma, dan Maatschap": 2,
    "s.  Korporasi Perkumpulan Badan Hukum": 2,
    "t.  Korporasi Perkumpulan Tidak Badan Hukum": 2,
    "u.  Pengurus Parpol": 2,
    "v.  Bertindak berdasarkan Kuasa" : 2,
    "w. Lain-lain" : 1
}

bisnis_pengguna = {
    "a. Perdagangan": 9,
    "b. Pertambangan": 4,
    "c. Pertanian": 1,
    "d. Perikanan": 1,
    "e. Perkebunan": 1,
    "f. Perindustrian": 2,
    "g. Perbankan": 3,
    "h. Pembiayaan": 4,
    "i. Pembangunan Property": 3,
    "j. Kontraktor": 2,
    "k. Konsultan": 1,
    "l. Transportasi Barang dan Orang": 1,
    "m. Usaha Sewa Menyewa": 2,
    "n. Lain-lain....": 1
}

jasa = {
    "a.  Pembelian dan Penjualan Properti": 9,
    "b.  Pengurusan Perizinan Badan Usaha": 7,
    "c.  Penitipan Pembayaran Pajak terkait Pengalihan Property": 3,
    "d.  Pengurusan Pembelian dan Penjualan Badan Usaha": 3,
    "e.  Pengelolaan terhadap Uang, Efek, dan/atau Produk Jasa Keuangan lainnya": 4,
    "f.  Pengelolaan Rekening Giro, Rekening Tabungan, Rekening Deposito, dan/atau Rekening Efek": 2,
    "g.  Pengoperasian dan Pengelolaan Perusahaan": 3,
    "h. Lain-lain": 1
}

produk = {
    "a. Akta pembayaran uang sewa, bunga, dan pensiun ": 4,
    "b. Akta penawaran pembayaran tunai ": 4,
    "c.  Akta protes terhadap tidak dibayarnya atau tidak diterimanya surat berharga ": 2,
    "d. Akta Kuasa": 4,
    "e. Akta keterangan kepemilikan": 5,
    "f. Akta Hibah (Barang Bergerak)": 4,
    "g. Akta Wasiat": 2,
    "h. Akta Jaminan Fidusia ": 3,
    "i. Akta Pendirian Perseroan Terbatas ": 8,
    "j. Akta Perubahan Perseroan Terbatas  ": 5,
    "k. Akta Pendirian dan Perubahan Koperasi ": 3,
    "l. Akta Pendirian dan Perubahan Yayasan (Nirlaba) ": 3,
    "m. Akta Pendirian dan Perubahan CV, Firma dan Maatschap (Persekutuan Perdata) - Badan usaha yang tidak berbadan hukum ": 3,
    "n. Akta Pendirian dan Perubahan Perkumpulan Badan Hukum (Sosial/Nirlaba) ": 3,
    "o. Akta Pendirian dan Perubahan Perkumpulan Tidak Berbadan Hukum (Sosial/Nirlaba) ": 3,
    "p. Akta Pendirian dan Perubahan Partai Politik ": 2,
    "q. Akta Perjanjian Sewa Menyewa ": 3,
    "r. Akta Perjanjian Pengikatan Jual Beli ": 8,
    "s. Akta Perjanjian Kerjasama ": 4,
    "t. Akta Perjanjian BOT (Build Operate Transfer/Bangun Kelola Serah) ": 2,
    "u. Akta Perjanjian JO (Joint Operation/Kerjasama Operasional Mengelola Proyek) ": 2,
    "v. Akta Perjanjian Kredit ": 4,
    "w. Akta Pinjam Meminjam/Pengakuan Hutang ": 4,
    "x. Akta lainnya sesuai dengan ketentuan peraturan perundang-undangan ": 3
}

negara = {
    "a.  Tax Haven Country": 6,
    "b.  RRT (Tiongkok)": 8,
    "c.  Malaysia": 7,
    "d.  Singapura": 7,
    "e.  Asia lainnya": 8,
    "f.  Afrika": 1,
    "g.  Amerika": 5,
    "h.  Eropa": 6,
    "i.  Australia dan Selandia Baru": 5
}

apgakkum = {"YA": 6, "TIDAK": 1}

# --- Fungsi hitung inherent risk ---
def hitung_risiko(inputs):
    def pilih_terbesar(mapping_dict, user_inputs, default=None):
        if all(v == 0 for v in user_inputs.values()):
            return default, mapping_dict.get(default, 0)
        terbaik = max(user_inputs, key=user_inputs.get)
        return terbaik, mapping_dict.get(terbaik, 0)

    jawaban_profil, skor_profil   = pilih_terbesar(profil, inputs["profil"],  default="w. Lain-lain")
    jawaban_bisnis, skor_bisnis   = pilih_terbesar(bisnis_pengguna, inputs["bisnis"], default="n. Lain-lain....")
    jawaban_jasa, skor_jasa       = pilih_terbesar(jasa, inputs["jasa"],     default="h. Lain-lain")
    jawaban_negara, skor_negara   = pilih_terbesar(negara, inputs["negara"], default="e.  Asia lainnya")
    skor_apgakkum                 = apgakkum.get(inputs["apgakkum"], 0)

    total = skor_profil + skor_bisnis + skor_jasa + skor_negara + skor_apgakkum + 6

    if 6 <= total <= 17: kategori = "Rendah"
    elif 18 <= total <= 29: kategori = "Sedang"
    elif 30 <= total <= 41: kategori = "Tinggi"
    elif 42 <= total <= 52: kategori = "Sangat Tinggi"
    else: kategori = "Diluar Rentang"

    return {
        "jawaban_profil": jawaban_profil, "skor_profil": skor_profil,
        "jawaban_bisnis": jawaban_bisnis, "skor_bisnis": skor_bisnis,
        "jawaban_jasa": jawaban_jasa,     "skor_jasa": skor_jasa,
        "jawaban_negara": jawaban_negara, "skor_negara": skor_negara,
        "jawaban_apgakkum": inputs["apgakkum"], "skor_apgakkum": skor_apgakkum,
        "total_skor": total, "kategori_risiko": kategori
    }

# --- OCR PDF ---
def validasi_ocr_pdf(uploaded_file1, kata_kunci_list, judul=""):
    if uploaded_file1 is None:
        return False, "Tidak ada file.", 0

    try:
        pdf_bytes = uploaded_file1.read()
        uploaded_file1.seek(0)

        all_text = ""

        # --- 1️⃣ Ekstrak teks langsung (kalau PDF text-based) ---
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages[:5]:
                    extracted = page.extract_text()
                    if extracted:
                        all_text += extracted.lower() + "\n"
        except Exception as e:
            st.warning(f"⚠️ Gagal ekstrak teks langsung: {e}")

        # --- 2️⃣ Kalau teks kosong, fallback ke OCR (hasil scan) ---
        if not all_text.strip():
            st.info("📸 Proses scan PDF sedan berjalan...")
            try:
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    for i, page in enumerate(pdf.pages[:5]):
                        # 🔹 Render halaman PDF jadi image
                        page_image = page.to_image(resolution=200).original
                        # 🔹 OCR dari image
                        text = pytesseract.image_to_string(
                            page_image, lang="ind+eng", config="--psm 6"
                        )
                        all_text += text.lower() + "\n"
            except Exception as e:
                st.error(f"❌ Gagal OCR dari gambar PDF: {e}")
                return False, "Error OCR", 0

        # --- 3️⃣ Cek kata kunci dengan fuzzy matching ---
        variasi_kata = [
            "formulir customer due diligence perorangan",
            "analisis resiko", "analisis risiko",
            "formulir customer due diligence",
            "enhanced due diligence", "formulir customer due diligence korporasi"
        ]

        jumlah_ditemukan = 0
        for kata_utama in kata_kunci_list:
            kata_lower = kata_utama.lower()
            variasi_relevan = [v for v in variasi_kata if kata_lower in v.lower()]

            def fuzzy_found(keyword):
                panjang = len(keyword)
                for i in range(0, len(all_text) - panjang + 1):
                    potongan = all_text[i:i+panjang+3]
                    if difflib.SequenceMatcher(None, keyword, potongan).ratio() > 0.6:
                        return True
                return False

            found = (
                kata_lower in all_text
                or any(v in all_text for v in variasi_relevan)
                or fuzzy_found(kata_lower)
            )
            if found:
                jumlah_ditemukan += 1

        # --- 4️⃣ Kalau OCR gak nemu teks sama sekali ---
        if not all_text.strip():
            st.warning(f"⚠️ OCR tidak menemukan teks di file {judul}.")
            return False, "Tidak ada teks terdeteksi", 0

        return True, all_text, jumlah_ditemukan

    except Exception as e:
        st.error(f"❌ Error umum saat OCR: {e}")
        return False, "Error OCR", 0
    
# --- Fungsi hitung internal control ---
def hitung_internal_control(q1, uploaded_file1, is_valid_ocr_q1):
    if q1 == "TIDAK" or uploaded_file1 is None:
        nilai = 141  # Lemah jika q1=TIDAK atau no file
    else:  # q1 == "YA" dan file ada
        nilai = 37 if is_valid_ocr_q1 else 141  # Sangat Baik jika valid, else Lemah
    def kategori_ic(nilai):
        if 37 <= nilai <= 62: return "Sangat Baik"
        elif 63 <= nilai <= 88: return "Baik"
        elif 89 <= nilai <= 114: return "Cukup"
        elif 115 <= nilai <= 141: return "Lemah"
        return "Diluar Rentang"
    return nilai, kategori_ic(nilai)

# --- Residual Risk ---
def hitung_residual_risk(kategori_inherent, kategori_internal):
    residual_matrix = {
        "Lemah":       {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Sangat Tinggi", "Sangat Tinggi": "Sangat Tinggi"},
        "Cukup":       {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Tinggi",        "Sangat Tinggi": "Sangat Tinggi"},
        "Baik":        {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Sedang",        "Sangat Tinggi": "Tinggi"},
        "Sangat Baik": {"Rendah": "Rendah", "Sedang": "Rendah", "Tinggi": "Sedang",        "Sangat Tinggi": "Tinggi"}
    }
    risk_value = {"Rendah": 1, "Sedang": 2, "Tinggi": 3, "Sangat Tinggi": 4}
    kategori_residual = residual_matrix.get(kategori_internal, {}).get(kategori_inherent, "Sangat Tinggi")  # Default jika miss
    return kategori_residual, risk_value.get(kategori_residual, 4)

# --- Risiko Pengguna Jasa ---
def risiko_pengguna_jasa(jumlah_klien):
    if jumlah_klien <= 100: return 1, "Rendah"
    if jumlah_klien <= 200: return 2, "Sedang"
    if jumlah_klien <= 300: return 3, "Tinggi"
    return 4, "Sangat Tinggi"

# --- Final Risk Priority ---
def final_risk(df):
    risk_priority = {
        4: {1: "Tinggi", 2: "Tinggi", 3: "Sangat Tinggi", 4: "Sangat Tinggi"},
        3: {1: "Sedang", 2: "Sedang", 3: "Tinggi",       4: "Sangat Tinggi"},
        2: {1: "Rendah", 2: "Sedang", 3: "Sedang",       4: "Tinggi"},
        1: {1: "Rendah", 2: "Rendah", 3: "Sedang",       4: "Tinggi"}
    }
    df["Tingkat Risiko"] = df.apply(lambda r: risk_priority.get(r["Nilai Residual Risk"], {}).get(r["Nilai Risiko Pengguna Jasa"]), axis=1)
    return df

# --- UI Streamlit ---
st.title("📊 Kuisioner PMPJ Notaris - Kementerian Hukum Jawa Timur")

with st.form("risk_form"):
    st.subheader("Identitas Notaris")
    nama_notaris = st.text_input("1. Nama Notaris (contoh: Herman Setiawan, S.H., M.Kn)", "")
    NIK_KTP = st.text_input("NIK KTP (16 digit angka)")
    username = st.text_input("Username Akun AHU Online", "")
    nomor_HP = st.text_input("Nomor HP")
    alamat = st.text_input("Alamat Lengkap Notaris", "")
    daftar_kota = [
    "Kabupaten Bangkalan",
    "Kabupaten Banyuwangi",
    "Kabupaten Blitar",
    "Kabupaten Bojonegoro",
    "Kabupaten Bondowoso",
    "Kabupaten Gresik",
    "Kabupaten Jember",
    "Kabupaten Jombang",
    "Kabupaten Kediri",
    "Kabupaten Lamongan",
    "Kabupaten Lumajang",
    "Kabupaten Madiun",
    "Kabupaten Magetan",
    "Kabupaten Malang",
    "Kabupaten Mojokerto",
    "Kabupaten Nganjuk",
    "Kabupaten Ngawi",
    "Kabupaten Pacitan",
    "Kabupaten Pamekasan",
    "Kabupaten Pasuruan",
    "Kabupaten Ponorogo",
    "Kabupaten Probolinggo",
    "Kabupaten Sampang",
    "Kabupaten Sidoarjo",
    "Kabupaten Situbondo",
    "Kabupaten Sumenep",
    "Kabupaten Trenggalek",
    "Kabupaten Tuban",
    "Kabupaten Tulungagung",
    "Kota Batu",
    "Kota Blitar",
    "Kota Kediri",
    "Kota Madiun",
    "Kota Malang",
    "Kota Mojokerto",
    "Kota Pasuruan",
    "Kota Probolinggo",
    "Kota Surabaya"
]


# Pilihan kota/kabupaten
    kota = st.selectbox("Pilih Kota/Kabupaten", daftar_kota)

    st.subheader("Jumlah Klien Sesuai Profesi")
    inputs_profil = {k: st.number_input(k, min_value=0, value=0) for k in profil.keys()}

    jumlah_klien = sum(inputs_profil.values())

    st.subheader("Jumlah Klien Total")
    st.write(jumlah_klien)

    st.subheader("Jumlah Klien Sesuai Bisnis")
    inputs_bisnis = {k: st.number_input(k, min_value=0, value=0) for k in bisnis_pengguna.keys()}

    st.subheader("Jumlah Klien Sesuai Jasa yang Digunakan")
    inputs_jasa = {k: st.number_input(k, min_value=0, value=0) for k in jasa.keys()}

    st.subheader("Jumlah Dokumen/Produk Jasa yang Diurus Klien")
    # FIX: gunakan nama berbeda agar tidak menimpa inputs_jasa
    inputs_produk = {k: st.number_input(k, min_value=0, value=0) for k in produk.keys()}

    st.subheader("Jumlah Klien Sesuai Negara")
    inputs_negara = {k: st.number_input(k, min_value=0, value=0) for k in negara.keys()}

    st.subheader("Terkait Aparat Penegak Hukum")
    inputs_apgakkum = st.radio("Apakah Notaris pernah dipanggil atau diminta informasi oleh Aparat Penegak Hukum?", ["YA", "TIDAK"])

    st.subheader("Pertanyaan Kepatuhan Notaris")
    q1 = st.radio("1. Apakah Kantor Notaris anda memiliki mekanisme analisis risiko Pengguna Jasa? (form cdd, edd dan analisa resiko)?", ["YA", "TIDAK"])
    uploaded_file1 = st.file_uploader("Upload Dokumen Pendukung (Form CDD, EDD dan Analisa Resiko) dengan format PDF", type=["pdf"])
    if uploaded_file1 is not None:
        st.success(f"File berhasil diupload: {uploaded_file1.name}")
        
    q2 = st.radio("2.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko tinggi pencucian uang dan/atau pendanaan terorisme, termasuk PEP dan negara yang berisiko tinggi sebagaimana diatur dalam Pasal 17 PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    uploaded_file2 = st.file_uploader("Upload Dokumen Pendukung (SOP PMPJ) dengan format PDF", type=["pdf"])
    # FIX: cek variabel yang benar
    if uploaded_file2 is not None:
        st.success(f"File berhasil diupload: {uploaded_file2.name}")

    q3 = st.radio("3.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko sedang pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 7 PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    q4 = st.radio("4.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko rendah pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 16 PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    q5 = st.radio("5.  Apakah Kantor Notaris Anda memiliki kebijakan larangan untuk membuka atau memelihara hubungan usaha yang menggunakan nama fiktif?", ["YA", "TIDAK"])
    q6 = st.radio("6.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi pengguna jasa orang perseorangan sebagaimana dimaksud dalam Pasal 7 ayat (1) PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    q7 = st.radio("7.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi pengguna jasa korporasi sebagaimana dimaksud dalam Pasal 7 ayat (2) PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    q8 = st.radio("8.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi pengguna jasa perikatan lainnya (legal arrangements) sebagaimana dimaksud dalam Pasal 7 ayat (1) PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    q9 = st.radio("9.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi Beneficial Owner (Pemilik Manfaat) dari Korporasi sebagaimana dimaksud dalam Pasal 8 PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    q10 = st.radio("10.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi Beneficial Owner (Pemilik Manfaat) dari perikatan lainnya (legal arrangement) sebagaimana dimaksud dalam Pasal 9 PerMenkumham Nomor 9 Tahun 2017?", ["YA", "TIDAK"])
    q11 = st.radio("11.  Apakah Kantor Notaris Anda memliliki kebijakan bertemu langsung dengan pegguna jasa dalam rangka pengumpulan informasi pengguna jasa?", ["YA", "TIDAK"])
    q12 = st.radio("12. Apakah Kantor Notaris Anda melakukan konfirmasi atas dokumen Pengguna Jasa melalui meminta keterangan kepada pengguna jasa untuk mengetahui kebenaran formil dokumen dimaksud?", ["YA", "TIDAK"])
    q13 = st.radio("13. Apakah Kantor Notaris Anda memiliki kebijakan untuk meminta dokumen pendukung lainnya dari pihak yang berwenang dalam hal terdapat keraguan atas kebenaran formil dokumen?", ["YA", "TIDAK"])
    q14 = st.radio("14. Apakah Kantor Notaris anda melakukan pemantauan kewajaran transaksi pengguna jasa?", ["YA", "TIDAK"])
    q15 = st.radio("15. Apakah Kantor Notaris anda melakukan pencatatan transaksi pengguna jasa?", ["YA", "TIDAK"])
    q16 = st.radio("16. Apakah Kantor Notaris Anda memiliki sistem informasi mengenai identifikasi, pemantauan transaksi, dan penyediaan laporan mengenai transaksi yang dilakukan oleh pengguna jasa?", ["YA", "TIDAK"])
    q17 = st.radio("17. Apakah Kantor Notaris anda menatausahakan dokumen seluruh informasi dan dokumen transaksi pengguna jasa dan beneficial owner (pemilik manfaat)?", ["YA", "TIDAK"])
    q18 = st.radio("18. Apakah Kantor Notaris anda menatausahakan dokumen seluruh informasi dan dokumen pengguna jasa dan beneficial owner (pemilik manfaat) yang diperoleh dalam rangka penerapan prinsip mengenali pengguna jasa?", ["YA", "TIDAK"])
    q19 = st.radio("19. Apakah Kantor Notaris anda menatausahakan dokumen analisis kewajaran atas transaksi pengguna jasa dan beneficial owner (pemilik manfaat)?", ["YA", "TIDAK"])
    q20 = st.radio("20. Apakah Kantor Notaris anda akan memutuskan hubungan usaha dengan pengguna jasa jika Pengguna Jasa menolak untuk mematuhi prinsip mengenali Pengguna Jasa?", ["YA", "TIDAK"])
    q21 = st.radio("21. Apakah Kantor Notaris anda akan memutuskan hubungan usaha dengan pengguna jasa jika Notaris meragukan kebenaran informasi yang disampaikan oleh Pengguna Jasa?", ["YA", "TIDAK"])
    q22 = st.radio("22. Apakah Kantor Notaris anda akan melaporkan kepada PPATK mengenai tindakan sebagaimana dimaksud pada pertanyaan nomor 20 dan nomor 21 di atas?", ["YA", "TIDAK"])
    q23 = st.radio("23. Apakah Kantor Notaris anda melakukan upaya pemutakhiran informasi dan dokumen pengguna jasa dalam hal terdapat perubahan yang diketahui oleh Notaris yang bersumber dari Pengguna Jasa yang sama atau informasi lain yang dapat dipertanggungjawabkan?", ["YA", "TIDAK"])
    q24 = st.radio("24. Apakah Kantor Notaris Anda mendokumentasikan hasil pemutakhiran informasi dan/atau dokumen sebagaimana dimaksud dalam pertanyaan Nomor 23?", ["YA", "TIDAK"])
    q25 = st.radio("25. Apakah Kantor Notaris Anda melakukan pengendalian internal melalui pelaksanaan pemeriksaan berkala terhadap penerapan prinsip mengenali Pengguna Jasa?", ["YA", "TIDAK"])
    q26 = st.radio("26. Apakah Kantor Notaris Anda melakukan pengendalian internal melalui Pemutakhiran daftar Pengguna Jasa atau pemberi kuasa yang memenuhi kriteria berisiko tinggi?", ["YA", "TIDAK"])
    q27 = st.radio("27. Apakah Kantor Notaris Anda melakukan prosedur penyaringan untuk penerimaan karyawan baru (pre-employee screening)?", ["YA", "TIDAK"])
    q28 = st.radio("28. Apakah Kantor Notaris Anda melakukan pengenalan dan pemantauan terhadap profil karyawan?", ["YA", "TIDAK"])
    q29 = st.radio("29. Apakah Kantor Notaris anda melakukan sosialisasi atau pelatihan mengenai penerapan peraturan perundang-undangan yang terkait dengan prinsip mengenali Pengguna Jasa? (baik diselenggarakan secara mandiri atau instansi terkait)", ["YA", "TIDAK"])
    q30 = st.radio("30. Apakah Kantor Notaris anda melakukan sosialisasi atau pelatihan mengenai teknik, metode, dan tipologi pencucian uang dan/atau pendanaan terorisme? (baik diselenggarakan secara mandiri atau instansi terkait)", ["YA", "TIDAK"])
    q31 = st.radio("31. Apakah Kantor Notaris anda melakukan sosialisasi atau pelatihan mengenai kebijakan dan prosedur penerapan prinsip mengenali Pengguna Jasa serta peran dan tanggung jawab pegawai dalam mencegah dan memberantas pencucian uang dan/atau pendanaan terorisme? (baik diselenggarakan secara mandiri atau instansi terkait)", ["YA", "TIDAK"])
    q32 = st.radio("32. Apakah Kantor Notaris Anda memanfaatkan teknologi baru dalam memberikan pelayanan kepada pengguna jasa?", ["YA", "TIDAK"])
    q33 = st.radio("33. Apabila jawaban Nomor 32 adalah iya, apakah Kantor Notaris Anda melakukan pengidentifikasian dan pengukuran mengenai risiko terjadinya tindak pidana pencucian uang dan pendanaan terorisme, sebelum pemanfaatan atau pengembangan teknologi baru tersebut pada pertanyaan Nomor 32?", ["YA", "TIDAK"])
    q34 = st.radio("34.  Apakah Kantor Notaris anda pernah melakukan kerja sama dengan penegak hukum dan otoritas yang berwenang untuk memberantas tindak pidana pencucian uang dan pendanaan terorisme?", ["YA", "TIDAK"])

    submitted = st.form_submit_button("Submit")

# ------------------------- Validasi & Hitung -------------------------
if submitted:
    required_fields = [nama_notaris, NIK_KTP, username, nomor_HP, alamat, kota, q1, q2, q34]
    missing = any(f is None or f == "" for f in required_fields)

    if missing:
        st.error("⚠️ Semua data wajib diisi (kecuali dokumen pendukung).")
    elif not NIK_KTP.isdigit() or not nomor_HP.isdigit():
        st.error("⚠️ NIK KTP dan Nomor HP harus berupa angka.")
    elif len(NIK_KTP) != 16:
        st.error("⚠️ NIK KTP harus 16 digit.")
    else:
        kata_kunci_list = [
            "Formulir Customer Due Diligence",
            "formulir customer due diligence perorangan",
            "Analisis Risiko", "Analisis Resiko",
            "Enhanced Due Diligence",
            "CDD",
            "EDD"]
        is_valid_ocr_q1, teks_ocr_q1, jumlah_kata_ditemukan_q1 = validasi_ocr_pdf(
            uploaded_file1, kata_kunci_list, judul="Dokumen Q1 (CDD/EDD/Analisis Risiko)"
        )

        # Simpan file pendukung (jika ada) & catat path
        # --- Simpan file & upload ke Google Drive ---
        os.makedirs("uploads", exist_ok=True)
        doc1_path, doc2_path = "", ""
        file_link_1, file_link_2 = "", ""

        def upload_to_drive(local_path, original_name):
            """Upload file ke Google Drive dan return link publik"""
            if drive_service is None:
                return local_path  # fallback: pakai path lokal
            try:
                FOLDER_ID = "1v0HSHab3hTRLPBDX4Sk5SzfHay2-rG8N" 

                file_metadata = {
                    "name": original_name,
                    "parents": [FOLDER_ID]
                }
                media = MediaFileUpload(local_path, mimetype="application/pdf")
                uploaded = drive_service.files().create(
                    body=file_metadata, media_body=media, fields="id"
                ).execute()

                file_id = uploaded.get("id")

                # Ubah jadi publik (bisa diakses semua orang dengan link)
                drive_service.permissions().create(
                    fileId=file_id,
                    body={"type": "anyone", "role": "reader"}
                ).execute()

                # Buat link view
                file_link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"
                return file_link

            except Exception as e:
                st.error(f"❌ Gagal upload ke Google Drive: {e}")
                return local_path  # fallback: simpan path lokal

        # Simpan & upload file 1
        if uploaded_file1 is not None:
            filename_1 = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_doc1_{uploaded_file1.name}"
            doc1_path = os.path.join("uploads", filename_1)
            with open(doc1_path, "wb") as f:
                f.write(uploaded_file1.getbuffer())
            file_link_1 = upload_to_drive(doc1_path, uploaded_file1.name)

        # Simpan & upload file 2
        if uploaded_file2 is not None:
            filename_2 = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_doc2_{uploaded_file2.name}"
            doc2_path = os.path.join("uploads", filename_2)
            with open(doc2_path, "wb") as f:
                f.write(uploaded_file2.getbuffer())
            file_link_2 = upload_to_drive(doc2_path, uploaded_file2.name)

        # Hitung risiko
        hasil_inherent = hitung_risiko({
            "profil": inputs_profil,
            "bisnis": inputs_bisnis,
            "jasa": inputs_jasa,
            "negara": inputs_negara,
            "apgakkum": inputs_apgakkum
        })
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        nilai_ic, kategori_ic = hitung_internal_control(q1, uploaded_file1, is_valid_ocr_q1)
        kategori_residual, nilai_residual = hitung_residual_risk(
            hasil_inherent["kategori_risiko"], kategori_ic
        )
        nilai_pengguna, kategori_pengguna = risiko_pengguna_jasa(jumlah_klien)

        df_temp = pd.DataFrame([{
            "Nilai Residual Risk": nilai_residual,
            "Nilai Risiko Pengguna Jasa": nilai_pengguna
        }])
        kategori_final = final_risk(df_temp).loc[0, "Tingkat Risiko"]

        # --- Susun baris lengkap ---
        # Bagian identitas
        data = {
            "Timestamp": timestamp,
            "Nama Notaris": nama_notaris.title(),
            "NIK KTP": NIK_KTP,
            "Username Akun AHU Online": username,
            "Nomor HP": nomor_HP,
            "2. Alamat Lengkap Notaris": alamat,
            "Kota/Kabupaten": kota,
            "3. Jumlah Klien Tahun 2024-2025": jumlah_klien,
        }

        # Rinci: profil, bisnis, jasa, produk, negara
        data.update({k: inputs_profil.get(k, 0) for k in profil.keys()})
        data.update({k: inputs_bisnis.get(k, 0) for k in bisnis_pengguna.keys()})
        data.update({k: inputs_jasa.get(k, 0) for k in jasa.keys()})
        data.update({k: inputs_produk.get(k, 0) for k in produk.keys()})
        data.update({k: inputs_negara.get(k, 0) for k in negara.keys()})

        # 34 pertanyaan + dokumen pendukung
        q_list = [q1,q2,q3,q4,q5,q6,q7,q8,q9,q10,q11,q12,q13,q14,q15,q16,q17,q18,q19,q20,q21,q22,q23,q24,q25,q26,q27,q28,q29,q30,q31,q32,q33,q34]
        q_cols = [
            "1.  Apakah Kantor Notaris anda memiliki mekanisme analisis risiko Pengguna Jasa? (form cdd, edd dan analisa resiko)",
            "2.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko tinggi ... Pasal 17 PerMenkumham 9/2017?",
            "3.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko sedang pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 7 PerMenkumham Nomor 9 Tahun 2017?",
            "4.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko rendah pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 16 PerMenkumham Nomor 9 Tahun 2017?",
            "5.  Kebijakan larangan nama fiktif?",
            "6.  Pengumpulan informasi OP (Pasal 7 ayat 1)?",
            "7.  Pengumpulan informasi Korporasi (Pasal 7 ayat 2)?",
            "8.  Pengumpulan informasi legal arrangements?",
            "9.  Pengumpulan BO Korporasi (Pasal 8)?",
            "10. Pengumpulan BO legal arrangement (Pasal 9)?",
            "11. Kebijakan bertemu langsung dengan pengguna jasa?",
            "12. Konfirmasi kebenaran formil dokumen ke pengguna jasa?",
            "13. Meminta dokumen pendukung dari otoritas berwenang bila ragu?",
            "14. Pemantauan kewajaran transaksi?",
            "15. Pencatatan transaksi pengguna jasa?",
            "16. Sistem informasi identifikasi/pemantauan/laporan transaksi?",
            "17. Penatausahaan dokumen transaksi & BO?",
            "18. Penatausahaan dokumen PJP & BO (prinsip PJP)?",
            "19. Penatausahaan dokumen analisis kewajaran transaksi?",
            "20. Putus hubungan bila PJP menolak prinsip PJP?",
            "21. Putus hubungan bila info PJP diragukan?",
            "22. Laporkan ke PPATK atas tindakan di no.20 & no.21?",
            "23. Pemutakhiran informasi & dokumen PJP bila ada perubahan?",
            "24. Dokumentasi hasil pemutakhiran?",
            "25. Pemeriksaan berkala penerapan prinsip PJP?",
            "26. Pemutakhiran daftar PJP/kuasa berisiko tinggi?",
            "27. Screening penerimaan karyawan (pre-employee)?",
            "28. Pengenalan & pemantauan profil karyawan?",
            "29. Sosialisasi/perlatihan regulasi PJP?",
            "30. Sosialisasi/perlatihan tipologi TPPU/TPPT?",
            "31. Sosialisasi/perlatihan kebijakan & prosedur PJP?",
            "32. Pemanfaatan teknologi baru?",
            "33. Identifikasi & pengukuran risiko sebelum adopsi teknologi baru?",
            "34. Kerja sama dengan penegak hukum & otoritas berwenang?"
        ]
        for i, col in enumerate(q_cols):
            data[col] = q_list[i]
        data["Dokumen_Pendukung (Q1)"] = file_link_1
        data["Dokumen Pendukung (SOP PMPJ) (Q2)"] = file_link_2

        # APGAKKUM + skor pilihan terbesar
        data["Apakah Notaris pernah dipanggil atau diminta informasi oleh Aparat Penegak Hukum?"] = inputs_apgakkum
        data["jawaban_profil"]   = hasil_inherent["jawaban_profil"]
        data["skor_profil"]      = hasil_inherent["skor_profil"]
        data["jawaban_bisnis"]   = hasil_inherent["jawaban_bisnis"]
        data["skor_bisnis"]      = hasil_inherent["skor_bisnis"]
        data["jawaban_jasa"]     = hasil_inherent["jawaban_jasa"]
        data["skor_jasa"]        = hasil_inherent["skor_jasa"]
        data["jawaban_negara"]   = hasil_inherent["jawaban_negara"]
        data["skor_negara"]      = hasil_inherent["skor_negara"]
        data["jawaban_apgakkum"] = hasil_inherent["jawaban_apgakkum"]
        data["skor_apgakkum"]    = hasil_inherent["skor_apgakkum"]

        # Ringkasan skor/tingkat risiko
        data["Nilai Inherent Risk"]     = hasil_inherent["total_skor"]
        data["Tingkat Inherent Risk"]   = hasil_inherent["kategori_risiko"]
        data["Nilai Internal Control"]  = nilai_ic
        data["Tingkat Internal Control"]= kategori_ic
        data["Tingkat Residual Risk"]   = kategori_residual
        data["Nilai Residual Risk"]     = nilai_residual
        data["Nilai Risiko Pengguna Jasa"]   = nilai_pengguna
        data["Tingkat Risiko Pengguna Jasa"] = kategori_pengguna
        data["Tingkat Risiko"]               = kategori_final

        # --- Urutan kolom (opsional): taruh kolom identitas & ringkasan dulu, sisanya mengikuti ---
        ident_cols = [
            "Timestamp","Nama Notaris","NIK KTP","Username Akun AHU Online","Nomor HP",
            "2. Alamat Lengkap Notaris","Kota/Kabupaten","3. Jumlah Klien Tahun 2024-2025"
        ]
        ringkasan_cols = [
            "Nilai Inherent Risk","Tingkat Inherent Risk",
            "Nilai Internal Control","Tingkat Internal Control",
            "Tingkat Residual Risk","Nilai Residual Risk",
            "Nilai Risiko Pengguna Jasa","Tingkat Risiko Pengguna Jasa","Tingkat Risiko"
        ]
        q_cols_with_docs = q_cols + ["Dokumen_Pendukung (Q1)","Dokumen Pendukung (SOP PMPJ) (Q2)"]
        detail_cols = list(profil.keys()) + list(bisnis_pengguna.keys()) + list(jasa.keys()) + list(produk.keys()) + list(negara.keys())
        pilihan_cols = ["Apakah Notaris pernah dipanggil atau diminta informasi oleh Aparat Penegak Hukum?",
                        "jawaban_profil","skor_profil","jawaban_bisnis","skor_bisnis",
                        "jawaban_jasa","skor_jasa","jawaban_negara","skor_negara",
                        "jawaban_apgakkum","skor_apgakkum"]

        column_order = ident_cols + detail_cols + q_cols_with_docs + pilihan_cols + ringkasan_cols

        # --- Simpan ke Excel ---
        
    # Helper: konversi nomor kolom ke huruf Excel (A, B, ..., AA, AB, ...)
    def colnum_to_excel(n: int) -> str:
        """Convert nomor kolom (1-based) ke huruf Excel (A, B, ..., AA, AB, ...)."""
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    if client is None:
        st.error("Gagal autentikasi Google Sheets. Data tidak disimpan.")
    else:
        SPREADSHEET_NAME = "Hasil Penilaian Risiko"
        SPREADSHEET_ID = "1sSzjDwgmqO6YhOGzSk4kqOybdlzAEXagWW36r3nY1OM"  
        try:
            try:
                sh = client.open_by_key(SPREADSHEET_ID)
            except gspread.SpreadsheetNotFound:
                sh = client.open(SPREADSHEET_NAME)

            worksheet = sh.sheet1  # sheet pertama

            # Ambil semua data sheet
            records = worksheet.get_all_records()
            existing = pd.DataFrame(records)

            # Kalau kosong, siapkan header
            if existing.empty:
                existing = pd.DataFrame(columns=column_order)

            existing = existing.reindex(columns=column_order)

            # Data baru
            row_df = pd.DataFrame([data]).reindex(columns=column_order)

            # Cek duplikat (Nama + NIK)
            nama_baru = data.get("Nama Notaris", "")
            nama_baru_lower = str(nama_baru).lower() if nama_baru else ""
            nik_baru = data.get("NIK KTP", "")

            if nama_baru_lower and nik_baru and not existing.empty:
                existing_nama_lower = existing["Nama Notaris"].astype(str).str.lower()
                existing_nik_str = existing["NIK KTP"].astype(str)
                mask_duplikat = (existing_nama_lower == nama_baru_lower) & (existing_nik_str == str(nik_baru))
            else:
                mask_duplikat = pd.Series([False]*len(existing))

            if mask_duplikat.any():
                existing_filtered = existing[~mask_duplikat].copy()
                df_all = pd.concat([existing_filtered, row_df], ignore_index=True)
                st.warning(f"⚠️ Data lama untuk '{nama_baru}' (NIK: {nik_baru}) ditemukan dan diganti.")
            else:
                df_all = pd.concat([existing, row_df], ignore_index=True)
                st.info("✅ Data baru ditambahkan.")

            # Clear & update sheet
            worksheet.clear()

            header = list(column_order)
            values = df_all.fillna("").astype(str).values.tolist()
            data_to_write = [header] + values

            resp = worksheet.update("A1", data_to_write)

            # Tangani hasil resp
            if isinstance(resp, dict):
                st.success(f"✅ Data berhasil disimpan")
            # else:
            #     st.success("✅ Data tersimpan ke Google Spreadsheet.")
            #     st.write("Response:", resp)   # debug

        except Exception as e:
            import traceback
            st.error(f"❌ Error saat menyimpan:\n{traceback.format_exc()}")



