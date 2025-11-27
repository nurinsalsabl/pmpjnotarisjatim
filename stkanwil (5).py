# streamlit_app.py
import streamlit as st
import pandas as pd
import json
from datetime import datetime
import os
import io
import pdfplumber
import pytesseract
from PIL import Image
import difflib

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# -----------------------
# CONFIG: scopes & keys
# -----------------------
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
]

# -----------------------
# AUTH: SERVICE ACCOUNT
# -----------------------
def get_google_clients():
    """Return (gspread_client, drive_service) using service account from st.secrets."""
    if "gcp_service_account" not in st.secrets:
        st.error("❌ Service account credentials tidak ditemukan di Streamlit secrets. "
                 "Tambahkan `gcp_service_account` dengan isi JSON key.")
        return None, None

    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=SCOPES
        )
        gspread_client = gspread.authorize(creds)
        drive_service = build("drive", "v3", credentials=creds)
        return gspread_client, drive_service
    except Exception as e:
        st.error(f"❌ Gagal autentikasi Service Account: {e}")
        return None, None

client, drive_service = get_google_clients()

# -----------------------
# CONSTANTS: spreadsheet
# -----------------------
SPREADSHEET_NAME = "Hasil Penilaian Risiko"
SPREADSHEET_ID = "1sSzjDwgmqO6YhOGzSk4kqOybdlzAEXagWW36r3nY1OM"  # jika mau pakai key langsung

# -----------------------
# Lookup dictionaries (tetap dari kode lama)
# -----------------------
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

wilayah_skor = {
    "DKI Jakarta": 9,
    "Jawa Barat": 6,
    "Jawa Timur": 6,
    "Aceh": 5,
    "Jawa Tengah": 4,
    "Kalimantan Timur": 4,
    "Banten": 3,
    "Kepulauan Riau": 3,
    "Lampung": 3,
    "Sulawasi Selatan": 3,
    "Sumatera Utara": 3,
    "Sulawasi Tenggara": 3,
    "Sulawesi Utara": 3,
    "Sumatera Selatan": 3,
    "DI Yogyakarta": 3,
    "Bali": 2,
    "Riau": 2,
    "Bangka Belitung": 2,
    "Bengkulu": 2,
    "Kalimantan Tengah": 2,
    "Maluku Utara": 2,
    "Nusa Tenggara Timur": 2,
    "Papua": 2,
    "Sulawesi Barat": 2,
    "Sulawesi Tengah": 2,
    "Gorontalo": 2,
    "Jambi": 2,
    "Kalimantan Selatan": 2,
    "Maluku": 2,
    "Nusa Tenggara Barat": 2,
    "Papua Barat": 2,
    "Sumatera Barat": 2,
    "Kalimantan Barat": 1,
    "Kalimantan Utara": 1
}

# -----------------------
# RISK FUNCTIONS (dari kode lama, cuma rapikan)
# -----------------------
def pilih_terbesar(mapping_dict, user_inputs, default=None):
    if all(v == 0 for v in user_inputs.values()):
        return default, mapping_dict.get(default, 0)
    terbaik = max(user_inputs, key=user_inputs.get)
    return terbaik, mapping_dict.get(terbaik, 0)

def hitung_risiko(inputs):
    jawaban_profil, skor_profil   = pilih_terbesar(profil, inputs["profil"],  default="w. Lain-lain")
    jawaban_bisnis, skor_bisnis   = pilih_terbesar(bisnis_pengguna, inputs["bisnis"], default="n. Lain-lain....")
    jawaban_jasa, skor_jasa       = pilih_terbesar(jasa, inputs["jasa"],     default="h. Lain-lain")
    jawaban_negara, skor_negara   = pilih_terbesar(negara, inputs["negara"], default="e.  Asia lainnya")
    skor_apgakkum                 = apgakkum.get(inputs["apgakkum"], 0)
    jawaban_wilayah = inputs["wilayah"]
    skor_wilayah = wilayah_skor.get(jawaban_wilayah, 0)

    total = skor_profil + skor_bisnis + skor_jasa + skor_negara + skor_apgakkum + skor_wilayah

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
        "jawaban_wilayah" : jawaban_wilayah, "skor_wilayah": skor_wilayah,
        "total_skor": total, "kategori_risiko": kategori
    }

def hitung_internal_control(q1, uploaded_file1, is_valid_ocr_q1):
    if q1 == "TIDAK" or uploaded_file1 is None:
        nilai = 141
    else:
        nilai = 37 if is_valid_ocr_q1 else 141
    def kategori_ic(nilai):
        if 37 <= nilai <= 62: return "Sangat Baik"
        elif 63 <= nilai <= 88: return "Baik"
        elif 89 <= nilai <= 114: return "Cukup"
        elif 115 <= nilai <= 141: return "Lemah"
        return "Diluar Rentang"
    return nilai, kategori_ic(nilai)

def hitung_residual_risk(kategori_inherent, kategori_internal):
    residual_matrix = {
        "Lemah":       {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Sangat Tinggi", "Sangat Tinggi": "Sangat Tinggi"},
        "Cukup":       {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Tinggi",        "Sangat Tinggi": "Sangat Tinggi"},
        "Baik":        {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Sedang",        "Sangat Tinggi": "Tinggi"},
        "Sangat Baik": {"Rendah": "Rendah", "Sedang": "Rendah", "Tinggi": "Sedang",        "Sangat Tinggi": "Tinggi"}
    }
    risk_value = {"Rendah": 1, "Sedang": 2, "Tinggi": 3, "Sangat Tinggi": 4}
    kategori_residual = residual_matrix.get(kategori_internal, {}).get(kategori_inherent, "Sangat Tinggi")
    return kategori_residual, risk_value.get(kategori_residual, 4)

def risiko_pengguna_jasa(jumlah_klien):
    if jumlah_klien <= 100: return 1, "Rendah"
    if jumlah_klien <= 200: return 2, "Sedang"
    if jumlah_klien <= 300: return 3, "Tinggi"
    return 4, "Sangat Tinggi"

def final_risk(df):
    risk_priority = {
        4: {1: "Tinggi", 2: "Tinggi", 3: "Sangat Tinggi", 4: "Sangat Tinggi"},
        3: {1: "Sedang", 2: "Sedang", 3: "Tinggi",       4: "Sangat Tinggi"},
        2: {1: "Rendah", 2: "Sedang", 3: "Sedang",       4: "Tinggi"},
        1: {1: "Rendah", 2: "Rendah", 3: "Sedang",       4: "Tinggi"}
    }
    df["Tingkat Risiko"] = df.apply(lambda r: risk_priority.get(r["Nilai Residual Risk"], {}).get(r["Nilai Risiko Pengguna Jasa"]), axis=1)
    return df

# -----------------------
# OCR PDF validation
# -----------------------
def validasi_ocr_pdf(uploaded_file1, kata_kunci_list, judul=""):
    if uploaded_file1 is None:
        return False, "Tidak ada file.", 0

    try:
        pdf_bytes = uploaded_file1.read()
        uploaded_file1.seek(0)
        all_text = ""

        # 1) extract text via pdfplumber
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages[:5]:
                    extracted = page.extract_text()
                    if extracted:
                        all_text += extracted.lower() + "\n"
        except Exception:
            pass

        # 2) fallback OCR if no text found
        if not all_text.strip():
            try:
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    for page in pdf.pages[:5]:
                        page_image = page.to_image(resolution=200).original
                        text = pytesseract.image_to_string(page_image, lang="ind+eng", config="--psm 6")
                        all_text += text.lower() + "\n"
            except Exception as e:
                return False, f"Error OCR: {e}", 0

        # 3) fuzzy keyword search
        variasi_kata = [
            "formulir customer due diligence perorangan",
            "formulir customer due diligence",
            "enhanced due diligence", "formulir customer due diligence korporasi"
        ]

        jumlah_ditemukan = 0
        for kata_utama in kata_kunci_list:
            kata_lower = kata_utama.lower()
            variasi_relevan = [v for v in variasi_kata if kata_lower in v.lower()]

            def fuzzy_found(keyword):
                panjang = len(keyword)
                for i in range(0, max(0, len(all_text) - panjang + 1)):
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

        if not all_text.strip():
            return False, "Tidak ada teks terdeteksi", 0

        return True, all_text, jumlah_ditemukan

    except Exception as e:
        return False, f"Error umum saat OCR: {e}", 0

# -----------------------
# Helper: upload to drive
# -----------------------
def upload_to_drive(local_path, original_name, drive_service, folder_id=None):
    """Upload local file to Google Drive and return share link. If drive_service None -> return local_path."""
    if drive_service is None:
        return local_path
    try:
        metadata = {"name": original_name}
        if folder_id:
            metadata["parents"] = [folder_id]
        media = MediaFileUpload(local_path, mimetype="application/pdf")
        uploaded = drive_service.files().create(body=metadata, media_body=media, fields="id").execute()
        file_id = uploaded.get("id")
        # set anyone permission
        drive_service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"}
        ).execute()
        return f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"
    except Exception as e:
        st.error(f"❌ Gagal upload ke Google Drive: {e}")
        return local_path

# -----------------------
# UI: Streamlit form
# -----------------------
st.title("📊 Kuisioner PMPJ Notaris - Kementerian Hukum Jawa Timur")

with st.form("risk_form"):
    st.subheader("Identitas Notaris")
    nama_notaris = st.text_input("1. Nama Notaris (contoh: Herman Setiawan, S.H., M.Kn)", "")
    NIK_KTP = st.text_input("NIK KTP (16 digit angka)")
    username = st.text_input("Username Akun AHU Online", "")
    nomor_HP = st.text_input("Nomor HP")
    alamat = st.text_input("Alamat Lengkap Kantor Notaris", "")

    daftar_kota = [
        "Kabupaten Bangkalan","Kabupaten Banyuwangi","Kabupaten Blitar","Kabupaten Bojonegoro",
        "Kabupaten Bondowoso","Kabupaten Gresik","Kabupaten Jember","Kabupaten Jombang",
        "Kabupaten Kediri","Kabupaten Lamongan","Kabupaten Lumajang","Kabupaten Madiun",
        "Kabupaten Magetan","Kabupaten Malang","Kabupaten Mojokerto","Kabupaten Nganjuk",
        "Kabupaten Ngawi","Kabupaten Pacitan","Kabupaten Pamekasan","Kabupaten Pasuruan",
        "Kabupaten Ponorogo","Kabupaten Probolinggo","Kabupaten Sampang","Kabupaten Sidoarjo",
        "Kabupaten Situbondo","Kabupaten Sumenep","Kabupaten Trenggalek","Kabupaten Tuban",
        "Kabupaten Tulungagung","Kota Batu","Kota Blitar","Kota Kediri","Kota Madiun",
        "Kota Malang","Kota Mojokerto","Kota Pasuruan","Kota Probolinggo","Kota Surabaya"
    ]

    daftar_wilayah = list(wilayah_skor.keys())

    kota = st.selectbox("Pilih Kedudukan Kota/Kabupaten", daftar_kota)
    wilayah_input = st.selectbox("Pilih Wilayah Provinsi Kedudukan", daftar_wilayah)

    st.subheader("Jumlah Klien Sesuai Profesi")
    inputs_profil = {k: st.number_input(k, min_value=0, value=0, key=f"prof_{i}") for i,k in enumerate(profil.keys())}
    jumlah_klien = sum(inputs_profil.values())
    st.write("Jumlah Klien (total):", jumlah_klien)

    st.subheader("Jumlah Klien Sesuai Bisnis")
    inputs_bisnis = {k: st.number_input(k, min_value=0, value=0, key=f"bis_{i}") for i,k in enumerate(bisnis_pengguna.keys())}

    st.subheader("Jumlah Klien Sesuai Jasa yang Digunakan")
    inputs_jasa = {k: st.number_input(k, min_value=0, value=0, key=f"jasa_{i}") for i,k in enumerate(jasa.keys())}

    st.subheader("Jumlah Dokumen/Produk Jasa yang Diurus Klien")
    inputs_produk = {k: st.number_input(k, min_value=0, value=0, key=f"prod_{i}") for i,k in enumerate(produk.keys())}

    st.subheader("Jumlah Klien Sesuai Negara")
    inputs_negara = {k: st.number_input(k, min_value=0, value=0, key=f"neg_{i}") for i,k in enumerate(negara.keys())}

    st.subheader("Terkait Aparat Penegak Hukum")
    inputs_apgakkum = st.radio("Apakah Notaris pernah dipanggil atau diminta informasi oleh Aparat Penegak Hukum?", ["YA", "TIDAK"])

    st.subheader("Pertanyaan Kepatuhan Notaris")
    q1 = st.radio("1. Apakah Kantor Notaris anda memiliki mekanisme analisis risiko Pengguna Jasa? (form cdd, edd dan analisa resiko)?", ["YA", "TIDAK"])
    uploaded_file1 = st.file_uploader("Upload Dokumen Pendukung (Form CDD, EDD dan Analisa Resiko) dengan format PDF", type=["pdf"])
    if uploaded_file1 is not None:
        st.success(f"File berhasil diupload: {uploaded_file1.name}")

    q2 = st.radio("2. Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko tinggi ...?", ["YA", "TIDAK"])
    uploaded_file2 = st.file_uploader("Upload Dokumen Pendukung (SOP PMPJ) dengan format PDF", type=["pdf"])
    if uploaded_file2 is not None:
        st.success(f"File berhasil diupload: {uploaded_file2.name}")

    # Additional Qs (q3..q34) - create as radio inputs similar to original
    q_list = []
    for i in range(3, 35):
        q = st.radio(f"{i}. (Pertanyaan {i})", ["YA", "TIDAK"], key=f"q{i}")
        q_list.append(q)

    submitted = st.form_submit_button("Submit")

# -----------------------
# On submit: validate, compute, save
# -----------------------
if submitted:
    required_fields = [nama_notaris, NIK_KTP, username, nomor_HP, alamat, kota, q1, q2]  # q34 etc included in q_list
    missing = any(f is None or f == "" for f in required_fields)

    if missing:
        st.error("⚠️ Semua data wajib diisi (kecuali dokumen pendukung).")
    elif not (NIK_KTP.isdigit() and len(NIK_KTP) == 16):
        st.error("⚠️ NIK KTP harus 16 digit angka.")
    elif not nomor_HP.isdigit():
        st.error("⚠️ Nomor HP harus berupa angka.")
    else:
        kata_kunci_list = [
            "Formulir Customer Due Diligence",
            "formulir customer due diligence perorangan",
            "Analisis Risiko", "Analisis Resiko",
            "Enhanced Due Diligence",
            "CDD",
            "EDD"
        ]
        is_valid_ocr_q1, teks_ocr_q1, jumlah_kata_ditemukan_q1 = validasi_ocr_pdf(
            uploaded_file1, kata_kunci_list, judul="Dokumen Q1 (CDD/EDD/Analisis Risiko)"
        )

        # save uploaded files locally then upload to drive
        os.makedirs("uploads", exist_ok=True)
        doc1_path, doc2_path = "", ""
        file_link_1, file_link_2 = "", ""
        DRIVE_FOLDER_ID = None  # set jika mau upload ke folder spesifik

        if uploaded_file1 is not None:
            filename_1 = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_doc1_{uploaded_file1.name}"
            doc1_path = os.path.join("uploads", filename_1)
            with open(doc1_path, "wb") as f:
                f.write(uploaded_file1.getbuffer())
            file_link_1 = upload_to_drive(doc1_path, uploaded_file1.name, drive_service, DRIVE_FOLDER_ID)

        if uploaded_file2 is not None:
            filename_2 = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_doc2_{uploaded_file2.name}"
            doc2_path = os.path.join("uploads", filename_2)
            with open(doc2_path, "wb") as f:
                f.write(uploaded_file2.getbuffer())
            file_link_2 = upload_to_drive(doc2_path, uploaded_file2.name, drive_service, DRIVE_FOLDER_ID)

        # calculations
        hasil_inherent = hitung_risiko({
            "profil": inputs_profil,
            "bisnis": inputs_bisnis,
            "jasa": inputs_jasa,
            "negara": inputs_negara,
            "apgakkum": inputs_apgakkum,
            "wilayah": wilayah_input
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

        # prepare data dict (identitas + detail + q's)
        data = {
            "Timestamp": timestamp,
            "Nama Notaris": nama_notaris.title(),
            "NIK KTP": NIK_KTP,
            "Username Akun AHU Online": username,
            "Nomor HP": nomor_HP,
            "2. Alamat Lengkap Kantor Notaris": alamat,
            "Kedudukan Kota/Kabupaten": kota,
            "3. Jumlah Klien Tahun 2024-2025": jumlah_klien,
            "Wilayah": wilayah_input
        }

        data.update({k: inputs_profil.get(k, 0) for k in profil.keys()})
        data.update({k: inputs_bisnis.get(k, 0) for k in bisnis_pengguna.keys()})
        data.update({k: inputs_jasa.get(k, 0) for k in jasa.keys()})
        data.update({k: inputs_produk.get(k, 0) for k in produk.keys()})
        data.update({k: inputs_negara.get(k, 0) for k in negara.keys()})

        # questions: q1, q2 and q_list (q3..q34)
        data["1.  Apakah Kantor Notaris ... (CDD/EDD)?"] = q1
        data["2.  Apakah Kantor Notaris ... (SOP)?"] = q2
        for idx, qval in enumerate(q_list, start=3):
            data[f"{idx}. Pertanyaan {idx}"] = qval

        data["Dokumen_Pendukung (Q1)"] = file_link_1
        data["Dokumen Pendukung (SOP PMPJ) (Q2)"] = file_link_2

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
        data["jawaban_wilayah"]  = hasil_inherent["jawaban_wilayah"]
        data["skor_wilayah"]     = hasil_inherent["skor_wilayah"]

        data["Nilai Inherent Risk"]     = hasil_inherent["total_skor"]
        data["Tingkat Inherent Risk"]   = hasil_inherent["kategori_risiko"]
        data["Nilai Internal Control"]  = nilai_ic
        data["Tingkat Internal Control"]= kategori_ic
        data["Tingkat Residual Risk"]   = kategori_residual
        data["Nilai Residual Risk"]     = nilai_residual
        data["Nilai Risiko Pengguna Jasa"]   = nilai_pengguna
        data["Tingkat Risiko Pengguna Jasa"] = kategori_pengguna
        data["Tingkat Risiko"]               = kategori_final

        # -----------------------
        # Save to Google Sheets (merge with existing, replace by NIK jika ada)
        # -----------------------
        if client is None:
            st.error("Gagal autentikasi Google Sheets. Data tidak disimpan.")
        else:
            try:
                # open sheet
                try:
                    sh = client.open_by_key(SPREADSHEET_ID)
                except Exception:
                    sh = client.open(SPREADSHEET_NAME)
                worksheet = sh.sheet1

                # read existing
                records = worksheet.get_all_records()
                existing = pd.DataFrame(records)
                # build column order from keys of data + ensure existing columns preserved
                # We'll place identity & summary first, then the rest alphabetically (simple approach)
                identity_cols = ["Timestamp","Nama Notaris","NIK KTP","Username Akun AHU Online","Nomor HP","Wilayah","2. Alamat Lengkap Kantor Notaris","Kedudukan Kota/Kabupaten","3. Jumlah Klien Tahun 2024-2025"]
                summary_cols = ["Nilai Inherent Risk","Tingkat Inherent Risk","Nilai Internal Control","Tingkat Internal Control","Tingkat Residual Risk","Nilai Residual Risk","Nilai Risiko Pengguna Jasa","Tingkat Risiko Pengguna Jasa","Tingkat Risiko"]
                other_cols = [c for c in data.keys() if c not in identity_cols + summary_cols]
                column_order = identity_cols + other_cols + summary_cols

                # ensure existing has these columns
                if existing.empty:
                    existing = pd.DataFrame(columns=column_order)
                else:
                    for c in column_order:
                        if c not in existing.columns:
                            existing[c] = ""

                # prepare new row dataframe
                row_df = pd.DataFrame([data])
                # ensure same columns
                row_df = row_df.reindex(columns=existing.columns)

                # deduplicate by NIK KTP
                nik_baru = str(data.get("NIK KTP", "")).strip()
                if nik_baru and not existing.empty:
                    existing_nik = existing["NIK KTP"].astype(str).str.strip()
                    mask_duplikat = (existing_nik == nik_baru)
                else:
                    mask_duplikat = pd.Series([False] * len(existing))

                if mask_duplikat.any():
                    existing_filtered = existing[~mask_duplikat].copy()
                    df_all = pd.concat([existing_filtered, row_df], ignore_index=True)
                    st.warning(f"⚠️ Data lama untuk NIK {nik_baru} ditemukan dan digantikan.")
                else:
                    df_all = pd.concat([existing, row_df], ignore_index=True)
                    st.success(f"✅ Data baru ditambahkan.")

                # write back: clear and update
                worksheet.clear()
                header = df_all.columns.tolist()
                values = df_all.fillna("").astype(str).values.tolist()
                data_to_write = [header] + values
                worksheet.update("A1", data_to_write)
                st.success("✅ Data berhasil disimpan ke Google Spreadsheet.")
            except Exception as e:
                st.error(f"❌ Error saat menyimpan ke Google Sheets: {e}")
