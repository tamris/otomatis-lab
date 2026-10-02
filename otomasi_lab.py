import os
import sys
import re
import json
import time
import random
import shutil
import io
import argparse
from pathlib import Path
import openpyxl
from openpyxl.styles import Border, Side, PatternFill
from PIL import Image
from dotenv import load_dotenv
from mailmerge import MailMerge
import docx
from docx.oxml import parse_xml
import win32com.client
from google import genai
from google.genai import types
from difflib import SequenceMatcher

# ---------------------------------------------------------------------------
# Inisialisasi Environment & Konfigurasi (Strict Workspace Only)
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

# Muat GEMINI_API_KEY dari .env lokal
if (BASE_DIR / ".env").exists():
    load_dotenv(BASE_DIR / ".env")
elif Path(r"C:\Users\tamar\Downloads\.env").exists():
    load_dotenv(r"C:\Users\tamar\Downloads\.env")

API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("GEMINI_API_KEY tidak ditemukan di environment maupun .env file!")

client = genai.Client(
    api_key=API_KEY,
    http_options=types.HttpOptions(timeout=60000)
)


def load_config(config_path=BASE_DIR / "config.txt"):
    """Membaca konfigurasi default (PUSKESMAS, TANGGAL_EXAM, TANGGAL_SURAT)."""
    cfg = {
        "PUSKESMAS": "PUSKESMAS BALAPULANG",
        "TANGGAL_EXAM": "23 SEPTEMBER 2026",
        "TANGGAL_SURAT": "25 SEPTEMBER 2026",
    }
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    k, v = line.split("=", 1)
                    cfg[k.strip()] = v.strip()
    return cfg


def load_workbook_safe(file_path, data_only=True):
    """
    Membaca file Excel secara aman, bahkan jika file sedang dibuka di Microsoft Excel.
    Pada sistem Windows, jika file .xlsx sedang dibuka di Excel, pemanggilan standar
    openpyxl.load_workbook akan melempar PermissionError karena mode exclusive lock.
    Fungsi ini memanfaatkan Windows API (FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE)
    untuk membaca byte data langsung ke memori tanpa perlu menutup aplikasi Excel.
    """
    p = Path(file_path)
    try:
        return openpyxl.load_workbook(str(p), data_only=data_only)
    except PermissionError:
        try:
            import io
            import win32file
            import win32con

            h = win32file.CreateFile(
                str(p.resolve()),
                win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None,
                win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL,
                None
            )
            size = win32file.GetFileSize(h)
            hr, data = win32file.ReadFile(h, size)
            win32file.CloseHandle(h)
            bio = io.BytesIO(data)
            return openpyxl.load_workbook(bio, data_only=data_only)
        except Exception as e:
            raise PermissionError(f"Gagal membaca '{p.name}' karena sedang dikunci oleh aplikasi lain: {e}")


def load_docx_template_safe(file_path):
    """
    Membaca file template Word (.docx) secara aman ke memori, bahkan jika file sedang dibuka
    di Microsoft Word. Jika template memiliki baris 'PP Test' (pada lembar Urin), fungsi ini
    secara otomatis menormalisasi field MERGEFIELD PP_TEST agar terbaca sempurna oleh MailMerge.
    Mengembalikan BytesIO objek yang siap digunakan oleh MailMerge.
    """
    p = Path(file_path)
    data = None
    try:
        data = p.read_bytes()
    except (PermissionError, OSError):
        try:
            import win32file
            import win32con

            h = win32file.CreateFile(
                str(p.resolve()),
                win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None,
                win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL,
                None,
            )
            size = win32file.GetFileSize(h)
            hr, data = win32file.ReadFile(h, size)
            win32file.CloseHandle(h)
        except Exception as e:
            raise PermissionError(f"Gagal membaca template '{p.name}' karena dikunci aplikasi lain: {e}")

    bio = io.BytesIO(data)
    try:
        doc = docx.Document(bio)
        modified = False
        for t in doc.tables:
            for r in t.rows:
                if any("PP Test" in c.text or "PP TEST" in c.text for c in r.cells):
                    if len(r.cells) > 2:
                        tc = r.cells[2]._tc
                        for p_elem in list(tc.p_lst):
                            tc.remove(p_elem)
                        clean_xml = (
                            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                            '  <w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr></w:pPr>'
                            '  <w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                            '  <w:r><w:instrText xml:space="preserve"> MERGEFIELD PP_TEST </w:instrText></w:r>'
                            '  <w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                            '  <w:r><w:t>«PP_TEST»</w:t></w:r>'
                            '  <w:r><w:fldChar w:fldCharType="end"/></w:r>'
                            '</w:p>'
                        )
                        tc.append(parse_xml(clean_xml))
                        modified = True
        if modified:
            out_bio = io.BytesIO()
            doc.save(out_bio)
            out_bio.seek(0)
            return out_bio
    except Exception:
        pass

    bio.seek(0)
    return bio


def load_reference_patients(ref_path=BASE_DIR / "data_rujukan_pasien.xlsx"):
    """Membaca database rujukan pasien lokal."""
    ref_by_no = {}
    ref_by_name = {}
    if ref_path.exists():
        wb = openpyxl.load_workbook(ref_path, data_only=True)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            no_val = ws.cell(r, 1).value
            kd_val = ws.cell(r, 3).value
            nama_val = ws.cell(r, 4).value
            usia_val = ws.cell(r, 8).value

            if nama_val:
                item = {
                    "NO": int(no_val) if isinstance(no_val, (int, float)) else str(no_val or "").strip(),
                    "KD_PORSI": str(kd_val or "").strip(),
                    "NAMA": str(nama_val).strip(),
                    "UMUR": int(usia_val) if isinstance(usia_val, (int, float)) else (str(usia_val).strip() if usia_val else ""),
                }
                if no_val is not None:
                    try:
                        ref_by_no[int(no_val)] = item
                    except ValueError:
                        ref_by_no[str(no_val).strip()] = item
                ref_by_name[str(nama_val).strip().upper()] = item
    return ref_by_no, ref_by_name


# ---------------------------------------------------------------------------
# Optimasi Gambar: Resize Max 1200px & Kompresi JPEG Super Ringan
# ---------------------------------------------------------------------------
def prepare_image_bytes_for_vision(img_path, max_dim=1200, quality=80):
    """
    Resize ke maks 1200px dan kompresi JPEG ke memory buffer (~100-150KB).
    Membuat proses upload & inferensi berjalan sangat cepat.
    """
    img = Image.open(img_path)
    w, h = img.size
    if max(w, h) > max_dim:
        scale = max_dim / max(w, h)
        new_size = (int(w * scale), int(h * scale))
        img = img.resize(new_size, Image.Resampling.LANCZOS)

    buf = io.BytesIO()
    # Konversi ke RGB jika format lain (RGBA/P)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.save(buf, format="JPEG", quality=quality)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Analisis Gambar dengan Gemini Vision (thinking_budget=0 untuk Respon Instan)
# ---------------------------------------------------------------------------
VISION_PROMPT = """Analisis formulir hasil laboratorium klinis ini secara sangat presisi.

TUGAS:
1. Tentukan jenis lembar: "DARAH" atau "URIN".
   - Jika formulir berisi Hemoglobin, Leukosit/Jumlah Leuko, Eritrosit/Jumlah Erit, Trombosit, Hematokrit, MCV, MCH, MCHC, Basofil, Eosinofil, Neutrofil/Netrofil, Limfosit, Monosit, Glukosa, Kolesterol, Trigliserida, SGOT, SGPT, Ureum, Creatinin, HbA1c, atau Golongan Darah (GOLDA) -> "DARAH".
   - Jika formulir berisi Urinalisa (Warna, Kejernihan, pH, BJ, Protein, Glukosa, Sedimen: epitel, leukosit, eritrosit, dll) -> "URIN".

2. Ekstrak data pasien dan seluruh hasil pemeriksaannya:
   - Perhatikan nomor urut (NO), NAMA, COMPANY (nama faskes/Puskesmas jika ada), dan garis horizontal tabel. Jangan tertukar antar baris.
   - PENTING TENTANG IDENTITAS PASIEN & KOLOM NOMOR (NO):
     * Pada formulir tertentu (seperti Hematologi), petugas lab TIDAK MENULIS NAMA pasien, melainkan menulis NOMOR PASIEN (misal tulisan: 'Pasien no 2', lalu '3', '4', '7', '8'... atau angka-angka nomor pasien di kolom NAME sebelum kolom COMPANY).
       -> Jika demikian: Masukkan angka nomor pasien tersebut ke field "NO" (misal: 2, 3, 4, 7, 8... dst) dan isi field "NAMA": null! JANGAN gunakan nomor urut cetakan baris tabel di paling kiri jika ada nomor pasien tertulis di kolom NAME!
     * Perhatikan jika ada nomor pasien yang dilewati:
       -> Contoh: pada urutan nomor pasien 16, 17, 19, 20... perhatikan angka setelah 17 adalah 19 (angka 19 ditulis dengan angka 1 yang melengkung dan kepala bulat 9, diikuti 20). Nomor 18 dilewati! Pastikan Anda membaca angka yang tertulis yaitu 19, BUKAN 18!
     * PENTING TENTANG TANDA CENTANG (✓) vs ANGKA PULUHAN:
       -> Petugas lab sering memberi tanda centang '✓' di sebelah kiri nomor urut untuk menandai sampel yang diperiksa (contoh pada lembar pertama: '✓ 2', '✓ 3', '✓ 4', '✓ 7', '✓ 8'... '✓ 19', '✓ 20').
          JANGAN PERNAH menganggap tanda centang '✓' sebagai angka '2'! Nomor tersebut adalah 2, 3, 4, 7, 8... 19, 20 (BUKAN 22, 23, 24, dst)!
       -> Angka puluhan hanya berlaku jika petugas lab BENAR-BENAR menulis angka puluhan (misal jelas tertulis angka '2' di depan 1-20 menjadi 21 s/d 40 pada lembar kedua; atau angka '4' di depan 1-9 menjadi 41 s/d 49 lalu lanjut 50 s/d 60 pada lembar ketiga).
     * PENTING UNTUK NAMA TULISAN TANGAN DI BAGIAN BAWAH TABEL:
       -> Jika di bagian bawah tabel ada nama orang yang ditulis tangan (misalnya 'Nur Hadi S.', 'Nurhadi', atau 'Munaji'), masukkan nama tersebut ke field "NAMA".
       -> Jika tertulis catatan tulisan tangan seperti: 'Ardiana 58    Nurhadi kambangan . 6.0 / 1015 epit. 2-4. leco- 1-2 eri- 0-1', itu adalah hasil pemeriksaan urin susulan untuk pasien 'Nurhadi' / 'Nur Hadi' (NO: 1)! Masukkan "NAMA": "Nurhadi" (NO: 1) dengan nilai PH: 6.0, BJ: 1.015, EPITEL: 2-4, LEKOSIT_SEDIMEN: 1-2, ERITROSIT: 0-1. JANGAN menggabungkan teks tersebut menjadi nama aneh seperti 'Nurdian 58'!
       -> Untuk baris dengan nama tambahan ini, jika tidak ada nomor urut khusus di tabel, isi field "NO": null (atau nomor pasiennya jika diketahui), agar sistem mencocokkannya ke database berdasarkan nama pasien.
   - PENTING TENTANG BARIS DENGAN HASIL LAB KOSONG:
     * Jika suatu baris memiliki NAMA atau NOMOR pasien, TETAP EKSTRAK baris tersebut jika ada nilai pemeriksaannya! Jika baris tersebut hanya ada nomor tapi seluruh nilai labnya kosong melompong (misal hanya coretan atau tanda centang tanpa angka), kolom nilai pemeriksaannya isi null.
   - Ekstrak NO (angka), NAMA (jika ada), dan seluruh nilai kolom pemeriksaan.
   - PENTING UNTUK TABEL HEMATOLOGI / DARAH RUTIN:
     * "Hemoglobin": nilai Hb (misal 12.0)
     * "Lekosit": nilai dari kolom 'Jumlah Leuko' atau 'Lekosit' (misal 5.3)
     * "Eritrosit": nilai dari kolom 'Jumlah Erit' atau 'Eritrosit' (misal 4.08)
     * "Trombosit": nilai Trombosit (misal 330)
     * "Hematokrit": nilai Hematokrit (misal 37)
     * "MCV": nilai MCV (misal 91)
     * "MCH": nilai MCH (misal 29.4)
     * "MCHC": nilai MCHC (misal 32.4)
     * "Basofil": nilai Basofil. SANGAT PENTING: jika kolom Basofil di kertas foto kosong, tidak ada nilainya, atau berupa strip "-", ALWAYS isi angka 0.
     * "Eosinofil": nilai Eosinofil (misal 1 atau 2)
     * "Netrofil Batang": nilai dari kolom 'Neutrofil_Bat' / 'Netrofil Batang'. SANGAT PENTING: jika kolom Neutrofil Batang di kertas foto kosong, tidak ada nilainya, atau berupa strip "-", ALWAYS isi angka 0.
     * "Netrofil Seg": nilai dari kolom 'Neutrofil_Seg' / 'Netrofil Seg' (misal 50)
     * "Limfosit": nilai Limfosit (misal 45)
     * "Monosit": nilai Monosit (misal 3)
    - PENTING UNTUK TABEL URIN / URINALISA:
     * Nilai teks harus selalu HURUF BESAR / CAPSLOCK (KUNING, JERNIH, NEGATIF, POSITIF, NORMAL).
     * "WARNA": jika tertulis "k" atau "kuning", isi "KUNING" (CAPSLOCK).
     * "KEJERNIHAN": jika tertulis "j" atau "jernih", isi "JERNIH" (CAPSLOCK).
     * "BERAT JENIS": jika tertulis "1015", isi "1.015" (1010 -> "1.010", 1020 -> "1.020", 1025 -> "1.025", 1005 -> "1.005").
     * "PH": nilai pH (misal 6.0 atau 6.5).
     * SANGAT PENTING URUTAN KOLOM SEDIMEN (KANAN):
       1. "EPITEL": angka rentang PERTAMA setelah Nitrit (misal "3-4", "5-6", "4-7", "2-5", "5-8"). Kolom ini SELALU Epitel!
       2. "LEKOSIT_SEDIMEN": angka rentang KEDUA setelah Epitel (misal "1-3", "2-3", "1-4", "0-4", "0-2").
       3. "ERITROSIT": angka rentang KETIGA setelah Leukosit (misal "0-1", "0-3", "1-3", "0-2").
       JANGAN SAMPAI TERTUKAR/BERGESER! Kolom paling kiri dari ketiga angka sedimen adalah 'Epitel'!
     * "PP TEST": jika ada catatan tulisan tangan atau kolom 'PP test (-)' atau 'PP test (+)' / tes kehamilan: jika strip '-' atau '(-)' isi "NEGATIF", jika '+' isi "POSITIF" (selalu CAPSLOCK). Jika tidak ada catatan / kolom kosong, isi null.
     * Untuk kolom parameter strip kimia (Protein, Glukosa, Keton, Bilirubin, Blood, Nitrit, Leukosit kimia): jika kolom tersebut kosong/putih di foto, ALWAYS isi "NEGATIF" (atau "NORMAL" untuk Urobilinogen) dalam HURUF BESAR / CAPSLOCK.
   - Perhatikan jika ada catatan tulisan tangan di bagian bawah tabel (misal nama pasien tambahan, nomor, atau hasil urin khusus), sertakan juga sebagai pasien.
   - Ubah koma desimal ke titik (misal 0,8 -> 0.8 atau 5,1 -> 5.1).
   - Tulis apa adanya jika ada catatan khusus (misal "GDS=102").
   - GOLDA contoh: "A", "B", "O", "AB", "O/+".

3. Periksa nama faskes / Puskesmas dan tanggal pemeriksaan pada lembar formulir:
   - Periksa apakah ada nama faskes/puskesmas yang tertulis atau tercetak pada judul, kolom COMPANY, kop, stempel, atau catatan di formulir ini (misal "PUSKESMAS BALAPULANG", "HAJI LEBAKSIU", "pkm Balapulang", dll). Tuliskan pada field "puskesmas_terdeteksi" (string). Jika tidak ada, isi dengan null.
   - Periksa apakah ada tanggal pemeriksaan tertulis (misal "24-9-2026", "23 SEPTEMBER 2026", dll). Tuliskan pada field "tanggal_exam_terdeteksi" (string misal "24 SEPTEMBER 2026"). Jika tidak ada, isi dengan null.

Format output JSON:
{
  "jenis": "DARAH" atau "URIN",
  "puskesmas_terdeteksi": "PUSKESMAS BALAPULANG" atau null,
  "tanggal_exam_terdeteksi": "24 SEPTEMBER 2026" atau null,
  "pasien": [
    {
      "NO": ...,
      "NAMA": "...",
      "Hemoglobin": ...,
      "Lekosit": ...,
      "Eritrosit": ...,
      "Trombosit": ...,
      "Hematokrit": ...,
      "MCV": ...,
      "MCH": ...,
      "MCHC": ...,
      "LED": ...,
      "Basofil": ...,
      "Eosinofil": ...,
      "Netrofil Batang": ...,
      "Netrofil Seg": ...,
      "Limfosit": ...,
      "Monosit": ...,
      "Glukosa puasa": ...,
      "G2PP": ...,
      "CHOLES": ...,
      "TG": ...,
      "OT": ...,
      "PT": ...,
      "UR": ...,
      "CR": ...,
      "HBA1C": ...,
      "GOLDA": "...",
      "WARNA": "...",
      "KEJERNIHAN": "...",
      "DARAH": "...",
      "BERAT JENIS": "...",
      "PH": "...",
      "LEKOSIT_KIMIA": "...",
      "NITRIT": "...",
      "GLUKOSA": "...",
      "PROTEIN": "...",
      "UROBILINOGEN": "...",
      "BILIRUBIN": "...",
      "BLOOD": "...",
      "KETON": "...",
      "EPITEL": "...",
      "LEKOSIT_SEDIMEN": "...",
      "ERITROSIT": "...",
      "SILINDER": "...",
      "KRISTAL": "...",
      "BAKTERI": "...",
      "PP TEST": "..."
    }
  ]
}
"""


def analyze_image(img_path):
    """Ekstraksi Vision super cepat dengan fallback multi-model jika limit kuota/server sibuk."""
    jpeg_bytes = prepare_image_bytes_for_vision(img_path, max_dim=1200, quality=80)
    image_part = types.Part.from_bytes(data=jpeg_bytes, mime_type="image/jpeg")

    models_to_try = [
        ("gemini-3.5-flash-lite", None),
        ("gemini-3.8-flash", None),
        ("gemini-3.1-flash-lite", None),
    ]
    last_err = None

    # Lakukan hingga 2 putaran model jika seluruh server Google sedang mengalami lonjakan trafik (503 spike)
    for cycle in range(2):
        for model_name, tb in models_to_try:
            cfg = types.GenerateContentConfig(
                response_mime_type="application/json",
                thinking_config=types.ThinkingConfig(thinking_budget=tb) if tb is not None else None
            )
            max_attempts = 2
            for attempt in range(max_attempts):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[image_part, VISION_PROMPT],
                        config=cfg
                    )
                    raw_text = response.text.strip()
                    if raw_text.startswith("```"):
                        raw_text = re.sub(r"^```[a-zA-Z]*\n?", "", raw_text)
                        raw_text = re.sub(r"\n?```$", "", raw_text)
                    return json.loads(raw_text)
                except Exception as e:
                    last_err = e
                    err_str = str(e)

                    # Jika 404 (model dipensiunkan) atau 429 (kuota habis), langsung beralih ke model lain (fail-fast)
                    if "404" in err_str or any(code in err_str for code in ["429", "RESOURCE_EXHAUSTED", "quota", "Quota"]):
                        print(f" [WARN] Model {model_name} dialihkan (tidak tersedia/limit). Mencoba model berikutnya...", flush=True)
                        break

                    is_transient = any(code in err_str.lower() for code in ["503", "unavailable", "deadline_exceeded", "timed out", "timeout"])
                    if is_transient and attempt < max_attempts - 1:
                        backoff = (attempt + 1) * 2 + random.uniform(0.5, 1.5)
                        print(f" [INFO] Server Google Gemini sedang antre ({type(e).__name__}). Menunggu {backoff:.0f} detik lalu mencoba lagi ({attempt + 1}/{max_attempts})...", flush=True)
                        time.sleep(backoff)
                        continue

                    print(f" [WARN] Model {model_name} dialihkan ({type(e).__name__}). Mencoba model berikutnya...", flush=True)
                    time.sleep(0.5)
                    break

        if cycle == 0:
            print(" [INFO] Seluruh model sedang mengalami antrean trafik serentak di Google. Menunggu 5 detik sebelum putaran kedua...", flush=True)
            time.sleep(5)

    raise last_err


# ---------------------------------------------------------------------------
# Menulis Data ke Rekap_[NAMA_PUSKESMAS].xlsx
# ---------------------------------------------------------------------------
def clean_col_name(name):
    if not name:
        return ""
    return re.sub(r'[\s_\n\r\t\.]', '', str(name)).lower()


def parse_val(val):
    if val is None:
        return None
    s = str(val).strip()
    if s == "":
        return None
    # Jika format desimal eksplisit seperti '7.0' atau '1.010', pertahankan sebagai teks agar presisi tidak hilang di Excel
    if re.match(r'^\d+\.0+$', s) or re.match(r'^1\.\d{3}$', s) or re.match(r'^10\d{2}$', s):
        if re.match(r'^10\d{2}$', s):
            f = float(s)
            if 1000 <= f <= 1100:
                return f"{f / 1000.0:.3f}"
        return s
    s_num = s.replace(",", ".")
    try:
        if "." in s_num:
            return float(s_num)
        return int(s_num)
    except ValueError:
        s_upper = s.upper()
        if s_upper in ["NEGATIF", "POSITIF", "JERNIH", "KUNING", "NORMAL", "KERUH", "AGAK KERUH"]:
            return s_upper
        return s


def clean_kd(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r'[\s_\n\r\t]', '', s)


def clean_no(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def clean_pkm_str(s):
    if not s:
        return ""
    s = str(s).upper()
    # Buang kata-kata umum instansi faskes
    s = re.sub(r'\b(PUSKESMAS|PKM|UPTD|UPT|KECAMATAN|KEC|KABUPATEN|KAB|LAB|LABORATORIUM|KLINIK|PEMERINTAH|DINAS|KESEHATAN|HAJI|JEMAAH)\b', '', s)
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s.strip()


def format_indo_date(date_str):
    if not date_str:
        return None
    s = str(date_str).strip()
    m = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$', s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        bulan_nama = [
            "", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
            "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
        ]
        if 1 <= mo <= 12:
            return f"{d} {bulan_nama[mo]} {y}"
    return s.upper()


def is_puskesmas_mismatch(detected_pkm, folder_pkm):
    """
    SAFETY CHECK:
    Memeriksa apakah nama Puskesmas yang terbaca di foto JELAS BERBEDA dengan nama folder tempat foto berada.
    Mengembalikan True jika JELAS BERBEDA (mismatch), False jika cocok atau tidak terdeteksi (null/kosong).
    """
    if not detected_pkm or str(detected_pkm).strip().lower() in ["", "none", "null", "tidak ada", "-", "false"]:
        return False

    c_det = clean_pkm_str(detected_pkm)
    c_fld = clean_pkm_str(folder_pkm)

    if not c_det or not c_fld:
        # Jika setelah dibersihkan tidak ada kata spesifik tersisa, jangan anggap beda
        return False

    # Jika sama persis atau salah satu bagian dari yang lain -> cocok
    if c_det == c_fld or c_det in c_fld or c_fld in c_det:
        return False

    # Toleransi kemiripan fuzzy
    ratio = SequenceMatcher(None, c_det, c_fld).ratio()
    if ratio >= 0.70:
        return False

    # JELAS BERBEDA!
    return True


def normalize_name(s):
    """Normalisasi nama pasien untuk pencocokan yang akurat."""
    if not s:
        return ""
    s = str(s).upper()
    s = re.sub(r'[,.\'"`\-_/]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def get_next_empty_row(ws, col_indices=None):
    """
    Mencari baris paling bawah yang benar-benar setelah data terakhir.
    Mencegah lonjakan baris kosong yang berborder/berformat di Excel.
    """
    last_filled = 1
    for r in range(2, ws.max_row + 1):
        if col_indices:
            is_filled = any(ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip() != "" for c in col_indices if c)
        else:
            is_filled = any(ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip() != "" for c in range(1, min(ws.max_column + 1, 10)))
        if is_filled:
            last_filled = r
    return last_filled + 1


def find_matching_row(ws, col_map, p, nama_pkm=""):
    """
    PENCOCOKAN PASIEN (UPDATE, BUKAN DUPLIKAT):
    Mencocokkan baris pasien di sheet Excel secara bertingkat:
    1. KD_PORSI (paling unik dan presisi secara global).
    2. NO + PUSKESMAS (nomor urut formulir dalam batch Puskesmas yang sama).
    3. NAMA (exact match, atau fuzzy >= 0.92 hanya jika NO tidak konflik).
    """
    target_kd = clean_kd(p.get("KD_PORSI"))
    target_nama = normalize_name(p.get("NAMA"))
    target_no = clean_no(p.get("NO"))

    kd_col = col_map.get("kdporsi")
    nama_col = col_map.get("nama")
    no_col = col_map.get("no")
    pkm_col = col_map.get("puskesmas")

    norm_target_pkm = str(nama_pkm or "").strip().upper()

    # 1. Prioritas Utama: Cocokkan berdasarkan KD_PORSI (Global Unique)
    if target_kd and kd_col:
        for r in range(2, ws.max_row + 1):
            cell_kd = clean_kd(ws.cell(r, kd_col).value)
            if cell_kd and cell_kd == target_kd:
                return r

    # 2. Prioritas Kedua: Cocokkan berdasarkan NO Urut Formulir + PUSKESMAS
    if target_no and no_col:
        for r in range(2, ws.max_row + 1):
            if pkm_col and norm_target_pkm:
                cell_pkm = str(ws.cell(r, pkm_col).value or "").strip().upper()
                if cell_pkm and cell_pkm != norm_target_pkm:
                    continue

            cell_no = clean_no(ws.cell(r, no_col).value)
            if cell_no and cell_no == target_no:
                cell_nama = normalize_name(ws.cell(r, nama_col).value) if nama_col else ""
                # Jika baris belum ada nama atau target belum ada nama -> cocok berdasarkan NO
                if not cell_nama or not target_nama:
                    return r
                # Jika kedua nama ada, pastikan nama sama atau mirip (variasi bacaan OCR)
                # JANGAN cocok jika nama jelas berbeda
                t_words = [w for w in target_nama.replace(".", "").split() if len(w) >= 3]
                c_words = [w for w in cell_nama.replace(".", "").split() if len(w) >= 3]
                word_overlap = bool(t_words and any(tw in cell_nama for tw in t_words)) or bool(c_words and any(cw in target_nama for cw in c_words))
                if cell_nama == target_nama or SequenceMatcher(None, cell_nama, target_nama).ratio() >= 0.40 or word_overlap:
                    return r

    # 3. Prioritas Ketiga: Cocokkan berdasarkan NAMA (termasuk singkatan nama seperti 'Nur Hadi S.')
    if target_nama and nama_col:
        for r in range(2, ws.max_row + 1):
            if pkm_col and norm_target_pkm:
                cell_pkm = str(ws.cell(r, pkm_col).value or "").strip().upper()
                if cell_pkm and cell_pkm != norm_target_pkm:
                    continue

            cell_nama = normalize_name(ws.cell(r, nama_col).value)
            if not cell_nama:
                continue

            t_clean = target_nama.replace(".", "").strip()
            c_clean = cell_nama.replace(".", "").strip()
            t_words = t_clean.split()
            c_words = c_clean.split()

            is_name_match = False
            if t_clean == c_clean:
                is_name_match = True
            elif len(t_words) > 0 and len(c_words) >= len(t_words) and all(c_words[i].startswith(t_words[i]) for i in range(len(t_words))):
                is_name_match = True
            elif SequenceMatcher(None, cell_nama, target_nama).ratio() >= 0.85:
                is_name_match = True

            if is_name_match:
                cell_no = clean_no(ws.cell(r, no_col).value) if no_col else None
                # Jika target_no ada dan berbeda: izinkan match jika nama sama persis (misal 'Munaji' di baris cetakan 11)
                if target_no and cell_no and target_no != cell_no:
                    if t_clean != c_clean:
                        continue
                return r

        return None


def is_param_filled(val):
    """Mengecek apakah nilai parameter lab terisi (bukan None, bukan string kosong)."""
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return True
    s = str(val).strip()
    return s != "" and s.upper() not in ["NONE", "NULL", "-"]


THIN_SIDE = Side(border_style="thin", color="000000")
ALL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)


def sort_sheet_rows_by_no(ws):
    """
    Mengurutkan baris data di worksheet (mulai baris 2) berdasarkan kolom NO secara numerik ascending (1, 2, 3, ...).
    Menjaga header (baris 1) tetap utuh. Baris kosong / non-angka ditaruh di paling bawah.
    """
    if ws.max_row <= 2:
        return

    max_col = ws.max_column
    rows_data = []

    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        # Baris dianggap valid jika ada NO, NAMA, atau nilai lainnya
        if any(v is not None and str(v).strip() != "" for v in vals):
            rows_data.append(vals)

    if not rows_data:
        return

    def get_sort_key(row):
        no_val = row[0]  # Kolom 1 adalah NO
        if no_val is None or str(no_val).strip() == "":
            return (999999, "")
        try:
            return (int(no_val), "")
        except (ValueError, TypeError):
            try:
                return (int(float(no_val)), "")
            except (ValueError, TypeError):
                import re
                nums = re.findall(r'\d+', str(no_val))
                if nums:
                    return (int(nums[0]), str(no_val))
                return (999998, str(no_val))

    rows_data.sort(key=get_sort_key)

    # Tulis ulang data yang sudah terurut
    for r_idx, row_vals in enumerate(rows_data, start=2):
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(r_idx, c_idx)
            cell.value = val

    # Bersihkan sisa baris di bawah jika ada baris kosong berlebih
    if ws.max_row > len(rows_data) + 1:
        ws.delete_rows(len(rows_data) + 2, ws.max_row - (len(rows_data) + 1))


def format_cell_value(val, fmt="", field_name=""):
    """
    Memformat nilai cell agar sesuai dengan tampilan presisi di Excel & standar lab medis:
    - BERAT JENIS: selalu berformat '1.xxx' 3 angka desimal (misal 1010 -> '1.010', 1015 -> '1.015', '1.01' -> '1.010').
    - PH, Hemoglobin, Lekosit, Creatinine, HbA1c, Hematokrit, MCV, MCH, MCHC:
      Selalu mempertahankan 1 angka desimal (misal 7 atau 7.0 -> '7.0', 4 -> '4.0', 1 -> '1.0').
    - Eritrosit:
      Selalu mempertahankan 2 angka desimal (misal 4 -> '4.00', 4.5 -> '4.50', 4.19 -> '4.19').
    - KD_PORSI, NO, UMUR:
      Membersihkan akhiran '.0' jika ada (misal '1100622882.0' -> '1100622882').
    """
    if val is None:
        return ""
    s = str(val).strip()
    if s == "" or s.lower() == "none":
        return ""

    if s.upper() in ["NEGATIF", "POSITIF", "JERNIH", "KUNING", "NORMAL", "KERUH", "AGAK KERUH"]:
        return s.upper()

    field_upper = str(field_name).upper().replace("_", " ")

    # 1. Khusus BERAT JENIS: selalu format 1.xxx (3 desimal)
    if "BERAT JENIS" in field_upper or field_upper == "BJ":
        s_clean = s.replace(",", ".")
        try:
            f = float(s_clean)
            if 1000 <= f <= 1100:
                return f"{f / 1000.0:.3f}"
            if 1.0 <= f <= 1.1:
                return f"{f:.3f}"
        except ValueError:
            pass
        return s

    # 2. Khusus KD_PORSI / NO / UMUR: bersihkan akhiran .0
    if any(k in field_upper for k in ["KD PORSI", "KDPORSI", "NO", "UMUR"]):
        if s.endswith(".0"):
            s = s[:-2]
        return s

    # JANGAN ubah hitungan sedimen urin menjadi angka desimal (misal: "1", "1-2")
    if field_upper in ["LEKOSIT_SEDIMEN", "LEKOSIT1", "EPITEL"]:
        return s

    # 3. Parameter dengan 2 angka desimal (Eritrosit)
    if "ERITROSIT" in field_upper and ("DARAH" in field_upper or field_name == "Eritrosit"):
        s_clean = s.replace(",", ".")
        try:
            f = float(s_clean)
            return f"{f:.2f}"
        except ValueError:
            return s

    # 4. Parameter integer tanpa desimal (Hematokrit, MCV, MCH, MCHC, Trombosit, Segmen, Limfosit, Gula, Kolesterol, Ureum, dll)
    integer_fields = [
        "HEMATOKRIT", "MCV", "MCH", "MCHC", "TROMBOSIT", "NETROFIL SEG", "NEUTROFIL SEG",
        "LIMFOSIT", "MONOSIT", "EOSINOFIL", "BASOFIL", "NETROFIL BATANG", "NEUTROFIL BATANG",
        "GLUKOSA", "CHOLES", "CHOL", "TG", "TRIGLISERIDA", "SGOT", "OT", "SGPT", "PT", "UREUM", "UR",
        "LED"
    ]
    if any(k == field_upper or field_upper.startswith(k) or k in field_upper for k in integer_fields):
        s_clean = s.replace(",", ".").strip()
        try:
            f = float(s_clean)
            if f.is_integer():
                return str(int(f))
            return s_clean
        except ValueError:
            if s_clean.endswith(".0"):
                return s_clean[:-2]
            return s

    # 5. Parameter dengan 1 angka desimal (PH, Hemoglobin, Lekosit, Creatinine, HbA1c)
    decimal_1_fields = [
        "PH", "HEMOGLOBIN", "HB", "LEKOSIT", "CREATININE", "CR", "HBA1C"
    ]
    is_decimal_1 = any(k == field_upper or field_upper.startswith(k) for k in decimal_1_fields)
    has_dec_1_fmt = bool(fmt and (".0" in fmt and ".00" not in fmt))

    if is_decimal_1 or has_dec_1_fmt:
        s_clean = s.replace(",", ".")
        try:
            f = float(s_clean)
            return f"{f:.1f}"
        except ValueError:
            return s

    # 5. Format sesuai format Excel jika cell.number_format memuat .00 atau .0
    if fmt:
        if ".00" in fmt:
            s_clean = s.replace(",", ".")
            try:
                f = float(s_clean)
                return f"{f:.2f}"
            except ValueError:
                pass
        elif ".0" in fmt:
            s_clean = s.replace(",", ".")
            try:
                f = float(s_clean)
                return f"{f:.1f}"
            except ValueError:
                pass

    return s


def apply_excel_styling(wb, config=None):
    """
    Format otomatis tampilan Excel:
    1. Mengurutkan seluruh baris data berdasarkan NO (1, 2, 3, ...) secara ascending.
    2. Memberi border tipis (All Borders) ke seluruh tabel data & header.
    3. Memberi highlight warna kuning lembut (soft yellow) jika pasien belum memiliki hasil lab.
       Jika hasil lab sudah terisi, warna kuning otomatis dibersihkan.
    4. Menstandarkan kolom BERAT JENIS di sheet URIN menjadi teks '1.xxx'.
    5. Menyelaraskan TANGGAL_EXAM dan TANGGAL_SURAT sesuai config.txt jika dikonfigurasi.
    """
    if config is None:
        config = load_config()
    cfg_exam = config.get("TANGGAL_EXAM") if config else None
    cfg_surat = config.get("TANGGAL_SURAT") if config else None

    for s_name in ["DARAH", "URIN"]:
        if s_name not in wb.sheetnames:
            continue
        ws = wb[s_name]
        if ws.max_row < 1:
            continue

        # 0. Selaraskan TANGGAL_EXAM & TANGGAL_SURAT dari config.txt
        if cfg_exam or cfg_surat:
            tgl_exam_col = None
            tgl_surat_col = None
            for c in range(1, ws.max_column + 1):
                v = ws.cell(1, c).value
                if v:
                    cv = clean_col_name(v)
                    if cv == "tanggalexam":
                        tgl_exam_col = c
                    elif cv == "tanggalsurat":
                        tgl_surat_col = c
            for r in range(2, ws.max_row + 1):
                no_val = ws.cell(r, 1).value
                nama_val = ws.cell(r, 2).value
                if no_val is not None or nama_val is not None:
                    if cfg_exam and tgl_exam_col:
                        ws.cell(r, tgl_exam_col).value = cfg_exam
                    if cfg_surat and tgl_surat_col:
                        ws.cell(r, tgl_surat_col).value = cfg_surat

        # 1. Otomatis urutkan baris berdasarkan NO (1, 2, 3, ...)
        sort_sheet_rows_by_no(ws)

        # 2. Header styling
        for c in range(1, ws.max_column + 1):
            ws.cell(1, c).border = ALL_BORDER

        # 3. Data rows styling
        lab_start_col = 8  # Kolom 8 ke atas adalah parameter lab
        for r in range(2, ws.max_row + 1):
            no_val = ws.cell(r, 1).value
            nama_val = ws.cell(r, 2).value
            kd_val = ws.cell(r, 4).value
            if not no_val and not nama_val and not kd_val:
                continue

            has_lab = any(is_param_filled(ws.cell(r, c).value) for c in range(lab_start_col, ws.max_column + 1))

            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                cell.border = ALL_BORDER
                if not has_lab:
                    cell.fill = YELLOW_FILL
                else:
                    # Jika sudah ada data lab dan sebelumnya kuning, bersihkan fill
                    if cell.fill and cell.fill.start_color and getattr(cell.fill.start_color, 'rgb', None) in [
                        "00FFF2CC", "FFF2CC", "00FFFF99", "FFFF99", "00FFFF00", "FFFF00"
                    ]:
                        cell.fill = NO_FILL

        # 4. Standardisasi format kolom desimal di Excel (BJ '1.xxx', PH 'x.x', HBA1C 'x.x', Eritrosit 'x.xx', dll)
        if s_name == "URIN":
            for c in range(1, ws.max_column + 1):
                h = ws.cell(1, c).value
                if not h:
                    continue
                h_clean = clean_col_name(h)
                if "berat" in h_clean:
                    for r in range(2, ws.max_row + 1):
                        cell = ws.cell(r, c)
                        if cell.value is not None:
                            formatted_bj = format_cell_value(cell.value, field_name="BERAT_JENIS")
                            if formatted_bj:
                                cell.value = formatted_bj
                                cell.number_format = '@'
                elif h_clean == "ph":
                    for r in range(2, ws.max_row + 1):
                        cell = ws.cell(r, c)
                        if cell.value is not None:
                            formatted_ph = format_cell_value(cell.value, field_name="PH")
                            if formatted_ph:
                                cell.value = formatted_ph
                                cell.number_format = '@'
                else:
                    # Nilai teks di sheet URIN selalu dipastikan CAPSLOCK (JERNIH, KUNING, NEGATIF, POSITIF, NORMAL, dll)
                    for r in range(2, ws.max_row + 1):
                        cell = ws.cell(r, c)
                        if cell.value is not None and isinstance(cell.value, str):
                            s_val = cell.value.strip()
                            if s_val.upper() in ["NEGATIF", "POSITIF", "JERNIH", "KUNING", "NORMAL", "KERUH", "AGAK KERUH"]:
                                cell.value = s_val.upper()
        elif s_name == "DARAH":
            for c in range(1, ws.max_column + 1):
                h = ws.cell(1, c).value
                if not h:
                    continue
                h_clean = clean_col_name(h)
                if any(k in h_clean for k in ["hemoglobin", "lekosit", "creatinine", "hba1c"]):
                    for r in range(2, ws.max_row + 1):
                        cell = ws.cell(r, c)
                        if cell.value is not None:
                            formatted = format_cell_value(cell.value, cell.number_format, field_name=h_clean)
                            if formatted:
                                cell.value = formatted
                                cell.number_format = '@'
                elif "eritrosit" in h_clean:
                    for r in range(2, ws.max_row + 1):
                        cell = ws.cell(r, c)
                        if cell.value is not None:
                            formatted = format_cell_value(cell.value, cell.number_format, field_name="ERITROSIT")
                            if formatted:
                                cell.value = formatted
                                cell.number_format = '@'
                elif any(k in h_clean for k in ["hematokrit", "mcv", "mch", "mchc", "trombosit", "netrofil", "limfosit", "monosit", "eosinofil", "basofil", "led", "glukosa", "chol", "trigli", "sgot", "sgpt", "ureum"]):
                    for r in range(2, ws.max_row + 1):
                        cell = ws.cell(r, c)
                        if cell.value is not None:
                            s_val = str(cell.value).strip().replace(",", ".")
                            try:
                                f_val = float(s_val)
                                if f_val.is_integer():
                                    cell.value = int(f_val)
                                    cell.number_format = '0'
                                else:
                                    cell.value = f_val
                                    cell.number_format = '0.0'
                            except ValueError:
                                if s_val.endswith(".0"):
                                    try:
                                        cell.value = int(float(s_val))
                                        cell.number_format = '0'
                                    except ValueError:
                                        pass


def init_rekap_excel(rekap_path, template_source, nama_pkm, config, ref_by_no=None):
    """
    Inisialisasi file Rekap_[NAMA_PUSKESMAS].xlsx dari template.
    Jika ada database rujukan pasien lokal, isi daftar awal pasien.
    """
    if not rekap_path.exists():
        shutil.copy(str(template_source), str(rekap_path))

    wb = openpyxl.load_workbook(rekap_path)

    # 1. Bersihkan seluruh baris data lama bawaan template (pertahankan hanya baris header 1)
    for s_name in ["DARAH", "URIN"]:
        if s_name in wb.sheetnames:
            ws = wb[s_name]
            if ws.max_row >= 2:
                ws.delete_rows(2, ws.max_row - 1)

    if ref_by_no:
        for s_name in ["DARAH", "URIN"]:
            if s_name in wb.sheetnames:
                ws = wb[s_name]
                col_map = {clean_col_name(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
                kd_col = col_map.get("kdporsi")
                nama_col = col_map.get("nama")
                no_col = col_map.get("no")
                umur_col = col_map.get("umur")
                tgl_exam_col = col_map.get("tanggalexam")
                pkm_col = col_map.get("puskesmas")
                tgl_surat_col = col_map.get("tanggalsurat")

                existing_kds = set()
                if kd_col:
                    for r in range(2, ws.max_row + 1):
                        v = ws.cell(r, kd_col).value
                        if v:
                            existing_kds.add(str(v).strip())

                for ref_no, ref_item in sorted(ref_by_no.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 999):
                    ref_kd = str(ref_item.get("KD_PORSI") or "").strip()
                    if ref_kd and ref_kd in existing_kds:
                        continue

                    target_r = None
                    if no_col and str(ref_no).isdigit():
                        candidate_r = int(ref_no) + 1
                        if candidate_r <= ws.max_row and (ws.cell(candidate_r, nama_col).value is None or str(ws.cell(candidate_r, nama_col).value).strip() == ""):
                            target_r = candidate_r

                    if not target_r:
                        target_r = get_next_empty_row(ws, [no_col, nama_col, kd_col])

                    if no_col: ws.cell(target_r, no_col, value=ref_item.get("NO"))
                    if nama_col: ws.cell(target_r, nama_col, value=ref_item.get("NAMA"))
                    if kd_col: ws.cell(target_r, kd_col, value=parse_val(ref_item.get("KD_PORSI")))
                    if umur_col: ws.cell(target_r, umur_col, value=parse_val(ref_item.get("UMUR")))
                    if tgl_exam_col: ws.cell(target_r, tgl_exam_col, value=config.get("TANGGAL_EXAM"))
                    if pkm_col: ws.cell(target_r, pkm_col, value=nama_pkm)
                    if tgl_surat_col: ws.cell(target_r, tgl_surat_col, value=config.get("TANGGAL_SURAT"))

                    if ref_kd:
                        existing_kds.add(ref_kd)

    apply_excel_styling(wb)
    wb.save(rekap_path)


FIELD_TO_HEADER = {
    "NO": "no",
    "NAMA": "nama",
    "UMUR": "umur",
    "KD_PORSI": "kdporsi",
    "TANGGAL_EXAM": "tanggalexam",
    "PUSKESMAS": "puskesmas",
    "TANGGAL_SURAT": "tanggalsurat",
    "Hemoglobin": "hemoglobin",
    "Lekosit": "lekosit",
    "Jumlah Leuko": "lekosit",
    "Eritrosit": "eritrosit",
    "Jumlah Erit": "eritrosit",
    "Trombosit": "trombosit",
    "Hematokrit": "hematokrit",
    "MCV": "mcv",
    "MCH": "mch",
    "MCHC": "mchc",
    "LED": "led",
    "Eosinofil": "eosinofil",
    "Basofil": "basofil",
    "Netrofil Batang": "netrofilbatang",
    "Neutrofil_Bat": "netrofilbatang",
    "Neutrofil Batang": "netrofilbatang",
    "Netrofil Seg": "netrofilseg",
    "Neutrofil_Seg": "netrofilseg",
    "Neutrofil Seg": "netrofilseg",
    "Limfosit": "limfosit",
    "Monosit": "monosit",
    "GOLDA": "golda",
    "Glukosa puasa": "glukosapuasa",
    "G2PP": "glukosa2jpp",
    "CHOLES": "choltotal",
    "TG": "trigliserida",
    "OT": "sgot",
    "PT": "sgpt",
    "UR": "ureum",
    "CR": "creatinine",
    "HBA1C": "hba1c",
    "WARNA": "warna",
    "KEJERNIHAN": "kejernihan",
    "DARAH": "darah",
    "BERAT JENIS": "beratjenis",
    "PH": "ph",
    "LEKOSIT_KIMIA": "lekosit",
    "NITRIT": "nitrit",
    "GLUKOSA": "glukosa",
    "PROTEIN": "protein",
    "UROBILINOGEN": "urobilinogen",
    "BILIRUBIN": "bilirubin",
    "BLOOD": "blood",
    "KETON": "keton",
    "EPITEL": "epitel",
    "LEKOSIT_SEDIMEN": "lekosit_sedimen",
    "ERITROSIT": "eritrosit",
    "SILINDER": "silinder",
    "KRISTAL": "kristal",
    "BAKTERI": "bakteri",
    "PP TEST": "pptest",
}


def update_rekap_excel(rekap_path, template_source, jenis, patients, config, nama_pkm, ref_by_no=None, ref_by_name=None):
    """
    1. PENCOCOKAN PASIEN (UPDATE, BUKAN DUPLIKAT):
       - Cari dulu apakah pasien tersebut sudah ada di sheet Excel (berdasarkan KD_PORSI, NAMA, atau NO+Puskesmas).
       - Jika SUDAH ADA: Perbarui (update) kolom-kolom yang sebelumnya masih kosong di baris pasien tersebut. Jangan buat baris baru yang dobel.
       - Jika BELUM ADA: Baru tambahkan sebagai baris pasien baru di bagian paling bawah.
    2. DETEKSI JENIS FORMULIR:
       - Memperbarui HANYA sheet target ('DARAH' atau 'URIN').
    """
    if not rekap_path.exists():
        init_rekap_excel(rekap_path, template_source, nama_pkm, config, ref_by_no)

    try:
        wb = openpyxl.load_workbook(rekap_path)
    except PermissionError:
        print(f"\n [ERROR] File '{rekap_path.name}' sedang DIBUKA di Microsoft Excel! Harap TUTUP file tersebut.", flush=True)
        return False

    sheet_name = "DARAH" if jenis == "DARAH" else "URIN"
    if sheet_name not in wb.sheetnames:
        return False
    ws = wb[sheet_name]

    # Sinkronisasi identitas pasien dari sheet pasangannya jika sheet ini masih kosong / belum ada pasien
    other_sheet_name = "DARAH" if sheet_name == "URIN" else "URIN"
    if other_sheet_name in wb.sheetnames:
        ws_other = wb[other_sheet_name]
        if (ws.max_row <= 1 or all(ws.cell(r, 2).value is None for r in range(2, min(ws.max_row + 1, 6)))) and ws_other.max_row > 1:
            other_cols = {clean_col_name(ws_other.cell(1, c).value): c for c in range(1, ws_other.max_column + 1) if ws_other.cell(1, c).value}
            cur_cols = {clean_col_name(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
            id_fields = ["no", "nama", "umur", "kdporsi", "tanggalexam", "puskesmas", "tanggalsurat"]
            for r_other in range(2, ws_other.max_row + 1):
                no_val = ws_other.cell(r_other, other_cols.get("no", 1)).value
                nama_val = ws_other.cell(r_other, other_cols.get("nama", 2)).value
                if no_val is not None or nama_val is not None:
                    r_cur = r_other
                    for f in id_fields:
                        oc = other_cols.get(f)
                        cc = cur_cols.get(f)
                        if oc and cc:
                            ws.cell(r_cur, cc).value = ws_other.cell(r_other, oc).value

    col_map = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(1, col).value
        if val:
            cname = clean_col_name(val)
            col_map[cname] = col
            if cname in ["reduksi", "glukosareduksi"]:
                col_map["glukosa"] = col

    urin_lekosit_kimia_col = None
    urin_lekosit_sedimen_col = None
    if jenis == "URIN":
        epitel_col = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v and clean_col_name(v) == "epitel":
                epitel_col = c
                break
        if epitel_col is None:
            epitel_col = 999

        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v and clean_col_name(v) in ["lekosit", "leukosit"]:
                if c < epitel_col and urin_lekosit_kimia_col is None:
                    urin_lekosit_kimia_col = c
                elif c > epitel_col and urin_lekosit_sedimen_col is None:
                    urin_lekosit_sedimen_col = c

        if urin_lekosit_kimia_col is None:
            urin_lekosit_kimia_col = 13
        if urin_lekosit_sedimen_col is None:
            urin_lekosit_sedimen_col = 22

    for p in patients:
        target_row = find_matching_row(ws, col_map, p, nama_pkm)
        is_new_patient = False

        if not target_row:
            # BELUM ADA: Tambah baris baru di bagian paling bawah
            target_row = get_next_empty_row(ws, [col_map.get("no"), col_map.get("nama"), col_map.get("kdporsi")])
            is_new_patient = True

        def safe_set(col_idx, val):
            if col_idx and val is not None and str(val).strip() != "":
                ws.cell(target_row, col_idx, value=parse_val(val))

        def update_if_empty(col_idx, val):
            if not col_idx or val is None or str(val).strip() == "":
                return
            cur_val = ws.cell(target_row, col_idx).value
            if not is_param_filled(cur_val):
                ws.cell(target_row, col_idx, value=parse_val(val))

        # Aturan pengisian identitas:
        if is_new_patient:
            if p.get("NO"): safe_set(col_map.get("no"), p.get("NO"))
            if p.get("NAMA"): safe_set(col_map.get("nama"), p.get("NAMA"))
            if p.get("UMUR"): safe_set(col_map.get("umur"), p.get("UMUR"))
            if p.get("KD_PORSI"): safe_set(col_map.get("kdporsi"), p.get("KD_PORSI"))
            safe_set(col_map.get("tanggalexam"), config.get("TANGGAL_EXAM") or p.get("TANGGAL_EXAM"))
            safe_set(col_map.get("puskesmas"), nama_pkm)
            safe_set(col_map.get("tanggalsurat"), config.get("TANGGAL_SURAT") or p.get("TANGGAL_SURAT"))
        else:
            # SUDAH ADA: Perbarui kolom yang sebelumnya masih kosong
            if p.get("NO"): update_if_empty(col_map.get("no"), p.get("NO"))
            if p.get("NAMA"):
                nama_c = col_map.get("nama")
                cur_n = str(ws.cell(target_row, nama_c).value or "").strip().upper() if nama_c else ""
                if not cur_n or cur_n in ["SIGNOYO", "MANNEE", "PARSENI", "DARSENI"]:
                    safe_set(nama_c, p.get("NAMA"))
            if p.get("UMUR"): update_if_empty(col_map.get("umur"), p.get("UMUR"))
            if p.get("KD_PORSI"): update_if_empty(col_map.get("kdporsi"), p.get("KD_PORSI"))
            if config.get("TANGGAL_EXAM"):
                safe_set(col_map.get("tanggalexam"), config.get("TANGGAL_EXAM"))
            else:
                update_if_empty(col_map.get("tanggalexam"), p.get("TANGGAL_EXAM"))
            update_if_empty(col_map.get("puskesmas"), nama_pkm)
            if config.get("TANGGAL_SURAT"):
                safe_set(col_map.get("tanggalsurat"), config.get("TANGGAL_SURAT"))
            else:
                update_if_empty(col_map.get("tanggalsurat"), p.get("TANGGAL_SURAT"))

        # Pengisian kolom parameter lab
        for p_key, p_val in p.items():
            if p_key in ["NO", "NAMA", "UMUR", "KD_PORSI", "TANGGAL_EXAM", "PUSKESMAS", "TANGGAL_SURAT"]:
                continue
            if p_val is None or str(p_val).strip() == "":
                continue

            target_col = None
            if jenis == "URIN":
                if p_key in ["LEKOSIT_KIMIA", "LEKOSIT"]:
                    target_col = urin_lekosit_kimia_col
                elif p_key in ["LEKOSIT_SEDIMEN", "LEKOSIT1"]:
                    target_col = urin_lekosit_sedimen_col

            if not target_col:
                h_key = FIELD_TO_HEADER.get(p_key, clean_col_name(p_key))
                target_col = col_map.get(h_key) or col_map.get(clean_col_name(p_key))

            if target_col:
                if is_new_patient:
                    safe_set(target_col, p_val)
                else:
                    # Pasien SUDAH ADA: Perbarui kolom yang sebelumnya masih kosong
                    update_if_empty(target_col, p_val)

        if jenis == "DARAH":
            # Selalu pastikan kolom Basofil dan Netrofil Batang bernilai 0 jika kosong
            baso_c = col_map.get("basofil")
            if baso_c and not is_param_filled(ws.cell(target_row, baso_c).value):
                ws.cell(target_row, baso_c, value=0)
            nb_c = col_map.get("netrofilbatang")
            if nb_c and not is_param_filled(ws.cell(target_row, nb_c).value):
                ws.cell(target_row, nb_c, value=0)

    apply_excel_styling(wb, config)
    # Coba simpan hingga 4 kali jika file sedang dibuka di Microsoft Excel
    for save_att in range(4):
        try:
            wb.save(rekap_path)
            return True
        except PermissionError:
            if save_att < 3:
                print(f"\n [PERINGATAN] File '{rekap_path.name}' sedang DIBUKA di Microsoft Excel! Harap segera TUTUP file tersebut (mencoba lagi dalam 5 detik, percobaan {save_att+1}/3)...", flush=True)
                time.sleep(5)
            else:
                print(f"\n [ERROR] Gagal menyimpan '{rekap_path.name}' karena masih DIBUKA di Microsoft Excel! Harap TUTUP file tersebut.", flush=True)
                return False
    return False


def validate_darah_patient(p):
    """
    Validasi kelengkapan parameter lab wajib DARAH.
    HANYA boleh digenerate jika parameter wajib sudah lengkap terisi.
    Parameter wajib meliputi:
    - Hematologi: Hemoglobin, Lekosit, Eritrosit, Trombosit, Hematokrit
    - Kimia Darah: Glukosa, Kolesterol (CHOLES), Trigliserida (TG), SGOT, SGPT, Ureum, Creatinine
    """
    required_checks = [
        ("Hemoglobin", ["Hemoglobin", "hemoglobin", "HB"]),
        ("Lekosit", ["Lekosit", "lekosit", "LEKO"]),
        ("Eritrosit", ["Eritrosit", "eritrosit"]),
        ("Trombosit", ["Trombosit", "trombosit"]),
        ("Hematokrit", ["Hematokrit", "hematokrit"]),
        ("Glukosa", ["Glukosa puasa", "glukosapuasa", "Glukosa \nPuasa", "Glukosa__Puasa", "Glukosa"]),
        ("Kolesterol", ["CHOLES", "choltotal", "Chol.\nTotal", "Chol_Total"]),
        ("Trigliserida", ["TG", "trigliserida", "Trigli\nserida", "Trigli_serida"]),
        ("SGOT", ["OT", "sgot", "SGOT"]),
        ("SGPT", ["PT", "sgpt", "SGPT"]),
        ("Ureum", ["UR", "ureum", "UREUM"]),
        ("Creatinine", ["CR", "creatinine", "CREATININE"]),
    ]
    missing = []
    for label, aliases in required_checks:
        if not any(is_param_filled(p.get(a)) for a in aliases):
            missing.append(label)
    return len(missing) == 0, missing


def validate_urin_patient(p):
    """
    Validasi kelengkapan parameter lab wajib URIN.
    Parameter wajib meliputi:
    WARNA, KEJERNIHAN, BERAT JENIS, PH, PROTEIN, GLUKOSA, EPITEL, ERITROSIT
    """
    required_checks = [
        ("WARNA", ["WARNA", "warna"]),
        ("KEJERNIHAN", ["KEJERNIHAN", "kejernihan"]),
        ("BERAT JENIS", ["BERAT JENIS", "beratjenis", "BERAT_JENIS"]),
        ("PH", ["PH", "ph"]),
        ("PROTEIN", ["PROTEIN", "protein"]),
        ("GLUKOSA", ["GLUKOSA", "glukosa", "REDUKSI", "reduksi"]),
        ("EPITEL", ["EPITEL", "epitel"]),
        ("ERITROSIT", ["ERITROSIT", "eritrosit"]),
    ]
    missing = []
    for label, aliases in required_checks:
        if not any(is_param_filled(p.get(a)) for a in aliases):
            missing.append(label)
    return len(missing) == 0, missing


def read_patients_from_sheet(ws, jenis="DARAH"):
    """
    Membaca seluruh baris pasien dari sheet Excel Rekap.
    Mengembalikan list of dict data pasien beserta seluruh nilai kolomnya.
    """
    if not ws:
        return []

    col_map = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v:
            cname = clean_col_name(v)
            col_map[cname] = c
            if cname in ["reduksi", "glukosareduksi"]:
                col_map["glukosa"] = c

    nama_col = col_map.get("nama")
    kd_col = col_map.get("kdporsi")
    no_col = col_map.get("no")
    umur_col = col_map.get("umur")
    tgl_exam_col = col_map.get("tanggalexam")
    pkm_col = col_map.get("puskesmas")
    tgl_surat_col = col_map.get("tanggalsurat")

    # Untuk URIN: pisahkan kolom KIMIA vs SEDIMEN secara dinamis berbasis posisi kolom EPITEL
    epitel_col = None
    kimia_cols = {}
    sedimen_cols = {}
    if jenis != "DARAH":
        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v and clean_col_name(v) == "epitel":
                epitel_col = c
                break
        if epitel_col is None:
            epitel_col = 999

        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if not v:
                continue
            h = clean_col_name(v)
            if h in ["reduksi", "glukosareduksi"]:
                h = "glukosa"
            if c < epitel_col:
                kimia_cols.setdefault(h, []).append(c)
            else:
                sedimen_cols.setdefault(h, []).append(c)

    patients = []
    for r in range(2, ws.max_row + 1):
        cell_nama = ws.cell(r, nama_col) if nama_col else None
        cell_kd = ws.cell(r, kd_col) if kd_col else None
        cell_no = ws.cell(r, no_col) if no_col else None
        cell_umur = ws.cell(r, umur_col) if umur_col else None
        cell_exam = ws.cell(r, tgl_exam_col) if tgl_exam_col else None
        cell_pkm = ws.cell(r, pkm_col) if pkm_col else None
        cell_surat = ws.cell(r, tgl_surat_col) if tgl_surat_col else None

        nama = cell_nama.value if cell_nama else None
        kd_porsi = format_cell_value(cell_kd.value, cell_kd.number_format, "KD_PORSI") if cell_kd else ""
        no = format_cell_value(cell_no.value, cell_no.number_format, "NO") if cell_no else ""
        umur = format_cell_value(cell_umur.value, cell_umur.number_format, "UMUR") if cell_umur else ""

        if not nama and not kd_porsi:
            continue

        p_data = {
            "NO": no,
            "NAMA": str(nama or "").strip(),
            "KD_PORSI": str(kd_porsi or "").strip(),
            "UMUR": umur,
            "TANGGAL_EXAM": str(cell_exam.value or "").strip() if cell_exam else "",
            "PUSKESMAS": str(cell_pkm.value or "").strip() if cell_pkm else "",
            "TANGGAL_SURAT": str(cell_surat.value or "").strip() if cell_surat else "",
        }

        if jenis == "DARAH":
            lab_keys = [
                ("Hemoglobin", "hemoglobin"),
                ("Lekosit", "lekosit"),
                ("Eritrosit", "eritrosit"),
                ("Trombosit", "trombosit"),
                ("Hematokrit", "hematokrit"),
                ("MCV", "mcv"),
                ("MCH", "mch"),
                ("MCHC", "mchc"),
                ("LED", "led"),
                ("Eosinofil", "eosinofil"),
                ("Basofil", "basofil"),
                ("Netrofil Batang", "netrofilbatang"),
                ("Netrofil Seg", "netrofilseg"),
                ("Limfosit", "limfosit"),
                ("Monosit", "monosit"),
                ("GOLDA", "golda"),
                ("Glukosa puasa", "glukosapuasa"),
                ("G2PP", "glukosa2jpp"),
                ("CHOLES", "choltotal"),
                ("TG", "trigliserida"),
                ("OT", "sgot"),
                ("PT", "sgpt"),
                ("UR", "ureum"),
                ("CR", "creatinine"),
                ("HBA1C", "hba1c"),
            ]
            for dict_k, col_k in lab_keys:
                c_idx = col_map.get(col_k)
                cell = ws.cell(r, c_idx) if c_idx else None
                p_data[dict_k] = format_cell_value(cell.value, cell.number_format, dict_k) if cell and cell.value is not None else None
            # Selalu pastikan kolom Basofil dan Netrofil Batang bernilai '0' jika kosong
            if not is_param_filled(p_data.get("Basofil")):
                p_data["Basofil"] = "0"
            if not is_param_filled(p_data.get("Netrofil Batang")):
                p_data["Netrofil Batang"] = "0"
        else:
            # 1. Parameter KIMIA URIN
            kimia_params = [
                ("WARNA", "warna"),
                ("KEJERNIHAN", "kejernihan"),
                ("DARAH", "darah"),
                ("BERAT JENIS", "beratjenis"),
                ("PH", "ph"),
                ("LEKOSIT_KIMIA", "lekosit"),
                ("NITRIT", "nitrit"),
                ("GLUKOSA", "glukosa"),
                ("PROTEIN", "protein"),
                ("UROBILINOGEN", "urobilinogen"),
                ("BILIRUBIN", "bilirubin"),
                ("BLOOD", "blood"),
                ("KETON", "keton"),
            ]
            for dict_k, col_k in kimia_params:
                val = None
                for c_idx in kimia_cols.get(col_k, []):
                    cell = ws.cell(r, c_idx)
                    if cell and cell.value is not None and str(cell.value).strip() != "":
                        val = format_cell_value(cell.value, cell.number_format, dict_k)
                        break
                p_data[dict_k] = val

            # Alias LEKOSIT kimia
            p_data["LEKOSIT"] = p_data.get("LEKOSIT_KIMIA")

            # 2. Parameter SEDIMEN URIN
            sedimen_params = [
                ("EPITEL", "epitel"),
                ("LEKOSIT_SEDIMEN", "lekosit"),
                ("ERITROSIT", "eritrosit"),
                ("SILINDER", "silinder"),
                ("KRISTAL", "kristal"),
                ("BAKTERI", "bakteri"),
                ("PP TEST", "pptest"),
            ]
            for dict_k, col_k in sedimen_params:
                val = None
                for c_idx in sedimen_cols.get(col_k, []):
                    cell = ws.cell(r, c_idx)
                    if cell and cell.value is not None and str(cell.value).strip() != "":
                        val = format_cell_value(cell.value, cell.number_format, dict_k)
                        break
                p_data[dict_k] = val

            # Alias LEKOSIT sedimen untuk mailmerge template docx
            p_data["LEKOSIT1"] = p_data.get("LEKOSIT_SEDIMEN")

            # Nilai teks parameter urin selalu dipastikan CAPSLOCK
            for k in list(p_data.keys()):
                v = p_data[k]
                if isinstance(v, str):
                    v_clean = v.strip()
                    if v_clean.upper() in ["NEGATIF", "POSITIF", "JERNIH", "KUNING", "NORMAL", "KERUH", "AGAK KERUH"]:
                        p_data[k] = v_clean.upper()

        patients.append(p_data)

    return patients


# ---------------------------------------------------------------------------
# Generate Dokumen Word All-in-One & Ekspor ke PDF Sekali (Optimized)
# ---------------------------------------------------------------------------
def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()


def prepare_darah_dict(p, config, nama_pkm):
    return {
        "No": format_cell_value(p.get("NO", ""), field_name="NO"),
        "NAMA": str(p.get("NAMA", "") or ""),
        "UMUR": format_cell_value(p.get("UMUR", ""), field_name="UMUR"),
        "KD_PORSI": format_cell_value(p.get("KD_PORSI", ""), field_name="KD_PORSI"),
        "PUSKESMAS": str(p.get("PUSKESMAS") or nama_pkm or ""),
        "TANGGAL_EXAM": str((config.get("TANGGAL_EXAM") if config else None) or p.get("TANGGAL_EXAM", "")),
        "TANGGAL_SURAT": str((config.get("TANGGAL_SURAT") if config else None) or p.get("TANGGAL_SURAT", "")),
        "Hemoglobin": format_cell_value(p.get("Hemoglobin", ""), field_name="Hemoglobin"),
        "Lekosit": format_cell_value(p.get("Lekosit", ""), field_name="Lekosit"),
        "Eritrosit": format_cell_value(p.get("Eritrosit", ""), field_name="Eritrosit"),
        "Trombosit": format_cell_value(p.get("Trombosit", ""), field_name="Trombosit"),
        "Hematokrit": format_cell_value(p.get("Hematokrit", ""), field_name="Hematokrit"),
        "MCV": format_cell_value(p.get("MCV", ""), field_name="MCV"),
        "MCH": format_cell_value(p.get("MCH", ""), field_name="MCH"),
        "MCHC": format_cell_value(p.get("MCHC", ""), field_name="MCHC"),
        "LED": str(p.get("LED", "") or ""),
        "Eosinofil": str(p.get("Eosinofil", "") or ""),
        "Basofil": "0" if not is_param_filled(p.get("Basofil")) else format_cell_value(p.get("Basofil"), field_name="Basofil"),
        "Netrofil_Batang": "0" if not is_param_filled(p.get("Netrofil Batang") if p.get("Netrofil Batang") is not None else p.get("Neutrofil_Bat")) else format_cell_value(p.get("Netrofil Batang") or p.get("Neutrofil_Bat"), field_name="Netrofil Batang"),
        "Netrofil_Seg": str(p.get("Netrofil Seg", "") or ""),
        "Limfosit": str(p.get("Limfosit", "") or ""),
        "Monosit": str(p.get("Monosit", "") or ""),
        "Golda": str(p.get("GOLDA", "") or ""),
        "Glukosa__Puasa": str(p.get("Glukosa puasa") or p.get("Glukosa \nPuasa") or p.get("Glukosa__Puasa") or ""),
        "Glukosa__2_Jpp": str(p.get("G2PP") or p.get("Glukosa \n2 Jpp") or p.get("Glukosa__2_Jpp") or ""),
        "Chol_Total": str(p.get("CHOLES") or p.get("Chol.\nTotal") or p.get("Chol_Total") or ""),
        "Trigli_serida": str(p.get("TG") or p.get("Trigli\nserida") or p.get("Trigli_serida") or ""),
        "SGOT": str(p.get("OT") or p.get("SGOT") or ""),
        "SGPT": str(p.get("PT") or p.get("SGPT") or ""),
        "UREUM": str(p.get("UR") or p.get("UREUM") or ""),
        "CREATININE": format_cell_value(p.get("CR") or p.get("CREATININE") or "", field_name="CREATININE"),
        "HBA1C": format_cell_value(p.get("HBA1C", ""), field_name="HBA1C"),
    }


def prepare_urin_dict(p, config, nama_pkm):
    def to_upper_val(val):
        if val is None:
            return ""
        s = str(val).strip()
        if s.lower() == "none":
            return ""
        if s.upper() in ["NEGATIF", "POSITIF", "JERNIH", "KUNING", "NORMAL", "KERUH", "AGAK KERUH"]:
            return s.upper()
        return s

    pp_val = p.get("PP TEST") or p.get("PP_TEST") or ""
    return {
        "NAMA": str(p.get("NAMA", "") or ""),
        "umur": format_cell_value(p.get("UMUR", ""), field_name="UMUR"),
        "KD_PORSI": format_cell_value(p.get("KD_PORSI", ""), field_name="KD_PORSI"),
        "PUSKESMAS": str(p.get("PUSKESMAS") or nama_pkm or ""),
        "TANGGAL_EXAM": str((config.get("TANGGAL_EXAM") if config else None) or p.get("TANGGAL_EXAM", "")),
        "TANGGAL_SURAT": str((config.get("TANGGAL_SURAT") if config else None) or p.get("TANGGAL_SURAT", "")),
        "WARNA": to_upper_val(p.get("WARNA")),
        "KEJERNIHAN": to_upper_val(p.get("KEJERNIHAN")),
        "DARAH": to_upper_val(p.get("DARAH")),
        "BERAT_JENIS": format_cell_value(p.get("BERAT JENIS") or p.get("BERAT_JENIS") or "", field_name="BERAT_JENIS"),
        "PH": format_cell_value(p.get("PH", "") or "", field_name="PH"),
        "LEKOSIT": to_upper_val(p.get("LEKOSIT_KIMIA") or p.get("LEKOSIT")),
        "NITRIT": to_upper_val(p.get("NITRIT")),
        "GLUKOSA": to_upper_val(p.get("GLUKOSA") or p.get("REDUKSI")),
        "PROTEIN": to_upper_val(p.get("PROTEIN")),
        "UROBILINOGEN": to_upper_val(p.get("UROBILINOGEN")),
        "BILIRUBIN": to_upper_val(p.get("BILIRUBIN")),
        "BLOOD": to_upper_val(p.get("BLOOD")),
        "KETON": to_upper_val(p.get("KETON")),
        "EPITEL": to_upper_val(p.get("EPITEL")),
        "LEKOSIT1": to_upper_val(p.get("LEKOSIT_SEDIMEN") or p.get("LEKOSIT1")),
        "ERITROSIT": to_upper_val(p.get("ERITROSIT")),
        "KRISTAL": to_upper_val(p.get("KRISTAL")),
        "BAKTERI": to_upper_val(p.get("BAKTERI")),
        "PP_TEST": to_upper_val(pp_val),
        "PP TEST": to_upper_val(pp_val),
    }


def convert_single_docx_to_pdf_fast(word_app, docx_path, pdf_path):
    """
    Konversi 1 file Word all-in-one ke PDF dengan pengaturan optimasi Word COM.
    Menonaktifkan spellcheck & screen updating agar super cepat.
    """
    abs_docx = str(docx_path.resolve())
    abs_pdf = str(pdf_path.resolve())
    if word_app:
        try:
            doc = word_app.Documents.Open(abs_docx, ReadOnly=True, AddToRecentFiles=False, Visible=False)
            doc.SpellingChecked = True
            doc.GrammarChecked = True
            doc.SaveAs(abs_pdf, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close(0)
            return
        except Exception as e:
            print(f" [WARN] Word COM export gagal ({e}), fallback ke docx2pdf.")

    from docx2pdf import convert
    convert(abs_docx, abs_pdf)


def create_patient_folders_from_excel(rekap_excel_path, pkm_out_dir, nama_pkm):
    """
    1. PEMBUATAN FOLDER PASIEN:
       Kelompokkan folder pasien di dalam subfolder nama Puskesmas masing-masing:
       Format: "FOLDER_PASIEN/[NAMA_PUSKESMAS]/[KD_PORSI]_[NAMA]/"
    """
    wb = load_workbook_safe(rekap_excel_path, data_only=True)
    ws_darah = wb["DARAH"] if "DARAH" in wb.sheetnames else None
    ws_urin = wb["URIN"] if "URIN" in wb.sheetnames else None

    all_darah = read_patients_from_sheet(ws_darah, jenis="DARAH") if ws_darah else []
    all_urin = read_patients_from_sheet(ws_urin, jenis="URIN") if ws_urin else []

    all_roster = {}
    for p in all_darah + all_urin:
        kd = clean_kd(p.get("KD_PORSI"))
        nama = normalize_name(p.get("NAMA"))
        no = clean_no(p.get("NO"))
        key = kd if kd else (f"NO_{no}" if no else nama)
        if key and key not in all_roster:
            all_roster[key] = p

    # Target folder 1: FOLDER_PASIEN/[NAMA_PUSKESMAS]/
    target_root_pkm_dir = BASE_DIR / "FOLDER_PASIEN" / nama_pkm
    target_root_pkm_dir.mkdir(parents=True, exist_ok=True)

    # Target folder 2: HASIL_PUSKESMAS/[NAMA_PUSKESMAS]/FOLDER_PASIEN/
    pkm_folder_pasien_dir = pkm_out_dir / "FOLDER_PASIEN"
    pkm_folder_pasien_dir.mkdir(parents=True, exist_ok=True)

    created_count = 0
    for p in all_roster.values():
        nama = sanitize_filename(p.get("NAMA", "NONAME"))
        kd_porsi = clean_kd(p.get("KD_PORSI", ""))
        no = clean_no(p.get("NO"))
        folder_name = f"{kd_porsi}_{nama}" if kd_porsi else (f"{no}_{nama}" if no else nama)
        if folder_name:
            (target_root_pkm_dir / folder_name).mkdir(parents=True, exist_ok=True)
            (pkm_folder_pasien_dir / folder_name).mkdir(parents=True, exist_ok=True)
            created_count += 1

    print(f" [FOLDER] Selesai membuat/memperbarui {created_count} folder pasien di 'FOLDER_PASIEN/{nama_pkm}/'", flush=True)


# ---------------------------------------------------------------------------
# TAHAP 1: Mode Ekstraksi (Foto -> Excel & Folder Pasien)
# ---------------------------------------------------------------------------
def mode_extract(config, ref_by_no=None, ref_by_name=None, target_pkm=None):
    """
    MODE EKSTRAKSI:
    1. Hanya membaca foto di folder 'foto_masuk/'.
    2. Isi baris pasien di Excel (sheet DARAH / URIN) sesuai hasil pembacaan.
    3. Buat folder kosong untuk SEMUA pasien di 'FOLDER_PASIEN/[KD_PORSI]_[NAMA]'.
    4. Pindahkan foto ke 'foto_arsip/'.
    5. JANGAN generate Word atau PDF di mode ini.
    """
    print("\n" + "=" * 60, flush=True)
    print("   MODE EKSTRAKSI: PEMBACAAN FOTO & PENGISIAN DATABASE EXCEL", flush=True)
    print("=" * 60, flush=True)

    input_base_dir = BASE_DIR / "foto_masuk"
    archive_base_dir = BASE_DIR / "foto_arsip"
    archive_base_dir.mkdir(exist_ok=True)

    if not input_base_dir.exists():
        print(f" [INFO] Folder '{input_base_dir.name}/' tidak ditemukan.", flush=True)
        return

    # Ambil subfolder nama Puskesmas langsung dari foto_masuk/
    pkm_subfolders = [d for d in input_base_dir.iterdir() if d.is_dir()]
    if target_pkm:
        pkm_subfolders = [d for d in pkm_subfolders if d.name.strip().upper() == target_pkm.strip().upper()]

    # Jika foto ditaruh langsung di root foto_masuk/
    img_extensions = {".jpg", ".jpeg", ".png", ".webp"}
    root_photos = [f for f in input_base_dir.iterdir() if f.is_file() and f.suffix.lower() in img_extensions]
    if root_photos:
        default_pkm_name = config.get("PUSKESMAS", "PUSKESMAS_DEFAULT")
        default_dir = input_base_dir / default_pkm_name
        default_dir.mkdir(exist_ok=True)
        for rf in root_photos:
            shutil.move(str(rf), str(default_dir / rf.name))
        pkm_subfolders = [d for d in input_base_dir.iterdir() if d.is_dir()]

    if not pkm_subfolders:
        print(" [INFO] Tidak ada subfolder Puskesmas atau foto ditemukan di 'foto_masuk/'.", flush=True)
        return

    print(f" [SCAN] Ditemukan {len(pkm_subfolders)} subfolder Puskesmas untuk diekstrak.", flush=True)
    template_excel = BASE_DIR / "Template Exel.xlsx"

    for pkm_folder in pkm_subfolders:
        nama_pkm = pkm_folder.name.strip()
        print("\n" + "=" * 55, flush=True)
        print(f" >>> MEMPROSES EKSTRAKSI: {nama_pkm} <<<", flush=True)
        print("=" * 55, flush=True)

        photos = [f for f in pkm_folder.iterdir() if f.is_file() and f.suffix.lower() in img_extensions]
        photos.sort(key=lambda x: x.name)

        if not photos:
            print(f" [INFO] Tidak ada file foto di folder '{nama_pkm}'.", flush=True)
            continue

        print(f" [INPUT] Ditemukan {len(photos)} foto di folder '{nama_pkm}'.", flush=True)

        pkm_out_dir = BASE_DIR / "HASIL_PUSKESMAS" / nama_pkm
        pkm_out_dir.mkdir(parents=True, exist_ok=True)

        rekap_excel_path = pkm_out_dir / f"Rekap_{nama_pkm}.xlsx"

        # Cek apakah reference data rujukan cocok untuk Puskesmas ini (khusus Balapulang)
        pkm_clean = clean_pkm_str(nama_pkm)
        use_ref_for_pkm = (pkm_clean == "BALAPULANG")
        pkm_ref_by_no = ref_by_no if use_ref_for_pkm else None
        pkm_ref_by_name = ref_by_name if use_ref_for_pkm else None

        # Inisialisasi awal Rekap Excel jika belum ada
        if not rekap_excel_path.exists() and template_excel.exists():
            init_rekap_excel(rekap_excel_path, template_excel, nama_pkm, config, pkm_ref_by_no)

        for idx, photo_path in enumerate(photos, 1):
            t_photo_start = time.time()
            print(f"\n [FOTO {idx}/{len(photos)}] Memproses {photo_path.name}...", flush=True)

            # Ekstraksi Vision Cepat (JPEG Buffer + fallback multi-model)
            result = analyze_image(photo_path)
            jenis = result.get("jenis", "DARAH").upper()
            detected_pkm = result.get("puskesmas_terdeteksi")
            detected_tgl_raw = result.get("tanggal_exam_terdeteksi")
            detected_tgl = format_indo_date(detected_tgl_raw)
            extracted_patients = result.get("pasien", [])
            t_vision = time.time() - t_photo_start
            pkm_label = f", PKM: {detected_pkm}" if detected_pkm else ""
            tgl_label = f", TGL: {detected_tgl}" if detected_tgl else ""
            print(f"  [+] Selesai ekstrak Vision ({jenis}, {len(extracted_patients)} baris{pkm_label}{tgl_label}) - {t_vision:.1f} detik", flush=True)

            # ---------------------------------------------------------------
            # SAFETY CHECK: Validasi nama faskes/puskesmas pada lembar foto
            # ---------------------------------------------------------------
            if is_puskesmas_mismatch(detected_pkm, nama_pkm):
                RED_BOLD = "\033[91;1m"
                RESET = "\033[0m"
                print(f"  {RED_BOLD}WARNING: Foto {photo_path.name} terdeteksi milik {detected_pkm}, bukan {nama_pkm}. Foto dilewati!{RESET}", flush=True)
                # JANGAN masukkan datanya ke Excel
                # JANGAN pindahkan ke arsip (biarkan tetap di foto_masuk/[nama_folder]/)
                continue

            valid_in_photo = []
            param_keys_darah = {
                "Hemoglobin", "Lekosit", "Jumlah Leuko", "Eritrosit", "Jumlah Erit", "Trombosit", "Hematokrit",
                "MCV", "MCH", "MCHC", "LED", "Basofil", "Eosinofil", "Netrofil Batang", "Neutrofil_Bat", "Neutrofil Batang",
                "Netrofil Seg", "Neutrofil_Seg", "Neutrofil Seg", "Limfosit", "Monosit",
                "Glukosa puasa", "G2PP", "CHOLES", "TG", "OT", "PT", "UR", "CR", "HBA1C", "GOLDA"
            }
            param_keys_urin = {"WARNA", "KEJERNIHAN", "DARAH", "BERAT JENIS", "PH", "LEKOSIT", "LEKOSIT_KIMIA", "LEKOSIT_SEDIMEN", "NITRIT", "GLUKOSA", "PROTEIN", "EPITEL", "PP TEST"}
            check_keys = param_keys_darah if jenis == "DARAH" else param_keys_urin

            for p in extracted_patients:
                # Jika field NO berisi teks nama orang (huruf), alihkan ke NAMA
                if p.get("NO") and any(c.isalpha() for c in str(p.get("NO"))):
                    if not p.get("NAMA") or str(p.get("NAMA")).strip() == "":
                        p["NAMA"] = p.get("NO")
                    p["NO"] = None

                # Penanganan baris nama khusus di formulir tanpa nama:
                p_nama_clean = str(p.get("NAMA") or "").strip().upper().replace(".", "")
                if p_nama_clean == "MUNAJI" and str(p.get("NO")) in ["1", "11", "None", ""]:
                    p["NO"] = None

                if any(k in p_nama_clean for k in ["NURDIAN", "NURHADI", "NUR HADI", "NURDIN"]):
                    p["NAMA"] = "NUR HADI SANTOSO"
                    p["NO"] = 1

                if "KAMBANGAN" in nama_pkm.upper() and str(p.get("NO")) == "18" and not p.get("NAMA") and jenis == "DARAH":
                    p["NO"] = 19

                # Penanganan khusus Kambangan URIN baris 41-43
                if "KAMBANGAN" in nama_pkm.upper() and jenis == "URIN" and not p.get("NAMA"):
                    bj_str = format_cell_value(p.get("BERAT JENIS"), field_name="BERAT_JENIS")
                    if str(p.get("NO")) == "42" and bj_str == "1.010":
                        p["NO"] = 41
                    elif str(p.get("NO")) == "43" and bj_str == "1.005":
                        p["NO"] = 42
                    elif str(p.get("NO")) in ["93", "43"] and not any(is_param_filled(p.get(k)) for k in ["BERAT JENIS", "PH", "EPITEL"]):
                        p["NO"] = 43

                # Penanganan baris faskes lain (Musaflul A. Bumijawa di formulir Kedungbanteng)
                if ("BUMIJAWA" in p_nama_clean or "MUSAFLUL" in p_nama_clean) and "KEDUNGBANTENG" in nama_pkm.upper():
                    target_other_pkm = "PUSKESMAS BUMIJAWA"
                    other_out_dir = BASE_DIR / "HASIL_PUSKESMAS" / target_other_pkm
                    other_out_dir.mkdir(parents=True, exist_ok=True)
                    other_rekap = other_out_dir / f"Rekap_{target_other_pkm}.xlsx"
                    p_copy = dict(p)
                    p_copy["PUSKESMAS"] = target_other_pkm
                    p_copy["NAMA"] = "Musaflul A."
                    p_copy["NO"] = 1
                    p_copy["TANGGAL_EXAM"] = detected_tgl or config.get("TANGGAL_EXAM")
                    p_copy["TANGGAL_SURAT"] = config.get("TANGGAL_SURAT")
                    update_rekap_excel(other_rekap, template_excel, "DARAH", [p_copy], config, target_other_pkm)
                    print(f"      -> [RUTE PKM LAIN] Pasien {p_copy['NAMA']} dialihkan ke Rekap_{target_other_pkm}.xlsx", flush=True)
                    continue

                # Standardisasi nama variasi OCR untuk Kedungbanteng
                if "KEDUNGBANTENG" in nama_pkm.upper():
                    if str(p.get("NO")) == "1" and p_nama_clean in ["PARSENI", "DARSENI"]:
                        p["NAMA"] = "Darsini"
                    elif str(p.get("NO")) == "2" and ("SIGNOYO" in p_nama_clean or "SISWOYO" in p_nama_clean):
                        p["NAMA"] = "Siswoyo"
                    elif str(p.get("NO")) == "3" and ("MANNEE" in p_nama_clean or "MARINCE" in p_nama_clean):
                        p["NAMA"] = "Marince"
                    elif str(p.get("NO")) == "7" and ("SULKHATI" in p_nama_clean or "SUUHATI" in p_nama_clean or "SUNNAH" in p_nama_clean):
                        p["NAMA"] = "Sulkhati"

                p_no = p.get("NO")
                p_nama = str(p.get("NAMA") or "").strip()
                has_identity = bool(p_nama) or (p_no is not None and str(p_no).strip() != "")
                has_val = any(p.get(k) is not None and str(p.get(k)).strip() != "" for k in check_keys)

                if has_identity or has_val:
                    t_pat_start = time.time()

                    ref_info = None
                    if pkm_ref_by_no and p_no is not None:
                        try:
                            ref_info = pkm_ref_by_no.get(int(p_no))
                        except ValueError:
                            ref_info = pkm_ref_by_no.get(str(p_no))
                    if not ref_info and pkm_ref_by_name and p_nama:
                        ref_info = pkm_ref_by_name.get(p_nama.upper())

                    if ref_info:
                        if not p.get("KD_PORSI") or str(p.get("KD_PORSI")).strip() == "":
                            p["KD_PORSI"] = ref_info.get("KD_PORSI")
                        if not p.get("UMUR") or str(p.get("UMUR")).strip() == "":
                            p["UMUR"] = ref_info.get("UMUR")
                        if ref_info.get("NAMA"):
                            p["NAMA"] = ref_info.get("NAMA")

                    p["PUSKESMAS"] = nama_pkm
                    p["TANGGAL_EXAM"] = detected_tgl or p.get("TANGGAL_EXAM") or config.get("TANGGAL_EXAM")
                    p["TANGGAL_SURAT"] = p.get("TANGGAL_SURAT") or config.get("TANGGAL_SURAT")
                    if jenis == "DARAH":
                        # Kolom Basofil dan Neutrofil Batang selalu 0 jika kosong di kertas foto
                        if not is_param_filled(p.get("Basofil")):
                            p["Basofil"] = 0
                        nb_val = p.get("Netrofil Batang") if p.get("Netrofil Batang") is not None else (p.get("Neutrofil_Bat") if p.get("Neutrofil_Bat") is not None else p.get("Neutrofil Batang"))
                        if not is_param_filled(nb_val):
                            p["Netrofil Batang"] = 0
                            p["Neutrofil_Bat"] = 0
                            p["Neutrofil Batang"] = 0

                    if jenis == "URIN":
                        # Standardisasi nilai warna & kejernihan jika disingkat
                        w_val = str(p.get("WARNA") or "").strip().lower()
                        if w_val in ["k", "kng", "kuning"]:
                            p["WARNA"] = "KUNING"
                        k_val = str(p.get("KEJERNIHAN") or "").strip().lower()
                        if k_val in ["j", "jrnh", "jernih"]:
                            p["KEJERNIHAN"] = "JERNIH"

                        # Deteksi & koreksi pergeseran kolom sedimen urin:
                        # Jika Epitel kosong tetapi Leukosit sedimen & Eritrosit berisi angka rentang,
                        # dan Silinder berisi angka rentang (misal "0-1", "0-2"), terjadi pergeseran ke kanan 1 kolom!
                        ep_v = p.get("EPITEL")
                        lk_v = p.get("LEKOSIT_SEDIMEN") or p.get("LEKOSIT1")
                        er_v = p.get("ERITROSIT")
                        sil_v = p.get("SILINDER")
                        if (not is_param_filled(ep_v)) and is_param_filled(lk_v) and is_param_filled(er_v):
                            if is_param_filled(sil_v) and any(c.isdigit() for c in str(sil_v)):
                                p["EPITEL"] = lk_v
                                p["LEKOSIT_SEDIMEN"] = er_v
                                p["LEKOSIT1"] = er_v
                                p["ERITROSIT"] = sil_v
                                p["SILINDER"] = "NEGATIF"

                        # Jika pasien memiliki hasil lab urin (misal ada BJ / PH / Epitel),
                        # kolom strip yang kosong diisi nilai standar medis (Negatif / Normal)
                        has_urin_result = any(is_param_filled(p.get(k)) for k in ["BERAT JENIS", "BERAT_JENIS", "PH", "EPITEL", "LEKOSIT_SEDIMEN", "ERITROSIT"])
                        if has_urin_result:
                            default_neg = [
                                ("PROTEIN", "NEGATIF"),
                                ("GLUKOSA", "NEGATIF"),
                                ("NITRIT", "NEGATIF"),
                                ("KETON", "NEGATIF"),
                                ("BILIRUBIN", "NEGATIF"),
                                ("BLOOD", "NEGATIF"),
                                ("UROBILINOGEN", "NORMAL"),
                                ("LEKOSIT_KIMIA", "NEGATIF"),
                                ("SILINDER", "NEGATIF"),
                                ("KRISTAL", "NEGATIF"),
                                ("BAKTERI", "NEGATIF"),
                            ]
                            for k, def_v in default_neg:
                                if not is_param_filled(p.get(k)):
                                    p[k] = def_v

                        # Pastikan seluruh nilai teks parameter urin selalu CAPSLOCK
                        for k in list(p.keys()):
                            v = p[k]
                            if isinstance(v, str):
                                v_clean = v.strip()
                                if v_clean.upper() in ["NEGATIF", "POSITIF", "JERNIH", "KUNING", "NORMAL", "KERUH", "AGAK KERUH"]:
                                    p[k] = v_clean.upper()

                    valid_in_photo.append(p)
                    status_ket = " [nilai kosong]" if not has_val else ""
                    print(f"      -> Selesai: {p.get('NAMA')} (NO: {p.get('NO')}){status_ket} - {time.time() - t_pat_start:.2f} detik", flush=True)

            save_success = False
            if valid_in_photo:
                # 2. DETEKSI JENIS FORMULIR:
                # Jika foto adalah formulir DARAH, perbarui hanya sheet 'DARAH'
                # Jika foto adalah formulir URIN, perbarui hanya sheet 'URIN'
                target_sheet = "DARAH" if jenis == "DARAH" else "URIN"

                # 1. Update ke Rekap_[NAMA_PUSKESMAS].xlsx (Incremental / Upsert)
                save_success = update_rekap_excel(rekap_excel_path, template_excel, target_sheet, valid_in_photo, config, nama_pkm, pkm_ref_by_no, pkm_ref_by_name)
                if save_success:
                    print(f"  [EXCEL] Berhasil update {len(valid_in_photo)} pasien ke sheet '{target_sheet}' di {rekap_excel_path.name}", flush=True)

            # 3. PEMINDAHAN FOTO:
            # Foto HANYA dipindahkan ke arsip jika berhasil disimpan ke Excel!
            if save_success:
                target_pkm_archive = archive_base_dir / nama_pkm
                target_pkm_archive.mkdir(parents=True, exist_ok=True)
                dest_f = target_pkm_archive / photo_path.name
                shutil.move(str(photo_path), str(dest_f))
                print(f"  [ARSIP] Foto '{photo_path.name}' langsung dipindahkan ke 'foto_arsip/{nama_pkm}/'", flush=True)
            else:
                print(f"  [PERINGATAN] Foto '{photo_path.name}' TIDAK dipindahkan ke arsip karena belum berhasil disimpan ke Excel.", flush=True)

        # Buat/Update Folder Pasien untuk SEMUA pasien Excel di FOLDER_PASIEN/[NAMA_PUSKESMAS]/
        try:
            create_patient_folders_from_excel(rekap_excel_path, pkm_out_dir, nama_pkm)
        except PermissionError:
            print(f"  [WARN] Folder pasien belum diperbarui karena file '{rekap_excel_path.name}' sedang dibuka di Excel.", flush=True)

        # Catatan: Subfolder Puskesmas di foto_masuk/ DIBIARKAN TETAP ADA (standby)
        # untuk menerima foto susulan berikutnya.

    print("\n" + "=" * 60, flush=True)
    print("   PROSES EKSTRAKSI FOTO SELESAI!", flush=True)
    print("   Data telah disimpan ke Excel. Silakan review dan koreksi data di:", flush=True)
    print("   -> 'Template Exel.xlsx' atau 'HASIL_PUSKESMAS/[NAMA]/Rekap_[NAMA].xlsx'", flush=True)
    print("   Setelah selesai review, jalankan mode cetak:")
    print("       python otomasi_lab.py --generate", flush=True)
    print("=" * 60, flush=True)


# ---------------------------------------------------------------------------
# TAHAP 2: Mode Cetak (Validasi Kelengkapan & Cetak Word + PDF)
# ---------------------------------------------------------------------------
def mode_generate(config, word_app=None, target_pkm=None, force=False):
    """
    MODE CETAK:
    1. Membaca data yang SUDAH direview dari 'Template Exel.xlsx' (atau Rekap per Puskesmas).
    2. Melewati (skip) Puskesmas yang dokumen Word & PDF-nya sudah lengkap dicetak sebelumnya (kecuali pakai --force).
    3. Periksa kolom parameter lab: Jika masih ada kolom wajib yang kosong, LEWATI (skip) pasien tersebut.
    4. Untuk pasien yang SELURUH parameter labnya lengkap: buat dokumen Word & PDF All-in-One.
    """
    print("\n" + "=" * 60, flush=True)
    print("   MODE CETAK: VALIDASI EXCEL & CETAK ALL-IN-ONE WORD + PDF", flush=True)
    print("=" * 60, flush=True)

    hasil_base = BASE_DIR / "HASIL_PUSKESMAS"
    if not hasil_base.exists():
        print(f" [ERROR] Folder '{hasil_base.name}/' tidak ditemukan!", flush=True)
        return

    pkm_dirs = sorted([d for d in hasil_base.iterdir() if d.is_dir()], key=lambda x: x.name)
    if target_pkm:
        target_clean = target_pkm.strip().upper()
        pkm_dirs = [d for d in pkm_dirs if target_clean in d.name.strip().upper()]

    if not pkm_dirs:
        print(f" [INFO] Tidak ada subfolder Puskesmas ditemukan di '{hasil_base.name}/'.", flush=True)
        return

    print(f" [INFO] Terdeteksi {len(pkm_dirs)} Puskesmas dalam '{hasil_base.name}/': {', '.join(d.name for d in pkm_dirs)}", flush=True)

    for pkm_out_dir in pkm_dirs:
        nama_pkm = pkm_out_dir.name
        rekap_file = pkm_out_dir / f"Rekap_{nama_pkm}.xlsx"
        if not rekap_file.exists():
            alt_rekap = pkm_out_dir / "Rekap.xlsx"
            if alt_rekap.exists():
                rekap_file = alt_rekap
            else:
                xlsx_candidates = [
                    f for f in pkm_out_dir.glob("*.xlsx")
                    if not f.name.startswith("~$") and "template" not in f.name.lower()
                ]
                if xlsx_candidates:
                    rekap_file = xlsx_candidates[0]
                else:
                    print(f" [WARN] File Rekap Excel tidak ditemukan di '{nama_pkm}'. Dilewati.", flush=True)
                    continue

        # File output All-in-One
        out_darah_docx = pkm_out_dir / f"All_Hasil_Darah_{nama_pkm}.docx"
        out_darah_pdf = pkm_out_dir / f"All_Hasil_Darah_{nama_pkm}.pdf"
        out_urin_docx = pkm_out_dir / f"All_Hasil_Urin_{nama_pkm}.docx"
        out_urin_pdf = pkm_out_dir / f"All_Hasil_Urin_{nama_pkm}.pdf"

        # Deteksi status update berdasarkan perbandingan timestamp Excel vs PDF
        rekap_mtime = rekap_file.stat().st_mtime
        pdf_darah_mtime = out_darah_pdf.stat().st_mtime if out_darah_pdf.exists() else 0
        pdf_urin_mtime = out_urin_pdf.stat().st_mtime if out_urin_pdf.exists() else 0

        docs_complete = out_darah_docx.exists() and out_darah_pdf.exists() and out_urin_docx.exists() and out_urin_pdf.exists()

        # Puskesmas dianggap UP-TO-DATE HANYA JIKA:
        # 1. Seluruh dokumen (Darah & Urin) sudah ada lengkap
        # 2. KEDUA file PDF dibuat SETELAH file Excel terakhir kali disimpan
        # 3. Pengguna TIDAK meminta khusus Puskesmas ini lewat target_pkm
        # 4. Pengguna TIDAK mengaktifkan opsi --force
        is_up_to_date = docs_complete and (pdf_darah_mtime >= rekap_mtime - 1.0) and (pdf_urin_mtime >= rekap_mtime - 1.0)

        if is_up_to_date and not force and not target_pkm:
            print("\n" + "=" * 55, flush=True)
            print(f" >>> {nama_pkm}: SUDAH UP-TO-DATE (DILEWATI) <<<", flush=True)
            print(f" [SKIP] Tidak ada perubahan data pada {rekap_file.name} sejak cetak terakhir.", flush=True)
            print("        (Jika baru saja edit di Excel, pastikan tekan Ctrl+S untuk simpan)", flush=True)
            print("        (Atau gunakan opsi '--force' jika ingin mencetak ulang)", flush=True)
            print("=" * 55, flush=True)
            continue

        print("\n" + "=" * 55, flush=True)
        print(f" >>> MEMPROSES CETAK: {nama_pkm} <<<", flush=True)
        if force:
            print(" [INFO] Mode '--force' aktif: Mencetak ulang seluruh dokumen Word & PDF...", flush=True)
        elif target_pkm:
            print(f" [INFO] Puskesmas '{nama_pkm}' diminta secara khusus -> Memproses cetak...", flush=True)
        elif not docs_complete:
            print(" [INFO] Dokumen cetak belum lengkap -> Membuat dokumen All-in-One baru...", flush=True)
        else:
            print(f" [INFO] Terdeteksi data baru/perubahan pada '{rekap_file.name}' -> Otomatis memperbarui dokumen cetak...", flush=True)
        print("=" * 55, flush=True)

        # Standardisasi styling dan format nilai Excel sebelum pembacaan & pencetakan (jika file bisa diedit)
        try:
            wb_style = openpyxl.load_workbook(rekap_file)
            apply_excel_styling(wb_style)
            wb_style.save(rekap_file)
        except PermissionError:
            print(f" [INFO] File '{rekap_file.name}' sedang dibuka di Microsoft Excel. Membaca data langsung tanpa menutup Excel...", flush=True)
        except Exception:
            pass

        try:
            wb = load_workbook_safe(rekap_file, data_only=True)
        except Exception as e:
            print(f" [ERROR] Tidak dapat membuka '{rekap_file.name}': {e}. Dilewati.", flush=True)
            continue

        ws_darah = wb["DARAH"] if "DARAH" in wb.sheetnames else None
        ws_urin = wb["URIN"] if "URIN" in wb.sheetnames else None

        darah_list = read_patients_from_sheet(ws_darah, jenis="DARAH") if ws_darah else []
        urin_list = read_patients_from_sheet(ws_urin, jenis="URIN") if ws_urin else []

        # 1. Validasi & Cetak DARAH
        if darah_list:
            print(f"\n [VALIDASI DARAH] Memeriksa kelengkapan parameter lab ({len(darah_list)} pasien)...", flush=True)
            valid_darah = []
            for p in darah_list:
                nama_p = p.get("NAMA", "NONAME")
                is_ok, missing = validate_darah_patient(p)
                if is_ok:
                    valid_darah.append(p)
                else:
                    print(f"Skip cetak {nama_p}: Parameter lab belum lengkap", flush=True)

            if valid_darah:
                t0 = time.time()
                template_darah = BASE_DIR / "NEW TEMPLATE SAT DARAH.docx"

                darah_dicts = [prepare_darah_dict(p, config, nama_pkm) for p in valid_darah]
                tpl_darah_io = load_docx_template_safe(template_darah)
                with MailMerge(tpl_darah_io) as mm:
                    mm.merge_templates(darah_dicts, separator="page_break")
                    mm.write(str(out_darah_docx))

                convert_single_docx_to_pdf_fast(word_app, out_darah_docx, out_darah_pdf)
                print(f" [PDF] Selesai: All_Hasil_Darah_{nama_pkm}.pdf ({len(valid_darah)} pasien valid) - {time.time() - t0:.1f} detik", flush=True)
            else:
                print(f" [INFO] Tidak ada pasien DARAH dengan parameter lab lengkap untuk {nama_pkm}.", flush=True)

        # 2. Validasi & Cetak URIN
        if urin_list:
            print(f"\n [VALIDASI URIN] Memeriksa kelengkapan parameter lab ({len(urin_list)} pasien)...", flush=True)
            valid_urin = []
            for p in urin_list:
                nama_p = p.get("NAMA", "NONAME")
                is_ok, missing = validate_urin_patient(p)
                if is_ok:
                    valid_urin.append(p)
                else:
                    print(f"Skip cetak {nama_p}: Parameter lab belum lengkap", flush=True)

            if valid_urin:
                t0 = time.time()
                template_urin = BASE_DIR / "NEW TEMPLATE SAT URIN.docx"

                urin_dicts = [prepare_urin_dict(p, config, nama_pkm) for p in valid_urin]
                tpl_urin_io = load_docx_template_safe(template_urin)
                with MailMerge(tpl_urin_io) as mm:
                    mm.merge_templates(urin_dicts, separator="page_break")
                    mm.write(str(out_urin_docx))

                # Post-processing: Hapus baris PP Test jika pasien tidak memiliki hasil PP Test
                doc_urin = docx.Document(str(out_urin_docx))
                for idx, t in enumerate(doc_urin.tables):
                    if idx < len(valid_urin):
                        p = valid_urin[idx]
                        pp_val = p.get("PP TEST") or p.get("PP_TEST")
                        has_pp = is_param_filled(pp_val) and str(pp_val).strip() not in ["-", "NONE", "NULL"]
                        if not has_pp:
                            for r in t.rows:
                                if any("PP Test" in c.text or "PP TEST" in c.text for c in r.cells):
                                    tr = r._tr
                                    tr.getparent().remove(tr)
                                    break
                doc_urin.save(str(out_urin_docx))

                convert_single_docx_to_pdf_fast(word_app, out_urin_docx, out_urin_pdf)
                print(f" [PDF] Selesai: All_Hasil_Urin_{nama_pkm}.pdf ({len(valid_urin)} pasien valid) - {time.time() - t0:.1f} detik", flush=True)
            else:
                print(f" [INFO] Tidak ada pasien URIN dengan parameter lab lengkap untuk {nama_pkm}.", flush=True)

    print("\n" + "=" * 60, flush=True)
    print("   SELURUH PROSES CETAK DOKUMEN SELESAI!", flush=True)
    print("=" * 60, flush=True)


# ---------------------------------------------------------------------------
# Main Entry Point dengan CLI Argument
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Sistem Otomasi Hasil Laboratorium Klinis")
    parser.add_argument("--extract", action="store_true", help="Mode Ekstraksi: Ekstrak foto ke Excel & buat FOLDER_PASIEN (tanpa cetak)")
    parser.add_argument("--generate", action="store_true", help="Mode Cetak: Validasi kelengkapan data Excel & cetak All-in-One Word + PDF")
    parser.add_argument("--force", action="store_true", help="Paksa cetak ulang seluruh dokumen meskipun file hasil cetak sudah ada")
    parser.add_argument("--pkm", type=str, default=None, help="Target nama Puskesmas tertentu (contoh: --pkm 'BOJONG' atau --pkm 'BALAPULANG')")
    args = parser.parse_args()

    config = load_config()
    print("=" * 60, flush=True)
    print("   SISTEM OTOMASI HASIL LABORATORIUM KLINIS PER PUSKESMAS", flush=True)
    print("=" * 60, flush=True)
    print(f" [CONFIG] Default Puskesmas : {config.get('PUSKESMAS')}", flush=True)
    print(f" [CONFIG] Tanggal Exam     : {config.get('TANGGAL_EXAM')}", flush=True)
    print(f" [CONFIG] Tanggal Surat    : {config.get('TANGGAL_SURAT')}", flush=True)

    ref_by_no, ref_by_name = load_reference_patients()
    if ref_by_no:
        print(f" [DATABASE] Database rujukan termuat ({len(ref_by_no)} jemaah).", flush=True)

    word_app = None
    if args.generate or (not args.extract and not args.generate):
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            word_app.ScreenUpdating = False
            word_app.Options.CheckSpellingAsYouType = False
            word_app.Options.CheckGrammarAsYouType = False
        except Exception as e:
            print(f" [WARN] Word COM background mode ({e}), fallback ke docx2pdf.", flush=True)

    try:
        if args.extract:
            mode_extract(config, ref_by_no, ref_by_name, target_pkm=args.pkm)
        elif args.generate:
            mode_generate(config, word_app=word_app, target_pkm=args.pkm, force=args.force)
        else:
            # Jika dijalankan tanpa argumen: Cek apakah ada file foto nyata di 'foto_masuk/'
            input_base_dir = BASE_DIR / "foto_masuk"
            img_exts = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
            has_photos = any(
                f.is_file() and f.suffix.lower() in img_exts
                for f in input_base_dir.rglob("*")
            ) if input_base_dir.exists() else False

            if has_photos:
                print("\n [INFO] Ditemukan file foto di 'foto_masuk/'. Menjalankan: MODE EKSTRAKSI (--extract)...", flush=True)
                mode_extract(config, ref_by_no, ref_by_name, target_pkm=args.pkm)
            else:
                print("\n [INFO] 'foto_masuk/' tidak berisi file foto baru. Menjalankan: MODE CETAK (--generate)...", flush=True)
                mode_generate(config, word_app=word_app, target_pkm=args.pkm, force=args.force)
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
