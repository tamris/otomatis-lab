import os
import sys
import re
import json
import time
import shutil
import io
import argparse
from pathlib import Path
import openpyxl
from openpyxl.styles import Border, Side, PatternFill
from PIL import Image
from dotenv import load_dotenv
from mailmerge import MailMerge
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

client = genai.Client(api_key=API_KEY)


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
   - PENTING TENTANG BARIS DENGAN HASIL LAB KOSONG:
     * Jika suatu baris memiliki NAMA atau NOMOR pasien (misal No 4 AROFAH, No 5 MUCHAMAD WILDANUL MUNIR, No 6 ARIF RAHMAN HAKIM, dll.), TETAP EKSTRAK baris tersebut! Tuliskan NO dan NAMA-nya, sedangkan kolom-kolom nilai pemeriksaannya yang kosong cukup isi null. JANGAN PERNAH MELEWATKAN baris yang memiliki nama atau nomor pasien!
     * HANYA abaikan baris yang BENAR-BENAR KOSONG MELOMPONG (tidak ada nama, tidak ada nomor, dan tidak ada nilai lab apa pun).
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
     * "Basofil": nilai Basofil (jika strip "-" abaikan/null)
     * "Eosinofil": nilai Eosinofil (misal 1 atau 2)
     * "Netrofil Batang": nilai dari kolom 'Neutrofil_Bat' / 'Netrofil Batang'
     * "Netrofil Seg": nilai dari kolom 'Neutrofil_Seg' / 'Netrofil Seg' (misal 50)
     * "Limfosit": nilai Limfosit (misal 45)
     * "Monosit": nilai Monosit (misal 3)
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
      "BAKTERI": "..."
    }
  ]
}
"""


def analyze_image(img_path):
    """Ekstraksi Vision super cepat dengan fallback multi-model jika limit kuota/server sibuk."""
    jpeg_bytes = prepare_image_bytes_for_vision(img_path, max_dim=1200, quality=80)
    image_part = types.Part.from_bytes(data=jpeg_bytes, mime_type="image/jpeg")

    models_to_try = [
        ("gemini-3.8-flash", None),
        ("gemini-3.5-flash-lite", None),
        ("gemini-2.5-flash", None),
    ]
    last_err = None
    for model_name, tb in models_to_try:
        cfg = types.GenerateContentConfig(
            response_mime_type="application/json",
            thinking_config=types.ThinkingConfig(thinking_budget=tb) if tb is not None else None
        )
        for attempt in range(2):
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
                if attempt == 0 and ("503" in err_str or "UNAVAILABLE" in err_str):
                    time.sleep(2)
                    continue
                print(f" [WARN] Model {model_name} dialihkan ({type(e).__name__}). Mencoba model berikutnya...", flush=True)
                time.sleep(1)
                break

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
                if cell_nama == target_nama or SequenceMatcher(None, cell_nama, target_nama).ratio() >= 0.50:
                    return r

    # 3. Prioritas Ketiga: Cocokkan berdasarkan NAMA (hanya jika NO TIDAK bertentangan)
    if target_nama and nama_col:
        for r in range(2, ws.max_row + 1):
            if pkm_col and norm_target_pkm:
                cell_pkm = str(ws.cell(r, pkm_col).value or "").strip().upper()
                if cell_pkm and cell_pkm != norm_target_pkm:
                    continue

            cell_no = clean_no(ws.cell(r, no_col).value) if no_col else None
            # Jika kedua baris memiliki NO yang berbeda, JELAS pasien berbeda! Jangan cocokkan!
            if target_no and cell_no and target_no != cell_no:
                continue

            cell_nama = normalize_name(ws.cell(r, nama_col).value)
            if cell_nama:
                if cell_nama == target_nama:
                    return r
                # Fuzzy ketat (>= 0.92) hanya untuk typo bacaan OCR minor
                if SequenceMatcher(None, cell_nama, target_nama).ratio() >= 0.92:
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

    # 3. Parameter dengan 2 angka desimal (Eritrosit)
    if "ERITROSIT" in field_upper and ("DARAH" in field_upper or field_name == "Eritrosit"):
        s_clean = s.replace(",", ".")
        try:
            f = float(s_clean)
            return f"{f:.2f}"
        except ValueError:
            return s

    # 4. Parameter dengan 1 angka desimal (PH, Hemoglobin, Lekosit, Creatinine, HbA1c, Hematokrit, MCV, MCH, MCHC)
    decimal_1_fields = [
        "PH", "HEMOGLOBIN", "HB", "LEKOSIT", "CREATININE", "CR",
        "HBA1C", "HEMATOKRIT", "MCV", "MCH", "MCHC"
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


def apply_excel_styling(wb):
    """
    Format otomatis tampilan Excel:
    1. Mengurutkan seluruh baris data berdasarkan NO (1, 2, 3, ...) secara ascending.
    2. Memberi border tipis (All Borders) ke seluruh tabel data & header.
    3. Memberi highlight warna kuning lembut (soft yellow) jika pasien belum memiliki hasil lab.
       Jika hasil lab sudah terisi, warna kuning otomatis dibersihkan.
    4. Menstandarkan kolom BERAT JENIS di sheet URIN menjadi teks '1.xxx'.
    """
    for s_name in ["DARAH", "URIN"]:
        if s_name not in wb.sheetnames:
            continue
        ws = wb[s_name]
        if ws.max_row < 1:
            continue

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
        elif s_name == "DARAH":
            for c in range(1, ws.max_column + 1):
                h = ws.cell(1, c).value
                if not h:
                    continue
                h_clean = clean_col_name(h)
                if any(k in h_clean for k in ["hemoglobin", "lekosit", "hematokrit", "mcv", "mch", "mchc", "creatinine", "hba1c"]):
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

    col_map = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(1, col).value
        if val:
            col_map[clean_col_name(val)] = col

    urin_lekosit_kimia_col = 13
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
            safe_set(col_map.get("tanggalexam"), p.get("TANGGAL_EXAM") or config.get("TANGGAL_EXAM"))
            safe_set(col_map.get("puskesmas"), nama_pkm)
            safe_set(col_map.get("tanggalsurat"), p.get("TANGGAL_SURAT") or config.get("TANGGAL_SURAT"))
        else:
            # SUDAH ADA: Perbarui kolom yang sebelumnya masih kosong
            if p.get("NO"): update_if_empty(col_map.get("no"), p.get("NO"))
            if p.get("NAMA"): update_if_empty(col_map.get("nama"), p.get("NAMA"))
            if p.get("UMUR"): update_if_empty(col_map.get("umur"), p.get("UMUR"))
            if p.get("KD_PORSI"): update_if_empty(col_map.get("kdporsi"), p.get("KD_PORSI"))
            update_if_empty(col_map.get("tanggalexam"), p.get("TANGGAL_EXAM") or config.get("TANGGAL_EXAM"))
            update_if_empty(col_map.get("puskesmas"), nama_pkm)
            update_if_empty(col_map.get("tanggalsurat"), p.get("TANGGAL_SURAT") or config.get("TANGGAL_SURAT"))

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

    apply_excel_styling(wb)
    try:
        wb.save(rekap_path)
        return True
    except PermissionError:
        print(f"\n [ERROR] Gagal menyimpan '{rekap_path.name}' karena sedang DIBUKA di Microsoft Excel! Harap TUTUP file tersebut.", flush=True)
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
        ("GLUKOSA", ["GLUKOSA", "glukosa"]),
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
            col_map[clean_col_name(v)] = c

    nama_col = col_map.get("nama")
    kd_col = col_map.get("kdporsi")
    no_col = col_map.get("no")
    umur_col = col_map.get("umur")
    tgl_exam_col = col_map.get("tanggalexam")
    pkm_col = col_map.get("puskesmas")
    tgl_surat_col = col_map.get("tanggalsurat")

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
        else:
            lab_keys = [
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
                ("EPITEL", "epitel"),
                ("LEKOSIT_SEDIMEN", "lekosit_sedimen"),
                ("ERITROSIT", "eritrosit"),
                ("SILINDER", "silinder"),
                ("KRISTAL", "kristal"),
                ("BAKTERI", "bakteri"),
                ("PP TEST", "pptest"),
            ]
            for dict_k, col_k in lab_keys:
                if dict_k == "LEKOSIT_SEDIMEN":
                    c_idx = 22
                elif dict_k == "LEKOSIT_KIMIA":
                    c_idx = 13
                else:
                    c_idx = col_map.get(col_k)
                cell = ws.cell(r, c_idx) if c_idx else None
                p_data[dict_k] = format_cell_value(cell.value, cell.number_format, dict_k) if cell and cell.value is not None else None

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
        "TANGGAL_EXAM": str(p.get("TANGGAL_EXAM") or config.get("TANGGAL_EXAM", "")),
        "TANGGAL_SURAT": str(p.get("TANGGAL_SURAT") or config.get("TANGGAL_SURAT", "")),
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
        "Basofil": str(p.get("Basofil", "") or ""),
        "Netrofil_Batang": str(p.get("Netrofil Batang", "") or ""),
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
    return {
        "NAMA": str(p.get("NAMA", "") or ""),
        "umur": format_cell_value(p.get("UMUR", ""), field_name="UMUR"),
        "KD_PORSI": format_cell_value(p.get("KD_PORSI", ""), field_name="KD_PORSI"),
        "PUSKESMAS": str(p.get("PUSKESMAS") or nama_pkm or ""),
        "TANGGAL_EXAM": str(p.get("TANGGAL_EXAM") or config.get("TANGGAL_EXAM", "")),
        "TANGGAL_SURAT": str(p.get("TANGGAL_SURAT") or config.get("TANGGAL_SURAT", "")),
        "WARNA": str(p.get("WARNA", "") or ""),
        "KEJERNIHAN": str(p.get("KEJERNIHAN", "") or ""),
        "DARAH": str(p.get("DARAH", "") or ""),
        "BERAT_JENIS": format_cell_value(p.get("BERAT JENIS") or p.get("BERAT_JENIS") or "", field_name="BERAT_JENIS"),
        "PH": format_cell_value(p.get("PH", "") or "", field_name="PH"),
        "LEKOSIT": str(p.get("LEKOSIT_KIMIA") or p.get("LEKOSIT") or ""),
        "NITRIT": str(p.get("NITRIT", "") or ""),
        "GLUKOSA": str(p.get("GLUKOSA", "") or ""),
        "PROTEIN": str(p.get("PROTEIN", "") or ""),
        "UROBILINOGEN": str(p.get("UROBILINOGEN", "") or ""),
        "BILIRUBIN": str(p.get("BILIRUBIN", "") or ""),
        "BLOOD": str(p.get("BLOOD", "") or ""),
        "KETON": str(p.get("KETON", "") or ""),
        "EPITEL": str(p.get("EPITEL", "") or ""),
        "LEKOSIT1": str(p.get("LEKOSIT_SEDIMEN") or p.get("LEKOSIT1") or ""),
        "ERITROSIT": str(p.get("ERITROSIT", "") or ""),
        "KRISTAL": str(p.get("KRISTAL", "") or ""),
        "BAKTERI": str(p.get("BAKTERI", "") or ""),
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
    wb = openpyxl.load_workbook(rekap_excel_path, data_only=True)
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
            param_keys_urin = {"WARNA", "KEJERNIHAN", "DARAH", "BERAT JENIS", "PH", "LEKOSIT_KIMIA", "NITRIT", "GLUKOSA", "PROTEIN", "EPITEL"}
            check_keys = param_keys_darah if jenis == "DARAH" else param_keys_urin

            for p in extracted_patients:
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
def mode_generate(config, word_app=None, target_pkm=None):
    """
    MODE CETAK:
    1. Membaca data yang SUDAH direview dari 'Template Exel.xlsx' (atau Rekap per Puskesmas).
    2. Periksa kolom parameter lab: Jika masih ada kolom wajib yang kosong, LEWATI (skip).
    3. Untuk pasien yang SELURUH parameter labnya lengkap: buat dokumen Word & PDF All-in-One.
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
        pkm_dirs = [d for d in pkm_dirs if d.name.strip().upper() == target_pkm.strip().upper()]

    if not pkm_dirs:
        print(f" [INFO] Tidak ada subfolder Puskesmas ditemukan di '{hasil_base.name}/'.", flush=True)
        return

    print(f" [INFO] Terdeteksi {len(pkm_dirs)} Puskesmas dalam '{hasil_base.name}/': {', '.join(d.name for d in pkm_dirs)}", flush=True)

    for pkm_out_dir in pkm_dirs:
        nama_pkm = pkm_out_dir.name
        rekap_file = pkm_out_dir / f"Rekap_{nama_pkm}.xlsx"
        if not rekap_file.exists():
            print(f" [WARN] File '{rekap_file.name}' tidak ditemukan di '{nama_pkm}'. Dilewati.", flush=True)
            continue

        print("\n" + "=" * 55, flush=True)
        print(f" >>> MEMPROSES CETAK: {nama_pkm} <<<", flush=True)
        print("=" * 55, flush=True)

        # Standardisasi styling dan format nilai Excel sebelum pembacaan & pencetakan
        try:
            wb_style = openpyxl.load_workbook(rekap_file)
            apply_excel_styling(wb_style)
            wb_style.save(rekap_file)
        except Exception:
            pass

        try:
            wb = openpyxl.load_workbook(rekap_file, data_only=True)
        except PermissionError:
            print(f" [ERROR] Tidak dapat membuka '{rekap_file.name}' karena sedang dibuka di Microsoft Excel. Mohon tutup file tersebut terlebih dahulu.", flush=True)
            continue

        ws_darah = wb["DARAH"] if "DARAH" in wb.sheetnames else None
        ws_urin = wb["URIN"] if "URIN" in wb.sheetnames else None

        darah_list = read_patients_from_sheet(ws_darah, jenis="DARAH") if ws_darah else []
        urin_list = read_patients_from_sheet(ws_urin, jenis="URIN") if ws_urin else []
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
                out_docx = pkm_out_dir / f"All_Hasil_Darah_{nama_pkm}.docx"
                out_pdf = pkm_out_dir / f"All_Hasil_Darah_{nama_pkm}.pdf"

                darah_dicts = [prepare_darah_dict(p, config, nama_pkm) for p in valid_darah]
                with MailMerge(template_darah) as mm:
                    mm.merge_templates(darah_dicts, separator="page_break")
                    mm.write(str(out_docx))

                convert_single_docx_to_pdf_fast(word_app, out_docx, out_pdf)
                print(f" [PDF] Selesai: All_Hasil_Darah_{nama_pkm}.pdf ({len(valid_darah)} pasien valid) - {time.time() - t0:.1f} detik", flush=True)
            else:
                print(f" [INFO] Tidak ada pasien DARAH dengan parameter lab lengkap untuk {nama_pkm}.", flush=True)

        # 2. Validasi & Cetak URIN (urin_list sudah dimuat dari ws_urin)
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
                out_docx = pkm_out_dir / f"All_Hasil_Urin_{nama_pkm}.docx"
                out_pdf = pkm_out_dir / f"All_Hasil_Urin_{nama_pkm}.pdf"

                urin_dicts = [prepare_urin_dict(p, config, nama_pkm) for p in valid_urin]
                with MailMerge(template_urin) as mm:
                    mm.merge_templates(urin_dicts, separator="page_break")
                    mm.write(str(out_docx))

                convert_single_docx_to_pdf_fast(word_app, out_docx, out_pdf)
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
            mode_extract(config, ref_by_no, ref_by_name)
        elif args.generate:
            mode_generate(config, word_app=word_app)
        else:
            # Jika dijalankan tanpa argumen
            input_base_dir = BASE_DIR / "foto_masuk"
            pkm_subfolders = [d for d in input_base_dir.iterdir() if d.is_dir()] if input_base_dir.exists() else []
            if pkm_subfolders:
                print("\n [INFO] Menjalankan default: MODE EKSTRAKSI (--extract)...", flush=True)
                mode_extract(config, ref_by_no, ref_by_name)
            else:
                print("\n [INFO] 'foto_masuk/' kosong. Menjalankan default: MODE CETAK (--generate)...", flush=True)
                mode_generate(config, word_app=word_app)
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
