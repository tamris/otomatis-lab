import os
import sys
import shutil
import subprocess
import threading
from pathlib import Path
from typing import Optional, List
import openpyxl
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware

import otomasi_lab
import reset_tool

BASE_DIR = Path(__file__).resolve().parent

app = FastAPI(title="Sistem Otomasi Lab Klinis", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Lock untuk mencegah race condition ekstraksi bersamaan
task_lock = threading.Lock()
current_task = {
    "status": "idle",  # idle, extracting, generating, done, error
    "message": "Sistem siap digunakan.",
    "progress": 0,
    "last_error": None
}


def is_file_locked(file_path: Path) -> bool:
    """Mengecek apakah file Excel sedang dibuka dan dikunci oleh proses lain untuk mode WRITE."""
    if not file_path.exists():
        return False
    try:
        with open(file_path, "r+b"):
            return False
    except (PermissionError, IOError):
        return True


def read_excel_shared(file_path: Path):
    """
    Membaca file Excel secara aman bahkan jika sedang dibuka di Microsoft Excel.
    Menggunakan Windows CreateFile dengan FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE.
    """
    try:
        import io, win32file, win32con
        handle = win32file.CreateFile(
            str(file_path.resolve()),
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None
        )
        size = win32file.GetFileSize(handle)
        _, data = win32file.ReadFile(handle, size)
        win32file.CloseHandle(handle)
        return openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        return openpyxl.load_workbook(file_path, data_only=True)


def get_all_puskesmas_list() -> List[str]:
    """Mendapatkan daftar seluruh Puskesmas yang ada di foto_masuk dan HASIL_PUSKESMAS."""
    pkm_set = set()
    fm = BASE_DIR / "foto_masuk"
    if fm.exists():
        for d in fm.iterdir():
            if d.is_dir() and not d.name.startswith("."):
                pkm_set.add(d.name)

    hp = BASE_DIR / "HASIL_PUSKESMAS"
    if hp.exists():
        for d in hp.iterdir():
            if d.is_dir() and not d.name.startswith("."):
                pkm_set.add(d.name)

    fa = BASE_DIR / "foto_arsip"
    if fa.exists():
        for d in fa.iterdir():
            if d.is_dir() and not d.name.startswith("."):
                pkm_set.add(d.name)

    # Tambahkan default dari config jika ada
    cfg = otomasi_lab.load_config()
    def_pkm = cfg.get("PUSKESMAS")
    if def_pkm:
        pkm_set.add(def_pkm)

    return sorted(list(pkm_set))


def get_pkm_stats(pkm_name: str):
    """Mendapatkan statistik foto dan data rekap untuk Puskesmas tertentu."""
    fm_dir = BASE_DIR / "foto_masuk" / pkm_name
    fa_dir = BASE_DIR / "foto_arsip" / pkm_name
    hp_dir = BASE_DIR / "HASIL_PUSKESMAS" / pkm_name
    rekap_file = hp_dir / f"Rekap_{pkm_name}.xlsx"

    img_exts = {".jpg", ".jpeg", ".png", ".webp"}
    photos_queue = []
    if fm_dir.exists():
        photos_queue = [f.name for f in fm_dir.iterdir() if f.is_file() and f.suffix.lower() in img_exts]

    archived_count = 0
    if fa_dir.exists():
        archived_count = len([f for f in fa_dir.iterdir() if f.is_file() and f.suffix.lower() in img_exts])

    rekap_exists = rekap_file.exists()
    locked = is_file_locked(rekap_file)

    total_patients = 0
    complete_count = 0
    pending_count = 0

    if rekap_exists:
        try:
            wb = read_excel_shared(rekap_file)
            for s_name in ["DARAH", "URIN"]:
                if s_name in wb.sheetnames:
                    ws = wb[s_name]
                    for r in range(2, ws.max_row + 1):
                        nama_v = ws.cell(r, 2).value
                        no_v = ws.cell(r, 1).value
                        if not nama_v and not no_v:
                            continue
                        total_patients += 1
                        has_lab = any(
                            ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip() != ""
                            for c in range(8, ws.max_column + 1)
                        )
                        if has_lab:
                            complete_count += 1
                        else:
                            pending_count += 1
        except Exception:
            pass

    docx_exists = (hp_dir / f"All_Hasil_Darah_{pkm_name}.docx").exists()
    pdf_exists = (hp_dir / f"All_Hasil_Darah_{pkm_name}.pdf").exists()

    return {
        "pkm_name": pkm_name,
        "queue_photos": photos_queue,
        "queue_count": len(photos_queue),
        "archived_count": archived_count,
        "rekap_exists": rekap_exists,
        "is_locked": locked,
        "total_patients": total_patients,
        "complete_count": complete_count,
        "pending_count": pending_count,
        "docx_exists": docx_exists,
        "pdf_exists": pdf_exists,
        "pdf_filename": f"All_Hasil_Darah_{pkm_name}.pdf" if pdf_exists else None
    }


# ---------------------------------------------------------------------------
# API Endpoints
# ---------------------------------------------------------------------------

@app.get("/api/system-status")
def api_system_status():
    """Mengambil status global sistem, daftar Puskesmas, dan konfigurasi."""
    cfg = otomasi_lab.load_config()
    pkm_list = get_all_puskesmas_list()
    pkm_stats = {pkm: get_pkm_stats(pkm) for pkm in pkm_list}

    tmpl_file = BASE_DIR / "Template Exel.xlsx"
    tmpl_clean = True
    if tmpl_file.exists():
        try:
            wb = openpyxl.load_workbook(tmpl_file, data_only=True)
            for s in ["DARAH", "URIN"]:
                if s in wb.sheetnames and wb[s].max_row > 1:
                    tmpl_clean = False
        except Exception:
            pass

    return {
        "config": cfg,
        "puskesmas_list": pkm_list,
        "puskesmas_stats": pkm_stats,
        "template_clean": tmpl_clean,
        "current_task": current_task
    }


@app.get("/api/patients")
def api_get_patients(pkm: str):
    """Membaca daftar pasien dari Rekap_[pkm].xlsx untuk ditampilkan di tabel preview."""
    rekap_file = BASE_DIR / "HASIL_PUSKESMAS" / pkm / f"Rekap_{pkm}.xlsx"
    if not rekap_file.exists():
        return {"darah": [], "urin": [], "is_locked": False, "exists": False}

    locked = is_file_locked(rekap_file)
    try:
        wb = read_excel_shared(rekap_file)
    except Exception:
        return {"darah": [], "urin": [], "is_locked": locked, "exists": True}

    results = {"darah": [], "urin": [], "is_locked": locked, "exists": True}

    if "DARAH" in wb.sheetnames:
        ws = wb["DARAH"]
        for r in range(2, ws.max_row + 1):
            nama_v = ws.cell(r, 2).value
            no_v = ws.cell(r, 1).value
            if not nama_v and not no_v:
                continue

            has_lab = any(
                ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip() != ""
                for c in range(8, ws.max_column + 1)
            )

            results["darah"].append({
                "no": ws.cell(r, 1).value,
                "nama": ws.cell(r, 2).value,
                "umur": ws.cell(r, 3).value,
                "kd_porsi": ws.cell(r, 4).value,
                "golda": ws.cell(r, 23).value,
                "glukosa_puasa": ws.cell(r, 24).value,
                "glukosa_2jpp": ws.cell(r, 25).value,
                "chol_total": ws.cell(r, 26).value,
                "trigliserida": ws.cell(r, 27).value,
                "sgot": ws.cell(r, 28).value,
                "sgpt": ws.cell(r, 29).value,
                "ureum": ws.cell(r, 30).value,
                "creatinine": ws.cell(r, 31).value,
                "hba1c": ws.cell(r, 32).value,
                "has_lab": has_lab
            })

    if "URIN" in wb.sheetnames:
        ws = wb["URIN"]
        for r in range(2, ws.max_row + 1):
            nama_v = ws.cell(r, 2).value
            no_v = ws.cell(r, 1).value
            if not nama_v and not no_v:
                continue

            has_lab = any(
                ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip() != ""
                for c in range(8, ws.max_column + 1)
            )

            results["urin"].append({
                "no": ws.cell(r, 1).value,
                "nama": ws.cell(r, 2).value,
                "umur": ws.cell(r, 3).value,
                "kd_porsi": ws.cell(r, 4).value,
                "warna": ws.cell(r, 8).value,
                "kejernihan": ws.cell(r, 9).value,
                "protein": ws.cell(r, 16).value,
                "glukosa": ws.cell(r, 15).value,
                "has_lab": has_lab
            })

    return results


@app.post("/api/upload")
async def api_upload(pkm: str = Form(...), files: List[UploadFile] = File(...)):
    """Menyimpan foto yang di-drag & drop langsung ke folder foto_masuk/[pkm]/."""
    pkm = pkm.strip()
    target_dir = BASE_DIR / "foto_masuk" / pkm
    target_dir.mkdir(parents=True, exist_ok=True)

    saved_files = []
    for file in files:
        if not file.filename:
            continue
        dest_path = target_dir / file.filename
        with open(dest_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        saved_files.append(file.filename)

    return {"status": "success", "pkm": pkm, "saved_files": saved_files, "total": len(saved_files)}


def run_extraction_worker(pkm: Optional[str] = None):
    global current_task
    with task_lock:
        current_task["status"] = "extracting"
        current_task["message"] = f"Mengekstrak foto untuk {pkm or 'Seluruh Puskesmas'}..."
        current_task["last_error"] = None

    try:
        config = otomasi_lab.load_config()
        # Muat database rujukan lokal jika ada
        ref_by_no, ref_by_name = otomasi_lab.load_reference_patients()

        otomasi_lab.mode_extract(config, ref_by_no, ref_by_name, target_pkm=pkm)

        with task_lock:
            current_task["status"] = "done"
            current_task["message"] = f"Ekstraksi foto untuk {pkm or 'Puskesmas'} selesai! Data tersimpan di Excel."
    except Exception as e:
        import traceback
        traceback.print_exc()
        with task_lock:
            current_task["status"] = "error"
            current_task["message"] = f"Terjadi kesalahan saat ekstraksi: {str(e)}"
            current_task["last_error"] = str(e)


@app.post("/api/extract")
def api_extract(pkm: Optional[str] = Form(None), background_tasks: BackgroundTasks = BackgroundTasks()):
    """Memicu proses ekstraksi AI Gemini untuk foto di folder foto_masuk."""
    if pkm:
        rekap_file = BASE_DIR / "HASIL_PUSKESMAS" / pkm / f"Rekap_{pkm}.xlsx"
        if is_file_locked(rekap_file):
            raise HTTPException(
                status_code=400,
                detail=f"File 'Rekap_{pkm}.xlsx' sedang dibuka di Microsoft Excel! Harap simpan dan tutup file tersebut terlebih dahulu."
            )

    if current_task["status"] in ["extracting", "generating"]:
        raise HTTPException(status_code=409, detail="Proses lain sedang berjalan. Harap tunggu.")

    background_tasks.add_task(run_extraction_worker, pkm)
    return {"status": "started", "target_pkm": pkm, "message": "Proses ekstraksi dimulai di background."}


def run_generate_worker(pkm: Optional[str] = None):
    global current_task
    with task_lock:
        current_task["status"] = "generating"
        current_task["message"] = f"Membuat dokumen Word & PDF untuk {pkm or 'Seluruh Puskesmas'}..."
        current_task["last_error"] = None

    try:
        config = otomasi_lab.load_config()
        otomasi_lab.mode_generate(config, target_pkm=pkm)

        with task_lock:
            current_task["status"] = "done"
            current_task["message"] = "Dokumen Word & PDF All-in-One berhasil dicetak!"
    except Exception as e:
        with task_lock:
            current_task["status"] = "error"
            current_task["message"] = f"Terjadi kesalahan saat cetak: {str(e)}"
            current_task["last_error"] = str(e)


@app.post("/api/generate")
def api_generate(pkm: Optional[str] = Form(None), background_tasks: BackgroundTasks = BackgroundTasks()):
    """Memicu validasi kelengkapan data & generate Word + PDF."""
    if current_task["status"] in ["extracting", "generating"]:
        raise HTTPException(status_code=409, detail="Proses lain sedang berjalan. Harap tunggu.")

    background_tasks.add_task(run_generate_worker, pkm)
    return {"status": "started", "target_pkm": pkm, "message": "Proses cetak dimulai di background."}


@app.post("/api/reset-pkm")
def api_reset_pkm(pkm: str = Form(...)):
    """Mereset data Puskesmas (kembalikan foto arsip ke antrean & bersihkan rekap)."""
    pkm = pkm.strip()
    rekap_file = BASE_DIR / "HASIL_PUSKESMAS" / pkm / f"Rekap_{pkm}.xlsx"
    if is_file_locked(rekap_file):
        raise HTTPException(
            status_code=400,
            detail=f"File 'Rekap_{pkm}.xlsx' sedang dibuka di Microsoft Excel! Tutup file terlebih dahulu sebelum mereset."
        )

    try:
        reset_tool.reset_puskesmas(pkm)
        return {"status": "success", "message": f"Data '{pkm}' berhasil direset total."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/open-folder")
def api_open_folder(pkm: str = Form(...), folder_type: str = Form("hasil")):
    """Membuka folder langsung di Windows Explorer menggunakan native ShellExecute."""
    pkm = pkm.strip()
    if folder_type == "hasil":
        target = BASE_DIR / "HASIL_PUSKESMAS" / pkm
    elif folder_type == "masuk":
        target = BASE_DIR / "foto_masuk" / pkm
    elif folder_type == "arsip":
        target = BASE_DIR / "foto_arsip" / pkm
    else:
        target = BASE_DIR

    target.mkdir(parents=True, exist_ok=True)
    try:
        os.startfile(str(target.resolve()))
        return {"status": "opened", "path": str(target), "message": f"Folder {folder_type} ({pkm}) berhasil dibuka di Windows Explorer."}
    except Exception as e:
        subprocess.Popen(f'explorer "{target.resolve()}"')
        return {"status": "opened", "path": str(target), "message": f"Folder {folder_type} ({pkm}) dibuka via explorer."}


@app.post("/api/open-excel")
def api_open_excel(pkm: str = Form(...)):
    """Membuka file Rekap_[pkm].xlsx langsung di Microsoft Excel. Otomatis buat file dari template jika belum ada."""
    pkm = pkm.strip()
    target_dir = BASE_DIR / "HASIL_PUSKESMAS" / pkm
    target_dir.mkdir(parents=True, exist_ok=True)
    rekap_file = target_dir / f"Rekap_{pkm}.xlsx"

    # Jika file belum ada, inisialisasi dari Template Exel.xlsx agar tidak error 404
    if not rekap_file.exists():
        cfg = otomasi_lab.load_config()
        tmpl = BASE_DIR / "Template Exel.xlsx"
        ref_by_no = None
        if "BALAPULANG" in pkm.upper():
            ref_by_no, _ = otomasi_lab.load_reference_patients()
        otomasi_lab.init_rekap_excel(rekap_file, tmpl, pkm, cfg, ref_by_no=ref_by_no)

    try:
        os.startfile(str(rekap_file.resolve()))
        return {"status": "opened", "path": str(rekap_file), "message": f"File 'Rekap_{pkm}.xlsx' berhasil dibuka di Microsoft Excel."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal membuka Microsoft Excel: {str(e)}")


@app.post("/api/create-puskesmas")
def api_create_puskesmas(pkm: str = Form(...)):
    """Membuat folder kerja dan file rekap untuk Puskesmas baru."""
    pkm = pkm.strip().upper()
    if not pkm.startswith("PUSKESMAS"):
        pkm = f"PUSKESMAS {pkm}"

    (BASE_DIR / "foto_masuk" / pkm).mkdir(parents=True, exist_ok=True)
    (BASE_DIR / "foto_arsip" / pkm).mkdir(parents=True, exist_ok=True)
    target_hasil = BASE_DIR / "HASIL_PUSKESMAS" / pkm
    target_hasil.mkdir(parents=True, exist_ok=True)

    rekap_file = target_hasil / f"Rekap_{pkm}.xlsx"
    if not rekap_file.exists():
        cfg = otomasi_lab.load_config()
        tmpl = BASE_DIR / "Template Exel.xlsx"
        ref_by_no = None
        if "BALAPULANG" in pkm:
            ref_by_no, _ = otomasi_lab.load_reference_patients()
        otomasi_lab.init_rekap_excel(rekap_file, tmpl, pkm, cfg, ref_by_no=ref_by_no)

    return {"status": "success", "pkm": pkm, "message": f"{pkm} berhasil ditambahkan dan siap digunakan."}


@app.post("/api/dismiss-error")
def api_dismiss_error():
    """Mereset status task dari error kembali ke idle."""
    global current_task
    with task_lock:
        current_task["status"] = "idle"
        current_task["message"] = "Sistem siap digunakan."
        current_task["last_error"] = None
    return {"status": "success"}


@app.post("/api/close-excel")
def api_close_excel():
    """Menutup proses Microsoft Excel yang sedang berjalan di background."""
    try:
        subprocess.run("taskkill /F /IM excel.exe", shell=True, capture_output=True)
        return {"status": "success", "message": "Microsoft Excel berhasil ditutup."}
    except Exception as e:
        return {"status": "error", "message": str(e)}


@app.get("/api/download-pdf/{pkm}")
def api_download_pdf(pkm: str):
    """Mengunduh file PDF All-in-One hasil cetak."""
    pdf_path = BASE_DIR / "HASIL_PUSKESMAS" / pkm / f"All_Hasil_Darah_{pkm}.pdf"
    if not pdf_path.exists():
        raise HTTPException(status_code=404, detail="File PDF belum dicetak.")
    return FileResponse(path=str(pdf_path), filename=f"All_Hasil_Darah_{pkm}.pdf", media_type="application/pdf")


# ---------------------------------------------------------------------------
# Serve Frontend HTML
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
def index():
    html_file = BASE_DIR / "templates" / "index.html"
    if html_file.exists():
        return html_file.read_text(encoding="utf-8")
    return "<h1>Frontend sedang disiapkan...</h1>"


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
