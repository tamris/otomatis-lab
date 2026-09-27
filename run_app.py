import sys
import time
import socket
import webbrowser
import subprocess
import threading
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

def kill_port(port=8000):
    """Mematikan proses apa pun yang sedang menduduki port 8000 agar tidak bentrok."""
    try:
        cmd = f'powershell -Command "Get-NetTCPConnection -LocalPort {port} -ErrorAction SilentlyContinue | ForEach-Object {{ Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }}"'
        subprocess.run(cmd, shell=True, capture_output=True)
    except Exception:
        pass

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8000")

if __name__ == "__main__":
    print("\n" + "=" * 58)
    print("       SISTEM OTOMASI LAB KLINIS - WEB DASHBOARD")
    print("=" * 58)

    # 1. Bersihkan port 8000 jika ada server lama nyangkut
    kill_port(8000)
    time.sleep(0.5)

    # 2. Buka browser otomatis
    threading.Thread(target=open_browser, daemon=True).start()

    print(" [1/2] Server aktif di   : http://127.0.0.1:8000")
    print(" [2/2] Membuka antarmuka di browser Anda...")
    print("\n Dashboard siap digunakan! Biarkan jendela ini tetap terbuka.")
    print(" (Tekan Ctrl + C untuk mematikan server)\n" + "-" * 58 + "\n")

    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, log_level="info")
