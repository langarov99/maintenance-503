"""Bot launcher — finds a free port (5000-5010) and starts uvicorn on it."""
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser

ROOT = os.path.dirname(os.path.abspath(__file__))


def find_free_port(start: int = 5000, end: int = 5010) -> int:
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError(f"No free port found between {start} and {end}")


def open_browser(port: int):
    url = f"http://127.0.0.1:{port}/"
    for _ in range(40):
        time.sleep(0.5)
        try:
            urllib.request.urlopen(url, timeout=1)
            webbrowser.open(f"http://localhost:{port}")
            return
        except Exception:
            pass


if __name__ == "__main__":
    port = find_free_port()

    print()
    print(" ================================================")
    print(f"  Data Extraction Bot  -  http://localhost:{port}")
    print(" ================================================")
    print("  Zatvorete tozi prozorec za da spirete bota.")
    print()

    threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    os.chdir(ROOT)
    subprocess.run([
        sys.executable, "-m", "uvicorn", "app.main:app",
        "--host", "127.0.0.1",
        "--port", str(port),
        "--app-dir", ROOT,
    ])
