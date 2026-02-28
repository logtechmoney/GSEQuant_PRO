#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GSEQuant Launcher v1.0
======================
Launcher liviano que:
1. Verifica si hay una nueva version disponible en GitHub
2. Descarga y actualiza automaticamente si es necesario
3. Lanza GSEQuant_PRO.exe

Requisitos para compilar este launcher:
    pip install pyinstaller
    pyinstaller launcher.spec

El launcher compilado pesa ~5-8MB y NO necesita actualizarse nunca
(a menos que cambies la URL de version.json).
"""

import sys
import os
import json
import queue
import shutil
import zipfile
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACION — Cambia estos valores con tu usuario/repo de GitHub
# ─────────────────────────────────────────────────────────────────────────────

# URL del archivo version.json publicado en tu repo (raw GitHub o Vercel/supabase/etc.)
VERSION_JSON_URL = (
    "https://raw.githubusercontent.com/logtechmoney/GSEQuant_PRO/main/version.json"
)

# Nombre del ejecutable principal de la app
APP_EXE_NAME = "GSEQuant_PRO.exe"

# Nombre del archivo local de version instalada
LOCAL_VERSION_FILE = "installed_version.json"

# Timeout en segundos para conexion al servidor
NETWORK_TIMEOUT = 10

# Version del launcher (no se actualiza; es fija en el .exe)
LAUNCHER_VERSION = "1.0.0"

# ─────────────────────────────────────────────────────────────────────────────


def get_base_dir() -> Path:
    """Directorio donde vive el launcher (y donde se instalara la app)."""
    if getattr(sys, "frozen", False):
        # Corriendo como .exe compilado
        return Path(sys.executable).parent
    else:
        # Corriendo como script .py
        return Path(__file__).parent


def get_installed_version(base_dir: Path) -> str:
    """Lee la version local instalada."""
    vfile = base_dir / LOCAL_VERSION_FILE
    try:
        with open(vfile, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("version", "0.0.0")
    except Exception:
        return "0.0.0"


def set_installed_version(base_dir: Path, version: str):
    """Guarda la version local instalada."""
    vfile = base_dir / LOCAL_VERSION_FILE
    with open(vfile, "w", encoding="utf-8") as f:
        json.dump({"version": version, "launcher_version": LAUNCHER_VERSION}, f)


def fetch_remote_version_info() -> Optional[dict]:
    """Obtiene el version.json del servidor."""
    try:
        import urllib.request
        with urllib.request.urlopen(VERSION_JSON_URL, timeout=NETWORK_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data
    except Exception:
        return None


def version_tuple(v: str):
    """Convierte '4.1.2' -> (4, 1, 2) para comparar."""
    try:
        return tuple(int(x) for x in str(v).split("."))
    except Exception:
        return (0, 0, 0)


def download_file(url: str, dest: Path, progress_callback=None) -> bool:
    """Descarga un archivo grande de forma robusta con chunking."""
    try:
        import urllib.request
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        with urllib.request.urlopen(req, timeout=30) as resp, open(dest, 'wb') as f:
            total_size_str = resp.getheader('Content-Length')
            total_size = int(total_size_str) if total_size_str else 0
            downloaded = 0
            block_size = 65536  # 64KB chunks
            
            while True:
                chunk = resp.read(block_size)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                
                if progress_callback and total_size > 0:
                    pct = min(100, int((downloaded / total_size) * 100))
                    progress_callback(pct, downloaded, total_size)
        return True
    except Exception as e:
        print(f"Error descargando {url}: {e}")
        return False


def extract_and_install(zip_path: Path, dest_dir: Path, progress_callback=None) -> bool:
    """Extrae un ZIP directamente en el directorio destino filtrando archivos de config."""
    KEEP_LOCAL = {"user_config.json", LOCAL_VERSION_FILE}
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            members = zf.infolist()
            total = len(members)
            for i, info in enumerate(members):
                file_name = Path(info.filename).name
                if file_name not in KEEP_LOCAL:
                    zf.extract(info, dest_dir)
                
                # Reportar progreso a la UI cada cierta cantidad para no saturar
                if progress_callback and i % 50 == 0:
                    progress_callback(int((i / total) * 100))
                    
            if progress_callback:
                progress_callback(100)
        return True
    except Exception as e:
        print(f"Error extrayendo ZIP: {e}")
        return False


def launch_app(base_dir: Path) -> bool:
    """Lanza el ejecutable principal de la app."""
    app_path = base_dir / APP_EXE_NAME
    if not app_path.exists():
        return False
    try:
        subprocess.Popen(
            [str(app_path)],
            cwd=str(base_dir),
            creationflags=subprocess.DETACHED_PROCESS if sys.platform == "win32" else 0,
        )
        return True
    except Exception as e:
        print(f"Error lanzando app: {e}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# GUI del Launcher
# ─────────────────────────────────────────────────────────────────────────────

class LauncherWindow(tk.Tk):
    """Ventana principal del launcher con barra de progreso y status."""

    COLORS = {
        "bg":         "#0d1117",
        "card":       "#161b22",
        "border":     "#30363d",
        "accent":     "#1f6feb",
        "accent2":    "#388bfd",
        "success":    "#3fb950",
        "warning":    "#d29922",
        "error":      "#f85149",
        "text":       "#e6edf3",
        "text_dim":   "#8b949e",
        "progress_bg":"#21262d",
    }

    def __init__(self):
        super().__init__()
        self.base_dir = get_base_dir()
        self._setup_window()
        self._build_ui()
        # Inicia el proceso de verificacion en hilo secundario
        threading.Thread(target=self._run_launcher, daemon=True).start()

    def _setup_window(self):
        self.title("GSEQuant Launcher")
        self.resizable(False, False)
        W, H = 520, 320
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - W) // 2
        y = (sh - H) // 2
        self.geometry(f"{W}x{H}+{x}+{y}")
        self.configure(bg=self.COLORS["bg"])
        self.overrideredirect(False)
        # Icono (si existe)
        icon_path = self.base_dir / "gse_app_icon.png"
        if icon_path.exists():
            try:
                img = tk.PhotoImage(file=str(icon_path))
                self.iconphoto(True, img)
            except Exception:
                pass

    def _build_ui(self):
        C = self.COLORS
        outer = tk.Frame(self, bg=C["bg"], padx=20, pady=20)
        outer.pack(fill="both", expand=True)

        # ── Encabezado ──────────────────────────────────────────────────────
        header = tk.Frame(outer, bg=C["card"], bd=0, relief="flat",
                          highlightthickness=1, highlightbackground=C["border"])
        header.pack(fill="x", pady=(0, 16))

        inner_h = tk.Frame(header, bg=C["card"], padx=20, pady=16)
        inner_h.pack(fill="x")

        tk.Label(
            inner_h, text="GSEQuant PRO", font=("Segoe UI", 20, "bold"),
            bg=C["card"], fg=C["text"]
        ).pack(anchor="w")

        tk.Label(
            inner_h, text="Ground Support Equipment Quantifier  ·  GTA",
            font=("Segoe UI", 10), bg=C["card"], fg=C["text_dim"]
        ).pack(anchor="w")

        # ── Panel de estado ──────────────────────────────────────────────────
        status_frame = tk.Frame(outer, bg=C["bg"])
        status_frame.pack(fill="x", pady=(0, 12))

        self.lbl_status = tk.Label(
            status_frame, text="Inicializando...",
            font=("Segoe UI", 10), bg=C["bg"], fg=C["text_dim"], anchor="w"
        )
        self.lbl_status.pack(fill="x")

        self.lbl_detail = tk.Label(
            status_frame, text="",
            font=("Segoe UI", 8), bg=C["bg"], fg=C["text_dim"], anchor="w"
        )
        self.lbl_detail.pack(fill="x")

        # ── Barra de progreso ────────────────────────────────────────────────
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "GSE.Horizontal.TProgressbar",
            troughcolor=C["progress_bg"],
            background=C["accent"],
            bordercolor=C["border"],
            lightcolor=C["accent2"],
            darkcolor=C["accent"],
            thickness=8,
        )
        self.progress = ttk.Progressbar(
            outer, style="GSE.Horizontal.TProgressbar",
            orient="horizontal", length=480, mode="indeterminate"
        )
        self.progress.pack(fill="x", pady=(0, 16))
        self.progress.start(12)

        # ── Version info ─────────────────────────────────────────────────────
        ver_frame = tk.Frame(outer, bg=C["bg"])
        ver_frame.pack(fill="x")

        installed = get_installed_version(self.base_dir)
        self.lbl_ver_installed = tk.Label(
            ver_frame, text=f"Instalada: v{installed}",
            font=("Segoe UI", 8), bg=C["bg"], fg=C["text_dim"]
        )
        self.lbl_ver_installed.pack(side="left")

        tk.Label(
            ver_frame, text=f"Launcher v{LAUNCHER_VERSION}",
            font=("Segoe UI", 8), bg=C["bg"], fg=C["text_dim"]
        ).pack(side="right")

    # ── Helpers para actualizar la UI desde el hilo ──────────────────────────

    def set_status(self, msg: str, detail: str = "", color: str = None):
        C = self.COLORS
        color = color or C["text_dim"]
        def _update():
            self.lbl_status.config(text=msg, fg=color)
            self.lbl_detail.config(text=detail)
        self.after(0, _update)

    def set_progress(self, pct: int, mode: str = "determinate"):
        def _update():
            self.progress.stop()
            self.progress.config(mode=mode)
            if mode == "determinate":
                self.progress["value"] = pct
            else:
                self.progress.start(12)
        self.after(0, _update)

    def set_progress_indeterminate(self):
        def _update():
            self.progress.config(mode="indeterminate")
            self.progress.start(12)
        self.after(0, _update)

    def update_installed_label(self, version: str):
        def _update():
            self.lbl_ver_installed.config(text=f"Instalada: v{version}")
        self.after(0, _update)

    # ── Lógica principal del launcher ────────────────────────────────────────

    def _run_launcher(self):
        """Flujo principal: verificar → (descargar) → lanzar.
        Corre en hilo secundario; se comunica con la UI via self.after().
        """
        C = self.COLORS
        base = self.base_dir
        installed_v = get_installed_version(base)

        # 1) Verificar conectividad y version remota
        self.set_status("Verificando actualizaciones...", "Conectando con el servidor...")
        remote_info = fetch_remote_version_info()

        if remote_info is None:
            # Sin internet — intentar lanzar lo que hay instalado
            self.set_status(
                "Sin conexión  —  Lanzando versión local",
                f"Versión instalada: {installed_v}",
                C["warning"]
            )
            self._try_launch_or_error(base)
            return

        remote_v = remote_info.get("version", "0.0.0")
        download_url = remote_info.get("download_url", "")
        release_notes = remote_info.get("release_notes", "")
        app_exe = remote_info.get("app_exe_name", APP_EXE_NAME)

        # 2) Comparar versiones
        needs_update = version_tuple(remote_v) > version_tuple(installed_v)
        app_exists = (base / app_exe).exists()

        if not needs_update and app_exists:
            self.set_status(
                "✓  App al día",
                f"Versión {installed_v}  ·  {release_notes}",
                C["success"]
            )
            self.set_progress(100)
            self.after(1200, lambda: self._launch_and_close(base, app_exe))
            return

        if not download_url:
            # No hay URL de descarga
            if app_exists:
                self.set_status(
                    "Sin URL de descarga  —  Lanzando versión local",
                    f"Versión instalada: {installed_v}",
                    C["warning"]
                )
                self._try_launch_or_error(base)
            else:
                self.set_status(
                    "Error: App no encontrada y sin descarga disponible",
                    "Contacta al administrador.",
                    C["error"]
                )
            return

        # 3) Hay actualización disponible → preguntar
        action = self._ask_update(installed_v, remote_v, release_notes)
        if action == "skip" and app_exists:
            self.set_status("Omitiendo actualización...", "", C["warning"])
            self._launch_and_close(base, app_exe)
            return
        elif action == "cancel":
            self.after(0, self.destroy)
            return

        # 4) Descargar actualización
        self.set_status(
            f"Descargando v{remote_v}...",
            f"Desde: {download_url}",
            C["text"]
        )
        self.set_progress_indeterminate()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            zip_dest = tmp_path / "update.zip"

            ok = download_file(
                download_url,
                zip_dest,
                progress_callback=lambda pct, dl, tot: self._on_download_progress(pct, dl, tot)
            )

            if not ok:
                self.set_status(
                    "Error al descargar",
                    "Verifica tu conexión a internet.",
                    C["error"]
                )
                if app_exists:
                    self.after(3000, lambda: self._launch_and_close(base, app_exe))
                return

            # 5) Extraer directo a base_dir
            self.set_status("Instalando actualización...", "Extrayendo archivos finales...")
            self.set_progress_indeterminate()
            
            def _install_prog(pct):
                self.set_status(f"Instalando... {pct}%", "Descomprimiendo archivos...")
                self.set_progress(pct, mode="determinate")
            
            ok = extract_and_install(zip_dest, base, progress_callback=_install_prog)
            if not ok:
                self.set_status("Error al instalar", "Permisos insuficientes o disco lleno.", C["error"])
                return

        # 6) Guardar version instalada
        set_installed_version(base, remote_v)
        self.update_installed_label(remote_v)

        self.set_status(
            f"✓  Actualización a v{remote_v} completada",
            release_notes,
            C["success"]
        )
        self.set_progress(100)
        self.after(1500, lambda: self._launch_and_close(base, app_exe))

    def _on_download_progress(self, pct: int, downloaded: int, total: int):
        mb_dl = downloaded / 1_048_576
        mb_tot = total / 1_048_576
        self.set_status(
            f"Descargando...  {pct}%",
            f"{mb_dl:.1f} MB / {mb_tot:.1f} MB"
        )
        self.set_progress(pct)

    def _copy_update_files(self, src: Path, dest: Path):
        """Copia archivos de la actualización, ignorando user_config.json."""
        KEEP_LOCAL = {"user_config.json", LOCAL_VERSION_FILE}
        for item in src.rglob("*"):
            if item.is_file():
                rel = item.relative_to(src)
                # No sobreescribir datos del usuario
                if rel.name in KEEP_LOCAL:
                    continue
                target = dest / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(item, target)

    def _try_launch_or_error(self, base: Path):
        """Intenta lanzar la app, si no existe muestra error."""
        app_path = base / APP_EXE_NAME
        if app_path.exists():
            self.after(1500, lambda: self._launch_and_close(base, APP_EXE_NAME))
        else:
            self.set_status(
                "Error: App no encontrada",
                f"No se encontró {APP_EXE_NAME} en {base}",
                self.COLORS["error"]
            )

    def _launch_and_close(self, base: Path, exe_name: str):
        """Lanza la app y cierra el launcher."""
        app_path = base / exe_name
        if app_path.exists():
            try:
                subprocess.Popen(
                    [str(app_path)],
                    cwd=str(base),
                    creationflags=subprocess.DETACHED_PROCESS if sys.platform == "win32" else 0,
                )
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo lanzar la app:\n{e}")
                return
        self.destroy()

    def _ask_update(self, current: str, new: str, notes: str) -> str:
        """Pregunta al usuario si quiere actualizar.
        
        Usa una Queue para comunicación thread-safe entre el hilo worker
        y el hilo principal de tkinter.
        Retorna: 'update', 'skip' o 'cancel'.
        """
        q: queue.Queue = queue.Queue()

        def _show_dialog():
            """Se ejecuta en el hilo principal (via self.after)."""
            dlg = tk.Toplevel(self)
            dlg.title("Actualización disponible")
            dlg.resizable(False, False)
            dlg.configure(bg=self.COLORS["bg"])
            dlg.grab_set()
            dlg.focus_force()
            W2, H2 = 440, 230
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            dlg.geometry(f"{W2}x{H2}+{(sw-W2)//2}+{(sh-H2)//2}")

            # Si el usuario cierra la ventana con X → cancel
            dlg.protocol("WM_DELETE_WINDOW", lambda: _put_result("cancel"))

            def _put_result(choice: str):
                q.put(choice)
                dlg.destroy()

            inner = tk.Frame(dlg, bg=self.COLORS["bg"], padx=24, pady=20)
            inner.pack(fill="both", expand=True)

            tk.Label(
                inner, text="\U0001F680  Nueva versión disponible",
                font=("Segoe UI", 13, "bold"),
                bg=self.COLORS["bg"], fg=self.COLORS["text"]
            ).pack(anchor="w")

            tk.Label(
                inner, text=f"v{current}  →  v{new}",
                font=("Segoe UI", 11), bg=self.COLORS["bg"], fg=self.COLORS["accent2"]
            ).pack(anchor="w", pady=(4, 2))

            if notes:
                tk.Label(
                    inner, text=notes, font=("Segoe UI", 9),
                    bg=self.COLORS["bg"], fg=self.COLORS["text_dim"],
                    wraplength=390, justify="left"
                ).pack(anchor="w", pady=(0, 12))

            btn_frame = tk.Frame(inner, bg=self.COLORS["bg"])
            btn_frame.pack(anchor="e")

            tk.Button(
                btn_frame, text="  Actualizar  ",
                command=lambda: _put_result("update"),
                bg=self.COLORS["accent"], fg="white",
                font=("Segoe UI", 10, "bold"),
                relief="flat", padx=14, pady=6, cursor="hand2", bd=0,
            ).pack(side="right", padx=(8, 0))

            tk.Button(
                btn_frame, text="  Omitir  ",
                command=lambda: _put_result("skip"),
                bg=self.COLORS["card"], fg=self.COLORS["text_dim"],
                font=("Segoe UI", 10), relief="flat", padx=14, pady=6,
                cursor="hand2", bd=0,
            ).pack(side="right", padx=(8, 0))

            tk.Button(
                btn_frame, text="  Cancelar  ",
                command=lambda: _put_result("cancel"),
                bg=self.COLORS["card"], fg=self.COLORS["text_dim"],
                font=("Segoe UI", 10), relief="flat", padx=14, pady=6,
                cursor="hand2", bd=0,
            ).pack(side="right")

        # Programar la apertura del diálogo en el hilo principal
        self.after(0, _show_dialog)
        # El hilo worker espera la respuesta del usuario de forma bloqueante
        return q.get(block=True)


# ─────────────────────────────────────────────────────────────────────────────

def main():
    app = LauncherWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
