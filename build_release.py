#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_release.py — Script de release para GSEQuant PRO
=======================================================

USO:
    python build_release.py --version 4.1.0 --notes "Descripcion de cambios"

Qué hace:
1. Actualiza version.json con la nueva version
2. Compila GSEQuant_PRO con PyInstaller
3. Crea el ZIP de release (GSEQuant_PRO_v4.1.0.zip)
4. El ZIP contiene:
    - GSEQuant_PRO/  (la carpeta completa compilada)
    Listo para subir como asset a GitHub Releases

Luego manualmente:
    git add version.json
    git commit -m "Release v4.1.0"
    git push
    # Crear Release en GitHub y subir el ZIP como asset

Requisitos:
    pip install pyinstaller
"""

import argparse
import json
import os
import sys
import shutil
import subprocess
import zipfile
from datetime import date
from pathlib import Path


BASE = Path(__file__).parent
DIST_DIR = BASE / "dist"
BUILD_DIR = BASE / "build"
VERSION_FILE = BASE / "version.json"
APP_EXE_NAME = "GSEQuant_PRO.exe"
APP_FOLDER_NAME = "GSEQuant_PRO"

# URL base de GitHub Releases (cambia TU_USUARIO y NOMBRE_REPO)
GITHUB_USER = "logtechmoney"
GITHUB_REPO = "GSEQuant_PRO"
GITHUB_RELEASES_BASE = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest/download"


def parse_args():
    p = argparse.ArgumentParser(description="Build y release de GSEQuant PRO")
    p.add_argument("--version", required=True, help="Numero de version, ej: 4.1.0")
    p.add_argument("--notes", default="Nueva versión de GSEQuant PRO", help="Release notes (descripcion breve)")
    p.add_argument("--no-compile", action="store_true", help="Omitir compilacion (solo empaquetar)")
    return p.parse_args()


def update_version_json(version: str, notes: str):
    """Actualiza version.json con la nueva version."""
    zip_name = f"GSEQuant_PRO_v{version}.zip"
    data = {
        "version": version,
        "release_date": date.today().isoformat(),
        "download_url": f"{GITHUB_RELEASES_BASE}/{zip_name}",
        "release_notes": notes,
        "min_launcher_version": "1.0.0",
        "app_exe_name": APP_EXE_NAME,
    }
    with open(VERSION_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"✓ version.json actualizado → v{version}")
    return zip_name


def compile_app():
    """Compila GSEQuant PRO con PyInstaller."""
    print("→ Compilando GSEQuant PRO con PyInstaller...")
    spec_file = BASE / "GSEQuant_PRO.spec"
    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", str(spec_file), "--noconfirm"],
        cwd=str(BASE),
    )
    if result.returncode != 0:
        print("✗ Error en la compilacion")
        sys.exit(1)
    print("✓ Compilacion exitosa")


def create_release_zip(version: str, zip_name: str):
    """Crea el ZIP de release."""
    compiled_dir = DIST_DIR / APP_FOLDER_NAME
    if not compiled_dir.exists():
        print(f"✗ No se encontro el directorio compilado: {compiled_dir}")
        sys.exit(1)

    releases_dir = BASE / "releases"
    releases_dir.mkdir(exist_ok=True)
    zip_path = releases_dir / zip_name

    print(f"→ Creando ZIP: {zip_path}")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for item in compiled_dir.rglob("*"):
            if item.is_file():
                # No incluimos datos del usuario en el release
                if item.name in ("user_config.json", "installed_version.json"):
                    continue
                arcname = item.relative_to(compiled_dir)
                zf.write(item, arcname)
                print(f"   + {arcname}")

    size_mb = zip_path.stat().st_size / 1_048_576
    print(f"✓ ZIP creado: {zip_path.name}  ({size_mb:.1f} MB)")
    return zip_path


def print_instructions(version: str, zip_path: Path):
    """Indica los pasos a seguir para publicar el release."""
    print()
    print("=" * 60)
    print(f"  RELEASE v{version} LISTO — PASOS PARA PUBLICAR:")
    print("=" * 60)
    print()
    print("  1. Sube los cambios a GitHub:")
    print()
    print("     git add version.json")
    print(f'     git commit -m "Release v{version}"')
    print("     git push")
    print()
    print("  2. Crea un Release en GitHub:")
    print("     → Ve a: github.com/logtechmoney/GSEQuant_PRO/releases/new")
    print(f"     → Tag: v{version}")
    print(f"     → Sube el archivo: {zip_path.name}")
    print()
    print("  3. Los usuarios con el Launcher verán la actualización")
    print("     automáticamente la próxima vez que lo abran.")
    print()
    print(f"  ZIP: {zip_path}")
    print()


def main():
    args = parse_args()
    version = args.version.strip()

    print()
    print(f"  ╔══════════════════════════════════════╗")
    print(f"  ║  GSEQuant PRO — Build v{version:<13} ║")
    print(f"  ╚══════════════════════════════════════╝")
    print()

    # 1. Actualizar version.json
    zip_name = update_version_json(version, args.notes)

    # 2. Compilar (opcional)
    if not args.no_compile:
        compile_app()
    else:
        print("ℹ Compilacion omitida (--no-compile)")

    # 3. Crear ZIP de release
    zip_path = create_release_zip(version, zip_name)

    # 4. Instrucciones finales
    print_instructions(version, zip_path)


if __name__ == "__main__":
    main()
