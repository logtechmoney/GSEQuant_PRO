from __future__ import annotations
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GSEQuant – Map & Graph Builder (PyQt5 + Leaflet) — PRO v4 (stable)

Qué cambia vs tu versión que "no carga el mapa":
1) **Dependencia opcional**: `requests` ahora es opcional (si no está instalado, el app NO se cae; el buscador por texto solo no geocodifica).
2) **Leaflet robusto**: ya no usamos `marker.setTooltipContent(...)` (que no existe en algunos builds). Ahora: `marker.getTooltip()?.setContent(...)` o bind nuevamente.
3) **Click en marcadores**: detenemos la propagación con `L.DomEvent.stop(e.originalEvent)` (cuando existe) + `bubblingMouseEvents:false`.
4) **Funciones JS presentes**: todo lo que llama Python (add/update/remove node/edge, flyTo, draw/clear) está definido y expuesto en `window.*`.
5) **Rutas**: tabla muestra **Nombre real + (ID)** para `Desde`/`Hasta`. Botón y menú contextual para **eliminar ruta** (borra mapa + grafo + tabla).
6) **Sin nodos fantasma**: si estás uniendo dos nodos, los clics en el mapa no crean nodos nuevos.

Requisitos:
    pip install PyQt5 PyQtWebEngine networkx
    # (opcional) pip install requests

Uso:
    python GSEQuant_int.py
"""

import json
import math

import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

import os
import re
import xml.etree.ElementTree as ET
import datetime
from dataclasses import dataclass, asdict, field
from typing import Dict, List, Tuple, Optional, Set
from collections import defaultdict

# requests pasa a ser opcional
try:
    import requests  # type: ignore
except Exception:
    requests = None  # geocode queda deshabilitado si falta

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import pyqtSlot, QObject
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWebChannel import QWebChannel

import networkx as nx

# Lectura de Excel (opcional)
try:
    import pandas as pd  # type: ignore
except Exception:
    pd = None  # si falta, se pedirá instalar "pip install pandas openpyxl"

# ------------------------------ Utilidades ------------------------------ #

def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6371.0088
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlmb/2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def _dms_to_dd(deg: float, minu: float, sec: float, hemi: str) -> float:
    dd = abs(deg) + (abs(minu) / 60.0) + (abs(sec) / 3600.0)
    hemi = (hemi or '').strip().upper()
    if hemi in ('S', 'W'):
        dd = -dd
    return dd

def _parse_compact_dms(token: str) -> Optional[float]:
    """Parsea una coordenada en formato DMS compacto, p.ej. 314239S o 0604841W."""
    t = (token or '').strip().upper()
    m = re.match(r'^(\d{2,3})(\d{2})(\d{2}(?:\.\d+)?)([NSWE])$', t)
    if not m:
        return None
    deg = float(m.group(1))
    minu = float(m.group(2))
    sec = float(m.group(3))
    hemi = m.group(4)
    return _dms_to_dd(deg, minu, sec, hemi)

def parse_coords(text: str) -> Optional[Tuple[float, float]]:
    s = (text or '').strip().replace(',', ' ')
    s = re.sub(r'\s+', ' ', s)
    m = re.match(r'^([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)$', s)
    if m:
        lat, lon = float(m.group(1)), float(m.group(2))
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return lat, lon
    dms = re.findall(
        r'([NSWE])?\s*(\d{1,3})[°\s]\s*(\d{1,2})[\'\s]\s*(\d{1,2}(?:\.\d+)?)\s*("?)([NSWE])?',
        s, flags=re.IGNORECASE
    )
    if len(dms) >= 2:
        def as_dd(piece):
            hemi1, deg, minu, sec, _, hemi2 = piece
            hemi = (hemi1 or hemi2 or '').upper()
            return _dms_to_dd(float(deg), float(minu), float(sec), hemi)
        dd1 = as_dd(dms[0]); dd2 = as_dd(dms[1])
        lat, lon = dd1, dd2
        if abs(dd1) > 90 and abs(dd2) <= 90:
            lat, lon = dd2, dd1
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return lat, lon
    # Formato DMS compacto por token, p.ej. "314239S 0604841W"
    parts = s.split()
    if len(parts) == 2:
        dd1 = _parse_compact_dms(parts[0])
        dd2 = _parse_compact_dms(parts[1])
        if dd1 is not None and dd2 is not None:
            lat, lon = dd1, dd2
            if abs(dd1) > 90 and abs(dd2) <= 90:
                lat, lon = dd2, dd1
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                return lat, lon
    return None


@dataclass
class FlightTypeParams:
    code: str
    name: str
    coef_multiplier: float = 1.0


class Bridge(QObject):
    mapClicked = pyqtSlot(float, float)
    nodeClicked = pyqtSlot(str)
    nodeRightClicked = pyqtSlot(str)
    edgeClicked = pyqtSlot(str)

    @pyqtSlot(float, float)
    def onMapClick(self, lat, lon):
        self.mapClicked.emit(lat, lon)

    @pyqtSlot(str)
    def onNodeClick(self, nid):
        self.nodeClicked.emit(nid)

    @pyqtSlot(str)
    def onNodeRightClick(self, nid):
        self.nodeRightClicked.emit(nid)

    @pyqtSlot(str)
    def onEdgeClick(self, eid):
        self.edgeClicked.emit(eid)

def _utm_to_latlon(zone: int, northern: bool, easting: float, northing: float) -> Tuple[float, float]:
    """Convierte coordenadas UTM (WGS84) a lat/lon en grados decimales."""
    # Constantes WGS84
    a = 6378137.0
    f = 1 / 298.257223563
    k0 = 0.9996
    e2 = f * (2 - f)
    e4 = e2 * e2
    e6 = e4 * e2
    ep2 = e2 / (1 - e2)

    x = easting - 500000.0
    y = northing
    if not northern:
        # En el hemisferio sur UTM suma 10 000 000 m
        y -= 10000000.0

    m = y / k0
    mu = m / (a * (1 - e2 / 4 - 3 * e4 / 64 - 5 * e6 / 256))

    # Series de corrección
    e1 = (1 - math.sqrt(1 - e2)) / (1 + math.sqrt(1 - e2))
    j1 = 3 * e1 / 2 - 27 * e1**3 / 32
    j2 = 21 * e1**2 / 16 - 55 * e1**4 / 32
    j3 = 151 * e1**3 / 96
    j4 = 1097 * e1**4 / 512

    fp = (
        mu
        + j1 * math.sin(2 * mu)
        + j2 * math.sin(4 * mu)
        + j3 * math.sin(6 * mu)
        + j4 * math.sin(8 * mu)
    )

    sin_fp = math.sin(fp)
    cos_fp = math.cos(fp)
    tan_fp = math.tan(fp)

    c1 = ep2 * cos_fp**2
    t1 = tan_fp**2
    r1 = a * (1 - e2) / math.pow(1 - e2 * sin_fp**2, 1.5)
    n1 = a / math.sqrt(1 - e2 * sin_fp**2)
    d = x / (n1 * k0)

    # Latitud
    lat = (
        fp
        - (n1 * tan_fp / r1)
        * (
            d**2 / 2
            - (5 + 3 * t1 + 10 * c1 - 4 * c1**2 - 9 * ep2) * d**4 / 24
            + (61 + 90 * t1 + 298 * c1 + 45 * t1**2 - 252 * ep2 - 3 * c1**2) * d**6 / 720
        )
    )

    # Longitud
    lon0 = math.radians((zone - 1) * 6 - 180 + 3)  # meridiano central
    lon = (
        lon0
        + (
            d
            - (1 + 2 * t1 + c1) * d**3 / 6
            + (5 - 2 * c1 + 28 * t1 - 3 * c1**2 + 8 * ep2 + 24 * t1**2) * d**5 / 120
        )
        / cos_fp
    )

    return math.degrees(lat), math.degrees(lon)


def parse_utm_coords(text: str) -> Optional[Tuple[float, float]]:
    """Parsea un string UTM simple y devuelve (lat, lon) en grados decimales.

    Ejemplos válidos:
      - "21J 700000 6300000"
      - "21 J 700000 6300000"
    """
    if not text:
        return None
    s = re.sub(r"\s+", " ", text.strip().upper())
    parts = s.split(" ")
    zone = None
    band = None
    easting = None
    northing = None

    try:
        if len(parts) == 3:
            # "21J 700000 6300000"
            m = re.match(r"^(\d{1,2})([C-X])$", parts[0])
            if not m:
                return None
            zone = int(m.group(1))
            band = m.group(2)
            easting = float(parts[1])
            northing = float(parts[2])
        elif len(parts) == 4:
            # "21 J 700000 6300000"
            zone = int(parts[0])
            band = parts[1]
            easting = float(parts[2])
            northing = float(parts[3])
        else:
            return None
    except Exception:
        return None

    if zone is None or band is None or easting is None or northing is None:
        return None

    # Bandas C..M -> Sur, N..X -> Norte (aprox)
    band = band.upper()
    northern = band >= "N"

    try:
        return _utm_to_latlon(zone, northern, easting, northing)
    except Exception:
        return None


def coerce_bool(x):
    """Coerce various representations to a Python bool."""
    if isinstance(x, bool):
        return x
    if x is None:
        return False
    if isinstance(x, (int, float)):
        return bool(x)
    s = str(x).strip().lower()
    if s in ('1', 'true', 't', 'yes', 'y', 'si', 'sí'):
        return True
    if s in ('0', 'false', 'f', 'no', 'n', ''):
        return False
    return bool(s)

def geocode_text(query: str) -> Optional[Tuple[float, float, str]]:
    """Geocodifica usando Nominatim vía urllib (sin requerir 'requests')."""
    import urllib.request, urllib.parse, json as _json
    try:
        params = urllib.parse.urlencode({"q": query, "format": "json", "limit": 1})
        url = f"https://nominatim.openstreetmap.org/search?{params}"
        req = urllib.request.Request(url, headers={"User-Agent": "GSEQuant/1.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = _json.loads(resp.read().decode())
        if not data:
            return None
        item = data[0]
        return float(item['lat']), float(item['lon']), item.get('display_name', query)
    except Exception:
        return None


def reverse_geocode(lat: float, lon: float) -> Optional[str]:
    """Obtiene un nombre de lugar legible a partir de coordenadas usando Nominatim.

    Devuelve una cadena corta (pueblo / ciudad / aeropuerto, etc.) o None si falla.
    """
    if requests is None:
        return None
    try:
        params = {"lat": lat, "lon": lon, "format": "json", "zoom": 14, "addressdetails": 1}
        headers = {"User-Agent": "GSEQuant/1.0 (contacto: ops@example.com)"}
        r = requests.get("https://nominatim.openstreetmap.org/reverse", params=params, headers=headers, timeout=10)
        r.raise_for_status()
        data = r.json()
        if not data:
            return None
        addr = data.get("address", {}) or {}

        # 1) Priorizar ciudad / pueblo / municipio
        for key in ("city", "town", "village", "municipality"):
            val = addr.get(key)
            if val:
                # Opcionalmente agregar provincia/estado si está disponible
                state = addr.get("state") or addr.get("region")
                country = addr.get("country")
                parts = [val]
                if state:
                    parts.append(state)
                if country:
                    parts.append(country)
                return ", ".join(parts)

        # 2) Si no hay ciudad/pueblo, usar provincia/estado + país
        state = addr.get("state") or addr.get("region")
        country = addr.get("country")
        if state and country:
            return f"{state}, {country}"
        if state:
            return state
        if country:
            return country

        # 3) Fallback al display_name truncado
        name = data.get("display_name")
        if name:
            return name.split(",")[0]
        return None
    except Exception:
        return None

# ------------------------------ Modelo ------------------------------ #

class ConfigManager:
    def __init__(self) -> None:
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.defaults_dir = os.path.join(self.base_dir, "config_defaults")
        self.user_cfg_path = os.path.join(self.base_dir, "user_config.json")
        # Datasets que se administran internamente
        self.dataset_keys = [
            "EF",
            "gsexaeronaves",
            "coef_vehiculos",
            "circulacion",
            "tipos_vuelo"
        ]
        self.datasets: Dict[str, dict] = {}
        self._load_defaults()
        self._load_user_overrides()

    def _read_json(self, path: str) -> dict:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _standardize(self, data: dict) -> dict:
        # Asegura estructura {columns: [...], rows: [...]} incluso si vino con "sheets"
        if "columns" in data and "rows" in data:
            return data
        sheets = data.get("sheets", {})
        if len(sheets) == 1:
            k = next(iter(sheets))
            return {"columns": sheets[k].get("columns", []), "rows": sheets[k].get("rows", []), "sheet": k}
        # si hay múltiples hojas, tomar la primera
        for k, v in sheets.items():
            return {"columns": v.get("columns", []), "rows": v.get("rows", []), "sheet": k}
        return {"columns": [], "rows": []}

    def _normalize_gsexaeronaves(self, data: dict) -> dict:
        """Normaliza encabezados del dataset gsexaeronaves.

        Reemplaza columnas 'Unnamed:*' por nombres más descriptivos derivados del
        código de aeronave, manteniendo la estructura de grupos de 3 columnas
        (rampa, remota, tiempo) sin alterar la lógica de cálculo.
        """
        cols = list(data.get("columns", []) or [])
        rows = list(data.get("rows", []) or [])
        if not cols:
            return data

        new_cols = cols[:]
        rename_map: Dict[str, str] = {}
        i = 1
        while i + 2 < len(new_cols):
            ac = str(new_cols[i] or "").strip()
            if not ac:
                i += 3
                continue
            r_old = new_cols[i + 1]
            t_old = new_cols[i + 2]
            r_new = f"{ac}_S_Rampa"
            t_new = f"{ac}_t_hr"
            if isinstance(r_old, str) and r_old.startswith("Unnamed"):
                new_cols[i + 1] = r_new
                rename_map[r_old] = r_new
            if isinstance(t_old, str) and t_old.startswith("Unnamed"):
                new_cols[i + 2] = t_new
                rename_map[t_old] = t_new
            i += 3

        if not rename_map:
            return {"columns": new_cols, "rows": rows}

        new_rows = []
        for r in rows:
            nr = {}
            for k, v in (r or {}).items():
                nk = rename_map.get(k, k)
                nr[nk] = v
            new_rows.append(nr)
        return {"columns": new_cols, "rows": new_rows}

    def _normalize_coef_vehiculos(self, data: dict) -> dict:
        """Limpia filas auxiliares o vacías en coef_vehiculos.

        Elimina filas donde la columna 'GSE' está vacía o es NaN, de modo que
        sólo queden las filas de vehículos reales visibles en la tabla.
        """
        cols = list(data.get("columns", []) or [])
        rows = list(data.get("rows", []) or [])
        if not rows:
            return {"columns": cols, "rows": rows}

        def _is_empty_or_nan(x) -> bool:
            if x is None:
                return True
            # NaN numérico
            if isinstance(x, float) and math.isnan(x):
                return True
            # Cadenas vacías o marcadores tipo "nan", "NA", "null" guardados desde la UI
            s = str(x).strip()
            if not s:
                return True
            if s.lower() in ("nan", "na", "none", "null"):
                return True
            return False

        cleaned = []
        for r in rows:
            gse = r.get("GSE")
            if _is_empty_or_nan(gse):
                continue
            cleaned.append(r)
        return {"columns": cols, "rows": cleaned}

    def _load_defaults(self):
        # ... existing logic ...
        for key in self.dataset_keys:
            if key == "tipos_vuelo":
                self.datasets[key] = {
                    "columns": ["TipoVuelo", "Multiplicador"],
                    "rows": [
                        {"TipoVuelo": "INTERNACIONAL", "Multiplicador": 1.2},
                        {"TipoVuelo": "CABOTAJE", "Multiplicador": 1.0},
                        {"TipoVuelo": "REGIONAL", "Multiplicador": 1.1},
                        {"TipoVuelo": "GENERAL", "Multiplicador": 1.0}
                    ]
                }
                continue
            # ... resto de la lógica ...
        for key in self.dataset_keys:
            try:
                raw = self._read_json(os.path.join(self.defaults_dir, f"{key}.defaults.json"))
                std = self._standardize(raw)
                if key == "gsexaeronaves":
                    std = self._normalize_gsexaeronaves(std)
                if key == "coef_vehiculos":
                    std = self._normalize_coef_vehiculos(std)
                self.datasets[key] = std
            except Exception:
                self.datasets[key] = {"columns": [], "rows": []}

    def _load_user_overrides(self):
        try:
            data = self._read_json(self.user_cfg_path)
        except Exception:
            data = {}
        user_sets = data.get("datasets", {})
        for k, v in user_sets.items():
            if k in self.dataset_keys:
                std = self._standardize(v)
                if k == "gsexaeronaves":
                    std = self._normalize_gsexaeronaves(std)
                if k == "coef_vehiculos":
                    std = self._normalize_coef_vehiculos(std)
                self.datasets[k] = std

    def save_user_config(self):
        with open(self.user_cfg_path, 'w', encoding='utf-8') as f:
            json.dump({"datasets": self.datasets}, f, ensure_ascii=False, indent=2)

    def reset_to_defaults(self, key: Optional[str] = None):
        if key is None:
            self._load_defaults()
            try:
                if os.path.exists(self.user_cfg_path):
                    os.remove(self.user_cfg_path)
            except Exception:
                pass
        else:
            try:
                raw = self._read_json(os.path.join(self.defaults_dir, f"{key}.defaults.json"))
                std = self._standardize(raw)
                if key == "gsexaeronaves":
                    std = self._normalize_gsexaeronaves(std)
                if key == "coef_vehiculos":
                    std = self._normalize_coef_vehiculos(std)
                self.datasets[key] = std
            except Exception:
                self.datasets[key] = {"columns": [], "rows": []}
            self.save_user_config()

    # API pública
    def list_datasets(self) -> List[str]:
        return list(self.dataset_keys)

    def get_dataset(self, key: str) -> dict:
        return self.datasets.get(key, {"columns": [], "rows": []})

    def set_dataset_rows(self, key: str, rows: list):
        if key in self.datasets:
            self.datasets[key]["rows"] = rows
            # no guardamos automáticamente: se guarda desde el diálogo

    def set_dataset(self, key: str, columns: list, rows: list):
        if key not in self.datasets:
            self.datasets[key] = {"columns": [], "rows": []}
        self.datasets[key]["columns"] = columns
        self.datasets[key]["rows"] = rows

    @staticmethod
    def coerce_numeric(value):
        try:
            s = str(value).strip()
            if s == "":
                return None
            return float(s.replace(',', '.'))
        except Exception:
            return value

class HubAssignmentDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget, nodes: Dict[str, Node]):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Configurar hubs de circulación")
        self.resize(420, 220)
        self._nodes = nodes
        layout = QtWidgets.QVBoxLayout(self)

        info = QtWidgets.QLabel(
            "Seleccione qué nodos funcionarán como bases intermedias (hubs) por vehículo.\n"
            "Cada GSE puede tener un hub distinto (BUS, STA, BAG, BEL)."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        form = QtWidgets.QFormLayout()
        self.cmbBus = QtWidgets.QComboBox()
        self.cmbSta = QtWidgets.QComboBox()
        self.cmbBag = QtWidgets.QComboBox()
        self.cmbBel = QtWidgets.QComboBox()
        self._populate_combo(self.cmbBus, target="bus")
        self._populate_combo(self.cmbSta, target="sta")
        self._populate_combo(self.cmbBag, target="bag")
        self._populate_combo(self.cmbBel, target="bel")
        form.addRow("Hub BUS:", self.cmbBus)
        form.addRow("Hub STA:", self.cmbSta)
        form.addRow("Hub BAG:", self.cmbBag)
        form.addRow("Hub BEL:", self.cmbBel)
        layout.addLayout(form)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _populate_combo(self, combo: QtWidgets.QComboBox, target: str):
        combo.addItem("(sin hub)", "")
        current_id = None
        for node in self._nodes.values():
            combo.addItem(f"{node.name} ({node.id})", node.id)
            if (
                (target == "bus" and node.is_hub_bus)
                or (target == "sta" and node.is_hub_sta)
                or (target == "bag" and node.is_hub_bag)
                or (target == "bel" and node.is_hub_bel)
            ):
                current_id = node.id
        if current_id:
            idx = combo.findData(current_id)
            if idx >= 0:
                combo.setCurrentIndex(idx)

    def selected_hubs(self) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
        bus_id = self.cmbBus.currentData() or None
        sta_id = self.cmbSta.currentData() or None
        bag_id = self.cmbBag.currentData() or None
        bel_id = self.cmbBel.currentData() or None
        return bus_id, sta_id, bag_id, bel_id


class BaseAssignmentDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget, nodes: Dict[str, Node], veh_base_map: Optional[Dict[str, str]] = None):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Configurar bases por vehículo")
        self.resize(420, 260)
        self._nodes = nodes
        self._veh_base_map = veh_base_map or {}
        layout = QtWidgets.QVBoxLayout(self)

        info = QtWidgets.QLabel("Seleccione qué nodos funcionarán como base inicial por tipo de vehículo GSE.")
        info.setWordWrap(True)
        layout.addWidget(info)

        form = QtWidgets.QFormLayout()
        self._combos: Dict[str, QtWidgets.QComboBox] = {}
        vehs = ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]
        for code in vehs:
            combo = QtWidgets.QComboBox()
            self._populate_combo(combo, self._veh_base_map.get(code))
            form.addRow(f"Base {code}:", combo)
            self._combos[code] = combo
        layout.addLayout(form)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _populate_combo(self, combo: QtWidgets.QComboBox, current_id: Optional[str]):
        combo.addItem("(sin base específica)", "")
        for node in self._nodes.values():
            if getattr(node, "kind", "").lower() == "base":
                combo.addItem(f"{node.name} ({node.id})", node.id)
        if current_id:
            idx = combo.findData(current_id)
            if idx >= 0:
                combo.setCurrentIndex(idx)

    def selected_bases(self) -> Dict[str, str]:
        out: Dict[str, str] = {}
        for code, combo in self._combos.items():
            nid = combo.currentData() or None
            if nid:
                out[code] = nid
        return out


class TableEditorDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget, key: str, data: dict):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle(f"Editar datos · {key}")
        self.resize(900, 600)
        self.key = key
        # acceso al config manager desde la ventana principal
        self.config: ConfigManager = parent.config  # type: ignore

        layout = QtWidgets.QVBoxLayout(self)
        # info/top bar
        info = QtWidgets.QLabel("Edite valores celda por celda. Use Guardar para persistir en esta estación. 'Restaurar originales' vuelve a los defaults iniciales embebidos.")
        info.setWordWrap(True)
        layout.addWidget(info)

        # tabla
        self.table = QtWidgets.QTableWidget()
        # permitir redimensionar columnas/filas manualmente
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        layout.addWidget(self.table, stretch=1)

        # barra de acciones de tabla
        table_actions = QtWidgets.QHBoxLayout()
        self.btnAddRow = QtWidgets.QPushButton("Agregar fila")
        self.btnDelRow = QtWidgets.QPushButton("Eliminar fila")
        self.btnAddCol = QtWidgets.QPushButton("Agregar columna")
        self.btnDelCol = QtWidgets.QPushButton("Eliminar columna")
        self.btnRenCol = QtWidgets.QPushButton("Renombrar columna")
        self.btnAutoCols = QtWidgets.QPushButton("Auto ancho columnas")
        self.btnAutoRows = QtWidgets.QPushButton("Auto alto filas")
        table_actions.addWidget(self.btnAddRow)
        table_actions.addWidget(self.btnDelRow)
        table_actions.addSpacing(8)
        table_actions.addWidget(self.btnAddCol)
        table_actions.addWidget(self.btnDelCol)
        table_actions.addWidget(self.btnRenCol)
        table_actions.addStretch(1)
        table_actions.addWidget(self.btnAutoCols)
        table_actions.addWidget(self.btnAutoRows)
        layout.addLayout(table_actions)

        # botones inferiores
        btns = QtWidgets.QHBoxLayout()
        self.btnSave = QtWidgets.QPushButton("Guardar y cerrar")
        self.btnReset = QtWidgets.QPushButton("Restaurar originales")
        self.btnCancel = QtWidgets.QPushButton("Cancelar")
        btns.addStretch(1)
        btns.addWidget(self.btnReset)
        btns.addWidget(self.btnCancel)
        btns.addWidget(self.btnSave)
        layout.addLayout(btns)

        # poblar
        self._fill_table(data)

        # señales
        self.table.itemChanged.connect(self._on_item_changed)
        self.btnAddRow.clicked.connect(self._on_add_row)
        self.btnDelRow.clicked.connect(self._on_del_row)
        self.btnAddCol.clicked.connect(self._on_add_col)
        self.btnDelCol.clicked.connect(self._on_del_col)
        self.btnRenCol.clicked.connect(self._on_rename_col)
        self.btnAutoCols.clicked.connect(self.table.resizeColumnsToContents)
        self.btnAutoRows.clicked.connect(self.table.resizeRowsToContents)
        self.btnSave.clicked.connect(self._on_save)
        self.btnCancel.clicked.connect(self.reject)
        self.btnReset.clicked.connect(self._on_reset)

    def _fill_table(self, data: dict):
        cols = data.get("columns", [])
        rows = data.get("rows", [])
        self.table.blockSignals(True)
        self.table.clear()
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self._apply_header_tooltips(cols)
        self.table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            for j, col in enumerate(cols):
                val = row.get(col, None)
                item = QtWidgets.QTableWidgetItem("" if val is None else str(val))
                item.setToolTip(col)
                self.table.setItem(i, j, item)
        self.table.blockSignals(False)
        # por defecto dejamos modo interactivo; tamaño inicial al contenido
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

    def _apply_header_tooltips(self, cols: list) -> None:
        key = getattr(self, "key", "") or ""
        for j, name in enumerate(cols):
            item = self.table.horizontalHeaderItem(j)
            if item is None:
                continue
            tip = self._tooltip_for_column(key, str(name or ""))
            if tip:
                item.setToolTip(tip)

    def _tooltip_for_column(self, key: str, name: str) -> str:
        k = key.strip().lower()
        col = name or ""
        col_stripped = col.strip()
        if k == "coef_vehiculos":
            if col_stripped == "GSE":
                return "Nombre del vehículo GSE (texto identificador)."
            if col_stripped in ("PM", "Unnamed: 3"):
                return "Potencia del vehículo PM en caballos de fuerza [HP]."
            gases = {"CO2", "CO", "HC", "NOx", "SOx", "PM10", "PM.1"}
            if col_stripped in gases:
                return f"Factor de emisión base de {col_stripped} en [g/(HP·h)]."
            if col_stripped.startswith("FCD"):
                return "Coeficiente de operación FCD (modo descarga), adimensional."
            if col_stripped.startswith("FC"):
                return "Coeficiente de operación FC (modo carga), adimensional."
            if col_stripped.startswith("tD"):
                return "Tiempo de operación en modo descarga tD [h]."
            if col_stripped.startswith("t") and len(col_stripped) <= 3 and col_stripped[1:].isdigit():
                return "Tiempo de operación en modo carga t [h]."
            return "Campo de coeficientes del vehículo (sin unidades físicas específicas o adimensional)."
        if k == "ef":
            if col_stripped.lower() == "combustible":
                return "Tipo de combustible para el rango de potencia (texto)."
            if col_stripped.lower().startswith("hp min"):
                return "Límite inferior de potencia del rango [HP]."
            if col_stripped.lower().startswith("hp max"):
                return "Límite superior de potencia del rango [HP]."
            if col_stripped.endswith("DF A"):
                base = col_stripped.replace("DF A", "").strip()
                return f"Parámetro A del factor de deterioro para {base} (adimensional)."
            if col_stripped.endswith("DF B"):
                base = col_stripped.replace("DF B", "").strip()
                return f"Parámetro B del factor de deterioro para {base} (adimensional)."
            return "Parámetro de la tabla EF (factores de deterioro o configuración, normalmente adimensional)."
        if k == "gsexaeronaves":
            if col_stripped == "GSE":
                return "Código del vehículo GSE (GPU, CAT, TUG, BAG, BEL, WAT, BRE, LAV, FUE, STA, BUS, CLE)."
            if col_stripped.endswith("_t_hr"):
                ac = col_stripped[:-5]
                return f"Tiempo efectivo de servicio por operación y por unidad de GSE para aeronave {ac} [h]."
            if col_stripped.endswith("_S_Rampa"):
                ac = col_stripped.replace("_S_Rampa", "")
                return f"Cantidad de unidades de GSE por operación en puesto remoto (sin manga) para aeronave {ac} (adimensional)."
            return f"Cantidad de unidades de GSE por operación en puesto con manga para aeronave {col_stripped} (adimensional)."
        if k == "circulacion":
            if col_stripped == "Categoria":
                return "Tipo de registro: 'Nodo' o 'Ruta'."
            if col_stripped == "ID":
                return "Identificador único del nodo o ruta (texto)."
            if col_stripped == "Nombre":
                return "Nombre legible del nodo o ruta (texto)."
            if col_stripped == "Desde":
                return "Nodo de origen (ID) de la ruta."
            if col_stripped == "Hasta":
                return "Nodo de destino (ID) de la ruta."
            if col_stripped == "Dist_km":
                return "Longitud de la ruta en kilómetros [km]."
            if col_stripped == "Sentido":
                return "Sentido de circulación de la ruta ('Doble' o 'Solo ida')."
            if col_stripped in ("Es_hub_BUS", "Es_hub_STA", "Es_hub_BAG", "Es_hub_BEL"):
                return "Marca si el nodo actúa como hub para este tipo de vehículo GSE (sí/no)."
            if col_stripped == "Manga":
                return "Indica si el nodo de tipo 'puesto' tiene manga (sí/no)."
            if col_stripped == "Lat":
                return "Latitud del nodo en grados decimales [°]."
            if col_stripped == "Lon":
                return "Longitud del nodo en grados decimales [°]."
            return "Campo de la tabla de circulación (texto o valor adimensional)."
        return "Campo de dataset (texto o valor adimensional, sin unidades físicas específicas)."

    def _collect_rows(self) -> list:
        cols = [self.table.horizontalHeaderItem(j).text() for j in range(self.table.columnCount())]
        out = []
        for i in range(self.table.rowCount()):
            rec = {}
            for j, col in enumerate(cols):
                it = self.table.item(i, j)
                rec[col] = None if it is None else it.text()
            out.append(rec)
        return out

    def _collect_columns(self) -> list:
        return [self.table.horizontalHeaderItem(j).text() for j in range(self.table.columnCount())]

    def _on_item_changed(self, item: QtWidgets.QTableWidgetItem):
        # coerciones numéricas suaves
        itxt = item.text()
        coerced = ConfigManager.coerce_numeric(itxt)
        if isinstance(coerced, float) and (itxt != str(coerced)):
            self.table.blockSignals(True)
            item.setText(str(coerced))
            self.table.blockSignals(False)

    def _on_save(self):
        rows = self._collect_rows()
        cols = self._collect_columns()
        # preferir guardar columnas también
        if hasattr(self.config, "set_dataset"):
            self.config.set_dataset(self.key, cols, rows)  # type: ignore
        else:
            # fallback
            if self.key in self.config.datasets:
                self.config.datasets[self.key]["columns"] = cols
            self.config.set_dataset_rows(self.key, rows)
        self.config.save_user_config()
        self.accept()

    def _on_reset(self):
        # restaura solo este dataset y recarga en pantalla
        self.config.reset_to_defaults(self.key)
        data = self.config.get_dataset(self.key)
        self._fill_table(data)

    def _on_add_row(self):
        cols = self.table.columnCount()
        r = self.table.rowCount()
        self.table.insertRow(r)
        for j in range(cols):
            self.table.setItem(r, j, QtWidgets.QTableWidgetItem(""))

    def _on_del_row(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)

    def _on_add_col(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Nueva columna", "Nombre de la columna:")
        if not ok:
            return
        name = (name or "").strip()
        if not name:
            return
        # Validar duplicados
        existing = self._collect_columns()
        if name in existing:
            QtWidgets.QMessageBox.warning(self, "Columna", "Ya existe una columna con ese nombre.")
            return
        c = self.table.columnCount()
        self.table.insertColumn(c)
        self.table.setHorizontalHeaderItem(c, QtWidgets.QTableWidgetItem(name))
        for i in range(self.table.rowCount()):
            self.table.setItem(i, c, QtWidgets.QTableWidgetItem(""))
        self.table.resizeColumnsToContents()

    def _on_del_col(self):
        c = self.table.currentColumn()
        if c < 0:
            return
        name = self.table.horizontalHeaderItem(c).text() if self.table.horizontalHeaderItem(c) else ""
        if QtWidgets.QMessageBox.question(self, "Eliminar columna", f"¿Eliminar columna '{name}'?") != QtWidgets.QMessageBox.Yes:
            return
        self.table.removeColumn(c)
        self.table.resizeColumnsToContents()

    def _on_rename_col(self):
        c = self.table.currentColumn()
        if c < 0:
            return
        old = self.table.horizontalHeaderItem(c).text() if self.table.horizontalHeaderItem(c) else ""
        new, ok = QtWidgets.QInputDialog.getText(self, "Renombrar columna", "Nuevo nombre:", text=old)
        if not ok:
            return
        new = (new or "").strip()
        if not new or new == old:
            return
        existing = self._collect_columns()
        if new in existing:
            QtWidgets.QMessageBox.warning(self, "Columna", "Ya existe una columna con ese nombre.")
            return
        self.table.horizontalHeaderItem(c).setText(new)
        self.table.resizeColumnsToContents()

class EmissionsResultsDialog(QtWidgets.QDialog):
    def __init__(
        self,
        parent: QtWidgets.QWidget,
        resultados: Dict[str, Dict[str, float]],
        rates_gps: Optional[Dict[str, Dict[str, float]]] = None,
    ):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Emisiones en servicio – Resultados")
        self.resize(850, 600)
        self._resultados = resultados
        self._rates = rates_gps or {}

        layout = QtWidgets.QVBoxLayout(self)
        title = QtWidgets.QLabel("Resultados de emisiones por vehículo (servicio)")
        title.setWordWrap(True)
        layout.addWidget(title)

        gases = ["CO2","CO","HC","NOx","SOx","PM10"]
        # Orden canónico de presentación de GSE
        veh_order = ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]
        vehs = [v for v in veh_order if v in resultados] + [v for v in sorted(resultados.keys()) if v not in veh_order]

        # Tabla principal: resultados en gramos y g/s
        base_headers = ["Vehículo"]
        gas_headers = []
        for gas in gases:
            gas_headers.append(f"{gas} (g)")
            gas_headers.append(f"{gas} (g/s)")
        headers = base_headers + gas_headers

        self.table = QtWidgets.QTableWidget(len(vehs), len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        # Anchos iniciales razonables; luego el usuario puede ajustarlos.
        self.table.setColumnWidth(0, 130)
        # Restaurar anchos guardados de sesiones anteriores.
        self._restore_column_widths(self.table, "EmissionsResultsDialog/main")
        header.sectionResized.connect(lambda *_: self._save_column_widths(self.table, "EmissionsResultsDialog/main"))

        # Tooltips de encabezados (tabla principal)
        item_veh = self.table.horizontalHeaderItem(0)
        if item_veh is not None:
            item_veh.setToolTip("Vehículo / tipo de GSE.")
        col_tool = 1
        for gas in gases:
            item_g = self.table.horizontalHeaderItem(col_tool)
            if item_g is not None:
                item_g.setToolTip(
                    f"Emisiones de {gas} en servicio en gramos [g]."
                )
            item_rate = self.table.horizontalHeaderItem(col_tool + 1)
            if item_rate is not None:
                item_rate.setToolTip(
                    f"Tasa media de emisiones de {gas} en servicio en gramos por segundo [g/s]."
                )
            col_tool += 2

        for i, v in enumerate(vehs):
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(v))
            col = 1
            for g in gases:
                mass_val = resultados.get(v, {}).get(g, 0.0)
                it_mass = QtWidgets.QTableWidgetItem(f"{mass_val:.3f}")
                it_mass.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table.setItem(i, col, it_mass)

                rate_val = 0.0
                if v in self._rates:
                    rate_val = float(self._rates[v].get(g, 0.0) or 0.0)
                it_rate = QtWidgets.QTableWidgetItem(f"{rate_val:.4f}")
                it_rate.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table.setItem(i, col + 1, it_rate)

                col += 2
        layout.addWidget(self.table, stretch=1)

        # Tabla secundaria: mismos resultados convertidos a kg y toneladas
        base_headers_kg = ["Vehículo"]
        gas_headers_kg = []
        for gas in gases:
            gas_headers_kg.append(f"{gas} (kg)")
            gas_headers_kg.append(f"{gas} (Tn)")
        headers_kg = base_headers_kg + gas_headers_kg

        self.table_kg = QtWidgets.QTableWidget(len(vehs), len(headers_kg))
        self.table_kg.setHorizontalHeaderLabels(headers_kg)
        header_kg = self.table_kg.horizontalHeader()
        header_kg.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        self.table_kg.setColumnWidth(0, 130)
        self._restore_column_widths(self.table_kg, "EmissionsResultsDialog/kg")
        header_kg.sectionResized.connect(lambda *_: self._save_column_widths(self.table_kg, "EmissionsResultsDialog/kg"))

        # Tooltips de encabezados (tabla en kg/Tn)
        item_veh_kg = self.table_kg.horizontalHeaderItem(0)
        if item_veh_kg is not None:
            item_veh_kg.setToolTip("Vehículo / tipo de GSE.")
        col_tool_kg = 1
        for gas in gases:
            item_kg = self.table_kg.horizontalHeaderItem(col_tool_kg)
            if item_kg is not None:
                item_kg.setToolTip(
                    f"Emisiones de {gas} en servicio en kilogramos [kg]."
                )
            item_tn = self.table_kg.horizontalHeaderItem(col_tool_kg + 1)
            if item_tn is not None:
                item_tn.setToolTip(
                    f"Emisiones de {gas} en servicio en toneladas [Tn]."
                )
            col_tool_kg += 2

        for i, v in enumerate(vehs):
            self.table_kg.setItem(i, 0, QtWidgets.QTableWidgetItem(v))
            col = 1
            for g in gases:
                mass_val = float(resultados.get(v, {}).get(g, 0.0) or 0.0)
                kg_val = mass_val / 1000.0
                tn_val = kg_val / 1000.0
                it_kg = QtWidgets.QTableWidgetItem(f"{kg_val:.3f}")
                it_kg.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table_kg.setItem(i, col, it_kg)

                it_tn = QtWidgets.QTableWidgetItem(f"{tn_val:.4f}")
                it_tn.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table_kg.setItem(i, col + 1, it_tn)

                col += 2
        layout.addWidget(self.table_kg, stretch=1)

        btns = QtWidgets.QHBoxLayout()
        self.btnExport = QtWidgets.QPushButton("Exportar CSV…")
        self.btnClose = QtWidgets.QPushButton("Cerrar")
        btns.addStretch(1)
        btns.addWidget(self.btnExport)
        btns.addWidget(self.btnClose)
        layout.addLayout(btns)

        self.btnExport.clicked.connect(self._export_csv)
        self.btnClose.clicked.connect(self.accept)

    def _save_column_widths(self, table: QtWidgets.QTableWidget, key: str) -> None:
        """Guarda los anchos de columnas de una tabla en QSettings bajo una clave dada."""
        try:
            settings = QtCore.QSettings("GSEQuant", "GSEQuant")
            widths = [str(table.columnWidth(i)) for i in range(table.columnCount())]
            settings.setValue(f"{key}/widths", ",".join(widths))
        except Exception:
            # No bloquear el diálogo por errores de persistencia visual
            pass

    def _restore_column_widths(self, table: QtWidgets.QTableWidget, key: str) -> None:
        """Restaura los anchos de columnas de una tabla desde QSettings si existen."""
        try:
            settings = QtCore.QSettings("GSEQuant", "GSEQuant")
            value = settings.value(f"{key}/widths")
            if not value:
                return
            if isinstance(value, (list, tuple)):
                parts = [str(v) for v in value]
            else:
                parts = str(value).split(",")
            for i, w in enumerate(parts):
                if i >= table.columnCount():
                    break
                try:
                    table.setColumnWidth(i, int(w))
                except Exception:
                    continue
        except Exception:
            pass

    def _export_csv(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Exportar emisiones", os.getcwd(), "CSV (*.csv)")
        if not fn:
            return
        gases = ["CO2","CO","HC","NOx","SOx","PM10"]
        try:
            with open(fn, 'w', encoding='utf-8') as f:
                headers = ["vehiculo"]
                for g in gases:
                    headers.append(f"{g}_g")
                    headers.append(f"{g}_gps")
                f.write(",".join(headers) + "\n")
                for v, rec in self._resultados.items():
                    rate_rec = self._rates.get(v, {}) if hasattr(self, "_rates") else {}
                    row = [v]
                    for g in gases:
                        mass_val = float(rec.get(g, 0.0) or 0.0)
                        rate_val = float(rate_rec.get(g, 0.0) or 0.0)
                        row.append(f"{mass_val:.3f}")
                        row.append(f"{rate_val:.4f}")
                    f.write(",".join(row) + "\n")
            QtWidgets.QMessageBox.information(self, "Exportación", f"Emisiones exportadas a\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

class TotalEmissionsDialog(QtWidgets.QDialog):
    def __init__(
        self,
        parent: QtWidgets.QWidget,
        data: Dict[str, Dict[str, Dict[str, float]]],
        gases: List[str],
        min_fleet: Optional[Dict[str, int]] = None,
    ):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Emisiones totales – Servicio + Circulación")
        self.resize(950, 600)
        self._data = data
        self._gases = gases
        self._min_fleet = min_fleet or {}

        layout = QtWidgets.QVBoxLayout(self)
        title = QtWidgets.QLabel("Resultados combinados por vehículo (servicio + circulación)")
        title.setWordWrap(True)
        layout.addWidget(title)

        # Orden canónico de presentación de GSE
        veh_order = ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]
        vehs = [v for v in veh_order if v in data] + [v for v in sorted(data.keys()) if v not in veh_order]

        base_headers = ["Vehículo", "FMT"]
        gas_headers = []
        for gas in gases:
            gas_headers.append(f"{gas} total (g)")
            gas_headers.append(f"{gas} total (g/s)")
        headers = base_headers + gas_headers

        self.table = QtWidgets.QTableWidget(len(vehs), len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        # Anchos iniciales razonables; luego el usuario puede ajustarlos.
        self.table.setColumnWidth(0, 130)
        self.table.setColumnWidth(1, 60)
        # Restaurar anchos guardados de sesiones anteriores.
        self._restore_column_widths(self.table, "TotalEmissionsDialog/main")
        header.sectionResized.connect(lambda *_: self._save_column_widths(self.table, "TotalEmissionsDialog/main"))

        # Tooltips de encabezados (tabla principal)
        item_veh = self.table.horizontalHeaderItem(0)
        if item_veh is not None:
            item_veh.setToolTip("Vehículo / tipo de GSE.")
        item_fmt = self.table.horizontalHeaderItem(1)
        if item_fmt is not None:
            item_fmt.setToolTip(
                "FMT: Flota mínima total necesaria para cubrir servicio y circulación,\n"
                "considerando tiempos de servicio y tiempos de circulación."
            )
        col_tool = 2
        for gas in gases:
            item_g = self.table.horizontalHeaderItem(col_tool)
            if item_g is not None:
                item_g.setToolTip(
                    f"Emisiones totales de {gas} en gramos [g] (servicio + circulación)."
                )
            item_rate = self.table.horizontalHeaderItem(col_tool + 1)
            if item_rate is not None:
                item_rate.setToolTip(
                    f"Tasa media de emisiones totales de {gas} en gramos por segundo [g/s],\n"
                    "considerando tiempos de servicio + tiempos de circulación."
                )
            col_tool += 2

        for i, v in enumerate(vehs):
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(v))
            fleet_val = 0
            try:
                fleet_val = int(self._min_fleet.get(v, 0) or 0)
            except Exception:
                fleet_val = 0
            it_fleet = QtWidgets.QTableWidgetItem("—" if fleet_val <= 0 else str(fleet_val))
            it_fleet.setTextAlignment(QtCore.Qt.AlignCenter)
            self.table.setItem(i, 1, it_fleet)
            col = 2
            rec_gases = data.get(v, {}) or {}
            for gas in gases:
                rec = rec_gases.get(gas, {}) or {}
                total_g = float(rec.get("total", 0.0) or 0.0)
                total_gps = float(rec.get("total_gps", 0.0) or 0.0)
                it_g = QtWidgets.QTableWidgetItem(f"{total_g:.3f}")
                it_g.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table.setItem(i, col, it_g)
                it_rate = QtWidgets.QTableWidgetItem(f"{total_gps:.4f}")
                it_rate.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table.setItem(i, col + 1, it_rate)
                col += 2
        layout.addWidget(self.table, stretch=1)

        base_headers_kg = ["Vehículo", "FMT"]
        gas_headers_kg = []
        for gas in gases:
            gas_headers_kg.append(f"{gas} total (kg)")
            gas_headers_kg.append(f"{gas} total (Tn)")
        headers_kg = base_headers_kg + gas_headers_kg

        self.table_kg = QtWidgets.QTableWidget(len(vehs), len(headers_kg))
        self.table_kg.setHorizontalHeaderLabels(headers_kg)
        header_kg = self.table_kg.horizontalHeader()
        header_kg.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        self.table_kg.setColumnWidth(0, 130)
        self.table_kg.setColumnWidth(1, 60)
        self._restore_column_widths(self.table_kg, "TotalEmissionsDialog/kg")
        header_kg.sectionResized.connect(lambda *_: self._save_column_widths(self.table_kg, "TotalEmissionsDialog/kg"))

        # Tooltips de encabezados (tabla en kg/Tn)
        item_veh_kg = self.table_kg.horizontalHeaderItem(0)
        if item_veh_kg is not None:
            item_veh_kg.setToolTip("Vehículo / tipo de GSE.")
        item_fmt_kg = self.table_kg.horizontalHeaderItem(1)
        if item_fmt_kg is not None:
            item_fmt_kg.setToolTip(
                "FMT: Flota mínima total necesaria para cubrir servicio y circulación,\n"
                "considerando tiempos de servicio y tiempos de circulación."
            )
        col_tool_kg = 2
        for gas in gases:
            item_kg = self.table_kg.horizontalHeaderItem(col_tool_kg)
            if item_kg is not None:
                item_kg.setToolTip(
                    f"Emisiones totales de {gas} en kilogramos [kg] (servicio + circulación)."
                )
            item_tn = self.table_kg.horizontalHeaderItem(col_tool_kg + 1)
            if item_tn is not None:
                item_tn.setToolTip(
                    f"Emisiones totales de {gas} en toneladas [Tn] (servicio + circulación)."
                )
            col_tool_kg += 2

        for i, v in enumerate(vehs):
            self.table_kg.setItem(i, 0, QtWidgets.QTableWidgetItem(v))
            fleet_val = 0
            try:
                fleet_val = int(self._min_fleet.get(v, 0) or 0)
            except Exception:
                fleet_val = 0
            it_fleet = QtWidgets.QTableWidgetItem("—" if fleet_val <= 0 else str(fleet_val))
            it_fleet.setTextAlignment(QtCore.Qt.AlignCenter)
            self.table_kg.setItem(i, 1, it_fleet)
            col = 2
            rec_gases = data.get(v, {}) or {}
            for gas in gases:
                rec = rec_gases.get(gas, {}) or {}
                total_g = float(rec.get("total", 0.0) or 0.0)
                kg_val = total_g / 1000.0
                tn_val = kg_val / 1000.0
                it_kg = QtWidgets.QTableWidgetItem(f"{kg_val:.3f}")
                it_kg.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table_kg.setItem(i, col, it_kg)
                it_tn = QtWidgets.QTableWidgetItem(f"{tn_val:.4f}")
                it_tn.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table_kg.setItem(i, col + 1, it_tn)
                col += 2
        layout.addWidget(self.table_kg, stretch=1)

        btns = QtWidgets.QHBoxLayout()
        self.btnExport = QtWidgets.QPushButton("Exportar CSV…")
        self.btnReport = QtWidgets.QPushButton("Informe completo…")
        self.btnClose = QtWidgets.QPushButton("Cerrar")
        btns.addStretch(1)
        btns.addWidget(self.btnExport)
        btns.addWidget(self.btnReport)
        btns.addWidget(self.btnClose)
        layout.addLayout(btns)

        self.btnExport.clicked.connect(self._export_csv)
        self.btnReport.clicked.connect(self._on_full_report)
        self.btnClose.clicked.connect(self.accept)

    def _save_column_widths(self, table: QtWidgets.QTableWidget, key: str) -> None:
        """Guarda los anchos de columnas de una tabla en QSettings bajo una clave dada."""
        try:
            settings = QtCore.QSettings("GSEQuant", "GSEQuant")
            widths = [str(table.columnWidth(i)) for i in range(table.columnCount())]
            settings.setValue(f"{key}/widths", ",".join(widths))
        except Exception:
            # No bloquear el diálogo por errores de persistencia visual
            pass

    def _restore_column_widths(self, table: QtWidgets.QTableWidget, key: str) -> None:
        """Restaura los anchos de columnas de una tabla desde QSettings si existen."""
        try:
            settings = QtCore.QSettings("GSEQuant", "GSEQuant")
            value = settings.value(f"{key}/widths")
            if not value:
                return
            if isinstance(value, (list, tuple)):
                parts = [str(v) for v in value]
            else:
                parts = str(value).split(",")
            for i, w in enumerate(parts):
                if i >= table.columnCount():
                    break
                try:
                    table.setColumnWidth(i, int(w))
                except Exception:
                    continue
        except Exception:
            pass

    def _export_csv(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Exportar emisiones totales", os.getcwd(), "CSV (*.csv)")
        if not fn:
            return
        try:
            with open(fn, "w", encoding="utf-8") as f:
                headers = ["vehiculo", "dist_km", "circulation_time_h", "service_time_h", "flota_min_total"]
                for g in self._gases:
                    headers.append(f"{g}_total_g")
                    headers.append(f"{g}_total_gps")
                f.write(",".join(headers) + "\n")
                for v in sorted(self._data.keys()):
                    rec_gases = self._data.get(v, {}) or {}
                    try:
                        res_obj = self.parent()._data_circ[v]
                        dist = f"{res_obj.distance_km:.3f}"
                        circ = f"{res_obj.circulation_time_h:.3f}"
                        serv = f"{res_obj.service_time_h:.3f}"
                    except Exception:
                        dist, circ, serv = "0.000", "0.000", "0.000"

                    fleet_val = 0
                    try:
                        fleet_val = int(self._min_fleet.get(v, 0) or 0)
                    except Exception:
                        fleet_val = 0
                    row = [v, dist, circ, serv, str(fleet_val)]
                    for g in self._gases:
                        rec = rec_gases.get(g, {}) or {}
                        total_g = float(rec.get("total", 0.0) or 0.0)
                        total_gps = float(rec.get("total_gps", 0.0) or 0.0)
                        row.append(f"{total_g:.3f}")
                        row.append(f"{total_gps:.4f}")
                    f.write(",".join(row) + "\n")
            QtWidgets.QMessageBox.information(self, "Exportación", f"Emisiones exportadas a\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    def _on_full_report(self):
        try:
            parent = self.parent()
            path = None
            if parent is not None and hasattr(parent, "generate_full_emissions_report"):
                path = parent.generate_full_emissions_report()
            if not path:
                QtWidgets.QMessageBox.information(
                    self,
                    "Informe",
                    "No hay datos suficientes para generar el informe completo.\nEjecute primero el cálculo de emisiones totales.",
                )
                return
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Informe", str(e))


class DateRangeDialog(QtWidgets.QDialog):
    def __init__(
        self,
        parent: QtWidgets.QWidget,
        min_date: Optional[datetime.date] = None,
        max_date: Optional[datetime.date] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Seleccionar rango de fechas")
        self.resize(360, 180)

        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()

        today = QtCore.QDate.currentDate()
        self.date_from = QtWidgets.QDateEdit(today)
        self.date_from.setCalendarPopup(True)
        self.date_to = QtWidgets.QDateEdit(today)
        self.date_to.setCalendarPopup(True)

        if min_date is not None:
            qmin = QtCore.QDate(min_date.year, min_date.month, min_date.day)
            self.date_from.setMinimumDate(qmin)
            self.date_to.setMinimumDate(qmin)
        if max_date is not None:
            qmax = QtCore.QDate(max_date.year, max_date.month, max_date.day)
            self.date_from.setMaximumDate(qmax)
            self.date_to.setMaximumDate(qmax)

        if min_date is not None:
            self.date_from.setDate(QtCore.QDate(min_date.year, min_date.month, min_date.day))
        if max_date is not None:
            self.date_to.setDate(QtCore.QDate(max_date.year, max_date.month, max_date.day))

        form.addRow("Desde:", self.date_from)
        form.addRow("Hasta:", self.date_to)
        layout.addLayout(form)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_selected_range(self) -> Tuple[datetime.date, datetime.date]:
        return self.date_from.date().toPyDate(), self.date_to.date().toPyDate()


class EmissionsCalculator:
    def __init__(self, config: ConfigManager, age_years: float = 8.0, t_util_years: float = 10.0, combustible: str = "Diesel") -> None:
        self.config = config
        self.age = age_years
        self.t_util = t_util_years
        self.combustible = combustible
        self._ops_df = None  # DataFrame de operaciones (opcional)
        self._overrides = None  # parámetros de simulación (dict)
        self._veh_map = {
            "GPU": "Ground Power Units",
            "CAT": "Catering",
            "TUG": "Tugs/ Aircraft tractor",
            "BAG": "Baggage",
            "BEL": "belt loader",
            "WAT": "Water truck",
            "BRE": "Camion pasajeros mov reducida",
            "LAV": "Lavatory truck",
            "FUE": "Fuel truck",
            "STA": "Passenger Stands",
            "BUS": "Transporte de pasajeros",
            "CLE": "Limpieza/servicios",
        }
        self._gases = ["CO2", "CO", "HC", "NOx", "SOx", "PM10"]

    def set_operations_df(self, df):
        # Guardar sin validar; la UI controla la carga
        self._ops_df = df

    def set_overrides(self, overrides: dict):
        # overrides: {
        #   'default': {'age':..., 't_util':..., 'combustible':..., 'ef_hp_max':..., 'hp_vehicle':...},
        #   'veh': { 'GPU': {...}, 'CAT': {...}, ...}
        # }
        self._overrides = overrides or {}

    def _ovr(self, veh: str, key: str, fallback=None):
        vmap = (self._overrides or {}).get('veh', {})
        dflt = (self._overrides or {}).get('default', {})
        if veh in vmap and key in vmap[veh] and vmap[veh][key] not in (None, ""):
            return vmap[veh][key]
        if key in dflt and dflt[key] not in (None, ""):
            return dflt[key]
        return fallback

    @staticmethod
    def _to_float(x):
        try:
            if x is None:
                return None
            s = str(x)
            if s.lower() == 'nan':
                return None
            return float(str(x).replace(',', '.'))
        except Exception:
            return None

    def _load_coef_vehiculos(self) -> List[dict]:
        data = self.config.get_dataset("coef_vehiculos")
        cols = data.get("columns", [])
        rows = data.get("rows", [])
        out = []
        for r in rows:
            rec = {}
            for c in cols:
                rec[c] = r.get(c)
            out.append(rec)
        return [r for r in out if (r.get("GSE") is not None and str(r.get("GSE")).strip() != '')]

    def _load_EF(self) -> List[dict]:
        data = self.config.get_dataset("EF")
        cols = data.get("columns", [])
        rows = data.get("rows", [])
        out = []
        for r in rows:
            rec = {}
            for c in cols:
                rec[c] = r.get(c)
            out.append(rec)
        return out

    def _find_vehicle_row(self, veh_code: str, coef_rows: List[dict]) -> Optional[dict]:
        gse_name = self._veh_map.get(veh_code, veh_code)
        for r in coef_rows:
            if str(r.get("GSE", "")).strip().lower() == str(gse_name).strip().lower():
                return r
        return None

    def _select_EF_row_for_veh(self, veh: str, hp_value: Optional[float], ef_rows: List[dict]) -> Optional[dict]:
        comb = (self._ovr(veh, 'combustible', self.combustible) or self.combustible).strip().lower()
        # Para seleccionar la fila EF, solo usamos ef_hp_max si es un override
        # específico por vehículo. El valor por defecto no debe forzar un HP max
        # fijo, para que abrir el diálogo de parámetros sin cambiar nada no
        # altere los resultados.
        ov = self._overrides or {}
        veh_map = ov.get('veh', {}) or {}
        veh_rec = veh_map.get(veh, {}) or {}
        target_hp_max = veh_rec.get('ef_hp_max', None)
        candidates = [r for r in ef_rows if str(r.get("Combustible", "")).strip().lower() == comb]
        if not candidates:
            return None
        if target_hp_max is not None:
            # Selección exacta por HP max
            for r in candidates:
                try:
                    hp_mx = float(str(r.get("HP max")).replace(',', '.'))
                except Exception:
                    continue
                if abs(hp_mx - float(target_hp_max)) < 1e-6:
                    return r
        if hp_value is None:
            return candidates[0]
        # Rango por HP del vehículo
        for r in candidates:
            try:
                hp_min = float(str(r.get("HP min")).replace(',', '.'))
                hp_max = float(str(r.get("HP max")).replace(',', '.'))
            except Exception:
                continue
            if hp_min <= hp_value <= hp_max:
                return r
        return candidates[0]

    def _vehicle_hp(self, vrow: dict) -> Optional[float]:
        pm = self._to_float(vrow.get("PM"))
        if pm is not None:
            return pm
        return self._to_float(vrow.get("Unnamed: 3"))

    def _k_for(self, vrow: dict, efrow: dict, gas: str, veh: str, age_override=None, t_util_override=None, hp_override=None) -> Optional[float]:
        base_col = gas if gas != 'PM10' else 'PM.1'
        base = self._to_float(vrow.get(base_col))
        if base is None:
            return None
        hp = (self._to_float(hp_override) if hp_override not in (None, "") else self._vehicle_hp(vrow)) or 0.0
        if gas == 'SOx':
            df_a = 0.0; df_b = 1.0
        else:
            a_col = f"{gas} DF A" if gas != 'PM10' else "PM10 DF A"
            b_col = f"{gas} DF B" if gas != 'PM10' else "PM10 DF B"
            df_a = self._to_float(efrow.get(a_col)) or 0.0
            df_b = self._to_float(efrow.get(b_col)) or 1.0
        age = self._to_float(age_override) if age_override not in (None, "") else self.age
        t_util = self._to_float(t_util_override) if t_util_override not in (None, "") else self.t_util
        # Verificación estricta: todo equipo tiene una vida útil (no existe entropía nula).
        # Si la vida útil no está configurada o es cero, es un error de carga de datos.
        if t_util is None or t_util <= 0.0:
            raise ValueError(
                f"La vida útil de la tecnología (t_util) para el vehículo '{veh}' "
                f"tiene un valor inválido ({t_util}). Todos los equipos deben tener un "
                "tiempo de utilidad establecido en los parámetros para calcular su desgaste. "
                "Por favor verifique la configuración."
            )
            
        deterioro = df_a * (age / t_util) ** df_b
            
        k = hp * base * (1.0 + deterioro)
        return k

    def compute_emisiones_servicio(self) -> Dict[str, Dict[str, float]]:
        coef_rows = self._load_coef_vehiculos()
        ef_rows = self._load_EF()
        resultados: Dict[str, Dict[str, float]] = {}
        vehs = ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]
        for veh in vehs:
            vrow = self._find_vehicle_row(veh, coef_rows)
            if not vrow:
                continue
            hp_vehicle = self._vehicle_hp(vrow)
            efrow = self._select_EF_row_for_veh(veh, hp_vehicle, ef_rows)
            if not efrow:
                continue
            FCD = [self._to_float(vrow.get("FCD 1 ")), self._to_float(vrow.get("FCD 2")), self._to_float(vrow.get("FCD 3")), self._to_float(vrow.get("FCD 4"))]
            FC  = [self._to_float(vrow.get("FC 1 ")),  self._to_float(vrow.get("FC 2")),  self._to_float(vrow.get("FC 3")),  self._to_float(vrow.get("FC 4"))]
            tD  = [self._to_float(vrow.get("tD1")),    self._to_float(vrow.get("tD2")),    self._to_float(vrow.get("tD3")),    self._to_float(vrow.get("tD4"))]
            tC  = [self._to_float(vrow.get("t1")),     self._to_float(vrow.get("t2")),     self._to_float(vrow.get("t3")),     self._to_float(vrow.get("t4"))]
            FCD = [x or 0.0 for x in FCD]; FC = [x or 0.0 for x in FC]; tD = [x or 0.0 for x in tD]; tC = [x or 0.0 for x in tC]
            res_gases: Dict[str, float] = {}
            age_o = self._ovr(veh, 'age', None)
            tutil_o = self._ovr(veh, 't_util', None)
            hp_o = self._ovr(veh, 'hp_vehicle', None)
            for gas in self._gases:
                k = self._k_for(vrow, efrow, gas, veh, age_override=age_o, t_util_override=tutil_o, hp_override=hp_o)
                if k is None:
                    continue
                e_c = sum(k * FC[i] * tC[i] for i in range(4))
                e_d = sum(k * FCD[i] * tD[i] for i in range(4))
                total = e_c + e_d
                # Debug paso a paso solo para GPU · CO2
                if veh == 'GPU' and gas == 'HC':
                    try:
                        base = self._to_float(vrow.get('HC'))
                        hp_used = (self._to_float(hp_o) if hp_o not in (None, "") else self._vehicle_hp(vrow)) or 0.0
                        age_used = self._to_float(age_o) if age_o not in (None, "") else self.age
                        tutil_used = self._to_float(tutil_o) if tutil_o not in (None, "") else self.t_util
                        # DFs como en _k_for para CO2 (toma DF A/B de EF si existen)
                        a_col = "HC DF A"; b_col = "HC DF B"
                        df_a = self._to_float(efrow.get(a_col)) or 0.0
                        df_b = self._to_float(efrow.get(b_col)) or 1.0
                        df_val = 1.0 + df_a * (age_used / tutil_used) ** df_b
                        k_manual = hp_used * (base or 0.0) * df_val
                        comp_c = [k * (FC[i] or 0.0) * (tC[i] or 0.0) for i in range(4)]
                        comp_d = [k * (FCD[i] or 0.0) * (tD[i] or 0.0) for i in range(4)]
                        print("==== Debug · GPU · HC ====")
                        print(f"Combustible (EF): {(self._ovr(veh, 'combustible', self.combustible) or self.combustible)}")
                        print(f"EF HP min/max: [{efrow.get('HP min')}, {efrow.get('HP max')}]  (seleccionada por potencia/rango)")
                        print(f"HP (usado) = {hp_used}")
                        print(f"base_HC = {base}")
                        print(f"age = {age_used}")
                        print(f"t_util = {tutil_used}")
                        print(f"HC DF_A = {df_a}")
                        print(f"HC DF_B = {df_b}")
                        print(f"DF = 1 + A*(age/t_util)^B = {df_val}")
                        print(f"K = HP * base_HC * DF = {hp_used} * {base} * {df_val} = {k_manual}")
                        print(f"K (en uso) = {k}")
                        print(f"FC  = {FC}")
                        print(f"tC  = {tC}")
                        print(f"FCD = {FCD}")
                        print(f"tD  = {tD}")
                        print(f"Componentes carga (K*FC[i]*tC[i]) = {comp_c}")
                        print(f"E_c = {e_c}")
                        print(f"Componentes descarga (K*FCD[i]*tD[i]) = {comp_d}")
                        print(f"E_d = {e_d}")
                        print(f"E_total = E_c + E_d = {e_c} + {e_d} = {total}")
                        print("===========================")
                    except Exception:
                        pass
                    print(">>>> DEBUG GPU HP <<<<")
                    print("hp_override (hp_o) =", hp_o)
                    print("PM =", vrow.get("PM"))
                    print("Unnamed: 3 =", vrow.get("Unnamed: 3"))
                    print("hp_used (hp_used) =", hp_used)
                    print(">>>>>>>>>>>>>>>>>>>>>>")
                res_gases[gas] = total
            resultados[veh] = res_gases
        return resultados


@dataclass
class OperationRecord:
    index: int
    arr: float
    dep: float
    stand_id: str
    stand_name: str
    has_jetbridge: bool
    aircraft: str
    tipo_ser: str
    date: Optional[datetime.date] = None


@dataclass
class CirculationResult:
    distance_km: float
    circulation_time_h: float
    service_time_h: float
    fleet: int
    warnings: List[str]
    gases: Dict[str, Dict[str, float]]  # gas -> {"g": value, "gps": value}
    # Desglose por puesto: {stand_name -> {"dist_km": float, "svc_h": float, gases...}}
    by_stand: Dict[str, Dict] = None  # type: ignore

    def __post_init__(self):
        if self.by_stand is None:
            self.by_stand = {}


class CirculationCalculator:
    """Simula circulación de GSE para estimar distancias, tiempos y emisiones."""

    _gases = ["CO2", "CO", "HC", "NOx", "SOx", "PM10"]

    def __init__(self, model: GraphModel, config: ConfigManager,
                 circ_params: Optional[Dict[str, Dict[str, float]]] = None,
                 sim_params: Optional[Dict] = None,
                 dataset: Optional[dict] = None,
                 debug_enabled: bool = False) -> None:
        self.model = model
        self.config = config
        self._circ_params = circ_params or {}
        self._sim_params = sim_params or {}
        self._ops: List[OperationRecord] = []
        self._gse_matrix: Dict[str, Dict[str, Dict[str, float]]] = {}
        self._emis_helper = EmissionsCalculator(self.config)
        if self._sim_params:
            self._emis_helper.set_overrides(self._sim_params)
        self._coef_rows = self._emis_helper._load_coef_vehiculos()
        self._ef_rows = self._emis_helper._load_EF()
        self._ds_nodes: Dict[str, Node] = {}
        self._ds_graph: Optional[nx.DiGraph] = None
        self._ds_base_id: Optional[str] = None
        # Etiquetas de stand que no pudieron resolverse como puestos reales
        self._unknown_stands: Set[str] = set()
        if dataset and dataset.get("rows"):
            self._load_dataset_graph(dataset)
        self._debug_enabled = debug_enabled
        self._date_filter_from: Optional[datetime.date] = None
        self._date_filter_to: Optional[datetime.date] = None

    # --------- API ---------
    def set_operations_df(self, df) -> None:
        self._ops = self._parse_operations(df)
        if self._date_filter_from is not None and self._date_filter_to is not None:
            filtered: List[OperationRecord] = []
            for op in self._ops:
                d = getattr(op, "date", None)
                if d is not None and self._date_filter_from <= d <= self._date_filter_to:
                    filtered.append(op)
            self._ops = filtered

    def set_date_filter(self, date_from: datetime.date, date_to: datetime.date) -> None:
        self._date_filter_from = date_from
        self._date_filter_to = date_to

    def compute(self) -> Tuple[Dict[str, CirculationResult], List[str]]:
        self._gse_matrix = self._load_gse_aircraft_matrix()
        global_warnings: List[str] = []
        results: Dict[str, CirculationResult] = {}
        if not self._ops:
            return results, ["No hay operaciones válidas para calcular circulación."]
        for veh in self._vehicle_codes():
            res = self._simulate_vehicle(veh)
            if res:
                results[veh] = res
                global_warnings.extend(res.warnings)
        # Advertir si hay etiquetas de puesto que no se pudieron asociar a ningún 'puesto'
        if self._unknown_stands:
            lst = ", ".join(sorted(self._unknown_stands))
            global_warnings.append(
                f"No se encontraron puestos en el grafo para las etiquetas de stand: {lst}. "
                "Revise la columna Puerta_asignada de la tabla de operaciones y la tabla de circulación."
            )
        if not results:
            global_warnings.append("Ningún vehículo tuvo operaciones aplicables.")
        return results, global_warnings

    def diagnostic_report(self) -> str:
        """Genera un pequeño escenario sintético para verificar cálculos básicos."""
        puestos = [n for n in self.model.nodes.values() if n.kind == 'puesto']
        base_id = self.model.default_base_id
        if not base_id or len(puestos) < 2:
            return "Prueba sintética omitida (requiere base y al menos dos puestos)."
        p1, p2 = puestos[0], puestos[1]
        veh = "GPU"
        dist1 = self._distance_for_vehicle_move(base_id, p1.id, veh)
        dist2 = self._distance_for_vehicle_move(p1.id, p2.id, veh)
        if dist1 is None or dist2 is None:
            return "Prueba sintética: no hay caminos disponibles base↔puestos."
        params = self._params_for_vehicle(veh)
        total_dist = dist1 + dist2
        total_time = total_dist / params['vel_kmh']
        k = self._k_value(veh, "CO2")
        if k is None:
            return "Prueba sintética: no se pudo obtener K para GPU/CO2."
        emis = (k * params['fc_cir'] / params['vel_kmh']) * total_dist
        # Intentar recuperar también las rutas explícitas de nodos para el diagnóstico
        def _path_ids(start: str, end: str) -> Optional[List[str]]:
            try:
                if self.model.G.number_of_nodes() > 0:
                    return nx.shortest_path(self.model.G, start, end, weight='weight')
            except (nx.NetworkXNoPath, nx.NodeNotFound):
                pass
            if self._ds_graph is not None:
                try:
                    return nx.shortest_path(self._ds_graph, start, end, weight='weight')
                except (nx.NetworkXNoPath, nx.NodeNotFound):
                    return None
            return None

        def _fmt_path(ids: Optional[List[str]]) -> str:
            if not ids:
                return "(camino no disponible)"
            labels = []
            for nid in ids:
                n = self.model.nodes.get(nid) or self._ds_nodes.get(nid)
                if n is not None:
                    labels.append(f"{nid} ({n.name})")
                else:
                    labels.append(nid)
            return " → ".join(labels)

        path1 = _path_ids(base_id, p1.id)
        path2 = _path_ids(p1.id, p2.id)

        lines = []
        lines.append(("Prueba sintética GPU: Base→{p1}→{p2}: {dist:.3f} km, {time:.3f} h, "
                      "CO2 ≈ {emis:.2f} g.").format(p1=p1.name, p2=p2.name,
                                                      dist=total_dist, time=total_time, emis=emis))
        lines.append("Camino Base → {p1}: {path}".format(p1=p1.name, path=_fmt_path(path1)))
        lines.append("Camino {p1} → {p2}: {path}".format(p1=p1.name, p2=p2.name, path=_fmt_path(path2)))
        return "\n".join(lines)

    def synthetic_debug_report(self) -> str:
        if not self._debug_enabled:
            return ""
        fc = 0.5
        vel = 15.0
        gases_k = {
            "CO2": 120.0,
            "CO": 3.5,
            "HC": 0.7,
            "NOx": 1.2,
            "SOx": 0.5,
            "PM10": 0.18,
        }
        legs = [
            {"idx": 1, "from": "BASE", "to": "P1", "dist": 0.45, "arr": 7.00, "svc": 0.20},
            {"idx": 2, "from": "P1", "to": "P3", "dist": 0.62, "arr": 7.30, "svc": 0.18},
            {"idx": 3, "from": "P3", "to": "P2", "dist": 0.38, "arr": 8.00, "svc": 0.25},
            {"idx": 4, "from": "P2", "to": "P4", "dist": 0.55, "arr": 8.50, "svc": 0.22},
            {"idx": 5, "from": "P4", "to": "P1", "dist": 0.50, "arr": 9.10, "svc": 0.20},
            {"idx": 6, "from": "P1", "to": "P2", "dist": 0.33, "arr": 9.50, "svc": 0.18},
        ]
        lines = []
        lines.append("=== Informe de prueba de circulación (circuito de 4 puestos, 6 operaciones) ===")
        lines.append(f"Parámetros: FC_CIR={fc}, Velocidad={vel} km/h")
        lines.append("Fórmula de emisiones por gas: E_g = (K_g * FC_CIR / Velocidad) * Distancia")
        lines.append("Flujo másico: mf_g = E_g / (tiempo_h * 3600)")
        lines.append("")
        for leg in legs:
            time_h = leg['dist'] / vel
            lines.append(f"Operación #{leg['idx']} | Recorrido {leg['from']} → {leg['to']}")
            lines.append(f"  Distancia = {leg['dist']:.3f} km  |  Tiempo = {time_h:.3f} h ({time_h*60:.1f} min)")
            lines.append(f"  Llegada programada = {leg['arr']:.2f} h | Tiempo servicio estimado = {leg['svc']:.2f} h")
            for gas, k in gases_k.items():
                emis = (k * fc / vel) * leg['dist']
                flow = emis / (time_h * 3600) if time_h > 0 else 0.0
                lines.append(
                    f"    {gas}: E = ({k:.2f} * {fc:.2f} / {vel:.2f}) * {leg['dist']:.3f} = {emis:.3f} g, "
                    f"mf = {emis:.3f} / ({time_h:.3f} * 3600) = {flow:.6f} g/s"
                )
            lines.append("")
        total_dist = sum(l['dist'] for l in legs)
        total_time = total_dist / vel
        lines.append(f"Total distancia = {total_dist:.3f} km, tiempo rodado = {total_time:.3f} h")
        for gas, k in gases_k.items():
            emis = (k * fc / vel) * total_dist
            lines.append(f"Total {gas} = {emis:.3f} g")
        return "\n".join(lines)

    # --------- Preparación de datos ---------
    def _parse_operations(self, df) -> List[OperationRecord]:
        ops: List[OperationRecord] = []
        if df is None:
            return ops

        try:
            cols = list(df.columns)
        except Exception:
            cols = []
        ncols = len(cols)

        def _col_index(default_idx: Optional[int]) -> Optional[int]:
            if default_idx is None:
                return None
            if default_idx < 0 or default_idx >= ncols:
                return None
            return default_idx

        # Layout real de la planilla (0-based):
        # 0: DIA
        # 1: TIPOVUEL
        # 2: Tipo_de_operación
        # 3: Aerolínea
        # 4: Aeronave
        # 5: Numero de Vuelo
        # 6: Puerta_asignada
        # 7: Hora_IN_GATE
        # 8: Hora_OUT_Gate
        # 9: TIPO_SER (opcional)
        dia_col = _col_index(0)
        stand_col = _col_index(6)
        arr_col = _col_index(7)
        dep_col = _col_index(8)
        ac_col = _col_index(4)
        tipo_col = _col_index(9)

        for idx, row in df.iterrows():
            op_date: Optional[datetime.date] = None
            if dia_col is not None:
                try:
                    dia_val = row.iloc[dia_col]
                except Exception:
                    dia_val = None
                if isinstance(dia_val, datetime.datetime):
                    op_date = dia_val.date()
                elif isinstance(dia_val, datetime.date):
                    op_date = dia_val
                elif isinstance(dia_val, str):
                    s = dia_val.strip()
                    if s:
                        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                            try:
                                op_date = datetime.datetime.strptime(s, fmt).date()
                                break
                            except Exception:
                                continue
            # Puerta / stand: si está vacía o no existe en el grafo,
            # igualmente usamos un stand sintético sin manga para NO
            # descartar la operación en el cálculo de servicio.
            try:
                stand_val = row.iloc[stand_col] if stand_col is not None else None
            except Exception:
                stand_val = None
            stand_txt = str(stand_val).strip() if stand_val is not None else ""
            if not stand_txt:
                stand_txt = f"OP{idx}"
            stand = self._resolve_stand(stand_txt)
            if stand:
                stand_id = stand.id
                stand_name = stand.name
                has_jet = stand.has_jetbridge
            else:
                # Registrar stands no resueltos (si no son sintéticos OP#)
                if not stand_txt.startswith("OP"):
                    self._unknown_stands.add(stand_txt)
                # Para circulación, usar un ID sintético que no coincida con nodos reales,
                # de modo que no se reutilicen vías/hubs/bases con el mismo identificador.
                stand_id = f"OP{idx}"
                # Mantener el texto original como nombre legible para tablas de servicio.
                stand_name = stand_txt
                has_jet = False

            # Horas de arribo/salida
            try:
                arr_val = row.iloc[arr_col] if arr_col is not None else None
            except Exception:
                arr_val = None
            try:
                dep_val = row.iloc[dep_col] if dep_col is not None else None
            except Exception:
                dep_val = None
            arr = self._to_float(arr_val)
            dep = self._to_float(dep_val)
            if arr is None:
                continue
            if dep is None or dep < arr:
                dep = arr

            # Aeronave y tipo de servicio (este último hoy no se usa para el cálculo)
            try:
                ac_val = row.iloc[ac_col] if ac_col is not None else ""
            except Exception:
                ac_val = ""
            try:
                tipo_val = row.iloc[tipo_col] if tipo_col is not None else ""
            except Exception:
                tipo_val = ""
            aircraft = str(ac_val or "").strip().upper()
            tipo_ser = str(tipo_val or "").strip().upper()

            ops.append(OperationRecord(
                index=idx,
                arr=arr,
                dep=dep,
                stand_id=stand_id,
                stand_name=stand_name,
                has_jetbridge=has_jet,
                aircraft=aircraft,
                tipo_ser=tipo_ser,
                date=op_date,
            ))
        ops.sort(key=lambda x: (x.arr, x.index))
        return ops

    def _load_gse_aircraft_matrix(self) -> Dict[str, Dict[str, Dict[str, float]]]:
        data = self.config.get_dataset('gsexaeronaves')
        cols = list(data.get('columns', []))
        rows = data.get('rows', [])
        if not cols or not rows:
            return {}
        aircraft_map: List[Tuple[str, str, str, str]] = []  # (ac, rampa, remota, tiempo)
        i = 1
        while i < len(cols):
            ac = cols[i]
            r_col = cols[i]
            remote_col = cols[i+1] if i+1 < len(cols) else cols[i]
            time_col = cols[i+2] if i+2 < len(cols) else cols[i]
            aircraft_map.append((str(ac).strip().upper(), r_col, remote_col, time_col))
            i += 3
        matrix: Dict[str, Dict[str, Dict[str, float]]] = {}
        for row in rows:
            veh = str(row.get('GSE') or '').strip()
            if not veh:
                continue
            veh_matrix: Dict[str, Dict[str, float]] = {}
            for ac, col_r, col_s, col_t in aircraft_map:
                try:
                    count_r = float(row.get(col_r, 0) or 0)
                except Exception:
                    count_r = 0.0
                try:
                    count_s = float(row.get(col_s, 0) or 0)
                except Exception:
                    count_s = 0.0
                try:
                    time_val = float(row.get(col_t, 0) or 0)
                except Exception:
                    time_val = 0.0
                veh_matrix[ac] = {
                    'count_rampa': count_r,
                    'count_remota': count_s,
                    'service_time': time_val,
                }
            matrix[veh] = veh_matrix
        return matrix

    def _load_dataset_graph(self, dataset: dict) -> None:
        rows = dataset.get("rows", [])
        if not rows:
            return
        graph = nx.DiGraph()
        nodes: Dict[str, Node] = {}
        base_candidate: Optional[str] = None

        def _bool(val) -> bool:
            if isinstance(val, bool):
                return val
            if val is None:
                return False
            text = str(val).strip().lower()
            return text in {"1", "true", "si", "sí", "yes", "y"}

        for row in rows:
            cat = str(row.get("Categoria", "")).strip().lower()
            if cat == "nodo":
                nid = str(row.get("ID") or row.get("Id") or row.get("Nodo") or "").strip()
                if not nid:
                    continue
                name = str(row.get("Nombre") or nid).strip()
                kind = str(row.get("Tipo") or row.get("Sentido") or "via").strip().lower()
                lat = float(row.get("Lat", 0) or 0)
                lon = float(row.get("Lon", 0) or 0)
                has_jet = _bool(row.get("Manga"))
                node = Node(
                    id=nid,
                    name=name,
                    lat=lat,
                    lon=lon,
                    kind=kind,
                    has_jetbridge=has_jet,
                    is_hub_bus=_bool(row.get("Es_hub_BUS")),
                    is_hub_sta=_bool(row.get("Es_hub_STA")),
                    is_hub_bag=_bool(row.get("Es_hub_BAG")),
                    is_hub_bel=_bool(row.get("Es_hub_BEL")),
                    is_hub_gpu=_bool(row.get("Es_hub_GPU")),
                    is_hub_fue=_bool(row.get("Es_hub_FUE")),
                )
                nodes[nid] = node
                graph.add_node(nid)
                if node.kind == 'base' and not base_candidate:
                    base_candidate = nid
            elif cat == "ruta":
                n1 = str(row.get("Desde") or row.get("Origen") or row.get("n1") or "").strip()
                n2 = str(row.get("Hasta") or row.get("Destino") or row.get("n2") or "").strip()
                if not n1 or not n2:
                    continue
                try:
                    length = float(row.get("Dist_km", row.get("Dist", 0)) or 0)
                except Exception:
                    length = 0.0
                sentido = str(row.get("Sentido") or "").lower()
                is_one_way = sentido.startswith("solo")
                graph.add_edge(n1, n2, weight=max(length, 0.01))
                if not is_one_way:
                    graph.add_edge(n2, n1, weight=max(length, 0.01))

        self._ds_nodes = nodes
        self._ds_graph = graph if nodes else None
        self._ds_base_id = base_candidate or (next(iter(nodes.keys())) if nodes else None)

    # --------- Simulación ---------
    def _simulate_vehicle(self, veh_code: str) -> Optional[CirculationResult]:
        """Corre la simulación para un tipo de vehículo."""
        gse_ops = self._gse_matrix.get(veh_code)
        if not gse_ops:
            return None

        # Cargar parámetros de circulación
        params = self._params_for_vehicle(veh_code)
        vel_kmh = params["vel_kmh"]
        fc_cir = params["fc_cir"]

        # Cargar multiplicador por tipo de vuelo (ej: Internacional x 1.2, Nacional x 1.0)
        # Se lee desde el dataset 'tipos_vuelo' configurado por el usuario
        flight_coeffs = {}
        try:
            tv_ds = self.config.get_dataset("tipos_vuelo")
            for r in tv_ds.get("rows", []):
                code = str(r.get("TipoVuelo", "")).strip().upper()
                try:
                    val = float(str(r.get("Multiplicador", "1.0")).replace(",", "."))
                except:
                    val = 1.0
                if code: flight_coeffs[code] = val
        except:
            pass

        # Simulación de flota (FMT) y Distancia
        base_id = self._base_for_vehicle(veh_code)
        relevant_ops = []
        skipped_zero = 0
        for op in self._ops:
            count, svc = self._operation_demand(veh_code, gse_ops, op)
            if count <= 0:
                skipped_zero += 1
                continue  # Operación sin demanda de este GSE (count=0 o aeronave no configurada)
            # Aplicar coeficiente de vuelo al consumo de combustible
            flight_type_coeff = flight_coeffs.get(op.tipo_ser, 1.0)
            relevant_ops.append((op, count, svc, flight_type_coeff))

        if not relevant_ops:
            # Si se saltaron ops por count=0, no es un error sino comportamiento esperado.
            # No retornar resultado -> este vehículo no aparece en los resultados.
            return None

        state_list: List[Dict] = []  # cada dict: {'pos':node_id, 'available':hora}
        total_dist = 0.0
        total_service_time_h = 0.0
        warnings: List[str] = []
        intervals: List[Tuple[float, float]] = []
        total_gases_per_type: Dict[str, float] = defaultdict(float)
        # Acumulador por puesto para discriminación de emisiones
        stand_accum: Dict[str, Dict] = {}

        for op, units, svc_time, flight_type_coeff in relevant_ops:
            # units puede ser fraccionario (ej: 0.3 tanques de combustible)
            # pero para la simulación de flota se necesitan unidades enteras reales.
            # Si count=2, hay que despachar 2 vehículos simultáneos al mismo puesto.
            units_int = max(1, int(math.ceil(float(units))))

            # Emitir UN movimiento por cada unidad requerida.
            # Cada vehículo cubre el tiempo de servicio completo (svc_time),
            # no el tiempo total multiplicado. Así la flota mínima se estima correctamente.
            for _unit_idx in range(units_int):
                move = self._assign_vehicle(state_list, op, svc_time, vel_kmh, veh_code)
                if not move:
                    warnings.append(
                        f"[{veh_code}] No se pudo asignar vehículo #{_unit_idx+1}/{units_int} "
                        f"para operación #{op.index} en {op.stand_name}."
                    )
                    continue
                total_dist += move['distance']
                intervals.append((move['travel_start'], move['service_end']))
                total_service_time_h += svc_time

                # Calcular gases para este movimiento y sumarlos
                gases_for_move = self._emissions_for_vehicle(
                    veh_code, move['distance'], vel_kmh, fc_cir * flight_type_coeff)
                for gas_type, amount in gases_for_move.items():
                    total_gases_per_type[gas_type] += amount

                stand_key = op.stand_name or op.stand_id
                if stand_key not in stand_accum:
                    stand_accum[stand_key] = {'dist_km': 0.0, 'svc_h': 0.0}
                    for _g in self._gases:
                        stand_accum[stand_key][_g] = 0.0
                stand_accum[stand_key]['dist_km'] += move['distance']
                stand_accum[stand_key]['svc_h'] += svc_time
                for gas_type, g_data in gases_for_move.items():
                    stand_accum[stand_key][gas_type] = (
                        stand_accum[stand_key].get(gas_type, 0.0) + g_data.get('g', 0.0))

        max_fleet = self._max_concurrent(intervals)
        res_gases = dict(total_gases_per_type) # Convert defaultdict to dict

        return CirculationResult(
            distance_km=total_dist,
            circulation_time_h=total_dist / vel_kmh,
            service_time_h=total_service_time_h,
            fleet=max_fleet,
            warnings=warnings,
            gases=res_gases,
            by_stand=dict(stand_accum)
        )

    def simulate_vehicle_with_events(self, veh_code: str) -> List[Dict]:
        """Versión de simulación que devuelve una lista de diccionarios con eventos para el StepByStep."""
        events = []
        gse_ops = self._gse_matrix.get(veh_code)
        if not gse_ops: return []
        
        applicable_ops = []
        for op in self._ops:
            m = gse_ops.get(op.aircraft)
            if not m: continue
            units = m['count_rampa'] if (not op.has_jetbridge) else m['count_remota']
            if units <= 0: continue
            applicable_ops.append((op, units, m['service_time']))
            
        if not applicable_ops: return []
        
        # Lógica simplificada para obtener el flujo cronológico
        base_id = self._base_for_vehicle(veh_code)
        pos_actual = base_id
        
        for i, (op, units, svc_time) in enumerate(applicable_ops):
            # Movimiento a puesto
            dist = self._distance_for_vehicle_move(pos_actual, op.stand_id, veh_code)
            detail = f"Misión {i+1}: Atender {op.aircraft} en {op.stand_name}"
            
            # Obtener path para el JS
            path = self._get_path_coords(pos_actual, op.stand_id, veh_code)
            
            events.append({
                'time': op.arr,
                'type': 'DESPLAZAMIENTO',
                'from': pos_actual,
                'to': op.stand_id,
                'detail': f"Hacia puesto: {detail}",
                'path_coords': path
            })
            
            events.append({
                'time': op.arr,
                'type': 'SERVICIO',
                'from': op.stand_id,
                'to': op.stand_id,
                'detail': f"Iniciando servicio a {op.aircraft} ({units} unidades)",
                'path_coords': None
            })
            
            pos_actual = op.stand_id
            
        return events

    def _get_path_coords(self, start: str, end: str, veh: str) -> Optional[List[List[float]]]:
        """Devuelve lista de [lat, lon] para una ruta."""
        hub = self._hub_node_for(veh)
        req_hub = self._requires_hub_cycle(veh)
        
        ids = []
        try:
            if req_hub and hub and start != hub and end != hub:
                p1 = nx.shortest_path(self.model.G, start, hub, weight='weight')
                p2 = nx.shortest_path(self.model.G, hub, end, weight='weight')
                ids = p1 + p2[1:]
            else:
                ids = nx.shortest_path(self.model.G, start, end, weight='weight')
        except (nx.NetworkXNoPath, nx.NodeNotFound):
            # Fallback to direct path if hub path fails
            try:
                ids = nx.shortest_path(self.model.G, start, end, weight='weight')
            except (nx.NetworkXNoPath, nx.NodeNotFound):
                return None
            
        coords = []
        for nid in ids:
            n = self.model.nodes.get(nid)
            if n: coords.append([n.lat, n.lon])
        return coords

    def _assign_vehicle(self, states: List[Dict], op: OperationRecord,
                         svc_time: float, vel: float, veh: str) -> Optional[Dict[str, float]]:
        candidates = []
        req_hub_cycle = self._requires_hub_cycle(veh)
        hub_node = self._hub_node_for(veh) if req_hub_cycle else None

        for st in states:
            travel_parts = []
            valid_path = True
            pos_actual = st['pos']

            # Si el vehículo requiere obligatoriamente pasar por su HUB entre servicios y NO está en el HUB
            # ni el destino es el HUB, el trayecto es [Posición Actual -> HUB -> Puesto Destino]
            if req_hub_cycle and hub_node and pos_actual != hub_node and op.stand_id != hub_node:
                d1 = self._shortest_distance(pos_actual, hub_node)
                d2 = self._shortest_distance(hub_node, op.stand_id)
                if d1 is None or d2 is None:
                    valid_path = False
                else:
                    travel_parts = [d1, d2]
            else:
                dist = self._distance_for_vehicle_move(pos_actual, op.stand_id, veh)
                if dist is None:
                    valid_path = False
                else:
                    travel_parts = [dist]

            if not valid_path:
                continue

            total_dist = sum(travel_parts)
            travel_time = total_dist / vel
            depart = st['available']
            travel_end = depart + travel_time
            arrival_time = max(op.arr, travel_end)
            service_end = arrival_time + max(svc_time, op.dep - op.arr)
            candidates.append({
                'distance': total_dist,
                'travel_start': depart,
                'arrival_time': arrival_time,
                'service_end': service_end,
                'state': st,
                'travel_time': travel_time,
                'new_vehicle': False,
            })
        base_move = self._new_vehicle_move(op, svc_time, vel, veh)
        if base_move:
            candidates.append(base_move)
        if not candidates:
            return None

        # Ordenar preferiendo el candidato que llegue igual de temprano que el resto
        # pero viajando la menor distancia. base_move SIEMPRE llegará exactamente en op.arr
        candidates.sort(key=lambda m: (m['arrival_time'], m['distance']))
        chosen = candidates[0]

        if chosen['new_vehicle']:
            st = chosen['state']
            states.append(st)
        else:
            st = chosen['state']
            st['pos'] = op.stand_id
            st['available'] = chosen['service_end']
        return chosen

    def _new_vehicle_move(self, op: OperationRecord, svc_time: float, vel: float, veh: str) -> Optional[Dict[str, float]]:
        hub_node = self._hub_node_for(veh) if self._requires_hub_cycle(veh) else None
        base_id = hub_node if hub_node else self._base_for_vehicle(veh)
        if not base_id:
            return None
        dist = self._distance_for_vehicle_move(base_id, op.stand_id, veh)
        if dist is None:
            return None
        travel_time = dist / vel
        depart = max(0.0, op.arr - travel_time)
        arrival_time = op.arr # Ideal dispatch
        service_end = arrival_time + max(svc_time, op.dep - op.arr)
        state = {'pos': op.stand_id, 'available': service_end}
        return {
            'distance': dist,
            'travel_start': depart,
            'arrival_time': arrival_time,
            'service_end': service_end,
            'state': state,
            'travel_time': travel_time,
            'new_vehicle': True,
        }

    # --------- Helpers ---------
    def _vehicle_codes(self) -> List[str]:
        return ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]

    def _operation_demand(self, veh: str, matrix: Dict[str, Dict[str, float]],
                          op: OperationRecord) -> Tuple[float, float]:
        data = matrix.get(op.aircraft)
        if not data:
            # Aeronave no configurada en gsexaeronaves -> no demanda este GSE.
            return 0.0, 0.0
        # Seleccionar cantidad según tipo de puesto (manga / remota)
        count = data['count_rampa'] if op.has_jetbridge else data['count_remota']
        # Si el usuario configuró count=0 -> este GSE no se usa para esta aeronave.
        # Esto es correcto y se respeta (no se fuerza a 1).
        # Tiempo de servicio: usar el de la tabla si está definido y > 0,
        # si no, usar el tiempo de permanencia de la operación (dep-arr).
        svc_time_tbl = data.get('service_time', 0)
        try:
            svc_time_tbl = float(svc_time_tbl) if svc_time_tbl else 0.0
        except Exception:
            svc_time_tbl = 0.0
        svc_time = svc_time_tbl if svc_time_tbl > 0.0 else max(op.dep - op.arr, 0.1)
        if veh in ('BUS', 'STA', 'BAG', 'BEL') and op.has_jetbridge:
            # Vehículos que no operan cuando hay manga.
            return 0.0, svc_time
        # NOTA: si el usuario puso 0 vehículos para esta aeronave,
        # se respeta: la operación NO requiere ese GSE.
        # No forzar count=1 aunque no haya grafo propio.
        return count, svc_time

    def _distance_for_vehicle_move(self, start: str, end: str, veh: str) -> Optional[float]:
        if start == end:
            return 0.0
        if self._requires_hub_cycle(veh):
            hub = self._hub_node_for(veh)
            if hub and hub not in (start, end):
                d1 = self._shortest_distance(start, hub)
                d2 = self._shortest_distance(hub, end)
                if d1 is None or d2 is None:
                    return None
                return d1 + d2
        
        dist = self._shortest_distance(start, end)
        if dist is None: return None
        
        # Ajuste por Rear Entry si el destino es un puesto con esa opción
        node_end = self.model.nodes.get(end) or self._ds_nodes.get(end)
        if node_end and getattr(node_end, 'is_rear_entry', False):
            # Asumimos un ajuste de 0.02 km (20m) si es entrada trasera?
            # O mejor, lo dejamos como una propiedad de distancia adicional.
            dist += 0.02 
            
        return dist

    def _requires_hub_cycle(self, veh: str) -> bool:
        return veh in ("BUS", "STA", "BAG", "BEL", "GPU", "FUE")

    def _shortest_distance(self, start: str, end: str) -> Optional[float]:
        try:
            if self.model.G.number_of_nodes() > 0:
                return nx.shortest_path_length(self.model.G, start, end, weight='weight')
        except (nx.NetworkXNoPath, nx.NodeNotFound):
            pass
        if self._ds_graph is not None:
            try:
                return nx.shortest_path_length(self._ds_graph, start, end, weight='weight')
            except (nx.NetworkXNoPath, nx.NodeNotFound):
                return None
        return None

    def _base_for_vehicle(self, veh: str) -> Optional[str]:
        # 1) Base específica por tipo de vehículo, si está configurada en el modelo
        try:
            veh_bases = getattr(self.model, "vehicle_bases", {}) or {}
        except Exception:
            veh_bases = {}
        base_spec = veh_bases.get(veh)
        if base_spec:
            return base_spec

        # 2) Fallback al comportamiento clásico: una única base global
        base = self.model.default_base_id
        if base:
            return base
        for node in self.model.nodes.values():
            if node.kind == 'base':
                return node.id
        if self._ds_base_id:
            return self._ds_base_id
        return None

    def _hub_node_for(self, veh: str) -> Optional[str]:
        """Devuelve el nodo hub específico según tipo de vehículo.

        BUS → is_hub_bus
        STA → is_hub_sta
        BAG → is_hub_bag
        BEL → is_hub_bel
        GPU → is_hub_gpu
        FUE → is_hub_fue
        """
        attr = None
        if veh == 'BUS':
            attr = 'is_hub_bus'
        elif veh == 'STA':
            attr = 'is_hub_sta'
        elif veh == 'BAG':
            attr = 'is_hub_bag'
        elif veh == 'BEL':
            attr = 'is_hub_bel'
        elif veh == 'GPU':
            attr = 'is_hub_gpu'
        elif veh == 'FUE':
            attr = 'is_hub_fue'
        if not attr:
            return None
        for node in self.model.nodes.values():
            if getattr(node, attr, False):
                return node.id
        for node in self._ds_nodes.values():
            if getattr(node, attr, False):
                return node.id
        return None

    def _params_for_vehicle(self, veh: str) -> Dict[str, float]:
        defaults = {"fc_cir": 0.5, "vel_kmh": 15.0}
        return self._circ_params.get(veh, defaults)

    def _k_value(self, veh: str, gas: str) -> Optional[float]:
        vrow = self._emis_helper._find_vehicle_row(veh, self._coef_rows)
        if not vrow:
            return None
        hp_vehicle = self._emis_helper._vehicle_hp(vrow)
        efrow = self._emis_helper._select_EF_row_for_veh(veh, hp_vehicle, self._ef_rows)
        if not efrow:
            return None
        age_o = self._emis_helper._ovr(veh, 'age', None)
        tutil_o = self._emis_helper._ovr(veh, 't_util', None)
        hp_o = self._emis_helper._ovr(veh, 'hp_vehicle', None)
        return self._emis_helper._k_for(vrow, efrow, gas, veh,
                                        age_override=age_o,
                                        t_util_override=tutil_o,
                                        hp_override=hp_o)

    def _emissions_for_vehicle(self, veh: str, distance_km: float,
                               vel: float, fc: float) -> Dict[str, Dict[str, float]]:
        gases: Dict[str, Dict[str, float]] = {}
        if distance_km <= 0 or vel <= 0:
            return gases
        for gas in self._gases:
            k = self._k_value(veh, gas)
            if k is None:
                continue
            emis_g = (k * fc / vel) * distance_km
            time_h = distance_km / vel
            emis_gps = emis_g / (time_h * 3600) if time_h > 0 else 0.0
            gases[gas] = {"g": emis_g, "gps": emis_gps}
        return gases

    def _max_concurrent(self, intervals: List[Tuple[float, float]]) -> int:
        if not intervals:
            return 0
        events = []
        for start, end in intervals:
            start = max(0.0, start)
            events.append((start, 1))
            events.append((max(start, end), -1))
        events.sort()
        active = 0
        max_active = 0
        for _, delta in events:
            active += delta
            max_active = max(max_active, active)
        return max_active

    def _resolve_stand(self, label: str) -> Optional[Node]:
        """Resuelve el puesto a partir de una etiqueta (texto de la tabla de operaciones).

        Solo devuelve nodos de tipo 'puesto'. Si no hay ningún puesto cuyo nombre o ID
        coincida con la etiqueta, devuelve None en lugar de reutilizar vías, hubs o
        la base aunque el ID sea igual.
        """
        target = label.strip().lower()

        # 1) Buscar por nombre en nodos del modelo (solo puestos)
        for node in self.model.nodes.values():
            if node.kind == 'puesto' and str(node.name).strip().lower() == target:
                return node

        # 2) Buscar por ID en nodos del modelo (solo puestos)
        for node in self.model.nodes.values():
            if node.kind == 'puesto' and node.id.lower() == target:
                return node

        # 3) Buscar en nodos provenientes del dataset de circulación (solo puestos)
        for node in self._ds_nodes.values():
            if node.kind == 'puesto' and node.name.strip().lower() == target:
                return node

        for node in self._ds_nodes.values():
            if node.kind == 'puesto' and node.id.lower() == target:
                return node

        # Si no se encontró ningún puesto, devolver None (no usar vías/hubs/bases por error)
        return None

    def _to_float(self, value) -> Optional[float]:
        if value in (None, ""):
            return None
        try:
            return float(value)
        except Exception:
            return None


class CirculationResultsDialog(QtWidgets.QDialog):
    """Diálogo de resultados de circulación con vista por GSE, por unidad (kg/Tn)
    y desglose por puesto. Incluye exportación a CSV completa."""

    _gases = CirculationCalculator._gases

    def __init__(self, parent, data: Dict[str, CirculationResult],
                 global_warnings: List[str], diagnostics: str,
                 debug_report_path: Optional[str]):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Resultados – Emisiones por circulación")
        self.resize(1050, 560)
        self._data = data
        self._global_warnings = global_warnings

        veh_order = ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]
        # Mostrar SIEMPRE los 12 GSE. Los que no tienen datos => fila en gris con ceros.
        extra = [v for v in sorted(data.keys()) if v not in veh_order]
        self._vehs = veh_order + extra
        _zero_cr = CirculationResult(distance_km=0.0, circulation_time_h=0.0,
                                    service_time_h=0.0, fleet=0,
                                    warnings=[], gases={}, by_stand={})
        self._data_full = {v: data.get(v, _zero_cr) for v in self._vehs}
        layout = QtWidgets.QVBoxLayout(self)
        tabs = QtWidgets.QTabWidget()
        layout.addWidget(tabs, 1)

        # ── Pestaña 1: Por GSE (gramos y g/s) ───────────────────────────
        tab1 = QtWidgets.QWidget()
        t1l = QtWidgets.QVBoxLayout(tab1)
        self._tbl_g = self._make_gse_table(unit="g")
        t1l.addWidget(self._tbl_g)
        tabs.addTab(tab1, "Por GSE (g / g·s⁻¹)")

        # ── Pestaña 2: Por GSE (kg y toneladas) ─────────────────────────
        tab2 = QtWidgets.QWidget()
        t2l = QtWidgets.QVBoxLayout(tab2)
        self._tbl_kg = self._make_gse_table(unit="kg")
        t2l.addWidget(self._tbl_kg)
        tabs.addTab(tab2, "Por GSE (kg / Tn)")

        # ── Pestaña 3: Desglose por Puesto ───────────────────────────────
        tab3 = QtWidgets.QWidget()
        t3l = QtWidgets.QVBoxLayout(tab3)

        # Selector de GSE
        sel_row = QtWidgets.QHBoxLayout()
        sel_row.addWidget(QtWidgets.QLabel("GSE:"))
        self._cmbStandVeh = QtWidgets.QComboBox()
        self._cmbStandVeh.addItems(self._vehs)
        self._cmbStandVeh.currentTextChanged.connect(self._refresh_stand_table)
        sel_row.addWidget(self._cmbStandVeh)
        sel_row.addStretch(1)
        t3l.addLayout(sel_row)

        self._tbl_stand = QtWidgets.QTableWidget()
        self._tbl_stand.setAlternatingRowColors(True)
        self._tbl_stand.setSortingEnabled(True)
        self._tbl_stand.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        t3l.addWidget(self._tbl_stand)
        tabs.addTab(tab3, "Por Puesto")

        # Populate stand table on start
        if self._vehs:
            self._refresh_stand_table(self._vehs[0])

        # ── Botones ──────────────────────────────────────────────────────
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)

        btnExportGSE = QtWidgets.QPushButton("Exportar (Por GSE) CSV…")
        btnExportGSE.clicked.connect(self._export_gse_csv)
        btn_row.addWidget(btnExportGSE)

        btnExportStand = QtWidgets.QPushButton("Exportar (Por Puesto) CSV…")
        btnExportStand.clicked.connect(self._export_stand_csv)
        btn_row.addWidget(btnExportStand)

        btnClose = QtWidgets.QPushButton("Cerrar")
        btnClose.clicked.connect(self.accept)
        btn_row.addWidget(btnClose)

        layout.addLayout(btn_row)

    # ── helpers internos ─────────────────────────────────────────────────

    def _make_gse_table(self, unit: str) -> QtWidgets.QTableWidget:
        """Crea la tabla resumen por GSE, en gramos (unit='g') o kg/Tn (unit='kg')."""
        tbl = QtWidgets.QTableWidget()
        tbl.setAlternatingRowColors(True)
        tbl.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)

        base_h = ["GSE", "Dist [km]", "Circ. [h]", "Serv. [h]", "Flota mín."]
        if unit == "g":
            gas_h = sum([[f"{g} (g)", f"{g} [g/s]"] for g in self._gases], [])
        else:
            gas_h = sum([[f"{g} (kg)", f"{g} (Tn)"] for g in self._gases], [])
        headers = base_h + gas_h
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(len(self._vehs))
        tbl.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)

        _gray = QtGui.QColor("#9E9E9E")
        _italic_font = QtGui.QFont(); _italic_font.setItalic(True)
        for row, veh in enumerate(self._vehs):
            _has_data = veh in self._data
            res = self._data_full[veh]
            def _mk(txt, is_zero=not _has_data):
                it = QtWidgets.QTableWidgetItem(str(txt))
                if is_zero:
                    it.setForeground(QtGui.QColor("#9E9E9E"))
                    it.setFont(_italic_font)
                    it.setToolTip("GSE sin operaciones en este cálculo (count=0 o aeronave no configurada)")
                return it
            tbl.setItem(row, 0, _mk(veh))
            tbl.setItem(row, 1, _mk(f"{res.distance_km:.3f}"))
            tbl.setItem(row, 2, _mk(f"{res.circulation_time_h:.3f}"))
            tbl.setItem(row, 3, _mk(f"{res.service_time_h:.3f}"))
            tbl.setItem(row, 4, _mk("-" if not _has_data else str(res.fleet)))
            col = 5
            for gas in self._gases:
                rec = res.gases.get(gas)
                g_val = (rec.get("g", 0.0) or 0.0) if rec else 0.0
                if unit == "g":
                    gps_val = (rec.get("gps", 0.0) or 0.0) if rec else 0.0
                    tbl.setItem(row, col,     _mk(f"{g_val:.3f}"))
                    tbl.setItem(row, col + 1, _mk(f"{gps_val:.5f}"))
                else:
                    kg_val = g_val / 1000.0
                    tn_val = kg_val / 1000.0
                    tbl.setItem(row, col,     _mk(f"{kg_val:.4f}"))
                    tbl.setItem(row, col + 1, _mk(f"{tn_val:.7f}"))
                col += 2
        return tbl

    def _refresh_stand_table(self, veh: str) -> None:
        """Rellena la tabla de desglose por puesto para el GSE seleccionado."""
        tbl = self._tbl_stand
        tbl.setSortingEnabled(False)
        res = self._data_full.get(veh)
        by_stand = (res.by_stand or {}) if res else {}

        gas_h = [f"{g} [g]" for g in self._gases]
        headers = ["Puesto", "Dist [km]", "Serv. [h]"] + gas_h
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(len(by_stand))
        tbl.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)

        for row, (stand, vals) in enumerate(sorted(by_stand.items())):
            tbl.setItem(row, 0, QtWidgets.QTableWidgetItem(stand))
            tbl.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{vals.get('dist_km', 0):.3f}"))
            tbl.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{vals.get('svc_h', 0):.3f}"))
            col = 3
            for gas in self._gases:
                g_val = vals.get(gas, 0.0)
                tbl.setItem(row, col, QtWidgets.QTableWidgetItem(f"{g_val:.3f}"))
                col += 1
        tbl.setSortingEnabled(True)

    # ── Exportación CSV ──────────────────────────────────────────────────

    def _export_gse_csv(self) -> None:
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Exportar resultados por GSE", os.getcwd(), "CSV (*.csv)")
        if not fn:
            return
        try:
            with open(fn, "w", encoding="utf-8-sig") as f:
                headers = ["GSE", "dist_km", "circ_h", "serv_h", "flota_min"]
                for g in self._gases:
                    headers += [f"{g}_g", f"{g}_kg", f"{g}_Tn"]
                f.write(",".join(headers) + "\n")
                for veh in self._vehs:
                    res = self._data_full[veh]
                    row = [veh,
                           f"{res.distance_km:.4f}",
                           f"{res.circulation_time_h:.4f}",
                           f"{res.service_time_h:.4f}",
                           str(res.fleet)]
                    for gas in self._gases:
                        rec = res.gases.get(gas)
                        g_val = (rec.get("g", 0.0) or 0.0) if rec else 0.0
                        row += [f"{g_val:.4f}",
                                f"{g_val/1000:.4f}",
                                f"{g_val/1e6:.9f}"]
                    f.write(",".join(row) + "\n")
            QtWidgets.QMessageBox.information(self, "Exportado", f"CSV guardado en:\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    def _export_stand_csv(self) -> None:
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Exportar desglose por puesto", os.getcwd(), "CSV (*.csv)")
        if not fn:
            return
        try:
            with open(fn, "w", encoding="utf-8-sig") as f:
                headers = ["GSE", "Puesto", "dist_km", "serv_h"] +                           [f"{g}_g" for g in self._gases]
                f.write(",".join(headers) + "\n")
                for veh in self._vehs:
                    res = self._data_full[veh]
                    by_stand = res.by_stand or {}
                    for stand, vals in sorted(by_stand.items()):
                        row = [veh, stand,
                               f"{vals.get('dist_km', 0):.4f}",
                               f"{vals.get('svc_h', 0):.4f}"]
                        for gas in self._gases:
                            row.append(f"{vals.get(gas, 0):.4f}")
                        f.write(",".join(row) + "\n")
            QtWidgets.QMessageBox.information(self, "Exportado", f"CSV por puesto guardado en:\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    def _save_column_widths(self, table: QtWidgets.QTableWidget, key: str) -> None:
        try:
            settings = QtCore.QSettings("GSEQuant", "GSEQuant")
            widths = [str(table.columnWidth(i)) for i in range(table.columnCount())]
            settings.setValue(f"{key}/widths", ",".join(widths))
        except Exception:
            pass

    def _restore_column_widths(self, table: QtWidgets.QTableWidget, key: str) -> None:
        try:
            settings = QtCore.QSettings("GSEQuant", "GSEQuant")
            value = settings.value(f"{key}/widths")
            if not value:
                return
            parts = str(value).split(",") if not isinstance(value, (list, tuple)) else [str(v) for v in value]
            for i, w in enumerate(parts):
                if i >= table.columnCount():
                    break
                try:
                    table.setColumnWidth(i, int(w))
                except Exception:
                    continue
        except Exception:
            pass


class SyntheticCirculationDialog(QtWidgets.QDialog):
    def __init__(self, parent, results: Dict[str, CirculationResult], warnings: List[str], diagnostics: str, debug_path: str = ""):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Simulación sintética – Circulación (NO REAL)")
        self.resize(980, 520)

        layout = QtWidgets.QVBoxLayout(self)

        title = QtWidgets.QLabel(
            "Resultados de una simulación sintética de circulación basada en datos de ejemplo.\n"
            "Este escenario NO utiliza sus operaciones ni su tabla de circulación reales."
        )
        title.setStyleSheet("color:#E65100; font-weight:600;")
        title.setWordWrap(True)
        layout.addWidget(title)

        gases = CirculationCalculator._gases

        # Tabla principal: resultados en gramos y g/s
        table = QtWidgets.QTableWidget(self)
        base_headers = ["GSE", "Dist [km]", "Tiempo Circ. [h]", "Tiempo Serv. [h]", "Flota mín."]
        gas_headers = []
        for gas in gases:
            gas_headers.append(f"{gas} (g)")
            gas_headers.append(f"{gas} (g/s)")
        headers = base_headers + gas_headers
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(results))

        for row, (veh, res) in enumerate(sorted(results.items())):
            table.setItem(row, 0, QtWidgets.QTableWidgetItem(veh))
            table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{res.distance_km:.2f}"))
            table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{res.circulation_time_h:.3f}"))
            table.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{res.service_time_h:.3f}"))
            table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(res.fleet)))
            col = 5
            for gas in gases:
                rec = res.gases.get(gas)
                if rec:
                    g_val = rec.get("g", 0.0) or 0.0
                    gps_val = rec.get("gps", 0.0) or 0.0
                    table.setItem(row, col,     QtWidgets.QTableWidgetItem(f"{g_val:.2f}"))
                    table.setItem(row, col + 1, QtWidgets.QTableWidgetItem(f"{gps_val:.5f}"))
                else:
                    table.setItem(row, col,     QtWidgets.QTableWidgetItem("—"))
                    table.setItem(row, col + 1, QtWidgets.QTableWidgetItem("—"))
                col += 2
        table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        table.setToolTip(
            "Esta tabla muestra distancias, tiempos, flota mínima y emisiones por circulación\n"
            "para un conjunto sintético de operaciones y de nodos/rutas de ejemplo."
        )
        layout.addWidget(table)

        # Tabla secundaria: mismos resultados convertidos a kg y toneladas
        table_kg = QtWidgets.QTableWidget(self)
        base_headers_kg = ["GSE", "Dist [km]", "Tiempo Circ. [h]", "Tiempo Serv. [h]", "Flota mín."]
        gas_headers_kg = []
        for gas in gases:
            gas_headers_kg.append(f"{gas} (kg)")
            gas_headers_kg.append(f"{gas} (Tn)")
        headers_kg = base_headers_kg + gas_headers_kg
        table_kg.setColumnCount(len(headers_kg))
        table_kg.setHorizontalHeaderLabels(headers_kg)
        table_kg.setRowCount(len(results))

        for row, (veh, res) in enumerate(sorted(results.items())):
            table_kg.setItem(row, 0, QtWidgets.QTableWidgetItem(veh))
            table_kg.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{res.distance_km:.2f}"))
            table_kg.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{res.circulation_time_h:.3f}"))
            table_kg.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{res.service_time_h:.3f}"))
            table_kg.setItem(row, 4, QtWidgets.QTableWidgetItem(str(res.fleet)))
            col = 5
            for gas in gases:
                rec = res.gases.get(gas)
                if rec:
                    g_val = rec.get("g", 0.0) or 0.0
                    kg_val = g_val / 1000.0
                    tn_val = kg_val / 1000.0
                    table_kg.setItem(row, col,     QtWidgets.QTableWidgetItem(f"{kg_val:.3f}"))
                    table_kg.setItem(row, col + 1, QtWidgets.QTableWidgetItem(f"{tn_val:.4f}"))
                else:
                    table_kg.setItem(row, col,     QtWidgets.QTableWidgetItem("—"))
                    table_kg.setItem(row, col + 1, QtWidgets.QTableWidgetItem("—"))
                col += 2
        table_kg.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        layout.addWidget(table_kg)

        txtWarnings = QtWidgets.QTextEdit(self)
        txtWarnings.setReadOnly(True)
        warn_lines = [
            "ATENCIÓN: Escenario sintético. Úselo sólo para entender el procedimiento y depurar la lógica.",
            "No utilice estos valores para análisis reales ni reportes.",
            "",
        ]
        if diagnostics:
            warn_lines.append(f"Diagnóstico sintético: {diagnostics}")
        if warnings:
            warn_lines.append("Avisos:")
            warn_lines.extend(f" • {w}" for w in warnings)
        txtWarnings.setText("\n".join(warn_lines))
        txtWarnings.setToolTip(
            "Resumen textual de la simulación sintética, incluyendo advertencias y detalles de diagnóstico."
        )
        layout.addWidget(txtWarnings)

        btnClose = QtWidgets.QPushButton("Cerrar", self)
        btnClose.clicked.connect(self.accept)
        layout.addWidget(btnClose)


class StepByStepSimDialog(QtWidgets.QDialog):
    def __init__(self, parent, calculator: 'CirculationCalculator'):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMaximizeButtonHint)
        self.setWindowTitle("Simulación Step-by-Step · Análisis de Recorrido")
        self.resize(1100, 700)
        self.setWindowIcon(QtGui.QIcon(resource_path("gse_app_icon.png")))
        self.calculator = calculator
        self.events = []
        self.current_idx = -1
        
        layout = QtWidgets.QVBoxLayout(self)
        
        info = QtWidgets.QLabel("Utilice los controles para avanzar paso a paso por la simulación de cada vehículo.")
        info.setStyleSheet("font-weight: bold; color: #1565C0;")
        layout.addWidget(info)
        
        # Selección de vehículo
        header_lay = QtWidgets.QHBoxLayout()
        header_lay.addWidget(QtWidgets.QLabel("Vehículo:"))
        self.cmbVeh = QtWidgets.QComboBox()
        self.cmbVeh.addItems(["BUS","STA","BAG","BEL","GPU","FUE","TUG","BRE","LAV","CAT","WAT","CLE"])
        self.cmbVeh.currentIndexChanged.connect(self.reload_vehicle_events)
        header_lay.addWidget(self.cmbVeh)
        
        self.btnStart = QtWidgets.QPushButton("Iniciar / Reset")
        self.btnStart.clicked.connect(self.reload_vehicle_events)
        header_lay.addWidget(self.btnStart)
        header_lay.addStretch()
        layout.addLayout(header_lay)
        
        # Tabla de eventos
        self.table = QtWidgets.QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Tiempo", "Evento", "Desde", "Hasta", "Detalle"])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        layout.addWidget(self.table)
        
        # Controles
        ctrl_lay = QtWidgets.QHBoxLayout()
        self.btnPrev = QtWidgets.QPushButton("← Anterior")
        self.btnNext = QtWidgets.QPushButton("Siguiente →")
        self.btnPrev.clicked.connect(self.prev_step)
        self.btnNext.clicked.connect(self.next_step)
        self.lblStep = QtWidgets.QLabel("Paso: 0 / 0")
        ctrl_lay.addWidget(self.btnPrev)
        ctrl_lay.addWidget(self.lblStep)
        ctrl_lay.addWidget(self.btnNext)
        layout.addLayout(ctrl_lay)
        
        self.reload_vehicle_events()

    def reload_vehicle_events(self):
        veh = self.cmbVeh.currentText()
        # Simular solo este vehículo capturando los pasos
        events = self.calculator.simulate_vehicle_with_events(veh)
        self.events = events
        self.current_idx = -1
        self.table.setRowCount(0)
        for i, ev in enumerate(events):
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(str(ev.get('time', ''))))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(ev.get('type', '')))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(ev.get('from', '')))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(ev.get('to', '')))
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(ev.get('detail', '')))
        
        self.lblStep.setText(f"Paso: 0 / {len(events)}")
        self.parent()._js_clear_routes()

    def next_step(self):
        if self.current_idx < len(self.events) - 1:
            self.current_idx += 1
            self.apply_step(self.current_idx)

    def prev_step(self):
        if self.current_idx > 0:
            self.current_idx -= 1
            self.apply_step(self.current_idx)

    def apply_step(self, idx):
        self.table.selectRow(idx)
        self.lblStep.setText(f"Paso: {idx+1} / {len(self.events)}")
        ev = self.events[idx]
        
        # Dibujar ruta en el mapa si hay coordenadas
        path = ev.get('path_coords')
        if path:
            self.parent()._js_clear_routes()
            self.parent()._js_draw_route(path)
            # Fly to middle or end
            if path:
                mid = path[-1]
                self.parent()._js_fly_to(mid[0], mid[1], 18)
        
        self.parent().statusBar().showMessage(f"Simulación: {ev.get('detail', '')}")

@dataclass
class Node:
    id: str
    name: str
    lat: float
    lon: float
    kind: str  # 'via' | 'puesto' | 'base' | 'hub'
    has_jetbridge: bool = field(default=False)  # "manga" en puestos
    is_hub_bus: bool = field(default=False)  # Terminal ↔ avión pasajeros
    is_hub_sta: bool = field(default=False)  # Escaleras estacionadas en plataforma
    is_hub_bag: bool = field(default=False)  # BAG ↔ terminal equipaje
    is_hub_bel: bool = field(default=False)  # BEL (belt loader) zona plataforma
    is_hub_gpu: bool = field(default=False)  # GPU
    is_hub_fue: bool = field(default=False)  # Combustible
    is_rear_entry: bool = field(default=False) # Ingreso por atrás

@dataclass
class Edge:
    id: str
    n1: str
    n2: str
    length_km: float
    edge_class: str = "via"      # "via" | "connector"
    is_one_way: bool = False     # True => solo ida (n1->n2)

class GraphModel:
    def __init__(self) -> None:
        self.nodes: Dict[str, Node] = {}
        self.edges: List[Edge] = []
        self.G = nx.DiGraph()  # dirigido
        self._counter = 0
        self._edge_counter = 1
        self.default_base_id: Optional[str] = None
        self.vehicle_bases: Dict[str, str] = {}

    def next_id(self) -> str:
        """Devuelve el próximo ID de nodo como N1, N2, ...

        En lugar de depender solo de _counter, mira los nodos existentes para
        reutilizar el siguiente número libre al final de la secuencia. Ejemplo:
        si existe N15 y N16, se crea N17; si N17 se borra, el siguiente alta
        vuelve a ser N17 (no N18).
        """
        max_num = 0
        for nid in self.nodes.keys():
            if isinstance(nid, str) and nid.startswith("N"):
                try:
                    num = int(nid[1:])
                except ValueError:
                    continue
                if num > max_num:
                    max_num = num
        next_num = max_num + 1
        # Mantener _counter en sincronía por si otras partes lo usan
        self._counter = max(self._counter, next_num)
        return f"N{next_num}"

    def next_edge_id(self) -> str:
        eid = f"E{self._edge_counter:04d}"
        self._edge_counter += 1
        return eid

    def add_node(self, name: str, lat: float, lon: float, kind: str = "via",
                 node_id: Optional[str] = None, has_jetbridge: bool = False) -> Node:
        if node_id is None:
            node_id = self.next_id()
        node = Node(id=node_id, name=name, lat=lat, lon=lon, kind=kind, has_jetbridge=has_jetbridge)
        self.nodes[node.id] = node
        self.G.add_node(node.id)
        return node

    def update_node(self, node_id: str, **kwargs):
        node = self.nodes.get(node_id)
        if not node:
            return
        for k, v in kwargs.items():
            if hasattr(node, k):
                setattr(node, k, v)

    def remove_node(self, node_id: str) -> bool:
        if node_id not in self.nodes:
            return False
        del self.nodes[node_id]
        # Eliminar del grafo de networkx
        if node_id in self.G:
            self.G.remove_node(node_id)
        # Eliminar aristas asociadas
        self.edges = [e for e in self.edges if e.n1 != node_id and e.n2 != node_id]
        # Si era la base por defecto, resetear
        if self.default_base_id == node_id:
            self.default_base_id = None
        # Si era una base de vehículo, eliminar
        self.vehicle_bases = {k: v for k, v in self.vehicle_bases.items() if v != node_id}
        return True

    def add_edge(self, n1: str, n2: str, edge_class: str = "via", is_one_way: bool = False) -> Edge:
        if n1 == n2:
            return None # Evitar self-loops
        if n1 not in self.nodes or n2 not in self.nodes:
            raise ValueError("Ambos nodos deben existir para crear la arista.")
        if edge_class == "connector":
            a, b = self.nodes[n1], self.nodes[n2]
            ok = (a.kind == "puesto" and b.kind == "via") or (a.kind == "via" and b.kind == "puesto")
            if not ok:
                raise ValueError("Un conector debe unir un 'puesto' con una 'via'.")
        a, b = self.nodes[n1], self.nodes[n2]
        length = haversine_km(a.lat, a.lon, b.lat, b.lon)
        edge = Edge(id=self.next_edge_id(), n1=n1, n2=n2, length_km=length, edge_class=edge_class, is_one_way=is_one_way)
        self.edges.append(edge)
        self.G.add_edge(n1, n2, weight=length, edge_id=edge.id, edge_class=edge_class, is_one_way=is_one_way)
        if not is_one_way:
            self.G.add_edge(n2, n1, weight=length, edge_id=edge.id, edge_class=edge_class, is_one_way=is_one_way)
        return edge

    def remove_edge(self, edge_id: str) -> bool:
        idx = next((i for i, e in enumerate(self.edges) if e.id == edge_id), None)
        if idx is None:
            return False
        e = self.edges.pop(idx)
        # borrar del grafo
        try:
            self.G.remove_edge(e.n1, e.n2)
        except Exception:
            pass
        if not e.is_one_way:
            try:
                self.G.remove_edge(e.n2, e.n1)
            except Exception:
                pass
        return True

    def export_json(self, filepath: str) -> None:
        data = {
            "nodes": {nid: asdict(node) for nid, node in self.nodes.items()},
            "edges": [asdict(e) for e in self.edges],
            "meta": {"default_base_id": self.default_base_id,
                     "edge_counter": self._edge_counter,
                     "node_counter": self._counter,
                     "vehicle_bases": self.vehicle_bases},
        }
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


    def load_json(self, filepath: str) -> None:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.nodes.clear()
        self.edges.clear()
        self.G.clear()
        for nid, info in data.get("nodes", {}).items():
            node = Node(**info)
            self.nodes[nid] = node
            self.G.add_node(nid)
        for e in data.get("edges", []):
            # normalizar campos y tipos que puedan venir como cadenas desde JSON/exportes
            e.setdefault("edge_class", "via")
            # coercionar booleano correctamente
            e["is_one_way"] = coerce_bool(e.get("is_one_way", False))
            # asegurar que length_km sea float
            try:
                e["length_km"] = float(e.get("length_km", 0.0))
            except Exception:
                e["length_km"] = 0.0
            # construir objeto Edge y añadir al grafo respetando el sentido
            edge = Edge(**e)
            self.edges.append(edge)
            # añadir arista al grafo dirigido; si no es one-way, añadir la inversa
            try:
                self.G.add_edge(edge.n1, edge.n2, weight=edge.length_km, edge_id=edge.id, edge_class=edge.edge_class, is_one_way=edge.is_one_way)
                if not edge.is_one_way:
                    self.G.add_edge(edge.n2, edge.n1, weight=edge.length_km, edge_id=edge.id, edge_class=edge.edge_class, is_one_way=edge.is_one_way)
            except Exception:
                # si hay inconsistencia en los nodos, ignorar esa arista en vez de fallar
                pass
        meta = data.get("meta", {}) or {}
        self.default_base_id = meta.get("default_base_id")
        self.vehicle_bases = meta.get("vehicle_bases", {}) or {}
        # asegurarse de que counters sean enteros
        try:
            self._edge_counter = int(meta.get("edge_counter", len(self.edges)+1))
        except Exception:
            self._edge_counter = len(self.edges) + 1
        try:
            self._counter = int(meta.get("node_counter", len(self.nodes)))
        except Exception:
            self._counter = len(self.nodes)

# ------------------------------ Puente JS <-> Python ------------------------------ #

class Bridge(QObject):
    mapClicked = QtCore.pyqtSignal(float, float)  # lat, lon
    nodeClicked = QtCore.pyqtSignal(str)          # node id
    nodeRightClicked = QtCore.pyqtSignal(str)     # node id for right-click
    edgeClicked = QtCore.pyqtSignal(str)          # edge id

    @pyqtSlot(float, float)
    def onMapClick(self, lat: float, lon: float):
        self.mapClicked.emit(lat, lon)

    @pyqtSlot(str)
    def onNodeClick(self, node_id: str):
        self.nodeClicked.emit(node_id)

    @pyqtSlot(str)
    def onNodeRightClick(self, node_id: str):
        self.nodeRightClicked.emit(node_id)

    @pyqtSlot(str)
    def onEdgeClick(self, edge_id: str):
        self.edgeClicked.emit(edge_id)

# ------------------------------ HTML Leaflet embebido ------------------------------ #

LEAFLET_HTML = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GSEQuant Map</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://unpkg.com/leaflet-polylinedecorator@1.6.0/dist/leaflet.polylineDecorator.js"></script>
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<style>
  html, body, #map { height: 100%; margin: 0; }
  .label { background:white; padding:2px 4px; border-radius:4px; border:1px solid #999; }
  #place-label {
    position: absolute;
    left: 10px;
    bottom: 10px;
    padding: 4px 8px;
    background: rgba(0,0,0,0.6);
    color: #fff;
    border-radius: 4px;
    font-size: 11px;
    max-width: 60%;
    z-index: 1000;
    pointer-events: none;
  }
  #place-label.hidden { display:none; }
  #place-label span.caption { font-weight:bold; margin-right:4px; }
</style>
</head>
<body>
<div id="map"></div>
<div id="place-label" class="hidden"><span class="caption">Lugar:</span><span id="place-label-text">—</span></div>
<script>
let map = L.map('map', { zoomControl: true }).setView([0,0], 2);
L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
    maxZoom: 19, attribution: 'Tiles &copy; Esri'
}).addTo(map);

let markers = {};    // nodeId -> marker
let polylines = [];  // rutas dibujadas desde Python
let edgeLines = {};  // edgeId -> { line, n1, n2 }
let placeLabelEl = null;

new QWebChannel(qt.webChannelTransport, function(channel) {
  window.pyHandler = channel.objects.pyHandler;
  placeLabelEl = document.getElementById('place-label');
});

map.on('click', function(e){
  if (window.pyHandler && window.pyHandler.onMapClick) {
    window.pyHandler.onMapClick(e.latlng.lat, e.latlng.lng);
  }
});

function colorFor(kind, hasJet, is_hub_gpu, is_hub_fue) {
  if (kind === 'base') return 'red';
  if (kind === 'hub') {
      if (is_hub_gpu) return '#9b59b6'; // Purple
      if (is_hub_fue) return '#34495e'; // Dark Slate
      return '#e67e22'; // Orange/gold default hub
  }
  if (kind === 'puesto') return hasJet ? '#2ecc71' : 'orange';
  return 'blue'; // via
}

function tooltipHtml(name, kind, lat, lon, hasJet, is_hub_gpu, is_hub_fue) {
  let extra = '';
  if (kind === 'puesto') extra = '<br>manga: ' + (hasJet ? 'sí' : 'no');
  if (kind === 'hub') {
      if (is_hub_gpu) extra += '<br>HUB: GPU';
      if (is_hub_fue) extra += '<br>HUB: FUE';
  }
  return `<div class="label"><b>${name}</b><br>${kind}${extra}<br>${lat.toFixed(6)}, ${lon.toFixed(6)}</div>`;
}

function addNode(node) {
  const {id, name, lat, lon, kind, has_jetbridge, is_hub_gpu, is_hub_fue} = node;
  const col = colorFor(kind, has_jetbridge, is_hub_gpu, is_hub_fue);
  const marker = L.circleMarker([lat, lon], {
    radius:6,
    color: col,
    fillColor: col,
    fillOpacity: 0.9,
    bubblingMouseEvents: false
  });
  marker.addTo(map);
  marker.bindTooltip(tooltipHtml(name, kind, lat, lon, has_jetbridge, is_hub_gpu, is_hub_fue));
  marker.on('click', (e) => {
    if (e && e.originalEvent) { L.DomEvent.stop(e.originalEvent); }
    if (window.pyHandler && window.pyHandler.onNodeClick) {
      window.pyHandler.onNodeClick(id);
    }
  });
  marker.on('contextmenu', (e) => {
    if (e && e.originalEvent) { L.DomEvent.stop(e.originalEvent); }
    if (window.pyHandler && window.pyHandler.onNodeRightClick) {
      window.pyHandler.onNodeRightClick(id);
    }
  });
  markers[id] = marker;
}

function updateNode(node) {
  const {id, name, lat, lon, kind, has_jetbridge, is_hub_gpu, is_hub_fue} = node;
  let m = markers[id];
  if (!m) { addNode(node); return; }
  m.setLatLng([lat, lon]);
  const col = colorFor(kind, has_jetbridge, is_hub_gpu, is_hub_fue);
  if (m.setStyle) m.setStyle({color: col, fillColor: col});
  const tt = m.getTooltip && m.getTooltip();
  const htmlTt = tooltipHtml(name, kind, lat, lon, has_jetbridge, is_hub_gpu, is_hub_fue);
  if (tt && tt.setContent) {
    tt.setContent(htmlTt);
  } else {
    m.bindTooltip(htmlTt);
  }
  refreshEdgesFor(id);
}

function removeNode(id) {
  const m = markers[id];
  if (m) { map.removeLayer(m); delete markers[id]; }
  // limpiar aristas relacionadas
  Object.entries(edgeLines).forEach(([eid, rec]) => {
    if (rec.n1 === id || rec.n2 === id) {
      map.removeLayer(rec.line);
      if (rec.decorator) map.removeLayer(rec.decorator);
      delete edgeLines[eid];
    }
  });
}

function addEdge(edge) {
  const {id, n1, n2, edge_class, is_one_way} = edge;
  const m1 = markers[n1], m2 = markers[n2];
  if (!m1 || !m2) return;
  const latlngs = [m1.getLatLng(), m2.getLatLng()];
  let opts = {weight: 3, opacity: 0.95};
  if (edge_class === 'connector') {
    opts.color = '#888';
    opts.dashArray = '6,6';
  } else {
    opts.color = '#00BCD4';
  }
  
  const line = L.polyline(latlngs, opts).addTo(map);
  
  // Flecha decorativa si es de un solo sentido
  let decorator = null;
  // Dibujamos la flecha si la arista es "solo ida" (is_one_way). 
  // Ahora lo permitimos tanto para 'via' como para 'connector'.
  if (is_one_way) {
    decorator = L.polylineDecorator(line, {
      patterns: [
        { offset: '50%', repeat: 0, symbol: L.Symbol.arrowHead({ pixelSize: 12, polygon: false, pathOptions: { stroke: true, color: opts.color, weight: 3 } }) }
      ]
    }).addTo(map);
  }

  line.on('click', (e) => {
    if (e && e.originalEvent) { L.DomEvent.stop(e.originalEvent); }
    if (window.pyHandler && window.pyHandler.onEdgeClick) {
      window.pyHandler.onEdgeClick(id);
    }
  });
  edgeLines[id] = { line, n1, n2, decorator };
  return line;
}

function removeEdge(edgeId) {
  const rec = edgeLines[edgeId];
  if (rec) { 
    map.removeLayer(rec.line);
    if (rec.decorator) map.removeLayer(rec.decorator);
    delete edgeLines[edgeId]; 
  }
}

function refreshEdgesFor(nodeId) {
  Object.entries(edgeLines).forEach(([eid, rec]) => {
    if (rec.n1 === nodeId || rec.n2 === nodeId) {
      const m1 = markers[rec.n1], m2 = markers[rec.n2];
      if (m1 && m2) {
        const p1 = m1.getLatLng(), p2 = m2.getLatLng();
        rec.line.setLatLngs([p1, p2]);
        if (rec.decorator) {
           rec.decorator.setPaths([p1, p2]);
        }
      }
    }
  });
}

function drawRoute(coords) {
  const line = L.polyline(coords, {weight: 5}).addTo(map);
  polylines.push(line);
  map.fitBounds(line.getBounds(), {padding:[40,40]});
}

function clearRoutes() {
  polylines.forEach(p => map.removeLayer(p));
  polylines = [];
}

function flyTo(lat, lon, z=17) { map.setView([lat, lon], z); }

function setPlaceLabel(text) {
  if (!placeLabelEl) return;
  const span = document.getElementById('place-label-text');
  if (span) {
    span.textContent = text || '—';
  }
  if (text && text.trim() !== '') {
    placeLabelEl.classList.remove('hidden');
  } else {
    placeLabelEl.classList.add('hidden');
  }
}

// Exponer
window.addNode = addNode;
window.updateNode = updateNode;
window.removeNode = removeNode;
window.addEdge = addEdge;
window.removeEdge = removeEdge;
window.drawRoute = drawRoute;
window.clearRoutes = clearRoutes;
window.flyTo = flyTo;
window.setPlaceLabel = setPlaceLabel;
</script>
</body>
</html>
"""

# ------------------------------ UI Principal ------------------------------ #

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        # Título base de la aplicación (nombre corto y profesional)
        self._base_title = "GSEQuant - PRO v.01"
        self.setWindowTitle(self._base_title)
        # Arrancar maximizada
        self.setWindowState(QtCore.Qt.WindowMaximized)
        self.setMinimumSize(1024, 680)

        # Cargar iconos Logo GTA y App GSE
        self.setWindowIcon(QtGui.QIcon(resource_path("gse_app_icon.png")))
        # Logo discreto en esquina: 52px de ancho máximo
        self._logo_pixmap = QtGui.QPixmap(resource_path("gta.png")).scaledToWidth(52, QtCore.Qt.SmoothTransformation)

        self.model = GraphModel()
        self.config = ConfigManager()
        self._ops_df = None  # operaciones (Excel) cargadas
        self._last_emis_servicio = None  # últimos resultados de emisiones por servicio
        self._last_circ_results = None  # últimos resultados de emisiones por circulación
        self._circ_dataset_is_synthetic: bool = False  # True cuando la tabla de circulación es de ejemplo
        self._session_path: Optional[str] = None  # ruta del archivo de sesión XML actual (None => sin archivo)
        # Ventanas hijas (otras sesiones abiertas en paralelo)
        self._child_windows: list["MainWindow"] = []
        # Parámetros de circulación por GSE (FC_CIR y velocidad km/h)
        self._circ_params: dict[str, dict[str, float]] = {}
        self._circ_debug_enabled: bool = True  # Genera reporte sintético hasta finalizar implementación
        self._date_filter_from: Optional[datetime.date] = None
        self._date_filter_to: Optional[datetime.date] = None
        # La tabla de circulación debe comenzar vacía salvo que se cargue explícitamente
        try:
            self.config.set_dataset("circulacion", [], [])
        except Exception:
            pass
        self.view = QWebEngineView()
        self.channel = QWebChannel(self.view.page())
        self.bridge = Bridge()
        self.channel.registerObject('pyHandler', self.bridge)
        self.view.page().setWebChannel(self.channel)
        self.view.setHtml(LEAFLET_HTML)

        self.panel = self._build_side_panel()
        self.tabs_bottom = self._build_bottom_tabs()

        # Splitter principal vertical (Mapa+Panel arriba, HUD abajo)
        self._main_splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)

        # Parte superior (Mapa + Panel lateral)
        top_container = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        top_container.addWidget(self.view)
        top_container.addWidget(self.panel)
        top_container.setStretchFactor(0, 5)
        top_container.setStretchFactor(1, 0)

        self._main_splitter.addWidget(top_container)
        self._main_splitter.addWidget(self.tabs_bottom)

        # Relación inicial de tamaños para el HUD ajustable
        self._main_splitter.setStretchFactor(0, 3)
        self._main_splitter.setStretchFactor(1, 1)

        self.setCentralWidget(self._main_splitter)

        # Barra de estado: indicar archivo de sesión actual y estado
        self._status_session_label = QtWidgets.QLabel()
        self._status_session_label.setObjectName("statusSessionLabel")
        self._status_session_label.setStyleSheet("color:#546E7A;")
        sb = self.statusBar()
        sb.addPermanentWidget(self._status_session_label)
        self._update_session_title_and_status()

        # Menú de datos para abrir editores por dataset
        self._build_menu()

        # Conexiones desde el mapa (Bridge JS)
        self.bridge.mapClicked.connect(self.on_map_clicked)
        self.bridge.nodeClicked.connect(self.on_node_clicked)
        self.bridge.nodeRightClicked.connect(self.delete_node_by_id)
        self.bridge.edgeClicked.connect(self.on_edge_clicked)

        # Estado del modo "conectar nodos"
        self.edge_first: Optional[str] = None
        self.edge_mode_active: bool = False

    def _build_side_panel(self) -> QtWidgets.QWidget:
        # Contenedor exterior con scroll
        outer = QtWidgets.QWidget()
        outer.setMinimumWidth(300)
        outer.setMaximumWidth(360)
        outer.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        inner = QtWidgets.QWidget()
        inner.setStyleSheet("""
            QWidget { background-color: #F4F6F8; }
            QPushButton {
                border: 1px solid #B0BEC5; border-radius: 3px;
                padding: 4px 8px; background: #ECEFF1;
                font-size: 11px; color: #263238;
            }
            QPushButton:hover   { background: #CFD8DC; }
            QPushButton:pressed { background: #B0BEC5; }
            QPushButton:checked { background: #FFF176; border-color: #F9A825; font-weight:bold; }
            QLineEdit, QComboBox, QDoubleSpinBox {
                border: 1px solid #CFD8DC; border-radius: 3px;
                padding: 3px 6px; background: #FFFFFF; font-size: 11px;
            }
            QCheckBox { font-size: 11px; color: #37474F; }
        """)

        vbox = QtWidgets.QVBoxLayout(inner)
        vbox.setContentsMargins(8, 0, 8, 12)
        vbox.setSpacing(3)

        # helpers
        def sec(text):
            lbl = QtWidgets.QLabel(text.upper())
            lbl.setStyleSheet(
                "color:#FFF; background:#37474F; padding:4px 8px;"
                "font-weight:700; font-size:10px; letter-spacing:1px;"
            )
            return lbl

        def lrow(lb_text, widget):
            h = QtWidgets.QHBoxLayout()
            h.setContentsMargins(0,0,0,0); h.setSpacing(6)
            lb = QtWidgets.QLabel(lb_text)
            lb.setStyleSheet("font-size:11px;color:#546E7A;min-width:50px;")
            h.addWidget(lb); h.addWidget(widget,1)
            return h

        # LOGO BAR
        logo_bar = QtWidgets.QWidget()
        logo_bar.setStyleSheet("background:#ECEFF1;")
        ll = QtWidgets.QHBoxLayout(logo_bar)
        ll.setContentsMargins(6,5,6,5); ll.setSpacing(7)
        lblLogo = QtWidgets.QLabel()
        lpx = QtGui.QPixmap(resource_path("gta.png")).scaledToHeight(32, QtCore.Qt.SmoothTransformation)
        lblLogo.setPixmap(lpx)
        t1 = QtWidgets.QLabel("GSEQuant PRO")
        t1.setStyleSheet("color:#263238;font-size:12px;font-weight:700;font-family:'Segoe UI';"
                         "background:transparent;letter-spacing:0.5px;")
        t2 = QtWidgets.QLabel("Ground Support Emissions")
        t2.setStyleSheet("color:#607D8B;font-size:8px;font-family:'Segoe UI';background:transparent;")
        text_col = QtWidgets.QVBoxLayout()
        text_col.setContentsMargins(0,0,0,0); text_col.setSpacing(1)
        text_col.addWidget(t1); text_col.addWidget(t2)
        ll.addWidget(lblLogo)
        ll.addStretch(1)        # empuja el texto al extremo derecho
        ll.addLayout(text_col)
        vbox.addWidget(logo_bar)

        # AEROPUERTO
        vbox.addSpacing(4)
        vbox.addWidget(sec("Aeropuerto"))
        self.txtSearch = QtWidgets.QLineEdit()
        self.txtSearch.setPlaceholderText("Nombre, IATA o coordenadas…")
        self.txtSearch.setToolTip(""" 🔍 Formatos aceptados:

📍 Nombre de lugar o aeropuerto:
   • Aeroparque
   • Ezeiza
   • Buenos Aires

🌐 Código IATA:
   • AEP  (Aeroparque Jorge Newbery)
   • EZE  (Aeropuerto Ministro Pistarini)
   • COR  (Córdoba)
   • MDZ  (Mendoza)
   • BHI  (Bahía Blanca)

📐 Coordenadas decimales:
   • -34.5597 -58.4156
   • -34.5597, -58.4156

📐 Coordenadas DMS:
   • 34°33'35"S 58°24'56"W

📐 DMS compacto:
   • 343535S 0582456W

↵ Presioná Enter o el botón Buscar/Ir """)
        self.btnSearch = QtWidgets.QPushButton("Buscar / Ir")
        self.btnSearch.setStyleSheet(
            "QPushButton{background:#1565C0;color:white;border:none;border-radius:3px;"
            "padding:4px 10px;font-size:11px;font-weight:600;}"
            "QPushButton:hover{background:#1976D2;}")
        self.btnSearch.clicked.connect(self.on_search)
        self.txtSearch.returnPressed.connect(self.on_search)
        sr = QtWidgets.QHBoxLayout(); sr.setSpacing(4)
        sr.addWidget(self.txtSearch,1); sr.addWidget(self.btnSearch)
        vbox.addLayout(sr)

        # NUEVO NODO
        vbox.addSpacing(6); vbox.addWidget(sec("Nuevo Nodo"))
        self.txtName = QtWidgets.QLineEdit()
        self.txtName.setPlaceholderText("Nombre del nodo…")
        vbox.addLayout(lrow("Nombre:", self.txtName))
        self.cmbKind = QtWidgets.QComboBox()
        self.cmbKind.addItems(["via","puesto","base","hub"])
        vbox.addLayout(lrow("Tipo:", self.cmbKind))
        self.chkManga = QtWidgets.QCheckBox("Puesto con manga")
        vbox.addWidget(self.chkManga)
        self.dsbLat = QtWidgets.QDoubleSpinBox()
        self.dsbLat.setRange(-90,90); self.dsbLat.setDecimals(6)
        self.dsbLon = QtWidgets.QDoubleSpinBox()
        self.dsbLon.setRange(-180,180); self.dsbLon.setDecimals(6)
        vbox.addLayout(lrow("Lat:", self.dsbLat))
        vbox.addLayout(lrow("Lon:", self.dsbLon))
        self.txtCoordRaw = QtWidgets.QLineEdit()
        self.txtCoordRaw.setPlaceholderText("DMS/UTM (ej: 314239S 0604841W)")
        self.btnCoordConvert = QtWidgets.QPushButton("→ Dec")
        self.btnCoordConvert.setFixedWidth(52)
        self.btnCoordConvert.clicked.connect(self.apply_raw_coords)
        cr = QtWidgets.QHBoxLayout(); cr.setSpacing(4)
        cr.addWidget(self.txtCoordRaw,1); cr.addWidget(self.btnCoordConvert)
        vbox.addLayout(cr)
        self.btnAddNode = QtWidgets.QPushButton("⊕  Agregar Nodo en Mapa")
        self.btnAddNode.setCheckable(True)
        self.btnAddNode.setStyleSheet(
            "QPushButton{background:#2E7D32;color:white;border:none;border-radius:4px;"
            "padding:7px;font-weight:700;font-size:12px;}"
            "QPushButton:hover{background:#388E3C;}"
            "QPushButton:checked{background:#F9A825;color:#263238;border:2px solid #F57F17;}")
        self.btnAddNode.clicked.connect(self.on_toggle_add_mode)
        vbox.addWidget(self.btnAddNode)

        # RUTAS
        vbox.addSpacing(6); vbox.addWidget(sec("Rutas"))
        self.cmbEdgeClass = QtWidgets.QComboBox()
        self.cmbEdgeClass.addItems(["Vía (calle)","Conector Puesto\u2194Vía"])
        self.cmbEdgeDir = QtWidgets.QComboBox()
        self.cmbEdgeDir.addItems(["Doble (ida y vuelta)","Solo ida (según orden)"])
        vbox.addLayout(lrow("Tipo:", self.cmbEdgeClass))
        vbox.addLayout(lrow("Sentido:", self.cmbEdgeDir))
        self.btnStartEdge  = QtWidgets.QPushButton("① Seleccionar Origen")
        self.btnFinishEdge = QtWidgets.QPushButton("② Seleccionar Destino")
        self.btnSeqEdge    = QtWidgets.QPushButton("⟳ Modo Secuencial")
        self.btnSeqEdge.setCheckable(True)
        self.btnSeqEdge.setStyleSheet(
            "QPushButton{border:1px solid #B0BEC5;border-radius:3px;padding:4px 8px;"
            "background:#ECEFF1;font-size:11px;color:#263238;}"
            "QPushButton:hover{background:#CFD8DC;}"
            "QPushButton:checked{background:#B3E5FC;border-color:#039BE5;font-weight:bold;}")
        self.btnStartEdge.clicked.connect(self.start_edge_mode)
        self.btnFinishEdge.clicked.connect(self.finish_edge_mode)
        self.btnSeqEdge.clicked.connect(self.on_toggle_sequential_routing_mode)
        rt = QtWidgets.QHBoxLayout(); rt.setSpacing(4)
        rt.addWidget(self.btnStartEdge,1); rt.addWidget(self.btnFinishEdge,1)
        vbox.addLayout(rt)
        vbox.addWidget(self.btnSeqEdge)

        # ARCHIVO
        vbox.addSpacing(6); vbox.addWidget(sec("Archivo"))
        self.btnExport = QtWidgets.QPushButton("↑ Exportar JSON…")
        self.btnImport = QtWidgets.QPushButton("↓ Importar JSON…")
        self.btnExport.clicked.connect(self.export_json)
        self.btnImport.clicked.connect(self.import_json)
        fl = QtWidgets.QHBoxLayout(); fl.setSpacing(4)
        fl.addWidget(self.btnExport,1); fl.addWidget(self.btnImport,1)
        vbox.addLayout(fl)

        # VERIFICADOR
        vbox.addSpacing(6); vbox.addWidget(sec("Verificador de Rutas"))
        self.cmbCalcFrom = QtWidgets.QComboBox()
        self.cmbCalcTo   = QtWidgets.QComboBox()
        vbox.addLayout(lrow("Desde:", self.cmbCalcFrom))
        vbox.addLayout(lrow("Hasta:", self.cmbCalcTo))
        self.chkDrawPath = QtWidgets.QCheckBox("Dibujar camino en mapa")
        vbox.addWidget(self.chkDrawPath)
        self.btnCalc = QtWidgets.QPushButton("Calcular distancia mínima")
        self.btnCalc.clicked.connect(self.calc_min_distance)
        vbox.addWidget(self.btnCalc)

        # BASES
        vbox.addSpacing(6); vbox.addWidget(sec("Bases"))
        self.btnSetDefaultBase = QtWidgets.QPushButton("Fijar base por defecto")
        self.btnCenterBase     = QtWidgets.QPushButton("Centrar en base")
        self.btnConfigVehBases = QtWidgets.QPushButton("Config. bases por vehículo…")
        self.btnSetDefaultBase.clicked.connect(self.set_default_base_from_table)
        self.btnCenterBase.clicked.connect(self.center_on_default_base)
        bl = QtWidgets.QHBoxLayout(); bl.setSpacing(4)
        bl.addWidget(self.btnSetDefaultBase,1); bl.addWidget(self.btnCenterBase,1)
        vbox.addLayout(bl)
        vbox.addWidget(self.btnConfigVehBases)

        vbox.addStretch(1)
        scroll.setWidget(inner)
        ol = QtWidgets.QVBoxLayout(outer)
        ol.setContentsMargins(0,0,0,0); ol.setSpacing(0)
        ol.addWidget(scroll)
        return outer

    def _build_bottom_tabs(self) -> QtWidgets.QTabWidget:
        tabs = QtWidgets.QTabWidget()

        # ---- NODOS ----
        self.tblNodes = QtWidgets.QTableWidget(0, 7)
        self.tblNodes.setHorizontalHeaderLabels(["ID", "Nombre", "Tipo", "Lat", "Lon", "Manga", "Hub"])
        self.tblNodes.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.tblNodes.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.SelectedClicked)
        self.tblNodes.itemChanged.connect(self.on_table_item_changed)
        self.tblNodes.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tblNodes.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.tblNodes.customContextMenuRequested.connect(self._nodes_context_menu)
        # Centrar mapa al hacer clic o doble clic en un nodo de la tabla
        self.tblNodes.itemClicked.connect(self._center_on_node_from_row)
        self.tblNodes.itemDoubleClicked.connect(self._center_on_node_from_row)

        btns_nodes = QtWidgets.QHBoxLayout()
        btns_nodes.setContentsMargins(0, 4, 0, 0)
        btns_nodes.setSpacing(16)

        base_btn_style = (
            "QPushButton {padding:6px 14px; border:1px solid #C7CCD1; border-radius:4px; background-color:#FAFBFC; color:#222;}"
            "QPushButton:hover {background-color:#F3F4F6;}"
            "QPushButton:pressed {background-color:#E9EBEF;}"
            "QPushButton:focus {border:2px solid #90A4AE;}"
            "QPushButton:disabled {background-color:#FFFFFF; color:#B0BEC5; border:1px dashed #CFD8DC;}"
        )
        primary_green_style = (
            "QPushButton {padding:6px 16px; border:1px solid #2E7D32; border-radius:4px; color:white;"
            " background-color:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #43A047, stop:1 #2E7D32);}"
            "QPushButton:hover {background-color:#3FA244;}"
            "QPushButton:pressed {background-color:#2E7D32;}"
            "QPushButton:disabled {background-color:#E8F5E9; color:#A5D6A7; border:1px solid #C8E6C9;}"
        )
        total_blue_style = (
            "QPushButton {padding:6px 16px; border:1px solid #1565C0; border-radius:4px; color:white;"
            " background-color:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #42A5F5, stop:1 #1565C0); font-weight:600;}"
            "QPushButton:hover {background-color:#1E88E5;}"
            "QPushButton:pressed {background-color:#1565C0;}"
            "QPushButton:disabled {background-color:#E3F2FD; color:#90CAF9; border:1px solid #BBDEFB;}"
        )

        # Botones
        self.btnExportNodesCSV = QtWidgets.QPushButton("Exportar nodos a CSV…"); self.btnExportNodesCSV.setStyleSheet(base_btn_style)
        self.btnCopyNodes = QtWidgets.QPushButton("Copiar nodos (CSV)"); self.btnCopyNodes.setStyleSheet(base_btn_style)
        self.btnDeleteNode = QtWidgets.QPushButton("Eliminar fila seleccionada"); self.btnDeleteNode.setStyleSheet(base_btn_style)
        self.btnCalcEmisServicio = QtWidgets.QPushButton("Calcular emisiones por servicio"); self.btnCalcEmisServicio.setStyleSheet(primary_green_style)
        self.btnCalcEmisServicio.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DialogApplyButton))
        self.btnCalcEmisServicio.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnCalcEmisServicio.setToolTip(
            "Calcula las emisiones en servicio para cada GSE usando el histórico de operaciones y la configuración de puestos/hubs.\n"
            "Se habilita cuando: (1) se cargó una planilla de operaciones (Excel) compatible y (2) existe una tabla/distribución de puestos (circulación) real,\n"
            "ya sea cargada desde Excel o generada/guardada desde el grafo."
        )
        # Botones relacionados con circulación
        self.btnCirculacionLoadExcel = QtWidgets.QPushButton("Cargar planilla circulación (Excel)…")
        self.btnCirculacionLoadExcel.setStyleSheet(base_btn_style)
        self.btnCirculacionLoadExcel.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DirOpenIcon))
        self.btnCirculacionLoadExcel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        # Botón discreto (tipo tool button) para captura desde grafo, al costado del Excel
        self.btnCirculacionFromGraph = QtWidgets.QToolButton()
        self.btnCirculacionFromGraph.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_FileDialogInfoView))
        self.btnCirculacionFromGraph.setToolTip("Capturar/editar tabla de circulación desde el grafo")
        self.btnCirculacionFromGraph.setAutoRaise(True)
        self.btnCirculacionFromGraph.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnCirculacionFromGraph.setFixedSize(22, 22)
        self.btnCirculacionFromGraph.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        circ_menu = QtWidgets.QMenu(self)
        act_capture = circ_menu.addAction("Generar tabla desde grafo")
        act_capture.triggered.connect(self._capture_circulation_from_graph)
        act_edit = circ_menu.addAction("Editar tabla de circulación…")
        act_edit.triggered.connect(self._open_circulation_editor)
        self.btnCirculacionFromGraph.setMenu(circ_menu)
        self.btnCalcEmisCirculacion = QtWidgets.QPushButton("Calcular emisiones por circulación")
        self.btnCalcEmisCirculacion.setStyleSheet(primary_green_style)
        self.btnCalcEmisCirculacion.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_BrowserReload))
        self.btnCalcEmisCirculacion.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnCalcEmisCirculacion.setToolTip(
            "Calcula las emisiones por circulación usando la red de nodos y la tabla de circulación.\n"
            "Se habilita cuando se carga una planilla de operaciones (Excel) y existe una tabla de circulación\n"
            "(cargada desde Excel o generada/guardada desde el grafo)."
        )

        # Simulación sintética de circulación (siempre disponible, separada de los cálculos reales)
        self.btnSyntheticCirc = QtWidgets.QPushButton("Simulación sintética de circulación…")
        self.btnSyntheticCirc.setStyleSheet(
            "QPushButton {padding:6px 16px; border:1px solid #FF9800; border-radius:4px; color:#FF6F00;"
            " background-color:#FFF3E0; font-weight:600;}"
            "QPushButton:hover {background-color:#FFE0B2;}"
            "QPushButton:pressed {background-color:#FFCC80;}"
        )
        self.btnSyntheticCirc.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_MessageBoxWarning))
        self.btnSyntheticCirc.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnSyntheticCirc.setToolTip(
            "Ejecuta una prueba sintética de circulación con datos de ejemplo.\n"
            "No utiliza las operaciones reales ni la tabla de circulación real.\n"
            "Úselo sólo para entender el funcionamiento del modelo."
        )

        self.btnCalcEmisTotal = QtWidgets.QPushButton("Calcular emisiones totales (servicio + circulación)")
        self.btnCalcEmisTotal.setStyleSheet(total_blue_style)
        self.btnCalcEmisTotal.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DialogApplyButton))
        self.btnCalcEmisTotal.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnCalcEmisTotal.setToolTip(
            "Calcula el total de emisiones combinando servicio + circulación para cada GSE.\n"
            "Se habilita cuando se cargan operaciones (Excel) y se dispone de una tabla de circulación real."
        )

        self.btnDateFilter = QtWidgets.QPushButton("Filtrar por fechas…")
        self.btnDateFilter.setStyleSheet(base_btn_style)
        self.btnDateFilter.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_FileDialogToParent))
        self.btnDateFilter.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnDateFilter.setToolTip(
            "Limita los cálculos de emisiones al rango de fechas seleccionado en la tabla de operaciones."
        )

        # Tamaño y políticas (auto-ajustable: el texto siempre debe leerse completo)
        for b in (self.btnExportNodesCSV, self.btnCopyNodes, self.btnDeleteNode, self.btnCalcEmisServicio,
                  self.btnCirculacionLoadExcel, self.btnCalcEmisCirculacion, self.btnCalcEmisTotal, self.btnSyntheticCirc):
            b.setMinimumHeight(28)
            b.setIconSize(QtCore.QSize(14, 14))
            b.setFont(QtGui.QFont("Segoe UI", 8))
            b.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
            b.setMinimumWidth(b.sizeHint().width())  # nunca cortar el texto

        # Botón para cargar Excel (neutro) y habilitar cálculo sólo cuando haya archivo válido
        self.btnLoadOpsExcel = QtWidgets.QPushButton("Cargar operaciones (Excel)…")
        self.btnLoadOpsExcel.setStyleSheet(base_btn_style)
        self.btnLoadOpsExcel.setMinimumHeight(30)
        self.btnLoadOpsExcel.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DirOpenIcon))
        self.btnLoadOpsExcel.setIconSize(QtCore.QSize(16,16))
        self.btnLoadOpsExcel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnLoadOpsExcel.setToolTip(
            "Estructura sugerida (fila 1):\n"
            "A: DIA\nB: TIPOVUELO\nC: Tipo de operación\nD: Aerolinea\nE: Aeronave\n"
            "F: Numero de vuelo\nG: Puerta asignada\nH: Hora IN GATE\nI: Hora OUT Gate\nJ: TIPO  SER\n\n"
            "Se leerá la primera hoja del archivo."
        )
        # El cálculo debe esperar a que se cargue un Excel válido
        self.btnCalcEmisServicio.setEnabled(False)

        # Botón para cargar Excel (neutro) y habilitar cálculo/preview sólo cuando haya archivo válido
        self.btnLoadOpsExcel = QtWidgets.QPushButton("Cargar operaciones (Excel)…")
        self.btnLoadOpsExcel.setStyleSheet(base_btn_style)
        self.btnLoadOpsExcel.setMinimumHeight(30)
        self.btnLoadOpsExcel.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DirOpenIcon))
        self.btnLoadOpsExcel.setIconSize(QtCore.QSize(16,16))
        self.btnLoadOpsExcel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        # Botón de previsualización opcional
        self.btnPreviewOps = QtWidgets.QPushButton("Previsualizar operaciones…")
        self.btnPreviewOps.setStyleSheet(base_btn_style)
        self.btnPreviewOps.setMinimumHeight(30)
        self.btnPreviewOps.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_FileDialogInfoView))
        self.btnPreviewOps.setIconSize(QtCore.QSize(16,16))
        self.btnPreviewOps.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        # El cálculo/preview deben esperar a que se cargue un Excel válido
        self.btnCalcEmisServicio.setEnabled(False)
        self.btnCalcEmisCirculacion.setEnabled(False)
        self.btnCalcEmisTotal.setEnabled(False)
        self.btnPreviewOps.setEnabled(False)

        # Encabezados
        def header(label: str) -> QtWidgets.QLabel:
            h = QtWidgets.QLabel(label)
            h.setStyleSheet("color:#607D8B; font-weight:600;")
            return h

        # Grupos (ancho contenido controlado)
        def column_group(title: str, widgets: list, maxw: int = 0) -> QtWidgets.QWidget:
            w = QtWidgets.QWidget()
            lay = QtWidgets.QVBoxLayout(w)
            lay.setContentsMargins(0,0,0,0)
            lay.setSpacing(6)
            lay.addWidget(header(title))
            for ww in widgets:
                lay.addWidget(ww)
            lay.addStretch(1)
            if maxw > 0:
                w.setMaximumWidth(maxw)
            return w

        grpDatos = column_group("Datos de nodos", [self.btnExportNodesCSV, self.btnCopyNodes])

        self.btnImportExcelConfig = QtWidgets.QPushButton("Importar Config (Excel)…")
        self.btnImportExcelConfig.setToolTip("Carga EF, GSE x Aeronave o Coeficientes desde un Excel externo.")
        self.btnImportExcelConfig.clicked.connect(self._import_config_from_excel)

        grpGestion = column_group("Gestión", [self.btnDeleteNode, self.btnImportExcelConfig])
        # Fila compacta: botón cargar + icono de información (preview)
        ops_row = QtWidgets.QWidget()
        ops_row_lay = QtWidgets.QHBoxLayout(ops_row)
        ops_row_lay.setContentsMargins(0,0,0,0)
        ops_row_lay.setSpacing(6)
        ops_row_lay.addWidget(self.btnLoadOpsExcel)
        self.btnPreviewOps = QtWidgets.QToolButton()
        self.btnPreviewOps.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_MessageBoxInformation))
        self.btnPreviewOps.setToolTip(
            "Previsualizar operaciones (muestra 10 primeras filas).\n\n"
            "Estructura sugerida (fila 1):\n"
            "A: DIA\nB: TIPOVUELO\nC: Tipo de operación\nD: Aerolinea\nE: Aeronave\n"
            "F: Numero de vuelo\nG: Puerta asignada\nH: Hora IN GATE\nI: Hora OUT Gate\nJ: TIPO  SER"
        )
        self.btnPreviewOps.setAutoRaise(True)
        self.btnPreviewOps.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnPreviewOps.setEnabled(False)
        self.btnPreviewOps.setFixedSize(22, 22)
        # Botón engranaje: parámetros de simulación
        self.btnSimParams = QtWidgets.QToolButton()
        self.btnSimParams.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))
        self.btnSimParams.setToolTip("Parámetros de simulación (por vehículo)")
        self.btnSimParams.setAutoRaise(True)
        self.btnSimParams.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        # Columna izquierda: operaciones (arriba: cargar+iconos, abajo: calcular servicio)
        col_ops = QtWidgets.QWidget()
        col_ops_lay = QtWidgets.QVBoxLayout(col_ops)
        col_ops_lay.setContentsMargins(0,0,0,0)
        col_ops_lay.setSpacing(4)

        row_ops_top = QtWidgets.QWidget()
        row_ops_top_lay = QtWidgets.QHBoxLayout(row_ops_top)
        row_ops_top_lay.setContentsMargins(0,0,0,0)
        row_ops_top_lay.setSpacing(4)
        row_ops_top_lay.addWidget(self.btnLoadOpsExcel)
        row_ops_top_lay.addWidget(self.btnPreviewOps)
        row_ops_top_lay.addWidget(self.btnSimParams)
        
        # Filtro de Aeronaves
        self.cmbFilterAircraft = QtWidgets.QComboBox()
        self.cmbFilterAircraft.addItem("Todas las aeronaves")
        self.cmbFilterAircraft.setToolTip("Permite aislar el cálculo de emisiones a una aeronave específica.")
        self.cmbFilterAircraft.setMinimumHeight(28)
        self.cmbFilterAircraft.setMinimumWidth(160)
        self.cmbFilterAircraft.setStyleSheet("QComboBox { border: 1px solid #C7CCD1; border-radius: 4px; padding-left: 6px; }")
        
        row_ops_top_lay.addWidget(self.cmbFilterAircraft)
        row_ops_top_lay.addStretch(1)

        row_ops_bottom = QtWidgets.QWidget()
        row_ops_bottom_lay = QtWidgets.QHBoxLayout(row_ops_bottom)
        row_ops_bottom_lay.setContentsMargins(0,0,0,0)
        row_ops_bottom_lay.setSpacing(4)
        # Botón de cálculo de servicio ocupa todo el ancho disponible de la columna
        row_ops_bottom_lay.addWidget(self.btnCalcEmisServicio)

        col_ops_lay.addWidget(row_ops_top)
        col_ops_lay.addWidget(row_ops_bottom)

        # Columna derecha: circulación (arriba: cargar+icono, abajo: calcular circulación)
        col_circ = QtWidgets.QWidget()
        col_circ_lay = QtWidgets.QVBoxLayout(col_circ)
        col_circ_lay.setContentsMargins(0,0,0,0)
        col_circ_lay.setSpacing(4)

        row_circ_top = QtWidgets.QWidget()
        row_circ_top_lay = QtWidgets.QHBoxLayout(row_circ_top)
        row_circ_top_lay.setContentsMargins(0,0,0,0)
        row_circ_top_lay.setSpacing(4)
        row_circ_top_lay.addWidget(self.btnCirculacionLoadExcel)
        row_circ_top_lay.addWidget(self.btnCirculacionFromGraph)
        # Botón para parámetros de circulación (FC_CIR y velocidad)
        self.btnCircParams = QtWidgets.QToolButton()
        self.btnCircParams.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))
        self.btnCircParams.setToolTip("Parámetros de circulación por GSE (FC_CIR y velocidad)")
        self.btnCircParams.setAutoRaise(True)
        self.btnCircParams.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btnCircParams.setFixedSize(22, 22)
        row_circ_top_lay.addWidget(self.btnCircParams)
        row_circ_top_lay.addStretch(1)

        row_circ_bottom = QtWidgets.QWidget()
        row_circ_bottom_lay = QtWidgets.QHBoxLayout(row_circ_bottom)
        row_circ_bottom_lay.setContentsMargins(0,0,0,0)
        row_circ_bottom_lay.setSpacing(4)
        row_circ_bottom_lay.addWidget(self.btnCalcEmisCirculacion)
        row_circ_bottom_lay.addStretch(1)

        col_circ_lay.addWidget(row_circ_top)
        col_circ_lay.addWidget(row_circ_bottom)

        # Contenedor de ambas columnas con separación entre ellas
        calc_cols = QtWidgets.QWidget()
        calc_cols_lay = QtWidgets.QHBoxLayout(calc_cols)
        calc_cols_lay.setContentsMargins(0,0,0,0)
        calc_cols_lay.setSpacing(16)
        calc_cols_lay.addWidget(col_ops)
        calc_cols_lay.addWidget(col_circ)

        grpCalculo = column_group("Cálculo (servicio y circulación)", [calc_cols])  # sin maxw: auto-ajusta

        # Grupo específico para el total integral
        grpTotal = QtWidgets.QWidget()
        grpTotal_lay = QtWidgets.QVBoxLayout(grpTotal)
        grpTotal_lay.setContentsMargins(0, 0, 0, 0)
        grpTotal_lay.setSpacing(6)
        grpTotal_lay.addWidget(header("Emisiones totales (integral)"))
        grpTotal_lay.addWidget(self.btnCalcEmisTotal)
        grpTotal_lay.addWidget(self.btnDateFilter)
        grpTotal_lay.addStretch(1)
        # grpTotal sin límite fijo de ancho

        grpSynthetic = column_group("Simulación sintética (NO REAL)", [self.btnSyntheticCirc])

        # Separadores verticales (general y específico para el integral)
        def vline():
            line = QtWidgets.QFrame()
            line.setFrameShape(QtWidgets.QFrame.VLine)
            line.setFrameShadow(QtWidgets.QFrame.Sunken)
            line.setStyleSheet("color:#CFD8DC;")
            return line

        def vline_integral():
            line = QtWidgets.QFrame()
            # Usar sólo un borde izquierdo punteado para marcar el integral
            line.setFrameShape(QtWidgets.QFrame.NoFrame)
            line.setStyleSheet("border:0; border-left:1px dashed #1565C0;")
            line.setFixedWidth(4)
            line.setMinimumHeight(30)
            return line

        # Alinear los grupos: datos, gestión, cálculo, separador punteado (integral) y módulos a la derecha
        btns_nodes.addWidget(grpDatos)
        btns_nodes.addWidget(vline())
        btns_nodes.addWidget(grpGestion)
        btns_nodes.addWidget(vline())
        btns_nodes.addWidget(grpCalculo)
        btns_nodes.addWidget(vline_integral())
        btns_nodes.addWidget(grpTotal)
        btns_nodes.addWidget(vline())
        btns_nodes.addWidget(grpSynthetic)
        btns_nodes.addStretch(1)

        # Conexiones
        self.btnExportNodesCSV.clicked.connect(self.export_nodes_csv)
        self.btnCopyNodes.clicked.connect(self.copy_nodes_to_clipboard)
        self.btnDeleteNode.clicked.connect(self.delete_selected_node_row)
        self.btnLoadOpsExcel.clicked.connect(self._load_operations_excel)
        self.btnPreviewOps.clicked.connect(self._preview_operations_excel)
        self.btnSimParams.clicked.connect(self._open_sim_params)
        self.btnCalcEmisServicio.clicked.connect(self.open_emissions_dialog)
        self.btnCirculacionLoadExcel.clicked.connect(self._open_circulation_excel)
        # Menú del botón incluye captura y edición
        self.btnCalcEmisCirculacion.clicked.connect(self.open_circulation_emissions_dialog)
        self.btnCalcEmisTotal.clicked.connect(self.open_total_emissions_dialog)
        self.btnCircParams.clicked.connect(self._open_circ_params)
        self.btnSyntheticCirc.clicked.connect(self.open_synthetic_circulation_dialog)
        self.btnStepSim = QtWidgets.QPushButton("Simulación Step-by-Step…")
        self.btnStepSim.setStyleSheet("background-color: #e1f5fe; border: 1px solid #01579b;")
        self.btnStepSim.clicked.connect(self.open_step_by_step_sim)
        grpSynthetic_lay = grpSynthetic.layout()
        if grpSynthetic_lay: grpSynthetic_lay.addWidget(self.btnStepSim)

        self.btnConfigVehBases.clicked.connect(self._open_vehicle_bases_dialog)
        self.btnDateFilter.clicked.connect(self._open_date_filter_dialog)

        # Fin footer
        wrap_nodes = QtWidgets.QWidget(); v1 = QtWidgets.QVBoxLayout(wrap_nodes)
        v1.addWidget(self.tblNodes); v1.addLayout(btns_nodes)

        # ---- RUTAS ----
        self.tblEdges = QtWidgets.QTableWidget(0, 6)
        self.tblEdges.setHorizontalHeaderLabels(["ID", "Desde (Nodo)", "Hasta (Nodo)", "Tipo", "Sentido", "Long_km"])
        self.tblEdges.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.tblEdges.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tblEdges.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tblEdges.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.tblEdges.customContextMenuRequested.connect(self._edges_context_menu)

        btns_edges = QtWidgets.QHBoxLayout()
        self.btnExportEdgesCSV = QtWidgets.QPushButton("Exportar rutas a CSV…")
        self.btnDeleteEdge = QtWidgets.QPushButton("Eliminar ruta seleccionada")
        self.btnExportEdgesCSV.clicked.connect(self.export_edges_csv)
        self.btnDeleteEdge.clicked.connect(self.delete_selected_edge_row)
        btns_edges.addWidget(self.btnExportEdgesCSV); btns_edges.addWidget(self.btnDeleteEdge)

        wrap_edges = QtWidgets.QWidget(); v2 = QtWidgets.QVBoxLayout(wrap_edges)
        v2.addWidget(self.tblEdges); v2.addLayout(btns_edges)

        tabs.addTab(wrap_nodes, "Nodos")
        tabs.addTab(wrap_edges, "Rutas")
        return tabs

    # ---------------- Estado de sesión / título de ventana ---------------- #

    def _update_session_title_and_status(self) -> None:
        """Actualiza el título de la ventana y la barra de estado según la sesión.

        - Si no hay archivo asociado (self._session_path is None), indica que es una
          sesión nueva / sin guardar.
        - Si hay archivo, muestra el nombre base del archivo en el título y en la barra.
        """
        base = getattr(self, "_base_title", "GSEQuant - PRO v.01")
        path = getattr(self, "_session_path", None)
        if path:
            name = os.path.basename(path)
            title = f"{base} — {name}"
            status = f"Sesión: {name} (archivo existente en disco)"
        else:
            title = f"{base} — Sesión sin archivo"
            status = "Sesión sin archivo (no guardada aún)"
        self.setWindowTitle(title)
        if hasattr(self, "_status_session_label") and self._status_session_label is not None:
            self._status_session_label.setText(status)

    # ---------------- Menú y Diálogos de Datos ---------------- #
    def _build_menu(self):
        menubar = self.menuBar()

        # Menú de sesión (archivos XML con todo el estado)
        m_sesion = menubar.addMenu("Sesión")
        act_new = m_sesion.addAction("Nueva sesión en ventana")
        act_open = m_sesion.addAction("Abrir sesión…")
        act_save = m_sesion.addAction("Guardar sesión…")
        act_save_as = m_sesion.addAction("Guardar sesión como…")
        act_new.triggered.connect(self._menu_new_session_window)
        act_open.triggered.connect(self._menu_open_session)
        act_save.triggered.connect(self._menu_save_session)
        act_save_as.triggered.connect(self._menu_save_session_as)

        # Menú de datos para abrir editores por dataset
        menu = menubar.addMenu("Datos")
        for key in self.config.list_datasets():
            act = menu.addAction(f"Editar {key}…")
            act.triggered.connect(lambda checked=False, k=key: self._open_dataset_editor(k))
        # (Menú de cálculos eliminado; acción de emisiones está en botones inferiores de Nodos)

    def _open_dataset_editor(self, key: str):
        data = self.config.get_dataset(key)
        dlg = TableEditorDialog(self, key, data)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            # Guardado ya fue manejado en el diálogo
            pass

    def _menu_new_session_window(self):
        """Abre una nueva ventana de GSEQuant con una sesión vacía."""
        child = MainWindow()
        # Mantener referencia para que no se recolecte
        self._child_windows.append(child)
        child.show()

    # ---------------- Sesión XML (guardar/cargar proyecto) ---------------- #

    def _menu_open_session(self):
        start = os.getcwd()
        fn, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Abrir sesión GSEQuant",
            start,
            "Sesiones GSEQuant (*.gseq.xml);;XML (*.xml)",
        )
        if not fn:
            return
        try:
            self.load_session_from_xml(fn)
            self._session_path = fn
            self._update_session_title_and_status()
            QtWidgets.QMessageBox.information(self, "Sesión", f"Sesión cargada desde:\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Sesión", f"Error al cargar sesión:\n{e}")

    def _menu_save_session(self):
        if not self._session_path:
            self._menu_save_session_as()
            return
        try:
            self.save_session_to_xml(self._session_path)
            self._update_session_title_and_status()
            QtWidgets.QMessageBox.information(self, "Sesión", f"Sesión guardada en:\n{self._session_path}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Sesión", f"Error al guardar sesión:\n{e}")

    def _menu_save_session_as(self):
        start = os.getcwd()
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Guardar sesión GSEQuant",
            start,
            "Sesiones GSEQuant (*.gseq.xml);;XML (*.xml)",
        )
        if not fn:
            return
        try:
            self.save_session_to_xml(fn)
            self._session_path = fn
            self._update_session_title_and_status()
            QtWidgets.QMessageBox.information(self, "Sesión", f"Sesión guardada en:\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Sesión", f"Error al guardar sesión:\n{e}")

    def save_session_to_xml(self, path: str):
        """Serializa el estado actual a un XML de sesión GSEQuant.

        Incluye grafo, datasets de ConfigManager, parámetros de simulación,
        operaciones (Excel) y último resultado de emisiones por servicio.
        """
        NS_GSEQ = "urn:gsequant"
        NS_AIXM = "http://www.aixm.aero/schema/5.1"
        ET.register_namespace("gseq", NS_GSEQ)
        ET.register_namespace("aixm", NS_AIXM)

        root = ET.Element(f"{{{NS_GSEQ}}}Session", attrib={"version": "1.0"})

        # Grafo
        g_graph = ET.SubElement(root, f"{{{NS_GSEQ}}}Graph")
        g_nodes = ET.SubElement(g_graph, f"{{{NS_GSEQ}}}Nodes")
        for nid, node in self.model.nodes.items():
            el = ET.SubElement(g_nodes, f"{{{NS_GSEQ}}}Node", attrib={
                "id": node.id,
                "kind": node.kind,
                "hasJetbridge": "true" if node.has_jetbridge else "false",
                "hubBus": "true" if node.is_hub_bus else "false",
                "hubSta": "true" if node.is_hub_sta else "false",
                "hubBag": "true" if node.is_hub_bag else "false",
                "hubBel": "true" if node.is_hub_bel else "false",
                "hubGpu": "true" if node.is_hub_gpu else "false",
                "hubFue": "true" if node.is_hub_fue else "false",
                "rearEntry": "true" if node.is_rear_entry else "false",
            })
            ET.SubElement(el, f"{{{NS_GSEQ}}}name").text = node.name
            pos = ET.SubElement(el, f"{{{NS_GSEQ}}}position")
            pos.set("lat", str(node.lat))
            pos.set("lon", str(node.lon))
        g_edges = ET.SubElement(g_graph, f"{{{NS_GSEQ}}}Edges")
        for e in self.model.edges:
            el = ET.SubElement(g_edges, f"{{{NS_GSEQ}}}Edge", attrib={
                "id": e.id,
                "n1": e.n1,
                "n2": e.n2,
                "edgeClass": e.edge_class,
                "isOneWay": "true" if e.is_one_way else "false",
            })
            el.set("lengthKm", str(e.length_km))
        if self.model.default_base_id:
            ET.SubElement(g_graph, f"{{{NS_GSEQ}}}DefaultBase").text = self.model.default_base_id
        # Bases específicas por vehículo (opcional)
        if getattr(self.model, "vehicle_bases", None):
            g_vbases = ET.SubElement(g_graph, f"{{{NS_GSEQ}}}VehicleBases")
            for vcode, nid in (self.model.vehicle_bases or {}).items():
                vb = ET.SubElement(g_vbases, f"{{{NS_GSEQ}}}VehicleBase", attrib={"code": vcode})
                vb.text = nid

        # Datasets (ConfigManager)
        g_datasets = ET.SubElement(root, f"{{{NS_GSEQ}}}Datasets")
        for key in self.config.list_datasets():
            data = self.config.get_dataset(key)
            d_el = ET.SubElement(g_datasets, f"{{{NS_GSEQ}}}Dataset", attrib={"key": key})
            cols = data.get("columns", [])
            rows = data.get("rows", [])
            cols_el = ET.SubElement(d_el, f"{{{NS_GSEQ}}}Columns")
            for c in cols:
                c_el = ET.SubElement(cols_el, f"{{{NS_GSEQ}}}Column")
                c_el.text = str(c)
            rows_el = ET.SubElement(d_el, f"{{{NS_GSEQ}}}Rows")
            for r in rows:
                row_el = ET.SubElement(rows_el, f"{{{NS_GSEQ}}}Row")
                for c in cols:
                    v = r.get(c, "")
                    cell = ET.SubElement(row_el, f"{{{NS_GSEQ}}}Cell", attrib={"col": str(c)})
                    cell.text = "" if v is None else str(v)

        # Parámetros de simulación
        if hasattr(self, "_sim_params") and self._sim_params:
            g_sim = ET.SubElement(root, f"{{{NS_GSEQ}}}SimParams")
            dflt = self._sim_params.get("default", {}) or {}
            d_el = ET.SubElement(g_sim, f"{{{NS_GSEQ}}}Default")
            for k, v in dflt.items():
                el = ET.SubElement(d_el, f"{{{NS_GSEQ}}}{k}")
                el.text = "" if v is None else str(v)
            vehs = self._sim_params.get("veh", {}) or {}
            v_map_el = ET.SubElement(g_sim, f"{{{NS_GSEQ}}}Vehicles")
            for vcode, rec in vehs.items():
                v_el = ET.SubElement(v_map_el, f"{{{NS_GSEQ}}}Vehicle", attrib={"code": vcode})
                for k, v in (rec or {}).items():
                    el = ET.SubElement(v_el, f"{{{NS_GSEQ}}}{k}")
                    el.text = "" if v is None else str(v)

        # Operaciones (Excel) embebidas
        if getattr(self, "_ops_df", None) is not None and pd is not None:
            df = self._ops_df
            g_ops = ET.SubElement(root, f"{{{NS_GSEQ}}}Operations")
            cols = [str(c) for c in df.columns]
            cols_el = ET.SubElement(g_ops, f"{{{NS_GSEQ}}}Columns")
            for c in cols:
                c_el = ET.SubElement(cols_el, f"{{{NS_GSEQ}}}Column")
                c_el.text = c
            rows_el = ET.SubElement(g_ops, f"{{{NS_GSEQ}}}Rows")
            for i in range(len(df)):
                row_el = ET.SubElement(rows_el, f"{{{NS_GSEQ}}}Row")
                for j, c in enumerate(cols):
                    val = df.iloc[i, j]
                    cell = ET.SubElement(row_el, f"{{{NS_GSEQ}}}Cell", attrib={"col": c})
                    if pd.isna(val):
                        cell.text = ""
                    else:
                        cell.text = str(val)

        # Resultados de emisiones por servicio
        if self._last_emis_servicio:
            g_emis = ET.SubElement(root, f"{{{NS_GSEQ}}}EmissionsServicio")
            for veh, gases in self._last_emis_servicio.items():
                v_el = ET.SubElement(g_emis, f"{{{NS_GSEQ}}}Vehicle", attrib={"code": veh})
                for gas, val in gases.items():
                    g_el = ET.SubElement(v_el, f"{{{NS_GSEQ}}}Gas", attrib={"name": gas})
                    g_el.text = str(val)

        tree = ET.ElementTree(root)
        tree.write(path, encoding="utf-8", xml_declaration=True)

    def load_session_from_xml(self, path: str):
        """Carga una sesión GSEQuant desde un XML, restaurando grafo, datos y estado."""
        NS_GSEQ = "urn:gsequant"
        tree = ET.parse(path)
        root = tree.getroot()

        # Limpiar estado actual básico
        self.model = GraphModel()
        self.config = ConfigManager()
        self._ops_df = None
        self._last_emis_servicio = None

        # Grafo desde XML (respaldo para sesiones antiguas)
        g_graph = root.find(f".//{{{NS_GSEQ}}}Graph")
        if g_graph is not None:
            nodes_el = g_graph.find(f"{{{NS_GSEQ}}}Nodes")
            if nodes_el is not None:
                for n_el in nodes_el.findall(f"{{{NS_GSEQ}}}Node"):
                    nid = n_el.get("id")
                    kind = n_el.get("kind", "via")
                    has_jet = n_el.get("hasJetbridge", "false").lower() == "true"
                    name_el = n_el.find(f"{{{NS_GSEQ}}}name")
                    name = name_el.text if name_el is not None else nid
                    pos_el = n_el.find(f"{{{NS_GSEQ}}}position")
                    if pos_el is not None:
                        try:
                            lat = float(pos_el.get("lat", "0"))
                            lon = float(pos_el.get("lon", "0"))
                        except Exception:
                            lat, lon = 0.0, 0.0
                    else:
                        lat, lon = 0.0, 0.0
                    node = self.model.add_node(name=name, lat=lat, lon=lon, kind=kind, node_id=nid, has_jetbridge=has_jet)
                    node.is_hub_bus = n_el.attrib.get("hubBus", "false").lower() == "true"
                    node.is_hub_sta = n_el.attrib.get("hubSta", "false").lower() == "true"
                    node.is_hub_bag = n_el.attrib.get("hubBag", "false").lower() == "true"
                    node.is_hub_bel = n_el.attrib.get("hubBel", "false").lower() == "true"
                    node.is_hub_gpu = n_el.attrib.get("hubGpu", "false").lower() == "true"
                    node.is_hub_fue = n_el.attrib.get("hubFue", "false").lower() == "true"
                    node.is_rear_entry = n_el.attrib.get("rearEntry", "false").lower() == "true"
            edges_el = g_graph.find(f"{{{NS_GSEQ}}}Edges")
            if edges_el is not None:
                for e_el in edges_el.findall(f"{{{NS_GSEQ}}}Edge"):
                    eid = e_el.get("id")
                    n1 = e_el.get("n1")
                    n2 = e_el.get("n2")
                    eclass = e_el.get("edgeClass", "via")
                    is_one_way = e_el.get("isOneWay", "false").lower() == "true"
                    try:
                        length = float(e_el.get("lengthKm", "0"))
                    except Exception:
                        length = 0.0
                    try:
                        edge = self.model.add_edge(n1, n2, edge_class=eclass, is_one_way=is_one_way)
                        edge.id = eid
                        edge.length_km = length
                    except Exception:
                        continue
            dbase_el = g_graph.find(f"{{{NS_GSEQ}}}DefaultBase")
            if dbase_el is not None and dbase_el.text:
                self.model.default_base_id = dbase_el.text
            # Bases específicas por vehículo (si existen en la sesión)
            vbases_el = g_graph.find(f"{{{NS_GSEQ}}}VehicleBases")
            vmap: Dict[str, str] = {}
            if vbases_el is not None:
                for vb in vbases_el.findall(f"{{{NS_GSEQ}}}VehicleBase"):
                    code = vb.get("code")
                    nid = vb.text or ""
                    if code and nid:
                        vmap[code] = nid
            try:
                self.model.vehicle_bases = vmap
            except Exception:
                pass

        # Datasets
        g_datasets = root.find(f".//{{{NS_GSEQ}}}Datasets")
        if g_datasets is not None:
            for d_el in g_datasets.findall(f"{{{NS_GSEQ}}}Dataset"):
                key = d_el.get("key")
                if not key:
                    continue
                cols_el = d_el.find(f"{{{NS_GSEQ}}}Columns")
                cols = []
                if cols_el is not None:
                    for c_el in cols_el.findall(f"{{{NS_GSEQ}}}Column"):
                        cols.append(c_el.text or "")
                rows_el = d_el.find(f"{{{NS_GSEQ}}}Rows")
                rows = []
                if rows_el is not None:
                    for r_el in rows_el.findall(f"{{{NS_GSEQ}}}Row"):
                        rec = {}
                        for cell in r_el.findall(f"{{{NS_GSEQ}}}Cell"):
                            col_name = cell.get("col", "")
                            rec[col_name] = cell.text or ""
                        rows.append(rec)
                self.config.set_dataset(key, cols, rows)
            self.config.save_user_config()

        # Si existe una tabla de circulación, reconstruir el grafo desde ella
        try:
            circ_data = self.config.get_dataset("circulacion")
        except Exception:
            circ_data = {"columns": [], "rows": []}
        circ_rows = circ_data.get("rows") or []
        if circ_rows:
            # Reemplazar el modelo actual por uno derivado de la tabla de circulación
            # Conservar nodos actuales (cargados desde el XML de respaldo) para
            # poder reutilizar coordenadas u otros atributos si la tabla no los
            # define explícitamente (por ejemplo, cuando fue generada sin Lat/Lon).
            old_nodes = {}
            try:
                old_nodes = dict(getattr(self.model, "nodes", {}) or {})
            except Exception:
                old_nodes = {}
            # Conservar también el mapa de bases por vehículo, si existiera
            old_vbases = {}
            try:
                old_vbases = dict(getattr(self.model, "vehicle_bases", {}) or {})
            except Exception:
                old_vbases = {}

            self.model = GraphModel()
            try:
                self.model.vehicle_bases = old_vbases
            except Exception:
                pass

            def _bool(val) -> bool:
                if isinstance(val, bool):
                    return val
                if val is None:
                    return False
                text = str(val).strip().lower()
                return text in {"1", "true", "si", "sí", "yes", "y"}

            base_candidate = None
            # Primero nodos
            for row in circ_rows:
                cat = str(row.get("Categoria", "")).strip().lower()
                if cat != "nodo":
                    continue
                nid = str(row.get("ID") or row.get("Id") or row.get("Nodo") or "").strip()
                if not nid:
                    continue
                name = str(row.get("Nombre") or nid).strip()

                # Usar el mismo criterio que CirculationCalculator._load_dataset_graph:
                # priorizar columna "Tipo" y luego "Sentido" para el tipo de nodo.
                existing = old_nodes.get(nid)
                raw_kind = row.get("Tipo") or row.get("Sentido")
                if raw_kind is None and existing is not None:
                    raw_kind = getattr(existing, "kind", None)
                kind = str(raw_kind or "via").strip().lower()

                # Coordenadas: si la tabla no trae Lat/Lon (o están vacías),
                # reutilizar las del grafo original cargado desde XML.
                lat_val = row.get("Lat") if "Lat" in row else None
                lon_val = row.get("Lon") if "Lon" in row else None
                if lat_val is not None and str(lat_val).strip() != "":
                    try:
                        lat = float(lat_val)
                    except Exception:
                        lat = 0.0
                elif existing is not None:
                    try:
                        lat = float(getattr(existing, "lat", 0.0) or 0.0)
                    except Exception:
                        lat = 0.0
                else:
                    lat = 0.0

                if lon_val is not None and str(lon_val).strip() != "":
                    try:
                        lon = float(lon_val)
                    except Exception:
                        lon = 0.0
                elif existing is not None:
                    try:
                        lon = float(getattr(existing, "lon", 0.0) or 0.0)
                    except Exception:
                        lon = 0.0
                else:
                    lon = 0.0

                has_jet = _bool(row.get("Manga"))
                node = self.model.add_node(name=name, lat=lat, lon=lon, kind=kind, node_id=nid, has_jetbridge=has_jet)
                node.is_hub_bus = _bool(row.get("Es_hub_BUS"))
                node.is_hub_sta = _bool(row.get("Es_hub_STA"))
                node.is_hub_bag = _bool(row.get("Es_hub_BAG"))
                node.is_hub_bel = _bool(row.get("Es_hub_BEL"))
                if node.kind == "base" and base_candidate is None:
                    base_candidate = nid

            # Luego rutas
            for row in circ_rows:
                cat = str(row.get("Categoria", "")).strip().lower()
                if cat != "ruta":
                    continue
                n1 = str(row.get("Desde") or row.get("Origen") or row.get("n1") or "").strip()
                n2 = str(row.get("Hasta") or row.get("Destino") or row.get("n2") or "").strip()
                if not n1 or not n2:
                    continue
                try:
                    length = float(row.get("Dist_km", row.get("Dist", 0)) or 0)
                except Exception:
                    length = 0.0
                sentido = str(row.get("Sentido") or "").strip().lower()
                is_one_way = sentido.startswith("solo")
                eid = str(row.get("ID") or row.get("Id") or "").strip()
                try:
                    edge = self.model.add_edge(n1, n2, edge_class="via", is_one_way=is_one_way)
                    edge.length_km = length
                    if eid:
                        edge.id = eid
                except Exception:
                    continue

            # Base por defecto: la primera marcada como base o, en su defecto, cualquier nodo existente
            if base_candidate and base_candidate in self.model.nodes:
                self.model.default_base_id = base_candidate
            elif self.model.nodes and not self.model.default_base_id:
                self.model.default_base_id = next(iter(self.model.nodes.keys()))

        # Parámetros de simulación
        g_sim = root.find(f".//{{{NS_GSEQ}}}SimParams")
        if g_sim is not None:
            dflt_el = g_sim.find(f"{{{NS_GSEQ}}}Default")
            dflt = {}
            if dflt_el is not None:
                for el in list(dflt_el):
                    tag = el.tag.split("}")[-1]
                    dflt[tag] = el.text
            vehs_el = g_sim.find(f"{{{NS_GSEQ}}}Vehicles")
            vmap = {}
            if vehs_el is not None:
                for v_el in vehs_el.findall(f"{{{NS_GSEQ}}}Vehicle"):
                    code = v_el.get("code")
                    if not code:
                        continue
                    rec = {}
                    for el in list(v_el):
                        tag = el.tag.split("}")[-1]
                        rec[tag] = el.text
                    vmap[code] = rec
            self._sim_params = {"default": dflt, "veh": vmap}

        # Operaciones
        g_ops = root.find(f".//{{{NS_GSEQ}}}Operations")
        if g_ops is not None and pd is not None:
            cols_el = g_ops.find(f"{{{NS_GSEQ}}}Columns")
            cols = []
            if cols_el is not None:
                for c_el in cols_el.findall(f"{{{NS_GSEQ}}}Column"):
                    cols.append(c_el.text or "")
            rows_el = g_ops.find(f"{{{NS_GSEQ}}}Rows")
            data_rows = []
            if rows_el is not None:
                for r_el in rows_el.findall(f"{{{NS_GSEQ}}}Row"):
                    rec = {}
                    for cell in r_el.findall(f"{{{NS_GSEQ}}}Cell"):
                        col_name = cell.get("col", "")
                        rec[col_name] = cell.text
                    data_rows.append(rec)
            if cols:
                try:
                    import pandas as _pd
                    self._ops_df = _pd.DataFrame(data_rows, columns=cols)
                except Exception:
                    self._ops_df = None
        # Si se cargaron operaciones, activar botones de cálculo y previsualización
        try:
            if getattr(self, "_ops_df", None) is not None:
                if hasattr(self, "btnCalcEmisServicio"):
                    self.btnCalcEmisServicio.setEnabled(True)
                if hasattr(self, "btnPreviewOps"):
                    self.btnPreviewOps.setEnabled(True)
        except Exception:
            pass

        # Resultados de emisiones por servicio
        g_emis = root.find(f".//{{{NS_GSEQ}}}EmissionsServicio")
        if g_emis is not None:
            res = {}
            for v_el in g_emis.findall(f"{{{NS_GSEQ}}}Vehicle"):
                code = v_el.get("code")
                if not code:
                    continue
                gases = {}
                for g_el in v_el.findall(f"{{{NS_GSEQ}}}Gas"):
                    name = g_el.get("name")
                    try:
                        val = float(g_el.text) if g_el.text is not None else 0.0
                    except Exception:
                        val = 0.0
                    if name:
                        gases[name] = val
                res[code] = gases
            self._last_emis_servicio = res

        # Refrescar UI desde el modelo y datos cargados
        self._reload_map_from_model()
        self._reload_node_table_from_model()
        self._reload_edge_table_from_model()
        self._refresh_calc_combos()

    def _mark_hub(self, node_id: str, *, bus: bool = False,
                   sta: bool = False, bag: bool = False, bel: bool = False, gpu: bool = False, fue: bool = False, rear_entry: bool = False):
        node = self.model.nodes.get(node_id)
        if not node:
            return
        if bus or sta or bag or bel or gpu or fue or rear_entry:
            node.kind = 'hub'
        node.is_hub_bus = bus
        node.is_hub_sta = sta
        node.is_hub_bag = bag
        node.is_hub_bel = bel
        node.is_hub_gpu = gpu
        node.is_hub_fue = fue
        node.is_rear_entry = rear_entry
        self._table_add_or_update_node(node)
        self._js_add_or_update_node(node)

    def _toggle_rear_entry(self, node_id: str):
        node = self.model.nodes.get(node_id)
        if not node:
            return
        node.is_rear_entry = not node.is_rear_entry
        self._table_add_or_update_node(node)
        self._js_add_or_update_node(node)

    def _nodes_context_menu(self, pos):
        row = self.tblNodes.rowAt(pos.y())
        if row < 0:
            return
        self.tblNodes.selectRow(row)
        nid_item = self.tblNodes.item(row, 0)
        if not nid_item:
            return
        nid = nid_item.text()
        node = self.model.nodes.get(nid)
        if not node:
            return
        menu = QtWidgets.QMenu(self)
        act_center = menu.addAction("Centrar en el mapa")
        act_del = menu.addAction("Eliminar nodo")
        menu.addSeparator()
        act_set_bus = menu.addAction("Marcar como hub BUS (terminal pax)")
        act_set_sta = menu.addAction("Marcar como hub STA (escaleras plataforma)")
        act_set_bag = menu.addAction("Marcar como hub BAG (equipaje ↔ terminal)")
        act_set_bel = menu.addAction("Marcar como hub BEL (belt loader plataforma)")
        act_set_gpu = menu.addAction("Marcar como hub GPU")
        act_set_fue = menu.addAction("Marcar como hub FUE (combustible)")
        act_mark_rear = menu.addAction("Alternar Ingreso Trasero (Rear Entry)")
        act_mark_rear.setCheckable(True)
        act_mark_rear.setChecked(node.is_rear_entry)
        act_clear_hub = menu.addAction("Quitar rol de hub")
        action = menu.exec_(self.tblNodes.viewport().mapToGlobal(pos))
        if action == act_center:
            self._center_on_node_from_row()
        elif action == act_del:
            self.delete_selected_node_row()
        elif action == act_set_bus:
            self._mark_hub(
                nid,
                bus=True,
                sta=node.is_hub_sta,
                bag=node.is_hub_bag,
                bel=node.is_hub_bel,
                gpu=node.is_hub_gpu,
                fue=node.is_hub_fue,
                rear_entry=node.is_rear_entry,
            )
        elif action == act_set_sta:
            self._mark_hub(
                nid,
                bus=node.is_hub_bus,
                sta=True,
                bag=node.is_hub_bag,
                bel=node.is_hub_bel,
                gpu=node.is_hub_gpu,
                fue=node.is_hub_fue,
                rear_entry=node.is_rear_entry,
            )
        elif action == act_set_bag:
            self._mark_hub(
                nid,
                bus=node.is_hub_bus,
                sta=node.is_hub_sta,
                bag=True,
                bel=node.is_hub_bel,
                gpu=node.is_hub_gpu,
                fue=node.is_hub_fue,
                rear_entry=node.is_rear_entry,
            )
        elif action == act_set_bel:
            self._mark_hub(
                nid,
                bus=node.is_hub_bus,
                sta=node.is_hub_sta,
                bag=node.is_hub_bag,
                bel=True,
                gpu=node.is_hub_gpu,
                fue=node.is_hub_fue,
                rear_entry=node.is_rear_entry,
            )
        elif action == act_set_gpu:
            self._mark_hub(
                nid,
                bus=node.is_hub_bus,
                sta=node.is_hub_sta,
                bag=node.is_hub_bag,
                bel=node.is_hub_bel,
                gpu=True,
                fue=node.is_hub_fue,
                rear_entry=node.is_rear_entry,
            )
        elif action == act_set_fue:
            self._mark_hub(
                nid,
                bus=node.is_hub_bus,
                sta=node.is_hub_sta,
                bag=node.is_hub_bag,
                bel=node.is_hub_bel,
                gpu=node.is_hub_gpu,
                fue=True,
                rear_entry=node.is_rear_entry,
            )
        elif action == act_mark_rear:
            self._toggle_rear_entry(nid)
        elif action == act_clear_hub:
            self._mark_hub(nid, bus=False, sta=False, bag=False, bel=False, gpu=False, fue=False, rear_entry=False)

    def _edges_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        act_del = menu.addAction("Eliminar ruta")
        action = menu.exec_(self.tblEdges.viewport().mapToGlobal(pos))
        if action == act_del:
            self.delete_selected_edge_row()

    # ---------------- Eventos de mapa ---------------- #
    def on_map_clicked(self, lat: float, lon: float):
        # Si el botón de agregar nodo está activo, creamos nodo directamente
        if self.btnAddNode.isChecked():
            name = self.txtName.text().strip() or f"Nodo_{len(self.model.nodes)+1}"
            kind = self.cmbKind.currentText()
            has_jet = self.chkManga.isChecked()
            new_node = self.model.add_node(name=name, lat=lat, lon=lon, kind=kind, has_jetbridge=has_jet)
            self._js_add_or_update_node(new_node)
            self._table_add_or_update_node(new_node)
            return

        # Comportamiento normal (seleccionar para rutas)
        if self.edge_mode_active:
            # En modo "conectar", clic en el vacío cancela? No, usualmente esperamos click en nodo.
            # Pero si quisiéramos crear un nodo intermedio rápido:
            pass
        else:
            # Actualizar spinboxes para referencia
            self.dsbLat.setValue(lat)
            self.dsbLon.setValue(lon)

    def on_node_clicked(self, nid: str):
        # Lógica secuencial
        if self.btnSeqEdge.isChecked():
            if self.edge_first and self.edge_first != nid:
                # Crear ruta automática desde el anterior al actual
                eclass = "via" if self.cmbEdgeClass.currentIndex() == 0 else "connector"
                oneway = (self.cmbEdgeDir.currentIndex() == 1)
                try:
                    edge = self.model.add_edge(self.edge_first, nid, edge_class=eclass, is_one_way=oneway)
                    if edge:
                        self._js_add_edge(edge)
                        self._table_add_edge(edge)
                        # Avanzar el puntero: el nuevo destino es el nuevo origen
                        self.edge_first = nid
                        self.statusBar().showMessage(f"Ruta secuencial creada. Origen actual: {nid}")
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self, "Error", str(e))
            else:
                self.edge_first = nid
                self.statusBar().showMessage(f"Origen secuencial fijado: {nid}")
            return

        if self.edge_mode_active:
            if not self.edge_first:
                self.edge_first = nid
                self.btnStartEdge.setText(f"Desde: {nid}")
            else:
                n2 = nid
                eclass = "via" if self.cmbEdgeClass.currentIndex() == 0 else "connector"
                oneway = (self.cmbEdgeDir.currentIndex() == 1)
                try:
                    edge = self.model.add_edge(self.edge_first, n2, edge_class=eclass, is_one_way=oneway)
                    if edge:
                        self._js_add_edge(edge)
                        self._table_add_edge(edge)
                    self.edge_mode_active = False
                    self.edge_first = None
                    self.btnStartEdge.setText("Crear Ruta: Sel. Origen")
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self, "Error", str(e))
        else:
            # Seleccionar en tabla
            for i in range(self.tblNodes.rowCount()):
                if self.tblNodes.item(i, 0).text() == nid:
                    self.tblNodes.selectRow(i)
                    break

    def on_toggle_add_mode(self, checked: bool):
        if checked:
            self.btnAddNode.setText("Modo: AGREGANDO NODO (click mapa)")
            # Desactivar modo de conexión de aristas si está activo
            if self.edge_mode_active:
                self.finish_edge_mode()
            if self.btnSeqEdge.isChecked():
                self.btnSeqEdge.setChecked(False)
                self.on_toggle_sequential_routing_mode(False)
        else:
            self.btnAddNode.setText("MANDAR RECURSOS (Agregar Nodo)")

    def on_toggle_sequential_routing_mode(self, checked: bool):
        if checked:
            self.btnSeqEdge.setText("Modo: RUTAS SECUENCIALES (click nodos)")
            self.edge_first = None # Resetear el primer nodo para la secuencia
            self.edge_mode_active = False # Asegurarse de que el modo normal de aristas no esté activo
            self.btnAddNode.setChecked(False) # Desactivar modo de agregar nodo
            self.on_toggle_add_mode(False)
            self.statusBar().showMessage("Modo de rutas secuenciales activado. Clic en el primer nodo.")
        else:
            self.btnSeqEdge.setText("Modo Rutas Secuencial")
            self.edge_first = None
            self.statusBar().showMessage("Modo de rutas secuenciales desactivado.")

    def delete_node_by_id(self, nid: str):
        if nid not in self.model.nodes:
            return
        node = self.model.nodes[nid]
        res = QtWidgets.QMessageBox.question(
            self, "Eliminar Nodo",
            f"¿Desea eliminar el nodo '{node.name}' ({nid}) y sus rutas asociadas?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if res == QtWidgets.QMessageBox.Yes:
            self.model.remove_node(nid)
            self._js_remove_node(nid)
            self._reload_node_table_from_model()
            self._reload_edge_table_from_model()
            self._refresh_calc_combos() # Actualizar combos de calculadora si se eliminó un puesto

    def add_node_manual(self):
        try:
            lat = float(self.dsbLat.value()); lon = float(self.dsbLon.value())
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Error", "Coordenadas inválidas")
            return
        name = self.txtName.text().strip() or self.model.next_id()
        kind = self.cmbKind.currentText()
        has_jet = bool(self.chkManga.isChecked()) if kind == 'puesto' else False
        node = self.model.add_node(name=name, lat=lat, lon=lon, kind=kind, has_jetbridge=has_jet)
        self._js_add_or_update_node(node)
        self._table_add_or_update_node(node)
        if kind == 'base' and self.model.default_base_id is None:
            self.model.default_base_id = node.id
        self._refresh_calc_combos()

    def apply_raw_coords(self):
        """Convierte texto en otros formatos (DMS / UTM) a decimales y escribe en Lat/Lon."""
        text = (self.txtCoordRaw.text() or "").strip()
        if not text:
            return

        # 1) Intentar DD/DMS usando el parser existente
        coords = parse_coords(text)
        if coords is not None:
            lat, lon = coords
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                self.dsbLat.setValue(lat)
                self.dsbLon.setValue(lon)
                return

        # 2) Intentar UTM
        coords = parse_utm_coords(text)
        if coords is not None:
            lat, lon = coords
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                self.dsbLat.setValue(lat)
                self.dsbLon.setValue(lon)
                return

        QtWidgets.QMessageBox.warning(
            self,
            "Coordenadas",
            "No se pudo interpretar el formato de coordenadas.\n"
            "Formatos soportados: DD, DMS (con o sin símbolos) y UTM simple.",
        )

    def start_edge_mode(self):
        self.edge_first = None
        self.edge_mode_active = True
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Modo arista: clickee el primer nodo en el mapa")

    def finish_edge_mode(self):
        # Salir explícitamente del modo arista
        self.edge_first = None
        self.edge_mode_active = False
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Modo arista finalizado. Puede volver a agregar nodos normalmente.")

    # ---------------- Buscador ---------------- #
    def on_search(self):
        q = self.txtSearch.text().strip()
        if not q:
            return
        # 1) buscar por nombre de nodo
        for node in self.model.nodes.values():
            if q.lower() in (node.name or '').lower():
                self._js_fly_to(node.lat, node.lon); return
        # 2) buscar por coords DD/DMS
        coords = parse_coords(q)
        if coords:
            self._js_fly_to(coords[0], coords[1]); return
        # 3) geocodificar (opcional)
        res = geocode_text(q)
        if res:
            lat, lon, _ = res; self._js_fly_to(lat, lon)
        else:
            QtWidgets.QMessageBox.information(self, "Buscar", "No se encontraron resultados.")

    # ---------------- Archivo ---------------- #
    def export_json(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Exportar grafo", os.getcwd(), "JSON (*.json)")
        if not fn: return
        try:
            self.model.export_json(fn)
            QtWidgets.QMessageBox.information(self, "Exportación", f"Grafo exportado a\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    def import_json(self):
        fn, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Importar grafo", os.getcwd(), "JSON (*.json)")
        if not fn: return
        try:
            self.model.load_json(fn)
            self._reload_map_from_model()
            self._reload_node_table_from_model()
            self._reload_edge_table_from_model()
            self._refresh_calc_combos()
            QtWidgets.QMessageBox.information(self, "Importación", f"Grafo importado desde\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    def set_default_base_from_table(self):
        row = self.tblNodes.currentRow()
        if row < 0:
            QtWidgets.QMessageBox.information(self, "Base", "Seleccione una fila (nodo) en la tabla.")
            return
        nid = self.tblNodes.item(row, 0).text()
        self.model.default_base_id = nid
        QtWidgets.QMessageBox.information(self, "Base", f"Base por defecto fijada en: {nid} – {self.model.nodes[nid].name}")

    def center_on_default_base(self):
        nid = self.model.default_base_id
        if not nid or nid not in self.model.nodes:
            QtWidgets.QMessageBox.information(self, "Base", "No hay base por defecto definida.")
            return
        node = self.model.nodes[nid]
        self._js_fly_to(node.lat, node.lon)

    def _open_vehicle_bases_dialog(self):
        # Permite asignar una base inicial específica por tipo de vehículo GSE.
        if not self.model.nodes:
            QtWidgets.QMessageBox.information(
                self,
                "Bases por vehículo",
                "No hay nodos cargados en el grafo. Cree al menos una base antes de configurar.",
            )
            return
        has_base = any(getattr(n, "kind", "").lower() == "base" for n in self.model.nodes.values())
        if not has_base:
            QtWidgets.QMessageBox.information(
                self,
                "Bases por vehículo",
                "No hay nodos de tipo 'base'. Defina al menos una base en el grafo para poder asignarla a los vehículos.",
            )
            return
        try:
            current_map = getattr(self.model, "vehicle_bases", {}) or {}
        except Exception:
            current_map = {}
        dlg = BaseAssignmentDialog(self, self.model.nodes, current_map)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            selected = dlg.selected_bases()
            try:
                self.model.vehicle_bases = selected
            except Exception:
                pass

    # ---------------- Calculadora ---------------- #
    def _refresh_calc_combos(self):
        puestos = sorted([(nid, n) for nid, n in self.model.nodes.items() if n.kind == "puesto"], key=lambda x: x[1].name.lower())
        self.cmbCalcFrom.blockSignals(True); self.cmbCalcTo.blockSignals(True)
        self.cmbCalcFrom.clear(); self.cmbCalcTo.clear()
        for nid, n in puestos:
            label = f"{n.name} ({nid})"
            self.cmbCalcFrom.addItem(label, nid)
            self.cmbCalcTo.addItem(label, nid)
        self.cmbCalcFrom.blockSignals(False); self.cmbCalcTo.blockSignals(False)

    def _graph_debug_report(self, nid_from: str, nid_to: str) -> str:
        """Genera un informe legible del grafo para depuración."""
        import networkx as _nx
        lines = []
        lines.append(f'Nodos totales: {len(self.model.nodes)}')
        lines.append(f'Aristas totales (lista interna): {len(self.model.edges)}')
        lines.append('--- Nodos (id : name | tipo) ---')
        for nid, n in sorted(self.model.nodes.items()):
            lines.append(f'  {nid} : {n.name} | {n.kind}')
        lines.append('--- Aristas (id : n1 -> n2 | tipo | is_one_way | length_km) ---')
        for e in self.model.edges:
            lines.append(f'  {e.id} : {e.n1} -> {e.n2} | {e.edge_class} | one_way={e.is_one_way} | {e.length_km:.3f} km')
        lines.append('--- Adyacencia en el grafo (salientes) ---')
        G = self.model.G
        if nid_from in G:
            outs = list(G.successors(nid_from))
            lines.append(f'  Salientes desde {nid_from}: {outs}')
            for v in outs:
                attr = G.get_edge_data(nid_from, v, default={})
                lines.append(f'    {nid_from} -> {v} : {attr}')
        else:
            lines.append(f'  Origen {nid_from} no existe en DiGraph.')
        if nid_to in G:
            ins = list(G.predecessors(nid_to))
            lines.append(f'  Entrantes a {nid_to}: {ins}')
        else:
            lines.append(f'  Destino {nid_to} no existe en DiGraph.')
        try:
            has = _nx.has_path(G, nid_from, nid_to)
            lines.append(f'Path exists (nx.has_path): {has}')
        except Exception as ex:
            lines.append(f'Error comprobando path: {ex}')
        if nid_from in G and not _nx.has_path(G, nid_from, nid_to):
            reachable = list(_nx.descendants(G, nid_from))
            lines.append(f'  Nodos alcanzables desde {nid_from} (count={len(reachable)}): {reachable[:100]}')
        return '\n'.join(lines)



    def calc_min_distance(self):
        nid_from = self.cmbCalcFrom.currentData()
        nid_to = self.cmbCalcTo.currentData()
        # Normalizar a str si vienen como QVariant u otro
        try:
            nid_from = str(nid_from) if nid_from is not None else None
            nid_to = str(nid_to) if nid_to is not None else None
        except Exception:
            pass

        if not nid_from or not nid_to:
            QtWidgets.QMessageBox.information(self, "Calculadora", "Seleccione origen y destino (puestos)." )
            return
        if nid_from == nid_to:
            QtWidgets.QMessageBox.information(self, "Calculadora", "Origen y destino son el mismo." )
            return

        # validación extra: existen los nodos en el grafo?
        missing = [x for x in (nid_from, nid_to) if x not in self.model.G]
        if missing:
            QtWidgets.QMessageBox.warning(self, "Calculadora", f"Los siguientes nodos no existen en el grafo: {missing}\nCompruebe import/edición." )
            return

        try:
            length = nx.shortest_path_length(self.model.G, nid_from, nid_to, weight="weight")
            path = nx.shortest_path(self.model.G, nid_from, nid_to, weight="weight")
        except (nx.NetworkXNoPath, nx.NodeNotFound):
            # generar informe de depuración y mostrar al usuario
            info = self._graph_debug_report(nid_from, nid_to)
            dlg = QtWidgets.QMessageBox(self)
            dlg.setWindowTitle("Calculadora - no hay camino")
            dlg.setIcon(QtWidgets.QMessageBox.Warning)
            dlg.setText("No existe un camino entre esos puestos. Use conectores y sentido correcto.")
            dlg.setDetailedText(info)
            dlg.exec_()
            return

        msg = f"Distancia mínima (red dirigida): {length:.3f} km\nCamino: {' → '.join(path)}"
        if self.chkDrawPath.isChecked():
            self._js_clear_routes()
            coords = [[self.model.nodes[n].lat, self.model.nodes[n].lon] for n in path]
            self._js_draw_route(coords)
        QtWidgets.QMessageBox.information(self, "Resultado", msg)

    # ---------------- Tabla de NODOS ---------------- #
    def _table_add_or_update_node(self, node: Node):
        row = None
        for i in range(self.tblNodes.rowCount()):
            if self.tblNodes.item(i, 0) and self.tblNodes.item(i, 0).text() == node.id:
                row = i; break
        if row is None:
            row = self.tblNodes.rowCount()
            self.tblNodes.insertRow(row)
            self.tblNodes.setItem(row, 0, QtWidgets.QTableWidgetItem(node.id))
            self.tblNodes.item(row, 0).setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)

        self.tblNodes.blockSignals(True)
        self.tblNodes.setItem(row, 1, QtWidgets.QTableWidgetItem(node.name))
        self.tblNodes.setItem(row, 2, QtWidgets.QTableWidgetItem(node.kind))
        self.tblNodes.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{node.lat:.6f}"))
        self.tblNodes.setItem(row, 4, QtWidgets.QTableWidgetItem(f"{node.lon:.6f}"))
        self.tblNodes.setItem(row, 5, QtWidgets.QTableWidgetItem("sí" if node.has_jetbridge else "no"))
        labels = []
        if node.is_hub_bus:
            labels.append("BUS")
        if node.is_hub_sta:
            labels.append("STA")
        if node.is_hub_bag:
            labels.append("BAG")
        if node.is_hub_bel:
            labels.append("BEL")
        if getattr(node, 'is_hub_gpu', False):
            labels.append("GPU")
        if getattr(node, 'is_hub_fue', False):
            labels.append("FUE")
        hub_txt = ", ".join(labels) if labels else "—"
        self.tblNodes.setItem(row, 6, QtWidgets.QTableWidgetItem(hub_txt))
        
        # Columna extra manga + rear? No, vamos a usar Manga y agregar Rear en el texto o columna
        rear_txt = "Sí" if getattr(node, 'is_rear_entry', False) else "No"
        self.tblNodes.setItem(row, 5, QtWidgets.QTableWidgetItem(f"Manga: {'Sí' if node.has_jetbridge else 'No'}, Trasero: {rear_txt}"))
        
        self.tblNodes.blockSignals(False)

    def _reload_node_table_from_model(self):
        self.tblNodes.blockSignals(True)
        self.tblNodes.setRowCount(0)
        for node in self.model.nodes.values():
            self._table_add_or_update_node(node)
        self.tblNodes.blockSignals(False)

    def on_table_item_changed(self, item: QtWidgets.QTableWidgetItem):
        row = item.row()
        nid = self.tblNodes.item(row, 0).text() if self.tblNodes.item(row, 0) else None
        if not nid or nid not in self.model.nodes:
            return
        node = self.model.nodes[nid]

        name = self.tblNodes.item(row, 1).text() if self.tblNodes.item(row, 1) else node.name
        kind = self.tblNodes.item(row, 2).text() if self.tblNodes.item(row, 2) else node.kind
        lat_txt = self.tblNodes.item(row, 3).text() if self.tblNodes.item(row, 3) else str(node.lat)
        lon_txt = self.tblNodes.item(row, 4).text() if self.tblNodes.item(row, 4) else str(node.lon)
        manga_txt = self.tblNodes.item(row, 5).text() if self.tblNodes.item(row, 5) else ("sí" if node.has_jetbridge else "no")

        coords = parse_coords(f"{lat_txt} {lon_txt}")
        if coords:
            lat, lon = coords
        else:
            try:
                lat = float(lat_txt); lon = float(lon_txt)
            except Exception:
                QtWidgets.QMessageBox.warning(self, "Tabla", "Coordenadas inválidas (use DD o DMS)." )
                self._table_add_or_update_node(node)
                return

        has_jet = manga_txt.strip().lower() in ("si", "sí", "true", "1", "y", "yes")
        if kind not in ("via", "puesto", "base", "hub"):
            QtWidgets.QMessageBox.warning(self, "Tabla", "Tipo inválido. Use: via | puesto | base | hub" )
            self._table_add_or_update_node(node); return
        if kind != "puesto":
            has_jet = False

        self.model.update_node(nid, name=name, kind=kind, lat=lat, lon=lon, has_jetbridge=has_jet)
        self._js_add_or_update_node(self.model.nodes[nid])
        self._table_add_or_update_node(self.model.nodes[nid])
        self._refresh_calc_combos()

        # actualizar longitudes de edges afectados y tabla
        changed = False
        for e in self.model.edges:
            if e.n1 == nid or e.n2 == nid:
                a = self.model.nodes[e.n1]; b = self.model.nodes[e.n2]
                e.length_km = haversine_km(a.lat, a.lon, b.lat, b.lon)
                changed = True
        if changed:
            self._reload_edge_table_from_model()

    def _center_on_node_from_row(self, *args):
        row = self.tblNodes.currentRow()
        if row < 0: return
        nid = self.tblNodes.item(row, 0).text()
        if nid in self.model.nodes:
            n = self.model.nodes[nid]
            self._js_fly_to(n.lat, n.lon)

    def export_nodes_csv(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Exportar nodos", os.getcwd(), "CSV (*.csv)")
        if not fn: return
        try:
            with open(fn, 'w', encoding='utf-8') as f:
                f.write("id,nombre,tipo,lat,lon,manga\n")
                for node in self.model.nodes.values():
                    f.write(f"{node.id},{node.name},{node.kind},{node.lat:.6f},{node.lon:.6f},{int(node.has_jetbridge)}\n")
            QtWidgets.QMessageBox.information(self, "Exportación", f"Nodos exportados a\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    def copy_nodes_to_clipboard(self):
        rows = ["id,nombre,tipo,lat,lon,manga"]
        for node in self.model.nodes.values():
            rows.append(f"{node.id},{node.name},{node.kind},{node.lat:.6f},{node.lon:.6f},{int(node.has_jetbridge)}")
        QtWidgets.QApplication.clipboard().setText("\n".join(rows))
        QtWidgets.QMessageBox.information(self, "Copiado", "Nodos copiados al portapapeles (CSV)." )

    def delete_selected_node_row(self):
        row = self.tblNodes.currentRow()
        if row < 0: return
        nid = self.tblNodes.item(row, 0).text()
        if nid in self.model.nodes:
            if QtWidgets.QMessageBox.question(self, "Confirmar", f"Eliminar nodo {nid}?\nSe eliminarán también sus rutas.") != QtWidgets.QMessageBox.Yes:
                return
            try:
                self.model.G.remove_node(nid)
            except Exception:
                pass
            # borrar edges asociados y del mapa
            to_del = [e for e in self.model.edges if e.n1 == nid or e.n2 == nid]
            for e in to_del:
                self._js_remove_edge(e.id)
            self.model.nodes.pop(nid, None)
            self.model.edges = [e for e in self.model.edges if e.n1 != nid and e.n2 != nid]
            self._reload_edge_table_from_model()
            self._js_remove_node(nid)
            self.tblNodes.blockSignals(True)
            self.tblNodes.removeRow(row)
            self.tblNodes.blockSignals(False)
            self._refresh_calc_combos()

            # Si el nodo borrado estaba siendo usado como origen en el modo arista, reiniciar ese estado
            if self.edge_first == nid:
                self.edge_first = None

    # ---------------- Tabla de RUTAS ---------------- #
    def _table_add_edge(self, edge: Edge):
        row = self.tblEdges.rowCount()
        self.tblEdges.insertRow(row)
        sentido = "Solo ida ({}→{})".format(edge.n1, edge.n2) if edge.is_one_way else "Doble"
        tipo = "Conector" if edge.edge_class == "connector" else "Vía"
        n1 = self.model.nodes.get(edge.n1); n2 = self.model.nodes.get(edge.n2)
        desde_txt = f"{n1.name if n1 else edge.n1} ({edge.n1})"
        hasta_txt = f"{n2.name if n2 else edge.n2} ({edge.n2})"
        self.tblEdges.setItem(row, 0, QtWidgets.QTableWidgetItem(edge.id))
        self.tblEdges.setItem(row, 1, QtWidgets.QTableWidgetItem(desde_txt))
        self.tblEdges.setItem(row, 2, QtWidgets.QTableWidgetItem(hasta_txt))
        self.tblEdges.setItem(row, 3, QtWidgets.QTableWidgetItem(tipo))
        self.tblEdges.setItem(row, 4, QtWidgets.QTableWidgetItem(sentido))
        self.tblEdges.setItem(row, 5, QtWidgets.QTableWidgetItem(f"{edge.length_km:.3f}"))

    def _reload_edge_table_from_model(self):
        self.tblEdges.setRowCount(0)
        for e in self.model.edges:
            self._table_add_edge(e)

    def on_edge_clicked(self, edge_id: str):
        """Selecciona en la tabla de rutas la fila correspondiente a la arista clickeada en el mapa."""
        if not edge_id:
            return
        try:
            target_row = -1
            for r in range(self.tblEdges.rowCount()):
                it = self.tblEdges.item(r, 0)
                if it and it.text() == edge_id:
                    target_row = r
                    break
            if target_row >= 0:
                self.tblEdges.setCurrentCell(target_row, 0)
                self.tblEdges.scrollToItem(
                    self.tblEdges.item(target_row, 0),
                    QtWidgets.QAbstractItemView.PositionAtCenter,
                )
        except Exception:
            pass

    def delete_selected_edge_row(self):
        row = self.tblEdges.currentRow()
        if row < 0: return
        edge_id = self.tblEdges.item(row, 0).text()
        if QtWidgets.QMessageBox.question(self, "Confirmar", f"Eliminar ruta {edge_id}?") != QtWidgets.QMessageBox.Yes:
            return
        if self.model.remove_edge(edge_id):
            self.tblEdges.removeRow(row)
            self._js_remove_edge(edge_id)
        else:
            QtWidgets.QMessageBox.warning(self, "Rutas", "No se pudo eliminar la ruta.")

    def export_edges_csv(self):
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Exportar rutas", os.getcwd(), "CSV (*.csv)")
        if not fn: return
        try:
            with open(fn, 'w', encoding='utf-8') as f:
                f.write("id,desde(humano),desde(id),hasta(humano),hasta(id),tipo,sentido,long_km\n")
                for e in self.model.edges:
                    n1 = self.model.nodes.get(e.n1); n2 = self.model.nodes.get(e.n2)
                    tipo = "conector" if e.edge_class == "connector" else "via"
                    sentido = "solo_ida" if e.is_one_way else "doble"
                    f.write(f"{e.id},{(n1.name if n1 else e.n1)},{e.n1},{(n2.name if n2 else e.n2)},{e.n2},{tipo},{sentido},{e.length_km:.3f}\n")
            QtWidgets.QMessageBox.information(self, "Exportación", f"Rutas exportadas a\n{fn}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    # ---------------- Eventos de ventana ---------------- #
    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        """Pregunta si se desea guardar la sesión antes de cerrar."""
        resp = QtWidgets.QMessageBox.question(
            self,
            "Cerrar GSEQuant",
            "¿Desea guardar la sesión actual antes de salir?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel,
            QtWidgets.QMessageBox.Yes,
        )
        if resp == QtWidgets.QMessageBox.Cancel:
            event.ignore()
            return
        if resp == QtWidgets.QMessageBox.Yes:
            try:
                # Si ya hay ruta de sesión, reutilizarla; si no, pedir "Guardar como"
                if getattr(self, "_session_path", None):
                    try:
                        self.save_session_to_xml(self._session_path)
                    except Exception:
                        # Si falla el guardado directo, ofrecer "Guardar como"
                        self._menu_save_session_as()
                else:
                    self._menu_save_session_as()
            except Exception:
                # En caso de error, dejamos continuar el cierre
                pass
        event.accept()

    # ---------------- JS helpers ---------------- #
    def _js_eval(self, script: str):
        self.view.page().runJavaScript(script)

    def _js_add_or_update_node(self, node: Node):
        payload = json.dumps(asdict(node))
        self._js_eval(f"updateNode({payload});")

    def _js_remove_node(self, nid: str):
        self._js_eval(f"removeNode('{nid}');")

    def _js_add_edge(self, edge: Edge):
        payload = json.dumps(asdict(edge))
        self._js_eval(f"addEdge({payload});")

    def _js_remove_edge(self, edge_id: str):
        self._js_eval(f"removeEdge('{edge_id}');")

    def _js_clear_routes(self):
        self._js_eval("clearRoutes();")

    def _js_draw_route(self, coords: List[List[float]]):
        self._js_eval(f"drawRoute({json.dumps(coords)});")

    def _js_fly_to(self, lat: float, lon: float, z: int = 17):
        # Mover el mapa
        self._js_eval(f"flyTo({lat}, {lon}, {z});")
        # Actualizar etiqueta de lugar (reverse geocoding simple)
        try:
            place = reverse_geocode(lat, lon)
        except Exception:
            place = None
        if place:
            self._js_eval(f"setPlaceLabel({json.dumps(place)});")
        else:
            self._js_eval("setPlaceLabel('');")

    def _reload_map_from_model(self):
        self.view.setHtml(LEAFLET_HTML)
        QtCore.QTimer.singleShot(300, self._redraw_after_reload)

    def _redraw_after_reload(self):
        for node in self.model.nodes.values():
            self._js_add_or_update_node(node)
        for e in self.model.edges:
            self._js_add_edge(e)
        # Centrar mapa en la zona de trabajo: base por defecto o centroide de nodos
        try:
            if self.model.default_base_id and self.model.default_base_id in self.model.nodes:
                n = self.model.nodes[self.model.default_base_id]
                self._js_fly_to(n.lat, n.lon)
            elif self.model.nodes:
                lats = [n.lat for n in self.model.nodes.values()]
                lons = [n.lon for n in self.model.nodes.values()]
                self._js_fly_to(sum(lats) / len(lats), sum(lons) / len(lons))
        except Exception:
            pass

    # ---------------- Emisiones en servicio ---------------- #
    def _compute_service_emissions_from_history(self) -> Dict[str, Dict[str, float]]:
        """Calcula emisiones en servicio por vehículo usando el histórico de operaciones.

        Integra la matriz GSE×aeronaves (gsexaeronaves) y distingue rampas/remotas
        según la presencia de manga. Para BUS y BAG añade tiempo de servicio
        equivalente en el hub (carga de pasajeros/equipaje en terminal).
        """
        # El cálculo de servicio debe respetar siempre la configuración de puestos
        # (manga/no manga, hubs, etc.), por lo que requiere una tabla REAL de
        # circulación. Si no existe, no hay datos suficientes para un resultado
        # coherente.
        dataset = self._ensure_circulation_dataset()
        if dataset is None:
            return {}, {}

        # Reutilizar la lógica de operaciones y demanda del CirculationCalculator,
        # pero alimentado con la misma tabla de circulación que se usa para el
        # cálculo de emisiones por circulación.
        circ = CirculationCalculator(
            self.model,
            self.config,
            circ_params=getattr(self, "_circ_params", None),
            sim_params=getattr(self, "_sim_params", None),
            dataset=dataset,
            debug_enabled=False,
        )
        if self._date_filter_from is not None and self._date_filter_to is not None:
            circ.set_date_filter(self._date_filter_from, self._date_filter_to)
        circ.set_operations_df(self._get_active_ops_df())
        gse_matrix = circ._load_gse_aircraft_matrix()
        ops = list(getattr(circ, "_ops", []))
        veh_codes = circ._vehicle_codes()

        # Comprobar si existen hubs específicos para BUS/BAG, ya sea en el grafo
        # principal o en el grafo derivado del dataset de circulación.
        hub_bus_exists = any(n.is_hub_bus for n in self.model.nodes.values()) or \
                         any(n.is_hub_bus for n in getattr(circ, "_ds_nodes", {}).values())
        hub_bag_exists = any(n.is_hub_bag for n in self.model.nodes.values()) or \
                         any(n.is_hub_bag for n in getattr(circ, "_ds_nodes", {}).values())

        # Acumular horas efectivas de servicio por vehículo y estimar flota mínima
        # de servicio usando concurrencia de intervalos (sin circulación).
        total_service_h: Dict[str, float] = {}
        service_min_fleet: Dict[str, int] = {}
        for veh in veh_codes:
            veh_matrix = gse_matrix.get(veh, {})
            if not veh_matrix:
                continue
            acc = 0.0
            intervals: List[Tuple[float, float]] = []
            for op in ops:
                count, svc_time = circ._operation_demand(veh, veh_matrix, op)
                if count <= 0 or svc_time <= 0:
                    continue
                eff = svc_time
                # Para BUS y BAG, considerar tiempo en hub además del puesto
                if veh == "BUS" and hub_bus_exists:
                    eff = svc_time * 2.0
                elif veh == "BAG" and hub_bag_exists:
                    eff = svc_time * 2.0
                acc += float(count) * eff
                units_int = int(math.ceil(float(count)))
                start = op.arr
                end = op.arr + eff
                for _ in range(units_int):
                    intervals.append((start, end))
            if acc > 0.0:
                total_service_h[veh] = acc
            if intervals:
                try:
                    service_min_fleet[veh] = circ._max_concurrent(intervals)
                except Exception:
                    pass

        # Guardar para uso posterior (por ejemplo, en el cálculo de totales)
        self._service_min_fleet = service_min_fleet

        if not total_service_h:
            return {}, {}

        # Usar EmissionsCalculator como helper para obtener K y patrones FCD/FC/t
        emis = EmissionsCalculator(self.config)
        if hasattr(self, "_sim_params") and self._sim_params:
            emis.set_overrides(self._sim_params)
        coef_rows = emis._load_coef_vehiculos()
        ef_rows = emis._load_EF()
        gases = ["CO2", "CO", "HC", "NOx", "SOx", "PM10"]

        resultados: Dict[str, Dict[str, float]] = {}
        for veh, T_h in total_service_h.items():
            vrow = emis._find_vehicle_row(veh, coef_rows)
            if not vrow:
                continue
            hp_vehicle = emis._vehicle_hp(vrow)
            efrow = emis._select_EF_row_for_veh(veh, hp_vehicle, ef_rows)
            if not efrow:
                continue
            # Patrones de factores y tiempos (como en compute_emisiones_servicio)
            FCD = [emis._to_float(vrow.get("FCD 1 ")), emis._to_float(vrow.get("FCD 2")), emis._to_float(vrow.get("FCD 3")), emis._to_float(vrow.get("FCD 4"))]
            FC  = [emis._to_float(vrow.get("FC 1 ")),  emis._to_float(vrow.get("FC 2")),  emis._to_float(vrow.get("FC 3")),  emis._to_float(vrow.get("FC 4"))]
            tD  = [emis._to_float(vrow.get("tD1")),    emis._to_float(vrow.get("tD2")),    emis._to_float(vrow.get("tD3")),    emis._to_float(vrow.get("tD4"))]
            tC  = [emis._to_float(vrow.get("t1")),     emis._to_float(vrow.get("t2")),     emis._to_float(vrow.get("t3")),     emis._to_float(vrow.get("t4"))]
            FCD = [x or 0.0 for x in FCD]; FC = [x or 0.0 for x in FC]; tD = [x or 0.0 for x in tD]; tC = [x or 0.0 for x in tC]
            base_time = sum(tC) + sum(tD)
            if base_time <= 0.0:
                base_time = 1.0

            res_gases: Dict[str, float] = {}
            age_o = emis._ovr(veh, 'age', None)
            tutil_o = emis._ovr(veh, 't_util', None)
            hp_o = emis._ovr(veh, 'hp_vehicle', None)
            for gas in gases:
                k = emis._k_for(vrow, efrow, gas, veh,
                                age_override=age_o,
                                t_util_override=tutil_o,
                                hp_override=hp_o)
                if k is None:
                    continue
                e_c = sum(k * FC[i] * tC[i] for i in range(4))
                e_d = sum(k * FCD[i] * tD[i] for i in range(4))
                total_base = e_c + e_d
                emis_per_h = total_base / base_time
                res_gases[gas] = emis_per_h * T_h
            if res_gases:
                resultados[veh] = res_gases

        return resultados, total_service_h

    def open_emissions_dialog(self):
        try:
            # Bloquear si no hay Excel cargado y válido
            if getattr(self, "_ops_df", None) is None:
                QtWidgets.QMessageBox.information(self, "Operaciones", "Debe cargar un archivo Excel de operaciones compatible antes de calcular.")
                return
            # También requiere una tabla de circulación (desde Excel o generada
            # desde el grafo) para conocer mangas y hubs por puesto.
            dataset = self._ensure_circulation_dataset()
            if dataset is None:
                QtWidgets.QMessageBox.information(
                    self,
                    "Circulación",
                    "Debe cargar o generar una tabla de circulación (con puestos y rutas) antes de calcular emisiones en servicio.",
                )
                return
            # Verificar que todas las aeronaves del Excel tengan configuración
            # en gsexaeronaves antes de calcular
            if not self._check_missing_aircraft_for_operations():
                return
            # Verificar que existan parámetros de servicio para todos los GSE
            if not self._validate_service_config():
                return
            resultados, svc_hours = self._compute_service_emissions_from_history()
            if not resultados:
                return
            # guardar último resultado de emisiones por servicio para la sesión
            self._last_emis_servicio = resultados

            # Construir tasas medias [g/s] por gas a partir de las horas efectivas de servicio
            rates: Dict[str, Dict[str, float]] = {}
            for veh, gases in resultados.items():
                T_h = float(svc_hours.get(veh, 0.0) or 0.0)
                if T_h <= 0.0:
                    continue
                veh_rates: Dict[str, float] = {}
                for gas, mass_g in gases.items():
                    try:
                        m = float(mass_g)
                    except Exception:
                        m = 0.0
                    veh_rates[gas] = m / (T_h * 3600.0)
                if veh_rates:
                    rates[veh] = veh_rates

            dlg = EmissionsResultsDialog(self, resultados, rates)
            dlg.exec_()
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Emisiones", str(e))

    def open_circulation_emissions_dialog(self):
        # Requiere operaciones reales y tabla de circulación real
        if getattr(self, "_ops_df", None) is None or getattr(self._ops_df, 'empty', False):
            QtWidgets.QMessageBox.information(
                self,
                "Operaciones",
                "Debe cargar un archivo Excel de operaciones compatible antes de calcular.",
            )
            return
        dataset = self._ensure_circulation_dataset()
        if dataset is None:
            QtWidgets.QMessageBox.information(
                self,
                "Circulación",
                "Debe cargar o generar una tabla de circulación antes de calcular.",
            )
            return
        # Verificar aeronaves sin configuración en gsexaeronaves
        if not self._check_missing_aircraft_for_operations():
            return
        # Verificar parámetros de servicio (coef_vehiculos/EF) para todos los GSE
        if not self._validate_service_config():
            return
        calc = CirculationCalculator(self.model, self.config,
                                     circ_params=getattr(self, "_circ_params", None),
                                     sim_params=getattr(self, "_sim_params", None),
                                     dataset=dataset,
                                     debug_enabled=getattr(self, "_circ_debug_enabled", False))
        if self._date_filter_from is not None and self._date_filter_to is not None:
            calc.set_date_filter(self._date_filter_from, self._date_filter_to)
        calc.set_operations_df(self._get_active_ops_df())
        results, warnings = calc.compute()
        self._last_circ_results = results
        diag = calc.diagnostic_report()

        # Mostrar SOLO advertencias sobre puestos inexistentes/no resueltos.
        # Si hay al menos una, cancelar el cálculo de circulación hasta que se corrijan.
        if warnings:
            stand_warnings = [
                w for w in warnings
                if "etiquetas de stand" in w or "puestos en el grafo" in w
            ]
            if stand_warnings:
                msg = "\n".join(stand_warnings)
                QtWidgets.QMessageBox.information(
                    self,
                    "Circulación – puestos no encontrados",
                    msg,
                )
                return

        if not results:
            QtWidgets.QMessageBox.information(
                self,
                "Circulación",
                "No hubo operaciones aplicables para circulación. Revise la configuración y los datos cargados.",
            )
            return
        # En resultados REALES de circulación no adjuntamos el informe sintético
        # de ejemplo con el circuito de 4 puestos; ese informe queda sólo en la
        # funcionalidad de simulación sintética.
        dlg = CirculationResultsDialog(self, results, warnings, diag, None)
        dlg.exec_()

    def open_total_emissions_dialog(self):
        """Calcula y muestra emisiones totales (servicio + circulación) por vehículo.

        Utiliza la misma tabla de operaciones para ambos componentes. Si no hay
        circulación aplicable, muestra solo las emisiones en servicio.
        """
        try:
            # Requiere operaciones reales
            if getattr(self, "_ops_df", None) is None or getattr(self._ops_df, 'empty', False):
                QtWidgets.QMessageBox.information(
                    self,
                    "Operaciones",
                    "Debe cargar un archivo Excel de operaciones compatible antes de calcular.",
                )
                return

            # Validaciones generales de datos antes de cualquier cálculo
            if not self._check_missing_aircraft_for_operations():
                return
            if not self._validate_service_config():
                return

            # Servicio (usando histórico de operaciones)
            serv, svc_hours = self._compute_service_emissions_from_history()
            self._last_emis_servicio = serv

            # Circulación
            dataset = self._ensure_circulation_dataset()
            if dataset is None:
                QtWidgets.QMessageBox.information(
                    self,
                    "Circulación",
                    "Debe cargar o generar una tabla de circulación antes de calcular.",
                )
                return
            calc_circ = CirculationCalculator(
                self.model,
                self.config,
                circ_params=getattr(self, "_circ_params", None),
                sim_params=getattr(self, "_sim_params", None),
                dataset=dataset,
                debug_enabled=getattr(self, "_circ_debug_enabled", False),
            )
            if self._date_filter_from is not None and self._date_filter_to is not None:
                calc_circ.set_date_filter(self._date_filter_from, self._date_filter_to)
            calc_circ.set_operations_df(self._get_active_ops_df())
            circ_results, warnings = calc_circ.compute()
            self._last_circ_results = circ_results

            # Si hay advertencias de puestos inexistentes, bloquear cálculo total
            if warnings:
                stand_warnings = [
                    w for w in warnings
                    if "etiquetas de stand" in w or "puestos en el grafo" in w
                ]
                if stand_warnings:
                    msg = "\n".join(stand_warnings)
                    QtWidgets.QMessageBox.information(
                        self,
                        "Emisiones totales – puestos no encontrados",
                        msg,
                    )
                    return

            gases = ["CO2", "CO", "HC", "NOx", "SOx", "PM10"]
            merged: Dict[str, Dict[str, Dict[str, float]]] = {}

            # Conjunto de vehículos presentes en servicio o circulación
            vehs = set(serv.keys()) | set(circ_results.keys())

            # Construir totales combinados por vehículo y gas
            for veh in vehs:
                merged[veh] = {}
                svc_h = float(svc_hours.get(veh, 0.0) or 0.0)
                circ_time_h = 0.0
                res_circ = circ_results.get(veh)
                if res_circ is not None:
                    try:
                        circ_time_h = float(res_circ.circulation_time_h or 0.0)
                    except Exception:
                        circ_time_h = 0.0
                total_time_s = (svc_h + circ_time_h) * 3600.0

                for gas in gases:
                    s_val = serv.get(veh, {}).get(gas, 0.0)
                    c_val = 0.0
                    if veh in circ_results:
                        rec = circ_results[veh].gases.get(gas)
                        if rec:
                            c_val = rec.get("g", 0.0)
                    total_g = float(s_val or 0.0) + float(c_val or 0.0)
                    total_gps = 0.0
                    if total_time_s > 0.0:
                        total_gps = total_g / total_time_s
                    merged[veh][gas] = {
                        "total": total_g,
                        "total_gps": total_gps,
                    }

            # Flota mínima por vehículo considerando servicio + circulación.
            # Servicio: concurrencia estimada en _compute_service_emissions_from_history
            # Circulación: atributo fleet de cada CirculationResult.
            svc_fleet: Dict[str, int] = getattr(self, "_service_min_fleet", {}) or {}
            min_fleet_sizes: Dict[str, int] = {}
            for veh in vehs:
                try:
                    f_svc = int(svc_fleet.get(veh, 0) or 0)
                except Exception:
                    f_svc = 0
                try:
                    res_circ = circ_results.get(veh)
                    f_circ = int(getattr(res_circ, "fleet", 0) or 0) if res_circ is not None else 0
                except Exception:
                    f_circ = 0
                # La flota mínima total es el máximo entre el pico de servicio
                # y el pico de circulación (no se suman, porque el mismo vehículo
                # podría estar en ambos estados en diferentes momentos).
                # El mayor de los dos determina cuántas unidades se necesitan en paralelo.
                min_fleet_sizes[veh] = max(f_svc, f_circ)

            # Guardar últimos resultados agregados para informes posteriores
            self._last_service_hours = svc_hours
            self._last_total_merged = merged
            self._last_total_min_fleet = min_fleet_sizes

            dlg = TotalEmissionsDialog(self, merged, gases, min_fleet_sizes)
            dlg.exec_()
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Emisiones", str(e))

    def generate_full_emissions_report(self) -> Optional[str]:
        """Genera un informe textual completo (servicio + circulación) con rutas básicas.

        Utiliza los últimos resultados de emisiones calculados en open_total_emissions_dialog.
        Devuelve la ruta del archivo generado o None si no hay datos suficientes.
        """
        try:
            serv: Dict[str, Dict[str, float]] = getattr(self, "_last_emis_servicio", {}) or {}
            svc_hours: Dict[str, float] = getattr(self, "_last_service_hours", {}) or {}
            circ_results: Dict[str, CirculationResult] = getattr(self, "_last_circ_results", {}) or {}
            merged: Dict[str, Dict[str, Dict[str, float]]] = getattr(self, "_last_total_merged", {}) or {}
            min_fleet: Dict[str, int] = getattr(self, "_last_total_min_fleet", {}) or {}
            if not merged:
                return None

            gases = ["CO2", "CO", "HC", "NOx", "SOx", "PM10"]
            veh_order = ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]
            vehs = [v for v in veh_order if v in merged] + [v for v in sorted(merged.keys()) if v not in veh_order]

            lines: List[str] = []
            lines.append("INFORME DE EMISIONES TOTALES (SERVICIO + CIRCULACIÓN)")
            lines.append("=" * 72)
            lines.append("")

            # Resumen global por gas
            lines.append("RESUMEN GLOBAL POR GAS (FLOTA COMPLETA)")
            for gas in gases:
                total_fleet_g = 0.0
                for v in vehs:
                    rec = merged.get(v, {}).get(gas, {}) or {}
                    total_fleet_g += float(rec.get("total", 0.0) or 0.0)
                lines.append(f"  {gas}: {total_fleet_g:.3f} g ({total_fleet_g/1000.0:.3f} kg)")
            lines.append("")

            # Detalle por vehículo
            for v in vehs:
                lines.append("-" * 72)
                human_name = EmissionsCalculator(self.config)._veh_map.get(v, v)
                lines.append(f"VEHÍCULO: {v} – {human_name}")
                fleet_val = 0
                try:
                    fleet_val = int(min_fleet.get(v, 0) or 0)
                except Exception:
                    fleet_val = 0
                lines.append(f"Flota mínima total estimada (FMT): {'—' if fleet_val <= 0 else fleet_val}")

                # Servicio
                Th = float(svc_hours.get(v, 0.0) or 0.0)
                lines.append("")
                lines.append("  SERVICIO (en puesto)")
                if Th <= 0.0 or v not in serv:
                    lines.append("    Sin horas de servicio registradas para este vehículo.")
                else:
                    lines.append(f"    Horas efectivas de servicio: {Th:.3f} h")
                    for gas in gases:
                        m_g = float(serv.get(v, {}).get(gas, 0.0) or 0.0)
                        rate = m_g / (Th * 3600.0) if Th > 0 else 0.0
                        lines.append(f"    {gas}: {m_g:.3f} g ({m_g/1000.0:.3f} kg)  |  media: {rate:.6f} g/s")

                # Circulación
                lines.append("")
                lines.append("  CIRCULACIÓN")
                res_circ = circ_results.get(v)
                if not res_circ:
                    lines.append("    Sin circulación registrada para este vehículo.")
                else:
                    lines.append(f"    Distancia total: {res_circ.distance_km:.3f} km")
                    lines.append(f"    Tiempo total de rodado: {res_circ.time_h:.3f} h")
                    lines.append(f"    Flota mínima en circulación: {res_circ.fleet}")
                    for gas in gases:
                        rec = res_circ.gases.get(gas)
                        if not rec:
                            continue
                        g_val = float(rec.get("g", 0.0) or 0.0)
                        gps_val = float(rec.get("gps", 0.0) or 0.0)
                        lines.append(f"    {gas}: {g_val:.3f} g ({g_val/1000.0:.3f} kg)  |  media: {gps_val:.4f} g/s")

                # Totales combinados
                lines.append("")
                lines.append("  TOTALES COMBINADOS (servicio + circulación)")
                rec_tot = merged.get(v, {}) or {}
                for gas in gases:
                    t_rec = rec_tot.get(gas, {}) or {}
                    total_g = float(t_rec.get("total", 0.0) or 0.0)
                    total_gps = float(t_rec.get("total_gps", 0.0) or 0.0)
                    lines.append(
                        f"    {gas}: {total_g:.3f} g ({total_g/1000.0:.3f} kg)  |  media combinada: {total_gps:.4f} g/s"
                    )
                lines.append("")

            # Escribir informe a archivo en la carpeta Resultados
            out_dir = os.path.join(os.getcwd(), "Resultados")
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception:
                pass
            out_path = os.path.join(out_dir, "informe_emisiones_total.txt")
            try:
                with open(out_path, "w", encoding="utf-8") as fh:
                    fh.write("\n".join(lines))
            except Exception:
                return None
            return out_path
        except Exception:
            return None

    def _write_circ_debug_report(self, text: str) -> Optional[str]:
        """Escribe el informe de depuración de circulación a un archivo de texto.

        Devuelve la ruta del archivo o None si falla o si el texto está vacío.
        """
        if not text:
            return None
        try:
            path = os.path.join(os.getcwd(), "circulation_debug_report.txt")
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(text)
            return path
        except Exception:
            return None

    def _open_circulation_excel(self):
        """Carga una planilla específica de circulación (Excel) y la abre en una tabla editable.

        La estructura se guarda como dataset 'circulacion' en ConfigManager para
        poder reutilizarla en futuros trabajos (y también editarla luego desde
        el menú Datos → Editar circulacion…).
        """
        try:
            if pd is None:
                QtWidgets.QMessageBox.information(
                    self,
                    "Circulación – Excel",
                    "Debe instalar dependencias para Excel:\n pip install pandas openpyxl",
                )
                return

            start = os.path.join(os.getcwd(), "Input") if os.path.isdir(os.path.join(os.getcwd(), "Input")) else os.getcwd()
            fn, _ = QtWidgets.QFileDialog.getOpenFileName(
                self,
                "Seleccione planilla de circulación (Excel)",
                start,
                "Excel (*.xlsx)",
            )
            if not fn:
                return

            df = pd.read_excel(fn, engine="openpyxl")
            cols = [str(c) for c in df.columns]
            rows = []
            for i in range(len(df)):
                rec = {}
                for j, c in enumerate(cols):
                    val = df.iloc[i, j]
                    # Normalizar NaN a cadena vacía para la tabla editable
                    if pd.isna(val):
                        rec[c] = ""
                    else:
                        rec[c] = val
                rows.append(rec)

            # Guardar como dataset 'circulacion' para uso futuro
            self.config.set_dataset("circulacion", cols, rows)
            self.config.save_user_config()

            data = {"columns": cols, "rows": rows}
            dlg = TableEditorDialog(self, "circulacion", data)
            dlg.exec_()
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Circulación – Excel", str(e))

        # Activar botones de circulación y total sólo si también hay operaciones cargadas
        try:
            if getattr(self, "_ops_df", None) is not None and not getattr(self._ops_df, 'empty', False):
                # Emisiones por servicio ahora también requieren la tabla de circulación
                if hasattr(self, "btnCalcEmisServicio"):
                    self.btnCalcEmisServicio.setEnabled(True)
                if hasattr(self, "btnCalcEmisCirculacion"):
                    self.btnCalcEmisCirculacion.setEnabled(True)
                if hasattr(self, "btnCalcEmisTotal"):
                    self.btnCalcEmisTotal.setEnabled(True)
        except Exception:
            pass

    def _open_circulation_editor(self):
        """Abre el editor de la tabla de circulación (dataset 'circulacion')."""
        try:
            data = self.config.get_dataset("circulacion")
        except Exception:
            data = {"columns": [], "rows": []}
        dlg = TableEditorDialog(self, "circulacion", data)
        dlg.exec_()

    def _capture_circulation_from_graph(self):
        if not self.model.nodes:
            QtWidgets.QMessageBox.information(self, "Circulación", "No hay nodos cargados en el grafo.")
            return
        cols = [
            "Categoria", "ID", "Nombre", "Desde", "Hasta", "Dist_km",
            "Sentido", "Es_hub_BUS", "Es_hub_STA", "Es_hub_BAG", "Es_hub_BEL", "Manga"
        ]
        rows: List[dict] = []
        for node in self.model.nodes.values():
            rows.append({
                "Categoria": "Nodo",
                "ID": node.id,
                "Nombre": node.name,
                "Desde": "",
                "Hasta": "",
                "Dist_km": "",
                "Sentido": node.kind,
                "Es_hub_BUS": "sí" if node.is_hub_bus else "no",
                "Es_hub_STA": "sí" if node.is_hub_sta else "no",
                "Es_hub_BAG": "sí" if node.is_hub_bag else "no",
                "Es_hub_BEL": "sí" if node.is_hub_bel else "no",
                "Manga": "sí" if node.has_jetbridge else "no",
            })
        for edge in self.model.edges:
            rows.append({
                "Categoria": "Ruta",
                "ID": edge.id,
                "Nombre": "",
                "Desde": edge.n1,
                "Hasta": edge.n2,
                "Dist_km": f"{edge.length_km:.3f}",
                "Sentido": "Solo ida" if edge.is_one_way else "Doble",
                "Es_hub_BUS": "",
                "Es_hub_STA": "",
                "Es_hub_BAG": "",
                "Es_hub_BEL": "",
                "Manga": "",
            })
        self.config.set_dataset("circulacion", cols, rows)
        self.config.save_user_config()
        data = {"columns": cols, "rows": rows}
        dlg = TableEditorDialog(self, "circulacion", data)
        dlg.exec_()

        # Activar botones de circulación y total sólo si también hay operaciones cargadas
        try:
            if getattr(self, "_ops_df", None) is not None and not getattr(self._ops_df, 'empty', False):
                # Emisiones por servicio ahora también requieren la tabla de circulación
                if hasattr(self, "btnCalcEmisServicio"):
                    self.btnCalcEmisServicio.setEnabled(True)
                if hasattr(self, "btnCalcEmisCirculacion"):
                    self.btnCalcEmisCirculacion.setEnabled(True)
                if hasattr(self, "btnCalcEmisTotal"):
                    self.btnCalcEmisTotal.setEnabled(True)
        except Exception:
            pass

    def _ensure_circulation_dataset(self) -> Optional[dict]:
        """Devuelve la tabla de circulación REAL si existe.

        Si no existe, devuelve None. El escenario sintético se maneja por separado
        en la funcionalidad de simulación sintética.
        """
        try:
            data = self.config.get_dataset("circulacion")
        except Exception:
            data = {"columns": [], "rows": []}
        rows = data.get("rows") or []
        if rows:
            # Considerar dataset como real (no sintético) salvo que se asigne explícitamente
            self._circ_dataset_is_synthetic = False
            return data
        return None

    def _check_missing_aircraft_for_operations(self) -> bool:
        """Verifica que todas las aeronaves del Excel tengan datos en gsexaeronaves.

        Devuelve True si todo está OK. Si encuentra aeronaves sin configurar,
        muestra un aviso indicando dónde completarlas y devuelve False para
        cancelar el cálculo.
        """
        df = getattr(self, "_ops_df", None)
        if df is None:
            return True
        try:
            cols = list(df.columns)
        except Exception:
            return True

        ac_series = None
        # Preferir columna llamada "Aeronave"
        for name in cols:
            if str(name).strip().lower() == "aeronave":
                ac_series = df[name]  # pylint: disable=E1136
                break
        if ac_series is None:
            # Fallback al índice 4 si existe
            if len(cols) > 4:
                try:
                    ac_series = df.iloc[:, 4]
                except Exception:
                    return True
            else:
                return True

        used_ac = set()
        for x in ac_series:
            if x is None:
                continue
            s = str(x).strip().upper()
            if not s:
                continue
            if s.lower() in ("nan", "na", "none", "null"):
                continue
            used_ac.add(s)
        if not used_ac:
            return True

        # Aeronaves conocidas en gsexaeronaves (una cada grupo de 3 columnas)
        try:
            data = self.config.get_dataset("gsexaeronaves")
        except Exception:
            data = {"columns": [], "rows": []}
        cols_gse = list(data.get("columns", []) or [])
        known_ac = set()
        i = 1
        while i < len(cols_gse):
            ac = str(cols_gse[i] or "").strip().upper()
            if ac:
                known_ac.add(ac)
            i += 3

        if not known_ac:
            missing = sorted(used_ac)
        else:
            missing = sorted(ac for ac in used_ac if ac not in known_ac)

        if not missing:
            return True

        lines = [
            "Las siguientes aeronaves aparecen en el Excel de operaciones,",
            "pero no tienen configuración en el dataset 'gsexaeronaves':",
            "",
        ]
        for ac in missing:
            lines.append(f"  - {ac}")
        lines.extend([
            "",
            "Debe agregarlas en:",
            "  Menú Datos → Editar gsexaeronaves…",
            "",
            "Para cada aeronave agregue un grupo de 3 columnas consecutivas:",
            "  CODIGO, CODIGO_S_Rampa, CODIGO_t_hr",
            "y complete las cantidades y tiempos de servicio para cada tipo de GSE.",
            "",
            "El cálculo se cancelará hasta que complete estos datos.",
        ])
        QtWidgets.QMessageBox.warning(
            self,
            "Aeronaves sin configuración de servicio",
            "\n".join(lines),
        )
        return False

    def _validate_service_config(self) -> bool:
        """Verifica que coef_vehiculos y EF tengan filas para todos los vehículos estándar.

        Devuelve True si los parámetros son suficientes. Si faltan datos de
        servicio para algún vehículo, muestra un aviso indicando en qué dataset
        completarlos y devuelve False.
        """
        emis = EmissionsCalculator(self.config)
        if hasattr(self, "_sim_params") and self._sim_params:
            emis.set_overrides(self._sim_params)

        try:
            coef_rows = emis._load_coef_vehiculos()
            ef_rows = emis._load_EF()
        except Exception:
            return True

        vehs = ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]
        missing_msgs: List[str] = []
        for veh in vehs:
            vrow = emis._find_vehicle_row(veh, coef_rows)
            human_name = emis._veh_map.get(veh, veh)
            if not vrow:
                missing_msgs.append(
                    f"- Vehículo {veh} ({human_name}): no hay fila en el dataset 'coef_vehiculos' (columna GSE)."
                )
                continue
            hp_vehicle = emis._vehicle_hp(vrow)
            efrow = emis._select_EF_row_for_veh(veh, hp_vehicle, ef_rows)
            if not efrow:
                comb = (emis._ovr(veh, 'combustible', emis.combustible) or emis.combustible)
                missing_msgs.append(
                    f"- Vehículo {veh} ({human_name}): no se encontró fila en el dataset 'EF' para combustible '{comb}'."
                )

        if not missing_msgs:
            return True

        text = (
            "Faltan datos para completar el cálculo de emisiones en servicio:\n\n"
            + "\n".join(missing_msgs)
            + "\n\nComplete la información en:\n"
            + "  - Menú Datos → Editar coef_vehiculos… (potencia, factores FCD/FC y tiempos t/tD)\n"
            + "  - Menú Datos → Editar EF… (factores de emisión por gas / combustible / rango de HP)\n\n"
            + "El cálculo se cancelará hasta que complete estos datos."
        )
        QtWidgets.QMessageBox.warning(self, "Parámetros de servicio incompletos", text)
        return False

    def _create_synthetic_circulation_dataset(self) -> dict:
        cols = [
            "Categoria", "ID", "Nombre", "Desde", "Hasta", "Dist_km",
            "Sentido", "Es_hub_BUS", "Es_hub_STA", "Es_hub_BAG", "Es_hub_BEL", "Manga", "Lat", "Lon"
        ]
        rows: List[dict] = []
        nodes = [
            ("B0", "BASE", "base", 0.0, 0.0, False, False, False, False),
            ("HB", "hub_BUS", "hub", 0.002, 0.002, True, False, False, False),   # hub BUS
            ("HG", "hub_BAG", "hub", -0.002, -0.002, False, False, True, False), # hub BAG
            # STA/BEL hubs no se usan en el ejemplo sintético pero dejamos columnas en "no"
            ("P1", "P1", "puesto", 0.01, 0.00, False, False, False, False),
            ("P2", "P2", "puesto", 0.015, 0.005, False, False, False, False),
            ("P3", "P3", "puesto", 0.02, -0.002, False, False, False, False),
            ("P4", "P4", "puesto", 0.025, 0.003, False, False, False, False),
        ]
        for nid, name, kind, lat, lon, hub_bus, hub_sta, hub_bag, hub_bel in nodes:
            rows.append({
                "Categoria": "Nodo",
                "ID": nid,
                "Nombre": name,
                "Desde": "",
                "Hasta": "",
                "Dist_km": "",
                "Sentido": kind,
                "Es_hub_BUS": "sí" if hub_bus else "no",
                "Es_hub_STA": "sí" if hub_sta else "no",
                "Es_hub_BAG": "sí" if hub_bag else "no",
                "Es_hub_BEL": "sí" if hub_bel else "no",
                "Manga": "sí" if (kind == "puesto" and name == "P3") else "no",
                "Lat": lat,
                "Lon": lon,
            })
        edges = [
            ("E001", "B0", "P1", 0.45, False),
            ("E002", "P1", "P2", 0.33, False),
            ("E003", "P2", "P3", 0.42, False),
            ("E004", "P3", "P4", 0.55, False),
            ("E005", "P4", "P1", 0.50, False),
            ("E006", "HB", "P2", 0.30, True),
            ("E007", "HG", "P4", 0.28, True),
        ]
        for eid, n1, n2, dist, one_way in edges:
            rows.append({
                "Categoria": "Ruta",
                "ID": eid,
                "Nombre": "",
                "Desde": n1,
                "Hasta": n2,
                "Dist_km": f"{dist:.3f}",
                "Sentido": "Solo ida" if one_way else "Doble",
                "Es_hub_BUS": "",
                "Es_hub_STA": "",
                "Es_hub_BAG": "",
                "Es_hub_BEL": "",
                "Manga": "",
                "Lat": "",
                "Lon": "",
            })
        return {"columns": cols, "rows": rows}

    def _synthetic_operations_data(self) -> list:
        ops = []
        base_time = 7.0
        stands = ["P1", "P3", "P2", "P4", "P1", "P2"]
        durations = [0.3, 0.35, 0.4, 0.45, 0.4, 0.35]
        for idx, stand in enumerate(stands):
            arr = base_time + idx * 0.4
            dep = arr + durations[idx]
            ops.append({
                "Puerta_asignada": stand,
                "Hora_IN_GATE": arr,
                "Hora_OUT_Gate": dep,
                "Aeronave": "A320",
                "TIPO_SER": "SYN",
            })
        return ops

    def open_synthetic_circulation_dialog(self):
        # ... (código existente de simulación sintética)
        pass

    def open_step_by_step_sim(self):
        calc = CirculationCalculator(self.model, self.config, circ_params=self._circ_params, sim_params=self._sim_params)
        if self._ops_df is not None:
            calc.set_operations_df(self._get_active_ops_df())
        dlg = StepByStepSimDialog(self, calc)
        dlg.exec_()
        """Ejecuta una simulación de circulación totalmente sintética.

        No usa self._ops_df ni el dataset 'circulacion' persistido. Genera un
        conjunto de operaciones y una red de circulación de ejemplo y ejecuta
        CirculationCalculator sólo para fines demostrativos.
        """
        try:
            # Crear dataset sintético de circulación (no se guarda en config)
            circ_dataset = self._create_synthetic_circulation_dataset()

            # Crear DataFrame sintético de operaciones si pandas está disponible
            if pd is None:
                QtWidgets.QMessageBox.information(
                    self,
                    "Simulación sintética",
                    "Se requiere pandas para generar operaciones sintéticas.",
                )
                return
            ops_data = self._synthetic_operations_data()
            df_ops = pd.DataFrame(ops_data)

            # Instanciar un CirculationCalculator desconectado del estado real
            calc = CirculationCalculator(
                self.model,
                self.config,
                circ_params=getattr(self, "_circ_params", None),
                sim_params=getattr(self, "_sim_params", None),
                dataset=circ_dataset,
                debug_enabled=True,
            )
            calc.set_operations_df(df_ops)
            results, warnings = calc.compute()
            diag = calc.diagnostic_report()
            debug_report = calc.synthetic_debug_report()
            debug_path = self._write_circ_debug_report(debug_report)

            if not results:
                QtWidgets.QMessageBox.information(
                    self,
                    "Simulación sintética",
                    "La simulación sintética no produjo resultados. Revise la configuración interna.",
                )
                return

            dlg = SyntheticCirculationDialog(self, results, warnings, diag, debug_path)
            dlg.exec_()
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Simulación sintética", str(e))

    def _load_operations_excel(self):
        try:
            if pd is None:
                QtWidgets.QMessageBox.information(self, "Operaciones", "Debe instalar dependencias para Excel:\n pip install pandas openpyxl")
                return
            start = os.path.join(os.getcwd(), "Input") if os.path.isdir(os.path.join(os.getcwd(), "Input")) else os.getcwd()
            # Permitir cualquier nombre .xlsx
            fn, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Seleccione tabla de operaciones (Excel)", start, "Excel (*.xlsx)")
            if not fn:
                return
            df = self._read_ops_excel(fn)
            # Aceptar cualquier distribución de columnas; el usuario verificará visualmente si lo desea
            self._date_filter_from = None
            self._date_filter_to = None
            self._ops_df = df
            if hasattr(self, 'btnPreviewOps'):
                self.btnPreviewOps.setEnabled(True)
            
            # Poblar filtro de aeronaves
            if hasattr(self, 'cmbFilterAircraft'):
                self.cmbFilterAircraft.blockSignals(True)
                self.cmbFilterAircraft.clear()
                self.cmbFilterAircraft.addItem("Todas las aeronaves")
                if "Aeronave" in df.columns:
                    autos = sorted(list(set([str(x).strip() for x in df["Aeronave"].dropna() if str(x).strip()])))
                    for a in autos:
                        self.cmbFilterAircraft.addItem(a)
                self.cmbFilterAircraft.blockSignals(False)

            # Sólo habilitar cálculos cuando también exista una tabla REAL de
            # circulación (para conocer mangas y hubs por puesto).
            has_circ = False
            try:
                circ = self._ensure_circulation_dataset()
                has_circ = circ is not None
            except Exception:
                has_circ = False

            # Emisiones en servicio requieren operaciones + circulación
            if hasattr(self, "btnCalcEmisServicio"):
                self.btnCalcEmisServicio.setEnabled(has_circ)

            # Si ya existe una tabla de circulación, habilitar también circulación y total
            if has_circ:
                if hasattr(self, "btnCalcEmisCirculacion"):
                    self.btnCalcEmisCirculacion.setEnabled(True)
                if hasattr(self, "btnCalcEmisTotal"):
                    self.btnCalcEmisTotal.setEnabled(True)
            QtWidgets.QMessageBox.information(self, "Operaciones", f"Operaciones cargadas:\n{os.path.basename(fn)}")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Operaciones", str(e))

    def _get_active_ops_df(self):
        """Devuelve las operaciones filtradas según la selección del usuario."""
        df = self._ops_df
        if df is None or df.empty or not hasattr(self, 'cmbFilterAircraft'):
            return df
        sel = self.cmbFilterAircraft.currentText()
        if sel == "Todas las aeronaves" or not sel:
            return df
        if "Aeronave" in df.columns:
            return df[df["Aeronave"].astype(str).str.strip() == sel.strip()].copy()
        return df

    def _read_ops_excel(self, path: str):
        # lectura simple (primera hoja) respetando encabezados
        return pd.read_excel(path, engine='openpyxl') if pd is not None else None

    def _validate_ops_df(self, df) -> bool:
        # Mantener por si se usa en el futuro, pero actualmente no se invoca
        try:
            return True
        except Exception:
            return False

    def _vehicle_codes(self) -> List[str]:
        return ["GPU","CAT","TUG","BAG","BEL","WAT","BRE","LAV","FUE","STA","BUS","CLE"]

    def _open_sim_params(self):
        # Diálogo profesional para parámetros por vehículo
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Parámetros de simulación – Emisiones en servicio")
        dlg.resize(780, 420)
        v = QtWidgets.QVBoxLayout(dlg)

        # Cargar existentes (si los hay)
        existing = getattr(self, '_sim_params', {}) or {}
        dflt = existing.get('default', {}) or {}
        veh_over = existing.get('veh', {}) or {}

        # Default params
        boxDefault = QtWidgets.QGroupBox("Valores por defecto (se aplican donde no haya override)")
        g = QtWidgets.QGridLayout(boxDefault)
        spAge = QtWidgets.QDoubleSpinBox(); spAge.setRange(0, 50); spAge.setDecimals(1); spAge.setSuffix(" años")
        spTU = QtWidgets.QDoubleSpinBox(); spTU.setRange(0.1, 50); spTU.setDecimals(1); spTU.setSuffix(" años")
        cbComb = QtWidgets.QComboBox(); cbComb.addItems(["Diesel","Gasolina","Nafta","GNC","Otro"]) 
        spEFHP = QtWidgets.QDoubleSpinBox(); spEFHP.setRange(1, 2000); spEFHP.setSuffix(" HP max (EF)")

        # Prefill defaults si existen
        try:
            spAge.setValue(float(dflt.get('age', 8.0)))
        except Exception:
            spAge.setValue(8.0)
        try:
            spTU.setValue(float(dflt.get('t_util', 10.0)))
        except Exception:
            spTU.setValue(10.0)
        cbComb.setCurrentText(str(dflt.get('combustible', 'Diesel')))
        try:
            spEFHP.setValue(float(dflt.get('ef_hp_max', 175.0)))
        except Exception:
            spEFHP.setValue(175.0)

        g.addWidget(QtWidgets.QLabel("Antigüedad (age)"), 0,0); g.addWidget(spAge, 0,1)
        g.addWidget(QtWidgets.QLabel("Tiempo de utilización (t_util)"), 0,2); g.addWidget(spTU, 0,3)
        g.addWidget(QtWidgets.QLabel("Combustible (EF)"), 1,0); g.addWidget(cbComb, 1,1)
        g.addWidget(QtWidgets.QLabel("EF – HP max preferido"), 1,2); g.addWidget(spEFHP, 1,3)

        # Tabla por vehículo
        tbl = QtWidgets.QTableWidget(); tbl.setColumnCount(6)
        tbl.setHorizontalHeaderLabels(["Vehículo","Override","Edad","t_util","EF HP max","HP vehículo (opcional)"])
        vehs = self._vehicle_codes()
        tbl.setRowCount(len(vehs))
        for i, veh in enumerate(vehs):
            tbl.setItem(i, 0, QtWidgets.QTableWidgetItem(veh))
            chk = QtWidgets.QTableWidgetItem(); chk.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled); chk.setCheckState(QtCore.Qt.Unchecked)
            tbl.setItem(i, 1, chk)
            age = QtWidgets.QTableWidgetItem(""); tu = QtWidgets.QTableWidgetItem(""); efhp = QtWidgets.QTableWidgetItem(""); hpv = QtWidgets.QTableWidgetItem("")
            tbl.setItem(i, 2, age); tbl.setItem(i,3, tu); tbl.setItem(i,4, efhp); tbl.setItem(i,5, hpv)
            # Prefill por-veh si existe
            if veh in veh_over:
                rec = veh_over.get(veh, {}) or {}
                chk.setCheckState(QtCore.Qt.Checked)
                def set_if_not_none(item, key):
                    val = rec.get(key, None)
                    if val is not None and item is not None:
                        item.setText(str(val))
                set_if_not_none(age, 'age')
                set_if_not_none(tu, 't_util')
                set_if_not_none(efhp, 'ef_hp_max')
                set_if_not_none(hpv, 'hp_vehicle')
        tbl.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        v.addWidget(boxDefault)
        v.addWidget(QtWidgets.QLabel("Overrides por vehículo (dejar vacío para usar defecto)"))
        v.addWidget(tbl)

        # Botones
        btns = QtWidgets.QHBoxLayout()
        btnApply = QtWidgets.QPushButton("Aplicar")
        btnClose = QtWidgets.QPushButton("Cerrar")
        btns.addStretch(1); btns.addWidget(btnApply); btns.addWidget(btnClose)
        v.addLayout(btns)

        def on_apply():
            overrides = {
                'default': {
                    'age': float(spAge.value()),
                    't_util': float(spTU.value()),
                    'combustible': cbComb.currentText(),
                    'ef_hp_max': float(spEFHP.value()),
                },
                'veh': {}
            }
            for i, veh in enumerate(vehs):
                if tbl.item(i,1).checkState() == QtCore.Qt.Checked:
                    rec = {}
                    def getf(j):
                        try:
                            val = tbl.item(i,j).text().strip()
                            return float(val) if val != '' else None
                        except Exception:
                            return None
                    rec['age'] = getf(2)
                    rec['t_util'] = getf(3)
                    rec['ef_hp_max'] = getf(4)
                    rec['hp_vehicle'] = getf(5)
                    rec['combustible'] = cbComb.currentText()  # combustible igual por ahora; se puede abrir por-veh si se necesita
                    overrides['veh'][veh] = rec
            self._sim_params = overrides
            QtWidgets.QMessageBox.information(self, "Parámetros", "Parámetros de simulación guardados.")

        btnApply.clicked.connect(on_apply)
        btnClose.clicked.connect(dlg.accept)
        dlg.exec_()

    def _open_date_filter_dialog(self):
        if getattr(self, "_ops_df", None) is None or getattr(self._ops_df, "empty", False):
            QtWidgets.QMessageBox.information(
                self,
                "Filtro de fechas",
                "Debe cargar primero un archivo de operaciones (Excel).",
            )
            return

        df = self._ops_df
        try:
            cols = list(df.columns)
        except Exception:
            QtWidgets.QMessageBox.warning(
                self,
                "Filtro de fechas",
                "No se pudo leer la tabla de operaciones para determinar el rango de fechas.",
            )
            return
        if not cols:
            QtWidgets.QMessageBox.warning(
                self,
                "Filtro de fechas",
                "La tabla de operaciones no tiene columnas.",
            )
            return

        try:
            series = df.iloc[:, 0]
        except Exception:
            QtWidgets.QMessageBox.warning(
                self,
                "Filtro de fechas",
                "No se pudo acceder a la columna de fechas (DIA).",
            )
            return

        min_date: Optional[datetime.date] = None
        max_date: Optional[datetime.date] = None
        for val in series:
            d: Optional[datetime.date] = None
            if isinstance(val, datetime.datetime):
                d = val.date()
            elif isinstance(val, datetime.date):
                d = val
            elif isinstance(val, str):
                s = val.strip()
                if s:
                    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                        try:
                            d = datetime.datetime.strptime(s, fmt).date()
                            break
                        except Exception:
                            continue
            if d is None:
                continue
            if min_date is None or d < min_date:
                min_date = d
            if max_date is None or d > max_date:
                max_date = d

        if min_date is None or max_date is None:
            QtWidgets.QMessageBox.warning(
                self,
                "Filtro de fechas",
                "No se encontraron fechas válidas en la columna DIA del Excel de operaciones.",
            )
            return

        dlg = DateRangeDialog(self, min_date, max_date)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return
        date_from, date_to = dlg.get_selected_range()
        if date_from > date_to:
            QtWidgets.QMessageBox.warning(
                self,
                "Filtro de fechas",
                "La fecha 'Desde' no puede ser posterior a la fecha 'Hasta'.",
            )
            return

        try:
            delta_days = (date_to - date_from).days
        except Exception:
            delta_days = 0
        max_days = 5 * 366
        if delta_days > max_days:
            QtWidgets.QMessageBox.warning(
                self,
                "Filtro de fechas",
                "El rango máximo permitido es de 5 años.",
            )
            return

        self._date_filter_from = date_from
        self._date_filter_to = date_to
        QtWidgets.QMessageBox.information(
            self,
            "Filtro de fechas",
            f"Filtro de fechas aplicado:\nDesde {date_from} hasta {date_to}.",
        )

    def _circ_defaults(self, code: str) -> dict[str, float]:
        """Valores por defecto para parámetros de circulación de un vehículo."""
        return {"fc_cir": 0.5, "vel_kmh": 15.0}

    def _open_circ_params(self):
        """Diálogo para editar FC_CIR y velocidad de circulación por GSE."""
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Parámetros de circulación – GSE")
        dlg.resize(640, 360)
        v = QtWidgets.QVBoxLayout(dlg)

        info = QtWidgets.QLabel(
            "FC_CIR: factor de carga en circulación (sin unidad, típico 0.5).\n"
            "Velocidad: velocidad media de circulación del vehículo [km/h]."
        )
        info.setWordWrap(True)
        v.addWidget(info)

        vehs = self._vehicle_codes()
        table = QtWidgets.QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Vehículo", "FC_CIR", "Velocidad [km/h]"])
        table.setRowCount(len(vehs))
        circ_cfg = getattr(self, "_circ_params", {}) or {}

        for row, veh in enumerate(vehs):
            table.setItem(row, 0, QtWidgets.QTableWidgetItem(veh))
            table.item(row, 0).setFlags(QtCore.Qt.ItemIsEnabled)
            params = circ_cfg.get(veh, self._circ_defaults(veh))
            fc_item = QtWidgets.QTableWidgetItem(f"{params.get('fc_cir', 0.5):.3f}")
            vel_item = QtWidgets.QTableWidgetItem(f"{params.get('vel_kmh', 15.0):.2f}")
            table.setItem(row, 1, fc_item)
            table.setItem(row, 2, vel_item)

        table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        v.addWidget(table)

        btns = QtWidgets.QHBoxLayout()
        btns.addStretch(1)
        btnApply = QtWidgets.QPushButton("Guardar")
        btnClose = QtWidgets.QPushButton("Cerrar")
        btns.addWidget(btnApply)
        btns.addWidget(btnClose)
        v.addLayout(btns)

        def on_apply():
            new_cfg: dict[str, dict[str, float]] = {}
            for row, veh in enumerate(vehs):
                try:
                    fc = float(table.item(row, 1).text()) if table.item(row, 1) else 0.5
                except Exception:
                    fc = 0.5
                try:
                    vel = float(table.item(row, 2).text()) if table.item(row, 2) else 15.0
                except Exception:
                    vel = 15.0
                fc = max(0.01, fc)
                vel = max(0.1, vel)
                new_cfg[veh] = {"fc_cir": fc, "vel_kmh": vel}
            self._circ_params = new_cfg
            QtWidgets.QMessageBox.information(self, "Parámetros", "Parámetros de circulación guardados.")

        btnApply.clicked.connect(on_apply)
        btnClose.clicked.connect(dlg.accept)
        dlg.exec_()



    def _preview_operations_excel(self):
        if getattr(self, "_ops_df", None) is None:
            QtWidgets.QMessageBox.information(self, "Operaciones", "Cargue primero un archivo para previsualizar.")
            return
        # Mostrar primeras 10 filas en un diálogo compacto
        df = self._ops_df.head(10)
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Previsualización – Operaciones (solo 10 primeras filas)")
        dlg.resize(900, 360)
        lay = QtWidgets.QVBoxLayout(dlg)
        # Atajo: cerrar con Esc
        esc = QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Escape), dlg)
        esc.activated.connect(dlg.accept)
        table = QtWidgets.QTableWidget(dlg)
        table.setColumnCount(len(df.columns))
        table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        table.setRowCount(len(df))
        for i in range(len(df)):
            for j, c in enumerate(df.columns):
                val = df.iloc[i, j]
                it = QtWidgets.QTableWidgetItem(str(val))
                table.setItem(i, j, it)
        table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        table.setFocus()
        btnClose = QtWidgets.QPushButton("Cerrar")
        btnClose.setMinimumWidth(110)
        btnClose.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        btnClose.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        lay.addWidget(table)
        # Fila de acciones compacta: botón alineado a la derecha
        actions = QtWidgets.QHBoxLayout()
        actions.addStretch(1)
        actions.addWidget(btnClose)
        lay.addLayout(actions)
        btnClose.clicked.connect(dlg.accept)
        dlg.exec_()

    def _import_config_from_excel(self):
        if pd is None:
            QtWidgets.QMessageBox.warning(self, "Error", "Se requiere pandas para importar desde Excel.")
            return

        fn, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Importar Configuración desde Excel", "", "Excel (*.xlsx *.xls)")
        if not fn: return

        # Preguntar qué queremos importar
        items = ["GSE por Aeronave", "Factores de Emisión (EF)", "Coeficientes de Vehículos", "Tipos de Vuelo"]
        item, ok = QtWidgets.QInputDialog.getItem(self, "Seleccionar Dataset", "¿Qué desea importar desde este archivo?", items, 0, False)
        if not ok: return

        key_map = {
            "GSE por Aeronave": "gsexaeronaves",
            "Factores de Emisión (EF)": "EF",
            "Coeficientes de Vehículos": "coef_vehiculos",
            "Tipos de Vuelo": "tipos_vuelo"
        }
        key = key_map[item]

        try:
            import pandas as pd
            df = pd.read_excel(fn)
            # Reemplazar NaNs por cadenas vacías para que el editor de tablas no falle
            df = df.fillna("")
            cols = [str(c) for c in df.columns]
            rows = []
            for i in range(len(df)):
                row_rec = {}
                for c in df.columns:
                    row_rec[str(c)] = df.iloc[i][c]
                rows.append(row_rec)
            
            self.config.set_dataset(key, cols, rows)
            self.config.save_user_config()
            QtWidgets.QMessageBox.information(self, "Éxito", f"Se han importado {len(rows)} filas al dataset '{key}'.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Error al leer el archivo:\n{e}")


# ------------------------------ SplashScreen ------------------------------ #

class GSESplashScreen(QtWidgets.QSplashScreen):
    def __init__(self):
        # Crear un mapa de bits base azul oscuro elegante con gradiente
        pixmap = QtGui.QPixmap(600, 340)
        pixmap.fill(QtCore.Qt.transparent)
        
        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        
        # Fondo redondeado con gradiente
        rect = QtCore.QRectF(0, 0, 600, 340)
        gradient = QtGui.QLinearGradient(0, 0, 0, 340)
        gradient.setColorAt(0, QtGui.QColor("#0a192f"))
        gradient.setColorAt(1, QtGui.QColor("#112240"))
        
        path = QtGui.QPainterPath()
        path.addRoundedRect(rect, 10, 10)
        painter.fillPath(path, QtGui.QBrush(gradient))
        
        # Borde sutil
        pen = QtGui.QPen(QtGui.QColor("#4b6cb7"), 2)
        painter.setPen(pen)
        painter.drawPath(path)
        
        # Dibujar titulo principal
        painter.setPen(QtGui.QColor("#64ffda"))
        font_title = QtGui.QFont("Segoe UI", 28, QtGui.QFont.Bold)
        painter.setFont(font_title)
        painter.drawText(QtCore.QRectF(0, 40, 600, 50), QtCore.Qt.AlignCenter, "GSEQuant PRO")
        
        # Subtitulo
        painter.setPen(QtGui.QColor("#8892b0"))
        font_sub = QtGui.QFont("Segoe UI", 12)
        painter.setFont(font_sub)
        painter.drawText(QtCore.QRectF(0, 95, 600, 30), QtCore.Qt.AlignCenter, "Ground Support Equipment Quantifier")
        
        # Dibujar logo del GTA si existe
        if os.path.exists(resource_path("gta.png")):
            # Lo hacemos considerablemente mas grande (el doble, 120px)
            logo = QtGui.QPixmap(resource_path("gta.png")).scaled(120, 120, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
            # Centrado en X: (600 - 120)/2 = 240, bajando un poco el Y=135
            painter.drawPixmap(240, 135, logo)        
        # Copyright info
        painter.setPen(QtGui.QColor("#ccd6f6"))
        font_copy = QtGui.QFont("Segoe UI", 10, QtGui.QFont.Bold)
        painter.setFont(font_copy)
        painter.drawText(QtCore.QRectF(0, 240, 600, 20), QtCore.Qt.AlignCenter, "Área de Desarrollo de Software")
        
        painter.setPen(QtGui.QColor("#8892b0"))
        font_copy.setBold(False)
        painter.setFont(font_copy)
        painter.drawText(QtCore.QRectF(0, 260, 600, 20), QtCore.Qt.AlignCenter, "Grupo de Transporte Aéreo (GTA)")
        
        painter.setPen(QtGui.QColor("#4b6cb7"))
        painter.drawText(QtCore.QRectF(0, 280, 600, 20), QtCore.Qt.AlignCenter, "https://gta.ing.unlp.edu.ar/")
        
        # Indicador de estado
        painter.setPen(QtGui.QColor("#a8b2d1"))
        font_status = QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold)
        painter.setFont(font_status)
        painter.drawText(QtCore.QRectF(20, 310, 560, 20), QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, "Iniciando módulos del sistema...")
        
        painter.end()
        
        super().__init__(pixmap, QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.FramelessWindowHint)

# ------------------------------ main ------------------------------ #

def main():
    import sys

    # ── Ícono en barra de tareas de Windows ──────────────────
    # Debe hacerse ANTES de crear QApplication
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            u'GTA.GSEQuant.PRO.1'
        )
    except Exception:
        pass  # No es Windows o no está disponible

    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("GSEQuant PRO")
    app.setOrganizationName("GTA")

    # Ícono de la aplicación (aparece en taskbar y alt-tab)
    _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), resource_path("gse_app_icon.png"))
    if os.path.exists(_icon_path):
        app.setWindowIcon(QtGui.QIcon(_icon_path))

    import time
    
    splash = GSESplashScreen()
    splash.show()
    
    # Simular carga
    for msg, delay in [("Cargando motor geográfico Leaflet...", 0.6), 
                       ("Validando parámetros termodinámicos...", 0.6), 
                       ("Cargando base de datos global de aeronaves...", 0.7), 
                       ("Inicializando motor de routing...", 0.6),
                       ("Iniciando interfaz GSEQuant PRO...", 0.5)]:
        splash.showMessage(msg, QtCore.Qt.AlignLeft | QtCore.Qt.AlignBottom, QtGui.QColor("#a8b2d1"))
        app.processEvents()
        time.sleep(delay)
    
    win = MainWindow()
    win.show()
    # Para que desaparezca suavemente (fade out simple via timer en qt, pero hide() alcanza para que sea rápido y limpio)
    splash.finish(win)
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

