# -*- coding: utf-8 -*-
"""
oozt_recs.py — подготовка природоохранных рекомендаций (мероприятий) для ООЗТ Москвы
по кварталам и выделам на основе ГИС-слоёв и таблицы «Мероприятия по Приказу».

Запуск
------
    python oozt_recs.py                   # диалоговый режим (tkinter) — все пути выбираются в окнах
    python oozt_recs.py --config cfg.json # повторный запуск с сохранённой конфигурацией, без диалогов
    python oozt_recs.py --config cfg.json --kv 1,2,5   # отчёт только по выбранным кварталам

Выходные данные (папка out_dir)
-------------------------------
    measures_long.xlsx / .csv   — длинная таблица ООЗТ–Квартал–Выдел–Мероприятие (+ основание)
    measures_by_vydel.xlsx      — сводка: одна строка на выдел, мероприятия через «; »
    evidence.gpkg               — слои-основания (точки/линии/полигоны) + выделы с мероприятиями
    report/report.md            — отчёт по кварталам (схема + таблица), рисунки report/fig_kv_*.png
    config_used.json            — конфигурация запуска (можно передать в --config)

Правила (ключ Метод выявления в Excel → расчёт) описаны в разделе RULES.
Все входы, кроме таблицы мероприятий и слоя ООЗТ, опциональны.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import tempfile
import warnings
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Iterable

# ---- локальная папка pylib рядом со скриптом (чистые Python-пакеты: openpyxl, contextily, python-docx) ----
_pylib = Path(__file__).resolve().parent / "pylib"
if _pylib.is_dir() and str(_pylib) not in sys.path:
    sys.path.append(str(_pylib))          # append: пакеты окружения имеют приоритет

import faulthandler
faulthandler.enable()                     # аварийный стек в stderr при падении C-расширений

# ---- окружение GDAL/matplotlib (conda на Windows часто не выставляет GDAL_DATA) ----
_env_root = Path(sys.prefix)
# Windows: при запуске python.exe без activate папки с DLL окружения не попадают в PATH,
# и C-расширения matplotlib падают с 0xc06d007f (delay-load DLL). Добавляем их сами.
def enable_win_dll_path():
    """Windows: python.exe без `conda activate` не видит DLL окружения, и C-расширения matplotlib
    падают с 0xc06d007f (delay-load DLL). Включаем PATH окружения только перед отрисовкой:
    делать это на старте нельзя — меняется порядок загрузки DLL для GDAL."""
    if os.name != "nt" or os.environ.get("_OOZT_DLLPATH"):
        return
    dirs = [str(p) for p in (_env_root, _env_root / "Library" / "bin",
                             _env_root / "Library" / "mingw-w64" / "bin",
                             _env_root / "Library" / "usr" / "bin",
                             _env_root / "DLLs", _env_root / "Scripts") if p.is_dir()]
    for d in dirs:
        try:
            os.add_dll_directory(d)
        except OSError:
            pass
    os.environ["PATH"] = os.pathsep.join(dirs + [os.environ.get("PATH", "")])
    os.environ["_OOZT_DLLPATH"] = "1"
for _cand in (_env_root / "Library" / "share" / "gdal", _env_root / "share" / "gdal"):
    if _cand.exists() and not os.environ.get("GDAL_DATA"):
        os.environ["GDAL_DATA"] = str(_cand)
for _cand in (_env_root / "Library" / "share" / "proj", _env_root / "share" / "proj"):
    if _cand.exists() and not os.environ.get("PROJ_LIB"):
        os.environ["PROJ_LIB"] = str(_cand)
# Windows/Python ≥ 3.12.4: tempfile.mkdtemp создаёт папку с ограниченным ACL; под сервисными учётными записями
# это ломает matplotlib/joblib/contextily. Заменяем на обычный os.makedirs.
if os.name == "nt":
    import tempfile as _tf
    import uuid as _uuid

    def _mkdtemp(suffix=None, prefix=None, dir=None):
        base = dir or _tf.gettempdir()
        for _ in range(20):
            p = os.path.join(base, (prefix or "tmp") + _uuid.uuid4().hex[:10] + (suffix or ""))
            if not os.path.exists(p):
                os.makedirs(p)
                return p
        raise FileExistsError("не удалось создать временную папку")
    _tf.mkdtemp = _mkdtemp

os.environ.setdefault("MPLCONFIGDIR", os.path.join(os.environ.get("TEMP", "."), "mpl_cache"))
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
os.environ.setdefault("SHAPE_RESTORE_SHX", "YES")  # восстанавливать .shx, если файл отсутствует/повреждён
try:
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except Exception:
    pass

warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import geopandas as gpd
import pyogrio
import rasterio
from rasterio import features as rfeatures
from rasterio.warp import reproject, Resampling, calculate_default_transform
from rasterio.transform import from_origin
from shapely.geometry import shape, box, Point
from shapely.ops import unary_union
from scipy import ndimage

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D

log = logging.getLogger("oozt")
logging.getLogger("pyogrio").setLevel(logging.WARNING)
logging.getLogger("rasterio").setLevel(logging.WARNING)

# =====================================================================================
#                                    ПАРАМЕТРЫ
# =====================================================================================
WORK_EPSG = 32637          # рабочая проекция (WGS84 / UTM 37N, метры)
CELL = 10.0                # шаг сетки для растров плотности/уклона, м
TRAIL_WINDOW_M = 50.0      # радиус скользящего окна плотности троп, м
CORVUS_WINDOW_M = 500.0    # радиус окна плотности врановых, м
RARE_BUFFER_M = 50.0       # буфер от точек редких видов, м
SHORE_BUFFER_M = 50.0      # буфер вдоль водных объектов для «укрепления берегов», м
SDM_THRESHOLD = 0.75       # порог вероятности SDM (по умолчанию, редкие виды)
SDM_THRESHOLD_INVASIVE = 0.95   # порог вероятности SDM для инвазионных видов
SDM_THRESHOLD_RARE = SDM_THRESHOLD  # порог вероятности SDM для редких видов
MIN_PATCH_HA = 0.05        # минимальная площадь пятна (уклон/плотность/SDM) в выделе для назначения, га
QUANTILE = 0.75            # «верхний квартиль»

# Типы дорог OSM, считающиеся дорожно-тропиночной сетью (если в слое есть поле HIGHWAY/fclass)
TRAIL_HIGHWAY_TYPES = {"path", "footway", "track", "cycleway", "steps", "bridleway",
                       "pedestrian", "living_street", "service", "unclassified"}

TYPE_FOREST = ("лес",)
TYPE_MEADOW = ("луг",)
TYPE_WATER = ("водн",)          # «Водные»
TYPE_NEARWATER = ("околовод", "болот")

# =====================================================================================
#                                   КОНФИГУРАЦИЯ
# =====================================================================================
@dataclass
class Config:
    measures_xlsx: str = ""
    oozt_path: str = ""
    oozt_layer: str | None = None
    dem_path: str | None = None
    trails: list[str] = field(default_factory=list)
    hydro_lines: list[str] = field(default_factory=list)
    hydro_polys: list[str] = field(default_factory=list)
    invasive_points: list[str] = field(default_factory=list)
    invasive_polys: list[str] = field(default_factory=list)
    invasive_rasters: list[str] = field(default_factory=list)
    rare_points: list[str] = field(default_factory=list)
    rare_polys: list[str] = field(default_factory=list)
    rare_rasters: list[str] = field(default_factory=list)
    fauna_points: list[str] = field(default_factory=list)     # птицы и др. (для Corvus)
    out_dir: str = ""
    report_kv: list[int] = field(default_factory=list)        # пусто = все кварталы
    basemap: bool = True
    fields: dict = field(default_factory=dict)                # {'OOZT':..,'KV':..,'VYD':..,'TYPE':..,'NAME':..}

    def save(self, path: str | Path):
        Path(path).write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")

    @staticmethod
    def load(path: str | Path) -> "Config":
        d = json.loads(Path(path).read_text(encoding="utf-8"))
        return Config(**d)


# =====================================================================================
#                               ДИАЛОГИ ВЫБОРА ФАЙЛОВ (tkinter)
# =====================================================================================
VEC_TYPES = [("Векторные данные", "*.shp *.gpkg *.geojson *.json *.kml *.dxf *.gml"), ("Все файлы", "*.*")]
RAS_TYPES = [("Растры", "*.tif *.tiff *.asc *.img *.vrt"), ("Все файлы", "*.*")]
XLS_TYPES = [("Excel", "*.xlsx *.xlsm *.xls"), ("Все файлы", "*.*")]


def gui_collect_config() -> Config:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    cfg = Config()

    def one(title, types, required=False):
        while True:
            p = filedialog.askopenfilename(title=title, filetypes=types)
            if p or not required:
                return p or None
            messagebox.showwarning("Обязательный файл", f"Нужно выбрать: {title}")

    def many(title, types):
        ps = filedialog.askopenfilenames(title=title + "  (можно несколько; Отмена = пропустить)", filetypes=types)
        return list(ps) if ps else []

    cfg.measures_xlsx = one("1/12  Таблица мероприятий (Excel «Мероприятия по Приказу»)", XLS_TYPES, required=True)
    cfg.oozt_path = one("2/12  Слой ООЗТ с кварталами и выделами", VEC_TYPES, required=True)
    if cfg.oozt_path.lower().endswith(".gpkg"):
        layers = [l[0] for l in pyogrio.list_layers(cfg.oozt_path)]
        if len(layers) > 1:
            cfg.oozt_layer = simpledialog.askstring("Слой GPKG", "Имя слоя ООЗТ:\n" + "\n".join(layers),
                                                    initialvalue=layers[0])
        else:
            cfg.oozt_layer = layers[0]
    cfg.dem_path = one("3/12  ЦМР (tif, один слой высот) — Отмена = пропустить", RAS_TYPES)
    cfg.trails = many("4/12  Дорожно-тропиночная сеть (линии)", VEC_TYPES)
    cfg.hydro_lines = many("5/12  Гидрология — линии", VEC_TYPES)
    cfg.hydro_polys = many("6/12  Гидрология — полигоны", VEC_TYPES)
    cfg.invasive_points = many("7/12  Инвазионные виды — точки встреч", VEC_TYPES)
    cfg.invasive_polys = many("8/12  Инвазионные виды — полигоны произрастания", VEC_TYPES)
    cfg.invasive_rasters = many("9/12  Инвазионные виды — растры SDM (asc/tif)", RAS_TYPES)
    cfg.rare_points = many("10/12  Редкие виды — точки встреч", VEC_TYPES)
    cfg.rare_polys = many("10/12  Редкие виды — полигоны", VEC_TYPES)
    cfg.rare_rasters = many("10/12  Редкие виды — растры SDM (asc/tif)", RAS_TYPES)
    cfg.fauna_points = many("11/12  Фауна — точки встреч (птицы: для Corvus)", VEC_TYPES)
    out = filedialog.askdirectory(title="12/12  Папка для результатов")
    cfg.out_dir = out or str(Path(cfg.oozt_path).parent / "oozt_recs_out")
    kv = simpledialog.askstring("Кварталы для отчёта",
                                "Номера кварталов через запятую (пусто = все кварталы):") or ""
    cfg.report_kv = [int(x) for x in re.findall(r"\d+", kv)]
    cfg.basemap = messagebox.askyesno("Подложка", "Добавлять подложку OpenStreetMap на схемы? (нужен интернет)")
    root.destroy()
    return cfg


# =====================================================================================
#                                 ЗАГРУЗКА ДАННЫХ
# =====================================================================================
def guess_crs_from_bounds(b) -> int | None:
    """Эвристика CRS по координатам, если она не задана или задана неверно."""
    minx, miny, maxx, maxy = b
    if -180 <= minx <= 180 and -90 <= miny <= 90 and maxx <= 180 and maxy <= 90:
        return 4326
    if 1e5 < minx < 9e5 and 5.5e6 < miny < 7e6:
        return WORK_EPSG          # UTM 37N (Москва ≈ 350–470 тыс. E, 6.1–6.3 млн N)
    if 3e6 < minx < 5e6 and 7e6 < miny < 8e6:
        return 3857               # Web Mercator (Москва ≈ 4.1–4.3 млн E, 7.4–7.6 млн N)
    return None


def fix_crs(obj_crs, bounds, what: str):
    """Возвращает CRS, исправленную по границам (частая ошибка: tif «в UTM», а координаты Web Mercator)."""
    guess = guess_crs_from_bounds(bounds)
    declared = None
    try:
        declared = obj_crs.to_epsg() if obj_crs else None
    except Exception:
        pass
    if declared is None:
        if guess is None:
            log.warning("%s: CRS не задана и не распознана по границам — считаем EPSG:%s", what, WORK_EPSG)
            return rasterio.crs.CRS.from_epsg(WORK_EPSG)
        log.warning("%s: CRS не задана, по границам принята EPSG:%s", what, guess)
        return rasterio.crs.CRS.from_epsg(guess)
    if guess is not None and guess != declared and declared in (4326, 3857, WORK_EPSG):
        log.warning("%s: объявлена EPSG:%s, но координаты соответствуют EPSG:%s — исправлено", what, declared, guess)
        return rasterio.crs.CRS.from_epsg(guess)
    return obj_crs


def read_vector(path: str, layer: str | None = None, bbox_wgs84=None) -> gpd.GeoDataFrame:
    """Чтение вектора любого формата с приведением к WORK_EPSG. bbox_wgs84 — (minx,miny,maxx,maxy) в EPSG:4326."""
    kw = {}
    if layer:
        kw["layer"] = layer
    # bbox в CRS файла
    src_crs = None
    try:
        info = pyogrio.read_info(path, **kw)
        src_crs = info.get("crs")
        n = info.get("features", -1)
    except Exception:
        n = -1
    if bbox_wgs84 is not None and n > 20000:
        try:
            bbs = gpd.GeoSeries([box(*bbox_wgs84)], crs=4326)
            bb = bbs.to_crs(src_crs).total_bounds if src_crs else bbs.total_bounds
            kw["bbox"] = tuple(bb)
        except Exception:
            pass
    try:
        g = gpd.read_file(path, **kw)
    except UnicodeDecodeError:
        g = gpd.read_file(path, encoding="cp1251", **kw)
    if g.empty:
        return g
    g = fix_text_encoding(g, path, kw)
    g = g[~g.geometry.isna() & ~g.geometry.is_empty].copy()
    crs = fix_crs(g.crs, g.total_bounds, Path(path).name)
    g = g.set_crs(crs, allow_override=True)
    if g.crs.to_epsg() != WORK_EPSG:
        g = g.to_crs(WORK_EPSG)
    g["__src"] = Path(path).stem if not layer else f"{Path(path).stem}:{layer}"
    return g


_CYR_RE = re.compile(r"[\u0400-\u04FF]")
_MOJI_RE = re.compile(r"[\u00C0-\u00FF][\u0080-\u00BF]")   # UTF-8, прочитанный как latin-1


def _remojibake(s):
    """«Р›РµСЃ» / «Ð›ÐµÑ» → «Лес»: пробуем обратить неверное декодирование UTF-8 как cp1251/latin1."""
    if not isinstance(s, str) or _CYR_RE.search(s) and not re.search(r"[РС][\u0080-\u04FF]", s):
        return s
    for enc in ("cp1251", "latin1"):
        try:
            r = s.encode(enc).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        if _CYR_RE.search(r) or r.isascii():
            return r
    return s


def fix_text_encoding(g: gpd.GeoDataFrame, path: str, kw: dict) -> gpd.GeoDataFrame:
    """Shapefile без .cpg часто читается с «кракозябрами». Пробуем latin1→utf8, иначе перечитываем как cp1251."""
    # имена полей (кириллица в shp без .cpg)
    ren = {c: _remojibake(c) for c in g.columns if isinstance(c, str) and c != g.geometry.name}
    ren = {k: v for k, v in ren.items() if v != k and v not in g.columns}
    if ren:
        g = g.rename(columns=ren)
        log.info("%s: имена полей перекодированы: %s", Path(path).name, ", ".join(ren.values()))
    cols = [c for c in g.columns if c != g.geometry.name and (g[c].dtype == object or pd.api.types.is_string_dtype(g[c]))]
    if not cols:
        return g
    sample = " ".join(str(v) for c in cols for v in g[c].dropna().astype(str).head(30))
    if not sample.strip():
        return g
    moji_cp = re.search(r"[РС][\u0402-\u045F\u2018-\u2122]", sample)   # UTF-8, прочитанный как cp1251 («Р›РµСЃ»)
    if _CYR_RE.search(sample) and not moji_cp:
        return g                                  # уже нормальная кириллица
    if _MOJI_RE.search(sample) or moji_cp:
        for c in cols:
            g[c] = g[c].map(_remojibake)
        log.info("%s: атрибуты перекодированы latin1→utf-8", Path(path).name)
        return g
    if re.search(r"[\u0080-\u00FF]", sample):      # похоже на cp1251, прочитанный как latin1
        try:
            g2 = gpd.read_file(path, encoding="cp1251", **kw)
            log.info("%s: атрибуты перечитаны в cp1251", Path(path).name)
            return g2
        except Exception:
            pass
    return g


def read_many(paths: Iterable[str], bbox_wgs84=None, clip=None) -> gpd.GeoDataFrame | None:
    parts = []
    for p in paths:
        layer = None
        if "|" in p:                      # «file.gpkg|layer»
            p, layer = p.split("|", 1)
        try:
            if p.lower().endswith(".gpkg") and layer is None:
                for ly in [l[0] for l in pyogrio.list_layers(p)]:
                    parts.append(read_vector(p, ly, bbox_wgs84))
            else:
                parts.append(read_vector(p, layer, bbox_wgs84))
            log.info("Загружен слой %s (%d объектов)", Path(p).name, len(parts[-1]))
        except Exception as e:
            log.error("Не удалось прочитать %s: %s", p, e)
    parts = [x for x in parts if x is not None and not x.empty]
    if not parts:
        return None
    g = gpd.GeoDataFrame(pd.concat(parts, ignore_index=True), crs=WORK_EPSG)
    if clip is not None:
        g = g[g.intersects(clip)].copy()
    return g


LATIN_RE = re.compile(r"^[A-Z][a-z]+(\s+(sp\.?|[a-z\-]+))?(\s+[a-z\-]+)?$")


def detect_species_column(g: gpd.GeoDataFrame) -> str | None:
    """Ищет колонку с латинскими названиями видов: сначала по имени, потом по содержимому."""
    def latin_share(c):
        s = g[c].dropna().astype(str).str.strip()
        s = s[s != ""].head(300)
        return s.str.match(LATIN_RE).mean() if not s.empty else 0.0

    prefer = ["species", "Species", "SPECIES", "taxon", "Taxon", "latin", "Latin", "lat_name", "Latin_name",
              "Вид_лат", "Вид", "вид"]
    for c in prefer:
        if c in g.columns and latin_share(c) >= 0.5:     # «Вид» может быть и русским названием — проверяем содержимое
            return c
    best, best_score = None, 0.0
    for c in g.columns:
        if c == g.geometry.name or not (g[c].dtype == object or pd.api.types.is_string_dtype(g[c])):
            continue
        score = latin_share(c)
        if score > best_score:
            best, best_score = c, score
    return best if best_score >= 0.5 else None


def species_from_filename(path: str) -> str:
    stem = Path(path).stem
    stem = re.sub(r"[_\-]+", " ", stem).strip()
    return stem[0].upper() + stem[1:] if stem else Path(path).stem


def read_raster_clipped(path: str, clip_bounds_work, resolution: float | None = None, what="raster"):
    """Читает растр, исправляет CRS, перепроецирует в WORK_EPSG, обрезает по bounds (в WORK_EPSG).
    Возвращает (array, transform, nodata)."""
    with rasterio.open(path) as src:
        crs = fix_crs(src.crs, src.bounds, Path(path).name)
        src_nodata = src.nodata
        if src_nodata is None and path.lower().endswith(".asc"):
            src_nodata = -9999.0
        minx, miny, maxx, maxy = clip_bounds_work
        if resolution is None:
            # оценка разрешения источника в метрах
            if crs.to_epsg() == 4326:
                resolution = float(abs(src.res[0])) * 111320 * np.cos(np.deg2rad((miny + maxy) / 2 / 111320))
                resolution = max(1.0, round(resolution / 2.5, 0) * 2.5) if resolution > 0 else CELL
            else:
                resolution = float(abs(src.res[0]))
        width = int(np.ceil((maxx - minx) / resolution))
        height = int(np.ceil((maxy - miny) / resolution))
        if width <= 0 or height <= 0:
            return None, None, None
        dst_transform = from_origin(minx, maxy, resolution, resolution)
        dst = np.full((height, width), np.nan, dtype="float32")
        # reproject поддерживает пересчёт из исправленной CRS (src_crs передаём явно)
        reproject(
            source=rasterio.band(src, 1), destination=dst,
            src_transform=src.transform, src_crs=crs, src_nodata=src_nodata,
            dst_transform=dst_transform, dst_crs=rasterio.crs.CRS.from_epsg(WORK_EPSG),
            dst_nodata=np.nan, resampling=Resampling.bilinear,
        )
        if src_nodata is not None:
            dst[dst == src_nodata] = np.nan
        return dst, dst_transform, np.nan


def polygonize_mask(mask: np.ndarray, transform, min_area_m2: float = 0.0) -> gpd.GeoDataFrame:
    """Булева маска → полигоны (WORK_EPSG)."""
    m = mask.astype("uint8")
    geoms = [shape(geom) for geom, val in rfeatures.shapes(m, mask=m.astype(bool), transform=transform) if val == 1]
    g = gpd.GeoDataFrame(geometry=geoms, crs=WORK_EPSG)
    if min_area_m2 > 0 and not g.empty:
        g = g[g.area >= min_area_m2].copy()
    return g


# =====================================================================================
#                              РАСТРЫ ПЛОТНОСТИ И УКЛОНА
# =====================================================================================
def make_grid(bounds, cell=CELL):
    minx, miny, maxx, maxy = bounds
    width = int(np.ceil((maxx - minx) / cell))
    height = int(np.ceil((maxy - miny) / cell))
    return from_origin(minx, maxy, cell, cell), (height, width)


def circular_kernel(radius_m: float, cell: float) -> np.ndarray:
    r = int(np.ceil(radius_m / cell))
    y, x = np.ogrid[-r:r + 1, -r:r + 1]
    return ((x * x + y * y) <= r * r).astype("float32")


def line_density(lines: gpd.GeoDataFrame, bounds, radius_m: float, cell=CELL):
    """Плотность линий (м/га) в круговом скользящем окне radius_m. Линии дискретизируются с шагом cell/2."""
    transform, shp = make_grid(bounds, cell)
    acc = np.zeros(shp, dtype="float32")
    step = cell / 2
    minx, miny, maxx, maxy = bounds
    for geom in lines.geometry:
        for part in getattr(geom, "geoms", [geom]):
            L = part.length
            if L == 0:
                continue
            n = max(2, int(L / step) + 1)
            d = np.linspace(0, L, n)
            pts = np.array([part.interpolate(t).coords[0] for t in d])
            seg = L / (n - 1)
            cols = ((pts[:, 0] - minx) / cell).astype(int)
            rows = ((maxy - pts[:, 1]) / cell).astype(int)
            ok = (cols >= 0) & (cols < shp[1]) & (rows >= 0) & (rows < shp[0])
            np.add.at(acc, (rows[ok], cols[ok]), seg)
    k = circular_kernel(radius_m, cell)
    total = ndimage.convolve(acc, k, mode="constant", cval=0.0)
    area_ha = k.sum() * cell * cell / 1e4
    return total / area_ha, transform            # м/га


def point_density(points: gpd.GeoDataFrame, bounds, radius_m: float, cell=CELL, weights=None):
    """Плотность точек (шт/га) в круговом окне radius_m."""
    transform, shp = make_grid(bounds, cell)
    acc = np.zeros(shp, dtype="float32")
    minx, miny, maxx, maxy = bounds
    xs = points.geometry.x.values
    ys = points.geometry.y.values
    w = np.ones(len(points), dtype="float32") if weights is None else np.asarray(weights, dtype="float32")
    cols = ((xs - minx) / cell).astype(int)
    rows = ((maxy - ys) / cell).astype(int)
    ok = (cols >= 0) & (cols < shp[1]) & (rows >= 0) & (rows < shp[0])
    np.add.at(acc, (rows[ok], cols[ok]), w[ok])
    k = circular_kernel(radius_m, cell)
    total = ndimage.convolve(acc, k, mode="constant", cval=0.0)
    area_ha = k.sum() * cell * cell / 1e4
    return total / area_ha, transform


def slope_percent(dem: np.ndarray, cell: float) -> np.ndarray:
    dem_f = dem.astype("float64")
    dem_f = np.where(np.isnan(dem_f), np.nanmedian(dem_f), dem_f)
    dem_f = ndimage.uniform_filter(dem_f, size=3, mode="nearest")   # лёгкое сглаживание шума
    dzdy, dzdx = np.gradient(dem_f, cell)
    slope = np.sqrt(dzdx ** 2 + dzdy ** 2) * 100.0
    slope[np.isnan(dem)] = np.nan
    return slope.astype("float32")


def upper_quantile_mask(density: np.ndarray, inside: np.ndarray, q=QUANTILE):
    """Маска ячеек ≥ q-квантиля, рассчитанного по ячейкам внутри ООЗТ с ненулевой плотностью."""
    vals = density[inside & (density > 0)]
    if vals.size == 0:
        return np.zeros_like(density, dtype=bool), np.nan
    thr = float(np.quantile(vals, q))
    return inside & (density >= thr) & (density > 0), thr


# =====================================================================================
#                                 КОНТЕКСТ ДАННЫХ
# =====================================================================================
class Data:
    """Все загруженные слои в WORK_EPSG + служебные объекты."""

    def __init__(self, cfg: Config, oozt_only: bool = False):
        self.cfg = cfg
        self.oozt = self._load_oozt()
        self.union = unary_union(self.oozt.geometry.values)
        self.bounds = tuple(self.oozt.total_bounds)
        pad = max(TRAIL_WINDOW_M, CORVUS_WINDOW_M, RARE_BUFFER_M, SHORE_BUFFER_M) + 2 * CELL
        b = self.bounds
        self.bounds_pad = (b[0] - pad, b[1] - pad, b[2] + pad, b[3] + pad)
        self.bbox_wgs84 = tuple(gpd.GeoSeries([box(*self.bounds_pad)], crs=WORK_EPSG).to_crs(4326).total_bounds)
        self.clip = box(*self.bounds_pad)
        self.trails = self.hydro_lines = self.hydro_polys = None
        self.inv_pts = self.inv_polys = self.rare_pts = self.rare_polys = self.fauna_pts = None
        self.dem = None
        self._inside = None
        if oozt_only:                      # режим --report-only: остальные слои не нужны
            return

        self.trails = self._load_trails()
        self.hydro_lines = read_many(cfg.hydro_lines, self.bbox_wgs84, self.clip)
        self.hydro_polys = read_many(cfg.hydro_polys, self.bbox_wgs84, self.clip)
        self.inv_pts = self._load_species_points(cfg.invasive_points)
        self.inv_polys = self._load_species_polys(cfg.invasive_polys)
        self.rare_pts = self._load_species_points(cfg.rare_points)
        self.rare_polys = self._load_species_polys(cfg.rare_polys)
        self.fauna_pts = self._load_species_points(cfg.fauna_points)
        self.dem = None
        if cfg.dem_path:
            try:
                arr, tr, _ = read_raster_clipped(cfg.dem_path, self.bounds_pad, resolution=CELL, what="ЦМР")
                self.dem = (arr, tr)
                log.info("ЦМР загружена: %s, ячейка %.0f м", Path(cfg.dem_path).name, CELL)
            except Exception as e:
                log.error("ЦМР не загружена: %s", e)
        self._inside = None

    # ---- ООЗТ ----
    def _load_oozt(self) -> gpd.GeoDataFrame:
        g = read_vector(self.cfg.oozt_path, self.cfg.oozt_layer)
        f = self.cfg.fields or {}
        cols = list(g.columns)

        def pick(key, candidates):
            if f.get(key) and f[key] in cols:
                return f[key]
            for c in candidates:
                if c in cols:
                    return c
            for c in cols:
                if c.lower() in [x.lower() for x in candidates]:
                    return c
            raise KeyError(f"В слое ООЗТ не найдено поле {key} (искал {candidates}); задайте cfg.fields['{key}']")

        self.F_OOZT = pick("OOZT", ["OOZT", "ООЗТ", "Name_OOZT"])
        self.F_KV = pick("KV", ["KV", "FID_kv", "kv", "Квартал", "kvartal"])
        self.F_VYD = pick("VYD", ["VYD", "FID_ecosys", "vyd", "Выдел", "vydel"])
        self.F_TYPE = pick("TYPE", ["Type", "TYPE", "Тип"])
        self.F_NAME = pick("NAME", ["Name", "NAME", "Название", "Ecosystem"])
        self.cfg.fields = {"OOZT": self.F_OOZT, "KV": self.F_KV, "VYD": self.F_VYD,
                           "TYPE": self.F_TYPE, "NAME": self.F_NAME}
        g = g.rename(columns={self.F_OOZT: "OOZT", self.F_KV: "KV", self.F_VYD: "VYD",
                              self.F_TYPE: "Type", self.F_NAME: "Name"})
        for f in ("KV", "VYD"):
            try:
                g[f] = pd.to_numeric(g[f]).astype("Int64")
            except (ValueError, TypeError):
                pass                                  # оставляем как есть (текстовые номера)
        g["Type"] = g["Type"].astype(str).str.strip()
        g["Name"] = g["Name"].astype(str).str.strip()
        g["uid"] = g["OOZT"].astype(str) + "|" + g["KV"].astype(str) + "|" + g["VYD"].astype(str)
        g["area_ha"] = g.geometry.area / 1e4
        g = g.reset_index(drop=True)
        log.info("ООЗТ: %d выделов, %d кварталов, ООЗТ: %s", len(g), g["KV"].nunique(),
                 ", ".join(map(str, g["OOZT"].unique())))
        return g

    # ---- тропы ----
    def _load_trails(self):
        g = read_many(self.cfg.trails, self.bbox_wgs84, self.clip)
        if g is None:
            return None
        col = next((c for c in ("HIGHWAY", "highway", "fclass", "FCLASS", "type") if c in g.columns), None)
        if col is not None:
            types = g[col].astype(str).str.lower()
            keep = types.isin(TRAIL_HIGHWAY_TYPES) | types.isna() | (types == "none") | (types == "nan")
            if keep.sum() > 0:
                log.info("Тропы: отфильтровано по %s — оставлено %d из %d (типы: %s)", col, keep.sum(), len(g),
                         ", ".join(sorted(types[keep].unique())[:12]))
                g = g[keep].copy()
        g = g[g.geom_type.isin(["LineString", "MultiLineString"])].copy()
        return g if not g.empty else None

    # ---- виды ----
    def _load_species_points(self, paths):
        g = read_many(paths, self.bbox_wgs84, self.clip)
        if g is None:
            return None
        g = g[g.geom_type.isin(["Point", "MultiPoint"])].explode(index_parts=False)
        if g.empty:
            log.warning("Точки видов %s: нет объектов в границах ООЗТ — слой не используется",
                        ", ".join(map(str, paths if isinstance(paths, (list, tuple)) else [paths])))
            return None
        col = detect_species_column(g)
        g["species"] = g[col].astype(str).str.strip() if col else g["__src"]
        if col:
            log.info("Точки видов %s: поле вида «%s», %d видов, %d точек", ", ".join(map(str, g["__src"].unique())), col,
                     g["species"].nunique(), len(g))
        else:
            log.warning("Точки видов %s: поле с латинскими названиями не найдено — вид = имя файла", ", ".join(map(str, g["__src"].unique())))
        return g[["species", "__src", "geometry"]]

    def _load_species_polys(self, paths):
        g = read_many(paths, self.bbox_wgs84, self.clip)
        if g is None:
            return None
        g = g[g.geom_type.isin(["Polygon", "MultiPolygon"])].copy()
        if g.empty:
            log.warning("Полигоны видов: нет объектов в границах ООЗТ — слой не используется")
            return None
        col = detect_species_column(g)
        g["species"] = g[col].astype(str).str.strip() if col else g["__src"]
        return g[["species", "__src", "geometry"]]

    # ---- маска «внутри ООЗТ» на сетке CELL ----
    def inside_mask(self, transform, shp):
        m = rfeatures.rasterize([(self.union, 1)], out_shape=shp, transform=transform, fill=0, dtype="uint8")
        return m.astype(bool)

    def vydels_of_type(self, *keys) -> gpd.GeoDataFrame:
        t = self.oozt["Type"].str.lower()
        m = np.zeros(len(t), dtype=bool)
        for k in keys:
            m |= t.str.contains(k, regex=False).values
        return self.oozt[m]


# =====================================================================================
#                                    РЕЗУЛЬТАТЫ
# =====================================================================================
class Results:
    def __init__(self):
        self.rows: list[dict] = []          # длинная таблица
        self.evidence: list[gpd.GeoDataFrame] = []

    def add(self, vydel_row, measure: dict, basis: str, value=None):
        self.rows.append({
            "ООЗТ": vydel_row["OOZT"], "Квартал": vydel_row["KV"], "Выдел": vydel_row["VYD"],
            "Тип выдела": vydel_row["Type"], "Экосистема": vydel_row["Name"],
            "Тип мероприятий": measure.get("group", ""), "Мероприятие": measure["name"],
            "Тип экосистемы мероприятия": measure.get("eco", ""),
            "Пункт Приказа": measure.get("prikaz", ""), "Сезонность": measure.get("season", ""),
            "Основание": basis, "Значение": value,
        })

    def add_evidence(self, g: gpd.GeoDataFrame, measure_name: str, layer: str, label_col: str | None = None):
        if g is None or g.empty:
            return
        n = len(g)
        labels = g[label_col].astype(str).values if label_col and label_col in g.columns else [""] * n
        e = gpd.GeoDataFrame({"measure": [measure_name] * n, "layer": [layer] * n, "label": list(labels)},
                             geometry=g.geometry.values, crs=WORK_EPSG)
        self.evidence.append(e)

    def long_table(self) -> pd.DataFrame:
        df = pd.DataFrame(self.rows)
        if df.empty:
            return df
        # объединяем дубли (один выдел — одно мероприятие) с конкатенацией оснований
        key = ["ООЗТ", "Квартал", "Выдел", "Мероприятие"]
        agg = df.groupby(key, as_index=False, sort=False).agg({
            "Тип выдела": "first", "Экосистема": "first", "Тип мероприятий": "first",
            "Тип экосистемы мероприятия": "first",
            "Пункт Приказа": "first", "Сезонность": "first",
            "Основание": lambda s: "; ".join(dict.fromkeys([x for x in s if x])),
            "Значение": "first",
        })
        agg = agg.sort_values(["ООЗТ", "Квартал", "Выдел", "Тип мероприятий", "Мероприятие"])
        cols = ["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема", "Тип мероприятий", "Мероприятие",
                "Тип экосистемы мероприятия", "Пункт Приказа", "Сезонность", "Основание", "Значение"]
        return agg[cols].reset_index(drop=True)

    def evidence_gdf(self) -> gpd.GeoDataFrame | None:
        if not self.evidence:
            return None
        return gpd.GeoDataFrame(pd.concat(self.evidence, ignore_index=True), crs=WORK_EPSG)


# =====================================================================================
#                                     ПРАВИЛА
# =====================================================================================
def _overlay_patches_to_vydels(data: Data, patches: gpd.GeoDataFrame, target: gpd.GeoDataFrame,
                               min_ha=MIN_PATCH_HA):
    """Пересечение полигонов-пятен с выделами; возвращает DataFrame uid → площадь пятен (га), доля."""
    if patches is None or patches.empty or target.empty:
        return pd.DataFrame(columns=["uid", "patch_ha", "share"])
    pu = unary_union(patches.geometry.values)
    inter = target.geometry.intersection(pu)
    ha = inter.area / 1e4
    df = pd.DataFrame({"uid": target["uid"].values, "patch_ha": ha.values,
                       "share": (ha / target["area_ha"]).values})
    return df[df["patch_ha"] >= min_ha]


def rule_type(data: Data, res: Results, measure: dict, keys, label):
    """Назначение по типу выдела."""
    sel = data.vydels_of_type(*keys)
    for _, r in sel.iterrows():
        res.add(r, measure, f"тип выдела «{r['Type']}» ({label})")
    log.info("  → %d выделов", len(sel))


def rule_trail_density_vydel(data: Data, res: Results, measure: dict, type_keys=TYPE_FOREST):
    """Плотность троп (м/га) по выделу; верхний квартиль среди выделов всех ООЗТ заданного типа."""
    if data.trails is None:
        log.warning("  тропы не загружены — пропуск")
        return
    target = data.vydels_of_type(*type_keys) if type_keys else data.oozt
    tu = unary_union(data.trails.geometry.values)
    dens = target.geometry.intersection(tu).length / target["area_ha"]
    dens = dens.replace([np.inf], np.nan).fillna(0)
    pos = dens[dens > 0]
    if pos.empty:
        return
    thr = float(np.quantile(pos, QUANTILE))
    sel = target[dens >= thr]
    for (_, r), d in zip(sel.iterrows(), dens[dens >= thr]):
        res.add(r, measure, f"плотность троп {d:.0f} м/га ≥ Q75 = {thr:.0f} м/га", round(d, 1))
    trails_in = None
    if not sel.empty:
        # Красным показываем не целые исходные линии, а только их фрагменты
        # внутри локальных пятен высокой плотности (Q75) и выбранных выделов.
        hot_polys, _, _ = _trail_hotspots(data)
        if hot_polys is not None and not hot_polys.empty:
            clip_geom = unary_union(hot_polys.geometry.values).intersection(
                unary_union(sel.geometry.values)
            )
            trails_in = data.trails[data.trails.intersects(clip_geom)].copy()
            if not trails_in.empty:
                trails_in["geometry"] = trails_in.geometry.intersection(clip_geom)
                trails_in = trails_in.explode(index_parts=False)
                trails_in = trails_in[
                    (~trails_in.is_empty)
                    & trails_in.geom_type.isin(["LineString", "MultiLineString"])
                ]
    res.add_evidence(trails_in, measure["name"],
                     "тропы в зонах повышенной плотности")
    log.info("  → %d выделов (порог %.0f м/га)", len(sel), thr)


def _trail_hotspots(data: Data):
    """Кэш: маска верхнего квартиля плотности троп в окне TRAIL_WINDOW_M."""
    if getattr(data, "_hot", None) is not None:
        return data._hot
    if data.trails is None:
        data._hot = (None, None, None)
        return data._hot
    dens, tr = line_density(data.trails, data.bounds_pad, TRAIL_WINDOW_M, CELL)
    inside = data.inside_mask(tr, dens.shape)
    mask, thr = upper_quantile_mask(dens, inside)
    polys = polygonize_mask(mask, tr)
    log.info("  плотность троп: порог Q75 = %.0f м/га, %d пятен", thr, len(polys))
    data._hot = (polys, thr, (dens, tr))
    return data._hot


def rule_trail_hotspot(data: Data, res: Results, measure: dict, type_keys=TYPE_FOREST):
    polys, thr, _ = _trail_hotspots(data)
    if polys is None:
        log.warning("  тропы не загружены — пропуск")
        return
    target = data.vydels_of_type(*type_keys) if type_keys else data.oozt
    hit = _overlay_patches_to_vydels(data, polys, target)
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        res.add(r, measure, f"участки повышенной плотности троп (окно {TRAIL_WINDOW_M:.0f} м, ≥Q75={thr:.0f} м/га): "
                            f"{h['patch_ha']:.2f} га", round(h["patch_ha"], 2))
    if not hit.empty:
        sel = target[target["uid"].isin(hit["uid"])]
        pieces = polys[polys.intersects(unary_union(sel.geometry.values))]
        res.add_evidence(pieces, measure["name"],
                         "зоны повышенной плотности троп (верхний квартиль)")
    log.info("  → %d выделов", len(hit))


def rule_shore_trails(data: Data, res: Results, measure: dict):
    """Буфер вдоль водных объектов ∩ пятна повышенной плотности троп."""
    polys, thr, _ = _trail_hotspots(data)
    if polys is None or (data.hydro_lines is None and data.hydro_polys is None):
        log.warning("  нет троп или гидрологии — пропуск")
        return
    parts = [g.geometry for g in (data.hydro_lines, data.hydro_polys) if g is not None]
    water = unary_union(pd.concat(parts).values)
    shore = water.buffer(SHORE_BUFFER_M).difference(water) if water.area > 0 else water.buffer(SHORE_BUFFER_M)
    hot_shore = polys[polys.intersects(shore)].copy()
    hot_shore["geometry"] = hot_shore.geometry.intersection(shore)
    hot_shore = hot_shore[~hot_shore.is_empty]
    hit = _overlay_patches_to_vydels(data, hot_shore, data.oozt, min_ha=MIN_PATCH_HA / 5)
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        res.add(r, measure, f"тропы повышенной плотности в {SHORE_BUFFER_M:.0f} м от водных объектов: "
                            f"{h['patch_ha']:.2f} га", round(h["patch_ha"], 2))
    res.add_evidence(hot_shore, measure["name"],
                     f"тропы на расстоянии {SHORE_BUFFER_M:.0f} м от водных объектов (Q75)")
    # сами водные объекты рядом с найденными участками — как контекст на схеме
    if not hot_shore.empty:
        near = hot_shore.geometry.buffer(SHORE_BUFFER_M).union_all()
        if data.hydro_lines is not None:
            res.add_evidence(data.hydro_lines[data.hydro_lines.intersects(near)], measure["name"], "водные объекты (линии)")
        if data.hydro_polys is not None:
            res.add_evidence(data.hydro_polys[data.hydro_polys.intersects(near)], measure["name"], "водные объекты (полигоны)")
    log.info("  → %d выделов", len(hit))


def _species_group(data: Data, kind: str):
    if kind == "invasive":
        return data.inv_pts, data.inv_polys, data.cfg.invasive_rasters
    return data.rare_pts, data.rare_polys, data.cfg.rare_rasters


def _sdm_patches(data: Data, raster_paths, thr=SDM_THRESHOLD):
    """Растры SDM → полигоны ≥ thr (по каждому виду)."""
    out = []
    for p in raster_paths:
        try:
            arr, tr, _ = read_raster_clipped(p, data.bounds_pad)
            if arr is None:
                continue
            inside = data.inside_mask(tr, arr.shape)
            mask = np.nan_to_num(arr, nan=-1) >= thr
            polys = polygonize_mask(mask & inside, tr)
            polys["species"] = species_from_filename(p)
            out.append(polys)
            log.info("  SDM %s: %d пятен ≥ %.2f", Path(p).name, len(polys), thr)
        except Exception as e:
            log.error("  SDM %s не обработан: %s", p, e)
    if not out:
        return None
    return gpd.GeoDataFrame(pd.concat(out, ignore_index=True), crs=WORK_EPSG)


def rule_invasive(data: Data, res: Results, measure: dict):
    pts, polys, rasters = _species_group(data, "invasive")
    # точки
    if pts is not None and not pts.empty:
        j = gpd.sjoin(pts, data.oozt[["uid", "geometry"]], how="inner", predicate="within")
        for uid, grp in j.groupby("uid"):
            r = data.oozt[data.oozt["uid"] == uid].iloc[0]
            sp = grp.groupby("species").size().sort_values(ascending=False)
            coords = []
            for s, g in grp.groupby("species"):
                gg = gpd.GeoSeries(g.geometry, crs=WORK_EPSG).to_crs(4326)
                coords.append(f"{s} ({gg.x.mean():.6f}° в.д., {gg.y.mean():.6f}° с.ш.)")
            res.add(r, measure, f"точечное удаление: {', '.join(f'{k} ({v})' for k, v in sp.items())}; "
                                f"координаты WGS 84: {'; '.join(coords)}", int(len(grp)))
        res.add_evidence(j, measure["name"], "инвазионные виды — точки", "species")
        log.info("  точки: %d выделов", j["uid"].nunique())
    # полигоны
    if polys is not None and not polys.empty:
        j = gpd.overlay(polys, data.oozt[["uid", "geometry"]], how="intersection")
        j["ha"] = j.area / 1e4
        for uid, grp in j.groupby("uid"):
            r = data.oozt[data.oozt["uid"] == uid].iloc[0]
            txt = ", ".join(f"{s} ({a:.2f} га)" for s, a in grp.groupby("species")["ha"].sum().items())
            res.add(r, measure, f"удаление в пределах полигонов произрастания: {txt}", round(grp["ha"].sum(), 2))
        res.add_evidence(j, measure["name"], "инвазионные виды — полигоны", "species")
    # растры
    sdm = _sdm_patches(data, rasters, thr=SDM_THRESHOLD_INVASIVE)
    if sdm is not None and not sdm.empty:
        m2 = dict(measure)
        m2["name"] = measure["name"] + " — мониторинг (SDM ≥ %.2f)" % SDM_THRESHOLD_INVASIVE
        j = gpd.overlay(sdm, data.oozt[["uid", "area_ha", "geometry"]], how="intersection")
        j["ha"] = j.area / 1e4
        for uid, grp in j.groupby("uid"):
            tot = grp.groupby("species")["ha"].sum()
            tot = tot[tot >= MIN_PATCH_HA]
            if tot.empty:
                continue
            r = data.oozt[data.oozt["uid"] == uid].iloc[0]
            txt = ", ".join(f"{s} ({a:.2f} га)" for s, a in tot.items())
            res.add(r, m2, f"вероятность SDM ≥ {SDM_THRESHOLD_INVASIVE}: {txt}", round(float(tot.sum()), 2))
        res.add_evidence(j, m2["name"],
                         f"инвазионные виды — SDM (≥ {SDM_THRESHOLD_INVASIVE:.2f})", "species")


def rule_rare_trails(data: Data, res: Results, measure: dict):
    """Ликвидация троп (ограничения) в 50 м от точек редких видов / в полигонах / в SDM ≥ порога для редких видов (мониторинг)."""
    pts, polys, rasters = _species_group(data, "rare")
    trails = data.trails
    zones = []
    if pts is not None and not pts.empty:
        b = pts.copy()
        b["geometry"] = b.geometry.buffer(RARE_BUFFER_M)
        b["src"] = "точки (буфер %d м)" % RARE_BUFFER_M
        zones.append(b)
    if polys is not None and not polys.empty:
        p = polys.copy()
        p["src"] = "полигоны"
        zones.append(p)
    if zones:
        z = gpd.GeoDataFrame(pd.concat(zones, ignore_index=True), crs=WORK_EPSG)
        if trails is not None:
            zu = unary_union(z.geometry.values)
            t = trails[trails.intersects(zu)].copy()
            t["geometry"] = t.geometry.intersection(zu)
            t = t[~t.is_empty]
            if not t.empty:
                j = gpd.overlay(t.explode(index_parts=False), data.oozt[["uid", "geometry"]], how="intersection",
                                keep_geom_type=False)
                j = j[j.geom_type.isin(["LineString", "MultiLineString"])]
                j["len"] = j.length
                for uid, grp in j.groupby("uid"):
                    r = data.oozt[data.oozt["uid"] == uid].iloc[0]
                    # какие виды
                    zz = z[z.intersects(r.geometry)]
                    sp = ", ".join(sorted(zz["species"].unique())[:8])
                    res.add(r, measure, f"тропы в зоне редких видов ({sp}): {grp['len'].sum():.0f} м",
                            round(grp["len"].sum()))
                res.add_evidence(j, measure["name"],
                                 f"тропы на расстоянии {RARE_BUFFER_M:.0f} м от редких видов")
                log.info("  тропы на расстоянии %.0f м от редких видов: %d выделов",
                         RARE_BUFFER_M, j["uid"].nunique())
            res.add_evidence(pts[pts.intersects(data.union)] if pts is not None else None,
                             measure["name"], "редкие виды — точки", "species")
            res.add_evidence(polys, measure["name"], "редкие виды — полигоны", "species")
        else:
            # без троп — назначаем по самим зонам
            j = gpd.overlay(z, data.oozt[["uid", "geometry"]], how="intersection")
            for uid, grp in j.groupby("uid"):
                r = data.oozt[data.oozt["uid"] == uid].iloc[0]
                res.add(r, measure, "местообитания редких видов: " + ", ".join(sorted(grp["species"].unique())[:8]))
            res.add_evidence(j, measure["name"], "редкие виды", "species")
    sdm = _sdm_patches(data, rasters, thr=SDM_THRESHOLD_RARE)
    if sdm is not None and not sdm.empty:
        m2 = dict(measure)
        m2["name"] = measure["name"] + " — мониторинг (SDM ≥ %.2f)" % SDM_THRESHOLD_RARE
        j = gpd.overlay(sdm, data.oozt[["uid", "geometry"]], how="intersection")
        j["ha"] = j.area / 1e4
        for uid, grp in j.groupby("uid"):
            tot = grp.groupby("species")["ha"].sum()
            tot = tot[tot >= MIN_PATCH_HA]
            if tot.empty:
                continue
            r = data.oozt[data.oozt["uid"] == uid].iloc[0]
            res.add(r, m2, "вероятность SDM ≥ %.2f: " % SDM_THRESHOLD_RARE +
                    ", ".join(f"{s} ({a:.2f} га)" for s, a in tot.items()), round(float(tot.sum()), 2))
        res.add_evidence(j, m2["name"], "редкие виды — SDM", "species")


def rule_slope(data: Data, res: Results, measure: dict, threshold_pct: float, type_keys=None):
    if data.dem is None:
        log.warning("  ЦМР не загружена — пропуск")
        return
    arr, tr = data.dem
    if getattr(data, "_slope", None) is None:
        data._slope = slope_percent(arr, CELL)
    slope = data._slope
    inside = data.inside_mask(tr, slope.shape)
    mask = inside & (np.nan_to_num(slope, nan=0) > threshold_pct)
    # убираем одиночные ячейки
    mask = ndimage.binary_opening(mask, structure=np.ones((2, 2)))
    polys = polygonize_mask(mask, tr)
    target = data.vydels_of_type(*type_keys) if type_keys else data.oozt
    hit = _overlay_patches_to_vydels(data, polys, target)
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        eco_label = "в луговых экосистемах" if type_keys else "в любых экосистемах"
        res.add(r, measure, f"уклон более {threshold_pct:g}% {eco_label}: "
                           f"{h['patch_ha']:.2f} га ({h['share'] * 100:.0f} % выдела)",
                round(h["patch_ha"], 2))
    if not hit.empty:
        sel = target[target["uid"].isin(hit["uid"])]
        pieces = polys[polys.intersects(unary_union(sel.geometry.values))]
        eco_label = "в луговых экосистемах" if type_keys else "в любых экосистемах"
        res.add_evidence(pieces, measure["name"], f"уклон более {threshold_pct:g}% {eco_label}")
    log.info("  → %d выделов (%d пятен)", len(hit), len(polys))


def rule_corvus(data: Data, res: Results, measure: dict, genus="Corvus", type_keys=TYPE_FOREST):
    pts = data.fauna_pts
    if pts is None or pts.empty:
        log.warning("  точки фауны не загружены — пропуск")
        return
    c = pts[pts["species"].str.contains(rf"\b{genus}\b", case=False, regex=True)]
    if c.empty:
        log.warning("  точки рода %s не найдены", genus)
        return
    dens, tr = point_density(c, data.bounds_pad, CORVUS_WINDOW_M, CELL)
    inside = data.inside_mask(tr, dens.shape)
    mask, thr = upper_quantile_mask(dens, inside)
    polys = polygonize_mask(mask, tr)
    target = data.vydels_of_type(*type_keys) if type_keys else data.oozt
    hit = _overlay_patches_to_vydels(data, polys, target)
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        res.add(r, measure, f"плотность {genus} (окно {CORVUS_WINDOW_M:.0f} м) ≥ Q75 = {thr:.3f} шт/га: "
                            f"{h['patch_ha']:.2f} га", round(h["patch_ha"], 2))
    res.add_evidence(polys[polys.intersects(data.union)], measure["name"], f"плотность {genus} (Q75)")
    res.add_evidence(c[c.intersects(data.union.buffer(CORVUS_WINDOW_M))], measure["name"], f"{genus} — точки", "species")
    log.info("  → %d выделов (%d точек %s)", len(hit), len(c), genus)


# --- сопоставление текста «Метод выявления» с правилом -------------------------------
def dispatch_rule(measure: dict, data: Data, res: Results) -> bool:
    m = (measure.get("method") or "").lower()
    nm = (measure.get("name") or "").lower()
    if not m or "нет данных" in m:
        return False
    log.info("Мероприятие: %s", measure["name"])

    if "corvus" in m or "ворон" in nm:
        rule_corvus(data, res, measure)
    elif "уклон" in m or "цмр" in m:
        mt = re.search(r"(\d+(?:[.,]\d+)?)\s*%", m)
        thr = float(mt.group(1).replace(",", ".")) if mt else 7.0
        keys = TYPE_MEADOW if "лугов" in m else None
        if keys is None and "люб" in str(measure.get("eco", "")).lower():
            thr = 7.0
        rule_slope(data, res, measure, thr, keys)
    elif ("инвази" in m or "чужерод" in nm or "инвази" in nm) and "троп" not in m:
        rule_invasive(data, res, measure)
    elif "точек видов" in m or ("ликвидация троп" in m and "50" in m):
        rule_rare_trails(data, res, measure)
    elif "периметр" in m or "берег" in nm or "водо" in m:
        rule_shore_trails(data, res, measure)
    elif "скольз" in m and "троп" in m:
        rule_trail_hotspot(data, res, measure)
    elif "плотност" in m and ("дорог" in m or "троп" in m):
        rule_trail_density_vydel(data, res, measure)
    elif "лугов" in m:
        rule_type(data, res, measure, TYPE_MEADOW, "луговые")
    elif "околовод" in m:
        rule_type(data, res, measure, TYPE_WATER + TYPE_NEARWATER, "водные и околоводные")
    elif "водн" in m:
        rule_type(data, res, measure, TYPE_WATER, "водные")
    elif "лесн" in m:
        rule_type(data, res, measure, TYPE_FOREST, "лесные")
    else:
        log.warning("  метод не распознан: «%s» — пропуск", measure.get("method"))
        return False
    return True


# =====================================================================================
#                              ТАБЛИЦА МЕРОПРИЯТИЙ (Excel)
# =====================================================================================
def read_measures(path: str) -> list[dict]:
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]

    def col(*names):
        for n in names:
            for c in df.columns:
                if c.lower().startswith(n.lower()):
                    return c
        return None

    c_name, c_on, c_method = col("Мероприяти"), col("Вкл"), col("Метод")
    c_group, c_eco, c_prikaz, c_season = col("Тип мероприят"), col("Тип экосистем"), col("Соответствие"), col("Сезон")
    if not (c_name and c_on and c_method):
        raise ValueError(f"В Excel не найдены колонки «Мероприятия», «Вкл/выкл», «Метод выявления»: {list(df.columns)}")
    df[c_group] = df[c_group].ffill() if c_group else ""
    out = []
    for _, r in df.iterrows():
        if pd.isna(r[c_name]):
            continue
        on = str(r[c_on]).strip().lower() in ("1", "1.0", "true", "да", "yes", "вкл")
        out.append({
            "name": re.sub(r"\s+", " ", str(r[c_name])).strip(),
            "on": on,
            "method": None if pd.isna(r[c_method]) else str(r[c_method]).strip(),
            "group": re.sub(r"\s+", " ", str(r[c_group])).strip() if c_group and not pd.isna(r[c_group]) else "",
            "eco": str(r[c_eco]).strip() if c_eco and not pd.isna(r[c_eco]) else "",
            "prikaz": str(r[c_prikaz]).strip() if c_prikaz and not pd.isna(r[c_prikaz]) else "",
            "season": str(r[c_season]).strip() if c_season and not pd.isna(r[c_season]) else "",
        })
    log.info("Таблица мероприятий: %d строк, включено %d", len(out), sum(m["on"] for m in out))
    return out


# =====================================================================================
#                                       ОТЧЁТ
# =====================================================================================
LAYER_STYLE = {
    # layer → (цвет, тип: point/line/poly)
    "тропы в зонах повышенной плотности": ("#d62728", "line"),
    "зоны повышенной плотности троп (верхний квартиль)": ("#ff7f0e", "poly"),
    f"тропы на расстоянии {SHORE_BUFFER_M:.0f} м от водных объектов (Q75)": ("#1f77b4", "poly"),
    "водные объекты (линии)": ("#3182bd", "line"),
    "водные объекты (полигоны)": ("#9ecae1", "poly"),
    f"тропы на расстоянии {RARE_BUFFER_M:.0f} м от редких видов": ("#9467bd", "line"),
    "редкие виды — точки": ("#2ca02c", "point"),
    "редкие виды — полигоны": ("#2ca02c", "poly"),
    "редкие виды — SDM": ("#98df8a", "poly"),
    "редкие виды": ("#2ca02c", "poly"),
    "инвазионные виды — точки": ("#e377c2", "point"),
    "инвазионные виды — полигоны": ("#e377c2", "poly"),
    f"инвазионные виды — SDM (≥ {SDM_THRESHOLD_INVASIVE:.2f})": ("#f7b6d2", "poly"),
    "Corvus — точки": ("#000000", "point"),
    "плотность Corvus (Q75)": ("#7f7f7f", "poly"),
}
TYPE_COLORS = {"лес": "#b7dfb0", "луг": "#f5e9a6", "околовод": "#bfe3ea", "водн": "#9ec9ff"}


_TYPE_FALLBACK = ["#e6d3f5", "#f7d6c4", "#d9d9d9", "#cfe8cf", "#ffe0e0"]


def _type_color(t: str):
    t = str(t).lower()
    for k, c in TYPE_COLORS.items():
        if k in t:
            return c
    # неизвестный тип выдела: устойчивый цвет из резервной палитры (чтобы попал в легенду)
    return _TYPE_FALLBACK[hash(t) % len(_TYPE_FALLBACK)]


def _slope_style(layer):
    if layer.startswith("уклон"):
        # 50% и 70% чёрного соответственно.
        return ("#808080" if "5%" in layer else "#4D4D4D", "slope")
    return LAYER_STYLE.get(layer, ("#17becf", "poly"))


_BASEMAP_PROVIDER: list = []


def _basemap_candidates():
    import contextily as cx
    # Любой загруженный растр ниже принудительно переводится в серый.
    # OSM ставим первым: CartoDB Positron в некоторых средах показывает водяной знак API KEY REQUIRED.
    return [cx.providers.OpenStreetMap.Mapnik, cx.providers.Esri.WorldGrayCanvas]


def pick_basemap_provider(skip=0):
    """Выбор доступного тайлового источника; отображение затем переводится в оттенки серого."""
    if _BASEMAP_PROVIDER and not skip:
        return _BASEMAP_PROVIDER[0]
    import contextily as cx
    import requests
    cx.tile.USER_AGENT = "oozt_recs/1.0 (natural-area management report generator)"
    try:                                   # кэш тайлов: повторные прогоны не тянут их заново
        cx.set_cache_dir(str(Path(tempfile.gettempdir()) / "oozt_tiles"))
    except Exception:
        pass
    candidates = _basemap_candidates()[skip:]
    if skip:
        _BASEMAP_PROVIDER.clear()
    for prov in candidates:
        try:
            url = prov.build_url(x=39530, y=20465, z=16)   # тайл над Москвой
            r = requests.get(url, headers={"user-agent": cx.tile.USER_AGENT}, timeout=15)
            if r.status_code == 200 and r.headers.get("content-type", "").startswith("image"):
                _BASEMAP_PROVIDER.append(prov)
                log.info("Подложка: %s", prov.get("name"))
                return prov
            log.warning("Тайлы %s недоступны (HTTP %s)", prov.get("name"), r.status_code)
        except Exception as e:
            log.warning("Тайлы %s недоступны: %s", prov.get("name"), e)
    _BASEMAP_PROVIDER.append(None)
    return None


def desaturate_last_basemap(ax):
    """Перевод последнего растрового слоя осей в оттенки серого."""
    if not ax.images:
        return
    im = ax.images[-1]
    arr = np.asarray(im.get_array())
    if arr.ndim != 3 or arr.shape[2] < 3:
        return
    rgb = arr[..., :3].astype(float)
    gray = 0.2126 * rgb[..., 0] + 0.7152 * rgb[..., 1] + 0.0722 * rgb[..., 2]
    out = np.repeat(gray[..., None], 3, axis=2)
    if arr.shape[2] == 4:
        out = np.dstack([out, arr[..., 3]])
    im.set_data(out.astype(arr.dtype))


def _deg_label(v: float, step: float) -> str:
    """Подпись координатной сетки в десятичных градусах WGS 84."""
    nd = max(3, int(np.ceil(-np.log10(max(step, 1e-12)))) + 1)
    return f"{v:.{nd}f}°"


def set_degree_ticks(ax, ext):
    """Оси подписываются градусами широты/долготы (WGS 84), сама карта остаётся в UTM."""
    from pyproj import Transformer
    fwd = Transformer.from_crs(f"EPSG:{WORK_EPSG}", "EPSG:4326", always_xy=True)
    inv = Transformer.from_crs("EPSG:4326", f"EPSG:{WORK_EPSG}", always_xy=True)
    xc, yc = (ext[0] + ext[1]) / 2, (ext[2] + ext[3]) / 2
    lon0, lat0 = fwd.transform(ext[0], yc)
    lon1, lat1 = fwd.transform(ext[1], yc)
    _, latmin = fwd.transform(xc, ext[2])
    _, latmax = fwd.transform(xc, ext[3])
    def ticks(v0, v1, to_axis):
        span = max(v1 - v0, 1e-12)
        target = span / 4
        magnitude = 10 ** np.floor(np.log10(target))
        step = next(m * magnitude for m in (1, 2, 2.5, 5, 10)
                    if m * magnitude >= target)
        first = np.ceil(v0 / step) * step
        vals = np.arange(first, v1 + step / 2, step)
        return [to_axis(v) for v in vals], [_deg_label(v, step) for v in vals]

    xt, xl = ticks(lon0, lon1, lambda v: inv.transform(v, (lat0 + lat1) / 2)[0])
    yt, yl = ticks(latmin, latmax, lambda v: inv.transform((lon0 + lon1) / 2, v)[1])
    ax.set_xticks(xt); ax.set_xticklabels(xl)
    ax.set_yticks(yt); ax.set_yticklabels(yl)
    ax.set_xlabel("Долгота, в.д. (WGS 84)", fontsize=12)
    ax.set_ylabel("Широта, с.ш.", fontsize=12)


def _legend_label(layer: str) -> str:
    """Перенос длинных подписей, чтобы легенда не растягивала схему."""
    if len(layer) > 45 and " м от " in layer:
        return layer.replace(" м от ", " м\nот ", 1)
    return layer


def draw_kvartal(data: Data, evidence: gpd.GeoDataFrame | None, kv, oozt_name, out_png: Path, basemap=True):
    vyd = data.oozt[(data.oozt["KV"] == kv) & (data.oozt["OOZT"] == oozt_name)]
    if vyd.empty:
        return None
    focus_geom = unary_union(vyd.geometry.values)
    minx, miny, maxx, maxy = vyd.total_bounds
    pad = max((maxx - minx), (maxy - miny)) * 0.12 + 30
    ext = (minx - pad, maxx + pad, miny - pad, maxy + pad)
    w, h = ext[1] - ext[0], ext[3] - ext[2]
    fig_w = 11
    fig_h = max(6, min(11, fig_w * h / w))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=200)
    plt.rcParams.update({"font.size": 12})

    # Лесные выделы имеют зелёную заливку, луговые прозрачны.
    # Границы всех выделов повторно рисуются поверх тематических слоёв ниже.
    forest = vyd[vyd["Type"].astype(str).str.lower().str.contains("лес", regex=False)]
    if not forest.empty:
        forest.plot(ax=ax, color="#b7dfb0", edgecolor="none", alpha=0.55, zorder=1)
    # все тропы в окне карты — тонкий серый контекстный слой
    if data.trails is not None and not data.trails.empty:
        map_box = box(ext[0], ext[2], ext[1], ext[3])
        all_trails = data.trails[data.trails.intersects(map_box)]
        if not all_trails.empty:
            all_trails.plot(ax=ax, color="#707070", linewidth=0.7, alpha=0.8, zorder=2)
    # основания
    handles = []
    if data.trails is not None and not data.trails.empty:
        handles.append(Line2D([], [], color="#707070", linewidth=1.0, label="все тропы"))
    used_layers = []
    if evidence is not None and not evidence.empty:
        kv_geom = unary_union(vyd.geometry.values)
        ev = evidence[evidence.intersects(kv_geom.buffer(5))]
        for layer, grp in ev.groupby("layer"):
            color, kind = _slope_style(layer)
            if kind == "point":
                species_point = layer in ("редкие виды — точки", "инвазионные виды — точки")
                point_size = 60 if species_point else 40
                legend_size = 13.5 if species_point else 9
                grp.plot(ax=ax, color=color, markersize=point_size,
                         edgecolor="white", linewidth=0.6, zorder=6)
                handles.append(Line2D([], [], marker="o", color="w", markerfacecolor=color,
                                      markersize=legend_size, label=_legend_label(layer)))
            elif kind == "line":
                grp.plot(ax=ax, color=color, linewidth=2.2, zorder=5)
                handles.append(Line2D([], [], color=color, linewidth=2.5,
                                      label=_legend_label(layer)))
            elif kind == "slope":
                grp.plot(ax=ax, facecolor=color, edgecolor=color,
                         alpha=1.0, linewidth=0.8, zorder=4)
                handles.append(Patch(facecolor=color, edgecolor=color,
                                     label=_legend_label(layer)))
            elif "SDM" in layer:      # модельные ареалы занимают большие площади — только штриховка
                grp.plot(ax=ax, facecolor="none", edgecolor=color, alpha=0.9,
                         linewidth=0.9, hatch="////", zorder=4)
                handles.append(Patch(facecolor="none", edgecolor=color, hatch="////",
                                     label=_legend_label(layer)))
            else:
                grp.plot(ax=ax, facecolor=color, edgecolor=color, alpha=0.45, linewidth=0.8, zorder=4)
                handles.append(Patch(facecolor=color, alpha=0.5, edgecolor=color,
                                     label=_legend_label(layer)))
            used_layers.append(layer)
    # границы выделов поверх всех тематических слоёв
    vyd.boundary.plot(ax=ax, color="#333333", linewidth=1.15, zorder=8)
    # подписи выделов
    for _, r in vyd.iterrows():
        p = r.geometry.representative_point()
        ax.annotate(str(r["VYD"]), (p.x, p.y), ha="center", va="center", fontsize=13, fontweight="bold",
                    color="#111111", zorder=11)
    # соседние кварталы: границы и номера в пределах окна карты
    map_box = box(ext[0], ext[2], ext[1], ext[3])
    same_oozt = data.oozt[data.oozt["OOZT"] == oozt_name]
    neighbors = same_oozt[(same_oozt["KV"] != kv) & same_oozt.intersects(map_box)]
    for n_kv, ng in neighbors.groupby("KV"):
        ngeom = unary_union(ng.geometry.values)
        clipped = ngeom.intersection(map_box)
        if clipped.is_empty:
            continue
        gpd.GeoSeries([ngeom], crs=WORK_EPSG).boundary.plot(
            ax=ax, color="#444444", linewidth=1.5, linestyle="--", zorder=9)
        p = clipped.representative_point()
        px = float(np.clip(p.x, ext[0] + w * 0.07, ext[1] - w * 0.07))
        py = float(np.clip(p.y, ext[2] + h * 0.09, ext[3] - h * 0.07))
        n_label = str(n_kv).removesuffix(".0")
        ax.annotate(n_label, (px, py), ha="center", va="center",
                    fontsize=16, fontweight="bold", color="#222222", zorder=12)
    # граница квартала в фокусе
    gpd.GeoSeries([focus_geom], crs=WORK_EPSG).boundary.plot(
        ax=ax, color="black", linewidth=2.4, zorder=10)
    ax.set_xlim(ext[0], ext[1]); ax.set_ylim(ext[2], ext[3])
    if basemap:
        try:
            import contextily as cx
            for _attempt in range(3):
                prov = pick_basemap_provider(skip=_attempt)
                if prov is None:
                    break
                try:
                    cx.add_basemap(ax, crs=f"EPSG:{WORK_EPSG}", source=prov,
                                   attribution=prov.get("attribution", "© OpenStreetMap"),
                                   attribution_size=10, zorder=0)
                    desaturate_last_basemap(ax)
                    break
                except Exception as e:     # таймаут/лимит запросов — пробуем следующий источник
                    log.warning("подложка %s недоступна (%s) — пробую другой источник", prov.get("name"), e)
        except Exception as e:
            log.warning("подложка OSM недоступна: %s", e)
    ax.legend(handles=handles, loc="upper left", bbox_to_anchor=(1.01, 1.0), fontsize=12, frameon=True,
              title="Легенда", title_fontsize=12)
    ax.set_title(f"{oozt_name}. Квартал {kv}", fontsize=15, fontweight="bold")
    ax.tick_params(labelsize=12)
    try:
        set_degree_ticks(ax, ext)
    except Exception as e:                # запасной вариант — метры UTM
        log.warning("не удалось построить сетку в градусах (%s) — оси в метрах", e)
        ax.set_xlabel("E, м (UTM 37N)", fontsize=12); ax.set_ylabel("N, м", fontsize=12)
        ax.ticklabel_format(style="plain", useOffset=False)
    # масштабная линейка
    target = w / 4                        # «круглая» длина линейки: 1/2/5 × 10^k
    dec = 10 ** np.floor(np.log10(target))
    L = next(m * dec for m in (5, 2, 1) if m * dec <= target) if target >= dec else dec
    x0, y0 = ext[0] + w * 0.04, ext[2] + h * 0.04
    ax.plot([x0, x0 + L], [y0, y0], color="black", linewidth=4, zorder=11)
    ax.text(x0 + L / 2, y0 + h * 0.012, f"{int(L)} м", ha="center", fontsize=12, zorder=11,
            bbox=dict(fc="white", ec="none", alpha=0.7, pad=1))
    fig.tight_layout()
    fig.savefig(out_png, bbox_inches="tight")
    plt.close(fig)
    return used_layers


def md_table(df: pd.DataFrame, cols: list[str]) -> str:
    def esc(x):
        s = "" if pd.isna(x) else str(x)
        return s.replace("|", "／").replace("\n", " ")
    lines = ["| " + " | ".join(cols) + " |", "|" + "|".join(["---"] * len(cols)) + "|"]
    for _, r in df.iterrows():
        lines.append("| " + " | ".join(esc(r[c]) for c in cols) + " |")
    return "\n".join(lines)


def build_report(data: Data, table: pd.DataFrame, evidence, out_dir: Path, kv_list: list[int], basemap=True):
    enable_win_dll_path()
    rep = out_dir / "report"
    rep.mkdir(parents=True, exist_ok=True)
    kv_all = data.oozt[["OOZT", "KV"]].drop_duplicates().sort_values(["OOZT", "KV"])
    if kv_list:
        kv_all = kv_all[kv_all["KV"].isin(kv_list)]
    md = [f"# Природоохранные мероприятия по кварталам ООЗТ\n",
          f"Слой ООЗТ: `{Path(data.cfg.oozt_path).name}`; таблица мероприятий: `{Path(data.cfg.measures_xlsx).name}`.  ",
          f"Всего назначений: {len(table)}; выделов с мероприятиями: "
          f"{table[['ООЗТ', 'Квартал', 'Выдел']].drop_duplicates().shape[0] if not table.empty else 0}.\n",
          "Параметры: окно плотности троп %.0f м, окно плотности врановых %.0f м, буфер от точек редких видов %.0f м, "
          "порог SDM %.2f (инвазионные) / %.2f (редкие), верхний квартиль — по всем ООЗТ.\n"
          % (TRAIL_WINDOW_M, CORVUS_WINDOW_M, RARE_BUFFER_M, SDM_THRESHOLD_INVASIVE, SDM_THRESHOLD_RARE)]
    # сводка по мероприятиям
    if not table.empty:
        summ = table.groupby(["Тип мероприятий", "Мероприятие"]).agg(
            Выделов=("Выдел", "size"), Кварталов=("Квартал", "nunique")).reset_index()
        md.append("## Сводка по мероприятиям\n")
        md.append(md_table(summ, ["Тип мероприятий", "Мероприятие", "Выделов", "Кварталов"]) + "\n")
    report_table = table.rename(columns={
        "Экосистема": "Экосистема выдела",
        "Мероприятие": "Мероприятия",
        "Тип экосистемы мероприятия": "Тип экосистемы",
        "Пункт Приказа": "Соответствие Пункту Приказа",
    })
    group_cols = ["Тип мероприятий", "Тип экосистемы", "Соответствие Пункту Приказа"]
    cols = ["Выдел", "Экосистема выдела", "Мероприятия", "Основание"]
    for _, kr in kv_all.iterrows():
        oozt, kv = kr["OOZT"], kr["KV"]
        png = rep / f"fig_{re.sub(r'[^\w]+', '_', str(oozt))}_kv_{kv}.png"
        log.info("Схема: %s кв. %s", oozt, kv)
        draw_kvartal(data, evidence, kv, oozt, png, basemap=basemap)
        sub = report_table[
            (report_table["ООЗТ"] == oozt) & (report_table["Квартал"] == kv)
        ] if not report_table.empty else report_table
        md.append(f"\n## {oozt}. Квартал {kv}\n")
        md.append(f"![Квартал {kv}]({png.relative_to(rep).as_posix()})\n")
        vyd = data.oozt[(data.oozt["KV"] == kv) & (data.oozt["OOZT"] == oozt)]
        md.append(f"Выделов: {len(vyd)}; площадь {vyd['area_ha'].sum():.1f} га; "
                  f"назначено мероприятий: {len(sub)}.\n")
        if sub.empty:
            md.append("_Мероприятия по загруженным данным не назначены._\n")
        else:
            md.append("### Мероприятия по приказу\n")
            grouped = sub.sort_values(group_cols + ["Выдел"]).groupby(
                group_cols, sort=False, dropna=False
            )
            for (measure_type, ecosystem_type, order_clause), group in grouped:
                md.append(f"#### Тип мероприятий: {measure_type}\n")
                md.append(f"##### Тип экосистемы: {ecosystem_type}\n")
                md.append(
                    f"###### Соответствие Пункту Приказа: {order_clause}\n"
                )
                md.append(md_table(group.sort_values(["Выдел"]), cols) + "\n")
    (rep / "report.md").write_text("\n".join(md), encoding="utf-8")
    # DOCX: сначала pandoc (если установлен), иначе python-docx (если установлен)
    done = False
    try:
        import shutil, subprocess
        if shutil.which("pandoc"):
            subprocess.run(["pandoc", "report.md", "-o", "report.docx", "--resource-path", "."],
                           cwd=rep, check=True, timeout=300)
            log.info("Отчёт сконвертирован в DOCX (pandoc)")
            done = True
    except Exception as e:
        log.info("pandoc недоступен (%s)", e)
    if not done:
        try:
            build_docx(data, table, kv_all, rep, cols)
            log.info("Отчёт записан в DOCX (python-docx)")
        except ImportError:
            log.info("python-docx не установлен — отчёт только в Markdown (pip install python-docx)")
        except Exception as e:
            log.warning("DOCX не создан: %s", e)
    return rep / "report.md"


def build_docx(data: Data, table: pd.DataFrame, kv_all: pd.DataFrame, rep: Path, cols: list[str]):
    """Word-версия отчёта: заголовок, сводка, по кварталу — схема (ширина страницы) + таблица."""
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.enum.section import WD_ORIENT

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(1.5))
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"; st.font.size = Pt(11)

    def add_table(df: pd.DataFrame, columns: list[str], font_pt=10, widths_cm=None):
        t = doc.add_table(rows=1, cols=len(columns))
        t.style = "Table Grid"
        t.autofit = widths_cm is None
        for i, c in enumerate(columns):
            cell = t.rows[0].cells[i]; cell.text = ""
            run = cell.paragraphs[0].add_run(str(c)); run.bold = True; run.font.size = Pt(font_pt)
        for _, r in df.iterrows():
            cells = t.add_row().cells
            for i, c in enumerate(columns):
                v = r[c]
                cells[i].text = ""
                run = cells[i].paragraphs[0].add_run("" if pd.isna(v) else str(v)); run.font.size = Pt(font_pt)
        if widths_cm:
            for row in t.rows:
                for i, w in enumerate(widths_cm):
                    row.cells[i].width = Cm(w)
        return t

    doc.add_heading("Природоохранные мероприятия по кварталам ООЗТ", level=0)
    doc.add_paragraph(f"Слой ООЗТ: {Path(data.cfg.oozt_path).name}; таблица мероприятий: "
                      f"{Path(data.cfg.measures_xlsx).name}. Всего назначений: {len(table)}.")
    doc.add_paragraph("Параметры: окно плотности троп %.0f м, окно плотности врановых %.0f м, буфер от точек редких "
                      "видов %.0f м, порог SDM %.2f (инвазионные) / %.2f (редкие), верхний квартиль — "
                      "по всем ООЗТ."
                      % (TRAIL_WINDOW_M, CORVUS_WINDOW_M, RARE_BUFFER_M,
                         SDM_THRESHOLD_INVASIVE, SDM_THRESHOLD_RARE))
    if not table.empty:
        doc.add_heading("Сводка по мероприятиям", level=1)
        summ = table.groupby(["Тип мероприятий", "Мероприятие"]).agg(
            Выделов=("Выдел", "size"), Кварталов=("Квартал", "nunique")).reset_index()
        add_table(summ, ["Тип мероприятий", "Мероприятие", "Выделов", "Кварталов"])
    for _, kr in kv_all.iterrows():
        oozt, kv = kr["OOZT"], kr["KV"]
        png = rep / f"fig_{re.sub(r'[^\w]+', '_', str(oozt))}_kv_{kv}.png"
        doc.add_page_break()
        doc.add_heading(f"{oozt}. Квартал {kv}", level=1)
        if png.exists():
            doc.add_picture(str(png), width=Cm(26))
        doc_table = table.rename(columns={
            "Экосистема": "Экосистема выдела",
            "Мероприятие": "Мероприятия",
            "Тип экосистемы мероприятия": "Тип экосистемы",
            "Пункт Приказа": "Соответствие Пункту Приказа",
        })
        sub = doc_table[
            (doc_table["ООЗТ"] == oozt) & (doc_table["Квартал"] == kv)
        ] if not doc_table.empty else doc_table
        vyd = data.oozt[(data.oozt["KV"] == kv) & (data.oozt["OOZT"] == oozt)]
        doc.add_paragraph(f"Выделов: {len(vyd)}; площадь {vyd['area_ha'].sum():.1f} га; назначено мероприятий: {len(sub)}.")
        if sub.empty:
            doc.add_paragraph("Мероприятия по загруженным данным не назначены.")
        else:
            doc.add_heading("Мероприятия по приказу", level=2)
            grouped = sub.sort_values(group_cols + ["Выдел"]).groupby(
                group_cols, sort=False, dropna=False
            )
            for (measure_type, ecosystem_type, order_clause), group in grouped:
                doc.add_heading(f"Тип мероприятий: {measure_type}", level=3)
                doc.add_heading(f"Тип экосистемы: {ecosystem_type}", level=4)
                doc.add_heading(
                    f"Соответствие Пункту Приказа: {order_clause}", level=5
                )
                add_table(group.sort_values(["Выдел"]), cols, font_pt=8,
                          widths_cm=[1.2, 4.5, 7.0, 11.0])
    doc.save(str(rep / "report.docx"))


# =====================================================================================
#                                       MAIN
# =====================================================================================
def _setup_logging(out: Path, name="run.log"):
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s  %(message)s", datefmt="%H:%M:%S",
                        handlers=[logging.StreamHandler(sys.stdout),
                                  logging.FileHandler(out / name, encoding="utf-8", mode="w")])


def report_only(cfg: Config):
    """Перестроить отчёт по уже рассчитанным measures_long.csv и evidence.gpkg (без пересчёта правил)."""
    out = Path(cfg.out_dir)
    _setup_logging(out, "report.log")
    data = Data(cfg, oozt_only=True)
    # Для схем нужен полный контекст троп, а не только фрагменты,
    # послужившие основанием для назначения мероприятий.
    data.trails = data._load_trails()
    table = pd.read_csv(out / "measures_long.csv", encoding="utf-8-sig")
    for c in ("Квартал", "Выдел"):
        table[c] = pd.to_numeric(table[c], errors="coerce").astype("Int64")
    gpkg = out / "evidence.gpkg"
    parts = []
    for ly, _ in pyogrio.list_layers(gpkg):
        if ly.startswith("evidence_"):
            g = gpd.read_file(gpkg, layer=ly)
            parts.append(g.to_crs(WORK_EPSG))
    ev = gpd.GeoDataFrame(pd.concat(parts, ignore_index=True), crs=WORK_EPSG) if parts else None
    rep = build_report(data, table, ev, out, cfg.report_kv, basemap=cfg.basemap)
    log.info("Готово. Отчёт: %s", rep)


def run(cfg: Config):
    out = Path(cfg.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    _setup_logging(out)
    measures = read_measures(cfg.measures_xlsx)
    data = Data(cfg)
    cfg.save(out / "config_used.json")
    res = Results()
    for m in measures:
        if not m["on"]:
            continue
        try:
            dispatch_rule(m, data, res)
        except Exception as e:
            log.exception("Ошибка в мероприятии «%s»: %s", m["name"], e)

    table = res.long_table()
    if table.empty:
        log.warning("Ни одно мероприятие не назначено")
        table = pd.DataFrame(columns=["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема", "Тип мероприятий",
                                      "Мероприятие", "Тип экосистемы мероприятия", "Пункт Приказа",
                                      "Сезонность", "Основание", "Значение"])
    table.to_csv(out / "measures_long.csv", index=False, encoding="utf-8-sig")
    try:
        with pd.ExcelWriter(out / "measures_long.xlsx", engine="openpyxl") as xw:
            table.to_excel(xw, index=False, sheet_name="long")
            if not table.empty:
                order_cols = ["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема",
                              "Тип мероприятий", "Мероприятие", "Тип экосистемы мероприятия",
                              "Пункт Приказа"]
                by_order = table[order_cols].rename(columns={
                    "Мероприятие": "Мероприятия",
                    "Тип экосистемы мероприятия": "Тип экосистемы",
                    "Пункт Приказа": "Соответствие Пункту Приказа",
                })
                by_order.to_excel(xw, index=False, sheet_name="мероприятия по приказу")

                # Значения не дедуплицируем: каждая строка во всех четырёх колонках
                # должна соответствовать одному и тому же мероприятию.
                join_lines = lambda s: "\n".join(
                    "" if pd.isna(x) else str(x) for x in s
                )
                wide = table.groupby(
                    ["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема"],
                    as_index=False, dropna=False
                ).agg(**{
                    "Тип мероприятий": ("Тип мероприятий", join_lines),
                    "Мероприятия": ("Мероприятие", join_lines),
                    "Тип экосистемы": ("Тип экосистемы мероприятия", join_lines),
                    "Соответствие Пункту Приказа": ("Пункт Приказа", join_lines),
                    "Число": ("Мероприятие", "size"),
                })
                wide.to_excel(xw, index=False, sheet_name="by_vydel")
                from openpyxl.styles import Alignment, Font, PatternFill
                for ws in (xw.book["мероприятия по приказу"], xw.book["by_vydel"]):
                    # Двухуровневый заголовок: четыре колонки объединены
                    # общей шапкой «Мероприятия по приказу».
                    ws.insert_rows(1)
                    ws.merge_cells("A1:E1")
                    ws["A1"] = "Территориальная привязка"
                    ws.merge_cells("F1:I1")
                    ws["F1"] = "Мероприятия по приказу"
                    if ws.max_column >= 10:
                        ws["J1"] = "Число мероприятий"
                    ws.freeze_panes = "A3"
                    ws.auto_filter.ref = f"A2:{chr(64 + ws.max_column)}{ws.max_row}"
                    for cell in ws[1] + ws[2]:
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill("solid", fgColor="D9EAD3")
                        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                                   horizontal="center")
                    widths = [16, 11, 10, 16, 28, 34, 55, 22, 25, 10]
                    for i, width in enumerate(widths[:ws.max_column], 1):
                        ws.column_dimensions[chr(64 + i)].width = width
                    for row in ws.iter_rows(min_row=3):
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
    except Exception as e:
        log.warning("Excel не записан (%s) — есть CSV", e)

    ev = res.evidence_gdf()
    try:
        gpkg = out / "evidence.gpkg"
        if gpkg.exists():
            gpkg.unlink()
        vm = data.oozt.copy()
        if not table.empty:
            agg = table.groupby(["ООЗТ", "Квартал", "Выдел"])["Мероприятие"].agg(lambda s: "; ".join(dict.fromkeys(s)))
            vm = vm.merge(agg.rename("measures").reset_index().rename(
                columns={"ООЗТ": "OOZT", "Квартал": "KV", "Выдел": "VYD"}), on=["OOZT", "KV", "VYD"], how="left")
        vm.drop(columns=["__src"], errors="ignore").to_crs(4326).to_file(
            gpkg, layer="vydel_measures", driver="GPKG")
        if ev is not None and not ev.empty:
            for gt, grp in ev.groupby(ev.geom_type.str.replace("Multi", "")):
                grp.to_crs(4326).to_file(
                    gpkg, layer=f"evidence_{gt.lower()}", driver="GPKG")
        log.info("Слои-основания записаны: %s", gpkg)
    except Exception as e:
        log.warning("GPKG не записан: %s", e)

    rep = build_report(data, table, ev, out, cfg.report_kv, basemap=cfg.basemap)
    log.info("Готово. Таблица: %s; отчёт: %s", out / "measures_long.xlsx", rep)
    return table, rep


def main():
    ap = argparse.ArgumentParser(description="Природоохранные мероприятия для ООЗТ по кварталам/выделам")
    ap.add_argument("--config", help="JSON-конфигурация (без диалогов)")
    ap.add_argument("--kv", help="кварталы для отчёта, через запятую (переопределяет конфиг)")
    ap.add_argument("--no-basemap", action="store_true", help="не добавлять подложку OSM")
    ap.add_argument("--save-config-only", action="store_true", help="только собрать конфиг диалогами и сохранить")
    ap.add_argument("--report-only", action="store_true",
                    help="не пересчитывать: построить отчёт по measures_long.csv и evidence.gpkg из out_dir")
    args = ap.parse_args()

    if args.config:
        cfg = Config.load(args.config)
    else:
        cfg = gui_collect_config()
    if args.kv:
        cfg.report_kv = [int(x) for x in re.findall(r"\d+", args.kv)]
    if args.no_basemap:
        cfg.basemap = False
    if args.save_config_only:
        Path(cfg.out_dir).mkdir(parents=True, exist_ok=True)
        cfg.save(Path(cfg.out_dir) / "config.json")
        print("Конфигурация сохранена:", Path(cfg.out_dir) / "config.json")
        return
    if args.report_only:
        report_only(cfg)
        return
    run(cfg)


if __name__ == "__main__":
    main()
