# -*- coding: utf-8 -*-
"""
oozt_recs.py — подготовка природоохранных рекомендаций (мероприятий) для ООЗТ Москвы
на основе ГИС-слоёв и таблицы «Мероприятия по Приказу».

Единицы назначения мероприятий (режим выбирается в окне или ключом unit_mode в конфигурации)
--------------------------------------------------------------------------------------------
    kv_vyd  — кварталы и выделы: поля ООЗТ, квартала, выдела, типа; отчёт и схемы — по кварталам;
    kv_cut  — экосистемы, разрезанные квартальной сетью: слой экосистем + вектор квартальной сети
             (kvnet_path, поле номера квартала kvnet_field); экосистемы режутся границами кварталов,
             в каждом квартале выделям присваиваются номера с 1 — с северо-запада, с запада на восток,
             с севера на юг; отчёт и схемы — по кварталам (квартал X, выдел Y);
    eco     — слой экосистем с уникальным ID, без кварталов: поля ID и типа (ООЗТ и название — по желанию);
             отчёт и схемы — по листам регулярной сетки (sheet_size_m, по умолчанию 1000 м)
             или по значениям поля группировки (report_group_field).

Запуск
------
    python oozt_recs.py                   # окно выбора данных: режим, слои, живая сводка мероприятий
    python oozt_recs.py --config cfg.json # повторный запуск с сохранённой конфигурацией, без окна
    python oozt_recs.py --config cfg.json --gui        # окно, предзаполненное из конфигурации
    python oozt_recs.py --config cfg.json --kv 1,2,5   # отчёт только по выбранным кварталам / листам
    python oozt_recs.py --config cfg.json --report-only --formats docx,pdf,doc   # пересобрать отчёт

Выходные данные (папка out_dir)
-------------------------------
    measures_long.xlsx / .csv   — длинная таблица ООЗТ–Квартал(Лист)–Выдел(ID)–Мероприятие (+ основание)
    evidence.gpkg               — полигоны с мероприятиями (vydel_measures / eco_measures), measures (полигон × мероприятие),
                                  слои-основания evidence_*
    oozt_measures.qgs           — проект QGIS с этими слоями и стилями (относительные пути)
    report/report.md            — отчёт по кварталам или листам (схема + таблицы), рисунки report/fig_*.png
    report/report.docx, .pdf    — тот же отчёт в Word и PDF (report_formats / --formats; .doc — по запросу)
    config_used.json            — конфигурация запуска (можно передать в --config)
    run.log                     — журнал, в начале — статус каждого мероприятия при выбранных данных

Правила (ключ Метод выявления в Excel → расчёт) — classify_measure() и раздел ПРАВИЛА;
связь правил со слоями — RULES (data / opt). Обязательны только таблица мероприятий и слой территории.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import warnings
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Iterable

# ---- локальная папка pylib рядом со скриптом (чистые Python-пакеты: openpyxl, contextily, python-docx) ----
# Jupyter: __file__ не определён — используем текущую папку ноутбука
try:
    _pylib = Path(__file__).resolve().parent / "pylib"
except NameError:
    _pylib = Path.cwd() / "pylib"
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
from shapely.geometry import shape, box, Point, LineString
from shapely.ops import unary_union
from shapely.validation import make_valid
import shapely
from scipy import ndimage

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D

log = logging.getLogger("oozt")


class _Progress:
    """Ход расчёта для окна прогресса: этап (диапазон %), доля внутри этапа, подпись.
    Без окна (запуск с --config) ничего не делает."""

    def __init__(self):
        self.q = None                     # queue.Queue окна прогресса или None
        self.lo, self.hi, self.v = 0.0, 100.0, 0.0
        self.title = ""

    def _emit(self, *item):
        if self.q is not None:
            self.q.put(item)

    def phase(self, title, lo, hi):
        self.lo, self.hi = float(lo), float(hi)
        self.v = max(self.v, self.lo)
        self.title = title
        self._emit("phase", title, self.v)

    def frac(self, f, title=None):
        f = min(1.0, max(0.0, float(f)))
        self.v = max(self.v, self.lo + (self.hi - self.lo) * f)
        if title:
            self.title = title
        self._emit("phase", self.title, self.v)

    def tick(self):
        """Сообщение журнала внутри этапа: полоска чуть продвигается, не выходя за этап."""
        if self.q is not None and self.hi > self.v:
            self.v += (self.hi - self.v) * 0.02
            self._emit("value", self.v)


PROG = _Progress()
logging.getLogger("pyogrio").setLevel(logging.WARNING)
logging.getLogger("rasterio").setLevel(logging.WARNING)


# =====================================================================================
#        ПОЧИНКА ГЕОМЕТРИЙ: невалидные полигоны (TopologyException) и многочастные
# =====================================================================================
def _keep_polygons_only(g: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Оставить только полигоны: буферные/сервисные слои должны состоять из полигонов,
    остальные типы (линии, точки) выбрасываются с записью в журнал."""
    if g is None or g.empty:
        return g
    ok = g.geom_type.isin(["Polygon", "MultiPolygon"])
    if not ok.all():
        log.warning("  отброшено %d неполигональных объектов", int((~ok).sum()))
        g = g[ok].copy()
    return g


def _poly_part(geom):
    """make_valid может вернуть коллекцию (полигоны + линии): берём полигональную часть."""
    if geom is not None and geom.geom_type == "GeometryCollection":
        polys = [g for g in geom.geoms if g.geom_type in ("Polygon", "MultiPolygon")]
        if polys:
            return polys[0] if len(polys) == 1 else unary_union(polys)
    return geom


def repair_geometries(g: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Починить невалидные геометрии: make_valid (для полигонов — с выделением
    полигональной части); пустые и NaN — выбросить. Тип слоя (линии, точки) сохраняется.
    Возвращает GeoDataFrame или None, если объектов не осталось."""
    if g is None:
        return None
    g = g[~g.geometry.isna() & ~g.geometry.is_empty]
    if g.empty:
        return None
    invalid = ~g.geometry.is_valid
    if invalid.any():
        log.info("  починено невалидных геометрий: %d", int(invalid.sum()))
        g = g.copy()

        def _fix(geom):
            try:
                v = make_valid(geom)
            except Exception:
                try:
                    v = geom.buffer(0)
                except Exception:
                    return None
            return _poly_part(v) if geom.geom_type in ("Polygon", "MultiPolygon") else v

        g.loc[invalid, "geometry"] = g.loc[invalid, "geometry"].apply(_fix)
        g = g[~g.geometry.isna() & ~g.geometry.is_empty]
    return g if not g.empty else None


def _safe_union(geoms) -> "BaseGeometry":
    """unary_union с починкой: make_valid по каждому; при TopologyException —
    buffer(0) по одному; если и это не спасает — пересечение коробок."""
    if geoms is None:
        return None
    try:
        return unary_union(geoms)
    except Exception:
        pass
    fixed, union = [], None
    for g in geoms:
        try:
            g2 = g if g.is_valid else make_valid(g)
        except Exception:
            try:
                g2 = g.buffer(0)
            except Exception:
                continue
        if g2 is None or g2.is_empty:
            continue
        fixed.append(g2)
        try:
            union = g2 if union is None else unary_union([union, g2])
        except Exception:
            try:
                union = g2 if union is None else union.union(g2.buffer(0))
            except Exception:
                continue
    if union is None:
        log.warning("  объединение геометрий не удалось даже с починкой — пусто")
        return None
    return union


def _safe_overlay(left: gpd.GeoDataFrame, right: gpd.GeoDataFrame, how: str = "intersection",
                  keep_geom_type: bool = True) -> gpd.GeoDataFrame:
    """gpd.overlay с починкой геометрий: при TopologyException обе стороны проходят
    через make_valid/buffer(0) и пересечение повторяется; итог — только полигоны."""
    try:
        return gpd.overlay(left, right, how=how, keep_geom_type=keep_geom_type)
    except Exception:
        pass
    log.warning("  пересечение слоёв упало (невалидные геометрии) — чиню и повторяю")
    a, b = repair_geometries(left), repair_geometries(right)
    if a is None or b is None:
        return gpd.GeoDataFrame(columns=["geometry"], geometry=[], crs=WORK_EPSG)
    try:
        j = gpd.overlay(a, b, how=how, keep_geom_type=keep_geom_type)
        return j
    except Exception:
        pass
    # последняя попытка: buffer(0) по каждой стороне
    try:
        a = a.copy(); a["geometry"] = a.geometry.buffer(0)
        b = b.copy(); b["geometry"] = b.geometry.buffer(0)
        return gpd.overlay(_keep_polygons_only(a), _keep_polygons_only(b), how=how,
                           keep_geom_type=keep_geom_type)
    except Exception as e:
        log.error("  пересечение слоёв не удалось даже после починки: %s", e)
        return gpd.GeoDataFrame(columns=["geometry"], geometry=[], crs=WORK_EPSG)

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
KV_SPLIT_GAP_M = 300             # м: части квартала дальше друг от друга считаются разными массивами
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

# Антропогенные луга: газоны (тип 503) и бурьян (тип 222); остальные луговые — естественные.
# Маркеры ищутся подстрокой в значении поля типа (в т. ч. с кодами: «Газоны 503», «бурьян 222»).
ANTHRO_MEADOW_MARKS = ("газон", "бурьян", "бурья", "503", "222")


# Структурные классы экосистем (поле «Структурный класс»; коды class_value — три последние цифры)
STRUCT_CODE_CLASSES = {
    **{c: "Широколиственные леса" for c in (110, 111, 112, 113, 121)},
    **{c: "Мелколиственные леса" for c in (131, 132, 134, 135, 171)},
    **{c: "Хвойные леса" for c in (141, 142)},
    **{c: "Смешанные леса" for c in (151, 152, 153, 161)},
    211: "Суходольные луга", 212: "Суходольные луга",
    221: "Мезофитные луга", 223: "Мезофитные луга", 224: "Мезофитные луга",
    222: "Мезофитные луга (антропогенные)",
    231: "Гидроморфные луга", 232: "Гидроморфные луга", 233: "Гидроморфные луга",
    310: "Болота", 400: "Вода", 501: "Антропогенные", 503: "Антропогенные",
}
# I.7 «Мозаичное выкашивание»: косим все луга, включая бурьян (222); гидроморфные луга (231–233) не косим
STRUCT_MOW_DEFAULT = ("Суходольные луга", "Мезофитные луга", "Мезофитные луга (антропогенные)")


def _norm_txt(s) -> str:
    """Текст для сравнения: строчные, ё → е, без лишних пробелов."""
    return re.sub(r"\s+", " ", str(s if s is not None else "").strip().lower().replace("ё", "е"))


def struct_class_of(value) -> str:
    """Значение поля «Структурный класс» → название класса. Число/код (class_value, например 5222 или 222)
    понимается по трём последним цифрам; текст возвращается как есть."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    t = str(value).strip()
    m = re.fullmatch(r"(\d+)(?:\.0+)?", t)
    if m:
        code = int(m.group(1)[-3:])
        return STRUCT_CODE_CLASSES.get(code, t)
    return t


def _anthro_meadow(text) -> bool:
    """Тип полигона — антропогенный луг (газоны 503 / бурьян 222)."""
    t = str(text or "").lower().replace("ё", "е")
    return any(m in t for m in ANTHRO_MEADOW_MARKS)

# =====================================================================================
#                                   КОНФИГУРАЦИЯ
# =====================================================================================
UNIT_MODES = {
    "kv_vyd": "Кварталы и выделы",
    "kv_cut": "Экосистемы, разрезанные квартальной сетью",
    "eco": "Слой экосистем с уникальным ID (без кварталов)",
}

# кандидаты для автоподбора поля номера квартала в слое квартальной сети
KVNET_KV_CANDIDATES = ["KV", "Квартал", "kv", "kvartal", "kvartal_no", "Kvartal", "квартал",
                       "nom_kv", "Номер квартала", "quarter", "kv_num", "num_kv", "FID_kv"]


@dataclass
class Config:
    measures_xlsx: str = ""
    unit_mode: str = "kv_vyd"            # kv_vyd — кварталы/выделы; kv_cut — экосистемы, разрезанные
                                         # квартальной сетью; eco — полигоны экосистем с уникальным ID
    oozt_path: str = ""                  # слой единиц территории (выделы или экосистемы)
    oozt_layer: str | None = None
    kvnet_path: str = ""                 # kv_cut: вектор квартальной сети
    kvnet_layer: str | None = None
    kvnet_field: str = ""                # kv_cut: поле с уникальным номером квартала в сети
    kvnet_oozt_field: str = ""           # kv_cut: поле ООЗТ в сети (необязательно)
    min_cut_ha: float = 0.01             # kv_cut: минимальная площадь куска экосистемы, га
    oozt_name: str = ""                  # eco / kv_cut: название ООЗТ, если в слое нет поля ООЗТ
    report_group_field: str | None = None  # eco: поле группировки отчёта; пусто — листы регулярной сетки
    sheet_size_m: float = 1000.0         # eco: размер листа сетки для схем, м
    dem_path: str | None = None
    trails: list[str] = field(default_factory=list)
    informal_trails: list = field(default_factory=lambda: ["auto"])   # ["auto"] — скачать из OSM (informal=yes); список путей — свой слой; [] — не использовать
    osm_trail_kinds: list = field(default_factory=lambda: ["informal"])  # что скачивать из OSM: informal — стихийные, unpaved — грунтовые
    osm_source: str = "auto"         # откуда брать тропы OSM: auto (Overpass → BBBike → Geofabrik) | overpass | bbbike | geofabrik | pbf
    osm_pbf: str = ""                # свой файл OSM (.osm.pbf / .osm) при osm_source = "pbf"
    osm_cache_dir: str = ""          # папка для скачанных .pbf (пусто — osm_cache рядом со скриптом)
    osm_cache_days: int = 30         # скачанный .pbf используется повторно, пока не старше N дней
    roads_paved: list[str] = field(default_factory=list)   # дороги с твёрдым покрытием: «Колокольчик», карты
    roads_dirt: list[str] = field(default_factory=list)   # грунтовые дороги: «Колокольчик», карты
    hydro_lines: list[str] = field(default_factory=list)
    hydro_polys: list[str] = field(default_factory=list)
    drainage: list[str] = field(default_factory=list)    # дренажные канавы (линии) — мероприятие III.8 и слой рек
    wetlands: list[str] = field(default_factory=list)    # болота (полигоны) — мероприятие I.8
    recreation: list[str] = field(default_factory=list)   # рекреационные объекты (точки) — тепловая карта на схемах
    lesopat_points: list[str] = field(default_factory=list)  # лесопатология: ПВН (Excel или вектор с категорией)
    # справочник водоёмов и болот ООЗТ (Excel; отметки мероприятий I.8, I.9, I.11, III.7); пусто — refdata/
    water_ref: str | None = None
    # справочник цветения охраняемых растений (Excel; сроки III.4 и синонимы); пусто — refdata/
    flowering_ref: str | None = None
    # слои видов и редких экосистем: [{path, category, group, kind, field}], см. SPECIES_*
    species_layers: list = field(default_factory=list)
    out_dir: str = ""
    report_kv: list = field(default_factory=list)   # кварталы / листы / группы для отчёта; пусто = все
    basemap: bool = True
    min_share_pct: float = 25.0          # мин. доля полигона под участком-основанием (тропы, уклон, вороны, SDM), %
    split_multipart: bool = True         # многочастный объект → отдельные контуры (ID-1, ID-2, …)
    report_formats: list = field(default_factory=lambda: ["docx", "pdf"])  # docx, pdf, doc; report.md — всегда
    fields: dict = field(default_factory=dict)      # kv_vyd: OOZT, KV, VYD, TYPE, NAME; eco: ID, TYPE, NAME, OOZT

    def __post_init__(self):
        # справочники по умолчанию — из папки refdata/ рядом со скриптом
        for key, name in REF_DEFAULTS.items():
            if not getattr(self, key, None):
                d = _REF_DIR / name
                setattr(self, key, str(d) if d.exists() else None)

    def save(self, path: str | Path):
        Path(path).write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")

    @staticmethod
    def load(path: str | Path) -> "Config":
        d = json.loads(Path(path).read_text(encoding="utf-8"))
        known = set(Config.__dataclass_fields__)
        if "shore_min_share_pct" in d and "min_share_pct" not in d:   # прежнее имя параметра
            d["min_share_pct"] = d["shore_min_share_pct"]
        layers = list(d.get("species_layers") or [])
        for key, (cat, grp, kind) in LEGACY_SPECIES_KEYS.items():  # прежние поля конфигурации
            for pth in d.get(key) or []:
                if not any(l.get("path") == pth for l in layers):
                    layers.append(dict(path=pth, category=cat, group=grp, kind=kind, field=""))
        d["species_layers"] = layers
        return Config(**{k: v for k, v in d.items() if k in known})


# прежние поля конфигурации (до версии со списком слоёв видов) → категория, группа, тип данных
LEGACY_SPECIES_KEYS = {
    "invasive_points": ("инвазивные", "не указана", "points"),
    "invasive_polys": ("инвазивные", "не указана", "polygons"),
    "invasive_rasters": ("инвазивные", "не указана", "sdm"),
    "rare_points": ("охраняемые", "не указана", "points"),
    "rare_polys": ("охраняемые", "не указана", "polygons"),
    "rare_rasters": ("охраняемые", "не указана", "sdm"),
    "fauna_points": ("прочие", "птицы", "points"),
}


def unit_labels(cfg: Config) -> dict:
    """Подписи территориальных единиц в таблицах, отчёте и на схемах (зависят от режима)."""
    if cfg.unit_mode == "eco":
        grp = cfg.report_group_field
        return {
            "group": grp if grp else "Лист", "group_pl": f"группам «{grp}»" if grp else "листам",
            "groups_n": "Групп" if grp else "Листов",
            "unit": "ID экосистемы", "units_n": "Экосистем", "units_gen": "экосистем",
            "unit_type": "Тип полигона", "unit_name": "Экосистема",
        }
    return {"group": "Квартал", "group_pl": "кварталам", "groups_n": "Кварталов",
            "unit": "Выдел", "units_n": "Выделов", "units_gen": "выделов",
            "unit_type": "Тип выдела", "unit_name": "Экосистема выдела"}


def display_columns(cfg: Config) -> dict:
    """Внутренние имена колонок (Квартал/Выдел/Тип выдела) → подписи текущего режима."""
    L = unit_labels(cfg)
    return {"Квартал": L["group"], "Выдел": L["unit"], "Тип выдела": L["unit_type"]}


# =====================================================================================
#                 ВХОДНЫЕ СЛОИ И ИХ СВЯЗЬ С МЕРОПРИЯТИЯМИ (для интерфейса и журнала)
# =====================================================================================
# key — поле Config; multi — можно выбрать несколько файлов
LAYER_SPECS = [
    dict(key="dem_path", title="Цифровая модель рельефа", multi=False, kind="ras",
         what="растр высот, 1 канал (tif, asc); уклон считается на сетке 10 м"),
    dict(key="trails", title="Дорожно-тропиночная сеть", multi=True, kind="vec",
         what="линии; OSM highway (path, footway, track…) или полевая съёмка"),
    dict(key="roads_paved", title="Дороги с твёрдым покрытием", multi=True, kind="vec",
         what="линии асфальтированных дорог (для карт и «Колокольчика»); если не задан, "
              "используются линии дорожно-тропиночной сети с признаком асфальта"),
    dict(key="roads_dirt", title="Дороги грунтовые", multi=True, kind="vec",
         what="линии грунтовых дорог (для карт); если не задан, при наличии признака "
              "используются соответствующие линии сети"),
    dict(key="hydro_lines", title="Гидрография — линии", multi=True, kind="vec",
         what="реки, ручьи, каналы"),
    dict(key="hydro_polys", title="Гидрография — полигоны", multi=True, kind="vec",
         what="пруды, озёра, водохранилища; из всех водных полигонов собирается "
              "справочник «ООЗТ — водоём — тип — площадь»"),
    dict(key="drainage", title="Дренажные канавы", multi=True, kind="vec",
         what="линии; добавляются к рекам на картах и включают мероприятие "
              "«Прочистка дренажной системы» (III.8)"),
    dict(key="wetlands", title="Болота (вектор)", multi=True, kind="vec",
         what="полигоны болот/заболоченных участков — для мероприятия "
              "«Частичное удаление деревьев и кустарников на болотах» (I.8)"),
    dict(key="recreation", title="Рекреационные объекты (точки)", multi=True, kind="vec",
         what="точки (площадки, навесы, беседки, инфраструктура отдыха) — тепловая карта "
              "концентрации на схемах кварталов; поле категории объекта, если есть"),
    dict(key="water_ref", title="Справочник водоёмов и болот (Excel)", multi=False, kind="xls",
         what="лист «Битца и Измайлово» / «Все ООЗТ Москвы»: объекты с координатами и отметками "
              "мероприятий I.8, I.9, I.11, III.7 (1 / да / + / ●); по умолчанию — refdata/"
              "Справочник_водоёмов_и_болот_ООЗТ_Москвы.xlsx"),
    dict(key="flowering_ref", title="Справочник цветения охраняемых растений (Excel)", multi=False, kind="xls",
         what="лист «Справочник»: латинские названия (в слое и по Красной книге — синонимы), месяцы "
              "цветения IV–IX (●) — сроки III.4; по умолчанию — refdata/okhranyaemye_rasteniya_tsvetenie.xlsx"),
    dict(key="lesopat_points", title="Лесопатология: ПВН (точки)", multi=True, kind="vec_xls",
         what="Excel с описаниями ПВН (PVN_…xlsx: «№ ПВН», «Рекомендации», «Широта/Долгота») "
              "либо векторный слой (.gpkg, .shp) с полями «Рекомендации» / «recomend*» или «Категория». "
              "Мероприятия V.1–V.6 назначаются по этим точкам"),
]
LAYER_TITLE = {s["key"]: s["title"] for s in LAYER_SPECS}
LAYER_TITLE["species"] = "Слои видов и редких экосистем"

# ---- слои видов и редких экосистем ---------------------------------------------------
# Каждый слой описывается категорией, группой организмов и типом данных. Мероприятия
# выбирают слои по этим признакам (столбцы «Виды: …» таблицы мероприятий).
SPECIES_CATEGORIES = ["инвазивные", "охраняемые", "прочие"]
SPECIES_CATEGORY_HELP = {
    "инвазивные": "чужеродные инвазивные виды",
    "охраняемые": "виды Красной книги и другие охраняемые или редкие виды, редкие экосистемы",
    "прочие": "прочие наблюдения (например, все птицы — для плотности врановых)",
}
SPECIES_GROUPS = ["сосудистые растения", "грибы", "лишайники", "млекопитающие", "птицы",
                  "низшие позвоночные", "насекомые", "редкие экосистемы", "не указана"]
SPECIES_KINDS = {"points": "точки", "polygons": "полигоны", "sdm": "SDM (растр 0–1)"}
RASTER_EXT = (".tif", ".tiff", ".asc", ".img", ".vrt")


def norm_category(v) -> str:
    t = str(v or "").strip().lower().replace("ё", "е")
    if not t:
        return ""
    if t.startswith(("инваз", "чужерод")):
        return "инвазивные"
    if t.startswith(("охран", "редк", "краснокниж")):
        return "охраняемые"
    if t.startswith(("проч", "друг", "фаун", "наблюд")):
        return "прочие"
    return t


def norm_group(v) -> str:
    t = re.sub(r"\s+", " ", str(v or "").strip().lower().replace("ё", "е"))
    return t or "не указана"


def group_matches(layer_group: str, groups) -> bool:
    """Группа слоя входит в список групп мероприятия (сравнение без учёта окончаний)."""
    if not groups:
        return True
    lg = norm_group(layer_group)
    for g in groups:
        g = norm_group(g)
        n = max(4, min(len(g), len(lg)) - 2)
        if lg == g or lg[:n] == g[:n]:
            return True
    return False


def detect_species_kind(path: str) -> str:
    """Тип данных слоя видов: SDM — растр, иначе по геометрии векторного слоя."""
    p = str(path).split("|")[0]
    if p.lower().endswith(RASTER_EXT):
        return "sdm"
    layer = str(path).split("|")[1] if "|" in str(path) else None
    for enc in (None, "latin1"):            # latin1 — если имена полей в другой кодировке
        try:
            info = pyogrio.read_info(p, layer=layer, **({"encoding": enc} if enc else {}))
        except Exception:
            continue
        gt = str(info.get("geometry_type") or "")
        if "Point" in gt:
            return "points"
        if "Polygon" in gt:
            return "polygons"
        return ""
    return ""


def species_layers(cfg) -> list[dict]:
    """Нормализованный список слоев видов из конфигурации."""
    out = []
    for l in getattr(cfg, "species_layers", None) or []:
        if isinstance(l, str):
            l = dict(path=l)
        pth = l.get("path")
        if not pth:
            continue
        out.append(dict(path=pth, category=norm_category(l.get("category")) or "прочие",
                        group=norm_group(l.get("group")),
                        kind=l.get("kind") or detect_species_kind(pth), field=l.get("field") or ""))
    return out


def species_layer_label(l: dict) -> str:
    return (f"{Path(str(l['path']).split('|')[0]).name} ({l['category']}, {l['group']}, "
            f"{SPECIES_KINDS.get(l['kind'], l['kind'] or '?')})")


# ---- правила назначения ---------------------------------------------------------------
REQ, SHARED = "__обязательно__", "__окно__"      # параметр обязателен / берётся из окна скрипта
SPECIES_FILTERS = ("category", "groups", "taxon")
RULES = {
    "type": dict(
        title="по типу полигона", data=[], opt=["hydro"],
        params=dict(types=REQ, water_only=None, water_except=None),
        help="Назначается всем полигонам указанных типов. Если задан фильтр водоёмов, назначаются "
             "только полигоны, пересекающие выбранные водные объекты (по названию из слоя гидрографии)."),
    "trail_density": dict(
        title="плотность троп в полигоне", data=["trails"], opt=[],
        params=dict(types=None, quantile=0.75, abs_m_ha=None),
        help="Длина троп внутри полигона / площадь полигона (м/га); назначается полигонам с плотностью "
             "не ниже квантиля среди полигонов этих типов."),
    "trail_hotspot": dict(
        title="участки повышенной плотности троп", data=["trails"], opt=[],
        params=dict(types=None, window=50.0, quantile=0.75, min_ha=0.05, min_share=SHARED),
        help="Растр плотности троп в скользящем окне; участки не ниже квантиля; назначается полигонам, "
             "где такие участки занимают заданную площадь и долю."),
    "shore": dict(
        title="тропы у водных объектов", data=["trails", "hydro"], opt=[],
        params=dict(types=None, dist=50.0, window=50.0, quantile=0.75, min_ha=0.01, min_share=SHARED),
        help="Участки повышенной плотности троп в полосе заданной ширины вокруг линий и полигонов "
             "гидрографии."),
    "slope": dict(
        title="уклон", data=["dem"], opt=[],
        params=dict(types=None, slope=REQ, min_ha=0.05, min_share=SHARED),
        help="Участки с уклоном больше порога по ЦМР."),
    "species": dict(
        title="виды: находки", data=["species"], opt=[],
        params=dict(types=None, category=None, groups=None, taxon=None, sources=None, kk_categories=None,
                    sdm=0.75, min_ha=0.05, min_share=SHARED),
        help="Полигоны, где есть точки или полигоны выбранных видов; по моделям SDM — отдельная строка "
             "«мониторинг». Можно ограничить категориями Красной книги Москвы (столбец «КК Москвы: категории»)."),
    "species_trails": dict(
        title="виды: тропы у находок", data=["species"], opt=["trails"],
        params=dict(types=None, category=None, groups=None, taxon=None, sources=None, kk_categories=None,
                    dist=50.0, sdm=0.75, informal_only=False,
                    min_ha=0.05, min_share=SHARED, vydel_min_ha=None, vydel_max_ha=None),
        help="Тропы в заданном расстоянии от точек и внутри полигонов выбранных видов; по моделям SDM — "
             "отдельная строка «мониторинг». «Только стихийные тропы» — считать по слою OSM informal=yes; "
             "можно ограничить категориями Красной книги Москвы."),
    "species_density": dict(
        title="виды: плотность находок", data=["species"], opt=[],
        params=dict(types=None, category=None, groups=None, taxon=None, window=500.0, quantile=0.75,
                    min_ha=0.05, min_share=SHARED),
        help="Растр плотности точек выбранных видов в скользящем окне; участки не ниже квантиля."),
    "road_buffer": dict(
        title="буфер дорожной сети", data=["trails"], opt=[],
        params=dict(types=None, dist_min=50.0, dist_max=100.0),
        help="Назначается полигонам, пересекающим полосу между двумя расстояниями от дорожной сети "
             "(например, 50–100 м от дорог)."),
    "drainage": dict(
        title="дренажные канавы", data=["drainage"], opt=[],
        params=dict(types=None, min_m=50.0),
        help="Назначается полигонам, пересекающим слой дренажных канав (линии); основание — "
             "суммарная длина канав внутри полигона."),
    "wetland": dict(
        title="болота (вектор)", data=["wetlands"], opt=[],
        params=dict(types=None),
        help="Назначается полигонам, пересекающим загруженный слой болот и заболоченных участков "
             "(вектор); основание — площадь пересечения."),
    "water_catalog": dict(
        title="водоёмы из справочника", data=["hydro"], opt=[],
        params=dict(types=None, water_only=None),
        help="Из всех загруженных водных полигонов собирается справочник «ООЗТ — водоём — тип — площадь» "
             "(water_catalog.csv в результатах). Мероприятие назначается полигонам, содержащим выбранные "
             "по названию водоёмы (столбец «Водоёмы (только)»); пусто — все водоёмы."),
    "water_ref": dict(
        title="по справочнику водоёмов", data=["water_ref"], opt=["hydro", "wetlands"],
        params=dict(),
        help="Объекты берутся из справочника водоёмов и болот ООЗТ (Excel): строки с отметкой в столбце "
             "кода мероприятия (I.8, I.9, I.11, III.7; отметка — 1, да, + или ●). Точка объекта (широта, "
             "долгота WGS 84) привязывается к контуру: водоём — к полигону гидрографии, болото — к слою болот "
             "(если нет — к гидрографии); сначала полигон, содержащий точку, затем ближайший не дальше 30 м. "
             "ООЗТ и квартал определяются по геометрии (повторяющиеся номера кварталов Терлецкого массива "
             "разделяются пространственно). Участок работ — контур объекта в границах ООЗТ; одна запись на "
             "объект. Флаги «Вкл/выкл» и «Вкл_<ООЗТ>» действуют поверх справочника. Объекты без контура — "
             "в water_ref_unmatched.csv."),
    "struct_class": dict(
        title="по структурному классу", data=[], opt=[],
        params=dict(struct_classes=STRUCT_MOW_DEFAULT),
        help="Назначается полигонам заданных структурных классов (столбец «Структурные классы»; по умолчанию — "
             "суходольные, мезофитные и мезофитные антропогенные луга; гидроморфные луга не косят). Класс берётся "
             "из поля «Структурный класс», выбранного в окне скрипта (если не задано — из полей типа и названия); "
             "числовой код class_value (например, 5222) понимается по трём последним цифрам (222). Участок работ — "
             "полигон экосистемы в границах ООЗТ."),
    "rare_trail_vydel": dict(
        title="виды КК: стихийные тропы в выделе", data=["species"], opt=["trails", "flowering"],
        params=dict(category=None, groups=None, taxon=None, sources=None, kk_categories=(1, 2, 3),
                    abs_m_ha=100.0, informal_only=True, vydel_min_ha=None, vydel_max_ha=None,
                    season_flower=False),
        help="Выдел любого типа назначается, если плотность стихийных троп в нём (длина стихийных троп — OSM "
             "informal=yes или свой слой — внутри выдела / площадь, м/га) больше порога (100 м/га) и в выделе есть находка вида Красной "
             "книги Москвы заданных категорий (1–3; с учётом синонимов). Ограничение по площади выдела: «Мин. "
             "площадь выдела» — больше, «Макс. площадь выдела» — не больше. «Сроки по цветению» = да — "
             "сезонность записи берётся из справочника цветения (объединение месяцев видов выдела). Участок "
             "работ — весь выдел."),
    "points_category": dict(
        title="точки по категории", data=["lesopat"], opt=[],
        params=dict(types=None, cat_value=REQ),
        help="Точечный слой лесопатологии (ПВН): файл Excel (PVN_…xlsx с колонками «№ ПВН», «Рекомендации», "
             "«Широта (WGS84)») или вектор с полем категории. Назначается полигонам с точками заданной "
             "категории (столбец «Точки: категория»: санитарная рубка, рубки ухода, надзор, посадка, "
             "эталон, биотехния)."),
    "kolokolchik": dict(
        title="«Колокольчик»: высаживание КК видов", data=[], opt=["roads"],
        params=dict(types=REQ, ring=(500.0, 1000.0), road_dist=100.0),
        help="Высаживание видов Красной книги в типичных местообитаниях: зона 500–1000 м от границы "
             "парка (внутрь), полоса 100 м от дорог с твёрдым покрытием и полоса вдоль границ выделов. "
             "Список видов — по типу экосистемы: широколиственный лес — первоцветы и колокольчик "
             "широколиственный; луга естественные — колокольчики персиколистный и раскидистый; "
             "понижения и западины — купальница и калужница."),
}
RULE_TITLES = [v["title"] for v in RULES.values()]

# столбцы таблицы мероприятий, задающие правило (ключ, заголовок, начало заголовка для поиска, пояснение)
SPEC_COLUMNS = [
    ("rule", "Правило назначения", "правило", "одно из правил (список); пусто — правило по тексту «Метода выявления»"),
    ("types", "Типы полигонов", "типы полиг",
     "типы полигонов слоя территории через запятую: лесные, луговые, водные, околоводные, болотные…; "
     "«любые» или пусто — все полигоны"),
    ("category", "Виды: категория", "виды: катег", "инвазивные, охраняемые, прочие (через запятую); «любые» или пусто — все"),
    ("groups", "Виды: группы организмов", "виды: групп",
     "сосудистые растения, грибы, лишайники, млекопитающие, птицы, низшие позвоночные, насекомые, "
     "редкие экосистемы…; «любые» или пусто — все"),
    ("taxon", "Виды: таксоны (лат.)", "виды: такс",
     "латинские названия видов или родов через запятую (например, Corvus); пусто — все виды"),
    ("kk_categories", "КК Москвы: категории", "кк моск",
     "категории Красной книги Москвы (0–5, «О») через запятую; например «1, 2, 3» — под угрозой исчезновения, "
     "редкие и уязвимые; пусто — без учёта Красной книги"),
    ("informal_only", "Только стихийные тропы", "только стихийные троп",
     "«да» — тропы у находок считаются только по слою стихийных троп (скачивается из OSM informal=yes "
     "или выбирается свой слой в окне); «нет» или пусто — по всей дорожно-тропиночной сети"),
    ("sources", "Виды: данные", "виды: данн", "точки, полигоны, SDM (через запятую); «любые» или пусто — все"),
    ("dist", "Расстояние, м", "расстоян", "ширина полосы у воды или радиус вокруг точек видов"),
    ("window", "Окно плотности, м", "окно", "радиус скользящего окна для растра плотности"),
    ("quantile", "Квантиль", "квантил", "порог плотности; 0,75 — верхний квартиль"),
    ("slope", "Уклон более, %", "уклон", "порог уклона, %"),
    ("sdm", "Порог SDM (0–1)", "порог sdm", "вероятность модели, начиная с которой участок учитывается"),
    ("min_ha", "Мин. площадь участка, га", "мин. площадь участ", "минимальная площадь участка внутри полигона"),
    ("min_share", "Мин. доля полигона, %", "мин. дол",
     "минимальная доля площади полигона под участками; пусто — общий порог из окна скрипта"),
    ("vydel_min_ha", "Мин. площадь выдела, га", "мин. площадь выдела",
     "правило назначается только выделам больше этой площади (например, III.3 — больше 5 га)"),
    ("vydel_max_ha", "Макс. площадь выдела, га", "макс. площадь выдела",
     "правило назначается только выделам не больше этой площади (например, III.4 — не больше 5 га)"),
    ("abs_m_ha", "Плотность троп, м/га (порог)", "плотность троп",
     "абсолютный порог плотности троп в полигоне, м/га; пусто — квантиль"),
    ("water_only", "Водоёмы (только)", "водоемы (только)",
     "названия водных объектов через запятую; назначаются только полигоны, пересекающие их "
     "(совпадение по части названия, например «Лебедянский»); требует слой гидрографии"),
    ("min_m", "Мин. длина в полигоне, м", "мин. длина",
     "минимальная суммарная длина линий слоя внутри полигона, м; для «дренажные канавы»"),
    ("cat_value", "Точки: категория", "точки: катег",
     "категория точек лесопатологии для правила «точки по категории»: санитарная рубка, рубки ухода, "
     "надзор, посадка, эталон, биотехния"),
    ("ring", "Зона от границы, м", "зона от границы",
     "кольцо от границы парка в метрах «мин-макс» (например «500-1000»); для «Колокольчика»"),
    ("road_dist", "Полоса от дорог, м", "полоса от дорог",
     "ширина полосы от дорог с твёрдым покрытием, м; для «Колокольчика»"),
    ("water_except", "Водоёмы (кроме)", "водоемы (кроме)",
     "названия водных объектов через запятую; эти объекты исключаются (например, «отстойник»); "
     "требует слой гидрографии"),
    ("dist_min", "Буфер от дорог, мин., м", "буфер от дорог, мин",
     "внутренний радиус полосы от дорожной сети (0 — от самих дорог)"),
    ("dist_max", "Буфер от дорог, макс., м", "буфер от дорог, макс",
     "внешний радиус полосы от дорожной сети"),
    ("struct_classes", "Структурные классы", "структурные класс",
     "структурные классы экосистем через запятую для правила «по структурному классу» (например, «Суходольные "
     "луга, Мезофитные луга, Мезофитные луга (антропогенные)»); сравнение по названию целиком без учёта "
     "регистра; коды class_value — по таблице кодов"),
    ("season_flower", "Сроки по цветению", "сроки по цвет",
     "«да» — сезонность записи по справочнику цветения охраняемых растений (объединение месяцев цветения "
     "видов КК в выделе; плаун годичный — VII–VIII); «нет» или пусто — сезонность из столбца «Сезонность»"),
]
SPEC_TITLE = {k: t for k, t, _, _ in SPEC_COLUMNS}
DESC_COLUMN = "Процедура назначения (как работает скрипт)"

TYPE_WORDS = [("околовод", "околовод"), ("болот", "болот"), ("водн", "водн"), ("вод", "водн"),
              ("лес", "лес"), ("луг", "луг")]


def _type_word_match(kw: str, t_words: list[str]) -> bool:
    """Слово ключа совпадает со словом типа: по стемам TYPE_WORDS (луг ~ луга, луговые)
    либо по общему началу не короче 4 символов (естественные ~ естественных)."""
    stem = next((s for s, _ in TYPE_WORDS if s in kw), None)
    # «нелесные» не должны попадать в «лесные»: слово с приставкой «не» не матчится по стему
    if stem is not None and any(stem in w and not (w.startswith("не") and stem in w[2:])
                                for w in t_words):
        return True
    n = max(4, len(kw))
    return any(w.startswith(kw[:n]) or (len(w) >= 4 and kw.startswith(w)) for w in t_words)


def type_matches(type_value, keys, name_value=None) -> bool:
    """Тип полигона относится к одному из ключей. «водн» не совпадает с «околоводные»:
    иначе мероприятия только для водных экосистем (аэрация) попадают и на околоводные.
    Ключ из нескольких слов требует совпадения каждого слова; слово-стем (лес, луг,
    болот, водн) ищется в поле «Тип», слова-уточнения («широколиственные»,
    «естественные») — в «Тип» или названии полигона (name_value).
    Луга: «луговые естественные» — луговые типы, кроме газонов (503) и бурьяна (222);
    «луговые антропогенные» — газоны и бурьян (слово «луг» в типе не обязательно)."""
    t = str(type_value or "").lower()
    t_water = t.replace("околовод", "")          # «водн» ищем вне слова «околоводные»
    tw = [w for w in re.split(r"[^а-яa-z0-9]+", t) if w]
    tw_w = [w for w in re.split(r"[^а-яa-z0-9]+", t_water) if w]
    ta = tw + [w for w in re.split(r"[^а-яa-z0-9]+", str(name_value or "").lower())
               if w and w not in tw]
    is_meadow = any("луг" in w for w in tw)
    for k in keys:
        kl = k.lower().replace("ё", "е")
        if "луг" in kl and "естественн" in kl:
            if is_meadow and not _anthro_meadow(t + " " + str(name_value or "")):
                return True
            continue
        if "луг" in kl and "антропоген" in kl:
            if _anthro_meadow(t + " " + str(name_value or "")):
                return True
            continue
        kws = [w for w in re.split(r"[^а-яa-z0-9]+", k) if w]
        if not kws:
            continue
        ok = True
        for kw in kws:
            stem = next((s for s, _ in TYPE_WORDS if s in kw), None)
            src = tw_w if stem == "водн" else (tw if stem else ta)
            if not _type_word_match(kw, src):
                ok = False
                break
        if ok:
            return True
    return False


def classify_measure(measure: dict) -> dict:
    """Метод выявления → правило. Возвращает {'rule', 'params', 'reason'} (rule=None — не рассчитывается).
    Правило определяется по тексту метода выявления; название мероприятия используется только там,
    где метод сам по себе неоднозначен (общий текст «точечные / полигональные / растровые данные»,
    поиск троп без указания объекта)."""
    m = re.sub(r"\s+", " ", (measure.get("method") or "").lower())
    nm = (measure.get("name") or "").lower()
    if not m:
        return dict(rule=None, params={}, reason="метод выявления не задан")
    if "нет данных" in m:
        return dict(rule=None, params={}, reason="по таблице: нет данных")
    water_words = ("водн", "водо", "берег", "акватор", "реки", "ручь", "пруд")
    species_data = any(w in m for w in ("точечн", "полигональн", "растров"))
    trails = "троп" in m or "дорог" in m
    if "corvus" in m or "врано" in m or ("птиц" in m and "плотн" in m):
        return dict(rule="corvus", params={}, reason="")
    if "уклон" in m or "цмр" in m:
        mt = re.search(r"(\d+(?:[.,]\d+)?)\s*%", m)
        thr = float(mt.group(1).replace(",", ".")) if mt else 7.0
        keys = TYPE_MEADOW if "лугов" in m else None
        if keys is None and "люб" in str(measure.get("eco", "")).lower():
            thr = float(mt.group(1).replace(",", ".")) if mt else 7.0
        return dict(rule="slope", params=dict(threshold_pct=thr, type_keys=keys), reason="")
    if "точек видов" in m or ("троп" in m and "точ" in m and "50" in m and "вид" in m):
        return dict(rule="rare", params={}, reason="")
    if ("инвази" in m or "чужерод" in m or "чежерод" in m) or (
            species_data and not trails and any(w in nm for w in ("инвази", "чужерод", "чежерод"))):
        return dict(rule="invasive", params={}, reason="")
    if trails and ("периметр" in m or any(w in m for w in water_words) or "берег" in nm):
        return dict(rule="shore", params={}, reason="")
    if "скольз" in m and "троп" in m:
        return dict(rule="trail_hotspot", params={}, reason="")
    if "плотност" in m and trails:
        return dict(rule="trail_density", params={}, reason="")
    if "лугов" in m:
        return dict(rule="type", params=dict(keys=TYPE_MEADOW, label="луговые"), reason="")
    if "околовод" in m or "болот" in m:
        return dict(rule="type", params=dict(keys=TYPE_WATER + TYPE_NEARWATER, label="водные и околоводные"),
                    reason="")
    if any(w in m for w in ("водн", "водо", "акватор")):
        return dict(rule="type", params=dict(keys=TYPE_WATER, label="водные"), reason="")
    if "лесн" in m:
        return dict(rule="type", params=dict(keys=TYPE_FOREST, label="лесные"), reason="")
    if "ворон" in nm and "плотн" in m:
        return dict(rule="corvus", params={}, reason="")
    return dict(rule=None, params={}, reason="метод не распознан")


_MIN_SHARE = [25.0]     # текущий порог доли полигона (обновляется из конфигурации)


def _fmt(x, nd=2) -> str:
    """Число для описаний: 0,75; 50; 0,05."""
    if x is None:
        return "—"
    x = float(x)
    s = f"{x:.{nd}f}".rstrip("0").rstrip(".") if abs(x - round(x)) > 1e-9 else str(int(round(x)))
    return s.replace(".", ",")


def _is_any(t: str) -> bool:
    t = t.strip().lower()
    return t in ("", "-", "—", "любые", "любой", "любая", "все", "всё", "any")


def _split_list(t) -> list[str]:
    return [x.strip() for x in re.split(r"[,;/]|\s+и\s+", str(t)) if x.strip()]


def parse_types(text):
    """«лесные, луговые» → (('лес', 'луг'), 'лесные, луговые'); «любые» → None.
    Уточнённые типы из нескольких слов сохраняются целиком: «луговые естественные» —
    отдельный ключ; луговые типы без газонов (503) и бурьяна (222) — естественные
    (см. type_matches), газоны и бурьян — антропогенные."""
    t = str(text or "").strip()
    if _is_any(t):
        return None
    keys = []
    for tok in _split_list(t.lower().replace("ё", "е")):
        k = next((key for stem, key in TYPE_WORDS if stem in tok), None)
        if k is None:                      # другой тип: сравнение по началу слова
            k = tok[:max(4, len(tok) - 2)]
        elif len(tok.split()) > 1:         # «луговые естественные», «леса хвойные» — как есть
            k = tok
        if k not in keys:
            keys.append(k)
    return (tuple(keys), t) if keys else None


def _parse_value(key: str, raw):
    """Значение ячейки → параметр правила. Возвращает (значение, ошибка)."""
    if raw is None or (isinstance(raw, float) and np.isnan(raw)) or str(raw).strip() == "":
        return None, None
    t = str(raw).strip()
    if key == "types":
        return parse_types(t), None
    if key in ("category", "groups", "sources"):
        if _is_any(t):
            return None, None
        items = _split_list(t)
        if key == "category":
            vals = [norm_category(x) for x in items]
            bad = [x for x, v in zip(items, vals) if v not in SPECIES_CATEGORIES]
            return (vals, None) if not bad else (None, f"неизвестная категория: {', '.join(bad)}")
        if key == "groups":
            return [norm_group(x) for x in items], None
        vals = []
        for x in items:
            xl = x.lower()
            v = ("points" if xl.startswith("точ") else "polygons" if xl.startswith("полиг") else
                 "sdm" if xl.startswith(("sdm", "растр", "модел")) else None)
            if v is None:
                return None, f"неизвестный тип данных: {x}"
            vals.append(v)
        return vals, None
    if key == "taxon":
        return (None, None) if _is_any(t) else (_split_list(t), None)
    if key in ("water_only", "water_except"):
        return (None, None) if _is_any(t) else (_split_list(t), None)
    if key == "kk_categories":                    # категории Красной книги Москвы: 0–5, «О»
        if _is_any(t):
            return None, None
        vals = []
        for x in _split_list(t):
            if x.upper() in ("О", "O"):
                vals.append("О")
                continue
            try:
                v = int(float(x))
            except ValueError:
                return None, f"неизвестная категория Красной книги: {x}"
            if not 0 <= v <= 5:
                return None, f"категория Красной книги должна быть от 0 до 5: {x}"
            vals.append(v)
        return vals, None
    if key == "struct_classes":                   # структурные классы — названия через запятую / «;»
        if _is_any(t):
            return None, None
        return [x.strip() for x in re.split(r"[,;]", t) if x.strip()], None
    if key in ("informal_only", "season_flower"):  # «да» / «нет»
        v = t.lower().replace("ё", "е")
        if v in ("да", "1", "истина", "true", "вкл", "только стихийные", "стихийные"):
            return True, None
        if v in ("нет", "0", "ложь", "false", "выкл", "все", "любые", "-"):
            return False, None
        return None, f"да или нет: «{t}»"
    if key == "cat_value":                        # категория точек — текст
        return t, None
    if key == "ring":                             # «500-1000» → (500, 1000)
        m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[-–—]\s*(\d+(?:\.\d+)?)\s*$", t)
        if not m:
            return None, "зона от границы: «мин-макс», например «500-1000»"
        lo, hi = float(m.group(1)), float(m.group(2))
        if hi <= lo:
            return None, "максимум зоны должен быть больше минимума"
        return (lo, hi), None
    try:
        v = float(t.replace(",", ".").replace("%", "").strip())
    except ValueError:
        return None, f"не число: «{t}»"
    if key == "quantile" and not 0 < v < 1:
        return None, "квантиль должен быть между 0 и 1"
    if key == "sdm" and not 0 <= v <= 1:
        return None, "порог SDM должен быть от 0 до 1"
    if v < 0:
        return None, "отрицательное значение"
    return v, None


def rule_from_title(t) -> str | None:
    def _norm(x):
        return re.sub(r"\s+", " ", str(x or "").strip().lower().replace("ё", "е"))
    s = _norm(t)
    if not s:
        return None
    for k, v in RULES.items():
        if s == _norm(v["title"]) or s == k:
            return k
    for k, v in RULES.items():                       # без учёта двоеточий, кавычек и регистра
        if re.sub(r"[^а-яa-z ]", "", s) == re.sub(r"[^а-яa-z ]", "", _norm(v["title"])):
            return k
    return "?"


TYPE_KEY_WORD = {"лес": "лесные", "луг": "луговые", "водн": "водные", "околовод": "околоводные", "болот": "болотные"}


def types_label(keys) -> str:
    return ", ".join(TYPE_KEY_WORD.get(k, k) for k in keys)


def classify_method_text(measure: dict) -> dict:
    """Прежний способ: правило по ключевым словам «Метода выявления» (если столбец «Правило назначения» пуст)."""
    c = classify_measure(measure)
    r, pr = c["rule"], c["params"]
    if r is None:
        return dict(rule=None, p={}, reason=c["reason"])
    forest = (TYPE_FOREST, types_label(TYPE_FOREST))
    if r == "corvus":
        return dict(rule="species_density", p=dict(types=forest, taxon=["Corvus"], window=CORVUS_WINDOW_M))
    if r == "slope":
        return dict(rule="slope", p=dict(types=(TYPE_MEADOW, types_label(TYPE_MEADOW)) if pr.get("type_keys") else None,
                                         slope=pr["threshold_pct"]))
    if r == "invasive":
        return dict(rule="species", p=dict(category=["инвазивные"], sdm=SDM_THRESHOLD_INVASIVE))
    if r == "rare":
        return dict(rule="species_trails", p=dict(category=["охраняемые"], dist=RARE_BUFFER_M,
                                                  sdm=SDM_THRESHOLD_RARE))
    if r == "shore":
        return dict(rule="shore", p=dict(dist=SHORE_BUFFER_M, window=TRAIL_WINDOW_M, min_ha=MIN_PATCH_HA / 5))
    if r in ("trail_hotspot", "trail_density"):
        return dict(rule=r, p=dict(types=forest))
    return dict(rule="type", p=dict(types=(tuple(pr["keys"]), types_label(pr["keys"]))))


def measure_spec(measure: dict) -> dict:
    """Полное описание правила мероприятия: rule, p (все параметры), src (откуда взят параметр),
    origin («таблица» / «текст метода»), errors, reason (почему не рассчитывается)."""
    raw = measure.get("spec") or {}
    errors = []
    rule_raw = raw.get("rule")
    if rule_raw is not None and str(rule_raw).strip() and str(rule_raw).strip().lower() not in ("nan",):
        rule = rule_from_title(rule_raw)
        if rule == "?":
            return dict(rule=None, p={}, src={}, origin="таблица", errors=[],
                        reason=f"правило «{rule_raw}» не распознано; допустимы: " + "; ".join(RULE_TITLES))
        origin, given = "таблица", {}
        for k, v in raw.items():
            if k == "rule" or k not in RULES[rule]["params"]:
                continue
            val, err = _parse_value(k, v)
            if err:
                errors.append(f"{SPEC_TITLE[k]}: {err}")
            elif val is not None:
                given[k] = val
    else:
        lc = classify_method_text(measure)
        if lc["rule"] is None:
            return dict(rule=None, p={}, src={}, origin="текст метода", errors=[], reason=lc["reason"])
        rule, given, origin = lc["rule"], lc["p"], "текст метода"
    p, src = {}, {}
    for k, d in RULES[rule]["params"].items():
        if k in given:
            p[k], src[k] = given[k], origin
        elif d == REQ:
            p[k], src[k] = None, "не задано"
            errors.append(f"не задан обязательный параметр «{SPEC_TITLE[k]}»")
        elif d == SHARED:
            p[k], src[k] = None, "окно скрипта"
        else:
            p[k], src[k] = d, "по умолчанию"
    if rule == "water_ref":
        # объекты задаёт справочник водоёмов: прежние фильтры (типы, водоёмы только/кроме) не применяются
        for k in ("types", "water_only", "water_except"):
            p.pop(k, None)
            src.pop(k, None)
    if rule == "wetland":
        # болота задаёт сам вектор: мероприятие назначается любому полигону,
        # пересекающему слой болот — столбец «Типы полигонов» не ограничивает
        p.pop("types", None)
        src.pop("types", None)
    return dict(rule=rule, p=p, src=src, origin=origin, errors=errors, reason="")


def spec_share(p: dict) -> float:
    v = p.get("min_share")
    return float(_MIN_SHARE[0] if v is None else v)


def _types_txt(p, any_txt="полигоны любого типа") -> str:
    t = p.get("types")
    return any_txt if not t else f"полигоны типа «{t[1]}»"


def _q_txt(q) -> str:
    return "верхний квартиль (0,75)" if abs(float(q) - 0.75) < 1e-9 else f"квантиль {_fmt(q)}"


def _sel_txt(p) -> str:
    """Какие слои видов выбирает мероприятие."""
    cat = ", ".join(p.get("category") or []) or "любые"
    grp = ", ".join(p.get("groups") or []) or "любые"
    parts = [f"категория — {cat}", f"группы — {grp}"]
    if p.get("taxon"):
        parts.append("таксоны — " + ", ".join(p["taxon"]))
    if p.get("kk_categories"):
        parts.append("КК Москвы — категории " + ", ".join(str(c) for c in p["kk_categories"]))
    if "sources" in p:
        parts.append("данные — " + (", ".join(SPECIES_KINDS[s] for s in p["sources"]) if p.get("sources")
                                    else "точки, полигоны, SDM"))
    return "; ".join(parts)


def _patch_txt(p) -> str:
    sh = spec_share(p)
    return (f"участок ≥ {_fmt(p['min_ha'])} га" + (f" и ≥ {_fmt(sh)}% площади полигона" if sh > 0 else ""))


def _uses(p, kind) -> bool:
    return not p.get("sources") or kind in p["sources"]


def describe_spec(spec: dict, full: bool = False) -> str:
    """Описание правила: короткое (журнал, отчёт, окно) или полное (столбец таблицы, окно)."""
    r, p = spec.get("rule"), spec.get("p", {})
    if r is None:
        return "не рассчитывается: " + spec.get("reason", "")
    if spec.get("errors"):
        return "не рассчитывается — ошибка в параметрах: " + "; ".join(spec["errors"])
    if not full:
        if r == "type":
            if p.get("water_only") or p.get("water_except"):
                wf = (f"; водоёмы: только {', '.join(p['water_only'])}" if p.get("water_only") else "") \
                     + (f"; кроме {', '.join(p['water_except'])}" if p.get("water_except") else "")
                return f"все {_types_txt(p)}{wf}"
            return f"все {_types_txt(p)}"
        if r == "road_buffer":
            return (f"пересечение с полосой {_fmt(p['dist_min'])}–{_fmt(p['dist_max'])} м "
                    f"от дорожной сети; {_types_txt(p)}")
        if r == "trail_density":
            if p.get("abs_m_ha"):
                return (f"плотность троп на полигон ≥ {_fmt(p['abs_m_ha'])} м/га (абсолютный порог); "
                        f"{_types_txt(p)}")
            return f"плотность троп на полигон (м/га) ≥ {_q_txt(p['quantile'])}; {_types_txt(p)}"
        if r == "trail_hotspot":
            return (f"участки повышенной плотности троп (окно {_fmt(p['window'])} м, {_q_txt(p['quantile'])}); "
                    f"{_types_txt(p)}; {_patch_txt(p)}")
        if r == "shore":
            return (f"участки повышенной плотности троп (окно {_fmt(p['window'])} м, {_q_txt(p['quantile'])}) "
                    f"в {_fmt(p['dist'])} м от водных объектов; {_types_txt(p)}; {_patch_txt(p)}")
        if r == "slope":
            return f"уклон по ЦМР более {_fmt(p['slope'])}%; {_types_txt(p)}; {_patch_txt(p)}"
        sel = _sel_txt(p)
        sdm = (f"; SDM ≥ {_fmt(p['sdm'])} — отдельная строка «мониторинг» ({_patch_txt(p)})"
               if "sdm" in p and _uses(p, "sdm") else "")
        if r == "species":
            return f"находки видов ({sel}): полигоны с точками или полигонами находок{sdm}; {_types_txt(p)}"
        if r == "species_trails":
            at = _vydel_area_txt(p, "; только выделы ")
            inf = "; только стихийные тропы OSM (informal=yes)" if p.get("informal_only") else ""
            return (f"тропы в {_fmt(p['dist'])} м от точек и в полигонах видов ({sel}){sdm}{at}{inf}; "
                    f"{_types_txt(p)}")
        if r == "species_density":
            return (f"плотность точек видов ({sel}) в окне {_fmt(p['window'])} м ≥ {_q_txt(p['quantile'])}; "
                    f"{_types_txt(p)}; {_patch_txt(p)}")
        if r == "drainage":
            return f"пересечение со слоем дренажных канав (≥ {_fmt(p['min_m'])} м внутри); {_types_txt(p)}"
        if r == "wetland":
            return f"пересечение со слоем болот (вектор); {_types_txt(p)}"
        if r == "water_catalog":
            wf = f"; только {', '.join(p['water_only'])}" if p.get("water_only") else ""
            return f"водоёмы из справочника{wf}; {_types_txt(p)}"
        if r == "points_category":
            return f"точки лесопатологии (ПВН) категории «{p.get('cat_value')}»; {_types_txt(p)}"
        if r == "water_ref":
            return "объекты, отмеченные в справочнике водоёмов и болот ООЗТ (столбец кода мероприятия)"
        if r == "struct_class":
            return "полигоны структурных классов: " + ", ".join(p.get("struct_classes") or STRUCT_MOW_DEFAULT)
        if r == "rare_trail_vydel":
            kk = ", ".join(str(c) for c in (p.get("kk_categories") or []))
            return (f"выделы с плотностью стихийных троп > {_fmt(p.get('abs_m_ha'))} м/га и находками видов КК "
                    f"Москвы (кат. {kk}){_vydel_area_txt(p, '; выделы ')}"
                    + ("; сроки — по справочнику цветения" if p.get("season_flower") else ""))
        if r == "kolokolchik":
            lo, hi = p.get("ring") or (500.0, 1000.0)
            return (f"зоны: кольцо {lo:.0f}–{hi:.0f} м от границы, полоса {_fmt(p['road_dist'])} м от дорог, "
                    f"граница выдела; {_types_txt(p)}")
    # ---- полное описание
    trails = ("слой троп (если в слое есть поле HIGHWAY или fclass, учитываются типы "
              + ", ".join(sorted(TRAIL_HIGHWAY_TYPES)) + ")")
    sh = spec_share(p)
    share = (f"не меньше {_fmt(p.get('min_ha'))} га и не меньше {_fmt(sh)}% площади полигона" if sh > 0
             else f"не меньше {_fmt(p.get('min_ha'))} га (порог доли полигона отключён)")
    tgt_each = "каждого полигона любого типа" if not p.get("types") else f"каждого полигона типа «{p['types'][1]}»"
    q = _q_txt(p.get("quantile", 0.75))
    hot = (f"По слою троп строится растр плотности: сетка {_fmt(CELL)} м, круговое скользящее окно радиусом "
           f"{_fmt(p.get('window'))} м, значение — длина троп в окне на гектар (м/га). По ячейкам внутри ООЗТ с "
           f"ненулевой плотностью вычисляется {q}; ячейки не ниже него объединяются в участки повышенной "
           f"плотности троп.")
    fin = (f"Мероприятие назначается, если такие участки занимают в полигоне {share}. В основании указываются "
           f"площадь участков и доля полигона; участки работ записываются в слой measure_zones.")
    sdm = (f"Модели SDM: выделяются участки с вероятностью не ниже {_fmt(p.get('sdm'))}; полигону назначается "
           f"отдельная строка «… — мониторинг (SDM ≥ {_fmt(p.get('sdm'))})», если участки одного вида занимают в нём "
           f"не меньше {_fmt(p.get('min_ha'))} га, а все вместе — "
           + (f"не меньше {_fmt(sh)}% площади полигона." if sh > 0 else "любую долю полигона.")
           ) if "sdm" in p and _uses(p, "sdm") else ""
    sp = (f"Слои видов выбираются по признакам, заданным при загрузке: {_sel_txt(p)}. Название вида берётся из "
          "поля с латинскими названиями (если его нет — из имени файла; у растров SDM — всегда из имени файла).")
    if r == "type":
        wf = ""
        if p.get("water_only") or p.get("water_except"):
            wf = (" Назначаются только полигоны, пересекающие водные объекты, выбранные по названию из "
                  "полигональной гидрографии: " + (f"только {', '.join(p['water_only'])}" if p.get("water_only") else "все")
                  + (f", кроме {', '.join(p['water_except'])}" if p.get("water_except") else "")
                  + ". Совпадение по части названия.")
        return (f"Назначается всем {('полигонам типа «' + p['types'][1] + '»') if p.get('types') else 'полигонам'} "
                "слоя территории. Тип определяется по началу слова (лес…, луг…, болот…); «водные» не включают "
                f"«околоводные».{wf} Других данных не требуется.")
    if r == "road_buffer":
        return (f"Данные: {trails}. Вокруг дорожной сети строится полоса между расстояниями "
                f"{_fmt(p['dist_min'])} и {_fmt(p['dist_max'])} м (сами дороги и акватория не входят). "
                f"Мероприятие назначается каждому полигону, пересекающему полосу. В основании указывается ширина "
                f"полосы; сама полоса показывается на схеме.")
    if r == "trail_density":
        if p.get("abs_m_ha"):
            return (f"Данные: {trails}. Для {tgt_each} считается плотность троп — длина троп внутри полигона, "
                    f"делённая на его площадь (м/га). Мероприятие назначается полигонам с плотностью не ниже "
                    f"абсолютного порога {_fmt(p['abs_m_ha'])} м/га (шкала рекреационной нагрузки: 50 — повышенная, "
                    "100–150 — критическая). Полигон оценивается целиком, порог доли не применяется.")
        return (f"Данные: {trails}. Для {tgt_each} считается плотность троп — длина троп внутри полигона, "
                f"делённая на его площадь (м/га). Среди этих полигонов с ненулевой плотностью вычисляется {q}. "
                "Мероприятие назначается полигонам с плотностью не ниже порога; полигон оценивается целиком, "
                "порог доли не применяется. На схеме красным показаны тропы внутри зон повышенной плотности.")
    if r == "trail_hotspot":
        return f"Данные: {trails}. 1) {hot} 2) Для {tgt_each} считается площадь таких участков внутри него. {fin}"
    if r == "shore":
        return (f"Данные: {trails}; гидрография — линии и полигоны. 1) {hot} 2) Водные объекты объединяются, "
                f"вокруг них строится полоса шириной {_fmt(p['dist'])} м; сама акватория исключается. 3) Участки "
                f"повышенной плотности троп обрезаются этой полосой. 4) Для {tgt_each} считается площадь таких "
                f"участков. {fin}")
    if r == "slope":
        return (f"Данные: цифровая модель рельефа, приведённая к сетке {_fmt(CELL)} м. 1) Уклон считается в "
                "процентах после лёгкого сглаживания (окно 3 × 3 ячейки). 2) Отбираются ячейки внутри ООЗТ с "
                f"уклоном более {_fmt(p['slope'])}%; одиночные ячейки убираются (морфологическое открытие 2 × 2), "
                f"остальные объединяются в участки. 3) Для {tgt_each} считается площадь участков. {fin}")
    if r == "species":
        parts = [sp]
        if _uses(p, "points"):
            parts.append("Точки: мероприятие назначается каждому полигону, внутри которого есть хотя бы одна "
                         "точка, без порогов площади; в основании — виды, число точек и средние координаты "
                         "WGS 84 по каждому виду.")
        if _uses(p, "polygons"):
            parts.append("Полигоны находок: мероприятие назначается каждому полигону, который с ними "
                         "пересекается; указывается площадь пересечения по видам.")
        if sdm:
            parts.append(sdm)
        parts.append("Учитываются " + _types_txt(p) + ".")
        return " ".join(parts)
    if r == "species_trails":
        z = []
        if _uses(p, "points"):
            z.append(f"круг радиусом {_fmt(p['dist'])} м вокруг каждой точки")
        if _uses(p, "polygons"):
            z.append("сами полигоны находок")
        parts = [sp]
        if z:
            at = _vydel_area_txt(p, " Мероприятие назначается только полигонам площадью ")
            at = at + "." if at else ""
            parts.append(f"1) Зоны видов — {' и '.join(z)}. 2) Тропы ({trails}) обрезаются по этим зонам. "
                         "3) Мероприятие назначается каждому полигону, в котором есть хотя бы фрагмент таких "
                         "троп, без порогов длины и доли; в основании — виды, длина троп в зоне (м) и координаты "
                         "до трёх самых длинных фрагментов WGS 84." + at + " Если слой "
                         "троп не загружен, мероприятие назначается полигонам, пересекающим сами зоны.")
        if sdm:
            parts.append(sdm)
        parts.append("Учитываются " + _types_txt(p) + ".")
        return " ".join(parts)
    if r == "species_density":
        return (f"{sp} Используются только точки. 1) Строится растр плотности: сетка {_fmt(CELL)} м, круговое окно "
                f"радиусом {_fmt(p['window'])} м, значение — число точек на гектар. 2) По ячейкам внутри ООЗТ с "
                f"ненулевой плотностью вычисляется {q}; ячейки не ниже него объединяются в участки. 3) Для "
                f"{tgt_each} считается площадь участков. {fin}")
    if r == "drainage":
        return (f"Данные: слой дренажных канав (линии). Мероприятие назначается каждому полигону, который "
                f"канавы пересекают, если суммарная длина канав внутри полигона не меньше {_fmt(p['min_m'])} м; "
                f"канавы также добавляются к рекам на схемах и в ГИС-проекте. Учитываются {_types_txt(p)}.")
    if r == "wetland":
        return (f"Данные: слой болот и заболоченных участков (полигоны). Мероприятие назначается каждому "
                f"полигону, пересекающему болота; в основании — площадь пересечения. "
                f"Учитываются {_types_txt(p)}.")
    if r == "water_catalog":
        wf = (f" Из справочника берутся только водоёмы «{', '.join(p['water_only'])}» (совпадение по части названия)."
              if p.get("water_only") else " Берутся все водоёмы справочника.")
        return (f"Данные: полигональная гидрография. Из всех загруженных водных полигонов собирается "
                f"справочник «ООЗТ — название — тип — площадь» (файл water_catalog.csv).{wf} Мероприятие "
                f"назначается полигону, содержащему водоём (при пересечении нескольких — с наибольшей "
                f"площадью пересечения); в основании — название, тип и площадь водоёма.")
    if r == "points_category":
        return (f"Данные: точки лесопатологии (ПВН). Файл Excel с описаниями ПВН (PVN_…xlsx) классифицируется "
                f"по 6 категориям рекомендаций (санитарная рубка, рубки ухода, надзор, посадка, эталон, "
                f"биотехния) либо используется готовый вектор с полем «Категория». Мероприятие назначается "
                f"каждому полигону с точками категории «{p.get('cat_value')}»; в основании — номера ПВН, "
                f"число точек и координаты WGS 84.")
    if r == "water_ref":
        return ("Данные: справочник водоёмов и болот ООЗТ (Excel, листы «Битца и Измайлово» и «Все ООЗТ Москвы»; "
                "лист «Сводка» не читается), полигоны гидрографии и слой болот. 1) Берутся строки с отметкой в "
                "столбце кода мероприятия (1, да, + или ●); строки «ИТОГО» и без координат пропускаются, дубли "
                "(ООЗТ + квартал + название) объединяются. 2) Точка объекта (широта и долгота WGS 84) "
                "привязывается к контуру: водоём — к полигону гидрографии, болото — к слою болот (если его нет — "
                "к гидрографии); сначала полигон, содержащий точку, затем ближайший не дальше 30 м; не найденные "
                "объекты записываются в water_ref_unmatched.csv и журнал. 3) ООЗТ и квартал определяются по "
                "геометрии: запись получает выдел с наибольшим пересечением с контуром, квартал «N (Терлецкий)» — "
                "квартал N Терлецкого массива. 4) Участок работ — контур объекта в границах ООЗТ; одна запись на "
                "объект; в основании — название, тип, площадь по справочнику, квартал и примечание. Флаги "
                "«Вкл/выкл» и «Вкл_<ООЗТ>» действуют поверх справочника; если отметок нет, мероприятие не "
                "назначается (замены другим правилом нет).")
    if r == "struct_class":
        cls = ", ".join(p.get("struct_classes") or STRUCT_MOW_DEFAULT)
        return (f"Назначается полигонам экосистем структурных классов: {cls}. Класс берётся из поля «Структурный "
                "класс» слоя территории (выбирается в окне скрипта; если не задано — из полей типа и названия) и "
                "сравнивается с названием целиком без учёта регистра и ё/е; числовой код class_value понимается "
                "по трём последним цифрам (211, 212 — суходольные луга; 221, 223, 224 — мезофитные; 222 — "
                "мезофитные антропогенные; 231–233 — гидроморфные). Участок работ — полигон экосистемы в "
                "границах ООЗТ.")
    if r == "rare_trail_vydel":
        kk = ", ".join(str(c) for c in (p.get("kk_categories") or []))
        at = _vydel_area_txt(p, " Учитываются только выделы площадью ")
        return (f"{sp} Для каждого выдела любого типа считается плотность стихийных троп — длина стихийных "
                f"троп (OSM informal=yes или свой слой) внутри выдела, делённая на его площадь (м/га). Мероприятие назначается, если "
                f"плотность больше {_fmt(p.get('abs_m_ha'))} м/га и в выделе есть хотя бы одна находка вида "
                f"Красной книги Москвы категорий {kk} (названия сравниваются с учётом синонимов справочника "
                f"цветения).{at + '.' if at else ''} Участок работ — весь выдел; в основании — виды и их "
                "категории, плотность троп и площадь выдела."
                + (" Сроки — по справочнику цветения: объединение месяцев цветения видов выдела "
                   "(плаун годичный — VII–VIII)." if p.get("season_flower") else ""))
    if r == "kolokolchik":
        lo, hi = p.get("ring") or (500.0, 1000.0)
        return (f"Высаживание видов Красной книги в типичных местообитаниях. Зоны: кольцо "
                f"{lo:.0f}–{hi:.0f} м от границы парка (внутрь), полоса {_fmt(p['road_dist'])} м от дорог "
                f"с твёрдым покрытием (если слой дорог не задан — полоса не учитывается) и полоса вдоль "
                f"границ выделов. Мероприятие назначается полигонам заданных типов местообитаний, "
                f"пересекающим зоны; в основании — перечень зон и список высаживаемых видов: "
                f"широколиственный лес — первоцветы и колокольчик широколиственный; луга естественные — "
                f"колокольчики персиколистный и раскидистый; понижения и западины — купальница и калужница.")
    return ""


def _vydel_area_txt(p: dict, prefix: str) -> str:
    """Ограничение по площади выдела: «больше мин.» и «не больше макс.»."""
    lo, hi = p.get("vydel_min_ha"), p.get("vydel_max_ha")
    if lo is None and hi is None:
        return ""
    parts = ([f"больше {_fmt(lo)} га"] if lo is not None else []) + \
            ([f"не больше {_fmt(hi)} га"] if hi is not None else [])
    return prefix + " и ".join(parts)


def describe_rule(measure: dict, full: bool = False) -> str:
    return describe_spec(measure_spec(measure), full)


def spec_params_lines(spec: dict) -> list[str]:
    """Параметры правила с источником значения — для окна скрипта."""
    out = []
    for k, v in spec.get("p", {}).items():
        src = spec["src"].get(k, "")
        if k == "types":
            val = v[1] if v else "любые"
        elif k in ("category", "groups", "taxon"):
            val = ", ".join(v) if v else "любые"
        elif k == "sources":
            val = ", ".join(SPECIES_KINDS[s] for s in v) if v else "точки, полигоны, SDM"
        elif k == "min_share":
            val = f"{_fmt(spec_share(spec['p']))}"
        elif isinstance(v, bool):
            val = "да" if v else "нет"
        elif k == "ring" and isinstance(v, (tuple, list)):
            val = f"{_fmt(v[0])}–{_fmt(v[1])}"
        elif isinstance(v, (tuple, list)):
            val = ", ".join(str(x) for x in v) if v else "любые"
        elif isinstance(v, str):
            val = v
        else:
            val = _fmt(v)
        out.append(f"{SPEC_TITLE[k]}: {val}  ({src})")
    return out


def _cfg_has(cfg: Config, key: str) -> bool:
    v = getattr(cfg, key, None)
    if key in REF_DEFAULTS:                 # справочники: файл должен существовать
        return bool(v) and Path(str(v)).exists()
    return bool(v) if not isinstance(v, list) else len(v) > 0


def matching_species_layers(cfg, p: dict, kinds=None) -> list[dict]:
    """Загруженные слои видов, подходящие мероприятию по категории, группе и типу данных."""
    out = []
    for l in species_layers(cfg):
        if p.get("category") and l["category"] not in p["category"]:
            continue
        if not group_matches(l["group"], p.get("groups")):
            continue
        if kinds is not None and l["kind"] not in kinds:
            continue
        if p.get("sources") and l["kind"] not in p["sources"]:
            continue
        out.append(l)
    return out


DATA_KEYS = {"dem": ["dem_path"], "trails": ["trails"], "hydro": ["hydro_lines", "hydro_polys"],
             "species": ["species"], "drainage": ["drainage"], "wetlands": ["wetlands"],
             "roads": ["roads_paved", "trails"], "lesopat": ["lesopat_points"],
             "water_ref": ["water_ref"], "flowering": ["flowering_ref"]}


_LESOPAT_PREVIEW_CACHE: dict = {}


def _lesopat_preview_counts(paths) -> tuple[bool, dict, str]:
    """Быстрая проверка выбранных файлов ПВН для окна предпросмотра."""
    key = tuple(str(p).strip() for p in (paths or []) if str(p).strip())
    if not key:
        return False, {}, ""
    sig = []
    for p in key:
        try:
            st = os.stat(p)
            sig.append((p, st.st_mtime_ns, st.st_size))
        except OSError:
            sig.append((p, 0, 0))
    ckey = tuple(sig)
    if ckey in _LESOPAT_PREVIEW_CACHE:
        return _LESOPAT_PREVIEW_CACHE[ckey]
    counts: dict[str, int] = {}
    ok_any = False
    bad_names = []
    for p in key:
        try:
            g = read_lesopat(p, quiet=True)
        except Exception:
            g = None
        if g is not None and not g.empty:
            ok_any = True
            for c, n in g["cat"].value_counts().items():
                counts[str(c)] = counts.get(str(c), 0) + int(n)
        else:
            bad_names.append(Path(p).name)
    warn = (f"в файле {', '.join(bad_names)} нет столбца «Рекомендации»/«Категория» или валидных координат"
            if bad_names else "")
    res = (ok_any, counts, warn)
    _LESOPAT_PREVIEW_CACHE[ckey] = res
    return res


def measure_status(measure: dict, cfg: Config, type_counts: dict | None = None) -> dict:
    """Статус мероприятия при текущем наборе данных: ok / missing / off / na (+ текст, нужные слои)."""
    spec = measure_spec(measure)
    rule, p = spec["rule"], spec.get("p", {})
    need, layers = [], []
    if rule:
        for d in RULES[rule]["data"]:
            if d == "species":
                need.append("слои видов (" + _sel_txt(p) + ")")
            else:
                need.append(" или ".join(LAYER_TITLE[k] for k in DATA_KEYS[d]))
        if p.get("types"):
            need.append("поле типа в слое территории")
    need_txt = "; ".join(need) if need else "только слой территории"
    base = dict(need=need_txt, rule=rule, spec=spec, layers=layers)
    if not measure.get("on"):
        return dict(base, state="off", text="выключено в таблице")
    if rule is None:
        return dict(base, state="na", text=spec["reason"], need="—")
    if spec["errors"]:
        return dict(base, state="missing", text="ошибка в таблице: " + "; ".join(spec["errors"]))
    missing = []
    for d in RULES[rule]["data"]:
        if d == "species":
            kinds = ["points"] if rule == "species_density" else None
            layers[:] = matching_species_layers(cfg, p, kinds)
            if not layers:
                missing.append("подходящие слои видов (" + _sel_txt(p) + ")")
        elif not any(_cfg_has(cfg, k) for k in DATA_KEYS[d]):
            missing.append(" или ".join(LAYER_TITLE[k] for k in DATA_KEYS[d]))
    if missing:
        return dict(base, state="missing", text="не хватает: " + "; ".join(missing))
    txt = "будет рассчитано"
    if rule == "water_ref":
        try:
            _CUR_CFG_prev = globals().get("_CUR_CFG")
            globals()["_CUR_CFG"] = cfg
            wr = water_ref_marked(str(measure.get("code") or ""))
            globals()["_CUR_CFG"] = _CUR_CFG_prev
            if wr is not None:
                txt += f" (отмечено в справочнике: {len(wr)})"
                if wr.empty:
                    return dict(base, state="missing", text="нет отметок в справочнике водоёмов")
        except Exception:
            pass
    if rule == "points_category" and _cfg_has(cfg, "lesopat_points"):
        ok_lp, lp_cnt, lp_warn = _lesopat_preview_counts(cfg.lesopat_points)
        if not ok_lp:
            return dict(base, state="missing",
                        text=lp_warn or "в файле ПВН не найдены рекомендации или координаты")
        cat_k = str(p.get("cat_value") or "").strip().lower().replace("ё", "е")
        txt += f" (точек «{cat_k}»: {lp_cnt.get(cat_k, 0)})"
        if lp_warn:
            txt += f"; внимание: {lp_warn}"
    keys = (p.get("types") or (None,))[0]
    if keys and type_counts is not None:
        n = sum(v for t, v in type_counts.items() if type_matches(t, keys))
        if n == 0:
            return dict(base, state="missing", text="в слое нет полигонов типа «%s»" % p["types"][1])
        txt += f" ({n} полигонов типа «{p['types'][1]}»)"
    if layers:
        txt += f"; слоёв видов: {len(layers)}"
    if rule == "species_trails" and not cfg.trails:
        txt += " — без троп: по самим зонам видов"
    return dict(base, state="ok", text=txt)


# =====================================================================================
#                       ОКНО ВЫБОРА ИСХОДНЫХ ДАННЫХ (tkinter, одно окно)
# =====================================================================================
VEC_TYPES = [("Векторные данные", "*.shp *.gpkg *.geojson *.json *.kml *.dxf *.gml"), ("Все файлы", "*.*")]
RAS_TYPES = [("Растры", "*.tif *.tiff *.asc *.img *.vrt"), ("Все файлы", "*.*")]
XLS_TYPES = [("Excel", "*.xlsx *.xlsm *.xls"), ("Все файлы", "*.*")]
VEC_XLS_TYPES = [
    ("Таблицы Excel и векторные слои", "*.xlsx *.xlsm *.xls *.shp *.gpkg *.geojson *.json *.kml *.gml"),
    ("Таблицы Excel", "*.xlsx *.xlsm *.xls"),
    ("Векторные данные", "*.shp *.gpkg *.geojson *.json *.kml *.gml"),
    ("Все файлы", "*.*"),
]
JSON_TYPES = [("Конфигурация", "*.json"), ("Все файлы", "*.*")]
NONE_FIELD = "— нет —"

# поле структурного класса (I.7 «Мозаичное выкашивание»): задаётся пользователем, необязательно
STRUCT_FIELD, STRUCT_FIELD_TITLE = "STRUCT", "Структурный класс (для выкашивания)"
STRUCT_FIELD_CANDS = ["Структурный класс", "Структурный_класс", "Struct_class", "STRUCT"]

# поля слоя территории: ключ, подпись, обязательность, кандидаты для автоподбора
FIELD_SPECS = {
    "kv_vyd": [
        ("OOZT", "ООЗТ", True, ["OOZT", "ООЗТ", "Name_OOZT"]),
        ("KV", "Квартал", True, ["KV", "FID_kv", "kv", "Квартал", "kvartal"]),
        ("VYD", "Выдел", True, ["VYD", "FID_ecosys", "vyd", "Выдел", "vydel"]),
        ("TYPE", "Тип (лес / луг / водные / околоводные)", True, ["Type", "TYPE", "Тип"]),
        ("NAME", "Название экосистемы", False, ["Name", "NAME", "Название", "Ecosystem"]),
        (STRUCT_FIELD, STRUCT_FIELD_TITLE, False, STRUCT_FIELD_CANDS),
    ],
    "eco": [
        ("ID", "Уникальный ID экосистемы", True, ["ID", "Id", "id", "eco_id", "ECO_ID", "uid", "UID",
                                                  "FID_ecosys", "fid_1"]),
        ("TYPE", "Тип (лес / луг / водные / околоводные)", True, ["Type", "TYPE", "Тип"]),
        ("NAME", "Название экосистемы", False, ["Name", "NAME", "Название", "Ecosystem"]),
        ("OOZT", "ООЗТ (если поля нет — название ниже)", False, ["OOZT", "ООЗТ", "Name_OOZT"]),
        (STRUCT_FIELD, STRUCT_FIELD_TITLE, False, STRUCT_FIELD_CANDS),
    ],
    # kv_cut: поля слоя экосистем; квартал и выдел назначаются разрезанием (kvnet_*)
    "kv_cut": [
        ("TYPE", "Тип (лес / луг / водные / околоводные)", True, ["Type", "TYPE", "Тип", "Расшифровка"]),
        ("NAME", "Название экосистемы", False, ["Name", "NAME", "Название", "Ecosystem", "Расшифровка"]),
        ("OOZT", "ООЗТ (если поля нет — из квартальной сети)", False, ["OOZT", "ООЗТ", "Name_OOZT"]),
        ("ID", "Исходный ID экосистемы (по желанию)", False, ["ID", "Id", "id", "eco_id", "ECO_ID", "uid",
                                                              "FID_ecosys", "fid_1"]),
        (STRUCT_FIELD, STRUCT_FIELD_TITLE, False, STRUCT_FIELD_CANDS),
    ],
}


def _vector_fields(path: str, layer: str | None):
    """Имена полей и таблица атрибутов (без геометрии) — для подбора полей в интерфейсе."""
    df = pyogrio.read_dataframe(path, layer=layer, read_geometry=False)
    ren = {c: _remojibake(c) for c in df.columns}
    ren = {k: v for k, v in ren.items() if v != k and v not in df.columns}
    if ren:
        df = df.rename(columns=ren)
    return list(df.columns), df


class SetupWindow:
    """Одно окно: режим территории, обязательные и тематические слои, живая сводка мероприятий."""

    C_OK, C_WARN, C_OFF, C_MUTED = "#2E7D32", "#B35C00", "#8A8A8A", "#5F6368"

    def __init__(self, cfg: Config | None = None):
        import tkinter as tk
        from tkinter import ttk, font as tkfont
        self.tk, self.ttk = tk, ttk
        if os.name == "nt":
            try:
                import ctypes
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass
        self.root = tk.Tk()
        self.root.title("ООЗТ: природоохранные мероприятия — исходные данные")
        self.root.geometry("1480x900")
        self.root.minsize(1180, 700)
        self.cfg = cfg or Config()
        self.result: Config | None = None
        self.measures: list[dict] = []
        self.type_counts: dict | None = None
        self.attr_df = None
        self.field_vars: dict = {}
        self.layer_rows: dict = {}

        fam = "Segoe UI" if os.name == "nt" else ("Noto Sans" if "Noto Sans" in tkfont.families() else None)
        for fn in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont", "TkCaptionFont"):
            try:
                f_ = tkfont.nametofont(fn)
                f_.configure(size=10, **({"family": fam} if fam else {}))
            except Exception:
                pass
        base = tkfont.nametofont("TkDefaultFont")
        self.f_bold = base.copy(); self.f_bold.configure(weight="bold")
        self.f_small = base.copy(); self.f_small.configure(size=9)
        self.f_head = base.copy(); self.f_head.configure(size=13, weight="bold")

        st = ttk.Style(self.root)
        st.configure("Treeview", rowheight=24)
        # обход ошибки Tk 8.6.9: цвета тегов Treeview не применяются
        def _fixed(opt):
            return [e for e in st.map("Treeview", query_opt=opt) if e[:2] != ("!disabled", "!selected")]
        st.map("Treeview", foreground=_fixed("foreground"), background=_fixed("background"))

        self._build()
        self._load_from_cfg(self.cfg)

    # ------------------------------------------------------------------ разметка
    def _build(self):
        tk, ttk = self.tk, self.ttk
        top = ttk.Frame(self.root, padding=(12, 10, 12, 4))
        top.pack(fill="x")
        ttk.Label(top, text="Исходные данные для расчёта мероприятий", font=self.f_head).pack(anchor="w")
        ttk.Label(top, foreground=self.C_MUTED, text=(
            "Обязательны таблица мероприятий и слой территории. Остальные слои необязательны: справа видно, "
            "какие мероприятия будут рассчитаны при выбранном наборе данных и чего не хватает остальным.")
        ).pack(anchor="w")

        pw = ttk.PanedWindow(self.root, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=12, pady=6)

        # ---- левая колонка (прокручиваемая)
        left_outer = ttk.Frame(pw)
        canvas = tk.Canvas(left_outer, highlightthickness=0, borderwidth=0)
        vsb = ttk.Scrollbar(left_outer, orient="vertical", command=canvas.yview)
        self.left = ttk.Frame(canvas)
        self.left.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=self.left, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-2, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(2, "units"))
        pw.add(left_outer, weight=5)

        # 1. таблица мероприятий
        s1 = ttk.LabelFrame(self.left, text=" 1. Таблица мероприятий (обязательно) ", padding=8)
        s1.pack(fill="x", pady=(0, 8))
        self.v_xlsx = tk.StringVar()
        r = ttk.Frame(s1); r.pack(fill="x")
        ttk.Entry(r, textvariable=self.v_xlsx, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(r, text="Выбрать…", command=self._pick_xlsx).pack(side="left", padx=(6, 0))
        self.l_xlsx = ttk.Label(s1, foreground=self.C_MUTED, font=self.f_small,
                                text="Excel «Мероприятия по Приказу»: колонки «Мероприятия», «Вкл/выкл», «Метод выявления»")
        self.l_xlsx.pack(anchor="w", pady=(4, 0))
        ttk.Label(s1, foreground=self.C_MUTED, font=self.f_small, wraplength=700, justify="left",
                  text="Способ назначения задаётся столбцами таблицы: «Правило назначения», «Типы полигонов», "
                       "«Виды: категория», «Виды: группы организмов», «Виды: таксоны (лат.)», «Виды: данные», "
                       "«Расстояние, м», «Окно плотности, м», «Квантиль», «Уклон более, %», «Порог SDM (0–1)», "
                       "«Мин. площадь участка, га», «Мин. доля полигона, %». Если «Правило назначения» пусто, "
                       "правило определяется по тексту «Метода выявления». Справка — на листе «Справка по правилам»."
                  ).pack(anchor="w", pady=(2, 0))
        r = ttk.Frame(s1); r.pack(fill="x", pady=(4, 0))
        ttk.Button(r, text="Записать правила и описания в таблицу…", command=self._update_table).pack(side="left")
        ttk.Label(r, foreground=self.C_MUTED, font=self.f_small,
                  text="  столбцы правил, справка, описания расчёта").pack(side="left")

        # 2. слой территории
        s2 = ttk.LabelFrame(self.left, text=" 2. Слой территории — единицы назначения мероприятий (обязательно) ",
                            padding=8)
        s2.pack(fill="x", pady=(0, 8))
        self.v_mode = tk.StringVar(value="kv_vyd")
        mf = ttk.Frame(s2); mf.pack(fill="x")
        ttk.Radiobutton(mf, text="Кварталы и выделы", value="kv_vyd", variable=self.v_mode,
                        command=self._on_mode).grid(row=0, column=0, sticky="w")
        ttk.Label(mf, foreground=self.C_MUTED, font=self.f_small,
                  text="полигоны выделов с полями ООЗТ, квартала и выдела; отчёт и схемы — по кварталам"
                  ).grid(row=1, column=0, sticky="w", padx=(24, 0))
        ttk.Radiobutton(mf, text="Экосистемы, разрезанные квартальной сетью", value="kv_cut",
                       variable=self.v_mode, command=self._on_mode).grid(row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Label(mf, foreground=self.C_MUTED, font=self.f_small,
                  text="слой экосистем + вектор квартальной сети; экосистемы режутся границами кварталов, "
                       "в каждом квартале — своя нумерация выделов с северо-запада; отчёт — по кварталам"
                  ).grid(row=3, column=0, sticky="w", padx=(24, 0))
        ttk.Radiobutton(mf, text="Слой экосистем с уникальным ID (без кварталов)", value="eco",
                        variable=self.v_mode, command=self._on_mode).grid(row=4, column=0, sticky="w", pady=(4, 0))
        ttk.Label(mf, foreground=self.C_MUTED, font=self.f_small,
                  text="полигоны экосистем; отчёт и схемы — по листам регулярной сетки или по полю группировки"
                  ).grid(row=5, column=0, sticky="w", padx=(24, 0))

        self.v_oozt = tk.StringVar()
        r = ttk.Frame(s2); r.pack(fill="x", pady=(8, 0))
        ttk.Label(r, text="Файл:", width=8).pack(side="left")
        ttk.Entry(r, textvariable=self.v_oozt, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(r, text="Выбрать…", command=self._pick_oozt).pack(side="left", padx=(6, 0))
        r = ttk.Frame(s2); r.pack(fill="x", pady=(4, 0))
        ttk.Label(r, text="Слой GPKG:", width=11).pack(side="left")
        self.v_layer = tk.StringVar()
        self.cb_layer = ttk.Combobox(r, textvariable=self.v_layer, state="disabled", width=40)
        self.cb_layer.pack(side="left")
        self.cb_layer.bind("<<ComboboxSelected>>", lambda e: self._read_oozt_fields())

        self.fields_frame = ttk.Frame(s2)
        self.fields_frame.pack(fill="x", pady=(8, 0))
        # kv_cut: файл квартальной сети, поле номера квартала, поле ООЗТ сети, название ООЗТ
        self.kvnet_frame = ttk.Frame(s2)
        self.v_kvnet = tk.StringVar()
        self.v_kvnet_layer = tk.StringVar()
        self.v_kvnet_field = tk.StringVar()
        self.v_kvnet_oozt = tk.StringVar(value=NONE_FIELD)
        self.v_oozt_name = tk.StringVar()
        r = ttk.Frame(self.kvnet_frame); r.pack(fill="x")
        ttk.Label(r, text="Квартальная сеть:", width=18).pack(side="left")
        ttk.Entry(r, textvariable=self.v_kvnet, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(r, text="Выбрать…", command=self._pick_kvnet).pack(side="left", padx=(6, 0))
        r = ttk.Frame(self.kvnet_frame); r.pack(fill="x", pady=(3, 0))
        ttk.Label(r, text="Слой GPKG:", width=18).pack(side="left")
        self.cb_kvnet_layer = ttk.Combobox(r, textvariable=self.v_kvnet_layer, state="disabled", width=38)
        self.cb_kvnet_layer.pack(side="left")
        self.cb_kvnet_layer.bind("<<ComboboxSelected>>", lambda e: self._read_kvnet_fields())
        r = ttk.Frame(self.kvnet_frame); r.pack(fill="x", pady=(3, 0))
        ttk.Label(r, text="Поле номера квартала:", width=18).pack(side="left")
        self.cb_kvnet_field = ttk.Combobox(r, textvariable=self.v_kvnet_field, state="readonly", width=28)
        self.cb_kvnet_field.pack(side="left")
        self.cb_kvnet_field.bind("<<ComboboxSelected>>", lambda e: self._refresh())
        r = ttk.Frame(self.kvnet_frame); r.pack(fill="x", pady=(3, 0))
        ttk.Label(r, text="Поле ООЗТ сети:", width=18).pack(side="left")
        self.cb_kvnet_oozt = ttk.Combobox(r, textvariable=self.v_kvnet_oozt, state="readonly", width=28)
        self.cb_kvnet_oozt.pack(side="left")
        self.cb_kvnet_oozt.bind("<<ComboboxSelected>>", lambda e: self._refresh())
        ttk.Label(r, foreground=self.C_MUTED, font=self.f_small,
                  text="необязательно").pack(side="left", padx=(6, 0))
        r = ttk.Frame(self.kvnet_frame); r.pack(fill="x", pady=(3, 0))
        ttk.Label(r, text="Название ООЗТ:", width=18).pack(side="left")
        ttk.Entry(r, textvariable=self.v_oozt_name, width=30).pack(side="left")
        ttk.Label(r, foreground=self.C_MUTED, font=self.f_small,
                  text="если нет поля ООЗТ ни в экосистемах, ни в сети").pack(side="left", padx=(6, 0))
        self.eco_frame = ttk.Frame(s2)
        self.v_group = tk.StringVar(value=NONE_FIELD)
        self.v_sheet = tk.StringVar(value="1000")
        ttk.Label(self.eco_frame, text="Название ООЗТ:").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.eco_frame, textvariable=self.v_oozt_name, width=30).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(self.eco_frame, foreground=self.C_MUTED, font=self.f_small,
                  wraplength=300, justify="left", text="используется, если поле ООЗТ не выбрано").grid(row=0, column=2, sticky="w")
        ttk.Label(self.eco_frame, text="Группировка отчёта:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.cb_group = ttk.Combobox(self.eco_frame, textvariable=self.v_group, state="readonly", width=28)
        self.cb_group.grid(row=1, column=1, sticky="w", padx=6, pady=(4, 0))
        self.cb_group.bind("<<ComboboxSelected>>", lambda e: self._refresh())
        ttk.Label(self.eco_frame, foreground=self.C_MUTED, font=self.f_small,
                  wraplength=300, justify="left", text="поле (например, участок или район); «нет» — листы регулярной сетки"
                  ).grid(row=1, column=2, sticky="w")
        ttk.Label(self.eco_frame, text="Размер листа, м:").grid(row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Spinbox(self.eco_frame, from_=200, to=5000, increment=100, textvariable=self.v_sheet, width=8,
                    command=self._refresh).grid(row=2, column=1, sticky="w", padx=6, pady=(4, 0))
        ttk.Label(self.eco_frame, foreground=self.C_MUTED, font=self.f_small,
                  wraplength=300, justify="left", text="лист = квадрат сетки; полигон относится к листу по своей внутренней точке"
                  ).grid(row=2, column=2, sticky="w")
        self.l_types = ttk.Label(s2, foreground=self.C_MUTED, font=self.f_small, wraplength=700, justify="left")
        self.l_types.pack(anchor="w", pady=(6, 0))

        # 3. тематические слои
        s3 = ttk.LabelFrame(self.left, text=" 3. Тематические слои (необязательно) ", padding=8)
        s3.pack(fill="x", pady=(0, 8))
        ttk.Label(s3, foreground=self.C_MUTED, font=self.f_small, wraplength=700, justify="left",
                  text="✔ — слой выбран;  ! — нужен включённым мероприятиям, но не выбран;  ○ — не используется "
                       "включёнными мероприятиями. Коды мероприятий — по таблице (колонка с номерами I.1, II.1…).",
                  ).pack(anchor="w", pady=(0, 4))
        for spec in LAYER_SPECS:
            self._layer_row(s3, spec)
        # стихийные тропы: скачать из OSM / свой слой / не использовать
        fi = ttk.Frame(s3)
        fi.pack(fill="x", pady=(6, 0))
        ttk.Label(fi, text="Стихийные тропы (для мероприятий со столбцом «Только стихийные тропы» = да):",
                  wraplength=600, justify="left").grid(row=0, column=0, columnspan=2, sticky="w")
        self.v_inf_mode = tk.StringVar(value="auto")
        self._inf_files: list[str] = []
        f_osm = ttk.Frame(fi)
        f_osm.grid(row=1, column=0, columnspan=2, sticky="w")
        ttk.Radiobutton(f_osm, text="из OSM:", value="auto",
                        variable=self.v_inf_mode, command=self._inf_changed).pack(side="left")
        self.v_osm_inf = tk.BooleanVar(value=True)
        self.v_osm_unp = tk.BooleanVar(value=False)
        ttk.Checkbutton(f_osm, text="стихийные", variable=self.v_osm_inf,
                        command=self._inf_changed).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(f_osm, text="грунтовые", variable=self.v_osm_unp,
                        command=self._inf_changed).pack(side="left", padx=(8, 0))
        f_src = ttk.Frame(fi)
        f_src.grid(row=2, column=0, columnspan=2, sticky="w", padx=(22, 0))
        ttk.Label(f_src, text="источник OSM:").pack(side="left")
        self.v_osm_src = tk.StringVar(value=OSM_SOURCES["auto"][0])
        self._osm_src_keys = {v[0]: k for k, v in OSM_SOURCES.items()}
        cb = ttk.Combobox(f_src, textvariable=self.v_osm_src, state="readonly", width=46,
                          values=[v[0] for v in OSM_SOURCES.values()])
        cb.pack(side="left", padx=(6, 0))
        cb.bind("<<ComboboxSelected>>", lambda e: self._inf_changed())
        self._osm_pbf = ""
        ttk.Button(f_src, text="Файл .pbf…", command=self._pick_osm_pbf).pack(side="left", padx=(8, 0))
        f_own = ttk.Frame(fi)
        f_own.grid(row=3, column=0, columnspan=2, sticky="w")
        ttk.Radiobutton(f_own, text="свой слой", value="files",
                        variable=self.v_inf_mode, command=self._inf_changed).pack(side="left")
        ttk.Button(f_own, text="Выбрать…", command=self._pick_informal).pack(side="left", padx=(8, 0))
        ttk.Radiobutton(fi, text="не использовать", value="none",
                        variable=self.v_inf_mode, command=self._inf_changed).grid(row=4, column=0, sticky="w")
        self.l_inf = ttk.Label(fi, foreground=self.C_MUTED, font=self.f_small, wraplength=600, justify="left")
        self.l_inf.grid(row=5, column=0, columnspan=2, sticky="w")
        self._inf_changed()
        self._species_block(s3)

        # ---- правая колонка
        right = ttk.Frame(pw, padding=(8, 0, 0, 0))
        pw.add(right, weight=4)
        self._pw = pw
        self.root.bind("<Map>", self._place_sash, add="+")
        s4 = ttk.LabelFrame(right, text=" 4. Что будет рассчитано ", padding=8)
        s4.pack(fill="both", expand=True)
        self.l_sum = ttk.Label(s4, font=self.f_bold, text="Выберите таблицу мероприятий")
        self.l_sum.pack(anchor="w", pady=(0, 6))
        tf = ttk.Frame(s4); tf.pack(fill="both", expand=True)
        cols = ("code", "name", "status")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c, t, w, st in (("code", "Код", 60, False), ("name", "Мероприятие", 300, True),
                            ("status", "Статус при выбранных данных", 300, True)):
            self.tree.heading(c, text=t, anchor="w")
            self.tree.column(c, width=w, stretch=st, anchor="w")
        self.tree.tag_configure("ok", foreground=self.C_OK)
        self.tree.tag_configure("missing", foreground=self.C_WARN)
        self.tree.tag_configure("off", foreground=self.C_OFF)
        self.tree.tag_configure("na", foreground=self.C_OFF)
        ys = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        df_ = ttk.Frame(s4); df_.pack(fill="both", expand=False, pady=(6, 0))
        self.detail = tk.Text(df_, height=14, wrap="word", relief="flat", background="#F4F5F6",
                              font=self.f_small, padx=6, pady=4)
        dsb = ttk.Scrollbar(df_, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.pack(side="left", fill="both", expand=True)
        dsb.pack(side="right", fill="y")
        self.detail.insert("1.0", "Выберите мероприятие в списке, чтобы увидеть правило назначения, его параметры, "
                                  "полное описание расчёта и нужные слои.")
        self.detail.configure(state="disabled")

        s5 = ttk.LabelFrame(right, text=" 5. Результаты ", padding=8)
        s5.pack(fill="x", pady=(8, 0))
        self.v_out = tk.StringVar()
        r = ttk.Frame(s5); r.pack(fill="x")
        ttk.Label(r, text="Папка:", width=8).pack(side="left")
        ttk.Entry(r, textvariable=self.v_out).pack(side="left", fill="x", expand=True)
        ttk.Button(r, text="Выбрать…", command=self._pick_out).pack(side="left", padx=(6, 0))
        r = ttk.Frame(s5); r.pack(fill="x", pady=(6, 0))
        self.l_kv = ttk.Label(r, text="Кварталы для отчёта:", width=22)
        self.l_kv.pack(side="left")
        self.v_kv = tk.StringVar()
        ttk.Entry(r, textvariable=self.v_kv, width=30).pack(side="left")
        ttk.Label(r, foreground=self.C_MUTED, font=self.f_small, text="  через запятую; пусто — все").pack(side="left")
        self.v_basemap = tk.BooleanVar(value=True)
        ttk.Checkbutton(s5, text="Подложка OpenStreetMap на схемах (нужен интернет)",
                        variable=self.v_basemap).pack(anchor="w", pady=(6, 0))
        r = ttk.Frame(s5); r.pack(fill="x", pady=(4, 0))
        ttk.Label(r, text="Общий порог: участок-основание не меньше").pack(side="left")
        self.v_minshare = tk.StringVar(value="25")
        ttk.Entry(r, textvariable=self.v_minshare, width=5).pack(side="left", padx=4)
        ttk.Label(r, text="% площади полигона (0 — без порога; если в таблице не указан свой)").pack(side="left")
        self.v_split = tk.BooleanVar(value=True)
        ttk.Checkbutton(s5, text="Делить многочастные объекты на отдельные контуры (ID-1, ID-2, …) — "
                                 "мероприятие назначается каждому контуру",
                        variable=self.v_split).pack(anchor="w", pady=(4, 0))
        r = ttk.Frame(s5); r.pack(fill="x", pady=(6, 0))
        ttk.Label(r, text="Отчёт:", width=8).pack(side="left")
        self.v_formats = {k: tk.BooleanVar(value=k in ("docx", "pdf")) for k in REPORT_FORMATS}
        for k, t in REPORT_FORMATS.items():
            ttk.Checkbutton(r, text=t, variable=self.v_formats[k]).pack(side="left", padx=(0, 10))
        ttk.Label(s5, foreground=self.C_MUTED, font=self.f_small, wraplength=640, justify="left",
                  text="report.md создаётся всегда. PDF и DOC делаются через Microsoft Word (если установлен), "
                       "иначе через LibreOffice; без них PDF собирается встроенной вёрсткой, DOC — не создаётся."
                  ).pack(anchor="w", pady=(2, 0))

        # ---- кнопки
        bot = ttk.Frame(self.root, padding=(12, 4, 12, 12))
        bot.pack(fill="x")
        ttk.Button(bot, text="Загрузить конфигурацию…", command=self._load_cfg_file).pack(side="left")
        ttk.Button(bot, text="Сохранить конфигурацию…", command=self._save_cfg_file).pack(side="left", padx=6)
        ttk.Button(bot, text="Отмена", command=self._cancel).pack(side="right")
        ttk.Button(bot, text="Запустить расчёт", command=self._run).pack(side="right", padx=6)
        self.root.protocol("WM_DELETE_WINDOW", self._cancel)

    def _place_sash(self, _e=None):
        if getattr(self, "_sash_done", False):
            return
        self.root.update_idletasks()
        w = self._pw.winfo_width()
        if w > 100:
            self._pw.sashpos(0, int(w * 0.53))
            self._sash_done = True

    def _layer_row(self, parent, spec):
        ttk = self.ttk
        key = spec["key"]
        card = ttk.Frame(parent, padding=(0, 5, 0, 5))
        card.pack(fill="x")
        ttk.Separator(parent, orient="horizontal").pack(fill="x")
        card.columnconfigure(1, weight=1)
        st = ttk.Label(card, text="○", width=2, foreground=self.C_OFF, font=self.f_bold)
        st.grid(row=0, column=0, sticky="w")
        ttk.Label(card, text=spec["title"], font=self.f_bold).grid(row=0, column=1, sticky="w")
        bf = ttk.Frame(card)
        bf.grid(row=0, column=2, rowspan=2, sticky="ne")
        ttk.Button(bf, text="Добавить…" if spec["multi"] else "Выбрать…", width=11,
                   command=lambda s_=spec: self._pick_layer(s_)).pack(side="left")
        ttk.Button(bf, text="Очистить", width=9, command=lambda k=key: self._clear_layer(k)).pack(side="left", padx=(4, 0))
        ttk.Label(card, text=spec["what"] + ("; можно несколько файлов" if spec["multi"] else ""),
                  foreground=self.C_MUTED, font=self.f_small, wraplength=520, justify="left"
                  ).grid(row=1, column=1, sticky="w")
        used = ttk.Label(card, text="", font=self.f_small, wraplength=640, justify="left")
        used.grid(row=2, column=1, columnspan=2, sticky="w")
        files = ttk.Label(card, text="не выбран", foreground=self.C_OFF, font=self.f_small,
                          wraplength=640, justify="left")
        files.grid(row=3, column=1, columnspan=2, sticky="w")
        self.layer_rows[key] = dict(status=st, used=used, files=files, spec=spec)

    # ------------------------------------------------------------------ слои видов
    SP_TYPES = [("Слои видов (векторы и растры SDM)",
                 "*.shp *.gpkg *.geojson *.json *.kml *.gml *.tif *.tiff *.asc *.img *.vrt"), ("Все файлы", "*.*")]

    def _species_block(self, parent):
        tk, ttk = self.tk, self.ttk
        card = ttk.Frame(parent, padding=(0, 6, 0, 4))
        card.pack(fill="x")
        card.columnconfigure(1, weight=1)
        self.sp_status = ttk.Label(card, text="○", width=2, foreground=self.C_OFF, font=self.f_bold)
        self.sp_status.grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="Виды и редкие экосистемы", font=self.f_bold).grid(row=0, column=1, sticky="w")
        bf = ttk.Frame(card)
        bf.grid(row=0, column=2, sticky="ne")
        ttk.Button(bf, text="Добавить…", width=11, command=self._sp_add).pack(side="left")
        ttk.Button(bf, text="Удалить", width=9, command=self._sp_remove).pack(side="left", padx=(4, 0))
        ttk.Label(card, foreground=self.C_MUTED, font=self.f_small, wraplength=640, justify="left",
                  text="Точки и полигоны находок (shp, gpkg…) или модели SDM (растры 0–1); можно выбрать сразу "
                       "несколько файлов. Для каждого слоя укажите категорию (инвазивные / охраняемые / прочие) и "
                       "группу организмов; тип данных определяется автоматически. Мероприятия выбирают слои по "
                       "этим признакам — столбцы «Виды: …» таблицы мероприятий."
                  ).grid(row=1, column=1, columnspan=2, sticky="w")
        tf = ttk.Frame(card)
        tf.grid(row=2, column=1, columnspan=2, sticky="we", pady=(4, 0))
        cols = ("file", "kind", "category", "group", "field")
        self.sp_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended", height=5)
        for c, t, w in (("file", "Файл", 190), ("kind", "Тип данных", 130), ("category", "Категория", 100),
                        ("group", "Группа организмов", 150), ("field", "Поле названия", 90)):
            self.sp_tree.heading(c, text=t, anchor="w")
            self.sp_tree.column(c, width=w, stretch=(c == "file"), anchor="w")
        ys = ttk.Scrollbar(tf, orient="vertical", command=self.sp_tree.yview)
        self.sp_tree.configure(yscrollcommand=ys.set)
        self.sp_tree.pack(side="left", fill="x", expand=True)
        ys.pack(side="right", fill="y")
        self.sp_tree.bind("<<TreeviewSelect>>", self._sp_on_select)
        ed = ttk.Frame(card)
        ed.grid(row=3, column=1, columnspan=2, sticky="w", pady=(4, 0))
        self.v_sp_cat = tk.StringVar(value="охраняемые")
        self.v_sp_grp = tk.StringVar(value="сосудистые растения")
        self.v_sp_kind = tk.StringVar(value="")
        self.v_sp_field = tk.StringVar(value="авто")
        ttk.Label(ed, text="Категория:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(ed, textvariable=self.v_sp_cat, values=SPECIES_CATEGORIES, state="readonly",
                     width=12).grid(row=0, column=1, sticky="w", padx=(4, 10))
        ttk.Label(ed, text="Группа:").grid(row=0, column=2, sticky="w")
        ttk.Combobox(ed, textvariable=self.v_sp_grp, values=SPECIES_GROUPS, width=20
                     ).grid(row=0, column=3, sticky="w", padx=(4, 10))
        ttk.Label(ed, text="Тип данных:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Combobox(ed, textvariable=self.v_sp_kind, values=list(SPECIES_KINDS.values()), state="readonly",
                     width=14).grid(row=1, column=1, sticky="w", padx=(4, 10), pady=(4, 0))
        ttk.Label(ed, text="Поле названия:").grid(row=1, column=2, sticky="w", pady=(4, 0))
        self.cb_sp_field = ttk.Combobox(ed, textvariable=self.v_sp_field, values=["авто"], width=20)
        self.cb_sp_field.grid(row=1, column=3, sticky="w", padx=(4, 10), pady=(4, 0))
        ttk.Button(ed, text="Применить к выбранным", command=self._sp_apply).grid(row=0, column=4, rowspan=2,
                                                                                 sticky="w")
        ttk.Label(card, foreground=self.C_MUTED, font=self.f_small, wraplength=640, justify="left",
                  text="Группу можно ввести вручную (новые группы организмов). «Поле названия: авто» — поле с "
                       "латинскими названиями ищется само; для редких экосистем выберите поле с названием. "
                       "У растров SDM название вида — имя файла."
                  ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(2, 0))
        self.sp_used = ttk.Label(card, text="", font=self.f_small, wraplength=640, justify="left")
        self.sp_used.grid(row=5, column=1, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Separator(parent, orient="horizontal").pack(fill="x")

    @staticmethod
    def _sp_guess(path: str, cat: str, grp: str):
        n = Path(path).stem.lower()
        if any(w in n for w in ("invas", "инваз", "alien", "чужерод")):
            cat = "инвазивные"
        elif any(w in n for w in ("rare", "redbook", "red_book", "редк", "красн", "охран", "protect")):
            cat = "охраняемые"
        for words, g in ((("vascular", "plant", "сосуд", "flora", "растен"), "сосудистые растения"),
                         (("fung", "гриб", "myco"), "грибы"), (("lichen", "лишай"), "лишайники"),
                         (("mamm", "млекоп", "theri"), "млекопитающие"), (("aves", "bird", "птиц", "орнит"), "птицы"),
                         (("amph", "rept", "herp", "fish", "земновод", "пресмык", "рыб", "низш"), "низшие позвоночные"),
                         (("insect", "насеком", "entom"), "насекомые"),
                         (("ecosys", "экосист", "habitat", "сообществ"), "редкие экосистемы")):
            if any(w in n for w in words):
                grp = g
                break
        return cat, grp

    def _sp_add(self):
        from tkinter import filedialog
        ps = list(filedialog.askopenfilenames(parent=self.root, title="Слои видов и редких экосистем — можно "
                                              "выбрать несколько", filetypes=self.SP_TYPES))
        if not ps:
            return
        cur = list(self.cfg.species_layers or [])
        have = {l.get("path") for l in cur}
        for pth in ps:
            if pth in have:
                continue
            cat, grp = self._sp_guess(pth, self.v_sp_cat.get(), norm_group(self.v_sp_grp.get()))
            cur.append(dict(path=pth, category=cat, group=grp, kind=detect_species_kind(pth), field=""))
        self.cfg.species_layers = cur
        self._sp_render()
        self._refresh()

    def _sp_remove(self):
        sel = {int(i) for i in self.sp_tree.selection()}
        if not sel:
            return
        self.cfg.species_layers = [l for i, l in enumerate(self.cfg.species_layers or []) if i not in sel]
        self._sp_render()
        self._refresh()

    def _sp_on_select(self, _e=None):
        sel = [int(i) for i in self.sp_tree.selection()]
        if not sel:
            return
        l = (self.cfg.species_layers or [])[sel[0]]
        self.v_sp_cat.set(norm_category(l.get("category")) or "прочие")
        self.v_sp_grp.set(norm_group(l.get("group")))
        self.v_sp_kind.set(SPECIES_KINDS.get(l.get("kind") or "", ""))
        self.v_sp_field.set(l.get("field") or "авто")
        vals = ["авто"]
        if len(sel) == 1 and l.get("kind") in ("points", "polygons"):
            try:
                pth = str(l["path"])
                cols, _ = _vector_fields(pth.split("|")[0], pth.split("|")[1] if "|" in pth else None)
                vals += [c for c in cols if c.lower() not in ("geometry",)]
            except Exception:
                pass
        self.cb_sp_field.configure(values=vals)

    def _sp_apply(self):
        sel = [int(i) for i in self.sp_tree.selection()]
        if not sel:
            from tkinter import messagebox
            messagebox.showinfo("Слои видов", "Выберите в списке один или несколько слоёв.", parent=self.root)
            return
        kind = {v: k for k, v in SPECIES_KINDS.items()}.get(self.v_sp_kind.get())
        fld = self.v_sp_field.get().strip()
        for i in sel:
            l = self.cfg.species_layers[i]
            l["category"] = norm_category(self.v_sp_cat.get()) or "прочие"
            l["group"] = norm_group(self.v_sp_grp.get())
            if kind:
                l["kind"] = kind
            if len(sel) == 1:
                l["field"] = "" if fld in ("", "авто") else fld
        self._sp_render()
        for i in sel:
            self.sp_tree.selection_add(str(i))
        self._refresh()

    def _sp_render(self):
        self.sp_tree.delete(*self.sp_tree.get_children())
        for i, l in enumerate(self.cfg.species_layers or []):
            kind = l.get("kind") or detect_species_kind(l.get("path", ""))
            l["kind"] = kind
            self.sp_tree.insert("", "end", iid=str(i), values=(
                Path(str(l.get("path", "")).split("|")[0]).name, SPECIES_KINDS.get(kind, "? не определён"),
                norm_category(l.get("category")) or "прочие", norm_group(l.get("group")), l.get("field") or "авто"))

    def _sp_refresh_status(self, statuses):
        n = len(self.cfg.species_layers or [])
        parts, bad = [], False
        for m, s in zip(self.measures, statuses):
            if not (s["rule"] and "species" in RULES[s["rule"]]["data"]) or not m["on"]:
                continue
            k = len(matching_species_layers(self.cfg, s["spec"].get("p", {}),
                                            ["points"] if s["rule"] == "species_density" else None))
            parts.append(f"{m.get('code') or '?'} ({_sel_txt(s['spec'].get('p', {}))}) — слоёв: {k}")
            bad |= k == 0
        txt = ("Нужны для включённых: " + ";  ".join(parts)) if parts else "Не нужны включённым мероприятиям"
        self.sp_used.configure(text=txt, foreground="#202124" if parts else self.C_MUTED)
        if n and not bad:
            self.sp_status.configure(text="✔", foreground=self.C_OK)
        elif parts and bad:
            self.sp_status.configure(text="!", foreground=self.C_WARN)
        else:
            self.sp_status.configure(text="✔" if n else "○", foreground=self.C_OK if n else self.C_OFF)

    def _update_table(self):
        from tkinter import messagebox
        if not self.cfg.measures_xlsx:
            messagebox.showinfo("Таблица мероприятий", "Сначала выберите таблицу мероприятий.", parent=self.root)
            return
        if not messagebox.askyesno(
                "Записать правила и описания в таблицу",
                "В таблицу мероприятий будут записаны:\n\n"
                "• столбцы правил («Правило назначения», «Типы полигонов», «Виды: …», «Расстояние, м» и др.) — "
                "только в строках, где «Правило назначения» пусто: правило и параметры, которые скрипт сейчас "
                "определяет по тексту «Метода выявления»;\n"
                "• выпадающие списки в этих столбцах;\n"
                "• лист «Справка по правилам»;\n"
                f"• столбец «{DESC_COLUMN}» — полное описание расчёта для каждой строки.\n\n"
                "Перед записью создаётся копия таблицы с окончанием «_до_правил». Файл должен быть закрыт в Excel. "
                "Продолжить?", parent=self.root):
            return
        try:
            info = update_measures_xlsx(self.cfg.measures_xlsx, self._collect())
        except PermissionError:
            messagebox.showerror("Таблица мероприятий", "Файл открыт в другой программе (Excel). Закройте его "
                                 "и повторите.", parent=self.root)
            return
        except Exception as e:
            messagebox.showerror("Таблица мероприятий", f"Не удалось записать: {e}", parent=self.root)
            return
        self._read_measures()
        self._refresh()
        messagebox.showinfo("Таблица мероприятий", info, parent=self.root)

    # ------------------------------------------------------------------ стихийные тропы
    def _inf_changed(self):
        m = self.v_inf_mode.get()
        if m == "files" and not self._inf_files:
            txt = "свой слой не выбран — нажмите «Выбрать…» (линии; можно несколько файлов)"
        elif m == "files":
            txt = "свой слой: " + "; ".join(Path(x).name for x in self._inf_files)
        elif m == "auto":
            parts = []
            if self.v_osm_inf.get():
                parts.append("стихийные (informal=yes)")
            if self.v_osm_unp.get():
                parts.append(f"грунтовые (highway={OSM_UNPAVED_HIGHWAY} и surface={OSM_UNPAVED_SURFACE})")
            src = self._osm_src_keys.get(self.v_osm_src.get(), "auto")
            if src == "pbf":
                where = ("файл " + Path(self._osm_pbf).name) if self._osm_pbf else \
                    "файл .pbf не выбран — нажмите «Файл .pbf…»"
            elif src == "auto":
                where = "Overpass, при неудаче — выгрузки BBBike (Москва), затем Geofabrik (ЦФО); нужен интернет" + \
                    (f"; первым — файл {Path(self._osm_pbf).name}" if self._osm_pbf else "")
            elif src == "overpass":
                where = "Overpass API, нужен интернет"
            else:
                where = OSM_SOURCES[src][0] + " — скачивается один раз в папку osm_cache рядом со скриптом"
            txt = ("тропы OSM в границах территории: " + "; ".join(parts) + ". Источник: " + where) if parts else \
                "отметьте, какие тропы брать из OSM: стихийные и/или грунтовые"
        else:
            txt = "стихийные тропы не используются — мероприятия с «Только стихийные тропы» = да не назначаются"
        self.l_inf.configure(text=txt)

    def _pick_osm_pbf(self):
        from tkinter import filedialog
        p = filedialog.askopenfilename(parent=self.root, title="Файл OpenStreetMap (.osm.pbf / .osm)",
                                       filetypes=PBF_TYPES)
        if p:
            self._osm_pbf = p
            if self._osm_src_keys.get(self.v_osm_src.get()) != "auto":
                self.v_osm_src.set(OSM_SOURCES["pbf"][0])
            self.v_inf_mode.set("auto")
            self._inf_changed()

    def _pick_informal(self):
        ps = self._ask("vec", True, "Стихийные тропы — линейный слой (можно несколько файлов)")
        if ps:
            self._inf_files = ps
            self.v_inf_mode.set("files")
        self._inf_changed()
        try:
            self._refresh()
        except Exception:
            pass

    # ------------------------------------------------------------------ выбор файлов
    def _ask(self, kind, multi, title):
        from tkinter import filedialog
        types = {"vec": VEC_TYPES, "ras": RAS_TYPES, "xls": XLS_TYPES, "vec_xls": VEC_XLS_TYPES}[kind]
        if multi:
            return list(filedialog.askopenfilenames(parent=self.root, title=title, filetypes=types))
        p = filedialog.askopenfilename(parent=self.root, title=title, filetypes=types)
        return [p] if p else []

    def _pick_xlsx(self):
        ps = self._ask("xls", False, "Таблица мероприятий (Excel)")
        if ps:
            self.cfg.measures_xlsx = ps[0]
            self._read_measures()
            self._refresh()

    def _pick_oozt(self):
        ps = self._ask("vec", False, "Слой территории: " + UNIT_MODES[self.v_mode.get()])
        if not ps:
            return
        self.cfg.oozt_path = ps[0]
        self.cfg.oozt_layer = None
        if not self.v_out.get():
            self.v_out.set(str(Path(ps[0]).parent / "oozt_recs_out"))
        self._setup_layer_combo()
        self._read_oozt_fields()

    def _pick_layer(self, spec):
        ps = self._ask(spec["kind"], spec["multi"], spec["title"] + (" — можно выбрать несколько" if spec["multi"] else ""))
        if not ps:
            return
        if spec["multi"]:
            cur = list(getattr(self.cfg, spec["key"]) or [])
            setattr(self.cfg, spec["key"], cur + [p for p in ps if p not in cur])
        else:
            setattr(self.cfg, spec["key"], ps[0])
        self._refresh()

    def _clear_layer(self, key):
        setattr(self.cfg, key, [] if isinstance(getattr(self.cfg, key), list) else None)
        self._refresh()

    def _pick_out(self):
        from tkinter import filedialog
        p = filedialog.askdirectory(parent=self.root, title="Папка для результатов")
        if p:
            self.v_out.set(p)

    # ------------------------------------------------------------------ чтение входов
    def _read_measures(self):
        try:
            self.measures = read_measures(self.cfg.measures_xlsx)
            self.v_xlsx.set(self.cfg.measures_xlsx)
            n_on = sum(m["on"] for m in self.measures)
            self.l_xlsx.configure(foreground=self.C_OK,
                                  text=f"✔ {len(self.measures)} мероприятий, включено (Вкл/выкл = 1): {n_on}")
        except Exception as e:
            self.measures = []
            self.v_xlsx.set(self.cfg.measures_xlsx)
            self.l_xlsx.configure(foreground=self.C_WARN, text=f"! таблица не прочитана: {e}")

    def _setup_layer_combo(self):
        self.v_oozt.set(self.cfg.oozt_path)
        p = self.cfg.oozt_path or ""
        if p.lower().endswith(".gpkg") and Path(p).exists():
            layers = [l[0] for l in pyogrio.list_layers(p)]
            self.cb_layer.configure(values=layers, state="readonly")
            self.v_layer.set(self.cfg.oozt_layer if self.cfg.oozt_layer in layers else layers[0])
        else:
            self.cb_layer.configure(values=[], state="disabled")
            self.v_layer.set("")

    def _read_oozt_fields(self):
        self.attr_df = None
        cols = []
        if self.cfg.oozt_path and Path(self.cfg.oozt_path).exists():
            try:
                cols, self.attr_df = _vector_fields(self.cfg.oozt_path, self.v_layer.get() or None)
            except Exception as e:
                self.l_types.configure(foreground=self.C_WARN, text=f"! слой не прочитан: {e}")
        self.columns = cols
        self._build_fields(guess=True)
        self._refresh()

    def _read_kvnet_fields(self):
        self.kvnet_columns = []
        p = self.cfg.kvnet_path or ""
        if p and Path(p).exists():
            try:
                self.kvnet_columns, _ = _vector_fields(p, self.v_kvnet_layer.get() or None)
            except Exception:
                self.kvnet_columns = []
        cols = self.kvnet_columns
        self.cb_kvnet_field.configure(values=cols)
        if cols and self.v_kvnet_field.get() not in cols:
            cur = self.cfg.kvnet_field
            if cur not in cols:
                cur = next((c for c in KVNET_KV_CANDIDATES if c in cols), None) or next(
                    (c for c in cols if c.lower() in [x.lower() for x in KVNET_KV_CANDIDATES]), None)
            self.v_kvnet_field.set(cur or "")
        self.cb_kvnet_oozt.configure(values=[NONE_FIELD] + cols)
        if cols and self.v_kvnet_oozt.get() not in [NONE_FIELD] + cols:
            cur = self.cfg.kvnet_oozt_field
            if cur not in cols:
                cur = next((c for c in ("OOZT", "ООЗТ", "Name_OOZT") if c in cols), NONE_FIELD)
            self.v_kvnet_oozt.set(cur)
        self._refresh()

    def _pick_kvnet(self):
        ps = self._ask("vec", False, "Вектор квартальной сети (кварталы)")
        if not ps:
            return
        self.cfg.kvnet_path = ps[0]
        self.cfg.kvnet_layer = None
        self.v_kvnet.set(ps[0])
        p = ps[0]
        if p.lower().endswith(".gpkg"):
            layers = [l[0] for l in pyogrio.list_layers(p)]
            self.cb_kvnet_layer.configure(values=layers, state="readonly")
            self.v_kvnet_layer.set(layers[0])
        else:
            self.cb_kvnet_layer.configure(values=[], state="disabled")
            self.v_kvnet_layer.set("")
        self._read_kvnet_fields()

    def _build_fields(self, guess=False):
        ttk = self.ttk
        for w in self.fields_frame.winfo_children():
            w.destroy()
        mode = self.v_mode.get()
        cols = getattr(self, "columns", [])
        prev = {k: v.get() for k, v in self.field_vars.items()}
        self.field_vars = {}
        ttk.Label(self.fields_frame, text="Поля слоя:", font=self.f_bold).grid(row=0, column=0, sticky="w")
        for i, (key, label, req, cands) in enumerate(FIELD_SPECS[mode], start=1):
            v = self.tk.StringVar()
            values = cols if req else [NONE_FIELD] + cols
            cur = (self.cfg.fields or {}).get(key) or prev.get(key)
            if cur not in cols and guess:
                cur = next((c for c in cands if c in cols), None) or next(
                    (c for c in cols if c.lower() in [x.lower() for x in cands]), None)
            v.set(cur if cur in cols else ("" if req else NONE_FIELD))
            ttk.Label(self.fields_frame, text=label + (" *" if req else "")).grid(row=i, column=0, sticky="w", pady=1)
            cb = ttk.Combobox(self.fields_frame, textvariable=v, values=values, state="readonly", width=26)
            cb.grid(row=i, column=1, sticky="w", padx=6, pady=1)
            cb.bind("<<ComboboxSelected>>", lambda e: self._refresh())
            self.field_vars[key] = v
        if mode == "eco":
            self.eco_frame.pack(fill="x", pady=(6, 0), before=self.l_types)
            self.cb_group.configure(values=[NONE_FIELD] + cols)
            if self.v_group.get() not in [NONE_FIELD] + cols:
                self.v_group.set(NONE_FIELD)
        else:
            self.eco_frame.pack_forget()
        if mode == "kv_cut":
            self.kvnet_frame.pack(fill="x", pady=(6, 0), before=self.l_types)
        else:
            self.kvnet_frame.pack_forget()

    def _on_mode(self):
        self.cfg.fields = {}
        self._build_fields(guess=True)
        self._refresh()

    # ------------------------------------------------------------------ сводка
    def _collect(self) -> Config:
        c = self.cfg
        c.unit_mode = self.v_mode.get()
        c.oozt_layer = self.v_layer.get() or None
        c.fields = {k: v.get() for k, v in self.field_vars.items() if v.get() and v.get() != NONE_FIELD}
        c.oozt_name = self.v_oozt_name.get().strip()
        g = self.v_group.get()
        c.report_group_field = g if (c.unit_mode == "eco" and g and g != NONE_FIELD) else None
        c.kvnet_layer = self.v_kvnet_layer.get() or None
        c.kvnet_field = self.v_kvnet_field.get().strip()
        ko = self.v_kvnet_oozt.get()
        c.kvnet_oozt_field = ko if (c.unit_mode == "kv_cut" and ko and ko != NONE_FIELD) else ""
        try:
            c.sheet_size_m = max(50.0, float(str(self.v_sheet.get()).replace(",", ".")))
        except ValueError:
            c.sheet_size_m = 1000.0
        c.out_dir = self.v_out.get().strip()
        c.report_kv = [x.strip() for x in re.split(r"[,;\s]+", self.v_kv.get()) if x.strip()]
        c.report_kv = [int(x) if x.isdigit() else x for x in c.report_kv]
        c.basemap = bool(self.v_basemap.get())
        c.split_multipart = bool(self.v_split.get())
        _im = self.v_inf_mode.get()
        c.informal_trails = (["auto"] if _im == "auto" else
                             list(self._inf_files) if _im == "files" else [])
        c.osm_trail_kinds = [k for k, v in (("informal", self.v_osm_inf), ("unpaved", self.v_osm_unp))
                             if v.get()] or ["informal"]
        c.osm_source = self._osm_src_keys.get(self.v_osm_src.get(), "auto")
        c.osm_pbf = self._osm_pbf or ""
        try:
            c.min_share_pct = max(0.0, float(self.v_minshare.get().replace(",", ".") or 0))
        except ValueError:
            c.min_share_pct = 25.0
        c.report_formats = [k for k, v in self.v_formats.items() if v.get()]
        return c

    def _refresh(self):
        cfg = self._collect()
        L = unit_labels(cfg)
        self.l_kv.configure(text=("Кварталы для отчёта:" if cfg.unit_mode != "eco" else
                                  "Листы для отчёта:" if not cfg.report_group_field else f"{L['group']}: для отчёта"))
        # типы полигонов
        self.type_counts = None
        tf = cfg.fields.get("TYPE")
        if self.attr_df is not None and tf in self.attr_df.columns:
            vc = self.attr_df[tf].map(_remojibake).astype(str).str.strip().value_counts()
            self.type_counts = vc.to_dict()
            known = (TYPE_FOREST + TYPE_MEADOW + TYPE_WATER + TYPE_NEARWATER)
            parts, unk = [], []
            for t, n in vc.items():
                (parts if any(k in t.lower() for k in known) else unk).append(f"{t} — {n}")
            txt = f"Полигонов: {len(self.attr_df)}. Типы, распознаваемые правилами: " + (", ".join(parts) or "нет")
            if unk:
                txt += ". Не используются правилами по типу: " + ", ".join(unk[:8])
            idf = cfg.fields.get("ID")
            if cfg.unit_mode == "eco" and idf in self.attr_df.columns:
                dup = int(self.attr_df[idf].duplicated().sum())
                txt += (f". ID уникальны." if dup == 0 else f". ! В поле ID {dup} повторов — ID должен быть уникальным.")
            self.l_types.configure(foreground=self.C_WARN if (not parts or "!" in txt) else self.C_MUTED, text=txt)
        elif self.attr_df is not None:
            self.l_types.configure(foreground=self.C_WARN, text="! выберите поле типа полигона")
        # тематические слои
        need_by = {k: [] for k in self.layer_rows}
        statuses = []
        for m in self.measures:
            s = measure_status(m, cfg, self.type_counts)
            statuses.append(s)
            if s["rule"]:
                for d in RULES[s["rule"]]["data"] + RULES[s["rule"]]["opt"]:
                    for k in DATA_KEYS[d]:
                        if k in need_by:
                            need_by[k].append((m.get("code") or "?", m["on"]))
        self._sp_refresh_status(statuses)
        for key, row in self.layer_rows.items():
            val = getattr(cfg, key)
            files = val if isinstance(val, list) else ([val] if val else [])
            codes_on = sorted(dict.fromkeys(c for c, on in need_by[key] if on), key=_code_key)
            codes_off = sorted(dict.fromkeys(c for c, on in need_by[key] if not on and c not in codes_on), key=_code_key)
            if not self.measures:
                utxt = "Нужен для мероприятий: выберите таблицу мероприятий"
            else:
                utxt = "Нужен для включённых: " + (", ".join(codes_on) if codes_on else "—")
                if codes_off:
                    utxt += ";  выключенных: " + ", ".join(codes_off)
            row["used"].configure(text=utxt, foreground="#202124" if codes_on else self.C_MUTED)
            if files:
                names = [Path(f.split("|")[0]).name for f in files]
                row["files"].configure(foreground=self.C_OK,
                                       text="Выбрано" + (f" ({len(files)})" if len(files) > 1 else "") + ": " +
                                       ", ".join(names[:5]) + (" …" if len(names) > 5 else ""))
                row["status"].configure(text="✔", foreground=self.C_OK)
            else:
                row["files"].configure(foreground=self.C_WARN if codes_on else self.C_OFF, text="Не выбран")
                if codes_on:
                    row["status"].configure(text="!", foreground=self.C_WARN)
                else:
                    row["status"].configure(text="○", foreground=self.C_OFF)
        # таблица мероприятий
        self.tree.delete(*self.tree.get_children())
        icon = {"ok": "✔", "missing": "!", "off": "—", "na": "—"}
        for i, (m, s) in enumerate(zip(self.measures, statuses)):
            self.tree.insert("", "end", iid=str(i), tags=(s["state"],),
                             values=(m.get("code", ""), m["name"], f"{icon[s['state']]} {s['text']}"))
        self._statuses = statuses
        if self.measures:
            n_on = sum(m["on"] for m in self.measures)
            n_ok = sum(s["state"] == "ok" for s in statuses)
            n_miss = sum(s["state"] == "missing" for s in statuses)
            self.l_sum.configure(text=f"Будет рассчитано {n_ok} из {n_on} включённых мероприятий"
                                      + (f"; не хватает данных: {n_miss}" if n_miss else ""),
                                 foreground=self.C_OK if n_miss == 0 else self.C_WARN)

    def _on_select(self, _e=None):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        m, s = self.measures[i], self._statuses[i]
        spec = s["spec"]
        lines = [f"{m.get('code', '')} {m['name']}".strip(),
                 f"Пункт Приказа: {m.get('prikaz') or '—'};  тип экосистемы: {m.get('eco') or '—'};  "
                 f"Вкл/выкл: {'1' if m['on'] else '0'}",
                 f"Метод выявления (Excel): {m.get('method') or '—'}", ""]
        if spec["rule"]:
            how = ("из столбца «Правило назначения»" if spec["origin"] == "таблица" else
                   "определено по тексту «Метода выявления» (столбец «Правило назначения» пуст)")
            lines.append(f"Правило: {RULES[spec['rule']]['title']} — {how}")
            lines.append("Параметры (в скобках — откуда взято значение):")
            lines += ["  • " + x for x in spec_params_lines(spec)]
            lines += ["", "Как рассчитывается: " + describe_spec(spec, full=True), ""]
            lines.append(f"Нужные данные: {s['need']}")
            if RULES[spec["rule"]]["opt"]:
                lines.append("Дополнительно используется: " +
                             ", ".join(LAYER_TITLE[k] for d in RULES[spec["rule"]]["opt"] for k in DATA_KEYS[d]))
            if "species" in RULES[spec["rule"]]["data"]:
                ls = matching_species_layers(self.cfg, spec["p"],
                                             ["points"] if spec["rule"] == "species_density" else None)
                lines.append("Подходящие слои видов: " + ("; ".join(species_layer_label(l) for l in ls) or "нет"))
        else:
            lines.append("Правило: не задано — " + spec.get("reason", ""))
        if spec.get("errors"):
            lines.append("Ошибки в столбцах правил: " + "; ".join(spec["errors"]))
        lines.append(f"Статус: {s['text']}")
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", "\n".join(lines))
        self.detail.configure(state="disabled")

    # ------------------------------------------------------------------ конфигурация
    def _load_from_cfg(self, cfg: Config):
        self.cfg = cfg
        self.v_mode.set(cfg.unit_mode if cfg.unit_mode in UNIT_MODES else "kv_vyd")
        self.v_oozt_name.set(cfg.oozt_name or "")
        self.v_group.set(cfg.report_group_field or NONE_FIELD)
        self.v_sheet.set(str(int(cfg.sheet_size_m)) if cfg.sheet_size_m else "1000")
        self.v_out.set(cfg.out_dir or "")
        self.v_kv.set(", ".join(map(str, cfg.report_kv or [])))
        self.v_basemap.set(bool(cfg.basemap))
        self.v_split.set(bool(getattr(cfg, "split_multipart", True)))
        _inf = [str(x).strip() for x in (cfg.informal_trails or []) if str(x).strip()]
        self._inf_files = [x for x in _inf if x.lower() not in ("auto", "авто", "osm")]
        self.v_inf_mode.set("files" if self._inf_files else "auto" if _inf else "none")
        _k = getattr(cfg, "osm_trail_kinds", None) or ["informal"]
        self.v_osm_inf.set("informal" in _k)
        self.v_osm_unp.set("unpaved" in _k)
        _src = str(getattr(cfg, "osm_source", "auto") or "auto").lower()
        self.v_osm_src.set(OSM_SOURCES.get(_src, OSM_SOURCES["auto"])[0])
        self._osm_pbf = str(getattr(cfg, "osm_pbf", "") or "")
        self._inf_changed()
        self.v_minshare.set(f"{float(getattr(cfg, 'min_share_pct', 25.0)):g}")
        for k, v in self.v_formats.items():
            v.set(k in (cfg.report_formats or []))
        self.field_vars = {}
        self._sp_render()
        if cfg.measures_xlsx:
            self._read_measures()
        self._setup_layer_combo()
        if cfg.oozt_path and Path(cfg.oozt_path).exists():
            try:
                self.columns, self.attr_df = _vector_fields(cfg.oozt_path, self.v_layer.get() or None)
            except Exception:
                self.columns, self.attr_df = [], None
        self.v_kvnet.set(cfg.kvnet_path or "")
        self.v_kvnet_field.set(cfg.kvnet_field or "")
        self.v_kvnet_oozt.set(cfg.kvnet_oozt_field or NONE_FIELD)
        if cfg.kvnet_path and str(cfg.kvnet_path).lower().endswith(".gpkg") and Path(cfg.kvnet_path).exists():
            layers = [l[0] for l in pyogrio.list_layers(cfg.kvnet_path)]
            self.cb_kvnet_layer.configure(values=layers, state="readonly")
            self.v_kvnet_layer.set(cfg.kvnet_layer if cfg.kvnet_layer in layers else layers[0])
        else:
            self.cb_kvnet_layer.configure(values=[], state="disabled")
            self.v_kvnet_layer.set("")
        self._build_fields(guess=True)
        if cfg.unit_mode == "kv_cut" and cfg.kvnet_path and Path(cfg.kvnet_path).exists():
            try:
                self.kvnet_columns, _ = _vector_fields(cfg.kvnet_path, self.v_kvnet_layer.get() or None)
                self.cb_kvnet_field.configure(values=self.kvnet_columns)
                self.cb_kvnet_oozt.configure(values=[NONE_FIELD] + self.kvnet_columns)
            except Exception:
                pass
        self._refresh()

    def _load_cfg_file(self):
        from tkinter import filedialog, messagebox
        p = filedialog.askopenfilename(parent=self.root, title="Конфигурация (json)", filetypes=JSON_TYPES)
        if not p:
            return
        try:
            self._load_from_cfg(Config.load(p))
        except Exception as e:
            messagebox.showerror("Конфигурация", f"Не удалось прочитать {p}:\n{e}", parent=self.root)

    def _save_cfg_file(self):
        from tkinter import filedialog
        p = filedialog.asksaveasfilename(parent=self.root, title="Сохранить конфигурацию", defaultextension=".json",
                                         filetypes=JSON_TYPES, initialfile="config_local.json")
        if p:
            self._collect().save(p)

    def _validate(self) -> list[str]:
        c = self._collect()
        err = []
        if not c.measures_xlsx or not self.measures:
            err.append("не выбрана или не прочитана таблица мероприятий")
        if not c.oozt_path:
            err.append("не выбран слой территории")
        for key, label, req, _ in FIELD_SPECS[c.unit_mode]:
            if req and not c.fields.get(key):
                err.append(f"не задано поле «{label}»")
        if c.unit_mode == "eco" and not c.fields.get("OOZT") and not c.oozt_name:
            err.append("укажите поле ООЗТ или название ООЗТ")
        if c.unit_mode == "kv_cut":
            if not c.kvnet_path:
                err.append("не выбран вектор квартальной сети")
            elif not c.kvnet_field:
                err.append("не задано поле номера квартала в квартальной сети")
            elif not (c.fields.get("OOZT") or c.kvnet_oozt_field or c.oozt_name):
                err.append("укажите поле ООЗТ (в экосистемах или в сети) или название ООЗТ")
        if c.unit_mode == "eco" and self.attr_df is not None and c.fields.get("ID") in self.attr_df.columns:
            if self.attr_df[c.fields["ID"]].duplicated().any():
                err.append("в поле ID экосистемы есть повторы")
        if not c.out_dir:
            err.append("не указана папка результатов")
        return err

    def _run(self):
        from tkinter import messagebox
        err = self._validate()
        if err:
            messagebox.showwarning("Не хватает данных", "• " + "\n• ".join(err), parent=self.root)
            return
        miss = [f"{m.get('code', '')} {m['name']}: {s['text']}" for m, s in zip(self.measures, self._statuses)
                if s["state"] == "missing"]
        if miss and not messagebox.askyesno(
                "Часть мероприятий не будет рассчитана",
                "Для этих включённых мероприятий не хватает данных:\n\n" + "\n".join(miss[:15]) +
                ("\n…" if len(miss) > 15 else "") + "\n\nЗапустить расчёт без них?", parent=self.root):
            return
        self.result = self._collect()
        self.root.destroy()

    def _cancel(self):
        self.result = None
        self.root.destroy()

    def show(self) -> Config | None:
        self.root.mainloop()
        return self.result


def _code_key(code: str):
    """Сортировка кодов вида «III.5»: римская часть, затем номер."""
    rom = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6}
    mt = re.match(r"^([IVX]+)\.(\d+)", str(code))
    return (rom.get(mt.group(1), 99), int(mt.group(2))) if mt else (99, 0)


def gui_collect_config(cfg: Config | None = None) -> Config:
    res = SetupWindow(cfg).show()
    if res is None:
        print("Отменено пользователем")
        sys.exit(0)
    return res


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
    def _read_with_enc(pth):
        try:
            return gpd.read_file(pth, **kw), pth
        except UnicodeDecodeError:
            return gpd.read_file(pth, encoding="cp1251", **kw), pth

    _tmp_shp_dir = None
    try:
        try:
            g, read_path = _read_with_enc(path)
        except Exception as e:
            p_obj = Path(path)
            if p_obj.suffix.lower() == ".shp" and (".shx" in str(e).lower() or "for writing" in str(e).lower()):
                import shutil
                _tmp_shp_dir = tempfile.TemporaryDirectory()
                td = Path(_tmp_shp_dir.name)
                for ext in (".shp", ".shx", ".dbf", ".prj", ".cpg", ".qpj"):
                    cand = p_obj.with_suffix(ext)
                    if cand.exists():
                        shutil.copy2(cand, td / cand.name)
                g, read_path = _read_with_enc(str(td / p_obj.name))
            else:
                raise
        if g.empty:
            return g
        g = fix_text_encoding(g, read_path, kw)
    finally:
        if _tmp_shp_dir is not None:
            _tmp_shp_dir.cleanup()
    g = g[~g.geometry.isna() & ~g.geometry.is_empty].copy()
    crs = fix_crs(g.crs, g.total_bounds, Path(path).name)
    g = g.set_crs(crs, allow_override=True)
    if g.crs.to_epsg() != WORK_EPSG:
        g = g.to_crs(WORK_EPSG)
    g = repair_geometries(g)          # make_valid: защита от TopologyException в overlay/union
    if g is None:
        return gpd.GeoDataFrame(columns=["__src"], geometry=[], crs=WORK_EPSG)
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
    g = repair_geometries(g)          # make_valid: защита от TopologyException
    if g is None:
        return None
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
def _nat_key(v):
    """Естественный порядок: числа — по значению (1110 < 5110 < 44504), текст — по фрагментам
    («кв2» < «кв10», «5-12» < «5-110»)."""
    s_ = "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip()
    try:
        return (0, float(s_.replace(",", ".")), ())
    except ValueError:
        pass
    parts = tuple((0, int(p), "") if p.isdigit() else (1, 0, p.lower()) for p in re.split(r"(\d+)", s_) if p)
    return (1, 0.0, parts)


def _nat_rank(values) -> np.ndarray:
    """Ранги значений в естественном порядке (для сортировки смешанных / текстовых полей)."""
    vals = ["" if pd.isna(v) else str(v).strip() for v in values]
    order = {v: i for i, v in enumerate(sorted(set(vals), key=_nat_key))}
    return np.array([order[v] for v in vals], dtype=np.int64)


def _nat_sort_key(*cols):
    """key= для DataFrame.sort_values: указанные колонки — в естественном порядке."""
    def key(sr: pd.Series):
        return pd.Series(_nat_rank(sr.values), index=sr.index) if sr.name in cols else sr
    return key


def _maybe_int(sr: pd.Series) -> pd.Series:
    """Номера кварталов/выделов/ID → Int64, если все значения целые; иначе — текст как есть."""
    try:
        num = pd.to_numeric(sr)
        if (num.dropna() % 1 == 0).all():
            return num.astype("Int64")
    except (ValueError, TypeError):
        pass
    return sr.astype(str).str.strip()


MIN_PART_HA = 0.01      # контур многочастного объекта меньше этой площади — не отдельная единица


def _split_multipart(g: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Многочастный полигон → отдельные контуры. Мелкие части (< MIN_PART_HA) остаются при крупнейшем контуре."""
    from shapely.geometry import MultiPolygon
    rows = []
    for i, (idx, r) in enumerate(g.iterrows()):
        geom = r.geometry
        parts = list(geom.geoms) if geom.geom_type == "MultiPolygon" else [geom]
        parts.sort(key=lambda x: -x.area)
        big = [p for p in parts if p.area / 1e4 >= MIN_PART_HA] or parts[:1]
        small = [p for p in parts if p not in big]
        if small:
            # мелкие части — к ближайшему крупному контуру
            groups = [[b] for b in big]
            for sp in small:
                k = min(range(len(big)), key=lambda j: big[j].distance(sp))
                groups[k].append(sp)
            big = [grp[0] if len(grp) == 1 else MultiPolygon(grp) for grp in groups]
        for k, p in enumerate(big, 1):
            rr = r.copy()
            rr["geometry"] = p
            rr["_part"], rr["_nparts"], rr["_src_row"] = k, len(big), i
            rows.append(rr)
    return gpd.GeoDataFrame(rows, geometry="geometry", crs=g.crs).reset_index(drop=True)


def _part_ids(ids: pd.Series, g: gpd.GeoDataFrame) -> pd.Series:
    """ID контура: исходный ID, для частей многочастного объекта — «ID-1», «ID-2», …"""
    multi = g["_nparts"] > 1
    out = ids.copy().astype(object)
    out[multi] = ids[multi].astype(str) + "-" + g.loc[multi, "_part"].astype(str)
    return out


# =====================================================================================
#        СПРАВОЧНИКИ: КРАСНАЯ КНИГА МОСКВЫ, ИНВАЗИВНЫЕ ВИДЫ, СТИХИЙНЫЕ ТРОПЫ OSM
# =====================================================================================
try:
    _REF_DIR = Path(__file__).resolve().parent / "refdata"
except NameError:                        # Jupyter: __file__ не определён
    _REF_DIR = Path.cwd() / "refdata"
_REF_CACHE: dict = {}
# справочники Excel по умолчанию (поле Config → файл в refdata/)
REF_DEFAULTS = {"water_ref": "Справочник_водоёмов_и_болот_ООЗТ_Москвы.xlsx",
                "flowering_ref": "okhranyaemye_rasteniya_tsvetenie.xlsx"}


def _ref_path(key: str) -> Path | None:
    """Путь к справочнику: из текущей конфигурации, иначе — refdata/ по умолчанию."""
    v = getattr(_CUR_CFG, key, None) if _CUR_CFG is not None else None
    if v and Path(str(v)).exists():
        return Path(str(v))
    d = _REF_DIR / REF_DEFAULTS[key]
    return d if (v is None or not str(v).strip()) and d.exists() else None


def _norm_latin(s) -> str:
    """«Pulmonaria obscura Duch. [P. officinalis]» → «pulmonaria obscura»: первые два слова, строчными."""
    t = re.sub(r"\[.*?\]|\(.*?\)", " ", str(s or "")).lower()
    parts = [p for p in re.split(r"[^a-zà-ÿ-]+", t) if p]
    return " ".join(parts[:2]) if len(parts) >= 2 else (parts[0] if parts else "")


def _load_ref(name: str, cols: list) -> dict:
    """CSV из refdata/ → {нормализованное латинское название: {столбец: значение}}."""
    if name in _REF_CACHE:
        return _REF_CACHE[name]
    ref = {}
    p = _REF_DIR / name
    try:
        df = pd.read_csv(p, encoding="utf-8-sig")
        for _, row in df.iterrows():
            key = _norm_latin(row.get("latin"))
            if key:
                ref[key] = {c: row.get(c) for c in cols}
        log.info("Справочник %s: %d позиций", p.name, len(ref))
    except Exception as e:
        log.warning("Справочник %s не загружен (%s): определение категории Красной книги недоступно",
                    p.name, e)
    _REF_CACHE[name] = ref
    return ref


def kk_ref() -> dict:
    """Красная книга Москвы, 3-е издание (ПП Москвы № 745-ПП от 02.07.2019, Приложение 2):
    латинское название → категория (0–5 или «О») и русское название."""
    ref = _load_ref("kk_moscow.csv", ["cat", "ru"])
    for v in ref.values():                     # категория: int, «О» или NaN
        c = v["cat"]
        try:
            v["cat"] = int(float(c))
        except (TypeError, ValueError):
            v["cat"] = str(c).strip() if str(c).strip() else np.nan
    return ref


def _latin_keys(s) -> list[str]:
    """Ключи латинского названия для сравнения: полное (до 3 слов, без авторов в скобках), затем 2 слова.
    «Dactylorhiza majalis baltica» → ['dactylorhiza majalis baltica', 'dactylorhiza majalis']."""
    t = re.sub(r"\[.*?\]|\(.*?\)", " ", str(s or ""))
    t = re.sub(r"(?i)\b(subsp|ssp|var|aggr?)\b\.?|\bs\.\s*(l|str)\.", " ", t)
    raw = [p for p in re.split(r"[^A-Za-zÀ-ÿ-]+", t) if p]
    parts = [p.lower() for p in raw]
    keys = []
    # третье слово — подвид (строчное), а не автор («Duch.», «L.»)
    if len(raw) >= 3 and raw[2][:1].islower() and len(raw[2]) > 2:
        keys.append(" ".join(parts[:3]))
    if len(parts) >= 2:
        keys.append(" ".join(parts[:2]))
    elif parts:
        keys.append(parts[0])
    return list(dict.fromkeys(keys))


# синонимы латинских названий (по умолчанию; дополняются справочником цветения, столбцы C/D)
SYNONYMS_DEFAULT = [
    ("Anthoxanthum nitens", "Hierochloe odorata"), ("Bistorta officinalis", "Bistorta major"),
    ("Silene flos-cuculi", "Coccyganthe flos-cuculi"), ("Myosotis scorpioides", "Myosotis palustris"),
    ("Viscaria vulgaris", "Steris viscaria"), ("Bromus benekenii", "Bromopsis benekenii"),
    ("Dactylorhiza majalis baltica", "Dactylorhiza baltica"), ("Cota tinctoria", "Anthemis tinctoria"),
    ("Neottia ovata", "Listera ovata"), ("Spinulum annotinum", "Lycopodium annotinum"),
]


def synonym_groups() -> dict:
    """Ключ латинского названия → множество равноправных ключей (синонимы). Источники: справочник цветения
    (столбцы «Латинское название (в слое)» и «… по справочнику ККМ»), refdata/species_synonyms.csv и
    встроенный список."""
    ck = ("_syn", str(_ref_path("flowering_ref")))
    if ck in _REF_CACHE:
        return _REF_CACHE[ck]
    pairs = list(SYNONYMS_DEFAULT)
    p = _REF_DIR / "species_synonyms.csv"
    if p.exists():
        try:
            df = pd.read_csv(p, encoding="utf-8-sig")
            c1, c2 = df.columns[:2]
            pairs += [(a, b) for a, b in zip(df[c1], df[c2]) if str(a).strip() and str(b).strip()]
        except Exception as e:
            log.warning("Синонимы %s не прочитаны: %s", p.name, e)
    pairs += flowering_ref().get("pairs", [])
    groups: dict = {}
    for a, b in pairs:
        ka, kb = _latin_keys(a), _latin_keys(b)
        if not ka or not kb or ka[0] == kb[0]:
            continue
        g = set(groups.get(ka[0], {ka[0]})) | set(groups.get(kb[0], {kb[0]}))
        for k in list(g):
            g |= groups.get(k, set())
        for k in g:
            groups[k] = g
    _REF_CACHE[ck] = groups
    return groups


def _lookup_latin(species, ref: dict):
    """Поиск вида в словаре по ключам названия; затем — по синонимам. Возвращает (значение, через_синоним)."""
    keys = _latin_keys(species)
    for k in keys:
        if k in ref:
            return ref[k], False
    syn = synonym_groups()
    for k in keys:
        for k2 in sorted(syn.get(k, ())):
            if k2 != k and k2 in ref:
                return ref[k2], True
            k2b = " ".join(k2.split()[:2])
            if k2b != k and k2b in ref:
                return ref[k2b], True
    return None, False


def kk_category(species):
    """Категория Красной книги Москвы для вида: 0–5, «О» или NaN (вида нет в КК). Синонимы равноправны."""
    r, _ = _lookup_latin(species, kk_ref())
    return np.nan if r is None else r["cat"]


def kk_via_synonym(species) -> bool:
    """Вид найден в Красной книге Москвы только через синоним."""
    r, syn = _lookup_latin(species, kk_ref())
    return r is not None and syn


# ---- месяцы (сезонность) ----------------------------------------------------------------
ROMAN_MONTHS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
_ROMAN_IDX = {r: i + 1 for i, r in enumerate(ROMAN_MONTHS)}


def months_text(months) -> str:
    """{5, 6, 7} → «V–VII»; {4, 5, 8} → «IV–V, VIII»."""
    ms = sorted({int(m) for m in months if 1 <= int(m) <= 12})
    out, i = [], 0
    while i < len(ms):
        j = i
        while j + 1 < len(ms) and ms[j + 1] == ms[j] + 1:
            j += 1
        out.append(ROMAN_MONTHS[ms[i] - 1] if i == j else f"{ROMAN_MONTHS[ms[i] - 1]}–{ROMAN_MONTHS[ms[j] - 1]}")
        i = j + 1
    return ", ".join(out)


def season_months(text):
    """«IV–V, VIII» → {4, 5, 8}; текст не из римских месяцев → None. Диапазон через год (XII–III) — по кругу."""
    t = str(text or "").strip()
    if not t:
        return set()
    out = set()
    for tok in re.split(r"[,;]\s*", t):
        tok = tok.strip().upper().replace("—", "-").replace("–", "-")
        if not tok:
            continue
        m = re.fullmatch(r"([IVX]+)\s*(?:-\s*([IVX]+))?", tok)
        if not m or m.group(1) not in _ROMAN_IDX or (m.group(2) and m.group(2) not in _ROMAN_IDX):
            return None
        a = _ROMAN_IDX[m.group(1)]
        b = _ROMAN_IDX[m.group(2)] if m.group(2) else a
        k = a
        while True:
            out.add(k)
            if k == b:
                break
            k = k % 12 + 1
    return out


def merge_seasons(values) -> str:
    """Сезонность нескольких записей: римские месяцы — объединение («V–VII» + «VIII» → «V–VIII»),
    иначе — уникальные значения через «; »."""
    vals = [str(v).strip() for v in values
            if v is not None and not (isinstance(v, float) and np.isnan(v)) and str(v).strip()
            and str(v).strip().lower() != "nan"]
    vals = list(dict.fromkeys(vals))
    if len(vals) <= 1:
        return vals[0] if vals else ""
    sets = [season_months(v) for v in vals]
    if all(x is not None for x in sets) and not any(12 in x and 1 in x for x in sets):
        return months_text(set().union(*sets))
    return "; ".join(vals)


# ---- справочник цветения охраняемых растений ---------------------------------------------
def flowering_ref() -> dict:
    """Лист «Справочник»: {"months": ключ латинского названия → множество месяцев, "pairs": синонимы,
    "kk": ключ → категория ККМ}. Оба латинских названия (в слое и по ККМ) — ключи."""
    p = _ref_path("flowering_ref")
    ck = ("_flow", str(p))
    if ck in _REF_CACHE:
        return _REF_CACHE[ck]
    res = {"months": {}, "pairs": [], "kk": {}, "path": None}
    _REF_CACHE[ck] = res
    if p is None:
        log.warning("Справочник цветения не найден — сроки III.4 по цветению и синонимы из него недоступны")
        return res
    try:
        import openpyxl
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = next((w for w in wb.worksheets if _norm_txt(w.title).startswith("справочник")), wb.worksheets[0])
        rows = list(ws.iter_rows(values_only=True))
        h_i = next(i for i, r in enumerate(rows[:10]) if any(_norm_txt(c).startswith("латинск") for c in r))
        hdr = [_norm_txt(c) for c in rows[h_i]]
        lat = [i for i, h in enumerate(hdr) if h.startswith("латинск")]
        c_layer = next((i for i in lat if "в слое" in hdr[i]), lat[0])
        c_kk = next((i for i in lat if i != c_layer), None)
        c_cat = next((i for i, h in enumerate(hdr) if h.startswith("категория")), None)
        mcols = {i: _ROMAN_IDX[str(rows[h_i][i]).strip().upper()] for i in range(len(hdr))
                 if str(rows[h_i][i] or "").strip().upper() in _ROMAN_IDX}
        n = 0
        for r in rows[h_i + 1:]:
            if not r or c_layer >= len(r) or not str(r[c_layer] or "").strip():
                continue
            names = [str(r[c_layer]).strip()] + ([str(r[c_kk]).strip()] if c_kk is not None and r[c_kk] else [])
            ms = {m for i, m in mcols.items() if i < len(r) and str(r[i] or "").strip() not in ("", "-", "—")}
            for nm in names:
                for k in _latin_keys(nm):
                    res["months"].setdefault(k, set()).update(ms)
                    if c_cat is not None and r[c_cat] not in (None, ""):
                        res["kk"].setdefault(k, str(r[c_cat]).strip())
            if len(names) == 2 and _latin_keys(names[0])[:1] != _latin_keys(names[1])[:1]:
                res["pairs"].append((names[0], names[1]))
            n += 1
        res["path"] = p
        log.info("Справочник цветения %s: %d видов, синонимов — %d", p.name, n, len(res["pairs"]))
    except Exception as e:
        log.warning("Справочник цветения %s не прочитан: %s", p.name, e)
    return res


def flowering_months(species) -> set:
    """Месяцы цветения вида по справочнику (с учётом синонимов); пусто — нет в справочнику."""
    r, _ = _lookup_latin(species, flowering_ref()["months"])
    return set(r or ())


# ---- справочник водоёмов и болот ООЗТ ------------------------------------------------------
WATER_REF_MARKS = {"1", "1.0", "да", "+", "●", "•", "yes", "true", "x", "х"}
WATER_REF_SNAP_M = 30.0                # допуск привязки точки к ближайшему контуру, м
WATER_REF_TERLETSKY_MAX_KV = 3         # в Терлецком массиве 3 квартала


def _wr_val(v):
    """Значение ячейки справочника: ошибки Excel (#REF! и т. п.) и пустые строки → None."""
    if v is None:
        return None
    if isinstance(v, str):
        t = v.strip()
        if not t or t.startswith("#") or t.startswith("="):
            return None
        return t
    if isinstance(v, float) and np.isnan(v):
        return None
    return v


def _wr_mark(v) -> bool:
    v = _wr_val(v)
    if v is None:
        return False
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v) == 1.0
    return _norm_txt(v) in WATER_REF_MARKS


def read_water_ref(path) -> pd.DataFrame:
    """Справочник водоёмов и болот ООЗТ → таблица объектов с отметками мероприятий.
    Читаются листы «Битца и Измайлово» и «Все ООЗТ Москвы» («Сводка» — нет); заголовок — строка с «№ п/п»;
    столбцы — по началу названия (без учёта регистра и ё/е), мероприятия — по коду в заголовке («I.9 …»).
    Строки «ИТОГО» и без координат пропускаются; дубли (ООЗТ + квартал + название) объединяются по ИЛИ."""
    import openpyxl
    key = ("_wref", str(path), os.path.getmtime(path) if os.path.exists(path) else 0)
    if key in _REF_CACHE:
        return _REF_CACHE[key]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    recs, skipped = [], {"итого": 0, "без координат": 0}
    for ws in wb.worksheets:
        tn = _norm_txt(ws.title)
        if not (tn.startswith("битца и измайлово") or tn.startswith("все оозт")):
            continue
        rows = list(ws.iter_rows(values_only=True))
        h_i = next((i for i, r in enumerate(rows[:30])
                    if any(_norm_txt(c).replace(" ", "") == "№п/п" for c in r)), None)
        if h_i is None:
            log.warning("Справочник водоёмов, лист «%s»: не найдена строка заголовка «№ п/п»", ws.title)
            continue
        hdr = [_norm_txt(c) for c in rows[h_i]]

        def col(*stems):
            for st in stems:
                for i, h in enumerate(hdr):
                    if h.startswith(st):
                        return i
            return None
        c = dict(n=next(i for i, h in enumerate(hdr) if h.replace(" ", "") == "№п/п"),
                 oozt=col("оозт"), kv=col("квартал"), name=col("название"), cat=col("категория"),
                 otype=col("тип объекта", "тип"), area=col("площадь"), lat=col("широта"), lon=col("долгота"),
                 note=col("примечание"))
        mcols = {}
        for i, h in enumerate(rows[h_i]):
            m = re.match(r"^\s*([IVX]+\.\d+)\b", str(h or ""))
            if m:
                mcols[m.group(1)] = i
        for ri, r in enumerate(rows[h_i + 1:], start=h_i + 2):
            g = lambda k: _wr_val(r[c[k]]) if c.get(k) is not None and c[k] < len(r) else None
            if any("итого" in _norm_txt(x) for x in r[:max(c["n"], c["name"] or 0) + 1] if x is not None):
                skipped["итого"] += 1
                continue
            name = g("name")
            if name is None:
                continue
            try:
                lat, lon = float(str(g("lat")).replace(",", ".")), float(str(g("lon")).replace(",", "."))
                if not (40 < lat < 70 and 20 < lon < 60):
                    raise ValueError
            except (TypeError, ValueError):
                skipped["без координат"] += 1
                continue
            try:
                area = float(str(g("area")).replace(",", "."))
            except (TypeError, ValueError):
                area = np.nan
            recs.append(dict(sheet=ws.title, row=ri, oozt=str(g("oozt") or "").strip(),
                             kv=str(g("kv") or "").strip().removesuffix(".0"), name=str(name).strip(),
                             cat=str(g("cat") or "").strip(), otype=str(g("otype") or "").strip(),
                             area=area, lat=lat, lon=lon, note=str(g("note") or "").strip(),
                             **{f"m:{k}": _wr_mark(r[i]) if i < len(r) else False for k, i in mcols.items()}))
    df = pd.DataFrame(recs)
    if not df.empty:
        mc = [x for x in df.columns if x.startswith("m:")]
        df[mc] = df[mc].fillna(False).astype(bool)
        df["_k"] = df["oozt"].map(_norm_txt) + "|" + df["kv"].map(_norm_txt) + "|" + df["name"].map(_norm_txt)
        agg = {x: "first" for x in df.columns if x not in mc + ["_k", "note"]}
        agg.update({x: "max" for x in mc})
        agg["note"] = lambda s: "; ".join(dict.fromkeys(v for v in s if v))
        n0 = len(df)
        df = df.groupby("_k", as_index=False, sort=False).agg(agg)
        if len(df) < n0:
            log.info("Справочник водоёмов: дубли объединены (%d → %d строк)", n0, len(df))
    log.info("Справочник водоёмов %s: %d объектов с координатами (пропущено: ИТОГО — %d, без координат — %d)",
             Path(path).name, len(df), skipped["итого"], skipped["без координат"])
    _REF_CACHE[key] = df
    return df


def water_ref_marked(code: str) -> pd.DataFrame | None:
    """Объекты справочника, отмеченные для мероприятия code; None — справочник недоступен."""
    p = _ref_path("water_ref")
    if p is None:
        return None
    try:
        df = read_water_ref(str(p))
    except Exception as e:
        log.error("Справочник водоёмов %s не прочитан: %s", p.name, e)
        return None
    col = f"m:{code}"
    if df.empty or col not in df.columns:
        return df.iloc[0:0] if not df.empty else df
    return df[df[col]]


def invasive_ref() -> dict:
    """Инвазивные виды Москвы (перечень к ПП Москвы № 369-ПП от 03.03.2026):
    латинское название → группа (I–IV), жизненная форма, русское название."""
    return _load_ref("invasive_moscow.csv", ["group", "life_form", "ru"])


OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
    "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
]


# ---- лесопатология: ПВН → точки с категорией рекомендации ------------------------------
# 6 категорий по тексту рекомендаций (порядок важен: «эталон»/«надзор» проверяются раньше рубок)
LESOPAT_CATEGORIES = ["санитарная рубка", "рубки ухода", "надзор", "посадка", "эталон", "биотехния"]
LESOPAT_CAT_RE = [
    ("эталон", r"эталон|ппп\b"),
    ("биотехния", r"синичник"),
    ("надзор", r"надзор"),
    ("санитарная рубка",
     r"\bвср\b|\bяиуз\b|изумрудн|графиоз|некроз|категори[ия] состояния|пораж[её]нн"),
    ("рубки ухода", r"рубк|уход за подростом|разреживани|прочистк|осветлени|угнетающ"),
    ("посадка", r"посадк|дополнени|реставрац|высаживани"),
]


def classify_lesopat_recs(text) -> list[str]:
    """Текст рекомендаций ПВН → список категорий (может быть несколько)."""
    t = str(text or "").strip().lower().replace("ё", "е")
    if not t or t in ("-", "—", "нет", "nan", "none", "нет рекомендаций"):
        return []
    return [name for name, pat in LESOPAT_CAT_RE if re.search(pat, t)]


_DMS_RE = re.compile(
    r"^\s*([+-]?\d+(?:[.,]\d+)?)\s*[°ºdд]\s*"
    r"(?:(\d+(?:[.,]\d+)?)\s*['′mм]\s*)?"
    r"(?:(\d+(?:[.,]\d+)?)\s*(?:[\"″sс]|'')\s*)?"
    r"([NSEWСЮЗВnsewсюзв])?\s*$"
)


def _parse_coord_val(val) -> float | None:
    """Десятичные градусы (с точкой или запятой) либо градусы-минуты-секунды (55°35'54.20\"С)."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (int, float, np.number)):
        v = float(val)
        return v if np.isfinite(v) else None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "-", "—"):
        return None
    try:
        v = float(s.replace(",", "."))
        return v if np.isfinite(v) else None
    except ValueError:
        pass
    m = _DMS_RE.match(s)
    if m:
        deg = float(m.group(1).replace(",", "."))
        mn = float(m.group(2).replace(",", ".")) if m.group(2) else 0.0
        sec = float(m.group(3).replace(",", ".")) if m.group(3) else 0.0
        sign = -1.0 if (m.group(4) or "").upper() in ("S", "W", "Ю", "З") or deg < 0 else 1.0
        return sign * (abs(deg) + mn / 60.0 + sec / 3600.0)
    return None


def _norm_wgs84_pair(lat, lon) -> tuple[float, float, bool] | None:
    """Проверка пары широта/долгота WGS 84 (при перепутанных широте и долготе для Москвы — автоисправление)."""
    la, lo = _parse_coord_val(lat), _parse_coord_val(lon)
    if la is None or lo is None:
        return None
    swapped = False
    if 34.0 <= la <= 42.0 and 53.0 <= lo <= 59.0:
        la, lo = lo, la
        swapped = True
    if not (-90.0 <= la <= 90.0 and -180.0 <= lo <= 180.0):
        return None
    return la, lo, swapped


def _read_excel_lesopat_df(p: str) -> pd.DataFrame:
    """Чтение Excel ПВН с сохранением строковых номеров ПВН и автопоиском строки заголовка (строки 1–6)."""
    raw = pd.read_excel(p, header=None, dtype=object)
    hdr_idx = 0
    for idx in range(min(6, len(raw))):
        vals = " | ".join(str(x or "").strip().lower().replace("ё", "е") for x in raw.iloc[idx].values)
        if ("широт" in vals or "lat" in vals) and ("долгот" in vals or "lon" in vals):
            hdr_idx = idx
            break
    cols, seen = [], {}
    for j, x in enumerate(raw.iloc[hdr_idx].values):
        nm = str(x).strip() if x is not None and not (isinstance(x, float) and np.isnan(x)) else f"col_{j}"
        cnt = seen.get(nm, 0)
        seen[nm] = cnt + 1
        cols.append(nm if cnt == 0 else f"{nm}_{cnt + 1}")
    df = raw.iloc[hdr_idx + 1:].copy()
    df.columns = cols
    return df.reset_index(drop=True)


def read_lesopat(path: str, quiet: bool = False, return_all: bool = False):
    """Слой лесопатологии (ПВН) → точки с категорией рекомендации в WORK_EPSG.

    Поддерживает:
    • Excel (.xlsx/.xls): десятичные WGS 84 и/или DMS («55°35'54.20"С»), автопоиск заголовка;
    • Вектор (.gpkg, .shp, .geojson, .kml): поле готовой категории («Категория» / «cat») либо
      текст рекомендаций на русском или латинице («Рекомендации», «recomend», а также разбитые
      в DBF поля шейп-файлов «recomend_1» + «recomend_2»).
    При return_all=True возвращает кортеж (gdf_categorized, gdf_all_points)."""
    p = str(path)
    wlog = (lambda *a, **k: None) if quiet else log.warning
    if p.lower().endswith((".xlsx", ".xlsm", ".xls")):
        df = _read_excel_lesopat_df(p)
        cols = [str(c).strip() for c in df.columns]

        def col(*stems, allow_dms=False):
            for c in cols:
                if not allow_dms and "dms" in c.lower():
                    continue
                cl = c.lower().replace("ё", "е")
                if any(s in cl for s in stems):
                    return c
            return None

        c_num = col("№ пвн", "номер пвн", "пвн", "pvn", "point")
        c_lat = col("широт", "lat") or col("широт", "lat", allow_dms=True)
        c_lon = col("долгот", "lon", "long") or col("долгот", "lon", "long", allow_dms=True)
        c_lat_dms = col("широт", allow_dms=True)
        c_lon_dms = col("долгот", allow_dms=True)
        c_rec = col("рекомендац", "recomend", "recommend")
        c_cat = col("категория меропр", "категория пвн") or next(
            (c for c in cols if c.lower().replace("ё", "е") in ("категория", "cat", "category")), None)
        c_kv, c_vyd = col("квартал", "kv"), col("выдел", "vyd")
        if not (c_lat and c_lon and (c_rec or c_cat)):
            wlog("Лесопатология %s: не найдены колонки координат и «Рекомендации» (выберите файл "
                 "PVN_Lesopatologiya_*.xlsx либо вектор PVN_*.gpkg/.shp с рекомендациями) — слой не использован",
                 Path(p).name)
            return (None, None) if return_all else None
        rows, all_rows, n_swap = [], [], 0
        for _, r in df.iterrows():
            pair = _norm_wgs84_pair(r.get(c_lat), r.get(c_lon))
            if pair is None and c_lat_dms and c_lon_dms:
                pair = _norm_wgs84_pair(r.get(c_lat_dms), r.get(c_lon_dms))
            if pair is None:
                continue
            lat, lon, sw = pair
            n_swap += int(sw)
            pvn_v = "" if not c_num or r.get(c_num) is None else str(r.get(c_num)).strip()
            if pvn_v.lower() in ("nan", "none"):
                pvn_v = ""
            kv_v = "" if not c_kv or r.get(c_kv) is None else str(r.get(c_kv)).strip()
            vyd_v = "" if not c_vyd or r.get(c_vyd) is None else str(r.get(c_vyd)).strip()
            rec_txt = str(r.get(c_rec) or "").strip() if c_rec else ""
            pt = Point(lon, lat)
            all_rows.append({"pvn": pvn_v, "kv": kv_v, "vyd": vyd_v, "recs": rec_txt, "geometry": pt})
            cats = ([str(r.get(c_cat)).strip().lower().replace("ё", "е")]
                    if c_cat and str(r.get(c_cat) or "").strip() not in ("", "nan", "None", "-")
                    else classify_lesopat_recs(rec_txt))
            for cat in cats:
                rows.append({"pvn": pvn_v, "cat": cat, "kv": kv_v, "vyd": vyd_v, "geometry": pt})
        if n_swap and not quiet:
            log.warning("Лесопатология %s: в %d строках широта и долгота были переставлены — исправлено",
                        Path(p).name, n_swap)
        g_all = gpd.GeoDataFrame(all_rows, crs=4326).to_crs(WORK_EPSG) if all_rows else None
        if not rows:
            wlog("Лесопатология %s: рекомендации не распознаны — слой не использован", Path(p).name)
            return (None, g_all) if return_all else None
        g_cat = gpd.GeoDataFrame(rows, crs=4326).to_crs(WORK_EPSG)
        return (g_cat, g_all) if return_all else g_cat

    g = read_vector(p)
    if g is None or g.empty:
        return (None, None) if return_all else None
    g = g[g.geometry.notna() & ~g.geometry.is_empty].copy()
    if g.empty:
        return (None, None) if return_all else None
    if g.crs is not None and g.crs.is_geographic:
        bx = g.total_bounds
        if 53.0 <= bx[0] <= 59.0 and 34.0 <= bx[1] <= 42.0:
            if not quiet:
                log.warning("Лесопатология %s: координаты вектора имеют порядок (широта, долгота) — исправлено",
                            Path(p).name)
            g["geometry"] = g.geometry.map(lambda pt: Point(pt.y, pt.x))
    g = g.to_crs(WORK_EPSG)
    g["geometry"] = g.geometry.map(
        lambda geom: Point(geom.x, geom.y) if geom.geom_type == "Point" else geom.representative_point())

    cols_nv = [c for c in g.columns if c != "geometry"]

    def vcol(*exact_or_stems):
        for s in exact_or_stems:
            for c in cols_nv:
                cl = str(c).strip().lower().replace("ё", "е")
                if cl == s or cl.startswith(s):
                    return c
        for s in exact_or_stems:
            for c in cols_nv:
                cl = str(c).strip().lower().replace("ё", "е")
                if s in cl:
                    return c
        return None

    cpvn = vcol("pvn", "no_pvn", "num_pvn", "№ пвн", "номер пвн", "пвн", "point", "name")
    ckv = vcol("kv", "kvartal", "квартал")
    cvyd = vcol("vyd", "vydel", "выдел")
    ccat = next((c for c in cols_nv if str(c).strip().lower().replace("ё", "е")
                 in ("категория", "cat", "category")), None)
    # все столбцы рекомендаций: кириллица и латиница, включая разбитые DBF-поля recomend_1, recomend_2
    rec_cols = [c for c in cols_nv
                if any(k in str(c).strip().lower().replace("ё", "е")
                       for k in ("рекоменд", "recomend", "recommend", "rekomen"))]

    def _clean_s(v) -> str:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return ""
        s = str(v).strip()
        return "" if s.lower() in ("nan", "none") else s

    all_rows = []
    for _, r in g.iterrows():
        parts = [_clean_s(r.get(c)) for c in rec_cols]
        parts = [x for x in parts if x]
        all_rows.append({
            "pvn": _clean_s(r.get(cpvn)) if cpvn else "",
            "kv": _clean_s(r.get(ckv)) if ckv else "",
            "vyd": _clean_s(r.get(cvyd)) if cvyd else "",
            "recs": " ".join(parts),
            "geometry": r.geometry,
        })
    g_all = gpd.GeoDataFrame(all_rows, crs=WORK_EPSG) if all_rows else None

    if ccat:
        rows = []
        for _, r in g.iterrows():
            cat_v = _clean_s(r.get(ccat)).lower().replace("ё", "е")
            if not cat_v or cat_v in ("-", "—", "нет"):
                continue
            rows.append({
                "pvn": _clean_s(r.get(cpvn)) if cpvn else "",
                "cat": cat_v,
                "kv": _clean_s(r.get(ckv)) if ckv else "",
                "vyd": _clean_s(r.get(cvyd)) if cvyd else "",
                "geometry": r.geometry,
            })
        if not rows:
            wlog("Лесопатология %s: в поле «%s» нет категорий — слой не использован", Path(p).name, ccat)
            return (None, g_all) if return_all else None
        g_cat = gpd.GeoDataFrame(rows, crs=WORK_EPSG)
        return (g_cat, g_all) if return_all else g_cat

    if not rec_cols:
        wlog("Лесопатология %s: нет поля «Категория» или «Рекомендации»/«recomend*» (доступны поля: %s) — "
             "слой не использован", Path(p).name, ", ".join(map(str, cols_nv[:10])))
        return (None, g_all) if return_all else None

    rows = []
    for _, r in g.iterrows():
        parts = [_clean_s(r.get(c)) for c in rec_cols]
        parts = [x for x in parts if x]
        if not parts:
            continue
        # В шейп-файлах DBF обрезает строку на 254 байтах прямо внутри слова (recomend_1 + recomend_2):
        # проверяем как слитную склейку без пробела, так и склейку через пробел.
        cats = list(dict.fromkeys(
            classify_lesopat_recs("".join(parts)) + classify_lesopat_recs(" ".join(parts))
        ))
        pvn_v = _clean_s(r.get(cpvn)) if cpvn else ""
        kv_v = _clean_s(r.get(ckv)) if ckv else ""
        vyd_v = _clean_s(r.get(cvyd)) if cvyd else ""
        for cat in cats:
            rows.append({"pvn": pvn_v, "cat": cat, "kv": kv_v, "vyd": vyd_v, "geometry": r.geometry})
    if not rows:
        wlog("Лесопатология %s: рекомендации не распознаны — слой не использован", Path(p).name)
        return (None, g_all) if return_all else None
    g_cat = gpd.GeoDataFrame(rows, crs=WORK_EPSG)
    return (g_cat, g_all) if return_all else g_cat


OSM_UNPAVED_HIGHWAY = "path|footway|track|bridleway"
OSM_UNPAVED_SURFACE = "ground|dirt|earth|grass|mud|sand|unpaved"
OSM_TRAIL_KIND_TITLE = {"informal": "стихийные", "unpaved": "грунтовые"}


def osm_trails_query(bbox_wgs84, kinds=("informal",)) -> str:
    """Запрос Overpass: стихийные (informal=yes) и/или грунтовые тропы (highway=path|footway|track|
    bridleway с surface=ground|dirt|earth|grass|mud|sand|unpaved)."""
    w, s, e, n = [float(x) for x in bbox_wgs84]          # total_bounds: (minx, miny, maxx, maxy)
    bb = f"({s:.6f},{w:.6f},{n:.6f},{e:.6f})"
    parts = []
    if "informal" in kinds:
        parts.append(f'way["informal"="yes"]{bb};')
    if "unpaved" in kinds:
        parts.append(f'way["highway"~"^({OSM_UNPAVED_HIGHWAY})$"]["surface"~"^({OSM_UNPAVED_SURFACE})$"]{bb};')
    return "[out:json][timeout:90];(" + "".join(parts) + ");out geom;"


def _osm_trail_kind(tags: dict) -> str:
    t = tags or {}
    inf = str(t.get("informal", "")).lower() == "yes"
    unp = (re.fullmatch(OSM_UNPAVED_HIGHWAY, str(t.get("highway", ""))) is not None
           and re.fullmatch(OSM_UNPAVED_SURFACE, str(t.get("surface", ""))) is not None)
    return "стихийная, грунтовая" if inf and unp else "стихийная" if inf else "грунтовая" if unp else ""


OSM_SOURCES = {     # ключ → (подпись в окне, URL .pbf или None)
    "auto": ("автоматически: Overpass → BBBike → Geofabrik", None),
    "overpass": ("Overpass API (запрос по границам территории)", None),
    "bbbike": ("BBBike: Москва (.pbf, ~85 МБ)", "https://download.bbbike.org/osm/bbbike/Moscow/Moscow.osm.pbf"),
    "geofabrik": ("Geofabrik: ЦФО (.pbf, ~850 МБ)",
                  "https://download.geofabrik.de/russia/central-fed-district-latest.osm.pbf"),
    "pbf": ("свой файл OSM (.osm.pbf / .osm)", None),
}
BBBIKE_MOSCOW_EXTENT = (37.32, 55.57, 37.88, 55.92)   # охват выгрузки BBBike «Moscow» (W, S, E, N)
PBF_TYPES = [("OpenStreetMap", "*.osm.pbf *.pbf *.osm"), ("Все файлы", "*.*")]


def _osm_trails_overpass(bbox_wgs84, kinds):
    """Тропы через Overpass API (по очереди зеркала OVERPASS_ENDPOINTS) → GeoDataFrame или None."""
    q = osm_trails_query(bbox_wgs84, kinds)
    log.info("Тропы OSM: запрос Overpass: %s", q)
    data = None
    for url in OVERPASS_ENDPOINTS:
        try:
            req = urllib.request.Request(url, data=q.encode("utf-8"),
                                         headers={"User-Agent": "oozt-recs/1.0 (protected area recommendations)"})
            with urllib.request.urlopen(req, timeout=45) as r:
                data = json.loads(r.read().decode("utf-8"))
            break
        except Exception as e:
            log.warning("Overpass %s: %s", url, e)
    if not data:
        log.warning("Тропы OSM: Overpass API недоступен")
        return None
    lines, ids, kd, hw, sf = [], [], [], [], []
    for el in data.get("elements", []):
        if el.get("type") != "way" or "geometry" not in el:
            continue
        coords = [(p["lon"], p["lat"]) for p in el["geometry"]]
        if len(coords) >= 2:
            tg = el.get("tags") or {}
            lines.append(LineString(coords))
            ids.append(el.get("id"))
            kd.append(_osm_trail_kind(tg))
            hw.append(tg.get("highway", ""))
            sf.append(tg.get("surface", ""))
    return gpd.GeoDataFrame({"osm_id": ids, "kind": kd, "highway": hw, "surface": sf},
                            geometry=lines, crs=4326)


_HSTORE_RE = re.compile(r'"((?:[^"\\]|\\.)*)"=>"((?:[^"\\]|\\.)*)"')


def _parse_other_tags(v) -> dict:
    """Поле other_tags драйвера GDAL OSM: "k"=>"v","k2"=>"v2" → dict."""
    if not isinstance(v, str) or not v:
        return {}
    return {k.replace('\\"', '"'): x.replace('\\"', '"') for k, x in _HSTORE_RE.findall(v)}


def _osm_trails_pbf(path, bbox_wgs84, kinds):
    """Тропы из файла OSM (.osm.pbf / .osm) через драйвер GDAL OSM (слой lines) → GeoDataFrame или None."""
    import pyogrio
    p = Path(path)
    if not p.is_file():
        log.warning("Тропы OSM: файл не найден — %s", p)
        return None
    w, s_, e, n = [float(x) for x in bbox_wgs84]
    log.info("Тропы OSM: чтение %s (%.0f МБ), охват %.4f, %.4f — %.4f, %.4f",
             p.name, p.stat().st_size / 1e6, w, s_, e, n)
    df = pyogrio.read_dataframe(str(p), layer="lines", bbox=(w, s_, e, n))
    if df is None or df.empty:
        return gpd.GeoDataFrame({"osm_id": [], "kind": [], "highway": [], "surface": []},
                                geometry=[], crs=4326)
    ot = df["other_tags"] if "other_tags" in df.columns else pd.Series([None] * len(df), index=df.index)
    rows = []
    for idx, v in ot.items():
        tg = _parse_other_tags(v)
        for c in ("highway", "surface", "informal"):
            if c in df.columns and isinstance(df.at[idx, c], str) and df.at[idx, c]:
                tg[c] = df.at[idx, c]
        k = _osm_trail_kind(tg)
        ok = ("informal" in kinds and "стихийная" in k) or ("unpaved" in kinds and "грунтовая" in k)
        if ok:
            rows.append((idx, k, tg.get("highway", ""), tg.get("surface", "")))
    sel = df.loc[[r[0] for r in rows]]
    g = gpd.GeoDataFrame({"osm_id": sel["osm_id"].values if "osm_id" in sel.columns else range(len(sel)),
                          "kind": [r[1] for r in rows], "highway": [r[2] for r in rows],
                          "surface": [r[3] for r in rows]},
                         geometry=sel.geometry.values, crs=df.crs or 4326)
    return g.to_crs(4326)


def _osm_cache_dir(cache_dir="") -> Path:
    d = Path(cache_dir) if str(cache_dir or "").strip() else Path(__file__).resolve().parent / "osm_cache"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _download_pbf(key, cache_dir="", max_age_days=30) -> Path | None:
    """Скачивает выгрузку OSM (.pbf) в кэш; файл моложе max_age_days используется повторно."""
    url = OSM_SOURCES[key][1]
    p = _osm_cache_dir(cache_dir) / Path(urllib.parse.urlparse(url).path).name
    if p.is_file() and p.stat().st_size > 0:
        age = (time.time() - p.stat().st_mtime) / 86400
        if age <= float(max_age_days or 0):
            log.info("Тропы OSM: используется скачанный ранее %s (%.0f дн., %.0f МБ)", p, age, p.stat().st_size / 1e6)
            return p
    log.info("Тропы OSM: скачивание %s → %s", url, p)
    tmp = p.with_suffix(p.suffix + ".part")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "oozt-recs/1.0 (protected area recommendations)"})
        with urllib.request.urlopen(req, timeout=60) as r, open(tmp, "wb") as f:
            total = int(r.headers.get("Content-Length") or 0)
            got, step = 0, 0
            while True:
                b = r.read(1 << 20)
                if not b:
                    break
                f.write(b)
                got += len(b)
                if total and got * 10 // total > step:
                    step = got * 10 // total
                    log.info("  скачано %d%% (%.0f из %.0f МБ)", step * 10, got / 1e6, total / 1e6)
        if total and got < total:
            raise IOError(f"получено {got} из {total} байт")
        tmp.replace(p)
        return p
    except Exception as e:
        log.warning("Тропы OSM: %s не скачан (%s)", url, e)
        try:
            tmp.unlink()
        except Exception:
            pass
        return p if p.is_file() else None      # старый файл из кэша лучше, чем ничего


def _bbox_inside(bb, ext) -> bool:
    return bb[0] >= ext[0] and bb[1] >= ext[1] and bb[2] <= ext[2] and bb[3] <= ext[3]


def download_informal_trails_osm(bbox_wgs84, out_dir, kinds=("informal",), source="auto", pbf="",
                                 cache_dir="", cache_days=30) -> Path | None:
    """Тропы OSM → GeoJSON (WGS 84): стихийные (informal=yes) и/или грунтовые.
    Источник: Overpass API, выгрузка BBBike (Москва), Geofabrik (ЦФО) или свой файл .osm.pbf;
    auto — по очереди Overpass → BBBike → Geofabrik. При неудаче возвращает None — скрипт продолжит без слоя."""
    kinds = [k for k in (kinds or ["informal"]) if k in OSM_TRAIL_KIND_TITLE] or ["informal"]
    source = str(source or "auto").strip().lower()
    if source not in OSM_SOURCES:
        log.warning("Тропы OSM: неизвестный источник «%s» — используется auto", source)
        source = "auto"
    bb = [float(x) for x in bbox_wgs84]
    order = {"auto": ["overpass", "bbbike", "geofabrik"]}.get(source, [source])
    if source == "auto" and str(pbf or "").strip():
        order = ["pbf"] + order                  # указан свой файл — он первым
    log.info("Тропы OSM (%s): источник — %s", ", ".join(OSM_TRAIL_KIND_TITLE[k] for k in kinds),
             " → ".join(OSM_SOURCES[k][0] for k in order))
    g, used = None, ""
    for src in order:
        try:
            if src == "overpass":
                g = _osm_trails_overpass(bb, kinds)
            elif src == "pbf":
                if not str(pbf or "").strip():
                    log.warning("Тропы OSM: файл .osm.pbf не выбран")
                    continue
                g = _osm_trails_pbf(pbf, bb, kinds)
            else:
                if src == "bbbike" and not _bbox_inside(bb, BBBIKE_MOSCOW_EXTENT):
                    log.warning("Тропы OSM: территория выходит за охват выгрузки BBBike «Moscow» %s — пропуск",
                                BBBIKE_MOSCOW_EXTENT)
                    continue
                fp = _download_pbf(src, cache_dir, cache_days)
                g = _osm_trails_pbf(fp, bb, kinds) if fp else None
        except Exception as e:
            log.warning("Тропы OSM: источник «%s» — ошибка (%s)", OSM_SOURCES[src][0], e)
            g = None
        if g is not None:
            used = src
            break
    if g is None:
        log.warning("Тропы OSM не получены ни из одного источника — слой пропущен, "
                    "правила «только стихийные тропы» используют всю дорожно-тропиночную сеть")
        return None
    if g.empty:
        log.warning("В OSM нет троп (%s) в границах территории (источник — %s)",
                    ", ".join(OSM_TRAIL_KIND_TITLE[k] for k in kinds), OSM_SOURCES[used][0])
        return None
    g["source"] = OSM_SOURCES[used][0]
    out = Path(out_dir) if str(out_dir or "").strip() else Path.cwd()
    out.mkdir(parents=True, exist_ok=True)
    p = out / "informal_trails_osm.geojson"
    g.to_file(p, driver="GeoJSON", encoding="utf-8")
    log.info("Тропы OSM: %d сегментов (%s), источник — %s → %s", len(g),
             ", ".join(f"{k} — {v}" for k, v in g["kind"].value_counts().items()), OSM_SOURCES[used][0], p.name)
    return p


class Data:
    """Все загруженные слои в WORK_EPSG + служебные объекты."""

    def __init__(self, cfg: Config, oozt_only: bool = False, pad_m: float = 0.0):
        global _CUR_CFG
        self.cfg = cfg
        _CUR_CFG = cfg
        self.kvnet = None                   # kv_cut: квартальная сеть (все кварталы, включая пустые)
        self.oozt = self._load_oozt()
        u_parts = list(self.oozt.geometry.values)
        if self.kvnet is not None and not self.kvnet.empty:
            u_parts.extend(self.kvnet.geometry.values)
        self.union = _safe_union(u_parts)
        self.bounds = tuple(self.union.bounds)
        pad = max(TRAIL_WINDOW_M, CORVUS_WINDOW_M, RARE_BUFFER_M, SHORE_BUFFER_M, pad_m) + 2 * CELL
        b = self.bounds
        self.bounds_pad = (b[0] - pad, b[1] - pad, b[2] + pad, b[3] + pad)
        self.bbox_wgs84 = tuple(gpd.GeoSeries([box(*self.bounds_pad)], crs=WORK_EPSG).to_crs(4326).total_bounds)
        self.clip = box(*self.bounds_pad)
        self.trails = self.hydro_lines = self.hydro_polys = None
        self.trails_informal = None                # стихийные тропы OSM (informal=yes)
        self.sp_points = self.sp_polys = None     # находки видов: species, category, group, layer, kk_cat, __src
        self.sp_rasters: list[dict] = []          # модели SDM: path, category, group, species
        self.dem = None
        self._inside = None
        # все слои инициализированы: режим --report-only выходит раньше загрузки,
        # но схемы/проект QGIS обращаются к атрибутам
        self.drainage = self.wetlands = self.rec_points = None
        self.roads_paved = self.roads_dirt = None
        self.lesopat = None
        self.lesopat_all = None
        if oozt_only:                      # режим --report-only: остальные слои не нужны
            return

        self.trails = self._load_trails()
        self.trails_informal = self._load_informal_trails()
        self.hydro_lines = read_many(cfg.hydro_lines, self.bbox_wgs84, self.clip)
        self.hydro_polys = read_many(cfg.hydro_polys, self.bbox_wgs84, self.clip)
        self.drainage = read_many(cfg.drainage, self.bbox_wgs84, self.clip)
        self.wetlands = read_many(cfg.wetlands, self.bbox_wgs84, self.clip)
        self.rec_points = read_many(cfg.recreation, self.bbox_wgs84, self.clip)
        self.roads_paved = read_many(cfg.roads_paved, self.bbox_wgs84, self.clip)
        self.roads_dirt = read_many(cfg.roads_dirt, self.bbox_wgs84, self.clip)
        self.lesopat = self._load_lesopat()
        # дренажные канавы добавляются к рекам (подложка карт и мероприятия у воды)
        if self.drainage is not None and not self.drainage.empty:
            parts = [g for g in (self.hydro_lines, self.drainage) if g is not None and not g.empty]
            if parts:
                self.hydro_lines = gpd.GeoDataFrame(pd.concat(parts, ignore_index=True), crs=WORK_EPSG)
                log.info("Дренажные канавы добавлены к рекам (линии гидрографии): +%d", len(self.drainage))
        self.sp_points, self.sp_polys, self.sp_rasters = self._load_species_layers(species_layers(cfg))
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
        cfg = self.cfg
        g = read_vector(cfg.oozt_path, cfg.oozt_layer)
        g = g[g.geom_type.isin(["Polygon", "MultiPolygon"])].copy()
        g["_part"], g["_nparts"] = 1, 1
        if getattr(cfg, "split_multipart", True):
            g = _split_multipart(g)
        f = cfg.fields or {}
        # поле «Структурный класс» (I.7): копия до переименования полей; cfg.fields ниже перезаписывается
        f_struct = f.get(STRUCT_FIELD)
        if f_struct and f_struct in g.columns:
            g["Struct"] = g[f_struct].values
        elif f_struct:
            log.warning("Поле «Структурный класс» «%s» не найдено в слое территории", f_struct)
            f_struct = None
        cols = list(g.columns)

        def pick(key, candidates, required=True):
            if f.get(key) and f[key] in cols:
                return f[key]
            for c in candidates:
                if c in cols:
                    return c
            for c in cols:
                if c.lower() in [x.lower() for x in candidates]:
                    return c
            if required:
                raise KeyError(f"В слое территории не найдено поле {key} (искал {candidates}); "
                               f"задайте его в окне выбора или в cfg.fields['{key}']")
            return None

        spec = {k: (req, cands) for k, _, req, cands in FIELD_SPECS[cfg.unit_mode]}
        if cfg.unit_mode == "eco":
            self.F_ID = pick("ID", spec["ID"][1])
            self.F_TYPE = pick("TYPE", spec["TYPE"][1])
            self.F_NAME = pick("NAME", spec["NAME"][1], required=False)
            self.F_OOZT = pick("OOZT", spec["OOZT"][1], required=False)
            cfg.fields = {k: v for k, v in {"ID": self.F_ID, "TYPE": self.F_TYPE, "NAME": self.F_NAME,
                                            "OOZT": self.F_OOZT}.items() if v}
            _dup = g[[self.F_ID, "_part"]].duplicated()
            if _dup.any():
                dup = g.loc[_dup, self.F_ID].astype(str).unique()[:10]
                raise ValueError(f"Поле ID «{self.F_ID}» не уникально: повторы {', '.join(dup)}")
            gf = cfg.report_group_field
            if gf and gf not in cols:
                raise KeyError(f"Поле группировки «{gf}» не найдено в слое территории")
            grp_vals = g[gf].astype(str).str.strip().values if gf else None
            g = g.drop(columns=[c for c in ("OOZT", "KV", "VYD", "Type", "Name", "KV_SORT")
                                if c in cols and c not in (self.F_ID, self.F_TYPE, self.F_NAME, self.F_OOZT)])
            if self.F_NAME:
                g["Name"] = g[self.F_NAME]
            if not self.F_OOZT:
                g["OOZT"] = cfg.oozt_name or Path(cfg.oozt_path).stem
            else:
                g["OOZT"] = g[self.F_OOZT].astype(str).str.strip()
            g[self.F_ID] = _part_ids(g[self.F_ID], g)
            g["VYD"] = g[self.F_ID]
            g["Type"] = g[self.F_TYPE]
            if not self.F_NAME:
                g["Name"] = g["Type"]
            g["KV"], g["KV_SORT"] = self._report_groups(g, grp_vals)
        elif cfg.unit_mode == "kv_cut":
            g = self._cut_by_kvartals(g, pick)
        else:
            self.F_OOZT = pick("OOZT", spec["OOZT"][1])
            self.F_KV = pick("KV", spec["KV"][1])
            self.F_VYD = pick("VYD", spec["VYD"][1])
            self.F_TYPE = pick("TYPE", spec["TYPE"][1])
            self.F_NAME = pick("NAME", spec["NAME"][1], required=False)
            cfg.fields = {k: v for k, v in {"OOZT": self.F_OOZT, "KV": self.F_KV, "VYD": self.F_VYD,
                                            "TYPE": self.F_TYPE, "NAME": self.F_NAME}.items() if v}
            g = g.rename(columns={self.F_OOZT: "OOZT", self.F_KV: "KV", self.F_VYD: "VYD",
                                  self.F_TYPE: "Type", **({self.F_NAME: "Name"} if self.F_NAME else {})})
            if not self.F_NAME:
                g["Name"] = g["Type"]
            g["VYD"] = _part_ids(g["VYD"], g)
        n_split = int((g["_nparts"] > 1).sum())
        if n_split:
            log.info("Многочастные объекты разделены на отдельные контуры: %d объектов → %d контуров "
                     "(ID-1, ID-2, …; контуры < %.2f га присоединены к соседнему крупному контуру)",
                     g.loc[g["_nparts"] > 1, "_src_row"].nunique(), n_split, MIN_PART_HA)
        g = g.drop(columns=["_part", "_nparts", "_src_row"], errors="ignore")
        for c in ("KV", "VYD"):
            g[c] = _maybe_int(g[c])
        if cfg.unit_mode != "eco":
            g["KV_SORT"] = _nat_rank(g["KV"])
        g["VYD_SORT"] = _nat_rank(g["VYD"])
        g["Type"] = g["Type"].astype(str).str.strip()
        g["Name"] = g["Name"].astype(str).str.strip()
        if f_struct and "Struct" in g.columns:
            cfg.fields[STRUCT_FIELD] = f_struct
            g["Struct"] = g["Struct"].map(struct_class_of)
            log.info("Структурный класс — поле «%s»: %s", f_struct, ", ".join(
                f"{k} — {v}" for k, v in g["Struct"].value_counts().head(12).items()))
        else:
            g["Struct"] = ""
        g["uid"] = g["OOZT"].astype(str) + "|" + g["KV"].astype(str) + "|" + g["VYD"].astype(str)
        g["area_ha"] = g.geometry.area / 1e4
        g = g.reset_index(drop=True)
        L = unit_labels(cfg)
        log.info("Слой территории (%s): %d %s, %s: %d, ООЗТ: %s", UNIT_MODES[cfg.unit_mode], len(g),
                 L["units_gen"], L["groups_n"].lower(), g[["OOZT", "KV"]].drop_duplicates().shape[0],
                 ", ".join(map(str, g["OOZT"].unique())))
        return g

    def _split_dup_kv_parts(self, kv: gpd.GeoDataFrame, cfg, eco_oozt: str = "") -> gpd.GeoDataFrame:
        """Разделение по кластерам ООЗТ при дублировании нумерации кварталов.

        Номера кварталов иногда совпадают в разных массивах одной сети (например, 1–3
        есть и в основном массиве Измайлово, и в Терлецкой дубраве). После dissolve по
        номеру они слились в один многочастный квартал, и мероприятия назначались обеим
        частям сразу — «Северный Терлецкий пруд» попадал в квартал 1 основного массива.

        Части дублированных кварталов группируются в массивы по взаимному касанию
        (разрыв больше KV_SPLIT_GAP_M внутри одного номера — уже разные массивы).
        Главный массив — тот, чьи кварталы граничат с большим числом прочих кварталов
        сети (у небольшого массива соседей мало). Он сохраняет имя ООЗТ; остальные
        получают метку «Измайлово (юго-восточный массив)» — по направлению от центра
        главного массива. Так меры назначаются в пределах своего массива, а списки
        кварталов в отчёте остаются уникальными.
        """
        import math
        from collections import defaultdict
        from shapely.ops import unary_union
        from shapely.strtree import STRtree
        fallback = ((cfg.oozt_name or "").strip() or eco_oozt
                   or Path(cfg.oozt_path).stem)
        # все части с их номерами и исходными метками ООЗТ
        parts, p_kv, p_oo = [], [], []
        for _, r in kv.iterrows():
            oo = str(r["KVOOZT"]).strip() or fallback
            for g in (r.geometry.geoms if r.geometry.geom_type in (
                    "MultiPolygon", "GeometryCollection") else [r.geometry]):
                if not g.is_empty and g.area > 0:
                    parts.append(g); p_kv.append(r["KV"]); p_oo.append(oo)
        n = len(parts)

        def _clusters(kvsel):
            """Кластеры частей одного номера: разрыв ≤ KV_SPLIT_GAP_M — один массив."""
            idx = [i for i in range(n) if p_kv[i] == kvsel]
            parent = list(range(len(idx)))

            def find(a):
                while parent[a] != a:
                    parent[a] = parent[parent[a]]
                    a = parent[a]
                return a
            tr = STRtree([parts[i] for i in idx])
            for ai in range(len(idx)):
                for aj in tr.query(parts[idx[ai]], predicate="dwithin",
                                   distance=KV_SPLIT_GAP_M):
                    aj = int(aj)
                    if aj > ai:
                        ra, rb = find(ai), find(aj)
                        if ra != rb:
                            parent[rb] = ra
            cl = defaultdict(list)
            for ai in range(len(idx)):
                cl[find(ai)].append(idx[ai])
            return list(cl.values())

        dup_kvs = {kvv for kvv in set(p_kv) if len(_clusters(kvv)) > 1}
        if not dup_kvs:
            out = gpd.GeoDataFrame({"KV": p_kv, "KVOOZT": p_oo, "geometry": parts},
                                   crs=kv.crs or WORK_EPSG)
            return out.dissolve(by=["KV", "KVOOZT"], as_index=False)
        # массивы: группы частей разных дублированных номеров, касающиеся друг друга
        dup_idx = [i for i in range(n) if p_kv[i] in dup_kvs]
        cl2id = {}
        for kvv in dup_kvs:
            for ci, idxs in enumerate(_clusters(kvv)):
                for i in idxs:
                    cl2id[i] = (kvv, ci)
        ids = list(cl2id.values())
        gid = {c: k for k, c in enumerate(ids)}
        parent = list(range(len(ids)))

        def find(a):
            while parent[a] != a:
                parent[a] = parent[parent[a]]
                a = parent[a]
            return a
        tr = STRtree([parts[i] for i in dup_idx])
        for i in dup_idx:
            for j in tr.query(parts[i], predicate="dwithin", distance=KV_SPLIT_GAP_M):
                j = dup_idx[int(j)]            # STRtree возвращает позицию в dup_idx
                if j <= i or p_kv[i] == p_kv[j]:
                    continue
                ra, rb = find(gid[cl2id[i]]), find(gid[cl2id[j]])
                if ra != rb:
                    parent[rb] = ra
        groups = defaultdict(list)
        for i in dup_idx:
            groups[find(gid[cl2id[i]])].append(i)
        # главный массив: максимум соседних прочих (недублированных) кварталов сети
        main_root = max(groups, key=lambda g: (
            sum(1 for i in range(n) if p_kv[i] not in dup_kvs and parts[i].distance(
                unary_union([parts[j] for j in groups[g]])) <= KV_SPLIT_GAP_M),
            sum(parts[i].area for i in groups[g])))
        main_c = unary_union([parts[i] for i in groups[main_root]]).centroid
        dirs = ["восточный", "северо-восточный", "северный", "северо-западный",
                "западный", "юго-западный", "южный", "юго-восточный"]   # atan2: 0° = восток
        gname, used = {}, {}
        for g in groups:
            if g == main_root:
                continue
            c = unary_union([parts[i] for i in groups[g]]).centroid
            ang = math.degrees(math.atan2(c.y - main_c.y, c.x - main_c.x)) % 360
            d = dirs[int((ang + 22.5) % 360 // 45)]
            used[d] = used.get(d, 0) + 1
            d = d if used[d] == 1 else f"{d}-{used[d]}"
            gname[g] = f"{p_oo[groups[g][0]]} ({d} массив)"
            log.warning(
                "Дублирование нумерации кварталов: массив из кварталов %s отделён от "
                "главного и выделен в отдельное ООЗТ «%s»", 
                ", ".join(sorted({p_kv[i] for i in groups[g]}, key=lambda x: (len(x), x))),
                gname[g])
        for g, nm in gname.items():
            for i in groups[g]:
                p_oo[i] = nm
        out = gpd.GeoDataFrame({"KV": p_kv, "KVOOZT": p_oo, "geometry": parts},
                               crs=kv.crs or WORK_EPSG)
        return out.dissolve(by=["KV", "KVOOZT"], as_index=False)

    def _cut_by_kvartals(self, g: gpd.GeoDataFrame, pick) -> gpd.GeoDataFrame:
        """Режим kv_cut: экосистемы разрезаются границами кварталов. В каждом квартале — новая
        нумерация получившихся выделов с 1: от северо-западного угла квартала, с запада на восток,
        с севера на юг (по положению центрода: сначала север, при равной северности — западнее левее)."""
        cfg = self.cfg
        spec = {k: cands for k, _, _, cands in FIELD_SPECS["kv_cut"]}
        self.F_TYPE = pick("TYPE", spec["TYPE"])
        self.F_NAME = pick("NAME", spec["NAME"], required=False)
        self.F_OOZT = pick("OOZT", spec["OOZT"], required=False)
        self.F_ID = pick("ID", spec["ID"], required=False)
        cfg.fields = {k: v for k, v in {"TYPE": self.F_TYPE, "NAME": self.F_NAME, "OOZT": self.F_OOZT,
                                        "ID": self.F_ID}.items() if v}
        ren = {self.F_TYPE: "Type"}
        if self.F_NAME and self.F_NAME != self.F_TYPE:
            ren[self.F_NAME] = "Name"
        if self.F_OOZT:
            ren[self.F_OOZT] = "OOZT"
        if self.F_ID:
            ren[self.F_ID] = "ECO_ID"
        g = g.rename(columns=ren)
        g = g.drop(columns=[c for c in ("KV", "VYD", "KVOOZT", "KV_SORT", "VYD_SORT", "uid", "area_ha")
                             if c in g.columns], errors="ignore")
        if "Name" not in g.columns:
            g["Name"] = g["Type"]
        g = repair_geometries(g)          # make_valid: защита от TopologyException при разрезании

        if not cfg.kvnet_path:
            raise KeyError("Не выбран вектор квартальной сети (kvnet_path)")
        kv = read_vector(cfg.kvnet_path, cfg.kvnet_layer)
        kv = kv[kv.geom_type.isin(["Polygon", "MultiPolygon"])].copy()
        if kv.empty:
            raise ValueError("В векторе квартальной сети нет полигонов")
        kf = cfg.kvnet_field or next((c for c in KVNET_KV_CANDIDATES if c in kv.columns), "")
        if kf not in kv.columns:
            raise KeyError(f"В квартальной сети не найдено поле номера квартала «{kf or '?'}» "
                           f"(задайте kvnet_field)")
        kv["KV"] = kv[kf].astype(str).str.strip()
        ko = cfg.kvnet_oozt_field
        if ko and ko in kv.columns:
            kv["KVOOZT"] = kv[ko].astype(str).str.strip().replace({"nan": "", "None": ""})
        else:
            kv["KVOOZT"] = ""
        kv = kv[["KV", "KVOOZT", "geometry"]].copy()
        kv = repair_geometries(kv)
        kv = kv.dissolve(by=["KV", "KVOOZT"], as_index=False)   # многочастный квартал → один полигон
        # Дублирование нумерации кварталов: один номер (1–3, …) встречается в нескольких
        # несвязанных массивах сети. Растворение превратило их в один многочастный квартал,
        # и мероприятия назначались обеим частям разом (например, «Северный Терлецкий пруд»
        # оказывался в квартале 1 основного массива Измайлово). Решение: каждая связная
        # часть получает отдельную метку ООЗТ. Крупнейший массив — под именем самого ООЗТ,
        # остальные — «Имя (направление массив)» по компасу относительно главного массива.
        # базовое имя ООЗТ для кварталов сети без своей метки: если экосистемы
        # однородны, берём их имя (иначе — имя файла или cfg.oozt_name)
        eco_oozt = ""
        if "OOZT" in g.columns and not g.empty:
            vc = g["OOZT"].astype(str).str.strip().value_counts()
            if len(vc) == 1:
                eco_oozt = vc.index[0]
        kv = self._split_dup_kv_parts(kv, cfg, eco_oozt)
        log.info("Квартальная сеть: %s, поле «%s»: %d кварталов", Path(cfg.kvnet_path).name, kf,
                 len(kv))
        self.kvnet = kv                      # все кварталы сети — для схем пустых кварталов

        eco_area_ha = g.geometry.area.sum() / 1e4
        inter = _safe_overlay(g, kv, how="intersection", keep_geom_type=True)
        inter = inter.explode(index_parts=False).reset_index(drop=True)
        if inter.empty:
            raise ValueError("Пересечение экосистем с квартальной сетью пусто: слои не перекрываются")
        min_ha = float(getattr(cfg, "min_cut_ha", 0.01) or 0)
        if min_ha > 0:
            a = inter.geometry.area / 1e4
            small = a < min_ha
            if small.any():
                log.info("Разрезание: отброшено %d мельчайших кусочков < %.3f га (всего %.2f га)",
                         int(small.sum()), min_ha, float(a[small].sum()))
                inter = inter[~small].reset_index(drop=True)
        covered = inter.geometry.area.sum() / 1e4
        if eco_area_ha - covered > 0.01:
            log.info("Разрезание: %.1f га (%.1f%%) экосистем лежит вне квартальной сети — в выделы не вошли",
                     eco_area_ha - covered, 100 * (eco_area_ha - covered) / eco_area_ha)

        o = (inter["OOZT"].astype(str).str.strip() if "OOZT" in inter.columns
             else pd.Series("", index=inter.index)).replace({"nan": "", "None": ""})
        kvo = inter["KVOOZT"].astype(str).str.strip().replace({"nan": "", "None": ""})
        inter["OOZT"] = o.where(o != "", kvo)
        fallback = (cfg.oozt_name or "").strip() or Path(cfg.oozt_path).stem
        inter["OOZT"] = inter["OOZT"].where(inter["OOZT"] != "", fallback)
        # дублированные номера: метка массива из сети главнее имени ООЗТ слоя экосистем —
        # иначе обе части квартала снова сольются в одно ООЗТ
        _dup_kv = set(kv.loc[kv["KV"].duplicated(keep=False), "KV"].astype(str))
        _m = inter["KV"].astype(str).isin(_dup_kv) & (kvo != "")
        inter["OOZT"] = inter["OOZT"].where(~_m, kvo[_m])
        if "Name" not in inter.columns:
            inter["Name"] = inter["Type"]
        # заполняем KVOOZT у кварталов сети, если оно осталось пустым (для сводки площадей и пустых кварталов)
        if self.kvnet is not None and not self.kvnet.empty:
            kn_oo = self.kvnet["KVOOZT"].astype(str).str.strip().replace({"nan": "", "None": ""})
            if (kn_oo == "").any():
                dom = (inter.groupby("KV")["OOZT"].agg(lambda s: s.value_counts().index[0])
                       if not inter.empty else pd.Series(dtype=str))
                def_oo = dom.value_counts().index[0] if len(dom) else fallback
                self.kvnet = self.kvnet.copy()
                self.kvnet["KVOOZT"] = [
                    cur if cur else str(dom.get(str(kvv), def_oo))
                    for kvv, cur in zip(self.kvnet["KV"], kn_oo)
                ]

        # нумерация выделов в каждом (ООЗТ, квартал): с северо-запада, с запада на восток, с севера на юг
        c_ = inter.geometry.centroid
        inter["_y"], inter["_x"] = c_.y.values, c_.x.values
        inter["_kvs"] = _nat_rank(inter["KV"].astype(str))
        inter["_oos"] = _nat_rank(inter["OOZT"].astype(str))
        inter = inter.sort_values(["_oos", "_kvs", "_y", "_x"], ascending=[True, True, False, True],
                                  kind="mergesort").reset_index(drop=True)
        inter["VYD"] = inter.groupby(["OOZT", "KV"], sort=False).cumcount() + 1
        log.info("Экосистемы разрезаны кварталами: %d экосистем → %d выделов; нумерация в каждом "
                 "квартале с 1 (северо-запад → восток, север → юг)", len(g), len(inter))
        return inter.drop(columns=[c for c in ("_x", "_y", "_kvs", "_oos", "KVOOZT") if c in inter.columns])

    def _report_groups(self, g: gpd.GeoDataFrame, grp_vals=None):
        """Режим eco: группа отчёта — значение поля группировки или лист регулярной сетки (A1, A2, …)."""
        cfg = self.cfg
        self.sheet_grid = None
        if grp_vals is not None:
            return grp_vals, _nat_rank(grp_vals)
        size = float(cfg.sheet_size_m or 1000.0)
        minx, miny, maxx, maxy = g.total_bounds
        pts = g.geometry.representative_point()
        col = np.floor((pts.x.values - minx) / size).astype(int)
        row = np.floor((maxy - pts.y.values) / size).astype(int)     # ряды сверху вниз (с севера)

        def letters(n):
            s = ""
            n += 1
            while n:
                n, r = divmod(n - 1, 26)
                s = chr(65 + r) + s
            return s
        labels = [f"{letters(r)}{c + 1}" for r, c in zip(row, col)]
        self.sheet_grid = dict(minx=minx, maxy=maxy, size=size)
        log.info("Листы сетки %.0f × %.0f м: %d листов с полигонами", size, size, len(set(labels)))
        return labels, row * 10000 + col

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

    # ---- стихийные тропы (OSM informal=yes) ----
    def _load_informal_trails(self):
        cfg = self.cfg
        paths = [str(x).strip() for x in (cfg.informal_trails or []) if str(x).strip()]
        if getattr(cfg, "_need_informal", None) is False:
            return None                      # ни одному включённому мероприятию слой не нужен
        auto = any(x.lower() in ("auto", "авто", "osm") for x in paths)
        files = [x for x in paths if x.lower() not in ("auto", "авто", "osm")]
        if auto:
            try:
                p = download_informal_trails_osm(self.bbox_wgs84, cfg.out_dir,
                                                 getattr(cfg, "osm_trail_kinds", None) or ["informal"],
                                                 getattr(cfg, "osm_source", "auto"), getattr(cfg, "osm_pbf", ""),
                                                 getattr(cfg, "osm_cache_dir", ""),
                                                 getattr(cfg, "osm_cache_days", 30))
                if p is not None:
                    files.append(str(p))
            except Exception as e:
                log.warning("Стихийные тропы: скачивание из OSM не удалось (%s)", e)
        if not files:
            if not paths:
                log.info("Стихийные тропы: не используются (выбрано «не использовать»)")
            return None
        own = [x for x in files if not Path(x).name.startswith("informal_trails_osm")]
        if own:
            log.info("Стихийные тропы: свой слой — %s", "; ".join(Path(x).name for x in own))
        g = read_many(files, self.bbox_wgs84, self.clip)
        if g is None or g.empty:
            log.warning("Стихийные тропы: в выбранных файлах нет объектов в границах территории")
            return None
        g = repair_geometries(g)          # make_valid: защита от TopologyException
        if g is None or g.empty:
            return None
        g = g[g.geom_type.isin(["LineString", "MultiLineString"])].copy()
        if g.empty:
            log.warning("Стихийные тропы: в выбранных файлах нет линий")
            return None
        log.info("Стихийные тропы: %d сегментов (слой для правил «только стихийные тропы»)", len(g))
        return g

    # ---- виды ----
    def _load_species_layers(self, layers):
        """Слои видов → общие таблицы точек и полигонов + список растров SDM (с категорией и группой)."""
        pts, polys, rasters = [], [], []
        for i, l in enumerate(layers):
            name = Path(str(l["path"]).split("|")[0]).name
            tag = f"{l['category']}, {l['group']}"
            if l["kind"] == "sdm":
                sp = species_from_filename(str(l["path"]).split("|")[0])
                rasters.append(dict(l, species=sp, kk_cat=kk_category(sp), layer=i))
                log.info("Слой видов %s (%s): модель SDM, вид «%s»", name, tag, sp)
                continue
            if l["kind"] not in ("points", "polygons"):
                log.warning("Слой видов %s: не удалось определить тип данных — слой не используется", name)
                continue
            g = read_many([l["path"]], self.bbox_wgs84, self.clip)
            g = repair_geometries(g)      # make_valid: защита от TopologyException
            want = ["Point", "MultiPoint"] if l["kind"] == "points" else ["Polygon", "MultiPolygon"]
            if g is not None:
                g = g[g.geom_type.isin(want)]
                if l["kind"] == "points":
                    g = g.explode(index_parts=False)
            if g is None or g.empty:
                log.warning("Слой видов %s (%s): нет объектов в границах ООЗТ — слой не используется", name, tag)
                continue
            g = g.copy()
            col = l["field"] if l.get("field") in g.columns else detect_species_column(g)
            g["species"] = g[col].astype(str).str.strip() if col else g["__src"]
            g["kk_cat"] = g["species"].map(kk_category)      # категория Красной книги Москвы
            n_syn = int(g["species"].map(kk_via_synonym).sum())
            if n_syn:
                log.info("Слой видов %s: %d находок распознаны в Красной книге через синоним (%s)", name, n_syn,
                         ", ".join(sorted(g.loc[g["species"].map(kk_via_synonym), "species"].unique())[:6]))
            g["category"], g["group"], g["layer"] = l["category"], l["group"], i
            if col:
                log.info("Слой видов %s (%s, %s): поле вида «%s», %d видов, %d объектов", name, tag,
                         SPECIES_KINDS[l["kind"]], col, g["species"].nunique(), len(g))
            else:
                log.warning("Слой видов %s (%s): поле с латинскими названиями не найдено — вид = имя файла",
                            name, tag)
            kk_sp = g.loc[g["kk_cat"].notna()].groupby("species")["kk_cat"].first()
            if len(kk_sp):
                log.info("Слой видов %s: в Красной книге Москвы — %s", name,
                         ", ".join(f"{s} (кат. {c})" for s, c in kk_sp.head(8).items())
                         + ("…" if len(kk_sp) > 8 else ""))
            (pts if l["kind"] == "points" else polys).append(
                g[["species", "category", "group", "layer", "kk_cat", "__src", "geometry"]])
        cat = lambda parts: gpd.GeoDataFrame(pd.concat(parts, ignore_index=True), crs=WORK_EPSG) if parts else None
        return cat(pts), cat(polys), rasters

    # ---- маска «внутри ООЗТ» на сетке CELL ----
    def inside_mask(self, transform, shp):
        m = rfeatures.rasterize([(self.union, 1)], out_shape=shp, transform=transform, fill=0, dtype="uint8")
        return m.astype(bool)

    # ---- лесопатология: точки ПВН с категорией рекомендации ----
    def _load_lesopat(self):
        layers, all_layers = [], []
        for pth in self.cfg.lesopat_points or []:
            try:
                g, g_all = read_lesopat(str(pth), return_all=True)
            except Exception as e:
                log.error("Лесопатология %s: %s", Path(pth).name, e)
                continue
            if g_all is not None and not g_all.empty:
                all_layers.append(g_all)
            if g is not None and not g.empty:
                layers.append(g)
        if all_layers:
            ga = gpd.GeoDataFrame(pd.concat(all_layers, ignore_index=True), crs=WORK_EPSG)
            if self.clip is not None and not ga.empty:
                ga = ga[ga.geometry.within(self.clip.buffer(30))]
            if not ga.empty:
                self.lesopat_all = ga
        if not layers:
            return None
        g = gpd.GeoDataFrame(pd.concat(layers, ignore_index=True), crs=WORK_EPSG)
        if self.clip is not None and not g.empty:
            n_all = len(g)
            # допуск 30 м: точки на границе ООЗТ (ПВН ставится по границе выдела)
            # не отбрасываются молча — предупреждаем, если что-то отфильтровано
            g = g[g.geometry.within(self.clip.buffer(30))]
            if len(g) < n_all:
                log.info("Лесопатология: отобрано %d из %d записей в пределах территории (остальные относятся к другим ООЗТ)",
                         len(g), n_all)
        if g.empty:
            log.warning("Лесопатология: ни одна точка из загруженного файла не попала в границы территории")
            return None
        log.info("Лесопатология: %d назначений по точкам ПВН (%s)", len(g),
                 ", ".join(f"{c}: {n}" for c, n in g["cat"].value_counts().items()))
        return g

    # ---- справочник водоёмов: «ООЗТ — название — тип — площадь» ----
    def water_catalog(self) -> gpd.GeoDataFrame:
        cols = ["ООЗТ", "Водоём", "Тип", "Площадь, га"]
        if self.hydro_polys is None or self.hydro_polys.empty:
            return gpd.GeoDataFrame(columns=cols + ["geometry"], crs=WORK_EPSG)
        wb = self.hydro_polys
        names = _water_names(wb)
        tcol = next((c for c in wb.columns if c != "geometry" and str(c).lower()
                     in ("natural", "waterway", "wetland", "type", "тип", "water")), None)
        oozt_d = self.oozt.dissolve(by="OOZT")
        rows = []
        for idx, geom in zip(wb.index, wb.geometry):
            nm = str(names.loc[idx]).strip() if names is not None else ""
            if nm in ("", "nan", "None"):
                nm = "без названия"
            typ = str(wb.at[idx, tcol]).strip() if tcol else ""
            if typ in ("", "nan", "None"):
                typ = "не указан"
            hits = oozt_d[oozt_d.intersects(geom)]
            oo = ""
            if len(hits):
                inter = hits.geometry.intersection(geom)
                oo = str(inter.area.idxmax())
            ha = geom.area / 1e4
            rows.append({"ООЗТ": oo, "Водоём": nm, "Тип": typ,
                         "Площадь, га": round(ha, 2), "geometry": geom})
        return gpd.GeoDataFrame(rows, crs=WORK_EPSG)

    def vydels_of_type(self, *keys) -> gpd.GeoDataFrame:
        has_name = "Name" in self.oozt.columns
        m = self.oozt.apply(
            lambda r: type_matches(r.get("Type"), keys,
                                   r.get("Name") if has_name else None),
            axis=1).values.astype(bool)
        return self.oozt[m]

    def quarters(self) -> gpd.GeoDataFrame | None:
        """Границы кварталов: в режиме kv_cut — полная квартальная сеть (включая кварталы без выделов),
        иначе — выделы ООЗТ, объединённые по ООЗТ и номеру квартала."""
        kn = getattr(self, "kvnet", None)
        if self.cfg.unit_mode == "kv_cut" and kn is not None and not kn.empty:
            kv = kn.copy()
            oo = kv["KVOOZT"].astype(str).str.strip().replace({"nan": "", "None": ""})
            if (oo == "").any():
                u_oo = self.oozt["OOZT"].astype(str).unique() if self.oozt is not None and not self.oozt.empty else []
                fb = str(u_oo[0]) if len(u_oo) == 1 else ((self.cfg.oozt_name or "").strip() or Path(self.cfg.oozt_path).stem)
                oo = oo.replace("", fb)
            kv["OOZT"] = oo
            q = kv[["OOZT", "KV", "geometry"]].dissolve(by=["OOZT", "KV"], as_index=False)
            q["KV"] = _maybe_int(q["KV"])
            q["KV_SORT"] = _nat_rank(q["KV"])
            q["area_ha"] = (q.geometry.area / 1e4).round(2)
            return q.sort_values(["OOZT", "KV_SORT"]).reset_index(drop=True)
        if self.oozt is None or self.oozt.empty:
            return None
        by = [c for c in ("OOZT", "KV", "KV_SORT") if c in self.oozt.columns]
        q = self.oozt.dissolve(by=by, as_index=False)
        q["area_ha"] = (q.geometry.area / 1e4).round(2)
        return q


# =====================================================================================
#                                    РЕЗУЛЬТАТЫ
# =====================================================================================
def _compact_quarter_basis(bases: list[str]) -> str:
    """Сжатие однотипных оснований десятков выделов внутри одного квартала (в режиме учёта по кварталам),
    с полным сохранением всех координат WGS 84 и номеров ПВН."""
    uniq = list(dict.fromkeys([x for x in bases if x]))
    if len(uniq) <= 3:
        return "; ".join(uniq)
    if any("координаты WGS 84" in x or "точки ПВН" in x for x in uniq):
        return "; ".join(uniq)
    # «Колокольчик»: агрегируем зоны и уникальные местообитания вместо повтора одинаковых предложений
    if all(x.startswith("зоны:") and "Местообитание:" in x for x in uniq):
        ztypes, habs = [], []
        tot_ring, tot_strip = 0.0, 0.0
        for x in uniq:
            head, _, tail = x.partition("Местообитание:")
            m_r = re.search(r"кольцо ([^;(]+?)\s*\(([\d.,]+)\s*га\)", head)
            if m_r:
                lbl_r = f"кольцо {m_r.group(1).strip()}"
                if lbl_r not in ztypes:
                    ztypes.append(lbl_r)
                tot_ring += float(m_r.group(2).replace(",", "."))
            m_s = re.search(r"полоса ([^;(]+?)\s*\(([\d.,]+)\s*га\)", head)
            if m_s:
                lbl_s = f"полоса {m_s.group(1).strip()}"
                if lbl_s not in ztypes:
                    ztypes.append(lbl_s)
                tot_strip += float(m_s.group(2).replace(",", "."))
            if "по границе выдела" in head and "по границе выдела" not in ztypes:
                ztypes.append("по границе выдела")
            ht = tail.strip()
            if ht and ht not in habs:
                habs.append(ht)
        z_out = []
        for zt in ztypes:
            if zt.startswith("кольцо ") and tot_ring > 0:
                z_out.append(f"{zt} ({tot_ring:.2f} га)")
            elif zt.startswith("полоса ") and tot_strip > 0:
                z_out.append(f"{zt} ({tot_strip:.2f} га)")
            else:
                z_out.append(zt)
        return (f"зоны ({len(uniq)} уч.): {'; '.join(z_out)}. "
                f"Местообитания: {' | '.join(habs)}")
    return "; ".join(uniq[:4]) + f" (и ещё {len(uniq) - 4} уч. в квартале)"


def _oozt_key_match(flag_key, oozt) -> bool:
    """Ключ столбца «Вкл_<ООЗТ>» / ООЗТ справочника совпадает с ООЗТ слоя: точно, по началу
    («Измайлово» ~ «Измайлово (юго-восточный массив)») или по общему корню из 4 букв («Битца» ~ «Битцевский лес»)."""
    a, b = _norm_txt(flag_key), _norm_txt(oozt)
    if not a or not b:
        return False
    if a == b or b.startswith(a) or a.startswith(b):
        return True
    return len(a) >= 4 and len(b) >= 4 and a[:4] == b[:4]


def oozt_flag(measure: dict, oozt):
    """Флаг «Вкл_<ООЗТ>» мероприятия для ООЗТ: True / False / None (нет ограничения)."""
    flags = measure.get("oozt_on") or {}
    if not flags:
        return None
    oo = str(oozt).strip()
    if oo in flags:
        return flags[oo]
    for k, v in flags.items():
        if _oozt_key_match(k, oo):
            return v
    return None


# =====================================================================================
#        ЖУРНАЛ ПРИЧИН: почему мероприятие назначено или не назначено (по каждой ООЗТ)
# =====================================================================================
class Diag:
    """Собирает по каждому мероприятию: статус до расчёта, сообщения правила, воронку отбора
    по ООЗТ (сколько полигонов прошло каждое условие), отброшенные флагами ООЗТ и итог."""

    def __init__(self, data=None):
        self.items: list[dict] = []
        self.cur: dict | None = None
        self.data = data
        self._oo_union = None

    def start(self, measure: dict, status: dict | None = None):
        self.cur = dict(code=str(measure.get("code") or ""), name=measure["name"], measure=measure,
                        state=(status or {}).get("state"), status=(status or {}).get("text", ""),
                        rule=(status or {}).get("rule"), lines=[], steps=[], drop={}, error=None)
        self.items.append(self.cur)
        return self.cur

    def stop(self):
        self.cur = None

    def note(self, level: str, msg: str):
        if self.cur is not None:
            self.cur["lines"].append((level, msg.strip()))

    def dropped(self, oozt):
        if self.cur is not None:
            self.cur["drop"][oozt] = self.cur["drop"].get(oozt, 0) + 1

    def step(self, title: str, g, col: str | None = "OOZT"):
        """Шаг воронки: число объектов всего и по ООЗТ (если в g есть поле ООЗТ)."""
        if self.cur is None or g is None:
            return
        try:
            if col and col in g.columns:
                by = g[col].value_counts().to_dict()
            else:
                by = self._by_oozt(g)
            title = title if len(title) <= 160 else title[:157] + "…"
            self.cur["steps"].append((title, int(len(g)), {str(k): int(v) for k, v in by.items()}))
        except Exception:
            pass

    def _by_oozt(self, g) -> dict:
        """Число объектов по ООЗТ пространственно (для слоёв без поля ООЗТ: находки видов и т. п.)."""
        d = self.data
        if d is None or getattr(d, "oozt", None) is None or "geometry" not in getattr(g, "columns", []):
            return {}
        try:
            if self._oo_union is None:
                self._oo_union = d.oozt[["OOZT", "geometry"]].dissolve(by="OOZT").reset_index()
            pts = gpd.GeoDataFrame(geometry=g.geometry.representative_point().values, crs=g.crs)
            j = gpd.sjoin(pts, self._oo_union, predicate="within", how="left")
            j = j[~j.index.duplicated()]
            by = j["OOZT"].value_counts().to_dict()
            for oo in self._oo_union["OOZT"]:
                by.setdefault(oo, 0)
            return by
        except Exception:
            return {}

    def step_counts(self, title: str, by: dict):
        if self.cur is not None:
            self.cur["steps"].append((title, int(sum(by.values())), {str(k): int(v) for k, v in by.items()}))


DIAG = Diag()


class _DiagHandler(logging.Handler):
    def emit(self, record):
        try:
            DIAG.note(record.levelname, record.getMessage())
        except Exception:
            pass


_FLAG_TXT = "выключено для этой ООЗТ (столбец «Вкл_…» таблицы мероприятий)"


def _mentions(msg: str, oozt: str, all_oozt) -> bool:
    """Сообщение относится к ООЗТ (без ложных совпадений «Измайлово» в «Измайлово (Терлецкий)»)."""
    if not oozt:
        return False
    t = msg
    for o in sorted(all_oozt, key=len, reverse=True):
        if o != oozt and oozt in o:
            t = t.replace(o, "")
    return oozt in t


def _diag_reason_zero(it: dict, oozt: str, all_oozt=(), n_other: int = 0) -> str:
    """Причина, по которой мероприятие не назначено в ООЗТ (из воронки и сообщений правила)."""
    if it.get("error"):
        return "ошибка при расчёте: " + it["error"]
    # первый шаг воронки, на котором в этой ООЗТ не осталось объектов
    for title, n, by in it["steps"]:
        if by and by.get(oozt, 0) == 0:
            return (f"на шаге «{title}» в этой ООЗТ 0 (всего по территории: {n})")
        if not by and n == 0:
            return f"на шаге «{title}» не осталось ни одного объекта"
    msgs = [(lv, m) for lv, m in it["lines"] if not m.startswith("Мероприятие:")]
    own = [(lv, m) for lv, m in msgs if _mentions(m, oozt, all_oozt)]
    own_w = [m for lv, m in own if lv in ("WARNING", "ERROR")]
    if own_w:
        return own_w[-1]
    if own:
        return own[-1][1]
    if it.get("rule") == "water_ref":
        return f"в справочнике водоёмов нет отметок {it['code']} для объектов этой ООЗТ"
    if n_other:
        return (f"в этой ООЗТ нет участков, удовлетворяющих условиям правила "
                f"(в других ООЗТ назначено записей: {n_other})")
    warn = [m for lv, m in msgs if lv in ("WARNING", "ERROR")]
    if warn:
        return warn[-1]
    zero = [m for lv, m in msgs if re.search(r"(— |: |→ )0\b", m)]
    if zero:
        return zero[-1]
    return "ни один участок не прошёл условия правила" + (f" (последнее сообщение: {msgs[-1][1]})" if msgs else "")


def diag_table(data, res) -> pd.DataFrame:
    """Строка на каждое мероприятие × ООЗТ: итог, число участков, площадь, причина, ход расчёта."""
    oozts = [str(x) for x in pd.unique(data.oozt["OOZT"])] if data is not None and data.oozt is not None else []
    rows_by = {}
    for r in res.rows:
        k = (r["Мероприятие"], str(r["ООЗТ"]))
        a = rows_by.setdefault(k, [0, 0.0, set()])
        a[0] += 1
        a[1] += float(r.get("_work_ha") or 0)
        a[2].add(str(r.get("Квартал")))
    out = []
    for it in DIAG.items:
        m = it["measure"]
        steps_txt = "\n".join(
            f"{t}: {n}" + (" (" + ", ".join(f"{k} — {v}" for k, v in by.items()) + ")" if by else "")
            for t, n, by in it["steps"])
        log_txt = "\n".join(f"{'! ' if lv in ('WARNING', 'ERROR') else ''}{msg}" for lv, msg in it["lines"]
                             if not msg.startswith("Мероприятие:"))
        for oo in oozts or [""]:
            n, ha, kvs = rows_by.get((it["name"], oo), (0, 0.0, set()))
            flag = oozt_flag(m, oo) if oo else None
            n_other = sum(v[0] for (nm, o2), v in rows_by.items() if nm == it["name"] and o2 != oo)
            if it["state"] == "off":
                res_txt, why = "выключено", "выключено в таблице (столбец «Вкл/выкл»)"
            elif flag is False:
                res_txt = "выключено для ООЗТ"
                d = it["drop"].get(oo, 0)
                why = _FLAG_TXT + (f"; подходящих участков отброшено: {d}" if d else "")
            elif it["state"] == "na":
                res_txt, why = "без правила", "правило не задано: " + (it["status"] or "назначается вручную")
            elif it["state"] == "missing":
                res_txt, why = "пропущено", it["status"]
            elif n:
                res_txt = "назначено"
                why = f"{n} записей, кварталов: {len(kvs)}, {ha:.2f} га"
            else:
                res_txt, why = "не назначено", _diag_reason_zero(it, oo, oozts, n_other)
            out.append({"Код": it["code"], "Мероприятие": it["name"], "ООЗТ": oo, "Итог": res_txt,
                        "Записей": n, "Кварталов": len(kvs), "Площадь, га": round(ha, 2), "Причина": why,
                        "Правило": RULES[it["rule"]]["title"] if it.get("rule") in RULES else "",
                        "Воронка отбора": steps_txt, "Сообщения расчёта": log_txt})
    return pd.DataFrame(out)


def write_diagnostics(out: Path, data, res):
    """diagnostics.xlsx + diagnostics.txt + итоговый блок в журнале."""
    df = diag_table(data, res)
    if df.empty:
        return df
    try:
        with pd.ExcelWriter(out / "diagnostics.xlsx", engine="openpyxl") as xw:
            df.to_excel(xw, sheet_name="Причины", index=False)
            ws = xw.sheets["Причины"]
            from openpyxl.styles import Alignment, PatternFill, Font
            widths = {"A": 8, "B": 45, "C": 22, "D": 18, "E": 9, "F": 10, "G": 11, "H": 70, "I": 28,
                      "J": 60, "K": 90}
            for c, w in widths.items():
                ws.column_dimensions[c].width = w
            fills = {"назначено": "E2F0D9", "не назначено": "FCE4D6", "пропущено": "FFF2CC",
                     "выключено для ООЗТ": "EDEDED", "выключено": "EDEDED", "без правила": "EDEDED"}
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
                f = fills.get(row[3].value)
                if f:
                    row[3].fill = PatternFill("solid", fgColor=f)
            for c in ws[1]:
                c.font = Font(bold=True)
            ws.freeze_panes = "D2"
            ws.auto_filter.ref = ws.dimensions
    except Exception as e:
        log.warning("diagnostics.xlsx не записан: %s", e)
    lines = ["ПОЧЕМУ МЕРОПРИЯТИЯ НАЗНАЧЕНЫ ИЛИ НЕ НАЗНАЧЕНЫ", "=" * 60]
    for (code, name), g in df.groupby(["Код", "Мероприятие"], sort=False):
        lines.append(f"\n{code} {name}".rstrip())
        st = g.iloc[0]
        if st["Правило"]:
            lines.append(f"  правило: {st['Правило']}")
        for _, r in g.iterrows():
            lines.append(f"  [{r['ООЗТ'] or '—'}] {r['Итог'].upper()}: {r['Причина']}")
        if st["Воронка отбора"]:
            lines.append("  воронка отбора:")
            lines += ["    " + x for x in str(st["Воронка отбора"]).split("\n")]
    txt = "\n".join(lines)
    (out / "diagnostics.txt").write_text(txt + "\n", encoding="utf-8")
    log.info("\n%s", txt)
    cnt = df["Итог"].value_counts().to_dict()
    log.info("Итог по мероприятиям × ООЗТ: %s. Подробно: diagnostics.xlsx / diagnostics.txt",
             ", ".join(f"{k} — {v}" for k, v in cnt.items()))
    return df


def _inf_label() -> str:
    """Подпись слоя стихийных троп: «… OSM» при автоскачивании, «… (свой слой)» при своём файле."""
    v = [str(x).strip().lower() for x in (getattr(_CUR_CFG, "informal_trails", None) or ["auto"])]
    if not any(x in ("auto", "авто", "osm") for x in v):
        return "стихийные тропы (свой слой)"
    k = getattr(_CUR_CFG, "osm_trail_kinds", None) or ["informal"]
    if "unpaved" in k and "informal" in k:
        return "стихийные и грунтовые тропы OSM"
    return "грунтовые тропы OSM" if "unpaved" in k else "стихийные тропы OSM"


class Results:
    def __init__(self):
        self.rows: list[dict] = []          # длинная таблица (+ геометрия участка работ _geom в WORK_EPSG)
        self.evidence: list[gpd.GeoDataFrame] = []

    def add(self, vydel_row, measure: dict, basis: str, value=None, work_geom=None,
            season: str | None = None, obj=None) -> bool:
        """Запись назначения. season — сезонность записи (вместо столбца таблицы); obj — идентификатор
        объекта (водоём/болото справочника) для подсчёта «Объектов, шт». False — выключено флагом ООЗТ."""
        if oozt_flag(measure, vydel_row["OOZT"]) is False:   # мероприятие выключено для этой ООЗТ
            DIAG.dropped(vydel_row["OOZT"])
            return False
        wg = work_geom if work_geom is not None else vydel_row.get("geometry")
        if wg is not None and not wg.is_empty:
            if wg.geom_type in ("Point", "MultiPoint", "LineString", "MultiLineString"):
                wg = wg.buffer(5.0).intersection(vydel_row.geometry)
            elif wg.geom_type == "GeometryCollection":
                polys = [p for p in wg.geoms if p.geom_type in ("Polygon", "MultiPolygon") and not p.is_empty]
                wg = _safe_union(polys) if polys else vydel_row.geometry
        if wg is None or wg.is_empty:
            wg = vydel_row.get("geometry")
        work_ha = round(float(wg.area / 1e4), 4) if wg is not None and not wg.is_empty else 0.0
        self.rows.append({
            "ООЗТ": vydel_row["OOZT"], "Квартал": vydel_row["KV"], "Выдел": vydel_row["VYD"],
            "uid": vydel_row.get("uid"),
            "_sort": vydel_row.get("KV_SORT", vydel_row["KV"]),
            "_vsort": vydel_row.get("VYD_SORT", 0),
            "Тип выдела": vydel_row["Type"], "Экосистема": vydel_row["Name"],
            "Тип мероприятий": measure.get("group", ""), "Мероприятие": measure["name"],
            "Тип экосистемы мероприятия": measure.get("eco", ""),
            "Пункт Приказа": measure.get("prikaz", ""),
            "Сезонность": season if season is not None else measure.get("season", ""),
            "Основание": basis, "Значение": value,
            "_geom": wg, "_work_ha": work_ha, "_obj": obj,
        })
        return True

    def add_evidence(self, g: gpd.GeoDataFrame, measure_name: str, layer: str, label_col: str | None = None):
        if g is None or g.empty:
            return
        g = g[g.geometry.notna() & ~g.geometry.is_empty]
        if g.empty:
            return
        n = len(g)
        labels = g[label_col].astype(str).values if label_col and label_col in g.columns else [""] * n
        oo_col = next((c for c in ("_oozt", "oozt", "OOZT") if c in g.columns), None)
        kv_col = next((c for c in ("_kv", "kv", "KV") if c in g.columns), None)
        oo_vals = g[oo_col].fillna("").astype(str).values if oo_col else [""] * n
        kv_vals = g[kv_col].fillna("").astype(str).values if kv_col else [""] * n
        e = gpd.GeoDataFrame(
            {"measure": [measure_name] * n, "layer": [layer] * n, "label": list(labels),
             "oozt": list(oo_vals), "kv": list(kv_vals)},
            geometry=g.geometry.values, crs=WORK_EPSG,
        )
        self.evidence.append(e)

    def long_table(self) -> pd.DataFrame:
        df = pd.DataFrame([{k: v for k, v in r.items() if k not in ("_geom", "_obj")} for r in self.rows])
        if df.empty:
            return df
        # объединяем дубли (один выдел — одно мероприятие) с конкатенацией оснований
        key = ["ООЗТ", "Квартал", "Выдел", "Мероприятие"]
        agg = df.groupby(key, as_index=False, sort=False).agg({
            "Тип выдела": "first", "Экосистема": "first", "Тип мероприятий": "first", "_sort": "first", "_vsort": "first",
            "Тип экосистемы мероприятия": "first",
            "Пункт Приказа": "first", "Сезонность": merge_seasons,
            "Основание": lambda s: "; ".join(dict.fromkeys([x for x in s if x])),
            "Значение": "first",
        })
        if _CUR_CFG is not None and _CUR_CFG.unit_mode != "eco":
            # учётная единица — квартал: строки одного квартала и мероприятия объединяются,
            # номера выделов и типы экосистем из таблицы не показываются (экосистемы задают
            # только границы участков работ на схемах и в ГИС)
            def _join_val(s):
                vals = ["" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x) for x in s]
                nums = pd.to_numeric(pd.Series(vals), errors="coerce")
                if len(vals) >= 2 and nums.notna().sum() >= max(2, int(len(vals) * 0.8)):
                    return round(float(np.nansum(nums)), 2)
                return "; ".join(dict.fromkeys([v for v in vals if v]))
            agg = agg.groupby(["ООЗТ", "Квартал", "Мероприятие"], as_index=False, sort=False).agg({
                "Тип мероприятий": "first", "_sort": "first",
                "Тип экосистемы мероприятия": "first",
                "Пункт Приказа": "first", "Сезонность": merge_seasons,
                "Основание": _compact_quarter_basis,
                "Значение": _join_val,
            })
        if _CUR_CFG is not None and _CUR_CFG.unit_mode != "eco":
            agg = agg.sort_values(["ООЗТ", "_sort", "Тип мероприятий", "Мероприятие"])
        else:
            agg = agg.sort_values(["ООЗТ", "_sort", "_vsort", "Тип мероприятий", "Мероприятие"])
        if _CUR_CFG is not None and _CUR_CFG.unit_mode != "eco":
            cols = ["ООЗТ", "Квартал", "Тип мероприятий", "Мероприятие",
                    "Тип экосистемы мероприятия", "Пункт Приказа", "Сезонность", "Основание", "Значение"]
        else:
            cols = ["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема", "Тип мероприятий", "Мероприятие",
                    "Тип экосистемы мероприятия", "Пункт Приказа", "Сезонность", "Основание", "Значение"]
        return agg[cols].reset_index(drop=True)

    def measures_gdf(self, table: pd.DataFrame, clip_geom=None) -> gpd.GeoDataFrame | None:
        """Географически очерченные полигоны проведения мероприятий (БЕЗ целых кварталов) в WORK_EPSG."""
        if table is None or table.empty or not self.rows:
            return None
        qmode = _CUR_CFG is None or _CUR_CFG.unit_mode != "eco"
        mkeys = ["ООЗТ", "Квартал", "Мероприятие"] if qmode else ["ООЗТ", "Квартал", "Выдел", "Мероприятие"]
        by_key: dict[tuple, list] = {}
        for r in self.rows:
            g = r.get("_geom")
            if g is None or g.is_empty:
                continue
            k = tuple(str(r[c]) for c in mkeys)
            by_key.setdefault(k, []).append(g)
        if not by_key:
            return None
        geoms, areas = [], []
        keep_idx = []
        for idx, row in table.iterrows():
            k = tuple(str(row[c]) for c in mkeys)
            gs = by_key.get(k)
            if not gs:
                continue
            u = _safe_union(gs)
            if clip_geom is not None and not clip_geom.is_empty:
                u = u.intersection(clip_geom)
            if u is not None and u.geom_type == "GeometryCollection":
                polys = [p for p in u.geoms if p.geom_type in ("Polygon", "MultiPolygon") and not p.is_empty]
                u = _safe_union(polys) if polys else None
            if u is None or u.is_empty or u.area <= 0.1:
                continue
            keep_idx.append(idx)
            geoms.append(u)
            areas.append(round(float(u.area / 1e4), 3))
        if not geoms:
            return None
        out = table.loc[keep_idx].copy().reset_index(drop=True)
        out["Площадь участка, га"] = areas
        return gpd.GeoDataFrame(out, geometry=geoms, crs=WORK_EPSG)

    def oozt_covered_ha(self, clip_geom=None) -> dict[str, float]:
        """Истинная площадь покрытия ООЗТ мероприятиями (геометрическое объединение всех участков работ, га)."""
        by_oo: dict[str, list] = {}
        for r in self.rows:
            g = r.get("_geom")
            if g is not None and not g.is_empty:
                by_oo.setdefault(str(r["ООЗТ"]), []).append(g)
        res_d = {}
        for oo, gs in by_oo.items():
            u = _safe_union(gs)
            if clip_geom is not None and not clip_geom.is_empty:
                u = u.intersection(clip_geom)
            res_d[oo] = round(float(u.area / 1e4), 2) if u is not None and not u.is_empty else 0.0
        return res_d

    def evidence_gdf(self) -> gpd.GeoDataFrame | None:
        if not self.evidence:
            return None
        ev = gpd.GeoDataFrame(pd.concat(self.evidence, ignore_index=True), crs=WORK_EPSG)

        def _homog(g):
            if g is None or g.is_empty or g.geom_type != "GeometryCollection":
                return g
            for kinds in (("Polygon", "MultiPolygon"), ("LineString", "MultiLineString"), ("Point", "MultiPoint")):
                parts = [p for p in g.geoms if p.geom_type in kinds and not p.is_empty]
                if parts:
                    return _safe_union(parts)
            return None

        ev["geometry"] = ev.geometry.map(_homog)
        ev = ev[ev.geometry.notna() & ~ev.geometry.is_empty].reset_index(drop=True)
        return ev if not ev.empty else None


# =====================================================================================
#                                     ПРАВИЛА
# =====================================================================================
_CUR_CFG: Config | None = None     # текущая конфигурация (для подписей в журнале и основаниях)


def _units() -> str:
    return unit_labels(_CUR_CFG)["units_gen"] if _CUR_CFG else "выделов"


def _unit_word() -> str:
    return "полигона" if (_CUR_CFG and _CUR_CFG.unit_mode == "eco") else "выдела"


def _overlay_patches_to_vydels(data: Data, patches: gpd.GeoDataFrame, target: gpd.GeoDataFrame,
                               min_ha=MIN_PATCH_HA, share_pct: float | None = None):
    """Пересечение полигонов-пятен с выделами; возвращает DataFrame uid → площадь пятен (га), доля."""
    if patches is None or patches.empty or target.empty:
        return pd.DataFrame(columns=["uid", "patch_ha", "share"])
    pu = _safe_union(patches.geometry.values)
    inter = target.geometry.intersection(pu)
    ha = inter.area / 1e4
    df = pd.DataFrame({"uid": target["uid"].values, "patch_ha": ha.values,
                       "share": (ha / target["area_ha"]).values})
    df = df[df["patch_ha"] >= min_ha]
    thr = _min_share(data) if share_pct is None else float(share_pct)
    if thr > 0 and len(df):
        n0 = len(df)
        df = df[df["share"] * 100 >= thr]
        log.info("  порог: участок ≥ %g%% площади полигона — оставлено %d из %d", thr, len(df), n0)
    return df


def _min_share(data) -> float:
    try:
        return float(getattr(data.cfg, "min_share_pct", 0) or 0)
    except (TypeError, ValueError):
        return 0.0


def _target(data: Data, p: dict) -> gpd.GeoDataFrame:
    t = p.get("types")
    g = data.vydels_of_type(*t[0]) if t else data.oozt
    if t:
        if g is None or g.empty:
            log.warning("  нет полигонов типа «%s» в слое территории", t[1])
        else:
            DIAG.step("полигонов типа «%s»" % t[1], g)
    return g


def _qtag(q) -> str:
    return f"Q{int(round(float(q) * 100))}"


def _eco_label(p: dict) -> str:
    t = p.get("types")
    if not t:
        return "в любых экосистемах"
    if len(t[0]) == 1 and t[0][0] == "луг":
        return "в луговых экосистемах"
    return f"в полигонах типа «{t[1]}»"


def _water_names(g: gpd.GeoDataFrame) -> gpd.GeoSeries:
    """Столбец с названиями водных объектов (NAME_RU / NAME / name), если есть."""
    for c in g.columns:
        if c != "geometry" and str(c).upper() in ("NAME_RU", "NAME"):
            return g[c].astype(str).fillna("")   # пропуски → пустая строка, не NaN
    return None


def _select_water_bodies(data: Data, p: dict) -> gpd.GeoDataFrame | None:
    """Водные объекты по фильтру «только/кроме» (совпадение по части названия)."""
    if data.hydro_polys is None or data.hydro_polys.empty:
        return None
    wb = data.hydro_polys
    names = _water_names(wb)
    only, except_ = p.get("water_only"), p.get("water_except")
    if names is None:
        log.warning("  у полигональной гидрографии нет поля с названиями — фильтр водоёмов пропущен")
        return wb if only is None else None
    low = names.fillna("").str.lower()          # NaN в названиях не должен ломать фильтр
    if only:
        wb = wb[low.apply(lambda s: any(t.lower() in s for t in only))]
        log.info("  водоёмы «только %s»: %d из %d", ", ".join(only), len(wb), len(low))
    if except_ and len(wb):
        low2 = names.loc[wb.index].str.lower()
        wb = wb[~low2.apply(lambda s: any(t.lower() in s for t in except_))]
        log.info("  водоёмы «кроме %s»: исключено %d", ", ".join(except_), int((low2.apply(
            lambda s: any(t.lower() in s for t in except_))).sum()))
    return wb


def rule_type(data: Data, res: Results, measure: dict, p: dict):
    """Назначение по типу выдела; для водных мер — фильтр по названиям водоёмов.
    Участок проведения мероприятия — строго границы целевой экосистемы (выдела), а не квартал."""
    keys, label = p["types"]
    sel = data.vydels_of_type(*keys)
    water_note = ""
    wb_geom = None
    if p.get("water_only") or p.get("water_except"):
        wb = _select_water_bodies(data, p)
        if wb is None or wb.empty:
            sel = sel.iloc[0:0]
        else:
            wb_geom = _safe_union(wb.geometry.values)
            sel = sel[sel.intersects(wb_geom)]
            if p.get("water_only"):
                water_note = f"; водоёмы: только {', '.join(p['water_only'])}"
            if p.get("water_except"):
                water_note += f"; кроме {', '.join(p['water_except'])}"
    ev_geoms = []
    for _, r in sel.iterrows():
        wg = r.geometry.intersection(wb_geom) if wb_geom is not None else r.geometry
        if wg.is_empty or wg.area / 1e4 < 0.005:
            continue
        ev_geoms.append(wg)
        res.add(r, measure, f"тип {_unit_word()} «{r['Type']}» ({label}: {wg.area / 1e4:.2f} га){water_note}",
                round(float(wg.area / 1e4), 2), work_geom=wg)
    if ev_geoms:
        res.add_evidence(gpd.GeoDataFrame(geometry=ev_geoms, crs=WORK_EPSG),
                         measure["name"], f"участки работ ({label})")
    log.info("  → %d %s", len(ev_geoms), _units())


def _trail_hotspots(data: Data, window=TRAIL_WINDOW_M, q=QUANTILE):
    """Кэш: участки ≥ квантиля плотности троп в окне window (по ячейкам внутри ООЗТ)."""
    cache = data.__dict__.setdefault("_hot_cache", {})
    key = (float(window), float(q))
    if key in cache:
        return cache[key]
    if data.trails is None:
        cache[key] = (None, None, None)
        return cache[key]
    dens, tr = line_density(data.trails, data.bounds_pad, window, CELL)
    inside = data.inside_mask(tr, dens.shape)
    mask, thr = upper_quantile_mask(dens, inside, q)
    polys = polygonize_mask(mask, tr)
    log.info("  плотность троп (окно %.0f м): порог %s = %.0f м/га, %d пятен", window, _qtag(q), thr, len(polys))
    cache[key] = (polys, thr, (dens, tr))
    return cache[key]


def rule_trail_density_vydel(data: Data, res: Results, measure: dict, p: dict):
    """Плотность троп (м/га) по выделу; квантиль среди выделов всех ООЗТ заданного типа."""
    if data.trails is None:
        log.warning("  тропы не загружены — пропуск")
        return
    q = p["quantile"]
    abs_thr = p.get("abs_m_ha")
    target = _target(data, p)
    tu = _safe_union(data.trails.geometry.values)
    dens = target.geometry.intersection(tu).length / target["area_ha"]
    dens = dens.replace([np.inf], np.nan).fillna(0)
    pos = dens[dens > 0]
    if pos.empty:
        log.warning("  тропы не пересекают ни одного подходящего полигона — плотность везде 0")
        return
    if abs_thr:                                   # абсолютный порог, м/га (вместо квартиля)
        thr = float(abs_thr)
        sel = target[dens >= thr]
        log.info("  порог: абсолютный %g м/га", thr)

        def _reason(d):
            return f"плотность троп {d:.0f} м/га ≥ {thr:.0f} м/га (абсолютный порог)"
    else:
        thr = float(np.quantile(pos, q))
        sel = target[dens >= thr]

        def _reason(d):
            return f"плотность троп {d:.0f} м/га ≥ {_qtag(q)} = {thr:.0f} м/га"
    hot_polys = None
    if not sel.empty:
        hot_polys, _, _ = _trail_hotspots(data, TRAIL_WINDOW_M, q)
    hot_union = (_safe_union(hot_polys.geometry.values)
                 if hot_polys is not None and not hot_polys.empty else None)
    ev_sel = []
    for (_, r), d in zip(sel.iterrows(), dens[dens >= thr]):
        wg = r.geometry
        if data.cfg.unit_mode == "kv":
            if hot_union is not None:
                cand = wg.intersection(hot_union)
                wg = cand if not cand.is_empty and cand.area > 1.0 else tu.intersection(wg).buffer(15.0).intersection(wg)
            else:
                wg = tu.intersection(wg).buffer(15.0).intersection(wg)
        res.add(r, measure, _reason(d), round(d, 1), work_geom=wg)
        ev_sel.append(wg)
    trails_in = None
    if not sel.empty:
        res.add_evidence(gpd.GeoDataFrame(geometry=ev_sel, crs=WORK_EPSG),
                         measure["name"], "лесные участки самовосстановления покрова")
        # Красным показываем не целые исходные линии, а только их фрагменты
        # внутри локальных пятен высокой плотности и выбранных выделов.
        if hot_union is not None:
            clip_geom = hot_union.intersection(_safe_union(sel.geometry.values))
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
    log.info("  → %d %s (порог %.0f м/га)", len(sel), _units(), thr)


def rule_trail_hotspot(data: Data, res: Results, measure: dict, p: dict):
    polys, thr, _ = _trail_hotspots(data, p["window"], p["quantile"])
    if polys is None:
        log.warning("  тропы не загружены — пропуск")
        return
    target = _target(data, p)
    hit = _overlay_patches_to_vydels(data, polys, target, p["min_ha"], spec_share(p))
    pu = _safe_union(polys.geometry.values) if not hit.empty else None
    ev_geoms = []
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        wg = r.geometry.intersection(pu)
        ev_geoms.append(wg)
        res.add(r, measure, f"участки повышенной плотности троп (окно {p['window']:.0f} м, "
                            f"≥{_qtag(p['quantile'])}={thr:.0f} м/га): "
                            f"{h['patch_ha']:.2f} га ({100 * h['share']:.0f}% полигона)", round(h["patch_ha"], 2),
                work_geom=wg)
    if ev_geoms:
        res.add_evidence(gpd.GeoDataFrame(geometry=ev_geoms, crs=WORK_EPSG), measure["name"],
                         "зоны повышенной плотности троп (верхний квартиль)" if abs(p["quantile"] - 0.75) < 1e-9
                         else f"зоны повышенной плотности троп ({_qtag(p['quantile'])})")
    log.info("  → %d %s", len(hit), _units())


def rule_shore_trails(data: Data, res: Results, measure: dict, p: dict):
    """Полоса вдоль водных объектов ∩ пятна повышенной плотности троп."""
    polys, thr, _ = _trail_hotspots(data, p["window"], p["quantile"])
    if polys is None or (data.hydro_lines is None and data.hydro_polys is None):
        log.warning("  нет троп или гидрологии — пропуск")
        return
    dist = p["dist"]
    parts = [g.geometry for g in (data.hydro_lines, data.hydro_polys) if g is not None]
    water = _safe_union(pd.concat(parts).values)
    shore = water.buffer(dist).difference(water) if water.area > 0 else water.buffer(dist)
    hot_shore = polys[polys.intersects(shore)].copy()
    hot_shore["geometry"] = hot_shore.geometry.intersection(shore)
    hot_shore = hot_shore[~hot_shore.is_empty]
    hit = _overlay_patches_to_vydels(data, hot_shore, _target(data, p), p["min_ha"], spec_share(p))
    hsu = _safe_union(hot_shore.geometry.values) if not hit.empty else None
    ev_geoms = []
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        wg = r.geometry.intersection(hsu)
        ev_geoms.append(wg)
        res.add(r, measure, f"тропы повышенной плотности в {dist:.0f} м от водных объектов: "
                            f"{h['patch_ha']:.2f} га ({100 * h['share']:.0f}% полигона)", round(h["patch_ha"], 2),
                work_geom=wg)
    if ev_geoms:
        ev_shore = gpd.GeoDataFrame(geometry=ev_geoms, crs=WORK_EPSG)
        res.add_evidence(ev_shore, measure["name"],
                         f"тропы на расстоянии {dist:.0f} м от водных объектов ({_qtag(p['quantile'])})")
        # сами водные объекты рядом с найденными участками — как контекст на схеме
        near = _safe_union(ev_geoms).buffer(dist)
        if data.hydro_lines is not None:
            res.add_evidence(data.hydro_lines[data.hydro_lines.intersects(near)], measure["name"], "водные объекты (линии)")
        if data.hydro_polys is not None:
            res.add_evidence(data.hydro_polys[data.hydro_polys.intersects(near)], measure["name"], "водные объекты (полигоны)")
    log.info("  → %d %s", len(hit), _units())


# ---- виды ----------------------------------------------------------------------------
CATEGORY_LABEL = {"инвазивные": ("инвазивные виды", "инвазивных видов"),
                  "охраняемые": ("охраняемые виды", "охраняемых видов"),
                  "прочие": ("виды", "видов")}


def _sp_label(p: dict, gen=False) -> str:
    """Компактная подпись выбранных видов для легенды и слоёв-обоснований (без длинных списков латыни)."""
    cats = p.get("category") or []
    taxa = p.get("taxon") or []
    if taxa:
        if len(taxa) <= 2:
            if len(cats) == 1 and cats[0] == "инвазивные":
                return ("инвазивных видов (" if gen else "инвазивные виды (") + ", ".join(taxa) + ")"
            return ", ".join(taxa)
        # более 2 таксонов (например, 35 травянистых инвазивных видов ПП № 369-ПП): короткая подпись
        if len(cats) == 1 and cats[0] in CATEGORY_LABEL:
            base = CATEGORY_LABEL[cats[0]][1 if gen else 0]
            return f"{base} ({len(taxa)} такс.)"
        return (f"выбранных видов ({len(taxa)} такс.)" if gen else f"виды ({len(taxa)} такс.)")
    lab = CATEGORY_LABEL.get(cats[0], ("виды", "видов")) if len(cats) == 1 else ("виды", "видов")
    return lab[1] if gen else lab[0]


def _taxon_mask(species: pd.Series, taxa) -> np.ndarray:
    """Вид относится к таксону (род / вид); для полных биномов учитываются синонимы."""
    tt = list(taxa)
    try:
        syn = synonym_groups()
        for t in list(tt):
            ks = _latin_keys(t)
            if ks and len(ks[0].split()) >= 2:
                tt += [k for k in syn.get(ks[0], ()) if k not in tt]
    except Exception:
        pass
    pat = "|".join(rf"\b{re.escape(t)}\b" for t in tt)
    return species.astype(str).str.contains(pat, case=False, regex=True).values


def select_species(data: Data, p: dict):
    """Точки, полигоны и растры SDM, подходящие мероприятию (категория, группы, таксоны, тип данных)."""
    def filt(g):
        if g is None or g.empty:
            return None
        m = np.ones(len(g), dtype=bool)
        if p.get("category"):
            m &= g["category"].isin(p["category"]).values
        if p.get("kk_categories"):
            if "kk_cat" not in g.columns:           # справочник КК не загрузился
                return None
            m &= g["kk_cat"].isin(p["kk_categories"]).values
        if p.get("groups"):
            m &= g["group"].map(lambda x: group_matches(x, p["groups"])).values.astype(bool)
        if p.get("taxon"):
            m &= _taxon_mask(g["species"], p["taxon"])
        g = g[m]
        return g if not g.empty else None
    pts = filt(data.sp_points) if _uses(p, "points") else None
    polys = filt(data.sp_polys) if _uses(p, "polygons") else None
    if _uses(p, "points"):
        DIAG.step("находок видов в загруженных слоях (точки)", data.sp_points)
        DIAG.step("из них подходящих мероприятию (точки: %s)" % _sel_txt(p), pts)
        if pts is None and DIAG.cur is not None:
            t_ = "из них подходящих мероприятию (точки: %s)" % _sel_txt(p)
            DIAG.cur["steps"].append((t_ if len(t_) <= 160 else t_[:157] + "…", 0, {}))
    if _uses(p, "polygons"):
        DIAG.step("находок видов в загруженных слоях (полигоны)", data.sp_polys)
        DIAG.step("из них подходящих мероприятию (полигоны)", polys)
    rasters = []
    if _uses(p, "sdm"):
        for r in data.sp_rasters:
            if p.get("category") and r["category"] not in p["category"]:
                continue
            if p.get("kk_categories") and r.get("kk_cat") not in p["kk_categories"]:
                continue
            if not group_matches(r["group"], p.get("groups")):
                continue
            if p.get("taxon") and not _taxon_mask(pd.Series([r["species"]]), p["taxon"])[0]:
                continue
            rasters.append(r["path"])
    return pts, polys, rasters


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


def _sdm_monitoring(data: Data, res: Results, measure: dict, p: dict, rasters, label: str):
    """Модели SDM: отдельная строка «… — мониторинг (SDM ≥ порог)» по участкам ≥ порога."""
    thr = p["sdm"]
    sdm = _sdm_patches(data, rasters, thr=thr)
    if sdm is None or sdm.empty:
        log.info("  SDM: нет участков с пригодностью ≥ %s", thr)
        return
    m2 = dict(measure)
    m2["name"] = measure["name"] + " — мониторинг (SDM ≥ %.2f)" % thr
    target = _target(data, p)
    j = _safe_overlay(sdm, target[["uid", "area_ha", "geometry"]], how="intersection")
    j["ha"] = j.area / 1e4
    share = spec_share(p)
    ev_parts = []
    for uid, grp in j.groupby("uid"):
        tot = grp.groupby("species")["ha"].sum()
        tot = tot[tot >= p["min_ha"]]
        if tot.empty:
            continue
        r = data.oozt[data.oozt["uid"] == uid].iloc[0]
        if 100 * tot.sum() / max(r["area_ha"], 1e-9) < share:
            continue
        grp_ok = grp[grp["species"].isin(tot.index)]
        wg = _safe_union(grp_ok.geometry.values)
        ev_parts.append(grp_ok)
        res.add(r, m2, f"вероятность SDM ≥ {_fmt(thr)}: " + ", ".join(f"{s} ({a:.2f} га)" for s, a in tot.items()),
                round(float(wg.area / 1e4), 2), work_geom=wg)
    if ev_parts:
        res.add_evidence(gpd.GeoDataFrame(pd.concat(ev_parts, ignore_index=True), crs=WORK_EPSG),
                         m2["name"], f"{label} — SDM (≥ {thr:.2f})", "species")


def rule_species(data: Data, res: Results, measure: dict, p: dict):
    """Находки видов: локальные участки у точек / полигоны находок; SDM — мониторинг."""
    pts, polys, rasters = select_species(data, p)
    lab = _sp_label(p)
    target = _target(data, p)
    if pts is not None:
        j = gpd.sjoin(pts, target[["uid", "geometry"]], how="inner", predicate="within")
        for uid, grp in j.groupby("uid"):
            r = data.oozt[data.oozt["uid"] == uid].iloc[0]
            sp = grp.groupby("species").size().sort_values(ascending=False)
            coords = []
            for s_, g in grp.groupby("species"):
                gg = gpd.GeoSeries(g.geometry, crs=WORK_EPSG).to_crs(4326)
                coords.append(f"{s_} ({gg.x.mean():.6f}° в.д., {gg.y.mean():.6f}° с.ш.)")
            # локальный участок проведения работ вокруг точек находки (радиус 15 м внутри выдела)
            wg = _safe_union(grp.geometry.values).buffer(15.0).intersection(r.geometry)
            res.add(r, measure, f"точечные находки: {', '.join(f'{k} ({v})' for k, v in sp.items())}; "
                                f"координаты WGS 84: {'; '.join(coords)}", int(len(grp)), work_geom=wg)
        # борщевик Сосновского: на схемах — все точки в границах территории (без прореживания)
        if pts["species"].astype(str).str.contains("heracleum", case=False).any():
            ev_pts = pts[pts.within(data.union)]
            log.info("  точки борщевика в границах территории: %d (на схемах — все)", len(ev_pts))
        else:
            ev_pts = j
        res.add_evidence(ev_pts, measure["name"], f"{lab} — точки", "species")
        log.info("  точки: %d %s", j["uid"].nunique(), _units())
    if polys is not None:
        j = _safe_overlay(polys, target[["uid", "geometry"]], how="intersection")
        j["ha"] = j.area / 1e4
        for uid, grp in j.groupby("uid"):
            r = data.oozt[data.oozt["uid"] == uid].iloc[0]
            wg = _safe_union(grp.geometry.values)
            txt = ", ".join(f"{s_} ({a:.2f} га)" for s_, a in grp.groupby("species")["ha"].sum().items())
            res.add(r, measure, f"в пределах полигонов находок: {txt}", round(float(wg.area / 1e4), 2),
                    work_geom=wg)
        res.add_evidence(j, measure["name"], f"{lab} — полигоны", "species")
        log.info("  полигоны: %d %s", j["uid"].nunique(), _units())
    _sdm_monitoring(data, res, measure, p, rasters, lab)


def rule_species_trails(data: Data, res: Results, measure: dict, p: dict):
    """Тропы в заданном расстоянии от точек и внутри полигонов видов; SDM — мониторинг."""
    pts, polys, rasters = select_species(data, p)
    lab, lab_gen, dist = _sp_label(p), _sp_label(p, gen=True), p["dist"]
    trails = data.trails
    if p.get("informal_only"):
        if data.trails_informal is not None:
            trails = data.trails_informal
            log.info("  только стихийные тропы OSM (informal=yes): %d сегментов", len(trails))
        else:
            log.warning("  слой стихийных троп не загружен — используется вся дорожно-тропиночная сеть")
    target = _target(data, p)
    # ограничение по площади выдела (например, III.3 — ≥ 0,5 га, III.4 — < 0,5 га)
    amin, amax = p.get("vydel_min_ha"), p.get("vydel_max_ha")
    area_txt = ""
    if amin is not None or amax is not None:
        before = len(target)
        if amin is not None:
            target = target[target["area_ha"] > float(amin)]
        if amax is not None:
            target = target[target["area_ha"] <= float(amax)]
        area_txt = _vydel_area_txt(p, "; площадь выдела ")
        log.info("  по площади выдела: %d из %d %s", len(target), before, _units())
    zones = []
    if pts is not None:
        b = pts.copy()
        b["geometry"] = b.geometry.buffer(dist)
        b["src"] = "точки (буфер %d м)" % dist
        zones.append(b)
    if polys is not None:
        q_ = polys.copy()
        q_["src"] = "полигоны"
        zones.append(q_)
    if zones:
        z = gpd.GeoDataFrame(pd.concat(zones, ignore_index=True), crs=WORK_EPSG)
        if trails is not None:
            zu = _safe_union(z.geometry.values)
            t = trails[trails.intersects(zu)].copy()
            t["geometry"] = t.geometry.intersection(zu)
            t = t[~t.is_empty]
            if not t.empty:
                j = _safe_overlay(t.explode(index_parts=False), target[["uid", "geometry"]], how="intersection",
                                keep_geom_type=False)
                j = j[j.geom_type.isin(["LineString", "MultiLineString"])]
                j["len"] = j.length
                for uid, grp in j.groupby("uid"):
                    r = data.oozt[data.oozt["uid"] == uid].iloc[0]
                    zz = z[z.intersects(r.geometry)]
                    sp = ", ".join(sorted(zz["species"].unique())[:8])
                    frag = grp.sort_values("len", ascending=False).geometry.head(3)
                    mid = gpd.GeoSeries(frag, crs=WORK_EPSG).to_crs(4326).geometry.interpolate(0.5, normalized=True) \
                        if len(frag) else None
                    coords = "; ".join(f"({p_.y:.5f}, {p_.x:.5f})" for p_ in mid) if mid is not None else ""
                    # коридор проведения работ вдоль выявленных фрагментов троп внутри охранной зоны и выдела
                    wg = _safe_union(grp.geometry.values).buffer(10.0).intersection(zu).intersection(r.geometry)
                    res.add(r, measure, f"тропы в зоне {lab_gen} ({sp}): {grp['len'].sum():.0f} м{area_txt}"
                            + (f"; фрагменты (с.ш., в.д.): {coords}" if coords else ""),
                            round(grp["len"].sum()), work_geom=wg)
                res.add_evidence(j, measure["name"], f"тропы на расстоянии {dist:.0f} м от {lab_gen}")
                log.info("  тропы на расстоянии %.0f м от %s: %d %s", dist, lab_gen, j["uid"].nunique(), _units())
            res.add_evidence(pts[pts.intersects(data.union)] if pts is not None else None,
                             measure["name"], f"{lab} — точки", "species")
            res.add_evidence(polys, measure["name"], f"{lab} — полигоны", "species")
        else:
            # без троп — назначаем по самим зонам местообитаний
            j = _safe_overlay(z, target[["uid", "geometry"]], how="intersection")
            for uid, grp in j.groupby("uid"):
                r = data.oozt[data.oozt["uid"] == uid].iloc[0]
                wg = _safe_union(grp.geometry.values)
                res.add(r, measure, f"местообитания {lab_gen}: " + ", ".join(sorted(grp["species"].unique())[:8]),
                        round(float(wg.area / 1e4), 2), work_geom=wg)
            res.add_evidence(j, measure["name"], lab, "species")
    _sdm_monitoring(data, res, measure, p, rasters, lab)


def rule_slope(data: Data, res: Results, measure: dict, p: dict):
    if data.dem is None:
        log.warning("  ЦМР не загружена — пропуск")
        return
    threshold_pct = p["slope"]
    arr, tr = data.dem
    if getattr(data, "_slope", None) is None:
        data._slope = slope_percent(arr, CELL)
    slope = data._slope
    inside = data.inside_mask(tr, slope.shape)
    mask = inside & (np.nan_to_num(slope, nan=0) > threshold_pct)
    # убираем одиночные ячейки
    mask = ndimage.binary_opening(mask, structure=np.ones((2, 2)))
    polys = polygonize_mask(mask, tr)
    target = _target(data, p)
    hit = _overlay_patches_to_vydels(data, polys, target, p["min_ha"], spec_share(p))
    eco_label = _eco_label(p)
    pu = _safe_union(polys.geometry.values) if not hit.empty else None
    ev_geoms = []
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        wg = r.geometry.intersection(pu)
        ev_geoms.append(wg)
        res.add(r, measure, f"уклон более {threshold_pct:g}% {eco_label}: "
                           f"{h['patch_ha']:.2f} га ({h['share'] * 100:.0f} % {_unit_word()})",
                round(h["patch_ha"], 2), work_geom=wg)
    if ev_geoms:
        res.add_evidence(gpd.GeoDataFrame(geometry=ev_geoms, crs=WORK_EPSG),
                         measure["name"], f"уклон более {threshold_pct:g}% {eco_label}")
    log.info("  → %d %s (%d пятен)", len(hit), _units(), len(polys))


def rule_species_density(data: Data, res: Results, measure: dict, p: dict):
    """Плотность точек выбранных видов (например, рода Corvus) в окне; участки ≥ квантиля."""
    pts, _, _ = select_species(data, dict(p, sources=["points"]))
    lab = _sp_label(p)
    if pts is None or pts.empty:
        log.warning("  точки видов (%s) не найдены — пропуск", _sel_txt(p))
        return
    w, q = p["window"], p["quantile"]
    dens, tr = point_density(pts, data.bounds_pad, w, CELL)
    inside = data.inside_mask(tr, dens.shape)
    mask, thr = upper_quantile_mask(dens, inside, q)
    polys = polygonize_mask(mask, tr)
    target = _target(data, p)
    hit = _overlay_patches_to_vydels(data, polys, target, p["min_ha"], spec_share(p))
    pu = _safe_union(polys.geometry.values) if not hit.empty else None
    ev_geoms = []
    for _, h in hit.iterrows():
        r = data.oozt[data.oozt["uid"] == h["uid"]].iloc[0]
        wg = r.geometry.intersection(pu)
        ev_geoms.append(wg)
        res.add(r, measure, f"плотность {lab} (окно {w:.0f} м) ≥ {_qtag(q)} = {thr:.3f} шт/га: "
                            f"{h['patch_ha']:.2f} га ({100 * h['share']:.0f}% полигона)", round(h["patch_ha"], 2),
                work_geom=wg)
    if ev_geoms:
        res.add_evidence(gpd.GeoDataFrame(geometry=ev_geoms, crs=WORK_EPSG),
                         measure["name"], f"плотность {lab} ({_qtag(q)})")
    res.add_evidence(pts[pts.intersects(data.union.buffer(w))], measure["name"], f"{lab} — точки", "species")
    log.info("  → %d %s (%d точек: %s)", len(hit), _units(), len(pts), lab)


def rule_road_buffer(data: Data, res: Results, measure: dict, p: dict):
    """Назначение полигонам, пересекающим полосу между двумя расстояниями от дорожной сети.

    Проверка и слой-основание строятся по каждому полигону отдельно: буферизуется только
    участок дорожной сети рядом с полигоном (иначе буфер всей сети переполняет память)."""
    if data.trails is None:
        log.warning("  дорожная сеть не загружена — пропуск")
        return
    dmin, dmax = float(p["dist_min"]), float(p["dist_max"])
    target = _target(data, p)
    if target.empty:
        return
    roads = data.trails
    sidx = roads.sindex
    sel_rows, ev_geoms = [], []
    for _, r in target.iterrows():
        g = r.geometry
        gb = g.buffer(dmax + 2.0)
        near_idx = list(sidx.query(gb, predicate="intersects"))
        if not near_idx:
            continue                                   # дороги дальше dmax — полосу не пересекает
        segs = [roads.geometry.values[i].intersection(gb) for i in near_idx]
        segs = [sg for sg in segs if sg is not None and not sg.is_empty]
        if not segs:
            continue
        outer = _safe_union([sg.buffer(dmax, quad_segs=8) for sg in segs])
        if dmin > 0:
            inner = _safe_union([sg.buffer(dmin, quad_segs=8) for sg in segs])
            if inner.covers(g):
                continue            # полигон целиком в ближней зоне — до полосы не достаёт
            annulus = outer.difference(inner)
        else:
            annulus = outer
        wg = annulus.intersection(g)
        if wg.is_empty or wg.area / 1e4 < 0.01:
            continue
        sel_rows.append((r, wg))
        ev_geoms.append(wg)
    band = (f"полоса {dmin:.0f}–{dmax:.0f} м от дорожной сети" if dmax > dmin > 0
            else f"полоса до {dmax:.0f} м от дорожной сети" if dmin <= 0
            else f"полоса от {dmin:.0f} м от дорожной сети")
    for r, wg in sel_rows:
        res.add(r, measure, f"пересечение с {band} ({wg.area / 1e4:.2f} га)",
                round(float(wg.area / 1e4), 2), work_geom=wg)
    if ev_geoms:
        ev = gpd.GeoDataFrame(geometry=ev_geoms, crs=WORK_EPSG)
        res.add_evidence(ev, measure["name"], band)
    log.info("  %s: → %d %s", band, len(sel_rows), _units())


def rule_drainage(data: Data, res: Results, measure: dict, p: dict):
    """Дренажные канавы: полигоны, которые канавы пересекают; участок работ — коридор вдоль канав."""
    if data.drainage is None or data.drainage.empty:
        log.warning("  %s: слой дренажных канав не загружен — пропуск", measure["name"])
        return
    target = _target(data, p)
    if target is None or target.empty:
        return
    min_m = float(p.get("min_m") or 0.0)
    d = data.drainage
    n = 0
    for _, r in target.iterrows():
        seg = gpd.clip(d, r.geometry)
        if seg is None or seg.empty:
            continue
        ln = float(seg.length.sum())
        if ln < max(min_m, 1.0):
            continue
        wg = _safe_union(seg.geometry.values).buffer(5.0).intersection(r.geometry)
        res.add(r, measure, f"дренажные канавы: {_fmt(ln)} м внутри полигона", round(ln, 1), work_geom=wg)
        res.add_evidence(seg, measure["name"], "дренажные канавы")
        n += 1
    log.info("  дренажные канавы ≥ %.0f м: → %d %s", min_m, n, _units())


def rule_wetland(data: Data, res: Results, measure: dict, p: dict):
    """Болота: полигоны, пересекающие загруженный вектор болот; участок работ — полигон болота."""
    if data.wetlands is None or data.wetlands.empty:
        log.warning("  %s: слой болот не загружен — пропуск", measure["name"])
        return
    target = _target(data, p)
    if target is None or target.empty:
        return
    try:
        j = _safe_overlay(target[["uid", "geometry"]], data.wetlands[["geometry"]],
                        how="intersection", keep_geom_type=True)
    except Exception as e:
        log.error("  оверлей с болотами: %s", e)
        return
    if j.empty:
        log.warning("  болота не пересекают ни одного подходящего полигона")
        return
    j["_ha"] = j.area / 1e4
    j = j[j["_ha"] >= 0.01]          # касания с нулевой площадью не считаются болотом
    if j.empty:
        log.warning("  болота только касаются полигонов (пересечение < 0,01 га)")
        return
    for uid, grp in j.groupby("uid"):
        r = data.oozt[data.oozt["uid"] == uid].iloc[0]
        wg = _safe_union(grp.geometry.values)
        ha = float(wg.area / 1e4)
        res.add(r, measure, f"болота/заболоченные участки: {_fmt(ha)} га внутри полигона",
                round(ha, 2), work_geom=wg)
    res.add_evidence(j, measure["name"], "болота (вектор)")
    log.info("  болота: → %d %s", j["uid"].nunique(), _units())


def rule_water_catalog(data: Data, res: Results, measure: dict, p: dict):
    """Водоёмы из справочника: выбор по названию; участок проведения работ — контур самого водоёма."""
    cat = data.water_catalog()
    if cat is None or cat.empty:
        log.warning("  %s: водные полигоны не загружены — пропуск", measure["name"])
        return
    only = p.get("water_only")
    if only:
        low = cat["Водоём"].astype(str).str.lower()
        sel = cat[low.apply(lambda s: any(t.lower() in s for t in only))]
        log.info("  водоёмы «только %s»: %d из %d", ", ".join(only), len(sel), len(cat))
    else:
        sel = cat
        log.info("  водоёмы не отфильтрованы: %d из справочника", len(sel))
    if sel.empty:
        log.warning("  после фильтра по названиям водоёмов не осталось ни одного")
        return
    target = _target(data, p)
    if target is None or target.empty:
        return
    for _, row in sel.iterrows():
        cand = target[target.geometry.intersects(row.geometry)]
        if cand.empty:
            continue
        a = cand.geometry.intersection(row.geometry).area
        uid = cand.loc[a.idxmax(), "uid"]
        r = data.oozt[data.oozt["uid"] == uid].iloc[0]
        wg = row.geometry.intersection(data.union) if data.union is not None else row.geometry
        if wg.is_empty:
            wg = row.geometry
        typ = f", тип: {row['Тип']}" if str(row["Тип"]) not in ("", "не указан") else ""
        res.add(r, measure, f"водоём «{row['Водоём']}»{typ}: {_fmt(row['Площадь, га'])} га",
                float(row["Площадь, га"]), work_geom=wg)
        res.add_evidence(gpd.GeoDataFrame(geometry=[wg], crs=WORK_EPSG),
                         measure["name"], "водные объекты — выбранные")
    log.info("  водоёмы из справочника: %d шт.", len(sel))


# ---- болота и пруды на схемах и в справочнике -----------------------------------------------
_WETLAND_RE = r"wetland|болот|marsh|bog|swamp|fen\b|топь"


def _wetland_flag(g: gpd.GeoDataFrame | None) -> pd.Series:
    """Полигоны гидрографии с признаком болота (natural=wetland, тип «болото» и т. п.)."""
    if g is None or g.empty:
        return pd.Series(dtype=bool)
    cols = [c for c in g.columns if c != "geometry" and str(c).lower() in
            ("natural", "wetland", "water", "type", "тип", "landuse", "fclass", "category", "категория", "class")]
    m = pd.Series(False, index=g.index)
    for c in cols:
        m |= g[c].astype(str).str.lower().str.contains(_WETLAND_RE, regex=True, na=False)
    return m


def wetland_polys(data) -> gpd.GeoDataFrame | None:
    """Болота: слой болот (вектор) + полигоны гидрографии с признаком болота."""
    parts = []
    if getattr(data, "wetlands", None) is not None and not data.wetlands.empty:
        w = data.wetlands[data.wetlands.geom_type.isin(["Polygon", "MultiPolygon"])]
        parts.append(w[["geometry"]])
    hp = getattr(data, "hydro_polys", None)
    if hp is not None and not hp.empty:
        f = _wetland_flag(hp)
        if f.any():
            parts.append(hp.loc[f, ["geometry"]])
    if not parts:
        return None
    return gpd.GeoDataFrame(pd.concat(parts, ignore_index=True), crs=WORK_EPSG)


def pond_polys(data) -> gpd.GeoDataFrame | None:
    """Пруды и озёра: полигоны гидрографии без признака болота."""
    hp = getattr(data, "hydro_polys", None)
    if hp is None or hp.empty:
        return hp
    return hp[~_wetland_flag(hp)]


WETLAND_COLOR = "#B4D796"     # болота на подложке схем (бледно-зелёный, RGB 180 215 150)


# ---- I.8, I.9, I.11, III.7: объекты справочника водоёмов и болот ----------------------------
def _wr_find_polygon(pt, layers):
    """Контур для точки: полигон, содержащий точку (наименьший), иначе ближайший ≤ WATER_REF_SNAP_M."""
    for L in layers:
        if L is None or L.empty:
            continue
        cand = L[L.geometry.contains(pt)]
        if not cand.empty:
            i = cand.geometry.area.idxmin()
            return L, i, 0.0
    best = None
    for L in layers:
        if L is None or L.empty:
            continue
        near = L[L.geometry.intersects(pt.buffer(WATER_REF_SNAP_M))]
        if near.empty:
            continue
        d = near.geometry.distance(pt)
        i = d.idxmin()
        if best is None or d[i] < best[2]:
            best = (L, i, float(d[i]))
    return best if best is not None else (None, None, None)


def _wr_write_unmatched(data):
    rows = data.__dict__.get("_wr_unmatched") or []
    out = Path(data.cfg.out_dir) if getattr(data.cfg, "out_dir", "") else None
    if out is None or not rows:
        return
    try:
        out.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(rows).to_csv(out / "water_ref_unmatched.csv", index=False, encoding="utf-8-sig")
    except Exception as e:
        log.warning("  water_ref_unmatched.csv не записан: %s", e)


def rule_water_ref(data: Data, res: Results, measure: dict, p: dict):
    """Объекты справочника водоёмов и болот с отметкой мероприятия; участок работ — контур объекта ∩ ООЗТ."""
    code = str(measure.get("code") or "").strip()
    ref = water_ref_marked(code)
    if ref is None:
        log.warning("  %s: справочник водоёмов и болот не найден — пропуск", measure["name"])
        return
    if ref.empty:
        log.warning("  %s: нет отметок в справочнике водоёмов (столбец %s) — мероприятие не назначается",
                    measure["name"], code or "?")
        return
    pts = gpd.GeoDataFrame(ref.reset_index(drop=True),
                           geometry=gpd.points_from_xy(ref["lon"], ref["lat"]), crs=4326).to_crs(WORK_EPSG)
    near = pts.geometry.intersects(data.union.buffer(WATER_REF_SNAP_M))
    n_out = int((~near).sum())
    pts = pts[near]
    if n_out:
        log.info("  справочник: %d отмеченных объектов вне территории расчёта — не рассматриваются", n_out)
    if pts.empty:
        log.warning("  %s: в границах территории нет объектов, отмеченных в справочнике (%s)",
                    measure["name"], code)
        return
    hp = data.hydro_polys if data.hydro_polys is not None and not data.hydro_polys.empty else None
    wl = wetland_polys(data)
    stats: dict = {}
    geo_cnt: dict = {str(x): 0 for x in pd.unique(data.oozt["OOZT"])}
    unmatched = data.__dict__.setdefault("_wr_unmatched", [])
    seen: dict = {}
    ev_w, ev_b = [], []
    for _, o in pts.iterrows():
        is_bog = _norm_txt(o["cat"]).startswith("болот") or _norm_txt(o["otype"]).startswith("болот")
        lbl = f"«{o['name']}» ({o['oozt']}, кв. {o['kv'] or '—'}; лист «{o['sheet']}», строка {o['row']})"
        if code == "I.8" and not is_bog:
            log.warning("  %s: непрофильная категория «%s» для I.8 (ожидается болото)", lbl, o["cat"] or "—")
        if code in ("I.9", "I.11", "III.7") and is_bog:
            log.warning("  %s: непрофильная категория «Болото» для %s", lbl, code)
        layers = [wl, hp] if is_bog else [hp]
        L, idx, dist = _wr_find_polygon(o.geometry, layers)
        st_key = str(o["oozt"]) or "—"
        stt = stats.setdefault(st_key, dict(marked=0, matched=0, unmatched=0, dropped=0))
        stt["marked"] += 1
        if L is None:
            stt["unmatched"] += 1
            unmatched.append(dict(Мероприятие=f"{code} {measure['name']}".strip(), ООЗТ=o["oozt"],
                                  Квартал=o["kv"], Название=o["name"], Категория=o["cat"], Тип=o["otype"],
                                  **{"Площадь, га": o["area"], "Широта": o["lat"], "Долгота": o["lon"],
                                     "Лист": o["sheet"], "Строка Excel": o["row"],
                                     "Причина": f"нет полигона {'болот/гидрографии' if is_bog else 'гидрографии'} "
                                                f"под точкой и ближе {WATER_REF_SNAP_M:.0f} м"}))
            log.warning("  %s: контур не найден (точка %.6f, %.6f) — в water_ref_unmatched.csv",
                        lbl, o["lat"], o["lon"])
            continue
        geom = L.geometry.loc[idx]
        if not geom.is_valid:
            geom = make_valid(geom)
        gkey = (id(L), idx)
        if gkey in seen:                     # тот же контур уже назначен (дубль в справочнике)
            log.info("  %s: контур совпадает с «%s» — одна запись на объект", lbl, seen[gkey])
            stt["matched"] += 1
            continue
        if dist:
            log.info("  %s: привязан к ближайшему контуру (%.0f м)", lbl, dist)
        g_ha = geom.area / 1e4
        if pd.notna(o["area"]) and o["area"] > 0 and g_ha > 0 and max(g_ha / o["area"], o["area"] / g_ha) > 2:
            log.warning("  %s: площадь по справочнику %.3f га, контура — %.3f га (расхождение более 2 раз)",
                        lbl, o["area"], g_ha)
        wg = geom.intersection(data.union)
        if wg.geom_type == "GeometryCollection":
            pl = [x for x in wg.geoms if x.geom_type in ("Polygon", "MultiPolygon")]
            wg = _safe_union(pl) if pl else wg
        if wg.is_empty or wg.area < 1.0:
            stt["unmatched"] += 1
            unmatched.append(dict(Мероприятие=f"{code} {measure['name']}".strip(), ООЗТ=o["oozt"],
                                  Квартал=o["kv"], Название=o["name"], Категория=o["cat"], Тип=o["otype"],
                                  **{"Площадь, га": o["area"], "Широта": o["lat"], "Долгота": o["lon"],
                                     "Лист": o["sheet"], "Строка Excel": o["row"],
                                     "Причина": "контур вне границ территории"}))
            continue
        cand = data.oozt[data.oozt.geometry.intersects(wg)]
        if cand.empty:
            stt["unmatched"] += 1
            continue
        a = cand.geometry.intersection(wg).area
        r = cand.loc[a.idxmax()]
        oo_geo = str(r["OOZT"])
        if o["oozt"] and not _oozt_key_match(o["oozt"], oo_geo):
            log.warning("  %s: в справочнике ООЗТ «%s», по геометрии — «%s» (взято по геометрии)",
                        lbl, o["oozt"], oo_geo)
        mkv = re.match(r"^\s*(\d+)\s*(?:\((.*?)\))?", str(o["kv"]))
        if mkv and mkv.group(2) and "терлец" in _norm_txt(mkv.group(2)) \
                and int(mkv.group(1)) > WATER_REF_TERLETSKY_MAX_KV:
            log.warning("  %s: в Терлецком массиве %d квартала — номер %s сомнителен",
                        lbl, WATER_REF_TERLETSKY_MAX_KV, mkv.group(1))
        if mkv and str(r["KV"]).removesuffix(".0") != mkv.group(1):
            log.info("  %s: квартал по справочнику %s, по геометрии — %s (%s)", lbl, o["kv"], r["KV"], oo_geo)
        if oozt_flag(measure, oo_geo) is False:
            stt["dropped"] += 1
            continue
        area_txt = _fmt(o["area"], 3) if pd.notna(o["area"]) else _fmt(g_ha, 3)
        kind = o["otype"] or o["cat"] or ("болото" if is_bog else "водоём")
        basis = (f"по справочнику водоёмов и болот: «{o['name']}», {kind}, {area_txt} га; квартал {o['kv'] or '—'}"
                 + (f"; {o['note']}" if o["note"] else ""))
        ok = res.add(r, measure, basis, round(float(o["area"] if pd.notna(o["area"]) else g_ha), 3),
                     work_geom=wg, obj=f"{o['sheet']}:{o['row']}")
        if ok:
            stt["matched"] += 1
            geo_cnt[oo_geo] = geo_cnt.get(oo_geo, 0) + 1
            seen[gkey] = o["name"]
            (ev_b if is_bog else ev_w).append(dict(geometry=wg, label=o["name"], oozt=oo_geo))
        else:
            stt["dropped"] += 1
    if ev_w:
        res.add_evidence(gpd.GeoDataFrame(ev_w, crs=WORK_EPSG), measure["name"], "водные объекты — выбранные", "label")
    if ev_b:
        res.add_evidence(gpd.GeoDataFrame(ev_b, crs=WORK_EPSG), measure["name"], "болота — по справочнику", "label")
    for oo, v in stats.items():
        log.info("  справочник %s, %s: отмечено %d, привязано %d, не найдено %d, отброшено флагами %d",
                 code, oo, v["marked"], v["matched"], v["unmatched"], v["dropped"])
    log.info("  назначено по ООЗТ, определённой по геометрии: %s",
             ", ".join(f"{k} — {v}" for k, v in geo_cnt.items()))
    DIAG.step_counts("объектов справочника, привязанных к контурам (ООЗТ по геометрии)", geo_cnt)
    _wr_write_unmatched(data)


# ---- I.7: мозаичное выкашивание по структурному классу ----------------------------------------
def rule_struct_class(data: Data, res: Results, measure: dict, p: dict):
    """Полигоны заданных структурных классов; участок работ — полигон экосистемы ∩ ООЗТ."""
    classes = list(p.get("struct_classes") or STRUCT_MOW_DEFAULT)
    want = {_norm_txt(c): c for c in classes}
    g = data.oozt
    has_struct = "Struct" in g.columns and g["Struct"].astype(str).str.strip().ne("").any()
    if not has_struct:
        log.info("  поле «Структурный класс» не задано — класс определяется по полям типа и названия")
    n = 0
    for _, r in g.iterrows():
        vals = [r["Struct"]] if has_struct else [struct_class_of(r["Type"]), struct_class_of(r["Name"])]
        hit = next((v for v in vals if _norm_txt(v) in want), None)
        if hit is None:
            continue
        wg = r.geometry.intersection(data.union)
        if wg.is_empty:
            continue
        if res.add(r, measure, f"структурный класс «{hit}» ({wg.area / 1e4:.2f} га)",
                   round(float(wg.area / 1e4), 2), work_geom=wg):
            n += 1
    log.info("  структурные классы (%s): → %d %s", ", ".join(classes), n, _units())


# ---- III.3 / III.4: стихийные тропы в выделах с видами Красной книги ---------------------------
def rule_rare_trail_vydel(data: Data, res: Results, measure: dict, p: dict):
    """Выделы с плотностью стихийных троп > порога и находками видов КК Москвы заданных категорий."""
    trails = data.trails_informal if p.get("informal_only", True) else data.trails
    if trails is None or trails.empty:
        log.warning("  %s: слой %s не загружен — пропуск", measure["name"],
                    "стихийных троп (OSM informal=yes или свой слой)" if p.get("informal_only", True) else "троп")
        return
    p_sel = dict(p, groups=None, sources=[s for s in (p.get("sources") or ["points", "polygons"]) if s != "sdm"]
                 or ["points", "polygons"])
    pts, polys, _ = select_species(data, p_sel)
    groups = p.get("groups")

    def _plants(g):
        # сосудистые растения: группа слоя или вид из справочника цветения охраняемых растений
        if g is None or g.empty or not groups:
            return g
        fl = flowering_ref()["months"]
        m = g["group"].map(lambda x: group_matches(x, groups)).astype(bool) | \
            g["species"].map(lambda sp: _lookup_latin(sp, fl)[0] is not None).astype(bool)
        g = g[m]
        return g if not g.empty else None
    pts, polys = _plants(pts), _plants(polys)
    if pts is None and polys is None:
        log.warning("  %s: нет находок видов КК Москвы категорий %s", measure["name"],
                    ", ".join(str(c) for c in p.get("kk_categories") or []))
        return
    thr = float(p.get("abs_m_ha") or 100.0)
    target = data.oozt
    amin, amax = p.get("vydel_min_ha"), p.get("vydel_max_ha")
    if amin is not None:
        target = target[target["area_ha"] > float(amin)]
    if amax is not None:
        target = target[target["area_ha"] <= float(amax)]
    DIAG.step("выделов всего", data.oozt)
    if target.empty:
        log.warning("  нет выделов%s", _vydel_area_txt(p, " площадью "))
        return
    DIAG.step("выделов" + _vydel_area_txt(p, " площадью "), target)
    if pts is not None:
        DIAG.step("находок растений КК (точки)", pts, col=None)
    if polys is not None:
        DIAG.step("находок растений КК (полигоны)", polys, col=None)
    tu = _safe_union(trails.geometry.values)
    dens = (target.geometry.intersection(tu).length / target["area_ha"]).replace([np.inf], np.nan).fillna(0)
    sel = target[dens > thr]
    log.info("  выделов%s: %d; плотность стихийных троп > %g м/га — %d", _vydel_area_txt(p, " площадью "),
             len(target), thr, len(sel))
    for oo, dd in dens.groupby(target["OOZT"]):
        log.info("    %s: плотность стихийных троп в выделах — медиана %.0f, макс. %.0f м/га; выше порога %g — %d",
                 oo, float(dd.median()), float(dd.max()), thr, int((dd > thr).sum()))
    DIAG.step("выделов с плотностью троп > %g м/га" % thr, sel)
    n, ev_t, ev_p = 0, [], []
    n_with, oo_with = 0, {}
    for (_, r), d in zip(sel.iterrows(), dens[dens > thr]):
        finds = []
        if pts is not None:
            fp = pts[pts.geometry.within(r.geometry)]
            finds.append(fp)
            if not fp.empty:
                ev_p.append(fp)
        if polys is not None:
            fq = polys[polys.geometry.intersects(r.geometry)]
            fq = fq[fq.geometry.intersection(r.geometry).area > 1.0]
            finds.append(fq)
        f = pd.concat([x for x in finds if x is not None and not x.empty], ignore_index=True) \
            if any(x is not None and not x.empty for x in finds) else None
        if f is None or f.empty:
            continue
        n_with += 1
        oo_with[r["OOZT"]] = oo_with.get(r["OOZT"], 0) + 1
        sp = f.groupby("species")["kk_cat"].first()
        season = None
        if p.get("season_flower"):
            ms, miss = set(), []
            for s_ in sp.index:
                m_ = flowering_months(s_)
                (ms.update(m_) if m_ else miss.append(s_))
            if miss:
                log.warning("  нет сроков цветения в справочнике: %s", ", ".join(miss[:6]))
            season = months_text(ms) if ms else None
        sp_txt = ", ".join(f"{s_} (кат. {c})" for s_, c in sp.items())
        basis = (f"виды КК Москвы: {sp_txt}; плотность стихийных троп {d:.0f} м/га > {thr:.0f} м/га; "
                 f"площадь выдела {r['area_ha']:.2f} га" + (f"; сроки по цветению: {season}" if season else ""))
        if res.add(r, measure, basis, round(float(d), 1), work_geom=r.geometry, season=season):
            n += 1
            ev_t.append(r.geometry)
    if ev_t:
        tt = trails[trails.intersects(_safe_union(ev_t))].copy()
        tt["geometry"] = tt.geometry.intersection(_safe_union(ev_t))
        res.add_evidence(tt[~tt.is_empty], measure["name"], "тропы в зонах повышенной плотности")
        if ev_p:
            res.add_evidence(gpd.GeoDataFrame(pd.concat(ev_p, ignore_index=True), crs=WORK_EPSG),
                             measure["name"], "редкие виды — точки", "species")
    DIAG.step_counts("из них с находкой растения КК", oo_with)
    if len(sel) and not n_with:
        log.warning("  ни в одном выделе с плотностью троп > %g м/га нет находок растений КК", thr)
    log.info("  → %d %s", n, _units())


def rule_points_category(data: Data, res: Results, measure: dict, p: dict):
    """Точки по категории (лесопатология): выделы леса, содержащие точки ПВН заданной категории."""
    if data.lesopat is None or data.lesopat.empty:
        log.warning("  %s: слой лесопатологии (ПВН) не загружен — пропуск", measure["name"])
        return
    cat = str(p.get("cat_value") or "").strip().lower().replace("ё", "е")
    pts = data.lesopat[data.lesopat["cat"].astype(str).str.lower().replace("ё", "е") == cat]
    if pts.empty:
        log.info("  нет точек категории «%s»", cat)
        return
    target = _target(data, p)
    if target is None or target.empty:
        return
    j = gpd.sjoin(pts, target[["uid", "geometry"]], how="left", predicate="within")
    # точки в «дырах» или на границе выдела: ближайший целевой полигон в пределах 50 м
    miss = j[j["uid"].isna()]
    if not miss.empty and len(target):
        try:
            (inp, tree), dists = target.sindex.nearest(
                miss.geometry.values, max_distance=50, return_distance=True)
            for k in np.argsort(dists):
                i, t = int(inp[k]), int(tree[k])
                uid_val = target.iloc[t]["uid"]
                if pd.isna(j.at[miss.index[i], "uid"]):
                    j.at[miss.index[i], "uid"] = uid_val
        except Exception as e:
            log.warning("  ближайший полигон для %d точек не найден: %s", len(miss), e)
    j = j[j["uid"].notna()].copy()
    if j.empty:
        return
    uid_oo = dict(zip(data.oozt["uid"], data.oozt["OOZT"].astype(str)))
    uid_kv = dict(zip(data.oozt["uid"], data.oozt["KV"].astype(str)))
    j["_oozt"] = j["uid"].map(uid_oo).fillna("")
    j["_kv"] = j["uid"].map(uid_kv).fillna("")
    wg_rows = []
    for uid, grp in j.groupby("uid"):
        r = data.oozt[data.oozt["uid"] == uid].iloc[0]
        ids = ", ".join(dict.fromkeys(x for x in grp["pvn"] if x and x not in ("nan", "None")))
        gg = gpd.GeoSeries(grp.geometry, crs=WORK_EPSG).to_crs(4326)
        coords = "; ".join(f"({lo:.6f}° в.д., {la:.6f}° с.ш.)" for lo, la in zip(gg.x, gg.y))
        base = (f"точки ПВН ({len(grp)} шт.)" + (f": № {ids}" if ids else "") +
                f"; координаты WGS 84: {coords}")
        # в режиме готовых кварталов (без выделов экосистем) ограничиваем участок 35-м буфером вокруг точки ПВН
        wg = (_safe_union(grp.geometry.values).buffer(35.0).intersection(r.geometry)
              if data.cfg.unit_mode == "kv"
              else r.geometry)
        res.add(r, measure, base, int(len(grp)), work_geom=wg)
        if wg is not None and not wg.is_empty:
            wg_rows.append(wg)
    if wg_rows:
        res.add_evidence(gpd.GeoDataFrame(geometry=wg_rows, crs=WORK_EPSG),
                         measure["name"], "участки лесопатологических мероприятий (ПВН)")
    res.add_evidence(j, measure["name"], f"лесопатология — {cat}", "pvn")
    log.info("  категория «%s»: %d точек → %d %s", cat, len(pts), j["uid"].nunique(), _units())


def _outer_shell(geom):
    """Внешний контур ООЗТ без внутренних мелких дыр (чтобы внутренний буфер 500–1000 м шёл только от внешней границы парка)."""
    if geom is None or geom.is_empty:
        return geom
    from shapely.geometry import Polygon as _Poly
    parts = list(geom.geoms) if hasattr(geom, "geoms") else [geom]
    shells = [_Poly(p.exterior) for p in parts if p.geom_type == "Polygon" and not p.is_empty]
    return _safe_union(shells) if shells else geom


# «Колокольчик»: местообитания → список высаживаемых видов Красной книги Москвы
KOLOKOL_HABITATS = [
    (("широколиствен", "дубрав"), "широколиственный лес",
     "первоцветы (Primula veris и др.) и колокольчик широколиственный (Campanula latifolia)"),
    (("луг",), "луга естественные",
     "колокольчик персиколистный (Campanula persicifolia) и колокольчик раскидистый (Campanula patula)"),
    (("болот", "пониж", "западин", "влажн", "сыр"), "понижения и западины",
     "купальница (Trollius europaeus) и калужница (Caltha palustris)"),
]
PAVED_ROAD_WORDS = ("asphalt", "paved", "motorway", "trunk", "primary", "secondary", "tertiary",
                    "residential", "unclassified", "living_street", "service", "road")


def _paved_roads(data: Data):
    """Дороги с твёрдым покрытием: отдельный слой либо линии сети с признаком асфальта."""
    if data.roads_paved is not None and not data.roads_paved.empty:
        return data.roads_paved
    t = data.trails
    if t is None or t.empty:
        return None
    for c in t.columns:
        if c == "geometry":
            continue
        cl = str(c).lower()
        if cl in ("surface", "highway", "fclass", "type", "road_type", "class"):
            vals = t[c].astype(str).str.lower()
            mask = vals.apply(lambda s: any(w in s for w in PAVED_ROAD_WORDS))
            if mask.any() and not mask.all():
                log.info("  дороги с твёрдым покрытием: %d из %d линий сети (по полю «%s»)",
                         int(mask.sum()), len(t), c)
                return t[mask]
    return None


def rule_kolokolchik(data: Data, res: Results, measure: dict, p: dict):
    """«Колокольчик»: высаживание КК видов в типичных местообитаниях.

    Полигоны проведения мероприятия строго очерчены географическими и экологическими условиями
    (пересечение кольца 500–1000 м от внешней границы ООЗТ и/или 100-м полосы от дорог с твёрдым
    покрытием с конкретными выделами широколиственных лесов, естественных лугов и понижений/западин;
    целые кварталы никогда не назначаются)."""
    target = _target(data, p)
    if target is None or target.empty:
        return
    lo, hi = p.get("ring") or (500.0, 1000.0)
    rd = float(p.get("road_dist") or 100.0)
    # 1) кольцо 500–1000 м от внешней границы каждой ООЗТ внутрь территории
    rings: dict[str, object] = {}
    for oo, geom in data.oozt.dissolve(by="OOZT").geometry.items():
        shell = _outer_shell(geom)
        mid = shell.buffer(-lo)
        if mid.is_empty:
            continue
        inner = shell.buffer(-hi)
        ring = (mid if inner.is_empty else mid.difference(inner)).intersection(geom)
        if not ring.is_empty and ring.area > 0:
            rings[str(oo)] = ring
    # 2) полоса от дорог с твёрдым покрытием (локально по пространственному индексу около каждого выдела)
    roads = _paved_roads(data)
    r_sidx = roads.sindex if roads is not None and not roads.empty else None
    if r_sidx is None:
        log.warning("  слой дорог с твёрдым покрытием не задан — полоса от дорог не учитывается")
    ev_rows = []
    for _, r in target.iterrows():
        s = f"{r.get('Type', '')} {r.get('Name', '')}".lower().replace("ё", "е")
        hab = next(((nm, sp) for keys, nm, sp in KOLOKOL_HABITATS
                    if ("луг" not in keys or ("луг" in s and not _anthro_meadow(s)))
                    and any(k in s for k in keys)), None)
        if hab is None:
            continue
        g = r.geometry
        zones, z_geoms = [], []
        ring = rings.get(str(r["OOZT"]))
        if ring is not None:
            zr = ring.intersection(g)
            if not zr.is_empty and zr.area / 1e4 >= 0.01:
                zones.append(f"кольцо {lo:.0f}–{hi:.0f} м от границы парка ({zr.area / 1e4:.2f} га)")
                z_geoms.append(zr)
        if r_sidx is not None:
            gb = g.buffer(rd + 2.0)
            near_idx = list(r_sidx.query(gb, predicate="intersects"))
            if near_idx:
                segs = [roads.geometry.values[i].intersection(gb) for i in near_idx]
                segs = [sg for sg in segs if sg is not None and not sg.is_empty]
                if segs:
                    zs = _safe_union([sg.buffer(rd, quad_segs=8) for sg in segs]).intersection(g)
                    if not zs.is_empty and zs.area / 1e4 >= 0.01:
                        zones.append(f"полоса {rd:.0f} м от дорог с твёрдым покрытием ({zs.area / 1e4:.2f} га)")
                        z_geoms.append(zs)
        if not z_geoms:
            continue
        wg = _safe_union(z_geoms)
        ha = float(wg.area / 1e4)
        if ha < 0.01:
            continue
        zones.append("по границе выдела")
        res.add(r, measure, f"зоны: {'; '.join(zones)}. Местообитание: {hab[0]} — высадить: {hab[1]}",
                round(ha, 2), work_geom=wg)
        ev_rows.append({"hab": f"{hab[0]}: {hab[1]}", "geometry": wg})
    if ev_rows:
        res.add_evidence(gpd.GeoDataFrame(ev_rows, crs=WORK_EPSG),
                         measure["name"], "участки высаживания («Колокольчик»)", "hab")
    log.info("  «Колокольчик»: кольцо %s–%s м, полоса %s м от дорог → %d участков (%.1f га)",
             lo, hi, rd, len(ev_rows),
             sum(x["geometry"].area for x in ev_rows) / 1e4 if ev_rows else 0.0)


RULE_FUNCS = {
    "type": rule_type, "trail_density": rule_trail_density_vydel, "trail_hotspot": rule_trail_hotspot,
    "shore": rule_shore_trails, "slope": rule_slope, "species": rule_species,
    "species_trails": rule_species_trails, "species_density": rule_species_density,
    "road_buffer": rule_road_buffer, "drainage": rule_drainage, "wetland": rule_wetland,
    "water_catalog": rule_water_catalog, "points_category": rule_points_category,
    "kolokolchik": rule_kolokolchik, "water_ref": rule_water_ref, "struct_class": rule_struct_class,
    "rare_trail_vydel": rule_rare_trail_vydel,
}


def dispatch_rule(measure: dict, data: Data, res: Results) -> bool:
    spec = measure_spec(measure)
    if spec["rule"] is None:
        if spec["reason"] not in ("метод выявления не задан", "по таблице: нет данных"):
            log.warning("  %s: %s — пропуск", measure["name"], spec["reason"])
        return False
    log.info("Мероприятие: %s%s", (measure.get("code") + " ") if measure.get("code") else "", measure["name"])
    st = measure_status(measure, data.cfg)
    if st["state"] == "missing":
        log.warning("  %s — пропуск", st["text"])
        return False
    RULE_FUNCS[spec["rule"]](data, res, measure, spec["p"])
    return True


def spec_pad_m(measures) -> float:
    """Запас вокруг ООЗТ для расчётов: наибольшее окно или расстояние среди включённых мероприятий."""
    v = 0.0
    for m in measures or []:
        if m.get("on"):
            p = measure_spec(m).get("p", {})
            v = max([v] + [float(p[k]) for k in ("window", "dist", "dist_max") if p.get(k)])
    return v


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
    code_re = re.compile(r"^[IVX]+\.\d+$")
    c_code = next((c for c in df.columns
                   if df[c].dropna().astype(str).str.strip().str.match(code_re).mean() > 0.6), None)
    spec_cols = {}
    for key, title, stem, _ in SPEC_COLUMNS:
        c = next((c for c in df.columns if re.sub(r"\s+", " ", c.lower().replace("ё", "е")).startswith(stem)), None)
        if c is not None:
            spec_cols[key] = c
    if spec_cols:
        log.info("Столбцы правил в таблице: %s", ", ".join(spec_cols.values()))
    # флаги включения по отдельным ООЗТ: столбцы «Вкл_<ООЗТ>» (Вкл/Выкл; пусто — без ограничения)
    oozt_flag_cols = {c: str(c).strip()[4:].strip() for c in df.columns
                      if str(c).strip().lower().startswith("вкл_")}

    def _flag(v):
        t = str(v).strip().lower() if v is not None and not (isinstance(v, float) and np.isnan(v)) else ""
        if not t:
            return None
        if t in ("0", "нет", "выкл", "выключено", "no", "false"):
            return False
        if t in ("1", "да", "вкл", "включено", "yes", "true"):
            return True
        return None

    out = []
    for i, r in df.iterrows():
        if pd.isna(r[c_name]):
            continue
        on = str(r[c_on]).strip().lower() in ("1", "1.0", "true", "да", "yes", "вкл")
        flags = {}
        for c, oo in oozt_flag_cols.items():
            v = _flag(r.get(c))
            if v is not None:
                flags[oo] = v
        out.append({
            "name": re.sub(r"\s+", " ", str(r[c_name])).strip(),
            "code": str(r[c_code]).strip() if c_code and not pd.isna(r[c_code]) else "",
            "on": on,
            "oozt_on": flags,
            "method": None if pd.isna(r[c_method]) else str(r[c_method]).strip(),
            "group": re.sub(r"\s+", " ", str(r[c_group])).strip() if c_group and not pd.isna(r[c_group]) else "",
            "eco": str(r[c_eco]).strip() if c_eco and not pd.isna(r[c_eco]) else "",
            "prikaz": str(r[c_prikaz]).strip() if c_prikaz and not pd.isna(r[c_prikaz]) else "",
            "season": str(r[c_season]).strip() if c_season and not pd.isna(r[c_season]) else "",
            "spec": {k: (None if pd.isna(r[c]) else r[c]) for k, c in spec_cols.items()},
            "row": int(i) + 2,                 # номер строки в Excel (для записи описаний)
        })
    log.info("Таблица мероприятий: %d строк, включено %d", len(out), sum(m["on"] for m in out))
    return out


def _cell_value(key: str, v):
    """Параметр правила → значение ячейки Excel."""
    if key == "types":
        return v[1] if v else "любые"
    if key in ("category", "groups"):
        return ", ".join(v) if v else "любые"
    if key == "sources":
        return ", ".join({"points": "точки", "polygons": "полигоны", "sdm": "SDM"}[x] for x in v) if v else "любые"
    if key == "taxon":
        return ", ".join(v) if v else None
    return None if v is None else float(v)


def update_measures_xlsx(path: str, cfg: Config | None = None, fill_from_method: bool = True,
                         backup: bool = True) -> str:
    """Дополняет таблицу мероприятий: столбцы правил (заполняются из текста метода там, где правило не задано),
    выпадающие списки, лист «Справка по правилам», столбец с полным описанием расчёта. Возвращает сводку."""
    import shutil
    import openpyxl
    from copy import copy as _cp
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    if cfg is not None:
        _MIN_SHARE[0] = float(getattr(cfg, "min_share_pct", 0) or 0)
    path = str(path)
    measures = read_measures(path)
    if backup:
        bak = Path(path).with_name(Path(path).stem + "_до_правил" + Path(path).suffix)
        if not bak.exists():
            shutil.copy2(path, bak)
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    norm = lambda v: re.sub(r"\s+", " ", str(v or "").strip().lower().replace("ё", "е"))
    hdr = {c: norm(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)}
    last = max([c for c, h in hdr.items() if h] or [1])
    proto = ws.cell(1, last)

    def ensure(title, stem, width):
        nonlocal last
        c = next((c for c, h in hdr.items() if h.startswith(stem)), None)
        if c is None:
            last += 1
            c = last
            cell = ws.cell(1, c, title)
            cell.font, cell.border = _cp(proto.font), _cp(proto.border)
            cell.fill = PatternFill("solid", fgColor="E8F0E0")
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            hdr[c] = norm(title)
            ws.column_dimensions[get_column_letter(c)].width = width
        return c

    widths = dict(rule=24, types=16, category=14, groups=18, taxon=14, sources=14)
    cols = {k: ensure(t, stem, widths.get(k, 11)) for k, t, stem, _ in SPEC_COLUMNS}
    c_desc = ensure(DESC_COLUMN, "процедура назначения", 95)
    n_fill = n_desc = 0
    wrap = Alignment(wrap_text=True, vertical="top")
    for m in measures:
        r = m["row"]
        spec = measure_spec(m)
        if fill_from_method and m["on"] and spec["rule"] and spec["origin"] == "текст метода":
            ws.cell(r, cols["rule"], RULES[spec["rule"]]["title"])
            for k, v in spec["p"].items():
                val = _cell_value(k, v)
                if val is not None and k != "min_share":
                    ws.cell(r, cols[k], val)
            m = dict(m, spec={k: ws.cell(r, c).value for k, c in cols.items()})
            spec = measure_spec(m)
            n_fill += 1
        for k, c in cols.items():
            ws.cell(r, c).alignment = wrap
        if spec["rule"] or m["on"]:
            txt = describe_spec(spec, full=True)
            ws.cell(r, c_desc, txt).alignment = wrap
            n_desc += 1
            need = (len(txt) // int(95 * 1.15) + 2) * 15
            if (ws.row_dimensions[r].height or 15) < need:
                ws.row_dimensions[r].height = need
    # выпадающие списки (подсказки: можно ввести и своё значение / несколько через запятую)
    ours = set(cols.values())
    ws.data_validations.dataValidation = [
        dv for dv in ws.data_validations.dataValidation
        if not any(rng.min_col in ours for rng in dv.sqref.ranges)]
    nrow = max(ws.max_row, 2)
    lists = {"rule": RULE_TITLES, "types": ["лесные", "луговые", "водные", "околоводные", "болотные",
                                             "водные, околоводные, болотные", "любые"],
             "category": SPECIES_CATEGORIES + ["любые"], "groups": SPECIES_GROUPS[:-1] + ["любые"],
             "sources": ["точки", "полигоны", "SDM", "точки, полигоны", "любые"]}
    # длинные списки (> 255 знаков Excel не принимает в формуле) — на скрытом листе «_списки»
    if "_списки" in wb.sheetnames:
        del wb["_списки"]
    ls = None
    for li, (k, vals) in enumerate(lists.items(), 1):
        f = '"' + ",".join(v.replace(", ", "; ") for v in vals) + '"'
        if len(f) > 250:
            if ls is None:
                ls = wb.create_sheet("_списки")
                ls.sheet_state = "hidden"
            cl = get_column_letter(li)
            for vi, v in enumerate(vals, 1):
                ls[f"{cl}{vi}"] = v.replace(", ", "; ")
            f = f"'_списки'!${cl}$1:${cl}${len(vals)}"
        dv = DataValidation(type="list", formula1=f, allow_blank=True, showErrorMessage=(k == "rule"),
                            errorTitle="Правило назначения",
                            error="Выберите правило из списка (см. лист «Справка по правилам»).")
        dv.add(f"{get_column_letter(cols[k])}2:{get_column_letter(cols[k])}{nrow}")
        ws.add_data_validation(dv)
    # лист-справка
    if "Справка по правилам" in wb.sheetnames:
        del wb["Справка по правилам"]
    h = wb.create_sheet("Справка по правилам")
    bold = Font(bold=True)
    rows = [("Как скрипт читает способ назначения мероприятия", None, None, None), (None,) * 4,
            ("Строка таблицы рассчитывается, если «Вкл/выкл» = 1. Правило берётся из столбца «Правило назначения»; "
             "параметры — из столбцов справа. Пустой параметр — значение по умолчанию (указано ниже). Если "
             "«Правило назначения» пусто, правило и параметры определяются по ключевым словам «Метода выявления» "
             "(как в прежних версиях). Столбец «" + DESC_COLUMN + "» заполняется скриптом и только описывает расчёт.",
             None, None, None), (None,) * 4,
            ("Правило назначения", "Что делает", "Нужные данные", "Параметры (по умолчанию)")]
    data_t = {"trails": "тропы", "hydro": "гидрография", "dem": "ЦМР", "species": "слои видов",
              "drainage": "дренажные канавы", "wetlands": "болота (вектор)",
              "lesopat": "точки лесопатологии (ПВН)", "roads": "дороги с твёрдым покрытием",
              "water_ref": "справочник водоёмов и болот", "flowering": "справочник цветения"}
    for k, v in RULES.items():
        ps = []
        for pk, d in v["params"].items():
            dv_ = ("обязательно" if d == REQ else "общий порог из окна скрипта" if d == SHARED else
                   "любые" if d is None else ("да" if d else "нет") if isinstance(d, bool) else
                   f"{_fmt(d[0])}–{_fmt(d[1])}" if pk == "ring" else
                   ", ".join(str(x) for x in d) if isinstance(d, (tuple, list)) else _fmt(d))
            ps.append(f"{SPEC_TITLE[pk]}: {dv_}")
        rows.append((v["title"], v["help"],
                     ", ".join(data_t[d] for d in v["data"]) + ((" (+ " if v["data"] else "")
                     + ", ".join(data_t[o] for o in v["opt"]) + (", если есть)" if v["data"] else ", если есть")
                     if v["opt"] else "")
                     or "только слой территории", "; ".join(ps)))
    rows += [(None,) * 4, ("Столбец", "Что указывать", None, None)]
    rows += [(t, help_, None, None) for _, t, _, help_ in SPEC_COLUMNS]
    rows += [(None,) * 4, ("Слои видов (задаются в окне скрипта)", "Значения", None, None),
             ("Категория", "; ".join(f"{k} — {v}" for k, v in SPECIES_CATEGORY_HELP.items()), None, None),
             ("Группа организмов", ", ".join(SPECIES_GROUPS[:-1]) + "; можно ввести новую группу", None, None),
             ("Тип данных", "точки, полигоны — векторные слои (поле с латинскими названиями находится само); "
                            "SDM — растр вероятности 0–1, название вида = имя файла", None, None)]
    for row in rows:
        h.append(list(row))
    for rr in h.iter_rows():
        for cell in rr:
            cell.alignment = wrap
    for rr in (1, 5):
        for cell in h[rr]:
            cell.font = bold
    for rr in range(1, h.max_row + 1):
        if h.cell(rr, 1).value in ("Столбец", "Слои видов (задаются в окне скрипта)"):
            for cell in h[rr]:
                cell.font = bold
    h["A1"].font = Font(bold=True, size=13)
    h["A1"].alignment = Alignment(wrap_text=False)
    h.page_setup.orientation = "landscape"
    h.sheet_properties.pageSetUpPr.fitToPage = True
    h.page_setup.fitToWidth, h.page_setup.fitToHeight = 1, 0
    h.freeze_panes = "A6"
    for col, w in zip("ABCD", (30, 70, 26, 60)):
        h.column_dimensions[col].width = w
    h.merge_cells("A3:D3")
    h.row_dimensions[3].height = 60
    wb.save(path)
    return (f"Таблица обновлена: {Path(path).name}. Правила записаны в {n_fill} строк(и), описания — в {n_desc}; "
            f"добавлен лист «Справка по правилам».")


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
    if layer in LAYER_STYLE:
        return LAYER_STYLE[layer]
    return _pattern_style(layer)


def _pattern_style(layer: str):
    """Стиль слоёв-обоснований и географически очерченных участков работ."""
    t = str(layer).lower()
    inv = "инваз" in t
    if "колокольчик" in t:
        return ("#6a3d9a", "poly")
    if "лесные участки самовосстановления" in t:
        return ("#33a02c", "outline")
    if "участки лесопатологических мероприятий" in t:
        return ("#8c564b", "outline")
    if "участки проведения мероприятия" in t:
        return ("#d95f02", "poly")
    if t.startswith("участки работ ("):
        if "луг" in t:
            return ("#d9a404", "poly")
        if "вод" in t or "болот" in t:
            return ("#1f78b4", "poly")
        return ("#2b8cbe", "poly")
    if "водные объекты — выбранные" in t:
        return ("#08519c", "poly")
    if "дренажные канавы" in t:
        return ("#008b8b", "line")
    if "болота" in t:
        return ("#2ca25f", "poly")
    if t.startswith("тропы на расстоянии"):
        return ("#1f77b4", "poly") if "водн" in t else ("#9467bd", "line")
    if "— sdm" in t:
        return ("#f7b6d2" if inv else "#98df8a", "poly")
    if t.startswith("плотность"):
        return ("#7f7f7f", "poly")
    if t.startswith("зоны повышенной плотности троп"):
        return ("#ff7f0e", "poly")
    if t.startswith("полоса"):
        return ("#7fbf7b", "poly") if ("50–100" in t or "50-100" in t) else ("#d8b365", "poly")
    if t.startswith("лесопатология"):
        for k, v in (("санитарная рубка", "#d62728"), ("рубки ухода", "#8c564b"),
                     ("надзор", "#1f77b4"), ("посадка", "#2ca02c"),
                     ("эталон", "#9467bd"), ("биотехния", "#e377c2")):
            if k in t:
                return (v, "point")
        return ("#8c564b", "point")
    if "— точки" in t:
        return ("#e377c2" if inv else "#2ca02c" if ("охран" in t or "редк" in t) else "#000000", "point")
    if "— полигоны" in t or "виды" in t or "видов" in t:
        return ("#e377c2" if inv else "#2ca02c", "poly")
    return ("#17becf", "poly")


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
    ax.set_xlabel("Долгота, в.д. (WGS 84, UTM 37N)", fontsize=10)
    ax.set_ylabel("Широта, с.ш.", fontsize=10)


def _legend_label(layer: str) -> str:
    """Подпись слоя в легенде схемы — компактная с переносом строк (чтобы не сжимать поле карты)."""
    import textwrap
    s = str(layer)
    if len(s) > 42 and " — " in s and not s.startswith(("лесопатология", "полоса", "уклон")):
        m = re.match(r"^(.*?)\s—\s(точки|полигоны|SDM.*)$", s)
        if m and len(m.group(1)) > 36:
            taxa = [t.strip() for t in m.group(1).split(",") if t.strip()]
            if len(taxa) > 2:
                s = f"{taxa[0]} и ещё {len(taxa) - 1} — {m.group(2)}"
    if len(s) > 38 and "\n" not in s:
        s = textwrap.fill(s, width=38, break_long_words=False, break_on_hyphens=False)
    return s


def _place_labels(ax, items, sizes=(11, 10, 9, 8, 7), leaders=True, max_labels=150, reserved=None):
    """Подписи полигонов и точек без наложений.
    items: [(текст, геометрия)] в порядке приоритета. Подпись ставится внутрь полигона, если помещается
    (перебор кеглей sizes); иначе — рядом с объектом с выноской; если и так не помещается — пропускается.
    reserved: список уже занятых областей (геометрии shapely или кортежи (cx, cy, text, fs)).
    Возвращает число размещённых подписей."""
    import matplotlib.patheffects as pe
    from shapely.geometry import box as _box
    fig = ax.figure
    ax.apply_aspect()
    pos = ax.get_position()
    x0, x1 = ax.get_xlim()
    y0, y1 = ax.get_ylim()
    m_per_pt = (x1 - x0) / (pos.width * fig.get_figwidth() * 72)
    frame = _box(x0, y0, x1, y1)
    placed = []
    halo = [pe.withStroke(linewidth=2.6, foreground="white")]

    def lbl_box(cx, cy, t, fs):
        hw = (len(t) * fs * 0.62 + 3) * m_per_pt / 2
        hh = (fs * 1.05 + 2) * m_per_pt / 2
        return _box(cx - hw, cy - hh, cx + hw, cy + hh)

    if reserved:
        for r in reserved:
            if isinstance(r, tuple) and len(r) == 4:
                placed.append(lbl_box(r[0], r[1], str(r[2]), float(r[3])))
            elif r is not None and not getattr(r, "is_empty", True):
                placed.append(r)

    def free(b):
        return frame.contains(b) and not any(b.intersects(q) for q in placed)
    n = 0
    for t, geom in items:
        if n >= max_labels:
            break
        if geom is None or geom.is_empty:
            continue
        p = geom.representative_point()
        done = False
        for fs in sizes:                                   # 1) внутри полигона
            b = lbl_box(p.x, p.y, t, fs)
            if geom.buffer(fs * 0.25 * m_per_pt).contains(b) and free(b):
                ax.annotate(t, (p.x, p.y), ha="center", va="center", fontsize=fs, fontweight="bold",
                            color="#111111", zorder=13, path_effects=halo)
                placed.append(b); done = True
                break
        if done or not leaders:
            n += done
            continue
        fs = sizes[-1] + 1                                 # 2) рядом с выноской
        gx0, gy0, gx1, gy1 = geom.bounds
        b0 = lbl_box(0, 0, t, fs)
        bw, bh = b0.bounds[2] - b0.bounds[0], b0.bounds[3] - b0.bounds[1]
        off = 6 * m_per_pt
        mx, my = (gx0 + gx1) / 2, (gy0 + gy1) / 2
        cands = []
        for mult in (1.0, 2.4, 4.0, 6.0):
            dx, dy = mult * off + bw / 2, mult * off + bh / 2
            cands.extend([
                (gx1 + dx, my), (gx0 - dx, my),
                (mx, gy1 + dy), (mx, gy0 - dy),
                (gx1 + dx, gy1 + dy), (gx0 - dx, gy1 + dy),
                (gx1 + dx, gy0 - dy), (gx0 - dx, gy0 - dy),
            ])
        for cx, cy in cands:
            b = lbl_box(cx, cy, t, fs)
            if free(b):
                ax.annotate(t, xy=(p.x, p.y), xytext=(cx, cy), textcoords="data", ha="center", va="center",
                            fontsize=fs, fontweight="bold", color="#111111", zorder=13, path_effects=halo,
                            arrowprops=dict(arrowstyle="-", color="#333333", lw=0.6, shrinkA=0, shrinkB=0))
                placed.append(b); n += 1
                break
    return n


def _utm_figsize(w: float, h: float) -> tuple[float, float]:
    """Размер фигуры под метрический экстент UTM 37N без вертикального или горизонтального сжатия."""
    ratio = max(h, 1.0) / max(w, 1.0)
    if ratio <= 1.0:
        fw = 10.5
        fh = max(5.6, fw * ratio)
    else:
        fh = min(12.0, max(7.0, 9.5 * ratio))
        fw = max(6.2, fh / ratio)
    return fw, fh


def _plot_evidence_group(ax, handles: list, layer: str, grp: gpd.GeoDataFrame, overview: bool = False):
    """Единообразная отрисовка одного слоя участков работ / оснований и добавление элемента в легенду."""
    color, kind = _slope_style(layer)
    lbl = _legend_label(layer)
    if kind == "point":
        is_pvn = layer.startswith("лесопатология")
        is_her = "heracleum" in layer.lower()
        species_point = ("виды — точки" in layer or "Corvus — точки" in layer or is_pvn or is_her)
        if overview:
            pt_sz = 36 if is_pvn else (28 if species_point else 22)
            lg_sz = 8.5 if species_point else 7.5
        else:
            pt_sz = 72 if is_pvn else (60 if species_point else 42)
            lg_sz = 11.5 if species_point else 9.0
        if is_her:                       # борщевик Сосновского — значок ×1,5
            pt_sz *= 1.5
            lg_sz *= 1.5 ** 0.5
        grp.plot(ax=ax, color=color, markersize=pt_sz, marker="D" if is_pvn else "o",
                 edgecolor="white", linewidth=0.7, zorder=7 if is_pvn else 6.5)
        handles.append(Line2D([], [], marker="D" if is_pvn else "o", color="w",
                              markerfacecolor=color, markersize=lg_sz, label=lbl))
    elif kind == "line":
        lw = 1.5 if overview else 2.2
        grp.plot(ax=ax, color=color, linewidth=lw, zorder=5.2)
        handles.append(Line2D([], [], color=color, linewidth=lw + 0.4, label=lbl))
    elif kind == "slope":
        grp.plot(ax=ax, facecolor=color, edgecolor=color, alpha=0.95, linewidth=0.7, zorder=4.2)
        handles.append(Patch(facecolor=color, edgecolor=color, label=lbl))
    elif kind == "outline":
        grp.plot(ax=ax, facecolor=color, edgecolor=color, alpha=0.18, linewidth=1.0, linestyle="-.", zorder=4.0)
        handles.append(Patch(facecolor=color, alpha=0.25, edgecolor=color, linestyle="-.", label=lbl))
    elif "SDM" in layer:
        grp.plot(ax=ax, facecolor="none", edgecolor=color, alpha=0.9,
                 linewidth=0.75 if overview else 0.9, hatch="////", zorder=4.1)
        handles.append(Patch(facecolor="none", edgecolor=color, hatch="////", label=lbl))
    else:
        al = 0.42 if overview else 0.48
        grp.plot(ax=ax, facecolor=color, edgecolor=color, alpha=al,
                 linewidth=0.65 if overview else 0.85, zorder=4.3)
        handles.append(Patch(facecolor=color, alpha=0.55, edgecolor=color, label=lbl))


def _oozt_map_ctx(data: Data, oozt_name: str) -> dict | None:
    """Предрасчёт контура ООЗТ, экстента UTM 37N, кварталов и подрезанных базовых слоёв
    для быстрой отрисовки общей обзорной схемы и всех частных схем по мероприятиям."""
    same = data.oozt[data.oozt["OOZT"].astype(str) == str(oozt_name)]
    kn = getattr(data, "kvnet", None)
    kn_same = None
    if kn is not None and not kn.empty:
        if "KVOOZT" in kn.columns and (kn["KVOOZT"].astype(str).str.strip() != "").any():
            kn_same = kn[kn["KVOOZT"].astype(str) == str(oozt_name)]
        elif data.oozt["OOZT"].astype(str).nunique() <= 1:
            kn_same = kn
    gparts = []
    if not same.empty:
        gparts.extend(same.geometry.values)
    if kn_same is not None and not kn_same.empty:
        gparts.extend(kn_same.geometry.values)
    if not gparts:
        return None
    geom = _safe_union(gparts)
    minx, miny, maxx, maxy = geom.bounds
    pad = max(maxx - minx, maxy - miny) * 0.06 + 40
    ext = (minx - pad, maxx + pad, miny - pad * 1.35, maxy + pad)
    w, h = ext[1] - ext[0], ext[3] - ext[2]
    fig_w, fig_h = _utm_figsize(w, h)
    _pp = pond_polys(data)
    hp = gpd.clip(_pp, geom) if (_pp is not None and not _pp.empty) else None
    _wl = wetland_polys(data)
    wl = gpd.clip(_wl, geom) if (_wl is not None and not _wl.empty) else None
    hl = gpd.clip(data.hydro_lines, geom) if (data.hydro_lines is not None and not data.hydro_lines.empty) else None
    paved = _paved_roads(data)
    dirt = None
    if data.roads_dirt is not None and not data.roads_dirt.empty:
        dirt = data.roads_dirt
    elif (paved is not None and (data.roads_paved is None or data.roads_paved.empty)
          and data.trails is not None and not data.trails.empty):
        dirt = data.trails[~data.trails.index.isin(paved.index)]
    elif data.trails is not None and not data.trails.empty:
        dirt = data.trails
    tl = gpd.clip(dirt, geom) if (dirt is not None and not dirt.empty) else None
    it = gpd.clip(data.trails_informal, geom) if (data.trails_informal is not None and not data.trails_informal.empty) else None
    pv = gpd.clip(paved, geom) if (paved is not None and not paved.empty) else None
    q_rows = []
    if kn_same is not None and not kn_same.empty and data.cfg.unit_mode != "eco":
        for kv_id, kg in kn_same.groupby("KV"):
            q_rows.append((str(kv_id).removesuffix(".0"), _safe_union(kg.geometry.values)))
    else:
        q = data.quarters()
        if q is not None and not q.empty:
            qq = q[q["OOZT"].astype(str) == str(oozt_name)] if "OOZT" in q.columns else q
            seen_kv = set()
            for _, r in qq.iterrows():
                kv_s = str(r["KV"]).removesuffix(".0")
                seen_kv.add(kv_s)
                q_rows.append((kv_s, r.geometry))
            if kn_same is not None and not kn_same.empty:
                for kv_id, kg in kn_same.groupby("KV"):
                    kv_s = str(kv_id).removesuffix(".0")
                    if kv_s not in seen_kv:
                        q_rows.append((kv_s, _safe_union(kg.geometry.values)))
    return dict(geom=geom, ext=ext, w=w, h=h, fig_w=fig_w, fig_h=fig_h,
                hp=hp, wl=wl, hl=hl, tl=tl, it=it, pv=pv, q_rows=q_rows, basemap_cache=None)


def _season_suffix(season) -> str:
    """Сезонность в заголовке: « (V–VIII)»; если в тексте уже есть скобки — « — IV–VII (…)»."""
    s = str(season or "").strip()
    if not s:
        return ""
    return f" — {s}" if "(" in s else f" ({s})"


def _measure_fig_name(oozt_name: str, idx: int, code: str, measure_name: str) -> str:
    """Короткое безопасное имя PNG-файла для частной схемы мероприятия (под Windows MAX_PATH)."""
    oo_s = re.sub(r"\W+", "_", str(oozt_name)).strip("_")[:26] or "oozt"
    cd_s = re.sub(r"\W+", "_", str(code)).strip("_")[:8]
    nm_s = re.sub(r"\W+", "_", str(measure_name)).strip("_")[:34] or "measure"
    mid = f"{cd_s}_{nm_s}" if cd_s else nm_s
    return f"fig_{oo_s}_m{idx:02d}_{mid}.png"


def draw_oozt(data: Data, evidence: gpd.GeoDataFrame | None, oozt_name: str, out_png: Path,
               basemap=True, ctx: dict | None = None, measure_name: str | None = None,
               measure_code: str = "", active_kvs: set[str] | None = None,
               work_geoms: list | None = None, measure_season: str = ""):
    """Схема на одно ООЗТ в проекции UTM Zone 37N (EPSG:32637, масштаб 1:1 без сжатия):
    при measure_name=None — общая обзорная схема по всем мероприятиям;
    при заданном measure_name — частная схема конкретного мероприятия с выделением кварталов и участков работ."""
    import textwrap
    import matplotlib.patheffects as pe
    if ctx is None:
        ctx = _oozt_map_ctx(data, oozt_name)
    if ctx is None:
        return None
    geom, ext, w, h = ctx["geom"], ctx["ext"], ctx["w"], ctx["h"]
    fig_w, fig_h = ctx["fig_w"], ctx["fig_h"]
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=200)
    ax.set_aspect("equal", adjustable="box")
    ax.set_xlim(ext[0], ext[1])
    ax.set_ylim(ext[2], ext[3])
    handles = []
    if ctx["hp"] is not None and not ctx["hp"].empty:
        ctx["hp"].plot(ax=ax, color="#A6CEE3", edgecolor="none", alpha=0.85, zorder=2)
        handles.append(Patch(facecolor="#A6CEE3", alpha=0.9, edgecolor="none", label="озёра и пруды"))
    if ctx.get("wl") is not None and not ctx["wl"].empty:
        ctx["wl"].plot(ax=ax, color=WETLAND_COLOR, edgecolor="none", alpha=0.9, zorder=2)
        handles.append(Patch(facecolor=WETLAND_COLOR, alpha=0.9, edgecolor="none", label="болота"))
    if ctx["hl"] is not None and not ctx["hl"].empty:
        ctx["hl"].plot(ax=ax, color="#1F78B4", linewidth=0.9, alpha=0.9, zorder=2.5)
        handles.append(Line2D([], [], color="#1F78B4", linewidth=1.6, label="реки и канавы"))
    if ctx["tl"] is not None and not ctx["tl"].empty:
        ctx["tl"].plot(ax=ax, color="#b5b5b5", linewidth=0.4, alpha=0.7, zorder=3.4)
        handles.append(Line2D([], [], color="#b5b5b5", linewidth=1.0, label="грунтовые дороги и тропы"))
    if ctx["it"] is not None and not ctx["it"].empty:
        ctx["it"].plot(ax=ax, color="#c98d5a", linewidth=0.65, alpha=0.85, zorder=3.5)
        handles.append(Line2D([], [], color="#c98d5a", linewidth=1.2, label=_inf_label()))
    if ctx["pv"] is not None and not ctx["pv"].empty:
        ctx["pv"].plot(ax=ax, color="#5f6368", linewidth=1.35, zorder=3.8)
        handles.append(Line2D([], [], color="#5f6368", linewidth=2.0, label="дороги с твёрдым покрытием"))
    ev_src = evidence
    if measure_name is not None and ev_src is not None and not ev_src.empty:
        ev_src = ev_src[ev_src["measure"].astype(str) == str(measure_name)]
    pvn_lab_items = []
    reserved_boxes: list = []
    has_poly_ev = False
    if ev_src is not None and not ev_src.empty:
        is_pt = ev_src.geom_type.isin(["Point", "MultiPoint"])
        has_oo = ("oozt" in ev_src.columns
                  and ev_src["oozt"].fillna("").astype(str).str.strip().ne(""))
        ev_pts = ev_src[is_pt & (
            (has_oo & (ev_src["oozt"].astype(str) == str(oozt_name)))
            | (~has_oo & ev_src.intersects(geom.buffer(5)))
        )].copy()
        if not ev_pts.empty:
            pt_lbl = ev_pts["label"].fillna("").astype(str) if "label" in ev_pts.columns else ""
            pt_c = ev_pts.geometry.centroid
            ev_pts["_dedup_k"] = (
                ev_pts["layer"].astype(str) + "|" + pt_lbl + "|"
                + pt_c.x.round(1).astype(str) + "|" + pt_c.y.round(1).astype(str)
            )
            # борщевик Сосновского — все точки, без прореживания совпадающих
            her = ev_pts["layer"].astype(str).str.lower().str.contains("heracleum")
            ev_pts = pd.concat([ev_pts[~her].drop_duplicates("_dedup_k"), ev_pts[her]]).drop(columns=["_dedup_k"])
        ev_nonpt = ev_src[~is_pt & ev_src.intersects(geom.buffer(1))].copy()
        if not ev_nonpt.empty:
            ev_nonpt = gpd.clip(ev_nonpt, geom.buffer(0.5))
            has_poly_ev = bool(ev_nonpt.geom_type.isin(["Polygon", "MultiPolygon"]).any())
        if measure_name is not None and not has_poly_ev and work_geoms:
            wg_gdf = gpd.clip(gpd.GeoDataFrame(geometry=work_geoms, crs=WORK_EPSG), geom)
            wg_gdf = wg_gdf[~wg_gdf.geometry.is_empty]
            if not wg_gdf.empty:
                _plot_evidence_group(ax, handles, "участки проведения мероприятия", wg_gdf, overview=False)
                has_poly_ev = True
        ev_parts = [x for x in (ev_nonpt, ev_pts) if x is not None and not x.empty]
        if ev_parts:
            ev_plot = gpd.GeoDataFrame(pd.concat(ev_parts, ignore_index=True), crs=WORK_EPSG)
            use_ov = True if measure_name is None else (len(ev_pts) > 35)
            for layer, grp in ev_plot.groupby("layer"):
                _plot_evidence_group(ax, handles, layer, grp, overview=use_ov)
            if measure_name is not None and not ev_pts.empty and "label" in ev_pts.columns:
                seen_pvn = set()
                for _, pr in ev_pts.iterrows():
                    lb = str(pr.get("label", "") or "").strip()
                    if lb and lb not in ("nan", "None") and str(pr.get("layer", "")).startswith("лесопатология"):
                        txt_p = lb if lb.upper().startswith("ПВН") else f"ПВН {lb}"
                        pc = pr.geometry.centroid
                        kp = (txt_p, round(pc.x, 1), round(pc.y, 1))
                        if kp not in seen_pvn:
                            seen_pvn.add(kp)
                            pvn_lab_items.append((txt_p, pc))
                            reserved_boxes.append((pc.x, pc.y, "ММ", 10.0))
    elif measure_name is not None and work_geoms:
        wg_gdf = gpd.clip(gpd.GeoDataFrame(geometry=work_geoms, crs=WORK_EPSG), geom)
        wg_gdf = wg_gdf[~wg_gdf.geometry.is_empty]
        if not wg_gdf.empty:
            _plot_evidence_group(ax, handles, "участки проведения мероприятия", wg_gdf, overview=False)
    q_rows = ctx["q_rows"]
    q_halo = [pe.withStroke(linewidth=2.8, foreground="white")]
    if measure_name is None or active_kvs is None:
        for kv_s, qg in q_rows:
            if qg is None or qg.is_empty:
                continue
            gpd.GeoSeries([qg], crs=WORK_EPSG).boundary.plot(ax=ax, color="#333333", linewidth=1.1, zorder=9)
            c = qg.representative_point()
            ax.annotate(kv_s, (c.x, c.y), ha="center", va="center", fontsize=9.5, fontweight="bold",
                        color="#1a1a1a", zorder=12, path_effects=q_halo)
            reserved_boxes.append((c.x, c.y, kv_s, 11.0))
        handles.append(Line2D([], [], color="#333333", linewidth=1.2, label="границы кварталов"))
    else:
        act_set = {str(x).removesuffix(".0") for x in active_kvs}
        has_act = has_inact = False
        for kv_s, qg in q_rows:
            if qg is None or qg.is_empty:
                continue
            if kv_s in act_set:
                continue
            has_inact = True
            gpd.GeoSeries([qg], crs=WORK_EPSG).boundary.plot(
                ax=ax, color="#8a8a8a", linewidth=0.75, linestyle="--", zorder=8.6)
            c = qg.representative_point()
            ax.annotate(kv_s, (c.x, c.y), ha="center", va="center", fontsize=8.5,
                        color="#666666", zorder=11.5, path_effects=q_halo)
            reserved_boxes.append((c.x, c.y, kv_s, 10.0))
        for kv_s, qg in q_rows:
            if qg is None or qg.is_empty or kv_s not in act_set:
                continue
            has_act = True
            gpd.GeoSeries([qg], crs=WORK_EPSG).boundary.plot(
                ax=ax, color="#1f1f1f", linewidth=1.35, zorder=9.3)
            c = qg.representative_point()
            ax.annotate(kv_s, (c.x, c.y), ha="center", va="center", fontsize=9.5, fontweight="bold",
                        color="#111111", zorder=12.2, path_effects=q_halo)
            reserved_boxes.append((c.x, c.y, kv_s, 11.0))
        if has_act:
            handles.append(Line2D([], [], color="#1f1f1f", linewidth=1.5, label="кварталы с мероприятием"))
        if has_inact:
            handles.append(Line2D([], [], color="#8a8a8a", linewidth=1.0, linestyle="--", label="прочие кварталы"))
    gpd.GeoSeries([geom], crs=WORK_EPSG).boundary.plot(ax=ax, color="black", linewidth=2.1, zorder=10)
    ax.set_xlim(ext[0], ext[1])
    ax.set_ylim(ext[2], ext[3])
    if basemap:
        if ctx.get("basemap_cache") is not None:
            arr_b, ext_b, orig_b, attr_b = ctx["basemap_cache"]
            ax.imshow(arr_b, extent=ext_b, origin=orig_b, interpolation="bilinear", zorder=0)
            if attr_b:
                ax.text(0.99, 0.01, attr_b, transform=ax.transAxes, ha="right", va="bottom",
                        fontsize=7.5, color="#555555", zorder=14,
                        bbox=dict(boxstyle="square,pad=0.15", facecolor="white", alpha=0.65, edgecolor="none"))
        else:
            try:
                import contextily as cx
                for _attempt in range(3):
                    prov = pick_basemap_provider(skip=_attempt)
                    if prov is None:
                        break
                    try:
                        attr_b = prov.get("attribution", "© OpenStreetMap")
                        cx.add_basemap(ax, crs=f"EPSG:{WORK_EPSG}", source=prov,
                                       attribution=attr_b, attribution_size=9, zorder=0)
                        desaturate_last_basemap(ax)
                        if ax.images:
                            im_b = ax.images[-1]
                            ctx["basemap_cache"] = (
                                np.asarray(im_b.get_array()).copy(),
                                tuple(im_b.get_extent()),
                                im_b.origin,
                                attr_b,
                            )
                        break
                    except Exception as e:
                        log.warning("подложка %s недоступна (%s) — пробую другой источник",
                                    prov.get("name"), e)
            except Exception as e:
                log.warning("подложка OSM недоступна: %s", e)
    ax.legend(handles=handles, loc="upper left", bbox_to_anchor=(1.01, 1.0), fontsize=9.5,
              frameon=True, title="Легенда", title_fontsize=10.5)
    if measure_name is None:
        ax.set_title(f"{oozt_name}. Обзорная схема мероприятий", fontsize=14, fontweight="bold")
    else:
        raw_title = f"{oozt_name}. {(measure_code + ' ') if measure_code else ''}{measure_name}" \
            + _season_suffix(measure_season)
        wrapped_t = textwrap.fill(raw_title, width=56, break_long_words=False, break_on_hyphens=False)
        t_fs = 12.0 if "\n" in wrapped_t else 13.5
        ax.set_title(wrapped_t, fontsize=t_fs, fontweight="bold", pad=10)
    ax.tick_params(labelsize=10.5)
    try:
        set_degree_ticks(ax, ext)
    except Exception as e:
        log.warning("не удалось построить сетку в градусах (%s) — оси в метрах", e)
        ax.set_xlabel("E, м (UTM 37N)", fontsize=11)
        ax.set_ylabel("N, м (UTM 37N)", fontsize=11)
        ax.ticklabel_format(style="plain", useOffset=False)
    # масштабная линейка
    target_l = w / 4
    dec = 10 ** np.floor(np.log10(max(target_l, 1.0)))
    L_bar = next(m * dec for m in (5, 2, 1) if m * dec <= target_l) if target_l >= dec else dec
    x0, y0 = ext[0] + w * 0.04, ext[2] + h * 0.032
    ax.plot([x0, x0 + L_bar], [y0, y0], color="black", linewidth=3.5, zorder=14,
            path_effects=[pe.withStroke(linewidth=6.0, foreground="white")])
    bar_txt = f"{L_bar / 1000:g} км" if L_bar >= 1000 else f"{int(L_bar)} м"
    ax.text(x0 + L_bar / 2, y0 + h * 0.018, bar_txt, ha="center", va="bottom", fontsize=10.5, zorder=14,
            path_effects=[pe.withStroke(linewidth=3, foreground="white")])
    reserved_boxes.append(box(x0 - w * 0.02, ext[2], x0 + L_bar + w * 0.02, y0 + h * 0.055))
    ax.set_aspect("equal", adjustable="box")
    fig.tight_layout()
    ax.set_xlim(ext[0], ext[1])
    ax.set_ylim(ext[2], ext[3])
    ax.set_aspect("equal", adjustable="box")
    if pvn_lab_items:
        try:
            _place_labels(ax, pvn_lab_items, sizes=(8.5,), leaders=True, max_labels=60,
                          reserved=reserved_boxes)
        except Exception:
            pass
    fig.savefig(out_png, bbox_inches="tight")
    plt.close(fig)
    return True


def draw_kvartal(data: Data, evidence: gpd.GeoDataFrame | None, kv, oozt_name, out_png: Path, basemap=True):
    """Схема квартала в проекции UTM Zone 37N (EPSG:32637, строгий масштаб 1:1 без сжатия по вертикали)."""
    vyd = data.oozt[(data.oozt["KV"].astype(str) == str(kv)) & (data.oozt["OOZT"].astype(str) == str(oozt_name))]
    kn = getattr(data, "kvnet", None)
    kn_oozt = None
    if kn is not None and not kn.empty:
        if "KVOOZT" in kn.columns and (kn["KVOOZT"].astype(str).str.strip() != "").any():
            kn_oozt = kn[kn["KVOOZT"].astype(str) == str(oozt_name)]
        else:
            kn_oozt = kn
    focus_geom = None
    if kn_oozt is not None and not kn_oozt.empty and data.cfg.unit_mode == "kv_cut":
        krow = kn_oozt[kn_oozt["KV"].astype(str) == str(kv)]
        if not krow.empty:
            focus_geom = _safe_union(krow.geometry.values)
    if focus_geom is None:
        if vyd.empty:
            if kn_oozt is not None and not kn_oozt.empty:
                krow = kn_oozt[kn_oozt["KV"].astype(str) == str(kv)]
                if not krow.empty:
                    focus_geom = _safe_union(krow.geometry.values)
            if focus_geom is None:
                return None
        else:
            focus_geom = _safe_union(vyd.geometry.values)
    minx, miny, maxx, maxy = focus_geom.bounds
    pad = max((maxx - minx), (maxy - miny)) * 0.12 + 30
    ext = (minx - pad, maxx + pad, miny - pad, maxy + pad)
    w, h = ext[1] - ext[0], ext[3] - ext[2]
    fig_w, fig_h = _utm_figsize(w, h)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=200)
    ax.set_aspect("equal", adjustable="box")
    plt.rcParams.update({"font.size": 12})

    # Лесные выделы имеют зелёную заливку, луговые прозрачны.
    # Границы всех выделов повторно рисуются поверх тематических слоёв ниже.
    cfg = data.cfg
    group_mode = cfg.unit_mode == "eco" and bool(cfg.report_group_field)   # индекс = тип, полигоны разбросаны
    map_box0 = box(ext[0], ext[2], ext[1], ext[3])
    same_oozt0 = data.oozt[data.oozt["OOZT"] == oozt_name]
    others = same_oozt0[(same_oozt0["KV"] != kv) & same_oozt0.intersects(map_box0)]
    if cfg.unit_mode == "eco" and not others.empty:      # прочие полигоны — тонкий контур без подписей
        others.boundary.plot(ax=ax, color="#9AA0A6", linewidth=0.45, zorder=3)
    # ---- подложка (снизу вверх): OSM → тепловая карта рекреации → реки/канавы/озёра →
    # горизонтали → грунтовые дороги → дороги с твёрдым покрытием → мероприятия → кварталы.
    # Слои показываются только в пределах границы ООЗТ; вне её видна только подложка OSM.
    handles = []
    u = focus_geom if vyd.empty else data.union
    if data.rec_points is not None and not data.rec_points.empty:
        try:
            pts = data.rec_points[data.rec_points.within(u)]
            if not pts.empty:
                dens, tr = point_density(pts, data.bounds_pad, radius_m=150.0)
                inside = data.inside_mask(tr, dens.shape)
                vals = dens[inside & (dens > 0)]
                if len(vals) > 10:
                    thr = float(np.quantile(vals, 0.95))
                    field = np.ma.masked_less(np.where(inside, dens, np.nan), thr)
                    nmin, nmax = tr.c, tr.f
                    ax.imshow(field, extent=(nmin, nmin + tr.a * dens.shape[1],
                                             nmax - tr.e * dens.shape[0], nmax),
                              origin="upper", cmap="YlOrRd", alpha=0.4, zorder=1, interpolation="bilinear")
                    handles.append(Patch(facecolor="#f4a582", alpha=0.55, edgecolor="none",
                                         label="концентрация рекреац. объектов (≥ 0,95 квантиля)"))
                cat_col = next((c for c in pts.columns if c != "geometry" and str(c).lower()
                                in ("category", "cat", "fclass", "type", "категория")), None)
                if cat_col:
                    cmap_r = plt.get_cmap("tab10")
                    for i, cv in enumerate(sorted(pts[cat_col].astype(str).unique())[:10]):
                        sub = pts[pts[cat_col].astype(str) == cv]
                        sub.plot(ax=ax, color=cmap_r(i), markersize=14, edgecolor="white",
                                 linewidth=0.4, zorder=6)
                        handles.append(Line2D([], [], marker="o", color="w", markerfacecolor=cmap_r(i),
                                              markersize=7, label=f"рекреация: {cv}"))
                else:
                    pts.plot(ax=ax, color="#d95f02", markersize=14, edgecolor="white", linewidth=0.4, zorder=6)
                    handles.append(Line2D([], [], marker="o", color="w", markerfacecolor="#d95f02",
                                          markersize=7, label="рекреационные объекты"))
        except Exception as e:
            log.warning("  тепловая карта рекреации не построена: %s", e)
    _pp = pond_polys(data)
    if _pp is not None and not _pp.empty:
        hp = gpd.clip(_pp, u)
        if not hp.empty:
            hp.plot(ax=ax, color="#A6CEE3", edgecolor="none", alpha=0.85, zorder=2)
            handles.append(Patch(facecolor="#A6CEE3", alpha=0.9, edgecolor="none", label="озёра и пруды"))
    _wl = wetland_polys(data)
    if _wl is not None and not _wl.empty:
        wl = gpd.clip(_wl, u)
        if not wl.empty:
            wl.plot(ax=ax, color=WETLAND_COLOR, edgecolor="none", alpha=0.9, zorder=2)
            handles.append(Patch(facecolor=WETLAND_COLOR, alpha=0.9, edgecolor="none", label="болота"))
    if data.hydro_lines is not None and not data.hydro_lines.empty:
        hl = gpd.clip(data.hydro_lines, u)
        if not hl.empty:
            hl.plot(ax=ax, color="#1F78B4", linewidth=1.1, alpha=0.9, zorder=2.5)
            handles.append(Line2D([], [], color="#1F78B4", linewidth=1.6, label="реки и канавы"))
    if data.dem is not None:
        try:
            arr, tr = data.dem
            inside = data.inside_mask(tr, arr.shape)
            dem_i = np.where(inside, arr, np.nan)
            if np.nanmax(dem_i) - np.nanmin(dem_i) > 2:
                step = 2.5
                levels = np.arange(np.floor(np.nanmin(dem_i) / step) * step,
                                    np.nanmax(dem_i) + step, step)
                nmin, nmax = tr.c, tr.f
                ax.contour(np.arange(dem_i.shape[1]) * tr.a + nmin,
                           nmax - np.arange(dem_i.shape[0]) * tr.e,
                           dem_i, levels=levels, colors="#9e9e9e", linewidths=0.4, alpha=0.75, zorder=3)
                handles.append(Line2D([], [], color="#9e9e9e", linewidth=0.9, label="горизонтали (2,5 м)"))
        except Exception as e:
            log.warning("  горизонтали не построены: %s", e)
    paved = _paved_roads(data)
    dirt = None
    if data.roads_dirt is not None and not data.roads_dirt.empty:
        dirt = data.roads_dirt
    elif (paved is not None and (data.roads_paved is None or data.roads_paved.empty)
          and data.trails is not None and not data.trails.empty):
        dirt = data.trails[~data.trails.index.isin(paved.index)]
    elif data.trails is not None and not data.trails.empty:
        dirt = data.trails
    if dirt is not None and not dirt.empty:
        tl = gpd.clip(dirt, u)
        if not tl.empty:
            tl.plot(ax=ax, color="#b5b5b5", linewidth=0.55, alpha=0.75, zorder=3.5)
            handles.append(Line2D([], [], color="#b5b5b5", linewidth=1.0, label="грунтовые дороги и тропы"))
    if data.trails_informal is not None and not data.trails_informal.empty:
        it = gpd.clip(data.trails_informal, u)
        if not it.empty:
            it.plot(ax=ax, color="#c98d5a", linewidth=0.8, alpha=0.85, zorder=3.6)
            handles.append(Line2D([], [], color="#c98d5a", linewidth=1.2, label=_inf_label()))
    if paved is not None and not paved.empty:
        pv = gpd.clip(paved, u)
        if not pv.empty:
            pv.plot(ax=ax, color="#5f6368", linewidth=1.7, zorder=4)
            handles.append(Line2D([], [], color="#5f6368", linewidth=2.2, label="дороги с твёрдым покрытием"))
    if cfg.unit_mode == "eco" and not others.empty:
        handles.append(Patch(facecolor="none", edgecolor="#9AA0A6", label="прочие полигоны"))
    used_layers = []
    pvn_lab_items = []
    if evidence is not None and not evidence.empty:
        is_pt = evidence.geom_type.isin(["Point", "MultiPoint"])
        has_kv = ("kv" in evidence.columns and "oozt" in evidence.columns
                  and evidence["kv"].fillna("").astype(str).str.strip().ne(""))
        ev_pts = evidence[is_pt & (
            (has_kv & (evidence["oozt"].astype(str) == str(oozt_name))
             & (evidence["kv"].astype(str) == str(kv)))
            | (~has_kv & evidence.intersects(focus_geom))
        )].copy()
        ev_nonpt = evidence[~is_pt & evidence.intersects(focus_geom)].copy()
        if not ev_nonpt.empty:
            # подрезаем участки работ строго к границе квартала (без захода на соседние кварталы)
            ev_nonpt = gpd.clip(ev_nonpt, focus_geom)
        ev_parts = [x for x in (ev_nonpt, ev_pts) if x is not None and not x.empty]
        if ev_parts:
            ev_plot = gpd.GeoDataFrame(pd.concat(ev_parts, ignore_index=True), crs=WORK_EPSG)
            for layer, grp in ev_plot.groupby("layer"):
                _plot_evidence_group(ax, handles, layer, grp, overview=False)
                if layer.startswith("лесопатология") and "label" in grp.columns:
                    for _, pr in grp.iterrows():
                        pv_id = str(pr.get("label") or "").strip()
                        if pv_id and pv_id not in ("nan", "None"):
                            pvn_lab_items.append((f"ПВН {pv_id}", pr.geometry.buffer(1.0)))
                used_layers.append(layer)
    # границы выделов — только в режиме полигонов экосистем; в режиме кварталов
    # границы выделов на схемах не показываются (учётная единица — квартал)
    ax.set_xlim(ext[0], ext[1]); ax.set_ylim(ext[2], ext[3])
    lab_items = []
    if cfg.unit_mode == "eco":
        vyd.boundary.plot(ax=ax, color="#222222", linewidth=1.6 if group_mode else 1.15, zorder=8)
        lab_items = [(str(r["VYD"]).removesuffix(".0"), r.geometry)
                     for _, r in vyd.assign(_a=vyd.geometry.area).sort_values("_a", ascending=False).iterrows()]
    # соседние кварталы / листы: границы и номера в пределах окна карты (в режиме индексов — не нужны)
    map_box = box(ext[0], ext[2], ext[1], ext[3])
    import matplotlib.patheffects as pe
    nb_geoms: list[tuple[str, object]] = []
    if not group_mode:
        if kn_oozt is not None and not kn_oozt.empty and cfg.unit_mode != "eco":
            nb_kn = kn_oozt[(kn_oozt["KV"].astype(str) != str(kv)) & kn_oozt.intersects(map_box)]
            for n_kv, ng in nb_kn.groupby("KV"):
                nb_geoms.append((str(n_kv).removesuffix(".0"), _safe_union(ng.geometry.values)))
        else:
            same_oozt = data.oozt[data.oozt["OOZT"].astype(str) == str(oozt_name)]
            neighbors = same_oozt[(same_oozt["KV"].astype(str) != str(kv)) & same_oozt.intersects(map_box)]
            for n_kv, ng in neighbors.groupby("KV"):
                nb_geoms.append((str(n_kv).removesuffix(".0"), _safe_union(ng.geometry.values)))
    for n_label, ngeom in nb_geoms:
        clipped = ngeom.intersection(map_box)
        if clipped.is_empty:
            continue
        gpd.GeoSeries([ngeom], crs=WORK_EPSG).boundary.plot(
            ax=ax, color="#444444", linewidth=1.5, linestyle="--", zorder=9)
        p = clipped.representative_point()
        px = float(np.clip(p.x, ext[0] + w * 0.07, ext[1] - w * 0.07))
        py = float(np.clip(p.y, ext[2] + h * (0.15 if px < ext[0] + w * 0.35 else 0.09), ext[3] - h * 0.07))
        ax.annotate(n_label, (px, py), ha="center", va="center", fontsize=13, fontweight="bold",
                    color="#555555", zorder=12, path_effects=[pe.withStroke(linewidth=3, foreground="white")])
    # граница квартала / листа в фокусе (у разбросанных полигонов индекса совпадает с их контурами)
    if not group_mode:
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
                                   attribution_size=9, zorder=0)
                    desaturate_last_basemap(ax)
                    break
                except Exception as e:     # таймаут/лимит запросов — пробуем следующий источник
                    log.warning("подложка %s недоступна (%s) — пробую другой источник", prov.get("name"), e)
        except Exception as e:
            log.warning("подложка OSM недоступна: %s", e)
    ax.legend(handles=handles, loc="upper left", bbox_to_anchor=(1.01, 1.0), fontsize=10.5, frameon=True,
              title="Легенда", title_fontsize=11)
    ax.set_title(f"{oozt_name}. {unit_labels(data.cfg)['group']} {kv}", fontsize=14.5, fontweight="bold")
    ax.tick_params(labelsize=11)
    try:
        set_degree_ticks(ax, ext)
    except Exception as e:                # запасной вариант — метры UTM
        log.warning("не удалось построить сетку в градусах (%s) — оси в метрах", e)
        ax.set_xlabel("E, м (UTM 37N)", fontsize=11); ax.set_ylabel("N, м (UTM 37N)", fontsize=11)
        ax.ticklabel_format(style="plain", useOffset=False)
    # масштабная линейка
    target = w / 4                        # «круглая» длина линейки: 1/2/5 × 10^k
    dec = 10 ** np.floor(np.log10(max(target, 1.0)))
    L = next(m * dec for m in (5, 2, 1) if m * dec <= target) if target >= dec else dec
    x0, y0 = ext[0] + w * 0.04, ext[2] + h * 0.04
    ax.plot([x0, x0 + L], [y0, y0], color="black", linewidth=4, zorder=14,
            path_effects=[pe.withStroke(linewidth=6.5, foreground="white")])
    ax.text(x0 + L / 2, y0 + h * 0.022, f"{int(L)} м", ha="center", va="bottom", fontsize=11.5, zorder=14,
            path_effects=[pe.withStroke(linewidth=3, foreground="white")])
    ax.set_aspect("equal", adjustable="box")
    fig.tight_layout()
    ax.set_xlim(ext[0], ext[1]); ax.set_ylim(ext[2], ext[3])
    ax.set_aspect("equal", adjustable="box")
    # подписи точек ПВН (если есть) и полигонов — последними, когда известен масштаб
    if pvn_lab_items:
        seen_pvn = set()
        uniq_pvn = []
        for lbl_p, gm_p in pvn_lab_items:
            key_p = (lbl_p, round(gm_p.centroid.x, 1), round(gm_p.centroid.y, 1))
            if key_p not in seen_pvn:
                seen_pvn.add(key_p)
                uniq_pvn.append((lbl_p, gm_p))
        _place_labels(ax, uniq_pvn, sizes=(9,), leaders=True, max_labels=40)
    n_lab = _place_labels(ax, lab_items, sizes=(13, 12, 11, 10, 9, 8) if cfg.unit_mode != "eco" else (10, 9, 8, 7),
                          leaders=True)
    if n_lab < len(lab_items):
        ax.text(0.995, 0.005, f"подписаны {n_lab} из {len(lab_items)} полигонов; полный перечень — в таблице",
                transform=ax.transAxes, ha="right", va="bottom", fontsize=9, color="#444444", zorder=14,
                path_effects=[pe.withStroke(linewidth=3, foreground="white")])
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


def units_table(data: Data, res: Results) -> pd.DataFrame:
    """Участки проведения мероприятий (uid × мероприятие) с точной географической площадью участка работ (га).
    Основа всех сводок по ООЗТ, типам, мероприятиям и кварталам."""
    cols = ["ООЗТ", "Квартал", "Выдел", "uid", "Тип мероприятий", "Мероприятие", "Пункт Приказа", "Площадь, га"]
    if not res.rows:
        return pd.DataFrame(columns=cols)
    stand_areas = dict(zip(data.oozt["uid"], data.oozt["area_ha"])) if data.oozt is not None else {}
    by_um: dict[tuple, dict] = {}
    for r in res.rows:
        uid = r.get("uid")
        if uid is None or (isinstance(uid, float) and pd.isna(uid)) or str(uid) == "":
            uid = f"{r['ООЗТ']}|{r['Квартал']}|{r['Выдел']}"
        m = str(r["Мероприятие"])
        key = (str(uid), m)
        rec = by_um.get(key)
        if rec is None:
            rec = {
                "ООЗТ": r["ООЗТ"], "Квартал": r["Квартал"], "Выдел": r["Выдел"], "uid": uid,
                "Тип мероприятий": r.get("Тип мероприятий", ""),
                "Мероприятие": m,
                "Пункт Приказа": r.get("Пункт Приказа", ""),
                "_geoms": [], "_ha_fb": 0.0,
            }
            by_um[key] = rec
        g = r.get("_geom")
        if g is not None and not g.is_empty:
            rec["_geoms"].append(g)
        else:
            rec["_ha_fb"] += float(r.get("_work_ha") or stand_areas.get(uid, 0.0) or 0.0)
    out_rows = []
    for rec in by_um.values():
        if rec["_geoms"]:
            ug = _safe_union(rec["_geoms"])
            ha = float(ug.area / 1e4) if ug is not None and not ug.is_empty else 0.0
        else:
            ha = rec["_ha_fb"]
        out_rows.append({
            "ООЗТ": rec["ООЗТ"], "Квартал": rec["Квартал"], "Выдел": rec["Выдел"], "uid": rec["uid"],
            "Тип мероприятий": rec["Тип мероприятий"], "Мероприятие": rec["Мероприятие"],
            "Пункт Приказа": rec["Пункт Приказа"], "Площадь, га": round(ha, 4),
        })
    return pd.DataFrame(out_rows, columns=cols)


def _oozt_areas(data: Data) -> dict:
    """Площадь каждого ООЗТ (га) и число кварталов сети: из квартальной сети
    (все кварталы, включая без мероприятий); без сети — сумма полигонов экосистем."""
    kn = getattr(data, "kvnet", None)
    if kn is not None and not kn.empty and "KVOOZT" in kn.columns:
        kv = kn.copy()
        kv["KVOOZT"] = kv["KVOOZT"].astype(str).str.strip().replace({"nan": "", "None": ""})
        if (kv["KVOOZT"] != "").any():
            kv["_a"] = kv.geometry.area / 1e4
            g = kv.groupby("KVOOZT").agg(areas=("_a", "sum"), kvs=("KV", "nunique"))
            return {o: (round(float(r.areas), 2), int(r.kvs)) for o, r in g.iterrows()}
    if data.oozt is not None and not data.oozt.empty:
        g = data.oozt.groupby(data.oozt["OOZT"].astype(str)).agg(
            areas=("area_ha", "sum"), kvs=("KV", "nunique"))
        return {o: (round(float(r.areas), 2), int(r.kvs)) for o, r in g.iterrows()}
    return {}


def _oozt_union_geoms(data: Data) -> dict[str, object]:
    """Геометрия контура каждого ООЗТ в WORK_EPSG (из квартальной сети или полигонов территории)."""
    qs = data.quarters()
    if qs is not None and not qs.empty and "OOZT" in qs.columns:
        return {str(oo): _safe_union(g.geometry.values) for oo, g in qs.groupby(qs["OOZT"].astype(str))}
    if data.oozt is not None and not data.oozt.empty:
        return {str(oo): _safe_union(g.geometry.values) for oo, g in data.oozt.groupby(data.oozt["OOZT"].astype(str))}
    return {}


def _covered_ha_by_group(data: Data, units: pd.DataFrame, group_cols: list[str],
                         res: "Results | None" = None,
                         measures_gdf: gpd.GeoDataFrame | None = None) -> dict[tuple, float]:
    """Истинная площадь без наложений (га) по ключу group_cols: через геометрическое объединение участков работ."""
    oo_geoms = _oozt_union_geoms(data)
    out: dict[tuple, float] = {}
    if res is not None and res.rows:
        buckets: dict[tuple, list] = {}
        for r in res.rows:
            g = r.get("_geom")
            if g is None or g.is_empty:
                continue
            k = tuple(str(r.get(c, "")) for c in group_cols)
            buckets.setdefault(k, []).append(g)
        for k, gs in buckets.items():
            ug = _safe_union(gs)
            if "ООЗТ" in group_cols:
                oo = k[group_cols.index("ООЗТ")]
                if oo in oo_geoms and oo_geoms[oo] is not None:
                    ug = ug.intersection(oo_geoms[oo])
            out[k] = round(float(ug.area / 1e4), 2) if ug is not None and not ug.is_empty else 0.0
        return out
    if measures_gdf is not None and not measures_gdf.empty and all(c in measures_gdf.columns for c in group_cols):
        mg = measures_gdf.to_crs(WORK_EPSG) if measures_gdf.crs and measures_gdf.crs.to_epsg() != WORK_EPSG else measures_gdf
        for k_vals, grp in mg.groupby([mg[c].astype(str) for c in group_cols]):
            k = (k_vals,) if len(group_cols) == 1 and not isinstance(k_vals, tuple) else tuple(k_vals)
            ug = _safe_union(grp.geometry.values)
            if "ООЗТ" in group_cols:
                oo = k[group_cols.index("ООЗТ")]
                if oo in oo_geoms and oo_geoms[oo] is not None:
                    ug = ug.intersection(oo_geoms[oo])
            out[k] = round(float(ug.area / 1e4), 2) if ug is not None and not ug.is_empty else 0.0
        return out
    if units is not None and not units.empty and all(c in units.columns for c in group_cols):
        u = units.copy()
        for c in group_cols:
            u[c] = u[c].astype(str)
        u["Площадь, га"] = u["Площадь, га"].astype(float)
        stand_a = dict(zip(data.oozt["uid"].astype(str), data.oozt["area_ha"].astype(float))) if data.oozt is not None else {}
        # внутри одного выдела сумма участков разных мер ограничена площадью самого выдела
        by_uid = u.groupby(group_cols + ["uid"], as_index=False)["Площадь, га"].sum()
        by_uid["Площадь, га"] = [
            min(float(a), float(stand_a.get(str(uid), a))) for uid, a in zip(by_uid["uid"], by_uid["Площадь, га"])
        ]
        for k_vals, grp in by_uid.groupby(group_cols):
            k = (k_vals,) if len(group_cols) == 1 and not isinstance(k_vals, tuple) else tuple(k_vals)
            out[k] = round(float(grp["Площадь, га"].sum()), 2)
    return out


def oozt_summary(data: Data, table: pd.DataFrame, units: pd.DataFrame,
                 res: "Results | None" = None, measures_gdf: gpd.GeoDataFrame | None = None,
                 include_total: bool = True) -> pd.DataFrame:
    """Сводка по каждому ООЗТ: общая площадь, кварталы, мероприятия в штуках, гектарах и % от площади ООЗТ."""
    cols = [
        "ООЗТ", "Площадь ООЗТ, га", "Кварталов всего, шт",
        "Кварталов с мероприятиями, шт", "Кварталов без мероприятий, шт", "Доля кварталов с мероприятиями, %",
        "Типов мероприятий, шт", "Уникальных мероприятий, шт",
        "Назначений (квартал × мероприятие), шт", "Участков проведения работ, шт",
        "Площадь с мероприятиями (без наложений), га", "Покрытие, % от площади ООЗТ",
        "Площадь без мероприятий, га", "Без мероприятий, % от площади ООЗТ",
        "Суммарная площадь назначений (с наложениями), га", "Назначений, % от площади ООЗТ",
    ]
    areas = _oozt_areas(data)
    t = table if table is not None and not table.empty else pd.DataFrame()
    u = units if units is not None and not units.empty else pd.DataFrame()
    cov_map = _covered_ha_by_group(data, u, ["ООЗТ"], res=res, measures_gdf=measures_gdf)
    all_oos = list(dict.fromkeys(
        list(areas.keys())
        + (t["ООЗТ"].astype(str).tolist() if not t.empty else [])
        + (u["ООЗТ"].astype(str).tolist() if not u.empty else [])
    ))
    if not all_oos:
        return pd.DataFrame(columns=cols)
    recs = []
    for o in sorted(all_oos):
        a_tot, n_kv = areas.get(o, (0.0, 0))
        tg = t[t["ООЗТ"].astype(str) == o] if not t.empty else pd.DataFrame()
        ug = u[u["ООЗТ"].astype(str) == o] if not u.empty else pd.DataFrame()
        n_asg = int(len(tg))
        n_uniq = int(tg["Мероприятие"].nunique()) if not tg.empty else 0
        n_types = int(tg["Тип мероприятий"].nunique()) if (not tg.empty and "Тип мероприятий" in tg.columns) else 0
        n_kv_m = int(tg["Квартал"].nunique()) if (not tg.empty and "Квартал" in tg.columns) else 0
        n_kv_empty = max(0, n_kv - n_kv_m)
        pct_kv = round(100.0 * n_kv_m / n_kv, 1) if n_kv else None
        n_patches = int(len(ug))
        cov_ha = min(round(float(cov_map.get((o,), 0.0)), 2), round(float(a_tot), 2)) if a_tot else round(float(cov_map.get((o,), 0.0)), 2)
        sum_ha = round(float(ug["Площадь, га"].astype(float).sum()), 2) if not ug.empty else 0.0
        free_ha = round(max(0.0, float(a_tot) - cov_ha), 2) if a_tot else 0.0
        pct_cov = round(100.0 * cov_ha / a_tot, 2) if a_tot else None
        pct_free = round(100.0 * free_ha / a_tot, 2) if a_tot else None
        pct_sum = round(100.0 * sum_ha / a_tot, 2) if a_tot else None
        recs.append([
            o, round(float(a_tot), 2), int(n_kv), int(n_kv_m), int(n_kv_empty), pct_kv,
            n_types, n_uniq, n_asg, n_patches,
            cov_ha, pct_cov, free_ha, pct_free, sum_ha, pct_sum,
        ])
    df = pd.DataFrame(recs, columns=cols)
    if include_total and len(df) > 1:
        tot_a = round(float(df["Площадь ООЗТ, га"].sum()), 2)
        tot_kv = int(df["Кварталов всего, шт"].sum())
        tot_kvm = int(df["Кварталов с мероприятиями, шт"].sum())
        tot_kve = int(df["Кварталов без мероприятий, шт"].sum())
        tot_cov = round(float(df["Площадь с мероприятиями (без наложений), га"].sum()), 2)
        tot_free = round(max(0.0, tot_a - tot_cov), 2)
        tot_sum = round(float(df["Суммарная площадь назначений (с наложениями), га"].sum()), 2)
        tot_row = pd.DataFrame([[
            "ИТОГО", tot_a, tot_kv, tot_kvm, tot_kve,
            round(100.0 * tot_kvm / tot_kv, 1) if tot_kv else None,
            int(t["Тип мероприятий"].nunique()) if (not t.empty and "Тип мероприятий" in t.columns) else 0,
            int(t["Мероприятие"].nunique()) if not t.empty else 0,
            int(df["Назначений (квартал × мероприятие), шт"].sum()),
            int(df["Участков проведения работ, шт"].sum()),
            tot_cov, round(100.0 * tot_cov / tot_a, 2) if tot_a else None,
            tot_free, round(100.0 * tot_free / tot_a, 2) if tot_a else None,
            tot_sum, round(100.0 * tot_sum / tot_a, 2) if tot_a else None,
        ]], columns=cols)
        df = pd.concat([df, tot_row], ignore_index=True)
    return df


def oozt_measure_summary(data: Data, table: pd.DataFrame, units: pd.DataFrame,
                         res: "Results | None" = None,
                         measures_gdf: gpd.GeoDataFrame | None = None) -> pd.DataFrame:
    """Сводка ООЗТ × мероприятие: в штуках (кварталы, участки), гектарах и % от общей площади ООЗТ."""
    cols = [
        "ООЗТ", "Площадь ООЗТ, га", "Тип мероприятий", "Мероприятие", "Пункт Приказа",
        "Кварталов с мероприятием, шт", "% кварталов ООЗТ",
        "Участков проведения, шт", "Объектов, шт", "Площадь проведения, га", "% от площади ООЗТ",
    ]
    if units is None or units.empty:
        return pd.DataFrame(columns=cols)
    areas = _oozt_areas(data)
    # «Объектов, шт» — уникальные объекты справочника водоёмов и болот (правило water_ref)
    n_obj: dict = {}
    for r_ in (res.rows if res is not None else []):
        if r_.get("_obj") is not None:
            n_obj.setdefault((str(r_["ООЗТ"]), str(r_["Мероприятие"])), set()).add(r_["_obj"])
    cov_om = _covered_ha_by_group(data, units, ["ООЗТ", "Мероприятие"], res=res, measures_gdf=measures_gdf)
    u = units.copy()
    u["ООЗТ"] = u["ООЗТ"].astype(str)
    u["Площадь, га"] = u["Площадь, га"].astype(float)
    if "Тип мероприятий" not in u.columns and table is not None and not table.empty and "Тип мероприятий" in table.columns:
        mt_map = dict(zip(table["Мероприятие"], table["Тип мероприятий"]))
        u["Тип мероприятий"] = u["Мероприятие"].map(mt_map).fillna("")
    elif "Тип мероприятий" not in u.columns:
        u["Тип мероприятий"] = ""
    if "Пункт Приказа" not in u.columns and table is not None and not table.empty and "Пункт Приказа" in table.columns:
        pr_map = dict(zip(table["Мероприятие"], table["Пункт Приказа"]))
        u["Пункт Приказа"] = u["Мероприятие"].map(pr_map).fillna("")
    elif "Пункт Приказа" not in u.columns:
        u["Пункт Приказа"] = ""
    g = u.groupby(["ООЗТ", "Мероприятие"], as_index=False).agg(
        **{
            "Тип мероприятий": ("Тип мероприятий", "first"),
            "Пункт Приказа": ("Пункт Приказа", "first"),
            "Участков проведения, шт": ("uid", "size"),
            "Кварталов с мероприятием, шт": ("Квартал", "nunique"),
            "_sum_ha": ("Площадь, га", "sum"),
        }
    )
    g["Площадь ООЗТ, га"] = [areas.get(o, (0.0, 0))[0] for o in g["ООЗТ"]]
    g["Объектов, шт"] = [len(n_obj[(str(o), str(m))]) if (str(o), str(m)) in n_obj else None
                        for o, m in zip(g["ООЗТ"], g["Мероприятие"])]
    g["% кварталов ООЗТ"] = [
        round(100.0 * k / areas[o][1], 1) if o in areas and areas[o][1] else None
        for o, k in zip(g["ООЗТ"], g["Кварталов с мероприятием, шт"])
    ]
    g["Площадь проведения, га"] = [
        cov_om.get((str(o), str(m)), round(float(sh), 2))
        for o, m, sh in zip(g["ООЗТ"], g["Мероприятие"], g["_sum_ha"])
    ]
    g["% от площади ООЗТ"] = [
        round(100.0 * a / areas[o][0], 2) if o in areas and areas[o][0] else None
        for o, a in zip(g["ООЗТ"], g["Площадь проведения, га"])
    ]
    order = {}
    if table is not None and not table.empty:
        for i, m in enumerate(table["Мероприятие"].drop_duplicates()):
            order[m] = i
    g = g.sort_values(["ООЗТ", "Мероприятие"],
                      key=lambda s: s.map(order) if s.name == "Мероприятие" else s)
    return g[cols].reset_index(drop=True)


def oozt_type_summary(data: Data, table: pd.DataFrame, units: pd.DataFrame,
                      res: "Results | None" = None,
                      measures_gdf: gpd.GeoDataFrame | None = None) -> pd.DataFrame:
    """Сводка ООЗТ × тип мероприятий: мероприятия и назначения в штуках, площадь в га и % от ООЗТ."""
    cols = [
        "ООЗТ", "Площадь ООЗТ, га", "Тип мероприятий",
        "Уникальных мероприятий, шт", "Назначений (квартал × мероприятие), шт",
        "Кварталов с данным типом, шт", "Участков проведения, шт",
        "Площадь проведения (без наложений), га", "% от площади ООЗТ",
        "Суммарная площадь мер типа (с наложениями), га",
    ]
    if units is None or units.empty:
        return pd.DataFrame(columns=cols)
    areas = _oozt_areas(data)
    u = units.copy()
    u["ООЗТ"] = u["ООЗТ"].astype(str)
    u["Площадь, га"] = u["Площадь, га"].astype(float)
    if "Тип мероприятий" not in u.columns:
        mt_map = (dict(zip(table["Мероприятие"], table["Тип мероприятий"]))
                  if table is not None and not table.empty and "Тип мероприятий" in table.columns else {})
        u["Тип мероприятий"] = u["Мероприятие"].map(mt_map).fillna("")
    cov_ot = _covered_ha_by_group(data, u, ["ООЗТ", "Тип мероприятий"], res=res, measures_gdf=measures_gdf)
    asg_cnt = {}
    if table is not None and not table.empty and "Тип мероприятий" in table.columns:
        for (o, tp), grp in table.groupby([table["ООЗТ"].astype(str), table["Тип мероприятий"].astype(str)]):
            asg_cnt[(o, tp)] = int(len(grp))
    rows = []
    for (o, tp), grp in u.groupby(["ООЗТ", "Тип мероприятий"]):
        a_tot = areas.get(str(o), (0.0, 0))[0]
        cov_ha = cov_ot.get((str(o), str(tp)), round(float(grp["Площадь, га"].sum()), 2))
        if a_tot:
            cov_ha = min(cov_ha, round(float(a_tot), 2))
        sum_ha = round(float(grp["Площадь, га"].sum()), 2)
        rows.append([
            str(o), round(float(a_tot), 2), str(tp),
            int(grp["Мероприятие"].nunique()),
            asg_cnt.get((str(o), str(tp)), int(grp[["Квартал", "Мероприятие"]].drop_duplicates().shape[0])),
            int(grp["Квартал"].nunique()),
            int(len(grp)),
            cov_ha,
            round(100.0 * cov_ha / a_tot, 2) if a_tot else None,
            sum_ha,
        ])
    return pd.DataFrame(rows, columns=cols).sort_values(["ООЗТ", "Тип мероприятий"]).reset_index(drop=True)


def oozt_kvartal_summary(data: Data, table: pd.DataFrame, units: pd.DataFrame,
                         res: "Results | None" = None,
                         measures_gdf: gpd.GeoDataFrame | None = None) -> pd.DataFrame:
    """Поквартальная сводка по каждому ООЗТ (включая кварталы без мероприятий): штуки, га, % от квартала и % от ООЗТ."""
    cols = [
        "ООЗТ", "Квартал", "Площадь квартала, га",
        "Мероприятий, шт", "Участков проведения, шт",
        "Площадь с мероприятиями, га", "% площади квартала",
        "Суммарная площадь назначений, га", "% от площади ООЗТ", "Перечень мероприятий",
    ]
    qs = data.quarters()
    if qs is None or qs.empty:
        return pd.DataFrame(columns=cols)
    areas = _oozt_areas(data)
    u = units if units is not None and not units.empty else pd.DataFrame()
    t = table if table is not None and not table.empty else pd.DataFrame()
    cov_ok = _covered_ha_by_group(data, u, ["ООЗТ", "Квартал"], res=res, measures_gdf=measures_gdf)
    rows = []
    for _, qr in qs.iterrows():
        o = str(qr["OOZT"])
        kv = qr["KV"]
        kvs = str(kv)
        kv_ha = round(float(qr.get("area_ha", qr.geometry.area / 1e4)), 2)
        oo_ha = areas.get(o, (0.0, 0))[0]
        tg = t[(t["ООЗТ"].astype(str) == o) & (t["Квартал"].astype(str) == kvs)] if not t.empty else pd.DataFrame()
        ug = u[(u["ООЗТ"].astype(str) == o) & (u["Квартал"].astype(str) == kvs)] if not u.empty else pd.DataFrame()
        m_list = list(dict.fromkeys(tg["Мероприятие"].astype(str).tolist())) if not tg.empty else []
        cov_ha = min(cov_ok.get((o, kvs), 0.0), kv_ha) if kv_ha else cov_ok.get((o, kvs), 0.0)
        sum_ha = round(float(ug["Площадь, га"].astype(float).sum()), 2) if not ug.empty else 0.0
        rows.append([
            o, kv, kv_ha,
            len(m_list), int(len(ug)),
            round(float(cov_ha), 2),
            round(100.0 * cov_ha / kv_ha, 2) if kv_ha else None,
            sum_ha,
            round(100.0 * cov_ha / oo_ha, 2) if oo_ha else None,
            "; ".join(m_list),
        ])
    return pd.DataFrame(rows, columns=cols)


def oozt_pivot_summary(data: Data, ms: pd.DataFrame) -> pd.DataFrame:
    """Перекрёстная матрица: строки — все назначенные мероприятия, столбцы — по каждому ООЗТ (кварталов шт., га, % от ООЗТ)."""
    base_cols = ["Тип мероприятий", "Мероприятие", "Пункт Приказа"]
    if ms is None or ms.empty:
        return pd.DataFrame(columns=base_cols)
    oos = sorted(ms["ООЗТ"].astype(str).unique())
    m_meta = ms[base_cols].drop_duplicates(["Мероприятие"]).reset_index(drop=True)
    idx_map = {(str(r["ООЗТ"]), str(r["Мероприятие"])): r for _, r in ms.iterrows()}
    out_rows = []
    for _, mr in m_meta.iterrows():
        m = str(mr["Мероприятие"])
        row = [mr["Тип мероприятий"], m, mr["Пункт Приказа"]]
        tot_kv, tot_ha = 0, 0.0
        for o in oos:
            rec = idx_map.get((o, m))
            if rec is not None:
                kv_n = int(rec["Кварталов с мероприятием, шт"])
                ha = round(float(rec["Площадь проведения, га"]), 2)
                pct = rec["% от площади ООЗТ"]
                tot_kv += kv_n
                tot_ha += ha
                row.extend([kv_n, ha, pct])
            else:
                row.extend([0, 0.0, 0.0])
        if len(oos) > 1:
            tot_area = sum(_oozt_areas(data).get(o, (0.0, 0))[0] for o in oos)
            row.extend([tot_kv, round(tot_ha, 2), round(100.0 * tot_ha / tot_area, 2) if tot_area else None])
        out_rows.append(row)
    cols = list(base_cols)
    for o in oos:
        cols.extend([f"{o}: кварталов, шт", f"{o}: площадь, га", f"{o}: % от площади ООЗТ"])
    if len(oos) > 1:
        cols.extend(["ВСЕГО: кварталов, шт", "ВСЕГО: площадь, га", "ВСЕГО: % от общей площади"])
    return pd.DataFrame(out_rows, columns=cols)


def _format_summary_sheets(book, sheet_names: list[str]):
    """Единообразное оформление сводных листов Excel (шапка, автофильтр, форматы чисел, автоширина)."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    hdr_fill = PatternFill("solid", fgColor="D9EAD3")
    tot_fill = PatternFill("solid", fgColor="EFEFEF")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for sh in sheet_names:
        if sh not in book.sheetnames:
            continue
        ws = book[sh]
        ws.freeze_panes = "A2"
        if ws.max_row >= 1 and ws.max_column >= 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        headers = [str(c.value or "") for c in ws[1]]
        for cell in ws[1]:
            cell.font = Font(bold=True, size=10)
            cell.fill = hdr_fill
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.row_dimensions[1].height = 34
        for r_idx in range(2, ws.max_row + 1):
            is_total = str(ws.cell(r_idx, 1).value or "") == "ИТОГО"
            for c_idx in range(1, ws.max_column + 1):
                cell = ws.cell(r_idx, c_idx)
                hdr = headers[c_idx - 1]
                is_num = isinstance(cell.value, (int, float))
                cell.border = border
                if is_total:
                    cell.font = Font(bold=True, size=10)
                    cell.fill = tot_fill
                else:
                    cell.font = Font(size=10)
                if is_num:
                    if "га" in hdr or "%" in hdr:
                        cell.number_format = "0.00"
                    else:
                        cell.number_format = "#,##0"
                    cell.alignment = Alignment(vertical="top", horizontal="right")
                else:
                    cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        for c_idx, hdr in enumerate(headers, 1):
            col_let = get_column_letter(c_idx)
            if "Перечень" in hdr:
                w = 56
            elif hdr == "Мероприятие":
                w = 46
            elif hdr in ("ООЗТ", "Тип мероприятий"):
                w = 28
            elif "га" in hdr or "%" in hdr or "шт" in hdr:
                w = max(14, min(24, len(hdr) // 2 + 6))
            else:
                w = max(12, min(32, len(hdr) + 4))
            ws.column_dimensions[col_let].width = w


def write_oozt_summary_excel(path: Path, data: Data, table: pd.DataFrame, units: pd.DataFrame,
                             res: "Results | None" = None,
                             measures_gdf: gpd.GeoDataFrame | None = None) -> dict[str, pd.DataFrame]:
    """Создаёт отдельную сводную книгу Excel со всеми разрезами по каждому ООЗТ (шт., га, % от площади ООЗТ)."""
    sm = oozt_summary(data, table, units, res=res, measures_gdf=measures_gdf, include_total=True)
    ms = oozt_measure_summary(data, table, units, res=res, measures_gdf=measures_gdf)
    ts = oozt_type_summary(data, table, units, res=res, measures_gdf=measures_gdf)
    ks = oozt_kvartal_summary(data, table, units, res=res, measures_gdf=measures_gdf)
    pv = oozt_pivot_summary(data, ms)
    sheets = [
        ("Сводка по ООЗТ", sm),
        ("ООЗТ × мероприятие", ms),
        ("ООЗТ × тип мероприятий", ts),
        ("Сравнительная матрица", pv),
        ("ООЗТ × кварталы", ks),
    ]
    written = []
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for sh_name, df_s in sheets:
            if df_s is not None and (not df_s.empty or sh_name == "Сводка по ООЗТ"):
                df_s.to_excel(xw, index=False, sheet_name=sh_name)
                written.append(sh_name)
        _format_summary_sheets(xw.book, written)
    return {"sm": sm, "ms": ms, "ts": ts, "ks": ks, "pv": pv}


def _report_groups_table(data: Data, kv_list) -> pd.DataFrame:
    """Список групп отчёта (кварталов / листов / групп) в порядке KV_SORT с учётом фильтра."""
    kv_all = (data.oozt[["OOZT", "KV", "KV_SORT"]].drop_duplicates(["OOZT", "KV"])
              .sort_values(["OOZT", "KV_SORT"]))
    # kv_cut: в отчёт включаются ВСЕ кварталы сети — и те, куда не попало ни одного
    # полигона экосистем (для них строятся схемы и разделы без мероприятий)
    kn = getattr(data, "kvnet", None)
    if data.cfg.unit_mode == "kv_cut" and kn is not None and not kn.empty:
        c_ = data.cfg
        fb = (c_.oozt_name or "").strip() or Path(c_.oozt_path).stem
        if kn["KVOOZT"].replace("", pd.NA).notna().any():
            kn_oo = kn["KVOOZT"].where(kn["KVOOZT"] != "", fb).astype(str)
        else:
            u_oo = data.oozt["OOZT"].astype(str).unique()
            kn_oo = pd.Series(str(u_oo[0]) if len(u_oo) == 1 else fb, index=kn.index)
        have_pairs = set(zip(kv_all["OOZT"].astype(str), kv_all["KV"].astype(str)))
        miss_mask = [(o, str(k)) not in have_pairs for o, k in zip(kn_oo, kn["KV"])]
        miss = kn[miss_mask]
        if not miss.empty:
            extra = pd.DataFrame({"OOZT": kn_oo[miss_mask].astype(str).values,
                                  "KV": _maybe_int(miss["KV"]).values})
            extra = extra.drop_duplicates(["OOZT", "KV"])
            extra["KV_SORT"] = _nat_rank(extra["KV"])
            kv_all = pd.concat([kv_all, extra], ignore_index=True)
            kv_all = kv_all.sort_values(["OOZT", "KV_SORT"]).reset_index(drop=True)
            log.info("Кварталы сети без полигонов экосистем (в отчёт добавлены пустыми): %s",
                     ", ".join(f"{o} кв.{k}" for o, k in zip(extra["OOZT"], extra["KV"])))
    if kv_list:
        want = {str(x).strip() for x in kv_list}
        kv_all = kv_all[kv_all["KV"].astype(str).isin(want)]
        miss = want - set(kv_all["KV"].astype(str))
        if miss:
            log.warning("В слое нет запрошенных для отчёта значений (%s): %s",
                        unit_labels(data.cfg)["groups_n"].lower(), ", ".join(sorted(miss)))
    return kv_all


def _report_rename(cfg: Config) -> dict:
    L = unit_labels(cfg)
    return {"Экосистема": L["unit_name"], "Мероприятие": "Мероприятия",
            "Тип экосистемы мероприятия": "Тип экосистемы", "Пункт Приказа": "Соответствие Пункту Приказа",
            "Выдел": L["unit"]}


def _fig_name(oozt, kv) -> str:
    safe = lambda x: re.sub(r"[^\w]+", "_", str(x))
    return "fig_" + safe(oozt) + "_kv_" + safe(kv) + ".png"


def _params_line() -> str:
    return ("Параметры каждого мероприятия (правило, расстояния, окна, квантили, пороги) заданы в таблице "
            "мероприятий и перечислены в разделе «Как рассчитаны мероприятия»; верхний квартиль считается по всем "
            "ООЗТ; общий порог доли полигона — %g%% (если в таблице не указан свой)." % _MIN_SHARE[0])


def _summary_table(table: pd.DataFrame, L: dict) -> pd.DataFrame:
    """Сводка: на скольких единицах назначено мероприятие (кварталы или выделы/полигоны)."""
    if "Выдел" in table.columns:
        return table.groupby(["Тип мероприятий", "Мероприятие"]).agg(
            **{L["units_n"]: ("Выдел", "size"), L["groups_n"]: ("Квартал", "nunique")}).reset_index()
    return table.groupby(["Тип мероприятий", "Мероприятие"]).agg(
        **{"Назначений": ("Квартал", "size"), L["groups_n"]: ("Квартал", "nunique")}).reset_index()


REPORT_FORMATS = {"docx": "Word (DOCX)", "pdf": "PDF", "doc": "Word 97–2003 (DOC)"}
PAGE_W_CM, PAGE_H_CM, MARGIN_CM = 29.7, 21.0, 1.5          # А4 альбомная
TEXT_W_CM = PAGE_W_CM - 2 * MARGIN_CM


def rules_table(measures, table: pd.DataFrame, cfg: Config) -> pd.DataFrame:
    """Как рассчитано каждое включённое мероприятие: метод из таблицы → правило скрипта → число полигонов.
    Позволяет сверить, что название мероприятия и основание назначения соответствуют друг другу."""
    _MIN_SHARE[0] = float(getattr(cfg, "min_share_pct", 0) or 0)
    rows = []
    cnt = table["Мероприятие"].value_counts() if not table.empty else pd.Series(dtype=int)
    for m in measures or []:
        if not m.get("on"):
            continue
        n = int(cnt.get(m["name"], 0))
        extra = [k for k in cnt.index if str(k).startswith(m["name"] + " — ")]
        note = ""
        if extra:
            note = "; " + "; ".join(f"{str(k)[len(m['name']) + 3:]}: {int(cnt[k])}" for k in extra)
        st = measure_status(m, cfg)
        rows.append({
            "Код": m.get("code", ""),
            "Мероприятие": m["name"],
            "Пункт Приказа": m.get("prikaz", ""),
            "Тип экосистемы (по таблице)": m.get("eco", ""),
            "Метод выявления (из таблицы)": m.get("method") or "",
            "Правило": (RULES[st["rule"]]["title"] if st["rule"] else "—") + (
                " (по тексту метода)" if st["spec"].get("origin") == "текст метода" and st["rule"] else ""),
            "Как рассчитано скриптом": describe_rule(m),
            "Назначено полигонов": f"{n}{note}" if st["state"] != "missing" else f"0 ({st['text']})",
        })
    return pd.DataFrame(rows)


def execution_list(table: pd.DataFrame, cfg: Config, data: "Data | None" = None) -> tuple[pd.DataFrame, str]:
    """Перечень мероприятий к выполнению: одна строка на полигон, мероприятия с пунктами Приказа списком.
    Порядок — по возрастанию поля группировки (индекса), затем ID; для листов сетки и кварталов — см. ниже.
    Возвращает (таблица с подписями колонок, описание порядка)."""
    L = unit_labels(cfg)
    if table is None or table.empty:
        return pd.DataFrame(), ""
    t = table.copy()
    def _with_clause(m, pk):
        pk = "" if pd.isna(pk) else str(pk).strip().rstrip(".")
        if pk.lower() in ("", "nan", "нет", "-", "—"):
            return str(m)
        return f"{m} (п. {pk})" if re.match(r"^\d", pk) else f"{m} ({pk})"
    t["_m"] = [_with_clause(m, pk) for pk, m in zip(t["Пункт Приказа"], t["Мероприятие"])]
    t["_tk"] = t["Тип мероприятий"].astype(str)
    t = t.sort_values(["_tk", "Пункт Приказа", "Мероприятие"], key=_nat_sort_key("Пункт Приказа"))
    qmode = cfg.unit_mode != "eco"      # учётная единица — квартал
    if qmode:
        keys = ["ООЗТ", "Квартал"]
        lst = t.groupby(keys, as_index=False, sort=False, dropna=False).agg(
            **{"Мероприятия": ("_m", "\n".join), "Число": ("Мероприятие", "size")})
        lst = lst.sort_values(["ООЗТ", "Квартал"], key=_nat_sort_key("Квартал")).reset_index(drop=True)
        cols = ["Квартал", "Мероприятия", "Число"]
        if lst["ООЗТ"].nunique() > 1:
            cols = ["ООЗТ"] + cols
        out = lst[cols].rename(columns={"Квартал": L["group"], "Мероприятия": "Мероприятия (пункт Приказа)",
                                        "Число": "Число мероприятий"})
        return out, f"по возрастанию «{L['group']}»"
    keys = ["ООЗТ", "Квартал", "Выдел"]
    lst = t.groupby(keys, as_index=False, sort=False, dropna=False).agg(
        **{"Экосистема": ("Экосистема", "first"), "Мероприятия": ("_m", "\n".join),
           "Число": ("Мероприятие", "size")})
    grid = cfg.unit_mode == "eco" and not cfg.report_group_field
    same = bool(cfg.unit_mode == "eco" and cfg.report_group_field and data is not None
                and cfg.report_group_field == getattr(data, "F_ID", None))
    order = ["Выдел", "Квартал"] if grid else ["Квартал", "Выдел"]
    lst = lst.sort_values(["ООЗТ"] + order, key=_nat_sort_key("Квартал", "Выдел")).reset_index(drop=True)
    cols = (["Выдел", "Квартал"] if grid else ["Квартал"] + ([] if same else ["Выдел"])) + \
           ["Экосистема", "Мероприятия", "Число"]
    if lst["ООЗТ"].nunique() > 1:
        cols = ["ООЗТ"] + cols
    out = lst[cols].rename(columns={"Квартал": L["group"], "Выдел": L["unit"], "Экосистема": L["unit_name"],
                                    "Мероприятия": "Мероприятия (пункт Приказа)", "Число": "Число мероприятий"})
    if grid:
        how = f"по возрастанию «{L['unit']}»"
    elif same:
        how = f"по возрастанию «{L['group']}»"
    else:
        how = f"по возрастанию «{L['group']}», внутри — по «{L['unit']}»"
    return out, how


def _work_geoms_by_oozt_measure(res: "Results | None",
                                measures_gdf: gpd.GeoDataFrame | None) -> dict[tuple[str, str], list]:
    """Геометрии участков проведения работ по ключу (ООЗТ, Мероприятие) для частных карт."""
    out: dict[tuple[str, str], list] = {}
    if res is not None and res.rows:
        for r in res.rows:
            g = r.get("_geom")
            if g is not None and not g.is_empty:
                out.setdefault((str(r.get("ООЗТ", "")), str(r.get("Мероприятие", ""))), []).append(g)
    elif (measures_gdf is not None and not measures_gdf.empty
          and {"ООЗТ", "Мероприятие"}.issubset(measures_gdf.columns)):
        mg = (measures_gdf.to_crs(WORK_EPSG)
              if measures_gdf.crs and measures_gdf.crs.to_epsg() != WORK_EPSG else measures_gdf)
        for (oo, m), grp in mg.groupby([mg["ООЗТ"].astype(str), mg["Мероприятие"].astype(str)]):
            gs = [g for g in grp.geometry.values if g is not None and not g.is_empty]
            if gs:
                out[(str(oo), str(m))] = gs
    return out


def _ordered_oozt_measures(data: Data, sub_oo: pd.DataFrame) -> list[tuple[str, str]]:
    """Упорядоченный список (Мероприятие, Код) для ООЗТ в порядке справочника мероприятий."""
    if sub_oo is None or sub_oo.empty or "Мероприятие" not in sub_oo.columns:
        return []
    present = list(dict.fromkeys(sub_oo["Мероприятие"].astype(str)))
    m_list = getattr(data, "measures", None) or []
    rank_map: dict[str, tuple[int, int, str]] = {}
    for i, m in enumerate(m_list):
        nm = str(m.get("name", ""))
        cd = str(m.get("code", "") or "").strip()
        if nm and nm not in rank_map:
            rank_map[nm] = (i, 0, cd)
    items = []
    for j, nm in enumerate(present):
        if nm in rank_map:
            rk, sub, cd = rank_map[nm]
            items.append((rk, sub, j, nm, cd))
        elif " — " in nm and nm.split(" — ", 1)[0] in rank_map:
            rk, _, cd = rank_map[nm.split(" — ", 1)[0]]
            items.append((rk, 1, j, nm, cd))
        else:
            items.append((10_000 + j, 0, j, nm, ""))
    items.sort(key=lambda t: (t[0], t[1], t[2]))
    return [(nm, cd) for _, _, _, nm, cd in items]
def _report_blocks(data: Data, table: pd.DataFrame, evidence, rep: Path, kv_all: pd.DataFrame,
                   basemap=True, draw=True, units: pd.DataFrame | None = None,
                   res: "Results | None" = None,
                   measures_gdf: gpd.GeoDataFrame | None = None) -> list:
    """Содержание отчёта в нейтральном виде; из него собираются MD, DOCX и PDF.
    Блоки: ("title", t) ("h", level, t) ("p", t) ("note", t) ("img", path) ("break",)
           ("table", header, rows, widths_cm, font_pt)"""
    cfg = data.cfg
    L = unit_labels(cfg)
    qmode = cfg.unit_mode != "eco"      # учётная единица — квартал
    mode_txt = UNIT_MODES[cfg.unit_mode]
    if cfg.unit_mode == "eco" and not cfg.report_group_field:
        mode_txt += f"; листы сетки {cfg.sheet_size_m:.0f} × {cfg.sheet_size_m:.0f} м"
    if qmode:
        n_units = table[["ООЗТ", "Квартал"]].drop_duplicates().shape[0] if not table.empty else 0
        units_txt = f"{L['groups_n'].lower()} с мероприятиями: {n_units}"
    else:
        n_units = table[["ООЗТ", "Квартал", "Выдел"]].drop_duplicates().shape[0] if not table.empty else 0
        units_txt = f"{L['units_gen']} с мероприятиями: {n_units}"
    B = [("title", f"Природоохранные мероприятия по {L['group_pl']} ООЗТ"),
         ("p", f"Слой территории: {Path(cfg.oozt_path).name} ({mode_txt}); "
               f"таблица мероприятий: {Path(cfg.measures_xlsx).name}."),
         ("p", f"Всего назначений: {len(table)}; {units_txt}; "
               f"{L['groups_n'].lower()} в отчёте: {len(kv_all)}."),
         ("p", _params_line())]
    s = lambda v: "" if pd.isna(v) else str(v)
    ms_all = pd.DataFrame()
    if not table.empty:
        B.append(("h", 1, "Сводка по мероприятиям"))
        u = units if units is not None and not units.empty else None
        sm = oozt_summary(data, table, units if u is not None else pd.DataFrame(),
                          res=res, measures_gdf=measures_gdf, include_total=True)
        if not sm.empty:
            for _, sr in sm[sm["ООЗТ"] != "ИТОГО"].iterrows():
                oo = sr["ООЗТ"]
                area_all = float(sr["Площадь ООЗТ, га"])
                n_kv = int(sr["Кварталов всего, шт"])
                n_kv_m = int(sr["Кварталов с мероприятиями, шт"])
                n_asg = int(sr["Назначений (квартал × мероприятие), шт"])
                cov_ha = float(sr["Площадь с мероприятиями (без наложений), га"])
                pct_cov = sr["Покрытие, % от площади ООЗТ"]
                txt = (f"{oo}: {L['groups_n'].lower()} — {n_kv}, площадь ООЗТ — {area_all:.1f} га; "
                       f"с мероприятиями — {n_kv_m} из {n_kv}; назначений — {n_asg}; "
                       f"площадь проведения мероприятий (без наложений) — {cov_ha:.1f} га"
                       + (f" ({pct_cov:.1f}% площади ООЗТ)." if pct_cov is not None and not pd.isna(pct_cov) else "."))
                B.append(("p", txt))
            shdr = [
                "ООЗТ", "Площадь ООЗТ, га", "Кварталов всего, шт", "Кварталов с мероприятиями, шт",
                "Уникальных мероприятий, шт", "Назначений (квартал × мероприятие), шт",
                "Площадь с мероприятиями (без наложений), га", "Покрытие, % от площади ООЗТ",
            ]
            shdr_disp = [
                "ООЗТ", "Площадь ООЗТ, га", "Кварталов всего", "С мерами, кв.",
                "Уник. мер, шт", "Назначений, шт", "Площадь с мерами, га", "% от ООЗТ",
            ]
            B.append(("table", shdr_disp, [[s(r[c]) for c in shdr] for _, r in sm.iterrows()],
                      [5.2, 3.0, 2.5, 2.5, 2.6, 2.9, 4.3, 3.7], 8.5))
        summ = _summary_table(table, L)
        hdr = ["Тип мероприятий", "Мероприятие",
               L["units_n"] if "Выдел" in table.columns else "Назначений", L["groups_n"]]
        widths = [5.0, 15.7, 3.0, 3.0]
        if u is not None:
            ms_all = oozt_measure_summary(data, table, u, res=res, measures_gdf=measures_gdf)
            ar = (ms_all.groupby("Мероприятие")["Площадь проведения, га"].sum().round(2)
                  if not ms_all.empty else u.groupby("Мероприятие")["Площадь, га"].sum().round(2))
            summ = summ.copy()
            summ["Площадь проведения, га"] = summ["Мероприятие"].map(ar).fillna(0.0)
            hdr.append("Площадь проведения, га")
            widths = [4.4, 12.7, 2.4, 2.4, 4.8]
        B.append(("table", hdr, [[s(r[c]) for c in hdr] for _, r in summ.iterrows()], widths, 9))
    # Общая обзорная схема на каждое ООЗТ (все мероприятия) + частные схемы по каждому мероприятию
    # в проекции UTM zone 37 (WORK_EPSG) без сжатия по вертикали.
    work_by_om = _work_geoms_by_oozt_measure(res, measures_gdf)
    oo_areas = _oozt_areas(data)
    ms_lookup: dict[tuple[str, str], tuple[float, float | None]] = {}
    if not ms_all.empty:
        for _, mr in ms_all.iterrows():
            ms_lookup[(str(mr["ООЗТ"]), str(mr["Мероприятие"]))] = (
                float(mr["Площадь проведения, га"]),
                None if pd.isna(mr["% от площади ООЗТ"]) else float(mr["% от площади ООЗТ"]),
            )
    rep_m = rep / "measures"
    all_oos = list(dict.fromkeys(kv_all["OOZT"].astype(str))) if not kv_all.empty else (
        list(dict.fromkeys(data.oozt["OOZT"].astype(str))) if data.oozt is not None and not data.oozt.empty else []
    )
    PROG.phase("Обзорные и частные схемы", 62, 76)
    for oo_i, oo in enumerate(all_oos):
        PROG.frac(oo_i / max(1, len(all_oos)), "Обзорная схема: %s (%d из %d)" % (oo, oo_i + 1, len(all_oos)))
        ctx_oo = _oozt_map_ctx(data, oo) if draw else None
        png_o = rep / ("fig_oozt_" + re.sub(r"\W+", "_", str(oo)) + ".png")
        if draw:
            log.info("Обзорная схема (все мероприятия): %s", oo)
            try:
                draw_oozt(data, evidence, oo, png_o, basemap=basemap, ctx=ctx_oo)
            except Exception as e:
                log.warning("Обзорная схема %s не построена: %s", oo, e)
        if png_o.exists():
            B += [("break",), ("h", 1, f"{oo}. Обзорная схема (все мероприятия)"), ("img", png_o)]
        sub_oo = (table[table["ООЗТ"].astype(str) == str(oo)]
                  if (table is not None and not table.empty and "ООЗТ" in table.columns)
                  else pd.DataFrame())
        ord_ms = _ordered_oozt_measures(data, sub_oo)
        if ord_ms:
            rep_m.mkdir(parents=True, exist_ok=True)
            if draw:
                oo_s = re.sub(r"\W+", "_", str(oo)).strip("_")[:26] or "oozt"
                for old_p in rep_m.glob(f"fig_{oo_s}_m[0-9][0-9]_*.png"):
                    old_p.unlink(missing_ok=True)
            oo_ha_tot, n_kv_tot = oo_areas.get(str(oo), (0.0, 0))
            for m_idx, (m_name, m_code) in enumerate(ord_ms, 1):
                png_m = rep_m / _measure_fig_name(oo, m_idx, m_code, m_name)
                sub_m = sub_oo[sub_oo["Мероприятие"].astype(str) == str(m_name)]
                kvs_raw = [str(x).removesuffix(".0") for x in sub_m["Квартал"].dropna().unique()]
                kvs_sorted = [kvs_raw[i] for i in np.argsort(_nat_rank(kvs_raw))] if kvs_raw else []
                wgs_m = work_by_om.get((str(oo), str(m_name)), [])
                season_m = merge_seasons(sub_m["Сезонность"]) if "Сезонность" in sub_m.columns else ""
                if draw:
                    PROG.frac((oo_i + m_idx / (len(ord_ms) + 1)) / max(1, len(all_oos)),
                              "Частная схема %d из %d: %s — %s" % (m_idx, len(ord_ms), oo, m_name))
                    log.info("  Частная схема [%d/%d]: %s — %s", m_idx, len(ord_ms), oo, m_name)
                    try:
                        draw_oozt(data, evidence, oo, png_m, basemap=basemap, ctx=ctx_oo,
                                  measure_name=m_name, measure_code=m_code,
                                  active_kvs=set(kvs_sorted), work_geoms=wgs_m,
                                  measure_season=season_m)
                    except Exception as e:
                        log.warning("  Частная схема %s (%s) не построена: %s", oo, m_name, e)
                if png_m.exists():
                    lbl_code = f"{m_code} " if m_code else ""
                    B += [("break",),
                          ("h", 2, f"{oo}. Мероприятие: {lbl_code}{m_name}"
                                     + _season_suffix(season_m)),
                          ("img", png_m)]
                    area_m, pct_m = ms_lookup.get((str(oo), str(m_name)), (None, None))
                    if area_m is None and wgs_m:
                        ug_m = _safe_union(wgs_m)
                        if ctx_oo is not None and ctx_oo.get("geom") is not None:
                            ug_m = ug_m.intersection(ctx_oo["geom"])
                        area_m = round(float(ug_m.area / 1e4), 2) if ug_m is not None and not ug_m.is_empty else 0.0
                        pct_m = round(100.0 * area_m / oo_ha_tot, 2) if oo_ha_tot > 0 else None
                    kv_txt = (", ".join(kvs_sorted[:24])
                              + (f" и ещё {len(kvs_sorted) - 24}" if len(kvs_sorted) > 24 else ""))
                    cap = (f"{L['groups_n']} с мероприятием: {len(kvs_sorted)}"
                           + (f" из {n_kv_tot}" if n_kv_tot else "")
                           + (f" ({L['group'].lower()} {kv_txt})" if kv_txt else ""))
                    if area_m is not None:
                        cap += f"; площадь проведения работ: {area_m:.2f} га"
                        if pct_m is not None:
                            cap += f" ({pct_m:.2f}% площади ООЗТ)"
                    B.append(("p", cap + "."))
    rt = rules_table(getattr(data, "measures", None), table, cfg)
    if not rt.empty:
        hdr2_src = ["Код", "Мероприятие", "Метод выявления (из таблицы)", "Правило", "Как рассчитано скриптом",
                    "Назначено полигонов"]
        hdr2_disp = hdr2_src[:-1] + [f"Назначено {L['groups_n'].lower()}" if qmode else "Назначено полигонов"]
        unit_word = L["groups_n"].lower() if qmode else "полигонов"
        B += [("break",),
              ("h", 1, "Как рассчитаны мероприятия"),
              ("p", "Для каждого включённого мероприятия указаны метод выявления из таблицы мероприятий, "
                    "правило назначения с параметрами (из столбцов «Правило назначения», «Типы полигонов», "
                    "«Виды: …», «Расстояние, м» и др.; если правило не задано — по тексту метода) и число "
                    f"{unit_word}, которым назначено мероприятие."),
              ("table", hdr2_disp, [[s(r[c]) for c in hdr2_src] for _, r in rt.iterrows()],
               [1.3, 4.6, 5.0, 2.8, 7.6, 2.4], 8)]
    report_table = table.rename(columns=_report_rename(cfg))
    group_cols = ["Тип мероприятий", "Тип экосистемы", "Соответствие Пункту Приказа"]
    if qmode:
        cols = ["Мероприятия", "Основание"]
        widths = [8.0, TEXT_W_CM - 8.0]
    else:
        cols = [L["unit"], L["unit_name"], "Мероприятия", "Основание"]
        w0 = 1.8 if cfg.unit_mode == "eco" else 1.3
        widths = [w0, 4.6, 7.4, TEXT_W_CM - w0 - 4.6 - 7.4]
    PROG.phase("Схемы кварталов", 76, 95)
    for kv_i, (_, kr) in enumerate(kv_all.iterrows()):
        oozt, kv = kr["OOZT"], kr["KV"]
        PROG.frac(kv_i / max(1, len(kv_all)), "Схема %d из %d: %s, %s %s" % (kv_i + 1, len(kv_all), oozt,
                                                                    L["group"].lower(), kv))
        png = rep / _fig_name(oozt, kv)
        if draw:
            log.info("Схема: %s, %s %s", oozt, L["group"].lower(), kv)
            draw_kvartal(data, evidence, kv, oozt, png, basemap=basemap)
        sub = report_table[
            (report_table["ООЗТ"] == oozt) & (report_table["Квартал"].astype(str) == str(kv))
        ] if not report_table.empty else report_table
        vyd = data.oozt[(data.oozt["KV"].astype(str) == str(kv)) & (data.oozt["OOZT"] == oozt)]
        B += [("break",), ("h", 1, f"{oozt}. {L['group']} {kv}")]
        if png.exists():
            B.append(("img", png))
        if vyd.empty and getattr(data, "kvnet", None) is not None:
            kn_ = data.kvnet
            if "KVOOZT" in kn_.columns and (kn_["KVOOZT"].astype(str).str.strip() != "").any():
                krow = kn_[(kn_["KV"].astype(str) == str(kv)) & (kn_["KVOOZT"].astype(str) == str(oozt))]
            else:
                krow = kn_[kn_["KV"].astype(str) == str(kv)]
            kv_ha = float(krow.geometry.area.sum() / 1e4) if not krow.empty else 0.0
            area_txt = f"площадь квартала {kv_ha:.1f} га (полигонов экосистем нет)"
        else:
            area_txt = f"площадь {vyd['area_ha'].sum():.1f} га"
        B.append(("p", f"{L['units_n']}: {len(vyd)}; {area_txt}; "
                       f"назначено мероприятий: {len(sub)}."))
        if sub.empty:
            B.append(("note", "Мероприятия по загруженным данным не назначены."))
            continue
        B.append(("h", 2, "Мероприятия по приказу"))
        if qmode:
            grouped = sub.sort_values(group_cols, key=_nat_sort_key("Соответствие Пункту Приказа")
                                      ).groupby(group_cols, sort=False, dropna=False)
        else:
            grouped = sub.sort_values(group_cols + [L["unit"]], key=_nat_sort_key(L["unit"], "Соответствие Пункту Приказа")
                                      ).groupby(group_cols, sort=False, dropna=False)
        for (measure_type, ecosystem_type, order_clause), group in grouped:
            B += [("h", 3, f"Тип мероприятий: {s(measure_type)}"),
                  ("h", 4, f"Тип экосистемы: {s(ecosystem_type)}"),
                  ("h", 5, f"Соответствие Пункту Приказа: {s(order_clause)}")]
            if qmode:
                rows = [[s(r[c]) for c in cols] for _, r in group.iterrows()]
            else:
                rows = [[s(r[c]) for c in cols] for _, r in group.sort_values(
                    [L["unit"]], key=_nat_sort_key(L["unit"])).iterrows()]
            B.append(("table", cols, rows, widths, 9))
    # перечень мероприятий к выполнению (по выбранным в отчёт кварталам / листам / индексам)
    if not table.empty and len(kv_all):
        sel = table.merge(kv_all[["OOZT", "KV"]].rename(columns={"OOZT": "ООЗТ", "KV": "Квартал"}),
                          on=["ООЗТ", "Квартал"], how="inner") if len(kv_all) else table
        ex, how = execution_list(sel, cfg, data)
        if not ex.empty:
            hdr = list(ex.columns)
            ex = ex.rename(columns={"Число мероприятий": "Число"})
            fixed = {L["group"]: 2.4, L["unit"]: 2.2, L["unit_name"]: 5.6, "Число": 1.5, "ООЗТ": 2.6}
            wd = [fixed.get(c, 0) for c in hdr]
            rest = TEXT_W_CM - sum(wd)
            wd = [w if w else rest for w in wd]
            B += [("break",), ("h", 1, "Перечень мероприятий к выполнению"),
                  ("p", f"Одна строка — {'один квартал' if qmode else 'один полигон'}; порядок — {how}. "
                        f"{'Кварталов' if qmode else 'Полигонов'}: {len(ex)}; "
                        f"мероприятий: {int(ex['Число'].sum())}."),
                  ("table", hdr, [[s(v) for v in r] for r in ex.itertuples(index=False)], wd, 9)]
    return B


# ------------------------------------------------------------------ Markdown
def _write_md(blocks: list, path: Path):
    def esc(x):
        return str(x).replace("|", "／").replace("\n", " ")
    out = []
    for b in blocks:
        k = b[0]
        if k == "title":
            out.append(f"# {b[1]}\n")
        elif k == "h":
            out.append(f"{'#' * (b[1] + 1)} {b[2]}\n")
        elif k == "p":
            out.append(b[1] + "\n")
        elif k == "note":
            out.append(f"_{b[1]}_\n")
        elif k == "img":
            try:
                rel = b[1].resolve().relative_to(path.parent.resolve()).as_posix()
            except ValueError:
                rel = Path(os.path.relpath(b[1].resolve(), path.parent.resolve())).as_posix()
            out.append(f"![{b[1].stem}]({rel})\n")
        elif k == "table":
            hdr, rows = b[1], b[2]
            t = ["| " + " | ".join(hdr) + " |", "|" + "|".join(["---"] * len(hdr)) + "|"]
            t += ["| " + " | ".join(esc(v) for v in r) + " |" for r in rows]
            out.append("\n".join(t) + "\n")
    path.write_text("\n".join(out), encoding="utf-8")


# ------------------------------------------------------------------ DOCX (без внешних библиотек)
_W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
         'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
         'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
         'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
_BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _x(t) -> str:
    from xml.sax.saxutils import escape
    return escape(_BAD_XML.sub("", str(t)), {'"': "&quot;"})


def _png_size(path: Path) -> tuple[int, int]:
    import struct
    with open(path, "rb") as f:
        head = f.read(24)
    if head[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", head[16:24])
    from PIL import Image
    with Image.open(path) as im:
        return im.size


def _docx_styles() -> str:
    def para(sid, name, sz, bold=False, ital=False, before=0, after=120, lvl=None, color=None, keep=False):
        ppr = ("<w:keepNext/>" if keep else "") + f'<w:spacing w:before="{before}" w:after="{after}"/>'
        if lvl is not None:
            ppr += f'<w:outlineLvl w:val="{lvl}"/>'
        rpr = ("<w:b/>" if bold else "") + ("<w:i/>" if ital else "") + \
              (f'<w:color w:val="{color}"/>' if color else "") + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>'
                f'<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
                f'<w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles {_W_NS}>'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="252" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        + para("Title", "Title", 34, bold=True, after=200, color="1F3B4D")
        + para("Heading1", "heading 1", 28, bold=True, before=120, after=120, lvl=0, color="1F3B4D", keep=True)
        + para("Heading2", "heading 2", 25, bold=True, before=160, after=80, lvl=1, color="1F3B4D", keep=True)
        + para("Heading3", "heading 3", 23, bold=True, before=160, after=40, lvl=2, keep=True)
        + para("Heading4", "heading 4", 22, bold=False, ital=True, before=0, after=40, lvl=3, keep=True)
        + para("Heading5", "heading 5", 22, bold=False, ital=True, before=0, after=80, lvl=4, color="44546A", keep=True)
        + para("Note", "Note", 22, ital=True, color="5F6368")
        + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
          '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
          '<w:left w:w="85" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="85" w:type="dxa"/>'
          '</w:tblCellMar></w:tblPr></w:style>'
          '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>'
          '<w:tblPr><w:tblBorders>'
        + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="8C9AA5"/>'
                  for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
        + '</w:tblBorders></w:tblPr></w:style></w:styles>')


def _write_docx(blocks: list, path: Path, title: str):
    import zipfile, datetime
    tw = lambda cm: int(round(cm * 567))          # твипы
    emu = lambda cm: int(round(cm * 360000))
    body, rels, media = [], [], []

    def run(t, bold=False, sz=None, ital=False):
        rpr = ("<w:b/>" if bold else "") + ("<w:i/>" if ital else "") + \
              (f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>' if sz else "")
        parts = str(t).split("\n")
        inner = "<w:br/>".join(f'<w:t xml:space="preserve">{_x(p)}</w:t>' for p in parts)
        return f"<w:r>{'<w:rPr>' + rpr + '</w:rPr>' if rpr else ''}{inner}</w:r>"

    def p(t, style=None, extra_ppr=""):
        ppr = (f'<w:pStyle w:val="{style}"/>' if style else "") + extra_ppr
        return f"<w:p>{'<w:pPr>' + ppr + '</w:pPr>' if ppr else ''}{run(t)}</w:p>"

    pending_break = False
    for b in blocks:
        k = b[0]
        if k == "break":
            pending_break = True
            continue
        brk = "<w:pageBreakBefore/>" if pending_break else ""
        pending_break = False
        if k == "title":
            body.append(p(b[1], "Title"))
        elif k == "h":
            body.append(p(b[2], f"Heading{min(b[1], 5)}", brk))
        elif k == "p":
            body.append(p(b[1], None, brk))
        elif k == "note":
            body.append(p(b[1], "Note", brk))
        elif k == "img":
            src: Path = b[1]
            n = len(media) + 1
            rid = f"rIdImg{n}"
            media.append((f"media/image{n}.png", src))
            rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                        f'relationships/image" Target="media/image{n}.png"/>')
            wpx, hpx = _png_size(src)
            w_cm = TEXT_W_CM - 0.7
            h_cm = w_cm * hpx / max(wpx, 1)
            max_h = PAGE_H_CM - 2 * MARGIN_CM - 3.2
            if h_cm > max_h:
                w_cm, h_cm = w_cm * max_h / h_cm, max_h
            cx, cy = emu(w_cm), emu(h_cm)
            body.append(
                f'<w:p><w:pPr><w:keepNext/>{brk}<w:spacing w:after="60"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
                f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
                f'<wp:docPr id="{n}" name="Схема {n}"/><wp:cNvGraphicFramePr>'
                f'<a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
                f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="{_x(src.name)}"/><pic:cNvPicPr/></pic:nvPicPr>'
                f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
        elif k == "table":
            hdr, rows, widths, fpt = b[1], b[2], b[3], b[4]
            sz = int(fpt * 2)
            ws = [tw(w) for w in widths]

            def cell(t, w, head):
                shd = '<w:shd w:val="clear" w:color="auto" w:fill="E3EAEE"/>' if head else ""
                return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}</w:tcPr>'
                        f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                        f'{run(t, bold=head, sz=sz)}</w:p></w:tc>')
            t = [f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="{sum(ws)}" w:type="dxa"/>'
                 f'<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
                 f'w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>'
                 + "".join(f'<w:gridCol w:w="{w}"/>' for w in ws) + "</w:tblGrid>",
                 '<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>'
                 + "".join(cell(h, w, True) for h, w in zip(hdr, ws)) + "</w:tr>"]
            for r in rows:
                t.append("<w:tr><w:trPr><w:cantSplit/></w:trPr>"
                         + "".join(cell(v, w, False) for v, w in zip(r, ws)) + "</w:tr>")
            t.append("</w:tbl>")
            if brk:
                body.append(f"<w:p><w:pPr>{brk}<w:spacing w:after='0'/></w:pPr></w:p>".replace("'", '"'))
            body.append("".join(t))
            body.append('<w:p><w:pPr><w:spacing w:after="60"/></w:pPr></w:p>')
    m = tw(MARGIN_CM)
    sect = (f'<w:sectPr><w:footerReference w:type="default" r:id="rIdFtr1"/>'
            f'<w:pgSz w:w="{tw(PAGE_W_CM)}" w:h="{tw(PAGE_H_CM)}" w:orient="landscape"/>'
            f'<w:pgMar w:top="{m}" w:right="{m}" w:bottom="{m}" w:left="{m}" w:header="425" w:footer="425" '
            f'w:gutter="0"/></w:sectPr>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document {_W_NS}><w:body>{"".join(body)}{sect}</w:body></w:document>')
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<w:ftr {_W_NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
              '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
              '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
              '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
              '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r>'
              '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')
    rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/styles" Target="styles.xml"/>'
                '<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/footer" Target="footer1.xml"/>' + "".join(rels) + "</Relationships>")
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{_x(title)}</dc:title><dc:creator>oozt_recs.py</dc:creator><dc:language>ru-RU</dc:language>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
    ctypes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              '<Default Extension="png" ContentType="image/png"/>'
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
              'wordprocessingml.document.main+xml"/>'
              '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
              'wordprocessingml.styles+xml"/>'
              '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.'
              'wordprocessingml.footer+xml"/>'
              '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.'
              'core-properties+xml"/></Types>')
    pkg_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/officeDocument" Target="word/document.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/'
                'metadata/core-properties" Target="docProps/core.xml"/></Relationships>')
    tmp = path.with_name(path.stem + ".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ctypes)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", _docx_styles())
        z.writestr("word/footer1.xml", footer)
        z.writestr("word/_rels/document.xml.rels", rels_xml)
        for name, src in media:
            z.write(src, "word/" + name, compress_type=zipfile.ZIP_STORED)
    os.replace(tmp, path)


# ------------------------------------------------------------------ PDF / DOC через Word или LibreOffice
def _convert_with_word(docx: Path, targets: dict) -> list[str]:
    """targets: {"pdf": Path, "doc": Path}. Возвращает список созданных форматов."""
    import win32com.client  # noqa: F401  (pywin32; есть в conda-окружениях Windows)
    import pythoncom
    fmt = {"pdf": 17, "doc": 0}           # wdFormatPDF, wdFormatDocument (Word 97–2003)
    done = []
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        d = word.Documents.Open(str(docx.resolve()), ConfirmConversions=False, ReadOnly=True,
                                AddToRecentFiles=False, Visible=False)
        try:
            d.Fields.Update()
            for k, p in targets.items():
                d.SaveAs2(str(p.resolve()), FileFormat=fmt[k], AddToRecentFiles=False)
                done.append(k)
        finally:
            d.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
    return done


def _find_soffice() -> str | None:
    import shutil
    for n in ("soffice", "libreoffice"):
        p = shutil.which(n)
        if p:
            return p
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        p = Path(base) / "LibreOffice" / "program" / "soffice.exe"
        if p.exists():
            return str(p)
    return None


def _convert_with_soffice(soffice: str, docx: Path, targets: dict) -> list[str]:
    import subprocess, tempfile
    done = []
    for k, p in targets.items():
        with tempfile.TemporaryDirectory() as td:
            subprocess.run([soffice, "--headless", f"-env:UserInstallation=file:///{Path(td).as_posix()}/u",
                            "--convert-to", k, "--outdir", td, str(docx)],
                           check=True, timeout=900, capture_output=True)
            res = Path(td) / (docx.stem + "." + k)
            if not res.exists():
                raise RuntimeError(f"LibreOffice не создал {res.name}")
            shutil_move(res, p)
            done.append(k)
    return done


def shutil_move(src: Path, dst: Path):
    import shutil
    if dst.exists():
        dst.unlink()
    shutil.move(str(src), str(dst))


# ------------------------------------------------------------------ PDF собственной вёрсткой (matplotlib)
def _write_pdf_mpl(blocks: list, path: Path, title: str):
    """Запасной путь, если нет Word и LibreOffice: та же структура, страницы А4 альбомные."""
    enable_win_dll_path()
    import matplotlib.image as mpimg
    from matplotlib.backends.backend_pdf import PdfPages
    from matplotlib.patches import Rectangle
    PT = 72 / 2.54
    W, H, M = PAGE_W_CM * PT, PAGE_H_CM * PT, MARGIN_CM * PT
    TW = W - 2 * M
    from matplotlib.font_manager import FontProperties, findfont, get_font
    fonts = {b: get_font(findfont(FontProperties(family="DejaVu Sans", weight="bold" if b else "normal")))
             for b in (False, True)}
    _wcache = {}

    def measure(t, size, bold):
        key = (t, size, bold)
        if key not in _wcache:
            f = fonts[bold]
            f.set_size(size, 72)
            f.set_text(t, 0.0)
            _wcache[key] = f.get_width_height()[0] / 64.0 * 1.02
        return _wcache[key]
    st = {"title": (17, True), 1: (14, True), 2: (12.5, True), 3: (11.5, True), 4: (11, False), 5: (11, False),
          "p": (10.5, False), "note": (10.5, False)}
    matplotlib.rcParams["pdf.fonttype"] = 42          # шрифты TrueType: текст копируется и ищется
    pdf = PdfPages(path, metadata={"Title": title, "Creator": "oozt_recs.py"})
    state = {"fig": None, "y": 0.0, "page": 0}

    def wrap(t, size, bold, width_pt):
        out = []
        for para in str(t).split("\n"):
            words = re.split(r"(?<=-)|\s+", para.strip())
            line = ""
            for w in [w for w in words if w]:
                cand = w if not line else (line + w if line.endswith("-") else line + " " + w)
                if measure(cand, size, bold) <= width_pt:
                    line = cand
                    continue
                if line:
                    out.append(line)
                while measure(w, size, bold) > width_pt and len(w) > 1:   # слово длиннее колонки
                    k = len(w) - 1
                    while k > 1 and measure(w[:k] + "-", size, bold) > width_pt:
                        k -= 1
                    out.append(w[:k] + "-")
                    w = w[k:]
                line = w
            out.append(line)
        return out

    def new_page():
        if state["fig"] is not None:
            finish()
        fig = plt.figure(figsize=(W / 72, H / 72))
        state.update(fig=fig, y=H - M, page=state["page"] + 1)

    def finish():
        fig = state["fig"]
        fig.text(0.5, (M * 0.45) / H, str(state["page"]), ha="center", va="center", fontsize=8, color="#5F6368")
        pdf.savefig(fig)
        plt.close(fig)
        state["fig"] = None

    def txt(x, y, t, size, bold=False, color="#202124", style="normal"):
        state["fig"].text(x / W, y / H, t, fontsize=size, fontweight="bold" if bold else "normal",
                          color=color, va="top", ha="left", fontstyle=style)

    def ensure(h):
        if state["fig"] is None or state["y"] - h < M:
            new_page()

    def para(t, size, bold, color="#202124", style="normal", before=0, after=4, keep=0):
        lines = wrap(t, size, bold, TW)
        lh = size * 1.3
        ensure(before + lh * min(len(lines), 2) + keep)
        state["y"] -= before
        for ln in lines:
            ensure(lh)
            txt(M, state["y"], ln, size, bold, color, style)
            state["y"] -= lh
        state["y"] -= after

    def table(hdr, rows, widths, fpt):
        ws = [w / sum(widths) * TW for w in widths]
        lh, pad = fpt * 1.25, 2.5

        def row_lines(r, bold):
            return [wrap(v, fpt, bold, w - 2 * pad - 1) for v, w in zip(r, ws)]

        def draw_row(cells, bold, fill=None):
            n = max(len(c) for c in cells)
            h = n * lh + 2 * pad
            x, y = M, state["y"]
            for c, w in zip(cells, ws):
                state["fig"].add_artist(Rectangle((x / W, (y - h) / H), w / W, h / H, transform=state["fig"].transFigure,
                                                  facecolor=fill or "none", edgecolor="#8C9AA5", linewidth=0.5))
                for i, ln in enumerate(c):
                    txt(x + pad, y - pad - i * lh, ln, fpt, bold)
                x += w
            state["y"] -= h

        head = row_lines(hdr, True)
        hh = max(len(c) for c in head) * lh + 2 * pad
        ensure(hh + 3 * lh)
        draw_row(head, True, "#E3EAEE")
        for r in rows:
            cells = row_lines(r, False)
            while cells:
                h = max(len(c) for c in cells) * lh + 2 * pad
                if state["y"] - h < M:
                    room = int((state["y"] - M - 2 * pad) / lh)
                    if room >= 3:                       # длинную строку делим между страницами
                        draw_row([c[:room] for c in cells], False)
                        cells = [c[room:] for c in cells]
                        if not any(cells):
                            break
                    new_page()
                    draw_row(head, True, "#E3EAEE")
                    continue
                draw_row(cells, False)
                break
        state["y"] -= 6

    def img_size(path):
        img = mpimg.imread(str(path))
        ih, iw = img.shape[0], img.shape[1]
        w = TW - 20
        h = w * ih / iw
        max_h = H - 2 * M - 90
        if h > max_h:
            w, h = w * max_h / h, max_h
        return img, w, h

    pending_break = False
    for i, b in enumerate(blocks):
        k = b[0]
        if k == "break":
            pending_break = True
            continue
        if pending_break:
            new_page()
            pending_break = False
        if k == "title":
            para(b[1], *st["title"], color="#1F3B4D", after=10)
        elif k == "h":
            size, bold = st[min(b[1], 5)]
            keep = {1: 60, 2: 140, 3: 110, 4: 85, 5: 60}[min(b[1], 5)]
            if i + 1 < len(blocks) and blocks[i + 1][0] == "img":      # заголовок не отрываем от схемы
                keep = img_size(blocks[i + 1][1])[2] + 8
            para(b[2], size, bold, color="#1F3B4D" if b[1] <= 2 else "#202124",
                 style="italic" if b[1] >= 4 else "normal", before=8 if b[1] <= 3 else 0, after=3, keep=keep)
        elif k == "p":
            para(b[1], *st["p"])
        elif k == "note":
            para(b[1], *st["note"], color="#5F6368", style="italic")
        elif k == "img":
            img, w, h = img_size(b[1])
            ensure(h + 6)
            x0 = M + (TW - w) / 2
            ax = state["fig"].add_axes([x0 / W, (state["y"] - h) / H, w / W, h / H])
            ax.imshow(img, interpolation="lanczos")
            ax.set_axis_off()
            state["y"] -= h + 6
        elif k == "table":
            table(b[1], b[2], b[3], b[4])
    if state["fig"] is None:
        new_page()
    finish()
    pdf.close()


def _write_report_files(blocks: list, rep: Path, title: str, formats) -> dict:
    """Записывает report.md и выбранные форматы; возвращает {формат: путь}."""
    formats = [f for f in (formats or []) if f in REPORT_FORMATS]
    made = {}
    PROG.phase("Запись отчёта (MD" + "".join(", " + f.upper() for f in formats) + ")", 95, 99.5)
    _write_md(blocks, rep / "report.md")
    made["md"] = rep / "report.md"
    if not formats:
        return made
    docx = rep / "report.docx"
    try:
        _write_docx(blocks, docx, title)
        if "docx" in formats:
            made["docx"] = docx
            log.info("Отчёт DOCX: %s", docx)
    except PermissionError:
        log.warning("report.docx открыт в другой программе — закройте его и пересоберите отчёт (--report-only)")
        return made
    targets = {k: rep / f"report.{k}" for k in ("pdf", "doc") if k in formats}
    methods = []
    if os.name == "nt":
        methods.append(("Microsoft Word", lambda t: _convert_with_word(docx, t)))
    so = _find_soffice()
    if so:
        methods.append(("LibreOffice", lambda t: _convert_with_soffice(so, docx, t)))
    for name, fn in methods:
        if not targets:
            break
        try:
            for k in fn(dict(targets)):
                made[k] = targets.pop(k)
                log.info("Отчёт %s (%s): %s", k.upper(), name, made[k])
        except ImportError:
            log.info("%s: нет модуля pywin32 — пропускаю", name)
        except Exception as e:
            log.info("%s недоступен для конвертации (%s)", name, str(e).splitlines()[0][:200] if str(e) else type(e).__name__)
    if "pdf" in targets:
        try:
            _write_pdf_mpl(blocks, targets["pdf"], title)
            made["pdf"] = targets.pop("pdf")
            log.info("Отчёт PDF (встроенная вёрстка): %s", made["pdf"])
        except PermissionError:
            log.warning("report.pdf открыт в другой программе — закройте его и пересоберите отчёт (--report-only)")
        except Exception as e:
            log.warning("PDF не создан: %s", e)
    if "doc" in targets:
        log.warning("Формат DOC (Word 97–2003) требует установленного Microsoft Word или LibreOffice — не создан")
    if "docx" not in formats and docx.exists() and "docx" not in made:
        try:
            docx.unlink()                       # промежуточный файл, если DOCX не заказан
        except OSError:
            pass
    return made


def build_report(data: Data, table: pd.DataFrame, evidence, out_dir: Path, kv_list, basemap=True,
                 formats=("docx", "pdf"), units: pd.DataFrame | None = None,
                 res: "Results | None" = None,
                 measures_gdf: gpd.GeoDataFrame | None = None):
    enable_win_dll_path()
    cfg = data.cfg
    rep = out_dir / "report"
    rep.mkdir(parents=True, exist_ok=True)
    kv_all = _report_groups_table(data, kv_list)
    blocks = _report_blocks(data, table, evidence, rep, kv_all, basemap=basemap,
                            units=units, res=res, measures_gdf=measures_gdf)
    title = blocks[0][1]
    made = _write_report_files(blocks, rep, title, formats)
    return made.get("pdf") or made.get("docx") or made["md"]


# =====================================================================================
#                                 ПРОЕКТ QGIS
# =====================================================================================
QGIS_PROJECT_NAME = "oozt_measures.qgs"
_MEASURE_COLORS = ["#E6194B", "#3CB44B", "#4363D8", "#F58231", "#911EB4", "#42D4F4", "#F032E6", "#BFEF45",
                   "#469990", "#9A6324", "#800000", "#808000", "#000075", "#DCBEFF", "#FFE119", "#AAFFC3",
                   "#FABED4", "#A9A9A9"]


def _rgba(hex_color: str, alpha: float = 1.0) -> str:
    h = hex_color.lstrip("#")
    return f"{int(h[0:2], 16)},{int(h[2:4], 16)},{int(h[4:6], 16)},{int(round(alpha * 255))}"


def _q_opts(d: dict) -> str:
    """Свойства слоя символа: формат <Option> (QGIS ≥ 3.26) и <prop> (более ранние версии)."""
    opt = "".join(f'<Option name="{_x(k)}" type="QString" value="{_x(v)}"/>' for k, v in d.items())
    prop = "".join(f'<prop k="{_x(k)}" v="{_x(v)}"/>' for k, v in d.items())
    return f'<Option type="Map">{opt}</Option>{prop}'


def _q_symbol(kind: str, name: str, color: str, alpha=1.0, outline="#333333", width=None, size=None,
              outline_width=0.26) -> str:
    if kind == "fill":
        lay = ("SimpleFill", {"color": _rgba(color, alpha), "outline_color": _rgba(outline, 1.0),
                              "outline_style": "solid", "outline_width": str(outline_width),
                              "outline_width_unit": "MM", "style": "solid" if alpha > 0 else "no",
                              "joinstyle": "bevel"})
    elif kind == "line":
        lay = ("SimpleLine", {"line_color": _rgba(color, alpha), "line_width": str(width or 0.6),
                              "line_width_unit": "MM", "line_style": "solid", "capstyle": "round",
                              "joinstyle": "round"})
    else:
        lay = ("SimpleMarker", {"name": "circle", "color": _rgba(color, alpha), "outline_color": "255,255,255,255",
                                "outline_style": "solid", "outline_width": "0.3", "outline_width_unit": "MM",
                                "size": str(size or 2.6), "size_unit": "MM"})
    return (f'<symbol type="{kind}" name="{_x(name)}" alpha="1" clip_to_extent="1" force_rhr="0">'
            f'<layer class="{lay[0]}" enabled="1" locked="0" pass="0">{_q_opts(lay[1])}</layer></symbol>')


def _q_single(sym: str) -> str:
    return (f'<renderer-v2 type="singleSymbol" forceraster="0" symbollevels="0" enableorderby="0">'
            f'<symbols>{sym}</symbols></renderer-v2>')


def _q_categorized(attr: str, cats: list) -> str:
    """cats: [(значение, подпись, символ_xml_с_именем_i)]"""
    c = "".join(f'<category render="true" symbol="{i}" value="{_x(v)}" label="{_x(lbl)}"/>'
                for i, (v, lbl, _) in enumerate(cats))
    s_ = "".join(sym for _, _, sym in cats)
    return (f'<renderer-v2 type="categorizedSymbol" attr="{_x(attr)}" forceraster="0" symbollevels="0" '
            f'enableorderby="0"><categories>{c}</categories><symbols>{s_}</symbols></renderer-v2>')


def _q_labels(field_name: str, max_scale=20000, size=8) -> str:
    return (f'<labeling type="simple"><settings calloutType="simple">'
            f'<text-style fieldName="{_x(field_name)}" isExpression="0" fontFamily="Arial" fontSize="{size}" '
            f'fontSizeUnit="Point" fontWeight="75" fontItalic="0" textColor="20,20,20,255" textOpacity="1" '
            f'multilineHeight="1" blendMode="0" previewBkgrdColor="255,255,255,255">'
            f'<text-buffer bufferDraw="1" bufferSize="0.8" bufferSizeUnits="MM" bufferColor="255,255,255,255" '
            f'bufferOpacity="1" bufferJoinStyle="128" bufferNoFill="1" bufferBlendMode="0"/>'
            f'</text-style><text-format wrapChar="" multilineAlign="1" plussign="0" decimals="3" formatNumbers="0"/>'
            f'<placement placement="1" centroidInside="1" centroidWhole="0" fitInPolygonOnly="0" priority="5" '
            f'offsetType="0" dist="0" overrunDistance="0"/>'
            f'<rendering scaleVisibility="1" scaleMin="1" scaleMax="{max_scale}" obstacle="1" '
            f'displayAll="0" upsidedownLabels="0" labelPerPart="0" mergeLines="0" minFeatureSize="0" '
            f'limitNumLabels="0" maxNumLabels="2000" drawLabels="1" fontLimitPixelSize="0"/>'
            f'</settings></labeling>')


def _q_srs(epsg: int) -> str:
    from pyproj import CRS
    c = CRS.from_epsg(epsg)
    return (f'<spatialrefsys nativeFormat="Wkt"><wkt>{_x(c.to_wkt())}</wkt><proj4>{_x(c.to_proj4())}</proj4>'
            f'<srid>{epsg}</srid><authid>EPSG:{epsg}</authid><description>{_x(c.name)}</description>'
            f'<projectionacronym></projectionacronym><ellipsoidacronym></ellipsoidacronym>'
            f'<geographicflag>{"true" if c.is_geographic else "false"}</geographicflag></spatialrefsys>')


def write_qgis_project(out: Path, cfg: Config, data: "Data", gpkg: Path, terr_layer: str, label_field: str,
                       measures_layer: str | None, measure_types: list, evidence_layers: list, nmax: int = 0,
                       zones_layer: str | None = None, zone_measures=(), ctx_layers=(), heatmap_tif=None) -> Path:
    """Файл проекта QGIS (.qgs, относительные пути) с результатами из evidence.gpkg.
    measure_types: [(тип мероприятий, [мероприятие, …])]; evidence_layers: [(имя слоя gpkg, геометрия, [значения layer])]"""
    import uuid
    L = unit_labels(cfg)
    rel = "./" + gpkg.name
    layers_xml, tree = [], []
    wgs = _q_srs(4326)
    b = gpd.GeoSeries([data.union], crs=WORK_EPSG).to_crs(3857).total_bounds
    pad = max(b[2] - b[0], b[3] - b[1]) * 0.05

    def add_vec(name, layer, geom, renderer, subset=None, labels=None, visible=True):
        lid = f"{re.sub(r'[^0-9A-Za-z]+', '_', layer)}_{uuid.uuid4().hex[:10]}"
        src = f"{rel}|layername={layer}" + (f"|subset={subset}" if subset else "")
        gtype = {"poly": "Polygon", "line": "Line", "point": "Point"}[geom]
        lab_attr = ' labelsEnabled="1"' if labels else ""
        layers_xml.append(
            f'<maplayer type="vector" geometry="{gtype}" autoRefreshEnabled="0" refreshOnNotifyEnabled="0" '
            f'hasScaleBasedVisibilityFlag="0" readOnly="0" simplifyDrawingHints="1"{lab_attr}>'
            f'<id>{lid}</id><datasource>{_x(src)}</datasource><layername>{_x(name)}</layername>'
            f'<srs>{wgs}</srs><provider encoding="UTF-8">ogr</provider>{renderer}{labels or ""}'
            f'<blendMode>0</blendMode><featureBlendMode>0</featureBlendMode><layerOpacity>1</layerOpacity>'
            f'</maplayer>')
        return (f'<layer-tree-layer id="{lid}" name="{_x(name)}" checked="{"Qt::Checked" if visible else "Qt::Unchecked"}" '
                f'expanded="0" providerKey="ogr" source="{_x(src)}"><customproperties/></layer-tree-layer>')

    def group(name, children, checked=True, expanded=True):
        return (f'<layer-tree-group name="{_x(name)}" checked="{"Qt::Checked" if checked else "Qt::Unchecked"}" '
                f'expanded="{1 if expanded else 0}"><customproperties/>{"".join(children)}</layer-tree-group>')

    # 1) основания — цвета как на схемах
    ev_nodes = []
    for lyr, geom, values in sorted(evidence_layers, key=lambda t: {"point": 0, "line": 1, "poly": 2}[t[1]]):
        cats = []
        for i, v in enumerate(sorted(values)):
            color, kind = _slope_style(v)
            if geom == "point":
                sym = _q_symbol("marker", str(i), color, size=3.0 if "виды" in v else 2.2)
            elif geom == "line":
                sym = _q_symbol("line", str(i), color, width=0.9)
            else:
                alpha = 0.0 if "SDM" in v else (0.85 if kind == "slope" else 0.45)
                sym = _q_symbol("fill", str(i), color, alpha=alpha, outline=color, outline_width=0.3)
            cats.append((v, v, sym))
        title = {"point": "Основания — точки", "line": "Основания — линии", "poly": "Основания — участки"}[geom]
        ev_nodes.append(add_vec(title, lyr, geom, _q_categorized("layer", cats)))
    # 2) мероприятия: по типам, слой на каждое мероприятие (в measures лежат только
    #    географически очерченные полигоны проведения работ, без целых кварталов)
    m_nodes = []
    if measures_layer:
        ci = 0
        for mtype, names in measure_types:
            kids = []
            for nm in names:
                col = _MEASURE_COLORS[ci % len(_MEASURE_COLORS)]
                ci += 1
                subset = '"Мероприятие" = ' + "'" + str(nm).replace("'", "''") + "'"
                kids.append(add_vec(nm, measures_layer, "poly",
                                    _q_single(_q_symbol("fill", "0", col, alpha=0.58, outline=col, outline_width=0.45)),
                                    subset=subset, visible=False))
            m_nodes.append(group(mtype or "Без типа", kids, checked=True, expanded=False))
    # 3) кварталы (верхний слой: контуры и номера по порядку слоёв п. 9) + выключенная картограмма числа мер
    terr_outline_title = "Кварталы (границы и номера)" if label_field == "KV" \
        else f"{L['units_gen'].capitalize()} (границы и номера)"
    terr_outline = add_vec(
        terr_outline_title, terr_layer, "poly",
        _q_single(_q_symbol("fill", "0", "#FFFFFF", alpha=0.0, outline="#202020", outline_width=0.48)),
        labels=_q_labels(label_field), visible=True,
    )
    reds = ["#FFFFFF", "#FEE8C8", "#FDBB84", "#FC8D59", "#E34A33", "#B30000", "#7F0000"]
    cats = [(str(k), (f"{k} мероприят." if k < 6 or nmax <= 6 else f"{k}"), _q_symbol(
        "fill", str(k), reds[min(k, len(reds) - 1)], alpha=0.0 if k == 0 else 0.72, outline="#555555",
        outline_width=0.2)) for k in range(0, nmax + 1)]
    terr_title = "Кварталы: число мероприятий" if label_field == "KV" \
        else f"{L['units_gen'].capitalize()}: число мероприятий"
    terr_choro = add_vec(terr_title, terr_layer, "poly",
                         _q_categorized("n_measures", cats), labels=_q_labels(label_field), visible=False)
    # 3а) контекстные слои подложки (в границах ООЗТ): дороги, гидрография,
    #     горизонтали, рекреация, стихийные тропы, все точки ПВН — из evidence.gpkg
    _CTX_TITLE = {"rivers_ditches": "Реки и канавы", "lakes_ponds": "Озёра и пруды",
                  "roads_paved": "Дороги с твёрдым покрытием", "roads_dirt": "Дороги грунтовые",
                  "contours": "Горизонтали (2,5 м)", "recreation_points": "Рекреация: объекты",
                  "informal_trails": _inf_label()[:1].upper() + _inf_label()[1:], "pvn_points": "Лесопатология: все точки ПВН",
                  "wetlands": "Болота"}
    ctx_by_name = {}
    for c in ctx_layers:
        nm = c["name"]
        if c["geom"] == "line":
            sym = _q_symbol("line", "0", c["color"], width=c.get("width", 0.8))
            rend = _q_single(sym)
        elif c["geom"] == "poly":
            sym = _q_symbol("fill", "0", c["color"], alpha=c.get("alpha", 0.8),
                            outline=c["color"], outline_width=0.2)
            rend = _q_single(sym)
        else:
            rend = _q_single(_q_symbol("marker", "0", c["color"], size=2.5))
            if c.get("cat_field") and c.get("cat_values"):
                # точки рекреации по категориям — как на схемах
                from matplotlib.colors import to_hex as _tohex
                cmap = plt.get_cmap("tab10")
                pcats = [(str(v), str(v), _q_symbol("marker", str(i), _tohex(cmap(i % 10)), size=2.5))
                         for i, v in enumerate(c["cat_values"])]
                if pcats:
                    rend = _q_categorized(c["cat_field"], pcats)
        ctx_by_name[nm] = add_vec(_CTX_TITLE.get(nm, nm), nm, c["geom"], rend)
    # 4) исходные векторные слои (выключены)
    src_nodes = []
    for title, key, geom, color in (("Тропы", "trails", "line", "#707070"),
                                    ("Гидрография — линии", "hydro_lines", "line", "#1F78B4"),
                                    ("Гидрография — полигоны", "hydro_polys", "poly", "#A6CEE3"),
                                    ("Болота (вектор)", "wetlands", "poly", WETLAND_COLOR)):
        for pth in getattr(cfg, key) or []:
            pth = Path(pth)
            try:
                rp = os.path.relpath(pth.resolve(), out.resolve())
                ds = "./" + Path(rp).as_posix() if not os.path.isabs(rp) else pth.as_posix()
            except ValueError:                  # другой диск — абсолютный путь
                ds = pth.as_posix()
            lid = f"src_{key}_{uuid.uuid4().hex[:10]}"
            sym = _q_symbol("line", "0", color, width=0.3) if geom == "line" else \
                _q_symbol("fill", "0", color, alpha=0.6, outline="#1F78B4", outline_width=0.2)
            layers_xml.append(
                f'<maplayer type="vector" geometry="{"Line" if geom == "line" else "Polygon"}">'
                f'<id>{lid}</id><datasource>{_x(ds)}</datasource><layername>{_x(title + " — " + pth.stem)}</layername>'
                f'<provider encoding="UTF-8">ogr</provider>{_q_single(sym)}</maplayer>')
            src_nodes.append(f'<layer-tree-layer id="{lid}" name="{_x(title + " — " + pth.stem)}" '
                             f'checked="Qt::Unchecked" expanded="0" providerKey="ogr" source="{_x(ds)}">'
                             f'<customproperties/></layer-tree-layer>')
    # 5) подложка OSM
    osm_id = f"osm_{uuid.uuid4().hex[:10]}"
    osm_src = "type=xyz&url=https://tile.openstreetmap.org/%7Bz%7D/%7Bx%7D/%7By%7D.png&zmax=19&zmin=0"
    layers_xml.append(
        f'<maplayer type="raster" autoRefreshEnabled="0"><id>{osm_id}</id><datasource>{_x(osm_src)}</datasource>'
        f'<layername>OpenStreetMap</layername><srs>{_q_srs(3857)}</srs><provider>wms</provider>'
        f'<pipe><rasterrenderer type="singlebandcolordata" opacity="0.8" alphaBand="-1" band="1">'
        f'<rasterTransparency/></rasterrenderer><brightnesscontrast brightness="0" contrast="0"/>'
        f'<huesaturation saturation="-70" grayscaleMode="0" colorizeOn="0" colorizeRed="255" colorizeGreen="128" '
        f'colorizeBlue="128" colorizeStrength="100"/><rasterresampler maxOversampling="2"/></pipe>'
        f'<blendMode>0</blendMode></maplayer>')
    osm_node = (f'<layer-tree-layer id="{osm_id}" name="OpenStreetMap" checked="Qt::Checked" expanded="0" '
                f'providerKey="wms" source="{_x(osm_src)}"><customproperties/></layer-tree-layer>')
    # тепловая карта рекреации (растр, полупрозрачно)
    heat_node = None
    if heatmap_tif is not None and Path(heatmap_tif).exists():
        hid = f"heat_{uuid.uuid4().hex[:10]}"
        hds = "./" + Path(heatmap_tif).name
        layers_xml.append(
            f'<maplayer type="raster" autoRefreshEnabled="0"><id>{hid}</id><datasource>{_x(hds)}</datasource>'
            f'<layername>Рекреация: тепловая карта</layername><srs>{_q_srs(WORK_EPSG)}</srs>'
            f'<provider>gdal</provider>'
            f'<pipe><rasterrenderer type="singlebandgray" opacity="0.55" grayBand="1" gradient="BlackToWhite">'
            f'<minMaxOrigin><limits>CumulativeCut</limits><extent>WholeRaster</extent>'
            f'<statPrecision>Estimated</statPrecision></minMaxOrigin></rasterrenderer>'
            f'<brightnesscontrast brightness="0" contrast="0"/>'
            f'<huesaturation saturation="0" grayscaleMode="0" colorizeOn="0"/>'
            f'<rasterresampler maxOversampling="2"/></pipe>'
            f'<blendMode>0</blendMode></maplayer>')
        heat_node = (f'<layer-tree-layer id="{hid}" name="Рекреация: тепловая карта" checked="Qt::Checked" '
                     f'expanded="0" providerKey="gdal" source="{_x(hds)}"><customproperties/></layer-tree-layer>')

    tree = []
    tree.append(terr_outline)
    tree.append(terr_choro)
    if "pvn_points" in ctx_by_name:
        tree.append(ctx_by_name["pvn_points"])
    if m_nodes:
        kids = list(m_nodes)
        if ev_nodes:
            kids.append(group("Основания назначения", ev_nodes))
        tree.append(group("Мероприятия (включайте нужные)", kids, expanded=True))
    elif ev_nodes:
        tree.append(group("Основания назначения", ev_nodes))
    # порядок подложки (сверху вниз): дороги, горизонтали, гидрография, рекреация, тропы
    for nm in ("roads_paved", "roads_dirt", "contours", "rivers_ditches", "lakes_ponds", "wetlands"):
        if nm in ctx_by_name:
            tree.append(ctx_by_name[nm])
    if heat_node:
        tree.append(heat_node)
    if "recreation_points" in ctx_by_name:
        tree.append(ctx_by_name["recreation_points"])
    if "informal_trails" in ctx_by_name:
        tree.append(ctx_by_name["informal_trails"])
    if src_nodes:
        tree.append(group("Исходные слои", src_nodes, checked=False, expanded=False))
    tree.append(osm_node)
    title = f"Природоохранные мероприятия: {', '.join(map(str, data.oozt['OOZT'].unique()[:3]))}"
    xml = ("<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>\n"
           f'<qgis projectname="{_x(title)}" version="3.28.0-Firenze">'
           f'<homePath path=""/><title>{_x(title)}</title>'
           f'<projectCrs>{_q_srs(3857)}</projectCrs>'
           f'<layer-tree-group><customproperties/>{"".join(tree)}</layer-tree-group>'
           f'<mapcanvas annotationsVisible="1" name="theMapCanvas"><units>meters</units>'
           f'<extent><xmin>{b[0] - pad}</xmin><ymin>{b[1] - pad}</ymin><xmax>{b[2] + pad}</xmax>'
           f'<ymax>{b[3] + pad}</ymax></extent><rotation>0</rotation>'
           f'<destinationsrs>{_q_srs(3857)}</destinationsrs><rendermaptile>0</rendermaptile></mapcanvas>'
           f'<projectlayers>{"".join(layers_xml)}</projectlayers>'
           f'<properties><Paths><Absolute type="bool">false</Absolute></Paths>'
           f'<Measurement><DistanceUnits type="QString">meters</DistanceUnits>'
           f'<AreaUnits type="QString">ha</AreaUnits></Measurement></properties>'
           f'</qgis>')
    path = out / QGIS_PROJECT_NAME
    path.write_text(xml, encoding="utf-8")
    return path


# =====================================================================================
#                                       MAIN
# =====================================================================================
def _setup_logging(out: Path, name="run.log"):
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s  %(message)s", datefmt="%H:%M:%S",
                        handlers=[logging.StreamHandler(sys.stdout),
                                  logging.FileHandler(out / name, encoding="utf-8", mode="w")])


def _prep_gpkg_gdf(gdf: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Убирает регистронезависимые дубли колонок и служебное поле fid перед записью в GeoPackage."""
    if gdf is None:
        return gdf
    seen, keep, rename = set(), [], {}
    for c in gdf.columns:
        if c == "geometry":
            keep.append(c)
            continue
        cl = str(c).lower()
        if cl == "fid":
            if "src_fid" not in seen:
                rename[c] = "src_fid"
                seen.add("src_fid")
                keep.append(c)
            continue
        if cl in seen:
            continue
        seen.add(cl)
        keep.append(c)
    out_g = gdf[keep].copy()
    return out_g.rename(columns=rename) if rename else out_g


def report_only(cfg: Config):
    """Перестроить отчёт по уже рассчитанным measures_long.csv и evidence.gpkg (без пересчёта правил)."""
    out = Path(cfg.out_dir)
    _setup_logging(out, "report.log")
    data = Data(cfg, oozt_only=True)
    try:
        data.measures = read_measures(cfg.measures_xlsx)
    except Exception as e:
        log.warning("Таблица мероприятий не прочитана (%s) — раздел «Как рассчитаны мероприятия» пропущен", e)
    # Для схем загружаем контекст подложки (дороги, гидрографию, рельеф, рекреацию, ПВН)
    data.trails = data._load_trails()
    data.roads_paved = read_many(cfg.roads_paved, data.bbox_wgs84, data.clip)
    data.roads_dirt = read_many(cfg.roads_dirt, data.bbox_wgs84, data.clip)
    data.hydro_lines = read_many(cfg.hydro_lines, data.bbox_wgs84, data.clip)
    data.hydro_polys = read_many(cfg.hydro_polys, data.bbox_wgs84, data.clip)
    data.wetlands = read_many(cfg.wetlands, data.bbox_wgs84, data.clip)
    data.rec_points = read_many(cfg.recreation, data.bbox_wgs84, data.clip)
    data.lesopat = data._load_lesopat()
    if cfg.dem_path:
        try:
            arr, tr, _ = read_raster_clipped(cfg.dem_path, data.bounds_pad, resolution=CELL, what="ЦМР")
            data.dem = (arr, tr)
        except Exception as e:
            log.warning("ЦМР не загружена (%s)", e)
    disp = display_columns(cfg)
    dtype_d = {disp["Квартал"]: str}
    if disp.get("Выдел"):
        dtype_d[disp["Выдел"]] = str
    table = pd.read_csv(out / "measures_long.csv", encoding="utf-8-sig", dtype=dtype_d)
    table = table.rename(columns={v: k for k, v in disp.items()})
    for c in ("Квартал", "Выдел"):
        if c in table.columns:
            table[c] = _maybe_int(table[c])
    gpkg = out / "evidence.gpkg"
    parts = []
    measures_gdf = None
    if gpkg.exists():
        gpkg_layers = [ly for ly, _ in pyogrio.list_layers(gpkg)]
        for ly in gpkg_layers:
            if ly.startswith("evidence_"):
                g = gpd.read_file(gpkg, layer=ly)
                parts.append(g.to_crs(WORK_EPSG))
        if "measures" in gpkg_layers:
            try:
                mg = gpd.read_file(gpkg, layer="measures").to_crs(WORK_EPSG)
                measures_gdf = mg.rename(columns={v: k for k, v in disp.items()})
            except Exception as e:
                log.warning("Слой measures из evidence.gpkg не прочитан: %s", e)
        if data.trails_informal is None and "informal_trails" in gpkg_layers:
            try:
                data.trails_informal = gpd.read_file(gpkg, layer="informal_trails").to_crs(WORK_EPSG)
            except Exception:
                pass
    ev = gpd.GeoDataFrame(pd.concat(parts, ignore_index=True), crs=WORK_EPSG) if parts else None
    units = None
    up = out / "units_long.csv"
    if up.exists():
        try:
            units = pd.read_csv(up, encoding="utf-8-sig")
        except Exception as e:
            log.warning("units_long.csv не прочитан — сводка площадей в отчёте пропущена: %s", e)
    if units is not None:
        try:
            write_oozt_summary_excel(out / "oozt_summary.xlsx", data, table, units, measures_gdf=measures_gdf)
        except Exception as e:
            log.warning("oozt_summary.xlsx не обновлён: %s", e)
    rep = build_report(data, table, ev, out, cfg.report_kv, basemap=cfg.basemap,
                       formats=cfg.report_formats, units=units, measures_gdf=measures_gdf)
    log.info("Готово. Отчёт: %s", rep)


def run(cfg: Config):
    out = Path(cfg.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    _setup_logging(out)
    _MIN_SHARE[0] = float(getattr(cfg, "min_share_pct", 0) or 0)
    PROG.phase("Чтение таблицы мероприятий", 0, 2)
    measures = read_measures(cfg.measures_xlsx)
    PROG.phase("Загрузка слоёв", 2, 15)
    # слой стихийных троп скачивается, только если он нужен включённому мероприятию
    cfg._need_informal = any(
        m.get("on") and measure_spec(m).get("p", {}).get("informal_only") for m in measures)
    data = Data(cfg, pad_m=spec_pad_m(measures))
    data.measures = measures
    cfg.save(out / "config_used.json")
    res = Results()
    log.info("Мероприятия при выбранных данных:")
    for m in measures:
        stt = measure_status(m, cfg)
        log.info("  %-6s %-60s %s", m.get("code", ""), m["name"][:60], stt["text"])
        if m.get("on"):
            sp_ = stt["spec"]
            if sp_["rule"]:
                log.info("         правило: %s (%s)", RULES[sp_["rule"]]["title"],
                         "из столбцов таблицы" if sp_["origin"] == "таблица" else "по тексту «Метода выявления»")
            log.info("         как рассчитывается: %s", describe_rule(m))
            for k in stt.get("layers") or []:
                log.info("         слой видов: %s", species_layer_label(k))
    DIAG.__init__(data)
    _dh = _DiagHandler(level=logging.INFO)
    logging.getLogger().addHandler(_dh)
    PROG.phase("Расчёт мероприятий", 15, 45)
    n_on, i_on = sum(1 for m in measures if m.get("on")), 0
    try:
        for m in measures:
            it = DIAG.start(m, measure_status(m, cfg))
            if not m["on"]:
                DIAG.stop()
                continue
            i_on += 1
            PROG.frac((i_on - 1) / max(1, n_on), "Мероприятие %d из %d: %s%s" % (
                i_on, n_on, (m.get("code") + " ") if m.get("code") else "", m["name"]))
            try:
                dispatch_rule(m, data, res)
            except Exception as e:
                it["error"] = f"{type(e).__name__}: {e}"
                log.exception("Ошибка в мероприятии «%s»: %s", m["name"], e)
            DIAG.stop()
    finally:
        logging.getLogger().removeHandler(_dh)
    PROG.phase("Журнал причин и таблицы мероприятий", 45, 52)
    try:
        write_diagnostics(out, data, res)
    except Exception as e:
        log.warning("журнал причин не сформирован: %s", e)

    table = res.long_table()
    if table.empty:
        log.warning("Ни одно мероприятие не назначено")
        if cfg.unit_mode != "eco":
            table = pd.DataFrame(columns=["ООЗТ", "Квартал", "Тип мероприятий", "Мероприятие",
                                          "Тип экосистемы мероприятия", "Пункт Приказа",
                                          "Сезонность", "Основание", "Значение"])
        else:
            table = pd.DataFrame(columns=["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема", "Тип мероприятий",
                                          "Мероприятие", "Тип экосистемы мероприятия", "Пункт Приказа",
                                          "Сезонность", "Основание", "Значение"])
    units = units_table(data, res)
    units.to_csv(out / "units_long.csv", index=False, encoding="utf-8-sig")
    disp = display_columns(cfg)
    table.rename(columns=disp).to_csv(out / "measures_long.csv", index=False, encoding="utf-8-sig")
    try:
        with pd.ExcelWriter(out / "measures_long.xlsx", engine="openpyxl") as xw:
            table.rename(columns=disp).to_excel(xw, index=False, sheet_name="long")
            rt = rules_table(measures, table, cfg)
            if not rt.empty:
                rt.to_excel(xw, index=False, sheet_name="правила")
                from openpyxl.styles import Alignment as _Al2, Font as _Ft2
                wr = xw.book["правила"]
                wr.freeze_panes = "A2"
                for i, w in enumerate([8, 40, 10, 18, 50, 24, 60, 22], 1):
                    wr.column_dimensions[chr(64 + i)].width = w
                for cell in wr[1]:
                    cell.font = _Ft2(bold=True)
                for row in wr.iter_rows():
                    for cell in row:
                        cell.alignment = _Al2(wrap_text=True, vertical="top")
            if not table.empty:
                qmode = cfg.unit_mode != "eco"     # учётная единица — квартал
                order_cols = (["ООЗТ", "Квартал", "Тип мероприятий", "Мероприятие",
                               "Тип экосистемы мероприятия", "Пункт Приказа"] if qmode else
                              ["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема",
                               "Тип мероприятий", "Мероприятие", "Тип экосистемы мероприятия",
                               "Пункт Приказа"])
                by_order = table[order_cols].rename(columns={
                    "Мероприятие": "Мероприятия",
                    "Тип экосистемы мероприятия": "Тип экосистемы",
                    "Пункт Приказа": "Соответствие Пункту Приказа",
                })
                by_order.rename(columns=disp).to_excel(xw, index=False, sheet_name="мероприятия по приказу")

                # Значения не дедуплицируем: каждая строка во всех четырёх колонках
                # должна соответствовать одному и тому же мероприятию.
                join_lines = lambda s: "\n".join(
                    "" if pd.isna(x) else str(x) for x in s
                )
                wide = table.groupby(
                    (["ООЗТ", "Квартал"] if qmode else
                     ["ООЗТ", "Квартал", "Выдел", "Тип выдела", "Экосистема"]),
                    as_index=False, dropna=False
                ).agg(**{
                    "Тип мероприятий": ("Тип мероприятий", join_lines),
                    "Мероприятия": ("Мероприятие", join_lines),
                    "Тип экосистемы": ("Тип экосистемы мероприятия", join_lines),
                    "Соответствие Пункту Приказа": ("Пункт Приказа", join_lines),
                    "Число": ("Мероприятие", "size"),
                })
                wide.rename(columns=disp).to_excel(xw, index=False,
                                                  sheet_name="by_kvartal" if qmode else "by_vydel")
                ex, how = execution_list(table, cfg, data)
                ex.to_excel(xw, index=False, sheet_name="перечень")
                ws = xw.book["перечень"]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for i, c in enumerate(ex.columns, 1):
                    ws.column_dimensions[chr(64 + i)].width = 70 if c.startswith("Мероприятия") else (
                        30 if i == len(ex.columns) - 2 else 14)
                from openpyxl.styles import Alignment as _Al, Font as _Ft, PatternFill as _Pf
                for cell in ws[1]:
                    cell.font = _Ft(bold=True); cell.fill = _Pf("solid", fgColor="D9EAD3")
                    cell.alignment = _Al(wrap_text=True, vertical="center", horizontal="center")
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.alignment = _Al(wrap_text=True, vertical="top")
                from openpyxl.styles import Alignment, Font, PatternFill
                for ws in (xw.book["мероприятия по приказу"], xw.book["by_kvartal" if qmode else "by_vydel"]):
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

    # Сводные таблицы по каждому ООЗТ: общая площадь ООЗТ, мероприятия в штуках, гектарах и %
    # от общей площади — в отдельную книгу oozt_summary.xlsx, в листы measures_long.xlsx и в CSV.
    PROG.phase("Сводки по ООЗТ", 52, 56)
    measures_poly = res.measures_gdf(table, data.union)
    try:
        sum_book = out / "oozt_summary.xlsx"
        sdfs = write_oozt_summary_excel(sum_book, data, table, units, res=res, measures_gdf=measures_poly)
        sm, ms, ts, ks, pv = sdfs["sm"], sdfs["ms"], sdfs["ts"], sdfs["ks"], sdfs["pv"]
        if not sm.empty:
            sm.to_csv(out / "oozt_summary.csv", index=False, encoding="utf-8-sig")
        if not ms.empty:
            ms.to_csv(out / "oozt_summary_by_measure.csv", index=False, encoding="utf-8-sig")
        if not ts.empty:
            ts.to_csv(out / "oozt_summary_by_type.csv", index=False, encoding="utf-8-sig")
        if not ks.empty:
            ks.to_csv(out / "oozt_summary_by_kvartal.csv", index=False, encoding="utf-8-sig")
        if (out / "measures_long.xlsx").exists() and not sm.empty:
            with pd.ExcelWriter(out / "measures_long.xlsx", engine="openpyxl",
                                 mode="a", if_sheet_exists="replace") as xw:
                added_sh = []
                for sh_n, df_s in (("сводка по ООЗТ", sm), ("ООЗТ × мероприятие", ms),
                                   ("ООЗТ × тип мероприятий", ts), ("сравнительная матрица", pv),
                                   ("ООЗТ × кварталы", ks)):
                    if df_s is not None and not df_s.empty:
                        df_s.to_excel(xw, index=False, sheet_name=sh_n)
                        added_sh.append(sh_n)
                _format_summary_sheets(xw.book, added_sh)
        log.info("Сводка по ООЗТ записана: %s (и листы сводки в measures_long.xlsx)", sum_book)
    except Exception as e:
        log.warning("Сводка по ООЗТ не записана: %s", e)

    PROG.phase("GeoPackage и проект QGIS", 56, 62)
    ev = res.evidence_gdf()
    try:
        gpkg = out / "evidence.gpkg"
        if gpkg.exists():
            gpkg.unlink()
        qmode = cfg.unit_mode != "eco"     # учётная единица — квартал
        vm = data.oozt.copy()
        if qmode:
            qs = data.quarters()
            if qs is None or qs.empty:
                qs = vm
            vm = qs.copy()
            if not table.empty:
                agg = table.groupby(["ООЗТ", "Квартал"])["Мероприятие"].agg(
                    lambda s: "; ".join(dict.fromkeys(s)))
                vm = vm.merge(agg.rename("measures").reset_index().rename(
                    columns={"ООЗТ": "OOZT", "Квартал": "KV"}), on=["OOZT", "KV"], how="left")
        else:
            if not table.empty:
                agg = table.groupby(["ООЗТ", "Квартал", "Выдел"])["Мероприятие"].agg(
                    lambda s: "; ".join(dict.fromkeys(s)))
                vm = vm.merge(agg.rename("measures").reset_index().rename(
                    columns={"ООЗТ": "OOZT", "Квартал": "KV", "Выдел": "VYD"}), on=["OOZT", "KV", "VYD"], how="left")
        vm["n_measures"] = (vm["measures"].fillna("").map(lambda x: len([y for y in x.split("; ") if y]))
                            if "measures" in vm.columns else 0)
        vm = vm.drop(columns=["__src", "KV_SORT", "VYD_SORT"], errors="ignore")
        if cfg.unit_mode == "eco":
            # исходное поле ID уже есть в слое; служебный VYD не дублируем, KV — лист / группа отчёта
            vm = vm.drop(columns=["VYD"] if data.F_ID in vm.columns and data.F_ID != "VYD" else [])
            vm = vm.rename(columns={"KV": "report_sheet" if not cfg.report_group_field else "report_group"})
        terr_layer = "kvartal_measures" if qmode else ("vydel_measures" if cfg.unit_mode != "eco" else "eco_measures")
        _prep_gpkg_gdf(vm.to_crs(4326)).to_file(gpkg, layer=terr_layer, driver="GPKG")
        # полигоны проведения мероприятий: СТРОГО географически очерченные участки работ (work_geom),
        # целых кварталов в слое measures нет.
        measures_layer = None
        if measures_poly is not None and not measures_poly.empty:
            ml = measures_poly.copy()
            if qmode and "OOZT" in vm.columns and "KV" in vm.columns:
                q_area_map = {
                    (str(r["OOZT"]), str(r["KV"])): float(r.geometry.area / 1e4)
                    for _, r in vm.iterrows() if r.geometry is not None and not r.geometry.is_empty
                }
                q_areas = [q_area_map.get((str(o), str(k)), 0.0)
                           for o, k in zip(ml["ООЗТ"], ml["Квартал"])]
                ml["Доля квартала, %"] = [
                    round(min(100.0, 100.0 * float(a) / qa), 1) if qa > 0 else None
                    for a, qa in zip(ml["Площадь участка, га"], q_areas)
                ]
            elif not qmode and "Выдел" in ml.columns:
                st_map = {
                    (str(r["OOZT"]), str(r["KV"]), str(r["VYD"])): float(r.get("area_ha", r.geometry.area / 1e4))
                    for _, r in data.oozt.iterrows()
                }
                st_areas = [st_map.get((str(o), str(k), str(v)), 0.0)
                            for o, k, v in zip(ml["ООЗТ"], ml["Квартал"], ml["Выдел"])]
                ml["Доля полигона, %"] = [
                    round(min(100.0, 100.0 * float(a) / sa), 1) if sa > 0 else None
                    for a, sa in zip(ml["Площадь участка, га"], st_areas)
                ]
            ml_out = gpd.GeoDataFrame(ml.rename(columns=disp), geometry="geometry", crs=WORK_EPSG)
            _prep_gpkg_gdf(ml_out.to_crs(4326)).to_file(gpkg, layer="measures", driver="GPKG")
            measures_layer = "measures"
        zones_layer, zone_measures = None, set()
        ev_layers = []
        if ev is not None and not ev.empty:
            for gt, grp in ev.groupby(ev.geom_type.str.replace("Multi", "")):
                _prep_gpkg_gdf(grp.to_crs(4326)).to_file(
                    gpkg, layer=f"evidence_{gt.lower()}", driver="GPKG")
                ev_layers.append((f"evidence_{gt.lower()}",
                                  {"Point": "point", "LineString": "line"}.get(gt, "poly"),
                                  sorted(grp["layer"].astype(str).unique())))
        # ---- контекстные слои подложки (в границах ООЗТ): дороги, гидрография,
        # рекреация, стихийные тропы, горизонтали, все точки ПВН — в GeoPackage для проекта QGIS
        ctx_layers, heatmap_tif = [], None
        try:
            def _ctx_write(name, g, geom, color, width=0.8, alpha=None, cat_field=None, clip_to_oozt=True, keep_cols=None):
                if g is None or g.empty:
                    return
                g = g.to_crs(WORK_EPSG)
                if clip_to_oozt:
                    g = gpd.clip(g, data.union)
                if g.empty:
                    return
                cat_values = (sorted(g[cat_field].astype(str).unique().tolist()[:10])
                              if cat_field and cat_field in g.columns else None)
                cols_k = ["geometry"]
                if keep_cols:
                    cols_k.extend([c for c in keep_cols if c in g.columns and c != "geometry"])
                elif cat_field and cat_field in g.columns:
                    cols_k.append(cat_field)
                _prep_gpkg_gdf(g[cols_k].to_crs(4326)).to_file(gpkg, layer=name, driver="GPKG")
                ctx_layers.append(dict(name=name, geom=geom, color=color, width=width,
                                       alpha=alpha, cat_field=cat_field, cat_values=cat_values))
            if data.lesopat_all is not None and not data.lesopat_all.empty:
                pvn_keep = [c for c in ("pvn", "kv", "vyd", "state", "recomend", "recs") if c in data.lesopat_all.columns]
                _ctx_write("pvn_points", data.lesopat_all, "point", "#d62728", clip_to_oozt=False, keep_cols=pvn_keep)
            _ctx_write("rivers_ditches", data.hydro_lines, "line", "#1F78B4", width=0.8)
            _ctx_write("lakes_ponds", pond_polys(data), "poly", "#A6CEE3", alpha=0.85)
            _ctx_write("wetlands", wetland_polys(data), "poly", WETLAND_COLOR, alpha=0.9)
            paved = _paved_roads(data)
            _ctx_write("roads_paved", paved, "line", "#5f6368", width=1.5)
            if data.roads_dirt is not None and not data.roads_dirt.empty:
                _ctx_write("roads_dirt", data.roads_dirt, "line", "#b5b5b5", width=0.5)
            elif (paved is not None and not paved.empty and (data.roads_paved is None or data.roads_paved.empty)
                  and data.trails is not None and not data.trails.empty):
                _ctx_write("roads_dirt", data.trails[~data.trails.index.isin(paved.index)],
                           "line", "#b5b5b5", width=0.5)
            else:
                _ctx_write("roads_dirt", data.trails, "line", "#b5b5b5", width=0.5)
            _ctx_write("informal_trails", data.trails_informal, "line", "#c98d5a", width=0.8)
            if data.rec_points is not None and not data.rec_points.empty:
                rcat = next((c for c in data.rec_points.columns if c != "geometry" and str(c).lower()
                             in ("category", "cat", "fclass", "type", "категория")), None)
                _ctx_write("recreation_points", data.rec_points, "point", "#d95f02", cat_field=rcat)
                try:
                    pts = data.rec_points[data.rec_points.within(data.union)]
                    if not pts.empty:
                        dens, tr = point_density(pts, data.bounds_pad, radius_m=150.0)
                        inside = data.inside_mask(tr, dens.shape)
                        vals = dens[inside & (dens > 0)]
                        if len(vals) > 10:
                            thr = float(np.quantile(vals, 0.95))
                            field = np.where(inside & (dens >= thr), dens, np.nan).astype("float32")
                            import rasterio as _rio
                            heatmap_tif = out / "recreation_heatmap.tif"
                            with _rio.open(heatmap_tif, "w", driver="GTiff", height=field.shape[0],
                                          width=field.shape[1], count=1, dtype="float32",
                                          crs=f"EPSG:{WORK_EPSG}", transform=tr, nodata=np.nan) as _dst:
                                _dst.write(field, 1)
                except Exception as e:
                    log.warning("  тепловая карта рекреации не записана: %s", e)
            # горизонтали по ЦМР (2,5 м) — вектор линий
            if data.dem is not None:
                try:
                    arr, tr = data.dem
                    inside = data.inside_mask(tr, arr.shape)
                    dem_i = np.where(inside, arr, np.nan)
                    if np.nanmax(dem_i) - np.nanmin(dem_i) > 2:
                        step = 2.5
                        levels = np.arange(np.floor(np.nanmin(dem_i) / step) * step,
                                           np.nanmax(dem_i) + step, step)
                        x = tr.c + np.arange(dem_i.shape[1]) * tr.a
                        y = tr.f - np.arange(dem_i.shape[0]) * tr.e
                        fig_tmp = plt.figure()
                        ax_tmp = fig_tmp.add_subplot(111)
                        cs = ax_tmp.contour(x, y, dem_i, levels=levels)
                        try:
                            segs_by_level = cs.allsegs
                        except Exception:
                            segs_by_level = [[p.vertices for p in coll.get_paths()]
                                              for coll in cs.collections]
                        from shapely.geometry import LineString as _LS
                        lines = []
                        for lvl, segs in zip(levels, segs_by_level):
                            for seg in segs:
                                if len(seg) >= 2:
                                    lines.append({"Высота, м": round(float(lvl), 1),
                                                  "geometry": _LS(seg)})
                        plt.close(fig_tmp)
                        if lines:
                            contours = gpd.GeoDataFrame(lines, crs=WORK_EPSG)
                            contours.to_crs(4326).to_file(gpkg, layer="contours", driver="GPKG")
                            ctx_layers.append(dict(name="contours", geom="line", color="#9e9e9e",
                                                   width=0.3, alpha=None, cat_field=None))
                except Exception as e:
                    log.warning("  горизонтали не записаны: %s", e)
        except Exception as e:
            log.warning("Контекстные слои не записаны: %s", e)
        # справочник водоёмов
        try:
            if data.hydro_polys is not None and not data.hydro_polys.empty:
                wc = data.water_catalog().drop(columns="geometry")
                wc.to_csv(out / "water_catalog.csv", index=False, encoding="utf-8-sig")
                log.info("Справочник водоёмов: %s (%d шт.)", out / "water_catalog.csv", len(wc))
        except Exception as e:
            log.warning("Справочник водоёмов не записан: %s", e)
        log.info("Результаты записаны в GeoPackage: %s (слои %s%s%s)", gpkg, terr_layer,
                 ", measures" if measures_layer else "", "".join(", " + x[0] for x in ev_layers))
        try:
            mt = []
            if not table.empty:
                for t_, g_ in table.groupby("Тип мероприятий", sort=True):
                    mt.append((t_, list(dict.fromkeys(g_.sort_values("Пункт Приказа", key=_nat_sort_key(
                        "Пункт Приказа"))["Мероприятие"]))))
            qp = write_qgis_project(out, cfg, data, gpkg, terr_layer,
                                    data.F_ID if cfg.unit_mode == "eco" else ("KV" if qmode else "VYD"),
                                    measures_layer, mt, ev_layers,
                                    nmax=int(vm["n_measures"].max()) if len(vm) else 0,
                                    zones_layer=zones_layer, zone_measures=zone_measures,
                                    ctx_layers=ctx_layers, heatmap_tif=heatmap_tif)
            log.info("Проект QGIS: %s", qp)
        except Exception as e:
            log.warning("Проект QGIS не создан: %s", e)
    except Exception as e:
        log.warning("GPKG не записан: %s", e)

    rep = build_report(data, table, ev, out, cfg.report_kv, basemap=cfg.basemap,
                       formats=cfg.report_formats, units=units, res=res, measures_gdf=measures_poly)
    log.info("Готово. Таблицы: %s, %s; отчёт: %s",
             out / "measures_long.xlsx", out / "oozt_summary.xlsx", rep)
    return table, rep


class _ProgressLogHandler(logging.Handler):
    """Сообщения журнала → очередь окна прогресса."""

    def __init__(self, q):
        super().__init__(level=logging.INFO)
        self.q = q

    def emit(self, record):
        try:
            self.q.put(("log", record.levelno, record.getMessage()))
            PROG.tick()
        except Exception:
            pass


def run_with_progress(fn, cfg: Config):
    """Запускает fn() (расчёт) в фоновом потоке и показывает окно прогресса: этап, текущее
    мероприятие или схема, последнее сообщение журнала, полоска в %, время, журнал.
    Если окно создать нельзя (нет tkinter / экрана) — обычный запуск без окна."""
    try:
        import tkinter as tk
        from tkinter import ttk, messagebox
        root = tk.Tk()
    except Exception:
        return fn()
    import queue, threading
    q = queue.Queue()
    PROG.__init__()
    PROG.q = q
    h = _ProgressLogHandler(q)
    logging.getLogger().addHandler(h)
    st = {"err": None, "done": False, "warn": 0, "t0": time.time()}

    def work():
        try:
            fn()
        except SystemExit as e:
            if e.code:
                st["err"] = f"выход с кодом {e.code}"
        except BaseException as e:
            st["err"] = f"{type(e).__name__}: {e}"
            log.exception("Ошибка выполнения")
        finally:
            q.put(("done",))

    root.title("Расчёт мероприятий ООЗТ")
    root.geometry("860x560")
    root.minsize(600, 380)
    fr = ttk.Frame(root, padding=12)
    fr.pack(fill="both", expand=True)
    v_phase = tk.StringVar(value="Подготовка…")
    v_step = tk.StringVar(value="")
    v_msg = tk.StringVar(value="")
    v_pct = tk.StringVar(value="0 %")
    v_time = tk.StringVar(value="")
    import tkinter.font as tkfont
    f_head = tkfont.nametofont("TkDefaultFont").copy()
    f_head.configure(size=max(11, int(f_head.cget("size") or 10) + 3), weight="bold")
    ttk.Label(fr, textvariable=v_phase, font=f_head).pack(anchor="w")
    ttk.Label(fr, textvariable=v_step, wraplength=820, justify="left").pack(anchor="w", pady=(4, 0))
    ttk.Label(fr, textvariable=v_msg, foreground="#5F6368", wraplength=820, justify="left").pack(anchor="w",
                                                                                                pady=(2, 6))
    fb = ttk.Frame(fr)
    fb.pack(fill="x")
    pb = ttk.Progressbar(fb, orient="horizontal", mode="determinate", maximum=100)
    pb.pack(side="left", fill="x", expand=True)
    ttk.Label(fb, textvariable=v_pct, width=7, anchor="e").pack(side="left", padx=(8, 0))
    ttk.Label(fr, textvariable=v_time, foreground="#5F6368").pack(anchor="w", pady=(4, 6))
    ft = ttk.Frame(fr)
    ft.pack(fill="both", expand=True)
    txt = tk.Text(ft, height=14, wrap="word", font=("TkFixedFont", 9), state="disabled")
    sb = ttk.Scrollbar(ft, command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    txt.tag_configure("warn", foreground="#B35C00")
    txt.tag_configure("err", foreground="#C62828")
    txt.tag_configure("head", font=("TkFixedFont", 9, "bold"))
    fbt = ttk.Frame(fr)
    fbt.pack(fill="x", pady=(8, 0))
    out_dir = Path(cfg.out_dir) if str(cfg.out_dir or "").strip() else Path.cwd()

    def open_dir():
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(out_dir))            # noqa — только Windows
            else:
                import subprocess
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(out_dir)])
        except Exception as e:
            messagebox.showwarning("Папка результатов", str(e), parent=root)

    def close():
        if not st["done"]:
            if not messagebox.askyesno("Прервать расчёт?",
                                       "Расчёт ещё идёт. Прервать его и закрыть окно?\n"
                                       "Результаты будут неполными.", parent=root):
                return
            logging.shutdown()
            os._exit(1)
        root.destroy()

    b_open = ttk.Button(fbt, text="Открыть папку результатов", command=open_dir, state="disabled")
    b_open.pack(side="left")
    b_close = ttk.Button(fbt, text="Прервать", command=close)
    b_close.pack(side="right")
    root.protocol("WM_DELETE_WINDOW", close)

    def add_line(text, tag=None):
        txt.configure(state="normal")
        txt.insert("end", text + "\n", tag or ())
        n = int(txt.index("end-1c").split(".")[0])
        if n > 3000:
            txt.delete("1.0", f"{n - 3000}.0")
        txt.see("end")
        txt.configure(state="disabled")

    def fmt_t(sec):
        sec = int(sec)
        return f"{sec // 3600}:{sec // 60 % 60:02d}:{sec % 60:02d}" if sec >= 3600 else f"{sec // 60}:{sec % 60:02d}"

    last_phase = [""]

    def poll():
        try:
            while True:
                it = q.get_nowait()
                if it[0] == "phase":
                    title, v = it[1], it[2]
                    m = re.match(r"^(Мероприятие|Обзорная схема|Частная схема|Схема)\b", title)
                    if m:
                        v_step.set(title)
                    else:
                        v_phase.set(title)
                        v_step.set("")
                        if title != last_phase[0]:
                            add_line("— " + title + " —", "head")
                            last_phase[0] = title
                    pb["value"] = v
                    v_pct.set(f"{v:.0f} %")
                elif it[0] == "value":
                    pb["value"] = it[1]
                    v_pct.set(f"{it[1]:.0f} %")
                elif it[0] == "log":
                    lvl, msg = it[1], str(it[2])
                    first = msg.splitlines()[0] if msg else ""
                    if lvl >= logging.WARNING:
                        st["warn"] += 1
                    add_line(first if lvl < logging.WARNING else ("! " + first),
                             "err" if lvl >= logging.ERROR else "warn" if lvl >= logging.WARNING else None)
                    v_msg.set(first[:300])
                elif it[0] == "done":
                    st["done"] = True
        except queue.Empty:
            pass
        el = time.time() - st["t0"]
        v = float(pb["value"] or 0)
        left = f"; осталось примерно {fmt_t(el * (100 - v) / v)}" if 3 < v < 100 and not st["done"] else ""
        v_time.set(f"прошло {fmt_t(el)}{left}; предупреждений: {st['warn']}")
        if st["done"] and b_close["text"] == "Прервать":
            logging.getLogger().removeHandler(h)
            PROG.q = None
            b_close.configure(text="Закрыть")
            b_open.configure(state="normal")
            if st["err"]:
                v_phase.set("Ошибка — расчёт не завершён")
                v_step.set(st["err"])
                add_line("Ошибка: " + st["err"], "err")
                root.bell()
            else:
                pb["value"] = 100
                v_pct.set("100 %")
                v_phase.set("Готово")
                v_step.set(f"Результаты: {out_dir}")
                add_line(f"Готово за {fmt_t(el)}", "head")
        root.after(150, poll)

    threading.Thread(target=work, daemon=True).start()
    root.after(100, poll)
    root.mainloop()
    if st["err"]:
        raise RuntimeError(st["err"])


def _running_under_ipython() -> bool:
    """Ячейка Jupyter (код вставлен в ноутбук): get_ipython есть, а sys.argv —
    аргументы ядра (ipykernel_launcher). При %run sys.argv — путь к скрипту:
    это обычный запуск, автозапуск срабатывает."""
    try:
        get_ipython()                     # noqa: F821 — определён только в IPython
    except NameError:
        return False
    return "ipykernel" in str(sys.argv[0])


def _clean_argv() -> list[str]:
    """Аргументы командной строки без имени программы и служебных аргументов ядра
    Jupyter (ipykernel). Значения флагов (--config путь, --kv 1,2) сохраняются."""
    argv = list(sys.argv[1:])              # без имени программы — как для argparse
    if not _running_under_ipython():
        return argv
    value_flags = {"--config", "--kv", "--formats", "--update-table"}
    bool_flags = {"--gui", "--no-basemap", "--no-split", "--save-config-only", "--report-only",
                  "--progress", "--no-progress"}
    out, i = [], 0
    while i < len(argv):
        a = argv[i]
        if a in value_flags and i + 1 < len(argv):
            out.extend([a, argv[i + 1]])
            i += 2
        elif a in bool_flags:
            out.append(a)
            i += 1
        elif a.startswith("-f"):           # «-f /path/kernel.json» — файл ядра
            i += 1
            if i < len(argv) and not argv[i].startswith("-"):
                i += 1
        else:
            i += 1
    return out


def _build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Природоохранные мероприятия для ООЗТ: кварталы/выделы "
                                             "или слой экосистем с уникальным ID")
    ap.add_argument("--config", help="JSON-конфигурация (без диалогов)")
    ap.add_argument("--kv", help="кварталы / листы / группы для отчёта, через запятую (переопределяет конфиг)")
    ap.add_argument("--gui", action="store_true",
                    help="открыть окно выбора данных, предзаполненное из --config")
    ap.add_argument("--no-basemap", action="store_true", help="не добавлять подложку OSM")
    ap.add_argument("--formats", help="форматы отчёта через запятую: docx,pdf,doc (report.md — всегда); "
                                      "пусто — только Markdown")
    ap.add_argument("--no-split", action="store_true",
                    help="не делить многочастные объекты слоя территории на отдельные контуры")
    ap.add_argument("--save-config-only", action="store_true", help="только собрать конфиг диалогами и сохранить")
    ap.add_argument("--update-table", metavar="XLSX",
                    help="записать в таблицу мероприятий столбцы правил, справку и описания расчёта и выйти")
    ap.add_argument("--progress", action="store_true",
                    help="показывать окно прогресса и при запуске с --config")
    ap.add_argument("--no-progress", action="store_true", help="не показывать окно прогресса")
    ap.add_argument("--report-only", action="store_true",
                    help="не пересчитывать: построить отчёт по measures_long.csv и evidence.gpkg из out_dir")
    return ap


def _apply_args(args, ap):
    """Аргументы → конфигурация (окно выбора данных — если нет --config или задан --gui)."""
    if args.config and not args.gui:
        cfg = Config.load(args.config)
    else:
        cfg = gui_collect_config(Config.load(args.config) if args.config else None)
    if args.kv:
        cfg.report_kv = [int(x) if x.isdigit() else x for x in re.split(r"[,;\s]+", args.kv) if x]
    if args.no_basemap:
        cfg.basemap = False
    if args.no_split:
        cfg.split_multipart = False
    if args.formats is not None:
        cfg.report_formats = [x.strip().lower().lstrip(".") for x in re.split(r"[,;\s]+", args.formats) if x.strip()]
        bad = [x for x in cfg.report_formats if x not in REPORT_FORMATS]
        if bad:
            ap.error("неизвестные форматы: " + ", ".join(bad) + " (допустимы: docx, pdf, doc)")
    return cfg


def _run_with_args(args, ap):
    if args.update_table:
        print(update_measures_xlsx(args.update_table, Config.load(args.config) if args.config else None))
        return
    cfg = _apply_args(args, ap)
    if args.save_config_only:
        Path(cfg.out_dir).mkdir(parents=True, exist_ok=True)
        cfg.save(Path(cfg.out_dir) / "config.json")
        print("Конфигурация сохранена:", Path(cfg.out_dir) / "config.json")
        return
    fn = (lambda: report_only(cfg)) if args.report_only else (lambda: run(cfg))
    gui_used = args.gui or not args.config
    if (gui_used or getattr(args, "progress", False)) and not getattr(args, "no_progress", False):
        run_with_progress(fn, cfg)
    else:
        fn()


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-5s %(message)s",
                        datefmt="%H:%M:%S")
    ap = _build_parser()
    args = ap.parse_args(_clean_argv())
    _run_with_args(args, ap)


def _safe_main() -> int:
    """main с перехватом ошибок: traceback в журнал, код 1 — вместо падения процесса."""
    try:
        main()
        return 0
    except SystemExit as e:
        return int(e.code or 0)
    except Exception:
        log.exception("Ошибка выполнения")
        return 1


if __name__ == "__main__" and not _running_under_ipython():
    sys.exit(_safe_main())
elif __name__ == "__main__":
    # Код вставлен в ячейку Jupyter: всё определено, автозапуска нет.
    # Подсказываем готовые команды с найденными конфигами в текущей папке.
    _cfgs = [p.name for p in Path.cwd().glob("config*.json") if p.name != "config.example.json"]
    print("Код oozt_recs.py загружен (Jupyter, автозапуск выключен). Запуск из ячейки:")
    if _cfgs:
        for _c in sorted(_cfgs)[:3]:
            print(f'    cfg = Config.load("{_c}"); run(cfg)')
        print("  или полный запуск с окном выбора данных: main()")
    else:
        print('    cfg = Config.load("config.json"); run(cfg)   # подставьте свой конфиг')
        print("    или полный запуск с окном выбора данных: main()")
