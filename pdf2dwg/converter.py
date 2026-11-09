"""
Ядро конвертера PDF -> DXF / DWG.

Принцип работы
--------------
1. PyMuPDF читает векторную графику страницы (линии, кривые Безье,
   прямоугольники, четырёхугольники) и текстовые фрагменты.
2. Координаты переводятся из пунктов PDF (1 pt = 1/72") в миллиметры
   с учётом масштаба чертежа (1:N) и поворота страницы, ось Y
   разворачивается (в PDF она направлена вниз, в CAD — вверх).
3. ezdxf создаёт DXF: LWPOLYLINE, HATCH (заливки), TEXT, IMAGE.
4. Для DWG готовый DXF прогоняется через бесплатный ODA File Converter
   (единственный надёжный бесплатный способ записать настоящий DWG).

Растровые (сканированные) PDF векторов не содержат — для них можно
только вставить картинку подложкой (опция «Растровые изображения»).
"""

from __future__ import annotations

import glob
import math
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

import ezdxf
import pymupdf
from ezdxf import colors as dxf_colors
from ezdxf.enums import TextEntityAlignment

PT_TO_MM = 25.4 / 72.0

# Версии DWG, которые понимает ODA File Converter
DWG_VERSIONS = {
    "AutoCAD 2018": "ACAD2018",
    "AutoCAD 2013": "ACAD2013",
    "AutoCAD 2010": "ACAD2010",
    "AutoCAD 2007": "ACAD2007",
    "AutoCAD 2004": "ACAD2004",
    "AutoCAD 2000": "ACAD2000",
}

# Соответствующие версии DXF для ezdxf
_DXF_FOR_DWG = {
    "ACAD2018": "R2018",
    "ACAD2013": "R2013",
    "ACAD2010": "R2010",
    "ACAD2007": "R2007",
    "ACAD2004": "R2004",
    "ACAD2000": "R2000",
}


class ConversionError(Exception):
    pass


@dataclass
class ConvertOptions:
    out_format: str = "DWG"            # "DWG" | "DXF"
    dwg_version: str = "ACAD2018"
    scale: float = 1.0                 # масштаб чертежа 1:N -> N
    pages: str = ""                    # "" = все, иначе "1-3,5"
    pages_mode: str = "separate"       # "separate" | "single"
    include_text: bool = True
    include_fills: bool = True
    include_images: bool = False
    layers_by_color: bool = True
    curve_segments: int = 8            # на сколько отрезков дробить кривую Безье
    oda_path: str = ""                 # путь к ODAFileConverter.exe


@dataclass
class PageStats:
    polylines: int = 0
    hatches: int = 0
    texts: int = 0
    images: int = 0

    def add(self, other: "PageStats") -> None:
        self.polylines += other.polylines
        self.hatches += other.hatches
        self.texts += other.texts
        self.images += other.images


@dataclass
class FileResult:
    source: Path
    outputs: list[Path] = field(default_factory=list)
    stats: PageStats = field(default_factory=PageStats)
    warnings: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# ODA File Converter
# --------------------------------------------------------------------------

def find_oda_converter(custom: str = "") -> str | None:
    """Ищет ODAFileConverter.exe в стандартных местах."""
    if custom and Path(custom).is_file():
        return custom
    candidates: list[str] = []
    if sys.platform.startswith("win"):
        for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(env)
            if base:
                candidates += glob.glob(os.path.join(base, "ODA", "*", "ODAFileConverter.exe"))
                candidates += glob.glob(os.path.join(base, "ODA", "ODAFileConverter.exe"))
    else:
        candidates += glob.glob("/usr/bin/ODAFileConverter*")
        candidates += glob.glob("/opt/ODAFileConverter*/ODAFileConverter")
    exe = shutil.which("ODAFileConverter")
    if exe:
        candidates.append(exe)
    # самая свежая версия — последняя по имени папки
    candidates = sorted(set(c for c in candidates if Path(c).is_file()))
    return candidates[-1] if candidates else None


def dxf_to_dwg(dxf_files: list[Path], out_dir: Path, version: str, oda_exe: str) -> list[Path]:
    """Пакетно конвертирует DXF-файлы (из одной папки) в DWG через ODA."""
    if not dxf_files:
        return []
    src_dir = dxf_files[0].parent
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="pdf2dwg_out_") as tmp_out:
        args = [oda_exe, str(src_dir), tmp_out, version, "DWG", "0", "1", "*.dxf"]
        kwargs: dict = {"capture_output": True, "text": True, "timeout": 600}
        if sys.platform.startswith("win"):
            kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
        try:
            proc = subprocess.run(args, **kwargs)
        except subprocess.TimeoutExpired as exc:
            raise ConversionError("ODA File Converter не ответил за 10 минут") from exc
        except OSError as exc:
            raise ConversionError(f"Не удалось запустить ODA File Converter: {exc}") from exc

        results = []
        for dxf in dxf_files:
            produced = Path(tmp_out) / (dxf.stem + ".dwg")
            if not produced.exists():
                # ODA может сохранить расширение в верхнем регистре
                alt = Path(tmp_out) / (dxf.stem + ".DWG")
                produced = alt if alt.exists() else produced
            if not produced.exists():
                details = (proc.stdout or "") + (proc.stderr or "")
                raise ConversionError(
                    f"ODA File Converter не создал {dxf.stem}.dwg. {details.strip()[:300]}"
                )
            target = out_dir / (dxf.stem + ".dwg")
            shutil.move(str(produced), target)
            results.append(target)
        return results


# --------------------------------------------------------------------------
# Вспомогательные функции
# --------------------------------------------------------------------------

def parse_page_range(spec: str, page_count: int) -> list[int]:
    """'1-3, 5, 8-' -> [0,1,2,4,7,8,...] (индексы с нуля)."""
    spec = (spec or "").strip()
    if not spec:
        return list(range(page_count))
    pages: list[int] = []
    for part in spec.replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start = int(a) if a.strip() else 1
            end = int(b) if b.strip() else page_count
        else:
            start = end = int(part)
        if start > end:
            start, end = end, start
        for p in range(start, end + 1):
            if 1 <= p <= page_count and (p - 1) not in pages:
                pages.append(p - 1)
    if not pages:
        raise ConversionError(f"Диапазон страниц «{spec}» не попадает в документ ({page_count} стр.)")
    return pages


def _rgb(color) -> tuple[int, int, int] | None:
    if color is None:
        return None
    if len(color) == 1:  # grayscale
        color = (color[0],) * 3
    return tuple(max(0, min(255, round(c * 255))) for c in color[:3])  # type: ignore[return-value]


def _layer_name(prefix: str, rgb: tuple[int, int, int] | None) -> str:
    if rgb is None:
        return prefix
    return f"{prefix}_{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _bezier(p0, p1, p2, p3, n: int) -> list[tuple[float, float]]:
    pts = []
    for i in range(1, n + 1):
        t = i / n
        mt = 1 - t
        x = mt**3 * p0[0] + 3 * mt**2 * t * p1[0] + 3 * mt * t**2 * p2[0] + t**3 * p3[0]
        y = mt**3 * p0[1] + 3 * mt**2 * t * p1[1] + 3 * mt * t**2 * p2[1] + t**3 * p3[1]
        pts.append((x, y))
    return pts


def _close(a, b, tol=1e-6) -> bool:
    return abs(a[0] - b[0]) < tol and abs(a[1] - b[1]) < tol


# --------------------------------------------------------------------------
# Конвертация одной страницы в modelspace
# --------------------------------------------------------------------------

class _PageWriter:
    def __init__(self, doc: ezdxf.document.Drawing, opts: ConvertOptions,
                 image_dir: Path | None, image_prefix: str):
        self.doc = doc
        self.msp = doc.modelspace()
        self.opts = opts
        self.image_dir = image_dir
        self.image_prefix = image_prefix
        self.k = PT_TO_MM * opts.scale

    # --- слои --------------------------------------------------------------
    def _layer(self, prefix: str, rgb) -> str:
        name = _layer_name(prefix, rgb) if self.opts.layers_by_color else prefix
        if name not in self.doc.layers:
            layer = self.doc.layers.add(name)
            if self.opts.layers_by_color and rgb is not None:
                layer.rgb = rgb
                # ближайший ACI — для старых программ, не понимающих true color
                layer.color = _nearest_aci(rgb)
        return name

    def _attribs(self, prefix: str, rgb, lineweight_mm: float | None = None) -> dict:
        attrs: dict = {"layer": self._layer(prefix, rgb)}
        if rgb is not None:
            attrs["true_color"] = dxf_colors.rgb2int(rgb)
        if lineweight_mm is not None:
            attrs["lineweight"] = _nearest_lineweight(lineweight_mm)
        return attrs

    # --- страница ----------------------------------------------------------
    def write(self, page: pymupdf.Page, offset_x: float) -> PageStats:
        stats = PageStats()
        # get_drawings / get_text отдают координаты НЕповёрнутой страницы
        rot = page.rotation_matrix
        height = page.rect.height  # высота с учётом поворота

        def tr(p) -> tuple[float, float]:
            q = pymupdf.Point(p) * rot
            return (q.x * self.k + offset_x, (height - q.y) * self.k)

        self._draw_frame(page, tr, offset_x)
        self._write_vectors(page, tr, stats)
        if self.opts.include_images:
            self._write_images(page, tr, stats)
        if self.opts.include_text:
            self._write_text(page, tr, rot, stats)
        return stats

    def _draw_frame(self, page, tr, offset_x):
        r = page.rect
        w, h = r.width * self.k, r.height * self.k
        name = "PDF_PAGE_FRAME"
        if name not in self.doc.layers:
            self.doc.layers.add(name, color=8)
        self.msp.add_lwpolyline(
            [(offset_x, 0), (offset_x + w, 0), (offset_x + w, h), (offset_x, h)],
            close=True, dxfattribs={"layer": name},
        )

    def _write_vectors(self, page, tr, stats: PageStats):
        n = max(2, self.opts.curve_segments)
        for path in page.get_drawings():
            stroke = _rgb(path.get("color"))
            fill = _rgb(path.get("fill"))
            ptype = path.get("type", "s")  # 'f' fill, 's' stroke, 'fs' both
            width_pt = path.get("width") or 0.0
            lw_mm = width_pt * PT_TO_MM if width_pt else None

            # собираем непрерывные контуры
            contours: list[list[tuple[float, float]]] = []
            closed_flags: list[bool] = []
            current: list[tuple[float, float]] = []

            def flush(closed=False):
                nonlocal current
                if len(current) >= 2:
                    contours.append(current)
                    closed_flags.append(closed)
                current = []

            for item in path["items"]:
                op = item[0]
                if op == "l":
                    a, b = tr(item[1]), tr(item[2])
                    if current and _close(current[-1], a):
                        current.append(b)
                    else:
                        flush()
                        current = [a, b]
                elif op == "c":
                    p0, p1, p2, p3 = (tr(item[i]) for i in range(1, 5))
                    if not (current and _close(current[-1], p0)):
                        flush()
                        current = [p0]
                    current.extend(_bezier(p0, p1, p2, p3, n))
                elif op == "re":
                    flush()
                    rect = item[1]
                    quad = rect.quad if hasattr(rect, "quad") else pymupdf.Rect(rect).quad
                    pts = [tr(quad.ul), tr(quad.ur), tr(quad.lr), tr(quad.ll)]
                    contours.append(pts)
                    closed_flags.append(True)
                elif op == "qu":
                    flush()
                    q = item[1]
                    pts = [tr(q.ul), tr(q.ur), tr(q.lr), tr(q.ll)]
                    contours.append(pts)
                    closed_flags.append(True)
            flush(bool(path.get("closePath")))

            # закрытые контуры: последняя точка совпадает с первой
            for i, pts in enumerate(contours):
                if len(pts) > 2 and _close(pts[0], pts[-1]):
                    pts.pop()
                    closed_flags[i] = True

            # заливка
            if self.opts.include_fills and "f" in ptype and fill is not None:
                hatch = self.msp.add_hatch(dxfattribs=self._attribs("PDF_FILL", fill))
                hatch.set_solid_fill(rgb=fill)
                ok = False
                for pts in contours:
                    if len(pts) >= 3:
                        hatch.paths.add_polyline_path(pts, is_closed=True)
                        ok = True
                if ok:
                    stats.hatches += 1
                else:
                    self.msp.delete_entity(hatch)

            # контур
            if "s" in ptype or ("f" in ptype and not self.opts.include_fills):
                color = stroke if "s" in ptype else fill
                attrs = self._attribs("PDF_LINES", color, lw_mm if "s" in ptype else None)
                for pts, closed in zip(contours, closed_flags):
                    if len(pts) >= 2:
                        self.msp.add_lwpolyline(pts, close=closed, dxfattribs=attrs)
                        stats.polylines += 1

    def _write_text(self, page, tr, rot, stats: PageStats):
        data = page.get_text("dict", flags=pymupdf.TEXT_PRESERVE_WHITESPACE)
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                dx, dy = line.get("dir", (1, 0))
                # направление в повёрнутой системе, затем инверсия Y
                d = pymupdf.Point(dx, dy) * rot - pymupdf.Point(0, 0) * rot
                angle = math.degrees(math.atan2(-d.y, d.x))
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if not text.strip():
                        continue
                    size = span.get("size", 0) or 0
                    if size <= 0:
                        continue
                    rgb = _int_to_rgb(span.get("color", 0))
                    attrs = self._attribs("PDF_TEXT", rgb)
                    attrs.update({
                        "height": size * self.k * 0.72,  # cap-height ≈ 0.72 * кегль
                        "rotation": angle,
                        "style": "PDF_TEXT",
                    })
                    x, y = tr(span["origin"])
                    ent = self.msp.add_text(text.strip("\n"), dxfattribs=attrs)
                    ent.set_placement((x, y), align=TextEntityAlignment.LEFT)
                    stats.texts += 1

    def _write_images(self, page, tr, stats: PageStats):
        if self.image_dir is None:
            return
        pdf = page.parent
        for info in page.get_image_info(xrefs=True):
            xref = info.get("xref", 0)
            bbox = pymupdf.Rect(info["bbox"])
            if bbox.is_empty or bbox.width < 1 or bbox.height < 1:
                continue
            try:
                if xref:
                    pix = pymupdf.Pixmap(pdf, xref)
                    if pix.n - pix.alpha >= 4:  # CMYK -> RGB
                        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                else:  # инлайн-картинка — рендерим область страницы
                    pix = page.get_pixmap(clip=bbox, dpi=200)
            except Exception:
                continue
            self.image_dir.mkdir(parents=True, exist_ok=True)
            fname = self.image_dir / f"{self.image_prefix}_p{page.number + 1}_{stats.images + 1}.png"
            pix.save(fname)

            # нижний левый угол и векторы осей (в повёрнутой системе страницы)
            p_ll = tr(bbox.bl)
            p_lr = tr(bbox.br)
            p_ul = tr(bbox.tl)
            u = ((p_lr[0] - p_ll[0]) / pix.width, (p_lr[1] - p_ll[1]) / pix.width)
            v = ((p_ul[0] - p_ll[0]) / pix.height, (p_ul[1] - p_ll[1]) / pix.height)

            img_def = self.doc.add_image_def(
                filename=os.path.relpath(fname, self.image_dir.parent),
                size_in_pixel=(pix.width, pix.height),
            )
            img = self.msp.add_image(
                img_def, insert=p_ll, size_in_units=(1, 1), rotation=0,
                dxfattribs={"layer": self._layer("PDF_IMAGES", None)},
            )
            img.dxf.u_pixel = (u[0], u[1], 0)
            img.dxf.v_pixel = (v[0], v[1], 0)
            stats.images += 1


def _int_to_rgb(value: int) -> tuple[int, int, int]:
    return ((value >> 16) & 255, (value >> 8) & 255, value & 255)


_LINEWEIGHTS = [0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70,
                80, 90, 100, 106, 120, 140, 158, 200, 211]


def _nearest_lineweight(mm: float) -> int:
    target = mm * 100
    return min(_LINEWEIGHTS, key=lambda w: abs(w - target))


def _nearest_aci(rgb: tuple[int, int, int]) -> int:
    # чёрный/белый -> ACI 7 (автоматически инвертируется под фон CAD)
    if max(rgb) < 40 or min(rgb) > 235:
        return 7
    best, best_d = 7, 1e9
    for aci in range(1, 256):
        r, g, b = dxf_colors.aci2rgb(aci)
        d = (r - rgb[0]) ** 2 + (g - rgb[1]) ** 2 + (b - rgb[2]) ** 2
        if d < best_d:
            best, best_d = aci, d
    return best


# --------------------------------------------------------------------------
# Публичный API
# --------------------------------------------------------------------------

def _new_doc(opts: ConvertOptions) -> ezdxf.document.Drawing:
    dxf_version = _DXF_FOR_DWG.get(opts.dwg_version, "R2018")
    doc = ezdxf.new(dxf_version, setup=True, units=ezdxf.units.MM)
    if "PDF_TEXT" not in doc.styles:
        doc.styles.add("PDF_TEXT", font="arial.ttf")
    return doc


def convert_pdf(
    pdf_path: Path | str,
    out_dir: Path | str,
    opts: ConvertOptions,
    log: Callable[[str], None] = print,
    progress: Callable[[float], None] = lambda f: None,
    cancelled: Callable[[], bool] = lambda: False,
) -> FileResult:
    """Конвертирует один PDF. Возвращает список созданных файлов."""
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    result = FileResult(source=pdf_path)

    oda = None
    if opts.out_format == "DWG":
        oda = find_oda_converter(opts.oda_path)
        if not oda:
            raise ConversionError(
                "Для сохранения в DWG нужен бесплатный ODA File Converter. "
                "Установите его или выберите формат DXF."
            )

    try:
        pdf = pymupdf.open(pdf_path)
    except Exception as exc:
        raise ConversionError(f"Не удалось открыть PDF: {exc}") from exc

    with pdf:
        if pdf.needs_pass:
            raise ConversionError("PDF защищён паролем")
        pages = parse_page_range(opts.pages, pdf.page_count)
        stem = pdf_path.stem
        work_dir = Path(tempfile.mkdtemp(prefix="pdf2dwg_")) if oda else out_dir
        image_dir = out_dir / f"{stem}_images" if opts.include_images else None

        dxf_files: list[Path] = []
        try:
            single = opts.pages_mode == "single" or len(pages) == 1
            doc = _new_doc(opts) if single else None
            offset_x = 0.0
            gap = 20.0 * opts.scale  # 20 мм между листами

            for i, pno in enumerate(pages):
                if cancelled():
                    raise ConversionError("Отменено пользователем")
                page = pdf[pno]
                if not single:
                    doc = _new_doc(opts)
                    offset_x = 0.0
                writer = _PageWriter(doc, opts, image_dir, stem)
                stats = writer.write(page, offset_x)
                result.stats.add(stats)

                if stats.polylines == 0 and stats.hatches == 0 and stats.texts == 0:
                    if page.get_images():
                        result.warnings.append(
                            f"стр. {pno + 1}: векторов нет — похоже на скан (только растр)"
                        )
                    else:
                        result.warnings.append(f"стр. {pno + 1}: страница пустая")

                log(f"   стр. {pno + 1}: линий {stats.polylines}, заливок {stats.hatches}, "
                    f"текстов {stats.texts}" + (f", картинок {stats.images}" if stats.images else ""))

                if single:
                    offset_x += page.rect.width * PT_TO_MM * opts.scale + gap
                else:
                    name = f"{stem}_p{pno + 1}.dxf" if len(pages) > 1 else f"{stem}.dxf"
                    dxf_path = work_dir / name
                    doc.saveas(dxf_path)
                    dxf_files.append(dxf_path)
                progress((i + 1) / len(pages) * (0.8 if oda else 1.0))

            if single:
                dxf_path = work_dir / f"{stem}.dxf"
                doc.saveas(dxf_path)
                dxf_files.append(dxf_path)

            if oda:
                log("   ODA File Converter: DXF → DWG …")
                result.outputs = dxf_to_dwg(dxf_files, out_dir, opts.dwg_version, oda)
                if image_dir is not None and result.stats.images:
                    result.warnings.append(
                        "картинки сохранены рядом как PNG и подгружаются в DWG как внешние ссылки"
                    )
            else:
                result.outputs = dxf_files
            progress(1.0)
        finally:
            if oda and work_dir.exists():
                shutil.rmtree(work_dir, ignore_errors=True)
    return result


def iter_pdfs(paths: Iterable[str]) -> list[Path]:
    """Разворачивает папки в список PDF."""
    out: list[Path] = []
    for p in paths:
        path = Path(p)
        if path.is_dir():
            out += sorted(x for x in path.rglob("*") if x.suffix.lower() == ".pdf")
        elif path.suffix.lower() == ".pdf" and path.is_file():
            out.append(path)
    return out
