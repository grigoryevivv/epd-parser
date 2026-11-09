"""
PDF → DWG Converter — графическое приложение для Windows.

Запуск из исходников:  python app.py
Сборка exe:            build.bat  (или GitHub Actions → workflow «Build PDF2DWG exe»)
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converter import (
    DWG_VERSIONS,
    ConversionError,
    ConvertOptions,
    convert_pdf,
    find_oda_converter,
    iter_pdfs,
)

try:  # drag & drop — необязательная зависимость
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _DND = True
except Exception:  # pragma: no cover
    _DND = False

APP_NAME = "PDF → DWG Converter"
APP_VERSION = "1.0.0"
ODA_URL = "https://www.opendesign.com/guestfiles/oda_file_converter"

ACCENT = "#2F7DF6"
ACCENT_HOVER = "#1F63CC"
OK = "#2BB673"
ERR = "#E5484D"
WARN = "#F5A524"
MUTED = ("#6B7280", "#9CA3AF")
CARD = ("#FFFFFF", "#1E2229")
BG = ("#F2F4F7", "#14171C")
ROW = ("#F7F8FA", "#262B33")


def resource_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def settings_file() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home() / ".config")
    p = Path(base) / "PDF2DWG"
    p.mkdir(parents=True, exist_ok=True)
    return p / "settings.json"


def open_in_explorer(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


_Base = ctk.CTk
if _DND:
    class _Base(ctk.CTk, TkinterDnD.DnDWrapper):  # type: ignore[no-redef]
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            self.TkdndVersion = TkinterDnD._require(self)


class FileRow(ctk.CTkFrame):
    """Строка в списке файлов."""

    def __init__(self, master, path: Path, on_remove):
        super().__init__(master, fg_color=ROW, corner_radius=8)
        self.path = path
        self.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self, text="PDF", width=38, height=26, corner_radius=6,
                     fg_color=ERR, text_color="white",
                     font=ctk.CTkFont(size=11, weight="bold")).grid(row=0, column=0, rowspan=2, padx=(10, 10), pady=8)
        ctk.CTkLabel(self, text=path.name, anchor="w",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=1, sticky="ew", pady=(6, 0))
        self.sub = ctk.CTkLabel(self, text=str(path.parent), anchor="w", text_color=MUTED,
                                font=ctk.CTkFont(size=11))
        self.sub.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        self.status = ctk.CTkLabel(self, text="в очереди", width=90, text_color=MUTED,
                                   font=ctk.CTkFont(size=12))
        self.status.grid(row=0, column=2, rowspan=2, padx=6)
        self.btn = ctk.CTkButton(self, text="✕", width=28, height=28, fg_color="transparent",
                                 hover_color=("#E5E7EB", "#374151"), text_color=MUTED,
                                 command=lambda: on_remove(self))
        self.btn.grid(row=0, column=3, rowspan=2, padx=(0, 8))

    def set_status(self, text: str, color=None):
        self.status.configure(text=text, text_color=color or MUTED)


class App(_Base):
    def __init__(self):
        super().__init__()
        ctk.set_default_color_theme("blue")
        self.title(f"{APP_NAME}  v{APP_VERSION}")
        self.geometry("1120x760")
        self.minsize(980, 660)
        self.configure(fg_color=BG)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        self.rows: list[FileRow] = []
        self.msgs: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_flag = threading.Event()
        self.last_out_dir: Path | None = None

        self.s = self._load_settings()
        ctk.set_appearance_mode(self.s.get("theme", "dark"))

        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_files()
        self._build_settings()
        self._build_footer()
        self._refresh_oda()
        self._refresh_empty()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll)

    # ------------------------------------------------------------------ UI
    def _card(self, **grid):
        card = ctk.CTkFrame(self, fg_color=CARD, corner_radius=14)
        card.grid(**grid)
        return card

    def _build_header(self):
        h = ctk.CTkFrame(self, fg_color="transparent")
        h.grid(row=0, column=0, columnspan=2, sticky="ew", padx=20, pady=(16, 8))
        h.grid_columnconfigure(1, weight=1)

        logo = ctk.CTkLabel(h, text="⇄", width=46, height=46, corner_radius=12, fg_color=ACCENT,
                            text_color="white", font=ctk.CTkFont(size=24, weight="bold"))
        logo.grid(row=0, column=0, rowspan=2, padx=(0, 12))
        ctk.CTkLabel(h, text="PDF → DWG", anchor="w",
                     font=ctk.CTkFont(size=22, weight="bold")).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(h, text="Векторные чертежи из PDF в AutoCAD / Revit / nanoCAD — линии, заливки, текст, слои по цвету",
                     anchor="w", text_color=MUTED, font=ctk.CTkFont(size=12)).grid(row=1, column=1, sticky="w")

        self.theme = ctk.CTkSegmentedButton(h, values=["Светлая", "Тёмная"], command=self._set_theme,
                                            width=160)
        self.theme.set("Тёмная" if self.s.get("theme", "dark") == "dark" else "Светлая")
        self.theme.grid(row=0, column=2, rowspan=2)

    def _build_files(self):
        card = self._card(row=1, column=0, sticky="nsew", padx=(20, 10), pady=8)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(card, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        top.grid_columnconfigure(0, weight=1)
        self.files_title = ctk.CTkLabel(top, text="Файлы PDF", anchor="w",
                                        font=ctk.CTkFont(size=16, weight="bold"))
        self.files_title.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(top, text="＋ Добавить файлы", width=150, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=self._add_files).grid(row=0, column=1, padx=4)
        ctk.CTkButton(top, text="📁 Папка", width=90, fg_color=("#E5E7EB", "#2F3540"),
                      hover_color=("#D1D5DB", "#3B4250"), text_color=("#111827", "#E5E7EB"),
                      command=self._add_folder).grid(row=0, column=2, padx=4)
        ctk.CTkButton(top, text="Очистить", width=90, fg_color="transparent", border_width=1,
                      border_color=("#D1D5DB", "#3B4250"), text_color=MUTED,
                      hover_color=("#F3F4F6", "#262B33"), command=self._clear).grid(row=0, column=3, padx=(4, 0))

        self.list = ctk.CTkScrollableFrame(card, fg_color="transparent")
        self.list.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 12))
        self.list.grid_columnconfigure(0, weight=1)

        hint = "Перетащите PDF-файлы или папку сюда" if _DND else "Нажмите «Добавить файлы»"
        self.empty = ctk.CTkFrame(self.list, fg_color="transparent", border_width=2,
                                  border_color=("#D1D5DB", "#333A45"), corner_radius=14, height=260)
        ctk.CTkLabel(self.empty, text="⬇", font=ctk.CTkFont(size=40), text_color=ACCENT).pack(pady=(50, 4))
        ctk.CTkLabel(self.empty, text=hint, font=ctk.CTkFont(size=15, weight="bold")).pack()
        ctk.CTkLabel(self.empty, text="поддерживаются векторные PDF (экспорт из CAD, Revit, Office)",
                     text_color=MUTED).pack(pady=(4, 50))
        self.empty.bind("<Button-1>", lambda e: self._add_files())

        if _DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)

    def _build_settings(self):
        card = self._card(row=1, column=1, sticky="nsew", padx=(10, 20), pady=8)
        card.grid_columnconfigure(0, weight=1)
        sc = ctk.CTkScrollableFrame(card, fg_color="transparent")
        sc.pack(fill="both", expand=True, padx=6, pady=6)
        sc.grid_columnconfigure(0, weight=1)
        r = 0

        def section(text):
            nonlocal r
            ctk.CTkLabel(sc, text=text, anchor="w", font=ctk.CTkFont(size=13, weight="bold")).grid(
                row=r, column=0, sticky="w", padx=10, pady=(12 if r else 6, 4))
            r += 1

        # --- папка вывода
        section("Куда сохранить")
        self.same_dir = ctk.BooleanVar(value=self.s.get("same_dir", True))
        ctk.CTkCheckBox(sc, text="Рядом с исходным PDF", variable=self.same_dir,
                        command=self._toggle_out).grid(row=r, column=0, sticky="w", padx=10, pady=2)
        r += 1
        f = ctk.CTkFrame(sc, fg_color="transparent")
        f.grid(row=r, column=0, sticky="ew", padx=10, pady=4)
        f.grid_columnconfigure(0, weight=1)
        r += 1
        self.out_dir = ctk.CTkEntry(f, placeholder_text="Папка для результатов…")
        self.out_dir.insert(0, self.s.get("out_dir", str(Path.home() / "Documents" / "PDF2DWG")))
        self.out_dir.grid(row=0, column=0, sticky="ew")
        self.out_btn = ctk.CTkButton(f, text="Обзор…", width=80, command=self._choose_out)
        self.out_btn.grid(row=0, column=1, padx=(6, 0))

        # --- формат
        section("Формат")
        self.fmt = ctk.CTkSegmentedButton(sc, values=["DWG", "DXF"], command=lambda _: self._refresh_oda())
        self.fmt.set(self.s.get("format", "DWG"))
        self.fmt.grid(row=r, column=0, sticky="ew", padx=10, pady=2)
        r += 1
        self.ver = ctk.CTkOptionMenu(sc, values=list(DWG_VERSIONS))
        self.ver.set(self.s.get("version", "AutoCAD 2018"))
        self.ver.grid(row=r, column=0, sticky="ew", padx=10, pady=6)
        r += 1

        self.oda_box = ctk.CTkFrame(sc, corner_radius=10)
        self.oda_box.grid(row=r, column=0, sticky="ew", padx=10, pady=(2, 4))
        self.oda_box.grid_columnconfigure(0, weight=1)
        r += 1
        self.oda_lbl = ctk.CTkLabel(self.oda_box, text="", anchor="w", justify="left", wraplength=330,
                                    font=ctk.CTkFont(size=12))
        self.oda_lbl.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(8, 4))
        self.oda_box.bind("<Configure>", lambda e: self.oda_lbl.configure(wraplength=max(150, e.width - 30)))
        self.oda_dl = ctk.CTkButton(self.oda_box, text="Скачать ODA (бесплатно)", height=26,
                                    command=lambda: webbrowser.open(ODA_URL))
        self.oda_dl.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 8))
        ctk.CTkButton(self.oda_box, text="Указать путь…", height=26, width=110,
                      fg_color="transparent", border_width=1, text_color=("#111827", "#E5E7EB"),
                      command=self._choose_oda).grid(row=1, column=1, sticky="e", padx=10, pady=(0, 8))

        # --- масштаб и страницы
        section("Масштаб и страницы")
        g = ctk.CTkFrame(sc, fg_color="transparent")
        g.grid(row=r, column=0, sticky="ew", padx=10)
        g.grid_columnconfigure(3, weight=1)
        r += 1
        ctk.CTkLabel(g, text="Масштаб  1 :").grid(row=0, column=0, sticky="w")
        self.scale = ctk.CTkComboBox(g, values=["1", "10", "20", "25", "50", "100", "200", "500", "1000"],
                                     width=100)
        self.scale.set(str(self.s.get("scale", "1")))
        self.scale.grid(row=0, column=1, sticky="w", padx=(6, 12))
        ctk.CTkLabel(g, text="Страницы").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.pages = ctk.CTkEntry(g, placeholder_text="все   (или например: 1-3, 5, 8-)")
        self.pages.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(6, 0), pady=(6, 0))
        hint = ctk.CTkLabel(sc, text="Единицы — мм. При 1:100 лист PDF превращается в реальные размеры здания.",
                            text_color=MUTED, font=ctk.CTkFont(size=11), anchor="w", justify="left")
        hint.grid(row=r, column=0, sticky="ew", padx=10, pady=(2, 4))
        sc.bind("<Configure>", lambda e: hint.configure(wraplength=max(150, e.width - 40)), add="+")
        r += 1
        self.mode = ctk.CTkSegmentedButton(sc, values=["Файл на страницу", "Все листы в один файл"])
        self.mode.set(self.s.get("mode", "Файл на страницу"))
        self.mode.grid(row=r, column=0, sticky="ew", padx=10, pady=4)
        r += 1

        # --- что переносить
        section("Что переносить")
        self.opt_text = ctk.BooleanVar(value=self.s.get("text", True))
        self.opt_fill = ctk.BooleanVar(value=self.s.get("fills", True))
        self.opt_img = ctk.BooleanVar(value=self.s.get("images", False))
        self.opt_layers = ctk.BooleanVar(value=self.s.get("layers", True))
        for text, var in (("Текст (TEXT)", self.opt_text),
                          ("Заливки (HATCH)", self.opt_fill),
                          ("Растровые изображения (PNG-подложки)", self.opt_img),
                          ("Раскладывать по слоям по цвету", self.opt_layers)):
            ctk.CTkSwitch(sc, text=text, variable=var, progress_color=ACCENT).grid(
                row=r, column=0, sticky="w", padx=10, pady=3)
            r += 1
        q = ctk.CTkFrame(sc, fg_color="transparent")
        q.grid(row=r, column=0, sticky="ew", padx=10, pady=(6, 10))
        q.grid_columnconfigure(1, weight=1)
        r += 1
        ctk.CTkLabel(q, text="Гладкость дуг").grid(row=0, column=0, sticky="w")
        self.curve = ctk.CTkSlider(q, from_=2, to=32, number_of_steps=30, command=self._curve_lbl)
        self.curve.set(self.s.get("curve", 8))
        self.curve.grid(row=0, column=1, sticky="ew", padx=8)
        self.curve_val = ctk.CTkLabel(q, text="", width=30)
        self.curve_val.grid(row=0, column=2)
        self._curve_lbl(self.curve.get())

        self._toggle_out()

    def _build_footer(self):
        card = self._card(row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=(8, 18))
        card.grid_columnconfigure(0, weight=1)

        self.log = ctk.CTkTextbox(card, height=110, fg_color=("#F7F8FA", "#171A20"),
                                  font=ctk.CTkFont(family="Consolas", size=12), wrap="word")
        self.log.grid(row=0, column=0, columnspan=4, sticky="ew", padx=14, pady=(12, 8))
        self.log.configure(state="disabled")

        self.progress = ctk.CTkProgressBar(card, progress_color=ACCENT, height=10)
        self.progress.set(0)
        self.progress.grid(row=1, column=0, sticky="ew", padx=(14, 12), pady=(0, 14))
        self.open_btn = ctk.CTkButton(card, text="Открыть папку", width=130, state="disabled",
                                      fg_color=("#E5E7EB", "#2F3540"), hover_color=("#D1D5DB", "#3B4250"),
                                      text_color=("#111827", "#E5E7EB"), command=self._open_out)
        self.open_btn.grid(row=1, column=1, padx=4, pady=(0, 14))
        self.cancel_btn = ctk.CTkButton(card, text="Стоп", width=80, state="disabled", fg_color=ERR,
                                        hover_color="#C53B40", command=self.cancel_flag.set)
        self.cancel_btn.grid(row=1, column=2, padx=4, pady=(0, 14))
        self.go = ctk.CTkButton(card, text="Конвертировать  ▶", width=200, height=40, fg_color=ACCENT,
                                hover_color=ACCENT_HOVER, font=ctk.CTkFont(size=15, weight="bold"),
                                command=self._start)
        self.go.grid(row=1, column=3, padx=(4, 14), pady=(0, 14))

    # ------------------------------------------------------------ actions
    def _curve_lbl(self, v):
        self.curve_val.configure(text=str(int(float(v))))

    def _set_theme(self, value):
        mode = "dark" if value == "Тёмная" else "light"
        ctk.set_appearance_mode(mode)
        self.s["theme"] = mode

    def _toggle_out(self):
        state = "disabled" if self.same_dir.get() else "normal"
        self.out_dir.configure(state=state)
        self.out_btn.configure(state=state)

    def _choose_out(self):
        d = filedialog.askdirectory(title="Папка для DWG/DXF", initialdir=self.out_dir.get() or None)
        if d:
            self.out_dir.delete(0, "end")
            self.out_dir.insert(0, d)

    def _choose_oda(self):
        f = filedialog.askopenfilename(title="Укажите ODAFileConverter.exe",
                                       filetypes=[("ODAFileConverter", "ODAFileConverter*.exe"),
                                                  ("Программы", "*.exe"), ("Все файлы", "*.*")])
        if f:
            self.s["oda_path"] = f
            self._refresh_oda()

    def _refresh_oda(self):
        if self.fmt.get() == "DXF":
            self.ver.configure(state="disabled")
            self.oda_box.configure(fg_color=("#EEF2FF", "#1E2A44"))
            self.oda_lbl.configure(text="DXF открывается в любой CAD-программе. ODA File Converter не нужен.",
                                   text_color=("#1F2937", "#D1D5DB"))
            self.oda_dl.grid_remove()
            return
        self.ver.configure(state="normal")
        path = find_oda_converter(self.s.get("oda_path", ""))
        if path:
            self.oda_box.configure(fg_color=("#E8F7EF", "#16302A"))
            self.oda_lbl.configure(text=f"✔ ODA File Converter найден\n{path}", text_color=("#14532D", "#86EFAC"))
            self.oda_dl.grid_remove()
            self.s["oda_path"] = path
        else:
            self.oda_box.configure(fg_color=("#FFF7E6", "#33280F"))
            self.oda_lbl.configure(
                text="⚠ Для записи DWG нужен бесплатный ODA File Converter. Установите его "
                     "(один раз) — программа найдёт его автоматически. Без него можно сохранить в DXF.",
                text_color=("#7C4A03", "#FCD34D"))
            self.oda_dl.grid()

    def _refresh_empty(self):
        if self.rows:
            self.empty.grid_forget()
        else:
            self.empty.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        n = len(self.rows)
        self.files_title.configure(text=f"Файлы PDF  ·  {n}" if n else "Файлы PDF")

    def _add_paths(self, paths):
        existing = {r.path.resolve() for r in self.rows}
        added = 0
        for p in iter_pdfs(paths):
            if p.resolve() in existing:
                continue
            row = FileRow(self.list, p, self._remove_row)
            row.grid(row=len(self.rows) + 1, column=0, sticky="ew", padx=4, pady=3)
            self.rows.append(row)
            existing.add(p.resolve())
            added += 1
        if paths and not added:
            self._log("Новых PDF не найдено.")
        self._refresh_empty()

    def _add_files(self):
        files = filedialog.askopenfilenames(title="Выберите PDF", filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")],
                                            initialdir=self.s.get("last_in") or None)
        if files:
            self.s["last_in"] = str(Path(files[0]).parent)
            self._add_paths(files)

    def _add_folder(self):
        d = filedialog.askdirectory(title="Папка с PDF (включая подпапки)", initialdir=self.s.get("last_in") or None)
        if d:
            self.s["last_in"] = d
            self._add_paths([d])

    def _on_drop(self, event):
        self._add_paths(self.tk.splitlist(event.data))

    def _remove_row(self, row: FileRow):
        if self.worker and self.worker.is_alive():
            return
        self.rows.remove(row)
        row.destroy()
        for i, rw in enumerate(self.rows):
            rw.grid(row=i + 1)
        self._refresh_empty()

    def _clear(self):
        if self.worker and self.worker.is_alive():
            return
        for r in self.rows:
            r.destroy()
        self.rows.clear()
        self._refresh_empty()

    def _open_out(self):
        if self.last_out_dir and self.last_out_dir.exists():
            open_in_explorer(self.last_out_dir)

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---------------------------------------------------------- convert
    def _collect_options(self) -> ConvertOptions:
        try:
            scale = float(self.scale.get().replace(",", ".").split(":")[-1])
            if scale <= 0:
                raise ValueError
        except ValueError:
            raise ConversionError("Масштаб должен быть положительным числом, например 100")
        return ConvertOptions(
            out_format=self.fmt.get(),
            dwg_version=DWG_VERSIONS[self.ver.get()],
            scale=scale,
            pages=self.pages.get(),
            pages_mode="single" if self.mode.get().startswith("Все") else "separate",
            include_text=self.opt_text.get(),
            include_fills=self.opt_fill.get(),
            include_images=self.opt_img.get(),
            layers_by_color=self.opt_layers.get(),
            curve_segments=int(self.curve.get()),
            oda_path=self.s.get("oda_path", ""),
        )

    def _start(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.rows:
            messagebox.showinfo(APP_NAME, "Сначала добавьте PDF-файлы.")
            return
        try:
            opts = self._collect_options()
        except ConversionError as e:
            messagebox.showerror(APP_NAME, str(e))
            return
        if opts.out_format == "DWG" and not find_oda_converter(opts.oda_path):
            if messagebox.askyesno(APP_NAME, "ODA File Converter не найден — DWG записать нельзя.\n\n"
                                             "Сохранить в DXF? (DXF открывается в AutoCAD, Revit, nanoCAD и др.)"):
                self.fmt.set("DXF")
                self._refresh_oda()
                opts.out_format = "DXF"
            else:
                return

        out_dir = None if self.same_dir.get() else Path(self.out_dir.get().strip() or ".")
        jobs = [(r, r.path) for r in self.rows]
        for r in self.rows:
            r.set_status("в очереди")
        self._save_settings()

        self.cancel_flag.clear()
        self.go.configure(state="disabled", text="Конвертация…")
        self.cancel_btn.configure(state="normal")
        self.progress.set(0)
        self._log(f"── Старт: {len(jobs)} файл(ов) → {opts.out_format}, масштаб 1:{opts.scale:g}")
        self.worker = threading.Thread(target=self._run, args=(jobs, opts, out_dir), daemon=True)
        self.worker.start()

    def _run(self, jobs, opts: ConvertOptions, out_dir: Path | None):
        put = self.msgs.put
        ok = fail = 0
        total = len(jobs)
        for i, (row, path) in enumerate(jobs):
            if self.cancel_flag.is_set():
                put(("status", row, "отменено", MUTED))
                continue
            put(("status", row, "в работе…", ACCENT))
            put(("log", f"▶ {path.name}"))
            target = out_dir or path.parent
            try:
                res = convert_pdf(
                    path, target, opts,
                    log=lambda m: put(("log", m)),
                    progress=lambda f, i=i: put(("progress", (i + f) / total)),
                    cancelled=self.cancel_flag.is_set,
                )
                for w in res.warnings:
                    put(("log", f"   ⚠ {w}"))
                for o in res.outputs:
                    put(("log", f"   ✔ {o}"))
                put(("status", row, "готово ✔", OK))
                put(("outdir", target))
                ok += 1
            except ConversionError as e:
                put(("log", f"   ✖ {e}"))
                put(("status", row, "ошибка", ERR))
                fail += 1
            except Exception as e:  # неожиданная ошибка — не роняем приложение
                put(("log", f"   ✖ {type(e).__name__}: {e}"))
                put(("status", row, "ошибка", ERR))
                fail += 1
            put(("progress", (i + 1) / total))
        put(("done", ok, fail))

    def _poll(self):
        try:
            while True:
                msg = self.msgs.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._log(msg[1])
                elif kind == "status":
                    msg[1].set_status(msg[2], msg[3])
                elif kind == "progress":
                    self.progress.set(min(1.0, msg[1]))
                elif kind == "outdir":
                    self.last_out_dir = Path(msg[1])
                    self.open_btn.configure(state="normal")
                elif kind == "done":
                    ok, fail = msg[1], msg[2]
                    self._log(f"── Готово: успешно {ok}, ошибок {fail}\n")
                    self.go.configure(state="normal", text="Конвертировать  ▶")
                    self.cancel_btn.configure(state="disabled")
                    if fail == 0 and ok:
                        self.progress.set(1)
        except queue.Empty:
            pass
        self.after(80, self._poll)

    # ---------------------------------------------------------- settings
    def _load_settings(self) -> dict:
        try:
            return json.loads(settings_file().read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_settings(self):
        self.s.update({
            "same_dir": self.same_dir.get(),
            "out_dir": self.out_dir.get(),
            "format": self.fmt.get(),
            "version": self.ver.get(),
            "scale": self.scale.get(),
            "mode": self.mode.get(),
            "text": self.opt_text.get(),
            "fills": self.opt_fill.get(),
            "images": self.opt_img.get(),
            "layers": self.opt_layers.get(),
            "curve": int(self.curve.get()),
        })
        try:
            settings_file().write_text(json.dumps(self.s, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _on_close(self):
        self._save_settings()
        self.cancel_flag.set()
        self.destroy()


def main():
    if sys.platform.startswith("win"):
        try:  # чёткий текст на HiDPI-мониторах и своя иконка на панели задач
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PDF2DWG.Converter")
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
