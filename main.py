# -*- coding: utf-8 -*-
"""
M3D -> STEP / STL для КОМПАС-3D v24.

ИСПРАВЛЕНИЕ STEP:
M3D теперь открывается через API7:
    IApplication.Documents.Open(full_path, visible, read_only)

После этого активный 3D-документ получается через API5:
    KompasObject.ActiveDocument3D()

Это обходит проблему:
    pywintypes.com_error: (-2147352573, 'Kan lid niet vinden.')

STL оставлен по тому же рабочему механизму.
"""

import threading
import traceback
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    import pythoncom
    from win32com.client import Dispatch, GetActiveObject, gencache
except ImportError:
    pythoncom = None
    Dispatch = None
    GetActiveObject = None
    gencache = None


API5_TLIB = "{0422828C-F174-495E-AC5D-D31014DBBE87}"
API7_TLIB = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"

FORMAT_STEP = 3
FORMAT_STL = 6
FORMAT_STEP_AP203 = 203
FORMAT_STEP_AP214 = 214
FORMAT_STEP_AP242 = 242


class KompasConverter:
    def __init__(self, log):
        self.log = log
        self.api5 = None
        self.api7 = None
        self.started_by_script = False
        self.com_initialized = False

    def connect(self):
        if pythoncom is None:
            raise RuntimeError(
                "Не установлен pywin32.\n\n"
                "В CMD выполните:\n"
                "python -m pip install pywin32"
            )

        pythoncom.CoInitialize()
        self.com_initialized = True

        try:
            kapi5 = gencache.EnsureModule(
                API5_TLIB, 0, 1, 0
            )
            kapi7 = gencache.EnsureModule(
                API7_TLIB, 0, 1, 0
            )

            # API5
            try:
                raw5 = GetActiveObject(
                    "Kompas.Application.5"
                )
            except Exception:
                raw5 = Dispatch(
                    "Kompas.Application.5"
                )
                self.started_by_script = True

            self.api5 = kapi5.KompasObject(
                raw5._oleobj_.QueryInterface(
                    kapi5.KompasObject.CLSID,
                    pythoncom.IID_IDispatch
                )
            )

            # API7
            try:
                raw7 = GetActiveObject(
                    "Kompas.Application.7"
                )
            except Exception:
                raw7 = Dispatch(
                    "Kompas.Application.7"
                )
                # Не меняем started_by_script, если API5 уже
                # был найден запущенным пользователем.
                if self.api5 is None:
                    self.started_by_script = True

            self.api7 = kapi7.IApplication(
                raw7._oleobj_.QueryInterface(
                    kapi7.IApplication.CLSID,
                    pythoncom.IID_IDispatch
                )
            )

            try:
                self.api7.Visible = True
            except Exception:
                pass

            self.log(
                "API5 и API7 КОМПАС-3D подключены."
            )

        except Exception as exc:
            self.disconnect()
            raise RuntimeError(
                "Не удалось подключить API5/API7 КОМПАС-3D.\n\n"
                f"{exc}"
            ) from exc

    def disconnect(self):
        try:
            # Закрываем приложение только если именно наш скрипт
            # запускал его. Обычно пользовательский КОМПАС не трогаем.
            if self.started_by_script:
                try:
                    if self.api7 is not None:
                        self.api7.Quit()
                    elif self.api5 is not None:
                        self.api5.Quit()
                except Exception:
                    pass
        finally:
            self.api5 = None
            self.api7 = None

            if self.com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                self.com_initialized = False

    def _open_via_api7(
        self, source: Path
    ):
        """
        Открывает M3D через API7 IDocuments.Open().
        Затем делает документ активным и получает API5
        ActiveDocument3D() для последующего экспорта.
        """
        if self.api7 is None:
            raise RuntimeError("API7 не подключён.")

        documents = self.api7.Documents
        if documents is None:
            raise RuntimeError(
                "Не удалось получить IApplication.Documents."
            )

        # Официальный KsAPI:
        # IKompasDocumentPtr Open(fileName, visible, readOnly)
        api7_doc = documents.Open(
            str(source),
            True,
            False
        )

        if api7_doc is None:
            # Запросим информацию об ошибке приложения, если доступна.
            error_text = ""
            try:
                err = self.api7.GetKompasError()
                if err:
                    error_text = f"\nОшибка КОМПАС: {err}"
            except Exception:
                pass

            raise RuntimeError(
                f"API7 не смог открыть M3D: {source}"
                + error_text
            )

        try:
            api7_doc.SetActive()
        except Exception:
            pass

        # API5 видит активный 3D-документ.
        doc3d = self.api5.ActiveDocument3D()

        if doc3d is None:
            # Последняя попытка — повторно активируем.
            try:
                self.api7.SetActiveDocument(api7_doc)
            except Exception:
                pass

            doc3d = self.api5.ActiveDocument3D()

        if doc3d is None:
            try:
                api7_doc.Close(False)
            except Exception:
                try:
                    api7_doc.Close()
                except Exception:
                    pass

            raise RuntimeError(
                "M3D открылся через API7, но API5 не получил "
                "ActiveDocument3D()."
            )

        return api7_doc, doc3d

    def _open_via_api5_fallback(
        self, source: Path
    ):
        """
        Запасной путь только для случая, когда API7 Open()
        недоступен в конкретной установке.
        """
        self.log(
            "API7 Open() не сработал, пробуем API5 Document3D()."
        )

        doc3d = self.api5.Document3D()

        if doc3d is None:
            raise RuntimeError(
                "API5 Document3D() вернул пустой интерфейс."
            )

        if not doc3d.Open(str(source), False):
            raise RuntimeError(
                f"КОМПАС не смог открыть M3D: {source}"
            )

        return None, doc3d

    def convert_cdw_to_pdf(
        self,
        source: Path,
        destination: Path,
        overwrite: bool,
    ):
        """
        Конвертация CDW -> PDF без конфигурационных файлов.

        Использует штатный конвертер Pdf2d.dll через API7.IConverter.
        Путь к Bin берётся через API5 ksSystemPath(5), с fallback
        на стандартные каталоги установки КОМПАС-3D v24.
        """
        target = destination / (source.stem + ".pdf")

        if target.exists() and not overwrite:
            return "skipped", str(target)

        pdf_dll = None

        # Стандартный способ: API5 сообщает системный путь Bin.
        try:
            bin_dir = self.api5.ksSystemPath(5)
            if bin_dir:
                candidate = Path(str(bin_dir)) / "Pdf2d.dll"
                if candidate.is_file():
                    pdf_dll = candidate
        except Exception:
            pass

        # Дополнительный поиск по типичным каталогам.
        if pdf_dll is None:
            candidates = [
                Path(r"C:\Program Files\ASCON\KOMPAS-3D v24\Bin\Pdf2d.dll"),
                Path(r"C:\Program Files\ASCON\KOMPAS-3D V24\Bin\Pdf2d.dll"),
                Path(r"C:\Program Files (x86)\ASCON\KOMPAS-3D v24\Bin\Pdf2d.dll"),
                Path(r"C:\Program Files (x86)\ASCON\KOMPAS-3D V24\Bin\Pdf2d.dll"),
            ]
            for candidate in candidates:
                if candidate.is_file():
                    pdf_dll = candidate
                    break

        if pdf_dll is None:
            raise RuntimeError(
                "Не найден Pdf2d.dll — штатный конвертер PDF КОМПАС-3D."
            )

        try:
            converter = self.api7.Converter(str(pdf_dll))
        except Exception as exc:
            raise RuntimeError(
                f"Не удалось получить Pdf2d.dll через API7:\n"
                f"{pdf_dll}\n\n{exc}"
            ) from exc

        if converter is None:
            raise RuntimeError(
                f"API7 не вернул IConverter для:\n{pdf_dll}"
            )

        try:
            result = converter.Convert(
                str(source),
                str(target),
                0,
                False,
            )
        except Exception as exc:
            raise RuntimeError(
                f"Ошибка конвертации CDW -> PDF:\n{exc}"
            ) from exc

        if not result:
            # Попытка получить штатное описание ошибки КОМПАС.
            error_text = ""
            try:
                error_text = str(self.api7.KompasError)
            except Exception:
                pass

            raise RuntimeError(
                "Pdf2d.dll не смог выполнить конвертацию."
                + (f"\nОшибка КОМПАС: {error_text}" if error_text else "")
            )

        if not target.exists():
            raise RuntimeError(
                "Конвертер сообщил об успехе, но PDF-файл не найден."
            )

        return "ok", str(target)

    def convert_one(
        self,
        source: Path,
        destination: Path,
        output_format: str,
        step_format: int,
        config_path: Optional[str],
        overwrite: bool,
    ):
        api7_doc = None
        doc3d = None

        try:
            target = destination / (
                source.stem
                + (".step" if output_format == "STEP" else ".stl")
            )

            if target.exists() and not overwrite:
                return "skipped", str(target)

            # Основной путь для обоих форматов — API7 Open.
            # Это особенно важно для STEP.
            try:
                api7_doc, doc3d = self._open_via_api7(source)
            except Exception as open_exc:
                # Не маскируем ошибку API7 для STEP:
                # fallback оставляем только как резервный механизм.
                self.log(
                    f"    API7 Open(): {open_exc}"
                )
                api7_doc, doc3d = (
                    self._open_via_api5_fallback(source)
                )

            params = doc3d.AdditionFormatParam()
            if params is None:
                raise RuntimeError(
                    "Не удалось получить AdditionFormatParam()."
                )

            params.Init()

            if output_format == "STEP":
                params.format = step_format

            else:
                params.format = FORMAT_STL
                try:
                    params.formatBinary = True
                except Exception:
                    pass

            if config_path:
                if not params.LoadConfigurationFile(
                    str(config_path)
                ):
                    raise RuntimeError(
                        "Не удалось загрузить конфигурацию:\n"
                        + str(config_path)
                    )

            if not doc3d.SaveAsToAdditionFormat(
                str(target), params
            ):
                raise RuntimeError(
                    "SaveAsToAdditionFormat() вернул FALSE."
                )

            if not target.exists():
                raise RuntimeError(
                    "КОМПАС вернул TRUE, но файл не появился."
                )

            return "ok", str(target)

        finally:
            # Предпочтительно закрываем API7-документ.
            if api7_doc is not None:
                try:
                    api7_doc.Close(False)
                except Exception:
                    try:
                        api7_doc.Close()
                    except Exception:
                        pass
            elif doc3d is not None:
                try:
                    doc3d.Close()
                except Exception:
                    pass



class App(tk.Tk):
    BG = "#0f172a"
    PANEL = "#111c31"
    PANEL2 = "#16233d"
    BORDER = "#243555"
    TEXT = "#e5edf9"
    MUTED = "#91a4c4"
    ACCENT = "#4f8cff"
    ACCENT_HOVER = "#6ba0ff"
    SUCCESS = "#39c78a"
    WARNING = "#f6c85f"
    ERROR = "#ff6b81"
    INPUT = "#0c1528"

    def __init__(self):
        super().__init__()

        self.title("M3D Converter • КОМПАС-3D v24")
        self.geometry("980x760")
        self.minsize(900, 680)
        self.configure(bg=self.BG)

        self.operation = tk.StringVar(value="3D")
        self.source = tk.StringVar()
        self.destination = tk.StringVar()
        self.format = tk.StringVar(value="STEP")
        self.step = tk.StringVar(value="AP242")
        self.sxcf = tk.StringVar()

        self.use_lxcf = tk.BooleanVar(value=False)
        self.lxcf = tk.StringVar()

        self.recursive = tk.BooleanVar(value=False)
        self.overwrite = tk.BooleanVar(value=False)

        self.running = False
        self.total = 0
        self.ok = 0
        self.skipped = 0
        self.errors = 0

        self._setup_style()
        self._build_ui()
        self._set_initial_mode()

    # ---------- styling ----------

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            ".",
            background=self.BG,
            foreground=self.TEXT,
            font=("Segoe UI", 10),
        )

        style.configure(
            "TFrame",
            background=self.BG,
        )
        style.configure(
            "Panel.TFrame",
            background=self.PANEL,
        )
        style.configure(
            "Inner.TFrame",
            background=self.PANEL2,
        )

        style.configure(
            "TLabel",
            background=self.BG,
            foreground=self.TEXT,
        )
        style.configure(
            "Panel.TLabel",
            background=self.PANEL,
            foreground=self.TEXT,
        )
        style.configure(
            "Muted.TLabel",
            background=self.BG,
            foreground=self.MUTED,
        )
        style.configure(
            "PanelMuted.TLabel",
            background=self.PANEL,
            foreground=self.MUTED,
        )
        style.configure(
            "Title.TLabel",
            background=self.BG,
            foreground=self.TEXT,
            font=("Segoe UI Semibold", 21),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.BG,
            foreground=self.MUTED,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Section.TLabel",
            background=self.PANEL,
            foreground=self.TEXT,
            font=("Segoe UI Semibold", 11),
        )
        style.configure(
            "Value.TLabel",
            background=self.PANEL2,
            foreground=self.TEXT,
            font=("Segoe UI Semibold", 11),
        )

        style.configure(
            "TButton",
            background=self.PANEL2,
            foreground=self.TEXT,
            borderwidth=0,
            focusthickness=0,
            padding=(14, 9),
            font=("Segoe UI Semibold", 9.5),
        )
        style.map(
            "TButton",
            background=[("active", self.BORDER)],
            foreground=[("disabled", "#5d6b84")],
        )

        style.configure(
            "Accent.TButton",
            background=self.ACCENT,
            foreground="#ffffff",
            padding=(18, 11),
            font=("Segoe UI Semibold", 10),
        )
        style.map(
            "Accent.TButton",
            background=[
                ("active", self.ACCENT_HOVER),
                ("pressed", self.ACCENT_HOVER),
                ("disabled", "#314a78"),
            ],
        )

        style.configure(
            "Small.TButton",
            padding=(10, 7),
            font=("Segoe UI Semibold", 9),
        )

        style.configure(
            "TEntry",
            fieldbackground=self.INPUT,
            foreground=self.TEXT,
            insertcolor=self.TEXT,
            bordercolor=self.BORDER,
            lightcolor=self.BORDER,
            darkcolor=self.BORDER,
            padding=(10, 8),
        )

        style.map(
            "TEntry",
            fieldbackground=[("disabled", "#10192b")],
            foreground=[("disabled", "#64748b")],
        )

        style.configure(
            "TCombobox",
            fieldbackground=self.INPUT,
            background=self.INPUT,
            foreground=self.TEXT,
            arrowcolor=self.TEXT,
            bordercolor=self.BORDER,
            lightcolor=self.BORDER,
            darkcolor=self.BORDER,
            padding=(9, 7),
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", self.INPUT)],
            foreground=[("readonly", self.TEXT)],
            selectbackground=[("readonly", self.ACCENT)],
            selectforeground=[("readonly", "#ffffff")],
        )

        style.configure(
            "TCheckbutton",
            background=self.PANEL,
            foreground=self.TEXT,
            font=("Segoe UI Semibold", 10),
            padding=(3, 3),
        )
        style.map(
            "TCheckbutton",
            background=[("active", self.PANEL)],
            foreground=[("active", self.TEXT)],
        )

        style.configure(
            "Horizontal.TProgressbar",
            background=self.ACCENT,
            troughcolor=self.INPUT,
            bordercolor=self.INPUT,
            lightcolor=self.ACCENT,
            darkcolor=self.ACCENT,
            thickness=9,
        )

    def _card(self, parent):
        frame = tk.Frame(
            parent,
            bg=self.PANEL,
            highlightbackground=self.BORDER,
            highlightthickness=1,
            bd=0,
        )
        return frame

    def _make_badge(self, parent, text, color):
        return tk.Label(
            parent,
            text=text,
            bg=color,
            fg="#ffffff",
            font=("Segoe UI Semibold", 9),
            padx=10,
            pady=5,
        )

    def _section_title(self, parent, title, subtitle=None):
        wrap = tk.Frame(parent, bg=self.PANEL)
        wrap.pack(fill="x", padx=16, pady=(14, 8))

        tk.Label(
            wrap,
            text=title,
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w")

        if subtitle:
            tk.Label(
                wrap,
                text=subtitle,
                bg=self.PANEL,
                fg=self.MUTED,
                font=("Segoe UI", 9),
            ).pack(anchor="w", pady=(2, 0))

        return wrap

    # ---------- UI ----------

    def _set_initial_mode(self):
        # Все виджеты уже созданы.
        self.operation.set("3D")
        self.format.set("STEP")

        self.op_3d.configure(style="Accent.TButton")
        self.op_cdw.configure(style="TButton")

        self.pdf_card.pack_forget()
        self.export_card.pack(
            fill="x", pady=(0, 12),
            before=self._action_card
        )

        self.start_button.configure(
            text="▶  НАЧАТЬ КОНВЕРТАЦИЮ"
        )

        self._update_path_labels(
            "M3D", "Исходная папка M3D"
        )
        self.update_controls()

    def _build_ui(self):
        outer = tk.Frame(self, bg=self.BG)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            outer,
            bg=self.BG,
            highlightthickness=0,
            bd=0,
        )
        scrollbar = ttk.Scrollbar(
            outer, orient="vertical", command=canvas.yview
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content = tk.Frame(canvas, bg=self.BG)
        window_id = canvas.create_window(
            (0, 0), window=content, anchor="nw"
        )

        def on_content_configure(_):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            canvas.itemconfigure(window_id, width=event.width)

        content.bind("<Configure>", on_content_configure)
        canvas.bind("<Configure>", on_canvas_configure)

        # Header
        header = tk.Frame(content, bg=self.BG)
        header.pack(fill="x", padx=28, pady=(24, 16))

        logo = tk.Frame(
            header,
            bg=self.ACCENT,
            width=48,
            height=48,
        )
        logo.pack(side="left")
        logo.pack_propagate(False)
        tk.Label(
            logo,
            text="M3",
            bg=self.ACCENT,
            fg="#ffffff",
            font=("Segoe UI Black", 15),
        ).pack(expand=True)

        title_box = tk.Frame(header, bg=self.BG)
        title_box.pack(side="left", padx=14)

        ttk.Label(
            title_box,
            text="Пакетный конвертер M3D → STEP / STL",
            style="Title.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            title_box,
            text="КОМПАС-3D v24 • M3D→STEP/STL • CDW→PDF",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        self.header_badge = self._make_badge(
            header, "ГОТОВО", self.SUCCESS
        )
        self.header_badge.pack(side="right", anchor="n")

        # Operation switch
        op_card = self._card(content)
        op_card.pack(fill="x", padx=28, pady=(0, 12))

        op_inner = tk.Frame(op_card, bg=self.PANEL)
        op_inner.pack(fill="x", padx=16, pady=14)

        tk.Label(
            op_inner,
            text="Режим",
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 11),
        ).pack(side="left", padx=(0, 14))

        self.op_3d = ttk.Button(
            op_inner,
            text="▣  3D МОДЕЛИ  →  STEP / STL",
            style="Accent.TButton",
            command=lambda: self.select_operation("3D"),
        )
        self.op_3d.pack(side="left", padx=(0, 8))

        self.op_cdw = ttk.Button(
            op_inner,
            text="▤  ЧЕРТЕЖИ  →  PDF",
            style="TButton",
            command=lambda: self.select_operation("CDW"),
        )
        self.op_cdw.pack(side="left")

        self.op_hint = tk.Label(
            op_inner,
            text="Без конфигураций",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        )
        self.op_hint.pack(side="right")

        # Main cards
        main = tk.Frame(content, bg=self.BG)
        main.pack(fill="both", expand=True, padx=28, pady=(0, 24))

        # Paths card
        paths = self._card(main)
        paths.pack(fill="x", pady=(0, 12))

        self._section_title(
            paths,
            "1  •  Исходные данные",
            "Укажите исходную папку и папку для сохранения результатов.",
        )

        path_rows = tk.Frame(paths, bg=self.PANEL)
        self.path_rows = path_rows
        path_rows.pack(fill="x", padx=16, pady=(0, 16))
        path_rows.columnconfigure(1, weight=1)

        self._path_row(
            path_rows, 0, "M3D",
            "Исходная папка", self.source,
            self.pick_source
        )
        self._path_row(
            path_rows, 1, "OUT",
            "Папка назначения", self.destination,
            self.pick_destination
        )

        # Export card
        export = self._card(main)
        self.export_card = export
        export.pack(fill="x", pady=(0, 12))

        self._section_title(
            export,
            "2  •  Формат экспорта",
            "Выберите формат. Параметры STEP и STL переключаются автоматически.",
        )

        format_wrap = tk.Frame(export, bg=self.PANEL)
        format_wrap.pack(fill="x", padx=16)

        self._format_card(
            format_wrap,
            "STEP",
            "CAD-модель",
            "step",
            self.ACCENT,
        )
        self._format_card(
            format_wrap,
            "STL",
            "3D-печать",
            "stl",
            self.SUCCESS,
        )

        sep = tk.Frame(export, bg=self.BORDER, height=1)
        sep.pack(fill="x", padx=16, pady=16)

        settings = tk.Frame(export, bg=self.PANEL)
        settings.pack(fill="x", padx=16, pady=(0, 16))
        settings.columnconfigure(1, weight=1)

        tk.Label(
            settings,
            text="STEP AP",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI Semibold", 9),
        ).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=6)

        self.step_box = ttk.Combobox(
            settings,
            textvariable=self.step,
            values=("AP203", "AP214", "AP242"),
            state="readonly",
            width=12,
        )
        self.step_box.grid(
            row=0, column=1, sticky="w", pady=6
        )

        tk.Label(
            settings,
            text=".sxcf",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI Semibold", 9),
        ).grid(row=1, column=0, sticky="w", padx=(0, 12), pady=6)

        self.sxcf_entry = ttk.Entry(
            settings, textvariable=self.sxcf
        )
        self.sxcf_entry.grid(
            row=1, column=1, sticky="ew", pady=6
        )
        self.sxcf_button = ttk.Button(
            settings,
            text="Выбрать",
            style="Small.TButton",
            command=self.pick_sxcf,
        )
        self.sxcf_button.grid(
            row=1, column=2, padx=(8, 0), pady=6
        )

        self.lxcf_check = ttk.Checkbutton(
            settings,
            text="Использовать .lxcf для STL",
            variable=self.use_lxcf,
            command=self.update_controls,
        )
        self.lxcf_check.grid(
            row=2, column=0, columnspan=2,
            sticky="w", pady=8
        )

        self.lxcf_entry = ttk.Entry(
            settings, textvariable=self.lxcf
        )
        self.lxcf_entry.grid(
            row=3, column=0, columnspan=2,
            sticky="ew", pady=6
        )
        self.lxcf_button = ttk.Button(
            settings,
            text="Выбрать",
            style="Small.TButton",
            command=self.pick_lxcf,
        )
        self.lxcf_button.grid(
            row=3, column=2, padx=(8, 0), pady=6
        )

        # Additional options
        opts = tk.Frame(export, bg=self.PANEL2)
        opts.pack(fill="x", padx=16, pady=(0, 16))

        ttk.Checkbutton(
            opts,
            text="Обрабатывать подпапки",
            variable=self.recursive,
        ).pack(side="left", padx=12, pady=10)

        ttk.Checkbutton(
            opts,
            text="Перезаписывать файлы",
            variable=self.overwrite,
        ).pack(side="left", padx=12, pady=10)

        # PDF card (CDW only)
        self.pdf_card = self._card(main)

        self._section_title(
            self.pdf_card,
            "2  •  Чертежи CDW → PDF",
            "Экспорт выполняется штатным конвертером КОМПАС без .lxcf/.sxcf.",
        )

        pdf_note = tk.Frame(
            self.pdf_card, bg=self.PANEL
        )
        pdf_note.pack(
            fill="x", padx=16, pady=(0, 16)
        )

        tk.Label(
            pdf_note,
            text="PDF создаётся с тем же именем, что и исходный CDW.",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        self.pdf_info = tk.Label(
            pdf_note,
            text="",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        )
        self.pdf_info.pack(anchor="w", pady=(6, 0))

        # Action / stats
        action = self._card(main)
        self._action_card = action
        action.pack(fill="x", pady=(0, 12))

        action_inner = tk.Frame(action, bg=self.PANEL)
        action_inner.pack(fill="x", padx=16, pady=16)

        self.start_button = ttk.Button(
            action_inner,
            text="▶  НАЧАТЬ КОНВЕРТАЦИЮ",
            style="Accent.TButton",
            command=self.start_conversion,
        )
        self.start_button.pack(side="left")

        self.stats = tk.Label(
            action_inner,
            text="0 файлов  •  0 успешно  •  0 ошибок",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        )
        self.stats.pack(side="right")

        self.progress = ttk.Progressbar(
            action,
            style="Horizontal.TProgressbar",
            mode="determinate",
        )
        self.progress.pack(
            fill="x", padx=16, pady=(0, 6)
        )

        self.status = tk.StringVar(
            value="Готово к запуску."
        )
        tk.Label(
            action,
            textvariable=self.status,
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(
            fill="x", padx=16, pady=(0, 14)
        )

        # Log
        log_card = self._card(main)
        log_card.pack(fill="both", expand=True)

        log_header = tk.Frame(log_card, bg=self.PANEL)
        log_header.pack(fill="x", padx=16, pady=(14, 8))

        tk.Label(
            log_header,
            text="Журнал обработки",
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 11),
        ).pack(side="left")

        ttk.Button(
            log_header,
            text="Очистить",
            style="Small.TButton",
            command=self.clear_log,
        ).pack(side="right")

        text_frame = tk.Frame(
            log_card,
            bg=self.INPUT,
            highlightbackground=self.BORDER,
            highlightthickness=1,
        )
        text_frame.pack(
            fill="both", expand=True,
            padx=16, pady=(0, 16)
        )

        self.logbox = ScrolledText(
            text_frame,
            wrap="word",
            bg=self.INPUT,
            fg="#cbd7ea",
            insertbackground=self.TEXT,
            selectbackground=self.ACCENT,
            selectforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Consolas", 9),
            padx=12,
            pady=10,
        )
        self.logbox.pack(
            fill="both", expand=True
        )
        self.logbox.configure(state="disabled")

        self.logbox.tag_configure(
            "ok", foreground=self.SUCCESS
        )
        self.logbox.tag_configure(
            "error", foreground=self.ERROR
        )
        self.logbox.tag_configure(
            "warn", foreground=self.WARNING
        )
        self.logbox.tag_configure(
            "header", foreground="#8fb6ff"
        )

        # Mouse wheel over main content
        def wheel(event):
            canvas.yview_scroll(
                int(-1 * (event.delta / 120)),
                "units"
            )

        canvas.bind_all("<MouseWheel>", wheel)

    def _path_row(
        self, parent, row, short, label,
        variable, command
    ):
        badge = tk.Label(
            parent,
            text=short,
            bg=self.PANEL2,
            fg=self.ACCENT_HOVER,
            font=("Segoe UI Semibold", 8),
            padx=8, pady=5,
        )
        badge.grid(
            row=row, column=0,
            padx=(0, 10), pady=6
        )
        if row == 0:
            self.source_badge = badge

        box = tk.Frame(parent, bg=self.PANEL)
        box.grid(
            row=row, column=1,
            sticky="ew", pady=6
        )
        box.columnconfigure(0, weight=1)

        label_widget = tk.Label(
            box,
            text=label,
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        )
        label_widget.grid(row=0, column=0, sticky="w")
        if row == 0:
            self.source_path_label = label_widget

        ttk.Entry(
            box, textvariable=variable
        ).grid(
            row=1, column=0,
            sticky="ew", pady=(2, 0)
        )

        ttk.Button(
            parent,
            text="Выбрать",
            style="Small.TButton",
            command=command,
        ).grid(
            row=row, column=2,
            padx=(10, 0), pady=6
        )

    def _format_card(
        self, parent, title, subtitle, value, accent
    ):
        frame = tk.Frame(
            parent,
            bg=self.PANEL2,
            highlightbackground=self.BORDER,
            highlightthickness=1,
            width=220,
            height=92,
            cursor="hand2",
        )
        frame.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 10) if value == "step" else (10, 0),
        )
        frame.pack_propagate(False)

        top = tk.Frame(frame, bg=self.PANEL2)
        top.pack(fill="x", padx=14, pady=(13, 3))

        dot = tk.Label(
            top,
            text="●",
            bg=self.PANEL2,
            fg=accent,
            font=("Segoe UI", 12),
        )
        dot.pack(side="left")

        tk.Label(
            top,
            text=title,
            bg=self.PANEL2,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 11),
        ).pack(side="left", padx=7)

        tk.Label(
            frame,
            text=subtitle,
            bg=self.PANEL2,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=14)

        def select(_=None):
            self.format.set(title)
            self.update_controls()

        for widget in (frame, top, dot):
            widget.bind("<Button-1>", select)

    # ---------- controls / actions ----------

    def select_operation(self, operation):
        self.operation.set(operation)

        if operation == "3D":
            self.op_3d.configure(style="Accent.TButton")
            self.op_cdw.configure(style="TButton")

            self.pdf_card.pack_forget()
            self.export_card.pack(
                fill="x",
                pady=(0, 12),
                before=self._action_card
            )

            self.format.set(
                "STEP"
                if self.format.get() == "PDF"
                else self.format.get()
            )
            self.start_button.configure(
                text="▶  НАЧАТЬ КОНВЕРТАЦИЮ"
            )
            self._update_path_labels(
                "M3D", "Исходная папка M3D"
            )

        else:
            self.op_3d.configure(style="TButton")
            self.op_cdw.configure(style="Accent.TButton")

            self.export_card.pack_forget()
            self.pdf_card.pack(
                fill="x",
                pady=(0, 12),
                before=self._action_card
            )

            self.format.set("PDF")
            self.start_button.configure(
                text="▶  КОНВЕРТИРОВАТЬ CDW → PDF"
            )
            self._update_path_labels(
                "CDW", "Исходная папка CDW"
            )
            self._update_pdf_info()

        self.update_controls()
        self._update_badge()

    def _update_path_labels(self, short, label):
        try:
            self.source_badge.configure(text=short)
            self.source_path_label.configure(text=label)
        except Exception:
            pass

    def _update_pdf_info(self):
        self.pdf_info.configure(
            text="Конфигурации .lxcf и .sxcf не используются."
        )

    def update_controls(self):
        if self.operation.get() == "CDW":
            self._update_pdf_info()
            return

        is_step = self.format.get() == "STEP"

        self.step_box.configure(
            state="readonly" if is_step else "disabled"
        )
        self.sxcf_entry.configure(
            state="normal" if is_step else "disabled"
        )
        self.sxcf_button.configure(
            state="normal" if is_step else "disabled"
        )

        self.lxcf_check.configure(
            state="disabled" if is_step else "normal"
        )
        lxcf_enabled = (
            not is_step and self.use_lxcf.get()
        )
        self.lxcf_entry.configure(
            state="normal" if lxcf_enabled else "disabled"
        )
        self.lxcf_button.configure(
            state="normal" if lxcf_enabled else "disabled"
        )

        self._update_badge()

    def _update_badge(self):
        if self.running:
            self.header_badge.configure(
                text="В РАБОТЕ",
                bg=self.ACCENT,
            )
        elif self.errors:
            self.header_badge.configure(
                text="ЕСТЬ ОШИБКИ",
                bg=self.ERROR,
            )
        else:
            self.header_badge.configure(
                text="ГОТОВО",
                bg=self.SUCCESS,
            )

    def pick_source(self):
        p = filedialog.askdirectory(
            title="Выберите папку с M3D"
        )
        if p:
            self.source.set(p)

    def pick_destination(self):
        p = filedialog.askdirectory(
            title="Выберите папку назначения"
        )
        if p:
            self.destination.set(p)

    def pick_sxcf(self):
        p = filedialog.askopenfilename(
            title="Выберите конфигурацию STEP",
            filetypes=[
                ("STEP configuration", "*.sxcf"),
                ("Все файлы", "*.*"),
            ],
        )
        if p:
            self.sxcf.set(p)

    def pick_lxcf(self):
        p = filedialog.askopenfilename(
            title="Выберите конфигурацию STL",
            filetypes=[
                ("STL configuration", "*.lxcf"),
                ("Все файлы", "*.*"),
            ],
        )
        if p:
            self.lxcf.set(p)
            self.use_lxcf.set(True)
            self.update_controls()

    def clear_log(self):
        self.logbox.configure(state="normal")
        self.logbox.delete("1.0", "end")
        self.logbox.configure(state="disabled")

    def write_log(self, text, tag=None):
        def put():
            self.logbox.configure(state="normal")
            if tag:
                self.logbox.insert("end", text + "\n", tag)
            else:
                self.logbox.insert("end", text + "\n")
            self.logbox.see("end")
            self.logbox.configure(state="disabled")
        self.after(0, put)

    def set_status(self, text):
        self.after(
            0, lambda: self.status.set(text)
        )

    def set_progress(self, value):
        self.after(
            0, lambda: self.progress.configure(value=value)
        )

    def update_stats(self):
        text = (
            f"{self.total} файлов  •  "
            f"{self.ok} успешно  •  "
            f"{self.skipped} пропущено  •  "
            f"{self.errors} ошибок"
        )
        self.after(
            0, lambda: self.stats.configure(text=text)
        )

    def start_conversion(self):
        if self.running:
            return

        src = Path(self.source.get().strip())
        dst = Path(self.destination.get().strip())

        if not src.is_dir():
            messagebox.showerror(
                "Ошибка",
                "Выберите существующую исходную папку.",
            )
            return

        dst.mkdir(
            parents=True, exist_ok=True
        )

        # CDW -> PDF режим
        if self.operation.get() == "CDW":
            pattern = (
                "**/*.cdw"
                if self.recursive.get()
                else "*.cdw"
            )
            files = sorted(
                [p for p in src.glob(pattern) if p.is_file()],
                key=lambda p: str(p).lower(),
            )

            if not files:
                messagebox.showinfo(
                    "Нет файлов",
                    "В выбранной папке нет файлов .cdw.",
                )
                return

            self.running = True
            self.total = len(files)
            self.ok = 0
            self.skipped = 0
            self.errors = 0

            self.start_button.configure(
                state="disabled"
            )
            self.progress.configure(
                maximum=len(files), value=0
            )
            self.clear_log()
            self._update_badge()
            self.update_stats()

            self.write_log(
                "M3D Converter • КОМПАС-3D v24",
                "header",
            )
            self.write_log(
                "Режим: CDW → PDF • Pdf2d.dll • без конфигураций",
                "header",
            )
            self.write_log(
                f"Файлов: {len(files)}"
            )
            self.write_log("=" * 72)

            threading.Thread(
                target=self.worker_cdw,
                args=(
                    src,
                    dst,
                    files,
                    self.overwrite.get(),
                ),
                daemon=True,
            ).start()
            return

        fmt = self.format.get()
        config = None

        if fmt == "STEP":
            step_map = {
                "AP203": FORMAT_STEP_AP203,
                "AP214": FORMAT_STEP_AP214,
                "AP242": FORMAT_STEP_AP242,
            }
            step_format = step_map[self.step.get()]

            cfg = self.sxcf.get().strip()
            if cfg:
                p = Path(cfg)
                if (
                    not p.is_file()
                    or p.suffix.lower() != ".sxcf"
                ):
                    messagebox.showerror(
                        "Ошибка",
                        "Укажите существующий файл .sxcf.",
                    )
                    return
                config = str(p)

        else:
            step_format = FORMAT_STEP

            if self.use_lxcf.get():
                cfg = self.lxcf.get().strip()
                if not cfg:
                    messagebox.showerror(
                        "Ошибка",
                        "Выбрано использование .lxcf, "
                        "но файл не указан.",
                    )
                    return

                p = Path(cfg)
                if (
                    not p.is_file()
                    or p.suffix.lower() != ".lxcf"
                ):
                    messagebox.showerror(
                        "Ошибка",
                        "Укажите существующий файл .lxcf.",
                    )
                    return
                config = str(p)

        pattern = (
            "**/*.m3d"
            if self.recursive.get()
            else "*.m3d"
        )
        files = sorted(
            [p for p in src.glob(pattern) if p.is_file()],
            key=lambda p: str(p).lower(),
        )

        if not files:
            messagebox.showinfo(
                "Нет файлов",
                "В выбранной папке нет файлов .m3d.",
            )
            return

        self.running = True
        self.total = len(files)
        self.ok = 0
        self.skipped = 0
        self.errors = 0

        self.start_button.configure(
            state="disabled"
        )
        self.progress.configure(
            maximum=len(files), value=0
        )
        self.clear_log()
        self._update_badge()
        self.update_stats()

        self.write_log(
            "M3D Converter • КОМПАС-3D v24",
            "header",
        )
        self.write_log(
            "API7 Open() → API5 ActiveDocument3D() → экспорт",
            "header",
        )
        self.write_log(
            f"Формат: {fmt} | Файлов: {len(files)}"
        )
        self.write_log("=" * 72)

        threading.Thread(
            target=self.worker,
            args=(
                src,
                dst,
                files,
                fmt,
                step_format,
                config,
                self.overwrite.get(),
            ),
            daemon=True,
        ).start()

    def worker_cdw(
        self,
        src,
        dst,
        files,
        overwrite,
    ):
        conv = KompasConverter(self.write_log)

        try:
            conv.connect()

            for i, source in enumerate(files, 1):
                self.set_status(
                    f"Конвертация CDW {i} из {len(files)} • {source.name}"
                )
                self.write_log(
                    f"[{i}/{len(files)}] {source}"
                )

                if self.recursive.get():
                    rel = source.parent.relative_to(src)
                    target_dir = dst / rel
                else:
                    target_dir = dst

                target_dir.mkdir(
                    parents=True, exist_ok=True
                )

                try:
                    result, target = conv.convert_cdw_to_pdf(
                        source,
                        target_dir,
                        overwrite,
                    )

                    if result == "ok":
                        self.ok += 1
                        self.write_log(
                            f"    ✓ OK → {target}",
                            "ok",
                        )
                    else:
                        self.skipped += 1
                        self.write_log(
                            f"    ↷ SKIP → {target}",
                            "warn",
                        )

                except Exception as exc:
                    self.errors += 1
                    self.write_log(
                        f"    ✕ ОШИБКА: {exc}",
                        "error",
                    )
                    self.write_log(
                        traceback.format_exc(),
                        "error",
                    )

                self.set_progress(i)
                self.update_stats()

            self.write_log("=" * 72)
            final = (
                f"ГОТОВО  •  {self.ok} успешно  •  "
                f"{self.skipped} пропущено  •  "
                f"{self.errors} ошибок"
            )
            self.write_log(
                final,
                "error" if self.errors else "ok",
            )
            self.set_status(final)

            self.after(
                0,
                lambda: messagebox.showinfo(
                    "CDW → PDF завершено",
                    f"Всего: {self.total}\n"
                    f"Успешно: {self.ok}\n"
                    f"Пропущено: {self.skipped}\n"
                    f"Ошибок: {self.errors}",
                ),
            )

        except Exception as exc:
            self.write_log("=" * 72, "error")
            self.write_log(
                f"КРИТИЧЕСКАЯ ОШИБКА: {exc}",
                "error",
            )
            self.write_log(
                traceback.format_exc(),
                "error",
            )
            self.set_status(
                "Ошибка подключения к КОМПАС-3D."
            )
            self.after(
                0,
                lambda: messagebox.showerror(
                    "Ошибка",
                    str(exc),
                ),
            )

        finally:
            conv.disconnect()
            self.running = False
            self.after(
                0,
                lambda: self.start_button.configure(
                    state="normal"
                )
            )
            self.after(0, self._update_badge)

    def worker(
        self,
        src,
        dst,
        files,
        fmt,
        step_format,
        config,
        overwrite,
    ):
        conv = KompasConverter(self.write_log)

        try:
            conv.connect()

            for i, source in enumerate(files, 1):
                self.set_status(
                    f"Обработка {i} из {len(files)} • {source.name}"
                )
                self.write_log(
                    f"[{i}/{len(files)}] {source}"
                )

                if self.recursive.get():
                    rel = source.parent.relative_to(src)
                    target_dir = dst / rel
                else:
                    target_dir = dst

                target_dir.mkdir(
                    parents=True, exist_ok=True
                )

                try:
                    result, target = conv.convert_one(
                        source,
                        target_dir,
                        fmt,
                        step_format,
                        config,
                        overwrite,
                    )

                    if result == "ok":
                        self.ok += 1
                        self.write_log(
                            f"    ✓ OK → {target}",
                            "ok",
                        )
                    else:
                        self.skipped += 1
                        self.write_log(
                            f"    ↷ SKIP → {target}",
                            "warn",
                        )

                except Exception as exc:
                    self.errors += 1
                    self.write_log(
                        f"    ✕ ОШИБКА: {exc}",
                        "error",
                    )
                    self.write_log(
                        traceback.format_exc(),
                        "error",
                    )

                self.set_progress(i)
                self.update_stats()

            self.write_log("=" * 72)
            final = (
                f"ГОТОВО  •  {self.ok} успешно  •  "
                f"{self.skipped} пропущено  •  "
                f"{self.errors} ошибок"
            )
            self.write_log(
                final,
                "error" if self.errors else "ok",
            )
            self.set_status(final)

            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Конвертация завершена",
                    f"Всего: {self.total}\n"
                    f"Успешно: {self.ok}\n"
                    f"Пропущено: {self.skipped}\n"
                    f"Ошибок: {self.errors}",
                ),
            )

        except Exception as exc:
            self.write_log(
                "=" * 72,
                "error",
            )
            self.write_log(
                f"КРИТИЧЕСКАЯ ОШИБКА: {exc}",
                "error",
            )
            self.write_log(
                traceback.format_exc(),
                "error",
            )
            self.set_status(
                "Ошибка подключения к КОМПАС-3D."
            )

            self.after(
                0,
                lambda: messagebox.showerror(
                    "Ошибка",
                    str(exc),
                ),
            )

        finally:
            conv.disconnect()
            self.running = False
            self.after(
                0,
                lambda: self.start_button.configure(
                    state="normal"
                ),
            )
            self.after(
                0, self._update_badge
            )


if __name__ == "__main__":
    App().mainloop()
