"""
Однофайловое GUI-приложение для пакетного геокодирования.

Можно передавать этот файл как есть или собрать из него один EXE:
    pyinstaller --onefile --windowed --name GeocodeEXE geocode_exe.py

Зависимости для запуска .py:
- Python 3.10+
- tkinter (обычно входит в Python для Windows)
- openpyxl нужен только для чтения/записи Excel .xlsx/.xlsm

CSV/TXT и запросы к сервису работают только на стандартной библиотеке Python.
"""

from __future__ import annotations

import csv
import importlib
import importlib.util
import json
import ssl
import threading
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

openpyxl = importlib.import_module("openpyxl") if importlib.util.find_spec("openpyxl") else None

DEFAULT_FORWARD_URL = "https://dadatatest.t2.ru/suggestions/ui/service/api-proxy/address/suggest"
DEFAULT_REVERSE_URL = "https://dadatatest.t2.ru/suggestions/ui/service/api-proxy/address/geolocate"
RESULT_ADDRESS_COLUMN = "Найденный адрес"
RESULT_LAT_COLUMN = "Широта результата"
RESULT_LON_COLUMN = "Долгота результата"
FILE_TYPES = [
    ("Табличные файлы", "*.xlsx *.xlsm *.csv *.txt"),
    ("Excel", "*.xlsx *.xlsm"),
    ("CSV", "*.csv"),
    ("TXT", "*.txt"),
    ("Все файлы", "*.*"),
]


class GeocodingError(RuntimeError):
    """Ошибка запроса к сервису геокодирования."""


@dataclass(slots=True)
class TableData:
    headers: list[str]
    rows: list[dict[str, Any]]

    def copy(self) -> "TableData":
        return TableData(self.headers[:], [row.copy() for row in self.rows])


@dataclass(slots=True)
class GeocodingClient:
    forward_url: str = DEFAULT_FORWARD_URL
    reverse_url: str = DEFAULT_REVERSE_URL
    timeout: float = 20.0
    verify_ssl: bool = False

    def geocode_address(self, address: str) -> dict[str, str]:
        address = (address or "").strip()
        if not address:
            return {RESULT_ADDRESS_COLUMN: "", RESULT_LAT_COLUMN: "", RESULT_LON_COLUMN: ""}

        data = self._post(self.forward_url, {"query": address})
        suggestions = data.get("suggestions") or []
        if not suggestions:
            return {RESULT_ADDRESS_COLUMN: "", RESULT_LAT_COLUMN: "", RESULT_LON_COLUMN: ""}

        first = suggestions[0]
        details = first.get("data") or {}
        return {
            RESULT_ADDRESS_COLUMN: first.get("unrestricted_value") or first.get("value") or "",
            RESULT_LAT_COLUMN: str(details.get("geo_lat") or ""),
            RESULT_LON_COLUMN: str(details.get("geo_lon") or ""),
        }

    def reverse_geocode(self, lat: Any, lon: Any) -> dict[str, str]:
        lat_text = str(lat or "").strip().replace(",", ".")
        lon_text = str(lon or "").strip().replace(",", ".")
        if not lat_text or not lon_text:
            return {RESULT_ADDRESS_COLUMN: "", RESULT_LAT_COLUMN: lat_text, RESULT_LON_COLUMN: lon_text}

        data = self._post(self.reverse_url, {"lat": lat_text, "lon": lon_text})
        suggestions = data.get("suggestions") or []
        if not suggestions:
            return {RESULT_ADDRESS_COLUMN: "", RESULT_LAT_COLUMN: lat_text, RESULT_LON_COLUMN: lon_text}

        first = suggestions[0]
        return {
            RESULT_ADDRESS_COLUMN: first.get("unrestricted_value") or first.get("value") or "",
            RESULT_LAT_COLUMN: lat_text,
            RESULT_LON_COLUMN: lon_text,
        }

    def _post(self, url: str, request_payload: dict[str, str]) -> dict[str, Any]:
        payload = urllib.parse.urlencode({"apiRequest": json.dumps(request_payload, ensure_ascii=False)}).encode("utf-8")
        request = urllib.request.Request(
            url,
            data=payload,
            method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded; charset=utf-8"},
        )
        context = None if self.verify_ssl else ssl._create_unverified_context()
        try:
            with urllib.request.urlopen(request, timeout=self.timeout, context=context) as response:
                text = response.read().decode("utf-8")
        except urllib.error.URLError as exc:
            raise GeocodingError(f"Ошибка запроса к сервису: {exc}") from exc
        except TimeoutError as exc:
            raise GeocodingError("Превышено время ожидания сервиса") from exc

        try:
            parsed = json.loads(text)
        except json.JSONDecodeError as exc:
            raise GeocodingError("Сервис вернул ответ не в формате JSON") from exc
        if not isinstance(parsed, dict):
            raise GeocodingError("Сервис вернул неожиданный формат ответа")
        return parsed


def read_table(
    path: str | Path,
    *,
    sheet_name: str = "",
    delimiter: str | None = None,
    encoding: str = "utf-8-sig",
    has_header: bool = True,
    start_row: int = 1,
) -> TableData:
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return read_excel(file_path, sheet_name=sheet_name, has_header=has_header, start_row=start_row)
    if suffix in {".csv", ".txt"}:
        return read_delimited(file_path, delimiter=delimiter, encoding=encoding, has_header=has_header, start_row=start_row)
    raise ValueError(f"Неподдерживаемый формат файла: {suffix}")


def read_delimited(
    path: Path,
    *,
    delimiter: str | None = None,
    encoding: str = "utf-8-sig",
    has_header: bool = True,
    start_row: int = 1,
) -> TableData:
    raw = path.read_text(encoding=encoding, errors="replace")
    lines = raw.splitlines()[max(start_row - 1, 0) :]
    if not lines:
        return TableData(headers=[], rows=[])
    if delimiter:
        reader = csv.reader(lines, delimiter=delimiter)
    else:
        sample = "\n".join(lines[:25])
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel_tab if path.suffix.lower() == ".txt" else csv.excel
        reader = csv.reader(lines, dialect=dialect)
    values = list(reader)
    if not values:
        return TableData(headers=[], rows=[])
    if has_header:
        headers = [str(value or "") for value in values[0]]
        data_rows = values[1:]
    else:
        width = max(len(row) for row in values)
        headers = [f"Столбец {index}" for index in range(1, width + 1)]
        data_rows = values
    rows = []
    for source_row in data_rows:
        row = {header: source_row[index] if index < len(source_row) else "" for index, header in enumerate(headers)}
        rows.append(row)
    return TableData(headers=headers, rows=rows)


def read_excel(path: Path, *, sheet_name: str = "", has_header: bool = True, start_row: int = 1) -> TableData:
    if openpyxl is None:
        raise ValueError("Для Excel нужен пакет openpyxl. Для EXE он должен быть установлен перед сборкой.")
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook[sheet_name] if sheet_name and sheet_name in workbook.sheetnames else workbook.active
    values = list(sheet.iter_rows(values_only=True))[max(start_row - 1, 0) :]
    if not values:
        return TableData(headers=[], rows=[])
    if has_header:
        headers = [str(value or "") for value in values[0]]
        data_rows = values[1:]
    else:
        width = max(len(row) for row in values)
        headers = [f"Столбец {index}" for index in range(1, width + 1)]
        data_rows = values
    rows = []
    for source_row in data_rows:
        row = {header: source_row[index] if index < len(source_row) else "" for index, header in enumerate(headers)}
        rows.append(row)
    return TableData(headers=headers, rows=rows)


def write_table(table: TableData, path: str | Path) -> None:
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        write_excel(table, file_path)
        return
    if suffix == ".csv":
        write_delimited(table, file_path, delimiter=",")
        return
    if suffix == ".txt":
        write_delimited(table, file_path, delimiter="\t")
        return
    raise ValueError(f"Неподдерживаемый формат сохранения: {suffix}")


def write_delimited(table: TableData, path: Path, delimiter: str) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=table.headers, delimiter=delimiter, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(table.rows)


def write_excel(table: TableData, path: Path) -> None:
    if openpyxl is None:
        raise ValueError("Для сохранения Excel нужен пакет openpyxl. Сохраните в CSV/TXT или соберите EXE с openpyxl.")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Result"
    sheet.append(table.headers)
    for row in table.rows:
        sheet.append([row.get(header, "") for header in table.headers])
    workbook.save(path)


def append_result_columns(table: TableData) -> TableData:
    result = table.copy()
    for column in [RESULT_ADDRESS_COLUMN, RESULT_LAT_COLUMN, RESULT_LON_COLUMN]:
        if column not in result.headers:
            result.headers.append(column)
    return result


def process_addresses(
    table: TableData,
    address_column: str,
    client: GeocodingClient,
    progress: Callable[[int, int, str], None] | None = None,
) -> TableData:
    result = append_result_columns(table)
    total = len(result.rows)
    for index, row in enumerate(result.rows, start=1):
        source = str(row.get(address_column, ""))
        row.update(client.geocode_address(source))
        if progress:
            progress(index, total, source)
    return result


def process_coordinates(
    table: TableData,
    lat_column: str,
    lon_column: str,
    client: GeocodingClient,
    progress: Callable[[int, int, str], None] | None = None,
) -> TableData:
    result = append_result_columns(table)
    total = len(result.rows)
    for index, row in enumerate(result.rows, start=1):
        lat = row.get(lat_column, "")
        lon = row.get(lon_column, "")
        row.update(client.reverse_geocode(lat, lon))
        if progress:
            progress(index, total, f"{lat}, {lon}")
    return result


class RoundedField(tk.Frame):
    """Контейнер с мягкой скруглённой обводкой для полей выбора."""

    def __init__(self, parent: tk.Widget, *, active: bool = False) -> None:
        super().__init__(parent, bg="#222846", highlightthickness=0, bd=0)
        self._active = active
        self._canvas = tk.Canvas(self, height=48, bg="#222846", highlightthickness=0, bd=0)
        self._canvas.pack(fill="both", expand=True)
        self.content = tk.Frame(self._canvas, bg="#f3f4f6", bd=0, highlightthickness=0)
        self._window = self._canvas.create_window(6, 6, anchor="nw", window=self.content)
        self._canvas.bind("<Configure>", self._redraw)

    def set_active(self, active: bool) -> None:
        self._active = active
        self._redraw()

    def _redraw(self, _event: tk.Event | None = None) -> None:
        width = max(self._canvas.winfo_width(), 1)
        height = max(self._canvas.winfo_height(), 48)
        border = "#f02fb3" if self._active else "#4b5275"
        self._canvas.delete("border")
        self._round_rect(2, 2, width - 2, height - 2, radius=14, fill=border, tags="border")
        self._round_rect(5, 5, width - 5, height - 5, radius=11, fill="#f3f4f6", tags="border")
        self._canvas.coords(self._window, 8, 8)
        self._canvas.itemconfigure(self._window, width=max(width - 16, 1), height=max(height - 16, 1))

    def _round_rect(self, x1: int, y1: int, x2: int, y2: int, *, radius: int, **kwargs: Any) -> None:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        self._canvas.create_polygon(points, smooth=True, **kwargs)


class RoundedRadio(tk.Frame):
    """Радиокнопка с закруглённой цветной подложкой выбранного состояния."""

    def __init__(
        self,
        parent: tk.Widget,
        *,
        text: str,
        variable: tk.Variable,
        value: str,
        command: Callable[[], None] | None = None,
        bg: str = "#222846",
    ) -> None:
        super().__init__(parent, bg=bg, highlightthickness=0, bd=0)
        self.variable = variable
        self.value = value
        self.bg = bg
        self.accent = "#f02fb3"
        self._canvas = tk.Canvas(self, height=34, bg=bg, highlightthickness=0, bd=0)
        self._canvas.pack(fill="both", expand=True)
        self.button = tk.Radiobutton(
            self._canvas,
            text=text,
            variable=variable,
            value=value,
            command=command,
            bg=bg,
            activebackground="#343b62",
            fg="#ffffff",
            activeforeground="#ffffff",
            selectcolor="#ffffff",
            font=("Segoe UI", 9, "bold"),
            indicatoron=True,
            borderwidth=0,
            highlightthickness=0,
            padx=8,
            pady=4,
        )
        self._canvas.configure(width=self.button.winfo_reqwidth() + 16)
        self._window = self._canvas.create_window(8, 5, anchor="nw", window=self.button)
        self._canvas.bind("<Configure>", self._redraw)
        self.variable.trace_add("write", lambda *_args: self._redraw())
        self._redraw()

    def _redraw(self, _event: tk.Event | None = None) -> None:
        selected = self.variable.get() == self.value
        width = max(self._canvas.winfo_width(), self.button.winfo_reqwidth() + 16)
        height = max(self._canvas.winfo_height(), 34)
        fill = self.accent if selected else self.bg
        foreground = "#171a2e" if selected else "#ffffff"
        active_bg = self.accent if selected else "#343b62"
        self._canvas.delete("background")
        self._round_rect(1, 1, width - 1, height - 1, radius=12, fill=fill, tags="background")
        self._canvas.tag_lower("background")
        self._canvas.itemconfigure(self._window, width=max(width - 16, 1), height=max(height - 10, 1))
        self.button.configure(bg=fill, activebackground=active_bg, fg=foreground)

    def _round_rect(self, x1: int, y1: int, x2: int, y2: int, *, radius: int, **kwargs: Any) -> None:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        self._canvas.create_polygon(points, smooth=True, **kwargs)


class RoundedButton(tk.Frame):
    """Кнопка с розовой скруглённой подложкой в стиле полей с обводкой."""

    def __init__(
        self,
        parent: tk.Widget,
        *,
        text: str,
        command: Callable[[], None],
        bg: str = "#222846",
        fill: str = "#f02fb3",
        active_fill: str = "#ff5ac8",
        disabled_fill: str = "#6f3a64",
        foreground: str = "#ffffff",
        padx: int = 18,
        pady: int = 10,
        height: int = 44,
    ) -> None:
        super().__init__(parent, bg=bg, highlightthickness=0, bd=0)
        self.command = command
        self.text = text
        self.fill = fill
        self.active_fill = active_fill
        self.disabled_fill = disabled_fill
        self.foreground = foreground
        self._state = "normal"
        self._hovered = False
        self._padx = padx
        self._pady = pady
        self._min_width = len(self.text) * 9 + self._padx * 2
        self._height = height
        self._canvas = tk.Canvas(self, width=self._min_width, height=self._height, bg=bg, highlightthickness=0, bd=0, cursor="hand2")
        self._canvas.pack(fill="both", expand=True)
        self._canvas.bind("<Configure>", self._redraw)
        self._canvas.bind("<Enter>", self._on_enter)
        self._canvas.bind("<Leave>", self._on_leave)
        self._canvas.bind("<Button-1>", self._on_click)
        self._redraw()

    def configure(self, cnf: dict[str, Any] | None = None, **kwargs: Any) -> None:  # type: ignore[override]
        options = dict(cnf or {}, **kwargs)
        if "state" in options:
            self._state = str(options.pop("state"))
            self._canvas.configure(cursor="" if self._state == "disabled" else "hand2")
        if "text" in options:
            self.text = str(options.pop("text"))
            self._min_width = len(self.text) * 9 + self._padx * 2
            self._canvas.configure(width=self._min_width)
        if "command" in options:
            self.command = options.pop("command")
        if options:
            super().configure(**options)
        self._redraw()

    config = configure

    def _on_enter(self, _event: tk.Event) -> None:
        self._hovered = True
        self._redraw()

    def _on_leave(self, _event: tk.Event) -> None:
        self._hovered = False
        self._redraw()

    def _on_click(self, _event: tk.Event) -> None:
        if self._state != "disabled":
            self.command()

    def _redraw(self, _event: tk.Event | None = None) -> None:
        width = max(self._canvas.winfo_width(), self._min_width)
        height = max(self._canvas.winfo_height(), self._height)
        fill = self.disabled_fill if self._state == "disabled" else self.active_fill if self._hovered else self.fill
        self._canvas.delete("all")
        self._round_rect(1, 1, width - 1, height - 1, radius=14, fill=fill, outline="")
        self._canvas.create_text(
            width // 2,
            height // 2,
            text=self.text,
            fill=self.foreground,
            font=("Segoe UI", 10, "bold"),
        )

    def _round_rect(self, x1: int, y1: int, x2: int, y2: int, *, radius: int, **kwargs: Any) -> None:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        self._canvas.create_polygon(points, smooth=True, **kwargs)


class CheckColumnDropdown(ttk.Frame):
    """Выпадающий список с чекбоксами для выбора рабочих столбцов."""

    def __init__(self, parent: tk.Widget, textvariable: tk.StringVar, command: Callable[[], None] | None = None) -> None:
        super().__init__(parent)
        self.textvariable = textvariable
        self.command = command
        self.columns: list[str] = []
        self.column_vars: dict[str, tk.BooleanVar] = {}
        self.all_var = tk.BooleanVar(value=True)
        self.popup: tk.Toplevel | None = None
        self._updating = False

        self.entry = ttk.Entry(self, textvariable=self.textvariable, state="readonly")
        self.entry.pack(side="left", fill="x", expand=True)
        self.button = ttk.Button(self, text="▼", width=2, command=self.toggle_popup, style="DropdownArrow.TButton")
        self.button.pack(side="right", fill="y")
        self.entry.bind("<Button-1>", lambda _event: self.toggle_popup())

    def set_columns(self, columns: list[str]) -> None:
        self.columns = columns[:]
        self.column_vars = {column: tk.BooleanVar(value=True) for column in self.columns}
        self.all_var.set(True)
        self._refresh_text()
        if self.popup is not None and self.popup.winfo_exists():
            self._rebuild_popup()

    def selected_columns(self) -> list[str]:
        if self.all_var.get():
            return self.columns[:]
        return [column for column in self.columns if self.column_vars[column].get()]

    def toggle_popup(self) -> None:
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.destroy()
            self.popup = None
            return
        self.popup = tk.Toplevel(self)
        self.popup.overrideredirect(True)
        self.popup.transient(self.winfo_toplevel())
        self.popup.configure(bg="#f3f4f6", padx=2, pady=2)
        self._rebuild_popup()
        self._position_popup()
        self.popup.bind("<FocusOut>", self._on_popup_focus_out)
        self.popup.bind("<Escape>", lambda _event: self._close_popup())
        self.popup.update_idletasks()
        width = max(self.winfo_width(), 260)
        height = self.popup.winfo_reqheight()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        self.popup.geometry(f"{width}x{height}{x:+d}{y:+d}")
        self.popup.bind("<FocusOut>", lambda _event: self._close_popup())
        self.popup.focus_force()

    def _rebuild_popup(self) -> None:
        if self.popup is None:
            return
        for child in self.popup.winfo_children():
            child.destroy()
        all_check = ttk.Checkbutton(self.popup, text="Выбрать все", variable=self.all_var, command=self._on_all_changed, style="Dropdown.TCheckbutton")
        all_check.pack(anchor="w", fill="x")
        for column in self.columns:
            check = ttk.Checkbutton(
                self.popup,
                text=column,
                variable=self.column_vars[column],
                command=self._on_column_changed,
                style="Dropdown.TCheckbutton",
            )
            check.pack(anchor="w", fill="x")

    def _position_popup(self) -> None:
        if self.popup is None:
            return
        self.popup.update_idletasks()
        width = max(self.winfo_width(), 260)
        height = self.popup.winfo_reqheight()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        self.popup.geometry(f"{width}x{height}{x:+d}{y:+d}")

    def _on_popup_focus_out(self, _event: tk.Event) -> None:
        if self.popup is None:
            return
        self.after_idle(self._close_popup_if_unfocused)

    def _close_popup_if_unfocused(self) -> None:
        if self.popup is None or not self.popup.winfo_exists():
            return
        focused = self.focus_get()
        if focused is None or not str(focused).startswith(str(self.popup)):
            self._close_popup()

    def _close_popup(self) -> None:
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.destroy()
        self.popup = None

    def _on_all_changed(self) -> None:
        if self._updating:
            return
        self._updating = True
        try:
            if self.all_var.get():
                for variable in self.column_vars.values():
                    variable.set(True)
            elif self.columns:
                for variable in self.column_vars.values():
                    variable.set(False)
        finally:
            self._updating = False
        self._refresh_text()
        self._rebuild_popup()
        self._position_popup()
        if self.command:
            self.command()

    def _on_column_changed(self) -> None:
        if self._updating:
            return
        selected_count = sum(variable.get() for variable in self.column_vars.values())
        self.all_var.set(selected_count == len(self.columns) and bool(self.columns))
        self._refresh_text()
        if self.command:
            self.command()

    def _refresh_text(self) -> None:
        total = len(self.columns)
        selected = total if self.all_var.get() else len(self.selected_columns())
        self.textvariable.set(f"Выбрано {selected} из {total}")


class GeocodeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Геокодирование файлов")
        self.geometry("1040x720")
        self.minsize(940, 640)

        self.table_data: TableData | None = None
        self.result_data: TableData | None = None
        self.loaded_path: Path | None = None
        self.worker_events: list[tuple[str, Any]] = []
        self.worker_lock = threading.Lock()

        self.mode = tk.StringVar(value="address_to_coords")
        self.address_column = tk.StringVar()
        self.lat_column = tk.StringVar()
        self.lon_column = tk.StringVar()
        self.source_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.work_columns = tk.StringVar(value="Все")
        self.delimiter = tk.StringVar(value=";")
        self.has_header = tk.BooleanVar(value=True)
        self.start_row = tk.IntVar(value=1)
        self.encoding = tk.StringVar(value="utf-8-sig")
        self.status = tk.StringVar(value="Загрузите Excel, CSV или TXT файл")

        self._configure_style()
        self._build_ui()
        self.after(150, self._poll_worker_events)

    def _configure_style(self) -> None:
        self.configure(bg="#171a2e")
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        base = "#171a2e"
        card = "#222846"
        field = "#f3f4f6"
        accent = "#f02fb3"
        cyan = "#38d6e8"
        style.configure("TFrame", background=base)
        style.configure("Card.TFrame", background=card, relief="flat")
        style.configure("FieldWrap.TFrame", background=card)
        style.configure("ActiveFieldWrap.TFrame", background=accent)
        style.configure("Status.TFrame", background="#1d2340", relief="flat")
        style.configure("Status.TLabel", background="#1d2340", foreground="#dbeafe", font=("Segoe UI", 9, "bold"))
        style.configure("TLabel", background=base, foreground="#eef2ff", font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=card, foreground="#f8fafc", font=("Segoe UI", 10, "bold"))
        style.configure("Muted.TLabel", foreground="#b8bfd9", background=card)
        style.configure("Title.TLabel", background=base, foreground="#ffffff", font=("Segoe UI", 22, "bold"))
        style.configure("Subtitle.TLabel", background=base, foreground="#aab4d6", font=("Segoe UI", 10))
        style.configure("TLabelframe", background=card, bordercolor=card, lightcolor=card, darkcolor=card, relief="flat")
        style.configure("TLabelframe.Label", background=card, foreground="#ffffff", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=(18, 10), background="#343b62", foreground="#ffffff", borderwidth=0, relief="flat")
        style.map("TButton", background=[("active", "#4a527f"), ("disabled", "#2d334f")])
        style.configure("Accent.TButton", foreground="#ffffff", background=accent)
        style.map("Accent.TButton", background=[("active", "#ff5ac8"), ("disabled", "#6f3a64")])
        style.configure("Tool.TButton", foreground="#ffffff", background="#4a4f66", padding=(12, 8), relief="flat")
        style.configure("DropdownArrow.TButton", foreground="#ffffff", background="#4a4f66", padding=(2, 0), relief="flat")
        style.configure("TEntry", fieldbackground=field, foreground="#121827", bordercolor="#4b5275", lightcolor=cyan, darkcolor="#4b5275", padding=7, relief="flat")
        style.configure("TSpinbox", fieldbackground=field, foreground="#121827", arrowsize=12)
        style.configure("TCheckbutton", background=card, foreground="#ffffff", font=("Segoe UI", 9, "bold"))
        style.configure("Dropdown.TCheckbutton", background=field, foreground="#121827", font=("Segoe UI", 9))
        style.map("Dropdown.TCheckbutton", background=[("active", "#e5e7eb"), ("disabled", field)], foreground=[("disabled", "#6b7280")])
        style.configure("TRadiobutton", background=card, foreground="#ffffff", font=("Segoe UI", 9, "bold"), padding=(10, 6))
        style.map("TRadiobutton", background=[("selected", "#f02fb3"), ("active", "#343b62")], foreground=[("selected", "#171a2e"), ("active", "#ffffff")])
        style.configure("TCombobox", padding=7, fieldbackground=field, foreground="#121827", arrowcolor="#4b5275", relief="flat")
        style.map("TCombobox", fieldbackground=[("disabled", "#313650"), ("readonly", field)], foreground=[("disabled", "#8c93ae"), ("readonly", "#121827")])
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 9), background="#202540", fieldbackground="#202540", foreground="#eef2ff", borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#2d3457", foreground="#ffffff")
        style.configure("Horizontal.TProgressbar", troughcolor="#111827", background=accent, bordercolor="#1d2340", lightcolor="#ff5ac8", darkcolor=accent)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=20)
        root.pack(fill="both", expand=True)

        hero = tk.Canvas(root, height=88, highlightthickness=0, bg="#171a2e")
        hero.pack(fill="x", pady=(0, 12))
        hero.create_arc(-80, 16, 330, 180, start=10, extent=120, outline="#38d6e8", width=3, style="arc")
        hero.create_arc(650, -90, 1160, 160, start=190, extent=130, outline="#f02fb3", width=4, style="arc")
        hero.create_text(8, 26, anchor="w", text="Геокодирование файлов", fill="#ffffff", font=("Segoe UI", 22, "bold"))
        hero.create_text(10, 58, anchor="w", text="Настройте входной файл, выберите колонки и сохраните результат.", fill="#aab4d6", font=("Segoe UI", 10))

        io = self._make_section(root, "Исходный файл/Настройка вывода", fill="x")
        io.columnconfigure(1, weight=1)
        io.columnconfigure(4, weight=1)

        ttk.Label(io, text="Исходный файл (txt, csv, excel):", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(io, textvariable=self.source_file).grid(row=0, column=1, columnspan=4, sticky="ew", padx=(6, 6), pady=4)
        RoundedButton(io, text="Выбрать", command=self.open_file, bg="#222846", padx=12, pady=8, height=34).grid(row=0, column=5, sticky="ew", pady=4)

        ttk.Label(io, text="Выходной файл:", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(io, textvariable=self.output_file).grid(row=1, column=1, columnspan=4, sticky="ew", padx=(6, 6), pady=4)
        RoundedButton(io, text="Выбрать", command=self.choose_output_file, bg="#222846", padx=12, pady=8, height=34).grid(row=1, column=5, sticky="ew", pady=4)

        ttk.Label(io, text="Лист Excel:", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        self.sheet_combo = ttk.Combobox(io, textvariable=self.sheet_name, state="readonly", values=[])
        self.sheet_combo.grid(row=2, column=1, columnspan=5, sticky="ew", padx=(6, 0), pady=4)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self.reload_file())

        ttk.Label(io, text="Столбцы для работы:", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=4)
        self.columns_dropdown = CheckColumnDropdown(io, textvariable=self.work_columns, command=self._on_work_columns_changed)
        self.columns_dropdown.grid(row=3, column=1, columnspan=5, sticky="ew", padx=(6, 0), pady=4)

        ttk.Label(io, text="Разделитель:", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=4)
        self.delimiter_entry = ttk.Entry(io, textvariable=self.delimiter, width=4)
        self.delimiter_entry.grid(row=4, column=1, sticky="w", padx=(6, 6), pady=4)
        ttk.Checkbutton(io, text="Наличие заголовка", variable=self.has_header, command=self.reload_file).grid(row=4, column=2, sticky="w", pady=4)
        ttk.Label(io, text="Данные начинаются со строки:", style="Card.TLabel").grid(row=4, column=3, sticky="e", padx=(12, 6), pady=4)
        ttk.Spinbox(io, from_=1, to=9999, textvariable=self.start_row, width=6, command=self.reload_file).grid(row=4, column=4, sticky="w", pady=4)

        ttk.Label(io, text="Кодировка файла:", style="Card.TLabel").grid(row=5, column=0, sticky="w", pady=4)
        self.encoding_radios = [
            RoundedRadio(io, text="UTF-8", variable=self.encoding, value="utf-8-sig", command=self.reload_file),
            RoundedRadio(io, text="CP1251", variable=self.encoding, value="cp1251", command=self.reload_file),
        ]
        self.encoding_radios[0].grid(row=5, column=1, sticky="w", padx=(6, 0), pady=4)
        self.encoding_radios[1].grid(row=5, column=2, sticky="w", pady=4)

        settings = self._make_section(root, "Настройки обработки", padding=14, fill="x", pady=12)
        RoundedRadio(settings, text="Адрес → координаты", variable=self.mode, value="address_to_coords", command=self._refresh_controls).grid(row=0, column=0, sticky="w")
        RoundedRadio(settings, text="Координаты → адрес", variable=self.mode, value="coords_to_address", command=self._refresh_controls).grid(row=0, column=1, sticky="w", padx=20)

        ttk.Label(settings, text="Колонка адреса", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.address_wrap = RoundedField(settings, active=True)
        self.address_wrap.grid(row=2, column=0, sticky="we", pady=(2, 0))
        self.address_combo = ttk.Combobox(self.address_wrap.content, textvariable=self.address_column, state="readonly", width=32)
        self.address_combo.pack(fill="x", expand=True)
        ttk.Label(settings, text="Широта", style="Card.TLabel").grid(row=1, column=1, sticky="w", pady=(10, 0))
        self.lat_wrap = RoundedField(settings)
        self.lat_wrap.grid(row=2, column=1, sticky="we", pady=(2, 0), padx=(20, 0))
        self.lat_combo = ttk.Combobox(self.lat_wrap.content, textvariable=self.lat_column, state="readonly", width=24)
        self.lat_combo.pack(fill="x", expand=True)
        ttk.Label(settings, text="Долгота", style="Card.TLabel").grid(row=1, column=2, sticky="w", pady=(10, 0))
        self.lon_wrap = RoundedField(settings)
        self.lon_wrap.grid(row=2, column=2, sticky="we", pady=(2, 0), padx=(20, 0))
        self.lon_combo = ttk.Combobox(self.lon_wrap.content, textvariable=self.lon_column, state="readonly", width=24)
        self.lon_combo.pack(fill="x", expand=True)
        settings.columnconfigure(0, weight=2)
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(2, weight=1)

        action_row = ttk.Frame(root)
        action_row.pack(fill="x", pady=(0, 8))
        self.start_button = RoundedButton(action_row, text="Запустить обработку", command=self.start_processing, bg="#171a2e")
        self.start_button.pack(side="left", fill="y")
        status_panel = ttk.Frame(action_row, style="Status.TFrame", padding=(12, 8))
        status_panel.pack(side="left", fill="x", expand=True, padx=(12, 0))
        self.progress_label = ttk.Label(status_panel, text="0%", style="Status.TLabel", width=5, anchor="center")
        self.progress_label.pack(side="right", padx=(10, 0))
        ttk.Label(status_panel, textvariable=self.status, style="Status.TLabel").pack(side="top", anchor="w", fill="x")
        self.progress = ttk.Progressbar(status_panel, mode="determinate", style="Horizontal.TProgressbar")
        self.progress.pack(side="top", fill="x", expand=True, pady=(6, 0))

        table_frame = self._make_section(root, "Предпросмотр", padding=6, fill="both", expand=True)
        self.preview = ttk.Treeview(table_frame, show="headings")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.preview.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.preview.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        self._refresh_controls()

    def _make_section(self, parent: tk.Widget, title: str, padding: int = 10, **pack_options: Any) -> ttk.Frame:
        section = ttk.Frame(parent, style="Card.TFrame", padding=padding)
        section.pack(**pack_options)
        ttk.Label(section, text=title, style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        content = ttk.Frame(section, style="Card.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        section.columnconfigure(0, weight=1)
        section.rowconfigure(1, weight=1)
        return content

    def open_file(self) -> None:
        filename = filedialog.askopenfilename(filetypes=FILE_TYPES)
        if not filename:
            return
        self.source_file.set(filename)
        source_path = Path(filename)
        self._load_sheet_names(source_path)
        self._refresh_file_format_controls(source_path)
        self.reload_file()

    def choose_output_file(self) -> None:
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=FILE_TYPES)
        if filename:
            self.output_file.set(filename)

    def reload_file(self) -> None:
        filename = self.source_file.get().strip()
        if not filename:
            return
        try:
            self.table_data = read_table(
                filename,
                sheet_name=self.sheet_name.get(),
                delimiter=self.delimiter.get() or None,
                encoding=self.encoding.get(),
                has_header=self.has_header.get(),
                start_row=max(int(self.start_row.get()), 1),
            )
        except Exception as exc:
            messagebox.showerror("Ошибка открытия файла", str(exc))
            return
        self.result_data = None
        self.loaded_path = Path(filename)
        self._refresh_file_format_controls(self.loaded_path)
        self.status.set(f"Загружено строк: {len(self.table_data.rows)}")
        self._fill_columns()
        self._show_preview(self.table_data)

    def _load_sheet_names(self, source_path: Path) -> None:
        if source_path.suffix.lower() not in {".xlsx", ".xlsm"} or openpyxl is None:
            self.sheet_combo.configure(values=[])
            self.sheet_name.set("")
            return
        try:
            workbook = openpyxl.load_workbook(source_path, read_only=True, data_only=True)
        except Exception:
            self.sheet_combo.configure(values=[])
            self.sheet_name.set("")
            return
        self.sheet_combo.configure(values=workbook.sheetnames)
        if workbook.sheetnames and self.sheet_name.get() not in workbook.sheetnames:
            self.sheet_name.set(workbook.sheetnames[0])

    def _refresh_file_format_controls(self, source_path: Path | None = None) -> None:
        path = source_path or self.loaded_path
        suffix = path.suffix.lower() if path else ""
        delimited_file = suffix in {".csv", ".txt"}
        excel_file = suffix in {".xlsx", ".xlsm"}
        delimiter_state = "normal" if delimited_file or not excel_file else "disabled"
        encoding_state = "normal" if delimited_file or not excel_file else "disabled"
        self.delimiter_entry.configure(state=delimiter_state)
        for radio in self.encoding_radios:
            radio.button.configure(state=encoding_state)

    def start_processing(self) -> None:
        if self.table_data is None:
            messagebox.showwarning("Нет файла", "Сначала загрузите файл для обработки.")
            return
        if self.mode.get() == "address_to_coords" and not self.address_column.get():
            messagebox.showwarning("Нет колонки", "Выберите колонку с адресом.")
            return
        if self.mode.get() == "coords_to_address" and (not self.lat_column.get() or not self.lon_column.get()):
            messagebox.showwarning("Нет колонок", "Выберите колонки широты и долготы.")
            return
        if not self.output_file.get().strip():
            self.choose_output_file()
        if not self.output_file.get().strip():
            messagebox.showwarning("Нет пути сохранения", "Выберите путь для автоматического сохранения результата.")
            return

        self.start_button.config(state="disabled")
        self.progress_label.configure(text="0%")
        self.progress.configure(maximum=max(len(self.table_data.rows), 1), value=0)
        self.status.set("Обработка...")
        threading.Thread(target=self._process_in_thread, daemon=True).start()

    def save_result(self) -> None:
        if self.result_data is None:
            messagebox.showwarning("Нет результата", "Сначала выполните обработку.")
            return
        filename = self.output_file.get().strip()
        if not filename:
            self.choose_output_file()
            filename = self.output_file.get().strip()
        if not filename:
            return
        try:
            write_table(self.result_data, filename)
        except Exception as exc:
            messagebox.showerror("Ошибка сохранения", str(exc))
            return
        self.status.set(f"Результат сохранён: {filename}")

    def _process_in_thread(self) -> None:
        assert self.table_data is not None
        client = GeocodingClient()

        def progress(current: int, total: int, label: str) -> None:
            self._add_worker_event("progress", (current, total, label))

        try:
            if self.mode.get() == "address_to_coords":
                result = process_addresses(self.table_data, self.address_column.get(), client, progress)
            else:
                result = process_coordinates(self.table_data, self.lat_column.get(), self.lon_column.get(), client, progress)
        except Exception as exc:
            self._add_worker_event("error", str(exc))
        else:
            self._add_worker_event("done", result)

    def _add_worker_event(self, event: str, payload: Any) -> None:
        with self.worker_lock:
            self.worker_events.append((event, payload))

    def _poll_worker_events(self) -> None:
        with self.worker_lock:
            events = self.worker_events[:]
            self.worker_events.clear()
        for event, payload in events:
            if event == "progress":
                current, total, label = payload
                self.progress.configure(maximum=total, value=current)
                percent = int(current / max(total, 1) * 100)
                self.progress_label.configure(text=f"{percent}%")
                self.status.set(f"Обработка {current}/{total}: {label}")
            elif event == "error":
                self.start_button.config(state="normal")
                self.progress_label.configure(text="Ошибка")
                messagebox.showerror("Ошибка обработки", str(payload))
                self.status.set("Обработка остановлена")
            elif event == "done":
                self.start_button.config(state="normal")
                self.result_data = payload
                self._show_preview(self.result_data)
                output_path = self.output_file.get().strip()
                try:
                    write_table(self.result_data, output_path)
                except Exception as exc:
                    self.progress_label.configure(text="Ошибка")
                    messagebox.showerror("Ошибка сохранения", str(exc))
                    self.status.set("Обработка завершена, но файл не сохранён")
                else:
                    self.progress.configure(value=self.progress.cget("maximum"))
                    self.progress_label.configure(text="100%")
                    self.status.set(f"Готово. Файл сохранён: {output_path}")
        self.after(150, self._poll_worker_events)

    def _fill_columns(self) -> None:
        if self.table_data is None:
            return
        columns = self.table_data.headers
        for combo in (self.address_combo, self.lat_combo, self.lon_combo):
            combo.configure(values=columns)
        self.columns_dropdown.set_columns(columns)
        self.address_column.set(guess_column(columns, ["адрес", "address", "addr"]))
        self.lat_column.set(guess_column(columns, ["lat", "latitude", "шир", "широта"]))
        self.lon_column.set(guess_column(columns, ["lon", "lng", "longitude", "долг", "долгота"]))
        self._refresh_controls()

    def _on_work_columns_changed(self) -> None:
        if self.table_data is not None:
            self._show_preview(self.table_data)

    def _selected_preview_columns(self, table: TableData) -> list[str]:
        selected = self.columns_dropdown.selected_columns()
        visible = selected or table.headers
        extra_columns = [column for column in table.headers if column not in self.columns_dropdown.columns]
        return [column for column in [*visible, *extra_columns] if column in table.headers]

    def _refresh_controls(self) -> None:
        address_mode = self.mode.get() == "address_to_coords"
        self.address_combo.configure(state="readonly" if address_mode else "disabled")
        self.lat_combo.configure(state="readonly" if not address_mode else "disabled")
        self.lon_combo.configure(state="readonly" if not address_mode else "disabled")
        self.address_wrap.set_active(address_mode)
        self.lat_wrap.set_active(not address_mode)
        self.lon_wrap.set_active(not address_mode)

    def _show_preview(self, table: TableData) -> None:
        columns = self._selected_preview_columns(table)
        self.preview.delete(*self.preview.get_children())
        self.preview.configure(columns=columns)
        for column in columns:
            self.preview.heading(column, text=column)
            self.preview.column(column, width=160, minwidth=80, stretch=True)
        for row in table.rows[:200]:
            self.preview.insert("", "end", values=[str(row.get(column, "")) for column in columns])


def guess_column(columns: list[str], needles: list[str]) -> str:
    for column in columns:
        lowered = column.lower()
        if any(needle in lowered for needle in needles):
            return column
    return columns[0] if columns else ""


def main() -> None:
    app = GeocodeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
