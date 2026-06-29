"""
Однофайловое GUI-приложение для пакетного геокодирования.

Можно передавать этот файл как есть или собрать из него один EXE:
    pyinstaller --onefile --windowed --name GeocodeEXE geocode_exe.py

Зависимости для запуска .py:
- Python 3.10+
- tkinter (обычно входит в Python для Windows)
- openpyxl нужен только для чтения/записи Excel .xlsx/.xlsm

CSV/TXT и API работают только на стандартной библиотеке Python.
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
            return {"api_address": "", "api_lat": "", "api_lon": "", "api_raw": ""}

        data = self._post(self.forward_url, {"query": address})
        suggestions = data.get("suggestions") or []
        if not suggestions:
            return {"api_address": "", "api_lat": "", "api_lon": "", "api_raw": json.dumps(data, ensure_ascii=False)}

        first = suggestions[0]
        details = first.get("data") or {}
        return {
            "api_address": first.get("unrestricted_value") or first.get("value") or "",
            "api_lat": str(details.get("geo_lat") or ""),
            "api_lon": str(details.get("geo_lon") or ""),
            "api_raw": json.dumps(first, ensure_ascii=False),
        }

    def reverse_geocode(self, lat: Any, lon: Any) -> dict[str, str]:
        lat_text = str(lat or "").strip().replace(",", ".")
        lon_text = str(lon or "").strip().replace(",", ".")
        if not lat_text or not lon_text:
            return {"api_address": "", "api_lat": lat_text, "api_lon": lon_text, "api_raw": ""}

        data = self._post(self.reverse_url, {"lat": lat_text, "lon": lon_text})
        suggestions = data.get("suggestions") or []
        if not suggestions:
            return {"api_address": "", "api_lat": lat_text, "api_lon": lon_text, "api_raw": json.dumps(data, ensure_ascii=False)}

        first = suggestions[0]
        return {
            "api_address": first.get("unrestricted_value") or first.get("value") or "",
            "api_lat": lat_text,
            "api_lon": lon_text,
            "api_raw": json.dumps(first, ensure_ascii=False),
        }

    def _post(self, url: str, api_request: dict[str, str]) -> dict[str, Any]:
        payload = urllib.parse.urlencode({"apiRequest": json.dumps(api_request, ensure_ascii=False)}).encode("utf-8")
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
            raise GeocodingError(f"Ошибка запроса к API: {exc}") from exc
        except TimeoutError as exc:
            raise GeocodingError("Превышено время ожидания API") from exc

        try:
            parsed = json.loads(text)
        except json.JSONDecodeError as exc:
            raise GeocodingError("API вернул ответ не в формате JSON") from exc
        if not isinstance(parsed, dict):
            raise GeocodingError("API вернул неожиданный формат ответа")
        return parsed


def read_table(path: str | Path) -> TableData:
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return read_excel(file_path)
    if suffix in {".csv", ".txt"}:
        return read_delimited(file_path)
    raise ValueError(f"Неподдерживаемый формат файла: {suffix}")


def read_delimited(path: Path) -> TableData:
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    sample = raw[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel_tab if path.suffix.lower() == ".txt" else csv.excel
    rows = list(csv.DictReader(raw.splitlines(), dialect=dialect))
    headers = list(rows[0].keys()) if rows else []
    return TableData(headers=headers, rows=rows)


def read_excel(path: Path) -> TableData:
    if openpyxl is None:
        raise ValueError("Для Excel нужен пакет openpyxl. Для EXE он должен быть установлен перед сборкой.")
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    values = list(sheet.iter_rows(values_only=True))
    if not values:
        return TableData(headers=[], rows=[])
    headers = [str(value or "") for value in values[0]]
    rows = []
    for source_row in values[1:]:
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


def append_api_columns(table: TableData) -> TableData:
    result = table.copy()
    for column in ["api_address", "api_lat", "api_lon", "api_raw"]:
        if column not in result.headers:
            result.headers.append(column)
    return result


def process_addresses(
    table: TableData,
    address_column: str,
    client: GeocodingClient,
    progress: Callable[[int, int, str], None] | None = None,
) -> TableData:
    result = append_api_columns(table)
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
    result = append_api_columns(table)
    total = len(result.rows)
    for index, row in enumerate(result.rows, start=1):
        lat = row.get(lat_column, "")
        lon = row.get(lon_column, "")
        row.update(client.reverse_geocode(lat, lon))
        if progress:
            progress(index, total, f"{lat}, {lon}")
    return result


class GeocodeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Геокодирование файлов")
        self.geometry("980x700")
        self.minsize(900, 620)

        self.table_data: TableData | None = None
        self.result_data: TableData | None = None
        self.loaded_path: Path | None = None
        self.worker_events: list[tuple[str, Any]] = []
        self.worker_lock = threading.Lock()

        self.mode = tk.StringVar(value="address_to_coords")
        self.address_column = tk.StringVar()
        self.lat_column = tk.StringVar()
        self.lon_column = tk.StringVar()
        self.forward_url = tk.StringVar(value=DEFAULT_FORWARD_URL)
        self.reverse_url = tk.StringVar(value=DEFAULT_REVERSE_URL)
        self.status = tk.StringVar(value="Загрузите Excel, CSV или TXT файл")

        self._build_ui()
        self.after(150, self._poll_worker_events)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        file_row = ttk.Frame(root)
        file_row.pack(fill="x")
        ttk.Button(file_row, text="Открыть файл", command=self.open_file).pack(side="left")
        ttk.Button(file_row, text="Сохранить результат", command=self.save_result).pack(side="left", padx=(8, 0))
        self.file_label = ttk.Label(file_row, text="Файл не выбран")
        self.file_label.pack(side="left", padx=12)

        settings = ttk.LabelFrame(root, text="Настройки обработки", padding=10)
        settings.pack(fill="x", pady=12)
        ttk.Radiobutton(settings, text="Адрес → координаты", variable=self.mode, value="address_to_coords", command=self._refresh_controls).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(settings, text="Координаты → адрес", variable=self.mode, value="coords_to_address", command=self._refresh_controls).grid(row=0, column=1, sticky="w", padx=20)

        ttk.Label(settings, text="Колонка адреса").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.address_combo = ttk.Combobox(settings, textvariable=self.address_column, state="readonly", width=32)
        self.address_combo.grid(row=2, column=0, sticky="we", pady=(2, 0))
        ttk.Label(settings, text="Широта").grid(row=1, column=1, sticky="w", pady=(10, 0))
        self.lat_combo = ttk.Combobox(settings, textvariable=self.lat_column, state="readonly", width=24)
        self.lat_combo.grid(row=2, column=1, sticky="we", pady=(2, 0), padx=(20, 0))
        ttk.Label(settings, text="Долгота").grid(row=1, column=2, sticky="w", pady=(10, 0))
        self.lon_combo = ttk.Combobox(settings, textvariable=self.lon_column, state="readonly", width=24)
        self.lon_combo.grid(row=2, column=2, sticky="we", pady=(2, 0), padx=(20, 0))

        ttk.Label(settings, text="API адрес → координаты").grid(row=3, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(settings, textvariable=self.forward_url).grid(row=4, column=0, columnspan=3, sticky="we")
        ttk.Label(settings, text="API координаты → адрес").grid(row=5, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(settings, textvariable=self.reverse_url).grid(row=6, column=0, columnspan=3, sticky="we")
        settings.columnconfigure(0, weight=2)
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(2, weight=1)

        action_row = ttk.Frame(root)
        action_row.pack(fill="x", pady=(0, 8))
        self.start_button = ttk.Button(action_row, text="Запустить обработку", command=self.start_processing)
        self.start_button.pack(side="left")
        self.progress = ttk.Progressbar(action_row, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=12)
        ttk.Label(action_row, textvariable=self.status).pack(side="left")

        table_frame = ttk.LabelFrame(root, text="Предпросмотр", padding=6)
        table_frame.pack(fill="both", expand=True)
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

    def open_file(self) -> None:
        filename = filedialog.askopenfilename(filetypes=FILE_TYPES)
        if not filename:
            return
        try:
            self.table_data = read_table(filename)
        except Exception as exc:
            messagebox.showerror("Ошибка открытия файла", str(exc))
            return
        self.result_data = None
        self.loaded_path = Path(filename)
        self.file_label.config(text=str(self.loaded_path))
        self.status.set(f"Загружено строк: {len(self.table_data.rows)}")
        self._fill_columns()
        self._show_preview(self.table_data)

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

        self.start_button.config(state="disabled")
        self.progress.configure(maximum=max(len(self.table_data.rows), 1), value=0)
        self.status.set("Обработка...")
        threading.Thread(target=self._process_in_thread, daemon=True).start()

    def save_result(self) -> None:
        if self.result_data is None:
            messagebox.showwarning("Нет результата", "Сначала выполните обработку.")
            return
        initial_name = "result.xlsx"
        if self.loaded_path:
            initial_name = f"{self.loaded_path.stem}_result.xlsx"
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=initial_name, filetypes=FILE_TYPES)
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
        client = GeocodingClient(forward_url=self.forward_url.get(), reverse_url=self.reverse_url.get())

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
                self.status.set(f"{current}/{total}: {label}")
            elif event == "error":
                self.start_button.config(state="normal")
                messagebox.showerror("Ошибка обработки", str(payload))
                self.status.set("Обработка остановлена")
            elif event == "done":
                self.start_button.config(state="normal")
                self.result_data = payload
                self._show_preview(self.result_data)
                self.status.set(f"Готово. Обработано строк: {len(self.result_data.rows)}")
        self.after(150, self._poll_worker_events)

    def _fill_columns(self) -> None:
        if self.table_data is None:
            return
        columns = self.table_data.headers
        for combo in (self.address_combo, self.lat_combo, self.lon_combo):
            combo.configure(values=columns)
        self.address_column.set(guess_column(columns, ["адрес", "address", "addr"]))
        self.lat_column.set(guess_column(columns, ["lat", "latitude", "шир", "широта"]))
        self.lon_column.set(guess_column(columns, ["lon", "lng", "longitude", "долг", "долгота"]))
        self._refresh_controls()

    def _refresh_controls(self) -> None:
        address_mode = self.mode.get() == "address_to_coords"
        self.address_combo.configure(state="readonly" if address_mode else "disabled")
        self.lat_combo.configure(state="readonly" if not address_mode else "disabled")
        self.lon_combo.configure(state="readonly" if not address_mode else "disabled")

    def _show_preview(self, table: TableData) -> None:
        columns = table.headers
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
