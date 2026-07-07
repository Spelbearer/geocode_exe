# GeocodeEXE

Теперь приложение собрано в один переносимый файл: `geocode_exe.py`.

## Как передавать

- Для разработчика/пользователя с Python: передайте только `geocode_exe.py`.
- Для пользователя без Python: соберите один `GeocodeEXE.exe` и передайте только этот `.exe`.

## Запуск одного Python-файла

```bash
python geocode_exe.py
```

CSV/TXT и запросы к сервису геокодирования работают без сторонних Python-библиотек. Для открытия/сохранения Excel `.xlsx/.xlsm` нужен `openpyxl`:

```bash
python -m pip install openpyxl
```

Для сборки EXE используйте отдельный минимальный список зависимостей из `requirements-exe.txt`: он содержит только `pyinstaller` для сборки и `openpyxl` для Excel. Зависимость `et_xmlfile` будет установлена автоматически как обязательная зависимость `openpyxl`.

## Сборка одного EXE

```bash
python -m pip install --upgrade -r requirements-exe.txt
pyinstaller --onefile --windowed --name GeocodeEXE --hidden-import=openpyxl geocode_exe.py
```

Готовый файл будет здесь:

```text
dist/GeocodeEXE.exe
```

После сборки можно передавать пользователю только `dist/GeocodeEXE.exe`; папки проекта и Python ему не нужны.

Если собираете другой входной файл, например `kml_generator_app.py`, оставьте тот же принцип: сначала установите `requirements-exe.txt`, затем добавьте `--hidden-import=openpyxl` в команду PyInstaller:

```bash
python -m pip install --upgrade -r requirements-exe.txt
pyinstaller --onefile --windowed --hidden-import=openpyxl kml_generator_app.py
```

Опция `--hidden-import=openpyxl` нужна, потому что приложение подключает Excel-поддержку динамически, а PyInstaller не всегда видит такие импорты автоматически.

## Что умеет интерфейс

- Загружать `.xlsx`, `.xlsm`, `.csv`, `.txt`.
- Определять координаты по адресу.
- Определять адрес по координатам.
- Показывать предпросмотр первых 200 строк.
- Сохранять результат в Excel, CSV или TXT.
- Использовать встроенные настройки сервиса без показа технических URL пользователю.
- Добавлять в результат понятные колонки: `Найденный адрес`, `Широта результата`, `Долгота результата`.
