# Конвертер Jira CSV ↔ XLSX

<p align="center">
  <img src="assets/icon_128.png" width="96" alt="Core-Cat icon"/>
</p>

Десктопное приложение для Windows — преобразует экспорт задач Jira (CSV) в Excel (XLSX) и обратно.

**Автор:** [maxsteff](https://t.me/maxsteff)

---

## Возможности

- **Drag & Drop** — перетащите один или несколько CSV в окно
- **Склейка CSV** — несколько файлов объединяются в один XLSX, дубли убираются
- **Очистка заголовков** — `Пользовательское поле (Статус)` → `Статус`
- **Фильтрация колонок** через `columns.txt` (секции `[columns]` / `[strip]`)
- **Конвертация дат** Jira (`05/фев/26 12:00 AM`) → `05.02.2026`
- **22 стиля таблиц** Excel с визуальным превью
- **Обратная конвертация** XLSX → CSV для загрузки обратно в Jira
- **Тёмная / светлая тема** — автоопределение из реестра Windows
- **8 языков** интерфейса: RU, EN, DE, FR, ES, ZH, AR, PT

---

## Установка и запуск

### Готовый .exe

Скачайте `Конвертер_Jira_CSV_XLSX.exe` из [Releases](../../releases) — установка не нужна.

### Из исходного кода

```bash
git clone https://github.com/<user>/jira-csv-xlsx-converter.git
cd jira-csv-xlsx-converter
pip install -r requirements.txt
python app.py
```

### Сборка .exe

```bash
pip install pyinstaller Pillow
python make_icons.py          # регенерировать иконки (уже есть в assets/)
python -m PyInstaller app.spec
# → dist/Конвертер_Jira_CSV_XLSX.exe
```

---

## Файл columns.txt

Положите рядом с `.exe` — подхватится автоматически.

```
[columns]
Ключ проблемы
Тема
Статус
Дата регистрации замечания

[strip]
Пользовательское поле 
Custom field 
```

- **`[columns]`** — какие колонки оставить, в каком порядке
- **`[strip]`** — префиксы для автоочистки имён (пробел в конце значим!)

Нет файла — кнопка **«Создать пример»** сгенерирует шаблон из текущего CSV.

---

## Тесты

```bash
pip install pytest
pytest tests/ -v
```

86 тестов покрывают всю бизнес-логику (без GUI):

| Группа | Тестов |
|---|---|
| `parse_jira_date` — парсинг дат | 11 |
| `auto_rename` — очистка заголовков | 8 |
| `parse_columns_txt` — разбор фильтра | 8 |
| `load_columns_filter` | 3 |
| `convert` CSV → XLSX | 16 |
| `convert_xlsx_to_csv` XLSX → CSV | 6 |
| `generate_columns_example` | 6 |
| `ALL_TABLE_STYLES` — палитра | 9 |
| `STRINGS` — локализация (8 языков) | 18 |
| Round-trip интеграция | 3 |

---

## Структура проекта

```
├── app.py                  ← основной код
├── app.spec                ← конфиг PyInstaller
├── requirements.txt
├── pytest.ini
├── columns.txt.example     ← пример фильтра колонок
├── assets/
│   ├── icon.ico            ← иконка для .exe (все размеры)
│   ├── icon_256.png
│   └── ...
├── tests/
│   └── test_core.py        ← 86 тестов
└── .github/
    └── workflows/
        └── tests.yml       ← CI на windows-latest
```

---

## Требования

- Windows 10 / 11 (для `.exe`)
- Python 3.11+ (для запуска из исходников)
- `pandas`, `openpyxl`, `tkinterdnd2` (опционально)

---

## Лицензия

MIT
