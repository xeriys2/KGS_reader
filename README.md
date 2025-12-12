# KGS_Reader

Приложение (Tkinter) для пакетной обработки PDF: извлечение текста и формирование Excel-отчёта. При необходимости может использовать OCR (Tesseract).

## Возможности

- Выбор PDF-файлов и папок (с учётом вложенности).
- Обработка PDF через PyMuPDF (fitz).
- OCR страниц через Tesseract (если установлен или лежит рядом с приложением в `tesseract\`).
- Выгрузка результата в Excel (`.xlsx`) через openpyxl.
- Редактор типов коммуникаций (чекбоксы/ПКМ), хранится в `comm_types.json` рядом с программой.
- Прогресс-бар, счётчик файлов/страниц и кнопка “Отмена”.
- Сохранение настроек и последних путей в `settings.json` рядом с программой.

## Запуск из исходников

Требуется Python 3.11+.

Установка зависимостей:

```bash
pip install -r requirements_portable.txt
```

Запуск:

```bash
python "KGS_Reader v6.py"
```

## Портативная сборка (Windows)

Рекомендуемый способ (сборка “лёгкая”, в чистом venv):

```powershell
.\build_portable_light.ps1
```

По умолчанию скрипт старается положить рядом portable-OCR (Tesseract + `rus/eng`). Если OCR не нужен:

```powershell
.\build_portable_light.ps1 -WithoutOcr
```

Альтернатива (если `pyinstaller` уже установлен в текущем окружении; может получиться “тяжёлая” сборка):

```powershell
.\build_portable.ps1
```

Выход:

- Папка: `dist\KGS_Reader\` (копируйте/переносите целиком, не только `.exe`).
- Архив: `release\KGS_Reader_portable.zip` (удобно раздавать “без заморочек”: распаковал и запустил).

### OCR в portable

Если рядом есть папка `tesseract\` (например `tesseract\tesseract.exe` и `tesseract\tessdata\...`), приложение автоматически попытается использовать её.

## Лицензия

Проект распространяется под GNU AGPL-3.0; PyMuPDF также под AGPL. См. `LICENSE`.
