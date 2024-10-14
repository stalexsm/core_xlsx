## Вспомогательная библиотека __boriy_core__.

`core_xlsx` - это Python-библиотека для работы с данными в виде Excel (XLSX) для проекта, предоставляющая удобные инструменты для манипуляции данными в таблицах.

## Установка

```
poetry add core-xlsx
```

## Основные компоненты

- `XLSXSheet`: Представляет лист Excel.
- `XLSXSheetCell`: Представляет ячейку в листе Excel.
- `Service`: Базовый класс для создания сервисов обработки данных и создания отчетов.
- `Formatter`: Базовый класс для форматирования данных.
- `HelperSheet`: Помощник для работы с несколькими листами и ячейками.
- `HelperSheetCell`: Помощник для работы с ячейками.

## Возможности

- Поиск листов по имени или шаблону
- Поиск ячеек по значению, регулярному выражению или адресу
- Манипуляция данными ячеек (установка значений, формул, форматов)
- Получение информации о ячейках (тип данных, числовой формат, стиль)
- Работа с датами и временем
- Создание пользовательских сервисов и форматтеров для создпния отчетов.

## Пример использования

```python
from core_xlsx import XLSXSheet, Service, HelperSheet

class MyService(Service):
    def summary_0(self, sheets):
        h = HelperSheet(sheets)
        sheet = h.find_sheet_by_pattern("Отчет")

        if sheet:
            cell = sheet.find_cell_pattern_regex("Итого:")
            if cell:
                total = float(cell.value)
                print(f"Итоговая сумма: {total}")

        return sheets

# Использование
service = MyService(uow="my_unit_of_work")
sheets = [...]  # Ваши данные листов

processed_sheets = service.summary_0(sheets)
```
