## 0.5.3 (2024-10-23)

### Fix

- Ускорение методов write_style_for_cell и write_formula_for_cell

## 0.5.2 (2024-10-23)

### Fix

- Правки в методе write_cell, write_style_for_cell, write_formula_for_cell, write_cell_with_formula. Ускорение методов

## 0.5.1 (2024-10-18)

### Fix

- Исправлена обработка XLSXSheetCell::value, если из python передается None
- Оптимизация метода write_cell_with_formula
- Оптимизация метода write_cell

## 0.5.0 (2024-10-18)

### Feat

- Добавление новых методов find_cells_by_range_rows и find_cells_by_range_cols для поиска ячеек

### Fix

- Правки в HelperSheetCell, добавление метода iter_cells

## 0.4.0 (2024-10-18)

### Feat

- Добавление функции для преобразование xlsxsheets в PyDict
- Добавление функции для конвертации данных xlsxsheets в json

## 0.3.0 (2024-10-17)

### Feat

- Добавление нового метода в XLSXSheet и HelperSheetCell iter_cells. Для получения ячеек по координатам

## 0.2.0 (2024-10-17)

### Feat

- Добавление методов в XLSXSheet для записи формулы и стиля ячейки

### Fix

- Правки метода write_cell. Убрал возврат ячейки с данного метода

## 0.1.2 (2024-10-16)

### Fix

- Удаление класса Formatter и перенос методов в класс Service
- Убрал лишнюю расраковку
- Правки названия метода в pyi для HelperSheetCell

## 0.1.1 (2024-10-15)

### Fix

- Правки в названии метода в pyi
- Перенос метода merged_cells в sheet. Удаление merge_cell из ячейки
- Правки метода write_cell в core_rs. Исправлено, возврат мутабельной ссылки для ячейки

## 0.1.0 (2024-10-14)
