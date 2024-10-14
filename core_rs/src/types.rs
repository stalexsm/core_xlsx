use std::sync::Arc;

use anyhow::{bail, Result};
use chrono::NaiveDateTime;
use serde::{Deserialize, Serialize};

use crate::{
    datatype::{CellRawValue, CellValue},
    helper::HelperSheetCell,
    utils::{column_number_to_letter, get_number_format_by_datatype},
};

// Список типов
const DATA_TYPES: [&str; 4] = ["s", "n", "d", "b"];

#[derive(Clone, Debug, Serialize, Deserialize, Default)]
pub struct XLSXSheet {
    pub name: String,
    pub max_row: u32,
    pub max_column: u32,
    pub index: i32,
    pub main: bool,
    pub cells: Vec<XLSXSheetCell>,
}

impl XLSXSheet {
    pub fn new(name: String, index: i32, cols: u32, rows: u32) -> Self {
        let mut cells = Vec::with_capacity((rows * cols) as usize);
        for r in 1..=rows {
            for c in 1..=cols {
                let cell = XLSXSheetCell::new(r, c, None);
                cells.push(cell);
            }
        }

        Self {
            name,
            index,
            cells,
            max_row: rows,
            max_column: cols,
            ..Default::default()
        }
    }

    /// Добавление значения в ячейку по координате.
    /// Дополнительно создаются несуществующие ячейки.
    pub fn write_cell(&mut self, row: u32, col: u32, value: String) -> Result<Arc<XLSXSheetCell>> {
        let cell_index = self
            .cells
            .iter()
            .position(|c| c.row == row && c.column == col);

        match cell_index {
            Some(index) => {
                self.cells[index].set_value(value)?;

                Ok(Arc::new(self.cells[index].clone()))
            }
            None => {
                // Обновим максимальные значения
                self.max_row = self.max_row.max(row);
                self.max_column = self.max_column.max(col);

                // Добавление заданной ячейки
                let new_cell = XLSXSheetCell::new(row, col, Some(value));
                self.cells.push(new_cell.clone());

                // Генерация несуществующих ячеек.
                for r in 1..=row {
                    for c in 1..=col {
                        if !self.cells.iter().any(|x| x.row == r && x.column == c) {
                            let cell = XLSXSheetCell::new(r, c, None);
                            self.cells.push(cell);
                        }
                    }
                }

                // Сортируем, чтобы все было упорядочено.
                self.cells
                    .sort_by(|a, b| a.row.cmp(&b.row).then_with(|| a.column.cmp(&b.column)));

                Ok(Arc::new(new_cell))
            }
        }
    }

    /// Метод для удаления колонок
    pub fn delete_cols(&mut self, idx: u32, amount: u32) -> Result<()> {
        // Remove cells in the specified columns
        self.cells
            .retain(|cell| cell.column < idx || cell.column >= idx + amount);

        // Update column numbers for remaining cells
        for cell in &mut self.cells {
            if cell.column > idx {
                cell.column -= amount;
                // Update the cell's letter coordinate
                let new_letter = column_number_to_letter(cell.column);
                cell.cell = format!("{}{}", new_letter, cell.row);
            }
        }

        // Update max_column if necessary
        self.max_column = self.max_column.saturating_sub(amount);

        // Sort cells to maintain order
        self.cells
            .sort_by(|a, b| a.row.cmp(&b.row).then_with(|| a.column.cmp(&b.column)));

        Ok(())
    }

    /// Метод для удаления  строк
    pub fn delete_rows(&mut self, idx: u32, amount: u32) -> Result<()> {
        // Remove cells in the specified rows
        self.cells
            .retain(|cell| cell.row < idx || cell.row >= idx + amount);

        // Update row numbers for remaining cells
        for cell in &mut self.cells {
            if cell.row > idx {
                cell.row -= amount;
                // Update the cell's coordinate
                let letter = column_number_to_letter(cell.column);
                cell.cell = format!("{}{}", letter, cell.row);
            }
        }

        // Update max_row if necessary
        self.max_row = self.max_row.saturating_sub(amount);

        // Sort cells to maintain order
        self.cells
            .sort_by(|a, b| a.row.cmp(&b.row).then_with(|| a.column.cmp(&b.column)));

        Ok(())
    }

    /// Поиск ячейки по шаблону
    pub fn find_cell_pattern_regex(&self, pattern: &str) -> Result<Option<XLSXSheetCell>> {
        HelperSheetCell::find_cell_pattern_regex(pattern, &self.cells)
    }

    // Поиск ячеек по шаблону
    pub fn find_cells_pattern_regex(&self, pattern: &str) -> Result<Vec<XLSXSheetCell>> {
        HelperSheetCell::find_cells_pattern_regex(pattern, &self.cells)
    }

    /// Поиск ячеек колонок для строк которые соответствуют патерну
    pub fn find_cells_for_rows_pattern_regex(
        &self,
        pattern: &str,
        col_stop: Option<u32>,
    ) -> Result<Vec<XLSXSheetCell>> {
        HelperSheetCell::find_cells_for_rows_pattern_regex(pattern, &self.cells, col_stop)
    }

    /// Поиск ячеек строк для колонок которые соответствуют патерну
    pub fn find_cells_for_cols_pattern_regex(
        &self,
        pattern: &str,
        row_stop: Option<u32>,
    ) -> Result<Vec<XLSXSheetCell>> {
        HelperSheetCell::find_cells_for_cols_pattern_regex(pattern, &self.cells, row_stop)
    }

    /// Поиск ячеек с помощью ИЛИ ячейки по патернам
    pub fn find_cells_multi_pattern_regex(
        &self,
        pattern_1: &str,
        pattern_2: &str,
    ) -> Result<Vec<XLSXSheetCell>> {
        HelperSheetCell::find_cells_multi_pattern_regex(pattern_1, pattern_2, &self.cells)
    }

    /// Поиск ячейки по буквенной координате A1 (cell)
    pub fn find_cell_by_cell(&self, cell: &str) -> Result<Option<XLSXSheetCell>> {
        HelperSheetCell::find_cell_by_cell(cell, &self.cells)
    }

    /// Поиск ячейки по координате
    pub fn find_cell_by_coordinate(&self, row: u32, col: u32) -> Result<Option<XLSXSheetCell>> {
        HelperSheetCell::find_cell_by_coordinate(row, col, &self.cells)
    }

    /// Поиск ячеек между шаьлонами
    pub fn find_cells_between_patterns(
        &self,
        pattern_after: &str,
        pattern_before: &str,
    ) -> Result<Vec<XLSXSheetCell>> {
        HelperSheetCell::find_cells_between_patterns(pattern_after, pattern_before, &self.cells)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, Default)]
pub struct XLSXSheetCell {
    pub row: u32,
    pub column: u32,
    pub cell: String,
    pub value: CellValue,
    pub formula: Option<String>,
    pub data_type: String,
    pub number_format: String,
    pub cell_type: Option<String>,
    pub el_type: Option<String>,
    pub is_merge: bool,
    pub start_row: Option<u32>,
    pub end_row: Option<u32>,
    pub start_column: Option<u32>,
    pub end_column: Option<u32>,
    pub style_id: Option<String>,
    // Даннык поля только в движке пока что.
    pub hidden_value: Option<String>,
    pub comment: Option<String>,
}

impl XLSXSheetCell {
    pub fn new(row: u32, col: u32, value: Option<String>) -> Self {
        // Получение  letter (cell)
        let cell = column_number_to_letter(col);
        let cell = format!("{}{}", cell, row);

        // Обработка значения
        let mut raw_value = CellRawValue::Empty;
        if let Some(val) = value {
            raw_value = CellValue::quess_typed_value(&val)
        }

        let value = CellValue {
            raw_value: raw_value.clone(),
        };
        // Определение datetype
        let data_type = raw_value.get_date_type().to_string();
        // Определение number format
        let number_format = get_number_format_by_datatype(&data_type);

        Self {
            row,
            column: col,
            cell,
            value,
            data_type,
            number_format,
            ..Default::default()
        }
    }

    /// Метод для получения значения ячейки.
    pub fn set_value(&mut self, value: String) -> Result<()> {
        let cell_value = self.value.set_value(value);

        if self.formula.is_none() {
            let data_type = cell_value.get_data_type();

            // Обновим данные, так как это формула
            self.number_format = get_number_format_by_datatype(data_type);
            self.data_type = data_type.to_string();
        }

        Ok(())
    }

    /// Метод для добавления данных по объединению ячеек.
    pub fn set_merge(
        &mut self,
        start_row: u32,
        end_row: u32,
        start_column: u32,
        end_column: u32,
    ) -> Result<()> {
        self.is_merge = true;

        self.start_row = Some(start_row);
        self.end_row = Some(end_row);
        self.start_column = Some(start_column);
        self.end_column = Some(end_column);

        Ok(())
    }

    /// Метод для получения значения ячейки Numbers.
    pub fn set_value_number(&mut self, value: f64) -> Result<()> {
        self.value.set_value_number(value);

        Ok(())
    }

    /// Метод для получения значения ячейки Bool.
    pub fn set_value_bool(&mut self, value: bool) -> Result<()> {
        self.value.set_value_bool(value);

        Ok(())
    }

    /// Метод для получения значения ячейки String.
    pub fn set_value_str(&mut self, value: String) -> Result<()> {
        self.value.set_value_str(value);

        Ok(())
    }

    /// Метод для получения значения ячейки Datetime.
    pub fn set_value_datetime(&mut self, value: NaiveDateTime) -> Result<()> {
        self.value.set_value_datatime(value);

        Ok(())
    }

    /// Метод для получения значения ячейки String.
    pub fn set_formula(&mut self, value: String) -> Result<()> {
        self.formula = Some(value);
        // если идет установка формулы, то и тип установим как формула
        self.data_type = "f".to_string();

        Ok(())
    }

    /// Метод для получения значения ячейки data_type.
    pub fn set_data_type(&mut self, value: String) -> Result<()> {
        if !DATA_TYPES.contains(&value.as_str()) {
            bail!(format!("value not in {:?}", DATA_TYPES))
        }

        self.data_type = value;

        Ok(())
    }

    /// Метод для получения значения ячейки number_format.
    pub fn set_number_format(&mut self, value: String) -> Result<()> {
        self.number_format = value;

        Ok(())
    }

    /// Метод для получения значения ячейки cell_type.
    pub fn set_cell_type(&mut self, value: String) -> Result<()> {
        self.cell_type = Some(value);

        Ok(())
    }

    /// Метод для получения значения ячейки el_type.
    pub fn set_el_type(&mut self, value: String) -> Result<()> {
        self.el_type = Some(value);

        Ok(())
    }

    /// Метод для получения флага, ячейка с формулой или нет.
    pub fn is_formula(&self) -> Result<bool> {
        Ok(self.formula.is_some() && self.data_type == *"f")
    }

    /// Проверить, является ли значение ячейки boolean
    pub fn is_value_bool(&self) -> Result<bool> {
        Ok(self.value.is_bool())
    }

    /// Проверить, является ли значение ячейки numeric
    pub fn is_value_numeric(&self) -> Result<bool> {
        Ok(self.value.is_numeric())
    }

    /// Проверить, является ли значение ячейки datetime
    pub fn is_value_datetime(&self) -> Result<bool> {
        Ok(self.value.is_datetime())
    }

    /// Проверить, является ли значение ячейки empty
    pub fn is_value_empty(&self) -> Result<bool> {
        Ok(self.value.is_empty())
    }

    /// Метод для добавления стиля к ячейки
    pub fn set_style_id(&mut self, value: String) -> Result<()> {
        self.style_id = Some(value);

        Ok(())
    }
}
