use crate::types::XLSXSheet;

/// Преобразование номера колонки в букву.
/// Данный метод используется по обработки формирования данных excel.
/// 1. Принимает номер колонки как `u32`.
/// 2. Создает пустую строку для хранения результата.
/// 3. В цикле:
///    - Уменьшает номер колонки на 1 (так как в Excel колонки начинаются с 1, а не с 0).
///    - Вычисляет остаток от деления на 26 (количество букв в английском алфавите).
///    - Преобразует остаток в соответствующую букву и добавляет её к результату.
///    - Делит номер колонки на 26 для перехода к следующей "цифре" в системе счисления с основанием 26.
/// 4. Переворачивает полученную строку, так как буквы были добавлены в обратном порядке.
pub fn column_number_to_letter(mut column: u32) -> String {
    let mut result = String::new();

    while column > 0 {
        column -= 1; // Уменьшаем на 1, так как в Excel колонки начинаются с 1, а не с 0

        let remainder = (column % 26) as u8;
        result.push((b'A' + remainder) as char);
        column /= 26;
    }

    result.chars().rev().collect()
}

/// конвертация данных в json
pub fn xlsxheets_to_json(sheets: Vec<XLSXSheet>) -> String {
    serde_json::to_string(&sheets).unwrap()
}

/// Определение формата по типу
pub(crate) fn get_number_format_by_datatype(data_type: &str) -> String {
    match data_type {
        "n" => "General".to_string(),
        "d" => "mm-dd-yy".to_string(),
        _ => "@".to_string(),
    }
}
