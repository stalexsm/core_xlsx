pub(crate) mod helper;
pub(crate) mod services;
pub(crate) mod types;

use helper::{WrapperHelperSheet, WrapperHelperSheetCell};
use pyo3::prelude::*;
use services::Service;
use types::{WrapperXLSXSheet, WrapperXLSXSheetCell};

/// Преобразование номера колонки в букву.
#[pyfunction]
fn column_number_to_letter(col: u32) -> PyResult<String> {
    let letter = core_rs::utils::column_number_to_letter(col);

    Ok(letter)
}

/// Преобразование данных в json.
#[pyfunction]
fn xlsxheets_to_json(sheets: Vec<WrapperXLSXSheet>) -> PyResult<String> {
    let xlsx_sheets: Vec<core_rs::types::XLSXSheet> = sheets.into_iter().map(|s| s.0).collect();

    let json = core_rs::utils::xlsxheets_to_json(xlsx_sheets);

    Ok(json)
}

/// Returns the version of the underlying queue_rs library.
///
/// Returns
/// -------
/// version : str
///   The version of the underlying queue_rs library.
///
#[pyfunction]
fn version() -> String {
    core_rs::version().to_string()
}

#[pymodule]
#[pyo3(name = "_core_xlsx")]
fn boriy_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Service>()?;
    m.add_class::<WrapperXLSXSheet>()?;
    m.add_class::<WrapperXLSXSheetCell>()?;
    m.add_class::<WrapperHelperSheet>()?;
    m.add_class::<WrapperHelperSheetCell>()?;

    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add_function(wrap_pyfunction!(column_number_to_letter, m)?)?;
    m.add_function(wrap_pyfunction!(xlsxheets_to_json, m)?)?;

    Ok(())
}
