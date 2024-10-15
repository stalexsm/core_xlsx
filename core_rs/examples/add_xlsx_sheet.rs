use core_rs::types::XLSXSheet;

fn main() {
    let mut sheet = XLSXSheet::new("A".to_string(), 1, 5, 5);

    println!("Добавление значения в ячейку 100:100");
    let cell = sheet.write_cell(100, 100, "Жопа".to_string());
    println!("Cell: {:?}", cell);

    println!("Добавление формулы в ячейку 100:100");
    if let Ok(cell) = cell {
        let _ = cell.set_formula("=A1+B1".to_string());
    }

    if let Ok(cell) = sheet.find_cell_by_coordinate(100, 100) {
        println!("Find {:?}", cell);
        if let Some(cell) = cell {
            println!("Formula {:?}", cell.formula);
        }
    }

    for cell in sheet.cells.iter() {
        // println!("{} -> {}:{}", cell.cell, cell.row, cell.column);

        if cell.row == 100 && cell.column == 100 {
            println!(
                "{} -> {}:{} __ value: {:?}",
                cell.cell, cell.row, cell.column, cell.value
            );
        }
    }

    println!("max row: {}, max col: {}", sheet.max_row, sheet.max_column);
    println!("cnt {}", sheet.cells.len());
}
