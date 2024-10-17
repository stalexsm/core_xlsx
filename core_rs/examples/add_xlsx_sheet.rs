use core_rs::types::XLSXSheet;

fn main() {
    let mut sheet = XLSXSheet::new("A".to_string(), 1, 5, 5);

    println!("Добавление значения в ячейку 100:100");
    let _ = sheet.write_cell(100, 100, "Жопа".to_string());

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
