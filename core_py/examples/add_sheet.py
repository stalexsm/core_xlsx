from core_xlsx import XLSXSheet


if __name__ == "__main__":
    sheet = XLSXSheet("A", 1)
    print("Sheet", sheet)

    print("Write Cell 100x100")
    cell = sheet.write_cell(100, 100, "Жопа")

    print("Sheet", sheet)

    cell = sheet.find_cell_by_coordinate(100, 100)

    print("Find Cell 100x100 value", cell.value if cell else cell)
