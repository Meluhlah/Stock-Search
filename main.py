from openpyxl import load_workbook
import sys

# * cell value -> at which cell the item is located at, example: A7

TARGET_COLUMN = "D"  # Which column the data will be inserted at.
SOURCE_COLUMN = 4  # Location of cell value from source file.


def run(stock_file, bom_file):
    stock_wb = load_workbook(filename=stock_file)
    stock_sheetnames = stock_wb.sheetnames
    bom_wb = load_workbook(filename=bom_file)
    stock_dict = get_dict(stock_wb, stock_sheetnames)
    insert_cell_location(bom_wb, stock_dict, bom_file)


def get_dict(workbook, sheetnames):
    stk_dict = {}
    for sheet in sheetnames:  # iterate through all the sheets in a wb
        workbook.active = workbook[sheet]
        ws = workbook.active
        for cell in ws.iter_rows(min_row=2, max_row=100, min_col=2, max_col=2):
            location = ws.cell(column=SOURCE_COLUMN, row=cell[0].row).value
            stk_dict[location] = cell[0].internal_value
    return stk_dict


def insert_cell_location(bom_wb, stock_dict, bom_file):
    bom_ws = bom_wb.active
    bom_ws.cell(row=1, column=4).value = "CELL"  # Write a new column header "CELL"
    for cell in bom_ws.iter_rows(min_row=2, max_row=100, min_col=2, max_col=2):
        if cell is not None:
            for key, val in stock_dict.items():
                if cell[0].internal_value == val:
                    destination = str(TARGET_COLUMN + str(cell[0].row))  # Set the cell location to be written
                    bom_ws[destination] = key
                    break
        else:
            continue
    bom_wb.save(bom_file)


if __name__ == '__main__':
    # stock = "C:\\Users\Lidor-lenovo\PycharmProjects\AltiumStockSearcher\stock.xlsx"
    stock = sys.argv[1]
    bom = sys.argv[2]
    # bom = "C:\\Users\Lidor-lenovo\PycharmProjects\AltiumStockSearcher\\bom.xlsx"
    run(stock, bom)
