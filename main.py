from openpyxl import load_workbook
import sys

# * cell value -> at which cell the item is located at, example: A7

SOURCE_COLUMN = 'D'  # Location of cell value from source file.
MAX_ROWS = 500
MAX_COLUMNS = 500
TARGET_COL = 4  # In which column the cell location will be inerted at
DESCRIPTION_CELLS = ['Description', 'Manufacturer', 'MPN', 'QTY', 'CELL']


def run(stock_file, bom_file):
    stock_wb = load_workbook(filename=stock_file, read_only=True)
    stock_sheetnames = stock_wb.sheetnames
    bom_wb = load_workbook(filename=bom_file)
    stock_dict = get_dict(stock_wb, stock_sheetnames)
    insert_cell_location(bom_wb, stock_dict, bom_file)


def get_dict(workbook, sheetnames):
    stk_dict = {}
    for i, sheet in enumerate(sheetnames):  # iterate through all the sheets in a wb
        workbook.active = i
        ws = workbook.active
        for row in ws.iter_rows(min_row=0, max_row=100, min_col=2, max_col=2):
            for cell in row:
                if type(cell).__name__ != 'MergedCell' and cell.value is not None and (
                        cell.value not in DESCRIPTION_CELLS):
                    cell_info_location = SOURCE_COLUMN + str(cell.row)
                    cell_location = ws[cell_info_location].value
                    stk_dict[cell.value] = cell_location
    return stk_dict


def insert_cell_location(bom_wb, stock_dict, bom_file):
    bom_ws = bom_wb.active
    bom_ws.cell(row=1, column=4).value = "CELL"  # Write a new column header "CELL"
    for row in bom_ws.iter_rows(min_row=2, max_row=100, min_col=2, max_col=2):
        for cell in row:
            if cell.value is not None and (cell.value in stock_dict.keys()):
                insert_location = SOURCE_COLUMN + str(cell.row)
                bom_ws[insert_location] = stock_dict[cell.value]
    bom_wb.save(bom_file)


if __name__ == '__main__':
    stock = sys.argv[1]
    bom = sys.argv[2]
    run(stock, bom)
