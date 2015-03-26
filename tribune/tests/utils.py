# utility functions to find a row or column for testing. Reduces need for
#   hard-coded row/col indices in tests.
#   call generally with the worksheet, desired cell value, the column/row to search
#       within, and starting/ending indices
#   returns -1 if the value is not found


def find_sheet_row(sheet, needle, column, start=0, end=100):
    try:
        for x in range(start, end+1):
            if sheet.cell_value(x, column) == needle:
                return x
    except IndexError:
        pass
    return -1


def find_sheet_col(sheet, needle, row, start=0, end=100):
    try:
        for x in range(start, end+1):
            if sheet.cell_value(row, x) == needle:
                return x
    except IndexError:
        pass
    return -1
