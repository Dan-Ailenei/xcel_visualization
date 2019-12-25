import xlrd

from xml_measurements.exceptions import FileFormatError


def extract_info(book, sheet_num, orientation):
    sheet = book.sheet_by_index(sheet_num)

    orientation = 0 if orientation == "COL" else 1
    origin = find_origin(orientation, sheet)


def find_origin(orientation, sheet):
    origin = 0, 0
    while not sheet.cell(*origin).value.strip():
        origin[orientation] += 1
        if origin[orientation] > 20:
            raise FileFormatError("The direction is wrong or an offset is needed")

    return origin


book = xlrd.open_workbook('/Users/dan.ailenei/myprojects/measurements/static_files/media/profesori.xlsx')
extract_info(book, 0, "COL") # orientation -> (COL, ROW)
