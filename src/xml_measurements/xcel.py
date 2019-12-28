from collections import defaultdict
import xlrd
import xlsxwriter
from koala import Spreadsheet
from xml_measurements.exceptions import FileFormatError


def generate_new_xml(path, orientation, rules, out_file, sheet_num=0):
    orientation = 0 if orientation == "COL" else 1

    # copy the file first because for some reason the evaluation library is not working without this
    worksheet_in, workbook_out, worksheet_out = open_in_and_out(path, sheet_num, out_file)
    workbook_out.close()

    rules_coords = extract_rules_coords(worksheet_in, orientation, rules)
    sp = Spreadsheet(out_file)
    worksheet_in, workbook_out, worksheet_out = open_in_and_out(path, sheet_num, out_file)

    cell_format_err = workbook_out.add_format({'bg_color': 'red'})
    write_rules(worksheet_in, worksheet_out, rules_coords, orientation, sheet_num, cell_format_err, sp)
    workbook_out.close()

    # TODO: use this when you make better UI
    # for val in sheet_iterator(open_in(out_file, sheet_num)):
    #     print(val)


def open_in_and_out(path, sheet_num, path_out):
    worksheet_in = open_in(path, sheet_num)
    worksheet_out, workbook_out = create_out(path_out, worksheet_in)
    return worksheet_in, workbook_out, worksheet_out


def create_out(path_out, worksheet_in):
    workbook_out = xlsxwriter.Workbook(path_out)
    worksheet_out = workbook_out.add_worksheet()
    copy_worksheet(worksheet_in, worksheet_out)
    return worksheet_out, workbook_out


def open_in(path, sheet_num):
    book = xlrd.open_workbook(path)
    worksheet_in = book.sheet_by_index(sheet_num)

    return worksheet_in


def compute_formula_value(sp, coord, rule, sheet_num):
    name = xlrd.cellname(*coord)
    adress = f"Sheet{sheet_num + 1}!{name}"
    before_formula_value = sp.cell_evaluate(adress)
    sp.cell_set_formula(adress, rule)
    formula_value = sp.cell_evaluate(adress)
    sp.cell_set_formula(adress, '')
    sp.cell_set_value(adress, before_formula_value)

    return formula_value, before_formula_value


def create_xml_rule(coords, orientation, i, rule):
    for j, coord in enumerate(coords[:-1]):
        coord[orientation] = i
        xl_coord = xlrd.cellname(*coord)
        rule = rule.replace(f'${j + 1}', xl_coord)
    return rule


def write_rules(worksheet_in, worksheet_out, rules_coords, orientation, sheet_num, cell_format_err, sp):
    orientation = (orientation + 1) % 2
    length = [worksheet_in.nrows, worksheet_in.ncols][orientation]
    for i in range(1, length):
        for (_, rule), coords in rules_coords.items():
            rule = create_xml_rule(coords, orientation, i, rule)
            coord = coords[-1]
            coord[orientation] = i

            formula_value, old_value = compute_formula_value(sp, coord, rule, sheet_num)

            try:
                formula_value = f'{float(formula_value):.2f}'
                old_value = f'{float(old_value):.2f}'
            except ValueError:
                pass

            cell_format = cell_format_err if formula_value != old_value else None
            worksheet_out.write(*coord, formula_value, cell_format)


def copy_worksheet(worksheet_in, worksheet_out):
    for i in range(worksheet_in.nrows):
        for j in range(worksheet_in.ncols):
            worksheet_out.write(i, j, worksheet_in.cell_value(i, j))


def extract_rules_coords(sheet, orientation, rules):
    coord = find_origin(orientation, sheet)
    length = [sheet.nrows, sheet.ncols][orientation]

    rules_coords = defaultdict(list)

    for i in range(coord[orientation], length):
        coord[orientation] = i
        current_val = sheet.cell_value(*coord)
        add_coords_occurences(rules, current_val, rules_coords, coord)

    for coords in rules_coords.values():
        coords.sort()
    for key in rules_coords:
        rules_coords[key] = [coord for _, coord in rules_coords[key]]

    return rules_coords


def find_origin(orientation, sheet):
    origin = [0, 0]
    length_or = [sheet.nrows, sheet.ncols][orientation]
    length_nor = [sheet.nrows, sheet.ncols][(orientation + 1) % 2]

    for i in range(length_nor):
        for j in range(length_or):
            origin[orientation] = i
            if sheet.cell_value(*origin).strip():
                return origin
    raise FileFormatError("Content couldn't be found! Is the file empty ?")


def add_coords_occurences(rules, current_val, rules_coords, coord):
    for rule in rules:
        for i, names in enumerate(rule.names):
            if current_val in names:
                rules_coords[(rule.pk, rule.rule)].append((i, [*coord]))


def sheet_iterator(sheet):
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            yield sheet.cell_value(i, j)


class FakeRule:
    def __init__(self, names, rule, pk):
        self.names = names
        self.rule = rule
        self.pk = pk


if __name__ == '__main__':
    rules = [FakeRule(names="Profesori\nEmail\nVorbit", rule="=count($1:$2)", pk=1)]
    for rule in rules:
        rule.names = [line.split(',') for line in rule.names.splitlines()]

    print(generate_new_xml('/Users/dan.ailenei/myprojects/measurements/static_files/media/profesori.xlsx',
          "ROW", rules, "figure_out_the_name.xlsx")) # orientation -> (COL, ROW)
