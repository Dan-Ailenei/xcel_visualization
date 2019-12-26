from collections import defaultdict

import xlrd
import xlsxwriter
from xml_measurements.exceptions import FileFormatError


def generate_new_xml(book, sheet_num, orientation, rules):
    orientation = 0 if orientation == "COL" else 1
    worksheet_in = book.sheet_by_index(sheet_num)
    rules_coords = extract_rules_coords(worksheet_in, orientation, rules)

    workbook_out = xlsxwriter.Workbook("figure_out_the_name.xlsx")
    worksheet_out = workbook_out.add_worksheet()

    copy_worksheet(worksheet_in, worksheet_out)
    write_rules(worksheet_in, worksheet_out, rules_coords, orientation)
    for (_, rule), coords in rules_coords.items():
        print(coords[-1])
        # print(xlrd.cellname(*coords[-1]))

    workbook_out.close()


def write_rules(worksheet_in, worksheet_out, rules_coords, orientation):
    orientation = (orientation + 1) % 2
    length = [worksheet_in.nrows, worksheet_in.ncols][orientation]
    for i in range(1, length):
        for (_, rule), coords in rules_coords.items():
            for j, coord in enumerate(coords[:-1]):
                coord[orientation] = i
                xl_coord = xlrd.cellname(*coord)
                rule = rule.replace(f'${j + 1}', xl_coord)
            coord = coords[-1]
            coord[orientation] = i
            worksheet_out.write(*coords[-1], rule)


def copy_worksheet(worksheet_in, worksheet_out):
    for i in range(worksheet_in.nrows):
        for j in range(worksheet_in.ncols):
            worksheet_out.write(i, j, worksheet_in.cell_value(i, j))


def extract_rules_coords(sheet, orientation, rules):
    coord = find_origin(orientation, sheet)

    rules_coords = defaultdict(list)
    current_val = sheet.cell(*coord).value
    while True:
        poz_pk_rule = find_rule_poz_and_pk_rule(rules, current_val)
        if poz_pk_rule:
            i, pk, rule = poz_pk_rule
            rules_coords[(pk, rule)].append((i, [*coord]))

        coord[orientation] += 1

        # TODO: better stop condition
        try:
            current_val = sheet.cell(*coord).value
        except IndexError:
            break

    for coords in rules_coords.values():
        coords.sort()
    for key in rules_coords:
        rules_coords[key] = [coord for _, coord in rules_coords[key]]

    return rules_coords


def find_origin(orientation, sheet):
    origin = [0, 0]
    while not sheet.cell(*origin).value.strip():
        origin[orientation] += 1
        if origin[orientation] > 20:
            raise FileFormatError("The direction is wrong or an offset is needed")

    return origin


def find_rule_poz_and_pk_rule(rules, current_val):
    for rule in rules:
        for i, names in enumerate(rule.names):
            if current_val in names:
                return i, rule.pk, rule.rule


book = xlrd.open_workbook('/Users/dan.ailenei/myprojects/measurements/static_files/media/profesori.xlsx')


class Rule:
    def __init__(self, names, rule, pk):
        self.names = names
        self.rule = rule
        self.pk = pk


rules = [Rule(names="Profesori\nEmail\nVorbit", rule="=sum($1, $2)", pk=1)]
for rule in rules:
    rule.names = [line.split(',') for line in rule.names.splitlines()]

print(generate_new_xml(book, 0, "ROW", rules)) # orientation -> (COL, ROW)
