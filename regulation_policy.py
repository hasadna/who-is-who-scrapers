
import json
import xlrd
import logging
import argparse
from enum import Enum

class Department(Enum):
    economics = 1
    health = 2
    agriculture = 3
    education = 4
    environment = 5
    law = 6
    transportation = 7
    energy = 8
    religion = 9
    communication = 10
    interior = 11
    welfare = 12
    treasure = 13
    construction = 14
    public_security = 15
    tourism = 16
    culture = 17
    defence = 18
    prime_minister = 19
    senior_citizens = 20


DEPARTMENTS_PARAMS = ((4, 6, Department.economics),
                      (8, 9, Department.health),
                      (11, 12, Department.agriculture),
                      (14, 15, Department.education),
                      (17, 17, Department.environment),
                      (19, 20, Department.law),
                      (22, 23, Department.transportation),
                      (25, 26, Department.energy),
                      (28, 28, Department.religion),
                      (30, 31, Department.communication),
                      (33, 33, Department.interior),
                      (35, 35, Department.welfare),
                      (37, 37, Department.treasure),
                      (39, 39, Department.construction),
                      (41, 41, Department.public_security),
                      (43, 43, Department.tourism),
                      (45, 45, Department.culture),
                      (47, 47, Department.defence),
                      (49, 49, Department.prime_minister),
                      (51, 51, Department.senior_citizens))


def save_json_pretty(output_filepath, json_dict):
    with open(output_filepath, 'w') as output_file:
        json_data = json.dumps(json_dict, sort_keys=False, indent=2, separators=(',', ': '))
        output_file.write(json_data)


def append_str(str, new_str):
    if str is None:
        return new_str
    if new_str is None:
        return str
    return str + new_str


def reverse_lines(str):
    if str is None:
        return None
    return " ".join(str.split('\n')[::-1])

class RegulatorBuilder:

    def __init__(self):
        self.index = None
        self.unit = None
        self.manager = None
        self.subject_to = None
        self.superior = None
        self.main_activities = None
        self.department = None

    def build(self):
        return Regulator(self.index, reverse_lines(self.unit), reverse_lines(self.manager),
                         reverse_lines(self.subject_to), reverse_lines(self.superior),
                         reverse_lines(self.main_activities),
                         self.department)

    def append_unit(self, unit):
        self.unit = append_str(self.unit, unit)

    def append_manager(self, manager):
        self.manager = append_str(self.manager, manager)

    def append_subject_to(self, subject_to):
        self.subject_to = append_str(self.subject_to, subject_to)

    def append_superior(self, superior):
        self.superior = append_str(self.superior, superior)

    def append_main_activities(self, main_activities):
        self.main_activities = append_str(self.main_activities, main_activities)


class Regulator:

    def __init__(self, index, unit, manager, subject_to, superior, main_activities, department):
        self.index = index
        self.unit = unit
        self.manager = manager
        self.subject_to = subject_to
        self.superior = superior
        self.main_activities = main_activities
        self.department = department


    def as_entry(self):
        return {"unit": self.unit, "manager": self.manager, "subjectTo": self.subject_to, "superior" : self.superior,
                "mainActivities": self.main_activities, "department": self.department.value}

    def __repr__(self):
        return "[Regulator: index - %d, unit - %s, manager - %s, subject_to - %s, superior - %s, mainActivities - %s, department - %s]" \
               % (self.index, self.unit, self.manager, self.subject_to, self.superior, self.main_activities, self.department)

def reverse(str):
    if str is None:
        return None
    return str[::-1]


def reverse_text_cell_to_string(text_cell):
    if text_cell.ctype == xlrd.XL_CELL_TEXT:
        return reverse(text_cell.value).encode("utf-8")
    return None

def number_cell_to_positive_int(number_cell):
    if number_cell.ctype == xlrd.XL_CELL_NUMBER:
        return number_cell.value
    return -1

def parse_table(book, start_page, end_page, department):
    current_builder = None
    regulators = []
    current_index = 0
    for page in xrange(start_page, end_page + 1):
        page_sheet = book.sheet_by_name("Page " + str(page))

        for i in xrange(page_sheet.nrows):
            row = page_sheet.row(i)
            next_index = number_cell_to_positive_int(row[-1])
            if i != 0 and next_index != -1 and next_index > current_index + 1:
                logging.warn("missing indices between %d and %d, jumping to %d" % (current_index, next_index, next_index))

            if page == start_page:
                if i == 0 or (current_builder is None and next_index == -1):
                    # Skip table headers
                    continue
            else:
                if i + 1 < page_sheet.nrows and next_index == -1 and number_cell_to_positive_int(page_sheet.row(i + 1)[-1]) != -1:
                    next_index = current_index + 1

            if next_index != -1:
                if current_builder is not None:
                    regulators.append(current_builder.build())
                current_builder = RegulatorBuilder()
                current_index = next_index
                current_builder.index = current_index
            current_builder.append_main_activities(reverse_text_cell_to_string(row[0]))
            current_builder.append_superior(reverse_text_cell_to_string(row[1]))
            current_builder.append_subject_to(reverse_text_cell_to_string(row[2]))
            current_builder.append_manager(reverse_text_cell_to_string(row[3]))
            current_builder.append_unit(reverse_text_cell_to_string(row[4]))
            current_builder.department = department

    if current_builder is not None:
        regulators.append(current_builder.build())

    return regulators


def setup_logging(args=None):
    logging_level = logging.WARNING
    if args.verbose:
        logging_level = logging.INFO
    config = {"level": logging_level}

    if args is not None and args.log_path != "":
        config["filename"] = args.log_path

    logging.basicConfig(**config)


def parse_arguments():

    parser = argparse.ArgumentParser(description= "Script for parsing regulators from xlsx file")

    parser.add_argument("input_file", help="The xlsx input file")
    parser.add_argument("--output_file", default="regulators.json", help="The output file")
    parser.add_argument("--log_path", default="", help="The log file path")
    parser.add_argument("--verbose", help="Increase logging verbosity", action="store_true")

    return parser.parse_args()



if __name__ == "__main__":

    parsed_args = parse_arguments()
    setup_logging(parsed_args)

    book = xlrd.open_workbook(parsed_args.input_file)
    regulators = []
    for dep_param in DEPARTMENTS_PARAMS:
        new_regulators = parse_table(book, dep_param[0], dep_param[1], dep_param[2])
        regulators.extend(new_regulators)

    json_regulators = []
    for regulator in regulators:
        json_regulators.append(regulator.as_entry())

    logging.info("Saving regulators to: %s" % parsed_args.output_file)
    save_json_pretty(parsed_args.output_file, json_regulators)