
import json
import logging
import argparse
import datetime
import urllib
import xlrd
import os
from lxml import html
import requests
from enum import Enum

# Constants

HAARRETZ_MINUYIM_URL = "http://www.haaretz.co.il/st/inter/DB/tm/2015/minuyim.xlsx"
THE_MARKER_MINUYIM_URL = "http://www.themarker.com/st/inter/DB/tm/2015/minuyim.xlsx"
MINUYIM_TMP_FILE_PATH = "minuyim_temp.xlsx"

MINUYIM_SHEET_NAME = "edit"

WOMAN_GENDER_STRING = "woman"
MAN_GENDER_STRING = "man"

CALCALIST_MINUYIM_URL_TEMPLATE = "http://www.calcalist.co.il/Ext/Comp/NominationList/CdaNominationList_Iframe/1,15014,L-0-XXX,00.html?parms=5543"


# Enum Definitions

class Gender(Enum):
    male = 1
    female = 2
    unavailable = 3


class Source(Enum):
    calcalist = 1
    themarker = 2


# Classes

class Minuy:

    def __init__(self, name, date, title, company, details, img_path, gender, source):
        self.name = name
        self.date = date
        self.title = title
        self.company = company
        self.details = details
        self.img_path = img_path
        self.gender = gender
        self.source = source

    # Assuming two minuyim are considered equal in case the name and date are equal

    def __hash__(self):
        return hash((self.name, self.date))

    def __eq__(self, other):
        return (self.name, self.date) == (other.name, other.date)

    def is_within_hours_range(self, hours_range):
        if hours_range == -1:
            return True

        time_delta = datetime.datetime.now() - datetime.datetime(self.date.year, self.date.month, self.date.day)
        return time_delta < datetime.timedelta(hours=hours_range)

    def as_entry(self):
        return {"name": self.name, "date": self.date, "title": self.title, "company": self.company,
                "details": self.details, "imagePath": self.img_path, "gender": self.gender.value,
                "source": self.source.value}

    def __repr__(self):
        return "[Minuy: name - %s, date - %s, title - %s, company - %s, details - %s, imgPath - %s, Gender: %d, \
            Source: %d]" % (self.name, self.date, self.title, self.company, self.details, self.img_path,
                             self.gender.value, self.source.value)


# Json Utilities

def date_handler(obj):
    if hasattr(obj, 'isoformat'):
        return obj.isoformat()
    else:
        raise TypeError


def save_json_pretty(output_filepath, json_dict):
    with open(output_filepath, 'w') as output_file:
        json_data = json.dumps(json_dict, default=date_handler, sort_keys=False, indent=2, separators=(',', ': '))
        output_file.write(json_data)



def date_cell_to_date_obj(date_cell, datemode):
    if date_cell.ctype == xlrd.XL_CELL_DATE:
        datetuple = xlrd.xldate_as_tuple(date_cell.value, datemode)
        if datetuple[3:] == (0, 0, 0):
            return datetime.date(datetuple[0], datetuple[1], datetuple[2])
        return None
    if date_cell.ctype == xlrd.XL_CELL_EMPTY:
        return None
    if date_cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return date_cell.value == 1
    if date_cell.ctype == xlrd.XL_CELL_TEXT:
        return text_call_to_date(date_cell)
    return date_cell.value


def text_cell_to_string(text_cell):
    if text_cell.ctype == xlrd.XL_CELL_TEXT:
        return text_cell.value.encode("utf-8")
    return None


def text_to_datetime(date_str):
    if date_str is None:
        return None

    # Hacks to fix errors in excel
    date_str = date_str.replace("//", "/").replace(".", "/")
    if date_str == "6/616":
        date_str = "6/6/16"

    year_format = "%Y"
    if len(date_str.split("/")[-1]) == 2:
        year_format = "%y"
    return datetime.datetime.strptime(date_str, "%d/%m/" + year_format)


def text_call_to_date(text_cell):
    if text_cell.ctype != xlrd.XL_CELL_TEXT:
        return None
    try:
        return text_to_datetime(text_cell.value)


    except Exception as e:
        logging.exception(e.message)
        return None


def gender_cell_to_enum(gender_cell):
    gender_as_string = text_cell_to_string(gender_cell)
    if gender_as_string is None:
        return None
    gender_as_string = gender_as_string.strip()
    return {
        WOMAN_GENDER_STRING: Gender.female,
        MAN_GENDER_STRING: Gender.male,
    }.get(gender_as_string, Gender.unavailable)


def row_values_to_minuy_obj(row, row_index, datemode):
    if len(row) < 7:
        return None

    # This is a hack for a specific invalid row
    name_index = 0
    date_index = 1
    if row[0].ctype == xlrd.XL_CELL_DATE and row[1].ctype == xlrd.XL_CELL_TEXT:
        name_index = 1
        date_index = 0

    return Minuy(text_cell_to_string(row[name_index]), date_cell_to_date_obj(row[date_index], datemode), text_cell_to_string(row[2]),
                 text_cell_to_string(row[3]), text_cell_to_string(row[4]), text_cell_to_string(row[5]),
                 gender_cell_to_enum(row[6]), Source.themarker)


def setup_logging(args=None):
    logging_level = logging.WARNING
    if args.verbose:
        logging_level = logging.INFO
    config = {"level": logging_level}

    if args is not None and args.log_path != "":
        config["filename"] = args.log_path

    logging.basicConfig(**config)


def parse_arguments():

    parser = argparse.ArgumentParser(description= "Script for saving all minuyim data to a json file")

    parser.add_argument("--hours_range", type=int, default=-1, help="Hours")
    parser.add_argument("--output_file", default="minuyim.json", help="The minuyim output file")
    parser.add_argument("--log_path", default="", help="The log file path")
    parser.add_argument("--verbose", help="Increase logging verbosity", action="store_true")

    return parser.parse_args()


def the_marker_minuyim_from_url(excel_url, hours_range):

    minuyim = set()

    urllib.urlretrieve(excel_url, MINUYIM_TMP_FILE_PATH)
    book = xlrd.open_workbook(MINUYIM_TMP_FILE_PATH)
    try:
        minuyim_sheet = book.sheet_by_name(MINUYIM_SHEET_NAME)
        logging.info("Found %d minuyim at %s" % (minuyim_sheet.nrows - 1, excel_url))
        for i in xrange(1, minuyim_sheet.nrows):
            minuy_obj = row_values_to_minuy_obj(minuyim_sheet.row(i), i, book.datemode)
            if minuy_obj is None:
                logging.warn("Minuy at row: " + i + " was not parsed properly")

            if minuy_obj.is_within_hours_range(hours_range):
                minuyim.add(minuy_obj)

    except Exception as e:
        logging.exception("Exception caught")
        return None
    finally:
        os.remove(MINUYIM_TMP_FILE_PATH)

    return minuyim


def the_marker_minuyim(hours_range):
    minuyim = set()
    minuyim.update(the_marker_minuyim_from_url(HAARRETZ_MINUYIM_URL, hours_range))
    minuyim.update(the_marker_minuyim_from_url(THE_MARKER_MINUYIM_URL, hours_range))
    return minuyim


def xpath_single_field_value(element, path):
    l = element.xpath(path)
    if len(l) != 1:
        return None
    return l[0]


def calcalist_minuyim(hours_range):
    minuyim = set()

    is_done = False
    idx = 1
    while not is_done:
        current_calcalist_page = CALCALIST_MINUYIM_URL_TEMPLATE.replace("XXX", str(idx))
        logging.info("Processing calcalist page: %s", current_calcalist_page)
        page = requests.get(current_calcalist_page)
        tree = html.fromstring(page.content)

        tables = tree.xpath('//table')
        if len(tables) <= 1:
            is_done = True
            continue

        for i in xrange(len(tables)-1):
            table = tables[i]
            date_str = xpath_single_field_value(table, './/div[@class="Nom_Date"][last()]/text()')
            title = xpath_single_field_value(table, './/div[@class="Nom_Title"]/a/text()')
            company = xpath_single_field_value(table, './/div[@class="Nom_Comp"]/text()')
            details = xpath_single_field_value(table, './/div[@class="Nom_SubTitle"]/a/text()')
            image_path = xpath_single_field_value(table, './/@src')
            minuy = Minuy(title, text_to_datetime(date_str).date(), None, company, details, image_path,
                          Gender.unavailable, Source.calcalist)

            if minuy.is_within_hours_range(hours_range):
                minuyim.add(minuy)
            else:
                is_done = True
                break

        idx += 1

    return minuyim

if __name__ == "__main__":

    parsed_args = parse_arguments()
    setup_logging(parsed_args)

    minuyim = set()

    minuyim.update(the_marker_minuyim(parsed_args.hours_range))
    minuyim.update(calcalist_minuyim(parsed_args.hours_range))

    logging.info("Found total of %d relevant minuyim (within time frame, without duplications)" % len(minuyim))
    minuyim_as_json_entries = []
    for minuy in minuyim:
        minuyim_as_json_entries.append(minuy.as_entry())

    logging.info("Saving to output json file: %s" % parsed_args.output_file)
    save_json_pretty(parsed_args.output_file, minuyim_as_json_entries)

