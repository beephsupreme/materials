# -*- coding: utf-8 -*-

# DATA
DATAFILE_PATH = "./data/"
INVENTORY_EXPORT = "data.txt"
BACKLOG_EXPORT = "bl.txt"
HFR_EXPORT = "hfr.txt"
SCHEDULE_URL = "https://www.toki.co.jp/purchasing/TLIHTML.files/sheet001.htm"
VALIDATION_DB = "validate.csv"
TRANSLATION_DB = "translate.csv"
MATERIALS_OUT = "materials.xlsx"
REORDER_OUT = "reorder.xlsx"

# REPORT
INVENTORY_LABELS = ['Part Number', 'On Hand', 'On Order', 'Reorder']
BACKLOG_LABELS = ['Part Number', 'Backlog', 'Factor']
HFR_LABELS = ['Part Number', 'HFR', 'Factor']
VALIDATE_LABELS = ['Part Number', 'Valid PN']
TRANSLATE_LABELS = ['Part Number', 'Factor']
SCHEDULE_USELESS_COLS = [1, 2, 3, 4]
SCHEDULE_USELESS_ROWS = 5
DATES_ROW = 3
DATES_COL_START = 5

HEADER = ["Part Number", "On Hand", "Backlog", "Released", "HFR", "On Order", "T-Avail", "R-Avail", "Reorder"]
HEADER_WIDTH = 9
PARSER = "html.parser"
ENGINE = "xlsxwriter"
SHEET_NAME = "Sheet1"

# COLUMN NAMES
PN = "Part Number"
OH = "On Hand"
OO = "On Order"
RO = "Reorder"
BL = "Backlog"
RLS = "Released"
HFR = "HFR"
TA = "T-Avail"
RA = "R-Avail"
MP = "Multiplier"
VP = "Vendor Part Num"
FAC = "Factor"
