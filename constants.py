# -*- coding: utf-8 -*-

# DATA
DATAFILE_PATH = "./data/"
INVENTORY_AV_EXPORT = "data.txt"
BACKLOG_EXPORT = "bl.txt"
HFR_EXPORT = "hfr.txt"
SCHEDULE_URL = "https://www.toki.co.jp/purchasing/TLIHTML.files/sheet001.htm"
VALIDATION_DB = "validate.csv"
TRANSLATION_DB = "translate.csv"
OUTFILE = "_materials.xlsx"

# SHIPPING
SHIPPING_DATES = []
SHIPPING_WIDTH = 0
SCHEDULE_WIDTH = 5
SCHEDULE_LENGTH = 0
FIRSTLINE_TEXT = "TOKISTAR CODE"
WHITESPACE = "\u3000"
NEW_WHITESPACE = "0"
ELEMENT = "td"

# REPORT
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
